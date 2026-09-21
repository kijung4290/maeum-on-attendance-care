import http from "node:http";
import { readFile } from "node:fs/promises";
import { existsSync, readFileSync } from "node:fs";
import { extname, join, normalize } from "node:path";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";
import { calculateFeatures, localRiskDecision, buildReasons, normalizeJevAnswer, RISK_LABELS } from "./risk-engine.js";
import { STATUS_TO_KO, normalizeStatus, formatExcelDate, validateMember } from "./spreadsheet.js";
import { loadXlsx } from "./xlsx-loader.js";
import { generateTemplateBuffer } from "./template-builder.js";

const root = fileURLToPath(new URL(".", import.meta.url));
loadEnv(join(root, ".env"));
const port = Number(process.env.PORT || 4173);
const seedDb = JSON.parse(await readFile(join(root, "data", "dummy-data.json"), "utf8"));
let db = structuredClone(seedDb);
const cache = new Map();

function loadEnv(path) {
  if (!existsSync(path)) return;
  for (const line of readFileSync(path, "utf8").split(/\r?\n/)) {
    const match = line.match(/^\s*([A-Z_][A-Z0-9_]*)\s*=\s*(.*)\s*$/);
    if (match && !process.env[match[1]]) process.env[match[1]] = match[2].replace(/^['"]|['"]$/g, "");
  }
}

function publicMember(member) {
  const features = calculateFeatures(member, db.meta.sessionDates);
  const decision = cache.get(member.id) || localRiskDecision(features);
  return { ...member, features, risk: { ...decision, meta: RISK_LABELS[decision.choice], reasons: buildReasons(features) } };
}

function sendJson(res, status, data) {
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" });
  res.end(JSON.stringify(data));
}

function dashboardPayload() {
  const members = db.members.map(publicMember);
  const summary = {
    total: members.length,
    high: members.filter((m) => m.risk.choice === "high").length,
    watch: members.filter((m) => m.risk.choice === "watch").length,
    stable: members.filter((m) => m.risk.choice === "stable").length,
    avgAttendance: members.length ? Math.round(members.reduce((sum, m) => sum + m.features.attendanceRate, 0) / members.length) : 0
  };
  return { meta: db.meta, mode: process.env.TYPESAFE_API_KEY ? "jev" : "demo", summary, members };
}

function readBody(req, maxBytes = 10 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > maxBytes) {
        reject(Object.assign(new Error("파일은 10MB 이하만 업로드할 수 있습니다."), { status: 413 }));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

async function createTemplate() {
  return generateTemplateBuffer(db.meta.sessionDates);
}

async function parseWorkbook(buffer) {
  let workbook;
  try {
    workbook = await loadXlsx(buffer);
  } catch (error) {
    console.warn("XLSX parsing failed:", error.message);
    throw Object.assign(new Error("XLSX 파일 구조를 읽을 수 없습니다. 파일을 Microsoft Excel 또는 LibreOffice에서 XLSX 형식으로 다시 저장해 주세요."), { status: 400 });
  }
  const sheet = workbook.getWorksheet("출석입력");
  if (!sheet) throw Object.assign(new Error("'출석입력' 시트를 찾을 수 없습니다. 제공된 양식을 사용해 주세요."), { status: 400 });
  const expectedHeaders = ["참여자번호", "이름", "나이", "성별", "연락처", "보호자/관계", "등록일", "담당자메모", ...db.meta.sessionDates];
  const headers = sheet.getRow(1).values.slice(1).map((value) => String(value ?? "").trim());
  if (expectedHeaders.some((header, index) => headers[index] !== header)) throw Object.assign(new Error("양식의 열 제목 또는 출석 날짜가 변경되었습니다. 새 양식을 내려받아 작성해 주세요."), { status: 400 });
  const members = [];
  const errors = [];
  const ids = new Set();
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    const values = row.values.slice(1);
    if (!values.some((value) => String(value ?? "").trim())) continue;
    const member = {
      id: String(row.getCell(1).text || "").trim(), name: String(row.getCell(2).text || "").trim(),
      age: Number(row.getCell(3).value), gender: String(row.getCell(4).text || "").trim(),
      phone: String(row.getCell(5).text || "").trim(), guardian: String(row.getCell(6).text || "").trim(),
      joinedAt: formatExcelDate(row.getCell(7).value), memo: String(row.getCell(8).text || "").trim(),
      attendance: db.meta.sessionDates.map((_, index) => normalizeStatus(row.getCell(9 + index).text))
    };
    errors.push(...validateMember(member, rowNumber, db.meta.sessionDates.length));
    if (ids.has(member.id)) errors.push(`${rowNumber}행: 참여자번호 ${member.id}가 중복되었습니다.`);
    ids.add(member.id);
    members.push(member);
    if (members.length > 1000) errors.push("한 번에 최대 1,000명까지 업로드할 수 있습니다.");
    if (errors.length >= 20) break;
  }
  if (!members.length) errors.push("입력된 참여자 데이터가 없습니다.");
  if (errors.length) throw Object.assign(new Error(errors.slice(0, 20).join("\n")), { status: 400 });
  return members;
}

async function callJev(member) {
  const features = calculateFeatures(member, db.meta.sessionDates);
  const fallback = localRiskDecision(features);
  if (!process.env.TYPESAFE_API_KEY) return fallback;

  const state = {
    subject_id: member.id,
    program_type: "senior_wellness_class",
    attendance_summary: {
      attendance_rate_percent: features.attendanceRate,
      consecutive_unexcused_absences: features.consecutiveAbsences,
      recent_4_unexcused_absences: features.recentAbsences,
      recent_4_late_arrivals: features.recentLate,
      trend_drop_percentage_points: features.trendDrop,
      status_by_date: features.records
    },
    staff_note: member.memo
  };
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 12000);
  try {
    const response = await fetch("https://api.typesafe.ai/v1/systemone", {
      method: "POST",
      headers: { "Authorization": `Bearer ${process.env.TYPESAFE_API_KEY}`, "Content-Type": "application/json" },
      signal: controller.signal,
      body: JSON.stringify({
        model: "jev-latest",
        state,
        questions: {
          attendance_risk: {
            type: "choice",
            instructions: "Classify the attendance disengagement risk. Focus on unexplained absence streaks, recent decline, and repeated lateness. Excused absences must not increase risk. This supports a Korean social worker's follow-up priority and is not a medical diagnosis.",
            criteria: {
              high: "Urgent human follow-up is appropriate due to a strong, recent disengagement signal.",
              watch: "Some concerning attendance change exists and should be monitored or checked.",
              stable: "Attendance is broadly stable with no meaningful recent disengagement signal."
            }
          }
        }
      })
    });
    const body = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(body?.message || `Jev API 오류 (${response.status})`);
    return normalizeJevAnswer(body?.answers?.attendance_risk, fallback);
  } finally {
    clearTimeout(timeout);
  }
}

async function api(req, res, url) {
  if (req.method === "GET" && url.pathname === "/api/dashboard") {
    return sendJson(res, 200, dashboardPayload());
  }
  if (req.method === "GET" && url.pathname === "/api/template") {
    const file = await createTemplate();
    res.writeHead(200, { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Content-Disposition": "attachment; filename*=UTF-8''maeum-on-attendance-template.xlsx", "Content-Length": file.length });
    res.end(file);
    return;
  }
  if (req.method === "POST" && url.pathname === "/api/upload") {
    try {
      if (!String(req.headers["content-type"] || "").includes("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) return sendJson(res, 415, { error: "XLSX 파일만 업로드할 수 있습니다." });
      const members = await parseWorkbook(await readBody(req));
      db = { ...db, members };
      cache.clear();
      return sendJson(res, 200, { ...dashboardPayload(), uploaded: members.length });
    } catch (error) {
      return sendJson(res, error.status || 400, { error: error.message || "엑셀 파일을 읽지 못했습니다." });
    }
  }
  const match = url.pathname.match(/^\/api\/analyze\/([^/]+)$/);
  if (req.method === "POST" && match) {
    const member = db.members.find((m) => m.id === decodeURIComponent(match[1]));
    if (!member) return sendJson(res, 404, { error: "대상자를 찾을 수 없습니다." });
    try {
      const result = await callJev(member);
      cache.set(member.id, result);
      return sendJson(res, 200, publicMember(member));
    } catch (error) {
      return sendJson(res, 502, { error: error.name === "AbortError" ? "Jev API 응답 시간이 초과되었습니다." : error.message });
    }
  }
  if (req.method === "POST" && url.pathname === "/api/analyze-all") {
    try {
      const results = new Array(db.members.length);
      let cursor = 0;
      async function worker() {
        while (cursor < db.members.length) {
          const index = cursor++;
          const member = db.members[index];
          const decision = await callJev(member);
          cache.set(member.id, decision);
          results[index] = publicMember(member);
        }
      }
      await Promise.all(Array.from({ length: Math.min(5, db.members.length) }, worker));
      return sendJson(res, 200, { members: results, mode: process.env.TYPESAFE_API_KEY ? "jev" : "demo" });
    } catch (error) {
      return sendJson(res, 502, { error: error.name === "AbortError" ? "Jev API 응답 시간이 초과되었습니다." : error.message });
    }
  }
  return false;
}

const mime = { ".html": "text/html; charset=utf-8", ".css": "text/css; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".svg": "image/svg+xml", ".json": "application/json; charset=utf-8" };
const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, `http://${req.headers.host || "localhost"}`);
    if (url.pathname.startsWith("/api/")) {
      const handled = await api(req, res, url);
      if (handled !== false) return;
      return sendJson(res, 404, { error: "API 경로를 찾을 수 없습니다." });
    }
    const relative = url.pathname === "/" ? "index.html" : decodeURIComponent(url.pathname.slice(1));
    const filePath = normalize(join(root, "docs", relative));
    if (!filePath.startsWith(normalize(join(root, "docs")))) return sendJson(res, 403, { error: "접근할 수 없습니다." });
    const content = await readFile(filePath);
    res.writeHead(200, { "Content-Type": mime[extname(filePath)] || "application/octet-stream" });
    res.end(content);
  } catch (error) {
    if (error.code === "ENOENT") return sendJson(res, 404, { error: "페이지를 찾을 수 없습니다." });
    console.error(error);
    sendJson(res, 500, { error: "서버 오류가 발생했습니다." });
  }
});

server.listen(port, () => {
  console.log(`\n  마음온 출석 돌봄  http://localhost:${port}`);
  console.log(`  분석 모드: ${process.env.TYPESAFE_API_KEY ? "Jev API" : "데모(로컬 규칙)"}\n`);
});
