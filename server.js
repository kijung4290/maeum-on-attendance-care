import http from "node:http";
import { readFile } from "node:fs/promises";
import { existsSync, readFileSync } from "node:fs";
import { extname, join, normalize } from "node:path";
import { fileURLToPath } from "node:url";
import { calculateFeatures, localRiskDecision, buildReasons, normalizeJevAnswer, RISK_LABELS } from "./risk-engine.js";

const root = fileURLToPath(new URL(".", import.meta.url));
loadEnv(join(root, ".env"));
const port = Number(process.env.PORT || 4173);
const db = JSON.parse(await readFile(join(root, "data", "dummy-data.json"), "utf8"));
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
    const members = db.members.map(publicMember);
    const summary = {
      total: members.length,
      high: members.filter((m) => m.risk.choice === "high").length,
      watch: members.filter((m) => m.risk.choice === "watch").length,
      stable: members.filter((m) => m.risk.choice === "stable").length,
      avgAttendance: Math.round(members.reduce((sum, m) => sum + m.features.attendanceRate, 0) / members.length)
    };
    return sendJson(res, 200, { meta: db.meta, mode: process.env.TYPESAFE_API_KEY ? "jev" : "demo", summary, members });
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
      const results = [];
      for (const member of db.members) {
        const decision = await callJev(member);
        cache.set(member.id, decision);
        results.push(publicMember(member));
      }
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
