import { readFile, writeFile, copyFile, mkdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { join } from "node:path";
import { calculateFeatures, localRiskDecision, buildReasons, RISK_LABELS } from "../risk-engine.js";
import { generateTemplateBuffer } from "../template-builder.js";

const root = fileURLToPath(new URL("..", import.meta.url));
const db = JSON.parse(await readFile(join(root, "data", "dummy-data.json"), "utf8"));
const members = db.members.map((member) => {
  const features = calculateFeatures(member, db.meta.sessionDates);
  const decision = localRiskDecision(features);
  return { ...member, features, risk: { ...decision, meta: RISK_LABELS[decision.choice], reasons: buildReasons(features) } };
});
const summary = {
  total: members.length,
  high: members.filter((member) => member.risk.choice === "high").length,
  watch: members.filter((member) => member.risk.choice === "watch").length,
  stable: members.filter((member) => member.risk.choice === "stable").length,
  avgAttendance: Math.round(members.reduce((sum, member) => sum + member.features.attendanceRate, 0) / members.length)
};

// 1. Write demo-dashboard.json
await writeFile(join(root, "docs", "demo-dashboard.json"), `${JSON.stringify({ meta: db.meta, mode: "demo", summary, members })}\n`, "utf8");
console.log(`GitHub Pages용 정적 데이터 ${members.length}명을 생성했습니다.`);

// 2. Generate template xlsx for static download
const templateBuffer = await generateTemplateBuffer(db.meta.sessionDates);
await writeFile(join(root, "docs", "maeum-on-attendance-template.xlsx"), templateBuffer);
console.log("GitHub Pages용 엑셀 양식(docs/maeum-on-attendance-template.xlsx)을 생성했습니다.");

// 3. Ensure vendor libraries exist in docs/vendor
const vendorDir = join(root, "docs", "vendor");
await mkdir(vendorDir, { recursive: true });
await copyFile(join(root, "node_modules", "exceljs", "dist", "exceljs.min.js"), join(vendorDir, "exceljs.min.js"));
await copyFile(join(root, "node_modules", "jszip", "dist", "jszip.min.js"), join(vendorDir, "jszip.min.js"));
console.log("브라우저용 엑셀 라이브러리(docs/vendor)를 복사했습니다.");
