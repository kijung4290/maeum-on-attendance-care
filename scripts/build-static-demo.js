import { readFile, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { calculateFeatures, localRiskDecision, buildReasons, RISK_LABELS } from "../risk-engine.js";

const root = fileURLToPath(new URL("..", import.meta.url));
const db = JSON.parse(await readFile(`${root}/data/dummy-data.json`, "utf8"));
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
await writeFile(`${root}/docs/demo-dashboard.json`, `${JSON.stringify({ meta: db.meta, mode: "demo", summary, members })}\n`, "utf8");
console.log(`GitHub Pages용 정적 데이터 ${members.length}명을 생성했습니다.`);
