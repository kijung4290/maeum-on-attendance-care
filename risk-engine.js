export const RISK_LABELS = {
  high: { label: "집중 확인", color: "red" },
  watch: { label: "관심 관찰", color: "amber" },
  stable: { label: "안정", color: "green" }
};

export function calculateFeatures(member, dates) {
  const records = dates.map((date, index) => ({ date, status: member.attendance[index] ?? "absent" }));
  const counts = records.reduce((acc, row) => {
    acc[row.status] = (acc[row.status] || 0) + 1;
    return acc;
  }, { present: 0, late: 0, absent: 0, excused: 0 });

  let consecutiveAbsences = 0;
  for (let i = records.length - 1; i >= 0 && records[i].status === "absent"; i -= 1) consecutiveAbsences += 1;
  const recent = records.slice(-4);
  const recentAbsences = recent.filter((r) => r.status === "absent").length;
  const recentLate = recent.filter((r) => r.status === "late").length;
  const attended = counts.present + counts.late;
  const eligible = records.length - counts.excused;
  const attendanceRate = eligible ? Math.round((attended / eligible) * 100) : 100;
  const previousHalf = records.slice(0, 6);
  const latestHalf = records.slice(-6);
  const rateOf = (items) => items.length ? items.filter((r) => r.status === "present" || r.status === "late").length / items.filter((r) => r.status !== "excused").length : 1;
  const trendDrop = Math.max(0, Math.round((rateOf(previousHalf) - rateOf(latestHalf)) * 100));

  return { records, counts, consecutiveAbsences, recentAbsences, recentLate, attendanceRate, trendDrop };
}

export function localRiskDecision(features) {
  let score = 0;
  score += features.consecutiveAbsences * 24;
  score += features.recentAbsences * 12;
  score += features.recentLate * 4;
  score += Math.max(0, 80 - features.attendanceRate) * 0.7;
  score += features.trendDrop * 0.35;
  score = Math.min(100, Math.round(score));

  const choice = score >= 58 ? "high" : score >= 25 ? "watch" : "stable";
  const lead = choice === "high" ? Math.min(.94, .55 + score / 250) : choice === "watch" ? Math.min(.88, .54 + score / 180) : Math.min(.96, .68 + (100 - score) / 350);
  const rest = (1 - lead) / 2;
  const probabilities = { high: rest, watch: rest, stable: rest, [choice]: lead };

  return { choice, probabilities, confidence: lead, source: "demo", score };
}

export function buildReasons(features) {
  const reasons = [];
  if (features.consecutiveAbsences >= 2) reasons.push(`최근 ${features.consecutiveAbsences}회 연속 무단결석`);
  if (features.recentAbsences >= 2 && features.consecutiveAbsences < 2) reasons.push(`최근 4회 중 ${features.recentAbsences}회 결석`);
  if (features.consecutiveAbsences === 1 && features.recentAbsences < 2) reasons.push("최근 회차 결석 사유 확인 필요");
  if (features.trendDrop >= 20) reasons.push(`이전 대비 출석률 ${features.trendDrop}%p 하락`);
  if (features.recentLate >= 2) reasons.push(`최근 4회 중 ${features.recentLate}회 지각`);
  if (features.attendanceRate >= 90 && !reasons.length) reasons.push("최근 출석 흐름이 안정적임");
  if (!reasons.length) reasons.push("뚜렷한 위험 징후 없음");
  return reasons;
}

export function normalizeJevAnswer(answer, fallback) {
  if (!answer || !["high", "watch", "stable"].includes(answer.choice)) return fallback;
  const probabilities = answer.probabilities || {};
  const confidence = Number(answer.confidence ?? probabilities[answer.choice] ?? 0);
  return { choice: answer.choice, probabilities, confidence, source: "jev", score: Math.round((probabilities.high || 0) * 100) };
}
