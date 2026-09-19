import test from "node:test";
import assert from "node:assert/strict";
import { calculateFeatures, localRiskDecision, normalizeJevAnswer } from "../risk-engine.js";

const dates = Array.from({ length: 12 }, (_, i) => `2026-09-${String(i + 1).padStart(2, "0")}`);

test("연속 무단결석은 집중 확인으로 판정한다", () => {
  const member = { attendance: ["present","present","present","present","present","present","present","present","present","absent","absent","absent"] };
  const features = calculateFeatures(member, dates);
  assert.equal(features.consecutiveAbsences, 3);
  assert.equal(localRiskDecision(features).choice, "high");
});

test("사전 연락한 인정결석은 출석률과 위험도를 낮추지 않는다", () => {
  const member = { attendance: ["present","present","present","present","present","present","present","present","present","present","excused","excused"] };
  const features = calculateFeatures(member, dates);
  assert.equal(features.attendanceRate, 100);
  assert.equal(features.consecutiveAbsences, 0);
  assert.equal(localRiskDecision(features).choice, "stable");
});

test("JEV 응답을 앱의 표준 판정 형태로 변환한다", () => {
  const fallback = { choice: "stable", source: "demo" };
  const result = normalizeJevAnswer({ choice: "watch", probabilities: { high: .1, watch: .75, stable: .15 }, confidence: .75 }, fallback);
  assert.equal(result.choice, "watch");
  assert.equal(result.source, "jev");
  assert.equal(result.score, 10);
});

test("예상하지 못한 JEV 선택지는 안전하게 로컬 판정으로 대체한다", () => {
  const fallback = { choice: "stable", source: "demo" };
  assert.equal(normalizeJevAnswer({ choice: "unknown" }, fallback), fallback);
});
