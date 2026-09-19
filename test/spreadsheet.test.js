import test from "node:test";
import assert from "node:assert/strict";
import { normalizeStatus, formatExcelDate, validateMember } from "../spreadsheet.js";

test("한글과 영문 출석 상태를 내부 값으로 변환한다", () => {
  assert.equal(normalizeStatus("출석"), "present");
  assert.equal(normalizeStatus("인정결석"), "excused");
  assert.equal(normalizeStatus("late"), "late");
  assert.equal(normalizeStatus("미정"), null);
});

test("엑셀 날짜를 표준 날짜 문자열로 변환한다", () => {
  assert.equal(formatExcelDate("2026/9/7"), "2026-09-07");
  assert.equal(formatExcelDate(new Date(2026, 8, 7)), "2026-09-07");
});

test("필수 출석값이 빠진 행을 검증한다", () => {
  const errors = validateMember({ id: "NB-1", name: "테스트", age: 80, joinedAt: "2026-01-01", attendance: ["present", null] }, 2, 2);
  assert.ok(errors.some((error) => error.includes("모든 출석일")));
});
