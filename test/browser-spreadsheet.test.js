import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { parseWorkbookInBrowser } from "../docs/spreadsheet.js";
import { calculateFeatures, localRiskDecision, buildReasons } from "../docs/risk-engine.js";

const root = fileURLToPath(new URL("..", import.meta.url));

// Setup simulated window environment for ExcelJS and JSZip
global.window = {
  ExcelJS,
  JSZip
};

test("브라우저 환경(window.ExcelJS)에서 1000명 XLSX 파일을 성공적으로 파싱하고 분석한다", async () => {
  const fileBuffer = await readFile(join(root, "sample-data", "가상_어르신_출석데이터_1000명.xlsx"));
  const arrayBuffer = fileBuffer.buffer.slice(fileBuffer.byteOffset, fileBuffer.byteOffset + fileBuffer.byteLength);

  const { members, sessionDates } = await parseWorkbookInBrowser(arrayBuffer);
  assert.equal(members.length, 1000);
  assert.equal(sessionDates.length, 12);

  // Verify first member
  const first = members[0];
  assert.equal(first.id, "SYN-0001");
  assert.equal(first.name, "가상참여자0001");
  assert.equal(first.attendance.length, 12);

  // Run risk engine
  const features = calculateFeatures(first, sessionDates);
  const decision = localRiskDecision(features);
  const reasons = buildReasons(features);

  assert.ok(["high", "watch", "stable"].includes(decision.choice));
  assert.ok(decision.confidence > 0);
  assert.ok(reasons.length > 0);
});

test("브라우저 환경에서 템플릿 XLSX 양식의 헤더 및 구성을 올바르게 인식한다", async () => {
  const fileBuffer = await readFile(join(root, "docs", "maeum-on-attendance-template.xlsx"));
  const arrayBuffer = fileBuffer.buffer.slice(fileBuffer.byteOffset, fileBuffer.byteOffset + fileBuffer.byteLength);

  const error = await parseWorkbookInBrowser(arrayBuffer).catch((err) => err);
  // Empty template should throw "입력된 참여자 데이터가 없습니다."
  assert.match(error.message, /입력된 참여자 데이터가 없습니다/);
});
