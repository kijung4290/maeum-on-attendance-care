import test from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { loadXlsx } from "../xlsx-loader.js";

test("비정상 app.xml 메타데이터가 있는 XLSX도 복구해 읽는다", async () => {
  const source = new ExcelJS.Workbook();
  const sheet = source.addWorksheet("출석입력");
  sheet.getCell("A1").value = "참여자번호";
  const zip = await JSZip.loadAsync(Buffer.from(await source.xlsx.writeBuffer()));
  zip.file("docProps/app.xml", "<?xml version=\"1.0\"?><NotProperties />");
  const malformed = await zip.generateAsync({ type: "nodebuffer" });

  const recovered = await loadXlsx(malformed);
  assert.equal(recovered.getWorksheet("출석입력").getCell("A1").text, "참여자번호");
});
