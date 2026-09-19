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

test("x: 접두사가 붙은 스프레드시트 XML도 정규화해 읽는다", async () => {
  const source = new ExcelJS.Workbook();
  const sheet = source.addWorksheet("출석입력");
  sheet.getCell("A1").value = "참여자번호";
  const zip = await JSZip.loadAsync(Buffer.from(await source.xlsx.writeBuffer()));
  const sharedStrings = await zip.file("xl/sharedStrings.xml").async("string");
  const prefixed = sharedStrings
    .replace(/<(\/?)sst(?=[ >])/g, "<$1x:sst")
    .replace(/<(\/?)si(?=[ >])/g, "<$1x:si")
    .replace(/<(\/?)t(?=[ >])/g, "<$1x:t")
    .replace("xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"", "xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
  zip.file("xl/sharedStrings.xml", prefixed);

  const recovered = await loadXlsx(await zip.generateAsync({ type: "nodebuffer" }));
  assert.equal(recovered.getWorksheet("출석입력").getCell("A1").text, "참여자번호");
});
