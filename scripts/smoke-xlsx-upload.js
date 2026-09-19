import ExcelJS from "exceljs";
import JSZip from "jszip";

const baseUrl = process.env.APP_URL || "http://localhost:4173";
const templateResponse = await fetch(`${baseUrl}/api/template`);
if (!templateResponse.ok) throw new Error(`양식 다운로드 실패: ${templateResponse.status}`);

const workbook = new ExcelJS.Workbook();
await workbook.xlsx.load(Buffer.from(await templateResponse.arrayBuffer()));
const sheet = workbook.getWorksheet("출석입력");
if (!sheet) throw new Error("출석입력 시트가 없습니다.");

const values = ["TEST-001", "테스트참여자", 80, "여", "010-****-0000", "딸", "2026-01-01", "통합 테스트"];
values.forEach((value, index) => { sheet.getCell(2, index + 1).value = value; });
for (let column = 9; column <= 20; column += 1) sheet.getCell(2, column).value = column === 20 ? "결석" : "출석";

let uploadBuffer = Buffer.from(await workbook.xlsx.writeBuffer());
if (process.env.MALFORMED_APP_XML === "1") {
  const zip = await JSZip.loadAsync(uploadBuffer);
  zip.file("docProps/app.xml", "<?xml version=\"1.0\"?><NotProperties />");
  uploadBuffer = await zip.generateAsync({ type: "nodebuffer" });
}
const uploadResponse = await fetch(`${baseUrl}/api/upload`, {
  method: "POST",
  headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  body: uploadBuffer
});
const result = await uploadResponse.json();
if (!uploadResponse.ok) throw new Error(result.error || `업로드 실패: ${uploadResponse.status}`);
if (result.uploaded !== 1 || result.members?.[0]?.id !== "TEST-001") throw new Error("업로드 결과가 예상과 다릅니다.");
console.log(`XLSX 통합 테스트 성공${process.env.MALFORMED_APP_XML === "1" ? "(비정상 메타데이터 복구)" : ""}: ${result.uploaded}명, ${result.members[0].risk.choice}`);
