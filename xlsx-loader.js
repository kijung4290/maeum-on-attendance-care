import ExcelJS from "exceljs";
import JSZip from "jszip";

export async function loadXlsx(buffer) {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(buffer);
    return workbook;
  } catch (error) {
    const malformedAppProperties = /reading ['"]company['"]/.test(String(error?.message));
    if (!malformedAppProperties) throw error;

    const zip = await JSZip.loadAsync(buffer);
    if (!zip.file("docProps/app.xml")) throw error;
    zip.remove("docProps/app.xml");
    const sanitized = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
    const recovered = new ExcelJS.Workbook();
    await recovered.xlsx.load(sanitized);
    return recovered;
  }
}
