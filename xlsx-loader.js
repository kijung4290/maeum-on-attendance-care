import ExcelJS from "exceljs";
import JSZip from "jszip";

async function loadXlsxWithRecovery(buffer, attempts = 0) {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(buffer);
    return workbook;
  } catch (error) {
    const malformedAppProperties = /reading ['"]company['"]/.test(String(error?.message));
    const prefixedSpreadsheetXml = /Unexpected xml node.*"name":"x:/.test(String(error?.message));
    if ((!malformedAppProperties && !prefixedSpreadsheetXml) || attempts >= 3) throw error;
    const zip = await JSZip.loadAsync(buffer);
    if (malformedAppProperties) {
      if (!zip.file("docProps/app.xml")) throw error;
      zip.remove("docProps/app.xml");
    }
    if (prefixedSpreadsheetXml) {
      const xmlFiles = Object.values(zip.files).filter((entry) => !entry.dir && entry.name.endsWith(".xml"));
      for (const entry of xmlFiles) {
        const xml = await entry.async("string");
        if (/<\/?x:/.test(xml)) zip.file(entry.name, xml.replace(/(<\/?)(?:x):/g, "$1"));
      }
    }
    const sanitized = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
    return loadXlsxWithRecovery(sanitized, attempts + 1);
  }
}

export function loadXlsx(buffer) {
  return loadXlsxWithRecovery(buffer);
}
