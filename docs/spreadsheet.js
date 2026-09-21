export const STATUS_TO_KO = { present: "출석", late: "지각", absent: "결석", excused: "인정결석" };
export const KO_TO_STATUS = Object.fromEntries(Object.entries(STATUS_TO_KO).map(([key, value]) => [value, key]));

export function normalizeStatus(value) {
  const text = String(value ?? "").trim().toLowerCase();
  if (KO_TO_STATUS[text]) return KO_TO_STATUS[text];
  if (STATUS_TO_KO[text]) return text;
  return null;
}

export function formatExcelDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }
  const text = String(value ?? "").trim().replaceAll(".", "-").replaceAll("/", "-");
  const match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  return match ? `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}` : text;
}

export function validateMember(member, rowNumber, expectedAttendanceCount) {
  const errors = [];
  if (!member.id) errors.push(`${rowNumber}행: 참여자번호가 필요합니다.`);
  if (!member.name) errors.push(`${rowNumber}행: 이름이 필요합니다.`);
  if (!Number.isInteger(member.age) || member.age < 0 || member.age > 120) errors.push(`${rowNumber}행: 나이는 0~120 사이 숫자로 입력하세요.`);
  if (!member.joinedAt || !/^\d{4}-\d{2}-\d{2}$/.test(member.joinedAt)) errors.push(`${rowNumber}행: 등록일을 YYYY-MM-DD 형식으로 입력하세요.`);
  if (member.attendance.length !== expectedAttendanceCount || member.attendance.some((status) => !status)) {
    errors.push(`${rowNumber}행: 모든 출석일에 출석, 지각, 결석, 인정결석 중 하나를 입력하세요.`);
  }
  return errors;
}

export async function loadXlsxInBrowser(buffer, attempts = 0) {
  if (typeof window === "undefined" || !window.ExcelJS) {
    throw new Error("엑셀 라이브러리(ExcelJS)를 불러오지 못했습니다. 페이지를 새로고침하거나 브라우저 네트워크 연결을 확인해 주세요.");
  }
  const workbook = new window.ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(buffer);
    return workbook;
  } catch (error) {
    const malformedAppProperties = /reading ['"]company['"]/.test(String(error?.message));
    const prefixedSpreadsheetXml = /Unexpected xml node.*"name":"x:/.test(String(error?.message));
    if ((!malformedAppProperties && !prefixedSpreadsheetXml) || attempts >= 3 || !window.JSZip) {
      throw error;
    }
    const zip = await window.JSZip.loadAsync(buffer);
    if (malformedAppProperties) {
      if (!zip.file("docProps/app.xml")) throw error;
      zip.remove("docProps/app.xml");
    }
    if (prefixedSpreadsheetXml) {
      const xmlFiles = Object.values(zip.files).filter((entry) => !entry.dir && entry.name.endsWith(".xml"));
      for (const entry of xmlFiles) {
        const xml = await entry.async("string");
        if (/<\/?x:/.test(xml)) {
          zip.file(entry.name, xml.replace(/(<\/?)(?:x):/g, "$1"));
        }
      }
    }
    const sanitized = await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
    return loadXlsxInBrowser(sanitized, attempts + 1);
  }
}

export async function parseWorkbookInBrowser(buffer, defaultSessionDates = []) {
  let workbook;
  try {
    workbook = await loadXlsxInBrowser(buffer);
  } catch (error) {
    console.warn("XLSX parsing failed:", error);
    throw new Error("XLSX 파일 구조를 읽을 수 없습니다. 파일을 Microsoft Excel 또는 LibreOffice에서 XLSX 형식으로 다시 저장해 주세요.");
  }

  const sheet = workbook.getWorksheet("출석입력");
  if (!sheet) {
    throw new Error("'출석입력' 시트를 찾을 수 없습니다. 제공된 양식을 사용해 주세요.");
  }

  const baseHeaders = ["참여자번호", "이름", "나이", "성별", "연락처", "보호자/관계", "등록일", "담당자메모"];
  const rawHeaders = sheet.getRow(1).values.slice(1).map((val) => String(val ?? "").trim());

  for (let i = 0; i < baseHeaders.length; i++) {
    if (rawHeaders[i] !== baseHeaders[i]) {
      throw new Error("양식의 열 제목 또는 형식이 올바르지 않습니다. 새 양식을 내려받아 작성해 주세요.");
    }
  }

  let sessionDates = rawHeaders.slice(baseHeaders.length).map(formatExcelDate).filter(Boolean);
  if (!sessionDates.length) {
    if (defaultSessionDates && defaultSessionDates.length) {
      sessionDates = defaultSessionDates;
    } else {
      throw new Error("출석일자 열(9열 이후)이 비어있습니다. 출석일을 입력해 주세요.");
    }
  }

  const members = [];
  const errors = [];
  const ids = new Set();

  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    const values = row.values.slice(1);
    if (!values.some((val) => String(val ?? "").trim())) continue;

    const member = {
      id: String(row.getCell(1).text || "").trim(),
      name: String(row.getCell(2).text || "").trim(),
      age: Number(row.getCell(3).value),
      gender: String(row.getCell(4).text || "").trim(),
      phone: String(row.getCell(5).text || "").trim(),
      guardian: String(row.getCell(6).text || "").trim(),
      joinedAt: formatExcelDate(row.getCell(7).value),
      memo: String(row.getCell(8).text || "").trim(),
      attendance: sessionDates.map((_, idx) => normalizeStatus(row.getCell(9 + idx).text))
    };

    errors.push(...validateMember(member, rowNumber, sessionDates.length));
    if (ids.has(member.id)) {
      errors.push(`${rowNumber}행: 참여자번호 ${member.id}가 중복되었습니다.`);
    }
    ids.add(member.id);
    members.push(member);

    if (members.length > 1000) {
      errors.push("한 번에 최대 1,000명까지 업로드할 수 있습니다.");
    }
    if (errors.length >= 20) break;
  }

  if (!members.length) {
    errors.push("입력된 참여자 데이터가 없습니다.");
  }
  if (errors.length) {
    throw new Error(errors.slice(0, 20).join("\n"));
  }

  return { members, sessionDates };
}
