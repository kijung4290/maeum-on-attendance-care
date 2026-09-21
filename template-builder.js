import ExcelJS from "exceljs";

export async function createTemplateWorkbook(sessionDates) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "마음온 출석 돌봄";
  const sheet = workbook.addWorksheet("출석입력", { views: [{ state: "frozen", ySplit: 1, xSplit: 2 }] });
  sheet.columns = [
    { header: "참여자번호", key: "id", width: 15 },
    { header: "이름", key: "name", width: 13 },
    { header: "나이", key: "age", width: 9 },
    { header: "성별", key: "gender", width: 9 },
    { header: "연락처", key: "phone", width: 18 },
    { header: "보호자/관계", key: "guardian", width: 18 },
    { header: "등록일", key: "joinedAt", width: 14 },
    { header: "담당자메모", key: "memo", width: 35 },
    ...sessionDates.map((date, index) => ({ header: date, key: `attendance_${index}`, width: 14 }))
  ];
  const header = sheet.getRow(1);
  header.height = 28;
  header.font = { bold: true, color: { argb: "FFFFFFFF" } };
  header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4E7968" } };
  header.alignment = { vertical: "middle", horizontal: "center" };
  sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1001, column: sheet.columnCount } };
  for (let row = 2; row <= 1001; row += 1) {
    sheet.getCell(row, 7).numFmt = "yyyy-mm-dd";
    for (let column = 9; column <= sheet.columnCount; column += 1) {
      sheet.getCell(row, column).dataValidation = {
        type: "list",
        allowBlank: false,
        formulae: ['"출석,지각,결석,인정결석"'],
        showErrorMessage: true,
        errorTitle: "출석 상태 확인",
        error: "출석, 지각, 결석, 인정결석 중 하나를 선택하세요."
      };
      sheet.getCell(row, column).alignment = { horizontal: "center" };
    }
  }
  const guide = workbook.addWorksheet("작성안내");
  guide.columns = [{ width: 24 }, { width: 82 }];
  [
    ["마음온 출석 데이터 작성 안내", "양식의 열 제목과 시트 이름은 변경하지 마세요."],
    ["필수 항목", "참여자번호, 이름, 나이, 등록일, 모든 회차의 출석 상태"],
    ["출석 상태", "출석 / 지각 / 결석 / 인정결석 중 하나를 선택하세요."],
    ["개인정보", "실제 자료를 사용할 경우 기관의 개인정보 처리 기준과 정보주체 동의를 확인하세요."],
    ["AI 전송 범위", "JEV 분석에는 이름과 연락처를 제외하고 참여자번호, 출석 요약, 담당자 메모만 전송됩니다."],
    ["업로드 제한", "최대 1,000명, XLSX 파일 10MB 이하"]
  ].forEach((values) => guide.addRow(values));
  guide.getRow(1).font = { bold: true, size: 15, color: { argb: "FF355F50" } };
  guide.eachRow((row) => {
    row.alignment = { vertical: "top", wrapText: true };
    row.height = 35;
  });
  return workbook;
}

export async function generateTemplateBuffer(sessionDates) {
  const workbook = await createTemplateWorkbook(sessionDates);
  return Buffer.from(await workbook.xlsx.writeBuffer());
}
