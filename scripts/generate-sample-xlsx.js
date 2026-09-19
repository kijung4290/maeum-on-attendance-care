import ExcelJS from "exceljs";
import { mkdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const outputDir = `${root}/sample-data`;
const outputPath = `${outputDir}/가상_어르신_출석데이터_1000명.xlsx`;
const dates = ["2026-08-11", "2026-08-14", "2026-08-18", "2026-08-21", "2026-08-25", "2026-08-28", "2026-09-01", "2026-09-04", "2026-09-08", "2026-09-11", "2026-09-15", "2026-09-18"];
const patterns = [
  ["출석","출석","출석","출석","출석","출석","출석","출석","출석","출석","출석","출석"],
  ["출석","지각","출석","출석","출석","출석","지각","출석","출석","출석","지각","출석"],
  ["출석","출석","결석","출석","출석","출석","출석","결석","출석","출석","출석","출석"],
  ["출석","출석","출석","출석","출석","출석","출석","출석","지각","출석","출석","결석"],
  ["출석","출석","출석","출석","출석","출석","출석","출석","출석","인정결석","인정결석","출석"],
  ["출석","출석","출석","출석","출석","출석","지각","출석","결석","출석","결석","출석"],
  ["출석","출석","출석","출석","출석","지각","출석","출석","지각","출석","지각","결석"],
  ["출석","출석","출석","출석","출석","출석","출석","출석","출석","결석","결석","결석"],
  ["출석","출석","출석","출석","출석","출석","출석","출석","결석","출석","결석","결석"],
  ["출석","출석","출석","지각","출석","출석","출석","출석","출석","출석","출석","출석"]
];
const memos = [
  "가상 데이터: 안정적으로 참여 중입니다.", "가상 데이터: 교통 사정으로 간헐적 지각이 있습니다.",
  "가상 데이터: 가족 일정으로 간헐적 결석이 있었습니다.", "가상 데이터: 최근 회차 결석 사유 확인이 필요합니다.",
  "가상 데이터: 사전 연락한 일정으로 인정결석 처리했습니다.", "가상 데이터: 최근 출석 흐름이 다소 불규칙합니다.",
  "가상 데이터: 이동 지원 필요 여부 확인이 필요합니다.", "가상 데이터: 최근 3회 연속 결석하여 안부 확인이 필요합니다.",
  "가상 데이터: 최근 결석이 증가했습니다.", "가상 데이터: 특별한 변동 없이 참여 중입니다."
];

await mkdir(outputDir, { recursive: true });
const workbook = new ExcelJS.Workbook();
workbook.creator = "마음온 출석 돌봄";
workbook.created = new Date("2026-09-19T00:00:00+09:00");
const sheet = workbook.addWorksheet("출석입력", { views: [{ state: "frozen", ySplit: 1, xSplit: 2 }] });
sheet.columns = [
  { header: "참여자번호", key: "id", width: 15 }, { header: "이름", key: "name", width: 18 },
  { header: "나이", key: "age", width: 9 }, { header: "성별", key: "gender", width: 9 },
  { header: "연락처", key: "phone", width: 18 }, { header: "보호자/관계", key: "guardian", width: 18 },
  { header: "등록일", key: "joinedAt", width: 14 }, { header: "담당자메모", key: "memo", width: 42 },
  ...dates.map((date, index) => ({ header: date, key: `attendance_${index}`, width: 14 }))
];
const header = sheet.getRow(1);
header.height = 30;
header.font = { bold: true, color: { argb: "FFFFFFFF" } };
header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4E7968" } };
header.alignment = { vertical: "middle", horizontal: "center" };
sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1001, column: sheet.columnCount } };

for (let index = 0; index < 1000; index += 1) {
  const number = index + 1;
  const patternIndex = index % patterns.length;
  const row = sheet.addRow({
    id: `SYN-${String(number).padStart(4, "0")}`,
    name: `가상참여자${String(number).padStart(4, "0")}`,
    age: 65 + ((index * 7) % 31),
    gender: index % 2 === 0 ? "여" : "남",
    phone: `010-****-${String(1000 + ((index * 37) % 9000)).padStart(4, "0")}`,
    guardian: index % 11 === 0 ? "본인" : `가상 보호자 / ${index % 2 === 0 ? "자녀" : "배우자"}`,
    joinedAt: `2026-${String(1 + (index % 7)).padStart(2, "0")}-${String(1 + (index % 27)).padStart(2, "0")}`,
    memo: memos[patternIndex],
    ...Object.fromEntries(patterns[patternIndex].map((status, attendanceIndex) => [`attendance_${attendanceIndex}`, status]))
  });
  row.height = 22;
  row.getCell(7).numFmt = "yyyy-mm-dd";
  for (let column = 9; column <= sheet.columnCount; column += 1) {
    row.getCell(column).dataValidation = { type: "list", allowBlank: false, formulae: ['"출석,지각,결석,인정결석"'] };
    row.getCell(column).alignment = { horizontal: "center" };
    const color = { "출석": "FFE8F1EB", "지각": "FFFFF0CF", "결석": "FFF9DFDB", "인정결석": "FFE3E8EC" }[row.getCell(column).value];
    row.getCell(column).fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } };
  }
}

const guide = workbook.addWorksheet("작성안내");
guide.columns = [{ width: 25 }, { width: 88 }];
[
  ["마음온 가상 데이터 안내", "이 파일의 1,000명은 모두 테스트용 가상 참여자이며 실존 인물과 무관합니다."],
  ["사용 방법", "로컬 마음온 웹앱의 참여자 관리 → 작성 파일 업로드에서 이 파일을 선택하세요."],
  ["출석 상태", "출석 / 지각 / 결석 / 인정결석 중 하나로 구성되어 있습니다."],
  ["분석 실행", "업로드 후 AI로 전체 다시 분석을 누르면 JEV API 호출이 발생합니다."],
  ["주의", "1,000명 전체 JEV 분석은 1,000회의 API 판정을 실행하므로 사용량과 비용을 먼저 확인하세요."]
].forEach((values) => guide.addRow(values));
guide.getRow(1).font = { bold: true, size: 15, color: { argb: "FF355F50" } };
guide.eachRow((row) => { row.alignment = { vertical: "top", wrapText: true }; row.height = 36; });

await workbook.xlsx.writeFile(outputPath);
console.log(outputPath);
