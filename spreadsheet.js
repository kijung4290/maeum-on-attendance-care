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
  if (member.attendance.length !== expectedAttendanceCount || member.attendance.some((status) => !status)) errors.push(`${rowNumber}행: 모든 출석일에 출석, 지각, 결석, 인정결석 중 하나를 입력하세요.`);
  return errors;
}
