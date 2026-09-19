import { readFile, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const path = `${root}/data/dummy-data.json`;
const data = JSON.parse(await readFile(path, "utf8"));
const original = data.members.slice(0, 10);

const surnames = ["강", "고", "권", "김", "남", "문", "박", "배", "백", "서", "손", "송", "신", "안", "양"];
const givenNames = ["경자", "광수", "금례", "동식", "명자", "병호"];
const patterns = [
  { attendance:["present","present","present","present","present","present","present","present","present","present","present","present"], memo:"꾸준히 참여하며 다른 참여자들과도 잘 어울립니다." },
  { attendance:["present","late","present","present","present","present","late","present","present","present","late","present"], memo:"대중교통 사정으로 가끔 지각합니다." },
  { attendance:["present","present","absent","present","present","present","present","absent","present","present","present","present"], memo:"간헐적 결석은 가족 일정으로 확인되었습니다." },
  { attendance:["present","present","present","present","present","present","present","present","late","present","present","absent"], memo:"마지막 회차 결석 사유 확인이 필요합니다." },
  { attendance:["present","present","present","present","present","present","present","present","present","excused","excused","present"], memo:"병원 진료 일정을 사전에 알려주었습니다." },
  { attendance:["present","present","present","present","present","present","late","present","absent","present","absent","present"], memo:"최근 출석 흐름이 다소 불규칙해졌습니다." },
  { attendance:["present","present","present","present","present","late","present","present","late","present","late","absent"], memo:"이동 지원 필요 여부를 확인해 주세요." },
  { attendance:["present","present","present","present","present","present","present","present","present","absent","absent","absent"], memo:"최근 3회 연락 없이 결석했습니다. 안부 확인이 필요합니다." },
  { attendance:["present","present","present","present","present","present","present","present","absent","present","absent","absent"], memo:"최근 결석이 늘었으며 연락이 닿지 않았습니다." },
  { attendance:["present","present","present","late","present","present","present","present","present","present","present","present"], memo:"특이사항 없이 안정적으로 참여 중입니다." }
];

const generated = [];
for (let index = 0; index < 90; index += 1) {
  const number = index + 11;
  const pattern = patterns[index % patterns.length];
  const surname = surnames[index % surnames.length];
  const given = givenNames[Math.floor(index / surnames.length)];
  const gender = ["경자", "금례", "명자"].includes(given) ? "여" : "남";
  generated.push({
    id: `NB-${String(number).padStart(3, "0")}`,
    name: `${surname}${given}`,
    age: 70 + ((index * 7) % 22),
    gender,
    phone: `010-****-${String(3100 + index * 47).slice(-4)}`,
    guardian: index % 9 === 0 ? "본인" : `${index % 2 ? "딸" : "아들"} ${surname}○○`,
    joinedAt: `2026-${String(1 + (index % 7)).padStart(2, "0")}-${String(2 + (index % 25)).padStart(2, "0")}`,
    memo: pattern.memo,
    attendance: [...pattern.attendance]
  });
}

data.members = [...original, ...generated];
await writeFile(path, `${JSON.stringify(data, null, 2)}\n`, "utf8");
console.log(`더미 데이터 ${data.members.length}명을 생성했습니다.`);
