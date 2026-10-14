// L8 -thing + 형용사 문법 학습지 생성기
// 사용법: node make_worksheet.js [student|teacher] [출력경로.docx]
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType, VerticalAlign,
} = require("docx");

const VARIANT = process.argv[2] === "teacher" ? "teacher" : "student";
const OUT = process.argv[3] || `L8_thing+형용사_${VARIANT === "teacher" ? "교사용" : "학생용"}.docx`;
const TEACHER = VARIANT === "teacher";

// ── 설정 ──────────────────────────────────────────────
const LESSON = "Lesson 8. Life in Space"; // 단원명 확인 필요 시 여기만 수정
const BADGE = "-thing + 형용사";
const PAGE_START = 1; // 페이지 라벨 시작 번호 (현재완료 학습지에 이어 붙이면 변경)

const FONT = "Noto Sans CJK KR";
const NAVY = "2E4057", ACCENT = "4A90D9", RED = "E74C3C", GREEN = "27AE60",
  ORANGE = "F39C12", PURPLE = "7B68EE", GRAY = "777777", DARK = "222222";
const W = 9026; // 콘텐츠 폭(DXA)

// ── 기본 헬퍼 ─────────────────────────────────────────
const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const line = (color, size = 6) => ({ style: BorderStyle.SINGLE, size, color });
const allBorders = (color, size = 6) => ({
  top: line(color, size), bottom: line(color, size), left: line(color, size), right: line(color, size),
});

function run(text, o = {}) {
  return new TextRun({
    text, font: FONT, size: o.size || 20, bold: o.bold, italics: o.italics,
    color: o.color || DARK, underline: o.underline ? {} : undefined,
  });
}

// 빈칸: 학생용 = 밑줄 공백, 교사용 = 파란 정답
function blank(answer, o = {}) {
  if (TEACHER) return run(` ${answer} `, { ...o, bold: true, color: ACCENT, underline: true });
  const w = [...answer].reduce((a, c) => a + (/[\u3131-\uD7A3]/.test(c) ? 4.5 : 2.4), 0);
  const n = Math.max(10, Math.round(w) + 4);
  return run(" ".repeat(n), { ...o, underline: true });
}

// 간단 마크업: **굵게**  [[빈칸정답]]  <r>빨강</r> <g>초록</g> <a>파랑</a> <o>주황</o> <p>보라</p> <u>밑줄</u> <s>회색</s>
const COLORS = { r: RED, g: GREEN, a: ACCENT, o: ORANGE, p: PURPLE };
function rich(str, base = {}) {
  const out = [];
  const re = /(\*\*.+?\*\*|\[\[.+?\]\]|<([rgaopus])>.+?<\/\2>)/g;
  let last = 0, m;
  while ((m = re.exec(str))) {
    if (m.index > last) out.push(run(str.slice(last, m.index), base));
    const t = m[0];
    if (t.startsWith("**")) out.push(...rich(t.slice(2, -2), { ...base, bold: true }));
    else if (t.startsWith("[[")) out.push(blank(t.slice(2, -2), base));
    else {
      const tag = m[2], inner = t.slice(3, -4);
      if (tag === "u") out.push(run(inner, { ...base, underline: true }));
      else if (tag === "s") out.push(run(inner, { ...base, color: GRAY }));
      else out.push(run(inner, { ...base, bold: true, color: COLORS[tag] }));
    }
    last = m.index + t.length;
  }
  if (last < str.length) out.push(run(str.slice(last), base));
  return out;
}

function P(str, o = {}) {
  return new Paragraph({
    children: typeof str === "string" ? rich(str, o) : str,
    alignment: o.align, indent: o.indent ? { left: o.indent } : undefined,
    spacing: { before: o.before || 0, after: o.after ?? 40, line: o.line || 276 },
    keepNext: o.keepNext,
  });
}
const gap = (pt = 80) => new Paragraph({ children: [run("", { size: 4 })], spacing: { before: 0, after: pt } });
const pageBreak = () => new Paragraph({ pageBreakBefore: true, children: [run("", { size: 2 })], spacing: { before: 0, after: 0 } });

function cell(children, o = {}) {
  return new TableCell({
    children: children.length ? children : [P("")],
    width: { size: o.width || W, type: WidthType.DXA },
    shading: o.fill ? { type: ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
    borders: o.borders || noBorders,
    margins: { top: o.mt ?? 80, bottom: o.mb ?? 80, left: o.ml ?? 140, right: o.mr ?? 140 },
    verticalAlign: o.valign || VerticalAlign.CENTER,
    columnSpan: o.span,
  });
}
function table(rows, widths) {
  return new Table({ rows, width: { size: W, type: WidthType.DXA }, columnWidths: widths });
}
// 한 칸짜리 박스
function box(children, fill, border, o = {}) {
  return table([new TableRow({ children: [cell(children, { fill, borders: border ? allBorders(border, o.bsize || 8) : noBorders, mt: 110, mb: 110, ml: 200, mr: 200 })] })], [W]);
}

// ── 하우스 포맷 블록 ─────────────────────────────────
function headerBar() {
  return table([new TableRow({ children: [
    cell([P(`${LESSON}   Language in Use`, { bold: true, color: "FFFFFF", size: 24, after: 0 })], { width: 6626, fill: NAVY, mt: 120, mb: 120, ml: 220 }),
    cell([P(BADGE, { bold: true, color: NAVY, size: 22, align: AlignmentType.CENTER, after: 0 })], { width: 2400, fill: "FFFFFF", borders: allBorders(NAVY, 18), mt: 120, mb: 120 }),
  ] })], [6626, 2400]);
}
function infoBar(pageNo) {
  const label = `${PAGE_START + pageNo - 1}p`;
  return table([new TableRow({ children: [
    cell([P("Class ______   No. ______   Date 2026.   .   .   Name ________________", { size: 19, after: 0 })], { width: 7826, fill: "F5F5F5", mt: 90, mb: 90, ml: 220 }),
    cell([P(label, { bold: true, color: ACCENT, size: 22, align: AlignmentType.RIGHT, after: 0 })], { width: 1200, fill: "F5F5F5", mt: 90, mb: 90 }),
  ] })], [7826, 1200]);
}
function pageTop(pageNo) {
  return [headerBar(), gap(60), infoBar(pageNo), gap(140)];
}
function section(num, title, sub) {
  return P(`<a>${num}</a>  **${title}**` + (sub ? `  <s>${sub}</s>` : ""), { size: 25, before: 60, after: 90, keepNext: true });
}
// 예문 박스: [[en, ko], ...]
function examples(pairs, o = {}) {
  const ch = [];
  pairs.forEach(([en, ko], i) => {
    ch.push(P(en, { size: 21, after: 0 }));
    ch.push(P(`→ ${ko}`, { size: 18, color: GRAY, after: i === pairs.length - 1 ? 0 : 90, indent: 200 }));
  });
  return box(ch, o.fill || "F5F5F5", o.border);
}
function formula(text, sub) {
  const ch = [P(text, { size: 30, bold: true, color: ACCENT, align: AlignmentType.CENTER, after: sub ? 40 : 0 })];
  if (sub) ch.push(P(sub, { size: 19, color: GRAY, align: AlignmentType.CENTER, after: 0 }));
  return box(ch, "E8F0FE", ACCENT, { bsize: 12 });
}
// 일반 표 (헤더 행 포함)
function grid(head, rows, widths, o = {}) {
  const hdr = new TableRow({ tableHeader: true, children: head.map((h, i) =>
    cell([P(h, { bold: true, color: "FFFFFF", size: 19, align: AlignmentType.CENTER, after: 0 })], { width: widths[i], fill: o.headFill || NAVY, borders: allBorders("BBBBBB", 4), mt: 60, mb: 60 })) });
  const body = rows.map((r, ri) => new TableRow({ children: r.map((c, i) =>
    cell([].concat(c).map((s) => P(s, { size: 19, after: 0, align: i === 0 && o.centerFirst ? AlignmentType.CENTER : undefined })),
      { width: widths[i], fill: ri % 2 ? "FAFAFA" : "FFFFFF", borders: allBorders("BBBBBB", 4), mt: 60, mb: 60, ml: 110, mr: 110 })) }));
  return table([hdr, ...body], widths);
}

// ── 문제 헬퍼 ─────────────────────────────────────────
const LINE = "______________________________________________";
function qTitle(letter, text, pts) {
  return P(`<a>${letter}</a>  **${text}**` + (pts ? `  <s>${pts}</s>` : ""), { size: 22, before: 140, after: 80, keepNext: true });
}
const q = (n, str, o = {}) => P(`${n}. ${str}`, { size: 20, indent: 240, after: o.after ?? 70, line: 300 });
const ans = (pre = "→", tail = "") => P(`${pre} ${LINE}${tail}`, { size: 20, indent: 560, after: 110, color: GRAY });

// ═════════════════════════════════════════════════════
// 1페이지: Warm Up → ① 핵심 개념 → ② 형태와 쓰임
// ═════════════════════════════════════════════════════
const page1 = [
  ...pageTop(1),
  box([
    P("<o>☀ Warm Up</o>  **형용사는 어디에 올까?**", { size: 23, after: 80 }),
    P("형용사는 명사를 꾸며 줘요. 지금까지는 늘 **명사 [[앞]]** 에 썼죠.", { after: 40 }),
    P("a <a>big</a> sign (큰 표지판)      a <a>huge</a> dinosaur (거대한 공룡)      <a>funny</a> pictures (재미있는 그림)", { indent: 200, after: 100 }),
    P("**Q. ‘아이들이 무언가 재미있는 것을 그리고 있다.’** 바른 문장에 ✔, 틀린 문장에 ✘ 해 보세요.", { after: 50 }),
    P("(a) The kids are drawing <u>funny pictures</u>.                     (      )", { indent: 200, after: 30 }),
    P("(b) The kids are drawing <u>funny something</u>.                  (      )", { indent: 200, after: 30 }),
    P("(c) The kids are drawing <u>something funny</u>.                  (      )", { indent: 200, after: 90 }),
    P("💡 **발견!** something처럼 -thing으로 끝나는 말은 형용사가 [[뒤]] 에서 꾸며 준다.", { color: DARK, after: 40 }),
    P("<s>※ 이미 만난 문장: I saw something strange. (L6 지각동사) / Have you ever seen anything strange? (L8 현재완료)</s>", { size: 18, after: 0 }),
  ], "FFF8E7", ORANGE),
  gap(110),

  section("①", "핵심 개념 : 형용사의 두 자리"),
  P("• 일반 명사 : 형용사가 명사 **[[앞]]** 에서 꾸민다.   a <a>cold</a> drink, an <a>interesting</a> book", { indent: 120 }),
  P("• -thing으로 끝나는 **[[부정대명사]]** (something, anything, nothing, everything) : 형용사가 **[[뒤]]** 에서 꾸민다.", { indent: 120, after: 100 }),
  formula("-thing  +  형용사", "something cold  /  anything strange  /  nothing special  /  everything new"),
  gap(60),
  P("<s>※ 부정대명사 = ‘정해지지 않은’ 사람·사물을 가리키는 대명사 (‘부정(否定)’이 아니라 ‘부정(不定)’!)</s>", { size: 18, indent: 120, after: 80 }),

  section("②", "형태와 쓰임 : -thing 4형제"),
  grid(["단어", "뜻", "주로 쓰는 문장"], [
    ["**something**", "무언가, 어떤 것", "[[긍정문]]"],
    ["**anything**", "무엇이든, 아무것도", "[[부정문]] · [[의문문]]"],
    ["**nothing**", "아무것도 ~ 않다", "그 자체로 부정 (= not ~ [[anything]])"],
    ["**everything**", "모든 것", "긍정문"],
  ], [2000, 3000, 4026], { centerFirst: true }),
  gap(80),
  examples([
    ["I saw <u>something strange</u> in the dark.", "나는 어둠 속에서 이상한 것을 봤다."],
    ["I don’t want you to do <u>anything foolish</u>.", "나는 네가 어리석은 일을 하지 않기를 바란다."],
    ["There is <u>nothing wrong</u> with the car. (= There isn’t <u>anything wrong</u> with the car.)", "그 차에는 잘못된 것이 아무것도 없다."],
    ["We did <u>everything possible</u> to help him.", "우리는 그를 돕기 위해 가능한 모든 것을 했다."],
  ]),
  gap(80),
  box([
    P("<o>★</o> **Tip**  권유·부탁하는 의문문에는 anything이 아니라 **something**을 써요.", { after: 30 }),
    P("Would you like <u>something cold</u> to drink? (차가운 것 좀 마실래?)", { indent: 280, after: 30 }),
    P("<o>★</o> **Tip**  -thing은 **[[단수]]** 취급 → <u>Something strange</u> <a>was</a> in the box. (were ✘)", { after: 0 }),
  ], "FFF3CD", ORANGE),
];

// ═════════════════════════════════════════════════════
// 2페이지: ③ +to부정사 → ④ 비교표 → ⑤ 시험 포인트
// ═════════════════════════════════════════════════════
const page2 = [
  pageBreak(),
  ...pageTop(2),
  section("③", "-thing + 형용사 + to부정사", "‘~할 …한 것’"),
  P("to부정사도 명사를 **뒤에서** 꾸며요(형용사적 용법). 둘 다 붙을 때는 **형용사가 먼저, to부정사가 나중!**", { indent: 120, after: 100 }),
  formula("-thing  +  형용사  +  to부정사", "무엇(-thing)  →  어떤(형용사)  →  ~할(to부정사)"),
  gap(80),
  examples([
    ["I need something cold to drink.", "나는 마실 차가운 것이 필요해."],
    ["Do you have anything interesting to read?", "재미있는 읽을거리가 있니?"],
    ["I have something important to tell you.", "나는 너에게 말할 중요한 것이 있어."],
  ]),
  gap(80),
  table([new TableRow({ children: [
    cell([P("✔  something **cold** <a>to drink</a>", { size: 20, color: GREEN, bold: true, align: AlignmentType.CENTER, after: 0 })], { width: 3008, fill: "E8F5E9", mt: 70, mb: 70 }),
    cell([P("✘  cold something to drink", { size: 20, color: RED, align: AlignmentType.CENTER, after: 0 })], { width: 3009, fill: "FDE8E8", mt: 70, mb: 70 }),
    cell([P("✘  something to drink cold", { size: 20, color: RED, align: AlignmentType.CENTER, after: 0 })], { width: 3009, fill: "FDE8E8", mt: 70, mb: 70 }),
  ] })], [3008, 3009, 3009]),
  gap(160),

  section("④", "한눈에 비교 : 일반 명사 vs -thing"),
  grid(["구분", "일반 명사", "-thing 부정대명사"], [
    ["형용사 위치", "명사 [[앞]]", "-thing [[뒤]]"],
    ["예", "a **funny** picture / **cold** water", "something **funny** / nothing **cold**"],
    ["+ to부정사", "a **good** book <a>to read</a>", "something **good** <a>to read</a>"],
  ], [1800, 3613, 3613], { centerFirst: true }),
  gap(80),
  box([
    P("<p>＋ 더 알아보기</p>  -one, -body로 끝나는 말(someone, anybody, nobody, everyone 등)도 형용사가 뒤에 와요.", { after: 30 }),
    P("I want to meet <u>someone kind</u>. (나는 친절한 누군가를 만나고 싶다.)", { indent: 280, after: 0 }),
  ], "F3E8FD", PURPLE),
  gap(160),

  section("⑤", "시험 포인트", "이렇게 틀려요!"),
  table([
    ["어순", "something **sweet**", "sweet something"],
    ["형용사 자리에 부사 ✘", "something **special**", "something specially"],
    ["형용사 형태 (명사·동사 ✘)", "anything **interesting**", "anything interest"],
    ["형용사 앞에 to ✘", "anything **wrong**", "anything to wrong"],
    ["3요소 순서", "something **important** to say", "something to say important"],
    ["이중 부정 ✘", "I didn’t do **anything** wrong. = I did **nothing** wrong.", "I didn’t do nothing wrong."],
    ["단수 취급", "Something strange **was** there.", "Something strange were there."],
  ].map(([k, o, x]) => new TableRow({ children: [
    cell([P(`**${k}**`, { size: 19, after: 0 })], { width: 2300, fill: "FDE8E8", borders: allBorders("F5B7B1", 4), mt: 55, mb: 55 }),
    cell([P(`<g>✔</g>  ${o}`, { size: 19, after: 0 })], { width: 3863, fill: "FFFFFF", borders: allBorders("F5B7B1", 4), mt: 55, mb: 55 }),
    cell([P(`<r>✘</r>  ${x}`, { size: 19, after: 0, color: GRAY })], { width: 2863, fill: "FFFFFF", borders: allBorders("F5B7B1", 4), mt: 55, mb: 55 }),
  ] })), [2300, 3863, 2863]),
];

// ═════════════════════════════════════════════════════
// 3~4페이지: 확인문제
// ═════════════════════════════════════════════════════
const page3 = [
  pageBreak(),
  ...pageTop(3),
  P("**✎ 확인문제**  <s>-thing + 형용사 (+ to부정사)</s>", { size: 26, color: NAVY, after: 60 }),

  qTitle("A", "괄호 안에서 알맞은 것을 고르시오."),
  q(1, "There is ( wrong nothing  /  nothing wrong ) with the car."),
  q(2, "Did you buy ( anything good  /  good anything ) at the store?"),
  q(3, "I have ( something important  /  important something ) to tell you."),
  q(4, "Would you like ( cold something  /  something cold ) to drink?"),
  q(5, "Something strange ( was  /  were ) in the box."),
  q(6, "I didn’t eat ( anything  /  nothing ) spicy today."),
  q(7, "She wants ( something  /  anything ) sweet for dessert."),
  q(8, "Do you have ( anything to read interesting  /  anything interesting to read )?", { after: 40 }),

  qTitle("B", "밑줄 친 부분을 바르게 고쳐 쓰시오."),
  ...[
    "I want <u>sweet something</u> after lunch.",
    "She heard <u>strange something</u> last night.",
    "Is there <u>anything to wrong</u> with my computer?",
    "Did you find <u>anything interest</u> in his room?",
    "They bought <u>something specially</u> for you.",
  ].map((s, i) => P(`${i + 1}. ${s}     →  ____________________________`, { size: 20, indent: 240, after: 90, line: 300 })),

  qTitle("C", "어법상 옳으면 O, 틀리면 X에 표시하고, 틀린 부분을 바르게 고치시오."),
  ...[
    "I want to eat something spicy.",
    "Did you learn useful anything from this book?",
    "Have you ever seen anything strange in the sky?",
    "He has something to say important.",
    "There was new nothing about his idea.",
    "I don’t have nothing special to do today.",
  ].map((s, i) => table([new TableRow({ children: [
    cell([P(`${i + 1}. ${s}`, { size: 20, after: 0 })], { width: 5026, mt: 45, mb: 45, ml: 240 }),
    cell([P("( O / X )", { size: 20, after: 0, align: AlignmentType.CENTER })], { width: 1100, mt: 45, mb: 45, ml: 40, mr: 40 }),
    cell([P("→ __________________", { size: 20, after: 0, color: GRAY })], { width: 2900, mt: 45, mb: 45, ml: 60 }),
  ] })], [5026, 1100, 2900])),
];

const page4 = [
  pageBreak(),
  ...pageTop(4),
  qTitle("D", "괄호 안의 말을 바르게 배열하여 문장을 완성하시오."),
  q(1, "They ______________________________________________ .  (brought, something, us, delicious, for)"),
  q(2, "Did ______________________________________ ?  (say, anything, he, important, at the meeting)"),
  q(3, "I ________________________________________________ .  (nothing, helpful, found, on that website)"),
  q(4, "Do you have ______________________________________ ?  (to, interesting, read, anything)"),
  q(5, "I ________________________________________________ .  (something, need, warm, wear, to)", { after: 40 }),

  qTitle("E", "우리말과 같도록 괄호 안의 말을 이용하여 영작하시오.", "(필요하면 형태를 바꿀 것)"),
  q(1, "나는 어둠 속에서 무언가 이상한 것을 봤다. (see, strange, in the dark)"), ans(),
  q(2, "오늘은 TV에 재미있는 것이 없다. (there, nothing, interesting)"), ans(),
  q(3, "그는 건강에 좋지 않은 어떤 것도 절대 먹지 않는다. (never, anything, unhealthy)"), ans(),
  q(4, "나는 마실 차가운 것이 필요하다. (need, cold, drink)"), ans(),
  q(5, "너는 우주에서 놀라운 것을 본 적이 있니? (ever, see, amazing, in space)"), ans(),

  qTitle("F", "대화를 읽고, 괄호 안의 말을 바르게 배열하여 대화를 완성하시오."),
  q(1, "A: What do you think is in the box?", { after: 20 }),
  P("B: I don’t know. But ________________________ in the box.  (is, heavy, something)", { size: 20, indent: 560, after: 90 }),
  q(2, "A: Let’s ________________________ for lunch.  (something, eat, delicious)", { after: 20 }),
  P("B: Sounds good. I’m so hungry.", { size: 20, indent: 560, after: 90 }),
  q(3, "A: Did you find ________________________ ?  (anything, around here, strange)", { after: 20 }),
  P("B: No, everything looks fine.", { size: 20, indent: 560, after: 90 }),
  q(4, "A: I’m so bored. What should we do this weekend?", { after: 20 }),
  P("B: Let’s ________________________ .  (to do, find, exciting, something)", { size: 20, indent: 560, after: 90 }),

  qTitle("G", "나만의 문장 쓰기", "-thing + 형용사 (+ to부정사)를 넣어 우주 생활에 관한 문장을 쓰시오."),
  box([
    P("e.g.  I want to eat **something delicious** in space.  /  I want to see **something beautiful** from space.", { size: 19, color: GRAY, after: 80 }),
    P(`1. ${LINE}______________`, { after: 80 }),
    P(`2. ${LINE}______________`, { after: 0 }),
  ], "E3F2FD"),
];

// ═════════════════════════════════════════════════════
// 교사용: 정답·해설
// ═════════════════════════════════════════════════════
const ANS = [
  ["A", [
    "1. nothing wrong  (-thing + 형용사)", "2. anything good  (의문문 → anything, 형용사는 뒤)",
    "3. something important  (-thing + 형용사 + to부정사)", "4. something cold  (권유 의문문 → something)",
    "5. was  (-thing은 단수 취급)", "6. anything  (didn’t가 이미 부정 → nothing 쓰면 이중 부정)",
    "7. something  (긍정문 → something)", "8. anything interesting to read  (형용사 먼저, to부정사 나중)",
  ]],
  ["B", [
    "1. something sweet", "2. something strange  (L6 지각동사 heard + 목적어)", "3. anything wrong  (형용사 앞 to 불필요)",
    "4. anything interesting  (명사 interest ✘ → 형용사)", "5. something special  (부사 specially ✘ → 형용사)",
  ]],
  ["C", [
    "1. O", "2. X → anything useful", "3. O  (현재완료 경험 + anything strange)",
    "4. X → something important to say", "5. X → nothing new",
    "6. X → anything special  (don’t + nothing = 이중 부정)  ※ I have nothing special to do today. 도 정답",
  ]],
  ["D", [
    "1. They brought something delicious for us.", "2. Did he say anything important at the meeting?",
    "3. I found nothing helpful on that website.", "4. Do you have anything interesting to read?",
    "5. I need something warm to wear.",
  ]],
  ["E", [
    "1. I saw something strange in the dark.", "2. There is nothing interesting on TV today.",
    "3. He never eats anything unhealthy.  (never가 부정어 → anything)", "4. I need something cold to drink.",
    "5. Have you ever seen anything amazing in space?  (현재완료 경험, 의문문 → anything)",
  ]],
  ["F", [
    "1. something heavy is  (B: But something heavy is in the box.)", "2. eat something delicious",
    "3. anything strange around here", "4. find something exciting to do",
  ]],
  ["G", ["예시) I want to find something interesting to do on the space station. / Did the astronauts bring anything special? 등 -thing 뒤 형용사 위치 확인"]],
];
const answerPage = TEACHER ? [
  pageBreak(),
  P("**[정답 및 해설]**  <s>교사용</s>", { size: 24, color: NAVY, after: 120 }),
  ...ANS.flatMap(([k, list]) => [
    P(`<a>${k}</a>`, { size: 21, before: 60, after: 30 }),
    ...list.map((s) => P(s, { size: 19, indent: 240, after: 20, color: "444444" })),
  ]),
  gap(100),
  box([
    P("**지도 포인트**", { size: 20, after: 40 }),
    P("• -thing 부정대명사는 형용사가 뒤에서 수식 / -thing + 형용사 + to부정사 순서", { size: 19, after: 20 }),
    P("• something(긍정·권유) / anything(부정·의문) / nothing(= not ~ anything, 이중 부정 금지)", { size: 19, after: 20 }),
    P("• 형용사 자리에 부사(specially)·명사(interest)·to 삽입 오답 빈출, -thing은 단수 동사", { size: 19, after: 0 }),
  ], "E3F2FD"),
] : [];

// ── 문서 생성 ─────────────────────────────────────────
const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 20 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 620, right: 800, bottom: 560, left: 800 } } },
    children: [...page1, ...page2, ...page3, ...page4, ...answerPage],
  }],
});
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("wrote", OUT); });
