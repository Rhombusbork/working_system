// 발표 슬라이드 v4 — 미리캔버스 「흰색과 하늘색의 심플한 웹사이트 회사소개서」 레이아웃을 따른다
// 계획: 프로젝트의 CLAUDE.md
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("./apply_theme.js");
const OUT = process.argv[2] || "deck.pptx";

// ── 색과 글꼴 ───────────────────────────────────────────
const BLUE = "3D8FE0", BLUE_D = "2B78C9", BLUE_L = "EAF3FC", BLUE_M = "B9D7F4", BLUE_S = "7FB6EC";
const INK = "2B2F36", GRAY = "6B7280", GRAY_L = "C9D1DA", ORANGE = "D9651F", KAKAO = "FEE500", WHITE = "FFFFFF";
const FONT = "Malgun Gothic";
const THEME = { name: "Sky Website", headFontFace: FONT, bodyFontFace: FONT,
  colors: { dk1: INK, lt1: WHITE, dk2: BLUE_D, lt2: BLUE_L, accent1: BLUE, accent2: BLUE_D, accent3: GRAY_L, accent4: ORANGE, accent5: KAKAO, accent6: GRAY, hlink: BLUE_D, folHlink: GRAY } };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "IT 업종 근무제의 네 원칙과 카카오 적용";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
const SH = pres.ShapeType;

const seen = new Set();
function newSlide(master, section) {
  if (!seen.has(section)) { pres.addSection({ title: section }); seen.add(section); }
  return pres.addSlide({ masterName: master, sectionTitle: section });
}
const NOTES = {};
for (const part of fs.readFileSync(path.join(__dirname, "script.md"), "utf8").split(/^## /m).slice(1)) {
  const m = part.match(/^(\d+)\.\s.*\n([\s\S]*)$/);
  if (m) NOTES[+m[1]] = m[2].trim();
}
const notes = (s, n) => { if (NOTES[n]) s.addNotes(NOTES[n]); };

// ── 프레임: 상단 탭 바 + 둥근 본문 카드 ──────────────────
const TABS = ["배경", "일정", "공정성", "측정", "강도", "카카오"];
const TW = 1.62, TX0 = (13.333 - TW * TABS.length) / 2;
function tabs(active) {
  return TABS.map((t, i) => {
    const o = { x: TX0 + i * TW + 0.12, y: 0.2, w: TW - 0.24, h: 0.42, color: WHITE, fontFace: FONT, fontSize: 14, align: "center", valign: "middle", margin: 0 };
    if (i === active) Object.assign(o, { shape: SH.roundRect, rectRadius: 0.21, fill: { color: BLUE_D }, bold: true });
    return { text: { text: t, options: o } };
  });
}
const frame = (active) => [
  { rect: { x: 0, y: 0, w: 13.333, h: 1.2, fill: { color: BLUE } } },
  { text: { text: "", options: { shape: SH.roundRect, rectRadius: 0.35, x: 0, y: 0.88, w: 13.333, h: 6.62, fill: { color: WHITE } } } },
  ...tabs(active),
  { text: { text: "IT 업종 근무제의 네 원칙 · 카카오 적용", options: { x: 0.6, y: 7.06, w: 6, h: 0.26, color: "9AA3AF", fontFace: FONT, fontSize: 9, margin: 0 } } },
];
const SLNUM = { x: 12.23, y: 7.06, w: 0.5, h: 0.26, fontFace: FONT, fontSize: 9, color: "9AA3AF", align: "right" };
for (let a = -1; a < TABS.length; a++) {
  const k = a < 0 ? "N" : a;
  // 가운데 위 제목
  pres.defineSlideMaster({ title: "TOP_" + k, background: { color: WHITE }, slideNumber: SLNUM, objects: [...frame(a),
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.22, w: 12.13, h: 0.66, fontFace: FONT, fontSize: 28, bold: true, color: INK, align: "center", valign: "middle", margin: 0 }, text: "" } }] });
  // 왼쪽 열 제목
  pres.defineSlideMaster({ title: "LEFT_" + k, background: { color: WHITE }, slideNumber: SLNUM, objects: [...frame(a),
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.78, w: 3.65, h: 1.3, fontFace: FONT, fontSize: 30, bold: true, color: INK, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.05 }, text: "" } }] });
  // 섹션 구분 장
  pres.defineSlideMaster({ title: "DIV_" + k, background: { color: WHITE }, slideNumber: SLNUM, objects: [...frame(a),
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 2.55, w: 12.13, h: 1.75, fontFace: FONT, fontSize: 46, bold: true, color: INK, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.05 }, text: "" } }] });
}
pres.defineSlideMaster({ title: "COVER", background: { color: WHITE }, objects: [
  { text: { text: "", options: { shape: SH.roundRect, rectRadius: 0.4, x: 0, y: -0.5, w: 13.333, h: 6.9, fill: { color: BLUE } } } }, ...tabs(-1)] });

// ── 도우미 ─────────────────────────────────────────────
function T(s, text, o) { s.addText(text, Object.assign({ isTextBox: true, fontFace: FONT, color: INK, margin: 0, valign: "top", fontSize: 15 }, o)); }
function card(s, o) { s.addShape(SH.roundRect, Object.assign({ rectRadius: 0.12, fill: { color: BLUE_L }, line: { type: "none" } }, o)); }
function ocard(s, o) { s.addShape(SH.roundRect, Object.assign({ rectRadius: 0.12, fill: { color: WHITE }, line: { color: BLUE_M, width: 1.25 } }, o)); }
function hline(s, x, y, w, c, wt) { s.addShape(SH.line, { x, y, w, h: 0, line: { color: c || GRAY_L, width: wt || 1 } }); }
function vline(s, x, y, h) { s.addShape(SH.line, { x, y, w: 0, h, line: { color: GRAY_L, width: 1 } }); }
function pill(s, text, o) { s.addText(text, Object.assign({ shape: SH.roundRect, rectRadius: 0.2, fill: { color: BLUE_D }, color: WHITE, fontFace: FONT, fontSize: 14, bold: true, align: "center", valign: "middle", margin: 0 }, o)); }
function title(s, parts) { s.addText(parts.map(p => typeof p === "string" ? { text: p } : { text: p.k, options: { color: BLUE_D } }), { placeholder: "title" }); }
// 번호 배지: 테두리만 있는 알약
function badge(s, label, x, y) {
  const w = 0.32 + label.length * 0.13;
  s.addText(label, { shape: SH.roundRect, rectRadius: 0.18, x, y, w, h: 0.36, fill: { color: WHITE }, line: { color: BLUE_D, width: 1.5 }, color: BLUE_D, fontFace: FONT, fontSize: 13, bold: true, align: "center", valign: "middle", margin: 0 });
  return w;
}
// 가운데 위 제목 앞에 배지를 붙인다 (제목 글자 폭을 어림해서)
function textW(str, pt) {
  let em = 0;
  for (const ch of str) em += /[ㄱ-힝]/.test(ch) ? 1.0 : ch === " " ? 0.32 : 0.58;
  return em * pt / 72;
}
function topTitle(s, num, parts, sub) {
  title(s, parts);
  if (num) {
    const w = textW(parts.map(p => typeof p === "string" ? p : p.k).join(""), 28);
    badge(s, num, 6.667 - w / 2 - 0.85, 1.37);
  }
  if (sub) T(s, sub, { x: 0.6, y: 1.9, w: 12.13, h: 0.34, fontSize: 13, color: GRAY, align: "center" });
}
function leftTitle(s, num, parts, desc, descY) {
  badge(s, num, 0.6, 1.32);
  title(s, parts);
  if (desc) T(s, desc, { x: 0.6, y: descY || 3.35, w: 3.55, h: 2.0, fontSize: 12.5, color: GRAY, lineSpacingMultiple: 1.4 });
}
function source(s, text) { T(s, "출처: " + text, { x: 0.6, y: 6.74, w: 12.13, h: 0.24, fontSize: 9, color: "9AA3AF" }); }
// 실제 자료 캡처: 흰 테두리 + 옅은 그림자
function clip(s, file, x, y, w, h, cap) {
  s.addShape(SH.rect, { x: x - 0.03, y: y - 0.03, w: w + 0.06, h: h + 0.06, fill: { color: WHITE }, line: { color: GRAY_L, width: 0.75 }, shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.16 } });
  s.addImage({ path: path.join(__dirname, "img", file), x, y, w, h });
  if (cap) T(s, cap, { x, y: y + h + 0.07, w: Math.max(w, 2.5), h: 0.24, fontSize: 10, color: GRAY });
}
function chartBase(extra) {
  return Object.assign({ barDir: "col", catAxisLabelColor: INK, catAxisLabelFontFace: FONT, catAxisLabelFontSize: 13,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, catAxisLineShow: true, catAxisLineColor: GRAY_L,
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: INK, dataLabelFontFace: FONT, dataLabelFontSize: 16, dataLabelFontBold: true,
    showLegend: false, barGapWidthPct: 70 }, extra);
}
const NONE = { type: "none" }, RULE = { type: "solid", pt: 1, color: GRAY_L };

// ═══ 1. 표지 (L-COVER) ═══════════════════════════════════
{
  const s = newSlide("COVER", "배경");
  [[11.15, 4.7, 1.0], [11.72, 4.0, 1.7], [12.29, 3.2, 2.5]].forEach(([x, y, h], i) => s.addShape(SH.rect, { x, y, w: 0.42, h, fill: { color: WHITE, transparency: 86 }, line: NONE, objectName: "deco-" + i }));
  T(s, "회사들은 왜 재택을 되돌릴까?  그 이유를 풀면서 장점은 지킬 수 없을까?", { x: 1.2, y: 1.7, w: 10.9, h: 0.4, fontSize: 16, color: WHITE, align: "center" });
  T(s, "같은 요일, 공개된 기준,\n측정되는 시간, 상한 있는 출근", { x: 1.2, y: 2.2, w: 10.9, h: 1.85, fontSize: 44, bold: true, color: WHITE, align: "center", valign: "middle", lineSpacingMultiple: 1.05 });
  pill(s, "AI 시대 IT 업종의 근무제 원칙과 카카오 적용", { x: 3.97, y: 4.35, w: 5.4, h: 0.5, fontSize: 15, rectRadius: 0.25 });
  T(s, "[조 이름]   ·   [발표자]   ·   [날짜]", { x: 0.6, y: 6.75, w: 8, h: 0.35, fontSize: 13, color: GRAY });
  notes(s, 1);
}

// ═══ 2. 목차 (L-TOC) ════════════════════════════════════
{
  const s = newSlide("TOP_N", "배경");
  topTitle(s, null, ["발표 ", { k: "목차" }]);
  const items = [["배경", "회사들은 왜 재택을 되돌릴까"], ["일정", "언제 모이나"], ["공정성", "예외는 누가 정하나"], ["측정", "일한 시간은 어떻게 재나"], ["강도", "출근은 어디까지 요구하나"], ["카카오", "원칙을 실제 회사에"]];
  items.forEach(([h, d], i) => {
    const x = i % 2 === 0 ? 0.95 : 6.95, y = 2.55 + Math.floor(i / 2) * 1.3;
    card(s, { x, y, w: 5.45, h: 1.0, rectRadius: 0.5 });
    s.addText(String(i + 1).padStart(2, "0"), { shape: SH.ellipse, x: x + 0.22, y: y + 0.2, w: 0.6, h: 0.6, fill: { color: BLUE_D }, color: WHITE, fontFace: FONT, fontSize: 15, bold: true, align: "center", valign: "middle", margin: 0 });
    T(s, h, { x: x + 1.05, y: y + 0.17, w: 4.2, h: 0.36, fontSize: 19, bold: true });
    T(s, d, { x: x + 1.05, y: y + 0.55, w: 4.2, h: 0.3, fontSize: 13, color: GRAY });
  });
  notes(s, 2);
}

// ═══ 3. 재택을 줄이는 이유 (L-KEY) ═══════════════════════
{
  const s = newSlide("TOP_0", "배경");
  topTitle(s, "01", ["재택을 줄이는 ", { k: "이유" }], "회사들이 내세운 이유를 연구와 맞춰 보니, 적어도 두 가지는 근거가 있었다");
  const stats = [[0.6, "협업", "−25%", "전사 재택 뒤 다른 부서와\n협업하는 시간의 비중", "Microsoft 직원 6만 명의 업무 기록", INK], [4.65, "신입의 성장", "+22%", "같은 건물에 팀원이 있을 때\n받는 코드 피드백", "신입일수록 차이가 컸다", BLUE_D]];
  stats.forEach(([x, h, n, c, note, col]) => {
    card(s, { x, y: 2.55, w: 3.85, h: 4.0 });
    T(s, h, { x: x + 0.35, y: 2.85, w: 3.2, h: 0.4, fontSize: 18, bold: true, color: BLUE_D });
    T(s, n, { x: x + 0.35, y: 3.35, w: 3.3, h: 1.15, fontSize: 64, bold: true, color: col });
    T(s, c, { x: x + 0.35, y: 4.65, w: 3.3, h: 0.75, fontSize: 15, lineSpacingMultiple: 1.25 });
    hline(s, x + 0.35, 5.6, 3.15, BLUE_M);
    T(s, note, { x: x + 0.35, y: 5.75, w: 3.3, h: 0.3, fontSize: 12, color: GRAY });
  });
  T(s, "원문", { x: 8.85, y: 2.6, w: 3.8, h: 0.3, fontSize: 12, bold: true, color: GRAY });
  const w = 3.88; clip(s, "src_nhb.png", 8.85, 2.98, w, w / 4.317, "Yang 외 (2021), Nature Human Behaviour");
  ocard(s, { x: 8.82, y: 4.45, w: 3.94, h: 1.55 });
  T(s, [{ text: "“The Power of Proximity\nto Coworkers”", options: { bold: true, breakLine: true } }, { text: "Emanuel · Harrington · Pallais\n워킹페이퍼 (2023)", options: { color: GRAY, fontSize: 12 } }], { x: 9.05, y: 4.6, w: 3.55, h: 1.3, fontSize: 14, lineSpacingMultiple: 1.2 });
  source(s, "[2-5] Yang 외 (2021) · [2-6] Emanuel·Harrington·Pallais (워킹페이퍼)");
  notes(s, 3);
}

// ═══ 4. 같은 재택, 다른 결과 (L-VS) ══════════════════════
{
  const s = newSlide("LEFT_0", "배경");
  leftTitle(s, "01", ["같은 재택,\n", { k: "다른 결과" }], "그렇다면 재택은\n적을수록 좋을까?\n\n같은 회사가 10여 년 간격으로\n한 두 실험. 결과를 가른 것은\n재택을 했느냐보다\n어떻게 운영했느냐였다.");
  const cols = [[4.55, "Ctrip 콜센터 · 2015", BLUE_S, [["재택 주 4일", INK], ["성과 +13%", INK], ["같은 성과에도 승진 불리", ORANGE]]],
                [8.85, "Trip.com · 2024", BLUE_D, [["주 2일, 수·금 같은 요일", INK], ["퇴사율 7.2% → 4.8%", INK], ["승진 · 평가 차이 없음", BLUE_D]]]];
  cols.forEach(([x, h, hc, rows]) => {
    s.addText(h, { shape: SH.roundRect, rectRadius: 0.08, x, y: 1.35, w: 3.88, h: 0.5, fill: { color: hc }, color: WHITE, fontFace: FONT, fontSize: 15, bold: true, align: "center", valign: "middle", margin: 0 });
    rows.forEach(([t, c], i) => {
      const y = 2.0 + i * 0.86;
      ocard(s, { x, y, w: 3.88, h: 0.72 });
      s.addShape(SH.ellipse, { x: x + 0.2, y: y + 0.24, w: 0.24, h: 0.24, fill: { color: c === ORANGE ? ORANGE : BLUE_D }, line: NONE });
      T(s, t, { x: x + 0.58, y, w: 3.2, h: 0.72, fontSize: 16, bold: true, color: c, valign: "middle" });
    });
  });
  s.addText("VS", { shape: SH.ellipse, x: 8.37, y: 2.75, w: 0.4, h: 0.4, fill: { color: BLUE_D }, color: WHITE, fontFace: FONT, fontSize: 11, bold: true, align: "center", valign: "middle", margin: 0 });
  ocard(s, { x: 4.55, y: 4.72, w: 3.88, h: 1.33 });
  T(s, [{ text: "“Does Working from Home Work?”", options: { bold: true, breakLine: true } }, { text: "Bloom · Liang · Roberts · Ying\nQuarterly Journal of Economics (2015)", options: { color: GRAY, fontSize: 12 } }], { x: 4.8, y: 4.85, w: 3.5, h: 1.1, fontSize: 14, lineSpacingMultiple: 1.2 });
  const w = 3.88; clip(s, "news_nature.png", 8.85, 4.72, w, w / (760 / 274));
  T(s, [{ text: "결과를 가른 것  ", options: { color: GRAY } }, { text: "일정 · 공정성 · 측정 · 강도", options: { bold: true, color: BLUE_D } }], { x: 4.55, y: 6.22, w: 8.2, h: 0.4, fontSize: 18, valign: "middle" });
  source(s, "[2-1] Bloom 외 (2015), QJE · [2-2] Bloom·Han·Liang (2024), Nature (CC BY)");
  notes(s, 4);
}

// ═══ 5. 네 원칙 구분 장 (L-DIV) ══════════════════════════
{
  const s = newSlide("DIV_N", "네 원칙");
  badge(s, "02 – 05", 6.667 - 0.53, 2.05);
  s.addText([{ text: "IT 기업이 따를", options: { color: BLUE_D, breakLine: true } }, { text: "네 원칙" }], { placeholder: "title" });
  card(s, { x: 3.2, y: 4.65, w: 6.93, h: 0.7, rectRadius: 0.35 });
  T(s, "일정   ·   공정성   ·   측정   ·   강도", { x: 3.2, y: 4.65, w: 6.93, h: 0.7, fontSize: 18, bold: true, color: BLUE_D, align: "center", valign: "middle" });
  notes(s, 5);
}

// ═══ 6. 한 주의 달력 ═══════════════════════════════════
{
  const s = newSlide("TOP_N", "네 원칙");
  topTitle(s, null, ["네 원칙을 ", { k: "한 주" }, "로"]);
  ["월", "화", "수", "목", "금"].forEach((d, i) => {
    const on = i < 3, x = 0.6 + i * 2.45;
    T(s, d, { x, y: 2.25, w: 2.3, h: 0.35, fontSize: 16, bold: true, color: GRAY, align: "center" });
    s.addText(on ? "출근" : "선택", { shape: SH.roundRect, rectRadius: 0.12, x, y: 2.65, w: 2.3, h: 1.05, fill: { color: on ? BLUE_D : WHITE }, line: on ? NONE : { color: BLUE_M, width: 1.5, dashType: "dash" }, color: on ? WHITE : BLUE_D, fontFace: FONT, fontSize: 26, bold: true, align: "center", valign: "middle", margin: 0 });
  });
  T(s, "회의 · 코드 리뷰 · 신입 온보딩", { x: 0.6, y: 3.8, w: 7.2, h: 0.4, fontSize: 16, color: BLUE_D, align: "center" });
  T(s, "집중 업무, 재택이든 출근이든", { x: 7.95, y: 3.8, w: 4.75, h: 0.4, fontSize: 16, color: GRAY, align: "center" });
  const P = [["02", "일정", "같은 요일, 1년 단위로 미리 고정"], ["03", "공정성", "예외는 미리 공개한 기준으로만"], ["04", "측정", "출퇴근 시각은 자율, 일한 시간은 기록"], ["05", "강도", "주 3일은 최소가 아니라 상한"]];
  P.forEach(([n, k, t], i) => {
    const x = i % 2 === 0 ? 0.6 : 6.8, y = 4.65 + Math.floor(i / 2) * 0.95;
    card(s, { x, y, w: 5.95, h: 0.78 });
    badge(s, n, x + 0.25, y + 0.21);
    T(s, [{ text: k + "   ", options: { bold: true, color: BLUE_D } }, { text: t }], { x: x + 1.05, y, w: 4.8, h: 0.78, fontSize: 17, valign: "middle" });
  });
  notes(s, 6);
}

// ═══ 7. 같은 요일 (L-GRAPH) ═════════════════════════════
{
  const s = newSlide("LEFT_1", "원칙 1 일정");
  leftTitle(s, "02", ["출근은\n", { k: "같은 요일" }, "에"], "요일을 각자, 또는 팀마다\n고르면 함께 있는 날이\n흩어진다.\n\n그래서 회사 전체가 같은 날\n모이고, 요일은 1년 단위로\n미리 정한다.");
  card(s, { x: 4.55, y: 1.35, w: 8.18, h: 3.3 });
  T(s, "근로자가 매긴 값 · 임금 대비", { x: 4.85, y: 1.55, w: 5, h: 0.3, fontSize: 13, bold: true, color: GRAY });
  s.addChart(pres.ChartType.bar, [{ name: "가치", labels: ["재택 선택권", "갑작스런 일정 피하기"], values: [8, 20] }], chartBase({
    x: 4.75, y: 1.9, w: 5.3, h: 2.6, chartColors: [GRAY_L, BLUE_D], dataLabelFormatCode: "0\"%\"", dataLabelFontSize: 20, catAxisLabelFontSize: 13, valAxisMinVal: 0, valAxisMaxVal: 24 }));
  T(s, "2.5배", { x: 10.3, y: 2.15, w: 2.2, h: 0.8, fontSize: 44, bold: true, color: BLUE_D });
  T(s, "미리 아는 일정에\n재택 선택권의\n2.5배 값을 매겼다", { x: 10.3, y: 2.95, w: 2.3, h: 1.1, fontSize: 14, lineSpacingMultiple: 1.25 });
  [["팀마다 요일을 고르면", "부서 사이 협업이 줄어든다", "재택 뒤 부서 간 협업 비중 −25%"], ["사람마다 요일을 고르면", "덜 보이는 사람이 생긴다", "같은 성과에도 재택자 승진 불리"]].forEach(([a, b, c], i) => {
    const x = 4.55 + i * 4.16;
    ocard(s, { x, y: 4.85, w: 4.02, h: 1.7 });
    T(s, a, { x: x + 0.3, y: 5.0, w: 3.5, h: 0.32, fontSize: 13, color: GRAY });
    T(s, b, { x: x + 0.3, y: 5.33, w: 3.5, h: 0.45, fontSize: 18, bold: true });
    T(s, c, { x: x + 0.3, y: 5.95, w: 3.5, h: 0.3, fontSize: 12, color: BLUE_D });
  });
  source(s, "[3-3] Mas & Pallais (2017), AER, 콜센터 구직자 실험 · [2-5] Yang 외 (2021) · [2-1] Bloom 외 (2015)");
  notes(s, 7);
}

// ═══ 8. 신입과 보안 (L-QUOTE) ════════════════════════════
{
  const s = newSlide("TOP_1", "원칙 1 일정");
  topTitle(s, "02", ["신입 교육과 ", { k: "보안" }], "신입은 출근일에 배우고, 원격 여부는 사람이 아니라 다루는 데이터로 정한다");
  const w = 4.55; clip(s, "src_ibm.png", 0.6, 2.55, w, w / 3.665, "IBM · Ponemon, Cost of a Data Breach (2021.7.28)");
  T(s, [{ text: "+100만 달러", options: { bold: true, fontSize: 24, breakLine: true } }, { text: "원격 요인이 낀 유출의 추가 피해", options: { fontSize: 13, color: GRAY } }], { x: 0.6, y: 4.3, w: 4.5, h: 0.85 });
  T(s, [{ text: "−176만 달러", options: { bold: true, fontSize: 24, color: BLUE_D, breakLine: true } }, { text: "제로트러스트를 갖춘 조직의 피해 감소", options: { fontSize: 13, color: GRAY } }], { x: 0.6, y: 5.3, w: 4.5, h: 0.85 });
  card(s, { x: 5.55, y: 2.55, w: 7.18, h: 4.05 });
  T(s, "보안은 다루는 데이터의 등급으로", { x: 5.85, y: 2.75, w: 6.6, h: 0.38, fontSize: 17, bold: true });
  const hd = (t) => ({ text: t, options: { bold: true, color: GRAY, fontSize: 12, border: [NONE, NONE, RULE, NONE] } });
  const c = (t, o) => ({ text: t, options: Object.assign({ border: [NONE, NONE, RULE, NONE] }, o) });
  s.addTable([[hd("등급"), hd("예"), hd("어디서")],
    [c("고위험", { bold: true, color: BLUE_D }), c("개인정보 원본, 핵심 모델", { color: GRAY }), c("원격 금지, 출근일에", { bold: true })],
    [c("내부 기밀", { bold: true }), c("사내 문서, 코드", { color: GRAY }), c("제로트러스트 원격", { bold: true })],
    [c("일반", { bold: true }), c("그 밖의 업무", { color: GRAY }), c("제한 없음", { bold: true })]],
    { x: 5.85, y: 3.2, w: 6.6, colW: [1.3, 2.6, 2.7], fontFace: FONT, fontSize: 15, color: INK, rowH: [0.36, 0.5, 0.5, 0.5], margin: [0, 0.06, 0, 0.06], valign: "middle" });
  hline(s, 5.85, 5.3, 6.6, BLUE_M);
  T(s, [{ text: "신입  ", options: { bold: true, color: BLUE_D } }, { text: "선배 가까이 있으면 코드 피드백 +22%" }], { x: 5.85, y: 5.42, w: 6.6, h: 0.38, fontSize: 15 });
  T(s, "→ 온보딩과 설계 회의는 전 직원이 있는 출근일에", { x: 5.85, y: 5.85, w: 6.6, h: 0.38, fontSize: 15, bold: true });
  source(s, "[R-5] IBM·Ponemon (2021) · [2-6] Emanuel·Harrington·Pallais");
  notes(s, 8);
}

// ═══ 9. 공개된 예외 (L-3ROW) ════════════════════════════
{
  const s = newSlide("LEFT_2", "원칙 2 공정성");
  leftTitle(s, "03", ["예외는\n", { k: "공개된 기준" }, "으로"], "리더가 사정을 보고 승인하면\n리더마다 기준이 달라지고\n“저 사람은 되는데 나는 왜?”\n라는 말이 나온다.", 3.35);
  const w = 3.6; clip(s, "src_c6.png", 0.6, 4.95, w, w / 5.902, "데일리안 (2023.1.16)");
  const rows = [["육아 · 가족돌봄 근로시간 단축", "남녀고용평등법 제19조의2, 제22조의3"], ["본인의 질병 · 장애", "공개된 기준에 해당하면 누구나 같은 판정"], ["제도가 바뀔 때의 유예", "원격을 전제로 멀리 이사한 직원은 첫해 주 1일 출근"]];
  rows.forEach(([a, b], i) => {
    const y = 1.4 + i * 1.0;
    s.addText(String(i + 1).padStart(2, "0"), { shape: SH.roundRect, rectRadius: 0.08, x: 4.55, y, w: 0.85, h: 0.85, fill: { color: BLUE_D }, color: WHITE, fontFace: FONT, fontSize: 22, bold: true, align: "center", valign: "middle", margin: 0 });
    card(s, { x: 5.5, y, w: 7.23, h: 0.85 });
    T(s, a, { x: 5.8, y: y + 0.1, w: 6.8, h: 0.36, fontSize: 17, bold: true });
    T(s, b, { x: 5.8, y: y + 0.47, w: 6.8, h: 0.3, fontSize: 12.5, color: GRAY });
  });
  T(s, "︾", { x: 8.3, y: 4.42, w: 0.7, h: 0.35, fontSize: 18, color: BLUE, align: "center" });
  s.addShape(SH.roundRect, { x: 4.55, y: 4.85, w: 8.18, h: 1.3, rectRadius: 0.12, fill: { color: BLUE_L }, line: { color: BLUE_M, width: 1.25 } });
  T(s, [{ text: "직급이 높거나 일이 많다는 것은 사유가 아니다", options: { breakLine: true } }, { text: "판정은 문서로 남기고, 이의를 제기하면 14일 안에 답한다", options: { bold: true, color: BLUE_D } }], { x: 4.75, y: 4.85, w: 7.8, h: 1.3, fontSize: 16, align: "center", valign: "middle", lineSpacingMultiple: 1.35 });
  source(s, "[4-1] Rousseau 외 (2006) · [4-3] Colquitt (2001) · [C-6] 데일리안 · [L-6] 남녀고용평등법");
  notes(s, 9);
}

// ═══ 10. 덜 보이는 사람을 위한 장치 (L-ASIS) ═════════════
{
  const s = newSlide("TOP_2", "원칙 2 공정성");
  topTitle(s, "03", ["덜 보이는 사람을 위한 ", { k: "장치" }]);
  T(s, "문제", { x: 0.95, y: 2.2, w: 3, h: 0.3, fontSize: 13, bold: true, color: GRAY });
  T(s, "저희 안", { x: 7.15, y: 2.2, w: 3, h: 0.3, fontSize: 13, bold: true, color: BLUE_D });
  const rows = [["성과가 같아도 재택자가 승진에서 불리", "콜센터 실험", "평가 · 승진 · 좋은 프로젝트 배정을 매년 공개"],
                ["유연근무를 쓰면 헌신이 낮다는 낙인", "유연성 낙인 연구", "같은 직군 · 직급에서 5%p 넘게 차이 나면 노사가 재검토"],
                ["효과가 확인된 제도도 합병 뒤 폐지", "STAR 실험", "근무제를 취업규칙 · 단체협약에 적는다"]];
  rows.forEach(([a, tag, b], i) => {
    const y = 2.55 + i * 1.0;
    ocard(s, { x: 0.6, y, w: 5.55, h: 0.82 });
    s.addShape(SH.ellipse, { x: 0.85, y: y + 0.17, w: 0.22, h: 0.22, fill: { color: BLUE_D }, line: NONE });
    T(s, [{ text: a, options: { bold: true, breakLine: true } }, { text: tag, options: { fontSize: 11, color: GRAY } }], { x: 1.25, y: y + 0.1, w: 4.8, h: 0.65, fontSize: 15 });
    T(s, "›››", { x: 6.22, y, w: 0.6, h: 0.82, fontSize: 18, bold: true, color: BLUE, align: "center", valign: "middle" });
    card(s, { x: 6.9, y, w: 5.83, h: 0.82 });
    T(s, b, { x: 7.1, y, w: 5.45, h: 0.82, fontSize: 15, bold: true, color: BLUE_D, align: "center", valign: "middle" });
  });
  const w = 6.2; clip(s, "src_star.png", 0.6, 5.75, w, w / 11.556, "STAR 실험 후속 논문, Social Problems 64(1) · MIT DSpace");
  source(s, "[2-1] Bloom 외 (2015) · [4-4] Williams 외 (2013) · [3-2] Moen·Kelly 외 · 5%p는 저희가 정한 설계값");
  notes(s, 10);
}

// ═══ 11. 밤 11시 장애 (L-PROC) ══════════════════════════
{
  const s = newSlide("TOP_3", "원칙 3 측정");
  topTitle(s, "04", ["밤 11시, ", { k: "장애" }, "가 나면"], "연결 차단 정책만으로는 업무시간 밖 연락이 줄지 않았다. 그래서 기록과 규칙을 얹는다");
  const steps = [["23:00", "장애 발생", "응답 의무의 예외는\n서비스 장애와\n보안 사고 두 가지뿐"], ["23:05", "대기 순번이 받는다", "연결을 끈 사람에게는\n연락하지 않는다"], ["01:00", "복구, 시간 기록", "연장·야간근로로 기록\n간주근로는 쓰지 않는다"], ["12:00", "11시간 쉰 뒤 업무", "가산임금과 보상휴가\n가운데 고른다"]];
  steps.forEach(([t, h, d], i) => {
    const x = 0.6 + i * 3.12;
    ocard(s, { x, y: 2.75, w: 2.78, h: 2.75 });
    pill(s, t, { x: x + 0.79, y: 2.57, w: 1.2, h: 0.38, fontSize: 13, fill: { color: i === 3 ? BLUE_D : BLUE } });
    T(s, h, { x: x + 0.15, y: 3.3, w: 2.48, h: 0.45, fontSize: 18, bold: true, align: "center" });
    hline(s, x + 0.9, 3.88, 1.0, BLUE_M);
    T(s, d, { x: x + 0.15, y: 4.05, w: 2.48, h: 1.2, fontSize: 13.5, color: GRAY, align: "center", lineSpacingMultiple: 1.3 });
    if (i < 3) T(s, "››", { x: x + 2.8, y: 3.85, w: 0.32, h: 0.4, fontSize: 16, bold: true, color: BLUE, align: "center" });
  });
  card(s, { x: 0.6, y: 5.75, w: 12.13, h: 0.82 });
  T(s, [{ text: "유럽 조사  ", options: { bold: true, color: BLUE_D } }, { text: "연결 차단 정책이 있는 회사는 일·생활 균형 만족 " }, { text: "92%", options: { bold: true } }, { text: " (없는 회사 80%). 그러나 업무시간 밖 연락은 " }, { text: "줄지 않았다", options: { bold: true, color: ORANGE } }], { x: 0.9, y: 5.75, w: 11.6, h: 0.82, fontSize: 15, valign: "middle" });
  source(s, "[P-2] Eurofound (2023), Right to disconnect · [K-3] 네이버 노조 조사 · [L-2] 근로기준법 제52조 · [L-4] 제57조");
  notes(s, 11);
}

// ═══ 12. AI 효과는 기록으로 (L-KEY) ══════════════════════
{
  const s = newSlide("TOP_3", "원칙 3 측정");
  topTitle(s, "04", ["AI 효과는 ", { k: "기록" }, "으로"], "AI가 일을 줄여 주니 근무시간도 줄이면 되지 않을까?");
  card(s, { x: 0.6, y: 2.55, w: 6.15, h: 4.05 });
  T(s, "AI를 쓴 숙련 개발자 16명의 작업 속도", { x: 0.9, y: 2.75, w: 5.6, h: 0.3, fontSize: 13, bold: true, color: GRAY });
  T(s, "느낌", { x: 0.9, y: 3.15, w: 2.6, h: 0.3, fontSize: 15, color: GRAY });
  T(s, "+20%", { x: 0.9, y: 3.45, w: 2.8, h: 0.9, fontSize: 54, bold: true, color: BLUE_D });
  T(s, "실제", { x: 3.75, y: 3.15, w: 2.6, h: 0.3, fontSize: 15, color: GRAY });
  T(s, "−19%", { x: 3.75, y: 3.45, w: 2.8, h: 0.9, fontSize: 54, bold: true, color: ORANGE });
  const w = 5.55; clip(s, "src_metr.png", 0.9, 4.55, w, w / 3.683, "METR (2025.7.10) 연구 발표");
  const rules = [["시간이 오히려 늘었다", "업무량부터 조정", "총 근로시간 · 업무시간 외 활동 · 재작업 가운데 하나라도"], ["주 2시간 넘게 줄고 품질 유지", "격주 금요일 오후 휴무", "시범 기간의 유급휴가, 계약상 근무시간은 그대로"], ["처리량만 늘었다 (+15~26%)", "돌려줄 방식은 노사가 결정", "덴마크 2만 5천 명: 근로시간·임금은 거의 그대로"]];
  rules.forEach(([a, b, c], i) => {
    const y = 2.55 + i * 1.38;
    (i === 1 ? card : ocard)(s, { x: 7.0, y, w: 5.73, h: 1.22 });
    T(s, a, { x: 7.3, y: y + 0.13, w: 5.2, h: 0.3, fontSize: 13, color: GRAY });
    T(s, "→ " + b, { x: 7.3, y: y + 0.43, w: 5.2, h: 0.4, fontSize: 18, bold: true, color: i === 1 ? BLUE_D : INK });
    if (c) T(s, c, { x: 7.3, y: y + 0.86, w: 5.2, h: 0.26, fontSize: 11, color: GRAY });
  });
  source(s, "[5-3] METR (2025) · [5-1] Brynjolfsson 외 (2025) · [5-2] Cui 외 (2025) · [5-4] Humlum & Vestergaard (2025) · [P-6] 워라밸+4.5");
  notes(s, 12);
}

// ═══ 13. 주 3일은 상한 (L-GRAPH) ═════════════════════════
{
  const s = newSlide("LEFT_4", "원칙 4 강도");
  leftTitle(s, "05", ["주 3일은\n", { k: "상한" }], "출근은 많을수록 좋을까?\n아마존은 주 5일 출근으로\n돌아갔다.", 3.35);
  const w = 3.6; clip(s, "src_amzn.png", 0.6, 4.2, w, w / 2.964, "Amazon CEO 메시지 (2024.9.16)");
  card(s, { x: 4.55, y: 1.35, w: 8.18, h: 3.15 });
  T(s, "출근 요구가 강할수록 오래 일한 사람이 더 떠났다", { x: 4.85, y: 1.52, w: 7.6, h: 0.36, fontSize: 16, bold: true });
  T(s, "출근 의무화 뒤 재직 0~3년 비중의 증가 · 이력서 2억 6천만 건 · 학술지 심사 전", { x: 4.85, y: 1.88, w: 7.6, h: 0.28, fontSize: 11.5, color: GRAY });
  const bars = [["Apple · 주 1일", 4, "4% 미만"], ["Microsoft · 출근 50%", 6, "약 6%"], ["SpaceX · 전일 출근", 15, "최대 15%"]];
  const base = 4.05, u = 0.1;
  bars.forEach(([lab, v, t], i) => {
    const x = 5.55 + i * 2.35;
    s.addShape(SH.rect, { x, y: base - v * u, w: 1.2, h: v * u, fill: { color: i === 2 ? BLUE_D : GRAY_L }, line: NONE });
    T(s, t, { x: x - 0.4, y: base - v * u - 0.4, w: 2.0, h: 0.36, fontSize: 16, bold: true, color: i === 2 ? BLUE_D : INK, align: "center" });
    T(s, lab, { x: x - 0.5, y: base + 0.06, w: 2.2, h: 0.28, fontSize: 11.5, color: GRAY, align: "center" });
  });
  hline(s, 4.95, base, 7.4);
  T(s, "반대쪽 경계: 재택 일수", { x: 4.6, y: 4.75, w: 4.3, h: 0.32, fontSize: 14, bold: true });
  const sx = 4.8, sw = 4.0, sy = 5.65;
  s.addShape(SH.rect, { x: sx, y: sy, w: sw / 2, h: 0.2, fill: { color: BLUE_L }, line: NONE });
  s.addShape(SH.rect, { x: sx + sw / 2, y: sy, w: sw / 2, h: 0.2, fill: { color: "F6D3BD" }, line: NONE });
  [0, 1, 2, 3, 4, 5].forEach(d => T(s, d + "일", { x: sx + sw * d / 5 - 0.3, y: sy + 0.27, w: 0.6, h: 0.26, fontSize: 11, color: GRAY, align: "center" }));
  s.addShape(SH.ellipse, { x: sx + sw * 0.4 - 0.16, y: sy - 0.06, w: 0.32, h: 0.32, fill: { color: BLUE_D }, line: { color: WHITE, width: 2 } });
  T(s, "재택 2일", { x: sx + sw * 0.4 - 0.8, y: sy - 0.45, w: 1.6, h: 0.34, fontSize: 15, bold: true, color: BLUE_D, align: "center" });
  T(s, "2.5일 넘으면 관계 악화", { x: sx + sw * 0.5 + 0.05, y: sy - 0.42, w: 2.0, h: 0.3, fontSize: 11.5, bold: true, color: ORANGE, align: "right" });
  vline(s, 9.15, 4.8, 1.7);
  T(s, "‘3일 이상’이 아니라\n딱 ‘3일’", { x: 9.5, y: 4.85, w: 3.2, h: 0.95, fontSize: 21, bold: true, color: BLUE_D, lineSpacingMultiple: 1.1 });
  T(s, "상한이 없으면 4일, 5일이\n될 거라는 불안이 남는다", { x: 9.5, y: 5.85, w: 3.2, h: 0.6, fontSize: 12.5, color: GRAY });
  source(s, "[R-3] Van Dijcke 외 (2024), BFI 워킹페이퍼 · [2-4] Gajendran & Harrison (2007), 연구 46개 · [C-9] Amazon · [B-1][B-4] 우아한형제들");
  notes(s, 13);
}

// ═══ 14. 카카오 구분 장 (L-DIV) ══════════════════════════
{
  const s = newSlide("DIV_5", "카카오 적용");
  badge(s, "06", 6.667 - 0.29, 2.05);
  s.addText([{ text: "카카오에", options: { color: BLUE_D, breakLine: true } }, { text: "적용합니다" }], { placeholder: "title" });
  card(s, { x: 3.2, y: 4.65, w: 6.93, h: 0.7, rectRadius: 0.35 });
  T(s, "재택을 가장 크게 되돌렸다가 다시 늘리고 있는 회사", { x: 3.2, y: 4.65, w: 6.93, h: 0.7, fontSize: 17, bold: true, color: BLUE_D, align: "center", valign: "middle" });
  notes(s, 14);
}

// ═══ 15. 카카오 근무제 연혁 (L-HIST) ═════════════════════
{
  const s = newSlide("TOP_5", "카카오 적용");
  topTitle(s, "06", ["카카오 근무제 ", { k: "연혁" }], "기둥 높이는 재택을 얼마나 허용했는지 (정성적 비교)");
  const cols = [["2020", "전사 재택", 2.6, BLUE_S], ["2022", "장소 무관\n주 1회 모임", 2.85, BLUE_S], ["2023", "오피스 원칙\n원격은 예외", 4.2, BLUE_S], ["2024", "전원 출근", 5.15, ORANGE], ["2025", "주 1회 원격\n요일은 각자", 4.6, BLUE_D]];
  cols.forEach(([yr, t, top, c], i) => {
    const x = 0.6 + i * 1.56;
    card(s, { x, y: top, w: 1.42, h: 6.55 - top, rectRadius: 0.08 });
    pill(s, yr, { x: x + 0.21, y: top - 0.2, w: 1.0, h: 0.36, fontSize: 13, fill: { color: c } });
    T(s, t, { x: x + 0.05, y: top + 0.28, w: 1.32, h: 0.75, fontSize: 13.5, bold: true, color: c === ORANGE ? ORANGE : INK, align: "center", lineSpacingMultiple: 1.15 });
  });
  card(s, { x: 8.6, y: 2.45, w: 4.13, h: 1.25 });
  T(s, "주 1회 원격 뒤 원격근무 이용자", { x: 8.85, y: 2.57, w: 3.7, h: 0.3, fontSize: 12, color: GRAY });
  T(s, [{ text: "14명 → ", options: { color: "9AA3AF" } }, { text: "3,854명" }], { x: 8.85, y: 2.9, w: 3.8, h: 0.65, fontSize: 32, bold: true, valign: "middle" });
  const H = 0.58;
  [["clip_ka7.png", 3.872, 3.95], ["clip_ka8.png", 6.27, 4.6], ["clip_ka10.png", 5.307, 5.25], ["clip_ka9.png", 5.222, 5.9]].forEach(([f, r, y], i) => clip(s, f, 8.63 + (i % 2) * 0.2, y, H * r, H));
  source(s, "[KA-7] 디지털데일리 (2024.10.30) · [KA-8] 디지털투데이 (2024.11.21) · [KA-10] 경기일보 (2025.11.17) · [KA-9] 뉴스웨이 (2026.4.9), 이용자 수는 보도 수치");
  notes(s, 15);
}

// ═══ 16. 카카오의 한 주 ════════════════════════════════
{
  const s = newSlide("TOP_5", "카카오 적용");
  topTitle(s, "06", ["카카오의 ", { k: "한 주" }, "는 이렇게 바뀝니다"]);
  const days = ["월", "화", "수", "목", "금"];
  function grid(x0, label, rows, together, caption, hi) {
    T(s, label, { x: x0, y: 2.25, w: 5.9, h: 0.38, fontSize: 15, bold: true, color: hi ? BLUE_D : GRAY });
    days.forEach((d, j) => {
      const x = x0 + 0.6 + j * 1.0, on = together.includes(j);
      T(s, d, { x, y: 2.75, w: 0.92, h: 0.32, fontSize: 14, bold: on, color: on ? INK : GRAY, align: "center" });
      if (on) s.addShape(SH.rect, { x: x + 0.14, y: 3.08, w: 0.64, h: 0.06, fill: { color: KAKAO }, line: NONE });
    });
    rows.forEach(([who, cells], r) => {
      const y = 3.25 + r * 0.6;
      T(s, who, { x: x0, y, w: 0.5, h: 0.5, fontSize: 15, bold: true, color: GRAY, align: "center", valign: "middle" });
      cells.forEach((c, j) => {
        const x = x0 + 0.6 + j * 1.0;
        const st = c === "출" ? { fill: { color: BLUE_D }, color: WHITE } : c === "재" ? { fill: { color: WHITE }, color: GRAY, line: { color: GRAY_L, width: 1 } } : { fill: { color: BLUE_L }, color: BLUE_D };
        s.addText(c === "출" ? "출근" : c === "재" ? "재택" : "선택", Object.assign({ shape: SH.roundRect, rectRadius: 0.06, x, y, w: 0.92, h: 0.5, fontFace: FONT, fontSize: 13, bold: c !== "재", align: "center", valign: "middle", margin: 0 }, st));
      });
    });
    T(s, caption, { x: x0, y: 5.15, w: 5.9, h: 0.38, fontSize: 16, bold: true, color: hi ? BLUE_D : INK });
  }
  grid(0.6, "지금 · 원격 요일을 각자 고름, 주 4일 출근", [["A", ["출", "출", "재", "출", "출"]], ["B", ["재", "출", "출", "출", "출"]], ["C", ["출", "출", "출", "출", "재"]]], [1, 3], "셋이 함께 있는 날은 화·목, 이틀", false);
  s.addText("VS", { shape: SH.ellipse, x: 6.47, y: 3.7, w: 0.4, h: 0.4, fill: { color: BLUE_D }, color: WHITE, fontFace: FONT, fontSize: 11, bold: true, align: "center", valign: "middle", margin: 0 });
  grid(6.95, "적용 · 공통 출근일 3일, 주 3일 출근", [["A", ["출", "출", "출", "선", "선"]], ["B", ["출", "출", "출", "선", "선"]], ["C", ["출", "출", "출", "선", "선"]]], [0, 1, 2], "셋이 함께 있는 날은 월·화·수, 사흘", true);
  card(s, { x: 0.6, y: 5.8, w: 12.13, h: 0.75 });
  T(s, [{ text: "코워크 타임(11~16시)은 어떨까?  ", options: { bold: true, color: BLUE_D } }, { text: "2024년 회사 제안, 노조 반대로 권장에 그쳤다. 시간대 대신 요일로 묶는다" }], { x: 0.9, y: 5.8, w: 11.6, h: 0.75, fontSize: 15, valign: "middle" });
  notes(s, 16);
}

// ═══ 17. 공정성·측정 (L-ASIS) ═══════════════════════════
{
  const s = newSlide("TOP_5", "카카오 적용");
  topTitle(s, "06", ["있는 제도와 데이터를 ", { k: "규정과 공개" }, "로"]);
  T(s, "카카오 공개 자료 · 2024 ESG 보고서", { x: 1.95, y: 2.2, w: 5, h: 0.3, fontSize: 13, bold: true, color: GRAY });
  T(s, "적용", { x: 7.55, y: 2.2, w: 5, h: 0.3, fontSize: 13, bold: true, color: BLUE_D });
  const rows = [["공정성", "esg92_kakaoon_text.png", 4.326, 4.85, "92쪽 · 근무 장소는 오피스가 원칙, 근무제는 2024 임단협으로", ["예외 3가지와 판정 절차까지 단체협약에", "이의는 라운드테이블과 노조로", "평가 격차는 매년 공개"]],
                ["측정", "esg92_health.png", 3.5596, 3.9, "92쪽 · 조직건강성 조사 응답률 34%, 결과 수치는 비공개", ["정산 결과를 팀별로 공개", "감독 지적 사항은 바로 반영", "진단 결과를 근무 유형별로 공개"]]];
  rows.forEach(([lab, f, r, w, cap, app], i) => {
    const y = 2.6 + i * 2.1, h = w / r;
    s.addText(lab, { shape: SH.roundRect, rectRadius: 0.08, x: 0.6, y, w: 1.15, h: 1.55, fill: { color: BLUE_D }, color: WHITE, fontFace: FONT, fontSize: 16, bold: true, align: "center", valign: "middle", margin: 0 });
    clip(s, f, 1.98, y + 0.03, w, h, cap);
    if (i === 1) s.addShape(SH.rect, { x: 1.98 + w * 0.378, y: y + 0.03 + h * 0.70, w: w * 0.08, h: h * 0.13, fill: { color: KAKAO, transparency: 45 }, line: NONE });
    T(s, "›››", { x: 6.92, y, w: 0.55, h: 1.55, fontSize: 18, bold: true, color: BLUE, align: "center", valign: "middle" });
    card(s, { x: 7.55, y, w: 5.18, h: 1.55 });
    T(s, app.join("\n"), { x: 7.8, y, w: 4.8, h: 1.55, fontSize: 15.5, bold: true, color: BLUE_D, valign: "middle", lineSpacingMultiple: 1.35 });
  });
  source(s, "[KA-1][KA-2] 카카오 2024 ESG 보고서 p.92 · [KA-8] 2024 임단협 보도 · [KA-10] 경기일보 (2025.11.17), 감독 결과는 미확인");
  notes(s, 17);
}

// ═══ 18. 6개월 실험 (간트) ══════════════════════════════
{
  const s = newSlide("TOP_5", "카카오 적용");
  topTitle(s, "06", [{ k: "6개월" }, " 비교 실험으로 결정"], "이 안이 정말 카카오에 적합할까? 저희도 장담하지 못합니다");
  const months = ["11", "12", "1", "2", "3", "4", "5", "6", "7", "8"];
  const gx = 3.4, mw = 0.93, gy = 2.85;
  card(s, { x: 0.6, y: 2.4, w: 12.13, h: 3.2 });
  T(s, "2026", { x: gx, y: 2.5, w: 1.8, h: 0.26, fontSize: 11, bold: true, color: GRAY });
  T(s, "2027", { x: gx + 2 * mw, y: 2.5, w: 1.8, h: 0.26, fontSize: 11, bold: true, color: GRAY });
  months.forEach((m, i) => T(s, m + "월", { x: gx + i * mw, y: gy - 0.08, w: mw, h: 0.28, fontSize: 11, color: GRAY, align: "center" }));
  vline(s, gx + 2 * mw, 2.55, 2.85);
  const tasks = [["설계 · 노사 합의", 0, 2, "9CC8F2"], ["기준선 측정 · 출근일 지정", 2, 1, BLUE_S], ["현행 vs 저희 안 운영", 2, 6, BLUE_D], ["판정 · 공개", 8, 2, "2F84D6"]];
  tasks.forEach(([n, st, len, col], i) => {
    const y = gy + 0.35 + i * 0.5;
    T(s, n, { x: 0.85, y, w: 2.5, h: 0.34, fontSize: 13.5, bold: i === 2, valign: "middle" });
    s.addShape(SH.roundRect, { x: gx + st * mw + 0.03, y: y + 0.03, w: len * mw - 0.06, h: 0.28, rectRadius: 0.06, fill: { color: col }, line: NONE });
  });
  const ry = gy + 0.35 + 2 * 0.5, cx = gx + 5 * mw;
  s.addShape(SH.diamond, { x: cx - 0.13, y: ry + 0.04, w: 0.26, h: 0.26, fill: { color: ORANGE }, line: { color: WHITE, width: 1.5 } });
  T(s, "3개월 차 점검", { x: cx - 0.8, y: ry - 0.27, w: 1.6, h: 0.24, fontSize: 11, bold: true, color: ORANGE, align: "center" });
  T(s, "→ 다음 단체협약에 반영, 최소 2년 유지", { x: 7.0, y: gy + 2.35, w: 5.5, h: 0.3, fontSize: 13, bold: true, color: BLUE_D, align: "right" });
  [["비교", "비슷한 사업부끼리 짝지어 추첨\n판단 기준은 시작 전에 노사가 공개", BLUE_D], ["중단", "경력자 이직이 짝 사업부의 2배 이상\n그리고 3%p 넘게 차이 나면 멈춘다", ORANGE]].forEach(([h, t, c], i) => {
    const x = 0.6 + i * 6.15;
    ocard(s, { x, y: 5.8, w: 5.98, h: 0.82, line: { color: c === ORANGE ? "F2B48E" : BLUE_M, width: 1.25 } });
    T(s, h, { x: x + 0.25, y: 5.8, w: 0.8, h: 0.82, fontSize: 15, bold: true, color: c, valign: "middle" });
    T(s, t, { x: x + 1.0, y: 5.8, w: 4.85, h: 0.82, fontSize: 13, valign: "middle", lineSpacingMultiple: 1.2 });
  });
  notes(s, 18);
}

// ═══ 19. 마무리 (L-END) ════════════════════════════════
{
  const s = newSlide("TOP_N", "결론");
  s.addText(" ", { placeholder: "title" });
  s.addShape(SH.roundRect, { x: 0.6, y: 1.45, w: 12.13, h: 4.45, rectRadius: 0.2, fill: { color: BLUE }, line: NONE });
  pill(s, "대면의 가치는 지키고, 그 대가는 공정하게", { x: 1.1, y: 2.0, w: 4.6, h: 0.46, fontSize: 14, rectRadius: 0.23 });
  T(s, "감사합니다", { x: 1.1, y: 2.65, w: 6, h: 1.1, fontSize: 56, bold: true, color: WHITE, valign: "middle" });
  T(s, "현장에서 보시기에 어디가 비현실적인지\n말씀해 주십시오", { x: 1.1, y: 3.9, w: 6, h: 0.8, fontSize: 17, color: WHITE, lineSpacingMultiple: 1.25 });
  T(s, [{ text: "한계", options: { bold: true, breakLine: true } }, { text: "네 원칙을 한꺼번에 시험한 연구는 아직 없다", options: { breakLine: true } }, { text: "같은 요일의 근거는 한 회사의 실험이다", options: { breakLine: true } }, { text: "→ 마지막 판단은 각 회사의 실험으로", options: { bold: true } }], { x: 7.3, y: 3.75, w: 5.1, h: 1.6, fontSize: 13, color: WHITE, align: "right", lineSpacingMultiple: 1.35 });
  T(s, "질의응답", { x: 0.6, y: 6.15, w: 6, h: 0.4, fontSize: 15, bold: true, color: BLUE_D });
  notes(s, 19);
}

// ═══ 부록. 참고 자료 ════════════════════════════════════
{
  const s = newSlide("TOP_N", "부록");
  topTitle(s, null, ["참고 자료: 카카오 ", { k: "2024 ESG 보고서" }], "질의응답용 · 발표 중에는 넘어간다");
  const wL = 5.9; clip(s, "esg146_turnover.png", 0.6, 2.45, wL, wL / 1.9966);
  T(s, "146쪽 이직 현황 · 이직률 9.6% → 5.4% → 5.6%. 같은 기간 신규채용 870 → 452 → 314명이라\n이직률 하락을 근무제 효과로 단정하지 않는다", { x: 0.6, y: 5.5, w: 5.9, h: 0.55, fontSize: 10.5, color: GRAY });
  const wR = 5.8; clip(s, "esg97_remote.png", 6.93, 2.45, wR, wR / 3.454, "97쪽 · 2025년 주 1회 원격근무 시행, 매월 마지막 주 금요일은 리커버리데이");
  const hN = 1.6, wN = hN * (705 / 315); clip(s, "news_ka9.png", 6.93, 4.75, wN, hN);
  T(s, "뉴스웨이 (2026.4.9)\n같은 기사의 정규직 수는\n3,753명이라, 이용자 수\n3,854명의 집계 기준은\n밝혀지지 않았다", { x: 6.93 + wN + 0.2, y: 4.75, w: 12.73 - (6.93 + wN + 0.2), h: 1.6, fontSize: 10.5, color: GRAY, valign: "middle" });
  s.addNotes("질의응답용 참고 슬라이드입니다. 이직률이나 원격근무 시행 시기, 3,854명에 대한 질문이 나오면 띄웁니다.");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
