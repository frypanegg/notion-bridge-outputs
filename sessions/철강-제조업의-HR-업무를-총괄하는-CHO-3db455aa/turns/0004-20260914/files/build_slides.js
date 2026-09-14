// HR 중기전략('26~'30 Rolling) 3장 PPTX 생성 스크립트 (Ⅰ Rolling 방향 / Ⅱ '30 추진전략 / Ⅲ 2035 HR 미래 청사진)
// 실행: node build_slides.js  (pptxgenjs 3.x 필요)
const path = require('path');
const pptxgen = require('pptxgenjs');

const ONLY = process.env.ONLY_SLIDE ? Number(process.env.ONLY_SLIDE) : 0;
const OUT = process.env.OUT_FILE || path.join(__dirname, 'HR_중기전략_2030_인당생산성_Rolling.pptx');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5 in
pptx.title = "HR 중기전략('26~'30 Rolling) - 인당생산성 향상 및 2035 HR 미래 청사진";

const FONT = '맑은 고딕';
const C = {
  navy: '0B2D5B', blue: '1F5FA8', mid: '16498A', sky: 'D6E4F5', sky2: 'EEF3FA',
  gray: 'F2F4F7', line: 'BFC7D2', text: '222222', sub: '555555', mute: '8A94A3',
  orange: 'E8541E', orangeD: 'C2410C', orangeL: 'FDEDE6', light: '8FB3E0', hyup: 'C9D1DB', white: 'FFFFFF',
};

// ---------- 인력 운영계획 원본 (TO, 명 / 기준시점: '24년 말) ----------
const YEARS = ["'24.E", "'25.E", "'26.E", "'27.E", "'28.E", "'29.E", "'30.E"];
const DATA = {
  jeong: [16940, 16306, 17476, 20812, 19342, 18180, 16769], // 직영 정규직 (원본 수치, S직군 포함)
  sj: null, // 직영 S직군: 연도별 수치 확정 시 배열 입력 → 정규직에서 차감해 별도 축으로 표시
  gye: [1057, 1427, 1802, 2125, 2304, 2352, 2164], // 직영 계약직
  hyup: [10107, 9966, 8130, 3630, 3557, 3555, 3298], // 협력
};
const TOTAL_CHECK = [28104, 27699, 27408, 26567, 25203, 24087, 22231];

const SJ_READY = Array.isArray(DATA.sj);
const sjv = SJ_READY ? DATA.sj : YEARS.map(() => 0);
const SERIES = [
  { name: '직영 정규직', color: C.navy, label: C.white, v: DATA.jeong.map((x, i) => x - sjv[i]) },
  { name: '직영 S직군', color: C.blue, label: C.white, v: sjv },
  { name: '직영 계약직', color: C.light, label: C.navy, v: DATA.gye },
  { name: '협력', color: C.hyup, label: C.navy, v: DATA.hyup },
];
const TOTAL = YEARS.map((_, i) => SERIES.reduce((a, se) => a + se.v[i], 0));
TOTAL.forEach((t, i) => {
  if (t !== TOTAL_CHECK[i]) throw new Error(`총계 불일치 ${YEARS[i]}: ${t} vs ${TOTAL_CHECK[i]}`);
});
const LAST = YEARS.length - 1;
const fmt = (n) => Math.abs(n).toLocaleString('en-US');
const sign = (n) => (n > 0 ? '+' : n < 0 ? '△' : '');
const dNum = (n) => sign(n) + fmt(n);
const dPct = (n) => sign(n) + Math.abs(n).toFixed(1) + '%';
const D_TOTAL = TOTAL[LAST] - TOTAL[0];
const D_TOTAL_PCT = (D_TOTAL / TOTAL[0]) * 100;
const D_JIK = DATA.jeong[LAST] + DATA.gye[LAST] - DATA.jeong[0] - DATA.gye[0];
const D_HYUP = DATA.hyup[LAST] - DATA.hyup[0];
const PROD_UP = (TOTAL[0] / TOTAL[LAST] - 1) * 100; // 생산량 동일 가정 時 인당생산성 증가율

// ---------- helpers ----------
function T(slide, text, o) {
  slide.addText(text, Object.assign({
    fontFace: FONT, fontSize: 9, color: C.text, margin: 2, valign: 'middle',
  }, o));
}
function R(slide, o) {
  slide.addShape(pptx.ShapeType.rect, Object.assign({ fill: { color: C.white }, line: { color: C.line, width: 0.75 } }, o));
}
function TB(slide, text, o) { // 텍스트가 들어간 도형
  T(slide, text, Object.assign({ shape: pptx.ShapeType.rect, fill: { color: C.white }, line: { color: C.line, width: 0.75 } }, o));
}
function seg(slide, x1, y1, x2, y2, line) {
  const o = { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), line };
  if ((x2 - x1) * (y2 - y1) < 0) o.flipV = true;
  slide.addShape(pptx.ShapeType.line, o);
}
function run(text, o) { return { text, options: Object.assign({ fontFace: FONT }, o) }; }

function header(slide, title, tag, msg, top = "HR 중기전략 ('26~'30 Rolling)  |  원가 경쟁력") {
  T(slide, top, { x: 0.4, y: 0.16, w: 6, h: 0.22, fontSize: 9, color: C.mute, margin: 0 });
  T(slide, title, { x: 0.4, y: 0.38, w: 9.6, h: 0.5, fontSize: 22, bold: true, color: C.navy, margin: 0 });
  TB(slide, tag, { x: 10.13, y: 0.47, w: 2.8, h: 0.34, fill: { color: C.orangeL }, line: { color: C.orange, width: 1 }, fontSize: 10, bold: true, color: C.orangeD, align: 'center' });
  slide.addShape(pptx.ShapeType.line, { x: 0.4, y: 0.93, w: 12.53, h: 0, line: { color: C.navy, width: 2 } });
  R(slide, { x: 0.4, y: 1.03, w: 12.53, h: 0.56, fill: { color: C.sky2 }, line: { color: C.sky2, width: 0 } });
  R(slide, { x: 0.4, y: 1.03, w: 0.07, h: 0.56, fill: { color: C.navy }, line: { color: C.navy, width: 0 } });
  T(slide, msg, { x: 0.55, y: 1.03, w: 12.3, h: 0.56, fontSize: 12, bold: true, color: C.navy, margin: 2 });
}
function footer(slide, note, page) {
  T(slide, note, { x: 0.4, y: 7.15, w: 11.8, h: 0.24, fontSize: 8, color: C.sub, margin: 0 });
  T(slide, String(page), { x: 12.43, y: 7.15, w: 0.5, h: 0.24, fontSize: 9, color: C.sub, align: 'right', margin: 0 });
}
function secHead(slide, text, x, y, w, h = 0.34, align = 'center', size = 11) {
  TB(slide, text, { x, y, w, h, fill: { color: C.navy }, line: { color: C.navy, width: 0 }, fontSize: size, bold: true, color: C.white, align, margin: 6 });
}

// =====================================================================
// Slide 1. Rolling 방향 : 여건 변화 → 인력운영 패러다임 전환
// =====================================================================
function slide1() {
  const s = pptx.addSlide();
  header(s, 'Ⅰ. 인당생산성 향상 중기전략 Rolling 방향', '원가 ④ 인당생산성 향상',
    "근지위 소송·협력작업 혁신에 따른 인력구조 변화와 저수익 구조 장기화를 반영하여, 인력운영 관리기준을 " +
    "『직영 인원수(Volume)』에서 『Total 인력·노무비(Cost)』 중심으로 전환");

  // ---- Col A : 전사 중기경영전략 연계 ----
  const ax = 0.4, aw = 3.0;
  secHead(s, '전사 중기경영전략 연계', ax, 1.72, aw);
  R(s, { x: ax, y: 2.06, w: aw, h: 4.38 });
  T(s, "'30 중기 경영전략 5대 축", { x: 0.5, y: 2.12, w: 2.8, h: 0.26, fontSize: 9.5, bold: true, color: C.navy, margin: 0 });
  ['성장', '원가', '기술\n제품', '미래', '사업\n재편'].forEach((t, i) => {
    const on = i === 1;
    TB(s, t, {
      x: 0.5 + i * 0.57, y: 2.42, w: 0.52, h: 0.48, margin: 0, align: 'center', fontSize: 8.5, bold: on,
      fill: { color: on ? C.orange : C.gray }, line: { color: on ? C.orange : C.line, width: 0.75 }, color: on ? C.white : C.sub,
    });
  });
  s.addShape(pptx.ShapeType.downArrow, { x: 1.2, y: 2.95, w: 0.26, h: 0.22, fill: { color: C.orange }, line: { color: C.orange, width: 0 } });
  TB(s, [
    run('④ 인당생산성 향상', { fontSize: 12, bold: true, color: C.orangeD, breakLine: true }),
    run('자동화 및 WX 기반 인력 최적화', { fontSize: 9.5, color: C.text }),
  ], { x: 0.5, y: 3.2, w: 2.8, h: 0.72, align: 'center', fill: { color: C.orangeL }, line: { color: C.orange, width: 1.5 } });

  T(s, "'25 수립 HR 중기전략 (3대 방향)", { x: 0.5, y: 4.0, w: 2.8, h: 0.26, fontSize: 9.5, bold: true, color: C.navy, margin: 0 });
  ['DX 연계 인적경쟁력 제고/우수인재 육성', '노무경쟁력 혁신 통한 노동환경 변화 대응', 'AI·DX 활용 일하는 방식 근본적 개선'].forEach((t, i) => {
    const y = 4.3 + i * 0.42;
    R(s, { x: 0.5, y, w: 2.8, h: 0.37, fill: { color: C.gray }, line: { color: C.gray, width: 0 } });
    TB(s, String(i + 1), { shape: pptx.ShapeType.ellipse, x: 0.56, y: y + 0.075, w: 0.22, h: 0.22, margin: 0, align: 'center', fontSize: 8, bold: true, color: C.white, fill: { color: C.navy }, line: { color: C.navy, width: 0 } });
    T(s, t, { x: 0.83, y, w: 2.45, h: 0.37, fontSize: 8.5, color: C.text, margin: 1 });
  });
  s.addShape(pptx.ShapeType.downArrow, { x: 1.77, y: 5.58, w: 0.26, h: 0.2, fill: { color: C.navy }, line: { color: C.navy, width: 0 } });
  TB(s, [
    run("'26 Rolling 관점", { fontSize: 8.5, bold: true, color: 'FFC7A8', breakLine: true }),
    run("인당생산성 향상의 '30년 구체화", { fontSize: 10.5, bold: true, color: C.white, breakLine: true }),
    run('PO가 아닌 TO 감축 중심 재설계', { fontSize: 8.5, color: C.white }),
  ], { x: 0.5, y: 5.8, w: 2.8, h: 0.6, align: 'center', fill: { color: C.navy }, line: { color: C.navy, width: 0 }, margin: 1 });

  s.addShape(pptx.ShapeType.chevron, { x: 3.45, y: 4.05, w: 0.2, h: 0.36, fill: { color: C.blue }, line: { color: C.blue, width: 0 } });

  // ---- Col B : 여건 변화 / Col C : 패러다임 전환 ----
  const bx = 3.72, bw = 3.9, cx = 7.95;
  secHead(s, '여건 변화 (현상 및 문제점)', bx, 1.72, bw);
  secHead(s, '인력운영 패러다임 전환 (As-Is → To-Be)', cx, 1.72, 4.98);

  const rows = [
    {
      tag: '인력구성', title: 'S직군 직영 전환에 따른 인력구조 변화',
      body: "근지위 소송·협력작업 혁신('26.4월~)으로 직영 인력 증가", so: '직영/협력 구분의 단순 인원수(Volume) 관리 한계 有',
      key: '관리\n대상', asis: '직영/협력\n구분 관리', tobe: '정규·S직군·계약·협력\n4개 축 Total 인력 통합 관리',
    },
    {
      tag: '비용구조', title: '통상임금 확대 등 노무비 상승 압력 가중',
      body: '통상임금 범위 확대로 퇴직급여·복리후생 비용 증대', so: '임금차액 소송에 따른 잠재적 매몰비용 증가',
      key: '관리\n기준', asis: '인원수\n(Head Count)', tobe: 'Total 노무비(Labor Cost)\n기반 원가 관리',
    },
    {
      tag: '생산주체', title: 'AI/로보틱스 등 생산·업무처리 주체 다변화',
      body: '직영/협력 노동력을 대체하는 AI·로보틱스 및 유지보수 인력 확대', so: '사람 중심 인력계획으로는 실질 생산역량 파악 한계',
      key: '관리\n범위', asis: '사람(직영/협력)\n중심 인력계획', tobe: '人+AI/로보틱스 포함\nTotal Workforce 설계',
    },
    {
      tag: '규모통제', title: '유연한 인력규모 통제 難',
      body: '親노동 정책 기조下 채용규모 축소 부담, 대의기구의 인력 충원 요구 확대', so: '인위적 PO 조정 곤란 → 구조적 TO 관리 필요',
      key: '관리\n방식', asis: '충원요구 대응형\nPO 사후 통제', tobe: '기준인력(TO) 선제 설계\n및 상시 Rolling 통제',
    },
  ];
  rows.forEach((r, i) => {
    const y = 2.14 + i * 1.1, h = 1.0;
    // 여건 변화 카드
    R(s, { x: bx, y, w: bw, h });
    TB(s, [
      run('①②③④'[i], { fontSize: 12, bold: true, color: C.navy, breakLine: true }),
      run(r.tag, { fontSize: 8.5, bold: true, color: C.navy }),
    ], { x: bx, y, w: 0.72, h, align: 'center', margin: 0, fill: { color: C.sky }, line: { color: C.line, width: 0.75 } });
    T(s, r.title, { x: 4.5, y: y + 0.05, w: 3.08, h: 0.3, fontSize: 10, bold: true, color: C.text, margin: 1 });
    T(s, [
      run(r.body, { fontSize: 8.5, color: C.sub, breakLine: true }),
      run('→ ' + r.so, { fontSize: 8.5, bold: true, color: C.orangeD }),
    ], { x: 4.5, y: y + 0.35, w: 3.08, h: 0.6, valign: 'top', margin: 1 });

    s.addShape(pptx.ShapeType.rightArrow, { x: 7.64, y: y + 0.37, w: 0.28, h: 0.26, fill: { color: C.blue }, line: { color: C.blue, width: 0 } });

    // As-Is → To-Be
    TB(s, r.key, { x: cx, y, w: 0.8, h, align: 'center', margin: 0, fontSize: 9.5, bold: true, color: C.white, fill: { color: C.mid }, line: { color: C.mid, width: 0 } });
    TB(s, [
      run('As-Is', { fontSize: 7.5, bold: true, color: C.mute, breakLine: true }),
      run(r.asis, { fontSize: 9, color: '444444' }),
    ], { x: 8.8, y, w: 1.62, h, align: 'center', margin: 2, fill: { color: C.gray }, line: { color: C.line, width: 0.75 } });
    s.addShape(pptx.ShapeType.rightArrow, { x: 10.46, y: y + 0.36, w: 0.3, h: 0.28, fill: { color: C.orange }, line: { color: C.orange, width: 0 } });
    TB(s, [
      run('To-Be', { fontSize: 7.5, bold: true, color: C.blue, breakLine: true }),
      run(r.tobe, { fontSize: 9.5, bold: true, color: C.navy }),
    ], { x: 10.8, y, w: 2.13, h, align: 'center', margin: 2, fill: { color: C.sky2 }, line: { color: C.blue, width: 1.25 } });
  });

  // ---- Bottom : Rolling 핵심 ----
  TB(s, 'Rolling 핵심', { x: 0.4, y: 6.56, w: 1.45, h: 0.52, align: 'center', fontSize: 11, bold: true, color: C.white, fill: { color: C.orange }, line: { color: C.orange, width: 0 } });
  TB(s, [
    run('PO(실제인력)가 아닌 ', { fontSize: 10, bold: true, color: C.navy }),
    run('TO(기준인력)를 구조적으로 낮추는 활동', { fontSize: 10, bold: true, color: C.orangeD }),
    run('으로 재정의 ─ 공정·프로세스 자동화, 저효율 라인 효율화, 개인별 업무 Span 확대로 TO를 선행 감축하고, PO는 자연감소·재배치를 통해 TO에 수렴', { fontSize: 10, bold: true, color: C.navy }),
  ], { x: 1.85, y: 6.56, w: 11.08, h: 0.52, margin: 6, fill: { color: C.orangeL }, line: { color: C.orange, width: 1 } });

  footer(s, "※ TO: 기준인력, PO: 실제 운영인력, S직군: 근지위 소송·특별채용을 통해 협력사에서 직영으로 전환된 직군  /  기준시점: '24년 말", 1);

  s.addNotes([
    "[발표 포인트] 작년 수립한 HR 중기전략 3대 방향은 유지하되, 올해는 전사 '원가' 축의 ④ 인당생산성 향상을 '30년까지 어떻게 구체화할 것인가의 관점에서 Rolling했습니다.",
    '작년 대비 달라진 여건은 네 가지입니다. ① S직군 직영 전환으로 직영/협력 구분 관리가 무의미해졌고, ② 통상임금 확대로 인원수보다 노무비 단가가 원가를 좌우하며, ③ AI·로보틱스가 노동력을 대체하고, ④ 인력규모를 인위적으로 조정하기 어려워졌습니다.',
    '이에 따라 관리 대상은 정규직·S직군·계약직·협력 4개 축의 Total 인력, 관리 기준은 노무비, 관리 범위는 AI·로보틱스를 포함한 Total Workforce, 관리 방식은 TO 선제 설계로 전환합니다.',
    '핵심은 사람(PO)을 줄이는 것이 아니라 일의 기준(TO)을 낮추고, PO는 자연감소와 재배치로 따라오게 하는 것입니다.',
  ].join('\n\n'));
}

// =====================================================================
// Slide 2. '30 추진전략 : 목표 / 4개 축 TO 계획·3대 Lever / 실행기반 / Roadmap
// =====================================================================
function slide2() {
  const s = pptx.addSlide();
  header(s, "Ⅱ. '30 인당생산성 향상 추진전략 및 Roadmap", '원가 ④ 인당생산성 향상',
    `'30년까지 Total 인력(TO)을 '24년 말 대비 ${dNum(D_TOTAL)}명(${dPct(D_TOTAL_PCT)}) 감축 ─ S직군 직영 전환 등에 따른 ` +
    `직영 증가(${dNum(D_JIK)}명)를 협력 감축(${dNum(D_HYUP)}명)으로 흡수하는 정규·S직군·계약·협력 4개 축 통합 관리 추진`);

  // ---- 목표 ----
  TB(s, "'30 목표", { x: 0.4, y: 1.66, w: 1.3, h: 0.58, align: 'center', fontSize: 13, bold: true, color: C.white, fill: { color: C.navy }, line: { color: C.navy, width: 0 } });
  [
    { k: 'Total 인력(TO)\n4개 축 합산', v: `'24.E比 ${dPct(D_TOTAL_PCT)}`, d: `${fmt(TOTAL[0])} → ${fmt(TOTAL[LAST])}명 (${dNum(D_TOTAL)}명)` },
    { k: '인당 생산성', v: "'24.E比 OO%↑", d: `생산량 동일 가정 時 +${PROD_UP.toFixed(1)}%` },
    { k: 'Total 노무비', v: '매출액比 OO% 이내', d: '4개 축 노무비 통합 KPI 관리' },
  ].forEach((k, i) => {
    const x = 1.78 + i * 3.73;
    R(s, { x, y: 1.66, w: 3.66, h: 0.58, fill: { color: C.sky2 }, line: { color: C.sky, width: 1 } });
    T(s, k.k, { x: x + 0.06, y: 1.66, w: 1.26, h: 0.58, fontSize: 9, bold: true, color: C.navy, align: 'center', margin: 0 });
    seg(s, x + 1.35, 1.75, x + 1.35, 2.15, { color: C.line, width: 0.75 });
    T(s, k.v, { x: x + 1.42, y: 1.67, w: 2.2, h: 0.32, fontSize: 15, bold: true, color: C.orange, margin: 0 });
    T(s, k.d, { x: x + 1.42, y: 1.98, w: 2.2, h: 0.22, fontSize: 8.5, color: C.sub, margin: 0 });
  });

  // ---- 좌 : 4개 축 Total 인력(TO) 운영계획 + 3대 감축 Lever ----
  secHead(s, "Total 인력(TO) 운영계획 및 3대 감축 Lever  ('24년 말 기준)", 0.4, 2.32, 7.55, 0.28, 'left', 10);

  // 범례
  let lx = 0.45;
  SERIES.forEach((se, i) => {
    const pending = i === 1 && !SJ_READY;
    const label = pending ? `${se.name}(정규직 內 포함)` : se.name;
    const w = pending ? 1.45 : i === 3 ? 0.6 : 1.0;
    R(s, { x: lx, y: 2.665, w: 0.12, h: 0.12, fill: { color: pending ? C.white : se.color }, line: { color: se.color, width: 0.75, dashType: pending ? 'dash' : 'solid' } });
    T(s, label, { x: lx + 0.16, y: 2.64, w: w - 0.16, h: 0.17, fontSize: 7.5, color: pending ? C.mute : C.text, margin: 0 });
    lx += w;
  });
  T(s, '(단위: 명)', { x: 4.6, y: 2.64, w: 0.95, h: 0.17, fontSize: 7.5, color: C.mute, align: 'right', margin: 0 });

  // 연도별 누적 막대
  const CH = { x: 0.45, w: 5.1, base: 4.2, hmax: 1.2 };
  const slotW = CH.w / YEARS.length, barW = 0.46, k = CH.hmax / Math.max(...TOTAL);
  YEARS.forEach((yr, i) => {
    const x = CH.x + i * slotW + (slotW - barW) / 2;
    let top = CH.base;
    SERIES.forEach((se) => {
      const h = se.v[i] * k;
      if (h <= 0) return;
      top -= h;
      R(s, { x, y: top, w: barW, h, fill: { color: se.color }, line: { color: C.white, width: 0.5 } });
      if (h >= 0.13) T(s, fmt(se.v[i]), { x: x - 0.05, y: top, w: barW + 0.1, h, fontSize: 7, color: se.label, align: 'center', margin: 0 });
    });
    const edge = i === 0 || i === LAST;
    T(s, fmt(TOTAL[i]), { x: x - 0.12, y: top - 0.17, w: barW + 0.24, h: 0.16, fontSize: 8, bold: true, color: edge ? C.orangeD : C.navy, align: 'center', margin: 0 });
    T(s, yr, { x: CH.x + i * slotW, y: CH.base + 0.02, w: slotW, h: 0.16, fontSize: 8, bold: edge, color: C.text, align: 'center', margin: 0 });
  });
  seg(s, CH.x, CH.base, CH.x + CH.w, CH.base, { color: C.line, width: 1 });

  // '24.E → '30.E 증감표
  const cell = (text, o = {}) => ({ text, options: Object.assign({ fontFace: FONT, fontSize: 8, color: C.text, align: 'center', valign: 'middle' }, o) });
  const border = (left) => [{ type: 'solid', pt: 0.5, color: C.line }, { type: 'solid', pt: 0.5, color: C.line }, { type: 'solid', pt: 0.5, color: C.line }, left];
  const dCell = (n, bold = true) => cell(dNum(n), { bold, color: n < 0 ? C.orangeD : C.blue });
  const rowsT = [[
    cell('구분', { bold: true, color: C.white, fill: { color: C.navy } }),
    cell("'24.E", { bold: true, color: C.white, fill: { color: C.navy } }),
    cell("'30.E", { bold: true, color: C.white, fill: { color: C.navy } }),
    cell('증감', { bold: true, color: C.white, fill: { color: C.navy } }),
  ]];
  SERIES.forEach((se, i) => {
    const b = border({ type: 'solid', pt: 3, color: se.color });
    const lab = cell(se.name, { align: 'left', border: b, bold: true });
    if (i === 1 && !SJ_READY) {
      rowsT.push([lab, cell('정규직 內 포함 (분리 後 반영)', { colspan: 3, fontSize: 7, color: C.mute })]);
    } else {
      rowsT.push([lab, cell(fmt(se.v[0])), cell(fmt(se.v[LAST])), dCell(se.v[LAST] - se.v[0])]);
    }
  });
  const totFill = { fill: { color: C.sky } };
  rowsT.push([
    cell('총계', Object.assign({ bold: true, color: C.navy, align: 'left' }, totFill)),
    cell(fmt(TOTAL[0]), Object.assign({ bold: true }, totFill)),
    cell(fmt(TOTAL[LAST]), Object.assign({ bold: true }, totFill)),
    cell(dNum(D_TOTAL), Object.assign({ bold: true, color: C.orangeD }, totFill)),
  ]);
  s.addTable(rowsT, {
    x: 5.68, y: 2.64, w: 2.27, colW: [0.74, 0.5, 0.5, 0.53], rowH: 0.29,
    border: { type: 'solid', pt: 0.5, color: C.line }, margin: [0.01, 0.04, 0.01, 0.05],
  });

  // 3대 감축 Lever
  [
    { t: '① 공정·프로세스 자동화', b: ['Intelligent Factory·AI/로봇 적용 확대', 'Inno TF 설비정예화 연계 운전·정비 TO↓'] },
    { t: '② 저효율 제조라인 효율화(폐쇄)', b: ['저수익·저가동 라인 통폐합/폐쇄', '잉여인력 신사업·해외법인 전환배치'] },
    { t: '③ 개인별 업무 Span 확대', b: ['직무통합·다기능화, 협력사間 중복기능 통합', '사무·지원업무 AI Agent 활용(WX)'] },
  ].forEach((c, i) => {
    const x = 0.4 + i * 2.55;
    TB(s, [
      run(c.t, { fontSize: 9, bold: true, color: C.orangeD }),
      run('  △OO명', { fontSize: 8, bold: true, color: C.navy }),
    ], { x, y: 4.44, w: 2.45, h: 0.23, fill: { color: C.orangeL }, line: { color: C.orange, width: 0.75 }, margin: 4 });
    TB(s, c.b.map((b, j) => run('· ' + b, { fontSize: 8, color: C.text, breakLine: j < c.b.length - 1 })),
      { x, y: 4.67, w: 2.45, h: 0.45, valign: 'middle', margin: 3, line: { color: C.line, width: 0.75 } });
  });

  // ---- 우 : TO-PO 운영 원칙 (개념도) ----
  const rx = 8.1, rw = 4.83;
  secHead(s, 'TO-PO 운영 원칙  (개념도)', rx, 2.32, rw, 0.28, 'left', 10);
  seg(s, 8.3, 2.73, 8.6, 2.73, { color: C.navy, width: 2.25 });
  T(s, 'TO(기준인력) : 선행 감축', { x: 8.65, y: 2.64, w: 1.8, h: 0.18, fontSize: 7.5, bold: true, color: C.navy, margin: 0 });
  seg(s, 10.6, 2.73, 10.9, 2.73, { color: C.orange, width: 2, dashType: 'dash' });
  T(s, 'PO(실제인력) : 후행 수렴', { x: 10.95, y: 2.64, w: 1.95, h: 0.18, fontSize: 7.5, bold: true, color: C.orangeD, margin: 0 });

  const ox = 8.3, oy = 4.0;
  seg(s, ox, 2.9, ox, oy, { color: C.line, width: 1 });
  seg(s, ox, oy, 12.9, oy, { color: C.line, width: 1 });
  const px = YEARS.map((_, i) => 8.6 + i * 0.683);
  const TOy = TOTAL.map((t) => 3.0 + ((TOTAL[0] - t) / -D_TOTAL) * 0.8); // 실제 TO 계획 추이를 반영한 개념선
  const POy = [2.98, 3.0, 2.99, 3.06, 3.2, 3.37, 3.7];
  YEARS.forEach((yr, i) => T(s, yr.replace('.E', ''), { x: px[i] - 0.25, y: oy + 0.02, w: 0.5, h: 0.18, fontSize: 8, color: C.sub, align: 'center', margin: 0 }));
  for (let i = 0; i < LAST; i++) {
    seg(s, px[i], TOy[i], px[i + 1], TOy[i + 1], { color: C.navy, width: 2.25 });
    seg(s, px[i], POy[i], px[i + 1], POy[i + 1], { color: C.orange, width: 2, dashType: 'dash' });
  }
  px.forEach((x, i) => {
    s.addShape(pptx.ShapeType.ellipse, { x: x - 0.04, y: TOy[i] - 0.04, w: 0.08, h: 0.08, fill: { color: C.navy }, line: { color: C.white, width: 0.5 } });
  });
  s.addShape(pptx.ShapeType.upDownArrow, { x: px[4] - 0.06, y: 3.23, w: 0.12, h: 0.15, fill: { color: C.mute }, line: { color: C.mute, width: 0 } });
  T(s, [
    run('과원(PO > TO) 발생 구간', { fontSize: 7.5, bold: true, color: C.sub, breakLine: true }),
    run('→ 자연감소·재배치로 단계적 해소', { fontSize: 7.5, color: C.sub }),
  ], { x: 8.45, y: 3.5, w: 2.1, h: 0.36, margin: 0 });

  TB(s, [
    run('· TO 선행 감축 → 정년퇴직 등 자연감소 범위 內 채용 조절로 PO 수렴', { fontSize: 8, color: C.text, breakLine: true }),
    run('· 잉여인력 Re-skilling 後 신사업·해외(인도 2기, 美 H社 합작) 재배치', { fontSize: 8, color: C.text, breakLine: true }),
    run('· 인위적 PO 조정 지양, 구조적 TO 관리로 노사 수용성 확보', { fontSize: 8, color: C.text, breakLine: true }),
    run('· 조직혁신 TF 주관, 반기 단위 4개 축 TO·노무비 Rolling 점검', { fontSize: 8, color: C.text }),
  ], { x: rx, y: 4.3, w: rw, h: 0.82, margin: 5, fill: { color: C.gray }, line: { color: C.line, width: 0.75 } });

  // ---- 실행기반 (Enabler) ----
  TB(s, '실행기반\n(Enabler)', { x: 0.4, y: 5.2, w: 1.0, h: 0.97, align: 'center', fontSize: 10, bold: true, color: C.white, fill: { color: C.navy }, line: { color: C.navy, width: 0 }, margin: 0 });
  [
    { t: '① 노무비 기반 인력관리체계 전환', tag: '신규', b: ['4개 축 Total 노무비 Baseline·KPI 정립','TO 증감의 노무비 환산, 사업부별 목표 부여', `계약직 증가(${dNum(DATA.gye[LAST] - DATA.gye[0])}명) 등 고용형태별 적정성 점검`] },
    { t: '② 노무 Risk 선제 대응', tag: "'25 전략2", b: ['통상임금·임금차액 소송 매몰비용 최소화', '노란봉투법 대비 협력/용역계약 유형별 대응', '정년연장·근로시간 규제 대비 제도 재정비'] },
    { t: '③ 인재 전환·재배치', tag: "'25 전략1", b: ['Domain&AI 전문인력 양성, 뉴칼라 2.0 개편', '잉여인력 Re-skilling 後 전략분야 재배치', 'S직군 조기 조직안정·직원 간 화합 도모'] },
    { t: '④ 일하는 방식·현장 리더십', tag: "'25 전략3", b: ['AI 기반 일하는 방식(Workway) 재정립·확산', '직책자 역할·권한 재정립, 솔선수범 리더십', '全 직원 DX 이해 증진 및 변화관리 강화'] },
  ].forEach((e, i) => {
    const x = 1.48 + i * 2.88, w = 2.79;
    TB(s, [
      run(e.t, { fontSize: 9, bold: true, color: C.navy }),
      run('  [' + e.tag + ']', { fontSize: 7.5, bold: true, color: e.tag === '신규' ? C.orangeD : C.mute }),
    ], { x, y: 5.2, w, h: 0.26, fill: { color: C.sky }, line: { color: C.line, width: 0.75 }, margin: 4 });
    TB(s, e.b.map((b, j) => run('· ' + b, { fontSize: 8, color: C.text, breakLine: j < e.b.length - 1 })),
      { x, y: 5.46, w, h: 0.71, margin: 4, line: { color: C.line, width: 0.75 } });
  });

  // ---- Roadmap ----
  const peak = TOTAL.map((_, i) => DATA.jeong[i] + DATA.gye[i]);
  const peakIdx = peak.indexOf(Math.max(...peak));
  TB(s, '추진\nRoadmap', { x: 0.4, y: 6.25, w: 1.0, h: 0.8, align: 'center', fontSize: 10, bold: true, color: C.white, fill: { color: C.navy }, line: { color: C.navy, width: 0 }, margin: 0 });
  [
    { t: "'26   기반 구축", c: C.blue, b: ['4개 축 Total TO·노무비 Baseline 및 관리체계 정립', "S직군 직영 전환('26.4월~) 조기 조직안정"] },
    { t: "'27~'28   본격 실행", c: C.mid, b: [`직영 TO Peak(${YEARS[peakIdx].replace('.E', '')}, ${fmt(peak[peakIdx])}명) 이후 단계적 감축`, `협력 ${fmt(DATA.hyup[3])}명 수준 재편, 자동화·라인 효율화 연계`] },
    { t: "'29~'30   정착·고도화", c: C.navy, b: ['人+AI/로보틱스 Hybrid 인력구조 정착', `'30 Total TO ${fmt(TOTAL[LAST])}명(${dPct(D_TOTAL_PCT)}) 달성`] },
  ].forEach((p, i) => {
    const x = 1.48 + i * 3.76;
    TB(s, p.t, { shape: i === 0 ? pptx.ShapeType.homePlate : pptx.ShapeType.chevron, x, y: 6.25, w: 3.9, h: 0.3, align: 'center', fontSize: 9.5, bold: true, color: C.white, fill: { color: p.c }, line: { color: p.c, width: 0 }, margin: 0 });
    T(s, p.b.map((b, j) => run('· ' + b, { fontSize: 8, color: C.text, breakLine: j < p.b.length - 1 })),
      { x: x + 0.2, y: 6.58, w: 3.5, h: 0.46, valign: 'top', margin: 1 });
  });

  footer(s, "※ '24.E: '24년 말 기준 인력 운영계획(TO)" +
    (SJ_READY ? '' : '  /  직영 S직군은 분리 수치 확정 後 별도 반영(현재 정규직 內 포함)') +
    '  /  Lever별 감축 규모(OO)는 사업계획 확정 後 반영', 2);

  s.addNotes([
    `[발표 포인트] '24년 말 기준 Total 인력(TO)은 ${fmt(TOTAL[0])}명이며, '30년까지 ${fmt(TOTAL[LAST])}명으로 ${fmt(D_TOTAL)}명(${Math.abs(D_TOTAL_PCT).toFixed(1)}%) 감축하는 것이 목표입니다.`,
    `4개 축으로 보면 직영은 S직군 전환과 계약직 증가로 ${dNum(D_JIK)}명 늘어나지만, 협력이 ${dNum(D_HYUP)}명 줄어 총량은 감소합니다. 직영만 보면 늘어나는 것처럼 보이기 때문에 Total 관점 관리가 필요합니다.`,
    `직영은 ${YEARS[peakIdx]} ${fmt(peak[peakIdx])}명으로 정점을 찍은 뒤, 자동화·라인 효율화·업무 Span 확대의 3대 Lever로 단계적으로 줄어듭니다. 생산량이 동일하다면 인당생산성은 약 ${PROD_UP.toFixed(1)}% 향상되는 효과입니다.`,
    'TO를 먼저 낮추고, PO는 정년퇴직 등 자연감소와 Re-skilling 후 신사업·해외 재배치를 통해 단계적으로 수렴시켜 노사 수용성을 확보합니다.',
  ].join('\n\n'));
}

// =====================================================================
// Slide 3. 2035 HR 미래 청사진 : Vision / 3대 축('35 모습·'30 연계) / 불확실성 대응
// =====================================================================
const V35 = 20000; // '35 Total 인력 '2만명 수준' (방향성 수치)
const D35_PCT = ((V35 - TOTAL[0]) / TOTAL[0]) * 100;
const PROD35 = TOTAL[0] / V35; // 생산량 동일 가정 時 인당생산성 배수

function slide3() {
  const s = pptx.addSlide();
  header(s, 'Ⅲ. 2035 HR 미래 청사진 (Future Blueprint)', "'30 이후 HR 지향점",
    "'35년 『Human+AI 협업 기반 정예·자율 전문가 조직』 지향 ─ 기술·규제·시황 불확실성을 감안하여 확정 목표가 아닌 " +
    '방향성(North Star)으로 제시하고, 매년 Rolling 時 핵심 변수(Signpost) 점검을 통해 경로 보정',
    "HR 중기전략 ('26~'30 Rolling)  |  2035 장기 비전");

  // ---- Vision ----
  TB(s, '2035\nHR Vision', { x: 0.4, y: 1.68, w: 1.3, h: 0.52, align: 'center', fontSize: 11, bold: true, color: C.white, fill: { color: C.orange }, line: { color: C.orange, width: 0 }, margin: 0 });
  TB(s, [
    run('Human+AI 협업으로 성과를 창출하는 ', { fontSize: 16, bold: true, color: C.white }),
    run('정예·자율 전문가 조직', { fontSize: 16, bold: true, color: 'FFC7A8' }),
  ], { x: 1.7, y: 1.68, w: 11.23, h: 0.52, align: 'center', fill: { color: C.navy }, line: { color: C.navy, width: 0 } });

  // ---- 3대 축 ----
  const PILLARS = [
    {
      name: '① 조직 & WX (일하는 방식)', dir: 'AX를 경영성과로 연결하는 Human+AI 조직·운영체계 혁신',
      from: '공정·부서별 分業', to: '가치흐름 통합 자율운영',
      a: { tag: '제조\n분야', title: 'AI 기반 제철소 자율제조', b: ['제철소 全 공정을 하나의 흐름으로 통합 운영', '공정별 조직 → 전문성 조직 전환, 전체 최적화', '사람은 감시·판단·예외대응 중심 역할 수행'] },
      b: { tag: '사무\n분야', title: '지능형 업무자율화', b: ['End-to-End 가치흐름 통합, 부서 경계 최소화', '인간+AI(Agent) 팀 중심 조직 운영', '정형업무 AI 자율처리, 사람은 기획·판단 집중'] },
      link: ['공정·프로세스 자동화, AI Agent 활용 업무 Span 확대', 'Inno TF 설비정예화 연계 → 자율제조 전환 기반 확보'],
    },
    {
      name: '② HR 제도', dir: '성과·역할 중심 유연 인사체계 기반 창의적 인재 육성',
      from: '중앙통제·직급 중심', to: '자율·책임·역할 중심',
      a: { tag: '운영\n방식', title: 'HR Decentralization', b: ["'중앙통제' → '자율·책임 기반 성과 창출 지원'", '현업 리더 주도 인사운영, HR은 기준·데이터 지원', '성과 창출 방식 변화에 대응하는 창의적 인재 육성'] },
      b: { tag: '인사\n체계', title: '전문성·역할 중심 인력 운영', b: ['수평화된 직급/호칭, 직무급제 기반 보상', '高직급 적체 해소 完, MZ세대가 주축 계층 리딩', '직무 전문성 기반 경력경로(전문가 트랙) 운영'] },
      link: ['뉴칼라 2.0 개편, Domain&AI 전문인력 양성', '직책자 역할·권한 재정립, 정년연장 대비 제도 정비'],
    },
    {
      name: '③ 인력 규모', dir: '인당생산성 향상 포화 下 AI/로봇 협업 기반 최적 규모 운영',
      from: '인원 감축(Volume)', to: 'Human+AI 최적 규모 안정화',
      b: { tag: '운영\n모습', title: 'Human+AI/로봇 합동 근무', b: ['全 직군 인력 정예화, 재채용 계약직 대폭 축소', '자동화·무인화 투자 성과 창출, 최적 효율 인력', 'AGI·AI/로봇과 사람이 합동 근무하는 운영체계'] },
      link: [`Total TO ${fmt(TOTAL[LAST])}명(${dPct(D_TOTAL_PCT)}), TO·노무비 통합관리`, "'31년~ 감축 속도 조절, 자연감소 연동 규모 안정화"],
    },
  ];
  const cw = 4.09, gap = 0.13;
  const Y = { head: 2.28, shift: 2.86, a: 3.2, b: 4.19, link: 5.2 }, subH = 0.95;
  const bullets = (arr, size = 8) => arr.map((b, j) => run('· ' + b, { fontSize: size, color: C.text, breakLine: j < arr.length - 1 }));
  function subBox(x, y, d) {
    TB(s, d.tag, { x, y, w: 0.6, h: subH, align: 'center', margin: 0, fontSize: 8.5, bold: true, color: C.navy, fill: { color: C.sky }, line: { color: C.line, width: 0.75 } });
    R(s, { x: x + 0.6, y, w: cw - 0.6, h: subH });
    T(s, [run(d.title, { fontSize: 9.5, bold: true, color: C.navy, breakLine: true })].concat(bullets(d.b)),
      { x: x + 0.66, y, w: cw - 0.7, h: subH, valign: 'middle', margin: 2 });
  }

  PILLARS.forEach((p, i) => {
    const x = 0.4 + i * (cw + gap);
    TB(s, [
      run(p.name, { fontSize: 12, bold: true, color: C.white, breakLine: true }),
      run(p.dir, { fontSize: 8.5, color: C.sky }),
    ], { x, y: Y.head, w: cw, h: 0.54, align: 'center', margin: 2, fill: { color: C.mid }, line: { color: C.mid, width: 0 } });
    TB(s, [
      run(p.from, { fontSize: 8.5, color: C.sub }),
      run('   ▶   ', { fontSize: 8, bold: true, color: C.orange }),
      run(p.to, { fontSize: 9.5, bold: true, color: C.orangeD }),
    ], { x, y: Y.shift, w: cw, h: 0.3, align: 'center', margin: 0, fill: { color: C.orangeL }, line: { color: C.orange, width: 0.75 } });

    if (p.a) {
      subBox(x, Y.a, p.a);
    } else {
      // 인력 규모 : Total TO 추이 ('24.E → '30.E → '35)
      const ay = Y.a;
      TB(s, '인력\n규모', { x, y: ay, w: 0.6, h: subH, align: 'center', margin: 0, fontSize: 8.5, bold: true, color: C.navy, fill: { color: C.sky }, line: { color: C.line, width: 0.75 } });
      R(s, { x: x + 0.6, y: ay, w: cw - 0.6, h: subH });
      T(s, '全 직군 Total 인력(TO) 2만명 수준 안정화', { x: x + 0.66, y: ay + 0.03, w: cw - 0.7, h: 0.22, fontSize: 9.5, bold: true, color: C.navy, margin: 2 });
      const bx = x + 1.1, maxLen = 1.0;
      [
        { yr: "'24.E", v: TOTAL[0], lab: fmt(TOTAL[0]), c: C.navy },
        { yr: "'30.E", v: TOTAL[LAST], lab: `${fmt(TOTAL[LAST])} (${dPct(D_TOTAL_PCT)})`, c: C.blue },
        { yr: "'35", v: V35, lab: `약 2만 (△약 ${Math.round(Math.abs(D35_PCT))}%)`, dash: true },
      ].forEach((r, j) => {
        const ry = ay + 0.31 + j * 0.205, len = (r.v / TOTAL[0]) * maxLen;
        T(s, r.yr, { x: x + 0.66, y: ry, w: 0.42, h: 0.16, fontSize: 7.5, bold: r.dash, color: C.text, align: 'right', margin: 0 });
        R(s, r.dash
          ? { x: bx, y: ry, w: len, h: 0.16, fill: { color: C.sky2 }, line: { color: C.blue, width: 1, dashType: 'dash' } }
          : { x: bx, y: ry, w: len, h: 0.16, fill: { color: r.c }, line: { color: r.c, width: 0 } });
        T(s, r.lab, { x: bx + len + 0.04, y: ry, w: 1.0, h: 0.16, fontSize: 7.5, bold: true, color: r.dash ? C.orangeD : C.navy, margin: 0 });
      });
      seg(s, x + 2.98, ay + 0.32, x + 2.98, ay + 0.88, { color: C.line, width: 0.75 });
      T(s, [
        run('인당생산성', { fontSize: 7.5, bold: true, color: C.navy, breakLine: true }),
        run(`약 ${PROD35.toFixed(1)}배`, { fontSize: 13, bold: true, color: C.orange, breakLine: true }),
        run("'24.E比 (생산량 동일)", { fontSize: 6.5, color: C.sub }),
      ], { x: x + 3.0, y: ay + 0.28, w: cw - 3.02, h: 0.64, align: 'center', margin: 0 });
    }
    subBox(x, Y.b, p.b);

    TB(s, "'30 전략\n연계", { x, y: Y.link, w: 0.6, h: 0.44, align: 'center', margin: 0, fontSize: 8, bold: true, color: C.white, fill: { color: C.blue }, line: { color: C.blue, width: 0 } });
    TB(s, bullets(p.link), { x: x + 0.6, y: Y.link, w: cw - 0.6, h: 0.44, margin: 4, fill: { color: C.gray }, line: { color: C.line, width: 0.75 } });
  });

  // ---- 불확실성 대응 ----
  const by = 5.74;
  TB(s, '불확실성\n대응', { x: 0.4, y: by, w: 1.0, h: 1.34, align: 'center', fontSize: 10, bold: true, color: C.white, fill: { color: C.navy }, line: { color: C.navy, width: 0 }, margin: 0 });

  TB(s, "핵심 불확실성 (Signpost)  ─  '35 경로 보정 Trigger로 매년 점검", { x: 1.48, y: by, w: 6.5, h: 0.24, align: 'left', fontSize: 9, bold: true, color: C.navy, fill: { color: C.sky }, line: { color: C.sky, width: 0 }, margin: 6 });
  [
    { t: 'AI·로봇 기술 성숙 속도', d: 'AGI 상용화, 자율제조 적용 범위·투자 회수 시점', imp: '①③' },
    { t: '노동 규제·노사 환경', d: '정년연장, 노란봉투법, 근로시간 단축 등 유연성 제약', imp: '②③' },
    { t: '철강 시황·사업 포트폴리오', d: '저수익 장기화, 해외거점(인도·美) 확대 및 사업재편', imp: '①③' },
    { t: '인구·세대 구조 변화', d: '생산가능인구 감소, 베이비부머 은퇴·MZ 주축 전환', imp: '②③' },
  ].forEach((u, i) => {
    const ux = 1.48 + (i % 2) * 3.28, uy = by + 0.28 + Math.floor(i / 2) * 0.55, uw = 3.22;
    R(s, { x: ux, y: uy, w: uw, h: 0.51 });
    T(s, u.t, { x: ux + 0.06, y: uy + 0.03, w: 2.3, h: 0.22, fontSize: 9, bold: true, color: C.text, margin: 1 });
    TB(s, '영향 ' + u.imp, { x: ux + uw - 0.72, y: uy + 0.06, w: 0.64, h: 0.17, align: 'center', margin: 0, fontSize: 7, bold: true, color: C.orangeD, fill: { color: C.orangeL }, line: { color: C.orange, width: 0.5 } });
    T(s, u.d, { x: ux + 0.06, y: uy + 0.26, w: uw - 0.1, h: 0.22, fontSize: 8, color: C.sub, margin: 1 });
  });

  TB(s, '경로 관리 원칙', { x: 8.08, y: by, w: 4.85, h: 0.24, align: 'left', fontSize: 9, bold: true, color: C.navy, fill: { color: C.sky }, line: { color: C.sky, width: 0 }, margin: 6 });
  [
    { k: 'No-regret', v: '시나리오 무관 선행 추진 ─ 직무·역할 중심 제도, AI 인재 확보, Total TO·노무비 관리체계' },
    { k: '조건부 확대', v: '기술 성숙·투자 성과 확인 後 자율제조·무인화 적용 범위, 조직 통합 속도 단계적 확대' },
    { k: 'Rolling 보정', v: "기술 가속 時 조기 달성, 지연·규제 강화 時 속도 조절 ─ 매년 '35 규모·시점 재설정" },
  ].forEach((r, i) => {
    const ry = by + 0.28 + i * 0.365;
    TB(s, r.k, { x: 8.08, y: ry, w: 0.95, h: 0.35, align: 'center', margin: 0, fontSize: 8.5, bold: true, color: C.orangeD, fill: { color: C.orangeL }, line: { color: C.orange, width: 0.75 } });
    TB(s, r.v, { x: 9.03, y: ry, w: 3.9, h: 0.35, fontSize: 8, color: C.text, margin: 1.5, line: { color: C.line, width: 0.75 } });
  });

  footer(s, "※ '35 청사진은 확정 목표가 아닌 방향성(North Star)으로, 규모·시점은 매년 중기전략 Rolling 時 재점검  /  " +
    "인당생산성: 생산량 동일 가정 時 Total TO('24.E) 대비 산출", 3);

  s.addNotes([
    "[발표 포인트] 앞의 '30 Rolling이 '어떻게 줄일 것인가'였다면, 이 장은 '30 이후 HR이 도달할 모습입니다. AI 기술·노동 규제·철강 시황의 불확실성이 커서 확정 목표가 아닌 방향성으로 제시드립니다.",
    '조직·WX: 제철소는 공정별로 나뉜 운영에서 하나의 흐름으로 통합된 AI 자율제조로 전환되고, 사람은 감시·판단·예외대응을 맡습니다. 사무는 End-to-End 가치흐름 단위로 인간+AI 팀이 일합니다.',
    'HR 제도: 중앙통제형 HR에서 현업이 자율과 책임으로 성과를 내도록 지원하는 체계로 바뀝니다. 직급·호칭은 수평화되고 직무급 기반으로 운영되며, 고직급 적체가 해소되어 MZ세대가 조직을 이끄는 주축이 됩니다.',
    `인력 규모: '24년 말 ${fmt(TOTAL[0])}명에서 '30년 ${fmt(TOTAL[LAST])}명, '35년 2만명 수준으로 안정화됩니다. '30년까지가 구조적 감축기라면, 이후는 인당생산성 향상이 포화되면서 Human+AI 최적 규모를 유지하는 단계입니다. 생산량이 같다면 인당생산성은 약 ${PROD35.toFixed(1)}배입니다.`,
    '불확실성은 네 가지 변수로 매년 점검합니다. 어떤 시나리오에서도 필요한 직무·역할 중심 제도, AI 인재 확보, TO·노무비 관리체계는 먼저 추진하고, 자율제조·무인화 범위는 기술 성숙과 투자 성과를 확인하며 단계적으로 넓히겠습니다.',
  ].join('\n\n'));
}

if (!ONLY || ONLY === 1) slide1();
if (!ONLY || ONLY === 2) slide2();
if (!ONLY || ONLY === 3) slide3();

pptx.writeFile({ fileName: OUT }).then((f) => console.log('saved:', f));
