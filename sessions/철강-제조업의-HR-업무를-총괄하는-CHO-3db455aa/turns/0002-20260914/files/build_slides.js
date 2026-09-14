// HR 중기전략('26~'30 Rolling) 2장 PPTX 생성 스크립트
// 실행: node build_slides.js  (pptxgenjs 3.x 필요)
const path = require('path');
const pptxgen = require('pptxgenjs');

const ONLY = process.env.ONLY_SLIDE ? Number(process.env.ONLY_SLIDE) : 0;
const OUT = process.env.OUT_FILE || path.join(__dirname, 'HR_중기전략_2030_인당생산성_Rolling.pptx');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5 in
pptx.title = "HR 중기전략('26~'30 Rolling) - 인당생산성 향상";

const FONT = '맑은 고딕';
const C = {
  navy: '0B2D5B', blue: '1F5FA8', mid: '16498A', sky: 'D6E4F5', sky2: 'EEF3FA',
  gray: 'F2F4F7', line: 'BFC7D2', text: '222222', sub: '555555', mute: '8A94A3',
  orange: 'E8541E', orangeD: 'C2410C', orangeL: 'FDEDE6', light: '8FB3E0', white: 'FFFFFF',
};

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

function header(slide, title, tag, msg) {
  T(slide, "HR 중기전략 ('26~'30 Rolling)  |  원가 경쟁력", { x: 0.4, y: 0.16, w: 6, h: 0.22, fontSize: 9, color: C.mute, margin: 0 });
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
      key: '관리\n대상', asis: '직영/협력\n구분 관리', tobe: '직영+협력\nTotal 인력 통합 관리',
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

  footer(s, '※ TO: 기준인력, PO: 실제 운영인력, S직군: 근지위 소송·특별채용을 통해 협력사에서 직영으로 전환된 직군', 1);

  s.addNotes([
    "[발표 포인트] 작년 수립한 HR 중기전략 3대 방향은 유지하되, 올해는 전사 '원가' 축의 ④ 인당생산성 향상을 '30년까지 어떻게 구체화할 것인가의 관점에서 Rolling했습니다.",
    '작년 대비 달라진 여건은 네 가지입니다. ① S직군 직영 전환으로 직영/협력 구분 관리가 무의미해졌고, ② 통상임금 확대로 인원수보다 노무비 단가가 원가를 좌우하며, ③ AI·로보틱스가 노동력을 대체하고, ④ 인력규모를 인위적으로 조정하기 어려워졌습니다.',
    '이에 따라 관리 대상은 Total 인력, 관리 기준은 노무비, 관리 범위는 AI·로보틱스를 포함한 Total Workforce, 관리 방식은 TO 선제 설계로 전환합니다.',
    '핵심은 사람(PO)을 줄이는 것이 아니라 일의 기준(TO)을 낮추고, PO는 자연감소와 재배치로 따라오게 하는 것입니다.',
  ].join('\n\n'));
}

// =====================================================================
// Slide 2. '30 추진전략 : 목표 / 3대 Lever / 실행기반 / Roadmap
// =====================================================================
function slide2() {
  const s = pptx.addSlide();
  header(s, "Ⅱ. '30 인당생산성 향상 추진전략 및 Roadmap", '원가 ④ 인당생산성 향상',
    "'30년까지 Total TO(직영+협력) OO% 감축을 목표로 3대 TO 감축 Lever를 단계적으로 실행하고, " +
    '노무비 기반 관리체계 전환·노무 Risk 대응 등 실행기반을 병행 구축');

  // ---- 목표 ----
  TB(s, "'30 목표", { x: 0.4, y: 1.68, w: 1.3, h: 0.62, align: 'center', fontSize: 13, bold: true, color: C.white, fill: { color: C.navy }, line: { color: C.navy, width: 0 } });
  [
    { k: 'Total TO\n(직영+협력)', v: "'25比 OO%↓", d: '직영 OO%↓  /  협력 OO%↓' },
    { k: '인당 생산성', v: "'25比 OO%↑", d: '인당 생산량(톤/인) 기준' },
    { k: 'Total 노무비', v: '매출액比 OO% 이내', d: '직영+협력 노무비 통합 KPI 관리' },
  ].forEach((k, i) => {
    const x = 1.78 + i * 3.73;
    R(s, { x, y: 1.68, w: 3.66, h: 0.62, fill: { color: C.sky2 }, line: { color: C.sky, width: 1 } });
    T(s, k.k, { x: x + 0.08, y: 1.68, w: 1.22, h: 0.62, fontSize: 9.5, bold: true, color: C.navy, align: 'center', margin: 0 });
    seg(s, x + 1.35, 1.78, x + 1.35, 2.2, { color: C.line, width: 0.75 });
    T(s, k.v, { x: x + 1.42, y: 1.7, w: 2.2, h: 0.34, fontSize: 15, bold: true, color: C.orange, margin: 0 });
    T(s, k.d, { x: x + 1.42, y: 2.03, w: 2.2, h: 0.24, fontSize: 8.5, color: C.sub, margin: 0 });
  });

  // ---- 좌 : 3대 TO 감축 Lever (Waterfall 개념도) ----
  secHead(s, '3대 TO 감축 Lever  (Total TO Waterfall 개념도)', 0.4, 2.42, 7.55, 0.3, 'left', 10);
  const base = 3.75, slot = 7.55 / 6, bw = 0.6;
  const bx = (i) => 0.4 + i * slot + (slot - bw) / 2;
  seg(s, 0.45, base, 7.9, base, { color: C.line, width: 1 });
  const stack = (i, jik, hyup) => { // 직영(하단) + 협력(상단)
    R(s, { x: bx(i), y: base - jik, w: bw, h: jik, fill: { color: C.navy }, line: { color: C.navy, width: 0 } });
    R(s, { x: bx(i), y: base - jik - hyup, w: bw, h: hyup, fill: { color: C.light }, line: { color: C.light, width: 0 } });
    if (jik >= 0.2) T(s, '직영', { x: bx(i), y: base - jik, w: bw, h: jik, fontSize: 8, bold: true, color: C.white, align: 'center', margin: 0 });
    if (hyup >= 0.18) T(s, '협력', { x: bx(i), y: base - jik - hyup, w: bw, h: hyup, fontSize: 8, bold: true, color: C.navy, align: 'center', margin: 0 });
  };
  const tops = [2.95, 2.95, 3.13, 3.27, 3.37]; // 각 단계 후 Total TO 수준(개념)
  stack(0, 0.44, 0.36);
  stack(1, 0.6, 0.2);
  T(s, '총량 불변(구성 변화)', { x: 0.4 + slot, y: 2.75, w: slot, h: 0.18, fontSize: 7.5, color: C.sub, align: 'center', margin: 0 });
  [0, 1, 2].forEach((j) => { // 감축 Lever 막대
    const i = j + 2, top = tops[i - 1], bot = tops[i];
    R(s, { x: bx(i), y: top, w: bw, h: bot - top, fill: { color: C.orange }, line: { color: C.orange, width: 0 } });
    T(s, '▼ OO', { x: bx(i) - 0.1, y: bot + 0.02, w: bw + 0.2, h: 0.2, fontSize: 8.5, bold: true, color: C.orangeD, align: 'center', margin: 0 });
  });
  stack(5, 0.28, 0.10);
  T(s, "'25比 OO%↓", { x: 0.4 + 5 * slot, y: 3.13, w: slot, h: 0.22, fontSize: 9, bold: true, color: C.orangeD, align: 'center', margin: 0 });
  for (let i = 0; i < 5; i++) { // 연결 점선
    seg(s, bx(i) + bw, tops[i], bx(i + 1), tops[i], { color: '999999', width: 0.75, dashType: 'dash' });
  }
  ["'25 Total TO", "S직군 직영전환\n('26.4월~)", '① 공정·프로세스\n자동화', '② 저효율 라인\n효율화', '③ 개인별\n업무 Span 확대', "'30 Total TO"].forEach((t, i) => {
    T(s, t, { x: 0.4 + i * slot, y: 3.78, w: slot, h: 0.3, fontSize: 8, bold: i === 0 || i === 5, color: i >= 2 && i <= 4 ? C.orangeD : C.text, align: 'center', margin: 0 });
  });

  [
    { t: '① 공정·프로세스 자동화', b: ['Intelligent Factory·AI/로봇 적용 확대', 'Inno TF 설비정예화 연계 운전·정비 TO↓', '위험·단순 반복작업 우선 자동화'] },
    { t: '② 저효율 제조라인 효율화(폐쇄)', b: ['저수익·저가동 라인 통폐합/폐쇄', '사업재편 연계 라인별 적정 TO 재산정', '잉여인력 신사업·해외법인 전환배치'] },
    { t: '③ 개인별 업무 Span 확대', b: ['직무통합·다기능화로 1인당 관리범위↑', '직영-협력, 협력사間 중복기능 통합', '사무·지원업무 AI Agent 활용(WX)'] },
  ].forEach((c, i) => {
    const x = 0.4 + i * 2.55;
    TB(s, c.t, { x, y: 4.12, w: 2.45, h: 0.25, fontSize: 9, bold: true, color: C.orangeD, fill: { color: C.orangeL }, line: { color: C.orange, width: 0.75 }, margin: 4 });
    TB(s, c.b.map((b, k) => run('· ' + b, { fontSize: 8, color: C.text, breakLine: k < c.b.length - 1 })),
      { x, y: 4.37, w: 2.45, h: 0.61, valign: 'middle', margin: 3, line: { color: C.line, width: 0.75 } });
  });

  // ---- 우 : TO-PO 운영 원칙 (개념도) ----
  const rx = 8.1, rw = 4.83;
  secHead(s, 'TO-PO 운영 원칙  (개념도)', rx, 2.42, rw, 0.3, 'left', 10);
  const ox = 8.3, oy = 3.88;
  seg(s, ox, 2.8, ox, oy, { color: C.line, width: 1 });
  seg(s, ox, oy, 12.9, oy, { color: C.line, width: 1 });
  const px = [0, 1, 2, 3, 4, 5].map((i) => 8.6 + i * 0.82);
  const TO = [2.98, 2.98, 3.18, 3.38, 3.55, 3.7];
  const PO = [2.95, 2.92, 3.0, 3.17, 3.4, 3.66];
  ["'25", "'26", "'27", "'28", "'29", "'30"].forEach((y, i) => T(s, y, { x: px[i] - 0.25, y: oy + 0.02, w: 0.5, h: 0.18, fontSize: 8, color: C.sub, align: 'center', margin: 0 }));
  for (let i = 0; i < 5; i++) {
    seg(s, px[i], TO[i], px[i + 1], TO[i + 1], { color: C.navy, width: 2.25 });
    seg(s, px[i], PO[i], px[i + 1], PO[i + 1], { color: C.orange, width: 2, dashType: 'dash' });
  }
  px.forEach((x, i) => {
    s.addShape(pptx.ShapeType.ellipse, { x: x - 0.04, y: TO[i] - 0.04, w: 0.08, h: 0.08, fill: { color: C.navy }, line: { color: C.white, width: 0.5 } });
  });
  // 범례
  seg(s, 11.2, 2.86, 11.5, 2.86, { color: C.navy, width: 2.25 });
  T(s, 'TO(기준인력) 선행 감축', { x: 11.55, y: 2.77, w: 1.4, h: 0.18, fontSize: 7.5, bold: true, color: C.navy, margin: 0 });
  seg(s, 11.2, 3.06, 11.5, 3.06, { color: C.orange, width: 2, dashType: 'dash' });
  T(s, 'PO(실제인력) 후행 수렴', { x: 11.55, y: 2.97, w: 1.4, h: 0.18, fontSize: 7.5, bold: true, color: C.orangeD, margin: 0 });
  s.addShape(pptx.ShapeType.upDownArrow, { x: 11.0, y: 3.19, w: 0.12, h: 0.18, fill: { color: C.mute }, line: { color: C.mute, width: 0 } });
  T(s, [
    run('과원(PO > TO) 발생 구간', { fontSize: 7.5, bold: true, color: C.sub, breakLine: true }),
    run('→ 자연감소·재배치로 단계적 해소', { fontSize: 7.5, color: C.sub }),
  ], { x: 8.45, y: 3.44, w: 2.4, h: 0.36, margin: 0 });

  TB(s, [
    run('· TO 선행 감축 → 정년퇴직 등 자연감소 범위 內 채용 조절로 PO 수렴', { fontSize: 8.5, color: C.text, breakLine: true }),
    run('· 잉여인력 Re-skilling 後 신사업·해외(인도 2기, 美 H社 합작) 재배치', { fontSize: 8.5, color: C.text, breakLine: true }),
    run('· 인위적 PO 조정 지양, 구조적 TO 관리로 노사 수용성 확보', { fontSize: 8.5, color: C.text }),
  ], { x: rx, y: 4.12, w: rw, h: 0.86, margin: 5, fill: { color: C.gray }, line: { color: C.line, width: 0.75 } });

  // ---- 실행기반 (Enabler) ----
  TB(s, '실행기반\n(Enabler)', { x: 0.4, y: 5.1, w: 1.0, h: 1.12, align: 'center', fontSize: 10, bold: true, color: C.white, fill: { color: C.navy }, line: { color: C.navy, width: 0 }, margin: 0 });
  [
    { t: '① 노무비 기반 인력관리체계 전환', tag: '신규', b: ['직영+협력 Total 노무비 Baseline·KPI 정립', 'TO 증감의 노무비 환산, 사업부별 목표 부여', '인력계획-예산-성과평가 연계 강화'] },
    { t: '② 노무 Risk 선제 대응', tag: "'25 전략2", b: ['통상임금·임금차액 소송 매몰비용 최소화', '노란봉투법 대비 협력/용역계약 유형별 대응', '정년연장·근로시간 규제 대비 제도 재정비'] },
    { t: '③ 인재 전환·재배치', tag: "'25 전략1", b: ['Domain&AI 전문인력 양성, 뉴칼라 2.0 개편', '잉여인력 Re-skilling 後 전략분야 재배치', 'S직군 조기 조직안정·직원 간 화합 도모'] },
    { t: '④ 일하는 방식·현장 리더십', tag: "'25 전략3", b: ['AI 기반 일하는 방식(Workway) 재정립·확산', '직책자 역할·권한 재정립, 솔선수범 리더십', '全 직원 DX 이해 증진 및 변화관리 강화'] },
  ].forEach((e, i) => {
    const x = 1.48 + i * 2.88, w = 2.79;
    TB(s, [
      run(e.t, { fontSize: 9, bold: true, color: C.navy }),
      run('  [' + e.tag + ']', { fontSize: 7.5, bold: true, color: e.tag === '신규' ? C.orangeD : C.mute }),
    ], { x, y: 5.1, w, h: 0.28, fill: { color: C.sky }, line: { color: C.line, width: 0.75 }, margin: 4 });
    TB(s, e.b.map((b, k) => run('· ' + b, { fontSize: 8, color: C.text, breakLine: k < e.b.length - 1 })),
      { x, y: 5.38, w, h: 0.84, margin: 4, line: { color: C.line, width: 0.75 } });
  });

  // ---- Roadmap ----
  TB(s, '추진\nRoadmap', { x: 0.4, y: 6.3, w: 1.0, h: 0.8, align: 'center', fontSize: 10, bold: true, color: C.white, fill: { color: C.navy }, line: { color: C.navy, width: 0 }, margin: 0 });
  [
    { t: "'26   기반 구축", c: C.blue, b: ['Total TO·노무비 Baseline 및 관리체계 정립', '조직혁신 TF 실행 로드맵 수립, S직군 조직안정'] },
    { t: "'27~'28   본격 실행", c: C.mid, b: ['자동화·라인 효율화 연계 TO 단계적 감축', '직영-협력 중복기능 통합, 협력사 재편'] },
    { t: "'29~'30   정착·고도화", c: C.navy, b: ["人+AI/로보틱스 Hybrid 인력구조 정착", "'30 Total TO OO%↓·인당생산성 OO%↑ 달성"] },
  ].forEach((p, i) => {
    const x = 1.48 + i * 3.76;
    TB(s, p.t, { shape: i === 0 ? pptx.ShapeType.homePlate : pptx.ShapeType.chevron, x, y: 6.3, w: 3.9, h: 0.32, align: 'center', fontSize: 9.5, bold: true, color: C.white, fill: { color: p.c }, line: { color: p.c, width: 0 }, margin: 0 });
    T(s, p.b.map((b, k) => run('· ' + b, { fontSize: 8, color: C.text, breakLine: k < p.b.length - 1 })),
      { x: x + 0.2, y: 6.64, w: 3.5, h: 0.44, valign: 'top', margin: 1 });
  });

  footer(s, '※ 추진체계: 조직혁신 TF(양사업부·인사·DX전략·기술연구원) 주관, 반기 단위 TO·노무비 Rolling 점검  /  OO 수치는 사업계획 확정 後 반영', 2);

  s.addNotes([
    "[발표 포인트] '30년 목표는 Total TO(직영+협력) OO% 감축, 인당 생산성 OO% 향상, Total 노무비 매출액 대비 OO% 이내 관리입니다.",
    'S직군 직영 전환은 총량이 변하지 않는 구성 변화이므로, 실질 감축은 자동화·라인 효율화·업무 Span 확대의 3대 Lever로 달성합니다.',
    'TO를 먼저 낮추고, PO는 정년퇴직 등 자연감소와 Re-skilling 후 신사업·해외 재배치를 통해 단계적으로 수렴시켜 노사 수용성을 확보합니다.',
    "이를 위해 노무비 기반 관리체계(신규)와 작년 수립한 3대 전략을 실행기반으로 재배치하고, '26 기반 구축 → '27~'28 본격 실행 → '29~'30 정착의 3단계로 추진합니다.",
  ].join('\n\n'));
}

if (!ONLY || ONLY === 1) slide1();
if (!ONLY || ONLY === 2) slide2();

pptx.writeFile({ fileName: OUT }).then((f) => console.log('saved:', f));
