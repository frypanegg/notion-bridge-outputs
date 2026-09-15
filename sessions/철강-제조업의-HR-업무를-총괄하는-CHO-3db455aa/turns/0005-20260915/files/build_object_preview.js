// Ⅰ장 '인력운영 패러다임 전환' 5번째 축(관리 주체) 추가안 — PPT 미수정, 미리보기용 SVG 생성
// 실행: node build_object_preview.js  → SVG 2개 생성 (PPT 삽입 > 그림으로 넣은 뒤 '도형으로 변환' 시 편집 가능)
// 좌표·색·폰트 크기는 4/build_slides.js slide1()과 동일 (단위: inch → SVG pt)
const fs = require('fs');
const path = require('path');

const FONT = "'Malgun Gothic', '맑은 고딕', 'Apple SD Gothic Neo', sans-serif";
const C = {
  navy: '0B2D5B', blue: '1F5FA8', mid: '16498A', sky: 'D6E4F5', sky2: 'EEF3FA',
  gray: 'F2F4F7', line: 'BFC7D2', text: '222222', sub: '555555', mute: '8A94A3',
  orange: 'E8541E', orangeD: 'C2410C', orangeL: 'FDEDE6', white: 'FFFFFF',
};
const PT = 72;
const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const L = (t, size, color, bold = false) => ({ t, size, color, bold });

class Canvas {
  constructor(x0, y0, w, h, scale = 4) {
    Object.assign(this, { x0, y0, w, h, scale, el: [] });
    this.rect(x0, y0, w, h, { fill: C.white, line: null });
  }
  X(x) { return ((x - this.x0) * PT).toFixed(2); }
  Y(y) { return ((y - this.y0) * PT).toFixed(2); }
  rect(x, y, w, h, { fill = C.white, line = C.line, lw = 0.75, dash = false, rx = 0 } = {}) {
    const stroke = line ? ` stroke="#${line}" stroke-width="${lw}"${dash ? ' stroke-dasharray="5 3"' : ''}` : '';
    this.el.push(`<rect x="${this.X(x)}" y="${this.Y(y)}" width="${(w * PT).toFixed(2)}" height="${(h * PT).toFixed(2)}" rx="${rx}" fill="${fill ? '#' + fill : 'none'}"${stroke}/>`);
  }
  rightArrow(x, y, w, h, color) { // PowerPoint rightArrow 기본형(몸통 50%, 머리 길이 = 짧은 변 × 0.5)
    const hl = Math.min(w, h) * 0.5;
    const pts = [[0, h * 0.25], [w - hl, h * 0.25], [w - hl, 0], [w, h / 2], [w - hl, h], [w - hl, h * 0.75], [0, h * 0.75]]
      .map(([px, py]) => `${this.X(x + px)},${this.Y(y + py)}`).join(' ');
    this.el.push(`<polygon points="${pts}" fill="#${color}"/>`);
  }
  text(x, y, w, h, lines, { align = 'center', valign = 'middle', margin = 2 } = {}) {
    const lh = 1.2, total = lines.reduce((a, l) => a + l.size * lh, 0);
    let top = Number(this.Y(y)) + (valign === 'top' ? margin : (h * PT - total) / 2);
    const anchor = { center: 'middle', left: 'start', right: 'end' }[align];
    const tx = align === 'center' ? Number(this.X(x + w / 2)) : align === 'left' ? Number(this.X(x)) + margin : Number(this.X(x + w)) - margin;
    lines.forEach((l) => {
      const base = top + (l.size * lh) / 2 + l.size * 0.36;
      this.el.push(`<text x="${tx.toFixed(2)}" y="${base.toFixed(2)}" font-size="${l.size}"${l.bold ? ' font-weight="bold"' : ''} fill="#${l.color}" text-anchor="${anchor}">${esc(l.t)}</text>`);
      top += l.size * lh;
    });
  }
  box(x, y, w, h, lines, o = {}) {
    this.rect(x, y, w, h, o);
    this.text(x, y, w, h, lines, o);
  }
  save(file) {
    const W = this.w * PT, H = this.h * PT;
    const svg = `<?xml version="1.0" encoding="UTF-8"?>\n<svg xmlns="http://www.w3.org/2000/svg" width="${(W * this.scale).toFixed(0)}" height="${(H * this.scale).toFixed(0)}" viewBox="0 0 ${W.toFixed(2)} ${H.toFixed(2)}" font-family="${FONT}">\n${this.el.join('\n')}\n</svg>\n`;
    fs.writeFileSync(path.join(__dirname, file), svg);
    console.log('saved:', file);
  }
}

// ---------- 패러다임 전환 5개 축 (⑤ 관리 주체 신규) ----------
const ROWS = [
  {
    no: '①', tag: '인력구성', title: 'S직군 직영 전환에 따른 인력구조 변화',
    body: ["근지위 소송·협력작업 혁신('26.4월~)으로 직영 인력 증가"], so: '직영/협력 구분의 단순 인원수(Volume) 관리 한계 有',
    key: ['관리', '대상'], asis: ['직영/협력', '구분 관리'], tobe: ['정규·S직군·계약·협력', '4개 축 Total 인력 통합 관리'],
  },
  {
    no: '②', tag: '비용구조', title: '통상임금 확대 등 노무비 상승 압력 가중',
    body: ['통상임금 범위 확대로 퇴직급여·복리후생 비용 증대'], so: '임금차액 소송에 따른 잠재적 매몰비용 증가',
    key: ['관리', '기준'], asis: ['인원수', '(Head Count)'], tobe: ['Total 노무비(Labor Cost)', '기반 원가 관리'],
  },
  {
    no: '③', tag: '생산주체', title: 'AI/로보틱스 등 생산·업무처리 주체 다변화',
    body: ['직영/협력 노동력을 대체하는 AI·로보틱스 및', '유지보수 인력 확대'], so: '사람 중심 인력계획으로는 실질 생산역량 파악 한계',
    key: ['관리', '범위'], asis: ['사람(직영/협력)', '중심 인력계획'], tobe: ['人+AI/로보틱스 포함', 'Total Workforce 설계'],
  },
  {
    no: '④', tag: '규모통제', title: '유연한 인력규모 통제 難',
    body: ['親노동 정책 기조下 채용규모 축소 부담,', '대의기구의 인력 충원 요구 확대'], so: '인위적 PO 조정 곤란 → 구조적 TO 관리 필요',
    key: ['관리', '방식'], asis: ['충원요구 대응형', 'PO 사후 통제'], tobe: ['기준인력(TO) 선제 설계', '및 상시 Rolling 통제'],
  },
  {
    no: '⑤', tag: '운영책임', title: '인사부서 주도 인력 감축의 실행력 한계', draft: true,
    body: ['현업은 충원 요구, 인사부서는 감축 통제하는 이원화 구조'], so: '권한·책임 불일치로 현업의 자율적 효율화 동기 부족',
    key: ['관리', '주체'], asis: ['인사부서 주도下', '인력 감축 노력'], tobe: ['인력운영 부서장 책임경영', '적정 노무비 內 자율 운영'], tobeSub: '재량·권한·책임 확대',
  },
];

// As-Is → To-Be 객체 (key / As-Is / 화살표 / To-Be) : slide1() 좌표 그대로
function paradigmRow(c, r, y, h) {
  c.box(7.95, y, 0.8, h, r.key.map((t) => L(t, 9.5, C.white, true)), { fill: C.mid, line: null, margin: 0 });
  c.box(8.8, y, 1.62, h, [L('As-Is', 7.5, C.mute, true)].concat(r.asis.map((t) => L(t, 9, '444444'))), { fill: C.gray });
  c.rightArrow(10.46, y + h / 2 - 0.14, 0.3, 0.28, C.orange);
  const tobe = [L('To-Be', 7.5, C.blue, true)].concat(r.tobe.map((t) => L(t, 9.5, C.navy, true)));
  if (r.tobeSub) tobe.push(L(r.tobeSub, 8.5, C.orangeD, true));
  c.box(10.8, y, 2.13, h, tobe, { fill: C.sky2, line: C.blue, lw: 1.25 });
}

// ---------- ① 신규 객체 단품 (현행 행 높이 1.0in) ----------
{
  const c = new Canvas(7.85, 0, 5.18, 1.2, 5);
  paradigmRow(c, ROWS[4], 0.1, 1.0);
  c.save('관리주체_AsIs_ToBe_객체.svg');
}

// ---------- ② 5개 축 배치 미리보기 (여건 변화 + 패러다임 전환, 행 높이 1.0 → 0.8in) ----------
{
  const c = new Canvas(2.85, 1.62, 10.23, 4.92, 3);
  const bx = 3.72, bw = 3.9, h = 0.8, pitch = 0.875;
  c.box(bx, 1.72, bw, 0.34, [L('여건 변화 (현상 및 문제점)', 11, C.white, true)], { fill: C.navy, line: null });
  c.box(7.95, 1.72, 4.98, 0.34, [L('인력운영 패러다임 전환 (As-Is → To-Be)', 11, C.white, true)], { fill: C.navy, line: null });

  ROWS.forEach((r, i) => {
    const y = 2.14 + i * pitch;
    c.rect(bx, y, bw, h);
    c.box(bx, y, 0.72, h, [L(r.no, 12, C.navy, true), L(r.tag, 8.5, C.navy, true)], { fill: C.sky, margin: 0 });
    c.text(4.5, y + 0.03, 3.08, 0.26, [L(r.title, 10, C.text, true)], { align: 'left', margin: 1 });
    if (r.draft) c.text(4.5, y + 0.03, 3.08, 0.26, [L('(참고안)', 7.5, C.mute, true)], { align: 'right', margin: 1 });
    c.text(4.5, y + 0.3, 3.08, h - 0.3, r.body.map((t) => L(t, 8, C.sub)).concat([L('→ ' + r.so, 8, C.orangeD, true)]), { align: 'left', valign: 'top', margin: 1 });
    c.rightArrow(7.64, y + h / 2 - 0.13, 0.28, 0.26, C.blue);
    paradigmRow(c, r, y, h);

    if (r.draft) { // 신규 행 강조 표시 (미리보기 전용, PPT 반영 대상 아님)
      c.rect(3.66, y - 0.035, 9.33, h + 0.07, { fill: null, line: C.orange, lw: 1.5, dash: true, rx: 3 });
      c.box(2.93, y + h / 2 - 0.22, 0.62, 0.44, [L('추가', 9, C.white, true), L('(안)', 8, C.white, true)], { fill: C.orange, line: null, margin: 0, rx: 3 });
    }
  });
  c.save('패러다임전환_5개축_배치_미리보기.svg');
}
