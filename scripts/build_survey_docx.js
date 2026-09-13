const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, Header, Footer, PageNumber,
  VerticalAlign, VerticalMergeType, convertInchesToTwip,
} = require('docx');
const fs = require('fs');

const FONT = 'Malgun Gothic';
const NAVY = '1F3A5F';
const NAVY_SOFT = 'E7EDF4';
const GREY = '595959';
const LINE = 'BFBFBF';
const GOLD = '9C6B23';

const pageW = 11906;
const marginLR = convertInchesToTwip(0.85);
const usableW = pageW - marginLR * 2;

function blank(width) { return '＿'.repeat(width || 10); }

function p(runsOrText, opts = {}) {
  const children = typeof runsOrText === 'string'
    ? [new TextRun({ text: runsOrText, font: FONT, size: opts.size || 20, color: opts.color || '262626', bold: !!opts.bold })]
    : runsOrText;
  return new Paragraph({
    spacing: { before: opts.before ?? 0, after: opts.after ?? 120, line: 300 },
    indent: opts.indent,
    border: opts.border,
    shading: opts.shading,
    children,
  });
}

function sectionHead(num, title, period) {
  return new Paragraph({
    spacing: { before: 420, after: 140 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 6 } },
    children: [
      new TextRun({ text: `${num}  `, bold: true, font: FONT, size: 30, color: NAVY }),
      new TextRun({ text: title, bold: true, font: FONT, size: 30, color: NAVY }),
      new TextRun({ text: `      (${period})`, italics: true, font: FONT, size: 18, color: GREY }),
    ],
  });
}

function qHead(num, text) {
  return p([
    new TextRun({ text: `${num}. `, bold: true, font: FONT, size: 21, color: NAVY }),
    new TextRun({ text, font: FONT, size: 21, color: '141414' }),
  ], { before: 220, after: 80 });
}

function subqHead(num, text) {
  return p([
    new TextRun({ text: `${num}. `, bold: true, font: FONT, size: 20, color: GREY }),
    new TextRun({ text, font: FONT, size: 20, color: '262626' }),
  ], {
    before: 140, after: 60,
    indent: { left: 260 },
    shading: { type: ShadingType.CLEAR, fill: 'F2F1EC' },
  });
}

function options(list, opts = {}) {
  if (opts.stacked) {
    return list.map(o => p(o, { indent: { left: 300 }, after: 40, size: 20 }));
  }
  const text = list.join('    ');
  return [p(text, { indent: { left: 300 }, after: 100, size: 20 })];
}

function fillLine(label, unit) {
  return p(`${label} : ${blank(12)} ${unit || ''}`.trim(), { indent: { left: 300 }, after: 100, size: 20 });
}

function help(text) {
  return p(`◆ ${text}`, { indent: { left: 200 }, after: 100, size: 18, color: GREY });
}

function footnote(text) {
  return p(text, { after: 40, size: 17, color: GREY });
}

function hr() {
  return new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 4 } }, spacing: { before: 60, after: 200 }, children: [] });
}

// ---------- table helpers ----------
function tc(text, opts = {}) {
  const children = Array.isArray(text) ? text : [new TextRun({
    text: String(text), font: FONT, size: opts.size || 17,
    bold: !!opts.header, color: opts.header ? 'FFFFFF' : '262626',
  })];
  return new TableCell({
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    columnSpan: opts.colSpan,
    rowSpan: opts.rowSpan,
    verticalMerge: opts.vMerge,
    shading: opts.header
      ? { type: ShadingType.CLEAR, fill: NAVY }
      : (opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill } : undefined),
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    children: [new Paragraph({ alignment: opts.align || AlignmentType.CENTER, children })],
  });
}
function blankCell(width, unit) {
  return tc(`${blank(6)}${unit ? ' ' + unit : ''}`, { width, size: 16 });
}
function grid(colWidths, rows) {
  const total = colWidths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: colWidths,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      left: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      right: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: LINE },
      insideVertical: { style: BorderStyle.SINGLE, size: 4, color: LINE },
    },
    rows: rows.map(cells => new TableRow({ children: cells })),
  });
}

const W = usableW;
function frac(...parts) {
  const sum = parts.reduce((a, b) => a + b, 0);
  return parts.map(x => Math.round(W * x / sum));
}

// ============================================================
// Build all repeated matrix tables
// ============================================================
function tblQ3() {
  const [c0, ...rest] = frac(2.2, 1, 1, 1, 1, 1, 1);
  const cw = [c0, ...rest];
  const rows = [];
  rows.push([
    tc('구분', { header: true, width: cw[0], rowSpan: 2 }),
    tc('등록회원 수', { header: true, width: cw[1] + cw[2], colSpan: 2 }),
    tc('이용자 수', { header: true, width: cw[3] + cw[4], colSpan: 2 }),
    tc('대출자 수', { header: true, width: cw[5] + cw[6], colSpan: 2 }),
  ]);
  rows.push([
    tc('', { header: true, width: 0, vMerge: VerticalMergeType.CONTINUE }),
    tc('2025년', { header: true, width: cw[1] }), tc('누계', { header: true, width: cw[2] }),
    tc('2025년', { header: true, width: cw[3] }), tc('누계', { header: true, width: cw[4] }),
    tc('2025년', { header: true, width: cw[5] }), tc('누계', { header: true, width: cw[6] }),
  ]);
  ['시각장애', '청각장애', '지체·뇌병변장애', '발달장애', '기타 유형 장애', '계'].forEach(label => {
    rows.push([tc(label, { width: cw[0], align: AlignmentType.LEFT }),
      ...Array(6).fill(0).map((_, i) => blankCell(cw[i + 1], '명'))]);
  });
  return grid(cw, rows);
}

function tblAmount(rowLabels) {
  const [c0, c1] = frac(2.5, 1.5);
  const rows = [[tc('구분', { header: true, width: c0 }), tc('금액(원)', { header: true, width: c1 })]];
  rowLabels.forEach(l => rows.push([tc(l, { width: c0, align: AlignmentType.LEFT }), blankCell(c1, '원')]));
  return grid([c0, c1], rows);
}

function tblQ4() {
  const [c0, c1, c2] = frac(1.2, 1.5, 1.5);
  const rows = [[tc('구분', { header: true, width: c0 }), tc('', { header: true, width: c1 }), tc('금액(원)', { header: true, width: c2 })]];
  const groups = [
    ['자체예산', ['자체 재원', '후원금', '기타(회비 등)']],
    ['정부보조금', ['문화체육관광부', '보건복지부', '기타 정부부처(교육부 등)', '지자체']],
  ];
  groups.forEach(([g, items]) => {
    items.forEach((it, i) => {
      const row = [];
      if (i === 0) row.push(tc(g, { width: c0, rowSpan: items.length }));
      else row.push(tc('', { width: 0, vMerge: VerticalMergeType.CONTINUE }));
      row.push(tc(it, { width: c1, align: AlignmentType.LEFT }));
      row.push(blankCell(c2, '원'));
      rows.push(row);
    });
  });
  rows.push([tc('계', { width: c0 + c1, colSpan: 2, align: AlignmentType.LEFT, header: false, fill: 'F2F1EC' }), blankCell(c2, '원')]);
  return grid([c0, c1, c2], rows);
}

function tblQ62() {
  const [c0, c1] = frac(1.3, 2.7);
  const rows = [[tc('구분', { header: true, width: c0 }), tc('내용', { header: true, width: c1 })]];
  const basic = [['자료실(코너) 총 면적', '㎡'], ['설치 위치', '층'], ['총 이용좌석 수', '석']];
  basic.forEach((r, i) => {
    const row = [];
    if (i === 0) row.push(tc('면적 및 위치, 총 좌석 수', { width: c0, rowSpan: basic.length }));
    else row.push(tc('', { width: 0, vMerge: VerticalMergeType.CONTINUE }));
    row.push(tc(`${r[0]}  :  ${blank(6)} ${r[1]}`, { width: c1, align: AlignmentType.LEFT, size: 16 }));
    rows.push(row);
  });
  const spaces = ['안내데스크', '대면낭독실', '영상실(멀티미디어실)', '정보검색실(정보검색대)', '열람실(열람석)', '대체자료 제작실', '기타(직접기입: ________ )'];
  spaces.forEach((s, i) => {
    const row = [];
    if (i === 0) row.push(tc('세부공간 및 설비 등 구성 여부', { width: c0, rowSpan: spaces.length }));
    else row.push(tc('', { width: 0, vMerge: VerticalMergeType.CONTINUE }));
    row.push(tc(`${s}   ①있다   ②없다`, { width: c1, align: AlignmentType.LEFT, size: 16 }));
    rows.push(row);
  });
  return grid([c0, c1], rows);
}

function tblQ7() {
  const [c0, c1, c2, c3] = frac(1, 2, 1.4, 1.1);
  const rows = [[tc('구분', { header: true, width: c0 }), tc('항목', { header: true, width: c1 }),
    tc('설치 기준', { header: true, width: c2 }), tc('설치유무', { header: true, width: c3 })]];
  const data = [
    ['매개시설', [['주출입구 접근로', '의무'], ['장애인 전용주차구역', '의무'], ['주출입구 높이 차이 제거', '의무']]],
    ['내부시설', [['출입구(문)', '의무'], ['복도', '의무'], ['계단 또는 승강기', '의무']]],
    ['위생시설', [['대변기', '의무'], ['소변기', '권장'], ['세면대', '권장']]],
    ['안내시설', [['점자블록', '의무'], ['유도 및 안내설비', '권장'], ['경보 및 피난설비', '의무']]],
    ['기타시설', [['접수대/작업대', '의무']]],
  ];
  data.forEach(([g, items]) => {
    items.forEach((it, i) => {
      const row = [];
      if (i === 0) row.push(tc(g, { width: c0, rowSpan: items.length }));
      else row.push(tc('', { width: 0, vMerge: VerticalMergeType.CONTINUE }));
      row.push(tc(it[0], { width: c1, align: AlignmentType.LEFT }));
      row.push(tc(it[1], { width: c2 }));
      row.push(tc('①있다   ②없다', { width: c3, size: 16 }));
      rows.push(row);
    });
  });
  return grid([c0, c1, c2, c3], rows);
}

function tblQ8() {
  const [c0, c1, c2] = frac(1.1, 2.2, 0.9);
  const rows = [[tc('구분', { header: true, width: c0 }), tc('품목', { header: true, width: c1 }), tc('수량', { header: true, width: c2 })]];
  const groups = [
    ['시각장애', ['확대경', '휴대용 독서확대기', '이동용 독서확대기', '탁상용 독서확대기', '데이지 플레이어', '음성바코드 리더기', '점자정보 단말기', '화면낭독 S/W', '화면확대 S/W', '점역 S/W', '문서인식 S/W 및 H/W', '점자프린터', '점자라벨 프린터']],
    ['청각/언어장애', ['소리증폭장치(보청기)', '신호,경보장치', '강의,강연 청취기', '화상전화기', '골도전화기(헤드셋)', '보완대체 의사소통기기(ACC)']],
    ['지체/지적/뇌병변장애', ['특수마우스', '특수키보드', '자세보조장치', '휠체어용 작업테이블', '모니터 이동보조기', '원고 홀더', '터치스크린', '경사각 작업테이블', '필기구 홀더', '높낮이 조절책상', '컴퓨터 입력 S/W', '스위치', '입력보조장치']],
  ];
  groups.forEach(([g, items]) => {
    items.forEach((it, i) => {
      const row = [];
      if (i === 0) row.push(tc(g, { width: c0, rowSpan: items.length }));
      else row.push(tc('', { width: 0, vMerge: VerticalMergeType.CONTINUE }));
      row.push(tc(it, { width: c1, align: AlignmentType.LEFT }));
      row.push(blankCell(c2, '개'));
      rows.push(row);
    });
  });
  rows.push([tc('기타', { width: c0 }), tc(`직접기입 : ${blank(10)}`, { width: c1, align: AlignmentType.LEFT }), blankCell(c2, '개')]);
  return grid([c0, c1, c2], rows);
}

function tblQ11_2() {
  const [c0, c1, c2, c3, c4, c5, c6] = frac(1.3, 1, 1, 1, 1, 1, 1);
  const rows = [
    [tc('구분', { header: true, width: c0, rowSpan: 2 }),
      tc('정규직(A)', { header: true, width: c1 + c2, colSpan: 2 }),
      tc('비정규직(B)', { header: true, width: c3 + c4, colSpan: 2 }),
      tc('계(A+B)', { header: true, width: c5 + c6, colSpan: 2 })],
    [tc('', { width: 0, vMerge: VerticalMergeType.CONTINUE }),
      tc('남', { header: true, width: c1 }), tc('여', { header: true, width: c2 }),
      tc('남', { header: true, width: c3 }), tc('여', { header: true, width: c4 }),
      tc('남', { header: true, width: c5 }), tc('여', { header: true, width: c6 })],
    [tc('인원 수(명)', { width: c0, align: AlignmentType.LEFT }),
      ...[c1, c2, c3, c4, c5, c6].map(w => blankCell(w, '명'))],
  ];
  return grid([c0, c1, c2, c3, c4, c5, c6], rows);
}

function tblQ11_3() {
  const cols = ['사서1급', '사서2급', '준사서', '사복1급', '사복2급', '점역1급', '점역2급', '점역3급', '독서지도사', '수어통역사', '농통역사', '보조공학사', '특수교사', '기타'];
  const widths = cols.map(() => Math.round(W / cols.length));
  const header = [
    tc('사서', { header: true, width: widths[0] + widths[1] + widths[2], colSpan: 3 }),
    tc('사회복지사', { header: true, width: widths[3] + widths[4], colSpan: 2 }),
    tc('점역·교정사', { header: true, width: widths[5] + widths[6] + widths[7], colSpan: 3 }),
    tc('독서\n지도사', { header: true, width: widths[8] }),
    tc('수어\n통역사', { header: true, width: widths[9] }),
    tc('농\n통역사', { header: true, width: widths[10] }),
    tc('보조\n공학사', { header: true, width: widths[11] }),
    tc('특수\n교사', { header: true, width: widths[12] }),
    tc('기타', { header: true, width: widths[13] }),
  ];
  const sub = ['1급', '2급', '준사서', '1급', '2급', '1급', '2급', '3급', '', '', '', '', '', ''].map((t, i) =>
    tc(t, { header: true, width: widths[i] }));
  const data = widths.map(w => blankCell(w, '명'));
  return grid(widths, [header, sub, [tc('인원 수(명)', { width: 0, colSpan: 14, align: AlignmentType.LEFT, fill: 'F2F1EC' })], data]);
}

function tblQ16() {
  const cols = [1.6, 1.1, 0.75, 0.75, 0.75, 0.75, 0.75, 0.75, 0.75, 0.75, 0.9];
  const w = frac(...cols);
  const header1 = [
    tc('구분', { header: true, width: w[0] + w[1], colSpan: 2, rowSpan: 3 }),
    tc('1. 소장현황', { header: true, width: w.slice(2, 10).reduce((a, b) => a + b, 0), colSpan: 8 }),
    tc('2.이용현황\n대출현황(권/점)', { header: true, width: w[10], rowSpan: 3 }),
  ];
  const header2 = [
    tc('', { width: 0, vMerge: VerticalMergeType.CONTINUE }), tc('', { width: 0, vMerge: VerticalMergeType.CONTINUE }),
    tc('2025년', { header: true, width: w[2] + w[3] + w[4] + w[5], colSpan: 4 }),
    tc('소장 자료', { header: true, width: w[6] + w[7] + w[8] + w[9], colSpan: 4 }),
    tc('', { width: 0, vMerge: VerticalMergeType.CONTINUE }),
  ];
  const header3 = [
    tc('', { width: 0, vMerge: VerticalMergeType.CONTINUE }), tc('', { width: 0, vMerge: VerticalMergeType.CONTINUE }),
    ...['구입', '기증', '자체제작', '계', '구입', '기증', '자체제작', '계'].map((t, i) => tc(t, { header: true, width: w[2 + i] })),
    tc('', { width: 0, vMerge: VerticalMergeType.CONTINUE }),
  ];
  const rows = [header1, header2, header3];
  const groups = [
    ['시각자료', [['점자', ['점자도서', '묵점자 혼용도서', '점자라벨도서', '전자점자파일']],
      ['영상', ['화면해설영상자료']],
      ['음성', ['CD-ROM, DVD 녹음자료', 'MP3 녹음파일', '데이지(DAISY) 자료']]]],
    ['청각자료', [['', ['한국수어&자막영상자료', '한국수어영상자료', '자막영상자료']]]],
    ['기타자료', [['', ['촉각도서', '큰글자도서', '읽기쉬운도서']]]],
  ];
  groups.forEach(([macro, subs]) => {
    const macroRows = subs.reduce((s, x) => s + x[1].length, 0);
    let macroPrinted = false;
    subs.forEach(([sub, items]) => {
      items.forEach((it, i) => {
        const row = [];
        if (!macroPrinted) { row.push(tc(macro, { width: w[0], rowSpan: macroRows })); macroPrinted = true; }
        else row.push(tc('', { width: 0, vMerge: VerticalMergeType.CONTINUE }));
        if (i === 0) row.push(tc(sub, { width: w[1], rowSpan: items.length }));
        else row.push(tc('', { width: 0, vMerge: VerticalMergeType.CONTINUE }));
        row.push(tc(it, { width: w[1], align: AlignmentType.LEFT, size: 15 }));
        for (let c = 0; c < 8; c++) row.push(blankCell(w[2 + c], c % 2 === 0 ? '종' : '권'));
        row.push(blankCell(w[10], '권'));
        rows.push(row);
      });
    });
  });
  const trFree = [tc('기타자료', { width: w[0] }), tc(`기입: ${blank(4)}`, { width: w[1], align: AlignmentType.LEFT, size: 15 })];
  for (let c = 0; c < 8; c++) trFree.push(blankCell(w[2 + c], c % 2 === 0 ? '종' : '권'));
  trFree.push(blankCell(w[10], '권'));
  rows.push(trFree);
  const trTotal = [tc('계', { width: w[0] + w[1], colSpan: 2, fill: 'F2F1EC' })];
  for (let c = 0; c < 8; c++) trTotal.push(blankCell(w[2 + c], c % 2 === 0 ? '종' : '권'));
  trTotal.push(blankCell(w[10], '권'));
  rows.push(trTotal);
  return grid(w, rows);
}

function tblProvide(rowsDef, extraCols) {
  // rowsDef: [ [label, footnote?], ... ]  extraCols: array of {title, unit}
  const [c0, c1, ...rest] = frac(1.8, 1.1, ...extraCols.map(() => 1));
  const cw = [c0, c1, ...rest];
  const header = [tc('구분', { header: true, width: c0 }), tc('서비스 제공 여부', { header: true, width: c1 })];
  extraCols.forEach((c, i) => header.push(tc(c.title, { header: true, width: cw[2 + i] })));
  const rows = [header];
  rowsDef.forEach(([label]) => {
    const row = [tc(label, { width: c0, align: AlignmentType.LEFT, size: 16 }),
      tc('①제공함  ②제공안함', { width: c1, size: 15 })];
    extraCols.forEach((c, i) => row.push(blankCell(cw[2 + i], c.unit)));
    rows.push(row);
  });
  return grid(cw, rows);
}

function tblQ19_1() {
  const groups = ['도서관 이용교육', '정보 활용교육', '문화 프로그램', '독서 관련 프로그램', '기타'];
  const w0 = Math.round(W * 0.16);
  const wc = Math.round((W - w0) / 10);
  const widths = [w0, ...Array(10).fill(wc)];
  const header1 = [tc('장애 유형', { header: true, width: w0, rowSpan: 2 }),
    ...groups.map(g => tc(g, { header: true, width: wc * 2, colSpan: 2 }))];
  const header2 = [tc('', { width: 0, vMerge: VerticalMergeType.CONTINUE })];
  groups.forEach(() => { header2.push(tc('회', { header: true, width: wc })); header2.push(tc('명', { header: true, width: wc })); });
  const rows = [header1, header2];
  ['시각장애', '청각장애', '지체·뇌병변장애', '발달장애', '기타 유형 장애', '계'].forEach(label => {
    rows.push([tc(label, { width: w0, align: AlignmentType.LEFT }), ...Array(10).fill(0).map((_, i) => blankCell(wc, i % 2 === 0 ? '회' : '명'))]);
  });
  return grid(widths, rows);
}

// ============================================================
// Document assembly
// ============================================================
const body = [];

body.push(new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: '설문지 초안 · Review Draft', font: FONT, size: 16, color: GOLD, bold: true })] }));
body.push(new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: '대학도서관 대체자료 소장 현황 및 서비스 실태조사 연구', bold: true, font: FONT, size: 38, color: NAVY })] }));
body.push(new Paragraph({
  spacing: { after: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 16, color: NAVY, space: 8 } },
  children: [new TextRun({ text: '대학도서관 · 장애학생지원센터용', italics: true, font: FONT, size: 22, color: NAVY })],
}));
body.push(p('설문지는 국립장애인도서관이 2년 주기로 실시하는 「전국 도서관 장애인서비스 현황조사」와 교육부에서 3년 주기로 실시하는 「장애대학생 교육복지지원 실태조사」 문항 체계를 바탕으로 작성한 것입니다. 조사 대상인 대학도서관(대학 내 장애학생지원센터 포함) 상황에 맞게 재구성하였으며, 문항의 내용과 구성은 원 설문과 최대한 동일하게 유지하고자 하였으며, 기관 유형을 나타내는 표현(예: 공공도서관 → 대학도서관/장애학생지원센터) 등을 조정하였습니다.', { after: 200 }));

body.push(p('조사 개요', { bold: true, size: 22, after: 80 }));
body.push(grid(frac(1, 3), [
  [tc('조사 기간', { header: true }), tc('[조사기간 입력]', { align: AlignmentType.LEFT })],
  [tc('조사 대상', { header: true }), tc('전국 대학도서관 담당자, 대학 내 장애학생지원센터 담당자', { align: AlignmentType.LEFT })],
  [tc('주최 기관', { header: true }), tc('[주최기관 입력]', { align: AlignmentType.LEFT })],
  [tc('주관/조사 기관', { header: true }), tc('[주관·조사기관 입력]', { align: AlignmentType.LEFT })],
  [tc('문의처', { header: true }), tc('[문의처 입력]', { align: AlignmentType.LEFT })],
]));
body.push(p('', { after: 100 }));
[
  '설문지 작성 기준은 2026.09.30. ~ 2026.10.16. 입니다.',
  '문항의 입력 내용 중 해당 내역이 없는 경우 “0”을 반드시 표시해 주십시오.',
  '문항 중 ‘누계’ 기준은 2025.12.31. 기준 누계를 작성해 주십시오.',
  '대체자료에 관한 상세 설명은 별첨 설문응답 지침서 내 [보기카드]를 참고해 주시기 바랍니다.',
  '이하 문항의 ‘귀 기관’은 응답하시는 대학도서관 또는 장애학생지원센터를 의미합니다.',
].forEach(t => body.push(p(`•  ${t}`, { size: 18, color: GREY, after: 60 })));

body.push(hr());
body.push(p('◆ 응답자 인적 사항 ◆', { bold: true, size: 24, after: 100, color: NAVY }));
body.push(p('대학도서관에 장애인서비스 업무를 전담하는 담당자가 없는 경우, 대학 내 장애학생지원센터 담당자께서 응답해 주시기 바랍니다.', { size: 18, color: NAVY, after: 140 }));
body.push(p(`담당 기관 구분 :   ① 대학도서관 담당자     ② 장애학생지원센터 담당자`, { after: 100 }));
body.push(p(`대학명 : ${blank(14)}          소속 부서명 : ${blank(14)}`, { after: 100 }));
body.push(p(`성명 : ${blank(10)}          직책 : ${blank(10)}`, { after: 100 }));
body.push(p(`전화번호 : ( ${blank(4)} ) - ( ${blank(4)} ) - ( ${blank(4)} )`, { after: 100 }));
body.push(p('※ 응답자 인적 사항은 본 조사에 실제 참여해 주셨는지 확인하기 위한 자료로만 활용되며, 조사 완료 후 폐기됩니다.', { size: 16, color: GREY }));

// ---- SECTION 1 ----
body.push(sectionHead('①', '기관 운영 및 예산 현황', '2025년 1월 1일 ~ 12월 31일 기준'));
body.push(p(`고유번호(도서관부호 기재) : ${blank(12)}`, { after: 100 }));
body.push(p('홈페이지 웹접근성 지침 준수 여부 :   ① 준수     ② 향후 적용 예정', { after: 60 }));
body.push(help('홈페이지 웹 접근성 지침 : 「지능정보화기본법」에 따라 장애인이나 고령자분들이 웹사이트에서 제공하는 정보를 비장애인과 동등하게 접근·이용할 수 있도록 보장하는 것으로, 한국형 웹콘텐츠 접근성 지침(KWCAG) 혹은 국제표준 웹 접근성 지침(WCAG) 준수에 대한 사항을 의미'));

body.push(qHead('문1', '귀 기관의 운영 계획에는 장애대학(원)생에 대한 학습지원(대체자료 제작 및 서비스) 내용이 포함되어 있습니까?'));
body.push(...options(['① 있다', '② 없다']));

body.push(qHead('문2', '귀 기관에서는 장애대학(원)생에 대한 학습지원(대체자료 제작 및 서비스) 매뉴얼이나 지침을 마련하고 있습니까?'));
body.push(...options(['① 있다', '② 없다']));

body.push(qHead('문3', '귀 기관의 장애 유형별 등록회원 수, 이용자 수, 대출자 수를 기재해 주십시오. 해당 인원이 없는 경우 ‘0’을 기재해 주십시오.'));
body.push(tblQ3());
body.push(footnote('1) 등록회원 수 : 2025년 란에는 2025.1.1.~12.31. 신규 등록한 회원 수를, 누계란에는 개관년도 이후 2025.12.31.까지의 누적 회원 수를 기입(중간에 기관명이 바뀐 경우 통계 집계 시작 이후부터의 누계)'));
body.push(footnote('2) 이용자 수 : 2025년 란에는 2025.1.1.~12.31. 방문한 이용자 수를, 누계란에는 누적 이용자 수를 기입(출입구 Counter 수로 산정하되, 정확한 산출이 어려운 경우 1주일을 정해 계산 후 x52주로 확대 측정)'));
body.push(footnote('3) 대출자 수 : 2025년 란에는 실제 대출자 수를, 누계란에는 누적 대출자 수를 기입(1년간 1회 이상 대출기록을 가진 이용자 기준)'));

body.push(qHead('문4', '귀 기관의 2025년도 장애대학(원)생에 대한 학습지원(대체자료 제작 및 서비스) 관련 예산을 기재해 주십시오. 해당 내역이 없는 경우 ‘0’을 기재해 주십시오.'));
body.push(tblQ4());
body.push(footnote('4) 자체예산 : 귀 기관의 장애인서비스를 위한 기본 재정을 의미'));
body.push(footnote('5) 정부보조금 : 귀 기관이 장애인서비스를 위해 정부로부터 지원받은 보조금을 의미'));

body.push(qHead('문5', '귀 기관이 2025년도에 장애인서비스 운영을 위해 지출한 예산을 기재해 주십시오. 해당 내역이 없는 경우 ‘0’을 기재해 주십시오. (최종 결산액 기준으로, 앞의 예산 내역과 일치하지 않아도 무방함)'));
body.push(tblAmount(['인건비', '자료구입', '자료제작', '프로그램 운영비', '보조기기 구입/유지', '편의시설 보완/확충', '기타', '계']));

// ---- SECTION 2 ----
body.push(sectionHead('②', '시설 및 설비 환경', '2025년 12월 31일 기준'));

body.push(qHead('문6', '귀 기관에는 장애대학(원)생을 위한 자료실이나 전용 공간(코너)을 운영하고 있습니까?'));
body.push(...options(['① 장애인자료실을 운영하고 있다', '② 장애인 전용 공간(코너)을 운영하고 있다', '③ 없다']));
body.push(subqHead('문6-1', '귀 기관 장애인자료실(코너)의 명칭, 개실년도, 담당부서를 작성해 주십시오.'));
body.push(fillLine('자료실(코너)명', ''));
body.push(fillLine('개실년도', '년'));
body.push(fillLine('담당부서', ''));
body.push(subqHead('문6-2', '귀 기관 장애인자료실(코너)의 면적, 이용좌석 수, 세부공간별 현황을 기재해 주십시오.'));
body.push(help('참고 : 장애인자료실이 독립 공간으로 있는 경우에는 해당하는 자료실의 내용으로 기술하고, 일반적인 열람공간 내에 장애인 공간(코너)이 있는 경우에는 장애인 이용자가 이용 가능한 공간의 전체 면적의 추정치를 기재하고, 장애인들이 이용할 수 있는 좌석 수 등을 기준으로 정리함'));
body.push(tblQ62());
body.push(footnote('6) 대면낭독실 : 시각장애인을 위해 대면낭독이 가능한 별도의 공간을 의미'));
body.push(footnote('7) 대체자료 제작실 : 점자자료 및 녹음자료 등을 제작할 수 있는 별도의 공간'));

body.push(qHead('문7', '귀 기관의 장애인 편의시설 설치 유무를 기재해 주십시오. (설치기준은 「장애인·노인·임산부 등의 편의증진 보장에 관한 법률 시행령」 별표에 제시된 내용입니다.)'));
body.push(help('장애인 편의시설에 대한 사항은 설문응답 지침서 내 [보기카드 1]의 내용을 참고해 주시기 바랍니다. (추후 링크 연결 예정)'));
body.push(tblQ7());

body.push(qHead('문8', '귀 기관의 장애인 독서보조기기 보유 현황을 기재해 주십시오. 해당사항이 없는 경우에는 0으로 작성해 주시면 됩니다.'));
body.push(help('장애인 독서보조기기에 대한 사항은 설문응답 지침서 내 [보기카드 2]의 내용을 참고해 주시기 바랍니다. (추후 링크 연결 예정)'));
body.push(tblQ8());

body.push(qHead('문9', '귀 기관에서는 독서보조기기 지원서비스를 어떻게 제공하고 있는지, 해당하는 항목을 모두 체크해 주세요.'));
body.push(...options(['① 관내 기기 대여 및 이용', '② 관외 기기 대여', '③ 기기 사용법 안내 서비스', '④ 장애인복지관 등 관련기관에서 임대']));
body.push(fillLine('⑤ 기타', ''));

body.push(qHead('문10', '귀 기관에서는 직원(자원봉사자 포함)을 대상으로 ‘독서보조기기 사용 안내’에 대한 교육을 실시한 경험이 있습니까?'));
body.push(...options(['① 있다', '② 없다']));

// ---- SECTION 3 ----
body.push(sectionHead('③', '인력 및 교육', '2025년 1월 1일 ~ 12월 31일 기준'));

body.push(qHead('문11', '귀 기관에는 장애대학(원)생을 위한 대체자료 제작 및 서비스를 전담하는 직원이 있습니까?'));
body.push(...options(['① 있다', '② 없다']));
body.push(subqHead('문11-1', '장애대학(원)생을 위한 대체자료 제작 및 서비스 전담 직원은 없지만, 관련 업무를 담당하는 직원이 있습니까?'));
body.push(...options(['① 있다', '② 없다']));
body.push(subqHead('문11-2', '귀 기관의 장애대학(원)생을 위한 대체자료 제작 및 서비스 전담 혹은 담당 직원 현황을 기재해 주십시오. (해당 인원이 없는 경우 ‘0’을 기재)'));
body.push(tblQ11_2());
body.push(subqHead('문11-3', '귀 기관의 장애대학(원)생을 위한 대체자료 제작 및 서비스 전담 혹은 담당 직원의 자격증 총 보유현황을 기재해 주십시오. (여러 자격증을 동시 보유하는 경우 해당 항목을 모두 집계, 해당 인원이 없는 경우 ‘0’을 기재)'));
body.push(tblQ11_3());
body.push(footnote('8) 농통역사 : 청각장애인으로 비장애인과 청각장애인과의 의사소통 담당'));
body.push(footnote('9) 보조공학사 : 장애인의 원활한 생활을 지원하기 위한 보조공학 기기와 관련 서비스 담당'));

body.push(qHead('문12', '귀 기관에는 장애대학(원)생을 위한 대체자료 제작 및 서비스를 위한 자원봉사자(사회적 일자리 포함)가 있습니까?'));
body.push(...options(['① 있다', '② 없다']));
body.push(subqHead('문12-1', '귀 기관에서 활동하는 장애인 이용자를 위한 자원봉사자(사회적 일자리 포함)는 몇 명이며, 2025년 기준 자원봉사활동 횟수는 몇 회입니까?'));
body.push(fillLine('자원봉사자', '명'));
body.push(fillLine('2025년 자원봉사활동 누적 횟수', '회'));
body.push(help('자원봉사활동 횟수는 장애인서비스에 관한 내용으로 한정하며, 1일 1회로 집계'));
body.push(subqHead('문12-2', '귀 기관에서 장애인 이용자를 위하여 활동하는 자원봉사자(사회적 일자리 포함)는 주로 어떤 서비스 분야에서 활동하고 계십니까?'));
body.push(...options([
  '① 자료대출 서비스(방문대출, 우편대출, 택배서비스 등)',
  '② 자료제작 서비스(점역, 교정, 녹음 등)',
  '③ 정보이용 서비스(대면낭독, 정보검색 지원, 보조기기 이용지원 등)',
  '④ 의사소통 서비스(대필, 통역, 의사소통 지원 등)',
  '⑤ 독서·문화 프로그램 운영지원',
], { stacked: true }));
body.push(fillLine('⑥ 기타', ''));

body.push(qHead('문13', '귀 기관에서는 2025년에 국립중앙도서관 및 국립장애인도서관에서 진행한 장애인서비스 관련 교육(사서교육훈련, 워크숍, 세미나 등)에 참여한 경험이 있으십니까?'));
body.push(...options(['① 있다', '② 없다']));
body.push(subqHead('문13-1', '귀 기관에서 2025년에 국립중앙도서관 및 국립장애인도서관에서 진행한 장애인서비스 관련 교육(사서교육훈련, 워크숍, 세미나 등)에 참여한 횟수와 참가직원수를 기재해 주십시오.'));
body.push(fillLine('총 횟수', '회'));
body.push(fillLine('총 참가직원수', '명'));

body.push(qHead('문14', '귀 기관에서는 장애대학(원)생을 위한 대체자료 제작 및 서비스를 효율적으로 제공하기 위해 2025년에 직원(자원봉사자 포함)을 대상으로 자체적으로 교육을 실시한 경험이 있습니까?'));
body.push(...options(['① 있다', '② 없다']));
body.push(subqHead('문14-1', '그렇다면, 상기 교육을 실시한 횟수와 참가자 수를 기재해 주십시오.'));
body.push(fillLine('총 횟수', '회'));
body.push(fillLine('총 참가자수', '명'));

// ---- SECTION 4 ----
body.push(sectionHead('④', '장서(자료) 구성', '2025년 1월 1일 ~ 12월 31일 기준'));
body.push(p('※ 장애학생지원센터에서 별도의 장서(자료)를 소장·관리하지 않는 경우, 아래 문항은 해당사항 없음(0)으로 기재해 주시기 바랍니다.', { size: 18, color: GREY, after: 140 }));

body.push(qHead('문15', '귀 기관에서는 장서개발지침이나 자료수집 기준(지침)에 장애인 자료에 대한 내용을 포함하고 있습니까?'));
body.push(...options(['① 있다', '② 없다']));

body.push(qHead('문16', '귀 기관에서 2025년 한 해 동안 수집한 장애 유형별 대체자료 수와 장애인 대체자료별 총 소장 현황 및 이용현황을 기재해 주십시오. 표의 분류에 따라 종(권) 수를 기재하고 해당 자료가 없는 경우 ‘0’을 기재해 주십시오.'));
body.push(help('구입, 기증, 제작 모두 2025년란에는 2025.1.1.~12.31.의 대체자료 현황을 기입하고, 총 소장자료 란에는 2025.12.31. 기준 누적 대체자료 수를 기입(제작의 경우 외부기관 위탁사업에 의한 제작은 제외)'));
body.push(help('‘종’은 자료유형별 종류 수, ‘권(점)’은 실제 물리적 개수를 의미. 예) 점자도서 1종을 2권, 3권씩 복본 보유 시 전체 집계는 2종 5권'));
body.push(help('장애인 대체자료에 대한 사항은 설문응답 지침서 내 [보기카드 3]의 내용을 참고해 주시기 바랍니다.'));
body.push(tblQ16());

// ---- SECTION 5 ----
body.push(sectionHead('⑤', '관련 서비스 제공', '2025년 1월 1일 ~ 12월 31일 기준'));

body.push(qHead('문17', '귀 기관의 장애대학(원)생 대상 자료대출서비스 현황을 기재해 주십시오. 해당사항이 없는 경우에는 0으로 작성해 주시면 됩니다.'));
body.push(tblProvide([
  ['관외대출'], ['책나래서비스'], ['우편·택배서비스(책나래서비스 외)'], ['상호대차서비스'], ['방문대출'], ['이동도서관'], ['기타(기입: ________ )'],
], [{ title: '이용자 수', unit: '명' }, { title: '이용 건수', unit: '건' }, { title: '대출 권수', unit: '권' }]));
[
  '10) 관외대출 : 장애인이 기관을 직접 내방하여 자료를 빌린 경우',
  '11) 책나래서비스 : 장애인이 필요로 하는 자료를 우체국 택배를 이용해 무료로 집까지 제공하는 서비스',
  '12) 우편·택배 서비스 : 책나래서비스 외 해당 기관에서 자체적으로 시행하는 우편·택배서비스',
  '13) 상호대차서비스 : 협약을 맺은 다른 도서관(장애인도서관, 타 지역 도서관 등)에 장애인 대체자료를 신청해 소장 자료를 서로 이용한 경우',
  '14) 방문대출 : 직원(자원봉사자 등)이 직접 장애인 가정을 방문하여 자료를 빌려 준 경우',
  '15) 이동도서관 : 접근성이 떨어지는 지역을 중심으로 자료를 대출·반납하여 주는 서비스',
  '16) 기타 : 상기 서비스 외 자체적으로 운영하는 자료대출 서비스명 기입',
].forEach(t => body.push(footnote(t)));

body.push(qHead('문18', '귀 기관의 장애대학(원)생 대상 서비스 현황을 기재해 주십시오. 해당사항이 없는 경우에는 0으로 작성해 주시면 됩니다.'));
body.push(tblProvide([
  ['대면낭독'], ['정보검색'], ['녹음자료 제작·제공'], ['점역자료 제작·제공'], ['입력자료(텍스트파일) 음성변환 프로그램'],
  ['한국수어 대면낭독'], ['자막지원'], ['독서보조기기 사용 안내'], ['정보화기기(컴퓨터, 스마트폰, 태블릿 기기 등) 사용 안내'], ['기타(기입: ________ )'],
], [{ title: '이용자 수', unit: '명' }, { title: '이용 건수', unit: '건' }]));
[
  '19) 대면낭독 : 책 읽기가 불편한 장애인을 대상으로 장서를 낭독해주는 서비스',
  '20) 녹음자료 제작·제공 : 시각장애인을 위해 도서를 낭독하여 CD, DVD, TTS 등으로 제작 또는 제공',
  '21) 점역자료 제작·제공 : 시각장애인을 위해 신간도서를 점자형태로 번역 제작·제공',
  '22) 입력자료(텍스트파일) 음성변환 프로그램 : 컴퓨터가 텍스트파일을 음성으로 변환해 읽어주는 서비스',
  '23) 한국수어 대면낭독 : 청각장애인을 대상으로 한국수어통역사가 한국수어로 낭독·부연 설명하는 서비스',
].forEach(t => body.push(footnote(t)));

body.push(qHead('문19', '귀 기관에서는 2025년 한 해 동안 장애대학(원)생을 대상으로 교육 및 독서문화프로그램을 운영한 경험이 있습니까?'));
body.push(...options(['① 있다', '② 없다']));
body.push(subqHead('문19-1', '2025년에 장애대학(원)생을 대상으로 운영한 교육 및 독서문화프로그램 횟수와 참가자 수를 기재해 주십시오.'));
body.push(help('장애유형 구분 없이 ‘장애인’으로만 되어 있거나 구체적 유형 구분이 어려운 경우 ‘기타 유형 장애’에 기입'));
body.push(help('횟수는 기간에 상관없이 1회로 보며, 참가자 수는 매회 참가자(참석자)를 누적으로 기입. 해당 경험이 없는 경우 ‘0’을 기재'));
body.push(tblQ19_1());
[
  '24) 도서관 이용 교육 : 견학, 오리엔테이션, 이용 안내 교육 등',
  '25) 정보 활용 교육 : 정보검색, 정보원 소개, 정보표현, 인용 및 표절 등 정보 이용 방법 교육',
  '26) 문화 프로그램 : 전시회, 강연회, 대회 등 도서관·독서 관련 프로그램을 제외한 모든 프로그램',
  '27) 독서 관련 프로그램 : 서비스와 직접 관련되는 도서, 독서지도, 정보 관련 강좌·행사·세미나 등',
].forEach(t => body.push(footnote(t)));

body.push(qHead('문20', '귀 기관에서는 2025년 한 해 동안 비장애인을 대상으로 장애인 인식 개선 교육이나 프로그램을 실시한 경험이 있습니까?'));
body.push(...options(['① 있다', '② 없다']));
body.push(subqHead('문20-1', '실시 횟수와 참가자 수를 기재해 주십시오.'));
body.push(fillLine('총 횟수', '회'));
body.push(fillLine('총 참가자수', '명'));

// ---- SECTION 6 ----
body.push(sectionHead('⑥', '대외 협력', '2025년 1월 1일 ~ 12월 31일 기준'));

body.push(qHead('문21', '귀 기관에서는 2025년에 장애대학(원)생을 위한 도서관서비스 홍보를 진행한 경험이 있습니까?'));
body.push(...options(['① 있다', '② 없다']));
body.push(subqHead('문21-1', '장애대학(원)생을 위한 도서관서비스 홍보를 한 경우, 어떤 방법으로 홍보가 진행되었습니까? (해당 항목에 모두 응답)'));
body.push(...options(['① 관내 이용 안내', '② 온·오프라인 소식지', '③ 팸플릿 등 홍보자료', '④ 기관 홈페이지', '⑤ 기관 소셜미디어', '⑥ 보도자료 배포 등 언론 홍보'], { stacked: true }));
body.push(fillLine('⑦ 기타', ''));

body.push(qHead('문22', '귀 기관에서는 2025년에 유관기관과 연계(외부협력)하여 장애대학(원)생을 위한 도서관서비스 및 관련 사업을 운영한 경험이 있습니까?'));
body.push(...options(['① 있다', '② 없다']));
body.push(subqHead('문22-1', '2025년에 협업한 기관 수는 대략 몇 개 기관입니까?'));
body.push(fillLine('협업 기관 수', '개'));
body.push(subqHead('문22-2', '2025년에 협업한 기관은 어떤 유형의 기관입니까? (해당 항목에 모두 응답)'));
body.push(...options(['① 공공도서관', '② 국립도서관', '③ 장애인도서관', '④ 장애인복지관 및 단체', '⑤ 특수학교(학급)', '⑥ 타 대학 장애학생지원센터·대학도서관'], { stacked: true }));
body.push(fillLine('⑦ 기타', ''));
body.push(subqHead('문22-3', '2025년에 협업한 서비스 및 사업 내용은 무엇입니까? (해당 항목에 모두 응답)'));
body.push(...options(['① 대체자료 제작 및 보급', '② 독서문화프로그램 운영', '③ 장애인 관련 교육 운영', '④ 독서보조기기 등 각종 장비 지원'], { stacked: true }));
body.push(fillLine('⑤ 기타', ''));

// ---- SECTION 7 ----
body.push(sectionHead('⑦', '장애학생지원센터 운영 및 학습지원 현황', '2025년 1월 1일 ~ 12월 31일 기준'));
body.push(p('※ 본 섹션의 문항은 국가장애인평생교육진흥센터 「2023년 장애대학생 교육복지지원 실태조사」의 관련 문항을 대학도서관·장애학생지원센터 상황에 맞게 재구성한 것입니다.', { size: 18, color: GREY, after: 140 }));

body.push(qHead('문23', '귀 기관에는 장애대학(원)생의 학습지원을 위한 장애학생지원센터가 독립적인 전담기구로 설치되어 있으며, 운영이 적절하게 이루어지고 있습니까? 아래 세부 항목에 답해 주십시오.'));
body.push(subqHead('문23-1', '장애학생지원센터의 현재 운영 형태는 어떠합니까?'));
body.push(...options(['① 독립 전담기구로 되어 있음', '② 독립 전담기구가 아님']));
body.push(fillLine('③ 기타', ''));
body.push(subqHead('문23-2', '장애학생지원센터의 현재 직원 형태는 어떠합니까? (해당 항목에 모두 응답)'));
body.push(...options(['① 전담 직원 있음', '② 겸직 직원 있음', '③ 전담 직원 없음']));
body.push(fillLine('④ 기타', ''));
body.push(subqHead('문23-3', '장애학생지원센터의 예산 운영은 어떠합니까?'));
body.push(...options([
  '① 장애학생지원센터의 예산이 다양한 비목으로 구성되어 있고, 예산에 따른 지출이 충실히 이루어지고 있음',
  '② 장애학생지원센터의 예산이 다양한 비목으로 구성되어 있으나, 예산에 따른 지출이 충실히 이루어지지 않고 있음',
  '③ 장애학생지원센터의 예산이 마련되어 있으나 다양한 비목으로 구성되어 있지 않고, 예산에 따른 지출이 제대로 이루어지지 않고 있음',
  '④ 장애학생지원센터의 예산이 별도로 마련되어 있지 않음',
], { stacked: true }));
body.push(fillLine('⑤ 기타', ''));

body.push(qHead('문24', '장애학생지원센터 전담교직원이 소지하고 있는 전문자격은 무엇입니까? (해당 항목에 모두 응답)'));
body.push(...options([
  '① 없음', '② 사회복지사', '③ 상담전문가', '④ 취업전문가', '⑤ 특수학교 교사', '⑥ 장애인재활상담사',
  '⑦ 속기사', '⑧ 점역·교정사', '⑨ 언어재활사', '⑩ 한국수어통역사(구, 수화통역사)', '⑪ 의지·보조기 기사',
], { stacked: true }));
body.push(fillLine('⑫ 기타', ''));

body.push(qHead('문25', '대학 차원에서 공식적으로 제공된 강의계획서에 장애학생을 위한 조정 안내를 제시하고 있습니까?'));
body.push(subqHead('문25-1', '강좌의 강의계획서에 장애유형별 맞춤형 지원 및 조정에 대한 구체적인 안내를 제공하고 있는 정도는 어떠합니까?'));
body.push(...options(['① 전혀 그렇지 않다', '② 그렇지 않다', '③ 보통이다', '④ 그렇다', '⑤ 매우 그렇다']));

body.push(qHead('문26', '장애학생을 위한 별도의 수강신청지원제도가 있습니까?'));
body.push(...options(['① 있다', '② 없다']));

body.push(qHead('문27', '장애학생이 강의신청 정보를 적절하게 파악하여 수강신청할 수 있도록 강의계획 및 수강신청의 콘텐츠는 접근성을 확보하고 있습니까?'));
body.push(subqHead('문27-1', '강의계획 관련 콘텐츠의 접근성은 어떠합니까?'));
body.push(...options(['① 우수', '② 보통', '③ 미흡']));

body.push(qHead('문28', '담당교수는 개별 장애학생의 장애특성(장애유형, 장애정도, 학습의 어려움 등)을 파악하고 지속적으로 피드백을 제공하고 있습니까?'));
body.push(...options(['① 있다', '② 없다']));
body.push(subqHead('문28-1', '개별 장애학생의 학습에 필요한 지원내용과 방법(장애유형 및 특성을 고려한 자료제시 및 강의전달방법)을 제시하고 지속적인 피드백을 제공하는 정도는 어떠합니까?'));
body.push(...options(['① 전혀 그렇지 않다', '② 그렇지 않다', '③ 보통이다', '④ 그렇다', '⑤ 매우 그렇다']));
body.push(subqHead('문28-2', '개별 장애학생의 교수·학습에 필요한 학습 자료를 과목담당 교수에게 지원(점자자료 제작지원, 녹음자료 제작지원 등)하고 있습니까?'));
body.push(...options(['① 있다', '② 없다']));

body.push(qHead('문29', '온라인 수업 콘텐츠(웹사이트를 통해 수강하는 온라인 수업 자체, 동영상 또는 멀티미디어 기반 파일)는 장애학생을 위한 접근성을 확보하고 있습니까?'));
body.push(subqHead('문29-1', '수어자막, 문자자막, 속기록을 제공하는 실태는 어떠합니까? (해당 항목에 모두 응답)'));
body.push(...options(['① 수어자막 제공', '② 문자자막 제공', '③ 속기록 제공', '④ 제공하지 않음']));
body.push(subqHead('문29-2', '동영상(멀티미디어) 온라인 수업 콘텐츠의 접근성을 위한 제어 기능 정도는 어떠합니까?'));
body.push(grid(frac(1.6, 1, 1, 1), [
  [tc('제어 기능', { header: true }), tc('우수', { header: true }), tc('보통', { header: true }), tc('미흡', { header: true })],
  ...['재생 / 일시정지 / 정지하기', '음량 조절', '화면 크기 조절', '화면 속도(재생 속도) 조절', '특정 화면(장면)으로 이동'].map(l =>
    [tc(l, { align: AlignmentType.LEFT }), tc('○', {}), tc('○', {}), tc('○', {})]),
]));
body.push(subqHead('문29-3', '온라인 학습관리시스템(Learning Management System, LMS)은 장애학생을 위해 접근성을 확보하고 있습니까?'));
body.push(...options(['① 있다', '② 없다']));

body.push(qHead('문30', '귀 기관의 디지털 도서관은 장애학생을 위하여 접근성을 확보하고 있습니까?'));
body.push(subqHead('문30-1', '디지털 도서관의 주요 메뉴와 기능의 접근성은 어떠합니까?'));
body.push(grid(frac(1.6, 1, 1, 1), [
  [tc('장애유형', { header: true }), tc('우수', { header: true }), tc('보통', { header: true }), tc('미흡', { header: true })],
  ...['시각장애학생을 위한 접근성', '청각장애학생을 위한 접근성', '지체장애·뇌병변장애학생을 위한 접근성'].map(l =>
    [tc(l, { align: AlignmentType.LEFT }), tc('○', {}), tc('○', {}), tc('○', {})]),
]));
body.push(subqHead('문30-2', '디지털 도서관에서 제공하는 서적 및 자료의 정보접근성은 어떠합니까?'));
body.push(grid(frac(1.6, 1, 1, 1), [
  [tc('장애유형', { header: true }), tc('우수', { header: true }), tc('보통', { header: true }), tc('미흡', { header: true })],
  ...['시각장애학생을 위한 접근성', '청각장애학생을 위한 접근성', '지체장애·뇌병변장애학생을 위한 접근성'].map(l =>
    [tc(l, { align: AlignmentType.LEFT }), tc('○', {}), tc('○', {}), tc('○', {})]),
]));

// ---- SECTION 8 ----
body.push(sectionHead('⑧', '기타', '2025년 1월 1일 ~ 12월 31일 기준'));
body.push(qHead('문31', '추가로 대학도서관·장애학생지원센터의 장애대학(원)생의 대체자료 제작 및 서비스 발전을 위해 의견이 있으시면 자유롭게 기술해 주세요.'));
for (let i = 0; i < 4; i++) {
  body.push(new Paragraph({
    spacing: { after: 260 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 4 } },
    children: [new TextRun({ text: ' ', font: FONT, size: 20 })],
  }));
}

body.push(qHead('P1', '마지막으로 본 설문조사 일괄 종료 후 감사의 의미로 추첨을 통해 답례를 제공할 예정입니다. 이에 전달/연락 받으실 귀하의 핸드폰 번호를 수집하고자 합니다. 귀하의 개인정보는 경품발송 및 중복 수령방지 등을 위한 목적으로만 사용되며 다른 용도로 사용되지 않고 파기할 것입니다. 개인정보의 제3자 제공에 대한 동의 여부는 어떻게 되십니까? (동의하지 않음을 선택하여도 설문결과는 통계자료 및 평가 지표로 활용됩니다.)'));
body.push(...options(['① 취급위탁에 동의합니다.', '② 취급위탁에 동의하지 않습니다.(경품지급불가)']));
body.push(subqHead('P2', '(P1의 ① 응답자만) 귀하의 핸드폰 번호는 어떻게 되십니까? (‘-’를 제외하고 숫자로만 입력)'));
body.push(fillLine('핸드폰 번호', ''));

body.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 400 },
  children: [new TextRun({ text: '♥ 끝까지 응답해 주셔서 대단히 감사합니다 ♥', font: FONT, size: 22, bold: true, color: NAVY })],
}));

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 20 } } } },
  sections: [{
    properties: {
      page: {
        size: { width: pageW, height: 16838 },
        margin: { top: convertInchesToTwip(0.8), bottom: convertInchesToTwip(0.8), left: marginLR, right: marginLR },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 4 } },
          children: [new TextRun({ text: '대학도서관 대체자료 소장 현황 및 서비스 실태조사 연구 (설문지 초안)', font: FONT, size: 15, color: GREY })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 17, color: GREY })],
        })],
      }),
    },
    children: body,
  }],
});

Packer.toBuffer(doc).then(buf => {
  const out = process.argv[2] || 'survey_output.docx';
  fs.writeFileSync(out, buf);
  console.log('written', out, buf.length, 'bytes');
});
