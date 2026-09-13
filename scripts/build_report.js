const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, PageOrientation, Header, Footer,
  PageNumber, NumberFormat, VerticalAlign, convertInchesToTwip,
} = require('docx');

const FONT = 'Malgun Gothic';
const NAVY = '1F3A5F';
const NAVY_LIGHT = 'E8EEF5';
const GREY = '595959';
const LINE = 'BFBFBF';

const pageW = 11906; // A4 width in DXA (210mm)
const marginLR = convertInchesToTwip(1);
const usableW = pageW - marginLR * 2;

function h1(text, num) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 480, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 6 } },
    children: [
      new TextRun({ text: num ? `${num}. ` : '', bold: true, color: NAVY, font: FONT, size: 30 }),
      new TextRun({ text, bold: true, color: NAVY, font: FONT, size: 30 }),
    ],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 120 },
    children: [new TextRun({ text, bold: true, color: '2B4A73', font: FONT, size: 24 })],
  });
}
function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after ?? 160, line: 360 },
    alignment: opts.align || AlignmentType.JUSTIFIED,
    indent: opts.indent,
    children: Array.isArray(text) ? text : [new TextRun({ text, font: FONT, size: 21, color: '262626' })],
  });
}
function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'main-bullets', level },
    spacing: { after: 90, line: 340 },
    children: [new TextRun({ text, font: FONT, size: 21, color: '262626' })],
  });
}
function caption(text) {
  return new Paragraph({
    spacing: { before: 60, after: 220 },
    children: [new TextRun({ text, italics: true, font: FONT, size: 18, color: GREY })],
  });
}

function cell(text, opts = {}) {
  return new TableCell({
    width: { size: opts.width, type: WidthType.DXA },
    shading: opts.header ? { type: ShadingType.CLEAR, fill: NAVY, color: 'auto' } : (opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill, color: 'auto' } : undefined),
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({
        text, font: FONT, size: 19,
        bold: !!opts.header, color: opts.header ? 'FFFFFF' : '262626',
      })],
    })],
  });
}

function table(colWidths, headerRow, bodyRows) {
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
    rows: [
      new TableRow({
        tableHeader: true,
        children: headerRow.map((t, i) => cell(t, { width: colWidths[i], header: true, align: AlignmentType.CENTER })),
      }),
      ...bodyRows.map((r, ri) => new TableRow({
        children: r.map((t, i) => cell(t, { width: colWidths[i], fill: ri % 2 === 1 ? 'F5F7FA' : undefined })),
      })),
    ],
  });
}

const w = [
  Math.round(usableW * 0.30),
  Math.round(usableW * 0.70),
];
const w3 = [
  Math.round(usableW * 0.24),
  Math.round(usableW * 0.20),
  Math.round(usableW * 0.56),
];
const w4 = [
  Math.round(usableW * 0.30),
  Math.round(usableW * 0.16),
  Math.round(usableW * 0.14),
  Math.round(usableW * 0.40),
];

const doc = new Document({
  numbering: {
    config: [{
      reference: 'main-bullets',
      levels: [
        { level: 0, format: 'bullet', text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 260 } } } },
        { level: 1, format: 'bullet', text: '-', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 820, hanging: 260 } } } },
      ],
    }],
  },
  styles: {
    default: {
      document: { run: { font: FONT, size: 21 } },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: pageW, height: 16838 },
        margin: { top: convertInchesToTwip(0.9), bottom: convertInchesToTwip(0.9), left: marginLR, right: marginLR },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 4 } },
          children: [new TextRun({ text: '설문조사 대상기관 선정 방법 및 절차', font: FONT, size: 16, color: GREY })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: GREY }),
          ],
        })],
      }),
    },
    children: [
      // ---- Title block ----
      new Paragraph({
        spacing: { after: 40 },
        children: [new TextRun({ text: '2025년 대학도서관·장애학생지원센터 장애인서비스 현황조사', font: FONT, size: 20, color: GREY })],
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [new TextRun({ text: '설문조사 대상기관 선정 방법 및 절차', bold: true, font: FONT, size: 40, color: NAVY })],
      }),
      new Paragraph({
        spacing: { after: 360 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 16, color: NAVY, space: 8 } },
        children: [new TextRun({ text: '교육부 고시 자료·KERIS 대학도서관 통계·장애학생지원센터 현황조사 통합 정리 보고', italics: true, font: FONT, size: 20, color: GREY })],
      }),

      // ================= 기 =================
      h1('기(起) — 연구 배경 및 목적', 'Ⅰ'),
      p('본 연구는 전국 대학도서관 및 대학 내 장애학생지원센터를 대상으로 장애인서비스 현황을 파악하기 위한 설문조사를 계획하고 있다. 설문조사의 신뢰성과 대표성을 확보하기 위해서는 문항 설계 못지않게 “누구에게 보낼 것인가”, 즉 설문 대상기관 목록을 정확하게 확정하는 작업이 선행되어야 한다.'),
      p('그러나 대상기관을 확인할 수 있는 기존 자료는 단일하지 않았으며, 집계 목적과 단위가 서로 다른 세 종류의 공공데이터로 나뉘어 있었다.'),
      table(w3, ['자료', '집계 단위', '주요 내용'], [
        ['① 교육부 고시 「공시대상학교 리스트」(2026년 기준)', '학교 단위 · 403개교', '학교명, 본분교구분, 대학구분, 학교구분, 지역, 설립구분, 근거법령'],
        ['② KERIS 「대학도서관 통계」', '도서관 단위 · 445건', '도서관명, 본분교(캠퍼스) 구분, 지역, 시군구, 설립구분'],
        ['③ 「장애학생지원센터 현황조사」', '학부·대학원 구분 단위 · 249개교', '장애재학생 수(장애유형·중증/경증), 지원센터 유무, 지원인력 현황'],
      ]),
      caption('표 1. 대상기관 선정에 활용한 세 가지 원자료의 특성'),
      p('세 자료는 각각 학교, 도서관(캠퍼스), 학부·대학원이라는 서로 다른 집계 단위로 구성되어 있어, 단순히 세 파일을 나열하거나 이름으로만 대조할 경우 동일 기관이 중복 집계되거나 반대로 실재하는 기관이 누락될 위험이 컸다. 이에 세 자료를 하나의 기준으로 통합하여, 중복과 누락이 없는 단일 대상기관 목록을 구축하는 작업을 수행하였다.'),

      // ================= 승 =================
      h1('승(承) — 선정 원칙 및 처리 절차', 'Ⅱ'),
      h2('1. 대상기관 선정의 네 가지 원칙'),
      p('연구진 협의를 통해 다음 네 가지 원칙을 우선 확정하고, 이를 기준으로 세 자료를 정리하였다.'),
      bullet('① 교육부 고시 「공시대상학교 리스트」를 최종 판단 기준(master)으로 삼는다.'),
      bullet('② KERIS 대학도서관 통계에서 분교·캠퍼스별로 별도 집계된 도서관 통계는, 교육부 고시상 1개 학교로 등재되어 있는 경우 1개 대상기관으로 합산하여 처리한다.'),
      bullet('③ 최근 통폐합이 이루어진 학교는 통합 이후의 학교 1개 기관을 대상으로 한다.'),
      bullet('④ 장애학생지원센터가 설치·운영 중인 학교는 예외 없이 대상기관 목록에 포함한다.'),
      p('아울러 학교 유형의 포함 범위에 대해서는 연구진 검토를 거쳐, 유형(대학교·산업대학·교육대학·전문대학·사이버대학·대학원대학 등)을 따로 제한하지 않고 교육부 고시 목록에 포함된 403개교 전체를 대상기관으로 삼기로 결정하였다.'),

      h2('2. 자료 간 매칭(대조) 절차'),
      p('KERIS 통계와 장애학생지원센터 조사자료에 기재된 기관명은 “OO대학교 △△도서관”, “OO대학교(캠퍼스)_제2캠퍼스”처럼 교육부 고시 학교명에 도서관명·캠퍼스 표기가 덧붙는 형태였다. 이를 교육부 고시 학교명과 대조하기 위해 아래와 같이 단계적 매칭 절차를 설계하였다.'),
      table(w3, ['단계', '방법', '적용 사례'], [
        ['1단계', '원자료의 기관명이 교육부 고시 학교명과 완전히 일치하면 그대로 매칭', '“가톨릭대학교_제2캠퍼스” → “가톨릭대학교”'],
        ['2단계', '괄호·캠퍼스 표기 등 부가정보를 제거한 “기준명”이 교육부 고시상 유일하게 존재하면 해당 학교로 매칭', '“가야대학교(고령)_제2캠퍼스” → 유일한 “가야대학교(김해)”로 매칭'],
        ['3단계', '기준명이 교육부 고시상 2개 이상 존재하는 경우(본교/분교 별도 고시), 원자료에 포함된 캠퍼스 식별 키워드로 분기 판단', '“동국대학교 WISE캠퍼스…” → 키워드 “WISE” 확인 후 “동국대학교(WISE)”로 매칭(본교 “동국대학교”와 구분)'],
      ]),
      caption('표 2. 3단계 매칭 절차'),
      p('본교와 분교가 교육부 고시상 별도 학교로 등재되어 있는 경우는 5개교(건국대학교(글로컬), 고려대학교(세종), 동국대학교(WISE), 연세대학교(미래), 한양대학교(ERICA))로 확인되었으며, 이들은 3단계 절차에 따라 본교와 분교를 구분하여 각각 별도의 대상기관으로 유지하였다.'),
      p('또한 매칭 과정에서 “가톨릭대학교”라는 이름을 공유하지만 실제로는 서로 완전히 다른 독립 대학인 광주가톨릭대학교, 대구가톨릭대학교, 대전가톨릭대학교, 목포가톨릭대학교, 부산가톨릭대학교, 수원가톨릭대학교, 인천가톨릭대학교 7개교와, 이름이 비슷해 혼동하기 쉬운 세종대학교·고려대학교(세종)의 사례를 확인하였다. 이런 경우 단순 이름 포함 여부(부분일치)로 매칭하면 서로 다른 학교가 하나로 합쳐지는 오류가 발생하므로, 전체 기준명 일치를 우선 적용하여 오매칭을 방지하였다.'),
      p('이상의 절차를 적용한 결과, KERIS 도서관 통계 445건 중 442건(99.3%), 장애학생지원센터 조사 249개교 중 248개교(99.6%)가 자동으로 매칭되었다.'),

      h2('3. 사실관계 확인이 필요한 항목의 검증'),
      p('자동 매칭에 실패했거나 최근 명칭이 변경된 것으로 의심되는 항목은 웹 검색을 통해 통폐합·명칭변경 여부를 확인한 뒤 목록에 반영하였다. 확인된 사례는 다음과 같다.'),
      table(w4, ['원자료상 명칭', '반영 결과', '시기', '확인 내용'], [
        ['경남도립거창대학 / 경남도립남해대학', '국립창원대학교로 통합', '2026.3.', '글로컬대학 사업 선정에 따라 국립창원대학교로 통폐합'],
        ['서라벌대학교', '신경주대학교로 개편', '2024.3.', '경주대학교와 통합, 신경주대학교로 교명 변경'],
        ['원광보건대학교', '원광대학교로 흡수통합', '2026.2.', '원광대학교로 흡수통합'],
        ['국립강릉원주대학교', '강원대학교로 통합', '2026.3.', '강원대학교와 통합(’통합 강원대’), 4개 캠퍼스 체제로 개편'],
        ['전남도립대학교', '국립목포대학교로 통합', '2026.3.', '국립목포대학교와 2·4년제 통합 운영 모델로 통합'],
      ]),
      caption('표 3. 웹 검색으로 확인·반영한 최근 통폐합 5건'),
      p('아울러 자동 매칭에 실패한 항목 가운데 실제로는 설문 대상 범위 밖인 것으로 확인된 사례도 함께 정리하였다.'),
      table(w3, ['원자료상 명칭', '처리', '사유'], [
        ['국방대학교 / 육군사관학교', '교육부 고시 목록에 미등재 확인', '「국방대학교 설치법」 등 별도 법령 소관 특수(대학)원으로, 대학알리미 공시대상학교 목록에는 등재되어 있지 않음. 설문 포함 여부는 별도 정책 판단 필요'],
        ['국제예술대학교', '대상 제외', '평생교육법상 학교형태 평생교육시설(전공대학)로 고등교육법 제2조 학교가 아니어서 공시대상 아님(정상 운영 중, 폐교 아님)'],
        ['경남과학기술대학교', '중복 데이터로 제외', '2021년 경상대학교와 통합되어 경상국립대학교로 개편됨. 장애학생지원센터 조사자료에 옛 명칭의 데이터가 남아 있었으나 경상국립대학교 행에 현재 데이터가 별도로 존재'],
      ]),
      caption('표 4. 매칭 범위 밖으로 판단하여 별도 처리한 항목'),

      // ================= 전 =================
      h1('전(轉) — 정리 결과 및 특이사항', 'Ⅲ'),
      h2('1. 최종 대상기관 목록의 규모'),
      p('위 절차를 거쳐 최종적으로 확정한 대상기관은 총 403개교이며, 캠퍼스별 도서관 통계를 개별 행으로 구분할 경우 457행(도서관 단위)으로 구성된다. 장애학생지원센터 현황조사 결과를 반영한 분포는 다음과 같다.'),
      table(w, ['구분', '학교 수'], [
        ['장애학생지원센터 있음(O)', '198개교'],
        ['장애학생지원센터 없음(X)', '23개교'],
        ['조사대상 아님(전문대학·대학원대학 등, 데이터 없음)', '182개교'],
        ['합계', '403개교'],
      ]),
      caption('표 5. 장애학생지원센터 유무 분포'),
      p('가톨릭대학교(3개 캠퍼스), 강원대학교(4개 캠퍼스), 단국대학교(2개 캠퍼스) 등 하나의 학교에 여러 캠퍼스 도서관·지원센터 통계가 각각 집계되어 있는 경우, 이를 임의로 하나의 수치로 합산하지 않고 캠퍼스별 세부 내역을 별도로 함께 제공함으로써, 추후 캠퍼스 단위로 개별 발송이 필요할 경우에도 활용할 수 있도록 하였다.'),

      h2('2. 확인된 한계 및 유의사항'),
      bullet('교육부 고시 목록 자체도 최신 통합 현황을 완전하게 반영하지 못한 경우가 있었다. 강릉원주대학교→강원대학교, 전남도립대학교→국립목포대학교 통합은 2026년 3월 시행 예정이었음에도, 확인 시점 기준 교육부 고시 목록에는 이미 통합 후 명칭만 등재되어 있어 옛 명칭의 KERIS·지원센터 자료를 별도로 매칭해야 했다.'),
      bullet('국방대학교·육군사관학교는 대학정보공시 대상 기관에 해당함에도 교육부 고시 「공시대상학교 리스트」에는 나타나지 않아, 특수목적대학의 목록 편입 여부에 대한 별도의 정책적 판단이 필요하다.'),
      bullet('도서관부호·상세주소는 세 원자료 어디에도 포함되어 있지 않아, 국립중앙도서관 도서관코드 조회 시스템 등 별도 자료로 보완이 필요하다(현재 진행 중).'),

      // ================= 결 =================
      h1('결(結) — 결론 및 향후 과제', 'Ⅳ'),
      p('본 작업을 통해 집계 단위가 서로 다른 세 원자료(학교 단위·도서관 단위·학부/대학원 단위)를 교육부 고시 학교 목록을 기준점으로 통합함으로써, 중복 없이 그리고 누락 없이 정리된 403개교(도서관 단위 457행) 규모의 설문 대상기관 목록을 확보하였다. 특히 정적 데이터 대조만으로는 드러나지 않는 최근 통폐합 5건을 웹 검색을 통해 사실관계까지 확인하여 반영함으로써, 목록의 최신성과 신뢰도를 함께 확보할 수 있었다.'),
      p('본 정리 절차(① 기준자료 선정 → ② 단계적 명칭 매칭 → ③ 사실관계 검증 → ④ 예외 처리 원칙 수립)는 이후 유사한 전수조사형 설문의 대상기관 확정 작업에도 재사용할 수 있는 절차로 판단된다.'),
      p('향후 과제는 다음과 같다.'),
      bullet('① 국립중앙도서관 도서관코드 조회 결과를 반영하여 도서관부호·주소 정보를 보완한다.'),
      bullet('② 국방대학교·육군사관학교 등 특수목적대학의 설문 대상 포함 여부를 연구진 협의를 통해 확정한다.'),
      bullet('③ 확정된 목록을 공동연구진에게 공유하여 최종 검토·확정을 거친 뒤 설문 발송 명단으로 확정한다.'),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  require('fs').writeFileSync(process.argv[2] || 'output.docx', buf);
  console.log('written');
});
