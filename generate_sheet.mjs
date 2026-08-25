// 뷰니스 공동구매 사업부 모니터링 시스템 — 구글시트 업로드용 템플릿 생성 스크립트
// 실행: node generate_sheet.mjs  → 공동구매_모니터링_템플릿.xlsx 생성
// (D:\Claude\node_modules 의 exceljs 를 상위 디렉터리 탐색으로 그대로 사용합니다)
import ExcelJS from 'exceljs';

const wb = new ExcelJS.Workbook();
wb.creator = 'BEAUNESS'; wb.created = new Date();

// ── 색상 (기존 BEAUNESS 템플릿과 동일 규칙)
const C = {
  headerDark: 'FF1F4E79', headerMid: 'FF2E75B6',
  input: 'FFFFFACD',   // 노란색: 직접 입력
  auto: 'FFDDEEFF',    // 파란색: 수식 자동계산 (수정 금지)
  warn: 'FFFFE4E1', kpi: 'FFD6E4F0', section: 'FFEBF3FB',
  white: 'FFFFFFFF', black: 'FF000000',
  danger: 'FFFFB3B3', warning: 'FFFFE0B2', normal: 'FFC8E6C9',
};
const F = { krw: '#,##0"원"', pct: '0.0%', date: 'YYYY-MM-DD', num: '#,##0', score: '0.0' };

// ── 공통 헬퍼
const font = (o = {}) => ({ name: '맑은 고딕', size: 10, bold: false, color: { argb: C.black }, ...o });
const fill = a => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: a } });
const thin = { style: 'thin', color: { argb: 'FFBFBFBF' } };
const bdr = () => ({ top: thin, left: thin, bottom: thin, right: thin });
const ac = (h = 'center', v = 'middle', w = true) => ({ horizontal: h, vertical: v, wrapText: w });
const al = () => ({ horizontal: 'left', vertical: 'middle', wrapText: false });

function titleRow(ws, title, sub, lastCol = 'S') {
  ws.getRow(1).height = 28;
  ws.mergeCells(`A1:${lastCol}1`);
  Object.assign(ws.getCell('A1'), {
    value: title, font: font({ bold: true, size: 13, color: { argb: C.white } }),
    fill: fill(C.headerDark), alignment: ac(),
  });
  ws.getRow(2).height = 16;
  ws.mergeCells(`A2:${lastCol}2`);
  Object.assign(ws.getCell('A2'), {
    value: sub, font: font({ size: 9, color: { argb: 'FF595959' } }), alignment: al(),
  });
}

function hdrRow(ws, row, headers, bgArgb = C.headerMid, fgArgb = C.white) {
  ws.getRow(row).height = 32;
  headers.forEach((h, i) => {
    const c = ws.getCell(row, i + 1);
    c.value = h; c.font = font({ bold: true, color: { argb: fgArgb } });
    c.fill = fill(bgArgb); c.alignment = ac(); c.border = bdr();
  });
}

function cell(ws, r, col, val, { fmt, fillArgb = C.input, formula = false, align = 'left' } = {}) {
  const c = ws.getCell(r, col);
  if (formula) c.value = { formula: val };
  else if (val !== null && val !== undefined && val !== '') c.value = val;
  c.font = font(); c.border = bdr();
  c.alignment = align === 'center' ? ac() : align === 'right' ? { horizontal: 'right', vertical: 'middle' } : al();
  if (fmt) c.numFmt = fmt;
  if (fillArgb) c.fill = fill(fillArgb);
}

function dv(ws, ref, formulae) { ws.dataValidations.add(ref, { type: 'list', allowBlank: true, formulae: [formulae] }); }
function colW(ws, widths) { widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; }); }

// =====================================================================
// 1. 기본가정
// =====================================================================
{
  const ws = wb.addWorksheet('기본가정', { properties: { tabColor: { argb: C.headerDark } } });
  titleRow(ws, '기본가정', 'BEAUNESS 공동구매 모니터링 — 예측점수 가중치 및 기준값 (담당: 사업부 리더)', 'D');
  hdrRow(ws, 3, ['항목', '값', '설명', ''].slice(0, 3));
  const rows = [
    ['기준연월', '2026-08', 'YYYY-MM 형식'],
    ['셀러영향력점수 가중치', 0.4, '예측점수 = 셀러영향력×w1 + 공구모멘텀×w2 + 시장수요×w3'],
    ['공구모멘텀점수 가중치', 0.3, ''],
    ['시장수요점수 가중치', 0.3, '3개 가중치 합계는 1.0 이어야 합니다'],
  ];
  rows.forEach((r, i) => {
    const rn = 4 + i;
    cell(ws, rn, 1, r[0], { fillArgb: C.section, align: 'left' });
    cell(ws, rn, 2, r[1], { fillArgb: C.input, align: 'center', fmt: typeof r[1] === 'number' && r[1] < 1 ? F.pct : undefined });
    cell(ws, rn, 3, r[2], { fillArgb: null, align: 'left' });
  });
  cell(ws, 8, 1, '가중치 합계', { fillArgb: C.section });
  cell(ws, 8, 2, 'SUM(B5:B7)', { formula: true, fillArgb: C.auto, align: 'center', fmt: F.pct });
  ws.addConditionalFormatting({
    ref: 'B8', rules: [{ type: 'cellIs', operator: 'notEqual', formulae: ['1'], priority: 1,
      style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.danger }, fgColor: { argb: C.danger } }, font: { bold: true } } }],
  });
  colW(ws, [24, 14, 55]);
}

// =====================================================================
// 2. 카테고리마스터
// =====================================================================
{
  const ws = wb.addWorksheet('카테고리마스터', { properties: { tabColor: { argb: 'FF375623' } } });
  titleRow(ws, '카테고리마스터', '신규 제품 카테고리 추가 시 먼저 이 시트에 등록 (담당: 사업부)', 'D');
  hdrRow(ws, 3, ['카테고리ID', '대분류', '소분류', '비고']);
  const cats = [
    ['CAT01', '뷰티', '스킨케어'], ['CAT02', '뷰티', '색조/헤어'],
    ['CAT03', '헬스', '이너뷰티/건강기능식품'], ['CAT04', '헬스', '다이어트/식단'],
    ['CAT05', '생활', '유아동'], ['CAT06', '생활', '주방/생활가전'],
    ['CAT07', '식품', '간편식/디저트'], ['CAT08', '패션', '의류/잡화'],
  ];
  cats.forEach((r, i) => {
    const rn = 4 + i;
    cell(ws, rn, 1, r[0], { align: 'center' }); cell(ws, rn, 2, r[1]); cell(ws, rn, 3, r[2]); cell(ws, rn, 4, '');
  });
  colW(ws, [12, 14, 24, 30]);
}

// =====================================================================
// 3. 셀러마스터
// =====================================================================
{
  const ws = wb.addWorksheet('셀러마스터', { properties: { tabColor: { argb: 'FF375623' } } });
  titleRow(ws, '셀러마스터', '공동구매를 진행하는 셀러(인플루언서) 등록 (담당: 마케팅팀)', 'K');
  const headers = ['셀러ID', '셀러명', '대표플랫폼', '인스타그램 핸들', '유튜브 채널URL', '팔로워수(IG)', '구독자수(YT)', '등급', '등급점수(자동)', '주요카테고리ID', '소속MCN/연락처'];
  hdrRow(ws, 3, headers);
  const sellers = [
    ['SEL01', '뷰니스언니', '인스타그램', '@beauness_unni', '', 85000, 0, '매크로', 'CAT01', '자체'],
    ['SEL02', '헬시한윤아', '유튜브', '', 'youtube.com/@healthyyuna', 12000, 210000, '메가', 'CAT03', 'ABC MCN'],
    ['SEL03', '육아템민지', '인스타그램', '@mommy_minji', '', 32000, 0, '마이크로', 'CAT05', '자체'],
    ['SEL04', '다이어터소희', '인스타그램', '@sohee_diet', '', 8000, 0, '나노', 'CAT04', '자체'],
    ['SEL05', '리빙템수아', '유튜브', '', 'youtube.com/@livingsua', 0, 45000, '마이크로', 'CAT06', 'XYZ MCN'],
  ];
  sellers.forEach((r, i) => {
    const rn = 4 + i;
    r.forEach((v, ci) => cell(ws, rn, ci + 1, v, { align: [5, 6].includes(ci) ? 'right' : 'left', fmt: [5, 6].includes(ci) ? F.num : undefined }));
    // 등급점수(자동) = I열, 등급(H열) 기반 매핑
    cell(ws, rn, 9, `IFS(H${rn}="나노",25,H${rn}="마이크로",50,H${rn}="매크로",75,H${rn}="메가",100,TRUE,0)`,
      { formula: true, fillArgb: C.auto, align: 'center', fmt: F.num });
  });
  dv(ws, 'H4:H200', '"나노,마이크로,매크로,메가"');
  dv(ws, 'J4:J200', "카테고리마스터!$A$4:$A$100");
  colW(ws, [10, 16, 12, 18, 26, 12, 12, 10, 12, 14, 16]);
}

// =====================================================================
// 4. 채널마스터
// =====================================================================
{
  const ws = wb.addWorksheet('채널마스터', { properties: { tabColor: { argb: 'FF375623' } } });
  titleRow(ws, '채널마스터', '공동구매가 실제로 열리는 채널(URL) 등록 (담당: 마케팅팀)', 'D');
  hdrRow(ws, 3, ['채널ID', '채널유형', '채널명', 'URL']);
  const channels = [
    ['CH01', '인스타그램', '뷰니스언니 공구스토리', 'instagram.com/beauness_unni'],
    ['CH02', '유튜브', '헬시한윤아 커뮤니티탭', 'youtube.com/@healthyyuna/community'],
    ['CH03', '네이버카페', '육아템민지 카페 공구게시판', 'cafe.naver.com/mommyminji'],
    ['CH04', '오픈채팅', '다이어터소희 공구방', 'open.kakao.com/o/soheediet'],
    ['CH05', '자사몰', '뷰니스 공식몰 공구관', 'beauness.co.kr/gonggu'],
  ];
  channels.forEach((r, i) => { const rn = 4 + i; r.forEach((v, ci) => cell(ws, rn, ci + 1, v, { align: ci === 0 ? 'center' : 'left' })); });
  dv(ws, 'B4:B200', '"인스타그램,유튜브,네이버카페,오픈채팅,자사몰,기타"');
  colW(ws, [10, 14, 26, 34]);
}

// =====================================================================
// 5. 공동구매마스터 (핵심 Fact 테이블)
// =====================================================================
{
  const ws = wb.addWorksheet('공동구매마스터', { properties: { tabColor: { argb: 'FFC00000' } } });
  titleRow(ws, '공동구매마스터', '전체 공동구매 건 등록 — 신규 발견 즉시 입력 (담당: 전 담당자)', 'V');
  const headers = ['No', '제품명(SKU)', '카테고리ID', '셀러ID', '진행채널ID', '오픈일', '마감일', '상태(자동)',
    '정가', '판매가(공구가)', '할인율(자동)', '목표판매수량', '실판매수량', '예상매출액(자동)', '실매출액(자동)',
    '수수료율', '예상순이익(자동)', '셀러등급점수(자동)', '데이터출처', '등록자', '등록일', '비고'];
  hdrRow(ws, 3, headers);

  // 오늘(2026-08-25) 기준 과거/현재/미래 예시 데이터
  // [No, 제품명, 카테고리ID, 셀러ID, 채널ID, 오픈일, 마감일, 정가, 판매가, 목표판매수량, 실판매수량, 수수료율, 데이터출처, 등록자, 등록일]
  const rows = [
    [1, '콜라겐 이너샷 30포', 'CAT03', 'SEL02', 'CH02', '2026-06-01', '2026-06-10', 39000, 27300, 1200, 1180, 0.15, '수동입력', '김마케터', '2026-05-28'],
    [2, '유아 실리콘 식판 세트', 'CAT05', 'SEL03', 'CH03', '2026-07-15', '2026-07-22', 25000, 18000, 800, 760, 0.12, '수동입력', '이담당', '2026-07-10'],
    [3, '저당 단백질 그래놀라', 'CAT04', 'SEL04', 'CH04', '2026-08-10', '2026-08-20', 22000, 15800, 1500, 950, 0.10, '수동입력', '박사원', '2026-08-05'],
    [4, '수분 진정 크림 3종', 'CAT01', 'SEL01', 'CH01', '2026-08-20', '2026-08-30', 42000, 29900, 2000, 0, 0.15, '수동입력', '김마케터', '2026-08-18'],
    [5, '무선 스틱 청소기', 'CAT06', 'SEL05', 'CH02', '2026-09-05', '2026-09-15', 189000, 139000, 300, 0, 0.10, '수동입력', '이담당', '2026-08-22'],
    [6, '이너뷰티 콤부차 2박스', 'CAT03', 'SEL02', 'CH02', '2026-09-20', '2026-09-30', 48000, 34900, 1000, 0, 0.15, '수동입력', '박사원', '2026-08-24'],
  ];
  rows.forEach((r, i) => {
    const rn = 4 + i;
    cell(ws, rn, 1, r[0], { align: 'center' });
    cell(ws, rn, 2, r[1]);
    cell(ws, rn, 3, r[2], { align: 'center' });
    cell(ws, rn, 4, r[3], { align: 'center' });
    cell(ws, rn, 5, r[4], { align: 'center' });
    cell(ws, rn, 6, r[5], { align: 'center', fmt: F.date });
    cell(ws, rn, 7, r[6], { align: 'center', fmt: F.date });
    // 상태(자동) = H열
    cell(ws, rn, 8, `IF(TODAY()<F${rn},"예정",IF(TODAY()<=G${rn},"진행중","종료"))`, { formula: true, fillArgb: C.auto, align: 'center' });
    cell(ws, rn, 9, r[7], { align: 'right', fmt: F.krw });
    cell(ws, rn, 10, r[8], { align: 'right', fmt: F.krw });
    // 할인율(자동) = K열
    cell(ws, rn, 11, `1-J${rn}/I${rn}`, { formula: true, fillArgb: C.auto, align: 'right', fmt: F.pct });
    cell(ws, rn, 12, r[9], { align: 'right', fmt: F.num });
    cell(ws, rn, 13, r[10], { align: 'right', fmt: F.num });
    // 예상매출액(자동) = N열
    cell(ws, rn, 14, `J${rn}*L${rn}`, { formula: true, fillArgb: C.auto, align: 'right', fmt: F.krw });
    // 실매출액(자동) = O열
    cell(ws, rn, 15, `J${rn}*M${rn}`, { formula: true, fillArgb: C.auto, align: 'right', fmt: F.krw });
    cell(ws, rn, 16, r[11], { align: 'right', fmt: F.pct });
    // 예상순이익(자동) = Q열
    cell(ws, rn, 17, `IF(M${rn}=0,N${rn},O${rn})*(1-P${rn})`, { formula: true, fillArgb: C.auto, align: 'right', fmt: F.krw });
    // 셀러등급점수(자동) = R열 (셀러마스터 VLOOKUP)
    cell(ws, rn, 18, `IFERROR(VLOOKUP(D${rn},셀러마스터!$A$4:$I$200,9,FALSE),0)`, { formula: true, fillArgb: C.auto, align: 'center', fmt: F.num });
    cell(ws, rn, 19, r[12], { align: 'left' });
    cell(ws, rn, 20, r[13], { align: 'left' });
    cell(ws, rn, 21, r[14], { align: 'center', fmt: F.date });
    cell(ws, rn, 22, '', { align: 'left' });
  });
  dv(ws, 'C4:C500', '카테고리마스터!$A$4:$A$100');
  dv(ws, 'D4:D500', '셀러마스터!$A$4:$A$200');
  dv(ws, 'E4:E500', '채널마스터!$A$4:$A$200');
  dv(ws, 'S4:S500', '"수동입력,인스타링크 자동입력,유튜브API,네이버API,기타"');
  colW(ws, [5, 22, 11, 10, 11, 12, 12, 10, 11, 13, 11, 11, 11, 14, 14, 9, 14, 12, 12, 10, 12, 20]);
  ws.views = [{ state: 'frozen', ySplit: 3, xSplit: 2 }];
}

// =====================================================================
// 6. 시장수요트래킹
// =====================================================================
{
  const ws = wb.addWorksheet('시장수요트래킹', { properties: { tabColor: { argb: 'FF806000' } } });
  titleRow(ws, '시장수요트래킹', '오픈마켓/홈쇼핑/검색트렌드 등 외부 수요 신호 기록 — 예측점수 산출에 사용 (담당: 마케팅/사업개발)', 'J');
  const headers = ['No', '제품/키워드', '카테고리ID', '채널유형', '채널명', '지표유형', '수요지수(1~100)', '수집일', '데이터출처', '비고'];
  hdrRow(ws, 3, headers);
  const rows = [
    [1, '콜라겐', 'CAT03', '검색트렌드', '네이버 데이터랩', '검색량지수', 78, '2026-08-20', '수동입력'],
    [2, '무선청소기', 'CAT06', '오픈마켓', '쿠팡 베스트', '베스트순위 환산', 65, '2026-08-21', '수동입력'],
    [3, '단백질 그래놀라', 'CAT04', '검색트렌드', '네이버 데이터랩', '검색량지수', 58, '2026-08-22', '수동입력'],
    [4, '수분크림', 'CAT01', '홈쇼핑', 'GS홈쇼핑 편성', '방송매출 환산', 70, '2026-08-19', '수동입력'],
    [5, '유아 식판', 'CAT05', '오픈마켓', '네이버쇼핑', '인기도순위 환산', 60, '2026-08-18', '수동입력'],
  ];
  rows.forEach((r, i) => { const rn = 4 + i; r.forEach((v, ci) => cell(ws, rn, ci + 1, v, { align: [0, 6].includes(ci) ? 'center' : 'left', fmt: ci === 6 ? F.num : ci === 7 ? F.date : undefined })); });
  dv(ws, 'C4:C500', '카테고리마스터!$A$4:$A$100');
  dv(ws, 'D4:D500', '"오픈마켓,홈쇼핑,검색트렌드,기타"');
  dv(ws, 'I4:I500', '"수동입력,유튜브API,네이버API,쿠팡API,기타"');
  colW(ws, [5, 20, 11, 12, 18, 16, 14, 12, 14, 24]);
}

// =====================================================================
// 7. 예측점수 (전부 수식 — 직접입력 금지)
// =====================================================================
{
  const ws = wb.addWorksheet('예측점수', { properties: { tabColor: { argb: 'FF1F4E79' } } });
  titleRow(ws, '예측점수', '카테고리별 히트 예상 점수 — 전체 수식 자동계산 (직접 입력 금지)', 'F');
  hdrRow(ws, 3, ['카테고리ID', '카테고리명', '셀러영향력점수', '공구모멘텀점수', '시장수요점수', '종합히트예상점수', '등급'].slice(0, 7));
  for (let i = 0; i < 8; i++) {
    const rn = 4 + i;
    const catRow = 4 + i; // 카테고리마스터와 1:1 대응 (8개)
    cell(ws, rn, 1, `카테고리마스터!A${catRow}`, { formula: true, fillArgb: C.auto, align: 'center' });
    cell(ws, rn, 2, `IFERROR(카테고리마스터!C${catRow},"")`, { formula: true, fillArgb: C.auto, align: 'left' });
    // 셀러영향력점수 = 해당 카테고리 공구를 진행한 셀러들의 등급점수 평균
    cell(ws, rn, 3, `IFERROR(AVERAGEIF(공동구매마스터!$C$4:$C$500,A${rn},공동구매마스터!$R$4:$R$500),0)`,
      { formula: true, fillArgb: C.auto, align: 'right', fmt: F.score });
    // 공구모멘텀점수 = 최근3개월 건수 - 이전3개월 건수 기반 50점 중심 가감
    cell(ws, rn, 4,
      `MIN(100,MAX(0,50+(COUNTIFS(공동구매마스터!$C$4:$C$500,A${rn},공동구매마스터!$F$4:$F$500,">="&EDATE(TODAY(),-3))-COUNTIFS(공동구매마스터!$C$4:$C$500,A${rn},공동구매마스터!$F$4:$F$500,">="&EDATE(TODAY(),-6),공동구매마스터!$F$4:$F$500,"<"&EDATE(TODAY(),-3)))*10))`,
      { formula: true, fillArgb: C.auto, align: 'right', fmt: F.score });
    // 시장수요점수 = 시장수요트래킹 평균 수요지수
    cell(ws, rn, 5, `IFERROR(AVERAGEIF(시장수요트래킹!$C$4:$C$500,A${rn},시장수요트래킹!$G$4:$G$500),0)`,
      { formula: true, fillArgb: C.auto, align: 'right', fmt: F.score });
    // 종합점수 = 가중합 (기본가정 B5:B7)
    cell(ws, rn, 6, `C${rn}*기본가정!$B$5+D${rn}*기본가정!$B$6+E${rn}*기본가정!$B$7`,
      { formula: true, fillArgb: C.kpi, align: 'right', fmt: F.score });
    cell(ws, rn, 7, `IF(F${rn}>=80,"S",IF(F${rn}>=60,"A",IF(F${rn}>=40,"B","C")))`,
      { formula: true, fillArgb: C.kpi, align: 'center' });
  }
  colW(ws, [12, 20, 16, 16, 14, 18, 8]);
}

// =====================================================================
// 8. 대시보드 (요약 KPI — 시트 자체 열람용, 실제 사용은 웹 대시보드 권장)
// =====================================================================
{
  const ws = wb.addWorksheet('대시보드', { properties: { tabColor: { argb: 'FF1F4E79' } } });
  titleRow(ws, '대시보드', 'KPI 요약 (직접 입력 금지, 상세 화면은 웹 대시보드 이용)', 'D');
  hdrRow(ws, 3, ['지표', '값', '비고']);
  const kpis = [
    ['진행중 공구 수', `COUNTIF(공동구매마스터!$H$4:$H$500,"진행중")`],
    ['예정 공구 수', `COUNTIF(공동구매마스터!$H$4:$H$500,"예정")`],
    ['종료 공구 수', `COUNTIF(공동구매마스터!$H$4:$H$500,"종료")`],
    ['등록 셀러 수', `COUNTA(셀러마스터!$A$4:$A$200)`],
    ['이번달 누적 실매출액', `SUMPRODUCT((TEXT(공동구매마스터!$F$4:$F$500,"YYYY-MM")=TEXT(TODAY(),"YYYY-MM"))*공동구매마스터!$O$4:$O$500)`],
    ['히트예상 S등급 카테고리 수', `COUNTIF(예측점수!$G$4:$G$11,"S")`],
  ];
  kpis.forEach((r, i) => {
    const rn = 4 + i;
    cell(ws, rn, 1, r[0], { fillArgb: C.section, align: 'left' });
    cell(ws, rn, 2, r[1], { formula: true, fillArgb: C.kpi, align: 'right', fmt: F.num });
    cell(ws, rn, 3, '', { fillArgb: null });
  });
  colW(ws, [26, 18, 30]);
}

await wb.xlsx.writeFile('공동구매_모니터링_템플릿.xlsx');
console.log('생성 완료: 공동구매_모니터링_템플릿.xlsx');
