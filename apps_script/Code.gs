/**
 * 뷰니스 공동구매 모니터링 — 인스타그램 링크 빠른등록용 Apps Script
 *
 * 설치 방법:
 * 1. 구글시트(공동구매_모니터링_템플릿) 열기 → 확장 프로그램 → Apps Script
 * 2. 기본 생성된 Code.gs 내용을 전부 지우고 이 파일 내용을 붙여넣기
 * 3. 저장 → 배포 → 새 배포 → 유형: "웹 앱"
 *    - 실행 계정: 나(본인)
 *    - 액세스 권한이 있는 사용자: 전체(익명 사용자 포함) — quick-add.html이 로그인 없이 호출하기 위함
 * 4. 배포 후 생성되는 "웹 앱 URL"을 복사해서 dashboard/quick-add.html 상단 CONFIG.APPS_SCRIPT_URL에 붙여넣기
 *
 * 주의: 이 스크립트는 이 구글시트 소유자 권한으로 실행되며, 공동구매마스터 시트에 행을 추가합니다.
 * 등록창에서 "신규 셀러/채널 직접 입력"을 사용한 경우 셀러마스터/채널마스터 시트에도 자동으로 행이 추가됩니다.
 */

const SHEET_NAME = '공동구매마스터';
const SELLER_SHEET_NAME = '셀러마스터';
const CHANNEL_SHEET_NAME = '채널마스터';

// 마스터 시트(셀러/채널)에서 ID 접두어(SEL/CH) 기준 다음 순번 ID를 계산합니다.
// 예: 기존에 SEL01~SEL05가 있으면 SEL06을 반환.
function nextMasterId(sheet, prefix) {
  const lastRow = sheet.getLastRow();
  let maxNum = 0;
  if (lastRow >= 4) {
    const ids = sheet.getRange(4, 1, lastRow - 3, 1).getValues().flat();
    ids.forEach(id => {
      const m = String(id).match(new RegExp('^' + prefix + '(\\d+)$'));
      if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
    });
  }
  const nextNum = maxNum + 1;
  return prefix + String(nextNum).padStart(2, '0');
}

// payload.newSeller가 있으면 셀러마스터에 신규 행을 추가하고 새 셀러ID를 반환합니다.
function createSeller(ss, newSeller) {
  const sheet = ss.getSheetByName(SELLER_SHEET_NAME);
  const id = nextMasterId(sheet, 'SEL');
  const row = sheet.getLastRow() + 1;
  const platform = (newSeller.igHandle || '').trim() ? '인스타그램' : '';
  const values = [
    id,
    newSeller.name || '',
    platform,
    newSeller.igHandle || '',
    '', // 유튜브 채널URL
    '', // 팔로워수(IG)
    '', // 구독자수(YT)
    newSeller.grade || '나노',
    '', // 등급점수(자동) - 아래서 수식 설정
    newSeller.catId || '',
    '', // 소속MCN/연락처
  ];
  sheet.getRange(row, 1, 1, values.length).setValues([values]);
  sheet.getRange(row, 9).setFormula(
    `=IFS(H${row}="나노",25,H${row}="마이크로",50,H${row}="매크로",75,H${row}="메가",100,TRUE,0)`);
  return id;
}

// payload.newChannel이 있으면 채널마스터에 신규 행을 추가하고 새 채널ID를 반환합니다.
function createChannel(ss, newChannel) {
  const sheet = ss.getSheetByName(CHANNEL_SHEET_NAME);
  const id = nextMasterId(sheet, 'CH');
  const row = sheet.getLastRow() + 1;
  const values = [
    id,
    newChannel.type || '인스타그램',
    newChannel.name || '',
    newChannel.url || '',
  ];
  sheet.getRange(row, 1, 1, values.length).setValues([values]);
  return id;
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let newSellerId = '';
    let newChannelId = '';
    if (payload.newSeller && payload.newSeller.name) {
      newSellerId = createSeller(ss, payload.newSeller);
      payload.sellerId = newSellerId;
    }
    if (payload.newChannel && payload.newChannel.name) {
      newChannelId = createChannel(ss, payload.newChannel);
      payload.channelId = newChannelId;
    }

    const sheet = ss.getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    const no = lastRow - 3; // 헤더가 3행이므로 데이터 1행의 No = 1

    const today = Utilities.formatDate(new Date(), 'GMT+9', 'yyyy-MM-dd');

    const values = [
      no,
      payload.product || '',
      payload.catId || '',
      payload.sellerId || '',
      payload.channelId || '',
      payload.openDate || today,
      payload.closeDate || '',
      '', // H 상태 - 아래서 수식으로 설정
      payload.listPrice || '',
      payload.salePrice || '',
      '', // K 할인율
      payload.targetQty || '',
      '', // M 실판매수량
      '', // N 예상매출액
      '', // O 실매출액
      payload.feeRate || '',
      '', // Q 예상순이익
      '', // R 셀러등급점수
      '인스타링크 자동입력',
      payload.registrant || '',
      today,
      payload.note || ('원본: ' + (payload.sourceUrl || '')),
    ];
    sheet.getRange(newRow, 1, 1, values.length).setValues([values]);

    sheet.getRange(newRow, 8).setFormula(
      `=IF(TODAY()<F${newRow},"예정",IF(TODAY()<=G${newRow},"진행중","종료"))`);
    sheet.getRange(newRow, 11).setFormula(`=IFERROR(1-J${newRow}/I${newRow},"")`);
    sheet.getRange(newRow, 14).setFormula(`=J${newRow}*L${newRow}`);
    sheet.getRange(newRow, 15).setFormula(`=J${newRow}*M${newRow}`);
    sheet.getRange(newRow, 17).setFormula(
      `=IFERROR(IF(M${newRow}=0,N${newRow},O${newRow})*(1-P${newRow}),"")`);
    sheet.getRange(newRow, 18).setFormula(
      `=IFERROR(VLOOKUP(D${newRow},${SELLER_SHEET_NAME}!$A$4:$I$200,9,FALSE),0)`);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, row: newRow, newSellerId: newSellerId, newChannelId: newChannelId }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: '공동구매 빠른등록 API 정상 동작 중' }))
    .setMimeType(ContentService.MimeType.JSON);
}
