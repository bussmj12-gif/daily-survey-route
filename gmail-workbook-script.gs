/**
 * ============================================================
 *  오늘의 검사일정 - Gmail 자동 수집 스크립트
 * ============================================================
 *
 *  하는 일:
 *   1) Gmail에서 "제목에 Workbook 포함 + 첨부파일 있음" 메일을 찾음
 *   2) 아직 처리 안 한 메일이면 첨부 엑셀을 읽어서 표(rows) 데이터로 변환
 *   3) 검사관님의 GitHub Gist에 "staging-data.json"이라는 별도 파일로 업로드
 *      (기존에 사이트가 쓰는 route-data.json은 건드리지 않음 -- 이건 그냥
 *       "대기소"일 뿐이고, 실제 반영은 검사관님이 사이트에서 직접 "적용하기"를
 *       눌러야 일어남)
 *   4) 처리한 메일에는 라벨을 붙여서 다음 실행 때 중복 처리하지 않음
 *
 *  ── 설정 방법 (최초 1회) ──
 *   1) script.google.com 접속 (검사관님 개인 Gmail 계정으로 로그인)
 *   2) 새 프로젝트 만들기 → 이 코드 전체를 붙여넣기
 *   3) 아래 "설정값" 부분 3개를 본인 값으로 수정
 *   4) 왼쪽 메뉴 "서비스" (+ 아이콘) → "Drive API" 추가 (엑셀 -> 표 변환에 필요)
 *   5) 상단 함수 선택 드롭다운에서 checkForNewWorkbook 선택 후 ▶ 실행
 *      → 처음 실행 시 권한 승인 화면이 뜨면 전부 허용
 *   6) 왼쪽 메뉴 "트리거"(시계 아이콘) → "트리거 추가"
 *      → 함수: checkForNewWorkbook / 이벤트 소스: 시간 기반 / 분 단위 타이머 / 15분마다
 *      → 저장
 *
 *  이후로는 15분마다 자동으로 새 워크북이 있는지 확인하고, 있으면
 *  대기소에 올려둡니다. 사이트에서 "적용하기"를 누르기 전까지는
 *  실제 목록에 아무 영향을 주지 않습니다.
 * ============================================================
 */

// ==================== 설정값 (본인 것으로 수정) ====================
const GITHUB_TOKEN = 'ghp_여기에_본인_GitHub_토큰';   // 사이트 ⚙ 동기화 설정에서 쓰는 것과 동일한 토큰
const GIST_ID = '여기에_본인_Gist_ID';                 // 사이트 ⚙ 동기화 설정에서 쓰는 것과 동일한 Gist ID
const GMAIL_SEARCH_QUERY = 'subject:Workbook has:attachment newer_than:3d -label:workbook-processed';
// ===================================================================

const STAGING_FILENAME = 'staging-data.json';
const PROCESSED_LABEL = 'workbook-processed';

function checkForNewWorkbook() {
  const threads = GmailApp.search(GMAIL_SEARCH_QUERY, 0, 5);
  if (threads.length === 0) {
    Logger.log('새 워크북 메일 없음');
    return;
  }

  // 가장 최근 메일 하나만 처리 (여러 개 와도 최신 것 기준)
  threads.sort((a, b) => b.getLastMessageDate() - a.getLastMessageDate());
  const thread = threads[0];
  const messages = thread.getMessages();
  const message = messages[messages.length - 1];

  const attachment = message.getAttachments().find(a => /\.xlsx?$/i.test(a.getName()));
  if (!attachment) {
    Logger.log('첨부파일 없음, 건너뜀');
    markProcessed(thread);
    return;
  }

  const rows = xlsxAttachmentToRows(attachment);
  const dateText = extractDateText(rows);

  uploadToStaging(rows, dateText);
  markProcessed(thread);

  Logger.log('처리 완료: ' + dateText);
}

// 엑셀 첨부파일을 임시 구글시트로 변환해서 표(2차원 배열) 형태로 읽어옵니다.
// (사이트에서 XLSX.js로 읽는 것과 동일한 "행 단위 배열" 구조를 맞추기 위함)
function xlsxAttachmentToRows(attachment) {
  const blob = attachment.copyBlob();
  const tempFile = Drive.Files.insert(
    { title: 'temp-workbook-' + Date.now(), mimeType: MimeType.GOOGLE_SHEETS },
    blob,
    { convert: true }
  );

  try {
    const ss = SpreadsheetApp.openById(tempFile.id);
    const sheet = ss.getSheets()[0];
    const values = sheet.getDataRange().getValues();
    // 빈 셀은 사이트 파싱 로직과 맞춰 빈 문자열로 통일
    return values.map(row => row.map(cell => (cell === null || cell === undefined) ? '' : cell));
  } finally {
    // 변환용 임시 구글시트는 드라이브에 남기지 않고 정리
    DriveApp.getFileById(tempFile.id).setTrashed(true);
  }
}

// 메일 내용에서 "Date : 2026.07.30" 같은 패턴을 찾아 배너에 쓸 날짜 텍스트를 뽑습니다.
// (사이트의 parseWorkbook()이 하는 것과 동일한 방식)
function extractDateText(rows) {
  let dateText = '';
  rows.forEach(row => {
    row.forEach(cell => {
      const m = String(cell).match(/Date\s*:\s*([\d.\-\/]+)/i);
      if (m) dateText = m[1];
    });
  });
  return dateText;
}

function uploadToStaging(rows, dateText) {
  const payload = {
    rows: rows,
    dateText: dateText,
    stagedAt: new Date().toISOString()
  };
  const url = 'https://api.github.com/gists/' + GIST_ID;
  const res = UrlFetchApp.fetch(url, {
    method: 'patch',
    headers: { Authorization: 'token ' + GITHUB_TOKEN },
    contentType: 'application/json',
    payload: JSON.stringify({
      files: { [STAGING_FILENAME]: { content: JSON.stringify(payload) } }
    }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) {
    throw new Error('Gist 업로드 실패: ' + res.getResponseCode() + ' ' + res.getContentText());
  }
}

function markProcessed(thread) {
  let label = GmailApp.getUserLabelByName(PROCESSED_LABEL);
  if (!label) label = GmailApp.createLabel(PROCESSED_LABEL);
  thread.addLabel(label);
}
