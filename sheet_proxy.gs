// ══════════════════════════════════════════════════════
// sheet_proxy.gs  —  시험지 생성 로그 + 회원 관리
//
// 【설정 방법】
//   1. Google Sheets 새 시트 생성 후 시트 ID 복사
//      (URL에서: /spreadsheets/d/[여기가_SHEET_ID]/edit)
//   2. 아래 SHEET_ID 에 붙여넣기
//   3. Apps Script → 배포 → 웹 앱
//      - 액세스: 모든 사용자
//      - 배포 후 URL을 exam_generator.html 의 SHEET_PROXY_URL 에 입력
//
// 【재배포 주의】
//   배포 관리 → 기존 배포 편집 → 새 버전 (URL 유지)
// ══════════════════════════════════════════════════════

const SHEET_ID        = '1eV-P_D1FOyx43jTDsr9wfly6yaq1nYxIj2kr-GsHum8';
const SHEET_NAME      = '시험지로그';   // 로그 시트
const MEMBER_SHEET    = '회원목록';     // 회원 관리 시트

// GET 방식 처리 (웹에서 호출 — CORS 우회)
function doGet(e) {
  const action = e.parameter.action || '';
  Logger.log('action=' + action + ' / id=' + (e.parameter.id || '') + ' / name=' + (e.parameter.name || ''));

  // ── 관리자에게 신청 알림 ──
  if (action === 'notify') {
    return notifyAdmin(e.parameter.id || '');
  }

  // ── 이메일/전화번호 조회 ──
  if (action === 'check') {
    return checkMember(e.parameter.id || '');
  }

  // ── 회원 등록 (관리자) ──
  if (action === 'register') {
    return registerMember(e.parameter.id || '', e.parameter.memo || '');
  }

  // ── 로그 기록 ──
  if (e.parameter.name) {
    return writeLog(e);
  }

  // ── 연결 확인 ──
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, message: 'sheet_proxy 정상 작동 중' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 관리자에게 신청 알림 (MailApp) ────────────────────
function notifyAdmin(id) {
  try {
    const ADMIN_EMAIL = 'khsq2011@gmail.com';
    const msg =
      '[전과목 시험지 생성기(중학교)] 이용 신청\n\n' +
      '신청자: ' + id + '\n\n' +
      '승인하려면 아래 GAS에서 관리자 버튼으로 등록하거나\n' +
      '로그인 페이지에서 관리자 버튼으로 직접 등록해 주세요.\n' +
      'https://giths-ops.github.io/myQuiz4mid/';

    MailApp.sendEmail({
      to:      ADMIN_EMAIL,
      subject: '[시험지 생성기] 이용 신청이 들어왔습니다 — ' + id,
      body:    msg
    });
    return result(true, '관리자 알림 발송 완료');
  } catch(err) {
    Logger.log('관리자 알림 오류: ' + err.message);
    return result(false, err.message);
  }
}

// ── 회원 조회 ──────────────────────────────────────────
function checkMember(id) {
  try {
    id = id.trim().toLowerCase();
    if (!id) return result(false, '입력값 없음');

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    let   sheet = ss.getSheetByName(MEMBER_SHEET);
    if (!sheet) return result(false, '미등록');

    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const stored = (data[i][2] || '').toString().trim().toLowerCase().replace(/^'/, '');
      if (stored === id) {
        return result(true, '등록된 회원', { name: data[i][1] || '' });
      }
    }
    return result(false, '미등록');
  } catch(err) {
    return result(false, err.message);
  }
}

// ── 회원 등록 ──────────────────────────────────────────
function registerMember(id, memo) {
  try {
    id = id.trim().toLowerCase();
    if (!id) return result(false, '입력값 없음');

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    let   sheet = ss.getSheetByName(MEMBER_SHEET);

    // 시트 없으면 자동 생성
    if (!sheet) {
      sheet = ss.insertSheet(MEMBER_SHEET);
      sheet.appendRow(['등록일', '이름', '이메일/전화번호', '메모']);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#4a4a6a').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    // 중복 확인
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const stored = (data[i][2] || '').toString().trim().toLowerCase().replace(/^'/, '');
      if (stored === id) {
        return result(true, '이미 등록된 회원');
      }
    }

    const kstStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    // 숫자로 시작하는 값은 앞에 ' 붙여 텍스트로 저장 (구글 시트가 앞자리 0 제거 방지)
    const idForSheet = /^\d/.test(id) ? "'" + id : id;
    sheet.appendRow([kstStr, '', idForSheet, memo]);

    // 승인 알림 발송
    const approvalMsg =
      '[전과목 시험지 생성기] 이용 승인 안내\n\n' +
      '안녕하세요!\n회원 가입이 승인되었습니다.\n' +
      '아래 주소에서 서비스를 이용하실 수 있습니다.\n' +
      'https://giths-ops.github.io/myQuiz4mid/';

    const isPhone = /^01[016789]\d{7,8}$/.test(id.replace(/-/g, ''));
    const isEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(id);

    if (isPhone) sendApprovalSms(id.replace(/-/g, ''), approvalMsg);
    if (isEmail) sendApprovalEmail(id, approvalMsg);

    return result(true, '등록 완료');
  } catch(err) {
    return result(false, err.message);
  }
}

// ── 승인 SMS (솔라피) ──────────────────────────────────
function sendApprovalSms(receiver, msg) {
  try {
    const API_KEY    = PropertiesService.getScriptProperties().getProperty('SOLAPI_KEY');
    const API_SECRET = PropertiesService.getScriptProperties().getProperty('SOLAPI_SECRET');
    const SENDER     = '01026989056';

    const dateTime  = new Date().toISOString();
    const salt      = Utilities.getUuid().replace(/-/g, '');
    const signBytes = Utilities.computeHmacSha256Signature(
      dateTime + salt, API_SECRET, Utilities.Charset.UTF_8
    );
    const signature = signBytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
    const authHeader = `HMAC-SHA256 apiKey=${API_KEY}, date=${dateTime}, salt=${salt}, signature=${signature}`;

    const res = UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json; charset=utf-8', 'Authorization': authHeader },
      payload: JSON.stringify({ message: { to: receiver, from: SENDER, text: msg } }),
      muteHttpExceptions: true
    });
    Logger.log('솔라피 응답: ' + res.getResponseCode() + ' / ' + res.getContentText());
  } catch(err) {
    Logger.log('승인 SMS 오류: ' + err.message);
  }
}

// ── 승인 이메일 (GAS MailApp) ──────────────────────────
function sendApprovalEmail(toEmail, msg) {
  try {
    MailApp.sendEmail({
      to:      toEmail,
      subject: '[전과목 시험지 생성기] 이용이 승인되었습니다',
      body:    msg
    });
  } catch(err) {
    Logger.log('승인 이메일 오류: ' + err.message);
  }
}

// ── 로그 기록 ──────────────────────────────────────────
function writeLog(e) {
  try {
    const name   = e.parameter.name   || '(미입력)';
    const tokens = parseInt(e.parameter.tokens || '0');
    const input  = parseInt(e.parameter.input  || '0');
    const output = parseInt(e.parameter.output || '0');
    const email  = e.parameter.email  || '(미로그인)';
    const subj   = e.parameter.subj   || '';
    const ip     = e.parameter.ip     || '(알 수 없음)';
    const kstStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    let   sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['일시(KST)', '학생이름', '로그인계정', '과목', '입력토큰', '출력토큰', '합계토큰', 'IP주소']);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#4a4a6a').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([kstStr, name, email, subj, input, output, tokens, ip]);
    return result(true);
  } catch(err) {
    Logger.log('로그 오류: ' + err.message);
    return result(false, err.message);
  }
}

// ── 공통 응답 헬퍼 ────────────────────────────────────
function result(success, message, extra) {
  const obj = { success: success };
  if (message) obj.message = message;
  if (extra)   Object.assign(obj, extra);
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// POST 방식 처리 (blue-tennis-reservation.html 과 동일 방식)
function doPost(e) {
  // GAS 에디터에서 직접 실행 시 안내
  if (!e || !e.postData) {
    Logger.log('⚠️ doPost를 직접 실행하지 마세요. testLog() 함수를 실행하세요.');
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'testLog() 함수로 테스트하세요.' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    const payload = JSON.parse(e.postData.contents);
    const name   = payload.name   || '(미입력)';
    const tokens = parseInt(payload.tokens || '0');
    const input  = parseInt(payload.input  || '0');
    const output = parseInt(payload.output || '0');
    const email  = payload.email  || '(미로그인)';
    const subj   = payload.subj   || '';
    const ip     = payload.ip     || '(알 수 없음)';
    const ts     = payload.ts     || new Date().toISOString();

    // 한국 시간 변환
    const kstStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    let   sheet = ss.getSheetByName(SHEET_NAME);

    // 시트가 없으면 새로 생성 + 헤더 작성
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['일시(KST)', '학생이름', '로그인계정', '과목', '입력토큰', '출력토큰', '합계토큰', 'IP주소']);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#4a4a6a').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([kstStr, name, email, subj, input, output, tokens, ip]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    Logger.log('오류: ' + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 테스트 함수 — GAS 에디터에서 직접 실행
function testLog() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        name:   '홍길동',
        tokens: 1531,
        input:  546,
        output: 985,
        email:  'test@example.com',
        subj:   '중2학년 국어',
        ip:     '1.2.3.4',
        ts:     new Date().toISOString()
      })
    }
  };
  const result = doPost(fakeEvent);
  Logger.log(result.getContent());
}
