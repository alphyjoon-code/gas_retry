// =====================================================================
// 사전 PQ 예측 스크립트 v2
//
// [실행] 메뉴 → 사전PQ예측
//
// [활성 시트]
//   B1  : 용역명
//   B6  : 유사면적 (㎡)
//   B19 : 공고일
//
// [참조 시트 — 같은 스프레드시트]
//   업체매핑         : A=업체명(풀네임), B=사업자등록번호, C=단축명
//   경쟁업체실적5000이상
//   경쟁업체실적추정치
//
// [PDF 폴더]
//   Google Drive: 18Xw8bL7pxnHZo7VlU_M0sUkQ0roWZo-S
//
// [결과 파일]
//   ID: 1jgE8g_f1eowkX0xK-ABunDnveDnsbR0OCJBZuEr0ltw
// =====================================================================

function runPreBidPQEstimate() {
  var ui   = SpreadsheetApp.getUi();
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // ── 1. 활성 시트에서 기준값 읽기 ─────────────────────────────────
  var 유사면적Raw = sheet.getRange('B6').getValue();
  var 공고일Raw   = sheet.getRange('B19').getValue();

  if (!유사면적Raw || !공고일Raw) {
    ui.alert('오류: B6(유사면적) 또는 B19(공고일)이 비어 있습니다.');
    return;
  }

  var 유사면적     = parseFloat(유사면적Raw);
  var 공고일       = new Date(공고일Raw);
  var 유효기간시작 = new Date(공고일);
  유효기간시작.setFullYear(유효기간시작.getFullYear() - 5);

  // ── 2. 식별자 입력 팝업 ───────────────────────────────────────────
  var resp = ui.prompt(
    '🔮 사전 PQ 예측',
    '공고 식별자를 입력하세요 (예: A097)\n\n' +
    '유사면적: ' + 유사면적.toLocaleString() + '㎡  |  공고일: ' +
    Utilities.formatDate(공고일, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var code = resp.getResponseText().trim().toUpperCase();
  if (!code) { ui.alert('식별자가 없습니다.'); return; }

  ss.toast('드라이브에서 PDF 파일 검색 중...', '처리 중', 30);

  // ── 3. PDF 파일 탐색 ──────────────────────────────────────────────
  var folder = DriveApp.getFolderById('18Xw8bL7pxnHZo7VlU_M0sUkQ0roWZo-S');
  var iter   = folder.getFilesByType(MimeType.PDF);
  var target = null, targetMod = null;

  while (iter.hasNext()) {
    var f = iter.next();
    if (f.getName().toUpperCase().indexOf(code) !== -1) {
      var mod = f.getLastUpdated();
      if (!target || mod > targetMod) { target = f; targetMod = mod; }
    }
  }

  if (!target) {
    ui.alert('❌ 파일을 찾을 수 없습니다.\n폴더에 "' + code + '"이 포함된 PDF가 있는지 확인하세요.');
    return;
  }

  ss.toast(target.getName() + ' → OCR 변환 중...', '처리 중', 90);

  // ── 4. PDF → Google Docs 변환 → 텍스트 추출 ─────────────────────
  var pdfText = _pq2_extractText(target);
  if (!pdfText) {
    ui.alert('PDF 텍스트 추출에 실패했습니다.\n실행 로그를 확인해 주세요.');
    return;
  }

  // ── 5. 사업자등록번호 파싱 (줄 단위 XXX-XX-XXXXX 패턴) ───────────
  var brnSet = {};
  var BRN_RE = /^\d{3}-\d{2}-\d{5}$/;
  pdfText.split('\n').forEach(function(line) {
    var l = line.trim();
    if (BRN_RE.test(l)) brnSet[l.replace(/-/g, '')] = true;
  });

  if (Object.keys(brnSet).length === 0) {
    ui.alert('PDF에서 사업자등록번호를 추출하지 못했습니다.\n파일을 확인해 주세요.');
    return;
  }

  // ── 6. 업체매핑 / 실적 / 임계점 DB 로드 ──────────────────────────
  var mapSheet = ss.getSheetByName('업체매핑');
  var 실적Sheet  = ss.getSheetByName('경쟁업체실적5000이상');
  var 임계점Sheet = ss.getSheetByName('경쟁업체실적추정치');

  if (!mapSheet)   { ui.alert("'업체매핑' 시트가 없습니다."); return; }
  if (!실적Sheet)  { ui.alert("'경쟁업체실적5000이상' 시트가 없습니다."); return; }
  if (!임계점Sheet){ ui.alert("'경쟁업체실적추정치' 시트가 없습니다."); return; }

  var mapData   = mapSheet.getDataRange().getValues();   // A=업체명, B=사업자번호, C=단축명
  var 실적Data  = 실적Sheet.getDataRange().getValues();
  var 임계점Data = 임계점Sheet.getDataRange().getValues();

  // ── 7. 매핑 및 PQ 계산 ────────────────────────────────────────────
  var rows = [];

  for (var i = 1; i < mapData.length; i++) {
    var r       = mapData[i];
    var 풀네임  = (r[0] || '').toString().trim();
    var brn     = (r[1] || '').toString().replace(/-/g, '').trim();
    var 단축명  = (r[2] || '').toString().trim();

    if (!단축명 || !풀네임) continue;

    // 사업자번호 1순위, 풀네임 텍스트 포함 2순위
    var matched = (brn && brnSet[brn]) ||
                  (풀네임.length > 2 && pdfText.indexOf(풀네임) !== -1);
    if (!matched) continue;

    // PQ 계산
    var 건수, 근거;
    if (유사면적 >= 5000) {
      건수 = _pq2_countDB(실적Data, 단축명, 유사면적, 유효기간시작, 공고일);
      근거 = '실적DB ' + 건수 + '건';
    } else {
      var db  = _pq2_countDB(실적Data, 단축명, 5000, 유효기간시작, 공고일);
      var est = _pq2_threshold(임계점Data, 단축명, 유사면적);
      건수    = Math.min(5, Math.max(db, est));
      근거    = 'DB ' + db + '건 / 임계점 ' + est + '건 → ' + 건수 + '건';
    }

    rows.push([단축명, _pq2_score(건수), 건수, 근거, 풀네임]);
  }

  if (rows.length === 0) {
    ui.alert('매핑된 업체가 없습니다.\n업체매핑 시트의 사업자등록번호(B열)와 PDF 내용을 확인해 주세요.\n(추출된 사업자번호 수: ' + Object.keys(brnSet).length + ')');
    return;
  }

  // PQ 내림차순 정렬 후 번호 부여
  rows.sort(function(a, b) { return b[1] - a[1]; });
  for (var n = 0; n < rows.length; n++) rows[n].unshift(n + 1);

  // ── 8. 결과 파일에 새 시트 생성 ──────────────────────────────────
  var resultSS  = SpreadsheetApp.openById('1jgE8g_f1eowkX0xK-ABunDnveDnsbR0OCJBZuEr0ltw');
  var sheetName = code + ' PQ예측';
  var old = resultSS.getSheetByName(sheetName);
  if (old) resultSS.deleteSheet(old);
  var out = resultSS.insertSheet(sheetName, resultSS.getSheets().length);

  // 메타 헤더
  var 용역명 = (sheet.getRange('B1').getValue() || code).toString();
  out.getRange(1, 1, 1, 6).merge()
    .setValue(용역명 + '  |  유사면적 ' + 유사면적.toLocaleString() + '㎡  |  공고일 ' +
              Utilities.formatDate(공고일, Session.getScriptTimeZone(), 'yyyy-MM-dd'))
    .setBackground('#2D3748').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(11);

  // 컬럼 헤더
  out.getRange(2, 1, 1, 6)
    .setValues([['번호', '단축명', '추정PQ', '실적건수', '산출근거', '업체 풀네임']])
    .setBackground('#4A5568').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center');

  // 데이터
  out.getRange(3, 1, rows.length, 6).setValues(rows);
  out.getRange(3, 1, rows.length, 1).setHorizontalAlignment('center');
  out.getRange(3, 2, rows.length, 1).setFontWeight('bold');
  out.getRange(3, 3, rows.length, 1).setNumberFormat('0.0').setHorizontalAlignment('center');
  out.getRange(3, 4, rows.length, 1).setHorizontalAlignment('center');

  // PQ 그룹 경계선
  var prevPQ = null;
  for (var s = 0; s < rows.length; s++) {
    var cur = rows[s][2];
    if (prevPQ !== null && cur !== prevPQ) {
      out.getRange(s + 3, 1, 1, 6).setBorder(
        true, null, null, null, null, null,
        '#888888', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
      );
    }
    prevPQ = cur;
  }

  // 열 너비
  out.setColumnWidth(1, 45);
  out.setColumnWidth(2, 80);
  out.setColumnWidth(3, 70);
  out.setColumnWidth(4, 70);
  out.setColumnWidth(5, 220);
  out.setColumnWidth(6, 360);

  ss.toast('[' + sheetName + '] 생성 완료  ·  ' + rows.length + '개 업체', '✅ 완료', 8);
}


// ── PDF → 텍스트 추출 (Drive API OCR) ────────────────────────────────
function _pq2_extractText(file) {
  var docId = null;
  try {
    var ts       = new Date().getTime().toString();
    var boundary = 'PQ2_BOUND_' + ts;
    var CRLF     = '\r\n';
    var delim    = CRLF + '--' + boundary + CRLF;
    var close    = CRLF + '--' + boundary + '--';

    var meta = JSON.stringify({
      name    : 'pq2_ocr_' + ts,
      mimeType: MimeType.GOOGLE_DOCS
    });

    var pdfBytes = file.getBlob().getBytes();
    var b64      = Utilities.base64Encode(pdfBytes);

    var body =
      delim +
      'Content-Type: application/json; charset=UTF-8' + CRLF + CRLF +
      meta +
      delim +
      'Content-Type: application/pdf' + CRLF +
      'Content-Transfer-Encoding: base64' + CRLF + CRLF +
      b64 +
      close;

    var res = UrlFetchApp.fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true',
      {
        method         : 'POST',
        contentType    : 'multipart/related; boundary="' + boundary + '"',
        headers        : { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
        payload        : body,
        muteHttpExceptions: true
      }
    );

    var status = res.getResponseCode();
    var text   = res.getContentText();
    Logger.log('Drive upload status: ' + status);
    Logger.log('Drive upload response: ' + text.substring(0, 300));

    if (status !== 200) {
      Logger.log('OCR 업로드 실패: ' + text);
      return null;
    }

    var info = JSON.parse(text);
    docId = info.id;
    if (!docId) { Logger.log('docId 없음'); return null; }

    Utilities.sleep(4000);

    var doc     = DocumentApp.openById(docId);
    var content = doc.getBody().getText();
    Logger.log('추출 텍스트 앞 300자: ' + content.substring(0, 300));
    return content;

  } catch (e) {
    Logger.log('_pq2_extractText 오류: ' + e.toString());
    return null;
  } finally {
    if (docId) {
      try { DriveApp.getFileById(docId).setTrashed(true); } catch(e) {}
    }
  }
}


// ── 실적 DB 조회 ──────────────────────────────────────────────────────
// 열: E=4(준공일), F=5(발주면적), G=6(낙찰업체=단축명)
function _pq2_countDB(data, 단축명, 최소면적, 시작일, 공고일) {
  var cnt = 0;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if ((row[6] || '').toString().trim() !== 단축명) continue;
    if (!row[4]) continue;
    var 준공 = new Date(row[4]);
    if (준공 < 시작일 || 준공 > 공고일) continue;
    var 면적 = parseFloat((row[5] || '').toString().replace(/,/g, ''));
    if (isNaN(면적) || 면적 < 최소면적) continue;
    if (++cnt >= 5) break;
  }
  return cnt;
}


// ── 임계점 DB 조회 ────────────────────────────────────────────────────
// 열: A=0(단축명), D=3(이탈고착), E=4(이상고착)
function _pq2_threshold(data, 단축명, 유사면적) {
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if ((row[0] || '').toString().trim() !== 단축명) continue;
    var 이상 = row[4], 이탈 = row[3];
    if (이상 !== '' && 이상 != null && 유사면적 >= Number(이상)) return 1;
    if (이탈 !== '' && 이탈 != null && 유사면적 >= Number(이탈)) return 3;
    return 5;
  }
  return 3; // DB 미등록 업체: 보수적 처리
}


// ── PQ 점수 계산 (국방부 기준) ────────────────────────────────────────
function _pq2_score(건수) {
  var 부족 = Math.max(0, Math.min(4, 5 - Math.max(0, 건수)));
  return Math.round((100 - 부족 * 1.8) * 10) / 10;
}