/*
  송탄영광교회 중고등부 · 출석체크 저장용 Apps Script
  Copyright (c) 2026 visionspeaker (실명: 강유석). All rights reserved.

  [배포 방법]
  1) 출석 스프레드시트를 연다
     https://docs.google.com/spreadsheets/d/1-moFQc8npkYh3ARL4SIpIB-Yft60d9srv7YBrmWpXXo/edit
  2) 상단 메뉴 [확장 프로그램] → [Apps Script] 클릭
  3) 기본 코드(Code.gs)를 지우고, 이 파일 내용을 통째로 붙여넣기 → 저장
  4) 우측 상단 [배포] → [새 배포] → 유형: '웹 앱'
       - 실행 계정: '나'(visionspeaker)
       - 액세스 권한: '모든 사용자'
     → [배포] → 권한 승인
  5) 나온 '웹 앱 URL'(끝이 /exec)을 복사해서 전달해 주세요.
     그 주소를 앱(index.html)의 ATTEND_API 에 넣으면 저장이 연결됩니다.

  [동작] 저장할 때마다 그 주일(날짜) 컬럼(H열부터)에 학생별 '출석'/'결석'을 기록합니다.
         같은 날짜로 다시 저장하면 그 컬럼을 덮어씁니다.
*/

var ATT_SHEET_ID = 1874453017;   // 출석 탭(gid)
var HEADER_ROW   = 2;            // 헤더 행 (이름/반 등)
var COL_NAME     = 2;            // B열: 이름
var COL_CLASS    = 4;            // D열: 반
var FIRST_DATE_COL = 8;          // H열부터 날짜별 이력 기록

function doPost(e){
  try{
    var body = JSON.parse(e.postData.contents);
    var date = String(body.date || '').trim();
    if(!date) return out({ ok:false, error:'날짜가 없습니다' });

    // records: [{name, present, reason}] — 없으면 present 배열로 대체
    var val = {};
    if (body.records && body.records.length) {
      body.records.forEach(function(r){
        var nm = String(r.name || '').trim();
        if (!nm) return;
        val[nm] = r.present ? '출석' : ((r.reason && String(r.reason).trim()) ? String(r.reason).trim() : '결석');
      });
    } else {
      (body.present || []).forEach(function(n){ val[String(n).trim()] = '출석'; });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = null, sheets = ss.getSheets();
    for(var i=0;i<sheets.length;i++){ if(sheets[i].getSheetId() === ATT_SHEET_ID){ sh = sheets[i]; break; } }
    if(!sh) sh = ss.getActiveSheet();

    var lastRow = sh.getLastRow();
    var lastCol = Math.max(sh.getLastColumn(), FIRST_DATE_COL);

    // 날짜 컬럼 찾기(없으면 새로 만들기)
    var hdr = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
    var col = -1;
    for(var c=FIRST_DATE_COL; c<=lastCol; c++){ if(fmtDate(hdr[c-1]) === date){ col = c; break; } }
    if(col === -1){
      col = FIRST_DATE_COL;
      while(col <= lastCol && String(hdr[col-1] || '').trim() !== '') col++;
    }
    sh.getRange(HEADER_ROW, col).setValue(date);

    // 학생 행별 출석/결석 기록
    var data = sh.getRange(1, 1, lastRow, COL_CLASS).getValues();
    var writes = [], cnt = 0;
    for(var r=HEADER_ROW+1; r<=lastRow; r++){
      var nm = String(data[r-1][COL_NAME-1] || '').trim();
      var cl = String(data[r-1][COL_CLASS-1] || '').trim();
      if(!nm || !cl || nm==='출석' || nm==='결석' || nm==='출석률'){ writes.push(['']); continue; }
      var v = (nm in val) ? val[nm] : '결석';
      if(v === '출석') cnt++;
      writes.push([v]);
    }
    if(writes.length) sh.getRange(HEADER_ROW+1, col, writes.length, 1).setValues(writes);

    return out({ ok:true, date:date, column:col, present:cnt });
  }catch(err){
    return out({ ok:false, error:String(err) });
  }
}

function doGet(){ return out({ ok:true, msg:'출석 API 연결됨' }); }

function fmtDate(v){
  if(v instanceof Date){
    var p = function(n){ return ('0'+n).slice(-2); };
    return v.getFullYear() + '-' + p(v.getMonth()+1) + '-' + p(v.getDate());
  }
  return String(v || '').trim();
}

function out(o){
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
