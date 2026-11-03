/**
 * @OnlyCurrentDoc
 * 영업 성장률 대시보드 — Google Apps Script 웹앱
 *
 * 이 스크립트는 판매 데이터가 들어있는 구글 시트에 "붙어 있는" 스크립트입니다.
 * (시트 메뉴 > 확장 프로그램 > Apps Script 에서 생성)
 * @OnlyCurrentDoc 덕분에 이 스프레드시트 하나에만 접근 권한을 요청합니다.
 */

var SHEET_NAME = 'RawData';   // 데이터 시트 이름. 없으면 첫 번째 시트를 사용
var REQUIRED = ['연월', '팀', '담당자', '거래처', '제품', '매출액'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('영업 성장률 대시보드')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 대시보드가 호출: 헤더 + 행 데이터를 돌려줌 (날짜는 'yyyy-MM' 문자열로 변환) */
function getData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error('「' + sheet.getName() + '」 시트에 데이터가 없습니다.');

  var header = values[0].map(function (h) { return String(h).trim(); });
  var missing = REQUIRED.filter(function (c) { return header.indexOf(c) < 0; });
  if (missing.length) throw new Error('필수 컬럼이 없습니다: ' + missing.join(', '));

  var tz = ss.getSpreadsheetTimeZone();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (r.every(function (v) { return v === '' || v === null; })) continue;   // 빈 행 건너뜀
    rows.push(r.map(function (v) {
      // Date 객체는 웹앱으로 그대로 넘길 수 없으므로 문자열로 변환
      return Object.prototype.toString.call(v) === '[object Date]' ? Utilities.formatDate(v, tz, 'yyyy-MM') : v;
    }));
  }
  return {
    name: ss.getName(),
    header: header,
    rows: rows,
    fetchedAt: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm')
  };
}
