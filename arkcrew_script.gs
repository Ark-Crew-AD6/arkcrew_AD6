// ================================================================
// ArkCrew 참석 관리 - Google Apps Script 최종본 v11
// 시트 구조:
//   관리 시트 (Apps Script 배포) → 명단, 설정
//   훈련 시트 → 훈련_1주차 ~ raw_훈련_1주차
//   예배 시트 → 예배_1주차 ~ raw_예배_1주차
//   보강 시트 → 보강_1주차 ~ raw_보강_1주차
// ================================================================

// ⚙️ 여기만 수정하세요!
var SHEET_IDS = {
  "훈련": "1ngAb_blvmSfVmGt7WvoHzOItzKyeF3pggDuHKtcBqzE",
  "예배": "1IUACRVyssDR3euB6Mv9YjfiKK1UPR779673SWsQhni4",
  "보강": "1SXY654Bgad2WRUvoIg6NRrtQB5HpMTxBr1uGobS-kQg"
};
// ================================================================

// 관리 시트 (Apps Script 배포된 곳 - 명단/설정 있는 곳)
function getAdminSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// 타입별 데이터 시트
function getDataSS(type) {
  var id = SHEET_IDS[type];
  if (!id || id.includes("여기에")) return null; // 미설정 시 null 반환
  return SpreadsheetApp.openById(id);
}

// ----------------------------------------------------------------
// GET: 명단 조회 (관리 시트에서)
// ----------------------------------------------------------------
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  if (action === "getMembers")      return getMembersResponse();
  if (action === "getSubmitted")    return getSubmittedResponse(e.parameter);
  if (action === "getAllSubmitted")  return getAllSubmittedResponse(e.parameter);
  return ContentService.createTextOutput("ArkCrew 서버 정상 작동 중 ✅").setMimeType(ContentService.MimeType.TEXT);
}

// ----------------------------------------------------------------
// 해당 주차 전체 제출 데이터 조회 (type + week 기준, 크루 전체)
// ----------------------------------------------------------------
function getAllSubmittedResponse(params) {
  try {
    var type = params.type || "";
    var week = params.week || "";

    var ss = getDataSS(type);
    if (!ss) return ContentService.createTextOutput(JSON.stringify({ crews: {} })).setMimeType(ContentService.MimeType.JSON);

    var rawName = "raw_" + type + "_" + week;
    var rawSheet = ss.getSheetByName(rawName);
    if (!rawSheet || rawSheet.getLastRow() < 2) {
      return ContentService.createTextOutput(JSON.stringify({ crews: {} })).setMimeType(ContentService.MimeType.JSON);
    }

    // 전체 데이터 한번에 읽기
    var data = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 6).getValues();
    var crews = {}; // { "1크루": { "홍길동": { status, reason } } }

    data.forEach(function(row) {
      var crew   = String(row[2]).trim();
      var name   = String(row[3]).trim();
      var status = String(row[4]).trim();
      var reason = String(row[5]).trim();
      if (!crew || !name) return;
      if (!crews[crew]) crews[crew] = {};
      crews[crew][name] = { status: status, reason: reason };
    });

    return ContentService.createTextOutput(JSON.stringify({ crews: crews })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ crews: {}, error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ----------------------------------------------------------------
// 기존 제출 데이터 조회 (type + week + crew 기준)
// ----------------------------------------------------------------
function getSubmittedResponse(params) {
  try {
    var type = params.type || "";
    var week = params.week || "";
    var crew = params.crew || "";

    var ss = getDataSS(type);
    if (!ss) return ContentService.createTextOutput(JSON.stringify({ rows: [] })).setMimeType(ContentService.MimeType.JSON);

    var rawName = "raw_" + type + "_" + week;
    var rawSheet = ss.getSheetByName(rawName);
    if (!rawSheet || rawSheet.getLastRow() < 2) {
      return ContentService.createTextOutput(JSON.stringify({ rows: [] })).setMimeType(ContentService.MimeType.JSON);
    }

    var data = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 6).getValues();
    var rows = [];
    data.forEach(function(row) {
      if (String(row[2]).trim() === crew) {
        rows.push({
          name:   String(row[3]).trim(),
          status: String(row[4]).trim(),
          reason: String(row[5]).trim()
        });
      }
    });

    return ContentService.createTextOutput(JSON.stringify({ rows: rows })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ rows: [], error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getMembersResponse() {
  var sheet = getAdminSS().getSheetByName("명단");
  if (!sheet) return ContentService.createTextOutput(JSON.stringify({ error: "명단 시트 없음" })).setMimeType(ContentService.MimeType.JSON);
  var lastRow = sheet.getLastRow();
  var crews = {};
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 3).getValues().forEach(function(row) {
      var crew = String(row[0]).trim(), name = String(row[1]).trim();
      var role = String(row[2]).trim().toUpperCase();
      if (!crew || !name) return;
      if (!crews[crew]) crews[crew] = [];
      crews[crew].push({ name: name, isManager: role === "Y", isStaff: role === "S" });
    });
  }
  return ContentService.createTextOutput(JSON.stringify({ crews: crews })).setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------
// POST: 데이터 저장 → 정리본 생성
// ----------------------------------------------------------------
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var timestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm");

    var type = data.type || "훈련";
    var week = data.week;
    var crew = data.crew;
    var rows = data.rows;

    var ss = getDataSS(type);
    if (!ss) {
      var msg = type + " 시트 ID가 설정되지 않았습니다. SHEET_IDS를 확인해주세요.";
      return ContentService.createTextOutput(JSON.stringify({ result: "error", message: msg })).setMimeType(ContentService.MimeType.JSON);
    }

    // raw 시트 (숨김)
    var rawName = "raw_" + type + "_" + week;
    var rawSheet = ss.getSheetByName(rawName);
    if (!rawSheet) {
      rawSheet = ss.insertSheet(rawName);
      rawSheet.hideSheet();
      rawSheet.appendRow(["제출시간", "주차", "크루", "이름", "상태", "사유/시간"]);
      rawSheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#4f8ef7").setFontColor("white");
    }

    // 해당 크루 기존 데이터 삭제
    var rawLast = rawSheet.getLastRow();
    if (rawLast >= 2) {
      for (var i = rawLast; i >= 2; i--) {
        if (rawSheet.getRange(i, 3).getValue() === crew) rawSheet.deleteRow(i);
      }
    }

    // 새 데이터 일괄 입력 (F열 텍스트 형식 먼저)
    var insertStart = rawSheet.getLastRow() + 1;
    var newRawData = rows.map(function(row) {
      return [timestamp, row.week, row.crew, row.name, row.status, String(row.reason || "")];
    });
    rawSheet.getRange(insertStart, 6, newRawData.length, 1).setNumberFormat("@");
    rawSheet.getRange(insertStart, 1, newRawData.length, 6).setValues(newRawData);

    // raw 색상 - 새로 추가된 행만 적용 (전체 X → 빠름)
    var rawDataLast = rawSheet.getLastRow();
    if (rawDataLast >= insertStart) {
      var newBg = newRawData.map(function(row) {
        var c = "#ffffff";
        if (row[4] === "참석")      c = "#d1fae5";
        else if (row[4] === "불참") c = "#fee2e2";
        else if (row[4] === "부분참") c = "#fef3c7";
        return [c,c,c,c,c,c];
      });
      rawSheet.getRange(insertStart, 1, newRawData.length, 6).setBackgrounds(newBg);
    }

    // raw 저장 완료 즉시 성공 반환 (정리본은 트리거로 비동기 처리)
    // 트리거용 PropertiesService에 작업 저장
    var props = PropertiesService.getScriptProperties();
    props.setProperty("pending_report", JSON.stringify({ type: type, week: week, ssId: ss.getId() }));

    // 정리본 작성 (동기 - 트리거 없이도 동작하도록)
    writeReport(ss, type, week);

    return ContentService.createTextOutput(JSON.stringify({ result: "success" })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    writeErrorLog(err.toString(), data);
    return ContentService.createTextOutput(JSON.stringify({ result: "error", message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ----------------------------------------------------------------
// 정리본 작성
// ----------------------------------------------------------------
function writeReport(ss, type, week) {
  var rawSheet = ss.getSheetByName("raw_" + type + "_" + week);
  if (!rawSheet) return;

  // 명단/설정은 관리 시트에서 읽기
  var adminSS    = getAdminSS();
  var memberSheet = adminSS.getSheetByName("명단");
  if (!memberSheet) return;

  var weekNum = parseInt(week) || 0;
  var dateStr = getWeekDate(adminSS, week); // 설정도 관리 시트에서

  // 명단 읽기
  var allMembers = {};
  var mLast = memberSheet.getLastRow();
  if (mLast >= 2) {
    memberSheet.getRange(2, 1, mLast - 1, 3).getValues().forEach(function(row) {
      var crew = String(row[0]).trim(), name = String(row[1]).trim();
      var role = String(row[2]).trim().toUpperCase();
      if (!crew || !name) return;
      if (!allMembers[crew]) allMembers[crew] = [];
      allMembers[crew].push({ name: name, isManager: role === "Y", isStaff: role === "S" });
    });
  }

  // 제출 데이터 읽기
  var submitted = {};
  var rawLast = rawSheet.getLastRow();
  if (rawLast >= 2) {
    rawSheet.getRange(2, 1, rawLast - 1, 6).getValues().forEach(function(row) {
      if (row[3]) submitted[row[3]] = { status: row[4], reason: String(row[5] || "") };
    });
  }

  var crewNames = Object.keys(allMembers).sort(function(a, b) {
    return (parseInt(a) || 0) - (parseInt(b) || 0);
  });

  var traineeOnly = (type === "예배" || type === "보강");

  // 통계
  var attendMgr = 0, absentMgr = 0, lateMgr = 0, undecidedMgr = 0;
  var attendStaff = 0, absentStaff = 0, lateStaff = 0, undecidedStaff = 0, totalStaff = 0;
  var attendTr = 0, absentTr = 0, lateTr = 0, undecidedTr = 0;

  crewNames.forEach(function(crew) {
    (allMembers[crew] || []).forEach(function(m) {
      if (traineeOnly && (m.isManager || m.isStaff)) return;
      var status = (submitted[m.name] || {}).status || "";
      if (m.isManager) {
        if (status === "참석") attendMgr++; else if (status === "불참") absentMgr++;
        else if (status === "부분참") lateMgr++; else undecidedMgr++;
      } else if (m.isStaff) {
        totalStaff++;
        if (status === "참석") attendStaff++; else if (status === "불참") absentStaff++;
        else if (status === "부분참") lateStaff++; else undecidedStaff++;
      } else {
        if (status === "참석") attendTr++; else if (status === "불참") absentTr++;
        else if (status === "부분참") lateTr++; else undecidedTr++;
      }
    });
  });

  // 정리본 시트 (데이터 시트에)
  var reportName = type + "_" + week;
  var rSheet = ss.getSheetByName(reportName);
  if (!rSheet) rSheet = ss.insertSheet(reportName);
  rSheet.clearContents();
  rSheet.clearFormats();
  rSheet.getRange(1, 1, rSheet.getMaxRows(), 1).setNumberFormat("@");
  rSheet.setColumnWidth(1, 420);

  var r = 1;
  var typeLabel = { "훈련": "훈련참석여부", "예배": "예배참석", "보강": "보강" };
  var hasStaff  = !traineeOnly && totalStaff > 0;

  // 제목
  rSheet.getRange(r, 1).setValue("*AD " + (typeLabel[type] || type) + " " + weekNum + "주차").setFontWeight("bold").setFontSize(13); r++;
  if (dateStr) { rSheet.getRange(r, 1).setValue("-" + dateStr); r++; }
  r++;

  // 통계
  if (traineeOnly) {
    rSheet.getRange(r, 1).setValue("-훈련생").setFontWeight("bold"); r++;
    rSheet.getRange(r, 1).setValue("참석: " + attendTr + "명").setBackground("#d1fae5"); r++;
    rSheet.getRange(r, 1).setValue("불참: " + absentTr + "명").setBackground("#fee2e2"); r++;
    rSheet.getRange(r, 1).setValue("부분참: " + lateTr   + "명").setBackground("#fef3c7"); r++;
    rSheet.getRange(r, 1).setValue("미정: " + undecidedTr + "명").setBackground("#f3f4f6"); r++;
  } else if (hasStaff) {
    rSheet.getRange(r, 1).setValue("-간사 / 스태프 / 훈련생 / 전체").setFontWeight("bold"); r++;
    rSheet.getRange(r, 1).setValue("참석: " + attendMgr + " / " + attendStaff + " / " + attendTr + " / " + (attendMgr+attendStaff+attendTr) + "명").setBackground("#d1fae5"); r++;
    rSheet.getRange(r, 1).setValue("불참: " + absentMgr + " / " + absentStaff + " / " + absentTr + " / " + (absentMgr+absentStaff+absentTr) + "명").setBackground("#fee2e2"); r++;
    rSheet.getRange(r, 1).setValue("부분참: " + lateMgr   + " / " + lateStaff   + " / " + lateTr   + " / " + (lateMgr+lateStaff+lateTr)     + "명").setBackground("#fef3c7"); r++;
    rSheet.getRange(r, 1).setValue("미정: " + undecidedMgr + " / " + undecidedStaff + " / " + undecidedTr + " / " + (undecidedMgr+undecidedStaff+undecidedTr) + "명").setBackground("#f3f4f6"); r++;
  } else {
    rSheet.getRange(r, 1).setValue("-간사 / 훈련생 / 전체").setFontWeight("bold"); r++;
    rSheet.getRange(r, 1).setValue("참석: " + attendMgr + " / " + attendTr + " / " + (attendMgr+attendTr) + "명").setBackground("#d1fae5"); r++;
    rSheet.getRange(r, 1).setValue("불참: " + absentMgr + " / " + absentTr + " / " + (absentMgr+absentTr) + "명").setBackground("#fee2e2"); r++;
    rSheet.getRange(r, 1).setValue("부분참: " + lateMgr   + " / " + lateTr   + " / " + (lateMgr+lateTr)    + "명").setBackground("#fef3c7"); r++;
    rSheet.getRange(r, 1).setValue("미정: " + undecidedMgr + " / " + undecidedTr + " / " + (undecidedMgr+undecidedTr) + "명").setBackground("#f3f4f6"); r++;
  }
  r++;

  // 크루별 정리
  crewNames.forEach(function(crew) {
    var managers = [], mgAbsent = [], mgLate = [], mgUndecided = [];
    var staffs = [], stAbsent = [], stLate = [], stUndecided = [];
    var trainees = [], trAbsent = [], trLate = [], undecided = [];

    (allMembers[crew] || []).forEach(function(m) {
      if (traineeOnly && (m.isManager || m.isStaff)) return;
      var s = submitted[m.name] || {};
      var status = s.status || "";
      var label  = m.name + (s.reason ? "(" + s.reason + ")" : "");
      if (m.isManager) {
        if (status === "참석") managers.push(m.name); else if (status === "불참") mgAbsent.push(label);
        else if (status === "부분참") mgLate.push(label); else mgUndecided.push(m.name);
      } else if (m.isStaff) {
        if (status === "참석") staffs.push(m.name); else if (status === "불참") stAbsent.push(label);
        else if (status === "부분참") stLate.push(label); else stUndecided.push(m.name);
      } else {
        if (status === "참석") trainees.push(m.name); else if (status === "불참") trAbsent.push(label);
        else if (status === "부분참") trLate.push(label); else undecided.push(m.name);
      }
    });

    var hasContent = managers.length || mgAbsent.length || mgLate.length || mgUndecided.length ||
                     staffs.length   || stAbsent.length  || stLate.length  || stUndecided.length ||
                     trainees.length || trAbsent.length  || trLate.length  || undecided.length;
    if (!hasContent) return;

    rSheet.getRange(r, 1).setValue("-" + crew).setFontWeight("bold").setBackground("#e8f4fd"); r++;
    if (managers.length > 0)    { rSheet.getRange(r,1).setValue(managers.join(" ")); r++; }
    if (mgLate.length > 0)      { rSheet.getRange(r,1).setValue("간사 부분참: "   + mgLate.join(" ")).setBackground("#fef3c7"); r++; }
    if (mgAbsent.length > 0)    { rSheet.getRange(r,1).setValue("간사 불참: "   + mgAbsent.join(" ")).setBackground("#fee2e2"); r++; }
    if (mgUndecided.length > 0) { rSheet.getRange(r,1).setValue("간사 미정: "   + mgUndecided.join(" ")).setBackground("#f3f4f6"); r++; }
    if (staffs.length > 0)      { rSheet.getRange(r,1).setValue("스태프: "      + staffs.join(" ")).setBackground("#f3e8ff"); r++; }
    if (stLate.length > 0)      { rSheet.getRange(r,1).setValue("스태프 부분참: " + stLate.join(" ")).setBackground("#fef3c7"); r++; }
    if (stAbsent.length > 0)    { rSheet.getRange(r,1).setValue("스태프 불참: " + stAbsent.join(" ")).setBackground("#fee2e2"); r++; }
    if (stUndecided.length > 0) { rSheet.getRange(r,1).setValue("스태프 미정: " + stUndecided.join(" ")).setBackground("#f3f4f6"); r++; }
    if (trainees.length > 0)    { rSheet.getRange(r,1).setValue("참석: "        + trainees.join(" ")); r++; }
    if (trLate.length > 0)      { rSheet.getRange(r,1).setValue("부분참: "        + trLate.join(" ")).setBackground("#fef3c7"); r++; }
    if (trAbsent.length > 0)    { rSheet.getRange(r,1).setValue("불참: "        + trAbsent.join(" ")).setBackground("#fee2e2"); r++; }
    if (undecided.length > 0)   { rSheet.getRange(r,1).setValue("미정: "        + undecided.join(" ")).setBackground("#f3f4f6"); r++; }
    r++;
  });
}

// ----------------------------------------------------------------
// 에러 로그 (관리 시트에 기록)
// ----------------------------------------------------------------
function writeErrorLog(errMsg, data) {
  try {
    var adminSS = getAdminSS();
    var logSheet = adminSS.getSheetByName("에러로그");
    if (!logSheet) {
      logSheet = adminSS.insertSheet("에러로그");
      logSheet.appendRow(["시간", "타입", "주차", "크루", "에러내용", "원본데이터"]);
      logSheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#ef4444").setFontColor("white");
      logSheet.setColumnWidth(1, 140);
      logSheet.setColumnWidth(2, 60);
      logSheet.setColumnWidth(3, 60);
      logSheet.setColumnWidth(4, 80);
      logSheet.setColumnWidth(5, 300);
      logSheet.setColumnWidth(6, 400);
    }
    var timestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    var type = data ? (data.type || "") : "";
    var week = data ? (data.week || "") : "";
    var crew = data ? (data.crew || "") : "";
    var raw  = data ? JSON.stringify(data) : "";
    logSheet.appendRow([timestamp, type, week, crew, errMsg, raw]);
    logSheet.getRange(logSheet.getLastRow(), 1, 1, 6).setBackground("#fff5f5");
  } catch(e) {
    // 로그 실패 시 무시
  }
}

// ----------------------------------------------------------------
// 공통 함수
// ----------------------------------------------------------------
function getWeekDate(adminSS, week) {
  var s = adminSS.getSheetByName("설정");
  if (!s || s.getLastRow() < 2) return "";
  var data = s.getRange(2, 1, s.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === week) return String(data[i][1]).trim();
  }
  return "";
}

// 명단 색상 적용 (에디터에서 수동 실행)
function applyMemberColors() {
  var sheet = getAdminSS().getSheetByName("명단");
  if (!sheet) { Logger.log("명단 시트 없음"); return; }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var bg   = data.map(function(r) { var v = String(r[2]).trim().toUpperCase(); return v==="Y"?["#dbeafe","#dbeafe","#dbeafe"]:v==="S"?["#ede9fe","#ede9fe","#ede9fe"]:["#ffffff","#ffffff","#ffffff"]; });
  var fc   = data.map(function(r) { var v = String(r[2]).trim().toUpperCase(); return v==="Y"?["#1d4ed8","#1d4ed8","#1d4ed8"]:v==="S"?["#6d28d9","#6d28d9","#6d28d9"]:["#000000","#000000","#000000"]; });
  var bold = data.map(function(r) { var b=(String(r[2]).trim().toUpperCase()==="Y"||String(r[2]).trim().toUpperCase()==="S")?"bold":"normal"; return [b,b,b]; });
  sheet.getRange(2, 1, lastRow-1, 3).setBackgrounds(bg).setFontColors(fc).setFontWeights(bold);
  sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#1a1d27").setFontColor("white");
  sheet.getRange(1, 4).setValue("Y:간사, S:스태프");
  Logger.log("완료! " + (lastRow-1) + "명 적용됨");
}