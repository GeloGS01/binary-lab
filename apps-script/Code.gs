/**
 * Binary Lab - Beat the Clock global Top 10 (Google Apps Script web app).
 *
 * Setup (once): make a new Google Sheet -> Extensions -> Apps Script -> paste this file ->
 * Deploy -> New deployment -> Web app, Execute as: Me, Who has access: Anyone -> copy the /exec URL
 * into SCORE_URL in index.html.
 *
 * Moderation: every post is one row in the "Scores" tab. Type anything (e.g. x) in the
 * "Hide" column to remove that row from the board, or just delete the row.
 */
var BOARDS = ["4-guide", "4-noguide", "8-guide", "8-noguide",   // Binary Lab: bits + place-value guide on/off
              "ict7-showdown"];                                  // ICT 7 Logic Showdown (logic-escape/showdown)
var CLASSES = ["8A", "8B", "7A", "7B", "7C"];                                      // keep in step with CLASSES in index.html
var MAX_SCORE = 200;                                             // anything higher in 60 s is not believable
var TOP_N = 10;
var SHEET_NAME = "Scores";

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["Time", "Board", "Nickname", "Class", "Score", "Hide (type x)"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Each player (nickname + class) appears once, with their highest score; ties go to whoever got it first.
function top_(board) {
  var rows = sheet_().getDataRange().getValues().slice(1);
  var best = {};
  rows.forEach(function (r) {
    if (r[1] !== board || String(r[5]).trim()) return;
    var name = String(r[2]), cls = String(r[3]), score = Number(r[4]) || 0, t = new Date(r[0]).getTime() || 0;
    var key = name.toLowerCase() + "|" + cls;
    var b = best[key];
    if (!b || score > b.score || (score === b.score && t < b.t)) best[key] = { name: name, cls: cls, score: score, t: t };
  });
  return Object.keys(best).map(function (k) { return best[k]; })
    .sort(function (a, b) { return b.score - a.score || a.t - b.t; })
    .slice(0, board === "ict7-showdown" ? 500 : TOP_N)   // Logic Showdown lists every player
    .map(function (e) { return { name: e.name, cls: e.cls, score: e.score }; });
}

function doGet(e) {
  var board = e && e.parameter && e.parameter.board;
  if (BOARDS.indexOf(board) < 0) return json_({ ok: false, error: "Unknown board" });
  return json_({ ok: true, top: top_(board) });
}

function doPost(e) {
  var d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: "Bad request" }); }
  var board = String(d.board || "");
  var cls = String(d.cls || "");
  var score = Math.floor(Number(d.score));
  // letters (incl. Thai), digits, space, _ . - only; must start with a letter/digit (also blocks =formulas)
  var name = String(d.name || "").replace(/[^A-Za-z0-9฀-๿ _.\-]/g, "")
    .replace(/^[^A-Za-z0-9฀-๿]+/, "").trim().slice(0, 12);

  if (BOARDS.indexOf(board) < 0) return json_({ ok: false, error: "Unknown board" });
  if (CLASSES.indexOf(cls) < 0) return json_({ ok: false, error: "Pick your class" });
  if (!name) return json_({ ok: false, error: "Type a nickname" });
  if (!(score >= 1 && score <= MAX_SCORE)) return json_({ ok: false, error: "Score not accepted" });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { sheet_().appendRow([new Date(), board, name, cls, score, ""]); }
  finally { lock.releaseLock(); }
  return json_({ ok: true, top: top_(board) });
}
