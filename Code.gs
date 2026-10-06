/**
 * Cherrybrook Raging Bulls 3-2-1 voting web app.
 * Bound to spreadsheet "Cherrybrook Raging Bulls Parents player".
 * Deploy: Deploy > New deployment > Web app > Execute as Me > Anyone > Deploy.
 */
var SHEET_ID = '1IoukghjpeDi3AQU5G7zOLjGU6MLu-tlJz9ep9l4rT_Y';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Cherrybrook Raging Bulls Voting')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  return SpreadsheetApp.openById(SHEET_ID);
}

function getPlayers() {
  var sheet = getSpreadsheet_().getSheetByName('Players');
  var values = sheet.getDataRange().getValues();
  var players = [];
  for (var i = 1; i < values.length; i++) {
    var shirt = values[i][0];
    var name = values[i][1];
    if (shirt === '' || name === '') continue;
    players.push({
      shirt: String(shirt),
      name: String(name),
      total: Number(values[i][2]) || 0,
      average: Number(values[i][3]) || 0,
      votes: Number(values[i][4]) || 0
    });
  }
  var match = getMatch_();
  return { players: players, matchNo: match.matchNo, opposition: match.opposition };
}

function getMatch_() {
  var report = getSpreadsheet_().getSheetByName('Match Report');
  var matchNo = 1;
  var opposition = '';
  if (report) {
    matchNo = report.getRange('B1').getValue() || 1;
    opposition = report.getRange('C1').getValue() || '';
  }
  return { matchNo: String(matchNo), opposition: String(opposition) };
}

function submitVotes(payload) {
  if (!payload || !payload.parentName) throw new Error('Parent name is required.');
  var first = String(payload.first || '');
  var second = String(payload.second || '');
  var third = String(payload.third || '');
  if (!first || !second || !third) throw new Error('Pick 1st, 2nd and 3rd.');
  if (first === second || first === third || second === third) {
    throw new Error('Each place must be a different player.');
  }
  var match = getMatch_();
  var matchNo = payload.matchNo || match.matchNo;
  var opposition = payload.opposition || match.opposition;
  var sheet = getSpreadsheet_().getSheetByName('Voters');
  if (!sheet) throw new Error('Voters sheet is missing.');
  var now = new Date();
  var rows = [
    [now, payload.parentName, first, 3, matchNo, opposition],
    [now, payload.parentName, second, 2, matchNo, opposition],
    [now, payload.parentName, third, 1, matchNo, opposition]
  ];
  sheet.getRange(sheet.getLastRow() + 1, 1, 3, 6).setValues(rows);
  return { ok: true, message: 'Votes saved for ' + payload.parentName + '.' };
}
