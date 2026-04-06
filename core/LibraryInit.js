var _COLONY_SS = null;

function init(options) {
  if (!options || !options.spreadsheetId) {
    throw new Error('init() requires { spreadsheetId: "..." }');
  }
  _COLONY_SS = SpreadsheetApp.openById(options.spreadsheetId);
  return true;
}

function getColonySpreadsheet_() {
  if (_COLONY_SS) return _COLONY_SS;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No spreadsheet context. Call init({ spreadsheetId: "..." }) first.');
  return ss;
}
