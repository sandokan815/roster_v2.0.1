/**
 * @description get date of each Monday
 * @param diff 
 * @returns date
 */
function getMonday(diff) {
  var now = new Date();
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var day = today.getDay();
  if (day == 0) {
    day = 7;
  }
  var monday = new Date(today.setDate(today.getDate() - day + diff));
  return monday;
}

/**
 * @description Detect current crew based on date
 * @param now 
 * @returns string (A_C or B_D)
 */
function detectCrew(now) {
  var first = new Date(now.getFullYear(), 0, 1);
  var weekNo = Math.ceil((((now - first) / 86400000) + first.getDay() + 1) / 7);
  if (now.getDay() == 0) {
    weekNo = weekNo - 1;
  }
  if (weekNo % 2 == 0) {
    return "A_C";
  } else {
    return "B_D";
  }
}

/**
 * @description Get the current week or next week 
 * @param diff 
 * @returns array
 */
function getWeek(diff) {
  var curr = new Date();
  var day = curr.getDay();
  if (day == 0) {
    day = 7;
  }
  var first = curr.getDate() - day + diff;
  var week = [];
  for (var i = 0; i < 7; i++) {
    var next = new Date(curr.getTime());
    next.setDate(first + i);
    next.setHours(0, 0, 0, 0);
    week.push(next);
  }
  return week;
}

/**
 * @description get data of A_C crew from sheet
 * @returns sheet
 */
function getDataA_C() {
  return SpreadsheetApp
    .openById('1-HmBSO0ViOWjC4ttaAxOZ1sygnfWmKvMJBswBRyouQA')
    .getSheetByName('Internal Dashboard A/C')
    .getRange(1, 1, 137, 14)
    .getValues();
}

/**
 * @description get data of B_D crew from sheet
 * @returns sheet
 */
function getDataB_D() {
  return SpreadsheetApp
    .openById('1-HmBSO0ViOWjC4ttaAxOZ1sygnfWmKvMJBswBRyouQA')
    .getSheetByName('Internal Dashboard B/D')
    .getRange(1, 1, 137, 14)
    .getValues();
}
/**
 * @description Get maximum length of all columns
 * @param data 
 * @param ref 
 */
function getMaxLength(data, ref) {
  var row;
  var maxLength;
  switch (ref) {
    case 'rollFirst':
      row = 8;
      break;
    case 'rollSecond':
      row = 27;
      break;
    case 'inFirst':
      row = 51;
      break;
    case 'inSecond':
      row = 65;
      break;
    case 'ecoFirst':
      row = 84;
      break;
    case 'ecoSecond':
      row = 96;
      break;
    case 'northFirst':
      row = 113;
      break;
    case 'northSecond':
      row = 126;
      break;
  }
  maxLength = Math.max(data[row-1][1], data[row-1][3], data[row-1][5], data[row-1][7], data[row-1][9], data[row-1][11], data[row-1][13]);
  return [row, maxLength];
}

/**
 * @description make table-view historical data from Internal Dashboard sheet
 * @param data 
 * @param ref 
 */
function makeTableFromInternal(data, ref) {
  // get maxLength
  var row_length = getMaxLength(data, ref);
  var rowNum = row_length[0];
  var maxLength = row_length[1];
  
  // define table maxLength * 7
  var table = [];
  for (var i = 0; i < maxLength; i++) {
    table[i] = [];
  }
  // insert data from sheet to table
  for (var i = 0; i < maxLength; i++) {
    for (var j = 0; j < 7; j++) {
      if (data[i+rowNum][2*j+1] && data[i+rowNum][2*j+1] != '-') {
        table[i][j] = data[i+rowNum][2*j+1];
      } else {
        table[i][j] = null;
      }
    }
  }
  // remove null in table
  var columns = [];
  for (var i = 0; i < 7; i++) {
    columns[i] = [];
    for (j = 0; j < table.length; j++) {
      if (table[j][i]) columns[i].push(table[j][i]);
    }
  }
  var new_table = [];
  for (var i = 0; i < table.length; i++) {
    new_table[i] = [];
    for (j = 0; j < 7; j++) {
      if (columns[j][i]) new_table[i][j] = columns[j][i];
      else new_table[i][j] = null;
    }
  }
  return new_table;
}

/**
 * @description Main Function
 * @returns It will download historical data from company shedule sheet to another sheet
 */
function downloadData() {
  var spreadsheetId = '1AvHhKOEsYRUHWqsG6-k8_C-NQ8rcaIqdAbkwnBAoxfg'; // spreadsheet ID
  var sheetName = Utilities.formatDate(getMonday(1), "CST", "MM/dd/YY"); // Create Sheet Name as Monday-Date
  var activeSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var newSheet = activeSpreadsheet.getSheetByName(sheetName);

  var crew = detectCrew(new Date()); // Detect Current Crew based on date
  var week = getWeek(1);
  var days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  var dates = [];
  for (var i in week) {
    dates.push(Utilities.formatDate(week[i], "CST", "MM/dd/YY"))
  }
  if (crew == "A_C") {
    var currentData = getDataA_C();
    var rollFirstHead = ['A Crew', 'A Crew', 'B Crew', 'B Crew', 'A Crew', 'A Crew', 'A Crew'];
    var rollSecondHead = ['C Crew', 'C Crew', 'D Crew', 'D Crew', 'C Crew', 'C Crew', 'C Crew'];
    var ecoFirstHead = ['Days', 'Days', 'Days', 'Days', 'OFF', 'OFF', 'OFF'];
    var ecoSecondHead = ['Nights', 'Nights', 'Nights', 'Nights', 'OFF', 'OFF', 'OFF'];
  }
  if (crew == "B_D") {
    var currentData = getDataB_D();
    var rollFirstHead = ['B Crew', 'B Crew', 'A Crew', 'A Crew', 'B Crew', 'B Crew', 'B Crew'];
    var rollSecondHead = ['D Crew', 'D Crew', 'C Crew', 'C Crew', 'D Crew', 'D Crew', 'D Crew'];
    var ecoFirstHead = ['Days', 'Days', 'Days', 'OFF', 'OFF', 'OFF', 'OFF'];
    var ecoSecondHead = ['Nights', 'Nights', 'Nights', 'OFF', 'OFF', 'OFF', 'OFF'];
  }
  var northFirstHead = ['Days', 'Days', 'Days', 'Days', 'OFF', 'OFF', 'OFF'];
  var northSecondHead = ['Nights', 'Nights', 'Nights', 'Nights', 'OFF', 'OFF', 'OFF'];

  var rollFirst = makeTableFromInternal(currentData, 'rollFirst');
  var rollSecond = makeTableFromInternal(currentData, 'rollSecond');
  var inFirst = makeTableFromInternal(currentData, 'inFirst');
  var inSecond = makeTableFromInternal(currentData, 'inSecond');
  var ecoFirst = makeTableFromInternal(currentData, 'ecoFirst');
  var ecoSecond = makeTableFromInternal(currentData, 'ecoSecond');
  var northFirst = makeTableFromInternal(currentData, 'northFirst');
  var northSecond = makeTableFromInternal(currentData, 'northSecond');

  var rollFirstHeader = [];
  rollFirstHeader.push(days, dates, rollFirstHead);
  var rollSecondHeader = [];
  rollSecondHeader.push(rollSecondHead);
  var inFirstHeader = rollFirstHeader;
  var inSecondHeader = rollSecondHeader;
  var ecoFirstHeader = [];
  ecoFirstHeader.push(days, dates, ecoFirstHead);
  var ecoSecondHeader = [];
  ecoSecondHeader.push(ecoSecondHead);
  var northFirstHeader = [];
  northFirstHeader.push(days, dates, northFirstHead);
  var northSecondHeader = [];
  northSecondHeader.push(northSecondHead);

  // Fill Out historical data into new sheet
  if (newSheet == null) {
    newSheet = activeSpreadsheet.insertSheet();
    newSheet.setName(sheetName);
    //Roll Fed
    newSheet.getRange('A1').setValue('RollFed');
    newSheet.getRange('B2:H4').setValues(rollFirstHeader);
    var dataLength = rollFirst.length;
    if (rollFirst.length)
      newSheet.getRange('B5:H' + (4 + dataLength)).setValues(rollFirst);
    newSheet.getRange('B' + (6 + dataLength) + ':H' + (6 + dataLength)).setValues(rollSecondHeader);
    if (rollSecond.length)
      newSheet.getRange('B' + (7 + dataLength) + ':H' + (6 + dataLength + rollSecond.length)).setValues(rollSecond);
    dataLength += rollSecond.length;
    //Inline
    newSheet.getRange('A' + (7 + dataLength)).setValue('Inline');
    newSheet.getRange('B' + (8 + dataLength) + ':H' + (10 + dataLength)).setValues(inFirstHeader);
    if (inFirst.length)
      newSheet.getRange('B' + (11 + dataLength) + ':H' + (10 + dataLength + inFirst.length)).setValues(inFirst);
    dataLength += inFirst.length;
    newSheet.getRange('B' + (11 + dataLength) + ':H' + (11 + dataLength)).setValues(inSecondHeader);
    if (inSecond.length)
      newSheet.getRange('B' + (12 + dataLength) + ':H' + (11 + dataLength + inSecond.length)).setValues(inSecond);
    dataLength += inSecond.length;
    //Eco Star
    newSheet.getRange('A' + (12 + dataLength)).setValue('Eco star');
    newSheet.getRange('B' + (13 + dataLength) + ':H' + (15 + dataLength)).setValues(ecoFirstHeader);
    if (ecoFirst.length)
      newSheet.getRange('B' + (16 + dataLength) + ':H' + (15 + dataLength + ecoFirst.length)).setValues(ecoFirst);
    dataLength += ecoFirst.length;
    newSheet.getRange('B' + (16 + dataLength) + ':H' + (16 + dataLength)).setValues(ecoSecondHeader);
    if (ecoSecond.length)
      newSheet.getRange('B' + (17 + dataLength) + ':H' + (16 + dataLength + ecoSecond.length)).setValues(ecoSecond);
    dataLength += ecoSecond.length;
    //North Plant
    newSheet.getRange('A' + (17 + dataLength)).setValue('North Plant');
    newSheet.getRange('B' + (18 + dataLength) + ':H' + (20 + dataLength)).setValues(northFirstHeader);
    if (northFirst.length)
      newSheet.getRange('B' + (21 + dataLength) + ':H' + (20 + dataLength + northFirst.length)).setValues(northFirst);

    dataLength += northFirst.length;
    newSheet.getRange('B' + (21 + dataLength) + ':H' + (21 + dataLength)).setValues(northSecondHeader);
    if (northSecond.length)
      newSheet.getRange('B' + (22 + dataLength) + ':H' + (21 + dataLength + northSecond.length)).setValues(northSecond);
    dataLength += northSecond.length;
    return true;
  } else {
    return false;
  }
}