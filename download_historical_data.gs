function downloadData() {
  var spreadsheetId = '1AvHhKOEsYRUHWqsG6-k8_C-NQ8rcaIqdAbkwnBAoxfg'; // spreadsheet ID
  var sheetName = Utilities.formatDate(firstOfWeek(1), "CST", "MM/dd/YY")
  var activeSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var newSheet = activeSpreadsheet.getSheetByName(sheetName);

  var crew = detectCrew(new Date());
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