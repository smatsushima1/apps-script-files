

function getVariables2() {
  _ga = SpreadsheetApp.getActive();
  _gas = _ga.getActiveSheet();
  _frow = firstRow(_gas);
  _lrow = _gas.getLastRow();
  _lcol = _gas.getLastColumn();
  _today = new Date();
};


////////////////////////////////////////////////////////////////////////////////////////////////////
// Operating Functions
////////////////////////////////////////////////////////////////////////////////////////////////////


function onEdit(e) {
  var range = e.range;
  if (range.getColumn() == 13 || range.getColumn() == 16 || range.getColumn() == 19 || range.getColumn() == 21) {
    mainCheck(range.getRow());
  };
};


function onEditManual() {
  var range = SpreadsheetApp.getActive().getActiveRange();
  mainCheck(range.getRow());
};


/*
Column number: heading
0: 120 Day Notice By
1: 120 Day Notice Sent
2: Initiate Interim CPARS
3: 60 Day Notice to KTR By
4: 60 Day Notice to KTR Sent
5: Days to Exercise Option by Written Notice
6: Option Exercised By
7: Option Exercised
8: Follow-On Due By
9: Follow-On Awarded
*/
function mainCheck(row) {
  getVariables2();
  var data = _gas.getRange("L" + row + ":U" + row).getValues();
  var days_ms = 1000*60*60*24;
  var tdate_ms = _today.getTime();
  // Exit if no 120-day notice column and follow-on column are empty
  if (data[0][0] == "" && data[0][8] == "") {
    return;
  // Mark row as complete if option has been exercised and no follow-on required; check if new option is required
  } else if (data[0][7] != "" && data[0][8] == "") {
    if (tdate_ms >= new Date(data[0][6]).getTime()) {
      updateColorStatus(row, "white", "");
      createDate(row);   
      return;
    } else {
      updateColorStatus(row, "white", "");
      return;
    };
  // If follow-on has been awarded, order or contract is ready for closout
  } else if (data[0][9] != "" && tdate_ms >= new Date(data[0][8]).getTime()) {
    var con_ord = _gas.getRange(row, 1, 1, 2).getValues();
    // First check if the action is a contract
    if (con_ord[0][1] == "" || con_ord[0][1].trim().toLowerCase() == "n/a") {
      var contracts = _gas.getRange("A" + _frow + ":A" + _lrow).getValues();
      // If there are orders still present, then don't highlight blue, otherwise, highlight blue
      if (contracts.filter(x => x==con_ord[0][0]).length > 1) {
        updateColorStatus(row, "white", "");
        return;
      } else {
        updateColorStatus(row, "blue", "");
        return;
      };
    // All orders
    } else {
      updateColorStatus(row, "blue", "");
      return;
    };
  // If follow-on has been awarded, but order is still not ready for closeout
  } else if (data[0][9] != "" && tdate_ms < new Date(data[0][8]).getTime()) {
    updateColorStatus(row, "white", "");
    return;
  };
  // Check for all other dates
  var day_120_diff = Math.round((new Date(data[0][0]).getTime() - tdate_ms) / days_ms) + 1;
  var day_60_diff = Math.round((new Date(data[0][3]).getTime() - tdate_ms) / days_ms) + 1;
  var day_ex_diff = Math.round((new Date(data[0][6]).getTime() - tdate_ms) / days_ms) + 1;
  if (data[0][8] == "") {
    var day_fo_diff = 999;
  } else {
    var day_fo_diff = Math.round((new Date(data[0][8]).getTime() - tdate_ms) / days_ms) + 1;
  };
  // 120 day notice to client
  if (data[0][1] == "") {
    checkDays(row, day_120_diff, "120-day notice to client");
  // 60 day notice to contractor
  } else if (data[0][4] == "") {
    checkDays(row, day_60_diff, "60-day notice to contractor");
  // Exercise option notice
  } else if (data[0][7] == "") {
    checkDays(row, day_ex_diff, "Exercise option");
  // Follow-on notice
  } else if (data[0][8] != "" && data[0][9] == "") {
    checkDays(row, day_fo_diff, "Follow-on");
  // All other actions
  } else {
    checkDays(row, 999, "");
  };
};


function checkDays(row, day_diff, message) {
  // Non-follow-on items
  if (message.toLowerCase().indexOf("follow-on") == -1) {
    if (day_diff <= 60 && day_diff > 30) {
      updateColorStatus(row, "yellow", message + " due in " + day_diff + " days");
    } else if (day_diff <= 30 && day_diff > 0) {
      updateColorStatus(row, "orange", message + " due in " + day_diff + " days");
    } else if (day_diff <= 0) {
      updateColorStatus(row, "red", message + " past due");
    } else {
      updateColorStatus(row, "white", "");
    };
  // Follow-on items
  } else {
    if (day_diff <= 60 && day_diff > 0) {
      updateColorStatus(row, "purple", message + " due in " + day_diff + " days");
    } else if (day_diff <= 0) {
      updateColorStatus(row, "purple", message + " past due");
    } else {
      updateColorStatus(row, "white", "");
    };
  };
};


function updateColorStatus(row, color, status_message) {
  var hl = "#ffffff";
  if (color == "blue") {
    hl = "#c9daf8";
  } else if (color == "gray") {
    hl = "#d9d9d9";
  } else if (color == "yellow") {
    hl = "#fff2cc";
  } else if (color == "orange") {
    hl = "#fce5cd";
  } else if (color == "red") {
    hl = "#f4cccc";
  } else if (color == "purple") {
    hl = "#d9d2e9";
  };
  _gas.getRange(row, 1, 1, _lcol).setBackground(hl);
  // Highlight applicable cells
  var row_vals = _gas.getRange("A" + row + ":U" + row).getValues();
  // 13, 16, 18, 20
  if (row_vals[0][11] != "") {
    if (row_vals[0][12] == "") {
      _gas.getRange(row, 13).setBackground("#ffff00");
    } else {
      _gas.getRange(row, 13).setBackground(hl);
    };
    if (row_vals[0][15] == "") {
      _gas.getRange(row, 16).setBackground("#ffff00");
    } else {
      _gas.getRange(row, 16).setBackground(hl);
    };
    if (row_vals[0][18] == "") {
      _gas.getRange(row, 19).setBackground("#ffff00");
    } else {
      _gas.getRange(row, 19).setBackground(hl);
    };
  };
  // Follow-on dates go on all rows, if applicable
  if (row_vals[0][19] != "") {
    if (row_vals[0][20] == "") {
      _gas.getRange(row, 21).setBackground("#ffff00");
    } else {
      _gas.getRange(row, 21).setBackground(hl);
    };
  };  
  _gas.getRange(row, 4).setValue(status_message);
};


function updateDaily() {
  var st = Date.now();
  getVariables2();
  var sheets = _ga.getSheets();
  // Pull all data for report
  var data_list = [];
  for (var i = 0; i < sheets.length; i++) {
    var sh_name = sheets[i].getName();
    if (sh_name.indexOf("Closeout") == -1 && sh_name.indexOf("UEI") == -1 && sh_name.indexOf("Teams") == -1) {
      var csheet = _ga.getSheetByName(sh_name);
      csheet.activate();
      console.log("Working: " + csheet.getName());
      updateStatuses();
      // For testing:
      //generateEmailData().forEach(i => data_list.push(i));
      // For production:
      if (_today.getDay() == 1) generateEmailData().forEach(i => data_list.push(i));
    };
  };
  if (_today.getDay() == 1) generateEmail(data_list);
  endTime(st);
};


// To be ran daily in the morning
function updateStatuses() {
  for (var i = _frow; i <= _lrow; i++) {
    mainCheck(i);
  };
};


function generateEmailData() {
  var data = _gas.getRange(_frow, 1, _lrow - _frow + 1, _lcol - 1).getValues();
  var data_list = [];
  for (var i = 0; i < data.length; i++) {
    var st = data[i][3];
    if (data[i][3] != "") {
      // Identify due date
      var due = "";
      if (st.toLowerCase().indexOf("client") > -1) {
        due = prettyDate(data[i][11]);
      } else if (st.toLowerCase().indexOf("contractor") > -1) {
        due = prettyDate(data[i][14]);
      } else if (st.toLowerCase().indexOf("exercise") > -1) {
        due = prettyDate(data[i][17]);
      } else if (st.toLowerCase().indexOf("follow-on") > -1) {
        due = prettyDate(data[i][19]);
      };
      // Parse out amount of days and color to add; by default, it's set to red
      // red = "#f4cccc"
      // yellow = "#fff2cc"
      // purple = "d9d2e9"
      var color = "#f4cccc";
      if (st.indexOf("past due") > -1) {
        color = "#f4cccc";
      } else if (st.trim().substring(st.length - 7, st.length).trim().split(" ")[0] > 30) {
        color = "#fff2cc";
      } else if (st.trim().substring(st.length - 7, st.length).trim().split(" ")[0] > 0) {
        color = "#fce5cd";
      } else if (st.indexOf("Follow-on") > -1) {
        color = "#d9d2e9";
      };
      // Add in Team Lead
      var sh_teams = _ga.getSheetByName("Teams");
      var t_list = sh_teams.getRange(1, 1, sh_teams.getLastRow(), sh_teams.getLastColumn()).getValues();
      for (var j = 0; j < t_list.length; j++) {
        if (_gas.getName() == t_list[j][0]) {
          var t_lead = t_list[j][1];
        };
      };
      data_list.push([_gas.getName(), data[i][0], data[i][1], data[i][2], findOption(data[i][10]), st, due, t_lead, color]);
    };
  };
  return data_list;
};


function addDates() {
  getVariables2();
  createDate(_gas.getActiveCell().getRow());
  onEditManual();
};


function findOption(date) {
  var new_pop = [];
  date.split("\n").forEach(i => new_pop.push(i.trim().replace(/\s/g, "").split(":")));
  for (var i = 0; i < new_pop.length; i++) {
    var oy = new_pop[i][0].toLowerCase();
    if(_today.getTime() < new Date(new_pop[i][1].split("-")[0]).getTime()) {
      break;
    };
  };
  return oy.replace("oy", "Option ");
};


function generateEmail(list) {
  // Separate into teams
  var sh_teams = _ga.getSheetByName("Teams");
  var t_list = sh_teams.getRange(1, 1, sh_teams.getLastRow(), sh_teams.getLastColumn()).getValues();
  var t_leads = getUniqueValues(t_list, 1);
  var all_teams = [];
  for (i = 0; i < t_leads.length; i++) {
    all_teams[i] = [];
    all_teams[i].push(["CS", "Contract", "Order", "Description", "Option", "Status", "Due Date", "Team", "Color"]);
    for (var j = 0; j < list.length; j++) {
      if (list[j][7] == t_leads[i]) {
        all_teams[i].push(list[j]);
      };
    };
  };
  // Email generation
  // Body
  var body = "Good Morning,<br><br>";
  body += "The below requirements were identified in the Status Report to have an upcoming action.<br><br>";
  // Create tables for each Team Lead
  for (i = 0; i < t_leads.length; i ++) {
    body += "<b>" + t_leads[i] + "'s Team:</b><br>"
    if (all_teams[i].length == 1) {
      body += "No actions to display.<br><br>";
    } else {
      body += generateEmailTable(all_teams[i]) + "<br>";
    };
  };
  body += "Please consult the Status Report for more details. Have a good day!<br><br>"
  // Signature
  var signature = "<font color='darkgray';>--<br>";
  signature += parseFullName() + "<br>";
  signature += "NOAA, AGO<br>";
  signature += "Eastern Acquisition Division</font>";
  // Draft Email
  GmailApp.createDraft("heather.l.coleman@noaa.gov, dorothy.curling@noaa.gov, nicole.lawson@noaa.gov",
                       "Status Report Upcoming Actions - " + Utilities.formatDate(new Date(), "EST", "M/dd"),
                       "",
                       {cc: "stacy.dohse@noaa.gov",
                        htmlBody: body + signature
  });
};


function generateEmailTable(list) {
  var table_html = "<table style='border:1px solid black;border-collapse:collapse; '>";
  // First row
  table_html += "<tr>";
  for (var i = 0; i < list.length; i++) {
    table_html += "<tr>";
    /*
    CS
    Contract
    Order
    Description
    Option
    Status
    Due Date
    */
    if (i == 0) {
      table_html += ths(list[i][0], 1);
      table_html += ths(list[i][1], 1);
      table_html += ths(list[i][2], 1);
      table_html += ths(list[i][3], 200);
      table_html += ths(list[i][4], 60);
      table_html += ths(list[i][5], 150);
      table_html += ths(list[i][6], 1);
    } else {
      table_html += tds(list[i][0], 1, list[i][8]);
      table_html += tds(list[i][1], 1, list[i][8]);
      table_html += tds(list[i][2], 1, list[i][8]);
      table_html += tds(list[i][3], 200, list[i][8]);
      table_html += tds(list[i][4], 60, list[i][8]);
      table_html += tds(list[i][5], 150, list[i][8]);
      table_html += tds(list[i][6], 1, list[i][8]);
    };
    table_html += "</tr>";
  };
  table_html += "</table>";
  return table_html;
};


function ths(text, width) {
  return "<th style='border:1px solid black;padding:5px;width:" + width + "px';>" + text + "</th>";
};


function tds(text, width, color) {
  return "<td style='border:1px solid black;padding:5px;width:" + width + "px;background-color:" + color + ";'>" + text + "</td>";
};


function prettyDate(date) {
  var n_date = new Date(date);
  return (n_date.getMonth() + 1) + "/" + n_date.getDate() + "/" + (n_date.getYear() + 1900);
};


function addDatesCalendar() {
  getVariables2();
  var arow = _gas.getActiveCell().getRow();
  // Error check
  if (arow <= _frow || arow > _lrow) {
    Browser.msgBox("Row selected was out of bounds. Exiting...");
    return;
  };
  var data = _gas.getRange(arow, 1, 1, 20).getValues();
  var contract = (data[0][0] + " " + data[0][1]).trim();
  var ca = CalendarApp;
  if (data[0][11] != "") ca.createAllDayEvent("120-Day Notice due for " + contract, data[0][11]);
  if (data[0][14] != "") ca.createAllDayEvent("60-Day Notice due for " + contract, data[0][14]);
  if (data[0][17] != "") ca.createAllDayEvent("Exercise Option due for " + contract, data[0][17]);
  if (data[0][19] != "") ca.createAllDayEvent("Follow-On due for " + contract, data[0][19]);
  Browser.msgBox("Dates added. Exiting...");
};


// Get font color
function dev101() {
  getVariables2();
  // Font color
  console.log(_gas.getActiveCell().getTextStyle().getForegroundColor());
  // Background color
  console.log(_gas.getActiveCell().getBackgroundColor());
};


// Dev report run
function dev102() {
  getVariables2();
  var data_list = [ [ 'Other Person 1',
    'EA133C17BA0049',
    '1305M223FNCNS0207',
    'Business Management Division (BMD) Labs Call Order',
    'Option 1',
    'Exercise option due in 13 days',
    '10/15/2024',
    'Nicole',
    '#f4cccc' ],
  [ 'Other Person 1',
    '1305M219ANCNS0013',
    '',
    'Analytical Chemistry Services',
    'Option 5',
    '120-day notice to client due in 23 days',
    '10/25/2024',
    'Nicole',
    '#f4cccc' ],
  [ 'Other Person 1',
    'GS00F217CA',
    '1305M224F0378',
    'Marine Spatial Ecology (MSE) HQ Bridge Order',
    'Option 1',
    '120-day notice to client due in 31 days',
    '11/2/2024',
    'Nicole',
    '#fff2cc' ],
  [ 'Brian',
    'EA133C17BA0054',
    '1305M222FNCNP0232',
    'Watershed Coordination, Protection, and Restoration to Control Land Based Pollution for the Government Cut-Inlet in Southeast Florida',
    'Base',
    'Follow-on due in 29 days',
    '10/31/2024',
    'Dorothy',
    '#d9d2e9' ],
  [ 'Other Person 2',
    'EA133C17BA0049',
    '1305M223FNCNS0094',
    'Natural Resource Damage Assessment (NRDA) Labs Call Order ',
    'Option 1',
    'Follow-on due in 20 days',
    '10/22/2024',
    'Dorothy',
    '#d9d2e9' ],
  [ 'Other Person 2',
    'EA133C17BA0049',
    '1305M223FNCNS0207',
    'Business Management Division (BMD) Labs Call Order',
    'Option 1',
    'Exercise option past due',
    '9/15/2024',
    'Dorothy',
    '#f4cccc' ],
  [ 'Other Person 2',
    '1305M219ANCNS0013',
    '',
    'Analytical Chemistry Services',
    'Option 5',
    '120-day notice to client due in 23 days',
    '10/25/2024',
    'Dorothy',
    '#f4cccc' ],
  [ 'Other Person 2',
    'GS00F217CA',
    '1305M224F0378',
    'Stressor, Detection, and Impacts (SDI) HQ Call Order',
    'Option 1',
    '60-day notice to contractor due in 43 days',
    '11/14/2024',
    'Dorothy',
    '#fff2cc' ],
  [ 'Erik',
    '1305M220ANCNT0062',
    '1305M224F0192',
    'Sutron Order',
    'Base',
    'Follow-on past due',
    '8/11/2024',
    'Heather',
    '#f4cccc' ],
  [ 'Spencer',
    'GS00F217CA',
    '1305M224F0367',
    'Marine Spatial Ecology (MSE) HQ Bridge Order',
    'Option 1',
    '120-day notice to client due in 31 days',
    '11/2/2024',
    'Heather',
    '#fff2cc' ] ];
  generateEmail(data_list);
};


////////////////////////////////////////////////////////////////////////////////////////////////////
// Conversion Functions
////////////////////////////////////////////////////////////////////////////////////////////////////


/*
Before running functions:
1) Unmerge all column headings
2) Add "Status" at column D
3) Move "Excepted Category" and "Warranty Expiration" to end
4) Delete unneeded columns
5) Add "Acquisition Workspace" column after "M: Drive Location" column
7) Add seven columns to the left of column "L"
8) Rename all columns from "L" to "U" to the column headings on Spencer's sheet
9) Convert columns "L" to "P" and "R" to "U" to date format
10) Make sure all "Closeout" tabs have the word "Closeout" in the title
11) Substitute – with - in the POP if results aren't showing correctly
*/


function xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx() {
};


function createDate() {
  getVariables2();
  var row = _gas.getActiveCell().getRow();
  var data = _gas.getRange("K" + row + ":U" + row).getValues();
  var dates = findPOP(data[0][0]);
  // First add follow-on date if it's just a regular POP with no options
  if (data[0][0].trim().split("\n").length == 1) {
    data[0][9] = data[0][0].trim().split("\n")[0].replace("Base:", "").trim().replace(/\s/g, "").split("-")[1];
  // For all else, check options
  } else if (_today.getTime() < new Date(dates[1]).getTime()) {
    var st = new Date(dates[1]).getTime();
    var days_ms = 1000*60*60*24;
    // 120-day notice
    data[0][1] = Utilities.formatDate(new Date(st - (119 * days_ms)), "EST", "M/d/yyyy");
    data[0][2] = "";
    data[0][3] = "";
    // 60-day notice
    data[0][4] = Utilities.formatDate(new Date(st - (59 * days_ms)), "EST", "M/d/yyyy");
    data[0][5] = "";
    // Exercise option
    data[0][7] = Utilities.formatDate(new Date(st - ((data[0][6] - 1) * days_ms)), "EST", "M/d/yyyy");
    data[0][8] = "";
    // Follow-on, if applicable
    if (dates[3] == 1) {
      data[0][9] = dates[2];
    };
  };
  _gas.getRange(row, 11, 1, data[0].length).setValues(data);
};


function formatPOP() {
  getVariables2();
  var data = _gas.getRange(_frow, 11, _lrow - _frow + 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == "") continue;
    var new_pop = "";
    var split_pop = data[i][0].trim().split("\n");
    for (var j = 0; j < split_pop.length; j++) {
      if (j == 0) {
        new_pop = "Base: " + split_pop[j] + "\n";
      } else {
        new_pop += "OY" + j + ": " + split_pop[j] + "\n";
      };
    };
    data[i][0] = new_pop.trim();
  };
  console.log(data);
  _gas.getRange(_frow, 11, _lrow - _frow + 1, 1).setValues(data);
};


function createDatesDev() {
  createDates(0);
};


function createDatesProd() {
  createDates(1);
};


function createDates(test) {
  getVariables2();
  var data = _gas.getRange("K" + _frow + ":U" + _lrow).getValues();
  for (var i = 0; i < data.length; i++) {
    console.log(i);
    // Exit if empty row
    if (data[i][0] == "") continue;
    // Inputting just follow-on date for no option periods
    if (data[i][0].trim().split("\n").length == 1) {
      data[i][9] = data[i][0].trim().split("\n")[0].replace("Base:", "").trim().replace(/\s/g, "").split("-")[1];
    // All other contracts with options
    } else {
      var pop = findPOP(data[i][0].trim());
      var st = new Date(pop[1]).getTime();
      var days_ms = 1000*60*60*24;
      // 120-day notice
      data[i][1] = Utilities.formatDate(new Date(st - (119 * days_ms)), "EST", "M/d/yyyy");
      data[i][2] = "";
      data[i][3] = "";
      // 60-day notice
      data[i][4] = Utilities.formatDate(new Date(st - (59 * days_ms)), "EST", "M/d/yyyy");
      data[i][5] = "";
      data[i][6] = "";
      // Exercise option
      data[i][7] = Utilities.formatDate(new Date(st - ((data[i][6] - 1) * days_ms)), "EST", "M/d/yyyy");
      data[i][8] = "";
      // Follow-on, if applicable
      if (pop[3] == 1) {
        data[i][9] = pop[2];
      };
    };
  };
  if (test == 0) {
    console.log(data);
    console.log(data.length);
  } else {
    // Data must be a list within a list    
    _gas.getRange(_frow, 11, data.length, data[0].length).setValues(data);
  };
};


function findPOP(date) {
  var new_pop = [];
  date.split("\n").forEach(i => new_pop.push(i.trim().replace(/\s/g, "").split(":")));
  final_pop = [];
  for (var i = 0; i < new_pop.length; i++) {
    var pop = new_pop[i][1].split("-");
    final_pop.push([new_pop[i][0], pop[0], pop[1]]);
  };
  for (var i = 0; i < final_pop.length; i++) {
    var oy = final_pop[i][0];
    var s_date = final_pop[i][1].trim();
    var e_date = final_pop[i][2].trim();
    // If this is the last iteration of i, save as a follow-on
    if (i != final_pop.length - 1) {
      var fo_ind = 0;
    } else {
      var fo_ind = 1;
    };
    if (_today.getTime() < new Date(s_date).getTime()) {
      break;
    };
  };
  return [oy.toLowerCase().replace("oy", "Option "), s_date, e_date, fo_ind];
};


function convertToCompleted() {
  getVariables2();
  var data = _gas.getRange(_frow, 12, _lrow - _frow + 1, 10).getValues();
  var tdate = _today.getTime();
  for (var i = 0; i < data.length; i++) {
    var cust_date = (new Date(data[i][0])).getTime();
    var ktr_date = (new Date(data[i][3])).getTime();
    var ex_date = (new Date(data[i][6])).getTime();
    var fo_date = (new Date(data[i][8])).getTime();
    // Skip item
    if (data[i][0] == "" && data[i][8] == "") {
      continue;
    };
    // 120 day notice dates
    if (tdate > cust_date && data[i][1] == "") {
      data[i][1] = "-";
    };
    // KTR notice dates
    if (tdate > ktr_date && data[i][4] == "") {
      data[i][4] = "-";
    };
    // Exercised option dates
    if (tdate > ex_date && data[i][7] == "") {
      data[i][7] = "-";
    };
    // Follow-on dates
    if (tdate > fo_date && data[i][8] != "" && data[i][9] == "") {
      data[i][9] = "-";
    };
  };
  // Insert data
  //console.log(data);
  _gas.getRange(_frow, 12, data.length, data[0].length).setValues(data);
};


function updateStatusesSheet() {
  getVariables2();
  updateStatuses();
};


////////////////////////////////////////////////////////////////////////////////////////////////////
// Accessory Functions
////////////////////////////////////////////////////////////////////////////////////////////////////


function xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx() {
};


function firstRow(sheet) {
  var list = sheet.getRange("A1:A10").getValues();
  var frow = "";
  for (var i = 0; i < list.length; i++) {
    if (list[i][0] == "Contract Number") {
      frow = i + 2;
    };
  };
  return frow;
};


function columnLetters(column) {
  return _gas.getRange(1, column).getA1Notation().slice(0, -1);
};


function endTime(start) {
  var total_ms = Date.now() - start;
  var s_ms = (total_ms/1000).toString().split('.');
  if (s_ms[0] == 0) {
    console.log('Finished in: 0m 0.' + s_ms[1].substring(0,3) + 's');
  } else {
    var m_s = (s_ms[0]/60).toString().split('.');
    var final_s = (Number('.' + m_s[1])*60).toString().substring(0,3);
    var final_m = m_s[0];
    console.log('Finished in: ' + final_m + 'm ' + final_s.replace('.', '') + 's');
  };
};


function getUniqueValues(data_list, column_number) {
  var rng_list = data_list.map(x => x[column_number]);
  var noduplicates = new Set(rng_list);
  var unique_values = [];
  noduplicates.forEach(x => unique_values.push(x));
  // Remove first value, which normally would be the column header
  var first_value = unique_values[0];
  var final_list = unique_values.filter((value) => value != first_value);
  return final_list.sort();
};


function parseFullName() {
  var un = Session.getActiveUser().getUsername();
  var un2 = un.split(".");
  var nlength = un2.length;
  // First name
  var fname = un2[0];
  var fname_final = fname.substring(0,1).toUpperCase() + fname.substring(1, fname.length);
  // Last name
  var lname = un2[nlength - 1];
  var lname_final = lname.substring(0,1).toUpperCase() + lname.substring(1, lname.length);
  return fname_final + " " + lname_final;
};


////////////////////////////////////////////////////////////////////////////////////////////////////
// Deprecated Code
////////////////////////////////////////////////////////////////////////////////////////////////////


/*
function openForm() {
  var form = HtmlService.createHtmlOutputFromFile("add_contract");
  //SpreadsheetApp.getUi().showModalDialog(form, "Add Contract");
  SpreadsheetApp.getUi().showSidebar(form);
};


function addItem(data) {
  getVariables2();
  _gas.getRange(_gas.getLastRow() + 1, 1, 1, data[0].length).setValues(data);
};


function openOptions() {
  var form = HtmlService.createHtmlOutputFromFile("add_options");
  SpreadsheetApp.getUi().showSidebar(form);
};


function addOptions(years) {
  getVariables2();
  var a_row = _gas.getActiveCell().getRow();
  _gas.insertRowsAfter(a_row, years);
  var contract = _gas.getRange(a_row, 1).getValue();
  var order = _gas.getRange(a_row, 2).getValue();
  var options = [];
  for (var i = 1; i <= years; i++) {
    options.push([contract, order, "Option " + i]);
  };
  _gas.getRange(a_row + 1, 1, years, options[0].length).setValues(options);
};


function mainCheckOld(row) {
  getVariables2();
  var data = _gas.getRange(row, 3, 1, 20).getValues();
  // First check if row belongs to an option; exit if it isn't
  if (data[0][0].toLowerCase().indexOf("option") == -1) {
    return;
  // Then check if option is exercised; update and exit if it is
  } else if ((data[0][17] != "" && data[0][18] == "") || (data[0][17] != "" && data[0][18] != "" && data[0][19] != "")) {
    updateColorStatus(row, "green", "");
    return;
  };
  // Check for all other dates
  var days_ms = 1000*60*60*24;
  var tdate_ms = _today.getTime();
  var day_120_diff = Math.round((new Date(data[0][10]).getTime() - tdate_ms) / days_ms) + 1;
  var day_60_diff = Math.round((new Date(data[0][13]).getTime() - tdate_ms) / days_ms) + 1;
  var day_ex_diff = Math.round((new Date(data[0][16]).getTime() - tdate_ms) / days_ms) + 1;
  if (data[0][16] = "") {
    var day_fo_diff = 999;
  } else {
    var day_fo_diff = Math.round((new Date(data[0][18]).getTime() - tdate_ms) / days_ms) + 1;
  };
  var color = "white";
  var status = "";
  // 120 day notice to client
  if (data[0][11] == "") {
    // 60 days out
    if (day_120_diff <= 60 && day_120_diff > 30) {
      color = "yellow";
      status = "120 day notice to client due in " + day_120_diff + " days";
    // 30 days out
    } else if (day_120_diff <= 30 && day_120_diff > 0) {
      color = "red";
      status = "120 day notice to client due in " + day_120_diff + " days";
    };
  // 60 day notice to contractor
  } else if (data[0][14] == "") {
    // 60 days out
    if (day_60_diff <= 60 && day_60_diff > 30) {
      color = "yellow";
      status = "60 day notice to contractor due in " + day_60_diff + " days";
    // 30 days out
    } else if (day_60_diff <= 30 && day_60_diff > 0) {
      color = "red";
      status = "60 day notice to contractor due in " + day_60_diff + " days";
    };
  // Exercise option notice
  } else if (data[0][17] == "") {
    // 60 days out
    if (day_ex_diff <= 60 && day_ex_diff > 30) {
      color = "yellow";
      status = "Exercise option due in " + day_ex_diff + " days";
    // 30 days out
    } else if (day_ex_diff <= 30 && day_ex_diff > 0) {
      color = "red";
      status = "Exercise option due in " + day_ex_diff + " days";
    // Past due
    } else if (day_ex_diff <= 0) {
      color = "red";
      status = "Exercise option past due";
    };
  // Follow-on notice
  } else if (data[0][18] != "" && data[0][19] == "") {
    // 60 days out
    if (day_fo_diff <= 60 && day_fo_diff > 0) {
      color = "purple";
      status = "Follow-on due in " + day_fo_diff + " days";
    // Past due
    } else if (day_fo_diff <= 0) {
      color = "purple";
      status = "Follow-on past due";
    };
  };
  // Highlight row and update status
  updateColorStatus(row, color, status);
};
*/


/*
function mainCheck(row) {
  getVariables2();
  if (_gas.getRange(row, 3).getValue().toLowerCase().indexOf("option") == -1) {
    mainCheckBase(row);
  } else {
    mainCheckOption(row);
  };
};


function mainCheckBase(row) {
  var data = _gas.getRange(row, 3, 1, 20).getValues();
  // First check for empty spacer rows
  if (data[0][0] == "") {
    updateColorStatusBase(row, "black", "");
    return;
  // Then check for base rows
  } else if (data[0][18] == "" || (data[0][18] != "" && data[0][19] != "")) {
    updateColorStatusBase(row, "blue", "");
    return;
  };
  // Check for all other dates
  var days_ms = 1000*60*60*24;
  var tdate_ms = _today.getTime();
  if (data[0][18] == "") {
    var day_fo_diff = 999;
  } else {
    var day_fo_diff = Math.round((new Date(data[0][18]).getTime() - tdate_ms) / days_ms) + 1;
  };
  checkDaysBase(row, day_fo_diff, "Follow-on");
};


function checkDaysBase(row, day_diff, message) {
  // Exit if row is before the first row
  if (row < _frow) {
    return;
  };
  // Follow-on items
  if (day_diff <= 60 && day_diff > 0) {
    updateColorStatusBase(row, "purple", message + " due in " + day_diff + " days");
  } else if (day_diff <= 0) {
    updateColorStatusBase(row, "purple", message + " past due");
  } else {
    updateColorStatusBase(row, "blue", "");
  };
};


function updateColorStatusBase(row, color, status_message) {
  if (color == "blue") {
    hl = "#c9daf8";
  } else if (color == "purple") {
    hl = "#d9d2e9";
  } else if (color == "black") {
    hl = "#000000";
  };
  _gas.getRange(row, 1, 1, _lcol).setBackground(hl);
  // Follon-dates go on all rows, if applicable
  if (color != "black" && _gas.getRange(row, 21).getValue() != "") {
    _gas.getRange(row, 22).setBackground("#ffff00");
  };
  _gas.getRange(row, 4).setValue(status_message);
};


function mainCheckOption(row) {
  var data = _gas.getRange(row, 3, 1, 20).getValues();
  // Mark row as complete if option has been exercised and no follow-on required, or if everything has been fulfilled
  if ((data[0][17] != "" && data[0][18] == "") || (data[0][17] != "" && data[0][18] != "" && data[0][19] != "")) {
    updateColorStatusOption(row, "green", "");
    return;
  };
  // Check for all other dates
  var days_ms = 1000*60*60*24;
  var tdate_ms = _today.getTime();
  var day_120_diff = Math.round((new Date(data[0][10]).getTime() - tdate_ms) / days_ms) + 1;
  var day_60_diff = Math.round((new Date(data[0][13]).getTime() - tdate_ms) / days_ms) + 1;
  var day_ex_diff = Math.round((new Date(data[0][16]).getTime() - tdate_ms) / days_ms) + 1;
  if (data[0][18] == "") {
    var day_fo_diff = 999;
  } else {
    var day_fo_diff = Math.round((new Date(data[0][18]).getTime() - tdate_ms) / days_ms) + 1;
  };
  // 120 day notice to client
  if (data[0][11] == "") {
    checkDaysOption(row, day_120_diff, "120-day notice to client");
  // 60 day notice to contractor
  } else if (data[0][14] == "") {
    checkDaysOption(row, day_60_diff, "60-day notice to contractor");
  // Exercise option notice
  } else if (data[0][17] == "") {
    checkDaysOption(row, day_ex_diff, "Exercise option");
  // Follow-on notice
  } else if (data[0][18] != "" && data[0][19] == "") {
    checkDaysOption(row, day_fo_diff, "Follow-on");
  // All other actions
  } else {
    checkDaysOption(row, 999, "");
  };
};


function checkDaysOption(row, day_diff, message) {
  // Non-follow-on items
  if (message.toLowerCase().indexOf("follow-on") == -1) {
    if (day_diff <= 60 && day_diff > 30) {
      updateColorStatusOption(row, "yellow", message + " due in " + day_diff + " days");
    } else if (day_diff <= 30 && day_diff > 0) {
      updateColorStatusOption(row, "red", message + " due in " + day_diff + " days");
    } else if (day_diff <= 0) {
      updateColorStatusOption(row, "red", message + " past due");
    } else {
      updateColorStatusOption(row, "white", "");
    };
  // Follow-on items
  } else {
    if (day_diff <= 60 && day_diff > 0) {
      updateColorStatusOption(row, "purple", message + " due in " + day_diff + " days");
    } else if (day_diff <= 0) {
      updateColorStatusOption(row, "purple", message + " past due");
    } else {
      updateColorStatusOption(row, "white", "");
    };
  };
};


function updateColorStatusOption(row, color, status_message) {
  var hl = "#ffffff";
  if (color == "green") {
    hl = "#d9ead3";
  } else if (color == "yellow") {
    hl = "#fff2cc";
  } else if (color == "red") {
    hl = "#f4cccc";
  } else if (color == "purple") {
    hl = "#d9d2e9";
  };
  _gas.getRange(row, 1, 1, _lcol).setBackground(hl);
  // Keep cells yellow to indicate input
  _gas.getRange(row, 14).setBackground("#ffff00");
  _gas.getRange(row, 17).setBackground("#ffff00");
  _gas.getRange(row, 20).setBackground("#ffff00");
  _gas.getRange("A" + row + ":B" + row).setFontColor(hl);
  // Follow-on dates go on all rows, if applicable
  if (_gas.getRange(row, 21).getValue() != "") {
    _gas.getRange(row, 22).setBackground("#ffff00");
  };  
  _gas.getRange(row, 4).setValue(status_message);
};

function generateEmailTableOld(list) {
  var ths = "<th style='border:1px solid black;padding:5px;'>"
  var table_html = "<table style='border:1px solid black;border-collapse:collapse; '>";
  for (var i = 0; i < list.length; i++) {
    table_html += "<tr>";
    // Add color specific to each row
    var tds = "<td style='border:1px solid black;padding:5px;background-color:" + list[i][8] + ";'>" 
    for (var j = 0; j < list[i].length - 2; j++) {
      if (i == 0) {
        table_html += ths + list[i][j] + "</th>";
      } else {
        table_html += tds + list[i][j] + "</td>";
      };
    };
    table_html += "</tr>";
  };
  table_html += "</table>";
  return table_html;
};
*/


/*
function highlightRowsInitial() {
  getVariables2();
  for (var i = _frow; i <= _lrow; i++) {
    // Base year
    if (_gas.getRange(i, 3).getValue().indexOf("Option") == -1 && _gas.getRange(i, 3).getValue() != "") {
      updateColorStatusBase(i, "blue", "");
    // Spacer rows
    } else if (_gas.getRange(i, 3).getValue() == "") {
      updateColorStatusBase(i, "black", "");
    };
  };
};


// Create rows for option years
function createStartEndDates() {
  getVariables2();
  var dates = _gas.getRange("K" + _frow + ":L" + _lrow).getValues();
  var dates_final = [];
  for (var i = 0; i < dates.length; i++) {
    // Skip if start date already there if ran previously
    if (dates[i][0] != "") {
      dates_final.push([dates[i][0], dates[i][1]]);
      continue;
    // First run through, skip if no work is to be done
    } else if (dates[i][1] == "" || dates[i][1] == "TBD") {
      dates_final.push(["", ""]);
      continue;
    } else {
      // Split on newlines
      var dates_sp = String(dates[i][1]).split("\n");
      for (var j = 0; j < dates_sp.length; j++) {
        var dates_cl = dates_sp[j].trim().replace(/\s/g, "")
        // First separate if there is a colon
        var dates_col = dates_cl.split(":");
        if (dates_col.length > 1) {
          var st_date = dates_col[1].split("-");
        } else {
          var st_date = dates_col[0].split("-");
        };
        dates_final.push([st_date[0], st_date[1]]);
      };
      i += j - 1;
    };
  };
  // Insert data
  //console.log(dates_final);
  //console.log(dates_final.length);
  _gas.getRange(_frow, 11, dates_final.length, dates_final[0].length).setValues(dates_final);
};


function addNoticeDates() {
  getVariables2();
  var data_status = _gas.getRange("C" + _frow + ":S" + _lrow).getValues();
  var days_ms = 1000*60*60*24;
  var list_days = [];
  for (var i = 0; i < data_status.length; i++) {
    // Only run for options
    var option_check = data_status[i][0].toLowerCase().indexOf("option");
    if (i != data_status.length - 1) {
      var option_check_next = data_status[i + 1][0].toLowerCase().indexOf("option");
    } else {
      var option_check_next = -1;
    };
    if (option_check > -1) {
      var sdate = new Date(data_status[i][8]).getTime();
      // Add follow-on date to last row of option year
      if (option_check_next == -1) {
        var follow_on = Utilities.formatDate(new Date(new Date(data_status[i][9]).getTime() + days_ms), "EST", "M/d/yyyy");
      } else {
        var follow_on = "";
      };
      list_days.push([Utilities.formatDate(new Date(sdate - (119 * days_ms)), "EST", "M/d/yyyy"),
                      Utilities.formatDate(new Date(sdate - (59 * days_ms)), "EST", "M/d/yyyy"),
                      Utilities.formatDate(new Date(sdate - ((data_status[i][15] - 1) * days_ms)), "EST", "M/d/yyyy"),
                      follow_on
      ]);
    // Add follow-on date to base year actions with no options
    } else if (data_status[i][0] != "" && option_check == -1 && option_check_next == -1) {
      list_days.push(["",
                      "",
                      "",
                      Utilities.formatDate(new Date(new Date(data_status[i][9]).getTime() + days_ms), "EST", "M/d/yyyy")
      ]);
    } else {
      list_days.push(["", "", "", ""]);
    };
  };
  var list_120_60 = _gas.getRange("M" + _frow + ":U" + _lrow).getValues();
  for (i = 0; i < list_120_60.length; i++) {
    list_120_60[i][0] = list_days[i][0];
    list_120_60[i][3] = list_days[i][1];
    list_120_60[i][6] = list_days[i][2];
    list_120_60[i][8] = list_days[i][3];
  };
  _gas.getRange(_frow, 13, list_120_60.length, list_120_60[0].length).setValues(list_120_60);
};


function mainCheckOld(row) {
  getVariables2();
  var data = _gas.getRange("L" + row + ":U" + row).getValues();
  var days_ms = 1000*60*60*24;
  var tdate_ms = _today.getTime();
  // Exit if no 120-day notice column and follow-on column are empty
  if (data[0][0] == "" && data[0][8] == "") {
    return;
  // Mark row as complete if option has been exercised and no follow-on required; check if new option is required
  } else if (data[0][7] != "" && data[0][8] == "") {
    if (tdate_ms >= new Date(data[0][6]).getTime()) {
      updateColorStatus(row, "white", "");
      createDate(row);   
      return;
    } else {
      updateColorStatus(row, "white", "");
      return;
    };
  // If follow-on has been awarded, order or contract is ready for closout
  } else if (data[0][9] != "" && tdate_ms >= new Date(data[0][8]).getTime()) {
    updateColorStatus(row, "blue", "");
    return;
  // If follow-on has been awarded, but order is still not ready for closeout
  } else if (data[0][9] != "" && tdate_ms < new Date(data[0][8]).getTime()) {
    updateColorStatus(row, "white", "");
    return;
  };
  // Check for all other dates
  var day_120_diff = Math.round((new Date(data[0][0]).getTime() - tdate_ms) / days_ms) + 1;
  var day_60_diff = Math.round((new Date(data[0][3]).getTime() - tdate_ms) / days_ms) + 1;
  var day_ex_diff = Math.round((new Date(data[0][6]).getTime() - tdate_ms) / days_ms) + 1;
  if (data[0][8] == "") {
    var day_fo_diff = 999;
  } else {
    var day_fo_diff = Math.round((new Date(data[0][8]).getTime() - tdate_ms) / days_ms) + 1;
  };
  // 120 day notice to client
  if (data[0][1] == "") {
    checkDays(row, day_120_diff, "120-day notice to client");
  // 60 day notice to contractor
  } else if (data[0][4] == "") {
    checkDays(row, day_60_diff, "60-day notice to contractor");
  // Exercise option notice
  } else if (data[0][7] == "") {
    checkDays(row, day_ex_diff, "Exercise option");
  // Follow-on notice
  } else if (data[0][8] != "" && data[0][9] == "") {
    checkDays(row, day_fo_diff, "Follow-on");
  // All other actions
  } else {
    checkDays(row, 999, "");
  };
};


function updateColorStatusOld(row, color, status_message) {
  var hl = "#ffffff";
  if (color == "blue") {
    hl = "#c9daf8";
  } else if (color == "gray") {
    hl = "#d9d9d9";
  } else if (color == "yellow") {
    hl = "#fff2cc";
  } else if (color == "orange") {
    hl = "#fce5cd";
  } else if (color == "red") {
    hl = "#f4cccc";
  } else if (color == "purple") {
    hl = "#d9d2e9";
  };
  _gas.getRange(row, 1, 1, _lcol).setBackground(hl);
  // Highlight applicable cells
  if (_gas.getRange(row, 12).getValue() != "") {
    _gas.getRange(row, 13).setBackground("#ffff00");
    _gas.getRange(row, 16).setBackground("#ffff00");
    _gas.getRange(row, 19).setBackground("#ffff00");
  };
  // Follow-on dates go on all rows, if applicable
  if (_gas.getRange(row, 20).getValue() != "") {
    _gas.getRange(row, 21).setBackground("#ffff00");
  };  
  _gas.getRange(row, 4).setValue(status_message);
};
*/

