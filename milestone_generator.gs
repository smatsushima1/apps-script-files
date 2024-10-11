

function getVariables() {
  _ga = SpreadsheetApp.getActive();
  _mg = _ga.getSheetByName("Milestone Generator");
  _cms = _ga.getSheetByName("Copy of Master Spreadsheet");
  _dvs = _ga.getSheetByName("Data Values");
  _pop_st = _mg.getRange(1, 2).getValue();
  // Convert to string
  _pop_st = ((_pop_st.getMonth() + 1) + "/" + _pop_st.getDate() + "/" + (_pop_st.getYear() + 1900));
  _dv = _mg.getRange(2, 2).getValue();
  _type = _mg.getRange(3, 2).getValue();
  _comm = _mg.getRange(4, 2).getValue();
  _holidays = _dvs.getRange("F1:H" + getLastDataRow(_dvs, "H")).getValues();
  _last_days = _dvs.getRange("K1:M" + getLastDataRow(_dvs, "M")).getValues();
};


/////////////////////////////////////////////////////////////////////////////////
////////////////////////////// Generate Milestones //////////////////////////////
/////////////////////////////////////////////////////////////////////////////////


function generateMilestones() {
  getVariables();
  var data = getData();
  _mg.getRange("D:L").clear();
  _mg.getRange(1, 4, data.length, data[0].length).setValues(data);
  formatTable();
  moveButtons();
};


function getData() {
  var type_rng = _cms.getRange("2:2").getValues()[0];
  // Find column to start the type
  var col_st = "";
  for (var i = 0; i < type_rng.length; i++) {
    if (type_rng[i] == _type) {
      col_st = i + 1;
      break;
    } else if (i == type_rng.length) {
      col_st = "";
    };
  };
  // Find column to end
  var col_end = "";
  for (i = col_st; i < type_rng.length; i++) {
    if (type_rng[i] != "" || i == type_rng.length) {
      col_end = i;
      break;
    };
  };
  var row_max = _cms.getLastRow();
  var data_raw = _cms.getRange(columnLetters(col_st) + 3 + ":" + columnLetters(col_end) + row_max).getValues();
  // Only save applicable milestones for type
  // First identify column that applies based on dollar value and commerciality
  for (i = 0; i < data_raw[0].length; i++) {
    if ((_dv >= data_raw[0][i] && _dv <= data_raw[1][i]) && (data_raw[2][i] == "" || data_raw[2][i] == _type)) {
      var col_start = i;
      break;
    };
  };
  // Pull data for specific columns
  var data_consolidated = [];
  var data_items = _cms.getRange("A3:B" + row_max).getValues();
  for (i = 3; i < data_raw.length; i++) {
    var data_row = [];
    // Add two columns back in
    data_row.push(data_items[i][0]);
    data_row.push(data_items[i][1]);
    for (var j = col_start; j <= col_start + 3; j++) {
      data_row.push(data_raw[i][j]);
    };
    data_consolidated.push(data_row);
  };
  // Remove all blanks, N/A, and - rows
  var data_no_blanks = [];
  for (i = 0; i < data_consolidated.length; i++) {
    // Push the first row
    if (i == 0) {
      data_no_blanks.push(data_consolidated[i]);
    // Push all other items
    } else {
      var days = data_consolidated[i][2];
      if (days != "" && days != "N/A" && days != "-" && days != "--") {
        data_row = [];
        for (var j = 0; j < data_consolidated[0].length; j++) {
          data_row.push(data_consolidated[i][j]);
        };
        data_no_blanks.push(data_row);
      };
    };
  };
  return findStart(data_no_blanks, _pop_st);
};


function findStart(data, pop_end) {
  // Days, add one to each date to get the actual date after converting to time
  var end_date = pop_end;
  var after_lag = "";
  var start_date = "";
  // Add total column
  data.push(["", "Total Days", "", "", "", "", "", "", ""]);
  // Find start dates, working backwards
  for (var i = data.length - 1; i >= 0; i--) {
    // Save the entire row
    var data_row = data[i];
    // Add in column headers for last iteration
    if (i == 0) {
      data[i][0] = "Number";
      data[i].push("Start Date");
      data[i].push("End Date");
      data[i].push("Actual Date");
    // Add all days and input total day
    } else if (i == data.length - 1) {
      var total_days = 0;
      for (j = 1; j < data.length - 1; j++) {
        total_days += data[j][2];
      };
      data[i][2] = total_days;
    // All other items
    } else {
      after_lag = lagDays(end_date, data_row[4], data_row[3]);
      start_date = previousDays(after_lag, data_row[2], data_row[3]);
      // Push final data to list
      data[i].push(start_date);
      data[i].push(after_lag);
      data[i].push("");
      // Reset day for next iteration
      end_date = start_date;
    };
  };
  return data;
};


function previousDay(date) {
  var m = date[0];
  var d = date[1];
  var y = date[2];
  var pm = m;
  var pd = d - 1;
  var py = y;
  // Only check if current day was 1
  if (pd == 0) {
    for (var i = 0; i < _last_days.length; i++) {
      // Only need to check month and year
      if (pm == _last_days[i][0] && py == _last_days[i][2]) {
        pm = _last_days[i - 1][0];
        pd = _last_days[i - 1][1];
        py = _last_days[i - 1][2];
        break;
      };
    };
  };
  // Holiday check
  for (var i = 0; i < _holidays.length; i++) {
    if (pm == _holidays[i][0] && pd == _holidays[i][1] && py == _holidays[i][2]) {
      pd -= 1;
    };
  };
  // Month and year check again
  if (pd == 0) {
    for (var i = 0; i < _last_days.length; i++) {
      // Only need to check month and year
      if (pm == _last_days[i][0] && py == _last_days[i][2]) {
        pm = _last_days[i - 1][0];
        pd = _last_days[i - 1][1];
        py = _last_days[i - 1][2];
        break;
      };
    };
  };
  // Done
  return [pm, pd, py];
};


function workingDayCheck(end_date) {
  // Check for Sundays
  var pd_daynum = new Date(end_date[0] + "/" + end_date[1] + "/" + end_date[2]).getDay();
  if (pd_daynum == 0) {
    end_date = previousDay([end_date[0], end_date[1], end_date[2]]);
  };
  // Run again if on a Saturday
  pd_daynum = new Date(end_date[0] + "/" + end_date[1] + "/" + end_date[2]).getDay();
  if (pd_daynum == 6) {
    end_date = previousDay([end_date[0], end_date[1], end_date[2]]);
  };
  return end_date;
};


// Identicaly to previousDays, but designed for adds one day to the i loop
function lagDays(end_date, days, working_ind) {
  var date = convertStringToList(end_date);
  // All other days
  for (var i = 1; i <= days; i++) {
    date = previousDay(date);
    if (working_ind == "Working") {
      date = workingDayCheck(date);
    };
  };
  return (date[0] + "/" + date[1] + "/" + date[2]);
};


function previousDays(end_date, days, working_ind) {
  var date = convertStringToList(end_date);
  for (var i = 1; i <= days - 1; i++) {
    date = previousDay(date);
    if (working_ind == "Working") {
      date = workingDayCheck(date);
    };
  };
  return (date[0] + "/" + date[1] + "/" + date[2]);
};


function formatTable() {
  // Format columns
  _mg.getRange("D:I").setNumberFormat("General");
  _mg.getRange("J:L").setNumberFormat("dddd, m/d/yyyy");
  _mg.autoResizeColumns(4, 9);
  // Reset borders
  _mg.getRange("D:L").setBorder(false, false, false, false, false, false);
  // Add borders
  var last_row = getLastDataRow(_mg, "F");
  _mg.getRange("D1:L" + last_row).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  // Wrap columns
  _mg.getRange("E:E").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  // Bold headers and total row
  _mg.getRange("D1:L1").setHorizontalAlignment("center").setFontWeight("bold");
  _mg.getRange("D" + last_row + ":L" + last_row).setFontWeight("bold");
  // Convert day types to drop-down items
  var day_types = _dvs.getRange("C4:C5").getValues();
  var day_types_final = [];
  day_types.forEach((i) => day_types_final.push(i[0]));
  _mg.getRange('G2:G' + getLastDataRow(_mg, "G")).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(day_types_final, true).build());
};


///////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////// Update Milestones ////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////


// Update dates based on actuals
function updateMilestones() {
  getVariables();
  var data = updateDates();
  _mg.getRange("D:L").clear();
  _mg.getRange(1, 4, data.length, data[0].length).setValues(data);
  formatTable();
};


function updateDates() {
  var table = _mg.getRange("D1:L" + getLastDataRow(_mg, "F")).getValues();
  return findEnd(table);
};


function findEnd(table) {
  var start = convertDateToString(table[1][6]);
  var total_days = 0;
  for (var i = 1; i < table.length; i++) {
    var drow = table[i];
    // All other items
    if (i != table.length - 1) {
      console.log("i: " + i);
      // First update start date
      drow[6] = start;
      // If "Actual Date" is present, then that is the end_date before lag days
      if (drow[8] != "") {
        // Calculate total days, replace PALT timeline
        var num_days = dateDiff(drow[6], convertDateToString(drow[8]), drow[3]);
        total_days += num_days;
        drow[2] = num_days;
        start = lagDaysNext(convertDateToString(drow[8]), drow[4], drow[3]);
        continue;
      // Regular items with no actual dates
      } else {
        total_days += drow[2];
        // End date
        var end = nextDays(start, drow[2], drow[3]);
        drow[7] = end;
        // Reset start date for next iteration
        start = lagDaysNext(end, drow[4], drow[3]);
      };
    // Total row
    } else {
      drow[2] = total_days;
    };
  };
  return table;
};


function nextDays(start_date, days, working_ind) {
  var date = start_date;
  for (var i = 1; i <= days - 1; i++) {
    date = nextDay(date);
    if (working_ind == "Working") {
      date = workingDayCheckNext(date);
    };
  };
  return date;
};


// Identicaly to nextDays, but designed to add one day to the i loop
function lagDaysNext(start_date, days, working_ind) {
  var date = start_date;
  // All other days
  for (var i = 1; i <= days; i++) {
    date = nextDay(date);
    if (working_ind == "Working") {
      date = workingDayCheckNext(date);
    };
  };
  return date;
};


function nextDay(date) {
  var date = convertStringToList(date);
  var m = Number(date[0]);
  var d = Number(date[1]);
  var y = Number(date[2]);
  var nm = m;
  var nd = d + 1;
  var ny = y;
  // Only start check if day 28 comes
  if (nd >= 28) {
    // EOM check
    for (var i = 0; i < _last_days.length; i++) {
      if (nm == _last_days[i][0] && nd == (_last_days[i][1] + 1) && ny == _last_days[i][2]) {
        nm = _last_days[i + 1][0];
        nd = 1;
        ny = _last_days[i + 1][2]
        break;
      };
    };
  };
  // Holiday check
  for (var i = 0; i < _holidays.length; i++) {
    if (nm == _holidays[i][0] && nd == _holidays[i][1] && ny == _holidays[i][2]) {
      nd += 1;
    };
  };
  // EOM check 2
  if (nd >= 28) {
    // EOM check
    for (var i = 0; i < _last_days.length; i++) {
      if (nm == _last_days[i][0] && nd == (_last_days[i][1] + 1) && ny == _last_days[i][2]) {
        nm = _last_days[i + 1][0];
        nd = 1;
        ny = _last_days[i + 1][2]
        break;
      };
    };
  };  
  // Done
  return convertListToString([nm, nd, ny]);
};


function workingDayCheckNext(date) {
  // Check for Saturdays
  var nd_daynum = new Date(date).getDay();
  if (nd_daynum == 6) {
    date = nextDay(date);
  };
  // Run again for Sundays
  nd_daynum = new Date(date).getDay();
  if (nd_daynum == 0) {
    date = nextDay(date);
  };
  return date;
};


function dateDiff(start_date, end_date, working_ind) {
  // Convert date to a string
  var sdate = start_date;
  var total_days = 0;
  // Loop until 
  while (sdate != end_date) {
    total_days += 1;
    if (working_ind != undefined) {
      sdate = workingDayCheckNext(nextDay(sdate));
    } else {
      sdate = nextDay(sdate);
    };
  };
  return total_days + 1;
};


function convertStringToList(date) {
  var d = date.split("/");
  return [Number(d[0]), Number(d[1]), Number(d[2])];
};


// Turn date list into a string
function convertListToString(date) {
  return date[0] + "/" + date[1] + "/" + date[2];
};


// Turn date into string
function convertDateToString(date) {
  return (date.getMonth() + 1) + "/" + date.getDate() + "/" + (date.getYear() + 1900);
};


/////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////// Accessory Functions ////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////


function columnLetters(column) {
  return _cms.getRange(1, column).getA1Notation().slice(0, -1);
};


function getLastDataRow(sheet, letter) {
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(letter + lastRow);
  if (range.getValue() !== "") {
    return lastRow;
  } else {
    return range.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  };
};


function insertRow() {
  getVariables();
  var row = SpreadsheetApp.getCurrentCell().getRow();
  _mg.getRange("D" + row + ":L" + row).insertCells(SpreadsheetApp.Dimension.ROWS);
  moveButtons();
};


function deleteRow() {
  getVariables();
  var row = SpreadsheetApp.getCurrentCell().getRow();
  _mg.getRange("D" + row + ":L" + row).deleteCells(SpreadsheetApp.Dimension.ROWS);
  moveButtons();
  updateMilestones();
};


function moveButtons() {
  var drawings = _mg.getDrawings();
  drawings[0].setPosition(6, 2, 0, 0);
  drawings[1].setPosition(6, 2, 0, 50);
  drawings[2].setPosition(6, 2, 0, 100);
  drawings[3].setPosition(6, 2, 0, 175);
  _mg.getRange("B1").activate();
};


/* ////////////////////////////// Deprecated //////////////////////////////
function findStartOld(data, pop_end) {
  // Days, add one to each date to get the actual date after converting to time
  var end_date = new Date(pop_end).getTime() + _day_ms;
  var after_lag = "";
  var start_date = "";
  // Find start dates, working backwards
  for (var i = data.length - 1; i >= 0; i--) {
    console.log(Utilities.formatDate(new Date(end_date), "EST", "MM/dd/yyyy"));
    var data_row = data[i];
    // All other items
    if (i != 0) {
      // Working days
      if (data_row[3] == "Working") {
        // Working days
        after_lag = dayDiff(end_date, data_row[4], "Working", "lag"); 
        start_date = dayDiff(after_lag, data_row[2], "Working", );
      // Calendar days
      } else {
        // Subtract lag days
        after_lag = dayDiff(end_date, data_row[4], "Calendar", "lag"); 
        start_date = dayDiff(after_lag, data_row[2], "Calendar", );
      };
      // Push final data to list
      data[i].push(Utilities.formatDate(new Date(start_date), "EST", "MM/dd/yyyy"));
      data[i].push(Utilities.formatDate(new Date(after_lag), "EST", "MM/dd/yyyy"));
      // Reset day for next iteration
      end_date = start_date;
    // Add in column headers for last iteration
    } else {
      data[i].push('Start Date');
      data[i].push('End Date');
    };
  };
  return data;
};


// Returns the start date in ms
function dayDiff(end_time, days, type_ind, lag_ind) {
  if (lag_ind != "lag") {
    days -= 1;
  };
  var iday = 0;
  ////////////////////// Working day section ////////////////////
  if (type_ind == "Working") {
    for (var i = 0; i < days; i++) {
      iday += 1;
      iday += holidayCheck(end_time - (iday * _day_ms) + 1, type_ind);
      // Weekned check, add 2 if next day is Sunday
      if (new Date(end_time - (iday * _day_ms) + 1).getDay() - 1 == 0) {
        iday += 2;
        iday += holidayCheck(end_time - (iday * _day_ms) + 1, type_ind);
      };
    };
  //////////////////// Calendar day section ////////////////////
  } else {
    for (var i = 0; i < days; i++) {
      iday += 1;
      iday += holidayCheck(end_time - (iday * _day_ms) + 1, type_ind);
    };
  };
  //console.log(Utilities.formatDate(new Date(end_time - (iday * _day_ms)), "EST", "MM/dd/yyyy"));
  return end_time - (iday * _day_ms);
};


// Make holidays return a number to add onto iday in the daydiff functions: Monday = 3, normal holidays = 1, none = 0
function holidayCheck(end_time, day_type) {
  // Save holidays in a list to compare against later
  var holidays = [];
  for (var i = 0; i < _holidays.length; i++) {
    holidays.push(_holidays[i][0]);
  };
  var end_date = Utilities.formatDate(new Date(end_time), "EST", "MM/dd/yyyy");
  if (holidays.indexOf(end_date) > 0) {
    if (day_type == "Working") {
      // Add extra days if Monday is a holiday
      if (new Date(end_date).getDay() == 1) {
        return 3;
      // Add just one day for everything else
      } else {
        return 1;
      };
    // Calendar days
    } else {
      //console.log("holiday");
      return 1;
    };
  } else {
    return 0;
  };
};


// Returns the start date in ms
function dayDiffDev(end_time, days, type_ind, lag_ind) {
  if (lag_ind != "lag") {
    days -= 1;
  };
  var iday = 0;
  ////////////////////// Working day section ////////////////////
  if (type_ind == "Working") {
    for (var i = 0; i < days; i++) {
      iday += 1;
      iday += holidayCheck(end_time - (iday * _day_ms), type_ind);
      // Weekned check, add 2 if next day is Sunday
      if (new Date(end_time - (iday * _day_ms)).getDay() - 1 == 0) {
        iday += 2;
        iday += holidayCheck(end_time - (iday * _day_ms), type_ind);
      };
    };
  //////////////////// Calendar day section ////////////////////
  } else {
    for (var i = 0; i < days; i++) {
      iday += 1;
      iday += holidayCheck(end_time - (iday * _day_ms), type_ind);
    };
  };
  //console.log(Utilities.formatDate(new Date(end_time - (iday * _day_ms)), "EST", "MM/dd/yyyy"));
  return end_time - (iday * _day_ms);
};


function dayDiffDev2(end_time, days, type_ind, lag_ind) {
  // Working and calendar days
  if (lag_ind == undefined) {
    //days += 1;
    var iday = -1;
  // Lag days
  } else {
    var iday = 0;
  };
  var iday = 0;
  ////////////////////// Working day section ////////////////////
  if (type_ind == "Working") {
    for (var i = 1; i <= days; i++) {
      iday += 1;
      iday += holidayCheck(end_time - (iday * _day_ms), type_ind);
      // Weekned check, add 2 if next day is Sunday
      if (new Date(end_time - (iday * _day_ms)).getDay() == 0) {
        iday += 2;
        iday += holidayCheck(end_time - (iday * _day_ms), type_ind);
      };
      //console.log("i: " + i + "; iday: " + iday + "; Next day num: " + (new Date(end_time - (iday * _day_ms)).getDay()) + "; Next day date: " + Utilities.formatDate(new Date(end_time - (iday * _day_ms)), "EST", "MM/dd/yyyy"));
    };
  //////////////////// Calendar day section ////////////////////
  } else {
    for (i = 1; i <= days; i++) {
      iday += 1;
      iday += holidayCheck(end_time - (iday * _day_ms), type_ind);
      //console.log("i: " + i + "; iday: " + iday + "; Next day num: " + (new Date(end_time - (iday * _day_ms)).getDay()) + "; Next day date: " + Utilities.formatDate(new Date(end_time - (iday * _day_ms)), "EST", "MM/dd/yyyy"));
    };
  };
  //console.log(Utilities.formatDate(new Date(end_time - (iday * _day_ms)), "EST", "MM/dd/yyyy"));
  return end_time - (iday * _day_ms);
};


function dev002() {
  getVariables();
  console.log("-------------------- Working --------------------");
  var dtime = new Date("01/17/2024").getTime() + _day_ms;
  console.log(Utilities.formatDate(new Date(dayDiff(dtime, 1, "Working", )), "EST", "MM/dd/yyyy"));
  console.log("-------------------- Calendar --------------------");
  console.log(Utilities.formatDate(new Date(dayDiff(dtime, 1, "Calendar", )), "EST", "MM/dd/yyyy"));
  console.log("-------------------- Lag - Working --------------------");
  dtime = new Date("01/17/2024").getTime() + _day_ms;
  console.log(Utilities.formatDate(new Date(dayDiff(dtime, 1, "Working", "lag")), "EST", "MM/dd/yyyy"));
  console.log("-------------------- Lag - Calendar --------------------");
  console.log(Utilities.formatDate(new Date(dayDiff(dtime, 1, "Calendar", "lag")), "EST", "MM/dd/yyyy"));
};


function findStartOrig(data, pop_end) {
  // Days, add one to each date to get the actual date after converting to time
  var end_date = pop_end;
  var after_lag = "";
  var start_date = "";
  // Find start dates, working backwards
  for (var i = data.length; i >= 0; i--) {
    if (i == data.length) {

    }
    var data_row = data[i];
    // Add in column headers for last iteration
    if (i == 0) {
      data[i].push("Start Date");
      data[i].push("End Date");
      data[i].push("Actual Date")
    // All other items
    } else {
      after_lag = lagDays(end_date, data_row[4], data_row[3]);
      start_date = previousDays(after_lag, data_row[2], data_row[3]);
      // Push final data to list
      data[i].push(start_date);
      data[i].push(after_lag);
      data[i].push("");
      // Reset day for next iteration
      end_date = start_date;
    };
  };
  return data;
};


*/

