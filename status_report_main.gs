

////////////////////////////////////////////////////////////////////////////////////////////////////
// Template Generation
////////////////////////////////////////////////////////////////////////////////////////////////////


function getVariables1(co_ind) {
  _ga = SpreadsheetApp.getActive();
  _gas = _ga.getActiveSheet();
  _arow = _gas.getActiveCell().getRow();
  _today = new Date();
  // Auto-Generated Template folder
  _ag_folder = DriveApp.getFolderById("");
  // Save all values on the sheet
  var data = _gas.getRange(_arow, 1, 1, 23).getValues();
  // Save contract and type
  if (data[0][1] != "") {
    var contract = (data[0][0] + " " + data[0][1]).trim();
  } else {
    var contract = data[0][0].trim();
  };
  // Year
  if (data[0][10].trim().split("\n").length == 1) {
    var year = "Base";
    var pop = data[0][10].replace("Base:", "").trim().split("-");
    var pop_st = convertDate(pop[0]);
    var pop_end = convertDate(pop[1]);
    var pop_full = pop_st + " to " + pop_end;
    var pop_exp = pop_end;
  } else {
    var pop = findPOP(data[0][10].trim());
    var year = convertOption(pop[0]);
    var pop_st = convertDate(pop[1]);
    var pop_end = convertDate(pop[2]);
    var pop_full_exp = findPOPFull(data[0][10].trim());
    var pop_full = convertDate(pop_full_exp[0]) + " to " + convertDate(pop_full_exp[1]);
    var pop_exp = convertDate(pop_full_exp[2]);
  };
  // Identify type
  var type = "Contract";
  if (contract.substring(8, 9) == "A" && contract.indexOf(" ") > -1) {
    type = "BPA Call";
  } else if (contract.substring(8, 9) == "A") {
    type = "BPA";
  } else if (contract.substring(0, 2) == "GS" && contract.indexOf(" ") > -1) {
    type = "GSA Order";
  } else if (contract.indexOf(" ") > -1) {
    type = "Order";
  } else {
    type = "Contract";
  };
  // Find base contract information
  var description = data[0][2].trim().replace(/\n/g, " ");
  var contractor = data[0][4];
  var client = data[0][5];
  var cor = data[0][21];
  var contractor_poc = data[0][22];
  // Contracting Officer
  if (co_ind != undefined) {
    var co = Browser.inputBox("Enter Contracting Officer Name");
  } else {
    var co = "";
  };
  return[contract, // 0
         year, // 1
         type, // 2
         description, // 3
         contractor, // 4
         client, // 5
         co, // 6
         cor, // 7
         contractor_poc, // 8
         pop_st, // 9
         pop_end, // 10
         pop_full, // 11
         pop_exp // 12
  ];
};


function generate120DayLetter() {
  var values = getVariables1(1);
  // Initial error check
  consoleHeader("Initial Check");
  var errors = errorTest(values, 1);
  if (errors[0] == 1) {
    Browser.msgBox(errors[1]);
    return;
  };
  // Create copy of template and save ID
  consoleHeader("Creating Word Doc");
  var new_file = copyTemplate("", values[1] + " " + values[0]);
  // Open doc and replace fields with pulled data
  consoleHeader("Inputting Values");
  var doc = DocumentApp.openById(new_file[0]);
  var body = doc.getBody();
  body.replaceText("{Date}", Utilities.formatDate(_today, "EST", "MMMM d, yyyy"));
  body.replaceText("{Contract}", values[0]);
  body.replaceText("{Option}", values[1]);
  body.replaceText("{ContractType}", values[2]);
  body.replaceText("{Description}", values[3]);
  body.replaceText("{Contractor}", values[4]);
  body.replaceText("{Client}", values[5]);
  body.replaceText("{ContractingOfficer}", values[6]);
  body.replaceText("{COR}", values[7].split("\n"))[0];
  body.replaceText("{POPStart}", values[9]);
  body.replaceText("{POPEnd}", values[10]);
  body.replaceText("{POPExpiration}", values[12]);
  body.replaceText("{Return}", Utilities.formatDate(new Date(_today.getTime() + 3600000*24*30), "EST", "MMMM d, yyyy"));
  body.replaceText("{ContactName}", parseFullName());
  body.replaceText("{ContactEmail}", Session.getActiveUser().getEmail());
  doc.saveAndClose();
  consoleHeader("Finished");
  downloadFile(new_file[0], new_file[1]);
};


function generate60DayLetter() {
  var values = getVariables1(1);
  try {
    var contractor_full = parseNameEmail(values[8].split("\n"));
    var contractor_fullname = contractor_full[0].split(" ");
  } catch(err) {
    var contractor_full = ["[Contractor_Name]", "[Contractor_Email]"];
    var contractor_fullname = ["[Contractor_Name]"];
  };
  // Initial error check
  consoleHeader("Initial Check");
  var errors = errorTest(values, 1);
  if (errors[0] == 1) {
    Browser.msgBox(errors[1]);
    return;
  };
  // Create copy of template and save ID
  consoleHeader("Creating Word Doc");
  var new_file = copyTemplate("", values[1] + " " + values[0]);
  // Open doc and replace fields with pulled data
  consoleHeader("Inputting Values");
  var doc = DocumentApp.openById(new_file[0]);
  var body = doc.getBody();
  body.replaceText("{Date}", Utilities.formatDate(_today, "EST", "MMMM d, yyyy"));
  body.replaceText("{Contract}", values[0]);
  body.replaceText("{Option}", values[1]);
  body.replaceText("{ContractType}", values[2]);
  body.replaceText("{Description}", values[3]);
  body.replaceText("{Contractor}", values[4]);
  body.replaceText("{ContractingOfficer}", values[6]);
  body.replaceText("{POPStart}", values[9]);
  body.replaceText("{POPEnd}", values[10]);
  body.replaceText("{ContractorPOC}", contractor_full[0]);
  body.replaceText("{ContractorPOCEmail}", contractor_full[1]);
  body.replaceText("{ContractorGreeting}", contractor_fullname[contractor_fullname.length - 1]);
  body.replaceText("{ContactName}", parseFullName());
  body.replaceText("{ContactEmail}", Session.getActiveUser().getEmail());
  doc.saveAndClose();
  consoleHeader("Finished");
  downloadFile(new_file[0], new_file[1]);
};


function generateDash8Letter() {
  var values = getVariables1(1);
  try {
    var contractor_full = parseNameEmail(values[8].split("\n"));
    var contractor_fullname = contractor_full[0].split(" ");
  } catch(err) {
    var contractor_full = ["[Contractor_Name]", "[Contractor_Email]"];
    var contractor_fullname = ["[Contractor_Name]"];
  };
  // Initial error check
  consoleHeader("Initial Check");
  var errors = errorTest(values, 1);
  if (errors[0] == 1) {
    Browser.msgBox(errors[1]);
    return;
  };
  // Create copy of template and save ID
  consoleHeader("Creating Word Doc");
  var new_file = copyTemplate("", " " + values[0]);
  // Open doc and replace fields with pulled data
  consoleHeader("Inputting Values");
  var doc = DocumentApp.openById(new_file[0]);
  var body = doc.getBody();
  // Open doc and replace fields with pulled data
  consoleHeader("Inputting Values");
  var doc = DocumentApp.openById(new_file[0]);
  var body = doc.getBody();
  body.replaceText("{Date}", Utilities.formatDate(_today, "EST", "MMMM d, yyyy"));
  body.replaceText("{Contract}", values[0]);
  body.replaceText("{ContractType}", values[2]);
  body.replaceText("{Description}", values[3]);
  body.replaceText("{Contractor}", values[4]);
  body.replaceText("{ContractingOfficer}", values[6]);
  body.replaceText("{POPStart}", values[9]);
  body.replaceText("{POPEnd}", values[10]);
  body.replaceText("{ContractorPOC}", contractor_full[0]);
  body.replaceText("{ContractorPOCEmail}", contractor_full[1]);
  body.replaceText("{ContractorGreeting}", contractor_fullname[contractor_fullname.length - 1]);
  body.replaceText("{ContactName}", parseFullName());
  body.replaceText("{ContactEmail}", Session.getActiveUser().getEmail());
  doc.saveAndClose();
  consoleHeader("Finished");
  downloadFile(new_file[0], new_file[1]);
};


function generateCORDelegation() {
  var values = getVariables1(1);
  // Initial error check
  consoleHeader("Initial Check");
  var errors = errorTest(values);
  if (errors[0] == 1) {
    Browser.msgBox(errors[1]);
    return;
  };
  // Create copy of template and save ID
  consoleHeader("Creating Word Doc");
  var new_file = copyTemplate("", " " + values[0]);
  // Open doc and replace fields with pulled data
  consoleHeader("Inputting Values");
  var doc = DocumentApp.openById(new_file[0]);
  var body = doc.getBody();
  body.replaceText("{Date}", Utilities.formatDate(_today, "EST", "MMMM d, yyyy"));
  body.replaceText("{Contract}", values[0]);
  body.replaceText("{ContractType}", values[2]);
  body.replaceText("{Description}", values[3]);
  body.replaceText("{Contractor}", values[4]);
  body.replaceText("{Client}", values[5]);
  body.replaceText("{ContractingOfficer}", values[6]);
  body.replaceText("{COR}", values[7].split("\n"))[0];
  doc.saveAndClose();
  consoleHeader("Finished");
  downloadFile(new_file[0], new_file[1]);
};


function generateModMemorandum() {
  var values = getVariables1(1);
  // Initial error check
  consoleHeader("Initial Check");
  var errors = errorTest(values);
  if (errors[0] == 1) {
    Browser.msgBox(errors[1]);
    return;
  };
  // Create copy of template and save ID
  consoleHeader("Creating Word Doc");
  var new_file = copyTemplate("", " " + values[0]);
  // Open doc and replace fields with pulled data
  consoleHeader("Inputting Values");
  var doc = DocumentApp.openById(new_file[0]);
  var body = doc.getBody();
  body.replaceText("{Date}", Utilities.formatDate(_today, "EST", "MMMM d, yyyy"));
  body.replaceText("{Contract}", values[0]);
  body.replaceText("{ContractType}", values[2]);
  body.replaceText("{Description}", values[3]);
  body.replaceText("{Contractor}", values[4]);
  body.replaceText("{Client}", values[5]);
  body.replaceText("{ContractingOfficer}", values[6]);
  body.replaceText("{PeriodOfPerformance}", values[11]);
  body.replaceText("{ContactName}", parseFullName());
  doc.saveAndClose();
  consoleHeader("Finished");
  downloadFile(new_file[0], new_file[1]);
};


// Returns the file ID and name of the copied template file
function copyTemplate(template_id, file_name_additions) {
  // Open word doc and rename
  var gf = DriveApp.getFileById(template_id);
  var template_name = gf.getName();
  gf.makeCopy(_ag_folder)
  var new_name = template_name.replace("Template", "") + file_name_additions;
  DriveApp.getFilesByName("Copy of " + template_name).next().setName(new_name);
  // Loop through files in current folder to find the copied generated previously
  var files = _ag_folder.getFiles();
  var nid = "";
  // Loop through folder till you find the new file, then save it's Id
  while (files.hasNext()){
    file = files.next();
    // Return Id if its the copied file
    if (file.getName() == new_name) return [file.getId(), new_name];
  };
};


function findPOPFull(date) {
  var new_pop = [];
  date.split("\n").forEach(i => new_pop.push(i.trim().replace(/\s/g, "").split(":")[1].split("-")));
  var continuing = 0;
  var exp_date = "";
  for (var i = 0; i < new_pop.length; i++) {
    // First find start and end dates for full pop
    if (i == 0) {
      var st_date = new_pop[i][0];
    } else if (i == new_pop.length - 1) {
      var end_date = new_pop[i][1];
    };
    // Next find current expiration date
    if (continuing == 0 && _today.getTime() < new Date(new_pop[i][0]).getTime()) {
      var exp_date = new_pop[i - 1][1];
      continuing = 1;
    };
  };
  if (exp_date == "") exp_date = end_date;
  return [st_date, end_date, exp_date];
};


// Regex: https://support.google.com/a/answer/1371415?hl=en
function errorTest(values, option) {
  var sheet_name = trimClean(_gas.getName());
  // Return error code if running on "closed" sheets, "ueis", if contract doesn't have two digits, or if nothing there
  if (sheet_name.indexOf("close") > -1 || sheet_name.indexOf("ueis") > -1 || sheet_name.indexOf("teams") > -1) {
    return [1, "Function can't run on this sheet. Exiting..."];    
  } else  if (!trimClean(values[0]).match(/^(.*\d{2}.*)$/) || trimClean(values[0]) == "") {
    return [1, "No contract detected in selected row. Exiting..."];
  } else if (option != undefined && trimClean(values[1]).indexOf("option") == -1) {
    return [1, "No option detected in selected row. Exiting..."];
  } else {
    return [0, ""];
  };
};


// Because I don't trust Javascript with dates anymore
function convertDate(date) {
  var new_date = new Date(date);
  var m = new_date.getMonth() + 1;
  var month = "";
  if (m == 1) {
    month = "January";
  } else if (m == 2) {
    month = "February";
  } else if (m == 3) {
    month = "March";
  } else if (m == 4) {
    month = "April";
  } else if (m == 5) {
    month = "May";
  } else if (m == 6) {
    month = "June";
  } else if (m == 7) {
    month = "July";
  } else if (m == 8) {
    month = "August";
  } else if (m == 9) {
    month = "September";
  } else if (m == 10) {
    month = "October";
  } else if (m == 11) {
    month = "November";
  } else if (m == 12) {
    month = "December";
  }
  return (month + " " + new_date.getDate() + ", " + (new_date.getYear() + 1900));
};


// Pulls first value that has an email address
function parseNameEmail(poc_info) {
  var name = "";
  var email = "";
  for (var i = 0; i < poc_info.length; i++) {
    var txt = poc_info[i];
    if (txt.indexOf("@") == -1 && txt.trim().indexOf(" ") > 0) {
      name = txt;
    } else if (txt.indexOf("@") >= 0) {
      email = txt;
    };
    if (name != "" && email != "") return [name, email];
  };
};


function convertOption(option) {
  if (option.indexOf("-8") > -1) return option;
  return option.replace("1", "One").replace("2", "Two").replace("3", "Three").replace("4", "Four").replace("5", "Five").replace("6", "Six").replace("7", "Seven").replace("8", "Eight").replace("9", "Nine");
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


// Converts to string, converts to lower case, trims, removes dashes, removes white space
function trimClean(text) {
  return text.toString().toLowerCase().trim().replace(/-/g, "").replace(/\s/g, "");
};


function downloadFile(file_id, file_name) {
  // Fetch the docx blob
  var token = ScriptApp.getOAuthToken();
  var blob = UrlFetchApp.fetch('https://docs.google.com/feeds/download/documents/export/Export?id=' + file_id + '&exportFormat=docx', 
                               {headers: {Authorization: 'Bearer ' + token}}).getBlob();
  // Put the file in a drive folder
  _ag_folder.createFile(blob).setName(file_name + ".docx");
  // Loop through files in the auto-generated template folder to find the docx file; otherwise, set to trash
  var files = _ag_folder.getFiles();
  while (files.hasNext()){
    var file = files.next();
    if (file.getName() == file_name + ".docx") {
      var nid = file.getId();
    } else {
      file.setTrashed(true);
    };
  };
  // Download file
  var url = "https://drive.google.com/uc?id=" + nid + "&export=download";
  var html = "<script>window.open('" + url + "');google.script.host.close();</script>";
  var interface = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(interface, "Finished. Downloading file...");  
  // Put file in trash when finished
  DriveApp.getFileById(nid).setTrashed(true);
};


////////////////////////////////////////////////////////////////////////////////////////////////////
// Other Functions
////////////////////////////////////////////////////////////////////////////////////////////////////


function xxxxxxxxxxxxxxxxxxxxxxxxxxxxx() {
};


function pullUEI() {
  var values = getVariables1();
  // Initial error check
  consoleHeader("Initial Check");
  var errors = errorTest(values);
  if (errors[0] == 1) {
    Browser.msgBox(errors[1]);
    return;
  };
  consoleHeader("Pulling Values");
  //var contract = trimClean(sheet.getRange(arow, 1).getValue());
  // Calculate last row and save range in UEI sheet
  consoleHeader("Calculate Last Row");
  var usheet = _ga.getSheetByName("UEIs");
  var lrow = usheet.getLastRow();
  var usheet_range = usheet.getRange("A1:C" + lrow).getValues();
  // Loop through all contracts and save UEI if there is one
  consoleHeader("Save UEI");
  var uei = "";
  for (var i = 0; i < lrow; i++) {
    if (trimClean(usheet_range[i][0]) == trimClean(values[0].split(" ")[0])) {
      var uei = usheet_range[i][2];
      break;
    };
  };
  // Send email if contract was found, exit otherwise
  consoleHeader("Send Email");
  if (uei != "") {
    MailApp.sendEmail({to: "1@email.gov",
                      //cc: "ugh@derp.com",
                      subject: uei
                      //body: ""
                      });
    Browser.msgBox("Email sent...");
  } else {
    Browser.msgBox('Contract has not been found in "UEIs" sheet. Please add contract with associated UEI on "UEIs" sheet and try again...');
    return;
  };
};


// Survey Request 
function surveyRequest() {
  var values = getVariables1();
  // Initial error check
  consoleHeader("Initial Check");
  var errors = errorTest(values);
  if (errors[0] == 1) {
    Browser.msgBox(errors[1]);
    return;
  };
  var cor = parseNameEmail(values[7].trim().split("\n"));
  // Create draft email
  consoleHeader("Draft Email");
  var body = "Good Afternoon Vanassa,<br><br>";
  body += "Will you please send a survey request for the following action?<br><br>";
  body += "Modification: " + values[0] + " P25XXX<br>";
  body += "Modification Purpose: XXXXXXXXXX<br>";
  body += "For: " + values[3] + "<br>";
  body += "Contractor: " + values[4] + "<br>";
  body += "CS: " + parseFullName() + "<br>";
  body += "Email Survey to COR: " + cor[0] + "<br><br>";
  body += "Thank you and please let me know if you have any questions.<br><br>";
  // Signature
  var signature = "<font color='darkgray';>--<br>";
  signature += parseFullName() + "<br>";
  signature += "NOAA, AGO<br>";
  signature += "Eastern Acquisition Division</font>";
  // Create draft
  GmailApp.createDraft("1@email.gov",
                       "Survey Request for " + values[0] + " P25XXX",
                       "",
                       {cc: "",
                        htmlBody: body + signature
  });
  Browser.msgBox('Email drafted - check your "Drafts" folder in Gmail. Please update the modification number and purpose before sending...');
};


////////////////////////////////////////////////////////////////////////////////////////////////////
// Accessory Functions
////////////////////////////////////////////////////////////////////////////////////////////////////


function xxxxxxxxxxxxxxxxxxxxxxxxxxxx() {
};


// Add items to UI menu
function menuFunctions() {
  var menu = SpreadsheetApp.getUi().createMenu("Templates and Functions");
  menu.addItem("__________ Templates __________", "doNothing");
  menu.addItem("120-Day Letter Template", "generate120DayLetter");
  menu.addItem("60-Day Letter Template", "generate60DayLetter");
  menu.addItem("-8 Intent Letter Template", "generateDash8Letter");
  menu.addItem("COR Delegation Letter Template", "generateCORDelegation");
  menu.addItem("Modification Memorandum Template", "generateModMemorandum");
  menu.addItem("__________ Requests __________", "doNothing");
  menu.addItem("Vendor Report Request", "pullUEI");
  menu.addItem("Customer Survey Request", "surveyRequest");
  menu.addItem("__________ Notification Dates __________", "doNothing");
  menu.addItem("Add Notification Dates to Row", "addDates");
  menu.addItem("Add Notification Dates to Calendar", "addDatesCalendar");
  menu.addToUi();
};


function doNothing() {
};


function onOpen(e) {
  menuFunctions();
};


// Archive file and remove older files
function archiveFile() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_name = SpreadsheetApp.getActiveSpreadsheet().getName().trim();
  var dest = DriveApp.getFolderById("");
  var today = Utilities.formatDate(new Date(), "UTC", "YYYYMMdd");
  DriveApp.getFileById(sheet.getId()).makeCopy(sheet_name + " Archive " + today, dest);
  // Remove older archives
  var files = dest.getFiles();
  while (files.hasNext()) {
    // Every time you call .next(), it iterates through another file; only use it once in the while statement
    var file = files.next();
    var file_name = file.getName();
    var file_type = file.getMimeType();
    if (file_name.toLowerCase().indexOf(sheet_name.toLowerCase() + " archive") >= 0 && file_type == "application/vnd.google-apps.spreadsheet") {
      var file_date = file_name.slice(-8);
      var date_year = file_date.substring(0,4);
      var date_month = file_date.substring(4,6);
      var date_day = file_date.substring(6,8);
      // Convert to number of days to compare later
      var file_days = Number(date_year*365) + Number(date_month*30) + Number(date_day);
      var today_year = today.substring(0,4);
      var today_month = today.substring(4,6);
      var today_day = today.substring(6,8);
      // Convert to number of days to compare later
      var today_days = Number(today_year*365) + Number(today_month*30) + Number(today_day);
      // Delete older file
      if (today_days - file_days > 21) file.setTrashed(true);
    };
  };
};


function consoleHeader(txt) {
  var txt_len = txt.length;
  var len_diff = 80 - (txt_len + 2);
  console.log("#".repeat(len_diff/2) + " " + txt + " " + "#".repeat(len_diff/2));
};


function findFolderId() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  var folders = file.getParents();
  while (folders.hasNext()){
    console.log('folder name = ' + folders.next().getId());
  };
};


////////////////////////////////////////////////////////////////////////////////////////////////////
// Deprecated Code
////////////////////////////////////////////////////////////////////////////////////////////////////


/*
function pullValuesOld(row_num, co_ind) {
  _ag_folder = DriveApp.getFolderById("");
  var sheet = SpreadsheetApp.getActiveSheet();
  var contract = sheet.getRange(row_num, 1).getValue();
  var order = sheet.getRange(row_num, 2).getValue();
  if (order.toLowerCase() != "n/a") {
    contract = contract + " " + order;
    var contract_type = "Order";
  } else {
    var contract_type = "Contract";
  };
  var description = sheet.getRange(row_num, 3).getValue();
  var contractor = sheet.getRange(row_num, 4).getValue();
  var customer = sheet.getRange(row_num, 5).getValue();
  if (customer.indexOf("NCCOS") >= 0) {
    customer = "National Centers for Coastal Ocean Science (NCCOS)";
  };
  var type = sheet.getRange(row_num, 6).getValue();
  var total_value = sheet.getRange(row_num, 9).getValue();
  var pop = sheet.getRange(row_num, 10).getValue();
  var cor_poc = sheet.getRange(row_num, 16).getValue();
  var contractor_poc = sheet.getRange(row_num, 17).getValue();
  if (co_ind == 0) {
    var contracting_officer = Browser.inputBox("Enter Contracting Officer Name");
  } else {
    var contracting_officer = 1;
  };
  return ["", // 0
          contract, // 1
          order, // 2 XXXX
          contract_type, // 3
          description, // 4
          contractor, // 5
          customer, // 6
          type, // 7 XXXXX
          total_value, // 8 XXXXX
          pop, // 9
          cor_poc, // 10
          contractor_poc, // 11
          contracting_officer // 12
  ];
};


function generate120DayLetterOld() {
  // Get active row
  var arow = SpreadsheetApp.getSelection().getCurrentCell().getRow();
  // Initial check
  consoleHeader("Initial Check");
  var errors = errorTest(arow);
  if (errors[0] == 1) {
    Browser.msgBox(errors[1]);
    return;
  };
  // Pull values
  consoleHeader("Pulling Values");
  var values = pullValues(arow, 0);
  consoleHeader("Check PoP");
  var pop = parsePOP(values[9].split("\n"));
  if (pop == undefined) {
    Browser.msgBox("Cell does not appear to have option years available. Make sure the Period of Performance cell has option years in the future. Exiting...");
    return;
  } else if (pop[0] == "Error") {
    Browser.msgBox(pop[1]);
    return;
  } else if (pop.length > 7) {
    Browser.msgBox("Cell does not appear to have option years. Make sure the Period of Performance cell has option years specified. Exiting...");
    return;
  };
  // Open word doc and rename
  // All documents will be saved in root folder
  consoleHeader("Creating Word Doc");
  DriveApp.getFileById("").makeCopy(_ag_folder)
  var newname = "120-Day Letter Option " + pop[0] + " " + values[1];
  DriveApp.getFilesByName("Copy of 120-Day Letter Template").next().setName(newname);
  // Loop through files in current folder to find the copied generated previously
  var files = _ag_folder.getFiles();
  var nid = "";
  // Loop through folder till you find the new file, then save it's Id
  while (files.hasNext()){
    file = files.next();
    // Only save Id if its the copied file
    if (file.getName() == newname) {
      nid = file.getId();
    };
  };
  // Open doc and replace fields with pulled data
  consoleHeader("Inputting Values");
  var doc = DocumentApp.openById(nid);
  var body = doc.getBody();
  body.replaceText("{Date}", Utilities.formatDate(new Date(), "EST", "MMMM d, yyyy"));
  body.replaceText("{COR}", parseNameEmail(values[10].split("\n"))[0]);
  body.replaceText("{ContractingOfficer}", values[12]);
  body.replaceText("{Contract}", values[1]);
  body.replaceText("{Description}", values[4]);
  body.replaceText("{ContractType}", values[3]);
  body.replaceText("{Contractor}", values[5]);
  body.replaceText("{Customer}", values[6]);
  body.replaceText("{Expiration}", pop[6]);
  body.replaceText("{OptionYear}", pop[0]);
  body.replaceText("{POPStart}", pop[3]);
  body.replaceText("{POPEnd}", pop[4]);
  body.replaceText("{Return}", Utilities.formatDate(new Date(new Date().getTime() + 3600000*24*30), "EST", "MMMM d, yyyy"));
  body.replaceText("{ContactName}", parseFullName());
  body.replaceText("{ContactEmail}", Session.getActiveUser().getEmail());
  doc.saveAndClose();
  consoleHeader("Finished");
  Browser.msgBox("Letter drafted...");
  downloadFile(nid, newname + ".docx");
};


function generate60DayLetterOld() {
  // Get active row
  var arow = SpreadsheetApp.getSelection().getCurrentCell().getRow();
  // Initial check
  consoleHeader("Initial Check");
  var errors = errorTest(arow);
  if (errors[0] == 1) {
    Browser.msgBox(errors[1]);
    return;
  };  
  consoleHeader("Pulling Values");
  var values = pullValues(arow, 0);
  consoleHeader("Check PoP");
  var pop = parsePOP(values[9].split("\n"));
  if (pop == undefined) {
    Browser.msgBox("Cell does not appear to have option years available. Make sure the Period of Performance cell has option years in the future. Exiting...");
    return;
  } else if (pop[0] == "Error") {
    Browser.msgBox(pop[1]);
    return;
  } else if (pop.length > 7) {
    Browser.msgBox("Cell does not appear to have option years. Make sure the Period of Performance cell has option years specified. Exiting...");
    return;
  };
  var contractor_full = parseNameEmail(values[11].split("\n"));
  var contractor_fullname = contractor_full[0].split(" ");
  // Open word doc and rename
  // All documents will be saved in root folder
  consoleHeader("Creating Word Doc");
  DriveApp.getFileById("").makeCopy(_ag_folder)  
  var newname = "60-Day Letter Option " + pop[0] + " " + values[1];
  DriveApp.getFilesByName("Copy of 60-Day Letter Template").next().setName(newname);
  // Loop through files in current folder to find the copied generated previously
  var files = _ag_folder.getFiles();
  var nid = "";
  // Loop through folder till you find the new file, then save it's Id
  while (files.hasNext()){
    file = files.next();
    // Only save Id if its the copied file
    if (file.getName() == newname) {
      nid = file.getId();
    };
  };
  // Open doc and replace fields with pulled data
  consoleHeader("Inputting Values");
  var doc = DocumentApp.openById(nid);
  var body = doc.getBody();
  body.replaceText("{Date}", Utilities.formatDate(new Date(), "EST", "MMMM d, yyyy"));
  body.replaceText("{ContractingOfficer}", values[12]);
  body.replaceText("{Contract}", values[1]);
  body.replaceText("{Description}", values[4]);
  body.replaceText("{ContractType}", values[3]);
  body.replaceText("{Contractor}", values[5]);
  body.replaceText("{ContractorPOC}", contractor_full[0]);
  body.replaceText("{ContractorPOCEmail}", contractor_full[1]);
  body.replaceText("{ContractorGreeting}", contractor_fullname[contractor_fullname.length - 1]);
  body.replaceText("{OptionYear}", pop[0]);
  body.replaceText("{ContactName}", parseFullName());
  body.replaceText("{ContactEmail}", Session.getActiveUser().getEmail());
  doc.saveAndClose();
  consoleHeader("Finished");
  Browser.msgBox("Letter drafted...");
  redirectFolder();
};


function generateDash8LetterOld() {
// Get active row
  var arow = SpreadsheetApp.getSelection().getCurrentCell().getRow();
  // Initial check
  consoleHeader("Initial Check");
  var errors = errorTest(arow);
  if (errors[0] == 1) {
    Browser.msgBox(errors[1]);
    return;
  };
  consoleHeader("Pulling Values");
  var values = pullValues(arow, 0);
  var contractor_full = parseNameEmail(values[11].split("\n"));
  var contractor_fullname = contractor_full[0].split(" ");
  // Open word doc and rename
  // All documents will be saved in root folder
  consoleHeader("Creating Word Doc");
  DriveApp.getFileById("").makeCopy(_ag_folder)  
  var newname = "-8 Intent Letter " + values[1];
  DriveApp.getFilesByName("Copy of -8 Intent Letter Template").next().setName(newname);
  // Loop through files in current folder to find the copied generated previously
  var files = _ag_folder.getFiles();
  var nid = "";
  // Loop through folder till you find the new file, then save it's Id
  while (files.hasNext()){
    file = files.next();
    // Only save Id if its the copied file
    if (file.getName() == newname) {
      nid = file.getId();
    };
  };
  // Open doc and replace fields with pulled data
  consoleHeader("Inputting Values");
  var doc = DocumentApp.openById(nid);
  var body = doc.getBody();
  body.replaceText("{Date}", Utilities.formatDate(new Date(), "EST", "MMMM d, yyyy"));
  body.replaceText("{ContractingOfficer}", values[12]);
  body.replaceText("{Contract}", values[1]);
  body.replaceText("{Description}", values[4]);
  body.replaceText("{ContractType}", values[3]);
  body.replaceText("{Contractor}", values[5]);
  body.replaceText("{ContractorPOC}", contractor_full[0]);
  body.replaceText("{ContractorPOCEmail}", contractor_full[1]);
  body.replaceText("{ContractorGreeting}", contractor_fullname[contractor_fullname.length - 1]);
  body.replaceText("{ContactName}", parseFullName());
  body.replaceText("{ContactEmail}", Session.getActiveUser().getEmail());
  doc.saveAndClose();
  consoleHeader("Finished");
  Browser.msgBox("Letter drafted...");
  redirectFolder();
};


function generateCORDelegationOld() {
  // Get active row
  var arow = SpreadsheetApp.getSelection().getCurrentCell().getRow();
  // Initial check
  consoleHeader("Initial Check");
  var errors = errorTest(arow);
  if (errors[0] == 1) {
    Browser.msgBox(errors[1]);
    return;
  };
  consoleHeader("Pulling Values");
  var values = pullValues(arow, 0);
  // Open word doc and rename
  consoleHeader("Creating Word Doc");
  DriveApp.getFileById("").makeCopy(_ag_folder) 
  var newname = "COR Delegation Letter " + values[1];
  DriveApp.getFilesByName("Copy of COR Delegation Letter Template").next().setName(newname);
  // Loop through files in current folder to find the copied generated previously
  var files = _ag_folder.getFiles();
  var nid = ""
  // Loop through folder till you find the new file, then save it's Id
  while (files.hasNext()){
    file = files.next();
    // Only save Id if its the copied file
    if (file.getName() == newname) {
      nid = file.getId();
    };
  };
  // Open doc and replace fields with pulled data
  consoleHeader("Inputting Values");
  var doc = DocumentApp.openById(nid);
  var body = doc.getBody();
  body.replaceText("{Date}", Utilities.formatDate(new Date(), "EST", "MMMM d, yyyy"));
  body.replaceText("{COR}", parseNameEmail(values[10].split("\n"))[0]);
  body.replaceText("{ContractingOfficer}", values[12]);
  body.replaceText("{Contract}", values[1]);
  body.replaceText("{Description}", values[4]);
  body.replaceText("{ContractType}", values[3]);
  body.replaceText("{Contractor}", values[5]);
  body.replaceText("{Customer}", values[6]);
  doc.saveAndClose();
  consoleHeader("Finished");
  Browser.msgBox("Letter drafted...");
  redirectFolder();
};


function generateModMemorandumOld() {
  // Get active row
  var arow = SpreadsheetApp.getSelection().getCurrentCell().getRow();
  // Initial check
  consoleHeader("Initial Check");
  var errors = errorTest(arow);
  if (errors[0] == 1) {
    Browser.msgBox(errors[1]);
    return;
  };
  // Pull values
  consoleHeader("Pulling Values");
  var values = pullValues(arow, 0);
  consoleHeader("Check PoP");
  var pop = parseEntirePOP(values[9].split("\n"));
  if (pop[0] == "Error") {
    Browser.msgBox(pop[1]);
    return;
  };
  // Open word doc and rename
  // All documents will be saved in root folder
  consoleHeader("Creating Word Doc");
  DriveApp.getFileById("").makeCopy(_ag_folder)
  var newname = "Modification Memorandum " + values[1];
  DriveApp.getFilesByName("Copy of Modification Memorandum Template").next().setName(newname);
  // Loop through files in current folder to find the copied generated previously
  var files = _ag_folder.getFiles();
  var nid = "";
  // Loop through folder till you find the new file, then save it's Id
  while (files.hasNext()){
    file = files.next();
    // Only save Id if its the copied file
    if (file.getName() == newname) {
      nid = file.getId();
    };
  };
  // Open doc and replace fields with pulled data
  consoleHeader("Inputting Values");
  var doc = DocumentApp.openById(nid);
  var body = doc.getBody();
  body.replaceText("{Date}", Utilities.formatDate(new Date(), "EST", "MMMM d, yyyy"));
  body.replaceText("{ContractingOfficer}", values[12]);
  body.replaceText("{Contract}", values[1]);
  body.replaceText("{Description}", values[4]);
  body.replaceText("{ContractType}", values[3]);
  body.replaceText("{Contractor}", values[5]);
  body.replaceText("{Customer}", values[6]);
  body.replaceText("{PeriodOfPerformance}", pop);
  body.replaceText("{ContactName}", parseFullName());
  doc.saveAndClose();
  consoleHeader("Finished");
  Browser.msgBox("Memorandum drafted...");
  redirectFolder();
};


function pullUEIOld() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var arow = SpreadsheetApp.getSelection().getCurrentCell().getRow();
  // Initial check
  consoleHeader("Initial Check");
  var errors = errorTest(arow);
  if (errors[0] == 1) {
    Browser.msgBox(errors[1]);
    return;
  };  
  consoleHeader("Pulling Values");
  var contract = trimClean(sheet.getRange(arow, 1).getValue());
  // Calculate last row and save range in UEI sheet
  consoleHeader("Calculate Last Row");
  var dsheet = SpreadsheetApp.getActive().getSheetByName("UEIs");
  var lrow = dsheet.getDataRange().getLastRow();
  var dsheet_range = dsheet.getRange("A1:C" + lrow).getValues();
  // Loop through all contracts and save UEI if there is one
  consoleHeader("Save UEI");
  var uei = "";
  for (var i = 0; i < lrow; i++) {
    if (trimClean(dsheet_range[i][0]) == contract) {
      var uei = dsheet_range[i][2];
    };
  };
  // Send email if contract was found, exit otherwise
  consoleHeader("Send Email");
  if (uei != "") {
    MailApp.sendEmail({to: "vendorreport@feddatacheck.net",
                      //cc: "ugh@derp.com",
                      subject: uei,
                      //body: ""
                      });
    Browser.msgBox("Email sent...");
  } else {
    Browser.msgBox('Contract has not been found in "UEIs" sheet. Please add contract with associated UEI on "UEIs" sheet and try again...');
    return;
  };
};


// Parse Period of Performance and option year
function parsePOPOld(opop) {
  consoleHeader("Parsing POP");
  var oy = -1;
  var exdate = "";
  // Proceed if only one line for PoP; returns pretty start and end date
  if (opop.length == 1) {
    return parsePOPOneLine(opop);
  } else {
    for (var i = 0; i < opop.length; i++) {
      oy += 1;
      // Remove all whitespace
      //var npop = opop[i].trim();
      //if (npop.indexOf(":") > 0) {
      //  var fpop = npop.substring(npop.indexOf(":") + 1,npop.lenth).split("-");
      //} else {
      //  var fpop = npop.split("-");
      //};
      var s_index = opop[i].indexOf(opop[i].match(/\d\//));
      if (s_index == -1) {
        return ["Error", 'POP not detected - make sure it is in the format "MM/DD/YYYY - MM/DD/YYYY"; for example: "09/30/2023 to 09/29/2024"'];
      };
      var upd_date = opop[i].substring(s_index - 1, opop[i].length).trim();
      var npop = upd_date.split("-");
      var spop = npop[0].trim();
      var epop = npop[1].trim();
      // Convert strings to date format
      var tdate = Utilities.formatDate(new Date(), "EST", "M/d/yyyy").split("/");
      var sdate = Utilities.formatDate(new Date(new Date(spop).getTime() + 3600000), "EST", "M/d/yyyy").split("/");
      // For when we return the values, need to add one hour to the time for some reason to show properly
      var sdate2 = Utilities.formatDate(new Date(new Date(spop).getTime() + 3600000), "EST", "MMMM d, yyyy");
      var edate = Utilities.formatDate(new Date(new Date(epop).getTime() + 3600000), "EST", "MMMM d, yyyy");
      // Must convert dates to number of days in order to compare them
      var tdated = Number(tdate[0]*30) + Number(tdate[1]) + Number(tdate[2]*365);
      var sdated = Number(sdate[0]*30) + Number(sdate[1]) + Number(sdate[2]*365);
      // Return the first date that is later than today, meaning the first option year
      if (sdated > tdated) {
        return [numberText(oy), spop, epop, sdate2, edate, sdate2 + " through " + edate, exdate];
      } else {
        // Save the expiration date as last thing before going to next year
        exdate = Utilities.formatDate(new Date(new Date(epop).getTime() + 3600000), "EST", "MMMM d, yyyy");
        continue;
      };
    };
  };
};


// Only run for POPs with one line
function parsePOPOneLine(opop) {
  // Find first instance of a digit followed by a "/", denoting the month
  var s_index = opop[0].indexOf(opop[0].match(/\d\//));
  // Stop if POP not detected
  if (s_index == -1) {
    return ["Error", 'POP not detected - make sure it is in the format "MM/DD/YYYY - MM/DD/YYYY"; for example: "09/30/2023 to 09/29/2024"'];
  }
  var upd_date = opop[0].substring(s_index - 1, opop[0].length).trim();
  var npop = upd_date.split("-");
  var spop = npop[0].trim();
  var epop = npop[1].trim();
  var sdate = Utilities.formatDate(new Date(new Date(spop).getTime() + 3600000), "EST", "MMMM d, yyyy");
  var edate = Utilities.formatDate(new Date(new Date(epop).getTime() + 3600000), "EST", "MMMM d, yyyy");
  return sdate + " through " + edate;
};


// Parse Entire Period of Performace
function parseEntirePOP(opop) {
  consoleHeader("Parsing POP");
  // Proceed if only one line for PoP; returns pretty start and end date
  if (opop.length == 1) {
    return parsePOPOneLine(opop);
  } else {
    for (var i = 0; i < opop.length; i++) {
      var s_index = opop[i].indexOf(opop[i].match(/\d\//));
      if (s_index == -1) {
        return ["Error", 'POP not detected - make sure it is in the format "M/D/YYYY - M/D/YYYY"; for example: "9/10/2023 to 9/9/2024"'];
      };
      var upd_date = opop[i].substring(s_index - 1, opop[i].length).trim();
      var npop = upd_date.split("-");
      var spop = npop[0].trim();
      var epop = npop[1].trim();
      // For when we return the values, need to add one hour to the time for some reason to show properly
      // Save just for starting year
      if (i == 0) {
        var estdate = Utilities.formatDate(new Date(new Date(spop).getTime() + 3600000), "EST", "MMMM d, yyyy");
      };
      var edate = Utilities.formatDate(new Date(new Date(epop).getTime() + 3600000), "EST", "MMMM d, yyyy");
    };
  };
  return estdate + " through " + edate;
};


// DEPRECATED
// Parse Period of Performance for COR Delegation Letters
// Returns start date pretty, end date pretty
function parsePOPCOR(opop) {
  console.log("#################### Parsing POP ####################")
  var oy = -1;
  if (opop.length = 1) {
    npop = opop[0].replace(/\s/g, "").split("-");
    var sdate = Utilities.formatDate(new Date(new Date(npop[0]).getTime() + 3600000), "EST", "MMMM d, yyyy");
    var edate = Utilities.formatDate(new Date(new Date(npop[1]).getTime() + 3600000), "EST", "MMMM d, yyyy");
  } else {
    for (var i = 0; i < opop.length; i++) {
      oy += 1;
      // Remove all whitespace
      var npop = opop[i].replace(/\s/g, "");
      if (npop.indexOf(":") > 0) {
        var fpop = npop.substring(npop.indexOf(":") + 1,npop.lenth).split("-");
      } else {
        var fpop = npop.split("-");
      };
      // Convert strings to date format
      if (oy == 0) {
        var sdate = Utilities.formatDate(new Date(new Date(fpop[0]).getTime() + 3600000), "EST", "MMMM d, yyyy");
      } else {
        var edate = Utilities.formatDate(new Date(new Date(fpop[1]).getTime() + 3600000), "EST", "MMMM d, yyyy");
      };
    };
  };
  return sdate + " through " + edate;
};


function redirectFolder() {
  var url = "https://drive.google.com/drive/folders/XXXXXXXXXX";
  var html = "<script>window.open('" + url + "');google.script.host.close();</script>";
  var userInterface = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Redirecting...');
};


function surveyRequestOld() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var arow = SpreadsheetApp.getSelection().getCurrentCell().getRow();
  // Initial check
  consoleHeader("Initial Check");
  var errors = errorTest(arow);
  if (errors[0] == 1) {
    Browser.msgBox(errors[1]);
    return;
  }; 
  consoleHeader("Pulling Values");
  var values = pullValues(arow, 1);
  var cor = parseNameEmail(values[10].split("\n"));
  // Find latest modification
  consoleHeader("Find Latest Modification");
  var mods = sheet.getRange(arow, 18).getValue().split("\n");
  var mod_desc = [];
  for (var i = mods.length - 1; i >= 0; i--) {
    var cl_text = trimClean(mods[i]);
    var w_text = mods[i].trim();
    if (cl_text.substring(0, 1) == "p") {
      mod_desc.push(w_text.substring(0, 6));
      if (w_text.split(":").length > 1) {
        mod_desc.push(w_text.split(":")[1].trim());
      } else if (w_text.split("-").length > 1) {
        mod_desc.push(w_text.split("-")[1].trim());
      } else {
        Browser.msgBox('Could not parse any modification description. Please make sure modification numbers have a ":" or "-" after it to denote the description. For example, "P24001: COR Change" or "P24003 - Option 1 Exercise". Exiting...');
        return;
      };
      break;
    };
  };
  // Create draft email
  consoleHeader("Draft Email");
  var embody = "Good Afternoon Vanassa,\n\nWill you please send a survey request for the following action?\n\n";
  embody = embody + "Modification: " + values[1] + " " + mod_desc[0] + "\n";
  embody = embody + "Modification Purpose: " + mod_desc[1] + "\n";
  embody = embody + "For: " + values[6] + "\n";
  embody = embody + "Contractor: " + values[5] + "\n";
  embody = embody + "CS: " + parseFullName() + "\n";
  embody = embody + "Email Survey to COR: " + cor[0] + ", " + cor[1] + "\n\n";
  embody = embody + "Thank you and please let me know if you have any questions.\n\n";
  var signature = parseFullName() + "\n";
  signature = signature + "NOAA, AGO\n";
  signature = signature + "Eastern Acquisition Division\n";
  // Create draft
  GmailApp.createDraft("vanassa.m.barner@noaa.gov",
                       "Survey Request for " + values[1] + " " + mod_desc[0],
                       embody + signature
  );
  Browser.msgBox('Email drafted - check your "Drafts" folder in Gmail. Exiting...');
};


// Works
function dev005() {
  var strng = "OLD BPA";
  // does not work
  var regex = /^(?=.*[A-Za-z])(?=.*[0-9])[A-Za-z0-9]+$/;
  // works
  var regex2 = /^(.*[0-90-90-9].*)$/;
  // works, finds any two digits in a row
  var regex3 = /^(.*\d{2}.*)$/;
  console.log((strng.trim().match(regex2)));
  if (!strng.trim().match(regex3)) {
    console.log(1);
  } else {
    console.log(0);
  };
  //console.log(strng.match(/()*));
};
*/

