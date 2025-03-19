

// @NotOnlyCurrentDoc


// Save all variables to be used throughout
function getVariables(dev) {
  if (dev == undefined ) {
    // File IDs
    _file_ww = "";
    _file_nos = "";
    _file_noskc = "";
    _file_nws = "";
    _file_nwskc = "";
    _file_omao = "";
    _file_sap = "";
    _file_sapkc = "";
    //_file_closeout = "";
    _file_cw = "";
    // Archive folder IDs
    _folder_ww = "";
    _folder_nos = "";
    _folder_noskc = "";
    _folder_nws = "";
    _folder_nwskc = "";
    _folder_omao = "";
    _folder_sap = "";
    _folder_sapkc = "";
    //_folder_closeout = "";
    _folder_cw = "";
    // WIP folder IDs
    _wip_folder_nos = "";
    _wip_folder_noskc = "";
    _wip_folder_nws = "";
    _wip_folder_nwskc = "";
    _wip_folder_omao = "";
    _wip_folder_sap = "";
    _wip_folder_sapkc = "";
    _wip_folder_cw = "";
    _folder_ago_wip = "";
  // Dev operations
  } else {
    // File IDs
    _file_ww = "";
    _file_nos = "";
    _file_noskc = "";
    _file_nws = "";
    _file_nwskc = "";
    _file_omao = "";
    _file_sap = "";
    _file_sapkc = "";
    //_file_closeout = "";
    _file_cw = "";
    // Archive folder IDs
    _folder_ww = "";
    _folder_nos = "";
    _folder_noskc = "";
    _folder_nws = "";
    _folder_nwskc = "";
    _folder_omao = "";
    _folder_sap = "";
    _folder_sapkc = "";
    //_folder_closeout = "";
    _folder_cw = "";
    // WIP folder IDs
    _wip_folder_nos = "";
    _wip_folder_noskc = "";
    _wip_folder_nws = "";
    _wip_folder_nwskc = "";
    _wip_folder_omao = "";
    _wip_folder_sap = "";
    _wip_folder_sapkc = "";
    _wip_folder_cw = "";
    _folder_ago_wip = "";
  };
  // Working WIP
  _ga = SpreadsheetApp.openById(_file_ww);
  _ga_raw = _ga.getSheetByName('Raw AGO WIP');
  _ga_rew = _ga.getSheetByName('Raw EAD WIP');
  _ga_cw = _ga.getSheetByName('Current WIP');
  _ga_rf = _ga.getSheetByName('Raw Forecasting');
  _ga_rnw = _ga.getSheetByName('Requisitions Need Workspaces');
  _ga_di = _ga.getSheetByName('Duplicate Items');
  _ga_dv = _ga.getSheetByName('Data Values');
  // Consolidated WIP
  _gac = SpreadsheetApp.openById(_file_cw);
  _gac_e = _gac.getSheetByName('EAD Summary by Specialist');
  _gac_ec = _gac.getSheetByName("NOS and OMAO Summary");
  _gac_nd = _gac.getSheetByName("NWS Summary");
  _gac_sb = _gac.getSheetByName('Summary by Branch and Type');
  _gac_ss = _gac.getSheetByName('Summary by Specialist and Type');
  _gac_w = _gac.getSheetByName('WIP');
  _gac_d = _gac.getSheetByName('Data Values');
  _gac_hd = _gac.getSheetByName('Historical Data');
  // WIPs table
  _wips = [['NOS', _file_nos, _folder_nos],
           ['NOS KC', _file_noskc, _folder_noskc],
           ['NWS', _file_nws, _folder_nws],
           ['NWS KC', _file_nwskc, _folder_nwskc],
           ['OMAO', _file_omao, _folder_omao],
           ['SAP', _file_sap, _folder_sap],
           ['SAP KC', _file_sapkc, _folder_sapkc]
  ];
};


////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////// Run All ////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////


function runAll() {
  var st = Date.now();
  getVariables();
  allFunctions();
  endTime(st);
};


function runAllDev() {
  var st = Date.now();
  getVariables(1);
  allFunctions();
  endTime(st);
};


function allFunctions() {
  updateWIP();
  pullData();
  dupeCheck();
  compareData();
  updateAllBranchWIPs();
  cwUpdateWIP();
  cwUpdatePivots();
  draftEmail();
};


////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////// Test /////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////


// Pulls data from WIP, run after runAll and runAllDev
function copyDataToDev() {
  // Save data
  getVariables();
  var next_row = 1;
  var next_dv = 1;
  var dv = next_dv;
  var dv_row = "";
  for (var i = 0; i < _wips.length; i++) {
    dv += 1;
    dv_row = _ga_dv.getRange("P" + dv).getValue();
    _wips[i][2] = _ga_cw.getRange("A" + (next_row + 1) + ":V" + (dv_row + next_row)).getValues();
    next_row += dv_row;
  };
  var wips = _wips;
  // Copy data to DEV
  getVariables(1);
  _wips[0][2] = wips[0][2];
  _wips[1][2] = wips[1][2];
  _wips[2][2] = wips[2][2];
  _wips[3][2] = wips[3][2];
  _wips[4][2] = wips[4][2];
  _wips[5][2] = wips[5][2];
  _wips[6][2] = wips[6][2];
  for (i = 0; i < _wips.length; i++) {
    insertData(_wips[i][2], SpreadsheetApp.openById(_wips[i][1]).getSheetByName("WIP"));
  };
};


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////// Step 1 ////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function updateWIP() {
  var st = Date.now();
  SpreadsheetApp.flush();
  consoleHeader('Step 1: Update "Working WIP"');
  ////////////////////////////////////////////////////////////////////////////////
  console.log('Archive file');
  archiveFile(_file_ww, _folder_ww);
  console.log('Delete data');
  SpreadsheetApp.flush();
  clearData(_ga_raw);
  SpreadsheetApp.flush();
  ////////////////////////////////////////////////////////////////////////////////
  console.log('Identify the latest file');
  var dest = DriveApp.getFolderById(_folder_ago_wip);
  // Remove older archives
  var files = dest.getFiles();
  var latest_date = 0;
  while (files.hasNext()) {
    // Every time you call .next(), it iterates through another file; only use it once in the while statement
    var file = files.next();
    var file_name = file.getName().trim().replace(/-/g, '').replace(/\s/g, '');
    var loc = file_name.indexOf((new Date()).getFullYear());
    var date_file = file_name.substring(loc, loc + 8);
    // Only save information from latest file
    if (date_file > latest_date) {
      latest_date = date_file;
      var latest_id = file.getId();
    };
  };
  ////////////////////////////////////////////////////////////////////////////////
  // Save data from latest file
  console.log('Copy all data and save to "Raw WIP"');
  console.log("Using File ID: " + latest_id + "\nFile name: " + SpreadsheetApp.openById(latest_id).getName());
  var wip_dv = getDataValues(SpreadsheetApp.openById(latest_id).getSheetByName('workinprogress'), 0, 0);
  console.log(wip_dv.length);
  insertData(wip_dv, _ga_raw);
  // Format data
  formatDataSheet(_ga_raw);
  endTime(st);
};


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////// Step 2 ////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function pullData() {
  var st = Date.now();
  SpreadsheetApp.flush();
  consoleHeader('Step 2: Pull data from "Raw WIP"');
  var raw_forecasting = [];
  var raw_no_workspaces = [];
  var raw_wip = [];
  var data_rw = getDataValues(_ga_raw, 0, 0);
  ////////////////////////////////////////////////////////////////////////////////
  console.log('Separate data to three lists');
  for (var i = 0; i < data_rw.length; i++) {
    var copoc_group = data_rw[i][20].toLowerCase().trim().substring(0,6);
    var buyer_group = data_rw[i][9].toLowerCase().trim().substring(0,6);
    var workspace = data_rw[i][13];
    var requisition = data_rw[i][2];
    // Only pull data that belongs to an EAD group
    if (buyer_group == 'ead - ' || (buyer_group != 'ead - ' && copoc_group == 'ead - ')) {
      // Send to forecasting list
      if (data_rw[i][17] == 'Forecasting') {
        raw_forecasting.push(data_rw[i]);
      // Send to no workspaces list
      } else if (workspace == '' && requisition != '') {
        raw_no_workspaces.push(data_rw[i]);
      // Send to regular workspaces
      } else if (workspace != '') {
        raw_wip.push(data_rw[i]);
      };
    };
  };
  ////////////////////////////////////////////////////////////////////////////////
  console.log('Insert data to "Forecasting"');
  SpreadsheetApp.flush();
  clearData(_ga_rf);
  SpreadsheetApp.flush();
  var dv_list = vLookupList(1, 2);
  for (i = 0; i < raw_forecasting.length; i++) {
    // Add Branch
    if (raw_forecasting[i][19] != '') {
      var branch_poc = raw_forecasting[i][19];
    } else if (raw_forecasting[i][19] == '' && raw_forecasting[i][8] != '') {
      var branch_poc = raw_forecasting[i][8];
    } else {
      var branch_poc = '';
    };
    raw_forecasting[i].push(vLookup(branch_poc, dv_list, 2, 0));  
  };
  insertData(raw_forecasting, _ga_rf);
  ////////////////////////////////////////////////////////////////////////////////
  console.log('Insert data to "Requisitions Need Workspaces"');
  SpreadsheetApp.flush();
  clearData(_ga_rnw);
  SpreadsheetApp.flush();
  var nwc_list = [];
  var dv_list = vLookupList(1, 2);
  for (i = 0; i < raw_no_workspaces.length; i++) {
    var req_num = raw_no_workspaces[i][2];
    // Branch POC to determine BRANCH, preference given to Buyer
    if (raw_no_workspaces[i][8] != '') {
      var branch_poc = raw_no_workspaces[i][8];
    } else if (raw_no_workspaces[i][8] == '' && raw_no_workspaces[i][19] != '') {
      var branch_poc = raw_no_workspaces[i][19];
    } else {
      var branch_poc = '';
    };
    // Contracts Office POC
    if (raw_no_workspaces[i][19] == '') {
      var co_poc = "(No Contracts Office POC Assigned)";
    } else {
      var co_poc = raw_no_workspaces[i][19];
    };
    // Purpose
    if (raw_no_workspaces[i][7].trim() != '') {
      var purpose = raw_no_workspaces[i][7].trim();
    } else {
      var purpose = raw_no_workspaces[i][14].trim();
    };
    // Committed amount
    if (raw_no_workspaces[i][6] == '$ -' || raw_no_workspaces[i][6] == '') {
      var comm_amt = 0;
    } else {
      var comm_amt = raw_no_workspaces[i][6];
    };
    nwc_list.push([vLookup(branch_poc, dv_list, 2, 0), // Branch
                   raw_no_workspaces[i][8], // Buyer
                   co_poc, // Contracts Office POC
                   '', // Req/Mod Workspace Number
                   req_num, // Requisition
                   '', // Contract Office POC Assign Date
                   purpose, // Purpose
                   '', // PRISM AAP
                   comm_amt, // Committed Amount
                   '', // Total Obligation
                   '', // Total Value
                   '', // Client
                   '', // Type
                   '', // PALT
                   '', // Days Since Contract Office POC Assigned
                   '', // Accrued PALT
                   '', // Special Interest
                   '', // Current Contract Expiration
                   '', // Fiscal Year
                   '', // Projected Award Date
                   '', // Process Update
                   '' // Comments/Issues
                   //String(req_num.toString() + comm_amt) // Lookup ID
    ]);
  };
  insertData(nwc_list, _ga_rnw);
  _ga_rnw.getRange('A1').getFilter().sort(5, true);
  _ga_rnw.getRange('A1').getFilter().sort(2, true);
  _ga_rnw.getRange('A1').getFilter().sort(1, false);
  ////////////////////////////////////////////////////////////////////////////////
  console.log('Insert data to "Raw EAD WIP"');
  SpreadsheetApp.flush();
  clearData(_ga_rew);
  SpreadsheetApp.flush();
  var dv_list = vLookupList(1, 2);
  for (i = 0; i < raw_wip.length; i++) {
    // Add Branch
    if (raw_wip[i][8] != '') {
      var branch_poc = raw_wip[i][8];
    } else if (raw_wip[i][8] == '' && raw_wip[i][19] != '') {
      var branch_poc = raw_wip[i][19];
    } else {
      var branch_poc = '';
    };
    raw_wip[i].push(vLookup(branch_poc, dv_list, 2, 0));
  };
  insertData(raw_wip, _ga_rew);
  endTime(st);
};


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////// Step 3 ////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function dupeCheck() {
  var st = Date.now();
  consoleHeader('3. Check for dupes from the "EAD WIP"');
  SpreadsheetApp.flush();
  // Check for dupes, post them in the "Duplicate Items" tab
  clearData(_ga_di);
  SpreadsheetApp.flush();
  // Recreate raw_wip into new list with data
  var raw_wip = getDataValues(_ga_rew, 0, 0);
  var wip_list = [];
  var workspace_list = [];
  for (var i = 0; i < raw_wip.length; i++) {
    var workspace_num = raw_wip[i][13];
    // Contracts Office POC
    if (raw_wip[i][19] == '') {
      var co_poc = "(No Contracts Office POC Assigned)";
    } else {
      var co_poc = raw_wip[i][19];
    };
    // Committed amount
    if (raw_wip[i][6] == '$ -' || raw_wip[i][6] == '') {
      var comm_amt = 0;
    } else {
      var comm_amt = raw_wip[i][6];
    };
    // PR
    if (raw_wip[i][14].trim() != '') {
      var pr_name = raw_wip[i][14].trim();
    } else {
      var pr_name = raw_wip[i][7].trim();
    };
    // Amendment numbers
    if (raw_wip[i][3] == 'ORIG' || raw_wip[i][3] == '') {
      var amendment = 0;
    } else {
      var amendment = raw_wip[i][3];
    };
    wip_list.push([raw_wip[i][raw_wip[0].length - 1], // Branch
                  raw_wip[i][8], // Buyer
                  co_poc, // Contracts Office POC
                  workspace_num, // Req/Mod Workspace Number
                  raw_wip[i][2], // Requisition
                  raw_wip[i][22], // Contract Office POC Assign Date
                  pr_name, // Purpose
                  workspace_num.replace(/-/g, '').slice(-5).padStart(5, 0).toString(), // PRISM AAP
                  comm_amt, // Committed Amount
                  '', // Total Obligation
                  '', // Total Value of Action
                  '', // Client
                  '', // Type
                  '', // PALT
                  raw_wip[i][23], // Days Since Contract Office POC Assigned
                  raw_wip[i][23], // Accrued PALT
                  '', // Special Interest
                  '', // Current Contract Expiration
                  '', // Fiscal Year
                  '', // Projected Award Date
                  '', // Process Update
                  '', // Comments/Issues
                  amendment // Amendment
    ]);
    // Push item to separate list to check for dupes
    workspace_list.push([workspace_num]);
  };
  ////////////////////////////////////////////////////////////////////////////////
  // Count totals
  var workspaces = {};
  workspace_list.forEach(num => {
    if (workspaces[num]) {
        workspaces[num] += 1;
    } else {
        workspaces[num] = 1;
    };
  });
  // Identify only duplicate workspaces
  var dupes = []
  for (let x in workspaces) {
    if (workspaces[x] > 1) {
      dupes.push([x, workspaces[x]]);
    };
  };
  // Check if dupes present, if no dupes present, push to final list
  var final_list = [];
  if (dupes.length == 0) {
    for (i = 0; i < wip_list.length; i++) {
      final_list.push(wip_list[i]);
    };
  } else {
    // Identify duplicate workspaces with additive amendment count
    var dupes_data = [];
    for (i = 0; i < dupes.length; i++) {
      var amen_count = 0;
      for (var j = 0; j < wip_list.length; j++) {
        // Add all amendments into one number
        if (dupes[i][0] == wip_list[j][3]) {
          amen_count += wip_list[j][22];
        };
      };
      dupes_data.push([dupes[i][0], amen_count]);
    };
    ////////////////////////////////////////////////////////////////////////////////
    // List all dupes on "Duplicate Items"
    console.log('List dupes on "Duplicate Items"');
    SpreadsheetApp.flush();
    clearData(_ga_di);
    SpreadsheetApp.flush();
    var ga_di_values = [];
    for (i = 0; i < dupes_data.length; i++) {
      for (j = 0; j < wip_list.length; j++) {
        if (dupes_data[i][0] == wip_list[j][3]) {
          ga_di_values.push([wip_list[j][0], // Branch
                             wip_list[j][1], // Buyer
                             wip_list[j][2], // Contracts Office POC
                             wip_list[j][3], // Req/Mod Workspace Number
                             wip_list[j][4], // Requisition
                             wip_list[j][22], // Amendment
                             wip_list[j][6], // Purpose
                             wip_list[j][8] // Committed Amount
          ]);
        };
      };
    };
    insertData(ga_di_values, _ga_di);
    _ga_di.getRange('A1').getFilter().sort(4, true);
    _ga_di.getRange('A1').getFilter().sort(3, true);
    _ga_di.getRange('A1').getFilter().sort(1, false);
    // First create new list with data that are not duplicates
    ////////////////////////////////////////////////////////////////
    console.log('Remove dupes from the wip_list');
    for (i = 0; i < wip_list.length; i++) {
      var dupe_ind = 0;
      for (j = 0; j < dupes_data.length; j++) {
        if (wip_list[i][3] == dupes_data[j][0]) {
          dupe_ind = 1;
        };
      };
      // Pushes non-dupe data to final list
      if (dupe_ind == 0) {
        // Push new data, remove last column of data
        final_list.push(wip_list[i].slice(0, -1));
      };
    };
    // Add specific dupe data back into data w/o dupes
    for (i = 0; i < dupes_data.length; i++) {
      // Add data with no amendments
      if (dupes_data[i][1] == 0) {
        var comm_total = 0;
        var updated_item = '';
        for (j = 0; j < wip_list.length; j++) {
          if (dupes_data[i][0] == wip_list[j][3]) {
            comm_total += wip_list[j][8];
            updated_item = wip_list[j];
            updated_item[8] = comm_total;
          };
        };
        // Push new data, remove last column of data
        final_list.push(updated_item.slice(0, -1));
      // Add data with latest amendment
      } else {
        var latest_amen = -1;
        for (j = 0; j < wip_list.length; j++) {
          if (dupes_data[i][0] == wip_list[j][3]) {
            var curr_amen = wip_list[j][22];
            if (curr_amen > latest_amen) {
              latest_amen = curr_amen;
              updated_item = wip_list[j];
            };
          };
        };
        // Push new data, remove last column of data
        final_list.push(updated_item.slice(0, -1));
      };
    };
  };
  ////////////////////////////////////////////////////////////////////////////////
  // Transfer the data
  console.log('Transfer Data to "Current WIP"');
  SpreadsheetApp.flush();
  clearData(_ga_cw);
  SpreadsheetApp.flush();
  // Format PRISM column prior to adding data
  _ga_cw.getRange('H:H').setNumberFormat('@');
  insertData(final_list, _ga_cw);
  // Format data
  formatDataSheet(_ga_cw);
  _ga_cw.getRange('A1').getFilter().sort(4, true);
  _ga_cw.getRange('A1').getFilter().sort(3, true);
  _ga_cw.getRange('A1').getFilter().sort(1, false);
  endTime(st);
};


function formatDataSheet(sheet) {
  changeFont(sheet);
  sheet.getRange('I:K').setNumberFormat('_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)');
  sheet.getRange('G:G').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.autoResizeColumns(1, 6);
};


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////// Step 4 ////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function compareData() {
  var st = Date.now();
  SpreadsheetApp.flush();
  consoleHeader('Step 4: Compare data from WIPs to "Current WIP"');
  var data_cw = getDataValues(_ga_cw, 0, 0);
  var data_rnw = getDataValues(_ga_rnw, 0, 0);
  // Pull data from WIPs
  var data_wips = [];
  var data_wips_rnw = [];
  for (var i = 0; i < _wips.length; i++) {
    console.log('Working: ' + _wips[i][0]);
    var wip = SpreadsheetApp.openById(_wips[i][1]);
    // WIP data
    var data_wip = getDataValues(wip.getSheetByName("WIP"), 0, 0);
    for (var j = 0; j < data_wip.length; j++) data_wips.push(data_wip[j]);
    // Requisitions Need Workspaces
    var data_wip_rnw = getDataValues(wip.getSheetByName("Requisitions Need Workspaces"), 0, 0);
    for (var j = 0; j < data_wip_rnw.length; j++) data_wips_rnw.push(data_wip_rnw[j]);
  };
  ////////////////////////////////////////////////////////////////////////////////
  // Columns to pull data from
  var col_list = [9, 10, 11, 12, 13, 16, 17, 18, 19, 20, 21];
  for (i = 0; i < data_cw.length; i++) {
    for (j = 0; j < data_wips.length; j++) {
      // Compare workspace numbers; there should not be any duplicates, in theory...
      if (data_cw[i][3] == data_wips[j][3]) {
        for (var k = 0; k < col_list.length; k++) {
          data_cw[i][col_list[k]] = data_wips[j][col_list[k]];
        };
        // Remove row from list but only if it matches
        data_wips = data_wips.filter((value) => value != data_wips[j]);
      };
    };
  };
  ////////////////////////////////////////////////////////////////////////////////
  // Requisitions Need Workspaces
  for (var i = 0; i < data_rnw.length; i++) {
    for (var j = 0; j < data_wips_rnw.length; j++) {
      // Compare requisitions
      if (data_rnw[i][4] == data_wips_rnw[j][4]) {
        for (var k = 0; k < col_list.length; k++) {
          data_rnw[i][col_list[k]] = data_wips_rnw[j][col_list[k]];
        };
        // Remove row from list but only if it matches
        data_wips_rnw = data_wips_rnw.filter((value) => value != data_wips_rnw[j]);
      };
    };
  };
  console.log('Add data back');
  insertData(data_cw, _ga_cw);
  formatDataSheet(_ga_cw);
  insertData(data_rnw, _ga_rnw);
  formatDataSheet(_ga_rnw);
  endTime(st);
};


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////// Step 5 ////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function updateAllBranchWIPs(curr_fy_ind) {
  var st = Date.now();
  SpreadsheetApp.flush();
  consoleHeader('Step 5: Update Branch WIPs, pull data and format, create pivots')
  // Save data from all sheets
  var data_cw = getDataValues(_ga_cw, 0, 0);
  var data_rew = getDataValues(_ga_rew, 0, 0);
  var data_rf = getDataValues(_ga_rf, 0, 0);
  var data_rnw = getDataValues(_ga_rnw, -1, 0);
  // Mailbox items
  var data_mi = [];
  for (var i = 0; i < data_cw.length; i++) {
    if (data_cw[i][0].substring(0, 4) == "EAD-") {
      data_mi.push(data_cw[i]);
    };
  };
  // Loop through branch WIPs
  for (var i = 0; i < _wips.length; i++) {
    // Skip specific WIPs
    //if (i < 4) continue;
    var wname = _wips[i][0];
    console.log('Working: ' + wname);
    // Archive file first
    console.log(wname + ": Archiving file");
    archiveFile(_wips[i][1], _wips[i][2]);
    ////////////////////////////////////////////////////////////////////////////////
    console.log(wname + ': Add data for "WIP"');
    var data = [];
    for (var j = 0; j < data_cw.length; j++) {
      if (_wips[i][0] == data_cw[j][0]) {
        data.push(data_cw[j]);
      };
    };
    var sheet = SpreadsheetApp.openById(_wips[i][1]);
    var sheet_cw = sheet.getSheetByName('WIP');
    ////////////////////////////////////////////////////////////////////////////////
    // Reset filters
    console.log("Reset filters");
    var last_col = _ga_cw.getLastColumn();
    var rng_fltr = sheet_cw.getRange('A1:' + columnToLetter(last_col) + String(sheet_cw.getLastRow()));
    var fltr = rng_fltr.getFilter();
    // First check to see if there is a filter; if not, then add
    if (fltr) {
      for (var j = 1; j <= last_col; j++) {
        fltr.setColumnFilterCriteria(j, SpreadsheetApp.newFilterCriteria());
      };
    } else {
      rng_fltr.createFilter();
    };
    ////////////////////////////////////////////////////////////////////////////////
    deleteData(sheet_cw);
    // Format PRISM AAP column prior to adding data
    sheet_cw.getRange('H:H').setNumberFormat('@');
    insertData(data, sheet_cw);
    formatData(sheet_cw);
    ////////////////////////////////////////////////////////////////////////////////
    console.log(wname + ': Add data for "Raw WIP"');
    var data = [];
    for (var j = 0; j < data_rew.length; j++) {
      if (String(_wips[i][0]) == String(data_rew[j][data_rew[0].length - 1])) {
        data.push(data_rew[j].slice(0, -1));
      };
    };
    var sheet_rew = sheet.getSheetByName('Raw WIP');
    clearData(sheet_rew);
    insertData(data, sheet_rew);
    sheet_rew.getRange("A1:" + columnToLetter(sheet_rew.getLastColumn()) + sheet_rew.getLastRow()).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    ////////////////////////////////////////////////////////////////////////////////
    console.log(wname + ': Add data for "Forecasting"');
    var data = [];
    for (var j = 0; j < data_rf.length; j++) {
      if (_wips[i][0] == data_rf[j][data_rf[0].length - 1]) {
        data.push(data_rf[j].slice(0, -1));
      };
    };
    var sheet_rf = sheet.getSheetByName('Forecasting');
    deleteData(sheet_rf);
    insertData(data, sheet_rf);
    sheet_rf.getRange("A1:" + columnToLetter(sheet_rf.getLastColumn()) + sheet_rf.getLastRow()).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    ////////////////////////////////////////////////////////////////////////////////
    console.log(wname + ': Add data for "Requisitions Need Workspaces"');
    var data = [];
    for (var j = 0; j < data_rnw.length; j++) {
      if (_wips[i][0] == data_rnw[j][0]) {
        data.push(data_rnw[j].slice(0, -1));
      };
    };
    var sheet_nwc = sheet.getSheetByName('Requisitions Need Workspaces');
    deleteData(sheet_nwc);
    insertData(data, sheet_nwc);
    formatDataRNW(sheet_nwc);
    ////////////////////////////////////////////////////////////////////////////////
    console.log(wname + ': Add data for "Pivot Data"');
    var sheet_pd = sheet.getSheetByName("Pivot Data");
    clearData(sheet_pd);
    var wip_lrow = sheet_cw.getLastRow();
    var last_col = columnToLetter(sheet_cw.getLastColumn());
    // importRange formula is different for referencing sheets within same workbook
    sheet_pd.getRange("A2").setFormula("='" + sheet_cw.getName() + "'!A2:'" + sheet_cw.getName() + "'!" + last_col + wip_lrow);
    var nwc_lrow = sheet_nwc.getLastRow();
    sheet_pd.getRange("A" + (wip_lrow + 1)).setFormula("='" + sheet_nwc.getName() + "'!A2:'" + sheet_nwc.getName() + "'!" + last_col + nwc_lrow);
    formatDataSheet(sheet_pd);
    ////////////////////////////////////////////////////////////////////////////////  
    console.log(wname + ': Add data for "Mailbox Items"');
    var sheet_mi = sheet.getSheetByName("Mailbox Items");
    clearData(sheet_mi);
    insertData(data_mi, sheet_mi);
    formatDataSheet(sheet_mi);
    ////////////////////////////////////////////////////////////////////////////////
    // Create pivot tables
    console.log(wname + ': Create pivot tables');
    createPivotSheet(sheet, sheet_pd, wname, curr_fy_ind);
  };
  endTime(st);
};


// Only used on sheet "WIP"
function formatData(sheet) {
  var lrow = sheet.getLastRow();
  // Table gridlines
  sheet.getRange('A1:' + columnToLetter(sheet.getLastColumn()) + lrow).activate();
  sheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  // Background
  sheet.getRange('J1:M' + lrow).setBackground('#fff2cc');
  sheet.getRange('Q1:V' + lrow).setBackground('#fff2cc');
  // Drop-Down Menus
  if (lrow > 1) {
    // Clients
    var last_drow = getLastDataRow(_ga_dv, 4);
    var client_list = _ga_dv.getRange('D2:D' + last_drow).getValues();
    var client_list_final = [];
    for (var i = 0; i < client_list.length; i++) {
      client_list_final.push(client_list[i][0]);
    };
    // Type
    last_drow = getLastDataRow(_ga_dv, 6);
    var type_list = _ga_dv.getRange('F2:F' + last_drow).getValues();
    var type_list_final = [];
    for (i = 0; i < type_list.length; i++) {
      type_list_final.push(type_list[i][0]);
    };
    // Special Interest
    last_drow = getLastDataRow(_ga_dv, 9);
    var si_list = _ga_dv.getRange('I2:I' + last_drow).getValues();
    var si_list_final = [];
    for (i = 0; i < si_list.length; i++) {
      si_list_final.push(si_list[i][0]);
    };
    // Fiscal Year
    last_drow = getLastDataRow(_ga_dv, 11);
    var fy_list = _ga_dv.getRange('K2:K' + last_drow).getValues();
    var fy_list_final = [];
    for (i = 0; i < fy_list.length; i++) {
      fy_list_final.push(fy_list[i][0]);
    };
    // Process Update
    last_drow = getLastDataRow(_ga_dv, 13);
    var pu_list = _ga_dv.getRange('M2:M' + last_drow).getValues();
    var pu_list_final = [];
    for (i = 0; i < pu_list.length; i++) {
      pu_list_final.push(pu_list[i][0]);
    };
    // Populate ranges
    sheet.getRange('L2:L' + lrow).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(client_list_final, true).build());
    sheet.getRange('M2:M' + lrow).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(type_list_final, true).build());
    sheet.getRange('Q2:Q' + lrow).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(si_list_final, true).build());
    sheet.getRange('S2:S' + lrow).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(fy_list_final, true).build());
    sheet.getRange('U2:U' + lrow).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(pu_list_final, true).build());
  };
  // Formatting
  formatDataSheet(sheet);
  sheet.getRange('A1').getFilter().sort(4, true);
  sheet.getRange('A1').getFilter().sort(3, true);
};


// Only used sheet "Requisitions Need Workspaces"
function formatDataRNW(sheet) {
  var lrow = sheet.getLastRow();
  // Table gridlines
  sheet.getRange('A1:' + columnToLetter(sheet.getLastColumn()) + lrow).activate();
  sheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  // Yellow background
  sheet.getRange('J1:M' + lrow).setBackground('#fff2cc');
  sheet.getRange('S1:S' + lrow).setBackground('#fff2cc');
  sheet.getRange('U1:V' + lrow).setBackground('#fff2cc');
  // Red background
  sheet.getRange('D1:D' + lrow).setBackground('red');
  sheet.getRange('F1:F' + lrow).setBackground('red');
  sheet.getRange('H1:H' + lrow).setBackground('red');
  sheet.getRange('N1:R' + lrow).setBackground('red');
  sheet.getRange('T1:T' + lrow).setBackground('red');
  sheet.hideColumns(4);
  sheet.hideColumns(6);
  sheet.hideColumns(8);
  sheet.hideColumns(14, 5);
  sheet.hideColumns(20);
  // Drop-Down Menus
  if (lrow > 1) {
    // Client
    var last_drow = getLastDataRow(_ga_dv, 4);
    var client_list = _ga_dv.getRange('D2:D' + last_drow).getValues();
    var client_list_final = [];
    for (var i = 0; i < client_list.length; i++) {
      client_list_final.push(client_list[i][0]);
    };
    // Type
    last_drow = getLastDataRow(_ga_dv, 6);
    var type_list = _ga_dv.getRange('F2:F' + last_drow).getValues();
    var type_list_final = [];
    for (i = 0; i < type_list.length; i++) {
      type_list_final.push(type_list[i][0]);
    };
    // Fiscal Year
    last_drow = getLastDataRow(_ga_dv, 11);
    var fy_list = _ga_dv.getRange('K2:K' + last_drow).getValues();
    var fy_list_final = [];
    for (i = 0; i < fy_list.length; i++) {
      fy_list_final.push(fy_list[i][0]);
    };
    // Process Update
    last_drow = getLastDataRow(_ga_dv, 13);
    var pu_list = _ga_dv.getRange('M2:M' + last_drow).getValues();
    var pu_list_final = [];
    for (i = 0; i < pu_list.length; i++) {
      pu_list_final.push(pu_list[i][0]);
    };
    // Populate ranges
    sheet.getRange('L2:L' + lrow).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(client_list_final, true).build());
    sheet.getRange('M2:M' + lrow).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(type_list_final, true).build());
    sheet.getRange('S2:S' + lrow).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(fy_list_final, true).build());
    sheet.getRange('U2:U' + lrow).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(pu_list_final, true).build());
  };
  // Formatting
  formatDataSheet(sheet);
  sheet.getRange('A1').getFilter().sort(5, true);
  sheet.getRange('A1').getFilter().sort(2, true);
  sheet.getRange('A1').getFilter().sort(1, false);
};


// Ran at the start of each sheet, updating each Branch
function createPivotSheet(sheet, data_sheet, branch_name, curr_fy_ind) {
  // Delete all data
  var name = "Summary";
  sheet.getSheetByName(name).activate();
  SpreadsheetApp.flush();
  sheet.deleteActiveSheet();
  SpreadsheetApp.flush();
  sheet.insertSheet(1).setName(name);
  sum_sheet = sheet.getSheetByName('Summary');
  // Create pivots
  console.log('Start Creating Pivots');
  var drange = "'" + data_sheet.getName() + "'!A1:" + columnToLetter(data_sheet.getLastColumn()) + data_sheet.getLastRow();
  var mi_sheet = sheet.getSheetByName("Mailbox Items");
  var drange_mi = "'Mailbox Items'!A1:" + columnToLetter(mi_sheet.getLastColumn()) + mi_sheet.getLastRow();
  // Fiscal Year filter
  var type_values = getUniqueValues(data_sheet, "M", "", "");
  var proc_values = getUniqueValues(data_sheet, "U", "", "");
  var fy_values = getUniqueValues(data_sheet, "S", "", "");
  createPTMI(branch_name, "Mailbox Items", sum_sheet, drange_mi, getDataValues(mi_sheet, 0, 0), 1, 1, 0, 0);
  createPT01(branch_name, 'Branch Summary by Specialist', sum_sheet, drange, 1, 7, 3, "Contract Specialist", fy_values, curr_fy_ind);
  createPT01(branch_name, 'Branch Summary by Client', sum_sheet, drange, 1, 12, 12, "Client", fy_values, curr_fy_ind);
  createPT01(branch_name, 'Branch Summary by Type', sum_sheet, drange, 1, 17, 13, "Type", fy_values, curr_fy_ind);
  createPT02(branch_name, 'Specialist by Client', sum_sheet, drange, 1, 22, 12, "Client", fy_values, curr_fy_ind);
  createPT02(branch_name, 'Type by Specialist', sum_sheet, drange, 1, 27, 13, "Type", fy_values, curr_fy_ind);
  createPT03(branch_name, 'Update by Type by Specialist', sum_sheet, drange, 1, 32, type_values, proc_values, fy_values, curr_fy_ind);
  createPT04(branch_name, 'Branch Summary by Type and Specialist', sum_sheet, sheet.getSheetByName('WIP'), drange, 1, 32, fy_values, curr_fy_ind);
  createPTF(branch_name, sheet);
};


// Mailbox items
function createPTMI(branch, title, sheet, data_range, list_data, row_position, column_position, fy_values, curr_fy_ind) {
  console.log(branch + ' - ' + title);
  // Set title
  sheet.getRange(row_position, column_position).setValue(title);
  // Create Pivot
  var sourceData = sheet.getRange(data_range);
  var pivotTable = sheet.getRange(row_position + 1, column_position).createPivotTable(sourceData);
  // Rows
  var pivotGroup = pivotTable.addRowGroup(1);
  pivotGroup.setDisplayName("Branch");
  // Values
  pivotValue = pivotTable.addPivotValue(1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue.setDisplayName('Actions');
  pivotValue = pivotTable.addPivotValue(9, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName('Commit Amount');
  pivotValue = pivotTable.addPivotValue(10, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName('Total Obligations');
  pivotValue = pivotTable.addPivotValue(11, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName('Total Value of Actions');
  // Filters
  // Branch filters
  if (branch.indexOf("NOS") >= 0) {
    var branch_filter = ["EAD-NOS MAILBOX"];
  } else if (branch.indexOf("NWS") >= 0) {
    var branch_filter = ["EAD-NWS MAILBOX"];
  } else if (branch.indexOf("OMAO") >= 0) {
    var branch_filter = ["EAD-OMAO MAILBOX"];
  } else if (branch.indexOf("SAP") >= 0) {
    var branch_filter = ["EAD-SAP MAILBOX"];
  };
  // Don't show anything for the branch filter if there are no records
  var mi_list = [];
  for (var i = 0; i < list_data.length; i++) {
    mi_list.push(list_data[i][0]);
  };
  if (mi_list.filter(x => x == branch_filter).length == 0) {
    branch_filter = [];
  };
  var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(branch_filter).build();
  pivotTable.addFilter(1, criteria);
  /*
  // Only filter for current fiscal year if a value is inputted
  if (curr_fy_ind == undefined) {
    var fy_filter = fy_values;
  } else {
    var fy_filter = ['', String(Utilities.formatDate(new Date(), "EST", "yyyy"))];
  };
  var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(fy_filter).build();
  pivotTable.addFilter(19, criteria);
  */
  // Formatting
  pivotTableCleanup(sheet, row_position, column_position, 5, [3, 4, 5], );
};


// Branch Summary by Specialist, Branch Summary by Client, Branch Summary by Type
function createPT01(branch, title, sheet, data_range, row_position, column_position, row_group, row_name, filter1, curr_fy_ind) {
  console.log(branch + ' - ' + title);
  // Set title
  sheet.getRange(row_position, column_position).setValue(title);
  // Create Pivot
  var sourceData = sheet.getRange(data_range);
  var pivotTable = sheet.getRange(row_position + 1, column_position).createPivotTable(sourceData);
  // Rows
  var pivotGroup = pivotTable.addRowGroup(row_group);
  pivotGroup.setDisplayName(row_name);
  // Values
  pivotValue = pivotTable.addPivotValue(row_group, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue.setDisplayName('Actions');
  pivotValue = pivotTable.addPivotValue(9, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName('Funded Amount');
  pivotValue = pivotTable.addPivotValue(11, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName('Total Acquisition Value');
  // Filters
  // Only filter for current fiscal year if a value is inputted
  if (curr_fy_ind == undefined) {
    var fy_filter = filter1;
  } else {
    var fy_filter = ['', String(Utilities.formatDate(new Date(), "EST", "yyyy"))];
  };
  var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(fy_filter).build();
  pivotTable.addFilter(19, criteria);
  // Formatting
  pivotTableCleanup(sheet, row_position, column_position, 4, [3, 4], );
};


// Specialist by Client, Type by Specialist
function createPT02(branch, title, sheet, data_range, row_position, column_position, row_group, row_name, filter1, curr_fy_ind) {
  console.log(branch + ' - ' + title);
  // Set title
  sheet.getRange(row_position, column_position).setValue(title);
  // Create Pivot
  var sourceData = sheet.getRange(data_range);
  var pivotTable = sheet.getRange(row_position + 1, column_position).createPivotTable(sourceData);
  // Rows
  var pivotGroup = pivotTable.addRowGroup(3);
  pivotGroup.setDisplayName('Contract Specialist');
  pivotGroup = pivotTable.addRowGroup(row_group);
  pivotGroup.setDisplayName(row_name);
  pivotGroup.showTotals(false);
  // Values
  pivotValue = pivotTable.addPivotValue(3, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue.setDisplayName('Actions');
  pivotValue = pivotTable.addPivotValue(9, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName('Funded Amount');
  // Filters
  // Only filter for current fiscal year if a value is inputted
  if (curr_fy_ind == undefined) {
    var fy_filter = filter1;
  } else {
    var fy_filter = ['', String(Utilities.formatDate(new Date(), "EST", "yyyy"))];
  };
  var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(fy_filter).build();
  pivotTable.addFilter(19, criteria);
  // Formatting
  pivotTableCleanup(sheet, row_position, column_position, 4, [4], );
};


// Update by Type by Specialist
function createPT03(branch, title, sheet, data_range, row_position, column_position, filter1, filter2, filter3, curr_fy_ind) {
  console.log(branch + ' - ' + title);
  // Set title
  sheet.getRange(row_position, column_position).setValue(title);
  // Create Pivot
  var sourceData = sheet.getRange(data_range);
  var pivotTable = sheet.getRange(row_position + 1, column_position).createPivotTable(sourceData);
  // Rows
  var pivotGroup = pivotTable.addRowGroup(3);
  pivotGroup.setDisplayName('Contract Specialist');
  pivotGroup = pivotTable.addRowGroup(13);
  pivotGroup.setDisplayName('Type');
  pivotGroup.showTotals(false);
  pivotGroup = pivotTable.addRowGroup(21);
  pivotGroup.setDisplayName('Process Update');
  pivotGroup.showTotals(false);
  // Values
  pivotValue = pivotTable.addPivotValue(3, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue.setDisplayName('Actions');
  pivotValue = pivotTable.addPivotValue(9, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName('Funded Amount');
  // Filters
  // Type
  criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(filter1).build();
  pivotTable.addFilter(13, criteria);
  // Process Update
  criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(filter2).build();
  pivotTable.addFilter(21, criteria);
  // Only filter for current fiscal year if a value is inputted
  if (curr_fy_ind == undefined) {
    var fy_filter = filter3;
  } else {
    var fy_filter = ['', String(Utilities.formatDate(new Date(), "EST", "yyyy"))];
  };
  var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(fy_filter).build();
  pivotTable.addFilter(19, criteria);
  // Formatting
  pivotTableCleanup(sheet, row_position, column_position, 5, [5], );
};


// Branch Summary by Type and Specialist
function createPT04(branch, title, sheet, data_sheet, data_range, row_position, column_position, filter1, curr_fy_ind) {
  console.log(branch + ' - ' + title);
  // Set title
  sheet.getRange(row_position, column_position).setValue(title);
  // Create Pivot
  var sourceData = sheet.getRange(data_range);
  var pivotTable = sheet.getRange(row_position + 1, column_position).createPivotTable(sourceData);
  // Rows
  var pivotGroup = pivotTable.addRowGroup(13);
  pivotGroup.setDisplayName('Type');
  // Columns
  pivotGroup = pivotTable.addColumnGroup(3);
  pivotGroup.setDisplayName('Contract Specialist');
  pivotGroup.showTotals(false);
  // Values
  pivotValue = pivotTable.addPivotValue(3, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue.setDisplayName('Actions');
  // Filters
  // Only filter for current fiscal year if a value is inputted
  if (curr_fy_ind == undefined) {
    var fy_filter = filter1;
  } else {
    var fy_filter = ['', String(Utilities.formatDate(new Date(), "EST", "yyyy"))];
  };
  var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(fy_filter).build();
  pivotTable.addFilter(19, criteria);
  // Formatting
  pivotTableCleanup(sheet, row_position, column_position, getUniqueValues(data_sheet, 'C', "", "").length, );
};


// Forecasting
function createPTF(branch, sheet) {
  console.log(branch + ' - Forecasting Records Table');
  var summ_sheet = sheet.getSheetByName("Summary");
  var last_row = getLastDataRow(sheet, 1) + 10;
  summ_sheet.getRange(last_row, 1).setValue("Number of Forecasting Records");
  summ_sheet.getRange(last_row, 2).setValue(sheet.getSheetByName("Forecasting").getLastRow() - 1);
  // Formatting
  summ_sheet.getRange("A" + last_row + ":B" + last_row).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
};


function pivotTableCleanup(sheet, first_row, first_column, number_columns, acct_columns, resize_ind) {
  var fcol_letter = columnToLetter(first_column);
  var lcol = first_column + number_columns - 1;
  var lcol_letter = columnToLetter(lcol);
  // Title
  sheet.getRange(first_row, first_column).setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange(fcol_letter + (first_row + 1) + ':' + lcol_letter + (first_row + 1)).setHorizontalAlignment('center').setFontWeight('bold');
  // Formatting
  var ldrow = getLastDataRow(sheet, first_column);
  var pivot_range = sheet.getRange(fcol_letter + first_row + ':' + lcol_letter + ldrow);
  pivot_range.setNumberFormat('General');
  if (acct_columns != undefined) {
    for (var i = 0; i < acct_columns.length; i++) {
      var acct_column = columnToLetter(first_column + acct_columns[i] - 1);
      sheet.getRange((acct_column + first_row) + ':' + (acct_column + ldrow)).setNumberFormat('_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)');
    };
  };
  pivot_range.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  changeFont(sheet);
  sheet.autoResizeColumns(first_column, number_columns);
  // Resize left column
  if (resize_ind == undefined) {
    if (first_column > 1) {
      sheet.setColumnWidth(first_column - 1, 15);
    };
  };
  // Merging had to be last to prevent errors, still don't understand why
  sheet.getRange((fcol_letter + first_row + ':' + lcol_letter + first_row).toString()).mergeAcross();
};


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////// Step 6 ////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function cwUpdateWIP() {
  var st = Date.now();
  SpreadsheetApp.flush();
  consoleHeader('Step 6: Updating "Consolidated WIP"');
  console.log("Archive file");
  archiveFile(_file_cw, _folder_cw);
  ////////////////////////////////////////////////////////////////////////////////
  // Save data to sheet
  SpreadsheetApp.flush();
  clearData(_gac_w);
  SpreadsheetApp.flush();
  // Format PRISM column prior to adding data
  _gac_w.getRange('H:H').setNumberFormat('@');
  console.log('Import ranges from "Current WIP" and "Requisitions Need Workspaces"')
  var ga_cw_values = getDataValues(_ga_cw, 0, 0);
  var ga_nwc_values = getDataValues(_ga_rnw, 0, 0);
  // Combine the two lists
  ga_nwc_values.forEach(i => ga_cw_values.push(i));
  all_values = [];
  for (var i = 0; i < ga_cw_values.length; i++) {
    if (ga_cw_values[i][0] != "") all_values.push(ga_cw_values[i]);
  };
  ////////////////////////////////////////////////////////////////////////////////
  console.log('Add "Branch Chief" column');
  var dv_list = vLookupList(19, 20);
  for (var i = 0; i < all_values.length; i++) {
    all_values[i].push(vLookup(all_values[i][0], dv_list, 2, 0));
  };
  ////////////////////////////////////////////////////////////////////////////////
  insertData(all_values, _gac_w);
  formatDataSheet(_gac_w);
  endTime(st);
};


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////// Step 7 ////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function cwUpdatePivots(curr_fy_ind) {
  var st = Date.now();
  SpreadsheetApp.flush();
  consoleHeader('Step 7: Creating Pivots for "Conslidated WIP"');
  // First import historical data
  console.log('Find historical week data');
  var week_num = Utilities.formatDate(new Date(), "CST", "w");
  var rng_week = _gac_hd.getRange("A1:A" + _gac_hd.getLastRow()).getValues();
  for (var i = 0; i < rng_week.length; i++) {
    if (rng_week[i] == week_num) {
      var row_id = i;
      _gac_hd.getRange("B" + (i + 1) + ":" + columnToLetter(_gac_hd.getLastColumn()) + (i + 40)).copyTo(_gac_d.getRange("A1"), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      break;
    };
  };
  // Update Summary Tables first
  cwUpdateSummaryTables();
  // EAD Summary by Specialist
  cwCreatePivots01(curr_fy_ind);
  // NOS and OMAO's Branch Summary
  cwCreatePivots02(curr_fy_ind);
  // NWS's Branch Summary
  cwCreatePivots03(curr_fy_ind);
  // Summary by Branch and Type
  cwCreatePivots04(curr_fy_ind);
  // Summary by Specialist and Type
  cwCreatePivots05(curr_fy_ind);
  // Save values back to historical data
  console.log('Save data back to historical data');
  _gac_d.getRange("A1:" + columnToLetter(_gac_d.getLastColumn()) + "40").copyTo(_gac_hd.getRange("B" + (row_id + 1)), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  endTime(st);
};


function cwUpdateSummaryTables() {
  consoleHeader('Updating summary tables');
  console.log("NOS and OMAO Branch Summary");
  // LARGE
  _gac_d.getRange(12, 2).setValue(_gac_d.getRange(48, 3).getValue());
  _gac_d.getRange(12, 3).setValue(_gac_d.getRange(48, 4).getValue());
  // SAP
  _gac_d.getRange(13, 2).setValue(_gac_d.getRange(49, 3).getValue());
  _gac_d.getRange(13, 3).setValue(_gac_d.getRange(49, 4).getValue());
  // Mailboxes
  _gac_d.getRange(14, 2).setValue(_gac_d.getRange(53, 3).getValue());
  _gac_d.getRange(14, 3).setValue(_gac_d.getRange(53, 4).getValue());
  // Totals
  _gac_d.getRange(15, 2).setValue(_gac_d.getRange(12, 2).getValue() + _gac_d.getRange(13, 2).getValue() + _gac_d.getRange(14, 2).getValue());
  _gac_d.getRange(15, 3).setValue(_gac_d.getRange(12, 3).getValue() + _gac_d.getRange(13, 3).getValue() + _gac_d.getRange(14, 3).getValue());
  console.log("NWS Branch Summary");
  // LARGE
  _gac_d.getRange(29, 2).setValue(_gac_d.getRange(57, 3).getValue());
  _gac_d.getRange(29, 3).setValue(_gac_d.getRange(57, 4).getValue());
  // SAP
  _gac_d.getRange(30, 2).setValue(_gac_d.getRange(58, 3).getValue());
  _gac_d.getRange(30, 3).setValue(_gac_d.getRange(58, 4).getValue());
  // Mailboxes
  _gac_d.getRange(31, 2).setValue(_gac_d.getRange(61, 3).getValue());
  _gac_d.getRange(31, 3).setValue(_gac_d.getRange(61, 4).getValue());
  // Totals
  _gac_d.getRange(32, 2).setValue(_gac_d.getRange(29, 2).getValue() + _gac_d.getRange(30, 2).getValue() + _gac_d.getRange(31, 2).getValue());
  _gac_d.getRange(32, 3).setValue(_gac_d.getRange(29, 3).getValue() + _gac_d.getRange(30, 3).getValue() + _gac_d.getRange(31, 3).getValue());
  console.log("EAD Summary");
  for (var i = 4; i <= 7; i++) {
    for (var j = 2; j <= 3; j++) {
      // Mailboxes
      if (i == 6) {
        _gac_d.getRange(i, j).setValue(_gac_d.getRange(78, j + 1).getValue());
      // Totals
      } else if (i == 7) {
        _gac_d.getRange(i, j).setValue(_gac_d.getRange(i - 3, j).getValue() + _gac_d.getRange(i - 2, j).getValue() + _gac_d.getRange(i - 1, j).getValue());
      // Large and SAP
      } else {
        _gac_d.getRange(i, j).setValue(_gac_d.getRange(i + 8, j).getValue() + _gac_d.getRange(i + 25, j).getValue());
      };
    };
  };
  consoleHeader("Updating previous year totals");
  console.log("NOS and OMAO Branch Summary");
  // NOS
  _gac_d.getRange(20, 2).setValue(_gac_d.getRange(67, 3).getValue());
  _gac_d.getRange(20, 3).setValue(_gac_d.getRange(67, 4).getValue());
  // OMAO
  _gac_d.getRange(21, 2).setValue(_gac_d.getRange(64, 3).getValue());
  _gac_d.getRange(21, 3).setValue(_gac_d.getRange(64, 4).getValue());
  // SAP
  _gac_d.getRange(22, 2).setValue(_gac_d.getRange(63, 3).getValue());
  _gac_d.getRange(22, 3).setValue(_gac_d.getRange(63, 4).getValue());
  // Mailboxes
  _gac_d.getRange(23, 2).setValue(_gac_d.getRange(53, 3).getValue());
  _gac_d.getRange(23, 3).setValue(_gac_d.getRange(53, 4).getValue());
  // Totals
  _gac_d.getRange(24, 2).setValue(_gac_d.getRange(20, 2).getValue() + _gac_d.getRange(21, 2).getValue() + _gac_d.getRange(22, 2).getValue() + _gac_d.getRange(23, 2).getValue());
  _gac_d.getRange(24, 3).setValue(_gac_d.getRange(20, 3).getValue() + _gac_d.getRange(21, 3).getValue() + _gac_d.getRange(22, 3).getValue() + _gac_d.getRange(23, 3).getValue());
  console.log("NWS Branch Summary");
  // NWS
  _gac_d.getRange(37, 2).setValue(_gac_d.getRange(72, 3).getValue());
  _gac_d.getRange(37, 3).setValue(_gac_d.getRange(72, 4).getValue());
  // SAP
  _gac_d.getRange(38, 2).setValue(_gac_d.getRange(69, 3).getValue());
  _gac_d.getRange(38, 3).setValue(_gac_d.getRange(69, 4).getValue());
  // Mailboxes
  _gac_d.getRange(39, 2).setValue(_gac_d.getRange(61, 3).getValue());
  _gac_d.getRange(39, 3).setValue(_gac_d.getRange(61, 4).getValue());
  // Totals
  _gac_d.getRange(40, 2).setValue(_gac_d.getRange(37, 2).getValue() + _gac_d.getRange(38, 2).getValue() + _gac_d.getRange(39, 2).getValue());
  _gac_d.getRange(40, 3).setValue(_gac_d.getRange(37, 3).getValue() + _gac_d.getRange(38, 3).getValue() + _gac_d.getRange(39, 3).getValue());
};


// EAD Summary by Specialist
function cwCreatePivots01(curr_fy_ind) {
  var name = "EAD Summary by Specialist";
  consoleHeader('Updating "' + name + '" sheet');
  var drange = _gac_w.getRange('A1:' + columnToLetter(_gac_w.getLastColumn()) + _gac_w.getLastRow());
  // Create a new sheet
  _gac_e.activate();
  SpreadsheetApp.flush();
  _gac.deleteActiveSheet();
  // Code will time-out if flush() is not included here
  SpreadsheetApp.flush();
  _gac.insertSheet(0).setName(name);
  _gac_e = _gac.getSheetByName(name);
  // Create pivots
  // Row 1
  var bc_values = getUniqueValues(_gac_w, 'A', "", "");
  var pu_values = getUniqueValues(_gac_w, 'U', "", "");
  var fy_values = getUniqueValues(_gac_w, 'S', "", "");
  var filter_values = ['EAD-NOS MAILBOX', 'EAD-NWS MAILBOX', 'EAD-OMAO MAILBOX', 'EAD-SAP MAILBOX'];
  // Update table titles and Pivot FY totals based on EOFY status
  if (curr_fy_ind != undefined) {
    var fy = Number(Utilities.formatDate(new Date(), 'EST', 'yy'));
    //cwPivot01ESBS("FY" + String(fy) + ' EAD Total WIP', _gac_e, drange, 1, 1, bc_values, pu_values, fy_values, 0, curr_fy_ind, [5, 6, 7]);
    //cwPivot01ESBS("FY" + String(fy + 1) + ' EAD Total WIP', _gac_e, drange, 1, 9, bc_values, pu_values, String(fy + 1), 1, curr_fy_ind, [5, 6, 7]);
    //cwPivot01('Mailboxes', _gac_e, drange, 1, 17, 1, filter_values, pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
    cwPivot01("FY" + String(fy) + ' EAD Total WIP', _gac_e, drange, 1, 1, 23, bc_values, pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
    cwPivot01FY("FY" + String(fy + 1) + ' EAD Total WIP', _gac_e, drange, 1, 7, 23, bc_values, pu_values, String(fy + 1), curr_fy_ind, [3, 4, 5]);
    cwPivot01('Mailboxes', _gac_e, drange, 1, 13, 1, filter_values, pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  } else {
    //cwPivot01ESBS('EAD Total WIP', _gac_e, drange, 1, 1, bc_values, pu_values, fy_values, 0, curr_fy_ind, [5, 6, 7]);
    //cwPivot01('Mailboxes', _gac_e, drange, 1, 9, 1, filter_values, pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
    cwPivot01('EAD Total WIP', _gac_e, drange, 1, 1, 23, bc_values, pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
    cwPivot01('Mailboxes', _gac_e, drange, 1, 7, 1, filter_values, pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  };
  // Row 2
  var frow = _gac_e.getLastRow();
  cwPivot01("Stacy Dohse'S NOS Branch", _gac_e, drange, frow + 3, 1, 3, ['NOS'], pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  cwPivot01("James Price's NOS Branch", _gac_e, drange, frow + 3, 7, 3, ['NOS KC'], pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  // Row 3
  frow = _gac_e.getLastRow();
  cwPivot01("Jennifer Hildebrandt's NWS Branch", _gac_e, drange, frow + 3, 1, 3, ['NWS'], pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  cwPivot01("Jackie Shewmaker's NWS Branch", _gac_e, drange, frow + 3, 7, 3, ['NWS KC'], pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  // Row 4
  frow = _gac_e.getLastRow();
  cwPivot01("Dawn Dabney's SAP Branch", _gac_e, drange, frow + 3, 1, 3, ['SAP'], pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  cwPivot01("Steven Prado's SAP Branch", _gac_e, drange, frow + 3, 7, 3, ['SAP KC'], pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  // Row 5
  frow = _gac_e.getLastRow();
  cwPivot01("OMAO Branch", _gac_e, drange, frow + 3, 1, 3, ['OMAO'], pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  // Row 6
  frow = _gac_e.getLastRow();
  _gac_d.getRange('A1:U7').copyTo(_gac_e.getRange(frow + 3, 1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  _gac_e.getRange('A' + (frow + 3) + ':U' + (frow + 3)).mergeAcross();
  // Row 7
  frow = _gac_e.getLastRow();
  cwPivot01('Mailboxes', _gac_e, drange, frow + 3, 1, 1, filter_values, pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  // Formatting
  changeFont(_gac_e);
  //_gac_e.hideColumns(18, 7);
  // Set title
  _gac_e.insertRows(1, 2);
  _gac_e.getRange(1, 1).setValue('EAD - WIP - ' + Utilities.formatDate(new Date(), 'EST', 'MMMM dd, yyyy'));
  _gac_e.getRange('A1:K1').mergeAcross();
  _gac_e.getRange(1, 1).setHorizontalAlignment('center').setFontWeight('bold').setFontSize(18);
};


// NOS and OMAO Branch Summary
function cwCreatePivots02(curr_fy_ind) {
  var name = "NOS and OMAO Summary";
  consoleHeader('Updating "' + name + '" sheet');
  var drange = _gac_w.getRange('A1:' + columnToLetter(_gac_w.getLastColumn()) + _gac_w.getLastRow());
  // Create a new sheet
  _gac_ec.activate();
  SpreadsheetApp.flush();
  _gac.deleteActiveSheet();
  SpreadsheetApp.flush();
  _gac.insertSheet(1).setName(name);
  _gac_ec = _gac.getSheetByName(name);
  // Create pivots
  // Row 1
  var pu_values = getUniqueValues(_gac_w, 'U', "", "");
  var fy_values = getUniqueValues(_gac_w, 'S', "", "");
  var filter_values = ['NOS KC', 'SAP', 'NOS', 'OMAO'];
  cwPivot01("NOS and OMAO's Total WIP", _gac_ec, drange, 1, 1, 23, filter_values, pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  filter_values = ['EAD-SAP MAILBOX', 'EAD-NOS MAILBOX', 'EAD-OMAO MAILBOX'];
  cwPivot01('Mailboxes', _gac_ec, drange, 1, 7, 1, filter_values, pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  // Row 2
  var frow = _gac_ec.getLastRow();
  cwPivot01("Stacy Dohse's NOS Branch", _gac_ec, drange, frow + 3, 1, 3, ['NOS'], pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  cwPivot01("James Price's NOS Branch", _gac_ec, drange, frow + 3, 7, 3, ['NOS KC'], pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  // Row 3
  frow = _gac_ec.getLastRow();
  cwPivot01("Dawn Dabney's SAP Branch", _gac_ec, drange, frow + 3, 1, 3, ['SAP'], pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  cwPivot01("OMAO Branch", _gac_ec, drange, frow + 3, 7, 3, ['OMAO'], pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  // Row 4
  frow = _gac_ec.getLastRow();
  _gac_d.getRange('A9:U15').copyTo(_gac_ec.getRange(frow + 3, 1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  _gac_ec.getRange('A' + (frow + 3) + ':U' + (frow + 3)).mergeAcross();
  // Row 5
  frow = _gac_ec.getLastRow();
  _gac_d.getRange('A17:G24').copyTo(_gac_ec.getRange(frow + 3, 1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  _gac_ec.getRange('A' + (frow + 3) + ':G' + (frow + 3)).mergeAcross();
  // Extra column in row 1
  filter_values = ['NOS KC', 'SAP', 'NOS', 'OMAO', 'EAD-SAP MAILBOX', 'EAD-NOS MAILBOX', 'EAD-OMAO MAILBOX'];
  cwPivot01("NOS and OMAO's Summary", _gac_ec, drange, 1, 22, 1, filter_values, pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  // Formatting
  changeFont(_gac_ec);
  //_gac_ec.hideColumns(18, 7);
  // Set title
  _gac_ec.insertRows(1, 2);
  _gac_ec.getRange(1, 1).setValue(name + ' - ' + Utilities.formatDate(new Date(), 'EST', 'MMMM dd, yyyy'));
  _gac_ec.getRange('A1:K1').mergeAcross();
  _gac_ec.getRange(1, 1).setHorizontalAlignment('center').setFontWeight('bold').setFontSize(18);
};


// NWS Branch Summary
function cwCreatePivots03(curr_fy_ind) {
  var name = "NWS Summary";
  consoleHeader('Updating "' + name + '" sheet');
  var drange = _gac_w.getRange('A1:' + columnToLetter(_gac_w.getLastColumn()) + _gac_w.getLastRow());
  // Create a new sheet
  _gac_nd.activate();
  SpreadsheetApp.flush();
  _gac.deleteActiveSheet();
  SpreadsheetApp.flush();
  _gac.insertSheet(2).setName(name);
  _gac_nd = _gac.getSheetByName(name);
  // Create pivots
  // Row 1
  var pu_values = getUniqueValues(_gac_w, 'U', "", "");
  var fy_values = getUniqueValues(_gac_w, 'S', "", "");
  var filter_values = ['NWS', 'NWS KC', 'SAP KC'];
  cwPivot01("NWS Total WIP", _gac_nd, drange, 1, 1, 23, filter_values, pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  filter_values = ['EAD-SAP MAILBOX', 'EAD-NWS MAILBOX'];
  cwPivot01('Mailboxes', _gac_nd, drange, 1, 7, 1, filter_values, pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  // Row 2
  var frow = _gac_nd.getLastRow();
  cwPivot01("Jennifer Hildebrandt's NWS Branch", _gac_nd, drange, frow + 3, 1, 3, ['NWS'], pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  cwPivot01("Steven Prado's SAP Branch", _gac_nd, drange, frow + 3, 7, 3, ['SAP KC'], pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  // Row 3
  frow = _gac_nd.getLastRow();
  cwPivot01("Jackie Shewmaker's NWS Branch", _gac_nd, drange, frow + 3, 1, 3, ['NWS KC'], pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  // Row 4
  frow = _gac_nd.getLastRow();
  _gac_d.getRange('A26:U32').copyTo(_gac_nd.getRange(frow + 3, 1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  _gac_nd.getRange('A' + (frow + 3) + ':U' + (frow + 3)).mergeAcross();
  // Row 5
  frow = _gac_nd.getLastRow();
  _gac_d.getRange('A34:G40').copyTo(_gac_nd.getRange(frow + 3, 1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  _gac_nd.getRange('A' + (frow + 3) + ':G' + (frow + 3)).mergeAcross();
  // Extra column in row 1
  filter_values = ['SAP KC', 'NWS KC', 'NWS', 'EAD-SAP MAILBOX', 'EAD-NWS MAILBOX'];
  cwPivot01("NWS Summary", _gac_nd, drange, 1, 22, 1, filter_values, pu_values, fy_values, curr_fy_ind, [3, 4, 5]);
  // Formatting
  changeFont(_gac_nd);
  //_gac_nd.hideColumns(18, 7);
  // Set title
  _gac_nd.insertRows(1, 2);
  _gac_nd.getRange(1, 1).setValue(name + ' - ' + Utilities.formatDate(new Date(), 'EST', 'MMMM dd, yyyy'));
  _gac_nd.getRange('A1:K1').mergeAcross();
  _gac_nd.getRange(1, 1).setHorizontalAlignment('center').setFontWeight('bold').setFontSize(18);
};


// Summary by Branch and Type
function cwCreatePivots04(curr_fy_ind) {
  var name = "Summary by Branch and Type";
  consoleHeader('Updating "' + name + '" sheet');
  var drange = _gac_w.getRange('A1:' + columnToLetter(_gac_w.getLastColumn()) + _gac_w.getLastRow());
  // Create a new sheet
  _gac_sb.activate();
  SpreadsheetApp.flush();
  _gac.deleteActiveSheet();
  SpreadsheetApp.flush();
  _gac.insertSheet(3).setName(name);
  _gac_sb = _gac.getSheetByName(name);
  // Create pivots
  var fy_values = getUniqueValues(_gac_w, 'S', "", "");
  cwPivot02("Dawn Dabney's SAP Branch by Contract Type", _gac_sb, drange, 1, 1, ['SAP'], fy_values, curr_fy_ind, [3]);
  cwPivot02("Jennifer Hildebrandt's NWS Branch by Contract Type", _gac_sb, drange, 1, 5, ['NWS'], fy_values, curr_fy_ind, [3]);
  cwPivot02("Stacy Dohse's NOS Branch by Contract Type", _gac_sb, drange, 1, 9, ['NOS'], fy_values, curr_fy_ind, [3]);
  cwPivot02("OMAO Branch by Contract Type", _gac_sb, drange, 1, 13, ['OMAO'], fy_values, curr_fy_ind, [3]);
  cwPivot02("Steven Prado's SAP Branch by Contract Type", _gac_sb, drange, 1, 17, ['SAP KC'], fy_values, curr_fy_ind, [3]);
  cwPivot02("Jackie Shewmaker's NWS Branch by Contract Type", _gac_sb, drange, 1, 21, ['NWS KC'], fy_values, curr_fy_ind, [3]);
  cwPivot02("James Price's NOS Branch by Contract Type", _gac_sb, drange, 1, 25, ['NOS KC'], fy_values, curr_fy_ind, [3]);
  var filter_values = ['EAD CLOSEOUTS MAILBOX', 'EAD-NOS MAILBOX', 'EAD-NWS MAILBOX', 'EAD-OMAO MAILBOX', 'EAD-SAP MAILBOX'];
  cwPivot02('Mailboxes Unassigned by Contract Type', _gac_sb, drange, 1, 29, filter_values, fy_values, curr_fy_ind, [3]);
  // Formatting
  changeFont(_gac_sb);
  // Set title
  _gac_sb.insertRows(1, 3);
  _gac_sb.getRange(1, 1).setValue('EAD - WIP - ' + Utilities.formatDate(new Date(), 'EST', 'MMMM dd, yyyy'));
  _gac_sb.getRange(2, 1).setValue('Summary by Branch and Contract Type');
  _gac_sb.getRange(('A1:' + columnToLetter(_gac_sb.getLastColumn()) + '1').toString()).mergeAcross();
  _gac_sb.getRange(1, 1).setHorizontalAlignment('center').setFontWeight('bold').setFontSize(18);
  _gac_sb.getRange(('A2:' + columnToLetter(_gac_sb.getLastColumn()) + '2').toString()).mergeAcross();
  _gac_sb.getRange(2, 1).setHorizontalAlignment('center').setFontWeight('bold').setFontSize(18);
};


// Summary by Specialist and Type
function cwCreatePivots05(curr_fy_ind) {
  var name = 'Summary by Specialist and Type';
  consoleHeader('Updating "' + name + '" sheet');
  var drange = _gac_w.getRange('A1:' + columnToLetter(_gac_w.getLastColumn()) + _gac_w.getLastRow());
  // Create a new sheet
  _gac_ss.activate();
  SpreadsheetApp.flush();
  _gac.deleteActiveSheet();
  SpreadsheetApp.flush();
  _gac.insertSheet(4).setName(name);
  _gac_ss = _gac.getSheetByName(name);
  // Create pivots
  var fy_values = getUniqueValues(_gac_w, 'S', "", "");
  cwPivot03("Dawn Dabney's SAP Branch by Contract Type", _gac_ss, drange, 1, 1, ['SAP'], fy_values, curr_fy_ind, [4]);
  cwPivot03("Jennifer Hildebrandt's NWS Branch by Contract Type", _gac_ss, drange, 1, 6, ['NWS'], fy_values, curr_fy_ind, [4]);
  cwPivot03("Stacy Dohse's NOS Branch by Contract Type", _gac_ss, drange, 1, 11, ['NOS'], fy_values, curr_fy_ind, [4]);
  cwPivot03("OMAO Branch by Contract Type", _gac_ss, drange, 1, 16, ['OMAO'], fy_values, curr_fy_ind, [4]);
  cwPivot03("Steven Prado's SAP Branch by Contract Type", _gac_ss, drange, 1, 21, ['SAP KC'], fy_values, curr_fy_ind, [4]);
  cwPivot03("Jackie Shewmaker's NWS Branch by Contract Type", _gac_ss, drange, 1, 26, ['NWS KC'], fy_values, curr_fy_ind, [4]);
  cwPivot03("James Price's NOS Branch by Contract Type", _gac_ss, drange, 1, 31, ['NOS KC'], fy_values, curr_fy_ind, [4]);
  var filter_values = ['EAD CLOSEOUTS MAILBOX', 'EAD-NOS MAILBOX', 'EAD-NWS MAILBOX', 'EAD-OMAO MAILBOX', 'EAD-SAP MAILBOX'];
  cwPivot03('Mailboxes Unassigned by Specialist and Contract Type', _gac_ss, drange, 1, 36, filter_values, fy_values, curr_fy_ind, [4]);
  // Formatting
  changeFont(_gac_ss);
  // Set title
  _gac_ss.insertRows(1, 3);
  _gac_ss.getRange(1, 1).setValue('EAD - WIP - ' + Utilities.formatDate(new Date(), 'EST', 'MMMM dd, yyyy'));
  _gac_ss.getRange(2, 1).setValue('Summary by Specialist and Contract Type');
  _gac_ss.getRange(('A1:' + columnToLetter(_gac_ss.getLastColumn()) + '1').toString()).mergeAcross();
  _gac_ss.getRange(1, 1).setHorizontalAlignment('center').setFontWeight('bold').setFontSize(18);
  _gac_ss.getRange(('A2:' + columnToLetter(_gac_ss.getLastColumn()) + '2').toString()).mergeAcross();
  _gac_ss.getRange(2, 1).setHorizontalAlignment('center').setFontWeight('bold').setFontSize(18);
};


function cwPivot01(title, sheet, data_range, row_start, column_start, pivot_column, filter1, filter2, filter3, curr_fy_ind, formatting_fields) {
  console.log(title);
  // Set title
  sheet.getRange(row_start, column_start).setValue(title);
  // Create Pivot
  var pivotTable = sheet.getRange(row_start + 1, column_start).createPivotTable(data_range);
  // Rows
  pivotTable.addRowGroup(pivot_column);
  // Values
  pivotValue = pivotTable.addPivotValue(pivot_column, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue.setDisplayName('Actions');
  pivotValue = pivotTable.addPivotValue(9, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName('Commit Amount');
  pivotValue = pivotTable.addPivotValue(10, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName('Total Obligations');
  pivotValue = pivotTable.addPivotValue(11, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName('Total Value of Actions');
  // Filters
  // Branch
  var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(filter1).build();
  pivotTable.addFilter(1, criteria);
  // Process Update
  var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(filter2).build();
  pivotTable.addFilter(21, criteria);
  // Fiscla Year - update to current FY
  if (curr_fy_ind == undefined) {  
    var fy_filter = filter3;
  } else {
    var fy_filter = ['', String(Utilities.formatDate(new Date(), "EST", "yyyy"))];
  };
  var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(fy_filter).build();
  pivotTable.addFilter(19, criteria);
  // Formatting
  pivotTableCleanup(sheet, row_start, column_start, 5, formatting_fields, 1);
};


// Only used for next FY actions
function cwPivot01FY(title, sheet, data_range, row_start, column_start, pivot_column, filter1, filter2, filter3, curr_fy_ind, formatting_fields) {
  console.log(title);
  // Set title
  sheet.getRange(row_start, column_start).setValue(title);
  // Create Pivot
  var pivotTable = sheet.getRange(row_start + 1, column_start).createPivotTable(data_range);
  // Rows
  pivotTable.addRowGroup(pivot_column);
  // Values
  pivotValue = pivotTable.addPivotValue(pivot_column, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue.setDisplayName('Actions');
  pivotValue = pivotTable.addPivotValue(9, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName('Commit Amount');
  pivotValue = pivotTable.addPivotValue(10, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName('Total Obligations');
  pivotValue = pivotTable.addPivotValue(11, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName('Total Value of Actions');
  // Filters
  // Branch
  var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(filter1).build();
  pivotTable.addFilter(1, criteria);
  // Process Update
  var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(filter2).build();
  pivotTable.addFilter(21, criteria);
  // Fiscal Year add one to the FY
  var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(["20" + filter3]).build();
  pivotTable.addFilter(19, criteria);
  // Formatting
  pivotTableCleanup(sheet, row_start, column_start, 5, formatting_fields, 1);
};


// EAD Summary By Specialist only
function cwPivot01ESBS(title, sheet, data_range, row_start, column_start, filter1, filter2, filter3, eofy_ind, curr_fy_ind, formatting_fields) {
  console.log(title);
  // Set title
  sheet.getRange(row_start, column_start).setValue(title);
  // Create Pivot
  var pivotTable = sheet.getRange(row_start + 1, column_start).createPivotTable(data_range);
  // Rows
  var pivotGroup = pivotTable.addRowGroup(23);
  pivotGroup.setDisplayName('Branch Chief');
  pivotGroup = pivotTable.addRowGroup(2);
  pivotGroup.setDisplayName('Buyer');
  pivotGroup = pivotTable.addRowGroup(3);
  pivotGroup.setDisplayName('Contracts Office POC');
  pivotGroup.showTotals(false);
  // Values
  pivotValue = pivotTable.addPivotValue(23, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue.setDisplayName('Actions');
  pivotValue = pivotTable.addPivotValue(9, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName('Commit Amount');
  pivotValue = pivotTable.addPivotValue(10, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName('Total Obligations');
  pivotValue = pivotTable.addPivotValue(11, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName('Total Value of Actions');
  // Filters
  // Branch
  var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(filter1).build();
  pivotTable.addFilter(1, criteria);
  // Process Update
  var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(filter2).build();
  pivotTable.addFilter(21, criteria);
  // Fiscla Year - update to current FY
  // eofy 
  if (eofy_ind == 0) {
    if (curr_fy_ind == undefined) {  
      var fy_filter = filter3;
    } else {
      var fy_filter = ['', String(Utilities.formatDate(new Date(), "EST", "yyyy"))];
    };
    var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(fy_filter).build();
  } else {
    var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(["20" + filter3]).build();
  };
  pivotTable.addFilter(19, criteria);
  // Formatting
  pivotTableCleanup(sheet, row_start, column_start, 7, formatting_fields, 1);
};


// Summary by Branch and Type
function cwPivot02(title, sheet, data_range, row_start, column_start, filter1, filter2, curr_fy_ind, formatting_fields) {
  console.log(title);
  // Set title
  sheet.getRange(row_start, column_start).setValue(title);
  // Create Pivot
  var pivotTable = sheet.getRange(row_start + 1, column_start).createPivotTable(data_range);
  // Rows
  pivotTable.addRowGroup(13);
  // Values
  pivotValue = pivotTable.addPivotValue(1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue.setDisplayName('Actions');
  pivotValue = pivotTable.addPivotValue(9, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName('Commit Amount');
  // Filters
  var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(filter1).build();
  pivotTable.addFilter(1, criteria);
  // Only filter for current fiscal year if a value is inputted
  if (curr_fy_ind == undefined) {  
    var fy_filter = filter2;
  } else {
    var fy_filter = ['', String(Utilities.formatDate(new Date(), "EST", "yyyy"))];
  };
  var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(fy_filter).build();
  pivotTable.addFilter(19, criteria);
  // Formatting
  pivotTableCleanup(sheet, row_start, column_start, 3, formatting_fields, );
};


// Summary by Specialist and Type
function cwPivot03(title, sheet, data_range, row_start, column_start, filter1, filter2, curr_fy_ind, formatting_fields) {
  console.log(title);
  // Set title
  sheet.getRange(row_start, column_start).setValue(title);
  // Create Pivot
  var pivotTable = sheet.getRange(row_start + 1, column_start).createPivotTable(data_range);
  // Rows
  var pivotGroup = pivotTable.addRowGroup(3);
  pivotGroup = pivotTable.addRowGroup(13);
  pivotGroup.showTotals(false);
  // Values
  pivotValue = pivotTable.addPivotValue(1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue.setDisplayName('Actions');
  pivotValue = pivotTable.addPivotValue(9, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName('Commit Amount');
  // Filters
  var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(filter1).build();
  pivotTable.addFilter(1, criteria);
  // Only filter for current fiscal year if a value is inputted
  if (curr_fy_ind == undefined) {  
    var fy_filter = filter2;
  } else {
    var fy_filter = ['', String(Utilities.formatDate(new Date(), "EST", "yyyy"))];
  };
  var criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(fy_filter).build();
  pivotTable.addFilter(19, criteria);
  // Formatting
  pivotTableCleanup(sheet, row_start, column_start, 4, formatting_fields, );
};


// Run once every new FY
function cwUpdateFY() {
  getVariables();
  // Identify FY
  var today =  new Date();
  if (Number(today.getMonth()) + 1 >= 10) {
    var fy = "FY" + Utilities.formatDate(today, "EST", "yy") + 1;
  } else {
    var fy = "FY" + Number(Utilities.formatDate(today, "EST", "yy"));
  };
  // Insert data
  _gac_hd.insertColumns(3, 2);
  for (var i = 1; i <= getLastDataRow(_gac_hd, 2); i++) {
    if (_gac_hd.getRange(i, 2).getValue().toLowerCase() == "branch") {
      _gac_hd.getRange(i - 1, 3).setValue(fy);
      _gac_hd.getRange("C" + (i - 1) + ":D" + (i - 1)).mergeAcross();
      _gac_hd.getRange(i, 3).setValue("Actions");
      _gac_hd.getRange(i, 4).setValue("Commit Amount");
    };
  };
};


// Used to update FY25 values since adding FY25 column erroneously
// Pulls previous FY values to blank columns for current FY
function cwUpdateFY25() {
  getVariables();
  var values = _gac_hd.getRange("B1:F376").getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0].toLowerCase() == "large" || values[i][0].toLowerCase() == "sap" || values[i][0].toLowerCase() == "mailboxes" || values[i][0].toLowerCase() == "nos" || values[i][0].toLowerCase() == "omao" || values[i][0].toLowerCase() == "nws" || values[i][0].toLowerCase() == "grand total") {
      values[i][2] = values[i][4];
      values[i][1] = values[i][3];
      values[i][4] = "";
      values[i][3] = "";
    };
  };
  _gac_hd.getRange(1, 2, values.length, values[0].length).setValues(values);
};


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////// Step 8 ////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function draftEmail() {
  var st = Date.now();
  SpreadsheetApp.flush();
  consoleHeader('Step 8: Send emails to branch chiefs');
  //////////////////////////////////////////////////////////////////////
  // Body
  var body = "Good Morning,<br><br>";
  body += "This week's WIP's are ready for your review:<br><br>";
  body += "<a href=\"https://docs.google.com/spreadsheets/d/" + _file_nos + "\">NOS Stacy WIP</a>, ";
  body += "<a href=\"https://drive.google.com/drive/folders/" + _wip_folder_nos + "\">Folder</a><br>";
  body += "<a href=\"https://docs.google.com/spreadsheets/d/" + _file_noskc + "\">NOS James WIP</a>, ";
  body += "<a href=\"https://drive.google.com/drive/folders/" + _wip_folder_noskc + "\">Folder</a><br>";
  body += "<a href=\"https://docs.google.com/spreadsheets/d/" + _file_nws + "\">NWS Jennifer WIP</a>, ";
  body += "<a href=\"https://drive.google.com/drive/folders/" + _wip_folder_nws + "\">Folder</a><br>";
  body += "<a href=\"https://docs.google.com/spreadsheets/d/" + _file_nwskc + "\">NWS Jackie WIP</a>, ";
  body += "<a href=\"https://drive.google.com/drive/folders/" + _wip_folder_nwskc + "\">Folder</a><br>";
  body += "<a href=\"https://docs.google.com/spreadsheets/d/" + _file_omao + "\">OMAO WIP</a>, ";
  body += "<a href=\"https://drive.google.com/drive/folders/" + _wip_folder_omao + "\">Folder</a><br>";
  body += "<a href=\"https://docs.google.com/spreadsheets/d/" + _file_sap + "\">SAP Dawn WIP</a>, ";
  body += "<a href=\"https://drive.google.com/drive/folders/" + _wip_folder_sap + "\">Folder</a><br>";
  body += "<a href=\"https://docs.google.com/spreadsheets/d/" + _file_sapkc + "\">SAP Steven WIP</a>, ";
  body += "<a href=\"https://drive.google.com/drive/folders/" + _wip_folder_sapkc + "\">Folder</a><br>";
  body += "<a href=\"https://docs.google.com/spreadsheets/d/" + _file_cw + "\">Consolidated WIP</a>, ";
  body += "<a href=\"https://drive.google.com/drive/folders/" + _wip_folder_cw + "\">Folder</a><br><br>";
  body += "As always, here are some notes about the WIP:<br><br>";
  body += "<li>A workspace (requirements or modification) must be associated with a requisition in PRISM in order for the requisition dollar value to populate on the WIP.</li>";
  body += "<li>Please note that the requisition includes the funding and a workspace includes the supporting documents.</li>";
  body += "<li>Multiple requisitions and multiple workspaces should not be submitted for the same action. If a requisition requires additional funding or to be edited, then it should be returned to the client to be amended.</li>";
  body += "<li>If none of your actions are populating and the above two bullets do not resolve this, ensure that your Group in PRISM is accurate, i.e. \"EAD-NOS, EAD-NWS, EAD-OMAO, etc.\"</li>";
  body += "<li>A modification (or requirements) workspace should be closed once it is completed in order for it to be removed from the WIP.</li>";
  body += "<li>Please refer to the <a href=\"https://drive.google.com/file/d/1IHcoAemD0anuMn6SVYk350H8-MTcB0I-/view?usp=drive_link\">WIP SOP</a> for instructions on how to utilize the WIP.</li>";
  body += "<li>The PALT is from the <a href=\"https://docs.google.com/spreadsheets/d/1ZGR5b9fCF213jQ2trdov_PdLBgqjzRCk\">DOC Acquisition Planning Timeline Tool</a>.</li>";
  body += "<li>There will always be a number in the PRISM AAP column, which is the last five of the requisition number.</li><br>";
  body += "Please contact me or Carley if you have any questions.<br><br>";
  //////////////////////////////////////////////////////////////////////
  // Signature
  var signature = "<font color='darkgray';>--<br>";
  signature += parseFullName() + "<br>";
  signature += 'NOAA, AGO<br>';
  signature += 'Eastern Acquisition Division';
  //////////////////////////////////////////////////////////////////////
  // Draft Email
  GmailApp.createDraft("",
                       "EAD WIP Updated - " + Utilities.formatDate(new Date(), "EST", "M/dd"),
                       "",
                       {cc: "",
                        htmlBody: body + signature
  });
  endTime(st);
};


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////// Accessory Functions ///////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function archiveFile(file_id, folder_id) {
  var gao = DriveApp.getFileById(file_id);
  var gao_name = gao.getName(); 
  gao.makeCopy();
  var new_name = gao_name + " " + Utilities.formatDate(new Date(), "CST", "yyyyMMdd")
  /*
  var new_name = gao_name + ' ' + currentMonday();
  var folder = DriveApp.getFolderById(folder_id);
  var files = folder.getFiles();
  while (files.hasNext()){
    var file = files.next();
    var file_name = file.getName();
    if (file_name == new_name) {
      file.setTrashed(true);
    };
  };
  */
  DriveApp.getFilesByName('Copy of ' + gao_name).next().setName(new_name).moveTo(DriveApp.getFolderById(folder_id));
};


// Date format will be in yyyyMMdd; update accordingly
function currentMonday() {
  var date = new Date();
  var newDate = date.setDate(date.getDate() - date.getDay() + 1);
  return Utilities.formatDate(new Date(newDate), "CST", "yyyyMMdd");
};


function trimClean(text) {
  return text.toString().toLowerCase().trim().replace(/-/g, '').replace(/\s/g, '');
};


function dateDiff(start_date, end_date, unit) {
  var sdate_y = start_date.substring(0,4);
  var sdate_m = start_date.substring(4,6);
  var sdate_d = start_date.substring(6,8);
  var edate_y = end_date.substring(0,4);
  var edate_m = end_date.substring(4,6);
  var edate_d = end_date.substring(6,8);
  if (sdate_m > 12 || edate_m > 12) {
    return 'One or more of the months specified is greater than 12. Make sure the dates inputted are in the "yyyymmdd" format; for example: January 31, 2024 will be "20240131". Exiting...';
  } else if (sdate_d > 31 || edate_d > 31) {
    return 'One or more of the days specified is greater than 31. Make sure the dates inputted are in the "yyyymmdd" format; for example: January 31, 2024 will be "20240131". Exiting...';
  } else if (start_date.length != 8 || end_date.length != 8) {
    return 'One or more of the dates specified is not inputted in the "yyyymmdd" format. Exiting...';
  };
  // For some reason, the months need to be subtracted by 1 when converted to a new date
  var sdate_new = new Date(sdate_y, sdate_m - 1, sdate_d).getTime();
  var edate_new = new Date(edate_y, edate_m - 1, edate_d).getTime();
  var day_ms = 1000*60*60*24;
  var diff_date = (edate_new-sdate_new)/day_ms;
  if (unit == 'd') {
    return diff_date;
  } else if (unit == 'w') {
    return diff_date/7;
  } else if (unit == 'm') {
    return diff_date/(365/12);
  } else if (unit == 'y') {
    return diff_date/365;
  } else {
    return 'Improper unit specified. Input "d", "w", "m", or "y". Exiting...';
  };
};


function consoleHeader(txt) {
  var txt_len = txt.length;
  var len_diff = 80 - (txt_len + 2);
  console.log('#'.repeat(len_diff/2) + ' ' + txt + ' ' + '#'.repeat(len_diff/2));
};


function findLastPosition(spreadsheet, name_length, search_string) {
  console.log('Find Last Position for Searched Sheet');
  var sh_lst = [];
  var sh_name = ''
  var sheets = spreadsheet.getSheets();
  // Identify all applicable WIPs
  for (var i = 0 ; i < sheets.length ; i++) {
    sh_name = trimClean(sheets[i].getName());
    if (sh_name.length == name_length && sh_name.substring(0, search_string.length) == search_string) {
      sh_lst.push(sheets[i].getName());
    };
  };
  // Find latest WIP
  for (var i = 0 ; i < sh_lst.length ; i++) {
    sh_name = trimClean(sh_lst[i]);
    var wdate = sh_name.slice(-8);
    var ldate = 0;
    if (wdate > ldate) {
      ldate = wdate;
      var sh_n = sh_lst[i];
    };
  };
  return [spreadsheet.getSheetByName(sh_n).getIndex(), sh_n];
};


function vLookupList(start_column, end_column) {
  var lrow = _ga_dv.getRange(1, start_column).getDataRegion(SpreadsheetApp.Dimension.ROWS).getLastRow();
  var rng = columnToLetter(start_column) + "2:" + columnToLetter(end_column) + lrow
  return _ga_dv.getRange(rng).getValues();
};


function vLookup(search_value, list, col_return_num, case_match) {
  // Calculate last row depending on column searchedsheet
  for (var i = 0; i < list.length; i++) {
    if (case_match == 0) {
      if (search_value.toLowerCase() == list[i][0].toLowerCase()) {
        return list[i][col_return_num - 1];
      };
    } else {
      if (search_value == list[i][0]) {
        return list[i][col_return_num - 1];
      };
    };
  };
};


function vLookup2(search_value, sheet, col_search_num, col_return_num, case_match) {
  // Calculate last row depending on column searchedsheet
  var lrow = sheet.getRange(1, col_search_num).getDataRegion(SpreadsheetApp.Dimension.ROWS).getLastRow();
  for (var i = 2; i <= lrow; i++) {
    if (case_match == 0) {
      if (search_value.toLowerCase() == sheet.getRange(i, col_search_num).getValue().toLowerCase()) {
        return sheet.getRange(i, col_return_num).getValue();
      };
    } else {
      if (search_value == sheet.getRange(i, col_search_num).getValue()) {
        return sheet.getRange(i, col_return_num).getValue();
      };
    };
  };
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


function getUniqueValues(sheet, column_letter, data_list, column_number) {
  if (sheet != "" && data_list == "") {
    var rng = sheet.getRange(column_letter + '2:' + column_letter).getValues();
    var rng_list = [];
    for (i = 0; i < rng.length; i++) {
      rng_list.push(rng[i][0]);
    };    
  } else if (sheet == "" && data_list != "") {
    var rng_list = extractColumn(data_list, column_number);
  } else {
    console.log("Error: not enough inputs. Exiting")
    return;
  };
  var noduplicates = new Set(rng_list);
  var unique_values = [];
  noduplicates.forEach(x => unique_values.push(x));
  /*
  // Remove first value, which normally would be the column header
  var first_value = unique_values[0];
  var final_list = unique_values.filter((value) => value != first_value);
  return final_list.sort();
  */
  return unique_values.sort();
};


function extractColumn(arr, column) {
  return arr.map(x => x[column]);
};


// https://stackoverflow.com/questions/17632165/determining-the-last-row-in-a-single-column
function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  };
  return letter;
};
function letterToColumn(letter) {
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  };
  return column;
};
function getLastDataColumn(sheet) {
  var lastCol = sheet.getLastColumn();
  var range = sheet.getRange(columnToLetter(lastCol) + "1");
  if (range.getValue() !== '') {
    return lastCol;
  } else {
    return range.getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).getColumn();
  };
};
function getLastDataRow(sheet, column_number) {
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(columnToLetter(column_number) + lastRow);
  if (range.getValue() !== '') {
    return lastRow;
  } else {
    return range.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  };
};
function getLastDataColumnRow(sheet, row) {
  //var column_letter = columnToLetter(column_number);
  var last_column = sheet.getLastColumn();
  var range = sheet.getRange(row, last_column);
  if (range.getValue() !== '') {
    return last_column;
  } else {
    return range.getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).getColumn();
  };
};


function getFirstRow(sheet, first_column) {
  if (sheet.getRange(1, first_column).getValue() == '') {
    return 1;
  } else {
    return getLastDataRow(sheet, first_column) + 3;
  };
};


function deleteData(sheet) {
  sheet.getRange('A2:' + columnToLetter(sheet.getLastColumn()) + (sheet.getLastRow() + 1)).deleteCells(SpreadsheetApp.Dimension.ROWS);
};


function clearData(sheet) {
  sheet.getRange('A2:' + columnToLetter(sheet.getLastColumn()) + (sheet.getLastRow() + 1)).clearContent();
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


function addHistoricalTotals() {
  for (var i = 48; i <= 51; i++) {
    console.log("Working row: " + i);
    // EAD WIP Summary
    console.log("EAD WIP Summary")
    for (var j = 3; j <= 22; j++) {
      var row_id = 7 + (i * 42);
      _gac_hd.getRange(row_id, j).setValue(_gac_hd.getRange(row_id - 3, j).getValue() + _gac_hd.getRange(row_id - 2, j).getValue() + _gac_hd.getRange(row_id - 1, j).getValue());
    };
    // NOS and OMAO Summary
    console.log("NOS and OMAO Summary")
    for (j = 3; j <= 22; j++) {
      row_id = 15 + (i * 42);
      _gac_hd.getRange(row_id, j).setValue(_gac_hd.getRange(row_id - 3, j).getValue() + _gac_hd.getRange(row_id - 2, j).getValue() + _gac_hd.getRange(row_id - 1, j).getValue());
    };
    // NOS and OMAO WIP Summary
    console.log("NOS and OMAO WIP Summary")
    for (j = 3; j <= 8; j++) {
      row_id = 24 + (i * 42);
      _gac_hd.getRange(row_id, j).setValue(_gac_hd.getRange(row_id - 4, j).getValue() + _gac_hd.getRange(row_id - 3, j).getValue() + _gac_hd.getRange(row_id - 2, j).getValue() + _gac_hd.getRange(row_id - 1, j).getValue());
    };
    // NWS Summary
    console.log("NWS Summary")
    for (j = 3; j <= 22; j++) {
      row_id = 32 + (i * 42);
      _gac_hd.getRange(row_id, j).setValue(_gac_hd.getRange(row_id - 3, j).getValue() + _gac_hd.getRange(row_id - 2, j).getValue() + _gac_hd.getRange(row_id - 1, j).getValue());
    };
    // NWS WIP Summary
    console.log("NWS WIP Summary")
    for (j = 3; j <= 8; j++) {
      row_id = 40 + (i * 42);
      _gac_hd.getRange(row_id, j).setValue(_gac_hd.getRange(row_id - 3, j).getValue() + _gac_hd.getRange(row_id - 2, j).getValue() + _gac_hd.getRange(row_id - 1, j).getValue());
    };
  };
};


// Insert data from a list into a sheet
function insertData(list, sheet) {
  if (list.length != 0) {
    sheet.getRange(2, 1, list.length, list[0].length).setValues(list);
    changeFont(sheet)
  };
  /*
  if (column_offset == undefined || column_offset == 0) {
    sheet.getRange(2, 1, list.length, list[0].length + column_offset).setValues(list);
  } else {
    for (var i = 0; i < list.length; i++) {
      for (var j = 0; j < list[0].length + column_offset; j++) {
        sheet.getRange(i + 2, j + 1).setValue(list[i][j]);
      };
    };    
  };
  */
};


// Save all data from sheet
function getDataValues(sheet, column_offset, row_offset) {
  return sheet.getRange("A2:" + columnToLetter(sheet.getLastColumn() + column_offset) + (sheet.getLastRow() + row_offset)).getValues();
};


// Change font to Times New Roman
function changeFont(sheet) {
  sheet.getRange("A:" + columnToLetter(sheet.getLastColumn())).setFontFamily('Times New Roman');
};


/*
// To be ran in each file
function onEdit(e) {
  var range = e.range;
  if (range.getColumn() == 13 && SpreadsheetApp.getActiveSheet().getName() == "WIP") {
    //SpreadsheetApp.getUi().alert(val);
    var new_value = vLookup(range.getValue(), SpreadsheetApp.openById("12Mavqu4igKzmuHD87HIEAiycDs4JOTaQ2lxpxixL_ss").getSheetByName("Data Values"), 6, 7, 0);
    SpreadsheetApp.getActiveSheet().getRange(range.getRow(), range.getColumn() + 1).setValue(new_value);
  };
};


function vLookup(search_value, sheet, col_search_num, col_return_num, case_match) {
  // Calculate last row depending on column searchedsheet
  var lrow = sheet.getRange(1, col_search_num).getDataRegion(SpreadsheetApp.Dimension.ROWS).getLastRow();
  for (var i = 2; i <= lrow; i++) {
    if (case_match == 0) {
      if (search_value.toLowerCase() == sheet.getRange(i, col_search_num).getValue().toLowerCase()) {
        return sheet.getRange(i, col_return_num).getValue();
      };
    } else {
      if (search_value == sheet.getRange(i, col_search_num).getValue()) {
        return sheet.getRange(i, col_return_num).getValue();
      };
    };
  };
};
*/

