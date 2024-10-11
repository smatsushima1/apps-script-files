

function getVariables() {
  _ga = SpreadsheetApp.getActive();
  _cp = _ga.getSheetByName("Clauses and Provisions");
};


function uncheckAll() {
  getVariables();
  _cp.getRange("G:G").uncheck();
  // Reset filter
  var criteria = SpreadsheetApp.newFilterCriteria().build();
  _cp.getFilter().setColumnFilterCriteria(7, criteria);
  repositionButtons();
};


function showSelected() {
  getVariables();
  var criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(['', 'FALSE']).build();
  _cp.getFilter().setColumnFilterCriteria(7, criteria);
  _cp.getRange(8, 7).activate();
  repositionButtons();
};


function allSolicitations() {
  getVariables();
  allSolicitationsContracts("H");
  repositionButtons();
}


function allContracts() {
  getVariables();
  allSolicitationsContracts("I");
  repositionButtons();
};


function allSolicitationsContracts(column_letter) {
  var lrow = _cp.getLastRow();
  var include_list = _cp.getRange(column_letter + "9:" + column_letter + lrow).getValues();
  var prescription_list = _cp.getRange("F9:G" + lrow).getValues();
  var comm_list = _cp.getRange("J9:J" + lrow).getValues();
  // Update old values list
  for (var i = 0; i < include_list.length; i++) {
    if (prescription_list[i][0].length > 0 && (prescription_list[i][1] == true || prescription_list[i][1] == false)) {
      // Only update affected values
      if (include_list[i][0].length > 0 && comm_list[i][0].length == 0) {
        prescription_list[i][1] = true;
      };
    };
  };
  // Recreate new list with values to insert
  var select_list = [];
  for (i = 0; i < prescription_list.length; i++) {
    select_list.push([prescription_list[i][1]]);
  };
  // Save new values in the current table
  _cp.getRange(9, 7, select_list.length, 1).setValues(select_list);
};


function selectAll() {
  getVariables();
  // Save variables
  var dv = _cp.getRange(2, 4).getValue();
  var sc = _cp.getRange(3, 4).getValue();
  var mpt = _cp.getRange(4, 4).getValue();
  var sat = _cp.getRange(5, 4).getValue();
  var comm = _cp.getRange(6, 4).getValue();
  // First run all solicitations and contracts
  if (sc == "Solicitation") {
    allSolicitationsContracts("H");
    allSolicitationsContracts("I");   
  } else {
    allSolicitationsContracts("I"); 
  };
  // Save all clauess and provisions in separate list to be compared against based on sc
  var lrow = _cp.getLastRow();
  var cp_list = _cp.getRange("E9:E" + lrow).getValues();
  // Loop through all criteria
  var select_list = _cp.getRange("G9:G" + lrow).getValues();
  var criteria_list = _cp.getRange("J9:M" + lrow).getValues();
  for (var i = 0; i < cp_list.length; i++) {
    var c_comm = criteria_list[i][0];
    var c_mpt = criteria_list[i][1];
    var c_sat = criteria_list[i][2];
    var c_dv = criteria_list[i][3];
    // First skip over all provisions if selecting for contracts
    if (sc == "Contract" && cp_list[i][0] == "P") {
      continue;
    // Check for commerciality
    } else if (comm == "Commercial") {
      if (c_comm.length > 0) {
        // If either of the MPT or SAT criteria is met and is required, then mark as TRUE
        if ((mpt == "Yes" && c_mpt.length > 0) || (sat == "Yes" && c_sat.length > 0)) {
          select_list[i] = [true];
        // Dollar values
        } else if (String(c_dv).length > 0 && dv > c_dv) {
          select_list[i] = [true];
        // Commerciality outside of MPT and SAT
        } else if (c_mpt.length == 0 && c_sat.length == 0) {
          select_list[i] = [true];
        };
      // Check for dollar values outside of commerciality 
      } else if (c_comm.length == 0) {
        // MPT and SAT
        if ((mpt == "Yes" && c_mpt.length > 0) || (sat == "Yes" && c_sat.length > 0)) {
          select_list[i] = [true];
        // Dollar values
        } else if (String(c_dv).length > 0 && dv > c_dv) {
          select_list[i] = [true];
        };
        /*
        } else if (c_mpt.length == 0 && c_sat.length == 0) {
          select_list[i] = [true];
        */
      };
    } else if (comm == "Non-Commercial") {
      // Automatically skip over all commercial items
      if (c_comm.length > 0) {
        continue;
      // MPT and SAT
      } else if ((mpt == "Yes" && c_mpt.length > 0) || (sat == "Yes" && c_sat.length > 0)) {
        select_list[i] = [true];
      // Dollar values
      } else if (String(c_dv).length > 0 && dv > c_dv) {
        select_list[i] = [true];
      };
    };
  };
  // Save new values in the current table
  _cp.getRange(9, 7, select_list.length, 1).setValues(select_list);
  repositionButtons();
};


function repositionButtons() {
  var drawings = _cp.getDrawings();
  // All Criteria
  drawings[2].setPosition(2, 5, 25, 0);
  // Reset Clauses
  drawings[1].setPosition(2, 7, -325, 0);
  // Generate Clauses
  drawings[0].setPosition(2, 7, -125, 0);
};


// Hyperlink format: https://www.acquisition.gov/far/part-28#FAR_28_102_1
function createHyperlinks() {
  getVariables();
  for (var i = 9; i <= 198; i++) {
    var rng = _cp.getRange(i, 3);
    var ref = String(rng.getValue());
    // Move to next item if no prescription
    if (ref.length == 0) {
      continue;
    }
    var part_i = ref.indexOf(".");
    var part = ref.substring(0, part_i);
    var subpart_i_1 = ref.indexOf("-");
    var subpart_i_2 = ref.indexOf("(");
    var subpart_i = "";
    // First check if there is no - or (
    if (subpart_i_1 == -1 && subpart_i_2 == -1) {
      subpart_i = ref.length + 1;
    // Only - exists
    } else if (subpart_i_1 > 0 && subpart_i_2 == -1) {
      subpart_i = subpart_i_1 + 1;
    // Only ( exists
    } else if (subpart_i_2 > 0 && subpart_i_1 == -1) {
      subpart_i = subpart_i_2 + 1;
    // Both exists, but - comes first
    } else if (subpart_i_1 > 0 && subpart_i_2 > 0 && subpart_i_1 < subpart_i_2) {
      subpart_i = subpart_i_1 + 1;
    // Both exists, but ( comes first
    } else if (subpart_i_1 > 0 && subpart_i_2 > 0 && subpart_i_2 < subpart_i_1) {
      subpart_i = subpart_i_2 + 1;
    };
    // For some reason, - and ( are still popping up
    var subpart = ref.substring(part_i + 1, subpart_i - part_i + 1).replace(/-/g, '').replace(/\(/g, '');
    var hyperlink = "https://www.acquisition.gov/far/part-" + part + "#FAR_" + part + "_" + subpart;
    // Add section,if applicable
    var section_i_end = "";
    if (subpart_i_1 > 0) {
      // First save section if no (
      if (subpart_i_2 == -1) {
        section_i_end = ref.length + 1;
      } else if (subpart_i_2 > subpart_i_1) {
        section_i_end = subpart_i_2;
      };
      var section = ref.substring(subpart_i_1 + 1, section_i_end);
      hyperlink += "_" + section;
    };
    // Update hyperlink in each cell
    var rich_value = SpreadsheetApp.newRichTextValue().setText(ref).setLinkUrl(hyperlink).build();
    rng.setRichTextValue(rich_value);
  };
};

