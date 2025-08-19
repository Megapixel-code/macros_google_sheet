/** @OnlyCurrentDoc */

function strict_modulo(x, m){
  return ((x % m) + m) % m;
}
function days_in_month(month, year) {
  // Use 0 for January, 1 for February, etc.
  // return the number of days of the month
  date = new Date(year, month + 1, 0)
  return date.getDate();
}
function get_day(month, year) {
  // Use 0 for January, 1 for February, etc.
  // returns 0 for sunday, 1 for monday etc.
  date = new Date(year, month, 1);
  return date.getDay();
}
function get_letter_of_day(day, month, year){
  // Use 0 for January, 1 for February, etc.
  // returns A for monday, B for tuesday etc.
  var position_of_first_day = (get_day(month, year) - 1); // 0 is monday
  var position_of_day = strict_modulo((position_of_first_day + (day - 1)), 7);
  return String.fromCharCode(65 + position_of_day);
}


function onEdit(e){
  if (onEdit_check(e)){
    return;
  }
}
function onEdit_check(e){
  if (!(e.range.isChecked()) || !(e.value.includes("#SCRIPT_"))){
    return false;
  }
  const sheet = e.range.getSheet();
  const spreadsheet = sheet.getParent();

  var script_name = e.value.replace("#SCRIPT_", "");
  Utilities.sleep(1000);
  switch (script_name){
    case "CALENDAR_INIT":
      calendar_init();
      break;
    case "TODAY":
      update_color();
      break;
    case "UPDATE":
      calendar_update(sheet);
      break;
  }

  // Uncheck the checkbox and return true to indicate we handled the event
  e.range.uncheck();
  return true;
}


function calendar_init() {
  var spreadsheet = SpreadsheetApp.getActive();

  // create calendar sheet
  var year = Math.floor(spreadsheet.getRange('\'PROGRAMME\'!B4').getValue());
  var sheet_name = year + '-' + (year+1);
  
  if (spreadsheet.getSheetByName(sheet_name)){
    spreadsheet.getSheetByName(sheet_name).activate();
    if (spreadsheet.getRange("Z2").getValue()){ // if correctly init
      update_color();
      return false;
    }
    else { // else, re-init the sheet
      spreadsheet.deleteSheet(spreadsheet.getSheetByName(sheet_name));
    }
  }

  spreadsheet.insertSheet(1);
  spreadsheet.getActiveSheet().setName(sheet_name);

  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(sheet_name), true);
  spreadsheet.getRange('Z1').setValue(year);
  
  spreadsheet.getRange('Y1').activate();
  spreadsheet.getRange('\'NE PAS TOUCHER TEMPLATES\'!Y1:Y3').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(sheet_name), true);
  spreadsheet.getSheetByName(sheet_name).activate();
  spreadsheet.getRange('\'NE PAS TOUCHER TEMPLATES\'!A1:M202').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  // copies the correct months numbers
  for (var i = 8; i < 20; i++){
    var current_month_start_row = ((i - 8) * 17) + 1;
    var month = i % 12;
    if (i < 12){
      var current_year = year;
    }
    else{
      var current_year = year + 1;
    }
    var start = get_day(month, current_year);

    spreadsheet.getRange('A' + (current_month_start_row + 2)).activate();
    spreadsheet.getRange('\'NE PAS TOUCHER TEMPLATES\'!O'+ ((start*12)+1) +':U'+ ((start+1)*12)).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    
    if (days_in_month(month, current_year) != 31){
      var range = [];
      var nb_rows = Math.floor((strict_modulo(get_day(month, current_year) - 1, 7) + (days_in_month(month, current_year) - 1)) / 7);
      var offset = 0;
      if (nb_rows == 4){
        range.push('A'+ (current_month_start_row + 12) +':G' + (current_month_start_row + 13));
      }
      if (month == 11){
        var last_line_range = get_letter_of_day(1, 0, current_year + 1);
      }
      else{
        var last_line_range = get_letter_of_day(1, month + 1, current_year);
      }

      if (last_line_range == "A"){
        offset += 2;
      }

      last_line_range += ((nb_rows * 2) + 2 + current_month_start_row + offset) + ':G' + ((nb_rows * 2) + 3 + current_month_start_row + offset);
      range.push(last_line_range);

      spreadsheet.getRangeList(range).setBackground(null).clear({contentsOnly: true, skipFilteredRows: true});
    }
  }

  // ================================================================================ TODO add later
  // resize rows
  for (var x = 0; x < 12; x++){
    for (var i = 0; i < 6; i++){
      spreadsheet.getActiveSheet().setRowHeight(((x * 17) + 4) + (i * 2), 60);
    }
  }
  // resize columns
  // calendar columns
  for (var x = 0; x < 7; x++){
    spreadsheet.getActiveSheet().setColumnWidth(x + 1, 178);
  }
  // comments column
  spreadsheet.getActiveSheet().setColumnWidth(8, 178);
  // validation culumns
  for (var x = 0; x < 2; x++){
    spreadsheet.getActiveSheet().setColumnWidth(x + 9, 60);
  }

  calendar_update(spreadsheet.getActiveSheet());

  spreadsheet.getSheetByName(sheet_name).getRange('Z2').setValue(true);
  spreadsheet.toast("le calendrier a été initialisé correctement !");

  return true;
};


function set_current_month_color() {
  var spreadsheet = SpreadsheetApp.getActive();

  // ========================================= Set colors for current month
  var dark_color = '#6aa84f';
  var light_color = '#b6d7a8';

  var date = new Date();
  var current_year = date.getFullYear();
  var current_month = date.getMonth();
  var current_day = date.getDate();

  var current_month_start_row = (17 * ((current_month + 4) % 12 )) + 1;
  var nb_rows = Math.floor((strict_modulo(get_day(current_month, current_year) - 1, 7) + (days_in_month(current_month, current_year) - 1)) / 7);

  // set color top
  spreadsheet.getRangeList(['A'+ current_month_start_row +':G' + (current_month_start_row+1), 'K'+ (current_month_start_row + 1) +':M' + (current_month_start_row + 1)]).activate();
  spreadsheet.getActiveRangeList().setBackground(dark_color);
  // set color bottom
  // center range
  var light_color_range = ['A'+ (current_month_start_row+5) +':G' + (current_month_start_row+5), 'A'+ (current_month_start_row+7) +':G' + (current_month_start_row+7), 'A'+ (current_month_start_row+9) +':G' + (current_month_start_row+9)];
  if (nb_rows == 5){
    light_color_range.push('A'+ (current_month_start_row+11) +':G' + (current_month_start_row+11))
  }
  
  // first line range
  light_color_range.push(get_letter_of_day(1, current_month, current_year) + (current_month_start_row+3) +':G'+ (current_month_start_row+3));
  
  // last line range
  if (nb_rows == 4){ // always 4 or 5 nb_rows
    var offset = 11;
  }
  else{
    var offset = 13;
  }
  var last_line_range = 'A'+ (current_month_start_row+offset) +':';
  if (current_month == 11){
    last_line_range += get_letter_of_day(0, 0, current_year + 1);
  }
  else{
    last_line_range += get_letter_of_day(0, current_month + 1, current_year);
  }
  last_line_range += (current_month_start_row+offset);

  light_color_range.push(last_line_range);
  spreadsheet.getRangeList(light_color_range).activate();
  spreadsheet.getActiveRangeList().setBackground(light_color);

  // set current day color
  var row = Math.floor((strict_modulo(get_day(current_month, current_year) - 1, 7) + (current_day - 1)) / 7);
  var range = get_letter_of_day(current_day, current_month, current_year) + (((row + 1) * 2) + 1 + current_month_start_row);
  spreadsheet.getRange(range).setBackground(dark_color);

  spreadsheet.getRange("Z3").setValue(true); // this sheet contains green

  spreadsheet.getRange(range).activate();
}


function reset_color(sheet_name){
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(sheet_name), true);

  var year = Math.floor(spreadsheet.getRange('Z1').getValue());
  
  // ========================================= Set colors
  var dark_color = '#ff66cc';
  var light_color = '#ffccff';

  var light_color_range = [];
  var dark_color_range = [];

  for (var i = 8; i < 20; i++){
    var current_month = i % 12;
    if (i < 12){
      var current_year = year;
    }
    else{
      var current_year = year + 1;
    }

    var current_month_start_row = (17 * ((current_month + 4) % 12 )) + 1;
    var nb_rows = Math.floor((strict_modulo(get_day(current_month, current_year) - 1, 7) + (days_in_month(current_month, current_year) - 1)) / 7);

    // set color top
    dark_color_range.push('A'+ current_month_start_row +':G' + (current_month_start_row+1), 'K'+ (current_month_start_row + 1) +':M' + (current_month_start_row + 1));
    
    // set color bottom
    // center range
    light_color_range.push('A'+ (current_month_start_row+5) +':G' + (current_month_start_row+5), 'A'+ (current_month_start_row+7) +':G' + (current_month_start_row+7), 'A'+ (current_month_start_row+9) +':G' + (current_month_start_row+9));
    if (nb_rows == 5){
      light_color_range.push('A'+ (current_month_start_row+11) +':G' + (current_month_start_row+11))
    }
    
    // first line range
    light_color_range.push(get_letter_of_day(1, current_month, current_year) + (current_month_start_row+3) +':G'+ (current_month_start_row+3));
    
    // last line range
    if (nb_rows == 4){ // always 4 or 5 nb_rows
      var offset = 11;
    }
    else{
      var offset = 13;
    }
    var last_line_range = 'A'+ (current_month_start_row+offset) +':';
    if (current_month == 11){
      last_line_range += get_letter_of_day(0, 0, current_year + 1);
    }
    else{
      last_line_range += get_letter_of_day(0, current_month + 1, current_year);
    }
    last_line_range += (current_month_start_row+offset);

    light_color_range.push(last_line_range);
  }

  spreadsheet.getRangeList(dark_color_range).activate();
  spreadsheet.getActiveRangeList().setBackground(dark_color);

  spreadsheet.getRangeList(light_color_range).activate();
  spreadsheet.getActiveRangeList().setBackground(light_color);

  spreadsheet.getRange("Z3").setValue(false); // this sheet does not contains green
}


function update_color(){
  var spreadsheet = SpreadsheetApp.getActive();
  var sheets = spreadsheet.getSheets();

  var b_current_year_exist = false;

  for (var i = 0; i < sheets.length; i++){
    var sheet_name = sheets[i].getName();
    if (sheet_name[0] >= '0' && sheet_name[0] <= "9"){  
      sheets[i].activate();
      if (spreadsheet.getRange("Z3").getValue()){ // if there is green on the sheet
        reset_color(sheet_name);
        spreadsheet.getRange("A1:G1").activate();
      }
      if (spreadsheet.getRange("Y1").getValue()){ // if it's current year
        set_current_month_color();
        b_current_year_exist = true;
        var current_year_sheet_name = sheet_name;
      }
    }
  }

  if (b_current_year_exist){
    spreadsheet.getSheetByName(current_year_sheet_name).activate();

    var date = new Date();
    var current_year = date.getFullYear();
    var current_month = date.getMonth();
    var current_day = date.getDate();

    var current_month_start_row = (17 * ((current_month + 4) % 12 )) + 1;
    var row = Math.floor((strict_modulo(get_day(current_month, current_year) - 1, 7) + (current_day - 1)) / 7);
    var today_range = get_letter_of_day(current_day, current_month, current_year) + (((row + 1) * 2) + 1 + current_month_start_row);
    
    spreadsheet.getRange(today_range).activate();
  }
}


function update_passed_requests(){
  spreadsheet = SpreadsheetApp.getActive();
  var rows_to_treat = [];

  var i = 2;
  var row = spreadsheet.getRange('\'NE PAS TOUCHER DONNEES\'!Z' + i).getValue();
  while (row != "" && row != "#N/A"){
    rows_to_treat.push(Math.floor(row));
    i++;
    row = spreadsheet.getRange('\'NE PAS TOUCHER DONNEES\'!Z' + i).getValue();
  }

  for (var i = 0; i < rows_to_treat.length; i++){
    range = '\'NE PAS TOUCHER RESPONSES\'!K' + rows_to_treat[i];
    spreadsheet.getRange(range).setValue("PASSED");
  }
}

function calendar_update(sheet){
  spreadsheet = sheet.getParent();

  update_color();
}
