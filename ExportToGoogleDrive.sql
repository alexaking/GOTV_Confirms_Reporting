INFO_SHEET = SpreadsheetApp.openById('10UxU_C0BjKCxXIZWTh85-R_dNYuGx95jIys6Mbk6QC4');  

// global variables
VAN_LINK_INDEX = 0; 
VAN_ID_INDEX = 1; 
SHIFT_INDEX = 11; 

function createTrigger() {
  ScriptApp.newTrigger("runWriteToSheets")
    .timeBased()
    .atHour(5)
    .nearMinute(0)
    .everyDays(1) // Frequency is required if you are using atHour() or nearMinute()
    .create();
}

function createTriggerclear() {
  ScriptApp.newTrigger("clearAllSheets")
    .timeBased()
    .atHour(23)
    .nearMinute(0)
    .everyDays(1) // Frequency is required if you are using atHour() or nearMinute()
    .create();
}

function runCollectDataFromSheets() { 

  Logger.log('runCollectDataFromSheets START'); 
  INFO_SHEET.getRange('Functions Control Panel!B11').setValue((new Date()).toString());
  INFO_SHEET.getRange('Functions Control Panel!B13').setValue("Running");
  
  try { 
    let SL_range = 'SL Locations!E2:E'; // gets all SL names
    var SL_array = getSLNamesArray(SL_range, INFO_SHEET); 
    var SL_dict = makeSLDict(SL_array); // empty dictionary with all SLs
    
    var spreadsheets = findSLSpreadsheets(SL_dict); // populated dictionary with SLs and their corresponding spreadsheets
    
    // get and clear sheet
    var export_sheet = INFO_SHEET.getSheetByName('CIVIS Export');
    export_sheet.clear(); 
    export_sheet.getRange('A1:B1').setValues([['Event Signup ID', 'Status']]); 
   
    for (var location in spreadsheets) { 
      var sheets = getSheets(SpreadsheetApp.openById(spreadsheets[location])); 
      // const sheet = SpreadsheetApp.openById(spreadsheets[location]).getSheetByName('Sheet1'); // sheet with data
      
      // get event signup ids and statuses (columns B and J), and filter
      let data = []; 
      const event_signup_ids = sheets.confirm_sheet.getRange('B2:B').getValues().filter(function (el) { return el[0] != ''; });
      data = event_signup_ids; 
      const statuses = sheets.confirm_sheet.getRange('K2:K').getValues().filter(function (el) { return el[0] != ''; }); ;
      
      if (event_signup_ids.length == statuses.length && data.length != 0) { 
        // put columns B and K together into one array
        data.forEach((row, index) => { 
                     row.push(statuses[index][0]);
        row.push(location); 
        }); 
        
        // write to master sheet
        export_sheet.getRange(export_sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data); 
      } 
      else { 
        Logger.log('Number of event signup IDs and statuses is not the same'); 
      }  
    }
    
    INFO_SHEET.getRange('Functions Control Panel!B13').setValue("Success");
  }
  catch (err) { 
    INFO_SHEET.getRange('Functions Control Panel!B13').setValue("Error");
  } 
  
  Logger.log('runCollectDataFromSheets END'); 
}


function clearAllSheets() { 

  Logger.log('clearAllSheets() START'); 
  INFO_SHEET.getRange('Functions Control Panel!B18').setValue((new Date()).toString());
  INFO_SHEET.getRange('Functions Control Panel!B20').setValue("Running");
  
  try { 
    let SL_range = 'SL Locations!E2:E'; // gets all SL names
    var SL_array = getSLNamesArray(SL_range, INFO_SHEET); 
    var SL_dict = makeSLDict(SL_array); // empty dictionary with all SLs
    
    var spreadsheets = findSLSpreadsheets(SL_dict); // populated dictionary with SLs and their corresponding spreadsheets
    
    for (var location in spreadsheets) { 
      var spreadsheet = SpreadsheetApp.openById(spreadsheets[location]); 
      
      var sheets = getSheets(spreadsheet); 
      clearSpreadsheet(location, sheets); 
     
    } 
    
    INFO_SHEET.getRange('Functions Control Panel!B20').setValue("Success");
 
  }
  catch (err) { 
    INFO_SHEET.getRange('Functions Control Panel!B20').setValue("Error");
  } 
  
  Logger.log('clearAllSheets() END'); 

} 


function runWriteToSheets() {  
  
  Logger.log('runWriteToSheets START'); 
  INFO_SHEET.getRange('Functions Control Panel!B4').setValue((new Date()).toString());
  INFO_SHEET.getRange('Functions Control Panel!B6').setValue("Running");
  
  var is_civis = true; // true if from CIVIS VAN tables, false if from BQ tables
  var is_van = false; 
  
  let import_date = INFO_SHEET.getRange('Civis Import Timestamp!A2').getValue().toDateString();
  let current_date = new Date().toDateString(); 
  
  if (import_date == current_date) { 
  
      try {
        let SL_range = 'SL Locations!E2:E'; // gets all SL names
        var SL_array = getSLNamesArray(SL_range, INFO_SHEET); 
        var SL_dict = makeSLDict(SL_array); // empty dictionary with all SLs
        
        var spreadsheets = findSLSpreadsheets(SL_dict); // populated dictionary with SLs and their corresponding spreadsheets
        
        if (is_civis) { 
          var sl_data_range = 'confirms from raw!M2:M'; // gets all SL names that have data
          var data_range = 'confirms from raw!A2:L'; // gets all data from master spreadsheet
          Logger.log('Populate tables based on CIVIS VAN data'); 
        }
        if (!is_civis) { 
          var sl_data_range = 'Confirms from BQ!N2:N'; // gets all SL names that have data
          var data_range = 'Confirms from BQ!A2:M'; // gets all data from master spreadsheet
          Logger.log('Populate tables based on BQ VAN data'); 
        }
        var info_dict = getInfo(sl_data_range, data_range, INFO_SHEET); // all data keyed by SL
        
        var confirm_headers = [['VAN ID', 'Event Signup ID', 'Name', 'Landline Phone', 'Cell Phone', 'Email', 'Shift this Weekend', 'Total Shifts this Weekend', 'Next Shift', 'Total Shifts After Today', 'Status', 'Status Last Updated', 'Re-shifts (Date and Time)', 'Notes', 'Pass 1', 'Pass 2', 'Pass 3']];
        var walk_in_headers = [['Name', 'Phone', 'Email', 'Shift Date', 'Shift Time', 'Re-shifts (Date and Time)', 'Notes']];
      
        // dates for color coding volunteer shifts 
        const sat_date = new Date('October 17, 2020'); 
        const sun_date = new Date('October 18, 2020'); 
        var dates = []; 
        dates.push(sat_date); 
        dates.push(sun_date); 
        
        for (var location in spreadsheets) { 
          Logger.log('Iteration is: ' + location);
          var spreadsheet = SpreadsheetApp.openById(spreadsheets[location]); 
          spreadsheet.setSpreadsheetTimeZone('America/New_York');
          location = location.substr(0,4) //gets the shortname for location- 'digi' for digital
          if (location== 'VSL_'){location = 'Digital'};
//          Logger.log(info_dict[location]);
          if (info_dict[location] != null) { 
              var sheets = getSheets(spreadsheet); 
              clearSpreadsheet(location, sheets); 
              populateSLSpreadsheet(location, sheets, info_dict[location], confirm_headers, walk_in_headers, dates, is_van); 
          } 
        }
        
        INFO_SHEET.getRange('Functions Control Panel!B6').setValue('Success');
      }
      catch (err) { 
        INFO_SHEET.getRange('Functions Control Panel!B6').setValue('Error');
        throw(err);
      }
   } 
   else {  
     Logger.log('Data not updated'); 
     INFO_SHEET.getRange('Functions Control Panel!B6').setValue('Data not updated');
   }
  
   Logger.log('runWriteToSheets END'); 
}
////////////////////////////

function findSLSpreadsheets(SL_dict) { 

  Logger.log('findSLSpreadsheets START'); 

  var GOTVfolder = DriveApp.getFoldersByName('VA GOTV Confirm Trackers').next(); 
  var region_folders = GOTVfolder.getFolders(); // FolderIterator
  
  while (region_folders.hasNext()) { 
  
    var folder = region_folders.next(); // next folder
    //Logger.log(folder.getName());
    
    var files = folder.getFiles(); // FileIterator
    
    while (files.hasNext()) {
    
      var file = files.next(); // next file
      SL_dict[file.getName()] = file.getId(); 
      //Logger.log(file.getName());
    }
  }
  
 Logger.log('findSLSpreadsheets END'); 
 return SL_dict; 
} 


function getSLNamesArray(SL_range, sheet) {

  Logger.log('getSLNamesArray START'); 
  
  var SL_info = sheet.getRange(SL_range).getValues(); 
  
  if(!SL_info) {
    Logger.log('Something went wrong, null staging locations');
  }
  
  var SL_array = SL_info.flat().filter(x => x!=''); // put into array
  
  Logger.log('%s staging locations found found', SL_array.length);
  
  Logger.log('getSLNamesArray END'); 
  
  return SL_array; 
}

function makeSLDict(SL_array) { 
  
  Logger.log('getSLNamesDict START'); 
  
  var SL_dict = {}; 
  
  SL_array.forEach((SL) => {SL_dict[SL] = []}); // create dictionary
  
  Logger.log('getSLNamesDict END'); 
  
  return SL_dict; 

} 

function getInfo(sl_range, data_range, sheet, is_van) { 

  Logger.log('getInfo START'); 

  var sl_data = sheet.getRange(sl_range).getValues(); 
  
  var sl_data_filtered = sl_data.flat().filter(x => x!=''); // put into array, take out empty values
  
  Logger.log('%s locations found', sl_data_filtered.length);
  
  var info_dict = {}; 
  
  sl_data_filtered.forEach((info) => {info_dict[info] = []}); // create dictionary
  
  Logger.log('%s locations in array', Object.keys(info_dict).length);
  
  var data = sheet.getRange(data_range).getValues();
  var data_filtered = data.filter(function (el) {return el[1] != "";}); // filter all the data so no empty elements
  
  if(data_filtered.length != sl_data_filtered.length) { 
    Logger.log('Number of staging locations not the same as information'); 
    return;
  } 
  
  var van_link = 'https://www.votebuilder.com/ContactsDetails.aspx?VANID=EID'; 
  
  for(var i = 0; i < data_filtered.length; i++) {

    sl_loc = sl_data_filtered[i]; // staging location
    
    data_filtered[i][VAN_LINK_INDEX] = van_link + data_filtered[i][VAN_LINK_INDEX];  // concatenate VAN link
    if (data_filtered[i][SHIFT_INDEX] == 'Sched-Web' || data_filtered[i][SHIFT_INDEX] == 'Tentative' || data_filtered[i][SHIFT_INDEX] == 'Invited') { 
      data_filtered[i][SHIFT_INDEX] = 'Scheduled'; 
    } 
    if (data_filtered[i][SHIFT_INDEX] == 'Confrm-Web' || data_filtered[i][SHIFT_INDEX] == 'Conf Twice') { data_filtered[i][SHIFT_INDEX] = 'Confirmed'; } 
    
    if (data_filtered[i][SHIFT_INDEX] == 'Walk in') { 
      Logger.log('VAN event is %s', data_filtered[i][SHIFT_INDEX]); 
      Logger.log('event signup id %s', data_filtered[i][2]); 
    }
    else { 
      try { 
        info_dict[sl_loc].push(data_filtered[i]);
        //Logger.log(data_filtered[i]); 
      } 
      catch(err) {
        Logger.log("ERROR at row %s", i);
        Logger.log("%s: %s", err.name, err.message);
        throw(err);
      }
    }
  }
  
  Logger.log('getInfo END'); 
  return info_dict; 
}


/////////////

function populateSLSpreadsheet (location, sheets, sl_info_dict, confirm_headers, walk_in_headers, dates, is_van) { 

  Logger.log('populateSLSpreadsheet START %s', location); 
  
  // headers formatting
  sheets.confirm_sheet.getRange(1, 1, confirm_headers.length, confirm_headers[0].length).setValues(confirm_headers); 
  sheets.confirm_sheet.getRange(1, 1, confirm_headers.length, confirm_headers[0].length).setBackground("#4285f4"); // set background to blue
  sheets.confirm_sheet.getRange(1, 1, confirm_headers.length, confirm_headers[0].length).setFontColor("#ffffff"); // set font color to white
  sheets.walk_in_sheet.getRange(1, 1, walk_in_headers.length, walk_in_headers[0].length).setValues(walk_in_headers); 
  sheets.walk_in_sheet.getRange(1, 1, walk_in_headers.length, walk_in_headers[0].length).setBackground("#4285f4"); // set background to blue
  sheets.walk_in_sheet.getRange(1, 1, walk_in_headers.length, walk_in_headers[0].length).setFontColor("#ffffff"); // set font color to white
          
  // write VAN links to sheets
  let links = sl_info_dict.map(function(value,index) {return value[0];}); // 1-D array of VAN links
  let ids = sl_info_dict.map(function(value,index) {return value[1];}); // 1-D array of VAN IDs
  
  if(links.length == ids.length) { 
    if(is_van) { 
      var values = []; 
      links.forEach((link, index) => { values.push([SpreadsheetApp.newRichTextValue().setText(ids[index]).build()]); }); // build hyperlinks to go into Sheets
      sheets.confirm_sheet.getRange(2, 1, values.length, values[0].length).setRichTextValues(values); 
    } 
    else { 
      var values = []; 
      links.forEach((link, index) => { values.push([SpreadsheetApp.newRichTextValue().setText(ids[index]).setLinkUrl(link).build()]); }); // build hyperlinks to go into Sheets
      sheets.confirm_sheet.getRange(2, 1, values.length, values[0].length).setRichTextValues(values); 
    }
  }
  else { 
    Logger.log('Something wrong with the number of rows in this location %s', location); 
    return; 
  }

  // write the rest of the data into sheets
  var new_info_dict = sl_info_dict.map(function(value, index) { return value.slice(2); }); // all the information except in the first 2 columns
  sheets.confirm_sheet.getRange(2, 2, new_info_dict.length, new_info_dict[0].length).setValues(new_info_dict);
  
  // autoresize columns
  sheets.confirm_sheet.autoResizeColumns(1, 10);
  sheets.confirm_sheet.setColumnWidth(11, 100); 
  sheets.confirm_sheet.autoResizeColumn(12); 
  sheets.confirm_sheet.setColumnWidths(13, 2, 200); 
  sheets.confirm_sheet.setColumnWidths(15, 3, 100); // set column 11 to 200 pixels
  
  sheets.walk_in_sheet.setColumnWidths(1, 7, 200); 
  
  // format dates and number of shifts (columns E:H)  sheets.walk_in_sheet
  sheets.confirm_sheet.getRange('G2:G').setNumberFormat('mmm dd h:mm am/pm');
  sheets.confirm_sheet.getRange('I2:I').setNumberFormat('mmm dd h:mm am/pm');
  sheets.confirm_sheet.getRange('L2:L').setNumberFormat('mmm dd h:mm am/pm');
  sheets.confirm_sheet.getRange('H2:H').setNumberFormat('0');
  sheets.confirm_sheet.getRange('J2:J').setNumberFormat('0');
  
  sheets.walk_in_sheet.getRange('D2:D').setNumberFormat('mmm dd');
  sheets.walk_in_sheet.getRange('E2:E').setNumberFormat('h:mm am/pm');
  sheets.walk_in_sheet.getRange('F2:F').setNumberFormat('mmm dd h:mm am/pm');
  
  // make column J (status) drop down menu w/ data validation
  const status_options = ['Scheduled', 'Left Msg', 'Confirmed', 'Conf Twice', 'Completed', 'No Show', 'Declined']; 
  const status_helpText = 'Please choose one of the options from the drop down menu: Scheduled, Confirmed, Left Msg, Tentative, Conf Twice, Completed, No Show, or Declined';
  const status_validation = SpreadsheetApp.newDataValidation().requireValueInList(status_options, true).setAllowInvalid(false).setHelpText(status_helpText).build();
  const status_range = sheets.confirm_sheet.getRange('K2:K'); 
  status_range.setDataValidation(status_validation);
  
  // conditional formatting color rules for column E (green if Saturday, yellow if Sunday)
  const shift_range = sheets.confirm_sheet.getRange('G2:G'); 
  const shift_colors = ['#d9ead3', '#fff2cc']; 
  let rules = sheets.confirm_sheet.getConditionalFormatRules();
  shift_colors.forEach((color, index) => { rules.push(SpreadsheetApp.newConditionalFormatRule().whenDateEqualTo(dates[index]).setBackground(color).setRanges([shift_range]).build()); }); 
  
  // conditional formatting color rules for column K (status) 
  const status_colors = ['#ffe599', '#F6B26B', '#d9ead3', '#d9ead3', '#6aa84f', '#ea9999', '#ea9999']; 
  status_options.forEach((option, index) => { rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(option).setBackground(status_colors[index]).setRanges([status_range]).build()); }); 
  sheets.confirm_sheet.setConditionalFormatRules(rules);
  
  // write Notes column
  //let column = new_info_dict[0].length + values[0].length + 2; // index to insert column at
  //sheet.insertColumns(column);
  //sheet.getRange('K1:L').mergeAcross();
  
  // write pass options
  const pass_options = ['Call', 'Text', 'Email'];
  const pass_helpText = 'Choose the method of voter contact (call, text or email)'; 
  const pass_validation = SpreadsheetApp.newDataValidation().requireValueInList(pass_options, true).setAllowInvalid(false).setHelpText(pass_helpText).build();
  const pass_range = sheets.confirm_sheet.getRange('O2:Q');
  pass_range.setDataValidation(pass_validation);
  
  // left align columns A through K
  sheets.confirm_sheet.getRange('A2:L').setHorizontalAlignment("left");
  
  // hide column B
  sheets.confirm_sheet.hideColumn(sheets.confirm_sheet.getRange('B1'));
  
  // lock header row 
  sheets.confirm_sheet.setFrozenRows(1); 
  // lock columns A:C 
  sheets.confirm_sheet.setFrozenColumns(6); 
  
  // don't allow edits on columns A through J except for me
  var protection = sheets.confirm_sheet.getRange('A2:J').protect().setDescription('Protect volunteer information cells');
  var protection_date = sheets.confirm_sheet.getRange('L2:L').protect().setDescription('shift status last updated'); 
  //var me = Session.getEffectiveUser();
  //protection.addEditor(me);
  protection.addEditors(['alexa.king@2020victory.com']); 
  protection_date.addEditors(['alexa.king@2020victory.com']); 
  /*protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }*/
    
  Logger.log('populateSLSpreadsheets END %s', location);
}



function getSheets(spreadsheet) { 
  
  var confirm_sheet = spreadsheet.getSheetByName('Confirm List');
  var walk_in_sheet = spreadsheet.getSheetByName('Walk ins'); 
  
  if (!confirm_sheet) { 
    spreadsheet.getSheetByName('Sheet1').setName('Confirm List');
    confirm_sheet = spreadsheet.getSheetByName('Confirm List');
  }
  if (!walk_in_sheet) { 
    spreadsheet.insertSheet('Walk ins'); 
    walk_in_sheet = spreadsheet.getSheetByName('Walk ins'); 
  }
  
  return {confirm_sheet, walk_in_sheet}; 
} 


function clearSpreadsheet(location, sheets) { 
  Logger.log('clearSpreadsheet START %s', location); 
  
  // clear formatting, contents, and data validations
  sheets.confirm_sheet.clear(); 
  sheets.walk_in_sheet.clear(); 
  sheets.confirm_sheet.getRange('A2:P').clearDataValidations(); 
  sheets.walk_in_sheet.getRange('A2:P').clearDataValidations(); 
  
  var protections = sheets.confirm_sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i];
    if (protection.canEdit()) {
      protection.remove();
    }
  }
  
  Logger.log('clearSpreadsheet END %s', location); 
  
} 



function makeSLSpreadsheets(SL_dict) { 
  
  Logger.log('makeSLSpreadsheets START'); 
  
  for (var location in SL_dict) { 
    var ss_temp = SpreadsheetApp.create(location);
    SL_dict[location] = ss_temp.getId(); 
  }
  
  Logger.log('makeSLSpreadsheets END'); 
  
  return SL_dict;  
} 
