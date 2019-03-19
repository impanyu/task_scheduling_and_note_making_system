/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
 
var column_index={"priority":0,"assigning time":1,"requested starting time":2,"requested ending time":3,"starting time":4,"ending time":5,"task type":6,"task name":7,"task group":8,"task info":9,"link":10};
var archives_name=["papers","projects","projects2","readings","readings2","books","books2"];
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Assign Task...', functionName: 'assignTask'},
    {name: 'Sort Priority Queue...', functionName: 'sortTask'},
    {name: 'Archive Task...', functionName: 'archiveTask'}
   
  ];
  spreadsheet.addMenu('Directions', menuItems);
}


/**
 * A custom function that gets the driving distance between two addresses.
 *
 * @param {String} origin The starting address.
 * @param {String} destination The ending address.
 * @return {Number} The distance in meters.
 */
function drivingDistance(origin, destination) {
  var directions = getDirections_(origin, destination);
  return directions.routes[0].legs[0].distance.value;
}


/**
* A function that assigning the task into calendar
*
*/
function assignTask(){
   var priority_sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("priority_queue");  
   var all_priority_range = priority_sheet.getDataRange();
   var all_priority_data = all_priority_range.getValues();
   var eventCal=CalendarApp.getCalendarById("impanyu@gmail.com");
   var currentTime=new Date();
   for(i=1;i<all_priority_range.getNumRows();i++){//read all tasks for processing
        var current_row=all_priority_data[i];
        var requested_starting_time=current_row[column_index["requested starting time"]];
        var requested_ending_time=current_row[column_index["requested ending time"]];
        var requested_starting_date=new Date(requested_starting_time);
        if(requested_ending_time && requested_starting_time && currentTime<requested_starting_date) {//find a task need to be registered into calendar
           var task_name="";
           var link=current_row[column_index["link"]];
           var task_info="info: "+current_row[column_index["task info"]];
           if(link) task_info+="\nlink: "+link;
           if(current_row[column_index["task group"]]) task_name+=current_row[column_index["task group"]]+" ";
           task_name+=current_row[column_index["task name"]];
           eventCal.createEvent(task_name, requested_starting_time, requested_ending_time,{description:task_info});  
           
       }
   }
}


/**
* A function that archive completed task into corresponding archives and calendar
*
*/
function archiveTask(){
    var style = SpreadsheetApp.newTextStyle()
    //.setForegroundColor("red")
    .setFontSize(11)
    //.setBold(true)
    //.setUnderline(true)
    .build();
   var archived_task={};
   var archives={};
   
   for(var [i,name] in archives_name){
         archived_task[name]=[];
         archives[name]=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
   }
   //console.log(archived_task);
   var priority_sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("priority_queue");  
   var all_priority_range = priority_sheet.getDataRange();
   var all_priority_data = all_priority_range.getValues();
   var new_priority_data = [];
   var cols = all_priority_range.getNumColumns();
   var rows = all_priority_range.getNumRows();
   var currentTime=new Date();
   
   for(i=1;i<rows;i++){//read all tasks for processing
        var current_row=all_priority_data[i];
        var data_type=current_row[column_index["task type"]];
        var ending_time=current_row[column_index["ending time"]];
        if(ending_time && data_type) {//find a task needed to be archived 
             current_row.splice(column_index["task type"],1);
             archived_task[data_type].push(current_row);       
       }
       else{
        new_priority_data.push(current_row);
       }
   }
   
   for (var key in archives) {//archive all tasks into corresponding archives
    // check if the property/key is defined in the object itself, not in parent
    
    if (archives.hasOwnProperty(key)) {
      var value=archives[key];
      var archive_rows=value.getDataRange().getNumRows();
      var archive_cols=value.getDataRange().getNumColumns();
      var tasks=archived_task[key];
      if(tasks.length>0){
           //console.log(tasks);
           value.getRange(archive_rows+1,1,tasks.length,tasks[0].length).setValues(tasks);
           value.getRange(archive_rows+1,1,tasks.length,tasks[0].length).setTextStyle(style);
           }
    }
  }
  
  priority_sheet.getRange(2,1,rows,cols).clear();
  priority_sheet.getRange(2,1,rows,cols).setTextStyle(style);
  priority_sheet.getRange(2, 1, new_priority_data.length, new_priority_data[0].length).setValues(new_priority_data); 
  
   var eventCal=CalendarApp.getCalendarById("impanyu@gmail.com");
   for (var key in archives) {//writes archived tasks into calendar
       var tasks=archived_task[key];
       for([i,task] in tasks){
         var starting_time=task[column_index["starting time"]];
         var ending_time=task[column_index["ending time"]];
         var task_name="";
         var link=task[column_index["link"]-1];
         var task_info="info: "+task[column_index["task info"]-1];
         if(link) task_info+="\nlink: "+link;
         if(task[column_index["task group"]-1]) task_name+=task[column_index["task group"]-1]+" ";
         task_name+=task[column_index["task name"]-1];
         console.info(starting_time);
         eventCal.createEvent(task_name, starting_time, ending_time,{description:task_info});
       }
   }
}


/**
 * A function that sort the spreadsheet according to calculated priority values.
 */
function sortTask() {
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("priority_queue");
  //var sheet = SpreadsheetApp.getActiveSheet();//.setName('Settings');
  var col_num=sheet.getDataRange().getNumColumns();
  var row_num=sheet.getDataRange().getNumRows();
  var priority_relevant_cols=sheet.getRange(1,1,row_num,3).getValues();
  var calculated_priority=[];
  var currentTime=new Date();
  for(i=0;i<row_num;i++){//reading all tasks and caculate the priority for sorting
     var priority=priority_relevant_cols[i][0];
     var assigningTime=new Date(priority_relevant_cols[i][1]);
     var requestingTime=new Date(priority_relevant_cols[i][2]);
     calculated_priority.push([]);
    
     var difference= Math.floor((currentTime-assigningTime)/1000/3600/24/10);//every 10 days' delay equals one priority level up
     if(i==0)  calculated_priority[i].push("cal_pri");
     else  if(currentTime-requestingTime>=-1) calculated_priority[i].push(1000000);
     else calculated_priority[i].push(priority+ difference);

  }
  sheet.getRange(1, col_num+1, row_num, 1).setValues(calculated_priority);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, row_num, col_num+1).sort({column:col_num+1,ascending: false});
  sheet.deleteColumn(col_num+1);
}



/**
 * A custom function that converts meters to miles.
 *
 * @param {Number} meters The distance in meters.
 * @return {Number} The distance in miles.
 */
function metersToMiles(meters) {
  if (typeof meters != 'number') {
    return null;
  }
  return meters / 1000 * 0.621371;
}


/**
 * Creates a new sheet containing step-by-step directions between the two
 * addresses on the "Settings" sheet that the user selected.
 */
function generateStepByStep_() {
  var spreadsheet = SpreadsheetApp.getActive();
  var settingsSheet = spreadsheet.getSheetByName('Settings');
  settingsSheet.activate();

  // Prompt the user for a row number.
  var selectedRow = Browser.inputBox('Generate step-by-step',
      'Please enter the row number of the addresses to use' +
      ' (for example, "2"):',
      Browser.Buttons.OK_CANCEL);
  if (selectedRow == 'cancel') {
    return;
  }
  var rowNumber = Number(selectedRow);
  if (isNaN(rowNumber) || rowNumber < 2 ||
      rowNumber > settingsSheet.getLastRow()) {
    Browser.msgBox('Error',
        Utilities.formatString('Row "%s" is not valid.', selectedRow),
        Browser.Buttons.OK);
    return;
  }

  // Retrieve the addresses in that row.
  var row = settingsSheet.getRange(rowNumber, 1, 1, 2);
  var rowValues = row.getValues();
  var origin = rowValues[0][0];
  var destination = rowValues[0][1];
  if (!origin || !destination) {
    Browser.msgBox('Error', 'Row does not contain two addresses.',
        Browser.Buttons.OK);
    return;
  }

  // Get the raw directions information.
  var directions = getDirections_(origin, destination);

  // Create a new sheet and append the steps in the directions.
  var sheetName = 'Driving Directions for Row ' + rowNumber;
  var directionsSheet = spreadsheet.getSheetByName(sheetName);
  if (directionsSheet) {
    directionsSheet.clear();
    directionsSheet.activate();
  } else {
    directionsSheet =
        spreadsheet.insertSheet(sheetName, spreadsheet.getNumSheets());
  }
  var sheetTitle = Utilities.formatString('Driving Directions from %s to %s',
      origin, destination);
  var headers = [
    [sheetTitle, '', ''],
    ['Step', 'Distance (Meters)', 'Distance (Miles)']
  ];
  var newRows = [];
  for (var i = 0; i < directions.routes[0].legs[0].steps.length; i++) {
    var step = directions.routes[0].legs[0].steps[i];
    // Remove HTML tags from the instructions.
    var instructions = step.html_instructions.replace(/<br>|<div.*?>/g, '\n')
        .replace(/<.*?>/g, '');
    newRows.push([
      instructions,
      step.distance.value
    ]);
  }
  directionsSheet.getRange(1, 1, headers.length, 3).setValues(headers);
  directionsSheet.getRange(headers.length + 1, 1, newRows.length, 2)
      .setValues(newRows);
  directionsSheet.getRange(headers.length + 1, 3, newRows.length, 1)
      .setFormulaR1C1('=METERSTOMILES(R[0]C[-1])');

  // Format the new sheet.
  directionsSheet.getRange('A1:C1').merge().setBackground('#ddddee');
  directionsSheet.getRange('A1:2').setFontWeight('bold');
  directionsSheet.setColumnWidth(1, 500);
  directionsSheet.getRange('B2:C').setVerticalAlignment('top');
  directionsSheet.getRange('C2:C').setNumberFormat('0.00');
  var stepsRange = directionsSheet.getDataRange()
      .offset(2, 0, directionsSheet.getLastRow() - 2);
  setAlternatingRowBackgroundColors_(stepsRange, '#ffffff', '#eeeeee');
  directionsSheet.setFrozenRows(2);
  SpreadsheetApp.flush();
}

/**
 * Sets the background colors for alternating rows within the range.
 * @param {Range} range The range to change the background colors of.
 * @param {string} oddColor The color to apply to odd rows (relative to the
 *     start of the range).
 * @param {string} evenColor The color to apply to even rows (relative to the
 *     start of the range).
 */
function setAlternatingRowBackgroundColors_(range, oddColor, evenColor) {
  var backgrounds = [];
  for (var row = 1; row <= range.getNumRows(); row++) {
    var rowBackgrounds = [];
    for (var column = 1; column <= range.getNumColumns(); column++) {
      if (row % 2 == 0) {
        rowBackgrounds.push(evenColor);
      } else {
        rowBackgrounds.push(oddColor);
      }
    }
    backgrounds.push(rowBackgrounds);
  }
  range.setBackgrounds(backgrounds);
}

/**
 * A shared helper function used to obtain the full set of directions
 * information between two addresses. Uses the Apps Script Maps Service.
 *
 * @param {String} origin The starting address.
 * @param {String} destination The ending address.
 * @return {Object} The directions response object.
 */
function getDirections_(origin, destination) {
  var directionFinder = Maps.newDirectionFinder();
  directionFinder.setOrigin(origin);
  directionFinder.setDestination(destination);
  var directions = directionFinder.getDirections();
  if (directions.status !== 'OK') {
    throw directions.error_message;
  }
  return directions;
}
