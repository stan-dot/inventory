
/** TODO
 * simplification of the SheetObject constructor method into accepting one object as parameter
 * then creation of new such objects is more DRY, without repeating the constants
 * 
 */

/** NOTES
 * the comments in the form [sheetobject].output.setValue are for debugging purposes
 * useful if you want to tweak the code for your needs
 * 
 */

const JOURNEY_COLUMN = 3;
const ACTION_COLUMN = 13;
const CUSTOM_ROW = 1;
const SORT_ROW = 2;

const STARTING_ROW = 4;

const HEADERS_RANGE = 'B3:R3';
const OUTPUT_STRING = 'H1';
const DEFAULT_RANGE = 'B4:R600';
const DEFAULT_RANGE_STRING = 'B2:L2';
const DEFAULT_RANGE_UNBOUNDED = 'B4:R';

const LAST_INDEX = 'A1';
const COLUMN_AS_TYPE = 'K1';
const WHICH_ITEM_ON_COLUMN = 'M1';

// JOURNEY
const JOURNEY_COLUMN_IN_LIST = 5;
const LABELS_ROW = 2;  
const LABELS_STARTING_COLUMN = 2;
const LABELS_DISTANCE_BETWEEN_COLUMNS = 2;
  

function SheetObject(name, rangeString, sortRangeString){
  var sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (rangeString == 0){
   rangeString =  getDefaultRange(sheet);
  }
  var result = {
    sheet:  sheet,
    range: sheet.getRange(rangeString),
    sortingRange: sheet.getRange(sortRangeString),
    output: sheet.getRange(OUTPUT_STRING),
    name: name
  }
  return result;
}


function getDefaultRange(sheet){
  var lastIndex = sheet.getRange(LAST_INDEX).getValue().toString();
  return DEFAULT_RANGE_UNBOUNDED + lastIndex + STARTING_ROW;
}

  
// these cannot work as global variables (in apps script as oppposed to JS), so are here just for reference
var listSheetObject = new SheetObject('list', 0, DEFAULT_RANGE_STRING);
var customSheetObject = new SheetObject('custom', 0, DEFAULT_RANGE_STRING);
var journeySheetObject = new SheetObject('journey', 'K7:L600', 'K4');


//general function to redirect
function onEdit(event){
  var sheet = event.source.getActiveSheet();
  var sheetName = sheet.getName();
  var editedCell = sheet.getActiveCell();
  var listSheetObject = new SheetObject('list', 0, DEFAULT_RANGE_STRING);
  var customSheetObject = new SheetObject('custom', 0, DEFAULT_RANGE_STRING);
  var journeySheetObject = new SheetObject('journey', 'K7:L600', 'K4');
  listSheetObject.output.setValue("onedit: " + sheetName + ": vs " + listSheetObject.name);
  switch (sheetName){
    case listSheetObject.name:
      listSheetObject.output.setValue("sending to list functions");
      listFunctions(editedCell, listSheetObject);    
      break;
    case customSheetObject.name:
      if(editedCell.getRow() == CUSTOM_ROW){
        listSheetObject.output.setValue("sending to custom");
        moveToCustom();
      }else{
        listSheetObject.output.setValue("sorting the custom one");
        listFunctions(editedCell, customSheetObject);
      }
      break;
    case journeySheetObject.name:
      listSheetObject.output.setValue("sending to journey");
      moveToJourney();
      break;
      
    default:
      
  }
}


//functions concerning the main list 
function listFunctions(editedCell, sheetObject){
  var value = editedCell.getValue();
  var column = editedCell.getColumn();
  var row = editedCell.getRow();
  sheetObject.output.setValue("listing");
  if ( row == SORT_ROW ){ //sorting by custom column order - so two or more columns possible 
    sortListAdvanced(sheetObject);
  }else if( value == 'del' && column == ACTION_COLUMN ){
    // deleteRow(row, sheetObject);
    sheetObject.sheet.deleteRow(row);
  }else if( column == JOURNEY_COLUMN_IN_LIST ){
    sheetObject.output.setValue("moving to journey");
    moveToJourney();
  }
}


function deleteRow(row, sheetObject){
  var resetRange = sheetObject.sheet.getRange("A"+row+":R" + row);
  resetRange.clearContent();
  sortList(sheetObject.sheet);
}


function getSortDirections(sheetObject){
  sheetObject.output.setValue("getting sort directions");
  var values = sheetObject.sortingRange.getValues();
  sheetObject.output.setValue(values);
  const COLUMN_OFFSET = 2;
  var columnPriorityPairs = [];
  var justOneRow = values[0];
  for(var i = 0; i< justOneRow.length; i++){
    if(justOneRow[i] != 0){
      columnPriorityPairs.push({column: COLUMN_OFFSET + i, prio: justOneRow[i]});
    }
  }
  return columnPriorityPairs;  
}


//it's set for descending now, to reverse change (1 into -1) and vice versa
function comparePairs(a, b) {
  if (a.prio < b.prio) {
    return 1;
    //return -1;
  }
  if (a.prio > b.prio) {
    return -1;
    // return 1;
  }
  return 0;
}


function sortListAdvanced(sheetObject){
  var columnPriorityPairs = getSortDirections(sheetObject);
  sheetObject.output.setValue("starting sorting");
  sheetObject.output.setValue(typeof columnPriorityPairs);
  columnPriorityPairs.sort(comparePairs);
  var sortArray = []; 
  for(var i = 0; i < columnPriorityPairs.length; i++){
    sortArray.push({column: columnPriorityPairs[i].column, ascending: false});
  }
  sheetObject.output.setValue(sortArray.length);
  sheetObject.range.sort(sortArray);
}


function getColumnNumber(string){
  var listSheetObject = new SheetObject('list', 0, DEFAULT_RANGE_STRING);
  var values = listSheetObject.sheet.getRange(HEADERS_RANGE).getValues();
  var result = '-1';
  var counter = 0;
  for (var i = 0; i < values[0].length;i++){
    counter++;
    listSheetObject.output.setValue("that's the column value: " +  values[0][i] + ";that's values" + values + "string: " + string);
    if (values[0][i] == string){ //first non empty that matches
      result = i; 
      //listSheetObject.output.setValue("i value: " + i);
    }
  }
  listSheetObject.output.setValue("that's the column index: " +  result +";that's values" + values + "string: " + string + " c: " + counter);
  return result;
}


//functions for displaying in different sheets
function moveToCustom(){
  var listSheetObject = new SheetObject('list', 0, DEFAULT_RANGE_STRING);
  var customSheetObject = new SheetObject('custom', 0, DEFAULT_RANGE_STRING);
  //get column and input wanted
  const DESIRED_COLUMN_STRING = customSheetObject.sheet.getRange(COLUMN_AS_TYPE).getValue();
  const DESIRED_COLUMN_VALUE = customSheetObject.sheet.getRange(WHICH_ITEM_ON_COLUMN).getValue();
  const DESIRED_COLUMN_INDEX = getColumnNumber(DESIRED_COLUMN_STRING);
  // listSheetObject.output.setValue("actual column number: " + DESIRED_COLUMN_INDEX + " string: " + DESIRED_COLUMN_STRING);
  var data = listSheetObject.range.getValues();
  var newData = [];
  for (var i in data) {
    //listSheetObject.output.setValue("actual vs expected: " + data[i] + ":" + data[i][DESIRED_COLUMN_INDEX] + DESIRED_COLUMN_VALUE);
    if (data[i][DESIRED_COLUMN_INDEX] == DESIRED_COLUMN_VALUE) {
      listSheetObject.output.setValue("data: " + data[i] + "i values: " + [data[i][0], data[i][1]]);
      newData.push(data[i]);
    }
  }
  //copy to the new place
  customSheetObject.range.clearContent();
  if(newData[0] != undefined){
    customSheetObject.sheet.getRange(4, 2, newData.length, newData[0].length).setValues(newData);
    customSheetObject.output.setValue("success!");
  }else{
   customSheetObject.output.setValue("error in getting data"); 
  }
}


function getValueOfRangeString(sheetObject, string){
  return sheetObject.sheet.getRange(string).getValue();
}


function getValueOfRangeNums(sheetObject, row, column){
  return sheetObject.sheet.getRange(row, column).getValue();
}


function getColumnForItems(columnIndex){
 return columnIndex - 1; 
}

function labelsToString(labels){
  var string = "";
  for(var i = 0; i < labels.length; i++){
    string += labels[i].value + " " + labels[i].row + " " + labels[i].column + " : ";
  }
  return string;
}

function getLabels(sheetObject){
  const LABEL_LENGTH = 1;
  var labels = [];
  var currentColumn = LABELS_STARTING_COLUMN;
  var cellValue = getValueOfRangeNums(sheetObject, LABELS_ROW, currentColumn);
  var rowForItems = LABELS_ROW + 2;
  while(cellValue.length == LABEL_LENGTH){
    labels.push({
      value: cellValue,
      row: rowForItems, 
      column: getColumnForItems(currentColumn)
    });
    currentColumn += LABELS_DISTANCE_BETWEEN_COLUMNS;
    cellValue = getValueOfRangeNums(sheetObject, LABELS_ROW, currentColumn);
  }
  // sheetObject.output.setValue(labelsToString(labels));
  return labels;
}


// go through the items and putting into 3d array; but not equal sizes
// output the stuff into the correct rows; nope, do it all at one time, get and move
function moveToJourney(){
  var listSheetObject = new SheetObject('list', 0, DEFAULT_RANGE_STRING);
  var journeySheetObject = new SheetObject('journey', 'K7:L600', 'K4');
  journeySheetObject.sheet.getRange('A4:J200').clearContent();
  var data = listSheetObject.range.getValues();
  var labels = getLabels(journeySheetObject);
  const DATA_WIDTH = 2;
  var counter = new Array(labels.length).fill(0);
  for (var dataRowIndex in data) {
    var value = data[dataRowIndex][JOURNEY_COLUMN];
    for(var labelIndex = 0; labelIndex < labels.length; labelIndex++){
      if (value == labels[labelIndex].value) {
        var finalRow = labels[labelIndex].row + counter[labelIndex];
        // journeySheetObject.output.setValue("final row: "+ finalRow + "labels: " +  labels[1].value);    
        journeySheetObject.sheet.getRange(finalRow, labels[labelIndex].column, 1, DATA_WIDTH).setValues([[data[dataRowIndex][0], data[dataRowIndex][1]]]); 
        // copying just name and mass
        // journeySheetObject.output.setValue("i values: " + [data[dataRowIndex][0], data[dataRowIndex][1]]);
        counter[labelIndex]++;
      }
    // journeySheetObject.output.setValue("actual vs expected: " + counter + ":" + data[i] + ":" + data[i][JOURNEY_COLUMN] + " : " + JOURNEY_VALUE + "::" + JOURNEY_COLUMN);
    }
  }
  // journeySheetObject.output.setValue("succeeded custom moving");
}


//functions invoked by macros - not automatic

/**
 * Removes duplicate rows from the list table range
 */
function removeDuplicates(){
  var listSheetObject = new SheetObject('list', 0, DEFAULT_RANGE_STRING);
  var data = listSheetObject.range.getValues();
  var newData = [];
  for (var i in data) {
    var row = data[i];
    var duplicate = false;
    for (var j in newData) {
      if (row[0] == newData[j][0]) {
        duplicate = true;
      }
    }
    if (!duplicate) {
      newData.push(row);
    }
  }
  listSheetObject.range.clearContent(); //that is not a function error
  listSheetObject.sheet.getRange(4, 2, newData.length, newData[0].length).setValues(newData);
}
