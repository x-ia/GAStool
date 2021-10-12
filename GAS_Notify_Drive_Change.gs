// #################### Google Apps Script #####################
// #                                                           #
// #   Listing files in Google Drive                           #
// #     and notifying changed via E-mail                      #
// #                                                           #
// #  1st release: 2021-10-06  by Y.Kosaka                     #
// #  Last update: 2021-10-12  by Y.Kosaka                     #
// #  See more information :                                   #
// #    https://riskinlife.com/ja/gas-notify-drive-change      #
// #############################################################

var K = 0;
        // Sheet name for listing files (Please see the bottom bar of Spreadsheet)
const SHEET_NAME_OF_CONTROL = "control";
        // Row number for cell range of GDrive file attributes to collect
const ROW_ATTR    = 4;
        // Offset number of rows for cell range of GDrive IDs to search
const ROW_FOLDERS = 6;
const COL_VAL     = 2;
const STR_PROGRESS = ["Running", "Finished"];
const STR_CHANGE   = ["last changed", "created", "updated", "renamed", "located", "restored", "removed"];
const MENU = {title: "Action", list: [ {name: "Get files list", functionName: "mainFunc"} ]};
const FLAG_REMOVE  = -9;
const FLAG_IGNORE  = -2;
const LINE_BUF     = 100;
const URL_PREFIX   = "https://drive.google.com/file/d/";
const URL_POSTFIX  = "/view";
const CHR_NOSUB    = "/";

        // Open lists file on Spreadsheet
//var SPREADSHEET_4_LISTS = SpreadsheetApp.openById(SHEET_ID_OF_LISTS);
const SPREADSHEET_4_LISTS = SpreadsheetApp.getActiveSpreadsheet();
var SHEET_4_CONTROL = SPREADSHEET_4_LISTS.getSheetByName(SHEET_NAME_OF_CONTROL);
        // ID of Spreadsheet file for listing (Please see the URL)
//var SHEET_ID_OF_LISTS = "1RRCfIJDaktzsGaqZPGal1QKVh0651d61h_9Fa8SHOhU";
var SHEET_ID_OF_LISTS = SHEET_4_CONTROL.getSheetId();
        // The script sends Gmail to this E-mail address if detecting any files changed.
var MAIL_ADDRESS_2_NOTIFY = SHEET_4_CONTROL.getRange(2, COL_VAL).getValues();

        // Variable for incrementing row No.
var J = 0;

        // Function for outputting logs
function debugFunc(lineno, varname, description, code){
  Logger.log("LOGGING BEGIN (log id: " + K + ", line no.: " + lineno + ")");
  Logger.log("line No." + lineno);
  Logger.log(description + ":");
  Logger.log(varname);
  Logger.log("```\n" + code + "\n```");
  Logger.log("LOGGING END (log id: " + K + ", line no.: " + lineno + ")");
  K++;
}

        // Function for adding action menu
function onOpen(){
  SPREADSHEET_4_LISTS.addMenu(MENU.title, MENU.list);
}

        // Function for indicating present process
function putProgressFunc(val, row){
  SHEET_4_CONTROL.getRange(row, COL_VAL + 1).setValue(val);
}

        // Function for getting present time in UNIX time format
function getUnixTime(){
  return Date.now();
}

        // Function for getting present time
function getTime(){
  return new Date().toLocaleTimeString();
}

        // Function for getting elapsed seconds
function getElapsed(timeBegin){
  return ((getUnixTime() - timeBegin) * 0.001).toPrecision(3) + " secs";
}

        // Function for setting nameSearch
function setNameSearch(driveFolderSearch){
  var nameSearch = SHEET_4_CONTROL.getRange(ROW_FOLDERS + J, COL_VAL - 1).getValue();
  if(nameSearch.length < 1){
    nameSearch = driveFolderSearch;
  }
  return "[" + nameSearch + "]";
}

        // Function for getting file IDs in the folder from Google Drive recursively
function getListDriveFileIdFunc(driveFolderSearch, nameSearch){
  var listDriveFileId = [];
  var arrFiles = driveFolderSearch.getFiles();
        // Getting IDs of files
          putProgressFunc("Getting file IDs from GDrive", ROW_FOLDERS + J);
  while(arrFiles.hasNext()){
    listDriveFileId.push(arrFiles.next().getId());
  }

  if( nameSearch.indexOf(CHR_NOSUB) < 0){
        // Getting IDs of folders
    var arrFolderSub = driveFolderSearch.getFolders();
          putProgressFunc("Getting folder IDs from GDrive (" + arrFolderSub.length + " folders)", ROW_FOLDERS + J);
    while(arrFolderSub.hasNext()){
      var folderSub = arrFolderSub.next();
      listDriveFileId = listDriveFileId.concat( getListDriveFileIdFunc(folderSub, nameSearch) );
    }
  }

  return listDriveFileId;
}

        // Function for getting file IDs in the folder recursively
function getMapDriveFile(driveFolderSearch, nameSearch){
  var arrFolder = driveFolderSearch.getFolders();
  var arrFile = driveFolderSearch.getFiles();

  var arrFileId = getListDriveFileIdFunc(driveFolderSearch, nameSearch);
  var mapDriveFiles = {};
        // Iteration for getting attributes of each files
  arrFileId.forEach(
    function( value, i ){
      var fileName  = DriveApp.getFileById(value);
      var timestamp = Drive.Files.get(value).modifiedDate;
      var parentId  = Drive.Files.get(value).getParents()[0].id;
      mapDriveFiles[value] = {
        fileName: fileName,
        timestamp: timestamp,
        parentId : parentId,        
      };
    }
  );
  return mapDriveFiles;
}

        // Function for getting combination of file attributes to collect
function getHeaderFunc(){
  var arrHeader = [[]];
  i = 0;
  while(SHEET_4_CONTROL.getRange(ROW_ATTR, COL_VAL + i).getValue().length > 0){
    arrHeader[0].push(SHEET_4_CONTROL.getRange(ROW_ATTR, COL_VAL + i++).getValue());
  }
  return arrHeader;
}

        // Function for create cheet for listing file
function setSheetFunc(sheetName, arrHeader){
  var sheet = SPREADSHEET_4_LISTS.getSheetByName(sheetName);
  if(!sheet){
    sheet = SPREADSHEET_4_LISTS.insertSheet();
    sheet.setName(sheetName);
    sheet.getRange(1, 1).setValue("fileId");
    sheet.getRange(1, 2).setValue("flag");
    sheet.getRange(1, 3).setValue("fileName");
    sheet.getRange(1, 4).setValue("timestamp");
    sheet.getRange(1, 5).setValue("parentId");
    sheet.getRange(1, 6, 1, arrHeader[0].length).setValues(arrHeader);
  }
  return sheet;
}

        // Function for converting file table to map
function setMapFileListFunc(dataFiles){
          putProgressFunc("converting file table to map (" + dataFiles.length + " files)", ROW_FOLDERS + J);
  var mapFileList = {};
  for(var i = 1; i < dataFiles.length; i++){
    mapFileList[dataFiles[i][0]] = {
      fileId   : dataFiles[i][0],
      flag     : dataFiles[i][1],
      fileName : dataFiles[i][2],
      timestamp: dataFiles[i][3],
      parentId : dataFiles[i][4],
      rowNo    : i + 1
    };
  }
  return mapFileList;
}

        // Function for checking what changed about the file
function checkChangeFunc(previous, present){
  var change = 0;
  if(present.timestamp > previous.timestamp){
    change = 2;
  } else if(present.fileName != previous.fileName){
    change = 3;
  } else if(present.parentId != previous.parentId){
    change = 4;
  } else if(previous.flag == FLAG_REMOVE){
    change = 5;
  }
  return change;
}

        // Definition for generating function for getting custom file attribute
function getFileAttrFuncSub(file, attr){
  let genFuncAttr = new Function('file', 'return file.' + attr);
  try{
    return genFuncAttr(file);
  }
  catch(e){
          debugFunc(231, e, "Error", arguments.callee);
    return "-";
  }
}

        // Sub function for getting file attributes
function getFileAttrFunc(mapDriveFile, key, flag, arrHeader){
  var file = Drive.Files.get(key);
  var arrAttr = [];
  arrAttr.push(key);
  arrAttr.push(flag);
  arrAttr.push(mapDriveFile.fileName);
  arrAttr.push(mapDriveFile.timestamp);
  arrAttr.push(mapDriveFile.parentId);
  for(var i = 0; i < arrHeader[0].length; i++){
    arrAttr.push( getFileAttrFuncSub(file, arrHeader[0][i]) );
  }
  return arrAttr;
}

        // Sub function for filling files list
function putFileAttrFunc(arrAttr, sheetFiles, row){
  sheetFiles.getRange(row, 1, arrAttr.length, arrAttr[0].length).setValues(arrAttr);
}

        // Sub function for filling removed flag
function putFlagRemovedFunc(sheetFiles, row){
  sheetFiles.getRange(row, 2).setValue(FLAG_REMOVE);
}

        // Function for filling files list and getting changed entries
function putMapDriveFileFunc(mapDriveFiles, mapFileList, sheetFiles, arrHeader){
          putProgressFunc("Checking file changed (" + Object.keys(mapDriveFiles).length + " files)", ROW_FOLDERS + J);
  var mapFolderChanged = [];
  var arrAttrExist     = [];
  var arrAttrNew       = [];
  var rowNo            = 0;

  for(key in mapDriveFiles){
    var flag   = 0;
    var change = 0;
    if(key in mapFileList){
        // When already existing in Spreadsheet
      if(mapFileList[key].flag != FLAG_IGNORE){
        change = checkChangeFunc(mapFileList[key], mapDriveFiles[key]);
        if(change > 0){
        // When the file has changed
          flag  = (mapFileList[key].flag < 0) ? 0 : mapFileList[key].flag + 1;
          rowNo = mapFileList[key].rowNo;
          arrAttrExist = [getFileAttrFunc(mapDriveFiles[key], key, flag, arrHeader)];
          putFileAttrFunc(arrAttrExist, sheetFiles, rowNo);
        }
      }
    } else {
        // When not existing in Spreadsheet yet
      if( SHEET_ID_OF_LISTS == mapDriveFiles[key].fileId ){
        // set flag for ignoring this list file changed by this script
        flag = FLAG_IGNORE;
      }
      var change = 1;
      arrAttrNew.push(getFileAttrFunc(mapDriveFiles[key], key, flag, arrHeader));
      if(arrAttrNew.length >= LINE_BUF){
          putProgressFunc("Appending new files (" + LINE_BUF + " files)", ROW_FOLDERS + J);
        rowNo = sheetFiles.getLastRow() + 1;
        putFileAttrFunc(arrAttrNew, sheetFiles, rowNo);
        arrAttrNew = [];
      }
    }

    if(change > 0 && flag >= 0){
      fileName  = mapDriveFiles[key].fileName;
      timestamp = mapDriveFiles[key].timestamp;
      parentId  = mapDriveFiles[key].parentId;
      mapFolderChanged.push({
        fileId   : key,
        fileName : fileName,
        timestamp: timestamp,
        parentId : parentId,
        change   : change
      });
    }
  }
  if(arrAttrNew.length){
          putProgressFunc("Appending new files (" + arrAttrNew.length + " files)", ROW_FOLDERS + J);
    rowNo = sheetFiles.getLastRow() + 1;
    putFileAttrFunc(arrAttrNew, sheetFiles, rowNo);
  }

          putProgressFunc("Checking file removed (" + Object.keys(mapFileList).length + " files)", ROW_FOLDERS + J);
  for(key in mapFileList){
    var change = 0;
    if(mapFileList[key].flag > FLAG_REMOVE){
      if (key in mapDriveFiles){
      } else {
        fileName  = mapFileList[key].fileName;
        parentId  = mapFileList[key].parentId,
        change    = 6;
        rowNo     = mapFileList[key].rowNo;
      }
      if(change > 0){
        putFlagRemovedFunc(sheetFiles, rowNo);
        mapFolderChanged.push({
          fileId   : key,
          fileName : fileName,
          parentId : parentId,
          change   : change
        });
      }
    }
  }
  return mapFolderChanged;
}

        // Function for filling files list and getting changed entries
function genFolderChangedFunc(driveFolderSearch, nameSearch){
  var mapDriveFiles = getMapDriveFile(driveFolderSearch, nameSearch);
        // Getting combination os file attributes to collect
  var arrHeader = getHeaderFunc();
        // Getting list of files from Spreadsheet
  var sheetFiles = setSheetFunc(nameSearch, arrHeader);
  var dataFiles = sheetFiles.getDataRange().getValues();
        // Convert table to map
  var mapFileList = setMapFileListFunc(dataFiles);

        // Filling files list and getting changed entries
  var mapFolderChanged = putMapDriveFileFunc(mapDriveFiles, mapFileList, sheetFiles, arrHeader);
  return mapFolderChanged;
}

        // Function for constructing a message for mail
function setMsgNotifyFunc(mapFolderChanged){
          putProgressFunc("aggregating files changed (" + mapFolderChanged.length + " files)", ROW_FOLDERS + J);
  var msgNotify = "";
  for( key in mapFolderChanged ){
    item = mapFolderChanged[key];
    msgNotify +=
      item.fileName
      + " \t" + STR_CHANGE[item.change];
    if(item.change < 4){
      msgNotify += " \tin  ";
    } else if(item.change < 6){
      msgNotify += " \tinto  ";
    } else {
      msgNotify += " \tfrom  ";
    }

    try{
      msgNotify += DriveApp.getFileById(item.parentId);
    }
    catch(e){
          debugFunc(409, e, "Error", arguments.callee);
      msgNotify += "(missing folder)";
    }

    if(item.change < 3){
      msgNotify += 
        "\n \tby  " + Drive.Files.get(item.fileId).lastModifyingUserName
        + " \ton  " + item.timestamp.toLocaleString();
    }
    msgNotify += 
      "\n" + URL_PREFIX + item.fileId + URL_POSTFIX + "\n\n"
  }
  return msgNotify;
}

        // Function for sending Gmail
function sendNotifyMailFunc(msgNotify, nameSearch){
          putProgressFunc("Sending Notify mail", ROW_FOLDERS + J);
  GmailApp.sendEmail(
    MAIL_ADDRESS_2_NOTIFY,
    nameSearch + "GDrive Changed",
    msgNotify
  );
}

        // Main routine of this script
function mainFunc(){
        // Clear previous run
  if(SHEET_4_CONTROL.getRange(ROW_FOLDERS, COL_VAL + 1).getValue() == STR_PROGRESS[1]){
    while(!SHEET_4_CONTROL.getRange(ROW_FOLDERS + J, COL_VAL + 1).isBlank()){
      SHEET_4_CONTROL.getRange(ROW_FOLDERS + J++, COL_VAL + 1).clearContent();
    }
    return 0;
  }
          putProgressFunc(STR_PROGRESS[0], ROW_FOLDERS);

        // Iteration for number of folders to search
  while(SHEET_4_CONTROL.getRange(ROW_FOLDERS + ++J, COL_VAL).getValue().length > 9){
          var timeBegin = getUnixTime();
        // Skip searching succeeded on previous run
    if(!SHEET_4_CONTROL.getRange(ROW_FOLDERS + J + 1, COL_VAL + 1).isBlank()){
      continue;
    }
          putProgressFunc("Begin (" + getTime() + ")", ROW_FOLDERS + J);

        // ID of Google Drive folder to search (Please see the URL)
    var folderIdSearch = (SHEET_4_CONTROL.getRange(ROW_FOLDERS + J, COL_VAL).getValue()).replace(/(.*\/)?([\w\d_-]+)(\?.*)?/, '$2');
    var driveFolderSearch = DriveApp.getFolderById(folderIdSearch);
        // Sheet name for listing files (Please see the bottom bar of Spreadsheet)
    var nameSearch = setNameSearch(driveFolderSearch);
        // Fill files list and get changed entries
    mapFolderChanged = genFolderChangedFunc(driveFolderSearch, nameSearch);

        // Send Gmail
    if(mapFolderChanged.length > 0){
        // construct a message for mail
      var msgNotify = setMsgNotifyFunc(mapFolderChanged);
      sendNotifyMailFunc(msgNotify, nameSearch);
          putProgressFunc("End with notifying (" + getElapsed(timeBegin) + ")", ROW_FOLDERS + J);
    } else {
          putProgressFunc("End without notifying (" + getElapsed(timeBegin) + ")", ROW_FOLDERS + J);
    }

  }
          putProgressFunc(STR_PROGRESS[1], ROW_FOLDERS);
}
