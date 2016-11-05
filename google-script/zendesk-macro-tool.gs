/**
*  Tool to pull Zendesk Macros and push the update back to Zendesk.
*
*  @author Robin Müller
*/

/*#####################
*   CONFIGURATION     #
#####################*/
var zendeskUrl = "https://companyname.zendesk.com/api/v2/";
var zendeskUsername = "your-account@company.com/token"; // your username with /token at the end
var zendeskPassword = ""; // your token


var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = SpreadsheetApp.getActiveSheet();

function onInstall() {
  onOpen();
}

function onOpen() {
  var menuEntries =
      [
        {
          name:"Import Macros (OVERWRITES CONTENT!)",
          functionName: "runImportMacros"
        },
        {
          name:"Sync changes to Zendesk",
          functionName: "runUpdateZendesk"
        }
      ];

  ss.addMenu("Zendesk", menuEntries);
  _showMessageBox();
}

function _showMessageBox() {
  Browser.msgBox('Please read the instruction in the "How to use the sheet" section carefully before using this tool.');
}

function runImportMacros() {
  var data = [];
  var page = 1;
  var totalPages = 1;

  do {
    var url = zendeskUrl + "macros.json?page=" + page;
    var macros = _fetchData(url);

    for(var i = 0; i < macros['macros'].length; i++) {
      var cur = _extractMacroData(macros['macros'][i]);
      data.push(cur);
    }

    totalPages = Math.ceil(macros['count']/100);
    page++;
  } while(page <= totalPages);

  // update sheet
  _updateSheet(
      data,
      _getNumberOfRows(data),
      _getNumberOfColumns(data)
  );

}


function runUpdateZendesk() {
  var sheetData = SpreadsheetApp.getActiveSheet().getDataRange().getValues();

  for(var x=0; x < sheetData.length; x++) {
    if(sheetData[x][0] == "edited") {
      // pull macro data from zendesk again
      var macroUrl = zendeskUrl + "macros/" + sheetData[x][1] + ".json";
      var macroData = _fetchData(macroUrl);

      var newData = _mergeSyncData(macroData, sheetData[x]);

      // sync back to zendesk
      var updateSuccessful = _updateData(macroUrl, newData);

      // mark as synced back
      if(updateSuccessful) {
        _markMacroAsSynced(x+1);
      }
    }
  }
}


function _mergeSyncData(macroData, sheetData) {
  var newData =  {
    'macro': {
      'actions': []
    }
  };

  // copy over actions from Zendesk - the complete actions object needs to be synced back, otherwise data will be deleted on Zendesk side
  newData['macro']['actions'] = macroData['macro']['actions'];

  for(var z = 0; z < newData['macro']['actions'].length; z++) {

    // subject
    if(newData['macro']['actions'][z]['field'] == "subject") {
      newData['macro']['actions'][z]['value'] = sheetData[4];
    }

    // comment
    if(newData['macro']['actions'][z]['field'] == "comment_value") {
      if(newData['macro']['actions'][z]['value'][1]) {
        newData['macro']['actions'][z]['value'][1] = sheetData[5];
      }
    }

    // @TODO: Sync tags back

  }

  return newData;
}


function _updateData(url, data) {
  var options = {};
  options.headers = {
    "Authorization": "Basic " + Utilities.base64Encode(zendeskUsername + ":" + zendeskPassword),
    "Content-Type": "application/json"
  };
  options.method = 'put';
  options.muteHttpExceptions = false;
  options.payload = JSON.stringify(data);

  try {
    var response = UrlFetchApp.fetch(url, options);

    // handle response
    if(response.getResponseCode() == 200) {
      return true;
    }
  } catch (err) {}

  return false;
}


function _markMacroAsSynced(row) {
  SpreadsheetApp.getActiveSheet().getRange('A'+row).setValue('synced-to-zendesk');
}


function _fetchData(url) {
  var options = {};
  options.headers = {"Authorization": "Basic " + Utilities.base64Encode(zendeskUsername + ":" + zendeskPassword)};

  var response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}

function _extractMacroData(curMacro) {
  var extracted = [];

  extracted[0] = "new";
  extracted[1] = curMacro['id'];
  extracted[2] = curMacro['active'] ? "yes" : "no";
  extracted[3] = curMacro['title'];

  extracted[4] = "";
  extracted[5] = "";
  extracted[6] = "";

  for(var x = 0; x < curMacro['actions'].length; x++) {
    // subject
    if(curMacro['actions'][x]['field'] == "subject") {
      extracted[4] = curMacro['actions'][x]['value'];
    }

    // comment
    if(curMacro['actions'][x]['field'] == "comment_value") {
      if(curMacro['actions'][x]['value'][1]) {
        extracted[5] = curMacro['actions'][x]['value'][1];
      }
    }

    // tags
    if(curMacro['actions'][x]['field'] == "current_tags") {
      extracted[6] = curMacro['actions'][x]['value'];
    }
  }

  return extracted;
}


function _updateSheet(data, numberRows, numberColumns) {
  sheet.getRange(2, 1, numberRows, numberColumns).setValues(data);
}


function _getNumberOfRows(data) {
  return data.length;
}


function _getNumberOfColumns(data) {
  var columns = 0;

  for(var x = 0; x < data.length; x++) {
    if(data[x].length > columns) {
      columns = data[x].length;
    }
  }

  return columns;
}
