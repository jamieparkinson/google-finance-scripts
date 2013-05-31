function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Generate Full Data",
    functionName : "putAllPays"
  }];
  sheet.addMenu("Scripts", entries);
};

function mkXPays(spread,sheet,endDate) {
  //get exclusive payments and objectify
  var Xin = spread.getRangeByName("Xin");
  var Xout = spread.getRangeByName("Xout");
  
  //Make friendly
  var payObjs = getRowsData(sheet,Xin).concat(getRowsData(sheet,Xout));
  var XPayments = new Array();
  for (var p in payObjs) {
    if (payObjs[p].date.getTime() <= endDate) {
      var amount = payObjs[p].amount;
      XPayments.push([payObjs[p].date.getTime(),amount,amount,amount]);
    }
  }
  
  return XPayments;
}

function doDiscrete(thisDate,endDate,cost,bigArray,freq,limU,limL) {
  while (thisDate <= endDate) {
    bigArray.push([thisDate,cost,limU,limL]);
    thisDate += freq;
  }
  return bigArray;
}

function mkRegPays(spread,sheet,startDate,endDate) { 
  //get regular incoming and outgoing payment ranges
  var Rin = spread.getRangeByName("Rin");
  var Rout = spread.getRangeByName("Rout");
  
  //make payments into objects and create blank total payments object
  var payObjs = getRowsData(sheet,Rin).concat(getRowsData(sheet,Rout));
  var regPayments = new Array();
  
  //for each regular payment:
  for (var p in payObjs) { 
    //convert days to milliseconds
    payObjs[p].frequency = payObjs[p].frequencyDays * 86400000;
    payObjs[p].offset = payObjs[p].beginFrom.getTime() - startDate;
    
    //Create discrete payments up to endDate
    var thisDate = startDate + payObjs[p].offset;
    var amount = payObjs[p].amount;
    regPayments = doDiscrete(thisDate,endDate,amount,regPayments,payObjs[p].frequency,amount,amount);
  }
  return regPayments;
}

function mkLivePays(spread,sheet,startDate,endDate) {
  //get living costs and uncertainties
  var liveWeek = spread.getRangeByName("liveWeek").getValue() * (-1);
  var liveDelta = spread.getRangeByName("liveDelta").getValue();
  
  //make discrete payments
  var livePayments = new Array();
  livePayments = doDiscrete(startDate,endDate,liveWeek,livePayments,604800000,liveWeek-liveDelta,liveWeek+liveDelta);
  
  return livePayments;
}

//for sorting 2D array
function sortFn(a,b) {
  a = a[0]
  b = b[0]
  return a == b ? 0 : (a < b ? -1 : 1);
}

function putAllPays() {
  //open sheet and make arguments
  var spread = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var startDate = spread.getRangeByName("startDate").getValue().getTime();
  var endDate = spread.getRangeByName("endDate").getValue().getTime();
  
  //get regular, exclusive and living costs then concatenate
  var regPayments = mkRegPays(spread,sheet,startDate,endDate);
  var XPayments = mkXPays(spread,sheet,endDate);
  var livePayments = mkLivePays(spread,sheet,startDate,endDate);
  var allPayments = regPayments.concat(XPayments, livePayments);
 
  //sort payments by date
  allPayments = allPayments.sort(sortFn);
  
  //convert back to date objects
  allPayments = allPayments.map(function(obj) { return [new Date(obj[0]),obj[1],'',obj[2],'',obj[3]]; });
  Logger.log(allPayments);
  
  //put back into the sheet!
  var payRange = sheet.getRange(41,2,allPayments.length,6);
  payRange.setValues(allPayments);
}

//Google's functions for parsing etc
      function getRowsData(sheet, range, columnHeadersRowIndex) {
        columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
        var numColumns = range.getLastColumn() - range.getColumn() + 1;
        var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
        var headers = headersRange.getValues()[0];
        return getObjects(range.getValues(), normalizeHeaders(headers));
      }
      
      
      function getObjects(data, keys) {
        var objects = [];
        for (var i = 0; i < data.length; ++i) {
          var object = {};
          var hasData = false;
          for (var j = 0; j < data[i].length; ++j) {
            var cellData = data[i][j];
            if (isCellEmpty(cellData)) {
              continue;
            }
            object[keys[j]] = cellData;
            hasData = true;
          }
          if (hasData) {
            objects.push(object);
          }
        }
        return objects;
      }
      
      
      function normalizeHeaders(headers) {
        var keys = [];
        for (var i = 0; i < headers.length; ++i) {
          var key = normalizeHeader(headers[i]);
          if (key.length > 0) {
            keys.push(key);
          }
        }
        return keys;
      }
      
      
      function normalizeHeader(header) {
        var key = "";
        var upperCase = false;
        for (var i = 0; i < header.length; ++i) {
          var letter = header[i];
          if (letter == " " && key.length > 0) {
            upperCase = true;
            continue;
          }
          if (!isAlnum(letter)) {
            continue;
          }
          if (key.length == 0 && isDigit(letter)) {
            continue; // first character must be a letter
          }
          if (upperCase) {
            upperCase = false;
            key += letter.toUpperCase();
          } else {
            key += letter.toLowerCase();
          }
        }
        return key;
      }
      
      function isCellEmpty(cellData) {
        return typeof(cellData) == "string" && cellData == "";
      }
      
      function isAlnum(char) {
        return char >= 'A' && char <= 'Z' ||
          char >= 'a' && char <= 'z' ||
          isDigit(char);
      }
      
      function isDigit(char) {
        return char >= '0' && char <= '9';
      }
      
      function arrayTranspose(data) {
        if (data.length == 0 || data[0].length == 0) {
          return null;
        }
      
        var ret = [];
        for (var i = 0; i < data[0].length; ++i) {
          ret.push([]);
        }
      
        for (var i = 0; i < data.length; ++i) {
          for (var j = 0; j < data[i].length; ++j) {
            ret[j][i] = data[i][j];
          }
        }
      
        return ret;
      }
