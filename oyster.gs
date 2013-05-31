function putData() {
  var spread = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spread.getSheetByName("Oyster");
  
  //make sure we don't duplicate data
  var lastUp = spread.getRangeByName("msgID").getValue();
  var oysterJourneys = grabOysterJourneys(lastUp);
  if (!oysterJourneys) return false;
  spread.getRangeByName("msgID").setValue(oysterJourneys.msgID);
  
  var journeyObjs = arr2Obj(oysterJourneys.data);
  
  //make objects into a flat table
  var journeyTable = new Array();
  for (var i=0; i<journeyObjs.length; i++) {
    var jO = journeyObjs[i];
    journeyTable[i] = [jO.date, jO.journeyTime, jO.journey[0], jO.journey[1], jO.cost, jO.balance];
  }
  
  //put into sheet
  if (sheet.getLastRow() > 5) { sheet.insertRowsBefore(5,journeyObjs.length); }
  var dataRange = sheet.getRange(5,2,journeyObjs.length,6);
  dataRange.setValues(journeyTable);
}


function grabOysterJourneys(lastUp) {
  var threads = GmailApp.search("label:Oyster");
  
  var msgID = threads[0].getMessages()[0].getId();
  if (msgID == lastUp) return false;
  
  var att = threads[0].getMessages()[0].getAttachments()[0];
  var rawCSV = att.getDataAsString();
 
  //slice removes mystery blank 1st element and TfL's column headings
  var dataArray = Utilities.parseCsv(rawCSV,',').slice(2);
  return {data:dataArray, msgID:msgID};
}

function arr2Obj(journeys) {
  var objs = new Array();
  keys = ["dateRaw", "startTime", "endTime", "journeyRaw", "less", "more", "balance", "comment"]
  
  //loop through journeys
  for (var i=0; i<journeys.length; i++) {
    var journey = journeys[i];
    var obj = new Object();
  
    //loop through journey details
    for (var j=0; j<journey.length; j++) {
      obj[keys[j]] = journey[j];
    }
    
    //make date/times nicer
    obj.dateRaw = obj.dateRaw.replace(/-/g,' ');
    obj.date = new Date(obj.dateRaw + " " + obj.startTime);
    obj.endDate = (obj.endTime) ? (new Date(obj.dateRaw + " " + obj.endTime)) : '';
    obj.journeyTime = (obj.endDate) ? (obj.endDate.getTime() - obj.date.getTime())/60000 : '';
    
    //make journey nicer
    obj.journey = obj.journeyRaw.split(" to ");
    obj.journey[1] = (obj.journey[1] || '');
    
    //make costs simpler
    obj.cost = obj.more - obj.less;
    
    //remove leftover properties
    delete obj.dateRaw, obj.startTime, obj.endTime, obj.journeyRaw, obj.less, obj.more;
    
    objs.push(obj);
  }
  
  return objs;
}

function test() {
  var str = "Bus journey, route 29";
  Logger.log(str.split(" to ").length);
}
