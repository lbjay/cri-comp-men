/*
* Populate lineup work sheet with available rowers from teamsnap
* 
* Lineup WS needs to have a consistent date value cell
* Fetch list of roster members
* Use date value to fetch practice event
* Use event id to fetch list of availability entries
* Cross ref roster with availability data
* Populate range of cells with available rowers
*/

var SA = SpreadsheetApp;
var moment = Moment.load();
var _ = Underscore.load();

// Add one line to use BetterLog
Logger = BetterLog.useSpreadsheet('1s--UrSqClpcFQ_OWx_fGI7XE62J8HgSRW_zn_1F1TDY'); 

//Now you can log and it will also log to the spreadsheet
Logger.log("That's all you need to do");

var scriptProperties = PropertiesService.getScriptProperties();
var teamsnapApiBase = scriptProperties.getProperty('teamsnap_api_base');
var teamId = scriptProperties.getProperty('team_id');
var accessToken = scriptProperties.getProperty('access_token');
var fetchHeaders = {
  'Authorization': 'Bearer ' + accessToken,
  'Content-Type': 'application/json'
};


function apiFetch(path, params) {
  
  Logger.log("fetching from endpoint: " + path);
  
  var url = teamsnapApiBase + path;
  var queryString = Object.keys(params).map(function(key) {
    return key + '=' + params[key]
  }).join('&');
  Logger.log("full url: " + url + '?' + queryString);
    
  var resp = UrlFetchApp.fetch(url + '?' + queryString, { 'headers': fetchHeaders });
  var content = resp.getContentText();
  return JSON.parse(content);
}

function onOpen(e) {
  var ss = SA.getActive();
  ss.getSheetId()
  var criLineupMenu = [
    {name: 'Import Roster', functionName: 'importRoster'},
    {name: 'New Lineup...', functionName: 'createNewLineup'}
  ];
  Logger.info("Adding CRI Lineup menu options...");
  ss.addMenu('CRI Teamsnap', criLineupMenu);
}

function formulaValue() {
  var ui = SA.getUi();
  var sheet = SA.getActive().getSheetByName("Lineup Worksheet");
  var formula = sheet.getRange("E8").getFormula();
  ui.alert(formula);
}

function importRoster() {
  var ss = SA.getActive();
  var sheet = ss.getSheetByName("Roster");
  var members = fetchRoster();
  var targetRange = findSheetNamedRange(sheet, "Roster");
  targetRange.clearContent();
  var fillRange = sheet.getRange(targetRange.getRow(), targetRange.getColumn(), members.length, 7);
  var values = _.map(members, function(m) { 
    return [m.sheet_name, m.age(), m.rowing_age(), m.Side, m.email, m.phone, m.is_coxswain()]; 
  });
  fillRange.setValues(values);
  Logger.log("fill Roster starting at %s", fillRange.getA1Notation());
  findSheetNamedRange(sheet, "LastImported").setValue(moment().format());
  return;
}

function createNewLineup() {

  var lineupSheetName = getLineupSheetName();
  if (!lineupSheetName) { Logger.info("lineup creation cancelled"); return ; }
  
  Logger.info('Creating new lineup: ' + lineupSheetName);
  
  var ss = SA.getActive();
  var template = ss.getSheetByName("Lineup Worksheet");
  var newIndex = template.getIndex() + 1;
  var lineupSheet = ss.insertSheet(lineupSheetName, newIndex, { 'template': template });
  ss.setActiveSheet(lineupSheet);
  initLineupSheet(lineupSheet);
}

function getLineupSheetName() {
  // tomorrow's date
  var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  var tomorrow = moment().add(1, 'days'); 
  var tomorrowStr = tomorrow.format('YYYY-MM-DD');
    
  // get the line-up name w/ tomorrow's date as default
  var ui = SA.getUi();
  var resp = ui.prompt('Lineup Name', 'Leave empty for "' + tomorrowStr + '"', ui.ButtonSet.OK_CANCEL);

  if (resp.getSelectedButton() == ui.Button.CANCEL) {
    return '';
  } else {
    var name = resp.getResponseText();
    if (name + "foo" == "foo") {
      name = tomorrowStr;
    }
  }
  return name;
}

function initLineupSheet(sheet) {
  setLineupDate(sheet);
  fillAvailability(sheet);
}

function setLineupDate(sheet) {
  sheet = sheet || promptForSheet();
  // try to get the lineup date from the sheet name
  var lineupDate = moment(sheet.getName());
  var lineupDateRange = findSheetNamedRange(sheet, "LineupDate");
  lineupDateRange.setValue(lineupDate.format('YYYY-MM-DD'));
}

function fillAvailability(sheet) {
  sheet = sheet || promptForSheet();
  
  var lineupDateString = findSheetNamedRange(sheet, "LineupDate").getDisplayValue();
  var lineupDate = moment(lineupDateString);
  Logger.log("got lineup date: " + lineupDate.format('YYYY-MM-DD'));
  
  var members = fetchRoster();
  var eventId = getEventId(lineupDate);
  var availability = fetchAvailability(eventId);
  var availableRowers = new Array();
  var unavailableRowers = new Array();
  var availableCoxswains = new Array();
  var unavailableCoxswains = new Array();

  for each (var m in members) {
    var available = availability[m.id];
    switch (true) {
      case (m.is_coxswain() && available):
        availableCoxswains.push(m.sheet_name);
        break;
      case (m.is_coxswain()):
        unavailableCoxswains.push(m.sheet_name);
        break;
      case (available):
        availableRowers.push(m.sheet_name);
        break;
      default:
        unavailableRowers.push(m.sheet_name);        
    }
  }
  fillAvailabilityRange(sheet, "AvailableRowers", availableRowers);
  fillAvailabilityRange(sheet, "UnavailableRowers", unavailableRowers);
  fillAvailabilityRange(sheet, "AvailableCoxswains", availableCoxswains);

  return;
}

function fillAvailabilityRange(sheet, rangeName, memberNames) {
  if (!memberNames.length) {
    Logger.log("availability list for %s was empty", rangeName);
    return;
  }
  var targetRange = findSheetNamedRange(sheet, rangeName);
  targetRange.clearContent();
  var fillRange = sheet.getRange(targetRange.getRow(), targetRange.getColumn(), memberNames.length);
  var values = _.map(memberNames, function(n) { return [n]; });
  fillRange.setValues(values);
  Logger.log("fill %s starting at %s", rangeName, fillRange.getA1Notation());
}

function fetchAvailability(eventId) {
  eventId = eventId || "199572191";
  var availabilityData = apiFetch("availabilities/search", { event_id: eventId });
  var availabilityMap = {};
  availabilityData.collection.items.forEach( function(availabilityData) {
    var memberId, statusCode;
    availabilityData.data.forEach( function(dataItem) {
      if (dataItem.name == "member_id") memberId = dataItem.value;
      if (dataItem.name == "status_code") statusCode = dataItem.value;
    });
    availabilityMap[memberId] = statusCode == 1 ? true : false;
  });
  return availabilityMap;
}

function getEventId(eventDate) {
  // URL example: https://api.teamsnap.com/v3/events/search?team_id=4969644&started_after=2019-04-01&started_before=2019-04-02
  // Note this assumes no more than one event per day
  var startedAfter = eventDate.format('YYYY-MM-DD');
  var startedBefore = eventDate.add(1, 'days').format('YYYY-MM-DD');
  var params = { team_id: teamId, started_after: startedAfter, started_before: startedBefore };
  var eventResp = apiFetch("events/search", params);
  var event = eventResp.collection.items[0];
  var eventId = event.rel.split("-")[1];
  return eventId;
}

function fetchRoster() {
    
  var member_data = apiFetch("members/search", { team_id: teamId });
  var custom_data = getMemberCustomFieldData();
  
  var members = member_data.collection.items.map( function(item) {
    var member = new Member(item['data']);
    member.set_custom_fields(custom_data[member.id]);
    return member;
  });
  
  // filter out inactive members and coaches
  members = _.filter(members, function(m) { return m.is_activated && !m.is_non_player });
  
  Logger.log("Fetched " + members.length + " members");

  var sorted_members = members.sort( function(a, b) {
    if (a.last_name == b.last_name) return 0;
    return a.last_name > b.last_name ? 1 : -1;
  });
  return sorted_members;
}

function getMemberCustomFieldData() {
  var custom_field_data = apiFetch("custom_data/search", { team_id: teamId });
  var customDataMemberMap = {};
  custom_field_data.collection.items.forEach( function(fieldItem) {
    var fieldData = fieldItem.data;
    var memberId, fieldName, fieldValue;
    fieldData.forEach (function(dataItem) {
      if (dataItem.name == "member_id") memberId = dataItem.value;
      if (dataItem.name == "name") fieldName = dataItem.value;
      if (dataItem.name == "value") fieldValue = dataItem.value;
    });
    if (customDataMemberMap[memberId] == undefined) customDataMemberMap[memberId] = {};
    customDataMemberMap[memberId][fieldName] = fieldValue;
  });
  return customDataMemberMap;
}

var Member = function(member_data) {
  for (i in member_data) {
    var attr = member_data[i];
    var prop = attr["name"];
    var val = attr["value"];
    var type = attr["type"];
    if (type == "DateTime") {
      this[prop] = new Date(val);
    } else {
      this[prop] = val;
    }
  }
  this.name = this.first_name + " " + this.last_name;
  this.sheet_name = this.last_name + ", " + this.first_name.substring(0, 1);
  this.email = this.email_addresses.join("; ");
  this.phone = this.phone_numbers.join("; ");
  this.age = function(asOf) {
    if (!this.birthday) {
      return '';
    }
    if (asOf !== undefined) {
      var asOfDate = new Date(asOf);
    } else {
      var asOfDate = new Date();
    }
    var birthDate = new Date(this.birthday);
    var age = asOfDate.getFullYear() - birthDate.getFullYear();
    var month = asOfDate.getMonth() - birthDate.getMonth();
    if (month < 0 || (month === 0 && asOfDate.getDate() < birthDate.getDate())) {
        age--;
    }
    return age;
  };
  this.rowing_age = function() {
    var endOfYear = moment(moment().year() + '-12-31');
    return this.age(endOfYear);
  };
  this.is_coxswain = function() {
    return this.Coxswain == 1;
  };
  this.set_custom_fields = function(custom_field_data) {
    for (attr in custom_field_data) {
      this[attr] = custom_field_data[attr];
    }
  };
};

function findCell(value, sheet, direction) {
  var data = sheet.getDataRange().getValues();

  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) { 
      if (data[i][j] == value) {
        Logger.log("Found at i: " + i + ", j: " + j);
        if (direction) {
          i += direction[0];
          j += direction[1];
        }
        var found = sheet.getRange(i + 1, j + 1);
        Logger.log("Returning range: " + found.getA1Notation());
        return found;
      }
    }    
  }
  Logger.log("Failed to find range in sheet '%s' with value '%s'", sheet.getName(), value);
}

function findSheetNamedRange(sheet, n) {  // just a wrapper
  var fullName = sheet.getName() + "!" + n;
  Logger.log("finding named range " + fullName);
  return SA.getActive().getRangeByName(fullName);
}

function promptFor(title, subtitle, defaultVal) {
  var ui = SA.getUi();
  var resp = ui.prompt(title, subtitle, ui.ButtonSet.OK_CANCEL);

  if (resp.getSelectedButton() == ui.Button.CANCEL) {
    return '';
  } else {
    var inputVal = resp.getResponseText();
    if (inputVal + "foo" == "foo") {
      inputVal = defaultVal;
    }
    return inputVal;
  }
}

function promptForSheet() {
  var sheetName = promptFor("Import availablity to which sheet?");
  return SA.getActiveSpreadsheet().getSheetByName(sheetName);
}
