function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
        return respond({ error: "Empty request received. Verify deployment." });
    }
    
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (action === 'register') return handleRegister(ss, data);
    if (action === 'login') return handleLogin(ss, data);
    if (action === 'createGroup') return handleCreateGroup(ss, data);
    if (action === 'getGroups') return handleGetGroups(ss, data);
    if (action === 'getGroupDetails') return handleGetGroupDetails(ss, data);
    if (action === 'addParticipant') return handleAddParticipant(ss, data);
    if (action === 'updateRestrictions') return handleUpdateRestrictions(ss, data);
    if (action === 'dispatch') return handleDispatch(ss, data);
    
    throw new Error("Action not recognized by Backend: " + action);
  } catch (err) {
    return respond({ error: err.message });
  }
}

function doOptions(e) {
  // Handing CORS pre-flight OPTIONS request explicitly
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}

function respond(responseObj) {
  return ContentService.createTextOutput(JSON.stringify(responseObj))
                       .setMimeType(ContentService.MimeType.JSON);
}

function normalizeEmail(email) {
  if (!email) return "";
  email = email.trim().toLowerCase();
  var parts = email.split("@");
  if (parts.length !== 2) return email;
  var localPart = parts[0];
  var domain = parts[1];
  if (localPart.indexOf("+") !== -1) localPart = localPart.split("+")[0];
  if (domain === "gmail.com" || domain === "googlemail.com") localPart = localPart.replace(/\./g, "");
  return localPart + "@" + domain;
}

function handleRegister(ss, data) {
  var sheet = ss.getSheetByName('Users');
  var userId = Utilities.getUuid();
  var normalizedEmail = normalizeEmail(data.email);
  var dataRange = sheet.getDataRange().getValues();
  for (var i = 1; i < dataRange.length; i++) {
    if (normalizeEmail(dataRange[i][1]) === normalizedEmail) return respond({ error: "Email already exists. Please log in." });
  }
  sheet.appendRow([userId, normalizedEmail, data.password, 'user']);
  return respond({ success: true, userId: userId, role: 'user' });
}

function handleLogin(ss, data) {
  var sheet = ss.getSheetByName('Users');
  var normalizedEmail = normalizeEmail(data.email);
  var dataRange = sheet.getDataRange().getValues();
  for (var i = 1; i < dataRange.length; i++) {
    if (normalizeEmail(dataRange[i][1]) === normalizedEmail && dataRange[i][2] === data.password) {
      return respond({ success: true, userId: dataRange[i][0], role: dataRange[i][3] });
    }
  }
  return respond({ error: "Invalid email or password." });
}

function handleCreateGroup(ss, data) {
  var sheet = ss.getSheetByName('Groups');
  var groupId = Utilities.getUuid();
  sheet.appendRow([groupId, data.ownerId, data.groupName, data.exchangeDate]);
  return respond({ success: true, group: { id: groupId, name: data.groupName, date: data.exchangeDate } });
}

function handleGetGroups(ss, data) {
  var sheet = ss.getSheetByName('Groups');
  var dataRange = sheet.getDataRange().getValues();
  var groups = [];
  for (var i = 1; i < dataRange.length; i++) {
    if (dataRange[i][1] === data.ownerId) {
      groups.push({ id: dataRange[i][0], name: dataRange[i][2], date: dataRange[i][3] });
    }
  }
  return respond({ success: true, groups: groups });
}

function handleGetGroupDetails(ss, data) {
  var groupSheet = ss.getSheetByName('Groups');
  var partsSheet = ss.getSheetByName('Participants');
  
  var groupData = groupSheet.getDataRange().getValues();
  var group = null;
  for (var i = 1; i < groupData.length; i++) {
    if (groupData[i][0] === data.groupId && groupData[i][1] === data.userId) {
      group = { id: groupData[i][0], name: groupData[i][2], date: groupData[i][3] };
      break;
    }
  }
  
  if (!group) return respond({ error: "Group not found or access denied." });
  
  var partsData = partsSheet.getDataRange().getValues();
  var participants = [];
  for (var j = 1; j < partsData.length; j++) {
    if (partsData[j][1] === data.groupId) {
      participants.push({
        id: partsData[j][0],
        name: partsData[j][2],
        email: partsData[j][3],
        restrictions: partsData[j][4] ? JSON.parse(partsData[j][4]) : []
      });
    }
  }
  return respond({ success: true, group: group, participants: participants });
}

function handleAddParticipant(ss, data) {
  var sheet = ss.getSheetByName('Participants');
  var partId = Utilities.getUuid();
  sheet.appendRow([partId, data.groupId, data.name, data.email, "[]"]);
  return respond({ success: true, participant: { id: partId, name: data.name, email: data.email, restrictions: [] } });
}

function handleUpdateRestrictions(ss, data) {
  var sheet = ss.getSheetByName('Participants');
  var partsData = sheet.getDataRange().getValues();
  for (var i = 1; i < partsData.length; i++) {
    if (partsData[i][0] === data.participantId && partsData[i][1] === data.groupId) {
      sheet.getRange(i + 1, 5).setValue(JSON.stringify(data.restrictions)); // Column E is restrictions
      return respond({ success: true });
    }
  }
  return respond({ error: "Participant not found." });
}

function handleDispatch(ss, data) {
  var partsSheet = ss.getSheetByName('Participants');
  var partsData = partsSheet.getDataRange().getValues();
  var participants = [];
  
  for (var j = 1; j < partsData.length; j++) {
    if (partsData[j][1] === data.groupId) {
      participants.push({
        id: partsData[j][0],
        name: partsData[j][2],
        contact: partsData[j][3],
        restrictions: partsData[j][4] ? JSON.parse(partsData[j][4]) : []
      });
    }
  }
  
  if (participants.length < 3) return respond({ error: "You need at least 3 participants to run an exchange." });
  
  var result = {};
  var givers = participants.map(function(p) { return p.id; });
  var receivers = participants.map(function(p) { return p.id; });
  
  function backtrack(index) {
    if (index === givers.length) return true;
    var giverId = givers[index];
    var giver = participants.filter(function(p) { return p.id === giverId; })[0];
    
    var available = receivers.filter(function(r) { 
      var notAssigned = true;
      for (var key in result) { if (result[key] === r) notAssigned = false; }
      return notAssigned;
    });
    
    shuffleArray(available);
    
    for (var i = 0; i < available.length; i++) {
      var receiverId = available[i];
      if (receiverId !== giverId && giver.restrictions.indexOf(receiverId) === -1) {
        result[giverId] = receiverId;
        if (backtrack(index + 1)) return true;
        delete result[giverId]; // backtrack
      }
    }
    return false;
  }
  
  var success = backtrack(0);
  
  if (success) {
    for (var k = 0; k < participants.length; k++) {
      var giver = participants[k];
      var receiverId = result[giver.id];
      var receiver = participants.filter(function(p) { return p.id === receiverId; })[0];
      
      var emailBody = "Hi " + giver.name + ",\n\n" +
                      "You have been invited to a Secret Santa exchange!\n\n" +
                      "You are the Secret Santa for: " + receiver.name + "\n\n" +
                      "Remember, it's a secret! Happy gifting! 🎁";
                      
      MailApp.sendEmail(giver.contact, "Your Secret Santa Match! 🎅", emailBody);
    }
    return respond({ success: true });
  } else {
    return respond({ error: "Impossible to match! The restrictions are too strict." });
  }
}

function shuffleArray(array) {
  for (var i = array.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
}

function cleanupOldGroups() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var groupsSheet = ss.getSheetByName('Groups');
  var partsSheet = ss.getSheetByName('Participants');
  var groupsData = groupsSheet.getDataRange().getValues();
  var partsData = partsSheet.getDataRange().getValues();
  var today = new Date();
  var groupsToDelete = [];
  
  for (var i = groupsData.length - 1; i > 0; i--) {
    var exchangeDate = new Date(groupsData[i][3]);
    if ((today.getTime() - exchangeDate.getTime()) / (1000 * 3600 * 24) > 30) {
      groupsToDelete.push(groupsData[i][0]);
      groupsSheet.deleteRow(i + 1);
    }
  }
  for (var j = partsData.length - 1; j > 0; j--) {
    if (groupsToDelete.indexOf(partsData[j][1]) !== -1) partsSheet.deleteRow(j + 1);
  }
}
