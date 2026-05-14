// ============================================================
//  PRODUCTION OPS TRACKER — Google Apps Script Backend v2
//  Paste this entire file into your Apps Script editor
// ============================================================

var REPAIR_HEADERS  = ['id','campus','type','item','serial','priority','desc','assigned','status','reported','notes','comments'];
var PROJECT_HEADERS = ['id','name','campus','category','budget','requestedBy','desc','status','progress','notes'];
var USER_HEADERS    = ['email','name','role'];  // role: admin | operator | general

// ── Entry point: GET ─────────────────────────────────────────
function doGet(e) {
  var action = e.parameter.action;
  var result;
  try {
    if      (action === 'getRepairs')  result = getRepairs();
    else if (action === 'getProjects') result = getProjects();
    else if (action === 'getRole')     result = getRole(e.parameter.email);
    else if (action === 'getUsers')    result = getUsers();
    else result = { error: 'Unknown action' };
  } catch(err) {
    result = { error: err.message };
  }
  return jsonResponse(result);
}

// ── Entry point: POST ────────────────────────────────────────
function doPost(e) {
  var payload = JSON.parse(e.postData.contents);
  var action  = payload.action;
  var result;
  try {
    if      (action === 'saveRepair')    result = saveRepair(payload.data);
    else if (action === 'updateRepair')  result = updateRepair(payload.data);
    else if (action === 'deleteRepair')  result = deleteRow('Repairs', payload.id);
    else if (action === 'addComment')    result = addRepairComment(payload.repairId, payload.comment);
    else if (action === 'saveProject')   result = saveProject(payload.data);
    else if (action === 'updateProject') result = updateProject(payload.data);
    else if (action === 'deleteProject') result = deleteRow('Projects', payload.id);
    else if (action === 'notifyBulk')   result = notifyBulk(payload.ids);
    else result = { error: 'Unknown action' };
  } catch(err) {
    result = { error: err.message };
  }
  return jsonResponse(result);
}

// ── Role lookup ───────────────────────────────────────────────
function getRole(email) {
  if (!email) return { error: 'No email provided' };
  var sheet = getSheet('Users');
  var data  = sheet.getDataRange().getValues();
  var emailLower = email.toLowerCase().trim();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().trim() === emailLower) {
      return { role: data[i][2], name: data[i][1] };
    }
  }
  return { role: null };
}

// ── Sheet helpers ─────────────────────────────────────────────
function getSheet(name) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    var headers;
    if      (name === 'Repairs')  headers = REPAIR_HEADERS;
    else if (name === 'Projects') headers = PROJECT_HEADERS;
    else if (name === 'Users')    headers = USER_HEADERS;
    else headers = [];
    if (headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight('bold')
           .setBackground('#1a1f2e')
           .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
    if (name === 'Users') {
      sheet.appendRow(['admin@christfellowship.church', 'Admin User', 'admin']);
    }
  }
  return sheet;
}

function sheetToObjects(sheet, headers) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function genId() {
  return Utilities.getUuid().replace(/-/g,'').slice(0,12);
}

function getUsers() {
  return { data: sheetToObjects(getSheet('Users'), USER_HEADERS) };
}

function getEmailByName(name) {
  if (!name) return null;
  var userData = getSheet('Users').getDataRange().getValues();
  for (var i = 1; i < userData.length; i++) {
    if (String(userData[i][1]).toLowerCase().trim() === String(name).toLowerCase().trim()) {
      return userData[i][0] || null;
    }
  }
  return null;
}

function notifyAssigned(assignedName, itemName, campus, type, priority, isNew) {
  Logger.log('notifyAssigned called — name: ' + assignedName + ', item: ' + itemName);
  var email = getEmailByName(assignedName);
  Logger.log('Email lookup result: ' + email);
  if (!email) {
    Logger.log('ERROR: No email found for "' + assignedName + '" — check Users sheet spelling');
    return;
  }
  var subject = isNew
    ? 'New repair assigned to you: ' + itemName
    : 'You have been assigned to a repair: ' + itemName;
  var body = 'Hi ' + assignedName + ',\n\n'
    + (isNew ? 'A new repair ticket has been assigned to you.' : 'You have been assigned to an existing repair ticket.') + '\n\n'
    + 'Item:     ' + itemName + '\n'
    + 'Campus:   ' + campus + '\n'
    + 'Type:     ' + type + '\n'
    + 'Priority: ' + priority + '\n\n'
    + 'View it here: https://cfproduction.github.io/cf-production-ops/\n\n'
    + '— CF Production Ops';
  try {
    GmailApp.sendEmail(email, subject, body);
    Logger.log('Email sent successfully to ' + email);
  } catch(e) {
    Logger.log('ERROR sending email: ' + e.message);
  }
}

// ── Quick test — run this in the Apps Script editor to verify email ──
function testNotifyAssigned() {
  // Change these to match a real name in your Users sheet and confirm the email arrives
  var testName = 'Yoel Torres';
  var testItem = 'Test Repair Item';
  Logger.log('=== testNotifyAssigned START ===');
  notifyAssigned(testName, testItem, 'Palm Beach Gardens', 'Audio', 'High', false);
  Logger.log('=== testNotifyAssigned END ===');
}

function getRepairs() {
  return { data: sheetToObjects(getSheet('Repairs'), REPAIR_HEADERS) };
}

function addRepairComment(repairId, comment) {
  var sheet = getSheet('Repairs');
  var vals  = sheet.getDataRange().getValues();
  var hdrs  = vals[0];
  var commentsCol = hdrs.indexOf('comments');
  if (commentsCol === -1) {
    commentsCol = hdrs.length;
    sheet.getRange(1, commentsCol + 1).setValue('comments')
         .setFontWeight('bold').setBackground('#1a1f2e').setFontColor('#ffffff');
  }
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(repairId)) {
      var existing = [];
      try { existing = JSON.parse(vals[i][commentsCol] || '[]'); } catch(e) {}
      existing.push(comment);
      sheet.getRange(i + 1, commentsCol + 1).setValue(JSON.stringify(existing));

      var assignedName = vals[i][hdrs.indexOf('assigned')];
      var itemName     = vals[i][hdrs.indexOf('item')];
      if (assignedName) {
        var userSheet = getSheet('Users');
        var userData  = userSheet.getDataRange().getValues();
        for (var j = 1; j < userData.length; j++) {
          if (String(userData[j][1]).toLowerCase().trim() === String(assignedName).toLowerCase().trim()) {
            var email = userData[j][0];
            if (email) {
              try {
                GmailApp.sendEmail(
                  email,
                  'New comment on repair: ' + itemName,
                  comment.author + ' commented on "' + itemName + '":\n\n"' + comment.text + '"\n\nView: https://cfproduction.github.io/cf-production-ops/'
                );
              } catch(emailErr) {}
            }
            break;
          }
        }
      }
      return { success: true };
    }
  }
  return { error: 'Repair not found' };
}

function saveRepair(data) {
  var sheet = getSheet('Repairs');
  data.id       = genId();
  data.reported = data.reported || new Date().toISOString().slice(0,10);
  data.status   = data.status || 'Reported';
  var row = REPAIR_HEADERS.map(function(h) { return data[h] || ''; });
  sheet.appendRow(row);
  if (data.assigned) {
    notifyAssigned(data.assigned, data.item, data.campus, data.type, data.priority, true);
  }
  return { success: true, id: data.id };
}

function updateRepair(data) {
  var sheet = getSheet('Repairs');
  var vals  = sheet.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(data.id)) {
      var oldAssigned = String(vals[i][REPAIR_HEADERS.indexOf('assigned')] || '');
      var row = REPAIR_HEADERS.map(function(h) {
        var val = data[h] !== undefined ? data[h] : vals[i][REPAIR_HEADERS.indexOf(h)];
        return val === undefined || val === null ? '' : val;
      });
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      // Notify only when the caller explicitly requests it
      var newAssigned = String(row[REPAIR_HEADERS.indexOf('assigned')] || '');
      Logger.log('updateRepair — notifyAssignee flag: ' + data.notifyAssignee + ', assigned: ' + newAssigned);
      if (data.notifyAssignee && newAssigned) {
        var itemName = row[REPAIR_HEADERS.indexOf('item')];
        var campus   = row[REPAIR_HEADERS.indexOf('campus')];
        var type     = row[REPAIR_HEADERS.indexOf('type')];
        var priority = row[REPAIR_HEADERS.indexOf('priority')];
        notifyAssigned(newAssigned, itemName, campus, type, priority, false);
      }
      return { success: true };
    }
  }
  return { error: 'Record not found' };
}

function getProjects() {
  return { data: sheetToObjects(getSheet('Projects'), PROJECT_HEADERS) };
}

function saveProject(data) {
  var sheet = getSheet('Projects');
  data.id       = genId();
  data.status   = data.status || 'Requested';
  data.progress = data.progress || 0;
  var row = PROJECT_HEADERS.map(function(h) { return data[h] || ''; });
  sheet.appendRow(row);
  return { success: true, id: data.id };
}

function updateProject(data) {
  var sheet = getSheet('Projects');
  var vals  = sheet.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0] === data.id) {
      var row = PROJECT_HEADERS.map(function(h) {
        return data[h] !== undefined ? data[h] : vals[i][PROJECT_HEADERS.indexOf(h)];
      });
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return { success: true };
    }
  }
  return { error: 'Record not found' };
}

function deleteRow(sheetName, id) {
  var sheet = getSheet(sheetName);
  var vals  = sheet.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Record not found' };
}

// ── Bulk notify assignees ─────────────────────────────────────
function notifyBulk(ids) {
  if (!ids || !ids.length) return { success: true, count: 0 };
  var sheet = getSheet('Repairs');
  var vals  = sheet.getDataRange().getValues();

  // Build map: assigneeName -> [ticket info, ...]
  var assigneeTickets = {};
  ids.forEach(function(rid) {
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === String(rid)) {
        var name = String(vals[i][REPAIR_HEADERS.indexOf('assigned')] || '').trim();
        if (name) {
          if (!assigneeTickets[name]) assigneeTickets[name] = [];
          assigneeTickets[name].push({
            item:     vals[i][REPAIR_HEADERS.indexOf('item')],
            campus:   vals[i][REPAIR_HEADERS.indexOf('campus')],
            priority: vals[i][REPAIR_HEADERS.indexOf('priority')],
            status:   vals[i][REPAIR_HEADERS.indexOf('status')]
          });
        }
        break;
      }
    }
  });

  var count = 0;
  Object.keys(assigneeTickets).forEach(function(name) {
    var email = getEmailByName(name);
    if (!email) return;
    var tickets  = assigneeTickets[name];
    var itemList = tickets.map(function(t, i) {
      return (i + 1) + '. ' + t.item + ' — ' + t.campus + ' | ' + t.priority + ' | ' + t.status;
    }).join('\n');
    var subject = 'Repair reminder: ' + tickets.length + ' ticket' + (tickets.length !== 1 ? 's' : '') + ' assigned to you';
    var body = 'Hi ' + name + ',\n\n'
      + 'Here\'s a reminder about your currently assigned repair ticket' + (tickets.length !== 1 ? 's' : '') + ':\n\n'
      + itemList + '\n\n'
      + 'View them here: https://cfproduction.github.io/cf-production-ops/\n\n'
      + '— CF Production Ops';
    try { GmailApp.sendEmail(email, subject, body); count++; } catch(e) {}
  });

  return { success: true, count: count };
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
