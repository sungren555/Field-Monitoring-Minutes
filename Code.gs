/**
 * Field Activity Log — Apps Script backend
 * Reads staff + activities from tab "Staff/Work"
 * Writes submissions to a monthly tab (e.g. "Oct 2026")
 */
var SHEET_ID = "1wmMW8aXqnP8UwGqnvdw-ZIuMCSQUtKxytOXS0ADQXks";
var CONFIG_TAB = "Staff/Work";
var MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
var HEADERS = ["Timestamp","Date","Staff","Section","Project","Category","Activity","Location","Hours","Participants","Notes"];

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var a = (e.parameter.action || "config");
    if (a === "config") return json_(getConfig_());
    if (a === "logs") return json_(getLogs_(e.parameter.staff || "", e.parameter.month || ""));
    return json_({ok: false, error: "unknown action"});
  } catch (err) {
    return json_({ok: false, error: String(err)});
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var p = JSON.parse(e.postData.contents);
    var hours = Number(p.hours);
    if (!p.staff || !p.date || !p.activity || !p.location || !(hours > 0)) {
      return json_({ok: false, error: "missing fields"});
    }
    var parts = String(p.date).split("-");
    var y = Number(parts[0]), m = Number(parts[1]), d = Number(parts[2]);
    if (!y || !m || !d) return json_({ok: false, error: "bad date"});
    var tab = MONTHS[m - 1] + " " + y;
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(tab);
    if (!sh) {
      sh = ss.insertSheet(tab);
      sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
      sh.setFrozenRows(1);
    }
    var row = [
      new Date(), new Date(y, m - 1, d), p.staff, p.section || "", p.project || "",
      p.category || "", p.activity, p.location, hours, Number(p.participants) || "", p.notes || ""
    ];
    sh.appendRow(row);
    var last = sh.getLastRow();
    sh.getRange(last, 1).setNumberFormat("yyyy-mm-dd hh:mm");
    sh.getRange(last, 2).setNumberFormat("yyyy-mm-dd");
    return json_({ok: true, tab: tab});
  } catch (err) {
    return json_({ok: false, error: String(err)});
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function getConfig_() {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(CONFIG_TAB);
  var last = sh.getLastRow();
  var v = sh.getRange(2, 1, Math.max(last - 1, 1), 4).getValues(); // A:D from row 2
  var staff = [], act = {"IBIS Rice": [], "CE": []};
  var cur = {"IBIS Rice": null, "CE": null};
  var cols = {"IBIS Rice": 2, "CE": 3};
  for (var i = 0; i < v.length; i++) {
    var name = String(v[i][0]).trim();
    if (name && name !== "Staff") staff.push({name: name, section: String(v[i][1]).trim()});
    Object.keys(cols).forEach(function (proj) {
      var t = String(v[i][cols[proj]]).trim();
      if (!t || t === proj) return;
      if (/^Category\s*\d+/i.test(t)) {
        cur[proj] = {category: t, items: []};
        act[proj].push(cur[proj]);
      } else if (cur[proj]) {
        cur[proj].items.push(t);
      }
    });
  }
  return {ok: true, staff: staff, activities: act};
}

function getLogs_(staff, month) {
  var now = new Date();
  var tab = month || (MONTHS[now.getMonth()] + " " + now.getFullYear());
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(tab);
  if (!sh || sh.getLastRow() < 2) return {ok: true, tab: tab, logs: []};
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues();
  var tz = Session.getScriptTimeZone();
  var out = [];
  for (var i = v.length - 1; i >= 0 && out.length < 20; i--) {
    if (String(v[i][2]) !== staff) continue;
    out.push({
      date: Utilities.formatDate(new Date(v[i][1]), tz, "yyyy-MM-dd"),
      project: v[i][4], activity: v[i][6], location: v[i][7], hours: v[i][8]
    });
  }
  return {ok: true, tab: tab, logs: out};
}
