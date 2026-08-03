function doGet(e) { return handle(e); }
function doPost(e) { return handle(e); }

function handle(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var p = (e && e.parameter) || {};

  if ((p.mode || '') === 'habits') return handleHabits(ss, p);
  if ((p.mode || '') === 'read_diary') return handleReadDiary(ss, p);
  if ((p.mode || '') === 'read_habits') return handleReadHabitsLog(ss, p);
  if ((p.mode || '') === 'garmin') return handleGarminWrite(ss, p);
  if ((p.mode || '') === 'read_garmin') return handleGarminRead(ss, p);
  if ((p.mode || '') === 'diag') return handleDiag(ss, p);

  var sheet = ss.getSheetByName('рефлексия');
  var text = p.text || '';
  var date = p.date || '';
  if (!text) return respond('ERROR: empty text');

  var lastRow = sheet.getLastRow();
  var row = lastRow + 1;
  if (date) {
    var dateCell = sheet.getRange(row, 1);
    dateCell.setValue(date);
    dateCell.setBackground('#93c47d');
    dateCell.setFontWeight('bold');
    dateCell.setHorizontalAlignment('right');
    row++;
  }

  var lines = text.split('\n');
  var count = 0;
  for (var i = 0; i < lines.length; i++) {
    if (lines[i].trim()) { sheet.getRange(row, 1).setValue(lines[i]); row++; count++; }
  }
  return respond('OK: added ' + count + ' rows');
}

function handleHabits(ss, p) {
  var sheet = ss.getSheetByName('трекер привычек');
  var dateStr = p.date || '';
  var vals = (p.vals || '').split(',');
  if (!dateStr || !vals.length) return respond('ERROR: need date and vals');

  var parts = dateStr.split('.');
  var d = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
  var days = ['воскресенье','понедельник','вторник','среда','четверг','пятница','суббота'];
  var tz = ss.getSpreadsheetTimeZone();

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var cVals = sheet.getRange(1, 3, lastRow, 1).getValues();

  var targetRow = 0, lastDateRow = 0;
  for (var i = 0; i < cVals.length; i++) {
    var v = cVals[i][0];
    var norm = '';
    if (v instanceof Date) {
      norm = Utilities.formatDate(v, tz, 'dd.MM.yyyy');
    } else if (typeof v === 'string' && v.trim()) {
      var pp = v.trim().split('.');
      if (pp.length === 3 && pp[0].length <= 2 && pp[2].length === 4) {
        norm = ('0'+pp[0]).slice(-2) + '.' + ('0'+pp[1]).slice(-2) + '.' + pp[2];
      }
    }
    if (norm) {
      lastDateRow = i + 1;
      if (norm === dateStr) targetRow = i + 1;
    }
  }
  if (!lastDateRow) lastDateRow = 4;

  if (!targetRow) {
    sheet.insertRowsAfter(lastDateRow, 1);
    targetRow = lastDateRow + 1;
    // наследуем оформление строки-донора и ставим настоящие чекбоксы
    sheet.getRange(lastDateRow, 1, 1, lastCol)
         .copyTo(sheet.getRange(targetRow, 1, 1, lastCol), {formatOnly: true});
    sheet.getRange(targetRow, 4, 1, Math.max(1, lastCol - 3)).insertCheckboxes();
    sheet.getRange(targetRow, 2).setValue(days[d.getDay()]);
    sheet.getRange(targetRow, 3).setValue(d);
  } else {
    sheet.getRange(targetRow, 4, 1, Math.max(1, lastCol - 3)).insertCheckboxes();
  }

  for (var j = 0; j < vals.length && j < 40; j++) {
    sheet.getRange(targetRow, 4 + j).setValue(vals[j] === '1');
  }
  return respond('OK: habits ' + dateStr + ' row ' + targetRow);
}


function getOrCreateSheet(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, 2).setValues([['дата', 'данные Garmin']]);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 100);
    sh.setColumnWidth(2, 700);
  }
  return sh;
}

function handleGarminWrite(ss, p) {
  var sheet = getOrCreateSheet(ss, 'гармин');
  var dateStr = p.date || '';
  var text = p.text || '';
  if (!text) return respond('ERROR: empty text');
  if (!dateStr) {
    var now = new Date();
    dateStr = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'dd.MM.yyyy');
  }
  // если строка за эту дату уже есть — обновляем, иначе добавляем
  var lastRow = sheet.getLastRow();
  var targetRow = 0;
  if (lastRow > 1) {
    var dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < dates.length; i++) {
      if (String(dates[i][0]).trim() === dateStr) targetRow = i + 2;
    }
  }
  if (!targetRow) targetRow = lastRow + 1;
  sheet.getRange(targetRow, 1).setValue(dateStr);
  sheet.getRange(targetRow, 2).setValue(text);
  return respond('OK: garmin ' + dateStr + ' row ' + targetRow);
}

function handleGarminRead(ss, p) {
  var sheet = ss.getSheetByName('гармин');
  if (!sheet) return respond('');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return respond('');
  var limitRows = Number(p.limit) || 30;
  var startRow = Math.max(2, lastRow - limitRows + 1);
  var data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 2).getValues();
  var lines = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0] && !data[i][1]) continue;
    lines.push(String(data[i][0]) + ': ' + String(data[i][1]).replace(/\n/g, ' '));
  }
  var text = lines.join('\n');
  var maxChars = Number(p.maxChars) || 12000;
  if (text.length > maxChars) text = text.slice(text.length - maxChars);
  return respond(text);
}

function handleDiag(ss, p) {
  var sheet = ss.getSheetByName('трекер привычек');
  var lastRow = sheet.getLastRow();
  var tz = ss.getSpreadsheetTimeZone();
  var cVals = sheet.getRange(1, 3, lastRow, 1).getValues();
  var found = [], lastDateRow = 0;
  for (var i = 0; i < cVals.length; i++) {
    var v = cVals[i][0];
    if (v instanceof Date) {
      lastDateRow = i + 1;
      found.push((i+1) + '=' + Utilities.formatDate(v, tz, 'dd.MM.yyyy') + '(Date)');
    } else if (typeof v === 'string' && v.trim()) {
      var pp = v.trim().split('.');
      if (pp.length === 3 && pp[2].length === 4) {
        lastDateRow = i + 1;
        found.push((i+1) + '=' + v.trim() + '(текст)');
      }
    }
  }
  var tail = found.slice(-5).join(' | ');
  return respond('lastRow=' + lastRow + ' | найдено дат: ' + found.length +
                 ' | последняя строка с датой: ' + lastDateRow +
                 ' | последние: ' + tail);
}

function respond(msg) {
  return ContentService.createTextOutput(msg).setMimeType(ContentService.MimeType.TEXT);
}

function handleReadDiary(ss, p) {
  var sheet = ss.getSheetByName('рефлексия');
  var last = sheet.getLastRow();
  var limitRows = Number(p.limit) || 500;
  var start = Math.max(1, last - limitRows + 1);
  var vals = sheet.getRange(start, 1, last - start + 1, 1).getValues();
  var tz = ss.getSpreadsheetTimeZone();
  var lines = [];
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i][0];
    if (v === '' || v === null) continue;
    if (v instanceof Date) lines.push(Utilities.formatDate(v, tz, 'dd.MM.yyyy'));
    else lines.push(String(v));
  }
  var text = lines.join('\n');
  var maxChars = Number(p.maxChars) || 24000;
  if (text.length > maxChars) text = text.slice(text.length - maxChars);
  return respond(text);
}

function handleReadHabitsLog(ss, p) {
  var sheet = ss.getSheetByName('трекер привычек');
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(4, 4, 1, lastCol - 3).getValues()[0];
  var limitRows = Number(p.limit) || 60;
  var startRow = Math.max(5, lastRow - limitRows + 1);
  var data = sheet.getRange(startRow, 3, lastRow - startRow + 1, lastCol - 2).getValues();
  var tz = ss.getSpreadsheetTimeZone();
  var lines = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var dateVal = row[0];
    if (!dateVal) continue;
    var dateStr = dateVal instanceof Date ? Utilities.formatDate(dateVal, tz, 'dd.MM.yyyy') : String(dateVal);
    var done = [];
    for (var c = 0; c < headers.length; c++) {
      if (row[1 + c] === true) done.push(headers[c]);
    }
    lines.push(dateStr + ': ' + (done.length ? done.join(', ') : '—'));
  }
  var text = lines.join('\n');
  var maxChars = Number(p.maxChars) || 16000;
  if (text.length > maxChars) text = text.slice(text.length - maxChars);
  return respond(text);
}
