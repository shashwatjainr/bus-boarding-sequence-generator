const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const multer = require('multer');
const upload = multer({ dest: 'uploads/' });

const app = express();
app.use(cors());
app.use(express.json());

function getCol(row, keys) {
  keys = Array.isArray(keys) ? keys : [keys];
  for (var i = 0; i < keys.length; i++) {
    for (var k in row) {
      if (k.toLowerCase().replace(/[\s_]+/g, '').includes(keys[i].toLowerCase())) return k;
    }
  }
  return null;
}

function parseSeatLabel(label, warnings, booking_id) {
  label = label.trim();
  if (!label) return null;
  var match = label.match(/^([A-Za-z])0*([1-9]\d*)$/);
  if (!match) {
    warnings.push("Booking " + booking_id + ": invalid seat label '" + label + "'");
    return null;
  }
  var letter = match[1].toUpperCase();
  var num = parseInt(match[2], 10);
  if (num < 1 || num > 25) {
    warnings.push("Booking " + booking_id + ": invalid seat number '" + label + "'");
    return null;
  }
  return { letter: letter, num: num };
}

app.post('/api/upload', upload.single('file'), function (req, res) {
  var warnings = [];
  if (!req.file) return res.json({ sequence: [], warnings: ["No file uploaded"] });
  var workbook;
  try { workbook = XLSX.readFile(req.file.path); }
  catch (e) { return res.json({ sequence: [], warnings: ["Excel file could not be read"] }); }
  var sheetName = workbook.SheetNames[0];
  var data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  if (!data.length) return res.json({ sequence: [], warnings: ["Excel Sheet is empty"] });

  var row0 = data[0];
  var bookingCol = getCol(row0, ["booking"]);
  var seatCol = getCol(row0, ["seat"]);
  if (!bookingCol || !seatCol) return res.json({ sequence: [], warnings: ["Missing booking/seat column"] });

  var allLetters = {};
  var seatUsage = {};
  var bookings = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var booking = row[bookingCol] ? row[bookingCol].toString().trim() : "";
    var seatstr = row[seatCol] ? row[seatCol].toString().trim() : "";
    if (!booking || !seatstr) {
      warnings.push("Row " + (i+2) + " missing Booking_ID or Seats");
      continue;
    }
    var seatLabels = seatstr.split(/[,;]+/).map(function(s){ return s.trim(); }).filter(Boolean);
    var validSeats = [];
    var seenSeats = {};
    for (var j = 0; j < seatLabels.length; j++) {
      var seatInfo = parseSeatLabel(seatLabels[j], warnings, booking);
      if (seatInfo && !seenSeats[seatLabels[j].toUpperCase()]) {
        validSeats.push(seatInfo);
        seenSeats[seatLabels[j].toUpperCase()] = true;
        allLetters[seatInfo.letter] = 1;
        var seatStr = seatInfo.letter + seatInfo.num;
        if (seatUsage[seatStr] && seatUsage[seatStr] !== booking) {
          warnings.push("Seat " + seatStr + " is duplicated in Booking " + seatUsage[seatStr] + " and " + booking);
          return res.json({ sequence: [], warnings: warnings });
        }
        seatUsage[seatStr] = booking;
      }
    }
    if (validSeats.length) {
      bookings.push({
        booking_id: booking,
        validSeats: validSeats
      });
    } else {
      warnings.push("Booking " + booking + " skipped — no valid seats");
    }
  }

  var letterList = Object.keys(allLetters).sort();
  if (letterList.length > 6) {
    warnings.push("Too many seat letters found (>6): " + letterList.join(", "));
    return res.json({ sequence: [], warnings: warnings });
  }

  var leftWindow = letterList.length ? letterList[0] : null;
  var rightWindow = letterList.length ? letterList[letterList.length-1] : null;
  var isWindowLetter = function(letter){ return letter == leftWindow || letter == rightWindow; };

  for (var i = 0; i < bookings.length; i++) {
    var bk = bookings[i];
    var windowRows = [];
    var aisleRows = [];
    var seatLettersUsed = {};
    for (var j = 0; j < bk.validSeats.length; j++) {
      var s = bk.validSeats[j];
      seatLettersUsed[s.letter] = 1;
      (isWindowLetter(s.letter) ? windowRows : aisleRows).push(s.num);
    }
    bk.farthestWindowRow = windowRows.length ? Math.max.apply(null, windowRows) : null;
    bk.farthestAisleRow = aisleRows.length ? Math.max.apply(null, aisleRows) : null;
    bk.validSeatCount = bk.validSeats.length;
    bk.isWindowPriority = windowRows.length > 0;
    bk.seatLettersUsed = Object.keys(seatLettersUsed).sort();
    bk.aisleLettersUsed = Object.keys(seatLettersUsed).filter(function(l){ return !isWindowLetter(l);}).sort();
  }

  bookings.sort(function(a, b){
    var wa = a.farthestWindowRow === null ? -Infinity : a.farthestWindowRow;
    var wb = b.farthestWindowRow === null ? -Infinity : b.farthestWindowRow;
    if (wb != wa) return wb - wa;
    var aa = a.farthestAisleRow === null ? -Infinity : a.farthestAisleRow;
    var ab = b.farthestAisleRow === null ? -Infinity : b.farthestAisleRow;
    if (ab != aa) return ab - aa;
    if (a.validSeatCount != b.validSeatCount) return a.validSeatCount - b.validSeatCount;
    var nai = /^\d+$/.test(a.booking_id), nbi = /^\d+$/.test(b.booking_id);
    if (nai && nbi) return parseInt(a.booking_id, 10) - parseInt(b.booking_id, 10);
    return (a.booking_id < b.booking_id) ? -1 : (a.booking_id > b.booking_id) ? 1 : 0;
  });

  var sequence = bookings.map(function(b, i){
    return { Seq: i+1, Booking_ID: b.booking_id, reason: b.isWindowPriority ? "window, farthestWindowRow=" + b.farthestWindowRow : "aisle, farthestAisleRow=" + b.farthestAisleRow };
  });

  res.json({ sequence: sequence, warnings: warnings, windowLetters: [leftWindow, rightWindow] });
});

app.listen(3000, function () { console.log('Server running on port 3000'); });