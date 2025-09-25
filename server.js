const express = require("express");
const cors = require("cors");
const XLSX = require("xlsx");
const multer = require("multer");
const upload = multer({ dest: "uploads/" });

const app = express();
app.use(cors());
app.use(express.json());

function getCol(row, keys) {
  keys = Array.isArray(keys) ? keys : [keys];
  for (let i = 0; i < keys.length; i++) {
    for (let k in row) {
      if (
        k.toLowerCase().replace(/[\s\_\-]/g, "") ===
        keys[i].toLowerCase().replace(/[\s\_\-]/g, "")
      ) {
        return k;
      }
    }
  }
  return null;
}

function parseSeatLabel(label, warnings, bookingid) {
  label = label.trim();
  if (!label) return null;
  let match = label.match(/^([A-Za-z])\s*0*([1-9][0-9]*)$/);
  if (!match) {
    warnings.push(`Booking ${bookingid}: invalid seat label '${label}'`);
    return null;
  }
  let letter = match[1].toUpperCase();
  let num = parseInt(match[2], 10);
  if (num < 1 || num > 25) {
    warnings.push(`Booking ${bookingid}: invalid seat number '${label}'`);
    return null;
  }
  return { letter, num, raw: label };
}

app.post("/api/upload", upload.single("file"), function (req, res) {
  let warnings = [];
  if (!req.file)
    return res.json({ sequence: [], warnings: ["No file uploaded"] });

  let workbook;
  try {
    workbook = XLSX.readFile(req.file.path);
  } catch (e) {
    return res.json({
      sequence: [],
      warnings: ["Excel file could not be read"],
    });
  }

  let sheetName = workbook.SheetNames[0];
  let data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  if (!data.length)
    return res.json({ sequence: [], warnings: ["Excel Sheet is empty"] });

  let bookingCol = getCol(data[0], ["booking", "BookingID", "id"]);
  let seatCol = getCol(data[0], ["seat", "Seats", "seat_label"]);
  if (!bookingCol || !seatCol)
    return res.json({
      sequence: [],
      warnings: ["Missing booking/seat column"],
    });

  const minLetter = "A";
  const maxLetter = "F";
  const letterCodes = [];
  for (let c = minLetter.charCodeAt(0); c <= maxLetter.charCodeAt(0); c++) {
    letterCodes.push(String.fromCharCode(c));
  }

  const windowLetters = [minLetter, maxLetter]; // A, F
  const middleLetters = ["B", "E"];
  const aisleLetters = ["C", "D"];

  function inRowRange(letter) {
    return (
      /^[A-Z]$/.test(letter) &&
      letter.charCodeAt(0) >= minLetter.charCodeAt(0) &&
      letter.charCodeAt(0) <= maxLetter.charCodeAt(0)
    );
  }

  let seatUsage = {};
  let bookings = [];
  for (let i = 0; i < data.length; i++) {
    let row = data[i];
    let booking = row[bookingCol] ? row[bookingCol].toString().trim() : "";
    let seatstr = row[seatCol] ? row[seatCol].toString().trim() : "";
    if (!booking || !seatstr) {
      warnings.push(`Row ${i + 2}: missing BookingID or Seats`);
      continue;
    }
    let seatLabels = seatstr.split(/[,;\|]+/).map((s) => s.trim()).filter(Boolean);

    let validSeats = [];
    let seenSeats = {};

    for (let seat of seatLabels) {
      let seatInfo = parseSeatLabel(seat, warnings, booking);
      if (seatInfo) {
        if (!inRowRange(seatInfo.letter)) {
          warnings.push(`Booking ${booking}: Out of bound row range: ${seatInfo.raw}`);
          continue;
        }
        if (!seenSeats[seat.toUpperCase()]) {
          validSeats.push(seatInfo);
          seenSeats[seat.toUpperCase()] = true;
        }
        let seatStr = seatInfo.letter + seatInfo.num;
        if (seatUsage[seatStr] && seatUsage[seatStr] !== booking) {
          warnings.push(`Seat ${seatStr} is duplicated in Booking ${seatUsage[seatStr]} and ${booking}`);
        }
        seatUsage[seatStr] = booking;
      }
    }

    bookings.push({
      bookingid: booking,
      validSeats,
      rowIndex: i,
    });
  }

  let validatedBookings = bookings.filter((bk) => bk.validSeats.length > 0);

  function getFarthestSeatNum(bk) {
    return bk.validSeats.reduce((max, s) => (s.num > max ? s.num : max), -Infinity);
  }

  function getSeatLetterPriority(letter) {
    if (windowLetters.includes(letter)) return 1;
    else if (middleLetters.includes(letter)) return 2;
    else if (aisleLetters.includes(letter)) return 3;
    else return 4;
  }

  function getHighestPrioritySeatInRow(bk, rowNum) {
    let seatsInRow = bk.validSeats.filter((s) => s.num === rowNum);
    seatsInRow.sort((a, b) => getSeatLetterPriority(a.letter) - getSeatLetterPriority(b.letter));
    return seatsInRow.length ? getSeatLetterPriority(seatsInRow[0].letter) : 4;
  }

  validatedBookings.sort((a, b) => {
    let aMax = getFarthestSeatNum(a);
    let bMax = getFarthestSeatNum(b);
    if (aMax !== bMax) return bMax - aMax;

    let aPriority = getHighestPrioritySeatInRow(a, aMax);
    let bPriority = getHighestPrioritySeatInRow(b, bMax);
    if (aPriority !== bPriority) return aPriority - bPriority;

    let nai = /^\d+$/.test(a.bookingid);
    let nbi = /^\d+$/.test(b.bookingid);
    if (nai && nbi) return parseInt(a.bookingid, 10) - parseInt(b.bookingid, 10);
    return a.bookingid < b.bookingid ? -1 : a.bookingid > b.bookingid ? 1 : 0;
  });

  let sequence = validatedBookings.map((b, i) => ({
    Seq: i + 1,
    BookingID: b.bookingid,
  }));

  res.json({
    sequence,
    warnings,
    rowRange: `${minLetter}-${maxLetter}`,
  });
});

app.listen(3000, () => {
  console.log("Server running on port 3000");
});
