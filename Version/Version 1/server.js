const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const multer = require('multer');
const upload = multer({ dest: 'uploads/' });

const app = express();
app.use(cors());
app.use(express.json());

function getSeatNumbers(seatLabels) {
  return seatLabels.map(seat => {
    const match = seat.match(/\d+/);
    return match ? parseInt(match[0]) : 0;
  });
}

app.post('/api/upload', upload.single('file'), (req, res) => {
  const workbook = XLSX.readFile(req.file.path);
  const sheetName = workbook.SheetNames[0];
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

  
  const bookings = data.map(row => {
    const seats = row.Seats.split(',').map(s => s.trim());
    const seatNumbers = getSeatNumbers(seats);
    const maxSeat = Math.max(...seatNumbers);
    return {
      booking_id: row.Booking_id || row.Booking_ID || row['Booking ID'],
      maxSeat
    };
  });

  bookings.sort((a, b) => {
    if (a.maxSeat === b.maxSeat) {
      return parseInt(a.booking_id) - parseInt(b.booking_id);
    }
    return b.maxSeat - a.maxSeat;
  });

  const sequence = bookings.map((b, i) => ({
    Seq: i + 1,
    Booking_ID: b.booking_id
  }));

  res.json(sequence);
});

app.listen(3000, () => console.log('Server running on port 3000'));