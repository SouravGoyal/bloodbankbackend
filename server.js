const express = require('express');
const multer = require('multer');
const fs = require('fs');
const cors = require('cors');
const { utils, read, writeFile } = require('xlsx');
const path = require('path');

const app = express();
const upload = multer();

const FILE_PATH = path.join(__dirname, 'FileName', 'user_data.xlsx');

app.use(cors());
app.use(express.json());

// Endpoint to get data from the Excel file
app.get('/api/data', (req, res) => {
  try {
    if (fs.existsSync(FILE_PATH)) {
      const workbook = read(fs.readFileSync(FILE_PATH), { type: 'buffer' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = utils.sheet_to_json(sheet);
      res.json(data);
    } else {
      res.json([]);
    }
  } catch (error) {
    console.error('Error reading Excel file:', error);
    res.status(500).json({ error: 'Error reading Excel file' });
  }
});

// Endpoint to save data to the Excel file
app.post('/api/data', upload.none(), (req, res) => {
  try {
    const newData = req.body;
    let data = [];

    if (fs.existsSync(FILE_PATH)) {
      const workbook = read(fs.readFileSync(FILE_PATH), { type: 'buffer' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      data = utils.sheet_to_json(sheet);
    }

    data.push(newData);

    const ws = utils.json_to_sheet(data);
    const wb = utils.book_new();
    utils.book_append_sheet(wb, ws, 'Users');
    writeFile(wb, FILE_PATH);

    res.json({ message: 'Data saved successfully' });
  } catch (error) {
    console.error('Error saving data to Excel file:', error);
    res.status(500).json({ error: 'Error saving data to Excel file' });
  }
});

app.listen(5000, () => {
  console.log('Server is running on port 5000');
});