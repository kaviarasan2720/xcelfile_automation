const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const port = 3000;

// Directory to store uploaded files and created data file
const uploadsDirectory = path.join(__dirname, 'uploads');
const dataFile = path.join(uploadsDirectory, 'data.xlsx');

// Create uploads directory if it doesn't exist
if (!fs.existsSync(uploadsDirectory)) {
  fs.mkdirSync(uploadsDirectory);
}

// Multer configuration for file upload
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, uploadsDirectory); // Uploads folder where files will be stored
  },
  filename: function (req, file, cb) {
    cb(null, file.originalname); // Use original filename
  }
});

const upload = multer({ storage: storage });

// Serve uploaded files statically
app.use('/uploads', express.static('uploads'));

// POST endpoint for uploading Excel file
app.post('/upload', upload.single('excelFile'), (req, res) => {
  try {
    // Check if data file exists, create if not
    if (!fs.existsSync(dataFile)) {
      const newWorkbook = xlsx.utils.book_new(); // Create a new workbook
      const newSheet = xlsx.utils.json_to_sheet([]); // Create a new empty sheet
      xlsx.utils.book_append_sheet(newWorkbook, newSheet, 'Sheet1'); // Append the sheet to the workbook
      xlsx.writeFile(newWorkbook, dataFile); // Write the workbook to dataFile
    }

    const filePath = req.file.path;
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const excelData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    // Load the existing data file
    const existingWorkbook = xlsx.readFile(dataFile);
    const existingSheet = existingWorkbook.Sheets[existingWorkbook.SheetNames[0]];

    // Append the uploaded data to the existing sheet
    const updatedSheet = xlsx.utils.sheet_add_json(existingSheet, excelData, { skipHeader: true, origin: -1 });

    // Create a new workbook with the updated sheet
    const newWorkbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(newWorkbook, updatedSheet, sheetName);

    // Write the updated workbook back to dataFile
    xlsx.writeFile(newWorkbook, dataFile);

    // Remove the uploaded file after processing (optional)
    fs.unlinkSync(filePath);

    res.status(200).send('File uploaded and data inserted into data file successfully');
  } catch (err) {
    console.error(err);
    res.status(500).send('Error uploading file and inserting data');
  }
});

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
