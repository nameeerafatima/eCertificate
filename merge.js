const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const crypto = require('crypto');
const sqlite3 = require('sqlite3').verbose();
const url = require('url');
const fs = require('fs');
const path = require('path');

const app = express();
const startLink = "http://localhost:3000/api?hash=";

// Set up multer to handle file uploads
const storage = multer.memoryStorage();
const upload = multer({ storage });

// Create SQLite database connection
const db = new sqlite3.Database('data.db');


// Create the "data" table if it doesn't exist
db.run(`
  CREATE TABLE IF NOT EXISTS data (
    Sno INTEGER,
    Rollno TEXT,
    Name TEXT,
    Domain TEXT,
    Title TEXT,
    Mentor TEXT,
    Duration TEXT,
    Completion DATE,
    hashValue TEXT
  )
`);

// Serve the index.html file
app.get('/', (req, res) => {
  res.sendFile(__dirname + '/index.html');
});


// Handle the file upload and conversion
app.post('/upload', upload.single('excelFile'), (req, res) => {
  const workbook = xlsx.read(req.file.buffer, { type: 'buffer', cellDates: true });
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const jsonData = xlsx.utils.sheet_to_json(worksheet, { raw: false });

  // Generate and add unique hash values (using sha256) as a new column
  jsonData.forEach((row) => {
    // console.log(row);

    // Convert and format the date value
    const rawDate = row['Date of Completion'];
    const completionDate = new Date(rawDate);
    const formattedDate = completionDate.toLocaleDateString('en-US', { day: 'numeric', month: 'short' });
    row['Date of Completion'] = formattedDate;

    // const combinedString = row['S.No'] + row['Roll Number'] + row['Name'] + row['Domain'] + row['Project Title'] + row['Mentor'] + row['Duration (months)'] + row['Date of Completion'];
    const { 'S.No': sno, 'Roll Number': rollno, 'Name': name, 'Domain': domain, 'Project Title': projectTitle, 'Mentor': mentor, 'Duration (months)': duration, 'Date of Completion': completion } = row;
    const combinedString = sno + rollno + name + domain + projectTitle + mentor + duration + completionDate;
    // console.log(combinedString);

    const hashValue = crypto.createHash('sha256').update(combinedString).digest('hex');
    row['link'] = startLink + hashValue;

    // const { sno, rollno, name, domain, projectTitle, mentor, duration, completion } = row;
    // console.log(sno, rollno, name, domain, projectTitle, mentor, duration, completion);

    // Check if the row with the same hash value already exists in the database
    db.get('SELECT * FROM data WHERE hashValue = ?', [hashValue], (err, result) => {
      if (err) {
        console.error(err);
        return;
      }

      if (!result) {
        // Insert the row data into the database if it doesn't exist
        db.run('INSERT INTO data (Sno, Rollno, Name, Domain, Title, Mentor, Duration, Completion, hashValue) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)',
          [sno, rollno, name, domain, projectTitle, mentor, duration, completion, hashValue], (err) => {
            if (err) {
              console.error(err);
              return;
            }
          });
      }
    });
  });

  // Convert JSON data back to an Excel workbook
  const updatedWorksheet = xlsx.utils.json_to_sheet(jsonData);
  const updatedWorkbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(updatedWorkbook, updatedWorksheet);

  // Set the content type and header for downloading the file
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="updated_excel.xlsx"');

  // Send the updated Excel file as a response
  res.send(xlsx.write(updatedWorkbook, { type: 'buffer' }));
});


// Define a route handler for GET requests to '/api'
app.get('/api', (req, res) => {
    const parsedUrl = url.parse(req.url, true);
    const hashValue = parsedUrl.query.hash;
  
    const sql = 'SELECT * FROM data WHERE hashValue = ?';
  
    db.get(sql, [hashValue], (err, row) => {
      if (err) {
        console.error(err);
        res.status(500).json({ message: 'Internal Server Error' });
      } else if (row) {
        const templatePath = path.join(__dirname, 'temp.html');
        fs.readFile(templatePath, 'utf8', (err, template) => {
          if (err) {
            console.error(err);
            res.status(500).json({ message: 'Internal Server Error' });
          } else {
            const renderedHtml = replacePlaceholders(template, row);
            res.status(200).send(renderedHtml);
          }
        });
      } else {
        res.status(404).json({ message: 'Data not found' });
      }
    });
  });
  
  // Define a route handler for all other routes
  app.use((req, res) => {
    res.status(404).json({ message: 'Route not found' });
  });
  
  // Helper function to replace placeholders in the HTML template with actual values
  // Helper function to replace placeholders in the HTML template with actual values
  function replacePlaceholders(template, data) {
      let renderedHtml = template;
      for (const [key, value] of Object.entries(data)) {
        const placeholder = `<%= ${key} %>`;
        renderedHtml = renderedHtml.replace(new RegExp(placeholder, 'g'), value);
      }
      return renderedHtml;
    }



// Start the server
const port = 3000;
app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
