const express = require('express');
const bodyParser = require('body-parser');
const Excel = require('exceljs');
const mongoose = require('mongoose');
const db = require('./db');
const fs = require('fs');
const path = require('path');
const User = require('./User'); 

const app = express();
const PORT = 3000;
const filePath = path.join(__dirname, 'data.xlsx');

if (!fs.existsSync(filePath)) {
  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');
  worksheet.addRow(['id', 'name', 'email']);
  workbook.xlsx.writeFile(filePath);
}


app.use(bodyParser.json());


app.post('/users', async (req, res) => {
  try {
    const newUser = req.body;
    const user = new User(newUser);
    await user.save();

    
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(1);

    worksheet.addRow([user._id, user.name, user.email]);

    await workbook.xlsx.writeFile(filePath);

    res.status(201).json(user);
  } catch (err) {
    console.error('Error creating user:', err);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});


app.use((err, req, res, next) => {
  console.error('Error occurred:', err);
  res.status(500).json({ error: 'Internal Server Error' });
});

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
