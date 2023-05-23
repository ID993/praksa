const express = require('express');
const ExcelJS = require('exceljs');

const app = express();

app.get('/api/excel-data', async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('./data.xlsx');

    const worksheet = workbook.getWorksheet('List1');

    const headers = [];
    const data = [];


    worksheet.getRow(1).eachCell((cell) => {
      headers.push(cell.value);
    });

    worksheet.eachRow({ startingRow: 2 }, (row) => {
      const rowData = {};
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const header = headers[colNumber - 1];
        rowData[header] = cell.value;
      });
      data.push(rowData);
    });

    data.shift();

    res.json(data);
  } catch (error) {
    console.error('Error reading Excel file:', error);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

app.listen(3000, () => {
  console.log('API server is running on port 3000');
});
