const ExcelJS = require('exceljs');
const fs = require('fs');


function applyRichTextFormatting(worksheet, cellRef, boldText) {
  const cell = worksheet.getCell(cellRef);
  const parts = [
    { text: boldText, font: { bold: true, size: 14, horizontal: 'center' } },
    { text: cell.value.toString().substring(boldText.length) }
  ];
  cell.value = { richText: parts };
}

function applyRowBorderToCells(worksheet, rowIndex, startColumn, endColumn, borderStyle) {
  for (let col = startColumn; col <= endColumn; col++) {
    const cell = worksheet.getCell(`${String.fromCharCode(64 + col)}${rowIndex}`);
    cell.border = {
      top: { style: borderStyle },
      bottom: { style: borderStyle },
      right: { style: borderStyle },
      left: { style: borderStyle } };
    cell.font = {
      name: 'Calibri',
      bold: true,
      size: 11 };
    cell.alignment = {
      vertical: 'middle',
      horizontal: 'center',
      wrapText: true };
  }
}

function applyOneRowBorderToCells(worksheet, rowIndex, startColumn, endColumn, borderStyle) {
  for (let col = startColumn; col <= endColumn; col++) {
    const cell = worksheet.getCell(`${String.fromCharCode(64 + col)}${rowIndex}`);
    if (col == endColumn) {
      cell.border = {
          left: { style: borderStyle },
          bottom: { style: 'medium' },
          right: { style: 'medium' } };
    } else {
    cell.border = {
      right: { style: borderStyle },
      left: { style: borderStyle },
      bottom: { style: 'medium' } };
    }
  }
}

function applyBorderToCells(worksheet, startRow, endRow, startColumn, endColumn, borderStyle) {
  for (let row = startRow; row <= endRow; row++) {
    for (let col = startColumn; col <= endColumn; col++) {
      const cell = worksheet.getCell(`${String.fromCharCode(64 + col)}${row}`);
      if (col == endColumn) {
        cell.border = {
            top: { style: borderStyle },
            bottom: { style: borderStyle },
            right: { style: 'medium' },
            left: { style: borderStyle } };
      } else {
        cell.border = {
          top: { style: borderStyle },
          bottom: { style: borderStyle },
          right: { style: borderStyle },
          left: { style: borderStyle } };
      }
    }
  }
}

function applyRowColorToCells(worksheet, rowIndex, startColumn, endColumn) {
  for (let col = startColumn; col <= endColumn; col++) {
    const cell = worksheet.getCell(`${String.fromCharCode(64 + col)}${rowIndex}`);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'DDDDDD' } };
  }
}

function setNumFormat(worksheet, startRow, endRow, startCol, endCol, format) {
  for (let i = startRow; i <= endRow; i++) {
    for (let j = startCol; j <= endCol; j++) {
      worksheet.getCell(`${String.fromCharCode(64 + j)}${i}`).numFmt = format;
    }
  }
}

async function generateExcelTable() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Table 1');

  const logoImage = workbook.addImage({
    filename: 'logo.png', 
    extension: 'png',
  });

  worksheet.addImage(logoImage, {
    tl: { col: 0, row: 0 }, 
    br: { col: 2, row: 4 }, 
  });

    worksheet.getRow(5).height = 15;
    worksheet.getRow(12).height = 49.5;
    worksheet.getRow(15).height = 15;
    worksheet.getRow(16).height = 70.5;
    
  
    worksheet.getColumn('A').width = 5.89;
    worksheet.getColumn('B').width = 18.33;
    worksheet.getColumn('C').width = 21.11;
    worksheet.getColumn('D').width = 21.11;
    worksheet.getColumn('E').width = 6.11;
    worksheet.getColumn('F').width = 7.78;
    worksheet.getColumn('G').width = 7.11;
    worksheet.getColumn('H').width = 10.11;
    worksheet.getColumn('I').width = 9.89;
    worksheet.getColumn('J').width = 10.11;
    worksheet.getColumn('K').width = 8.11;
    worksheet.getColumn('L').width = 8.11;
    worksheet.getColumn('M').width = 8.11;
    worksheet.getColumn('N').width = 8.11;

    worksheet.mergeCells('A6:I11');
    worksheet.mergeCells('A12:B12');
    worksheet.mergeCells('A13:B13');
    worksheet.mergeCells('H12:I12');
    worksheet.mergeCells('H13:I13');
    worksheet.mergeCells('J34:L35');
    worksheet.mergeCells('A28:C29');
    worksheet.mergeCells('A34:C35');
    worksheet.mergeCells('A15:A16');
    worksheet.mergeCells('B15:B16');
    worksheet.mergeCells('C15:C16');
    worksheet.mergeCells('D15:D16');
    worksheet.mergeCells('H15:H16');
    worksheet.mergeCells('I15:I16');
    worksheet.mergeCells('J15:J16');
    worksheet.mergeCells('E15:G15');
    worksheet.mergeCells('K15:M15');
    worksheet.mergeCells('N15:N16');
    worksheet.mergeCells('A25:C25');

    worksheet.getCell('A5').value = 'Predmet: ';
    worksheet.getCell('A6').alignment = { vertical: 'middle', wrapText: true };
    worksheet.getCell('A6').value = 'NALOG ZA ISPLATU\nLorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.';
    worksheet.getCell('A28').value = 'Prodekanica za nastavu i studentska pitanja\nProf. dr. sc.: Ime Prezime';
    worksheet.getCell('A34').value = 'Prodekan za financije i upravljanje\nProf. dr. sc. Ime Prezime';
    worksheet.getCell('J34').value = 'Dekan\nProf. dr. sc. Ime Prezime';
    worksheet.getCell('A28').alignment = { vertical: 'middle', wrapText: true };
    worksheet.getCell('A34').alignment = { vertical: 'middle', wrapText: true };
    worksheet.getCell('J34').alignment = { vertical: 'middle', wrapText: true };

    applyRichTextFormatting(worksheet, 'A6', 'NALOG ZA ISPLATU');
  
    applyRowBorderToCells(worksheet, 12, 1, 9, 'medium');
    applyRowBorderToCells(worksheet, 13, 1, 9, 'medium');
    applyOneRowBorderToCells(worksheet, 13, 1, 9, 'thin');
    applyRowBorderToCells(worksheet, 15, 1, 14, 'medium');
    applyRowBorderToCells(worksheet, 16, 1, 14, 'medium');
    applyBorderToCells(worksheet, 17, 24, 1, 14, 'thin');
    applyRowColorToCells(worksheet, 12, 1, 9);
    applyRowColorToCells(worksheet, 15, 1, 14);
    applyRowColorToCells(worksheet, 16, 1, 14);
    applyRowBorderToCells(worksheet, 25, 1, 14, 'medium');

    worksheet.getCell('A12').value = 'Katedra';
    worksheet.getCell('C12').value = 'Studij';
    worksheet.getCell('D12').value = 'ak. god.';
    worksheet.getCell('E12').value = 'stud. god.';
    worksheet.getCell('F12').value = 'početak turnusa';
    worksheet.getCell('G12').value = 'kraj turnusa';
    worksheet.getCell('H12').value = 'br sati predviđen programom';
    worksheet.getCell('A15').value = 'Redni broj';
    worksheet.getCell('B15').value = 'Nastavnik/Suradnik';
    worksheet.getCell('C15').value = 'Zvanje';
    worksheet.getCell('D15').value = 'Status';
    worksheet.getCell('E15').value = 'Sati nastave';
    worksheet.getCell('E16').value = 'pred';
    worksheet.getCell('F16').value = 'sem';
    worksheet.getCell('G16').value = 'vjež';
    worksheet.getCell('H15').value = 'Bruto satnica predavanja (EUR)';
    worksheet.getCell('I15').value = 'Bruto satnica seminari (EUR)';
    worksheet.getCell('J15').value = 'Bruto satnica vježbe (EUR)';
    worksheet.getCell('K15').value = 'Bruto iznos';
    worksheet.getCell('K16').value = 'pred';
    worksheet.getCell('L16').value = 'sem';
    worksheet.getCell('M16').value = 'vjež';
    worksheet.getCell('N15').value = 'Ukupno za isplatu (EUR)';
    worksheet.getCell('A25').value = 'UKUPNO';

    worksheet.getRow(13).eachCell({ includeEmpty: true }, (cell) => {
      cell.font = {
        name: 'Calibri',
        bold: false,
        size: 11 };
      });

    const data = await readSecondTable();
      
    data.forEach((row, index) => {
      worksheet.getCell('A5').value = 'Predmet: ' + row['PredmetNaziv'];
      worksheet.getCell('A13').value = row['Katedra'];
      worksheet.getCell('C13').value = row['Studij'];
      worksheet.getCell('D13').value = row['SkolskaGodinaNaziv'];
      worksheet.getCell(`A${index + 17}`).value = index + 1;
      worksheet.getCell(`B${index + 17}`).value = row['NastavnikSuradnikNaziv'];
      worksheet.getCell(`C${index + 17}`).value = row['ZvanjeNaziv'];
      worksheet.getCell(`D${index + 17}`).value = row['NazivNastavnikStatus'];
      worksheet.getCell(`E${index + 17}`).value = row['PlaniraniSatiPredavanja'];
      worksheet.getCell(`F${index + 17}`).value = row['PlaniraniSatiSeminari'];
      worksheet.getCell(`G${index + 17}`).value = row['PlaniraniSatiVjezbe'];
      worksheet.getCell(`H${index + 17}`).value = row['NormaPlaniraniSatiPredavanja'];
      worksheet.getCell(`I${index + 17}`).value = row['NormaPlaniraniSatiSeminari'];
      worksheet.getCell(`J${index + 17}`).value = row['NormaPlaniraniSatiVjezbe'];
      worksheet.getCell(`K${index + 17}`).value = row['RealiziraniSatiPredavanja'];
      worksheet.getCell(`L${index + 17}`).value = row['RealiziraniSatiSeminari'];
      worksheet.getCell(`M${index + 17}`).value = row['RealiziraniSatiVjezbe'];
      worksheet.getCell(`N${index + 17}`).value = row['RealiziraniSatiPredavanja'] + row['RealiziraniSatiSeminari'] + row['RealiziraniSatiVjezbe'];
      worksheet.getCell('E25').value = { formula: 'SUM(E17:E24)'};
      worksheet.getCell('F25').value = { formula: 'SUM(F17:F24)'};
      worksheet.getCell('G25').value = { formula: 'SUM(G17:G24)'};
      worksheet.getCell('K25').value = { formula: 'SUM(E17:E24)'};
      worksheet.getCell('L25').value = { formula: 'SUM(F17:F24)'};
      worksheet.getCell('M25').value = { formula: 'SUM(G17:G24)'};
      worksheet.getCell('N25').value = { formula: 'SUM(E17:E24)'};

      setNumFormat(worksheet, 25, 25, 5, 7, '0.00');
      setNumFormat(worksheet, 17, 25, 11, 14, '0.00');

    });

    await workbook.xlsx.writeFile('projectTwo.xlsx');
      console.log('Excel file generated!');
}

async function readSecondTable() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('./data.xlsx');
  const worksheet = workbook.getWorksheet('List1');

  const data = [];

  worksheet.eachRow((row, rowIndex) => {
    if (rowIndex > 1) {
      const rowData = {};

      row.eachCell((cell, cellIndex) => {
        const headerCell = worksheet.getRow(1).getCell(cellIndex);
        const header = headerCell.value;

        rowData[header] = cell.value;
      });

      data.push(rowData);
    }
  });

  return data;
}


generateExcelTable().catch((error) => {
  console.log('An error occurred:', error);
});
readSecondTable().catch((error) => {
  console.log('An error occurred:', error);
});