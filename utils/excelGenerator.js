const ExcelJS = require('exceljs');

function generateStyledTimetableExcel({ timetable, subjectTeachers, university, faculty, wefDate, slots, days }) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Timetable', { views: [{ state: 'frozen', ySplit: 0 }] });

  const deepHeaderStyle = {
    font: { bold: true, size: 14, color: { argb: 'FFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '0F044C' } },
    alignment: { vertical: 'middle', horizontal: 'center' }
  };

  const lightHeaderStyle = {
    font: { bold: true, color: { argb: '000000' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F8D7DA' } },
    alignment: { vertical: 'middle', horizontal: 'center' }
  };

  const labCellStyle = {
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'CBDCEB' } },
    alignment: { vertical: 'middle', horizontal: 'center', wrapText: true }
  };

  const borderStyle = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };

  let rowCursor = 1;

  for (const [courseName, courseData] of Object.entries(timetable)) {
    const totalCols = slots.length + 1;
    const courseStartRow = rowCursor;

    sheet.mergeCells(rowCursor, 1, rowCursor, totalCols);
    const uniCell = sheet.getCell(rowCursor, 1);
    uniCell.value = university;
    uniCell.style = deepHeaderStyle;
    rowCursor++;

    const formattedDate = !isNaN(new Date(wefDate)) ? new Date(wefDate).toLocaleDateString('en-GB') : wefDate;
    const part1 = `Course: ${courseName}`;
    const part2 = `Faculty: ${faculty}`;
    const part3 = `W.E.F: ${formattedDate}`;

    const firstPartCols = Math.floor(totalCols / 3);
    const secondPartCols = Math.floor(totalCols / 3);
    const thirdPartCols = totalCols - firstPartCols - secondPartCols;

    sheet.mergeCells(rowCursor, 1, rowCursor, firstPartCols);
    sheet.mergeCells(rowCursor, firstPartCols + 1, rowCursor, firstPartCols + secondPartCols);
    sheet.mergeCells(rowCursor, firstPartCols + secondPartCols + 1, rowCursor, totalCols);

    sheet.getCell(rowCursor, 1).value = part1;
    sheet.getCell(rowCursor, firstPartCols + 1).value = part2;
    sheet.getCell(rowCursor, firstPartCols + secondPartCols + 1).value = part3;

    [1, firstPartCols + 1, firstPartCols + secondPartCols + 1].forEach(col => {
      sheet.getCell(rowCursor, col).style = lightHeaderStyle;
    });

    rowCursor++;

    const headerRow = sheet.getRow(rowCursor);
    headerRow.getCell(1).value = 'Day / Time';
    for (let i = 0; i < slots.length; i++) {
      headerRow.getCell(i + 2).value = slots[i];
    }
    headerRow.eachCell(cell => {
      cell.style = lightHeaderStyle;
      cell.border = borderStyle;
    });
    sheet.getRow(rowCursor).height = 25;
    rowCursor++;

    for (const day of days) {
      const row = sheet.getRow(rowCursor);
      row.getCell(1).value = day;
      row.getCell(1).style = lightHeaderStyle;
      row.getCell(1).border = borderStyle;

      let colCursor = 2;
      for (let i = 0; i < slots.length;) {
        const entry = courseData[day]?.[i];
        if (!entry) {
          row.getCell(colCursor).value = '';
          row.getCell(colCursor).border = borderStyle;
          row.getCell(colCursor).alignment = { vertical: 'middle', horizontal: 'center' };
          colCursor++;
          i++;
          continue;
        }

        const teacherDisplay = entry.teacher || (entry.subject?.toLowerCase().includes("lunch") ? '' : 'TBA');
        const content = `${entry.room || ''}\n${entry.subject || ''}\n${teacherDisplay}`;

        if (entry.subject?.toLowerCase().includes('lab')) {
          let colspan = 1;
          while (
            i + colspan < slots.length &&
            courseData[day]?.[i + colspan]?.subject === entry.subject
          ) {
            colspan++;
          }

          sheet.mergeCells(rowCursor, colCursor, rowCursor, colCursor + colspan - 1);
          const cell = sheet.getCell(rowCursor, colCursor);
          cell.value = content;
          cell.style = labCellStyle;
          cell.border = borderStyle;
          colCursor += colspan;
          i += colspan;
        } else {
          const cell = row.getCell(colCursor);
          cell.value = content;
          cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
          cell.border = borderStyle;
          colCursor++;
          i++;
        }
      }
      row.height = 60;
      rowCursor++;
    }

    rowCursor++;
    sheet.mergeCells(rowCursor, 1, rowCursor, totalCols);
    const mapTitle = sheet.getCell(rowCursor, 1);
    mapTitle.value = `Subject and Teacher Mapping for ${courseName}`;
    mapTitle.style = deepHeaderStyle;
    rowCursor++;

    // New Header for Mapping
    const mapHeader = sheet.getRow(rowCursor);
    mapHeader.getCell(1).value = 'Subject Short';
    sheet.mergeCells(rowCursor, 2, rowCursor, 4);
    mapHeader.getCell(2).value = 'Subject Full Name';
    mapHeader.getCell(5).value = 'Teacher Short';
    sheet.mergeCells(rowCursor, 6, rowCursor, 8);
    mapHeader.getCell(6).value = 'Teacher Full Name';

    [1, 2, 5, 6].forEach(col => {
      const cell = sheet.getCell(rowCursor, col);
      cell.style = lightHeaderStyle;
      cell.border = borderStyle;
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    rowCursor++;

    const mappings = subjectTeachers[courseName] || [];
    for (const m of mappings) {
      const row = sheet.getRow(rowCursor);
      row.getCell(1).value = m.subjectShort || '';
      sheet.mergeCells(rowCursor, 2, rowCursor, 4);
      sheet.getCell(rowCursor, 2).value = m.subjectLong || '';
      row.getCell(5).value = m.teacherShort || '';
      sheet.mergeCells(rowCursor, 6, rowCursor, 8);
      sheet.getCell(rowCursor, 6).value = m.teacherLong || '';

      [1, 2, 5, 6].forEach(col => {
        const cell = sheet.getCell(rowCursor, col);
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.border = borderStyle;
      });

      row.height = 20;
      rowCursor++;
    }

    const courseEndRow = rowCursor;

    // Apply border to entire block
    for (let i = courseStartRow; i <= courseEndRow; i++) {
      for (let j = 1; j <= totalCols; j++) {
        const cell = sheet.getCell(i, j);
        if (!cell.border) cell.border = borderStyle;
      }
    }

    rowCursor += 2;
  }

  for (let i = 1; i <= slots.length + 1; i++) {
    sheet.getColumn(i).width = 18;
  }

  return workbook;
}

module.exports = generateStyledTimetableExcel;
