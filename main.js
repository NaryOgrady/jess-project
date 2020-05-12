const Excel = require('exceljs');

const hearingsFileName = 'data/2018-2019_Hearings_Check.xlsx';
const generatedFileName = 'data/2018-2019_Hearings_Generated.xlsx';
const worksheetName = '2018-2019';
const dataColumn = 'A';
const typesRegex = [/(\w+\s){0,2}MCH/, /(\w+\s){0,2}ICH/, /BOND HEARING/];
const namesRegex = /(\w+\s)+(?=\s*\()/;
const alienNumberRegex = /(?<=\(A\s)(\d{3}(\s|-)){2}\d{3}/;

async function getHearingsData(filename) {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(filename);
  const sheet = workbook.getWorksheet(worksheetName);
  return sheet.getColumn(dataColumn).values;
}

function getHearingType(data) {
  let type = '';
  for (const re of typesRegex) {
    const matches = re.exec(data);
    if (matches) {
      type = matches[0];
    }
  }
  return type;
}

function getHearingName(data) {
  let name = '';
  const matches = namesRegex.exec(data);
  if (matches) {
    name = matches[0];
  }
  return name;
}

function getAlienNumber(data) {
  let alienNumber = '';
  const matches = alienNumberRegex.exec(data);
  if (matches) {
    alienNumber = matches[0];
  }
  return alienNumber.replace('-', ' ');
}

function extractInfo(data) {
  if (!data) {
    return;
  }
  const type = getHearingType(data);
  if (!type) {
    return;
  }
  data = data.split(type)[1].trim();
  const name = getHearingName(data);
  const alienNumber = getAlienNumber(data);
  return {type, name, alienNumber};
}

function markCell(cell) {
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFF7D7D' },
    bgColor: { argb: 'FF000000' }
  }
}

async function writeExcel(hearingsInfo) {
  const workbook = new Excel.Workbook();
  const sheet = workbook.addWorksheet(worksheetName);
  sheet.addRow(['Type', 'Name', 'Alien Number']);
  for (const info of hearingsInfo) {
    sheet.addRow([info.type, info.name, info.alienNumber]);
  }
  let column = sheet.getColumn('B');
  column.eachCell({includeEmpty: true}, (cell, rowNumber) => {
    if (!cell.value || cell.value.length < 10) {
      markCell(cell);
    }
  });
  column = sheet.getColumn('C');
  column.eachCell({includeEmpty: true}, (cell, rowNumber) => {
    if (!cell.value || cell.value.length < 6) {
      markCell(cell);
    }
  });
  await workbook.xlsx.writeFile(generatedFileName);
}

async function generateExcel() {
  const hearingsInfo = []
  const hearingsData = await getHearingsData(hearingsFileName);
  for (const data of hearingsData) {
    const info = extractInfo(data);
    if (info) {
      hearingsInfo.push(info);
    }
  }
  await writeExcel(hearingsInfo);
  console.log('File Generated');
}

generateExcel();
