const xl = require('excel4node');
const fs = require('fs');
const path = require('path');

const URL_SOURCE_JSON = path.resolve(__dirname, 'vi.json');

const isExistedFile = (path) => {
  if (fs.existsSync(path)) {
    return true;
  }
  return false;
}

const handleReadFile = () => {
  if (!isExistedFile(URL_SOURCE_JSON)) return {};

  const getFile = fs.readFileSync(URL_SOURCE_JSON);
  const rawData = !!getFile && JSON.parse(getFile);

  return rawData;
}

// const handleWriteFile = (data) => {
//   return fs.writeFileSync(URL_RESPONSE_JSON, JSON.stringify(data));
// }

const init = () => {
  const data = handleReadFile(URL_SOURCE_JSON);

  if (!data) {
    return console.log('Source file not found or empty!!!');
  }

  let wb = new xl.Workbook({
    defaultFont: {
      size: 12,
      name: 'Times New Roman'
    },
  });
  let ws = wb.addWorksheet('Sheet1');

  let border = { left: 'thin', right: 'thin', top: 'thin', bottom: 'thin' };
  // ws.row(1).setHeight(30);
  // ws.column(1).setWidth(5);
  // ws.column(4).freeze();
  // ws.row(7).freeze();
  let row = 1, col = 1;

  function renderCell(data) {
    if (typeof data == 'string') {
      ws.cell(row, col).string(data);
    } else {
      Object.keys(data).forEach(keyChild => {
        renderCell(data[keyChild]);
      })
    }
  }

  Object.keys(data).forEach(key => {
    col = 1;
    renderCell(key);
    col++;
    renderCell(data[key]);
    row++;
  })

  function getExcelCellRef(row, col) {
    return xl.getExcelCellRef(row, col);
  }

  wb.write('output.xlsx', (err) => {
    if (err) {
      console.error(err);
    } else {
      console.log("TaiTH ~ Done");
    }
  });
}

init();