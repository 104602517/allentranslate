const chineseFolder = './zh-CN';
const fsAsync = require('fs/promises');
const fs = require('fs')
const path = require('path');
const readLine = require('readline')
const ExcelJS = require('exceljs');

const split = (str, index) => {
  const result = [str.slice(0, index), str.slice(index)];

  return result;
}

const toEn = async () => {
  const res = await fsAsync.readdir(chineseFolder)
  const engExcel = new ExcelJS.Workbook()
  await engExcel.xlsx.readFile('./english.xlsx')
  const noFixWordBook = new ExcelJS.Workbook()
  const worksheet = noFixWordBook.addWorksheet('Sheet1');

  if (!fs.existsSync('en_US')) {
    fs.mkdirSync('en_US', { recursive: true });
  }
  
  res.forEach(async (file) => {
    
    if (path.extname(file) === '.ts') {
      const inputStream =  fs.createReadStream(`${chineseFolder}/${file}`)
      const writeStream = fs.createWriteStream(`./en_US/${file}`)
      const lineReader =  readLine.createInterface({input: inputStream})
      

      for await (const line of lineReader) {
        // 如果有冒號
        // 有 : { 代表是父層 不是key
        // 開頭是 " 代表不是key 是value的第二行
        if(line.includes(':') && (!line.includes(': {') || line.includes("'")) && line.trim()[0] !== '"') {
          const colonIndex = line.indexOf(':')
          const [left, right] = split(line, colonIndex + 1);
          let noFitKey = true
          engExcel.eachSheet(sheet => {
            let isFind = false
            sheet.eachRow(row => {
              row.eachCell((cell, index) => {
                if (typeof cell.value === 'string' && cell.value.trim() === left.replace(':', '').trim() && row.getCell(index + 2).value !== null && !isFind) {
                  console.log(left, row.getCell(index + 1).value, row.getCell(index + 2).value);
                  noFitKey = false
                  isFind = true
                  writeStream.write(`${left} "${row.getCell(index + 2).value}",\n`)
                } else if (typeof cell.value === 'string' && cell.value.trim() === left.replace(':', '').trim() && !isFind) {
                  writeStream.write(`${left} "${row.getCell(index + 1).value}",\n`)
                  noFitKey = false
                  isFind = true
                }
              })
            })
          })
          if (noFitKey) {
            worksheet.addRow([left, right]);
          }
        } else if (line.trim()[0] !== '"' && line.trim()[0] !== "'") {
          writeStream.write(line + '\n')
        }
      }
      noFixWordBook.xlsx.writeFile('noFix.xlsx')
        .then(() => {
            console.log('Excel檔案已成功建立並寫入欄位！');
        })
        .catch((error) => {
            console.error('建立Excel檔案時發生錯誤：', error);
        });
    }
  })
}

toEn()
