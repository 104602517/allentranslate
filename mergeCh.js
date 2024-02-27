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

let keyValueArr = []

const mergeCh = async () => {
  const res = await fsAsync.readdir(chineseFolder)
  let i = 0
  await Promise.all(
    res.map(async (file) => {
      // 檢查是不是都是ts檔案
      if (path.extname(file) === '.ts') {
        const inputStream =  fs.createReadStream(`${chineseFolder}/${file}`)
        const lineReader =  readLine.createInterface({input: inputStream})
  
        for await (const line of lineReader) {
          // 如果有冒號
          // 有 : { 代表是父層 不是key
          // 開頭是 " 代表不是key 是value的第二行
          if(line.includes(':') && (!line.includes(': {') || line.includes("'")) && line.trim()[0] !== '"') {
            const colonIndex = line.indexOf(':')
            const [left, right] = split(line, colonIndex + 1);
            keyValueArr.push({
              key: left.replace(':', '').trim(),
              // 把結尾的逗號刪掉
              value: right.trim().slice(0, right.trim().length - 1)
            })
            i++
          } else if (line.trim()[0] === '"' || line.trim()[0] === "'") {
            keyValueArr[i - 1].value = line
          }
        }
      }
    })
  )
  // 接下來把資料寫進表格
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('BB新市場後台翻譯');
  worksheet.addRow(['Key', 'Chinese', 'English']);
  keyValueArr.forEach(item => {
    worksheet.addRow([item.key, item.value]);
  });
  try {
    await workbook.xlsx.writeFile('BB新市場後台翻譯.xlsx')
    console.log('excel輸出完成')
  } catch(err) {
    console.log(err)
  }
}

mergeCh()

