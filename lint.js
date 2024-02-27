const ExcelJS = require('exceljs');
const fs = require('fs');

async function removeColons() {
  // 创建一个新的工作簿
  const workbook = new ExcelJS.Workbook();
  
  // 读取 Excel 文件
  await workbook.xlsx.readFile('noFit.xlsx');
  
  // 获取第一个工作表
  const worksheet = workbook.getWorksheet(1);
  
  // 遍历第一列，删除所有冒号
  worksheet.eachRow((row) => {
    const cellValue = row.getCell(1).value;
    if (typeof cellValue === 'string') {
      row.getCell(1).value = cellValue.replace(/:/g, '');
    }
  });

  worksheet.eachRow((row) => {
    const cellValue = row.getCell(3).value;
    if (typeof cellValue === 'string') {
      row.getCell(3).value = cellValue.slice(2, cellValue.length-2)
    }
  });
  
  // 保存修改后的工作簿到新的文件
  await workbook.xlsx.writeFile('noFit_updated.xlsx');
  
  console.log('冒号已删除并保存到 nofix_updated.xlsx 文件中。');
}

removeColons();


