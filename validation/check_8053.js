const XLSX = require('xlsx');

// 读取Excel文件
const workbook = XLSX.readFile('asset/output_from_docx.xlsx');
const worksheet = workbook.Sheets['国民经济行业分类和代码'];
const data = XLSX.utils.sheet_to_json(worksheet);

// 查找805相关的行
console.log('查找805相关的行:');
for (let i = 1560; i < 1570; i++) {
  if (i < data.length) {
    const row = data[i];
    console.log(`行${i+2}: 门类=${row.门类}, 大类=${row.大类}, 中类=${row.中类}, 小类=${row.小类}, 类别名称=${row.类别名称}`);
  }
}

// 检查是否有8053
console.log('\n查找8053:');
for (let i = 0; i < data.length; i++) {
  const row = data[i];
  if (row.中类 && row.中类 === '805') {
    if (row.小类 && row.小类 === '8053') {
      console.log(`找到8053行${i+2}: 类别名称=${row.类别名称}`);
      break;
    }
  }
}

// 检查是否有8052
console.log('\n查找8052:');
for (let i = 0; i < data.length; i++) {
  const row = data[i];
  if (row.中类 && row.中类 === '805') {
    if (row.小类 && row.小类 === '8052') {
      console.log(`找到8052行${i+2}: 类别名称=${row.类别名称}`);
      break;
    }
  }
}
