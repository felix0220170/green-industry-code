const XLSX = require('xlsx');
const workbook = XLSX.readFile('asset/output_from_docx.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(worksheet);

console.log('寻找A门类结束和B门类开始的部分...');
let foundB = false;

for (let i = 0; i < Math.min(200, data.length); i++) {
  const row = data[i];
  
  // 只显示A门类最后几行和B门类开始几行
  if (row.门类 === 'A' && i > 100) {
    console.log(`行${i+2}: 门类=${row.门类}, 大类=${row.大类}, 中类=${row.中类}, 小类=${row.小类}, 类别名称=${row.类别名称}`);
  }
  
  if (row.门类 === 'B') {
    foundB = true;
    console.log(`\nB门类开始:`);
    console.log(`行${i+2}: 门类=${row.门类}, 大类=${JSON.stringify(row.大类)}, 中类=${JSON.stringify(row.中类)}, 小类=${JSON.stringify(row.小类)}, 类别名称=${row.类别名称}`);
    
    // 显示B门类的前5行
    for (let j = 0; j < Math.min(5, data.length - i); j++) {
      const bRow = data[i + j];
      console.log(`行${i+2+j}: 门类=${bRow.门类}, 大类=${JSON.stringify(bRow.大类)}, 中类=${JSON.stringify(bRow.中类)}, 小类=${JSON.stringify(bRow.小类)}, 类别名称=${bRow.类别名称}`);
    }
    
    break;
  }
}
