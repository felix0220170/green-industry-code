const XLSX = require('xlsx');

// 读取Excel文件
const workbook = XLSX.readFile('asset/output_from_docx.xlsx');
const worksheet = workbook.Sheets['国民经济行业分类和代码'];
const data = XLSX.utils.sheet_to_json(worksheet);

// 查找B门类的第一行
console.log('开始检查B门类的第一行...');
let foundB = false;
for (let i = 0; i < data.length; i++) {
  const row = data[i];
  if (row.门类 === 'B') {
    foundB = true;
    console.log(`\nB门类开始:`);
    console.log(`行${i+2}: 门类=${row.门类}, 大类=${JSON.stringify(row.大类)}, 中类=${JSON.stringify(row.中类)}, 小类=${JSON.stringify(row.小类)}, 类别名称=${row.类别名称}`);
    
    // 显示B门类的前5行，以便观察
    console.log(`\nB门类的前5行:`);
    for (let j = 0; j < Math.min(5, data.length - i); j++) {
      const bRow = data[i + j];
      console.log(`行${i + j + 2}: 门类=${bRow.门类}, 大类=${JSON.stringify(bRow.大类)}, 中类=${JSON.stringify(bRow.中类)}, 小类=${JSON.stringify(bRow.小类)}, 类别名称=${bRow.类别名称}`);
    }
    break;
  }
}

if (!foundB) {
  console.log('未找到B门类');
}

// 检查是否有B门类包含05大类的情况
console.log('\n检查B门类中是否存在05大类或054中类...');
let foundIssue = false;
for (let i = 0; i < data.length; i++) {
  const row = data[i];
  if (row.门类 === 'B') {
    if (row.大类 === '05' || row.中类 === '054') {
      console.log(`发现问题行 ${i+2}: 门类=${row.门类}, 大类=${JSON.stringify(row.大类)}, 中类=${JSON.stringify(row.中类)}, 小类=${JSON.stringify(row.小类)}, 类别名称=${row.类别名称}`);
      foundIssue = true;
    }
  }
}

if (foundIssue) {
  console.log('\n存在问题！B门类中包含不应有的05大类或054中类。');
} else {
  console.log('\n未发现问题。B门类中没有05大类或054中类。');
}
