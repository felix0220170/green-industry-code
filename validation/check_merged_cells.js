const XLSX = require('xlsx');

// 读取Excel文件
const workbook = XLSX.readFile('asset/output_from_docx.xlsx');
const worksheet = workbook.Sheets['国民经济行业分类和代码'];
const data = XLSX.utils.sheet_to_json(worksheet);

// 查找有问题的行
console.log('开始检查有问题的行...');

// 检查277278相关行
console.log('\n查找277278相关行:');
for (let i = 0; i < data.length; i++) {
  const row = data[i];
  if (row.中类 && row.中类.includes('277278')) {
    console.log(`行${i+2}: 门类=${row.门类}, 大类=${row.大类}, 中类=${row.中类}, 小类=${row.小类}, 类别名称=${row.类别名称}`);
    // 显示前后几行以便上下文
    for (let j = Math.max(0, i-2); j <= Math.min(data.length-1, i+2); j++) {
      const contextRow = data[j];
      console.log(`  上下文行${j+2}: 门类=${contextRow.门类}, 大类=${contextRow.大类}, 中类=${contextRow.中类}, 小类=${contextRow.小类}, 类别名称=${contextRow.类别名称}`);
    }
    break;
  }
}

// 检查80528053相关行
console.log('\n查找80528053相关行:');
for (let i = 0; i < data.length; i++) {
  const row = data[i];
  if (row.小类 && row.小类.includes('80528053')) {
    console.log(`行${i+2}: 门类=${row.门类}, 大类=${row.大类}, 中类=${row.中类}, 小类=${row.小类}, 类别名称=${row.类别名称}`);
    // 显示前后几行以便上下文
    for (let j = Math.max(0, i-2); j <= Math.min(data.length-1, i+2); j++) {
      const contextRow = data[j];
      console.log(`  上下文行${j+2}: 门类=${contextRow.门类}, 大类=${contextRow.大类}, 中类=${contextRow.中类}, 小类=${contextRow.小类}, 类别名称=${contextRow.类别名称}`);
    }
    break;
  }
}

// 检查其他可能的合并单元格问题
console.log('\n查找其他可能的合并单元格问题:');
let issueCount = 0;
for (let i = 0; i < data.length; i++) {
  const row = data[i];
  
  // 检查代码是否看起来像是合并的
  if (row.小类 && row.小类.length > 4) {
    console.log(`可能的问题行${i+2}: 小类=${row.小类}, 名称=${row.类别名称}`);
    issueCount++;
  }
  if (row.中类 && row.中类.length > 3) {
    console.log(`可能的问题行${i+2}: 中类=${row.中类}, 名称=${row.类别名称}`);
    issueCount++;
  }
  
  if (issueCount >= 10) break; // 只显示前10个问题
}
