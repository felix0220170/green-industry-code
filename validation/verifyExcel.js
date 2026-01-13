const XLSX = require('xlsx');
const fs = require('fs');

// 检查文件是否存在
const filePath = 'output_from_docx.xlsx';
if (!fs.existsSync(filePath)) {
  console.error('文件不存在:', filePath);
  process.exit(1);
}

// 读取Excel文件
const workbook = XLSX.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// 将工作表转换为JSON
const data = XLSX.utils.sheet_to_json(worksheet);

console.log('Excel文件验证信息:');
console.log('- 工作表名称:', sheetName);
console.log('- 数据行数:', data.length);
console.log('- 列名:', Object.keys(data[0] || {}));

// 输出前5行数据
console.log('\n前5行数据示例:');
for (let i = 0; i < Math.min(5, data.length); i++) {
  console.log('行', i + 1, ':', data[i]);
}

// 检查数据完整性
if (data.length > 0) {
  console.log('\n数据完整性检查:');
  const requiredColumns = ['门类', '大类', '中类', '小类', '类别名称', '说明'];
  const missingColumns = requiredColumns.filter(col => !Object.keys(data[0]).includes(col));
  if (missingColumns.length === 0) {
    console.log('- 所有必填列都存在');
  } else {
    console.log('- 缺失列:', missingColumns);
  }
}
