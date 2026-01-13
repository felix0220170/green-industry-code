const XLSX = require('xlsx');
const fs = require('fs');

// 读取Excel文件
const workbook = XLSX.readFile('output_from_docx.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// 将Excel转换为JSON
const jsonData = XLSX.utils.sheet_to_json(worksheet);

console.log('Excel文件基本信息:');
console.log('- 总行数:', jsonData.length);
console.log('- 列名:', Object.keys(jsonData[0] || {}));

// 查找所有门类
const sections = new Set();
const sectionRows = [];

console.log('\n所有门类信息:');
jsonData.forEach((row, index) => {
  const section = row.门类;
  const category = row.大类;
  const subcategory = row.中类;
  const item = row.小类;
  const name = row.类别名称;
  
  if (section && !category && !subcategory && !item) {
    sections.add(section);
    sectionRows.push({
      index: index + 2, // 行号从2开始
      section: section,
      name: name,
      desc: row.说明
    });
  }
});

console.log('- 门类别:', Array.from(sections));
console.log('\n详细的门类信息:');
sectionRows.forEach(row => {
  console.log('行', row.index, ':', '门类代码:', row.section, '名称:', row.name, '说明:', row.desc);
});

// 输出前5行数据
console.log('\n前5行数据示例:');
for (let i = 0; i < Math.min(5, jsonData.length); i++) {
  console.log('行', i + 2, ':', jsonData[i]);
}

// 查找门类B的信息
console.log('\n门类B的信息:');
for (let i = 0; i < jsonData.length; i++) {
  const row = jsonData[i];
  if (row.门类 === 'B') {
    console.log('行', i + 2, ':', row);
  }
}
