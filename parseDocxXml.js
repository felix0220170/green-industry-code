const AdmZip = require('adm-zip');
const fs = require('fs');
const XLSX = require('xlsx');

// 读取并解压docx文件
const zip = new AdmZip('source.docx');
const zipEntries = zip.getEntries();

// 查找document.xml文件
let documentXml = null;
for (const entry of zipEntries) {
  if (entry.entryName === 'word/document.xml') {
    documentXml = entry.getData().toString('utf8');
    break;
  }
}

if (!documentXml) {
  console.error('未找到word/document.xml文件');
  process.exit(1);
}

// 将XML转换为可解析的格式（简单处理）
// 这里只是简单的处理，实际XML解析应该使用专门的XML库
let tableContent = documentXml;

// 提取所有表格
const tables = tableContent.match(/<w:tbl[^>]*>([\s\S]*?)<\/w:tbl>/g) || [];
console.log('找到', tables.length, '个表格');

// 查找包含"国民经济行业分类和代码"的表格
let targetTable = null;
for (const table of tables) {
  // 检查表格是否包含关键词
  if (table.includes('国民经济行业分类和代码')) {
    targetTable = table;
    break;
  }
}

if (!targetTable) {
  console.error('未找到包含"国民经济行业分类和代码"的表格');
  process.exit(1);
}

// 提取表格中的行
const rows = targetTable.match(/<w:tr[^>]*>([\s\S]*?)<\/w:tr>/g) || [];
console.log('目标表格包含', rows.length, '行');

// 解析每行的单元格
const parsedTable = [];
rows.forEach(row => {
  // 提取单元格
  const cells = row.match(/<w:tc[^>]*>([\s\S]*?)<\/w:tc>/g) || [];
  
  const parsedRow = [];
  cells.forEach(cell => {
    // 提取单元格文本
    const texts = cell.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g) || [];
    let cellText = '';
    
    texts.forEach(text => {
      // 移除标签，保留文本内容
      cellText += text.replace(/<\/w:t>/g, '').replace(/<w:t[^>]*>/g, '');
    });
    
    parsedRow.push(cellText);
  });
  
  parsedTable.push(parsedRow);
});

// 处理表格数据
const excelData = [];

// 查找表头行
let headerRowIndex = -1;
for (let i = 0; i < parsedTable.length; i++) {
  const row = parsedTable[i];
  if (row.some(cell => cell.includes('门类')) && row.some(cell => cell.includes('大类'))) {
    headerRowIndex = i;
    break;
  }
}

if (headerRowIndex === -1) {
  console.error('未找到表头行');
  process.exit(1);
}

// 从数据行开始处理
for (let i = headerRowIndex + 1; i < parsedTable.length; i++) {
  const row = parsedTable[i];
  
  // 确保行有足够的单元格
  if (row.length >= 6) {
    const entry = {
      门类: row[0] || '',
      大类: row[1] || '',
      中类: row[2] || '',
      小类: row[3] || '',
      类别名称: row[4] || '',
      说明: row[5] || ''
    };
    
    // 只添加有实际内容的行
    if (Object.values(entry).some(value => value !== '')) {
      excelData.push(entry);
    }
  }
}

console.log('表格数据处理完成，共提取', excelData.length, '条有效记录');

// 创建Excel工作簿
const workbook = XLSX.utils.book_new();
const worksheet = XLSX.utils.json_to_sheet(excelData);

// 设置列名顺序
const headers = ['门类', '大类', '中类', '小类', '类别名称', '说明'];
XLSX.utils.sheet_add_aoa(worksheet, [headers], { origin: 'A1' });

// 调整列宽
worksheet['!cols'] = [
  { wch: 8 },   // 门类
  { wch: 8 },   // 大类
  { wch: 8 },   // 中类
  { wch: 8 },   // 小类
  { wch: 40 },  // 类别名称
  { wch: 100 }  // 说明
];

// 添加工作表到工作簿
XLSX.utils.book_append_sheet(workbook, worksheet, '国民经济行业分类和代码');

// 保存Excel文件
const outputFilePath = 'output_from_docx.xlsx';
XLSX.writeFile(workbook, outputFilePath);

console.log('Excel文件生成成功：', outputFilePath);
