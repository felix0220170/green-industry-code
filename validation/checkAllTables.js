const AdmZip = require('adm-zip');
const fs = require('fs');

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

// 提取所有表格
const tables = documentXml.match(/<w:tbl[^>]*>([\s\S]*?)<\/w:tbl>/g) || [];
console.log('找到', tables.length, '个表格');

// 保存所有表格内容到文件以便检查
fs.writeFileSync('tables_content.txt', '', 'utf8');

// 解析每个表格的内容
for (let tableIndex = 0; tableIndex < tables.length; tableIndex++) {
  const table = tables[tableIndex];
  
  console.log(`\n=== 表格 ${tableIndex + 1} 内容 ===`);
  
  // 提取表格中的行
  const rows = table.match(/<w:tr[^>]*>([\s\S]*?)<\/w:tr>/g) || [];
  
  let tableText = `表格 ${tableIndex + 1} (${rows.length}行)\n`;
  tableText += '=' .repeat(50) + '\n';
  
  rows.forEach((row, rowIndex) => {
    // 提取单元格
    const cells = row.match(/<w:tc[^>]*>([\s\S]*?)<\/w:tc>/g) || [];
    
    let rowText = `行 ${rowIndex + 1}: `;
    
    cells.forEach((cell, cellIndex) => {
      // 提取单元格文本
      const texts = cell.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g) || [];
      let cellText = '';
      
      texts.forEach(text => {
        cellText += text.replace(/<\/w:t>/g, '').replace(/<w:t[^>]*>/g, '');
      });
      
      rowText += `[${cellIndex + 1}]: "${cellText}" `;
    });
    
    tableText += rowText + '\n';
  });
  
  // 将表格内容写入文件
  fs.appendFileSync('tables_content.txt', tableText + '\n\n', 'utf8');
  
  // 只在控制台显示前5行，避免输出过多
  const previewLines = tableText.split('\n').slice(0, 10).join('\n');
  console.log(previewLines);
  if (tableText.split('\n').length > 10) {
    console.log('... (更多内容请查看tables_content.txt文件)');
  }
}

console.log('\n所有表格内容已保存到tables_content.txt文件');
