const AdmZip = require('adm-zip');
const fs = require('fs');

// 读取并解压docx文件
const zip = new AdmZip('asset/source.docx');
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
const targetTable = tables[0];

// 提取表格中的行
const rows = targetTable.match(/<w:tr[^>]*>([\s\S]*?)<\/w:tr>/g) || [];

// 辅助函数：从字符串中提取纯文本（保留段落分隔符）
function extractPlainText(xmlString) {
  // 保留段落标签，用于后续分割
  let text = xmlString.replace(/<w:p[^>]*>/g, '\n');
  // 移除其他所有XML标签
  text = text.replace(/<[^>]*>/g, '');
  
  // 处理可能的XML转义字符
  text = text.replace(/&amp;/g, '&')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"')
            .replace(/&apos;/g, "'");
  
  // 去除多余的空白字符
  text = text.trim();
  
  return text;
}

// 解析行数据
const parsedTable = [];
rows.forEach(row => {
  const cells = row.match(/<w:tc[^>]*>([\s\S]*?)<\/w:tc>/g) || [];
  const parsedRow = [];
  cells.forEach(cell => {
    const cellText = extractPlainText(cell);
    parsedRow.push(cellText);
  });
  if (parsedRow.some(cell => cell.length > 0)) {
    parsedTable.push(parsedRow);
  }
});

// 查找包含277或278的行
console.log('查找包含277或278的行：');
for (let i = 0; i < parsedTable.length; i++) {
  const row = parsedTable[i];
  if (row.some(cell => cell.includes('277') || cell.includes('278'))) {
    console.log(`行 ${i + 1}:`);
    console.log('  门类:', JSON.stringify(row[0]));
    console.log('  大类:', JSON.stringify(row[1]));
    console.log('  中类:', JSON.stringify(row[2]));
    console.log('  小类:', JSON.stringify(row[3]));
    console.log('  名称:', JSON.stringify(row[4]));
    console.log('  说明:', JSON.stringify(row[5]));
    console.log('  中类包含换行符:', row[2].includes('\n'));
    console.log('  小类包含换行符:', row[3].includes('\n'));
    console.log('  名称包含换行符:', row[4].includes('\n'));
    console.log('---');
  }
}
