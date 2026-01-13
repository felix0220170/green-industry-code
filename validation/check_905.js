const XLSX = require('xlsx');

// 读取生成的Excel文件
const workbook = XLSX.readFile('asset/output_from_docx.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(worksheet);

console.log('开始检查905相关数据：');

// 查找所有与905相关的行
const rows905 = data.filter(row => {
  const subcategory = row['中类'] || '';
  const item = row['小类'] || '';
  const name = row['类别名称'] || '';
  
  return subcategory.includes('905') || item.includes('905');
});

console.log('找到与905相关的行：', rows905.length);

rows905.forEach((row, index) => {
  console.log(`\n第${index + 1}行：`);
  console.log('  门类:', row['门类']);
  console.log('  大类:', row['大类']);
  console.log('  中类:', row['中类']);
  console.log('  小类:', row['小类']);
  console.log('  类别名称:', JSON.stringify(row['类别名称'])); // 使用JSON.stringify查看换行符
  console.log('  说明:', row['说明']);
});

// 检查原始的Word文档解析结果
const AdmZip = require('adm-zip');
const zip = new AdmZip('asset/source.docx');
const zipEntries = zip.getEntries();

let documentXml = null;
for (const entry of zipEntries) {
  if (entry.entryName === 'word/document.xml') {
    documentXml = entry.getData().toString('utf8');
    break;
  }
}

if (documentXml) {
  // 查找包含905的表格行
  const rows = documentXml.match(/<w:tr[^>]*>([\s\S]*?)<\/w:tr>/g) || [];
  let found = false;
  
  for (const row of rows) {
    if (row.includes('905')) {
      console.log('\n找到包含905的Word原始行：');
      // 提取单元格内容
      const cells = row.match(/<w:tc[^>]*>([\s\S]*?)<\/w:tc>/g) || [];
      cells.forEach((cell, idx) => {
        console.log(`  单元格${idx + 1}：`);
        // 简单提取文本内容，保留段落结构
        const paragraphs = cell.match(/<w:p[^>]*>([\s\S]*?)<\/w:p>/g) || [];
        paragraphs.forEach((p, pIdx) => {
          // 提取文本
          let text = p.replace(/<[^>]*>/g, '');
          text = text.replace(/&amp;/g, '&')
                    .replace(/&lt;/g, '<')
                    .replace(/&gt;/g, '>')
                    .replace(/&quot;/g, '"')
                    .replace(/&apos;/g, "'");
          console.log(`    段落${pIdx + 1}：${JSON.stringify(text.trim())}`);
        });
      });
      found = true;
      break;
    }
  }
  
  if (!found) {
    console.log('未找到包含905的Word原始行');
  }
}
