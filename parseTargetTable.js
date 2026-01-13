const AdmZip = require('adm-zip');
const XLSX = require('xlsx');
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
console.log('找到', tables.length, '个表格');

// 使用第一个表格（从tables_content.txt中看到这是我们需要的表格）
const targetTable = tables[0];
if (!targetTable) {
  console.error('未找到表格');
  process.exit(1);
}

// 提取表格中的行
const rows = targetTable.match(/<w:tr[^>]*>([\s\S]*?)<\/w:tr>/g) || [];
console.log('目标表格包含', rows.length, '行');

// 解析每行的单元格
  const parsedTable = [];
  
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
  
  rows.forEach(row => {
    // 提取单元格
    const cells = row.match(/<w:tc[^>]*>([\s\S]*?)<\/w:tc>/g) || [];
    
    const parsedRow = [];
    cells.forEach(cell => {
      // 直接提取单元格中的所有纯文本
      const cellText = extractPlainText(cell);
      parsedRow.push(cellText);
    });
    
    // 只添加包含实际内容的行
    if (parsedRow.some(cell => cell.length > 0)) {
      parsedTable.push(parsedRow);
    }
  });

// 处理表格数据，从第3行开始（跳过标题和表头）
const excelData = [];

// 保存上一级的代码，用于处理合并单元格的情况
let lastSection = '';
let lastCategory = '';
let lastSubcategory = '';

for (let i = 2; i < parsedTable.length; i++) { // 从第3行开始
  const row = parsedTable[i];
  
  // 确保行有足够的单元格
  if (row.length >= 6) {
    // 先更新上一级代码，处理合并单元格的情况
    if (row[0].trim() !== '') {
      lastSection = row[0].trim();
      lastCategory = ''; // 切换门类时重置大类
      lastSubcategory = ''; // 切换门类时重置中类
    }
    if (row[1].trim() !== '') {
      lastCategory = row[1].trim();
      lastSubcategory = ''; // 切换大类时重置中类
    }
    if (row[2].trim() !== '') lastSubcategory = row[2].trim();
    
    // 使用最新的上一级代码填充当前行
    const section = row[0].trim() || lastSection;
    const category = row[1].trim() || lastCategory;
    
    // 先获取原始单元格内容，保留换行符信息
    const rawSubcategoryCell = row[2];
    const rawItemCell = row[3];
    const rawNameCell = row[4];
    
    // 处理可能包含换行符的单元格
    let subcategoryCell = rawSubcategoryCell.trim();
    const itemCell = rawItemCell.trim();
    let nameCell = rawNameCell.trim();
    const descCell = row[5].trim();
    
    // 去除名称前面的多余空格
    nameCell = nameCell.trim();
    
    // 检查是否需要分割行
    const hasLineBreaks = rawSubcategoryCell.includes('\n') || rawItemCell.includes('\n') || rawNameCell.includes('\n');
    
    // 处理合并单元格的情况
    if (subcategoryCell === '') subcategoryCell = lastSubcategory;
    
    if (hasLineBreaks) {
      // 特殊处理：当中类本身包含多个段落（如277\n\n\n\n278）
      // 先将多个连续换行符替换为单个换行符，再分割
      const normalizedSubcategory = rawSubcategoryCell.replace(/\n+/g, '\n');
      const normalizedItem = rawItemCell.replace(/\n+/g, '\n');
      const normalizedName = rawNameCell.replace(/\n+/g, '\n');
      
      const subcategoryLines = normalizedSubcategory.split('\n').filter(line => line.trim() !== '');
      
      // 初始化tempSubcategory
      let tempSubcategory = subcategoryCell;
      if (subcategoryLines.length > 1) {
        // 按换行符分割小类、名称和说明
        const itemLines = normalizedItem.split('\n').filter(line => line.trim() !== '');
        const nameLines = normalizedName.split('\n').filter(line => line.trim() !== '');
        const descLines = row[5].trim() ? row[5].trim().split('\n').filter(line => line.trim() !== '') : [];
        
        // 处理每个中类
        for (let j = 0; j < subcategoryLines.length; j++) {
          const subcategory = subcategoryLines[j].trim();
          const item = itemLines[j]?.trim() || '';
          const name = nameLines[j]?.trim() || '';
          const desc = descLines[j]?.trim() || '';
          
          // 创建中类行（如果有名称）
          if (name !== '') {
            const entry = {
              门类: section,
              大类: category,
              中类: subcategory,
              小类: item,
              类别名称: name,
              说明: desc
            };
            excelData.push(entry);
          }
          
          // 更新lastSubcategory
          lastSubcategory = subcategory;
        }
      } else {
        // 特殊处理：当存在中类但小类或名称包含多个段落时
        const hasSubcategoryWithMultipleItems = subcategoryCell !== '' && (rawItemCell.includes('\n') || rawNameCell.includes('\n'));
        
        // 先处理中类行（如果中类存在且小类/名称有多个段落）
        if (hasSubcategoryWithMultipleItems) {
          // 提取中类名称（第一个段落）
          const subcategoryName = nameCell.split('\n')[0]?.trim() || '';
          
          // 创建中类行
          if (subcategoryName !== '') {
            const subcategoryEntry = {
              门类: section,
              大类: category,
              中类: subcategoryCell,
              小类: '',
              类别名称: subcategoryName,
              说明: ''
            };
            excelData.push(subcategoryEntry);
          }
        }
        
        // 按换行符分割小类和名称
        const itemLines = itemCell.split('\n').filter(line => line.trim() !== '');
        const nameLines = nameCell.split('\n').filter(line => line.trim() !== '');
        
        // 确定最大行数
        const maxLines = Math.max(itemLines.length, nameLines.length);
        
        // 临时保存当前行的中类，用于后续行继承
        let tempSubcategory = subcategoryCell;
        
        // 创建小类数据
        for (let j = 0; j < maxLines; j++) {
          const item = itemLines[j] || '';
          // 对于多个名称段落的情况，跳过第一个（中类名称）
          const name = hasSubcategoryWithMultipleItems ? (nameLines[j + 1] || '') : (nameLines[j] || '');
          
          // 创建数据条目
          const entry = {
            门类: section,
            大类: category,
            中类: subcategoryCell,
            小类: item,
            类别名称: name,
            说明: descCell // 说明字段通常不会有换行符，如果有可以类似处理
          };
          
          // 只添加有实际内容的行（小类或名称必须有一个非空）
          if (item !== '' || name.trim() !== '') {
            excelData.push(entry);
          }
        }
        
        // 更新全局的lastSubcategory
        lastSubcategory = tempSubcategory;
      }
      
      // 更新全局的lastSubcategory
      lastSubcategory = tempSubcategory;
    } else {
      // 正常处理，不分割行
      let subcategory = subcategoryCell;
      const item = itemCell;
      
      // 如果有小类但中类为空，继承lastSubcategory
      if (item !== '' && subcategory === '') {
        subcategory = lastSubcategory;
      }
      
      // 创建数据条目
      const entry = {
        门类: section,
        大类: category,
        中类: subcategory,
        小类: item,
        类别名称: nameCell,
        说明: descCell
      };
      
      // 只添加有实际内容的行
      if (Object.values(entry).some(value => value !== '')) {
        excelData.push(entry);
      }
    }
  }
}

console.log('表格数据处理完成，共提取', excelData.length, '条有效记录');

// 创建Excel工作簿
const workbook = XLSX.utils.book_new();
const worksheet = XLSX.utils.json_to_sheet(excelData);

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
const outputFilePath = 'asset/output_from_docx.xlsx';
XLSX.writeFile(workbook, outputFilePath);

console.log('Excel文件生成成功：', outputFilePath);
