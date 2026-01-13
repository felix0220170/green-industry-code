const officeparser = require('officeparser');
const XLSX = require('xlsx');
const fs = require('fs');

async function convertDocxTableToExcel() {
  try {
    console.log('正在读取docx文件...');
    
    // 解析docx文件中的表格
    const tables = await officeparser.parseOffice('source.docx', { tables: true });
    
    console.log('成功读取docx文件，共找到', tables.length, '个表格');
    
    // 查找表1：国民经济行业分类和代码
    let targetTable = null;
    for (let i = 0; i < tables.length; i++) {
      const table = tables[i];
      console.log(`表格 ${i+1} 的大小：${table.length} 行 x ${table[0]?.length || 0} 列`);
      
      // 检查表格是否包含目标表头
      if (table.length > 0 && table[0].length >= 4) {
        const firstRow = table[0].map(cell => cell?.trim());
        if (firstRow.includes('门类') && firstRow.includes('大类') && firstRow.includes('中类') && firstRow.includes('小类')) {
          targetTable = table;
          console.log('找到目标表格：表1 国民经济行业分类和代码');
          break;
        }
      }
    }
    
    if (!targetTable) {
      console.error('未找到包含"门类、大类、中类、小类"的表格');
      return;
    }
    
    // 处理表格数据
    const excelData = [];
    
    // 跳过表头行（前几行可能是标题和表头）
    let startRow = 0;
    for (; startRow < targetTable.length; startRow++) {
      const row = targetTable[startRow].map(cell => cell?.trim());
      if (row.includes('门类') && row.includes('大类') && row.includes('中类') && row.includes('小类')) {
        break;
      }
    }
    
    // 从数据行开始处理
    for (let i = startRow + 1; i < targetTable.length; i++) {
      const row = targetTable[i];
      
      // 确保行有足够的单元格
      if (row.length >= 6) {
        const entry = {
          门类: row[0]?.trim() || '',
          大类: row[1]?.trim() || '',
          中类: row[2]?.trim() || '',
          小类: row[3]?.trim() || '',
          类别名称: row[4]?.trim() || '',
          说明: row[5]?.trim() || ''
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
    
  } catch (error) {
    console.error('处理过程中发生错误：', error);
  }
}

convertDocxTableToExcel();
