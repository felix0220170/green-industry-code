const mammoth = require('mammoth');
const XLSX = require('xlsx');
const fs = require('fs');

async function convertDocxToExcel() {
  try {
    // 读取docx文件
    const result = await mammoth.extractRawText({
      path: 'source.docx'
    });
    
    const text = result.value;
    console.log('文档读取成功，开始解析表格...');
    
    // 查找表1的位置
    const tableStartMarker = '表1 国民经济行业分类和代码';
    const startIndex = text.indexOf(tableStartMarker);
    
    if (startIndex === -1) {
      console.error('未找到"表1 国民经济行业分类和代码"');
      return;
    }
    
    // 提取表1的内容（简单处理，实际可能需要更复杂的解析）
    let tableContent = text.substring(startIndex);
    
    // 查找下一个表格或文档结束
    const tableEndMarker = '表2';
    const endIndex = tableContent.indexOf(tableEndMarker);
    if (endIndex !== -1) {
      tableContent = tableContent.substring(0, endIndex);
    }
    
    // 按行分割
    const lines = tableContent.split('\n').filter(line => line.trim() !== '');
    
    // 处理表格数据（这里需要根据实际表格格式调整）
    const excelData = [];
    
    // 跳过表头行
    for (let i = 5; i < lines.length; i++) {
      const line = lines[i].trim();
      // 简单的行解析，实际可能需要更复杂的处理
      // 这里假设每行的格式是：门类 大类 中类 小类 类别名称 说明
      const parts = line.split(/\s+/);
      
      if (parts.length >= 6) {
        const entry = {
          门类: parts[0],
          大类: parts[1],
          中类: parts[2],
          小类: parts[3],
          类别名称: parts[4],
          说明: parts.slice(5).join(' ')
        };
        excelData.push(entry);
      }
    }
    
    console.log('表格解析完成，共找到', excelData.length, '条记录');
    
    // 创建Excel工作簿
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(excelData);
    
    // 调整列宽
    const colWidths = [
      { wch: 6 },  // 门类
      { wch: 6 },  // 大类
      { wch: 6 },  // 中类
      { wch: 6 },  // 小类
      { wch: 40 }, // 类别名称
      { wch: 100 } // 说明
    ];
    worksheet['!cols'] = colWidths;
    
    // 添加工作表并保存
    XLSX.utils.book_append_sheet(workbook, worksheet, '行业分类');
    XLSX.writeFile(workbook, 'output_from_docx.xlsx');
    
    console.log('Excel文件生成成功：output_from_docx.xlsx');
    
  } catch (error) {
    console.error('处理过程中发生错误：', error);
  }
}

convertDocxToExcel();
