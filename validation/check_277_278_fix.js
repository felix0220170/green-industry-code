const XLSX = require('xlsx');
const fs = require('fs');

// 读取生成的Excel文件
const workbook = XLSX.readFile('asset/output_from_docx.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(worksheet);

console.log('检查277/278中类问题修复情况：');
console.log('---');

// 查找包含277或278的行
const foundRows = data.filter(row => {
  return row['中类']?.toString().includes('277') || row['中类']?.toString().includes('278');
});

foundRows.forEach((row, index) => {
  console.log(`找到行 ${index + 1}:`);
  console.log(`  门类: ${row['门类']}`);
  console.log(`  大类: ${row['大类']}`);
  console.log(`  中类: ${row['中类']}`);
  console.log(`  小类: ${row['小类']}`);
  console.log(`  名称: ${row['类别名称']}`);
  console.log('  中类包含换行符:', row['中类']?.includes('\n'));
  console.log('  中类包含277278:', row['中类']?.includes('277278'));
  console.log('---');
});

// 检查是否正确分离为两行
const has277 = foundRows.some(row => row['中类'] === '277' && row['小类'] === '2770');
const has278 = foundRows.some(row => row['中类'] === '278' && row['小类'] === '2780');

console.log('修复结果：');
console.log(`  277 2770 卫生材料及医药用品制造 行是否存在: ${has277}`);
console.log(`  278 2780 药用辅料及包装材料制造 行是否存在: ${has278}`);
console.log(`  中类是否包含277278合并: ${foundRows.some(row => row['中类']?.includes('277278'))}`);
console.log(`  中类是否包含换行符: ${foundRows.some(row => row['中类']?.includes('\n'))}`);

if (has277 && has278 && !foundRows.some(row => row['中类']?.includes('277278')) && !foundRows.some(row => row['中类']?.includes('\n'))) {
  console.log('✅ 修复成功！277/278问题已解决。');
} else {
  console.log('❌ 修复失败！277/278问题仍然存在。');
}
