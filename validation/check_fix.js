const XLSX = require('xlsx');

// 读取生成的Excel文件
const workbook = XLSX.readFile('asset/output_from_docx.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(worksheet);

console.log('开始检查修复结果：');
console.log('总记录数:', data.length);

// 检查8052和8053是否存在
const has8052 = data.some(row => row['小类'] === '8052');
const has8053 = data.some(row => row['小类'] === '8053');
const has80528053 = data.some(row => row['小类'] === '80528053');

console.log('\n检查8052/8053问题：');
console.log('是否存在8052:', has8052);
console.log('是否存在8053:', has8053);
console.log('是否存在80528053:', has80528053);

// 检查8053的中类是否正确
const row8053 = data.find(row => row['小类'] === '8053');
if (row8053) {
  console.log('8053的中类:', row8053['中类']);
  console.log('8053的类别名称:', row8053['类别名称']);
}

// 检查2770和2780是否存在
const has2770 = data.some(row => row['小类'] === '2770');
const has2780 = data.some(row => row['小类'] === '2780');
const has27702780 = data.some(row => row['小类'] === '27702780');
const has277278 = data.some(row => row['中类'] === '277278');

console.log('\n检查277/278问题：');
console.log('是否存在2770:', has2770);
console.log('是否存在2780:', has2780);
console.log('是否存在27702780:', has27702780);
console.log('是否存在中类277278:', has277278);

// 检查2770和2780的详细信息
const row2770 = data.find(row => row['小类'] === '2770');
if (row2770) {
  console.log('2770的中类:', row2770['中类']);
  console.log('2770的类别名称:', row2770['类别名称']);
}

const row2780 = data.find(row => row['小类'] === '2780');
if (row2780) {
  console.log('2780的中类:', row2780['中类']);
  console.log('2780的类别名称:', row2780['类别名称']);
}

// 检查是否有其他明显的合并问题
const mergedIssues = data.filter(row => {
  const subcategory = row['中类'] || '';
  const item = row['小类'] || '';
  const name = row['类别名称'] || '';
  
  // 检查代码是否过长（可能是合并的）
  const hasLongCode = (subcategory.length > 3 && subcategory.match(/^\d+$/)) || 
                     (item.length > 4 && item.match(/^\d+$/));
  
  // 检查名称是否包含多个类别的迹象
  const hasMultipleNames = name.includes('及') && name.includes('制造');
  
  return hasLongCode || hasMultipleNames;
});

if (mergedIssues.length > 0) {
  console.log('\n发现可能的合并问题：');
  mergedIssues.forEach((row, index) => {
    console.log(`${index + 1}. 中类:${row['中类']}, 小类:${row['小类']}, 名称:${row['类别名称']}`);
  });
} else {
  console.log('\n未发现明显的合并问题');
}

console.log('\n检查完成！');
