const fs = require('fs');

// 读取output.json文件
const outputData = JSON.parse(fs.readFileSync('dist/output.json', 'utf8'));

// 创建output_txt目录用于存放生成的txt文件
if (!fs.existsSync('dist/output_txt')) {
  fs.mkdirSync('dist/output_txt');
}

// 遍历所有大类
outputData.forEach(category => {
  // 生成文件名，使用code和name的组合
  const fileName = `${category.code}_${category.name.replace(/[\s、]/g, '_')}.txt`;
  const filePath = `dist/output_txt/${fileName}`;
  
  // 将大类的JSON内容写入txt文件
  fs.writeFileSync(filePath, JSON.stringify(category, null, 2), 'utf8');
  
  console.log(`已生成文件: ${filePath}`);
});

console.log('\n所有大类已成功生成对应的txt文件！');
