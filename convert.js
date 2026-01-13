const XLSX = require('xlsx');
const fs = require('fs');

// 读取Excel文件
const workbook = XLSX.readFile('asset/output_from_docx.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// 将Excel转换为JSON
const jsonData = XLSX.utils.sheet_to_json(worksheet);

// 创建层级映射
const sectionMap = new Map(); // 门类映射
const categoryMap = new Map(); // 大类映射
const subcategoryMap = new Map(); // 中类映射
const itemMap = new Map(); // 小类映射
const result = [];

// 遍历所有行，构建四层结构
jsonData.forEach(row => {
  const section = row.门类;
  const category = row.大类;
  const subcategory = row.中类;
  const item = row.小类;
  const name = row.类别名称;
  const desc = row.说明;
  
  // 确保section存在
  if (!section) return;
  
  // 1. 首先确保门类存在
  if (!sectionMap.has(section)) {
    const sectionNode = {
      code: section,
      name: section // 临时名称
    };
    sectionMap.set(section, sectionNode);
    sectionNode.children = [];
    result.push(sectionNode);
  }
  
  const sectionNode = sectionMap.get(section);
  
  // 2. 处理门类（一级） - 仅处理有名称的情况
  if (section && !category && !subcategory && !item && name && name.trim() !== section) {
    // 更新门类名称和说明
    sectionNode.name = name;
    if (desc && desc.trim() !== '') {
      sectionNode.desc = desc;
    }
  }
  // 3. 处理大类（二级）
  else if (section && category && !subcategory && !item) {
    const categoryKey = `${section}-${category}`;
    
    // 检查大类是否已存在
    if (!categoryMap.has(categoryKey)) {
      const categoryNode = {
        code: category,
        name: name
      };
      if (desc && desc.trim() !== '') {
        categoryNode.desc = desc;
      }
      categoryMap.set(categoryKey, categoryNode);
      categoryNode.children = [];
      sectionNode.children.push(categoryNode);
    }
  }
  // 4. 处理中类（三级）
  else if (section && category && subcategory && !item) {
    const categoryKey = `${section}-${category}`;
    const subcategoryKey = `${section}-${category}-${subcategory}`;
    
    // 确保大类存在
    if (!categoryMap.has(categoryKey)) {
      const categoryNode = {
        code: category,
        name: category // 临时名称
      };
      categoryMap.set(categoryKey, categoryNode);
      categoryNode.children = [];
      sectionNode.children.push(categoryNode);
    }
    
    const categoryNode = categoryMap.get(categoryKey);
    
    // 检查中类是否已存在
    if (!subcategoryMap.has(subcategoryKey)) {
      const subcategoryNode = {
        code: subcategory,
        name: name
      };
      if (desc && desc.trim() !== '') {
        subcategoryNode.desc = desc;
      }
      subcategoryMap.set(subcategoryKey, subcategoryNode);
      subcategoryNode.children = [];
      categoryNode.children.push(subcategoryNode);
    }
  }
  // 5. 处理小类（四级）
  else if (section && category && subcategory && item) {
    const categoryKey = `${section}-${category}`;
    const subcategoryKey = `${section}-${category}-${subcategory}`;
    
    // 确保大类存在
    if (!categoryMap.has(categoryKey)) {
      const categoryNode = {
        code: category,
        name: category // 临时名称
      };
      categoryMap.set(categoryKey, categoryNode);
      categoryNode.children = [];
      sectionNode.children.push(categoryNode);
    }
    
    // 确保中类存在
    if (!subcategoryMap.has(subcategoryKey)) {
      const categoryNode = categoryMap.get(categoryKey);
      const subcategoryNode = {
        code: subcategory,
        name: name // 当只有一个小类时，使用小类名称作为中类名称
      };
      subcategoryMap.set(subcategoryKey, subcategoryNode);
      subcategoryNode.children = [];
      categoryNode.children.push(subcategoryNode);
    }
    
    const subcategoryNode = subcategoryMap.get(subcategoryKey);
    const itemNode = {
      code: item,
      name: name
    };
    if (desc && desc.trim() !== '') {
      itemNode.desc = desc;
    }
    // 小类是叶子节点，不添加children属性
    subcategoryNode.children.push(itemNode);
  }
});

// 手动设置所有门类的名称和说明
// 这是根据国家标准GB/T 4754-2017的固定信息
const sectionInfo = {
  'A': { name: '农、林、牧、渔业', desc: '本门类包括01～05大类' },
  'B': { name: '采矿业', desc: '本门类包括06～12大类' },
  'C': { name: '制造业', desc: '本门类包括13～43大类' },
  'D': { name: '电力、热力、燃气及水生产和供应业', desc: '本门类包括44～46大类' },
  'E': { name: '建筑业', desc: '本门类包括47～50大类' },
  'F': { name: '批发和零售业', desc: '本门类包括51～52大类' },
  'G': { name: '交通运输、仓储和邮政业', desc: '本门类包括53～60大类' },
  'H': { name: '住宿和餐饮业', desc: '本门类包括61～62大类' },
  'I': { name: '信息传输、软件和信息技术服务业', desc: '本门类包括63～65大类' },
  'J': { name: '金融业', desc: '本门类包括66～69大类' },
  'K': { name: '房地产业', desc: '本门类包括70大类' },
  'L': { name: '租赁和商务服务业', desc: '本门类包括71～72大类' },
  'M': { name: '科学研究和技术服务业', desc: '本门类包括73～75大类' },
  'N': { name: '水利、环境和公共设施管理业', desc: '本门类包括76～79大类' },
  'O': { name: '居民服务、修理和其他服务业', desc: '本门类包括80～82大类' },
  'P': { name: '教育', desc: '本门类包括83大类' },
  'Q': { name: '卫生和社会工作', desc: '本门类包括84～85大类' },
  'R': { name: '文化、体育和娱乐业', desc: '本门类包括86～90大类' },
  'S': { name: '公共管理、社会保障和社会组织', desc: '本门类包括91～96大类' },
  'T': { name: '国际组织', desc: '本门类包括97大类' }
};

// 更新所有门类的名称和说明
result.forEach(node => {
  if (sectionInfo[node.code]) {
    node.name = sectionInfo[node.code].name;
    node.desc = sectionInfo[node.code].desc;
  }
});

// 移除没有子节点的children属性
function removeEmptyChildren(node) {
  if (node.children && node.children.length === 0) {
    delete node.children;
  } else if (node.children) {
    node.children.forEach(child => removeEmptyChildren(child));
  }
  return node;
}

// 处理所有节点
const finalResult = result.map(node => removeEmptyChildren(node));

// 将结果写入JSON文件
fs.writeFileSync('dist/output.json', JSON.stringify(finalResult, null, 2), 'utf8');

console.log('转换完成，结果已保存到dist/output.json');
console.log('共处理', jsonData.length, '行数据');
console.log('生成', finalResult.length, '个门类');

