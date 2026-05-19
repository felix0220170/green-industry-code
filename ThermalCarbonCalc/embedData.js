const fs = require('fs');

console.log('正在加载 JSON 数据...');

const wenshangData = fs.readFileSync('data/wenshang_table.json', 'utf8');
const steamData = fs.readFileSync('data/steam_table_easyquery_style.json', 'utf8');

console.log('正在读取原始 JavaScript 文件...');

let jsContent = fs.readFileSync('js/thermal-carbon-calc.js', 'utf8');

console.log('正在替换数据加载部分...');

const dataSection = `const REF_TEMP = 20;
let refEnthalpy = null;

let wenshangTableData = null;
let steamTableData = null;
let dataLoaded = false;
let pressureToTempMap = null;

function loadData() {
    return Promise.all([
        fetch('data/wenshang_table.json').then(response => response.json()),
        fetch('data/steam_table_easyquery_style.json').then(response => response.json())
    ]).then(([wenshangData, steamData]) => {
        wenshangTableData = wenshangData.table;
        steamTableData = steamData;
        dataLoaded = true;
        return { wenshangTable: wenshangTableData, steamTable: steamTableData };
    });
}`;

const embeddedSection = `const REF_TEMP = 20;
let refEnthalpy = null;

let wenshangTableData = null;
let steamTableData = null;
let dataLoaded = false;
let pressureToTempMap = null;

const wenshangJsonData = ${wenshangData};
const steamJsonData = ${steamData};

function loadData() {
    return new Promise((resolve) => {
        wenshangTableData = wenshangJsonData.table;
        steamTableData = steamJsonData;
        dataLoaded = true;
        resolve({ wenshangTable: wenshangTableData, steamTable: steamTableData });
    });
}`;

jsContent = jsContent.replace(dataSection, embeddedSection);

console.log('正在写入新的 JavaScript 文件...');

fs.writeFileSync('js/thermal-carbon-calc.js', jsContent);

console.log('✅ 数据嵌入完成！现在可以直接双击打开 index.html 文件。');