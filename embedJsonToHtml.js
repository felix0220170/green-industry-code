const fs = require('fs');

// 读取JSON数据
const jsonData = fs.readFileSync('dist/output.json', 'utf8');

// 读取HTML模板
let htmlContent = fs.readFileSync('cascading-select.html', 'utf8');

// 创建新的脚本内容，注意转义模板字符串
const newScript = `        let jsonData = ${jsonData};
        let currentSelections = {
            level1: null,
            level2: null,
            level3: null,
            level4: null
        };

        // 页面加载完成后初始化第一级选择
        window.onload = function() {
            populateLevel1();
        };

        // 填充第一级（门类）
        function populateLevel1() {
            const select = document.getElementById('level1');
            select.innerHTML = '<option value="">请选择门类</option>';
            
            jsonData.forEach(item => {
                const option = document.createElement('option');
                option.value = item.code;
                option.textContent = item.code + ' ' + item.name;
                select.appendChild(option);
            });
        }

        // 填充第二级（大类）
        function populateLevel2(level1Code) {
            const select = document.getElementById('level2');
            select.innerHTML = '<option value="">请选择大类</option>';
            
            const level1Item = jsonData.find(item => item.code === level1Code);
            if (level1Item && level1Item.children) {
                level1Item.children.forEach(item => {
                    const option = document.createElement('option');
                    option.value = item.code;
                    option.textContent = item.code + ' ' + item.name;
                    select.appendChild(option);
                });
            }
        }

        // 填充第三级（中类）
        function populateLevel3(level1Code, level2Code) {
            const select = document.getElementById('level3');
            select.innerHTML = '<option value="">请选择中类</option>';
            
            const level1Item = jsonData.find(item => item.code === level1Code);
            if (level1Item && level1Item.children) {
                const level2Item = level1Item.children.find(item => item.code === level2Code);
                if (level2Item && level2Item.children) {
                    level2Item.children.forEach(item => {
                        const option = document.createElement('option');
                        option.value = item.code;
                        option.textContent = item.code + ' ' + item.name;
                        select.appendChild(option);
                    });
                }
            }
        }

        // 填充第四级（小类）
        function populateLevel4(level1Code, level2Code, level3Code) {
            const select = document.getElementById('level4');
            select.innerHTML = '<option value="">请选择小类</option>';
            
            const level1Item = jsonData.find(item => item.code === level1Code);
            if (level1Item && level1Item.children) {
                const level2Item = level1Item.children.find(item => item.code === level2Code);
                if (level2Item && level2Item.children) {
                    const level3Item = level2Item.children.find(item => item.code === level3Code);
                    if (level3Item && level3Item.children) {
                        level3Item.children.forEach(item => {
                            const option = document.createElement('option');
                            option.value = item.code;
                            option.textContent = item.code + ' ' + item.name;
                            select.appendChild(option);
                        });
                    }
                }
            }
        }

        // 第一级选择变化
        function onLevel1Change() {
            const level1Code = document.getElementById('level1').value;
            currentSelections.level1 = level1Code;
            
            // 重置后续选择
            currentSelections.level2 = null;
            currentSelections.level3 = null;
            currentSelections.level4 = null;
            
            // 清空并重新填充第二级
            document.getElementById('level2').innerHTML = '<option value="">请选择大类</option>';
            document.getElementById('level3').innerHTML = '<option value="">请先选择大类</option>';
            document.getElementById('level4').innerHTML = '<option value="">请先选择中类</option>';
            
            if (level1Code) {
                populateLevel2(level1Code);
            }
            
            updateResult();
        }

        // 第二级选择变化
        function onLevel2Change() {
            const level1Code = currentSelections.level1;
            const level2Code = document.getElementById('level2').value;
            currentSelections.level2 = level2Code;
            
            // 重置后续选择
            currentSelections.level3 = null;
            currentSelections.level4 = null;
            
            // 清空并重新填充第三级
            document.getElementById('level3').innerHTML = '<option value="">请选择中类</option>';
            document.getElementById('level4').innerHTML = '<option value="">请先选择中类</option>';
            
            if (level1Code && level2Code) {
                populateLevel3(level1Code, level2Code);
            }
            
            updateResult();
        }

        // 第三级选择变化
        function onLevel3Change() {
            const level1Code = currentSelections.level1;
            const level2Code = currentSelections.level2;
            const level3Code = document.getElementById('level3').value;
            currentSelections.level3 = level3Code;
            
            // 重置后续选择
            currentSelections.level4 = null;
            
            // 清空并重新填充第四级
            document.getElementById('level4').innerHTML = '<option value="">请选择小类</option>';
            
            if (level1Code && level2Code && level3Code) {
                populateLevel4(level1Code, level2Code, level3Code);
            }
            
            updateResult();
        }

        // 第四级选择变化
        function onLevel4Change() {
            currentSelections.level4 = document.getElementById('level4').value;
            updateResult();
        }

        // 更新结果显示
        function updateResult() {
            const resultDiv = document.getElementById('resultText');
            const { level1, level2, level3, level4 } = currentSelections;
            
            if (!level1) {
                resultDiv.innerHTML = '<p>请选择完整的分类路径</p>';
                return;
            }
            
            let resultHTML = '';
            let currentLevel = jsonData;
            let selectedItems = [];
            
            // 查找第一级
            const level1Item = currentLevel.find(item => item.code === level1);
            if (level1Item) {
                selectedItems.push(level1Item);
                resultHTML += '<p><strong>门类：</strong>' + level1Item.code + ' ' + level1Item.name + '</p>';
                
                if (level2 && level1Item.children) {
                    const level2Item = level1Item.children.find(item => item.code === level2);
                    if (level2Item) {
                        selectedItems.push(level2Item);
                        resultHTML += '<p><strong>大类：</strong>' + level2Item.code + ' ' + level2Item.name + '</p>';
                        
                        if (level3 && level2Item.children) {
                            const level3Item = level2Item.children.find(item => item.code === level3);
                            if (level3Item) {
                                selectedItems.push(level3Item);
                                resultHTML += '<p><strong>中类：</strong>' + level3Item.code + ' ' + level3Item.name + '</p>';
                                
                                if (level4 && level3Item.children) {
                                    const level4Item = level3Item.children.find(item => item.code === level4);
                                    if (level4Item) {
                                        selectedItems.push(level4Item);
                                        resultHTML += '<p><strong>小类：</strong>' + level4Item.code + ' ' + level4Item.name + '</p>';
                                        
                                        // 显示说明
                                        if (level4Item.desc) {
                                            resultHTML += '<p><strong>说明：</strong>' + level4Item.desc + '</p>';
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            
            resultDiv.innerHTML = resultHTML;
        }`;

// 替换HTML中的脚本部分
htmlContent = htmlContent.replace(/<script>[\s\S]*?<\/script>/, `<script>${newScript}</script>`);

// 保存生成的HTML文件
fs.writeFileSync('cascading-select-standalone.html', htmlContent);

console.log('已生成包含嵌入JSON数据的独立HTML文件：cascading-select-standalone.html');
