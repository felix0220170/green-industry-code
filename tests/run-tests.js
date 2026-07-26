// 测试运行器 - 加载源文件并运行测试用例
const fs = require('fs');
const path = require('path');
const testCases = require('./test-cases');

// 报告存储
let results = [];
let startTime = null;

// 从源文件中提取核心计算函数
function loadCalculator() {
  const sourcePath = path.join(__dirname, '../ThermalCarbonCalc/js/thermal-carbon-calc.js');
  let sourceCode = fs.readFileSync(sourcePath, 'utf-8');
  
  // 创建沙箱环境
  const sandbox = {
    console: {
      log: () => {},
      error: (msg) => {
        sandbox._errors = sandbox._errors || [];
        sandbox._errors.push(msg);
      },
      warn: () => {}
    },
    window: {
      addEventListener: () => {}
    },
    document: {
      addEventListener: (event, callback) => {
        if (event === 'DOMContentLoaded') {
          callback();
        }
      },
      getElementById: () => ({ 
        addEventListener: () => {}, 
        appendChild: () => {},
        setAttribute: () => {},
        style: {},
        classList: { add: () => {}, remove: () => {} }
      }),
      querySelector: () => ({ addEventListener: () => {} }),
      querySelectorAll: () => []
    },
    alert: () => {},
    confirm: () => true,
    setTimeout: () => {}
  };

  // 在全局作用域中执行源代码（使用 globalThis）
  const wrappedCode = `
    // 在执行前设置一个全局引用
    var _testEnv = {};
    
    ${sourceCode}
    
    // 确保初始化完成
    if (!dataLoaded) {
      loadData();
    }
    if (refEnthalpy === null) {
      refEnthalpy = getSaturatedWaterEnthalpy(REF_TEMP);
    }
    if (pressureToTempMap === null) {
      pressureToTempMap = buildPressureToTempMap();
    }
    
    // 暴露核心计算函数到 globalThis
    _testEnv.calculate = calculate;
    _testEnv.calculateHotWater = calculateHotWater;
    _testEnv.calculateSaturatedSteam = calculateSaturatedSteam;
    _testEnv.calculateSuperheatedSteam = calculateSuperheatedSteam;
    _testEnv.getSaturatedWaterEnthalpy = getSaturatedWaterEnthalpy;
    _testEnv.getSaturatedSteamEnthalpy = getSaturatedSteamEnthalpy;
    _testEnv.getSaturationTempByPressure = getSaturationTempByPressure;
    _testEnv.getEnthalpyByPT = getEnthalpyByPT;
    _testEnv.getEnthalpyValue = getEnthalpyValue;
    _testEnv.linearInterpolation = linearInterpolation;
    _testEnv.loadData = loadData;
    _testEnv.wenshangTableData = wenshangTableData;
    _testEnv.steamTableData = steamTableData;
    _testEnv.refEnthalpy = refEnthalpy;
    _testEnv.REF_TEMP = REF_TEMP;
    _testEnv.dataLoaded = dataLoaded;
    _testEnv.buildPressureToTempMap = buildPressureToTempMap;
    
    globalThis._calcFunctions = _testEnv;
  `;
  
  try {
    // 执行代码
    const fn = new Function('console', 'window', 'document', 'alert', 'confirm', 'setTimeout', wrappedCode);
    fn(sandbox.console, sandbox.window, sandbox.document, sandbox.alert, sandbox.confirm, sandbox.setTimeout);
    
    const calc = global._calcFunctions;
    
    return calc;
  } catch (e) {
    console.error('加载源文件失败:', e.message);
    console.error(e.stack);
    return null;
  }
}

// 运行单个测试用例
function runTestCase(calc, testCase) {
  const result = {
    id: testCase.id,
    name: testCase.name,
    input: testCase.input,
    expected: testCase.expected,
    passed: false,
    actualGJ: null,
    error: null,
    details: null
  };

  try {
    const { mediumType, weight, temp, pressure } = testCase.input;
    
    // 验证输入
    const validation = validateInput(calc, mediumType, weight, temp, pressure);
    if (!validation.valid) {
      if (testCase.expected.success === false) {
        result.passed = true;
        result.message = `预期失败，验证正确: ${validation.message}`;
      } else {
        result.passed = false;
        result.message = `预期成功，但验证失败: ${validation.message}`;
      }
      return result;
    }
    
    // 执行计算
    const calcResult = calc.calculate(mediumType, weight, temp, pressure);
    
    if (calcResult === null) {
      if (testCase.expected.success === false) {
        result.passed = true;
        result.message = '预期计算失败（返回null）';
      } else {
        result.passed = false;
        result.message = '预期计算成功，但返回null';
      }
      return result;
    }
    
    // 验证结果
    result.actualGJ = calcResult.value;
    result.details = calcResult.details;
    
    if (testCase.expected.success === false) {
      result.passed = false;
      result.message = '预期失败，但计算成功';
      return result;
    }
    
    // 检查结果范围
    if (testCase.expected.minGJ !== undefined && testCase.expected.maxGJ !== undefined) {
      if (calcResult.value >= testCase.expected.minGJ && calcResult.value <= testCase.expected.maxGJ) {
        result.passed = true;
        result.message = `计算成功，GJ=${calcResult.value.toFixed(2)}，在预期范围内`;
      } else {
        result.passed = false;
        result.message = `GJ=${calcResult.value.toFixed(2)}，超出预期范围 [${testCase.expected.minGJ}, ${testCase.expected.maxGJ}]`;
      }
    } else {
      result.passed = calcResult !== null;
      result.message = `计算成功，GJ=${calcResult.value.toFixed(2)}`;
    }
    
  } catch (e) {
    result.error = e.message;
    if (testCase.expected.success === false) {
      result.passed = true;
      result.message = `预期失败，抛出异常: ${e.message}`;
    } else {
      result.passed = false;
      result.message = `预期成功，但抛出异常: ${e.message}`;
    }
  }

  return result;
}

// 简单的输入验证（与 validateBatchInputs 逻辑一致）
function validateInput(calc, mediumType, weight, temp, pressure) {
  if (isNaN(weight) || weight <= 0) {
    return { valid: false, message: '请输入有效的重量（大于0）' };
  }

  if (mediumType === 'hot_water') {
    if (isNaN(temp) || temp < 0 || temp > 374) {
      return { valid: false, message: '请输入有效的温度（0~374°C）' };
    }
  } else if (mediumType === 'saturated_steam') {
    if (isNaN(pressure) || pressure < 0.0006 || pressure > 22.064) {
      return { valid: false, message: '请输入有效的压力（0.0006~22.064 MPa）' };
    }
  } else if (mediumType === 'superheated_steam') {
    if (isNaN(pressure) || pressure < 0.001 || pressure > 100) {
      return { valid: false, message: '请输入有效的压力（0.001~100 MPa）' };
    }
    if (isNaN(temp) || temp < 0 || temp > 800) {
      return { valid: false, message: '请输入有效的温度（0~800°C）' };
    }
    const saturationTemp = calc.getSaturationTempByPressure(pressure);
    if (saturationTemp !== null && temp < saturationTemp) {
      return { valid: false, message: `温度 ${temp.toFixed(2)}°C 低于当前压力下的饱和温度 ${saturationTemp.toFixed(2)}°C` };
    }
  }

  return { valid: true, message: '' };
}

// 运行所有测试
function runAllTests() {
  console.log('='.repeat(60));
  console.log('热力购入量计算器 - 自动化测试套件');
  console.log('='.repeat(60));
  console.log('');
  
  // 加载计算器
  console.log('[1/3] 加载计算器模块...');
  const calc = loadCalculator();
  if (!calc) {
    console.error('❌ 加载失败，退出测试');
    return results;
  }
  console.log('✅ 加载成功');
  console.log('');
  
  // 显示测试用例统计
  console.log('[2/3] 运行测试用例...');
  console.log('');
  
  const categories = [
    { key: 'hot_water', label: '热水测试' },
    { key: 'saturated_steam', label: '饱和蒸汽测试' },
    { key: 'superheated_steam', label: '过热蒸汽测试' }
  ];
  
  for (const cat of categories) {
    const cases = testCases[cat.key];
    if (!cases || cases.length === 0) continue;
    
    console.log(`\n📋 ${cat.label} (${cases.length} 个用例):`);
    console.log('-'.repeat(50));
    
    let catPassed = 0;
    let catFailed = 0;
    
    for (const testCase of cases) {
      const result = runTestCase(calc, testCase);
      results.push(result);
      
      if (result.passed) {
        catPassed++;
        console.log(`  ✅ ${result.id}: ${result.name}`);
        console.log(`     ${result.message}`);
      } else {
        catFailed++;
        console.log(`  ❌ ${result.id}: ${result.name}`);
        console.log(`     ${result.message}`);
        if (result.error) {
          console.log(`     错误: ${result.error}`);
        }
      }
    }
    
    console.log(`  小计: ${catPassed}/${cases.length} 通过`);
  }
  
  console.log('');
  console.log('[3/3] 生成报告...');
  console.log('');
  
  return results;
}

// 生成报告
function generateReport(results) {
  const total = results.length;
  const passed = results.filter(r => r.passed).length;
  const failed = total - passed;
  
  let report = '';
  report += '='.repeat(70) + '\n';
  report += '热力购入量计算器 - 测试报告\n';
  report += '='.repeat(70) + '\n';
  report += `生成时间: ${new Date().toLocaleString('zh-CN')}\n`;
  report += `测试耗时: ${((Date.now() - startTime) / 1000).toFixed(2)} 秒\n`;
  report += '\n';
  report += '📊 测试统计:\n';
  report += `- 总用例数: ${total}\n`;
  report += `- 通过: ${passed}\n`;
  report += `- 失败: ${failed}\n`;
  report += `- 通过率: ${((passed / total) * 100).toFixed(1)}%\n`;
  report += '\n';
  
  // 按类别统计
  const categories = ['hot_water', 'saturated_steam', 'superheated_steam'];
  const catLabels = { hot_water: '热水', saturated_steam: '饱和蒸汽', superheated_steam: '过热蒸汽' };
  
  report += '📋 分类统计:\n';
  for (const cat of categories) {
    const catResults = results.filter(r => {
      if (!r.input || !r.input.mediumType) return false;
      return r.input.mediumType === cat;
    });
    if (catResults.length > 0) {
      const catPassed = catResults.filter(r => r.passed).length;
      report += `- ${catLabels[cat]}: ${catPassed}/${catResults.length} 通过\n`;
    }
  }
  report += '\n';
  
  // 详细结果
  report += '📝 详细结果:\n';
  report += '-'.repeat(70) + '\n';
  
  for (const result of results) {
    const status = result.passed ? '✅' : '❌';
    report += `${status} [${result.id}] ${result.name}\n`;
    report += `   输入: 介质=${result.input.mediumType}, 重量=${result.input.weight}吨`;
    if (result.input.temp !== null && result.input.temp !== undefined) {
      report += `, 温度=${result.input.temp}°C`;
    }
    if (result.input.pressure !== null && result.input.pressure !== undefined) {
      report += `, 压力=${result.input.pressure}MPa`;
    }
    report += '\n';
    report += `   预期: ${result.expected.success ? '成功' : '失败'}`;
    if (result.expected.minGJ !== undefined) {
      report += `, GJ范围=[${result.expected.minGJ}, ${result.expected.maxGJ}]`;
    }
    report += '\n';
    report += `   实际: ${result.message}\n`;
    
    if (result.actualGJ !== null) {
      report += `   计算值: ${result.actualGJ.toFixed(4)} GJ\n`;
    }
    
    report += '\n';
  }
  
  // 失败详情
  const failedResults = results.filter(r => !r.passed);
  if (failedResults.length > 0) {
    report += '⚠️  失败用例详情:\n';
    report += '-'.repeat(70) + '\n';
    for (const result of failedResults) {
      report += `\n❌ ${result.id}: ${result.name}\n`;
      report += `   ${result.message}\n`;
      if (result.error) {
        report += `   错误: ${result.error}\n`;
      }
      if (result.actualGJ !== null) {
        report += `   实际GJ: ${result.actualGJ.toFixed(4)}\n`;
      }
      if (result.expected.minGJ !== undefined) {
        report += `   预期范围: [${result.expected.minGJ}, ${result.expected.maxGJ}]\n`;
      }
    }
  }
  
  report += '\n' + '='.repeat(70) + '\n';
  report += `测试结论: ${failed === 0 ? '🎉 全部通过!' : `⚠️  ${failed} 个用例失败`}\n`;
  report += '='.repeat(70) + '\n';
  
  return report;
}

// 主程序
function main() {
  startTime = Date.now();
  
  // 运行测试
  const testResults = runAllTests();
  
  // 生成报告
  const report = generateReport(testResults);
  
  // 保存报告
  const reportPath = path.join(__dirname, 'test-report.txt');
  fs.writeFileSync(reportPath, report, 'utf-8');
  console.log(`📄 报告已保存到: ${reportPath}`);
  
  // 在控制台输出报告摘要
  const total = testResults.length;
  const passed = testResults.filter(r => r.passed).length;
  const failed = total - passed;
  
  console.log('');
  console.log('='.repeat(60));
  console.log(`📊 测试完成: ${passed}/${total} 通过 (${failed} 失败)`);
  console.log('='.repeat(60));
  
  // 输出失败用例
  if (failed > 0) {
    console.log('');
    console.log('失败用例列表:');
    for (const result of testResults.filter(r => !r.passed)) {
      console.log(`  ❌ ${result.id}: ${result.name}`);
      console.log(`     ${result.message}`);
    }
    // 返回非零退出码
    process.exit(1);
  }
  
  process.exit(0);
}

// 执行
main();