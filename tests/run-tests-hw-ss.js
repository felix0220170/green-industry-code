// 专项测试运行器 - 热水与饱和蒸汽验证
const fs = require('fs');
const path = require('path');
const testCases = require('./test-cases-hw-ss');

let results = [];
let startTime = null;

function loadCalculator() {
  const sourcePath = path.join(__dirname, '../ThermalCarbonCalc/js/thermal-carbon-calc.js');
  let sourceCode = fs.readFileSync(sourcePath, 'utf-8');
  
  const sandbox = {
    console: { log: () => {}, error: () => {}, warn: () => {} },
    window: { addEventListener: () => {} },
    document: {
      addEventListener: (event, callback) => {
        if (event === 'DOMContentLoaded') callback();
      },
      getElementById: () => ({ 
        addEventListener: () => {}, appendChild: () => {}, setAttribute: () => {},
        style: {}, classList: { add: () => {}, remove: () => {} }
      }),
      querySelector: () => ({ addEventListener: () => {} }),
      querySelectorAll: () => []
    },
    alert: () => {},
    confirm: () => true,
    setTimeout: () => {}
  };

  const wrappedCode = `
    var _testEnv = {};
    
    ${sourceCode}
    
    if (!dataLoaded) {
      loadData();
    }
    if (refEnthalpy === null) {
      refEnthalpy = getSaturatedWaterEnthalpy(REF_TEMP);
    }
    if (pressureToTempMap === null) {
      pressureToTempMap = buildPressureToTempMap();
    }
    
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
    
    result.actualGJ = calcResult.value;
    result.details = calcResult.details;
    
    if (testCase.expected.success === false) {
      result.passed = false;
      result.message = '预期失败，但计算成功';
      return result;
    }
    
    let allChecksPassed = true;
    let checkMessages = [];
    
    // 1. Check GJ range (minGJ/maxGJ)
    if (testCase.expected.minGJ !== undefined && testCase.expected.maxGJ !== undefined) {
      if (calcResult.value >= testCase.expected.minGJ && calcResult.value <= testCase.expected.maxGJ) {
        checkMessages.push(`GJ=${calcResult.value.toFixed(5)} 在范围 [${testCase.expected.minGJ}, ${testCase.expected.maxGJ}] 内`);
      } else {
        allChecksPassed = false;
        checkMessages.push(`GJ=${calcResult.value.toFixed(5)} 超出范围 [${testCase.expected.minGJ}, ${testCase.expected.maxGJ}]`);
      }
    }
    
    // 2. Check exact value (exactGJ + tolerance)
    if (testCase.expected.exactGJ !== undefined) {
      const tolerance = testCase.expected.tolerance || 0.01;
      const diff = Math.abs(calcResult.value - testCase.expected.exactGJ);
      if (diff <= tolerance) {
        checkMessages.push(`GJ=${calcResult.value.toFixed(5)} 与预期值 ${testCase.expected.exactGJ} 的偏差=${diff.toFixed(5)} ≤ ${tolerance}`);
      } else {
        allChecksPassed = false;
        checkMessages.push(`GJ=${calcResult.value.toFixed(5)} 与预期值 ${testCase.expected.exactGJ} 的偏差=${diff.toFixed(5)} > ${tolerance}`);
      }
    }
    
    if (testCase.expected.description) {
      checkMessages.push(`说明：${testCase.expected.description}`);
    }
    
    result.passed = allChecksPassed;
    result.message = checkMessages.join(' | ');
    
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

function runAllTests() {
  console.log('='.repeat(60));
  console.log('热力购入量计算器 - 专项测试套件（热水 & 饱和蒸汽）');
  console.log('='.repeat(60));
  console.log('');
  
  console.log('[1/3] 加载计算器模块...');
  const calc = loadCalculator();
  if (!calc) {
    console.error('❌ 加载失败，退出测试');
    return results;
  }
  console.log('✅ 加载成功');
  console.log('');
  
  console.log('[2/3] 运行测试用例...');
  console.log('');
  
  const categories = [
    { key: 'hot_water', label: '一、热水专项测试' },
    { key: 'saturated_steam', label: '二、饱和蒸汽专项测试' }
  ];
  
  for (const cat of categories) {
    const cases = testCases[cat.key];
    if (!cases || cases.length === 0) continue;
    
    console.log(`\n📋 ${cat.label} (${cases.length} 个用例):`);
    console.log('-'.repeat(60));
    
    let catPassed = 0;
    let catFailed = 0;
    
    for (const testCase of cases) {
      const result = runTestCase(calc, testCase);
      results.push(result);
      
      if (result.passed) {
        catPassed++;
        console.log(`  ✅ ${result.id}: ${result.name}`);
        console.log(`     ${result.message}`);
        if (result.actualGJ !== null) {
          console.log(`     GJ=${result.actualGJ.toFixed(5)}`);
        }
      } else {
        catFailed++;
        console.log(`  ❌ ${result.id}: ${result.name}`);
        console.log(`     ${result.message}`);
        if (result.actualGJ !== null) {
          console.log(`     GJ=${result.actualGJ.toFixed(5)}`);
        }
        if (result.error) {
          console.log(`     错误: ${result.error}`);
        }
      }
    }
    
    console.log(`  小计: ${catPassed}/${cases.length} 通过`);
  }
  
  console.log('');
  console.log('[3/3] 生成报告...');
  
  return results;
}

function generateReport(results) {
  const total = results.length;
  const passed = results.filter(r => r.passed).length;
  const failed = total - passed;
  
  let report = '';
  report += '='.repeat(70) + '\n';
  report += '热力购入量计算器 - 专项测试报告（热水 & 饱和蒸汽）\n';
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
  
  const categories = [
    { key: 'hot_water', label: '热水专项测试' },
    { key: 'saturated_steam', label: '饱和蒸汽专项测试' }
  ];
  
  report += '📋 分类统计:\n';
  for (const cat of categories) {
    const catResults = results.filter(r => {
      if (!r.id) return false;
      if (r.id.startsWith('HW-SP-')) return cat.key === 'hot_water';
      if (r.id.startsWith('SS-SP-')) return cat.key === 'saturated_steam';
      return false;
    });
    if (catResults.length > 0) {
      const catPassed = catResults.filter(r => r.passed).length;
      report += `- ${cat.label}: ${catPassed}/${catResults.length} 通过\n`;
    }
  }
  report += '\n';
  
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
    
    if (result.actualGJ !== null) {
      report += `   计算值: ${result.actualGJ.toFixed(5)} GJ\n`;
    }
    
    report += `   检查结果: ${result.message}\n`;
    report += '\n';
  }
  
  const failedResults = results.filter(r => !r.passed);
  if (failedResults.length > 0) {
    report += '⚠️  失败用例详情:\n';
    report += '-'.repeat(70) + '\n';
    for (const result of failedResults) {
      report += `\n❌ ${result.id}: ${result.name}\n`;
      report += `   ${result.message}\n`;
      if (result.actualGJ !== null) {
        report += `   实际GJ: ${result.actualGJ.toFixed(5)}\n`;
      }
    }
  }
  
  report += '\n' + '='.repeat(70) + '\n';
  report += `测试结论: ${failed === 0 ? '🎉 全部通过!' : `⚠️  ${failed} 个用例失败`}\n`;
  report += '='.repeat(70) + '\n';
  
  return report;
}

function main() {
  startTime = Date.now();
  
  const testResults = runAllTests();
  const report = generateReport(testResults);
  
  const reportPath = path.join(__dirname, 'test-report-hw-ss.txt');
  fs.writeFileSync(reportPath, report, 'utf-8');
  console.log(`📄 报告已保存到: ${reportPath}`);
  
  const total = testResults.length;
  const passed = testResults.filter(r => r.passed).length;
  const failed = total - passed;
  
  console.log('');
  console.log('='.repeat(60));
  console.log(`📊 测试完成: ${passed}/${total} 通过 (${failed} 失败)`);
  console.log('='.repeat(60));
  
  if (failed > 0) {
    console.log('');
    console.log('失败用例列表:');
    for (const result of testResults.filter(r => !r.passed)) {
      console.log(`  ❌ ${result.id}: ${result.name}`);
      console.log(`     ${result.message}`);
    }
    process.exit(1);
  }
  
  process.exit(0);
}

main();