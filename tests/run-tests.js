// 整合测试运行器 - 热力购入量计算器
// 支持基础功能、修正逻辑、角点验证、显示逻辑等所有测试

const fs = require('fs');
const path = require('path');
const testCases = require('./test-cases');

let results = [];
let startTime = null;

// 从源文件中提取核心计算函数
function loadCalculator() {
  const sourcePath = path.join(__dirname, '../ThermalCarbonCalc/js/thermal-carbon-calc.js');
  let sourceCode = fs.readFileSync(sourcePath, 'utf-8');
  
  const sandbox = {
    console: {
      log: () => {},
      error: (msg) => {},
      warn: () => {}
    },
    window: { addEventListener: () => {} },
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

// 解析计算过程中的温度区间信息
function extractIntervalInfo(details) {
  const info = { t1: null, t2: null, saturatedUsed: false, intervalFound: false };
  
  if (!details) return info;
  
  if (details.includes('使用饱和蒸汽焓') || details.includes('饱和温度')) {
    info.saturatedUsed = true;
  }
  
  const intervalMatch = details.match(/温度区间[：:]\s*([\d.]+)°C\s*~\s*([\d.]+)°C/);
  if (intervalMatch) {
    info.t1 = parseFloat(intervalMatch[1]);
    info.t2 = parseFloat(intervalMatch[2]);
    info.intervalFound = true;
  }
  
  return info;
}

// 检查角点焓值是否包含过冷水
// 只检查插值网格表格中的焓值（格式为 h=XX.XX kJ/kg）
function checkCornerEnthalpy(details) {
  const check = { noLiquidEnthalpy: true, foundLiquid: false, liquidValues: [], cornerValues: [] };
  
  if (!details) return check;
  
  // 提取表格中的焓值：h=XX.XX kJ/kg 格式
  const cornerEnthalpyMatches = details.match(/h=(\d+\.?\d*)\s*kJ\/kg/g) || [];
  for (const match of cornerEnthalpyMatches) {
    const value = parseFloat(match.replace('h=', ''));
    check.cornerValues.push(value);
    // 蒸汽焓值通常在 2000-3500 kJ/kg 范围内
    // 过冷水焓值通常 < 300 kJ/kg
    if (value < 300 && value > 0) {
      check.foundLiquid = true;
      check.liquidValues.push(value);
    }
  }
  
  if (check.foundLiquid) {
    check.noLiquidEnthalpy = false;
  }
  
  return check;
}

// 检查显示逻辑
function checkDisplayLogic(details) {
  const check = { 
    hasInterpolationSteps: false, 
    hasSaturatedLabel: false,
    hasSuperheatedLabel: false
  };
  
  if (!details) return check;
  
  if (details.includes('双线性插值步骤') || details.includes('沿温度方向插值') || details.includes('沿压力方向插值')) {
    check.hasInterpolationSteps = true;
  }
  
  if (details.includes('饱和') || details.includes('饱和蒸汽')) {
    check.hasSaturatedLabel = true;
  }
  
  if (details.includes('过热') || details.includes('过热蒸汽')) {
    check.hasSuperheatedLabel = true;
  }
  
  return check;
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
    details: null,
    intervalInfo: null,
    cornerCheck: null,
    displayCheck: null
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
    result.intervalInfo = extractIntervalInfo(calcResult.details);
    result.cornerCheck = checkCornerEnthalpy(calcResult.details);
    result.displayCheck = checkDisplayLogic(calcResult.details);
    
    if (testCase.expected.success === false) {
      result.passed = false;
      result.message = '预期失败，但计算成功';
      return result;
    }
    
    let allChecksPassed = true;
    let checkMessages = [];
    
    // 1. 检查 GJ 范围
    if (testCase.expected.minGJ !== undefined && testCase.expected.maxGJ !== undefined) {
      if (calcResult.value >= testCase.expected.minGJ && calcResult.value <= testCase.expected.maxGJ) {
        checkMessages.push(`GJ=${calcResult.value.toFixed(4)} 在范围 [${testCase.expected.minGJ}, ${testCase.expected.maxGJ}] 内`);
      } else {
        allChecksPassed = false;
        checkMessages.push(`GJ=${calcResult.value.toFixed(4)} 超出范围 [${testCase.expected.minGJ}, ${testCase.expected.maxGJ}]`);
      }
    }
    
    // 2. 检查精确值
    if (testCase.expected.exactGJ !== undefined) {
      const tolerance = testCase.expected.tolerance || 0.01;
      const diff = Math.abs(calcResult.value - testCase.expected.exactGJ);
      if (diff <= tolerance) {
        checkMessages.push(`GJ=${calcResult.value.toFixed(4)} 与预期值 ${testCase.expected.exactGJ} 的偏差=${diff.toFixed(4)} ≤ ${tolerance}`);
      } else {
        allChecksPassed = false;
        checkMessages.push(`GJ=${calcResult.value.toFixed(4)} 与预期值 ${testCase.expected.exactGJ} 的偏差=${diff.toFixed(4)} > ${tolerance}`);
      }
    }
    
    // 3. 检查温度区间
    if (testCase.expected.expectedInterval) {
      const expInterval = testCase.expected.expectedInterval;
      const actualInterval = result.intervalInfo;
      
      if (!actualInterval || !actualInterval.intervalFound) {
        allChecksPassed = false;
        checkMessages.push(`未能从计算详情中提取温度区间信息`);
      } else {
        if (expInterval.t1 !== undefined && actualInterval.t1 !== null) {
          const t1Diff = Math.abs(actualInterval.t1 - expInterval.t1);
          if (t1Diff > 1.0) {
            allChecksPassed = false;
            checkMessages.push(`温度区间下界 t1=${actualInterval.t1} 与预期 ${expInterval.t1} 不符（误差${t1Diff.toFixed(2)}°C）`);
          } else {
            checkMessages.push(`温度区间下界 t1=${actualInterval.t1} ✓`);
          }
        }
        
        if (expInterval.t2 !== undefined && actualInterval.t2 !== null) {
          const t2Diff = Math.abs(actualInterval.t2 - expInterval.t2);
          if (t2Diff > 1.0) {
            allChecksPassed = false;
            checkMessages.push(`温度区间上界 t2=${actualInterval.t2} 与预期 ${expInterval.t2} 不符（误差${t2Diff.toFixed(2)}°C）`);
          } else {
            checkMessages.push(`温度区间上界 t2=${actualInterval.t2} ✓`);
          }
        }
        
        const shouldUseSaturated = expInterval.description && expInterval.description.includes('使用饱和温度');
        const shouldNotUseSaturated = expInterval.description && expInterval.description.includes('不触发饱和修正');
        const hasCornerSaturation = expInterval.description && expInterval.description.includes('需饱和修正');

        if (shouldUseSaturated || hasCornerSaturation) {
          if (actualInterval.saturatedUsed) {
            checkMessages.push(`修正逻辑：使用饱和温度作为下界 ✓`);
          } else {
            allChecksPassed = false;
            checkMessages.push(`修正逻辑：预期使用饱和温度作为下界，但实际未触发`);
          }
        } else if (shouldNotUseSaturated) {
          if (actualInterval.saturatedUsed) {
            allChecksPassed = false;
            checkMessages.push(`修正逻辑：预期不触发修正，但实际触发了`);
          } else {
            checkMessages.push(`修正逻辑：未触发修正 ✓`);
          }
        }
      }
    }
    
    // 4. 检查角点焓值
    if (testCase.expected.cornerCheck) {
      if (testCase.expected.cornerCheck.noLiquidEnthalpy && result.cornerCheck) {
        if (result.cornerCheck.noLiquidEnthalpy) {
          checkMessages.push(`角点验证：所有角点焓值均为蒸汽（无过冷水） ✓`);
        } else {
          allChecksPassed = false;
          checkMessages.push(`角点验证：发现过冷水焓值 ${result.cornerCheck.liquidValues.join(', ')} kJ/kg`);
        }
      }
    }
    
    // 5. 检查显示逻辑
    if (testCase.expected.displayCheck && result.displayCheck) {
      const displayCheck = testCase.expected.displayCheck;
      const actualDisplay = result.displayCheck;
      
      if (displayCheck.hasInterpolationSteps) {
        if (actualDisplay.hasInterpolationSteps) {
          checkMessages.push(`显示验证：包含插值步骤 ✓`);
        } else {
          allChecksPassed = false;
          checkMessages.push(`显示验证：应包含插值步骤，但未找到`);
        }
      }
      
      if (displayCheck.hasSaturatedLabel) {
        if (actualDisplay.hasSaturatedLabel) {
          checkMessages.push(`显示验证：包含饱和标签 ✓`);
        } else {
          allChecksPassed = false;
          checkMessages.push(`显示验证：应包含饱和标签，但未找到`);
        }
      }
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

// 输入验证
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
    if (saturationTemp !== null && temp < saturationTemp - 0.01) {
      return { valid: false, message: `温度 ${temp.toFixed(2)}°C 低于当前压力下的饱和温度 ${saturationTemp.toFixed(2)}°C` };
    }
  }

  return { valid: true, message: '' };
}

// 运行所有测试
function runAllTests() {
  console.log('='.repeat(60));
  console.log('热力购入量计算器 - 整合测试套件');
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
    { key: 'saturated_steam', label: '二、饱和蒸汽专项测试' },
    { key: 'superheated_basic', label: '三、过热蒸汽基础测试' },
    { key: 'correction_logic', label: '四、修正逻辑验证' },
    { key: 'boundary_combinations', label: '五、边界组合测试' },
    { key: 'corner_enthalpy', label: '六、角点焓值验证' },
    { key: 'weight_linearity', label: '七、重量线性验证' },
    { key: 'display_logic', label: '八、显示逻辑验证' }
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
          console.log(`     GJ=${result.actualGJ.toFixed(4)}`);
        }
        if (result.intervalInfo && result.intervalInfo.intervalFound) {
          const intervalStr = `[${result.intervalInfo.t1}, ${result.intervalInfo.t2}]`;
          const satStr = result.intervalInfo.saturatedUsed ? ' (饱和修正)' : '';
          console.log(`     温度区间: ${intervalStr}${satStr}`);
        }
      } else {
        catFailed++;
        console.log(`  ❌ ${result.id}: ${result.name}`);
        console.log(`     ${result.message}`);
        if (result.actualGJ !== null) {
          console.log(`     GJ=${result.actualGJ.toFixed(4)}`);
        }
        if (result.intervalInfo && result.intervalInfo.intervalFound) {
          console.log(`     实际温度区间: [${result.intervalInfo.t1}, ${result.intervalInfo.t2}]`);
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

// 生成报告
function generateReport(results) {
  const total = results.length;
  const passed = results.filter(r => r.passed).length;
  const failed = total - passed;
  
  let report = '';
  report += '='.repeat(70) + '\n';
  report += '热力购入量计算器 - 整合测试报告\n';
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
    { key: 'hot_water', label: '热水测试' },
    { key: 'saturated_steam', label: '饱和蒸汽测试' },
    { key: 'superheated_basic', label: '过热蒸汽基础测试' },
    { key: 'correction_logic', label: '修正逻辑验证' },
    { key: 'boundary_combinations', label: '边界组合测试' },
    { key: 'corner_enthalpy', label: '角点焓值验证' },
    { key: 'weight_linearity', label: '重量线性验证' },
    { key: 'display_logic', label: '显示逻辑验证' }
  ];
  
  report += '📋 分类统计:\n';
  for (const cat of categories) {
    const catResults = results.filter(r => {
      if (!r.id) return false;
      return r.id.startsWith(cat.key.split('_')[0].split('-')[0]);
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
      report += `   计算值: ${result.actualGJ.toFixed(4)} GJ\n`;
    }
    
    if (result.intervalInfo && result.intervalInfo.intervalFound) {
      report += `   温度区间: [${result.intervalInfo.t1}, ${result.intervalInfo.t2}] °C`;
      if (result.intervalInfo.saturatedUsed) {
        report += ' (使用饱和温度修正)';
      }
      report += '\n';
    }
    
    if (result.cornerCheck && !result.cornerCheck.noLiquidEnthalpy) {
      report += `   ⚠️ 发现过冷水焓值: ${result.cornerCheck.liquidValues.join(', ')} kJ/kg\n`;
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
        report += `   实际GJ: ${result.actualGJ.toFixed(4)}\n`;
      }
      if (result.intervalInfo && result.intervalInfo.intervalFound) {
        report += `   实际温度区间: [${result.intervalInfo.t1}, ${result.intervalInfo.t2}] °C\n`;
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
  
  const testResults = runAllTests();
  const report = generateReport(testResults);
  
  const reportPath = path.join(__dirname, 'test-report.txt');
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
