// 专项测试用例 - 热水与饱和蒸汽验证
// 补充测试用例集

const testCases = {
  // ========== 一、热水专项测试（11个用例） ==========
  hot_water: [
    {
      id: 'HW-SP-01',
      name: '热水 - 0°C（基准温度边界）',
      input: { mediumType: 'hot_water', weight: 1, temp: 0, pressure: null },
      expected: { 
        success: true, 
        minGJ: -0.10, 
        maxGJ: 0.01,
        description: '0°C时焓值低于基准温度(20°C)，GJ为负值'
      }
    },
    {
      id: 'HW-SP-02',
      name: '热水 - 20°C（基准温度）',
      input: { mediumType: 'hot_water', weight: 1, temp: 20, pressure: null },
      expected: { 
        success: true, 
        minGJ: -0.01, 
        maxGJ: 0.01,
        description: '20°C等于基准温度，焓差为0，GJ应为0'
      }
    },
    {
      id: 'HW-SP-03',
      name: '热水 - 50°C（中间值）',
      input: { mediumType: 'hot_water', weight: 1, temp: 50, pressure: null },
      expected: { 
        success: true, 
        minGJ: 0.12, 
        maxGJ: 0.13,
        description: '50°C时h_liquid=209.34kJ/kg，GJ=(209.34-83.92)*1000/1000000≈0.1254'
      }
    },
    {
      id: 'HW-SP-04',
      name: '热水 - 90°C（工业常见）',
      input: { mediumType: 'hot_water', weight: 1, temp: 90, pressure: null },
      expected: { 
        success: true, 
        minGJ: 0.28, 
        maxGJ: 0.31,
        description: '90°C时GJ≈0.29305'
      }
    },
    {
      id: 'HW-SP-05',
      name: '热水 - 100°C（沸点）',
      input: { mediumType: 'hot_water', weight: 1, temp: 100, pressure: null },
      expected: { 
        success: true, 
        minGJ: 0.33, 
        maxGJ: 0.36,
        description: '100°C时GJ≈0.33518'
      }
    },
    {
      id: 'HW-SP-06',
      name: '热水 - 374°C（临界点）',
      input: { mediumType: 'hot_water', weight: 1, temp: 374, pressure: null },
      expected: { 
        success: true, 
        minGJ: 1.99, 
        maxGJ: 2.02,
        description: '临界温度374°C时GJ≈2.00368'
      }
    },
    {
      id: 'HW-SP-07',
      name: '热水 - 85°C（非网格点）',
      input: { mediumType: 'hot_water', weight: 1, temp: 85, pressure: null },
      expected: { 
        success: true, 
        minGJ: 0.27, 
        maxGJ: 0.28,
        description: '85°C位于80-90°C之间，线性插值GJ≈0.27203'
      }
    },
    {
      id: 'HW-SP-08',
      name: '热水 - 10吨, 90°C（重量线性）',
      input: { mediumType: 'hot_water', weight: 10, temp: 90, pressure: null },
      expected: { 
        success: true, 
        exactGJ: 2.93050,
        tolerance: 0.01,
        description: '10吨时GJ应为1吨的10倍（0.29305 * 10 = 2.9305）'
      }
    },
    {
      id: 'HW-SP-09',
      name: '热水 - 0.5吨, 90°C（重量线性）',
      input: { mediumType: 'hot_water', weight: 0.5, temp: 90, pressure: null },
      expected: { 
        success: true, 
        exactGJ: 0.14653,
        tolerance: 0.01,
        description: '0.5吨时GJ应为1吨的0.5倍（0.29305 * 0.5 = 0.14653）'
      }
    },
    {
      id: 'HW-SP-10',
      name: '热水 - -1°C（越界）',
      input: { mediumType: 'hot_water', weight: 1, temp: -1, pressure: null },
      expected: { 
        success: false,
        description: '温度低于0°C，应验证失败'
      }
    },
    {
      id: 'HW-SP-11',
      name: '热水 - 375°C（越界）',
      input: { mediumType: 'hot_water', weight: 1, temp: 375, pressure: null },
      expected: { 
        success: false,
        description: '温度高于374°C，应验证失败'
      }
    },
    {
      id: 'HW-SP-12',
      name: '热水 - 150°C（高温）',
      input: { mediumType: 'hot_water', weight: 1, temp: 150, pressure: null },
      expected: { 
        success: true, 
        minGJ: 0.54, 
        maxGJ: 0.56,
        description: '150°C时GJ≈0.54833'
      }
    },
    {
      id: 'HW-SP-13',
      name: '热水 - 200°C（高温）',
      input: { mediumType: 'hot_water', weight: 1, temp: 200, pressure: null },
      expected: { 
        success: true, 
        minGJ: 0.76, 
        maxGJ: 0.78,
        description: '200°C时GJ≈0.76847'
      }
    }
  ],

  // ========== 二、饱和蒸汽专项测试（13个用例） ==========
  saturated_steam: [
    {
      id: 'SS-SP-01',
      name: '饱和蒸汽 - 0.0006 MPa（最小边界）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 0.0006 },
      expected: { 
        success: true, 
        minGJ: 2.40, 
        maxGJ: 2.44,
        description: '0.0006MPa时GJ≈2.41699'
      }
    },
    {
      id: 'SS-SP-02',
      name: '饱和蒸汽 - 0.1 MPa（低压）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 0.1 },
      expected: { 
        success: true, 
        minGJ: 2.57, 
        maxGJ: 2.61,
        description: '0.1MPa时GJ≈2.59102'
      }
    },
    {
      id: 'SS-SP-03',
      name: '饱和蒸汽 - 0.5 MPa（中低压）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 0.5 },
      expected: { 
        success: true, 
        minGJ: 2.64, 
        maxGJ: 2.68,
        description: '0.5MPa时GJ≈2.66419'
      }
    },
    {
      id: 'SS-SP-04',
      name: '饱和蒸汽 - 1 MPa（中压）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 1 },
      expected: { 
        success: true, 
        minGJ: 2.67, 
        maxGJ: 2.71,
        description: '1MPa时GJ≈2.69320'
      }
    },
    {
      id: 'SS-SP-05',
      name: '饱和蒸汽 - 3 MPa（峰值附近）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 3 },
      expected: { 
        success: true, 
        minGJ: 2.70, 
        maxGJ: 2.74,
        description: '3MPa时GJ≈2.71934，位于焓值峰值附近'
      }
    },
    {
      id: 'SS-SP-06',
      name: '饱和蒸汽 - 10 MPa（高压）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 10 },
      expected: { 
        success: true, 
        minGJ: 2.62, 
        maxGJ: 2.66,
        description: '10MPa时GJ≈2.64155'
      }
    },
    {
      id: 'SS-SP-07',
      name: '饱和蒸汽 - 22.064 MPa（临界点）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 22.064 },
      expected: { 
        success: true, 
        minGJ: 1.99, 
        maxGJ: 2.02,
        description: '临界点22.064MPa时GJ≈2.00368'
      }
    },
    {
      id: 'SS-SP-08',
      name: '饱和蒸汽 - 0.05 MPa（非网格点）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 0.05 },
      expected: { 
        success: true, 
        minGJ: 2.55, 
        maxGJ: 2.57,
        description: '0.05MPa位于0.0006-0.1MPa之间，线性插值GJ≈2.56128'
      }
    },
    {
      id: 'SS-SP-09',
      name: '饱和蒸汽 - 5.5 MPa（非网格点）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 5.5 },
      expected: { 
        success: true, 
        minGJ: 2.70, 
        maxGJ: 2.72,
        description: '5.5MPa位于5-6MPa之间，线性插值GJ≈2.70580'
      }
    },
    {
      id: 'SS-SP-10',
      name: '饱和蒸汽 - 10吨, 1 MPa（重量线性）',
      input: { mediumType: 'saturated_steam', weight: 10, temp: null, pressure: 1 },
      expected: { 
        success: true, 
        exactGJ: 26.93199,
        tolerance: 0.01,
        description: '10吨时GJ应为1吨的10倍（2.69320 * 10 = 26.932）'
      }
    },
    {
      id: 'SS-SP-11',
      name: '饱和蒸汽 - 0.5吨, 1 MPa（重量线性）',
      input: { mediumType: 'saturated_steam', weight: 0.5, temp: null, pressure: 1 },
      expected: { 
        success: true, 
        exactGJ: 1.34660,
        tolerance: 0.01,
        description: '0.5吨时GJ应为1吨的0.5倍（2.69320 * 0.5 = 1.34660）'
      }
    },
    {
      id: 'SS-SP-12',
      name: '饱和蒸汽 - 0.0005 MPa（越界）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 0.0005 },
      expected: { 
        success: false,
        description: '压力低于0.0006MPa最小值，应验证失败'
      }
    },
    {
      id: 'SS-SP-13',
      name: '饱和蒸汽 - 23 MPa（越界）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 23 },
      expected: { 
        success: false,
        description: '压力高于22.064MPa最大值，应验证失败'
      }
    }
  ]
};

module.exports = testCases;