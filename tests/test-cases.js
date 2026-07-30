// 整合测试用例 - 热力购入量计算器
// 适配简化后的代码：标准双线性插值，无修正逻辑，数据已清洗

const testCases = {
  // ========== 一、热水专项测试（8个用例） ==========
  hot_water: [
    {
      id: 'HW-001',
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
      id: 'HW-002',
      name: '热水 - 20°C（基准温度，GJ应为0）',
      input: { mediumType: 'hot_water', weight: 1, temp: 20, pressure: null },
      expected: {
        success: true,
        minGJ: -0.01,
        maxGJ: 0.01,
        description: '20°C等于基准温度，焓差为0，GJ应为0'
      }
    },
    {
      id: 'HW-003',
      name: '热水 - 90°C（工业常见温度）',
      input: { mediumType: 'hot_water', weight: 1, temp: 90, pressure: null },
      expected: {
        success: true,
        minGJ: 0.28,
        maxGJ: 0.31,
        description: '90°C时GJ≈0.293'
      }
    },
    {
      id: 'HW-004',
      name: '热水 - 100°C（沸点）',
      input: { mediumType: 'hot_water', weight: 1, temp: 100, pressure: null },
      expected: {
        success: true,
        minGJ: 0.33,
        maxGJ: 0.36,
        description: '100°C时GJ≈0.335'
      }
    },
    {
      id: 'HW-005',
      name: '热水 - 150°C（高温）',
      input: { mediumType: 'hot_water', weight: 1, temp: 150, pressure: null },
      expected: {
        success: true,
        minGJ: 0.54,
        maxGJ: 0.56,
        description: '150°C时GJ≈0.548'
      }
    },
    {
      id: 'HW-006',
      name: '热水 - 10吨, 90°C（重量线性验证）',
      input: { mediumType: 'hot_water', weight: 10, temp: 90, pressure: null },
      expected: {
        success: true,
        exactGJ: 2.93050,
        tolerance: 0.02,
        description: '10吨时GJ应为1吨的10倍'
      }
    },
    {
      id: 'HW-007',
      name: '热水 - -1°C（越界值）',
      input: { mediumType: 'hot_water', weight: 1, temp: -1, pressure: null },
      expected: { success: false }
    },
    {
      id: 'HW-008',
      name: '热水 - 375°C（越界值）',
      input: { mediumType: 'hot_water', weight: 1, temp: 375, pressure: null },
      expected: { success: false }
    }
  ],

  // ========== 二、饱和蒸汽专项测试（8个用例） ==========
  saturated_steam: [
    {
      id: 'SS-001',
      name: '饱和蒸汽 - 0.1 MPa（低压）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 0.1 },
      expected: {
        success: true,
        minGJ: 2.57,
        maxGJ: 2.61,
        description: '0.1MPa时GJ≈2.591'
      }
    },
    {
      id: 'SS-002',
      name: '饱和蒸汽 - 1 MPa（中压）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 1 },
      expected: {
        success: true,
        minGJ: 2.67,
        maxGJ: 2.71,
        description: '1MPa时GJ≈2.693'
      }
    },
    {
      id: 'SS-003',
      name: '饱和蒸汽 - 3 MPa（峰值附近）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 3 },
      expected: {
        success: true,
        minGJ: 2.70,
        maxGJ: 2.74,
        description: '3MPa时GJ≈2.719，位于焓值峰值附近'
      }
    },
    {
      id: 'SS-004',
      name: '饱和蒸汽 - 10 MPa（高压）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 10 },
      expected: {
        success: true,
        minGJ: 2.62,
        maxGJ: 2.66,
        description: '10MPa时GJ≈2.642'
      }
    },
    {
      id: 'SS-005',
      name: '饱和蒸汽 - 5.5 MPa（非网格点，验证插值）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 5.5 },
      expected: {
        success: true,
        minGJ: 2.70,
        maxGJ: 2.72,
        description: '5.5MPa位于5-6MPa之间，线性插值'
      }
    },
    {
      id: 'SS-006',
      name: '饱和蒸汽 - 10吨, 1 MPa（重量线性验证）',
      input: { mediumType: 'saturated_steam', weight: 10, temp: null, pressure: 1 },
      expected: {
        success: true,
        exactGJ: 26.932,
        tolerance: 0.02,
        description: '10吨时GJ应为1吨的10倍'
      }
    },
    {
      id: 'SS-007',
      name: '饱和蒸汽 - 0.0005 MPa（越界值）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 0.0005 },
      expected: { success: false }
    },
    {
      id: 'SS-008',
      name: '饱和蒸汽 - 23 MPa（越界值）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 23 },
      expected: { success: false }
    }
  ],

  // ========== 三、过热蒸汽基础测试（5个用例） ==========
  superheated_basic: [
    {
      id: 'SH-BASIC-001',
      name: '过热蒸汽 - 1MPa, 200°C（命中网格点）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 200, pressure: 1 },
      expected: { success: true, minGJ: 2.64, maxGJ: 2.84 }
    },
    {
      id: 'SH-BASIC-002',
      name: '过热蒸汽 - 10MPa, 400°C（中温高压）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 400, pressure: 10 },
      expected: { success: true, minGJ: 2.93, maxGJ: 3.13 }
    },
    {
      id: 'SH-BASIC-003',
      name: '过热蒸汽 - 10MPa, 200°C（低于饱和温度，应报错）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 200, pressure: 10 },
      expected: { success: false }
    },
    {
      id: 'SH-BASIC-004',
      name: '过热蒸汽 - 压力越界（101MPa）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 500, pressure: 101 },
      expected: { success: false }
    },
    {
      id: 'SH-BASIC-005',
      name: '过热蒸汽 - 温度越界（801°C）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 801, pressure: 10 },
      expected: { success: false }
    }
  ],

  // ========== 四、边界组合测试（4个用例） ==========
  // 覆盖 P/T 在或不在网格点的各种组合（适配新增压力点）
  boundary_combinations: [
    {
      id: 'BOUND-001',
      name: '边界组合 - P=1.0(网格点), T=205(非网格点)',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 205, pressure: 1.0 },
      expected: {
        success: true,
        minGJ: 2.65,
        maxGJ: 2.85,
        description: '压力命中网格点，温度不在网格点'
      }
    },
    {
      id: 'BOUND-002',
      name: '边界组合 - P=1.0(网格点), T=200(网格点)',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 200, pressure: 1.0 },
      expected: {
        success: true,
        minGJ: 2.64,
        maxGJ: 2.84,
        description: '压力和温度都命中网格点'
      }
    },
    {
      id: 'BOUND-003',
      name: '边界组合 - P=0.55(非网格点), T=205(非网格点)',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 205, pressure: 0.55 },
      expected: {
        success: true,
        minGJ: 2.68,
        maxGJ: 2.88,
        description: '压力和温度都不在网格点'
      }
    },
    {
      id: 'BOUND-004',
      name: '边界组合 - P=0.55(非网格点), T=200(网格点)',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 200, pressure: 0.55 },
      expected: {
        success: true,
        minGJ: 2.67,
        maxGJ: 2.87,
        description: '压力不在网格点，温度在网格点'
      }
    }
  ],

  // ========== 五、重量线性验证（3个用例） ==========
  weight_linearity: [
    {
      id: 'WEIGHT-001',
      name: '重量线性 - 10MPa, 400°C, 2吨（应为1吨的2倍）',
      input: { mediumType: 'superheated_steam', weight: 2, temp: 400, pressure: 10 },
      expected: {
        success: true,
        exactGJ: 6.027,
        tolerance: 0.01,
        description: '2吨时GJ应为1吨的2倍'
      }
    },
    {
      id: 'WEIGHT-002',
      name: '重量线性 - 10MPa, 400°C, 0.5吨（应为1吨的0.5倍）',
      input: { mediumType: 'superheated_steam', weight: 0.5, temp: 400, pressure: 10 },
      expected: {
        success: true,
        exactGJ: 1.50675,
        tolerance: 0.01,
        description: '0.5吨时GJ应为1吨的0.5倍'
      }
    },
    {
      id: 'WEIGHT-003',
      name: '重量线性 - 10MPa, 400°C, 100吨（应为1吨的100倍）',
      input: { mediumType: 'superheated_steam', weight: 100, temp: 400, pressure: 10 },
      expected: {
        success: true,
        exactGJ: 301.35,
        tolerance: 0.02,
        description: '100吨时GJ应为1吨的100倍'
      }
    }
  ],

  // ========== 六、新增压力点验证（5个用例） ==========
  // 验证新增的压力点（0.15, 0.25, 0.35, 0.4, 0.45, 0.6, 0.7, 0.8, 0.9, 1.5, 2.5, 3.5, 4.5, 5.5, 6.5, 7.5, 8.5, 9.5, 21）正常工作
  new_pressure_points: [
    {
      id: 'SH-NEW-01',
      name: '新增压力点 - 0.25 MPa, 140°C',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 140, pressure: 0.25 },
      expected: {
        success: true,
        minGJ: 2.56,
        maxGJ: 2.76,
        description: '0.25MPa为新压力点，饱和温度约127.4°C，140°C为过热蒸汽'
      }
    },
    {
      id: 'SH-NEW-02',
      name: '新增压力点 - 0.35 MPa, 145°C',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 145, pressure: 0.35 },
      expected: {
        success: true,
        minGJ: 2.56,
        maxGJ: 2.76,
        description: '0.35MPa为新压力点，饱和温度约138.9°C，145°C为过热蒸汽'
      }
    },
    {
      id: 'SH-NEW-03',
      name: '新增压力点 - 4.5 MPa, 300°C',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 300, pressure: 4.5 },
      expected: {
        success: true,
        minGJ: 2.76,
        maxGJ: 2.96,
        description: '4.5MPa为新压力点，饱和温度约257.4°C，300°C为过热蒸汽'
      }
    },
    {
      id: 'SH-NEW-04',
      name: '新增压力点 - 8.5 MPa, 350°C',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 350, pressure: 8.5 },
      expected: {
        success: true,
        minGJ: 2.79,
        maxGJ: 2.99,
        description: '8.5MPa为新压力点，饱和温度约299.3°C，350°C为过热蒸汽'
      }
    },
    {
      id: 'SH-NEW-05',
      name: '新增压力点 - 21 MPa, 400°C',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 400, pressure: 21 },
      expected: {
        success: true,
        minGJ: 2.59,
        maxGJ: 2.79,
        description: '21MPa为新压力点，饱和温度约369.8°C，400°C为过热蒸汽'
      }
    }
  ]
};

module.exports = testCases;
