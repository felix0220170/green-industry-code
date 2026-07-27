// 整合测试用例 - 热力购入量计算器
// 包含基础功能、修正逻辑、角点验证等所有测试场景

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

  // ========== 四、修正逻辑验证（5个用例） ==========
  // 验证：当温度接近饱和温度时，使用饱和温度作为下界
  correction_logic: [
    {
      id: 'COR-001',
      name: '修正逻辑 - 0.5MPa, 155°C（刚过饱和温度151.83°C）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 155, pressure: 0.5 },
      expected: { 
        success: true, 
        minGJ: 2.57, 
        maxGJ: 2.77,
        expectedInterval: { t1: 151.8, t2: 160, description: '使用饱和温度作为下界' }
      }
    },
    {
      id: 'COR-002',
      name: '修正逻辑 - 2MPa, 215°C（刚过饱和温度212.38°C）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 215, pressure: 2 },
      expected: { 
        success: true, 
        minGJ: 2.62, 
        maxGJ: 2.82,
        expectedInterval: { t1: 212.4, t2: 220, description: '使用饱和温度作为下界' }
      }
    },
    {
      id: 'COR-003',
      name: '修正逻辑 - 5MPa, 265°C（刚过饱和温度263.94°C）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 265, pressure: 5 },
      expected: { 
        success: true, 
        minGJ: 2.61, 
        maxGJ: 2.81,
        expectedInterval: { t1: 263.9, t2: 270, description: '使用饱和温度作为下界' }
      }
    },
    {
      id: 'COR-004',
      name: '修正逻辑 - 15MPa, 345°C（刚过饱和温度342.16°C）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 345, pressure: 15 },
      expected: { 
        success: true, 
        minGJ: 2.46, 
        maxGJ: 2.66,
        expectedInterval: { t1: 342.1, t2: 350, description: '使用饱和温度作为下界' }
      }
    },
    {
      id: 'COR-005',
      name: '修正逻辑 - 20MPa, 368°C（刚过饱和温度365.74°C）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 368, pressure: 20 },
      expected: { 
        success: true, 
        minGJ: 2.29, 
        maxGJ: 2.49,
        expectedInterval: { t1: 365.7, t2: 370, description: '使用饱和温度作为下界' }
      }
    }
  ],

  // ========== 五、边界组合测试（4个用例） ==========
  // 覆盖 P/T 在或不在网格点的各种组合
  boundary_combinations: [
    {
      id: 'BOUND-001',
      name: '边界组合 - P=0.5, T=159（压力在网格点，温度不在）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 159, pressure: 0.5 },
      expected: { 
        success: true, 
        minGJ: 2.58, 
        maxGJ: 2.78,
        expectedInterval: { t1: 151.8, t2: 160, description: '使用饱和温度作为下界' }
      }
    },
    {
      id: 'BOUND-002',
      name: '边界组合 - P=0.5, T=160（压力和温度都在网格点）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 160, pressure: 0.5 },
      expected: { 
        success: true, 
        minGJ: 2.58, 
        maxGJ: 2.78,
        expectedInterval: { t1: 151.8, t2: 160, description: '使用饱和温度作为下界' }
      }
    },
    {
      id: 'BOUND-003',
      name: '边界组合 - P=0.4, T=159（压力和温度都不在网格点，P=0.5角点需饱和修正）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 159, pressure: 0.4 },
      expected: { 
        success: true, 
        minGJ: 2.59, 
        maxGJ: 2.79,
        expectedInterval: { t1: 150, t2: 160, description: 'P=0.5行角点需饱和修正' }
      }
    },
    {
      id: 'BOUND-004',
      name: '边界组合 - P=0.4, T=160（压力不在网格点，温度在网格点，P=0.5角点需饱和修正）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 160, pressure: 0.4 },
      expected: { 
        success: true, 
        minGJ: 2.59, 
        maxGJ: 2.79,
        expectedInterval: { t1: 150, t2: 160, description: 'P=0.5行角点需饱和修正' }
      }
    }
  ],

  // ========== 六、角点焓值验证（4个用例） ==========
  // 验证：插值网格表格中不再出现过冷水焓值
  corner_enthalpy: [
    {
      id: 'CORNER-001',
      name: '角点验证 - P=0.4, T=159（验证角点焓值为过热/饱和蒸汽）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 159, pressure: 0.4 },
      expected: { 
        success: true, 
        minGJ: 2.59, 
        maxGJ: 2.79,
        cornerCheck: {
          noLiquidEnthalpy: true,
          description: '所有角点焓值应为过热蒸汽或饱和蒸汽，不应为过冷水'
        }
      }
    },
    {
      id: 'CORNER-002',
      name: '角点验证 - P=0.4, T=160（验证角点焓值为过热/饱和蒸汽）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 160, pressure: 0.4 },
      expected: { 
        success: true, 
        minGJ: 2.59, 
        maxGJ: 2.79,
        cornerCheck: {
          noLiquidEnthalpy: true,
          description: '所有角点焓值应为过热蒸汽或饱和蒸汽，不应为过冷水'
        }
      }
    },
    {
      id: 'CORNER-003',
      name: '角点验证 - P=0.5, T=159（验证角点焓值为过热/饱和蒸汽）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 159, pressure: 0.5 },
      expected: { 
        success: true, 
        minGJ: 2.58, 
        maxGJ: 2.78,
        cornerCheck: {
          noLiquidEnthalpy: true,
          description: '所有角点焓值应为过热蒸汽或饱和蒸汽，不应为过冷水'
        }
      }
    },
    {
      id: 'CORNER-004',
      name: '角点验证 - P=0.5, T=160（验证角点焓值为过热/饱和蒸汽）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 160, pressure: 0.5 },
      expected: { 
        success: true, 
        minGJ: 2.58, 
        maxGJ: 2.78,
        cornerCheck: {
          noLiquidEnthalpy: true,
          description: '所有角点焓值应为过热蒸汽或饱和蒸汽，不应为过冷水'
        }
      }
    }
  ],

  // ========== 七、重量线性验证（3个用例） ==========
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

  // ========== 八、显示逻辑验证（3个用例） ==========
  // 验证：插值过程显示正确、饱和/过热标签正确
  display_logic: [
    {
      id: 'DISPLAY-001',
      name: '显示验证 - P=0.4, T=159（验证插值过程显示）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 159, pressure: 0.4 },
      expected: { 
        success: true, 
        minGJ: 2.59, 
        maxGJ: 2.79,
        displayCheck: {
          hasInterpolationSteps: true,
          hasSaturatedLabel: true,
          description: '应显示插值步骤和饱和标签'
        }
      }
    },
    {
      id: 'DISPLAY-002',
      name: '显示验证 - P=0.4, T=160（验证网格点命中显示）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 160, pressure: 0.4 },
      expected: { 
        success: true, 
        minGJ: 2.59, 
        maxGJ: 2.79,
        displayCheck: {
          hasInterpolationSteps: true,
          hasSaturatedLabel: true,
          description: '应显示插值步骤和饱和标签'
        }
      }
    },
    {
      id: 'DISPLAY-003',
      name: '显示验证 - P=1.0, T=180（验证高压饱和温度显示）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 180, pressure: 1.0 },
      expected: { 
        success: true, 
        minGJ: 2.62, 
        maxGJ: 2.82,
        displayCheck: {
          hasInterpolationSteps: true,
          hasSaturatedLabel: true,
          description: '应显示插值步骤和饱和标签，饱和温度约179.88°C'
        }
      }
    }
  ]
};

module.exports = testCases;
