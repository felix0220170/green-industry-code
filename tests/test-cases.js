// 测试用例数据定义
// 包含各种场景的测试用例

const testCases = {
  // ========== 热水测试 ==========
  hot_water: [
    {
      id: 'HW-001',
      name: '热水 - 温度0°C（边界值，低于基准温度）',
      input: { mediumType: 'hot_water', weight: 1, temp: 0, pressure: null },
      expected: { success: true, minGJ: -0.1, maxGJ: 0 }
    },
    {
      id: 'HW-002',
      name: '热水 - 温度100°C（中间值）',
      input: { mediumType: 'hot_water', weight: 1, temp: 100, pressure: null },
      expected: { success: true, minGJ: 0.3, maxGJ: 0.5 }
    },
    {
      id: 'HW-003',
      name: '热水 - 温度374°C（临界温度）',
      input: { mediumType: 'hot_water', weight: 1, temp: 374, pressure: null },
      expected: { success: true, minGJ: 1.8, maxGJ: 2.2 }
    },
    {
      id: 'HW-004',
      name: '热水 - 温度-1°C（越界值）',
      input: { mediumType: 'hot_water', weight: 1, temp: -1, pressure: null },
      expected: { success: false }
    },
    {
      id: 'HW-005',
      name: '热水 - 温度400°C（越界值）',
      input: { mediumType: 'hot_water', weight: 1, temp: 400, pressure: null },
      expected: { success: false }
    }
  ],

  // ========== 饱和蒸汽测试 ==========
  saturated_steam: [
    {
      id: 'SS-001',
      name: '饱和蒸汽 - 压力0.0006 MPa（最小边界）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 0.0006 },
      expected: { success: true, minGJ: 2.3, maxGJ: 2.5 }
    },
    {
      id: 'SS-002',
      name: '饱和蒸汽 - 压力0.1 MPa（低压）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 0.1 },
      expected: { success: true, minGJ: 2.4, maxGJ: 2.6 }
    },
    {
      id: 'SS-003',
      name: '饱和蒸汽 - 压力1 MPa（中间值）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 1 },
      expected: { success: true, minGJ: 2.6, maxGJ: 2.8 }
    },
    {
      id: 'SS-004',
      name: '饱和蒸汽 - 压力10 MPa（高压）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 10 },
      expected: { success: true, minGJ: 2.4, maxGJ: 2.8 }
    },
    {
      id: 'SS-005',
      name: '饱和蒸汽 - 压力22.064 MPa（临界点）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 22.064 },
      expected: { success: true, minGJ: 1.8, maxGJ: 2.2 }
    },
    {
      id: 'SS-006',
      name: '饱和蒸汽 - 压力0.0005 MPa（越界值）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 0.0005 },
      expected: { success: false }
    },
    {
      id: 'SS-007',
      name: '饱和蒸汽 - 压力23 MPa（越界值）',
      input: { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 23 },
      expected: { success: false }
    }
  ],

  // ========== 过热蒸汽测试 ==========
  superheated_steam: [
    {
      id: 'SH-001',
      name: '过热蒸汽 - 1MPa, 200°C（低压低温）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 200, pressure: 1 },
      expected: { success: true, minGJ: 2.6, maxGJ: 2.9 }
    },
    {
      id: 'SH-002',
      name: '过热蒸汽 - 1MPa, 500°C（中温中压）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 500, pressure: 1 },
      expected: { success: true, minGJ: 3.2, maxGJ: 3.6 }
    },
    {
      id: 'SH-003',
      name: '过热蒸汽 - 1MPa, 800°C（高温）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 800, pressure: 1 },
      expected: { success: true, minGJ: 3.8, maxGJ: 4.3 }
    },
    {
      id: 'SH-004',
      name: '过热蒸汽 - 10MPa, 312°C（刚过饱和温度）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 312, pressure: 10 },
      expected: { success: true, minGJ: 2.4, maxGJ: 2.8 }
    },
    {
      id: 'SH-005',
      name: '过热蒸汽 - 10MPa, 400°C（用户反馈问题场景）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 400, pressure: 10 },
      expected: { success: true, minGJ: 2.8, maxGJ: 3.2 }
    },
    {
      id: 'SH-006',
      name: '过热蒸汽 - 10MPa, 500°C（中温高压）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 500, pressure: 10 },
      expected: { success: true, minGJ: 3.0, maxGJ: 3.6 }
    },
    {
      id: 'SH-007',
      name: '过热蒸汽 - 10MPa, 200°C（低于饱和温度）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 200, pressure: 10 },
      expected: { success: false }
    },
    {
      id: 'SH-008',
      name: '过热蒸汽 - 10MPa, 311°C（等于饱和温度，自动切换）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 311, pressure: 10 },
      expected: { success: true, minGJ: 2.4, maxGJ: 2.8 }
    },
    {
      id: 'SH-009',
      name: '过热蒸汽 - 0.1MPa, 150°C（低压过热）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 150, pressure: 0.1 },
      expected: { success: true, minGJ: 2.6, maxGJ: 2.9 }
    },
    {
      id: 'SH-010',
      name: '过热蒸汽 - 0.1MPa, 80°C（低于饱和温度）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 80, pressure: 0.1 },
      expected: { success: false }
    },
    {
      id: 'SH-011',
      name: '过热蒸汽 - 0.001MPa, 100°C（最低压力边界）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 100, pressure: 0.001 },
      expected: { success: true, minGJ: 2.4, maxGJ: 2.8 }
    },
    {
      id: 'SH-012',
      name: '过热蒸汽 - 100MPa, 800°C（最高压力边界）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 800, pressure: 100 },
      expected: { success: true, minGJ: 3.5, maxGJ: 4.5 }
    },
    {
      id: 'SH-013',
      name: '过热蒸汽 - 压力越界（101MPa）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 500, pressure: 101 },
      expected: { success: false }
    },
    {
      id: 'SH-014',
      name: '过热蒸汽 - 温度越界（801°C）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 801, pressure: 10 },
      expected: { success: false }
    },
    {
      id: 'SH-015',
      name: '过热蒸汽 - 非网格点温度验证（10MPa, 395°C）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 395, pressure: 10 },
      expected: { success: true, minGJ: 2.6, maxGJ: 3.4 }
    },
    {
      id: 'SH-016',
      name: '过热蒸汽 - 非网格点温度验证（5MPa, 450°C）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 450, pressure: 5 },
      expected: { success: true, minGJ: 2.6, maxGJ: 3.4 }
    },
    {
      id: 'SH-017',
      name: '过热蒸汽 - 压力10MPa, 600°C（验证温度区间查找）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 600, pressure: 10 },
      expected: { success: true, minGJ: 3.2, maxGJ: 3.8 }
    },
    {
      id: 'SH-018',
      name: '过热蒸汽 - 压力10MPa, 700°C（验证高温插值）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 700, pressure: 10 },
      expected: { success: true, minGJ: 3.6, maxGJ: 4.2 }
    }
  ],

  // ========== 总 GJ 计算测试 ==========
  total_calc: [
    {
      id: 'TOTAL-001',
      name: '总GJ - 多介质混合计算',
      input: {
        items: [
          { mediumType: 'hot_water', weight: 1, temp: 90, pressure: null },
          { mediumType: 'saturated_steam', weight: 1, temp: null, pressure: 1 },
          { mediumType: 'superheated_steam', weight: 1, temp: 350, pressure: 5 }
        ]
      },
      expected: { success: true, minGJ: 6, maxGJ: 8 }
    }
  ]
};

module.exports = testCases;