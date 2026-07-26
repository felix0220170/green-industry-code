// 专项测试用例 - 过热蒸汽修正逻辑与温度区间查找验证
// 补充测试用例集

const testCases = {
  // ========== 一、修正逻辑专项验证（5个用例） ==========
  correction_logic: [
    {
      id: 'COR-001',
      name: '修正逻辑 - 0.5MPa, 155°C（刚过饱和温度）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 155, pressure: 0.5 },
      expected: { 
        success: true, 
        minGJ: 2.57, 
        maxGJ: 2.77,
        expectedInterval: { t1: 151.8, t2: 160, description: '使用饱和温度151.8作为下界' }
      }
    },
    {
      id: 'COR-002',
      name: '修正逻辑 - 2MPa, 215°C（刚过饱和温度）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 215, pressure: 2 },
      expected: { 
        success: true, 
        minGJ: 2.62, 
        maxGJ: 2.82,
        expectedInterval: { t1: 212.4, t2: 220, description: '使用饱和温度212.4作为下界' }
      }
    },
    {
      id: 'COR-003',
      name: '修正逻辑 - 5MPa, 265°C（刚过饱和温度）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 265, pressure: 5 },
      expected: { 
        success: true, 
        minGJ: 2.61, 
        maxGJ: 2.81,
        expectedInterval: { t1: 263.9, t2: 270, description: '使用饱和温度263.9作为下界' }
      }
    },
    {
      id: 'COR-004',
      name: '修正逻辑 - 15MPa, 345°C（刚过饱和温度）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 345, pressure: 15 },
      expected: { 
        success: true, 
        minGJ: 2.46, 
        maxGJ: 2.66,
        expectedInterval: { t1: 342.1, t2: 350, description: '使用饱和温度342.1作为下界' }
      }
    },
    {
      id: 'COR-005',
      name: '修正逻辑 - 20MPa, 368°C（刚过饱和温度）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 368, pressure: 20 },
      expected: { 
        success: true, 
        minGJ: 2.29, 
        maxGJ: 2.49,
        expectedInterval: { t1: 365.8, t2: 370, description: '使用饱和温度365.8作为下界' }
      }
    }
  ],

  // ========== 二、温度命中网格点验证（4个用例） ==========
  grid_point: [
    {
      id: 'GRID-001',
      name: '网格点验证 - 10MPa, 320°C（命中网格点，不触发修正）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 320, pressure: 10 },
      expected: { 
        success: true, 
        minGJ: 2.60, 
        maxGJ: 2.80,
        expectedInterval: { t1: 310, t2: 320, description: '命中网格点320°C，直接使用网格点焓值' }
      }
    },
    {
      id: 'GRID-002',
      name: '网格点验证 - 10MPa, 350°C（命中网格点）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 350, pressure: 10 },
      expected: { 
        success: true, 
        minGJ: 2.74, 
        maxGJ: 2.94,
        expectedInterval: { t1: 340, t2: 350, description: '直接命中网格点350°C' }
      }
    },
    {
      id: 'GRID-003',
      name: '网格点验证 - 1MPa, 200°C（命中网格点）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 200, pressure: 1 },
      expected: { 
        success: true, 
        minGJ: 2.64, 
        maxGJ: 2.84,
        expectedInterval: { t1: 190, t2: 200, description: '直接命中网格点200°C' }
      }
    },
    {
      id: 'GRID-004',
      name: '网格点验证 - 0.1MPa, 150°C（命中网格点）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 150, pressure: 0.1 },
      expected: { 
        success: true, 
        minGJ: 2.59, 
        maxGJ: 2.79,
        expectedInterval: { t1: 140, t2: 150, description: '直接命中网格点150°C' }
      }
    }
  ],

  // ========== 三、温度区间查找验证（4个用例） ==========
  interval_lookup: [
    {
      id: 'INTERVAL-001',
      name: '区间查找 - 10MPa, 405°C（应使用400~410区间）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 405, pressure: 10 },
      expected: { 
        success: true, 
        minGJ: 2.93, 
        maxGJ: 3.13,
        expectedInterval: { t1: 400, t2: 410, description: '远离饱和温度400~410区间，不触发饱和修正' }
      }
    },
    {
      id: 'INTERVAL-002',
      name: '区间查找 - 10MPa, 415°C（应使用410~420区间）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 415, pressure: 10 },
      expected: { 
        success: true, 
        minGJ: 2.96, 
        maxGJ: 3.16,
        expectedInterval: { t1: 410, t2: 420, description: '远离饱和温度410~420区间，不触发饱和修正' }
      }
    },
    {
      id: 'INTERVAL-003',
      name: '区间查找 - 10MPa, 425°C（应使用420~430区间）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 425, pressure: 10 },
      expected: { 
        success: true, 
        minGJ: 2.99, 
        maxGJ: 3.19,
        expectedInterval: { t1: 420, t2: 430, description: '远离饱和温度420~430区间，不触发饱和修正' }
      }
    },
    {
      id: 'INTERVAL-004',
      name: '区间查找 - 10MPa, 435°C（应使用430~440区间）',
      input: { mediumType: 'superheated_steam', weight: 1, temp: 435, pressure: 10 },
      expected: { 
        success: true, 
        minGJ: 3.02, 
        maxGJ: 3.22,
        expectedInterval: { t1: 430, t2: 440, description: '远离饱和温度430~440区间，不触发饱和修正' }
      }
    }
  ],

  // ========== 四、重量线性验证（3个用例） ==========
  weight_linearity: [
    {
      id: 'WEIGHT-001',
      name: '重量线性 - 10MPa, 400°C, 2吨（应为1吨的2倍）',
      input: { mediumType: 'superheated_steam', weight: 2, temp: 400, pressure: 10 },
      expected: { 
        success: true, 
        exactGJ: 6.027,
        tolerance: 0.01,
        description: '2吨时GJ应为1吨的2倍（约3.0135 × 2 = 6.027）'
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
        description: '0.5吨时GJ应为1吨的0.5倍（约3.0135 × 0.5 = 1.50675）'
      }
    },
    {
      id: 'WEIGHT-003',
      name: '重量线性 - 10MPa, 400°C, 100吨（应为1吨的100倍）',
      input: { mediumType: 'superheated_steam', weight: 100, temp: 400, pressure: 10 },
      expected: { 
        success: true, 
        exactGJ: 301.35,
        tolerance: 0.01,
        description: '100吨时GJ应为1吨的100倍（约3.0135 × 100 = 301.35）'
      }
    }
  ]
};

module.exports = testCases;