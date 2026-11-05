/* =====================================================================
   data.js - financial statements stored in the project (no Excel upload needed)
   Generated from nvda.xlsx and goog.xlsx (NVDA, GOOGL) and SEC 10-K filings (AMZN, AAPL, MSFT, META, TSLA, AVGO, AMD, NFLX). Amounts are USD millions, EPS and
   dividends per share in USD, share counts in millions.
   Each entry is keyed by ticker (must match the symbol in config.js) and has
   exactly the structure that script.js builds from an Excel file.
   To add a company: add one more entry below, or regenerate this file.
   ===================================================================== */
const DATA_LIBRARY = {
  NVDA: {
    company: "NVIDIA Corporation",
    ticker: "NVDA",
    currency: {"sym":"$","code":"USD"},
    subtitle: "",
    years: ["FY2022","FY2023","FY2024","FY2025","FY2026"],
    income: {
      revenue: [26914,26974,60922,130497,215938],
      costOfRevenue: [9439,11618,16621,32639,62475],
      grossProfit: [17475,15356,44301,97858,153463],
      sga: [2166,2440,2654,3491,4579],
      rnd: [5268,7339,8675,12914,18497],
      operatingIncome: [10041,5577,32972,81453,130387],
      interestExpense: [-236,-262,-257,-247,-259],
      investmentGain: [100,-45,238,1030,8918],
      netIncome: [9752,4368,29760,72880,120067],
      dilutedShares: [25350,25070,24940,24804,24514],
      epsDiluted: [0.39,0.17,1.19,2.94,4.9],
      dps: [0.016,0.016,0.016,0.034,0.04],
    },
    balance: {
      cash: [1990,3389,7280,8589,10605],
      securities: [19218,9907,18704,34621,51951],
      receivables: [4650,3827,9999,23065,38466],
      inventory: [2605,5159,5282,10080,21403],
      currentAssets: [28829,23073,44345,80126,125605],
      ppe: [3607,4845,5260,8076,13250],
      ltInvestments: [266,299,1321,3387,22251],
      goodwill: [4349,4372,4430,5188,20832],
      totalAssets: [44187,41182,65728,111601,206803],
      payables: [1783,1193,2699,6310,9812],
      currentLiabilities: [4335,6563,10631,18047,32163],
      equity: [26612,22101,42978,79327,157293],
      totalDebt: [11831,12031,11056,9982,11412],
      netCash: [9377,1265,14928,33228,51144],
    },
    cashflow: {
      ocf: [9108,5641,28090,64089,102718],
      capex: [-976,-1833,-1069,-3236,-6042],
      buybacks: [-1904,-11514,-12316,-40636,-48034],
      dividends: [-399,-398,-395,-834,-974],
      fcf: [8132,3808,27021,60853,96676],
    },
    notes: [],
    missing: [],
  },
  GOOGL: {
    company: "Alphabet Inc.",
    ticker: "GOOGL",
    currency: {"sym":"$","code":"USD"},
    subtitle: "",
    years: ["FY2021","FY2022","FY2023","FY2024","FY2025"],
    income: {
      revenue: [257637,282836,307394,350018,402836],
      costOfRevenue: [110939,126203,133332,146306,162535],
      grossProfit: [146698,156633,174062,203712,240301],
      sga: [36422,42291,44342,41996,50175],
      rnd: [31562,39500,45427,49326,61087],
      operatingIncome: [78714,74842,84293,112390,129039],
      interestExpense: [-346,-357,-308,-268,-736],
      investmentGain: [12270,-5519,-823,2671,24620],
      netIncome: [76033,59972,73795,100118,132170],
      dilutedShares: [13553,13159,12722,12447,12230],
      epsDiluted: [5.61,4.56,5.8,8.04,10.81],
      dps: [0,0,0,0,0],
    },
    balance: {
      cash: [20945,21879,24048,23466,30708],
      securities: [118704,91883,86868,72191,96135],
      receivables: [39304,40258,47964,52340,62886],
      inventory: [1170,0,0,0,0],
      currentAssets: [188143,164795,171530,163711,206038],
      ppe: [110558,127049,148436,184624,261818],
      ltInvestments: [29549,30492,31929,37982,68687],
      goodwill: [22956,28960,29198,31885,33380],
      totalAssets: [359268,365264,402392,450256,595281],
      payables: [6037,5128,7493,7987,12200],
      currentLiabilities: [64254,69300,81814,89122,102745],
      equity: [251635,256144,283379,325084,415265],
      totalDebt: [28508,29977,29787,30437,66996],
      netCash: [111141,83785,81129,65220,59847],
    },
    cashflow: {
      ocf: [91652,91495,101746,125299,164713],
      capex: [-24640,-31485,-32251,-52535,-91447],
      buybacks: [-50274,-59296,-61504,-62222,-45709],
      dividends: [0,0,0,-7363,-10049],
      fcf: [67012,60010,69495,72764,73266],
    },
    notes: [],
    missing: ["Dividends per share ($)"],
  },
  AMZN: {
    company: "Amazon.com, Inc.",
    ticker: "AMZN",
    currency: {"sym":"$","code":"USD"},
    subtitle: "",
    years: [
      "FY2021",
      "FY2022",
      "FY2023",
      "FY2024",
      "FY2025"
    ],
    income: {
      revenue: [469822,513983,574785,637959,716924],
      costOfRevenue: [272344,288831,304739,326288,356414],
      grossProfit: [197478,225152,270046,311671,360510],
      sga: [41374,54129,56186,55266,58301],
      rnd: [56052,73213,85622,88544,108521],
      operatingIncome: [24879,12248,36852,68593,79975],
      interestExpense: [-1809,-2367,-3182,-2406,-2274],
      investmentGain: [14633,-16806,938,-2250,15229],
      netIncome: [33364,-2722,30425,59248,77670],
      dilutedShares: [10296,10189,10492,10721,10827],
      epsDiluted: [3.24,-0.27,2.9,5.53,7.17],
      dps: [0,0,0,0,0]
    },
    balance: {
      cash: [36220,53888,73387,78779,86810],
      securities: [59829,16138,13393,22423,36219],
      receivables: [32891,42360,52253,55451,67729],
      inventory: [32640,34405,33318,34214,38325],
      currentAssets: [161580,146791,172351,190867,229083],
      ppe: [160281,186715,204177,252665,357025],
      ltInvestments: [0,0,0,0,0],
      goodwill: [15371,20288,22789,23074,23273],
      totalAssets: [420549,462675,527854,624894,818042],
      payables: [78664,79600,84981,94363,121909],
      currentLiabilities: [142266,155393,164917,179431,218005],
      equity: [138245,146043,201875,285970,411065],
      totalDebt: [116395,140118,135611,130900,152987],
      netCash: [-20346,-70092,-48831,-29698,-29958]
    },
    cashflow: {
      ocf: [46327,46752,84946,115877,139514],
      capex: [-55396,-58321,-48133,-77658,-128320],
      buybacks: [0,-6000,0,0,0],
      dividends: [0,0,0,0,0],
      fcf: [-9069,-11569,36813,38219,11194]
    },
    notes: [
      "Amazon: 'Research & development' is Technology and infrastructure; 'SG&A' is Sales and marketing plus General and administrative. Fulfillment costs (USD 109.1B in FY2025) are not shown as a separate line.",
      "Amazon: capex is purchases of property and equipment net of proceeds from sales and incentives, so free cash flow matches the figure Amazon reports. Debt is long-term debt plus long-term lease liabilities (current portions are inside accrued expenses and not split out).",
      "Amazon: 'Gain (loss) on investments' is Other income (expense), net, which is mostly mark-to-market moves on its Rivian stake. Long-term investments (including non-marketable stakes) sit inside Other assets and are not split out."
    ],
    missing: [
      "Dividends per share ($)",
      "Long-term investments (incl. non-marketable equity)"
    ]
  },
  AAPL: {
    company: "Apple Inc.",
    ticker: "AAPL",
    currency: {"sym":"$","code":"USD"},
    subtitle: "",
    years: [
      "FY2021",
      "FY2022",
      "FY2023",
      "FY2024",
      "FY2025"
    ],
    income: {
      revenue: [365817,394328,383285,391035,416161],
      costOfRevenue: [212981,223546,214137,210352,220960],
      grossProfit: [152836,170782,169148,180683,195201],
      sga: [21973,25094,24932,26097,27601],
      rnd: [21914,26251,29915,31370,34550],
      operatingIncome: [108949,119437,114301,123216,133050],
      interestExpense: [-2645,-2931,-3933,0,0],
      investmentGain: [0,0,0,0,0],
      netIncome: [94680,99803,96995,93736,112010],
      dilutedShares: [16865,16326,15813,15408,15005],
      epsDiluted: [5.61,6.11,6.13,6.08,7.46],
      dps: [0.85,0.9,0.94,0.98,1.02]
    },
    balance: {
      cash: [34940,23646,29965,29943,35934],
      securities: [155576,145463,132134,126707,96486],
      receivables: [26278,28184,29508,33410,39777],
      inventory: [6580,4946,6331,7286,5718],
      currentAssets: [134836,135405,143566,152987,147957],
      ppe: [39440,42117,43715,45680,49834],
      ltInvestments: [0,0,0,0,0],
      goodwill: [0,0,0,0,0],
      totalAssets: [351002,352755,352583,364980,359241],
      payables: [54763,64115,62611,68960,69860],
      currentLiabilities: [125481,153982,145308,176392,165631],
      equity: [63090,50672,62146,56950,73733],
      totalDebt: [124719,120069,111088,106629,98657],
      netCash: [65797,49040,51011,50021,33763]
    },
    cashflow: {
      ocf: [104038,122151,110543,118254,111482],
      capex: [-11085,-10708,-10959,-9447,-12715],
      buybacks: [-85971,-89402,-77550,-94949,-90711],
      dividends: [-14467,-14841,-15025,-15234,-15421],
      fcf: [92953,111443,99584,108807,98767]
    },
    notes: [
      "Apple: fiscal years end in late September. FY2026 is not included because Apple reports it on 2 November 2026.",
      "Apple: 'Cash + securities' includes both current and non-current marketable securities, so net cash is Apple's own definition (about USD 34B at FY2025). Debt is commercial paper plus term debt; lease liabilities are not split out.",
      "Apple: interest expense was not found as a separate line for FY2024 and FY2025, so it is shown as 0 and interest coverage reads n/a for those years. Apple has no goodwill line on its balance sheet."
    ],
    missing: [
      "Gain (loss) on sale of investments",
      "Goodwill",
      "Long-term investments (incl. non-marketable equity)"
    ]
  },
  MSFT: {
    company: "Microsoft Corporation",
    ticker: "MSFT",
    currency: {"sym":"$","code":"USD"},
    subtitle: "",
    years: ["FY2022","FY2023","FY2024","FY2025","FY2026"],
    income: {
      revenue: [198270,211915,245122,281724,331839],
      costOfRevenue: [62650,65863,74114,87831,106374],
      grossProfit: [135620,146052,171008,193893,225465],
      sga: [27725,30334,32065,32877,34666],
      rnd: [24512,27195,29510,32488,35562],
      operatingIncome: [83383,88523,109433,128528,155237],
      interestExpense: [0,0,0,0,0],
      investmentGain: [333,788,-1646,-4901,10697],
      netIncome: [72738,72361,88136,101832,133749],
      dilutedShares: [7540,7472,7469,7465,7453],
      epsDiluted: [9.65,9.68,11.8,13.64,17.95],
      dps: [2.48,2.72,3,3.32,3.64]
    },
    balance: {
      cash: [13931,34704,18315,30242,20935],
      securities: [90826,76558,57228,64323,55908],
      receivables: [44261,48688,56924,69905,80876],
      inventory: [3742,2500,1246,938,1397],
      currentAssets: [169684,184257,159734,191131,207710],
      ppe: [74398,95641,135591,204966,313076],
      ltInvestments: [6891,9879,14600,15405,36348],
      goodwill: [67524,67886,119220,119509,119651],
      totalAssets: [364840,411976,512163,619003,758376],
      payables: [19000,18095,21996,27724,42416],
      currentLiabilities: [95082,104149,125286,141218,168825],
      equity: [166542,206223,268477,343479,442387],
      totalDebt: [61270,59965,67134,60588,56826],
      netCash: [43487,51297,8409,33977,20017]
    },
    cashflow: {
      ocf: [89035,87582,118548,136162,182935],
      capex: [-23886,-28107,-44477,-64551,-115948],
      buybacks: [-32696,-22245,-17254,-18420,-22271],
      dividends: [-18135,-19800,-21771,-24082,-26445],
      fcf: [65149,59475,74071,71611,66987]
    },
    notes: [
      "Microsoft: fiscal years end on 30 June, so FY2026 ended 30 June 2026. 'SG&A' is Sales and marketing plus General and administrative. Dividends per share are the amounts declared in each fiscal year.",
      "Microsoft: 'Gain (loss) on investments' is Other income (expense), net, which mainly reflects investment gains and losses and equity-method results. Interest expense is not a separate income-statement line, so it is shown as 0 and interest coverage reads n/a.",
      "Microsoft: debt is short-term debt, current and long-term debt, plus long-term operating lease liabilities. Capex is additions to property and equipment and does not include finance-lease assets."
    ],
    missing: []
  },
  META: {
    company: "Meta Platforms, Inc.",
    ticker: "META",
    currency: {"sym":"$","code":"USD"},
    subtitle: "",
    years: ["FY2021","FY2022","FY2023","FY2024","FY2025"],
    income: {
      revenue: [117929,116609,134902,164501,200966],
      costOfRevenue: [22649,25249,25959,30161,36175],
      grossProfit: [95280,91360,108943,134340,164791],
      sga: [23872,27078,23709,21087,24143],
      rnd: [24655,35338,38483,43873,57372],
      operatingIncome: [46753,28944,46751,69380,83276],
      interestExpense: [0,0,0,0,0],
      investmentGain: [0,0,0,0,0],
      netIncome: [39370,23200,39098,62360,60458],
      dilutedShares: [2859,2702,2629,2614,2574],
      epsDiluted: [13.77,8.59,14.87,23.86,23.49],
      dps: [0,0,0,2,2.1]
    },
    balance: {
      cash: [16601,14681,41862,43889,35873],
      securities: [31397,26057,23541,33926,45719],
      receivables: [14039,13466,16169,16994,19769],
      inventory: [0,0,0,0,0],
      currentAssets: [66666,59549,85365,100045,108722],
      ppe: [57809,79518,96587,121346,176400],
      ltInvestments: [6775,6201,6141,6070,27524],
      goodwill: [19197,20306,20654,20654,24534],
      totalAssets: [165987,185727,229623,276054,366021],
      payables: [4083,4990,4849,7687,8894],
      currentLiabilities: [21135,27026,31960,33596,41836],
      equity: [124879,125713,153168,182637,217243],
      totalDebt: [13873,26591,37234,49060,83897],
      netCash: [34125,14147,28169,28755,-2305]
    },
    cashflow: {
      ocf: [57683,50475,71113,91328,115800],
      capex: [-18690,-31186,-27045,-37256,-69691],
      buybacks: [-44537,-27956,-19774,-30125,-26248],
      dividends: [0,0,0,-5072,-5324],
      fcf: [38993,19289,44068,54072,46109]
    },
    notes: [
      "Meta: 'SG&A' is Marketing and sales plus General and administrative. Meta started paying a dividend in 2024. Interest expense is not a separate income-statement line, so it is shown as 0 and interest coverage reads n/a.",
      "Meta: debt is long-term debt plus operating lease liabilities (current and non-current). Capex is purchases of property and equipment and does not include principal payments on finance leases. Meta has no inventory, and 'Long-term investments' are its non-marketable equity investments."
    ],
    missing: [
      "Gain (loss) on sale of investments"
    ]
  },
  TSLA: {
    company: "Tesla, Inc.",
    ticker: "TSLA",
    currency: {"sym":"$","code":"USD"},
    subtitle: "",
    years: ["FY2021","FY2022","FY2023","FY2024","FY2025"],
    income: {
      revenue: [53823,81462,96773,97690,94827],
      costOfRevenue: [40217,60609,79113,80240,77733],
      grossProfit: [13606,20853,17660,17450,17094],
      sga: [4517,3946,4800,5150,5834],
      rnd: [2593,3075,3969,4540,6411],
      operatingIncome: [6523,13656,8891,7076,4355],
      interestExpense: [-371,-191,-156,-350,-338],
      investmentGain: [27,-140,0,589,-68],
      netIncome: [5519,12556,14997,7091,3794],
      dilutedShares: [3386,3475,3485,3498,3528],
      epsDiluted: [1.63,3.62,4.3,2.04,1.08],
      dps: [0,0,0,0,0]
    },
    balance: {
      cash: [17576,16253,16398,16139,16513],
      securities: [131,5932,12696,20424,27546],
      receivables: [1986,3116,3319,4255,4598],
      inventory: [5757,12839,13626,12017,12392],
      currentAssets: [27100,40917,49616,58360,68642],
      ppe: [31176,36635,45123,51501,56186],
      ltInvestments: [0,0,0,0,0],
      goodwill: [200,194,253,0,0],
      totalAssets: [62131,82338,106618,122070,137806],
      payables: [10025,15255,14431,12474,13371],
      currentLiabilities: [19705,26709,28748,28821,31714],
      equity: [30189,44704,62634,72913,82137],
      totalDebt: [8873,5748,9573,13623,14719],
      netCash: [8834,16437,19521,22940,29340]
    },
    cashflow: {
      ocf: [11497,14724,13256,14923,14747],
      capex: [-6482,-7158,-8899,-11342,-8527],
      buybacks: [0,0,0,0,0],
      dividends: [0,0,0,0,0],
      fcf: [5015,7566,4357,3581,6220]
    },
    notes: [
      "Tesla: net income and EPS are attributable to common stockholders. 'SG&A' is Selling, general and administrative; restructuring and other charges are inside operating income but not shown as a separate line. Capex is purchases of property and equipment excluding finance leases, net of sales.",
      "Tesla: 'Gain (loss) on investments' is the gain or loss on digital assets (bitcoin). Debt is debt plus finance and operating leases (current and non-current). Property and receivables figures follow the S&P/StockAnalysis presentation, so they include lease-related assets and are broader than the narrow line on Tesla's balance sheet."
    ],
    missing: [
      "Dividends per share ($)",
      "Repurchase of common stock",
      "Long-term investments (incl. non-marketable equity)"
    ]
  },
  AVGO: {
    company: "Broadcom Inc.",
    ticker: "AVGO",
    currency: {"sym":"$","code":"USD"},
    subtitle: "",
    years: ["FY2021","FY2022","FY2023","FY2024","FY2025"],
    income: {
      revenue: [27450,33203,35819,51574,63887],
      costOfRevenue: [10606,11108,11129,19065,20593],
      grossProfit: [16844,22095,24690,32509,43294],
      sga: [1347,1382,1592,4959,4211],
      rnd: [4854,4919,5253,9310,10977],
      operatingIncome: [8519,14225,16207,13463,25484],
      interestExpense: [-1885,-1737,-1622,-3953,-3210],
      investmentGain: [0,0,0,0,0],
      netIncome: [6736,11495,14082,5895,23126],
      dilutedShares: [4290,4232,4272,4778,4853],
      epsDiluted: [1.5,2.65,3.3,1.23,4.77],
      dps: [1.44,1.64,1.84,2.1,2.36]
    },
    balance: {
      cash: [12163,12416,14189,9348,16178],
      securities: [0,0,0,0,0],
      receivables: [2071,2958,3653,6332,12150],
      inventory: [1297,1925,1898,1760,2270],
      currentAssets: [16586,18504,20847,19595,31573],
      ppe: [2936,2740,2617,3846,3848],
      ltInvestments: [0,0,0,0,0],
      goodwill: [43450,43614,43653,97873,97801],
      totalAssets: [75570,73249,72861,165645,171092],
      payables: [1086,998,1210,1662,1560],
      currentLiabilities: [6281,7052,7405,16697,18514],
      equity: [24989,22709,23988,67678,81292],
      totalDebt: [40273,39978,39648,68916,66461],
      netCash: [-28110,-27562,-25459,-59568,-50283]
    },
    cashflow: {
      ocf: [13764,16736,18085,19962,27537],
      capex: [-443,-424,-452,-548,-623],
      buybacks: [-1299,-8455,-7685,-12392,-6310],
      dividends: [-5913,-6733,-7645,-9814,-11142],
      fcf: [13321,16312,17633,19414,26914]
    },
    notes: [
      "Broadcom: fiscal years end in late October or early November (FY2025 ended 2 November 2025). Per-share data and share counts are adjusted for the 10-for-1 stock split of July 2024. Dividends per share are the quarterly rate times four, cross-checked against cash dividends paid.",
      "Broadcom: cost of revenue is the GAAP total and includes amortization of acquired intangibles (about USD 6.0B in FY2025), so gross margin is the GAAP figure. Operating income is GAAP, after amortization in operating expenses and restructuring charges. Debt includes lease liabilities; the FY2024 balance sheet reflects the VMware acquisition."
    ],
    missing: [
      "Marketable securities (short-term investments)",
      "Long-term investments (incl. non-marketable equity)",
      "Gain (loss) on sale of investments"
    ]
  },
  AMD: {
    company: "Advanced Micro Devices, Inc.",
    ticker: "AMD",
    currency: {"sym":"$","code":"USD"},
    subtitle: "",
    years: ["FY2021","FY2022","FY2023","FY2024","FY2025"],
    income: {
      revenue: [16434,23601,22680,25785,34639],
      costOfRevenue: [8505,12998,12220,13060,17487],
      grossProfit: [7929,10603,10460,12725,17152],
      sga: [1448,2336,2318,2735,4144],
      rnd: [2845,5005,5872,6456,8091],
      operatingIncome: [3648,1264,401,1900,3694],
      interestExpense: [-34,-88,-106,-92,-131],
      investmentGain: [0,0,1,0,341],
      netIncome: [3162,1320,854,1641,4335],
      dilutedShares: [1229,1571,1625,1637,1636],
      epsDiluted: [2.57,0.84,0.53,1,2.65],
      dps: [0,0,0,0,0]
    },
    balance: {
      cash: [2535,4835,3933,3787,5539],
      securities: [1073,1020,1840,1345,5013],
      receivables: [2706,4126,4323,6192,6315],
      inventory: [1955,3771,4351,5734,7920],
      currentAssets: [8583,15019,16768,19049,26947],
      ppe: [702,1513,1589,1802,2312],
      ltInvestments: [0,0,0,0,0],
      goodwill: [289,24177,24262,24839,25126],
      totalAssets: [12419,67580,67885,69226,76926],
      payables: [1321,2493,2055,2466,2929],
      currentLiabilities: [4240,6369,6689,7281,9455],
      equity: [7497,54750,55892,57568,62999],
      totalDebt: [661,2863,3003,2212,3847],
      netCash: [2947,2992,2770,2920,6705]
    },
    cashflow: {
      ocf: [3521,3565,1667,3041,7709],
      capex: [-301,-450,-546,-636,-974],
      buybacks: [-1762,-3702,-985,-862,-1316],
      dividends: [0,0,0,0,0],
      fcf: [3220,3115,1121,2405,6735]
    },
    notes: [
      "AMD: 'SG&A' is Marketing, general and administrative. Amortization of acquisition-related intangibles (USD 1.2B in FY2025) is a separate operating expense line in AMD's statement and is not part of 'SG&A' here, so R&D plus SG&A will not add up to total operating expenses. The Xilinx acquisition closed in February 2022, which is why total assets jump in FY2022.",
      "AMD: FY2025 net income of USD 4.3B includes a small gain from discontinued operations (the ZT Systems manufacturing business sold in 2025), and operating cash flow of USD 7.7B includes USD 1.2B from that business. Debt is long-term debt plus long-term operating leases. 'Gain on investments' is gains on long-term investments."
    ],
    missing: [
      "Dividends per share ($)",
      "Long-term investments (incl. non-marketable equity)"
    ]
  },
  NFLX: {
    company: "Netflix, Inc.",
    ticker: "NFLX",
    currency: {"sym":"$","code":"USD"},
    subtitle: "",
    years: ["FY2021","FY2022","FY2023","FY2024","FY2025"],
    income: {
      revenue: [29698,31616,33723,39001,45183],
      costOfRevenue: [17333,19168,19715,21038,23275],
      grossProfit: [12365,12447,14008,17963,21908],
      sga: [3897,4103,4378,4620,5190],
      rnd: [2274,2711,2676,2925,3391],
      operatingIncome: [6195,5633,6954,10418,13327],
      interestExpense: [-766,-706,-700,-719,-777],
      investmentGain: [0,0,0,0,0],
      netIncome: [5116,4492,5408,8712,10981],
      dilutedShares: [4554,4513,4495,4393,4344],
      epsDiluted: [1.124,0.995,1.203,1.98,2.53],
      dps: [0,0,0,0,0]
    },
    balance: {
      cash: [6028,5147,7117,7805,9034],
      securities: [0,911,21,1779,29],
      receivables: [0,0,0,0,0],
      inventory: [0,0,0,0,0],
      currentAssets: [8070,9266,9918,13100,13020],
      ppe: [1323,1398,1491,1594,2004],
      ltInvestments: [0,0,0,0,0],
      goodwill: [0,0,0,0,0],
      totalAssets: [44585,48595,48732,53630,55597],
      payables: [837,672,747,900,901],
      currentLiabilities: [8489,7931,8861,10755,10981],
      equity: [15849,20777,20588,24744,26615],
      totalDebt: [15393,14353,14543,15583,14463],
      netCash: [-9365,-8295,-7405,-5999,-5400]
    },
    cashflow: {
      ocf: [393,2026,7274,7361,10149],
      capex: [-525,-408,-349,-440,-688],
      buybacks: [-600,0,-6045,-6264,-9127],
      dividends: [0,0,0,0,0],
      fcf: [-132,1619,6926,6922,9461]
    },
    notes: [
      "Netflix: EPS and share counts for all years are adjusted for the 10-for-1 stock split of November 2025. 'SG&A' is Marketing plus General and administrative; 'R&D' is Technology and development. Cost of revenues is mostly content amortization.",
      "Netflix: cash flow from operations is after spending on content (USD 17.1B in FY2025), so free cash flow is far below net income plus amortization. Content assets (USD 32.8B) appear under 'Other' in the asset split. Interest and other income, which includes currency effects on debt, is not shown as an investment gain. Debt excludes lease liabilities."
    ],
    missing: [
      "Dividends per share ($)",
      "Accounts receivable",
      "Inventories",
      "Goodwill",
      "Long-term investments (incl. non-marketable equity)",
      "Gain (loss) on sale of investments"
    ]
  },
};