/**
 * ============================================================================
 * V2.3.1 性能基准（Performance Benchmark）—— 任务书 §19 / §40
 * ============================================================================
 *
 * 用途：在【修改代码之前】对当前版本建立基准，修改后同机复测并对比，
 *       杜绝"凭感觉说性能提高"。
 *
 * 测试项（任务书 §19）：
 *   Test A  单方案 8760 仿真            → simulationTime
 *   Test B  DataSummary                 → summaryTime
 *   Test C  Finance                     → financeTime
 *   Test D  Chart 数据生成              → chartDataTime
 *   Test E  8760 表格首次渲染           → tableRenderTime
 *   Test F  NSGA-II 20×10               → optimizationTime / evaluationCount / cacheHitRate
 *
 * 用法：
 *   node tests/performance-benchmark.js                 运行并输出报告 + benchmark-last.json
 *   node tests/performance-benchmark.js --save          把本次结果存为基线 benchmark-baseline.json
 *   node tests/performance-benchmark.js --compare       与基线对比，输出前后对照表
 *
 * 注意：单次测量受 JIT/负载抖动影响，每个用例先预热再取多轮中位数。
 */

'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = process.env.WB_BENCH_ROOT || path.resolve(__dirname, '..');
const OUT_FILE = path.join(__dirname, 'benchmark-report.txt');
const LAST_FILE = path.join(__dirname, 'benchmark-last.json');
const BASE_FILE = path.join(__dirname, 'benchmark-baseline.json');
const WS = 'C:/Users/lishuo/.workbuddy/binaries/node/workspace/node_modules';

let JSDOM = null;
try { JSDOM = require(path.join(WS, 'jsdom')).JSDOM; } catch (e) { JSDOM = null; }

const lines = [];
function out(s) { lines.push(s); }
function section(t) { out(''); out('='.repeat(78)); out(t); out('='.repeat(78)); }

/** 中位数计时：先预热 warmup 次，再跑 runs 次取中位数 */
function bench(fn, runs = 7, warmup = 2) {
    for (let i = 0; i < warmup; i++) fn();
    const t = [];
    for (let i = 0; i < runs; i++) {
        const t0 = process.hrtime.bigint();
        fn();
        t.push(Number(process.hrtime.bigint() - t0) / 1e6);
    }
    t.sort((a, b) => a - b);
    return t[Math.floor(t.length / 2)];
}
function ms(x) { return Number(x).toFixed(2) + ' ms'; }

// ---------------------------------------------------------------------------
// 0. 加载模块（与 index.html 的 script 顺序一致）
// ---------------------------------------------------------------------------
global.self = global;
global.window = global;

const XLSX = require(path.join(ROOT, 'js', 'vendor', 'xlsx.full.min.js'));
global.XLSX = XLSX;

const MODULES = [
    'js/utils.js',
    'js/parameter-manager.js',
    'js/result-data-store.js',
    'js/simulation-engine.js',
    'js/data-summary.js',
    'js/estimate.js',
    'js/finance-engine.js',
    'js/optimization-engine.js',
];
const bundle = MODULES
    .map(f => '/* ===== ' + f + ' ===== */\n' + fs.readFileSync(path.join(ROOT, f), 'utf8'))
    .join('\n;\n');
vm.runInThisContext(bundle, { filename: 'v231-benchmark-bundle.js' });

const DS = global.ResultDataStore;
const OptimizationEngine = global.OptimizationEngine;
const DataSummary = global.DataSummary;
const Estimate = global.Estimate;
const FinanceEngine = global.FinanceEngine;
const runSingleSimulation = global.runSingleSimulation;

// ---------------------------------------------------------------------------
// 输入数据与参数
// ---------------------------------------------------------------------------
const wb = XLSX.read(fs.readFileSync(path.join(ROOT, 'input.xlsx')), { type: 'buffer' });
const sheet = wb.Sheets[wb.SheetNames[0]];
const rawRows = XLSX.utils.sheet_to_json(sheet, { header: ['pv', 'wind'], range: 1 });
const N = rawRows.length;
const pvData = new Float64Array(N);
const windData = new Float64Array(N);
for (let i = 0; i < N; i++) {
    pvData[i] = Number(rawRows[i].pv) || 0;
    windData[i] = Number(rawRows[i].wind) || 0;
}

const SIM_CONFIG = {
    electrolyzerMinRatio: 0.3, maxExportHourly: 0.0, maxExportTotal: 0.0,
    maxImportRatio: 0.15, chargeEfficiency: 0.92, dischargeEfficiency: 0.92,
    hydrogenConsumption: 55,
};
const PRICES = {
    windUnitPrice: 3800, pvUnitPrice: 2500, storageUnitPrice: 800, electrolyzerUnitPrice: 1600,
    transmissionCost: 10000, landCost: 10000, otherFacilitiesRatio: 70,
};
const FINANCE = {
    calc_period: 22, construct_period: 2, batch_count: 1,
    capital_benchmark: 8.0, industry_benchmark_pre: 5.0, industry_benchmark_post: 4.5,
    capital_ratio: 20, loan_period: 15, long_term_loan_rate: 3.0,
    single_electrolyzer_mw: 5.0, capacity_elec_fee: 20.0, ongrid_price: 0.2, offgrid_price: 0.7,
    wheeling_fee: 0.12, h2_labor_cost: 3.0, electrolyzer_maint: 40.0, electrolyzer_overhaul: 240.0,
    water_unit_price: 2.0, wind_pv_om: 35.0, depreciation_years: 20, residual_rate: 5.0,
    h2_price: 30.0, output_vat_rate: 13.0, income_tax_rate: 25.0, surtax_rate: 10.0,
};

const SCHEME = { windCapacity: 200, pvCapacity: 360, storagePower: 100, storageDuration: 2, electrolyzerCapacity: 160 };
const SCHEME_F = { windCapacity: 150, pvCapacity: 250, storagePower: 80, storageDuration: 3, electrolyzerCapacity: 140 };

const results = {};

out('多能互补风光储氢分析软件 V2.3.1 —— 性能基准');
out('生成时间：' + new Date().toISOString());
out('运行环境：Node ' + process.version + '（中位数计时：预热 ' + 2 + ' 轮 + 测量 7 轮取中位）');
out('数据：input.xlsx ' + N + ' 小时；基准方案 ' + JSON.stringify(SCHEME));

// ---------------------------------------------------------------------------
section('Test A  单方案 8760 仿真（simulationTime）');
{
    const sim = runSingleSimulation(pvData, windData, SCHEME, SIM_CONFIG);
    results.simulationTime = bench(() => runSingleSimulation(pvData, windData, SCHEME, SIM_CONFIG));
    out('  simulationTime = ' + ms(results.simulationTime));
    out('  结果规模: Float64Array(' + sim.results.length + ') = ' + (sim.results.length * 8 / 1024).toFixed(1) + ' KB');
    global.__benchSim = sim;
}

// ---------------------------------------------------------------------------
section('Test B  DataSummary（summaryTime）');
{
    const sim = global.__benchSim;
    results.summaryTime = bench(() => DataSummary.generateSummary([sim], SCHEME));
    out('  summaryTime = ' + ms(results.summaryTime));
    global.__benchSummary = DataSummary.generateSummary([sim], SCHEME);
}

// ---------------------------------------------------------------------------
section('Test C  Finance（financeTime，含概算）');
{
    const summaryRow = global.__benchSummary[0];
    results.financeTime = bench(() => {
        const estimateRow = Estimate.batchEstimate([summaryRow], PRICES)[0];
        return FinanceEngine.calculateAll(FINANCE, summaryRow, estimateRow);
    });
    out('  financeTime（概算 + 财务评价 + LCOH）= ' + ms(results.financeTime));
}

// ---------------------------------------------------------------------------
section('Test D  Chart 数据生成（chartDataTime）');
{
    const rec = DS.create(global.__benchSim.results, SCHEME, { sums: global.__benchSim.sums });
    DS.ChartDataCache.clear();

    // 冷路径：首次构建（含降采样索引计算）
    results.chartDataCold = bench(() => {
        DS.ChartDataCache.clear();
        DS.ChartDataAdapter.getHourlyView(rec, 'total', 1500);
        DS.ChartDataAdapter.getOverviewView(rec, 1500);
        DS.ChartDataAdapter.getMonthlyView(rec);
    }, 5, 1);
    out('  chartDataTime（冷，3 类视图首次构建）= ' + ms(results.chartDataCold));

    // 热路径：命中缓存
    DS.ChartDataAdapter.getHourlyView(rec, 'total', 1500);
    DS.ChartDataAdapter.getOverviewView(rec, 1500);
    DS.ChartDataAdapter.getMonthlyView(rec);
    results.chartDataWarm = bench(() => {
        DS.ChartDataAdapter.getHourlyView(rec, 'total', 1500);
        DS.ChartDataAdapter.getOverviewView(rec, 1500);
        DS.ChartDataAdapter.getMonthlyView(rec);
    });
    out('  chartDataTime（热，全部命中缓存）= ' + ms(results.chartDataWarm));

    // 全分辨率视图取一列（表格/导出以外的显示路径）
    results.chartDataColumn = bench(() => DS.ChartDataAdapter.getFullView(rec).map(d => d['合计电量']));
    out('  chartDataTime（全分辨率视图取一列 8760 点）= ' + ms(results.chartDataColumn));

    // 临时数组审计：统计一次概览视图构建期间的 Float64Array 分配量
    let allocs = 0, allocBytes = 0;
    const OrigF64 = Float64Array;
    global.Float64Array = function (...args) {
        allocs++;
        allocBytes += (typeof args[0] === 'number' ? args[0] : (args[0] ? args[0].length : 0)) * 8;
        return new OrigF64(...args);
    };
    global.Float64Array.prototype = OrigF64.prototype;
    Object.setPrototypeOf(global.Float64Array, OrigF64);
    try {
        DS.ChartDataCache.clear();
        DS.ChartDataAdapter.getOverviewView(rec, 1500);
    } finally {
        global.Float64Array = OrigF64;
    }
    results.overviewAllocArrays = allocs;
    results.overviewAllocBytes = allocBytes;
    out('  概览视图单次构建的 Float64Array 分配: ' + allocs + ' 个 / ' + (allocBytes / 1024).toFixed(1) + ' KB');

    const hv = DS.ChartDataAdapter.getHourlyView(rec, 'total', 1500);
    const ov = DS.ChartDataAdapter.getOverviewView(rec, 1500);
    out('  降采样: 单列 ' + hv.originalPoints + ' → ' + hv.length + ' 点；7 列并集 → ' + ov.length + ' 点');
    DS.ChartDataCache.clear();
}

// ---------------------------------------------------------------------------
section('Test E  8760 表格首次渲染（tableRenderTime）');
if (!JSDOM) {
    out('  [SKIP] 未安装 jsdom，跳过 DOM 测试');
} else {
    try {
        const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
        const dom = new JSDOM(html, { url: 'http://127.0.0.1/', runScripts: 'outside-only', pretendToBeVisual: true });
        const w = dom.window;
        w.self = w;
        w.XLSX = XLSX;
        const SCRIPTS = [
            'js/utils.js', 'js/parameter-manager.js', 'js/result-data-store.js', 'js/excel-io.js',
            'js/simulation-engine.js', 'js/data-summary.js', 'js/estimate.js', 'js/finance-engine.js',
            'js/chart-module.js', 'js/optimization-engine.js', 'js/app.js',
        ];
        const pageErrors = [];
        w.addEventListener('error', e => pageErrors.push((e && e.error && e.error.message) || e.message));
        for (const f of SCRIPTS) w.eval(fs.readFileSync(path.join(ROOT, f), 'utf8'));
        if (typeof w.__wbTestHooks === 'undefined') {
            w.document.dispatchEvent(new w.Event('DOMContentLoaded'));
        }

        w.self.__sim = global.__benchSim;
        w.eval('window.__wbTestHooks.__injectSimulationResults([ self.__sim ]);');

        const sel = w.document.getElementById('schemeSelect');
        const opt = w.document.createElement('option');
        opt.value = '0';
        sel.appendChild(opt);

        // 稳定视口，保证前后可比
        const container = w.document.getElementById('simTableContainer');
        Object.defineProperty(container, 'clientHeight', { configurable: true, value: 480 });

        // 首次渲染（含表头/表尾构建）
        const t0 = process.hrtime.bigint();
        sel.selectedIndex = 0;
        sel.dispatchEvent(new w.Event('change'));
        results.tableRenderTime = Number(process.hrtime.bigint() - t0) / 1e6;

        // 重复渲染取中位数（交互路径的主导成本；首帧含一次性构建，波动大）
        sel.selectedIndex = -1;
        results.tableRenderMedian = bench(() => {
            sel.selectedIndex = 0;
            sel.dispatchEvent(new w.Event('change'));
            sel.selectedIndex = -1;
        }, 5, 1);

        const tbody = w.document.getElementById('simulationTable').tBodies[0];
        const rows = tbody.querySelectorAll('tr').length;
        const nodes = w.document.getElementById('simulationTable').querySelectorAll('*').length;
        out('  tableRenderTime（首次，含一次性构建）= ' + ms(results.tableRenderTime));
        out('  tableRenderTime（重复渲染中位数）    = ' + ms(results.tableRenderMedian));
        out('  DOM 行数: ' + rows + ' 行；DOM 节点: ' + nodes + ' 个（任务书 §25 目标 < 1000）');
        results.tableRows = rows;
        results.tableNodes = nodes;

        // 滚动重绘
        Object.defineProperty(container, 'scrollTop', { configurable: true, value: 5000 * 28 });
        results.tableScrollTime = bench(() => container.dispatchEvent(new w.Event('scroll')), 5, 1);
        out('  滚动重绘（跳到第 5000 小时）= ' + ms(results.tableScrollTime));
        void pageErrors;
    } catch (e) {
        out('  [FAIL] DOM 测试异常: ' + e.message);
    }
}

// ---------------------------------------------------------------------------
section('Test F  NSGA-II 20×10（optimizationTime）');
{
    const cfg = OptimizationEngine.normalizeConfig({
        variables: {
            windCapacity: { min: 100, max: 300, step: 50 },
            pvCapacity: { min: 100, max: 500, step: 50 },
            storagePower: { min: 0, max: 200, step: 50 },
            storageDuration: { min: 0, max: 4, step: 2 },
            electrolyzerCapacity: { min: 50, max: 200, step: 50 },
        },
        nsga2: {
            populationSize: 20, generations: 10,
            crossoverProbability: 0.9, mutationProbability: 0.1,
            randomSeed: 20260912, earlyStopping: false, patience: 15,
        },
        recommendationWeights: { eirr: 0.4, lcoh: 0.35, curtailmentRate: 0.25 },
        lcoh: { discountRate: 5.0 },
        baseline: SCHEME_F,
    });
    let lastStats = null;
    results.optimizationTime = bench(() => {
        const ctx = {
            pvData, windData, simulationConfig: SIM_CONFIG,
            prices: PRICES, financeParams: FINANCE, lcohDiscountRate: 5.0,
            cache: new Map(), stats: { evaluated: 0, cacheHits: 0 },
            config: cfg,
        };
        const s = OptimizationEngine.createSession({ config: cfg, context: ctx });
        s.ensureInitialized();
        while (!s.isFinished()) s.runNextGeneration();
        const r = s.getResult();
        // V2.3.1：createSession 现在会保留调用方的 stats 对象（修复前会重建导致恒为 0）；
        // 这里以结果自带的 statistics 为准，ctx.stats 作为交叉校验
        const st = r.statistics || {};
        const evaluated = st.totalEvaluated || ctx.stats.evaluated;
        const cacheHits = st.cacheHits || ctx.stats.cacheHits;
        lastStats = {
            stats: st,
            evaluated: evaluated,
            cacheHits: cacheHits,
            hitRate: (evaluated + cacheHits) > 0 ? cacheHits / (evaluated + cacheHits) : 0,
        };
        return r;
    }, 3, 1);
    out('  optimizationTime = ' + (results.optimizationTime / 1000).toFixed(3) + ' s');
    out('  evaluationCount  = ' + lastStats.evaluated + '（缓存命中 ' + lastStats.cacheHits + '，命中率 ' +
        (lastStats.hitRate * 100).toFixed(1) + '%）');
    out('  完成代数 = ' + (lastStats.stats.generationsCompleted || '-') +
        '，Pareto = ' + (lastStats.stats.paretoCount || '-'));
    results.evaluationCount = lastStats.evaluated;
    results.cacheHitRate = lastStats.hitRate;
}

// ---------------------------------------------------------------------------
section('基准汇总');
out('  simulationTime      = ' + ms(results.simulationTime));
out('  summaryTime         = ' + ms(results.summaryTime));
out('  financeTime         = ' + ms(results.financeTime));
out('  chartDataTime(冷)   = ' + ms(results.chartDataCold));
out('  chartDataTime(热)   = ' + ms(results.chartDataWarm));
out('  chartDataTime(列)   = ' + ms(results.chartDataColumn));
out('  概览视图分配        = ' + results.overviewAllocArrays + ' 个 Float64Array / ' +
    (results.overviewAllocBytes / 1024).toFixed(1) + ' KB');
out('  tableRenderTime(首帧)    = ' + ms(results.tableRenderTime));
out('  tableRenderTime(中位)    = ' + ms(results.tableRenderMedian || 0) +
    (results.tableRows ? '（' + results.tableRows + ' 行 / ' + results.tableNodes + ' 节点）' : ''));
out('  tableScrollTime     = ' + ms(results.tableScrollTime || 0));
out('  optimizationTime    = ' + (results.optimizationTime / 1000).toFixed(3) + ' s（20×10）');
out('  evaluationCount     = ' + results.evaluationCount);
out('  cacheHitRate        = ' + (results.cacheHitRate * 100).toFixed(1) + '%');

// ---------------------------------------------------------------------------
// 写文件与对比
// ---------------------------------------------------------------------------
fs.writeFileSync(OUT_FILE, lines.join('\n'), 'utf8');
fs.writeFileSync(LAST_FILE, JSON.stringify({
    date: new Date().toISOString(),
    node: process.version,
    results: results,
}, null, 2), 'utf8');

const args = process.argv.slice(2);
if (args.indexOf('--save') >= 0) {
    fs.writeFileSync(BASE_FILE, JSON.stringify({ date: new Date().toISOString(), node: process.version, results }, null, 2), 'utf8');
    out('');
    out('已保存为基线: tests/benchmark-baseline.json');
}
if (args.indexOf('--compare') >= 0 && fs.existsSync(BASE_FILE)) {
    const base = JSON.parse(fs.readFileSync(BASE_FILE, 'utf8')).results;
    out('');
    out('='.repeat(78));
    out('性能优化前后对比（任务书 §40；同机同参，中位数计时）');
    out('='.repeat(78));
    out('  项目                       优化前          优化后          变化');
    out('  ' + '-'.repeat(70));
    const rows = [
        ['单方案仿真', 'simulationTime', v => ms(v), false],
        ['Summary', 'summaryTime', v => ms(v), false],
        ['Finance', 'financeTime', v => ms(v), false],
        ['Chart数据(冷)', 'chartDataCold', v => ms(v), false],
        ['Chart数据(热)', 'chartDataWarm', v => ms(v), false],
        ['Chart数据(列)', 'chartDataColumn', v => ms(v), false],
        ['概览视图分配', 'overviewAllocArrays', v => v + ' 个', false],
        ['8760表格渲染(首帧)', 'tableRenderTime', v => ms(v), false],
        ['8760表格渲染(中位)', 'tableRenderMedian', v => ms(v), false],
        ['表格滚动重绘', 'tableScrollTime', v => ms(v), false],
        ['NSGA-II 20×10', 'optimizationTime', v => (v / 1000).toFixed(3) + ' s', false],
        ['缓存命中率', 'cacheHitRate', v => (v * 100).toFixed(1) + '%', true],
    ];
    for (const [label, key, fmt, higherBetter] of rows) {
        const b = base[key], a = results[key];
        if (b === undefined || a === undefined) continue;
        const delta = higherBetter ? (a - b) : (b - a);
        const pct = b !== 0 ? (delta / Math.abs(b) * 100) : 0;
        const sign = delta >= 0 ? '+' : '';
        out('  ' + label.padEnd(24) + fmt(b).padStart(12) + fmt(a).padStart(14) +
            '   ' + sign + pct.toFixed(1) + '%');
    }
}
process.stdout.write(lines.join('\n') + '\n');
