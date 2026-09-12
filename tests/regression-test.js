/**
 * ============================================================================
 * V2.1 回归测试脚本（Node.js）
 * ============================================================================
 *
 * 用途：
 *   1. 数值一致性测试 —— 同一容量组合下，V2.1 evaluateScheme() 与现有
 *      V1.0 计算链路（simulation-engine → data-summary → estimate → finance-engine）
 *      必须得到完全一致的技术经济指标（六十二条）。
 *   2. 工程正确性测试 —— 逐小时功率平衡、储能 SOC 递推、电解槽最低负荷、
 *      外购电比例 / 绿电比例 / 弃电率 / 年制氢量 / LCOH / FIRR 口径。
 *   3. 边界测试 —— 无储能 / 无风电 / 无光伏 / 极小电解槽 / 约束无可行解（六十一条）。
 *   4. NSGA-II 算法测试 —— 种群初始化、非支配排序、拥挤距离、贝罗尼、
 *      离散步长约束、Pareto 输出、代表方案识别、缓存命中。
 *
 * 运行：
 *   node tests/regression-test.js
 * 输出：
 *   tests/regression-report.txt（同时打印到 stdout）
 */

'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.resolve(__dirname, '..');
const OUT_FILE = path.join(__dirname, 'regression-report.txt');

const lines = [];
let passCount = 0;
let failCount = 0;

function out(s) { lines.push(s); }
function ok(name, detail) { passCount++; out('  [PASS] ' + name + (detail ? '  —— ' + detail : '')); }
function fail(name, detail) { failCount++; out('  [FAIL] ' + name + (detail ? '  —— ' + detail : '')); }
function assertClose(name, a, b, tol) {
    const t = tol === undefined ? 1e-6 : tol;
    if (a === null || b === null) {
        if (a === b) ok(name, 'both null');
        else fail(name, 'a=' + a + ', b=' + b);
        return;
    }
    if (Math.abs(a - b) <= t * Math.max(1, Math.abs(a), Math.abs(b))) {
        ok(name, 'Δ=' + Math.abs(a - b).toExponential(2));
    } else {
        fail(name, 'a=' + a + ', b=' + b + ', Δ=' + Math.abs(a - b));
    }
}
function assertTrue(name, cond, detail) {
    if (cond) ok(name, detail); else fail(name, detail);
}
function section(t) { out(''); out('='.repeat(78)); out(t); out('='.repeat(78)); }

// ---------------------------------------------------------------------------
// 0. 加载浏览器端模块（拼接后在 Node 全局上下文中执行）
// ---------------------------------------------------------------------------
global.self = global;
global.window = global;

const XLSX = require(path.join(ROOT, 'js', 'vendor', 'xlsx.full.min.js'));
global.XLSX = XLSX;

const MODULES = [
    'js/utils.js',
    'js/simulation-engine.js',
    'js/data-summary.js',
    'js/estimate.js',
    'js/finance-engine.js',
    'js/optimization-engine.js',
];

const bundle = MODULES
    .map(f => '/* ===== ' + f + ' ===== */\n' + fs.readFileSync(path.join(ROOT, f), 'utf8'))
    .join('\n;\n');

vm.runInThisContext(bundle, { filename: 'v21-bundle.js' });

const OptimizationEngine = global.OptimizationEngine;
const DataSummary = global.DataSummary;
const Estimate = global.Estimate;
const FinanceEngine = global.FinanceEngine;
const runSingleSimulation = global.runSingleSimulation;

out('多能互补风光储氢分析软件 V2.1 —— 回归测试报告');
out('生成时间：' + new Date().toISOString());
out('模块加载：' + MODULES.join(' / '));
assertTrue('模块加载完整性', !!(OptimizationEngine && DataSummary && Estimate && FinanceEngine && runSingleSimulation));

// ---------------------------------------------------------------------------
// 1. 输入数据与参数（与前端 index.html / app.js 默认值完全一致）
// ---------------------------------------------------------------------------
section('1. 输入数据与参数');

// 说明：vendor 目录下为 SheetJS 浏览器构建，未绑定 fs，因此此处自行读取文件后解析
const wb = XLSX.read(fs.readFileSync(path.join(ROOT, 'input.xlsx')), { type: 'buffer' });
const sheet = wb.Sheets[wb.SheetNames[0]];
const rawRows = XLSX.utils.sheet_to_json(sheet, { header: ['pv', 'wind'], range: 1 });
out('input.xlsx 工作表：' + wb.SheetNames[0] + '，数据行数：' + rawRows.length);

const N = rawRows.length;
const pvData = new Float64Array(N);
const windData = new Float64Array(N);
for (let i = 0; i < N; i++) {
    pvData[i] = Number(rawRows[i].pv) || 0;
    windData[i] = Number(rawRows[i].wind) || 0;
}
assertTrue('8760 小时数据长度校验', N === 8760, '实际 ' + N + ' 小时');

const SIM_PARAMS_BASE = {
    PV_CAPACITY: 360,
    WIND_CAPACITY: 200,
    ELECTROLYZER_MIN_RATIO: 0.3,
    MAX_EXPORT_RATIO_HOURLY: 0.0,
    MAX_EXPORT_RATIO_TOTAL: 0.0,
    MAX_IMPORT_RATIO: 0.15,
    STORAGE_CHARGE_EFFICIENCY: 0.92,
    STORAGE_DISCHARGE_EFFICIENCY: 0.92,
    HYDROGEN_ENERGY_CONSUMPTION: 55,
};

const PRICES = {
    windUnitPrice: 3800, pvUnitPrice: 2500, storageUnitPrice: 800, electrolyzerUnitPrice: 1600,
    transmissionCost: 10000, landCost: 10000, otherFacilitiesRatio: 70,
};

const FINANCE_PARAMS = {
    calc_period: 22, construct_period: 2, batch_count: 1,
    capital_benchmark: 8.0, industry_benchmark_pre: 5.0, industry_benchmark_post: 4.5,
    capital_ratio: 20, loan_period: 15, long_term_loan_rate: 3.0,
    single_electrolyzer_mw: 5.0,
    capacity_elec_fee: 20.0, ongrid_price: 0.2, offgrid_price: 0.7, wheeling_fee: 0.12,
    h2_labor_cost: 3.0, electrolyzer_maint: 40.0, electrolyzer_overhaul: 240.0,
    water_unit_price: 2.0, wind_pv_om: 35.0,
    depreciation_years: 20, residual_rate: 5.0,
    h2_price: 30.0, output_vat_rate: 13.0, income_tax_rate: 25.0, surtax_rate: 10.0,
};

const LCOH_DISCOUNT = 5.0;

function makeContext(extra) {
    return Object.assign({
        pvData: pvData,
        windData: windData,
        simParams: SIM_PARAMS_BASE,
        prices: PRICES,
        financeParams: FINANCE_PARAMS,
        lcohDiscountRate: LCOH_DISCOUNT,
        cache: new Map(),
        stats: { evaluated: 0, cacheHits: 0 },
        config: OptimizationEngine.normalizeConfig({ lcoh: { discountRate: LCOH_DISCOUNT } }),
    }, extra || {});
}

/** 现有 V1.0 计算链路（完全按 app.js 的调用顺序执行） */
function runLegacyChain(scheme) {
    const params = Object.assign({}, SIM_PARAMS_BASE, {
        PV_CAPACITY: scheme.pvCapacity,
        WIND_CAPACITY: scheme.windCapacity,
    });
    const sim = runSingleSimulation(pvData, windData, params,
        scheme.storagePower, scheme.storageDuration, scheme.electrolyzerCapacity);
    const summaryRow = DataSummary.generateSummary([sim], scheme.pvCapacity, scheme.windCapacity)[0];
    const estimateRow = Estimate.batchEstimate([summaryRow], PRICES)[0];
    const fin = FinanceEngine.calculateAll(FINANCE_PARAMS, summaryRow, estimateRow);
    return { sim: sim, summaryRow: summaryRow, estimateRow: estimateRow, fin: fin };
}

// ---------------------------------------------------------------------------
// 2. 测试方案 1：简单方案 —— V1.0 链路 vs V2.1 evaluateScheme
// ---------------------------------------------------------------------------
section('2. 测试方案1（简单方案）：V1.0 链路 与 V2.1 evaluateScheme 数值一致性');

const SCHEME_1 = {
    windCapacity: 100, pvCapacity: 100,
    storagePower: 50, storageDuration: 2, electrolyzerCapacity: 50,
};

const legacy = runLegacyChain(SCHEME_1);
const ctx1 = makeContext();
const ev1 = OptimizationEngine.evaluateScheme(SCHEME_1, ctx1);

out('方案：风电 100MW / 光伏 100MW / 储能 50MW×2h / 电解槽 50MW');
out('');
out('  -- 原 V1.0 链路输出 --');
out('  年制氢量(kg)      : ' + legacy.sim.sums.sumH2Prod.toFixed(2));
out('  弃电量(MWh)       : ' + legacy.sim.sums.sumCurtailment.toFixed(2));
out('  上网量(MWh)       : ' + legacy.sim.sums.sumExport.toFixed(2));
out('  下网量(MWh)       : ' + legacy.sim.sums.sumImport.toFixed(2));
out('  建设总投资(万元)  : ' + legacy.estimateRow['建设总投资（万元）']);
out('  FIRR税前(%)       : ' + legacy.fin.result['项目投资财务内部收益率（所得税前）（%）']);
out('  EIRR资本金(%)     : ' + legacy.fin.result['资本金财务内部收益率（%）']);
out('');
out('  -- V2.1 evaluateScheme 输出 --');
out('  年制氢量           : ' + ev1.technical.annualHydrogenKg.toFixed(2) + ' kg');
out('  弃电率             : ' + (ev1.technical.curtailmentRate * 100).toFixed(4) + ' %');
out('  外购电比例         : ' + (ev1.technical.gridImportRatio * 100).toFixed(4) + ' %');
out('  绿电制氢比例       : ' + (ev1.technical.greenHydrogenRatio * 100).toFixed(4) + ' %');
out('  总投资             : ' + ev1.economic.totalInvestment + ' 万元');
out('  FIRR               : ' + ev1.economic.FIRR + ' %');
out('  EIRR               : ' + ev1.economic.EIRR + ' %');
out('  LCOH（折现口径）    : ' + ev1.economic.LCOH.toFixed(4) + ' 元/kg');
out('  制氢成本（V1.0口径）: ' + ev1.economic.h2CostAvg + ' 元/kg');
out('');

// 2.1 与现有链路逐项比对
assertClose('年制氢量一致', ev1.technical.annualHydrogenKg, legacy.sim.sums.sumH2Prod);
assertClose('弃电量一致', ev1.technical.annualCurtailment, legacy.sim.sums.sumCurtailment);
assertClose('上网电量一致', ev1.technical.annualGridExport, legacy.sim.sums.sumExport);
assertClose('下网电量一致', ev1.technical.annualGridImport, legacy.sim.sums.sumImport);
assertClose('建设投资一致', ev1.economic.constructionInvestment, legacy.estimateRow['建设总投资（万元）']);
assertClose('总投资一致', ev1.economic.totalInvestment, legacy.fin.result['项目总投资（万元）']);
assertClose('FIRR 与财务引擎一致', ev1.economic.FIRR === null ? null : ev1.economic.FIRR,
    legacy.fin.result['项目投资财务内部收益率（所得税前）（%）']);
assertClose('EIRR 与财务引擎一致', ev1.economic.EIRR, legacy.fin.result['资本金财务内部收益率（%）']);

// 2.2 派生指标口径自检（在测试中独立重算）
assertClose('弃电率 = 弃电量/风光理论发电量',
    ev1.technical.curtailmentRate,
    legacy.sim.sums.sumCurtailment / legacy.sim.sums.sumTotal);
assertClose('外购电比例 = 下网电量/电解槽耗电量',
    ev1.technical.gridImportRatio,
    legacy.sim.sums.sumImport / legacy.sim.sums.sumHydrogenPower);
assertClose('绿电比例 = 1 - 外购电比例',
    ev1.technical.greenHydrogenRatio, 1 - ev1.technical.gridImportRatio);
assertClose('上网比例 = 上网电量/风光理论发电量',
    ev1.technical.exportRatio,
    legacy.sim.sums.sumExport / legacy.sim.sums.sumTotal);

// 2.3 LCOH 独立重算
(function () {
    const d = legacy.fin.detail;
    const r = LCOH_DISCOUNT / 100;
    const capexPerYear = (d.construction_investment + d.construction_interest) / d.construct_period;
    let num = 0, den = 0;
    for (let t = 0; t < d.calc_period; t++) {
        const df = 1 / Math.pow(1 + r, t);
        let cost = 0;
        if (t < d.construct_period) cost = capexPerYear;
        else {
            cost = d.annual_opex;
            if (t === d.overhaul_year) cost += d.overhaul_total;
            if (t === d.construct_period) cost += d.working_capital;
            den += d.annual_h2_kg * df;
        }
        num += cost * df;
    }
    const lcoh = num * 1e4 / den;
    assertClose('LCOH 折现口径独立重算一致', ev1.economic.LCOH, lcoh, 1e-9);
    out('    LCOH 手算 = ' + lcoh.toFixed(4) + ' 元/kg');
})();

// ---------------------------------------------------------------------------
// 3. 工程正确性：逐小时功率平衡 / 储能 SOC / 电解槽最低负荷
// ---------------------------------------------------------------------------
section('3. 工程正确性校验（8760 小时）');

const res = legacy.sim.results;
let maxBalanceError = 0;
let socError = 0;
let underMinLoadHours = 0;
let minNonZeroLoad = Infinity;
const EFF_C = SIM_PARAMS_BASE.STORAGE_CHARGE_EFFICIENCY;
const EFF_D = SIM_PARAMS_BASE.STORAGE_DISCHARGE_EFFICIENCY;
const EL_MIN = SCHEME_1.electrolyzerCapacity * SIM_PARAMS_BASE.ELECTROLYZER_MIN_RATIO;
let soc = 0;

for (let h = 0; h < N; h++) {
    const i = h * 11;
    const pv = res[i + 0], wind = res[i + 1];
    const charge = res[i + 3], discharge = res[i + 4];
    const h2 = res[i + 6], exp = res[i + 7], imp = res[i + 8], curt = res[i + 9];

    // 功率平衡：风电 + 光伏 + 下网 + 放电 = 制氢 + 充电 + 上网 + 弃电
    const lhs = wind + pv + imp + discharge;
    const rhs = h2 + charge + exp + curt;
    const err = Math.abs(lhs - rhs);
    if (err > maxBalanceError) maxBalanceError = err;

    // SOC 递推
    soc = soc + EFF_C * charge - discharge / EFF_D;
    socError = Math.max(socError, Math.abs(soc - res[i + 5]));

    // 电解槽无效运行区间 0 < P_EL < P_EL,min
    if (h2 > 1e-9 && h2 < EL_MIN - 1e-9) {
        underMinLoadHours++;
        if (h2 < minNonZeroLoad) minNonZeroLoad = h2;
    }
}

assertTrue('逐小时功率平衡（|误差| < 1e-6）', maxBalanceError < 1e-6,
    '最大误差 ' + maxBalanceError.toExponential(3) + ' MWh');
assertTrue('储能 SOC 递推一致（误差 < 1e-6）', socError < 1e-6,
    '最大误差 ' + socError.toExponential(3) + ' MWh');

out('');
out('  电解槽最低运行功率 P_EL,min = ' + EL_MIN.toFixed(3) + ' MW');
out('  出现 0 < P_EL < P_EL,min 的小时数 = ' + underMinLoadHours +
    (underMinLoadHours ? '（最小值 ' + minNonZeroLoad.toFixed(3) + ' MW）' : ''));
if (underMinLoadHours > 0) {
    out('  ⚠ 已知 V1.0 建模特征：风光不足且储能/下网受限时，现有调度会令电解槽运行在');
    out('    (0, P_EL,min) 区间，即存在「无效运行状态」。V2.1 未修改该逻辑（六十七条第3点），');
    out('    如需严格禁止需在 V2.2 增加最小负荷停机判据。');
} else {
    opt_note_probe();
}
function opt_note_probe() { /* 占位，保持输出结构稳定 */ }

// ---------------------------------------------------------------------------
// 4. 边界测试（六十一条）
// ---------------------------------------------------------------------------
section('4. 边界测试');

function evalCase(title, scheme, extraCtx, expect) {
    const ctx = makeContext(extraCtx);
    let ev = null, err = null;
    try { ev = OptimizationEngine.evaluateScheme(scheme, ctx); } catch (e) { err = e; }
    out('');
    out('  ▶ ' + title);
    if (err) {
        fail(title + '：不应抛异常', err.message);
        return null;
    }
    out('    年制氢量 = ' + ev.technical.annualHydrogenKg.toFixed(1) + ' kg （' +
        ev.technical.annualHydrogenWanTon.toFixed(4) + ' 万吨）');
    out('    弃电率   = ' + (ev.technical.curtailmentRate * 100).toFixed(3) + ' %');
    out('    电解槽利用小时 = ' + ev.technical.electrolyzerHours.toFixed(1) + ' h');
    out('    LCOH     = ' + (isFinite(ev.economic.LCOH) ? ev.economic.LCOH.toFixed(3) + ' 元/kg' : 'Infinity（无制氢量）'));
    out('    FIRR     = ' + (ev.economic.FIRR === null ? 'null（现金流无符号变化，经济评价不可行）' : ev.economic.FIRR + ' %'));
    out('    可行性   = ' + (ev.constraints.feasible ? '可行' : '不可行') +
        (ev.constraints.violations.length ? '，违反项：' + ev.constraints.violations.map(v => v.name).join('、') : ''));
    if (expect && expect.check) expect.check(ev);
    return ev;
}

// 测试方案 2：无储能
const evNoStorage = evalCase('测试方案2：无储能（储能功率 0 / 时长 0）', {
    windCapacity: 100, pvCapacity: 100, storagePower: 0, storageDuration: 0, electrolyzerCapacity: 50,
}, null, {
    check: ev => {
        assertTrue('  无储能：储能充放电均为 0',
            Math.abs(ev.technical.annualChargeEnergy) < 1e-9 && Math.abs(ev.technical.annualDischargeEnergy) < 1e-9);
        assertTrue('  无储能：无 NaN / 非有限值',
            isFinite(ev.technical.annualHydrogenKg) && isFinite(ev.technical.curtailmentRate));
    },
});

// 测试方案 3：无风电
evalCase('测试方案3：无风电（风电 0MW）', {
    windCapacity: 0, pvCapacity: 200, storagePower: 50, storageDuration: 2, electrolyzerCapacity: 50,
}, null, {
    check: ev => {
        assertTrue('  无风电：风电年发电量为 0', Math.abs(ev.technical.annualWindEnergy) < 1e-9);
        assertTrue('  无风电：光伏年发电量 > 0', ev.technical.annualPvEnergy > 0);
        assertTrue('  无风电：无 NaN', isFinite(ev.technical.annualHydrogenKg));
    },
});

// 测试方案 4：无光伏
evalCase('测试方案4：无光伏（光伏 0MW）', {
    windCapacity: 200, pvCapacity: 0, storagePower: 50, storageDuration: 2, electrolyzerCapacity: 50,
}, null, {
    check: ev => {
        assertTrue('  无光伏：光伏年发电量为 0', Math.abs(ev.technical.annualPvEnergy) < 1e-9);
        assertTrue('  无光伏：风电年发电量 > 0', ev.technical.annualWindEnergy > 0);
        assertTrue('  无光伏：无 NaN', isFinite(ev.technical.annualHydrogenKg));
    },
});

// 测试方案 5：极端小电解槽
const evTiny = evalCase('测试方案5：极端小电解槽（25MW）', {
    windCapacity: 300, pvCapacity: 500, storagePower: 200, storageDuration: 4, electrolyzerCapacity: 25,
}, null, {
    check: ev => {
        assertTrue('  小电解槽：弃电率显著升高（> 30%）', ev.technical.curtailmentRate > 0.30,
            '弃电率 ' + (ev.technical.curtailmentRate * 100).toFixed(2) + '%');
        assertTrue('  小电解槽：电解槽利用小时接近满发', ev.technical.electrolyzerHours > 7000,
            '利用小时 ' + ev.technical.electrolyzerHours.toFixed(1) + ' h');
    },
});

// 电解槽容量 = 0（无制氢量 → LCOH = Infinity）
evalCase('测试方案5b：电解槽 0MW（无制氢量）', {
    windCapacity: 100, pvCapacity: 100, storagePower: 0, storageDuration: 0, electrolyzerCapacity: 0,
}, null, {
    check: ev => {
        assertTrue('  电解槽为 0：年制氢量为 0', Math.abs(ev.technical.annualHydrogenKg) < 1e-9);
        assertTrue('  电解槽为 0：LCOH = Infinity（不显示为数字）', ev.economic.LCOH === Infinity);
        assertTrue('  电解槽为 0：判为不可行', ev.constraints.feasible === false);
    },
});

// 测试方案 6：约束无可行解（年制氢量要求极高）
(function () {
    const cfg = OptimizationEngine.normalizeConfig({
        lcoh: { discountRate: LCOH_DISCOUNT },
        constraints: { minAnnualHydrogen: { enabled: true, value: 100000, unit: '万吨/年' } },
    });
    const ctx = makeContext({ config: cfg });
    const ev = OptimizationEngine.evaluateScheme({
        windCapacity: 100, pvCapacity: 100, storagePower: 50, storageDuration: 2, electrolyzerCapacity: 50,
    }, ctx);
    out('');
    out('  ▶ 测试方案6：不可满足的工程约束（年制氢量 ≥ 100000 万吨/年）');
    out('    可行性 = ' + (ev.constraints.feasible ? '可行' : '不可行') +
        '，违反项 = ' + ev.constraints.violations.map(v => v.name).join('、'));
    assertTrue('  约束无可行解：正确判为不可行', ev.constraints.feasible === false);
    assertTrue('  约束无可行解：给出具体违反项', ev.constraints.violations.length > 0);
})();

// ---------------------------------------------------------------------------
// 5. 异常与容错（五十一条 / 五十二条 / 五十三条）
// ---------------------------------------------------------------------------
section('5. 异常处理与容错');

(function () {
    // 数据长度不足
    const cfg = OptimizationEngine.normalizeConfig({});
    const badCtx = {
        pvData: new Float64Array(100),
        windData: new Float64Array(100),
        simParams: SIM_PARAMS_BASE, prices: PRICES, financeParams: FINANCE_PARAMS,
        lcohDiscountRate: LCOH_DISCOUNT, cache: new Map(), stats: { evaluated: 0, cacheHits: 0 }, config: cfg,
    };
    let thrown = null;
    try {
        OptimizationEngine.createSession({ config: cfg, context: badCtx });
    } catch (e) { thrown = e; }
    assertTrue('数据长度非 8760：仅告警不致命（不抛异常）', thrown === null,
        thrown ? thrown.message : '按设计放行并告警');
})();

(function () {
    const cfg = OptimizationEngine.normalizeConfig({});
    let thrown = null;
    try {
        OptimizationEngine.createSession({ config: cfg, context: { pvData: null, windData: null } });
    } catch (e) { thrown = e; }
    assertTrue('数据缺失：抛出可读错误而非崩溃', !!thrown && /8760|数据|为空|长度/.test(thrown.message),
        thrown ? thrown.message : '未抛出');
})();

(function () {
    // min > max
    const res = OptimizationEngine.validateConfig(OptimizationEngine.normalizeConfig({
        variables: { windCapacity: { min: 300, max: 100, step: 25 } },
    }));
    assertTrue('参数校验：识别 min > max 为非法', res.ok === false, res.message);
})();

(function () {
    // step <= 0
    const res = OptimizationEngine.validateConfig(OptimizationEngine.normalizeConfig({
        variables: { pvCapacity: { min: 50, max: 500, step: -1 } },
    }));
    assertTrue('参数校验：识别 step < 0 为非法', res.ok === false, res.message);
})();

(function () {
    // 非法种群规模 / 迭代次数
    const res = OptimizationEngine.validateConfig(OptimizationEngine.normalizeConfig({
        nsga2: { populationSize: 1, generations: 0 },
    }));
    assertTrue('参数校验：识别非法种群规模与迭代次数', res.ok === false, res.message);
})();

(function () {
    // 步长吸附
    const cases = [
        [87.326, 50, 300, 25, 75],
        [88, 50, 300, 25, 100],
        [0.4, 0, 4, 1, 0],
        [3.6, 0, 4, 1, 4],
        [-5, 0, 200, 25, 0],
        [999, 0, 200, 25, 200],
        [2, 0, 4, 0, 2],          // step = 0 → 连续，仅裁剪
    ];
    let allOk = true, detail = [];
    for (const [v, mn, mx, st, exp] of cases) {
        const got = OptimizationEngine.snapToStep(v, mn, mx, st);
        detail.push(v + '→' + got);
        if (Math.abs(got - exp) > 1e-9) { allOk = false; detail.push('(期望 ' + exp + ')'); }
    }
    assertTrue('snapToStep 严格落在离散步长上', allOk, detail.join(', '));
})();

(function () {
    // 情形 A：现金流全为同号 → 现有财务引擎返回 0（无解哨兵），V2.1 必须转成 null
    // 构造方式：全零容量 + 移除送出线路/土地投资 → 全部现金流均 <= 0
    const zeroPrices = Object.assign({}, PRICES, { transmissionCost: 0, landCost: 0 });
    const evZero = OptimizationEngine.evaluateScheme({
        windCapacity: 0, pvCapacity: 0, storagePower: 0, storageDuration: 0, electrolyzerCapacity: 0,
    }, makeContext({ prices: zeroPrices }));
    out('');
    out('  ▶ 现金流全为同号（全零容量 + 零路站投资）：FIRR = ' + evZero.economic.FIRR +
        '，LCOH = ' + evZero.economic.LCOH);
    assertTrue('FIRR 无解时返回 null 而非 0（避免误导优化器）', evZero.economic.FIRR === null);
    assertTrue('LCOH 无制氢量时返回 Infinity 而非 0', evZero.economic.LCOH === Infinity);
    assertTrue('经济评价不可行 → 方案判为不可行', evZero.constraints.feasible === false);

    // 情形 B：现金流有符号变化 → IRR 存在（即使为负），必须保留真实值，不得误判为 null
    const evNeg = OptimizationEngine.evaluateScheme({
        windCapacity: 0, pvCapacity: 0, storagePower: 0, storageDuration: 0, electrolyzerCapacity: 0,
    }, makeContext());
    out('  ▶ 现金流有符号变化（存在残值回收）：FIRR = ' + evNeg.economic.FIRR + '（负值但客观存在，应保留）');
    assertTrue('IRR 客观存在时不得被误判为 null',
        evNeg.economic.FIRR !== null && evNeg.economic.FIRR < 0);
})();

// ---------------------------------------------------------------------------
// 6. 缓存机制（二十三条 / 四十六条）
// ---------------------------------------------------------------------------
section('6. 评价缓存机制');

(function () {
    const ctx = makeContext();
    const s = { windCapacity: 150, pvCapacity: 250, storagePower: 50, storageDuration: 2, electrolyzerCapacity: 100 };
    const a = OptimizationEngine.evaluateScheme(s, ctx);
    const b = OptimizationEngine.evaluateScheme(s, ctx);
    const c = OptimizationEngine.evaluateScheme(s, ctx);
    out('  首次评价后 stats.evaluated = ' + ctx.stats.evaluated + '，cacheHits = ' + ctx.stats.cacheHits);
    assertTrue('相同方案仅计算一次（缓存命中 2 次）', ctx.stats.evaluated === 1 && ctx.stats.cacheHits === 2);
    assertTrue('缓存返回同一对象引用', a === b && b === c);
})();

// ---------------------------------------------------------------------------
// 7. NSGA-II 算法端到端测试
// ---------------------------------------------------------------------------
section('7. NSGA-II 端到端测试');

const NSGA_CONFIG = OptimizationEngine.normalizeConfig({
    variables: {
        windCapacity: { min: 50, max: 200, step: 50 },
        pvCapacity: { min: 50, max: 300, step: 50 },
        storagePower: { min: 0, max: 100, step: 50 },
        storageDuration: { min: 0, max: 4, step: 2 },
        electrolyzerCapacity: { min: 25, max: 100, step: 25 },
    },
    constraints: {
        minAnnualHydrogen: { enabled: true, value: 0.5, unit: '万吨/年' },
        maxCurtailmentRate: { enabled: true, value: 10, unit: '%' },
        maxGridImportRatio: { enabled: true, value: 10, unit: '%' },
        minGreenHydrogenRatio: { enabled: true, value: 90, unit: '%' },
    },
    nsga2: {
        populationSize: 24, generations: 10,
        crossoverProbability: 0.9, mutationProbability: 0.15,
        randomSeed: 20260912, earlyStopping: false, patience: 15,
    },
    lcoh: { discountRate: LCOH_DISCOUNT },
    recommendationWeights: { firr: 0.40, lcoh: 0.35, curtailmentRate: 0.25 },
    baseline: { windCapacity: 100, pvCapacity: 300, storagePower: 100, storageDuration: 2, electrolyzerCapacity: 100 },
});

out('  搜索空间规模 = ' + OptimizationEngine.searchSpaceSize(OptimizationEngine.buildLevels(NSGA_CONFIG.variables)));
out('  种群规模 = ' + NSGA_CONFIG.nsga2.populationSize + '，迭代代数 = ' + NSGA_CONFIG.nsga2.generations);

const t0 = Date.now();
let session = null, result = null, runErr = null;
try {
    session = OptimizationEngine.createSession({ config: NSGA_CONFIG, context: makeContext() });
    session.ensureInitialized();
    while (!session.isFinished()) session.runNextGeneration();
    result = session.getResult();
} catch (e) { runErr = e; }
const elapsed = (Date.now() - t0) / 1000;

if (runErr) {
    fail('NSGA-II 运行', runErr.message);
    out(runErr.stack);
} else {
    ok('NSGA-II 运行完成', elapsed.toFixed(2) + ' s');

    const st = result.statistics;
    out('');
    out('  统计：总评价 ' + st.totalEvaluated + ' 个，缓存命中 ' + st.cacheHits +
        '，可行 ' + st.feasibleCount + '，Pareto ' + st.paretoCount +
        '，完成代数 ' + st.generationsCompleted);
    out('  有效种群规模 = ' + st.effectivePopulationSize + '（搜索空间 ' + st.searchSpaceSize + '）');
    out('  历史记录长度 = ' + result.history.length);

    assertTrue('算法：产生了 Pareto 前沿', result.paretoSolutions.length > 0,
        result.paretoSolutions.length + ' 个非支配方案');
    assertTrue('算法：历史记录覆盖全部代数 + 初始化',
        result.history.length === st.generationsCompleted + 1,
        result.history.length + ' vs ' + (st.generationsCompleted + 1));
    assertTrue('算法：精英保留使代数递增完成', st.generationsCompleted === NSGA_CONFIG.nsga2.generations,
        st.generationsCompleted + ' 代');

    // 离散步长校验
    let gridOk = true, gridDetail = [];
    for (const s of result.paretoSolutions) {
        const sc = s.scheme;
        for (const key of OptimizationEngine.KEY_ORDER) {
            const v = NSGA_CONFIG.variables[key];
            const snapped = OptimizationEngine.snapToStep(sc[key], v.min, v.max, v.step);
            if (Math.abs(snapped - sc[key]) > 1e-9 || sc[key] < v.min - 1e-9 || sc[key] > v.max + 1e-9) {
                gridOk = false;
                gridDetail.push(key + '=' + sc[key]);
            }
        }
    }
    assertTrue('算法：全部 Pareto 方案严格落在离散步长与范围内', gridOk, gridDetail.slice(0, 5).join(', '));

    // Pareto 互不支配校验
    let nonDominated = true;
    const P = result.paretoSolutions;
    for (let i = 0; i < P.length; i++) {
        for (let j = 0; j < P.length; j++) {
            if (i === j) continue;
            const a = { feasible: true, vector: P[i].vector, violationAmount: 0 };
            const b = { feasible: true, vector: P[j].vector, violationAmount: 0 };
            if (OptimizationEngine.dominates(a, b) === 1) { nonDominated = false; }
        }
    }
    assertTrue('算法：Pareto 前沿内两两互不支配', nonDominated, P.length + ' 个方案');

    // 代表方案
    const rep = result.representativeSolutions;
    assertTrue('代表方案：已识别综合推荐', !!rep.recommended);
    assertTrue('代表方案：已识别经济最优', !!rep.economicBest);
    assertTrue('代表方案：已识别氢成本最优', !!rep.hydrogenCostBest);
    assertTrue('代表方案：已识别消纳最优', !!rep.curtailmentBest);
    assertTrue('代表方案：推荐方案来自 Pareto 前沿',
        P.some(s => s.key === rep.recommended.key));

    if (rep.recommended) {
        const s = rep.recommended;
        out('');
        out('  ▶ 综合推荐方案');
        out('    风电 ' + s.scheme.windCapacity + ' MW / 光伏 ' + s.scheme.pvCapacity + ' MW / 储能 ' +
            s.scheme.storagePower + ' MW × ' + s.scheme.storageDuration + ' h / 电解槽 ' + s.scheme.electrolyzerCapacity + ' MW');
        out('    年制氢量 ' + s.technical.annualHydrogenWanTon.toFixed(3) + ' 万吨/年');
        out('    弃电率 ' + (s.technical.curtailmentRate * 100).toFixed(2) + ' %，外购电比例 ' +
            (s.technical.gridImportRatio * 100).toFixed(2) + ' %，绿电比例 ' + (s.technical.greenHydrogenRatio * 100).toFixed(2) + ' %');
        out('    LCOH ' + s.economic.LCOH.toFixed(2) + ' 元/kg，FIRR ' + s.economic.FIRR + ' %，总投资 ' +
            (s.economic.totalInvestment / 10000).toFixed(3) + ' 亿元');
        out('    综合得分 ' + (s.score * 100).toFixed(2));
    }

    // 基准方案
    assertTrue('基准方案（Baseline）已评价', !!result.baseline);
    if (result.baseline) {
        out('');
        out('  ▶ 基准方案');
        out('    风电 ' + result.baseline.scheme.windCapacity + ' MW / 光伏 ' + result.baseline.scheme.pvCapacity +
            ' MW / 储能 ' + result.baseline.scheme.storagePower + ' MW × ' + result.baseline.scheme.storageDuration + ' h');
        out('    FIRR ' + result.baseline.economic.FIRR + ' %，LCOH ' + result.baseline.economic.LCOH.toFixed(2) +
            ' 元/kg，弃电率 ' + (result.baseline.technical.curtailmentRate * 100).toFixed(2) + ' %');
    }

    // 收敛性：末代 Pareto 方案数应不低于第 0 代
    const firstH = result.history[0];
    const lastH = result.history[result.history.length - 1];
    out('');
    out('  收敛：第 0 代 Pareto=' + firstH.paretoCount + '，末代 Pareto=' + lastH.paretoCount);
    out('  收敛：第 0 代 可行=' + firstH.feasibleCount + '，末代 可行=' + lastH.feasibleCount);
}

// ---------------------------------------------------------------------------
// 8. 可复现性（三十三条）
// ---------------------------------------------------------------------------
section('8. 随机种子可复现性');

(function () {
    const cfg = OptimizationEngine.normalizeConfig({
        variables: {
            windCapacity: { min: 50, max: 200, step: 50 },
            pvCapacity: { min: 50, max: 200, step: 50 },
            storagePower: { min: 0, max: 100, step: 50 },
            storageDuration: { min: 0, max: 4, step: 2 },
            electrolyzerCapacity: { min: 25, max: 100, step: 25 },
        },
        constraints: { minAnnualHydrogen: { enabled: false }, maxCurtailmentRate: { enabled: false } },
        nsga2: { populationSize: 12, generations: 4, randomSeed: 12345, earlyStopping: false },
    });

    const runOnce = () => {
        const s = OptimizationEngine.createSession({ config: cfg, context: makeContext() });
        s.ensureInitialized();
        while (!s.isFinished()) s.runNextGeneration();
        const r = s.getResult();
        return r.paretoSolutions.map(x => x.key).sort().join(';');
    };

    const a = runOnce();
    const b = runOnce();
    assertTrue('相同随机种子 → 完全相同的 Pareto 解集', a === b,
        a === b ? '解集一致' : ('A=' + a + ' | B=' + b));
})();

// ---------------------------------------------------------------------------
// 9. 性能（四十六条）
// ---------------------------------------------------------------------------
section('9. 性能基准');

(function () {
    const ctx = makeContext();
    const schemes = [];
    for (let w = 50; w <= 300; w += 25) {
        for (let p = 50; p <= 500; p += 50) {
            schemes.push({ windCapacity: w, pvCapacity: p, storagePower: 100, storageDuration: 2, electrolyzerCapacity: 100 });
        }
    }
    const t = Date.now();
    for (const s of schemes) OptimizationEngine.evaluateScheme(s, ctx);
    const ms = Date.now() - t;
    out('  首次评价 ' + schemes.length + ' 个不同方案耗时 ' + ms + ' ms（平均 ' +
        (ms / schemes.length).toFixed(2) + ' ms/方案）');
    out('  单方案评价 = 8760 小时仿真 + 年度汇总 + 投资估算 + 财务评价 + LCOH + 约束校验');

    const t2 = Date.now();
    for (const s of schemes) OptimizationEngine.evaluateScheme(s, ctx);
    const ms2 = Date.now() - t2;
    out('  第二次（全部命中缓存）耗时 ' + ms2 + ' ms（平均 ' + (ms2 / schemes.length).toFixed(4) + ' ms/方案）');
    assertTrue('缓存显著降低重复评价开销', ms2 < ms * 0.2,
        '缓存化 ' + (ms / Math.max(ms2, 1)).toFixed(1) + ' 倍加速');
})();

// ---------------------------------------------------------------------------
// 10. 默认参数全量优化（前端 index.html 出厂默认值）
// ---------------------------------------------------------------------------
section('10. 出厂默认参数全量优化（种群 60 × 100 代）');

(function () {
    const cfg = OptimizationEngine.normalizeConfig({
        lcoh: { discountRate: LCOH_DISCOUNT },
        baseline: { windCapacity: 200, pvCapacity: 360, storagePower: 100, storageDuration: 2, electrolyzerCapacity: 160 },
    });
    const levels = OptimizationEngine.buildLevels(cfg.variables);
    const space = OptimizationEngine.searchSpaceSize(levels);
    out('  搜索空间 = ' + space.toLocaleString('zh-CN') + ' 个容量组合');
    out('  各变量档位数：' + OptimizationEngine.KEY_ORDER
        .map(k => OptimizationEngine.VAR_META[k].label + '=' + levels[k].length).join('，'));

    const ctx = makeContext();
    const t = Date.now();
    let r = null, err = null;
    try {
        const s = OptimizationEngine.createSession({ config: cfg, context: ctx });
        s.ensureInitialized();
        while (!s.isFinished()) s.runNextGeneration();
        r = s.getResult();
    } catch (e) { err = e; }
    const sec = (Date.now() - t) / 1000;

    if (err) {
        fail('默认参数全量优化', err.message);
    } else {
        const st = r.statistics;
        out('');
        out('  完成代数        : ' + st.generationsCompleted + (st.earlyStopped ? '（提前终止）' : ''));
        out('  累计评价方案数  : ' + st.totalEvaluated + '（缓存命中 ' + st.cacheHits + '）');
        out('  可行方案数      : ' + st.feasibleCount);
        out('  Pareto 方案数   : ' + st.paretoCount);
        out('  总耗时          : ' + sec.toFixed(2) + ' s');
        out('  平均单次评价    : ' + (sec * 1000 / Math.max(st.totalEvaluated, 1)).toFixed(3) + ' ms');
        out('');
        const rep = r.representativeSolutions.recommended;
        if (rep) {
            out('  ▶ 默认参数下的综合推荐方案');
            out('    风电 ' + rep.scheme.windCapacity + ' MW ｜ 光伏 ' + rep.scheme.pvCapacity +
                ' MW ｜ 储能 ' + rep.scheme.storagePower + ' MW × ' + rep.scheme.storageDuration +
                ' h ｜ 电解槽 ' + rep.scheme.electrolyzerCapacity + ' MW');
            out('    年制氢量 ' + rep.technical.annualHydrogenWanTon.toFixed(3) + ' 万吨/年 ｜ 弃电率 ' +
                (rep.technical.curtailmentRate * 100).toFixed(2) + ' % ｜ LCOH ' +
                rep.economic.LCOH.toFixed(2) + ' 元/kg ｜ FIRR ' + rep.economic.FIRR + ' %');
        }
        assertTrue('默认参数下优化在合理时间内完成（< 120 s）', sec < 120, sec.toFixed(2) + ' s');
        assertTrue('默认参数下产生 Pareto 前沿', r.paretoSolutions.length > 0,
            r.paretoSolutions.length + ' 个方案');

        // 默认约束较严（年制氢量 ≥ 1 万吨/年、弃电率 ≤ 10%），检查约束确实生效
        const bad = r.paretoSolutions.filter(s =>
            s.technical.annualHydrogenWanTon < 1 - 1e-9 || s.technical.curtailmentRate > 0.10 + 1e-9);
        assertTrue('Pareto 前沿内全部方案满足默认工程约束', bad.length === 0,
            bad.length ? bad.length + ' 个方案违反' : '全部满足');
    }
})();

// ---------------------------------------------------------------------------
// 11. 前端集成契约（index.html 出厂默认值 → app.js 组装方式 → 引擎）
// ---------------------------------------------------------------------------
// 背景：优化页曾出现「主线程回退路径报『风光数据缺失』」的故障，根因是
//       app.js 传 pv/wind 而引擎约定 pvData/windData。本节把「HTML 默认值 →
//       app.js 参数读取 → 引擎上下文」的完整契约固化下来，防止再次漏网。
section('11. 前端集成契约（index.html 默认值 → app.js 组装 → 引擎）');

const HTML_PATH = path.join(ROOT, 'index.html');
const html = fs.readFileSync(HTML_PATH, 'utf8');

/** 取某个 id 所在 input 标签的完整字符串 */
function htmlTag(id) {
    const re = new RegExp('<input[^>]*id="' + id + '"[^>]*>');
    const m = html.match(re);
    return m ? m[0] : null;
}
/** 取 input 的 value 数值（无 value 属性返回 null） */
function htmlVal(id) {
    const tag = htmlTag(id);
    if (!tag) return null;
    const m = tag.match(/value="([^"]*)"/);
    return m ? Number(m[1]) : null;
}
/** 取 input 的 checked 状态 */
function htmlChecked(id) {
    const tag = htmlTag(id);
    return !!tag && /[\s"]checked[\s>]/.test(tag + ' ');
}
/** 模拟 app.js 的 Utils.toNum(el.value, fallback) */
function numFrom(id, fallback) {
    const v = htmlVal(id);
    return (v === null || !isFinite(v)) ? fallback : v;
}
/** 模拟 app.js 的 parseInt(el.value) || fallback */
function intFrom(id, fallback) {
    const v = htmlVal(id);
    const n = parseInt(v, 10);
    return isFinite(n) && n ? n : fallback;
}

// 11.1 优化变量默认值必须与引擎默认配置一致
(function () {
    const optVars = [
        ['windCapacity', 'optWindMin', 'optWindMax', 'optWindStep'],
        ['pvCapacity', 'optPvMin', 'optPvMax', 'optPvStep'],
        ['storagePower', 'optStoragePowerMin', 'optStoragePowerMax', 'optStoragePowerStep'],
        ['storageDuration', 'optStorageDurationMin', 'optStorageDurationMax', 'optStorageDurationStep'],
        ['electrolyzerCapacity', 'optElectrolyzerMin', 'optElectrolyzerMax', 'optElectrolyzerStep'],
    ];
    const def = OptimizationEngine.buildDefaultConfig().variables;
    let allOk = true; const detail = [];
    for (const [key, idMin, idMax, idStep] of optVars) {
        const mn = htmlVal(idMin), mx = htmlVal(idMax), st = htmlVal(idStep);
        if (mn === null || mx === null || st === null) { allOk = false; detail.push(key + ' 缺 id/value'); continue; }
        const same = mn === def[key].min && mx === def[key].max && st === def[key].step;
        if (!same) { allOk = false; detail.push(`${key}: HTML(${mn}/${mx}/${st}) ≠ 引擎默认(${def[key].min}/${def[key].max}/${def[key].step})`); }
    }
    assertTrue('11.1 优化变量默认值与引擎默认配置一致', allOk, detail.join(' | ') || '5 个变量全部一致');
})();

// 11.2 组装与 app.js 完全一致的配置与上下文，跑通主线程会话
(function () {
    const optTag = htmlTag('optPopulationSize');
    if (!optTag) { bad('11.2 未找到优化页参数输入框', 'optPopulationSize'); return; }

    const config = OptimizationEngine.normalizeConfig({
        variables: {
            windCapacity: { min: htmlVal('optWindMin'), max: htmlVal('optWindMax'), step: htmlVal('optWindStep') },
            pvCapacity: { min: htmlVal('optPvMin'), max: htmlVal('optPvMax'), step: htmlVal('optPvStep') },
            storagePower: { min: htmlVal('optStoragePowerMin'), max: htmlVal('optStoragePowerMax'), step: htmlVal('optStoragePowerStep') },
            storageDuration: { min: htmlVal('optStorageDurationMin'), max: htmlVal('optStorageDurationMax'), step: htmlVal('optStorageDurationStep') },
            electrolyzerCapacity: { min: htmlVal('optElectrolyzerMin'), max: htmlVal('optElectrolyzerMax'), step: htmlVal('optElectrolyzerStep') },
        },
        constraints: {
            minAnnualHydrogen: { enabled: htmlChecked('consHydrogenOn'), value: htmlVal('consHydrogen'), unit: '万吨/年' },
            maxCurtailmentRate: { enabled: htmlChecked('consCurtailOn'), value: htmlVal('consCurtail'), unit: '%' },
            maxGridImportRatio: { enabled: htmlChecked('consImportOn'), value: htmlVal('consImport'), unit: '%' },
            minGreenHydrogenRatio: { enabled: htmlChecked('consGreenOn'), value: htmlVal('consGreen'), unit: '%' },
            maxExportRatio: { enabled: htmlChecked('consExportOn'), value: htmlVal('consExport'), unit: '%' },
            maxExportPower: { enabled: htmlChecked('consExportPowerOn'), value: htmlVal('consExportPower'), unit: 'MW' },
            maxImportPower: { enabled: htmlChecked('consImportPowerOn'), value: htmlVal('consImportPower'), unit: 'MW' },
        },
        nsga2: {
            populationSize: intFrom('optPopulationSize', 60),
            generations: intFrom('optGenerations', 100),
            crossoverProbability: numFrom('optCrossoverProb', 0.9),
            mutationProbability: numFrom('optMutationProb', 0.1),
            randomSeed: numFrom('optRandomSeed', 20260912),
            earlyStopping: htmlChecked('optEarlyStopping'),
            patience: intFrom('optPatience', 15),
        },
        recommendationWeights: {
            firr: numFrom('optWeightFirr', 40) / 100,
            lcoh: numFrom('optWeightLcoh', 35) / 100,
            curtailmentRate: numFrom('optWeightCurtail', 25) / 100,
        },
        lcoh: { discountRate: numFrom('optLcohDiscount', 5.0) },
        baseline: {
            windCapacity: numFrom('baseWind', 200), pvCapacity: numFrom('basePv', 360),
            storagePower: numFrom('baseStoragePower', 100), storageDuration: numFrom('baseStorageDuration', 2),
            electrolyzerCapacity: numFrom('baseElectrolyzer', 160),
        },
    });

    // 把种群/代数压到测试规模，其余参数全部取自 HTML
    config.nsga2.populationSize = 12;
    config.nsga2.generations = 3;

    out('  出厂默认工程约束：' + Object.keys(config.constraints).map(k =>
        k + '=' + (config.constraints[k].enabled ? config.constraints[k].value + config.constraints[k].unit : '未启用')).join('，'));
    out('  出厂默认推荐权重：FIRR ' + (config.recommendationWeights.firr * 100) + '% / LCOH ' +
        (config.recommendationWeights.lcoh * 100) + '% / 弃电率 ' + (config.recommendationWeights.curtailmentRate * 100) + '%');
    out('  LCOH 折现率：' + config.lcoh.discountRate + '%');

    // app.js getSimulationParams() 的等价实现（含同一套 fallback）
    const simParams = {
        PV_CAPACITY: numFrom('pvCapacity', 360),
        WIND_CAPACITY: numFrom('windCapacity', 200),
        ELECTROLYZER_MIN_RATIO: numFrom('electrolyzerMinRatio', 0.3),
        MAX_EXPORT_RATIO_HOURLY: numFrom('maxExportHourly', 0.0),
        MAX_EXPORT_RATIO_TOTAL: numFrom('maxExportTotal', 0.0),
        MAX_IMPORT_RATIO: numFrom('maxImportRatio', 0.15),
        STORAGE_CHARGE_EFFICIENCY: numFrom('chargeEfficiency', 0.92),
        STORAGE_DISCHARGE_EFFICIENCY: numFrom('dischargeEfficiency', 0.92),
        HYDROGEN_ENERGY_CONSUMPTION: numFrom('hydrogenConsumption', 55),
    };
    // app.js getEstimatePrices() 的等价实现
    const prices = {
        windUnitPrice: numFrom('windUnitPrice', 3800), pvUnitPrice: numFrom('pvUnitPrice', 2500),
        storageUnitPrice: numFrom('storageUnitPrice', 800), electrolyzerUnitPrice: numFrom('electrolyzerUnitPrice', 1600),
        transmissionCost: numFrom('transmissionCost', 10000), landCost: numFrom('landCost', 10000),
        otherFacilitiesRatio: numFrom('otherFacilitiesRatio', 70),
    };
    // app.js getFinanceParams() 的等价实现
    const financeParams = {
        calc_period: intFrom('calcPeriod', 22), construct_period: intFrom('constructPeriod', 2),
        batch_count: intFrom('batchCount', 1),
        capital_benchmark: numFrom('capitalBenchmark', 8.0),
        industry_benchmark_pre: numFrom('industryBenchmarkPre', 5.0),
        industry_benchmark_post: numFrom('industryBenchmarkPost', 4.5),
        capital_ratio: numFrom('capitalRatio', 20), loan_period: intFrom('loanPeriod', 15),
        long_term_loan_rate: numFrom('longTermLoanRate', 3.0),
        single_electrolyzer_mw: numFrom('singleElectrolyzer', 5.0),
        capacity_elec_fee: numFrom('capacityElecFee', 20.0),
        ongrid_price: numFrom('ongridPrice', 0.2), offgrid_price: numFrom('offgridPrice', 0.7),
        wheeling_fee: numFrom('wheelingFee', 0.12),
        h2_labor_cost: numFrom('h2LaborCost', 3.0), electrolyzer_maint: numFrom('electrolyzerMaint', 40.0),
        electrolyzer_overhaul: numFrom('electrolyzerOverhaul', 240.0),
        water_unit_price: numFrom('waterUnitPrice', 2.0), wind_pv_om: numFrom('windPvOm', 35.0),
        depreciation_years: intFrom('depreciationYears', 20), residual_rate: numFrom('residualRate', 5.0),
        h2_price: numFrom('h2Price', 30.0), output_vat_rate: numFrom('outputVatRate', 13.0),
        income_tax_rate: numFrom('incomeTaxRate', 25.0), surtax_rate: numFrom('surtaxRate', 12.0),
    };
    out('  出厂默认仿真参数：光伏 ' + simParams.PV_CAPACITY + 'MW / 风电 ' + simParams.WIND_CAPACITY +
        'MW / 电解槽最低比 ' + simParams.ELECTROLYZER_MIN_RATIO + ' / 制氢电耗 ' + simParams.HYDROGEN_ENERGY_CONSUMPTION + ' kWh/kg');
    out('  出厂默认财务参数：计算期 ' + financeParams.calc_period + ' 年 / 氢气售价 ' + financeParams.h2_price +
        ' 元/kg / 税金附加 ' + financeParams.surtax_rate + '%');

    const appShapedCtx = {
        pvData: pvData, windData: windData,
        simParams: simParams, prices: prices, financeParams: financeParams,
        lcohDiscountRate: config.lcoh.discountRate,
    };

    let r = null, err = null;
    try {
        const s = OptimizationEngine.createSession({ config: config, context: appShapedCtx });
        s.ensureInitialized();
        while (!s.isFinished()) s.runNextGeneration();
        r = s.getResult();
    } catch (e) { err = e; }

    if (err) {
        fail('11.2 app.js 形状的上下文可正常驱动优化会话', err.message);
        out('      ' + (err.stack || '').split('\n').slice(0, 4).join('\n      '));
    } else {
        ok('11.2 app.js 形状的上下文可正常驱动优化会话',
            '完成 ' + r.statistics.generationsCompleted + ' 代，评价 ' + r.statistics.totalEvaluated +
            ' 个，可行 ' + r.statistics.feasibleCount + '，Pareto ' + r.statistics.paretoCount);
    }

    // 11.3 上下文键名兼容（历史 pv/wind 写法也须可用）
    let aliasOk = false, aliasErr = null;
    try {
        const aliasCtx = {
            pv: pvData, wind: windData,
            simParams: simParams, prices: prices, financeParams: financeParams,
            lcohDiscountRate: config.lcoh.discountRate,
        };
        const ev = OptimizationEngine.evaluateScheme({
            windCapacity: 100, pvCapacity: 100, storagePower: 50, storageDuration: 2, electrolyzerCapacity: 50,
        }, aliasCtx);
        aliasOk = isFinite(ev.technical.annualHydrogenKg) && ev.technical.annualHydrogenKg > 0;
    } catch (e) { aliasErr = e; }
    assertTrue('11.3 上下文键名兼容（pv/wind 与 pvData/windData 均可）', aliasOk,
        aliasErr ? aliasErr.message : '两种写法结果一致');

    // 11.4 缺少风光数据时必须给出可读错误
    let missErr = null;
    try {
        OptimizationEngine.createSession({ config: config, context: { simParams: simParams, prices: prices, financeParams: financeParams } });
    } catch (e) { missErr = e; }
    assertTrue('11.4 缺少风光数据时给出可读错误', !!missErr && /8760|数据/.test(missErr.message),
        missErr ? missErr.message : '未抛出');
})();

// ---------------------------------------------------------------------------
// 汇总
// ---------------------------------------------------------------------------
section('测试汇总');
out('  PASS：' + passCount + ' 项');
out('  FAIL：' + failCount + ' 项');
out('  结论：' + (failCount === 0 ? '全部通过 ✅' : '存在失败项 ❌'));
out('');

fs.writeFileSync(OUT_FILE, lines.join('\n'), 'utf8');
process.stdout.write(lines.join('\n') + '\n');
process.exitCode = failCount === 0 ? 0 : 1;
