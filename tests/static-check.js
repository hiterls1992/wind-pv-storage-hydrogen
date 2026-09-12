/**
 * ============================================================================
 * V2.1 静态一致性检查（Node.js）
 * ============================================================================
 * 纯前端项目的典型故障是「JS 引用的 DOM id 在 HTML 中不存在」，浏览器不会报错，
 * 只会静默失效。本脚本在提交前做四类静态校验：
 *   1. 全部 js 文件语法检查（node --check 等价）
 *   2. HTML 引用的脚本文件是否真实存在
 *   3. JS 中 getElementById / optEl 引用的静态 id 是否都在 HTML 中定义
 *   4. data-tab / data-subtab / data-ftab 与对应面板 id 是否一一对应
 *
 * 运行：node tests/static-check.js
 */

'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.resolve(__dirname, '..');
const HTML = path.join(ROOT, 'index.html');

const lines = [];
let pass = 0, fail = 0;
function ok(n, d) { pass++; lines.push('  [PASS] ' + n + (d ? '  —— ' + d : '')); }
function bad(n, d) { fail++; lines.push('  [FAIL] ' + n + (d ? '  —— ' + d : '')); }
function section(t) { lines.push(''); lines.push('='.repeat(78)); lines.push(t); lines.push('='.repeat(78)); }

const html = fs.readFileSync(HTML, 'utf8');

// ---------------------------------------------------------------------------
section('1. JavaScript 语法检查');
const JS_FILES = [
    'js/utils.js', 'js/parameter-manager.js', 'js/result-data-store.js', 'js/excel-io.js',
    'js/simulation-engine.js', 'js/data-summary.js', 'js/estimate.js', 'js/finance-engine.js',
    'js/chart-module.js', 'js/optimization-engine.js', 'js/optimization-worker.js', 'js/app.js',
];
for (const f of JS_FILES) {
    const p = path.join(ROOT, f);
    if (!fs.existsSync(p)) { bad('文件存在', f); continue; }
    const src = fs.readFileSync(p, 'utf8');
    try {
        // 仅解析，不执行
        new vm.Script(src, { filename: f });
        ok('语法正确', f + '（' + src.split('\n').length + ' 行）');
    } catch (e) {
        bad('语法错误', f + ' → ' + e.message);
    }
}

// ---------------------------------------------------------------------------
section('2. HTML 脚本引用检查');
const scriptSrcs = [];
const reScript = /<script\s+src="([^"]+)"/g;
let m;
while ((m = reScript.exec(html)) !== null) scriptSrcs.push(m[1]);
lines.push('  HTML 引用脚本 ' + scriptSrcs.length + ' 个：');
for (const s of scriptSrcs) {
    const exists = fs.existsSync(path.join(ROOT, s));
    if (exists) ok('脚本文件存在', s);
    else bad('脚本文件缺失', s);
    lines.push('    - ' + s);
}

const cssRefs = [];
const reCss = /<link[^>]+href="([^"]+\.css)"/g;
while ((m = reCss.exec(html)) !== null) cssRefs.push(m[1]);
for (const c of cssRefs) {
    if (fs.existsSync(path.join(ROOT, c))) ok('样式文件存在', c);
    else bad('样式文件缺失', c);
}

// ---------------------------------------------------------------------------
section('3. DOM id 引用一致性（JS → HTML）');

const htmlIds = new Set();
const reId = /\bid="([^"]+)"/g;
while ((m = reId.exec(html)) !== null) htmlIds.add(m[1]);
lines.push('  HTML 中定义 id 共 ' + htmlIds.size + ' 个');

const jsRefs = new Map();   // id -> [文件]
for (const f of JS_FILES) {
    const src = fs.readFileSync(path.join(ROOT, f), 'utf8');
    const patterns = [
        /getElementById\(\s*'([^']+)'\s*\)/g,
        /optEl\(\s*'([^']+)'\s*\)/g,
    ];
    for (const re of patterns) {
        let mm;
        while ((mm = re.exec(src)) !== null) {
            const id = mm[1];
            if (id.indexOf('${') >= 0 || id.indexOf('`') >= 0) continue;   // 动态 id 跳过
            if (!jsRefs.has(id)) jsRefs.set(id, []);
            jsRefs.get(id).push(f);
        }
    }
}
lines.push('  JS 中静态引用 id 共 ' + jsRefs.size + ' 个');

const missing = [];
for (const [id, files] of jsRefs) {
    // 允许由 JS 动态生成的行内元素
    if (!htmlIds.has(id)) missing.push(id + '（' + Array.from(new Set(files)).join(', ') + '）');
}
if (missing.length === 0) {
    ok('全部静态 id 引用均可在 HTML 中找到定义');
} else {
    bad('存在 HTML 中不存在的 id 引用', missing.join(' | '));
}

// ---------------------------------------------------------------------------
section('4. 标签页 / 子标签页 面板对应关系');

function checkPairing(attr, idPrefix) {
    const vals = [];
    const re = new RegExp('data-' + attr + '="([^"]+)"', 'g');
    let mm;
    while ((mm = re.exec(html)) !== null) vals.push(mm[1]);
    const missingPanels = vals.filter(v => !htmlIds.has(idPrefix + v));
    if (vals.length === 0) { bad('未找到 data-' + attr + ' 按钮'); return; }
    if (missingPanels.length === 0) {
        ok('data-' + attr + ' ↔ #' + idPrefix + '* 全部匹配', vals.join(', '));
    } else {
        bad('存在缺少对应面板的按钮', missingPanels.map(v => v + '→#' + idPrefix + v).join(', '));
    }
}

checkPairing('tab', 'tab-');
checkPairing('subtab', 'subtab-');
checkPairing('ftab', 'ftab-');

// 反查：面板是否都有入口按钮
const panelIds = [];
const rePanel = /\bid="(tab|subtab|ftab)-([^"]+)"/g;
while ((m = rePanel.exec(html)) !== null) panelIds.push({ full: m[1] + '-' + m[2], attr: m[1], key: m[2] });
const orphans = panelIds.filter(p => html.indexOf('data-' + p.attr + '="' + p.key + '"') < 0);
if (orphans.length === 0) ok('全部面板均有对应的导航按钮');
else bad('存在无入口按钮的面板', orphans.map(o => o.full).join(', '));

// ---------------------------------------------------------------------------
section('5. app.js 内部函数引用完整性（优化模块）');
(function () {
    const src = fs.readFileSync(path.join(ROOT, 'js', 'app.js'), 'utf8');
    // 关键函数定义与调用
    const mustExist = [
        'function initOptimization', 'function startOptimization', 'function startWithWorker',
        'function startWithMainThread', 'function stopOptimization', 'function updateOptProgress',
        'function onOptimizationFinished', 'function renderOptimizationResult', 'function renderRecommendArea',
        'function renderParetoTable', 'function renderParetoChart', 'function renderConvergenceCharts',
        'function showSchemeDetail', 'function applySchemeToSimulation', 'function exportOptimizationExcel',
        'function getEstimatePrices', 'function collectOptimizationContext',
    ];
    const missingFn = mustExist.filter(k => src.indexOf(k) < 0);
    if (missingFn.length === 0) ok('优化模块关键函数全部存在', mustExist.length + ' 个');
    else bad('缺少关键函数', missingFn.join(', '));

    // 导出函数
    const excel = fs.readFileSync(path.join(ROOT, 'js', 'excel-io.js'), 'utf8');
    if (excel.indexOf('exportOptimizationResults') >= 0) ok('ExcelIO.exportOptimizationResults 已实现');
    else bad('ExcelIO.exportOptimizationResults 缺失');

    const chart = fs.readFileSync(path.join(ROOT, 'js', 'chart-module.js'), 'utf8');
    if (chart.indexOf('renderParetoScatter') >= 0 && chart.indexOf('renderConvergenceChart') >= 0) {
        ok('ChartModule 优化图表函数已实现');
    } else {
        bad('ChartModule 优化图表函数缺失');
    }
})();

// ---------------------------------------------------------------------------
section('6. 跨层字段契约（app.js ↔ 引擎 ↔ Worker）');
(function () {
    const app = fs.readFileSync(path.join(ROOT, 'js', 'app.js'), 'utf8');
    const eng = fs.readFileSync(path.join(ROOT, 'js', 'optimization-engine.js'), 'utf8');
    const wrk = fs.readFileSync(path.join(ROOT, 'js', 'optimization-worker.js'), 'utf8');

    // 该契约曾因 app.js 传 pv/wind 而引擎约定 pvData/windData 导致运行期失败
    // V2.2 起运行参数字段名统一为 simulationConfig，容量的承载者也从上下文改为 scheme
    const mustInApp = [
        ['pvData: pvData', 'collectOptimizationContext 返回 pvData'],
        ['windData: windData', 'collectOptimizationContext 返回 windData'],
        ['simulationConfig: getSimulationConfig()', 'collectOptimizationContext 返回 simulationConfig'],
        ['pvData: ctx.pvData', 'Worker 消息体携带 pvData'],
        ['windData: ctx.windData', 'Worker 消息体携带 windData'],
        ['simulationConfig: ctx.simulationConfig', 'Worker 消息体携带 simulationConfig'],
    ];
    const missApp = mustInApp.filter(([k]) => app.indexOf(k) < 0).map(([, d]) => d);
    if (missApp.length === 0) ok('app.js 与引擎的上下文字段名一致（pvData / windData / simulationConfig）');
    else bad('app.js 上下文字段名与引擎不一致', missApp.join('、'));

    if (app.indexOf('pv: ctx.pv') < 0) ok('app.js 未残留下线字段名 pv: ctx.pv');
    else bad('app.js 仍存在旧字段名 pv: ctx.pv');

    if (app.indexOf('simParams: getSimulationParams()') < 0) ok('app.js 不再通过 simParams 夹带运行参数');
    else bad('app.js 仍通过 simParams 传递上下文');

    if (wrk.indexOf('raw.pvData') >= 0 && wrk.indexOf('raw.windData') >= 0) {
        ok('optimization-worker.js 正确读取 pvData / windData');
    } else {
        bad('optimization-worker.js 未读取 pvData / windData');
    }

    if (wrk.indexOf('simulationConfig:') >= 0) ok('optimization-worker.js 使用 simulationConfig 字段名');
    else bad('optimization-worker.js 未使用 simulationConfig');

    if (eng.indexOf('function normalizeContext') >= 0) ok('引擎提供上下文键名兼容（normalizeContext）');
    else bad('引擎缺少上下文键名兼容处理');

    if (eng.indexOf('function toSimulationConfig') >= 0) ok('引擎提供运行参数字段名兼容层（toSimulationConfig）');
    else bad('引擎缺少运行参数兼容层');
})();

// ---------------------------------------------------------------------------
section('7. 参数唯一来源检查（V2.2 任务书 §12 / §36 / §50）');
(function () {
    const read = f => fs.readFileSync(path.join(ROOT, f), 'utf8');
    const app = read('js/app.js');
    const eng = read('js/optimization-engine.js');
    const sim = read('js/simulation-engine.js');
    const pm = read('js/parameter-manager.js');

    // 7.1 容量参数被注入运行参数对象的写法必须消失（§36）
    if (eng.indexOf('PV_CAPACITY: scheme.pvCapacity') < 0 && eng.indexOf('WIND_CAPACITY: scheme.windCapacity') < 0) {
        ok('优化引擎不再把容量 Object.assign 进运行参数对象');
    } else {
        bad('优化引擎仍存在容量注入运行参数对象的写法');
    }

    // 7.2 app.js 不得再定义 getSimulationParams（容量与运行参数混装）
    if (app.indexOf('function getSimulationParams') < 0 && app.indexOf('getSimulationParams(') < 0) {
        ok('app.js 已移除 getSimulationParams（容量与运行参数混装）');
    } else {
        bad('app.js 仍存在 getSimulationParams');
    }

    // 7.3 旧 DOM id 不得再被任何 JS 引用（已从 HTML 删除）
    const REMOVED_IDS = [
        'pvCapacity', 'windCapacity',
        'storagePowerMin', 'storagePowerMax', 'storagePowerStep',
        'storageDurationMin', 'storageDurationMax', 'storageDurationStep',
        'electrolyzerMin', 'electrolyzerMax', 'electrolyzerStep',
        'baseWind', 'basePv', 'baseStoragePower', 'baseStorageDuration', 'baseElectrolyzer',
    ];
    const staleRefs = [];
    for (const f of JS_FILES) {
        const src = read(f);
        for (const id of REMOVED_IDS) {
            if (src.indexOf("getElementById('" + id + "')") >= 0 || src.indexOf('optEl(\'' + id + '\')') >= 0) {
                staleRefs.push(f + " → '" + id + "'");
            }
        }
    }
    if (staleRefs.length === 0) ok('已删除的旧容量 DOM id 无任何残留引用', REMOVED_IDS.length + ' 个旧 id');
    else bad('仍引用已删除的旧 id', staleRefs.join(' | '));

    // 7.4 新命名契约必须同时存在于 HTML 与 ParameterManager
    const needSchemeIds = ['currentWindCapacity', 'currentPvCapacity', 'currentStoragePower',
        'currentStorageDuration', 'currentStorageEnergy', 'currentElectrolyzerCapacity'];
    const missHtml = needSchemeIds.filter(id => htmlIds.has(id) === false);
    if (missHtml.length === 0) ok('HTML 定义了全部「当前方案」输入 id', needSchemeIds.length + ' 个');
    else bad('HTML 缺少「当前方案」输入 id', missHtml.join(', '));

    const missPm = needSchemeIds.filter(id => pm.indexOf("'" + id + "'") < 0);
    if (missPm.length === 0) ok('ParameterManager 与 HTML 的「当前方案」id 一一对应');
    else bad('ParameterManager 中的 id 与 HTML 不一致', missPm.join(', '));

    // 7.5 储能容量必须是只读派生字段（§25）
    if (/id="currentStorageEnergy"[^>]*readonly/.test(html)) {
        ok('储能容量为只读派生字段（不可直接输入）');
    } else {
        bad('储能容量未设置为只读派生字段');
    }
    if (pm.indexOf('storageEnergy = roundTo(out.storagePower * out.storageDuration') >= 0) {
        ok('储能容量 = 储能功率 × 储能时长（派生计算）');
    } else {
        bad('ParameterManager 未实现储能容量派生计算');
    }

    // 7.6 ParameterManager 关键接口齐全（§23）
    const needApi = ['getCurrentScheme', 'setCurrentScheme', 'getSimulationConfig',
        'getOptimizationConfig', 'validateScheme', 'validateSimulationConfig',
        'validateOptimizationConfig', 'syncSchemeToUI', 'syncSchemeFromUI', 'schemeKey'];
    const missApi = needApi.filter(k => pm.indexOf(k + ':') < 0 && pm.indexOf(k + ' =') < 0 && pm.indexOf('function ' + k) < 0);
    if (missApi.length === 0) ok('ParameterManager 接口齐全', needApi.join(', '));
    else bad('ParameterManager 缺少接口', missApi.join(', '));

    // 7.7 仿真引擎新签名
    if (sim.indexOf('function runSingleSimulation(pvData, windData, scheme, simulationConfig') >= 0) {
        ok('runSingleSimulation 已改为 (pvData, windData, scheme, simulationConfig)');
    } else {
        bad('runSingleSimulation 签名未更新');
    }
    if (sim.indexOf('resolveSimulationArgs') >= 0 && sim.indexOf('TODO V2.3 REMOVE LEGACY') >= 0) {
        ok('保留旧签名兼容层并已标注 TODO V2.3 REMOVE LEGACY（§43）');
    } else {
        bad('缺少旧签名兼容层或未标注 TODO');
    }

    // 7.8 结果必须自带完整方案（§48）
    if (sim.indexOf('scheme: {') >= 0 && sim.indexOf('storageEnergy: storageEnergy') >= 0) {
        ok('仿真结果自带完整 scheme（含 storageEnergy）');
    } else {
        bad('仿真结果未携带完整 scheme');
    }

    // 7.9 app.js 必须通过 ParameterManager 读写参数，不得散读 DOM（§21）
    const hasSync = app.indexOf('ParameterManager.syncSchemeFromUI()') >= 0
        && app.indexOf('ParameterManager.syncSchemeToUI(') >= 0
        && app.indexOf('ParameterManager.setCurrentScheme(') >= 0;
    if (hasSync) ok('app.js 通过 ParameterManager 统一读写当前方案');
    else bad('app.js 未通过 ParameterManager 统一读写当前方案');

    // 7.10 baseline 唯一来源：不得再出现独立的 base* 输入框
    if (html.indexOf('id="baseWind"') < 0 && html.indexOf('id="basePv"') < 0) {
        ok('优化页已移除独立的基准方案输入框（§16 / §18）');
    } else {
        bad('优化页仍存在独立的基准方案输入框');
    }
})();

// ---------------------------------------------------------------------------
section('8. 数据流与性能契约（V2.3 任务书 §5 / §6 / §11 / §20 / §21）');
(function () {
    const read = f => fs.readFileSync(path.join(ROOT, f), 'utf8');
    const app = read('js/app.js');
    const chart = read('js/chart-module.js');
    const store = read('js/result-data-store.js');
    const worker = read('js/optimization-worker.js');

    // 8.1 显示路径禁止再物化 8760 个行对象（§5）
    //     app.js / chart-module.js 的**渲染与表格路径**不得调用 parseResults
    //     （允许出现在：chart-module 内的废弃标记注释与导出兜底）
    const bannedInApp = /ChartModule\.parseResults/.test(app);
    if (!bannedInApp) ok('app.js 显示/表格/图表路径已全部改用 ResultDataStore（不再物化 8760 行对象）');
    else bad('app.js 仍调用 ChartModule.parseResults（V2.3 禁止）');

    // 8.2 ResultDataStore 基础接口齐全（§4）
    const needApi = ['COLS', 'COL_COUNT', 'create', 'getValue', 'getHourRow', 'getRange',
        'getColumn', 'getLength', 'getAnnualSummary', 'getMonthlySummary'];
    const missApi = needApi.filter(k => store.indexOf(k + ':') < 0 && store.indexOf('function ' + k) < 0);
    if (missApi.length === 0) ok('ResultDataStore 接口齐全（§4）', needApi.join(', '));
    else bad('ResultDataStore 缺少接口', missApi.join(', '));

    // 8.3 显示降采样不得回流到计算（§8）
    if (store.indexOf('本函数输出**只允许用于绘图**') >= 0 && store.indexOf('禁止用于任何指标 / 约束计算') >= 0) {
        ok('降采样接口已明确标注「仅用于绘图」');
    } else {
        bad('降采样接口缺少「禁止用于指标计算」的警示标注');
    }

    // 8.4 逐时表必须虚拟滚动，禁止一次性生成 8760 行（§6）
    if (app.indexOf('function renderSimTableRows') >= 0 && app.indexOf('BUFFER_ROWS') >= 0
        && app.indexOf('vt-spacer') >= 0) {
        ok('逐时表已实现虚拟滚动（可视区间 + 缓冲 + 占位行）');
    } else {
        bad('逐时表未实现虚拟滚动');
    }
    if (/for \(let h = 0; h < .*8760/.test(app) === false && app.indexOf('vt-spacer') >= 0) {
        ok('逐时表不再一次性拼接全部行');
    }
    if (/ROW_HEIGHT|rowHeight/.test(app)) ok('虚拟滚动使用固定行高计算可视区间');

    // 8.5 图表实例复用，切换类型不再 dispose（§10）
    if (chart.indexOf('echarts.getInstanceByDom') >= 0 && chart.indexOf('existing.clear()') >= 0) {
        ok('ECharts 实例按容器复用并 clear() 旧 option（不再 dispose 重建）');
    } else {
        bad('ECharts 实例未复用');
    }
    if (app.indexOf('AppState.currentChart.dispose()') < 0) {
        ok('updateChart 不再销毁图表实例');
    } else {
        bad('updateChart 仍会 dispose 图表实例');
    }

    // 8.6 优化阶段不保存 8760 原始结果（§11）
    if (store.indexOf('OptimizationResultCache') >= 0 && store.indexOf('strip(result)') >= 0
        && worker.indexOf('storesHourlyResults: false') >= 0) {
        ok('优化缓存只保存指标（OptimizationResultCache.strip + Worker 内存审计字段）');
    } else {
        bad('优化结果内存约束未落实');
    }

    // 8.7 月度统计只算一次（§18 / §19）
    if (store.indexOf('_monthly') >= 0 && store.indexOf('getMonthlyColumn') >= 0
        && chart.indexOf('_monthlyTotals') >= 0) {
        ok('月度摘要统一缓存，图表不再每次渲染遍历 8760 小时');
    } else {
        bad('月度统计未走统一缓存');
    }

    // 8.8 日志限长（§20）
    if (app.indexOf('MAX_LOG_LINES') >= 0) ok('日志已限长（MAX_LOG_LINES）');
    else bad('日志未限长');

    // 8.9 禁止循环内 innerHTML +=（§21）
    const loopAppend = [];
    for (const f of JS_FILES) {
        const src = read(f);
        const re = /while\s*\([^)]*\)\s*\{[^}]*innerHTML\s*\+=|for\s*\([^)]*\)\s*\{[^}]*innerHTML\s*\+=/gs;
        if (re.test(src)) loopAppend.push(f);
    }
    if (loopAppend.length === 0) ok('无「循环内 innerHTML +=」的写法');
    else bad('仍存在循环内 innerHTML +=', loopAppend.join(', '));

    // 8.10 Worker 优先 + 诊断（§13 / §14 / §17）
    if (worker.indexOf('performanceStats') >= 0 && worker.indexOf('cacheHitRate') >= 0) {
        ok('Worker 输出性能诊断（performanceStats）');
    } else {
        bad('Worker 缺少性能诊断');
    }
    if (app.indexOf('worker.onerror') >= 0 && app.indexOf('diagnose') >= 0) {
        ok('Worker 失败原因已分类上报（不静默回退）');
    } else {
        bad('Worker 失败缺少诊断');
    }

    // 8.11 Worker 必须加载结果数据层（§13：全部计算在 Worker 内完成）
    if (worker.indexOf("'result-data-store.js'") >= 0) ok('Worker 已加载 result-data-store.js');
    else bad('Worker 未加载 result-data-store.js');

    // 8.12 主线程时间片（§16）
    if (/CHUNK_MS = 1[0-5]\b/.test(app)) ok('主线程 Fallback 时间片已降至 10~15ms');
    else bad('主线程 Fallback 时间片未调整（应为 10~15ms）');
})();

// ---------------------------------------------------------------------------
section('9. 检查汇总');
lines.push('  PASS：' + pass + ' 项');
lines.push('  FAIL：' + fail + ' 项');
lines.push('  结论：' + (fail === 0 ? '全部通过 ✅' : '存在失败项 ❌'));
lines.push('');

const OUT = path.join(__dirname, 'static-check-report.txt');
fs.writeFileSync(OUT, lines.join('\n'), 'utf8');
process.stdout.write(lines.join('\n') + '\n');
process.exitCode = fail === 0 ? 0 : 1;
