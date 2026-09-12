/**
 * ============================================================================
 * 统一参数管理器 ParameterManager  ——  V2.2 参数体系重构核心
 * ============================================================================
 *
 * 设计原则（V2.2 任务书 §2 / §24）：
 *     一个参数，一个定义，一个数据源，一个传递路径。
 *
 * ---------------------------------------------------------------------------
 * 三类参数严格分离（§5）
 * ---------------------------------------------------------------------------
 *   ① 方案参数 Scheme        ——「建多大」：5 个容量变量 + 派生的储能容量
 *   ② 系统运行参数 Simulation ——「设备怎么运行」：效率 / 比例 / 电耗
 *   ③ 优化参数 Optimization   ——「怎么搜索」：范围 / 约束 / NSGA-II / 权重
 *
 * 唯一真值：
 *   ParameterManager.getCurrentScheme()        ← 电量计算 / 方案汇总 / Baseline / 方案详情 / 方案复核
 *   ParameterManager.getSimulationConfig()
 *   ParameterManager.getOptimizationConfig()
 *
 * 禁止任何模块自行保存一套容量参数（§4 / §21 / §50）。
 *
 * ---------------------------------------------------------------------------
 * 单位约定（§46）——功率与电量不得混淆
 * ---------------------------------------------------------------------------
 *   风电容量 / 光伏容量 / 储能功率 / 电解槽容量   MW
 *   储能时长                                      h
 *   储能容量 = 储能功率 × 储能时长                 MWh（派生量，只读，禁止输入）
 *   电量                                          MWh
 *   制氢量                                        kg（内部）/ 万吨（展示）
 *   制氢电耗                                      kWh/kg
 *
 * ---------------------------------------------------------------------------
 * 运行环境
 * ---------------------------------------------------------------------------
 * 本文件同时被浏览器主线程（index.html）与 Web Worker（optimization-worker.js
 * 内 importScripts）加载，因此**所有 DOM 访问必须写在函数体内**，
 * 模块加载期不得触碰 document / window。
 */

(function () {
    'use strict';

    // =========================================================================
    // 常量：参数键名（全项目唯一权威定义）
    // =========================================================================

    /** 方案参数：5 个容量变量（顺序即 NSGA-II 决策变量顺序，不可调整） */
    const SCHEME_KEYS = [
        'windCapacity',
        'pvCapacity',
        'storagePower',
        'storageDuration',
        'electrolyzerCapacity',
    ];

    /** 方案参数中文名（用于状态栏 / 表格 / 日志） */
    const SCHEME_LABELS = {
        windCapacity: '风电容量',
        pvCapacity: '光伏容量',
        storagePower: '储能功率',
        storageDuration: '储能时长',
        electrolyzerCapacity: '电解槽容量',
    };

    /** 方案参数单位 */
    const SCHEME_UNITS = {
        windCapacity: 'MW',
        pvCapacity: 'MW',
        storagePower: 'MW',
        storageDuration: 'h',
        electrolyzerCapacity: 'MW',
    };

    /** 系统运行参数键名 */
    const SIM_KEYS = [
        'electrolyzerMinRatio',
        'maxExportHourly',
        'maxExportTotal',
        'maxImportRatio',
        'chargeEfficiency',
        'dischargeEfficiency',
        'hydrogenConsumption',
    ];

    /**
     * 出厂默认方案（与 index.html 中「当前方案」输入框的 value 必须一致）
     * 仅作为「输入框缺失 / 非法」时的兜底，不作为运行时真值。
     */
    const DEFAULT_SCHEME = {
        windCapacity: 200,
        pvCapacity: 360,
        storagePower: 100,
        storageDuration: 2,
        electrolyzerCapacity: 160,
    };

    /** 出厂默认运行参数（与 index.html 中「运行参数」输入框的 value 必须一致） */
    const DEFAULT_SIMULATION_CONFIG = {
        electrolyzerMinRatio: 0.30,
        maxExportHourly: 0.0,
        maxExportTotal: 0.0,
        maxImportRatio: 0.15,
        chargeEfficiency: 0.92,
        dischargeEfficiency: 0.92,
        hydrogenConsumption: 55,
    };

    /**
     * 出厂默认优化参数（与 index.html 优化页各输入框 value 必须一致）
     * 运行时真值由 app.js 每次「开始优化」前从 DOM 读取后经 setOptimizationConfig() 写入。
     */
    const DEFAULT_OPTIMIZATION_CONFIG = {
        variables: {
            windCapacity:         { min: 50, max: 300, step: 25 },
            pvCapacity:           { min: 50, max: 500, step: 25 },
            storagePower:         { min: 0,  max: 200, step: 25 },
            storageDuration:      { min: 0,  max: 4,   step: 1 },
            electrolyzerCapacity: { min: 25, max: 200, step: 25 },
        },
        constraints: {},
        nsga2: {
            populationSize: 60,
            generations: 100,
            crossoverProbability: 0.9,
            mutationProbability: 0.1,
            randomSeed: 20260912,
            earlyStopping: true,
            patience: 15,
        },
        objectives: {},
        recommendationWeights: { firr: 0.40, lcoh: 0.35, curtailmentRate: 0.25 },
        lcoh: { discountRate: 5.0 },
        baseline: null,
    };

    /** 方案参数 → DOM id（唯一映射，禁止在别处再写一遍 id 字面量） */
    const SCHEME_DOM = {
        windCapacity: 'currentWindCapacity',
        pvCapacity: 'currentPvCapacity',
        storagePower: 'currentStoragePower',
        storageDuration: 'currentStorageDuration',
        electrolyzerCapacity: 'currentElectrolyzerCapacity',
    };

    /** 储能容量（派生量）只读显示字段 id */
    const STORAGE_ENERGY_DOM = 'currentStorageEnergy';

    /** 运行参数 → DOM id */
    const SIM_DOM = {
        electrolyzerMinRatio: 'electrolyzerMinRatio',
        maxExportHourly: 'maxExportHourly',
        maxExportTotal: 'maxExportTotal',
        maxImportRatio: 'maxImportRatio',
        chargeEfficiency: 'chargeEfficiency',
        dischargeEfficiency: 'dischargeEfficiency',
        hydrogenConsumption: 'hydrogenConsumption',
    };

    // =========================================================================
    // 内部状态（唯一真值所在）
    // =========================================================================

    const state = {
        currentScheme: cloneScheme(DEFAULT_SCHEME),
        simulationConfig: Object.assign({}, DEFAULT_SIMULATION_CONFIG),
        optimizationConfig: deepClone(DEFAULT_OPTIMIZATION_CONFIG),
    };

    // =========================================================================
    // 工具
    // =========================================================================

    function deepClone(v) {
        return JSON.parse(JSON.stringify(v));
    }

    function num(v, fallback) {
        const n = Number(v);
        return isFinite(n) ? n : fallback;
    }

    function roundTo(v, digits) {
        const p = Math.pow(10, digits);
        return Math.round(v * p) / p;
    }

    // =========================================================================
    // 方案参数 Scheme
    // =========================================================================

    /**
     * 规范化方案对象：补齐 5 个键、做数值安全化、自动派生 storageEnergy。
     * 永远返回**新对象**，不修改入参，避免调用方共享引用。
     *
     * @param {Object} scheme 任意形态的方案对象
     * @param {Object} [fallback] 键缺失时的兜底值（默认用出厂默认方案）
     * @returns {Object} {5 个容量参数 + storageEnergy}
     */
    function normalizeScheme(scheme, fallback) {
        const fb = fallback || DEFAULT_SCHEME;
        const src = scheme || {};
        const out = {};
        for (const k of SCHEME_KEYS) {
            out[k] = num(src[k], num(fb[k], DEFAULT_SCHEME[k]));
        }
        // 储能容量永远是派生量（§25），禁止作为独立输入
        out.storageEnergy = roundTo(out.storagePower * out.storageDuration, 6);
        return out;
    }

    /** 内部：只取 5 个容量键（不含 storageEnergy），用于缓存键与落库 */
    function cloneScheme(scheme) {
        return normalizeScheme(scheme, DEFAULT_SCHEME);
    }

    /** 返回当前方案（新对象，外部修改不影响唯一真值） */
    function getCurrentScheme() {
        return Object.assign({}, state.currentScheme);
    }

    /**
     * 写入当前方案（唯一真值入口）。
     * @param {Object} scheme
     * @returns {{ok:boolean, scheme:Object, message:string, warnings:string[]}}
     */
    function setCurrentScheme(scheme) {
        const normalized = normalizeScheme(scheme, state.currentScheme);
        const v = validateScheme(normalized);
        if (v.ok) state.currentScheme = normalized;
        return {
            ok: v.ok,
            scheme: Object.assign({}, state.currentScheme),
            message: v.message,
            warnings: v.warnings,
        };
    }

    /** 当前方案的中文单行描述 */
    function schemeLabel(scheme) {
        const s = normalizeScheme(scheme, state.currentScheme);
        return '风电 ' + s.windCapacity + 'MW / 光伏 ' + s.pvCapacity + 'MW / ' +
            '储能 ' + s.storagePower + 'MW×' + s.storageDuration + 'h（' + s.storageEnergy + 'MWh）/ ' +
            '电解槽 ' + s.electrolyzerCapacity + 'MW';
    }

    /**
     * 方案唯一标识（§49）——全项目统一用于：
     * 缓存 / 结果索引 / 方案选择 / Pareto方案 / Baseline / Excel导出
     *
     * 例：W200|PV360|B100|H2|EL160
     */
    function schemeKey(scheme) {
        const s = normalizeScheme(scheme, state.currentScheme);
        return 'W' + fmt(s.windCapacity) +
            '|PV' + fmt(s.pvCapacity) +
            '|B' + fmt(s.storagePower) +
            '|H' + fmt(s.storageDuration) +
            '|EL' + fmt(s.electrolyzerCapacity);
    }

    function fmt(v) {
        return String(roundTo(num(v, 0), 6));
    }

    /**
     * 方案参数校验（§44）：5 个容量均必须为有限非负数。
     * @returns {{ok:boolean, message:string, warnings:string[]}}
     */
    function validateScheme(scheme) {
        const s = scheme || {};
        const bad = [];
        for (const k of SCHEME_KEYS) {
            const v = Number(s[k]);
            if (!isFinite(v)) bad.push(SCHEME_LABELS[k] + ' 不是有效数值');
            else if (v < 0) bad.push(SCHEME_LABELS[k] + ' 不能为负（当前 ' + v + '）');
        }
        const warnings = [];
        if (isFinite(Number(s.storageDuration)) && isFinite(Number(s.storagePower))) {
            if (Number(s.storageDuration) > 0 && Number(s.storagePower) === 0) {
                warnings.push('储能功率为 0，储能容量（' + Number(s.storageDuration) + 'h）无效，实际不配置储能');
            }
        }
        return {
            ok: bad.length === 0,
            message: bad.join('；'),
            warnings: warnings,
        };
    }

    // =========================================================================
    // 系统运行参数 SimulationConfig
    // =========================================================================

    function getSimulationConfig() {
        return Object.assign({}, state.simulationConfig);
    }

    function setSimulationConfig(config) {
        const merged = Object.assign({}, state.simulationConfig);
        for (const k of SIM_KEYS) {
            if (config && config[k] !== undefined) merged[k] = num(config[k], merged[k]);
        }
        state.simulationConfig = merged;
        return Object.assign({}, merged);
    }

    /**
     * 运行参数校验：比例类限 0~1，效率类限 (0,1]，制氢电耗 > 0。
     * 只做合理性提示，不阻断计算（保持与 V1.0 一致的宽容行为）。
     */
    function validateSimulationConfig(config) {
        const c = config || state.simulationConfig;
        const bad = [];
        const ratioKeys = ['electrolyzerMinRatio', 'maxExportHourly', 'maxExportTotal', 'maxImportRatio'];
        const effKeys = ['chargeEfficiency', 'dischargeEfficiency'];
        for (const k of ratioKeys) {
            const v = Number(c[k]);
            if (!isFinite(v) || v < 0 || v > 1) bad.push(k + ' 应在 0~1 之间（当前 ' + c[k] + '）');
        }
        for (const k of effKeys) {
            const v = Number(c[k]);
            if (!isFinite(v) || v <= 0 || v > 1) bad.push(k + ' 应在 (0,1] 之间（当前 ' + c[k] + '）');
        }
        const h2 = Number(c.hydrogenConsumption);
        if (!isFinite(h2) || h2 <= 0) bad.push('制氢电耗必须大于 0（当前 ' + c.hydrogenConsumption + '）');
        return { ok: bad.length === 0, message: bad.join('；'), warnings: [] };
    }

    // =========================================================================
    // 优化参数 OptimizationConfig
    // =========================================================================

    function getOptimizationConfig() {
        return deepClone(state.optimizationConfig);
    }

    function setOptimizationConfig(config) {
        if (config && typeof config === 'object') {
            state.optimizationConfig = deepClone(config);
        }
        return getOptimizationConfig();
    }

    /**
     * 优化参数校验（§45）：min ≤ max，step > 0；
     * 并检查「当前基准方案是否落在搜索范围内」——**只提示，绝不自动修改**（§28）。
     */
    function validateOptimizationConfig(config) {
        const c = config || state.optimizationConfig;
        const bad = [];
        const variables = (c && c.variables) || {};
        for (const k of SCHEME_KEYS) {
            const cfg = variables[k];
            if (!cfg) continue;
            const min = Number(cfg.min), max = Number(cfg.max), step = Number(cfg.step);
            if (!isFinite(min) || !isFinite(max)) bad.push(SCHEME_LABELS[k] + ' 范围必须为有效数值');
            else if (min > max) bad.push(SCHEME_LABELS[k] + ' 最小值不能大于最大值');
            if (!isFinite(step) || step <= 0) bad.push(SCHEME_LABELS[k] + ' 步长必须大于 0');
        }
        return { ok: bad.length === 0, message: bad.join('；'), warnings: [] };
    }

    /**
     * 判断基准方案是否落在优化搜索范围内（§27）。
     * @returns {{inRange:boolean, items:Array<{key,label,unit,current,min,max,text}>}}
     */
    function checkBaselineInRange(scheme, optimizationConfig) {
        const s = normalizeScheme(scheme, state.currentScheme);
        const c = optimizationConfig || state.optimizationConfig;
        const variables = (c && c.variables) || {};
        const items = [];
        for (const k of SCHEME_KEYS) {
            const cfg = variables[k];
            if (!cfg) continue;
            const cur = Number(s[k]);
            const min = Number(cfg.min), max = Number(cfg.max);
            if (!isFinite(min) || !isFinite(max)) continue;
            const inRange = cur >= min && cur <= max;
            items.push({
                key: k,
                label: SCHEME_LABELS[k],
                unit: SCHEME_UNITS[k],
                current: cur,
                min: min,
                max: max,
                inRange: inRange,
                text: SCHEME_LABELS[k] + ' 当前 ' + cur + SCHEME_UNITS[k] + '，范围 ' + min + '~' + max + SCHEME_UNITS[k] +
                    (inRange ? ' ✓ 在范围内' : ' ⚠ 超出范围'),
            });
        }
        return { inRange: items.every(it => it.inRange), items: items };
    }

    /**
     * 求出「把基准方案包含进搜索范围」后的最小改动范围（§27 可选功能）。
     * 只做区间扩张，不改变 step；不修改任何现有配置。
     */
    function expandRangeToIncludeBaseline(scheme, optimizationConfig) {
        const s = normalizeScheme(scheme, state.currentScheme);
        const next = deepClone(optimizationConfig || state.optimizationConfig);
        const variables = next.variables || {};
        for (const k of SCHEME_KEYS) {
            const cfg = variables[k];
            if (!cfg) continue;
            const cur = Number(s[k]);
            if (!isFinite(cur)) continue;
            const step = Number(cfg.step) > 0 ? Number(cfg.step) : 1;
            let min = Number(cfg.min), max = Number(cfg.max);
            if (!isFinite(min) || cur < min) min = cur;
            if (!isFinite(max) || cur > max) max = cur;
            cfg.min = min;
            cfg.max = max;
        }
        next.baseline = Object.assign({}, s);
        return next;
    }

    // =========================================================================
    // 档位数 / 搜索空间（与 Utils.getValues 的取值规则保持一致）
    // =========================================================================

    /** 取值个数：min + k×step ≤ max（step 非法时按单点处理） */
    function countLevels(min, max, step) {
        const lo = Number(min), hi = Number(max), st = Number(step);
        if (!isFinite(lo) || !isFinite(hi)) return 0;
        if (!isFinite(st) || st <= 0) return lo <= hi ? 1 : 0;
        if (lo > hi) return 0;
        return Math.floor((hi - lo) / st + 1e-9) + 1;
    }

    /** 生成取值数组（与 Utils.getValues 完全等价，供无 Utils 环境使用） */
    function buildLevels(min, max, step) {
        const lo = Number(min), hi = Number(max), st = Number(step);
        if (!isFinite(lo) || !isFinite(hi)) return [];
        if (!isFinite(st) || st <= 0) return lo <= hi ? [lo] : [];
        const out = [];
        for (let v = lo; v <= hi + st * 1e-9; v += st) {
            out.push(roundTo(v, 6));
            if (out.length > 100000) break; // 防御：极端步长导致的无限循环
        }
        return out;
    }

    /** 搜索空间大小（5 个容量变量取值个数的乘积） */
    function searchSpaceSize(variables) {
        const vars = variables || state.optimizationConfig.variables || {};
        let space = 1;
        for (const k of SCHEME_KEYS) {
            const cfg = vars[k];
            const n = cfg ? countLevels(cfg.min, cfg.max, cfg.step) : 0;
            space *= Math.max(1, n);
        }
        return space;
    }

    /**
     * 批量方案生成（§40「批量方案计算」模块使用）。
     * 产生 schemeList[]，**绝不修改 AppState.currentScheme**。
     *
     * @param {Object} variables {windCapacity:{min,max,step}, ...}
     * @returns {Array<Object>} 方案数组（含 storageEnergy）
     */
    function createBatchSchemes(variables) {
        const vars = variables || {};
        const lists = SCHEME_KEYS.map(k => {
            const cfg = vars[k];
            if (!cfg) return [DEFAULT_SCHEME[k]];
            return buildLevels(cfg.min, cfg.max, cfg.step);
        });
        // 5 层笛卡尔积
        let combos = [[]];
        for (const list of lists) {
            const next = [];
            for (const acc of combos) {
                for (const v of list) next.push(acc.concat([v]));
            }
            combos = next;
        }
        return combos.map(arr => {
            const raw = {};
            SCHEME_KEYS.forEach((k, i) => { raw[k] = arr[i]; });
            return normalizeScheme(raw, DEFAULT_SCHEME);
        });
    }

    // =========================================================================
    // UI 同步（§20 / §21 / §38）
    // =========================================================================

    function hasDom() {
        return typeof document !== 'undefined' && !!document.getElementById;
    }

    function readNum(id, fallback) {
        if (!hasDom()) return fallback;
        const el = document.getElementById(id);
        if (!el) return fallback;
        const raw = (el.value === '' || el.value === null || el.value === undefined) ? NaN : Number(el.value);
        return isFinite(raw) ? raw : fallback;
    }

    function writeVal(id, v) {
        if (!hasDom()) return;
        const el = document.getElementById(id);
        if (el) el.value = v;
    }

    function writeText(id, t) {
        if (!hasDom()) return;
        const el = document.getElementById(id);
        if (el) el.textContent = t;
    }

    /**
     * 当前方案 → UI 输入框（唯一写出路径）。
     * 同时刷新「储能容量」只读派生字段与「当前方案状态栏」。
     */
    function syncSchemeToUI(scheme) {
        const s = normalizeScheme(scheme || state.currentScheme, state.currentScheme);
        for (const k of SCHEME_KEYS) writeVal(SCHEME_DOM[k], s[k]);
        renderSchemeStatus(s);
        return s;
    }

    /**
     * UI 输入框 → 当前方案（唯一读入路径）。
     * 用户在「电量计算」页改动任一容量输入时立即调用，保证 DOM 与真值不漂移（§38）。
     */
    function syncSchemeFromUI() {
        const raw = {};
        for (const k of SCHEME_KEYS) {
            const el = hasDom() ? document.getElementById(SCHEME_DOM[k]) : null;
            raw[k] = el ? readNum(SCHEME_DOM[k], state.currentScheme[k]) : state.currentScheme[k];
        }
        return setCurrentScheme(raw);
    }

    /** 运行参数 → / ← UI */
    function syncSimulationConfigToUI(config) {
        const c = Object.assign({}, state.simulationConfig, config || {});
        for (const k of SIM_KEYS) writeVal(SIM_DOM[k], c[k]);
        return Object.assign({}, c);
    }

    function syncSimulationConfigFromUI() {
        const raw = {};
        for (const k of SIM_KEYS) raw[k] = readNum(SIM_DOM[k], state.simulationConfig[k]);
        return setSimulationConfig(raw);
    }

    /**
     * 「当前方案状态栏」+ 派生字段渲染（§9 / §25 / §39）。
     * 电量计算页与优化页显示的是同一份 AppState.currentScheme，禁止各自维护。
     *
     * 注意：储能容量是 <input readonly>，必须写 value 而不是 textContent。
     */
    function renderSchemeStatus(scheme) {
        if (!hasDom()) return;
        const s = normalizeScheme(scheme || state.currentScheme, state.currentScheme);

        // 储能容量：派生量（储能功率 × 储能时长），只读展示
        writeVal(STORAGE_ENERGY_DOM, s.storageEnergy);

        writeText('currentSchemeStatusText',
            '风电 ' + s.windCapacity + ' MW　|　光伏 ' + s.pvCapacity + ' MW　|　' +
            '储能 ' + s.storagePower + ' MW / ' + s.storageDuration + ' h（' + s.storageEnergy + ' MWh）　|　' +
            '电解槽 ' + s.electrolyzerCapacity + ' MW');
        writeText('optBaselineStatusText',
            '风电 ' + s.windCapacity + ' MW　|　光伏 ' + s.pvCapacity + ' MW　|　' +
            '储能 ' + s.storagePower + ' MW / ' + s.storageDuration + ' h　|　' +
            '电解槽 ' + s.electrolyzerCapacity + ' MW');
        writeText('optBaselineKey', schemeKey(s));
    }

    // =========================================================================
    // 导出
    // =========================================================================

    const ParameterManager = {
        // 常量（只读约定，供全项目引用，避免字面量散落）
        SCHEME_KEYS: SCHEME_KEYS,
        SIM_KEYS: SIM_KEYS,
        SCHEME_LABELS: SCHEME_LABELS,
        SCHEME_UNITS: SCHEME_UNITS,
        SCHEME_DOM: SCHEME_DOM,
        SIM_DOM: SIM_DOM,
        STORAGE_ENERGY_DOM: STORAGE_ENERGY_DOM,
        DEFAULT_SCHEME: DEFAULT_SCHEME,
        DEFAULT_SIMULATION_CONFIG: DEFAULT_SIMULATION_CONFIG,
        DEFAULT_OPTIMIZATION_CONFIG: DEFAULT_OPTIMIZATION_CONFIG,

        // 方案参数
        normalizeScheme: normalizeScheme,
        getCurrentScheme: getCurrentScheme,
        setCurrentScheme: setCurrentScheme,
        validateScheme: validateScheme,
        schemeLabel: schemeLabel,
        schemeKey: schemeKey,

        // 运行参数
        getSimulationConfig: getSimulationConfig,
        setSimulationConfig: setSimulationConfig,
        validateSimulationConfig: validateSimulationConfig,

        // 优化参数
        getOptimizationConfig: getOptimizationConfig,
        setOptimizationConfig: setOptimizationConfig,
        validateOptimizationConfig: validateOptimizationConfig,
        checkBaselineInRange: checkBaselineInRange,
        expandRangeToIncludeBaseline: expandRangeToIncludeBaseline,

        // 档位 / 搜索空间 / 批量
        countLevels: countLevels,
        buildLevels: buildLevels,
        searchSpaceSize: searchSpaceSize,
        createBatchSchemes: createBatchSchemes,

        // UI 同步
        syncSchemeToUI: syncSchemeToUI,
        syncSchemeFromUI: syncSchemeFromUI,
        syncSimulationConfigToUI: syncSimulationConfigToUI,
        syncSimulationConfigFromUI: syncSimulationConfigFromUI,
        renderSchemeStatus: renderSchemeStatus,
    };

    self.ParameterManager = ParameterManager;
})();
