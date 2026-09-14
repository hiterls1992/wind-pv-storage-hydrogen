/**
 * ============================================================================
 * 多能互补容量优化引擎（NSGA-II）  ——  V2.1 新增
 * ============================================================================
 *
 * 设计原则
 * --------
 * 1. 本模块只解决「建多大」，不解决「每小时如何动态优化调度」。
 *    8760 小时仿真、投资估算、财务评价全部复用现有模块，本文件不重复实现。
 * 2. 与现有模块解耦：通过运行时全局对象（浏览器 window / Worker self）
 *    按需取得 runSingleSimulation / DataSummary / Estimate / FinanceEngine，
 *    因此本文件既可运行于主线程，也可用 importScripts() 运行于 Web Worker。
 * 3. 统一评价入口：evaluateScheme()，一次调用完成
 *    8760仿真 → 年度技术指标 → 投资估算 → 财务评价 → LCOH → 约束校验 → 目标向量。
 *    NSGA-II 完全依赖该统一结果，不直接触碰任何底层计算。
 *
 * 决策变量（5 个，全部为工程离散档位）
 * ------------------------------------
 *   X = [ windCapacity, pvCapacity, storagePower, storageDuration, electrolyzerCapacity ]
 *   storageEnergy = storagePower × storageDuration  （派生量，不作为独立决策变量）
 *
 * 优化目标（NSGA-II 内部统一为最小化）
 * ------------------------------------
 *   objective1 = -EIRR           （最大化 资本金内部收益率 EIRR）
 *   objective2 =  LCOH           （最小化 元/kg-H₂）
 *   objective3 =  CurtailmentRate （最小化 弃电率）
 *
 * 约束处理：可行性优先
 * --------------------
 *   可行解 > 不可行解；均不可行时，归一化总违反程度更小者优先。
 *
 * @version 2.1
 */

(function (root) {
    'use strict';

    const VERSION = '2.1';

    /** 运行时全局对象（浏览器 window / Web Worker self 通用） */
    function rt() {
        if (typeof self !== 'undefined') return self;
        if (typeof globalThis !== 'undefined') return globalThis;
        return root || {};
    }

    /** 决策变量顺序（全局唯一，禁止随意调整） */
    const KEY_ORDER = ['windCapacity', 'pvCapacity', 'storagePower', 'storageDuration', 'electrolyzerCapacity'];

    /** 变量中文名 / 单位（用于结果表与导出） */
    const VAR_META = {
        windCapacity:       { label: '风电容量',   unit: 'MW' },
        pvCapacity:         { label: '光伏容量',   unit: 'MW' },
        storagePower:       { label: '储能功率',   unit: 'MW' },
        storageDuration:    { label: '储能时长',   unit: 'h'  },
        electrolyzerCapacity: { label: '电解槽容量', unit: 'MW' },
    };

    /** 数值比较容差（相对容差，兼顾 -EIRR / LCOH / 弃电率 三种量级） */
    const REL_EPS = 1e-9;

    // ======================================================================
    // 一、基础数值工具
    // ======================================================================

    /** 小数位数（用于消除浮点误差，如 0.1+0.2） */
    function decimals(n) {
        const s = String(n);
        const i = s.indexOf('.');
        if (i < 0) return 0;
        return Math.min(6, s.length - i - 1);
    }

    function roundTo(n, d) {
        const p = Math.pow(10, d);
        return Math.round(n * p) / p;
    }

    /**
     * 工程容量离散化：把任意数值吸附到「min + k×step」上。
     * 严格保证 变量 ∈ [min, max] 且落在离散步长上，避免 87.326 MW 这类无意义容量。
     * @param {number} value 原始值
     * @param {number} min   下限
     * @param {number} max   上限
     * @param {number} step  步长（<=0 视为连续变量，仅做范围裁剪）
     * @returns {number}
     */
    function snapToStep(value, min, max, step) {
        if (!isFinite(min)) min = 0;
        if (!isFinite(max)) max = min;
        if (min > max) { const t = min; min = max; max = t; }
        if (!isFinite(value)) value = min;

        const d = Math.max(decimals(min), decimals(step));

        if (!isFinite(step) || step <= 0) {
            return roundTo(Math.min(Math.max(value, min), max), Math.max(d, 6));
        }

        const maxK = Math.max(0, Math.floor((max - min) / step + 1e-9));
        const k = Math.min(Math.max(Math.round((value - min) / step), 0), maxK);
        return roundTo(min + k * step, d);
    }

    /** 生成某变量在 [min,max] 上按 step 离散后的全部档位 */
    function buildLevelsFor(min, max, step) {
        if (!isFinite(min)) min = 0;
        if (!isFinite(max)) max = min;
        if (min > max) { const t = min; min = max; max = t; }

        const d = Math.max(decimals(min), decimals(step));
        const vals = [];

        if (!isFinite(step) || step <= 0) {
            vals.push(roundTo(min, 6));
            if (max > min) vals.push(roundTo(max, 6));
        } else {
            const n = Math.floor((max - min) / step + 1e-9);
            for (let k = 0; k <= n; k++) vals.push(roundTo(min + k * step, d));
        }

        // 去重 + 升序
        return Array.from(new Set(vals)).sort((a, b) => a - b);
    }

    /** 按 variables 配置生成全部变量的档位表 */
    function buildLevels(variables) {
        const levels = {};
        for (const key of KEY_ORDER) {
            const v = (variables && variables[key]) || {};
            levels[key] = buildLevelsFor(v.min, v.max, v.step);
        }
        return levels;
    }

    /** 档位表 → 搜索空间规模（用于自动收缩有效种群规模） */
    function searchSpaceSize(levels) {
        return KEY_ORDER.reduce((acc, k) => acc * Math.max(1, levels[k].length), 1);
    }

    /** 相对容差比较：a<b → -1；a>b → 1；相等 → 0 */
    function relCmp(a, b) {
        if (a === b) return 0;
        if (!isFinite(a) && !isFinite(b)) return 0;
        if (!isFinite(a)) return 1;    // Infinity（如 LCOH 无法计算）视为最差
        if (!isFinite(b)) return -1;
        const m = Math.max(Math.abs(a), Math.abs(b), 1e-12);
        if (a < b - REL_EPS * m) return -1;
        if (a > b + REL_EPS * m) return 1;
        return 0;
    }

    /** 可复现随机数发生器（mulberry32），保证同种子同结果 */
    function createRng(seed) {
        let a = (seed >>> 0) || 1;
        return function () {
            a = (a + 0x6D2B79F5) | 0;
            let t = Math.imul(a ^ (a >>> 15), 1 | a);
            t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
            return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
        };
    }

    /** 跨环境计时 */
    function nowMs() {
        const g = rt();
        if (g.performance && typeof g.performance.now === 'function') return g.performance.now();
        return Date.now();
    }

    // ======================================================================
    // 二、默认配置
    // ======================================================================

    /**
     * 生成 V2.1 默认优化配置（全部字段均可在前端修改，内核不含任何硬编码上限）
     */
    function buildDefaultConfig() {
        return {
            variables: {
                windCapacity:         { min: 50, max: 300, step: 25 },
                pvCapacity:           { min: 50, max: 500, step: 25 },
                storagePower:         { min: 0,  max: 200, step: 25 },
                storageDuration:      { min: 0,  max: 4,   step: 1  },
                electrolyzerCapacity: { min: 25, max: 200, step: 25 },
            },

            /** 工程约束：null / enabled=false 表示不启用 */
            constraints: {
                minAnnualHydrogen:     { enabled: true,  value: 1,   unit: '万吨/年' },
                maxCurtailmentRate:    { enabled: true,  value: 10,  unit: '%' },
                maxGridImportRatio:    { enabled: true,  value: 10,  unit: '%' },
                minGreenHydrogenRatio: { enabled: true,  value: 90,  unit: '%' },
                maxExportRatio:        { enabled: true,  value: 10,  unit: '%' },
                maxExportPower:        { enabled: false, value: 100, unit: 'MW' },
                maxImportPower:        { enabled: false, value: 100, unit: 'MW' },
            },

            /** 参与优化的目标开关（本版本三个目标默认全开） */
            objectives: { eirr: true, lcoh: true, curtailmentRate: true },

            /** NSGA-II 进化参数 */
            nsga2: {
                populationSize: 60,
                generations: 100,
                crossoverProbability: 0.9,
                mutationProbability: 0.1,
                distributionIndexCrossover: 20,   // SBX 分布指数
                randomSeed: 20260912,
                earlyStopping: true,
                patience: 15,
            },

            /** LCOH 折现率（%），默认取行业财务基准收益率（税前） */
            lcoh: { discountRate: 5.0 },

            /** 综合推荐权重（仅用于从 Pareto 前沿挑选推荐方案） */
            recommendationWeights: { eirr: 0.40, lcoh: 0.35, curtailmentRate: 0.25 },

            /** 基准方案（Baseline），由前端填充 */
            baseline: null,
        };
    }

    /** 深拷贝配置（避免前端对象被算法内部改写） */
    function normalizeConfig(config) {
        const base = buildDefaultConfig();
        const c = JSON.parse(JSON.stringify(Object.assign({}, base, config || {})));
        // 保证子对象完整
        c.variables = Object.assign({}, base.variables, c.variables || {});
        for (const key of KEY_ORDER) {
            c.variables[key] = Object.assign({}, base.variables[key], c.variables[key] || {});
        }
        c.constraints = Object.assign({}, base.constraints, c.constraints || {});
        c.objectives = Object.assign({}, base.objectives, c.objectives || {});
        c.nsga2 = Object.assign({}, base.nsga2, c.nsga2 || {});
        c.recommendationWeights = Object.assign({}, base.recommendationWeights, c.recommendationWeights || {});
        c.lcoh = Object.assign({}, base.lcoh, c.lcoh || {});
        return c;
    }

    // ======================================================================
    // 三、LCOH（平准化制氢成本）
    // ======================================================================

    /**
     * LCOH = 项目全生命周期折现成本 / 项目全生命周期折现制氢量   [元/kg-H₂]
     *
     * 计算口径（与现有财务评价口径保持一致的「不含折旧、不含所得税」现金流成本口径）：
     *   分子 C_t（万元）：建设期 = (建设投资 + 建设期利息) / 建设期；投产年另加流动资金；
     *                     运营期 = 年经营成本（不含折旧、不含利息、不含所得税），大修年另加大修费
     *   分母 H_t（kg）  ：运营期 = 年制氢量 = 制氢量总和（万吨）× 1e7
     *   折现率 r        ：默认取行业财务基准收益率（税前）5%，可在前端修改
     *
     * 说明：不采用现有「制氢成本（元/kg）」作为 LCOH，后者为未折现的运营期平均成本口径，
     *      两者在界面上同时展示以便对照，但优化目标统一使用本函数结果。
     *
     * @param {Object} detail FinanceEngine.calculateAll() 返回的 detail
     * @param {number} discountRatePercent 折现率（%）
     * @returns {number} LCOH（元/kg）；制氢量为 0 或成本不可得时返回 Infinity
     */
    function calculateLCOH(detail, discountRatePercent) {
        if (!detail || !(detail.annual_h2_kg > 0)) return Infinity;

        const N = detail.calc_period;
        const cp = detail.construct_period;
        if (!(N > 0) || !(cp >= 0)) return Infinity;

        const r = (isFinite(discountRatePercent) ? discountRatePercent : 5.0) / 100.0;
        const capexTotal = (detail.construction_investment || 0) + (detail.construction_interest || 0);
        const capexPerYear = capexTotal / Math.max(cp, 1);

        let discountedCost = 0;   // 万元
        let discountedH2 = 0;     // kg

        for (let t = 0; t < N; t++) {
            const df = 1 / Math.pow(1 + r, t);

            let cost = 0;
            if (t < cp) {
                cost = capexPerYear;
            } else {
                cost = detail.annual_opex || 0;
                if (t === detail.overhaul_year) cost += (detail.overhaul_total || 0);
                if (t === cp) cost += (detail.working_capital || 0);   // 流动资金于投产年投入
                discountedH2 += detail.annual_h2_kg * df;
            }

            discountedCost += cost * df;
        }

        if (!(discountedH2 > 0)) return Infinity;
        return discountedCost * 1e4 / discountedH2;   // 万元 → 元
    }

    // ======================================================================
    // 四、统一单方案评价函数 evaluateScheme()
    // ======================================================================

    /**
     * 方案唯一键（任务书 §49）。
     *
     * V2.2 起**统一委托 ParameterManager.schemeKey()**，全项目只允许存在一种方案标识格式：
     *     W200|PV360|B100|H2|EL160
     * 统一用于：缓存 / 结果索引 / 方案选择 / Pareto方案 / Baseline / Excel导出。
     */
    function schemeKey(scheme) {
        const g = rt();
        if (g.ParameterManager && typeof g.ParameterManager.schemeKey === 'function') {
            return g.ParameterManager.schemeKey(scheme);
        }
        // 兜底（正常不会走到：collaborators() 已把 ParameterManager 列为必需模块）
        return [
            scheme.windCapacity,
            scheme.pvCapacity,
            scheme.storagePower,
            scheme.storageDuration,
            scheme.electrolyzerCapacity,
        ].join('|');
    }

    /** 补全派生量 storageEnergy 并做安全化 */
    function completeScheme(scheme) {
        const s = {
            windCapacity: Number(scheme.windCapacity) || 0,
            pvCapacity: Number(scheme.pvCapacity) || 0,
            storagePower: Number(scheme.storagePower) || 0,
            storageDuration: Number(scheme.storageDuration) || 0,
            electrolyzerCapacity: Number(scheme.electrolyzerCapacity) || 0,
        };
        s.storageEnergy = roundTo(s.storagePower * s.storageDuration, 6);
        return s;
    }

    /**
     * 运行参数归一化。
     *
     * TODO V2.3 REMOVE LEGACY
     *   兼容 V2.1 的 UPPER_SNAKE 字段名（ELECTROLYZER_MIN_RATIO 等）。
     *   兼容层只做「字段名翻译」，不改变任何取值；
     *   业务逻辑必须使用 V2.2 规范的 camelCase 字段（任务书 §37）。
     */
    function toSimulationConfig(raw) {
        if (!raw) return null;
        // 已是 V2.2 规范字段 → 原样返回
        if (raw.electrolyzerMinRatio !== undefined || raw.hydrogenConsumption !== undefined ||
            raw.chargeEfficiency !== undefined || raw.maxImportRatio !== undefined) {
            return raw;
        }
        // V2.1 旧字段名 → 翻译
        return {
            electrolyzerMinRatio: raw.ELECTROLYZER_MIN_RATIO,
            maxExportHourly: raw.MAX_EXPORT_RATIO_HOURLY,
            maxExportTotal: raw.MAX_EXPORT_RATIO_TOTAL,
            maxImportRatio: raw.MAX_IMPORT_RATIO,
            chargeEfficiency: raw.STORAGE_CHARGE_EFFICIENCY,
            dischargeEfficiency: raw.STORAGE_DISCHARGE_EFFICIENCY,
            hydrogenConsumption: raw.HYDROGEN_ENERGY_CONSUMPTION,
        };
    }

    /**
     * 上下文键名兼容。
     * 正式约定：pvData / windData / simulationConfig；
     * 同时兼容 pv / wind / simParams，避免调用方字段命名不一致时静默失败。
     */
    function normalizeContext(ctx) {
        if (!ctx) return ctx;
        if (!ctx.pvData && ctx.pv) ctx.pvData = ctx.pv;
        if (!ctx.windData && ctx.wind) ctx.windData = ctx.wind;
        // TODO V2.3 REMOVE LEGACY：旧字段 simParams → 新字段 simulationConfig
        if (!ctx.simulationConfig && ctx.simParams) {
            ctx.simulationConfig = toSimulationConfig(ctx.simParams);
        } else if (ctx.simulationConfig) {
            ctx.simulationConfig = toSimulationConfig(ctx.simulationConfig);
        }
        return ctx;
    }

    /** 取运行时协作模块（延迟解析，保证与脚本加载顺序解耦） */
    function collaborators() {
        const g = rt();
        const missing = [];
        if (typeof g.runSingleSimulation !== 'function') missing.push('simulation-engine.js (runSingleSimulation)');
        if (!g.ParameterManager || typeof g.ParameterManager.schemeKey !== 'function') missing.push('parameter-manager.js (ParameterManager)');
        if (!g.DataSummary || typeof g.DataSummary.generateSummary !== 'function') missing.push('data-summary.js (DataSummary)');
        if (!g.Estimate || typeof g.Estimate.batchEstimate !== 'function') missing.push('estimate.js (Estimate)');
        if (!g.FinanceEngine || typeof g.FinanceEngine.calculateAll !== 'function') missing.push('finance-engine.js (FinanceEngine)');
        if (missing.length) {
            throw new Error('优化引擎依赖模块未就绪：' + missing.join('、') + '（请检查脚本加载顺序）');
        }
        return {
            runSingleSimulation: g.runSingleSimulation,
            DataSummary: g.DataSummary,
            Estimate: g.Estimate,
            FinanceEngine: g.FinanceEngine,
        };
    }

    /**
     * 财务内部收益率解算结果规范化。
     * 现有财务引擎在「现金流全为同号 / 二分法找不到符号变化」时返回 0，
     * 该返回值并非真实 IRR，按规范（五十二）必须转成 null，避免误导优化器。
     */
    function resolveIrr(irrValue, cashflows) {
        if (irrValue === null || irrValue === undefined || !isFinite(irrValue)) return null;
        if (!Array.isArray(cashflows) || cashflows.length === 0) return irrValue;
        const hasPos = cashflows.some(v => v > 0);
        const hasNeg = cashflows.some(v => v < 0);
        if (!hasPos || !hasNeg) return null;
        if (irrValue === 0) return null;   // 现有引擎「无解」的返回值
        return irrValue;
    }

    /**
     * 仿真键（V2.3.1 任务书 §9 / §10）：委托 ResultDataStore.createSimulationKey。
     * 结果数据层未加载时退化为 schemeKey（与 V2.3 行为一致，不中断计算）。
     */
    function simulationKeyFor(scheme, schemeKeyStr, context) {
        const g = rt();
        if (g.ResultDataStore && typeof g.ResultDataStore.createSimulationKey === 'function') {
            return g.ResultDataStore.createSimulationKey({
                scheme: scheme,
                schemeKey: schemeKeyStr,
                simulationConfig: context ? context.simulationConfig : undefined,
                inputVersion: context ? context.inputVersion : undefined,
                engineVersion: (typeof SIMULATION_ENGINE_VERSION !== 'undefined')
                    ? SIMULATION_ENGINE_VERSION : undefined,
            });
        }
        return schemeKeyStr;
    }

    /**
     * 统一单方案评价：一组容量参数 → 完整技术经济评价
     *
     * 输入：scheme（5 个容量参数）+ context（风光8760数据 / 仿真参数 / 概算单价 / 财务参数）
     * 输出：{ scheme, technical, economic, constraints, objectives, vector }
     *
     * @param {Object} schemeIn
     * @param {Object} context
     * @returns {Object} 统一结果对象
     */
    function evaluateScheme(schemeIn, context) {
        const t0 = nowMs();
        normalizeContext(context);
        const scheme = completeScheme(schemeIn);
        const key = schemeKey(scheme);

        // V2.3.1（任务书 §9 / §13）：评价缓存的键从 schemeKey 升级为 simulationKey。
        // schemeKey 只描述"建多大"；simulationKey = 方案 + 运行参数 + 输入数据版本 + 算法版本，
        // 保证"同方案不同运行参数 / 不同输入数据"不会命中旧缓存。
        const simKey = simulationKeyFor(scheme, key, context);

        // ---- 缓存：同一方案禁止重复运行 8760 小时仿真 ----
        if (context.cache && context.cache.has(simKey)) {
            if (context.stats) context.stats.cacheHits++;
            return context.cache.get(simKey);
        }

        const co = collaborators();
        const pvData = context.pvData;
        const windData = context.windData;

        // ---- 1. 8760 小时仿真（唯一入口：容量走 scheme，运行规则走 simulationConfig） ----
        // 任务书 §36：禁止再把容量 Object.assign 进运行参数对象（PV_CAPACITY/WIND_CAPACITY 注入已删除）
        // V2.3：记录仿真耗时，用于 Worker 性能诊断（只累加统计量，不影响任何计算）
        const simT0 = nowMs();
        const sim = co.runSingleSimulation(
            pvData, windData, scheme, context.simulationConfig
        );
        const simElapsedMs = nowMs() - simT0;

        // ---- 2. 年度技术指标（复用 data-summary.js） ----
        const summaryRow = co.DataSummary.generateSummary([sim], scheme.pvCapacity, scheme.windCapacity)[0];

        // ---- 3. 投资估算（复用 estimate.js） ----
        const estimateRow = co.Estimate.batchEstimate([summaryRow], context.prices)[0];

        // ---- 4. 财务评价（复用 finance-engine.js） ----
        const fin = co.FinanceEngine.calculateAll(context.financeParams, summaryRow, estimateRow);

        // ---- 5. 逐时补充统计（理论上网/实际发电量、储能吞吐、逐时最大上网/下网功率） ----
        const sums = sim.sums;
        const hours = pvData.length;
        const res = sim.results;
        let sumPv = 0, sumWind = 0, sumCharge = 0, sumDischarge = 0;
        let maxHourlyExport = 0, maxHourlyImport = 0;
        for (let h = 0; h < hours; h++) {
            const idx = h * 11;
            sumPv += res[idx + 0];
            sumWind += res[idx + 1];
            sumCharge += res[idx + 3];
            sumDischarge += res[idx + 4];
            if (res[idx + 7] > maxHourlyExport) maxHourlyExport = res[idx + 7];
            if (res[idx + 8] > maxHourlyImport) maxHourlyImport = res[idx + 8];
        }

        const theoretical = sums.sumTotal;   // = 风电理论发电量 + 光伏理论发电量
        const electrolyzerElectricity = sums.sumHydrogenPower;   // 电解槽总耗电量

        const curtailmentRate = theoretical > 0 ? sums.sumCurtailment / theoretical : 0;
        const gridImportRatio = electrolyzerElectricity > 0 ? sums.sumImport / electrolyzerElectricity : 0;
        const greenHydrogenRatio = electrolyzerElectricity > 0 ? Math.max(0, 1 - gridImportRatio) : 0;
        const exportRatio = theoretical > 0 ? sums.sumExport / theoretical : 0;

        const annualHydrogenKg = sums.sumH2Prod;

        const technical = {
            annualWindEnergy: sumWind,                 // MWh
            annualPvEnergy: sumPv,                     // MWh
            theoreticalEnergy: theoretical,            // MWh（风光理论发电量）
            annualHydrogenKg: annualHydrogenKg,        // kg
            annualHydrogenTon: annualHydrogenKg / 1000,
            annualHydrogenWanTon: annualHydrogenKg / 1e7,
            electrolyzerElectricity: electrolyzerElectricity,     // MWh
            electrolyzerHours: summaryRow['电解槽利用小时数'] || 0,
            electrolyzerLoadFactor: (scheme.electrolyzerCapacity > 0 && hours > 0)
                ? (electrolyzerElectricity / (scheme.electrolyzerCapacity * hours)) : 0,
            annualChargeEnergy: sumCharge,
            annualDischargeEnergy: sumDischarge,
            annualGridImport: sums.sumImport,
            annualGridExport: sums.sumExport,
            annualCurtailment: sums.sumCurtailment,
            curtailmentRate: curtailmentRate,
            greenHydrogenRatio: greenHydrogenRatio,
            gridImportRatio: gridImportRatio,
            exportRatio: exportRatio,
            maxHourlyExportPower: maxHourlyExport,
            maxHourlyImportPower: maxHourlyImport,
        };

        // ---- 6. 财务指标 ----
        const R = fin.result;
        const preTaxCFs = (fin.detail.project_cf || []).map(cf => cf.net_cf_pre_tax);
        const firr = resolveIrr(R['项目投资财务内部收益率（所得税前）（%）'], preTaxCFs);
        const firrPostTax = resolveIrr(R['项目投资财务内部收益率（所得税后）（%）'], preTaxCFs);
        // V2.3.2：优化目标的经济口径切换为「资本金内部收益率（EIRR）」。
        // 与 FIRR 相同的无解判定（现金流无符号变化 / 引擎返回 0）→ null = 经济评价不可行。
        const equityCFs = (fin.detail.equity_cf || []).map(cf => cf.net_cf);
        const eirr = resolveIrr(R['资本金财务内部收益率（%）'], equityCFs);

        const economic = {
            totalInvestment: R['项目总投资（万元）'],
            constructionInvestment: R['建设投资（万元）'],
            constructionInterest: R['建设期利息（万元）'],
            workingCapital: R['流动资金（万元）'],
            annualRevenue: fin.detail.annual_revenue,
            annualCost: fin.detail.annual_total_cost,
            FIRR: firr,
            FIRRPostTax: firrPostTax,
            EIRR: R['资本金财务内部收益率（%）'],
            NPV: R['项目投资财务净现值（所得税前）（万元）'],
            paybackPeriod: R['项目投资回收期（所得税前）（年）'],
            ROI: R['总投资收益率（ROI）（%）'],
            ROE: R['项目资本金净利润率（ROE）（%）'],
            breakEvenPoint: R['盈亏平衡点（生产能力利用率）（%）'],
            /** LCOH：折现口径（本版本优化目标） */
            LCOH: calculateLCOH(fin.detail, context.lcohDiscountRate),
            /** 制氢成本：V1.0 未折现口径（仅作对照展示） */
            h2CostAvg: R['制氢成本（元/kg）'],
        };

        // ---- 7. 目标向量（统一最小化） ----
        const objectives = {
            eirr: eirr,                                    // %（null 表示经济评价不可行）
            lcoh: economic.LCOH,                           // 元/kg
            curtailmentRate: curtailmentRate,              // 0~1
        };
        const vector = [
            (eirr === null) ? 1e6 : -eirr,                 // 最大化 EIRR → 最小化 -EIRR
            economic.LCOH,                                  // 最小化 LCOH
            curtailmentRate,                                // 最小化弃电率
        ];

        const unified = {
            scheme: scheme,
            key: key,
            simulationKey: simKey,
            technical: technical,
            economic: economic,
            constraints: null,     // 见下
            objectives: objectives,
            vector: vector,
            durationMs: nowMs() - t0,
        };

        // ---- 8. 约束校验 ----
        unified.constraints = checkConstraints(unified, context.config);

        // ---- 写缓存 ----
        if (context.cache) context.cache.set(simKey, unified);
        if (context.stats) {
            context.stats.evaluated++;
            // V2.3：性能诊断用的累计耗时（§14），不参与任何优化计算
            context.stats.evalTimeMs = (context.stats.evalTimeMs || 0) + (nowMs() - t0);
            context.stats.simTimeMs = (context.stats.simTimeMs || 0) + simElapsedMs;
        }

        return unified;
    }

    // ======================================================================
    // 五、约束处理（可行性优先）
    // ======================================================================

    /**
     * 工程约束校验。
     * 每条约束独立支持 启用/禁用，并计算归一化违反程度，供不可行解之间排序使用。
     * @param {Object} result 统一评价结果（含 technical / economic）
     * @param {Object} config 优化配置
     * @returns {{feasible:boolean, violations:Array, violationAmount:number, economicViable:boolean}}
     */
    function checkConstraints(result, config) {
        const violations = [];
        const c = (config && config.constraints) || {};
        const tech = result.technical;
        const EPS = 1e-9;

        const on = (cfg) => !!(cfg && cfg.enabled && isFinite(Number(cfg.value)));

        /** 记录一条违反项，amount 为归一化违反程度（越大约严重） */
        function addViolation(name, actual, limit, unit, amount) {
            violations.push({ name: name, actual: actual, limit: limit, unit: unit, amount: Math.max(0, amount) });
        }

        // 8.1 年制氢量 ≥  MinimumHydrogen（前端 万吨/年，内部 kg）
        if (on(c.minAnnualHydrogen)) {
            const limWanTon = Number(c.minAnnualHydrogen.value);
            const limKg = limWanTon * 1e7;
            const act = tech.annualHydrogenKg;
            if (!(act >= limKg - EPS)) {
                addViolation('年制氢量', act / 1e7, limWanTon, '万吨/年',
                    (limKg - act) / Math.max(limKg, 1));
            }
        }

        // 11. 弃电率 ≤ MaxCurtailmentRate
        if (on(c.maxCurtailmentRate)) {
            const lim = Number(c.maxCurtailmentRate.value) / 100;
            const act = tech.curtailmentRate;
            if (act > lim + EPS) {
                addViolation('弃电率', act * 100, Number(c.maxCurtailmentRate.value), '%',
                    (act - lim) / Math.max(lim, 1e-6));
            }
        }

        // 9. 外购电比例 ≤ MaxGridImportRatio
        //    GridImportRatio = 年下网电量 / 电解槽总耗电量
        if (on(c.maxGridImportRatio)) {
            const lim = Number(c.maxGridImportRatio.value) / 100;
            const act = tech.gridImportRatio;
            if (act > lim + EPS) {
                addViolation('外购电比例', act * 100, Number(c.maxGridImportRatio.value), '%',
                    (act - lim) / Math.max(lim, 1e-6));
            }
        }

        // 10. 绿电制氢比例 ≥ MinimumGreenHydrogenRatio
        //     GreenHydrogenRatio = 用于制氢的新能源电量 / 电解槽总耗电量 = 1 - 外购电比例
        if (on(c.minGreenHydrogenRatio)) {
            const lim = Number(c.minGreenHydrogenRatio.value) / 100;
            const act = tech.greenHydrogenRatio;
            if (act < lim - EPS) {
                addViolation('绿电制氢比例', act * 100, Number(c.minGreenHydrogenRatio.value), '%',
                    (lim - act) / Math.max(lim, 1e-6));
            }
        }

        // 12. 上网比例 ≤ MaxExportRatio（复用现有总量上网比例上限逻辑）
        if (on(c.maxExportRatio)) {
            const lim = Number(c.maxExportRatio.value) / 100;
            const act = tech.exportRatio;
            if (act > lim + EPS) {
                addViolation('上网比例', act * 100, Number(c.maxExportRatio.value), '%',
                    (act - lim) / Math.max(lim, 1e-6));
            }
        }

        // 13. 逐时最大上网功率 ≤ MaxExportPower
        if (on(c.maxExportPower)) {
            const lim = Number(c.maxExportPower.value);
            const act = tech.maxHourlyExportPower;
            if (act > lim + EPS) {
                addViolation('最大上网功率', act, lim, 'MW', (act - lim) / Math.max(lim, 1e-6));
            }
        }

        // 14. 逐时最大下网功率 ≤ MaxImportPower
        if (on(c.maxImportPower)) {
            const lim = Number(c.maxImportPower.value);
            const act = tech.maxHourlyImportPower;
            if (act > lim + EPS) {
                addViolation('最大下网功率', act, lim, 'MW', (act - lim) / Math.max(lim, 1e-6));
            }
        }

        // 52. 财务指标异常：资本金内部收益率（EIRR，优化经济目标）无法计算的方案判为「经济评价不可行」
        const economicViable = (result.economic.EIRR !== null && isFinite(result.economic.EIRR));
        if (!economicViable) {
            addViolation('资本金内部收益率(EIRR)', null, null, '%', 1.0);
        }

        const violationAmount = violations.reduce((s, v) => s + v.amount, 0);

        return {
            feasible: violations.length === 0,
            violations: violations,
            violationAmount: violationAmount,
            economicViable: economicViable,
        };
    }

    // ======================================================================
    // 六、NSGA-II 核心算子
    // ======================================================================

    /**
     * NSGA-II 支配关系（含约束处理）
     * @returns {number} 1: a 支配 b；-1: b 支配 a；0: 互不支配
     */
    function dominates(a, b) {
        const af = !!a.feasible, bf = !!b.feasible;

        // 可行性优先
        if (af !== bf) return af ? 1 : -1;

        // 均不可行：总违反程度更小者优先
        if (!af) {
            const c = relCmp(a.violationAmount, b.violationAmount);
            if (c < 0) return 1;
            if (c > 0) return -1;
            return 0;
        }

        // 均可行：Pareto 支配（三个目标全部最小化）
        const va = a.vector, vb = b.vector;
        let better = false, worse = false;
        for (let i = 0; i < va.length; i++) {
            const c = relCmp(va[i], vb[i]);
            if (c < 0) better = true;
            else if (c > 0) worse = true;
        }
        if (better && !worse) return 1;
        if (worse && !better) return -1;
        return 0;
    }

    /**
     * 非支配排序 → Front 1 / Front 2 / ...
     * @returns {Array<Array>} fronts，并给每个个体写入 ind.rank
     */
    function nonDominatedSort(population) {
        const n = population.length;
        if (n === 0) return [];

        // 预建索引映射，避免在循环内使用 indexOf（性能关键）
        const idxMap = new Map();
        for (let i = 0; i < n; i++) idxMap.set(population[i], i);

        const dominatedSet = new Array(n);
        const domCount = new Array(n);
        const fronts = [[]];

        for (let i = 0; i < n; i++) {
            dominatedSet[i] = [];
            domCount[i] = 0;
            for (let j = 0; j < n; j++) {
                if (i === j) continue;
                const d = dominates(population[i], population[j]);
                if (d === 1) dominatedSet[i].push(j);
                else if (dominates(population[j], population[i]) === 1) domCount[i]++;
            }
            if (domCount[i] === 0) {
                population[i].rank = 1;
                fronts[0].push(population[i]);
            }
        }

        let fi = 0;
        while (fronts[fi] && fronts[fi].length > 0) {
            const next = [];
            for (const ind of fronts[fi]) {
                const idx = idxMap.get(ind);
                for (const k of dominatedSet[idx]) {
                    domCount[k]--;
                    if (domCount[k] === 0) {
                        population[k].rank = fi + 2;
                        next.push(population[k]);
                    }
                }
            }
            fi++;
            if (next.length > 0) fronts.push(next);
            else break;
        }

        return fronts;
    }

    /** 拥挤距离（三目标分别计算，边界方案取 Infinity） */
    function calculateCrowdingDistance(front) {
        const n = front.length;
        for (const ind of front) ind.crowding = 0;
        if (n === 0) return;
        if (n <= 2) {
            for (const ind of front) ind.crowding = Infinity;
            return;
        }

        const m = front[0].vector.length;
        for (let k = 0; k < m; k++) {
            const sorted = front.slice().sort((a, b) => (a.vector[k] - b.vector[k]) || 0);
            sorted[0].crowding = Infinity;
            sorted[n - 1].crowding = Infinity;

            const min = sorted[0].vector[k];
            const max = sorted[n - 1].vector[k];
            const span = max - min;
            if (!isFinite(span) || span <= 0) continue;

            for (let i = 1; i < n - 1; i++) {
                const d = (sorted[i + 1].vector[k] - sorted[i - 1].vector[k]) / span;
                if (isFinite(d)) sorted[i].crowding += d;
            }
        }
    }

    /** 个体优劣比较（供锦标赛选择使用）：rank 优先，其次拥挤距离 */
    function isBetterIndividual(a, b, rng) {
        if (a.rank !== b.rank) return a.rank < b.rank;
        const ca = isFinite(a.crowding) ? a.crowding : Infinity;
        const cb = isFinite(b.crowding) ? b.crowding : Infinity;
        if (ca !== cb) return ca > cb;
        return (rng ? rng() : 0.5) < 0.5;
    }

    /** 二元锦标赛选择 */
    function tournamentSelection(population, rng) {
        const n = population.length;
        const i = Math.min(n - 1, Math.floor(rng() * n));
        const j = Math.min(n - 1, Math.floor(rng() * n));
        if (i === j) return population[i];
        return isBetterIndividual(population[i], population[j], rng) ? population[i] : population[j];
    }

    /**
     * 模拟二进制交叉 SBX（子代严格落在变量范围内且吸附到步长）
     * @returns {[Array<number>, Array<number>]} 两个子代基因数组
     */
    function sbxCrossover(parent1, parent2, config, rng) {
        const eta = Number(config.nsga2.distributionIndexCrossover) || 20;
        const c1 = parent1.genes.slice();
        const c2 = parent2.genes.slice();

        if (rng() > Number(config.nsga2.crossoverProbability)) return [c1, c2];

        for (let i = 0; i < KEY_ORDER.length; i++) {
            if (rng() > 0.5) continue;   // 每个基因 50% 概率参与交叉

            const v = config.variables[KEY_ORDER[i]];
            const yl = Number(v.min), yu = Number(v.max);
            if (!(yu - yl > 1e-12)) continue;   // 单一档位，无需交叉

            const a = Math.min(c1[i], c2[i]);
            const b = Math.max(c1[i], c2[i]);
            if (b - a < 1e-12) continue;       // 两父代该基因相同

            const u = rng();

            // 子代 1
            const beta1 = 1 + 2 * (a - yl) / (b - a);
            const alpha1 = 2 - Math.pow(beta1, -(eta + 1));
            const betaq1 = (u <= 1 / alpha1)
                ? Math.pow(u * alpha1, 1 / (eta + 1))
                : Math.pow(1 / (2 - u * alpha1), 1 / (eta + 1));
            const child1 = 0.5 * ((a + b) - betaq1 * (b - a));

            // 子代 2
            const beta2 = 1 + 2 * (yu - b) / (b - a);
            const alpha2 = 2 - Math.pow(beta2, -(eta + 1));
            const betaq2 = (u <= 1 / alpha2)
                ? Math.pow(u * alpha2, 1 / (eta + 1))
                : Math.pow(1 / (2 - u * alpha2), 1 / (eta + 1));
            const child2 = 0.5 * ((a + b) + betaq2 * (b - a));

            // 交换
            if (rng() <= 0.5) { c1[i] = child2; c2[i] = child1; }
            else { c1[i] = child1; c2[i] = child2; }

            // 吸附到工程步长（禁止出现 87.326 MW 这类无意义容量）
            c1[i] = snapToStep(c1[i], yl, yu, v.step);
            c2[i] = snapToStep(c2[i], yl, yu, v.step);
        }

        return [c1, c2];
    }

    /**
     * 变异：针对离散工程变量，采用「随机向上/向下跳动 1~2 个步长」策略
     * （NSGA-II 标准中的 Polynomial Mutation 在整数档位变量上等价于该离散操作，
     *   且能天然保证子代落在步长上、不越界、不产生负值与 NaN）。
     */
    function mutation(genes, config, rng) {
        const out = genes.slice();
        for (let i = 0; i < KEY_ORDER.length; i++) {
            if (rng() > Number(config.nsga2.mutationProbability)) continue;

            const v = config.variables[KEY_ORDER[i]];
            const yl = Number(v.min), yu = Number(v.max);
            const step = (Number(v.step) > 0) ? Number(v.step) : (yu - yl);
            if (!(step > 0)) continue;

            const dir = (rng() < 0.5) ? -1 : 1;
            const k = 1 + Math.floor(rng() * 2);           // 跳动 1~2 个步长
            out[i] = snapToStep(out[i] + dir * step * k, yl, yu, v.step);
        }
        return out;
    }

    /**
     * 精英保留下生成下一代：父代 + 子代合并 → 非支配排序 → 按 rank、拥挤距离截断
     */
    function createNextGeneration(population, offspring, config) {
        const targetSize = config.nsga2.populationSize;
        const combined = population.concat(offspring);
        const fronts = nonDominatedSort(combined);

        const next = [];
        for (const front of fronts) {
            if (next.length >= targetSize) break;
            calculateCrowdingDistance(front);
            if (next.length + front.length <= targetSize) {
                for (const ind of front) next.push(ind);
            } else {
                front.sort((a, b) => {
                    const ca = isFinite(a.crowding) ? a.crowding : Infinity;
                    const cb = isFinite(b.crowding) ? b.crowding : Infinity;
                    return cb - ca;
                });
                for (let i = 0; i < targetSize - next.length; i++) next.push(front[i]);
            }
        }

        return next;
    }

    /** 为新种群重算 rank 与拥挤距离（锦标赛选择依赖这两个字段） */
    function refreshRankAndCrowding(population) {
        const fronts = nonDominatedSort(population);
        for (const front of fronts) calculateCrowdingDistance(front);
    }

    /** 取当前种群的 Pareto 前沿（Front 1 中的可行解） */
    function getParetoFront(population) {
        const fronts = nonDominatedSort(population);
        const first = fronts[0] || [];
        return first.filter(ind => ind.feasible);
    }

    // ======================================================================
    // 七、种群初始化
    // ======================================================================

    /** 生成一个随机个体（严格落在档位上） */
    function createIndividual(config, rng) {
        const genes = KEY_ORDER.map(key => {
            const v = config.variables[key];
            const min = Number(v.min), max = Number(v.max);
            const raw = min + rng() * Math.max(0, max - min);
            return snapToStep(raw, min, max, v.step);
        });
        return makeIndividual(genes, null);
    }

    /** 由基因数组构造个体对象 */
    function makeIndividual(genes, evaluation) {
        const scheme = completeScheme({
            windCapacity: genes[0],
            pvCapacity: genes[1],
            storagePower: genes[2],
            storageDuration: genes[3],
            electrolyzerCapacity: genes[4],
        });
        return {
            genes: genes.slice(),
            key: schemeKey(scheme),
            scheme: scheme,
            evaluation: evaluation || null,
            vector: evaluation ? evaluation.vector : null,
            feasible: evaluation ? evaluation.constraints.feasible : false,
            violationAmount: evaluation ? evaluation.constraints.violationAmount : Infinity,
            rank: null,
            crowding: 0,
        };
    }

    /** 把个体与缓存中的评价结果绑定 */
    function bindEvaluation(ind, ctx) {
        const ev = evaluateScheme(ind.scheme, ctx);
        ind.evaluation = ev;
        ind.vector = ev.vector;
        ind.feasible = ev.constraints.feasible;
        ind.violationAmount = ev.constraints.violationAmount;
        return ind;
    }

    /**
     * 初始化种群：优先保证个体唯一（Set 去重）。
     * 若搜索空间规模小于目标种群规模，自动降低有效种群数量；
     * 极端情况下（搜索空间很小）直接以全组合枚举作为初始种群。
     */
    function initializePopulation(config, ctx) {
        const levels = buildLevels(config.variables);
        const spaceSize = searchSpaceSize(levels);
        const target = Math.max(2, Math.min(Number(config.nsga2.populationSize) || 60, spaceSize));
        const effective = target;

        const seen = new Set();
        const population = [];
        const maxTries = effective * 60;

        // 搜索空间不大时，直接用穷举组合（保证初始种群尽可能分散）
        let enumerated = null;
        if (spaceSize <= Math.max(effective * 4, 400)) {
            enumerated = [];
            const walk = (i, acc) => {
                if (i === KEY_ORDER.length) { enumerated.push(acc.slice()); return; }
                for (const val of levels[KEY_ORDER[i]]) { acc.push(val); walk(i + 1, acc); acc.pop(); }
            };
            walk(0, []);
        }

        if (enumerated && enumerated.length >= effective) {
            // 从全组合中随机抽取，保证唯一性
            const pool = enumerated.slice();
            const picked = [];
            const need = Math.min(effective, pool.length);
            for (let i = 0; i < need; i++) {
                const k = Math.floor(ctx.rng() * pool.length);
                picked.push(pool.splice(k, 1)[0]);
            }
            for (const genes of picked) population.push(makeIndividual(genes, null));
        } else {
            let tries = 0;
            while (population.length < effective && tries < maxTries) {
                tries++;
                const ind = createIndividual(config, ctx.rng);
                if (seen.has(ind.key)) continue;
                seen.add(ind.key);
                population.push(ind);
            }
            // 兜底：随机生成仍不足时，用枚举组合补齐
            if (population.length < effective && enumerated) {
                for (const genes of enumerated) {
                    if (population.length >= effective) break;
                    const key = schemeKey(completeScheme({
                        windCapacity: genes[0], pvCapacity: genes[1], storagePower: genes[2],
                        storageDuration: genes[3], electrolyzerCapacity: genes[4],
                    }));
                    if (seen.has(key)) continue;
                    seen.add(key);
                    population.push(makeIndividual(genes, null));
                }
            }
        }

        return {
            population: population,
            searchSpaceSize: spaceSize,
            effectivePopulationSize: population.length,
            levels: levels,
        };
    }

    // ======================================================================
    // 八、综合推荐（权重仅用于 Pareto 前沿内部择优）
    // ======================================================================

    /**
     * 归一化目标值 → 各目标得分（越大越好）∈ [0,1]
     */
    function normalizeObjectives(solutions) {
        const pick = (fn) => solutions.map(fn).filter(v => v !== null && isFinite(v));
        const eirrVals = pick(s => s.objectives.eirr);
        const lcohVals = pick(s => s.objectives.lcoh);
        const curtVals = pick(s => s.objectives.curtailmentRate);

        function range(vals) {
            if (!vals.length) return null;
            let mn = Infinity, mx = -Infinity;
            for (const v of vals) { if (v < mn) mn = v; if (v > mx) mx = v; }
            if (!(mx - mn > 1e-12)) return null;
            return { min: mn, span: mx - mn };
        }

        const rEirr = range(eirrVals);
        const rLcoh = range(lcohVals);
        const rCurt = range(curtVals);

        return solutions.map(s => {
            const eirr = s.objectives.eirr;
            const lcoh = s.objectives.lcoh;
            const curt = s.objectives.curtailmentRate;

            // EIRR 越高越好；LCOH / 弃电率 越低越好
            // 目标值全部相同（range 为 null）时该维度视为满分，不影响排序
            let eirrScore;
            if (eirr === null || !isFinite(eirr)) eirrScore = 0;          // 经济不可行 → 最差
            else if (rEirr) eirrScore = (eirr - rEirr.min) / rEirr.span;
            else eirrScore = 1;

            let lcohScore;
            if (!isFinite(lcoh)) lcohScore = 0;
            else if (rLcoh) lcohScore = 1 - (lcoh - rLcoh.min) / rLcoh.span;
            else lcohScore = 1;

            let curtScore;
            if (!isFinite(curt)) curtScore = 0;
            else if (rCurt) curtScore = 1 - (curt - rCurt.min) / rCurt.span;
            else curtScore = 1;

            return {
                key: s.key,
                eirrScore: eirrScore,
                lcohScore: lcohScore,
                curtailmentScore: curtScore,
            };
        });
    }

    /**
     * 从 Pareto 前沿中挑选综合推荐方案。
     * 注意：权重只在此处使用，绝不进入 NSGA-II 的适应度。
     */
    function calculateRecommendedSolution(paretoSolutions, config) {
        if (!paretoSolutions || paretoSolutions.length === 0) return null;

        const scores = normalizeObjectives(paretoSolutions);
        const w = config.recommendationWeights;
        let wSum = (Number(w.eirr) || 0) + (Number(w.lcoh) || 0) + (Number(w.curtailmentRate) || 0);
        if (!(wSum > 0)) wSum = 1;
        const we = (Number(w.eirr) || 0) / wSum;
        const wl = (Number(w.lcoh) || 0) / wSum;
        const wc = (Number(w.curtailmentRate) || 0) / wSum;

        let best = null, bestScore = -Infinity;
        for (let i = 0; i < paretoSolutions.length; i++) {
            const s = paretoSolutions[i];
            const sc = scores[i];
            const total = we * sc.eirrScore + wl * sc.lcohScore + wc * sc.curtailmentScore;
            s.scores = sc;
            s.score = total;

            // 平局处理：投资较低者优先，其次制氢量较高者优先
            if (total > bestScore + 1e-12) { bestScore = total; best = s; }
            else if (Math.abs(total - bestScore) <= 1e-12 && best) {
                const invA = s.economic.totalInvestment, invB = best.economic.totalInvestment;
                if (invA < invB - 1e-9) { best = s; }
                else if (Math.abs(invA - invB) <= 1e-9 &&
                         s.technical.annualHydrogenKg > best.technical.annualHydrogenKg) { best = s; }
            }
        }
        return best;
    }

    /** 识别四类代表方案：经济最优 / 氢成本最优 / 消纳最优 / 综合推荐 */
    function pickRepresentativeSolutions(paretoSolutions, config) {
        if (!paretoSolutions || paretoSolutions.length === 0) {
            return { economicBest: null, hydrogenCostBest: null, curtailmentBest: null, recommended: null };
        }

        /** 平局时：投资较低 → 制氢量较高 */
        function tieBreak(cur, cand) {
            if (!cur) return cand;
            const invA = cand.economic.totalInvestment, invB = cur.economic.totalInvestment;
            if (invA < invB - 1e-9) return cand;
            if (Math.abs(invA - invB) <= 1e-9 &&
                cand.technical.annualHydrogenKg > cur.technical.annualHydrogenKg) return cand;
            return cur;
        }

        let economicBest = null, hydrogenCostBest = null, curtailmentBest = null;

        for (const s of paretoSolutions) {
            // 方案A：资本金内部收益率（EIRR）最高
            if (s.economic.EIRR !== null && isFinite(s.economic.EIRR)) {
                if (!economicBest || s.economic.EIRR > economicBest.economic.EIRR + 1e-12) economicBest = s;
                else if (Math.abs(s.economic.EIRR - economicBest.economic.EIRR) <= 1e-12) economicBest = tieBreak(economicBest, s);
            }
            // 方案B：LCOH 最低
            if (isFinite(s.economic.LCOH)) {
                if (!hydrogenCostBest || s.economic.LCOH < hydrogenCostBest.economic.LCOH - 1e-9) hydrogenCostBest = s;
                else if (Math.abs(s.economic.LCOH - hydrogenCostBest.economic.LCOH) <= 1e-9) hydrogenCostBest = tieBreak(hydrogenCostBest, s);
            }
            // 方案C：弃电率最低
            if (!curtailmentBest || s.technical.curtailmentRate < curtailmentBest.technical.curtailmentRate - 1e-9) curtailmentBest = s;
            else if (Math.abs(s.technical.curtailmentRate - curtailmentBest.technical.curtailmentRate) <= 1e-9) curtailmentBest = tieBreak(curtailmentBest, s);
        }

        const recommended = calculateRecommendedSolution(paretoSolutions, config);

        return {
            economicBest: economicBest,
            hydrogenCostBest: hydrogenCostBest,
            curtailmentBest: curtailmentBest,
            recommended: recommended,
        };
    }

    // ======================================================================
    // 九、优化会话（可分代驱动 → 支持 Worker 与主线程分片调度）
    // ======================================================================

    /** 校验运行上下文，异常抛出可读错误（前端捕获后提示，不崩溃页面） */
    function validateContext(context) {
        if (!context) throw new Error('缺少优化上下文 context');
        normalizeContext(context);
        if (!context.pvData || !context.windData) {
            throw new Error('优化上下文缺少风光 8760 小时数据（pvData / windData），请先在「电量计算」页加载 input.xlsx');
        }
        const n = context.pvData.length;
        if (!(n > 0)) throw new Error('风光 8760 小时数据为空（长度为 0）');
        if (n !== context.windData.length) throw new Error('光伏与风电数据长度不一致');
        if (n !== 8760) {
            // 不视为致命错误，但要显式提示（数据可能为闰年或非典型年）
            if (typeof context.onWarn === 'function') {
                context.onWarn('风光数据长度为 ' + n + ' 小时（标准应为 8760），结果仅供参考');
            }
        }
        if (!context.simulationConfig) {
            throw new Error('缺少系统运行参数（simulationConfig），请检查参数传递链路');
        }
        if (!context.prices) throw new Error('缺少概算设备单价参数');
        if (!context.financeParams) throw new Error('缺少财务评价参数');
    }

    /** 校验配置，返回 { ok, message } */
    function validateConfig(config) {
        const problems = [];
        for (const key of KEY_ORDER) {
            const v = config.variables[key];
            const label = VAR_META[key].label;
            const min = Number(v.min), max = Number(v.max), step = Number(v.step);
            if (!isFinite(min) || !isFinite(max)) problems.push(label + '：最小值/最大值必须为数字');
            else if (min > max) problems.push(label + '：最小值不能大于最大值');
            else if (min < 0) problems.push(label + '：最小值不能为负数');
            if (!isFinite(step) || step < 0) problems.push(label + '：步长必须 >= 0');
        }
        const ns = config.nsga2;
        if (!(ns.populationSize >= 2)) problems.push('种群规模过小（至少 2）');
        if (!(ns.generations >= 1)) problems.push('迭代次数过小（至少 1）');
        if (!(ns.crossoverProbability >= 0 && ns.crossoverProbability <= 1)) problems.push('交叉概率必须位于 0~1');
        if (!(ns.mutationProbability >= 0 && ns.mutationProbability <= 1)) problems.push('变异概率必须位于 0~1');

        const c = config.constraints;
        if (c.minGreenHydrogenRatio && c.minGreenHydrogenRatio.enabled &&
            Number(c.minGreenHydrogenRatio.value) > 100) problems.push('绿电制氢比例不能大于 100%');
        if (c.maxGridImportRatio && c.maxGridImportRatio.enabled &&
            Number(c.maxGridImportRatio.value) > 100) problems.push('外购电比例不能大于 100%');

        return { ok: problems.length === 0, message: problems.join('；') };
    }

    /**
     * 创建优化会话
     * @param {Object} options { config, context, onProgress, onGeneration }
     * @returns {Object} 会话对象 { runNextGeneration, isFinished, cancel, isCancelled, getResult, getLevels }
     */
    function createSession(options) {
        const opts = options || {};
        const config = normalizeConfig(opts.config);
        const ctx = Object.assign({}, opts.context);
        const onProgress = opts.onProgress || function () {};
        const onGeneration = opts.onGeneration || function () {};

        // 上下文补全
        ctx.config = config;
        ctx.lcohDiscountRate = Number(config.lcoh.discountRate);
        ctx.cache = ctx.cache || new Map();
        // V2.3.1 修复：保留调用方传入的 stats 对象。
        // 旧实现无条件重建 ctx.stats，导致调用方（Worker 的 performanceStats、
        // 基准脚本）累计的 evalTimeMs / simTimeMs 永远读不到（始终为 0）。
        ctx.stats = (opts.context && opts.context.stats) ? opts.context.stats
            : { evaluated: 0, cacheHits: 0, evalTimeMs: 0, simTimeMs: 0 };
        if (ctx.stats.evalTimeMs === undefined) ctx.stats.evalTimeMs = 0;
        if (ctx.stats.simTimeMs === undefined) ctx.stats.simTimeMs = 0;
        ctx.rng = createRng(Number(config.nsga2.randomSeed) || 20260912);

        validateContext(ctx);
        const vres = validateConfig(config);
        if (!vres.ok) throw new Error('优化参数非法：' + vres.message);

        const startTime = nowMs();

        let initialized = false;
        let finished = false;
        let cancelled = false;
        let generation = 0;
        let population = [];
        let searchSpace = 0;
        let effectivePopulation = 0;
        let levels = null;
        let stagnantGenerations = 0;
        let bestSnapshot = null;
        const history = [];

        /** 对一组个体做评价 */
        function evaluateIndividuals(list) {
            for (const ind of list) {
                if (ind.evaluation) continue;
                bindEvaluation(ind, ctx);
            }
        }

        /** 当前种群统计快照 */
        function snapshot(extra) {
            const feasible = population.filter(ind => ind.feasible);
            const pool = feasible.length > 0 ? feasible : population;

            let bestEirr = null, bestLCOH = Infinity, bestCurt = Infinity;
            for (const ind of pool) {
                const ev = ind.evaluation;
                if (!ev) continue;
                const e = ev.objectives.eirr;
                if (e !== null && isFinite(e) && (bestEirr === null || e > bestEirr)) bestEirr = e;
                if (isFinite(ev.objectives.lcoh) && ev.objectives.lcoh < bestLCOH) bestLCOH = ev.objectives.lcoh;
                if (isFinite(ev.objectives.curtailmentRate) && ev.objectives.curtailmentRate < bestCurt) bestCurt = ev.objectives.curtailmentRate;
            }

            const fronts = nonDominatedSort(population);
            const paretoCount = (fronts[0] || []).filter(i => i.feasible).length;

            return Object.assign({
                generation: generation,
                totalGenerations: Number(config.nsga2.generations),
                evaluated: ctx.stats.evaluated,
                cacheHits: ctx.stats.cacheHits,
                feasibleCount: feasible.length,
                paretoCount: paretoCount,
                bestEirr: bestEirr,
                bestLCOH: isFinite(bestLCOH) ? bestLCOH : null,
                bestCurtailmentRate: isFinite(bestCurt) ? bestCurt : null,
                effectivePopulationSize: effectivePopulation,
                searchSpaceSize: searchSpace,
                elapsedTime: (nowMs() - startTime) / 1000,
                done: finished,
                cancelled: cancelled,
            }, extra || {});
        }

        /** 记录一代历史（用于收敛曲线与「优化过程」导出） */
        function recordHistory(snap) {
            history.push({
                generation: snap.generation,
                evaluatedCount: snap.evaluated,
                feasibleCount: snap.feasibleCount,
                paretoCount: snap.paretoCount,
                bestEirr: snap.bestEirr,
                bestLCOH: snap.bestLCOH,
                bestCurtailmentRate: snap.bestCurtailmentRate,
            });
        }

        /** 提前终止判据：Pareto 前沿综合指标是否仍有明显改善 */
        function checkImprovement(snap) {
            if (!bestSnapshot) { bestSnapshot = snap; stagnantGenerations = 0; return; }
            const betterEirr = (snap.bestEirr !== null && isFinite(snap.bestEirr)) &&
                (bestSnapshot.bestEirr === null || snap.bestEirr > bestSnapshot.bestEirr + 1e-4);
            const betterLcoh = (snap.bestLCOH !== null) && (bestSnapshot.bestLCOH === null ||
                snap.bestLCOH < bestSnapshot.bestLCOH * (1 - 1e-4));
            const betterCurt = (snap.bestCurtailmentRate !== null) && (bestSnapshot.bestCurtailmentRate === null ||
                snap.bestCurtailmentRate < bestSnapshot.bestCurtailmentRate * (1 - 1e-4));

            if (betterEirr || betterLcoh || betterCurt) {
                stagnantGenerations = 0;
                bestSnapshot = snap;
            } else {
                stagnantGenerations++;
            }
        }

        /** 第一步：初始化种群并完成首次评价 */
        function ensureInitialized() {
            if (initialized) return;
            const init = initializePopulation(config, ctx);
            population = init.population;
            searchSpace = init.searchSpaceSize;
            effectivePopulation = init.effectivePopulationSize;
            levels = init.levels;

            // 有效种群规模回写（保持与搜索空间匹配）
            config.nsga2.populationSize = effectivePopulation;

            evaluateIndividuals(population);
            refreshRankAndCrowding(population);

            initialized = true;
            generation = 0;

            const snap = snapshot();
            recordHistory(snap);
            bestSnapshot = snap;
            onGeneration(snap);
            onProgress(snap);
        }

        /** 推进一步：跑完一代（选择 → 交叉 → 变异 → 评价 → 精英保留） */
        function runNextGeneration() {
            if (finished || cancelled) return;
            if (!initialized) { ensureInitialized(); return; }

            generation++;

            // 1. 繁殖：SBX 交叉 + 离散步长变异
            const offspring = [];
            const offspringKeys = new Set();
            const targetSize = effectivePopulation;
            let guard = 0;
            while (offspring.length < targetSize && guard < targetSize * 20) {
                guard++;
                const p1 = tournamentSelection(population, ctx.rng);
                const p2 = tournamentSelection(population, ctx.rng);
                const [g1, g2] = sbxCrossover(p1, p2, config, ctx.rng);

                const m1 = mutation(g1, config, ctx.rng);
                const m2 = mutation(g2, config, ctx.rng);

                for (const genes of [m1, m2]) {
                    if (offspring.length >= targetSize) break;
                    const ind = makeIndividual(genes, null);
                    if (offspringKeys.has(ind.key)) continue;   // 代内去重，减少重复评价
                    offspringKeys.add(ind.key);
                    offspring.push(ind);
                }
            }

            // 2. 评价子代
            evaluateIndividuals(offspring);

            // 3. 精英保留生成下一代，并重算 rank / 拥挤距离
            population = createNextGeneration(population, offspring, config);
            refreshRankAndCrowding(population);

            // 4. 统计 / 历史 / 回调
            const snap = snapshot();
            recordHistory(snap);
            checkImprovement(snap);
            onGeneration(snap);
            onProgress(snap);

            // 5. 终止判据
            if (generation >= Number(config.nsga2.generations)) {
                finished = true;
            } else if (config.nsga2.earlyStopping && stagnantGenerations >= Number(config.nsga2.patience)) {
                finished = true;
                snap.earlyStopped = true;
            }

            if (finished) {
                const finalSnap = snapshot();
                onProgress(finalSnap);
            }
        }

        /** 汇总最终结果 */
        function getResult() {
            const fronts = nonDominatedSort(population);
            const paretoFront = (fronts[0] || []).filter(ind => ind.feasible);

            // 组装 Pareto 方案列表（浅拷贝，避免污染缓存对象）
            const paretoSolutions = paretoFront.map((ind, i) => {
                const ev = ind.evaluation;
                return Object.assign({}, ev, {
                    id: 'P-' + String(i + 1).padStart(3, '0'),
                    rank: 1,
                    score: null,
                    scores: null,
                });
            });

            // 综合推荐 + 四类代表方案
            const representative = pickRepresentativeSolutions(paretoSolutions, config);

            // 基准方案（可选）
            let baseline = null;
            if (config.baseline) {
                try {
                    baseline = evaluateScheme(config.baseline, ctx);
                } catch (e) {
                    baseline = null;
                }
            }

            const feasibleCount = population.filter(ind => ind.feasible).length;

            return {
                success: paretoSolutions.length > 0,
                config: config,
                statistics: {
                    totalEvaluated: ctx.stats.evaluated,
                    cacheHits: ctx.stats.cacheHits,
                    feasibleCount: feasibleCount,
                    paretoCount: paretoSolutions.length,
                    generationsCompleted: generation,
                    elapsedTime: (nowMs() - startTime) / 1000,
                    searchSpaceSize: searchSpace,
                    effectivePopulationSize: effectivePopulation,
                    earlyStopped: (config.nsga2.earlyStopping && stagnantGenerations >= Number(config.nsga2.patience)),
                },
                paretoSolutions: paretoSolutions,
                representativeSolutions: representative,
                baseline: baseline,
                history: history,
                levels: levels,
            };
        }

        /** 实时获取当前 Pareto 前沿（供进度刷新时预览） */
        function getCurrentPareto() {
            const fronts = nonDominatedSort(population);
            return (fronts[0] || []).filter(ind => ind.feasible)
                .map(ind => ind.evaluation);
        }

        return {
            runNextGeneration: runNextGeneration,
            ensureInitialized: ensureInitialized,
            isFinished: function () { return finished; },
            isCancelled: function () { return cancelled; },
            isInitialized: function () { return initialized; },
            cancel: function () { cancelled = true; },
            getResult: getResult,
            getCurrentPareto: getCurrentPareto,
            getContext: function () { return ctx; },
            getProgressSnapshot: function () { return snapshot(); },
            getLevels: function () { return levels; },
        };
    }

    /**
     * 同步运行（Worker 内部使用）：一次跑完整个优化过程
     */
    function run(options) {
        const session = createSession(options);
        let guard = 0;
        const maxSteps = (Number(options.config && options.config.nsga2 && options.config.nsga2.generations) || 100) + 5;
        session.ensureInitialized();
        while (!session.isFinished() && !session.isCancelled() && guard++ < maxSteps) {
            session.runNextGeneration();
        }
        return session.getResult();
    }

    // ======================================================================
    // 十、导出
    // ======================================================================

    const OptimizationEngine = {
        VERSION: VERSION,
        KEY_ORDER: KEY_ORDER,
        VAR_META: VAR_META,

        // 配置
        buildDefaultConfig: buildDefaultConfig,
        normalizeConfig: normalizeConfig,
        validateConfig: validateConfig,

        // 数值工具
        snapToStep: snapToStep,
        buildLevels: buildLevels,
        buildLevelsFor: buildLevelsFor,
        searchSpaceSize: searchSpaceSize,
        createRng: createRng,

        // 统一评价
        evaluateScheme: evaluateScheme,
        calculateLCOH: calculateLCOH,
        schemeKey: schemeKey,
        checkConstraints: checkConstraints,

        // NSGA-II 算子
        dominates: dominates,
        nonDominatedSort: nonDominatedSort,
        calculateCrowdingDistance: calculateCrowdingDistance,
        tournamentSelection: tournamentSelection,
        sbxCrossover: sbxCrossover,
        mutation: mutation,
        createNextGeneration: createNextGeneration,
        getParetoFront: getParetoFront,

        // 初始化与推荐
        createIndividual: createIndividual,
        initializePopulation: initializePopulation,
        normalizeObjectives: normalizeObjectives,
        calculateRecommendedSolution: calculateRecommendedSolution,
        pickRepresentativeSolutions: pickRepresentativeSolutions,

        // 调度
        createSession: createSession,
        run: run,
    };

    const g = rt();
    g.OptimizationEngine = OptimizationEngine;

    // 兼容 CommonJS（Node 回归测试用），浏览器 / Worker 环境不触发
    if (typeof module !== 'undefined' && module.exports) module.exports = OptimizationEngine;

})(typeof self !== 'undefined' ? self : this);
