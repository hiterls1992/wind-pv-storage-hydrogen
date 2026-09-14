/**
 * ============================================================================
 * 结果数据层  ——  V2.3 前端性能与数据流优化
 * ============================================================================
 *
 * 任务书 §4 / §5 / §8 / §9 / §18 / §19 / §26
 *
 * 设计目标：
 *   「计算一次、统一存储、按需展示」—— 计算层 / 结果层 / 显示层彻底分离。
 *
 *   计算层   Float64Array(8760 × 11)       ← simulation-engine，保持不动
 *   结果层   ResultDataStore / 视图 / 缓存  ← 本文件
 *   显示层   图表 / 表格 / Excel           ← 按需读取，不持有第二份 8760 数据
 *
 * ---------------------------------------------------------------------------
 * 三个必须遵守的铁律
 * ---------------------------------------------------------------------------
 *   ① **显示降采样绝不回流到计算**。
 *      ChartDataAdapter 的降采样结果只允许用于绘图；任何技术指标、经济指标、
 *      约束判断都必须走 ResultDataStore 全分辨率读取（getAnnualSummary /
 *      getMonthlySummary / getColumn）。
 *
 *   ② **8760 小时数据只保留一份**（Float64Array）。禁止为了显示把整段结果
 *      永久转成 Array<Object>（8760 个对象 × 11 属性）。
 *      视图（HourView）用「复用同一个行对象」的方式遍历，分配量 O(1) 而不是 O(8760)。
 *
 *   ③ **优化阶段不保存 8760 原始结果**。见 OptimizationResultCache 的说明。
 *
 * ---------------------------------------------------------------------------
 * 运行环境
 * ---------------------------------------------------------------------------
 * 本文件同时被主线程与 Web Worker 加载（Worker 内 importScripts），
 * 因此**模块加载期不得触碰 document / window**。
 */

(function () {
    'use strict';

    // =========================================================================
    // 一、列定义（与 simulation-engine.js 的输出列顺序严格一致，不可调整）
    // =========================================================================

    /** 数值列索引（供计算与适配器使用） */
    const COLS = {
        pv: 0,
        wind: 1,
        total: 2,
        charge: 3,
        discharge: 4,
        storage: 5,
        hydrogenEnergy: 6,
        export: 7,
        import: 8,
        curtailment: 9,
        hydrogen: 10,
    };

    const COL_COUNT = 11;

    /** 列顺序 → 中文列名（与 V1.0 的 Excel 表头、图表系列名保持一致） */
    const COL_LABELS = [
        '光伏电量', '风电电量', '合计电量',
        '储能充电量', '储能放电量', '储能现存容量',
        '制氢电量', '上网电量', '下网电量',
        '弃电量', '制氢量',
    ];

    /** 图表列 key ↔ 中文列名（与 chart-module.js 的 COL_NAMES 一致） */
    const KEY_TO_LABEL = {
        pv: '光伏电量', wind: '风电电量', total: '合计电量',
        charge: '储能充电量', discharge: '储能放电量', storage: '储能现存容量',
        hydrogen: '制氢电量', export: '上网电量', import: '下网电量',
        curtailment: '弃电量', h2prod: '制氢量',
    };
    const LABEL_TO_KEY = {};
    for (const k of Object.keys(KEY_TO_LABEL)) LABEL_TO_KEY[KEY_TO_LABEL[k]] = k;

    /** 每月天数（非闰年，与 simulation-engine 的 8760 小时口径一致） */
    const MONTH_DAYS = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

    // =========================================================================
    // 二、工具
    // =========================================================================

    /** 把列名或列 key 统一解析为列索引；非法返回 -1 */
    function resolveColumn(column) {
        if (typeof column === 'number') return (column >= 0 && column < COL_COUNT) ? column : -1;
        if (typeof column === 'string') {
            if (Object.prototype.hasOwnProperty.call(COLS, column)) return COLS[column];
            const byLabel = COL_LABELS.indexOf(column);
            return byLabel;
        }
        return -1;
    }

    /** 结果对象 → 方案标识（优先用结果自带的 scheme，§48） */
    function schemeKeyOf(result) {
        if (!result) return '';
        if (result.key) return result.key;
        if (result.scheme && typeof ParameterManager !== 'undefined' && ParameterManager.schemeKey) {
            return ParameterManager.schemeKey(result.scheme);
        }
        return '';
    }

    // =========================================================================
    // 二.5 SimulationKey（V2.3.1 任务书 §9 / §10 / §36）
    // =========================================================================
    //
    //   SchemeKey      = 方案容量身份（建多大）
    //   SimulationKey  = Scheme + SimulationConfig + 输入数据版本 + 仿真算法版本
    //
    // 换气象数据 / 改制氢电耗 / 改储能效率 / 改上网约束 / 改仿真策略时，
    // 即使 Scheme 不变也会得到新的 SimulationKey —— 旧缓存因此不会污染新结果。

    /** 仿真算法版本（任务书 §36）：仿真公式发生变化时必须递增，使旧缓存自动失效 */
    const SIMULATION_ENGINE_VERSION = '2.3.1';

    /** FNV-1a 字符串哈希（用于 config 指纹；不扫描 8760 数据） */
    function fnv1a(str) {
        let h = 0x811c9dc5;
        for (let i = 0; i < str.length; i++) {
            h ^= str.charCodeAt(i);
            h = Math.imul(h, 0x01000193) >>> 0;
        }
        return ('0000000' + h.toString(16)).slice(-8);
    }

    /** 运行参数指纹（只哈希配置对象，代价 O(1)） */
    function configFingerprint(cfg) {
        if (!cfg) return 'none';
        try { return fnv1a(JSON.stringify(cfg)); } catch (e) { return 'err'; }
    }

    /**
     * 生成仿真键（任务书 §9）。
     *
     * @param {Object} opts { scheme, simulationConfig, inputVersion, schemeKey, engineVersion }
     * @returns {string} 形如 "2.3.1|W200|PV360|B100|H2|EL160|cfg1a2b3c4d|input5f6e7d8c-8760"
     */
    function createSimulationKey(opts) {
        const o = opts || {};
        const schemePart = o.schemeKey ||
            ((o.scheme && typeof ParameterManager !== 'undefined' && ParameterManager.schemeKey)
                ? ParameterManager.schemeKey(o.scheme) : 'noscheme');
        const cfgPart = configFingerprint(o.simulationConfig);
        const inputPart = (o.inputVersion === undefined || o.inputVersion === null)
            ? 'input-none' : String(o.inputVersion);
        const enginePart = o.engineVersion || SIMULATION_ENGINE_VERSION;
        return enginePart + '|' + schemePart + '|cfg' + cfgPart + '|' + inputPart;
    }

    /** 结果对象 → 仿真键（优先取结果自带的 simulationKey，缺则现场合成） */
    function simulationKeyOf(result) {
        if (!result) return '';
        if (result.simulationKey) return result.simulationKey;
        return createSimulationKey({
            scheme: result.scheme,
            simulationConfig: result.simulationConfig,
            schemeKey: result.key,
        });
    }

    /**
     * 生成输入数据版本号（任务书 §9 / §35）。
     *
     * 在**输入数据加载时调用一次**并复用，严禁在每次生成 simulationKey 时重新哈希 8760 数据。
     * 用 FNV-1a 逐元素混合浮点位模式，能可靠区分"数据真的变了"与"仅换个文件名"。
     *
     * @param {Array<{pv:number,wind:number}>} rows 解析后的逐时数据
     * @param {Object} [meta] 附加元信息（文件名 / 大小等）
     * @returns {string} 形如 "in-1a2b3c4d-8760"
     */
    function createInputVersionFromRows(rows, meta) {
        const f = new Float64Array(1);
        const u = new Uint32Array(f.buffer);
        let h = 0x811c9dc5;
        const mixU32 = function (v) {
            h ^= v >>> 0;
            h = Math.imul(h, 0x01000193) >>> 0;
        };
        const n = rows ? rows.length : 0;
        for (let i = 0; i < n; i++) {
            f[0] = rows[i].pv; mixU32(u[0]); mixU32(u[1]);
            f[0] = rows[i].wind; mixU32(u[0]); mixU32(u[1]);
        }
        if (meta) {
            const mh = fnv1a(JSON.stringify(meta));
            for (let i = 0; i < mh.length; i++) mixU32(mh.charCodeAt(i));
        }
        return 'in-' + ('0000000' + h.toString(16)).slice(-8) + '-' + n;
    }

    // =========================================================================
    // 三、ResultDataStore：统一结果数据存储层（§4）
    // =========================================================================

    /**
     * 把仿真输出规范化成「结果记录」。
     * 核心数据 results 保持 Float64Array 原样引用 —— **不做任何拷贝**。
     *
     * @param {Float64Array} rawResults 长度 = hours × 11
     * @param {Object} scheme  方案参数（5 个容量，含派生 storageEnergy）
     * @param {Object} [meta]  附加元数据（simulationConfig / systemVars / ratioData / sums / filename）
     * @returns {Object} 结果记录
     */
    function create(rawResults, scheme, meta) {
        if (!rawResults || typeof rawResults.length !== 'number') {
            throw new Error('ResultDataStore.create：rawResults 必须是 Float64Array');
        }
        if (rawResults.length % COL_COUNT !== 0) {
            throw new Error('ResultDataStore.create：rawResults 长度必须是 ' + COL_COUNT + ' 的整数倍（当前 ' + rawResults.length + '）');
        }
        const m = meta || {};
        const rec = {
            /** ★ 唯一的一份 8760 小时核心数据（不拷贝、不对象化） */
            results: rawResults,
            hours: rawResults.length / COL_COUNT,
            scheme: scheme ? Object.assign({}, scheme) : null,
            simulationConfig: m.simulationConfig || null,
            systemVars: m.systemVars || null,
            ratioData: m.ratioData || null,
            sums: m.sums || null,
            filename: m.filename || null,
            /** 惰性缓存：年度/月度摘要只算一次 */
            _annual: null,
            _monthly: null,
            _monthlyByLabel: null,
        };
        rec.key = m.key || schemeKeyOf(rec);
        // V2.3.1（任务书 §9）：结果记录自带 simulationKey，
        // 供图表缓存与结果缓存正确区分"同方案不同运行参数/不同输入数据"
        rec.simulationKey = m.simulationKey || createSimulationKey({
            scheme: rec.scheme,
            simulationConfig: rec.simulationConfig,
            schemeKey: rec.key,
        });
        return rec;
    }

    /** 小时数 */
    function getLength(result) {
        return result ? result.hours : 0;
    }

    /** 读第 hour 行第 column 列（按需读取，不产生行对象） */
    function getValue(result, hour, column) {
        const c = resolveColumn(column);
        if (!result || c < 0 || hour < 0 || hour >= result.hours) return 0;
        return result.results[hour * COL_COUNT + c];
    }

    /** 取第 hour 行的对象（单行按需，供 tooltip / 详情使用） */
    function getHourRow(result, hour) {
        const row = { 小时: hour };
        if (!result || hour < 0 || hour >= result.hours) return row;
        const base = hour * COL_COUNT;
        for (let c = 0; c < COL_COUNT; c++) row[COL_LABELS[c]] = result.results[base + c];
        return row;
    }

    /** 取 [startHour, endHour) 区间的行对象数组（按需，供导出/详情使用） */
    function getRange(result, startHour, endHour) {
        const out = [];
        if (!result) return out;
        const s = Math.max(0, startHour | 0);
        const e = Math.min(result.hours, endHour | 0);
        for (let h = s; h < e; h++) out.push(getHourRow(result, h));
        return out;
    }

    /**
     * 取整列的**连续副本**。
     *
     * ⚠ 任务书 §3 指正：本函数【不是零拷贝】。
     *   核心数据是 hour-major 布局（每小时 11 列连续存放），单列数据在内存中并不连续，
     *   无法用 subarray() 取视图，因此这里必然创建一个长度 = hours 的临时 Float64Array
     *   （8760 × 8B ≈ 70 KB）。
     *
     * 使用准则：
     *   · 只在"确实需要连续数组"的场景调用（例如降采样定位极值）；
     *   · 只需逐点读取时，请改用 getValue() 或 createColumnView()（零分配）；
     *   · 调用方不得长期持有返回值。
     */
    function getColumnArray(result, column) {
        const c = resolveColumn(column);
        if (!result || c < 0) return new Float64Array(0);
        const n = result.hours;
        const raw = result.results;
        const out = new Float64Array(n);
        for (let h = 0; h < n; h++) out[h] = raw[h * COL_COUNT + c];
        return out;
    }

    /**
     * @deprecated V2.4 移除 —— 名称暗示"取列"却伴随整列复制，易被误读为零拷贝。
     * 请改用语义明确的 getColumnArray()（开销相同）。
     * TODO V2.4 REMOVE LEGACY
     */
    function getColumn(result, column) {
        return getColumnArray(result, column);
    }

    /**
     * 年度摘要（11 列合计 + 若干派生量）。**全分辨率遍历，只做一次并缓存**。
     * 这是显示层唯一被允许用来获取"总量"的入口（§18）。
     */
    function getAnnualSummary(result) {
        if (!result) return null;
        if (result._annual) return result._annual;

        const n = result.hours;
        const raw = result.results;
        const colSum = new Float64Array(COL_COUNT);
        let maxExport = 0, maxImport = 0, lastStorage = 0;

        for (let h = 0; h < n; h++) {
            const base = h * COL_COUNT;
            for (let c = 0; c < COL_COUNT; c++) colSum[c] += raw[base + c];
            if (raw[base + 7] > maxExport) maxExport = raw[base + 7];
            if (raw[base + 8] > maxImport) maxImport = raw[base + 8];
            lastStorage = raw[base + 5];
        }

        const byLabel = {};
        for (let c = 0; c < COL_COUNT; c++) byLabel[COL_LABELS[c]] = colSum[c];
        // 储能现存容量是「存量」而非「流量」，合计无意义 → 取末值
        byLabel['储能现存容量'] = lastStorage;

        result._annual = {
            hours: n,
            colSum: colSum,
            byLabel: byLabel,
            totalStorageLast: lastStorage,
            maxHourlyExport: maxExport,
            maxHourlyImport: maxImport,
        };
        return result._annual;
    }

    /** 月度累计起始小时（基于各月实际天数） */
    function monthCumulativeHours() {
        const cum = [0];
        for (let i = 0; i < 12; i++) cum.push(cum[i] + MONTH_DAYS[i] * 24);
        return cum;
    }

    /**
     * 月度摘要：12 个月 × 11 列合计。**全分辨率遍历，只做一次并缓存**。
     * 逐月柱状图 / 逐月汇总图 / 分模块逐月图统一从这里取数（§18 / §19）。
     */
    function getMonthlySummary(result) {
        if (!result) return null;
        if (result._monthly) return result._monthly;

        const cum = monthCumulativeHours();
        const n = result.hours;
        const raw = result.results;
        const months = [];
        for (let m = 0; m < 12; m++) {
            months.push({ colSum: new Float64Array(COL_COUNT), byLabel: {}, lastStorage: 0, hours: 0 });
        }

        for (let h = 0; h < n; h++) {
            // 定位月份（8760 小时恰好覆盖 12 个月；异常长度时归入最后一个有效月）
            let m = 11;
            for (let i = 0; i < 12; i++) {
                if (h < cum[i + 1]) { m = i; break; }
            }
            const base = h * COL_COUNT;
            const bucket = months[m];
            for (let c = 0; c < COL_COUNT; c++) bucket.colSum[c] += raw[base + c];
            bucket.lastStorage = raw[base + 5];
            bucket.hours++;
        }

        for (const bucket of months) {
            for (let c = 0; c < COL_COUNT; c++) bucket.byLabel[COL_LABELS[c]] = bucket.colSum[c];
            bucket.byLabel['储能现存容量'] = bucket.lastStorage;
        }

        // 按列名横向展开（便于图表按列名索引：[列名][月]）
        const byLabelMonthly = {};
        for (let c = 0; c < COL_COUNT; c++) {
            const label = COL_LABELS[c];
            byLabelMonthly[label] = months.map(bucket => bucket.byLabel[label]);
        }

        result._monthly = months;
        result._monthlyByLabel = byLabelMonthly;
        return months;
    }

    /** 取某列的各月合计（全分辨率口径） */
    function getMonthlyColumn(result, column) {
        getMonthlySummary(result);
        const c = resolveColumn(column);
        if (!result || c < 0 || !result._monthly) return [];
        return result._monthly.map(bucket => bucket.colSum[c]);
    }

    // =========================================================================
    // 四、HourView：零对象分配的小时视图（§5 的关键机制）
    // =========================================================================
    //
    // 图表渲染函数需要 `data.map(d => d['光伏电量'])` 这类写法。
    // 若先把 8760 行都对象化再 map，会一次性产生 8760 个对象（每个 11 个属性）
    // —— 这正是 V2.2 卡顿的主因之一。
    //
    // HourView 的做法：**复用同一个行对象**，每次只填充当前小时，
    // 并且只保留调用方真正要的那一列结果数组。分配量从 O(8760) 降到 O(1)。
    //
    // ⚠ 使用约束：map/forEach 传给回调的 row 是**复用对象**，
    //    回调内不得把 row 存入数组或跨迭代持有（取值后立即拷贝出来）。
    //    既有渲染函数全部满足该约束（只做取值/求和）。

    const KEYS = COL_LABELS;

    function createHourView(result, indices) {
        const raw = result ? result.results : null;
        const view = {
            /** 视图长度 */
            length: indices ? indices.length : getLength(result),
            /** 视图对应的真实小时序号（null 表示连续的 0..length-1） */
            indices: indices || null,
            result: result,
            /** 当前小时在 raw 中的基址（惰性行对象依赖它取值） */
            _base: 0,
            /** 该视图的轴标签间隔建议（供图表按点数自适应） */
            axisInterval: 0,
            /** 视图是否经过降采样（只用于显示，禁止参与指标计算） */
            downsampled: false,
            originalPoints: getLength(result),
        };

        const n = view.length;
        view.axisInterval = n > 0 ? Math.max(1, Math.floor(n / 8)) : 1;

        // V2.3.1（任务书 §4）：行对象改为【惰性 getter】。
        // 旧实现每小时把 11 列全部写入行对象 —— 一次 map 就是 8760 × 11 = 96360 次
        // 属性写入，而回调通常只读其中 1~2 列。改为 getter 后：
        //   · 每视图只定义 11 个 getter（O(1) 分配）；
        //   · 遍历中每小时只剩 1 次基址赋值；
        //   · 属性访问按需读 raw，读几列就付出几列的代价。
        // ⚠ 约束不变：row 仍是复用对象，回调内不得把它存入数组或跨迭代持有。
        const row = {};
        if (raw) {
            for (let c = 0; c < COL_COUNT; c++) {
                const ci = c;
                Object.defineProperty(row, KEYS[ci], {
                    get: function () { return raw[view._base + ci]; },
                    enumerable: true,
                    configurable: true,
                });
            }
        }
        view._row = row;

        /** 真实小时序号 */
        view.hourAt = function (i) {
            return indices ? indices[i] : i;
        };

        /** 遍历：fn(row, realHour, i)，row 为复用对象（惰性取值） */
        view.forEach = function (fn) {
            const r = view._row;
            const idx = indices;
            for (let i = 0; i < n; i++) {
                const h = idx ? idx[i] : i;
                view._base = h * COL_COUNT;
                fn(r, h, i);
            }
        };

        /** map：只保留 fn 的返回值（通常是数字），不保留 row */
        view.map = function (fn) {
            const out = new Array(n);
            const r = view._row;
            const idx = indices;
            for (let i = 0; i < n; i++) {
                const h = idx ? idx[i] : i;
                view._base = h * COL_COUNT;
                out[i] = fn(r, h, i);
            }
            return out;
        };

        /** reduce */
        view.reduce = function (fn, init) {
            let acc = init;
            const r = view._row;
            const idx = indices;
            for (let i = 0; i < n; i++) {
                const h = idx ? idx[i] : i;
                view._base = h * COL_COUNT;
                acc = fn(acc, r, h, i);
            }
            return acc;
        };

        /** 切片（保持视图语义，仍无对象分配） */
        view.slice = function (a, b) {
            const s = Math.max(0, a | 0);
            const e = Math.min(n, b === undefined ? n : b | 0);
            const src = indices;
            const sub = new Int32Array(Math.max(0, e - s));
            for (let i = s; i < e; i++) sub[i - s] = src ? src[i] : i;
            const child = createHourView(result, sub);
            child.downsampled = view.downsampled;
            return child;
        };

        /** 一次性取出某列的副本（Float64Array，会创建临时数组；仅确需连续数组时使用） */
        view.column = function (column) {
            const c = resolveColumn(column);
            const out = new Float64Array(n);
            if (c < 0 || !raw) return out;
            const idx = indices;
            for (let i = 0; i < n; i++) {
                const h = idx ? idx[i] : i;
                out[i] = raw[h * COL_COUNT + c];
            }
            return out;
        };

        return view;
    }

    /**
     * 列视图（任务书 §4 模式1）：单列直接访问，零分配。
     *
     * 与 HourView 的区别：HourView 服务于"逐小时多列"的既有渲染函数；
     * 当一个图表只需要**一列**时，用本视图可完全避免行对象与整列复制。
     */
    function createColumnView(result, column) {
        const c = resolveColumn(column);
        const raw = (result && result.results) ? result.results : null;
        const n = (result && c >= 0) ? result.hours : 0;
        const view = {
            length: n,
            column: c,
            result: result,
            /** 第 i 小时的值（零分配，直接读 hour-major 基址） */
            get: function (i) {
                return (i >= 0 && i < n) ? raw[i * COL_COUNT + c] : 0;
            },
            /** 映射为普通数组（仅当调用方确需数组时才创建） */
            map: function (fn) {
                const out = new Array(n);
                for (let i = 0; i < n; i++) out[i] = fn(raw[i * COL_COUNT + c], i);
                return out;
            },
            /** 求和（单次遍历，零分配） */
            sum: function () {
                let s = 0;
                for (let i = 0; i < n; i++) s += raw[i * COL_COUNT + c];
                return s;
            },
            /** 极值（单次遍历，零分配，返回 [min, max]） */
            extent: function () {
                let mn = Infinity, mx = -Infinity;
                for (let i = 0; i < n; i++) {
                    const v = raw[i * COL_COUNT + c];
                    if (v < mn) mn = v;
                    if (v > mx) mx = v;
                }
                return [mn, mx];
            },
        };
        return view;
    }

    // =========================================================================
    // 五、ChartDataAdapter：显示数据适配（§8 / §9）
    // =========================================================================

    /**
     * 降采样：分桶取「该桶内的最小值与最大值」，再按真实小时升序输出。
     *
     * 为什么不用等距抽样：电量/功率曲线存在尖峰（如光伏正午峰、电解槽启停），
     * 等距抽样会把尖峰整点漏掉，图形失真。分桶极值法保证**每个桶的极值都出现在图上**，
     * 同时把点数控制在 maxPoints 以内。
     *
     * ⚠ 本函数输出**只允许用于绘图**。任何指标计算必须走全分辨率接口。
     *
     * @param {Float64Array} values 全分辨率序列
     * @param {number} maxPoints 目标点数上限
     * @returns {{indices:Int32Array, points:number, downsampled:boolean}}
     */
    function downsampleIndices(values, maxPoints) {
        return downsampleByAccessor(function (i) { return values[i]; }, values.length, maxPoints);
    }

    /**
     * 降采样核心：从访问器取值，避免为定位极值而先复制整列（任务书 §3 / §5）。
     * @param {Function} get (i) => value
     * @param {number} n 数据长度
     * @param {number} maxPoints 目标点数上限
     */
    function downsampleByAccessor(get, n, maxPoints) {
        if (!(maxPoints > 0) || n <= maxPoints) {
            const all = new Int32Array(n);
            for (let i = 0; i < n; i++) all[i] = i;
            return { indices: all, points: n, downsampled: false };
        }

        const bucketCount = Math.max(1, Math.floor(maxPoints / 2));   // 每桶 2 点（min + max）
        const size = n / bucketCount;
        const picked = [];

        for (let b = 0; b < bucketCount; b++) {
            const s = Math.floor(b * size);
            const e = Math.min(n, Math.floor((b + 1) * size));
            if (e <= s) continue;
            let iMin = s, iMax = s;
            let vMin = get(s), vMax = get(s);
            for (let i = s + 1; i < e; i++) {
                const v = get(i);
                if (v < vMin) { vMin = v; iMin = i; }
                if (v > vMax) { vMax = v; iMax = i; }
            }
            if (iMin === iMax) picked.push(iMin);
            else if (iMin < iMax) picked.push(iMin, iMax);
            else picked.push(iMax, iMin);
        }

        return { indices: Int32Array.from(picked), points: picked.length, downsampled: true };
    }

    const DEFAULT_MAX_POINTS = 1500;

    /**
     * 图表数据缓存（任务书 §8）。
     * Key = simulationKey ‖ chartType ‖ column ‖ month ‖ maxPoints
     *
     * V2.3.1（任务书 §9 / §10）：Key 的第一段从 schemeKey 改为 **simulationKey**——
     * 同一方案在不同运行参数 / 不同输入数据下必须命中不同缓存，否则会读到旧结果。
     * 顶层分隔符用 '||'（simulationKey 内部只含单 '|'），便于按方案前缀精准清除。
     */
    const ChartDataCache = {
        _m: new Map(),
        LIMIT: 80,
        stats: { hits: 0, misses: 0 },

        key(simulationKey, chartType, column, month, maxPoints) {
            return [simulationKey || '', chartType || '', column === undefined || column === null ? '-' : column,
                month === undefined || month === null ? '-' : month,
                maxPoints === undefined || maxPoints === null ? '-' : maxPoints].join('||');
        },

        get(k) {
            if (this._m.has(k)) {
                this.stats.hits++;
                // 简单 LRU：命中后移到末尾
                const v = this._m.get(k);
                this._m.delete(k);
                this._m.set(k, v);
                return v;
            }
            this.stats.misses++;
            return undefined;
        },

        set(k, v) {
            if (this._m.has(k)) this._m.delete(k);
            this._m.set(k, v);
            // 超限时淘汰最旧的若干条（简单 LRU）
            while (this._m.size > this.LIMIT) {
                const oldest = this._m.keys().next().value;
                this._m.delete(oldest);
            }
            return v;
        },

        /** 按 simulationKey 前缀清除（供 SimulationResultCache 淘汰时联动，任务书 §33 / §34） */
        deleteBySimulationKey(simulationKey) {
            if (!simulationKey) return 0;
            let removed = 0;
            const prefix = simulationKey + '||';
            for (const k of Array.from(this._m.keys())) {
                if (k === simulationKey || k.indexOf(prefix) === 0) {
                    this._m.delete(k);
                    removed++;
                }
            }
            return removed;
        },

        clear() {
            this._m.clear();
            this.stats.hits = 0;
            this.stats.misses = 0;
        },

        get size() { return this._m.size; },

        /** 命中率（0~1） */
        get hitRate() {
            const total = this.stats.hits + this.stats.misses;
            return total === 0 ? 0 : this.stats.hits / total;
        },
    };

    const ChartDataAdapter = {
        DEFAULT_MAX_POINTS: DEFAULT_MAX_POINTS,

        /**
         * 取某列的逐时显示序列（必要时降采样）。
         *
         * @param {Object} result 结果记录
         * @param {string} column 列 key 或中文列名
         * @param {Object} [opts] { maxPoints }
         * @returns {{hours:number[], values:number[], realHours:number[], indices:Int32Array,
         *            downsampled:boolean, originalPoints:number, points:number, cached:boolean,
         *            axisInterval:number, column:string}}
         */
        getHourlySeries(result, column, opts) {
            const maxPoints = (opts && opts.maxPoints) || DEFAULT_MAX_POINTS;
            const label = KEY_TO_LABEL[column] || column;
            const sk = simulationKeyOf(result);
            const ck = ChartDataCache.key(sk, 'hourly', label, null, maxPoints);

            const cached = ChartDataCache.get(ck);
            if (cached) {
                return Object.assign({}, cached, { cached: true });
            }

            // V2.3.1（任务书 §3）：直接在 hour-major 原始数组上以步长访问定位极值，
            // 不再先复制整列（省去 70 KB 临时 Float64Array 与一次 8760 次拷贝）
            const raw = result.results;
            const cIdx = resolveColumn(label);
            const ds = downsampleByAccessor(function (h) { return raw[h * COL_COUNT + cIdx]; },
                getLength(result), maxPoints);

            const values = new Array(ds.points);
            const realHours = new Array(ds.points);
            for (let i = 0; i < ds.points; i++) {
                const h = ds.indices[i];
                realHours[i] = h;
                values[i] = raw[h * COL_COUNT + cIdx];
            }

            const built = {
                column: label,
                hours: realHours.slice(),          // 类别轴直接用真实小时号
                values: values,
                realHours: realHours,
                indices: ds.indices,
                downsampled: ds.downsampled,
                originalPoints: getLength(result),
                points: ds.points,
                axisInterval: Math.max(1, Math.floor(ds.points / 8)),
            };
            ChartDataCache.set(ck, built);
            return Object.assign({}, built, { cached: false });
        },

        /**
         * 取全分辨率的整段视图（供典型日/典型周/逐月等需要精确口径的图使用）。
         * 视图内不产生行对象数组。
         *
         * @param {Object} result
         * @returns {Object} HourView（全 8760，无降采样）
         */
        getFullView(result) {
            const sk = simulationKeyOf(result);
            const ck = ChartDataCache.key(sk, 'fullview', '-', null, null);
            const hit = ChartDataCache.get(ck);
            if (hit) return hit;
            return ChartDataCache.set(ck, createHourView(result, null));
        },

        /**
         * 单列逐时曲线的**降采样视图**（仅用于绘图）。
         * 分桶极值法：每个桶保留最小值与最大值所在小时，尖峰不会丢失。
         *
         * ⚠ 返回的视图只能用于绘图，禁止用于任何指标 / 约束计算。
         *
         * @param {Object} result
         * @param {string} column 列 key 或中文列名
         * @param {number} [maxPoints] 点数上限（默认 1500）
         */
        getHourlyView(result, column, maxPoints) {
            const mp = maxPoints || DEFAULT_MAX_POINTS;
            const label = KEY_TO_LABEL[column] || column;
            const sk = simulationKeyOf(result);
            const ck = ChartDataCache.key(sk, 'hourlyview', label, null, mp);
            const hit = ChartDataCache.get(ck);
            if (hit) return hit;

            const full = getColumn(result, label);
            const ds = downsampleIndices(full, mp);
            const view = createHourView(result, ds.indices);
            view.downsampled = ds.downsampled;
            view.originalPoints = full.length;
            view.axisInterval = Math.max(1, Math.floor(ds.points / 8));
            return ChartDataCache.set(ck, view);
        },

        /**
         * 全年能量流动汇总图的**降采样视图**（7 条曲线）。
         * 对 7 列同时做分桶极值，取各列极值小时的并集，
         * 桶数按 maxPoints / (2 × 曲线数) 预算，保证并集不超过 maxPoints。
         *
         * ⚠ 仅用于绘图。
         *
         * @param {Object} result
         * @param {number} [maxPoints]
         * @param {string[]} [columns] 参与降采样的列（默认 7 条能量流曲线）
         */
        getOverviewView(result, maxPoints, columns) {
            const mp = maxPoints || DEFAULT_MAX_POINTS;
            const cols = (columns || ['pv', 'wind', 'discharge', 'import', 'charge', 'hydrogen', 'export'])
                .map(c => KEY_TO_LABEL[c] || c);
            const sk = simulationKeyOf(result);
            const ck = ChartDataCache.key(sk, 'overviewview', cols.join('+'), null, mp);
            const hit = ChartDataCache.get(ck);
            if (hit) return hit;

            const n = getLength(result);
            if (n <= mp) return ChartDataCache.set(ck, createHourView(result, null));

            const raw = result.results;
            const colIdx = cols.map(label => resolveColumn(label));

            // V2.3.1（任务书 §5）：单次遍历同时提取 7 列的极值索引。
            // 旧实现先调 7 次 getColumnArray（7 × 70 KB 临时数组 + 7 次 8760 拷贝），
            // 再各自扫描一遍；现在零中间数组，一遍扫完。
            const bucketCount = Math.max(1, Math.floor(mp / (2 * Math.max(1, cols.length))));
            const size = n / bucketCount;
            const picked = new Set();

            for (let b = 0; b < bucketCount; b++) {
                const s = Math.floor(b * size);
                const e = Math.min(n, Math.floor((b + 1) * size));
                if (e <= s) continue;
                for (let ci = 0; ci < colIdx.length; ci++) {
                    const c = colIdx[ci];
                    let iMin = s, iMax = s;
                    let vMin = raw[s * COL_COUNT + c], vMax = vMin;
                    for (let h = s + 1; h < e; h++) {
                        const v = raw[h * COL_COUNT + c];
                        if (v < vMin) { vMin = v; iMin = h; }
                        if (v > vMax) { vMax = v; iMax = h; }
                    }
                    picked.add(iMin);
                    picked.add(iMax);
                }
            }

            const indices = Int32Array.from(Array.from(picked).sort((a, b) => a - b));
            const view = createHourView(result, indices);
            view.downsampled = true;
            view.originalPoints = n;
            view.axisInterval = Math.max(1, Math.floor(indices.length / 8));
            return ChartDataCache.set(ck, view);
        },

        /**
         * 取指定小时区间的视图（全分辨率，供典型日 24h / 典型周 168h / 单月使用）。
         * @param {Object} result
         * @param {number} startHour
         * @param {number} hourCount
         */
        getRangeView(result, startHour, hourCount) {
            const n = getLength(result);
            const s = Math.max(0, Math.min(n, startHour | 0));
            const e = Math.max(s, Math.min(n, s + (hourCount | 0)));
            const idx = new Int32Array(Math.max(0, e - s));
            for (let i = 0; i < idx.length; i++) idx[i] = s + i;
            return createHourView(result, idx);
        },

        /** 全分辨率月度视图（供逐月图使用；内部只读 MonthSummary 缓存） */
        getMonthlyView(result) {
            const sk = simulationKeyOf(result);
            const ck = ChartDataCache.key(sk, 'monthview', '-', null, null);
            const hit = ChartDataCache.get(ck);
            if (hit) return hit;

            getMonthlySummary(result);   // 确保缓存就绪
            const view = {
                length: 12,
                result: result,
                byLabel: result._monthlyByLabel,
                buckets: result._monthly,
                /** 取某列 12 个月合计 */
                column(label) { return result._monthlyByLabel[label] || []; },
                monthIndexAt: i => i,
            };
            return ChartDataCache.set(ck, view);
        },
    };

    // =========================================================================
    // 六、结果注册表（§19：SimulationResultCache / OptimizationResultCache 分离）
    // =========================================================================

    /**
     * SimulationResultCache —— **只保存用户实际查看过的方案**的完整 8760 结果。
     *
     * 与 OptimizationResultCache 分开，是因为二者的容量特征完全不同：
     *   · 用户查看的方案：个位数，保留完整 8760 是有价值的（反复切图表要复用）；
     *   · 优化候选方案：成百上千，保存 8760 会迅速耗尽内存。
     */
    const SimulationResultCache = {
        _m: new Map(),
        LIMIT: 8,          // 最多保留 8 个方案的完整逐时结果（任务书 §12：8~10 个 + LRU）

        /**
         * 登记用户查看过的结果。
         * LRU 淘汰时联动清除该方案的图表缓存，避免已释放的结果仍被视图引用（任务书 §33）。
         */
        put(result) {
            const k = simulationKeyOf(result);
            if (!k) return result;
            if (this._m.has(k)) this._m.delete(k);
            this._m.set(k, result);
            while (this._m.size > this.LIMIT) {
                const oldest = this._m.keys().next().value;
                this._m.delete(oldest);
                // 被淘汰的结果不再驻留 → 其图表显示数据一并释放，防止内存缓慢增长
                ChartDataCache.deleteBySimulationKey(oldest);
            }
            return result;
        },

        get(schemeKey) {
            return this._m.get(schemeKey);
        },

        has(schemeKey) {
            return this._m.has(schemeKey);
        },

        keys() {
            return Array.from(this._m.keys());
        },

        clear() {
            this._m.clear();
        },

        get size() { return this._m.size; },
    };

    /**
     * OptimizationResultCache —— 优化过程使用，**只保存指标，绝不保存 8760 原始结果**（§11）。
     *
     * 保存字段：scheme / technical / economic / constraints / objectives / vector / key
     * 明确排除：results（Float64Array 8760×11）
     *
     * 若调用方传入含 results 的结果对象，这里会**剥离** results 后再保存。
     */
    const OptimizationResultCache = {
        _m: new Map(),

        /** 剥离 8760 原始数据，只留指标 */
        strip(result) {
            if (!result) return result;
            const out = {
                key: result.key,
                simulationKey: result.simulationKey,
                scheme: result.scheme,
                technical: result.technical,
                economic: result.economic,
                constraints: result.constraints,
                objectives: result.objectives,
                vector: result.vector,
                durationMs: result.durationMs,
            };
            return out;
        },

        put(result) {
            const k = simulationKeyOf(result);
            if (!k) return result;
            this._m.set(k, this.strip(result));
            return result;
        },

        get(simulationKey) {
            return this._m.get(simulationKey);
        },

        clear() {
            this._m.clear();
        },

        get size() { return this._m.size; },
    };

    // =========================================================================
    // 七、导出辅助（§5：导出确实需要对象数组，但只应显式、按需地物化一次）
    // =========================================================================

    /**
     * 物化为「逐时对象数组」。
     *
     * ⚠ 这是**显式的高开销操作**：8760 个对象。只允许导出 / 详情等一次性场景调用，
     * 显示与图表路径禁止使用（应走 ChartDataAdapter）。
     *
     * @param {Object} result
     * @returns {Array<Object>} 每行含 11 个中文列名
     */
    function materialize(result) {
        if (!result) return [];
        const n = result.hours;
        const raw = result.results;
        const out = new Array(n);
        for (let h = 0; h < n; h++) {
            const base = h * COL_COUNT;
            const row = {};
            for (let c = 0; c < COL_COUNT; c++) row[COL_LABELS[c]] = raw[base + c];
            out[h] = row;
        }
        return out;
    }

    // =========================================================================
    // 八、导出接口
    // =========================================================================

    /**
     * 缓存清理接口（任务书 §34）。
     * 在「更换输入数据」「输入数据版本变化」「用户主动清理」时调用（§35）。
     */
    function clearSimulationCache() { SimulationResultCache.clear(); }
    function clearChartCache() { ChartDataCache.clear(); }
    function clearOptimizationCache() { OptimizationResultCache.clear(); }
    function clearAllCaches() {
        clearSimulationCache();
        clearChartCache();
        clearOptimizationCache();
    }

    const ResultDataStore = {
        // 常量
        COLS: COLS,
        COL_COUNT: COL_COUNT,
        COL_LABELS: COL_LABELS,
        KEY_TO_LABEL: KEY_TO_LABEL,
        LABEL_TO_KEY: LABEL_TO_KEY,
        MONTH_DAYS: MONTH_DAYS,
        SIMULATION_ENGINE_VERSION: SIMULATION_ENGINE_VERSION,

        // §4 要求的基础接口
        create: create,
        getValue: getValue,
        getHourRow: getHourRow,
        getRange: getRange,
        getColumn: getColumn,
        getLength: getLength,
        getAnnualSummary: getAnnualSummary,
        getMonthlySummary: getMonthlySummary,

        // V2.3.1：数据访问微优化（任务书 §3 / §4）
        getColumnArray: getColumnArray,
        createColumnView: createColumnView,
        getHourObject: getHourRow,
        createSimulationKey: createSimulationKey,
        simulationKeyOf: simulationKeyOf,
        createInputVersionFromRows: createInputVersionFromRows,
        downsampleByAccessor: downsampleByAccessor,

        // 扩展
        resolveColumn: resolveColumn,
        schemeKeyOf: schemeKeyOf,
        monthCumulativeHours: monthCumulativeHours,
        getMonthlyColumn: getMonthlyColumn,
        createHourView: createHourView,
        downsampleIndices: downsampleIndices,
        materialize: materialize,

        // 缓存与适配器
        SimulationResultCache: SimulationResultCache,
        OptimizationResultCache: OptimizationResultCache,
        ChartDataCache: ChartDataCache,
        ChartDataAdapter: ChartDataAdapter,

        // V2.3.1：缓存清理接口（任务书 §34）
        clearSimulationCache: clearSimulationCache,
        clearChartCache: clearChartCache,
        clearOptimizationCache: clearOptimizationCache,
        clearAllCaches: clearAllCaches,
    };

    self.ResultDataStore = ResultDataStore;
    self.ChartDataAdapter = ChartDataAdapter;
    self.ChartDataCache = ChartDataCache;
})();
