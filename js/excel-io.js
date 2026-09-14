/**
 * Excel IO 模块
 * 使用 SheetJS 读取 Excel，ExcelJS 导出 Excel
 */

const ExcelIO = {
    /**
     * 读取 input.xlsx（8760小时风光单MW电量数据，MWh/MW）
     * 返回 [{pv, wind}, ...]，仿真时由引擎按装机容量缩放
     */
    readInputExcel(arrayBuffer) {
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        // skiprows=1, header=0: 跳过第一行标题，使用默认列名
        const data = XLSX.utils.sheet_to_json(sheet, { header: ['光伏电量', '风电电量'], range: 1 });
        return data.map(row => ({
            pv: Utils.toNum(row['光伏电量']),
            wind: Utils.toNum(row['风电电量'])
        }));
    },

    /**
     * 读取数据汇总 Excel
     */
    readDataSummaryExcel(arrayBuffer) {
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        return XLSX.utils.sheet_to_json(sheet);
    },

    /**
     * 读取概算结果 Excel
     */
    readEstimateExcel(arrayBuffer) {
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        return XLSX.utils.sheet_to_json(sheet);
    },

    /**
     * 导出仿真结果为 Excel（3个Sheet：逐时数据+系统参数+比值分析）
     */
    async exportSimulationResult(resultData) {
        const workbook = new ExcelJS.Workbook();

        // Sheet1: 逐时数据
        const ws1 = workbook.addWorksheet('Sheet1');
        const columns = ['光伏电量', '风电电量', '合计电量', '储能充电量', '储能放电量',
                         '储能现存容量', '制氢电量', '上网电量', '下网电量', '弃电量', '制氢量'];
        ws1.columns = columns.map(c => ({ header: c, key: c, width: 14 }));

        for (const row of resultData.hourlyData) {
            ws1.addRow(row);
        }
        // 总和行
        // V2.3.1（任务书 §30）：若调用方已提供年度摘要（totals），直接读取，
        // 避免为求和再做 10 列 × 8760 次对象属性遍历；否则回退到逐列 reduce。
        const totalRow = {};
        const sumCols = ['光伏电量', '风电电量', '合计电量', '储能充电量', '储能放电量',
                          '制氢电量', '上网电量', '下网电量', '弃电量', '制氢量'];
        if (resultData.totals) {
            for (const col of sumCols) totalRow[col] = resultData.totals[col];
            totalRow['储能现存容量'] = resultData.totals['储能现存容量'];
        } else {
            for (const col of sumCols) {
                totalRow[col] = resultData.hourlyData.reduce((s, r) => s + (r[col] || 0), 0);
            }
            totalRow['储能现存容量'] = resultData.hourlyData[resultData.hourlyData.length - 1]['储能现存容量'];
        }
        ws1.addRow(totalRow);

        // 系统参数
        const ws2 = workbook.addWorksheet('系统参数');
        ws2.columns = [{ header: '参数', key: '参数', width: 30 }, { header: '值', key: '值', width: 20 }];
        for (const [key, val] of Object.entries(resultData.systemVars)) {
            ws2.addRow({ '参数': key, '值': val });
        }

        // 比值分析
        const ws3 = workbook.addWorksheet('比值分析');
        ws3.columns = [{ header: '指标', key: '指标', width: 30 }, { header: '比值', key: '比值', width: 15 }];
        for (const r of resultData.ratioData) {
            ws3.addRow({ '指标': r['指标'], '比值': r['比值'] });
        }

        return await workbook.xlsx.writeBuffer();
    },

    /**
     * 导出数据汇总为 Excel
     */
    async exportSummaryExcel(summaryData) {
        const workbook = new ExcelJS.Workbook();
        const ws = workbook.addWorksheet('数据汇总');

        if (summaryData.length > 0) {
            const keys = Object.keys(summaryData[0]);
            ws.columns = keys.map(k => ({ header: k, key: k, width: Math.max(k.length * 2 + 4, 14) }));
            for (const row of summaryData) {
                ws.addRow(row);
            }
        }

        return await workbook.xlsx.writeBuffer();
    },

    /**
     * 导出概算结果为 Excel
     */
    async exportEstimateResult(estimateData) {
        const workbook = new ExcelJS.Workbook();
        const ws = workbook.addWorksheet('概算结果');

        if (estimateData.length > 0) {
            const keys = Object.keys(estimateData[0]);
            ws.columns = keys.map(k => ({ header: k, key: k, width: Math.min(k.length * 2 + 4, 40) }));
            for (const row of estimateData) {
                ws.addRow(row);
            }
        }

        return await workbook.xlsx.writeBuffer();
    },

    /**
     * 导出经评指标为 Excel
     */
    async exportFinanceResult(resultData) {
        const workbook = new ExcelJS.Workbook();
        const ws = workbook.addWorksheet('财务指标汇总');
        ws.columns = [{ header: '指标名称', key: '指标名称', width: 45 }, { header: '数值', key: '数值', width: 20 }];

        for (const [key, val] of Object.entries(resultData)) {
            ws.addRow({ '指标名称': key, '数值': val });
        }

        return await workbook.xlsx.writeBuffer();
    },

    /**
     * 导出现金流量表为 Excel（2个Sheet）
     */
    async exportCashflow(projectCF, equityCF) {
        const workbook = new ExcelJS.Workbook();

        // 项目投资现金流量表
        const ws1 = workbook.addWorksheet('项目投资现金流量表');
        const pcfCols = ['年份', '营业收入（不含增值税）', '经营成本', '建设投资', '流动资金',
                         '所得税', '净现金流（所得税前）', '净现金流（所得税后）'];
        ws1.columns = pcfCols.map(c => ({ header: c, key: c, width: 20 }));
        for (const cf of projectCF) {
            ws1.addRow({
                '年份': cf.year,
                '营业收入（不含增值税）': Utils.toNum(cf.revenue, 0),
                '经营成本': Utils.toNum(cf.operating_cost, 0),
                '建设投资': Utils.toNum(cf.investment, 0),
                '流动资金': Utils.toNum(cf.working_capital, 0),
                '所得税': Utils.toNum(cf.income_tax, 0),
                '净现金流（所得税前）': Utils.toNum(cf.net_cf_pre_tax, 0),
                '净现金流（所得税后）': Utils.toNum(cf.net_cf_post_tax, 0),
            });
        }

        // 资本金现金流量表
        const ws2 = workbook.addWorksheet('资本金现金流量表');
        const ecfCols = ['年份', '资本金投入', '借款收到', '营业收入', '经营成本',
                         '借款还本', '借款付息', '所得税', '净现金流'];
        ws2.columns = ecfCols.map(c => ({ header: c, key: c, width: 20 }));
        for (const cf of equityCF) {
            ws2.addRow({
                '年份': cf.year,
                '资本金投入': Utils.toNum(cf.capital_investment, 0),
                '借款收到': Utils.toNum(cf.loan_received, 0),
                '营业收入': Utils.toNum(cf.revenue, 0),
                '经营成本': Utils.toNum(cf.operating_cost, 0),
                '借款还本': Utils.toNum(cf.loan_repay_principal, 0),
                '借款付息': Utils.toNum(cf.loan_repay_interest, 0),
                '所得税': Utils.toNum(cf.income_tax, 0),
                '净现金流': Utils.toNum(cf.net_cf, 0),
            });
        }

        return await workbook.xlsx.writeBuffer();
    },

    /**
     * 导出经评批量结果
     */
    async exportBatchFinanceResult(batchData) {
        const workbook = new ExcelJS.Workbook();
        const ws = workbook.addWorksheet('经评批量结果');

        if (batchData.length > 0) {
            const keys = Object.keys(batchData[0]);
            ws.columns = keys.map(k => ({ header: k, key: k, width: Math.min(k.length * 2 + 4, 35) }));
            for (const row of batchData) {
                ws.addRow(row);
            }
        }

        return await workbook.xlsx.writeBuffer();
    },

    // ==================== V2.1 新增：优化结果导出 ====================

    /** 数值安全化：非有限值（Infinity / NaN）在 Excel 中留空，避免出现 Infinity 字样 */
    _safeNum(v, digits) {
        if (v === null || v === undefined) return '';
        const n = Number(v);
        if (!isFinite(n)) return '';
        if (digits === undefined) return n;
        const p = Math.pow(10, digits);
        return Math.round(n * p) / p;
    },

    /** 单个方案 → 一行扁平数据（Pareto 表与代表方案表共用） */
    _schemeRow(s, rank) {
        const sc = s.scheme, te = s.technical, ec = s.economic;
        return {
            '排名': rank,
            '方案编号': s.id || '',
            '风电容量（MW）': this._safeNum(sc.windCapacity, 4),
            '光伏容量（MW）': this._safeNum(sc.pvCapacity, 4),
            '储能功率（MW）': this._safeNum(sc.storagePower, 4),
            '储能时长（h）': this._safeNum(sc.storageDuration, 4),
            '储能容量（MWh）': this._safeNum(sc.storageEnergy, 4),
            '电解槽容量（MW）': this._safeNum(sc.electrolyzerCapacity, 4),
            '年制氢量（万吨）': this._safeNum(te.annualHydrogenWanTon, 4),
            '电解槽利用小时（h）': this._safeNum(te.electrolyzerHours, 2),
            '弃电率（%）': this._safeNum(te.curtailmentRate * 100, 2),
            '外购电比例（%）': this._safeNum(te.gridImportRatio * 100, 2),
            '绿电制氢比例（%）': this._safeNum(te.greenHydrogenRatio * 100, 2),
            '上网比例（%）': this._safeNum(te.exportRatio * 100, 2),
            '总投资（万元）': this._safeNum(ec.totalInvestment, 2),
            'LCOH（元/kg）': this._safeNum(ec.LCOH, 2),
            'FIRR（%）': this._safeNum(ec.FIRR, 2),
            'EIRR（%）': this._safeNum(ec.EIRR, 2),
            '综合得分': s.score === null || s.score === undefined ? '' : this._safeNum(s.score * 100, 2),
        };
    },

    /**
     * 导出优化结果 Excel（5 个 Sheet）
     * Sheet1 优化参数 / Sheet2 Pareto方案 / Sheet3 代表方案 / Sheet4 优化过程 / Sheet5 基准方案对比
     * @param {Object} payload { config, statistics, paretoSolutions, representativeSolutions, baseline, history }
     */
    async exportOptimizationResults(payload) {
        const p = payload || {};
        const cfg = p.config || {};
        const stats = p.statistics || {};
        const pareto = p.paretoSolutions || [];
        const rep = p.representativeSolutions || {};
        const history = p.history || [];

        const workbook = new ExcelJS.Workbook();
        workbook.creator = '多能互补系统仿真计算平台 WEB-V2.1';

        // ---------- Sheet1：优化参数 ----------
        const ws1 = workbook.addWorksheet('优化参数');
        ws1.columns = [
            { header: '类别', key: '类别', width: 16 },
            { header: '参数', key: '参数', width: 30 },
            { header: '最小值', key: '最小值', width: 14 },
            { header: '最大值', key: '最大值', width: 14 },
            { header: '步长', key: '步长', width: 14 },
            { header: '档位数', key: '档位数', width: 12 },
            { header: '说明', key: '说明', width: 46 },
        ];

        const VAR_LABEL = {
            windCapacity: '风电容量 (MW)', pvCapacity: '光伏容量 (MW)',
            storagePower: '储能功率 (MW)', storageDuration: '储能时长 (h)',
            electrolyzerCapacity: '电解槽容量 (MW)',
        };
        const KEY_ORDER = ['windCapacity', 'pvCapacity', 'storagePower', 'storageDuration', 'electrolyzerCapacity'];

        for (const key of KEY_ORDER) {
            const v = (cfg.variables && cfg.variables[key]) || {};
            let levels = 0;
            try { levels = window.OptimizationEngine.buildLevelsFor(v.min, v.max, v.step).length; } catch (e) { levels = ''; }
            ws1.addRow({
                '类别': '优化变量', '参数': VAR_LABEL[key],
                '最小值': v.min, '最大值': v.max, '步长': v.step, '档位数': levels,
                '说明': '决策变量（离散档位）',
            });
        }
        ws1.addRow({
            '类别': '派生量', '参数': '储能容量 (MWh)', '最小值': '', '最大值': '', '步长': '', '档位数': '',
            '说明': '储能功率 × 储能时长',
        });

        const consMap = [
            ['minAnnualHydrogen', '年制氢量 ≥', '万吨/年'],
            ['maxCurtailmentRate', '弃电率 ≤', '%'],
            ['maxGridImportRatio', '外购电比例 ≤', '%'],
            ['minGreenHydrogenRatio', '绿电制氢比例 ≥', '%'],
            ['maxExportRatio', '上网比例 ≤', '%'],
            ['maxExportPower', '最大上网功率 ≤', 'MW'],
            ['maxImportPower', '最大下网功率 ≤', 'MW'],
        ];
        for (const [key, label, unit] of consMap) {
            const c = (cfg.constraints && cfg.constraints[key]) || {};
            ws1.addRow({
                '类别': '工程约束', '参数': label + ' (' + unit + ')',
                '最小值': '', '最大值': '',
                '步长': c.enabled ? c.value : '', '档位数': '',
                '说明': c.enabled ? '已启用' : '未启用',
            });
        }

        const ns = cfg.nsga2 || {};
        ws1.addRow({ '类别': 'NSGA-II', '参数': '种群规模（有效）', '步长': ns.populationSize, '说明': '受搜索空间规模自动收缩' });
        ws1.addRow({ '类别': 'NSGA-II', '参数': '最大迭代次数', '步长': ns.generations, '说明': '实际完成 ' + (stats.generationsCompleted || 0) + ' 代' });
        ws1.addRow({ '类别': 'NSGA-II', '参数': '交叉概率（SBX）', '步长': ns.crossoverProbability, '说明': '分布指数 ηc = ' + ns.distributionIndexCrossover });
        ws1.addRow({ '类别': 'NSGA-II', '参数': '变异概率', '步长': ns.mutationProbability, '说明': '离散步长变异（±1~2 个步长）' });
        ws1.addRow({ '类别': 'NSGA-II', '参数': '随机种子', '步长': ns.randomSeed, '说明': (ns.randomSeed ? '固定种子，结果可复现' : '') });
        ws1.addRow({ '类别': 'NSGA-II', '参数': '提前终止', '步长': ns.earlyStopping ? ('启用（patience=' + ns.patience + '）') : '未启用', '说明': stats.earlyStopped ? '本次已触发提前终止' : '' });

        const w = cfg.recommendationWeights || {};
        ws1.addRow({ '类别': '推荐权重', '参数': 'EIRR（资本金）权重', '步长': w.eirr, '说明': '仅用于从 Pareto 前沿挑选综合推荐方案' });
        ws1.addRow({ '类别': '推荐权重', '参数': 'LCOH 权重', '步长': w.lcoh, '说明': '不参与 NSGA-II 适应度' });
        ws1.addRow({ '类别': '推荐权重', '参数': '弃电率 权重', '步长': w.curtailmentRate, '说明': '' });

        ws1.addRow({ '类别': 'LCOH', '参数': '折现率 (%)', '步长': (cfg.lcoh || {}).discountRate, '说明': 'LCOH = 全生命周期折现成本 / 全生命周期折现制氢量（不含折旧、不含所得税）' });

        ws1.addRow({ '类别': '统计', '参数': '有效评价方案数', '步长': stats.totalEvaluated });
        ws1.addRow({ '类别': '统计', '参数': '缓存命中次数', '步长': stats.cacheHits });
        ws1.addRow({ '类别': '统计', '参数': '可行方案数', '步长': stats.feasibleCount });
        ws1.addRow({ '类别': '统计', '参数': 'Pareto 方案数', '步长': stats.paretoCount });
        ws1.addRow({ '类别': '统计', '参数': '搜索空间规模', '步长': stats.searchSpaceSize });
        ws1.addRow({ '类别': '统计', '参数': '耗时 (s)', '步长': this._safeNum(stats.elapsedTime, 2) });

        // ---------- Sheet2：Pareto方案 ----------
        const ws2 = workbook.addWorksheet('Pareto方案');
        if (pareto.length > 0) {
            const rows = pareto.map((s, i) => this._schemeRow(s, i + 1));
            const keys = Object.keys(rows[0]);
            ws2.columns = keys.map(k => ({ header: k, key: k, width: Math.min(Math.max(k.length * 2 + 4, 12), 22) }));
            for (const r of rows) ws2.addRow(r);
        } else {
            ws2.addRow({ '提示': '本次优化未获得可行 Pareto 方案，请放宽约束或扩大容量搜索范围。' });
        }

        // ---------- Sheet3：代表方案 ----------
        const ws3 = workbook.addWorksheet('代表方案');
        const reps = [
            ['方案A 经济最优（EIRR最高）', rep.economicBest],
            ['方案B 氢成本最优（LCOH最低）', rep.hydrogenCostBest],
            ['方案C 消纳最优（弃电率最低）', rep.curtailmentBest],
            ['方案D 综合推荐（得分最高）', rep.recommended],
        ];
        const repRows = reps.map(([name, s]) => {
            const row = this._schemeRow(s || { scheme: {}, technical: {}, economic: {} }, '');
            row['方案编号'] = name;
            return row;
        });
        const repKeys = Object.keys(repRows[0]);
        ws3.columns = [{ header: '指标', key: '指标', width: 22 }].concat(
            reps.map(([name]) => ({ header: name, key: name, width: 30 }))
        );
        for (const k of repKeys) {
            if (k === '方案编号') continue;
            const row = { '指标': k };
            reps.forEach(([name], i) => { row[name] = repRows[i][k]; });
            ws3.addRow(row);
        }
        ws3.addRow({});
        ws3.addRow({ '指标': '说明', [reps[0][0]]: '代表方案均取自 Pareto 前沿；若某项指标不可得则留空。' });

        // ---------- Sheet4：优化过程 ----------
        const ws4 = workbook.addWorksheet('优化过程');
        ws4.columns = [
            { header: 'Generation', key: 'Generation', width: 12 },
            { header: 'Population', key: 'Population', width: 12 },
            { header: 'FeasibleCount', key: 'FeasibleCount', width: 14 },
            { header: 'ParetoCount', key: 'ParetoCount', width: 12 },
            { header: 'BestEIRR', key: 'BestEIRR', width: 12 },
            { header: 'BestLCOH', key: 'BestLCOH', width: 12 },
            { header: 'BestCurtailment', key: 'BestCurtailment', width: 16 },
        ];
        for (const h of history) {
            ws4.addRow({
                'Generation': h.generation,
                'Population': h.evaluatedCount,
                'FeasibleCount': h.feasibleCount,
                'ParetoCount': h.paretoCount,
                'BestEIRR': this._safeNum(h.bestEirr, 4),
                'BestLCOH': this._safeNum(h.bestLCOH, 4),
                'BestCurtailment': this._safeNum(h.bestCurtailmentRate, 6),
            });
        }
        if (history.length === 0) ws4.addRow({ 'Generation': '无数据' });

        // ---------- Sheet5：基准方案对比 ----------
        const ws5 = workbook.addWorksheet('基准方案对比');
        const base = p.baseline;
        const rec = rep.recommended;
        if (base && rec) {
            const bRow = this._schemeRow(base, '基准');
            const rRow = this._schemeRow(rec, '推荐');
            ws5.columns = [
                { header: '指标', key: '指标', width: 22 },
                { header: '基准方案 (Baseline)', key: 'base', width: 24 },
                { header: '综合推荐方案', key: 'rec', width: 24 },
                { header: '变化', key: 'delta', width: 20 },
            ];
            const numericDirs = {
                '风电容量（MW）': 0, '光伏容量（MW）': 0, '储能功率（MW）': 0, '储能时长（h）': 0,
                '储能容量（MWh）': 0, '电解槽容量（MW）': 0, '年制氢量（万吨）': 1,
                '电解槽利用小时（h）': 1, '弃电率（%）': -1, '外购电比例（%）': -1,
                '绿电制氢比例（%）': 1, '上网比例（%）': -1, '总投资（万元）': -1,
                'LCOH（元/kg）': -1, 'FIRR（%）': 1, 'EIRR（%）': 1,
            };
            for (const k of Object.keys(bRow)) {
                if (k === '排名' || k === '方案编号' || k === '综合得分') continue;
                const bv = bRow[k], rv = rRow[k];
                let delta = '';
                if (typeof bv === 'number' && typeof rv === 'number') {
                    const dir = numericDirs[k] || 0;
                    const d = Math.round((rv - bv) * 100) / 100;
                    delta = (d > 0 ? '+' : '') + d + (dir > 0 ? '（越大越好）' : dir < 0 ? '（越小越好）' : '');
                }
                ws5.addRow({ '指标': k, 'base': bv, 'rec': rv, 'delta': delta });
            }
            ws5.addRow({});
            ws5.addRow({
                '指标': '结论',
                'base': '优化后 EIRR（资本金）提高 ' + this._deltaText(rRow['EIRR（%）'], bRow['EIRR（%）'], '个百分点'),
                'rec': 'LCOH 降低 ' + this._deltaText(bRow['LCOH（元/kg）'], rRow['LCOH（元/kg）'], '元/kg'),
                'delta': '弃电率降低 ' + this._deltaText(bRow['弃电率（%）'], rRow['弃电率（%）'], '个百分点'),
            });
        } else {
            ws5.addRow({ '指标': '提示', 'base': '未提供基准方案（Baseline），无法对比。' });
        }

        return await workbook.xlsx.writeBuffer();
    },

    /** 生成「A 相对 B」的差值文本 */
    _deltaText(a, b, unit) {
        if (typeof a !== 'number' || typeof b !== 'number') return '—';
        const d = Math.round((a - b) * 100) / 100;
        return (d > 0 ? '+' : '') + d + ' ' + unit;
    }
};

window.ExcelIO = ExcelIO;
