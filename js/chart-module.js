/**
 * 图表模块
 * 使用 ECharts 实现交互式图表，替代 matplotlib
 * - 在线预览：深色主题（透明背景）
 * - 导出 PNG：浅色主题（白色背景），便于嵌入报告
 */

const ChartModule = {
    // 颜色映射（与原版 matplotlib 配色一致）
    COLORS: {
        pv: '#1f77b4',
        wind: '#ff7f0e',
        total: '#2ca02c',
        charge: '#d62728',
        discharge: '#9467bd',
        storage: '#8c564b',
        hydrogen: '#e377c2',
        export_: '#7f7f7f',
        import_: '#bcbd22',
        curtailment: '#17becf',
        h2prod: '#1a55FF',
    },

    // 主题配置
    THEMES: {
        dark: { bg: 'transparent', text: '#e0e0e0', axisLine: '#555', splitLine: '#333' },
        light: { bg: '#ffffff', text: '#333333', axisLine: '#999', splitLine: '#e0e0e0' },
    },

    MONTH_HOURS: [0, 744, 1416, 2160, 2880, 3624, 4344, 5088, 5832, 6552, 7296, 8016, 8760],
    MONTH_NAMES: ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'],
    MONTH_DAYS: [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31],

    COL_NAMES: {
        pv: '光伏电量', wind: '风电电量', total: '合计电量',
        charge: '储能充电量', discharge: '储能放电量', storage: '储能现存容量',
        hydrogen: '制氢电量', export: '上网电量', import: '下网电量',
        curtailment: '弃电量', h2prod: '制氢量',
    },

    // 7 项能量流动序列（正方向 + 负方向），用于汇总类图表
    FLOW_POS: [
        { key: '光伏电量', color: '#1f77b4' },
        { key: '风电电量', color: '#ff7f0e' },
        { key: '储能放电量', color: '#9467bd' },
        { key: '下网电量', color: '#bcbd22' },
    ],
    FLOW_NEG: [
        { key: '储能充电量', color: '#d62728' },
        { key: '制氢电量', color: '#e377c2' },
        { key: '上网电量', color: '#7f7f7f' },
    ],

    /** 获取主题对象 */
    _theme(forExport) {
        return this.THEMES[forExport ? 'light' : 'dark'];
    },

    /** 公共轴样式 */
    _axisStyle(forExport) {
        const t = this._theme(forExport);
        return {
            axisLine: { lineStyle: { color: t.axisLine } },
            axisLabel: { color: t.text },
            splitLine: { lineStyle: { color: t.splitLine } },
            nameTextStyle: { color: t.text },
        };
    },

    /**
     * 【已废弃 · V2.3】将仿真结果扁平数组转换为可读对象数组。
     *
     * 该函数一次产生 8760 个对象 × 11 个属性（约数 MB 垃圾），是 V2.2 浏览器卡顿的
     * 主要来源之一。V2.3 起**显示路径禁止调用**，改用 ResultDataStore 的 HourView
     * （零对象分配）与 ChartDataAdapter（按需降采样）。
     *
     * 仅保留给「确实需要持久对象数组」的场景（如 Excel 导出），且应优先改用
     * ResultDataStore.materialize(result)。
     *
     * TODO V2.4 REMOVE LEGACY
     */
    parseResults(resultsArray) {
        const COLS = 11;
        const keys = ['光伏电量', '风电电量', '合计电量', '储能充电量', '储能放电量',
                       '储能现存容量', '制氢电量', '上网电量', '下网电量', '弃电量', '制氢量'];
        const data = [];
        for (let h = 0; h < resultsArray.length / COLS; h++) {
            const row = {};
            for (let c = 0; c < COLS; c++) row[keys[c]] = resultsArray[h * COLS + c];
            data.push(row);
        }
        return data;
    },

    /** 计算每月累计起始小时（基于各月实际天数） */
    _monthCumulativeHours() {
        const cum = [0];
        for (let i = 0; i < 12; i++) cum.push(cum[i] + this.MONTH_DAYS[i] * 24);
        return cum;
    },

    /** 每月随机选取一个典型日（返回 [{month, dayStart}, ...]） */
    getRandomTypicalDays() {
        const cum = this._monthCumulativeHours();
        const result = [];
        for (let m = 0; m < 12; m++) {
            const randomDay = Math.floor(Math.random() * this.MONTH_DAYS[m]);
            result.push({ month: m, dayStart: cum[m] + randomDay * 24 });
        }
        return result;
    },

    /** 每月随机选取一个典型周（返回 [{month, weekStart}, ...]） */
    getRandomTypicalWeeks() {
        const cum = this._monthCumulativeHours();
        const result = [];
        for (let m = 0; m < 12; m++) {
            const monthHours = this.MONTH_DAYS[m] * 24;
            const maxStart = Math.max(0, monthHours - 168); // 确保完整 168 小时不跨月
            const weekStart = cum[m] + Math.floor(Math.random() * (maxStart + 1));
            result.push({ month: m, weekStart });
        }
        return result;
    },

    // ==================== 渲染函数 ====================

    /** 全年能量流动汇总图 */
    renderOverviewChart(container, data, forExport = false) {
        const t = this._theme(forExport);
        const chart = this._initChart(container, forExport);
        // V2.3：data 为 ResultDataStore 的 HourView（可能是降采样视图）。
        // map 的第二个参数是「真实小时序号」，因此类别轴仍使用真实小时号，不因降采样而错位。
        const hours = data.map((_, h) => h);
        const axisInterval = data.axisInterval || Math.max(1, Math.floor(hours.length / 8));
        const ax = this._axisStyle(forExport);

        chart.setOption({
            backgroundColor: t.bg,
            title: { text: '全年能量流动汇总', left: 'center', textStyle: { color: t.text } },
            tooltip: { trigger: 'axis', axisPointer: { type: 'cross' } },
            legend: { data: [...this.FLOW_POS, ...this.FLOW_NEG].map(s => s.key), bottom: 0, textStyle: { color: t.text } },
            grid: { left: 70, right: 40, top: 50, bottom: forExport ? 70 : 70 },
            xAxis: { type: 'category', data: hours, name: '小时', axisLabel: { interval: axisInterval, color: t.text }, ...ax },
            yAxis: { type: 'value', name: 'MWh', ...ax },
            dataZoom: forExport ? [] : [{ type: 'inside' }, { type: 'slider', bottom: 30 }],
            animation: !forExport,
            series: [
                ...this.FLOW_POS.map(s => ({ name: s.key, type: 'line', stack: 'pos', areaStyle: { opacity: 0.7 }, data: data.map(d => d[s.key]), itemStyle: { color: s.color }, symbol: 'none', sampling: 'lttb' })),
                ...this.FLOW_NEG.map(s => ({ name: s.key, type: 'line', stack: 'neg', areaStyle: { opacity: 0.7 }, data: data.map(d => -d[s.key]), itemStyle: { color: s.color }, symbol: 'none', sampling: 'lttb' })),
            ]
        });
        return chart;
    },

    /** 逐时曲线图（单列） */
    renderHourlyChart(container, data, colKey, forExport = false) {
        const t = this._theme(forExport);
        const chart = this._initChart(container, forExport);
        const colName = this.COL_NAMES[colKey] || colKey;
        // V2.3：data 为 HourView（可能是降采样视图），map 的第二参数是真实小时序号
        const hours = data.map((_, h) => h);
        const values = data.map(d => colKey === 'h2prod' ? d['制氢量'] / 1000 : d[colName]);
        const unit = colKey === 'h2prod' ? '吨' : 'MWh';
        const color = this.COLORS[colKey] || '#58a6ff';
        const ax = this._axisStyle(forExport);
        const axisInterval = data.axisInterval || Math.max(1, Math.floor(hours.length / 8));

        chart.setOption({
            backgroundColor: t.bg,
            title: { text: `${colName}逐时曲线`, left: 'center', textStyle: { color: t.text } },
            tooltip: { trigger: 'axis', formatter: p => `第${p[0].axisValue}小时: ${p[0].value.toFixed(2)} ${unit}` },
            grid: { left: 70, right: 40, top: 50, bottom: forExport ? 50 : 70 },
            xAxis: { type: 'category', data: hours, name: '小时', axisLabel: { interval: axisInterval, color: t.text }, ...ax },
            yAxis: { type: 'value', name: unit, ...ax },
            dataZoom: forExport ? [] : [{ type: 'inside' }, { type: 'slider', bottom: 30 }],
            animation: !forExport,
            series: [{ type: 'line', data: values, itemStyle: { color }, symbol: 'none', sampling: 'lttb', areaStyle: { opacity: 0.15 } }]
        });
        return chart;
    },

    /**
     * 取 12 个月的分项合计（任务书 §18 / §19）。
     *
     * 优先路径：data 为 ResultDataStore 的「月度视图」（含 byLabel），直接读缓存，
     *           完全不遍历 8760 小时；
     * 兜底路径：data 为全分辨率 HourView / 旧对象数组，逐月 reduce（仅导出等低频场景）。
     *
     * @returns {Array<Object>} 长度 12，每项含 8 个 MWh 分项 + 制氢量(吨)
     */
    _monthlyTotals(data) {
        if (data && data.byLabel) {
            const b = data.byLabel;
            return this.MONTH_NAMES.map((_, m) => ({
                '光伏电量': b['光伏电量'][m],
                '风电电量': b['风电电量'][m],
                '储能放电量': b['储能放电量'][m],
                '下网电量': b['下网电量'][m],
                '储能充电量': b['储能充电量'][m],
                '制氢电量': b['制氢电量'][m],
                '上网电量': b['上网电量'][m],
                '弃电量': b['弃电量'][m],
                '制氢量(吨)': b['制氢量'][m] / 1000,
            }));
        }
        const out = [];
        for (let m = 0; m < 12; m++) {
            const md = data.slice(this.MONTH_HOURS[m], this.MONTH_HOURS[m + 1]);
            out.push({
                '光伏电量': md.reduce((s, d) => s + d['光伏电量'], 0),
                '风电电量': md.reduce((s, d) => s + d['风电电量'], 0),
                '储能放电量': md.reduce((s, d) => s + d['储能放电量'], 0),
                '下网电量': md.reduce((s, d) => s + d['下网电量'], 0),
                '储能充电量': md.reduce((s, d) => s + d['储能充电量'], 0),
                '制氢电量': md.reduce((s, d) => s + d['制氢电量'], 0),
                '上网电量': md.reduce((s, d) => s + d['上网电量'], 0),
                '弃电量': md.reduce((s, d) => s + d['弃电量'], 0),
                '制氢量(吨)': md.reduce((s, d) => s + d['制氢量'], 0) / 1000,
            });
        }
        return out;
    },

    /**
     * 取 12 个月某列的合计（正负号按 FLOW 分组处理）。
     * 优先读月度缓存，兜底逐月 reduce。
     * @returns {Array<Array<number>>} [seriesIndex][month]
     */
    _monthlySeries(data, specs) {
        const hasCache = !!(data && data.byLabel);
        return specs.map(s => {
            const label = s.key;
            if (hasCache) {
                const col = data.byLabel[label];
                return this.MONTH_NAMES.map((_, m) => Math.round(col[m] * 100) / 100);
            }
            return this.MONTH_NAMES.map((_, m) => {
                const md = data.slice(this.MONTH_HOURS[m], this.MONTH_HOURS[m + 1]);
                return Math.round(md.reduce((sum, d) => sum + d[label], 0) * 100) / 100;
            });
        });
    },

    /** 逐月柱状图（9项独立柱） */
    renderMonthlyChart(container, data, forExport = false) {
        const t = this._theme(forExport);
        const chart = this._initChart(container, forExport);
        const ax = this._axisStyle(forExport);

        // V2.3：月度合计统一取自 ResultDataStore 月度缓存（全分辨率、只算一次），
        // 不再每次渲染都遍历 8760 小时（任务书 §18 / §19）
        const monthlyTotals = this._monthlyTotals(data);

        const barCols = [
            { key: '光伏电量', color: this.COLORS.pv },
            { key: '风电电量', color: this.COLORS.wind },
            { key: '储能放电量', color: this.COLORS.discharge },
            { key: '下网电量', color: this.COLORS.import_ },
            { key: '储能充电量', color: this.COLORS.charge },
            { key: '制氢电量', color: this.COLORS.hydrogen },
            { key: '上网电量', color: this.COLORS.export_ },
            { key: '弃电量', color: this.COLORS.curtailment },
            { key: '制氢量(吨)', color: this.COLORS.h2prod },
        ];

        chart.setOption({
            backgroundColor: t.bg,
            title: { text: '逐月电量柱状图', left: 'center', textStyle: { color: t.text } },
            tooltip: { trigger: 'axis' },
            legend: { data: barCols.map(c => c.key), bottom: 0, textStyle: { color: t.text, fontSize: 10 } },
            grid: { left: 70, right: 40, top: 50, bottom: 90 },
            xAxis: { type: 'category', data: this.MONTH_NAMES, ...ax },
            yAxis: { type: 'value', name: 'MWh', ...ax },
            animation: !forExport,
            series: barCols.map(c => ({
                name: c.key, type: 'bar',
                data: monthlyTotals.map(m => Math.round(m[c.key] * 100) / 100),
                itemStyle: { color: c.color }, barMaxWidth: 30,
            }))
        });
        return chart;
    },

    /** 逐月电量汇总图（7项正负堆叠面积图，对应原版 fill_between） */
    renderMonthlyOverviewChart(container, data, forExport = false) {
        const t = this._theme(forExport);
        const chart = this._initChart(container, forExport);
        const ax = this._axisStyle(forExport);

        // 按月汇总各项（V2.3：优先读 ResultDataStore 月度缓存，避免重复遍历 8760）
        const posSeries = this._monthlySeries(data, this.FLOW_POS);
        const negSeries = this._monthlySeries(data, this.FLOW_NEG);
        const posData = this.FLOW_POS.map((s, i) => posSeries[i]);
        const negData = this.FLOW_NEG.map((s, i) => negSeries[i].map(v => -v));

        chart.setOption({
            backgroundColor: t.bg,
            title: { text: '逐月电量汇总', left: 'center', textStyle: { color: t.text } },
            tooltip: { trigger: 'axis' },
            legend: { data: [...this.FLOW_POS, ...this.FLOW_NEG].map(s => s.key), bottom: 0, textStyle: { color: t.text, fontSize: 10 } },
            grid: { left: 70, right: 40, top: 50, bottom: 90 },
            xAxis: { type: 'category', data: this.MONTH_NAMES, ...ax },
            yAxis: { type: 'value', name: 'MWh', ...ax },
            animation: !forExport,
            series: [
                ...this.FLOW_POS.map((s, i) => ({ name: s.key, type: 'line', stack: 'pos', areaStyle: { opacity: 0.7 }, data: posData[i], itemStyle: { color: s.color }, symbol: 'none' })),
                ...this.FLOW_NEG.map((s, i) => ({ name: s.key, type: 'line', stack: 'neg', areaStyle: { opacity: 0.7 }, data: negData[i], itemStyle: { color: s.color }, symbol: 'none' })),
            ]
        });
        return chart;
    },

    /** 典型日曲线（指定月份 + 起始小时） */
    renderTypicalDayChart(container, data, month, dayStart = null, forExport = false) {
        const t = this._theme(forExport);
        const chart = this._initChart(container, forExport);
        const ax = this._axisStyle(forExport);
        const cum = this._monthCumulativeHours();
        const start = dayStart !== null ? dayStart : cum[month];
        const dayData = data.slice(start, start + 24);

        chart.setOption({
            backgroundColor: t.bg,
            title: { text: `${month + 1}月典型日电量曲线`, left: 'center', textStyle: { color: t.text } },
            tooltip: { trigger: 'axis' },
            legend: { data: [...this.FLOW_POS, ...this.FLOW_NEG].map(s => s.key), bottom: 0, textStyle: { color: t.text, fontSize: 10 } },
            grid: { left: 60, right: 30, top: 50, bottom: 70 },
            xAxis: { type: 'category', data: Array.from({ length: 24 }, (_, i) => `${i}时`), ...ax },
            yAxis: { type: 'value', name: 'MWh', ...ax },
            animation: !forExport,
            series: [
                ...this.FLOW_POS.map(s => ({ name: s.key, type: 'line', stack: 'pos', areaStyle: { opacity: 0.7 }, data: dayData.map(d => d[s.key]), itemStyle: { color: s.color }, symbol: 'none' })),
                ...this.FLOW_NEG.map(s => ({ name: s.key, type: 'line', stack: 'neg', areaStyle: { opacity: 0.7 }, data: dayData.map(d => -d[s.key]), itemStyle: { color: s.color }, symbol: 'none' })),
            ]
        });
        return chart;
    },

    /**
     * 逐月单月电量汇总图（对应 Python 中 1-12 月各月的 fill_between 宽幅图）
     * @param {HTMLElement} container
     * @param {Array}  data - 全年逐时数据
     * @param {number} month - 0-based 月份索引 (0=1月 … 11=12月)
     * @param {boolean} forExport
     */
    renderMonthlyDetailChart(container, data, month, forExport = false) {
        const t = this._theme(forExport);
        const chart = this._initChart(container, forExport);
        const ax = this._axisStyle(forExport);
        const cum = this._monthCumulativeHours();
        const start = cum[month];
        const end = cum[month + 1];
        const monthData = data.slice(start, end);
        const len = monthData.length;
        const xLabels = Array.from({ length: len }, (_, i) => i);

        chart.setOption({
            backgroundColor: t.bg,
            title: { text: `${month + 1}月电量汇总图`, left: 'center', textStyle: { color: t.text, fontSize: 16 } },
            tooltip: { trigger: 'axis', axisPointer: { type: 'cross' } },
            legend: {
                data: [...this.FLOW_POS, ...this.FLOW_NEG].map(s => s.key),
                bottom: 0,
                textStyle: { color: t.text, fontSize: 11 },
            },
            grid: { left: 70, right: 40, top: 55, bottom: forExport ? 75 : 80 },
            xAxis: {
                type: 'category',
                data: xLabels,
                name: '小时',
                axisLabel: { interval: Math.floor(len / 8), color: t.text },
                ...ax,
            },
            yAxis: { type: 'value', name: 'MWh', ...ax },
            dataZoom: forExport ? [] : [{ type: 'inside' }, { type: 'slider', bottom: 30 }],
            animation: !forExport,
            series: [
                ...this.FLOW_POS.map(s => ({
                    name: s.key, type: 'line', stack: 'pos',
                    areaStyle: { opacity: 0.7 },
                    data: monthData.map(d => d[s.key]),
                    itemStyle: { color: s.color }, symbol: 'none', sampling: 'lttb',
                })),
                ...this.FLOW_NEG.map(s => ({
                    name: s.key, type: 'line', stack: 'neg',
                    areaStyle: { opacity: 0.7 },
                    data: monthData.map(d => -d[s.key]),
                    itemStyle: { color: s.color }, symbol: 'none', sampling: 'lttb',
                })),
            ],
        });
        return chart;
    },

    /**
     * 分模块逐月柱状图（对应 Python 中 9 项独立柱状图，每项12个月）
     * @param {HTMLElement} container
     * @param {Array}  data - 全年逐时数据
     * @param {string} colKey - 数据列 key（与 COL_NAMES/COLORS 对应）
     * @param {boolean} forExport
     */
    renderModuleMonthlyChart(container, data, colKey, forExport = false) {
        const t = this._theme(forExport);
        const chart = this._initChart(container, forExport);
        const ax = this._axisStyle(forExport);

        // 制氢量(吨) 需特殊处理
        const isH2Ton = colKey === 'h2prod';
        const rawColName = isH2Ton ? '制氢量' : (this.COL_NAMES[colKey] || colKey);
        const displayName = isH2Ton ? '制氢量(吨)' : rawColName;
        const unit = isH2Ton ? '吨' : 'MWh';
        const color = this.COLORS[colKey] || '#58a6ff';

        // V2.3：优先读 ResultDataStore 月度缓存（全分辨率、只算一次），避免重复遍历 8760
        const hasCache = !!(data && data.byLabel);
        const monthlyValues = this.MONTH_NAMES.map((_, m) => {
            let sum;
            if (hasCache) {
                sum = data.byLabel[rawColName][m];
            } else {
                const md = data.slice(this.MONTH_HOURS[m], this.MONTH_HOURS[m + 1]);
                sum = md.reduce((s, d) => s + d[rawColName], 0);
            }
            if (isH2Ton) sum /= 1000;
            return Math.round(sum * 100) / 100;
        });

        chart.setOption({
            backgroundColor: t.bg,
            title: { text: `${displayName}逐月汇总`, left: 'center', textStyle: { color: t.text, fontSize: 16 } },
            tooltip: { trigger: 'axis', formatter: p => `${p[0].name}: ${p[0].value.toFixed(2)} ${unit}` },
            grid: { left: 70, right: 40, top: 55, bottom: forExport ? 50 : 60 },
            xAxis: { type: 'category', data: this.MONTH_NAMES, name: '月份', ...ax },
            yAxis: { type: 'value', name: unit, ...ax },
            animation: !forExport,
            series: [{
                type: 'bar',
                data: monthlyValues,
                itemStyle: { color },
                barMaxWidth: 50,
                label: {
                    show: forExport,
                    position: 'top',
                    formatter: p => p.value.toFixed(1),
                    color: t.text,
                    fontSize: 11,
                },
            }],
        });
        return chart;
    },

    /** 典型周曲线（指定月份 + 起始小时） */
    renderTypicalWeekChart(container, data, month, weekStart = null, forExport = false) {
        const t = this._theme(forExport);
        const chart = this._initChart(container, forExport);
        const ax = this._axisStyle(forExport);
        const cum = this._monthCumulativeHours();
        const start = weekStart !== null ? weekStart : cum[month];
        const weekData = data.slice(start, start + 168);

        chart.setOption({
            backgroundColor: t.bg,
            title: { text: `${month + 1}月典型周电量曲线`, left: 'center', textStyle: { color: t.text } },
            tooltip: { trigger: 'axis' },
            legend: { data: [...this.FLOW_POS, ...this.FLOW_NEG].map(s => s.key), bottom: 0, textStyle: { color: t.text, fontSize: 10 } },
            grid: { left: 60, right: 30, top: 50, bottom: 70 },
            xAxis: { type: 'category', data: Array.from({ length: 168 }, (_, i) => `${i}时`), axisLabel: { interval: 23, color: t.text }, ...ax },
            yAxis: { type: 'value', name: 'MWh', ...ax },
            dataZoom: forExport ? [] : [{ type: 'inside' }],
            animation: !forExport,
            series: [
                ...this.FLOW_POS.map(s => ({ name: s.key, type: 'line', stack: 'pos', areaStyle: { opacity: 0.7 }, data: weekData.map(d => d[s.key]), itemStyle: { color: s.color }, symbol: 'none', sampling: 'lttb' })),
                ...this.FLOW_NEG.map(s => ({ name: s.key, type: 'line', stack: 'neg', areaStyle: { opacity: 0.7 }, data: weekData.map(d => -d[s.key]), itemStyle: { color: s.color }, symbol: 'none', sampling: 'lttb' })),
            ]
        });
        return chart;
    },

    // ==================== V2.1 新增：优化结果图表 ====================

    /** Pareto 图坐标定义 */
    PARETO_AXES: {
        'firr-lcoh': {
            xName: 'LCOH (元/kg)', yName: 'FIRR (%)',
            x: s => s.economic.LCOH, y: s => s.economic.FIRR,
        },
        'lcoh-curtail': {
            xName: 'LCOH (元/kg)', yName: '弃电率 (%)',
            x: s => s.economic.LCOH, y: s => s.technical.curtailmentRate * 100,
        },
        'scale-firr': {
            xName: '风电+光伏装机 (MW)', yName: 'FIRR (%)',
            x: s => s.scheme.windCapacity + s.scheme.pvCapacity, y: s => s.economic.FIRR,
        },
    },

    /**
     * Pareto 前沿散点图（三类视图共用）
     * @param {HTMLElement} container
     * @param {Array}  solutions - Pareto 方案数组（统一评价结果）
     * @param {string} kind - 'firr-lcoh' | 'lcoh-curtail' | 'scale-firr'
     * @param {Object} opts - { forExport, onSelect(solution), repKeys:{recommended,economicBest,hydrogenCostBest,curtailmentBest}, selectedKey }
     */
    renderParetoScatter(container, solutions, kind, opts) {
        const o = opts || {};
        const forExport = !!o.forExport;
        const t = this._theme(forExport);
        const chart = this._initChart(container, forExport);
        const ax = this._axisStyle(forExport);
        const axis = this.PARETO_AXES[kind] || this.PARETO_AXES['firr-lcoh'];
        const rep = o.repKeys || {};

        const repDefs = [
            { key: rep.recommended, name: '综合推荐', color: '#f85149', symbol: 'diamond', size: 22 },
            { key: rep.economicBest, name: '经济最优', color: '#d29922', symbol: 'circle', size: 16 },
            { key: rep.hydrogenCostBest, name: '氢成本最优', color: '#00d4ff', symbol: 'triangle', size: 16 },
            { key: rep.curtailmentBest, name: '消纳最优', color: '#3fb950', symbol: 'rect', size: 14 },
        ];

        // 去重：同一方案同时命中多个代表方案时，只保留优先级最高的一类
        const usedKeys = new Set();
        const activeReps = [];
        for (const d of repDefs) {
            if (!d.key || usedKeys.has(d.key)) continue;
            usedKeys.add(d.key);
            activeReps.push(d);
        }

        const baseData = [];
        const repSeries = activeReps.map(d => ({ def: d, data: [] }));

        solutions.forEach((s, i) => {
            const x = axis.x(s);
            const y = axis.y(s);
            if (!isFinite(x) || !isFinite(y)) return;

            const isSelected = o.selectedKey && s.key === o.selectedKey;
            const item = {
                value: [x, y],
                schemeIndex: i,
                symbolSize: isSelected ? 20 : 10,
                itemStyle: isSelected ? { borderColor: '#ffffff', borderWidth: 2 } : undefined,
            };

            const hit = activeReps.findIndex(d => d.key === s.key);
            if (hit >= 0) {
                // 代表方案：使用系列级 symbolSize；被选中时放大并描边
                repSeries[hit].data.push({
                    value: [x, y],
                    schemeIndex: i,
                    symbolSize: isSelected ? activeReps[hit].size + 8 : activeReps[hit].size,
                    itemStyle: isSelected ? { borderColor: '#ffffff', borderWidth: 3 } : undefined,
                });
            } else {
                baseData.push(item);
            }
        });

        const series = [
            {
                name: 'Pareto 方案', type: 'scatter', data: baseData,
                symbolSize: 10, itemStyle: { color: '#58a6ff', opacity: 0.85 },
            },
            ...repSeries.map(rs => ({
                name: rs.def.name, type: 'scatter', data: rs.data,
                symbol: rs.def.symbol, symbolSize: rs.def.size,
                itemStyle: { color: rs.def.color, borderColor: '#fff', borderWidth: 1 },
                z: 10,
            })),
        ];

        const fmt = (v, d) => (v === null || v === undefined || !isFinite(v)) ? '—' : Number(v).toFixed(d);

        chart.setOption({
            backgroundColor: t.bg,
            title: {
                text: `Pareto 前沿（共 ${solutions.length} 个非支配方案）`,
                left: 'center',
                textStyle: { color: t.text, fontSize: 14 },
            },
            tooltip: {
                trigger: 'item',
                formatter: p => {
                    const s = solutions[p.data.schemeIndex];
                    if (!s) return '';
                    const sc = s.scheme;
                    return [
                        `<b>${s.id || '方案'}</b>`,
                        `风电 ${sc.windCapacity} MW ｜ 光伏 ${sc.pvCapacity} MW`,
                        `储能 ${sc.storagePower} MW / ${sc.storageDuration} h`,
                        `电解槽 ${sc.electrolyzerCapacity} MW`,
                        `年制氢量 ${fmt(s.technical.annualHydrogenTon, 0)} t/a`,
                        `弃电率 ${fmt(s.technical.curtailmentRate * 100, 2)} %`,
                        `LCOH ${fmt(s.economic.LCOH, 2)} 元/kg`,
                        `FIRR ${fmt(s.economic.FIRR, 2)} %`,
                        `<span style="color:#8b949e">点击查看方案详情</span>`,
                    ].join('<br/>');
                },
            },
            legend: { data: series.map(s => s.name), bottom: 0, textStyle: { color: t.text, fontSize: 10 } },
            grid: { left: 74, right: 40, top: 50, bottom: 60 },
            xAxis: { type: 'value', name: axis.xName, nameLocation: 'middle', nameGap: 28, scale: true, ...ax },
            yAxis: { type: 'value', name: axis.yName, nameLocation: 'middle', nameGap: 46, scale: true, ...ax },
            animation: !forExport,
            series: series,
        });

        if (!forExport && typeof o.onSelect === 'function') {
            chart.on('click', params => {
                const si = params && params.data && params.data.schemeIndex;
                if (si !== undefined && si !== null && solutions[si]) o.onSelect(solutions[si]);
            });
        }

        return chart;
    },

    /**
     * 优化收敛曲线（Generation vs 各代最优指标）
     * @param {HTMLElement} container
     * @param {Array}  history - 引擎返回的 history 数组
     * @param {string} metric - 'firr' | 'lcoh' | 'curtailment'
     * @param {boolean} forExport
     */
    renderConvergenceChart(container, history, metric, forExport = false) {
        const t = this._theme(forExport);
        const chart = this._initChart(container, forExport);
        const ax = this._axisStyle(forExport);

        const meta = {
            firr: { title: '收敛曲线：Generation vs Best FIRR', name: 'Best FIRR (%)', color: '#3fb950', get: h => h.bestFIRR },
            lcoh: { title: '收敛曲线：Generation vs Best LCOH', name: 'Best LCOH (元/kg)', color: '#58a6ff', get: h => h.bestLCOH },
            curtailment: { title: '收敛曲线：Generation vs Best Curtailment Rate', name: 'Best 弃电率 (%)', color: '#d29922', get: h => (h.bestCurtailmentRate === null ? null : h.bestCurtailmentRate * 100) },
        }[metric] || {
            title: '收敛曲线', name: '指标', color: '#58a6ff', get: h => h.bestFIRR,
        };

        const gens = (history || []).map(h => h.generation);
        const values = (history || []).map(h => {
            const v = meta.get(h);
            return (v === null || v === undefined || !isFinite(v)) ? null : Math.round(v * 1e4) / 1e4;
        });

        chart.setOption({
            backgroundColor: t.bg,
            title: { text: meta.title, left: 'center', textStyle: { color: t.text, fontSize: 13 } },
            tooltip: { trigger: 'axis' },
            grid: { left: 74, right: 30, top: 44, bottom: 40 },
            xAxis: { type: 'category', data: gens.length ? gens : [0], name: '代数', ...ax },
            yAxis: { type: 'value', name: meta.name, scale: true, ...ax },
            animation: !forExport,
            series: [{
                name: meta.name,
                type: 'line',
                data: values,
                connectNulls: true,
                symbol: 'circle',
                symbolSize: 4,
                itemStyle: { color: meta.color },
                lineStyle: { color: meta.color, width: 2 },
                areaStyle: { opacity: 0.10, color: meta.color },
            }],
        });

        return chart;
    },

    // ==================== 内部：图表初始化 ====================

    /** 初始化图表实例（在线预览用 dark 主题，导出用默认+白色背景） */
    /**
     * 获取（或复用）ECharts 实例。
     *
     * V2.3（任务书 §10）：在线预览路径**复用同一容器上的既有实例**，
     * 只在切换图表类型时 setOption，不再 dispose + 重建。
     * 复用前调用 clear() 清空旧 option，等价于 notMerge，避免残留上一次的系列。
     *
     * 导出路径（forExport=true）使用浅色主题且容器为一次性离屏节点，
     * 每次新建实例，由调用方负责 dispose。
     */
    _initChart(container, forExport) {
        if (!forExport && typeof echarts !== 'undefined' && echarts.getInstanceByDom) {
            const existing = echarts.getInstanceByDom(container);
            if (existing && !existing.isDisposed()) {
                existing.clear();
                return existing;
            }
        }
        // 导出时使用 vega/light 风格：不传 'dark' 主题，避免 ECharts 注入深色背景
        return echarts.init(container, forExport ? null : 'dark');
    },

    // ==================== 导出功能 ====================

    /** 导出当前在线图表为 PNG（白色背景） */
    exportChartPNG(chartInstance, filename) {
        const url = chartInstance.getDataURL({
            type: 'png', pixelRatio: 2,
            backgroundColor: '#ffffff',
        });
        Utils.downloadBase64(url.replace('data:image/png;base64,', ''), filename);
    },

    /**
     * 创建离屏渲染容器（固定尺寸，确保完整渲染）
     * 关键：必须挂载到 DOM 且有显式宽高，不能用 left:-9999px（会导致渲染异常）
     */
    _createOffscreenContainer(width = 1600, height = 600) {
        const div = document.createElement('div');
        div.style.cssText = `width:${width}px;height:${height}px;position:fixed;left:0;top:0;z-index:-1;opacity:0;pointer-events:none;`;
        document.body.appendChild(div);
        return div;
    },

    /**
     * 导出单个方案的全部图表为 ZIP
     * @param {Object} simResult - 单个方案的仿真结果
     * @param {string} schemeLabel - 方案显示名称（用于 ZIP 文件名）
     */
    async exportSchemeCharts(simResult, schemeLabel) {
        // V2.3：不再 parseResults 物化 8760 个对象。
        //   · 逐时类图 → 全分辨率 HourView（零对象分配）
        //   · 月度类图 → 月度视图（读 ResultDataStore 缓存）
        //   · 典型日/周 → 全分辨率区间视图（保证精确口径，不做降采样）
        const DS = (typeof ResultDataStore !== 'undefined') ? ResultDataStore : null;
        const fullView = DS
            ? DS.ChartDataAdapter.getFullView(simResult)
            : this.parseResults(simResult.results);   // 兜底：ResultDataStore 未加载
        const monthView = DS ? DS.ChartDataAdapter.getMonthlyView(simResult) : fullView;
        const data = fullView;
        const sv = simResult.systemVars;
        const suffix = `${sv['电解槽容量（MW）']}MW-${sv['储能功率（MW）']}MW-${sv['储能时长（小时）']}H`;

        const zip = new JSZip();
        // 用方案标识作为子文件夹，避免多方案混淆
        const folder = zip.folder(`图表_${suffix}`);

        // 离屏容器（固定尺寸，所有图共用一个容器，逐个渲染后销毁）
        const containerW = 1600, containerH = 600;
        const offDiv = this._createOffscreenContainer(containerW, containerH);

        try {
            // 每月随机典型日 / 典型周（共 24 张：12 日 + 12 周）
            const typicalDays = this.getRandomTypicalDays();
            const typicalWeeks = this.getRandomTypicalWeeks();

            const tasks = [];

            // 1. 逐时曲线 11 项（导出用全分辨率，保证报告图精度）
            const colKeys = ['pv', 'wind', 'total', 'charge', 'discharge', 'storage', 'hydrogen', 'export', 'import', 'curtailment', 'h2prod'];
            colKeys.forEach(key => {
                tasks.push({ name: `${this.COL_NAMES[key]}-${suffix}.png`, render: () => this.renderHourlyChart(offDiv, data, key, true) });
            });

            // 2. 全年能量流动汇总图
            tasks.push({ name: `能量流动汇总图-${suffix}.png`, render: () => this.renderOverviewChart(offDiv, data, true) });

            // 3. 逐月电量柱状图（9项）
            tasks.push({ name: `逐月电量柱状图-${suffix}.png`, render: () => this.renderMonthlyChart(offDiv, monthView, true) });

            // 4. 逐月电量汇总图（7项正负堆叠）
            tasks.push({ name: `逐月电量汇总图-${suffix}.png`, render: () => this.renderMonthlyOverviewChart(offDiv, monthView, true) });

            // 5. 典型日曲线 12 张（每月随机一天）
            typicalDays.forEach(({ month, dayStart }) => {
                tasks.push({
                    name: `典型日_${month + 1}月-${suffix}.png`,
                    render: () => this.renderTypicalDayChart(offDiv, data, month, dayStart, true)
                });
            });

            // 6. 典型周曲线 12 张（每月随机一周）
            typicalWeeks.forEach(({ month, weekStart }) => {
                tasks.push({
                    name: `典型周_${month + 1}月-${suffix}.png`,
                    render: () => this.renderTypicalWeekChart(offDiv, data, month, weekStart, true)
                });
            });

            // 7. 1-12 月逐月电量汇总图（宽幅 fill_between，12 张）
            for (let m = 0; m < 12; m++) {
                const mn = m;
                tasks.push({
                    name: `${mn + 1}月电量汇总图-${suffix}.png`,
                    render: () => this.renderMonthlyDetailChart(offDiv, data, mn, true),
                    size: [2000, 600],
                });
            }

            // 8. 分模块逐月柱状图（9 项，12 个月）
            const moduleKeys = ['pv', 'wind', 'discharge', 'import', 'charge', 'hydrogen', 'export', 'curtailment', 'h2prod'];
            moduleKeys.forEach(key => {
                const displayName = key === 'h2prod' ? '制氢量(吨)' : (this.COL_NAMES[key] || key);
                tasks.push({
                    name: `${displayName}逐月汇总-${suffix}.png`,
                    render: () => this.renderModuleMonthlyChart(offDiv, monthView, key, true),
                    size: [900, 500],
                });
            });

            // 逐个渲染并采集 PNG（forExport=true 已禁用动画，setOption 后立即可取图）
            for (const task of tasks) {
                // 支持 task 自定义尺寸，默认使用公共宽高
                const [w, h] = task.size || [containerW, containerH];
                offDiv.style.width = `${w}px`;
                offDiv.style.height = `${h}px`;
                const chart = task.render();
                chart.resize({ width: w, height: h });
                const url = chart.getDataURL({ type: 'png', pixelRatio: 2, backgroundColor: '#ffffff' });
                const base64 = url.replace('data:image/png;base64,', '');
                folder.file(task.name, base64, { base64: true });
                chart.dispose();
                // 让出主线程，避免长时间阻塞 UI
                await new Promise(r => setTimeout(r, 0));
            }
        } finally {
            document.body.removeChild(offDiv);
        }

        const blob = await zip.generateAsync({ type: 'blob' });
        saveAs(blob, `图表_${schemeLabel}_${Utils.timestamp()}.zip`);
    }
};

window.ChartModule = ChartModule;
