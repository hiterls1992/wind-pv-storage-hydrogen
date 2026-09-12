/**
 * 主程序入口
 * 管理全局状态、标签页切换、事件绑定、计算调度
 */

(function () {
    'use strict';

    // ========== 全局状态 ==========
    const AppState = {
        inputData: null,           // 原始Excel数据 [{pv, wind}, ...]
        inputDataRaw: null,        // ArrayBuffer for re-read
        simulationResults: [],     // 仿真结果数组
        summaryData: [],           // 数据汇总
        estimateData: [],          // 概算结果
        currentChart: null,        // 当前ECharts实例
        financeEstimateData: null, // 经评模块的概算结果（已含运营+投资字段）
        financeAllResults: {},     // 经评多方案结果 {schemeName: result}
        financeAllDetails: {},    // 经评多方案详情
        worker: null,              // Web Worker 实例

        // ---- V2.2 参数体系（任务书 §4 / §6 / §21）----
        // 唯一真值全部托管在 ParameterManager；此处只提供只读代理，
        // 禁止在 AppState 上另存一份容量参数或运行参数的真值。
        get currentScheme() { return ParameterManager.getCurrentScheme(); },
        get simulationConfig() { return ParameterManager.getSimulationConfig(); },
        get optimizationConfig() { return ParameterManager.getOptimizationConfig(); },
    };

    // ========== 日志系统 ==========
    function log(message, type = 'info') {
        const logContent = document.getElementById('logContent');
        const time = Utils.timeNow();
        const cls = type === 'error' ? 'log-error' : type === 'success' ? 'log-success' : type === 'warn' ? 'log-warn' : 'log-msg';
        logContent.innerHTML += `<div><span class="log-time">[${time}]</span> <span class="${cls}">${message}</span></div>`;
        logContent.scrollTop = logContent.scrollHeight;
    }

    function clearLog() {
        document.getElementById('logContent').innerHTML = '';
    }

    function setStatus(text) {
        document.getElementById('statusText').textContent = text;
    }

    // ========== 标签页切换 ==========
    function initTabs() {
        // 主标签页
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
                document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
                btn.classList.add('active');
                document.getElementById(`tab-${btn.dataset.tab}`).classList.add('active');
            });
        });

        // 仿真子标签页
        document.querySelectorAll('.sub-tab-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                document.querySelectorAll('.sub-tab-btn').forEach(b => b.classList.remove('active'));
                document.querySelectorAll('.sub-tab-panel').forEach(p => p.classList.remove('active'));
                btn.classList.add('active');
                document.getElementById(`subtab-${btn.dataset.subtab}`).classList.add('active');
            });
        });

        // 经评子标签页
        document.querySelectorAll('.finance-sub-tab').forEach(btn => {
            btn.addEventListener('click', () => {
                document.querySelectorAll('.finance-sub-tab').forEach(b => b.classList.remove('active'));
                document.querySelectorAll('.finance-sub-panel').forEach(p => p.classList.remove('active'));
                btn.classList.add('active');
                document.getElementById(`ftab-${btn.dataset.ftab}`).classList.add('active');
            });
        });
    }

    // ========== 仿真计算模块 ==========
    function initSimulation() {
        // 文件上传
        document.getElementById('inputFile').addEventListener('change', async function () {
            const file = this.files[0];
            if (!file) return;

            document.getElementById('inputFileName').textContent = file.name;
            log(`正在读取文件: ${file.name}`);

            try {
                const buffer = await file.arrayBuffer();
                AppState.inputDataRaw = buffer;
                AppState.inputData = ExcelIO.readInputExcel(buffer);

                const fileInfo = document.getElementById('fileInfo');
                fileInfo.textContent = `✅ 成功读取 ${AppState.inputData.length} 小时数据`;
                fileInfo.style.color = 'var(--accent-green)';

                log(`成功读取 ${AppState.inputData.length} 小时光伏/风电数据`, 'success');
                setStatus('数据已加载，可以开始仿真');
            } catch (err) {
                document.getElementById('fileInfo').textContent = `❌ 读取失败: ${err.message}`;
                document.getElementById('fileInfo').style.color = 'var(--accent-red)';
                log(`读取文件失败: ${err.message}`, 'error');
            }
        });

        // 开始仿真
        document.getElementById('btnRunSimulation').addEventListener('click', startSimulation);

        // 取消仿真
        document.getElementById('btnCancelSimulation').addEventListener('click', cancelSimulation);

        // 方案选择 → 显示数据
        document.getElementById('schemeSelect').addEventListener('change', function () {
            if (this.value !== '' && AppState.simulationResults.length > 0) {
                displaySimulationTable(parseInt(this.value));
                // 同步图表方案选择
                const chartSel = document.getElementById('chartSchemeSelect');
                chartSel.value = this.value;
            }
        });

        // 图表方案选择
        document.getElementById('chartSchemeSelect').addEventListener('change', function () {
            if (this.value !== '' && AppState.simulationResults.length > 0) {
                displaySimulationTable(parseInt(this.value));
                document.getElementById('schemeSelect').value = this.value;
            }
            updateChart();
        });

        // 图表类型选择
        document.getElementById('chartTypeSelect').addEventListener('change', function () {
            const colSelect = document.getElementById('chartColumnSelect');
            const monthSelect = document.getElementById('chartMonthSelect');
            // 显示列选择器：逐时曲线 或 分模块逐月柱状图
            colSelect.style.display = (this.value === 'hourly' || this.value === 'module-monthly') ? '' : 'none';
            // 显示月份选择器：典型日/周 或 单月电量汇总图
            const showMonth = this.value === 'typical-day' || this.value === 'typical-week' || this.value === 'monthly-detail';
            if (monthSelect) monthSelect.style.display = showMonth ? '' : 'none';
            updateChart();
        });

        document.getElementById('chartColumnSelect').addEventListener('change', updateChart);
        const _monthSel = document.getElementById('chartMonthSelect');
        if (_monthSel) _monthSel.addEventListener('change', updateChart);

        // 导出Excel
        document.getElementById('btnExportExcel').addEventListener('click', exportSingleExcel);
        document.getElementById('btnExportAllExcel').addEventListener('click', exportAllExcel);

        // 汇总
        document.getElementById('btnGenSummary').addEventListener('click', generateSummary);
        document.getElementById('btnExportSummary').addEventListener('click', exportSummary);

        // 图表导出（单张：当前在线图表）
        document.getElementById('btnExportChartPng').addEventListener('click', () => {
            if (AppState.currentChart) {
                const sel = document.getElementById('chartSchemeSelect');
                const schemeName = sel.options[sel.selectedIndex].text;
                const chartType = document.getElementById('chartTypeSelect').value;
                ChartModule.exportChartPNG(AppState.currentChart, `${schemeName}_${chartType}.png`);
            }
        });
        // 图表导出（当前方案全部图表 → ZIP）
        document.getElementById('btnExportAllCharts').addEventListener('click', exportCurrentSchemeCharts);
    }

    // ==================================================================
    // ========== 参数层（V2.2 参数体系重构核心，任务书 §20 / §21 / §23）==========
    // ==================================================================
    //
    //   HTML 输入框  ──syncSchemeFromUI()──▶  AppState.currentScheme  ──▶  仿真 / 优化 / 基准 / 结果
    //   HTML 输入框  ◀──syncSchemeToUI()───   AppState.currentScheme
    //
    //   任何模块都不允许再直接读写容量输入框的 DOM 值，必须走上面这两个函数。

    /**
     * 读取「运行参数」（设备怎么运行）—— 不含任何容量。
     * 容量一律来自 AppState.currentScheme（任务书 §5 / §12）。
     * 读取的同时写回 ParameterManager，保证「DOM 改动即真值改动」（§38）。
     */
    function getSimulationConfig() {
        return ParameterManager.syncSimulationConfigFromUI();
    }

    /**
     * 参数层初始化：把「当前方案」「运行参数」两组输入框接进唯一真值通道。
     * 首屏以 index.html 中的出厂默认值初始化真值（不写死在 JS 里，避免两份默认值漂移）。
     */
    function initParameterLayer() {
        // 1) 当前方案输入框 → 唯一真值（任务书 §38）
        ParameterManager.SCHEME_KEYS.forEach(key => {
            const el = document.getElementById(ParameterManager.SCHEME_DOM[key]);
            if (!el) return;
            el.addEventListener('input', () => {
                const res = ParameterManager.syncSchemeFromUI();
                if (!res.ok) return;             // 非法中间态（如空值）不写入真值
                updateSchemeStatusUI(res.scheme);
            });
            el.addEventListener('change', () => {
                const res = ParameterManager.syncSchemeFromUI();
                if (!res.ok) { log('方案参数非法：' + res.message, 'error'); return; }
                updateSchemeStatusUI(res.scheme);
            });
        });

        // 2) 运行参数输入框 → ParameterManager
        ParameterManager.SIM_KEYS.forEach(key => {
            const el = document.getElementById(ParameterManager.SIM_DOM[key]);
            if (el) el.addEventListener('change', () => ParameterManager.syncSimulationConfigFromUI());
        });

        // 3) 首屏：出厂默认值 → 唯一真值 → 回写状态栏与储能容量只读字段
        const first = ParameterManager.syncSchemeFromUI();
        if (!first.ok) {
            log('出厂默认方案参数异常（' + first.message + '），已回退到内置默认值', 'warn');
            ParameterManager.setCurrentScheme(ParameterManager.DEFAULT_SCHEME);
            ParameterManager.syncSchemeToUI();
        } else {
            ParameterManager.syncSchemeToUI(first.scheme);
        }
        ParameterManager.syncSimulationConfigFromUI();

        log(`参数体系就绪（V2.2）：当前方案 ${ParameterManager.schemeKey(ParameterManager.getCurrentScheme())}`);
    }

    /** 刷新「当前方案状态栏」，并同步优化页的基准方案显示与变量「当前值」列 */
    function updateSchemeStatusUI(scheme) {
        ParameterManager.renderSchemeStatus(scheme);
        refreshOptimizationBaselineUI();
    }

    /**
     * 开始仿真计算 —— V2.2：只计算「当前方案」这一个方案（任务书 §7 / §10）。
     *
     * 容量来自 AppState.currentScheme，运行规则来自 AppState.simulationConfig，
     * 二者分别对应 runSingleSimulation 的第 3、4 个参数，不再混在同一个对象里（§12 / §13）。
     * 多组容量组合的批量计算已迁至独立的「批量计算」页（§40）。
     */
    async function startSimulation() {
        if (!AppState.inputData || AppState.inputData.length === 0) {
            alert('请先选择并加载 input.xlsx 文件！');
            return;
        }

        // UI → 唯一真值（任务书 §10 / §38）
        const sres = ParameterManager.syncSchemeFromUI();
        if (!sres.ok) { alert('方案参数非法：\n' + sres.message); return; }
        const scheme = sres.scheme;
        const simulationConfig = getSimulationConfig();

        const vres = ParameterManager.validateSimulationConfig(simulationConfig);
        if (!vres.ok) log('运行参数提示：' + vres.message, 'warn');
        (sres.warnings || []).forEach(w => log('提示：' + w, 'warn'));

        log(`开始仿真计算：${ParameterManager.schemeLabel(scheme)}（方案编号 ${ParameterManager.schemeKey(scheme)}）`);

        // 转换为 TypedArray
        const pvArr = new Float64Array(AppState.inputData.length);
        const windArr = new Float64Array(AppState.inputData.length);
        for (let i = 0; i < AppState.inputData.length; i++) {
            pvArr[i] = AppState.inputData[i].pv;
            windArr[i] = AppState.inputData[i].wind;
        }

        // UI状态
        document.getElementById('btnRunSimulation').style.display = 'none';
        document.getElementById('btnCancelSimulation').style.display = '';
        document.getElementById('progressContainer').style.display = '';
        document.getElementById('workerStatus').classList.add('calculating');
        document.getElementById('workerStatusText').textContent = '计算中...';

        // 单方案 8760 小时仿真耗时在毫秒级，直接在主线程计算即可（V2.1 的批量分片调度与内联 Worker 代码已删除）
        setTimeout(() => {
            try {
                const result = runSingleSimulation(pvArr, windArr, scheme, simulationConfig);
                onSimulationComplete([result]);
            } catch (err) {
                log('仿真计算失败：' + err.message, 'error');
                resetSimulationUI();
                setStatus('计算失败');
            }
        }, 0);
    }

    /**
     * 取消计算。
     * V2.2 的单方案仿真在毫秒级完成，已不存在「分批可中断」的长任务；
     * 多组容量组合的批量计算页有各自的取消逻辑。此处仅保留为 UI 复位入口。
     */
    function cancelSimulation() {
        log('当前为单方案计算（毫秒级完成），无需取消', 'warn');
        resetSimulationUI();
    }

    function onSimulationComplete(results) {
        AppState.simulationResults = results;
        log(`仿真计算完成！共 ${results.length} 个方案`, 'success');

        // 填充方案选择器（方案编号统一取自结果自带的 scheme，§48 / §49）
        const sel = document.getElementById('schemeSelect');
        const chartSel = document.getElementById('chartSchemeSelect');
        sel.innerHTML = '';
        chartSel.innerHTML = '';

        results.forEach((r, i) => {
            const sv = r.systemVars;
            const label = r.scheme
                ? ParameterManager.schemeKey(r.scheme)
                : `方案${i + 1}`;
            const optText = `${label}: 电解槽${sv['电解槽容量（MW）']}MW, 储能${sv['储能功率（MW）']}MW×${sv['储能时长（小时）']}h`;
            sel.innerHTML += `<option value="${i}">${optText}</option>`;
            chartSel.innerHTML += `<option value="${i}">${optText}</option>`;
        });

        // 启用按钮
        document.getElementById('btnExportExcel').disabled = false;
        document.getElementById('btnExportAllExcel').disabled = false;
        document.getElementById('btnExportChartPng').disabled = false;
        document.getElementById('btnExportAllCharts').disabled = false;
        document.getElementById('btnGenSummary').disabled = false;

        // 显示第一个方案
        if (results.length > 0) {
            displaySimulationTable(0);
        }

        resetSimulationUI(true);
        setStatus(`计算完成，共 ${results.length} 个方案`);
    }

    function resetSimulationUI(completed = false) {
        document.getElementById('btnRunSimulation').style.display = '';
        document.getElementById('btnCancelSimulation').style.display = 'none';
        document.getElementById('workerStatus').classList.remove('calculating');
        document.getElementById('workerStatusText').textContent = completed ? '完成' : '就绪';

        if (completed) {
            document.getElementById('progressFill').style.width = '100%';
            document.getElementById('progressPercent').textContent = '100%';
        }
    }

    function displaySimulationTable(index) {
        const result = AppState.simulationResults[index];
        if (!result) return;

        const data = ChartModule.parseResults(result.results);
        const COLS = 11;
        // 列定义：表头、原始数据 key、是否求和、是否取末值
        // ⚠️ 原始 key 必须与 chart-module.js:75 parseResults() 完全一致
        // ⚠️ 单位：11 列电量字段均为 MWh（功率×1h），制氢量为 kg，储能现存容量为 MWh
        const dataCols = [
            { header: '光伏电量(MWh)',     key: '光伏电量',     sum: true  },
            { header: '风电电量(MWh)',     key: '风电电量',     sum: true  },
            { header: '合计电量(MWh)',     key: '合计电量',     sum: true  },
            { header: '储能充电电量(MWh)', key: '储能充电量',   sum: true  },
            { header: '储能放电电量(MWh)', key: '储能放电量',   sum: true  },
            { header: '储能现存容量(MWh)', key: '储能现存容量', sum: false, takeLast: true },
            { header: '制氢电量(MWh)',     key: '制氢电量',     sum: true  },
            { header: '上网电量(MWh)',     key: '上网电量',     sum: true  },
            { header: '下网电量(MWh)',     key: '下网电量',     sum: true  },
            { header: '弃电量(MWh)',       key: '弃电量',       sum: true  },
            { header: '制氢量(kg)',        key: '制氢量',       sum: true  }
        ];

        const table = document.getElementById('simulationTable');
        let html = '<thead><tr><th>小时</th>';
        for (const c of dataCols) html += `<th>${c.header}</th>`;
        html += '</tr></thead><tbody>';

        for (let h = 0; h < data.length; h++) {
            html += '<tr>';
            html += `<td>${h}</td>`;
            for (const c of dataCols) {
                html += `<td>${data[h][c.key].toFixed(2)}</td>`;
            }
            html += '</tr>';
        }

        // 总和行（与表头 11 列一一对应）
        const rawResults = result.results;
        const rows = rawResults.length / COLS;
        html += '<tr>';
        html += '<td><b>总和</b></td>';
        for (const c of dataCols) {
            let v;
            if (c.takeLast) {
                v = rawResults[(rows - 1) * COLS + colIndexOf(c.key)];
            } else {
                let s = 0;
                for (let h = 0; h < rows; h++) s += rawResults[h * COLS + colIndexOf(c.key)];
                v = s;
            }
            html += `<td><b>${v.toFixed(2)}</b></td>`;
        }
        html += '</tr>';

        html += '</tbody>';
        table.innerHTML = html;

        // 原始结果数组里的列索引（与 simulation-engine.js 列顺序完全一致）
        function colIndexOf(key) {
            const order = ['光伏电量', '风电电量', '合计电量', '储能充电量', '储能放电量',
                           '储能现存容量', '制氢电量', '上网电量', '下网电量', '弃电量', '制氢量'];
            return order.indexOf(key);
        }
    }

    function updateChart() {
        const chartContainer = document.getElementById('chartContainer');
        const schemeIdx = parseInt(document.getElementById('chartSchemeSelect').value);
        const chartType = document.getElementById('chartTypeSelect').value;

        if (isNaN(schemeIdx) || !AppState.simulationResults[schemeIdx]) return;

        // 销毁旧图表
        if (AppState.currentChart) {
            AppState.currentChart.dispose();
            AppState.currentChart = null;
        }

        chartContainer.innerHTML = '';
        const data = ChartModule.parseResults(AppState.simulationResults[schemeIdx].results);

        // 月份选择器（典型日/周使用）
        const monthSelect = document.getElementById('chartMonthSelect');
        const month = monthSelect ? parseInt(monthSelect.value) : 0;

        switch (chartType) {
            case 'overview':
                AppState.currentChart = ChartModule.renderOverviewChart(chartContainer, data);
                break;
            case 'hourly':
                const colKey = document.getElementById('chartColumnSelect').value;
                AppState.currentChart = ChartModule.renderHourlyChart(chartContainer, data, colKey);
                break;
            case 'monthly':
                AppState.currentChart = ChartModule.renderMonthlyChart(chartContainer, data);
                break;
            case 'monthly-overview':
                AppState.currentChart = ChartModule.renderMonthlyOverviewChart(chartContainer, data);
                break;
            case 'monthly-detail':
                AppState.currentChart = ChartModule.renderMonthlyDetailChart(chartContainer, data, month);
                break;
            case 'module-monthly': {
                const mColKey = document.getElementById('chartColumnSelect').value;
                AppState.currentChart = ChartModule.renderModuleMonthlyChart(chartContainer, data, mColKey);
                break;
            }
            case 'typical-day':
                AppState.currentChart = ChartModule.renderTypicalDayChart(chartContainer, data, month);
                break;
            case 'typical-week':
                AppState.currentChart = ChartModule.renderTypicalWeekChart(chartContainer, data, month);
                break;
        }
    }

    /** 导出当前选中方案的全部图表为 ZIP（白底 PNG） */
    async function exportCurrentSchemeCharts() {
        const schemeIdx = parseInt(document.getElementById('chartSchemeSelect').value);
        if (isNaN(schemeIdx) || !AppState.simulationResults[schemeIdx]) {
            alert('请先运行仿真计算并选择方案！');
            return;
        }

        const sel = document.getElementById('chartSchemeSelect');
        const schemeLabel = sel.options[sel.selectedIndex].text;
        const simResult = AppState.simulationResults[schemeIdx];

        log(`正在导出方案「${schemeLabel}」的全部图表（白底PNG）...`);
        document.getElementById('workerStatus').classList.add('calculating');
        document.getElementById('workerStatusText').textContent = '导出中...';

        try {
            await ChartModule.exportSchemeCharts(simResult, schemeLabel);
            log(`方案「${schemeLabel}」图表导出完成（逐时11项 + 汇总 + 月度柱 + 月度汇总 + 各月详图12 + 典型日12 + 典型周12 + 分模块月柱9）`, 'success');
        } catch (err) {
            log(`图表导出失败: ${err.message}`, 'error');
            console.error(err);
        } finally {
            document.getElementById('workerStatus').classList.remove('calculating');
            document.getElementById('workerStatusText').textContent = '就绪';
        }
    }

    // ========== Excel 导出 ==========
    async function exportSingleExcel() {
        const idx = parseInt(document.getElementById('schemeSelect').value);
        if (isNaN(idx)) return;

        const result = AppState.simulationResults[idx];
        const data = ChartModule.parseResults(result.results);

        // 构造hourlyData格式
        const hourlyData = data.map(row => {
            const obj = {};
            for (const k of Object.keys(row)) obj[k] = row[k];
            return obj;
        });

        const resultData = {
            hourlyData,
            systemVars: result.systemVars,
            ratioData: result.ratioData
        };

        const buffer = await ExcelIO.exportSimulationResult(resultData);
        const blob = new Blob([buffer], { type: 'application/octet-stream' });
        Utils.downloadBlob(blob, result.filename);
        log(`已导出: ${result.filename}`, 'success');
    }

    async function exportAllExcel() {
        if (AppState.simulationResults.length === 0) return;

        const zip = new JSZip();
        const folder = zip.folder('simulation_results');

        for (const result of AppState.simulationResults) {
            const data = ChartModule.parseResults(result.results);
            const hourlyData = data.map(row => { const obj = {}; for (const k of Object.keys(row)) obj[k] = row[k]; return obj; });
            const buffer = await ExcelIO.exportSimulationResult({ hourlyData, systemVars: result.systemVars, ratioData: result.ratioData });
            folder.file(result.filename, buffer);
        }

        const blob = await zip.generateAsync({ type: 'blob' });
        saveAs(blob, `仿真结果批量_${Utils.timestamp()}.zip`);
        log('已批量导出所有仿真结果Excel', 'success');
    }

    function generateSummary() {
        if (AppState.simulationResults.length === 0) return;

        // 容量取自每个结果自带的 scheme（§48），不再由调用方另外传一份容量造成第二处真值
        AppState.summaryData = DataSummary.generateSummary(AppState.simulationResults, AppState.currentScheme);

        // 显示汇总表
        const table = document.getElementById('summaryTable');
        if (AppState.summaryData.length === 0) {
            table.innerHTML = '<thead><tr><th>无数据</th></tr></thead><tbody></tbody>';
            return;
        }

        const keys = Object.keys(AppState.summaryData[0]).filter(k => !k.startsWith('_'));
        let html = '<thead><tr>';
        for (const k of keys) html += `<th>${k}</th>`;
        html += '</tr></thead><tbody>';

        for (const row of AppState.summaryData) {
            html += '<tr>';
            for (const k of keys) {
                html += `<td>${row[k]}</td>`;
            }
            html += '</tr>';
        }
        html += '</tbody>';
        table.innerHTML = html;

        document.getElementById('btnExportSummary').disabled = false;
        log(`数据汇总完成，共 ${AppState.summaryData.length} 个方案`, 'success');
    }

    async function exportSummary() {
        if (AppState.summaryData.length === 0) return;

        const cleanData = AppState.summaryData.map(row => {
            const clean = {};
            for (const k of Object.keys(row)) {
                if (!k.startsWith('_')) clean[k] = row[k];
            }
            return clean;
        });

        const buffer = await ExcelIO.exportSummaryExcel(cleanData);
        const blob = new Blob([buffer], { type: 'application/octet-stream' });
        Utils.downloadBlob(blob, `数据汇总${Utils.timestamp()}.xlsx`);
        log('已导出数据汇总Excel', 'success');
    }

    // ========== 概算模块 ==========
    function initEstimate() {
        document.getElementById('estimateDataFile').addEventListener('change', async function () {
            const file = this.files[0];
            if (!file) return;

            document.getElementById('estimateDataFileName').textContent = file.name;
            log(`概算模块: 正在读取 ${file.name}`);

            try {
                const buffer = await file.arrayBuffer();
                const data = ExcelIO.readDataSummaryExcel(buffer);
                AppState.estimateSourceData = data;
                log(`概算模块: 成功读取 ${data.length} 条方案数据`, 'success');
            } catch (err) {
                log(`概算模块: 读取失败 - ${err.message}`, 'error');
            }
        });

        document.getElementById('btnRunEstimate').addEventListener('click', runEstimate);
        document.getElementById('btnExportEstimate').addEventListener('click', exportEstimate);
    }

    function getEstimatePrices() {
        return {
            windUnitPrice: Utils.toNum(document.getElementById('windUnitPrice').value, 3800),
            pvUnitPrice: Utils.toNum(document.getElementById('pvUnitPrice').value, 2500),
            storageUnitPrice: Utils.toNum(document.getElementById('storageUnitPrice').value, 800),
            electrolyzerUnitPrice: Utils.toNum(document.getElementById('electrolyzerUnitPrice').value, 1600),
            transmissionCost: Utils.toNum(document.getElementById('transmissionCost').value, 10000),
            landCost: Utils.toNum(document.getElementById('landCost').value, 10000),
            otherFacilitiesRatio: Utils.toNum(document.getElementById('otherFacilitiesRatio').value, 70),
        };
    }

    function runEstimate() {
        if (!AppState.estimateSourceData || AppState.estimateSourceData.length === 0) {
            alert('请先选择数据汇总文件！');
            return;
        }

        const prices = getEstimatePrices();

        AppState.estimateData = Estimate.batchEstimate(AppState.estimateSourceData, prices);
        log(`概算分析完成，共 ${AppState.estimateData.length} 个方案`, 'success');

        // 显示表格
        const table = document.getElementById('estimateTable');
        if (AppState.estimateData.length === 0) {
            table.innerHTML = '<thead><tr><th>无数据</th></tr></thead><tbody></tbody>';
            return;
        }

        const keys = Object.keys(AppState.estimateData[0]);
        let html = '<thead><tr>';
        for (const k of keys) html += `<th>${k}</th>`;
        html += '</tr></thead><tbody>';

        for (const row of AppState.estimateData) {
            html += '<tr>';
            for (const k of keys) html += `<td>${row[k]}</td>`;
            html += '</tr>';
        }
        html += '</tbody>';
        table.innerHTML = html;

        document.getElementById('btnExportEstimate').disabled = false;
        setStatus(`概算完成，共 ${AppState.estimateData.length} 个方案`);
    }

    async function exportEstimate() {
        if (AppState.estimateData.length === 0) return;
        const buffer = await ExcelIO.exportEstimateResult(AppState.estimateData);
        const blob = new Blob([buffer], { type: 'application/octet-stream' });
        Utils.downloadBlob(blob, `概算结果_${Utils.timestamp()}.xlsx`);
        log('已导出概算结果Excel', 'success');
    }

    // ========== 经评模块 ==========
    function initFinance() {
        // 概算结果文件（已包含数据汇总的运营字段，无需另选汇总文件）
        document.getElementById('financeDataFile').addEventListener('change', async function () {
            const file = this.files[0];
            if (!file) return;
            document.getElementById('financeDataFileName').textContent = file.name;
            log(`经评模块: 正在读取概算结果 ${file.name}`);
            try {
                const buffer = await file.arrayBuffer();
                AppState.financeEstimateData = ExcelIO.readEstimateExcel(buffer);
                document.getElementById('btnLoadFinance').disabled = false;
                log(`经评模块: 成功读取 ${AppState.financeEstimateData.length} 条方案数据（含运营+投资）`, 'success');
            } catch (err) {
                log(`经评模块: 读取失败 - ${err.message}`, 'error');
            }
        });

        document.getElementById('btnLoadFinance').addEventListener('click', loadFinanceSchemes);
        document.getElementById('btnRunFinance').addEventListener('click', runFinanceSingle);
        document.getElementById('btnBatchFinance').addEventListener('click', runFinanceBatch);
        document.getElementById('btnExportFinance').addEventListener('click', exportFinanceResult);
        document.getElementById('btnExportCashflow').addEventListener('click', exportCashflow);
        document.getElementById('btnExportBatchFinance').addEventListener('click', exportBatchFinance);
    }

    function loadFinanceSchemes() {
        if (!AppState.financeEstimateData || AppState.financeEstimateData.length === 0) {
            alert('请先选择概算结果文件！');
            return;
        }

        const sel = document.getElementById('financeSchemeSelect');
        sel.innerHTML = '';

        // 概算结果文件已包含全部运营字段（来自上一步数据汇总）+ 投资字段（概算计算）
        // 因此经评模块只需读取这一个文件即可，无需再合并数据汇总文件
        AppState.financeMergedData = AppState.financeEstimateData;

        AppState.financeMergedData.forEach((row, i) => {
            const name = `方案${i + 1}: ${row['文件名称']}`;
            sel.innerHTML += `<option value="${i}">${name}</option>`;
        });

        document.getElementById('financeSchemeArea').style.display = '';
        document.getElementById('btnRunFinance').disabled = false;
        document.getElementById('btnBatchFinance').disabled = false;

        log(`经评模块: 已加载 ${AppState.financeMergedData.length} 个方案`, 'success');
    }

    function getFinanceParams() {
        return {
            calc_period: parseInt(document.getElementById('calcPeriod').value) || 22,
            construct_period: parseInt(document.getElementById('constructPeriod').value) || 2,
            batch_count: parseInt(document.getElementById('batchCount').value) || 1,
            capital_benchmark: Utils.toNum(document.getElementById('capitalBenchmark').value, 8.0),
            industry_benchmark_pre: Utils.toNum(document.getElementById('industryBenchmarkPre').value, 5.0),
            industry_benchmark_post: Utils.toNum(document.getElementById('industryBenchmarkPost').value, 4.5),
            capital_ratio: Utils.toNum(document.getElementById('capitalRatio').value, 20),
            loan_period: parseInt(document.getElementById('loanPeriod').value) || 15,
            long_term_loan_rate: Utils.toNum(document.getElementById('longTermLoanRate').value, 3.0),
            single_electrolyzer_mw: Utils.toNum(document.getElementById('singleElectrolyzer').value, 5.0),
            capacity_elec_fee: Utils.toNum(document.getElementById('capacityElecFee').value, 20.0),
            ongrid_price: Utils.toNum(document.getElementById('ongridPrice').value, 0.2),
            offgrid_price: Utils.toNum(document.getElementById('offgridPrice').value, 0.7),
            wheeling_fee: Utils.toNum(document.getElementById('wheelingFee').value, 0.12),
            h2_labor_cost: Utils.toNum(document.getElementById('h2LaborCost').value, 3.0),
            electrolyzer_maint: Utils.toNum(document.getElementById('electrolyzerMaint').value, 40.0),
            electrolyzer_overhaul: Utils.toNum(document.getElementById('electrolyzerOverhaul').value, 240.0),
            water_unit_price: Utils.toNum(document.getElementById('waterUnitPrice').value, 2.0),
            wind_pv_om: Utils.toNum(document.getElementById('windPvOm').value, 35.0),
            depreciation_years: parseInt(document.getElementById('depreciationYears').value) || 20,
            residual_rate: Utils.toNum(document.getElementById('residualRate').value, 5.0),
            h2_price: Utils.toNum(document.getElementById('h2Price').value, 30.0),
            output_vat_rate: Utils.toNum(document.getElementById('outputVatRate').value, 13.0),
            income_tax_rate: Utils.toNum(document.getElementById('incomeTaxRate').value, 25.0),
            surtax_rate: Utils.toNum(document.getElementById('surtaxRate').value, 12.0),
        };
    }

    function runFinanceSingle() {
        if (!AppState.financeMergedData || AppState.financeMergedData.length === 0) {
            alert('请先加载方案数据！');
            return;
        }

        const params = getFinanceParams();
        const idx = parseInt(document.getElementById('financeSchemeSelect').value);
        const row = AppState.financeMergedData[idx];
        const schemeName = `方案${idx + 1}: ${row['文件名称']}`;

        try {
            const { result, detail } = FinanceEngine.calculateAll(params, row, row);
            AppState.financeAllResults[schemeName] = result;
            AppState.financeAllDetails[schemeName] = detail;
            AppState.lastFinanceResult = result;
            AppState.lastFinanceDetail = detail;

            displayFinanceResults();
            document.getElementById('btnExportFinance').disabled = false;
            document.getElementById('btnExportCashflow').disabled = false;

            log(`经评: 「${schemeName}」经济评价计算完成`, 'success');
        } catch (err) {
            log(`经评计算失败: ${err.message}`, 'error');
            console.error(err);
        }
    }

    function runFinanceBatch() {
        if (!AppState.financeMergedData || AppState.financeMergedData.length === 0) {
            alert('请先加载方案数据！');
            return;
        }

        const params = getFinanceParams();
        AppState.financeAllResults = {};
        AppState.financeAllDetails = {};
        const batchRows = [];

        log(`经评: 开始批量分析 ${AppState.financeMergedData.length} 个方案...`);

        for (let i = 0; i < AppState.financeMergedData.length; i++) {
            const row = AppState.financeMergedData[i];
            const schemeName = `方案${i + 1}: ${row['文件名称']}`;

            try {
                const { result, detail } = FinanceEngine.calculateAll(params, row, row);
                AppState.financeAllResults[schemeName] = result;
                AppState.financeAllDetails[schemeName] = detail;

                batchRows.push({
                    '方案': schemeName,
                    '光伏容量（MW）': Utils.toNum(row['光伏容量（MW）']),
                    '风电容量（MW）': Utils.toNum(row['风电容量（MW）']),
                    '电解槽容量（MW）': Utils.toNum(row['电解槽容量（MW）']),
                    '储能功率（MW）': Utils.toNum(row['储能功率（MW）']),
                    '储能时长（小时）': Utils.toNum(row['储能时长（小时）']),
                    '制氢量总和（万吨）': Utils.toNum(row['制氢量总和（万吨）']),
                    '建设投资（万元）': result['建设投资（万元）'],
                    '项目总投资（万元）': result['项目总投资（万元）'],
                    '利润总额（万元）': result['利润总额（万元）'],
                    '项目投资IRR-税前（%）': result['项目投资财务内部收益率（所得税前）（%）'],
                    '项目投资IRR-税后（%）': result['项目投资财务内部收益率（所得税后）（%）'],
                    '资本金IRR（%）': result['资本金财务内部收益率（%）'],
                    '投资回收期-税前（年）': result['项目投资回收期（所得税前）（年）'],
                    '投资回收期-税后（年）': result['项目投资回收期（所得税后）（年）'],
                    'ROI（%）': result['总投资收益率（ROI）（%）'],
                    'ROE（%）': result['项目资本金净利润率（ROE）（%）'],
                    '盈亏平衡点（%）': result['盈亏平衡点（生产能力利用率）（%）'],
                    'EVA（万元）': result['经济增加值（EVA）（万元）'],
                    '制氢成本（元/kg）': result['制氢成本（元/kg）'],
                });
            } catch (err) {
                log(`方案${i + 1}计算失败: ${err.message}`, 'error');
            }
        }

        // 按IRR税后降序排序
        batchRows.sort((a, b) => b['项目投资IRR-税后（%）'] - a['项目投资IRR-税后（%）']);
        AppState.financeBatchData = batchRows;

        displayFinanceResults();
        document.getElementById('btnExportFinance').disabled = false;
        document.getElementById('btnExportCashflow').disabled = false;
        document.getElementById('btnExportBatchFinance').disabled = false;

        log(`经评: 批量分析完成，共 ${batchRows.length} 个方案`, 'success');
    }

    function displayFinanceResults() {
        const table = document.getElementById('financeTable');
        const allResults = AppState.financeAllResults;
        const schemeNames = Object.keys(allResults);

        if (schemeNames.length === 0) {
            table.innerHTML = '<thead><tr><th>暂无结果</th></tr></thead><tbody></tbody>';
            return;
        }

        // 指标定义
        const indicators = [
            ['项目总投资', '万元'], ['建设投资', '万元'], ['建设期利息', '万元'], ['流动资金', '万元'],
            ['销售收入总额（不含增值税）', '万元'], ['总成本费用', '万元'], ['销售税金附加总额', '万元'], ['利润总额', '万元'],
            ['项目投资回收期（所得税前）', '年'], ['项目投资回收期（所得税后）', '年'],
            ['项目投资财务内部收益率（所得税前）', '%'], ['项目投资财务内部收益率（所得税后）', '%'],
            ['项目投资财务净现值（所得税前）', '万元'], ['项目投资财务净现值（所得税后）', '万元'],
            ['资本金财务内部收益率', '%'], ['资本金财务净现值', '万元'],
            ['总投资收益率（ROI）', '%'], ['投资利税率', '%'], ['项目资本金净利润率（ROE）', '%'],
            ['资产负债率（最大值）', '%'], ['净资产收益率', '%'], ['营业现金比率', '%'],
            ['盈亏平衡点（生产能力利用率）', '%'], ['经济增加值（EVA）', '万元'], ['制氢成本', '元/kg'],
        ];

        let html = '<thead><tr><th>指标名称</th>';
        for (const name of schemeNames) {
            const short = name.length > 18 ? name.substring(0, 15) + '...' : name;
            html += `<th>${short}</th>`;
        }
        html += '</tr></thead><tbody>';

        for (const [indName, unit] of indicators) {
            html += `<tr><td>${indName}（${unit}）</td>`;

            for (const sName of schemeNames) {
                const result = allResults[sName];
                let val = findIndicatorValue(result, indName, unit);

                if (val !== null && val !== undefined) {
                    if (unit === '%') {
                        html += `<td>${Utils.formatPercent(val)}</td>`;
                    } else {
                        html += `<td>${Utils.formatNumber(val)}</td>`;
                    }
                } else {
                    html += '<td>N/A</td>';
                }
            }
            html += '</tr>';
        }

        html += '</tbody>';
        table.innerHTML = html;
    }

    function findIndicatorValue(result, indicatorName, unit) {
        const fullKey = `${indicatorName}（${unit}）`;
        if (fullKey in result) return result[fullKey];
        for (const k of Object.keys(result)) {
            if (k.startsWith(indicatorName)) return result[k];
        }
        return null;
    }

    async function exportFinanceResult() {
        if (!AppState.lastFinanceResult) return;
        const buffer = await ExcelIO.exportFinanceResult(AppState.lastFinanceResult);
        const blob = new Blob([buffer], { type: 'application/octet-stream' });
        Utils.downloadBlob(blob, `经评结果_${Utils.timestamp()}.xlsx`);
        log('已导出经评结果Excel', 'success');
    }

    async function exportCashflow() {
        if (!AppState.lastFinanceDetail) return;
        const { project_cf, equity_cf } = AppState.lastFinanceDetail;
        const buffer = await ExcelIO.exportCashflow(project_cf, equity_cf);
        const blob = new Blob([buffer], { type: 'application/octet-stream' });
        Utils.downloadBlob(blob, `经评现金流量表_${Utils.timestamp()}.xlsx`);
        log('已导出现金流量表Excel', 'success');
    }

    async function exportBatchFinance() {
        if (!AppState.financeBatchData) return;
        const buffer = await ExcelIO.exportBatchFinanceResult(AppState.financeBatchData);
        const blob = new Blob([buffer], { type: 'application/octet-stream' });
        Utils.downloadBlob(blob, `经评批量结果_${Utils.timestamp()}.xlsx`);
        log('已导出经评批量结果Excel', 'success');
    }

    // ==================================================================
    // ========== 方案优化模块（V2.1 新增 · NSGA-II 多目标容量优化） ==========
    // ==================================================================

    /** 优化页运行时状态 */
    const OptState = {
        running: false,
        cancelled: false,
        mode: 'idle',          // 'idle' | 'worker' | 'main'
        worker: null,
        session: null,
        result: null,
        liveHistory: [],
        selectedKey: null,
        paretoChart: null,
        convCharts: [],
        lastSnap: null,
        fallbackTimer: null
    };

    /** 优化变量 DOM 映射（顺序与 OptimizationEngine.KEY_ORDER 一致） */
    const OPT_VARS = [
        { key: 'windCapacity',         min: 'optWindMin',         max: 'optWindMax',         step: 'optWindStep',         levels: 'optWindLevels' },
        { key: 'pvCapacity',           min: 'optPvMin',           max: 'optPvMax',           step: 'optPvStep',           levels: 'optPvLevels' },
        { key: 'storagePower',         min: 'optStoragePowerMin', max: 'optStoragePowerMax', step: 'optStoragePowerStep', levels: 'optStoragePowerLevels' },
        { key: 'storageDuration',      min: 'optStorageDurationMin', max: 'optStorageDurationMax', step: 'optStorageDurationStep', levels: 'optStorageDurationLevels' },
        { key: 'electrolyzerCapacity', min: 'optElectrolyzerMin', max: 'optElectrolyzerMax', step: 'optElectrolyzerStep', levels: 'optElectrolyzerLevels' },
    ];

    /** 变量「当前值」列的 DOM 映射（只读，数据来自 AppState.currentScheme，§26） */
    const OPT_CURRENT_DOM = {
        windCapacity: 'optCurrentWind',
        pvCapacity: 'optCurrentPv',
        storagePower: 'optCurrentStoragePower',
        storageDuration: 'optCurrentStorageDuration',
        electrolyzerCapacity: 'optCurrentElectrolyzer',
    };

    /** 数值格式化（非有限值统一显示为 —，绝不出现 Infinity / NaN 字样） */
    function optNum(v, d) {
        if (v === null || v === undefined || !isFinite(Number(v))) return '—';
        return Number(v).toFixed(d === undefined ? 2 : d);
    }
    function optInt(v) {
        if (v === null || v === undefined || !isFinite(Number(v))) return '—';
        return Math.round(Number(v)).toLocaleString('zh-CN');
    }
    function optPct(v, d) {
        if (v === null || v === undefined || !isFinite(Number(v))) return '—';
        return Number(v).toFixed(d === undefined ? 2 : d) + '%';
    }
    function optLog(msg, type) { log('【优化】' + msg, type || 'info'); }
    function optEl(id) { return document.getElementById(id); }

    // -------------------- 参数读取 --------------------

    function readOptVariableConfig() {
        const out = {};
        for (const v of OPT_VARS) {
            out[v.key] = {
                min: Utils.toNum(optEl(v.min).value, NaN),
                max: Utils.toNum(optEl(v.max).value, NaN),
                step: Utils.toNum(optEl(v.step).value, NaN),
            };
        }
        return out;
    }

    function readOptConstraints() {
        const rd = (onId, valId, unit) => ({
            enabled: !!optEl(onId).checked,
            value: Utils.toNum(optEl(valId).value, NaN),
            unit: unit,
        });
        return {
            minAnnualHydrogen:     rd('consHydrogenOn',    'consHydrogen',    '万吨/年'),
            maxCurtailmentRate:    rd('consCurtailOn',     'consCurtail',     '%'),
            maxGridImportRatio:    rd('consImportOn',      'consImport',      '%'),
            minGreenHydrogenRatio: rd('consGreenOn',       'consGreen',       '%'),
            maxExportRatio:        rd('consExportOn',      'consExport',      '%'),
            maxExportPower:        rd('consExportPowerOn', 'consExportPower', 'MW'),
            maxImportPower:        rd('consImportPowerOn', 'consImportPower', 'MW'),
        };
    }

    function readOptNsgaConfig() {
        const useFixedSeed = !!optEl('optUseFixedSeed').checked;
        const seedInput = Utils.toNum(optEl('optRandomSeed').value, 20260912);
        return {
            populationSize: Math.round(Utils.toNum(optEl('optPopulationSize').value, 60)),
            generations: Math.round(Utils.toNum(optEl('optGenerations').value, 100)),
            crossoverProbability: Utils.toNum(optEl('optCrossoverProb').value, 0.9),
            mutationProbability: Utils.toNum(optEl('optMutationProb').value, 0.1),
            // 取消勾选「固定随机种子」时，每次优化使用新种子（便于验证算法稳健性）
            randomSeed: useFixedSeed ? Math.round(seedInput) : (Date.now() % 2147483647),
            earlyStopping: !!optEl('optEarlyStopping').checked,
            patience: Math.round(Utils.toNum(optEl('optPatience').value, 15)),
        };
    }

    function readOptWeights() {
        return {
            firr: Utils.toNum(optEl('optWeightFirr').value, 40) / 100,
            lcoh: Utils.toNum(optEl('optWeightLcoh').value, 35) / 100,
            curtailmentRate: Utils.toNum(optEl('optWeightCurtail').value, 25) / 100,
        };
    }

    /**
     * 基准方案唯一来源：AppState.currentScheme（任务书 §16 / §17 / §18）。
     *
     * V2.1 存在两处问题，均已删除：
     *   ① baseWind / basePv / baseStoragePower / baseStorageDuration / baseElectrolyzer
     *      一整套独立输入框——优化页自己维护了第二套「当前方案」；
     *   ② 把 storagePowerMin / storageDurationMin / electrolyzerMin（**扫描范围的下界**）
     *      当成基准值读取——语义错误。
     */
    function readBaselineScheme() {
        return ParameterManager.getCurrentScheme();
    }

    /** 权重合计实时显示 */
    function updateOptWeightSum() {
        const w = readOptWeights();
        const sum = (w.firr + w.lcoh + w.curtailmentRate) * 100;
        const el = optEl('optWeightSum');
        el.textContent = Math.round(sum * 100) / 100 + '%';
        el.style.color = Math.abs(sum - 100) < 1e-6 ? 'var(--accent-cyan)' : 'var(--accent-yellow)';
    }

    /** 变量范围 → 档位数 / 搜索空间 / 当前值 / 范围外提示（任务书 §26 / §27） */
    function updateOptVarPreview() {
        const variables = readOptVariableConfig();
        const scheme = ParameterManager.getCurrentScheme();
        let space = 1;
        for (const v of OPT_VARS) {
            const cfg = variables[v.key];
            let n = 0;
            if (isFinite(cfg.min) && isFinite(cfg.max)) {
                n = OptimizationEngine.buildLevelsFor(cfg.min, cfg.max, cfg.step).length;
            }
            optEl(v.levels).textContent = n > 0 ? n : '—';
            space *= Math.max(1, n);

            // 「当前值」列：只读显示，来自 AppState.currentScheme（§26）
            const curEl = optEl(OPT_CURRENT_DOM[v.key]);
            if (curEl) {
                const cur = scheme[v.key];
                const unit = ParameterManager.SCHEME_UNITS[v.key];
                const inRange = isFinite(cfg.min) && isFinite(cfg.max) && cur >= cfg.min && cur <= cfg.max;
                curEl.textContent = String(cur);
                curEl.classList.toggle('is-out-range', !inRange);
                curEl.title = ParameterManager.SCHEME_LABELS[v.key] + ' 当前值 ' + cur + unit +
                    '，优化范围 ' + cfg.min + ' ~ ' + cfg.max + unit +
                    (inRange ? '（在范围内）' : '（超出范围）');
            }
        }
        const pop = Math.round(Utils.toNum(optEl('optPopulationSize').value, 60));
        optEl('optSearchSpace').innerHTML =
            `搜索空间：<b>${space.toLocaleString('zh-CN')}</b> 个容量组合` +
            (space < pop ? `（小于种群规模 ${pop}，将自动收缩有效种群规模）` : '');

        renderOptRangeWarn(scheme, variables);
    }

    /**
     * 「当前基准方案是否落在优化搜索范围内」提示（任务书 §27 / §28）。
     * 只提示，**绝不自动修改**用户输入的任何数值。
     */
    function renderOptRangeWarn(scheme, variables) {
        const el = optEl('optRangeWarn');
        if (!el) return;
        const res = ParameterManager.checkBaselineInRange(scheme, { variables: variables });
        if (res.inRange) {
            el.style.display = 'none';
            el.innerHTML = '';
            return;
        }
        const out = res.items.filter(it => !it.inRange);
        el.style.display = '';
        el.innerHTML =
            '<span class="opt-range-warn-title">⚠ 当前基准方案不在优化搜索空间内</span>' +
            '<ul>' + out.map(it => `<li class="is-out">${it.text}</li>`).join('') + '</ul>' +
            '<button class="btn btn-sm btn-outline" id="btnExpandOptRange" type="button">将优化范围包含基准方案</button>' +
            '<span class="hint" style="margin-left:8px">可选操作；不点击则不会改动任何数值</span>';
    }

    /** 「将优化范围包含基准方案」（§27 可选功能）：只扩张区间，不改步长，且必须用户主动点击 */
    function expandOptRangeToBaseline() {
        const scheme = ParameterManager.getCurrentScheme();
        const next = ParameterManager.expandRangeToIncludeBaseline(scheme, { variables: readOptVariableConfig() });
        for (const v of OPT_VARS) {
            optEl(v.min).value = next.variables[v.key].min;
            optEl(v.max).value = next.variables[v.key].max;
        }
        updateOptVarPreview();
        optLog('已将优化范围扩张到包含基准方案（基准方案本身未做任何修改）', 'success');
    }

    /**
     * 刷新优化页的「基准方案」显示与变量「当前值」列。
     * V2.2：基准方案不再有一套独立输入框，直接显示 AppState.currentScheme（§18 / §39）。
     */
    function refreshOptimizationBaselineUI() {
        const scheme = ParameterManager.getCurrentScheme();
        ParameterManager.renderSchemeStatus(scheme);
        updateOptVarPreview();
    }

    // -------------------- 上下文组装 --------------------

    function collectOptimizationContext() {
        if (!AppState.inputData || AppState.inputData.length === 0) {
            throw new Error('请先在「电量计算」页加载 input.xlsx 风光 8760 小时数据');
        }
        const n = AppState.inputData.length;
        const pvData = new Float64Array(n);
        const windData = new Float64Array(n);
        for (let i = 0; i < n; i++) {
            pvData[i] = AppState.inputData[i].pv;
            windData[i] = AppState.inputData[i].wind;
        }
        // 字段名必须与 OptimizationEngine 的上下文约定一致（pvData / windData / simulationConfig）
        // 任务书 §35 / §36：上下文中只允许出现「数据 + 运行参数 + 单价 + 财务参数」，
        // 容量一律通过 evaluateScheme 的 scheme 参数传入，禁止再经由上下文夹带。
        return {
            pvData: pvData,
            windData: windData,
            simulationConfig: getSimulationConfig(),
            prices: getEstimatePrices(),
            financeParams: getFinanceParams(),
            lcohDiscountRate: Utils.toNum(optEl('optLcohDiscount').value, 5.0),
        };
    }

    // -------------------- 启动 / 停止 --------------------

    function initOptimization() {
        // 实时预览
        for (const v of OPT_VARS) {
            for (const id of [v.min, v.max, v.step]) {
                const el = optEl(id);
                if (el) el.addEventListener('input', updateOptVarPreview);
            }
        }
        ['optWeightFirr', 'optWeightLcoh', 'optWeightCurtail'].forEach(id => {
            const el = optEl(id);
            if (el) el.addEventListener('input', updateOptWeightSum);
        });
        optEl('optPopulationSize').addEventListener('input', updateOptVarPreview);
        updateOptWeightSum();
        updateOptVarPreview();

        // 按钮
        optEl('btnRunOptimization').addEventListener('click', startOptimization);
        optEl('btnStopOptimization').addEventListener('click', stopOptimization);
        // V2.2：基准方案 ＝「电量计算」页的当前方案，此按钮只做页面跳转，不再复制一套输入框（§16 / §18）
        optEl('btnLoadBaseline').addEventListener('click', () => {
            const tabBtn = document.querySelector('.tab-btn[data-tab="simulation"]');
            if (tabBtn) tabBtn.click();
            optLog('已切换到「电量计算」页：修改当前方案后，基准方案会自动同步', 'info');
        });
        optEl('btnExportOptimization').addEventListener('click', exportOptimizationExcel);

        // 「将优化范围包含基准方案」按钮（内容动态渲染，使用事件委托）
        optEl('optRangeWarn').addEventListener('click', (e) => {
            if (e.target && e.target.id === 'btnExpandOptRange') expandOptRangeToBaseline();
        });
        optEl('btnLoadRecommendedToSim').addEventListener('click', () => {
            if (!OptState.result || !OptState.result.representativeSolutions.recommended) return;
            applySchemeToSimulation(OptState.result.representativeSolutions.recommended, false);
        });
        optEl('btnView8760').addEventListener('click', () => {
            if (!OptState.result || !OptState.result.representativeSolutions.recommended) return;
            applySchemeToSimulation(OptState.result.representativeSolutions.recommended, true);
        });

        // Pareto 表排序 / 筛选
        optEl('optParetoSort').addEventListener('change', () => renderParetoTable());
        optEl('optParetoFilter').addEventListener('change', () => renderParetoTable());

        // Pareto 图切换
        optEl('optParetoChartType').addEventListener('change', () => renderParetoChart());
        optEl('btnExportOptChartPng').addEventListener('click', () => {
            if (OptState.paretoChart) ChartModule.exportChartPNG(OptState.paretoChart, 'Pareto前沿.png');
        });

        // 推荐方案卡片内的操作按钮（事件委托）
        optEl('optRecommendArea').addEventListener('click', (e) => {
            const btn = e.target.closest('[data-act]');
            if (!btn) return;
            const key = btn.dataset.key;
            if (!OptState.result) return;
            const s = OptState.result.paretoSolutions.find(x => x.key === key);
            if (!s) return;
            if (btn.dataset.act === 'detail') showSchemeDetail(s);
            else if (btn.dataset.act === 'apply') applySchemeToSimulation(s);
        });

        // Pareto 表格行点击（事件委托）
        optEl('optParetoTable').addEventListener('click', (e) => {
            const tr = e.target.closest('tr[data-key]');
            if (!tr || !OptState.result) return;
            const s = OptState.result.paretoSolutions.find(x => x.key === tr.dataset.key);
            if (s) showSchemeDetail(s);
        });

        // 切换到优化相关子标签页时，重绘图表（隐藏容器无法直接渲染）
        document.querySelectorAll('.sub-tab-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                if (!btn.dataset.subtab || btn.dataset.subtab.indexOf('opt-') !== 0) return;
                setTimeout(refreshOptChartsForVisibleTab, 30);
            });
        });
        const optTabBtn = document.querySelector('.tab-btn[data-tab="optimization"]');
        if (optTabBtn) {
            optTabBtn.addEventListener('click', () => setTimeout(refreshOptChartsForVisibleTab, 30));
        }

        // 初始化基准方案与「当前值」列：直接读取 AppState.currentScheme（唯一真值，§18）
        refreshOptimizationBaselineUI();
        optLog('方案优化模块已加载：基准方案自动取自「电量计算」页的当前方案；请先加载数据并确认基准，再设置搜索范围');
    }

    function startOptimization() {
        if (OptState.running) return;

        // 1. 组装并校验配置
        //    baseline 唯一来源 = AppState.currentScheme（任务书 §16 / §18），不再有独立输入框
        let config;
        try {
            config = OptimizationEngine.normalizeConfig({
                variables: readOptVariableConfig(),
                constraints: readOptConstraints(),
                nsga2: readOptNsgaConfig(),
                recommendationWeights: readOptWeights(),
                lcoh: { discountRate: Utils.toNum(optEl('optLcohDiscount').value, 5.0) },
                baseline: readBaselineScheme(),
            });
        } catch (err) {
            alert('优化参数读取失败：' + err.message);
            return;
        }

        const vres = OptimizationEngine.validateConfig(config);
        if (!vres.ok) {
            alert('优化参数非法：\n' + vres.message);
            return;
        }

        // 把本次搜索配置写入 ParameterManager，供全项目以统一接口读取（§6 / §23）
        ParameterManager.setOptimizationConfig(config);

        // 基准方案不在搜索范围内时只提示，不阻断（§28）
        const rangeRes = ParameterManager.checkBaselineInRange(config.baseline, config);
        if (!rangeRes.inRange) {
            optLog('提示：当前基准方案不在优化搜索空间内，优化结果可能不包含基准方案本身', 'warn');
        }

        // 2. 组装运行上下文
        let ctx;
        try {
            ctx = collectOptimizationContext();
        } catch (err) {
            alert(err.message);
            return;
        }

        // 3. 进入运行态
        OptState.running = true;
        OptState.cancelled = false;
        OptState.result = null;
        OptState.liveHistory = [];
        OptState.selectedKey = null;
        disposeOptCharts();

        optEl('btnRunOptimization').style.display = 'none';
        optEl('btnStopOptimization').style.display = '';
        optEl('optProgress').style.display = '';
        optEl('optProgressFill').style.width = '0%';
        optEl('optProgressTitle').textContent = '优化进行中...';
        optEl('optProgressGen').textContent = `第 0 / ${config.nsga2.generations} 代`;
        optEl('optProgressStats').innerHTML = '<span>正在初始化种群并计算 8760 小时仿真...</span>';
        optEl('btnExportOptimization').disabled = true;
        optEl('btnLoadRecommendedToSim').disabled = true;
        optEl('btnView8760').disabled = true;
        optEl('btnExportOptChartPng').disabled = true;
        optEl('optRecommendArea').innerHTML = '<div class="opt-empty">优化进行中，请稍候…</div>';

        optLog(`开始优化：种群 ${config.nsga2.populationSize}，迭代 ${config.nsga2.generations} 代，` +
               `交叉 ${config.nsga2.crossoverProbability}，变异 ${config.nsga2.mutationProbability}，` +
               `随机种子 ${config.nsga2.randomSeed}`);

        // 4. 优先 Web Worker，失败自动回退主线程分片调度
        try {
            startWithWorker(config, ctx);
        } catch (err) {
            optLog('Web Worker 不可用（' + err.message + '），已切换为主线程分片计算', 'warn');
            startWithMainThread(config, ctx);
        }
    }

    /** 路径 A：Web Worker */
    function startWithWorker(config, ctx) {
        const worker = new Worker('js/optimization-worker.js');
        OptState.worker = worker;
        OptState.mode = 'worker';
        optEl('optEngineMode').textContent = '计算模式：Web Worker（后台线程）';

        let gotProgress = false;
        let failed = false;

        const giveUp = (reason, type) => {
            if (failed || gotProgress) return;
            failed = true;
            if (OptState.fallbackTimer) clearTimeout(OptState.fallbackTimer);
            try { worker.terminate(); } catch (e) { /* ignore */ }
            OptState.worker = null;
            optLog(reason + '，已回退到主线程分片计算', type || 'warn');
            startWithMainThread(config, ctx);
        };

        // 兜底：若 Worker 长时间无响应（file:// 环境常见），自动回退
        OptState.fallbackTimer = setTimeout(() => giveUp('Worker 未在预期时间内响应'), 4000);

        worker.onerror = (e) => {
            giveUp('Web Worker 启动失败（' + ((e && e.message) || '脚本加载被浏览器安全策略拦截') + '）', 'error');
        };

        worker.onmessage = (e) => {
            const d = e.data || {};
            switch (d.type) {
                case 'ready':
                    break;
                case 'progress':
                    if (!gotProgress) { gotProgress = true; if (OptState.fallbackTimer) clearTimeout(OptState.fallbackTimer); }
                    updateOptProgress(d.progress);
                    break;
                case 'generation':
                    OptState.liveHistory.push(historyItemOf(d.progress));
                    break;
                case 'warn':
                    optLog(d.message, 'warn');
                    break;
                case 'complete':
                    if (OptState.fallbackTimer) clearTimeout(OptState.fallbackTimer);
                    try { worker.terminate(); } catch (err) { /* ignore */ }
                    OptState.worker = null;
                    onOptimizationFinished(d.result);
                    break;
                case 'error':
                    if (OptState.fallbackTimer) clearTimeout(OptState.fallbackTimer);
                    try { worker.terminate(); } catch (err) { /* ignore */ }
                    OptState.worker = null;
                    onOptimizationError(new Error(d.message));
                    break;
            }
        };

        // 注意：不使用 transfer，避免回退时源数组被 detach
        worker.postMessage({
            type: 'start',
            payload: {
                config: config,
                context: {
                    pvData: ctx.pvData,
                    windData: ctx.windData,
                    simulationConfig: ctx.simulationConfig,
                    prices: ctx.prices,
                    financeParams: ctx.financeParams,
                    lcohDiscountRate: ctx.lcohDiscountRate,
                },
            },
        });
    }

    /** 路径 B：主线程分片调度（每代让出事件循环，保证页面不卡死） */
    function startWithMainThread(config, ctx) {
        OptState.mode = 'main';
        optEl('optEngineMode').textContent = '计算模式：主线程分片调度（页面保持响应）';

        let session;
        try {
            session = OptimizationEngine.createSession({
                config: config,
                context: ctx,
                onProgress: (snap) => updateOptProgress(snap),
                onGeneration: (snap) => OptState.liveHistory.push(historyItemOf(snap)),
            });
        } catch (err) {
            onOptimizationError(err);
            return;
        }
        OptState.session = session;

        const CHUNK_MS = 60;

        const tick = () => {
            if (OptState.cancelled) {
                onOptimizationCancelled();
                return;
            }
            try {
                const t0 = performance.now();
                do {
                    if (session.isFinished()) break;
                    session.runNextGeneration();
                } while (!session.isFinished() && (performance.now() - t0) < CHUNK_MS);
            } catch (err) {
                onOptimizationError(err);
                return;
            }

            if (session.isFinished()) {
                onOptimizationFinished(session.getResult());
            } else {
                setTimeout(tick, 0);   // 让出主线程 → UI 保持响应
            }
        };

        // 先让 UI 完成一次渲染，再执行耗时的种群初始化
        setTimeout(() => {
            if (OptState.cancelled) { onOptimizationCancelled(); return; }
            try {
                session.ensureInitialized();
            } catch (err) {
                onOptimizationError(err);
                return;
            }
            if (session.isFinished()) onOptimizationFinished(session.getResult());
            else setTimeout(tick, 0);
        }, 60);
    }

    function stopOptimization() {
        if (!OptState.running) return;
        OptState.cancelled = true;
        optLog('已请求停止优化...', 'warn');

        if (OptState.session) {
            OptState.session.cancel();
            // 主线程分片调度会自动进入 onOptimizationCancelled
            return;
        }
        if (OptState.worker) {
            try {
                OptState.worker.postMessage({ type: 'cancel' });
            } catch (e) { /* ignore */ }
            // 若 Worker 在 1.5s 内未返回结果，直接终止
            setTimeout(() => {
                if (OptState.running && OptState.worker) {
                    try { OptState.worker.terminate(); } catch (e) { /* ignore */ }
                    OptState.worker = null;
                    onOptimizationCancelled();
                }
            }, 1500);
        }
    }

    function historyItemOf(snap) {
        if (!snap) return null;
        return {
            generation: snap.generation,
            evaluatedCount: snap.evaluated,
            feasibleCount: snap.feasibleCount,
            paretoCount: snap.paretoCount,
            bestFIRR: snap.bestFIRR,
            bestLCOH: snap.bestLCOH,
            bestCurtailmentRate: snap.bestCurtailmentRate,
        };
    }

    /** 进度刷新（每代一次，绝不逐个体操作 DOM） */
    function updateOptProgress(snap) {
        if (!snap) return;
        OptState.lastSnap = snap;
        const total = Math.max(1, snap.totalGenerations);
        const pct = Math.min(100, snap.generation / total * 100);
        optEl('optProgressFill').style.width = pct.toFixed(1) + '%';
        optEl('optProgressGen').textContent = `第 ${snap.generation} / ${snap.totalGenerations} 代`;
        optEl('optProgressStats').innerHTML = [
            `已评价方案 <b>${optInt(snap.evaluated)}</b>`,
            `缓存命中 <b>${optInt(snap.cacheHits)}</b>`,
            `可行方案 <b>${optInt(snap.feasibleCount)}</b>`,
            `Pareto方案 <b>${optInt(snap.paretoCount)}</b>`,
            `当前最佳FIRR <b>${optPct(snap.bestFIRR)}</b>`,
            `当前最低LCOH <b>${optNum(snap.bestLCOH, 2)} 元/kg</b>`,
            `当前最低弃电率 <b>${optPct(snap.bestCurtailmentRate === null ? null : snap.bestCurtailmentRate * 100)}</b>`,
            `耗时 <b>${optNum(snap.elapsedTime, 1)} s</b>`,
        ].join('');
    }

    function onOptimizationError(err) {
        const msg = (err && err.message) ? err.message : String(err);
        optLog('优化失败：' + msg, 'error');
        console.error(err);
        OptState.running = false;
        OptState.session = null;
        resetOptRunUI();
        optEl('optProgressTitle').textContent = '优化失败';
        optEl('optProgressStats').innerHTML = `<span style="color:var(--accent-red)">${msg}</span>`;
        optEl('optRecommendArea').innerHTML = `<div class="opt-warn">优化失败：${msg}</div>`;
    }

    function onOptimizationCancelled() {
        optLog('优化已停止', 'warn');
        OptState.running = false;
        OptState.session = null;
        resetOptRunUI();
        optEl('optProgress').style.display = 'none';
        optEl('optRecommendArea').innerHTML = '<div class="opt-empty">优化已停止。可调整参数后重新点击「开始优化」。</div>';
    }

    function resetOptRunUI() {
        optEl('btnRunOptimization').style.display = '';
        optEl('btnStopOptimization').style.display = 'none';
    }

    function onOptimizationFinished(result) {
        OptState.running = false;
        OptState.session = null;
        OptState.result = result;
        resetOptRunUI();

        const st = result.statistics || {};
        optEl('optProgressTitle').textContent = st.earlyStopped ? '优化完成（已提前终止）' : '优化完成';
        optEl('optProgressFill').style.width = '100%';

        optLog(`优化完成：共完成 ${st.generationsCompleted} 代，评价方案 ${st.totalEvaluated} 个（缓存命中 ${st.cacheHits}），` +
               `可行方案 ${st.feasibleCount} 个，Pareto 方案 ${st.paretoCount} 个，耗时 ${optNum(st.elapsedTime, 1)} s` +
               (st.earlyStopped ? '，已触发提前终止' : ''), 'success');

        if (!result.success) {
            optLog('当前约束条件下没有找到可行方案，请放宽约束或扩大容量搜索范围', 'warn');
        }

        renderOptimizationResult();
    }

    // -------------------- 结果渲染 --------------------

    function renderOptimizationResult() {
        if (!OptState.result) return;
        renderRecommendArea();
        renderParetoTable();
        renderOptimizationHistory();
        refreshOptChartsForVisibleTab();

        const ok = OptState.result.success;
        optEl('btnExportOptimization').disabled = !ok;
        optEl('btnLoadRecommendedToSim').disabled = !ok;
        optEl('btnView8760').disabled = !ok;
        optEl('btnExportOptChartPng').disabled = !ok;
    }

    /** 四类代表方案卡片 + 基准方案对比 + 优化说明 */
    function renderRecommendArea() {
        const area = optEl('optRecommendArea');
        const r = OptState.result;
        const st = r.statistics || {};
        const pareto = r.paretoSolutions || [];

        const statsLine = `<div class="opt-stat-line">
            <span>有效方案数 <b>${optInt(st.totalEvaluated)}</b></span>
            <span>可行方案数 <b>${optInt(st.feasibleCount)}</b></span>
            <span>Pareto方案数 <b>${optInt(st.paretoCount)}</b></span>
            <span>搜索空间 <b>${optInt(st.searchSpaceSize)}</b></span>
            <span>有效种群 <b>${optInt(st.effectivePopulationSize)}</b></span>
            <span>完成代数 <b>${optInt(st.generationsCompleted)}</b></span>
            <span>耗时 <b>${optNum(st.elapsedTime, 1)} s</b></span>
            <span>缓存命中 <b>${optInt(st.cacheHits)}</b></span>
        </div>`;

        const note = `<div class="opt-note">
            <b>优化说明：</b>本优化结果是在<b>当前风光 8760 小时数据、设备参数、投资参数、电价、氢价、财务参数及固定运行策略</b>条件下
            得到的 Pareto 最优解，<b>不代表脱离边界条件的绝对最优工程方案</b>。
            目标函数为：最大化 FIRR、最小化 LCOH（全生命周期折现口径，折现率 ${optNum((OptState.result.config.lcoh || {}).discountRate, 2)}%）、最小化新能源弃电率；
            综合推荐方案由权重（FIRR ${optNum(OptState.result.config.recommendationWeights.firr * 100, 0)}% /
            LCOH ${optNum(OptState.result.config.recommendationWeights.lcoh * 100, 0)}% /
            弃电率 ${optNum(OptState.result.config.recommendationWeights.curtailmentRate * 100, 0)}%）
            在 Pareto 前沿内部择优得到，权重不参与 NSGA-II 适应度。
        </div>`;

        if (!r.success) {
            area.innerHTML = statsLine +
                `<div class="opt-warn">当前约束条件下没有找到可行方案，请放宽约束或扩大容量搜索范围。</div>` + note;
            return;
        }

        const rep = r.representativeSolutions;
        const card = (title, cls, s) => {
            if (!s) {
                return `<div class="opt-card ${cls}"><div class="opt-card-title">${title}<span class="opt-card-badge">无</span></div>
                        <div class="opt-empty" style="min-height:60px">该类型方案未找到</div></div>`;
            }
            const sc = s.scheme, te = s.technical, ec = s.economic;
            return `<div class="opt-card ${cls}">
                <div class="opt-card-title">${title}<span class="opt-card-badge">${s.id || ''}</span></div>
                <div class="opt-card-row"><span>风电 / 光伏</span><b>${optNum(sc.windCapacity, 0)} / ${optNum(sc.pvCapacity, 0)} MW</b></div>
                <div class="opt-card-row"><span>储能</span><b>${optNum(sc.storagePower, 0)} MW / ${optNum(sc.storageDuration, 0)} h</b></div>
                <div class="opt-card-row"><span>电解槽</span><b>${optNum(sc.electrolyzerCapacity, 0)} MW</b></div>
                <div class="opt-card-row"><span>年制氢量</span><b>${optInt(te.annualHydrogenTon)} t/a</b></div>
                <div class="opt-card-row"><span>FIRR</span><b>${optPct(ec.FIRR)}</b></div>
                <div class="opt-card-row"><span>LCOH</span><b>${optNum(ec.LCOH, 2)} 元/kg</b></div>
                <div class="opt-card-row"><span>弃电率</span><b>${optPct(te.curtailmentRate * 100)}</b></div>
                <div class="opt-card-row"><span>总投资</span><b>${optNum(ec.totalInvestment / 10000, 3)} 亿元</b></div>
                <div class="opt-card-actions">
                    <button class="btn btn-xs btn-outline" data-act="detail" data-key="${s.key}">🔍 查看详情</button>
                    <button class="btn btn-xs btn-outline" data-act="apply" data-key="${s.key}">🔁 载入电量计算页复核</button>
                </div>
            </div>`;
        };

        const cards = `<div class="opt-card-grid">
            ${card('🏆 方案D 综合推荐', 'is-recommended', rep.recommended)}
            ${card('💰 方案A 经济最优（FIRR最高）', 'is-economic', rep.economicBest)}
            ${card('🧪 方案B 氢成本最优（LCOH最低）', 'is-h2cost', rep.hydrogenCostBest)}
            ${card('🌿 方案C 消纳最优（弃电率最低）', 'is-curtail', rep.curtailmentBest)}
        </div>`;

        // 基准方案对比
        let compare = '';
        const base = r.baseline;
        if (base && rep.recommended) {
            const rows = [
                ['风电容量', 'MW', s => s.scheme.windCapacity, 0],
                ['光伏容量', 'MW', s => s.scheme.pvCapacity, 0],
                ['储能功率', 'MW', s => s.scheme.storagePower, 0],
                ['储能时长', 'h', s => s.scheme.storageDuration, 0],
                ['储能容量', 'MWh', s => s.scheme.storageEnergy, 0],
                ['电解槽容量', 'MW', s => s.scheme.electrolyzerCapacity, 0],
                ['总投资', '亿元', s => s.economic.totalInvestment / 10000, 3],
                ['年制氢量', 't/a', s => s.technical.annualHydrogenTon, 0],
                ['电解槽利用小时', 'h', s => s.technical.electrolyzerHours, 1],
                ['弃电率', '%', s => s.technical.curtailmentRate * 100, 2],
                ['外购电比例', '%', s => s.technical.gridImportRatio * 100, 2],
                ['绿电制氢比例', '%', s => s.technical.greenHydrogenRatio * 100, 2],
                ['LCOH', '元/kg', s => s.economic.LCOH, 2],
                ['FIRR', '%', s => s.economic.FIRR, 2],
                ['EIRR', '%', s => s.economic.EIRR, 2],
                ['项目投资回收期', '年', s => s.economic.paybackPeriod, 2],
            ];
            const body = rows.map(([label, unit, getter, digits]) => {
                const b = getter(base), c = getter(rep.recommended);
                const d = (isFinite(b) && isFinite(c)) ? c - b : null;
                const cls = d === null ? 'opt-delta-none' : (Math.abs(d) < 1e-9 ? 'opt-delta-none' : 'opt-delta-up');
                const dtxt = d === null ? '—' : ((d > 0 ? '+' : '') + d.toFixed(digits === 0 ? 0 : digits));
                return `<tr><td>${label}（${unit}）</td><td>${optNum(b, digits)}</td><td>${optNum(c, digits)}</td><td class="${cls}">${dtxt}</td></tr>`;
            }).join('');

            const df = (base.economic.FIRR !== null && rep.recommended.economic.FIRR !== null)
                ? rep.recommended.economic.FIRR - base.economic.FIRR : null;
            const dl = (isFinite(base.economic.LCOH) && isFinite(rep.recommended.economic.LCOH))
                ? rep.recommended.economic.LCOH - base.economic.LCOH : null;
            const dc = (rep.recommended.technical.curtailmentRate - base.technical.curtailmentRate) * 100;

            compare = `<div class="opt-note">
                    <b>相较基准方案（Baseline）：</b>
                    FIRR ${df === null ? '—' : (df >= 0 ? '提高 +' : '降低 ') + df.toFixed(2) + ' 个百分点'}；
                    LCOH ${dl === null ? '—' : (dl <= 0 ? '降低 ' : '提高 +') + dl.toFixed(2) + ' 元/kg'}；
                    弃电率 ${(dc <= 0 ? '降低 ' : '提高 +') + dc.toFixed(2) + ' 个百分点'}。
                </div>
                <table class="opt-kv">
                    <caption>基准方案 vs 综合推荐方案</caption>
                    <thead><tr><th>指标</th><th>基准方案</th><th>综合推荐方案</th><th>变化</th></tr></thead>
                    <tbody>${body}</tbody>
                </table>`;
        } else {
            compare = `<div class="opt-note">未提供有效的基准方案，未能生成对比。可在左侧「基准方案」中填写或点击「读取电量计算页当前参数」。</div>`;
        }

        area.innerHTML = statsLine + cards + compare + note;
    }

    /** 当前显示的 Pareto 列表（排序 + 筛选后的结果） */
    function currentParetoList() {
        if (!OptState.result || !OptState.result.paretoSolutions) return [];
        let list = OptState.result.paretoSolutions.slice();

        const filter = optEl('optParetoFilter').value;
        if (filter === 'top20') list = list.slice(0, 20);
        else if (filter === 'economic') list = list.filter(s => s.economic.FIRR !== null && isFinite(s.economic.FIRR) && s.economic.FIRR >= 0);

        const sort = optEl('optParetoSort').value;
        const num = v => (v === null || v === undefined || !isFinite(Number(v))) ? (sort === 'lcoh' || sort === 'curtailment' || sort === 'investment' ? Infinity : -Infinity) : Number(v);
        const cmp = {
            score: (a, b) => num(b.score) - num(a.score),
            firr: (a, b) => num(b.economic.FIRR) - num(a.economic.FIRR),
            lcoh: (a, b) => num(a.economic.LCOH) - num(b.economic.LCOH),
            curtailment: (a, b) => num(a.technical.curtailmentRate) - num(b.technical.curtailmentRate),
            hydrogen: (a, b) => num(b.technical.annualHydrogenKg) - num(a.technical.annualHydrogenKg),
            investment: (a, b) => num(a.economic.totalInvestment) - num(b.economic.totalInvestment),
        }[sort] || ((a, b) => num(b.score) - num(a.score));

        return list.sort(cmp);
    }

    function renderParetoTable() {
        const table = optEl('optParetoTable');
        const list = currentParetoList();

        if (list.length === 0) {
            table.innerHTML = '<thead><tr><th>无非支配方案</th></tr></thead><tbody></tbody>';
            optEl('optParetoCount').textContent = '—';
            return;
        }

        optEl('optParetoCount').textContent = `共 ${list.length} 个方案（Pareto 前沿合计 ${OptState.result.paretoSolutions.length} 个）`;

        const cols = [
            ['排名', s => s.__rank],
            ['风电(MW)', s => optNum(s.scheme.windCapacity, 0)],
            ['光伏(MW)', s => optNum(s.scheme.pvCapacity, 0)],
            ['储能(MW)', s => optNum(s.scheme.storagePower, 0)],
            ['时长(h)', s => optNum(s.scheme.storageDuration, 0)],
            ['电解槽(MW)', s => optNum(s.scheme.electrolyzerCapacity, 0)],
            ['制氢量(万t)', s => optNum(s.technical.annualHydrogenWanTon, 3)],
            ['弃电率(%)', s => optNum(s.technical.curtailmentRate * 100, 2)],
            ['外购电(%)', s => optNum(s.technical.gridImportRatio * 100, 2)],
            ['绿电(%)', s => optNum(s.technical.greenHydrogenRatio * 100, 2)],
            ['LCOH(元/kg)', s => optNum(s.economic.LCOH, 2)],
            ['FIRR(%)', s => optNum(s.economic.FIRR, 2)],
            ['总投资(亿元)', s => optNum(s.economic.totalInvestment / 10000, 3)],
            ['综合得分', s => s.score === null || s.score === undefined ? '—' : optNum(s.score * 100, 2)],
        ];

        let html = '<thead><tr>' + cols.map(c => `<th>${c[0]}</th>`).join('') + '</tr></thead><tbody>';
        list.forEach((s, i) => {
            s.__rank = i + 1;
            const sel = (OptState.selectedKey && s.key === OptState.selectedKey) ? ' class="is-selected"' : '';
            html += `<tr data-key="${s.key}"${sel}>` + cols.map(c => `<td>${c[1](s)}</td>`).join('') + '</tr>';
        });
        html += '</tbody>';
        table.innerHTML = html;
    }

    function renderOptimizationHistory() {
        const table = optEl('optHistoryTable');
        const history = (OptState.result && OptState.result.history) || [];
        optEl('optHistorySummary').textContent = history.length
            ? `共 ${history.length} 代记录（含第 0 代初始化）`
            : '暂无优化历史';

        if (!history.length) {
            table.innerHTML = '<thead><tr><th>暂无优化历史</th></tr></thead><tbody></tbody>';
            return;
        }

        let html = '<thead><tr><th>Generation</th><th>Population（累计评价）</th><th>FeasibleCount</th>' +
                   '<th>ParetoCount</th><th>BestFIRR</th><th>BestLCOH</th><th>BestCurtailment</th></tr></thead><tbody>';
        for (const h of history) {
            html += `<tr>
                <td>${h.generation}</td>
                <td>${optInt(h.evaluatedCount)}</td>
                <td>${optInt(h.feasibleCount)}</td>
                <td>${optInt(h.paretoCount)}</td>
                <td>${optPct(h.bestFIRR)}</td>
                <td>${optNum(h.bestLCOH, 2)}</td>
                <td>${h.bestCurtailmentRate === null ? '—' : optPct(h.bestCurtailmentRate * 100)}</td>
            </tr>`;
        }
        html += '</tbody>';
        table.innerHTML = html;
    }

    // -------------------- 图表 --------------------

    function disposeOptCharts() {
        if (OptState.paretoChart) { try { OptState.paretoChart.dispose(); } catch (e) { } OptState.paretoChart = null; }
        for (const c of OptState.convCharts) { try { c.dispose(); } catch (e) { } }
        OptState.convCharts = [];
        ['optParetoChart', 'optConvFirr', 'optConvLcoh', 'optConvCurtail'].forEach(id => {
            const el = optEl(id);
            if (el) el.innerHTML = el.id === 'optParetoChart'
                ? '<div class="chart-placeholder">优化完成后在此处查看 Pareto 前沿</div>' : '';
        });
    }

    /** 依据当前可见子标签页渲染 / 重绘图表（隐藏容器宽高为 0，必须延迟渲染） */
    function refreshOptChartsForVisibleTab() {
        if (!OptState.result || !OptState.result.success) return;
        const active = document.querySelector('#tab-optimization .sub-tab-panel.active');
        if (!active) return;

        if (active.id === 'subtab-opt-charts') {
            renderParetoChart();
        } else if (active.id === 'subtab-opt-convergence') {
            renderConvergenceCharts();
        } else {
            // 其他页签下仅做尺寸同步，避免图表在切回时尺寸错误
            if (OptState.paretoChart) OptState.paretoChart.resize();
            for (const c of OptState.convCharts) if (c) c.resize();
        }
    }

    function renderParetoChart() {
        if (!OptState.result || !OptState.result.success) return;
        const container = optEl('optParetoChart');
        if (!container) return;

        const kind = optEl('optParetoChartType').value;
        const rep = OptState.result.representativeSolutions || {};

        if (OptState.paretoChart) { try { OptState.paretoChart.dispose(); } catch (e) { } }
        container.innerHTML = '';

        OptState.paretoChart = ChartModule.renderParetoScatter(container, OptState.result.paretoSolutions, kind, {
            repKeys: {
                recommended: rep.recommended ? rep.recommended.key : null,
                economicBest: rep.economicBest ? rep.economicBest.key : null,
                hydrogenCostBest: rep.hydrogenCostBest ? rep.hydrogenCostBest.key : null,
                curtailmentBest: rep.curtailmentBest ? rep.curtailmentBest.key : null,
            },
            selectedKey: OptState.selectedKey,
            onSelect: (s) => showSchemeDetail(s),
        });
    }

    function renderConvergenceCharts() {
        if (!OptState.result) return;
        const history = OptState.result.history || [];
        const defs = [
            ['optConvFirr', 'firr'],
            ['optConvLcoh', 'lcoh'],
            ['optConvCurtail', 'curtailment'],
        ];
        for (const c of OptState.convCharts) { try { c.dispose(); } catch (e) { } }
        OptState.convCharts = [];

        for (const [id, metric] of defs) {
            const el = optEl(id);
            if (!el) continue;
            el.innerHTML = '';
            OptState.convCharts.push(ChartModule.renderConvergenceChart(el, history, metric, false));
        }
    }

    // -------------------- 方案详情 --------------------

    function activateSubTab(subtab) {
        const btn = document.querySelector(`.sub-tab-btn[data-subtab="${subtab}"]`);
        if (btn) btn.click();
    }

    function showSchemeDetail(s) {
        if (!s) return;
        OptState.selectedKey = s.key;

        const sc = s.scheme, te = s.technical, ec = s.economic, co = s.constraints || { violations: [] };
        const kv = (caption, rows) => `<table class="opt-kv"><caption>${caption}</caption><tbody>` +
            rows.map(r => `<tr><td>${r[0]}</td><td colspan="3">${r[1]}</td></tr>`).join('') + '</tbody></table>';

        const viol = (co.violations && co.violations.length)
            ? co.violations.map(v => `<tr><td>${v.name}</td><td colspan="3" style="color:var(--accent-red)">实际 ${v.actual === null ? '—' : optNum(v.actual, 2)}${v.unit || ''} / 限值 ${v.limit === null ? '—' : optNum(v.limit, 2)}${v.unit || ''}</td></tr>`).join('')
            : '<tr><td>约束校验</td><td colspan="3" style="color:var(--accent-green)">全部满足</td></tr>';

        optEl('optSchemeDetail').innerHTML = `
            <div class="opt-stat-line">
                <span>方案编号 <b>${s.id || '—'}</b></span>
                <span>可行性 <b style="color:${co.feasible ? 'var(--accent-green)' : 'var(--accent-red)'}">${co.feasible ? '可行' : '不可行'}</b></span>
                <span>综合得分 <b>${s.score === null || s.score === undefined ? '—' : optNum(s.score * 100, 2)}</b></span>
            </div>
            ${kv('一、容量方案', [
                ['风电容量', optNum(sc.windCapacity, 0) + ' MW'],
                ['光伏容量', optNum(sc.pvCapacity, 0) + ' MW'],
                ['储能功率', optNum(sc.storagePower, 0) + ' MW'],
                ['储能时长', optNum(sc.storageDuration, 0) + ' h'],
                ['储能容量', optNum(sc.storageEnergy, 0) + ' MWh'],
                ['电解槽容量', optNum(sc.electrolyzerCapacity, 0) + ' MW'],
            ])}
            ${kv('二、年度技术指标', [
                ['新能源年发电量', optNum(te.theoreticalEnergy / 1000, 2) + ' GWh'],
                ['风电年发电量', optNum(te.annualWindEnergy / 1000, 2) + ' GWh'],
                ['光伏年发电量', optNum(te.annualPvEnergy / 1000, 2) + ' GWh'],
                ['年制氢量', optInt(te.annualHydrogenTon) + ' t/a（' + optNum(te.annualHydrogenWanTon, 3) + ' 万吨/年）'],
                ['电解槽耗电量', optNum(te.electrolyzerElectricity / 1000, 2) + ' GWh'],
                ['电解槽利用小时', optNum(te.electrolyzerHours, 1) + ' h'],
                ['电解槽负荷率', optPct(te.electrolyzerLoadFactor * 100)],
                ['储能充电量', optNum(te.annualChargeEnergy / 1000, 2) + ' GWh'],
                ['储能放电量', optNum(te.annualDischargeEnergy / 1000, 2) + ' GWh'],
                ['上网电量', optNum(te.annualGridExport / 1000, 2) + ' GWh'],
                ['下网电量', optNum(te.annualGridImport / 1000, 2) + ' GWh'],
                ['弃电量', optNum(te.annualCurtailment / 1000, 2) + ' GWh'],
                ['弃电率', optPct(te.curtailmentRate * 100)],
                ['外购电比例', optPct(te.gridImportRatio * 100)],
                ['绿电制氢比例', optPct(te.greenHydrogenRatio * 100)],
                ['上网比例', optPct(te.exportRatio * 100)],
                ['逐时最大上网功率', optNum(te.maxHourlyExportPower, 2) + ' MW'],
                ['逐时最大下网功率', optNum(te.maxHourlyImportPower, 2) + ' MW'],
            ])}
            ${kv('三、经济指标', [
                ['总投资', optNum(ec.totalInvestment / 10000, 3) + ' 亿元（' + optNum(ec.totalInvestment, 0) + ' 万元）'],
                ['建设投资', optNum(ec.constructionInvestment, 0) + ' 万元'],
                ['建设期利息', optNum(ec.constructionInterest, 0) + ' 万元'],
                ['流动资金', optNum(ec.workingCapital, 0) + ' 万元'],
                ['年均营业收入（不含税）', optNum(ec.annualRevenue, 0) + ' 万元'],
                ['年均总成本费用', optNum(ec.annualCost, 0) + ' 万元'],
                ['FIRR（项目投资，税前）', optPct(ec.FIRR)],
                ['项目投资 IRR（税后）', optPct(ec.FIRRPostTax)],
                ['EIRR（资本金）', optPct(ec.EIRR)],
                ['项目投资财务净现值（税前）', optNum(ec.NPV, 0) + ' 万元'],
                ['投资回收期（税前）', optNum(ec.paybackPeriod, 2) + ' 年'],
                ['总投资收益率 ROI', optPct(ec.ROI)],
                ['资本金净利润率 ROE', optPct(ec.ROE)],
                ['盈亏平衡点', optPct(ec.breakEvenPoint)],
                ['LCOH（折现口径）', optNum(ec.LCOH, 2) + ' 元/kg'],
                ['制氢成本（未折现口径，V1.0）', optNum(ec.h2CostAvg, 2) + ' 元/kg'],
            ])}
            <table class="opt-kv"><caption>四、约束校验</caption><tbody>${viol}</tbody></table>
            <div class="opt-note">
                FIRR 口径：项目投资财务内部收益率（所得税前）；EIRR 口径：资本金财务内部收益率。
                LCOH 口径：项目全生命周期折现成本 / 项目全生命周期折现制氢量，
                成本项为建设投资 + 建设期利息（按建设期均摊）+ 流动资金（投产年）+ 年经营成本 + 大修费，
                不含折旧与所得税，折现率 ${optNum((OptState.result.config.lcoh || {}).discountRate, 2)}%。
            </div>`;

        renderParetoTable();     // 同步表格选中态
        activateSubTab('opt-detail');
    }

    // -------------------- 8760 小时复核 --------------------

    /**
     * 把优化方案写回「电量计算」页；runSimulation 为 true 时立即运行一次仿真并展示 8760 小时结果。
     * 全程复用 V1.0 仿真链路（runSingleSimulation + ChartModule），不新建任何仿真逻辑。
     */
    function applySchemeToSimulation(s, runSimulation) {
        if (!s) return;
        const sc = s.scheme || s;
        const doRun = runSimulation !== false;

        if (!AppState.inputData || AppState.inputData.length === 0) {
            alert('请先在「电量计算」页加载 input.xlsx 文件！');
            return;
        }

        // 唯一写入路径：setCurrentScheme → syncSchemeToUI（任务书 §19 / §33）
        // 候选方案只是候选；只有用户点击「应用」才允许修改当前方案（§32）
        const res = ParameterManager.setCurrentScheme(sc);
        if (!res.ok) { alert('方案参数非法：\n' + res.message); return; }
        ParameterManager.syncSchemeToUI(res.scheme);
        refreshOptimizationBaselineUI();

        optLog(`已应用方案 ${ParameterManager.schemeKey(res.scheme)}（${ParameterManager.schemeLabel(res.scheme)}）到「当前方案」` +
            (doRun ? '，正在运行 8760 小时复核仿真...' : '，请点击「开始仿真计算」运行复核'), 'success');

        const tabBtn = document.querySelector('.tab-btn[data-tab="simulation"]');
        if (tabBtn) tabBtn.click();
        if (!doRun) return;

        startSimulation();

        // 仿真在主线程分片执行，完成后自动切到「图表分析」（轮询而非固定延时，避免竞态）
        let tries = 0;
        const waitSimulation = () => {
            tries++;
            if (AppState.simulationResults && AppState.simulationResults.length > 0) {
                const chartTab = document.querySelector('.sub-tab-btn[data-subtab="sim-charts"]');
                if (chartTab) chartTab.click();
                updateChart();
                log('8760 小时复核完成：可在「数据表格」查看逐时数据，在「图表分析」查看风电/光伏/电解槽/储能SOC/充放电/上网/下网/弃电曲线', 'success');
                return;
            }
            if (tries < 200) setTimeout(waitSimulation, 150);
        };
        setTimeout(waitSimulation, 150);
    }

    // -------------------- 导出 --------------------

    async function exportOptimizationExcel() {
        if (!OptState.result || !OptState.result.success) return;
        try {
            const buffer = await ExcelIO.exportOptimizationResults({
                config: OptState.result.config,
                statistics: OptState.result.statistics,
                paretoSolutions: OptState.result.paretoSolutions,
                representativeSolutions: OptState.result.representativeSolutions,
                baseline: OptState.result.baseline,
                history: OptState.result.history,
            });
            const blob = new Blob([buffer], { type: 'application/octet-stream' });
            Utils.downloadBlob(blob, `优化结果_${Utils.timestamp()}.xlsx`);
            optLog('已导出优化结果 Excel（优化参数 / Pareto方案 / 代表方案 / 优化过程 / 基准方案对比）', 'success');
        } catch (err) {
            optLog('优化结果导出失败：' + err.message, 'error');
            console.error(err);
        }
    }

    // ==================================================================
    // ========== 批量计算模块（V2.2 从「电量计算」主参数区独立出来）==========
    // ==================================================================
    //
    // 任务书 §7 / §40：多组容量组合的扫描不再混入「电量计算」的主参数区，
    // 单独成为「批量计算」页，并且按钮文案与「当前方案」彻底区分。
    //   · 只产生 schemeList[]，**绝不修改 AppState.currentScheme**；
    //   · 运行参数沿用「电量计算」页的设置；
    //   · 需要把某一组设为当前方案时，必须由用户主动点击（§32）。

    const BATCH_VARS = [
        { key: 'windCapacity',         min: 'batchWindMin',         max: 'batchWindMax',         step: 'batchWindStep',         levels: 'batchWindLevels' },
        { key: 'pvCapacity',           min: 'batchPvMin',           max: 'batchPvMax',           step: 'batchPvStep',           levels: 'batchPvLevels' },
        { key: 'storagePower',         min: 'batchStoragePowerMin', max: 'batchStoragePowerMax', step: 'batchStoragePowerStep', levels: 'batchStoragePowerLevels' },
        { key: 'storageDuration',      min: 'batchStorageDurationMin', max: 'batchStorageDurationMax', step: 'batchStorageDurationStep', levels: 'batchStorageDurationLevels' },
        { key: 'electrolyzerCapacity', min: 'batchElectrolyzerMin', max: 'batchElectrolyzerMax', step: 'batchElectrolyzerStep', levels: 'batchElectrolyzerLevels' },
    ];

    const BatchState = {
        running: false,
        cancelled: false,
        schemes: [],
        results: [],
        _cancel: null,
    };

    function readBatchVariables() {
        const out = {};
        for (const v of BATCH_VARS) {
            out[v.key] = {
                min: Utils.toNum(optEl(v.min).value, NaN),
                max: Utils.toNum(optEl(v.max).value, NaN),
                step: Utils.toNum(optEl(v.step).value, NaN),
            };
        }
        return out;
    }

    function updateBatchPreview() {
        const vars = readBatchVariables();
        let space = 1;
        for (const v of BATCH_VARS) {
            const cfg = vars[v.key];
            const n = (isFinite(cfg.min) && isFinite(cfg.max))
                ? ParameterManager.countLevels(cfg.min, cfg.max, cfg.step)
                : 0;
            optEl(v.levels).textContent = n > 0 ? n : '—';
            space *= Math.max(1, n);
        }
        optEl('batchSearchSpace').innerHTML = `搜索空间：<b>${space.toLocaleString('zh-CN')}</b> 个方案`;
        return space;
    }

    function initBatch() {
        for (const v of BATCH_VARS) {
            for (const id of [v.min, v.max, v.step]) {
                const el = optEl(id);
                if (el) el.addEventListener('input', updateBatchPreview);
            }
        }
        optEl('btnRunBatch').addEventListener('click', runBatch);
        optEl('btnCancelBatch').addEventListener('click', cancelBatch);
        optEl('btnExportBatchSchemes').addEventListener('click', exportBatchSchemes);

        // 「设为当前方案」按钮（行内动态渲染，使用事件委托）
        optEl('batchTable').addEventListener('click', (e) => {
            const btn = e.target.closest('[data-batch-act]');
            if (!btn) return;
            const s = BatchState.results[parseInt(btn.dataset.idx, 10)];
            if (s && s.scheme) applyBatchScheme(s.scheme);
        });

        updateBatchPreview();
    }

    function runBatch() {
        if (BatchState.running) return;
        if (!AppState.inputData || AppState.inputData.length === 0) {
            alert('请先在「电量计算」页加载 input.xlsx 文件！（两个页面共用同一份数据）');
            return;
        }

        const vars = readBatchVariables();
        const bad = [];
        for (const v of BATCH_VARS) {
            const cfg = vars[v.key];
            if (!isFinite(cfg.min) || !isFinite(cfg.max)) bad.push(ParameterManager.SCHEME_LABELS[v.key] + '：范围必须为有效数值');
            else if (cfg.min > cfg.max) bad.push(ParameterManager.SCHEME_LABELS[v.key] + '：最小值不能大于最大值');
            if (!isFinite(cfg.step) || cfg.step < 0) bad.push(ParameterManager.SCHEME_LABELS[v.key] + '：步长不能为负');
        }
        if (bad.length) { alert('批量计算参数非法：\n' + bad.join('\n')); return; }

        const schemes = ParameterManager.createBatchSchemes(vars);
        if (schemes.length === 0) { alert('搜索空间为空，请检查扫描范围与步长'); return; }
        if (schemes.length > 2000 && !confirm(`将计算 ${schemes.length} 个方案，可能耗时较久，是否继续？`)) return;

        // 运行参数沿用「电量计算」页设置（批量计算不引入第二套运行参数）
        const simulationConfig = getSimulationConfig();
        const n = AppState.inputData.length;
        const pvArr = new Float64Array(n);
        const windArr = new Float64Array(n);
        for (let i = 0; i < n; i++) {
            pvArr[i] = AppState.inputData[i].pv;
            windArr[i] = AppState.inputData[i].wind;
        }

        BatchState.running = true;
        BatchState.cancelled = false;
        BatchState.schemes = schemes;
        BatchState.results = [];

        optEl('btnRunBatch').style.display = 'none';
        optEl('btnCancelBatch').style.display = '';
        optEl('batchProgressContainer').style.display = '';
        optEl('batchProgressFill').style.width = '0%';
        optEl('batchProgressText').textContent = '准备中...';
        optEl('batchProgressPercent').textContent = '0%';
        optEl('btnExportBatchSchemes').disabled = true;
        optEl('batchSummaryText').textContent = '计算中...';
        renderBatchTable();

        log(`【批量】开始批量计算：共 ${schemes.length} 个方案（运行参数沿用「电量计算」页设置，不修改当前方案）`);

        let i = 0;
        const CHUNK_MS = 40;
        BatchState._cancel = () => { BatchState.cancelled = true; };

        const step = () => {
            if (BatchState.cancelled) { finishBatch(true); return; }
            const t0 = performance.now();
            try {
                while (i < schemes.length && (performance.now() - t0) < CHUNK_MS) {
                    BatchState.results.push(runSingleSimulation(pvArr, windArr, schemes[i], simulationConfig));
                    i++;
                }
            } catch (err) {
                log('【批量】计算失败：' + err.message, 'error');
                finishBatch(true);
                return;
            }
            const pct = (i / schemes.length) * 100;
            optEl('batchProgressFill').style.width = pct.toFixed(1) + '%';
            optEl('batchProgressText').textContent = `已完成 ${i}/${schemes.length} 个方案`;
            optEl('batchProgressPercent').textContent = pct.toFixed(0) + '%';
            if (i < schemes.length) setTimeout(step, 0);   // 让出主线程，页面保持响应
            else finishBatch(false);
        };
        setTimeout(step, 0);
    }

    function finishBatch(cancelled) {
        BatchState.running = false;
        BatchState._cancel = null;
        optEl('btnRunBatch').style.display = '';
        optEl('btnCancelBatch').style.display = 'none';
        if (cancelled) optEl('batchProgressContainer').style.display = 'none';
        optEl('btnExportBatchSchemes').disabled = BatchState.results.length === 0;

        if (cancelled) {
            optEl('batchSummaryText').textContent = `已取消（完成 ${BatchState.results.length} 个）`;
            log(`【批量】已取消，完成 ${BatchState.results.length} 个方案`, 'warn');
        } else {
            optEl('batchSummaryText').textContent = `完成 ${BatchState.results.length} 个方案`;
            log(`【批量】批量计算完成，共 ${BatchState.results.length} 个方案`, 'success');
        }
        renderBatchTable();
    }

    function cancelBatch() {
        if (BatchState._cancel) {
            BatchState._cancel();
            log('【批量】已请求取消...', 'warn');
        }
    }

    function renderBatchTable() {
        const tbody = optEl('batchTable').querySelector('tbody');
        if (!tbody) return;
        const curKey = ParameterManager.schemeKey(ParameterManager.getCurrentScheme());

        if (BatchState.results.length === 0) {
            tbody.innerHTML = BatchState.schemes.length > 0
                ? `<tr><td colspan="10" class="batch-empty">正在计算，共 ${BatchState.schemes.length} 个方案…</td></tr>`
                : '<tr><td colspan="10" class="batch-empty">设置扫描范围后点击「开始批量计算」</td></tr>';
            return;
        }

        // 汇总复用 data-summary（容量取自每个结果自带的 scheme）
        const summary = DataSummary.generateSummary(BatchState.results, ParameterManager.getCurrentScheme());

        let html = '';
        BatchState.results.forEach((r, i) => {
            const s = r.scheme;
            const key = ParameterManager.schemeKey(s);
            const row = summary[i] || {};
            const isCur = (key === curKey);
            const h2 = row['制氢量总和（万吨）'];
            const curt = row['弃电电量比例'];
            html += `<tr class="${isCur ? 'is-current' : ''}">
                <td>${i + 1}</td>
                <td>${key}</td>
                <td class="batch-num">${s.windCapacity}</td>
                <td class="batch-num">${s.pvCapacity}</td>
                <td class="batch-num">${s.storagePower} / ${s.storageDuration}</td>
                <td class="batch-num">${s.storageEnergy}</td>
                <td class="batch-num">${s.electrolyzerCapacity}</td>
                <td class="batch-num">${h2 === undefined ? '—' : h2}</td>
                <td class="batch-num">${curt === undefined ? '—' : curt}</td>
                <td>${isCur
                    ? '<span class="hint">＝当前方案</span>'
                    : `<button class="btn btn-sm btn-outline" data-batch-act="apply" data-idx="${i}" type="button">设为当前方案</button>`}</td>
            </tr>`;
        });
        tbody.innerHTML = html;
    }

    /**
     * 把批量计算中的某一组设为「当前方案」（任务书 §40）。
     * 批量计算本身绝不修改当前方案，只有用户主动点击才写入。
     */
    function applyBatchScheme(scheme) {
        const res = ParameterManager.setCurrentScheme(scheme);
        if (!res.ok) { alert('方案参数非法：\n' + res.message); return; }
        ParameterManager.syncSchemeToUI(res.scheme);
        refreshOptimizationBaselineUI();
        renderBatchTable();
        log(`【批量】已将 ${ParameterManager.schemeKey(res.scheme)} 设为当前方案，正在运行 8760 小时仿真...`, 'success');

        const tabBtn = document.querySelector('.tab-btn[data-tab="simulation"]');
        if (tabBtn) tabBtn.click();
        startSimulation();
    }

    async function exportBatchSchemes() {
        if (BatchState.results.length === 0) return;
        try {
            const summary = DataSummary.generateSummary(BatchState.results, ParameterManager.getCurrentScheme());
            const clean = summary.map((row, i) => {
                const o = { '方案编号': ParameterManager.schemeKey(BatchState.results[i].scheme) };
                for (const k of Object.keys(row)) {
                    if (!k.startsWith('_')) o[k] = row[k];
                }
                return o;
            });
            const buffer = await ExcelIO.exportSummaryExcel(clean);
            const blob = new Blob([buffer], { type: 'application/octet-stream' });
            Utils.downloadBlob(blob, `批量方案清单_${Utils.timestamp()}.xlsx`);
            log('【批量】已导出方案清单 Excel', 'success');
        } catch (err) {
            log('【批量】导出失败：' + err.message, 'error');
        }
    }

    // ========== 初始化 ==========
    function init() {
        initTabs();

        // 参数层必须先于其他模块初始化：它是 AppState.currentScheme 的唯一来源（§20 / §21）
        if (typeof ParameterManager === 'undefined') {
            log('⚠ 参数管理器未加载（js/parameter-manager.js），参数体系不可用', 'error');
            return;
        }
        initParameterLayer();

        initSimulation();
        initEstimate();
        initFinance();
        initOptimization();
        initBatch();

        document.getElementById('btnClearLog').addEventListener('click', clearLog);

        log('多能互补风光储氢分析软件 WEB-V2.2 已就绪（参数体系统一 + NSGA-II 多目标容量优化）', 'success');
        log('V2.2 参数体系：一个参数、一个定义、一个数据源、一个传递路径');
        log('　· 当前方案（风电/光伏/储能功率/储能时长/电解槽）唯一定义在「电量计算」页，优化页的基准方案直接读取它');
        log('　· 优化范围与工程约束只用于搜索，不参与单方案计算，也不会改变当前方案');
        log('　· 多组容量组合的批量扫描已独立为「批量计算」页');
        log('请先选择 input.xlsx 文件，然后配置参数并运行仿真计算');

        // 检查依赖库
        if (typeof XLSX === 'undefined') log('⚠ SheetJS 库未加载，Excel功能不可用', 'warn');
        if (typeof ExcelJS === 'undefined') log('⚠ ExcelJS 库未加载，Excel导出功能不可用', 'warn');
        if (typeof echarts === 'undefined') log('⚠ ECharts 库未加载，图表功能不可用', 'warn');
        if (typeof JSZip === 'undefined') log('⚠ JSZip 库未加载，批量导出功能不可用', 'warn');
        if (typeof OptimizationEngine === 'undefined') log('⚠ 优化引擎未加载，「方案优化」页不可用', 'error');
    }

    // 等待DOM加载完成
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

})();
