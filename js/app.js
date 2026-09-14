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
    /**
     * 日志上限（任务书 §20）：优化过程中日志会持续增长，
     * innerHTML += 每次都全量重解析整个容器，且 DOM 无限膨胀。
     * 超过上限后删除最早的行，保证 DOM 规模恒定。
     */
    const MAX_LOG_LINES = 300;

    function log(message, type = 'info') {
        const logContent = document.getElementById('logContent');
        const time = Utils.timeNow();
        const cls = type === 'error' ? 'log-error' : type === 'success' ? 'log-success' : type === 'warn' ? 'log-warn' : 'log-msg';

        // 一次性构建节点并追加（禁止在循环中 innerHTML +=）
        const row = document.createElement('div');
        const t = document.createElement('span');
        t.className = 'log-time';
        t.textContent = `[${time}]`;
        const m = document.createElement('span');
        m.className = cls;
        m.textContent = message;          // textContent 而非 innerHTML，杜绝注入与重复解析
        row.appendChild(t);
        row.appendChild(document.createTextNode(' '));
        row.appendChild(m);
        logContent.appendChild(row);

        // 限长：一次性删除超量的旧行
        const over = logContent.childElementCount - MAX_LOG_LINES;
        if (over > 0) {
            for (let i = 0; i < over; i++) logContent.removeChild(logContent.firstChild);
        }
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

                // V2.3.1（任务书 §9 / §35）：输入数据版本在加载时生成一次并复用，
                // 绝不在每次生成 simulationKey 时重复哈希 8760 数据。
                // 换数据 → inputVersion 变化 → simulationKey / 图表缓存全部失效。
                AppState.inputVersion = ResultDataStore.createInputVersionFromRows(AppState.inputData, {
                    name: file.name, size: file.size,
                });
                ResultDataStore.clearAllCaches();
                log(`输入数据版本 ${AppState.inputVersion}，已使全部缓存失效`, 'info');

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
        // V2.3：DocumentFragment 一次性插入，禁止循环内 innerHTML +=（任务书 §21）
        const sel = document.getElementById('schemeSelect');
        const chartSel = document.getElementById('chartSchemeSelect');
        sel.innerHTML = '';
        chartSel.innerHTML = '';

        const fragA = document.createDocumentFragment();
        const fragB = document.createDocumentFragment();
        results.forEach((r, i) => {
            const sv = r.systemVars;
            const label = r.scheme
                ? ParameterManager.schemeKey(r.scheme)
                : `方案${i + 1}`;
            const optText = `${label}: 电解槽${sv['电解槽容量（MW）']}MW, 储能${sv['储能功率（MW）']}MW×${sv['储能时长（小时）']}h`;
            const o1 = document.createElement('option');
            o1.value = String(i);
            o1.textContent = optText;
            const o2 = o1.cloneNode(true);
            fragA.appendChild(o1);
            fragB.appendChild(o2);
        });
        sel.appendChild(fragA);
        chartSel.appendChild(fragB);

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

    // ==================================================================
    // ========== 逐时数据表（V2.3 虚拟滚动，任务书 §6 / §7）==========
    // ==================================================================
    //
    // V2.2 的问题：一次性生成 8760 行 × 12 列 ≈ 10.5 万个 DOM 节点，
    // 且每次切换方案都重建 —— 这是浏览器卡顿的首要来源。
    //
    // V2.3 的做法：
    //   · 数据仍是完整的 8760 小时（Float64Array，经 ResultDataStore 零拷贝读取）；
    //   · DOM 只保留「视口内 + 上下缓冲」约 50~120 行；
    //   · 滚动时只重算可视区间并重建这一小段 DOM；
    //   · 「总和」行放入 <tfoot>，始终可见且不参与虚拟化；
    //   · 提供跳转小时 / 首页 / 上一页 / 下一页 / 末页 定位。

    /**
     * 取（或惰性创建）某个仿真结果的统一结果记录。
     * 结果层只有一个数据源：results 仍是原始 Float64Array，零拷贝（任务书 §4 / §26）。
     * 同时登记到 SimulationResultCache —— 只保留用户实际查看过的方案（§19）。
     */
    function ensureResultStore(simResult) {
        if (!simResult._store) {
            simResult._store = ResultDataStore.create(simResult.results, simResult.scheme, {
                simulationConfig: simResult.simulationConfig,
                systemVars: simResult.systemVars,
                ratioData: simResult.ratioData,
                sums: simResult.sums,
                filename: simResult.filename,
                inputVersion: AppState.inputVersion,
            });
            ResultDataStore.SimulationResultCache.put(simResult._store);
        }
        return simResult._store;
    }

    const SimTable = {
        ROW_HEIGHT: 28,     // 行高（首帧后按实际渲染高度校准）
        BUFFER_ROWS: 10,    // 视口上下各多渲染的行数，滚动时不露白
        PAGE_ROWS: 240,     // 上一页 / 下一页滚动的行数
        result: null,       // ResultDataStore 结果记录
        cols: [],           // 列定义
        totalRows: 0,
        rowHeight: 28,
        rendered: { start: -1, end: -1 },
    };

    /** 逐时表的列定义（顺序与 simulation-engine.js 的输出列严格一致） */
    const SIM_TABLE_COLS = [
        { header: '光伏电量(MWh)',     col: 'pv' },
        { header: '风电电量(MWh)',     col: 'wind' },
        { header: '合计电量(MWh)',     col: 'total' },
        { header: '储能充电电量(MWh)', col: 'charge' },
        { header: '储能放电电量(MWh)', col: 'discharge' },
        { header: '储能现存容量(MWh)', col: 'storage', takeLast: true },
        { header: '制氢电量(MWh)',     col: 'hydrogenEnergy' },
        { header: '上网电量(MWh)',     col: 'export' },
        { header: '下网电量(MWh)',     col: 'import' },
        { header: '弃电量(MWh)',       col: 'curtailment' },
        { header: '制氢量(kg)',        col: 'hydrogen' },
    ];

    /**
     * 显示某个方案的逐时数据表（虚拟滚动版）。
     * @param {number} index AppState.simulationResults 的下标
     */
    function displaySimulationTable(index) {
        return Utils.PerfTimer.measure('Table.render', () => displaySimulationTableInner(index));
    }

    function displaySimulationTableInner(index) {
        const simResult = AppState.simulationResults[index];
        if (!simResult) return;
        if (typeof ResultDataStore === 'undefined') {
            // 防御：结果数据层未加载时给出可读提示，而不是静默失败
            document.getElementById('simulationTable').innerHTML =
                '<thead><tr><th>结果数据层未加载（js/result-data-store.js）</th></tr></thead><tbody></tbody>';
            return;
        }

        // ---- 结果层：把原始结果包装成统一结果记录（零拷贝），并注册到查看缓存 ----
        SimTable.result = ensureResultStore(simResult);
        SimTable.cols = SIM_TABLE_COLS;
        SimTable.totalRows = SimTable.result.hours;
        SimTable.rendered = { start: -1, end: -1 };

        buildSimTableHead();
        buildSimTableFoot();
        renderSimTableRows(true);
    }

    /** 表头 + 列宽（一次性构建；table-layout:fixed 使列宽稳定且首帧更快） */
    function buildSimTableHead() {
        const table = document.getElementById('simulationTable');
        const thead = table.tHead || table.querySelector('thead');
        let html = '<colgroup><col style="width:64px">';
        for (let i = 0; i < SimTable.cols.length; i++) html += '<col>';
        html += '</colgroup><tr><th>小时</th>';
        for (const c of SimTable.cols) html += `<th>${c.header}</th>`;
        html += '</tr>';
        thead.innerHTML = html;
    }

    /** 「总和」行放入 tfoot：始终可见，不参与虚拟化（§18：复用年度摘要，不再二次遍历） */
    function buildSimTableFoot() {
        const table = document.getElementById('simulationTable');
        let tfoot = table.tFoot;
        if (!tfoot) {
            tfoot = document.createElement('tfoot');
            table.appendChild(tfoot);
        }
        const annual = ResultDataStore.getAnnualSummary(SimTable.result);
        let html = '<tr><td><b>总和</b></td>';
        for (const c of SimTable.cols) {
            const v = annual.byLabel[ResultDataStore.COL_LABELS[ResultDataStore.COLS[c.col]]];
            html += `<td><b>${Number(v).toFixed(2)}</b></td>`;
        }
        html += '</tr>';
        tfoot.innerHTML = html;
    }

    /** 只渲染可视区间内的行（含上下缓冲） */
    function renderSimTableRows(force) {
        const st = SimTable;
        if (!st.result) return;
        const container = document.getElementById('simTableContainer');
        const tbody = document.getElementById('simulationTable').tBodies[0];
        if (!container || !tbody) return;

        const scrollTop = container.scrollTop;
        const viewport = container.clientHeight || 480;
        const rowH = st.rowHeight;

        const first = Math.max(0, Math.floor(scrollTop / rowH) - st.BUFFER_ROWS);
        const visible = Math.ceil(viewport / rowH) + st.BUFFER_ROWS * 2;
        const last = Math.min(st.totalRows, first + visible);

        if (!force && first === st.rendered.start && last === st.rendered.end) return;
        st.rendered = { start: first, end: last };

        const frag = document.createDocumentFragment();
        const spacerTd = () => {
            const td = document.createElement('td');
            td.className = 'vt-spacer-cell';
            return td;
        };

        // 顶部占位（撑起已滚过的高度，使滚动条长度与 8760 行一致）
        if (first > 0) {
            const tr = document.createElement('tr');
            tr.className = 'vt-spacer';
            for (let i = 0; i <= st.cols.length; i++) {
                const td = spacerTd();
                if (i === 0) td.style.height = (first * rowH) + 'px';
                tr.appendChild(td);
            }
            frag.appendChild(tr);
        }

        // 可视数据行（直接读 Float64Array，不产生行对象）
        for (let h = first; h < last; h++) {
            const tr = document.createElement('tr');
            const tdH = document.createElement('td');
            tdH.textContent = h;
            tr.appendChild(tdH);
            for (const c of st.cols) {
                const td = document.createElement('td');
                td.textContent = ResultDataStore.getValue(st.result, h, c.col).toFixed(2);
                tr.appendChild(td);
            }
            frag.appendChild(tr);
        }

        // 底部占位
        if (last < st.totalRows) {
            const tr = document.createElement('tr');
            tr.className = 'vt-spacer';
            for (let i = 0; i <= st.cols.length; i++) {
                const td = spacerTd();
                if (i === 0) td.style.height = ((st.totalRows - last) * rowH) + 'px';
                tr.appendChild(td);
            }
            frag.appendChild(tr);
        }

        tbody.innerHTML = '';
        tbody.appendChild(frag);

        // 首帧后校准实际行高（不同 DPI / 字体下 28px 可能有偏差）
        if (tbody.rows.length > 0) {
            const sample = tbody.rows[first > 0 ? 1 : 0];
            const real = sample ? sample.getBoundingClientRect().height : 0;
            if (real > 4 && Math.abs(real - st.rowHeight) > 0.5) {
                st.rowHeight = real;
                if (force) { st.rendered = { start: -1, end: -1 }; renderSimTableRows(true); return; }
            }
        }

        updateSimTableStatus(first, last);
    }

    /** 定位状态栏：当前小时 / 可视区间 */
    function updateSimTableStatus(first, last) {
        const cur = document.getElementById('vtCurrentHour');
        const range = document.getElementById('vtRangeInfo');
        if (cur) cur.textContent = String(first);
        if (range) {
            range.textContent = `显示第 ${first} ~ ${Math.max(first, last - 1)} 小时 / 共 ${SimTable.totalRows} 小时（DOM 仅 ${Math.max(0, last - first)} 行）`;
        }
    }

    /** 跳转到指定小时（§7） */
    function jumpSimTableTo(hour) {
        const st = SimTable;
        if (!st.result) return;
        const h = Math.max(0, Math.min(st.totalRows - 1, Math.round(Number(hour) || 0)));
        const container = document.getElementById('simTableContainer');
        if (!container) return;
        container.scrollTop = h * st.rowHeight;
        renderSimTableRows(false);
        updateSimTableStatus(h, h + 1);
        const input = document.getElementById('vtJumpHour');
        if (input) input.value = String(h);
    }

    /** 逐时表翻页 / 跳转控件 */
    function initSimTableControls() {
        const container = document.getElementById('simTableContainer');
        if (!container) return;

        // 滚动 → 只重绘可视区间（rAF 节流，一帧最多一次）
        let rafPending = false;
        container.addEventListener('scroll', () => {
            if (rafPending) return;
            rafPending = true;
            requestAnimationFrame(() => {
                rafPending = false;
                renderSimTableRows(false);
            });
        }, { passive: true });

        // 容器尺寸变化时重算可视行数
        if (typeof ResizeObserver !== 'undefined') {
            new ResizeObserver(() => renderSimTableRows(true)).observe(container);
        }

        const on = (id, fn) => {
            const el = document.getElementById(id);
            if (el) el.addEventListener('click', fn);
        };
        on('vtFirst', () => jumpSimTableTo(0));
        on('vtPrev', () => jumpSimTableTo((SimTable.rendered.start || 0) - SimTable.PAGE_ROWS));
        on('vtNext', () => jumpSimTableTo((SimTable.rendered.start || 0) + SimTable.PAGE_ROWS));
        on('vtLast', () => jumpSimTableTo(SimTable.totalRows - 1));

        const input = document.getElementById('vtJumpHour');
        if (input) {
            input.addEventListener('change', () => jumpSimTableTo(input.value));
            input.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') jumpSimTableTo(input.value);
            });
        }
    }

    function updateChart() {
        return Utils.PerfTimer.measure('Chart.render', () => updateChartInner());
    }

    function updateChartInner() {
        const chartContainer = document.getElementById('chartContainer');
        const schemeIdx = parseInt(document.getElementById('chartSchemeSelect').value);
        const chartType = document.getElementById('chartTypeSelect').value;

        if (isNaN(schemeIdx) || !AppState.simulationResults[schemeIdx]) return;

        const simResult = AppState.simulationResults[schemeIdx];
        const DS = (typeof ResultDataStore !== 'undefined') ? ResultDataStore : null;
        if (!DS) { log('结果数据层未加载，图表不可用', 'error'); return; }

        // 结果层：与逐时表共用同一条结果记录（零拷贝）
        const store = ensureResultStore(simResult);

        // V2.3：不再 dispose + innerHTML='' 重建。
        // _initChart 会复用同一容器上的既有实例并 clear() 旧 option（任务书 §10）。
        const monthSelect = document.getElementById('chartMonthSelect');
        const month = monthSelect ? parseInt(monthSelect.value) : 0;

        switch (chartType) {
            case 'overview':
                // 全年 7 条曲线 → 降采样显示（仅绘图，不参与任何指标计算）
                AppState.currentChart = ChartModule.renderOverviewChart(chartContainer,
                    DS.ChartDataAdapter.getOverviewView(store));
                break;
            case 'hourly': {
                const colKey = document.getElementById('chartColumnSelect').value;
                AppState.currentChart = ChartModule.renderHourlyChart(chartContainer,
                    DS.ChartDataAdapter.getHourlyView(store, colKey), colKey);
                break;
            }
            case 'monthly':
                AppState.currentChart = ChartModule.renderMonthlyChart(chartContainer,
                    DS.ChartDataAdapter.getMonthlyView(store));
                break;
            case 'monthly-overview':
                AppState.currentChart = ChartModule.renderMonthlyOverviewChart(chartContainer,
                    DS.ChartDataAdapter.getMonthlyView(store));
                break;
            case 'monthly-detail':
                AppState.currentChart = ChartModule.renderMonthlyDetailChart(chartContainer,
                    DS.ChartDataAdapter.getRangeView(store, ChartModule.MONTH_HOURS[month],
                        ChartModule.MONTH_HOURS[month + 1] - ChartModule.MONTH_HOURS[month]), month);
                break;
            case 'module-monthly': {
                const mColKey = document.getElementById('chartColumnSelect').value;
                AppState.currentChart = ChartModule.renderModuleMonthlyChart(chartContainer,
                    DS.ChartDataAdapter.getMonthlyView(store), mColKey);
                break;
            }
            case 'typical-day': {
                const cum = ChartModule._monthCumulativeHours();
                AppState.currentChart = ChartModule.renderTypicalDayChart(chartContainer,
                    DS.ChartDataAdapter.getRangeView(store, cum[month], 24), month);
                break;
            }
            case 'typical-week': {
                const cum = ChartModule._monthCumulativeHours();
                AppState.currentChart = ChartModule.renderTypicalWeekChart(chartContainer,
                    DS.ChartDataAdapter.getRangeView(store, cum[month], 168), month);
                break;
            }
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
        // V2.3.1（任务书 §30）：导出确实需要对象数组，此处显式物化一次（仅导出场景）；
        // 总和行改用年度摘要，避免导出阶段再做 10 列 × 8760 次属性遍历
        const store = ensureResultStore(result);
        const hourlyData = ResultDataStore.materialize(store);
        const totals = ResultDataStore.getAnnualSummary(store).byLabel;

        const resultData = {
            hourlyData,
            totals,
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
            // V2.3：导出按需物化对象数组（仅导出场景，任务书 §5）
            const hourlyData = ResultDataStore.materialize(ensureResultStore(result));
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

        // V2.3：DocumentFragment 一次性插入（任务书 §21）
        const frag = document.createDocumentFragment();
        AppState.financeMergedData.forEach((row, i) => {
            const name = `方案${i + 1}: ${row['文件名称']}`;
            const opt = document.createElement('option');
            opt.value = String(i);
            opt.textContent = name;
            frag.appendChild(opt);
        });
        sel.innerHTML = '';
        sel.appendChild(frag);

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
            eirr: Utils.toNum(optEl('optWeightEirr').value, 40) / 100,
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
        const sum = (w.eirr + w.lcoh + w.curtailmentRate) * 100;
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
            // V2.3.1（任务书 §9 / §35）：输入数据版本参与 simulationKey，
            // 换数据后旧评价缓存自动失效
            inputVersion: AppState.inputVersion,
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
        ['optWeightEirr', 'optWeightLcoh', 'optWeightCurtail'].forEach(id => {
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
        optLog('已启用 Web Worker：NSGA-II + 8760 仿真 + 概算 + 财务评价全部在后台线程执行', 'info');

        let gotProgress = false;
        let failed = false;

        /**
         * Worker 启动/运行失败诊断（任务书 §17）。
         * 不静默回退：把失败原因分类后写入日志，便于定位是
         * 「脚本加载被拦截」「依赖缺失」还是「运行期异常」。
         */
        const diagnose = (e) => {
            const msg = (e && (e.message || e.reason)) || '';
            if (/cannot be accessed from origin|SecurityError|file:\/\//i.test(msg)) {
                return 'Worker 脚本被浏览器安全策略拦截（file:// 直开属正常现象）';
            }
            if (/importScripts|404|Failed to fetch|NetworkError/i.test(msg)) {
                return 'Worker 脚本或其依赖加载失败';
            }
            if (/SyntaxError|ReferenceError|TypeError/i.test(msg)) {
                return 'Worker 运行期异常';
            }
            return msg || 'Worker 未在预期时间内响应';
        };

        const giveUp = (reason, type) => {
            if (failed || gotProgress) return;
            failed = true;
            if (OptState.fallbackTimer) clearTimeout(OptState.fallbackTimer);
            try { worker.terminate(); } catch (e) { /* ignore */ }
            OptState.worker = null;
            optLog('Web Worker 不可用：' + reason, type || 'warn');
            optLog('已回退到主线程分片计算 —— ⚠ 性能可能下降，页面仍保持响应', 'warn');
            startWithMainThread(config, ctx);
        };

        // 兜底：若 Worker 长时间无响应（file:// 环境常见），自动回退
        OptState.fallbackTimer = setTimeout(() => giveUp(diagnose(null)), 4000);

        worker.onerror = (e) => {
            giveUp(diagnose(e), 'error');
        };

        worker.onmessageerror = () => {
            giveUp('Worker 消息反序列化失败（数据结构不可结构化克隆）', 'error');
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

        // 数据传输策略（任务书 §15）：
        //   经评估**不使用 Transferable**。原因：一旦 transfer，pvData/windData 的
        //   ArrayBuffer 会被 detach，用户「取消优化」后主线程原始数据即不可用，
        //   且无法二次发起优化。8760 × 8 字节 × 2 ≈ 140 KB，结构化克隆的开销可忽略，
        //   不值得为它牺牲数据可用性。
        //   若未来数据规模显著增大，应改为「主线程持有不可变副本 + 传输副本」的方案。
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
                    inputVersion: ctx.inputVersion,
                },
            },
        });
    }

    /**
     * 路径 B：主线程分片调度（仅在 Worker 不可用时使用）。
     *
     * 任务书 §16：时间片由 60ms 降至 12ms —— 单次长任务会阻塞输入与滚动，
     * 更短的时间片能让 UI 在每代之间获得调度机会，保持基本响应。
     */
    function startWithMainThread(config, ctx) {
        OptState.mode = 'main';
        optEl('optEngineMode').textContent = '计算模式：主线程 Fallback ⚠ 性能可能下降';
        optEl('optEngineMode').style.color = 'var(--accent-yellow, #e3b341)';

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

        const CHUNK_MS = 12;          // 任务书 §16：10~15ms
        const perf = { startedAt: Date.now(), initMs: 0, generationCount: 0, generationTimeMs: 0 };

        const tick = () => {
            if (OptState.cancelled) {
                onOptimizationCancelled();
                return;
            }
            try {
                const t0 = performance.now();
                do {
                    if (session.isFinished()) break;
                    const g0 = performance.now();
                    session.runNextGeneration();
                    perf.generationTimeMs += performance.now() - g0;
                    perf.generationCount++;
                } while (!session.isFinished() && (performance.now() - t0) < CHUNK_MS);
            } catch (err) {
                onOptimizationError(err);
                return;
            }

            if (session.isFinished()) {
                onOptimizationFinished(attachMainThreadPerf(session.getResult(), perf));
            } else {
                setTimeout(tick, 0);   // 让出主线程 → UI 保持响应
            }
        };

        // 先让 UI 完成一次渲染，再执行耗时的种群初始化
        setTimeout(() => {
            if (OptState.cancelled) { onOptimizationCancelled(); return; }
            try {
                const initT0 = Date.now();
                session.ensureInitialized();
                perf.initMs = Date.now() - initT0;
            } catch (err) {
                onOptimizationError(err);
                return;
            }
            if (session.isFinished()) {
                onOptimizationFinished(attachMainThreadPerf(session.getResult(), perf));
            }
            else setTimeout(tick, 0);
        }, 30);
    }

    /** 为主线程 Fallback 的结果补充同一结构的性能统计（任务书 §14） */
    function attachMainThreadPerf(result, perf) {
        if (!result) return result;
        const st = result.statistics || {};
        const evaluated = st.totalEvaluated || 0;
        const cacheHits = st.cacheHits || 0;
        const total = evaluated + cacheHits;
        const totalMs = Date.now() - perf.startedAt;
        result.performanceStats = {
            mode: 'main',
            initMs: perf.initMs,
            totalMs: totalMs,
            totalSeconds: Math.round(totalMs / 100) / 10,
            generationsCompleted: st.generationsCompleted || perf.generationCount,
            avgGenerationMs: perf.generationCount > 0 ? Math.round(perf.generationTimeMs / perf.generationCount) : 0,
            totalEvaluated: evaluated,
            cacheHits: cacheHits,
            cacheHitRate: total > 0 ? Math.round((cacheHits / total) * 1000) / 10 : 0,
            avgEvaluateMs: 0,
            avgSimulateMs: 0,
            storesHourlyResults: false,
        };
        return result;
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
            bestEirr: snap.bestEirr,
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
            `当前最佳EIRR <b>${optPct(snap.bestEirr)}</b>`,
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

        // 性能诊断（任务书 §14）：优化完成后统一汇报，便于持续优化
        const pf = result.performanceStats;
        if (pf) {
            optLog(`性能诊断：模式=${pf.mode === 'worker' ? 'Web Worker' : '主线程 Fallback'}` +
                   `，初始化 ${optNum(pf.initMs, 0)} ms，总耗时 ${optNum(pf.totalSeconds, 1)} s` +
                   `，平均每代 ${optNum(pf.avgGenerationMs, 0)} ms`, 'info');
            optLog(`性能诊断：评价方案 ${pf.totalEvaluated} 个，缓存命中 ${pf.cacheHits} 次（命中率 ${optNum(pf.cacheHitRate, 1)}%）` +
                   (pf.avgEvaluateMs ? `，平均单方案评价 ${optNum(pf.avgEvaluateMs, 2)} ms` +
                    `（其中 8760 仿真 ${optNum(pf.avgSimulateMs, 2)} ms）` : ''), 'info');
            optLog('内存审计：优化结果只保存技术/经济指标与方案参数，不保存 8760 小时逐时结果（§11）', 'info');
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
            目标函数为：最大化资本金内部收益率（EIRR）、最小化 LCOH（全生命周期折现口径，折现率 ${optNum((OptState.result.config.lcoh || {}).discountRate, 2)}%）、最小化新能源弃电率；
            综合推荐方案由权重（EIRR ${optNum(OptState.result.config.recommendationWeights.eirr * 100, 0)}% /
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
                <div class="opt-card-row"><span>资本金收益率 EIRR</span><b>${optPct(ec.EIRR)}</b></div>
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
            ${card('💰 方案A 经济最优（EIRR最高）', 'is-economic', rep.economicBest)}
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

            const de = (base.economic.EIRR !== null && rep.recommended.economic.EIRR !== null)
                ? rep.recommended.economic.EIRR - base.economic.EIRR : null;
            const dl = (isFinite(base.economic.LCOH) && isFinite(rep.recommended.economic.LCOH))
                ? rep.recommended.economic.LCOH - base.economic.LCOH : null;
            const dc = (rep.recommended.technical.curtailmentRate - base.technical.curtailmentRate) * 100;

            compare = `<div class="opt-note">
                    <b>相较基准方案（Baseline）：</b>
                    EIRR ${de === null ? '—' : (de >= 0 ? '提高 +' : '降低 ') + de.toFixed(2) + ' 个百分点'}；
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
        else if (filter === 'economic') list = list.filter(s => s.economic.EIRR !== null && isFinite(s.economic.EIRR) && s.economic.EIRR >= 0);

        const sort = optEl('optParetoSort').value;
        const num = v => (v === null || v === undefined || !isFinite(Number(v))) ? (sort === 'lcoh' || sort === 'curtailment' || sort === 'investment' ? Infinity : -Infinity) : Number(v);
        const cmp = {
            score: (a, b) => num(b.score) - num(a.score),
            eirr: (a, b) => num(b.economic.EIRR) - num(a.economic.EIRR),
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
            ['EIRR(%)', s => optNum(s.economic.EIRR, 2)],
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
                   '<th>ParetoCount</th><th>BestEIRR</th><th>BestLCOH</th><th>BestCurtailment</th></tr></thead><tbody>';
        for (const h of history) {
            html += `<tr>
                <td>${h.generation}</td>
                <td>${optInt(h.evaluatedCount)}</td>
                <td>${optInt(h.feasibleCount)}</td>
                <td>${optInt(h.paretoCount)}</td>
                <td>${optPct(h.bestEirr)}</td>
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
        ['optParetoChart', 'optConvEirr', 'optConvLcoh', 'optConvCurtail'].forEach(id => {
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
            ['optConvEirr', 'eirr'],
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
        initSimTableControls();

        document.getElementById('btnClearLog').addEventListener('click', clearLog);

        // 测试钩子：仅供 tests/*.js 驱动内部行为（非业务接口）
        window.__wbTestHooks = {
            log: log,
            clearLog: clearLog,
            /** 仅供测试：注入仿真结果并刷新方案选择器（绕开文件读取，便于度量渲染耗时） */
            __injectSimulationResults: (results) => onSimulationComplete(results),
        };

        // =====================================================================
        // V2.3.1 性能治理（任务书 §20 / §21 / §28）
        // =====================================================================

        // 性能计时开关：默认关闭；URL 带 ?perf=1 或控制台置 DEBUG_PERFORMANCE 时开启
        const wantsPerf = (function () {
            try {
                if (window.DEBUG_PERFORMANCE === true) return true;
                return /([?&])perf=1\b/.test(window.location.search || '');
            } catch (e) { return false; }
        })();
        if (wantsPerf) {
            Utils.PerfTimer.enable();
            log('[PERF] 性能计时已开启（DEBUG_PERFORMANCE）', 'warn');

            // 长任务检测（§21）：>50ms 记录一条；由 PerfTimer.enabled 控制生命周期，
            // 不做永久高频输出
            if (typeof PerformanceObserver !== 'undefined') {
                try {
                    const po = new PerformanceObserver((list) => {
                        const entries = list.getEntries() || [];
                        for (const e of entries) {
                            if (e.duration >= 50) {
                                log('Long Task detected: duration = ' + e.duration.toFixed(1) + ' ms', 'warn');
                            }
                        }
                    });
                    po.observe({ entryTypes: ['longtask'] });
                } catch (e) { /* 浏览器不支持 longtask 条目，忽略 */ }
            }
        }

        // 窗口缩放 → rAF 防抖后 resize（§28）：避免拖动过程中高频触发 chart.resize()
        let resizePending = false;
        window.addEventListener('resize', () => {
            if (resizePending) return;
            resizePending = true;
            requestAnimationFrame(() => {
                resizePending = false;
                if (AppState.currentChart && !AppState.currentChart.isDisposed()) AppState.currentChart.resize();
                if (OptState.paretoChart && !OptState.paretoChart.isDisposed()) OptState.paretoChart.resize();
                for (const c of OptState.convCharts) {
                    if (c && !c.isDisposed()) c.resize();
                }
            });
        }, { passive: true });

        log('多能互补风光储氢分析软件 WEB-V2.3.1 已就绪（数据访问层与仿真引擎性能审计版）', 'success');
        log('V2.3 架构：计算层（Float64Array）→ 结果层（ResultDataStore）→ 显示层（图表/表格/Excel 按需读取）');
        log('　· 8760 小时数据只计算并保存一份；逐时表采用虚拟滚动，DOM 只保留可视区间');
        log('　· 图表显示数据允许降采样（≤1500 点），指标计算始终使用全分辨率');
        log('　· 优化阶段只保存指标，不保存 8760 逐时结果；复核时按需重算');
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
