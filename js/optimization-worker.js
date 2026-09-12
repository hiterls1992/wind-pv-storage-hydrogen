/**
 * ============================================================================
 * 优化计算 Web Worker  ——  V2.1 新增
 * ============================================================================
 *
 * 职责划分：
 *   主线程（app.js）  ：仅负责 UI、参数读取、结果渲染
 *   Worker（本文件） ：承担全部 NSGA-II 计算与 8760 小时仿真评价
 *
 * 通信协议（postMessage）
 * ----------------------
 *   主线程 → Worker
 *     { type: 'start',  payload: { config, context } }
 *     { type: 'cancel' }
 *   Worker → 主线程
 *     { type: 'ready' }                                   Worker 就绪
 *     { type: 'progress',   progress: snapshot }           每代进度
 *     { type: 'generation', progress: snapshot }           每代历史
 *     { type: 'complete',   result: optimizationResult }   优化完成
 *     { type: 'error',      message: string }              异常（主线程捕获提示，页面不崩溃）
 *     { type: 'warn',       message: string }              非致命告警（如数据长度非 8760）
 *
 * 说明：本文件只在 Worker 环境生效；若被误引入主线程则静默退出。
 */

/* eslint-disable no-restricted-globals */

if (typeof importScripts !== 'function') {
    // 主线程误加载：本文件仅供 Worker 使用，不做任何事。
} else {

    // 复用现有全部计算模块（顺序不可调整：optimization-engine 依赖前者）
    importScripts(
        'utils.js',
        'simulation-engine.js',
        'data-summary.js',
        'estimate.js',
        'finance-engine.js',
        'optimization-engine.js'
    );

    /** 当前优化会话 */
    var currentSession = null;

    /** 把主线程传入的数据统一还原为 Float64Array */
    function toFloat64(v, name) {
        if (v instanceof Float64Array) return v;
        if (v instanceof ArrayBuffer) return new Float64Array(v);
        if (Array.isArray(v)) return new Float64Array(v);
        throw new Error(name + ' 数据格式不正确（应为 Float64Array / ArrayBuffer / Array）');
    }

    function postError(err) {
        var msg = (err && err.message) ? err.message : String(err);
        self.postMessage({ type: 'error', message: msg });
    }

    /** 启动一次优化 */
    function start(payload) {
        try {
            if (!payload || !payload.config || !payload.context) {
                throw new Error('优化启动参数不完整');
            }

            var raw = payload.context;
            var ctx = {
                pvData: toFloat64(raw.pvData !== undefined ? raw.pvData : raw.pv, '光伏 8760'),
                windData: toFloat64(raw.windData !== undefined ? raw.windData : raw.wind, '风电 8760'),
                simParams: raw.simParams,
                prices: raw.prices,
                financeParams: raw.financeParams,
                lcohDiscountRate: raw.lcohDiscountRate,
                onWarn: function (msg) { self.postMessage({ type: 'warn', message: msg }); }
            };

            currentSession = OptimizationEngine.createSession({
                config: payload.config,
                context: ctx,
                onProgress: function (snap) { self.postMessage({ type: 'progress', progress: snap }); },
                onGeneration: function (snap) { self.postMessage({ type: 'generation', progress: snap }); }
            });

            currentSession.ensureInitialized();
            tick();
        } catch (err) {
            currentSession = null;
            postError(err);
        }
    }

    /**
     * 分代推进：每代结束后主动让出事件循环，
     * 使 Worker 仍能接收并响应主线程的 'cancel' 消息。
     */
    function tick() {
        if (!currentSession) return;
        try {
            if (currentSession.isCancelled() || currentSession.isFinished()) {
                var result = currentSession.getResult();
                currentSession = null;
                self.postMessage({ type: 'complete', result: result });
                return;
            }
            currentSession.runNextGeneration();
            setTimeout(tick, 0);
        } catch (err) {
            currentSession = null;
            postError(err);
        }
    }

    self.onmessage = function (e) {
        var d = e.data || {};

        if (d.type === 'start') {
            start(d.payload);
            return;
        }

        if (d.type === 'cancel') {
            if (currentSession) {
                currentSession.cancel();
            } else {
                self.postMessage({ type: 'ready' });
            }
        }
    };

    // 就绪通知
    self.postMessage({ type: 'ready' });
}
