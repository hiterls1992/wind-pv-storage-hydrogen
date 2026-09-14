/**
 * 工具函数模块
 * 提供数值格式化、时间戳生成等通用功能
 */

const Utils = {
    /**
     * 格式化数字显示
     */
    formatNumber(value, decimals = 2) {
        if (value === null || value === undefined || isNaN(value)) return 'N/A';
        if (typeof value === 'string') value = parseFloat(value);
        if (isNaN(value)) return 'N/A';

        if (Math.abs(value) >= 100000) {
            return value.toLocaleString('zh-CN', { maximumFractionDigits: 0 });
        } else if (Math.abs(value) >= 100) {
            return value.toLocaleString('zh-CN', { maximumFractionDigits: 1 });
        } else if (Math.abs(value) >= 1) {
            return value.toLocaleString('zh-CN', { maximumFractionDigits: decimals });
        } else {
            return value.toFixed(4);
        }
    },

    /**
     * 格式化百分比
     */
    formatPercent(value, decimals = 2) {
        if (value === null || value === undefined || isNaN(value)) return 'N/A';
        return value.toFixed(decimals) + '%';
    },

    /**
     * 生成时间戳字符串
     */
    timestamp() {
        const now = new Date();
        const pad = (n, l = 2) => String(n).padStart(l, '0');
        return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
    },

    /**
     * 生成时间显示字符串 HH:MM:SS
     */
    timeNow() {
        const now = new Date();
        const pad = n => String(n).padStart(2, '0');
        return `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
    },

    /**
     * 安全读取数值
     */
    toNum(val, defaultVal = 0) {
        const n = parseFloat(val);
        return isNaN(n) ? defaultVal : n;
    },

    /**
     * 生成参数组合（笛卡尔积）
     */
    cartesianProduct(...arrays) {
        return arrays.reduce((acc, arr) => {
            const result = [];
            acc.forEach(a => arr.forEach(b => result.push([...a, b])));
            return result;
        }, [[]]);
    },

    /**
     * 处理步长为0的情况
     */
    getValues(min, max, step) {
        if (step === 0 && min === max) return [min];
        if (step <= 0) return [min];
        const result = [];
        for (let v = min; v <= max + step * 0.001; v += step) {
            result.push(Math.round(v * 1e6) / 1e6); // 避免浮点误差
        }
        return result;
    },

    /**
     * 深拷贝对象
     */
    deepClone(obj) {
        return JSON.parse(JSON.stringify(obj));
    },

    /**
     * 下载二进制数据
     */
    downloadBlob(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    },

    /**
     * 下载base64数据
     */
    downloadBase64(base64, filename, mimeType = 'image/png') {
        const link = document.createElement('a');
        link.href = `data:${mimeType};base64,${base64}`;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }
};

/**
 * ============================================================================
 * 统一性能计时器（V2.3.1 任务书 §20）
 * ============================================================================
 *
 * 用法：
 *   PerfTimer.start('Simulation');  ...耗时操作...;  PerfTimer.end('Simulation');
 *   PerfTimer.measure('Summary', () => DataSummary.generateSummary(...));
 *
 * 输出：
 *   [PERF] Simulation: 8.2 ms
 *   [PERF] Table.render: 120.3 ms  (long task)      —— 超过 50ms 单独标注
 *
 * 开关：生产环境默认关闭（enabled = false），不产生任何输出与计时开销。
 * 打开方式（任一）：
 *   · URL 加 ?perf=1
 *   · 控制台执行 Utils.PerfTimer.enabled = true
 */
const PerfTimer = {
    /** 默认关闭（任务书 §20：生产环境可通过 DEBUG_PERFORMANCE 控制） */
    enabled: false,
    _marks: {},

    /** 开启（返回 this 便于链式） */
    enable() { this.enabled = true; return this; },
    disable() { this.enabled = false; return this; },

    start(name) {
        if (!this.enabled) return;
        this._marks[name] = performance.now();
    },

    end(name) {
        if (!this.enabled) return;
        const t0 = this._marks[name];
        if (t0 === undefined) return;
        delete this._marks[name];
        this._emit(name, performance.now() - t0);
    },

    /** 同步测量：无论开关状态都执行 fn，仅在开启时输出 */
    measure(name, fn) {
        if (!this.enabled) return fn();
        const t0 = performance.now();
        try {
            return fn();
        } finally {
            this._emit(name, performance.now() - t0);
        }
    },

    _emit(name, ms) {
        const tag = ms >= 50 ? '  (long task)' : '';
        const fn = ms >= 50 ? console.warn : console.log;
        fn.call(console, '[PERF] ' + name + ': ' + ms.toFixed(1) + ' ms' + tag);
    },
};

Utils.PerfTimer = PerfTimer;

// 导出给其他模块使用（浏览器主线程 window / Web Worker self 通用）
self.Utils = Utils;
