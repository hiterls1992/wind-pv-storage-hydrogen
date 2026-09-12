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

// 导出给其他模块使用（浏览器主线程 window / Web Worker self 通用）
self.Utils = Utils;
