/**
 * 数据汇总模块
 * 从 data_summary.py 翻译，汇总仿真结果的关键指标
 */

const DataSummary = {
    /**
     * 生成汇总数据
     * @param {Array} simulationResults - 所有方案的仿真结果数组
     * @param {number} pvCapacity - 光伏容量 MW
     * @param {number} windCapacity - 风电容量 MW
     * @returns {Array} 汇总数据数组
     */
    generateSummary(simulationResults, pvCapacity, windCapacity) {
        const summaryData = [];

        for (const simResult of simulationResults) {
            const { sums, systemVars, filename } = simResult;

            // 制氢量 kg → 万吨
            const totalH2ProdTenThousandTons = Math.round(sums.sumH2Prod / 1e7 * 1000) / 1000;

            // 电量 MWh → 亿度
            const totalH2PowerBillion = Math.round(sums.sumHydrogenPower / 1e5 * 1e5) / 1e5;
            const totalExportBillion = Math.round(sums.sumExport / 1e5 * 1e5) / 1e5;
            const totalImportBillion = Math.round(sums.sumImport / 1e5 * 1e5) / 1e5;
            const totalCurtailmentBillion = Math.round(sums.sumCurtailment / 1e5 * 1e5) / 1e5;

            const electrolyzerCapacity = systemVars['电解槽容量（MW）'];
            const storagePower = systemVars['储能功率（MW）'];
            const storageDuration = systemVars['储能时长（小时）'];

            // 电解槽利用小时数
            let electrolyzerHours = 0;
            if (electrolyzerCapacity > 0) {
                electrolyzerHours = Math.round(
                    (totalH2PowerBillion / (electrolyzerCapacity * 24 * 365 / 100000)) * 8760 * 100
                ) / 100;
            }

            // 比值
            const ratioH2 = sums.sumTotal !== 0 ? sums.sumHydrogenPower / sums.sumTotal : 0;
            const ratioExport = sums.sumTotal !== 0 ? sums.sumExport / sums.sumTotal : 0;
            const ratioImport = sums.sumTotal !== 0 ? sums.sumImport / sums.sumTotal : 0;
            const ratioCurtailment = sums.sumTotal !== 0 ? sums.sumCurtailment / sums.sumTotal : 0;

            summaryData.push({
                '文件名称': filename,
                '光伏容量（MW）': pvCapacity,
                '风电容量（MW）': windCapacity,
                '电解槽容量（MW）': electrolyzerCapacity,
                '储能功率（MW）': storagePower,
                '储能时长（小时）': storageDuration,
                '制氢电量总和（亿度）': totalH2PowerBillion,
                '上网电量总和（亿度）': totalExportBillion,
                '下网电量总和（亿度）': totalImportBillion,
                '弃电量总和（亿度）': totalCurtailmentBillion,
                '制氢量总和（万吨）': totalH2ProdTenThousandTons,
                '制氢电量比例': (ratioH2 * 100).toFixed(2) + '%',
                '上网电量比例': (ratioExport * 100).toFixed(2) + '%',
                '下网电量比例': (ratioImport * 100).toFixed(2) + '%',
                '弃电电量比例': (ratioCurtailment * 100).toFixed(2) + '%',
                '电解槽利用小时数': electrolyzerHours,
                // 保留原始数据供概算/经评模块使用
                '_raw': sums,
                '_systemVars': systemVars,
            });
        }

        return summaryData;
    }
};

// 导出（浏览器主线程 window / Web Worker self 通用）
self.DataSummary = DataSummary;
