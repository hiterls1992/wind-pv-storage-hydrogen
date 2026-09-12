/**
 * 概算模块
 * 从 estimate_module.py 翻译，计算各方案的工程投资
 */

const Estimate = {
    /**
     * 批量概算分析
     * @param {Array} summaryData - 数据汇总数据
     * @param {Object} prices - 设备单价参数
     * @returns {Array} 概算结果数组
     */
    batchEstimate(summaryData, prices) {
        const windPrice = Utils.toNum(prices.windUnitPrice, 3800);
        const pvPrice = Utils.toNum(prices.pvUnitPrice, 2500);
        const storagePrice = Utils.toNum(prices.storageUnitPrice, 800);
        const electrolyzerPrice = Utils.toNum(prices.electrolyzerUnitPrice, 1600);
        const transmission = Utils.toNum(prices.transmissionCost, 10000);
        const land = Utils.toNum(prices.landCost, 10000);
        const otherRatio = Utils.toNum(prices.otherFacilitiesRatio, 70) / 100;

        const results = [];

        for (const row of summaryData) {
            const electrolyzerCapacity = Utils.toNum(row['电解槽容量（MW）']);
            const storagePower = Utils.toNum(row['储能功率（MW）']);
            const storageDuration = Utils.toNum(row['储能时长（小时）']);
            const pvCapacity = Utils.toNum(row['光伏容量（MW）']);
            const windCapacity = Utils.toNum(row['风电容量（MW）']);

            // 各部分成本（万元）
            const windTotal = windCapacity * 1000 * windPrice / 10000;
            const pvTotal = pvCapacity * 1000 * pvPrice / 10000;
            const storageCapacity = storagePower * storageDuration * 1000; // kWh
            const storageTotal = storageCapacity * storagePrice / 10000;
            const electrolyzerTotal = electrolyzerCapacity * 1000 * electrolyzerPrice / 10000;
            const otherFacilities = electrolyzerTotal * otherRatio;

            // 建设总投资
            const totalInvestment = windTotal + pvTotal + storageTotal + electrolyzerTotal + transmission + land + otherFacilities;

            // 先携带源数据汇总文件中的全部运营字段（供下游经评模块直接使用，无需再读汇总文件）
            // 过滤掉内部字段（以 _ 开头）
            const result = {};
            for (const k of Object.keys(row)) {
                if (!k.startsWith('_')) result[k] = row[k];
            }

            // 追加概算计算字段（运营字段在前，投资字段在后）
            result['风电单价（元/kW）'] = windPrice;
            result['光伏单价（元/kW）'] = pvPrice;
            result['储能单价（元/kWh）'] = storagePrice;
            result['电解槽单价（元/kW）'] = electrolyzerPrice;
            result['送出线路及升压站（万元）'] = transmission;
            result['配套设施及土地（万元）'] = land;
            result['其他配套设施比例系数（%）'] = otherRatio * 100;
            result['风电总价（万元）'] = Math.round(windTotal * 100) / 100;
            result['光伏总价（万元）'] = Math.round(pvTotal * 100) / 100;
            result['储能总价（万元）'] = Math.round(storageTotal * 100) / 100;
            result['电解槽总价（万元）'] = Math.round(electrolyzerTotal * 100) / 100;
            result['其他配套设施（万元）'] = Math.round(otherFacilities * 100) / 100;
            result['建设总投资（万元）'] = Math.round(totalInvestment * 100) / 100;

            results.push(result);
        }

        return results;
    }
};

// 导出（浏览器主线程 window / Web Worker self 通用）
self.Estimate = Estimate;
