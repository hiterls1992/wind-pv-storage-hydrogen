/**
 * 仿真计算引擎
 * 从 simulation_logic.py 1:1 翻译为纯 JavaScript
 * 设计为在 Web Worker 中运行
 */

/**
 * 执行单个参数组合的仿真计算
 * @param {Float64Array} pvData - 光伏单MW电量数据 (8760小时, MWh/MW)
 * @param {Float64Array} windData - 风电单MW电量数据 (8760小时, MWh/MW)
 * @param {Object} params - 仿真参数
 * @param {number} params.PV_CAPACITY - 光伏容量 MW
 * @param {number} params.WIND_CAPACITY - 风电容量 MW
 * @param {number} params.ELECTROLYZER_MIN_RATIO - 电解槽最低运行比例
 * @param {number} params.MAX_EXPORT_RATIO_HOURLY - 每小时上网比例上限
 * @param {number} params.MAX_EXPORT_RATIO_TOTAL - 总量上网比例上限
 * @param {number} params.MAX_IMPORT_RATIO - 下网电量比例上限
 * @param {number} params.STORAGE_CHARGE_EFFICIENCY - 储能充电效率
 * @param {number} params.STORAGE_DISCHARGE_EFFICIENCY - 储能放电效率
 * @param {number} params.HYDROGEN_ENERGY_CONSUMPTION - 制氢电耗 kWh/kg
 * @param {number} storagePower - 储能功率 MW
 * @param {number} storageDuration - 储能时长 小时
 * @param {number} electrolyzerCapacity - 电解槽容量 MW
 * @returns {Object} 仿真结果
 */
function runSingleSimulation(pvData, windData, params, storagePower, storageDuration, electrolyzerCapacity) {
    const PV_CAPACITY = params.PV_CAPACITY;
    const WIND_CAPACITY = params.WIND_CAPACITY;
    const ELECTROLYZER_MIN_RATIO = params.ELECTROLYZER_MIN_RATIO;
    const MAX_EXPORT_RATIO_HOURLY = params.MAX_EXPORT_RATIO_HOURLY;
    const MAX_EXPORT_RATIO_TOTAL = params.MAX_EXPORT_RATIO_TOTAL;
    const MAX_IMPORT_RATIO = params.MAX_IMPORT_RATIO;
    const CHARGE_EFF = params.STORAGE_CHARGE_EFFICIENCY;
    const DISCHARGE_EFF = params.STORAGE_DISCHARGE_EFFICIENCY;
    const H2_CONSUMPTION = params.HYDROGEN_ENERGY_CONSUMPTION;

    const TOTAL_CAPACITY = PV_CAPACITY + WIND_CAPACITY;
    const ELECTROLYZER_MIN = electrolyzerCapacity * ELECTROLYZER_MIN_RATIO;
    const MAX_EXPORT_HOURLY = TOTAL_CAPACITY * MAX_EXPORT_RATIO_HOURLY;
    const STORAGE_CAPACITY = storagePower * storageDuration;

    const hours = pvData.length;

    // 预分配结果数组（使用扁平数组提升性能）
    // 每小时 11 个字段
    const COLS = 11;
    const results = new Float64Array(hours * COLS);
    // 列顺序: 0=光伏 1=风电 2=合计 3=充电 4=放电 5=储能现存 6=制氢电量 7=上网 8=下网 9=弃电 10=制氢量

    let storageRemaining = 0.0;
    let totalGenerated = 0.0;
    let totalExported = 0.0;

    for (let h = 0; h < hours; h++) {
        const pv = pvData[h] * PV_CAPACITY;
        const wind = windData[h] * WIND_CAPACITY;
        const total = pv + wind;
        totalGenerated += total;

        const idx = h * COLS;
        results[idx + 0] = pv;     // 光伏电量
        results[idx + 1] = wind;   // 风电电量
        results[idx + 2] = total;  // 合计电量

        let charge = 0, discharge = 0, hydrogenPower = 0, export_ = 0, importPower = 0, curtailment = 0, hydrogenProduction = 0;

        if (total >= ELECTROLYZER_MIN) {
            // 情况1：风光电量足够维持最低制氢功率
            hydrogenPower = Math.min(total, electrolyzerCapacity);
            let remaining = total - hydrogenPower;

            // 储能充电
            const chargeCapacityLimit = (STORAGE_CAPACITY - storageRemaining) / CHARGE_EFF;
            const chargePowerLimit = storagePower;
            const chargePossible = Math.min(chargeCapacityLimit, chargePowerLimit);
            charge = Math.min(remaining, chargePossible);
            storageRemaining += charge * CHARGE_EFF;
            let remainingAfterCharge = remaining - charge;

            // 上网电量
            const exportHourlyLimit = MAX_EXPORT_HOURLY;
            const exportTotalLimit = totalGenerated * MAX_EXPORT_RATIO_TOTAL;
            const exportRemainingQuota = exportTotalLimit - totalExported;
            export_ = Math.min(remainingAfterCharge, exportHourlyLimit, exportRemainingQuota);
            curtailment = remainingAfterCharge - export_;

            totalExported += export_;

            hydrogenProduction = hydrogenPower * 1000 / H2_CONSUMPTION;
        } else {
            // 情况2：风光电量不足
            const needed = ELECTROLYZER_MIN - total;

            // 储能放电
            const dischargeCapacityLimit = storageRemaining * DISCHARGE_EFF;
            const dischargePowerLimit = storagePower;
            const dischargePossible = Math.min(dischargeCapacityLimit, dischargePowerLimit);
            discharge = Math.min(needed, dischargePossible);
            storageRemaining -= discharge / DISCHARGE_EFF;

            const remainingNeeded = needed - discharge;

            // 下网电量
            importPower = Math.min(remainingNeeded, TOTAL_CAPACITY * MAX_IMPORT_RATIO);
            hydrogenPower = total + discharge + importPower;

            hydrogenProduction = hydrogenPower * 1000 / H2_CONSUMPTION;
        }

        results[idx + 3] = charge;
        results[idx + 4] = discharge;
        results[idx + 5] = storageRemaining;
        results[idx + 6] = hydrogenPower;
        results[idx + 7] = export_;
        results[idx + 8] = importPower;
        results[idx + 9] = curtailment;
        results[idx + 10] = hydrogenProduction;
    }

    // 计算比值
    let sumTotal = 0, sumHydrogenPower = 0, sumExport = 0, sumImport = 0, sumCurtailment = 0, sumH2Prod = 0;
    for (let h = 0; h < hours; h++) {
        const idx = h * COLS;
        sumTotal += results[idx + 2];
        sumHydrogenPower += results[idx + 6];
        sumExport += results[idx + 7];
        sumImport += results[idx + 8];
        sumCurtailment += results[idx + 9];
        sumH2Prod += results[idx + 10];
    }

    const ratioData = [
        { '指标': '总制氢电量/总合计电量', '比值': sumTotal !== 0 ? sumHydrogenPower / sumTotal : 0 },
        { '指标': '总上网电量/总合计电量', '比值': sumTotal !== 0 ? sumExport / sumTotal : 0 },
        { '指标': '总下网电量/总合计电量', '比值': sumTotal !== 0 ? sumImport / sumTotal : 0 },
        { '指标': '总弃电电量/总合计电量', '比值': sumTotal !== 0 ? sumCurtailment / sumTotal : 0 },
    ];

    // 系统参数
    const systemVars = {
        '光伏容量（MW）': PV_CAPACITY,
        '风电容量（MW）': WIND_CAPACITY,
        '储能功率（MW）': storagePower,
        '储能时长（小时）': storageDuration,
        '电解槽容量（MW）': electrolyzerCapacity,
        '电解槽最低运行比例': ELECTROLYZER_MIN_RATIO,
        '原每小时上网比例上限': MAX_EXPORT_RATIO_HOURLY,
        '新增总量上网比例上限': MAX_EXPORT_RATIO_TOTAL,
        '下网电量比例上限': MAX_IMPORT_RATIO,
        '储能充电效率': CHARGE_EFF,
        '储能放电效率': DISCHARGE_EFF,
        '每千克氢气耗电量（kWh/kg）': H2_CONSUMPTION,
        '总容量（MW）': TOTAL_CAPACITY,
        '电解槽最小运行功率（MW）': ELECTROLYZER_MIN,
        '原每小时上网电量绝对上限（MW）': MAX_EXPORT_HOURLY,
        '储能总容量（MWh）': STORAGE_CAPACITY,
    };

    return {
        results,       // Float64Array (hours * 11)
        ratioData,
        systemVars,
        sums: { sumTotal, sumHydrogenPower, sumExport, sumImport, sumCurtailment, sumH2Prod },
        filename: `OUTPUT-${electrolyzerCapacity}MW-${storagePower}MW-${storageDuration}H-${MAX_EXPORT_RATIO_TOTAL.toFixed(1)}.xlsx`
    };
}

// 如果在 Worker 中运行，监听消息
if (typeof self !== 'undefined' && typeof window === 'undefined') {
    // Web Worker 环境
    self.onmessage = function(e) {
        const { type, params, inputData, storagePowerValues, storageDurationValues, electrolyzerValues } = e.data;

        if (type === 'simulate') {
            const pvData = new Float64Array(inputData.pv);
            const windData = new Float64Array(inputData.wind);
            const totalCombinations = storagePowerValues.length * storageDurationValues.length * electrolyzerValues.length;

            self.postMessage({ type: 'log', msg: `共有 ${totalCombinations} 种参数组合需要计算` });
            self.postMessage({ type: 'log', msg: `成功读取输入数据，共 ${pvData.length} 小时` });

            const allResults = [];
            let completed = 0;

            for (const sp of storagePowerValues) {
                for (const sd of storageDurationValues) {
                    for (const ec of electrolyzerValues) {
                        completed++;
                        const progress = (completed / totalCombinations) * 100;
                        self.postMessage({
                            type: 'progress',
                            progress,
                            msg: `正在计算第 ${completed}/${totalCombinations} 种组合 (储能${sp}MW×${sd}h, 电解槽${ec}MW)...`
                        });

                        const result = runSingleSimulation(pvData, windData, params, sp, sd, ec);
                        allResults.push(result);
                    }
                }
            }

            self.postMessage({ type: 'complete', results: allResults });
        }
    };
}

// 导出到全局（浏览器主线程 window / Web Worker self 通用）
self.runSingleSimulation = runSingleSimulation;
