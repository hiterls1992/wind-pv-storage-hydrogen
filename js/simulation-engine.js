/**
 * ============================================================================
 * 仿真计算引擎  ——  V2.2 参数体系重构版
 * ============================================================================
 * 从 simulation_logic.py 1:1 翻译为纯 JavaScript。
 *
 * ★★★ 本次重构只改「参数怎么传进来」，计算逻辑一行未动 ★★★
 *   V2.1：runSingleSimulation(pv, wind, params, storagePower, storageDuration, electrolyzerCapacity)
 *        —— 容量参数一部分藏在 params.PV_CAPACITY / WIND_CAPACITY 里，
 *           另一部分作为裸数字参数传入，参数来源不统一（任务书 §11）。
 *   V2.2：runSingleSimulation(pv, wind, scheme, simulationConfig)
 *        —— 容量全部来自 scheme，运行规则全部来自 simulationConfig，彻底分离（§13 / §36）。
 *
 * ---------------------------------------------------------------------------
 * 统一入口签名（任务书 §11 / §35）
 * ---------------------------------------------------------------------------
 *   runSingleSimulation(
 *       pvData,            // Float64Array 光伏「单 MW 每小时电量」MWh/MW
 *       windData,          // Float64Array 风电「单 MW 每小时电量」MWh/MW
 *       scheme,            // { windCapacity, pvCapacity, storagePower, storageDuration, electrolyzerCapacity }
 *       simulationConfig   // { electrolyzerMinRatio, maxExportHourly, maxExportTotal, maxImportRatio,
 *                          //   chargeEfficiency, dischargeEfficiency, hydrogenConsumption }
 *   )
 *
 * ---------------------------------------------------------------------------
 * 单位约定（任务书 §46）—— 功率与电量不得混淆
 * ---------------------------------------------------------------------------
 *   输入 pvData / windData 为**单位容量（1 MW）的每小时电量** MWh/MW，
 *   故「每小时电量 MWh = 单位容量电量 MWh/MW × 装机容量 MW」。
 *   这一约定来自 V1.0 的 input.xlsx（表头为「光伏电量(MWh/MW)」「风电电量(MWh/MW)」），
 *   V2.2 继续保持，且必须保持——改动会让全部历史结果失效。
 *
 *   容量 MW / 时长 h / 容量 MWh / 电量 MWh / 制氢量 kg / 制氢电耗 kWh/kg
 */

/**
 * 参数解析：兼容 V2.1 旧签名，统一产出 { scheme, simulationConfig }。
 *
 * TODO V2.3 REMOVE LEGACY
 *   兼容层仅为「不破坏已有调用方」而存在；业务逻辑必须逐步迁移到
 *   runSingleSimulation(..., scheme, simulationConfig) 这一唯一形式（§43）。
 */
function resolveSimulationArgs(schemeOrParams, simulationConfig, legacyDuration, legacyElectrolyzer) {
    // 判据：第 4 个形参（simulationConfig 位）为数字 → 说明调用方用的是 V2.1 旧签名
    if (typeof simulationConfig === 'number') {
        const p = schemeOrParams || {};
        return {
            scheme: {
                windCapacity: p.WIND_CAPACITY,
                pvCapacity: p.PV_CAPACITY,
                storagePower: simulationConfig,
                storageDuration: legacyDuration,
                electrolyzerCapacity: legacyElectrolyzer,
            },
            simulationConfig: {
                electrolyzerMinRatio: p.ELECTROLYZER_MIN_RATIO,
                maxExportHourly: p.MAX_EXPORT_RATIO_HOURLY,
                maxExportTotal: p.MAX_EXPORT_RATIO_TOTAL,
                maxImportRatio: p.MAX_IMPORT_RATIO,
                chargeEfficiency: p.STORAGE_CHARGE_EFFICIENCY,
                dischargeEfficiency: p.STORAGE_DISCHARGE_EFFICIENCY,
                hydrogenConsumption: p.HYDROGEN_ENERGY_CONSUMPTION,
            },
        };
    }
    return {
        scheme: schemeOrParams || {},
        simulationConfig: simulationConfig || {},
    };
}

/**
 * 执行单个方案（一个容量组合）的 8760 小时仿真计算。
 *
 * @param {Float64Array} pvData   光伏单 MW 每小时电量 (MWh/MW × hours)
 * @param {Float64Array} windData 风电单 MW 每小时电量 (MWh/MW × hours)
 * @param {Object} scheme         方案参数（唯一容量来源）
 * @param {Object} simulationConfig 系统运行参数（不含任何容量）
 * @returns {Object} 仿真结果 { results, ratioData, systemVars, sums, scheme, simulationConfig, filename }
 */
function runSingleSimulation(pvData, windData, scheme, simulationConfig, legacyDuration, legacyElectrolyzer) {
    const parsed = resolveSimulationArgs(scheme, simulationConfig, legacyDuration, legacyElectrolyzer);

    // ---- 类别 ①：方案参数（建多大）----
    const {
        windCapacity,
        pvCapacity,
        storagePower,
        storageDuration,
        electrolyzerCapacity,
    } = parsed.scheme;

    // ---- 类别 ②：系统运行参数（设备怎么运行）----
    const {
        electrolyzerMinRatio,
        maxExportHourly,
        maxExportTotal,
        maxImportRatio,
        chargeEfficiency,
        dischargeEfficiency,
        hydrogenConsumption,
    } = parsed.simulationConfig;

    // ---- 派生量 ----
    const totalCapacity = pvCapacity + windCapacity;                   // MW
    const electrolyzerMin = electrolyzerCapacity * electrolyzerMinRatio; // MW
    const maxExportHourlyPower = totalCapacity * maxExportHourly;        // MW
    const storageEnergy = storagePower * storageDuration;               // MWh（派生，禁止独立输入）

    if (typeof pvData !== 'object' || pvData === null) throw new Error('runSingleSimulation: 缺少光伏 8760 小时数据');
    if (typeof windData !== 'object' || windData === null) throw new Error('runSingleSimulation: 缺少风电 8760 小时数据');
    if (pvData.length !== windData.length) throw new Error('runSingleSimulation: 光伏与风电数据长度不一致');

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
        const pv = pvData[h] * pvCapacity;
        const wind = windData[h] * windCapacity;
        const total = pv + wind;
        totalGenerated += total;

        const idx = h * COLS;
        results[idx + 0] = pv;     // 光伏电量
        results[idx + 1] = wind;   // 风电电量
        results[idx + 2] = total;  // 合计电量

        let charge = 0, discharge = 0, hydrogenPower = 0, export_ = 0, importPower = 0, curtailment = 0, hydrogenProduction = 0;

        if (total >= electrolyzerMin) {
            // 情况1：风光电量足够维持最低制氢功率
            hydrogenPower = Math.min(total, electrolyzerCapacity);
            const remaining = total - hydrogenPower;

            // 储能充电
            const chargeCapacityLimit = (storageEnergy - storageRemaining) / chargeEfficiency;
            const chargePowerLimit = storagePower;
            const chargePossible = Math.min(chargeCapacityLimit, chargePowerLimit);
            charge = Math.min(remaining, chargePossible);
            storageRemaining += charge * chargeEfficiency;
            const remainingAfterCharge = remaining - charge;

            // 上网电量
            const exportHourlyLimit = maxExportHourlyPower;
            const exportTotalLimit = totalGenerated * maxExportTotal;
            const exportRemainingQuota = exportTotalLimit - totalExported;
            export_ = Math.min(remainingAfterCharge, exportHourlyLimit, exportRemainingQuota);
            curtailment = remainingAfterCharge - export_;

            totalExported += export_;

            hydrogenProduction = hydrogenPower * 1000 / hydrogenConsumption;
        } else {
            // 情况2：风光电量不足
            const needed = electrolyzerMin - total;

            // 储能放电
            const dischargeCapacityLimit = storageRemaining * dischargeEfficiency;
            const dischargePowerLimit = storagePower;
            const dischargePossible = Math.min(dischargeCapacityLimit, dischargePowerLimit);
            discharge = Math.min(needed, dischargePossible);
            storageRemaining -= discharge / dischargeEfficiency;

            const remainingNeeded = needed - discharge;

            // 下网电量
            importPower = Math.min(remainingNeeded, totalCapacity * maxImportRatio);
            hydrogenPower = total + discharge + importPower;

            hydrogenProduction = hydrogenPower * 1000 / hydrogenConsumption;
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
        '光伏容量（MW）': pvCapacity,
        '风电容量（MW）': windCapacity,
        '储能功率（MW）': storagePower,
        '储能时长（小时）': storageDuration,
        '电解槽容量（MW）': electrolyzerCapacity,
        '电解槽最低运行比例': electrolyzerMinRatio,
        '原每小时上网比例上限': maxExportHourly,
        '新增总量上网比例上限': maxExportTotal,
        '下网电量比例上限': maxImportRatio,
        '储能充电效率': chargeEfficiency,
        '储能放电效率': dischargeEfficiency,
        '每千克氢气耗电量（kWh/kg）': hydrogenConsumption,
        '总容量（MW）': totalCapacity,
        '电解槽最小运行功率（MW）': electrolyzerMin,
        '原每小时上网电量绝对上限（MW）': maxExportHourlyPower,
        '储能总容量（MWh）': storageEnergy,
    };

    return {
        results,       // Float64Array (hours * 11)
        ratioData,
        systemVars,
        sums: { sumTotal, sumHydrogenPower, sumExport, sumImport, sumCurtailment, sumH2Prod },
        // 结果必须自带完整方案（任务书 §48）：不依赖文件名 / DOM / 全局变量判断结果对应哪个方案
        scheme: {
            windCapacity: windCapacity,
            pvCapacity: pvCapacity,
            storagePower: storagePower,
            storageDuration: storageDuration,
            storageEnergy: storageEnergy,
            electrolyzerCapacity: electrolyzerCapacity,
        },
        // 结果同时记录运行参数，便于审计「这个结果是在什么运行规则下算出来的」
        simulationConfig: {
            electrolyzerMinRatio: electrolyzerMinRatio,
            maxExportHourly: maxExportHourly,
            maxExportTotal: maxExportTotal,
            maxImportRatio: maxImportRatio,
            chargeEfficiency: chargeEfficiency,
            dischargeEfficiency: dischargeEfficiency,
            hydrogenConsumption: hydrogenConsumption,
        },
        filename: `OUTPUT-${electrolyzerCapacity}MW-${storagePower}MW-${storageDuration}H-${maxExportTotal.toFixed(1)}.xlsx`
    };
}

// 导出到全局（浏览器主线程 window / Web Worker self 通用）
self.runSingleSimulation = runSingleSimulation;
