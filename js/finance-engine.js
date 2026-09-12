/**
 * 财务评价计算引擎
 * 从 finance_module.py FinanceCalculator 类翻译
 */

const FinanceEngine = {
    /**
     * 计算所有财务指标
     * @param {Object} params - 用户输入参数
     * @param {Object} dataSummary - 数据汇总行
     * @param {Object} estimateData - 概算结果行
     * @returns {{result: Object, detail: Object}}
     */
    calculateAll(params, dataSummary, estimateData) {
        const p = params;
        const d = dataSummary;
        const e = estimateData;

        // 预计算派生参数
        const electrolyzerCapacityMw = Utils.toNum(d['电解槽容量（MW）']);
        const singleElectrolyzerMw = p.single_electrolyzer_mw;
        const electrolyzerCount = Math.max(1, Math.floor(electrolyzerCapacityMw / singleElectrolyzerMw));

        const pvCapacity = Utils.toNum(d['光伏容量（MW）']);
        const windCapacity = Utils.toNum(d['风电容量（MW）']);

        const totalH2Power = Utils.toNum(d['制氢电量总和（亿度）']);       // 亿度/年
        const totalExportPower = Utils.toNum(d['上网电量总和（亿度）']);
        const totalImportPower = Utils.toNum(d['下网电量总和（亿度）']);
        const totalH2Prod = Utils.toNum(d['制氢量总和（万吨）']);           // 万吨/年

        const calcPeriod = p.calc_period;
        const constructPeriod = p.construct_period;
        const operatePeriod = calcPeriod - constructPeriod;

        // 投资数据
        const constructionInvestment = Utils.toNum(e['建设总投资（万元）']);

        // 流动资金
        const workingCapital = (30 * pvCapacity / 1.2 + 40 * windCapacity) * 1000 / 10000;

        // 建设期利息
        const capitalRatio = p.capital_ratio / 100.0;
        const loanRatio = 1.0 - capitalRatio;
        const loanRate = p.long_term_loan_rate / 100.0;
        const loanDuringConstruction = constructionInvestment * loanRatio;
        const constructionInterest = loanDuringConstruction * loanRate * constructPeriod / 2.0;

        // 项目总投资
        const totalInvestment = constructionInvestment + constructionInterest + workingCapital;

        // 资本金和贷款（资本金基数 = 建设投资；建设期利息、流动资金另由融资解决）
        const capitalFund = constructionInvestment * capitalRatio;
        const loanPrincipal = totalInvestment - capitalFund;

        // 折旧
        const fixedAssetOriginal = constructionInvestment + constructionInterest;
        const residualValue = fixedAssetOriginal * (p.residual_rate / 100.0);
        const annualDepreciation = (fixedAssetOriginal - residualValue) / p.depreciation_years;

        // 还款计划（等额本息）
        const loanPeriod = p.loan_period;
        let annualLoanPayment;
        if (loanPeriod > 0 && loanRate > 0) {
            annualLoanPayment = loanPrincipal * loanRate * Math.pow(1 + loanRate, loanPeriod) / (Math.pow(1 + loanRate, loanPeriod) - 1);
        } else {
            annualLoanPayment = loanPrincipal / Math.max(loanPeriod, 1);
        }

        // === 年度成本费用 ===
        // 容量电费
        const capacityDemandKva = electrolyzerCapacityMw * 0.3 * 1000;
        const annualCapacityElecFee = p.capacity_elec_fee * capacityDemandKva / 10000 * 12;

        // 上网电费（收入）
        const annualOngridFee = totalExportPower * p.ongrid_price * 100000000 / 10000;

        // 下网电费
        const annualOffgridFee = totalImportPower * p.offgrid_price * 100000000 / 10000;

        // 过网费
        const wheelingPower = totalH2Power - totalImportPower;
        const annualWheelingFee = p.wheeling_fee * wheelingPower * 10000;

        // 人工
        const annualH2Labor = p.h2_labor_cost * electrolyzerCount;

        // 维护
        const annualElectrolyzerMaint = p.electrolyzer_maint * electrolyzerCount;

        // 大修（第10年）
        const electrolyzerOverhaulTotal = p.electrolyzer_overhaul * electrolyzerCount;

        // 水费
        const waterAmount = totalH2Prod * 10000000 / 0.0899 * 0.006;
        const annualWaterFee = p.water_unit_price * waterAmount / 10000;

        // 运维
        const annualWindPvOm = p.wind_pv_om * (windCapacity + pvCapacity / 1.2) * 1000 / 10000;

        // === 辅助函数 ===
        function getAnnualRevenue() {
            return totalH2Prod * 10000000 * p.h2_price / 10000;
        }

        function getAnnualRevenueExclVat() {
            return getAnnualRevenue() / (1 + p.output_vat_rate / 100.0);
        }

        function getAnnualOperatingCost() {
            return annualCapacityElecFee + annualOffgridFee + annualWheelingFee +
                   annualH2Labor + annualElectrolyzerMaint + annualWaterFee +
                   annualWindPvOm + workingCapital;
        }

        function getAnnualTotalCost(year, includeDepreciation = true) {
            let cost = getAnnualOperatingCost();
            if (includeDepreciation && year >= constructPeriod) {
                if (year - constructPeriod < p.depreciation_years) {
                    cost += annualDepreciation;
                }
            }
            if (year === constructPeriod + 10) {
                cost += electrolyzerOverhaulTotal;
            }
            return cost;
        }

        const overhaulYear = constructPeriod + 10;

        // === 项目投资现金流量（融资前） ===
        const projectCF = [];
        const annualInvestment = constructionInvestment / Math.max(constructPeriod, 1);

        for (let year = 0; year < calcPeriod; year++) {
            const cf = { year };

            if (year < constructPeriod) {
                cf.revenue = 0;
                cf.operating_cost = 0;
                cf.investment = annualInvestment;
                cf.working_capital = 0;
                cf.income_tax = 0;
                cf.net_cf_pre_tax = -annualInvestment;
                cf.net_cf_post_tax = -annualInvestment;
            } else {
                const operateYear = year - constructPeriod;
                cf.revenue = getAnnualRevenueExclVat();

                let opCost = getAnnualOperatingCost();
                if (year === overhaulYear) opCost += electrolyzerOverhaulTotal;
                cf.operating_cost = opCost;

                cf.investment = 0;
                cf.working_capital = (year === constructPeriod) ? workingCapital : 0;

                const depreciation = (operateYear < p.depreciation_years) ? annualDepreciation : 0;

                const profitPreTax = cf.revenue - opCost - depreciation;

                // 增值税
                const outputVat = getAnnualRevenue() - cf.revenue;
                const inputVat = opCost * (p.output_vat_rate / 100.0) * 0.5;
                const vat = Math.max(0, outputVat - inputVat);
                const surtax = vat * (p.surtax_rate / 100.0);

                const taxableIncome = profitPreTax - surtax;
                cf.income_tax = Math.max(0, taxableIncome * (p.income_tax_rate / 100.0));

                cf.net_cf_pre_tax = cf.revenue - opCost - cf.investment - cf.working_capital - surtax;
                cf.net_cf_post_tax = cf.net_cf_pre_tax - cf.income_tax;

                if (year === calcPeriod - 1) {
                    cf.net_cf_pre_tax += residualValue + workingCapital;
                    cf.net_cf_post_tax += residualValue + workingCapital;
                }
            }

            projectCF.push(cf);
        }

        // === 资本金现金流量（融资后） ===
        const equityCF = [];
        let outstandingLoan = loanPrincipal;

        for (let year = 0; year < calcPeriod; year++) {
            const cf = { year };

            if (year < constructPeriod) {
                cf.capital_investment = annualInvestment * capitalRatio;
                cf.loan_received = annualInvestment * (1 - capitalRatio);
                cf.revenue = 0;
                cf.operating_cost = 0;
                cf.loan_repay_principal = 0;
                cf.loan_repay_interest = 0;
                cf.income_tax = 0;
                cf.net_cf = -cf.capital_investment;
            } else {
                const operateYear = year - constructPeriod;

                cf.capital_investment = (year === constructPeriod) ? workingCapital * capitalRatio : 0;
                cf.loan_received = (year === constructPeriod) ? workingCapital * (1 - capitalRatio) : 0;

                cf.revenue = getAnnualRevenueExclVat();

                let opCost = getAnnualOperatingCost();
                if (year === overhaulYear) opCost += electrolyzerOverhaulTotal;
                cf.operating_cost = opCost;

                const depreciation = (operateYear < p.depreciation_years) ? annualDepreciation : 0;

                // 利息和本金
                let interest = 0, principal = 0;
                if (operateYear < p.loan_period && outstandingLoan > 0) {
                    interest = outstandingLoan * loanRate;
                    principal = annualLoanPayment - interest;
                    if (principal > outstandingLoan) principal = outstandingLoan;
                    principal = Math.max(0, principal);
                }

                cf.loan_repay_principal = principal;
                cf.loan_repay_interest = interest;

                // 增值税及附加
                const outputVat = getAnnualRevenue() - cf.revenue;
                const inputVat = opCost * (p.output_vat_rate / 100.0) * 0.5;
                const vat = Math.max(0, outputVat - inputVat);
                const surtax = vat * (p.surtax_rate / 100.0);

                const profitPreTax = cf.revenue - opCost - depreciation - interest - surtax;
                cf.income_tax = Math.max(0, profitPreTax * (p.income_tax_rate / 100.0));

                cf.net_cf = cf.revenue - opCost - principal - interest - surtax -
                            cf.income_tax - cf.capital_investment + cf.loan_received;

                if (year === calcPeriod - 1) {
                    cf.net_cf += residualValue + workingCapital;
                }

                outstandingLoan -= principal;
            }

            equityCF.push(cf);
        }

        // === 计算财务指标 ===
        const preTaxCFs = projectCF.map(cf => cf.net_cf_pre_tax);
        const postTaxCFs = projectCF.map(cf => cf.net_cf_post_tax);
        const equityCFs = equityCF.map(cf => cf.net_cf);

        const capitalBenchmark = p.capital_benchmark / 100.0;
        const industryPre = p.industry_benchmark_pre / 100.0;
        const industryPost = p.industry_benchmark_post / 100.0;

        const npvPre = _calcNPV(preTaxCFs, industryPre);
        const npvPost = _calcNPV(postTaxCFs, industryPost);
        const npvCapital = _calcNPV(equityCFs, capitalBenchmark);

        const irrPre = _calcIRR(preTaxCFs);
        const irrPost = _calcIRR(postTaxCFs);
        const irrCapital = _calcIRR(equityCFs);

        const paybackPre = _calcPayback(preTaxCFs);
        const paybackPost = _calcPayback(postTaxCFs);

        // ROI
        const annualRevenueExclVat = getAnnualRevenueExclVat();
        let totalEbit = 0;
        for (let yr = constructPeriod; yr < calcPeriod; yr++) {
            const opYr = yr - constructPeriod;
            let opCost = getAnnualOperatingCost();
            if (yr === overhaulYear) opCost += electrolyzerOverhaulTotal;
            const dep = (opYr < p.depreciation_years) ? annualDepreciation : 0;
            totalEbit += annualRevenueExclVat - opCost - dep;
        }
        const avgEbit = totalEbit / operatePeriod;
        const roi = totalInvestment > 0 ? (avgEbit / totalInvestment) * 100 : 0;

        // 投资利税率
        let totalSurtax = 0;
        let totalProfitTax = 0;
        for (let yr = constructPeriod; yr < calcPeriod; yr++) {
            const opYr = yr - constructPeriod;
            let opCost = getAnnualOperatingCost();
            if (yr === overhaulYear) opCost += electrolyzerOverhaulTotal;
            const dep = (opYr < p.depreciation_years) ? annualDepreciation : 0;
            const outputVat = getAnnualRevenue() - annualRevenueExclVat;
            const inputVat = opCost * (p.output_vat_rate / 100.0) * 0.5;
            const vat = Math.max(0, outputVat - inputVat);
            const surtax = vat * (p.surtax_rate / 100.0);
            totalSurtax += surtax;
            const ebit = annualRevenueExclVat - opCost - dep;
            const tax = Math.max(0, (ebit - surtax) * (p.income_tax_rate / 100.0));
            totalProfitTax += vat + surtax + tax;
        }
        const totalRevenue = annualRevenueExclVat * operatePeriod;
        let totalCost2 = 0;
        let ol2 = loanPrincipal;
        for (let yr = constructPeriod; yr < calcPeriod; yr++) {
            const opYr = yr - constructPeriod;
            let ac = getAnnualTotalCost(yr, (opYr < p.depreciation_years));
            if (opYr < p.loan_period && ol2 > 0) {
                const interest2 = ol2 * loanRate;
                ac += interest2;
                const principal2 = annualLoanPayment - interest2;
                ol2 = Math.max(0, ol2 - Math.max(0, principal2));
            }
            totalCost2 += ac;
        }
        const totalProfit = totalRevenue - totalCost2;
        const investmentProfitTaxRate = totalInvestment > 0 ?
            ((totalProfit + totalProfitTax + totalSurtax) / operatePeriod / totalInvestment) * 100 : 0;

        // ROE
        let totalNetProfit = 0;
        let ol3 = loanPrincipal;
        for (let yr = constructPeriod; yr < calcPeriod; yr++) {
            const opYr = yr - constructPeriod;
            let opCost = getAnnualOperatingCost();
            if (yr === overhaulYear) opCost += electrolyzerOverhaulTotal;
            const dep = (opYr < p.depreciation_years) ? annualDepreciation : 0;
            let interest3 = 0;
            if (opYr < p.loan_period && ol3 > 0) {
                interest3 = ol3 * loanRate;
                const principal3 = annualLoanPayment - interest3;
                ol3 = Math.max(0, ol3 - Math.max(0, principal3));
            }
            const outputVat = getAnnualRevenue() - annualRevenueExclVat;
            const inputVat = opCost * (p.output_vat_rate / 100.0) * 0.5;
            const vat = Math.max(0, outputVat - inputVat);
            const surtax = vat * (p.surtax_rate / 100.0);
            const ebt = annualRevenueExclVat - opCost - dep - interest3 - surtax;
            const tax = Math.max(0, ebt * (p.income_tax_rate / 100.0));
            totalNetProfit += ebt - tax;
        }
        const avgNetProfit = totalNetProfit / operatePeriod;
        const roe = capitalFund > 0 ? (avgNetProfit / capitalFund) * 100 : 0;

        // 资产负债率
        const maxDebtRatio = totalInvestment > 0 ? (loanPrincipal / totalInvestment) * 100 : 0;

        // 净资产收益率
        const netAssetReturn = capitalFund > 0 ? (avgNetProfit / capitalFund) * 100 : 0;

        // 营业现金比率
        let avgOperatingCF = 0;
        let count = 0;
        for (let yr = constructPeriod; yr < calcPeriod; yr++) {
            if (yr !== calcPeriod - 1) {
                avgOperatingCF += projectCF[yr].net_cf_pre_tax;
                count++;
            }
        }
        avgOperatingCF = avgOperatingCF / Math.max(count, 1);
        const operatingCashRatio = annualRevenueExclVat > 0 ? (avgOperatingCF / annualRevenueExclVat) * 100 : 0;

        // 盈亏平衡点
        let totalFixedCost = 0;
        let totalVariableCost = 0;
        const avgInterest = loanPrincipal * loanRate * 0.5;
        for (let yr = constructPeriod; yr < calcPeriod; yr++) {
            const opYr = yr - constructPeriod;
            let fixed = ((opYr < p.depreciation_years) ? annualDepreciation : 0) +
                        annualH2Labor + annualElectrolyzerMaint + annualCapacityElecFee + annualWindPvOm;
            if (yr === overhaulYear) fixed += electrolyzerOverhaulTotal;
            fixed += avgInterest;
            totalFixedCost += fixed;
            totalVariableCost += annualOffgridFee + annualWheelingFee + annualWaterFee + workingCapital;
        }
        const avgFixed = totalFixedCost / operatePeriod;
        const avgVariable = totalVariableCost / operatePeriod;
        const bep = (annualRevenueExclVat - avgVariable > 0) ?
            (avgFixed / (annualRevenueExclVat - avgVariable)) * 100 : 100.0;

        // EVA
        const taxRate = p.income_tax_rate / 100.0;
        const nopat = avgEbit * (1 - taxRate);
        const wacc = capitalBenchmark * (p.capital_ratio / 100.0) +
                     loanRate * (1 - taxRate) * (1 - p.capital_ratio / 100.0);
        const eva = nopat - wacc * totalInvestment;

        // 制氢成本
        let totalCostForH2 = 0;
        for (let yr = constructPeriod; yr < calcPeriod; yr++) {
            totalCostForH2 += getAnnualTotalCost(yr, true);
        }
        const totalH2ProdKg = totalH2Prod * operatePeriod * 10000000;
        const h2CostPerKg = totalH2ProdKg > 0 ? totalCostForH2 / totalH2ProdKg * 10000 : 0;

        // 结果
        const result = {
            '项目总投资（万元）': Math.round(totalInvestment * 100) / 100,
            '建设投资（万元）': Math.round(constructionInvestment * 100) / 100,
            '建设期利息（万元）': Math.round(constructionInterest * 100) / 100,
            '流动资金（万元）': Math.round(workingCapital * 100) / 100,
            '销售收入总额（不含增值税）（万元）': Math.round(totalRevenue * 100) / 100,
            '总成本费用（万元）': Math.round(totalCost2 * 100) / 100,
            '销售税金附加总额（万元）': Math.round(totalSurtax * 100) / 100,
            '利润总额（万元）': Math.round(totalProfit * 100) / 100,
            '项目投资回收期（所得税前）（年）': Math.round(paybackPre * 100) / 100,
            '项目投资回收期（所得税后）（年）': Math.round(paybackPost * 100) / 100,
            '项目投资财务内部收益率（所得税前）（%）': Math.round(irrPre * 10000) / 100,
            '项目投资财务内部收益率（所得税后）（%）': Math.round(irrPost * 10000) / 100,
            '项目投资财务净现值（所得税前）（万元）': Math.round(npvPre * 100) / 100,
            '项目投资财务净现值（所得税后）（万元）': Math.round(npvPost * 100) / 100,
            '资本金财务内部收益率（%）': Math.round(irrCapital * 10000) / 100,
            '资本金财务净现值（万元）': Math.round(npvCapital * 100) / 100,
            '总投资收益率（ROI）（%）': Math.round(roi * 100) / 100,
            '投资利税率（%）': Math.round(investmentProfitTaxRate * 100) / 100,
            '项目资本金净利润率（ROE）（%）': Math.round(roe * 100) / 100,
            '资产负债率（最大值）（%）': Math.round(maxDebtRatio * 100) / 100,
            '净资产收益率（%）': Math.round(netAssetReturn * 100) / 100,
            '营业现金比率（%）': Math.round(operatingCashRatio * 100) / 100,
            '盈亏平衡点（生产能力利用率）（%）': Math.round(bep * 100) / 100,
            '经济增加值（EVA）（万元）': Math.round(eva * 100) / 100,
            '制氢成本（元/kg）': Math.round(h2CostPerKg * 100) / 100,
        };

        const detail = {
            annual_revenue: Math.round(annualRevenueExclVat * 100) / 100,
            annual_total_cost: Math.round(totalCost2 / operatePeriod * 100) / 100,
            annual_depreciation: Math.round(annualDepreciation * 100) / 100,
            annual_electrolyzer_count: electrolyzerCount,
            project_cf: projectCF,
            equity_cf: equityCF,
            electrolyzer_overhaul_year: overhaulYear,
            electrolyzer_overhaul_total: Math.round(electrolyzerOverhaulTotal * 100) / 100,

            // ===== V2.1 新增：LCOH（折现口径）所需的年度成本组成 =====
            // 说明：以下字段为纯新增，不参与任何原有财务指标计算，
            //       不影响 result 中的既有指标，也不影响现金流量表导出结果。
            calc_period: calcPeriod,
            construct_period: constructPeriod,
            construction_investment: constructionInvestment,      // 万元（建设投资）
            construction_interest: constructionInterest,          // 万元（建设期利息）
            working_capital: workingCapital,                      // 万元（流动资金）
            annual_opex: getAnnualOperatingCost(),                // 万元/年（经营成本：不含折旧、利息、所得税）
            overhaul_total: electrolyzerOverhaulTotal,            // 万元（电解槽大修费）
            overhaul_year: overhaulYear,                          // 大修发生年份（计算期年份索引）
            annual_h2_kg: totalH2Prod * 1e7,                      // kg/年（年制氢量）
        };

        return { result, detail };
    }
};

// === 内部辅助函数 ===

function _calcNPV(cashflows, rate) {
    let npv = 0;
    for (let t = 0; t < cashflows.length; t++) {
        npv += cashflows[t] / Math.pow(1 + rate, t);
    }
    return npv;
}

function _calcIRR(cashflows, maxIter = 1000, tol = 1e-6) {
    if (cashflows.every(cf => cf >= 0) || cashflows.every(cf => cf <= 0)) return 0;

    let lo = -0.99, hi = 2.0;
    let fLo = _npvAtRate(cashflows, lo);
    let fHi = _npvAtRate(cashflows, hi);

    if (fLo * fHi > 0) {
        for (hi of [5.0, 10.0, 50.0]) {
            fHi = _npvAtRate(cashflows, hi);
            if (fLo * fHi <= 0) break;
        }
    }

    if (fLo * fHi > 0) return 0;

    for (let i = 0; i < maxIter; i++) {
        const mid = (lo + hi) / 2;
        const fMid = _npvAtRate(cashflows, mid);
        if (Math.abs(fMid) < tol) return mid;
        if (fLo * fMid < 0) { hi = mid; fHi = fMid; }
        else { lo = mid; fLo = fMid; }
    }
    return (lo + hi) / 2;
}

function _npvAtRate(cashflows, rate) {
    let npv = 0;
    for (let t = 0; t < cashflows.length; t++) {
        npv += cashflows[t] / Math.pow(1 + rate, t);
    }
    return npv;
}

function _calcPayback(cashflows) {
    let cumulative = 0;
    for (let t = 0; t < cashflows.length; t++) {
        cumulative += cashflows[t];
        if (cumulative >= 0) {
            if (t > 0) {
                const prevCum = cumulative - cashflows[t];
                const fraction = cashflows[t] !== 0 ? Math.abs(prevCum) / cashflows[t] : 0;
                return t - 1 + fraction;
            }
            return t;
        }
    }
    return cashflows.length;
}

// 导出（浏览器主线程 window / Web Worker self 通用）
self.FinanceEngine = FinanceEngine;
