'use strict';// ═══════════════════════════════════════════════════════════════// DEFAULTS — sourced exactly from Annexure-A.I (20MT), Annexure-A.II// (30MT) and the CE_35_MT_Brand_New_TTU Excel workbook// (ANNEXURES, Capex_Deprecn, Wages_OT_PF_ESI, Interest_Maint_Insurance)// ═══════════════════════════════════════════════════════════════const n = (v,d=0)=>{ const p=parseFloat(v); return isNaN(p)?d; };

const DEFAULTS_20MT = Object.freeze({projectName:'New TTU Project', ttuType:'20MT',contractPeriod:4, workingDays:26, kmPerMonth:3000, contractStartYear:2026,

// ── Capital (Annexure-A.I rows 1.1/1.2 — ANNEXURES!F11/F12, ROUND(avg-of-3-quotes,0)) ──tractorCost:2430159, trailerCost:852667,loanPct:85, loanInterestRate:8.95, ownCapRate:7.5,usefulLife:8, salvagePct:5,// Capex_Deprecn: the capital cost is dated 2021/22; 4 years of SLM depreciation (on the// ORIGINAL/gross cost) are stripped out BEFORE the 2026 contract starts to arrive at the// "Net Capital Cost" used for the loan/EMI/maintenance-% base (ANNEXURES!F16). The same// annual depreciation continues during the 4 contract years to derive the salvage value// at contract end (Capex_Deprecn!F28). Tied to contractPeriod (both workbooks: 4 yrs// prior + 4 yrs during, matching D10/12=4).

// ── Wages — Operator (Highly Skilled) — Wages_OT_PF_ESI row 14 base ──opBasicBase:579, opVDABase:339, opEscalationPct:3.17, opSpecialAllowance:0,opBonusPct:0, // Excel: operator bonus = 0 for 2026-29 circularbonusEscalationPct:4.85,

// ── Wages — Helper (Semi-Skilled) — row 26 base — × 2 positions ──hlpBasicBase:410, hlpVDABase:240, hlpEscalationPct:3.17, hlpSpecialAllowance:0,hlpBonusPct:8.33,

// ── Wages — Supervisor (Skilled) — row 39 base — shared across 4 TTUs ──supBasicBase:494, supVDABase:287, supEscalationPct:3.17, supSpecialAllowance:0,supBonusPct:8.33, supervisionTTUShare:4,

// ── Reliever / OT structure ──relieverDaysFactor:84, // Wages_OT_PF_ESI P8singleOTHrs:1, doubleOTHrs:3,numOperators:1, numHelpers:2,

// ── PF / ESI / EC Insurance ──pfRatePct:13, pfDailyCap:75, esiRatePct:3.25, esiThreshold:21000, ecInsuranceAnnual:2738, employeePfRatePct:12,

// ── Uniform, Shoe & Medical (Annexure-A row 39, Annexure-III row 5.1) ──uniformShoePerHeadPerYear:6000, numUniformPersons:3,medicalCertPerHeadPerYear:3000, numMedicalPersons:3, medicalCertYearMultiplier:2, // Excel literal: 300032

// ── Toll (reimbursable, Annexure-C) ──tollPerTripAmount:1070, tripsPerYear:12,

// ── Fuel (Annexure-A rows 19-21) ──dieselPrice:89.77, defPriceLow:74.15, defPriceHigh:135,fuelEfficiency:2.5, defConsumptionPctOfFuel:3, lubeOilPct:0.5,

// ── Tyre (Annexure-A/D row 17) ──tyreCostExclGST:16839.357142857143, numTyres:14, tyreLife:30000, tyreSpareAdj:1,

// ── Battery (Annexure-A row 16) ──batteryUnitPrice:16500, batteryGSTFactor:1.18, batteryLife:2, batteryEscalationPct:1.34,

// ── Maintenance (% of Capital Cost, ANNEXURES!F57 — all 4 yrs = 4%) ──maintPct_yr1:4.0, maintPct_yr2:4.0, maintPct_yr3:4.0, maintPct_yr4:4.0,

// ── Insurance (Interest_Maint_Insurance sheet) ──// NOTE (audited against actual formulas, not the generic tariff write-up): the workbook// does NOT progress IDV-depreciation% or OD-rate% by vehicle age year-on-year. Every// contract year references the SAME fixed cells: depreciation = E57 (the ">4–5yr" tariff// bracket, 50%) and OD rate = O73 ("upto 7 yrs" column, 1.22%) — consistent with the fleet// already carrying ~4 yrs of notional prior use per the Capex_Deprecn logic above. The// age-graded schedule below IS still used for the separate Vintage-TTU dashboard feature// (which has no Excel equivalent), but the standard/new-TTU calculation uses the fixed// 50% / "upto 7 yr" values every year, matching G91/G123/G153/G183 exactly.gvw:40000, odRateUpto5yr:1.19, odRateUpto7yr:1.22, odRateAbove7yr:1.25,insFixedDepPct:50, // Interest_Maint_Insurance!E57 — used flat for all 4 yrs in the main calcliabilityTractor:785, liabilityTrailer:375,addlCoverRatePerKg:0.27, addlCoverGVWThreshold:12000,specialDiscountPct:35, tpiPremium:46435, paCoverPerPerson:100,// NCB progression actually hardcoded in the workbook (E76/E108/E138/E168), which differs// from the generic 0/20/25/35 tariff table: Yr1=20%, Yr2=20%, Yr3=25%, Yr4=25%.ncbPctYr1:20, ncbPctYr2:20, ncbPctYr3:25, ncbPctYr4:25,

// ── DEF cost/km literal constant (ANNEXURES!F52 hardcodes 104.58, distinct from the// F55 quote-average of 104.575 used only for display) ──defCostConstant:104.58,

// ── Road Tax, Permit & Pollution (ANNEXURES!F26 — single flat figure per annum, same// every contract year; not split by year or by sub-component in the workbook) ──roadTaxOnly_yr1:29980, roadTaxOnly_yrRest:29980, permitFee:0, pollutionCert:0,

// ── GST ──gstPct:18,

// ── Vintage TTU defaults ──vintagePurchaseYear:2022, vintageCurrentYear:2026, vintageSalvagePct:5,vintageUsefulLife:8, vintageContractPeriod:4,});

const DEFAULTS_30MT = Object.freeze({...DEFAULTS_20MT,ttuType:'30MT',tractorCost:3065905, trailerCost:1009000, // ANNEXURES!F11/F12, ROUND(avg-of-3-quotes,0)fuelEfficiency:2.0,tyreCostExclGST:16754.890909090907, numTyres:22,roadTaxOnly_yr1:37536, roadTaxOnly_yrRest:37536, permitFee:0, pollutionCert:0, // ANNEXURES!F26 flat figuregvw:40000,});

// ═══════════════════════════════════════════════════════════════// FINANCIAL PRIMITIVES// ═══════════════════════════════════════════════════════════════function PMT(rate, nper, pv, fv=0, type=0){if(!nper) return 0;if(rate===0) return -(pv+fv)/nper;const pvif = Math.pow(1+rate, nper);let pmt = rate/(pvif-1) * -(pvpvif+fv);if(type===1) pmt/=(1+rate);return pmt;}const round2 = v=>Math.round(v100)/100;

// ═══════════════════════════════════════════════════════════════// WAGE ENGINE — replicates Wages_OT_PF_ESI sheet exactly, incl.// the intentional Excel quirks (bonus escalation exponent 0.755,// the "+2×doubleOT" term appearing from year 2 for the operator// row but from year 1 for the helper row, PF capped at ₹75/day).// ═══════════════════════════════════════════════════════════════function computeWageRole(cfg, years){const rows=[];let basic = cfg.basicBaseMath.pow(1+cfg.escPct/100, 0.25);let vda   = cfg.vdaBase  Math.pow(1+cfg.escPct/100, 0.25);let bonus = (basic+vda)(cfg.bonusPct/100);for(let y=0;y<years;y++){if(y>0){basic = (1+cfg.escPct/100);vda   = (1+cfg.escPct/100);if(cfg.bonusFreshEachYear){bonus = (basic+vda)(cfg.bonusPct/100);} else if(y===1){bonus = bonusMath.pow(1+cfg.bonusEscPct/100, 0.755);} else {bonus = bonus(1+cfg.bonusEscPct/100);}}const sa = cfg.sa;const dailyRate = basic+vda+bonus+sa;const monthlyRate = dailyRatecfg.workingDays;const singleOT = basic/8;const doubleOT = (basic+vda)/82;const holidayOT = (basic+vda)/8;const V = basic+vda; // PF/ESI basis (S+T columns)const pfDaily = Math.min(round2(Vcfg.pfRatePct/100), cfg.pfCap);let esiOrEcDaily, mode;if(monthlyRate < cfg.esiThreshold){ esiOrEcDaily = round2(Vcfg.esiRatePct/100); mode='ESI'; }else { esiOrEcDaily = cfg.ecAnnual/365; mode='EC'; }const monthlyPF = pfDaily365/12;const monthlyESIorEC = esiOrEcDaily365/12;// Split so the UI can show "ESI" and "Insurance" as two distinct lines — only one is ever non-zero.const monthlyESI = mode==='ESI' ? monthlyESIorEC : 0;const monthlyInsurance = mode==='EC' ? monthlyESIorEC : 0;const dailyInclPFESI = dailyRate+pfDaily+esiOrEcDaily;const monthlyInclPFESI = dailyInclPFESI365/12;const relieverMonthly = dailyRatecfg.relieverFactor/12;const otAnnual = ((1singleOT)+(cfg.doubleOTHrsdoubleOT)+((y>0||cfg.extraDoubleFromY0)?2doubleOT:0))365cfg.headcount;// Employee's own PF share — informational only (standard 12% deducted from the// employee's own wage under EPF rules; NOT an addl. cost to the contractor, since// the employer-side column above already carries the statutory outgo billed to the client).const employeePfDaily = round2(V(cfg.employeePfRatePct??12)/100);const employeePfMonthly = employeePfDaily*365/12;rows.push({basic,vda,bonus,sa,dailyRate,monthlyRate,singleOT,doubleOT,holidayOT,V,pfDaily,esiOrEcDaily,mode,monthlyPF,monthlyESIorEC,monthlyESI,monthlyInsurance,dailyInclPFESI,monthlyInclPFESI,relieverMonthly,otAnnual,employeePfDaily,employeePfMonthly});}return rows;}

function computeSupervisorRole(cfg, years){const rows=[];let basic = cfg.basicBaseMath.pow(1+cfg.escPct/100, 0.25);let vda   = cfg.vdaBase  Math.pow(1+cfg.escPct/100, 0.25);for(let y=0;y<years;y++){if(y>0){ basic=(1+cfg.escPct/100); vda=(1+cfg.escPct/100); }const sa=cfg.sa;const basicPlusVdaSa = basic+vda+sa;const bonus = (basic+vda)(cfg.bonusPct/100); // recomputed fresh every yearconst dailyRate = basic+vda+bonus+sa;const V = basic+vda;const pfDaily = Math.min(round2(dailyRatecfg.pfRatePct/100), cfg.pfCap);const esiDaily = round2(basicPlusVdaSacfg.esiRatePct/100);const dailyInclPFESI = dailyRate+pfDaily+esiDaily;const monthlyShare = dailyRate(365/12)/cfg.ttuShare;const monthlyPF = pfDaily365/12, monthlyESI = esiDaily365/12;const employeePfDaily = round2(V*(cfg.employeePfRatePct??12)/100);const employeePfMonthly = employeePfDaily365/12;const singleOT = basic/8, doubleOT=(basic+vda)/82, holidayOT=(basic+vda)/8;rows.push({basic,vda,bonus,sa,dailyRate,V,pfDaily,esiDaily,dailyInclPFESI,monthlyShare,monthlyPF,monthlyESI,employeePfDaily,employeePfMonthly,singleOT,doubleOT,holidayOT});}return rows;}

// ═══════════════════════════════════════════════════════════════// INSURANCE ENGINE — one computeInsuranceYear() powers both://  1) the standard/new-TTU calc (fixed=true): replicates the workbook's ACTUAL formulas,//     which reference fixed cells (50% depreciation, 1.22% "upto 7yr" OD rate) every//     contract year rather than progressing by age — audited from Interest_Maint_Insurance//     rows 63-183 (C67/C99/C129/C159 all = E57; C72/C104/C134/C164 all = O73).//  2) the Vintage TTU dashboard (fixed=false): age-graded tariff schedule below//     (Interest_Maint_Insurance rows 50-57), keyed to exact calendar-date age — this//     feature has no Excel equivalent, so the general tariff table is used instead.// ═══════════════════════════════════════════════════════════════// Exact calendar-date age (Years/Months/Days) + fractional-year equivalent for// depreciation/insurance/maintenance calculations — replaces naive year subtraction.// e.g. 15-06-2020 → 03-07-2026 = 6 Years 0 Months 18 Days.function calendarAgeBreakdown(startDateStr, endDateStr){let s=new Date(startDateStr+'T00:00:00'), e=new Date(endDateStr+'T00:00:00');if(isNaN(s.getTime())||isNaN(e.getTime())) return {years:0,months:0,days:0,totalDays:0,fractionalYears:0};let swapped=false;if(e<s){ [s,e]=[e,s]; swapped=true; }let years=e.getFullYear()-s.getFullYear();let months=e.getMonth()-s.getMonth();let days=e.getDate()-s.getDate();if(days<0){ months--; const prevMonthLastDay=new Date(e.getFullYear(),e.getMonth(),0).getDate(); days+=prevMonthLastDay; }if(months<0){ years--; months+=12; }const totalDays=Math.round((e-s)/86400000);const fractionalYears=totalDays/365.25;return {years,months,days,totalDays,fractionalYears?-fractionalYears};}function idvDepreciationPct(ageYears){if(ageYears<=0.5) return 5;if(ageYears<=1)   return 15;if(ageYears<=2)   return 20;if(ageYears<=3)   return 30;if(ageYears<=4)   return 40;if(ageYears<=5)   return 50;return 50; // beyond 5 yrs: by mutual agreement — capped for estimation}function odPremiumRatePct(ageYears, cfg){if(ageYears<=5) return cfg.odRateUpto5yr;if(ageYears<=7) return cfg.odRateUpto7yr;return cfg.odRateAbove7yr;}function ncbPctForContractYear(yearIndex, cfg){// Interest_Maint_Insurance!E76/E108/E138/E168 — hardcoded 20/20/25/25, not the generic 0/20/25/35 tariff scaleconst arr = [cfg.ncbPctYr1,cfg.ncbPctYr2,cfg.ncbPctYr3,cfg.ncbPctYr4];return arr[yearIndex] ?? arr[3];}

function computeInsuranceYear(tractorCost, trailerCost, ageAtYearStart, yearIndex, cfg, fixed=true){// fixed=true replicates the actual workbook (used for the standard/new-TTU calc): every// contract year references the SAME cells — depreciation = E57 (50%) and OD rate = O73// ("upto 7 yr", 1.22%) — rather than progressing by vehicle age. fixed=false uses the// age-graded tariff schedule instead, for the Vintage-TTU dashboard (no Excel equivalent).const depPct = fixed ? cfg.insFixedDepPct : idvDepreciationPct(ageAtYearStart);const tractorIDV = tractorCost*(1-depPct/100);const trailerIDV = trailerCost*(1-depPct/100);const odRatePct = fixed ? cfg.odRateUpto7yr : odPremiumRatePct(ageAtYearStart, cfg);const tractorOD = tractorIDV*(odRatePct/100);const trailerOD = trailerIDV*(odRatePct/100);const addlCover = Math.max(0, cfg.gvw-cfg.addlCoverGVWThreshold)cfg.addlCoverRatePerKg;const liability = cfg.liabilityTractor+cfg.liabilityTrailer;const subTotalOD = tractorOD+trailerOD+liability+addlCover;const ncbPct = ncbPctForContractYear(yearIndex, cfg);const afterNCB = subTotalOD - subTotalOD(ncbPct/100);const afterDiscount = afterNCB - subTotalOD*(cfg.specialDiscountPct/100);const netODP = afterDiscount;const tpi = cfg.tpiPremium;const pa = cfg.paCoverPerPerson*(cfg.numOperators+cfg.numHelpers);const totalExGST = netODP+tpi+pa;const totalIncGST = totalExGST*(1+cfg.gstPct/100);return {depPct,tractorIDV,trailerIDV,odRatePct,tractorOD,trailerOD,addlCover,liability,subTotalOD,ncbPct,netODP,tpi,pa,totalExGST,totalIncGST};}

// ═══════════════════════════════════════════════════════════════// ENGINE — one unified computeAll() powers both the "new TTU"// Engine.calc() and the Vintage TTU dashboard (Engine.calcVintage())// by parameterising the vehicle's starting age and capital base.// ═══════════════════════════════════════════════════════════════const Engine = {

computeAll(inp, accessories=[], startAge=0, capitalOverride=null){const O = {};const cYears = Math.max(1, Math.round(n(inp.contractPeriod,4)));const cMonths = cYears*12;const kmPM = n(inp.kmPerMonth,3000);const workingDays = n(inp.workingDays,26);const loanRate = n(inp.loanInterestRate,8.95)/100;const ownRate  = n(inp.ownCapRate,7.5)/100;// Interest_Maint_Insurance!E10/E11 — Excel amortises on the EFFECTIVE MONTHLY rate,// (1+annual)^(1/12)-1, not annual/12. Used in every PMT() call below.const loanMonthlyRate = Math.pow(1+loanRate,1/12)-1;const ownMonthlyRate  = Math.pow(1+ownRate,1/12)-1;

// ── Capital & Depreciation (Companies Act 2013, Sch-II SLM) ──
O.vehicleCost = n(inp.tractorCost)+n(inp.trailerCost);
O.accessoriesCost = accessories.reduce((s,a)=>s+(n(a.qty)*n(a.unitCost)*(1+n(a.gstPct)/100))+n(a.installCost),0);
O.usefulLife = n(inp.usefulLife,8);
O.salvagePct = n(inp.salvagePct,5);
O.depRateAnnual = (1-O.salvagePct/100)/O.usefulLife;
O.annualDep = O.vehicleCost*O.depRateAnnual; // Capex_Deprecn!F13 — SLM always on the ORIGINAL/gross cost, never the net

if(capitalOverride!=null){
  // Vintage TTU: caller already supplies the age-adjusted current value — no further prior-depreciation strip.
  O.netCapitalCost = capitalOverride;
  O.salvageValueAtContractEnd = Math.max(0, O.netCapitalCost - O.annualDep*cYears);
} else {
  // Capex_Deprecn: the quoted capital cost is dated 2021/22. `cYears` years of SLM
  // depreciation (on the gross vehicle cost) are stripped out BEFORE the contract
  // starts to get "Net Capital Cost" (ANNEXURES!F16) — the base for loan/EMI/
  // maintenance-%. The same annual amount continues to be charged through the
  // contract to find the salvage value at contract end (Capex_Deprecn!F28).
  // Accessories are NOT part of this chain — added to net capital cost undepreciated.
  const priorDepreciation = O.annualDep*cYears;
  const netVehicleCost = O.vehicleCost - priorDepreciation;
  O.netCapitalCost = netVehicleCost + O.accessoriesCost;
  O.salvageValueAtContractEnd = Math.max(0, netVehicleCost - O.annualDep*cYears);
}

const loanPct = n(inp.loanPct,85)/100;
O.loanAmount = O.netCapitalCost*loanPct;
O.ownCapAmount = O.netCapitalCost*(1-loanPct);
O.loanInstalmentPM = PMT(loanMonthlyRate, cMonths, -O.loanAmount);
O.ownCapAmortPM_before = PMT(ownMonthlyRate, cMonths, -O.ownCapAmount);
O.salvageAmortPM = PMT(ownMonthlyRate, cMonths, 0, -O.salvageValueAtContractEnd); // ANNEXURES!F23 uses the FULL salvage value, not a loan-share-scaled figure
O.netOwnCapAmortPM = O.ownCapAmortPM_before - O.salvageAmortPM;
O.totalCapitalPM = O.loanInstalmentPM + O.netOwnCapAmortPM;

// 8-year full depreciation schedule for the TTU Dashboard table
O.depSchedule=[]; let ov=O.netCapitalCost;
for(let y=1;y<=O.usefulLife;y++){ const dep=O.annualDep; const cv=Math.max(0,ov-dep); O.depSchedule.push({year:y,opening:ov,dep,closing:cv}); ov=cv; }

// ── Wages ──
const opRows = computeWageRole({
  basicBase:n(inp.opBasicBase,579), vdaBase:n(inp.opVDABase,339), escPct:n(inp.opEscalationPct,3.17),
  sa:n(inp.opSpecialAllowance,0), bonusPct:n(inp.opBonusPct,0), bonusEscPct:n(inp.bonusEscalationPct,4.85),
  bonusFreshEachYear:false, workingDays, pfRatePct:n(inp.pfRatePct,13), pfCap:n(inp.pfDailyCap,75),
  esiRatePct:n(inp.esiRatePct,3.25), esiThreshold:n(inp.esiThreshold,21000), ecAnnual:n(inp.ecInsuranceAnnual,2738),
  relieverFactor:n(inp.relieverDaysFactor,84), doubleOTHrs:n(inp.doubleOTHrs,3), headcount:n(inp.numOperators,1),
  extraDoubleFromY0:true, employeePfRatePct:n(inp.employeePfRatePct,12), // Wages_OT_PF_ESI!Q14 includes +2*L14 from Yr1
}, cYears);
const hlpRows = computeWageRole({
  basicBase:n(inp.hlpBasicBase,410), vdaBase:n(inp.hlpVDABase,240), escPct:n(inp.hlpEscalationPct,3.17),
  sa:n(inp.hlpSpecialAllowance,0), bonusPct:n(inp.hlpBonusPct,8.33), bonusEscPct:n(inp.bonusEscalationPct,4.85),
  bonusFreshEachYear:false, workingDays, pfRatePct:n(inp.pfRatePct,13), pfCap:n(inp.pfDailyCap,75),
  esiRatePct:n(inp.esiRatePct,3.25), esiThreshold:n(inp.esiThreshold,21000), ecAnnual:n(inp.ecInsuranceAnnual,2738),
  relieverFactor:n(inp.relieverDaysFactor,84), doubleOTHrs:n(inp.doubleOTHrs,3), headcount:n(inp.numHelpers,2),
  extraDoubleFromY0:true, employeePfRatePct:n(inp.employeePfRatePct,12),
}, cYears);
const supRows = computeSupervisorRole({
  basicBase:n(inp.supBasicBase,494), vdaBase:n(inp.supVDABase,287), escPct:n(inp.supEscalationPct,3.17),
  sa:n(inp.supSpecialAllowance,0), bonusPct:n(inp.supBonusPct,8.33), pfRatePct:n(inp.pfRatePct,13),
  pfCap:n(inp.pfDailyCap,75), esiRatePct:n(inp.esiRatePct,3.25), ttuShare:n(inp.supervisionTTUShare,4),
  employeePfRatePct:n(inp.employeePfRatePct,12),
}, cYears);
O.opRows=opRows; O.hlpRows=hlpRows; O.supRows=supRows;

const numOp=n(inp.numOperators,1), numHlp=n(inp.numHelpers,2);
O.wagesPM=[]; O.otAnnual=[]; O.supervisionPM=[];
O.statutoryPM=[]; O.laborCostInclStatutoryPM=[]; O.laborCostAnnual=[]; O.dailyLaborCost=[];
O.pfPM=[]; O.esiPM=[]; O.insurancePM=[]; O.employeePfPM=[];
for(let y=0;y<cYears;y++){
  const op=opRows[y], hlp=hlpRows[y], sup=supRows[y];
  const wagesPM = op.monthlyRate*numOp + op.relieverMonthly*numOp
                + hlp.monthlyRate*numHlp + hlp.relieverMonthly*numHlp;
  O.wagesPM.push(wagesPM);
  O.otAnnual.push(op.otAnnual + hlp.otAnnual);
  O.supervisionPM.push(sup.monthlyShare);
  // Employer's statutory outgo (PF + ESI/EC Insurance) for the whole crew this year —
  // op/hlp per Wages_OT_PF_ESI columns Y & Z; supervisor per columns K & L (daily→monthly ×365/12).
  const supShareFactor = 1/n(inp.supervisionTTUShare,4);
  const supStatutoryPM = (sup.pfDaily+sup.esiDaily)*(365/12)*supShareFactor;
  const statutoryPM = op.monthlyPF*numOp + op.monthlyESIorEC*numOp
                     + hlp.monthlyPF*numHlp + hlp.monthlyESIorEC*numHlp
                     + supStatutoryPM;
  O.statutoryPM.push(statutoryPM);
  O.laborCostInclStatutoryPM.push(wagesPM + sup.monthlyShare + statutoryPM);
  O.laborCostAnnual.push((wagesPM + sup.monthlyShare + statutoryPM)*12);
  O.dailyLaborCost.push((wagesPM + sup.monthlyShare + statutoryPM)/workingDays);
  // PF / ESI / Insurance broken out separately (whole crew, per month)
  O.pfPM.push(op.monthlyPF*numOp + hlp.monthlyPF*numHlp + sup.monthlyPF*supShareFactor);
  O.esiPM.push(op.monthlyESI*numOp + hlp.monthlyESI*numHlp + sup.monthlyESI*supShareFactor);
  O.insurancePM.push(op.monthlyInsurance*numOp + hlp.monthlyInsurance*numHlp);
  // Employee's own PF share — informational only, NOT part of contractor's billed cost
  O.employeePfPM.push(op.employeePfMonthly*numOp + hlp.employeePfMonthly*numHlp + sup.employeePfMonthly*supShareFactor);
}
O.avgMonthlyWages = O.wagesPM.reduce((s,w)=>s+w,0)/cYears;
O.totalOTContract = O.otAnnual.reduce((s,v)=>s+v,0);
O.avgStatutoryPM = O.statutoryPM.reduce((s,v)=>s+v,0)/cYears;
O.avgLaborCostInclStatutoryPM = O.laborCostInclStatutoryPM.reduce((s,v)=>s+v,0)/cYears;
O.totalLaborCostContract = O.laborCostAnnual.reduce((s,v)=>s+v,0);
O.laborCostPerKM = kmPM>0 ? O.avgLaborCostInclStatutoryPM/kmPM : 0;
// "Minimum wage" headline figures (Year-1, per Wages_OT_PF_ESI col G/H — excl PF & ESI)
O.minDailyWageOp = opRows[0].dailyRate; O.minMonthlyWageOp = opRows[0].monthlyRate;
O.minDailyWageHlp = hlpRows[0].dailyRate; O.minMonthlyWageHlp = hlpRows[0].monthlyRate;

// ── Uniform / Medical / Toll (contract-level addl outgo) ──
O.uniformShoeTotal = n(inp.uniformShoePerHeadPerYear,6000)*n(inp.numUniformPersons,3)*cYears;
O.medicalCertTotal = n(inp.medicalCertPerHeadPerYear,3000)*n(inp.numMedicalPersons,3)*n(inp.medicalCertYearMultiplier,2);
O.tollPerYear = n(inp.tollPerTripAmount,1070)*n(inp.tripsPerYear,12);
O.tollTotal = O.tollPerYear*cYears;

// PF / ESI-Insurance addl outgo for contract (Annexure-III rows 6.1/7.1)
let pfAddl=0, insAddl=0;
for(let y=0;y<cYears;y++){
  pfAddl += opRows[y].monthlyPF*12*numOp + hlpRows[y].monthlyPF*12*numHlp;
  insAddl+= opRows[y].monthlyESIorEC*12*numOp + hlpRows[y].monthlyESIorEC*12*numHlp;
}
O.pfAddlContract = pfAddl; O.esiInsAddlContract = insAddl;

// ── Fuel ──
const fuelEff = n(inp.fuelEfficiency,2.5);
O.dieselPerKm = fuelEff>0 ? n(inp.dieselPrice,89.77)/fuelEff : 0;
const defPrice = (n(inp.defPriceLow,74.15)+n(inp.defPriceHigh,135))/2;
O.defPrice = defPrice; // display only (Interest_Maint_Insurance/ANNEXURES!F55, quote average)
// ANNEXURES!F52 — the DEF cost/km formula uses a literal hardcoded 104.58, NOT the F55
// quote-average (104.575) computed above. Kept as a separate constant to match exactly.
O.defPerKm = fuelEff>0 ? n(inp.defCostConstant,104.58)*(1/fuelEff)*(n(inp.defConsumptionPctOfFuel,3)/100) : 0;
O.lubePerKm = O.dieselPerKm*(n(inp.lubeOilPct,0.5)/100);
O.fuelLitPerMonth = fuelEff>0 ? kmPM/fuelEff : 0;
O.dieselCostPM = O.dieselPerKm*kmPM;
O.defLitPM = O.fuelLitPerMonth*(n(inp.defConsumptionPctOfFuel,3)/100);
O.defCostPM = O.defPerKm*kmPM;
O.lubeCostPM = O.lubePerKm*kmPM;
O.totalFuelPM = O.dieselCostPM+O.defCostPM+O.lubeCostPM;
O.totalFuelPerKm = O.dieselPerKm+O.defPerKm+O.lubePerKm;

// ── Tyre (Annexure formula incl. "-1 spare" quirk) ──
const tyreLife=n(inp.tyreLife,30000), numTyres=n(inp.numTyres,14), tyreCost=n(inp.tyreCostExclGST);
const rawChanges = tyreLife>0 ? (kmPM*cMonths)/tyreLife - n(inp.tyreSpareAdj,1) : 0;
O.tyreChangesInContract = Math.round((rawChanges<0?0:rawChanges)*10000)/10000;
O.tyrePerKm = (kmPM*cMonths)>0 ? (tyreCost*O.tyreChangesInContract*numTyres)/(kmPM*cMonths) : 0;
O.tyreCostForContract = O.tyrePerKm*kmPM*cMonths;
O.tyreCostPM = O.tyreCostForContract/cMonths;

// ── Battery (Excel per-km formula incl. escalation quirk) ──
const batLife=n(inp.batteryLife,2);
const batCostBasis = (n(inp.batteryUnitPrice,16500)/n(inp.batteryGSTFactor,1.18))*batLife;
O.batterySetCostDisplay = n(inp.batteryUnitPrice,16500)+n(inp.batteryUnitPrice,16500)/n(inp.batteryGSTFactor,1.18);
O.batteryPerKm = (kmPM*cMonths)>0 ? (batCostBasis*Math.pow(1+n(inp.batteryEscalationPct,1.34)/100,batLife))/(kmPM*cMonths) : 0;
O.batteryCostPM = O.batteryPerKm*kmPM;

// ── Maintenance ──
const mPcts=[n(inp.maintPct_yr1,1),n(inp.maintPct_yr2,2),n(inp.maintPct_yr3,3),n(inp.maintPct_yr4,4)];
O.maintYr=[]; O.maintAmortPM=[];
for(let y=0;y<cYears;y++){
  const pct = mPcts[y]!==undefined ? mPcts[y] : mPcts[3];
  const yearly = O.netCapitalCost*pct/100;
  const pv = yearly/Math.pow(1+ownRate,y+1);
  const amort = PMT(ownMonthlyRate,12,-pv);
  O.maintYr.push(yearly); O.maintAmortPM.push(amort);
}
O.avgMonthlyMaint = O.maintAmortPM.reduce((s,v)=>s+v,0)/cYears;
O.maintPerKm = kmPM>0 ? O.avgMonthlyMaint/kmPM : 0;

// ── Insurance (age-parameterised) ──
const insCfg = {odRateUpto5yr:n(inp.odRateUpto5yr,1.19), odRateUpto7yr:n(inp.odRateUpto7yr,1.22),
  odRateAbove7yr:n(inp.odRateAbove7yr,1.25), gvw:n(inp.gvw,40000),
  liabilityTractor:n(inp.liabilityTractor,785), liabilityTrailer:n(inp.liabilityTrailer,375),
  addlCoverRatePerKg:n(inp.addlCoverRatePerKg,0.27), addlCoverGVWThreshold:n(inp.addlCoverGVWThreshold,12000),
  specialDiscountPct:n(inp.specialDiscountPct,35), tpiPremium:n(inp.tpiPremium,46435),
  paCoverPerPerson:n(inp.paCoverPerPerson,100), numOperators:numOp, numHelpers:numHlp, gstPct:n(inp.gstPct,18),
  insFixedDepPct:n(inp.insFixedDepPct,50),
  ncbPctYr1:n(inp.ncbPctYr1,20), ncbPctYr2:n(inp.ncbPctYr2,20), ncbPctYr3:n(inp.ncbPctYr3,25), ncbPctYr4:n(inp.ncbPctYr4,25)};
O.insuranceYears=[]; O.insuranceAmortPM=[];
// capital used for IDV should reflect the (possibly overridden) vehicle cost proportion
const vehShare = O.vehicleCost>0 ? 1 : 0;
const tractorForIns = capitalOverride!=null ? O.netCapitalCost*(n(inp.tractorCost)/(O.vehicleCost||1)) : n(inp.tractorCost);
const trailerForIns = capitalOverride!=null ? O.netCapitalCost*(n(inp.trailerCost)/(O.vehicleCost||1)) : n(inp.trailerCost);
const fixedInsurance = capitalOverride==null; // standard/new-TTU calc uses the workbook's flat 50%/1.22% every year; Vintage TTU uses age-graded tariff
for(let y=0;y<cYears;y++){
  const age = startAge+y;
  const yr = computeInsuranceYear(tractorForIns, trailerForIns, age, y, insCfg, fixedInsurance);
  O.insuranceYears.push(yr);
  O.insuranceAmortPM.push(PMT(ownMonthlyRate,12,-yr.totalIncGST));
}
// Interest_Maint_Insurance!G195 = ROUND(SUM(yearly amortisations)/contractYears, -1) — rounds to the nearest 10
O.avgMonthlyInsurance = Math.round((O.insuranceAmortPM.reduce((s,v)=>s+v,0)/cYears)/10)*10;

// ── Road Tax & Permit (ANNEXURES!F26 — single flat annual figure, same every contract year) ──
O.roadTaxYears=[];
for(let y=0;y<cYears;y++){
  const only = y===0 ? n(inp.roadTaxOnly_yr1,29980) : n(inp.roadTaxOnly_yrRest,29980);
  const total = only+n(inp.permitFee,0)+n(inp.pollutionCert,0);
  O.roadTaxYears.push({only,permit:n(inp.permitFee,0),pollution:n(inp.pollutionCert,0),total});
}
O.avgRoadTaxPerYear = O.roadTaxYears.reduce((s,r)=>s+r.total,0)/cYears;
O.roadTaxAmortPM = PMT(ownMonthlyRate,12,-O.avgRoadTaxPerYear);

// ── Fixed Charge / Month (year-wise, Annexure-II A) ──
O.fixedChargePM=[];
for(let y=0;y<cYears;y++){
  O.fixedChargePM.push(O.wagesPM[y]+O.loanInstalmentPM+O.netOwnCapAmortPM+O.avgMonthlyInsurance+O.roadTaxAmortPM+O.supervisionPM[y]);
}
O.avgFixedChargePM = O.fixedChargePM.reduce((s,f)=>s+f,0)/cYears;

// ── Variable Charge / KM (Annexure-II E) ──
O.totalVarPerKm = O.dieselPerKm+O.defPerKm+O.batteryPerKm+O.tyrePerKm+O.lubePerKm+O.maintPerKm;

// ── Contract Totals (Annexure-III) ──
O.totalKmContract = kmPM*cMonths;
O.totalFixedForContract = O.avgFixedChargePM*12*cYears;
O.totalVarForContract = O.totalVarPerKm*O.totalKmContract;
O.baseContractCost = O.totalFixedForContract+O.totalVarForContract;
O.grandTotal = O.baseContractCost+O.totalOTContract+O.uniformShoeTotal+O.medicalCertTotal
              +O.pfAddlContract+O.esiInsAddlContract+O.tollTotal;
O.grandTotalWithGST = O.grandTotal*(1+n(inp.gstPct,18)/100);
O.perKmExGST = O.totalKmContract>0 ? O.grandTotal/O.totalKmContract : 0;
O.perKmWithGST = O.totalKmContract>0 ? O.grandTotalWithGST/O.totalKmContract : 0;

// ── Monthly / Annual "rate card" figures ──
O.monthlyVarCost = O.totalVarPerKm*kmPM;
O.monthlyTotalCost = O.avgFixedChargePM+O.monthlyVarCost;
O.annualTotalCost = O.monthlyTotalCost*12;

// ── Year-wise contract table for dashboard ──
for(let y=1;y<=cYears;y++){
  const fixedPA=(O.fixedChargePM[y-1]||O.fixedChargePM[0])*12;
  const varPA=O.totalVarPerKm*kmPM*12;
  O['yr'+y+'_fixed']=fixedPA; O['yr'+y+'_var']=varPA; O['yr'+y+'_maint']=O.maintYr[y-1]||0;
  O['yr'+y+'_total']=fixedPA+varPA;
}
O.cYears=cYears; O.cMonths=cMonths; O.kmPM=kmPM;

// ── % breakdown for dashboard bars ──
const tot=O.annualTotalCost||1;
O.pctCapital=(O.totalCapitalPM*12)/tot*100;
O.pctLabour=(O.avgMonthlyWages*12+O.supervisionPM.reduce((s,v)=>s+v,0)/cYears*12)/tot*100;
O.pctFuel=(O.totalFuelPerKm*kmPM*12)/tot*100;
O.pctTyre=(O.tyrePerKm*kmPM*12)/tot*100;
O.pctBattery=(O.batteryPerKm*kmPM*12)/tot*100;
O.pctMaint=(O.maintPerKm*kmPM*12)/tot*100;
O.pctIns=(O.avgMonthlyInsurance*12)/tot*100;
O.pctRoadTax=(O.roadTaxAmortPM*12)/tot*100;
O.pctToll=(O.tollPerYear)/tot*100;

return O;

},

calc(inp, accessories=[]){ return this.computeAll(inp, accessories, 0, null); },

// ─── VINTAGE TTU ───calcVintage(inp, v){const newTractor = n(v.newTractorPrice)||n(inp.tractorCost);const newTrailer = n(v.newTrailerPrice)||n(inp.trailerCost);const newPrice = newTractor+newTrailer;const ageBreakdown = calendarAgeBreakdown(v.purchaseDate||'2022-04-01', v.currentDate||new Date().toISOString().slice(0,10));const age = Math.max(0, ageBreakdown.fractionalYears);const usefulLife = n(v.usefulLife,8);const salvagePct = n(v.salvagePct,5);const salvageValue = newPrice*(salvagePct/100);const annualDep = (newPrice-salvageValue)/usefulLife;const accumDep = annualDep*Math.min(age,usefulLife);const currentValueRaw = newPrice-accumDep;const currentValue = Math.max(currentValueRaw, salvageValue);const remainingLife = Math.max(usefulLife-age,0);

const vInp = {...inp, tractorCost:newTractor, trailerCost:newTrailer, contractPeriod:n(v.contractPeriod,4)};
const O = this.computeAll(vInp, [], age, currentValue);

O.vintage = {
  age, ageBreakdown, newPrice, annualDep, accumDep, currentValue, salvageValue, remainingLife,
  monthlyCapitalCost: O.totalCapitalPM,
  estimatedFixedCharge: O.avgFixedChargePM,
  estimatedVarPerKm: O.totalVarPerKm,
  totalContractCost: O.grandTotal,
  totalContractCostGST: O.grandTotalWithGST,
};

// Comparison vs a brand-new TTU of the same spec
const newO = this.computeAll({...inp, tractorCost:newTractor, trailerCost:newTrailer, contractPeriod:n(v.contractPeriod,4)}, [], 0, null);
O.vintage.newComparison = {
  capitalCost:newO.netCapitalCost, fixedCharge:newO.avgFixedChargePM, varPerKm:newO.totalVarPerKm,
  totalContractCost:newO.grandTotal,
  savings:newO.grandTotal-O.grandTotal,
  savingsPct: newO.grandTotal>0 ? (newO.grandTotal-O.grandTotal)/newO.grandTotal*100 : 0,
};
return O;

}};

// ═══════════════════════════════════════════════════════════════// STATE// ═══════════════════════════════════════════════════════════════const State = {inputs, outputs, accessories:[], listeners:[],vintage, vintageOutputs,projId, saveTimer, undoStack:[], redoStack:[], changed Set(), defaults,customDefaults:{}, // {20MT:{...}, 30MT:{...}} — user-saved defaults, overrides factory constants

getFactoryDefaults(type){ return type==='30MT' ? {...DEFAULTS_30MT} : {...DEFAULTS_20MT}; },getDefaults(type){ return this.customDefaults[type] ? {...this.customDefaults[type]} : this.getFactoryDefaults(type); },getVintageDefaults(){ return {purchaseDate:'2022-04-01',currentDate Date().toISOString().slice(0,10),salvagePct:5,usefulLife:8,contractPeriod:4,newTractorPrice,newTrailerPrice}; },

async loadCustomDefaults(){for(const type of ['20MT','30MT']){try{ const r=await DB.g('settings','customDefault'+type); if(r&&r.value) this.customDefaults[type]=r.value; }catch(_){}}},async saveCurrentAsDefault(){const type=this.inputs.ttuType||'20MT';const factory=this.getFactoryDefaults(type);const snapshot={...this.inputs, projectName.projectName, ttuType};this.customDefaults[type]=snapshot;await DB.p('settings',{id:'customDefault'+type,value});this.defaults={...snapshot}; this.changed=new Set(); this._recomputeAll();},async restoreFactoryDefaults(){const type=this.inputs.ttuType||'20MT';delete this.customDefaults[type];await DB.d('settings','customDefault'+type).catch(()=>{});this._push();const d=this.getFactoryDefaults(type);this.defaults=d; this.inputs={...d}; this.changed=new Set(); this.accessories=[];this.vintage=this.getVintageDefaults();this._recomputeAll(); this._scheduleSave();},

init(){const d=this.getDefaults('20MT');this.inputs={...d}; this.defaults=d; this.accessories=[];this.vintage=this.getVintageDefaults();this._recomputeAll();},setInput(k,v){this._push(); this.inputs[k]=v;const def=this.defaults||{};(Math.abs(n(v)-n(def[k]))<0.0001) ? this.changed.delete(k) : this.changed.add(k);this._recompute(); this._scheduleSave();},setVintageInput(k,v){this._push(); this.vintage[k]=v; this._recomputeVintage(); this._scheduleSave();},switchTTUType(type){const d=this.getDefaults(type);this.defaults=d; this.inputs={...d,projectName.inputs.projectName||'New TTU Project',ttuType};this.changed=new Set(); this.accessories=[];this._recomputeAll(); this._scheduleSave();},resetToDefaults(){this._push(); const d=this.getDefaults(this.inputs.ttuType||'20MT');this.defaults=d; this.inputs={...d}; this.changed=new Set(); this.accessories=[];this.vintage=this.getVintageDefaults();this._recomputeAll(); this._scheduleSave();},loadProject(p){this.projId=p.id;const type=p.inputs?.ttuType||'20MT';const d=this.getDefaults(type);this.defaults=d; this.inputs={...d,...(p.inputs||{})};this.accessories=p.accessories||[];this.vintage={...this.getVintageDefaults(),...(p.vintage||{})};// Migrate legacy (pre-date-picker) saves that stored purchaseYear/currentYear only.if(p.vintage && !p.vintage.purchaseDate && p.vintage.purchaseYear) this.vintage.purchaseDate=p.vintage.purchaseYear+'-01-01';if(p.vintage && !p.vintage.currentDate && p.vintage.currentYear) this.vintage.currentDate=p.vintage.currentYear+'-01-01';this.changed=new Set(Object.keys(p.inputs||{}).filter(k=>n(p.inputs[k])!==n(d[k])));this.undoStack=[]; this.redoStack=[]; this._recomputeAll();},addAccessory(a){ this.push(); this.accessories.push({id:'acc'+Date.now()+Math.random().toString(36).slice(2,6),...a}); this._recompute(); this._scheduleSave(); },removeAccessory(id){ this._push(); this.accessories=this.accessories.filter(a=>a.id!==id); this._recompute(); this._scheduleSave(); },updateAccessory(id,key,val){ this._push(); const a=this.accessories.find(x=>x.id===id); if(a)a[key]=val; this._recompute(); this._scheduleSave(); },undo(){ if(!this.undoStack.length)return false; this.redoStack.push({i:{...this.inputs},a:[...this.accessories],c Set(this.changed),v:{...this.vintage}}); const s=this.undoStack.pop(); this.inputs=s.i; this.accessories=s.a; this.changed=s.c; this.vintage=s.v||this.vintage; this._recomputeAll(); return true; },redo(){ if(!this.redoStack.length)return false; this.undoStack.push({i:{...this.inputs},a:[...this.accessories],c Set(this.changed),v:{...this.vintage}}); const s=this.redoStack.pop(); this.inputs=s.i; this.accessories=s.a; this.changed=s.c; this.vintage=s.v||this.vintage; this._recomputeAll(); return true; },on(fn){ this.listeners.push(fn); },_recompute(){ this.outputs=Engine.calc(this.inputs,this.accessories); this._recomputeVintage(); this.listeners.forEach(fn=>fn(this.inputs,this.outputs)); },_recomputeVintage(){if(!this.vintage.newTractorPrice) this.vintage.newTractorPrice=n(this.inputs.tractorCost);if(!this.vintage.newTrailerPrice) this.vintage.newTrailerPrice=n(this.inputs.trailerCost);this.vintageOutputs=Engine.calcVintage(this.inputs,this.vintage);},_recomputeAll(){ this.outputs=Engine.calc(this.inputs,this.accessories); this._recomputeVintage(); this.listeners.forEach(fn=>fn(this.inputs,this.outputs)); },_push(){ this.undoStack.push({i:{...this.inputs},a.parse(JSON.stringify(this.accessories)),c Set(this.changed),v:{...this.vintage}}); if(this.undoStack.length>60)this.undoStack.shift(); this.redoStack=[]; },_scheduleSave(){const sl=document.getElementById('saveLabel'); if(sl)sl.textContent='Saving…';const sd=document.getElementById('saveDot'); if(sd)sd.style.background='var(--ye)';clearTimeout(this.saveTimer); this.saveTimer=setTimeout(()=>DB.save(),1600);}};

// ═══════════════════════════════════════════════════════════════// INDEXEDDB// ═══════════════════════════════════════════════════════════════const DB = {db,async open(){return new Promise((res,rej)=>{const r=indexedDB.open('TTUv6',1);r.onupgradeneeded=e=>{ const d=e.target.result; if(!d.objectStoreNames.contains('projects'))d.createObjectStore('projects',{keyPath:'id'}); if(!d.objectStoreNames.contains('settings'))d.createObjectStore('settings',{keyPath:'id'}); };r.onsuccess=e=>{ this.db=e.target.result; res(); };r.onerror=e=>rej(e.target.error);});},_tx(s,m='readonly'){ return this.db.transaction(s,m).objectStore(s); },_g(s,k){ return new Promise((res,rej)=>{ const r=this._tx(s).get(k); r.onsuccess=()=>res(r.result); r.onerror=()=>rej(); }); },_p(s,v){ return new Promise((res,rej)=>{ const r=this._tx(s,'readwrite').put(v); r.onsuccess=()=>res(); r.onerror=()=>rej(); }); },_d(s,k){ return new Promise((res,rej)=>{ const r=this._tx(s,'readwrite').delete(k); r.onsuccess=()=>res(); r.onerror=()=>rej(); }); },_a(s){ return new Promise((res,rej)=>{ const r=this.tx(s).getAll(); r.onsuccess=()=>res(r.result); r.onerror=()=>rej(); }); },async save(){const now=new Date().toISOString();const p={id.projId||('proj'+Date.now()),name.inputs.projectName||'Untitled',inputs:{...State.inputs},outputs:{...State.outputs},accessories.parse(JSON.stringify(State.accessories)),vintage:{...State.vintage},savedAt};await this._p('projects',p); await this._p('settings',{id:'lastId',value.id});State.projId=p.id;const sl=document.getElementById('saveLabel'); if(sl)sl.textContent='Saved';const sd=document.getElementById('saveDot'); if(sd)sd.style.background='var(--gr)';return p;},getAll(){ return this._a('projects'); },get(id){ return this._g('projects',id); },del(id){ return this.d('projects',id); },async dup(id){ const p=await this.get(id); if(!p)return; const c={...p,id:'proj'+Date.now(),name.name+' (Copy)',savedAt Date().toISOString()}; await this._p('projects',c); return c; }};

// ═══════════════════════════════════════════════════════════════// FORMATTERS & BUILDERS// ═══════════════════════════════════════════════════════════════function inrGrouped2dp(v){if(v==null||isNaN(v)) return '—';const neg = v<0; v=Math.abs(v);const [intPart,decPart]=v.toFixed(2).split('.');const grouped=intPart.replace(/\B(?=(\d\d)+(\d)(?!\d))/g,',');return (neg?'-':'')+grouped+'.'+decPart;}const F = v=>(v==null||isNaN(v))?'—':'₹\u202f'+inrGrouped2dp(v);const F2 = v=>(v==null||isNaN(v))?'—':'₹\u202f'+inrGrouped2dp(v);const esc = s=>String(s??'').replace(/&/g,'&').replace(/</g,'<').replace(/>/g,'>').replace(/"/g,'"');

// ── Numeric field helpers (cursor-stable editing) ──────────────// Raw value: plain digits/decimal/minus, no separators — used while the// field is focused and being typed into.function numRaw(v){if(v===null||v===undefined||v==='') return '';const s=String(v);return s;}// Display value: Indian thousands-separated, shown only when the field// is NOT focused (i.e. after blur / on initial render).function numDisplay(v){if(v===null||v===undefined||v==='') return '';const n=parseFloat(v);if(isNaN(n)) return String(v);// Preserve up to 2 decimal places only if present in the source value.const s=String(v);const dot=s.indexOf('.');const decimals = dot>=0 ? Math.min(s.length-dot-1,2) : 0;const parts = n.toFixed(decimals).split('.');parts[0] = parts[0].replace(/^-/,'').replace(/\B(?=(\d\d)+(\d)(?!\d))/g,',');const sign = n<0 ? '-' : '';return sign + parts.join('.');}// Strip thousands separators back to a raw editable numeric string.function numUnformat(s){return String(s??'').replace(/,/g,'');}// Sanitize a raw string as the user types: keep digits, at most one '.',// and an optional leading '-'. Returns {value,cursorShift} where// cursorShift is how many characters were removed before the cursor.function sanitizeNumericTyping(raw,cursorPos,allowNegative,allowDecimal=true){let out='', seenDot=false, removedBeforeCursor=0;for(let i=0;i<raw.length;i++){const c=raw[i];let keep=false;if(c>='0'&&c<='9') keep=true;else if(c==='.'&&allowDecimal&&!seenDot){ keep=true; seenDot=true; }else if(c==='-'&&allowNegative&&out.length===0) keep=true;if(keep) out+=c; else if(i<cursorPos) removedBeforeCursor++;}return {value, cursorShift};}

function INP(key,label,val,type='number',ph='',step='any',min='',max='',unit=''){const ch=State.changed.has(key)?' chg':'';const dot=State.changed.has(key)?'<span class="tye" style="font-size:9px"> ●</span>':'';const isYear = type==='year';const isNum = type==='number' || isYear;const allowNeg = min!=='' ? (parseFloat(min)<0) : true;// Years are identifiers, not financial quantities — never comma-group them.const shown = isYear ? String(Math.round(n(val,0))) : (isNum ? numDisplay(val) : esc(String(val??'')));return <div class="frow">
    <label class="fl">${label}${dot}</label>
    <input type="${isNum?'text':type}" ${isNum?inputmode="${(isYear||step==='1')?'numeric':'decimal'}":''} class="fi${ch}" value="${shown}" data-key="${key}"
      ${isNum?data-numeric="1" data-allow-neg="${allowNeg?1:0}":''}${isYear?' data-integer="1" data-nogroup="1"':''}
      placeholder="${esc(ph)}"
      ${min!==''?data-min="${min}":''}${max!==''? data-max="${max}":''}
      oninput="UI.inp(this)" onfocus="UI.inpFocus(this)" onblur="UI.inpBlur(this)">
    ${unit?<div style="font-size:10px;color:var(--mu);margin-top:2px">${unit}</div>`:''}

  </div>`;
}
function RINP(key,label,val){
  const ch=State.changed.has(key)?' chg':'';
  const dot=State.changed.has(key)?'<span class="tye" style="font-size:9px"> ●</span>':'';
  return `<div class="frow">
    <label class="fl">${label}${dot}</label>
    <div class="ig"><span class="pfx">₹</span>
      <input type="text" inputmode="decimal" class="fi${ch}" value="${numDisplay(val??0)}" data-key="${key}"
        data-numeric="1" data-allow-neg="0" data-min="0"
        oninput="UI.inp(this)" onfocus="UI.inpFocus(this)" onblur="UI.inpBlur(this)">
    </div>
  </div>`;
}
function SEL(key,label,val,opts){
  return `<div class="frow">
    <label class="fl">${label}</label>
    <select class="fs" data-key="${key}" onchange="UI.inp(this)">
      ${opts.map(([v,t])=>`<option value="${v}" ${v==val?'selected':''}>${t}</option>`).join('')}
    </select>
  </div>`;
}
function VINP(key,label,val,type='number',step='any'){
  const isNum = type==='number';
  const shown = isNum ? numDisplay(val) : esc(String(val??''));
  return `<div class="frow">
    <label class="fl">${label}</label>
    <input type="${isNum?'text':type}" ${isNum?'inputmode="decimal"':''} class="fi" value="${shown}" data-vkey="${key}"
      ${isNum?'data-numeric="1" data-allow-neg="0"':''}
      oninput="UI.vinp(this)" onfocus="UI.vinpFocus(this)" onblur="UI.vinpBlur(this)">
  </div>`;
}
function ageLabel(b){
  if(!b) return '—';
  return `${b.years} Yr${b.years!==1?'s':''} ${b.months} Mo ${b.days} Day${b.days!==1?'s':''}`;
}
function DATE_INP(key,label,val){
  return `<div class="frow">
    <label class="fl">${label}</label>
    <input type="date" class="fi" value="${esc(val||'')}" data-vkey="${key}"
      oninput="UI.vinp(this)" onblur="UI.vinpBlur(this)">
  </div>`;
}
function slugKey(s){ return String(s).toLowerCase().replace(/[^a-z0-9]+/g,'_').replace(/^_+|_+$/g,''); }
function COUT(label,value,formula,unit=''){
  return `<div class="frow">
    <label class="fl">${label}</label>
    <div class="co" data-outkey="${slugKey(label)}">${esc(value)}</div>
    ${formula?`<div class="ftag">${esc(formula)}</div>`:''}
    ${unit?`<div style="font-size:10px;color:var(--mu)">${unit}</div>`:''}
  </div>`;
}
function KPI(label,val,color,sub,infoKey){
  return `<div class="kpi" style="--kc:${color}" data-outkey="${slugKey(label)}">
    <div class="kl">${label}${infoKey?` <button type="button" class="finfo" onclick="UI.showFormula('${infoKey}')" title="How is this calculated?">ⓘ</button>`:''}</div>
    <div class="kv" style="color:${color}">${val}</div>
    <div class="ks">${sub}</div>
  </div>`;
}
function SR(label,val,color){
  return `<div class="sr"><span class="flex ic"><span class="sdot" style="background:${color}"></span>${label}</span><span class="mono bold">${F(val)}</span></div>`;
}
function mrow(l,v){
  return `<div style="display:flex;justify-content:space-between;align-items:center;padding:5px 0;border-bottom:1px solid var(--div)"><span style="font-size:11px;color:var(--mu)">${l}</span><span class="mono" style="font-size:12px">${v}</span></div>`;
}
function rbar(){
  if(!State.changed.size) return '';
  return `<div class="rbar no-print"><span>⚠️ <strong>${State.changed.size}</strong> field${State.changed.size!==1?'s':''} modified from engineering defaults</span><button class="btn bw sm" onclick="UI.confirmReset()">↺ Reset to Defaults</button></div>`;
}
function SEL_V(key,label,val,opts){
  return `<div class="frow">
    <label class="fl">${label}</label>
    <select class="fs" data-vkey="${key}" onchange="UI.vinp(this)">
      ${opts.map(([v,t])=>`<option value="${v}" ${v==val?'selected':''}>${t}</option>`).join('')}
    </select>
  </div>`;
}
// Renders a single role's (Operator / Helper) PF + ESI-vs-Insurance rule
// trace as a visible flow, so the user always sees WHICH rule fired and
// WHY — instead of the value silently switching between ESI and EC
// Insurance in the background.
function decisionPanel(roleLabel,row,esiThreshold,esiRatePct,pfRatePct,pfCap,ecAnnual){
  const th=n(esiThreshold,21000), isESI=row.mode==='ESI';
  const statusCls=isESI?'esi':'ins';
  const statusTxt=isESI?"✓ ESI Applicable":"✓ Employees' Compensation Insurance Applicable";
  const ruleTxt=isESI?"ESI applied · EC Insurance not applicable":"EC Insurance applied · ESI removed (wage over threshold)";
  return `<div data-outkey="decision_${slugKey(roleLabel)}">
    <div style="font-size:11px;font-weight:700;color:var(--tx2);text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px">${roleLabel}</div>
    <div class="dflow">
      <div class="dstep"><span class="dlbl">Monthly Wage</span><span class="dval">${F(row.monthlyRate)}</span></div>
      <div class="darrow">↓</div>
      <div class="dstep"><span class="dlbl">Threshold (statutory ESI ceiling)</span><span class="dval">${F(th)}</span></div>
      <div class="darrow">↓</div>
      <div class="dstep"><span class="dlbl">Test</span><span class="dval">Monthly Wage ${isESI?'&lt;':'≥'} ${F(th)} → ${isESI?'TRUE':'TRUE'}</span></div>
    </div>
    <div class="dstatus ${statusCls}" style="margin-top:10px">${statusTxt}</div>
    <div class="dreason">
      <div class="drbox"><div class="drl">Monthly PF</div><div class="drv">${F(row.monthlyPF)}</div></div>
      <div class="drbox"><div class="drl">Monthly ${isESI?'ESI':'Insurance'}</div><div class="drv">${F(row.monthlyESIorEC)}</div></div>
      <div class="drbox"><div class="drl">Employer Contribution</div><div class="drv">${F(row.monthlyPF+row.monthlyESIorEC)}</div></div>
    </div>
    <div class="callout mt3" style="font-family:inherit;white-space:normal">Applied Rule: <strong>${ruleTxt}</strong>. PF = MIN((Basic+VDA)×${pfRatePct}%, ₹${pfCap}/day cap) → ₹${round2ish(row.pfDaily)}/day.
${isESI?`ESI = (Basic+VDA)×${esiRatePct}%/day.`:`EC Insurance = ₹${n(ecAnnual,2738).toLocaleString('en-IN')}/year ÷ 365 days.`}</div>
  </div>`;
}
function round2ish(v){ return (Math.round(v*100)/100).toFixed(2); }
function BAR(label,val,total,color){
  const p=total>0?Math.max(0,Math.min(100,val/total*100)):0;
  return `<div class="bw2"><div class="bl2"><span class="bn">${label}</span><span class="bv">${F(val)} <span style="color:var(--mu)">(${p.toFixed(1)}%)</span></span></div><div class="bt"><div class="bf" style="width:${p}%;background:${color}"></div></div></div>`;
}

// ═══════════════════════════════════════════════════════════════// SMART FORMULA DISPLAY — click-to-explain registry for key// computed outputs. Each entry's detail(i,o) runs at click-time// so it always reflects the user's current inputs/outputs.// ═══════════════════════════════════════════════════════════════const FORMULA_DOCS = {minMonthlyWageOp:{title:'Monthly Minimum Wage — Operator',formula:'Monthly Wage = (Basic + VDA + Bonus + Spl.Allowance)/day × Working Days',detail:(i,o)=>{const r=o.opRows[0]; return [['Basic / day',F2(r.basic)],['VDA / day',F2(r.vda)],['Bonus / day',F2(r.bonus)],['Special Allowance / day',F2(r.sa)],['Working Days / month',i.workingDays||26],['= Daily Rate',F2(r.dailyRate)],['= Monthly Rate',F(r.monthlyRate)]];}},minMonthlyWageHlp:{title:'Monthly Minimum Wage — Helper',formula:'Monthly Wage = (Basic + VDA + Bonus + Spl.Allowance)/day × Working Days',detail:(i,o)=>{const r=o.hlpRows[0]; return [['Basic / day',F2(r.basic)],['VDA / day',F2(r.vda)],['Bonus / day',F2(r.bonus)],['Special Allowance / day',F2(r.sa)],['Working Days / month',i.workingDays||26],['= Daily Rate',F2(r.dailyRate)],['= Monthly Rate',F(r.monthlyRate)]];}},statutoryPM:{title:'Employer Statutory Contribution / Month (PF + ESI/Insurance)',formula:'PF = MIN(ROUND((Basic+VDA)×PF%,2), PF-cap/day); ESI or Insurance depends on the monthly-wage threshold test',detail:(i,o)=>{const op=o.opRows[0],hlp=o.hlpRows[0]; return [['PF Rate',i.pfRatePct+'%'],['PF Cap',F(i.pfDailyCap)+'/day'],['ESI Threshold (₹/month)',F(i.esiThreshold)],['Operator mode',op.mode+' (monthly wage '+F(op.monthlyRate)+')'],['Helper mode',hlp.mode+' (monthly wage '+F(hlp.monthlyRate)+')'],['Operator PF+ESI/Ins. per month',F(op.monthlyPF+op.monthlyESIorEC)],['Helper PF+ESI/Ins. per month (each)',F(hlp.monthlyPF+hlp.monthlyESIorEC)],['= Total statutory / month (whole crew)',F(o.statutoryPM[0])]];}},laborCostInclStatutoryPM:{title:'Monthly Labour Cost Including Statutory Charges',formula:'Wages (Op+RelOp+Hlp+RelHlp) + Supervisor share + Employer PF/ESI/Insurance',detail:(i,o)=>[['Wages / month (excl. statutory)',F(o.wagesPM[0])],['Supervisor share / month',F(o.supervisionPM[0])],['Employer statutory / month',F(o.statutoryPM[0])],['= Total Labour Cost / month',F(o.laborCostInclStatutoryPM[0])]],},laborCostPerKM:{title:'Labour Cost per KM',formula:'Avg. Monthly Labour Cost (incl. statutory) ÷ KM run per month',detail:(i,o)=>[['Avg. Monthly Labour Cost (incl. statutory)',F(o.avgLaborCostInclStatutoryPM)],['KM / month',F(o.kmPM)],['= Labour Cost / KM',F2(o.laborCostPerKM)]],},laborCostAnnual:{title:'Annual Labour Cost',formula:'Monthly Labour Cost (incl. statutory) × 12',detail:(i,o)=>[['Monthly Labour Cost (Yr 1)',F(o.laborCostInclStatutoryPM[0])],['× 12 months',''],['= Annual Labour Cost (Yr 1)',F(o.laborCostAnnual[0])]],},};

const Pages = {

dashboard(i,o){const total=o.annualTotalCost||1;return `${rbar()}

  <div class="sh mb4">
    <div>
      <div class="st">${esc(i.projectName)}</div>
      <div class="ss">${esc(i.ttuType)} TTU · OIL India · Contract: ${i.contractPeriod} Years · ${n(i.kmPerMonth).toLocaleString('en-IN')} KM/Month</div>
    </div>
    <div class="flex gap2">
      <button class="btn bs sm" onclick="UI.go('reports')">📄 Report</button>
      <button class="btn bp sm" onclick="Exp.csv()">⬇️ Export</button>
    </div>
  </div>
  <div class="gauto mb4">
    ${KPI('Fixed Charge / Month',F(o.avgFixedChargePM),'var(--ac)','avg. ex-GST per month')}
    ${KPI('Variable Charge / KM','₹ '+o.totalVarPerKm.toFixed(2),'#0da6e0','total variable per km')}
    ${KPI('Monthly Total',F(o.monthlyTotalCost),'#9775ea','ex-GST all-in')}
    ${KPI('Annual Total',F(o.annualTotalCost),'#1fb97a','ex-GST')}
    ${KPI(o.cYears+'-Yr Contract Cost',F(o.grandTotal),'var(--ye)','ex-GST, incl. reimbursables')}
    ${KPI('Grand Total incl. GST',F(o.grandTotalWithGST),'var(--re)','incl. '+i.gstPct+'% GST')}
  </div>
  <div class="g2 mb4">
    <div class="card">
      <div class="ch"><div class="ct">Annual Cost Breakdown</div></div>
      ${BAR('Capital (Loan+Own Capital)',o.totalCapitalPM*12,total,'#f47820')}
      ${BAR('Labour & Wages',(o.avgMonthlyWages+o.supervisionPM.reduce((s,v)=>s+v,0)/o.cYears)*12,total,'#0da6e0')}
      ${BAR('Fuel (Diesel+DEF+Lube)',o.totalFuelPerKm*o.kmPM*12,total,'#9775ea')}
      ${BAR('Tyre & Tubes',o.tyrePerKm*o.kmPM*12,total,'#1fb97a')}
      ${BAR('Battery',o.batteryPerKm*o.kmPM*12,total,'#f97316')}
      ${BAR('Maintenance (avg)',o.avgMonthlyMaint*12,total,'#06b6d4')}
      ${BAR('Insurance',o.avgMonthlyInsurance*12,total,'#a855f7')}
      ${BAR('Road Tax & Permits',o.roadTaxAmortPM*12,total,'#84cc16')}
      ${BAR('Toll',o.tollPerYear,total,'#14b8a6')}
    </div>
    <div class="card">
      <div class="ch"><div class="ct">Cost Summary</div></div>
      ${SR('Capital (Loan+Own Cap)',o.totalCapitalPM*12,'#f47820')}
      ${SR('Labour & Wages',(o.avgMonthlyWages+o.supervisionPM.reduce((s,v)=>s+v,0)/o.cYears)*12,'#0da6e0')}
      ${SR('Fuel (Diesel+DEF+Lube)',o.totalFuelPerKm*o.kmPM*12,'#9775ea')}
      ${SR('Tyre & Tubes',o.tyrePerKm*o.kmPM*12,'#1fb97a')}
      ${SR('Battery',o.batteryPerKm*o.kmPM*12,'#f97316')}
      ${SR('Maintenance (avg, yr-wise)',o.avgMonthlyMaint*12,'#06b6d4')}
      ${SR('Insurance',o.avgMonthlyInsurance*12,'#a855f7')}
      ${SR('Road Tax & Permits',o.roadTaxAmortPM*12,'#84cc16')}
      ${SR('Toll (reimbursable)',o.tollPerYear,'#14b8a6')}
      ${State.accessories.length?SR('Accessories',o.accessoriesCost,'#eaaa35'):''}
      <div class="sr stot"><span><strong>Monthly Total (ex-GST)</strong></span><span class="mono bold tac">${F(o.monthlyTotalCost)}</span></div>
    </div>
  </div>
  <div class="card mb4">
    <div class="ch"><div class="ct">${o.cYears}-Year Contract Cost Projection</div><span class="badge ba">Per Annexure-III Reference</span></div>
    <div class="tw"><table><thead><tr><th>Year</th><th class="nr">Fixed Cost (Annual)</th><th class="nr">Variable Cost (Annual)</th><th class="nr">Maint. Cost</th><th class="nr">Year Total</th><th class="nr">Cumulative</th></tr></thead>
    <tbody>${(()=>{let cum=0;let rows='';for(let y=1;y<=o.cYears;y++){const yt=o['yr'+y+'_total']||0;cum+=yt;rows+=`<tr><td><span class="badge ba">Year ${y} (${n(i.contractStartYear,2026)+y-1})</span></td><td class="nr">${F(o['yr'+y+'_fixed'])}</td><td class="nr">${F(o['yr'+y+'_var'])}</td><td class="nr">${F(o['yr'+y+'_maint'])}</td><td class="nr">${F(yt)}</td><td class="nr">${F(cum)}</td></tr>`}return rows})()}
    <tr class="ttr"><td><strong>Contract Total</strong></td><td class="nr">${F(o.totalFixedForContract)}</td><td class="nr">${F(o.totalVarForContract)}</td><td class="nr">${F(o.maintYr.reduce((s,m)=>s+m,0))}</td><td class="nr">${F(o.baseContractCost)}</td><td class="nr">${F(o.baseContractCost)}</td></tr>
    </tbody></table></div>
    <div class="callout mt3">Grand Total also includes: OT ${F(o.totalOTContract)} · Uniform &amp; Shoe ${F(o.uniformShoeTotal)} · Medical ${F(o.medicalCertTotal)} · Addl. PF ${F(o.pfAddlContract)} · Addl. ESI/Insurance ${F(o.esiInsAddlContract)} · Toll ${F(o.tollTotal)} (reimbursable) → <strong style="color:var(--ac)">Grand Total ex-GST ${F(o.grandTotal)}</strong></div>
  </div>
  <div class="g3">
    <div class="card"><div class="ch"><div class="ct">Fixed Charges/Month</div></div>
      ${o.fixedChargePM.map((f,idx)=>mrow('Year '+(idx+1)+' ('+(n(i.contractStartYear,2026)+idx)+')',F(f))).join('')}
      ${mrow('Average /Month',F(o.avgFixedChargePM))}
    </div>
    <div class="card"><div class="ch"><div class="ct">Variable Charges/KM</div></div>
      ${mrow('Fuel (Diesel)','₹ '+o.dieselPerKm.toFixed(2))}
      ${mrow('DEF','₹ '+o.defPerKm.toFixed(2))}
      ${mrow('Lube Oil','₹ '+o.lubePerKm.toFixed(2))}
      ${mrow('Tyre/Tube','₹ '+o.tyrePerKm.toFixed(2))}
      ${mrow('Battery','₹ '+o.batteryPerKm.toFixed(2))}
      ${mrow('Maintenance','₹ '+o.maintPerKm.toFixed(2))}
      ${mrow('Total/KM','₹ '+o.totalVarPerKm.toFixed(2))}
    </div>
    <div class="card"><div class="ch"><div class="ct">Rate Card</div></div>
      ${mrow('Monthly ex-GST',F(o.monthlyTotalCost))}
      ${mrow('Annual ex-GST',F(o.annualTotalCost))}
      ${mrow('Per KM ex-GST','₹ '+o.perKmExGST.toFixed(2))}
      ${mrow(o.cYears+'-Yr Grand Total ex-GST',F(o.grandTotal))}
      ${mrow('incl. GST ('+i.gstPct+'%)',F(o.grandTotalWithGST))}
    </div>
  </div>`;
},

ttu(i,o){const ttuOpts=[['20MT','20MT/24MT — Double Axle'],['30MT','30MT/35MT — Triple Axle']];return `${rbar()}

  <div class="sh mb4"><div><div class="st">TTU Dashboard</div><div class="ss">Vehicle specs, capital cost and depreciation per Companies Act 2013 Schedule II</div></div></div>
  <div class="g4 mb4">
    ${KPI('Net Capital Cost',F(o.vehicleCost+o.accessoriesCost),'var(--ac)','tractor + trailer + accessories')}
    ${KPI('Annual Depreciation',F(o.annualDep),'#0da6e0','SLM @ '+(o.depRateAnnual*100).toFixed(2)+'% p.a.')}
    ${KPI('Salvage @ contract end',F(o.salvageValueAtContractEnd),'#1fb97a','after '+o.cYears+' years')}
    ${KPI('Monthly Capital',F(o.totalCapitalPM),'#9775ea','loan + own cap/mo')}
  </div>
  <div class="g2 mb4">
    <div class="card"><div class="ch"><div class="ct">Vehicle Identity &amp; TTU Type</div></div>
      <div class="fg">
        ${INP('projectName','Project Name',i.projectName,'text')}
        ${SEL('ttuType','TTU Type',i.ttuType,ttuOpts)}
        ${INP('contractPeriod','Contract Period (Years)',i.contractPeriod,'number','','1','1','20')}
        ${INP('contractStartYear','Contract Start Year',i.contractStartYear,'year','','1','2020','2050')}
        ${INP('kmPerMonth','Avg KM per Month',i.kmPerMonth,'number','km','1','100')}
        ${INP('workingDays','Working Days/Month',i.workingDays,'number','','1','1','31')}
      </div>
    </div>
    <div class="card"><div class="ch"><div class="ct">Capital Costs (Excl. GST)</div></div>
      <div class="fg">
        ${RINP('tractorCost','Tractor Unit Cost (avg of 3 quotes)',i.tractorCost)}
        ${RINP('trailerCost','Trailer Cost (avg of 3 quotes)',i.trailerCost)}
        ${COUT('Vehicle Cost',F(o.vehicleCost),'tractorCost + trailerCost')}
        ${COUT('Accessories Cost',F(o.accessoriesCost),'sum of accessories')}
        ${COUT('Net Capital Cost',F(o.vehicleCost+o.accessoriesCost),'tractorCost + trailerCost + accessoriesCost (display only)')}
        ${INP('gvw','GVW (kg) — used for insurance addl. cover',i.gvw,'number','','100')}
      </div>
    </div>
  </div>
  <div class="g2 mb4">
    <div class="card"><div class="ch"><div class="ct">Depreciation — Companies Act 2013 Sch. II (SLM)</div></div>
      <div class="fg">
        ${INP('usefulLife','Useful Life (Years)',i.usefulLife,'number','','1','1','20')}
        ${INP('salvagePct','Salvage Value (%)',i.salvagePct,'number','','0.1','0','50')}
        ${COUT('Annual Dep. Rate',((1-n(i.salvagePct,5)/100)/n(i.usefulLife,8)*100).toFixed(2)+'%','(1 − salvage%) ÷ usefulLife')}
        ${COUT('Annual Depreciation',F(o.annualDep),'netCapitalCost × depRateAnnual')}
        ${COUT('Salvage at Contract End',F(o.salvageValueAtContractEnd),'cost − annualDep × contractYears')}
      </div>
      <div class="tw mt3"><table><thead><tr><th>Year</th><th class="nr">Opening</th><th class="nr">Depreciation</th><th class="nr">Closing Value</th></tr></thead>
      <tbody>${o.depSchedule.map(r=>`<tr><td><span class="badge ba">Yr ${r.year}</span></td><td class="nr">${F(r.opening)}</td><td class="nr">${F(r.dep)}</td><td class="nr">${F(r.closing)}</td></tr>`).join('')}</tbody></table></div>
    </div>
    <div class="card"><div class="ch"><div class="ct">Loan &amp; Capital Structure</div></div>
      <div class="fg">
        ${INP('loanPct','Loan Capital (%)',i.loanPct,'number','','0.1','0','100')}
        ${INP('loanInterestRate','Loan Interest Rate (% p.a.)',i.loanInterestRate,'number','Bank rate','0.001','0','30')}
        ${INP('ownCapRate','Own Capital Rate (% p.a.)',i.ownCapRate,'number','PF rate','0.001','0','20')}
        ${COUT('Loan Amount',F(o.loanAmount),'netCapitalCost × loanPct%')}
        ${COUT('Own Capital',F(o.ownCapAmount),'netCapitalCost × (1−loanPct%)')}
        ${COUT('Monthly Loan Instalment',F(o.loanInstalmentPM),'PMT((1+loanRate)^(1/12)-1, contractMonths, −loanAmt)')}
        ${COUT('Net Own Cap Amort/Month',F(o.netOwnCapAmortPM),'ownCapAmort − salvageAmort')}
        ${COUT('Total Capital/Month',F(o.totalCapitalPM),'loanInstalment + netOwnCapAmort')}
      </div>
    </div>
  </div>`;
},

labour(i,o){const op=o.opRows, hlp=o.hlpRows, sup=o.supRows, sy=n(i.contractStartYear,2026);const op1=op[0], hlp1=hlp[0], sup1=sup[0];const numOp=n(i.numOperators,1), numHlp=n(i.numHelpers,2), supShare=n(i.supervisionTTUShare,4);const opMonthlyLaborCost=op1.monthlyInclPFESInumOp+op1.relieverMonthlynumOp;const hlpMonthlyLaborCost=hlp1.monthlyInclPFESInumHlp+hlp1.relieverMonthlynumHlp;const supMonthlyLaborCost=sup1.monthlyShare+(sup1.monthlyPF+sup1.monthlyESI)/supShare;const yrHead=arr=>arr.map((_,idx)=><th class="nr">Yr ${idx+1} (${sy+idx})</th>).join('');return `${rbar()}

  <div class="sh mb4"><div><div class="st">Labour / Wages</div><div class="ss">Daily-rate wage model — Basic, VDA, Special Allowance, Bonus, OT, PF &amp; ESI/EC Insurance, exactly per Wages_OT_PF_ESI sheet</div></div></div>
  <div class="g4 mb4">
    ${KPI('Monthly Min. Wage — Operator',F(o.minMonthlyWageOp),'var(--ac)','excl. PF & ESI','minMonthlyWageOp')}
    ${KPI('Daily Min. Wage — Operator',F2(o.minDailyWageOp),'var(--ac)','excl. PF & ESI')}
    ${KPI('Monthly Min. Wage — Helper',F(o.minMonthlyWageHlp),'#0da6e0','excl. PF & ESI','minMonthlyWageHlp')}
    ${KPI('Daily Min. Wage — Helper',F2(o.minDailyWageHlp),'#0da6e0','excl. PF & ESI')}
  </div>
  <div class="g4 mb4">
    ${KPI('Total Wages Yr1/Month',F(o.wagesPM[0]),'var(--ac)','excl. PF & ESI, incl. relievers')}
    ${KPI('Employer Contribution/Month',F(o.statutoryPM[0]),'#f97316','PF + ESI/Insurance, whole crew','statutoryPM')}
    ${KPI('Labour Cost incl. Statutory/Month',F(o.laborCostInclStatutoryPM[0]),'#1fb97a','wages+supervision+PF/ESI','laborCostInclStatutoryPM')}
    ${KPI('Labour Cost / KM',F2(o.laborCostPerKM),'#9775ea','avg. monthly ÷ km run/month','laborCostPerKM')}
  </div>
  <div class="g4 mb4">
    ${KPI('Supervision Yr1/Month',F(o.supervisionPM[0]),'#0da6e0','1/'+i.supervisionTTUShare+' share')}
    ${KPI('Avg Monthly Wages',F(o.avgMonthlyWages),'#1fb97a','across contract')}
    ${KPI('Annual Labour Cost (Yr1)',F(o.laborCostAnnual[0]),'#e0507a','incl. statutory','laborCostAnnual')}
    ${KPI('Total OT (contract)',F(o.totalOTContract),'#9775ea','operator + helper')}
  </div>

  <div class="g3 mb4">
    <div class="card"><div class="ch"><div class="ct">Operator (Highly Skilled) — Daily Rate Inputs</div></div>
      <div class="fg">
        ${RINP('opBasicBase','Basic/day (2017 base → escalated)',i.opBasicBase)}
        ${RINP('opVDABase','VDA/day (base)',i.opVDABase)}
        ${RINP('opSpecialAllowance','Special Allowance/day',i.opSpecialAllowance)}
        ${INP('opBonusPct','Bonus % of (Basic+VDA)',i.opBonusPct,'number','','0.01','0','20','%')}
        ${INP('opEscalationPct','Annual Escalation % (Basic+VDA)',i.opEscalationPct,'number','','0.01','0','10','%')}
        ${INP('numOperators','No. of Operators',i.numOperators,'number','','1','1')}
      </div>
      <div class="ct" style="margin-top:14px;padding-top:10px;border-top:1px solid var(--bdr);font-size:12px">Live Calculated Output — Year 1</div>
      ${mrow('Basic/Day',F2(op1.basic))}
      ${mrow('VDA/Day',F2(op1.vda))}
      ${mrow('Bonus/Day',F2(op1.bonus))}
      ${mrow('Daily Rate (excl. PF &amp; ESI)',F2(op1.dailyRate))}
      ${mrow('Monthly Wage (excl. PF &amp; ESI)',F(op1.monthlyRate))}
      ${mrow('Single OT Rate/hr',F2(op1.singleOT))}
      ${mrow('Double OT Rate/hr',F2(op1.doubleOT))}
      ${mrow('Holiday OT Rate/hr',F2(op1.holidayOT))}
      ${mrow('Monthly PF',F(op1.monthlyPF))}
      ${op1.mode==='ESI'?mrow('Monthly ESI <span style="color:var(--gr);font-size:9px">(wage &lt; ₹'+n(i.esiThreshold,21000).toLocaleString('en-IN')+')</span>',F(op1.monthlyESI))
                        :mrow('Monthly EC Insurance <span style="color:var(--ac);font-size:9px">(wage ≥ ₹'+n(i.esiThreshold,21000).toLocaleString('en-IN')+')</span>',F(op1.monthlyInsurance))}
      ${mrow('Total Monthly Wage (incl. statutory)',F(op1.monthlyInclPFESI))}
      ${mrow('Reliever Wage/Month',F(op1.relieverMonthly))}
      ${mrow('Annual OT Amount (contract)',F(op1.otAnnual))}
      ${mrow('Employer Contribution/Month',F(op1.monthlyPF+op1.monthlyESIorEC))}
      ${mrow('<strong>Monthly Labour Cost</strong> (×'+numOp+')',F(opMonthlyLaborCost))}
      ${mrow('<strong>Annual Labour Cost</strong>',F(opMonthlyLaborCost*12))}
    </div>
    <div class="card"><div class="ch"><div class="ct">Helper (Semi-Skilled) × 2 — Daily Rate Inputs</div></div>
      <div class="fg">
        ${RINP('hlpBasicBase','Basic/day (base)',i.hlpBasicBase)}
        ${RINP('hlpVDABase','VDA/day (base)',i.hlpVDABase)}
        ${RINP('hlpSpecialAllowance','Special Allowance/day',i.hlpSpecialAllowance)}
        ${INP('hlpBonusPct','Bonus % of (Basic+VDA)',i.hlpBonusPct,'number','','0.01','0','20','%')}
        ${INP('hlpEscalationPct','Annual Escalation % (Basic+VDA)',i.hlpEscalationPct,'number','','0.01','0','10','%')}
        ${INP('numHelpers','No. of Helpers',i.numHelpers,'number','','1','1')}
      </div>
      <div class="ct" style="margin-top:14px;padding-top:10px;border-top:1px solid var(--bdr);font-size:12px">Live Calculated Output — Year 1 (per helper)</div>
      ${mrow('Daily Wage (excl. PF &amp; ESI)',F2(hlp1.dailyRate))}
      ${mrow('Monthly Wage (excl. PF &amp; ESI)',F(hlp1.monthlyRate))}
      ${mrow('Bonus/Day',F2(hlp1.bonus))}
      ${mrow('Monthly PF',F(hlp1.monthlyPF))}
      ${hlp1.mode==='ESI'?mrow('Monthly ESI <span style="color:var(--gr);font-size:9px">(wage &lt; ₹'+n(i.esiThreshold,21000).toLocaleString('en-IN')+')</span>',F(hlp1.monthlyESI))
                         :mrow('Monthly EC Insurance <span style="color:var(--ac);font-size:9px">(wage ≥ ₹'+n(i.esiThreshold,21000).toLocaleString('en-IN')+')</span>',F(hlp1.monthlyInsurance))}
      ${mrow('Employer Contribution/Month',F(hlp1.monthlyPF+hlp1.monthlyESIorEC))}
      ${mrow('OT/Month (annualised ÷12)',F(hlp1.otAnnual/12))}
      ${mrow('Reliever Cost/Month',F(hlp1.relieverMonthly))}
      ${mrow('<strong>Monthly Labour Cost</strong> (×'+numHlp+')',F(hlpMonthlyLaborCost))}
      ${mrow('<strong>Annual Labour Cost</strong>',F(hlpMonthlyLaborCost*12))}
    </div>
    <div class="card"><div class="ch"><div class="ct">Supervisor (Skilled) — Shared Across TTUs</div></div>
      <div class="fg">
        ${RINP('supBasicBase','Basic/day (base)',i.supBasicBase)}
        ${RINP('supVDABase','VDA/day (base)',i.supVDABase)}
        ${INP('supBonusPct','Bonus % of (Basic+VDA)',i.supBonusPct,'number','','0.01','0','20','%')}
        ${INP('supEscalationPct','Annual Escalation % (Basic+VDA)',i.supEscalationPct,'number','','0.01','0','10','%')}
        ${INP('supervisionTTUShare','Shared across N TTUs',i.supervisionTTUShare,'number','','1','1','10')}
      </div>
      <div class="ct" style="margin-top:14px;padding-top:10px;border-top:1px solid var(--bdr);font-size:12px">Live Calculated Output — Year 1 (1/${supShare} share)</div>
      ${mrow('Daily Wage (full, pre-share)',F2(sup1.dailyRate))}
      ${mrow('Monthly Share',F(sup1.monthlyShare))}
      ${mrow('Bonus/Day',F2(sup1.bonus))}
      ${mrow('Monthly PF (share)',F(sup1.monthlyPF/supShare))}
      ${mrow('Monthly ESI (share) <span style="color:var(--gr);font-size:9px">(Excel: supervisor is always ESI)</span>',F(sup1.monthlyESI/supShare))}
      ${mrow('Employer Contribution/Month',F((sup1.monthlyPF+sup1.monthlyESI)/supShare))}
      ${mrow('Shared Cost per TTU',F(sup1.monthlyShare))}
      ${mrow('<strong>Monthly Cost</strong>',F(supMonthlyLaborCost))}
      ${mrow('<strong>Annual Cost</strong>',F(supMonthlyLaborCost*12))}
    </div>
  </div>

  <div class="card mb4"><div class="ch"><div class="ct">OT, Reliever &amp; Bonus Escalation Settings</div></div>
    <div class="fg">
      ${INP('singleOTHrs','Single OT Hours/day (8th–9th hr)',i.singleOTHrs,'number','','0.5','0','4')}
      ${INP('doubleOTHrs','Double OT Hours/day (9th–12th hr)',i.doubleOTHrs,'number','','0.5','0','8')}
      ${INP('bonusEscalationPct','Bonus Escalation % p.a. (Helper)',i.bonusEscalationPct,'number','','0.01','0','15','%')}
      ${INP('relieverDaysFactor','Reliever Days Factor/Year',i.relieverDaysFactor,'number','','1','0','365')}
    </div>
    <div class="callout mt3">Crew: 1 Operator (Highly Skilled) + 1 Relief Operator + 2 Helpers (Semi-Skilled) + 2 Relief Helpers + 1/${i.supervisionTTUShare} Supervisor (Skilled).
OT contract-cost formula intentionally replicates the Excel quirk: the "+2×DoubleOT" uplift term is absent in the Operator's first contract year but present from Year 2 onward; for Helpers it is present from Year 1. This is carried over verbatim from Wages_OT_PF_ESI.</div>
  </div>

  <div class="card mb4"><div class="ch"><div class="ct">Year-wise Daily &amp; Monthly Rates</div></div>
    <div class="tw"><table><thead><tr><th>Component</th>${yrHead(op)}</tr></thead><tbody>
      <tr><td colspan="${op.length+1}" style="background:var(--surf);font-weight:700;color:var(--ac)">OPERATOR</td></tr>
      <tr><td>Basic/day</td>${op.map(r=>`<td class="nr">${r.basic.toFixed(2)}</td>`).join('')}</tr>
      <tr><td>VDA/day</td>${op.map(r=>`<td class="nr">${r.vda.toFixed(2)}</td>`).join('')}</tr>
      <tr><td>Bonus/day</td>${op.map(r=>`<td class="nr">${r.bonus.toFixed(2)}</td>`).join('')}</tr>
      <tr><td>Daily Rate (excl PF/ESI)</td>${op.map(r=>`<td class="nr">${r.dailyRate.toFixed(2)}</td>`).join('')}</tr>
      <tr><td>Monthly Rate (excl PF/ESI)</td>${op.map(r=>`<td class="nr">${F(r.monthlyRate)}</td>`).join('')}</tr>
      <tr><td>Single OT/hr</td>${op.map(r=>`<td class="nr">${r.singleOT.toFixed(2)}</td>`).join('')}</tr>
      <tr><td>Double OT/hr</td>${op.map(r=>`<td class="nr">${r.doubleOT.toFixed(2)}</td>`).join('')}</tr>
      <tr><td>Holiday OT/hr</td>${op.map(r=>`<td class="nr">${r.holidayOT.toFixed(2)}</td>`).join('')}</tr>
      <tr><td>PF Contribution/day</td>${op.map(r=>`<td class="nr">${r.pfDaily.toFixed(2)}</td>`).join('')}</tr>
      <tr><td>Monthly PF Contribution</td>${op.map(r=>`<td class="nr">${F(r.monthlyPF)}</td>`).join('')}</tr>
      <tr><td>ESI or Insurance/day</td>${op.map(r=>`<td class="nr">${r.esiOrEcDaily.toFixed(2)} <span style="color:var(--mu);font-size:9px">(${r.mode})</span></td>`).join('')}</tr>
      <tr><td>Monthly ESI Contribution</td>${op.map(r=>`<td class="nr">${r.mode==='ESI'?F(r.monthlyESI):'—'}</td>`).join('')}</tr>
      <tr><td>Monthly Insurance Amount (EC)</td>${op.map(r=>`<td class="nr">${r.mode==='EC'?F(r.monthlyInsurance):'—'}</td>`).join('')}</tr>
      <tr><td style="color:var(--mu)">Employee's PF Contribution/Month <span style="font-size:9px">(informational)</span></td>${op.map(r=>`<td class="nr" style="color:var(--mu)">${F(r.employeePfMonthly)}</td>`).join('')}</tr>
      <tr><td>Total Monthly incl. PF &amp; ESI</td>${op.map(r=>`<td class="nr">${F(r.monthlyInclPFESI)}</td>`).join('')}</tr>
      <tr><td>Reliever Wage/Month</td>${op.map(r=>`<td class="nr">${F(r.relieverMonthly)}</td>`).join('')}</tr>
      <tr><td>OT Amount P.A. (contract)</td>${op.map(r=>`<td class="nr">${F(r.otAnnual)}</td>`).join('')}</tr>

  <tr><td colspan="${hlp.length+1}" style="background:var(--surf);font-weight:700;color:var(--ac)">HELPER (each, × ${i.numHelpers})</td></tr>
  <tr><td>Basic/day</td>${hlp.map(r=>`<td class="nr">${r.basic.toFixed(2)}</td>`).join('')}</tr>
  <tr><td>VDA/day</td>${hlp.map(r=>`<td class="nr">${r.vda.toFixed(2)}</td>`).join('')}</tr>
  <tr><td>Bonus/day</td>${hlp.map(r=>`<td class="nr">${r.bonus.toFixed(2)}</td>`).join('')}</tr>
  <tr><td>Daily Rate (excl PF/ESI)</td>${hlp.map(r=>`<td class="nr">${r.dailyRate.toFixed(2)}</td>`).join('')}</tr>
  <tr><td>Monthly Rate (excl PF/ESI)</td>${hlp.map(r=>`<td class="nr">${F(r.monthlyRate)}</td>`).join('')}</tr>
  <tr><td>PF Contribution/day</td>${hlp.map(r=>`<td class="nr">${r.pfDaily.toFixed(2)}</td>`).join('')}</tr>
  <tr><td>Monthly PF Contribution</td>${hlp.map(r=>`<td class="nr">${F(r.monthlyPF)}</td>`).join('')}</tr>
  <tr><td>ESI or Insurance/day</td>${hlp.map(r=>`<td class="nr">${r.esiOrEcDaily.toFixed(2)} <span style="color:var(--mu);font-size:9px">(${r.mode})</span></td>`).join('')}</tr>
  <tr><td>Monthly ESI Contribution</td>${hlp.map(r=>`<td class="nr">${r.mode==='ESI'?F(r.monthlyESI):'—'}</td>`).join('')}</tr>
  <tr><td>Monthly Insurance Amount (EC)</td>${hlp.map(r=>`<td class="nr">${r.mode==='EC'?F(r.monthlyInsurance):'—'}</td>`).join('')}</tr>
  <tr><td style="color:var(--mu)">Employee's PF Contribution/Month <span style="font-size:9px">(informational)</span></td>${hlp.map(r=>`<td class="nr" style="color:var(--mu)">${F(r.employeePfMonthly)}</td>`).join('')}</tr>
  <tr><td>Reliever Wage/Month</td>${hlp.map(r=>`<td class="nr">${F(r.relieverMonthly)}</td>`).join('')}</tr>
  <tr><td>OT Amount P.A. (contract, per helper)</td>${hlp.map(r=>`<td class="nr">${F(r.otAnnual)}</td>`).join('')}</tr>

  <tr><td colspan="${sup.length+1}" style="background:var(--surf);font-weight:700;color:var(--ac)">SUPERVISOR (1/${i.supervisionTTUShare} share)</td></tr>
  <tr><td>Basic/day</td>${sup.map(r=>`<td class="nr">${r.basic.toFixed(2)}</td>`).join('')}</tr>
  <tr><td>VDA/day</td>${sup.map(r=>`<td class="nr">${r.vda.toFixed(2)}</td>`).join('')}</tr>
  <tr><td>Bonus/day</td>${sup.map(r=>`<td class="nr">${r.bonus.toFixed(2)}</td>`).join('')}</tr>
  <tr><td>Monthly PF Contribution (share)</td>${sup.map(r=>`<td class="nr">${F(r.monthlyPF/i.supervisionTTUShare)}</td>`).join('')}</tr>
  <tr><td>Monthly ESI Contribution (share)</td>${sup.map(r=>`<td class="nr">${F(r.monthlyESI/i.supervisionTTUShare)}</td>`).join('')}</tr>
  <tr><td style="color:var(--mu)">Employee's PF Contribution/Month <span style="font-size:9px">(informational, share)</span></td>${sup.map(r=>`<td class="nr" style="color:var(--mu)">${F(r.employeePfMonthly/i.supervisionTTUShare)}</td>`).join('')}</tr>
  <tr><td>Monthly Share/TTU</td>${sup.map(r=>`<td class="nr">${F(r.monthlyShare)}</td>`).join('')}</tr>

  <tr class="ttr"><td><strong>TOTAL WAGES/MONTH (Op+RelOp+2×Hlp+2×RelHlp)</strong></td>${o.wagesPM.map(w=>`<td class="nr">${F(w)}</td>`).join('')}</tr>
  <tr class="ttr"><td><strong>TOTAL PF/MONTH (whole crew)</strong></td>${o.pfPM.map(w=>`<td class="nr">${F(w)}</td>`).join('')}</tr>
  <tr class="ttr"><td><strong>TOTAL ESI/MONTH (whole crew)</strong></td>${o.esiPM.map(w=>`<td class="nr">${F(w)}</td>`).join('')}</tr>
  <tr class="ttr"><td><strong>TOTAL INSURANCE/MONTH (whole crew)</strong></td>${o.insurancePM.map(w=>`<td class="nr">${F(w)}</td>`).join('')}</tr>
  <tr class="ttr"><td><strong>LABOUR COST/MONTH incl. STATUTORY</strong></td>${o.laborCostInclStatutoryPM.map(w=>`<td class="nr">${F(w)}</td>`).join('')}</tr>
  <tr class="ttr"><td><strong>LABOUR COST/DAY incl. STATUTORY</strong></td>${o.dailyLaborCost.map(w=>`<td class="nr">${F(w)}</td>`).join('')}</tr>
  <tr class="ttr"><td><strong>ANNUAL LABOUR COST</strong></td>${o.laborCostAnnual.map(w=>`<td class="nr">${F(w)}</td>`).join('')}</tr>
</tbody></table></div>

  </div>

  <div class="card mb4"><div class="ch"><div class="ct">PF / ESI / EC Insurance — Decision Engine <span class="badge ba" style="margin-left:6px">Live — recalculates on every keystroke</span></div></div>
    <div class="g2">
      ${decisionPanel('Operator',op[0],i.esiThreshold,i.esiRatePct,i.pfRatePct,i.pfDailyCap,i.ecInsuranceAnnual)}
      ${decisionPanel('Helper (each)',hlp[0],i.esiThreshold,i.esiRatePct,i.pfRatePct,i.pfDailyCap,i.ecInsuranceAnnual)}
    </div>
  </div>

  <div class="g2 mb4">
    <div class="card"><div class="ch"><div class="ct">PF / ESI / EC Insurance Rules</div></div>
      <div class="fg">
        ${INP('pfRatePct','PF Rate (incl. Admin & EDLI) %',i.pfRatePct,'number','','0.01','0','15','%')}
        ${INP('pfDailyCap','PF Cap (₹/day)',i.pfDailyCap,'number','','1','0')}
        ${INP('esiRatePct','ESI Rate %',i.esiRatePct,'number','','0.01','0','5','%')}
        ${INP('esiThreshold','ESI Threshold (₹/month)',i.esiThreshold,'number','','100','0')}
        ${INP('ecInsuranceAnnual','EC Insurance (₹/year, if ≥ threshold)',i.ecInsuranceAnnual,'number','','1','0')}
        ${INP('employeePfRatePct',"Employee's PF Rate % (informational)",i.employeePfRatePct,'number','','0.01','0','15','%')}
      </div>
      <div class="callout mt3"><strong>Automatic rule:</strong> If Monthly Wage &lt; ₹${n(i.esiThreshold,21000).toLocaleString('en-IN')} → PF + <strong>ESI</strong> @ ${i.esiRatePct}% of (Basic+VDA) apply (Insurance hidden).
If Monthly Wage ≥ ₹${n(i.esiThreshold,21000).toLocaleString('en-IN')} → PF + <strong>Employees' Compensation Insurance</strong> @ ₹${n(i.ecInsuranceAnnual,2738).toLocaleString('en-IN')}/year apply instead (ESI hidden).
PF (Employer) @ ${i.pfRatePct}% of (Basic+VDA), capped at ₹${i.pfDailyCap}/day — billed to the client as part of Labour Cost.
Employee's own PF @ ${i.employeePfRatePct}% is shown for compliance reference only; it is deducted from the worker's own wage and is NOT added to the contractor's billed cost.</div>
    </div>
    <div class="card"><div class="ch"><div class="ct">Uniform, Shoe &amp; Medical Fitness</div></div>
      <div class="fg">
        ${RINP('uniformShoePerHeadPerYear','Uniform &amp; Shoe / Head / Year',i.uniformShoePerHeadPerYear)}
        ${INP('numUniformPersons','No. of Persons (Uniform)',i.numUniformPersons,'number','','1','1')}
        ${RINP('medicalCertPerHeadPerYear','Medical Cert. / Head / Year',i.medicalCertPerHeadPerYear)}
        ${INP('numMedicalPersons','No. of Persons (Medical)',i.numMedicalPersons,'number','','1','1')}
        ${INP('medicalCertYearMultiplier','Medical Cert. Year Multiplier',i.medicalCertYearMultiplier,'number','Excel literal = 2, not contract years','1','1','10')}
        ${COUT('Uniform & Shoe Total (Contract)',F(o.uniformShoeTotal),'perHead × persons × contractYears')}
        ${COUT('Medical Cert. Total (Contract)',F(o.medicalCertTotal),'perHead × persons × yearMultiplier')}
      </div>
      <div class="callout mt3">⚠️ Excel quirk preserved: Annexure-III cell F123 literally computes Medical Fitness Certificate as 3000×3×2 (=₹18,000), NOT ×4 years (₹36,000) as the general spec describes. The "Medical Cert. Year Multiplier" input defaults to 2 to match the Excel exactly — raise it to 4 if the correct multi-year intent is required.</div>
    </div>
  </div>`;
},

fuel(i,o){return `${rbar()}

  <div class="sh mb4"><div><div class="st">Fuel &amp; Lubricants</div><div class="ss">Diesel, DEF (AdBlue) and Lube Oil — per Annexure-A running charges section</div></div></div>
  <div class="g4 mb4">
    ${KPI('Diesel/Month',F(o.dieselCostPM),'var(--ac)',o.fuelLitPerMonth.toFixed(0)+' litres')}
    ${KPI('DEF/Month',F(o.defCostPM),'#0da6e0',o.defLitPM.toFixed(1)+' litres')}
    ${KPI('Lube Oil/Month',F(o.lubeCostPM),'#9775ea','0.5% of diesel cost')}
    ${KPI('Total Fuel/Month',F(o.totalFuelPM),'#1fb97a','₹ '+o.totalFuelPerKm.toFixed(2)+'/km')}
  </div>
  <div class="g2 mb4">
    <div class="card"><div class="ch"><div class="ct">Diesel</div></div>
      <div class="fg">
        ${RINP('dieselPrice','Diesel Price at Duliajan (₹/L)',i.dieselPrice)}
        ${INP('fuelEfficiency','Fuel Efficiency (KMPL)',i.fuelEfficiency,'number','km per litre','0.01','0.5')}
        ${COUT('Litres / Month',(o.fuelLitPerMonth).toFixed(1),'kmPerMonth ÷ fuelEfficiency','L/month')}
        ${COUT('Diesel Cost / Month',F(o.dieselCostPM),'litresPerMonth × dieselPrice')}
        ${COUT('Diesel / KM',F2(o.dieselPerKm),'dieselPrice ÷ fuelEfficiency')}
      </div>
    </div>
    <div class="card"><div class="ch"><div class="ct">DEF &amp; Lube Oil</div></div>
      <div class="fg">
        ${RINP('defPriceLow','DEF Quote 1 (₹/L)',i.defPriceLow)}
        ${RINP('defPriceHigh','DEF Quote 2 (₹/L)',i.defPriceHigh)}
        ${COUT('DEF Price (avg)',F2(o.defPrice),'(low + high) ÷ 2')}
        ${INP('defConsumptionPctOfFuel','DEF Consumption (% of fuel vol.)',i.defConsumptionPctOfFuel,'number','3% typical','0.01','0','20')}
        ${COUT('DEF Litres / Month',o.defLitPM.toFixed(2),'fuelLitres × defConsumptionPct%','L/month')}
        ${COUT('DEF Cost / Month',F(o.defCostPM),'defLitres × defPrice')}
        ${INP('lubeOilPct','Lube Oil Rate (% of diesel/km)',i.lubeOilPct,'number','0.5% typical','0.01','0','5')}
        ${COUT('Lube Cost / Month',F(o.lubeCostPM),'dieselPerKm × lubeOilPct% × kmPM')}
      </div>
    </div>
  </div>
  <div class="card"><div class="ch"><div class="ct">Fuel — Per KM breakdown (from Annexure formula)</div></div>
    <div class="tw"><table><thead><tr><th>Component</th><th>Formula (Annexure ref.)</th><th class="nr">₹ per KM</th><th class="nr">Monthly (₹)</th></tr></thead>
    <tbody>
      <tr><td>Diesel</td><td><span class="fchip">dieselPrice ÷ fuelEfficiency</span></td><td class="nr">${o.dieselPerKm.toFixed(2)}</td><td class="nr">${F(o.dieselCostPM)}</td></tr>
      <tr><td>DEF (${i.defConsumptionPctOfFuel}% of fuel vol)</td><td><span class="fchip">defPrice × (1/fuelEff) × defPct</span></td><td class="nr">${o.defPerKm.toFixed(2)}</td><td class="nr">${F(o.defCostPM)}</td></tr>
      <tr><td>Lube Oil (${i.lubeOilPct}% of diesel/km)</td><td><span class="fchip">dieselPerKm × lubeOilPct%</span></td><td class="nr">${o.lubePerKm.toFixed(2)}</td><td class="nr">${F(o.lubeCostPM)}</td></tr>
      <tr class="ttr"><td><strong>Total Fuel</strong></td><td></td><td class="nr">${o.totalFuelPerKm.toFixed(2)}</td><td class="nr">${F(o.totalFuelPM)}</td></tr>
    </tbody></table></div>
  </div>`;
},

tyre(i,o){return `${rbar()}

  <div class="sh mb4"><div><div class="st">Tyre &amp; Battery</div><div class="ss">Per Annexure-D tyre calculation — excl. 18% GST</div></div></div>
  <div class="g4 mb4">
    ${KPI('Cost/Tyre (excl. GST)',F(n(i.tyreCostExclGST)),'var(--ac)','avg per Annexure-D')}
    ${KPI('No. of Tyres',n(i.numTyres).toString(),'#0da6e0','excl. spare wheel')}
    ${KPI('Tyre Cost/Month',F(o.tyreCostPM),'#1fb97a',n(i.tyreLife).toLocaleString('en-IN')+' KM life')}
    ${KPI('Battery Cost/Month',F(o.batteryCostPM),'#9775ea',i.batteryLife+' yr life')}
  </div>
  <div class="g2 mb4">
    <div class="card"><div class="ch"><div class="ct">Tyre Parameters (Annexure-D)</div></div>
      <div class="fg">
        ${RINP('tyreCostExclGST','Avg Tyre Cost (excl. 18% GST)',i.tyreCostExclGST)}
        ${INP('numTyres','Number of Tyres (excl. spare)',i.numTyres,'number','','1','1')}
        ${INP('tyreLife','Life of Tyre &amp; Tube (KM)',i.tyreLife,'number','','1000','1000')}
        ${INP('tyreSpareAdj','Spare-Tyre Adjustment (sets subtracted)',i.tyreSpareAdj,'number','Excel quirk: 1 spare on purchase','1','0','2')}
        ${COUT('Total Set Cost (excl. GST)',F(n(i.tyreCostExclGST)*n(i.numTyres)),'tyreCost × numTyres')}
        ${COUT('Tyre Changes in Contract',(o.tyreChangesInContract).toFixed(2),'(contractKM ÷ tyreLife) − spareAdj')}
        ${COUT('Tyre Cost / Month',F(o.tyreCostPM),'tyreCost × changes × numTyres ÷ contractMonths')}
        ${COUT('Tyre Cost / KM',F2(o.tyrePerKm),'per Annexure formula')}
      </div>
    </div>
    <div class="card"><div class="ch"><div class="ct">Battery Parameters</div></div>
      <div class="fg">
        ${RINP('batteryUnitPrice','Battery Unit Price (₹, incl. GST)',i.batteryUnitPrice)}
        ${INP('batteryGSTFactor','GST Factor (1+GST%)',i.batteryGSTFactor,'number','','0.001','1','2')}
        ${INP('batteryLife','Battery Life (Years)',i.batteryLife,'number','','0.5','0.5')}
        ${INP('batteryEscalationPct','Battery Cost Escalation % (for per-km rate)',i.batteryEscalationPct,'number','','0.01','0','10')}
        ${COUT('Battery Set Cost (display)',F(o.batterySetCostDisplay),'unitPrice + unitPrice÷GSTFactor')}
        ${COUT('Battery Cost / Month',F(o.batteryCostPM),'batteryPerKm × kmPerMonth')}
        ${COUT('Battery Cost / KM',F2(o.batteryPerKm),'per Interest_Maint_Insurance!F99')}
      </div>
      <div class="callout mt3">Battery set cost basis (excl. GST) = unitPrice ÷ GSTFactor, multiplied by battery life in years, then escalated by batteryEscalationPct compounded over the life — replicated exactly from the Excel per-km formula (F99), including this "life-year multiplier" quirk.</div>
    </div>
  </div>`;
},

maintenance(i,o){const cap=o.netCapitalCost;const mPcts=[n(i.maintPct_yr1),n(i.maintPct_yr2),n(i.maintPct_yr3),n(i.maintPct_yr4)];const sy=n(i.contractStartYear,2026);return `${rbar()}

  <div class="sh mb4"><div><div class="st">Maintenance</div><div class="ss">Year-wise % of Capital Cost, PV-discounted and PMT-amortised monthly — per Annexure-A row 23 / Interest_Maint_Insurance sheet</div></div></div>
  <div class="maint-yr mb4">
    ${o.maintYr.map((m,idx)=>`<div class="yr-box">
      <div class="yr-lbl">Year ${idx+1} (${sy+idx})</div>
      <div class="yr-v">${F(m)}</div>
      <div class="yr-s">${mPcts[idx]!==undefined?mPcts[idx]:mPcts[3]}% of capital cost</div>
    </div>`).join('')}
  </div>
  <div class="g2 mb4">
    <div class="card"><div class="ch"><div class="ct">Maintenance % Inputs (Annexure-A row 23)</div></div>
      <div class="fg">
        ${INP('maintPct_yr1','Year 1 — % of Capital Cost',i.maintPct_yr1,'number','1% per Annexure','0.01','0','20','%')}
        ${INP('maintPct_yr2','Year 2 — % of Capital Cost',i.maintPct_yr2,'number','2% per Annexure','0.01','0','20','%')}
        ${INP('maintPct_yr3','Year 3 — % of Capital Cost',i.maintPct_yr3,'number','3% per Annexure','0.01','0','20','%')}
        ${INP('maintPct_yr4','Year 4 — % of Capital Cost',i.maintPct_yr4,'number','4% per Annexure','0.01','0','20','%')}
        ${COUT('Net Capital Cost (Base)',F(cap),'tractor + trailer + accessories')}
        ${COUT('Avg Monthly Maintenance',F(o.avgMonthlyMaint),'avg of PMT-amortised yr-wise PV')}
        ${COUT('Avg Maintenance / KM',F2(o.maintPerKm),'avgMonthlyMaint ÷ kmPerMonth')}
      </div>
    </div>
    <div class="card"><div class="ch"><div class="ct">Year-wise Maintenance Cost</div></div>
      <div class="tw"><table><thead><tr><th>Year</th><th class="nr">Capital Base</th><th>Rate</th><th class="nr">Annual Cost</th><th class="nr">Monthly Amort.</th></tr></thead>
      <tbody>
        ${o.maintYr.map((m,idx)=>`<tr><td><span class="badge ba">Year ${idx+1} (${sy+idx})</span></td><td class="nr">${F(cap)}</td><td class="nr">${mPcts[idx]!==undefined?mPcts[idx]:mPcts[3]}%</td><td class="nr">${F(m)}</td><td class="nr">${F(o.maintAmortPM[idx])}</td></tr>`).join('')}
        <tr class="ttr"><td><strong>Total / Average</strong></td><td></td><td></td><td class="nr">${F(o.maintYr.reduce((s,m)=>s+m,0))}</td><td class="nr">${F(o.avgMonthlyMaint)}</td></tr>
      </tbody></table></div>
      <div class="callout mt3">Each year's maintenance is PV-discounted at the Own Capital (PF) rate to the start of the contract, then amortised over 12 months via PMT — exactly as computed in Interest_Maint_Insurance rows 24-40.</div>
    </div>
  </div>`;
},

insurance(i,o){const sy=n(i.contractStartYear,2026);const yrs=o.insuranceYears;return `${rbar()}

  <div class="sh mb4"><div><div class="st">Insurance</div><div class="ss">IDV-based year-wise declining premium, per Motor Insurance Tariff schedule (Interest_Maint_Insurance sheet) — depreciation band driven by vehicle age</div></div></div>
  <div class="g4 mb4">
    ${KPI('Avg Monthly (amortised)',F(o.avgMonthlyInsurance),'var(--ac)','used in fixed charge')}
    ${KPI('Year 1 Premium (incl GST)',F(yrs[0].totalIncGST),'#0da6e0','age 0, IDV dep '+yrs[0].depPct+'%')}
    ${KPI('Year '+yrs.length+' Premium (incl GST)',F(yrs[yrs.length-1].totalIncGST),'#1fb97a','IDV dep '+yrs[yrs.length-1].depPct+'%')}
    ${KPI('TPI (Liability)',F(n(i.tpiPremium,46435)),'#9775ea','flat, tractor+trailer')}
  </div>
  <div class="g2 mb4">
    <div class="card"><div class="ch"><div class="ct">Tariff Inputs</div></div>
      <div class="fg">
        ${INP('odRateUpto5yr','OD Premium Rate — upto 5 yrs (%)',i.odRateUpto5yr,'number','','0.001','0','5')}
        ${INP('odRateUpto7yr','OD Premium Rate — upto 7 yrs (%)',i.odRateUpto7yr,'number','','0.001','0','5')}
        ${INP('odRateAbove7yr','OD Premium Rate — above 7 yrs (%)',i.odRateAbove7yr,'number','','0.001','0','5')}
        ${RINP('liabilityTractor','Liability Only Cover — Tractor',i.liabilityTractor)}
        ${RINP('liabilityTrailer','Liability Only Cover — Trailer',i.liabilityTrailer)}
        ${INP('addlCoverRatePerKg','Addl. Cover Rate (₹ per kg over threshold)',i.addlCoverRatePerKg,'number','','0.001','0','2')}
        ${INP('addlCoverGVWThreshold','Addl. Cover GVW Threshold (kg)',i.addlCoverGVWThreshold,'number','','100','0')}
      </div>
    </div>
    <div class="card"><div class="ch"><div class="ct">Discounts &amp; Fixed Covers</div></div>
      <div class="fg">
        ${INP('specialDiscountPct','Special Discount on OD Premium (%)',i.specialDiscountPct,'number','20–50% typical','0.1','0','50')}
        ${RINP('tpiPremium','TPI Premium (₹/year, Tractor+Trailer)',i.tpiPremium)}
        ${RINP('paCoverPerPerson','PA Cover (₹/person)',i.paCoverPerPerson)}
        ${INP('gstPct','GST Rate (%)',i.gstPct,'number','','0.01','0','30')}
      </div>
      <div class="callout mt3">IDV Depreciation Schedule (age of vehicle at start of each contract year):
≤6mo: 5% · 6mo–1yr: 15% · 1–2yr: 20% · 2–3yr: 30% · 3–4yr: 40% · 4–5yr: 50%
No-Claim Bonus by contract year: Yr1: 0% · Yr2: 20% · Yr3: 25% · Yr4: 35%
For a brand-new TTU the 4 contract years correspond to vehicle ages 0,1,2,3 → IDV dep. 5%/20%/30%/40%, matching the specification exactly. The same engine drives the Vintage TTU dashboard using the vehicle's actual age.</div>
    </div>
  </div>
  <div class="card mb4"><div class="ch"><div class="ct">Year-wise Insurance Premium (computed live)</div></div>
    <div class="ins-yr-grid">
      ${yrs.map((y,idx)=>`
      <div class="ins-yr-box">
        <div class="iyr-lbl">Year ${idx+1} (${sy+idx})</div>
        <div class="iyr-v">${F(y.totalIncGST)}</div>
        <div class="iyr-s">Dep: ${y.depPct}% · NCB: ${y.ncbPct}%</div>
      </div>`).join('')}
    </div>
    <div style="margin-top:10px" class="tw"><table><thead><tr><th>Component</th>${yrs.map((_,idx)=>`<th class="nr">Year ${idx+1}</th>`).join('')}</tr></thead>
    <tbody>
      <tr><td>Tractor IDV</td>${yrs.map(y=>`<td class="nr">${F(y.tractorIDV)}</td>`).join('')}</tr>
      <tr><td>Trailer IDV</td>${yrs.map(y=>`<td class="nr">${F(y.trailerIDV)}</td>`).join('')}</tr>
      <tr><td>OD Rate Applied</td>${yrs.map(y=>`<td class="nr">${y.odRatePct}%</td>`).join('')}</tr>
      <tr><td>Sub-Total OD + Liability + Addl.</td>${yrs.map(y=>`<td class="nr">${F(y.subTotalOD)}</td>`).join('')}</tr>
      <tr><td>Net ODP (after NCB &amp; Discount)</td>${yrs.map(y=>`<td class="nr">${F(y.netODP)}</td>`).join('')}</tr>
      <tr><td>TPI (Liability)</td>${yrs.map(y=>`<td class="nr">${F(y.tpi)}</td>`).join('')}</tr>
      <tr><td>PA Cover</td>${yrs.map(y=>`<td class="nr">${F(y.pa)}</td>`).join('')}</tr>
      <tr><td>Total ex-GST</td>${yrs.map(y=>`<td class="nr">${F(y.totalExGST)}</td>`).join('')}</tr>
      <tr class="ttr"><td><strong>Total incl. GST</strong></td>${yrs.map(y=>`<td class="nr">${F(y.totalIncGST)}</td>`).join('')}</tr>
      <tr><td>Monthly Amortisation</td>${o.insuranceAmortPM.map(v=>`<td class="nr">${F(v)}</td>`).join('')}</tr>
    </tbody></table></div>
  </div>`;
},

tax(i,o){const sy=n(i.contractStartYear,2026);const rt=o.roadTaxYears;return `${rbar()}

  <div class="sh mb4"><div><div class="st">Road Tax &amp; Permit</div><div class="ss">Year-wise Road Tax, Permit Fee, Pollution Certificate — per Annexure-B, plus GST &amp; Toll (Annexure-C)</div></div></div>
  <div class="g4 mb4">
    ${KPI('Year 1 Total',F(rt[0].total),'var(--ac)','higher — registration year')}
    ${KPI('Year 2+ Total',F(rt[rt.length-1].total),'#0da6e0','renewal years')}
    ${KPI('Avg / Year (amortised)',F(o.avgRoadTaxPerYear),'#1fb97a','used in fixed charge')}
    ${KPI('Toll (Reimbursable)',F(o.tollPerYear),'#9775ea','per Annexure-C')}
  </div>
  <div class="g2 mb4">
    <div class="card"><div class="ch"><div class="ct">Road Tax &amp; Permit Inputs (Annexure-B)</div></div>
      <div class="fg">
        ${RINP('roadTaxOnly_yr1','Road Tax Only — Year 1',i.roadTaxOnly_yr1)}
        ${RINP('roadTaxOnly_yrRest','Road Tax Only — Year 2-4 (each)',i.roadTaxOnly_yrRest)}
        ${RINP('permitFee','Permit Fee',i.permitFee)}
        ${RINP('pollutionCert','Pollution Certificate',i.pollutionCert)}
        ${COUT('Average / Year',F(o.avgRoadTaxPerYear),'average of yr-wise totals')}
        ${COUT('Monthly (PMT-amortised)',F(o.roadTaxAmortPM),'PMT((1+ownRate)^(1/12)-1,12,−avgRoadTax)')}
      </div>
    </div>
    <div class="card"><div class="ch"><div class="ct">GST &amp; Toll</div></div>
      <div class="fg">
        ${INP('gstPct','GST Rate (%)',i.gstPct,'number','Applied on grand total','0.01','0','30')}
        ${RINP('tollPerTripAmount','Toll per Trip (₹)',i.tollPerTripAmount)}
        ${INP('tripsPerYear','Trips per Year',i.tripsPerYear,'number','','1','0')}
        ${COUT('Toll / Year / TTU',F(o.tollPerYear),'tollPerTrip × tripsPerYear')}
        ${COUT('Toll / Month (amortised)',F(o.tollPerYear/12),'tollPerYear ÷ 12')}
        ${COUT('Grand Total Contract (ex-GST)',F(o.grandTotal),'baseContractCost + reimbursables')}
        ${COUT('Grand Total (incl. GST)',F(o.grandTotalWithGST),'grandTotal × (1 + gstPct%)')}
      </div>
    </div>
  </div>
  <div class="card"><div class="ch"><div class="ct">Year-wise Road Tax Breakdown</div></div>
    <div class="tw"><table><thead><tr><th>Year</th><th class="nr">Road Tax</th><th class="nr">Permit</th><th class="nr">Pollution</th><th class="nr">Total</th></tr></thead>
    <tbody>${rt.map((r,idx)=>`<tr><td><span class="badge ba">Year ${idx+1} (${sy+idx})</span></td><td class="nr">${F(r.only)}</td><td class="nr">${F(r.permit)}</td><td class="nr">${F(r.pollution)}</td><td class="nr">${F(r.total)}</td></tr>`).join('')}
    <tr class="ttr"><td><strong>Average</strong></td><td></td><td></td><td></td><td class="nr">${F(o.avgRoadTaxPerYear)}</td></tr>
    </tbody></table></div>
    <div class="callout mt3">20MT defaults: Yr1 ₹32,065 (Road Tax ₹29,205 + Permit ₹2,600 + Pollution ₹260); Yr2-4 ₹29,285 each. Average ≈ ₹29,980/year.
30MT defaults: Yr1 ₹39,423; Yr2-4 ₹36,907 each. Average ≈ ₹37,536/year.
Toll: ₹1,070/trip × 12 trips = ₹12,840/TTU/year (same for 20MT &amp; 30MT) — reimbursable, excluded from the base contract cost.</div>
  </div>`;
},

accessories(i,o){const accs=State.accessories;return `${rbar()}

  <div class="sh mb4"><div><div class="st">Accessories</div><div class="ss">Optional fitments that add to Net Capital Cost and therefore feed loan, own-capital, maintenance and insurance calculations</div></div>
    <button class="btn bp sm" onclick="UI.addAccModal()">➕ Add Accessory</button>
  </div>
  <div class="g4 mb4">
    ${KPI('No. of Accessories',accs.length.toString(),'var(--ac)','')}
    ${KPI('Accessories Cost (incl GST)',F(o.accessoriesCost),'#0da6e0','feeds into Net Capital Cost')}
    ${KPI('Net Capital Cost',F(o.netCapitalCost),'#1fb97a','vehicle + accessories')}
    ${KPI('Added Dep/Year',F(o.accessoriesCost*o.depRateAnnual),'#9775ea','accCost × depRate')}
  </div>
  ${accs.length===0?`<div class="card" style="text-align:center;padding:40px;color:var(--mu)"><div style="font-size:36px">🔧</div><div style="margin-top:8px">No accessories added.</div><div style="font-size:11px;margin-top:4px">Items such as GPS tracker, fire extinguisher, tool kit, spare wheel carrier, etc.</div></div>`:`
  <div class="card">
    <div class="ch"><div class="ct">Accessories List</div><span class="badge ba">${accs.length} item${accs.length!==1?'s':''}</span></div>
    <div class="tw"><table class="acc-tbl"><thead><tr><th>Name</th><th class="nr">Qty</th><th class="nr">Unit Cost (₹)</th><th class="nr">GST %</th><th class="nr">Install (₹)</th><th class="nr">Useful Life (yr)</th><th class="nr">Total (incl. GST)</th><th>Actions</th></tr></thead>
    <tbody>
      ${accs.map(a=>`<tr>
        <td><input value="${esc(a.name||'')}" oninput="State.updateAccessory('${a.id}','name',this.value)" placeholder="Name"></td>
        <td><input type="number" value="${a.qty||1}" style="width:50px" oninput="State.updateAccessory('${a.id}','qty',+this.value)"></td>
        <td><input type="number" value="${a.unitCost||0}" style="width:90px" oninput="State.updateAccessory('${a.id}','unitCost',+this.value)"></td>
        <td><input type="number" value="${a.gstPct||18}" style="width:50px" oninput="State.updateAccessory('${a.id}','gstPct',+this.value)"></td>
        <td><input type="number" value="${a.installCost||0}" style="width:80px" oninput="State.updateAccessory('${a.id}','installCost',+this.value)"></td>
        <td><input type="number" value="${a.usefulLife||8}" style="width:50px" oninput="State.updateAccessory('${a.id}','usefulLife',+this.value)"></td>
        <td class="nr">${F((n(a.qty)*n(a.unitCost)*(1+n(a.gstPct)/100))+n(a.installCost))}</td>
        <td>
          <button class="btn bs sm" onclick="UI.dupAccessory('${a.id}')">⧉</button>
          <button class="btn sm" style="color:var(--re)" onclick="State.removeAccessory('${a.id}');UI._render()">🗑</button>
        </td>
      </tr>`).join('')}
    </tbody></table></div>
  </div>`}`;
},

vintage(i,o){const v=State.vintage, vo=State.vintageOutputs, vd=vo.vintage;const cmp=vd.newComparison;return `

  <div class="sh mb4"><div><div class="st">Vintage TTU Dashboard</div><div class="ss">Used-vehicle valuation &amp; contract pricing — Straight Line Depreciation per Companies Act 2013</div></div></div>
  <div class="g4 mb4">
    ${KPI('Age of Vehicle',ageLabel(vd.ageBreakdown),(vd.age>=n(v.usefulLife,8)?'var(--re)':'var(--ac)'),'remaining life: '+vd.remainingLife.toFixed(2)+' yrs')}
    ${KPI('Current Value (Book)',F(vd.currentValue),'#0da6e0','vs new '+F(vd.newPrice))}
    ${KPI('Accumulated Depreciation',F(vd.accumDep),'#f97316','of '+F(vd.newPrice))}
    ${KPI('Savings vs New TTU',F(cmp.savings),cmp.savings>=0?'#1fb97a':'var(--re)',cmp.savingsPct.toFixed(2)+'% lower contract cost')}
  </div>
  <div class="g2 mb4">
    <div class="card"><div class="ch"><div class="ct">Vintage TTU Inputs</div></div>
      <div class="fg">
        ${VINP('newTractorPrice','New Tractor Price (₹, defaults from TTU setup)',v.newTractorPrice)}
        ${VINP('newTrailerPrice','New Trailer Price (₹, defaults from TTU setup)',v.newTrailerPrice)}
        ${DATE_INP('purchaseDate','Purchase Date',v.purchaseDate)}
        ${DATE_INP('currentDate','Current Date',v.currentDate)}
        ${VINP('salvagePct','Salvage Percentage (%)',v.salvagePct,'number','0.1')}
        ${VINP('usefulLife','Useful Life (Years)',v.usefulLife,'number','1')}
        ${VINP('contractPeriod','New Contract Period (Years)',v.contractPeriod,'number','1')}
      </div>
      <div class="callout mt3">Vehicle age is computed from the exact calendar-day difference between Purchase Date and Current Date — <strong>${ageLabel(vd.ageBreakdown)}</strong> (${vd.age.toFixed(2)} yrs internally) — not by simply subtracting years. Every age-dependent calculation below (depreciation, insurance IDV band, remaining life) uses this exact fractional age.</div>
    </div>
    <div class="card"><div class="ch"><div class="ct">Valuation (SLM, Companies Act 2013)</div></div>
      ${mrow('Age of Vehicle',ageLabel(vd.ageBreakdown))}
      ${mrow('New Price (Tractor+Trailer)',F(vd.newPrice))}
      ${mrow('Annual Depreciation',F(vd.annualDep))}
      ${mrow('Accumulated Depreciation',F(vd.accumDep))}
      ${mrow('Current Value (Book Value)',F(vd.currentValue))}
      ${mrow('Salvage Value',F(vd.salvageValue))}
      ${mrow('Remaining Useful Life',vd.remainingLife.toFixed(2)+' years')}
      <div class="callout mt3">Depreciation Rate = (1 − Salvage%) ÷ Useful Life. Annual Depreciation = (New Price − Salvage Value) ÷ Useful Life. Current Value is floored at the Salvage Value.
The same exact age drives the Insurance IDV band and OD-premium tariff tier used elsewhere in this tool, so Insurance and Maintenance below are computed for a vehicle of this actual age — not a fresh 0-age vehicle.</div>
    </div>
  </div>
  <div class="g4 mb4">
    ${KPI('Monthly Capital Cost',F(vd.monthlyCapitalCost),'var(--ac)','loan + own cap, on current value')}
    ${KPI('Est. Fixed Charge/Month',F(vd.estimatedFixedCharge),'#0da6e0','for new contract')}
    ${KPI('Est. Variable/KM',F2(vd.estimatedVarPerKm),'#1fb97a','fuel+tyre+battery+maint')}
    ${KPI(vd.newComparison?'Total Contract Cost':'—',F(vd.totalContractCost),'#9775ea',vo.cYears+'-yr, ex-GST')}
  </div>
  <div class="card mb4"><div class="ch"><div class="ct">Vintage vs New TTU — Side-by-Side Comparison</div></div>
    <div class="tw"><table><thead><tr><th>Metric</th><th class="nr">Vintage TTU (Age ${ageLabel(vd.ageBreakdown)})</th><th class="nr">Brand-New TTU</th><th class="nr">Difference</th></tr></thead>
    <tbody>
      <tr><td>Capital / Book Value</td><td class="nr">${F(vd.currentValue)}</td><td class="nr">${F(cmp.capitalCost)}</td><td class="nr">${F(cmp.capitalCost-vd.currentValue)}</td></tr>
      <tr><td>Fixed Charge / Month</td><td class="nr">${F(vd.estimatedFixedCharge)}</td><td class="nr">${F(cmp.fixedCharge)}</td><td class="nr">${F(cmp.fixedCharge-vd.estimatedFixedCharge)}</td></tr>
      <tr><td>Variable Charge / KM</td><td class="nr">₹ ${vd.estimatedVarPerKm.toFixed(2)}</td><td class="nr">₹ ${cmp.varPerKm.toFixed(2)}</td><td class="nr">₹ ${(cmp.varPerKm-vd.estimatedVarPerKm).toFixed(2)}</td></tr>
      <tr class="ttr"><td><strong>Total Contract Cost (ex-GST)</strong></td><td class="nr">${F(vd.totalContractCost)}</td><td class="nr">${F(cmp.totalContractCost)}</td><td class="nr">${F(cmp.savings)}</td></tr>
    </tbody></table></div>
    <div class="callout mt3" style="background:${cmp.savings>=0?'rgba(31,185,122,.1)':'rgba(230,80,80,.1)'};border-color:${cmp.savings>=0?'var(--gr)':'var(--re)'}">
      Contracting the Vintage TTU is <strong>${F(Math.abs(cmp.savings))} (${Math.abs(cmp.savingsPct).toFixed(2)}%) ${cmp.savings>=0?'cheaper':'more expensive'}</strong> than a brand-new TTU of the same specification over ${vo.cYears} years.
    </div>
  </div>`;
},

formula(i,o){const rows=[['Net Capital Cost','tractorCost + trailerCost + accessoriesCost',F(o.netCapitalCost)],['Annual Depreciation (SLM)','netCapitalCost × (1 − salvage%) ÷ usefulLife',F(o.annualDep)],['Monthly Loan Instalment','PMT((1+loanRate)^(1/12)-1, contractMonths, −loanAmount)',F(o.loanInstalmentPM)],['Net Own-Capital Amortisation','PMT(effMonthlyRate,months,−ownCap) − PMT(effMonthlyRate,months,0,−salvage)',F(o.netOwnCapAmortPM)],['Operator Daily Rate','Basic/day + VDA/day + Bonus/day + Special Allowance',F2(o.opRows[0].dailyRate)],['Single OT/hr','Basic/day ÷ 8',F2(o.opRows[0].singleOT)],['Double OT/hr','(Basic/day + VDA/day) ÷ 8 × 2',F2(o.opRows[0].doubleOT)],['Monthly PF','MIN(ROUND((Basic+VDA)×PF%,2), PF-cap) × 365/12',F(o.opRows[0].monthlyPF)],['Monthly ESI / EC Insurance','if MonthlyWage<₹21,000: ESI%×(Basic+VDA); else ECannual/365 — ×365/12',F(o.opRows[0].monthlyESIorEC)],['Diesel / KM','dieselPrice ÷ fuelEfficiency (KMPL)',F2(o.dieselPerKm)],['DEF / KM','defPrice × (1/fuelEfficiency) × defConsumption%',F2(o.defPerKm)],['Tyre / KM','(tyreCost × changes × numTyres) ÷ totalContractKM',F2(o.tyrePerKm)],['Battery / KM','((unitPrice÷GSTFactor)×life×(1+esc%)^life) ÷ totalContractKM',F2(o.batteryPerKm)],['Maintenance / KM','avg[ PMT(effMonthlyRate,12,−PV(yearlyMaint)) ] ÷ kmPerMonth',F2(o.maintPerKm)],['Insurance IDV','(Tractor+Trailer Cost) × (1 − IDV-depreciation%)',F(o.insuranceYears[0].tractorIDV+o.insuranceYears[0].trailerIDV)],['Insurance Net OD Premium','subTotalOD × (1 − NCB% − specialDiscount%)',F(o.insuranceYears[0].netODP)],['Road Tax Monthly','PMT(effMonthlyRate, 12, −avgRoadTaxPerYear)',F(o.roadTaxAmortPM)],['Fixed Charge / Month','wagesPM + loanInstalment + netOwnCapAmort + avgInsurance + roadTaxAmort + supervisionPM',F(o.fixedChargePM[0])],['Variable Charge / KM','diesel + DEF + lube + tyre + battery + maintenance (all per km)',F2(o.totalVarPerKm)],['Grand Total (Contract)','baseContractCost + OT + uniform&shoe + medical + addl.PF + addl.ESI/Ins + toll',F(o.grandTotal)],['Vintage Current Value','newPrice − (annualDep × min(age,usefulLife)), floored at salvage',F(State.vintageOutputs.vintage.currentValue)],];return `

  <div class="sh mb4"><div><div class="st">Formula Viewer</div><div class="ss">Live formula → value trace, computed from your current inputs (Excel → JavaScript reference)</div></div></div>
  <div class="card"><div class="ch"><div class="ct">Formula Reference</div></div>
    <div class="tw"><table><thead><tr><th>Metric</th><th>Formula</th><th class="nr">Live Value</th></tr></thead>
    <tbody>${rows.map(r=>`<tr><td>${esc(r[0])}</td><td><span class="fchip">${esc(r[1])}</span></td><td class="nr mono bold">${r[2]}</td></tr>`).join('')}</tbody></table></div>
  </div>`;
},

reports(i,o){const now=new Date().toLocaleDateString('en-IN',{day:'2-digit',month:'long',year:'numeric'});const mPcts=[n(i.maintPct_yr1),n(i.maintPct_yr2),n(i.maintPct_yr3),n(i.maintPct_yr4)];const sy=n(i.contractStartYear,2026);return `

  <div class="sh mb4 no-print">
    <div><div class="st">Cost Estimation Report</div><div class="ss">${esc(i.projectName)} · ${esc(i.ttuType)}</div></div>
    <div class="flex gap2">
      <button class="btn bs sm" onclick="Exp.csv()">⬇️ CSV/Excel</button>
      <button class="btn bs sm" onclick="Exp.exportSection()">🖨️ Export This Report</button>
      <button class="btn bp sm" onclick="Exp.fullReport()">📑 Complete Project Report</button>
    </div>
  </div>
  <div style="max-width:900px">
    <div class="card mb4" style="border-top:4px solid var(--ac)">
      <div class="flex jb ic mb4">
        <div>
          <div style="font-size:20px;font-weight:800">OIL India Limited — TTU Cost Estimation</div>
          <div style="font-size:12px;color:var(--mu)">${esc(i.projectName)} · ${now} · CONFIDENTIAL</div>
        </div>
        <span class="badge ba" style="padding:4px 12px;font-size:11px">Internal Estimate</span>
      </div>
      <div class="g4" style="gap:9px">
        ${[['TTU Type',i.ttuType],['Contract Period',i.contractPeriod+' Years'],['KM per Month',n(i.kmPerMonth).toLocaleString('en-IN')],['Diesel Price','₹ '+i.dieselPrice+'/L'],['Tractor Cost',F(n(i.tractorCost))],['Trailer Cost',F(n(i.trailerCost))],['No. Tyres',i.numTyres],['Battery Life',i.batteryLife+' yrs']].map(([l,v])=>`<div style="padding:5px 0;border-bottom:1px solid var(--div)"><div style="font-size:10px;color:var(--mu);text-transform:uppercase">${l}</div><div style="font-size:12px;font-weight:600">${v}</div></div>`).join('')}
      </div>
    </div>
    <div class="card mb4"><div class="ch"><div class="ct">ANNEXURE-II : Fixed &amp; Variable Charges</div></div>
      <div style="font-size:11px;color:var(--mu);margin-bottom:10px">A. FIXED CHARGE BREAKUP (₹/month)</div>
      <div class="tw"><table><thead><tr><th>Component</th>${o.fixedChargePM.map((f,idx)=>`<th class="nr">Year ${idx+1} (${sy+idx})</th>`).join('')}</tr></thead>
      <tbody>
        <tr><td>Wages (excl. PF &amp; ESI)</td>${o.wagesPM.map(w=>`<td class="nr">${F(w)}</td>`).join('')}</tr>
        <tr><td>Supervision (1/${i.supervisionTTUShare} share)</td>${o.supervisionPM.map(s=>`<td class="nr">${F(s)}</td>`).join('')}</tr>
        <tr><td>Loan Instalment</td>${o.fixedChargePM.map(()=>`<td class="nr">${F(o.loanInstalmentPM)}</td>`).join('')}</tr>
        <tr><td>Net Own Capital Amortisation</td>${o.fixedChargePM.map(()=>`<td class="nr">${F(o.netOwnCapAmortPM)}</td>`).join('')}</tr>
        <tr><td>Insurance (avg amortised)</td>${o.fixedChargePM.map(()=>`<td class="nr">${F(o.avgMonthlyInsurance)}</td>`).join('')}</tr>
        <tr><td>Road Tax, Permit, Pollution</td>${o.fixedChargePM.map(()=>`<td class="nr">${F(o.roadTaxAmortPM)}</td>`).join('')}</tr>
        <tr class="ttr"><td><strong>FIXED COST/MONTH (excl. GST)</strong></td>${o.fixedChargePM.map(f=>`<td class="nr">${F(f)}</td>`).join('')}</tr>
      </tbody></table></div>
      <div style="font-size:11px;color:var(--mu);margin-top:15px;margin-bottom:10px">E. VARIABLE CHARGES PER KM (excl. GST)</div>
      <div class="tw"><table><thead><tr><th>Component</th><th class="nr">₹ per KM</th><th>Notes</th></tr></thead>
      <tbody>
        <tr><td>Fuel (Diesel)</td><td class="nr">${o.dieselPerKm.toFixed(2)}</td><td>dieselPrice ÷ fuelEfficiency (${i.fuelEfficiency} KMPL)</td></tr>
        <tr><td>DEF Consumption</td><td class="nr">${o.defPerKm.toFixed(2)}</td><td>${i.defConsumptionPctOfFuel}% of fuel vol × avg DEF price</td></tr>
        <tr><td>Battery</td><td class="nr">${o.batteryPerKm.toFixed(2)}</td><td>${i.batteryLife}-yr life, set cost ${F(o.batterySetCostDisplay)}</td></tr>
        <tr><td>Tyre/Tube</td><td class="nr">${o.tyrePerKm.toFixed(2)}</td><td>${i.numTyres} tyres × ${F(n(i.tyreCostExclGST))} ÷ ${n(i.tyreLife).toLocaleString('en-IN')} KM life</td></tr>
        <tr><td>Lube Oil (${i.lubeOilPct}% of diesel)</td><td class="nr">${o.lubePerKm.toFixed(2)}</td><td>lubricant rate per km</td></tr>
        <tr><td>Maintenance (avg yr-wise)</td><td class="nr">${o.maintPerKm.toFixed(2)}</td><td>Yr1: ${mPcts[0]}%, Yr2: ${mPcts[1]}%, Yr3: ${mPcts[2]}%, Yr4: ${mPcts[3]}% of capex</td></tr>
        <tr class="ttr"><td><strong>VARIABLE CHARGES/KM (excl. GST)</strong></td><td class="nr">${o.totalVarPerKm.toFixed(2)}</td><td>total per km</td></tr>
      </tbody></table></div>
    </div>
    <div class="card mb4"><div class="ch"><div class="ct">ANNEXURE-III : Total Contract Cost</div></div>
      <div class="tw"><table><thead><tr><th>Item</th><th class="nr">Amount (₹)</th></tr></thead>
      <tbody>
        <tr><td>1. Total Fixed Cost (all years, excl. GST)</td><td class="nr">${F(o.totalFixedForContract)}</td></tr>
        <tr><td>2. Total Variable Cost (excl. GST)</td><td class="nr">${F(o.totalVarForContract)}</td></tr>
        <tr><td>BASE CONTRACT COST (excl. Reimbursables, excl. GST)</td><td class="nr mono bold">${F(o.baseContractCost)}</td></tr>
        <tr><td style="padding-top:10px;color:var(--mu)">ADDL OUTGO (Reimbursables):</td><td></td></tr>
        <tr><td>&nbsp;&nbsp;OT (contract total)</td><td class="nr">${F(o.totalOTContract)}</td></tr>
        <tr><td>&nbsp;&nbsp;Uniform &amp; Shoe (${o.cYears} yrs)</td><td class="nr">${F(o.uniformShoeTotal)}</td></tr>
        <tr><td>&nbsp;&nbsp;Medical Fitness Certificates</td><td class="nr">${F(o.medicalCertTotal)}</td></tr>
        <tr><td>&nbsp;&nbsp;Addl. PF (contract total)</td><td class="nr">${F(o.pfAddlContract)}</td></tr>
        <tr><td>&nbsp;&nbsp;Addl. ESI/EC Insurance (contract total)</td><td class="nr">${F(o.esiInsAddlContract)}</td></tr>
        <tr><td>&nbsp;&nbsp;Toll Tax (Reimbursable, ${o.cYears} yrs)</td><td class="nr">${F(o.tollTotal)}</td></tr>
        <tr class="ttr"><td><strong>GRAND TOTAL (excl. GST)</strong></td><td class="nr">${F(o.grandTotal)}</td></tr>
        <tr><td>GST @ ${i.gstPct}%</td><td class="nr">${F(o.grandTotalWithGST - o.grandTotal)}</td></tr>
        <tr style="background:var(--acg)"><td style="color:var(--ac)"><strong>GRAND TOTAL (incl. GST @ ${i.gstPct}%)</strong></td><td class="nr" style="color:var(--ac)">${F(o.grandTotalWithGST)}</td></tr>
        <tr><td style="padding-top:8px">Per-KM Rate (excl. GST)</td><td class="nr" style="padding-top:8px">₹ ${o.perKmExGST.toFixed(2)}</td></tr>
        <tr><td>Per-KM Rate (incl. GST)</td><td class="nr">₹ ${o.perKmWithGST.toFixed(2)}</td></tr>
      </tbody></table></div>
    </div>
    <div style="text-align:center;padding:14px;font-size:10px;color:var(--mu);border-top:1px solid var(--bdr)">Generated by TTU Cost Estimation System v6.0 · ${now} · OIL India Limited · CONFIDENTIAL · All values in INR</div>
  </div>`;
},

async projects(){return `

  <div>
    <div class="sh mb4">
      <div><div class="st">Projects</div><div class="ss">Manage cost estimation projects — 20MT and 30MT TTU configurations</div></div>
      <button class="btn bp" onclick="UI.newProject()">➕ New Project</button>
    </div>
    <div class="card mb4" style="border-left:3px solid var(--ac)">
      <div class="flex jb ic">
        <div><div style="font-size:10px;color:var(--mu);text-transform:uppercase">Currently Open</div><div style="font-size:15px;font-weight:700">${esc(State.inputs.projectName)}</div><div style="font-size:11px;color:var(--tx2)">${State.inputs.ttuType}</div></div>
        <button class="btn bp sm" onclick="DB.save().then(()=>UI.toast('Saved','💾'))">💾 Save Now</button>
      </div>
    </div>
    <div id="plist"><div class="flex ic gap2" style="padding:40px;justify-content:center"><div class="spin"></div><span style="color:var(--mu)">Loading…</span></div></div>
  </div>`;
}

};

const Exp = {csv(){const i=State.inputs, o=State.outputs, now=new Date().toISOString().slice(0,10), sy=n(i.contractStartYear,2026);const rows=[['OIL INDIA LIMITED — TTU COST ESTIMATION SYSTEM v6.0'],['Project',i.projectName],['Date',now],['TTU Type',i.ttuType],['Contract Period',i.contractPeriod+' Years'],['Tractor Cost (₹)',n(i.tractorCost)],['Trailer Cost (₹)',n(i.trailerCost)],['Net Capital Cost (₹)',o.netCapitalCost],[],['=== FIXED CHARGES (₹/MONTH) ===',...o.fixedChargePM.map((f,idx)=>'Year '+(idx+1)+' ('+(sy+idx)+')')],['Wages (excl PF & ESI)',...o.wagesPM],['Supervision',...o.supervisionPM],['Loan Instalment',...o.fixedChargePM.map(()=>o.loanInstalmentPM)],['Net Own Capital Amortisation',...o.fixedChargePM.map(()=>o.netOwnCapAmortPM)],['Insurance (avg amortised)',...o.fixedChargePM.map(()=>o.avgMonthlyInsurance)],['Road Tax, Permit',...o.fixedChargePM.map(()=>o.roadTaxAmortPM)],['TOTAL FIXED/MONTH',...o.fixedChargePM],[],['=== VARIABLE CHARGES (₹/KM) ===','₹/km'],['Diesel',o.dieselPerKm.toFixed(2)],['DEF',o.defPerKm.toFixed(2)],['Lube Oil',o.lubePerKm.toFixed(2)],['Battery',o.batteryPerKm.toFixed(2)],['Tyre/Tube',o.tyrePerKm.toFixed(2)],['Maintenance (avg)',o.maintPerKm.toFixed(2)],['TOTAL VARIABLE/KM',o.totalVarPerKm.toFixed(2)],[],['=== CONTRACT TOTALS (₹) ==='],['Total Fixed (contract)',o.totalFixedForContract],['Total Variable (contract)',o.totalVarForContract],['Base Contract Cost',o.baseContractCost],['OT (contract total)',o.totalOTContract],['Uniform & Shoe',o.uniformShoeTotal],['Medical Certificate',o.medicalCertTotal],['Addl. PF (contract)',o.pfAddlContract],['Addl. ESI/EC Insurance (contract)',o.esiInsAddlContract],['Toll (Reimbursable)',o.tollTotal],['GRAND TOTAL excl. GST',o.grandTotal],['GST @ '+i.gstPct+'%',o.grandTotalWithGST-o.grandTotal],['GRAND TOTAL incl. GST',o.grandTotalWithGST],['Per-KM Rate excl. GST',o.perKmExGST.toFixed(2)],['Per-KM Rate incl. GST',o.perKmWithGST.toFixed(2)],[],['=== YEAR-WISE MAINTENANCE (₹) ===','Annual'],...o.maintYr.map((m,idx)=>['Year '+(idx+1)+' ('+i['maintPct_yr'+(idx+1)]+'% of capex)',m]),[],['=== VINTAGE TTU (if applicable) ==='],['Age (yrs)',State.vintageOutputs.vintage.age],['Current Value',State.vintageOutputs.vintage.currentValue],['Total Contract Cost (vintage)',State.vintageOutputs.vintage.totalContractCost],];const csv=rows.map(r=>r.map(c=>{ const s=String(c??''); return s.includes(',')||s.includes('"')?'"'+s.replace(/"/g,'""')+'"'; }).join(',')).join('\r\n');const a=document.createElement('a');a.href='data/csv;charset=utf-8,\uFEFF'+encodeURIComponent(csv);a.download='OIL_TTU_'+i.ttuType+''+i.projectName.replace(/[^a-z0-9]/gi,'')+'_'+now+'.csv';a.click(); UI.toast('CSV exported','✅');},

// Section-wise export: print ONLY the module currently on screen.// Print CSS (see @media print) releases every height/overflow// constraint on #app/.main/.content first, so a module that is// 2, 5, or 10 pages long paginates in full instead of being clipped// to one viewport-height screenshot.exportSection(){if(typeof UI.hideModal==='function') UI.hideModal();window.print();},

// Complete project report: renders every module's HTML (cover page +// Main/TTU dashboards + every cost component + Formula Viewer +// Reports) into one off-screen container, swaps it in for print only,// then hands off to the browser's native paginated print/PDF engine.async fullReport(){const order=['dashboard','ttu','labour','fuel','tyre','maintenance','insurance','tax','accessories','vintage','formula','reports'];const titles={dashboard:'Main Dashboard',ttu:'TTU Dashboard',labour:'Labour / Wages',fuel:'Fuel & Lubricants',tyre:'Tyre & Battery',maintenance:'Maintenance',insurance:'Insurance',tax:'Road Tax & Permit',accessories:'Accessories',vintage:'Vintage TTU Dashboard',formula:'Formula Summary / Viewer',reports:'Reports & Final Cost Summary'};const i=State.inputs, now=new Date().toLocaleDateString('en-IN',{day:'2-digit',month:'long',year:'numeric'});const parts=[<div class="print-cover">
      <div style="font-size:12px;letter-spacing:.15em;color:#555">OIL INDIA LIMITED</div>
      <h1>TTU Cost Estimation — Complete Project Report</h1>
      <p><strong>Project:</strong> ${esc(i.projectName||'Untitled')}</p>
      <p><strong>TTU Type:</strong> ${esc(i.ttuType)} &nbsp;·&nbsp; <strong>Contract Period:</strong> ${esc(i.contractPeriod)} Years</p>
      <p>Generated ${now}</p>
    </div>];for(const p of order){const fn=Pages[p]; if(!fn) continue;let html;try{ html=fn(State.inputs,State.outputs); if(html instanceof Promise) html=await html; }catch(e){ html='<div class="card">Section could not be rendered.</div>'; }parts.push(<div class="print-section"><div class="print-section-title">${esc(titles[p]||p)}</div>${html}</div>);}document.getElementById('fullReportPrint').innerHTML=parts.join('');document.body.classList.add('printing-full');UI.toast('Preparing complete project report…','📑');requestAnimationFrame(()=>requestAnimationFrame(()=>window.print()));}};window.addEventListener('afterprint',()=>{document.body.classList.remove('printing-full');const c=document.getElementById('fullReportPrint'); if(c) c.innerHTML='';});

// ═══════════════════════════════════════════════════════════════// UI CONTROLLER — cursor-stable input handling// ═══════════════════════════════════════════════════════════════const UI = {current:'dashboard',

go(page){this.current=page;document.querySelectorAll('.ni').forEach(el=>el.classList.toggle('active',el.dataset.page===page));const titles={dashboard:'Main Dashboard',ttu:'TTU Dashboard',labour:'Labour / Wages',fuel:'Fuel',tyre:'Tyre & Battery',maintenance:'Maintenance',insurance:'Insurance',tax:'Road Tax & Permit',accessories:'Accessories',vintage:'Vintage TTU Dashboard',formula:'Formula Viewer',reports:'Reports',projects:'Projects'};document.getElementById('pageTitle').textContent=titles[page]||page;this._render();},

async _render(){const content=document.getElementById('content');const scrollTop=content.scrollTop;const fn=Pages[this.current];if(!fn){ content.innerHTML=<div style="text-align:center;padding:60px;color:var(--mu)">Page not found</div>; return; }let html=fn(State.inputs,State.outputs);if(html instanceof Promise) html=await html;content.innerHTML=html;content.scrollTop=scrollTop;this._flashChanges();if(this.current==='projects') this._loadProjects();document.getElementById('projType').textContent=State.inputs.ttuType;},

// Compare every [data-outkey] element's rendered text against what it// showed on the previous render of this page. Anything that changed as// a knock-on effect of the edit just made gets a ~500ms glow so the user// can see, at a glance, exactly which outputs were recalculated._prevOut:{},_flashChanges(){const store=this._prevOut[this.current]||(this._prevOut[this.current]={});document.querySelectorAll('#content [data-outkey]').forEach(el=>{const k=el.dataset.outkey, val=el.textContent;if(store[k]!==undefined && store[k]!==val){el.classList.remove('flash'); void el.offsetWidth; el.classList.add('flash');}store[k]=val;});},

// On focus: strip thousands separators so the user edits a plain raw// number. Formatting is re-applied only on blur (never mid-keystroke).inpFocus(el){if(el.dataset.numeric!=='1') return;const pos=el.selectionStart??el.value.length;const commasBefore=(el.value.slice(0,pos).match(/,/g)||[]).length;const raw=numUnformat(el.value);if(raw!==el.value){el.value=raw;const newPos=Math.max(0,pos-commasBefore);try{ el.setSelectionRange(newPos,newPos); }catch(_){}}},vinpFocus(el){ this.inpFocus(el); },

// Cursor-stable: update state on input without ever reformatting the// field the user is actively typing in. Dependent outputs elsewhere on// the page are refreshed via a soft patch that leaves the focused field// (and its raw, un-formatted value + caret position) untouched.inp(el){const key=el.dataset.key; if(!key)return;

if(el.tagName==='SELECT'){
  const val=el.value;
  if(key==='ttuType' && val!==State.inputs.ttuType){
    State.switchTTUType(val);
    document.getElementById('projName').textContent=State.inputs.projectName;
    document.getElementById('projType').textContent=val;
    this._render();
    return;
  }
  State.setInput(key,val);
  this._render();
  return;
}

if(key==='projectName'){
  State.inputs[key]=el.value;
  document.getElementById('projName').textContent=el.value||'Untitled';
  State._scheduleSave();
  return;
}

// Numeric fields: sanitize keystrokes in place (digits, one decimal
// point, optional leading minus) instead of relying on browser
// number-input validation, which breaks cursor/selection APIs.
if(el.dataset.numeric==='1'){
  const pos=el.selectionStart??el.value.length;
  const allowNeg=el.dataset.allowNeg==='1';
  const allowDecimal=el.dataset.integer!=='1';
  const {value,cursorShift}=sanitizeNumericTyping(el.value,pos,allowNeg,allowDecimal);
  if(value!==el.value){
    el.value=value;
    const newPos=Math.max(0,pos-cursorShift);
    try{ el.setSelectionRange(newPos,newPos); }catch(_){}
  }
}

State.inputs[key]=el.value;
const def=State.defaults||{};
const numVal=parseFloat(el.value), defVal=parseFloat(def[key]);
const isChanged = isNaN(numVal)||isNaN(defVal) ? el.value!==String(def[key]) : Math.abs(numVal-defVal)>=0.0001;
isChanged ? State.changed.add(key) : State.changed.delete(key);
el.classList.toggle('chg',isChanged);
State._recompute();
State._scheduleSave();
this._patchComputedFields();

},

// Refresh dependent computed fields elsewhere on the page without// disturbing the field currently being edited._patchComputedFields(){const active=document.activeElement;const isEditingInput = active && active.tagName==='INPUT' && active.dataset && (active.dataset.key||active.dataset.vkey);if(isEditingInput){this._softPatch();} else {this._render();}},

_softPatch(){const content=document.getElementById('content');const scrollTop=content.scrollTop;const active=document.activeElement;const key=active && active.dataset ? active.dataset.key : null;const selStart=active && active.selectionStart!=null ? active.selectionStart : null;const selEnd=active && active.selectionEnd!=null ? active.selectionEnd : null;const fn=Pages[this.current];if(!fn) return;let html=fn(State.inputs,State.outputs);if(html instanceof Promise){ html.then(h=>{ content.innerHTML=h; content.scrollTop=scrollTop; this._flashChanges(); this._restoreFocus(key,selStart,selEnd); }); return; }content.innerHTML=html;content.scrollTop=scrollTop;this._flashChanges();this._restoreFocus(key,selStart,selEnd);},

restoreFocus(key,selStart,selEnd){if(!key) return;const el=document.querySelector([data-key="${key}"]);if(el){// Always show the raw (unformatted) in-progress value for the// field being edited — the render pass may have applied display// formatting, which must never appear while typing.if(el.dataset.numeric==='1') el.value=numRaw(State.inputs[key]);el.focus();const len=el.value.length;try{ if(selStart!=null) el.setSelectionRange(Math.min(selStart,len),Math.min(selEnd??selStart,len)); }catch(){}}},

inpBlur(el){const key=el.dataset.key; if(!key)return;let val;if(el.dataset.numeric==='1'){val=parseFloat(numUnformat(el.value));if(isNaN(val)) return;if(el.dataset.min!==undefined && el.dataset.min!=='') val=Math.max(val,parseFloat(el.dataset.min));if(el.dataset.max!==undefined && el.dataset.max!=='') val=Math.min(val,parseFloat(el.dataset.max));if(el.dataset.integer==='1') val=Math.round(val);} else {val=el.value;}State.setInput(key,val);this._render();},

vinp(el){const key=el.dataset.vkey; if(!key)return;if(el.tagName==='SELECT'){State.vintage[key]=+el.value;State._recomputeVintage();this.render();return;}if(el.dataset.numeric==='1'){const pos=el.selectionStart??el.value.length;const {value,cursorShift}=sanitizeNumericTyping(el.value,pos,el.dataset.allowNeg==='1');if(value!==el.value){el.value=value;const newPos=Math.max(0,pos-cursorShift);try{ el.setSelectionRange(newPos,newPos); }catch(){}}}State.vintage[key]=el.value;State._recomputeVintage();this._softPatchVintage();},vinpBlur(el){const key=el.dataset.vkey; if(!key)return;let val;if(el.dataset.numeric==='1'){val=parseFloat(numUnformat(el.value));if(isNaN(val)) return;} else {val=el.value;}State.setVintageInput(key,val);this._render();},softPatchVintage(){const content=document.getElementById('content');const scrollTop=content.scrollTop;const active=document.activeElement;const key=active && active.dataset ? active.dataset.vkey : null;const selStart=active && active.selectionStart!=null ? active.selectionStart : null;const selEnd=active && active.selectionEnd!=null ? active.selectionEnd : null;content.innerHTML=Pages.vintage(State.inputs,State.outputs);content.scrollTop=scrollTop;if(key){const el=document.querySelector([data-vkey="${key}"]);if(el){if(el.dataset.numeric==='1') el.value=numRaw(State.vintage[key]);el.focus();const len=el.value.length;try{ if(selStart!=null) el.setSelectionRange(Math.min(selStart,len),Math.min(selEnd??selStart,len)); }catch(){}}}},

showFormula(key){const doc=FORMULA_DOCS[key];if(!doc){ this.toast('No formula reference for this field yet','ℹ️'); return; }const rows=doc.detail(State.inputs,State.outputs);const body=<div class="fdoc-formula">${esc(doc.formula)}</div>
      ${rows.map(([l,v])=><div class="fdoc-row"><span style="color:var(--mu)">${esc(l)}</span><span class="mono bold">${esc(String(v))}</span></div>).join('')};this.showModal(doc.title, body, [{label:'Close',cls:'bp',action:()=>this.hideModal()}]);},

showDefaultsMenu(){const type=State.inputs.ttuType||'20MT';const hasCustom=!!State.customDefaults[type];this.showModal('Default Values',<p style="margin-bottom:10px;font-size:12px;color:var(--mu)">Manage the engineering defaults used for <strong>${type}</strong> projects and the "Reset" action.</p>
       <div style="display:flex;flex-direction:column;gap:8px">
         <button class="btn bw" style="justify-content:flex-start" onclick="UI.hideModal();UI.confirmReset()">↺ Reset Current Inputs<br><span style="font-weight:400;font-size:10.5px;color:var(--mu)">Revert this project's inputs to the active default</span></button>
         <button class="btn bp" style="justify-content:flex-start" onclick="UI.confirmSaveAsDefault()">💾 Save Current Values as New Default<br><span style="font-weight:400;font-size:10.5px;color:var(--mu)">Make today's inputs the default for future ${type} projects</span></button>
         <button class="btn bd" style="justify-content:flex-start" onclick="UI.confirmRestoreFactory()">🏭 Restore Original Factory Defaults<br><span style="font-weight:400;font-size:10.5px">${hasCustom?'A custom default is currently active — this clears it':'Factory defaults are already active'}</span></button>
       </div>,[{label:'Close',cls:'bs',action:()=>this.hideModal()}]);},confirmSaveAsDefault(){const type=State.inputs.ttuType||'20MT';this.showModal('Save Current Values as New Default',<p>Do you want to make the current inputs the new default values for <strong>${type}</strong> projects?</p>
       <p style="color:var(--ye);font-size:12px;margin-top:7px">This persists across page refresh and future projects of this TTU type until changed again.</p>,[{label:'Cancel',cls:'bs',action:()=>this.hideModal()},{label:'Yes, Save as Default',cls:'bp',action()=>{ await State.saveCurrentAsDefault(); this.hideModal(); this._render(); this.toast('Saved as new default','💾'); }}]);},confirmRestoreFactory(){const type=State.inputs.ttuType||'20MT';this.showModal('Restore Original Factory Defaults',<p>This clears any custom default for <strong>${type}</strong> and resets the current inputs to the original engineering constants from the Excel workbook.</p>
       <p style="color:var(--re);font-size:12px;margin-top:7px">⚠️ Any custom default you saved for ${type} will be permanently removed.</p>,[{label:'Cancel',cls:'bs',action:()=>this.hideModal()},{label:'Restore Factory Defaults',cls:'bd',action()=>{ await State.restoreFactoryDefaults(); this.hideModal(); this._render(); this.toast('Factory defaults restored','🏭'); }}]);},

confirmReset(){this.showModal('Reset to Default Values',<p style="margin-bottom:9px">Restore ALL inputs to engineering defaults from the Annexure reference documents and Excel workbook.</p>
       <p style="color:var(--ye);font-size:12px">⚠️ ${State.changed.size} modified field${State.changed.size!==1?'s':''} and ${State.accessories.length} accessories will be cleared.</p>,[{label:'Cancel',cls:'bs',action:()=>this.hideModal()},{label:'↺ Reset',cls:'bw',action:()=>{ this.hideModal(); State.resetToDefaults(); this._render(); UI.toast('Reset to defaults','↺'); }}]);},

toast(msg,icon='ℹ️'){const el=document.createElement('div'); el.className='toast';el.innerHTML=<span>${icon}</span><span>${esc(msg)}</span>;document.getElementById('toasts').appendChild(el);setTimeout(()=>{ el.classList.add('out'); setTimeout(()=>el.remove(),280); },3000);},

showModal(title,body,buttons=[]){document.getElementById('mo-title').textContent=title;document.getElementById('mo-body').innerHTML=body;const foot=document.getElementById('mo-foot'); foot.innerHTML='';buttons.forEach(b=>{ const btn=document.createElement('button'); btn.className='btn '+(b.cls||'bs'); btn.textContent=b.label; btn.onclick=b.action; foot.appendChild(btn); });document.getElementById('mo').classList.add('open');},hideModal(){ document.getElementById('mo').classList.remove('open'); },

newProject(){this.showModal('New Project',
      <div class="frow mb3"><label class="fl">Project Name</label><input type="text" class="fi" id="np-n" value="New TTU Project"></div>
      <div class="frow"><label class="fl">TTU Type</label><select class="fs" id="np-t"><option value="20MT">20MT/24MT — Double Axle</option><option value="30MT">30MT/35MT — Triple Axle</option></select></div>,[{label:'Cancel',cls:'bs',action:()=>this.hideModal()},{label:'➕ Create',cls:'bp',action()=>{const name=document.getElementById('np-n').value.trim();const type=document.getElementById('np-t').value;if(!name){this.toast('Name required','⚠️');return;}const d=State.getDefaults(type);State.defaults=d; State.inputs={...d,projectName,ttuType};State.accessories=[]; State.changed=new Set(); State.projId=null; State.undoStack=[]; State.redoStack=[];State.vintage=State.getVintageDefaults();State.outputs=Engine.calc(State.inputs,[]);State._recomputeVintage();document.getElementById('projName').textContent=name;document.getElementById('projType').textContent=type;this.hideModal(); await DB.save(); this.go('dashboard');this.toast('Created: '+name,'✅');}}]);setTimeout(()=>document.getElementById('np-n')?.select(),80);},

addAccModal(){this.showModal('Add Accessory',
      <div class="fg">
        <div class="frow"><label class="fl">Name</label><input type="text" class="fi" id="acc-name" placeholder="e.g. GPS Tracker"></div>
        <div class="frow"><label class="fl">Quantity</label><input type="number" class="fi" id="acc-qty" value="1" min="1"></div>
        <div class="frow"><label class="fl">Unit Cost (₹)</label><input type="number" class="fi" id="acc-cost" value="0" min="0"></div>
        <div class="frow"><label class="fl">GST %</label><input type="number" class="fi" id="acc-gst" value="18" min="0"></div>
        <div class="frow"><label class="fl">Installation Cost (₹)</label><input type="number" class="fi" id="acc-inst" value="0" min="0"></div>
        <div class="frow"><label class="fl">Useful Life (Years)</label><input type="number" class="fi" id="acc-life" value="8" min="1"></div>
      </div>,[{label:'Cancel',cls:'bs',action:()=>this.hideModal()},{label:'Add',cls:'bp',action:()=>{const name=document.getElementById('acc-name').value.trim();if(!name){this.toast('Name required','⚠️');return;}State.addAccessory({name,qty:+document.getElementById('acc-qty').value,unitCost:+document.getElementById('acc-cost').value,gstPct:+document.getElementById('acc-gst').value,installCost:+document.getElementById('acc-inst').value,usefulLife:+document.getElementById('acc-life').value});this.hideModal(); this._render(); this.toast('Accessory added','✅');}}]);setTimeout(()=>document.getElementById('acc-name')?.focus(),80);},

dupAccessory(id){const a=State.accessories.find(x=>x.id===id);if(!a)return;const {id:_old, ...rest}=a;State.addAccessory({...rest,name.name+' (Copy)'});this._render(); this.toast('Duplicated','✅');},

async _loadProjects(){const all=await DB.getAll(); all.sort((a,b)=>new Date(b.savedAt)-new Date(a.savedAt));const el=document.getElementById('plist'); if(!el)return;if(!all.length){ el.innerHTML=<div class="card" style="text-align:center;padding:40px;color:var(--mu)"><div style="font-size:36px">📂</div><div style="margin-top:8px">No saved projects yet</div></div>; return; }el.innerHTML=<div style="font-size:11px;color:var(--mu);margin-bottom:9px">${all.length} project${all.length!==1?'s':''} saved locally</div>
    <div class="gauto">
      ${all.map(p=>{ const isOpen=p.id===State.projId; const upd=new Date(p.savedAt).toLocaleString('en-IN',{dateStyle:'medium',timeStyle:'short'});
        return <div style="background:var(--card);border:1px solid ${isOpen?'var(--ac)':'var(--bdr)'};border-radius:var(--r);padding:12px;cursor:pointer;position:relative;transition:border-color .1s" onclick="UI._openProj('${p.id}')">${isOpen?'<div style="margin-bottom:6px"><span class="badge ba">Open</span></div>':''}<div style="font-size:13px;font-weight:600;margin-bottom:2px">${esc(p.name)}</div><div style="font-size:11px;color:var(--mu)">${p.inputs?.ttuType||'—'} · ${esc(String(p.inputs?.contractPeriod||4))} yrs</div><div style="font-size:10px;color:var(--mu);margin-top:5px">Saved: ${upd}</div><div style="position:absolute;top:8px;right:8px;display:flex;gap:3px"><button class="tbtn" title="Duplicate" onclick="event.stopPropagation();UI._dupProj('${p.id}')">⧉</button><button class="tbtn" title="Delete" style="color:var(--re)" onclick="event.stopPropagation();UI._delProj('${p.id}','${esc(p.name)}')">🗑</button></div></div>; }).join('')}
    </div>;},

async _openProj(id){ const p=await DB.get(id); if(!p){this.toast('Not found','❌');return;} State.loadProject(p); document.getElementById('projName').textContent=p.name; document.getElementById('projType').textContent=p.inputs?.ttuType||'20MT'; this.go('dashboard'); this.toast('Opened: '+p.name,'✅'); },async _dupProj(id){ const c=await DB.dup(id); this.toast('Duplicated: '+c.name,'✅'); this._loadProjects(); },_delProj(id,name){ this.showModal('Delete Project',<p>Delete <strong>${esc(name)}</strong>?</p><p style="color:var(--re);margin-top:7px;font-size:12px">Cannot be undone.</p>,[{label:'Cancel',cls:'bs',action:()=>this.hideModal()},{label:'Delete',cls:'bd',action()=>{ await DB.del(id); this.hideModal(); this.toast('Deleted','✅'); this._loadProjects(); }}]); }};

// ═══════════════════════════════════════════════════════════════// BOOT// ═══════════════════════════════════════════════════════════════async function boot(){try {await DB.open();await State.loadCustomDefaults();State.init();const lastId=await DB._g('settings','lastId').then(r=>r?.value).catch(()=>null);if(lastId){ const p=await DB.get(lastId).catch(()=>null); if(p){ State.loadProject(p); document.getElementById('projName').textContent=p.name||''; document.getElementById('projType').textContent=p.inputs?.ttuType||'20MT'; } }document.querySelectorAll('.ni').forEach(el=>el.addEventListener('click',()=>UI.go(el.dataset.page)));document.getElementById('mo').addEventListener('click',e=>{ if(e.target===e.currentTarget)UI.hideModal(); });const t=localStorage.getItem('ttu5_theme')||'dark'; applyTheme(t);document.getElementById('themeBtn').addEventListener('click',()=>{ const nv=localStorage.getItem('ttu5_theme')==='light'?'dark':'light'; localStorage.setItem('ttu5_theme',nv); applyTheme(nv); });document.getElementById('searchBox').addEventListener('input',e=>{ const q=e.target.value.toLowerCase(); document.querySelectorAll('.fl').forEach(el=>{ const m=q&&el.textContent.toLowerCase().includes(q); el.style.background=m?'var(--acg)':''; el.style.borderRadius=m?'3px':''; }); });document.addEventListener('keydown',e=>{if((e.ctrlKey||e.metaKey)&&e.key==='s'){e.preventDefault();DB.save().then(()=>UI.toast('Saved','💾'));}if((e.ctrlKey||e.metaKey)&&e.key==='z'&&!e.shiftKey){e.preventDefault();if(State.undo()){UI._render();UI.toast('Undo','↩️');}}if((e.ctrlKey||e.metaKey)&&(e.key==='y'||(e.key==='z'&&e.shiftKey))){e.preventDefault();if(State.redo()){UI._render();UI.toast('Redo','↪️');}}});State.on(inp=>{ document.getElementById('projName').textContent=inp.projectName||'Untitled'; document.getElementById('projType').textContent=inp.ttuType||'20MT'; });UI.go('dashboard');} catch(err){document.getElementById('content').innerHTML=<div style="text-align:center;padding:60px 20px;color:var(--mu)"><div style="font-size:36px">⚠️</div><div style="font-size:15px;font-weight:600;margin-bottom:5px">Initialization Error</div><div style="font-size:12px">${esc(err.message)}</div><div style="font-size:11px;margin-top:7px">Try Chrome or Firefox</div></div>;console.error(err);}}

function applyTheme(t){const L={'--bg':'#f0f4f8','--surf':'#fff','--card':'#fff','--inp':'#f7fafc','--hov':'#eef2f7','--bdr':'#d5dfe9','--div':'#eaf0f7','--tx':'#1a2332','--tx2':'#4a6080','--mu':'#8ba3bc'};const D={'--bg':'#0b1623','--surf':'#0f1e30','--card':'#132337','--inp':'#0a1520','--hov':'#19304a','--bdr':'#1e3350','--div':'#152035','--tx':'#dce8f5','--tx2':'#7a9ab8','--mu':'#3e607e'};const vars=t==='light'?L;Object.entries(vars).forEach(([k,v])=>document.documentElement.style.setProperty(k,v));document.getElementById('themeBtn').textContent=t==='light'?'🌙 Dark Mode':'☀️ Light Mode';}

boot();
