// Farming CBA Tool — Newcastle Business School (extended version with project module, soil benefits, and simulation)
(() => {
  // ---------- MODEL ----------
  const thisYear = new Date().getFullYear();
  const todayIso = new Date().toISOString().slice(0, 10);

  const model = {
    project: {
      name: "Nitrogen Optimisation Trial",
      lead: "Project lead name",
      team: "Farm Econ Team",
      analysts: "Farm Econ Team",
      organisation: "Newcastle Business School, The University of Newcastle",
      contactEmail: "frank.agbola@newcastle.edu.au",
      contactPhone: "",
      projectStartYear: thisYear,
      summary:
        "Test fertiliser and soil strategies to raise wheat yield and protein across 500 ha over 5 years.",
      lastUpdated: todayIso,
      objectives: "Increase profitability and soil health while managing risk.",
      goal: "Increase yield by 10% and protein by 0.5 p.p. on 500 ha within 3 years.",
      withProject: "Adopt optimised nitrogen timing and rates; improved soil management on 500 ha.",
      withoutProject:
        "Business-as-usual fertilisation; yield/protein unchanged; rising costs.",
      activities: [
        "On-farm trials at multiple sites",
        "Soil sampling and lab analysis",
        "Farmer engagement workshops"
      ],
      stakeholders: [
        "Grain growers",
        "Advisers and agronomists",
        "Industry bodies"
      ]
    },
    time: {
      startYear: thisYear,
      years: 10,
      projectStartYear: thisYear,
      discMode: "constant", // "constant" or "schedule"
      discBase: 7,
      discLow: 4,
      discHigh: 10,
      // Time-varying schedule in per cent
      discSchedule: {
        low:    { p1: 2, p2: 4, p3: 4, p4: 3, p5: 2 },
        def:    { p1: 4, p2: 7, p3: 7, p4: 6, p5: 5 },
        high:   { p1: 6, p2: 10, p3: 10, p4: 9, p5: 8 }
      },
      mirrFinance: 6,
      mirrReinvest: 4
    },
    outputsConfig: {
      systemType: "single", // "single" or "mixed"
      assumptions: ""
    },
    outputs: [
      { id: uid(), name: "Yield", unit: "t/ha", value: 300, source: "Input Directly" },
      { id: uid(), name: "Biomass", unit: "t/ha", value: 40, source: "Input Directly" },
      { id: uid(), name: "Antlers/harvest", unit: "per head", value: 15, source: "Input Directly" },
      { id: uid(), name: "Sold output", unit: "t", value: 320, source: "Input Directly" }
    ],
    treatments: [
      {
        id: uid(),
        name: "Control (current practice)",
        area: 300,
        replications: 1,
        adoption: 1.0,
        deltas: {},
        annualCost: 25,
        materialsCost: 0,
        servicesCost: 0,
        capitalCost: 0,
        constrained: true,
        source: "Farm Trials",
        isControl: true,
        useDepreciation: false,
        deprMethod: "sl",
        deprLife: 5,
        deprRate: 20
      },
      {
        id: uid(),
        name: "Optimised N (rate + timing)",
        area: 300,
        replications: 1,
        adoption: 0.8,
        deltas: {},
        annualCost: 45,
        materialsCost: 10,
        servicesCost: 5,
        capitalCost: 5000,
        constrained: true,
        source: "Farm Trials",
        isControl: false,
        useDepreciation: false,
        deprMethod: "sl",
        deprLife: 5,
        deprRate: 20
      }
    ],
    controlTreatmentId: null,
    benefits: [
      {
        id: uid(),
        label: "Improved soil structure (physical)",
        domain: "Soil health – physical",
        category: "C4",
        frequency: "Annual",
        startYear: thisYear,
        endYear: thisYear + 4,
        year: thisYear,
        unitValue: 0,
        quantity: 0,
        abatement: 0,
        annualAmount: 15000,
        growthPct: 0,
        linkAdoption: true,
        linkRisk: true,
        p0: 0,
        p1: 0,
        consequence: 0,
        notes: "Project-wide OPEX saving due to better trafficability"
      },
      {
        id: uid(),
        label: "Reduced risk of grain quality downgrades",
        domain: "Risk reduction",
        category: "C7",
        frequency: "Annual",
        startYear: thisYear,
        endYear: thisYear + 9,
        year: thisYear,
        unitValue: 0,
        quantity: 0,
        abatement: 0,
        annualAmount: 0,
        growthPct: 0,
        linkAdoption: true,
        linkRisk: false,
        p0: 0.10,
        p1: 0.07,
        consequence: 120000,
        notes: ""
      },
      {
        id: uid(),
        label: "Soil carbon asset uplift",
        domain: "Soil carbon",
        category: "C6",
        frequency: "Once",
        startYear: thisYear,
        endYear: thisYear,
        year: thisYear + 5,
        unitValue: 0,
        quantity: 0,
        abatement: 0,
        annualAmount: 50000,
        growthPct: 0,
        linkAdoption: false,
        linkRisk: true,
        p0: 0,
        p1: 0,
        consequence: 0,
        notes: ""
      }
    ],
    otherCosts: [
      {
        id: uid(),
        label: "Project management & M&E",
        type: "annual",
        costCategory: "labour",
        annual: 20000,
        startYear: thisYear,
        endYear: thisYear + 4,
        capital: 0,
        year: thisYear,
        constrained: true,
        useDepreciation: false,
        deprMethod: "sl",
        deprLife: 5,
        deprRate: 20
      }
    ],
    adoption: { base: 0.9, low: 0.6, high: 1.0 },
    risk: {
      base: 0.15,
      low: 0.05,
      high: 0.30,
      tech: 0.05,
      nonCoop: 0.04,
      socio: 0.02,
      fin: 0.03,
      man: 0.02
    },
    sim: {
      n: 1000,
      targetBCR: 2,
      bcrMode: "all",
      seed: null,
      results: { npv: [], bcr: [] },
      details: [],
      sensitivity: []
    }
  };

  function initTreatmentDeltas() {
    model.treatments.forEach(t => {
      model.outputs.forEach(o => {
        if (!(o.id in t.deltas)) t.deltas[o.id] = 0;
      });
    });
  }
  initTreatmentDeltas();

  if (!model.controlTreatmentId && model.treatments.length) {
    const ctrl = model.treatments.find(t => t.isControl) || model.treatments[0];
    model.controlTreatmentId = ctrl.id;
  }

  // ---------- UTIL ----------
  function uid() { return Math.random().toString(36).slice(2, 10); }
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const fmt = n =>
    (isFinite(n)
      ? (Math.abs(n) >= 1000
          ? n.toLocaleString(undefined, { maximumFractionDigits: 0 })
          : n.toLocaleString(undefined, { maximumFractionDigits: 2 }))
      : "—");
  const money = n => (isFinite(n) ? "$" + fmt(n) : "—");
  const percent = n => (isFinite(n) ? fmt(n) + "%" : "—");
  const slug = s => (s || "project").toLowerCase().replace(/[^a-z0-9]+/g,"_").replace(/^_|_$/g,"");
  const annuityFactor = (N, rPct) => {
    const r = rPct / 100;
    return r === 0 ? N : (1 - Math.pow(1 + r, -N)) / r;
  };
  const esc = s => (s ?? "").toString().replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));

  function discountRateForYear(absYear, scenario) {
    const sched = model.time.discSchedule || {};
    const table = scenario === "low" ? sched.low : scenario === "high" ? sched.high : sched.def;
    if (!table) return model.time.discBase;
    const y = absYear;
    if (y >= 2025 && y <= 2034) return table.p1;
    if (y >= 2035 && y <= 2044) return table.p2;
    if (y >= 2045 && y <= 2054) return table.p3;
    if (y >= 2055 && y <= 2064) return table.p4;
    if (y >= 2065 && y <= 2074) return table.p5;
    return model.time.discBase;
  }

  function presentValue(series, ratePct, scenarioLabel) {
    let pv = 0;
    if (model.time.discMode === "schedule" && scenarioLabel) {
      const base = model.time.startYear;
      for (let t = 0; t < series.length; t++) {
        const absYear = base + t;
        const rPct = t === 0 ? 0 : discountRateForYear(absYear, scenarioLabel);
        const r = rPct / 100;
        pv += series[t] / Math.pow(1 + r, t);
      }
      return pv;
    }
    const r = ratePct / 100;
    for (let t = 0; t < series.length; t++) {
      pv += series[t] / Math.pow(1 + r, t);
    }
    return pv;
  }

  function partialPresentValue(series, ratePct, T, scenarioLabel) {
    const slice = series.slice(0, T + 1);
    return presentValue(slice, ratePct, scenarioLabel);
  }

  function saveWorkbook(filename, wb) {
    try {
      if (window.XLSX && typeof XLSX.writeFile === "function") {
        XLSX.writeFile(wb, filename, { compression: true });
        return;
      }
    } catch(_) {}
    try {
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array", compression: true });
      const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = filename;
      document.body.appendChild(a); a.click(); setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove();},300);
    } catch (err) {
      alert("Download failed. Ensure SheetJS script is loaded.\n\n"+(err?.message||err));
    }
  }

  // ---------- CASHFLOWS ----------
  function buildCashflows({ forRate, adoptMul, risk, sim }) {
    const N = model.time.years;
    const baseYear = model.time.startYear;

    const priceFactor = sim?.priceFactor ?? 1;
    const treatCostFactor = sim?.treatCostFactor ?? 1;
    const inputCostFactor = sim?.inputCostFactor ?? 1;

    const benefitByYear = new Array(N + 1).fill(0);
    const costByYear = new Array(N + 1).fill(0);
    const constrainedCostByYear = new Array(N + 1).fill(0);

    // Treatments × Outputs
    let annualBenefit = 0;
    let treatAnnualCostForGM = 0;

    model.treatments.forEach(t => {
      const adopt = clamp(t.adoption * adoptMul, 0, 1);
      const rep = t.replications || 1;
      const area = (Number(t.area) || 0) * rep;

      let valuePerHa = 0;
      model.outputs.forEach(o => {
        const delta = Number(t.deltas[o.id]) || 0;
        const v = (Number(o.value) || 0) * priceFactor;
        valuePerHa += delta * v;
      });
      const benefit = valuePerHa * area * (1 - clamp(risk, 0, 1)) * adopt;

      const perHaCost =
        (Number(t.annualCost) || 0) +
        (Number(t.materialsCost) || 0) +
        (Number(t.servicesCost) || 0);
      const opCost = perHaCost * area * treatCostFactor;
      const cap = (Number(t.capitalCost) || 0) * inputCostFactor;

      annualBenefit += benefit;
      treatAnnualCostForGM += opCost;

      for (let year = 1; year <= N; year++) {
        costByYear[year] += opCost;
        if (t.constrained) constrainedCostByYear[year] += opCost;
      }

      if (cap > 0) {
        // Capital costs are treated as up-front economic costs at Year 0.
        costByYear[0] += cap;
        if (t.constrained) constrainedCostByYear[0] += cap;
      }
    });

    // Other project costs
    const otherAnnualByYear = new Array(N + 1).fill(0);
    const otherConstrAnnualByYear = new Array(N + 1).fill(0);
    let otherCapitalY0 = 0, otherConstrCapitalY0 = 0;

    model.otherCosts.forEach(c => {
      if (c.type === "annual") {
        const a = (Number(c.annual) || 0) * inputCostFactor;
        const sy = Number(c.startYear) || baseYear;
        const ey = Number(c.endYear) || sy;
        for (let y = sy; y <= ey; y++) {
          const idx = y - baseYear + 1;
          if (idx >= 1 && idx <= N) {
            otherAnnualByYear[idx] += a;
            if (c.constrained) otherConstrAnnualByYear[idx] += a;
          }
        }
      } else if (c.type === "capital") {
        const cap = (Number(c.capital) || 0) * inputCostFactor;
        const cy = Number(c.year) || baseYear;
        const idx = cy - baseYear;
        if (idx === 0) {
          otherCapitalY0 += cap;
          if (c.constrained) otherConstrCapitalY0 += cap;
        } else if (idx > 0 && idx <= N) {
          costByYear[idx] += cap;
          if (c.constrained) constrainedCostByYear[idx] += cap;
        }
      }
    });

    costByYear[0] += otherCapitalY0;
    constrainedCostByYear[0] += otherConstrCapitalY0;
    for (let t = 1; t <= N; t++) {
      costByYear[t] += otherAnnualByYear[t];
      constrainedCostByYear[t] += otherConstrAnnualByYear[t];
    }

    // Additional benefits
    const extra = additionalBenefitsSeries(N, baseYear, adoptMul, risk);
    for (let i = 0; i < extra.length; i++) benefitByYear[i] += extra[i];

    const cf = new Array(N + 1).fill(0).map((_, i) => (benefitByYear[i] - costByYear[i]));
    const annualGM = annualBenefit - treatAnnualCostForGM;
    return { benefitByYear, costByYear, constrainedCostByYear, cf, annualGM };
  }

  function additionalBenefitsSeries(N, baseYear, adoptMul, risk) {
    const series = new Array(N + 1).fill(0);
    model.benefits.forEach(b => {
      const cat = String(b.category || "").toUpperCase();
      const linkA = !!b.linkAdoption, linkR = !!b.linkRisk;
      const A = linkA ? clamp(adoptMul, 0, 1) : 1;
      const R = linkR ? (1 - clamp(risk, 0, 1)) : 1;
      const g = Number(b.growthPct) || 0;

      const addAnnual = (yearIndex, baseAmount, tFromStart) => {
        const grown = baseAmount * Math.pow(1 + g / 100, tFromStart);
        if (yearIndex >= 1 && yearIndex <= N) series[yearIndex] += grown * A * R;
      };
      const addOnce = (absYear, amount) => {
        const idx = absYear - baseYear + 1;
        if (idx >= 0 && idx <= N) series[idx] += amount * A * R;
      };

      const sy = Number(b.startYear) || baseYear;
      const ey = Number(b.endYear) || sy;
      const yr = Number(b.year) || sy;

      if (b.frequency === "Once" || cat === "C6") {
        const amount = Number(b.annualAmount) || 0;
        addOnce(yr, amount);
        return;
      }

      for (let y = sy; y <= ey; y++) {
        const idx = y - baseYear + 1;
        const tFromStart = y - sy;
        let amt = 0;
        switch (cat) {
          case "C1":
          case "C2":
          case "C3": {
            const v = Number(b.unitValue) || 0;
            const q = Number(cat === "C3" ? b.abatement : b.quantity) || 0;
            amt = v * q;
            break;
          }
          case "C4":
          case "C5":
          case "C8":
            amt = Number(b.annualAmount) || 0;
            break;
          case "C7": {
            const p0 = Number(b.p0) || 0, p1 = Number(b.p1) || 0;
            const c = Number(b.consequence) || 0;
            amt = Math.max(p0 - p1, 0) * c;
            break;
          }
          default:
            amt = 0;
        }
        addAnnual(idx, amt, tFromStart);
      }
    });
    return series;
  }

  function computeAll(rate, adoptMul, risk, bcrMode, discountScenario, simOptions) {
    const { benefitByYear, costByYear, constrainedCostByYear, cf, annualGM } =
      buildCashflows({ forRate: rate, adoptMul, risk, sim: simOptions });

    const pvBenefits = presentValue(benefitByYear, rate, discountScenario || null);
    const pvCosts = presentValue(costByYear, rate, discountScenario || null);
    const pvCostsConstrained = presentValue(constrainedCostByYear, rate, discountScenario || null);

    const npv = pvBenefits - pvCosts;
    const denom = (bcrMode === "constrained") ? pvCostsConstrained : pvCosts;
    const bcr = denom > 0 ? pvBenefits / denom : NaN;

    const irrVal = irr(cf);
    const mirrVal = mirr(cf, model.time.mirrFinance, model.time.mirrReinvest);
    const roi = pvCosts > 0 ? ((pvBenefits - pvCosts) / pvCosts) * 100 : NaN;
    const profitMargin = benefitByYear[1] > 0 ? (annualGM / benefitByYear[1]) * 100 : NaN;
    const pb = payback(cf, rate);

    return {
      pvBenefits, pvCosts, pvCostsConstrained,
      npv, bcr, irrVal, mirrVal, roi,
      annualGM, profitMargin, paybackYears: pb,
      benefitByYear, costByYear
    };
  }

  function irr(cf) {
    const hasPos = cf.some(v => v > 0);
    const hasNeg = cf.some(v => v < 0);
    if (!hasPos || !hasNeg) return NaN;
    let lo = -0.99, hi = 5.0;
    const npvAt = r => cf.reduce((acc, v, t) => acc + v / Math.pow(1 + r, t), 0);
    let nLo = npvAt(lo), nHi = npvAt(hi);
    if (nLo * nHi > 0) {
      for (let k = 0; k < 20 && nLo * nHi > 0; k++) {
        hi *= 1.5; nHi = npvAt(hi);
      }
      if (nLo * nHi > 0) return NaN;
    }
    for (let i = 0; i < 80; i++) {
      const mid = (lo + hi) / 2; const nMid = npvAt(mid);
      if (Math.abs(nMid) < 1e-8) return mid * 100;
      if (nLo * nMid <= 0) { hi = mid; nHi = nMid; } else { lo = mid; nLo = nMid; }
    }
    return ((lo + hi) / 2) * 100;
  }

  function mirr(cf, financeRatePct, reinvestRatePct) {
    const n = cf.length - 1;
    const fr = financeRatePct / 100, rr = reinvestRatePct / 100;
    let pvNeg = 0, fvPos = 0;
    for (let t = 0; t <= n; t++) {
      const v = cf[t];
      if (v < 0) pvNeg += v / Math.pow(1 + fr, t);
      if (v > 0) fvPos += v * Math.pow(1 + rr, n - t);
    }
    if (pvNeg === 0) return NaN;
    const mirrVal = Math.pow(-fvPos / pvNeg, 1 / n) - 1;
    return mirrVal * 100;
  }

  function payback(cf, ratePct) {
    let cum = 0;
    const r = ratePct / 100;
    for (let t = 0; t < cf.length; t++) {
      cum += cf[t] / Math.pow(1 + r, t);
      if (cum >= 0) return t;
    }
    return null;
  }

  // ---------- DOM ----------
  const $ = sel => document.querySelector(sel);
  const $$ = sel => Array.from(document.querySelectorAll(sel));
  const num = sel => +(document.querySelector(sel)?.value || 0);
  const setVal = (sel, text) => (document.querySelector(sel).textContent = text);

  function switchTab(target){
    $$("#tabs button").forEach(b =>
      b.classList.toggle("active", b.dataset.tab === target)
    );
    $$(".tab-panel").forEach(p =>
      p.classList.toggle("show", p.id === `tab-${target}`)
    );
    if (target === "distribution") drawHists();
    if (target === "report") calcAndRender();
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function initTabs() {
    document.addEventListener("click", (e) => {
      const navBtn = e.target.closest("#tabs button[data-tab]");
      const jumpBtn = e.target.closest("[data-tab-jump]");
      const target = navBtn?.dataset.tab || jumpBtn?.dataset.tabJump;
      if (!target) return;
      switchTab(target);
    });
  }

  function initActions(){
    document.addEventListener("click", (e) => {
      if (e.target.closest("#recalc, #getResults, [data-action='recalc']")) {
        e.preventDefault(); calcAndRender();
      }
      if (e.target.closest("#runSim, [data-action='run-sim']")) {
        e.preventDefault(); runSimulation();
      }
      if (e.target.closest("#runSensitivity")) {
        e.preventDefault(); runStructuredSensitivity();
      }
    });
  }

  // ---------- BIND + RENDER FORMS ----------
  function setBasicsFieldsFromModel() {
    $("#projectName").value = model.project.name || "";
    $("#projectLead").value = model.project.lead || "";
    $("#projectTeam").value = model.project.team || "";
    $("#analystNames").value = model.project.analysts || "";
    $("#projectSummary").value = model.project.summary || "";
    $("#lastUpdated").value = model.project.lastUpdated || "";
    $("#projectObjectives").value = model.project.objectives || "";
    $("#projectGoal").value = model.project.goal || "";
    $("#withProject").value = model.project.withProject || "";
    $("#withoutProject").value = model.project.withoutProject || "";
    $("#organisation").value = model.project.organisation || "";
    $("#contactEmail").value = model.project.contactEmail || "";
    $("#contactPhone").value = model.project.contactPhone || "";
    $("#projectStartYear").value = model.project.projectStartYear || "";
    $("#timeProjectStartYear").value = model.time.projectStartYear || "";

    const acts = model.project.activities || [];
    for (let i = 1; i <= 10; i++) {
      const el = $("#activity"+i);
      if (el) el.value = acts[i-1] || "";
    }
    const stks = model.project.stakeholders || [];
    for (let i = 1; i <= 10; i++) {
      const el = $("#stakeholder"+i);
      if (el) el.value = stks[i-1] || "";
    }

    $("#startYear").value = model.time.startYear;
    $("#years").value = model.time.years;
    $("#discBase").value = model.time.discBase;
    $("#discLow").value = model.time.discLow;
    $("#discHigh").value = model.time.discHigh;
    $("#discMode").value = model.time.discMode;
    $("#mirrFinance").value = model.time.mirrFinance;
    $("#mirrReinvest").value = model.time.mirrReinvest;

    const sch = model.time.discSchedule || {};
    const low = sch.low || {}, def = sch.def || {}, high = sch.high || {};
    $("#disc_low_p1").value = low.p1 ?? 2;
    $("#disc_low_p2").value = low.p2 ?? 4;
    $("#disc_low_p3").value = low.p3 ?? 4;
    $("#disc_low_p4").value = low.p4 ?? 3;
    $("#disc_low_p5").value = low.p5 ?? 2;
    $("#disc_def_p1").value = def.p1 ?? 4;
    $("#disc_def_p2").value = def.p2 ?? 7;
    $("#disc_def_p3").value = def.p3 ?? 7;
    $("#disc_def_p4").value = def.p4 ?? 6;
    $("#disc_def_p5").value = def.p5 ?? 5;
    $("#disc_high_p1").value = high.p1 ?? 6;
    $("#disc_high_p2").value = high.p2 ?? 10;
    $("#disc_high_p3").value = high.p3 ?? 10;
    $("#disc_high_p4").value = high.p4 ?? 9;
    $("#disc_high_p5").value = high.p5 ?? 8;

    $("#adoptBase").value = model.adoption.base;
    $("#adoptLow").value = model.adoption.low;
    $("#adoptHigh").value = model.adoption.high;

    $("#riskBase").value = model.risk.base;
    $("#riskLow").value = model.risk.low;
    $("#riskHigh").value = model.risk.high;
    $("#rTech").value = model.risk.tech;
    $("#rNonCoop").value = model.risk.nonCoop;
    $("#rSocio").value = model.risk.socio;
    $("#rFin").value = model.risk.fin;
    $("#rMan").value = model.risk.man;

    $("#simN").value = model.sim.n;
    $("#targetBCR").value = model.sim.targetBCR;
    $("#bcrMode").value = model.sim.bcrMode;
    $("#simBcrTargetLabel").textContent = model.sim.targetBCR;

    $("#cropSystemType").value = model.outputsConfig.systemType || "single";
    $("#outputAssumptions").value = model.outputsConfig.assumptions || "";
  }

  function bindBasics() {
    setBasicsFieldsFromModel();

    $("#calcCombinedRisk").addEventListener("click", () => {
      const r = 1 - (1 - num("#rTech")) * (1 - num("#rNonCoop")) * (1 - num("#rSocio")) * (1 - num("#rFin")) * (1 - num("#rMan"));
      $("#combinedRiskOut").textContent = `Combined: ${(r * 100).toFixed(2)}%`;
      $("#riskBase").value = r.toFixed(3);
      model.risk.base = r;
      calcAndRender();
    });

    $("#addCost").addEventListener("click", () => {
      const c = {
        id: uid(),
        label: "New cost",
        type: "annual",
        costCategory: "other",
        annual: 0,
        startYear: model.time.startYear,
        endYear: model.time.startYear,
        capital: 0,
        year: model.time.startYear,
        constrained: true,
        useDepreciation: false,
        deprMethod: "sl",
        deprLife: 5,
        deprRate: 20
      };
      model.otherCosts.push(c);
      renderCosts();
      calcAndRender();
    });

    document.addEventListener("input", e => {
      const id = e.target.id;
      if (!id) return;

      const actMatch = id.match(/^activity(\d+)$/);
      if (actMatch) {
        const idx = Number(actMatch[1]) - 1;
        model.project.activities[idx] = e.target.value;
        return;
      }
      const stkMatch = id.match(/^stakeholder(\d+)$/);
      if (stkMatch) {
        const idx = Number(stkMatch[1]) - 1;
        model.project.stakeholders[idx] = e.target.value;
        return;
      }

      switch (id) {
        case "projectName": model.project.name = e.target.value; break;
        case "projectLead": model.project.lead = e.target.value; break;
        case "projectTeam": model.project.team = e.target.value; break;
        case "analystNames": model.project.analysts = e.target.value; break;
        case "projectSummary": model.project.summary = e.target.value; break;
        case "lastUpdated": model.project.lastUpdated = e.target.value; break;
        case "projectObjectives": model.project.objectives = e.target.value; break;
        case "projectGoal": model.project.goal = e.target.value; break;
        case "withProject": model.project.withProject = e.target.value; break;
        case "withoutProject": model.project.withoutProject = e.target.value; break;
        case "organisation": model.project.organisation = e.target.value; break;
        case "contactEmail": model.project.contactEmail = e.target.value; break;
        case "contactPhone": model.project.contactPhone = e.target.value; break;
        case "projectStartYear":
          model.project.projectStartYear = +e.target.value;
          model.time.projectStartYear = +e.target.value;
          $("#timeProjectStartYear").value = e.target.value;
          break;
        case "timeProjectStartYear":
          model.time.projectStartYear = +e.target.value;
          model.project.projectStartYear = +e.target.value;
          $("#projectStartYear").value = e.target.value;
          break;

        case "startYear": model.time.startYear = +e.target.value; break;
        case "years": model.time.years = +e.target.value; break;
        case "discBase": model.time.discBase = +e.target.value; break;
        case "discLow": model.time.discLow = +e.target.value; break;
        case "discHigh": model.time.discHigh = +e.target.value; break;
        case "discMode": model.time.discMode = e.target.value; break;
        case "mirrFinance": model.time.mirrFinance = +e.target.value; break;
        case "mirrReinvest": model.time.mirrReinvest = +e.target.value; break;

        case "adoptBase": model.adoption.base = +e.target.value; break;
        case "adoptLow": model.adoption.low = +e.target.value; break;
        case "adoptHigh": model.adoption.high = +e.target.value; break;

        case "riskBase": model.risk.base = +e.target.value; break;
        case "riskLow": model.risk.low = +e.target.value; break;
        case "riskHigh": model.risk.high = +e.target.value; break;
        case "rTech": model.risk.tech = +e.target.value; break;
        case "rNonCoop": model.risk.nonCoop = +e.target.value; break;
        case "rSocio": model.risk.socio = +e.target.value; break;
        case "rFin": model.risk.fin = +e.target.value; break;
        case "rMan": model.risk.man = +e.target.value; break;

        case "simN": model.sim.n = +e.target.value; break;
        case "targetBCR": model.sim.targetBCR = +e.target.value; $("#simBcrTargetLabel").textContent = e.target.value; break;
        case "bcrMode": model.sim.bcrMode = e.target.value; break;
        case "randSeed": model.sim.seed = e.target.value ? +e.target.value : null; break;

        case "cropSystemType": model.outputsConfig.systemType = e.target.value; break;
        case "outputAssumptions": model.outputsConfig.assumptions = e.target.value; break;
      }

      if (id.startsWith("disc_")) {
        const parts = id.split("_"); // disc_low_p1
        const which = parts[1]; // low/def/high
        const p = parts[2]; // p1..p5
        const val = +e.target.value;
        if (!model.time.discSchedule) model.time.discSchedule = { low:{},def:{},high:{} };
        if (which === "low") model.time.discSchedule.low[p] = val;
        if (which === "def") model.time.discSchedule.def[p] = val;
        if (which === "high") model.time.discSchedule.high[p] = val;
      }

      calcAndRenderDebounced();
    });

    $("#saveProject").addEventListener("click", () => {
      const data = JSON.stringify(model, null, 2);
      downloadFile(`cba_${(model.project.name || "project").replace(/\s+/g, "_")}.json`, data, "application/json");
    });
    $("#loadProject").addEventListener("click", () => $("#loadFile").click());
    $("#loadFile").addEventListener("change", async e => {
      const file = e.target.files?.[0]; if (!file) return;
      const text = await file.text();
      try {
        const obj = JSON.parse(text);
        Object.assign(model, obj);
        initTreatmentDeltas();
        renderAll();
        setBasicsFieldsFromModel();
        calcAndRender();
      } catch {
        alert("Invalid JSON file.");
      } finally { e.target.value = ""; }
    });

    $("#exportCsv").addEventListener("click", exportAllCsv);
    $("#exportCsvFoot").addEventListener("click", exportAllCsv);
    $("#exportPdf").addEventListener("click", exportPdf);
    $("#exportPdfFoot").addEventListener("click", exportPdf);

    $("#parseExcel").addEventListener("click", handleParseExcel);
    $("#importExcel").addEventListener("click", commitExcelToModel);

    $("#downloadTemplate").addEventListener("click", downloadExcelTemplate);
    $("#downloadSample").addEventListener("click", downloadSampleDataset);

    $("#startBtn")?.addEventListener("click", () => switchTab("project"));
  }

  // ---------- RENDERERS ----------
  function renderOutputs() {
    const root = $("#outputsList"); root.innerHTML = "";
    model.outputs.forEach(o => {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>Output: ${esc(o.name)}</h4>
        <div class="row-6">
          <div class="field"><label>Name</label><input value="${esc(o.name)}" data-k="name" data-id="${o.id}"/></div>
          <div class="field"><label>Unit</label><input value="${esc(o.unit)}" data-k="unit" data-id="${o.id}"/></div>
          <div class="field"><label>Value ($/unit)</label><input type="number" step="0.01" value="${o.value}" data-k="value" data-id="${o.id}"/></div>
          <div class="field"><label>Source</label>
            <select data-k="source" data-id="${o.id}">
              ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"].map(s => `<option ${s===o.source?"selected":""}>${s}</option>`).join("")}
            </select>
          </div>
          <div class="field"><label>&nbsp;</label><button class="danger" data-del-output="${o.id}">Remove</button></div>
        </div>
        <div class="kv"><small class="muted">id:</small> <code>${o.id}</code></div>
      `;
      root.appendChild(el);
    });
    root.oninput = onOutputEdit;
    root.onclick = onOutputDelete;
  }
  function onOutputEdit(e) {
    const k = e.target.dataset.k, id = e.target.dataset.id; if (!k || !id) return;
    const o = model.outputs.find(x => x.id === id); if (!o) return;
    if (k === "value") o[k] = +e.target.value; else o[k] = e.target.value;
    model.treatments.forEach(t => { if (!(id in t.deltas)) t.deltas[id] = 0; });
    renderTreatments(); renderDatabaseTags(); calcAndRenderDebounced();
  }
  function onOutputDelete(e) {
    const id = e.target.dataset.delOutput; if (!id) return;
    if (!confirm("Remove this output metric?")) return;
    model.outputs = model.outputs.filter(o => o.id !== id);
    model.treatments.forEach(t => delete t.deltas[id]);
    renderOutputs(); renderTreatments(); renderDatabaseTags(); calcAndRender();
  }
  $("#addOutput")?.addEventListener("click", () => {
    const id = uid();
    model.outputs.push({ id, name: "Custom Output", unit: "unit", value: 0, source: "Input Directly" });
    model.treatments.forEach(t => (t.deltas[id] = 0));
    renderOutputs(); renderTreatments(); renderDatabaseTags();
  });

  function renderTreatments() {
    const root = $("#treatmentsList"); root.innerHTML = "";
    model.treatments.forEach(t => {
      const perHa =
        (Number(t.annualCost) || 0) +
        (Number(t.materialsCost) || 0) +
        (Number(t.servicesCost) || 0);
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>🚜 Treatment: ${esc(t.name)}</h4>
        <div class="row">
          <div class="field"><label>Name</label><input value="${esc(t.name)}" data-tk="name" data-id="${t.id}" /></div>
          <div class="field"><label>Area (ha, per replication)</label><input type="number" step="0.01" value="${t.area}" data-tk="area" data-id="${t.id}" /></div>
          <div class="field"><label>Replications</label><input type="number" min="1" step="1" value="${t.replications||1}" data-tk="replications" data-id="${t.id}" /></div>
          <div class="field"><label>Adoption (0–1)</label><input type="number" min="0" max="1" step="0.01" value="${t.adoption}" data-tk="adoption" data-id="${t.id}" /></div>
          <div class="field"><label>Source</label>
            <select data-tk="source" data-id="${t.id}">
              ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"].map(s => `<option ${s===t.source?"selected":""}>${s}</option>`).join("")}
            </select>
          </div>
          <div class="field"><label>Annual base cost ($/ha)</label><input type="number" step="0.01" value="${t.annualCost}" data-tk="annualCost" data-id="${t.id}" /></div>
          <div class="field"><label>Materials cost ($/ha)</label><input type="number" step="0.01" value="${t.materialsCost||0}" data-tk="materialsCost" data-id="${t.id}" /></div>
          <div class="field"><label>Services cost ($/ha)</label><input type="number" step="0.01" value="${t.servicesCost||0}" data-tk="servicesCost" data-id="${t.id}" /></div>
          <div class="field"><label>Total variable cost ($/ha)</label><input value="${perHa.toFixed(2)}" disabled /></div>
          <div class="field"><label>Capital cost ($, Year 0)</label><input type="number" step="0.01" value="${t.capitalCost}" data-tk="capitalCost" data-id="${t.id}" /></div>
          <div class="field"><label>Constrained?</label>
            <select data-tk="constrained" data-id="${t.id}">
              <option value="true" ${t.constrained?"selected":""}>Yes</option>
              <option value="false" ${!t.constrained?"selected":""}>No</option>
            </select>
          </div>
          <div class="field">
            <label>Control treatment?</label>
            <input type="radio" name="controlTreatment" data-control-id="${t.id}" ${model.controlTreatmentId===t.id || t.isControl ? "checked" : ""}/>
          </div>
          <div class="field"><label>&nbsp;</label><button class="danger" data-del-treatment="${t.id}">Remove</button></div>
        </div>

        <h5>Output deltas (per ha)</h5>
        <div class="row">
          ${model.outputs.map(o => `
            <div class="field">
              <label>${esc(o.name)} (${esc(o.unit)})</label>
              <input type="number" step="0.0001" value="${t.deltas[o.id] ?? 0}" data-td="${o.id}" data-id="${t.id}" />
            </div>
          `).join("")}
        </div>
        <div class="kv"><small class="muted">id:</small> <code>${t.id}</code></div>
      `;
      root.appendChild(el);
    });
    root.oninput = e => {
      const id = e.target.dataset.id; if (!id) return;
      const t = model.treatments.find(x => x.id === id); if (!t) return;
      const tk = e.target.dataset.tk;
      if (tk) {
        if (tk === "constrained") t[tk] = e.target.value === "true";
        else if (tk === "name" || tk === "source") t[tk] = e.target.value;
        else if (tk === "replications") t[tk] = Math.max(1, +e.target.value||1);
        else t[tk] = +e.target.value;
      }
      const td = e.target.dataset.td;
      if (td) t.deltas[td] = +e.target.value;
      calcAndRenderDebounced();
    };
    root.onclick = e => {
      const id = e.target.dataset.delTreatment;
      if (id) {
        if (!confirm("Remove this treatment?")) return;
        model.treatments = model.treatments.filter(x => x.id !== id);
        if (model.controlTreatmentId === id) {
          model.controlTreatmentId = model.treatments[0]?.id || null;
        }
        renderTreatments(); renderDatabaseTags(); calcAndRender();
        return;
      }
      const ctrlId = e.target.dataset.controlId;
      if (ctrlId) {
        model.controlTreatmentId = ctrlId;
        model.treatments.forEach(t => t.isControl = (t.id === ctrlId));
      }
    };
  }
  $("#addTreatment")?.addEventListener("click", () => {
    if (model.treatments.length >= 64) {
      alert("For clarity, the tool supports up to 64 treatments.");
      return;
    }
    const t = {
      id: uid(),
      name: "New treatment",
      area: 100,
      replications: 1,
      adoption: 0.7,
      deltas: {},
      annualCost: 0,
      materialsCost: 0,
      servicesCost: 0,
      capitalCost: 0,
      constrained: true,
      source: "Input Directly",
      isControl: false,
      useDepreciation: false,
      deprMethod: "sl",
      deprLife: 5,
      deprRate: 20
    };
    model.outputs.forEach(o => { t.deltas[o.id] = 0; });
    model.treatments.push(t);
    renderTreatments(); renderDatabaseTags();
  });

  function renderBenefits() {
    const root = $("#benefitsList"); root.innerHTML = "";
    model.benefits.forEach(b => {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>🌱 ${esc(b.label || "Benefit")}</h4>
        <div class="row-6">
          <div class="field"><label>Label</label><input value="${esc(b.label||"")}" data-bk="label" data-id="${b.id}" /></div>
          <div class="field"><label>Domain</label>
            <select data-bk="domain" data-id="${b.id}">
              ${[
                "Soil health – chemical",
                "Soil health – physical",
                "Soil health – biological",
                "Soil carbon",
                "Soil pH by depth",
                "Soil nutrients by depth",
                "Soil properties by treatment",
                "Cost savings",
                "Water retention",
                "Risk reduction",
                "Other"
              ].map(d => `<option ${d===(b.domain||"Other")?"selected":""}>${d}</option>`).join("")}
            </select>
          </div>
          <div class="field"><label>Category</label>
            <select data-bk="category" data-id="${b.id}">
              ${["C1","C2","C3","C4","C5","C6","C7","C8"].map(c=>`<option ${c===b.category?"selected":""}>${c}</option>`).join("")}
            </select>
          </div>
          <div class="field"><label>Frequency</label>
            <select data-bk="frequency" data-id="${b.id}">
              <option ${b.frequency==="Annual"?"selected":""}>Annual</option>
              <option ${b.frequency==="Once"?"selected":""}>Once</option>
            </select>
          </div>
          <div class="field"><label>Start year</label><input type="number" value="${b.startYear||model.time.startYear}" data-bk="startYear" data-id="${b.id}" /></div>
          <div class="field"><label>End year</label><input type="number" value="${b.endYear||model.time.startYear}" data-bk="endYear" data-id="${b.id}" /></div>
          <div class="field"><label>Once year</label><input type="number" value="${b.year||model.time.startYear}" data-bk="year" data-id="${b.id}" /></div>
        </div>

        <div class="row-6">
          <div class="field"><label>Unit value ($)</label><input type="number" step="0.01" value="${b.unitValue||0}" data-bk="unitValue" data-id="${b.id}" /></div>
          <div class="field"><label>Quantity</label><input type="number" step="0.01" value="${b.quantity||0}" data-bk="quantity" data-id="${b.id}" /></div>
          <div class="field"><label>Abatement</label><input type="number" step="0.01" value="${b.abatement||0}" data-bk="abatement" data-id="${b.id}" /></div>
          <div class="field"><label>Annual amount ($)</label><input type="number" step="0.01" value="${b.annualAmount||0}" data-bk="annualAmount" data-id="${b.id}" /></div>
          <div class="field"><label>Growth (%/yr)</label><input type="number" step="0.01" value="${b.growthPct||0}" data-bk="growthPct" data-id="${b.id}" /></div>
          <div class="field"><label>Link adoption?</label>
            <select data-bk="linkAdoption" data-id="${b.id}">
              <option value="true" ${b.linkAdoption?"selected":""}>Yes</option>
              <option value="false" ${!b.linkAdoption?"selected":""}>No</option>
            </select>
          </div>
          <div class="field"><label>Link risk?</label>
            <select data-bk="linkRisk" data-id="${b.id}">
              <option value="true" ${b.linkRisk?"selected":""}>Yes</option>
              <option value="false" ${!b.linkRisk?"selected":""}>No</option>
            </select>
          </div>
        </div>

        <div class="row-6">
          <div class="field"><label>P0 (baseline prob)</label><input type="number" step="0.001" value="${b.p0||0}" data-bk="p0" data-id="${b.id}" /></div>
          <div class="field"><label>P1 (with-project prob)</label><input type="number" step="0.001" value="${b.p1||0}" data-bk="p1" data-id="${b.id}" /></div>
          <div class="field"><label>Consequence ($)</label><input type="number" step="0.01" value="${b.consequence||0}" data-bk="consequence" data-id="${b.id}" /></div>
          <div class="field"><label>Notes</label><input value="${esc(b.notes||"")}" data-bk="notes" data-id="${b.id}" /></div>
          <div class="field"><label>&nbsp;</label><button class="danger" data-del-benefit="${b.id}">Remove</button></div>
        </div>
      `;
      root.appendChild(el);
    });

    root.oninput = e => {
      const id = e.target.dataset.id; if (!id) return;
      const b = model.benefits.find(x => x.id === id); if (!b) return;
      const k = e.target.dataset.bk;
      if (!k) return;
      if (k === "label" || k === "category" || k === "frequency" || k === "notes" || k === "domain") b[k] = e.target.value;
      else if (k === "linkAdoption" || k === "linkRisk") b[k] = e.target.value === "true";
      else b[k] = +e.target.value;
      calcAndRenderDebounced();
    };
    root.onclick = e => {
      const id = e.target.dataset.delBenefit; if (!id) return;
      if (!confirm("Remove this benefit item?")) return;
      model.benefits = model.benefits.filter(x => x.id !== id);
      renderBenefits(); calcAndRender();
    };
  }
  $("#addBenefit")?.addEventListener("click", () => {
    model.benefits.push({
      id: uid(), label: "New benefit", domain: "Other", category: "C4", frequency: "Annual",
      startYear: model.time.startYear, endYear: model.time.startYear, year: model.time.startYear,
      unitValue: 0, quantity: 0, abatement: 0, annualAmount: 0, growthPct: 0,
      linkAdoption: true, linkRisk: true, p0: 0, p1: 0, consequence: 0, notes: ""
    });
    renderBenefits();
  });

  function renderCosts() {
    const root = $("#costsList"); root.innerHTML = "";
    model.otherCosts.forEach(c => {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>💰 Cost item: ${esc(c.label)}</h4>
        <div class="row-6">
          <div class="field"><label>Label</label><input value="${esc(c.label)}" data-ck="label" data-id="${c.id}" /></div>
          <div class="field"><label>Type</label>
            <select data-ck="type" data-id="${c.id}">
              <option value="annual" ${c.type==="annual"?"selected":""}>Annual</option>
              <option value="capital" ${c.type==="capital"?"selected":""}>Capital</option>
            </select>
          </div>
          <div class="field"><label>Cost category</label>
            <select data-ck="costCategory" data-id="${c.id}">
              <option value="capital" ${c.costCategory==="capital"?"selected":""}>Capital</option>
              <option value="labour" ${c.costCategory==="labour"?"selected":""}>Labour</option>
              <option value="materials" ${c.costCategory==="materials"?"selected":""}>Materials</option>
              <option value="services" ${c.costCategory==="services"?"selected":""}>Services</option>
              <option value="other" ${!c.costCategory || c.costCategory==="other"?"selected":""}>Other production</option>
            </select>
          </div>
          <div class="field"><label>Annual ($/yr)</label><input type="number" step="0.01" value="${c.annual ?? 0}" data-ck="annual" data-id="${c.id}" /></div>
          <div class="field"><label>Start year</label><input type="number" value="${c.startYear ?? model.time.startYear}" data-ck="startYear" data-id="${c.id}" /></div>
          <div class="field"><label>End year</label><input type="number" value="${c.endYear ?? model.time.startYear}" data-ck="endYear" data-id="${c.id}" /></div>
          <div class="field"><label>Capital ($)</label><input type="number" step="0.01" value="${c.capital ?? 0}" data-ck="capital" data-id="${c.id}" /></div>
          <div class="field"><label>Capital year</label><input type="number" value="${c.year ?? model.time.startYear}" data-ck="year" data-id="${c.id}" /></div>
          <div class="field"><label>Constrained?</label>
            <select data-ck="constrained" data-id="${c.id}">
              <option value="true" ${c.constrained?"selected":""}>Yes</option>
              <option value="false" ${!c.constrained?"selected":""}>No</option>
            </select>
          </div>
          <div class="field"><label>&nbsp;</label><button class="danger" data-del-cost="${c.id}">Remove</button></div>
        </div>
      `;
      root.appendChild(el);
    });
    root.oninput = e => {
      const id = e.target.dataset.id, k = e.target.dataset.ck; if (!id || !k) return;
      const c = model.otherCosts.find(x => x.id === id); if (!c) return;
      if (k === "label" || k === "type" || k==="costCategory") c[k] = e.target.value;
      else if (k === "constrained") c[k] = e.target.value === "true";
      else c[k] = +e.target.value;
      calcAndRenderDebounced();
    };
    root.onclick = e => {
      const id = e.target.dataset.delCost; if (!id) return;
      if (!confirm("Remove this cost item?")) return;
      model.otherCosts = model.otherCosts.filter(x => x.id !== id);
      renderCosts(); calcAndRender();
    };
  }

  function renderDatabaseTags() {
    const outRoot = $("#dbOutputs"); outRoot.innerHTML = "";
    model.outputs.forEach(o => {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <div class="row-2">
          <div class="field"><label>${esc(o.name)} (${esc(o.unit)})</label></div>
          <div class="field">
            <label>Source</label>
            <select data-db-out="${o.id}">
              ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"].map(s => `<option ${s===o.source?"selected":""}>${s}</option>`).join("")}
            </select>
          </div>
        </div>`;
      outRoot.appendChild(el);
    });
    outRoot.onchange = e => {
      const id = e.target.dataset.dbOut;
      const o = model.outputs.find(x => x.id === id);
      if (o) o.source = e.target.value;
    };

    const tRoot = $("#dbTreatments"); tRoot.innerHTML = "";
    model.treatments.forEach(t => {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <div class="row-2">
          <div class="field"><label>${esc(t.name)}</label></div>
          <div class="field">
            <label>Source</label>
            <select data-db-t="${t.id}">
              ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"].map(s => `<option ${s===t.source?"selected":""}>${s}</option>`).join("")}
            </select>
          </div>
        </div>`;
      tRoot.appendChild(el);
    });
    tRoot.onchange = e => {
      const id = e.target.dataset.dbT;
      const t = model.treatments.find(x => x.id === id);
      if (t) t.source = e.target.value;
    };
  }

  function renderTreatmentSummary(rate, adoptMul, risk) {
    const root = $("#treatmentSummary"); root.innerHTML = "";
    const rows = [];
    model.treatments.forEach(t => {
      let valuePerHa = 0;
      model.outputs.forEach(o => (valuePerHa += (Number(t.deltas[o.id]) || 0) * (Number(o.value) || 0)));
      const rep = t.replications || 1;
      const area = (Number(t.area) || 0) * rep;
      const adopt = clamp(t.adoption * adoptMul, 0, 1);
      const annualBen = valuePerHa * area * (1 - clamp(risk, 0, 1)) * adopt;
      const perHaCost =
        (Number(t.annualCost) || 0) +
        (Number(t.materialsCost) || 0) +
        (Number(t.servicesCost) || 0);
      const annualCost = perHaCost * area;
      const cap = Number(t.capitalCost) || 0;
      const pvBen = annualBen * annuityFactor(model.time.years, rate);
      const pvCost = cap + annualCost * annuityFactor(model.time.years, rate);
      const bcr = pvCost > 0 ? pvBen / pvCost : NaN;
      const npv = pvBen - pvCost;

      rows.push({
        name: t.name,
        isControl: t.id === model.controlTreatmentId,
        area,
        adopt,
        annualBen,
        annualCost,
        pvBen,
        pvCost,
        bcr,
        npv
      });
    });

    rows.sort((a,b) => b.npv - a.npv);

    rows.forEach(r => {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <div class="row-6">
          <div class="field"><label>Rank</label><div class="metric"><div class="value">${rows.indexOf(r)+1}</div></div></div>
          <div class="field"><label>Treatment</label><div class="metric"><div class="value">${esc(r.name)}${r.isControl?" (control)":""}</div></div></div>
          <div class="field"><label>Area (all reps)</label><div class="metric"><div class="value">${fmt(r.area)} ha</div></div></div>
          <div class="field"><label>Adoption</label><div class="metric"><div class="value">${fmt(r.adopt)}</div></div></div>
          <div class="field"><label>Annual benefit</label><div class="metric"><div class="value">${money(r.annualBen)}</div></div></div>
          <div class="field"><label>Annual cost</label><div class="metric"><div class="value">${money(r.annualCost)}</div></div></div>
          <div class="field"><label>PV benefit</label><div class="metric"><div class="value">${money(r.pvBen)}</div></div></div>
          <div class="field"><label>PV cost</label><div class="metric"><div class="value">${money(r.pvCost)}</div></div></div>
          <div class="field"><label>BCR</label><div class="metric"><div class="value">${isFinite(r.bcr)?fmt(r.bcr):"—"}</div></div></div>
          <div class="field"><label>NPV</label><div class="metric"><div class="value">${money(r.npv)}</div></div></div>
        </div>`;
      root.appendChild(el);
    });
  }

  function renderTimeProjections(rate, adoptMul, risk) {
    const root = $("#projectionSummary");
    if (!root) return;
    const N = model.time.years;
    const horizons = [5,10,15,20,25].filter(h => h <= N);
    if (!horizons.length) {
      root.innerHTML = "<p class='small muted'>Increase the analysis horizon to at least 5 years to see time-based projections.</p>";
      return;
    }

    const all = computeAll(rate, adoptMul, risk, model.sim.bcrMode,
      model.time.discMode==="schedule" ? "def" : null, null);
    const { benefitByYear, costByYear } = all;

    const rows = horizons.map(h => {
      const T = h;
      const pvB = partialPresentValue(benefitByYear, rate, T, model.time.discMode==="schedule"?"def":null);
      const pvC = partialPresentValue(costByYear, rate, T, model.time.discMode==="schedule"?"def":null);
      const npv = pvB - pvC;
      const bcr = pvC>0 ? pvB/pvC : NaN;
      return { horizon:T, pvB, pvC, npv, bcr };
    });

    let html = "<table class='projection-table'><thead><tr><th>Horizon (years)</th><th>PV benefits (with-project)</th><th>PV costs (with-project)</th><th>NPV (with-project)</th><th>BCR (with-project)</th></tr></thead><tbody>";
    rows.forEach(r => {
      html += `<tr>
        <td>${r.horizon}</td>
        <td>${money(r.pvB)}</td>
        <td>${money(r.pvC)}</td>
        <td>${money(r.npv)}</td>
        <td>${isFinite(r.bcr)?fmt(r.bcr):"—"}</td>
      </tr>`;
    });
    html += "</tbody></table>";
    root.innerHTML = html;
  }

  function renderAll() {
    renderOutputs();
    renderTreatments();
    renderBenefits();
    renderDatabaseTags();
    renderCosts();
  }

  // ---------- MAIN CALC / REPORT ----------
  function calcAndRender() {
    const rate = model.time.discBase;
    const adoptMul = model.adoption.base;
    const risk = model.risk.base;

    const all = computeAll(
      rate,
      adoptMul,
      risk,
      model.sim.bcrMode,
      model.time.discMode==="schedule" ? "def" : null,
      null
    );

    const {
      pvBenefits, pvCosts, npv, bcr, irrVal,
      mirrVal, roi, annualGM, profitMargin,
      paybackYears
    } = all;

    setVal("#pvBenefits", money(pvBenefits));
    setVal("#pvCosts", money(pvCosts));
    const npvEl = $("#npv");
    npvEl.textContent = money(npv);
    npvEl.className = "value " + (npv >= 0 ? "positive" : "negative");
    setVal("#bcr", isFinite(bcr) ? fmt(bcr) : "—");
    setVal("#irr", isFinite(irrVal) ? percent(irrVal) : "—");
    setVal("#mirr", isFinite(mirrVal) ? percent(mirrVal) : "—");
    setVal("#roi", isFinite(roi) ? percent(roi) : "—");
    setVal("#grossMargin", money(annualGM));
    setVal("#profitMargin", isFinite(profitMargin) ? percent(profitMargin) : "—");
    setVal("#payback", paybackYears != null ? paybackYears : "Not reached");

    renderTreatmentSummary(rate, adoptMul, risk);
    renderTimeProjections(rate, adoptMul, risk);
    $("#simBcrTargetLabel").textContent = model.sim.targetBCR;
  }

  let debTimer = null;
  function calcAndRenderDebounced() {
    clearTimeout(debTimer);
    debTimer = setTimeout(calcAndRender, 120);
  }

  // ---------- MONTE CARLO ----------
  function rng(seed) {
    let t = (seed || Math.floor(Math.random() * 2 ** 31)) >>> 0;
    return () => {
      t += 0x6D2B79F5;
      let x = t;
      x = Math.imul(x ^ (x >>> 15), 1 | x);
      x ^= x + Math.imul(x ^ (x >>> 7), 61 | x);
      return ((x ^ (x >>> 14)) >>> 0) / 4294967296;
    };
  }

  function triangular(r, a, c, b) {
    const F = (c - a) / (b - a);
    if (r < F) return a + Math.sqrt(r * (b - a) * (c - a));
    return b - Math.sqrt((1 - r) * (b - a) * (b - c));
  }

  async function runSimulation() {
    $("#simStatus").textContent = "Running…";
    await new Promise(r => setTimeout(r));
    const N = model.sim.n;
    const seed = model.sim.seed;
    const rand = rng(seed ?? undefined);

    const discLow = model.time.discLow, discBase = model.time.discBase, discHigh = model.time.discHigh;
    const adoptLow = model.adoption.low, adoptBase = model.adoption.base, adoptHigh = model.adoption.high;
    const riskLow = model.risk.low, riskBase = model.risk.base, riskHigh = model.risk.high;

    const levels = [0.05,0.10,0.15,0.20,0.25];

    const npvs = new Array(N);
    const bcrs = new Array(N);
    const details = [];

    for (let i = 0; i < N; i++) {
      const r1 = rand(), r2 = rand(), r3 = rand();
      const disc = triangular(r1, discLow, discBase, discHigh);
      const adoptMul = clamp(triangular(r2, adoptLow, adoptBase, adoptHigh), 0, 1);
      const risk = clamp(triangular(r3, riskLow, riskBase, riskHigh), 0, 1);

      const Lp = levels[Math.floor(rand()*levels.length)];
      const Lc = levels[Math.floor(rand()*levels.length)];
      const Li = levels[Math.floor(rand()*levels.length)];
      const priceFactor = 1 + (rand()<0.5 ? -1:1) * Lp;
      const treatCostFactor = 1 + (rand()<0.5 ? -1:1) * Lc;
      const inputCostFactor = 1 + (rand()<0.5 ? -1:1) * Li;

      const { pvBenefits, pvCosts, bcr, npv } = computeAll(
        disc,
        adoptMul,
        risk,
        model.sim.bcrMode,
        null,
        { priceFactor, treatCostFactor, inputCostFactor }
      );

      npvs[i] = npv;
      bcrs[i] = bcr;
      details.push({ run: i + 1, discount: disc, adoption: adoptMul, risk, pvBenefits, pvCosts, npv, bcr });
    }

    model.sim.results = { npv: npvs, bcr: bcrs };
    model.sim.details = details;
    $("#simStatus").textContent = "Done.";
    renderSimulationResults();
    drawHists();
  }

  function renderSimulationResults() {
    const { npv, bcr } = model.sim.results;
    if (!npv?.length) return;
    const sortedNpv = [...npv].sort((a, b) => a - b);
    const validBcr = bcr.filter(x => isFinite(x));
    const sortedBcr = [...validBcr].sort((a, b) => a - b);
    const N = npv.length, NB = sortedBcr.length;

    const stats = arr => ({
      min: arr[0],
      max: arr[arr.length - 1],
      mean: arr.reduce((a, c) => a + c, 0) / arr.length,
      median: arr.length ? (arr[Math.floor((arr.length - 1) / 2)] + arr[Math.ceil((arr.length - 1) / 2)]) / 2 : NaN
    });

    const sN = stats(sortedNpv);
    const sB = stats(sortedBcr.length ? sortedBcr : [NaN]);

    setVal("#simNpvMin", money(sN.min));
    setVal("#simNpvMax", money(sN.max));
    setVal("#simNpvMean", money(sN.mean));
    setVal("#simNpvMedian", money(sN.median));
    const pN = npv.filter(x => x > 0).length / N * 100;
    setVal("#simNpvProb", fmt(pN) + "%");

    setVal("#simBcrMin", isFinite(sB.min) ? fmt(sB.min) : "—");
    setVal("#simBcrMax", isFinite(sB.max) ? fmt(sB.max) : "—");
    setVal("#simBcrMean", isFinite(sB.mean) ? fmt(sB.mean) : "—");
    setVal("#simBcrMedian", isFinite(sB.median) ? fmt(sB.median) : "—");
    const pB1 = NB ? validBcr.filter(x => x > 1).length / NB * 100 : 0;
    setVal("#simBcrProb1", fmt(pB1) + "%");
    const tgt = model.sim.targetBCR;
    const pBt = NB ? validBcr.filter(x => x > tgt).length / NB * 100 : 0;
    setVal("#simBcrProbTarget", fmt(pBt) + "%");
  }

  function drawHist(canvasId, data, bins = 24, labelFmt = v => v.toFixed(0)) {
    const canvas = document.getElementById(canvasId);
    if (!canvas || !data?.length) return;
    const ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    const min = Math.min(...data), max = Math.max(...data);
    const padL = 54, padR = 14, padT = 10, padB = 34;
    const W = canvas.width - padL - padR, H = canvas.height - padT - padB;

    const counts = new Array(bins).fill(0);
    const span = (max - min) || 1e-9;
    data.forEach(v => {
      let idx = Math.floor(((v - min) / span) * bins);
      if (idx < 0) idx = 0;
      if (idx >= bins) idx = bins - 1;
      counts[idx]++;
    });
    const maxC = Math.max(...counts) || 1;

    ctx.strokeStyle = "#3c6a52";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(padL, padT);
    ctx.lineTo(padL, padT + H);
    ctx.lineTo(padL + W, padT + H);
    ctx.stroke();

    for (let i = 0; i < bins; i++) {
      const x = padL + (i * W) / bins + 1;
      const h = (counts[i] / maxC) * (H - 2);
      const y = padT + H - h;
      ctx.fillStyle = "rgba(116, 209, 140, 0.45)";
      ctx.fillRect(x, y, (W / bins) - 2, h);
    }

    ctx.fillStyle = "#c9efd6";
    ctx.font = "12px system-ui";
    ctx.textAlign = "center";
    const lbls = [min, (min + max) / 2, max];
    [0, 0.5, 1].forEach((p, i) => {
      const x = padL + p * W;
      ctx.fillText(labelFmt(lbls[i]), x, padT + H + 20);
    });
  }
  function drawHists() {
    const { npv, bcr } = model.sim.results;
    if (npv?.length) drawHist("histNpv", npv, 24, v => money(v));
    if (bcr?.length) drawHist("histBcr", bcr.filter(x => isFinite(x)), 24, v => v.toFixed(2));
  }

  // ---------- STRUCTURED SENSITIVITY ----------
  function runStructuredSensitivity() {
    const usePrices = $("#sensPrices")?.checked;
    const useTreatCosts = $("#sensTreatCosts")?.checked;
    const useInputCosts = $("#sensInputCosts")?.checked;
    const useDisc = $("#sensDiscounts")?.checked;
    const useAdopt = $("#sensAdoption")?.checked;
    const useRisk = $("#sensRisk")?.checked;

    const params = [];
    if (usePrices) params.push("Prices");
    if (useTreatCosts) params.push("Treatment costs");
    if (useInputCosts) params.push("Input costs");
    if (useDisc) params.push("Discount rate");
    if (useAdopt) params.push("Adoption");
    if (useRisk) params.push("Risk");

    const levels = [0.05,0.10,0.15,0.20,0.25];
    const baseRate = model.time.discBase;
    const discountScenario = model.time.discMode==="schedule" ? "def" : null;

    const rows = [];

    params.forEach(p => {
      levels.forEach(L => {
        ["-","+"].forEach(sign => {
          const mult = 1 + (sign === "-" ? -L : L);
          let rate = baseRate;
          let adoptMul = model.adoption.base;
          let risk = model.risk.base;
          let priceFactor = 1;
          let treatCostFactor = 1;
          let inputCostFactor = 1;

          if (p === "Prices") priceFactor = mult;
          if (p === "Treatment costs") treatCostFactor = mult;
          if (p === "Input costs") inputCostFactor = mult;
          if (p === "Discount rate") rate = baseRate * mult;
          if (p === "Adoption") adoptMul = clamp(model.adoption.base * mult,0,1);
          if (p === "Risk") risk = clamp(model.risk.base * mult,0,1);

          const res = computeAll(
            rate,
            adoptMul,
            risk,
            model.sim.bcrMode,
            discountScenario,
            { priceFactor, treatCostFactor, inputCostFactor }
          );

          rows.push({
            param: p,
            shock: sign + (L*100).toFixed(0) + "%",
            rate,
            adoption: adoptMul,
            risk,
            pvB: res.pvBenefits,
            pvC: res.pvCosts,
            npv: res.npv,
            bcr: res.bcr
          });
        });
      });
    });

    model.sim.sensitivity = rows;

    const root = $("#sensitivityTable");
    if (!rows.length) {
      root.innerHTML = "<p class='small muted'>Select at least one parameter to vary.</p>";
      return;
    }

    let html = "<div class='card'><h3>Sensitivity grid results</h3>";
    html += "<table class='sensitivity-table'><thead><tr><th>Parameter</th><th>Shock</th><th>Discount rate (%)</th><th>Adoption</th><th>Risk</th><th>PV benefits</th><th>PV costs</th><th>NPV</th><th>BCR</th></tr></thead><tbody>";
    rows.forEach(r => {
      html += `<tr>
        <td>${esc(r.param)}</td>
        <td>${esc(r.shock)}</td>
        <td>${fmt(r.rate)}</td>
        <td>${fmt(r.adoption)}</td>
        <td>${fmt(r.risk)}</td>
        <td>${money(r.pvB)}</td>
        <td>${money(r.pvC)}</td>
        <td>${money(r.npv)}</td>
        <td>${isFinite(r.bcr)?fmt(r.bcr):"—"}</td>
      </tr>`;
    });
    html += "</tbody></table></div>";
    root.innerHTML = html;
  }

  // ---------- EXPORTS ----------
  function toCsv(rows) {
    return rows.map(r => r.map(v => {
      const s = (v ?? "").toString();
      if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
      return s;
    }).join(",")).join("\n");
  }
  function downloadFile(filename, text, mime="text/csv") {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([text], { type: mime }));
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  function buildSummaryForCsv() {
    const rate = model.time.discBase;
    const adoptMul = model.adoption.base;
    const risk = model.risk.base;
    const all = computeAll(
      rate,
      adoptMul,
      risk,
      model.sim.bcrMode,
      model.time.discMode==="schedule"?"def":null,
      null
    );

    return {
      meta: {
        name: model.project.name,
        analysts: model.project.analysts,
        organisation: model.project.organisation,
        contact: model.project.contactEmail,
        updated: model.project.lastUpdated
      },
      params: {
        startYear: model.time.startYear,
        years: model.time.years,
        discountBase: model.time.discBase,
        discountLow: model.time.discLow,
        discountHigh: model.time.discHigh,
        mirrFinance: model.time.mirrFinance,
        mirrReinvest: model.time.mirrReinvest,
        adoptionBase: model.adoption.base,
        riskBase: model.risk.base,
        bcrMode: model.sim.bcrMode
      },
      results: all
    };
  }

  function exportAllCsv() {
    const s = buildSummaryForCsv();
    const summaryRows = [
      ["Project", s.meta.name],
      ["Project lead", model.project.lead],
      ["Analysts", s.meta.analysts],
      ["Organisation", s.meta.organisation],
      ["Contact", s.meta.contact],
      ["Last Updated", s.meta.updated],
      [],
      ["Start Year", s.params.startYear],
      ["Years", s.params.years],
      ["Discount Rate (Base)", s.params.discountBase],
      ["Discount Rate (Low)", s.params.discountLow],
      ["Discount Rate (High)", s.params.discountHigh],
      ["MIRR Finance %", s.params.mirrFinance],
      ["MIRR Reinvest %", s.params.mirrReinvest],
      ["Adoption Multiplier", s.params.adoptionBase],
      ["Risk (overall)", s.params.riskBase],
      ["BCR Mode", s.params.bcrMode],
      [],
      ["PV Benefits", s.results.pvBenefits],
      ["PV Costs", s.results.pvCosts],
      ["NPV", s.results.npv],
      ["BCR", s.results.bcr],
      ["IRR %", s.results.irrVal],
      ["MIRR %", s.results.mirrVal],
      ["ROI %", s.results.roi],
      ["Gross Margin (annual)", s.results.annualGM],
      ["Gross Profit Margin %", s.results.profitMargin],
      ["Payback (years)", s.results.paybackYears ?? "Not reached"]
    ];
    downloadFile(`cba_summary_${slug(s.meta.name)}.csv`, toCsv(summaryRows));

    const treatHeader = ["Treatment","Control?","Area(ha)","Adoption","Annual Benefit","Annual Cost","PV Benefit","PV Cost","BCR","NPV"];
    const treatRows = [treatHeader];
    const rate = model.time.discBase, adoptMul = model.adoption.base, risk = model.risk.base;
    model.treatments.forEach(t => {
      let valuePerHa = 0;
      model.outputs.forEach(o => valuePerHa += ((+t.deltas[o.id]||0) * (+o.value||0)));
      const rep = t.replications || 1;
      const area = (t.area||0) * rep;
      const adopt = clamp(t.adoption * adoptMul, 0, 1);
      const annBen = valuePerHa * area * (1 - risk) * adopt;
      const perHaCost =
        (Number(t.annualCost) || 0) +
        (Number(t.materialsCost) || 0) +
        (Number(t.servicesCost) || 0);
      const annCost = perHaCost * area;
      const pvB = annBen * annuityFactor(model.time.years, rate);
      const pvC = (t.capitalCost||0) + annCost * annuityFactor(model.time.years, rate);
      const bcr = pvC>0 ? pvB/pvC : "";
      const npv = pvB - pvC;
      treatRows.push([t.name, t.id===model.controlTreatmentId?"Yes":"No", area, adopt, annBen, annCost, pvB, pvC, bcr, npv]);
    });
    downloadFile(`cba_treatments_${slug(s.meta.name)}.csv`, toCsv(treatRows));

    const benRows = [["Label","Domain","Category","Frequency","StartYear","EndYear","Year","UnitValue","Quantity","Abatement","AnnualAmount","GrowthPct","LinkAdoption","LinkRisk","P0","P1","Consequence","Notes"]];
    model.benefits.forEach(b => benRows.push([b.label,b.domain,b.category,b.frequency,b.startYear,b.endYear,b.year,b.unitValue,b.quantity,b.abatement,b.annualAmount,b.growthPct,b.linkAdoption,b.linkRisk,b.p0,b.p1,b.consequence,b.notes]));
    downloadFile(`cba_benefits_${slug(s.meta.name)}.csv`, toCsv(benRows));

    const outRows = [["Output","Unit","$/unit","Source","Id"]];
    model.outputs.forEach(o => outRows.push([o.name,o.unit,o.value,o.source,o.id]));
    downloadFile(`cba_outputs_${slug(s.meta.name)}.csv`, toCsv(outRows));

    const { npv, bcr } = model.sim.results;
    if (npv?.length) {
      const validBcr = bcr.filter(x => isFinite(x));
      const stats = arr => {
        const a = [...arr].sort((x,y)=>x-y);
        const N = a.length;
        const med = (a[Math.floor((N-1)/2)] + a[Math.ceil((N-1)/2)]) / 2;
        const mean = a.reduce((u,v)=>u+v,0) / N;
        return { min:a[0], max:a[N-1], mean, median:med };
      };
      const sN = stats(npv);
      const sB = validBcr.length ? stats(validBcr) : {min:"",max:"",mean:"",median:""};
      const pN = npv.filter(x => x > 0).length / npv.length * 100;
      const tgt = model.sim.targetBCR;
      const pB1 = validBcr.filter(x => x > 1).length / (validBcr.length||1) * 100;
      const pBt = validBcr.filter(x => x > tgt).length / (validBcr.length||1) * 100;

      const simRows = [
        ["N", model.sim.n],
        ["BCR Mode", model.sim.bcrMode],
        ["NPV Min", sN.min],["NPV Max", sN.max],["NPV Mean", sN.mean],["NPV Median", sN.median],["Pr(NPV>0)%", pN],
        [],
        ["BCR Min", sB.min],["BCR Max", sB.max],["BCR Mean", sB.mean],["BCR Median", sB.median],["Pr(BCR>1)%", pB1],["Pr(BCR>Target)%", pBt]
      ];
      downloadFile(`cba_simulation_summary_${slug(s.meta.name)}.csv`, toCsv(simRows));

      const rawRows = [["run","discount","adoption","risk","pvBenefits","pvCosts","npv","bcr"]];
      model.sim.details.forEach(d => rawRows.push([d.run,d.discount,d.adoption,d.risk,d.pvBenefits,d.pvCosts,d.npv,d.bcr]));
      downloadFile(`cba_simulation_raw_${slug(s.meta.name)}.csv`, toCsv(rawRows));
    } else {
      const simRows = [["N", model.sim.n],["BCR Mode", model.sim.bcrMode],["Note","Run Monte Carlo to populate results."]];
      downloadFile(`cba_simulation_summary_${slug(s.meta.name)}.csv`, toCsv(simRows));
    }
  }

  function exportPdf() {
    drawHists();

    const npvCan = document.getElementById("histNpv");
    const bcrCan = document.getElementById("histBcr");
    const npvImg = (npvCan && npvCan.width) ? npvCan.toDataURL("image/png") : null;
    const bcrImg = (bcrCan && bcrCan.width) ? bcrCan.toDataURL("image/png") : null;

    const s = buildSummaryForCsv();

    const trRows = model.treatments.map(t => {
      let valuePerHa = 0;
      model.outputs.forEach(o => valuePerHa += ((+t.deltas[o.id]||0) * (+o.value||0)));
      const rep = t.replications || 1;
      const area = (t.area||0) * rep;
      const adopt = clamp(t.adoption * model.adoption.base, 0, 1);
      const annBen = valuePerHa * area * (1 - model.risk.base) * adopt;
      const perHaCost =
        (Number(t.annualCost) || 0) +
        (Number(t.materialsCost) || 0) +
        (Number(t.servicesCost) || 0);
      const annCost = perHaCost * area;
      const pvB = annBen * annuityFactor(model.time.years, model.time.discBase);
      const pvC = (t.capitalCost||0) + annCost * annuityFactor(model.time.years, model.time.discBase);
      const bcr = pvC>0 ? pvB/pvC : NaN;
      const npv = pvB - pvC;
      return `<tr>
        <td>${esc(t.name)}${t.id===model.controlTreatmentId?" (control)":""}</td><td>${fmt(area)}</td><td>${fmt(adopt)}</td>
        <td>${money(annBen)}</td><td>${money(annCost)}</td>
        <td>${money(pvB)}</td><td>${money(pvC)}</td>
        <td>${isFinite(bcr)?fmt(bcr):"—"}</td><td>${money(npv)}</td>
      </tr>`;
    }).join("");

    const benRows = model.benefits.map(b => `
      <tr><td>${esc(b.label)}</td><td>${esc(b.domain||"")}</td><td>${b.category}</td><td>${b.frequency}</td>
      <td>${b.startYear||""}</td><td>${b.endYear||""}</td><td>${b.year||""}</td>
      <td>${b.unitValue||""}</td><td>${b.quantity||""}</td><td>${b.abatement||""}</td>
      <td>${b.annualAmount||""}</td><td>${b.growthPct||""}</td>
      <td>${b.linkAdoption?"Yes":"No"}</td><td>${b.linkRisk?"Yes":"No"}</td>
      <td>${b.p0||""}</td><td>${b.p1||""}</td><td>${b.consequence||""}</td>
      <td>${esc(b.notes||"")}</td></tr>
    `).join("");

    const now = new Date().toLocaleString();

    const win = window.open("", "_blank");
    win.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>${esc(s.meta.name)} — CBA Report</title>
      <style>${getPrintCss()}</style></head><body>
      <div class="print-report">
        <div class="print-header">
          <div class="print-logo">🚜🌾</div>
          <div>
            <h1>${esc(s.meta.name)} <span class="print-badge">CBA Report</span></h1>
            <div class="print-small muted">Generated ${esc(now)}</div>
          </div>
        </div>

        <div class="print-cols">
          <div>
            <h2>Project</h2>
            <table>
              <tr><th>Lead</th><td>${esc(model.project.lead||"")}</td></tr>
              <tr><th>Team</th><td>${esc(model.project.team||"")}</td></tr>
              <tr><th>Analysts</th><td>${esc(s.meta.analysts)}</td></tr>
              <tr><th>Organisation</th><td>${esc(s.meta.organisation)}</td></tr>
              <tr><th>Contact</th><td><a href="mailto:${esc(s.meta.contact)}">${esc(s.meta.contact)}</a></td></tr>
              <tr><th>Last Updated</th><td>${esc(s.meta.updated)}</td></tr>
            </table>
          </div>
          <div>
            <h2>Parameters</h2>
            <table>
              <tr><th>Start Year</th><td>${s.params.startYear}</td></tr>
              <tr><th>Years</th><td>${s.params.years}</td></tr>
              <tr><th>Discount (L/B/H)</th><td>${s.params.discountLow}% / ${s.params.discountBase}% / ${s.params.discountHigh}%</td></tr>
              <tr><th>MIRR (Finance/Reinvest)</th><td>${s.params.mirrFinance}% / ${s.params.mirrReinvest}%</td></tr>
              <tr><th>Adoption Multiplier</th><td>${s.params.adoptionBase}</td></tr>
              <tr><th>Risk (overall)</th><td>${s.params.riskBase}</td></tr>
              <tr><th>BCR Mode</th><td>${esc(s.params.bcrMode)}</td></tr>
            </table>
          </div>
        </div>

        <h2>Economic indicators</h2>
        <table>
          <tr><th>PV Benefits</th><td>${money(s.results.pvBenefits)}</td></tr>
          <tr><th>PV Costs</th><td>${money(s.results.pvCosts)}</td></tr>
          <tr><th>NPV</th><td>${money(s.results.npv)}</td></tr>
          <tr><th>BCR</th><td>${isFinite(s.results.bcr)?fmt(s.results.bcr):"—"}</td></tr>
          <tr><th>IRR</th><td>${isFinite(s.results.irrVal)?percent(s.results.irrVal):"—"}</td></tr>
          <tr><th>MIRR</th><td>${isFinite(s.results.mirrVal)?percent(s.results.mirrVal):"—"}</td></tr>
          <tr><th>ROI</th><td>${isFinite(s.results.roi)?percent(s.results.roi):"—"}</td></tr>
          <tr><th>Gross Margin (annual)</th><td>${money(s.results.annualGM)}</td></tr>
          <tr><th>Gross Profit Margin</th><td>${isFinite(s.results.profitMargin)?percent(s.results.profitMargin):"—"}</td></tr>
          <tr><th>Payback (years)</th><td>${s.results.paybackYears ?? "Not reached"}</td></tr>
        </table>

        <h2>Treatments</h2>
        <table>
          <thead><tr>
            <th>Treatment</th><th>Area</th><th>Adoption</th><th>Annual Benefit</th><th>Annual Cost</th>
            <th>PV Benefit</th><th>PV Cost</th><th>BCR</th><th>NPV</th>
          </tr></thead>
          <tbody>${trRows}</tbody>
        </table>

        <h2>Additional benefits</h2>
        <table>
          <thead><tr>
            <th>Label</th><th>Domain</th><th>Cat</th><th>Freq</th><th>Start</th><th>End</th><th>Year</th>
            <th>UnitValue</th><th>Qty</th><th>Abatement</th><th>Annual</th><th>Growth%</th>
            <th>Adopt?</th><th>Risk?</th><th>P0</th><th>P1</th><th>Consequence</th><th>Notes</th>
          </tr></thead>
          <tbody>${benRows}</tbody>
        </table>

        <h2>Simulation highlights</h2>
        <div class="print-cols">
          <div>${npvImg ? `<img src="${npvImg}" style="width:100%;border:1px solid #ddd;border-radius:8px" />` : "<div class='muted'>NPV histogram not available.</div>"}</div>
          <div>${bcrImg ? `<img src="${bcrImg}" style="width:100%;border:1px solid #ddd;border-radius:8px" />` : "<div class='muted'>BCR histogram not available.</div>"}</div>
        </div>

        <hr />
        <div class="print-small muted">
          Newcastle Business School • The University of Newcastle • Contact: <a href="mailto:${esc(model.project.contactEmail)}">${esc(model.project.contactEmail)}</a>
        </div>
      </div>
    </body></html>`);
    win.document.close();
    setTimeout(() => { win.focus(); win.print(); }, 300);
  }

  function getPrintCss() {
    return `
      body{background:#fff;margin:0}
      .print-report{font:13px/1.5 system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#111;padding:24px;max-width:900px;margin:0 auto}
      .print-header{display:grid;grid-template-columns:60px 1fr;gap:12px;align-items:center;margin-bottom:6px}
      .print-logo{width:60px;height:60px;border-radius:14px;display:grid;place-items:center;font-size:26px;
        background:linear-gradient(135deg,#9be2ad,#ffd77e);border:1px solid #eee}
      .print-report h1{font-size:20px;margin:0 0 6px}
      .print-report h2{font-size:16px;margin:14px 0 6px}
      .print-report table{border-collapse:collapse;width:100%;margin:6px 0}
      .print-report th,.print-report td{border:1px solid #ddd;padding:6px;text-align:left}
      .print-report .muted{color:#555}
      .print-cols{display:grid;grid-template-columns:1fr 1fr;gap:12px}
      .print-small{font-size:12px}
      .print-badge{display:inline-block;border:1px solid #ddd;border-radius:8px;padding:2px 8px;margin-left:6px}
      @media print {.print-report img{page-break-inside:avoid}}
    `;
  }

  // ---------- EXCEL: TEMPLATE + PARSE + SAMPLE ----------
  let parsedExcel = null;

  async function handleParseExcel() {
    const file = $("#excelFile").files?.[0];
    const status = $("#loadStatus");
    const alertBox = $("#validation");
    const preview = $("#preview");
    parsedExcel = null;
    alertBox.classList.remove("show");
    alertBox.innerHTML = "";
    preview.innerHTML = "";
    $("#importExcel").disabled = true;

    if (!file) { status.textContent = "Select an Excel/CSV file first."; return; }
    status.textContent = "Parsing…";

    try {
      const buf = await file.arrayBuffer();
      let wb;
      if (file.name.toLowerCase().endsWith(".csv")) {
        const csvTxt = new TextDecoder().decode(new Uint8Array(buf));
        const ws = XLSX.utils.csv_to_sheet(csvTxt);
        wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, $("#csvSheetType").value || "Outputs");
      } else {
        wb = XLSX.read(buf, { type: "array" });
      }

      const getSheet = (name) => {
        const s = wb.Sheets[name];
        if (!s
      const getSheet = (name) => {
        const s = wb.Sheets[name];
        if (!s) return null;
        return XLSX.utils.sheet_to_json(s, { defval: "" });
      };

      // For CSV uploads we only have one sheet, named by csvSheetType
      // For XLSX we expect sheets called "Outputs", "Treatments", "Benefits", "OtherCosts"
      const outputs = getSheet("Outputs");
      const treatments = getSheet("Treatments");
      const benefits = getSheet("Benefits");
      const costs = getSheet("OtherCosts");

      if (
        (!outputs || !outputs.length) &&
        (!treatments || !treatments.length) &&
        (!benefits || !benefits.length) &&
        (!costs || !costs.length)
      ) {
        status.textContent = "No recognised sheets found (Outputs, Treatments, Benefits, OtherCosts).";
        return;
      }

      parsedExcel = { outputs, treatments, benefits, costs };

      const buildPreviewTable = (title, rows) => {
        if (!rows || !rows.length) return "";
        const cols = Object.keys(rows[0] || {});
        if (!cols.length) return "";
        let html = `<h4>${esc(title)}</h4><div class="table-wrapper"><table class="preview-table"><thead><tr>`;
        cols.forEach(c => { html += `<th>${esc(c)}</th>`; });
        html += "</tr></thead><tbody>";
        rows.slice(0, 8).forEach(r => {
          html += "<tr>";
          cols.forEach(c => {
            const v = r[c];
            html += `<td>${esc(v == null ? "" : String(v))}</td>`;
          });
          html += "</tr>";
        });
        if (rows.length > 8) {
          html += `<tr><td colspan="${cols.length}" class="muted small">… ${rows.length - 8} more rows not shown</td></tr>`;
        }
        html += "</tbody></table></div>";
        return html;
      };

      let html = "";
      html += buildPreviewTable("Outputs", outputs);
      html += buildPreviewTable("Treatments", treatments);
      html += buildPreviewTable("Benefits", benefits);
      html += buildPreviewTable("Other costs", costs);
      preview.innerHTML = html;

      status.textContent = "Parsed successfully. Review the preview and click “Import from Excel” to load.";
      $("#importExcel").disabled = false;
    } catch (err) {
      console.error(err);
      status.textContent = "Failed to parse file.";
      alertBox.classList.add("show");
      alertBox.innerHTML = `<p class="error">Parsing error: ${esc(err?.message || err)}</p>`;
    }
  }

  function commitExcelToModel() {
    const status = $("#loadStatus");
    const alertBox = $("#validation");
    alertBox.classList.remove("show");
    alertBox.innerHTML = "";

    if (!parsedExcel) {
      alertBox.classList.add("show");
      alertBox.innerHTML = "<p class='error'>Parse an Excel or CSV file first.</p>";
      return;
    }

    const { outputs, treatments, benefits, costs } = parsedExcel;

    // ----- Outputs -----
    if (outputs && outputs.length) {
      model.outputs = outputs.map((r, idx) => {
        const id = uid();
        const name =
          r.Name ?? r.Output ?? r.output ?? r.name ?? `Output ${idx + 1}`;
        const unit = r.Unit ?? r.unit ?? "";
        const valueRaw =
          r["$/unit"] ?? r.Value ?? r.value ?? r.Price ?? r.price ?? 0;
        const value = Number(valueRaw) || 0;
        const source =
          r.Source ?? r.source ?? "Input Directly";

        return {
          id,
          name: String(name),
          unit: String(unit),
          value,
          source: String(source)
        };
      });

      // Rebuild treatment deltas structure to align with new outputs
      model.treatments.forEach(t => {
        const newDeltas = {};
        model.outputs.forEach(o => {
          newDeltas[o.id] = t.deltas?.[o.id] ?? 0;
        });
        t.deltas = newDeltas;
      });
    }

    // ----- Treatments -----
    if (treatments && treatments.length) {
      const newTreatments = treatments.map((r, idx) => {
        const id = uid();
        const name =
          r.Treatment ?? r.Name ?? r.name ?? `Treatment ${idx + 1}`;
        const area = Number(r.Area ?? r.area ?? 0) || 0;
        const replications = Number(r.Replications ?? r.replications ?? 1) || 1;
        const adoption = Number(r.Adoption ?? r.adoption ?? 0.7) || 0;
        const annualCost =
          Number(
            r.AnnualCost ??
            r["AnnualCost($/ha)"] ??
            r["Annual cost ($/ha)"] ??
            r.annualCost ??
            0
          ) || 0;
        const materialsCost =
          Number(r.MaterialsCost ?? r.materialsCost ?? 0) || 0;
        const servicesCost =
          Number(r.ServicesCost ?? r.servicesCost ?? 0) || 0;
        const capitalCost =
          Number(r.CapitalCost ?? r.capitalCost ?? 0) || 0;

        const constrainedRaw =
          r.Constrained ?? r.constrained ?? "Yes";
        const constrainedStr = String(constrainedRaw).toLowerCase();
        const constrained =
          constrainedStr === "true" ||
          constrainedStr === "yes" ||
          constrainedStr === "y" ||
          constrainedStr === "1";

        const source =
          r.Source ?? r.source ?? "Input Directly";

        const t = {
          id,
          name: String(name),
          area,
          replications,
          adoption,
          deltas: {},
          annualCost,
          materialsCost,
          servicesCost,
          capitalCost,
          constrained,
          source: String(source),
          isControl: false,
          useDepreciation: false,
          deprMethod: "sl",
          deprLife: 5,
          deprRate: 20
        };

        // Map any delta columns of form d:OutputName or Delta:OutputName
        if (model.outputs && model.outputs.length) {
          model.outputs.forEach(o => {
            const k1 = `d:${o.name}`;
            const k2 = `delta:${o.name}`;
            const k3 = `Delta_${o.name}`;
            const v =
              Number(r[k1] ?? r[k2] ?? r[k3] ?? 0) || 0;
            t.deltas[o.id] = v;
          });
        }

        return t;
      });

      model.treatments = newTreatments;
      if (model.treatments.length) {
        model.controlTreatmentId = model.treatments[0].id;
        model.treatments.forEach((t, i) => {
          t.isControl = (i === 0);
        });
      }
    }

    // ----- Benefits -----
    if (benefits && benefits.length) {
      model.benefits = benefits.map(r => {
        const linkAdoptRaw = r.LinkAdoption ?? r.linkAdoption ?? "true";
        const linkAdoptStr = String(linkAdoptRaw).toLowerCase();
        const linkAdoption =
          linkAdoptStr === "true" ||
          linkAdoptStr === "yes" ||
          linkAdoptStr === "y" ||
          linkAdoptStr === "1";

        const linkRiskRaw = r.LinkRisk ?? r.linkRisk ?? "true";
        const linkRiskStr = String(linkRiskRaw).toLowerCase();
        const linkRisk =
          linkRiskStr === "true" ||
          linkRiskStr === "yes" ||
          linkRiskStr === "y" ||
          linkRiskStr === "1";

        return {
          id: uid(),
          label: r.Label ?? r.label ?? "",
          domain: r.Domain ?? r.domain ?? "Other",
          category: r.Category ?? r.category ?? "C4",
          frequency: r.Frequency ?? r.frequency ?? "Annual",
          startYear:
            Number(r.StartYear ?? r.startYear ?? model.time.startYear) ||
            model.time.startYear,
          endYear:
            Number(r.EndYear ?? r.endYear ?? model.time.startYear) ||
            model.time.startYear,
          year:
            Number(r.Year ?? r.year ?? model.time.startYear) ||
            model.time.startYear,
          unitValue:
            Number(r.UnitValue ?? r.unitValue ?? 0) || 0,
          quantity:
            Number(r.Quantity ?? r.quantity ?? 0) || 0,
          abatement:
            Number(r.Abatement ?? r.abatement ?? 0) || 0,
          annualAmount:
            Number(r.AnnualAmount ?? r.annualAmount ?? 0) || 0,
          growthPct:
            Number(r.GrowthPct ?? r.growthPct ?? 0) || 0,
          linkAdoption,
          linkRisk,
          p0: Number(r.P0 ?? r.p0 ?? 0) || 0,
          p1: Number(r.P1 ?? r.p1 ?? 0) || 0,
          consequence:
            Number(r.Consequence ?? r.consequence ?? 0) || 0,
          notes: r.Notes ?? r.notes ?? ""
        };
      });
    }

    // ----- Other costs -----
    if (costs && costs.length) {
      model.otherCosts = costs.map(r => {
        const typeRaw = r.Type ?? r.type ?? "annual";
        const typeStr = String(typeRaw).toLowerCase();
        const type = typeStr === "capital" ? "capital" : "annual";

        const cat = r.CostCategory ?? r.costCategory ?? "other";

        const constrainedRaw = r.Constrained ?? r.constrained ?? "true";
        const constrainedStr = String(constrainedRaw).toLowerCase();
        const constrained =
          constrainedStr === "true" ||
          constrainedStr === "yes" ||
          constrainedStr === "y" ||
          constrainedStr === "1";

        return {
          id: uid(),
          label: r.Label ?? r.label ?? "",
          type,
          costCategory: cat,
          annual:
            Number(r.Annual ?? r.annual ?? 0) || 0,
          startYear:
            Number(r.StartYear ?? r.startYear ?? model.time.startYear) ||
            model.time.startYear,
          endYear:
            Number(r.EndYear ?? r.endYear ?? model.time.startYear) ||
            model.time.startYear,
          capital:
            Number(r.Capital ?? r.capital ?? 0) || 0,
          year:
            Number(r.Year ?? r.year ?? model.time.startYear) ||
            model.time.startYear,
          constrained,
          useDepreciation: false,
          deprMethod: "sl",
          deprLife: 5,
          deprRate: 20
        };
      });
    }

    renderAll();
    calcAndRender();
    status.textContent = "Imported successfully from Excel/CSV.";
  }

  function downloadExcelTemplate() {
    if (!window.XLSX) {
      alert("SheetJS (XLSX) library is not loaded. Cannot create template.");
      return;
    }
    const wb = XLSX.utils.book_new();

    // Outputs sheet
    const outputsAoA = [
      ["Name", "Unit", "$/unit", "Source"]
    ];
    const wsOut = XLSX.utils.aoa_to_sheet(outputsAoA);
    XLSX.utils.book_append_sheet(wb, wsOut, "Outputs");

    // Treatments sheet
    const treatsAoA = [
      [
        "Name",
        "Area",
        "Replications",
        "Adoption",
        "AnnualCost",
        "MaterialsCost",
        "ServicesCost",
        "CapitalCost",
        "Constrained",
        "Source"
      ]
    ];
    const wsTr = XLSX.utils.aoa_to_sheet(treatsAoA);
    XLSX.utils.book_append_sheet(wb, wsTr, "Treatments");

    // Benefits sheet
    const benAoA = [[
      "Label",
      "Domain",
      "Category",
      "Frequency",
      "StartYear",
      "EndYear",
      "Year",
      "UnitValue",
      "Quantity",
      "Abatement",
      "AnnualAmount",
      "GrowthPct",
      "LinkAdoption",
      "LinkRisk",
      "P0",
      "P1",
      "Consequence",
      "Notes"
    ]];
    const wsBen = XLSX.utils.aoa_to_sheet(benAoA);
    XLSX.utils.book_append_sheet(wb, wsBen, "Benefits");

    // OtherCosts sheet
    const costAoA = [[
      "Label",
      "Type",
      "CostCategory",
      "Annual",
      "StartYear",
      "EndYear",
      "Capital",
      "Year",
      "Constrained"
    ]];
    const wsCost = XLSX.utils.aoa_to_sheet(costAoA);
    XLSX.utils.book_append_sheet(wb, wsCost, "OtherCosts");

    saveWorkbook("cba_excel_template.xlsx", wb);
  }

  function downloadSampleDataset() {
    if (!window.XLSX) {
      alert("SheetJS (XLSX) library is not loaded. Cannot create sample workbook.");
      return;
    }

    const wb = XLSX.utils.book_new();

    // Sample Outputs from current model
    const outAoA = [["Name", "Unit", "$/unit", "Source"]];
    model.outputs.forEach(o => {
      outAoA.push([o.name, o.unit, o.value, o.source]);
    });
    const wsOut = XLSX.utils.aoa_to_sheet(outAoA);
    XLSX.utils.book_append_sheet(wb, wsOut, "Outputs");

    // Sample Treatments
    const trAoA = [[
      "Name",
      "Area",
      "Replications",
      "Adoption",
      "AnnualCost",
      "MaterialsCost",
      "ServicesCost",
      "CapitalCost",
      "Constrained",
      "Source"
    ]];
    model.treatments.forEach(t => {
      trAoA.push([
        t.name,
        t.area,
        t.replications,
        t.adoption,
        t.annualCost,
        t.materialsCost,
        t.servicesCost,
        t.capitalCost,
        t.constrained ? "Yes" : "No",
        t.source
      ]);
    });
    const wsTr = XLSX.utils.aoa_to_sheet(trAoA);
    XLSX.utils.book_append_sheet(wb, wsTr, "Treatments");

    // Sample Benefits
    const benAoA = [[
      "Label",
      "Domain",
      "Category",
      "Frequency",
      "StartYear",
      "EndYear",
      "Year",
      "UnitValue",
      "Quantity",
      "Abatement",
      "AnnualAmount",
      "GrowthPct",
      "LinkAdoption",
      "LinkRisk",
      "P0",
      "P1",
      "Consequence",
      "Notes"
    ]];
    model.benefits.forEach(b => {
      benAoA.push([
        b.label,
        b.domain,
        b.category,
        b.frequency,
        b.startYear,
        b.endYear,
        b.year,
        b.unitValue,
        b.quantity,
        b.abatement,
        b.annualAmount,
        b.growthPct,
        b.linkAdoption,
        b.linkRisk,
        b.p0,
        b.p1,
        b.consequence,
        b.notes
      ]);
    });
    const wsBen = XLSX.utils.aoa_to_sheet(benAoA);
    XLSX.utils.book_append_sheet(wb, wsBen, "Benefits");

    // Sample OtherCosts
    const costAoA = [[
      "Label",
      "Type",
      "CostCategory",
      "Annual",
      "StartYear",
      "EndYear",
      "Capital",
      "Year",
      "Constrained"
    ]];
    model.otherCosts.forEach(c => {
      costAoA.push([
        c.label,
        c.type,
        c.costCategory,
        c.annual,
        c.startYear,
        c.endYear,
        c.capital,
        c.year,
        c.constrained ? "Yes" : "No"
      ]);
    });
    const wsCost = XLSX.utils.aoa_to_sheet(costAoA);
    XLSX.utils.book_append_sheet(wb, wsCost, "OtherCosts");

    saveWorkbook(`cba_sample_${slug(model.project.name || "project")}.xlsx`, wb);
  }

  // ---------- INIT ----------
  document.addEventListener("DOMContentLoaded", () => {
    try {
      initTabs();
      initActions();
      bindBasics();
      renderAll();
      calcAndRender();
    } catch (err) {
      console.error("Initialisation error:", err);
      const alertBox = $("#validation");
      if (alertBox) {
        alertBox.classList.add("show");
        alertBox.innerHTML = `<p class="error">Initialisation error: ${esc(err?.message || err)}</p>`;
      }
    }
  });
})();
