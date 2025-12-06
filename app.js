// Farming CBA Tool — Newcastle Business School (enhanced project, time, outputs, treatments, costs, simulation, report)
(() => {
  // ---------- CONSTANTS ----------
  const DEFAULT_DISCOUNT_SCHEDULE = [
    { label: "2025–2034", from: 2025, to: 2034, low: 2, base: 4, high: 6 },
    { label: "2035–2044", from: 2035, to: 2044, low: 4, base: 7, high: 10 },
    { label: "2045–2054", from: 2045, to: 2054, low: 4, base: 7, high: 10 },
    { label: "2055–2064", from: 2055, to: 2064, low: 3, base: 6, high: 9 },
    { label: "2065–2074", from: 2065, to: 2074, low: 2, base: 5, high: 8 }
  ];

  // ---------- MODEL ----------
  const model = {
    project: {
      name: "Nitrogen Optimization Trial",
      lead: "Project Lead",
      analysts: "Farm Econ Team",
      team: "",
      organisation: "Newcastle Business School, The University of Newcastle",
      contactEmail: "frank.agbola@newcastle.edu.au",
      contactPhone: "",
      summary:
        "Test fertilizer strategies to raise wheat yield and protein across 500 ha over 5 years.",
      objectives: "",
      activities: "",
      stakeholders: "",
      lastUpdated: new Date().toISOString().slice(0, 10),
      goal: "Increase yield by 10% and protein by 0.5 p.p. on 500 ha within 3 years.",
      withProject: "Adopt optimized nitrogen timing and rates; improved management on 500 ha.",
      withoutProject:
        "Business-as-usual fertilization; yield/protein unchanged; rising costs."
    },
    time: {
      startYear: new Date().getFullYear(),
      projectStartYear: new Date().getFullYear(),
      years: 10,
      discBase: 7,
      discLow: 4,
      discHigh: 10,
      mirrFinance: 6,
      mirrReinvest: 4,
      discountSchedule: JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE))
    },
    outputsMeta: {
      systemType: "single",
      assumptions: ""
    },
    outputs: [
      { id: uid(), name: "Yield", unit: "t/ha", value: 300, source: "Input Directly" },
      { id: uid(), name: "Protein", unit: "%-point", value: 12, source: "Input Directly" },
      { id: uid(), name: "Moisture", unit: "%-point", value: -5, source: "Input Directly" },
      { id: uid(), name: "Biomass", unit: "t/ha", value: 40, source: "Input Directly" }
    ],
    treatments: [
      {
        id: uid(),
        name: "Optimized N (Rate+Timing)",
        area: 300,
        adoption: 0.8,
        deltas: {},
        annualCost: 45,
        materialsCost: 0,
        servicesCost: 0,
        capitalCost: 5000,
        constrained: true,
        source: "Farm Trials",
        replications: 1,
        isControl: false,
        notes: ""
      },
      {
        id: uid(),
        name: "Slow-Release N",
        area: 200,
        adoption: 0.7,
        deltas: {},
        annualCost: 25,
        materialsCost: 0,
        servicesCost: 0,
        capitalCost: 0,
        constrained: true,
        source: "ABARES",
        replications: 1,
        isControl: false,
        notes: ""
      }
    ],
    benefits: [
      {
        id: uid(),
        label: "Reduced recurring costs (energy/water)",
        category: "C4",
        theme: "Cost savings",
        frequency: "Annual",
        startYear: new Date().getFullYear(),
        endYear: new Date().getFullYear() + 4,
        year: new Date().getFullYear(),
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
        notes: "Project-wide OPEX saving"
      },
      {
        id: uid(),
        label: "Reduced risk of quality downgrades",
        category: "C7",
        theme: "Risk reduction",
        frequency: "Annual",
        startYear: new Date().getFullYear(),
        endYear: new Date().getFullYear() + 9,
        year: new Date().getFullYear(),
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
        label: "Soil asset value uplift (carbon/structure)",
        category: "C6",
        theme: "Soil carbon",
        frequency: "Once",
        startYear: new Date().getFullYear(),
        endYear: new Date().getFullYear(),
        year: new Date().getFullYear() + 5,
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
        label: "Project Mgmt & M&E",
        type: "annual",
        category: "Services",
        annual: 20000,
        startYear: new Date().getFullYear(),
        endYear: new Date().getFullYear() + 4,
        capital: 0,
        year: new Date().getFullYear(),
        constrained: true,
        depMethod: "none",
        depLife: 5,
        depRate: 30
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
      variationPct: 20,
      varyOutputs: true,
      varyTreatCosts: true,
      varyInputCosts: false
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

  // ---------- UTIL ----------
  function uid() { return Math.random().toString(36).slice(2, 10); }
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const fmt = n => (isFinite(n) ? (Math.abs(n) >= 1000 ? n.toLocaleString(undefined, { maximumFractionDigits: 0 }) : n.toLocaleString(undefined, { maximumFractionDigits: 2 })) : "—");
  const money = n => (isFinite(n) ? "$" + fmt(n) : "—");
  const percent = n => (isFinite(n) ? fmt(n) + "%" : "—");
  const slug = s => (s || "project").toLowerCase().replace(/[^a-z0-9]+/g,"_").replace(/^_|_$/g,"");
  const annuityFactor = (N, rPct) => { const r = rPct / 100; return r === 0 ? N : (1 - Math.pow(1 + r, -N)) / r; };
  const esc = s => (s ?? "").toString().replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const horizons = [5, 10, 15, 20, 25];

  function saveWorkbook(filename, wb) {
    try { if (window.XLSX && typeof XLSX.writeFile === "function") { XLSX.writeFile(wb, filename, { compression: true }); return; } }
    catch(_) {}
    try {
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array", compression: true });
      const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = filename;
      document.body.appendChild(a); a.click(); setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove();},300);
    } catch (err) { alert("Download failed. Ensure SheetJS script is loaded.\n\n"+(err?.message||err)); }
  }

  // ---------- CASHFLOWS ----------
  function buildCashflows({ forRate = model.time.discBase, adoptMul = model.adoption.base, risk = model.risk.base }) {
    const N = model.time.years;
    const baseYear = model.time.startYear;

    const benefitByYear = new Array(N + 1).fill(0);
    const costByYear = new Array(N + 1).fill(0);
    const constrainedCostByYear = new Array(N + 1).fill(0);

    // Treatments × Outputs
    let annualBenefit = 0, treatAnnualCost = 0, treatConstrAnnualCost = 0;
    let treatCapitalY0 = 0, treatConstrCapitalY0 = 0;

    model.treatments.forEach(t => {
      const adopt = clamp(t.adoption * adoptMul, 0, 1);
      let valuePerHa = 0;
      model.outputs.forEach(o => {
        const delta = Number(t.deltas[o.id]) || 0;
        const v = Number(o.value) || 0;
        valuePerHa += delta * v;
      });
      const area = Number(t.area) || 0;
      const benefit = valuePerHa * area * (1 - clamp(risk, 0, 1)) * adopt;

      const annualCostPerHa = ((Number(t.materialsCost) || 0) + (Number(t.servicesCost) || 0)) || (Number(t.annualCost) || 0);
      const opCost = annualCostPerHa * area;
      const cap = Number(t.capitalCost) || 0;

      annualBenefit += benefit;
      treatAnnualCost += opCost;
      treatCapitalY0 += cap;

      if (t.constrained) {
        treatConstrAnnualCost += opCost;
        treatConstrCapitalY0 += cap;
      }
    });

    costByYear[0] += treatCapitalY0;
    constrainedCostByYear[0] += treatConstrCapitalY0;
    for (let t = 1; t <= N; t++) {
      benefitByYear[t] += annualBenefit;
      costByYear[t] += treatAnnualCost;
      constrainedCostByYear[t] += treatConstrAnnualCost;
    }

    // Other project costs
    const otherAnnualByYear = new Array(N + 1).fill(0);
    const otherConstrAnnualByYear = new Array(N + 1).fill(0);
    let otherCapitalY0 = 0, otherConstrCapitalY0 = 0;

    model.otherCosts.forEach(c => {
      if (c.type === "annual") {
        const a = Number(c.annual) || 0;
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
        const cap = Number(c.capital) || 0;
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
    const annualGM = annualBenefit - treatAnnualCost;
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

  function presentValue(series, ratePct) {
    let pv = 0;
    for (let t = 0; t < series.length; t++) pv += series[t] / Math.pow(1 + ratePct / 100, t);
    return pv;
  }

  function computeAll(rate, adoptMul, risk, bcrMode) {
    const { benefitByYear, costByYear, constrainedCostByYear, cf, annualGM } =
      buildCashflows({ forRate: rate, adoptMul, risk });
    const pvBenefits = presentValue(benefitByYear, rate);
    const pvCosts = presentValue(costByYear, rate);
    const pvCostsConstrained = presentValue(constrainedCostByYear, rate);

    const npv = pvBenefits - pvCosts;
    const denom = (bcrMode === "constrained") ? pvCostsConstrained : pvCosts;
    const bcr = denom > 0 ? pvBenefits / denom : NaN;

    const irrVal = irr(cf);
    const mirrVal = mirr(cf, model.time.mirrFinance, model.time.mirrReinvest);
    const roi = pvCosts > 0 ? ((pvBenefits - pvCosts) / pvCosts) * 100 : NaN;
    const profitMargin = benefitByYear[1] > 0 ? (annualGM / benefitByYear[1]) * 100 : NaN;
    const pb = payback(cf, rate);

    return { pvBenefits, pvCosts, pvCostsConstrained, npv, bcr, irrVal, mirrVal, roi, annualGM, profitMargin, paybackYears: pb, cf, benefitByYear, costByYear };
  }

  function irr(cf) {
    const hasPos = cf.some(v => v > 0);
    const hasNeg = cf.some(v => v < 0);
    if (!hasPos || !hasNeg) return NaN;
    let lo = -0.99, hi = 5.0;
    const npvAt = r => cf.reduce((acc, v, t) => acc + v / Math.pow(1 + r, t), 0);
    let nLo = npvAt(lo), nHi = npvAt(hi);
    if (nLo * nHi > 0) { for (let k = 0; k < 20 && nLo * nHi > 0; k++) { hi *= 1.5; nHi = npvAt(hi); } if (nLo * nHi > 0) return NaN; }
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
    for (let t = 0; t < cf.length; t++) {
      cum += cf[t] / Math.pow(1 + ratePct / 100, t);
      if (cum >= 0) return t;
    }
    return null;
  }

  // ---------- DOM ----------
  const $ = sel => document.querySelector(sel);
  const $$ = sel => Array.from(document.querySelectorAll(sel));
  const num = sel => +(document.querySelector(sel)?.value || 0);
  const setVal = (sel, text) => {
    const el = document.querySelector(sel);
    if (el) el.textContent = text;
  };

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
    });
  }

  // ---------- BIND + RENDER FORMS ----------
  function setBasicsFieldsFromModel() {
    if ($("#projectName")) $("#projectName").value = model.project.name || "";
    if ($("#projectLead")) $("#projectLead").value = model.project.lead || "";
    if ($("#analystNames")) $("#analystNames").value = model.project.analysts || "";
    if ($("#projectTeam")) $("#projectTeam").value = model.project.team || "";
    if ($("#projectSummary")) $("#projectSummary").value = model.project.summary || "";
    if ($("#projectObjectives")) $("#projectObjectives").value = model.project.objectives || "";
    if ($("#projectActivities")) $("#projectActivities").value = model.project.activities || "";
    if ($("#stakeholderGroups")) $("#stakeholderGroups").value = model.project.stakeholders || "";
    if ($("#lastUpdated")) $("#lastUpdated").value = model.project.lastUpdated || "";
    if ($("#projectGoal")) $("#projectGoal").value = model.project.goal || "";
    if ($("#withProject")) $("#withProject").value = model.project.withProject || "";
    if ($("#withoutProject")) $("#withoutProject").value = model.project.withoutProject || "";
    if ($("#organisation")) $("#organisation").value = model.project.organisation || "";
    if ($("#contactEmail")) $("#contactEmail").value = model.project.contactEmail || "";
    if ($("#contactPhone")) $("#contactPhone").value = model.project.contactPhone || "";

    if ($("#startYear")) $("#startYear").value = model.time.startYear;
    if ($("#projectStartYear")) $("#projectStartYear").value = model.time.projectStartYear || model.time.startYear;
    if ($("#years")) $("#years").value = model.time.years;
    if ($("#discBase")) $("#discBase").value = model.time.discBase;
    if ($("#discLow")) $("#discLow").value = model.time.discLow;
    if ($("#discHigh")) $("#discHigh").value = model.time.discHigh;
    if ($("#mirrFinance")) $("#mirrFinance").value = model.time.mirrFinance;
    if ($("#mirrReinvest")) $("#mirrReinvest").value = model.time.mirrReinvest;

    if ($("#adoptBase")) $("#adoptBase").value = model.adoption.base;
    if ($("#adoptLow")) $("#adoptLow").value = model.adoption.low;
    if ($("#adoptHigh")) $("#adoptHigh").value = model.adoption.high;

    if ($("#riskBase")) $("#riskBase").value = model.risk.base;
    if ($("#riskLow")) $("#riskLow").value = model.risk.low;
    if ($("#riskHigh")) $("#riskHigh").value = model.risk.high;
    if ($("#rTech")) $("#rTech").value = model.risk.tech;
    if ($("#rNonCoop")) $("#rNonCoop").value = model.risk.nonCoop;
    if ($("#rSocio")) $("#rSocio").value = model.risk.socio;
    if ($("#rFin")) $("#rFin").value = model.risk.fin;
    if ($("#rMan")) $("#rMan").value = model.risk.man;

    if ($("#simN")) $("#simN").value = model.sim.n;
    if ($("#targetBCR")) $("#targetBCR").value = model.sim.targetBCR;
    if ($("#bcrMode")) $("#bcrMode").value = model.sim.bcrMode;
    if ($("#simBcrTargetLabel")) $("#simBcrTargetLabel").textContent = model.sim.targetBCR;

    if ($("#simVarPct")) $("#simVarPct").value = String(model.sim.variationPct || 20);
    if ($("#simVaryOutputs")) $("#simVaryOutputs").value = model.sim.varyOutputs ? "true" : "false";
    if ($("#simVaryTreatCosts")) $("#simVaryTreatCosts").value = model.sim.varyTreatCosts ? "true" : "false";
    if ($("#simVaryInputCosts")) $("#simVaryInputCosts").value = model.sim.varyInputCosts ? "true" : "false";

    if ($("#systemType")) $("#systemType").value = model.outputsMeta.systemType || "single";
    if ($("#outputAssumptions")) $("#outputAssumptions").value = model.outputsMeta.assumptions || "";

    const sched = model.time.discountSchedule || DEFAULT_DISCOUNT_SCHEDULE;
    $$("input[data-disc-period]").forEach(inp => {
      const idx = +inp.dataset.discPeriod;
      const scenario = inp.dataset.scenario;
      const row = sched[idx];
      if (!row) return;
      let v = "";
      if (scenario === "low") v = row.low;
      else if (scenario === "base") v = row.base;
      else if (scenario === "high") v = row.high;
      inp.value = v ?? "";
    });
  }

  function bindBasics() {
    setBasicsFieldsFromModel();

    if ($("#calcCombinedRisk")) {
      $("#calcCombinedRisk").addEventListener("click", () => {
        const r = 1 - (1 - num("#rTech")) * (1 - num("#rNonCoop")) * (1 - num("#rSocio")) * (1 - num("#rFin")) * (1 - num("#rMan"));
        if ($("#combinedRiskOut")) $("#combinedRiskOut").textContent = `Combined: ${(r * 100).toFixed(2)}%`;
        if ($("#riskBase")) $("#riskBase").value = r.toFixed(3);
        model.risk.base = r;
        calcAndRender();
      });
    }

    if ($("#addCost")) {
      $("#addCost").addEventListener("click", () => {
        const c = {
          id: uid(),
          label: "New Cost",
          type: "annual",
          category: "Labour",
          annual: 0,
          startYear: model.time.startYear,
          endYear: model.time.startYear,
          capital: 0,
          year: model.time.startYear,
          constrained: true,
          depMethod: "none",
          depLife: 5,
          depRate: 30
        };
        model.otherCosts.push(c);
        renderCosts();
        calcAndRender();
      });
    }

    document.addEventListener("input", e => {
      const t = e.target;
      if (!t) return;

      // Discount schedule inputs
      if (t.dataset && t.dataset.discPeriod !== undefined) {
        const idx = +t.dataset.discPeriod;
        const scenario = t.dataset.scenario;
        if (!model.time.discountSchedule) model.time.discountSchedule = JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE));
        const row = model.time.discountSchedule[idx];
        if (row && scenario) {
          const val = +t.value;
          if (scenario === "low") row.low = val;
          else if (scenario === "base") row.base = val;
          else if (scenario === "high") row.high = val;
          calcAndRenderDebounced();
        }
        return;
      }

      const id = t.id;
      if (!id) return;
      switch (id) {
        case "projectName": model.project.name = t.value; break;
        case "projectLead": model.project.lead = t.value; break;
        case "analystNames": model.project.analysts = t.value; break;
        case "projectTeam": model.project.team = t.value; break;
        case "projectSummary": model.project.summary = t.value; break;
        case "projectObjectives": model.project.objectives = t.value; break;
        case "projectActivities": model.project.activities = t.value; break;
        case "stakeholderGroups": model.project.stakeholders = t.value; break;
        case "lastUpdated": model.project.lastUpdated = t.value; break;
        case "projectGoal": model.project.goal = t.value; break;
        case "withProject": model.project.withProject = t.value; break;
        case "withoutProject": model.project.withoutProject = t.value; break;
        case "organisation": model.project.organisation = t.value; break;
        case "contactEmail": model.project.contactEmail = t.value; break;
        case "contactPhone": model.project.contactPhone = t.value; break;

        case "startYear": model.time.startYear = +t.value; break;
        case "projectStartYear": model.time.projectStartYear = +t.value; break;
        case "years": model.time.years = +t.value; break;
        case "discBase": model.time.discBase = +t.value; break;
        case "discLow": model.time.discLow = +t.value; break;
        case "discHigh": model.time.discHigh = +t.value; break;
        case "mirrFinance": model.time.mirrFinance = +t.value; break;
        case "mirrReinvest": model.time.mirrReinvest = +t.value; break;

        case "adoptBase": model.adoption.base = +t.value; break;
        case "adoptLow": model.adoption.low = +t.value; break;
        case "adoptHigh": model.adoption.high = +t.value; break;

        case "riskBase": model.risk.base = +t.value; break;
        case "riskLow": model.risk.low = +t.value; break;
        case "riskHigh": model.risk.high = +t.value; break;
        case "rTech": model.risk.tech = +t.value; break;
        case "rNonCoop": model.risk.nonCoop = +t.value; break;
        case "rSocio": model.risk.socio = +t.value; break;
        case "rFin": model.risk.fin = +t.value; break;
        case "rMan": model.risk.man = +t.value; break;

        case "simN": model.sim.n = +t.value; break;
        case "targetBCR": model.sim.targetBCR = +t.value; if ($("#simBcrTargetLabel")) $("#simBcrTargetLabel").textContent = t.value; break;
        case "bcrMode": model.sim.bcrMode = t.value; break;
        case "randSeed": model.sim.seed = t.value ? +t.value : null; break;

        case "simVarPct": model.sim.variationPct = +t.value || 20; break;
        case "simVaryOutputs": model.sim.varyOutputs = t.value === "true"; break;
        case "simVaryTreatCosts": model.sim.varyTreatCosts = t.value === "true"; break;
        case "simVaryInputCosts": model.sim.varyInputCosts = t.value === "true"; break;

        case "systemType": model.outputsMeta.systemType = t.value; break;
        case "outputAssumptions": model.outputsMeta.assumptions = t.value; break;
      }
      calcAndRenderDebounced();
    });

    if ($("#saveProject")) {
      $("#saveProject").addEventListener("click", () => {
        const data = JSON.stringify(model, null, 2);
        downloadFile(`cba_${(model.project.name || "project").replace(/\s+/g, "_")}.json`, data, "application/json");
      });
    }
    if ($("#loadProject")) {
      $("#loadProject").addEventListener("click", () => $("#loadFile")?.click());
    }
    if ($("#loadFile")) {
      $("#loadFile").addEventListener("change", async e => {
        const file = e.target.files?.[0]; if (!file) return;
        const text = await file.text();
        try {
          const obj = JSON.parse(text);
          Object.assign(model, obj);
          if (!model.time.discountSchedule) model.time.discountSchedule = JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE));
          initTreatmentDeltas();
          renderAll();
          setBasicsFieldsFromModel();
          calcAndRender();
        } catch {
          alert("Invalid JSON file.");
        } finally { e.target.value = ""; }
      });
    }

    if ($("#exportCsv")) $("#exportCsv").addEventListener("click", exportAllCsv);
    if ($("#exportCsvFoot")) $("#exportCsvFoot").addEventListener("click", exportAllCsv);
    if ($("#exportPdf")) $("#exportPdf").addEventListener("click", exportPdf);
    if ($("#exportPdfFoot")) $("#exportPdfFoot").addEventListener("click", exportPdf);

    if ($("#parseExcel")) $("#parseExcel").addEventListener("click", handleParseExcel);
    if ($("#importExcel")) $("#importExcel").addEventListener("click", commitExcelToModel);

    if ($("#downloadTemplate")) $("#downloadTemplate").addEventListener("click", downloadExcelTemplate);
    if ($("#downloadSample")) $("#downloadSample").addEventListener("click", downloadSampleDataset);

    if ($("#startBtn")) $("#startBtn").addEventListener("click", () => switchTab("project"));
  }

  // ---------- RENDERERS ----------
  function renderOutputs() {
    const root = $("#outputsList"); if (!root) return;
    root.innerHTML = "";
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
  if ($("#addOutput")) {
    $("#addOutput").addEventListener("click", () => {
      const id = uid();
      model.outputs.push({ id, name: "Custom Output", unit: "unit", value: 0, source: "Input Directly" });
      model.treatments.forEach(t => (t.deltas[id] = 0));
      renderOutputs(); renderTreatments(); renderDatabaseTags();
    });
  }

  function renderTreatments() {
    const root = $("#treatmentsList"); if (!root) return;
    root.innerHTML = "";
    const list = [...model.treatments];
    list.forEach(t => {
      const materials = Number(t.materialsCost) || 0;
      const services = Number(t.servicesCost) || 0;
      const totalPerHa = (materials + services) || (Number(t.annualCost) || 0);
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>🚜 Treatment: ${esc(t.name)}</h4>
        <div class="row">
          <div class="field"><label>Name</label><input value="${esc(t.name)}" data-tk="name" data-id="${t.id}" /></div>
          <div class="field"><label>Area (ha)</label><input type="number" step="0.01" value="${t.area}" data-tk="area" data-id="${t.id}" /></div>
          <div class="field"><label>Adoption (0–1)</label><input type="number" min="0" max="1" step="0.01" value="${t.adoption}" data-tk="adoption" data-id="${t.id}" /></div>
          <div class="field"><label>Replications</label><input type="number" min="1" step="1" value="${t.replications || 1}" data-tk="replications" data-id="${t.id}" /></div>
          <div class="field"><label>Source</label>
            <select data-tk="source" data-id="${t.id}">
              ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"].map(s => `<option ${s===t.source?"selected":""}>${s}</option>`).join("")}
            </select>
          </div>
          <div class="field"><label>Control group?</label>
            <select data-tk="isControl" data-id="${t.id}">
              <option value="false" ${!t.isControl?"selected":""}>No</option>
              <option value="true" ${t.isControl?"selected":""}>Yes</option>
            </select>
          </div>
          <div class="field"><label>&nbsp;</label><button class="danger" data-del-treatment="${t.id}">Remove</button></div>
        </div>
        <div class="row-6">
          <div class="field"><label>Materials cost ($/ha)</label><input type="number" step="0.01" value="${t.materialsCost || 0}" data-tk="materialsCost" data-id="${t.id}" /></div>
          <div class="field"><label>Services cost ($/ha)</label><input type="number" step="0.01" value="${t.servicesCost || 0}" data-tk="servicesCost" data-id="${t.id}" /></div>
          <div class="field"><label>Total annual cost ($/ha)</label><input type="number" step="0.01" value="${totalPerHa}" readonly /></div>
          <div class="field"><label>Annual Cost (fallback, $/ha)</label><input type="number" step="0.01" value="${t.annualCost}" data-tk="annualCost" data-id="${t.id}" /></div>
          <div class="field"><label>Capital Cost ($, y0)</label><input type="number" step="0.01" value="${t.capitalCost}" data-tk="capitalCost" data-id="${t.id}" /></div>
          <div class="field"><label>Constrained?</label>
            <select data-tk="constrained" data-id="${t.id}">
              <option value="true" ${t.constrained?"selected":""}>Yes</option>
              <option value="false" ${!t.constrained?"selected":""}>No</option>
            </select>
          </div>
        </div>
        <div class="field">
          <label>Notes (e.g. replication design, control definition)</label>
          <textarea data-tk="notes" data-id="${t.id}" rows="2">${esc(t.notes || "")}</textarea>
        </div>
        <h5>Output Deltas (per ha)</h5>
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
        else if (tk === "name" || tk === "source" || tk === "notes") t[tk] = e.target.value;
        else if (tk === "isControl") {
          const val = e.target.value === "true";
          model.treatments.forEach(tt => { tt.isControl = false; });
          if (val) t.isControl = true;
          renderTreatments();
          calcAndRenderDebounced();
          return;
        } else if (tk === "replications") {
          t[tk] = Math.max(1, Math.round(+e.target.value || 1));
        } else {
          t[tk] = +e.target.value;
        }
      }
      const td = e.target.dataset.td;
      if (td) t.deltas[td] = +e.target.value;
      calcAndRenderDebounced();
    };
    root.addEventListener("click", e => {
      const id = e.target.dataset.delTreatment; if (!id) return;
      if (!confirm("Remove this treatment?")) return;
      model.treatments = model.treatments.filter(x => x.id !== id);
      renderTreatments(); renderDatabaseTags(); calcAndRender();
    });
  }
  if ($("#addTreatment")) {
    $("#addTreatment").addEventListener("click", () => {
      if (model.treatments.length >= 64) {
        alert("Maximum of 64 treatments reached.");
        return;
      }
      const t = {
        id: uid(),
        name: "New Treatment",
        area: 0,
        adoption: 0.5,
        deltas: {},
        annualCost: 0,
        materialsCost: 0,
        servicesCost: 0,
        capitalCost: 0,
        constrained: true,
        source: "Input Directly",
        replications: 1,
        isControl: false,
        notes: ""
      };
      model.outputs.forEach(o => { t.deltas[o.id] = 0; });
      model.treatments.push(t);
      renderTreatments(); renderDatabaseTags();
    });
  }

  function renderBenefits() {
    const root = $("#benefitsList"); if (!root) return;
    root.innerHTML = "";
    const THEMES = [
      "Soil chemical",
      "Soil physical",
      "Soil biological",
      "Soil carbon",
      "Soil pH by depth",
      "Soil nutrients by depth",
      "Soil properties by treatment",
      "Cost savings",
      "Water retention",
      "Risk reduction",
      "Other"
    ];
    model.benefits.forEach(b => {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>🌱 ${esc(b.label || "Benefit")}</h4>
        <div class="row-6">
          <div class="field"><label>Label</label><input value="${esc(b.label||"")}" data-bk="label" data-id="${b.id}" /></div>
          <div class="field"><label>Category</label>
            <select data-bk="category" data-id="${b.id}">
              ${["C1","C2","C3","C4","C5","C6","C7","C8"].map(c=>`<option ${c===b.category?"selected":""}>${c}</option>`).join("")}
            </select>
          </div>
          <div class="field"><label>Benefit type</label>
            <select data-bk="theme" data-id="${b.id}">
              ${THEMES.map(th=>`<option ${th===(b.theme||"")?"selected":""}>${th}</option>`).join("")}
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
        </div>

        <div class="row-6">
          <div class="field"><label>Once year</label><input type="number" value="${b.year||model.time.startYear}" data-bk="year" data-id="${b.id}" /></div>
          <div class="field"><label>Unit value ($)</label><input type="number" step="0.01" value="${b.unitValue||0}" data-bk="unitValue" data-id="${b.id}" /></div>
          <div class="field"><label>Quantity</label><input type="number" step="0.01" value="${b.quantity||0}" data-bk="quantity" data-id="${b.id}" /></div>
          <div class="field"><label>Abatement</label><input type="number" step="0.01" value="${b.abatement||0}" data-bk="abatement" data-id="${b.id}" /></div>
          <div class="field"><label>Annual amount ($)</label><input type="number" step="0.01" value="${b.annualAmount||0}" data-bk="annualAmount" data-id="${b.id}" /></div>
          <div class="field"><label>Growth (%/yr)</label><input type="number" step="0.01" value="${b.growthPct||0}" data-bk="growthPct" data-id="${b.id}" /></div>
        </div>

        <div class="row-6">
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
      if (["label","category","frequency","notes","theme"].includes(k)) b[k] = e.target.value;
      else if (k === "linkAdoption" || k === "linkRisk") b[k] = e.target.value === "true";
      else b[k] = +e.target.value;
      calcAndRenderDebounced();
    };
    root.addEventListener("click", e => {
      const id = e.target.dataset.delBenefit; if (!id) return;
      if (!confirm("Remove this benefit item?")) return;
      model.benefits = model.benefits.filter(x => x.id !== id);
      renderBenefits(); calcAndRender();
    });
  }
  if ($("#addBenefit")) {
    $("#addBenefit").addEventListener("click", () => {
      model.benefits.push({
        id: uid(), label: "New Benefit", category: "C4", theme: "Other", frequency: "Annual",
        startYear: model.time.startYear, endYear: model.time.startYear, year: model.time.startYear,
        unitValue: 0, quantity: 0, abatement: 0, annualAmount: 0, growthPct: 0,
        linkAdoption: true, linkRisk: true, p0: 0, p1: 0, consequence: 0, notes: ""
      });
      renderBenefits();
    });
  }

  function renderCosts() {
    const root = $("#costsList"); if (!root) return;
    root.innerHTML = "";
    model.otherCosts.forEach(c => {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>💰 Cost Item: ${esc(c.label)}</h4>
        <div class="row-6">
          <div class="field"><label>Label</label><input value="${esc(c.label)}" data-ck="label" data-id="${c.id}" /></div>
          <div class="field"><label>Type</label>
            <select data-ck="type" data-id="${c.id}">
              <option value="annual" ${c.type==="annual"?"selected":""}>Annual</option>
              <option value="capital" ${c.type==="capital"?"selected":""}>Capital</option>
            </select>
          </div>
          <div class="field"><label>Category</label>
            <select data-ck="category" data-id="${c.id}">
              <option ${c.category==="Capital"?"selected":""}>Capital</option>
              <option ${c.category==="Labour"?"selected":""}>Labour</option>
              <option ${c.category==="Materials"?"selected":""}>Materials</option>
              <option ${c.category==="Services"?"selected":""}>Services</option>
            </select>
          </div>
          <div class="field"><label>Annual ($/yr)</label><input type="number" step="0.01" value="${c.annual ?? 0}" data-ck="annual" data-id="${c.id}" /></div>
          <div class="field"><label>Start Year</label><input type="number" value="${c.startYear ?? model.time.startYear}" data-ck="startYear" data-id="${c.id}" /></div>
          <div class="field"><label>End Year</label><input type="number" value="${c.endYear ?? model.time.startYear}" data-ck="endYear" data-id="${c.id}" /></div>
        </div>
        <div class="row-6">
          <div class="field"><label>Capital ($)</label><input type="number" step="0.01" value="${c.capital ?? 0}" data-ck="capital" data-id="${c.id}" /></div>
          <div class="field"><label>Capital Year</label><input type="number" value="${c.year ?? model.time.startYear}" data-ck="year" data-id="${c.id}" /></div>
          <div class="field"><label>Depreciation method</label>
            <select data-ck="depMethod" data-id="${c.id}">
              <option value="none" ${c.depMethod==="none"?"selected":""}>None</option>
              <option value="straight" ${c.depMethod==="straight"?"selected":""}>Straight-line</option>
              <option value="declining" ${c.depMethod==="declining"?"selected":""}>Declining-balance</option>
            </select>
          </div>
          <div class="field"><label>Life (years)</label><input type="number" step="1" min="1" value="${c.depLife || 5}" data-ck="depLife" data-id="${c.id}" /></div>
          <div class="field"><label>Declining rate (%/yr)</label><input type="number" step="1" value="${c.depRate || 30}" data-ck="depRate" data-id="${c.id}" /></div>
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
      if (["label","type","category","depMethod"].includes(k)) c[k] = e.target.value;
      else if (k === "constrained") c[k] = e.target.value === "true";
      else c[k] = +e.target.value;
      calcAndRenderDebounced();
    };
    root.addEventListener("click", e => {
      const id = e.target.dataset.delCost; if (!id) return;
      if (!confirm("Remove this cost item?")) return;
      model.otherCosts = model.otherCosts.filter(x => x.id !== id);
      renderCosts(); calcAndRender();
    });
  }

  function renderDatabaseTags() {
    const outRoot = $("#dbOutputs"); if (outRoot) {
      outRoot.innerHTML = "";
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
    }

    const tRoot = $("#dbTreatments"); if (tRoot) {
      tRoot.innerHTML = "";
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
  }

  function computeSingleTreatmentMetrics(t, rate, years, adoptMul, risk) {
    let valuePerHa = 0;
    model.outputs.forEach(o => (valuePerHa += (Number(t.deltas[o.id]) || 0) * (Number(o.value) || 0)));
    const adopt = clamp(t.adoption * adoptMul, 0, 1);
    const area = Number(t.area) || 0;
    const annualBen = valuePerHa * area * (1 - clamp(risk, 0, 1)) * adopt;
    const annualCostPerHa = ((Number(t.materialsCost) || 0) + (Number(t.servicesCost) || 0)) || (Number(t.annualCost) || 0);
    const annualCost = annualCostPerHa * area;
    const cap = Number(t.capitalCost) || 0;
    const pvBen = annualBen * annuityFactor(years, rate);
    const pvCost = cap + annualCost * annuityFactor(years, rate);
    const bcr = pvCost > 0 ? pvBen / pvCost : NaN;
    const npv = pvBen - pvCost;
    const cf = new Array(years + 1).fill(0);
    cf[0] = -cap;
    for (let i = 1; i <= years; i++) cf[i] = annualBen - annualCost;
    const irrVal = irr(cf);
    const mirrVal = mirr(cf, model.time.mirrFinance, model.time.mirrReinvest);
    const roi = pvCost > 0 ? (npv / pvCost) * 100 : NaN;
    const gm = annualBen - annualCost;
    const gpm = annualBen > 0 ? (gm / annualBen) * 100 : NaN;
    const pb = payback(cf, rate);
    return { pvBen, pvCost, bcr, npv, irrVal, mirrVal, roi, gm, gpm, pb };
  }

  function renderTreatmentSummary(rate, adoptMul, risk) {
    const root = $("#treatmentSummary"); if (!root) return;
    root.innerHTML = "";
    const list = [...model.treatments].map(t => {
      const m = computeSingleTreatmentMetrics(t, rate, model.time.years, adoptMul, risk);
      return { t, m };
    }).sort((a, b) => {
      const bcrA = isFinite(a.m.bcr) ? a.m.bcr : -Infinity;
      const bcrB = isFinite(b.m.bcr) ? b.m.bcr : -Infinity;
      return bcrB - bcrA;
    });

    list.forEach((entry, idx) => {
      const t = entry.t, m = entry.m;
      const adopt = clamp(t.adoption * adoptMul, 0, 1);
      const area = Number(t.area) || 0;
      const annualBen = m.gm + ((Number(t.materialsCost) || 0) + (Number(t.servicesCost) || 0) || (Number(t.annualCost) || 0)) * area;
      const annualCostPerHa = ((Number(t.materialsCost) || 0) + (Number(t.servicesCost) || 0)) || (Number(t.annualCost) || 0);
      const annualCost = annualCostPerHa * area;
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <div class="row-6">
          <div class="field"><label>Rank</label><div class="metric"><div class="value">${idx + 1}</div></div></div>
          <div class="field"><label>Treatment</label><div class="metric"><div class="value">${esc(t.name)}${t.isControl ? " (Control)" : ""}</div></div></div>
          <div class="field"><label>Area</label><div class="metric"><div class="value">${fmt(area)} ha</div></div></div>
          <div class="field"><label>Adoption</label><div class="metric"><div class="value">${fmt(adopt)}</div></div></div>
          <div class="field"><label>Annual Benefit</label><div class="metric"><div class="value">${money(annualBen)}</div></div></div>
          <div class="field"><label>Annual Cost</label><div class="metric"><div class="value">${money(annualCost)}</div></div></div>
          <div class="field"><label>PV Benefit</label><div class="metric"><div class="value">${money(m.pvBen)}</div></div></div>
          <div class="field"><label>PV Cost</label><div class="metric"><div class="value">${money(m.pvCost)}</div></div></div>
          <div class="field"><label>BCR</label><div class="metric"><div class="value">${isFinite(m.bcr)?fmt(m.bcr):"—"}</div></div></div>
          <div class="field"><label>NPV</label><div class="metric"><div class="value">${money(m.npv)}</div></div></div>
          <div class="field"><label>IRR</label><div class="metric"><div class="value">${isFinite(m.irrVal)?percent(m.irrVal):"—"}</div></div></div>
          <div class="field"><label>Payback</label><div class="metric"><div class="value">${m.pb!=null?m.pb:"Not reached"}</div></div></div>
        </div>`;
      root.appendChild(el);
    });
  }

  function computeControlAndTreatmentGroupMetrics(rate, adoptMul, risk) {
    const years = model.time.years;
    const control = model.treatments.find(t => t.isControl);
    let controlMetrics = null;
    if (control) {
      controlMetrics = computeSingleTreatmentMetrics(control, rate, years, adoptMul, risk);
    }
    const combined = {
      area: 0,
      adoption: 1,
      deltas: {},
      materialsCost: 0,
      servicesCost: 0,
      annualCost: 0,
      capitalCost: 0
    };
    model.outputs.forEach(o => combined.deltas[o.id] = 0);
    model.treatments.forEach(t => {
      if (t.isControl) return;
      combined.capitalCost += Number(t.capitalCost) || 0;
      const area = Number(t.area) || 0;
      const adopt = clamp(t.adoption * adoptMul, 0, 1);
      combined.area += area;
      const annualCostPerHa = ((Number(t.materialsCost) || 0) + (Number(t.servicesCost) || 0)) || (Number(t.annualCost) || 0);
      combined.annualCost += annualCostPerHa * area;
      model.outputs.forEach(o => {
        combined.deltas[o.id] += (Number(t.deltas[o.id]) || 0) * (area * adopt);
      });
    });
    let treatMetrics = null;
    if (combined.capitalCost || combined.annualCost || Object.values(combined.deltas).some(v => v !== 0)) {
      const temp = {
        id: "treatGroup",
        name: "Treatment group",
        area: combined.area || 1,
        adoption: 1,
        deltas: {},
        annualCost: 0,
        materialsCost: 0,
        servicesCost: 0,
        capitalCost: combined.capitalCost
      };
      model.outputs.forEach(o => {
        const perHaDelta = combined.area ? combined.deltas[o.id] / combined.area : 0;
        temp.deltas[o.id] = perHaDelta;
      });
      const annualCostPerHa = combined.area ? combined.annualCost / combined.area : 0;
      temp.annualCost = annualCostPerHa;
      treatMetrics = computeSingleTreatmentMetrics(temp, rate, years, 1, risk);
    }
    return { controlMetrics, treatMetrics };
  }

  function renderDepreciationSummary() {
    const root = $("#depSummary"); if (!root) return;
    root.innerHTML = "";
    const N = model.time.years;
    const baseYear = model.time.startYear;
    const totalPerYear = new Array(N + 1).fill(0);
    const rows = [];

    model.otherCosts.forEach(c => {
      if (c.type !== "capital") return;
      const method = c.depMethod || "none";
      if (method === "none") return;
      const cost = Number(c.capital) || 0;
      if (!cost) return;
      const life = Math.max(1, Number(c.depLife) || 5);
      const rate = Number(c.depRate) || 30;
      const startIndex = (Number(c.year) || baseYear) - baseYear;
      const sched = [];
      if (method === "straight") {
        const annual = cost / life;
        for (let i = 0; i < life; i++) {
          const idx = startIndex + i;
          if (idx >= 0 && idx <= N) {
            sched[idx] = (sched[idx] || 0) + annual;
            totalPerYear[idx] += annual;
          }
        }
      } else if (method === "declining") {
        let book = cost;
        for (let i = 0; i < life; i++) {
          const dep = book * rate / 100;
          const idx = startIndex + i;
          if (idx >= 0 && idx <= N) {
            sched[idx] = (sched[idx] || 0) + dep;
            totalPerYear[idx] += dep;
          }
          book -= dep;
          if (book <= 0) break;
        }
      }
      const firstDep = sched.find(v => v > 0) || 0;
      rows.push({
        label: c.label,
        method: method === "straight" ? "Straight-line" : "Declining-balance",
        life,
        rate: method === "declining" ? rate : "",
        firstDep
      });
    });

    if (!rows.length) {
      const p = document.createElement("p");
      p.className = "small muted";
      p.textContent = "No capital items with depreciation configured. Set Depreciation method on capital costs to see a schedule.";
      root.appendChild(p);
      return;
    }

    const tbl = document.createElement("table");
    tbl.className = "dep-table";
    tbl.innerHTML = `
      <thead>
        <tr>
          <th>Cost item</th>
          <th>Method</th>
          <th>Life (yrs)</th>
          <th>Rate (%/yr)</th>
          <th>Approx. first-year depreciation ($)</th>
        </tr>
      </thead>
      <tbody>
        ${rows.map(r => `
          <tr>
            <td>${esc(r.label)}</td>
            <td>${esc(r.method)}</td>
            <td>${r.life}</td>
            <td>${r.rate||""}</td>
            <td>${money(r.firstDep)}</td>
          </tr>
        `).join("")}
      </tbody>
    `;
    root.appendChild(tbl);
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

    const all = computeAll(rate, adoptMul, risk, model.sim.bcrMode);

    setVal("#pvBenefits", money(all.pvBenefits));
    setVal("#pvCosts", money(all.pvCosts));
    const npvEl = $("#npv");
    if (npvEl) {
      npvEl.textContent = money(all.npv);
      npvEl.className = "value " + (all.npv >= 0 ? "positive" : "negative");
    }
    setVal("#bcr", isFinite(all.bcr) ? fmt(all.bcr) : "—");
    setVal("#irr", isFinite(all.irrVal) ? percent(all.irrVal) : "—");
    setVal("#mirr", isFinite(all.mirrVal) ? percent(all.mirrVal) : "—");
    setVal("#roi", isFinite(all.roi) ? percent(all.roi) : "—");
    setVal("#grossMargin", money(all.annualGM));
    setVal("#profitMargin", isFinite(all.profitMargin) ? percent(all.profitMargin) : "—");
    setVal("#payback", all.paybackYears != null ? all.paybackYears : "Not reached");

    renderTreatmentSummary(rate, adoptMul, risk);

    const { controlMetrics, treatMetrics } = computeControlAndTreatmentGroupMetrics(rate, adoptMul, risk);
    if (controlMetrics) {
      setVal("#pvBenefitsControl", money(controlMetrics.pvBen));
      setVal("#pvCostsControl", money(controlMetrics.pvCost));
      setVal("#npvControl", money(controlMetrics.npv));
      setVal("#bcrControl", isFinite(controlMetrics.bcr)?fmt(controlMetrics.bcr):"—");
      setVal("#irrControl", isFinite(controlMetrics.irrVal)?percent(controlMetrics.irrVal):"—");
      setVal("#roiControl", isFinite(controlMetrics.roi)?percent(controlMetrics.roi):"—");
      setVal("#paybackControl", controlMetrics.pb!=null?controlMetrics.pb:"Not reached");
      setVal("#gmControl", money(controlMetrics.gm));
    } else {
      ["#pvBenefitsControl","#pvCostsControl","#npvControl","#bcrControl","#irrControl","#roiControl","#paybackControl","#gmControl"]
        .forEach(sel => setVal(sel, "—"));
    }
    if (treatMetrics) {
      setVal("#pvBenefitsTreat", money(treatMetrics.pvBen));
      setVal("#pvCostsTreat", money(treatMetrics.pvCost));
      setVal("#npvTreat", money(treatMetrics.npv));
      setVal("#bcrTreat", isFinite(treatMetrics.bcr)?fmt(treatMetrics.bcr):"—");
      setVal("#irrTreat", isFinite(treatMetrics.irrVal)?percent(treatMetrics.irrVal):"—");
      setVal("#roiTreat", isFinite(treatMetrics.roi)?percent(treatMetrics.roi):"—");
      setVal("#paybackTreat", treatMetrics.pb!=null?treatMetrics.pb:"Not reached");
      setVal("#gmTreat", money(treatMetrics.gm));
    } else {
      ["#pvBenefitsTreat","#pvCostsTreat","#npvTreat","#bcrTreat","#irrTreat","#roiTreat","#paybackTreat","#gmTreat"]
        .forEach(sel => setVal(sel, "—"));
    }

    renderTimeProjections(all.benefitByYear, all.costByYear, rate);
    renderDepreciationSummary();

    if ($("#simBcrTargetLabel")) $("#simBcrTargetLabel").textContent = model.sim.targetBCR;
  }

  function renderTimeProjections(benefitByYear, costByYear, rate) {
    const tblBody = $("#timeProjectionTable tbody");
    if (!tblBody) return;
    tblBody.innerHTML = "";
    const maxYears = model.time.years;
    const npvSeries = [];
    const usedHorizons = [];

    horizons.forEach(H => {
      const h = Math.min(H, maxYears);
      if (h <= 0) return;
      const b = benefitByYear.slice(0, h + 1);
      const c = costByYear.slice(0, h + 1);
      const pvB = presentValue(b, rate);
      const pvC = presentValue(c, rate);
      const npv = pvB - pvC;
      const bcr = pvC > 0 ? pvB / pvC : NaN;
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${h}</td>
        <td>${money(pvB)}</td>
        <td>${money(pvC)}</td>
        <td>${money(npv)}</td>
        <td>${isFinite(bcr)?fmt(bcr):"—"}</td>
      `;
      tblBody.appendChild(tr);
      npvSeries.push(npv);
      usedHorizons.push(h);
    });

    drawTimeSeries("timeNpvChart", usedHorizons, npvSeries);
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
    if ($("#simStatus")) $("#simStatus").textContent = "Running…";
    await new Promise(r => setTimeout(r));
    const N = model.sim.n;
    const seed = model.sim.seed;
    const rand = rng(seed ?? undefined);

    const discLow = model.time.discLow, discBase = model.time.discBase, discHigh = model.time.discHigh;
    const adoptLow = model.adoption.low, adoptBase = model.adoption.base, adoptHigh = model.adoption.high;
    const riskLow = model.risk.low, riskBase = model.risk.base, riskHigh = model.risk.high;

    const npvs = new Array(N);
    const bcrs = new Array(N);
    const details = [];
    const varPct = (model.sim.variationPct || 0) / 100;

    for (let i = 0; i < N; i++) {
      const r1 = rand(), r2 = rand(), r3 = rand();
      const disc = triangular(r1, discLow, discBase, discHigh);
      const adoptMul = clamp(triangular(r2, adoptLow, adoptBase, adoptHigh), 0, 1);
      const risk = clamp(triangular(r3, riskLow, riskBase, riskHigh), 0, 1);

      const shockOutputs = model.sim.varyOutputs ? (1 + (rand() * 2 * varPct - varPct)) : 1;
      const shockTreatCosts = model.sim.varyTreatCosts ? (1 + (rand() * 2 * varPct - varPct)) : 1;
      const shockInputCosts = model.sim.varyInputCosts ? (1 + (rand() * 2 * varPct - varPct)) : 1;

      const origOutValues = model.outputs.map(o => o.value);
      const origTreatAnn = model.treatments.map(t => t.annualCost);
      const origTreatMat = model.treatments.map(t => t.materialsCost || 0);
      const origTreatServ = model.treatments.map(t => t.servicesCost || 0);
      const origTreatCap = model.treatments.map(t => t.capitalCost);
      const origOcAnn = model.otherCosts.map(c => c.annual);
      const origOcCap = model.otherCosts.map(c => c.capital);

      try {
        if (model.sim.varyOutputs) {
          model.outputs.forEach((o, idx) => { o.value = origOutValues[idx] * shockOutputs; });
        }
        if (model.sim.varyTreatCosts) {
          model.treatments.forEach((t, idx) => {
            t.annualCost = origTreatAnn[idx] * shockTreatCosts;
            t.materialsCost = origTreatMat[idx] * shockTreatCosts;
            t.servicesCost = origTreatServ[idx] * shockTreatCosts;
          });
        }
        if (model.sim.varyInputCosts) {
          model.otherCosts.forEach((c, idx) => {
            c.annual = origOcAnn[idx] * shockInputCosts;
            c.capital = origOcCap[idx] * shockInputCosts;
          });
        }

        const { pvBenefits, pvCosts, bcr, npv } = computeAll(disc, adoptMul, risk, model.sim.bcrMode);
        npvs[i] = npv;
        bcrs[i] = bcr;
        details.push({ run: i + 1, discount: disc, adoption: adoptMul, risk, pvBenefits, pvCosts, npv, bcr });
      } finally {
        model.outputs.forEach((o, idx) => { o.value = origOutValues[idx]; });
        model.treatments.forEach((t, idx) => {
          t.annualCost = origTreatAnn[idx];
          t.materialsCost = origTreatMat[idx];
          t.servicesCost = origTreatServ[idx];
          t.capitalCost = origTreatCap[idx];
        });
        model.otherCosts.forEach((c, idx) => {
          c.annual = origOcAnn[idx];
          c.capital = origOcCap[idx];
        });
      }
    }

    model.sim.results = { npv: npvs, bcr: bcrs };
    model.sim.details = details;
    if ($("#simStatus")) $("#simStatus").textContent = "Done.";
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

  function drawTimeSeries(canvasId, xs, ys) {
    const canvas = document.getElementById(canvasId);
    if (!canvas || !xs.length || !ys.length) return;
    const ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    const padL = 60, padR = 16, padT = 10, padB = 34;
    const W = canvas.width - padL - padR;
    const H = canvas.height - padT - padB;

    const minX = Math.min(...xs), maxX = Math.max(...xs);
    const minY = Math.min(...ys), maxY = Math.max(...ys);
    const yMin = Math.min(minY, 0);
    const yMax = Math.max(maxY, 0) || 1;

    const xScale = v => padL + ((v - minX) / (maxX - minX || 1)) * W;
    const yScale = v => padT + H - ((v - yMin) / (yMax - yMin || 1)) * H;

    ctx.strokeStyle = "#3c6a52";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(padL, padT);
    ctx.lineTo(padL, padT + H);
    ctx.lineTo(padL + W, padT + H);
    ctx.stroke();

    const zeroY = yScale(0);
    ctx.strokeStyle = "rgba(200,220,210,0.6)";
    ctx.setLineDash([4,4]);
    ctx.beginPath();
    ctx.moveTo(padL, zeroY);
    ctx.lineTo(padL + W, zeroY);
    ctx.stroke();
    ctx.setLineDash([]);

    ctx.strokeStyle = "rgba(116,209,140,0.9)";
    ctx.lineWidth = 2;
    ctx.beginPath();
    xs.forEach((xv, idx) => {
      const x = xScale(xv);
      const y = yScale(ys[idx]);
      if (idx === 0) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    });
    ctx.stroke();

    ctx.fillStyle = "#c9efd6";
    ctx.font = "12px system-ui";
    ctx.textAlign = "center";
    xs.forEach(xv => {
      const x = xScale(xv);
      ctx.fillText(String(xv), x, padT + H + 18);
    });
    ctx.textAlign = "right";
    ctx.fillText(money(yMax), padL - 6, yScale(yMax) + 4);
    ctx.fillText(money(0), padL - 6, zeroY + 4);
    ctx.fillText(money(yMin), padL - 6, yScale(yMin) + 4);
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
    const all = computeAll(rate, adoptMul, risk, model.sim.bcrMode);

    return {
      meta: {
        name: model.project.name,
        lead: model.project.lead,
        analysts: model.project.analysts,
        team: model.project.team,
        organisation: model.project.organisation,
        contact: model.project.contactEmail,
        phone: model.project.contactPhone,
        updated: model.project.lastUpdated
      },
      params: {
        startYear: model.time.startYear,
        projectStartYear: model.time.projectStartYear,
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
      ["Project lead", s.meta.lead],
      ["Analysts", s.meta.analysts],
      ["Project team", s.meta.team],
      ["Organisation", s.meta.organisation],
      ["Contact email", s.meta.contact],
      ["Contact phone", s.meta.phone],
      ["Last Updated", s.meta.updated],
      [],
      ["Analysis Start Year", s.params.startYear],
      ["Project Start Year", s.params.projectStartYear],
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

    const treatHeader = ["Treatment","Area(ha)","Adoption","Annual Benefit","Annual Cost","PV Benefit","PV Cost","BCR","NPV"];
    const treatRows = [treatHeader];
    const rate = model.time.discBase;
    const adoptMul = model.adoption.base;
    const risk = model.risk.base;
    model.treatments.forEach(t => {
      const m = computeSingleTreatmentMetrics(t, rate, model.time.years, adoptMul, risk);
      const adopt = clamp(t.adoption * adoptMul, 0, 1);
      const area = Number(t.area) || 0;
      const annualCostPerHa = ((Number(t.materialsCost) || 0) + (Number(t.servicesCost) || 0)) || (Number(t.annualCost) || 0);
      const annualCost = annualCostPerHa * area;
      const annualBen = m.gm + annualCost;
      treatRows.push([t.name, area, adopt, annualBen, annualCost, m.pvBen, m.pvCost, m.bcr, m.npv]);
    });
    downloadFile(`cba_treatments_${slug(s.meta.name)}.csv`, toCsv(treatRows));

    const benRows = [["Label","Category","BenefitType","Frequency","StartYear","EndYear","Year","UnitValue","Quantity","Abatement","AnnualAmount","GrowthPct","LinkAdoption","LinkRisk","P0","P1","Consequence","Notes"]];
    model.benefits.forEach(b => benRows.push([b.label,b.category,b.theme||"",b.frequency,b.startYear,b.endYear,b.year,b.unitValue,b.quantity,b.abatement,b.annualAmount,b.growthPct,b.linkAdoption,b.linkRisk,b.p0,b.p1,b.consequence,b.notes]));
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
      const m = computeSingleTreatmentMetrics(t, model.time.discBase, model.time.years, model.adoption.base, model.risk.base);
      const adopt = clamp(t.adoption * model.adoption.base, 0, 1);
      const area = Number(t.area) || 0;
      const annualCostPerHa = ((Number(t.materialsCost) || 0) + (Number(t.servicesCost) || 0)) || (Number(t.annualCost) || 0);
      const annualCost = annualCostPerHa * area;
      const annualBen = m.gm + annualCost;
      return `<tr>
        <td>${esc(t.name)}${t.isControl?" (Control)":""}</td><td>${fmt(area)}</td><td>${fmt(adopt)}</td>
        <td>${money(annualBen)}</td><td>${money(annualCost)}</td>
        <td>${money(m.pvBen)}</td><td>${money(m.pvCost)}</td>
        <td>${isFinite(m.bcr)?fmt(m.bcr):"—"}</td><td>${money(m.npv)}</td>
      </tr>`;
    }).join("");

    const benRows = model.benefits.map(b => `
      <tr><td>${esc(b.label)}</td><td>${b.category}</td><td>${esc(b.theme||"")}</td><td>${b.frequency}</td>
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
              <tr><th>Project lead</th><td>${esc(s.meta.lead)}</td></tr>
              <tr><th>Analysts</th><td>${esc(s.meta.analysts)}</td></tr>
              <tr><th>Organisation</th><td>${esc(s.meta.organisation)}</td></tr>
              <tr><th>Contact</th><td><a href="mailto:${esc(s.meta.contact)}">${esc(s.meta.contact)}</a> ${s.meta.phone?(" / "+esc(s.meta.phone)):""}</td></tr>
              <tr><th>Last Updated</th><td>${esc(s.meta.updated)}</td></tr>
            </table>
          </div>
          <div>
            <h2>Parameters</h2>
            <table>
              <tr><th>Analysis Start Year</th><td>${s.params.startYear}</td></tr>
              <tr><th>Project Start Year</th><td>${s.params.projectStartYear}</td></tr>
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
            <th>Label</th><th>Cat</th><th>Type</th><th>Freq</th><th>Start</th><th>End</th><th>Year</th>
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
    const file = $("#excelFile")?.files?.[0];
    const status = $("#loadStatus");
    const alertBox = $("#validation");
    const preview = $("#preview");
    parsedExcel = null;
    if (alertBox) {
      alertBox.classList.remove("show");
      alertBox.innerHTML = "";
    }
    if (preview) preview.innerHTML = "";
    if ($("#importExcel")) $("#importExcel").disabled = true;

    if (!file) {
      if (status) status.textContent = "Select an Excel/CSV file first.";
      return;
    }
    if (status) status.textContent = "Parsing…";

    try {
      const buf = await file.arrayBuffer();
      let wb;
      if (file.name.toLowerCase().endsWith(".csv")) {
        const csvTxt = new TextDecoder().decode(new Uint8Array(buf));
        const ws = XLSX.utils.csv_to_sheet(csvTxt);
        wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, $("#csvSheetType")?.value || "Outputs");
      } else {
        wb = XLSX.read(buf, { type: "array" });
      }

      const getSheet = (name) => {
        const s = wb.Sheets[name];
        if (!s) return null;
        return XLSX.utils.sheet_to_json(s, { defval: "", raw: true });
      };

      const outputs = getSheet("Outputs") || [];
      const treatments = getSheet("Treatments") || getSheet("Treatment") || [];
      const benefits = getSheet("Benefits") || getSheet("Benefit") || [];
      const costs = getSheet("Costs") || getSheet("Cost") || [];

      parsedExcel = { outputs, treatments, benefits, costs };

      const issues = [];
      if (!outputs.length) issues.push("No rows found in sheet “Outputs”. At least one output row is recommended.");
      if (!treatments.length) issues.push("No rows found in sheet “Treatments”. You can still import outputs, benefits, and costs.");
      if (!benefits.length) issues.push("No rows found in sheet “Benefits”.");
      if (!costs.length) issues.push("No rows found in sheet “Costs”.");

      if (alertBox && issues.length) {
        alertBox.innerHTML = `<ul>${issues.map(x => `<li>${esc(x)}</li>`).join("")}</ul>`;
        alertBox.classList.add("show");
      }

      if (preview) {
        const parts = [];
        parts.push(`<p><strong>Detected sheets</strong></p>`);
        parts.push("<ul>");
        parts.push(`<li>Outputs: ${outputs.length} row(s)</li>`);
        parts.push(`<li>Treatments: ${treatments.length} row(s)</li>`);
        parts.push(`<li>Benefits: ${benefits.length} row(s)</li>`);
        parts.push(`<li>Costs: ${costs.length} row(s)</li>`);
        parts.push("</ul>");

        const mkTable = (title, rows, maxRows = 3) => {
          if (!rows.length) return `<h4>${title}</h4><p class="small muted">No rows.</p>`;
          const cols = Object.keys(rows[0]);
          const head = cols.map(c => `<th>${esc(c)}</th>`).join("");
          const body = rows.slice(0, maxRows).map(r =>
            `<tr>${cols.map(c => `<td>${esc(r[c])}</td>`).join("")}</tr>`
          ).join("");
          return `
            <h4>${title}</h4>
            <table class="mini">
              <thead><tr>${head}</tr></thead>
              <tbody>${body}</tbody>
            </table>
          `;
        };

        parts.push(mkTable("Outputs (preview)", outputs));
        parts.push(mkTable("Treatments (preview)", treatments));
        parts.push(mkTable("Benefits (preview)", benefits));
        parts.push(mkTable("Costs (preview)", costs));

        preview.innerHTML = parts.join("");
      }

      if (status) {
        if (!outputs.length && !treatments.length && !benefits.length && !costs.length) {
          status.textContent = "Parsed file but no recognised sheets found. Expected: Outputs, Treatments, Benefits, Costs.";
        } else {
          status.textContent = "File parsed. Review preview then click “Import to model”.";
        }
      }

      if ($("#importExcel") && (outputs.length || treatments.length || benefits.length || costs.length)) {
        $("#importExcel").disabled = false;
      }
    } catch (err) {
      if (status) status.textContent = "Could not parse file.";
      if (alertBox) {
        alertBox.innerHTML = `<p class="error">Error: ${esc(err?.message || String(err))}</p>`;
        alertBox.classList.add("show");
      }
    }
  }

  function commitExcelToModel() {
    if (!parsedExcel) {
      alert("No parsed Excel data. Parse a file first.");
      return;
    }
    const { outputs, treatments, benefits, costs } = parsedExcel;
    const toBool = v => {
      const s = String(v).trim().toLowerCase();
      return s === "true" || s === "yes" || s === "y" || s === "1" || s === "t";
    };

    if (outputs && outputs.length) {
      model.outputs = outputs.map(row => ({
        id: uid(),
        name: row.Name || row.Output || row.output || "",
        unit: row.Unit || row.UnitOfMeasure || row.Unit_ || "",
        value: Number(row.ValuePerUnit ?? row.Value ?? row["$/unit"] ?? 0) || 0,
        source: row.Source || "Input Directly"
      }));
    }

    if (treatments && treatments.length) {
      model.treatments = treatments.map(row => ({
        id: uid(),
        name: row.Name || row.Treatment || "",
        area: Number(row.AreaHa ?? row.Area ?? 0) || 0,
        adoption: Number(row.Adoption_0to1 ?? row.Adoption ?? 0.5) || 0,
        deltas: {},
        annualCost: Number(row.AnnualCost_perHa ?? row.AnnualCost ?? 0) || 0,
        materialsCost: Number(row.MaterialsCost_perHa ?? row.MaterialsCost ?? 0) || 0,
        servicesCost: Number(row.ServicesCost_perHa ?? row.ServicesCost ?? 0) || 0,
        capitalCost: Number(row.CapitalCost_y0 ?? row.CapitalCost ?? 0) || 0,
        constrained: row.Constrained_trueFalse !== undefined ? toBool(row.Constrained_trueFalse) : true,
        source: row.Source || "Input Directly",
        replications: Number(row.Replications ?? 1) || 1,
        isControl: row.IsControl_trueFalse !== undefined ? toBool(row.IsControl_trueFalse) : false,
        notes: row.Notes || ""
      }));
    }

    if (benefits && benefits.length) {
      model.benefits = benefits.map(row => ({
        id: uid(),
        label: row.Label || "",
        category: row.Category || "C4",
        theme: row.Theme || row.BenefitType || "Other",
        frequency: row.Frequency || "Annual",
        startYear: Number(row.StartYear ?? model.time.startYear) || model.time.startYear,
        endYear: Number(row.EndYear ?? row.StartYear ?? model.time.startYear) || model.time.startYear,
        year: Number(row.YearOnce ?? row.Year ?? row.StartYear ?? model.time.startYear) || model.time.startYear,
        unitValue: Number(row.UnitValue ?? 0) || 0,
        quantity: Number(row.Quantity ?? 0) || 0,
        abatement: Number(row.Abatement ?? 0) || 0,
        annualAmount: Number(row.AnnualAmount ?? 0) || 0,
        growthPct: Number(row.GrowthPct ?? 0) || 0,
        linkAdoption: row.LinkAdoption_trueFalse !== undefined ? toBool(row.LinkAdoption_trueFalse) : true,
        linkRisk: row.LinkRisk_trueFalse !== undefined ? toBool(row.LinkRisk_trueFalse) : true,
        p0: Number(row.P0 ?? 0) || 0,
        p1: Number(row.P1 ?? 0) || 0,
        consequence: Number(row.Consequence ?? 0) || 0,
        notes: row.Notes || ""
      }));
    }

    if (costs && costs.length) {
      model.otherCosts = costs.map(row => ({
        id: uid(),
        label: row.Label || "",
        type: (row.Type_annualCapital || row.Type || "annual").toString().toLowerCase().startsWith("cap") ? "capital" : "annual",
        category: row.Category || "Labour",
        annual: Number(row.Annual ?? 0) || 0,
        startYear: Number(row.StartYear ?? model.time.startYear) || model.time.startYear,
        endYear: Number(row.EndYear ?? row.StartYear ?? model.time.startYear) || model.time.startYear,
        capital: Number(row.Capital ?? 0) || 0,
        year: Number(row.CapitalYear ?? row.Year ?? model.time.startYear) || model.time.startYear,
        constrained: row.Constrained_trueFalse !== undefined ? toBool(row.Constrained_trueFalse) : true,
        depMethod: (row.DepMethod || "none").toString().toLowerCase().startsWith("str") ? "straight"
          : (row.DepMethod || "none").toString().toLowerCase().startsWith("dec") ? "declining"
          : "none",
        depLife: Number(row.DepLife ?? 5) || 5,
        depRate: Number(row.DepRate ?? 30) || 30
      }));
    }

    // Reset deltas for all treatments against current outputs
    initTreatmentDeltas();

    renderAll();
    setBasicsFieldsFromModel();
    calcAndRender();

    if ($("#loadStatus")) $("#loadStatus").textContent = "Imported data from file into the model.";
    if ($("#importExcel")) $("#importExcel").disabled = true;
  }

  function downloadExcelTemplate() {
    if (!window.XLSX) {
      alert("SheetJS (XLSX) library not loaded.");
      return;
    }
    const wb = XLSX.utils.book_new();

    const outRows = [
      ["Name","Unit","ValuePerUnit","Source"],
      ["Yield","t/ha","","Input Directly"],
      ["Protein","%-point","","Input Directly"],
      ["Moisture","%-point","","Input Directly"],
      ["Biomass","t/ha","","Input Directly"]
    ];
    const wsOut = XLSX.utils.aoa_to_sheet(outRows);
    XLSX.utils.book_append_sheet(wb, wsOut, "Outputs");

    const treatRows = [
      ["Name","AreaHa","Adoption_0to1","MaterialsCost_perHa","ServicesCost_perHa","AnnualCost_perHa","CapitalCost_y0","Constrained_trueFalse","Source","Replications","IsControl_trueFalse","Notes"],
      ["Optimized N (Rate+Timing)",300,0.8,0,0,45,5000,"TRUE","Farm Trials",1,"FALSE",""],
      ["Slow-Release N",200,0.7,0,0,25,0,"TRUE","ABARES",1,"FALSE",""]
    ];
    const wsTr = XLSX.utils.aoa_to_sheet(treatRows);
    XLSX.utils.book_append_sheet(wb, wsTr, "Treatments");

    const benRows = [
      ["Label","Category","Theme","Frequency","StartYear","EndYear","YearOnce","UnitValue","Quantity","Abatement","AnnualAmount","GrowthPct","LinkAdoption_trueFalse","LinkRisk_trueFalse","P0","P1","Consequence","Notes"],
      ["Reduced recurring costs (energy/water)","C4","Cost savings","Annual",model.time.startYear,model.time.startYear+4,"",0,0,0,15000,0,"TRUE","TRUE",0,0,0,"Project-wide OPEX saving"],
      ["Reduced risk of quality downgrades","C7","Risk reduction","Annual",model.time.startYear,model.time.startYear+9,"",0,0,0,0,0,"TRUE","FALSE",0.1,0.07,120000,""],
      ["Soil asset value uplift (carbon/structure)","C6","Soil carbon","Once",model.time.startYear,model.time.startYear,model.time.startYear+5,0,0,0,50000,0,"FALSE","TRUE",0,0,0,""]
    ];
    const wsBen = XLSX.utils.aoa_to_sheet(benRows);
    XLSX.utils.book_append_sheet(wb, wsBen, "Benefits");

    const costRows = [
      ["Label","Type_annualCapital","Category","Annual","StartYear","EndYear","Capital","CapitalYear","DepMethod","DepLife","DepRate","Constrained_trueFalse"],
      ["Project Mgmt & M&E","annual","Services",20000,model.time.startYear,model.time.startYear+4,0,model.time.startYear,"none",5,30,"TRUE"]
    ];
    const wsCost = XLSX.utils.aoa_to_sheet(costRows);
    XLSX.utils.book_append_sheet(wb, wsCost, "Costs");

    saveWorkbook("cba_template.xlsx", wb);
  }

  function downloadSampleDataset() {
    if (!window.XLSX) {
      alert("SheetJS (XLSX) library not loaded.");
      return;
    }
    const wb = XLSX.utils.book_new();

    const outRows = [["Name","Unit","ValuePerUnit","Source"]];
    model.outputs.forEach(o => {
      outRows.push([o.name, o.unit, o.value, o.source]);
    });
    const wsOut = XLSX.utils.aoa_to_sheet(outRows);
    XLSX.utils.book_append_sheet(wb, wsOut, "Outputs");

    const treatRows = [["Name","AreaHa","Adoption_0to1","MaterialsCost_perHa","ServicesCost_perHa","AnnualCost_perHa","CapitalCost_y0","Constrained_trueFalse","Source","Replications","IsControl_trueFalse","Notes"]];
    model.treatments.forEach(t => {
      treatRows.push([
        t.name,
        t.area,
        t.adoption,
        t.materialsCost || 0,
        t.servicesCost || 0,
        t.annualCost || 0,
        t.capitalCost || 0,
        t.constrained ? "TRUE" : "FALSE",
        t.source || "Input Directly",
        t.replications || 1,
        t.isControl ? "TRUE" : "FALSE",
        t.notes || ""
      ]);
    });
    const wsTr = XLSX.utils.aoa_to_sheet(treatRows);
    XLSX.utils.book_append_sheet(wb, wsTr, "Treatments");

    const benRows = [["Label","Category","Theme","Frequency","StartYear","EndYear","YearOnce","UnitValue","Quantity","Abatement","AnnualAmount","GrowthPct","LinkAdoption_trueFalse","LinkRisk_trueFalse","P0","P1","Consequence","Notes"]];
    model.benefits.forEach(b => {
      benRows.push([
        b.label, b.category, b.theme || "", b.frequency,
        b.startYear, b.endYear, b.year,
        b.unitValue, b.quantity, b.abatement,
        b.annualAmount, b.growthPct,
        b.linkAdoption ? "TRUE" : "FALSE",
        b.linkRisk ? "TRUE" : "FALSE",
        b.p0, b.p1, b.consequence, b.notes || ""
      ]);
    });
    const wsBen = XLSX.utils.aoa_to_sheet(benRows);
    XLSX.utils.book_append_sheet(wb, wsBen, "Benefits");

    const costRows = [["Label","Type_annualCapital","Category","Annual","StartYear","EndYear","Capital","CapitalYear","DepMethod","DepLife","DepRate","Constrained_trueFalse"]];
    model.otherCosts.forEach(c => {
      costRows.push([
        c.label,
        c.type === "capital" ? "capital" : "annual",
        c.category,
        c.annual || 0,
        c.startYear,
        c.endYear,
        c.capital || 0,
        c.year,
        c.depMethod || "none",
        c.depLife || 5,
        c.depRate || 30,
        c.constrained ? "TRUE" : "FALSE"
      ]);
    });
    const wsCost = XLSX.utils.aoa_to_sheet(costRows);
    XLSX.utils.book_append_sheet(wb, wsCost, "Costs");

    saveWorkbook(`cba_sample_${slug(model.project.name)}.xlsx`, wb);
  }

  // ---------- INIT ----------
  document.addEventListener("DOMContentLoaded", () => {
    initTabs();
    initActions();
    bindBasics();
    renderAll();
    calcAndRender();
  });
})();
