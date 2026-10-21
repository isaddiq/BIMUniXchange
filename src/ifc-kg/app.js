/* BIMUniXchange · BIM Knowledge Graph Studio
 * app.js — UI controller: file loading, lenses, filters, inspector, search,
 * context menu, query builder, analytics / health / mapping / catalogue
 * drawers and exports.
 */
(function (KG) {
  "use strict";
  const S = KG.Schema;
  const V = KG.Views;
  const A = KG.Analytics;
  const X = KG.Export;

  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
  const short = (cls) => (cls || "").replace(/^Ifc/, "");
  const idOf = (x) => (typeof x === "object" && x ? x.id : x);
  const debounce = (fn, ms) => {
    let t;
    return (...a) => {
      clearTimeout(t);
      t = setTimeout(() => fn(...a), ms);
    };
  };
  function fmt(v, d = 3) {
    if (v == null || v === "") return "";
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    if (typeof v !== "number") return String(v);
    if (!Number.isFinite(v)) return String(v);
    if (Number.isInteger(v)) return v.toLocaleString();
    if (Math.abs(v) > 0 && Math.abs(v) < 0.001) return v.toExponential(2);
    return (+v.toFixed(d)).toLocaleString(undefined, { maximumFractionDigits: d });
  }
  const pct = (v) => (v == null ? "" : (v * 100).toFixed(v < 0.1 && v > 0 ? 1 : 0) + "%");
  const bytes = (b) => (b > 1e6 ? (b / 1e6).toFixed(1) + " MB" : (b / 1e3).toFixed(0) + " kB");

  const CAT_PALETTE = ["#4e79a7", "#f28e2b", "#e15759", "#76b7b2", "#59a14f", "#edc948", "#b07aa1", "#ff9da7", "#9c755f", "#17becf", "#1f77b4", "#2ca02c", "#d62728", "#9467bd", "#8c564b", "#e377c2", "#bcbd22", "#aec7e8", "#ffbb78", "#98df8a"];
  const seq = (t) => d3.interpolateViridis(0.12 + 0.84 * Math.max(0, Math.min(1, t)));
  const DEFAULT_KINDS = ["Project", "Spatial", "Element", "Type", "Material", "Classification", "ClassSystem", "System", "Zone", "Group", "Document", "Process", "Actor", "Control", "Resource", "Exterior"];
  const BASE_R = { Project: 10, Spatial: 6.5, Element: 3.6, Type: 5, PropertySet: 5, QuantitySet: 5, Property: 3, Value: 3.4, Material: 5, Classification: 4.4, ClassSystem: 7, System: 7.5, Zone: 6.5, Group: 6, Port: 2.2, Opening: 2.6, Document: 4, Process: 5, Actor: 5, Control: 5, Resource: 5, Class: 5, Aggregate: 4.5, Exterior: 9, Other: 4 };
  const SPATIAL_R = { IfcSite: 9, IfcBuilding: 8.5, IfcBuildingStorey: 7.5, IfcSpace: 4.8 };

  const ICONS = {
    kg: '<circle cx="6" cy="6" r="2.2"/><circle cx="18" cy="7" r="2.2"/><circle cx="12" cy="17" r="2.2"/><circle cx="20" cy="17" r="1.6"/><path d="M8 7l8 0M7 8l4 7M17 9l-4 6M14 17h4.4"/>',
    spatial: '<path d="M12 3v4M12 7l-6 5M12 7l6 5M6 12v4M18 12v4"/><rect x="10" y="2" width="4" height="3" rx="1"/><rect x="4" y="16" width="4" height="4" rx="1"/><rect x="16" y="16" width="4" height="4" rx="1"/>',
    breakdown: '<path d="M4 4h6v4H4zM14 10h6v4h-6zM14 17h6v4h-6zM7 8v11h7M7 12h7"/>',
    systems: '<path d="M3 12h4l2-4 4 8 2-4h6"/><circle cx="3" cy="12" r="1"/><circle cx="21" cy="12" r="1"/>',
    connectivity: '<rect x="3" y="4" width="7" height="7" rx="1"/><rect x="14" y="13" width="7" height="7" rx="1"/><path d="M10 8h4v5"/>',
    spaces: '<rect x="3" y="3" width="8" height="8" rx="1"/><rect x="13" y="3" width="8" height="8" rx="1"/><rect x="3" y="13" width="18" height="8" rx="1"/><path d="M11 7h2M7 11v2"/>',
    types: '<path d="M12 3l8 4-8 4-8-4z"/><path d="M4 12l8 4 8-4M4 17l8 4 8-4"/>',
    materials: '<path d="M4 6h16M4 10h16M4 14h16M4 18h16"/>',
    properties: '<rect x="4" y="3" width="16" height="18" rx="2"/><path d="M8 8h8M8 12h8M8 16h5"/>',
    schema: '<rect x="3" y="3" width="7" height="5" rx="1"/><rect x="14" y="3" width="7" height="5" rx="1"/><rect x="8.5" y="16" width="7" height="5" rx="1"/><path d="M6.5 8v3h11V8M12 11v5"/>',
    subgraph: '<circle cx="12" cy="12" r="3"/><circle cx="12" cy="12" r="8" stroke-dasharray="2 2"/>',
  };
  const icon = (k, s = 15) => `<svg width="${s}" height="${s}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${ICONS[k] || ""}</svg>`;

  function shapeIcon(kind, color, size = 14) {
    const shape = (S.KINDS[kind] || S.KINDS.Other).shape;
    const sh = shape === "ring" || shape === "double" || shape === "dot" ? "circle" : shape;
    const hollow = shape === "ring" || shape === "squareo";
    const c = size / 2;
    return `<svg class="swatch-shape" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">${KG.Renderer.shapeSVG(sh, c, c, size * 0.33)} fill="${hollow ? "none" : color}" stroke="${color}" stroke-width="${hollow ? 1.8 : 0}"/></svg>`;
  }

  // ── State ───────────────────────────────────────────────────────────────
  const state = {
    model: null,
    graph: null,
    analysis: null,
    health: null,
    healthIndex: new Map(),
    searchIndex: [],
    settings: null,
    style: { layout: "force", colorBy: "domain", sizeBy: "degree", labels: "auto", arrows: true, particles: true, inferred: true, edgeColor: true, colorProp: "" },
    selection: null,
    pathStart: null,
    drawerTab: "overview",
    legend: null,
    tablePage: 0,
    tableSort: { key: "degree", dir: -1 },
    queryResults: null,
    exportScope: "visible",
  };

  function defaultSettings() {
    return {
      lens: "kg",
      detail: "auto",
      maxNodes: 3000,
      kinds: new Set(DEFAULT_KINDS),
      cats: null,
      domains: null,
      classes: null,
      storeys: null,
      hidden: new Set(),
      subgraph: null,
      valueKey: "",
      showPorts: false,
      showOpenings: false,
      showReferences: false,
      showZones: true,
      showBoundaries: false,
      includeInstances: false,
      includeClassification: true,
      includeProps: false,
      includeTaxonomy: false,
      breakdown: ["Discipline", "Class", "Type"],
    };
  }
  state.settings = defaultSettings();

  function theme() {
    const t = document.documentElement.dataset.theme;
    if (t) return t;
    return window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
  }

  // ── Renderer ───────────────────────────────────────────────────────────
  let renderer;
  const mouse = { x: 0, y: 0 };
  let lastClick = { id: null, t: 0 };

  function initRenderer() {
    renderer = new KG.Renderer.Renderer($("graph"), {
      theme,
      onClick: (n, ev) => {
        hideCtx();
        const now = Date.now();
        if (lastClick.id === n.id && now - lastClick.t < 350) {
          renderer.focus(n.id);
          if (n.ref) exploreNeighborhood(n.ref, 1);
        } else select(n.id, { from: "graph" });
        lastClick = { id: n.id, t: now };
      },
      onRightClick: (n, ev) => openCtx(n, ev),
      onHover: (n) => showTip(n),
      onLinkHover: (l) => showLinkTip(l),
      onBackground: () => {
        hideCtx();
        select(null);
      },
      storeyCount: () => state.model?.storeys.length || 0,
      storeyIndexOf: (n) => {
        const id = n.cls === "IfcBuildingStorey" ? n.ref : n.storey;
        const s = id ? state.model.nodes.get(id) : null;
        return s ? s.storeyIndex : null;
      },
      groupOf: (n) => n._g,
    });
  }

  // ── Loading ────────────────────────────────────────────────────────────
  function showLoading(title, file) {
    $("loadTitle").textContent = title;
    $("loadFile").textContent = file || "";
    $("loadBar").style.width = "0%";
    $("loadText").textContent = "";
    $("loading").classList.add("show");
  }
  async function progress(frac, text) {
    $("loadBar").style.width = Math.round(frac * 100) + "%";
    if (text) $("loadText").textContent = text;
    await new Promise((r) => setTimeout(r, 0));
  }
  const hideLoading = () => $("loading").classList.remove("show");

  async function readFile(file) {
    const name = file.name || "model.ifc";
    if (/\.ifczip$|\.zip$/i.test(name)) {
      await loadJSZip();
      const zip = await JSZip.loadAsync(file);
      const entry = Object.values(zip.files).find((f) => /\.ifc$/i.test(f.name));
      if (!entry) throw new Error("No .ifc file found inside the archive");
      return { text: await entry.async("string"), name: entry.name, kind: "ifc", size: file.size };
    }
    const text = await file.text();
    const kind = /\.csv$/i.test(name) ? "csv" : "ifc";
    return { text, name, kind, size: file.size };
  }
  function loadJSZip() {
    if (window.JSZip) return Promise.resolve();
    return new Promise((res, rej) => {
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
      s.onload = res;
      s.onerror = () => rej(new Error("Could not load JSZip for .ifczip support"));
      document.head.appendChild(s);
    });
  }

  async function loadFromText({ text, name, kind, size }) {
    showLoading(kind === "csv" ? "Mapping CSV to the knowledge graph" : "Parsing IFC model", `${name} · ${bytes(size || text.length)}`);
    const t0 = performance.now();
    try {
      let model;
      if (kind === "csv") {
        model = await KG.CSV.buildFromCSV(text, { fileName: name, fileSize: size, onProgress: progress });
      } else {
        if (!/ISO-10303-21/.test(text.slice(0, 2000))) throw new Error("This does not look like an IFC STEP file (missing ISO-10303-21 header).");
        const step = new KG.Step.StepFile(text);
        await step.index((f) => progress(f * 0.4, `Indexed ${step.size.toLocaleString()} STEP instances`));
        model = await KG.Model.buildFromIFC(step, { fileName: name, fileSize: size, onProgress: (f, msg) => progress(0.4 + f * 0.55, msg) });
      }
      model.meta.loadMs = performance.now() - t0;
      await progress(0.97, "Auditing model health");
      onModelLoaded(model);
    } catch (err) {
      console.error(err);
      toast("Could not read file: " + err.message, 6000);
    } finally {
      hideLoading();
    }
  }

  async function openFile(file) {
    try {
      showLoading("Reading file", file.name);
      const data = await readFile(file);
      await loadFromText(data);
    } catch (err) {
      hideLoading();
      toast("Could not read file: " + err.message, 6000);
    }
  }

  const SAMPLES = {
    ifc: { file: "sample-building.ifc", kind: "ifc", title: "Loading sample building (IFC4)" },
    revit: { file: "revit.csv", kind: "csv", title: "Loading Revit CSV sample" },
    archicad: { file: "archicad.csv", kind: "csv", title: "Loading ArchiCAD CSV sample" },
  };
  async function loadSample(which = "ifc") {
    const smp = SAMPLES[which] || SAMPLES.ifc;
    showLoading(smp.title, smp.file);
    try {
      const res = await fetch("../sample/" + smp.file);
      if (!res.ok) throw new Error("HTTP " + res.status);
      const text = await res.text();
      await loadFromText({ text, name: smp.file, kind: smp.kind, size: text.length });
    } catch (err) {
      hideLoading();
      toast("Sample could not be loaded (" + err.message + "). If you opened the page from disk, serve it over HTTP or open a file instead.", 7000);
    }
  }

  // ── Model loaded ───────────────────────────────────────────────────────
  function onModelLoaded(model) {
    state.model = model;
    state.health = KG.Health.run(model);
    state.healthIndex = new Map();
    const rank = { error: 3, warning: 2, info: 1 };
    for (const c of state.health.checks) {
      for (const id of c.affected) {
        const cur = state.healthIndex.get(id);
        if (!cur || rank[c.severity] > rank[cur]) state.healthIndex.set(id, c.severity);
      }
    }
    state.settings = defaultSettings();
    state.selection = null;
    state.pathStart = null;
    state.queryResults = null;
    renderer.marked = new Set();
    renderer.pos.clear();
    state.settings.breakdown = model.meta.csvSource === "Revit" ? ["Category", "Family", "Type"] : model.meta.csvSource === "ArchiCAD" ? ["Category", "Layer"] : model.source === "csv" ? ["Category", "Class"] : ["Discipline", "Class", "Type"];
    state.settings.lens = model.source === "csv" ? "breakdown" : model.nodes.size <= 3000 ? "kg" : "spatial";
    state.style.layout = V.LENSES.find((l) => l.id === state.settings.lens).layout;
    buildSearchIndex();

    $("emptyState").hidden = true;
    document.querySelectorAll(".needs-model").forEach((el) => (el.hidden = false));
    $("searchBox").hidden = false;
    $("viewStats").hidden = false;
    $("stageTools").hidden = false;
    $("legend").hidden = false;
    $("legend").classList.toggle("collapsed", $("stage").clientWidth < 700 || $("stage").clientHeight < 520);
    $("btnExport").disabled = false;
    $("fileChip").hidden = false;
    $("fileName").textContent = model.meta.fileName;
    $("schemaPill").textContent = model.meta.schema || "IFC";
    $("fileMeta").textContent = `${bytes(model.meta.fileSize)} · ${(model.meta.entityCount || 0).toLocaleString()} ${model.source === "csv" ? "rows" : "instances"} · ${(model.meta.loadMs / 1000).toFixed(1)} s`;
    document.title = `${model.meta.fileName} · BIM Knowledge Graph Studio`;

    $("propKeys").innerHTML = [...model.catalog.props.keys()].sort().slice(0, 4000).map((k) => `<option value="${esc(k)}">`).join("");
    syncControls();
    renderLensList();
    renderFilters();
    renderLensOptions();
    renderDrawerTabs();
    if (window.innerHeight > 780) openDrawer("overview");
    else renderDrawer();
    rebuild();
    renderInspector();
    const inferred = model.meta.counts.inferred;
    toast(`Knowledge graph ready: ${model.meta.counts.nodes.toLocaleString()} nodes, ${model.meta.counts.edges.toLocaleString()} edges${inferred ? ` (${inferred.toLocaleString()} inferred)` : ""}`);
  }

  function syncControls() {
    $("layoutSel").value = state.style.layout;
    $("colorSel").value = state.style.colorBy;
    $("sizeSel").value = state.style.sizeBy;
    $("labelSel").value = state.style.labels;
    $("detailSel").value = state.settings.detail;
    $("maxNodes").value = state.settings.maxNodes;
    $("colorPropField").hidden = state.style.colorBy !== "property";
    $("colorProp").value = state.style.colorProp || "";
  }

  // ── Build & render the current view ─────────────────────────────────────
  function rebuild(opts = {}) {
    if (!state.model) return;
    const s = state.settings;
    const g = V.build(state.model, s);
    if (!state.style.inferred) g.links = g.links.filter((l) => !l.inferred);
    state.graph = g;
    const n = g.nodes.length;
    state.analysis = A.analyze(g, { full: n <= 9000, communities: n <= 25000 });
    if (s.lens === "spaces") spaceSyntax();
    computeStyle();
    renderer.style = { ...renderer.style, layout: state.style.layout, labels: state.style.labels, arrows: state.style.arrows, particles: state.style.particles };
    renderer.setData(g, { fit: opts.fit !== false });
    if (state.selection && !g.nodes.some((x) => x.id === state.selection)) {
      const viaRef = g.nodes.find((x) => x.ref === state.selection);
      state.selection = viaRef ? viaRef.id : state.selection;
    }
    renderer.select(g.nodes.some((x) => x.id === state.selection) ? state.selection : null);
    updateStats();
    updateNotes();
    updateLegend();
    renderLensList();
    if (["analytics", "table"].includes(state.drawerTab)) renderDrawer();
    if (opts.inspect !== false) renderInspector();
  }

  function spaceSyntax() {
    const g = state.graph;
    const spaces = g.nodes.filter((n) => n.cls === "IfcSpace" || n.kind === "Exterior");
    const access = g.links.filter((l) => l.pred === "accessTo");
    const links = access.length ? access : g.links.filter((l) => l.pred === "adjacentTo");
    const G = A.index({ nodes: spaces, links });
    const ss = A.spaceSyntax(G, spaces.some((n) => n.id === "exterior") ? "exterior" : null);
    state.analysis.spaceSyntaxBasis = access.length ? "door access" : "adjacency";
    for (const [id, v] of ss) Object.assign(state.analysis.metrics.get(id) || {}, { integration: v.integration, meanDepth: v.meanDepth, depth: v.depth });
  }

  // ── Colour & size ──────────────────────────────────────────────────────
  function domainColorOf(n) {
    const mn = n.ref ? state.model.nodes.get(n.ref) : null;
    if (n.kind === "System" && mn?.memberDomain) return { key: mn.memberDomain, color: S.DOMAINS[mn.memberDomain].color };
    if (["Element", "Opening", "Port", "Aggregate", "Class"].includes(n.kind) || n.kind === "Spatial" || n.kind === "Project") {
      const d = n.kind === "Project" ? "Spatial" : n.domain || "Other";
      return { key: d, color: (S.DOMAINS[d] || S.DOMAINS.Other).color };
    }
    const k = S.KINDS[n.kind] || S.KINDS.Other;
    return { key: "kind:" + n.kind, color: k.color, label: k.label };
  }

  function computeStyle() {
    const g = state.graph;
    const m = state.model;
    const metrics = state.analysis?.metrics || new Map();
    const st = state.style;
    const items = new Map();
    const bump = (key, color, label, kind) => {
      const it = items.get(key);
      if (it) it.count++;
      else items.set(key, { key, color, label: label || key, count: 1, kind });
    };
    let legend = { type: "cat", title: "", items: [] };

    // group key for clusters layout + categorical colour
    const catBy = (keyFn, colorFn, title) => {
      const freq = new Map();
      for (const n of g.nodes) {
        const k = keyFn(n);
        n._g = k;
        freq.set(k, (freq.get(k) || 0) + 1);
      }
      const order = [...freq.entries()].filter(([k]) => k != null && k !== "—").sort((a, b) => b[1] - a[1]);
      const cmap = new Map(order.map(([k], i) => [k, colorFn ? colorFn(k, i) : i < CAT_PALETTE.length ? CAT_PALETTE[i] : "#a3a3a3"]));
      for (const n of g.nodes) {
        n._c = cmap.get(n._g) || (theme() === "dark" ? "#4b5563" : "#cbd5e1");
        bump(n._g ?? "—", n._c, n._g ?? "—");
      }
      legend.title = title;
    };
    const seqBy = (valFn, title, fmtFn = fmt) => {
      const vals = g.nodes.map((n) => valFn(n)).filter((v) => v != null && Number.isFinite(v));
      const sorted = vals.slice().sort((a, b) => a - b);
      const rank = (v) => {
        if (sorted.length < 2) return 0.5;
        let lo = 0, hi = sorted.length - 1;
        while (lo < hi) {
          const mid = (lo + hi) >> 1;
          if (sorted[mid] < v) lo = mid + 1;
          else hi = mid;
        }
        return lo / (sorted.length - 1);
      };
      for (const n of g.nodes) {
        const v = valFn(n);
        n._c = v == null || !Number.isFinite(v) ? (theme() === "dark" ? "#374151" : "#d1d5db") : seq(rank(v));
        n._g = v == null ? "—" : Math.round(rank(v) * 5);
      }
      legend = { type: "seq", title, min: sorted.length ? fmtFn(sorted[0]) : "–", max: sorted.length ? fmtFn(sorted[sorted.length - 1]) : "–", missing: g.nodes.length - vals.length };
    };

    switch (st.colorBy) {
      case "kind":
        for (const n of g.nodes) {
          const k = S.KINDS[n.kind] || S.KINDS.Other;
          n._c = k.color;
          n._g = n.kind;
          bump(n.kind, k.color, k.label, n.kind);
        }
        legend.title = "Node kind";
        break;
      case "class":
        catBy((n) => n.cls || n.kind, null, "IFC class");
        break;
      case "storey": {
        const N = m.storeys.length;
        for (const n of g.nodes) {
          const sid = n.cls === "IfcBuildingStorey" ? n.ref : n.storey;
          const s = sid ? m.nodes.get(sid) : null;
          n._g = s ? s.label : "—";
          n._c = s ? seq(N > 1 ? s.storeyIndex / (N - 1) : 0.5) : theme() === "dark" ? "#4b5563" : "#cbd5e1";
          bump(n._g, n._c, n._g);
        }
        legend.title = "Storey (bottom → top)";
        legend.sortBy = (a, b) => (m.nodes.get(m.storeys.find((x) => m.nodes.get(x).label === a.key))?.storeyIndex ?? 99) - (m.nodes.get(m.storeys.find((x) => m.nodes.get(x).label === b.key))?.storeyIndex ?? 99);
        break;
      }
      case "system":
        catBy((n) => {
          if (n.kind === "System") return n.label;
          const mn = n.ref ? m.nodes.get(n.ref) : null;
          return mn?.systemName || "—";
        }, null, "Distribution system");
        break;
      case "community":
        catBy((n) => {
          const c = metrics.get(n.id)?.community;
          return c == null ? "—" : "C" + (c + 1);
        }, null, "Louvain community");
        break;
      case "health": {
        const col = { error: "#dc2626", warning: "#f59e0b", info: "#38bdf8", ok: "#16a34a" };
        const lab = { error: "Error", warning: "Warning", info: "Info", ok: "No issues" };
        for (const n of g.nodes) {
          let sev = "ok";
          if (n.members) {
            for (const mid of n.members) {
              const sv = state.healthIndex.get(mid);
              if (sv === "error" || (sv === "warning" && sev !== "error") || (sv === "info" && sev === "ok")) sev = sv;
            }
          } else sev = state.healthIndex.get(n.ref) || "ok";
          n._c = col[sev];
          n._g = sev;
          bump(sev, col[sev], lab[sev]);
        }
        legend.title = "Worst model-health finding";
        break;
      }
      case "degree":
        seqBy((n) => metrics.get(n.id)?.degree, "Degree");
        break;
      case "pagerank":
        seqBy((n) => metrics.get(n.id)?.pagerank, "PageRank");
        break;
      case "betweenness":
        seqBy((n) => metrics.get(n.id)?.betweenness, "Betweenness centrality");
        break;
      case "closeness":
        seqBy((n) => metrics.get(n.id)?.closeness, "Harmonic closeness");
        break;
      case "integration":
        seqBy((n) => metrics.get(n.id)?.integration, "Space-syntax integration");
        break;
      case "property": {
        const key = st.colorProp;
        const val = (n) => {
          const mn = n.ref ? m.nodes.get(n.ref) : null;
          if (!mn || !key) return undefined;
          if (key.startsWith("*.")) {
            const suf = key.slice(1);
            for (const [k, v] of m.propMap(mn)) if (k.endsWith(suf)) return v;
            return undefined;
          }
          return m.propValue(mn, key);
        };
        const vals = g.nodes.map(val).filter((v) => v != null);
        const numeric = vals.length && vals.every((v) => typeof v === "number");
        if (numeric) seqBy((n) => val(n), key || "Property");
        else catBy((n) => {
          const v = val(n);
          return v == null || v === "" ? "—" : String(v);
        }, null, key || "Property");
        break;
      }
      default:
        for (const n of g.nodes) {
          const d = domainColorOf(n);
          n._c = d.color;
          n._g = d.key;
          bump(d.key, d.color, d.label || d.key, d.key.startsWith("kind:") ? n.kind : null);
        }
        legend.title = "Discipline / kind";
    }
    if (legend.type === "cat") {
      legend.items = [...items.values()];
      if (legend.sortBy) legend.items.sort(legend.sortBy);
      else legend.items.sort((a, b) => b.count - a.count);
    }
    state.legend = legend;

    // sizes
    const sb = st.sizeBy;
    let maxV = 0;
    if (sb !== "kind" && sb !== "count") for (const n of g.nodes) maxV = Math.max(maxV, metrics.get(n.id)?.[sb] || 0);
    const degs = g.nodes.map((n) => metrics.get(n.id)?.degree || 0).sort((a, b) => b - a);
    const hubCut = degs[Math.min(degs.length - 1, Math.max(8, Math.floor(degs.length * 0.02)))] || Infinity;
    // only the largest groups get permanent labels; the rest appear on zoom / hover
    const groupCounts = g.nodes.filter((n) => n.kind === "Aggregate" || n.kind === "Class").map((n) => n.count || 0).sort((a, b) => b - a);
    const groupCut = groupCounts.length > 40 ? groupCounts[39] : 0;
    for (const n of g.nodes) {
      let r = n.kind === "Spatial" ? SPATIAL_R[n.cls] || BASE_R.Spatial : BASE_R[n.kind] || 4;
      const grouped = n.kind === "Aggregate" || n.kind === "Class";
      if (n.count && (grouped || sb === "count")) r += 1.25 * Math.log2(n.count + 1);
      if (sb !== "kind" && sb !== "count" && maxV > 0 && !grouped) {
        const v = metrics.get(n.id)?.[sb] || 0;
        r *= 0.75 + 1.6 * Math.sqrt(v / maxV);
      }
      n._r = Math.min(r, 24);
      const deg = metrics.get(n.id)?.degree || 0;
      const grouped2 = n.kind === "Aggregate" || n.kind === "Class";
      n._important = (grouped2 && (n.count || 0) >= groupCut && (n.count || 0) > 0) || ["Project", "System", "Zone", "ClassSystem", "Exterior"].includes(n.kind) || (n.kind === "Spatial" && (n.cls !== "IfcSpace" || state.settings.lens === "spaces")) || deg >= hubCut || n.focus;
    }
    for (const l of g.links) l._c = st.edgeColor ? S.CATEGORY_COLORS[l.cat] || S.CATEGORY_COLORS.Other : null;
  }

  // ── Stats, notes, legend ───────────────────────────────────────────────
  function updateStats() {
    const g = state.graph;
    const a = state.analysis;
    $("viewStats").innerHTML = [
      [g.nodes.length.toLocaleString(), "nodes"],
      [g.links.length.toLocaleString(), "edges"],
      [a.components.toLocaleString(), "components"],
      [a.communities != null ? a.communities.toLocaleString() : "–", "communities"],
      [a.density < 0.001 ? a.density.toExponential(1) : a.density.toFixed(3), "density"],
    ].map(([b, s]) => `<div class="stat"><b>${b}</b><span>${s}</span></div>`).join("");
  }

  function updateNotes() {
    const notes = (state.graph.info?.notes || []).slice();
    if (state.settings.lens === "subgraph") notes.unshift("");
    $("viewNotes").innerHTML = notes
      .map((t, i) =>
        t === "" && i === 0
          ? `<div class="note info">Focused subgraph. <button class="link-btn" data-act="leave-subgraph" style="position:static;font-size:12px;color:var(--accent)">Return to previous lens</button></div>`
          : `<div class="note"><span>${esc(t)}</span><button aria-label="Dismiss" data-act="dismiss">×</button></div>`,
      )
      .join("");
  }

  function updateLegend() {
    const L = state.legend;
    const el = $("legend");
    let h = `<h5><span>${esc(L.title)}</span><button class="link-btn" data-act="legend-toggle">${el.classList.contains("collapsed") ? "show" : "hide"}</button></h5>`;
    if (L.type === "seq") {
      const stops = Array.from({ length: 11 }, (_, i) => seq(i / 10)).join(",");
      h += `<div><div class="grad" style="background:linear-gradient(90deg,${stops})"></div><div class="grad-labels"><span>${esc(L.min)}</span><span>rank-scaled</span><span>${esc(L.max)}</span></div>${L.missing ? `<div class="muted" style="margin-top:3px">${L.missing.toLocaleString()} nodes without value</div>` : ""}</div>`;
    } else {
      const clickable = state.style.colorBy === "domain";
      h += "<div>" + L.items.slice(0, 12).map((it) => {
        const isDomain = clickable && !it.key.startsWith("kind:");
        return `<div class="legend-row${isDomain ? " click" : ""}" ${isDomain ? `data-domain="${esc(it.key)}" title="Click to show only this discipline"` : ""}>${shapeIcon(it.kind || "Element", it.color, 12)}<span>${esc(it.label)}</span><span class="n">${it.count.toLocaleString()}</span></div>`;
      }).join("") + (L.items.length > 12 ? `<div class="muted">+ ${L.items.length - 12} more</div>` : "") + "</div>";
    }
    // shapes
    const kinds = new Map();
    for (const n of state.graph.nodes) kinds.set(n.kind, (kinds.get(n.kind) || 0) + 1);
    if (state.style.colorBy !== "kind" && kinds.size > 1) {
      h += `<h5><span>Shapes</span></h5><div class="legend-shapes">` + [...kinds.entries()].sort((a, b) => b[1] - a[1]).slice(0, 10).map(([k]) => `<span title="${esc((S.KINDS[k] || S.KINDS.Other).label)}">${shapeIcon(k, theme() === "dark" ? "#a9b4c6" : "#64748b", 11)}${esc((S.KINDS[k] || S.KINDS.Other).label.replace(/^Building /, ""))}</span>`).join("") + "</div>";
    }
    // edges
    const cats = new Map();
    let inferred = 0;
    for (const l of state.graph.links) {
      cats.set(l.cat, (cats.get(l.cat) || 0) + 1);
      if (l.inferred) inferred++;
    }
    if (cats.size && state.style.edgeColor) {
      h += `<h5><span>Relationships</span></h5><div>` + [...cats.entries()].sort((a, b) => b[1] - a[1]).map(([c, k]) => `<div class="legend-row"><span class="legend-line" style="border-color:${S.CATEGORY_COLORS[c] || "#999"}"></span><span>${esc(c)}</span><span class="n">${k.toLocaleString()}</span></div>`).join("") + (inferred ? `<div class="legend-row"><span class="legend-line dash" style="border-color:var(--muted)"></span><span>Inferred</span><span class="n">${inferred.toLocaleString()}</span></div>` : "") + "</div>";
    }
    el.innerHTML = h;
  }

  // ── Sidebar: lenses, options, filters ──────────────────────────────────
  function lensCount(id) {
    const m = state.model;
    if (!m) return "";
    const c = m.catalog;
    switch (id) {
      case "kg": return m.nodes.size.toLocaleString();
      case "spatial": return (c.kindCount.get("Spatial") || 0) + " spatial";
      case "breakdown": return "pivot";
      case "systems": return ((c.kindCount.get("System") || 0) + (c.kindCount.get("Group") || 0)) + " sys · " + (c.predCount.get("feeds") || 0) + " flow";
      case "connectivity": return ((c.predCount.get("connectedTo") || 0) + (c.predCount.get("hosts") || 0)).toLocaleString();
      case "spaces": return (c.classCount.get("IfcSpace") || 0) + " spaces";
      case "types": return (c.kindCount.get("Type") || 0) + " types";
      case "materials": return (c.kindCount.get("Material") || 0) + " mat.";
      case "properties": return c.props.size.toLocaleString() + " props";
      case "schema": return c.classCount.size + " classes";
      default: return "";
    }
  }
  function lensEmpty(id) {
    const m = state.model;
    if (!m) return false;
    const c = m.catalog;
    if (id === "systems") return !(c.kindCount.get("System") || c.kindCount.get("Group") || c.predCount.get("feeds"));
    if (id === "spaces") return !c.classCount.get("IfcSpace");
    if (id === "connectivity") return !(c.predCount.get("connectedTo") || c.predCount.get("hosts") || c.predCount.get("coveredBy"));
    if (id === "materials") return !c.kindCount.get("Material");
    return false;
  }

  function renderLensList() {
    const cur = state.settings.lens;
    $("lensList").innerHTML = V.LENSES.filter((l) => !l.hidden || l.id === cur)
      .map((l) => `<button class="lens${l.id === cur ? " active" : ""}${lensEmpty(l.id) ? " disabled" : ""}" data-lens="${l.id}" title="${esc(l.desc)}">
          <span class="lens-icon">${icon(l.id)}</span>
          <span><div class="lens-title">${esc(l.title)}</div>${l.id === cur ? `<div class="lens-desc">${esc(l.desc)}</div>` : ""}</span>
          <span class="lens-count">${esc(lensCount(l.id))}</span></button>`)
      .join("");
  }

  function setLens(id, opts = {}) {
    const s = state.settings;
    if (s.lens !== "subgraph") state.prevLens = s.lens;
    s.lens = id;
    const def = V.LENSES.find((l) => l.id === id);
    if (!opts.keepLayout) state.style.layout = def.layout;
    syncControls();
    renderLensOptions();
    $("kindBox").hidden = id !== "kg";
    rebuild();
    closeSideOnMobile();
  }

  function renderLensOptions() {
    const s = state.settings;
    const opt = (key, label) => `<label class="check"><input type="checkbox" data-opt="${key}" ${s[key] ? "checked" : ""}/> ${label}</label>`;
    let h = "";
    switch (s.lens) {
      case "kg":
        h = `<div class="field"><label for="valueKey">Property values as nodes</label><input type="text" id="valueKey" list="propKeys" placeholder="e.g. Pset_WallCommon.FireRating" value="${esc(s.valueKey)}"/></div><div class="muted" style="font-size:11px">Toggle property sets, ports and openings in <b>Node kinds</b> below.</div>`;
        break;
      case "spatial":
        h = opt("showOpenings", "Show openings") + opt("showPorts", "Show ports") + opt("showReferences", "Show spatial references");
        break;
      case "breakdown": {
        const keys = state.model ? [...BD_ATTRS, ...[...state.model.catalog.props.values()].sort((a, b) => b.count - a.count).slice(0, 300).map((p) => p.key)] : BD_ATTRS;
        h = `<datalist id="bdKeys">${keys.map((k) => `<option value="${esc(k)}">`).join("")}</datalist>` +
          [0, 1, 2].map((i) => `<div class="field"><label>Level ${i + 1}</label><input type="text" list="bdKeys" data-bd="${i}" value="${esc(s.breakdown[i] || "")}" placeholder="${i ? "(optional)" : "e.g. Category"}"/></div>`).join("") +
          opt("includeInstances", "Show individual elements");
        break;
      }
      case "systems":
        h = opt("showPorts", "Show distribution ports");
        break;
      case "connectivity":
        h = opt("showOpenings", "Show openings (voids / fills)");
        break;
      case "spaces":
        h = opt("showZones", "Show zones") + opt("showBoundaries", "Show bounding elements");
        break;
      case "types":
        h = opt("includeInstances", "Include element instances") + opt("includeClassification", "Include classification crosswalk");
        break;
      case "properties":
        h = opt("includeProps", "Include individual properties");
        break;
      case "schema":
        h = opt("includeTaxonomy", "Include class inheritance");
        break;
    }
    $("lensOptions").innerHTML = h ? `<div style="margin-top:4px">${h}</div>` : "";
    $("kindBox").hidden = s.lens !== "kg";
  }

  function renderFilters() {
    const m = state.model;
    const s = state.settings;
    // domains
    const dc = m.catalog.domainCount;
    $("domainChips").innerHTML = [...dc.entries()].sort((a, b) => b[1] - a[1]).map(([d, c]) => {
      const on = !s.domains || s.domains.has(d);
      return `<button class="chip${on ? "" : " off"}" data-domain-chip="${esc(d)}" title="${c.toLocaleString()} elements"><span class="sw" style="background:${(S.DOMAINS[d] || S.DOMAINS.Other).color}"></span>${esc(d)} <small>${c.toLocaleString()}</small></button>`;
    }).join("");
    // kinds
    const kc = m.catalog.kindCount;
    const kindsAll = [...new Set([...kc.keys(), "PropertySet"])].filter((k) => k !== "Other" || kc.get("Other"));
    $("kindList").innerHTML = kindsAll.sort((a, b) => (kc.get(b) || 0) - (kc.get(a) || 0)).map((k) => {
      const meta = S.KINDS[k] || S.KINDS.Other;
      const count = k === "PropertySet" ? new Set([...m.psets.values()].map((p) => p.name)).size : kc.get(k) || 0;
      return `<label class="check"><input type="checkbox" data-kind="${k}" ${s.kinds.has(k) ? "checked" : ""}/> ${shapeIcon(k, meta.color, 12)} ${esc(k === "PropertySet" ? "Property sets (by name)" : meta.label)}<span class="count">${count.toLocaleString()}</span></label>`;
    }).join("");
    // storeys
    $("storeyList").innerHTML = m.storeys.length
      ? m.storeys.slice().reverse().map((id) => {
          const n = m.nodes.get(id);
          const cnt = [...m.nodes.values()].filter((x) => x.storey === id && x.kind === "Element").length;
          return `<label class="check"><input type="checkbox" data-storey="${esc(id)}" ${!s.storeys || s.storeys.has(id) ? "checked" : ""}/> ${esc(n.label)}${n.elevation != null ? ` <span class="muted">${fmt(n.elevation, 2)}</span>` : ""}<span class="count">${cnt.toLocaleString()}</span></label>`;
        }).join("")
      : `<div class="muted" style="padding:4px 0">No storeys</div>`;
    renderClassList();
    // categories
    const cc = m.catalog.catCount;
    $("catList").innerHTML = [...cc.entries()].sort((a, b) => b[1] - a[1]).map(([c, k]) => `<label class="check"><input type="checkbox" data-cat="${esc(c)}" ${!s.cats || s.cats.has(c) ? "checked" : ""}/> <span class="legend-line" style="border-color:${S.CATEGORY_COLORS[c] || "#999"};width:12px;flex:0 0 12px"></span> ${esc(c)}<span class="count">${k.toLocaleString()}</span></label>`).join("");
    updateHiddenRow();
  }

  function renderClassList() {
    const m = state.model;
    const s = state.settings;
    const q = ($("classSearch").value || "").toLowerCase();
    const entries = [...m.catalog.classCount.entries()].filter(([c]) => {
      const k = S.kindOf(c);
      return (k === "Element" || k === "Opening" || k === "Port" || m.source === "csv") && c.toLowerCase().includes(q);
    });
    $("classList").innerHTML = entries.sort((a, b) => b[1] - a[1]).map(([c, k]) => `<label class="check"><input type="checkbox" data-class="${esc(c)}" ${!s.classes || s.classes.has(c) ? "checked" : ""}/> ${esc(short(c))}<span class="count">${k.toLocaleString()}</span></label>`).join("") || `<div class="muted" style="padding:4px 0">No matching classes</div>`;
  }

  function updateHiddenRow() {
    const h = state.settings.hidden.size;
    $("hiddenRow").hidden = !h;
    $("hiddenLabel").textContent = `${h} hidden node${h === 1 ? "" : "s"}`;
  }

  const rebuildSoon = debounce(() => rebuild(), 160);

  function toggleSetFilter(key, value, allValues, on) {
    const s = state.settings;
    let set = s[key] ? new Set(s[key]) : new Set(allValues);
    if (on) set.add(value);
    else set.delete(value);
    s[key] = set.size === allValues.length ? null : set;
  }

  // ── Selection & inspector ──────────────────────────────────────────────
  function select(id, opts = {}) {
    state.selection = id;
    renderer.select(id);
    renderInspector();
    if (id && window.innerWidth <= 1100) $("inspector").classList.add("open");
  }

  function viewNode(id) {
    return state.graph?.nodes.find((n) => n.id === id) || null;
  }

  /** Navigate to a model node: select it if visible, otherwise explore its neighbourhood. */
  function goTo(modelId) {
    const vn = state.graph.nodes.find((n) => n.id === modelId) || state.graph.nodes.find((n) => n.members && n.members.includes(modelId));
    if (vn && vn.id === modelId) {
      select(vn.id);
      renderer.focus(vn.id);
      return;
    }
    exploreNeighborhood(modelId, 1);
  }

  function renderInspector() {
    const body = $("inspectorBody");
    const m = state.model;
    if (!m) {
      body.innerHTML = `<div class="insp-empty">${icon("kg", 40)}<div>Open an IFC or CSV file to build its knowledge graph.</div></div>`;
      return;
    }
    const id = state.selection;
    const vn = id ? viewNode(id) : null;
    const mn = id ? m.nodes.get(vn?.ref || id) : null;
    if (!id || (!vn && !mn)) {
      body.innerHTML = inspectorOverview();
      return;
    }
    if (mn) body.innerHTML = inspectModelNode(mn, vn);
    else body.innerHTML = inspectVirtual(vn);
  }

  function inspectorOverview() {
    const m = state.model;
    const lens = V.LENSES.find((l) => l.id === state.settings.lens);
    const c = m.meta.counts;
    return `<div class="insp-head"><div class="insp-title">${icon(lens.id, 22)}<div><h2>${esc(lens.title)}</h2><div class="sub">${esc(lens.desc)}</div></div></div>
      <div class="insp-actions"><button class="btn sm insp-close icon-btn" data-act="close-insp" style="width:auto">Close</button></div></div>
      <div class="insp-body" style="padding-top:12px">
        <div class="metric-grid">
          <div class="metric"><b>${c.nodes.toLocaleString()}</b><span>KG nodes</span></div>
          <div class="metric"><b>${c.edges.toLocaleString()}</b><span>KG edges</span></div>
          <div class="metric"><b>${c.inferred.toLocaleString()}</b><span>Inferred edges</span></div>
          <div class="metric"><b>${state.health.score}</b><span>Health score / 100</span></div>
        </div>
        <p class="muted" style="margin:14px 0 6px">Click a node to inspect its IFC attributes, property sets, materials, ontology mappings and relationships. Right-click for neighbourhood exploration, path finding and hiding. Double-click expands a node’s neighbourhood.</p>
        <p class="muted" style="margin:0">Keyboard: <span class="mono">/</span> search · <span class="mono">F</span> fit · <span class="mono">Space</span> freeze layout · <span class="mono">Esc</span> clear.</p>
      </div>`;
  }

  function kv(rows) {
    return `<table class="kv">${rows.filter((r) => r && r[1] !== undefined && r[1] !== null && r[1] !== "").map(([k, v, raw]) => `<tr><td>${esc(k)}</td><td>${raw ? v : esc(v)}</td></tr>`).join("")}</table>`;
  }
  const sec = (title, inner, open = true, n = "") => `<details class="insp-sec" ${open ? "open" : ""}><summary>${esc(title)}${n !== "" ? `<span class="n">${n}</span>` : ""}</summary><div class="insp-body">${inner}</div></details>`;

  function nodeColor(id) {
    const vn = state.graph.nodes.find((n) => n.id === id);
    if (vn) return vn._c;
    const mn = state.model.nodes.get(id);
    if (!mn) return "#999";
    return domainColorOf({ ...mn, ref: mn.id }).color;
  }

  function inspectModelNode(n, vn) {
    const m = state.model;
    const color = vn?._c || nodeColor(n.id);
    const kindLabel = (S.KINDS[n.kind] || S.KINDS.Other).label;
    const chips = [];
    if (n.domain && n.domain !== "Semantic") chips.push(`<span class="pill gray" title="${esc(n.domainReason || "")}"><span class="sw" style="width:7px;height:7px;border-radius:50%;background:${(S.DOMAINS[n.domain] || S.DOMAINS.Other).color}"></span>${esc(n.domain)}</span>`);
    if (n.predefinedType) chips.push(`<span class="pill gray">${esc(n.predefinedType)}</span>`);
    if (n.inferred || n.virtual) chips.push(`<span class="pill gray">virtual</span>`);
    const sev = state.healthIndex.get(n.id);
    if (sev) chips.push(`<span class="pill gray" style="color:${sev === "error" ? "var(--err)" : sev === "warning" ? "var(--warn)" : "var(--info)"}">${sev} findings</span>`);
    const inView = state.graph.nodes.some((x) => x.id === n.id);
    let h = `<div class="insp-head">
      <div class="insp-title">${shapeIcon(n.kind, color, 22)}<div style="min-width:0"><h2>${esc(n.label)}</h2><div class="sub">${esc(n.cls || "")} · ${esc(kindLabel)}${n.sub ? " · " + esc(n.sub) : ""}</div></div></div>
      <div class="insp-chips">${chips.join("")}</div>
      <div class="insp-actions">
        ${inView ? `<button class="btn sm" data-act="focus" data-id="${esc(n.id)}">Focus</button>` : ""}
        <button class="btn sm" data-act="explore" data-id="${esc(n.id)}" data-hops="1">Neighbourhood</button>
        <button class="btn sm" data-act="explore" data-id="${esc(n.id)}" data-hops="2">2-hop</button>
        <button class="btn sm" data-act="path-start" data-id="${esc(n.id)}">${state.pathStart && state.pathStart !== n.id ? "Path to here" : "Path from here"}</button>
        ${inView ? `<button class="btn sm" data-act="hide" data-id="${esc(n.id)}">Hide</button>` : ""}
        <button class="btn sm insp-close" data-act="close-insp">Close</button>
      </div></div>`;

    // Identity
    const ident = [
      ["GlobalId", n.guid ? `<span class="mono">${esc(n.guid)}</span> <button class="copy" data-copy="${esc(n.guid)}" title="Copy">⧉</button>` : "", true],
      ["STEP id", n.ifcId != null ? `<span class="mono">#${n.ifcId}</span>` : "", true],
      ["Name", n.name],
      ["Long name", n.longName],
      ["Description", n.description],
      ["Object type", n.objectType],
      ["Tag", n.tag],
      ["Predefined type", n.predefinedType],
      ["Elevation", n.elevation != null ? `${fmt(n.elevation)} ${m.meta.units.LENGTHUNIT || ""}` : ""],
      ["Latitude / longitude", n.latitude != null ? `${fmt(n.latitude, 6)}, ${fmt(n.longitude, 6)}` : ""],
      ["Flow direction", n.flowDirection],
      ["Classification code", n.code],
      ["Scheme", n.scheme],
      ["Category", n.category],
      ["Source category", n.sourceCategory],
      ["Location", n.location],
      ["Edition", n.edition],
    ];
    h += sec("Identity", kv(ident));

    // Location
    const path = m.spatialPath(n.id).reverse();
    if (path.length) {
      h += sec("Spatial location", `<div class="crumbs">${path.map((pid) => `<button class="crumb" data-goto="${esc(pid)}">${esc(m.nodes.get(pid)?.label)}</button>`).join('<span class="sep">›</span>')}<span class="sep">›</span><b>${esc(n.label)}</b></div>`);
    }

    // Semantics
    const chain = n.cls && n.cls.startsWith("Ifc") ? S.chain(n.cls, m.meta.schema) : [];
    const onto = S.ontologyFor(n.cls, n.kind);
    const iri = (q) => {
      const [p, l] = q.split(":");
      return `<a href="${esc((S.NS[p] || "") + l)}" target="_blank" rel="noopener">${esc(q)}</a>`;
    };
    let sem = "";
    if (chain.length) sem += `<div class="field-label" style="margin-bottom:4px">IFC inheritance</div><div class="chain">${chain.map((c) => `<span>${esc(c)}</span>`).join("")}</div>`;
    sem += `<div style="margin-top:10px">${kv([
      ["Discipline", n.domain ? `${esc(n.domain)}<div class="muted" style="font-size:11px">${esc(n.domainReason || "")}</div>` : "", true],
      ["ifcOWL", n.cls?.startsWith("Ifc") ? iri(onto.ifcowl) : "", true],
      ["BOT", onto.bot ? iri(onto.bot) : "", true],
      ["Brick", onto.brick ? iri(onto.brick) : "", true],
      ["Type object", n.typeId ? `<button class="crumb" data-goto="${esc(n.typeId)}">${esc(m.nodes.get(n.typeId)?.label)}</button>` : "", true],
      ["Hosted by", n.hostId ? `<button class="crumb" data-goto="${esc(n.hostId)}">${esc(m.nodes.get(n.hostId)?.label)}</button>` : "", true],
      ["Classification", (n.classRefs || []).map((c) => `<button class="crumb" data-goto="${esc(c)}">${esc(m.nodes.get(c)?.label)}</button>`).join(" "), true],
      ["System", m.outE(n.id, "memberOf").map((e) => `<button class="crumb" data-goto="${esc(e.t)}">${esc(m.nodes.get(e.t)?.label)}</button>`).join(" "), true],
    ])}</div>`;
    h += `<div class="onto">${sec("Semantics & ontology", sem)}</div>`;

    // Materials
    const mats = n.materials?.length ? n.materials : n.materialsInherited || [];
    if (mats.length) {
      const lu = m.meta.units.LENGTHUNIT || "";
      h += sec("Materials", `${!n.materials?.length ? '<div class="muted" style="margin-bottom:4px"><span class="badge t">from type</span></div>' : ""}${kv(mats.map((x, i) => [
        x.layer ? `Layer ${x.layer}${x.setName ? " · " + x.setName : ""}` : x.constituent || x.setName || `Material ${i + 1}`,
        `<button class="crumb" data-goto="${esc(x.id)}">${esc(x.name || m.nodes.get(x.id)?.label)}</button>${x.thickness != null ? ` <span class="muted">${fmt(x.thickness)} ${lu}</span>` : ""}${x.fraction != null ? ` <span class="muted">${pct(x.fraction)}</span>` : ""}${x.profile ? ` <span class="muted">${esc(x.profile)}</span>` : ""}`,
        true,
      ]))}`, true, mats.length);
    }

    // Properties
    const sets = m.propertySets(n);
    if (sets.length) {
      const total = sets.reduce((s, p) => s + p.props.length, 0);
      const inner = sets.map((ps, i) => `<details class="pset" ${i < 3 ? "open" : ""}><summary>${ps.kind === "Q" ? '<span class="badge q">Q</span>' : ""}${ps.inherited ? '<span class="badge t">type</span>' : ""}${esc(ps.name)}<span class="n">${ps.props.length}</span></summary>${kv(ps.props.map((p) => [p.name, `${esc(fmt(p.value))}${p.unit ? `<span class="unit">${esc(p.unit)}</span>` : ""}`, true]))}</details>`).join("");
      h += sec("Properties & quantities", inner, true, `${sets.length} sets · ${total}`);
    }

    // Relationships
    const groups = new Map();
    for (const e of m.outE(n.id)) {
      const k = "→ " + e.pred;
      if (!groups.has(k)) groups.set(k, { pred: e.pred, dir: "→", items: [] });
      groups.get(k).items.push({ id: e.t, e });
    }
    for (const e of m.inE(n.id)) {
      const k = "← " + e.pred;
      if (!groups.has(k)) groups.set(k, { pred: e.pred, dir: "←", items: [] });
      groups.get(k).items.push({ id: e.s, e });
    }
    if (groups.size) {
      let total = 0;
      const inner = [...groups.values()].sort((a, b) => b.items.length - a.items.length).map((g) => {
        total += g.items.length;
        const p = S.PRED[g.pred] || S.PRED.related;
        const label = g.dir === "→" ? p.label : `is “${p.label}” of`;
        const list = g.items.slice(0, 12).map((it) => {
          const t = m.nodes.get(it.id);
          const prov = it.e.inferred ? '<span class="badge inf">inferred</span>' : it.e.attrs?.rel ? `<span class="c">${esc(it.e.attrs.rel.replace(/^IfcRel/, ""))}</span>` : "";
          return `<button class="rel-item" data-goto="${esc(it.id)}" title="${esc(t?.cls || "")}"><span class="sw" style="background:${nodeColor(it.id)}"></span><span class="t">${esc(t?.label || it.id)}</span>${prov}</button>`;
        }).join("");
        return `<div class="rel-group"><h6><span class="dir">${g.dir}</span>${esc(label)} <span class="muted">(${g.items.length})</span>${g.items.length > 12 ? `<button class="link-btn" data-act="explore" data-id="${esc(n.id)}" data-hops="1" style="margin-left:auto">all</button>` : ""}</h6>${list}</div>`;
      }).join("");
      h += sec("Relationships", inner, true, total);
    }

    // Metrics
    const met = vn ? state.analysis.metrics.get(vn.id) : null;
    if (met) {
      const cells = [
        [met.degree, "Degree (view)"],
        [`${met.inDegree} / ${met.outDegree}`, "In / out degree"],
        [fmt(met.pagerank, 5), "PageRank"],
        [met.betweenness != null ? fmt(met.betweenness, 5) : "–", "Betweenness" + (state.analysis.betweennessSampled ? " (est.)" : "")],
        [met.closeness != null ? fmt(met.closeness, 4) : "–", "Harmonic closeness"],
        [fmt(met.clustering, 3), "Clustering coeff."],
        [met.community != null ? "C" + (met.community + 1) : "–", "Community"],
        [met.component, "Component"],
      ];
      if (met.integration !== undefined) cells.push([met.integration != null ? fmt(met.integration, 3) : "–", "Integration (1/RA)"], [met.meanDepth != null ? fmt(met.meanDepth, 2) : "–", "Mean depth"], [met.depth != null ? met.depth : "–", "Depth from exterior"]);
      h += sec("Network metrics (current view)", `<div class="metric-grid">${cells.map(([b, s]) => `<div class="metric"><b>${esc(b)}</b><span>${esc(s)}</span></div>`).join("")}</div>`, false);
    }

    // Health findings
    const findings = state.health.checks.filter((c) => c.affected.includes(n.id));
    if (findings.length) h += sec("Model-health findings", findings.map((c) => `<div style="display:flex;gap:8px;align-items:flex-start;margin-bottom:6px"><span class="sev ${c.severity}" style="margin-top:4px;flex:0 0 9px"></span><div><b>${esc(c.title)}</b><div class="muted" style="font-size:11.5px">${esc(c.description)}</div></div></div>`).join(""), true, findings.length);

    // Attributes & raw
    const attrs = m.attributes(n);
    if (attrs.length) h += sec("IFC attributes", kv(attrs.map((a) => [a.name, a.value])), false, attrs.length);
    const raw = m.raw(n);
    if (raw) h += sec(m.source === "csv" ? "Source row" : "STEP source", `<pre class="step">${esc(raw.length > 6000 ? raw.slice(0, 6000) + "…" : raw)}</pre>`, false);
    return h;
  }

  function inspectVirtual(vn) {
    const m = state.model;
    let h = `<div class="insp-head"><div class="insp-title">${shapeIcon(vn.kind, vn._c, 22)}<div style="min-width:0"><h2>${esc(vn.label)}</h2><div class="sub">${esc((S.KINDS[vn.kind] || S.KINDS.Other).label)}${vn.count ? " · " + vn.count.toLocaleString() + " members" : ""}</div></div></div>
      <div class="insp-actions">${vn.members ? `<button class="btn sm" data-act="expand-members" data-id="${esc(vn.id)}">Expand members</button>` : ""}<button class="btn sm insp-close" data-act="close-insp">Close</button></div></div>`;
    if (vn.kind === "Class") {
      const cls = vn.cls;
      const onto = S.ontologyFor(cls, S.kindOf(cls));
      h += sec("Class", `<div class="chain">${S.chain(cls, m.meta.schema).map((c) => `<span>${esc(c)}</span>`).join("")}</div><div style="margin-top:8px">${kv([["Instances", (vn.instances ?? vn.count ?? 0).toLocaleString()], ["Abstract (no direct instances)", vn.abstract ? "yes" : ""], ["ifcOWL", onto.ifcowl], ["BOT", onto.bot], ["Brick", onto.brick]])}</div>`);
      const props = [...m.catalog.props.values()].filter((p) => p.classes.has(cls)).sort((a, b) => b.classes.get(cls) - a.classes.get(cls)).slice(0, 25);
      const total = m.catalog.classCount.get(cls) || 1;
      if (props.length) h += sec("Most common properties", kv(props.map((p) => [p.key, pct(p.classes.get(cls) / total)])), true, props.length);
      if (!vn.abstract) h += `<div class="insp-body" style="padding-top:10px"><button class="btn sm" data-act="filter-class" data-cls="${esc(cls)}">Show only ${esc(short(cls))} elements</button></div>`;
    } else if (vn.members) {
      if (vn.facet) h += sec("Group", kv([["Breakdown level", vn.facet], ["Value", String(vn.facetValue)], ["Elements", vn.count.toLocaleString()], ["Dominant class", vn.cls], ["Dominant discipline", vn.domain]]));
      const list = vn.members.slice(0, 60).map((id) => `<button class="rel-item" data-goto="${esc(id)}"><span class="sw" style="background:${vn._c}"></span><span class="t">${esc(m.nodes.get(id)?.label || id)}</span></button>`).join("");
      h += sec("Members", list + (vn.members.length > 60 ? `<div class="muted">+ ${vn.members.length - 60} more</div>` : ""), true, vn.members.length);
    } else if (vn.psetName) {
      const props = [...m.catalog.props.values()].filter((p) => p.pset === vn.psetName);
      h += sec("Properties in this set", kv(props.map((p) => [p.name, `${p.count.toLocaleString()} values · ${pct(p.coverage)} coverage`])), true, props.length);
    } else if (vn.propKey) {
      const p = m.catalog.props.get(vn.propKey);
      if (p) h += sec("Property", kv([["Key", p.key], ["Occurrences", p.count.toLocaleString()], ["Coverage", pct(p.coverage)], ["Distinct values", p.distinct], ["Data types", [...p.datatypes].join(", ")], ["Unit", [...p.units].join(", ")], ["Value", vn.value != null ? fmt(vn.value) : ""]]));
    }
    return h;
  }

  // ── Exploration: neighbourhood, path, subgraphs ────────────────────────
  function exploreNeighborhood(id, hops = 1, cap = 600) {
    const m = state.model;
    if (!m.nodes.has(id)) return;
    const ids = new Set([id]);
    let frontier = [id];
    for (let h = 0; h < hops && ids.size < cap; h++) {
      const next = [];
      for (const v of frontier) {
        for (const w of m.neighbors(v)) {
          if (ids.size >= cap) break;
          if (!ids.has(w)) {
            ids.add(w);
            next.push(w);
          }
        }
      }
      frontier = next;
    }
    const n = m.nodes.get(id);
    showSubgraph(ids, `Neighbourhood of “${n.label}” (${hops} hop${hops > 1 ? "s" : ""}, ${ids.size} nodes${ids.size >= cap ? ", capped" : ""}).`, new Set([id]));
    state.selection = id;
    renderer.select(id);
    renderInspector();
    setTimeout(() => renderer.focus(id), 900);
  }

  function showSubgraph(ids, title, focus, pathEdges) {
    const s = state.settings;
    if (s.lens !== "subgraph") state.prevLens = s.lens;
    s.subgraph = { ids, title, focus: focus || new Set(), pathEdges };
    s.lens = "subgraph";
    state.style.layout = "force";
    syncControls();
    renderLensOptions();
    rebuild();
  }

  function modelPath(a, b) {
    const m = state.model;
    const prev = new Map([[a, null]]);
    const prevEdge = new Map();
    const q = [a];
    while (q.length) {
      const v = q.shift();
      if (v === b) break;
      for (const e of m.out.get(v) || []) if (!prev.has(e.t)) { prev.set(e.t, v); prevEdge.set(e.t, e.id); q.push(e.t); }
      for (const e of m.in.get(v) || []) if (!prev.has(e.s)) { prev.set(e.s, v); prevEdge.set(e.s, e.id); q.push(e.s); }
    }
    if (!prev.has(b)) return null;
    const path = [];
    const edges = new Set();
    for (let v = b; v != null; v = prev.get(v)) {
      path.push(v);
      if (prevEdge.has(v)) edges.add(prevEdge.get(v));
    }
    return { path: path.reverse(), edges };
  }

  function pathAction(id) {
    if (!state.pathStart || state.pathStart === id) {
      state.pathStart = id;
      toast(`Path start: “${state.model.nodes.get(id)?.label}”. Now choose “Path to here” on another node.`);
      renderInspector();
      return;
    }
    const res = modelPath(state.pathStart, id);
    const a = state.model.nodes.get(state.pathStart)?.label;
    const b = state.model.nodes.get(id)?.label;
    if (!res) {
      toast(`No path between “${a}” and “${b}”.`);
      state.pathStart = null;
      return;
    }
    showSubgraph(new Set(res.path), `Shortest path “${a}” → “${b}”: ${res.path.length - 1} hops through ${res.path.length - 2} intermediate nodes.`, new Set([state.pathStart, id]), res.edges);
    state.pathStart = null;
  }

  // ── Search ─────────────────────────────────────────────────────────────
  function buildSearchIndex() {
    state.searchIndex = [];
    for (const n of state.model.nodes.values()) {
      state.searchIndex.push({ id: n.id, t: [n.label, n.name, n.guid, n.cls, n.tag, n.objectType, n.longName, n.code, n.ifcId != null ? "#" + n.ifcId : ""].filter(Boolean).join(" ").toLowerCase(), n });
    }
  }
  let searchSel = -1;
  function runSearch(q) {
    const box = $("searchResults");
    q = q.trim().toLowerCase();
    if (!q) { box.classList.remove("open"); return; }
    const toks = q.split(/\s+/);
    const res = [];
    for (const it of state.searchIndex) {
      if (!toks.every((t) => it.t.includes(t))) continue;
      const lab = (it.n.label || "").toLowerCase();
      const score = (lab === q ? 0 : lab.startsWith(q) ? 1 : it.n.guid === q ? 0 : 2) + (it.n.kind === "Element" ? 0 : 0.5);
      res.push([score, it]);
      if (res.length > 2000) break;
    }
    res.sort((a, b) => a[0] - b[0]);
    searchSel = -1;
    box.innerHTML = res.length
      ? res.slice(0, 40).map(([, it], i) => {
          const n = it.n;
          const st = n.storey ? state.model.nodes.get(n.storey)?.label : "";
          return `<div class="sr" data-goto="${esc(n.id)}" data-i="${i}"><span class="sw" style="background:${nodeColor(n.id)}"></span><div style="min-width:0"><div class="sr-title">${esc(n.label)}</div><div class="sr-sub">${esc(n.cls)}${st ? " · " + esc(st) : ""}${n.guid ? " · " + esc(n.guid) : ""}</div></div><span class="sr-tag">${esc(n.kind)}</span></div>`;
        }).join("") + (res.length > 40 ? `<div class="sr" style="cursor:default"><span></span><span class="sr-sub">${res.length > 2000 ? "2000+" : res.length} matches — <button class="link-btn" data-act="search-subgraph">show all as subgraph</button></span></div>` : "")
      : `<div class="sr" style="cursor:default"><span></span><span class="sr-sub">No matches</span></div>`;
    box.classList.add("open");
    state.lastSearch = res.map(([, it]) => it.id);
  }

  // ── Tooltip ────────────────────────────────────────────────────────────
  function showTip(n) {
    const tip = $("tip");
    if (!n) { tip.classList.remove("show"); return; }
    const met = state.analysis.metrics.get(n.id);
    const st = n.storey ? state.model.nodes.get(n.storey)?.label : "";
    tip.innerHTML = `<b>${esc(n.label)}</b><div class="t-sub">${esc(n.cls || (S.KINDS[n.kind] || {}).label || n.kind)}${st ? " · " + esc(st) : ""}</div>${n.count ? `<div class="t-sub">${n.count.toLocaleString()} members</div>` : ""}${met ? `<div class="t-sub">degree ${met.degree}${met.community != null ? " · C" + (met.community + 1) : ""}</div>` : ""}`;
    placeTip();
    tip.classList.add("show");
  }
  function showLinkTip(l) {
    const tip = $("tip");
    if (!l) { tip.classList.remove("show"); return; }
    const p = S.PRED[l.pred] || S.PRED.related;
    const s = typeof l.source === "object" ? l.source.label : l.source;
    const t = typeof l.target === "object" ? l.target.label : l.target;
    const e = l.edgeId != null ? state.model.edges[l.edgeId] : null;
    tip.innerHTML = `<b>${esc(p.label)}</b><div class="t-sub">${esc(s)} → ${esc(t)}</div><div class="t-sub">${esc(l.cat)}${l.inferred ? " · inferred" : ""}${l.weight > 1 ? " · weight " + l.weight : ""}${e?.attrs?.rel ? " · " + esc(e.attrs.rel) : ""}${e?.attrs?.rule ? " · " + esc(e.attrs.rule) : ""}</div>`;
    placeTip();
    tip.classList.add("show");
  }
  function placeTip() {
    const tip = $("tip");
    const w = tip.offsetWidth || 200;
    tip.style.left = Math.min(window.innerWidth - w - 10, mouse.x + 14) + "px";
    tip.style.top = mouse.y + 14 + "px";
  }

  // ── Context menu ───────────────────────────────────────────────────────
  function openCtx(n, ev) {
    const menu = $("ctxMenu");
    const ref = n.ref;
    const items = [];
    items.push(["inspect", "Inspect"]);
    items.push(["focus", "Focus"]);
    if (ref) {
      items.push(["explore1", "Explore neighbourhood (1 hop)"]);
      items.push(["explore2", "Explore neighbourhood (2 hops)"]);
      items.push(["path", state.pathStart && state.pathStart !== ref ? "Shortest path to here" : "Set as path start"]);
    }
    if (n.members) items.push(["expand", `Expand ${n.count} members`]);
    items.push(["pin", n._pinned ? "Unpin" : "Pin position"]);
    items.push(["hide", "Hide node"]);
    if (n.cls && (n.kind === "Element" || n.kind === "Aggregate")) items.push(["only-class", `Show only ${short(n.cls)}`]);
    if (ref && state.model.nodes.get(ref)?.guid) items.push(["copy", "Copy GlobalId"]);
    menu.innerHTML = `<div class="ctx-head">${esc(n.label)}<small>${esc(n.cls || n.kind)}</small></div>` + items.map(([a, l]) => `<button class="menu-item" data-ctx="${a}">${esc(l)}</button>`).join("");
    menu.dataset.id = n.id;
    menu.classList.add("open");
    const x = Math.min(window.innerWidth - 240, ev.clientX);
    const y = Math.min(window.innerHeight - menu.offsetHeight - 10, ev.clientY);
    menu.style.left = x + "px";
    menu.style.top = y + "px";
  }
  const hideCtx = () => $("ctxMenu").classList.remove("open");

  function ctxAction(a, id) {
    const n = viewNode(id);
    if (!n) return;
    hideCtx();
    switch (a) {
      case "inspect": select(id); break;
      case "focus": select(id); renderer.focus(id); break;
      case "explore1": exploreNeighborhood(n.ref, 1); break;
      case "explore2": exploreNeighborhood(n.ref, 2); break;
      case "path": pathAction(n.ref); break;
      case "expand": showSubgraph(new Set(n.members), `Members of “${n.label}”.`); break;
      case "pin":
        if (n._pinned) { n._pinned = false; n.fx = undefined; n.fy = undefined; }
        else { n._pinned = true; n.fx = n.x; n.fy = n.y; }
        renderer.redraw();
        break;
      case "hide":
        state.settings.hidden.add(n.ref || id);
        updateHiddenRow();
        rebuild({ fit: false });
        break;
      case "only-class":
        state.settings.classes = new Set([n.cls]);
        renderFilters();
        rebuild();
        break;
      case "copy": copy(state.model.nodes.get(n.ref).guid); break;
    }
  }

  function copy(text) {
    navigator.clipboard?.writeText(text).then(() => toast("Copied " + text), () => toast(text));
  }

  // ── Toast ──────────────────────────────────────────────────────────────
  let toastT;
  function toast(msg, ms = 3200) {
    const t = $("toast");
    t.textContent = msg;
    t.classList.add("show");
    clearTimeout(toastT);
    toastT = setTimeout(() => t.classList.remove("show"), ms);
  }

  // ── Drawer ─────────────────────────────────────────────────────────────
  const TABS = [
    ["overview", "Overview"],
    ["analytics", "Network analytics"],
    ["health", "Model health"],
    ["mapping", "Semantic mapping"],
    ["catalog", "Property catalogue"],
    ["query", "Query"],
    ["table", "Node table"],
  ];
  function renderDrawerTabs() {
    const m = state.model;
    $("drawerTabs").innerHTML = TABS.map(([id, label]) => {
      let c = "";
      if (m && id === "health") c = state.health.score;
      if (m && id === "catalog") c = m.catalog.props.size.toLocaleString();
      if (m && id === "query" && state.queryResults) c = state.queryResults.length.toLocaleString();
      return `<button class="tab${state.drawerTab === id ? " active" : ""}" data-tab="${id}" ${!m ? "disabled" : ""}>${label}${c !== "" ? `<span class="tc">${c}</span>` : ""}</button>`;
    }).join("") + `<span style="flex:1"></span><button class="icon-btn" data-act="drawer-toggle" title="Collapse / expand"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><path d="${$("drawer").classList.contains("collapsed") ? "M6 15l6-6 6 6" : "M6 9l6 6 6-6"}"/></svg></button>`;
  }
  function openDrawer(tab) {
    if (tab) state.drawerTab = tab;
    $("drawer").classList.remove("collapsed");
    renderDrawerTabs();
    renderDrawer();
    setTimeout(() => renderer.resize(), 50);
  }
  function toggleDrawer() {
    $("drawer").classList.toggle("collapsed");
    renderDrawerTabs();
    if (!$("drawer").classList.contains("collapsed")) renderDrawer();
    setTimeout(() => renderer.resize(), 50);
  }

  function renderDrawer() {
    if (!state.model) return;
    const el = $("drawerBody");
    switch (state.drawerTab) {
      case "overview": el.innerHTML = tabOverview(); break;
      case "analytics": el.innerHTML = tabAnalytics(); break;
      case "health": el.innerHTML = tabHealth(); break;
      case "mapping": el.innerHTML = tabMapping(); break;
      case "catalog": el.innerHTML = tabCatalog(); break;
      case "query": el.innerHTML = tabQuery(); break;
      case "table": el.innerHTML = tabTable(); break;
    }
  }

  function hbars(rows, opts = {}) {
    const max = Math.max(1, ...rows.map((r) => r.value));
    return rows.map((r) => `<div class="hbar${r.act ? " click" : ""}" ${r.act ? `data-act="${r.act}" data-v="${esc(r.v)}"` : ""} title="${esc(r.title || r.label)}"><span class="hb-l">${r.icon || ""}${esc(r.label)}</span><span class="hb-t"><div style="width:${(100 * r.value) / max}%;background:${r.color || "var(--accent)"}"></div></span><span class="hb-n">${opts.pct ? pct(r.value) : r.value.toLocaleString()}</span></div>`).join("");
  }

  function tabOverview() {
    const m = state.model;
    const c = m.meta.counts;
    const h = m.meta.header || {};
    const kpis = [
      [(m.meta.entityCount || 0).toLocaleString(), m.source === "csv" ? "Source rows" : "STEP instances"],
      [c.nodes.toLocaleString(), "KG nodes"],
      [c.edges.toLocaleString(), "KG edges", `${c.inferred.toLocaleString()} inferred`],
      [c.elements.toLocaleString(), "Elements"],
      [c.spatial.toLocaleString(), "Spatial"],
      [c.types.toLocaleString(), "Types"],
      [c.psets.toLocaleString(), "Property sets"],
      [c.properties.toLocaleString(), "Distinct properties"],
      [m.catalog.classCount.size.toLocaleString(), "Classes"],
      [state.health.score + "/100", "Health score"],
    ];
    const units = Object.entries(m.meta.units || {}).map(([k, v]) => `${k.replace(/UNIT$/, "").toLowerCase()}: ${v}`).join(" · ");
    const source = kv([
      ["File", m.meta.fileName],
      ["Size", bytes(m.meta.fileSize)],
      ["Schema", m.meta.schema],
      ["Model view definition", h.mvd],
      ["Authoring application", m.meta.application || h.originatingSystem],
      ["Pre-processor", h.preprocessor],
      ["Author / organisation", [h.author, h.organization].filter(Boolean).join(" · ")],
      ["Time stamp", h.timeStamp || m.meta.created],
      ["Units", units],
      ["Parse time", (m.meta.loadMs / 1000).toFixed(2) + " s"],
      ["Unresolved references", m.meta.unresolved ? m.meta.unresolved.toLocaleString() : ""],
      ["Notes", (m.meta.warnings || []).join(" ")],
    ]);
    const dom = [...m.catalog.domainCount.entries()].sort((a, b) => b[1] - a[1]).map(([d, v]) => ({ label: d, value: v, color: (S.DOMAINS[d] || S.DOMAINS.Other).color, act: "only-domain", v: d }));
    const cls = [...m.catalog.classCount.entries()].filter(([k]) => S.kindOf(k) === "Element" || m.source === "csv").sort((a, b) => b[1] - a[1]).slice(0, 14).map(([k, v]) => ({ label: short(k), value: v, color: (S.DOMAINS[[...m.nodes.values()].find((n) => n.cls === k)?.domain] || S.DOMAINS.Other).color, act: "only-class", v: k }));
    const preds = [...m.catalog.predCount.entries()].sort((a, b) => b[1] - a[1]).slice(0, 14).map(([p, v]) => ({ label: (S.PRED[p] || S.PRED.related).label + (S.PRED[p]?.inferred ? " (inferred)" : ""), value: v, color: S.CATEGORY_COLORS[(S.PRED[p] || S.PRED.related).cat] }));
    const storeyCounts = new Map();
    for (const n of m.nodes.values()) if (n.kind === "Element" && n.storey) storeyCounts.set(n.storey, (storeyCounts.get(n.storey) || 0) + 1);
    const storeys = m.storeys.slice().reverse().map((id) => ({ label: m.nodes.get(id).label, value: storeyCounts.get(id) || 0, color: "#8b5cf6", act: "only-storey", v: id }));
    const geo = (m.meta.typeHistogram || []).slice(0, 12).map(([t, v]) => ({ label: S.canonical(t), value: v, color: "#94a3b8" }));
    return `<div class="kpis">${kpis.map(([b, s, e]) => `<div class="kpi"><b>${b}</b><span>${s}</span>${e ? `<em>${e}</em>` : ""}</div>`).join("")}</div>
      <div class="grid-3">
        <div class="card"><h4>Source & provenance</h4>${source}</div>
        <div class="card"><h4>Disciplines <small>click to filter</small></h4>${hbars(dom)}</div>
        <div class="card"><h4>Top IFC classes <small>click to filter</small></h4>${hbars(cls)}</div>
        <div class="card"><h4>Relationships by predicate</h4>${hbars(preds)}</div>
        ${storeys.length ? `<div class="card"><h4>Elements per storey <small>click to filter</small></h4>${hbars(storeys)}</div>` : ""}
        <div class="card"><h4>${m.source === "csv" ? "Classes" : "STEP entity histogram"} <small>incl. geometry</small></h4>${hbars(geo)}</div>
      </div>`;
  }

  function degreeChart(hist) {
    if (!hist?.length) return "";
    const W = 320, H = 170, P = 30;
    const pts = hist.filter(([d]) => d > 0);
    if (!pts.length) return "";
    const maxD = Math.max(...pts.map((p) => p[0]));
    const maxC = Math.max(...pts.map((p) => p[1]));
    const x = (d) => P + (Math.log10(d) / Math.log10(Math.max(maxD, 10))) * (W - P - 8);
    const y = (c) => H - P + 6 - (Math.log10(c) / Math.log10(Math.max(maxC, 10))) * (H - P - 4);
    const dots = pts.map(([d, c]) => `<circle cx="${x(d).toFixed(1)}" cy="${y(c).toFixed(1)}" r="2.6" fill="var(--accent)" fill-opacity="0.75"><title>degree ${d}: ${c} nodes</title></circle>`).join("");
    return `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:420px" font-size="9" fill="var(--muted)">
      <line x1="${P}" y1="${H - P + 6}" x2="${W - 4}" y2="${H - P + 6}" stroke="var(--border-strong)"/><line x1="${P}" y1="4" x2="${P}" y2="${H - P + 6}" stroke="var(--border-strong)"/>
      <text x="${W / 2}" y="${H - 4}" text-anchor="middle">degree (log)</text><text x="8" y="${H / 2}" transform="rotate(-90 8 ${H / 2})" text-anchor="middle">nodes (log)</text>
      <text x="${P}" y="${H - P + 16}" text-anchor="middle">1</text><text x="${W - 8}" y="${H - P + 16}" text-anchor="end">${maxD}</text><text x="${P - 3}" y="10" text-anchor="end">${maxC}</text>
      ${dots}</svg>`;
  }

  function topTable(key, label, n = 10) {
    const a = state.analysis;
    const rows = [...a.metrics.entries()].filter(([, v]) => v[key] != null).sort((x, y) => y[1][key] - x[1][key]).slice(0, n);
    return `<table class="data"><thead><tr><th>#</th><th>Node</th><th>Class</th><th style="text-align:right">${esc(label)}</th></tr></thead><tbody>${rows.map(([id, v], i) => {
      const vn = viewNode(id);
      return `<tr class="click" data-select="${esc(id)}"><td>${i + 1}</td><td>${esc(vn?.label)}</td><td class="muted">${esc(short(vn?.cls))}</td><td class="num">${fmt(v[key], 4)}</td></tr>`;
    }).join("")}</tbody></table>`;
  }

  function tabAnalytics() {
    const a = state.analysis;
    const g = state.graph;
    const kpis = [
      [a.n.toLocaleString(), "Nodes"],
      [a.m.toLocaleString(), "Edges"],
      [a.density < 0.001 ? a.density.toExponential(2) : a.density.toFixed(4), "Density"],
      [a.avgDegree.toFixed(2), "Avg degree"],
      [a.components.toLocaleString(), "Components"],
      [a.n ? pct(a.largestComponent / a.n) : "–", "Largest component"],
      [a.isolates.toLocaleString(), "Isolates"],
      [a.clustering.toFixed(3), "Avg clustering"],
      [a.diameter, "Diameter (est.)"],
      [a.communities ?? "–", "Communities"],
      [a.modularity != null ? a.modularity.toFixed(3) : "–", "Modularity Q"],
      [Math.round(a.ms) + " ms", "Compute time"],
    ];
    // communities
    const comm = new Map();
    for (const [id, v] of a.metrics) {
      if (v.community == null) continue;
      let c = comm.get(v.community);
      if (!c) comm.set(v.community, (c = { id: v.community, members: [], classes: new Map(), domains: new Map() }));
      c.members.push(id);
      const vn = viewNode(id);
      if (vn) {
        c.classes.set(vn.cls || vn.kind, (c.classes.get(vn.cls || vn.kind) || 0) + 1);
        if (vn.domain) c.domains.set(vn.domain, (c.domains.get(vn.domain) || 0) + 1);
      }
    }
    const top = (mp) => [...mp.entries()].sort((x, y) => y[1] - x[1])[0]?.[0] || "";
    const commRows = [...comm.values()].sort((x, y) => y.members.length - x.members.length).slice(0, 12).map((c) => `<tr class="click" data-act="community" data-v="${c.id}"><td>C${c.id + 1}</td><td class="num">${c.members.length}</td><td>${esc(short(top(c.classes)))}</td><td class="muted">${esc(top(c.domains))}</td></tr>`).join("");
    let ss = "";
    if (state.settings.lens === "spaces") {
      const rows = g.nodes.filter((n) => n.cls === "IfcSpace").map((n) => [n, a.metrics.get(n.id)]).filter(([, v]) => v).sort((x, y) => (y[1].integration ?? -1) - (x[1].integration ?? -1));
      ss = `<div class="card"><h4>Space syntax <small>basis: ${esc(a.spaceSyntaxBasis || "–")} · RA = 2(MD−1)/(k−2)</small></h4><div class="table-wrap" style="max-height:220px"><table class="data"><thead><tr><th>Space</th><th style="text-align:right">Integration</th><th style="text-align:right">Mean depth</th><th style="text-align:right">Depth from exterior</th><th style="text-align:right">Degree</th></tr></thead><tbody>${rows.map(([n, v]) => `<tr class="click" data-select="${esc(n.id)}"><td>${esc(n.label)}</td><td class="num">${v.integration != null ? fmt(v.integration, 3) : "–"}</td><td class="num">${v.meanDepth != null ? fmt(v.meanDepth, 2) : "–"}</td><td class="num">${v.depth ?? "–"}</td><td class="num">${v.degree}</td></tr>`).join("")}</tbody></table></div></div>`;
    }
    return `<div class="kpis">${kpis.map(([b, s]) => `<div class="kpi"><b>${b}</b><span>${s}</span></div>`).join("")}</div>
      ${a.betweennessSampled || a.closenessSampled ? `<p class="muted" style="margin:-6px 0 10px">Betweenness / closeness are estimated from a random sample of source nodes because the view is large.</p>` : ""}
      <div class="grid-3">
        ${ss}
        <div class="card"><h4>Degree distribution <small>log–log</small></h4>${degreeChart(a.degreeHistogram)}</div>
        <div class="card"><h4>Top PageRank <button class="btn sm" data-act="color" data-v="pagerank">Colour</button></h4>${topTable("pagerank", "PageRank")}</div>
        <div class="card"><h4>Top betweenness (bridges) <button class="btn sm" data-act="color" data-v="betweenness">Colour</button></h4>${topTable("betweenness", "Betweenness")}</div>
        <div class="card"><h4>Top degree (hubs) <button class="btn sm" data-act="color" data-v="degree">Colour</button></h4>${topTable("degree", "Degree")}</div>
        <div class="card"><h4>Louvain communities <button class="btn sm" data-act="color" data-v="community">Colour</button></h4><table class="data"><thead><tr><th>ID</th><th style="text-align:right">Size</th><th>Dominant class</th><th>Discipline</th></tr></thead><tbody>${commRows}</tbody></table></div>
      </div>`;
  }

  function tabHealth() {
    const H = state.health;
    const by = { error: 0, warning: 0, info: 0 };
    for (const c of H.checks) if (c.failed) by[c.severity]++;
    const ring = (score) => {
      const r = 26, C = 2 * Math.PI * r;
      const col = score >= 80 ? "var(--ok)" : score >= 55 ? "var(--warn)" : "var(--err)";
      return `<svg class="score-ring" viewBox="0 0 64 64"><circle cx="32" cy="32" r="${r}" fill="none" stroke="var(--border)" stroke-width="6"/><circle cx="32" cy="32" r="${r}" fill="none" stroke="${col}" stroke-width="6" stroke-dasharray="${(C * score) / 100} ${C}" transform="rotate(-90 32 32)" stroke-linecap="round"/><text x="32" y="37" text-anchor="middle" font-size="16" font-weight="700" fill="var(--text)">${score}</text></svg>`;
    };
    const groups = new Map();
    for (const c of H.checks) {
      if (!groups.has(c.group)) groups.set(c.group, []);
      groups.get(c.group).push(c);
    }
    const sevW = { error: 3, warning: 2, info: 1 };
    const badness = (c) => (c.population ? c.failed / c.population : 0) * sevW[c.severity] + (c.failed ? sevW[c.severity] : 0);
    for (const list of groups.values()) list.sort((a, b) => badness(b) - badness(a));
    const rows = [...groups.entries()].sort((a, b) => Math.max(...b[1].map(badness)) - Math.max(...a[1].map(badness))).map(([g, list]) => `<div class="card" style="margin-bottom:12px"><h4>${esc(g)}</h4>${list.map((c) => {
      const ratio = c.population ? c.failed / c.population : 0;
      const col = c.failed === 0 ? "var(--ok)" : c.severity === "error" ? "var(--err)" : c.severity === "warning" ? "var(--warn)" : "var(--info)";
      return `<div class="check-row"><span class="sev ${c.failed ? c.severity : ""}" style="${c.failed ? "" : "background:var(--ok)"}"></span>
        <div><div class="check-title">${esc(c.title)}</div><div class="check-desc">${esc(c.description)}</div></div>
        <div class="hb-t" title="${pct(ratio)} of ${c.population}"><div style="width:${Math.max(ratio * 100, c.failed ? 2 : 0)}%;background:${col}"></div></div>
        <div class="hb-n">${c.failed.toLocaleString()} / ${c.population.toLocaleString()}</div>
        <div style="display:flex;gap:4px">${c.affected.length ? `<button class="btn sm" data-act="health-show" data-v="${c.id}">Graph</button><button class="btn sm" data-act="health-mark" data-v="${c.id}">Mark</button>` : ""}</div></div>`;
    }).join("")}</div>`).join("");
    return `<div class="toolbar"><div class="score">${ring(H.score)}<div><b style="font-size:14px">Information quality score</b><div class="muted">${by.error} error, ${by.warning} warning and ${by.info} info checks with findings · weighted by severity and share of affected objects</div></div></div><span class="grow"></span>
      <button class="btn sm" data-act="color" data-v="health">Colour graph by health</button><button class="btn sm" data-act="health-clear">Clear marks</button><button class="btn sm" data-act="export" data-v="health">Export CSV</button></div>${rows}`;
  }

  function mappingRows() {
    const m = state.model;
    const byCls = new Map();
    for (const n of m.nodes.values()) {
      if (!n.cls) continue;
      let r = byCls.get(n.cls);
      if (!r) byCls.set(n.cls, (r = { cls: n.cls, kind: n.kind, count: 0, domains: new Map(), classified: 0, typed: 0, props: 0, withMat: 0 }));
      r.count++;
      if (n.domain) r.domains.set(n.domain, (r.domains.get(n.domain) || 0) + 1);
      if (n.classRefs?.length) r.classified++;
      if (n.typeId) r.typed++;
      if (n.materials?.length || n.materialsInherited?.length) r.withMat++;
      if (n.kind === "Element" || n.kind === "Spatial") r.props += m.propertySets(n).reduce((s, p) => s + p.props.length, 0);
    }
    return [...byCls.values()].sort((a, b) => b.count - a.count).map((r) => {
      const o = S.ontologyFor(r.cls, r.kind);
      const dom = [...r.domains.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || "";
      return { ifcClass: r.cls, kind: r.kind, count: r.count, discipline: dom, ifcOWL: r.cls.startsWith("Ifc") ? o.ifcowl : "", bot: o.bot || "", brick: o.brick || "", classified: r.classified / r.count, typed: r.typed / r.count, withMaterial: r.withMat / r.count, avgProperties: r.props / r.count };
    });
  }

  function tabMapping() {
    const m = state.model;
    const rows = mappingRows();
    const physical = rows.filter((r) => ["Element", "Spatial", "System", "Zone", "Opening", "Port", "Type"].includes(r.kind));
    const botCov = physical.filter((r) => r.bot).reduce((s, r) => s + r.count, 0) / Math.max(1, physical.reduce((s, r) => s + r.count, 0));
    const brickCov = physical.filter((r) => r.brick).reduce((s, r) => s + r.count, 0) / Math.max(1, physical.reduce((s, r) => s + r.count, 0));
    const systems = [...m.nodes.values()].filter((n) => n.kind === "ClassSystem");
    const sysRows = systems.map((s) => {
      const refs = [...m.nodes.values()].filter((n) => n.kind === "Classification" && (m.outE(n.id, "inSystem").some((e) => e.t === s.id) || n.scheme === s.label));
      const elems = new Set();
      for (const r of refs) for (const e of m.inE(r.id, "hasClassification")) elems.add(e.s);
      return `<tr><td>${esc(s.label)}</td><td>${esc(s.edition || "")}</td><td class="num">${refs.length}</td><td class="num">${elems.size}</td><td class="num">${pct(elems.size / Math.max(1, m.catalog.kindCount.get("Element") || 0))}</td></tr>`;
    }).join("");
    let cross = "";
    if (m.source === "csv") {
      const cw = new Map();
      for (const n of m.nodes.values()) if (n.sourceCategory) {
        const k = n.sourceCategory + "|" + n.cls;
        cw.set(k, (cw.get(k) || 0) + 1);
      }
      cross = `<div class="card"><h4>${esc(m.meta.csvSource)} category → IFC class crosswalk</h4><div class="table-wrap" style="max-height:240px"><table class="data"><thead><tr><th>Source category</th><th>IFC class</th><th style="text-align:right">Elements</th></tr></thead><tbody>${[...cw.entries()].sort((a, b) => b[1] - a[1]).map(([k, v]) => { const [a, b] = k.split("|"); return `<tr><td>${esc(a)}</td><td>${esc(b)}</td><td class="num">${v}</td></tr>`; }).join("")}</tbody></table></div></div>`;
    }
    return `<div class="kpis">
        <div class="kpi"><b>${rows.length}</b><span>Classes in model</span></div>
        <div class="kpi"><b>${pct(botCov)}</b><span>BOT coverage</span><em>of physical & spatial objects</em></div>
        <div class="kpi"><b>${pct(brickCov)}</b><span>Brick coverage</span><em>equipment & spaces</em></div>
        <div class="kpi"><b>${systems.length}</b><span>Classification systems</span></div>
      </div>
      <div class="grid-2" style="margin-bottom:14px">
        <div class="card"><h4>Classification systems</h4>${systems.length ? `<table class="data"><thead><tr><th>System</th><th>Edition</th><th style="text-align:right">References</th><th style="text-align:right">Elements</th><th style="text-align:right">Coverage</th></tr></thead><tbody>${sysRows}</tbody></table>` : `<div class="muted">No IfcClassification found. Associate Uniclass / OmniClass / CCI references to enable classification crosswalks.</div>`}</div>
        ${cross || `<div class="card"><h4>Ontology alignment</h4><div class="muted">Classes are mapped to <b>ifcOWL</b> (buildingSMART IFC4 ADD2 TC1 OWL), the <b>Building Topology Ontology</b> (bot:Site/Building/Storey/Space/Element/Zone) and <b>Brick</b> equipment classes. Relationships map to BOT topology (hasStorey, hasSpace, containsElement, hostsElement, adjacentElement, adjacentZone) and Brick (feeds, isPartOf). Export the result as RDF Turtle from the Export menu.</div></div>`}
      </div>
      <div class="toolbar"><b>Class mapping</b><span class="grow"></span><button class="btn sm" data-act="export" data-v="mapping">Export CSV</button></div>
      <div class="table-wrap"><table class="data"><thead><tr><th>IFC class</th><th>Kind</th><th style="text-align:right">Count</th><th>Discipline</th><th>ifcOWL</th><th>BOT</th><th>Brick</th><th style="text-align:right">Typed</th><th style="text-align:right">Classified</th><th style="text-align:right">Material</th><th style="text-align:right">Avg props</th></tr></thead>
      <tbody>${rows.map((r) => `<tr class="click" data-act="only-class" data-v="${esc(r.ifcClass)}"><td>${esc(r.ifcClass)}</td><td class="muted">${esc(r.kind)}</td><td class="num">${r.count.toLocaleString()}</td><td>${esc(r.discipline)}</td><td class="mono">${esc(r.ifcOWL)}</td><td class="mono">${esc(r.bot)}</td><td class="mono">${esc(r.brick)}</td><td class="num">${pct(r.typed)}</td><td class="num">${pct(r.classified)}</td><td class="num">${pct(r.withMaterial)}</td><td class="num">${r.avgProperties.toFixed(1)}</td></tr>`).join("")}</tbody></table></div>`;
  }

  function catalogRows(filter) {
    const q = (filter || "").toLowerCase();
    return [...state.model.catalog.props.values()].filter((p) => !q || p.key.toLowerCase().includes(q)).sort((a, b) => b.count - a.count);
  }
  function tabCatalog() {
    const rows = catalogRows(state.catalogFilter);
    const shown = rows.slice(0, 400);
    const sample = (p) => {
      if (p.min !== Infinity && p.datatypes.size && [...p.values.keys()].every((v) => !isNaN(+v))) return `${fmt(p.min)} … ${fmt(p.max)}`;
      return [...p.values.entries()].sort((a, b) => b[1] - a[1]).slice(0, 4).map(([v, c]) => `${v.length > 24 ? v.slice(0, 22) + "…" : v} (${c})`).join(", ");
    };
    return `<div class="toolbar"><input type="search" id="catFilter" placeholder="Filter properties…" value="${esc(state.catalogFilter || "")}"/><span class="muted">${rows.length.toLocaleString()} properties${rows.length > 400 ? " (showing 400)" : ""} · coverage = share of instances of the classes that use the property</span><span class="grow"></span><button class="btn sm" data-act="export" data-v="catalog">Export CSV</button></div>
      <div class="table-wrap"><table class="data"><thead><tr><th>Property set</th><th>Property</th><th></th><th>Data type</th><th>Unit</th><th style="text-align:right">Count</th><th style="text-align:right">Coverage</th><th style="text-align:right">Distinct</th><th>Values / range</th><th></th></tr></thead><tbody>
      ${shown.map((p) => `<tr><td>${esc(p.pset)}</td><td><b>${esc(p.name)}</b></td><td>${p.kind === "Q" ? '<span class="badge q">Q</span>' : ""}${p.inherited ? '<span class="badge t">type</span>' : ""}</td><td class="muted">${esc([...p.datatypes].map((d) => d.replace(/^Ifc/, "")).join(", "))}</td><td>${esc([...p.units].join(", "))}</td><td class="num">${p.count.toLocaleString()}</td><td class="num">${pct(p.coverage)}</td><td class="num">${p.distinct}</td><td class="muted" title="${esc(sample(p))}">${esc(sample(p))}</td>
        <td><button class="btn sm" data-act="prop-color" data-v="${esc(p.key)}">Colour</button> <button class="btn sm" data-act="prop-values" data-v="${esc(p.key)}">Value nodes</button> <button class="btn sm" data-act="prop-query" data-v="${esc(p.key)}">Query</button></td></tr>`).join("")}
      </tbody></table></div>`;
  }

  // ── Query builder ──────────────────────────────────────────────────────
  const ATTR_KEYS = ["Name", "GlobalId", "Class", "ObjectType", "Tag", "PredefinedType", "Discipline", "Storey", "Type", "Family", "Category", "Layer", "Phase", "Host", "Material", "Classification", "System"];
  const BD_ATTRS = ["Category", "Family", "Type", "Class", "Discipline", "Storey", "Layer", "Material", "System", "Classification", "Phase", "PredefinedType", "ObjectType"];
  const OPS = [["=", "="], ["!=", "≠"], ["contains", "contains"], [">", ">"], [">=", "≥"], ["<", "<"], ["<=", "≤"], ["exists", "exists"], ["missing", "is missing"]];
  state.query = { cls: "", domain: "", storey: "", conds: [{ key: "", op: "=", val: "" }] };
  const PRESETS = [
    { label: "External walls", q: { cls: "IfcWall", conds: [{ key: "*.IsExternal", op: "=", val: "true" }] } },
    { label: "Load-bearing elements", q: { cls: "", conds: [{ key: "*.LoadBearing", op: "=", val: "true" }] } },
    { label: "Doors without fire rating", q: { cls: "IfcDoor", conds: [{ key: "*.FireRating", op: "missing", val: "" }] } },
    { label: "Spaces > 20 m²", q: { cls: "IfcSpace", conds: [{ key: "*.NetFloorArea", op: ">", val: "20" }] } },
    { label: "Elements without material", q: { cls: "IfcElement", conds: [{ key: "Material", op: "missing", val: "" }] } },
    { label: "Proxies", q: { cls: "IfcBuildingElementProxy", conds: [] } },
  ];

  function qValue(n, key) {
    return state.model.valueOf(n, key);
  }
  function qTest(v, op, target) {
    if (op === "exists") return v != null && v !== "";
    if (op === "missing") return v == null || v === "";
    if (v == null) return false;
    const tn = parseFloat(target);
    const vn = typeof v === "number" ? v : parseFloat(v);
    const vs = String(v).toLowerCase();
    const ts = String(target).toLowerCase().trim();
    switch (op) {
      case "=": return typeof v === "boolean" ? String(v) === ts || (v && ["yes", "1", "t"].includes(ts)) || (!v && ["no", "0", "f"].includes(ts)) : !isNaN(tn) && !isNaN(vn) && typeof v === "number" ? vn === tn : vs === ts;
      case "!=": return !qTest(v, "=", target);
      case "contains": return vs.includes(ts);
      case ">": return vn > tn;
      case ">=": return vn >= tn;
      case "<": return vn < tn;
      case "<=": return vn <= tn;
    }
    return false;
  }
  function runQuery() {
    const q = state.query;
    const m = state.model;
    const out = [];
    for (const n of m.nodes.values()) {
      if (!["Element", "Spatial", "System", "Zone", "Type", "Opening", "Port"].includes(n.kind)) continue;
      if (q.cls && !(n.cls === q.cls || S.isA(n.cls, q.cls) || (S.isTypeClass(q.cls) && n.cls === q.cls))) continue;
      if (q.domain && n.domain !== q.domain) continue;
      if (q.storey && n.storey !== q.storey) continue;
      let ok = true;
      for (const c of q.conds) {
        if (!c.key) continue;
        if (!qTest(qValue(n, c.key), c.op, c.val)) { ok = false; break; }
      }
      if (ok) out.push(n.id);
    }
    state.queryResults = out;
    renderDrawerTabs();
    renderDrawer();
  }

  function tabQuery() {
    const m = state.model;
    const q = state.query;
    const classes = [...m.catalog.classCount.keys()].filter((c) => c.startsWith("Ifc")).sort();
    const abstract = ["IfcElement", "IfcBuiltElement", "IfcDistributionElement", "IfcDistributionFlowElement", "IfcFlowTerminal", "IfcFlowSegment", "IfcSpatialElement"].filter((a) => classes.some((c) => S.isA(c, a)));
    const opt = (v, l, cur) => `<option value="${esc(v)}" ${v === cur ? "selected" : ""}>${esc(l)}</option>`;
    const head = `<div class="q-head">
      <div class="field"><label>Class</label><select data-q="cls">${opt("", "Any", q.cls)}${abstract.map((c) => opt(c, c + " (and subclasses)", q.cls)).join("")}${classes.map((c) => opt(c, c, q.cls)).join("")}</select></div>
      <div class="field"><label>Discipline</label><select data-q="domain">${opt("", "Any", q.domain)}${[...m.catalog.domainCount.keys()].map((d) => opt(d, d, q.domain)).join("")}</select></div>
      <div class="field"><label>Storey</label><select data-q="storey">${opt("", "Any", q.storey)}${m.storeys.map((s) => opt(s, m.nodes.get(s).label, q.storey)).join("")}</select></div></div>`;
    const conds = q.conds.map((c, i) => `<div class="q-row"><input type="text" list="qKeys" data-qc="key" data-i="${i}" value="${esc(c.key)}" placeholder="Property (Pset.Name, *.Name) or attribute"/><select data-qc="op" data-i="${i}">${OPS.map(([v, l]) => opt(v, l, c.op)).join("")}</select><input type="text" data-qc="val" data-i="${i}" value="${esc(c.val)}" placeholder="value" ${c.op === "exists" || c.op === "missing" ? "disabled" : ""}/><button class="icon-btn" data-act="q-del" data-v="${i}" title="Remove">×</button></div>`).join("");
    const res = state.queryResults;
    let results = "";
    if (res) {
      const rows = res.slice(0, 300).map((id) => {
        const n = m.nodes.get(id);
        const vals = q.conds.filter((c) => c.key).map((c) => fmt(qValue(n, c.key))).join(" · ");
        return `<tr class="click" data-goto="${esc(id)}"><td>${esc(n.label)}</td><td class="muted">${esc(n.cls)}</td><td>${esc(n.storey ? m.nodes.get(n.storey)?.label : "")}</td><td class="mono">${esc(n.guid || "")}</td><td>${esc(vals)}</td></tr>`;
      }).join("");
      results = `<div class="toolbar" style="margin-top:12px"><b>${res.length.toLocaleString()} result${res.length === 1 ? "" : "s"}</b><span class="grow"></span>${res.length ? `<button class="btn sm" data-act="q-mark">Highlight in view</button><button class="btn sm" data-act="q-graph">Show as graph</button><button class="btn sm" data-act="q-graph-ctx">Graph + context</button><button class="btn sm" data-act="export" data-v="query">Export CSV</button>` : ""}</div>
        ${res.length ? `<div class="table-wrap"><table class="data"><thead><tr><th>Name</th><th>Class</th><th>Storey</th><th>GlobalId</th><th>Matched values</th></tr></thead><tbody>${rows}</tbody></table></div>` : ""}`;
    }
    return `<div class="presets"><span class="muted" style="align-self:center">Presets:</span>${PRESETS.map((p, i) => `<button class="chip" data-act="q-preset" data-v="${i}">${esc(p.label)}</button>`).join("")}</div>${head}<div class="field-label" style="margin-bottom:6px">Conditions (all must match)</div>${conds}
      <datalist id="qKeys">${ATTR_KEYS.map((k) => `<option value="${k}">`).join("")}${[...new Set([...m.catalog.props.values()].map((p) => "*." + p.name))].slice(0, 1500).map((k) => `<option value="${esc(k)}">`).join("")}${[...m.catalog.props.keys()].slice(0, 3000).map((k) => `<option value="${esc(k)}">`).join("")}</datalist>
      <div class="toolbar" style="margin-top:4px"><button class="btn sm" data-act="q-add">+ Condition</button><button class="btn sm primary" data-act="q-run">Run query</button><button class="btn sm ghost" data-act="q-reset">Reset</button></div>${results}`;
  }

  // ── Node table ─────────────────────────────────────────────────────────
  function tableRows() {
    const g = state.graph;
    const a = state.analysis;
    const q = (state.tableFilter || "").toLowerCase();
    let rows = g.nodes.map((n) => {
      const mt = a.metrics.get(n.id) || {};
      return { id: n.id, label: n.label, cls: n.cls || "", kind: n.kind, domain: n.domain || "", storey: n.storey ? state.model.nodes.get(n.storey)?.label || "" : "", count: n.count ?? "", degree: mt.degree, pagerank: mt.pagerank, betweenness: mt.betweenness, closeness: mt.closeness, clustering: mt.clustering, community: mt.community != null ? mt.community + 1 : "", integration: mt.integration };
    });
    if (q) rows = rows.filter((r) => (r.label + " " + r.cls + " " + r.storey + " " + r.domain).toLowerCase().includes(q));
    const { key, dir } = state.tableSort;
    rows.sort((x, y) => {
      const a1 = x[key], b1 = y[key];
      if (typeof a1 === "number" || typeof b1 === "number") return ((a1 ?? -Infinity) - (b1 ?? -Infinity)) * dir;
      return String(a1 ?? "").localeCompare(String(b1 ?? "")) * dir;
    });
    return rows;
  }
  function tabTable() {
    const rows = tableRows();
    const per = 150;
    const pages = Math.max(1, Math.ceil(rows.length / per));
    state.tablePage = Math.min(state.tablePage, pages - 1);
    const slice = rows.slice(state.tablePage * per, (state.tablePage + 1) * per);
    const cols = [["label", "Node"], ["cls", "Class"], ["kind", "Kind"], ["domain", "Discipline"], ["storey", "Storey"], ["count", "Members"], ["degree", "Degree"], ["pagerank", "PageRank"], ["betweenness", "Betweenness"], ["closeness", "Closeness"], ["clustering", "Clustering"], ["community", "Comm."]];
    if (state.settings.lens === "spaces") cols.push(["integration", "Integration"]);
    const num = new Set(["count", "degree", "pagerank", "betweenness", "closeness", "clustering", "community", "integration"]);
    const arrow = (k) => (state.tableSort.key === k ? (state.tableSort.dir > 0 ? " ▲" : " ▼") : "");
    return `<div class="toolbar"><input type="search" id="tableFilter" placeholder="Filter nodes…" value="${esc(state.tableFilter || "")}"/><span class="muted">${rows.length.toLocaleString()} nodes in current view</span><span class="grow"></span>
      <button class="btn sm" data-act="page" data-v="-1" ${state.tablePage <= 0 ? "disabled" : ""}>‹</button><span class="muted">${state.tablePage + 1} / ${pages}</span><button class="btn sm" data-act="page" data-v="1" ${state.tablePage >= pages - 1 ? "disabled" : ""}>›</button>
      <button class="btn sm" data-act="export" data-v="nodes">Export CSV</button></div>
      <div class="table-wrap"><table class="data"><thead><tr>${cols.map(([k, l]) => `<th class="sort" data-sort="${k}" style="${num.has(k) ? "text-align:right" : ""}">${l}${arrow(k)}</th>`).join("")}</tr></thead><tbody>
      ${slice.map((r) => `<tr class="click" data-select="${esc(r.id)}">${cols.map(([k]) => `<td class="${num.has(k) ? "num" : ""}">${esc(num.has(k) ? fmt(r[k], 4) : k === "cls" ? r[k] : r[k])}</td>`).join("")}</tr>`).join("")}
      </tbody></table></div>`;
  }

  // ── Exports ────────────────────────────────────────────────────────────
  function renderExportMenu() {
    const sc = state.exportScope;
    $("exportMenu").innerHTML = `
      <h6>Current view</h6>
      <button class="menu-item" data-exp="graphml">GraphML <small>Gephi · yEd · Cytoscape · NetworkX</small></button>
      <button class="menu-item" data-exp="json">JSON node-link <small>with properties & metrics</small></button>
      <button class="menu-item" data-exp="nodes">Nodes CSV <small>with centralities</small></button>
      <button class="menu-item" data-exp="edges">Edges CSV</button>
      <button class="menu-item" data-exp="svg">SVG <small>vector, publication-ready</small></button>
      <button class="menu-item" data-exp="png">PNG <small>screenshot</small></button>
      <h6>Knowledge graph</h6>
      <div class="menu-scope"><button class="chip ${sc === "visible" ? "on" : ""}" data-scope="visible">Visible nodes</button><button class="chip ${sc === "all" ? "on" : ""}" data-scope="all">Entire model</button></div>
      <button class="menu-item" data-exp="ttl">RDF Turtle <small>ifcOWL · BOT · Brick · PROPS</small></button>
      <button class="menu-item" data-exp="cypher">Neo4j Cypher <small>labelled property graph</small></button>
      <h6>Reports</h6>
      <button class="menu-item" data-exp="health">Model-health findings <small>CSV</small></button>
      <button class="menu-item" data-exp="mapping">Semantic mapping <small>CSV</small></button>
      <button class="menu-item" data-exp="catalog">Property catalogue <small>CSV</small></button>`;
  }

  async function doExport(kind) {
    const m = state.model;
    if (!m) return;
    const base = X.slug(m.meta.fileName) + "_" + state.settings.lens;
    const ids = state.exportScope === "all" ? null : X.modelIds(state.graph);
    try {
      switch (kind) {
        case "graphml": X.download(base + ".graphml", X.toGraphML(state.graph, m, state.analysis), "application/xml"); break;
        case "json": X.download(base + ".json", X.toJSON(state.graph, m, state.analysis, state.settings), "application/json"); break;
        case "nodes": X.download(base + "_nodes.csv", X.nodesCSV(state.graph, m, state.analysis), "text/csv"); break;
        case "edges": X.download(base + "_edges.csv", X.edgesCSV(state.graph), "text/csv"); break;
        case "svg": {
          if (renderer.mode !== "2d") { toast("Switch to 2D for SVG export"); return; }
          const title = `${m.meta.fileName} — ${V.LENSES.find((l) => l.id === state.settings.lens).title}`;
          X.download(base + ".svg", renderer.toSVG(title, state.legend.type === "cat" ? state.legend.items : []), "image/svg+xml");
          break;
        }
        case "png": X.download(base + ".png", await renderer.screenshotPNG(), "image/png"); break;
        case "ttl": X.download(X.slug(m.meta.fileName) + (ids ? "_view" : "") + ".ttl", X.toTurtle(m, ids), "text/turtle"); break;
        case "cypher": X.download(X.slug(m.meta.fileName) + (ids ? "_view" : "") + ".cypher", X.toCypher(m, ids), "text/plain"); break;
        case "health": X.download(X.slug(m.meta.fileName) + "_model_health.csv", X.healthCSV(state.health, m), "text/csv"); break;
        case "mapping": X.download(X.slug(m.meta.fileName) + "_semantic_mapping.csv", X.toCSV(mappingRows()), "text/csv"); break;
        case "catalog":
          X.download(X.slug(m.meta.fileName) + "_property_catalogue.csv", X.toCSV(catalogRows("").map((p) => ({ propertySet: p.pset, property: p.name, kind: p.kind === "Q" ? "quantity" : "property", datatypes: [...p.datatypes].join(" "), units: [...p.units].join(" "), count: p.count, coverage: +p.coverage.toFixed(4), distinct: p.distinct, min: p.min === Infinity ? "" : p.min, max: p.max === -Infinity ? "" : p.max, classes: [...p.classes.keys()].join(" ") }))), "text/csv");
          break;
        case "query":
          X.download(X.slug(m.meta.fileName) + "_query.csv", X.toCSV((state.queryResults || []).map((id) => {
            const n = m.nodes.get(id);
            const r = { id, globalId: n.guid || "", name: n.label, ifcClass: n.cls, storey: n.storey ? m.nodes.get(n.storey)?.label : "" };
            state.query.conds.filter((c) => c.key).forEach((c) => (r[c.key] = qValue(n, c.key) ?? ""));
            return r;
          })), "text/csv");
          break;
      }
      $("exportMenu").classList.remove("open");
    } catch (err) {
      console.error(err);
      toast("Export failed: " + err.message);
    }
  }

  // ── Mobile helpers ─────────────────────────────────────────────────────
  function closeSideOnMobile() {
    if (window.innerWidth <= 820) $("sidebar").classList.remove("open");
  }

  // ── Event wiring ───────────────────────────────────────────────────────
  function wire() {
    const fileInput = $("fileInput");
    $("btnOpen").onclick = () => fileInput.click();
    $("dropZone").onclick = (e) => {
      if (e.target.closest("[data-sample]")) return;
      fileInput.click();
    };
    $("dropZone").onkeydown = (e) => {
      if (e.key === "Enter" || e.key === " ") fileInput.click();
    };
    document.querySelectorAll("[data-sample]").forEach((b) => {
      b.onclick = (e) => {
        e.stopPropagation();
        loadSample(b.dataset.sample);
      };
    });
    fileInput.onchange = () => {
      if (fileInput.files[0]) openFile(fileInput.files[0]);
      fileInput.value = "";
    };
    // drag & drop anywhere
    let dragDepth = 0;
    document.addEventListener("dragenter", (e) => { e.preventDefault(); dragDepth++; $("dropZone").classList.add("over"); });
    document.addEventListener("dragleave", () => { if (--dragDepth <= 0) $("dropZone").classList.remove("over"); });
    document.addEventListener("dragover", (e) => e.preventDefault());
    document.addEventListener("drop", (e) => {
      e.preventDefault();
      dragDepth = 0;
      $("dropZone").classList.remove("over");
      const f = e.dataTransfer.files[0];
      if (f) openFile(f);
    });

    $("btnTheme").onclick = () => {
      const next = theme() === "dark" ? "light" : "dark";
      document.documentElement.dataset.theme = next;
      try { localStorage.setItem("kg-theme", next); } catch (e) {}
      if (state.graph) {
        computeStyle();
        updateLegend();
        if (renderer.fg3) renderer.fg3.backgroundColor(next === "dark" ? "#0c111b" : "#fbfcfe");
        renderer.restyle();
      }
    };
    $("btnExport").onclick = (e) => {
      e.stopPropagation();
      renderExportMenu();
      $("exportMenu").classList.toggle("open");
    };
    $("exportMenu").onclick = (e) => {
      e.stopPropagation();
      const sc = e.target.closest("[data-scope]");
      if (sc) { state.exportScope = sc.dataset.scope; renderExportMenu(); return; }
      const it = e.target.closest("[data-exp]");
      if (it) doExport(it.dataset.exp);
    };
    $("btnSide").onclick = () => $("sidebar").classList.toggle("open");

    // lenses
    $("lensList").onclick = (e) => {
      const b = e.target.closest("[data-lens]");
      if (b && state.model) setLens(b.dataset.lens);
    };
    $("lensOptions").addEventListener("change", (e) => {
      if (e.target.dataset.bd !== undefined) {
        state.settings.breakdown[+e.target.dataset.bd] = e.target.value.trim();
        rebuild();
        return;
      }
      const k = e.target.dataset.opt;
      if (k) { state.settings[k] = e.target.checked; rebuild(); }
      if (e.target.id === "valueKey") { state.settings.valueKey = e.target.value.trim(); rebuild(); }
    });
    $("detailSel").onchange = (e) => { state.settings.detail = e.target.value; rebuild(); };
    $("maxNodes").onchange = (e) => { state.settings.maxNodes = Math.max(100, +e.target.value || 3000); rebuild(); };

    // filters
    $("domainChips").onclick = (e) => {
      const c = e.target.closest("[data-domain-chip]");
      if (!c) return;
      const all = [...state.model.catalog.domainCount.keys()];
      const d = c.dataset.domainChip;
      const cur = state.settings.domains;
      toggleSetFilter("domains", d, all, cur ? !cur.has(d) : false);
      renderFilters();
      rebuildSoon();
    };
    $("sidebar").addEventListener("change", (e) => {
      const t = e.target;
      const m = state.model;
      if (!m) return;
      if (t.dataset.kind) {
        if (t.checked) state.settings.kinds.add(t.dataset.kind);
        else state.settings.kinds.delete(t.dataset.kind);
        rebuildSoon();
      } else if (t.dataset.storey) {
        toggleSetFilter("storeys", t.dataset.storey, m.storeys, t.checked);
        rebuildSoon();
      } else if (t.dataset.class) {
        const all = [...m.catalog.classCount.keys()];
        toggleSetFilter("classes", t.dataset.class, all, t.checked);
        rebuildSoon();
      } else if (t.dataset.cat) {
        toggleSetFilter("cats", t.dataset.cat, [...new Set([...m.catalog.catCount.keys(), "Properties", "Schema", "Typing", "Classification"])], t.checked);
        rebuildSoon();
      }
    });
    $("classSearch").oninput = () => renderClassList();
    document.querySelectorAll("[data-all]").forEach((b) => {
      b.onclick = () => {
        const k = { storey: "storeys", class: "classes", cat: "cats" }[b.dataset.all];
        state.settings[k] = null;
        renderFilters();
        rebuild();
      };
    });
    $("btnResetFilters").onclick = () => {
      Object.assign(state.settings, { domains: null, classes: null, storeys: null, cats: null, kinds: new Set(DEFAULT_KINDS), hidden: new Set() });
      $("classSearch").value = "";
      renderFilters();
      rebuild();
    };
    $("btnUnhide").onclick = () => {
      state.settings.hidden = new Set();
      updateHiddenRow();
      rebuild({ fit: false });
    };

    // appearance
    $("layoutSel").onchange = (e) => {
      state.style.layout = e.target.value;
      renderer.applyLayout(e.target.value);
    };
    $("colorSel").onchange = (e) => {
      state.style.colorBy = e.target.value;
      $("colorPropField").hidden = e.target.value !== "property";
      if (e.target.value === "integration" && state.settings.lens !== "spaces") toast("Integration is computed in the Space adjacency lens");
      if (e.target.value === "property" && !state.style.colorProp) { $("colorProp").focus(); return; }
      restyle();
    };
    $("colorProp").onchange = (e) => { state.style.colorProp = e.target.value.trim(); restyle(); };
    $("sizeSel").onchange = (e) => { state.style.sizeBy = e.target.value; restyle(); renderer.applyForces(); };
    $("labelSel").onchange = (e) => { state.style.labels = e.target.value; renderer.style.labels = e.target.value; renderer.restyle(); if (renderer.mode === "3d") renderer.fg3.nodeThreeObject(renderer.fg3.nodeThreeObject()); };
    $("optArrows").onchange = (e) => { state.style.arrows = renderer.style.arrows = e.target.checked; renderer.mode === "2d" ? renderer.fg.linkDirectionalArrowLength(renderer.fg.linkDirectionalArrowLength()) : renderer.fg3.linkDirectionalArrowLength(renderer.fg3.linkDirectionalArrowLength()); };
    $("optParticles").onchange = (e) => { state.style.particles = renderer.style.particles = e.target.checked; renderer.mode === "2d" ? renderer.fg.linkDirectionalParticles(renderer.fg.linkDirectionalParticles()) : renderer.fg3.linkDirectionalParticles(renderer.fg3.linkDirectionalParticles()); };
    $("optInferred").onchange = (e) => { state.style.inferred = e.target.checked; rebuild({ fit: false }); };
    $("optEdgeColor").onchange = (e) => { state.style.edgeColor = e.target.checked; restyle(); };

    // stage tools
    $("tFit").onclick = () => renderer.fit();
    $("tZoomIn").onclick = () => renderer.zoomBy(1.4);
    $("tZoomOut").onclick = () => renderer.zoomBy(1 / 1.4);
    $("tPause").onclick = () => $("tPause").classList.toggle("active", renderer.togglePause());
    $("tReheat").onclick = () => renderer.reheat();
    $("t3d").onclick = async () => {
      const to = renderer.mode === "2d" ? "3d" : "2d";
      try {
        if (to === "3d") toast("Loading 3D engine…");
        await renderer.setMode(to);
        $("t3d").classList.toggle("active", to === "3d");
      } catch (err) {
        toast(err.message);
      }
    };
    $("tPng").onclick = () => doExport("png");
    $("tSvg").onclick = () => doExport("svg");

    // search
    const search = $("search");
    search.oninput = debounce(() => runSearch(search.value), 120);
    search.onfocus = () => search.value && runSearch(search.value);
    search.onkeydown = (e) => {
      const items = [...$("searchResults").querySelectorAll(".sr[data-goto]")];
      if (e.key === "ArrowDown" || e.key === "ArrowUp") {
        e.preventDefault();
        searchSel = Math.max(0, Math.min(items.length - 1, searchSel + (e.key === "ArrowDown" ? 1 : -1)));
        items.forEach((it, i) => it.classList.toggle("kb", i === searchSel));
        items[searchSel]?.scrollIntoView({ block: "nearest" });
      } else if (e.key === "Enter") {
        const it = items[Math.max(0, searchSel)];
        if (it) { goTo(it.dataset.goto); $("searchResults").classList.remove("open"); }
      } else if (e.key === "Escape") {
        $("searchResults").classList.remove("open");
        search.blur();
      }
    };

    // global delegated clicks
    document.addEventListener("click", (e) => {
      if (!e.target.closest(".menu")) $("exportMenu").classList.remove("open");
      if (!e.target.closest("#searchBox")) $("searchResults").classList.remove("open");
      if (!e.target.closest("#ctxMenu")) hideCtx();
      const ctx = e.target.closest("[data-ctx]");
      if (ctx) { ctxAction(ctx.dataset.ctx, $("ctxMenu").dataset.id); return; }
      const go = e.target.closest("[data-goto]");
      if (go) {
        goTo(go.dataset.goto);
        if (go.closest("#searchResults")) $("searchResults").classList.remove("open");
        return;
      }
      const sel = e.target.closest("[data-select]");
      if (sel) { select(sel.dataset.select); renderer.focus(sel.dataset.select); return; }
      const tab = e.target.closest("[data-tab]");
      if (tab) {
        if (state.drawerTab === tab.dataset.tab && !$("drawer").classList.contains("collapsed")) toggleDrawer();
        else openDrawer(tab.dataset.tab);
        return;
      }
      const srt = e.target.closest("[data-sort]");
      if (srt) {
        const k = srt.dataset.sort;
        state.tableSort = { key: k, dir: state.tableSort.key === k ? -state.tableSort.dir : ["label", "cls", "kind", "domain", "storey"].includes(k) ? 1 : -1 };
        renderDrawer();
        return;
      }
      const ld = e.target.closest("[data-domain]");
      if (ld && ld.closest("#legend")) {
        state.settings.domains = new Set([ld.dataset.domain]);
        renderFilters();
        rebuild();
        return;
      }
      const cp = e.target.closest("[data-copy]");
      if (cp) { copy(cp.dataset.copy); return; }
      const act = e.target.closest("[data-act]");
      if (act) handleAct(act.dataset.act, act);
    });

    // drawer inputs
    $("drawerBody").addEventListener("input", debounce((e) => {
      if (e.target.id === "catFilter") { state.catalogFilter = e.target.value; renderDrawer(); const el = $("catFilter"); el.focus(); el.setSelectionRange(el.value.length, el.value.length); }
      if (e.target.id === "tableFilter") { state.tableFilter = e.target.value; state.tablePage = 0; renderDrawer(); const el = $("tableFilter"); el.focus(); el.setSelectionRange(el.value.length, el.value.length); }
    }, 250));
    $("drawerBody").addEventListener("change", (e) => {
      const t = e.target;
      if (t.dataset.q) { state.query[t.dataset.q] = t.value; }
      if (t.dataset.qc) {
        state.query.conds[+t.dataset.i][t.dataset.qc] = t.value;
        if (t.dataset.qc === "op") renderDrawer();
      }
    });
    $("drawerBody").addEventListener("keydown", (e) => {
      if (e.key === "Enter" && e.target.dataset.qc) {
        state.query.conds[+e.target.dataset.i][e.target.dataset.qc] = e.target.value;
        runQuery();
      }
    });

    // drawer resize
    let drag = null;
    $("drawerHandle").addEventListener("mousedown", (e) => {
      drag = { y: e.clientY, h: $("drawer").offsetHeight };
      document.body.style.userSelect = "none";
    });
    window.addEventListener("mousemove", (e) => {
      mouse.x = e.clientX;
      mouse.y = e.clientY;
      if (!drag) return;
      const h = Math.max(120, Math.min(window.innerHeight - 200, drag.h + drag.y - e.clientY));
      document.documentElement.style.setProperty("--drawer-h", h + "px");
      $("drawer").classList.remove("collapsed");
    });
    window.addEventListener("mouseup", () => {
      if (drag) { drag = null; document.body.style.userSelect = ""; renderer.resize(); }
    });

    // keyboard
    document.addEventListener("keydown", (e) => {
      const typing = /INPUT|SELECT|TEXTAREA/.test(document.activeElement?.tagName);
      if (e.key === "Escape") {
        hideCtx();
        $("exportMenu").classList.remove("open");
        if (!typing) { renderer.marked = new Set(); renderer.clearHighlight(); select(null); }
      }
      if (typing || !state.model) return;
      if (e.key === "/") { e.preventDefault(); $("search").focus(); }
      else if (e.key === "f" || e.key === "F") renderer.fit();
      else if (e.key === " ") { e.preventDefault(); $("tPause").click(); }
    });
    window.addEventListener("contextmenu", (e) => {
      if (e.target.closest("#graph")) e.preventDefault();
    });
  }

  function restyle() {
    if (!state.graph) return;
    computeStyle();
    updateLegend();
    renderer.restyle();
    if (state.style.layout === "clusters" || state.style.layout === "circular") renderer.applyLayout(state.style.layout);
  }

  function handleAct(act, el) {
    const v = el.dataset.v;
    const m = state.model;
    switch (act) {
      case "dismiss": el.closest(".note").remove(); break;
      case "leave-subgraph": setLens(state.prevLens || "kg"); break;
      case "legend-toggle": $("legend").classList.toggle("collapsed"); updateLegend(); break;
      case "drawer-toggle": toggleDrawer(); break;
      case "close-insp": $("inspector").classList.remove("open"); break;
      case "focus": renderer.focus(el.dataset.id); break;
      case "explore": exploreNeighborhood(el.dataset.id, +el.dataset.hops || 1); break;
      case "path-start": pathAction(el.dataset.id); break;
      case "hide":
        state.settings.hidden.add(el.dataset.id);
        updateHiddenRow();
        select(null);
        rebuild({ fit: false });
        break;
      case "expand-members": {
        const vn = viewNode(el.dataset.id);
        if (vn?.members) showSubgraph(new Set(vn.members), `Members of “${vn.label}”.`);
        break;
      }
      case "filter-class":
        state.settings.classes = new Set([el.dataset.cls]);
        renderFilters();
        rebuild();
        break;
      case "only-domain":
        state.settings.domains = new Set([v]);
        renderFilters();
        rebuild();
        break;
      case "only-class":
        if (S.kindOf(v) === "Element" || m.source === "csv") {
          state.settings.classes = new Set([v]);
          renderFilters();
          rebuild();
        } else {
          state.query = { cls: v, domain: "", storey: "", conds: [] };
          runQuery();
          openDrawer("query");
        }
        break;
      case "only-storey":
        state.settings.storeys = new Set([v]);
        renderFilters();
        rebuild();
        break;
      case "color":
        state.style.colorBy = v;
        syncControls();
        restyle();
        break;
      case "community": {
        const ids = new Set();
        for (const [id, mt] of state.analysis.metrics) if (mt.community === +v) { const vn = viewNode(id); if (vn?.ref) ids.add(vn.ref); else if (vn?.members) vn.members.forEach((x) => ids.add(x)); }
        showSubgraph(ids, `Community C${+v + 1} (${ids.size} nodes).`);
        break;
      }
      case "health-show": {
        const c = state.health.checks.find((x) => x.id === v);
        const ids = new Set(c.affected.slice(0, 2500));
        for (const id of [...ids]) for (const p of m.spatialPath(id).slice(0, 2)) ids.add(p);
        showSubgraph(ids, `Model health — ${c.title}: ${c.failed.toLocaleString()} affected objects (with their spatial containers).`);
        break;
      }
      case "health-mark": {
        const c = state.health.checks.find((x) => x.id === v);
        const set = new Set(c.affected);
        renderer.marked = new Set(state.graph.nodes.filter((n) => set.has(n.ref) || (n.members && n.members.some((x) => set.has(x)))).map((n) => n.id));
        renderer.redraw();
        toast(`${renderer.marked.size} nodes in the current view are marked (amber ring). Press Esc to clear.`);
        break;
      }
      case "health-clear": renderer.marked = new Set(); renderer.redraw(); break;
      case "export": doExport(v); break;
      case "prop-color":
        state.style.colorBy = "property";
        state.style.colorProp = v;
        syncControls();
        restyle();
        toast(`Colouring by ${v}`);
        break;
      case "prop-values":
        state.settings.valueKey = v;
        if (state.settings.lens !== "kg") setLens("kg");
        else { renderLensOptions(); rebuild(); }
        toast(`Property values of ${v} are shown as nodes — elements sharing a value cluster together.`);
        break;
      case "prop-query":
        state.query = { cls: "", domain: "", storey: "", conds: [{ key: v, op: "exists", val: "" }] };
        state.queryResults = null;
        openDrawer("query");
        break;
      case "q-add": state.query.conds.push({ key: "", op: "=", val: "" }); renderDrawer(); break;
      case "q-del": state.query.conds.splice(+v, 1); if (!state.query.conds.length) state.query.conds.push({ key: "", op: "=", val: "" }); renderDrawer(); break;
      case "q-reset": state.query = { cls: "", domain: "", storey: "", conds: [{ key: "", op: "=", val: "" }] }; state.queryResults = null; renderDrawerTabs(); renderDrawer(); break;
      case "q-run": runQuery(); break;
      case "q-preset": {
        const p = PRESETS[+v].q;
        state.query = { cls: p.cls, domain: "", storey: "", conds: p.conds.map((c) => ({ ...c })) };
        if (!state.query.conds.length) state.query.conds.push({ key: "", op: "=", val: "" });
        runQuery();
        break;
      }
      case "q-mark": {
        const set = new Set(state.queryResults);
        const ids = state.graph.nodes.filter((n) => set.has(n.ref) || (n.members && n.members.some((x) => set.has(x)))).map((n) => n.id);
        renderer.highlightSet(ids);
        toast(`${ids.length} matching nodes in the current view highlighted. Press Esc to clear.`);
        break;
      }
      case "q-graph":
      case "q-graph-ctx": {
        const ids = new Set(state.queryResults.slice(0, 3000));
        if (act === "q-graph-ctx") for (const id of [...ids]) for (const w of m.neighbors(id)) if (ids.size < 5000) ids.add(w);
        showSubgraph(ids, `Query result: ${state.queryResults.length.toLocaleString()} objects${act === "q-graph-ctx" ? " with their direct neighbours" : ""}.`, new Set(state.queryResults.slice(0, 3000)));
        break;
      }
      case "search-subgraph":
        if (state.lastSearch?.length) showSubgraph(new Set(state.lastSearch.slice(0, 3000)), `Search results (${state.lastSearch.length.toLocaleString()}).`);
        $("searchResults").classList.remove("open");
        break;
      case "page": state.tablePage += +v; renderDrawer(); break;
    }
  }

  // ── Boot ───────────────────────────────────────────────────────────────
  function boot() {
    initRenderer();
    wire();
    renderDrawerTabs();
    renderInspector();
    renderLensList();
    const params = new URLSearchParams(location.search);
    const smp = params.get("sample");
    if (smp) loadSample(SAMPLES[smp] ? smp : "ifc");
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot);
  else boot();

  KG.App = { state, loadFromText, rebuild, setLens, runQuery, select, goTo, exploreNeighborhood, pathAction, get renderer() { return renderer; } };
})((window.KG = window.KG || {}));
