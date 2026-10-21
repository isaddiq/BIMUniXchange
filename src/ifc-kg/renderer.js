/* BIMUniXchange · BIM Knowledge Graph Studio
 * renderer.js — 2D canvas (force-graph) and 3D WebGL (3d-force-graph, lazy)
 * rendering with semantic node shapes, layouts (force, hierarchical, radial,
 * storey layers, clusters, circular), highlighting and vector SVG export.
 */
(function (KG) {
  "use strict";
  const S = KG.Schema;
  const idOf = (x) => (typeof x === "object" && x ? x.id : x);

  const LIBS_3D = [
    "https://cdn.jsdelivr.net/npm/three@0.160.0/build/three.min.js",
    "https://cdn.jsdelivr.net/npm/three-spritetext@1.9.6/dist/three-spritetext.min.js",
    "https://cdn.jsdelivr.net/npm/3d-force-graph@1.77.0/dist/3d-force-graph.min.js",
  ];
  function loadScript(src) {
    return new Promise((res, rej) => {
      if (document.querySelector(`script[src="${src}"]`)) return res();
      const s = document.createElement("script");
      s.src = src;
      s.crossOrigin = "anonymous";
      s.onload = res;
      s.onerror = () => rej(new Error("Failed to load " + src));
      document.head.appendChild(s);
    });
  }

  function hexToRgb(hex) {
    const h = hex.replace("#", "");
    const n = parseInt(h.length === 3 ? h.split("").map((c) => c + c).join("") : h, 16);
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  }
  function rgba(color, a) {
    if (!color) return `rgba(150,150,150,${a})`;
    if (color.startsWith("rgb")) {
      const m = color.match(/[\d.]+/g);
      return `rgba(${m[0]},${m[1]},${m[2]},${a})`;
    }
    const [r, g, b] = hexToRgb(color);
    return `rgba(${r},${g},${b},${a})`;
  }

  /** Draw a semantic node shape centred on x,y with radius r. */
  function shapePath(ctx, shape, x, y, r) {
    ctx.beginPath();
    switch (shape) {
      case "square":
      case "squareo":
        ctx.rect(x - r * 0.88, y - r * 0.88, r * 1.76, r * 1.76);
        break;
      case "roundrect": {
        const w = r * 1.9, h = r * 1.5, rr = r * 0.35;
        ctx.moveTo(x - w / 2 + rr, y - h / 2);
        ctx.arcTo(x + w / 2, y - h / 2, x + w / 2, y + h / 2, rr);
        ctx.arcTo(x + w / 2, y + h / 2, x - w / 2, y + h / 2, rr);
        ctx.arcTo(x - w / 2, y + h / 2, x - w / 2, y - h / 2, rr);
        ctx.arcTo(x - w / 2, y - h / 2, x + w / 2, y - h / 2, rr);
        break;
      }
      case "diamond":
        ctx.moveTo(x, y - r * 1.2);
        ctx.lineTo(x + r * 1.2, y);
        ctx.lineTo(x, y + r * 1.2);
        ctx.lineTo(x - r * 1.2, y);
        ctx.closePath();
        break;
      case "triangle":
        ctx.moveTo(x, y - r * 1.2);
        ctx.lineTo(x + r * 1.1, y + r * 0.8);
        ctx.lineTo(x - r * 1.1, y + r * 0.8);
        ctx.closePath();
        break;
      case "hexagon":
        for (let i = 0; i < 6; i++) {
          const a = (Math.PI / 3) * i + Math.PI / 6;
          const px = x + r * 1.1 * Math.cos(a);
          const py = y + r * 1.1 * Math.sin(a);
          i ? ctx.lineTo(px, py) : ctx.moveTo(px, py);
        }
        ctx.closePath();
        break;
      case "tag":
        ctx.moveTo(x - r, y - r * 0.7);
        ctx.lineTo(x + r * 0.45, y - r * 0.7);
        ctx.lineTo(x + r * 1.15, y);
        ctx.lineTo(x + r * 0.45, y + r * 0.7);
        ctx.lineTo(x - r, y + r * 0.7);
        ctx.closePath();
        break;
      default:
        ctx.arc(x, y, r, 0, 2 * Math.PI);
    }
  }

  function shapeSVG(shape, x, y, r) {
    const f = (v) => +v.toFixed(2);
    switch (shape) {
      case "square":
      case "squareo":
        return `<rect x="${f(x - r * 0.88)}" y="${f(y - r * 0.88)}" width="${f(r * 1.76)}" height="${f(r * 1.76)}"`;
      case "roundrect":
        return `<rect x="${f(x - r * 0.95)}" y="${f(y - r * 0.75)}" width="${f(r * 1.9)}" height="${f(r * 1.5)}" rx="${f(r * 0.35)}"`;
      case "diamond":
        return `<polygon points="${f(x)},${f(y - r * 1.2)} ${f(x + r * 1.2)},${f(y)} ${f(x)},${f(y + r * 1.2)} ${f(x - r * 1.2)},${f(y)}"`;
      case "triangle":
        return `<polygon points="${f(x)},${f(y - r * 1.2)} ${f(x + r * 1.1)},${f(y + r * 0.8)} ${f(x - r * 1.1)},${f(y + r * 0.8)}"`;
      case "hexagon": {
        const pts = [];
        for (let i = 0; i < 6; i++) {
          const a = (Math.PI / 3) * i + Math.PI / 6;
          pts.push(`${f(x + r * 1.1 * Math.cos(a))},${f(y + r * 1.1 * Math.sin(a))}`);
        }
        return `<polygon points="${pts.join(" ")}"`;
      }
      case "tag":
        return `<polygon points="${f(x - r)},${f(y - r * 0.7)} ${f(x + r * 0.45)},${f(y - r * 0.7)} ${f(x + r * 1.15)},${f(y)} ${f(x + r * 0.45)},${f(y + r * 0.7)} ${f(x - r)},${f(y + r * 0.7)}"`;
      default:
        return `<circle cx="${f(x)}" cy="${f(y)}" r="${f(r)}"`;
    }
  }

  class Renderer {
    constructor(el, hooks) {
      this.el = el;
      this.hooks = hooks; // { onClick, onRightClick, onHover, onBackground, theme() }
      this.mode = "2d";
      this.graph = { nodes: [], links: [] };
      this.pos = new Map();
      this.hl = { nodes: new Set(), links: new Set(), active: false };
      this.selected = null;
      this.marked = new Set();
      this.style = { labels: "auto", arrows: true, particles: true, layout: "force", inferred: true };
      this.nbrs = new Map();
      this.fg = null;
      this.fg3 = null;
      this.paused = false;
      this.init2D();
      new ResizeObserver(() => this.resize()).observe(el);
    }

    theme() {
      return this.hooks.theme();
    }

    init2D() {
      const self = this;
      this.fg = ForceGraph()(this.el)
        .nodeId("id")
        .nodeLabel(() => "")
        .backgroundColor("rgba(0,0,0,0)")
        .nodeRelSize(1)
        .nodeVal((n) => (n._r || 4) ** 2)
        .nodeCanvasObject((n, ctx, scale) => self.drawNode(n, ctx, scale))
        .nodePointerAreaPaint((n, color, ctx) => {
          ctx.fillStyle = color;
          ctx.beginPath();
          ctx.arc(n.x, n.y, Math.max(n._r || 4, 3) * 1.25, 0, 2 * Math.PI);
          ctx.fill();
        })
        .linkColor((l) => self.linkColor(l))
        .linkWidth((l) => self.linkWidth(l))
        .linkLineDash((l) => (l.inferred ? [3, 2] : null))
        .linkCurvature((l) => l.curvature || 0)
        .linkDirectionalArrowLength((l) => (self.style.arrows ? Math.min(6, 2.5 + (self.linkWidth(l) || 1)) : 0))
        .linkDirectionalArrowRelPos(1)
        .linkDirectionalArrowColor((l) => self.linkColor(l, 0.9))
        .linkDirectionalParticles((l) => (self.style.particles && (l.pred === "feeds" || l.pred === "portConnectedTo") ? 2 : 0))
        .linkDirectionalParticleWidth(2.2)
        .linkDirectionalParticleSpeed(0.006)
        .linkDirectionalParticleColor((l) => self.linkColor(l, 1))
        .onNodeClick((n, ev) => self.hooks.onClick(n, ev))
        .onNodeRightClick((n, ev) => self.hooks.onRightClick(n, ev))
        .onNodeHover((n) => {
          self.el.style.cursor = n ? "pointer" : "";
          self.hover(n);
          self.hooks.onHover(n);
        })
        .onLinkHover((l) => self.hooks.onLinkHover && self.hooks.onLinkHover(l))
        .onBackgroundClick(() => self.hooks.onBackground())
        .onBackgroundRightClick(() => self.hooks.onBackground())
        .onNodeDragEnd((n) => {
          if (self.style.layout === "force" || self.style.layout === "clusters") {
            n.fx = n.x;
            n.fy = n.y;
            n._pinned = true;
          }
        })
        .onDagError(() => {})
        .d3AlphaDecay(0.028)
        .d3VelocityDecay(0.35)
        .cooldownTime(15000)
        .onEngineStop(() => {
          self.savePositions();
          if (self._fitOnStop && !self._userMoved) self.fit(700);
          self._fitOnStop = false;
        });
      this.c2d = this.el.firstElementChild;
      const moved = () => (this._userMoved = true);
      this.el.addEventListener("wheel", moved, { passive: true });
      this.el.addEventListener("pointerdown", moved);
      this.resize();
    }

    resize() {
      const w = this.el.clientWidth;
      const h = this.el.clientHeight;
      if (this.mode === "2d" && this.fg) this.fg.width(w).height(h);
      if (this.mode === "3d" && this.fg3) this.fg3.width(w).height(h);
    }

    // ── styling hooks ──
    linkColor(l, alpha) {
      const dark = this.theme() === "dark";
      const base = l._c || (dark ? "#64748b" : "#94a3b8");
      if (this.hl.active) {
        if (this.hl.links.has(l)) return rgba(base, alpha ?? 0.95);
        return rgba(base, 0.05);
      }
      if (l.onPath) return rgba("#f43f5e", 1);
      const a = alpha ?? (l.coverage != null ? 0.15 + 0.7 * l.coverage : this.graph.links.length > 4000 ? 0.28 : 0.5);
      return rgba(base, a);
    }
    linkWidth(l) {
      let w = 0.6;
      if (l.weight > 1) w = Math.min(6, 0.6 + Math.log2(l.weight) * 0.7);
      if (l.onPath) w = 3;
      if (this.hl.active && this.hl.links.has(l)) w = Math.max(w, 1.6);
      return w;
    }

    drawNode(n, ctx, scale) {
      const r = n._r || 4;
      const dim = this.hl.active && !this.hl.nodes.has(n.id);
      const dark = this.theme() === "dark";
      const bg = dark ? "#0c111b" : "#ffffff";
      const kind = S.KINDS[n.kind] || S.KINDS.Other;
      const shape = kind.shape;
      ctx.globalAlpha = dim ? 0.12 : 1;
      shapePath(ctx, shape === "ring" || shape === "double" || shape === "dot" ? "circle" : shape, n.x, n.y, r);
      if (shape === "ring" || shape === "squareo") {
        ctx.fillStyle = bg;
        ctx.fill();
        ctx.lineWidth = Math.max(r * 0.38, 1.2 / scale);
        ctx.strokeStyle = n._c;
        ctx.stroke();
      } else {
        ctx.fillStyle = n._c;
        ctx.fill();
        ctx.lineWidth = Math.min(1.2, 0.9 / Math.sqrt(scale) + 0.2);
        ctx.strokeStyle = bg;
        if (r * scale > 3) ctx.stroke();
      }
      if (shape === "double") {
        ctx.beginPath();
        ctx.arc(n.x, n.y, r * 1.28, 0, 2 * Math.PI);
        ctx.lineWidth = Math.max(r * 0.12, 0.8 / scale);
        ctx.strokeStyle = n._c;
        ctx.stroke();
      }
      if (n._pinned) {
        ctx.beginPath();
        ctx.arc(n.x + r * 0.8, n.y - r * 0.8, Math.max(1.5, r * 0.28), 0, 2 * Math.PI);
        ctx.fillStyle = dark ? "#e2e8f0" : "#334155";
        ctx.fill();
      }
      if (this.marked.has(n.id)) {
        ctx.beginPath();
        ctx.arc(n.x, n.y, r + 3.2 / scale + r * 0.25, 0, 2 * Math.PI);
        ctx.lineWidth = 2 / scale;
        ctx.strokeStyle = "#f59e0b";
        ctx.stroke();
      }
      if (this.selected === n.id) {
        ctx.beginPath();
        ctx.arc(n.x, n.y, r + 5 / scale + r * 0.3, 0, 2 * Math.PI);
        ctx.lineWidth = 2.4 / scale;
        ctx.strokeStyle = dark ? "#7cb0ff" : "#2563eb";
        ctx.stroke();
      }
      // labels
      const mode = this.style.labels;
      let show = false;
      if (mode === "all") show = true;
      else if (mode === "auto") {
        const screenR = r * scale;
        const grouped = n.kind === "Aggregate" || n.kind === "Class";
        show = screenR > (grouped ? 18 : 7) || (n._important && scale > 0.35) || (this.hl.active && this.hl.nodes.has(n.id) && this.hl.nodes.size < 60) || this.selected === n.id;
      } else if (mode === "selected") show = this.selected === n.id || (this.hl.active && this.hl.nodes.has(n.id) && this.hl.nodes.size < 80);
      if (show && !dim) {
        // label size in screen pixels: grows mildly with node size, capped for readability
        const fsPx = Math.min(15, Math.max(10.5, r * 0.55 * scale));
        const fs = fsPx / scale;
        ctx.font = `${n._important ? 600 : 500} ${fs}px Inter, system-ui, sans-serif`;
        ctx.textAlign = "center";
        ctx.textBaseline = "top";
        let label = n.label || "";
        if (label.length > 42) label = label.slice(0, 40) + "…";
        const y = n.y + r * (shape === "diamond" || shape === "triangle" ? 1.3 : 1.12) + 2 / scale;
        ctx.lineWidth = 3 / scale;
        ctx.strokeStyle = dark ? "rgba(12,17,27,0.9)" : "rgba(255,255,255,0.92)";
        ctx.strokeText(label, n.x, y);
        ctx.fillStyle = dark ? "#e6eaf2" : "#0f172a";
        ctx.fillText(label, n.x, y);
      }
      ctx.globalAlpha = 1;
    }

    // ── data ──
    setData(graph, opts = {}) {
      this.savePositions();
      this.graph = graph;
      this.nbrs = new Map();
      for (const n of graph.nodes) {
        this.nbrs.set(n.id, { nodes: new Set([n.id]), links: new Set() });
        const p = this.pos.get(n.id);
        if (p) {
          n.x = p.x;
          n.y = p.y;
          if (p.z != null) n.z = p.z;
        }
      }
      for (const l of graph.links) {
        const a = this.nbrs.get(l.source);
        const b = this.nbrs.get(l.target);
        if (a && b) {
          a.nodes.add(l.target);
          a.links.add(l);
          b.nodes.add(l.source);
          b.links.add(l);
        }
      }
      this.hl = { nodes: new Set(), links: new Set(), active: false };
      this._userMoved = false;
      this._fitOnStop = opts.fit !== false;
      this.applyForces();
      // pre-compute part of the layout so the first paint is already organised
      const n = graph.nodes.length;
      const warm = n <= 600 ? 160 : n <= 2000 ? 70 : n <= 5000 ? 25 : 0;
      const fixed = ["radial", "tree-td", "tree-lr", "circular"].includes(this.style.layout);
      (this.mode === "2d" ? this.fg : this.fg3)?.warmupTicks(fixed ? 0 : warm);
      if (this.mode === "2d") {
        this.fg.graphData(graph);
        this.applyLayout(this.style.layout, true);
        if (opts.fit !== false) setTimeout(() => this.fit(600), graph.nodes.length > 1500 ? 1400 : 700);
      } else if (this.fg3) {
        this.fg3.graphData(graph);
        this.applyLayout(this.style.layout, true);
        if (opts.fit !== false) setTimeout(() => this.fit(600), 1200);
      }
      this.paused = false;
    }

    savePositions() {
      for (const n of this.graph.nodes) if (n.x != null) this.pos.set(n.id, { x: n.x, y: n.y, z: n.z });
    }

    applyForces() {
      const n = this.graph.nodes.length;
      const fg = this.mode === "2d" ? this.fg : this.fg3;
      if (!fg) return;
      const k = n > 3000 ? 0.5 : n > 1000 ? 0.7 : n > 300 ? 0.9 : 1.15;
      fg.d3Force("charge").strength((d) => -(34 + (d._r || 4) * 9) * k).distanceMax(n > 2000 ? 320 : 900);
      fg.d3Force("link")
        .distance((l) => {
          const s = typeof l.source === "object" ? l.source : null;
          const t = typeof l.target === "object" ? l.target : null;
          const base = 22 + ((s?._r || 4) + (t?._r || 4)) * 1.8;
          return l.pred === "hasPropertySet" || l.pred === "hasValue" ? base * 1.4 : base;
        })
        .strength((l) => (l.weight > 1 ? 0.45 : 0.6));
      fg.d3Force("collide", d3.forceCollide((d) => (d._r || 4) * 1.15 + 2.5).strength(0.7).iterations(1));
    }

    clearConstraints() {
      for (const n of this.graph.nodes) {
        if (n._pinned) continue;
        n.fx = undefined;
        n.fy = undefined;
        n.fz = undefined;
      }
      const fg = this.mode === "2d" ? this.fg : this.fg3;
      fg.d3Force("x", null).d3Force("y", null).d3Force("z", null);
    }

    applyLayout(layout, silent) {
      this.style.layout = layout;
      const fg = this.mode === "2d" ? this.fg : this.fg3;
      if (!fg) return;
      this.clearConstraints();
      const nodes = this.graph.nodes;
      const n = nodes.length;
      fg.dagMode(null);
      if (layout === "radial" || layout === "tree-td" || layout === "tree-lr") this.treeLayout(layout);
      if (layout === "storeys") {
        const gap = Math.max(110, Math.min(260, Math.sqrt(n) * 7));
        const top = (this.hooks.storeyCount() || 1) + 1;
        for (const d of nodes) {
          const idx = this.hooks.storeyIndexOf(d);
          let level;
          if (idx != null) level = idx;
          else if (d.kind === "Project" || (d.kind === "Spatial" && !d.storey)) level = top;
          else level = -1.6;
          d.fy = -level * gap;
        }
        fg.d3Force("x", d3.forceX(0).strength(0.03));
      } else if (layout === "clusters") {
        const groups = new Map();
        for (const d of nodes) {
          const g = this.hooks.groupOf(d) ?? "—";
          if (!groups.has(g)) groups.set(g, groups.size);
        }
        const R = Math.max(200, Math.sqrt(n) * 22);
        const G = groups.size;
        const center = (d) => {
          const i = groups.get(this.hooks.groupOf(d) ?? "—");
          const a = (2 * Math.PI * i) / G;
          return [Math.cos(a) * R, Math.sin(a) * R];
        };
        fg.d3Force("x", d3.forceX((d) => center(d)[0]).strength(0.18));
        fg.d3Force("y", d3.forceY((d) => center(d)[1]).strength(0.18));
      } else if (layout === "force") {
        // weak gravity keeps disconnected components on screen
        const g = n > 2000 ? 0.03 : 0.06;
        fg.d3Force("x", d3.forceX(0).strength(g));
        fg.d3Force("y", d3.forceY(0).strength(g));
      } else if (layout === "circular") {
        const sorted = nodes.slice().sort((a, b) => String(this.hooks.groupOf(a)).localeCompare(String(this.hooks.groupOf(b))) || (a.label || "").localeCompare(b.label || ""));
        const R = Math.max(160, (n * 7) / (2 * Math.PI));
        sorted.forEach((d, i) => {
          const a = (2 * Math.PI * i) / n;
          d.fx = Math.cos(a) * R;
          d.fy = Math.sin(a) * R;
          if (this.mode === "3d") d.fz = 0;
        });
      }
      // 3d-force-graph builds its layout in a debounced update; reheating before that crashes
      if (this.mode === "3d") setTimeout(() => this.fg3 && this.fg3.d3ReheatSimulation(), 80);
      else fg.d3ReheatSimulation();
      const fixed = ["radial", "tree-td", "tree-lr", "circular"].includes(layout);
      if (!silent || fixed) setTimeout(() => this.fit(500), fixed ? 250 : 900);
    }

    /** Deterministic tidy-tree layout on a BFS spanning tree (works for any graph). */
    treeLayout(layout) {
      const nodes = this.graph.nodes;
      if (!nodes.length) return;
      const byId = new Map(nodes.map((n) => [n.id, n]));
      const adj = new Map(nodes.map((n) => [n.id, []]));
      const indeg = new Map(nodes.map((n) => [n.id, 0]));
      for (const l of this.graph.links) {
        const a = idOf(l.source), b = idOf(l.target);
        if (!adj.has(a) || !adj.has(b) || a === b) continue;
        adj.get(a).push(b);
        adj.get(b).push(a);
        if (l.pred !== "isTypedBy" && l.pred !== "hasMaterial" && l.pred !== "hasClassification" && l.pred !== "memberOf") indeg.set(b, indeg.get(b) + 1);
      }
      const rank = (n) => (n.kind === "Project" ? 3e6 : n.kind === "Exterior" ? 2e6 : 0) + (indeg.get(n.id) === 0 ? 1e6 : 0) + adj.get(n.id).length;
      const order = nodes.slice().sort((a, b) => rank(b) - rank(a));
      const children = new Map();
      const seen = new Set();
      const roots = [];
      for (const r of order) {
        if (seen.has(r.id)) continue;
        roots.push(r.id);
        seen.add(r.id);
        const q = [r.id];
        while (q.length) {
          const v = q.shift();
          const kids = [];
          // visit hubs first so subtrees stay compact
          for (const w of adj.get(v).slice().sort((x, y) => adj.get(y).length - adj.get(x).length)) {
            if (seen.has(w)) continue;
            seen.add(w);
            kids.push(w);
            q.push(w);
          }
          children.set(v, kids);
        }
      }
      const build = (id) => ({ id, children: (children.get(id) || []).map(build) });
      const data = { id: "__root", children: roots.map(build) };
      const h = d3.hierarchy(data);
      const leaves = h.leaves().length;
      const depth = h.height;
      if (layout === "radial") {
        const R = Math.max(220, (leaves * 9) / (2 * Math.PI), depth * 90);
        d3.tree().size([2 * Math.PI, R]).separation((a, b) => (a.parent === b.parent ? 1 : 2) / Math.max(1, a.depth))(h);
        h.each((d) => {
          const n = byId.get(d.data.id);
          if (!n) return;
          const r = roots.length === 1 ? d.y - R / depth : d.y;
          n.fx = Math.cos(d.x - Math.PI / 2) * Math.max(0, r);
          n.fy = Math.sin(d.x - Math.PI / 2) * Math.max(0, r);
          if (this.mode === "3d") n.fz = 0;
        });
      } else {
        const breadth = 12;
        const rOf = (d) => byId.get(d.data.id)?._r || 4;
        const level = Math.max(110, Math.min(240, 1800 / Math.max(1, depth)));
        // separation scales with node size so large class / group nodes do not overlap
        d3.tree().nodeSize([breadth, level]).separation((a, b) => ((rOf(a) + rOf(b)) * 1.25 + (layout === "tree-td" ? 10 : 4)) / breadth * (a.parent === b.parent ? 1 : 1.3))(h);
        h.each((d) => {
          const n = byId.get(d.data.id);
          if (!n) return;
          const along = d.x, down = d.y - level;
          if (layout === "tree-td") { n.fx = along; n.fy = down; }
          else { n.fx = down; n.fy = along; }
          if (this.mode === "3d") n.fz = 0;
        });
      }
      for (const n of nodes) if (n.fx != null) { n.x = n.fx; n.y = n.fy; }
    }

    // ── 3D ──
    async setMode(mode) {
      if (mode === this.mode) return;
      this.savePositions();
      if (mode === "3d") {
        for (const src of LIBS_3D) await loadScript(src);
        this.fg.pauseAnimation();
        this.c2d.style.display = "none";
        if (!this.el3) {
          this.el3 = document.createElement("div");
          this.el3.style.cssText = "position:absolute;inset:0";
          this.el.appendChild(this.el3);
        }
        this.el3.style.display = "";
        this.mode = "3d";
        this.init3D();
        this.applyForces();
        this.applyLayout(this.style.layout, true);
        setTimeout(() => this.resize(), 120);
        setTimeout(() => this.fit(600), 1500);
      } else {
        if (this.fg3) {
          this.fg3._destructor();
          this.fg3 = null;
        }
        if (this.el3) {
          this.el3.remove();
          this.el3 = null;
        }
        this.c2d.style.display = "";
        this.mode = "2d";
        this.fg.resumeAnimation();
        this.fg.graphData(this.graph);
        this.applyForces();
        this.applyLayout(this.style.layout, true);
        setTimeout(() => this.fit(500), 800);
      }
      this.resize();
    }

    init3D() {
      const self = this;
      const dark = this.theme() === "dark";
      this.fg3 = ForceGraph3D({ rendererConfig: { preserveDrawingBuffer: true, antialias: true } })(this.el3)
        .backgroundColor(dark ? "#0c111b" : "#fbfcfe")
        .width(this.el.clientWidth)
        .height(this.el.clientHeight)
        .nodeId("id")
        .nodeLabel(() => "")
        .nodeVal((n) => Math.max(1, (n._r || 4) ** 2 / 6))
        .nodeRelSize(2.2)
        .nodeResolution(12)
        .nodeOpacity(0.95)
        .nodeColor((n) => (self.hl.active && !self.hl.nodes.has(n.id) ? rgba(n._c, 0.12) : n.id === self.selected ? "#f43f5e" : n._c))
        .nodeThreeObjectExtend(true)
        .nodeThreeObject((n) => {
          if (self.style.labels === "none") return null;
          if (self.style.labels === "auto" && !n._important) return null;
          if (typeof SpriteText === "undefined") return null;
          const sp = new SpriteText(n.label.length > 36 ? n.label.slice(0, 34) + "…" : n.label);
          sp.color = dark ? "#e6eaf2" : "#0f172a";
          sp.textHeight = 3.2;
          sp.position.y = (n._r || 4) * 1.4 + 3;
          return sp;
        })
        .linkColor((l) => self.linkColor(l))
        .linkOpacity(0.55)
        .linkWidth((l) => (self.hl.active && self.hl.links.has(l) ? 1.2 : 0))
        .linkCurvature((l) => l.curvature || 0)
        .linkDirectionalArrowLength((l) => (self.style.arrows ? 3 : 0))
        .linkDirectionalArrowRelPos(1)
        .linkDirectionalParticles((l) => (self.style.particles && (l.pred === "feeds" || l.pred === "portConnectedTo") ? 2 : 0))
        .linkDirectionalParticleWidth(1.6)
        .onNodeClick((n, ev) => self.hooks.onClick(n, ev))
        .onNodeRightClick((n, ev) => self.hooks.onRightClick(n, ev))
        .onNodeHover((n) => {
          self.el.style.cursor = n ? "pointer" : "";
          self.hover(n);
          self.hooks.onHover(n);
        })
        .onBackgroundClick(() => self.hooks.onBackground())
        .onDagError(() => {})
        .cooldownTime(15000)
        .onEngineStop(() => {
          self.savePositions();
          if (self._fitOnStop && !self._userMoved) self.fit(700);
          self._fitOnStop = false;
        })
        .graphData(this.graph);
    }

    refresh3DStyle() {
      if (!this.fg3) return;
      this.fg3.nodeColor(this.fg3.nodeColor()).linkColor(this.fg3.linkColor()).linkWidth(this.fg3.linkWidth()).nodeThreeObject(this.fg3.nodeThreeObject());
    }

    // ── interaction ──
    hover(n) {
      if (this.lockHL) return;
      this.setHighlight(n ? n.id : this.selected);
    }

    setHighlight(id) {
      if (id && this.nbrs.has(id)) {
        const nb = this.nbrs.get(id);
        this.hl = { nodes: nb.nodes, links: nb.links, active: true };
      } else {
        this.hl = { nodes: new Set(), links: new Set(), active: false };
      }
      this.redraw();
    }

    highlightSet(ids, links) {
      const set = new Set(ids);
      this.hl = { nodes: set, links: new Set(links || []), active: set.size > 0 };
      if (!links) for (const l of this.graph.links) if (this.hl.nodes.has(idOf(l.source)) && this.hl.nodes.has(idOf(l.target))) this.hl.links.add(l);
      this.lockHL = this.hl.active;
      this.redraw();
    }

    clearHighlight() {
      this.lockHL = false;
      this.setHighlight(this.selected);
    }

    select(id) {
      this.selected = id;
      if (!this.lockHL) this.setHighlight(id);
      else this.redraw();
    }

    redraw() {
      if (this.mode === "2d" && this.fg) {
        // nudge the canvas to repaint without re-heating the simulation
        const z = this.fg.zoom();
        this._nudge = !this._nudge;
        this.fg.zoom(z * (this._nudge ? 1.0000001 : 0.9999999));
      } else this.refresh3DStyle();
    }

    restyle() {
      if (this.mode === "3d") this.refresh3DStyle();
      else this.redraw();
    }

    nodeById(id) {
      return this.graph.nodes.find((n) => n.id === id);
    }

    focus(id, ms = 700) {
      const n = this.nodeById(id);
      if (!n) return false;
      this._userMoved = true;
      if (this.mode === "2d") {
        this.fg.centerAt(n.x, n.y, ms);
        const k = this.fg.zoom();
        if (k < 1.6) this.fg.zoom(2.4, ms);
      } else if (this.fg3) {
        const d = 120;
        const ratio = 1 + d / Math.hypot(n.x || 1, n.y || 1, n.z || 1);
        this.fg3.cameraPosition({ x: (n.x || 1) * ratio, y: (n.y || 1) * ratio, z: (n.z || 1) * ratio }, n, ms);
      }
      return true;
    }

    fit(ms = 500) {
      if (this.mode === "2d") {
        const bb = this.fg.getGraphBbox();
        if (!bb || !Number.isFinite(bb.x[0])) return;
        const pad = 60;
        const w = Math.max(1, bb.x[1] - bb.x[0]);
        const h = Math.max(1, bb.y[1] - bb.y[0]);
        const k = Math.min(3, (this.fg.width() - pad * 2) / w, (this.fg.height() - pad * 2) / h);
        this.fg.centerAt((bb.x[0] + bb.x[1]) / 2, (bb.y[0] + bb.y[1]) / 2, ms);
        this.fg.zoom(Math.max(0.02, k), ms);
      } else if (this.fg3) this.fg3.zoomToFit(ms, 40);
    }
    zoomBy(f) {
      if (this.mode === "2d") this.fg.zoom(this.fg.zoom() * f, 250);
      else if (this.fg3) {
        const p = this.fg3.cameraPosition();
        this.fg3.cameraPosition({ x: p.x / f, y: p.y / f, z: p.z / f }, undefined, 250);
      }
    }
    togglePause() {
      const fg = this.mode === "2d" ? this.fg : this.fg3;
      this.paused = !this.paused;
      if (this.paused) {
        for (const n of this.graph.nodes) { n.fx = n.x; n.fy = n.y; if (this.mode === "3d") n.fz = n.z; }
      } else {
        this.applyLayout(this.style.layout, true);
      }
      fg.d3ReheatSimulation();
      return this.paused;
    }
    reheat() {
      const fg = this.mode === "2d" ? this.fg : this.fg3;
      for (const n of this.graph.nodes) if (n._pinned) { n._pinned = false; n.fx = undefined; n.fy = undefined; }
      this.pos.clear();
      this.applyLayout(this.style.layout);
      fg.d3ReheatSimulation();
    }

    screenshotPNG() {
      const canvas = this.mode === "2d" ? this.el.querySelector("canvas") : this.fg3.renderer().domElement;
      const dark = this.theme() === "dark";
      const out = document.createElement("canvas");
      out.width = canvas.width;
      out.height = canvas.height;
      const ctx = out.getContext("2d");
      ctx.fillStyle = dark ? "#0c111b" : "#ffffff";
      ctx.fillRect(0, 0, out.width, out.height);
      ctx.drawImage(canvas, 0, 0);
      return new Promise((res) => out.toBlob(res, "image/png"));
    }

    /** Publication-quality vector export of the current 2D layout. */
    toSVG(title, legendItems) {
      const nodes = this.graph.nodes.filter((n) => n.x != null);
      if (!nodes.length) return "";
      const dark = this.theme() === "dark";
      let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
      for (const n of nodes) {
        const r = (n._r || 4) * 1.3;
        minX = Math.min(minX, n.x - r);
        minY = Math.min(minY, n.y - r);
        maxX = Math.max(maxX, n.x + r);
        maxY = Math.max(maxY, n.y + r + 14);
      }
      const pad = 40;
      const W = maxX - minX + pad * 2;
      const H = maxY - minY + pad * 2 + 30;
      const ox = -minX + pad;
      const oy = -minY + pad + 30;
      const f = (v) => +v.toFixed(2);
      const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;");
      const out = [];
      out.push(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${f(W)} ${f(H)}" width="${f(W)}" height="${f(H)}" font-family="Inter, Helvetica, Arial, sans-serif">`);
      out.push(`<rect width="100%" height="100%" fill="${dark ? "#0c111b" : "#ffffff"}"/>`);
      out.push(`<text x="${pad}" y="28" font-size="16" font-weight="600" fill="${dark ? "#e6eaf2" : "#0f172a"}">${esc(title || "")}</text>`);
      out.push(`<g transform="translate(${f(ox)},${f(oy)})">`);
      out.push(`<g fill="none" stroke-linecap="round">`);
      for (const l of this.graph.links) {
        const s = l.source, t = l.target;
        if (typeof s !== "object" || typeof t !== "object" || s.x == null || t.x == null) continue;
        const col = l.onPath ? "#f43f5e" : l._c || "#94a3b8";
        const w = this.linkWidth(l);
        const dash = l.inferred ? ` stroke-dasharray="3 2"` : "";
        if (l.curvature) {
          const dx = t.x - s.x, dy = t.y - s.y;
          const cx = (s.x + t.x) / 2 + dy * l.curvature;
          const cy = (s.y + t.y) / 2 - dx * l.curvature;
          out.push(`<path d="M${f(s.x)},${f(s.y)} Q${f(cx)},${f(cy)} ${f(t.x)},${f(t.y)}" stroke="${col}" stroke-opacity="0.55" stroke-width="${f(w)}"${dash}/>`);
        } else {
          out.push(`<line x1="${f(s.x)}" y1="${f(s.y)}" x2="${f(t.x)}" y2="${f(t.y)}" stroke="${col}" stroke-opacity="0.55" stroke-width="${f(w)}"${dash}/>`);
        }
      }
      out.push(`</g><g>`);
      const bg = dark ? "#0c111b" : "#ffffff";
      for (const n of nodes) {
        const shape = (S.KINDS[n.kind] || S.KINDS.Other).shape;
        const r = n._r || 4;
        const hollow = shape === "ring" || shape === "squareo";
        out.push(`${shapeSVG(shape === "ring" || shape === "double" || shape === "dot" ? "circle" : shape, n.x, n.y, r)} fill="${hollow ? bg : n._c}" stroke="${hollow ? n._c : bg}" stroke-width="${hollow ? f(r * 0.38) : 0.8}"><title>${esc(n.label)}${n.cls ? " · " + esc(n.cls) : ""}</title></${shape === "hexagon" || shape === "diamond" || shape === "triangle" || shape === "tag" ? "polygon" : shape === "square" || shape === "squareo" || shape === "roundrect" ? "rect" : "circle"}>`);
        if (shape === "double") out.push(`<circle cx="${f(n.x)}" cy="${f(n.y)}" r="${f(r * 1.28)}" fill="none" stroke="${n._c}" stroke-width="${f(r * 0.12)}"/>`);
      }
      out.push(`</g><g font-size="7" fill="${dark ? "#e6eaf2" : "#0f172a"}" text-anchor="middle" paint-order="stroke" stroke="${bg}" stroke-width="2">`);
      const labelAll = this.style.labels === "all" || nodes.length <= 400;
      for (const n of nodes) {
        if (this.style.labels === "none") break;
        if (!labelAll && !n._important) continue;
        const r = n._r || 4;
        const lab = (n.label || "").length > 42 ? n.label.slice(0, 40) + "…" : n.label || "";
        out.push(`<text x="${f(n.x)}" y="${f(n.y + r * 1.15 + 7)}" font-size="${f(Math.max(6, r * 0.6))}"${n._important ? ' font-weight="600"' : ""}>${esc(lab)}</text>`);
      }
      out.push(`</g></g>`);
      if (legendItems?.length) {
        out.push(`<g transform="translate(${f(W - 200)},${pad})" font-size="10" fill="${dark ? "#a9b4c6" : "#475569"}">`);
        legendItems.slice(0, 24).forEach((it, i) => {
          out.push(`<circle cx="6" cy="${i * 15 + 6}" r="5" fill="${it.color}"/><text x="16" y="${i * 15 + 10}">${esc(it.label)}</text>`);
        });
        out.push(`</g>`);
      }
      out.push(`</svg>`);
      return out.join("\n");
    }
  }

  KG.Renderer = { Renderer, rgba, shapeSVG };
})((window.KG = window.KG || {}));
