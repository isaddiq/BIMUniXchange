/* BIMUniXchange · BIM Knowledge Graph Studio
 * analytics.js — network science on the visible graph: degree, PageRank,
 * betweenness (Brandes, sampled for large graphs), harmonic closeness,
 * clustering, connected components, Louvain communities, shortest paths,
 * and space-syntax integration for space-access graphs.
 */
(function (KG) {
  "use strict";

  const idOf = (x) => (typeof x === "object" && x !== null ? x.id : x);

  /** Compressed sparse row adjacency (undirected, de-duplicated) + directed lists. */
  function index(graph) {
    const ids = graph.nodes.map((n) => n.id);
    const idx = new Map(ids.map((id, i) => [id, i]));
    const n = ids.length;
    const und = Array.from({ length: n }, () => new Set());
    const outL = Array.from({ length: n }, () => []);
    const inDeg = new Int32Array(n);
    let m = 0;
    for (const l of graph.links) {
      const a = idx.get(idOf(l.source));
      const b = idx.get(idOf(l.target));
      if (a === undefined || b === undefined || a === b) continue;
      und[a].add(b);
      und[b].add(a);
      outL[a].push(b);
      inDeg[b]++;
      m++;
    }
    const off = new Int32Array(n + 1);
    for (let i = 0; i < n; i++) off[i + 1] = off[i] + und[i].size;
    const adj = new Int32Array(off[n]);
    for (let i = 0; i < n; i++) {
      let k = off[i];
      for (const j of und[i]) adj[k++] = j;
    }
    return { ids, idx, n, m, off, adj, outL, inDeg };
  }

  function degrees(G) {
    const out = new Map();
    for (let i = 0; i < G.n; i++) out.set(G.ids[i], { deg: G.off[i + 1] - G.off[i], outDeg: G.outL[i].length, inDeg: G.inDeg[i] });
    return out;
  }

  function components(G) {
    const comp = new Int32Array(G.n).fill(-1);
    const sizes = [];
    let c = 0;
    const q = new Int32Array(G.n);
    for (let s = 0; s < G.n; s++) {
      if (comp[s] >= 0) continue;
      let h = 0, t = 0;
      q[t++] = s;
      comp[s] = c;
      while (h < t) {
        const v = q[h++];
        for (let k = G.off[v]; k < G.off[v + 1]; k++) {
          const w = G.adj[k];
          if (comp[w] < 0) { comp[w] = c; q[t++] = w; }
        }
      }
      sizes.push(t);
      c++;
    }
    return { comp, sizes, count: c };
  }

  function pagerank(G, d = 0.85, iters = 50) {
    const n = G.n;
    if (!n) return new Float64Array(0);
    let pr = new Float64Array(n).fill(1 / n);
    for (let it = 0; it < iters; it++) {
      const next = new Float64Array(n).fill((1 - d) / n);
      let dangling = 0;
      for (let i = 0; i < n; i++) {
        const o = G.outL[i];
        if (!o.length) { dangling += pr[i]; continue; }
        const share = (d * pr[i]) / o.length;
        for (const j of o) next[j] += share;
      }
      const dd = (d * dangling) / n;
      let diff = 0;
      for (let i = 0; i < n; i++) {
        next[i] += dd;
        diff += Math.abs(next[i] - pr[i]);
      }
      pr = next;
      if (diff < 1e-9) break;
    }
    return pr;
  }

  /** Brandes betweenness (undirected). Samples `k` pivots when n is large. */
  function betweenness(G, maxExact = 2500, k = 400) {
    const n = G.n;
    const cb = new Float64Array(n);
    if (n < 3) return { values: cb, sampled: false };
    let sources = [...Array(n).keys()];
    let sampled = false;
    if (n > maxExact) {
      sampled = true;
      for (let i = n - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [sources[i], sources[j]] = [sources[j], sources[i]];
      }
      sources = sources.slice(0, k);
    }
    const sigma = new Float64Array(n);
    const dist = new Int32Array(n);
    const delta = new Float64Array(n);
    const stack = new Int32Array(n);
    const queue = new Int32Array(n);
    const predStart = G.off;
    const pred = new Int32Array(G.adj.length);
    const predCount = new Int32Array(n);
    for (const s of sources) {
      sigma.fill(0);
      dist.fill(-1);
      delta.fill(0);
      predCount.fill(0);
      sigma[s] = 1;
      dist[s] = 0;
      let qh = 0, qt = 0, sp = 0;
      queue[qt++] = s;
      while (qh < qt) {
        const v = queue[qh++];
        stack[sp++] = v;
        for (let e = G.off[v]; e < G.off[v + 1]; e++) {
          const w = G.adj[e];
          if (dist[w] < 0) { dist[w] = dist[v] + 1; queue[qt++] = w; }
          if (dist[w] === dist[v] + 1) {
            sigma[w] += sigma[v];
            pred[predStart[w] + predCount[w]++] = v;
          }
        }
      }
      while (sp > 0) {
        const w = stack[--sp];
        for (let p = 0; p < predCount[w]; p++) {
          const v = pred[predStart[w] + p];
          delta[v] += (sigma[v] / sigma[w]) * (1 + delta[w]);
        }
        if (w !== s) cb[w] += delta[w];
      }
    }
    const scale = (sampled ? n / sources.length : 1) / 2;
    const norm = ((n - 1) * (n - 2)) / 2;
    for (let i = 0; i < n; i++) cb[i] = (cb[i] * scale) / norm;
    return { values: cb, sampled };
  }

  /** Harmonic closeness centrality (exact for small graphs, pivot-sampled otherwise). */
  function harmonic(G, maxWork = 4e7, k = 300) {
    const n = G.n;
    const h = new Float64Array(n);
    if (n < 2) return { values: h, sampled: false };
    const exact = n * (G.adj.length + n) <= maxWork;
    let sources = [...Array(n).keys()];
    if (!exact) sources = sources.sort(() => Math.random() - 0.5).slice(0, k);
    const dist = new Int32Array(n);
    const q = new Int32Array(n);
    for (const s of sources) {
      dist.fill(-1);
      dist[s] = 0;
      let qh = 0, qt = 0;
      q[qt++] = s;
      while (qh < qt) {
        const v = q[qh++];
        for (let e = G.off[v]; e < G.off[v + 1]; e++) {
          const w = G.adj[e];
          if (dist[w] < 0) { dist[w] = dist[v] + 1; q[qt++] = w; h[w] += 1 / dist[w]; }
        }
      }
    }
    const scale = exact ? 1 : n / sources.length;
    for (let i = 0; i < n; i++) h[i] = (h[i] * scale) / (n - 1);
    return { values: h, sampled: !exact };
  }

  function clustering(G, maxDeg = 400) {
    const n = G.n;
    const cc = new Float64Array(n);
    const mark = new Int32Array(n).fill(-1);
    let sum = 0, cnt = 0;
    for (let v = 0; v < n; v++) {
      const d = G.off[v + 1] - G.off[v];
      if (d < 2 || d > maxDeg) continue;
      for (let e = G.off[v]; e < G.off[v + 1]; e++) mark[G.adj[e]] = v;
      let tri = 0;
      for (let e = G.off[v]; e < G.off[v + 1]; e++) {
        const u = G.adj[e];
        for (let f = G.off[u]; f < G.off[u + 1]; f++) if (mark[G.adj[f]] === v) tri++;
      }
      cc[v] = tri / (d * (d - 1));
      sum += cc[v];
      cnt++;
    }
    return { values: cc, average: n ? sum / n : 0, averageNonTrivial: cnt ? sum / cnt : 0 };
  }

  /** Louvain community detection (modularity optimisation, multi-level). */
  function louvain(G) {
    const n0 = G.n;
    let nbrs = [];
    for (let i = 0; i < n0; i++) {
      const mp = new Map();
      for (let e = G.off[i]; e < G.off[i + 1]; e++) mp.set(G.adj[e], 1);
      nbrs.push(mp);
    }
    const membership = Int32Array.from({ length: n0 }, (_, i) => i);
    for (let level = 0; level < 12; level++) {
      const res = oneLevel(nbrs);
      if (!res.moved) break;
      const remap = new Map();
      for (const c of res.comm) if (!remap.has(c)) remap.set(c, remap.size);
      const comm = res.comm.map((c) => remap.get(c));
      for (let i = 0; i < n0; i++) membership[i] = comm[membership[i]];
      const k = remap.size;
      const agg = Array.from({ length: k }, () => new Map());
      nbrs.forEach((mp, i) => {
        const ci = comm[i];
        for (const [j, w] of mp) {
          const cj = comm[j];
          agg[ci].set(cj, (agg[ci].get(cj) || 0) + w);
        }
      });
      const shrunk = k < nbrs.length;
      nbrs = agg;
      if (!shrunk) break;
    }
    // Relabel by community size (largest = 0)
    const sizes = new Map();
    for (const c of membership) sizes.set(c, (sizes.get(c) || 0) + 1);
    const order = [...sizes.entries()].sort((a, b) => b[1] - a[1]).map(([c]) => c);
    const rank = new Map(order.map((c, i) => [c, i]));
    for (let i = 0; i < n0; i++) membership[i] = rank.get(membership[i]);
    return { membership, count: order.length, modularity: modularity(G, membership) };
  }

  function oneLevel(nbrs) {
    const n = nbrs.length;
    const k = new Float64Array(n);
    let m2 = 0;
    for (let i = 0; i < n; i++) {
      let s = 0;
      for (const w of nbrs[i].values()) s += w;
      k[i] = s;
      m2 += s;
    }
    const comm = Array.from({ length: n }, (_, i) => i);
    if (!m2) return { comm, moved: false };
    const tot = Float64Array.from(k);
    const order = [...Array(n).keys()].sort(() => Math.random() - 0.5);
    let moved = false;
    for (let pass = 0; pass < 25; pass++) {
      let changed = false;
      for (const i of order) {
        const ci = comm[i];
        const wc = new Map();
        for (const [j, w] of nbrs[i]) {
          if (j === i) continue;
          const cj = comm[j];
          wc.set(cj, (wc.get(cj) || 0) + w);
        }
        tot[ci] -= k[i];
        let best = ci;
        let bestGain = (wc.get(ci) || 0) - (tot[ci] * k[i]) / m2;
        for (const [c, w] of wc) {
          const gain = w - (tot[c] * k[i]) / m2;
          if (gain > bestGain + 1e-12) { bestGain = gain; best = c; }
        }
        tot[best] += k[i];
        if (best !== ci) { comm[i] = best; moved = true; changed = true; }
      }
      if (!changed) break;
    }
    return { comm, moved };
  }

  function modularity(G, membership) {
    const m2 = G.adj.length;
    if (!m2) return 0;
    const tot = new Map();
    let inside = 0;
    for (let v = 0; v < G.n; v++) {
      const d = G.off[v + 1] - G.off[v];
      tot.set(membership[v], (tot.get(membership[v]) || 0) + d);
      for (let e = G.off[v]; e < G.off[v + 1]; e++) if (membership[G.adj[e]] === membership[v]) inside++;
    }
    let q = inside / m2;
    for (const t of tot.values()) q -= (t / m2) ** 2;
    return q;
  }

  function bfs(G, s) {
    const dist = new Int32Array(G.n).fill(-1);
    const prev = new Int32Array(G.n).fill(-1);
    const q = new Int32Array(G.n);
    let qh = 0, qt = 0;
    dist[s] = 0;
    q[qt++] = s;
    while (qh < qt) {
      const v = q[qh++];
      for (let e = G.off[v]; e < G.off[v + 1]; e++) {
        const w = G.adj[e];
        if (dist[w] < 0) { dist[w] = dist[v] + 1; prev[w] = v; q[qt++] = w; }
      }
    }
    return { dist, prev, last: q[qt - 1], reached: qt };
  }

  function shortestPath(G, a, b) {
    const s = G.idx.get(a);
    const t = G.idx.get(b);
    if (s === undefined || t === undefined) return null;
    const { dist, prev } = bfs(G, s);
    if (dist[t] < 0) return null;
    const path = [];
    for (let v = t; v >= 0; v = prev[v]) path.push(G.ids[v]);
    return path.reverse();
  }

  function diameterEstimate(G, comps) {
    if (!G.n) return 0;
    let big = 0;
    comps.sizes.forEach((s, i) => { if (s > comps.sizes[big]) big = i; });
    const start = comps.comp.findIndex((c) => c === big);
    const a = bfs(G, start);
    const b = bfs(G, a.last);
    return b.dist[b.last];
  }

  /** Space syntax measures on an access graph: mean depth, RA, integration. */
  function spaceSyntax(G, rootId) {
    const res = new Map();
    for (let v = 0; v < G.n; v++) {
      const { dist, reached } = bfs(G, v);
      if (reached < 3) { res.set(G.ids[v], { meanDepth: null, ra: null, integration: null }); continue; }
      let sum = 0;
      for (let i = 0; i < G.n; i++) if (dist[i] > 0) sum += dist[i];
      const md = sum / (reached - 1);
      const ra = (2 * (md - 1)) / (reached - 2);
      res.set(G.ids[v], { meanDepth: md, ra, integration: ra > 0 ? 1 / ra : null });
    }
    if (rootId != null && G.idx.has(rootId)) {
      const { dist } = bfs(G, G.idx.get(rootId));
      for (let i = 0; i < G.n; i++) res.get(G.ids[i]).depth = dist[i] >= 0 ? dist[i] : null;
    }
    return res;
  }

  /** Full metric suite for the visible graph. */
  function analyze(graph, opts = {}) {
    const t0 = performance.now();
    const G = index(graph);
    const deg = degrees(G);
    const comps = components(G);
    const pr = pagerank(G);
    const cl = clustering(G);
    const result = {
      G,
      n: G.n,
      m: G.m,
      density: G.n > 1 ? G.adj.length / 2 / ((G.n * (G.n - 1)) / 2) : 0,
      avgDegree: G.n ? G.adj.length / G.n : 0,
      components: comps.count,
      largestComponent: comps.sizes.length ? Math.max(...comps.sizes) : 0,
      isolates: [...deg.values()].filter((d) => d.deg === 0).length,
      clustering: cl.average,
      diameter: diameterEstimate(G, comps),
      metrics: new Map(),
    };
    let bc = null;
    let hc = null;
    if (opts.full !== false) {
      bc = betweenness(G);
      hc = harmonic(G);
      result.betweennessSampled = bc.sampled;
      result.closenessSampled = hc.sampled;
    }
    let lv = null;
    if (opts.communities !== false) {
      lv = louvain(G);
      result.communities = lv.count;
      result.modularity = lv.modularity;
    }
    for (let i = 0; i < G.n; i++) {
      const d = deg.get(G.ids[i]);
      result.metrics.set(G.ids[i], {
        degree: d.deg,
        inDegree: d.inDeg,
        outDegree: d.outDeg,
        pagerank: pr[i],
        betweenness: bc ? bc.values[i] : null,
        closeness: hc ? hc.values[i] : null,
        clustering: cl.values[i],
        component: comps.comp[i],
        community: lv ? lv.membership[i] : null,
      });
    }
    // Degree distribution
    const hist = new Map();
    for (const d of deg.values()) hist.set(d.deg, (hist.get(d.deg) || 0) + 1);
    result.degreeHistogram = [...hist.entries()].sort((a, b) => a[0] - b[0]);
    result.ms = performance.now() - t0;
    return result;
  }

  KG.Analytics = { index, degrees, components, pagerank, betweenness, harmonic, clustering, louvain, modularity, shortestPath, spaceSyntax, analyze, bfs };
})((window.KG = window.KG || {}));
