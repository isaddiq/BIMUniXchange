/* BIMUniXchange · BIM Knowledge Graph Studio
 * exporters.js — research & industry interchange formats:
 *   JSON node-link, GraphML (Gephi / yEd / Cytoscape / NetworkX),
 *   CSV node & edge tables, RDF Turtle (ifcOWL + BOT + Brick + LBD PROPS),
 *   Neo4j Cypher, model-health and semantic-mapping reports.
 */
(function (KG) {
  "use strict";
  const S = KG.Schema;

  function download(filename, content, mime = "text/plain") {
    const blob = content instanceof Blob ? content : new Blob([content], { type: mime + ";charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  const slug = (s) => (s || "model").replace(/\.[^.]+$/, "").replace(/[^A-Za-z0-9_-]+/g, "_").slice(0, 60) || "model";
  const xml = (s) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const csvCell = (v) => {
    if (v == null) return "";
    const s = String(v);
    return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const idOf = (x) => (typeof x === "object" && x ? x.id : x);

  function nodeRecord(n, model, metrics) {
    const mn = n.ref ? model.nodes.get(n.ref) : null;
    const m = metrics?.get(n.id) || {};
    return {
      id: n.id,
      label: n.label,
      kind: n.kind,
      ifcClass: n.cls || "",
      domain: n.domain || "",
      storey: n.storey ? model.nodes.get(n.storey)?.label || "" : "",
      globalId: mn?.guid || "",
      predefinedType: mn?.predefinedType || "",
      count: n.count ?? "",
      degree: m.degree ?? "",
      pagerank: m.pagerank ?? "",
      betweenness: m.betweenness ?? "",
      closeness: m.closeness ?? "",
      clustering: m.clustering ?? "",
      community: m.community ?? "",
      x: n.x != null ? +n.x.toFixed(2) : "",
      y: n.y != null ? +n.y.toFixed(2) : "",
    };
  }

  function linkRecord(l) {
    return {
      source: idOf(l.source),
      target: idOf(l.target),
      predicate: l.pred,
      category: l.cat || "",
      inferred: !!l.inferred,
      weight: l.weight || 1,
    };
  }

  function toJSON(graph, model, analysis, settings) {
    const metrics = analysis?.metrics;
    const out = {
      generator: "BIMUniXchange BIM Knowledge Graph Studio",
      exported: new Date().toISOString(),
      source: { file: model.meta.fileName, schema: model.meta.schema, application: model.meta.application || "", entities: model.meta.entityCount },
      view: { lens: settings?.lens, detail: graph.info?.detail, nodes: graph.nodes.length, links: graph.links.length },
      metrics: analysis ? { density: analysis.density, avgDegree: analysis.avgDegree, components: analysis.components, communities: analysis.communities, modularity: analysis.modularity, clustering: analysis.clustering, diameter: analysis.diameter } : null,
      nodes: graph.nodes.map((n) => {
        const r = nodeRecord(n, model, metrics);
        const mn = n.ref ? model.nodes.get(n.ref) : null;
        if (mn && (mn.kind === "Element" || mn.kind === "Spatial")) {
          r.properties = {};
          for (const ps of model.propertySets(mn)) {
            r.properties[ps.name] = Object.fromEntries(ps.props.map((p) => [p.name, p.value]));
          }
        }
        if (n.members) r.members = n.members;
        return r;
      }),
      links: graph.links.map(linkRecord),
    };
    return JSON.stringify(out, null, 1);
  }

  function toGraphML(graph, model, analysis) {
    const metrics = analysis?.metrics;
    const nkeys = ["label", "kind", "ifcClass", "domain", "storey", "globalId", "predefinedType", "count", "degree", "pagerank", "betweenness", "closeness", "clustering", "community", "x", "y"];
    const ntype = { count: "int", degree: "int", community: "int", pagerank: "double", betweenness: "double", closeness: "double", clustering: "double", x: "double", y: "double" };
    const ekeys = ["predicate", "category", "inferred", "weight"];
    const etype = { inferred: "boolean", weight: "double" };
    let s = '<?xml version="1.0" encoding="UTF-8"?>\n<graphml xmlns="http://graphml.graphdrawing.org/xmlns" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://graphml.graphdrawing.org/xmlns http://graphml.graphdrawing.org/xmlns/1.0/graphml.xsd">\n';
    s += `  <!-- BIMUniXchange BIM Knowledge Graph Studio · ${xml(model.meta.fileName)} · ${new Date().toISOString()} -->\n`;
    nkeys.forEach((k) => (s += `  <key id="n_${k}" for="node" attr.name="${k}" attr.type="${ntype[k] || "string"}"/>\n`));
    ekeys.forEach((k) => (s += `  <key id="e_${k}" for="edge" attr.name="${k}" attr.type="${etype[k] || "string"}"/>\n`));
    s += '  <graph id="G" edgedefault="directed">\n';
    for (const n of graph.nodes) {
      const r = nodeRecord(n, model, metrics);
      s += `    <node id="${xml(n.id)}">`;
      for (const k of nkeys) if (r[k] !== "" && r[k] != null) s += `<data key="n_${k}">${xml(r[k])}</data>`;
      s += "</node>\n";
    }
    graph.links.forEach((l, i) => {
      const r = linkRecord(l);
      s += `    <edge id="e${i}" source="${xml(r.source)}" target="${xml(r.target)}">`;
      for (const k of ekeys) s += `<data key="e_${k}">${xml(r[k])}</data>`;
      s += "</edge>\n";
    });
    s += "  </graph>\n</graphml>\n";
    return s;
  }

  function toCSV(rows) {
    if (!rows.length) return "";
    const cols = Object.keys(rows[0]);
    return [cols.join(","), ...rows.map((r) => cols.map((c) => csvCell(r[c])).join(","))].join("\n");
  }
  const nodesCSV = (graph, model, analysis) => toCSV(graph.nodes.map((n) => nodeRecord(n, model, analysis?.metrics)));
  const edgesCSV = (graph) => toCSV(graph.links.map(linkRecord));

  /** Expand view nodes (incl. collapsed groups) to semantic-model node ids. */
  function modelIds(graph) {
    const ids = new Set();
    for (const n of graph.nodes) {
      if (n.members) n.members.forEach((m) => ids.add(m));
      else if (n.ref) ids.add(n.ref);
    }
    return ids;
  }

  // ── RDF / Turtle ──
  const local = (id) => "n" + String(id).replace(/[^A-Za-z0-9_-]/g, (c) => "_" + c.charCodeAt(0).toString(16));
  const lit = (v) => {
    if (typeof v === "boolean") return v ? "true" : "false";
    if (typeof v === "number" && Number.isFinite(v)) return `"${v}"^^xsd:double`;
    return '"' + String(v).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\n/g, "\\n").replace(/\r/g, "\\r") + '"';
  };
  const camel = (s) => {
    const parts = String(s).replace(/[^A-Za-z0-9]+/g, " ").trim().split(/\s+/);
    const out = parts.map((p, i) => (i === 0 ? p.charAt(0).toLowerCase() + p.slice(1) : p.charAt(0).toUpperCase() + p.slice(1))).join("");
    return /^[A-Za-z_]/.test(out) ? out : "p" + out;
  };

  function toTurtle(model, ids) {
    const base = `urn:bimunixchange:${slug(model.meta.fileName)}:`;
    const lines = [];
    lines.push(`# Knowledge graph exported by BIMUniXchange BIM Knowledge Graph Studio`);
    lines.push(`# Source: ${model.meta.fileName} (${model.meta.schema || "?"}) · ${new Date().toISOString()}`);
    lines.push(`# Ontologies: ifcOWL (classes), BOT (topology), Brick (MEP), LBD PROPS level-1 (properties)`);
    for (const [p, iri] of Object.entries(S.NS)) lines.push(`@prefix ${p}: <${iri}> .`);
    lines.push(`@prefix props: <https://w3id.org/props#> .`);
    lines.push(`@prefix inst: <${base}> .`);
    lines.push("");
    const keep = (id) => !ids || ids.has(id);
    for (const n of model.nodes.values()) {
      if (!keep(n.id) || n.virtual) continue;
      const o = S.ontologyFor(n.cls, n.kind);
      const types = [n.cls && n.cls.startsWith("Ifc") ? o.ifcowl : null, o.bot, o.brick].filter(Boolean);
      const po = [];
      if (types.length) po.push(`a ${types.join(", ")}`);
      po.push(`rdfs:label ${lit(n.label)}`);
      if (n.guid) po.push(`bux:globalId ${lit(n.guid)}`);
      if (n.ifcId != null) po.push(`bux:stepId ${n.ifcId}`);
      po.push(`bux:kind ${lit(n.kind)}`);
      if (n.domain) po.push(`bux:discipline ${lit(n.domain)}`);
      if (n.predefinedType) po.push(`bux:predefinedType ${lit(n.predefinedType)}`);
      if (n.objectType) po.push(`bux:objectType ${lit(n.objectType)}`);
      if (n.tag) po.push(`bux:tag ${lit(n.tag)}`);
      if (n.elevation != null) po.push(`bux:elevation ${lit(n.elevation)}`);
      if (n.kind === "Element" || n.kind === "Spatial" || n.kind === "System" || n.kind === "Zone") {
        const seen = new Set();
        for (const ps of model.propertySets(n)) {
          for (const p of ps.props) {
            if (p.value == null || p.value === "") continue;
            const pk = camel(p.name);
            const key = pk + "|" + p.value;
            if (seen.has(key)) continue;
            seen.add(key);
            po.push(`props:${pk} ${lit(p.value)}`);
          }
        }
      }
      lines.push(`inst:${local(n.id)} ${po.join(" ;\n    ")} .`);
    }
    lines.push("");
    for (const e of model.edges) {
      if (!keep(e.s) || !keep(e.t)) continue;
      const p = S.PRED[e.pred] || S.PRED.related;
      let pred = p.rdf;
      if (e.pred === "aggregates") {
        const b = S.botDecomposition(model.nodes.get(e.s).cls, model.nodes.get(e.t).cls);
        if (b) lines.push(`inst:${local(e.s)} ${b} inst:${local(e.t)} .`);
      }
      lines.push(`inst:${local(e.s)} ${pred} inst:${local(e.t)} .`);
    }
    return lines.join("\n") + "\n";
  }

  // ── Neo4j Cypher ──
  const cyStr = (v) => '"' + String(v).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\n/g, "\\n") + '"';
  const cyVal = (v) => (typeof v === "number" && Number.isFinite(v) ? String(v) : typeof v === "boolean" ? String(v) : cyStr(v));
  const cyLabel = (s) => "`" + String(s).replace(/`/g, "") + "`";

  function toCypher(model, ids) {
    const keep = (id) => !ids || ids.has(id);
    const out = [];
    out.push(`// BIMUniXchange BIM Knowledge Graph Studio — ${model.meta.fileName}`);
    out.push(`// Run in Neo4j Browser or cypher-shell. Creates :Entity nodes and typed relationships.`);
    out.push(`CREATE CONSTRAINT bimunix_uid IF NOT EXISTS FOR (n:Entity) REQUIRE n.uid IS UNIQUE;`);
    const byLabel = new Map();
    for (const n of model.nodes.values()) {
      if (!keep(n.id) || n.virtual) continue;
      const labels = ["Entity", n.kind, n.cls && n.cls.startsWith("Ifc") ? n.cls : null].filter(Boolean).map(cyLabel).join(":");
      const props = { uid: n.id, name: n.name || n.label, kind: n.kind, ifcClass: n.cls };
      if (n.guid) props.globalId = n.guid;
      if (n.domain) props.discipline = n.domain;
      if (n.predefinedType) props.predefinedType = n.predefinedType;
      if (n.storey) props.storey = model.nodes.get(n.storey)?.label;
      if (n.elevation != null) props.elevation = n.elevation;
      if (n.kind === "Element" || n.kind === "Spatial") {
        for (const ps of model.propertySets(n)) for (const p of ps.props) if (p.value != null && p.value !== "" && typeof p.value !== "object") props[ps.name + "." + p.name] = p.value;
      }
      const map = "{" + Object.entries(props).filter(([, v]) => v != null && v !== "").map(([k, v]) => "`" + k.replace(/`/g, "") + "`: " + cyVal(v)).join(", ") + "}";
      let l = byLabel.get(labels);
      if (!l) byLabel.set(labels, (l = []));
      l.push(map);
    }
    for (const [labels, rows] of byLabel) {
      for (let i = 0; i < rows.length; i += 500) {
        out.push(`UNWIND [${rows.slice(i, i + 500).join(",\n  ")}] AS row\nCREATE (n:${labels}) SET n = row;`);
      }
    }
    const byPred = new Map();
    for (const e of model.edges) {
      if (!keep(e.s) || !keep(e.t)) continue;
      let l = byPred.get(e.pred);
      if (!l) byPred.set(e.pred, (l = []));
      l.push(`[${cyStr(e.s)}, ${cyStr(e.t)}]`);
    }
    for (const [pred, rows] of byPred) {
      const rel = pred.replace(/([a-z])([A-Z])/g, "$1_$2").toUpperCase();
      const inferred = !!S.PRED[pred]?.inferred;
      for (let i = 0; i < rows.length; i += 1000) {
        out.push(`UNWIND [${rows.slice(i, i + 1000).join(", ")}] AS r\nMATCH (a:Entity {uid: r[0]}), (b:Entity {uid: r[1]})\nCREATE (a)-[:${rel} {inferred: ${inferred}}]->(b);`);
      }
    }
    return out.join("\n\n") + "\n";
  }

  function healthCSV(health, model) {
    const rows = [];
    for (const c of health.checks) {
      if (!c.affected.length) rows.push({ check: c.title, severity: c.severity, group: c.group, nodeId: "", globalId: "", ifcClass: "", name: "", failed: c.failed, population: c.population });
      for (const id of c.affected) {
        const n = model.nodes.get(id);
        rows.push({ check: c.title, severity: c.severity, group: c.group, nodeId: id, globalId: n?.guid || "", ifcClass: n?.cls || "", name: n?.label || "", failed: c.failed, population: c.population });
      }
    }
    return toCSV(rows);
  }

  function mappingCSV(rows) {
    return toCSV(rows);
  }

  KG.Export = { download, slug, toJSON, toGraphML, nodesCSV, edgesCSV, toCSV, toTurtle, toCypher, healthCSV, mappingCSV, modelIds };
})((window.KG = window.KG || {}));
