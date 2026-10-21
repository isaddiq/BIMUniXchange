/* BIMUniXchange · BIM Knowledge Graph Studio
 * views.js — "lenses": each lens projects the semantic model onto a graph
 * that answers one kind of question (spatial breakdown, MEP flow, space
 * topology, class-level schema …). Lenses share filtering, abstraction
 * (collapse instances into class groups) and node-budget logic.
 */
(function (KG) {
  "use strict";
  const S = KG.Schema;

  const LENSES = [
    { id: "kg", title: "Knowledge graph", desc: "Every semantic entity and relationship, including inferred links.", layout: "force" },
    { id: "spatial", title: "Spatial hierarchy", desc: "Project → Site → Building → Storey → Space → elements (BOT-aligned).", layout: "radial" },
    { id: "breakdown", title: "Breakdown tree", desc: "Pivot elements into a hierarchy by any attributes or properties (e.g. Category → Family → Type).", layout: "radial" },
    { id: "systems", title: "MEP systems & flow", desc: "Distribution systems, members and port-derived flow (feeds).", layout: "force" },
    { id: "connectivity", title: "Physical topology", desc: "Element connections, hosting (wall → door/window), coverings, interferences.", layout: "force" },
    { id: "spaces", title: "Space adjacency", desc: "Space-to-space adjacency and door access from space boundaries, with space-syntax metrics.", layout: "force" },
    { id: "types", title: "Types & classification", desc: "IFC class taxonomy, type objects and classification-system crosswalk.", layout: "tree-lr" },
    { id: "materials", title: "Materials", desc: "Material usage across element classes or individual elements.", layout: "force" },
    { id: "properties", title: "Information structure", desc: "Which classes carry which property sets and properties (information requirements).", layout: "force" },
    { id: "schema", title: "Schema meta-graph", desc: "Class-level summary — relationships aggregated between IFC classes.", layout: "force" },
    { id: "subgraph", title: "Focused subgraph", desc: "Query result, path or neighbourhood exploration.", layout: "force", hidden: true },
  ];

  const ELEMENTISH = new Set(["Element", "Opening", "Port"]);

  function passes(n, s) {
    if (s.hidden && s.hidden.has(n.id)) return false;
    if (ELEMENTISH.has(n.kind)) {
      if (s.domains && !s.domains.has(n.domain)) return false;
      if (s.classes && !s.classes.has(n.cls)) return false;
      if (s.storeys && !(n.storey && s.storeys.has(n.storey))) return false;
    } else if (n.kind === "Spatial" && s.storeys) {
      if (n.cls === "IfcBuildingStorey" && !s.storeys.has(n.id)) return false;
      if (n.cls !== "IfcBuildingStorey" && n.storey && !s.storeys.has(n.storey)) return false;
    }
    return true;
  }

  function rn(n, extra) {
    return { id: n.id, ref: n.id, label: n.label, kind: n.kind, cls: n.cls, domain: n.domain, storey: n.storey, sub: n.sub, ...extra };
  }
  function rl(e, extra) {
    return { source: e.s, target: e.t, pred: e.pred, cat: e.cat, inferred: e.inferred, edgeId: e.id, ...extra };
  }
  const short = (cls) => (cls || "").replace(/^Ifc/, "");

  // ── Lens builders ─────────────────────────────────────────────────────

  function lensKG(model, s, notes) {
    const kinds = s.kinds;
    const include = new Set();
    for (const n of model.nodes.values()) if (kinds.has(n.kind) && passes(n, s)) include.add(n.id);
    let links = [];
    for (const e of model.edges) {
      if (include.has(e.s) && include.has(e.t) && (!s.cats || s.cats.has(e.cat))) links.push(rl(e));
    }
    const filtered = s.domains || s.classes || s.storeys;
    const deg = new Map();
    for (const l of links) {
      deg.set(l.source, (deg.get(l.source) || 0) + 1);
      deg.set(l.target, (deg.get(l.target) || 0) + 1);
    }
    const nodes = [];
    for (const id of include) {
      const n = model.nodes.get(id);
      if (filtered && !ELEMENTISH.has(n.kind) && n.kind !== "Project" && n.kind !== "Spatial" && !deg.get(id)) continue;
      nodes.push(rn(n));
    }
    if (kinds.has("PropertySet")) addPsetNodes(model, nodes, links, s);
    if (s.valueKey) addValueNodes(model, nodes, links, s.valueKey);
    return { nodes, links };
  }

  function addPsetNodes(model, nodes, links, s) {
    const added = new Map();
    for (const rnode of nodes.slice()) {
      if (rnode.kind !== "Element" && rnode.kind !== "Spatial") continue;
      const n = model.nodes.get(rnode.ref);
      for (const ps of model.propertySets(n)) {
        const id = "pset:" + ps.name;
        if (!added.has(id)) {
          const node = { id, label: ps.name, kind: ps.kind === "Q" ? "QuantitySet" : "PropertySet", cls: ps.kind === "Q" ? "IfcElementQuantity" : "IfcPropertySet", domain: "Semantic", virtual: true, count: 0, psetName: ps.name };
          added.set(id, node);
          nodes.push(node);
        }
        added.get(id).count++;
        links.push({ source: n.id, target: id, pred: "hasPropertySet", cat: "Properties", inferred: false, inheritedFromType: ps.inherited });
      }
    }
  }

  function addValueNodes(model, nodes, links, key) {
    const added = new Map();
    const prop = key.split(".").slice(1).join(".");
    for (const rnode of nodes.slice()) {
      if (rnode.kind !== "Element" && rnode.kind !== "Spatial") continue;
      const n = model.nodes.get(rnode.ref);
      const v = model.propValue(n, key);
      if (v == null || v === "") continue;
      const id = "val:" + key + "=" + v;
      if (!added.has(id)) {
        const node = { id, label: `${prop} = ${v}`, kind: "Value", cls: "Value", domain: "Semantic", virtual: true, count: 0, propKey: key, value: v };
        added.set(id, node);
        nodes.push(node);
      }
      added.get(id).count++;
      links.push({ source: n.id, target: id, pred: "hasValue", cat: "Properties", inferred: false });
    }
  }

  function lensSpatial(model, s, notes) {
    const include = new Set();
    for (const n of model.nodes.values()) {
      if ((n.kind === "Project" || n.kind === "Spatial") && passes(n, s)) include.add(n.id);
    }
    let orphan = 0;
    for (const n of model.nodes.values()) {
      if (n.kind !== "Element" && !(s.showOpenings && n.kind === "Opening") && !(s.showPorts && n.kind === "Port")) continue;
      if (!passes(n, s)) continue;
      const up = model.inE(n.id).some((e) => e.pred === "contains" || e.pred === "aggregates" || e.pred === "references" || e.pred === "hasOpening" || e.pred === "hasPort" || e.pred === "nests");
      if (up) include.add(n.id);
      else if (n.kind === "Element") orphan++;
    }
    if (orphan) notes.push(`${orphan.toLocaleString()} element${orphan === 1 ? " has" : "s have"} no spatial container and ${orphan === 1 ? "is" : "are"} not shown (see Model health).`);
    const preds = new Set(["aggregates", "contains", "nests", ...(s.showOpenings ? ["hasOpening"] : []), ...(s.showPorts ? ["hasPort"] : []), ...(s.showReferences ? ["references"] : [])]);
    const links = model.edges.filter((e) => preds.has(e.pred) && include.has(e.s) && include.has(e.t)).map((e) => rl(e));
    // Keep tree-like: one parent per child for contains/aggregates
    const seenChild = new Set();
    const tree = [];
    for (const l of links) {
      if (l.pred === "references") { tree.push(l); continue; }
      if (seenChild.has(l.target)) continue;
      seenChild.add(l.target);
      tree.push(l);
    }
    const nodes = [...include].map((id) => rn(model.nodes.get(id)));
    return { nodes, links: tree };
  }

  function lensBreakdown(model, s, notes) {
    const keys = (s.breakdown || []).filter(Boolean);
    const elements = [...model.nodes.values()].filter((n) => n.kind === "Element" && passes(n, s));
    const project = [...model.nodes.values()].find((n) => n.kind === "Project");
    const root = { id: "bd:root", label: project?.label || "Model", kind: "Project", cls: "", domain: "Spatial", count: elements.length, virtual: true };
    const nodes = [root];
    const links = [];
    const groups = new Map();
    const instances = s.includeInstances || elements.length <= 400;
    for (const n of elements) {
      let parent = root.id;
      let path = "";
      keys.forEach((k, i) => {
        let v = model.valueOf(n, k);
        if (v == null || v === "") v = "(none)";
        path += " › " + v;
        const id = "bd:" + i + path;
        let g = groups.get(id);
        if (!g) {
          g = { id, label: String(v), kind: "Aggregate", cls: "", count: 0, members: [], domains: new Map(), classes: new Map(), virtual: true, facet: k, facetValue: v };
          groups.set(id, g);
          nodes.push(g);
          links.push({ source: parent, target: id, pred: "groups", cat: "Decomposition" });
        }
        g.count++;
        g.members.push(n.id);
        g.domains.set(n.domain, (g.domains.get(n.domain) || 0) + 1);
        g.classes.set(n.cls, (g.classes.get(n.cls) || 0) + 1);
        parent = id;
      });
      if (instances) {
        nodes.push(rn(n));
        links.push({ source: parent, target: n.id, pred: "groups", cat: "Decomposition" });
      }
    }
    const top = (m) => [...m.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
    for (const g of groups.values()) {
      g.domain = top(g.domains) || "Other";
      g.cls = top(g.classes) || "";
      g.label = `${g.label} (${g.count})`;
      delete g.domains;
      delete g.classes;
    }
    if (!keys.length) notes.push("Choose one or more breakdown levels in the Scope panel.");
    else notes.push(`Breakdown: ${keys.join(" → ")}${instances ? " → elements" : ""}. Group size encodes element count; groups are coloured by their dominant discipline.`);
    return { nodes, links };
  }

  function lensSystems(model, s, notes) {
    const include = new Set();
    const systems = [...model.nodes.values()].filter((n) => (n.kind === "System" || n.kind === "Group") && !s.hidden?.has(n.id));
    for (const sys of systems) {
      const members = model.inE(sys.id, "memberOf").map((e) => model.nodes.get(e.s)).filter((n) => n && passes(n, s));
      if (!members.length && (s.domains || s.classes || s.storeys)) continue;
      include.add(sys.id);
      members.forEach((n) => include.add(n.id));
    }
    for (const e of model.edges) {
      if (e.pred === "feeds" || e.pred === "controlledBy") {
        const a = model.nodes.get(e.s);
        const b = model.nodes.get(e.t);
        if (passes(a, s) && passes(b, s)) { include.add(e.s); include.add(e.t); }
      }
    }
    for (const e of model.edges) {
      if (e.pred === "serves" && include.has(e.s)) include.add(e.t);
    }
    if (s.showPorts) {
      for (const id of [...include]) for (const e of model.outE(id, "hasPort")) include.add(e.t);
    }
    const preds = new Set(["memberOf", "feeds", "controlledBy", "serves", "connectedTo", ...(s.showPorts ? ["hasPort", "portConnectedTo"] : [])]);
    const links = model.edges.filter((e) => preds.has(e.pred) && include.has(e.s) && include.has(e.t)).map((e) => rl(e, { flow: e.pred === "feeds" || e.pred === "portConnectedTo" }));
    if (!include.size) notes.push("No IfcSystem / IfcDistributionSystem assignments or port connections (IfcRelConnectsPorts) were found in this model.");
    const feeds = links.filter((l) => l.pred === "feeds").length;
    if (include.size && !feeds) notes.push("No port-to-port connectivity found — only system membership is shown. Export with IfcDistributionPort + IfcRelConnectsPorts to obtain flow topology.");
    return { nodes: [...include].map((id) => rn(model.nodes.get(id))), links };
  }

  function lensConnectivity(model, s, notes) {
    const preds = new Set(["connectedTo", "hosts", "coveredBy", "interferesWith", "structurallyConnectedTo", "adheredBy", ...(s.showOpenings ? ["hasOpening", "fills"] : [])]);
    const include = new Set();
    const links = [];
    for (const e of model.edges) {
      let ok = preds.has(e.pred);
      if (!ok && e.pred === "aggregates") {
        const a = model.nodes.get(e.s);
        const b = model.nodes.get(e.t);
        ok = a.kind === "Element" && b.kind === "Element";
      }
      if (!ok) continue;
      const a = model.nodes.get(e.s);
      const b = model.nodes.get(e.t);
      if (!passes(a, s) || !passes(b, s)) continue;
      include.add(e.s);
      include.add(e.t);
      links.push(rl(e));
    }
    if (!links.length) notes.push("No IfcRelConnectsElements / IfcRelConnectsPathElements / IfcRelFillsElement relationships found in this model.");
    return { nodes: [...include].map((id) => rn(model.nodes.get(id))), links };
  }

  function lensSpaces(model, s, notes) {
    const include = new Set();
    for (const n of model.nodes.values()) if ((n.cls === "IfcSpace" || n.kind === "Exterior") && passes(n, s)) include.add(n.id);
    const preds = new Set(["adjacentTo", "accessTo"]);
    const links = model.edges.filter((e) => preds.has(e.pred) && include.has(e.s) && include.has(e.t)).map((e) => rl(e));
    if (s.showZones) {
      for (const e of model.edges) {
        if (e.pred === "memberOf" && include.has(e.s) && model.nodes.get(e.t)?.kind === "Zone") {
          include.add(e.t);
          links.push(rl(e));
        }
      }
    }
    if (s.showBoundaries) {
      for (const e of model.edges) {
        if (e.pred === "boundedBy" && include.has(e.s) && passes(model.nodes.get(e.t), s)) {
          include.add(e.t);
          links.push(rl(e));
        }
      }
    }
    if (!links.some((l) => preds.has(l.pred))) {
      notes.push("No space boundaries (IfcRelSpaceBoundary) found, so adjacency cannot be derived. Spaces are grouped by storey instead.");
      for (const n of [...include]) {
        const st = model.nodes.get(n)?.storey;
        if (st && passes(model.nodes.get(st), s)) {
          include.add(st);
          links.push({ source: st, target: n, pred: "aggregates", cat: "Decomposition" });
        }
      }
    }
    return { nodes: [...include].map((id) => rn(model.nodes.get(id))), links };
  }

  function classTaxonomy(classCounts, nodes, links, stopAt) {
    const added = new Map();
    const ensure = (cls, count) => {
      const id = "cls:" + cls;
      if (!added.has(id)) {
        const d = S.domainOf(cls, S.kindOf(cls), {});
        const node = { id, label: cls, kind: "Class", cls, domain: d.domain, count: 0, abstract: true, virtual: true };
        added.set(id, node);
        nodes.push(node);
      }
      const node = added.get(id);
      node.count += count;
      return node;
    };
    for (const [cls, count] of classCounts) {
      const leaf = ensure(cls, count);
      leaf.abstract = false;
      leaf.instances = count;
      const chain = S.chain(cls);
      for (let i = 1; i < chain.length; i++) {
        const parent = ensure(chain[i], count);
        const key = chain[i - 1] + ">" + chain[i];
        if (!added.has(key)) {
          added.set(key, true);
          links.push({ source: "cls:" + chain[i - 1], target: parent.id, pred: "subClassOf", cat: "Schema" });
        }
        if (chain[i] === stopAt) break;
      }
    }
    return added;
  }

  function lensTypes(model, s, notes) {
    const nodes = [];
    const links = [];
    const elements = [...model.nodes.values()].filter((n) => n.kind === "Element" && passes(n, s));
    const classCounts = new Map();
    for (const n of elements) classCounts.set(n.cls, (classCounts.get(n.cls) || 0) + 1);
    classTaxonomy(classCounts, nodes, links, "IfcProduct");
    const typeCount = new Map();
    for (const n of elements) if (n.typeId) typeCount.set(n.typeId, (typeCount.get(n.typeId) || 0) + 1);
    for (const [tid, c] of typeCount) {
      const t = model.nodes.get(tid);
      nodes.push(rn(t, { count: c }));
      const cls = t.elementClass && classCounts.has(t.elementClass) ? t.elementClass : null;
      const target = cls || [...classCounts.keys()].find((k) => elements.some((e) => e.typeId === tid && e.cls === k));
      if (target) links.push({ source: tid, target: "cls:" + target, pred: "typeFor", cat: "Typing" });
    }
    if (s.includeInstances) {
      for (const n of elements) {
        nodes.push(rn(n));
        if (n.typeId) links.push({ source: n.id, target: n.typeId, pred: "isTypedBy", cat: "Typing" });
        else links.push({ source: n.id, target: "cls:" + n.cls, pred: "instanceOf", cat: "Schema" });
      }
    }
    if (s.includeClassification) {
      const refNodes = new Map();
      const queue = [];
      const ensureRef = (id) => {
        if (refNodes.has(id)) return refNodes.get(id);
        const node = rn(model.nodes.get(id), { count: 0 });
        refNodes.set(id, node);
        nodes.push(node);
        queue.push(id);
        return node;
      };
      const refAgg = new Map();
      for (const n of elements) {
        for (const cr of n.classRefs || []) {
          ensureRef(cr).count++;
          if (s.includeInstances) links.push({ source: n.id, target: cr, pred: "hasClassification", cat: "Classification" });
          else {
            const k = n.cls + "|" + cr;
            refAgg.set(k, (refAgg.get(k) || 0) + 1);
          }
        }
      }
      for (const [k, w] of refAgg) {
        const [cls, cr] = k.split("|");
        links.push({ source: "cls:" + cls, target: cr, pred: "hasClassification", cat: "Classification", weight: w });
      }
      while (queue.length) {
        const id = queue.shift();
        for (const e of model.outE(id, "inSystem")) {
          ensureRef(e.t);
          links.push(rl(e));
        }
      }
    }
    return { nodes, links };
  }

  function lensMaterials(model, s, notes, detail) {
    const nodes = [];
    const links = [];
    const elements = [...model.nodes.values()].filter((n) => n.kind === "Element" && passes(n, s));
    const instances = detail === "instances" || (detail === "auto" && elements.length <= 400);
    const matNodes = new Map();
    const agg = new Map();
    const classNodes = new Map();
    for (const n of elements) {
      const mats = n.materials?.length ? n.materials : n.materialsInherited || [];
      if (!mats.length) continue;
      for (const mt of mats) {
        if (!matNodes.has(mt.id)) {
          const mn = model.nodes.get(mt.id);
          if (!mn) continue;
          matNodes.set(mt.id, rn(mn, { count: 0 }));
        }
        matNodes.get(mt.id).count++;
        if (instances) {
          links.push({ source: n.id, target: mt.id, pred: "hasMaterial", cat: "Material", inferred: !n.materials?.length, layer: mt.layer, thickness: mt.thickness });
        } else {
          const k = n.cls + "|" + mt.id;
          agg.set(k, (agg.get(k) || 0) + 1);
        }
      }
      if (instances) nodes.push(rn(n));
      else if (!classNodes.has(n.cls)) classNodes.set(n.cls, { id: "agg:" + n.cls, label: short(n.cls), kind: "Aggregate", cls: n.cls, domain: n.domain, count: 0, members: [] });
      if (!instances) {
        const c = classNodes.get(n.cls);
        c.count++;
        c.members.push(n.id);
      }
    }
    for (const [k, w] of agg) {
      const [cls, mid] = k.split("|");
      links.push({ source: "agg:" + cls, target: mid, pred: "hasMaterial", cat: "Material", weight: w });
    }
    nodes.push(...matNodes.values(), ...classNodes.values());
    if (!matNodes.size) notes.push("No material associations (IfcRelAssociatesMaterial) found for the current selection.");
    else if (!instances) notes.push("Elements are grouped by IFC class (edge weight = number of elements). Switch Detail to “Instances” to see individual elements.");
    return { nodes, links };
  }

  function lensProperties(model, s, notes) {
    const nodes = [];
    const links = [];
    const elements = [...model.nodes.values()].filter((n) => (n.kind === "Element" || n.kind === "Spatial") && passes(n, s));
    const clsNodes = new Map();
    const psetNodes = new Map();
    const propNodes = new Map();
    const cp = new Map();
    for (const n of elements) {
      if (!clsNodes.has(n.cls)) clsNodes.set(n.cls, { id: "agg:" + n.cls, label: short(n.cls), kind: "Aggregate", cls: n.cls, domain: n.domain, count: 0, members: [] });
      const c = clsNodes.get(n.cls);
      c.count++;
      c.members.push(n.id);
      for (const ps of model.propertySets(n)) {
        const pid = "pset:" + ps.name;
        if (!psetNodes.has(pid)) psetNodes.set(pid, { id: pid, label: ps.name, kind: ps.kind === "Q" ? "QuantitySet" : "PropertySet", cls: "", domain: "Semantic", count: 0, virtual: true, psetName: ps.name });
        psetNodes.get(pid).count++;
        const k = n.cls + "|" + pid;
        cp.set(k, (cp.get(k) || 0) + 1);
        if (s.includeProps) {
          for (const p of ps.props) {
            const prid = "prop:" + ps.name + "." + p.name;
            if (!propNodes.has(prid)) {
              propNodes.set(prid, { id: prid, label: p.name, kind: "Property", cls: "", domain: "Semantic", count: 0, virtual: true, propKey: ps.name + "." + p.name });
              links.push({ source: pid, target: prid, pred: "hasProperty", cat: "Properties" });
            }
            propNodes.get(prid).count++;
          }
        }
      }
    }
    for (const [k, w] of cp) {
      const [cls, pid] = k.split("|");
      const c = clsNodes.get(cls);
      links.push({ source: c.id, target: pid, pred: "hasPropertySet", cat: "Properties", weight: w, coverage: w / c.count });
    }
    const linked = new Set(links.map((l) => l.source));
    const bare = [...clsNodes.values()].filter((c) => !linked.has(c.id));
    nodes.push(...[...clsNodes.values()].filter((c) => linked.has(c.id)), ...psetNodes.values(), ...propNodes.values());
    if (bare.length) notes.push(`${bare.length} classes carry no property sets and are omitted: ${bare.slice(0, 8).map((c) => short(c.cls)).join(", ")}${bare.length > 8 ? "…" : ""}.`);
    notes.push("Class → property-set edges are weighted by the number of occurrences; edge opacity encodes coverage (share of the class carrying the set).");
    return { nodes, links };
  }

  function lensSchema(model, s, notes) {
    const nodes = [];
    const links = [];
    const counts = new Map();
    const ok = new Set();
    for (const n of model.nodes.values()) {
      if (!passes(n, s)) continue;
      ok.add(n.id);
      counts.set(n.cls, (counts.get(n.cls) || 0) + 1);
    }
    const agg = new Map();
    for (const e of model.edges) {
      if (!ok.has(e.s) || !ok.has(e.t)) continue;
      if (s.cats && !s.cats.has(e.cat)) continue;
      const a = model.nodes.get(e.s).cls;
      const b = model.nodes.get(e.t).cls;
      const k = a + "|" + e.pred + "|" + b;
      agg.set(k, (agg.get(k) || 0) + 1);
    }
    for (const [cls, c] of counts) {
      const kind = S.kindOf(cls);
      const any = [...model.nodes.values()].find((n) => n.cls === cls);
      nodes.push({ id: "cls:" + cls, label: cls, kind: "Class", cls, domain: any?.domain || S.domainOf(cls, kind, {}).domain, count: c, instances: c, virtual: true, nodeKind: any?.kind });
    }
    for (const [k, w] of agg) {
      const [a, pred, b] = k.split("|");
      const p = S.PRED[pred] || S.PRED.related;
      links.push({ source: "cls:" + a, target: "cls:" + b, pred, cat: p.cat, inferred: !!p.inferred, weight: w });
    }
    if (s.includeTaxonomy) {
      const extra = [];
      const extraLinks = [];
      classTaxonomy(new Map([...counts].filter(([c]) => c.startsWith("Ifc"))), extra, extraLinks, "IfcRoot");
      const have = new Set(nodes.map((n) => n.id));
      for (const n of extra) if (!have.has(n.id)) nodes.push(n);
      links.push(...extraLinks);
    }
    return { nodes, links };
  }

  function lensSubgraph(model, s, notes) {
    const ids = s.subgraph?.ids || new Set();
    const nodes = [];
    for (const id of ids) {
      const n = model.nodes.get(id);
      if (n && !s.hidden?.has(id)) nodes.push(rn(n, { focus: s.subgraph.focus?.has(id) }));
    }
    const links = model.edges.filter((e) => ids.has(e.s) && ids.has(e.t) && (!s.cats || s.cats.has(e.cat))).map((e) => rl(e, { onPath: s.subgraph.pathEdges?.has(e.id) }));
    if (s.subgraph?.title) notes.push(s.subgraph.title);
    return { nodes, links };
  }

  // ── Abstraction & budget ──────────────────────────────────────────────

  function collapse(graph, keyFn) {
    const map = new Map();
    const aggs = new Map();
    for (const n of graph.nodes) {
      const k = keyFn(n);
      if (!k) continue;
      const aid = "agg:" + k;
      map.set(n.id, aid);
      let a = aggs.get(aid);
      if (!a) {
        a = { id: aid, label: short(n.cls), kind: "Aggregate", cls: n.cls, domain: n.domain, storey: n.storey, count: 0, members: [] };
        aggs.set(aid, a);
      }
      a.count++;
      a.members.push(n.ref || n.id);
    }
    if (!map.size) return graph;
    for (const a of aggs.values()) a.label = `${short(a.cls)} ×${a.count}`;
    const nodes = graph.nodes.filter((n) => !map.has(n.id)).concat([...aggs.values()]);
    const merged = new Map();
    for (const l of graph.links) {
      const s = map.get(l.source) || l.source;
      const t = map.get(l.target) || l.target;
      if (s === t) continue;
      const key = s + "|" + l.pred + "|" + t;
      const ex = merged.get(key);
      if (ex) ex.weight = (ex.weight || 1) + (l.weight || 1);
      else merged.set(key, { ...l, source: s, target: t, weight: l.weight || 1 });
    }
    return { nodes, links: [...merged.values()] };
  }

  function curvatures(links) {
    const groups = new Map();
    for (const l of links) {
      const a = l.source < l.target ? l.source + "|" + l.target : l.target + "|" + l.source;
      let g = groups.get(a);
      if (!g) groups.set(a, (g = []));
      g.push(l);
    }
    for (const g of groups.values()) {
      if (g.length < 2) continue;
      g.forEach((l, i) => {
        const sign = l.source < l.target ? 1 : -1;
        l.curvature = sign * (0.15 + 0.12 * Math.floor(i / 2)) * (i % 2 ? -1 : 1);
      });
    }
  }

  function build(model, s) {
    const notes = [];
    const lens = s.lens;
    let g;
    switch (lens) {
      case "spatial": g = lensSpatial(model, s, notes); break;
      case "systems": g = lensSystems(model, s, notes); break;
      case "breakdown": g = lensBreakdown(model, s, notes); break;
      case "connectivity": g = lensConnectivity(model, s, notes); break;
      case "spaces": g = lensSpaces(model, s, notes); break;
      case "types": g = lensTypes(model, s, notes); break;
      case "materials": g = lensMaterials(model, s, notes, s.detail); break;
      case "properties": g = lensProperties(model, s, notes); break;
      case "schema": g = lensSchema(model, s, notes); break;
      case "subgraph": g = lensSubgraph(model, s, notes); break;
      default: g = lensKG(model, s, notes);
    }

    // Abstraction: collapse instances into class groups
    const collapsible = !["types", "materials", "properties", "schema", "subgraph", "breakdown"].includes(lens);
    const elemCount = g.nodes.filter((n) => ELEMENTISH.has(n.kind)).length;
    let detail = s.detail;
    if (collapsible && detail === "auto") detail = g.nodes.length > s.maxNodes ? (lens === "spatial" ? "class-parent" : "class-storey") : "instances";
    if (collapsible && detail !== "instances" && elemCount) {
      let keyFn;
      if (detail === "class") keyFn = (n) => (ELEMENTISH.has(n.kind) ? n.cls : null);
      else if (detail === "class-parent") {
        const parent = new Map();
        for (const l of g.links) if (l.pred === "contains" || l.pred === "aggregates") parent.set(l.target, l.source);
        keyFn = (n) => (ELEMENTISH.has(n.kind) ? n.cls + "@" + (parent.get(n.id) || "-") : null);
      } else keyFn = (n) => (ELEMENTISH.has(n.kind) ? n.cls + "@" + (n.storey || "-") : null);
      g = collapse(g, keyFn);
      if (s.detail === "auto") notes.push(`${elemCount.toLocaleString()} elements exceed the node budget, so they are grouped by class${detail === "class-storey" ? " and storey" : ""}. Choose Detail → Instances to expand.`);
    }

    // Budget
    let truncated = 0;
    if (g.nodes.length > s.maxNodes) {
      const deg = new Map();
      for (const l of g.links) {
        deg.set(l.source, (deg.get(l.source) || 0) + 1);
        deg.set(l.target, (deg.get(l.target) || 0) + 1);
      }
      const rank = (n) => (ELEMENTISH.has(n.kind) ? 0 : 1e6) + (deg.get(n.id) || 0);
      const sorted = g.nodes.slice().sort((a, b) => rank(b) - rank(a));
      const keep = new Set(sorted.slice(0, s.maxNodes).map((n) => n.id));
      truncated = g.nodes.length - keep.size;
      g = { nodes: g.nodes.filter((n) => keep.has(n.id)), links: g.links.filter((l) => keep.has(l.source) && keep.has(l.target)) };
      notes.push(`Node budget reached: ${truncated.toLocaleString()} lower-degree nodes hidden. Raise “Max nodes”, filter, or use a coarser Detail level.`);
    }

    // Drop dangling links
    const ids = new Set(g.nodes.map((n) => n.id));
    g.links = g.links.filter((l) => ids.has(l.source) && ids.has(l.target));
    curvatures(g.links);
    g.info = { lens, notes, truncated, detail };
    return g;
  }

  KG.Views = { LENSES, build, passes, collapse };
})((window.KG = window.KG || {}));
