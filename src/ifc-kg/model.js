/* BIMUniXchange · BIM Knowledge Graph Studio
 * model.js — semantic knowledge-graph model built from an IFC STEP file.
 *
 * Nodes  : IfcRoot objects (products, spatial elements, types, systems, zones,
 *          processes…) plus resource entities promoted to first-class nodes
 *          (materials, classification references, documents).
 * Edges  : one per objectified IfcRel* pair, with provenance, plus inferred
 *          edges (hosting, MEP flow, space adjacency / door access).
 * Data   : property & quantity sets are attached to nodes (shared by id) and
 *          resolved with units, including properties inherited from types.
 */
(function (KG) {
  "use strict";
  const S = KG.Schema;
  const { refId, refIds, str, plain, Ref, Enum, Typed } = KG.Step;

  const tick = () => new Promise((r) => setTimeout(r, 0));

  class KGModel {
    constructor(source) {
      this.source = source; // "ifc" | "csv"
      this.meta = { warnings: [], units: {}, counts: {} };
      this.nodes = new Map();
      this.edges = [];
      this.edgeKeys = new Set();
      this.psets = new Map();
      this.out = new Map();
      this.in = new Map();
      this.storeys = [];
      this.unresolved = 0;
      this.rawProvider = null;
      this.attrProvider = null;
    }

    addNode(n) {
      if (this.nodes.has(n.id)) return this.nodes.get(n.id);
      n.kind = n.kind || "Other";
      n.cls = n.cls || "";
      n.name = n.name || "";
      n.psetIds = n.psetIds || [];
      if (!n.label) n.label = n.name || n.longName || `${(n.cls || n.kind).replace(/^Ifc/, "")} ${n.ifcId != null ? "#" + n.ifcId : ""}`.trim();
      this.nodes.set(n.id, n);
      return n;
    }

    addEdge(s, t, pred, attrs) {
      if (s === t || !this.nodes.has(s) || !this.nodes.has(t)) {
        if (!this.nodes.has(s) || !this.nodes.has(t)) this.unresolved++;
        return null;
      }
      const key = s + "|" + pred + "|" + t;
      if (this.edgeKeys.has(key)) return null;
      this.edgeKeys.add(key);
      const p = S.PRED[pred] || S.PRED.related;
      const e = { id: this.edges.length, s, t, pred, cat: p.cat, inferred: !!p.inferred, attrs: attrs || null };
      this.edges.push(e);
      return e;
    }

    buildIndex() {
      this.out = new Map();
      this.in = new Map();
      for (const e of this.edges) {
        let o = this.out.get(e.s);
        if (!o) this.out.set(e.s, (o = []));
        o.push(e);
        let i = this.in.get(e.t);
        if (!i) this.in.set(e.t, (i = []));
        i.push(e);
      }
    }

    outE(id, pred) {
      const l = this.out.get(id) || [];
      return pred ? l.filter((e) => e.pred === pred) : l;
    }
    inE(id, pred) {
      const l = this.in.get(id) || [];
      return pred ? l.filter((e) => e.pred === pred) : l;
    }
    neighbors(id) {
      const s = new Set();
      for (const e of this.out.get(id) || []) s.add(e.t);
      for (const e of this.in.get(id) || []) s.add(e.s);
      return s;
    }

    /** All property/quantity sets of a node, including inherited type sets. */
    propertySets(node) {
      const out = [];
      for (const id of node.psetIds || []) {
        const p = this.psets.get(id);
        if (p) out.push({ ...p, inherited: false });
      }
      if (node.typeId) {
        const t = this.nodes.get(node.typeId);
        if (t) {
          const own = new Set(out.map((p) => p.name));
          for (const id of t.psetIds || []) {
            const p = this.psets.get(id);
            if (p && !own.has(p.name)) out.push({ ...p, inherited: true });
          }
        }
      }
      return out;
    }

    /** Map "Pset.Property" → value (cached per node). */
    propMap(node) {
      if (node._pv) return node._pv;
      const m = new Map();
      for (const ps of this.propertySets(node)) {
        for (const p of ps.props) {
          const k = ps.name + "." + p.name;
          if (!m.has(k)) m.set(k, p.value);
        }
      }
      node._pv = m;
      return m;
    }

    propValue(node, key) {
      return this.propMap(node).get(key);
    }

    /**
     * Resolve an attribute name (Name, Class, Storey, Family, …), a property key
     * ("Pset_WallCommon.FireRating") or a wildcard property key ("*.FireRating").
     */
    valueOf(n, key) {
      switch (key) {
        case "Name": return n.name || n.label;
        case "GlobalId": return n.guid;
        case "Class": return n.cls;
        case "Kind": return n.kind;
        case "ObjectType": return n.objectType;
        case "Tag": return n.tag;
        case "PredefinedType": return n.predefinedType;
        case "Discipline": return n.domain;
        case "Storey": return n.storey ? this.nodes.get(n.storey)?.label : undefined;
        case "Type": return n.typeId ? this.nodes.get(n.typeId)?.label : n.typeName || undefined;
        case "Family": return n.family || (n.typeId ? this.nodes.get(n.typeId)?.objectType?.split(":")[0] : undefined);
        case "Category": return n.sourceCategory || (n.classRefs?.length ? this.nodes.get(n.classRefs[0])?.label : undefined);
        case "Layer": return n.layer;
        case "Phase": return n.phaseCreated;
        case "Material": {
          const mats = n.materials?.length ? n.materials : n.materialsInherited || [];
          return mats.length ? mats.map((x) => x.name || this.nodes.get(x.id)?.label).join(", ") : undefined;
        }
        case "Classification": return n.classRefs?.length ? n.classRefs.map((c) => this.nodes.get(c)?.label).join(", ") : undefined;
        case "System": return (this.out.get(n.id) || []).filter((e) => e.pred === "memberOf").map((e) => this.nodes.get(e.t)?.label).join(", ") || undefined;
        case "Host": return n.hostId ? this.nodes.get(n.hostId)?.label : undefined;
      }
      if (key.startsWith("*.")) {
        const suf = key.slice(1).toLowerCase();
        for (const [k, v] of this.propMap(n)) if (k.toLowerCase().endsWith(suf)) return v;
        return undefined;
      }
      return this.propValue(n, key);
    }

    /** Ancestor chain through decomposition / containment (closest first). */
    spatialPath(id) {
      const path = [];
      const seen = new Set([id]);
      let cur = id;
      for (let guard = 0; guard < 20; guard++) {
        const ins = this.in.get(cur) || [];
        const up = ins.find((e) => e.pred === "contains" || e.pred === "aggregates" || e.pred === "hasOpening" || e.pred === "hasPort" || e.pred === "nests");
        if (!up || seen.has(up.s)) break;
        seen.add(up.s);
        path.push(up.s);
        cur = up.s;
      }
      return path;
    }

    raw(node) {
      return this.rawProvider ? this.rawProvider(node) : "";
    }
    attributes(node) {
      return this.attrProvider ? this.attrProvider(node) : [];
    }

    /** Compute storeys, inherited data, domains, catalogs. */
    finalize() {
      this.buildIndex();
      // Storeys ordered by elevation
      const storeys = [...this.nodes.values()].filter((n) => n.cls === "IfcBuildingStorey");
      storeys.sort((a, b) => (a.elevation ?? 0) - (b.elevation ?? 0) || a.label.localeCompare(b.label));
      storeys.forEach((s, i) => (s.storeyIndex = i));
      this.storeys = storeys.map((s) => s.id);

      // Storey of every node (via containment / decomposition chain)
      const memo = new Map();
      const storeyOf = (id, depth = 0) => {
        if (memo.has(id)) return memo.get(id);
        const n = this.nodes.get(id);
        if (!n || depth > 25) return null;
        if (n.cls === "IfcBuildingStorey") { memo.set(id, id); return id; }
        memo.set(id, null); // cycle guard
        let res = null;
        for (const e of this.in.get(id) || []) {
          if (e.pred === "contains" || e.pred === "aggregates" || e.pred === "hasOpening" || e.pred === "hasPort" || e.pred === "nests" || e.pred === "fills" || e.pred === "coveredBy") {
            const p = this.nodes.get(e.s);
            if (p && (p.kind === "Spatial" || p.kind === "Element" || p.kind === "Opening")) {
              res = storeyOf(e.s, depth + 1);
              if (res) break;
            }
          }
        }
        if (!res) {
          // doors fill openings: the opening is the target of "fills"
          for (const e of this.out.get(id) || []) {
            if (e.pred === "fills" || e.pred === "references") {
              res = storeyOf(e.t, depth + 1);
              if (res) break;
            }
          }
        }
        memo.set(id, res);
        return res;
      };
      for (const n of this.nodes.values()) {
        if (n.kind === "Element" || n.kind === "Spatial" || n.kind === "Opening" || n.kind === "Port") {
          const s = storeyOf(n.id);
          if (s && s !== n.id) n.storey = s;
          else if (s === n.id) n.storey = s;
        }
      }

      // Type & material inheritance
      for (const n of this.nodes.values()) {
        if (n.kind !== "Element") continue;
        const te = (this.out.get(n.id) || []).find((e) => e.pred === "isTypedBy");
        if (te) n.typeId = te.t;
        if (!n.materials?.length && n.typeId) {
          const t = this.nodes.get(n.typeId);
          if (t?.materials?.length) n.materialsInherited = t.materials;
        }
      }

      // Domains
      for (const n of this.nodes.values()) {
        if (n.domain) continue;
        const hints = {};
        if (n.kind === "Element") {
          const pm = this.propMap(n);
          for (const [k, v] of pm) {
            if (/\.LoadBearing$/i.test(k) && typeof v === "boolean") { hints.loadBearing = v; break; }
            if (/\.LoadBearing$/i.test(k) && typeof v === "string") { hints.loadBearing = /^(true|yes|1|t)$/i.test(v); break; }
          }
          const sys = (this.out.get(n.id) || []).find((e) => e.pred === "memberOf" && this.nodes.get(e.t)?.kind === "System");
          if (sys) {
            const s = this.nodes.get(sys.t);
            hints.systemLabel = [s.predefinedType, s.objectType, s.name, s.longName].filter(Boolean).join(" ");
            n.systemName = s.label;
          } else if (n.systemHint) hints.systemLabel = n.systemHint;
          if (hints.loadBearing === undefined && n.loadBearingHint !== undefined) hints.loadBearing = n.loadBearingHint;
        }
        const d = S.domainOf(n.cls, n.kind, hints);
        n.domain = d.domain;
        n.domainReason = d.reason;
      }
      // Systems inherit the dominant domain of their members
      for (const n of this.nodes.values()) {
        if (n.kind !== "System") continue;
        const fromLabel = S.domainFromSystemLabel([n.predefinedType, n.objectType, n.name, n.longName].filter(Boolean).join(" "));
        if (fromLabel) { n.memberDomain = fromLabel; continue; }
        const counts = {};
        for (const e of this.in.get(n.id) || []) {
          if (e.pred !== "memberOf") continue;
          const m = this.nodes.get(e.s);
          if (m?.domain) counts[m.domain] = (counts[m.domain] || 0) + 1;
        }
        const best = Object.entries(counts).sort((a, b) => b[1] - a[1])[0];
        if (best) n.memberDomain = best[0];
      }

      this.computeCatalogs();
    }

    computeCatalogs() {
      const classCount = new Map();
      const kindCount = new Map();
      const domainCount = new Map();
      for (const n of this.nodes.values()) {
        classCount.set(n.cls, (classCount.get(n.cls) || 0) + 1);
        kindCount.set(n.kind, (kindCount.get(n.kind) || 0) + 1);
        if (n.kind === "Element") domainCount.set(n.domain, (domainCount.get(n.domain) || 0) + 1);
      }
      const predCount = new Map();
      const catCount = new Map();
      for (const e of this.edges) {
        predCount.set(e.pred, (predCount.get(e.pred) || 0) + 1);
        catCount.set(e.cat, (catCount.get(e.cat) || 0) + 1);
      }
      // Property catalog
      const props = new Map();
      for (const n of this.nodes.values()) {
        if (n.kind !== "Element" && n.kind !== "Spatial" && n.kind !== "System" && n.kind !== "Zone") continue;
        for (const ps of this.propertySets(n)) {
          for (const p of ps.props) {
            const key = ps.name + "." + p.name;
            let c = props.get(key);
            if (!c) {
              c = { key, pset: ps.name, name: p.name, kind: ps.kind, count: 0, inherited: 0, datatypes: new Set(), units: new Set(), values: new Map(), overflow: false, min: Infinity, max: -Infinity, classes: new Map() };
              props.set(key, c);
            }
            c.count++;
            if (ps.inherited) c.inherited++;
            if (p.datatype) c.datatypes.add(p.datatype);
            if (p.unit) c.units.add(p.unit);
            c.classes.set(n.cls, (c.classes.get(n.cls) || 0) + 1);
            const v = p.value;
            if (typeof v === "number") {
              if (v < c.min) c.min = v;
              if (v > c.max) c.max = v;
            }
            const vk = v == null ? "∅" : String(v);
            if (c.values.has(vk)) c.values.set(vk, c.values.get(vk) + 1);
            else if (c.values.size < 400) c.values.set(vk, 1);
            else c.overflow = true;
          }
        }
      }
      for (const c of props.values()) {
        let applicable = 0;
        for (const cls of c.classes.keys()) applicable += classCount.get(cls) || 0;
        c.coverage = applicable ? c.count / applicable : 0;
        c.distinct = c.values.size + (c.overflow ? "+" : "");
      }
      this.catalog = { classCount, kindCount, domainCount, predCount, catCount, props };
      this.meta.counts = {
        nodes: this.nodes.size,
        edges: this.edges.length,
        elements: kindCount.get("Element") || 0,
        spatial: kindCount.get("Spatial") || 0,
        types: kindCount.get("Type") || 0,
        psets: this.psets.size,
        properties: props.size,
        inferred: this.edges.filter((e) => e.inferred).length,
      };
    }
  }

  // ───────────────────────── IFC → KG builder ─────────────────────────

  const SKIP_REL = new Set(["IFCRELDECLARES", "IFCRELDEFINESBYTEMPLATE", "IFCRELASSOCIATESCONSTRAINT", "IFCRELASSOCIATESAPPROVAL", "IFCRELASSOCIATESPROFILEDEF", "IFCRELASSOCIATESLIBRARY"]);
  const PSET_TYPES = new Set(["IFCPROPERTYSET", "IFCELEMENTQUANTITY"]);

  async function buildFromIFC(step, opts = {}) {
    const progress = opts.onProgress || (async () => {});
    const m = new KGModel("ifc");
    m.step = step;
    const schemaId = step.schema || "";
    m.meta.schema = schemaId;
    m.meta.header = step.header;
    m.meta.fileName = opts.fileName || step.header.name || "model.ifc";
    m.meta.fileSize = opts.fileSize || step.text.length;
    m.meta.entityCount = step.size;
    m.meta.parseErrors = step.errors;

    // Entity type histogram (all STEP instances)
    const typeHist = [];
    for (const [t, ids] of step.byType) typeHist.push([t, ids.length]);
    typeHist.sort((a, b) => b[1] - a[1]);
    m.meta.typeHistogram = typeHist;

    // ── Units ──
    const units = {};
    const describeUnit = (id) => {
      const t = step.type(id);
      const a = step.args(id);
      if (t === "IFCSIUNIT") {
        const name = str(a[3]);
        const prefix = str(a[2]);
        return { type: str(a[1]), symbol: (S.SI_PREFIX[prefix] || "") + (S.SI_SYMBOL[name] || name.toLowerCase()) };
      }
      if (t === "IFCCONVERSIONBASEDUNIT" || t === "IFCCONVERSIONBASEDUNITWITHOFFSET") {
        return { type: str(a[1]), symbol: str(a[2]).toLowerCase() };
      }
      if (t === "IFCCONTEXTDEPENDENTUNIT") return { type: str(a[1]), symbol: str(a[2]) };
      if (t === "IFCDERIVEDUNIT") {
        return { type: str(a[1]), symbol: str(a[2]) || str(a[1]).replace(/UNIT$/, "").toLowerCase() };
      }
      if (t === "IFCMONETARYUNIT") return { type: "MONETARYUNIT", symbol: str(a[0]) };
      return null;
    };
    for (const pid of step.ofType("IFCPROJECT")) {
      const ua = refId(step.args(pid)[8]);
      if (ua) for (const uid of refIds(step.args(ua)[0])) {
        const u = describeUnit(uid);
        if (u && u.type && !units[u.type]) units[u.type] = u.symbol;
      }
    }
    m.meta.units = units;
    const unitFor = (unitArg, measureType) => {
      const uid = refId(unitArg);
      if (uid) return describeUnit(uid)?.symbol || "";
      const ut = S.MEASURE_UNIT[measureType];
      return ut ? units[ut] || "" : "";
    };

    // ── Authoring application / provenance ──
    const app = step.ofType("IFCAPPLICATION")[0];
    if (app) {
      const a = step.args(app);
      m.meta.application = `${str(a[2])} ${str(a[1])}`.trim();
    }
    const oh = step.ofType("IFCOWNERHISTORY")[0];
    if (oh) {
      const a = step.args(oh);
      const ts = typeof a[7] === "number" ? a[7] : null;
      if (ts) m.meta.created = new Date(ts * 1000).toISOString();
    }
    await progress(0.05, "Indexed units & provenance");

    // ── Classify entity types ──
    const rootTypes = [];
    const relTypes = [];
    const psetTypes = [];
    for (const [type, ids] of step.byType) {
      if (!ids.length || !ids.slice(0, 3).every((id) => step.isRooted(id))) continue;
      if (type.startsWith("IFCREL")) relTypes.push(type);
      else if (PSET_TYPES.has(type) || /PROPERTIES$/.test(type)) psetTypes.push(type);
      else if (/TEMPLATE$/.test(type)) continue;
      else rootTypes.push(type);
    }

    // ── Nodes ──
    const isLegacy = /2X3/i.test(schemaId);
    let done = 0;
    const totalRoot = rootTypes.reduce((s, t) => s + step.ofType(t).length, 0) || 1;
    for (const type of rootTypes) {
      const cls = S.canonical(type);
      const kind = S.kindOf(cls);
      const names = S.attributeNames(cls);
      const idx = (nm) => names.indexOf(nm);
      for (const id of step.ofType(type)) {
        const a = step.args(id);
        const n = {
          id: "#" + id,
          ifcId: id,
          kind,
          cls,
          guid: str(a[0]),
          name: str(a[2]),
          description: str(a[3]),
        };
        if (idx("ObjectType") > 0) n.objectType = str(a[4]);
        if (idx("Tag") > 0) n.tag = str(a[idx("Tag")]);
        if (idx("LongName") > 0) n.longName = str(a[idx("LongName")]);
        if (cls === "IfcBuildingStorey") n.elevation = typeof a[9] === "number" ? a[9] : null;
        if (cls === "IfcSite") {
          const dms = (v) => (Array.isArray(v) && v.length >= 3 ? v[0] + v[1] / 60 + (v[2] + (v[3] || 0) / 1e6) / 3600 : null);
          n.latitude = dms(a[9]);
          n.longitude = dms(a[10]);
          n.refElevation = typeof a[11] === "number" ? a[11] : null;
        }
        if (cls === "IfcDistributionPort") n.flowDirection = str(a[7]);
        // PredefinedType
        let pt = idx("PredefinedType");
        if (cls === "IfcSpace" && isLegacy) pt = -1;
        if (pt > 0 && a[pt] instanceof Enum) n.predefinedType = a[pt].v;
        else if (kind === "Element" || kind === "Type") {
          for (let k = kind === "Type" ? 6 : 8; k < a.length; k++) {
            if (a[k] instanceof Enum && !["T", "F", "U"].includes(a[k].v)) { n.predefinedType = a[k].v; break; }
          }
        }
        if (n.predefinedType === "NOTDEFINED") n.predefinedType = "";
        if (kind === "Type") {
          n.psetIds = refIds(a[5]);
          n.elementClass = S.elementClassOfType(cls);
        }
        if (!n.name && n.longName) n.label = n.longName;
        if (cls === "IfcBuildingStorey" && n.elevation != null) n.sub = `Elev. ${round(n.elevation)} ${units.LENGTHUNIT || ""}`.trim();
        m.addNode(n);
        if (++done % 20000 === 0) await progress(0.05 + 0.3 * (done / totalRoot), `Created ${done.toLocaleString()} nodes`);
      }
    }
    await progress(0.35, "Resolving property sets");

    // ── Property & quantity sets ──
    const resolveProp = (pid, prefix = "") => {
      const t = step.type(pid);
      const a = step.args(pid);
      const name = prefix + str(a[0]);
      switch (t) {
        case "IFCPROPERTYSINGLEVALUE": {
          const v = a[2];
          const pv = plain(v);
          return [{ name, value: pv.value, datatype: v instanceof Typed ? S.canonical(v.type) : pv.datatype, unit: unitFor(a[3], v instanceof Typed ? v.type : null) }];
        }
        case "IFCPROPERTYENUMERATEDVALUE": {
          const vals = (Array.isArray(a[2]) ? a[2] : []).map((x) => plain(x).value);
          return [{ name, value: vals.join(", "), datatype: "IfcPropertyEnumeratedValue" }];
        }
        case "IFCPROPERTYBOUNDEDVALUE": {
          const up = plain(a[2]).value;
          const lo = plain(a[3]).value;
          const sp = plain(a[5]).value;
          const unitT = a[2] instanceof Typed ? a[2].type : a[3] instanceof Typed ? a[3].type : null;
          let val = `${lo ?? "–∞"} … ${up ?? "+∞"}`;
          if (sp != null) val += ` (set point ${sp})`;
          return [{ name, value: val, datatype: "IfcPropertyBoundedValue", unit: unitFor(a[4], unitT) }];
        }
        case "IFCPROPERTYLISTVALUE": {
          const vals = (Array.isArray(a[2]) ? a[2] : []).map((x) => plain(x).value);
          return [{ name, value: vals.join(", "), datatype: "IfcPropertyListValue", unit: unitFor(a[3], a[2]?.[0]?.type) }];
        }
        case "IFCPROPERTYTABLEVALUE": {
          const rows = Array.isArray(a[2]) ? a[2].length : 0;
          return [{ name, value: `(table · ${rows} rows)`, datatype: "IfcPropertyTableValue" }];
        }
        case "IFCPROPERTYREFERENCEVALUE": {
          const r = refId(a[3]);
          let label = r ? S.canonical(step.type(r)) + " #" + r : "";
          if (r && step.isRooted(r)) label = str(step.args(r)[2]) || label;
          else if (r) label = str(step.args(r)[0]) || label;
          return [{ name, value: label, datatype: "IfcPropertyReferenceValue" }];
        }
        case "IFCCOMPLEXPROPERTY": {
          const out = [];
          for (const c of refIds(a[3])) out.push(...resolveProp(c, name + "."));
          return out;
        }
        case "IFCQUANTITYLENGTH":
        case "IFCQUANTITYAREA":
        case "IFCQUANTITYVOLUME":
        case "IFCQUANTITYWEIGHT":
        case "IFCQUANTITYTIME":
        case "IFCQUANTITYCOUNT":
        case "IFCQUANTITYNUMBER": {
          const v = typeof a[3] === "number" ? a[3] : plain(a[3]).value;
          return [{ name, value: v, datatype: S.canonical(t), unit: t === "IFCQUANTITYCOUNT" ? "" : unitFor(a[2], t) }];
        }
        case "IFCPHYSICALCOMPLEXQUANTITY": {
          const out = [];
          for (const c of refIds(a[2])) out.push(...resolveProp(c, name + "."));
          return out;
        }
        default:
          return [];
      }
    };
    for (const type of psetTypes) {
      for (const id of step.ofType(type)) {
        const a = step.args(id);
        let props = [];
        let kind = "P";
        if (type === "IFCPROPERTYSET") props = refIds(a[4]).flatMap((p) => resolveProp(p));
        else if (type === "IFCELEMENTQUANTITY") {
          kind = "Q";
          props = refIds(a[5]).flatMap((p) => resolveProp(p));
        } else {
          // Pre-defined property sets (IfcDoorLiningProperties …): expose numeric attributes
          const cls = S.canonical(type);
          props = a.slice(4).map((v, i) => ({ name: `Attr${i + 4}`, value: plain(v).value, datatype: v instanceof Typed ? v.type : typeof v })).filter((p) => p.value != null && p.datatype !== "object");
          m.psets.set(id, { id, name: str(a[2]) || cls, kind: "P", props, cls });
          continue;
        }
        m.psets.set(id, { id, name: str(a[2]) || `${kind === "Q" ? "QuantitySet" : "PropertySet"} #${id}`, kind, props, method: kind === "Q" ? str(a[4]) : "" });
      }
    }
    await progress(0.5, "Resolving relationships");

    // ── Resource nodes: materials, classifications, documents ──
    const ensureMaterial = (mid) => {
      const key = "#" + mid;
      if (!m.nodes.has(key)) {
        const a = step.args(mid);
        m.addNode({ id: key, ifcId: mid, kind: "Material", cls: "IfcMaterial", name: str(a[0]), description: str(a[1]), category: str(a[2]) });
      }
      return key;
    };
    const materialCache = new Map();
    const resolveMaterial = (rid, depth = 0) => {
      if (materialCache.has(rid)) return materialCache.get(rid);
      const t = step.type(rid);
      const a = step.args(rid);
      let out = [];
      if (depth > 6) return out;
      switch (t) {
        case "IFCMATERIAL":
          out = [{ id: ensureMaterial(rid) }];
          break;
        case "IFCMATERIALLAYERSETUSAGE":
          out = resolveMaterial(refId(a[0]), depth + 1).map((x) => ({ ...x, usage: `${str(a[1])} ${str(a[2])}`.trim() }));
          break;
        case "IFCMATERIALLAYERSET": {
          const setName = str(a[1]);
          refIds(a[0]).forEach((lid, i) => {
            const la = step.args(lid);
            const mid = refId(la[0]);
            if (!mid) return;
            out.push({ id: ensureMaterial(mid), layer: i + 1, thickness: typeof la[1] === "number" ? la[1] : null, layerName: str(la[3]), setName, ventilated: la[2] instanceof Enum ? la[2].v === "T" : undefined });
          });
          break;
        }
        case "IFCMATERIALLAYER": {
          const mid = refId(a[0]);
          if (mid) out = [{ id: ensureMaterial(mid), thickness: typeof a[1] === "number" ? a[1] : null }];
          break;
        }
        case "IFCMATERIALLIST":
          out = refIds(a[0]).map((mid) => ({ id: ensureMaterial(mid) }));
          break;
        case "IFCMATERIALCONSTITUENTSET": {
          const setName = str(a[0]);
          for (const cid of refIds(a[2])) {
            const ca = step.args(cid);
            const mid = refId(ca[2]);
            if (mid) out.push({ id: ensureMaterial(mid), constituent: str(ca[0]), fraction: typeof ca[3] === "number" ? ca[3] : null, setName });
          }
          break;
        }
        case "IFCMATERIALPROFILESETUSAGE":
        case "IFCMATERIALPROFILESETUSAGETAPERING":
          out = resolveMaterial(refId(a[0]), depth + 1);
          break;
        case "IFCMATERIALPROFILESET": {
          const setName = str(a[0]);
          for (const pid of refIds(a[2])) {
            const pa = step.args(pid);
            const mid = refId(pa[2]);
            const prof = refId(pa[3]);
            if (mid) out.push({ id: ensureMaterial(mid), profile: prof ? str(step.args(prof)[1]) : "", setName });
          }
          break;
        }
        case "IFCMATERIALPROFILE": {
          const mid = refId(a[2]);
          if (mid) out = [{ id: ensureMaterial(mid) }];
          break;
        }
      }
      materialCache.set(rid, out);
      return out;
    };

    const ensureClassSystem = (cid) => {
      const key = "#" + cid;
      if (!m.nodes.has(key)) {
        const a = step.args(cid);
        m.addNode({ id: key, ifcId: cid, kind: "ClassSystem", cls: "IfcClassification", name: str(a[3]) || str(a[0]), source: str(a[0]), edition: str(a[1]), location: str(a[5]) });
      }
      return key;
    };
    const ensureClassRef = (rid, depth = 0) => {
      const key = "#" + rid;
      if (m.nodes.has(key)) return key;
      const t = step.type(rid);
      if (t === "IFCCLASSIFICATION") return ensureClassSystem(rid);
      if (t !== "IFCCLASSIFICATIONREFERENCE") return null;
      const a = step.args(rid);
      const code = str(a[1]);
      const name = str(a[2]);
      m.addNode({ id: key, ifcId: rid, kind: "Classification", cls: "IfcClassificationReference", name: name || code, code, label: [code, name].filter(Boolean).join(" · ") || "Classification #" + rid, location: str(a[0]) });
      const src = refId(a[3]);
      if (src && depth < 10) {
        const parent = ensureClassRef(src, depth + 1);
        if (parent) {
          m.addEdge(key, parent, "inSystem");
          const pn = m.nodes.get(parent);
          m.nodes.get(key).scheme = pn.kind === "ClassSystem" ? pn.label : pn.scheme;
        }
      }
      return key;
    };
    const ensureDocument = (did) => {
      const key = "#" + did;
      if (m.nodes.has(key)) return key;
      const t = step.type(did);
      const a = step.args(did);
      if (t === "IFCDOCUMENTREFERENCE") m.addNode({ id: key, ifcId: did, kind: "Document", cls: "IfcDocumentReference", name: str(a[2]) || str(a[1]) || str(a[0]), location: str(a[0]) });
      else if (t === "IFCDOCUMENTINFORMATION") m.addNode({ id: key, ifcId: did, kind: "Document", cls: "IfcDocumentInformation", name: str(a[1]) || str(a[0]), location: str(a[3]) });
      else return null;
      return key;
    };

    // ── Relationships ──
    const boundaries = [];
    const relStats = {};
    let relDone = 0;
    const totalRel = relTypes.reduce((s, t) => s + step.ofType(t).length, 0) || 1;
    for (const type of relTypes) {
      if (SKIP_REL.has(type)) continue;
      const relCls = S.canonical(type);
      let spec = S.REL[type];
      for (const id of step.ofType(type)) {
        const a = step.args(id);
        relStats[relCls] = (relStats[relCls] || 0) + 1;
        let sp = spec;
        if (!sp) {
          // Generic: first single reference → every listed reference
          let si = -1;
          let di = -1;
          for (let k = 4; k < a.length; k++) {
            if (a[k] instanceof Ref && si < 0) si = k;
            else if (Array.isArray(a[k]) && a[k].some((x) => x instanceof Ref) && di < 0) di = k;
          }
          if (si < 0 || di < 0) continue;
          sp = { src: si, dst: di, pred: "related" };
        }
        const srcs = refIds(a[sp.src]);
        const dsts = refIds(a[sp.dst]);
        const prov = { rel: relCls, relId: id };

        if (type === "IFCRELDEFINESBYPROPERTIES") {
          for (const d of dsts) {
            if (!m.psets.has(d)) continue;
            for (const s of srcs) {
              const n = m.nodes.get("#" + s);
              if (n) n.psetIds.push(d);
              else m.unresolved++;
            }
          }
          continue;
        }
        if (sp.special === "material") {
          for (const d of dsts) {
            const mats = resolveMaterial(d);
            for (const s of srcs) {
              const n = m.nodes.get("#" + s);
              if (!n) continue;
              n.materials = (n.materials || []).concat(mats.map((x) => ({ ...x, name: m.nodes.get(x.id)?.label })));
              for (const x of mats) m.addEdge(n.id, x.id, "hasMaterial", { ...prov, layer: x.layer, thickness: x.thickness, fraction: x.fraction });
            }
          }
          continue;
        }
        if (sp.special === "classification") {
          for (const d of dsts) {
            const ck = ensureClassRef(d);
            if (!ck) continue;
            for (const s of srcs) {
              const n = m.nodes.get("#" + s);
              if (!n) continue;
              (n.classRefs = n.classRefs || []).push(ck);
              m.addEdge(n.id, ck, "hasClassification", prov);
            }
          }
          continue;
        }
        if (sp.special === "document") {
          for (const d of dsts) {
            const dk = ensureDocument(d);
            if (!dk) continue;
            for (const s of srcs) m.addEdge("#" + s, dk, "hasDocument", prov);
          }
          continue;
        }
        const attrs = { ...prov };
        if (sp.special === "boundary") {
          attrs.physical = str(a[7]);
          attrs.side = str(a[8]);
        } else if (sp.special === "path") {
          attrs.relatingConnection = str(a[10]);
          attrs.relatedConnection = str(a[9]);
        }
        for (const s of srcs) {
          for (const d of dsts) {
            let pred = sp.pred;
            if (pred === "nests" && m.nodes.get("#" + d)?.kind === "Port") pred = "hasPort";
            m.addEdge("#" + s, "#" + d, pred, attrs);
            if (sp.special === "boundary") boundaries.push({ space: "#" + s, element: "#" + d, side: attrs.side, physical: attrs.physical });
          }
        }
        if (++relDone % 20000 === 0) await progress(0.5 + 0.3 * (relDone / totalRel), `Linked ${relDone.toLocaleString()} relationships`);
      }
    }
    m.meta.relationships = relStats;
    await progress(0.82, "Running inference rules");

    m.buildIndex();
    infer(m, boundaries);

    // Providers for inspector
    m.rawProvider = (n) => (n.ifcId != null ? step.line(n.ifcId) : "");
    m.attrProvider = (n) => {
      if (n.ifcId == null) return [];
      const a = step.args(n.ifcId);
      const names = S.attributeNames(n.cls);
      return a.map((v, i) => ({ name: names[i] || `Attribute ${i}`, value: KG.Step.display(v) }));
    };

    await progress(0.9, "Computing catalogs");
    m.finalize();
    m.meta.unresolved = m.unresolved;
    step.cache.clear();
    await progress(1, "Done");
    return m;
  }

  /** Semantic inference over asserted relationships. */
  function infer(m, boundaries) {
    // 1. Hosting: host —hasOpening→ opening ←fills— filler  ⇒ host —hosts→ filler
    for (const e of m.edges.slice()) {
      if (e.pred !== "fills") continue;
      const opening = e.t;
      for (const h of m.inE(opening, "hasOpening")) {
        m.addEdge(h.s, e.s, "hosts", { via: opening, rule: "IfcRelVoidsElement ∘ IfcRelFillsElement" });
        const filler = m.nodes.get(e.s);
        if (filler) filler.hostId = h.s;
      }
    }
    m.buildIndex();

    // 2. MEP flow: element —hasPort→ p1 —portConnectedTo→ p2 ←hasPort— element
    const ownerOf = (port) => {
      const h = m.inE(port, "hasPort")[0];
      return h ? h.s : null;
    };
    for (const e of m.edges.slice()) {
      if (e.pred !== "portConnectedTo") continue;
      const a = ownerOf(e.s);
      const b = ownerOf(e.t);
      if (!a || !b || a === b) continue;
      const da = m.nodes.get(e.s)?.flowDirection;
      const db = m.nodes.get(e.t)?.flowDirection;
      if (da === "SINK" || db === "SOURCE") m.addEdge(b, a, "feeds", { viaPorts: [e.t, e.s], rule: "Port flow direction" });
      else m.addEdge(a, b, "feeds", { viaPorts: [e.s, e.t], rule: da === "SOURCE" || db === "SINK" ? "Port flow direction" : "Port connection (direction undefined)", directed: da === "SOURCE" || db === "SINK" });
    }

    // 3. Space topology from space boundaries
    if (boundaries.length) {
      const byElement = new Map();
      for (const b of boundaries) {
        let l = byElement.get(b.element);
        if (!l) byElement.set(b.element, (l = []));
        l.push(b);
      }
      // Doors give horizontal access; stairs, ramps and lifts give vertical access.
      const isAccess = (id) => {
        const n = m.nodes.get(id);
        if (!n) return false;
        if (["IfcDoor", "IfcStair", "IfcStairFlight", "IfcRamp", "IfcRampFlight", "IfcTransportElement"].some((c) => S.isA(n.cls, c))) return true;
        if (n.kind === "Opening") return m.inE(id, "fills").some((x) => S.isA(m.nodes.get(x.s)?.cls || "", "IfcDoor"));
        return false;
      };
      let exteriorNeeded = false;
      for (const [el, list] of byElement) {
        const spaces = [...new Set(list.map((b) => b.space))];
        const door = isAccess(el);
        for (let i = 0; i < spaces.length; i++) {
          for (let j = i + 1; j < spaces.length; j++) {
            m.addEdge(spaces[i], spaces[j], door ? "accessTo" : "adjacentTo", { via: el, rule: door ? "Shared door / stair / ramp boundary" : "Shared IfcRelSpaceBoundary element" });
          }
        }
        if (door && list.some((b) => /EXTERNAL/.test(b.side || ""))) {
          exteriorNeeded = true;
          for (const b of list) if (/EXTERNAL/.test(b.side || "")) b.exteriorDoor = true;
        }
      }
      if (exteriorNeeded) {
        m.addNode({ id: "exterior", kind: "Exterior", cls: "Exterior", name: "Exterior", label: "Exterior", virtual: true });
        for (const b of boundaries) if (b.exteriorDoor) m.addEdge("exterior", b.space, "accessTo", { via: b.element, rule: "External door boundary" });
      }
    }
    m.buildIndex();
  }

  function round(v, d = 2) {
    if (typeof v !== "number") return v;
    const f = Math.pow(10, d);
    return Math.round(v * f) / f;
  }

  KG.Model = { KGModel, buildFromIFC, infer, round, tick };
})((window.KG = window.KG || {}));
