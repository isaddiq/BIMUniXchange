/* BIMUniXchange · BIM Knowledge Graph Studio
 * csv-adapter.js — maps BIMUniXchange Revit / ArchiCAD CSV exports (and any
 * CSV with IFC_* columns) onto the same semantic model used for IFC files.
 */
(function (KG) {
  "use strict";
  const S = KG.Schema;

  const IDENTITY = new Set([
    "Element_Id", "Element_GUID", "IFC_GlobalId", "IFC_Identification", "IFC_EntityType", "IFC_PredefinedType",
    "IFC_Definition", "Element_Name", "Element_Type", "IFC_Hierarchy", "IFC_Project", "IFC_Site", "IFC_Building",
    "IFC_Storey", "IFC_SchemaVersion", "IFC_PropertySets", "IFC_QuantitySets", "Revit_Category", "ArchiCAD_Category",
    "Family and Type", "Element_Color", "Material", "Material_Name", "Materials",
  ]);
  const SKIP_CATEGORIES = new Set(["Views", "Cameras", "Sheets", "Schedules", "Viewports"]);

  function splitList(v) {
    return String(v || "")
      .split(/[;|]/)
      .map((s) => s.trim())
      .filter(Boolean);
  }
  // ArchiCAD exports enum values as "NonLocalizedValueEnumId {'nonLocalizedValue': 'New', …}"
  function clean(v) {
    const m = /nonLocalizedValue'\s*:\s*'([^']*)'/.exec(v);
    return m ? m[1] : v;
  }
  function typed(v) {
    if (v === "") return null;
    const t = clean(v.trim());
    if (/^(true|false)$/i.test(t)) return { value: /^true$/i.test(t), datatype: "BOOLEAN" };
    if (/^[-+]?\d*\.?\d+(e[-+]?\d+)?$/i.test(t)) return { value: parseFloat(t), datatype: "REAL" };
    return { value: t, datatype: "STRING" };
  }

  async function buildFromCSV(text, opts = {}) {
    const progress = opts.onProgress || (async () => {});
    const rows = d3.csvParseRows(text.replace(/^﻿/, ""));
    if (rows.length < 2) throw new Error("CSV file has no data rows");
    const header = rows[0].map((h) => h.trim());
    const col = {};
    header.forEach((h, i) => (col[h] = i));
    const get = (r, name) => (col[name] !== undefined ? clean((r[col[name]] || "").trim()) : "");
    const source = col.Revit_Category !== undefined ? "Revit" : col.ArchiCAD_Category !== undefined ? "ArchiCAD" : "CSV";
    await progress(0.15, `Parsed ${rows.length - 1} rows`);

    const m = new KG.Model.KGModel("csv");
    m.meta.fileName = opts.fileName || "export.csv";
    m.meta.fileSize = opts.fileSize || text.length;
    m.meta.schema = get(rows[1], "IFC_SchemaVersion") || "IFC (via CSV)";
    m.meta.application = `${source} CSV export`;
    m.meta.header = { originatingSystem: source, description: "BIMUniXchange tabular export" };
    m.meta.entityCount = rows.length - 1;
    m.meta.csvSource = source;
    m.meta.columns = header.length;

    // Column groups
    const propCols = [];
    const relCols = [];
    header.forEach((h, i) => {
      if (IDENTITY.has(h) || /^MaterialLayer_\d+_/.test(h)) return;
      if (/^Relationships\./.test(h)) relCols.push([i, h.slice(14)]);
      else if (h.includes(".")) {
        const dot = h.indexOf(".");
        propCols.push([i, h.slice(0, dot), h.slice(dot + 1)]);
      } else propCols.push([i, `${source} Parameters`, h]);
    });
    const layerCols = header
      .map((h, i) => [h.match(/^MaterialLayer_(\d+)_(Name|Thickness)$/), i])
      .filter(([mm]) => mm)
      .map(([mm, i]) => ({ layer: +mm[1], field: mm[2], i }));

    // Some exports put element elevations into the storey column — detect and ignore
    const storeyCol = ["IFC_Storey", "Constraints.Reference_Level", "Level", "Reference Level", "Base Level"].find((c) => col[c] !== undefined && rows.slice(1, 400).some((r) => (r[col[c]] || "").trim()));
    let useStorey = !!storeyCol;
    if (storeyCol) {
      const vals = rows.slice(1).map((r) => (r[col[storeyCol]] || "").trim()).filter(Boolean);
      const distinct = new Set(vals);
      const numeric = vals.filter((v) => /^[-+]?\d*\.?\d+$/.test(v)).length / Math.max(1, vals.length);
      if (numeric > 0.6 && distinct.size > 25) {
        useStorey = false;
        m.meta.warnings.push(`Column “${storeyCol}” holds ${distinct.size} numeric values (element elevations, not storey names), so storeys were not derived.`);
      }
    }

    // Spatial skeleton
    const project = m.addNode({ id: "proj", kind: "Project", cls: "IfcProject", name: get(rows[1], "IFC_Project") || opts.fileName || "Project" });
    const ensureSpatial = (id, cls, name, parent) => {
      if (!m.nodes.has(id)) {
        m.addNode({ id, kind: "Spatial", cls, name });
        m.addEdge(parent, id, "aggregates", { rel: "CSV hierarchy" });
      }
      return id;
    };
    const catSystem = col.Revit_Category !== undefined || col.ArchiCAD_Category !== undefined
      ? m.addNode({ id: "cs:" + source, kind: "ClassSystem", cls: "IfcClassification", name: `${source} Categories` })
      : null;

    const psetSig = new Map();
    let psetSeq = 1;
    const sharedPset = (name, kind, props) => {
      const sig = name + "\u0001" + props.map((p) => p.name + "=" + p.value).join("\u0002");
      let id = psetSig.get(sig);
      if (id === undefined) {
        id = "ps" + psetSeq++;
        psetSig.set(sig, id);
        m.psets.set(id, { id, name, kind, props });
      }
      return id;
    };

    const byTag = new Map();
    const seenIds = new Map();
    const duplicates = [];
    const pendingRels = [];
    const pendingHosts = [];
    const storeyElev = new Map();

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row.length || (row.length === 1 && !row[0])) continue;
      const category = get(row, "Revit_Category") || get(row, "ArchiCAD_Category");
      if (SKIP_CATEGORIES.has(category)) continue;

      let cls = S.canonical(get(row, "IFC_EntityType") || "IfcBuildingElementProxy");
      if (!cls.startsWith("Ifc")) cls = "IfcBuildingElementProxy";
      let kind = S.kindOf(cls);
      if (kind === "Other" || kind === "Type") kind = "Element";
      const guid = get(row, "IFC_GlobalId") || get(row, "Element_GUID");
      const tag = get(row, "Element_Id");
      let id = "e:" + (get(row, "Element_GUID") || guid || tag || r);
      if (seenIds.has(id)) {
        duplicates.push(id);
        id = id + "~" + r;
      }
      seenIds.set(id, r);

      // Spatial placement
      const siteName = get(row, "IFC_Site");
      const bldgName = get(row, "IFC_Building");
      const storeyName = useStorey ? get(row, "IFC_Storey") || get(row, "Constraints.Reference_Level") || get(row, "Level") || get(row, "Reference Level") || get(row, "Base Level") : "";
      let parent = project.id;
      if (siteName) parent = ensureSpatial("site:" + siteName, "IfcSite", siteName, parent);
      if (bldgName) parent = ensureSpatial("bldg:" + bldgName, "IfcBuilding", bldgName, parent);
      if (storeyName) {
        if (!bldgName && !siteName) parent = ensureSpatial("bldg:default", "IfcBuilding", "Building", parent);
        const sid = ensureSpatial("st:" + storeyName, "IfcBuildingStorey", storeyName, parent);
        const elev = parseFloat(get(row, "Constraints.Reference_Level_Elevation"));
        if (!isNaN(elev) && !storeyElev.has(sid)) storeyElev.set(sid, elev);
        parent = sid;
      }

      const name = get(row, "Element_Name") || get(row, "Family and Type") || get(row, "Element_Type");
      const pt = get(row, "IFC_PredefinedType");
      const n = m.addNode({
        id, kind, cls, guid, tag, name,
        label: name || `${cls.replace(/^Ifc/, "")} ${tag}`,
        objectType: get(row, "Family and Type") || get(row, "Element_Type"),
        predefinedType: pt && pt !== "NOTDEFINED" ? pt : "",
        classDefinition: get(row, "IFC_Definition"),
        sourceCategory: category,
        family: get(row, "Family") || (get(row, "Family and Type").includes(":") ? get(row, "Family and Type").split(":")[0].trim() : ""),
        typeName: get(row, "Type") || (get(row, "Family and Type").split(":").pop() || "").trim(),
        layer: get(row, "Model_View.Layer_Name") || get(row, "Layer"),
        composite: get(row, "Construction.Composite_Name"),
        phaseCreated: get(row, "Phase Created") || get(row, "CategoryPropertyDefinitionGroup.Renovation_Status"),
        phaseDemolished: get(row, "Phase Demolished"),
        csvRow: r,
      });
      const sf = get(row, "Structural.StructuralFunction") || get(row, "CategoryPropertyDefinitionGroup.Structural_Function");
      if (/^LoadBearing/i.test(sf)) n.loadBearingHint = true;
      else if (/^NonLoadBearing/i.test(sf)) n.loadBearingHint = false;
      const hostTag = get(row, "Host Id");
      if (hostTag && hostTag !== "-1") pendingHosts.push([id, hostTag]);
      if (tag) byTag.set(tag, id);
      if (guid) byTag.set(guid, id);
      m.addEdge(parent, id, kind === "Spatial" ? "aggregates" : "contains", { rel: "CSV hierarchy" });

      // Type
      const famType = get(row, "Family and Type");
      if (famType) {
        const tid = "t:" + famType;
        if (!m.nodes.has(tid)) {
          const tcls = S.canonical(cls.toUpperCase().replace(/STANDARDCASE$/, "") + "TYPE");
          m.addNode({ id: tid, kind: "Type", cls: tcls.endsWith("Type") ? tcls : "IfcTypeObject", name: famType.split(":").pop().trim(), objectType: famType, elementClass: cls });
        }
        m.addEdge(id, tid, "isTypedBy", { rel: "Family and Type" });
      }

      // Category as classification
      if (catSystem && category) {
        const cid = "cat:" + category;
        if (!m.nodes.has(cid)) {
          m.addNode({ id: cid, kind: "Classification", cls: "IfcClassificationReference", name: category, code: category, scheme: catSystem.label });
          m.addEdge(cid, catSystem.id, "inSystem");
        }
        m.addEdge(id, cid, "hasClassification", { rel: `${source} category` });
        n.classRefs = [cid];
      }

      // Materials (layers first, then flat list)
      const layers = new Map();
      for (const lc of layerCols) {
        const v = (row[lc.i] || "").trim();
        if (!v) continue;
        const L = layers.get(lc.layer) || {};
        L[lc.field] = v;
        layers.set(lc.layer, L);
      }
      const mats = [];
      for (const [li, L] of [...layers].sort((a, b) => a[0] - b[0])) if (L.Name) mats.push({ name: L.Name, layer: li + 1, thickness: parseFloat(L.Thickness) || null });
      for (const mn of splitList(get(row, "Materials") || get(row, "Material_Name") || get(row, "Material"))) if (!mats.some((x) => x.name === mn)) mats.push({ name: mn });
      // ArchiCAD composites / building materials act as the material definition
      if (!mats.length && n.composite) mats.push({ name: n.composite, composite: true });
      if (mats.length) {
        n.materials = [];
        for (const mt of mats) {
          const mid = "mat:" + mt.name;
          if (!m.nodes.has(mid)) m.addNode({ id: mid, kind: "Material", cls: "IfcMaterial", name: mt.name });
          m.addEdge(id, mid, "hasMaterial", { layer: mt.layer, thickness: mt.thickness });
          n.materials.push({ id: mid, name: mt.name, layer: mt.layer, thickness: mt.thickness });
        }
      }

      // Systems
      const sysNames = splitList((get(row, "System Name") || "").replace(/,/g, ";"));
      const sysClass = get(row, "System Classification") || get(row, "System Type");
      for (const sn of sysNames) {
        const sid = "sys:" + sn;
        if (!m.nodes.has(sid)) m.addNode({ id: sid, kind: "System", cls: "IfcDistributionSystem", name: sn, predefinedType: sysClass.toUpperCase(), objectType: sysClass });
        m.addEdge(id, sid, "memberOf", { rel: "System Name" });
      }
      if (sysClass) n.systemHint = sysClass;

      // Properties → shared, de-duplicated property sets
      const groups = new Map();
      for (const [i, pset, prop] of propCols) {
        const raw = row[i];
        if (raw === undefined || raw === "") continue;
        const tv = typed(raw);
        if (!tv) continue;
        let g = groups.get(pset);
        if (!g) groups.set(pset, (g = []));
        g.push({ name: prop, value: tv.value, datatype: tv.datatype });
      }
      for (const [pset, props] of groups) n.psetIds.push(sharedPset(pset, /^Qto_|^Dimensions$/.test(pset) ? "Q" : "P", props));

      for (const [i, relName] of relCols) {
        const v = (row[i] || "").trim();
        if (v) pendingRels.push([id, relName, v]);
      }
      if (r % 2000 === 0) await progress(0.15 + 0.6 * (r / rows.length), `Mapped ${r.toLocaleString()} rows`);
    }

    for (const [sid, elev] of storeyElev) {
      const s = m.nodes.get(sid);
      s.elevation = elev;
      s.sub = `Elev. ${KG.Model.round(elev)}`;
    }

    // Revit "Host Id" → hosting (doors / windows / fixtures hosted by walls, floors, ceilings)
    for (const [id, hostTag] of pendingHosts) {
      const hid = byTag.get(hostTag);
      if (!hid) continue;
      const e = m.addEdge(hid, id, "hosts", { rel: "Host Id" });
      if (e) {
        e.inferred = false; // asserted by the authoring tool, not derived
        m.nodes.get(id).hostId = hid;
      }
    }

    // Relationship columns → edges
    for (const [id, relName, v] of pendingRels) {
      const targets = v.split(/[;,|\s]+/).filter(Boolean);
      const pred = /group|system/i.test(relName) ? "memberOf" : /host/i.test(relName) ? "hosts" : "connectedTo";
      for (const t of targets) {
        let tid = byTag.get(t);
        if (!tid && pred === "memberOf") {
          tid = "grp:" + t;
          if (!m.nodes.has(tid)) m.addNode({ id: tid, kind: "Group", cls: "IfcGroup", name: t });
        }
        if (tid) m.addEdge(id, tid, pred, { rel: "Relationships." + relName });
      }
    }

    m.meta.duplicateIds = duplicates;
    m.meta.typeHistogram = [...m.nodes.values()].reduce((acc, n) => acc.set(n.cls, (acc.get(n.cls) || 0) + 1), new Map());
    m.meta.typeHistogram = [...m.meta.typeHistogram].sort((a, b) => b[1] - a[1]);
    m.rawProvider = (n) => {
      if (n.csvRow == null) return "";
      const row = rows[n.csvRow];
      const lines = [];
      header.forEach((h, i) => {
        if (row[i] !== undefined && row[i] !== "") lines.push(`${h}: ${row[i]}`);
      });
      return lines.join("\n");
    };
    m.attrProvider = (n) => [
      { name: "GlobalId", value: n.guid || "" },
      { name: "Element_Id", value: n.tag || "" },
      { name: "Name", value: n.name || "" },
      { name: "ObjectType", value: n.objectType || "" },
      { name: "PredefinedType", value: n.predefinedType || "" },
      { name: "Source category", value: n.sourceCategory || "" },
      { name: "IFC definition", value: n.classDefinition || "" },
    ].filter((a) => a.value);
    await progress(0.85, "Computing catalogs");
    m.buildIndex();
    m.finalize();
    await progress(1, "Done");
    return m;
  }

  KG.CSV = { buildFromCSV };
})((window.KG = window.KG || {}));
