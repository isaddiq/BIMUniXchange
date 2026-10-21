/* BIMUniXchange · BIM Knowledge Graph Studio
 * health.js — information-quality audit of the semantic model
 * (identity, spatial structure, typing, materials, property completeness,
 * MEP topology, space boundaries). Each check returns affected node ids so
 * the result can be visualised as a subgraph.
 */
(function (KG) {
  "use strict";
  const S = KG.Schema;

  const COMMON_PSET = ["Wall", "Slab", "Beam", "Column", "Door", "Window", "Roof", "Stair", "StairFlight", "Railing", "Covering", "CurtainWall", "Member", "Plate", "Footing", "Pile", "Ramp", "RampFlight", "BuildingElementProxy", "Chimney", "ShadingDevice"];

  function baseName(cls) {
    return cls.replace(/^Ifc/, "").replace(/(StandardCase|ElementedCase)$/, "");
  }

  function run(model) {
    const nodes = [...model.nodes.values()];
    const elements = nodes.filter((n) => n.kind === "Element" && !n.virtual);
    const physical = elements.filter((n) => n.domain !== "Annotation & Grid");
    const checks = [];
    const add = (c) => checks.push({ ...c, failed: c.count ?? c.affected.length });

    // Identity
    const withGuid = nodes.filter((n) => n.guid);
    const byGuid = new Map();
    for (const n of withGuid) {
      const l = byGuid.get(n.guid);
      if (l) l.push(n.id);
      else byGuid.set(n.guid, [n.id]);
    }
    const dups = [...byGuid.values()].filter((l) => l.length > 1).flat();
    add({ id: "dup-guid", group: "Identity", severity: "error", title: "Duplicate GlobalIds", description: "IfcRoot.GlobalId must be unique across the model (IFC schema rule UR1). Duplicates break change tracking, BCF issues and model federation.", affected: dups, population: withGuid.length });
    const badGuid = model.source === "ifc" ? withGuid.filter((n) => !/^[0-9A-Za-z_$]{22}$/.test(n.guid)).map((n) => n.id) : [];
    if (model.source === "ifc") add({ id: "guid-format", group: "Identity", severity: "error", title: "Malformed GlobalIds", description: "GlobalIds must be 22 characters from the IFC base-64 alphabet (0-9, A-Z, a-z, _ and $).", affected: badGuid, population: withGuid.length });
    if (model.meta.unresolved) add({ id: "unresolved", group: "Identity", severity: "warning", title: "Unresolved relationship targets", description: "Relationships reference instances that are missing or are not IfcRoot objects, so those links could not be added to the graph.", affected: [], count: model.meta.unresolved, population: model.meta.unresolved + model.edges.length });

    // Spatial structure
    const uncontained = physical.filter((n) => !n.storey && model.spatialPath(n.id).length === 0).map((n) => n.id);
    add({ id: "no-container", group: "Spatial structure", severity: "warning", title: "Elements outside the spatial structure", description: "Elements that are neither contained in a spatial element (IfcRelContainedInSpatialStructure) nor aggregated into another element.", affected: uncontained, population: physical.length });
    const storeyUse = new Map();
    for (const n of elements) if (n.storey) storeyUse.set(n.storey, (storeyUse.get(n.storey) || 0) + 1);
    const storeys = nodes.filter((n) => n.cls === "IfcBuildingStorey");
    add({ id: "empty-storey", group: "Spatial structure", severity: "info", title: "Empty storeys", description: "Building storeys with no contained elements.", affected: storeys.filter((s) => !storeyUse.get(s.id)).map((s) => s.id), population: storeys.length });
    const spaces = nodes.filter((n) => n.cls === "IfcSpace");
    if (model.source === "ifc" && spaces.length) {
      add({ id: "space-no-boundary", group: "Spatial structure", severity: "info", title: "Spaces without space boundaries", description: "IfcRelSpaceBoundary is required for energy simulation hand-over and for deriving space adjacency and circulation graphs.", affected: spaces.filter((s) => !model.outE(s.id, "boundedBy").length).map((s) => s.id), population: spaces.length });
      add({ id: "space-no-area", group: "Spatial structure", severity: "info", title: "Spaces without floor area quantity", description: "No NetFloorArea / GrossFloorArea quantity found (Qto_SpaceBaseQuantities).", affected: spaces.filter((s) => ![...model.propMap(s).keys()].some((k) => /FloorArea$/i.test(k))).map((s) => s.id), population: spaces.length });
    }

    // Typing & classification
    add({ id: "no-type", group: "Semantics", severity: "info", title: "Elements without a type object", description: "Elements not related to an IfcTypeObject via IfcRelDefinesByType; shared type data cannot be inherited.", affected: physical.filter((n) => !n.typeId).map((n) => n.id), population: physical.length });
    add({ id: "proxy", group: "Semantics", severity: "warning", title: "Unclassified proxy elements", description: "IfcBuildingElementProxy instances carry no semantic class; they reduce the value of the model for automated checking and quantity take-off.", affected: physical.filter((n) => n.cls === "IfcBuildingElementProxy").map((n) => n.id), population: physical.length });
    add({ id: "unnamed", group: "Semantics", severity: "info", title: "Unnamed elements", description: "Elements with an empty Name attribute.", affected: physical.filter((n) => !n.name).map((n) => n.id), population: physical.length });
    add({ id: "no-classification", group: "Semantics", severity: "info", title: "Elements without classification", description: "No IfcClassificationReference (e.g. Uniclass, OmniClass, CCI) is associated with the element.", affected: physical.filter((n) => !(n.classRefs && n.classRefs.length)).map((n) => n.id), population: physical.length });
    const types = nodes.filter((n) => n.kind === "Type");
    add({ id: "orphan-type", group: "Semantics", severity: "info", title: "Unused type objects", description: "Type objects that no occurrence references.", affected: types.filter((t) => !model.inE(t.id, "isTypedBy").length).map((t) => t.id), population: types.length });

    // Materials & properties
    const matPop = physical.filter((n) => !S.isA(n.cls, "IfcDistributionControlElement"));
    add({ id: "no-material", group: "Properties & materials", severity: "warning", title: "Elements without material", description: "No IfcRelAssociatesMaterial on the element or its type — blocks LCA / embodied-carbon and cost workflows.", affected: matPop.filter((n) => !(n.materials?.length || n.materialsInherited?.length)).map((n) => n.id), population: matPop.length });
    add({ id: "no-pset", group: "Properties & materials", severity: "warning", title: "Elements without any property set", description: "Neither the occurrence nor its type carry property or quantity sets.", affected: physical.filter((n) => !model.propertySets(n).length).map((n) => n.id), population: physical.length });
    if (model.source === "ifc") {
      const commonPop = physical.filter((n) => COMMON_PSET.includes(baseName(n.cls)));
      const missingCommon = commonPop.filter((n) => {
        const expected = "Pset_" + baseName(n.cls) + "Common";
        return !model.propertySets(n).some((p) => p.name === expected);
      });
      add({ id: "missing-common", group: "Properties & materials", severity: "info", title: "Missing Pset_*Common", description: "Occurrences lacking their standard common property set (e.g. Pset_WallCommon for IfcWall).", affected: missingCommon.map((n) => n.id), population: commonPop.length });
    }
    const envelopePop = physical.filter((n) => ["Wall", "Window", "Door", "Slab", "Roof", "CurtainWall"].includes(baseName(n.cls)));
    const pmHas = (n, re) => [...model.propMap(n).keys()].some((k) => re.test(k));
    add({ id: "no-isexternal", group: "Properties & materials", severity: "info", title: "Envelope elements without IsExternal", description: "IsExternal is needed to derive the thermal envelope for energy analysis.", affected: envelopePop.filter((n) => !pmHas(n, /\.IsExternal$/i)).map((n) => n.id), population: envelopePop.length });
    const firePop = physical.filter((n) => ["Wall", "Door", "Slab", "Column", "Beam"].includes(baseName(n.cls)));
    add({ id: "no-firerating", group: "Properties & materials", severity: "info", title: "Fire-relevant elements without FireRating", description: "Walls, doors, slabs, columns and beams without a FireRating property (code-compliance checking).", affected: firePop.filter((n) => !pmHas(n, /\.FireRating$/i)).map((n) => n.id), population: firePop.length });

    // MEP topology
    const ports = nodes.filter((n) => n.kind === "Port");
    if (ports.length) add({ id: "port-unconnected", group: "MEP topology", severity: "warning", title: "Unconnected distribution ports", description: "Ports without IfcRelConnectsPorts — gaps in the MEP network topology.", affected: ports.filter((p) => !model.outE(p.id, "portConnectedTo").length && !model.inE(p.id, "portConnectedTo").length).map((p) => p.id), population: ports.length });
    const flow = physical.filter((n) => S.isA(n.cls, "IfcDistributionFlowElement"));
    if (flow.length) add({ id: "flow-no-system", group: "MEP topology", severity: "warning", title: "Distribution elements without system", description: "Flow elements not assigned to an IfcSystem / IfcDistributionSystem.", affected: flow.filter((n) => !model.outE(n.id, "memberOf").length).map((n) => n.id), population: flow.length });
    const groups = nodes.filter((n) => n.kind === "System" || n.kind === "Zone" || n.kind === "Group");
    if (groups.length) add({ id: "empty-group", group: "MEP topology", severity: "info", title: "Empty systems / zones / groups", description: "Groups with no assigned members.", affected: groups.filter((g) => !model.inE(g.id, "memberOf").length).map((g) => g.id), population: groups.length });

    const W = { error: 3, warning: 2, info: 1 };
    let num = 0, den = 0;
    for (const c of checks) {
      if (!c.population) continue;
      num += W[c.severity] * (c.failed / c.population);
      den += W[c.severity];
    }
    const score = den ? Math.round(100 * (1 - num / den)) : 100;
    return { checks, score };
  }

  KG.Health = { run };
})((window.KG = window.KG || {}));
