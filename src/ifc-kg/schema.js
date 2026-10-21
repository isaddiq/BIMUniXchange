/* BIMUniXchange · BIM Knowledge Graph Studio
 * schema.js — IFC class taxonomy, attribute names, relationship semantics,
 * discipline (domain) rules and ontology alignment (ifcOWL, BOT, Brick).
 */
(function (KG) {
  "use strict";

  // ── IFC class taxonomy (IFC2x3 / IFC4 / IFC4x3 union, parent → children) ──
  const TREE = {
    IfcRoot: ["IfcObjectDefinition", "IfcPropertyDefinition", "IfcRelationship"],
    IfcObjectDefinition: ["IfcObject", "IfcTypeObject", "IfcContext"],
    IfcContext: ["IfcProject", "IfcProjectLibrary"],
    IfcObject: ["IfcProduct", "IfcProcess", "IfcResource", "IfcControl", "IfcGroup", "IfcActor"],
    IfcProduct: ["IfcElement", "IfcSpatialElement", "IfcPort", "IfcAnnotation", "IfcPositioningElement", "IfcProxy", "IfcStructuralItem", "IfcStructuralActivity"],
    IfcPositioningElement: ["IfcGrid", "IfcAlignment", "IfcLinearPositioningElement", "IfcReferent"],
    IfcSpatialElement: ["IfcSpatialStructureElement", "IfcSpatialZone", "IfcExternalSpatialStructureElement"],
    IfcSpatialStructureElement: ["IfcSite", "IfcFacility", "IfcBuildingStorey", "IfcSpace", "IfcFacilityPart"],
    IfcFacility: ["IfcBuilding", "IfcBridge", "IfcMarineFacility", "IfcRailway", "IfcRoad"],
    IfcFacilityPart: ["IfcBridgePart", "IfcFacilityPartCommon", "IfcMarinePart", "IfcRailwayPart", "IfcRoadPart"],
    IfcExternalSpatialStructureElement: ["IfcExternalSpatialElement"],
    IfcElement: ["IfcBuiltElement", "IfcBuildingElement", "IfcCivilElement", "IfcDistributionElement", "IfcElementAssembly", "IfcElementComponent", "IfcFeatureElement", "IfcFurnishingElement", "IfcGeographicElement", "IfcTransportElement", "IfcVirtualElement", "IfcElectricalElement", "IfcEquipmentElement"],
    IfcBuiltElement: ["IfcBeam", "IfcBearing", "IfcBuildingElementProxy", "IfcChimney", "IfcColumn", "IfcCourse", "IfcCovering", "IfcCurtainWall", "IfcDeepFoundation", "IfcDoor", "IfcEarthworksElement", "IfcFooting", "IfcKerb", "IfcMember", "IfcMooringDevice", "IfcNavigationElement", "IfcPavement", "IfcPlate", "IfcRail", "IfcRailing", "IfcRamp", "IfcRampFlight", "IfcRoof", "IfcShadingDevice", "IfcSlab", "IfcStair", "IfcStairFlight", "IfcTrackElement", "IfcWall", "IfcWindow"],
    IfcDeepFoundation: ["IfcCaissonFoundation", "IfcPile"],
    IfcEarthworksElement: ["IfcEarthworksFill", "IfcReinforcedSoil"],
    IfcWall: ["IfcWallStandardCase", "IfcWallElementedCase"],
    IfcSlab: ["IfcSlabStandardCase", "IfcSlabElementedCase"],
    IfcBeam: ["IfcBeamStandardCase"],
    IfcColumn: ["IfcColumnStandardCase"],
    IfcMember: ["IfcMemberStandardCase"],
    IfcPlate: ["IfcPlateStandardCase"],
    IfcDoor: ["IfcDoorStandardCase"],
    IfcWindow: ["IfcWindowStandardCase"],
    IfcDistributionElement: ["IfcDistributionControlElement", "IfcDistributionFlowElement"],
    IfcDistributionControlElement: ["IfcActuator", "IfcAlarm", "IfcController", "IfcFlowInstrument", "IfcProtectiveDeviceTrippingUnit", "IfcSensor", "IfcUnitaryControlElement"],
    IfcDistributionFlowElement: ["IfcDistributionChamberElement", "IfcEnergyConversionDevice", "IfcFlowController", "IfcFlowFitting", "IfcFlowMovingDevice", "IfcFlowSegment", "IfcFlowStorageDevice", "IfcFlowTerminal", "IfcFlowTreatmentDevice"],
    IfcEnergyConversionDevice: ["IfcAirToAirHeatRecovery", "IfcBoiler", "IfcBurner", "IfcChiller", "IfcCoil", "IfcCondenser", "IfcCooledBeam", "IfcCoolingTower", "IfcElectricGenerator", "IfcElectricMotor", "IfcEngine", "IfcEvaporativeCooler", "IfcEvaporator", "IfcHeatExchanger", "IfcHumidifier", "IfcMotorConnection", "IfcSolarDevice", "IfcTransformer", "IfcTubeBundle", "IfcUnitaryEquipment"],
    IfcFlowController: ["IfcAirTerminalBox", "IfcDamper", "IfcDistributionBoard", "IfcElectricDistributionBoard", "IfcElectricTimeControl", "IfcFlowMeter", "IfcProtectiveDevice", "IfcSwitchingDevice", "IfcValve"],
    IfcFlowFitting: ["IfcCableCarrierFitting", "IfcCableFitting", "IfcDuctFitting", "IfcJunctionBox", "IfcPipeFitting"],
    IfcFlowMovingDevice: ["IfcCompressor", "IfcFan", "IfcPump"],
    IfcFlowSegment: ["IfcCableCarrierSegment", "IfcCableSegment", "IfcDuctSegment", "IfcPipeSegment"],
    IfcFlowStorageDevice: ["IfcElectricFlowStorageDevice", "IfcTank"],
    IfcFlowTerminal: ["IfcAirTerminal", "IfcAudioVisualAppliance", "IfcCommunicationsAppliance", "IfcElectricAppliance", "IfcFireSuppressionTerminal", "IfcLamp", "IfcLightFixture", "IfcLiquidTerminal", "IfcMedicalDevice", "IfcOutlet", "IfcSanitaryTerminal", "IfcSignal", "IfcSpaceHeater", "IfcStackTerminal", "IfcWasteTerminal"],
    IfcFlowTreatmentDevice: ["IfcDuctSilencer", "IfcElectricFlowTreatmentDevice", "IfcFilter", "IfcInterceptor"],
    IfcElementComponent: ["IfcBuildingElementPart", "IfcDiscreteAccessory", "IfcFastener", "IfcMechanicalFastener", "IfcReinforcingElement", "IfcVibrationIsolator", "IfcVibrationDamper", "IfcImpactProtectionDevice", "IfcSign"],
    IfcReinforcingElement: ["IfcReinforcingBar", "IfcReinforcingMesh", "IfcTendon", "IfcTendonAnchor", "IfcTendonConduit"],
    IfcFeatureElement: ["IfcFeatureElementAddition", "IfcFeatureElementSubtraction", "IfcSurfaceFeature"],
    IfcFeatureElementAddition: ["IfcProjectionElement"],
    IfcFeatureElementSubtraction: ["IfcOpeningElement", "IfcVoidingFeature", "IfcEarthworksCut"],
    IfcOpeningElement: ["IfcOpeningStandardCase"],
    IfcFurnishingElement: ["IfcFurniture", "IfcSystemFurnitureElement"],
    IfcPort: ["IfcDistributionPort"],
    IfcGroup: ["IfcSystem", "IfcAsset", "IfcInventory", "IfcStructuralLoadGroup", "IfcStructuralResultGroup"],
    IfcSystem: ["IfcDistributionSystem", "IfcBuildingSystem", "IfcBuiltSystem", "IfcStructuralAnalysisModel", "IfcZone", "IfcElectricalCircuit"],
    IfcDistributionSystem: ["IfcDistributionCircuit"],
    IfcProcess: ["IfcTask", "IfcEvent", "IfcProcedure"],
    IfcResource: ["IfcConstructionResource"],
    IfcConstructionResource: ["IfcConstructionEquipmentResource", "IfcConstructionMaterialResource", "IfcConstructionProductResource", "IfcCrewResource", "IfcLaborResource", "IfcSubContractResource"],
    IfcControl: ["IfcActionRequest", "IfcCostItem", "IfcCostSchedule", "IfcPerformanceHistory", "IfcPermit", "IfcProjectOrder", "IfcWorkCalendar", "IfcWorkControl"],
    IfcWorkControl: ["IfcWorkPlan", "IfcWorkSchedule"],
    IfcActor: ["IfcOccupant"],
    IfcStructuralItem: ["IfcStructuralConnection", "IfcStructuralMember"],
    IfcStructuralConnection: ["IfcStructuralCurveConnection", "IfcStructuralPointConnection", "IfcStructuralSurfaceConnection"],
    IfcStructuralMember: ["IfcStructuralCurveMember", "IfcStructuralSurfaceMember"],
    IfcTypeObject: ["IfcTypeProduct", "IfcTypeProcess", "IfcTypeResource"],
    IfcTypeProduct: ["IfcElementType", "IfcSpatialElementType", "IfcDoorStyle", "IfcWindowStyle"],
    IfcPropertyDefinition: ["IfcPropertySetDefinition", "IfcPropertyTemplateDefinition"],
    IfcPropertySetDefinition: ["IfcPropertySet", "IfcElementQuantity", "IfcPreDefinedPropertySet"],
    IfcRelationship: ["IfcRelAssigns", "IfcRelAssociates", "IfcRelConnects", "IfcRelDeclares", "IfcRelDecomposes", "IfcRelDefines"],
    IfcRelDecomposes: ["IfcRelAggregates", "IfcRelNests", "IfcRelProjectsElement", "IfcRelVoidsElement"],
    IfcRelAssigns: ["IfcRelAssignsToActor", "IfcRelAssignsToControl", "IfcRelAssignsToGroup", "IfcRelAssignsToProcess", "IfcRelAssignsToProduct", "IfcRelAssignsToResource"],
    IfcRelAssignsToGroup: ["IfcRelAssignsToGroupByFactor"],
    IfcRelAssociates: ["IfcRelAssociatesApproval", "IfcRelAssociatesClassification", "IfcRelAssociatesConstraint", "IfcRelAssociatesDocument", "IfcRelAssociatesLibrary", "IfcRelAssociatesMaterial", "IfcRelAssociatesProfileDef"],
    IfcRelConnects: ["IfcRelConnectsElements", "IfcRelConnectsPortToElement", "IfcRelConnectsPorts", "IfcRelConnectsStructuralActivity", "IfcRelConnectsStructuralMember", "IfcRelContainedInSpatialStructure", "IfcRelCoversBldgElements", "IfcRelCoversSpaces", "IfcRelFillsElement", "IfcRelFlowControlElements", "IfcRelInterferesElements", "IfcRelReferencedInSpatialStructure", "IfcRelSequence", "IfcRelServicesBuildings", "IfcRelSpaceBoundary", "IfcRelPositions", "IfcRelAdheresToElement"],
    IfcRelConnectsElements: ["IfcRelConnectsPathElements", "IfcRelConnectsWithRealizingElements"],
    IfcRelSpaceBoundary: ["IfcRelSpaceBoundary1stLevel"],
    IfcRelSpaceBoundary1stLevel: ["IfcRelSpaceBoundary2ndLevel"],
    IfcRelDefines: ["IfcRelDefinesByObject", "IfcRelDefinesByProperties", "IfcRelDefinesByTemplate", "IfcRelDefinesByType"],
  };

  const PARENT = {}; // canonical → canonical parent
  const CANON = {}; // UPPER → canonical
  for (const [parent, kids] of Object.entries(TREE)) {
    CANON[parent.toUpperCase()] = parent;
    for (const k of kids) {
      CANON[k.toUpperCase()] = k;
      if (!PARENT[k]) PARENT[k] = parent;
    }
  }
  // IfcBuildingElement (≤IFC4) and IfcBuiltElement (IFC4x3) are the same concept.
  PARENT.IfcBuildingElement = "IfcElement";

  // Extra canonical names that only appear as leaves of other concepts
  [
    "IfcPropertySingleValue", "IfcPropertyEnumeratedValue", "IfcPropertyBoundedValue", "IfcPropertyListValue",
    "IfcPropertyTableValue", "IfcPropertyReferenceValue", "IfcComplexProperty", "IfcQuantityLength", "IfcQuantityArea",
    "IfcQuantityVolume", "IfcQuantityCount", "IfcQuantityWeight", "IfcQuantityTime", "IfcQuantityNumber",
    "IfcMaterial", "IfcMaterialLayer", "IfcMaterialLayerSet", "IfcMaterialLayerSetUsage", "IfcMaterialList",
    "IfcMaterialConstituent", "IfcMaterialConstituentSet", "IfcMaterialProfile", "IfcMaterialProfileSet",
    "IfcMaterialProfileSetUsage", "IfcClassification", "IfcClassificationReference", "IfcDocumentReference",
    "IfcDocumentInformation", "IfcOwnerHistory", "IfcApplication", "IfcSIUnit", "IfcConversionBasedUnit",
    "IfcUnitAssignment", "IfcDerivedUnit", "IfcPerson", "IfcOrganization", "IfcPersonAndOrganization",
    "IfcPostalAddress", "IfcLibraryReference", "IfcDoorLiningProperties", "IfcDoorPanelProperties",
    "IfcWindowLiningProperties", "IfcWindowPanelProperties", "IfcPhysicalComplexQuantity",
    // defined types used as property values
    "IfcLengthMeasure", "IfcPositiveLengthMeasure", "IfcNonNegativeLengthMeasure", "IfcAreaMeasure", "IfcVolumeMeasure",
    "IfcCountMeasure", "IfcReal", "IfcInteger", "IfcBoolean", "IfcLogical", "IfcLabel", "IfcText", "IfcIdentifier",
    "IfcPowerMeasure", "IfcThermalTransmittanceMeasure", "IfcVolumetricFlowRateMeasure", "IfcPlaneAngleMeasure",
    "IfcPositivePlaneAngleMeasure", "IfcMassMeasure", "IfcElectricCurrentMeasure", "IfcElectricVoltageMeasure",
    "IfcPressureMeasure", "IfcRotationalFrequencyMeasure", "IfcPositiveRatioMeasure", "IfcRatioMeasure",
    "IfcNormalisedRatioMeasure", "IfcThermodynamicTemperatureMeasure", "IfcTimeMeasure", "IfcDate", "IfcDateTime",
    "IfcMonetaryMeasure", "IfcIlluminanceMeasure", "IfcLuminousFluxMeasure", "IfcFrequencyMeasure", "IfcEnergyMeasure",
    "IfcForceMeasure", "IfcHeatFluxDensityMeasure", "IfcThermalConductivityMeasure", "IfcMassDensityMeasure",
    "IfcSoundPowerLevelMeasure", "IfcSoundPressureLevelMeasure", "IfcDescriptiveMeasure", "IfcMassFlowRateMeasure",
    "IfcLinearVelocityMeasure", "IfcElectricResistanceMeasure", "IfcLuminousIntensityMeasure", "IfcDuration",
  ].forEach((c) => (CANON[c.toUpperCase()] = c));

  function canonical(name) {
    if (!name) return "";
    const up = String(name).toUpperCase();
    if (CANON[up]) return CANON[up];
    // Type objects: IFCWALLTYPE → IfcWallType when the element class is known
    if (up.endsWith("TYPE") && CANON[up.slice(0, -4)]) return CANON[up.slice(0, -4)] + "Type";
    if (up.endsWith("STYLE") && CANON[up.slice(0, -5)]) return CANON[up.slice(0, -5)] + "Style";
    // Fallback prettifier: IFCFOOBAR → IfcFoobar
    if (up.startsWith("IFC")) return "Ifc" + up.charAt(3) + up.slice(4).toLowerCase();
    return name;
  }

  /** Inheritance chain, most specific first. */
  function chain(cls, schemaId) {
    const out = [];
    let c = cls;
    const legacy = schemaId && !/4X3|IFC4X/i.test(schemaId);
    if (isTypeClass(c)) {
      out.push(c);
      const el = elementClassOfType(c);
      const elChain = el ? chain(el, schemaId) : [];
      for (const e of elChain) {
        if (e === "IfcElement") { out.push("IfcElementType"); break; }
        if (e === "IfcSpatialElement") { out.push("IfcSpatialElementType"); break; }
        if (e !== el) out.push(e + "Type");
      }
      out.push("IfcTypeProduct", "IfcTypeObject", "IfcObjectDefinition", "IfcRoot");
      return [...new Set(out)];
    }
    let guard = 0;
    while (c && guard++ < 30) {
      out.push(legacy && c === "IfcBuiltElement" ? "IfcBuildingElement" : c);
      c = PARENT[c];
    }
    return out;
  }

  function isA(cls, ancestor) {
    let c = cls;
    let guard = 0;
    while (c && guard++ < 30) {
      if (c === ancestor) return true;
      if (ancestor === "IfcBuiltElement" && c === "IfcBuildingElement") return true;
      c = PARENT[c];
    }
    return false;
  }

  function isTypeClass(cls) {
    return /Type$|Style$/.test(cls) && !/^IfcRel/.test(cls) && !PARENT[cls];
  }
  function elementClassOfType(cls) {
    const base = cls.replace(/Type$|Style$/, "");
    return PARENT[base] ? base : null;
  }

  // ── Attribute names by class family (index → name) ──
  const ROOT = ["GlobalId", "OwnerHistory", "Name", "Description"];
  const OBJ = [...ROOT, "ObjectType"];
  const PROD = [...OBJ, "ObjectPlacement", "Representation"];
  const ELEM = [...PROD, "Tag"];
  const SPAT = [...PROD, "LongName", "CompositionType"];
  const TYPEOBJ = [...ROOT, "ApplicableOccurrence", "HasPropertySets"];
  const TYPEPROD = [...TYPEOBJ, "RepresentationMaps", "Tag", "ElementType", "PredefinedType"];
  const ATTRS = {
    IfcProject: [...OBJ, "LongName", "Phase", "RepresentationContexts", "UnitsInContext"],
    IfcSite: [...SPAT, "RefLatitude", "RefLongitude", "RefElevation", "LandTitleNumber", "SiteAddress"],
    IfcBuilding: [...SPAT, "ElevationOfRefHeight", "ElevationOfTerrain", "BuildingAddress"],
    IfcBuildingStorey: [...SPAT, "Elevation"],
    IfcSpace: [...SPAT, "PredefinedType", "ElevationWithFlooring"],
    IfcDistributionPort: [...PROD, "FlowDirection", "PredefinedType", "SystemType"],
    IfcDistributionSystem: [...OBJ, "LongName", "PredefinedType"],
    IfcDistributionCircuit: [...OBJ, "LongName", "PredefinedType"],
    IfcBuildingSystem: [...OBJ, "PredefinedType", "LongName"],
    IfcZone: [...OBJ, "LongName"],
    IfcSpatialZone: [...PROD, "LongName", "PredefinedType"],
    IfcDoor: [...ELEM, "OverallHeight", "OverallWidth", "PredefinedType", "OperationType", "UserDefinedOperationType"],
    IfcWindow: [...ELEM, "OverallHeight", "OverallWidth", "PredefinedType", "PartitioningType", "UserDefinedPartitioningType"],
    IfcPropertySet: [...ROOT, "HasProperties"],
    IfcElementQuantity: [...ROOT, "MethodOfMeasurement", "Quantities"],
    IfcTask: [...OBJ, "Identification", "LongDescription", "Status", "WorkMethod", "IsMilestone", "Priority", "TaskTime", "PredefinedType"],
  };
  function attributeNames(cls) {
    if (ATTRS[cls]) return ATTRS[cls];
    if (isTypeClass(cls)) return TYPEPROD;
    if (isA(cls, "IfcSpatialElement")) return SPAT;
    if (isA(cls, "IfcElement")) return [...ELEM, "PredefinedType"];
    if (isA(cls, "IfcProduct")) return PROD;
    if (isA(cls, "IfcObject")) return OBJ;
    if (isA(cls, "IfcRelationship")) return ROOT;
    if (cls === "IfcTypeObject" || isA(cls, "IfcTypeObject")) return TYPEOBJ;
    return ROOT;
  }

  // ── Node kinds (semantic role of a node in the knowledge graph) ──
  const KINDS = {
    Project:        { label: "Project",            shape: "hexagon",  color: "#475569" },
    Spatial:        { label: "Spatial structure",  shape: "hexagon",  color: "#8b5cf6" },
    Element:        { label: "Building element",   shape: "circle",   color: "#3b82f6" },
    Type:           { label: "Element type",       shape: "square",   color: "#6366f1" },
    PropertySet:    { label: "Property set",       shape: "diamond",  color: "#f59e0b" },
    QuantitySet:    { label: "Quantity set",       shape: "diamond",  color: "#d97706" },
    Property:       { label: "Property",           shape: "dot",      color: "#fbbf24" },
    Value:          { label: "Property value",     shape: "dot",      color: "#fcd34d" },
    Material:       { label: "Material",           shape: "triangle", color: "#a8a29e" },
    Classification: { label: "Classification",     shape: "tag",      color: "#14b8a6" },
    ClassSystem:    { label: "Classification system", shape: "tag",   color: "#0f766e" },
    System:         { label: "System",             shape: "ring",     color: "#06b6d4" },
    Zone:           { label: "Zone",               shape: "ring",     color: "#c084fc" },
    Group:          { label: "Group",              shape: "ring",     color: "#22d3ee" },
    Port:           { label: "Port",               shape: "dot",      color: "#94a3b8" },
    Opening:        { label: "Opening",            shape: "squareo",  color: "#94a3b8" },
    Document:       { label: "Document",           shape: "square",   color: "#84cc16" },
    Process:        { label: "Process / task",     shape: "triangle", color: "#f43f5e" },
    Actor:          { label: "Actor",              shape: "triangle", color: "#fb7185" },
    Control:        { label: "Control",            shape: "square",   color: "#fda4af" },
    Resource:       { label: "Resource",           shape: "square",   color: "#fb923c" },
    Class:          { label: "IFC class",          shape: "roundrect", color: "#64748b" },
    Aggregate:      { label: "Collapsed group",    shape: "double",   color: "#64748b" },
    Exterior:       { label: "Exterior",           shape: "hexagon",  color: "#334155" },
    Other:          { label: "Other",              shape: "circle",   color: "#9ca3af" },
  };

  function kindOf(cls) {
    if (cls === "IfcProject" || cls === "IfcProjectLibrary") return "Project";
    if (isA(cls, "IfcSpatialElement")) return "Spatial";
    if (isA(cls, "IfcFeatureElementSubtraction")) return "Opening";
    if (isA(cls, "IfcPort")) return "Port";
    if (isA(cls, "IfcElement") || isA(cls, "IfcProxy")) return "Element";
    if (isA(cls, "IfcZone")) return "Zone";
    if (isA(cls, "IfcSystem")) return "System";
    if (isA(cls, "IfcGroup")) return "Group";
    if (isA(cls, "IfcProcess")) return "Process";
    if (isA(cls, "IfcActor")) return "Actor";
    if (isA(cls, "IfcControl")) return "Control";
    if (isA(cls, "IfcResource")) return "Resource";
    if (isA(cls, "IfcProduct")) return "Element";
    if (isTypeClass(cls) || isA(cls, "IfcTypeObject")) return "Type";
    return "Other";
  }

  // ── Discipline / domain classification ──
  const DOMAINS = {
    Spatial:          { color: "#8b5cf6" },
    Structural:       { color: "#dc2626" },
    Architectural:    { color: "#2563eb" },
    HVAC:             { color: "#0891b2" },
    Plumbing:         { color: "#16a34a" },
    Electrical:       { color: "#ca8a04" },
    "Fire Protection":{ color: "#ea580c" },
    "Controls & Sensors": { color: "#db2777" },
    Furnishing:       { color: "#a16207" },
    "Vertical Transport": { color: "#0f766e" },
    "Site & Civil":   { color: "#65a30d" },
    Infrastructure:   { color: "#475569" },
    "Openings & Features": { color: "#94a3b8" },
    "Components & Assemblies": { color: "#78716c" },
    "Annotation & Grid": { color: "#a3a3a3" },
    "Generic / Proxy": { color: "#64748b" },
    Semantic:         { color: "#6366f1" },
    Other:            { color: "#9ca3af" },
  };

  const HVAC = ["IfcDuctSegment", "IfcDuctFitting", "IfcAirTerminal", "IfcAirTerminalBox", "IfcDamper", "IfcFan", "IfcAirToAirHeatRecovery", "IfcCoil", "IfcChiller", "IfcBoiler", "IfcCoolingTower", "IfcUnitaryEquipment", "IfcHumidifier", "IfcEvaporativeCooler", "IfcEvaporator", "IfcCondenser", "IfcHeatExchanger", "IfcSpaceHeater", "IfcCooledBeam", "IfcFilter", "IfcDuctSilencer", "IfcCompressor", "IfcTubeBundle", "IfcBurner", "IfcEngine"];
  const PLUMB = ["IfcPipeSegment", "IfcPipeFitting", "IfcValve", "IfcPump", "IfcSanitaryTerminal", "IfcWasteTerminal", "IfcTank", "IfcInterceptor", "IfcStackTerminal", "IfcLiquidTerminal", "IfcFlowMeter", "IfcDistributionChamberElement"];
  const ELEC = ["IfcCableSegment", "IfcCableFitting", "IfcCableCarrierSegment", "IfcCableCarrierFitting", "IfcLamp", "IfcLightFixture", "IfcOutlet", "IfcSwitchingDevice", "IfcProtectiveDevice", "IfcElectricAppliance", "IfcElectricDistributionBoard", "IfcDistributionBoard", "IfcElectricFlowStorageDevice", "IfcElectricFlowTreatmentDevice", "IfcElectricGenerator", "IfcElectricMotor", "IfcElectricTimeControl", "IfcTransformer", "IfcJunctionBox", "IfcMotorConnection", "IfcSolarDevice", "IfcAudioVisualAppliance", "IfcCommunicationsAppliance", "IfcElectricalElement", "IfcSignal"];
  const STRUCT = ["IfcBeam", "IfcColumn", "IfcFooting", "IfcDeepFoundation", "IfcPile", "IfcCaissonFoundation", "IfcMember", "IfcPlate", "IfcReinforcingElement", "IfcBearing", "IfcStructuralItem", "IfcStructuralActivity"];
  const ARCH = ["IfcDoor", "IfcWindow", "IfcCurtainWall", "IfcRailing", "IfcCovering", "IfcRoof", "IfcStair", "IfcStairFlight", "IfcRamp", "IfcRampFlight", "IfcShadingDevice", "IfcChimney"];
  const CIVIL = ["IfcGeographicElement", "IfcCivilElement", "IfcEarthworksElement", "IfcEarthworksCut", "IfcPavement", "IfcKerb", "IfcCourse"];
  const INFRA = ["IfcRail", "IfcTrackElement", "IfcAlignment", "IfcLinearPositioningElement", "IfcReferent", "IfcMooringDevice", "IfcNavigationElement", "IfcBridge", "IfcRoad", "IfcRailway", "IfcMarineFacility"];

  // IfcDistributionSystemEnum / IfcBuildingSystemTypeEnum → domain
  const SYSTEM_DOMAIN = [
    [/AIR|VENT|EXHAUST|HVAC|HEATING|COOLING|CHILLED|CONDENSER|REFRIGERATION|FUEL|GAS|STEAM|DUCT|SUPPLY ?AIR|RETURN ?AIR|MECHANICAL/i, "HVAC"],
    [/FIRE|SPRINKLER/i, "Fire Protection"],
    [/WATER|DRAIN|SEWAGE|SANITARY|WASTE|STORM|RAIN|PLUMB|VENT_?PIPE|DOMESTIC|POTABLE|IRRIGATION/i, "Plumbing"],
    [/ELECTR|POWER|LIGHT|LIGHTING|LOW ?VOLTAGE|COMMUNICATION|DATA|SECURITY|SIGNAL|TELE|AUDIO|LIGHTNING|EARTH/i, "Electrical"],
    [/CONTROL|AUTOMATION|MONITOR|BMS/i, "Controls & Sensors"],
  ];

  function domainFromSystemLabel(label) {
    if (!label) return null;
    for (const [re, d] of SYSTEM_DOMAIN) if (re.test(label)) return d;
    return null;
  }

  function anyIsA(cls, list) {
    for (const a of list) if (isA(cls, a)) return true;
    return false;
  }

  /**
   * Classify an entity into an engineering discipline.
   * hints: { loadBearing: bool|undefined, systemLabel: string|undefined }
   * Returns { domain, reason }.
   */
  function domainOf(cls, kind, hints = {}) {
    if (kind === "Project" || kind === "Spatial" || kind === "Zone") return { domain: "Spatial", reason: "IfcSpatialElement / IfcZone subtype" };
    if (kind !== "Element" && kind !== "Opening" && kind !== "Port") return { domain: "Semantic", reason: "Non-physical semantic entity" };
    if (kind === "Opening" || isA(cls, "IfcFeatureElement")) return { domain: "Openings & Features", reason: "IfcFeatureElement subtype" };
    if (hints.systemLabel && isA(cls, "IfcDistributionElement")) {
      const d = domainFromSystemLabel(hints.systemLabel);
      if (d) return { domain: d, reason: `Assigned to system “${hints.systemLabel}”` };
    }
    if (anyIsA(cls, ["IfcWall", "IfcSlab"])) {
      const isSlab = isA(cls, "IfcSlab");
      if (hints.loadBearing === true) return { domain: "Structural", reason: "LoadBearing = TRUE (Pset_*Common)" };
      if (hints.loadBearing === false) return { domain: "Architectural", reason: "LoadBearing = FALSE (Pset_*Common)" };
      return isSlab
        ? { domain: "Structural", reason: "IfcSlab default (no LoadBearing property)" }
        : { domain: "Architectural", reason: "IfcWall default (no LoadBearing property)" };
    }
    if (cls === "IfcFireSuppressionTerminal" || cls === "IfcAlarm") return { domain: "Fire Protection", reason: "Fire safety class" };
    if (anyIsA(cls, STRUCT)) return { domain: "Structural", reason: "Structural element class" };
    if (anyIsA(cls, ARCH)) return { domain: "Architectural", reason: "Architectural element class" };
    if (anyIsA(cls, HVAC)) return { domain: "HVAC", reason: "HVAC distribution class" };
    if (anyIsA(cls, PLUMB)) return { domain: "Plumbing", reason: "Plumbing / piping class" };
    if (anyIsA(cls, ELEC)) return { domain: "Electrical", reason: "Electrical class" };
    if (isA(cls, "IfcDistributionControlElement")) return { domain: "Controls & Sensors", reason: "IfcDistributionControlElement subtype" };
    if (isA(cls, "IfcDistributionElement")) return { domain: "HVAC", reason: "Generic distribution element (assumed HVAC)" };
    if (isA(cls, "IfcFurnishingElement")) return { domain: "Furnishing", reason: "IfcFurnishingElement subtype" };
    if (isA(cls, "IfcTransportElement")) return { domain: "Vertical Transport", reason: "IfcTransportElement" };
    if (anyIsA(cls, INFRA)) return { domain: "Infrastructure", reason: "IFC4x3 infrastructure class" };
    if (anyIsA(cls, CIVIL)) return { domain: "Site & Civil", reason: "Civil / geographic class" };
    if (anyIsA(cls, ["IfcElementComponent", "IfcElementAssembly"])) return { domain: "Components & Assemblies", reason: "Component / assembly" };
    if (anyIsA(cls, ["IfcAnnotation", "IfcGrid", "IfcVirtualElement", "IfcPositioningElement"])) return { domain: "Annotation & Grid", reason: "Annotation / positioning" };
    if (cls === "IfcBuildingElementProxy" || cls === "IfcProxy") return { domain: "Generic / Proxy", reason: "Unclassified proxy element" };
    if (kind === "Port") return { domain: "Other", reason: "Distribution port" };
    return { domain: "Other", reason: "No rule matched" };
  }

  // ── Ontology alignment ──
  const NS = {
    ifc: "https://standards.buildingsmart.org/IFC/DEV/IFC4/ADD2_TC1/OWL#",
    bot: "https://w3id.org/bot#",
    brick: "https://brickschema.org/schema/Brick#",
    rdfs: "http://www.w3.org/2000/01/rdf-schema#",
    rdf: "http://www.w3.org/1999/02/22-rdf-syntax-ns#",
    xsd: "http://www.w3.org/2001/XMLSchema#",
    bux: "https://isaddiq.github.io/BIMUniXchange/ontology#",
  };

  const BRICK = {
    IfcSite: "Site", IfcBuilding: "Building", IfcBuildingStorey: "Floor", IfcSpace: "Space", IfcZone: "Zone",
    IfcSpatialZone: "Zone", IfcPump: "Pump", IfcFan: "Fan", IfcBoiler: "Boiler", IfcChiller: "Chiller",
    IfcValve: "Valve", IfcDamper: "Damper", IfcFilter: "Filter", IfcCoil: "Coil", IfcHeatExchanger: "Heat_Exchanger",
    IfcSensor: "Sensor", IfcFlowMeter: "Meter", IfcCoolingTower: "Cooling_Tower", IfcLightFixture: "Luminaire",
    IfcLamp: "Luminaire", IfcCompressor: "Compressor", IfcTransformer: "Transformer",
  };

  function ontologyFor(cls, kind) {
    const out = { ifcowl: "ifc:" + cls };
    if (cls === "IfcSite") out.bot = "bot:Site";
    else if (cls === "IfcBuilding") out.bot = "bot:Building";
    else if (cls === "IfcBuildingStorey") out.bot = "bot:Storey";
    else if (cls === "IfcSpace") out.bot = "bot:Space";
    else if (kind === "Zone" || isA(cls, "IfcSpatialZone") || (kind === "Spatial" && isA(cls, "IfcSpatialElement"))) out.bot = "bot:Zone";
    else if (kind === "Element" || kind === "Opening") out.bot = "bot:Element";
    if (BRICK[cls]) out.brick = "brick:" + BRICK[cls];
    else if (cls === "IfcUnitaryEquipment") out.brick = "brick:HVAC_Equipment";
    else if (kind === "System" && isA(cls, "IfcDistributionSystem")) out.brick = "brick:System";
    return out;
  }

  // ── Relationship semantics ──
  // src/dst: argument index; "list" flag when argument is an aggregate.
  // Edges are drawn src → dst with predicate `pred`.
  const REL = {
    IFCRELAGGREGATES:                  { src: 4, dst: 5, pred: "aggregates" },
    IFCRELNESTS:                       { src: 4, dst: 5, pred: "nests" },
    IFCRELCONTAINEDINSPATIALSTRUCTURE: { src: 5, dst: 4, pred: "contains" },
    IFCRELREFERENCEDINSPATIALSTRUCTURE:{ src: 5, dst: 4, pred: "references" },
    IFCRELDEFINESBYTYPE:               { src: 4, dst: 5, pred: "isTypedBy" },
    IFCRELDEFINESBYPROPERTIES:         { src: 4, dst: 5, pred: "hasPropertySet" },
    IFCRELASSOCIATESMATERIAL:          { src: 4, dst: 5, pred: "hasMaterial", special: "material" },
    IFCRELASSOCIATESCLASSIFICATION:    { src: 4, dst: 5, pred: "hasClassification", special: "classification" },
    IFCRELASSOCIATESDOCUMENT:          { src: 4, dst: 5, pred: "hasDocument", special: "document" },
    IFCRELASSIGNSTOGROUP:              { src: 4, dst: 6, pred: "memberOf" },
    IFCRELASSIGNSTOGROUPBYFACTOR:      { src: 4, dst: 6, pred: "memberOf" },
    IFCRELSERVICESBUILDINGS:           { src: 4, dst: 5, pred: "serves" },
    IFCRELCONNECTSELEMENTS:            { src: 5, dst: 6, pred: "connectedTo" },
    IFCRELCONNECTSPATHELEMENTS:        { src: 5, dst: 6, pred: "connectedTo", special: "path" },
    IFCRELCONNECTSWITHREALIZINGELEMENTS:{ src: 5, dst: 6, pred: "connectedTo", special: "realizing" },
    IFCRELCONNECTSPORTTOELEMENT:       { src: 5, dst: 4, pred: "hasPort" },
    IFCRELCONNECTSPORTS:               { src: 4, dst: 5, pred: "portConnectedTo" },
    IFCRELFILLSELEMENT:                { src: 5, dst: 4, pred: "fills" },
    IFCRELVOIDSELEMENT:                { src: 4, dst: 5, pred: "hasOpening" },
    IFCRELPROJECTSELEMENT:             { src: 4, dst: 5, pred: "hasProjection" },
    IFCRELSPACEBOUNDARY:               { src: 4, dst: 5, pred: "boundedBy", special: "boundary" },
    IFCRELSPACEBOUNDARY1STLEVEL:       { src: 4, dst: 5, pred: "boundedBy", special: "boundary" },
    IFCRELSPACEBOUNDARY2NDLEVEL:       { src: 4, dst: 5, pred: "boundedBy", special: "boundary" },
    IFCRELCOVERSBLDGELEMENTS:          { src: 4, dst: 5, pred: "coveredBy" },
    IFCRELCOVERSSPACES:                { src: 4, dst: 5, pred: "coveredBy" },
    IFCRELFLOWCONTROLELEMENTS:         { src: 5, dst: 4, pred: "controlledBy" },
    IFCRELINTERFERESELEMENTS:          { src: 4, dst: 5, pred: "interferesWith" },
    IFCRELASSIGNSTOPRODUCT:            { src: 6, dst: 4, pred: "hasAssignment" },
    IFCRELASSIGNSTOPROCESS:            { src: 6, dst: 4, pred: "operatesOn" },
    IFCRELASSIGNSTOACTOR:              { src: 6, dst: 4, pred: "responsibleFor" },
    IFCRELASSIGNSTOCONTROL:            { src: 6, dst: 4, pred: "controls" },
    IFCRELASSIGNSTORESOURCE:           { src: 6, dst: 4, pred: "resourceFor" },
    IFCRELSEQUENCE:                    { src: 4, dst: 5, pred: "precedes" },
    IFCRELDECLARES:                    { src: 4, dst: 5, pred: "declares" },
    IFCRELCONNECTSSTRUCTURALMEMBER:    { src: 4, dst: 5, pred: "structurallyConnectedTo" },
    IFCRELCONNECTSSTRUCTURALACTIVITY:  { src: 4, dst: 5, pred: "hasStructuralActivity" },
    IFCRELADHERESTOELEMENT:            { src: 4, dst: 5, pred: "adheredBy" },
    IFCRELPOSITIONS:                   { src: 4, dst: 5, pred: "positions" },
  };

  // Predicate vocabulary: category, display label, ontology predicate.
  const PRED = {
    aggregates:      { cat: "Decomposition", label: "aggregates", rdf: "bux:aggregates" },
    nests:           { cat: "Decomposition", label: "nests", rdf: "bux:nests" },
    groups:          { cat: "Decomposition", label: "groups", rdf: "bux:groups" },
    contains:        { cat: "Containment", label: "contains", rdf: "bot:containsElement" },
    references:      { cat: "Containment", label: "references", rdf: "bux:referencesElement" },
    isTypedBy:       { cat: "Typing", label: "is typed by", rdf: "bux:isTypedBy" },
    hasPropertySet:  { cat: "Properties", label: "has property set", rdf: "bux:hasPropertySet" },
    hasValue:        { cat: "Properties", label: "has value", rdf: "bux:hasValue" },
    hasProperty:     { cat: "Properties", label: "has property", rdf: "bux:hasProperty" },
    typeFor:         { cat: "Typing", label: "type for class", rdf: "bux:typeForClass" },
    hasMaterial:     { cat: "Material", label: "has material", rdf: "bux:hasMaterial" },
    hasClassification:{ cat: "Classification", label: "classified as", rdf: "bux:hasClassification" },
    inSystem:        { cat: "Classification", label: "belongs to scheme", rdf: "bux:inClassificationSystem" },
    hasDocument:     { cat: "Documentation", label: "has document", rdf: "bux:hasDocument" },
    memberOf:        { cat: "Systems & Groups", label: "member of", rdf: "brick:isPartOf" },
    serves:          { cat: "Systems & Groups", label: "serves", rdf: "bux:servesBuilding" },
    connectedTo:     { cat: "Physical connectivity", label: "connected to", rdf: "bux:connectedTo", sym: true },
    hasPort:         { cat: "MEP connectivity", label: "has port", rdf: "bux:hasPort" },
    portConnectedTo: { cat: "MEP connectivity", label: "port connected to", rdf: "bux:portConnectedTo" },
    feeds:           { cat: "MEP connectivity", label: "feeds", rdf: "brick:feeds", inferred: true },
    fills:           { cat: "Hosting", label: "fills", rdf: "bux:fills" },
    hasOpening:      { cat: "Hosting", label: "has opening", rdf: "bux:hasOpening" },
    hasProjection:   { cat: "Hosting", label: "has projection", rdf: "bux:hasProjection" },
    hosts:           { cat: "Hosting", label: "hosts", rdf: "bot:hostsElement", inferred: true },
    boundedBy:       { cat: "Space boundaries", label: "bounded by", rdf: "bot:adjacentElement" },
    adjacentTo:      { cat: "Space topology", label: "adjacent to", rdf: "bot:adjacentZone", inferred: true, sym: true },
    accessTo:        { cat: "Space topology", label: "door access to", rdf: "bux:hasAccessTo", inferred: true, sym: true },
    coveredBy:       { cat: "Hosting", label: "covered by", rdf: "bux:coveredBy" },
    controlledBy:    { cat: "MEP connectivity", label: "controlled by", rdf: "bux:controlledBy" },
    interferesWith:  { cat: "Physical connectivity", label: "interferes with", rdf: "bux:interferesWith" },
    hasAssignment:   { cat: "Process & control", label: "has assignment", rdf: "bux:hasAssignment" },
    operatesOn:      { cat: "Process & control", label: "operates on", rdf: "bux:operatesOn" },
    responsibleFor:  { cat: "Process & control", label: "responsible for", rdf: "bux:responsibleFor" },
    controls:        { cat: "Process & control", label: "controls", rdf: "bux:controls" },
    resourceFor:     { cat: "Process & control", label: "resource for", rdf: "bux:resourceFor" },
    precedes:        { cat: "Process & control", label: "precedes", rdf: "bux:precedes" },
    declares:        { cat: "Decomposition", label: "declares", rdf: "bux:declares" },
    structurallyConnectedTo: { cat: "Physical connectivity", label: "structurally connected to", rdf: "bux:structurallyConnectedTo" },
    hasStructuralActivity: { cat: "Physical connectivity", label: "has structural activity", rdf: "bux:hasStructuralActivity" },
    adheredBy:       { cat: "Physical connectivity", label: "adhered by", rdf: "bux:adheredBy" },
    positions:       { cat: "Containment", label: "positions", rdf: "bux:positions" },
    locatedIn:       { cat: "Containment", label: "located in", rdf: "bot:containsElement", inferred: true },
    subClassOf:      { cat: "Schema", label: "subclass of", rdf: "rdfs:subClassOf" },
    instanceOf:      { cat: "Schema", label: "instance of", rdf: "rdf:type" },
    related:         { cat: "Other", label: "related to", rdf: "bux:relatedTo" },
    relates:         { cat: "Schema", label: "relates", rdf: "bux:relatesTo" },
  };

  const CATEGORY_COLORS = {
    Decomposition: "#8b5cf6",
    Containment: "#a78bfa",
    Typing: "#6366f1",
    Properties: "#f59e0b",
    Material: "#a8a29e",
    Classification: "#14b8a6",
    Documentation: "#84cc16",
    "Systems & Groups": "#06b6d4",
    "Physical connectivity": "#ef4444",
    "MEP connectivity": "#0ea5e9",
    Hosting: "#f97316",
    "Space boundaries": "#c084fc",
    "Space topology": "#10b981",
    "Process & control": "#f43f5e",
    Schema: "#64748b",
    Other: "#9ca3af",
  };

  /** BOT predicate refinement for decomposition between spatial levels. */
  function botDecomposition(srcCls, dstCls) {
    if (dstCls === "IfcBuilding") return "bot:hasBuilding";
    if (dstCls === "IfcBuildingStorey") return "bot:hasStorey";
    if (dstCls === "IfcSpace") return "bot:hasSpace";
    if (isA(dstCls, "IfcElement") && isA(srcCls, "IfcElement")) return "bot:hasSubElement";
    return null;
  }

  // IfcSIUnit naming
  const SI_SYMBOL = {
    METRE: "m", SQUARE_METRE: "m²", CUBIC_METRE: "m³", GRAM: "g", SECOND: "s", RADIAN: "rad",
    DEGREE_CELSIUS: "°C", KELVIN: "K", WATT: "W", PASCAL: "Pa", NEWTON: "N", JOULE: "J", HERTZ: "Hz",
    VOLT: "V", AMPERE: "A", LUMEN: "lm", LUX: "lx", CANDELA: "cd", STERADIAN: "sr", MOLE: "mol",
    OHM: "Ω", COULOMB: "C", FARAD: "F", SIEMENS: "S", TESLA: "T", HENRY: "H", WEBER: "Wb", LITRE: "L",
  };
  const SI_PREFIX = { EXA: "E", PETA: "P", TERA: "T", GIGA: "G", MEGA: "M", KILO: "k", HECTO: "h", DECA: "da", DECI: "d", CENTI: "c", MILLI: "m", MICRO: "µ", NANO: "n", PICO: "p", FEMTO: "f", ATTO: "a" };
  const MEASURE_UNIT = {
    IFCLENGTHMEASURE: "LENGTHUNIT", IFCPOSITIVELENGTHMEASURE: "LENGTHUNIT", IFCNONNEGATIVELENGTHMEASURE: "LENGTHUNIT",
    IFCAREAMEASURE: "AREAUNIT", IFCVOLUMEMEASURE: "VOLUMEUNIT", IFCPLANEANGLEMEASURE: "PLANEANGLEUNIT",
    IFCPOSITIVEPLANEANGLEMEASURE: "PLANEANGLEUNIT", IFCMASSMEASURE: "MASSUNIT", IFCTIMEMEASURE: "TIMEUNIT",
    IFCTHERMODYNAMICTEMPERATUREMEASURE: "THERMODYNAMICTEMPERATUREUNIT", IFCPOWERMEASURE: "POWERUNIT",
    IFCELECTRICCURRENTMEASURE: "ELECTRICCURRENTUNIT", IFCELECTRICVOLTAGEMEASURE: "ELECTRICVOLTAGEUNIT",
    IFCPRESSUREMEASURE: "PRESSUREUNIT", IFCENERGYMEASURE: "ENERGYUNIT", IFCFORCEMEASURE: "FORCEUNIT",
    IFCFREQUENCYMEASURE: "FREQUENCYUNIT", IFCLUMINOUSFLUXMEASURE: "LUMINOUSFLUXUNIT", IFCILLUMINANCEMEASURE: "ILLUMINANCEUNIT",
    IFCQUANTITYLENGTH: "LENGTHUNIT", IFCQUANTITYAREA: "AREAUNIT", IFCQUANTITYVOLUME: "VOLUMEUNIT",
    IFCQUANTITYWEIGHT: "MASSUNIT", IFCQUANTITYTIME: "TIMEUNIT",
  };

  KG.Schema = {
    TREE, PARENT, KINDS, DOMAINS, REL, PRED, NS, CATEGORY_COLORS, SI_SYMBOL, SI_PREFIX, MEASURE_UNIT,
    canonical, chain, isA, isTypeClass, elementClassOfType, attributeNames, kindOf, domainOf,
    domainFromSystemLabel, ontologyFor, botDecomposition,
  };
})((window.KG = window.KG || {}));
