using UnityEngine;
using UnityEditor;
using UnityEditor.SceneManagement;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System;
using System.Text.RegularExpressions;

/// <summary>
/// Represents a node in the IFC schema tree structure.
/// </summary>
[Serializable]
public class SchemaNode
{
    public string Name;
    public string Kind; // TypeEnum, Entity, Pset, Qto, Package
    public string Package;
    public string Group;
    public string SubGroup;
    public List<SchemaNode> Children = new List<SchemaNode>();
    public GameObject GameObject; // Reference to the created GameObject in the scene

    public SchemaNode(string name, string kind = "", string package = "", string group = "", string subGroup = "")
    {
        Name = name;
        Kind = kind;
        Package = package;
        Group = group;
        SubGroup = subGroup;
    }

    public SchemaNode FindChild(string name)
    {
        return Children.Find(c => c.Name == name);
    }

    public SchemaNode GetOrCreateChild(string name, string kind = "", string package = "", string group = "", string subGroup = "")
    {
        var existing = FindChild(name);
        if (existing != null)
            return existing;

        var newChild = new SchemaNode(name, kind, package, group, subGroup);
        Children.Add(newChild);
        return newChild;
    }
}

/// <summary>
/// Component attached to BIM elements to store Pset and Qto associations.
/// </summary>
public class IfcPropertyData : MonoBehaviour
{
    [Serializable]
    public class PropertySetReference
    {
        public string PropertySetName;
        public List<PropertyValue> Properties = new List<PropertyValue>();
    }

    [Serializable]
    public class QuantitySetReference
    {
        public string QuantitySetName;
        public List<PropertyValue> Quantities = new List<PropertyValue>();
    }

    [Serializable]
    public class PropertyValue
    {
        public string Name;
        public string Value;

        public PropertyValue(string name, string value)
        {
            Name = name;
            Value = value;
        }
    }

    public string GlobalId;
    public string IfcClass;
    public List<PropertySetReference> PropertySets = new List<PropertySetReference>();
    public List<QuantitySetReference> QuantitySets = new List<QuantitySetReference>();

    /// <summary>
    /// Adds or updates a property in the specified property set.
    /// </summary>
    public void AddProperty(string psetName, string propertyName, string value)
    {
        var pset = PropertySets.Find(p => p.PropertySetName == psetName);
        if (pset == null)
        {
            pset = new PropertySetReference { PropertySetName = psetName };
            PropertySets.Add(pset);
        }

        var existingProp = pset.Properties.Find(p => p.Name == propertyName);
        if (existingProp != null)
            existingProp.Value = value;
        else
            pset.Properties.Add(new PropertyValue(propertyName, value));
    }

    /// <summary>
    /// Adds or updates a quantity in the specified quantity set.
    /// </summary>
    public void AddQuantity(string qtoName, string quantityName, string value)
    {
        var qto = QuantitySets.Find(q => q.QuantitySetName == qtoName);
        if (qto == null)
        {
            qto = new QuantitySetReference { QuantitySetName = qtoName };
            QuantitySets.Add(qto);
        }

        var existingQty = qto.Quantities.Find(q => q.Name == quantityName);
        if (existingQty != null)
            existingQty.Value = value;
        else
            qto.Quantities.Add(new PropertyValue(quantityName, value));
    }
}

/// <summary>
/// Component that links a schema hierarchy element to its original GameObject.
/// This allows the schema hierarchy to reference elements without moving them from the spatial hierarchy.
/// </summary>
public class SchemaElementReference : MonoBehaviour
{
    [Tooltip("Reference to the original element GameObject in the spatial hierarchy")]
    public GameObject OriginalElement;

    [Tooltip("The IFC class of this element (e.g., IfcBeam, IfcWall)")]
    public string IfcClass;

    /// <summary>
    /// Selects the original element in the hierarchy when this reference is clicked.
    /// </summary>
    public void SelectOriginal()
    {
        if (OriginalElement != null)
        {
#if UNITY_EDITOR
            UnityEditor.Selection.activeGameObject = OriginalElement;
            UnityEditor.EditorGUIUtility.PingObject(OriginalElement);
#endif
        }
    }
}

/// <summary>
/// Custom editor for SchemaElementReference to add a "Select Original" button.
/// </summary>
[CustomEditor(typeof(SchemaElementReference))]
public class SchemaElementReferenceEditor : Editor
{
    public override void OnInspectorGUI()
    {
        DrawDefaultInspector();

        SchemaElementReference reference = (SchemaElementReference)target;

        EditorGUILayout.Space();

        GUI.enabled = reference.OriginalElement != null;
        if (GUILayout.Button("Select Original Element", GUILayout.Height(25)))
        {
            reference.SelectOriginal();
        }
        GUI.enabled = true;

        if (reference.OriginalElement == null)
        {
            EditorGUILayout.HelpBox("Original element reference is missing.", MessageType.Warning);
        }
    }
}

/// <summary>
/// Custom editor for IfcPropertyData to display property and quantity sets nicely.
/// </summary>
[CustomEditor(typeof(IfcPropertyData))]
public class IfcPropertyDataEditor : Editor
{
    private bool showPropertySets = true;
    private bool showQuantitySets = true;

    public override void OnInspectorGUI()
    {
        IfcPropertyData propData = (IfcPropertyData)target;

        EditorGUILayout.LabelField("IFC Element Information", EditorStyles.boldLabel);
        EditorGUILayout.LabelField("Global ID", propData.GlobalId);
        EditorGUILayout.LabelField("IFC Class", propData.IfcClass);

        EditorGUILayout.Space();

        // Property Sets
        showPropertySets = EditorGUILayout.Foldout(showPropertySets, $"Property Sets ({propData.PropertySets.Count})");
        if (showPropertySets)
        {
            EditorGUI.indentLevel++;
            foreach (var pset in propData.PropertySets)
            {
                EditorGUILayout.LabelField(pset.PropertySetName, EditorStyles.boldLabel);
                EditorGUI.indentLevel++;
                foreach (var prop in pset.Properties)
                {
                    EditorGUILayout.LabelField(prop.Name, prop.Value);
                }
                EditorGUI.indentLevel--;
            }
            EditorGUI.indentLevel--;
        }

        EditorGUILayout.Space();

        // Quantity Sets
        showQuantitySets = EditorGUILayout.Foldout(showQuantitySets, $"Quantity Sets ({propData.QuantitySets.Count})");
        if (showQuantitySets)
        {
            EditorGUI.indentLevel++;
            foreach (var qto in propData.QuantitySets)
            {
                EditorGUILayout.LabelField(qto.QuantitySetName, EditorStyles.boldLabel);
                EditorGUI.indentLevel++;
                foreach (var qty in qto.Quantities)
                {
                    EditorGUILayout.LabelField(qty.Name, qty.Value);
                }
                EditorGUI.indentLevel--;
            }
            EditorGUI.indentLevel--;
        }
    }
}

public class UnifiedHierarchyOrganizer : EditorWindow
{
    public enum BIMType
    {
        Archicad,
        Revit
    }

    public enum HierarchyMode
    {
        // Universal modes for both BIM types
        ByCategory,              // Category → Elements
        ByFamily,                // Family → Type → Elements  
        ByMaterial,              // Material → Category → Elements
        ByLevel,                 // Level → Category → Family → Elements
        BySystem,                // System → Category → Family → Elements

        // Advanced organizational modes
        Modular,                 // Modular_Group → Category → Family → Type → Elements
        SystemDetailed,          // System_Name → System_Type → Category → Family → Elements
        MaterialDetailed,        // Material → Category → Family → Type → Elements
        LevelDetailed,           // Level → Category → Family → Type → Elements
        ModularDetailed,         // Modular_Group → Level → System → Category → Family → Type → Elements

        // Revit specific modes
        FlatByCategory,          // Category → Elements (flat structure)

        // ArchiCAD specific modes
        ConstructionDiscipline,  // Layer-based → Element Type → Elements
        SpatialDiscipline,       // Zone → Element Type → Elements
        MultiLevelClassification, // Element Type → Profile/Library → Elements

        // IFC Schema mode (uses element CSV to build IFC hierarchy)
        IfcSchema,               // IfcProject → IfcSite → IfcBuilding → IfcBuildingStorey → EntityType → Elements

        // Custom mode
        Custom                   // User-defined hierarchy levels
    }

    [SerializeField] private BIMType bimType = BIMType.Archicad;
    [SerializeField] private HierarchyMode archicadHierarchyMode = HierarchyMode.ByCategory;
    [SerializeField] private HierarchyMode revitHierarchyMode = HierarchyMode.ByCategory;
    [SerializeField] private GameObject modelRoot;

    // Property to get/set the current hierarchy mode based on BIM type
    private HierarchyMode CurrentHierarchyMode
    {
        get => bimType == BIMType.Archicad ? archicadHierarchyMode : revitHierarchyMode;
        set
        {
            if (bimType == BIMType.Archicad)
                archicadHierarchyMode = value;
            else
                revitHierarchyMode = value;
        }
    }

    // Property to get/set the current custom hierarchy based on BIM type
    private string CurrentCustomHierarchy
    {
        get => bimType == BIMType.Archicad ? archicadCustomHierarchy : revitCustomHierarchy;
        set
        {
            if (bimType == BIMType.Archicad)
                archicadCustomHierarchy = value;
            else
                revitCustomHierarchy = value;
        }
    }
    [SerializeField] private UnityEngine.Object csvFile;
    [SerializeField] private bool createUnmatchedGroup = true;
    [SerializeField] private bool preserveOriginalNames = false;
    [SerializeField] private bool useProgressBar = true;
    [SerializeField] private bool debugMode = false;
    [SerializeField] private string archicadCustomHierarchy = "Category,Family,Type_Type Name";
    [SerializeField] private string revitCustomHierarchy = "Category,Family,Type_Type Name";
    [SerializeField] private bool caseSensitiveMatching = false;
    [SerializeField] private string elementIdPrefix = "";
    [SerializeField] private string elementIdSuffix = "";

    // IFC Schema Hierarchy settings
    [SerializeField] private bool enableIfcSchemaHierarchy = true;
    [SerializeField] private UnityEngine.Object schemaCSVFile;
    [SerializeField] private string ifcSchemaRootName = "IFC_Schema_Root";
    [SerializeField] private bool createSchemaPropertyLinks = true;
    [SerializeField] private bool useElementCsvForSchema = true; // Use the element CSV to build schema hierarchy

    private Dictionary<string, Dictionary<string, string>> elementMetadata;
    private HashSet<string> availableColumns;
    private Vector2 scrollPosition;
    private Vector2 logScroll;
    private bool showAdvancedOptions = false;
    private bool showColumnInfo = false;
    private bool showIfcSchemaOptions = false;
    private string logText = "";
    private bool busy = false;

    // IFC Schema tree structure
    private SchemaNode schemaRoot;
    private Dictionary<string, SchemaNode> entityNodeLookup; // Quick lookup: IfcClass name -> SchemaNode
    private Dictionary<string, SchemaNode> psetNodeLookup;   // Quick lookup: Pset name -> SchemaNode
    private Dictionary<string, SchemaNode> qtoNodeLookup;    // Quick lookup: Qto name -> SchemaNode
    private GameObject ifcSchemaRootObject;

    // Performance tracking
    private System.Diagnostics.Stopwatch performanceTimer;
    private long memoryAtStart;
    private long peakMemoryUsed;
    private int csvRowsLoaded;
    private int hierarchyNodesCreated;
    private int schemaReferencesCreated;

    [MenuItem("Window/BIMUniXchange/Hierarchy Organizer", false, 31)]
    public static void ShowWindow()
    {
        var window = GetWindow<UnifiedHierarchyOrganizer>("BIM Hierarchy Organizer");
        window.minSize = new Vector2(500, 700);
    }

    private void OnGUI()
    {
        scrollPosition = EditorGUILayout.BeginScrollView(scrollPosition);

        DrawHeader();
        EditorGUILayout.Space();

        // BIM Type Selection
        DrawBIMTypeSelection();
        EditorGUILayout.Space();

        // Main settings
        DrawMainSettings();
        EditorGUILayout.Space();

        // Hierarchy mode selection
        DrawHierarchyModeSelection();
        EditorGUILayout.Space();

        // Advanced options
        DrawAdvancedOptions();
        EditorGUILayout.Space();

        // IFC Schema Hierarchy options
        DrawIfcSchemaOptions();
        EditorGUILayout.Space();

        // Column information
        if (csvFile != null)
        {
            DrawColumnInformation();
            EditorGUILayout.Space();
        }

        // Action buttons
        DrawActionButtons();
        EditorGUILayout.Space();

        // Log section
        DrawLogSection();

        EditorGUILayout.EndScrollView();
    }

    private void DrawHeader()
    {
        var titleStyle = new GUIStyle(EditorStyles.boldLabel)
        { fontSize = 16, alignment = TextAnchor.MiddleCenter };
        GUILayout.Label("BIM Hierarchy Organizer", titleStyle);
        GUILayout.Label("Advanced Multi-Level BIM Element Organization System", EditorStyles.centeredGreyMiniLabel);
        GUILayout.Label("", GUI.skin.horizontalSlider);
    }

    private void DrawBIMTypeSelection()
    {
        EditorGUILayout.LabelField("BIM Platform", EditorStyles.boldLabel);

        EditorGUILayout.BeginHorizontal();
        if (GUILayout.Toggle(bimType == BIMType.Archicad, "ArchiCAD (FBX)", EditorStyles.miniButtonLeft))
            bimType = BIMType.Archicad;
        if (GUILayout.Toggle(bimType == BIMType.Revit, "Revit (OBJ)", EditorStyles.miniButtonRight))
            bimType = BIMType.Revit;
        EditorGUILayout.EndHorizontal();

        string helpText = bimType == BIMType.Archicad ?
            "ArchiCAD mode: Direct GameObject name matching with Element ID column" :
            "Revit mode: Element ID extraction from GameObject name (after last underscore)";
        EditorGUILayout.HelpBox(helpText, MessageType.Info);
    }

    private void DrawMainSettings()
    {
        EditorGUILayout.LabelField("Main Settings", EditorStyles.boldLabel);
        using (new EditorGUILayout.VerticalScope(EditorStyles.helpBox))
        {
            string modelLabel = bimType == BIMType.Archicad ? "ArchiCAD FBX Model" : "Revit OBJ Model";
            EditorGUILayout.LabelField(modelLabel, EditorStyles.miniLabel);
            modelRoot = (GameObject)EditorGUILayout.ObjectField(modelRoot, typeof(GameObject), true);

            GUILayout.Space(4);
            EditorGUILayout.LabelField("Element Metadata CSV", EditorStyles.miniLabel);
            csvFile = EditorGUILayout.ObjectField(csvFile, typeof(UnityEngine.Object), false);

            if (csvFile != null && availableColumns == null)
            {
                LoadColumnInformation();
            }
        }
    }

    private void DrawHierarchyModeSelection()
    {
        EditorGUILayout.LabelField("Organizational Structure", EditorStyles.boldLabel);
        using (new EditorGUILayout.VerticalScope(EditorStyles.helpBox))
        {
            // Show which BIM platform's mode is being configured
            string platformLabel = $"Hierarchy Mode ({bimType})";

            if (bimType == BIMType.Revit)
            {
                CurrentHierarchyMode = DrawRevitHierarchyModePopup(platformLabel, CurrentHierarchyMode);
            }
            else
            {
                CurrentHierarchyMode = DrawArchicadHierarchyModePopup(platformLabel, CurrentHierarchyMode);
            }

            EditorGUILayout.HelpBox(GetHierarchyDescription(CurrentHierarchyMode), MessageType.Info);

            if (CurrentHierarchyMode == HierarchyMode.Custom)
            {
                EditorGUILayout.LabelField($"Custom Hierarchy Levels ({bimType}):", EditorStyles.boldLabel);
                CurrentCustomHierarchy = EditorGUILayout.TextField("Hierarchy Columns", CurrentCustomHierarchy);
                EditorGUILayout.HelpBox("Enter column names separated by commas. Example: Category,System_Name,Family,Type_Type Name", MessageType.Info);
            }

            // Show a hint about separate configurations
            EditorGUILayout.HelpBox($"Note: Each BIM platform maintains its own hierarchy mode configuration. Switch tabs to configure the other platform's mode.", MessageType.None);
        }
    }

    private HierarchyMode DrawRevitHierarchyModePopup(string label, HierarchyMode currentMode)
    {
        // Revit-specific modes based on the image
        var revitModes = new[]
        {
            HierarchyMode.ByCategory,        // NonModular: Category > Family > Type > Elements
            HierarchyMode.Modular,           // Modular: Modular Group > Category > Family > Type > Elements
            HierarchyMode.BySystem,          // System Based: System Name > System Type > Category > Family > Elements
            HierarchyMode.ByLevel,           // Level Based: Level > Category > Family > Type > Elements
            HierarchyMode.ByMaterial,        // Material Based: Material > Category > Family > Type > Elements
            HierarchyMode.ModularDetailed,   // Modular Detailed: Modular Group > Level > System Name > Category > Family > Type > Elements
            HierarchyMode.FlatByCategory,    // Flat By Category: Category > Elements flat structure
            HierarchyMode.IfcSchema,         // IFC Schema: IfcProject > IfcSite > IfcBuilding > IfcBuildingStorey > EntityType > Elements
            HierarchyMode.Custom             // Custom: User-defined comma-separated columns
        };

        var revitModeNames = new[]
        {
            "NonModular",
            "Modular",
            "System Based",
            "Level Based",
            "Material Based",
            "Modular Detailed",
            "Flat By Category",
            "IFC Schema",
            "Custom"
        };

        int currentIndex = Array.IndexOf(revitModes, currentMode);
        if (currentIndex == -1) currentIndex = 0; // Default to first mode if current mode is not in Revit modes

        int selectedIndex = EditorGUILayout.Popup(label, currentIndex, revitModeNames);
        return revitModes[selectedIndex];
    }

    private HierarchyMode DrawArchicadHierarchyModePopup(string label, HierarchyMode currentMode)
    {
        // Per updated requirement: Only show these ArchiCAD modes
        var archicadModes = new[]
        {
            HierarchyMode.ConstructionDiscipline,
            HierarchyMode.SpatialDiscipline,
            HierarchyMode.MultiLevelClassification,
            HierarchyMode.IfcSchema,
            HierarchyMode.Custom
        };

        var archicadModeNames = new[]
        {
            "Construction Discipline",
            "Spatial Discipline",
            "Multi-Level Classification",
            "IFC Schema",
            "Custom"
        };

        // If previously saved mode isn't in the restricted list, default to first
        int currentIndex = Array.IndexOf(archicadModes, currentMode);
        if (currentIndex == -1) currentIndex = 0;

        int selectedIndex = EditorGUILayout.Popup(label, currentIndex, archicadModeNames);
        return archicadModes[selectedIndex];
    }

    private void DrawAdvancedOptions()
    {
        showAdvancedOptions = EditorGUILayout.Foldout(showAdvancedOptions, "Advanced Options");

        if (showAdvancedOptions)
        {
            using (new EditorGUILayout.VerticalScope(EditorStyles.helpBox))
            {
                EditorGUI.indentLevel++;

                createUnmatchedGroup = EditorGUILayout.Toggle("Create Unmatched Elements Group", createUnmatchedGroup);
                preserveOriginalNames = EditorGUILayout.Toggle("Preserve Original Element Names", preserveOriginalNames);
                useProgressBar = EditorGUILayout.Toggle("Show Progress Bar", useProgressBar);
                debugMode = EditorGUILayout.Toggle("Debug Logging", debugMode);
                caseSensitiveMatching = EditorGUILayout.Toggle("Case-Sensitive Matching", caseSensitiveMatching);

                GUILayout.Space(4);
                GUILayout.Label("Element ID Processing", EditorStyles.miniLabel);
                elementIdPrefix = EditorGUILayout.TextField("Remove Prefix", elementIdPrefix);
                elementIdSuffix = EditorGUILayout.TextField("Remove Suffix", elementIdSuffix);

                EditorGUI.indentLevel--;
            }
        }
    }

    private void DrawIfcSchemaOptions()
    {
        showIfcSchemaOptions = EditorGUILayout.Foldout(showIfcSchemaOptions, "IFC Schema Hierarchy");

        if (showIfcSchemaOptions)
        {
            using (new EditorGUILayout.VerticalScope(EditorStyles.helpBox))
            {
                EditorGUI.indentLevel++;

                enableIfcSchemaHierarchy = EditorGUILayout.Toggle("Enable IFC Schema Hierarchy", enableIfcSchemaHierarchy);

                GUI.enabled = enableIfcSchemaHierarchy;

                // Option to use element CSV or separate schema CSV
                useElementCsvForSchema = EditorGUILayout.Toggle("Use Element CSV for Schema", useElementCsvForSchema);

                if (!useElementCsvForSchema)
                {
                    EditorGUILayout.LabelField("Schema CSV File", EditorStyles.miniLabel);
                    schemaCSVFile = EditorGUILayout.ObjectField(schemaCSVFile, typeof(UnityEngine.Object), false);
                }
                else
                {
                    EditorGUILayout.HelpBox(
                        "The IFC schema hierarchy will be built from the element metadata CSV using:\n" +
                        "• IFC_EntityType column for entity classification\n" +
                        "• IFC_Project / IFC_Site / IFC_Building / IFC_Storey for spatial structure\n" +
                        "• IFC_PropertySets / IFC_QuantitySets for Pset/Qto nodes",
                        MessageType.Info);
                }

                ifcSchemaRootName = EditorGUILayout.TextField("Schema Root Name", ifcSchemaRootName);

                createSchemaPropertyLinks = EditorGUILayout.Toggle("Link Pset/Qto to Schema", createSchemaPropertyLinks);

                EditorGUILayout.HelpBox(
                    "Creates a second hierarchy based on the IFC schema.\n" +
                    "Structure: IFC_Schema_Root / Project / Site / Building / Storey / EntityType / Elements",
                    MessageType.Info);

                if (!useElementCsvForSchema && schemaCSVFile != null)
                {
                    if (GUILayout.Button("Preview Schema Structure"))
                    {
                        PreviewSchemaStructure();
                    }
                }

                GUI.enabled = true;

                EditorGUI.indentLevel--;
            }
        }
    }

    private void PreviewSchemaStructure()
    {
        if (schemaCSVFile == null)
        {
            Log("No schema CSV file assigned.", true);
            return;
        }

        string csvPath = AssetDatabase.GetAssetPath(schemaCSVFile);
        if (!LoadSchemaCSV(csvPath))
        {
            return;
        }

        Log("<b>=== IFC SCHEMA STRUCTURE PREVIEW ===</b>");
        LogSchemaNode(schemaRoot, 0);
    }

    private void LogSchemaNode(SchemaNode node, int depth)
    {
        string indent = new string(' ', depth * 2);
        string kindInfo = string.IsNullOrEmpty(node.Kind) ? "" : $" [{node.Kind}]";
        Log($"{indent}• {node.Name}{kindInfo}");

        foreach (var child in node.Children.OrderBy(c => c.Name))
        {
            LogSchemaNode(child, depth + 1);
        }
    }

    private void DrawColumnInformation()
    {
        showColumnInfo = EditorGUILayout.Foldout(showColumnInfo, $"Available Columns ({availableColumns?.Count ?? 0})");

        if (showColumnInfo && availableColumns != null)
        {
            using (new EditorGUILayout.VerticalScope(EditorStyles.helpBox))
            {
                EditorGUI.indentLevel++;

                var sortedColumns = availableColumns.OrderBy(c => c).ToList();
                foreach (var column in sortedColumns.Take(20)) // Show first 20 columns
                {
                    EditorGUILayout.LabelField("• " + column, EditorStyles.miniLabel);
                }

                if (sortedColumns.Count > 20)
                {
                    EditorGUILayout.LabelField($"... and {sortedColumns.Count - 20} more columns", EditorStyles.miniLabel);
                }

                EditorGUI.indentLevel--;
            }
        }
    }

    private void DrawActionButtons()
    {
        EditorGUILayout.LabelField("Actions", EditorStyles.boldLabel);

        GUI.enabled = !busy && modelRoot != null && csvFile != null;

        if (GUILayout.Button("Organize Hierarchy", GUILayout.Height(30)))
        {
            if (ValidateInputs())
            {
                ExecuteHierarchyOrganization();
            }
        }

        GUI.enabled = !busy && modelRoot != null;

        if (GUILayout.Button("Reset to Original Structure"))
        {
            ResetHierarchy();
        }

        // IFC Schema hierarchy reset button
        GUI.enabled = !busy && enableIfcSchemaHierarchy;
        if (GUILayout.Button("Clear IFC Schema Hierarchy"))
        {
            ClearSchemaHierarchy();
        }

        GUI.enabled = !busy;

        if (GUILayout.Button("Export Current Hierarchy to CSV"))
        {
            ExportHierarchyToCSV();
        }

        if (GUILayout.Button("Clear Log", GUILayout.Width(100)))
        {
            logText = "";
        }

        GUI.enabled = true;
    }

    private void ClearSchemaHierarchy()
    {
        if (string.IsNullOrEmpty(ifcSchemaRootName))
            return;

        GameObject schemaRoot = GameObject.Find(ifcSchemaRootName);
        if (schemaRoot != null)
        {
            bool confirm = EditorUtility.DisplayDialog(
                "Clear IFC Schema Hierarchy",
                $"This will delete the entire '{ifcSchemaRootName}' GameObject and all its children. Continue?",
                "Yes", "Cancel");

            if (confirm)
            {
                DestroyImmediate(schemaRoot);
                ifcSchemaRootObject = null;
                this.schemaRoot = null;
                entityNodeLookup = null;
                psetNodeLookup = null;
                qtoNodeLookup = null;
                Log($"Cleared IFC Schema hierarchy: {ifcSchemaRootName}");
                EditorSceneManager.MarkSceneDirty(EditorSceneManager.GetActiveScene());
            }
        }
        else
        {
            Log($"No IFC Schema hierarchy found with name: {ifcSchemaRootName}");
        }
    }

    private void DrawLogSection()
    {
        EditorGUILayout.BeginHorizontal();
        GUILayout.Label("Processing Log", EditorStyles.boldLabel);
        GUILayout.FlexibleSpace();
        if (GUILayout.Button("Copy Log", GUILayout.Width(80)))
        {
            GUIUtility.systemCopyBuffer = logText;
            EditorUtility.DisplayDialog("Copied", "Log copied to clipboard!", "OK");
        }
        if (GUILayout.Button("Clear Log", GUILayout.Width(80)))
        {
            logText = "";
        }
        EditorGUILayout.EndHorizontal();

        EditorGUILayout.BeginVertical(EditorStyles.helpBox);
        logScroll = EditorGUILayout.BeginScrollView(logScroll, GUILayout.Height(250));

        var logStyle = new GUIStyle(EditorStyles.textArea)
        {
            wordWrap = false,
            richText = true,
            font = Font.CreateDynamicFontFromOSFont("Consolas", 11)
        };

        EditorGUILayout.SelectableLabel(logText, logStyle, GUILayout.ExpandHeight(true), GUILayout.ExpandWidth(true), GUILayout.MinHeight(230));
        EditorGUILayout.EndScrollView();
        EditorGUILayout.EndVertical();
    }

    private void LoadColumnInformation()
    {
        try
        {
            string csvPath = AssetDatabase.GetAssetPath(csvFile);
            if (File.Exists(csvPath))
            {
                string[] lines = File.ReadAllLines(csvPath);
                if (lines.Length > 0)
                {
                    var headers = ParseCSVLine(lines[0]);
                    availableColumns = new HashSet<string>(headers.Select(h => h.Trim().TrimStart('\uFEFF')));
                }
            }
        }
        catch (Exception e)
        {
            Log($"Error loading column information: {e.Message}", true);
        }
    }

    private bool ValidateInputs()
    {
        if (modelRoot == null)
        {
            EditorUtility.DisplayDialog("Error", $"Please assign a {bimType} model root object.", "OK");
            return false;
        }

        if (csvFile == null)
        {
            EditorUtility.DisplayDialog("Error", "Please assign a CSV metadata file.", "OK");
            return false;
        }

        if (CurrentHierarchyMode == HierarchyMode.Custom && string.IsNullOrWhiteSpace(CurrentCustomHierarchy))
        {
            EditorUtility.DisplayDialog("Error", "Please specify custom hierarchy columns.", "OK");
            return false;
        }

        return true;
    }

    private void ExecuteHierarchyOrganization()
    {
        busy = true;

        // Initialize performance tracking
        performanceTimer = System.Diagnostics.Stopwatch.StartNew();
        memoryAtStart = GC.GetTotalMemory(false);
        peakMemoryUsed = 0;
        csvRowsLoaded = 0;
        hierarchyNodesCreated = 0;
        schemaReferencesCreated = 0;

        try
        {
            Log($"<b>=== {bimType.ToString().ToUpper()} HIERARCHY ORGANIZATION START ===</b>");
            Log($"Organization Mode: {CurrentHierarchyMode}");
            Log($"BIM Platform: {bimType}");
            Log($"Start Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            Log($"Initial Memory: {FormatMemorySize(memoryAtStart)}");

            // Check if IFC Schema mode is selected
            bool useIfcSchemaMode = CurrentHierarchyMode == HierarchyMode.IfcSchema;

            // For non-IfcSchema modes, check if additional IFC schema hierarchy is enabled
            Log($"IFC Schema Hierarchy (secondary): {(enableIfcSchemaHierarchy && !useIfcSchemaMode ? "Enabled" : "Disabled")}");

            // Step 1: Unpack prefab if needed
            UnpackPrefabIfNeeded(modelRoot);
            UpdatePeakMemory();

            // Step 2: Load element metadata
            var csvLoadTimer = System.Diagnostics.Stopwatch.StartNew();
            if (!LoadElementMetadata())
            {
                Log("Failed to load element metadata from CSV", true);
                return;
            }
            csvLoadTimer.Stop();
            Log($"CSV Loading Time: {csvLoadTimer.ElapsedMilliseconds}ms");
            UpdatePeakMemory();

            // Step 3: Build secondary IFC schema hierarchy (only if enabled AND not using IFC Schema mode)
            // When IFC Schema mode is selected, the primary hierarchy IS the IFC schema structure
            if (enableIfcSchemaHierarchy && !useIfcSchemaMode)
            {
                var schemaTimer = System.Diagnostics.Stopwatch.StartNew();
                if (useElementCsvForSchema)
                {
                    // Build schema hierarchy from element metadata
                    if (BuildSchemaFromElementMetadata())
                    {
                        BuildSchemaHierarchyInScene();
                    }
                    else
                    {
                        Log("Warning: Failed to build IFC schema from element metadata.", true);
                    }
                }
                else if (schemaCSVFile != null)
                {
                    // Build schema from separate schema CSV file
                    string schemaPath = AssetDatabase.GetAssetPath(schemaCSVFile);
                    if (LoadSchemaCSV(schemaPath))
                    {
                        BuildSchemaHierarchyInScene();
                    }
                    else
                    {
                        Log("Warning: Failed to load IFC schema. Schema hierarchy will not be created.", true);
                    }
                }
                else
                {
                    Log("Warning: IFC Schema Hierarchy enabled but no schema source specified.", true);
                }
                schemaTimer.Stop();
                Log($"Schema Hierarchy Build Time: {schemaTimer.ElapsedMilliseconds}ms");
                UpdatePeakMemory();
            }

            // Step 4: Execute hierarchy organization (works for ALL modes including IFC Schema)
            var organizeTimer = System.Diagnostics.Stopwatch.StartNew();
            OrganizeHierarchy();
            organizeTimer.Stop();
            Log($"Hierarchy Organization Time: {organizeTimer.ElapsedMilliseconds}ms");
            UpdatePeakMemory();

            EditorUtility.SetDirty(modelRoot);
            if (ifcSchemaRootObject != null)
            {
                EditorUtility.SetDirty(ifcSchemaRootObject);
            }
            EditorSceneManager.MarkSceneDirty(EditorSceneManager.GetActiveScene());

            performanceTimer.Stop();
            Log($"<b>Hierarchy organization completed successfully using {CurrentHierarchyMode} mode.</b>");
            if (enableIfcSchemaHierarchy && !useIfcSchemaMode && schemaRoot != null)
            {
                Log($"<b>Secondary IFC Schema hierarchy created under '{ifcSchemaRootName}'.</b>");
            }
        }
        catch (Exception e)
        {
            performanceTimer.Stop();
            Log($"Critical error during hierarchy organization: {e.Message}", true);
            EditorUtility.DisplayDialog("Error", $"Failed to organize hierarchy: {e.Message}", "OK");
        }
        finally
        {
            busy = false;
            if (useProgressBar)
            {
                EditorUtility.ClearProgressBar();
            }
        }
    }

    private void UpdatePeakMemory()
    {
        long currentMemory = GC.GetTotalMemory(false);
        if (currentMemory > peakMemoryUsed)
        {
            peakMemoryUsed = currentMemory;
        }
    }

    private string FormatMemorySize(long bytes)
    {
        string[] sizes = { "B", "KB", "MB", "GB" };
        double len = bytes;
        int order = 0;
        while (len >= 1024 && order < sizes.Length - 1)
        {
            order++;
            len = len / 1024;
        }
        return $"{len:0.##} {sizes[order]}";
    }

    private void UnpackPrefabIfNeeded(GameObject root)
    {
        var status = PrefabUtility.GetPrefabInstanceStatus(root);
        if (status == PrefabInstanceStatus.Connected)
        {
            bool unpack = EditorUtility.DisplayDialog(
                "Prefab Detected",
                "The selected object is a prefab instance. Unpack it to modify hierarchy?",
                "Yes", "Cancel");

            if (unpack)
            {
                PrefabUtility.UnpackPrefabInstance(root, PrefabUnpackMode.Completely, InteractionMode.AutomatedAction);
                Log($"Unpacked prefab: {root.name}");
            }
            else
            {
                throw new OperationCanceledException("User cancelled prefab unpacking.");
            }
        }
    }

    private bool LoadElementMetadata()
    {
        string csvPath = AssetDatabase.GetAssetPath(csvFile);
        elementMetadata = ReadCsvFile(csvPath);

        if (elementMetadata == null || elementMetadata.Count == 0)
        {
            Log("Failed to load CSV data or CSV file is empty", true);
            return false;
        }

        Log($"Loaded metadata for {elementMetadata.Count} building elements");
        return true;
    }

    private void OrganizeHierarchy()
    {
        var allChildren = GetAllChildElements();
        int totalElements = allChildren.Count;
        int processedElements = 0;
        int matchedElements = 0;
        int unmatchedCount = 0;

        Log($"Found {totalElements} child elements to process (no limit on element count)");

        Transform unmatchedParent = null;
        if (createUnmatchedGroup)
        {
            unmatchedParent = GetOrCreateChild(modelRoot.transform, "Unmatched_Elements");
        }

        var hierarchyCache = new Dictionary<string, Transform>(StringComparer.OrdinalIgnoreCase);
        var organizationStats = new Dictionary<string, int>();
        var unmatchedElements = new List<Transform>();

        // Process ALL elements without any limit
        foreach (Transform child in allChildren)
        {
            if (useProgressBar && processedElements % 100 == 0) // Update progress every 100 elements for performance
            {
                float progress = (float)processedElements / totalElements;
                EditorUtility.DisplayProgressBar("Organizing Hierarchy",
                    $"Processing element {processedElements + 1} of {totalElements} ({matchedElements} matched)", progress);
            }

            // Try multiple matching strategies to find metadata
            var metadata = TryFindMetadata(child.name);

            if (metadata != null)
            {
                OrganizeElement(child, metadata, hierarchyCache, organizationStats);
                matchedElements++;

                if (debugMode)
                {
                    Log($"  ✓ MATCHED: {child.name}");
                }
            }
            else
            {
                unmatchedElements.Add(child);
                unmatchedCount++;

                if (debugMode)
                {
                    string elementId = ExtractElementId(child.name);
                    Log($"  ✗ UNMATCHED: {child.name} (ID: {elementId})");
                }
            }

            processedElements++;
        }

        // Move unmatched elements to unmatched parent AFTER processing all matched elements
        // This ensures we don't accidentally skip any potential matches
        if (unmatchedParent != null)
        {
            foreach (var element in unmatchedElements)
            {
                element.SetParent(unmatchedParent, true);
            }
        }

        LogOrganizationResults(totalElements, matchedElements, unmatchedCount, organizationStats);
    }

    /// <summary>
    /// Tries multiple strategies to find metadata for an element.
    /// Returns null if no metadata is found.
    /// </summary>
    private Dictionary<string, string> TryFindMetadata(string elementName)
    {
        if (string.IsNullOrEmpty(elementName) || elementMetadata == null)
            return null;

        // Strategy 1: Direct element ID extraction (primary method)
        string elementId = ExtractElementId(elementName);
        string lookupKey = CleanElementId(elementId);
        if (!caseSensitiveMatching && !string.IsNullOrEmpty(lookupKey))
            lookupKey = lookupKey.ToLower();

        if (!string.IsNullOrEmpty(lookupKey) && elementMetadata.ContainsKey(lookupKey))
            return elementMetadata[lookupKey];

        // Strategy 2: Try the full element name (cleaned)
        string fullNameKey = CleanElementId(elementName.Replace("(Clone)", "").Trim());
        if (!caseSensitiveMatching && !string.IsNullOrEmpty(fullNameKey))
            fullNameKey = fullNameKey.ToLower();

        if (!string.IsNullOrEmpty(fullNameKey) && elementMetadata.ContainsKey(fullNameKey))
            return elementMetadata[fullNameKey];

        // Strategy 3: Try with different ID extraction patterns
        var alternativeIds = ExtractAlternativeElementIds(elementName);
        foreach (var altId in alternativeIds)
        {
            string altKey = CleanElementId(altId);
            if (!caseSensitiveMatching && !string.IsNullOrEmpty(altKey))
                altKey = altKey.ToLower();

            if (!string.IsNullOrEmpty(altKey) && elementMetadata.ContainsKey(altKey))
                return elementMetadata[altKey];
        }

        // Strategy 4: Partial match - search through all metadata keys for a containing match
        if (!string.IsNullOrEmpty(elementId))
        {
            foreach (var kvp in elementMetadata)
            {
                string metaKey = caseSensitiveMatching ? kvp.Key : kvp.Key.ToLower();
                string searchId = caseSensitiveMatching ? elementId : elementId.ToLower();

                // Check if either contains the other
                if (metaKey.Contains(searchId) || searchId.Contains(metaKey))
                {
                    if (debugMode)
                        Log($"    Partial match found: '{elementId}' matched with '{kvp.Key}'");
                    return kvp.Value;
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Extracts alternative element IDs using various patterns.
    /// </summary>
    private List<string> ExtractAlternativeElementIds(string elementName)
    {
        var alternatives = new List<string>();
        if (string.IsNullOrEmpty(elementName))
            return alternatives;

        elementName = elementName.Replace("(Clone)", "").Trim();

        // Try extracting numeric IDs from various positions
        var matches = Regex.Matches(elementName, @"\d+");
        foreach (Match match in matches)
        {
            if (!alternatives.Contains(match.Value))
                alternatives.Add(match.Value);
        }

        // Try extracting after common separators
        string[] separators = { "_", "-", ".", " ", ":" };
        foreach (var sep in separators)
        {
            var parts = elementName.Split(new[] { sep }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length > 1)
            {
                // Add last part
                if (!alternatives.Contains(parts[parts.Length - 1]))
                    alternatives.Add(parts[parts.Length - 1]);

                // Add first part
                if (!alternatives.Contains(parts[0]))
                    alternatives.Add(parts[0]);
            }
        }

        // Try the name without common prefixes
        string[] prefixes = { "ifc", "bim", "model", "element" };
        foreach (var prefix in prefixes)
        {
            if (elementName.ToLower().StartsWith(prefix))
            {
                string stripped = elementName.Substring(prefix.Length).TrimStart('_', '-', ' ');
                if (!string.IsNullOrEmpty(stripped) && !alternatives.Contains(stripped))
                    alternatives.Add(stripped);
            }
        }

        return alternatives;
    }

    private List<Transform> GetAllChildElements()
    {
        var allChildren = new List<Transform>();
        var hierarchyNodes = new HashSet<Transform>(); // Track hierarchy nodes to exclude

        // First pass: Identify all transforms
        var allTransforms = modelRoot.GetComponentsInChildren<Transform>(true);

        // Identify potential hierarchy container nodes (nodes that only contain other nodes, no mesh/renderer)
        foreach (Transform t in allTransforms)
        {
            if (t == modelRoot.transform)
                continue;

            // Check if this is a pure hierarchy node (no visual components)
            bool hasVisualComponent = t.GetComponent<MeshRenderer>() != null ||
                                      t.GetComponent<MeshFilter>() != null ||
                                      t.GetComponent<SkinnedMeshRenderer>() != null ||
                                      t.GetComponent<Renderer>() != null;

            // If it has children but no visual components, it might be a hierarchy container
            // However, we should still process it if it could have metadata
            if (!hasVisualComponent && t.childCount > 0)
            {
                // This might be a hierarchy container, but don't exclude it yet
                // We'll check if it has metadata
            }
        }

        // Second pass: Collect all elements that should be organized
        // Include both leaf nodes AND nodes with visual components
        foreach (Transform t in allTransforms)
        {
            if (t == modelRoot.transform)
                continue;

            bool hasVisualComponent = t.GetComponent<MeshRenderer>() != null ||
                                      t.GetComponent<MeshFilter>() != null ||
                                      t.GetComponent<SkinnedMeshRenderer>() != null ||
                                      t.GetComponent<Renderer>() != null;

            // Include if:
            // 1. It's a leaf node (no children), OR
            // 2. It has visual components (mesh, renderer, etc.), OR  
            // 3. It has an IfcPropertyData component (indicating it's a BIM element)
            bool isLeaf = t.childCount == 0;
            bool hasBimData = t.GetComponent<IfcPropertyData>() != null;

            if (isLeaf || hasVisualComponent || hasBimData)
            {
                allChildren.Add(t);
            }
        }

        // Sort by hierarchy depth to process parents before children (optional, helps with organization)
        allChildren.Sort((a, b) => GetTransformDepth(a).CompareTo(GetTransformDepth(b)));

        Log($"Identified {allChildren.Count} elements to organize (leaf nodes: {allChildren.Count(t => t.childCount == 0)}, with visual components: {allChildren.Count(t => t.GetComponent<Renderer>() != null)})");

        return allChildren;
    }

    private int GetTransformDepth(Transform t)
    {
        int depth = 0;
        Transform current = t;
        while (current != null && current != modelRoot.transform)
        {
            depth++;
            current = current.parent;
        }
        return depth;
    }

    private void OrganizeElement(Transform element, Dictionary<string, string> metadata,
                               Dictionary<string, Transform> hierarchyCache, Dictionary<string, int> organizationStats)
    {
        var hierarchyLevels = GetHierarchyLevels(metadata);

        Transform currentParent = modelRoot.transform;

        foreach (string levelValue in hierarchyLevels)
        {
            if (!string.IsNullOrWhiteSpace(levelValue))
            {
                currentParent = GetOrCreateChild(currentParent, SanitizeName(levelValue), hierarchyCache);
            }
        }

        // Track organization statistics
        string hierarchyPath = string.Join(" → ", hierarchyLevels.Where(l => !string.IsNullOrWhiteSpace(l)));
        organizationStats[hierarchyPath] = organizationStats.GetValueOrDefault(hierarchyPath, 0) + 1;

        // Preserve original name or use sanitized version
        if (!preserveOriginalNames)
        {
            element.name = SanitizeName(element.name);
        }

        // Parent element in spatial hierarchy
        element.SetParent(currentParent, true);

        // Parent element reference in IFC schema hierarchy (if enabled)
        if (enableIfcSchemaHierarchy && schemaRoot != null)
        {
            ParentElementUnderSchemaHierarchy(element, metadata);
        }

        // Attach Pset/Qto property data to element (if enabled)
        if (createSchemaPropertyLinks)
        {
            AttachPropertyDataToElement(element, metadata);
        }
    }

    private List<string> GetHierarchyLevels(Dictionary<string, string> metadata)
    {
        var levels = new List<string>();

        // Helper function to get metadata with multiple fallback column names
        Func<string[], string, string> getWithFallbacks = (columns, defaultVal) =>
        {
            foreach (var col in columns)
            {
                string value = GetMetadataValue(metadata, col, "");
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }
            return defaultVal;
        };

        switch (CurrentHierarchyMode)
        {
            case HierarchyMode.ByCategory:
                levels.Add(getWithFallbacks(new[] {
                    "Element_Type",                          // Archicad primary element type
                    "General Parameters.Element Type",       // Archicad full column name
                    "Category",
                    "IFC_EntityType",
                    "IFC4RV_EntityType",
                    "Type"
                }, "Categorized_Elements"));
                break;

            case HierarchyMode.FlatByCategory:
                // Flat structure: only Category level, no further nesting
                levels.Add(getWithFallbacks(new[] {
                    "Element_Type",                          // Archicad primary element type
                    "General Parameters.Element Type",       // Archicad full column name
                    "Category",
                    "IFC_EntityType",
                    "IFC4RV_EntityType",
                    "Type"
                }, "Categorized_Elements"));
                break;

            case HierarchyMode.ByFamily:
                levels.AddRange(new[] {
                    getWithFallbacks(new[] {
                        "Element_Type",                          // Archicad primary element type
                        "General Parameters.Element Type",       // Archicad full column name
                        "Family",
                        "Category",
                        "IFC_EntityType",
                        "IFC4RV_EntityType"
                    }, "Categorized_Elements"),
                    getWithFallbacks(new[] {
                        "Construction.Composite Name",           // Archicad composite name
                        "Construction.Complex Profile Name",     // Archicad profile name
                        "Construction.Structure Type",           // Archicad structure type
                        "Type_Type Name",
                        "Type",
                        "Type Name"
                    }, "Default_Type")
                });
                break;

            case HierarchyMode.ByMaterial:
                levels.AddRange(new[] {
                    getWithFallbacks(new[] {
                        "Construction.Composite Name",           // Archicad composite material
                        "Construction.Complex Profile Name",     // Archicad profile material
                        "Material",
                        "Materials"
                    }, "Default_Material"),
                    getWithFallbacks(new[] {
                        "Element_Type",                          // Archicad primary element type
                        "General Parameters.Element Type",       // Archicad full column name
                        "Category",
                        "IFC_EntityType",
                        "IFC4RV_EntityType"
                    }, "Categorized_Elements")
                });
                break;

            case HierarchyMode.ByLevel:
                levels.AddRange(new[] {
                    getWithFallbacks(new[] {
                        "General Parameters.Top Link Story",     // Archicad story/level
                        "Positioning.Elevation to Project Zero", // Archicad elevation
                        "IFC4RV_Hierarchy",
                        "Level",
                        "IFC_Storey",
                        "Storey",
                        "Floor"
                    }, "Default_Level"),
                    getWithFallbacks(new[] {
                        "Element_Type",                          // Archicad primary element type
                        "General Parameters.Element Type",       // Archicad full column name
                        "Category",
                        "IFC_EntityType",
                        "IFC4RV_EntityType"
                    }, "Categorized_Elements"),
                    getWithFallbacks(new[] {
                        "Construction.Composite Name",           // Archicad composite name
                        "Construction.Complex Profile Name",     // Archicad profile name
                        "Family",
                        "Type"
                    }, "Default_Family")
                });
                break;

            case HierarchyMode.BySystem:
                levels.AddRange(new[] {
                    getWithFallbacks(new[] {
                        "Model View.Layer Name",                 // Archicad layer as system proxy
                        "System Name",
                        "SystemName",
                        "System"
                    }, "No_System"),
                    getWithFallbacks(new[] {
                        "Element_Type",                          // Archicad primary element type
                        "General Parameters.Element Type",       // Archicad full column name
                        "Category",
                        "IFC_EntityType",
                        "IFC4RV_EntityType"
                    }, "Categorized_Elements"),
                    getWithFallbacks(new[] {
                        "Construction.Composite Name",           // Archicad composite name
                        "Construction.Complex Profile Name",     // Archicad profile name
                        "Family",
                        "Type"
                    }, "Default_Family")
                });
                break;

            case HierarchyMode.Modular:
                levels.AddRange(new[] {
                    getWithFallbacks(new[] { "Modular_Group_Id", "Modular Group", "ModularGroup" }, "Non_Modular"),
                    getWithFallbacks(new[] { "Element_Type", "Category", "IFC_EntityType", "IFC4RV_EntityType" }, "Categorized_Elements"),
                    getWithFallbacks(new[] { "Family", "Type" }, "Default_Family"),
                    getWithFallbacks(new[] { "Type_Type Name", "Type", "Type Name" }, "Default_Type")
                });
                break;

            case HierarchyMode.SystemDetailed:
                levels.AddRange(new[] {
                    getWithFallbacks(new[] { "System Name", "SystemName", "System" }, "No_System"),
                    getWithFallbacks(new[] { "System Type", "SystemType" }, "Default_System_Type"),
                    getWithFallbacks(new[] { "Element_Type", "Category", "IFC_EntityType", "IFC4RV_EntityType" }, "Categorized_Elements"),
                    getWithFallbacks(new[] { "Family", "Type" }, "Default_Family")
                });
                break;

            case HierarchyMode.MaterialDetailed:
                levels.AddRange(new[] {
                    getWithFallbacks(new[] {
                        "Construction.Composite Name",           // Archicad composite material
                        "Construction.Complex Profile Name",     // Archicad profile material
                        "Material",
                        "Materials"
                    }, "Default_Material"),
                    getWithFallbacks(new[] {
                        "Element_Type",                          // Archicad primary element type
                        "General Parameters.Element Type",       // Archicad full column name
                        "Category",
                        "IFC_EntityType",
                        "IFC4RV_EntityType"
                    }, "Categorized_Elements"),
                    getWithFallbacks(new[] {
                        "Construction.Structure Type",           // Archicad structure type
                        "Family",
                        "Type"
                    }, "Default_Family"),
                    getWithFallbacks(new[] {
                        "General Parameters.Element ID",         // Archicad element ID
                        "Type_Type Name",
                        "Type",
                        "Type Name"
                    }, "Default_Type")
                });
                break;

            case HierarchyMode.LevelDetailed:
                levels.AddRange(new[] {
                    getWithFallbacks(new[] {
                        "General Parameters.Top Link Story",     // Archicad story/level
                        "Positioning.Elevation to Project Zero", // Archicad elevation
                        "IFC4RV_Hierarchy",
                        "Level",
                        "IFC_Storey",
                        "Storey",
                        "Floor"
                    }, "Default_Level"),
                    getWithFallbacks(new[] {
                        "Element_Type",                          // Archicad primary element type
                        "General Parameters.Element Type",       // Archicad full column name
                        "Category",
                        "IFC_EntityType",
                        "IFC4RV_EntityType"
                    }, "Categorized_Elements"),
                    getWithFallbacks(new[] {
                        "Construction.Composite Name",           // Archicad composite name
                        "Construction.Complex Profile Name",     // Archicad profile name
                        "Family",
                        "Type"
                    }, "Default_Family"),
                    getWithFallbacks(new[] {
                        "Construction.Structure Type",           // Archicad structure type
                        "Type_Type Name",
                        "Type",
                        "Type Name"
                    }, "Default_Type")
                });
                break;

            case HierarchyMode.ModularDetailed:
                levels.AddRange(new[] {
                    getWithFallbacks(new[] { "Modular_Group_Id", "Modular Group", "ModularGroup" }, "Non_Modular"),
                    getWithFallbacks(new[] { "IFC4RV_Hierarchy", "Level", "IFC_Storey", "Storey" }, "Default_Level"),
                    getWithFallbacks(new[] { "System Name", "SystemName", "System" }, "No_System"),
                    getWithFallbacks(new[] { "Element_Type", "Category", "IFC_EntityType", "IFC4RV_EntityType" }, "Categorized_Elements"),
                    getWithFallbacks(new[] { "Family", "Type" }, "Default_Family"),
                    getWithFallbacks(new[] { "Type_Type Name", "Type", "Type Name" }, "Default_Type")
                });
                break;

            case HierarchyMode.ConstructionDiscipline:
                levels.AddRange(new[] {
                    ClassifyConstructionDiscipline(getWithFallbacks(new[] {
                        "Model View.Layer Name",                 // Archicad layer name
                        "Layer Name",
                        "Layer",
                        "LayerName"
                    }, "NoLayer")),
                    getWithFallbacks(new[] {
                        "Element_Type",                          // Archicad primary element type
                        "General Parameters.Element Type",       // Archicad full column name
                        "Category",
                        "IFC_EntityType",
                        "IFC4RV_EntityType"
                    }, "Categorized_Elements")
                });
                break;

            case HierarchyMode.SpatialDiscipline:
                levels.AddRange(new[] {
                    getWithFallbacks(new[] { "Related Zone Name", "Zone Name", "Zone", "Room" }, "NoZone"),
                    getWithFallbacks(new[] { "Element_Type", "Category", "IFC_EntityType", "IFC4RV_EntityType" }, "Categorized_Elements")
                });
                break;

            case HierarchyMode.MultiLevelClassification:
                levels.AddRange(new[] {
                    getWithFallbacks(new[] { "Element_Type", "Category", "IFC_EntityType", "IFC4RV_EntityType" }, "Categorized_Elements"),
                    DetermineSubCategory(metadata)
                });
                break;

            case HierarchyMode.IfcSchema:
                // IFC Schema hierarchy: IfcProject → IfcSite → IfcBuilding → IfcBuildingStorey → EntityType → Elements
                levels.AddRange(new[] {
                    getWithFallbacks(new[] { "IFC4RV_Project", "IFC_Project", "Project", "IfcProject" }, "Default_Project"),
                    getWithFallbacks(new[] { "IFC4RV_Site", "IFC_Site", "Site", "IfcSite" }, "Default_Site"),
                    getWithFallbacks(new[] { "IFC4RV_Building", "IFC_Building", "Building", "IfcBuilding" }, "Default_Building"),
                    getWithFallbacks(new[] { "IFC4RV_Hierarchy", "IFC_Storey", "Level", "Storey", "Floor", "IfcBuildingStorey" }, "Default_Storey"),
                    getWithFallbacks(new[] { "IFC4RV_EntityType", "Element_Type", "IFC_EntityType", "Category", "IfcClass" }, "Categorized_Elements")
                });
                break;

            case HierarchyMode.Custom:
                var customColumns = CurrentCustomHierarchy.Split(',').Select(c => c.Trim()).Where(c => !string.IsNullOrEmpty(c));
                foreach (var column in customColumns)
                {
                    // For custom columns, try exact match first, then fallback to similar column names
                    string value = GetMetadataValue(metadata, column, "");
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        // Try case-insensitive match
                        var matchingKey = metadata.Keys.FirstOrDefault(k =>
                            k.Equals(column, StringComparison.OrdinalIgnoreCase) ||
                            k.Replace("_", " ").Equals(column.Replace("_", " "), StringComparison.OrdinalIgnoreCase) ||
                            k.Replace(" ", "_").Equals(column.Replace(" ", "_"), StringComparison.OrdinalIgnoreCase));

                        if (matchingKey != null)
                            value = metadata[matchingKey];
                        else
                            value = $"Default_{column.Replace(" ", "_")}";
                    }
                    levels.Add(value);
                }
                break;
        }

        return levels;
    }

    private string GetMetadataValue(Dictionary<string, string> metadata, string key, string defaultValue)
    {
        // Try exact match first
        if (metadata.TryGetValue(key, out string value) && !string.IsNullOrWhiteSpace(value))
        {
            return value.Trim();
        }

        // Try case-insensitive match
        foreach (var kvp in metadata)
        {
            if (kvp.Key.Equals(key, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(kvp.Value))
            {
                return kvp.Value.Trim();
            }
        }

        // Try with underscore/space variations
        string normalizedKey = key.Replace("_", " ").Replace("-", " ");
        foreach (var kvp in metadata)
        {
            string normalizedMetaKey = kvp.Key.Replace("_", " ").Replace("-", " ");
            if (normalizedMetaKey.Equals(normalizedKey, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(kvp.Value))
            {
                return kvp.Value.Trim();
            }
        }

        return defaultValue;
    }

    private Transform GetOrCreateChild(Transform parent, string childName, Dictionary<string, Transform> cache = null)
    {
        if (string.IsNullOrWhiteSpace(childName))
            childName = "Unknown";

        // Use cache if provided
        if (cache != null)
        {
            string cacheKey = $"{parent.GetInstanceID()}_{childName}";
            if (cache.TryGetValue(cacheKey, out Transform cachedChild))
                return cachedChild;
        }

        Transform child = parent.Find(childName);
        if (child == null)
        {
            GameObject newChild = new GameObject(childName);
            newChild.transform.SetParent(parent, false);
            child = newChild.transform;
            hierarchyNodesCreated++; // Track new nodes created
        }

        // Cache the result if cache is provided
        if (cache != null)
        {
            string cacheKey = $"{parent.GetInstanceID()}_{childName}";
            cache[cacheKey] = child;
        }

        return child;
    }

    private string ExtractElementId(string elementName)
    {
        if (string.IsNullOrEmpty(elementName))
            return null;

        // Clean up common suffixes
        elementName = elementName.Replace("(Clone)", "").Trim();

        if (bimType == BIMType.Archicad)
        {
            // ArchiCAD: Use the object name directly
            return elementName;
        }
        else // Revit
        {
            // Revit: Extract element ID from the end after the last underscore
            int lastUnderscoreIndex = elementName.LastIndexOf('_');
            if (lastUnderscoreIndex >= 0 && lastUnderscoreIndex < elementName.Length - 1)
            {
                string candidate = elementName.Substring(lastUnderscoreIndex + 1);
                if (int.TryParse(candidate, out _)) // Check if it's numeric
                {
                    return candidate;
                }
            }

            // Fallback: Extract numeric part from the end
            var match = Regex.Match(elementName, @"(\d+)$");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }

            // Final fallback: Use the full name
            return elementName;
        }
    }

    private string CleanElementId(string rawId)
    {
        if (string.IsNullOrEmpty(rawId)) return rawId;

        string cleanId = rawId.Trim();

        // Apply user-defined prefix/suffix removal
        if (!string.IsNullOrEmpty(elementIdPrefix) && cleanId.StartsWith(elementIdPrefix))
            cleanId = cleanId.Substring(elementIdPrefix.Length);
        if (!string.IsNullOrEmpty(elementIdSuffix) && cleanId.EndsWith(elementIdSuffix))
            cleanId = cleanId.Substring(0, cleanId.Length - elementIdSuffix.Length);

        return cleanId;
    }

    private string SanitizeName(string name)
    {
        if (string.IsNullOrEmpty(name))
            return "Unknown";

        // Remove invalid characters and normalize spaces
        return name.Trim()
                   .Replace('/', '_')
                   .Replace('\\', '_')
                   .Replace(':', '_')
                   .Replace('*', '_')
                   .Replace('?', '_')
                   .Replace('"', '_')
                   .Replace('<', '_')
                   .Replace('>', '_')
                   .Replace('|', '_');
    }

    private string ClassifyConstructionDiscipline(string layerName)
    {
        if (string.IsNullOrEmpty(layerName) || layerName == "NoLayer") return "Undefined_Discipline";

        string layer = layerName.ToLower();

        // Handle Archicad layer naming convention: "ENZ - Category - SubCategory"
        // Examples: "ENZ - Structure - Wall", "ENZ - Interior - Partition", "ENZ - Finish - Cladding"
        if (layer.Contains(" - "))
        {
            string[] parts = layerName.Split(new[] { " - " }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 2)
            {
                string category = parts[1].Trim().ToLower();

                // Map Archicad categories to disciplines
                if (category == "structure" || category.Contains("struct"))
                    return "01_Structural";
                if (category == "interior" || category.Contains("interior"))
                    return "02_Interior";
                if (category == "finish" || category.Contains("finish"))
                    return "03_Finish";
                if (category == "exterior" || category.Contains("exterior"))
                    return "04_Exterior";
                if (category == "site" || category.Contains("site"))
                    return "05_Site";

                // Return the category as-is if not matched
                return $"{SanitizeName(parts[1].Trim())}";
            }
        }

        // Structural discipline (fallback patterns)
        if (layer.Contains("structure") || layer.Contains("structural") ||
            layer.Contains("base") || layer.Contains("plate") || layer.Contains("beam") ||
            layer.Contains("column") || layer.Contains("purlin") || layer.Contains("볼트") ||
            layer.Contains("철골") || layer.Contains("기초") || layer.Contains("구조"))
            return "01_Structural";

        // Architectural discipline  
        if (layer.Contains("가구") || layer.Contains("문") || layer.Contains("창") ||
            layer.Contains("마감") || layer.Contains("벽") || layer.Contains("바닥") ||
            layer.Contains("천장") || layer.Contains("계단"))
            return "02_Architectural";

        // MEP discipline
        if (layer.Contains("덕트") || layer.Contains("배관") || layer.Contains("전기") ||
            layer.Contains("통신") || layer.Contains("기계") || layer.Contains("설비"))
            return "03_MEP";

        return $"04_Other_{SanitizeName(layerName)}";
    }

    private string DetermineSubCategory(Dictionary<string, string> metadata)
    {
        string elementType = GetMetadataValue(metadata, "Element_Type", "");

        // For structural elements, use profile category
        if (IsStructuralElement(elementType))
        {
            string profile = GetMetadataValue(metadata, "Profile Category", "");
            if (!string.IsNullOrEmpty(profile) && profile != "NoProfile")
            {
                return SanitizeName(profile);
            }
        }

        // For architectural elements, use library part name
        if (IsArchitecturalElement(elementType))
        {
            string library = GetMetadataValue(metadata, "Library Part Name", "");
            if (!string.IsNullOrEmpty(library) && library != "NoLibraryPart")
            {
                return SanitizeName(library);
            }
        }

        return "Standard";
    }

    private bool IsStructuralElement(string elementType)
    {
        if (string.IsNullOrEmpty(elementType)) return false;
        string type = elementType.ToLower();
        return type.Contains("beam") || type.Contains("column") ||
               type.Contains("slab") || type.Contains("wall");
    }

    private bool IsArchitecturalElement(string elementType)
    {
        if (string.IsNullOrEmpty(elementType)) return false;
        string type = elementType.ToLower();
        return type.Contains("door") || type.Contains("window") ||
               type.Contains("object") || type.Contains("stair");
    }

    private Dictionary<string, Dictionary<string, string>> ReadCsvFile(string filePath)
    {
        if (!File.Exists(filePath))
        {
            Log($"CSV file not found: {filePath}", true);
            return null;
        }

        var data = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

        try
        {
            string[] lines = File.ReadAllLines(filePath);

            if (lines.Length < 2)
            {
                Log("CSV file must have headers and at least one data row", true);
                return null;
            }

            // Parse headers and clean BOM
            var headers = ParseCSVLine(lines[0]);
            for (int i = 0; i < headers.Length; i++)
                headers[i] = headers[i].Trim().TrimStart('\uFEFF');

            // Find Element ID column (support multiple naming conventions)
            // For ArchiCAD, prioritize columns that match FBX object naming
            string[] possibleIdColumns = {
                "ID and Categories.Element ID",   // Primary Archicad column (W-001, W-002, etc.)
                "General Parameters.Element ID",  // Alternative Archicad CSV format
                "General Parameters.Unique ID",   // Unique ID column  
                "Element ID",
                "Element_ID",
                "ElementID",
                "ID",
                "Element_Id",
                "Name",
                "Element Name",
                "ElementName"
            };
            int elementIdIndex = -1;

            for (int i = 0; i < headers.Length; i++)
            {
                if (possibleIdColumns.Any(col => col.Equals(headers[i], StringComparison.OrdinalIgnoreCase)))
                {
                    elementIdIndex = i;
                    break;
                }
            }

            if (elementIdIndex == -1)
            {
                Log($"Element ID column not found. Available columns: {string.Join(", ", headers.Take(10))}", true);
                return null;
            }

            Log($"Using Element ID column: '{headers[elementIdIndex]}' (Column {elementIdIndex})");

            // Also find additional columns that can be used for alternative lookups
            int nameColumnIndex = Array.FindIndex(headers, h =>
                h.Equals("Name", StringComparison.OrdinalIgnoreCase) ||
                h.Equals("Element Name", StringComparison.OrdinalIgnoreCase));
            int globalIdIndex = Array.FindIndex(headers, h =>
                h.Equals("GlobalId", StringComparison.OrdinalIgnoreCase) ||
                h.Equals("Global_Id", StringComparison.OrdinalIgnoreCase) ||
                h.Equals("IFC_GlobalId", StringComparison.OrdinalIgnoreCase));

            int totalRows = 0;
            int duplicateRows = 0;

            // Process ALL data rows without any limit
            for (int i = 1; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i]))
                    continue;

                var cells = ParseCSVLine(lines[i]);
                if (cells.Length <= elementIdIndex)
                    continue;

                string elementId = cells[elementIdIndex].Trim();
                if (string.IsNullOrEmpty(elementId) || elementId == "<Undefined>")
                    continue;

                totalRows++;
                csvRowsLoaded++; // Track for performance summary

                var rowData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int j = 0; j < Math.Min(headers.Length, cells.Length); j++)
                {
                    string value = cells[j].Trim();
                    if (!string.IsNullOrEmpty(value) && value != "<Undefined>")
                    {
                        rowData[headers[j]] = value;
                    }
                }

                // Store with multiple key variations for better matching
                string primaryKey = CleanElementId(elementId);
                if (!caseSensitiveMatching && !string.IsNullOrEmpty(primaryKey))
                    primaryKey = primaryKey.ToLower();

                if (!string.IsNullOrEmpty(primaryKey))
                {
                    if (data.ContainsKey(primaryKey))
                        duplicateRows++;
                    else
                        data[primaryKey] = rowData;
                }

                // Also add with alternative keys for better matching
                // Add the raw element ID as well
                string rawKey = caseSensitiveMatching ? elementId : elementId.ToLower();
                if (!data.ContainsKey(rawKey))
                    data[rawKey] = rowData;

                // Add with Name column if different from Element ID
                if (nameColumnIndex >= 0 && nameColumnIndex < cells.Length && nameColumnIndex != elementIdIndex)
                {
                    string nameValue = cells[nameColumnIndex].Trim();
                    if (!string.IsNullOrEmpty(nameValue) && nameValue != "<Undefined>")
                    {
                        string nameKey = caseSensitiveMatching ? nameValue : nameValue.ToLower();
                        if (!data.ContainsKey(nameKey))
                            data[nameKey] = rowData;
                    }
                }

                // Add with GlobalId if available
                if (globalIdIndex >= 0 && globalIdIndex < cells.Length)
                {
                    string globalIdValue = cells[globalIdIndex].Trim();
                    if (!string.IsNullOrEmpty(globalIdValue) && globalIdValue != "<Undefined>")
                    {
                        string globalIdKey = caseSensitiveMatching ? globalIdValue : globalIdValue.ToLower();
                        if (!data.ContainsKey(globalIdKey))
                            data[globalIdKey] = rowData;
                    }
                }
            }

            Log($"Successfully loaded {totalRows} element records from CSV (stored as {data.Count} lookup keys)");
            if (duplicateRows > 0)
                Log($"  Note: {duplicateRows} duplicate element IDs found (first occurrence used)");

            return data;
        }
        catch (Exception e)
        {
            Log($"Error reading CSV file: {e.Message}", true);
            return null;
        }
    }

    private string[] ParseCSVLine(string line)
    {
        var result = new List<string>();
        bool inQuotes = false;
        var currentField = new System.Text.StringBuilder();

        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];

            if (c == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    // Handle escaped quotes
                    currentField.Append('"');
                    i++; // Skip next quote
                }
                else
                {
                    inQuotes = !inQuotes;
                }
            }
            else if (c == ',' && !inQuotes)
            {
                result.Add(currentField.ToString());
                currentField.Clear();
            }
            else if (c != '\r') // Skip carriage returns
            {
                currentField.Append(c);
            }
        }

        result.Add(currentField.ToString());
        return result.ToArray();
    }

    #region IFC Schema Hierarchy Methods

    /// <summary>
    /// Loads the IFC schema CSV file and builds the in-memory tree structure.
    /// Expected columns: Package, Group, SubGroup, Name, Kind
    /// </summary>
    private bool LoadSchemaCSV(string filePath)
    {
        if (!File.Exists(filePath))
        {
            Log($"Schema CSV file not found: {filePath}", true);
            return false;
        }

        try
        {
            schemaRoot = new SchemaNode("Root");
            entityNodeLookup = new Dictionary<string, SchemaNode>(StringComparer.OrdinalIgnoreCase);
            psetNodeLookup = new Dictionary<string, SchemaNode>(StringComparer.OrdinalIgnoreCase);
            qtoNodeLookup = new Dictionary<string, SchemaNode>(StringComparer.OrdinalIgnoreCase);

            string[] lines = File.ReadAllLines(filePath);

            if (lines.Length < 2)
            {
                Log("Schema CSV file must have headers and at least one data row", true);
                return false;
            }

            // Parse headers
            var headers = ParseCSVLine(lines[0]);
            for (int i = 0; i < headers.Length; i++)
                headers[i] = headers[i].Trim().TrimStart('\uFEFF');

            // Find column indices
            int packageIdx = Array.FindIndex(headers, h => h.Equals("Package", StringComparison.OrdinalIgnoreCase));
            int groupIdx = Array.FindIndex(headers, h => h.Equals("Group", StringComparison.OrdinalIgnoreCase));
            int subGroupIdx = Array.FindIndex(headers, h => h.Equals("SubGroup", StringComparison.OrdinalIgnoreCase));
            int nameIdx = Array.FindIndex(headers, h => h.Equals("Name", StringComparison.OrdinalIgnoreCase));
            int kindIdx = Array.FindIndex(headers, h => h.Equals("Kind", StringComparison.OrdinalIgnoreCase));

            if (packageIdx == -1 || groupIdx == -1 || nameIdx == -1 || kindIdx == -1)
            {
                Log($"Schema CSV missing required columns. Found: {string.Join(", ", headers)}", true);
                Log("Required columns: Package, Group, SubGroup (optional), Name, Kind", true);
                return false;
            }

            int schemaNodeCount = 0;

            // Process data rows
            for (int i = 1; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i]))
                    continue;

                var cells = ParseCSVLine(lines[i]);

                string package = packageIdx < cells.Length ? cells[packageIdx].Trim() : "";
                string group = groupIdx < cells.Length ? cells[groupIdx].Trim() : "";
                string subGroup = subGroupIdx >= 0 && subGroupIdx < cells.Length ? cells[subGroupIdx].Trim() : "";
                string name = nameIdx < cells.Length ? cells[nameIdx].Trim() : "";
                string kind = kindIdx < cells.Length ? cells[kindIdx].Trim() : "";

                if (string.IsNullOrEmpty(package) || string.IsNullOrEmpty(name))
                    continue;

                // Build tree: Root -> Package -> Group -> [SubGroup] -> Name
                SchemaNode packageNode = schemaRoot.GetOrCreateChild(package, "Package");
                SchemaNode groupNode = packageNode.GetOrCreateChild(group, "Group");

                SchemaNode parentForName = groupNode;
                if (!string.IsNullOrEmpty(subGroup))
                {
                    parentForName = groupNode.GetOrCreateChild(subGroup, "SubGroup");
                }

                SchemaNode nameNode = parentForName.GetOrCreateChild(name, kind, package, group, subGroup);
                schemaNodeCount++;

                // Add to lookup dictionaries based on Kind
                switch (kind.ToLower())
                {
                    case "entity":
                    case "typeenum":
                        entityNodeLookup[name] = nameNode;
                        break;
                    case "pset":
                        psetNodeLookup[name] = nameNode;
                        break;
                    case "qto":
                        qtoNodeLookup[name] = nameNode;
                        break;
                }
            }

            Log($"Loaded IFC schema with {schemaNodeCount} nodes");
            Log($"  Entities/Types: {entityNodeLookup.Count}");
            Log($"  Property Sets: {psetNodeLookup.Count}");
            Log($"  Quantity Sets: {qtoNodeLookup.Count}");

            return true;
        }
        catch (Exception e)
        {
            Log($"Error reading schema CSV file: {e.Message}", true);
            return false;
        }
    }

    /// <summary>
    /// Creates the IFC schema hierarchy GameObjects in the scene based on the loaded schema tree.
    /// </summary>
    private void BuildSchemaHierarchyInScene()
    {
        if (schemaRoot == null || schemaRoot.Children.Count == 0)
        {
            Log("No schema data loaded. Cannot build schema hierarchy.", true);
            return;
        }

        // Find or create the root object
        ifcSchemaRootObject = GameObject.Find(ifcSchemaRootName);
        if (ifcSchemaRootObject == null)
        {
            ifcSchemaRootObject = new GameObject(ifcSchemaRootName);
            Log($"Created IFC Schema root: {ifcSchemaRootName}");
        }
        else
        {
            Log($"Using existing IFC Schema root: {ifcSchemaRootName}");
        }

        schemaRoot.GameObject = ifcSchemaRootObject;

        // Recursively build the hierarchy
        int createdCount = 0;
        BuildSchemaNodeHierarchy(schemaRoot, ifcSchemaRootObject.transform, ref createdCount);

        Log($"Built IFC schema hierarchy with {createdCount} GameObjects");
    }

    /// <summary>
    /// Recursively creates GameObjects for each schema node.
    /// </summary>
    private void BuildSchemaNodeHierarchy(SchemaNode node, Transform parentTransform, ref int createdCount)
    {
        foreach (var child in node.Children.OrderBy(c => c.Name))
        {
            // Find or create the GameObject for this child
            Transform childTransform = parentTransform.Find(child.Name);
            if (childTransform == null)
            {
                GameObject childObj = new GameObject(child.Name);
                childObj.transform.SetParent(parentTransform, false);
                childTransform = childObj.transform;
                createdCount++;
            }

            child.GameObject = childTransform.gameObject;

            // Recurse for children
            BuildSchemaNodeHierarchy(child, childTransform, ref createdCount);
        }
    }

    /// <summary>
    /// Parents an element under its corresponding IFC entity node in the schema hierarchy.
    /// </summary>
    private void ParentElementUnderSchemaHierarchy(Transform element, Dictionary<string, string> metadata)
    {
        if (!enableIfcSchemaHierarchy || schemaRoot == null)
            return;

        // Get the IFC entity type from metadata (try multiple column names)
        string ifcClass = GetMetadataValue(metadata, "IFC_EntityType", "");
        if (string.IsNullOrEmpty(ifcClass))
            ifcClass = GetMetadataValue(metadata, "IfcClass", "");

        if (string.IsNullOrEmpty(ifcClass))
        {
            if (debugMode)
                Log($"  No IFC_EntityType/IfcClass found for element: {element.name}");
            return;
        }

        // Try to find the entity node using storey/entityType lookup first
        string storey = GetMetadataValue(metadata, "IFC_Storey", "");
        string lookupKey = !string.IsNullOrEmpty(storey) ? $"{storey}/{ifcClass}" : ifcClass;

        SchemaNode entityNode = null;

        // First try the storey-specific lookup
        if (!string.IsNullOrEmpty(storey) && entityNodeLookup.TryGetValue($"{storey}/{ifcClass}", out entityNode))
        {
            // Found storey-specific node
        }
        // Fall back to flat entity type lookup
        else if (entityNodeLookup.TryGetValue(ifcClass, out entityNode))
        {
            // Found flat entity node
        }

        if (entityNode != null && entityNode.GameObject != null)
        {
            // Create a linked instance (the element stays in spatial hierarchy too)
            GameObject schemaLinkedElement = new GameObject($"{element.name}_SchemaRef");
            schemaLinkedElement.transform.SetParent(entityNode.GameObject.transform, false);

            // Copy transform from original element
            schemaLinkedElement.transform.position = element.position;
            schemaLinkedElement.transform.rotation = element.rotation;
            schemaLinkedElement.transform.localScale = element.localScale;

            // Add a reference component to link back to the original
            var linkRef = schemaLinkedElement.AddComponent<SchemaElementReference>();
            linkRef.OriginalElement = element.gameObject;
            linkRef.IfcClass = ifcClass;

            schemaReferencesCreated++; // Track schema references created

            if (debugMode)
                Log($"  Linked {element.name} under schema: {lookupKey}");
        }
        else if (debugMode)
        {
            Log($"  IFC_EntityType '{ifcClass}' not found in schema for element: {element.name}");
        }
    }

    /// <summary>
    /// Attaches Pset and Qto property data to an element based on its metadata fields.
    /// </summary>
    private void AttachPropertyDataToElement(Transform element, Dictionary<string, string> metadata)
    {
        if (!createSchemaPropertyLinks)
            return;

        IfcPropertyData propData = null;

        // Process all metadata fields
        foreach (var kvp in metadata)
        {
            string fieldName = kvp.Key;
            string fieldValue = kvp.Value;

            if (string.IsNullOrEmpty(fieldValue))
                continue;

            // Check for Pset_ prefix (handles both "Pset_Common.IsExternal" and "Pset_BeamCommon")
            if (fieldName.StartsWith("Pset_", StringComparison.OrdinalIgnoreCase))
            {
                if (propData == null)
                    propData = GetOrAddPropertyData(element, metadata);

                // Extract the Pset name and property name
                // Format could be: Pset_Common.LoadBearing or just Pset_BeamCommon
                int dotIndex = fieldName.IndexOf('.');
                string psetName, propName;

                if (dotIndex > 0)
                {
                    psetName = fieldName.Substring(0, dotIndex);
                    propName = fieldName.Substring(dotIndex + 1);
                }
                else
                {
                    psetName = fieldName;
                    propName = "Value";
                }

                propData.AddProperty(psetName, propName, fieldValue);
            }
            // Check for Qto_ prefix (handles both "Qto_BaseQuantities.Length" and "Qto_BeamBaseQuantities")
            else if (fieldName.StartsWith("Qto_", StringComparison.OrdinalIgnoreCase))
            {
                if (propData == null)
                    propData = GetOrAddPropertyData(element, metadata);

                // Extract the Qto name and quantity name
                int dotIndex = fieldName.IndexOf('.');
                string qtoName, qtyName;

                if (dotIndex > 0)
                {
                    qtoName = fieldName.Substring(0, dotIndex);
                    qtyName = fieldName.Substring(dotIndex + 1);
                }
                else
                {
                    qtoName = fieldName;
                    qtyName = "Value";
                }

                propData.AddQuantity(qtoName, qtyName, fieldValue);
            }
        }

        // Also attach the IFC_PropertySets and IFC_QuantitySets info if present
        if (propData != null)
        {
            string psets = GetMetadataValue(metadata, "IFC_PropertySets", "");
            string qtos = GetMetadataValue(metadata, "IFC_QuantitySets", "");

            // Store as metadata on the property data component (for reference)
            if (!string.IsNullOrEmpty(psets) && propData.PropertySets.Count == 0)
            {
                foreach (var pset in psets.Split(','))
                {
                    var trimmed = pset.Trim();
                    if (!string.IsNullOrEmpty(trimmed))
                    {
                        propData.PropertySets.Add(new IfcPropertyData.PropertySetReference
                        {
                            PropertySetName = trimmed
                        });
                    }
                }
            }

            if (!string.IsNullOrEmpty(qtos) && propData.QuantitySets.Count == 0)
            {
                foreach (var qto in qtos.Split(','))
                {
                    var trimmed = qto.Trim();
                    if (!string.IsNullOrEmpty(trimmed))
                    {
                        propData.QuantitySets.Add(new IfcPropertyData.QuantitySetReference
                        {
                            QuantitySetName = trimmed
                        });
                    }
                }
            }
        }
    }

    private IfcPropertyData GetOrAddPropertyData(Transform element, Dictionary<string, string> metadata)
    {
        var propData = element.GetComponent<IfcPropertyData>();
        if (propData == null)
        {
            propData = element.gameObject.AddComponent<IfcPropertyData>();
            propData.GlobalId = GetMetadataValue(metadata, "GlobalId", "");
            propData.IfcClass = GetMetadataValue(metadata, "IfcClass", "");
        }
        return propData;
    }

    /// <summary>
    /// Derives IFC entity type from element type.
    /// Converts element types like "Wall", "Slab", "Beam" to IFC entity types like "IfcWall", "IfcSlab", "IfcBeam".
    /// </summary>
    private string DeriveIfcEntityType(string elementType)
    {
        if (string.IsNullOrEmpty(elementType))
            return "IfcBuildingElementProxy";

        // If already in IFC format, return as is
        if (elementType.StartsWith("Ifc", StringComparison.OrdinalIgnoreCase))
            return elementType;

        // Common element type mappings
        var lowerType = elementType.ToLower().Trim();

        // Direct mappings
        switch (lowerType)
        {
            case "wall":
                return "IfcWall";
            case "slab":
                return "IfcSlab";
            case "beam":
                return "IfcBeam";
            case "column":
                return "IfcColumn";
            case "door":
                return "IfcDoor";
            case "window":
                return "IfcWindow";
            case "roof":
                return "IfcRoof";
            case "stair":
            case "stairs":
                return "IfcStair";
            case "railing":
                return "IfcRailing";
            case "curtain wall":
            case "curtainwall":
                return "IfcCurtainWall";
            case "covering":
                return "IfcCovering";
            case "ceiling":
                return "IfcCovering";
            case "floor":
                return "IfcSlab";
            case "furniture":
                return "IfcFurniture";
            case "space":
            case "room":
            case "zone":
                return "IfcSpace";
            case "opening":
                return "IfcOpeningElement";
            case "footing":
                return "IfcFooting";
            case "pile":
                return "IfcPile";
            case "ramp":
                return "IfcRamp";
            case "plate":
                return "IfcPlate";
            case "member":
                return "IfcMember";
            case "pipe":
                return "IfcPipeSegment";
            case "duct":
                return "IfcDuctSegment";
            case "equipment":
                return "IfcBuildingElementProxy";
            default:
                // For unknown types, convert to IfcBuildingElementProxy
                // but try to preserve the original name
                return $"Ifc{char.ToUpper(elementType[0])}{elementType.Substring(1)}";
        }
    }

    /// <summary>
    /// Builds the IFC schema hierarchy from the element metadata CSV.
    /// Uses Element_Type to derive IFC entity types, and IFC4RV/IFC columns for spatial hierarchy.
    /// Also uses Top_Link_Story or similar columns to determine storey/level.
    /// </summary>
    private bool BuildSchemaFromElementMetadata()
    {
        if (elementMetadata == null || elementMetadata.Count == 0)
        {
            Log("No element metadata loaded. Cannot build schema from element CSV.", true);
            return false;
        }

        try
        {
            schemaRoot = new SchemaNode("Root");
            entityNodeLookup = new Dictionary<string, SchemaNode>(StringComparer.OrdinalIgnoreCase);
            psetNodeLookup = new Dictionary<string, SchemaNode>(StringComparer.OrdinalIgnoreCase);
            qtoNodeLookup = new Dictionary<string, SchemaNode>(StringComparer.OrdinalIgnoreCase);

            // Collect unique values from metadata
            var projects = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var sites = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var buildings = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var storeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var entityTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var propertySets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var quantitySets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Also collect hierarchical relationships
            var spatialStructure = new Dictionary<string, Dictionary<string, Dictionary<string, HashSet<string>>>>();
            // Project -> Site -> Building -> Storey

            foreach (var element in elementMetadata.Values)
            {
                // Use standard column "Design_Options.Design_Option_Name" or default for project
                string project = GetMetadataValue(element, "Design_Options.Design_Option_Name", "");
                if (string.IsNullOrEmpty(project))
                    project = GetMetadataValue(element, "Design_Options.Design_Option_Set_Name", "");
                if (string.IsNullOrEmpty(project))
                    project = "Default Project";

                // Site - use a default or derive from location if available
                string site = "Default Site";

                // Building - use Design Option Set Name or default
                string building = GetMetadataValue(element, "Design_Options.Design_Option_Set_Name", "");
                if (string.IsNullOrEmpty(building))
                    building = "Default Building";

                // Use Top_Link_Story for storey/level information
                string storey = GetMetadataValue(element, "General_Parameters.Top_Link_Story", "");
                if (string.IsNullOrEmpty(storey))
                    storey = GetMetadataValue(element, "Positioning.Top_Link_Story", "");
                if (string.IsNullOrEmpty(storey))
                    storey = GetMetadataValue(element, "Top_Link_Story", "");
                if (string.IsNullOrEmpty(storey))
                    storey = GetMetadataValue(element, "Level", "");
                if (string.IsNullOrEmpty(storey))
                    storey = "Default Storey";

                // Get element type from Element_Type column
                string elementType = GetMetadataValue(element, "Element_Type", "");
                if (string.IsNullOrEmpty(elementType))
                    elementType = GetMetadataValue(element, "Category", "");
                if (string.IsNullOrEmpty(elementType))
                    elementType = GetMetadataValue(element, "General_Parameters.Element_Type", "");

                // Convert element type to IFC entity type (e.g., "Wall" -> "IfcWall")
                string entityType = DeriveIfcEntityType(elementType);

                // Property sets and quantity sets - left empty as we're not using IFC-specific columns
                string psets = "";
                string qtos = "";

                if (!string.IsNullOrEmpty(project)) projects.Add(project);
                if (!string.IsNullOrEmpty(site)) sites.Add(site);
                if (!string.IsNullOrEmpty(building)) buildings.Add(building);
                if (!string.IsNullOrEmpty(storey)) storeys.Add(storey);
                if (!string.IsNullOrEmpty(entityType)) entityTypes.Add(entityType);

                // Parse property sets (comma-separated)
                if (!string.IsNullOrEmpty(psets))
                {
                    foreach (var pset in psets.Split(','))
                    {
                        var trimmed = pset.Trim();
                        if (!string.IsNullOrEmpty(trimmed))
                            propertySets.Add(trimmed);
                    }
                }

                // Parse quantity sets (comma-separated)
                if (!string.IsNullOrEmpty(qtos))
                {
                    foreach (var qto in qtos.Split(','))
                    {
                        var trimmed = qto.Trim();
                        if (!string.IsNullOrEmpty(trimmed))
                            quantitySets.Add(trimmed);
                    }
                }

                // Build spatial structure
                if (!string.IsNullOrEmpty(project))
                {
                    if (!spatialStructure.ContainsKey(project))
                        spatialStructure[project] = new Dictionary<string, Dictionary<string, HashSet<string>>>(StringComparer.OrdinalIgnoreCase);

                    if (!string.IsNullOrEmpty(site))
                    {
                        if (!spatialStructure[project].ContainsKey(site))
                            spatialStructure[project][site] = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

                        if (!string.IsNullOrEmpty(building))
                        {
                            if (!spatialStructure[project][site].ContainsKey(building))
                                spatialStructure[project][site][building] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                            if (!string.IsNullOrEmpty(storey))
                                spatialStructure[project][site][building].Add(storey);
                        }
                    }
                }
            }

            // Build the schema tree
            // Structure: Root -> Project -> Site -> Building -> Storey -> EntityTypes

            // Create spatial hierarchy nodes
            foreach (var projectKvp in spatialStructure)
            {
                var projectNode = schemaRoot.GetOrCreateChild(projectKvp.Key, "IfcProject");

                foreach (var siteKvp in projectKvp.Value)
                {
                    var siteNode = projectNode.GetOrCreateChild(siteKvp.Key, "IfcSite");

                    foreach (var buildingKvp in siteKvp.Value)
                    {
                        var buildingNode = siteNode.GetOrCreateChild(buildingKvp.Key, "IfcBuilding");

                        foreach (var storey in buildingKvp.Value)
                        {
                            var storeyNode = buildingNode.GetOrCreateChild(storey, "IfcBuildingStorey");

                            // Add entity type nodes under each storey
                            foreach (var entityType in entityTypes)
                            {
                                var entityNode = storeyNode.GetOrCreateChild(entityType, "Entity");

                                // Store in lookup for quick access
                                string lookupKey = $"{storey}/{entityType}";
                                if (!entityNodeLookup.ContainsKey(lookupKey))
                                    entityNodeLookup[lookupKey] = entityNode;
                            }
                        }
                    }
                }
            }

            // Also add a flat entity type lookup (for elements without full spatial data)
            foreach (var entityType in entityTypes)
            {
                if (!entityNodeLookup.ContainsKey(entityType))
                {
                    // Create under a default "Entities" node if not already categorized
                    var entitiesNode = schemaRoot.GetOrCreateChild("Entities", "Group");
                    var entityNode = entitiesNode.GetOrCreateChild(entityType, "Entity");
                    entityNodeLookup[entityType] = entityNode;
                }
            }

            // Create PropertySets node
            if (propertySets.Count > 0)
            {
                var psetsNode = schemaRoot.GetOrCreateChild("PropertySets", "Group");
                foreach (var pset in propertySets)
                {
                    var psetNode = psetsNode.GetOrCreateChild(pset, "Pset");
                    psetNodeLookup[pset] = psetNode;
                }
            }

            // Create QuantitySets node
            if (quantitySets.Count > 0)
            {
                var qtosNode = schemaRoot.GetOrCreateChild("QuantitySets", "Group");
                foreach (var qto in quantitySets)
                {
                    var qtoNode = qtosNode.GetOrCreateChild(qto, "Qto");
                    qtoNodeLookup[qto] = qtoNode;
                }
            }

            Log($"Built IFC schema from element metadata:");
            Log($"  Projects: {projects.Count}");
            Log($"  Sites: {sites.Count}");
            Log($"  Buildings: {buildings.Count}");
            Log($"  Storeys: {storeys.Count}");
            Log($"  Entity Types: {entityTypes.Count}");
            Log($"  Property Sets: {propertySets.Count}");
            Log($"  Quantity Sets: {quantitySets.Count}");

            return true;
        }
        catch (Exception e)
        {
            Log($"Error building schema from element metadata: {e.Message}", true);
            return false;
        }
    }

    #endregion

    private void ResetHierarchy()
    {
        if (modelRoot == null) return;

        bool confirm = EditorUtility.DisplayDialog(
            "Reset Hierarchy",
            "This will flatten all elements directly under the root object. Continue?",
            "Yes", "Cancel");

        if (!confirm) return;

        Log("Resetting hierarchy to flat structure...");

        var allElements = modelRoot.GetComponentsInChildren<Transform>(true)
                                 .Where(t => t != modelRoot.transform && t.childCount == 0)
                                 .ToList();

        foreach (var element in allElements)
        {
            element.SetParent(modelRoot.transform, true);
        }

        // Remove empty intermediate hierarchy objects
        var allChildren = modelRoot.GetComponentsInChildren<Transform>(true)
                                 .Where(t => t != modelRoot.transform && t.childCount == 0)
                                 .Select(t => t.parent)
                                 .Distinct()
                                 .Where(p => p != modelRoot.transform)
                                 .ToList();

        foreach (var emptyParent in allChildren)
        {
            if (emptyParent.childCount == 0)
            {
                DestroyImmediate(emptyParent.gameObject);
            }
        }

        EditorUtility.SetDirty(modelRoot);
        Log("Hierarchy reset to flat structure completed");
    }

    private void ExportHierarchyToCSV()
    {
        if (modelRoot == null) return;

        string path = EditorUtility.SaveFilePanel("Export Hierarchy", "", $"hierarchy_export_{bimType}.csv", "csv");
        if (string.IsNullOrEmpty(path)) return;

        try
        {
            var lines = new List<string>();
            lines.Add("Element_Name,Full_Path,Level,Parent,BIM_Type");

            var allElements = modelRoot.GetComponentsInChildren<Transform>(true)
                                     .Where(t => t != modelRoot.transform);

            foreach (var element in allElements)
            {
                string fullPath = GetFullPath(element);
                int level = GetHierarchyLevel(element);
                string parent = element.parent != null ? element.parent.name : "";

                lines.Add($"\"{element.name}\",\"{fullPath}\",{level},\"{parent}\",\"{bimType}\"");
            }

            File.WriteAllLines(path, lines);
            Log($"Hierarchy exported to: {path}");
        }
        catch (Exception e)
        {
            Log($"Failed to export hierarchy: {e.Message}", true);
        }
    }

    private string GetFullPath(Transform transform)
    {
        var path = new List<string>();
        Transform current = transform;

        while (current != null && current != modelRoot.transform)
        {
            path.Insert(0, current.name);
            current = current.parent;
        }

        return string.Join("/", path);
    }

    private int GetHierarchyLevel(Transform transform)
    {
        int level = 0;
        Transform current = transform;

        while (current != null && current != modelRoot.transform)
        {
            level++;
            current = current.parent;
        }

        return level;
    }

    private string GetHierarchyDescription(HierarchyMode mode)
    {
        // Provide context-aware descriptions based on BIM type
        if (bimType == BIMType.Revit)
        {
            return mode switch
            {
                HierarchyMode.ByCategory => "NonModular: Category → Family → Type → Elements",
                HierarchyMode.FlatByCategory => "Flat By Category: Category → Elements (flat structure)",
                HierarchyMode.ByMaterial => "Material Based: Material → Category → Family → Type → Elements",
                HierarchyMode.ByLevel => "Level Based: Level → Category → Family → Type → Elements",
                HierarchyMode.BySystem => "System Based: System Name → System Type → Category → Family → Elements",
                HierarchyMode.Modular => "Modular: Modular Group → Category → Family → Type → Elements",
                HierarchyMode.ModularDetailed => "Modular Detailed: Modular Group → Level → System Name → Category → Family → Type → Elements",
                HierarchyMode.IfcSchema => "IFC Schema: IfcProject → IfcSite → IfcBuilding → IfcStorey → EntityType → Elements",
                HierarchyMode.Custom => "Custom: User-defined comma-separated columns",
                _ => "Unknown hierarchy mode"
            };
        }
        else // ArchiCAD
        {
            return mode switch
            {
                HierarchyMode.ByCategory => "Category → Elements",
                HierarchyMode.ByFamily => "Family → Type → Elements",
                HierarchyMode.ByMaterial => "Material → Category → Elements",
                HierarchyMode.ByLevel => "Level → Category → Family → Elements",
                HierarchyMode.BySystem => "System → Category → Family → Elements",
                HierarchyMode.SystemDetailed => "System Name → System Type → Category → Family → Elements",
                HierarchyMode.MaterialDetailed => "Material → Category → Family → Type → Elements",
                HierarchyMode.LevelDetailed => "Level → Category → Family → Type → Elements",
                HierarchyMode.FlatByCategory => "Flat By Category: Category → Elements (flat structure)",
                HierarchyMode.ConstructionDiscipline => "Construction Discipline (Layer-based) → Element Type → Elements",
                HierarchyMode.SpatialDiscipline => "Spatial Location (Zone/Room) → Element Type → Elements",
                HierarchyMode.MultiLevelClassification => "Element Type → Profile/Library Category → Elements",
                HierarchyMode.IfcSchema => "IFC Schema: IfcProject → IfcSite → IfcBuilding → IfcStorey → EntityType → Elements",
                HierarchyMode.Custom => "User-defined hierarchy levels",
                _ => "Unknown hierarchy mode"
            };
        }
    }

    private void LogOrganizationResults(int totalElements, int matched, int unmatched, Dictionary<string, int> stats)
    {
        // Calculate final memory usage
        long memoryAtEnd = GC.GetTotalMemory(false);
        long memoryUsed = memoryAtEnd - memoryAtStart;
        UpdatePeakMemory();

        Log($"\n<b>╔══════════════════════════════════════════════════════════════╗</b>");
        Log($"<b>║           HIERARCHY ORGANIZATION SUMMARY                     ║</b>");
        Log($"<b>╚══════════════════════════════════════════════════════════════╝</b>");

        // Processing Statistics
        Log($"\n<b>📊 PROCESSING STATISTICS:</b>");
        Log($"  ├─ Total Elements Processed: {totalElements:N0}");
        Log($"  ├─ Successfully Matched: {matched:N0} ({(totalElements > 0 ? (float)matched / totalElements * 100 : 0):F2}%)");
        Log($"  ├─ Unmatched Elements: {unmatched:N0} ({(totalElements > 0 ? (float)unmatched / totalElements * 100 : 0):F2}%)");
        Log($"  ├─ Hierarchy Categories Created: {stats.Count:N0}");
        Log($"  ├─ Hierarchy Nodes Created: {hierarchyNodesCreated:N0}");
        Log($"  └─ CSV Records Loaded: {csvRowsLoaded:N0}");

        // Data Loss Analysis
        Log($"\n<b>⚠️ DATA LOSS ANALYSIS:</b>");
        int dataLoss = unmatched;
        float dataLossPercentage = totalElements > 0 ? (float)unmatched / totalElements * 100 : 0;
        string dataLossStatus = dataLossPercentage == 0 ? "<color=green>✓ NO DATA LOSS</color>" :
                                dataLossPercentage < 5 ? "<color=yellow>⚠ MINIMAL DATA LOSS</color>" :
                                dataLossPercentage < 20 ? "<color=orange>⚠ MODERATE DATA LOSS</color>" :
                                "<color=red>✗ SIGNIFICANT DATA LOSS</color>";
        Log($"  ├─ Status: {dataLossStatus}");
        Log($"  ├─ Elements without metadata match: {dataLoss:N0}");
        Log($"  ├─ Data Loss Percentage: {dataLossPercentage:F2}%");
        Log($"  └─ Successfully Categorized: {(100 - dataLossPercentage):F2}%");

        // Performance Metrics
        Log($"\n<b>⏱️ PERFORMANCE METRICS:</b>");
        Log($"  ├─ Total Processing Time: {performanceTimer.Elapsed.TotalSeconds:F2} seconds ({performanceTimer.ElapsedMilliseconds:N0} ms)");
        Log($"  ├─ Processing Speed: {(totalElements > 0 && performanceTimer.Elapsed.TotalSeconds > 0 ? totalElements / performanceTimer.Elapsed.TotalSeconds : 0):F1} elements/second");
        Log($"  └─ End Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");

        // Memory Usage
        Log($"\n<b>💾 MEMORY USAGE:</b>");
        Log($"  ├─ Memory at Start: {FormatMemorySize(memoryAtStart)}");
        Log($"  ├─ Memory at End: {FormatMemorySize(memoryAtEnd)}");
        Log($"  ├─ Peak Memory Used: {FormatMemorySize(peakMemoryUsed)}");
        Log($"  ├─ Net Memory Change: {(memoryUsed >= 0 ? "+" : "")}{FormatMemorySize(Math.Abs(memoryUsed))}");
        Log($"  └─ Memory per Element: {(totalElements > 0 ? FormatMemorySize(Math.Abs(memoryUsed) / totalElements) : "N/A")}");

        // Hierarchy Structure
        Log($"\n<b>🏗️ HIERARCHY STRUCTURE:</b>");
        Log($"  ├─ Organization Mode: {CurrentHierarchyMode}");
        Log($"  ├─ BIM Platform: {bimType}");
        Log($"  └─ Hierarchy Description: {GetHierarchyDescription(CurrentHierarchyMode)}");

        if (stats.Count > 0)
        {
            Log($"\n<b>📁 TOP ORGANIZATION CATEGORIES:</b>");
            int categoryRank = 1;
            foreach (var category in stats.OrderByDescending(s => s.Value).Take(15))
            {
                float categoryPercent = totalElements > 0 ? (float)category.Value / matched * 100 : 0;
                Log($"  {categoryRank,2}. {category.Key}: {category.Value:N0} elements ({categoryPercent:F1}%)");
                categoryRank++;
            }

            if (stats.Count > 15)
                Log($"  ... and {stats.Count - 15} more categories");
        }

        // IFC Schema hierarchy information
        if (enableIfcSchemaHierarchy && schemaRoot != null)
        {
            Log($"\n<b>📐 IFC SCHEMA HIERARCHY:</b>");
            Log($"  ├─ Schema Root: {ifcSchemaRootName}");
            Log($"  ├─ Entity Types in Schema: {entityNodeLookup?.Count ?? 0:N0}");
            Log($"  ├─ Property Sets in Schema: {psetNodeLookup?.Count ?? 0:N0}");
            Log($"  ├─ Quantity Sets in Schema: {qtoNodeLookup?.Count ?? 0:N0}");
            Log($"  └─ Schema References Created: {schemaReferencesCreated:N0}");
        }

        // Warnings and Recommendations
        if (unmatched > 0)
        {
            Log($"\n<b>💡 RECOMMENDATIONS:</b>");
            Log($"<color=yellow>  ⚠ {unmatched:N0} elements remain unmatched.</color>");
            Log($"  Possible causes:");
            Log($"    • Element ID mismatch between CSV and GameObject names");
            Log($"    • Missing metadata rows in CSV file");
            Log($"    • Case sensitivity issues (current: {(caseSensitiveMatching ? "Case-Sensitive" : "Case-Insensitive")})");
            Log($"  Suggested actions:");
            Log($"    • Enable Debug Mode to see individual matching results");
            Log($"    • Check the 'Unmatched_Elements' group in the hierarchy");
            Log($"    • Verify CSV Element ID column matches GameObject naming convention");
        }
        else
        {
            Log($"\n<color=green><b>✓ ALL ELEMENTS SUCCESSFULLY CATEGORIZED!</b></color>");
        }

        Log($"\n<b>════════════════════════════════════════════════════════════════</b>");

        // Save log to file
        SaveLogToFile("HierarchyOrganization");
    }

    /// <summary>
    /// Save log and metrics to a text file
    /// </summary>
    private void SaveLogToFile(string processName)
    {
        try
        {
            // Create logs directory in project
            string logsFolder = Path.Combine(Application.dataPath, "BIMUniXchange", "Logs");
            if (!Directory.Exists(logsFolder))
            {
                Directory.CreateDirectory(logsFolder);
            }

            // Generate filename with timestamp
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string modelName = modelRoot != null ? modelRoot.name : "Unknown";
            string fileName = $"{processName}_{modelName}_{timestamp}.txt";
            string logFilePath = Path.Combine(logsFolder, fileName);

            // Build full log content
            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"Process: {processName}");
            sb.AppendLine($"Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"BIM Type: {bimType}");
            sb.AppendLine($"Hierarchy Mode: {CurrentHierarchyMode}");
            sb.AppendLine($"Model Root: {(modelRoot != null ? modelRoot.name : "N/A")}");
            sb.AppendLine($"CSV File: {(csvFile != null ? AssetDatabase.GetAssetPath(csvFile) : "N/A")}");
            sb.AppendLine();

            // Strip HTML tags for file output
            string cleanLog = System.Text.RegularExpressions.Regex.Replace(logText, "<[^>]+>", "");
            sb.AppendLine(cleanLog);

            File.WriteAllText(logFilePath, sb.ToString());
            Log($"Log saved to: {logFilePath}");

            // Refresh AssetDatabase to show new file
            AssetDatabase.Refresh();
        }
        catch (Exception ex)
        {
            Debug.LogError($"[UnifiedHierarchyOrganizer] Failed to save log file: {ex.Message}");
        }
    }

    private void Log(string message, bool isError = false)
    {
        string timestamp = DateTime.Now.ToString("HH:mm:ss");
        string logEntry = $"[{timestamp}] {message}\n";

        if (isError)
        {
            logEntry = $"<color=red>{logEntry}</color>";
            Debug.LogError($"[UnifiedHierarchyOrganizer] {message}");
        }
        else if (debugMode)
        {
            Debug.Log($"[UnifiedHierarchyOrganizer] {message}");
        }

        logText += logEntry;
        logScroll.y = Mathf.Infinity;
        Repaint();
    }
}
