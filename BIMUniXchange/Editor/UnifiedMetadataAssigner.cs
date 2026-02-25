using UnityEngine;
using UnityEditor;
using UnityEditor.SceneManagement;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System;
using System.Diagnostics;

public class UnifiedMetadataAssigner : EditorWindow
{
    private enum BIMType
    {
        Archicad,
        Revit
    }

    private BIMType currentTab = BIMType.Archicad;

    // Common fields
    private GameObject targetObject;
    private UnityEngine.Object csvFile;
    private Vector2 logScroll;
    private List<string> logs = new List<string>();

    // Statistics
    private readonly Dictionary<string, MetadataStats> elementStats = new Dictionary<string, MetadataStats>();

    // Performance tracking
    private Stopwatch performanceTimer;
    private long memoryAtStart;
    private long peakMemoryUsed;
    private int csvRowsLoaded;
    private int totalPropertiesAssigned;
    private int totalEmptyProperties;
    private int elementsWithMetadata;
    private int elementsWithoutMatch;

    // Log file path
    private string _lastLogFilePath;

    [MenuItem("Window/BIMUniXchange/Metadata Assignment", false, 30)]
    public static void ShowWindow()
    {
        var window = GetWindow<UnifiedMetadataAssigner>("BIM Metadata Assignment");
        window.minSize = new Vector2(500, 400);
    }

    void OnGUI()
    {
        EditorGUILayout.LabelField("BIM Metadata Assignment", EditorStyles.boldLabel);
        EditorGUILayout.Space();

        // Tab selection
        EditorGUILayout.BeginHorizontal();
        if (GUILayout.Toggle(currentTab == BIMType.Archicad, "Archicad (FBX)", EditorStyles.miniButtonLeft))
            currentTab = BIMType.Archicad;
        if (GUILayout.Toggle(currentTab == BIMType.Revit, "Revit (OBJ)", EditorStyles.miniButtonRight))
            currentTab = BIMType.Revit;
        EditorGUILayout.EndHorizontal();

        EditorGUILayout.Space();

        // Common UI
        string objectLabel = currentTab == BIMType.Archicad ? "FBX Root Object" : "OBJ Object";
        targetObject = (GameObject)EditorGUILayout.ObjectField(objectLabel, targetObject, typeof(GameObject), true);
        csvFile = EditorGUILayout.ObjectField("CSV Metadata File", csvFile, typeof(UnityEngine.Object), false);

        EditorGUILayout.Space();

        // Assignment button
        using (new EditorGUI.DisabledScope(targetObject == null || csvFile == null))
        {
            if (GUILayout.Button("Assign Metadata", GUILayout.Height(30)))
            {
                logs.Clear();
                elementStats.Clear();
                AssignMetadata();
            }
        }

        EditorGUILayout.Space();

        // Log display with copy button
        EditorGUILayout.BeginHorizontal();
        EditorGUILayout.LabelField("Assignment Log", EditorStyles.boldLabel);
        GUILayout.FlexibleSpace();
        if (logs.Count > 0 && GUILayout.Button("Copy Log", GUILayout.Width(80)))
        {
            GUIUtility.systemCopyBuffer = string.Join("\n", logs);
            EditorUtility.DisplayDialog("Copied", "Log copied to clipboard!", "OK");
        }
        if (logs.Count > 0 && GUILayout.Button("Clear Log", GUILayout.Width(80)))
        {
            logs.Clear();
        }
        EditorGUILayout.EndHorizontal();

        EditorGUILayout.BeginVertical(EditorStyles.helpBox);
        logScroll = EditorGUILayout.BeginScrollView(logScroll, GUILayout.Height(300));

        var logStyle = new GUIStyle(EditorStyles.textArea)
        {
            wordWrap = false,
            richText = false,
            font = Font.CreateDynamicFontFromOSFont("Consolas", 11)
        };

        string logText = string.Join("\n", logs);
        EditorGUILayout.SelectableLabel(logText, logStyle, GUILayout.ExpandHeight(true), GUILayout.ExpandWidth(true), GUILayout.MinHeight(280));
        EditorGUILayout.EndScrollView();
        EditorGUILayout.EndVertical();
    }

    void Log(string msg)
    {
        logs.Add($"[{DateTime.Now:HH:mm:ss}] {msg}");
        Repaint();
        UnityEngine.Debug.Log(msg);
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

    void AssignMetadata()
    {
        if (targetObject == null || csvFile == null)
        {
            EditorUtility.DisplayDialog("Missing Input", "Please assign both target object and CSV file.", "OK");
            return;
        }

        // Initialize performance tracking
        performanceTimer = Stopwatch.StartNew();
        memoryAtStart = GC.GetTotalMemory(false);
        peakMemoryUsed = memoryAtStart;
        csvRowsLoaded = 0;
        totalPropertiesAssigned = 0;
        totalEmptyProperties = 0;
        elementsWithMetadata = 0;
        elementsWithoutMatch = 0;

        Log($"╔══════════════════════════════════════════════════════════════╗");
        Log($"║        BIM METADATA ASSIGNMENT - {currentTab.ToString().ToUpper()}                    ║");
        Log($"╚══════════════════════════════════════════════════════════════╝");
        Log($"Start Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        Log($"Initial Memory: {FormatMemorySize(memoryAtStart)}");
        Log($"");

        // Parse CSV data
        var csvLoadTimer = Stopwatch.StartNew();
        var csvData = ParseCSV();
        csvLoadTimer.Stop();

        if (csvData == null || csvData.Count == 0)
        {
            Log("Failed to parse CSV data");
            EditorUtility.DisplayDialog("CSV Error", "Could not parse CSV data.", "OK");
            return;
        }

        csvRowsLoaded = csvData.Count;
        Log($"CSV Loading Time: {csvLoadTimer.ElapsedMilliseconds}ms");
        Log($"Parsed {csvData.Count} CSV rows");
        UpdatePeakMemory();

        // Get child objects
        var children = targetObject.GetComponentsInChildren<Transform>(true)
            .Where(t => t.gameObject != targetObject)
            .Select(t => t.gameObject)
            .ToArray();

        int totalChildren = children.Length;
        Log($"Found {totalChildren} child objects to process");

        // Track assignment progress
        var assignmentTimer = Stopwatch.StartNew();
        int assignedCount = 0;
        int progressInterval = Math.Max(1, totalChildren / 100); // Update every 1%

        foreach (var go in children)
        {
            // Update progress bar periodically
            if (assignedCount % progressInterval == 0)
            {
                float progress = (float)assignedCount / totalChildren;
                EditorUtility.DisplayProgressBar("Assigning Metadata",
                    $"Processing {assignedCount + 1} of {totalChildren} ({elementsWithMetadata} matched)", progress);
            }

            string elementId = currentTab == BIMType.Archicad ?
                ExtractArchicadElementId(go.name) :
                ExtractRevitElementId(go.name);

            if (string.IsNullOrEmpty(elementId) || !csvData.ContainsKey(elementId))
            {
                elementsWithoutMatch++;
                continue;
            }

            var csvProperties = csvData[elementId];
            var metadata = go.GetComponent<Metadata>() ?? go.AddComponent<Metadata>();

            // Use the AssignCSVData method which properly sets up the metadata
            metadata.AssignCSVData(csvProperties);

            // Store stats for display
            elementStats[elementId] = new MetadataStats
            {
                TotalParameters = metadata.Stats.TotalParameters,
                EmptyParameters = metadata.Stats.EmptyParameters,
                NonEmptyParameters = metadata.Stats.NonEmptyParameters,
                AssignedParameters = metadata.Stats.AssignedParameters
            };

            // Accumulate totals
            totalPropertiesAssigned += metadata.Stats.NonEmptyParameters;
            totalEmptyProperties += metadata.Stats.EmptyParameters;
            elementsWithMetadata++;

            // Mark the metadata component and the game object as dirty
            EditorUtility.SetDirty(metadata);
            EditorUtility.SetDirty(go);

            assignedCount++;
            UpdatePeakMemory();
        }

        assignmentTimer.Stop();
        EditorUtility.ClearProgressBar();

        Log($"Assignment Processing Time: {assignmentTimer.ElapsedMilliseconds}ms");

        if (assignedCount > 0)
        {
            // Mark the scene as dirty to ensure changes are saved
            EditorSceneManager.MarkSceneDirty(EditorSceneManager.GetActiveScene());

            // Force refresh of the inspector
            UnityEditorInternal.InternalEditorUtility.RepaintAllViews();

            // Save assets and refresh
            AssetDatabase.SaveAssets();
            AssetDatabase.Refresh();
        }

        performanceTimer.Stop();

        // Display comprehensive results
        DisplayPerformanceResults(totalChildren, assignedCount);

        EditorUtility.DisplayDialog("Assignment Complete",
            $"Successfully assigned metadata to {assignedCount}/{totalChildren} objects.\n\n" +
            $"Data Loss: {(totalChildren > 0 ? (float)(totalChildren - assignedCount) / totalChildren * 100 : 0):F2}%\n" +
            $"Time: {performanceTimer.Elapsed.TotalSeconds:F2} seconds\n\n" +
            "Select any assigned object in the hierarchy to view its metadata in the inspector.", "OK");
    }

    private void DisplayPerformanceResults(int totalChildren, int assignedCount)
    {
        // Calculate final memory usage
        long memoryAtEnd = GC.GetTotalMemory(false);
        long memoryUsed = memoryAtEnd - memoryAtStart;
        UpdatePeakMemory();

        int unmatchedCount = totalChildren - assignedCount;
        float dataLossPercentage = totalChildren > 0 ? (float)unmatchedCount / totalChildren * 100 : 0;

        Log($"");
        Log($"╔══════════════════════════════════════════════════════════════╗");
        Log($"║           METADATA ASSIGNMENT SUMMARY                        ║");
        Log($"╚══════════════════════════════════════════════════════════════╝");

        // Processing Statistics
        Log($"");
        Log($"📊 PROCESSING STATISTICS:");
        Log($"  ├─ Total Objects Scanned: {totalChildren:N0}");
        Log($"  ├─ Successfully Matched: {assignedCount:N0} ({(totalChildren > 0 ? (float)assignedCount / totalChildren * 100 : 0):F2}%)");
        Log($"  ├─ Unmatched Objects: {unmatchedCount:N0} ({dataLossPercentage:F2}%)");
        Log($"  ├─ CSV Records Available: {csvRowsLoaded:N0}");
        Log($"  ├─ Total Properties Assigned: {totalPropertiesAssigned:N0}");
        Log($"  ├─ Empty Properties Skipped: {totalEmptyProperties:N0}");
        Log($"  └─ Avg Properties per Element: {(assignedCount > 0 ? (float)totalPropertiesAssigned / assignedCount : 0):F1}");

        // Data Loss Analysis
        Log($"");
        Log($"⚠️ DATA LOSS ANALYSIS:");
        string dataLossStatus = dataLossPercentage == 0 ? "✓ NO DATA LOSS" :
                                dataLossPercentage < 5 ? "⚠ MINIMAL DATA LOSS" :
                                dataLossPercentage < 20 ? "⚠ MODERATE DATA LOSS" :
                                "✗ SIGNIFICANT DATA LOSS";
        Log($"  ├─ Status: {dataLossStatus}");
        Log($"  ├─ Objects without metadata match: {unmatchedCount:N0}");
        Log($"  ├─ Data Loss Percentage: {dataLossPercentage:F2}%");
        Log($"  ├─ Successfully Processed: {(100 - dataLossPercentage):F2}%");
        Log($"  └─ CSV Utilization: {(csvRowsLoaded > 0 ? (float)assignedCount / csvRowsLoaded * 100 : 0):F2}% of CSV records used");

        // Performance Metrics
        Log($"");
        Log($"⏱️ PERFORMANCE METRICS:");
        Log($"  ├─ Total Processing Time: {performanceTimer.Elapsed.TotalSeconds:F2} seconds ({performanceTimer.ElapsedMilliseconds:N0} ms)");
        Log($"  ├─ Processing Speed: {(totalChildren > 0 && performanceTimer.Elapsed.TotalSeconds > 0 ? totalChildren / performanceTimer.Elapsed.TotalSeconds : 0):F1} objects/second");
        Log($"  ├─ Properties/Second: {(performanceTimer.Elapsed.TotalSeconds > 0 ? totalPropertiesAssigned / performanceTimer.Elapsed.TotalSeconds : 0):F1}");
        Log($"  └─ End Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");

        // Memory Usage
        Log($"");
        Log($"💾 MEMORY USAGE:");
        Log($"  ├─ Memory at Start: {FormatMemorySize(memoryAtStart)}");
        Log($"  ├─ Memory at End: {FormatMemorySize(memoryAtEnd)}");
        Log($"  ├─ Peak Memory Used: {FormatMemorySize(peakMemoryUsed)}");
        Log($"  ├─ Net Memory Change: {(memoryUsed >= 0 ? "+" : "")}{FormatMemorySize(Math.Abs(memoryUsed))}");
        Log($"  └─ Memory per Object: {(assignedCount > 0 ? FormatMemorySize(Math.Abs(memoryUsed) / assignedCount) : "N/A")}");

        // Configuration
        Log($"");
        Log($"⚙️ CONFIGURATION:");
        Log($"  ├─ BIM Platform: {currentTab}");
        Log($"  ├─ Target Object: {targetObject.name}");
        Log($"  └─ CSV File: {AssetDatabase.GetAssetPath(csvFile)}");

        // Top Elements by Properties
        if (elementStats.Count > 0)
        {
            Log($"");
            Log($"📁 TOP ELEMENTS BY PROPERTY COUNT:");
            int rank = 1;
            foreach (var stat in elementStats.OrderByDescending(s => s.Value.NonEmptyParameters).Take(10))
            {
                float fillRate = stat.Value.TotalParameters > 0 ?
                    (float)stat.Value.NonEmptyParameters / stat.Value.TotalParameters * 100 : 0;
                Log($"  {rank,2}. {stat.Key}: {stat.Value.NonEmptyParameters}/{stat.Value.TotalParameters} properties ({fillRate:F1}% fill rate)");
                rank++;
            }
            if (elementStats.Count > 10)
                Log($"  ... and {elementStats.Count - 10} more elements");
        }

        // Recommendations
        if (unmatchedCount > 0)
        {
            Log($"");
            Log($"💡 RECOMMENDATIONS:");
            Log($"  ⚠ {unmatchedCount:N0} objects could not be matched to CSV data.");
            Log($"  Possible causes:");
            Log($"    • Element ID mismatch between CSV and GameObject names");
            Log($"    • Missing rows in CSV file for some elements");
            Log($"    • Naming convention differences (Archicad vs Revit format)");
            Log($"  Suggested actions:");
            Log($"    • Verify CSV Element ID column matches GameObject naming");
            Log($"    • Check if all model elements have corresponding CSV entries");
            Log($"    • Try switching between Archicad/Revit mode if names don't match");
        }
        else
        {
            Log($"");
            Log($"✓ ALL OBJECTS SUCCESSFULLY MATCHED WITH METADATA!");
        }

        Log($"");
        Log($"════════════════════════════════════════════════════════════════");

        // Save log to file
        SaveLogToFile();
    }

    private void SaveLogToFile()
    {
        try
        {
            // Create the logs folder if it doesn't exist
            string logFolder = "Assets/BIMUniXchange/Logs";
            if (!Directory.Exists(logFolder))
            {
                Directory.CreateDirectory(logFolder);
            }

            // Generate filename with timestamp
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string fileName = $"MetadataAssignment_{timestamp}.txt";
            string filePath = Path.Combine(logFolder, fileName);

            // Build log content
            var content = new System.Text.StringBuilder();
            content.AppendLine("═══════════════════════════════════════════════════════════════════");
            content.AppendLine("               BIM METADATA ASSIGNMENT LOG");
            content.AppendLine("═══════════════════════════════════════════════════════════════════");
            content.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            content.AppendLine($"Target Object: {(targetObject != null ? targetObject.name : "N/A")}");
            content.AppendLine($"CSV File: {(csvFile != null ? AssetDatabase.GetAssetPath(csvFile) : "N/A")}");
            content.AppendLine($"BIM Platform: {currentTab}");
            content.AppendLine("═══════════════════════════════════════════════════════════════════");
            content.AppendLine();

            foreach (var log in logs)
            {
                content.AppendLine(log);
            }

            // Write to file
            File.WriteAllText(filePath, content.ToString());
            _lastLogFilePath = filePath;

            Log($"");
            Log($"📄 Log saved to: {filePath}");

            // Refresh asset database to show the new file
            AssetDatabase.Refresh();
        }
        catch (Exception ex)
        {
            UnityEngine.Debug.LogError($"Failed to save log file: {ex.Message}");
        }
    }

    private string ExtractArchicadElementId(string objectName)
    {
        // For Archicad FBX files, use the object name directly (trimmed)
        return objectName.Trim();
    }

    private string ExtractRevitElementId(string objectName)
    {
        // For Revit OBJ files, extract element ID from the end after the last underscore
        string name = objectName.Trim();
        int lastIndex = name.LastIndexOf('_');
        return lastIndex >= 0 && lastIndex < name.Length - 1 ?
            name.Substring(lastIndex + 1) : string.Empty;
    }

    Dictionary<string, Dictionary<string, string>> ParseCSV()
    {
        try
        {
            string csvPath = AssetDatabase.GetAssetPath(csvFile);
            string csvText;

            if (csvFile is TextAsset textAsset)
            {
                csvText = textAsset.text;
            }
            else
            {
                csvText = File.ReadAllText(csvPath);
            }

            var result = new Dictionary<string, Dictionary<string, string>>();
            var lines = csvText.Split('\n');

            if (lines.Length < 2)
            {
                Log("CSV has insufficient data (less than 2 lines)");
                return null;
            }

            // Parse headers
            var headers = SplitCSVLine(lines[0]);
            Log($"Found {headers.Length} columns in CSV");

            // Find Element ID column (support both naming conventions)
            int elementIdIndex = -1;
            // For Archicad CSV, prioritize columns in order of likelihood to match FBX object names
            string[] possibleIdColumns = {
                "ID and Categories.Element ID",   // Primary Archicad column (W-001, W-002, etc.)
                "General Parameters.Element ID",  // Alternative Archicad CSV format
                "General Parameters.Unique ID",   // Unique ID column
                "Element ID",
                "Element_Id",
                "ElementID",
                "ID"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                foreach (var possibleName in possibleIdColumns)
                {
                    if (headers[i].Equals(possibleName, StringComparison.OrdinalIgnoreCase))
                    {
                        elementIdIndex = i;
                        Log($"Found element ID column '{headers[i]}' at position {i}");
                        break;
                    }
                }
                if (elementIdIndex != -1) break;
            }

            if (elementIdIndex == -1)
            {
                Log($"No element ID column found. Available columns: {string.Join(", ", headers)}");
                return null;
            }

            // Parse data rows
            for (int i = 1; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;

                var values = SplitCSVLine(lines[i]);
                if (values.Length <= elementIdIndex) continue;

                var elementId = values[elementIdIndex].Trim();
                if (string.IsNullOrEmpty(elementId)) continue;

                var rowDict = new Dictionary<string, string>();
                for (int j = 0; j < headers.Length && j < values.Length; j++)
                {
                    rowDict[headers[j]] = values[j].Trim();
                }

                result[elementId] = rowDict;
            }

            Log($"Successfully parsed {result.Count} data rows");
            return result;
        }
        catch (Exception ex)
        {
            Log($"Error parsing CSV: {ex.Message}");
            return null;
        }
    }

    string[] SplitCSVLine(string line)
    {
        var result = new List<string>();
        var current = new System.Text.StringBuilder();
        bool inQuotes = false;

        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];

            if (c == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    current.Append('"');
                    i++; // Skip next quote
                }
                else
                {
                    inQuotes = !inQuotes;
                }
            }
            else if (c == ',' && !inQuotes)
            {
                result.Add(current.ToString());
                current.Clear();
            }
            else if (c != '\r') // Skip carriage returns
            {
                current.Append(c);
            }
        }

        result.Add(current.ToString());
        return result.ToArray();
    }
}
