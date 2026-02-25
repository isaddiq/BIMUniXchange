using UnityEngine;
using UnityEditor;
using System.Linq;
using System.Collections.Generic;

[CustomEditor(typeof(Metadata))]
[CanEditMultipleObjects]
public class MetadataInspector : Editor
{
    private Vector2 scrollPosition;
    private bool showAllProperties = false;
    private string searchFilter = "";
    private bool showEmptyProperties = true;
    private Dictionary<string, bool> foldoutStates = new Dictionary<string, bool>();

    public override void OnInspectorGUI()
    {
        var metadata = (Metadata)target;

        EditorGUI.BeginChangeCheck();

        EditorGUILayout.LabelField("BIM Element Metadata", EditorStyles.boldLabel);
        EditorGUILayout.Space();

        // Statistics section
        DrawStatisticsSection(metadata);
        EditorGUILayout.Space();

        // Filter and options
        DrawFilterSection();
        EditorGUILayout.Space();

        // Metadata Properties section
        DrawMetadataProperties(metadata);
        EditorGUILayout.Space();

        // Action buttons
        DrawActionButtons(metadata);

        // Apply changes if any
        if (EditorGUI.EndChangeCheck())
        {
            metadata.UpdateStats();
            EditorUtility.SetDirty(metadata);
        }
    }

    private void DrawStatisticsSection(Metadata metadata)
    {
        EditorGUILayout.LabelField("Statistics", EditorStyles.boldLabel);
        using (new EditorGUI.IndentLevelScope())
        {
            EditorGUILayout.LabelField($"Total Parameters: {metadata.Stats.TotalParameters}");
            EditorGUILayout.LabelField($"Assigned Parameters: {metadata.Stats.AssignedParameters}");

            // Show non-empty in green
            var oldColor = GUI.contentColor;
            GUI.contentColor = new Color(0.2f, 0.7f, 0.2f);
            EditorGUILayout.LabelField($"Non-empty Parameters: {metadata.Stats.NonEmptyParameters}");
            GUI.contentColor = oldColor;

            // Show empty in orange/red
            GUI.contentColor = new Color(0.9f, 0.4f, 0.1f);
            EditorGUILayout.LabelField($"Empty/Undefined Parameters: {metadata.Stats.EmptyParameters}");
            GUI.contentColor = oldColor;

            if (metadata.Stats.TotalParameters > 0)
            {
                float completionPercentage = (float)metadata.Stats.NonEmptyParameters / metadata.Stats.TotalParameters * 100f;
                EditorGUILayout.LabelField($"Completion: {completionPercentage:F1}%");

                // Progress bar
                Rect progressRect = GUILayoutUtility.GetRect(0, 20, GUILayout.ExpandWidth(true));
                EditorGUI.ProgressBar(progressRect, completionPercentage / 100f, $"{completionPercentage:F1}%");
            }
        }
    }

    private void DrawFilterSection()
    {
        EditorGUILayout.LabelField("Display Options", EditorStyles.boldLabel);
        using (new EditorGUI.IndentLevelScope())
        {
            EditorGUILayout.BeginHorizontal();
            EditorGUILayout.LabelField("Search:", GUILayout.Width(50));
            searchFilter = EditorGUILayout.TextField(searchFilter);
            if (GUILayout.Button("Clear", GUILayout.Width(50)))
            {
                searchFilter = "";
                GUI.FocusControl(null);
            }
            EditorGUILayout.EndHorizontal();

            EditorGUILayout.BeginHorizontal();
            bool newShowEmpty = EditorGUILayout.Toggle("Show Empty/Undefined", showEmptyProperties);
            if (newShowEmpty != showEmptyProperties)
            {
                showEmptyProperties = newShowEmpty;
                Repaint();
            }
            bool newShowAll = EditorGUILayout.Toggle("Show All Properties", showAllProperties);
            if (newShowAll != showAllProperties)
            {
                showAllProperties = newShowAll;
                Repaint();
            }
            EditorGUILayout.EndHorizontal();

            // Add a help box to clarify what empty means
            if (!showEmptyProperties)
            {
                EditorGUILayout.HelpBox("Empty properties (including '<undefined>') are hidden. Enable 'Show Empty/Undefined' to see them.", MessageType.Info);
            }
        }
    }

    private void DrawMetadataProperties(Metadata metadata)
    {
        EditorGUILayout.LabelField("Metadata Properties", EditorStyles.boldLabel);

        if (metadata.Properties.Count == 0)
        {
            EditorGUILayout.HelpBox("No metadata assigned to this element.", MessageType.Info);
            DrawAddPropertySection(metadata);
            return;
        }

        // Filter properties based on search and settings
        var filteredProperties = FilterProperties(metadata.Properties);

        if (filteredProperties.Count == 0)
        {
            EditorGUILayout.HelpBox("No properties match the current filter.", MessageType.Info);
            return;
        }

        // Scrollable area for properties
        scrollPosition = EditorGUILayout.BeginScrollView(scrollPosition, GUILayout.MaxHeight(400));

        DrawEditableProperties(filteredProperties, metadata);

        EditorGUILayout.EndScrollView();

        // Add new property section
        EditorGUILayout.Space();
        DrawAddPropertySection(metadata);
    }

    private List<MetadataProperty> FilterProperties(List<MetadataProperty> properties)
    {
        var filtered = properties.AsEnumerable();

        // Apply search filter
        if (!string.IsNullOrEmpty(searchFilter))
        {
            filtered = filtered.Where(p =>
                p.Key.ToLower().Contains(searchFilter.ToLower()) ||
                (p.Value != null && p.Value.ToLower().Contains(searchFilter.ToLower())));
        }

        // Apply empty property filter
        if (!showEmptyProperties)
        {
            filtered = filtered.Where(p => !p.IsEmpty);
        }

        var result = filtered.ToList();

        // Limit display if not showing all
        if (!showAllProperties && result.Count > 20)
        {
            result = result.Take(20).ToList();
        }

        return result;
    }

    private void DrawEditableProperties(List<MetadataProperty> properties, Metadata metadata)
    {
        for (int i = 0; i < properties.Count; i++)
        {
            var property = properties[i];

            // Highlight empty properties with a different background color
            if (property.IsEmpty)
            {
                GUI.backgroundColor = new Color(1f, 0.9f, 0.9f); // Light red tint for empty
            }

            EditorGUILayout.BeginHorizontal();

            // Key field (editable)
            EditorGUILayout.LabelField("Key:", GUILayout.Width(30));
            string newKey = EditorGUILayout.TextField(property.Key, GUILayout.MinWidth(120));

            // Value field (editable) with visual indicator for empty values
            EditorGUILayout.LabelField("Value:", GUILayout.Width(40));

            // Show placeholder text for truly empty values
            string displayValue = property.Value ?? "";
            if (property.IsEmpty && string.IsNullOrEmpty(displayValue))
            {
                displayValue = "";
            }

            string newValue = EditorGUILayout.TextField(displayValue, GUILayout.MinWidth(120));

            // Show empty indicator
            if (property.IsEmpty)
            {
                EditorGUILayout.LabelField("(empty)", EditorStyles.miniLabel, GUILayout.Width(50));
            }

            // Delete button
            if (GUILayout.Button("×", GUILayout.Width(25)))
            {
                metadata.Properties.Remove(property);
                EditorUtility.SetDirty(metadata);
                GUI.backgroundColor = Color.white;
                break;
            }

            EditorGUILayout.EndHorizontal();

            // Reset background color
            GUI.backgroundColor = Color.white;

            // Update property if changed
            if (newKey != property.Key)
            {
                property.Key = newKey;
                EditorUtility.SetDirty(metadata);
            }

            if (newValue != property.Value)
            {
                property.Value = newValue;
                metadata.UpdateStats();
                EditorUtility.SetDirty(metadata);
            }

            // Visual separator for better readability
            if (i < properties.Count - 1)
            {
                EditorGUILayout.Space(2);
            }
        }

        // Show truncation notice
        if (!showAllProperties && metadata.Properties.Count > 20)
        {
            EditorGUILayout.Space();
            EditorGUILayout.HelpBox($"Showing 20 of {metadata.Properties.Count} properties. Enable 'Show All Properties' to see more.", MessageType.Info);
        }
    }

    private string newPropertyKey = "";
    private string newPropertyValue = "";

    private void DrawAddPropertySection(Metadata metadata)
    {
        EditorGUILayout.LabelField("Add New Property", EditorStyles.boldLabel);
        using (new EditorGUI.IndentLevelScope())
        {
            EditorGUILayout.BeginHorizontal();
            EditorGUILayout.LabelField("Key:", GUILayout.Width(30));
            newPropertyKey = EditorGUILayout.TextField(newPropertyKey, GUILayout.MinWidth(120));
            EditorGUILayout.LabelField("Value:", GUILayout.Width(40));
            newPropertyValue = EditorGUILayout.TextField(newPropertyValue, GUILayout.MinWidth(120));

            using (new EditorGUI.DisabledScope(string.IsNullOrEmpty(newPropertyKey)))
            {
                if (GUILayout.Button("Add", GUILayout.Width(50)))
                {
                    metadata.AddProperty(newPropertyKey, newPropertyValue);
                    newPropertyKey = "";
                    newPropertyValue = "";
                    EditorUtility.SetDirty(metadata);
                }
            }
            EditorGUILayout.EndHorizontal();
        }
    }

    private void DrawActionButtons(Metadata metadata)
    {
        EditorGUILayout.BeginHorizontal();

        if (GUILayout.Button("Clear All Properties"))
        {
            if (EditorUtility.DisplayDialog("Clear Metadata",
                "Are you sure you want to clear all metadata properties for this element?",
                "Yes", "Cancel"))
            {
                metadata.ClearProperties();
                EditorUtility.SetDirty(metadata);
            }
        }

        if (metadata.Properties.Count > 0 && GUILayout.Button("Refresh Stats"))
        {
            metadata.UpdateStats();
            EditorUtility.SetDirty(metadata);
        }

        if (GUILayout.Button("Export to CSV"))
        {
            ExportMetadataToCSV(metadata);
        }

        EditorGUILayout.EndHorizontal();

        // Show a help box if this is a prefab
        if (PrefabUtility.IsPartOfPrefabAsset(metadata.gameObject))
        {
            EditorGUILayout.Space();
            EditorGUILayout.HelpBox("This is part of a prefab. Changes may affect all instances.", MessageType.Info);
        }
    }

    private void ExportMetadataToCSV(Metadata metadata)
    {
        string path = EditorUtility.SaveFilePanel("Export Metadata to CSV", "",
            $"{metadata.gameObject.name}_metadata.csv", "csv");

        if (!string.IsNullOrEmpty(path))
        {
            try
            {
                var csv = new System.Text.StringBuilder();
                csv.AppendLine("Key,Value");

                foreach (var prop in metadata.Properties)
                {
                    // Escape commas and quotes in CSV
                    string key = prop.Key?.Replace("\"", "\"\"") ?? "";
                    string value = prop.Value?.Replace("\"", "\"\"") ?? "";

                    if (key.Contains(",") || key.Contains("\"") || key.Contains("\n"))
                        key = $"\"{key}\"";
                    if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                        value = $"\"{value}\"";

                    csv.AppendLine($"{key},{value}");
                }

                System.IO.File.WriteAllText(path, csv.ToString());
                EditorUtility.DisplayDialog("Export Complete",
                    $"Metadata exported successfully to:\n{path}", "OK");
            }
            catch (System.Exception e)
            {
                EditorUtility.DisplayDialog("Export Failed",
                    $"Failed to export metadata:\n{e.Message}", "OK");
            }
        }
    }
}
