using UnityEngine;
using UnityEditor;

[CustomEditor(typeof(Metadata))]
public class UnifiedMetadataEditor : Editor
{
    public override void OnInspectorGUI()
    {
        Metadata metadata = (Metadata)target;

        EditorGUILayout.LabelField("Metadata Statistics", EditorStyles.boldLabel);
        EditorGUILayout.LabelField($"Has Metadata: {metadata.HasMetadata}");
        EditorGUILayout.LabelField($"Total Parameters: {metadata.Stats.TotalParameters}");
        EditorGUILayout.LabelField($"Non-empty Parameters: {metadata.Stats.NonEmptyParameters}");
        EditorGUILayout.LabelField($"Empty Parameters: {metadata.Stats.EmptyParameters}");
        EditorGUILayout.LabelField($"Assigned Parameters: {metadata.Stats.AssignedParameters}");

        EditorGUILayout.Space();

        if (metadata.Properties != null && metadata.Properties.Count > 0)
        {
            EditorGUILayout.LabelField("Metadata Properties", EditorStyles.boldLabel);

            EditorGUI.indentLevel++;
            foreach (var property in metadata.Properties)
            {
                EditorGUILayout.BeginVertical("box");

                EditorGUILayout.BeginHorizontal();
                EditorGUILayout.LabelField("Key:", EditorStyles.boldLabel, GUILayout.Width(40));
                EditorGUI.BeginChangeCheck();
                string newKey = EditorGUILayout.TextField(property.Key);
                if (EditorGUI.EndChangeCheck())
                {
                    property.Key = newKey;
                    EditorUtility.SetDirty(metadata);
                }
                EditorGUILayout.EndHorizontal();

                EditorGUILayout.BeginHorizontal();
                EditorGUILayout.LabelField("Value:", EditorStyles.boldLabel, GUILayout.Width(40));
                EditorGUI.BeginChangeCheck();
                string newValue = EditorGUILayout.TextField(property.Value);
                if (EditorGUI.EndChangeCheck())
                {
                    property.Value = newValue;
                    metadata.UpdateStats();
                    EditorUtility.SetDirty(metadata);
                }
                EditorGUILayout.EndHorizontal();

                EditorGUILayout.EndVertical();
                EditorGUILayout.Space();
            }
            EditorGUI.indentLevel--;
        }
        else
        {
            EditorGUILayout.HelpBox("No metadata assigned to this object.", MessageType.Info);
        }

        EditorGUILayout.Space();

        EditorGUILayout.BeginHorizontal();
        if (GUILayout.Button("Refresh Stats"))
        {
            metadata.UpdateStats();
            EditorUtility.SetDirty(metadata);
        }

        if (GUILayout.Button("Clear Metadata"))
        {
            if (EditorUtility.DisplayDialog("Clear Metadata",
                "Are you sure you want to clear all metadata from this object?",
                "Yes", "Cancel"))
            {
                metadata.ClearProperties();
                EditorUtility.SetDirty(metadata);
            }
        }
        EditorGUILayout.EndHorizontal();

        EditorGUILayout.Space();
        if (GUILayout.Button("Debug Info"))
        {
            Debug.Log($"Metadata Debug Info for {metadata.gameObject.name}:");
            Debug.Log($"- HasMetadata: {metadata.HasMetadata}");
            Debug.Log($"- Properties Count: {metadata.Properties.Count}");
            Debug.Log($"- Stats: Total={metadata.Stats.TotalParameters}, NonEmpty={metadata.Stats.NonEmptyParameters}");
            foreach (var prop in metadata.Properties)
            {
                Debug.Log($"  {prop.Key}: '{prop.Value}' (IsEmpty: {prop.IsEmpty})");
            }
        }

        if (GUI.changed)
        {
            serializedObject.ApplyModifiedProperties();
        }
    }
}
