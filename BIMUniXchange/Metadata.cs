using System;
using System.Collections.Generic;
using UnityEngine;
using System.Linq;

[System.Serializable]
public class MetadataProperty
{
    public string Key;
    public string Value;
    public bool IsEmpty => string.IsNullOrEmpty(Value) ||
                          Value.Equals("<undefined>", StringComparison.OrdinalIgnoreCase);
}

[System.Serializable]
public class MetadataStats
{
    [SerializeField] private int totalParameters;
    [SerializeField] private int emptyParameters;
    [SerializeField] private int nonEmptyParameters;
    [SerializeField] private int assignedParameters;

    public int TotalParameters { get => totalParameters; set => totalParameters = value; }
    public int EmptyParameters { get => emptyParameters; set => emptyParameters = value; }
    public int NonEmptyParameters { get => nonEmptyParameters; set => nonEmptyParameters = value; }
    public int AssignedParameters { get => assignedParameters; set => assignedParameters = value; }

    public void Reset()
    {
        totalParameters = 0;
        emptyParameters = 0;
        nonEmptyParameters = 0;
        assignedParameters = 0;
    }
}

public class Metadata : MonoBehaviour
{
    [SerializeField] private List<MetadataProperty> properties = new List<MetadataProperty>();
    [SerializeField] private MetadataStats stats = new MetadataStats();
    [SerializeField] private bool hasMetadata = false;

    public List<MetadataProperty> Properties => properties;
    public MetadataStats Stats => stats;
    public bool HasMetadata => hasMetadata && properties.Count > 0;

    public void AddProperty(string key, string value)
    {
        if (string.IsNullOrEmpty(key)) return;

        var existingProperty = properties.Find(prop => prop.Key.Equals(key, StringComparison.OrdinalIgnoreCase));
        if (existingProperty != null)
        {
            existingProperty.Value = value ?? "";
        }
        else
        {
            properties.Add(new MetadataProperty { Key = key, Value = value ?? "" });
        }
        RefreshMetadata();

#if UNITY_EDITOR
        // Mark the component as dirty when properties are added/modified
        UnityEditor.EditorUtility.SetDirty(this);
#endif
    }

    public void UpdateStats(Dictionary<string, string> csvData = null)
    {
        if (csvData != null)
        {
            stats.TotalParameters = csvData.Count;
            stats.EmptyParameters = csvData.Values.Count(value =>
                string.IsNullOrEmpty(value) ||
                value.Equals("<undefined>", StringComparison.OrdinalIgnoreCase));
            stats.NonEmptyParameters = csvData.Values.Count(value =>
                !string.IsNullOrEmpty(value) &&
                !value.Equals("<undefined>", StringComparison.OrdinalIgnoreCase));
        }
        else
        {
            stats.TotalParameters = properties.Count;
            stats.EmptyParameters = properties.Count(p => p.IsEmpty);
            stats.NonEmptyParameters = properties.Count(p => !p.IsEmpty);
        }
        stats.AssignedParameters = properties.Count;
    }

    public string GetProperty(string key)
    {
        var property = properties.Find(prop => prop.Key == key);
        return property?.Value;
    }

    public void ClearProperties()
    {
        properties.Clear();
        stats.Reset();
        hasMetadata = false;
    }

    public void AssignCSVData(Dictionary<string, string> row)
    {
        ClearProperties();
        foreach (var kvp in row)
        {
            if (!string.IsNullOrEmpty(kvp.Key))
            {
                // Add all properties, including empty ones for completeness
                AddProperty(kvp.Key, kvp.Value ?? "");
            }
        }
        UpdateStats(row);
        hasMetadata = true;

#if UNITY_EDITOR
        // Ensure the component is marked as dirty in the editor
        UnityEditor.EditorUtility.SetDirty(this);
#endif
    }

    private void RefreshMetadata()
    {
        hasMetadata = properties.Count > 0;
        UpdateStats();
    }

    // Compatibility methods for existing code
    public string GetValue(string key) => GetProperty(key);
    public bool HasKey(string key) => properties.Any(p => p.Key.Equals(key, StringComparison.OrdinalIgnoreCase));
    public void AddMetadataEntry(string key, string value) => AddProperty(key, value);
}
