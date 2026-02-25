using UnityEngine;
using UnityEditor;
using System.Collections.Generic;
using System.Xml;
using System.IO;
using Newtonsoft.Json;

public class ExportDifferencesByIdToXml : EditorWindow
{
    private GameObject originalObj;  // The original (unmodified) OBJ model
    private GameObject changedObj;   // The changed (modified) OBJ model

    #region Data Structures

    [System.Serializable]
    private class SceneObjectData
    {
        public string name;
        public int elementId;
        public Vector3 position;
        public Vector3 rotation;
        public Vector3 scale;
        public string materialName;
        public Color materialColor;
        // Add any additional metadata fields if needed
    }

    #endregion

    [MenuItem("Window/BIMUniXchange/ReUniXchange/Export Differences by Element ID to Revit XML", false, 50)]
    public static void ShowWindow()
    {
        GetWindow<ExportDifferencesByIdToXml>("Export Differences (By ID)");
    }

    private void OnGUI()
    {
        GUILayout.Label("Export Differences Between Two OBJ Models (By ID)", EditorStyles.boldLabel);

        originalObj = (GameObject)EditorGUILayout.ObjectField(
            "Original OBJ GameObject",
            originalObj,
            typeof(GameObject),
            true
        );

        changedObj = (GameObject)EditorGUILayout.ObjectField(
            "Changed OBJ GameObject",
            changedObj,
            typeof(GameObject),
            true
        );

        if (GUILayout.Button("Export Differences to XML"))
        {
            if (originalObj == null || changedObj == null)
            {
                EditorUtility.DisplayDialog(
                    "Missing Input",
                    "Please assign both an Original and a Changed GameObject.",
                    "OK"
                );
                return;
            }

            ExportDifferencesToXmlFile(originalObj, changedObj);
        }
    }

    /// <summary>
    /// Compares the original OBJ model and the changed OBJ model,
    /// gathers only the modified/new elements with a MeshFilter, 
    /// and writes them to an XML file.
    /// </summary>
    private void ExportDifferencesToXmlFile(GameObject original, GameObject changed)
    {
        // Collect data from the original model, keyed by element ID,
        // but only for objects that actually have a MeshFilter.
        Dictionary<int, SceneObjectData> originalDict = new Dictionary<int, SceneObjectData>();
        PopulateDataDictionary(original.transform, originalDict);

        // Compare with the changed model, gathering only differences
        // for objects that have a MeshFilter.
        List<SceneObjectData> differences = new List<SceneObjectData>();
        FindDifferences(changed.transform, originalDict, differences);

        // If no differences are found, bail out
        if (differences.Count == 0)
        {
            EditorUtility.DisplayDialog(
                "No Differences Found",
                "No elements were changed (or added) that have a MeshFilter. No export needed.",
                "OK"
            );
            return;
        }

        // Prompt user to choose a file path for the XML
        string filePath = EditorUtility.SaveFilePanel(
            "Save Revit-Style XML File",
            "",
            "DifferencesById.xml",
            "xml"
        );
        if (string.IsNullOrEmpty(filePath)) return;

        // Write only the modified/new elements to an XML file
        WriteRevitXml(filePath, differences);

        // Notify the user of completion
        Debug.Log($"Export complete. File saved at: {filePath}");
        EditorUtility.DisplayDialog(
            "Export Complete",
            $"Differences have been successfully exported to:\n{filePath}",
            "OK"
        );
    }

    /// <summary>
    /// Recursively populates a dictionary with element data from the original model,
    /// **only** for objects that have a MeshFilter, keyed by the element ID extracted from the name.
    /// </summary>
    private void PopulateDataDictionary(Transform parent, Dictionary<int, SceneObjectData> dict)
    {
        foreach (Transform child in parent)
        {
            // Check if this child has a MeshFilter
            MeshFilter filter = child.GetComponent<MeshFilter>();
            if (filter != null) // Only proceed if there's a MeshFilter
            {
                int elementId = ExtractIdFromName(child.name);
                if (elementId != 0)
                {
                    // Construct a data entry
                    SceneObjectData data = new SceneObjectData
                    {
                        name = child.name,
                        elementId = elementId,
                        position = child.localPosition,
                        rotation = child.localEulerAngles,
                        scale = child.localScale
                    };

                    // Capture material info if available
                    MeshRenderer rend = child.GetComponent<MeshRenderer>();
                    if (rend != null && rend.sharedMaterial != null)
                    {
                        data.materialName = rend.sharedMaterial.name;
                        if (rend.sharedMaterial.HasProperty("_Color"))
                        {
                            data.materialColor = rend.sharedMaterial.color;
                        }
                    }

                    // Store/overwrite this entry for the element ID
                    dict[elementId] = data;
                }
            }

            // Recurse regardless; even if this child doesn't have a MeshFilter,
            // its children might have one.
            PopulateDataDictionary(child, dict);
        }
    }

    /// <summary>
    /// Recursively checks each element in the changed model, comparing it
    /// to the originalDataDict by element ID. If differences are found,
    /// we add them to 'differences'. If the element is new (no ID match)
    /// but has a MeshFilter, it's considered different as well.
    /// </summary>
    private void FindDifferences(
        Transform changedTransform,
        Dictionary<int, SceneObjectData> originalDataDict,
        List<SceneObjectData> differences
    )
    {
        // Only proceed if there's a MeshFilter on the changed object
        MeshFilter filter = changedTransform.GetComponent<MeshFilter>();
        if (filter != null)
        {
            int changedId = ExtractIdFromName(changedTransform.name);

            // If this ID is in the original, compare
            if (changedId != 0 && originalDataDict.TryGetValue(changedId, out SceneObjectData originalData))
            {
                bool positionChanged = changedTransform.localPosition != originalData.position;
                bool rotationChanged = changedTransform.localEulerAngles != originalData.rotation;
                bool scaleChanged = changedTransform.localScale != originalData.scale;

                // Compare material
                MeshRenderer rend = changedTransform.GetComponent<MeshRenderer>();
                string changedMatName = "None";
                Color changedColor = Color.white;
                if (rend && rend.sharedMaterial)
                {
                    changedMatName = rend.sharedMaterial.name;
                    if (rend.sharedMaterial.HasProperty("_Color"))
                    {
                        changedColor = rend.sharedMaterial.color;
                    }
                }

                bool matNameDiffers = changedMatName != originalData.materialName;
                bool matColorDiffers = changedColor != originalData.materialColor;
                bool materialChanged = matNameDiffers || matColorDiffers;

                // If anything changed, record this element
                if (positionChanged || rotationChanged || scaleChanged || materialChanged)
                {
                    differences.Add(new SceneObjectData
                    {
                        name = changedTransform.name,
                        elementId = changedId,
                        position = changedTransform.localPosition,
                        rotation = changedTransform.localEulerAngles,
                        scale = changedTransform.localScale,
                        materialName = changedMatName,
                        materialColor = changedColor
                    });
                }
            }
            else
            {
                // If changedId is 0 or wasn't in originalDataDict -> new element,
                // but only if there's a MeshFilter (which we've already confirmed).
                MeshRenderer rend = changedTransform.GetComponent<MeshRenderer>();
                SceneObjectData newData = new SceneObjectData
                {
                    name = changedTransform.name,
                    elementId = changedId,
                    position = changedTransform.localPosition,
                    rotation = changedTransform.localEulerAngles,
                    scale = changedTransform.localScale,
                    materialName = rend && rend.sharedMaterial ? rend.sharedMaterial.name : "None",
                    materialColor = (rend && rend.sharedMaterial && rend.sharedMaterial.HasProperty("_Color"))
                                    ? rend.sharedMaterial.color
                                    : Color.white
                };
                differences.Add(newData);
            }
        }

        // Recurse for all children, even if the current object doesn't have a MeshFilter.
        foreach (Transform child in changedTransform)
        {
            FindDifferences(child, originalDataDict, differences);
        }
    }

    /// <summary>
    /// Writes only the "differences" (which all have MeshFilters) to an XML file
    /// in a Revit-like structure, including the element ID as both an attribute
    /// AND a child node.
    /// </summary>
    private void WriteRevitXml(string filePath, List<SceneObjectData> differences)
    {
        XmlDocument xmlDoc = new XmlDocument();

        // Root node
        XmlElement rootNode = xmlDoc.CreateElement("RevitExport");
        xmlDoc.AppendChild(rootNode);

        // Elements node
        XmlElement elementsNode = xmlDoc.CreateElement("Elements");
        rootNode.AppendChild(elementsNode);

        foreach (var diff in differences)
        {
            // <Element ... />
            XmlElement elemNode = xmlDoc.CreateElement("Element");
            elemNode.SetAttribute("name", diff.name);
            elemNode.SetAttribute("elementId", diff.elementId.ToString());

            // Optional: explicit ElementId node
            XmlElement idNode = xmlDoc.CreateElement("ElementId");
            idNode.InnerText = diff.elementId.ToString();
            elemNode.AppendChild(idNode);

            // Position
            XmlElement pos = xmlDoc.CreateElement("Position");
            pos.SetAttribute("x", diff.position.x.ToString("F3"));
            pos.SetAttribute("y", diff.position.y.ToString("F3"));
            pos.SetAttribute("z", diff.position.z.ToString("F3"));
            elemNode.AppendChild(pos);

            // Rotation
            XmlElement rot = xmlDoc.CreateElement("Rotation");
            rot.SetAttribute("x", diff.rotation.x.ToString("F3"));
            rot.SetAttribute("y", diff.rotation.y.ToString("F3"));
            rot.SetAttribute("z", diff.rotation.z.ToString("F3"));
            elemNode.AppendChild(rot);

            // Scale
            XmlElement scl = xmlDoc.CreateElement("Scale");
            scl.SetAttribute("x", diff.scale.x.ToString("F3"));
            scl.SetAttribute("y", diff.scale.y.ToString("F3"));
            scl.SetAttribute("z", diff.scale.z.ToString("F3"));
            elemNode.AppendChild(scl);

            // Material
            XmlElement mat = xmlDoc.CreateElement("Material");
            mat.SetAttribute("name", diff.materialName ?? "None");
            mat.SetAttribute("colorR", diff.materialColor.r.ToString("F3"));
            mat.SetAttribute("colorG", diff.materialColor.g.ToString("F3"));
            mat.SetAttribute("colorB", diff.materialColor.b.ToString("F3"));
            elemNode.AppendChild(mat);

            elementsNode.AppendChild(elemNode);
        }

        // Save the XML
        xmlDoc.Save(filePath);
    }

    /// <summary>
    /// Extracts a numeric ID from the GameObject name if it follows a pattern like:
    /// "A1-39-19A-M3_Interior_wall_90mm_1431734"
    /// and returns 1431734 as an integer.
    /// </summary>
    private int ExtractIdFromName(string objName)
    {
        // Split by underscore, parse the last part as an integer
        string[] parts = objName.Split('_');
        if (parts.Length < 2) return 0;

        string lastPart = parts[parts.Length - 1];
        if (int.TryParse(lastPart, out int parsedId))
        {
            return parsedId;
        }
        return 0;
    }
}
