


# 🏗️ BIMUniXchange

<div align="center">

![BIM Badge](https://img.shields.io/badge/BIM-Interoperability-blue?style=for-the-badge&logo=autodesk)
![Unity Badge](https://img.shields.io/badge/Unity-XR%20Ready-black?style=for-the-badge&logo=unity)
![XR Badge](https://img.shields.io/badge/XR-Compatible-purple?style=for-the-badge&logo=oculus)
![MIT License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

_Seamless AEC-to-XR Pipeline for Immersive Construction Visualization_

</div>

---
<img width="1672" height="941" alt="image" src="https://github.com/user-attachments/assets/3d1f0a98-5bfe-4764-bab2-172e382ff987" />

## 🛒 **Download & Installation**

### 🔧 **ReUniXchange for Revit**

**Now Available on Autodesk App Store!**

[![Autodesk App Store](https://img.shields.io/badge/Autodesk-App%20Store-0696D7?style=for-the-badge&logo=autodesk)](https://apps.autodesk.com/RVT/en/Detail/Index?id=6622158299438821562&appLang=en&os=Win64)

Download and install ReUniXchange directly from the Autodesk App Store for seamless integration with Revit 2024, 2025, and 2026.

### 🔧 ArchUniXchange for Archicad

**Available in this repository**

Download the **ArchUniXchange Python scripts** from this repository to enable the **Archicad** side of the BIM-to-Unity workflow.

### 🔧 BIMUniXchange for Unity

**BIMUniXchange_V1.1.0 is available in this repository in the releases on the right side of this page**

Download the **BIMUniXchange Unity package** from this repository to use the Unity-side integration for immersive BIM-to-XR workflows.
---

## 🚀 **Quick Start Guide**

### 🔧 **Revit Workflow**

#### 📐 Export from Autodesk Revit

```
1️⃣ Install ReUniXchange from Autodesk App Store or via ReUniXchange.msi
2️⃣ Launch the "Export to OBJ+CSV" command in Revit
3️⃣ Generate building geometry and element metadata
```

> **Compatible Versions:** Revit 2024, 2025, 2026

---

### 🏛️ **ArchiCAD Workflow**

#### 📊 Export from GRAPHISOFT ArchiCAD

```
1️⃣ Navigate to ArchiUniXchange/ directory
2️⃣ Execute Python automation scripts
3️⃣ Generate FBX geometry and CSV metadata files
```

> **Output Format:** FBX + CSV metadata for seamless Unity integration

---

### 🎮 **Unity Integration**

#### 🔗 Import into Unity Engine

```
1️⃣ Import UnityPackage/ into your Unity project
2️⃣ Drop exported OBJ/FBX and CSV files into Assets/
3️⃣ Launch BIMUniXchange window for metadata assignment
4️⃣ Build hierarchical BIM structures automatically
```

> **Features:** Automated metadata mapping, parametric object hierarchy, material assignment

---

### 🥽 **XR Deployment**

#### 🌐 Deploy to Extended Reality

```
1️⃣ Configure OpenXR settings for target platform
2️⃣ Build and deploy to your XR device
3️⃣ Experience immersive construction visualization
```

> **Supported Platforms:** Quest 3, Samsung S22 Ultra, Desktop VR Headsets

---

## 📊 **Performance Benchmarks**

<table>
<tr>
<th>🔄 <strong>Interoperability</strong></th>
<th>⚡ <strong>Performance</strong></th>
<th>👤 <strong>Usability</strong></th>
</tr>
<tr>
<td>
• CSV exports validated against IFC schema<br>
• Network-graph analysis for data integrity<br>
• Cross-platform BIM compatibility
</td>
<td>
• Tested with 4,000+ building elements<br>
• Maintains interactive frame rates in XR<br>
• Optimized geometry processing pipeline
</td>
<td>
• Clear metadata presentation<br>
• High-fidelity visual rendering<br>
• Effective 4D construction simulation
</td>
</tr>
</table>

---

## 🏆 **Key Features**

| Feature                      | Description                         | Status    |
| ---------------------------- | ----------------------------------- | --------- |
| 🔧 **Multi-CAD Support**     | Revit & ArchiCAD integration        | ✅ Active |
| 📱 **Cross-Platform XR**     | Quest, Mobile, Desktop VR           | ✅ Active |
| 🏗️ **4D Simulation**         | Construction sequence visualization | ✅ Active |
| 📊 **Metadata Preservation** | Complete BIM data transfer          | ✅ Active |
| ⚡ **Real-time Performance** | Interactive XR experience           | ✅ Active |

---

## 🕸️ **BIM Knowledge Graph Studio** (web)

**Launch:** [isaddiq.github.io/BIMUniXchange/src/ifctograph.html](https://isaddiq.github.io/BIMUniXchange/src/ifctograph.html) · try it instantly with the [sample building](https://isaddiq.github.io/BIMUniXchange/src/ifctograph.html?sample=ifc)

A browser-only workbench that turns IFC models (IFC2x3 / IFC4 / IFC4x3, `.ifc` or `.ifczip`) and BIMUniXchange Revit / ArchiCAD CSV exports into a semantic **knowledge graph**. Files never leave the browser. It replaces the former separate Revit, ArchiCAD and IFC analyzers — one Studio opens all three formats (samples: [IFC](https://isaddiq.github.io/BIMUniXchange/src/ifctograph.html?sample=ifc) · [Revit](https://isaddiq.github.io/BIMUniXchange/src/ifctograph.html?sample=revit) · [ArchiCAD](https://isaddiq.github.io/BIMUniXchange/src/ifctograph.html?sample=archicad)).

| Capability | Details |
| --- | --- |
| **Parsing** | Streaming ISO 10303-21 indexer with lazy argument decoding (geometry is never materialised; ~90 MB / 1.5 M instances index in about a second). Multi-line entities, `''` escapes and `\X2\` unicode are handled. |
| **Semantic model** | Every objectified `IfcRel*` becomes a typed edge with provenance; property & quantity sets resolved with project units; materials (layer / constituent / profile sets), classification references, documents, ports, zones, systems, tasks. |
| **Inference** | Hosting (`wall hosts door` via voids ∘ fills), MEP flow (`feeds` from port connectivity and flow direction), space adjacency and door / stair access (from space boundaries), storey location, type-inherited properties and materials, discipline classification (e.g. `LoadBearing` walls → Structural, system type → HVAC / Plumbing / Electrical). |
| **Ontology alignment** | Classes and relations mapped to **ifcOWL**, **BOT** (Site / Building / Storey / Space / Element, hasStorey, containsElement, hostsElement, adjacentZone) and **Brick** (equipment, `feeds`, `isPartOf`). |
| **Lenses** | Knowledge graph · Spatial hierarchy · Breakdown tree (e.g. Category → Family → Type) · MEP systems & flow · Physical topology · Space adjacency (with space-syntax integration) · Types & classification crosswalk · Materials · Information structure (class ↔ property set) · Schema meta-graph. |
| **Analytics** | Degree, PageRank, betweenness, harmonic closeness, clustering, components, Louvain communities & modularity, diameter, degree distribution, shortest paths, neighbourhood expansion. |
| **CSV semantics** | Revit: families, types, levels, systems, `Host Id` hosting, phases, category ↔ IFC class crosswalk. ArchiCAD: layers, composites as materials, structural function, zones. |
| **Model health** | Information-quality audit: duplicate / malformed GlobalIds, uncontained elements, missing types, materials, `Pset_*Common`, classification, IsExternal / FireRating, unconnected ports, flow elements without systems, spaces without boundaries or floor area. |
| **Exploration** | Global search, property query builder (wildcard `*.Property`, numeric operators, presets), property catalogue with coverage statistics, colour by any property, property values as nodes. |
| **Export** | RDF Turtle (ifcOWL + BOT + Brick + LBD PROPS), Neo4j Cypher, GraphML (Gephi / yEd / Cytoscape / NetworkX), JSON, node & edge CSV with metrics, publication-ready SVG, PNG, health and mapping reports. |

> The bundled [`sample/sample-building.ifc`](sample/sample-building.ifc) is a geometry-free IFC4 model with spaces, boundaries, MEP ports and systems, types, materials and Uniclass references. It contains a few deliberate data-quality issues for the model-health demo.

---

## 📜 **License & Attribution**

<div align="center">

**Released under the [MIT License](LICENSE)**

_Please cite the authors when implementing these tools in academic research or industry projects_

</div>

---

## 📚 **Academic Citation**

### 📖 **Citation Formats**

#### **APA (7th Edition)**

```
isaddiq. (2025). BIMUniXchange [Computer software]. GitHub.
https://github.com/isaddiq/BIMUniXchange
```

#### **BibTeX**

```bibtex
@software{isaddiq_BIMUniXchange_2025,
  author       = {Saddiq Ur Rehman},
  title        = {BIMUniXchange},
  year         = {2025},
  publisher    = {GitHub},
  howpublished = {\url{https://github.com/isaddiq/BIMUniXchange}},
}
```

---

## 👨‍🎓 **Research Contact**

<div align="center">

**Saddiq Ur Rehman**  
_PhD Candidate in Architecture_  
🏛️ **Kyung Hee University**  
📧 **saddiqurrehman@khu.ac.kr**

---

<sub>🏗️ _Advancing AEC Industry through Immersive Technology Solutions_ 🥽</sub>

</div>
