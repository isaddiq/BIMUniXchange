
---

## 💡 Getting Started

1. **Export from Revit**  
   - Install `ReUniXchange.bundle` into Revit (2024–2026).
   - Use the “Export to OBJ+CSV” command.

2. **Export from Archicad**  
   - Run the Python scripts in `ArchiUniXchange/`.
   - Generate FBX geometry and CSV metadata.

3. **Import into Unity**  
   - Add the `UnityPackage/` to your project.
   - Drop exported OBJ/FBX and CSV into `Assets/`.
   - Open the BIMUniXchange window to assign metadata and build hierarchies.

4. **Deploy to XR**  
   - Configure OpenXR settings.
   - Build and run on your target device (Quest 3, S22 Ultra, desktop VR).

---

## 🧪 Evaluation

- **Interoperability**  
  Compared Revit/ArchiCAD CSV exports against IFC schema with network‑graph analysis.

- **Performance**  
  Tested on models with >4,000 elements; maintains interactive frame rates in XR.

- **Usability**  
  User studies confirmed clear metadata presentation, high visual fidelity, and effective 4D simulation comprehension.

---

## 📜 License

Released under the [MIT License](LICENSE). Please cite the authors when using these tools in academic or industry projects.

---

## 📬 Contact

**Saddiq Ur Rehman**  
PhD Candidate, Kyung Hee University  
✉️ saddiqurrehman@khu.ac.kr
