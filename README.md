
---

## 💡 Getting Started

1. **Export from Revit**  
   - Install `ReUniXchange.msi` into Revit (2024–2026).
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
## 📚 Citation

If you use **BIMUniXchange** in your research, teaching, or development work, please cite it using one of the following formats:

**APA (7th Edition):**  
`isaddiq. (2025). BIMUniXchange [Computer software]. GitHub. https://github.com/isaddiq/BIMUniXchange`

**BibTeX:**
```bibtex
@software{isaddiq_BIMUniXchange_2025,
  author       = {Saddiq Ur Rehman},
  title        = {BIMUniXchange},
  year         = {2025},
  publisher    = {GitHub},
  howpublished = {\url{https://github.com/isaddiq/BIMUniXchange}},
}
---


## 📬 Contact

**Saddiq Ur Rehman**  
PhD Candidate, Kyung Hee University  
✉️ saddiqurrehman@khu.ac.kr
