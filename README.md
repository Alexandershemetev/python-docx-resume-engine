# Programmatic Document Architecture Engine (`python-docx-resume-engine`)

This repository contains a production-ready document generation engine built in Python. It programmatically compiles raw professional historical data, assets, and design grids into a highly polished, ATS-optimized corporate resume in Microsoft Word (.docx) format.

Rather than relying on manual visual styling or rigid templates, this engine manages deep typographic grids, custom spatial constraints, and high-frequency document formatting directly via low-level file architecture manipulation.

## 🛠️ Advanced Engineering Problem Solving & Overrides

This script goes beyond standard `python-docx` wrappers by actively overriding hardcoded layout limitations and rendering anomalies inherent to the Microsoft Word format:

* **Dynamic OpenXML Fixed Layout Table Constraints:** Standard Word processors automatically shift and compress text blocks based on word lengths, which often drops company tags onto a secondary line. This engine programmatically injects a strict fixed-layout grid schema (`w:tblLayout type="fixed"`) and enforces strict twentieth-of-a-point (`dxa`) dimensional limits directly onto cell nodes, achieving a permanent 14.0 cm layout clearance that guarantees zero text-wrapping or misaligned boundaries.
* **Zero-Division Metadata Image Pipeline:** To circumvent rendering engine failures caused by image files missing standard vertical DPI headers (`vert_dpi == 0`), this module captures image binaries via `io.BytesIO` streams, uses `Pillow (PIL)` to audit the structural headers on-the-fly, and automatically re-serializes malformed files to 96 DPI in memory prior to canvas injection.
* **Automated Two-Pass Text Separation:** Mitigates the common MS Word rendering bug where justified multi-line texts stretch solitary trailing words across the full width of the document. This is achieved by programmatically slicing complex data lists into isolated text node components instead of relying on inline line breaks (`\n`).

## 📊 Structural & Typography Architecture
* **ATS Compatibility Optimization:** Built around standard 0.75-inch geometric document boundaries to balance maximum information density with machine-readability.
* **Corporate Typography Schema:** Uses strict Calibri hierarchy paired with custom-drawn deep navy (`#1B365D`) bottom paragraph boundaries for premium structural sectioning.

## 🚀 Execution & Quick Start

### Prerequisites
Ensure your local workspace has the required programmatic image and document handling dependencies installed:
```bash
pip install python-docx Pillow
```

### Local Compiler Compilation
Execute the script from your command prompt environment to instantly generate and verify the resume within your relative file structure:
```bash
python "CV in Python.py"
```
The output file will automatically compile and save directly into the designated 