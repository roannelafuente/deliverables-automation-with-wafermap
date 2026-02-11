# Deliverables Automation Tool with Wafermap

## 📖 Description
Automates semiconductor deliverables by converting CSVs to Excel, generating pivot tables, validating End Test numbers, and creating wafermap visualizations for yield and defect tracking. Built with Python, Tkinter, OpenPyXL, and xlwings, it streamlines workflows and ensures reproducible, audit‑ready insights.

## 🚀 Features
- **CSV → Excel Conversion**  
  Converts raw CSV deliverables into structured Excel workbooks with professional formatting.
- **Pivot Table Generation**  
  Automates fallout analysis by filtering C1_MARK values and calculating End Test fallout percentages.
- **End Test Validation**  
  Locates and validates End Test numbers against reference tables, highlighting limit conditions.
- **Wafermap Visualization**  
  Generates wafermaps with color‑coded grids for yield/defect tracking, mirrored headers, and clean formatting.
- **GUI Interface**  
  Tkinter‑based interface for file selection, filter dropdowns, and status logging.

## 🛠️ Tech Stack
- Python (automation & GUI)  
- Tkinter (user interface)  
- OpenPyXL (Excel file handling)  
- xlwings (pivot tables & wafermap generation)  
- CSV (data parsing)  

## 📂 Sample Files
- **Input:** `DEMO_WAFERMAP_08.wmap.csv`  
- **Output:** `DEMO_WAFERMAP_08.wmap.xlsx`  
- **Dashboard Screenshot:** `Deliverables Automation with Wafermap GUI.png`  

These files are included for demonstration purposes so recruiters and collaborators can quickly test the workflow and visualize the results.

## 📸 Screenshots
### GUI Dashboard
![Deliverables Automation with Wafermap GUI](Deliverables%20Automation%20with%20Wafermap%20GUI.png)

##🌟 **Impact (v1.0.0)**
- Marked the first release to include wafermap visualization, enabling engineers to see yield and defect distribution visually.  
- Automated CSV → Excel conversion and pivot table generation reduced manual effort in preparing deliverables.  
- Introduced fallout analysis by filtering `C1_MARK` values and calculating End Test fallout percentages.  
- Provided a Tkinter GUI for file selection, filter dropdowns, and status logging, making the workflow more accessible.  
- Random wafermap coloring served as a prototype stage, laying the foundation for deterministic production color codes in later releases.
  
## 📦 Download
The latest release (with `.exe` build) is available here:  
[➡️ Deliverables Automation Tool v1.0](https://github.com/roannelafuente/deliverables-automation-with-wafermap/releases/tag/v1.0.0)

## 👩‍💻 Author
**Rose Anne Lafuente**  
Licensed Electronics Engineer | Product Engineer II | Python Automation  
GitHub: [@roannelafuente](https://github.com/roannelafuente)
