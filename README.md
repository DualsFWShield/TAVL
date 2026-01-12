# TAVL Relevés Tool

![Main Interface](showcase/desktop-edit.png)

Web tool for importing, editing, and exporting survey XLSX files (Matrix structure).

## Features

### Core
- **Excel Import**: Load `.xlsx` files (drag & drop). Automatic analysis of file structure (Categories, Questions, Data Types).
- **Premium Interface**: Dark, modern UI optimized for readability.
- **Dynamic Forms**: Input fields generated based on Excel headers:
  - **Radio Buttons**: `v/f` (True/False) and `o/n` (Yes/No).
  - **Tri-State Selectors**: `G/M/F` (Bleacher/Mobile/Fixed).
  - **Date Pickers**: Calendar selection with standard format (`dd/mm/yyyy`).
  - **Smart Inputs**: Number fields and auto-expanding text areas.
- **Sidebar Navigation**: Dynamically generated list of "Auditoires" (Auditoriums).
- **Auto-Save**: Changes are persisted locally (IndexedDB). Session restore available on reload.
  
  ![Session Restore](showcase/desktop-saved.png)

- **Smart Export**: Exports to XLSX preserving **original formatting** (colors, fonts, borders).

### Advanced Tools
- **Read-Only Protection**: Critical structural fields (Building, Auditorium, Announced Capacity) are locked by default to prevent accidental edits.
- **Force Edit Mode (🔓)**: Unlock all fields temporarily via the lock icon in the header.
- **Magic Fill (⚡)**: Automate filling of standard fields:
  - Sets Date to "Today".
  - Copies "Announced Capacity" to "Real Capacity" if empty.
  - Sets 'Yes'/'True' for standard checks.
  - **Smart Exception**: Sets 'No' for negative attributes like "Humidity" or "Water Infiltration".
  - **Safety**: Skips standard Optional fields.
- **Smart Navigation (⬇)**: customized Floating Action Button:
  - Jumps to the **next empty mandatory field**.
  - Starts searching *after* the currently focused field (cursor awareness).
  - Wraps around to the start of the form.
  - Skips sibling radio buttons for faster traversal.
- **Badges & Validation**:
  - **(Facultatif)**: Optional fields (detected via grey/patterned Excel cells) are clearly marked.
  - **Test Manual Requis**: Replaces "(testé)" labels with a clear red badge.

## User Guide

1. **Open**: Launch `index.html` in a modern browser (Chrome, Edge).
2. **Import**: Drag & drop your `.xlsx` file (e.g., `Barbe.xlsx`).
   
   <img src="showcase/desktop-drop.png" width="45%" /> <img src="showcase/mobile-drop.png" width="20%" />

3. **Navigate**: Click an Auditorium name in the sidebar.
4. **Edit**:
   - Tab through fields or use the **⬇ Button** to jump to the next empty task.
   - Use **⚡ Fill** to pre-fill standard "All Good" values for a room.
   - If a structural error exists in the source, use **🔓 Unlock** to fix it.
   
   ![Mobile Interface](showcase/mobile-edit.png)

5. **Export**: Click **"Exporter le relevé"** to download the completed file.

## Excel Structure & Constraints

The tool relies on a specific "Matrix" structure in the Excel file.

### Critical Rows (Fixed Positions)
- **Row 3**: **Categories** (Primary Header, e.g., "Mobilier", "Sécurité").
- **Row 4**: **Questions** (Secondary Header, e.g., "Nombre de places", "Extincteur présent ?").
- **Row 5**: **Data Types** (Defines input type).

### Supported Data Types (Row 5 - Case Insensitive)
- `v/f` : True/False (Vrai/Faux)
- `o/n` : Yes/No (Oui/Non)
- `date` or `..date..` : Date picker
- `nombre` : Numeric input
- `gmf` : Gradin/Mobile/Fixe (Tri-state)
- `text` (or empty) : Default text area

### Keywords & Logic Dependencies
Certain features rely on specific keywords in **Row 3 (Category)** or **Row 4 (Question)**. These logic rules are **Keyword Sensitive** (partial match, case insensitive).

| Feature | Trigger Keywords (in Category or Question) | Effect |
| :--- | :--- | :--- |
| **Identity** | `Auditoires` | Identifies the column used for the sidebar list. |
| **Read-Only** | `Bâtiment`, `Auditoires`, `Capacité annoncée` | Locks the field. |
| **Magic Fill** | `Capacité réelle`, `Réellement fonctionnelles` | Copies value from "Capacité annoncée". |
| **Magic Fill** | `Date de passage` | Fills with Today's date. |
| **Magic Fill** | `Humidité`, `Infiltration` | Defaults to "Non" (N) instead of "Oui". |
| **GMF** | `Gradin` + `Mobile` | Forces GMF radio type if not specified. |

### Safe Modifications (What you can change in Excel)
- ✅ **Add Columns**: You can add new columns anywhere if they have headers in Rows 3, 4, 5.
- ✅ **Rename Headers**: You can rename most headers, **EXCEPT** those containing the keywords listed above if you want to keep the special logic attached to them.
- ✅ **Change Colors**:
  - **Pattern/Hatch Fill**: Any cell with a pattern fill (dots, lines) will be detected as **Optional** (Facultatif).
  - **Solid Colors**: Preserved on export but ignored by logic.

### Unsafe Modifications (What breaks the tool)
- ❌ **Moving Header Rows**: Rows 3, 4, 5 **MUST** remain the header rows. Do not insert rows above them.
- ❌ **Deleting Identity Column**: One column must have "Auditoires" in the header to generate the list.

## Technologies

- **HTML5 / CSS3** (Vanilla)
- **JavaScript** (ES6+)
- **ExcelJS**: For high-fidelity Excel reading/writing.
- **IndexedDB**: For local data persistence.
