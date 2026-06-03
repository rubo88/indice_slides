# Project: indice_slides

## Overview
This repo contains VBA macros for Office (PowerPoint and Excel) packaged as add-ins (.ppam / .xlam) with custom ribbon tabs.

## Project Structure

```
indice_slides/
├── SectionNavBar_Addin.bas    # Main PowerPoint macro (latest version)
├── SectionNavBar.bas          # Original version (before add-in adaptation)
├── SectionNavBar.ppam         # Packaged PowerPoint add-in
├── SectionNavBar.pptm         # Intermediate macro-enabled file
├── customUI14.xml             # Ribbon XML for the PowerPoint add-in
├── Guia_SectionNavBar.md      # User documentation (Spanish)
├── Guia_SectionNavBar.pdf     # PDF version of the guide
├── Guia_crear_xlam_con_ribbon.md  # Guide: how to create .xlam with ribbon from .bas
├── HOWTO_Create_Addin.md      # Guide: how to create .ppam add-in
├── inject_ribbon.py           # Python script to inject customUI XML into .pptm/.ppam
│
├── mismacros.xlam             # Excel add-in: utility macros
├── mismacros_v2.xlam          # Updated version
├── mismacros_v2.xlsm          # Macro-enabled workbook version
├── customUI14_mismacros.xml   # Ribbon XML for mismacros
├── RibbonCallbacks_mismacros.bas  # Ribbon callback wrappers for mismacros
├── Módulo1.bas                # VBA module from mismacros
├── Módulo3.bas                # VBA module from mismacros
│
├── colores_tramas/            # Subfolder for the chart themes macro
│   ├── colores_tramas_oficiales_actualizacion.xlam
│   ├── customUI14_colores_tramas.xml
│   └── RibbonCallbacks_colores_tramas.bas
│
└── 250305_MTBEyPUX.pptx       # Sample presentation
```

## SectionNavBar macro (PowerPoint)

### What it does
Generates a chevron navigation bar at the top of every slide based on presentation sections. Each chevron is a clickable hyperlink to the first slide of that section.

### Special section handling
- **"Portada"**: No nav bar at all (clean slide)
- **"Apéndice" / "Apendice"**: Nav bar appears but NO section is highlighted
- **Any other section**: Nav bar with current section highlighted

### Adaptive color scheme (per slide)
The macro detects each slide's background color:
- **White background**: black borders, black text, active section = black fill + white text, inactive = white fill
- **Non-white background**: blue (#8DA4C8) borders/text, active section = white fill, inactive = transparent

### Key constants (in the .bas)
`NAV_TOP=11`, `NAV_HEIGHT=15.5`, `NAV_LEFT=15.5`, `NAV_RIGHT=15.5`, `BORDER_WEIGHT=1`, `FONT_SIZE=10`

### Ribbon callbacks
- `RibbonCreateNavBar(control As IRibbonControl)` → calls `CreateSectionNavBar`
- `RibbonRemoveNavBar(control As IRibbonControl)` → calls `RemoveSectionNavBar`

### Packaging workflow
1. Create blank .pptm → Alt+F11 → import SectionNavBar_Addin.bas → save as .ppam from PowerPoint
2. Inject customUI14.xml using Office RibbonX Editor (Insert Office 2010+ Custom UI Part)
3. The .ppam can be distributed to partners — they just add it via File → Options → Add-ins

## mismacros (Excel)

### Macros included
- `ChangeSeriesFormula` — find/replace in chart series formulas (active chart)
- `ChangeSeriesFormulaAllCharts` — same but all charts in sheet
- `DuplicateCurrentSheetMultipleTimes` — duplicate sheet using names from "SheetNames" sheet
- `Copy_paste_absolute` — copy/paste preserving absolute references
- `TASAPCT` — custom worksheet function for percentage rates
- `EvalFormula` — custom worksheet function

### Ribbon: "Mis Macros" tab with 3 groups
- Series de Gráficos (2 buttons)
- Hojas (1 button)
- Utilidades (1 button: Copy/Paste Absolute)

## colores_tramas (Excel)

### What it does
Applies official color themes with patterns (tramas) to Excel charts. Has two palettes: Fondo Blanco and Fondo Azul. Includes a UserForm for selection.

### Ribbon: "Temas Gráficos" tab with 3 groups
- Aplicar Tema (opens UserForm)
- Fondo Blanco (active chart / all charts)
- Fondo Azul (active chart / all charts)

## Important notes
- Ribbon callbacks must accept `IRibbonControl` parameter
- The `GetCustomUI` VBA approach does NOT work for .xlam — must inject XML into the ZIP
- Use Office RibbonX Editor to inject customUI14.xml into .xlam/.ppam files
- imageMso icon reference: https://bert-toolkit.com/imagemso-list.html
- All documentation is in Spanish
