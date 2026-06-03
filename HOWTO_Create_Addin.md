# How to Create the Section Nav Bar Add-in (.ppam)

## What You'll Get

A `.ppam` file that, when installed by anyone, adds a **"Section Nav"** tab
to the PowerPoint ribbon with two large buttons:

- **Create Nav Bar** — builds the navigation bar on all slides
- **Remove Nav Bar** — removes all navigation bar shapes

---

## Method A: Using the Python Script (Recommended)

This is the fastest method if you have Python installed.

### Step 1: Create the .pptm base file

1. Open PowerPoint → create a **blank presentation**
2. Press **Alt+F11** to open the VBA Editor
3. In the VBA Editor: **File → Import File** → select `SectionNavBar_Addin.bas`
4. Close the VBA Editor
5. **File → Save As** → choose type **PowerPoint Macro-Enabled Presentation (.pptm)**
6. Save as `SectionNavBar.pptm`

### Step 2: Inject the ribbon and create the .ppam

Open a terminal/command prompt in the folder containing the files and run:

```
python inject_ribbon.py SectionNavBar.pptm customUI14.xml SectionNavBar.ppam
```

Done! You now have `SectionNavBar.ppam` ready to distribute.

---

## Method B: Using Office RibbonX Editor (No Python Needed)

### Step 1: Create the .ppam base file

1. Open PowerPoint → create a **blank presentation**
2. Press **Alt+F11** to open the VBA Editor
3. In the VBA Editor: **File → Import File** → select `SectionNavBar_Addin.bas`
4. Close the VBA Editor
5. **File → Save As** → choose type **PowerPoint Add-in (.ppam)**
6. Save as `SectionNavBar.ppam`
7. **Close PowerPoint completely**

### Step 2: Inject the ribbon XML

1. Download **Office RibbonX Editor** from:
   https://github.com/fernandreu/office-ribbonx-editor/releases
2. Open `SectionNavBar.ppam` in the RibbonX Editor
3. Right-click the file → **Insert Office 2010+ Custom UI Part**
4. Paste the contents of `customUI14.xml` into the editor pane
5. Click the **Validate** button (green checkmark) to confirm no errors
6. **Save** and close

---

## How to Install (For You and Your Partners)

1. Copy the `.ppam` file to a local folder
   (e.g., `C:\Users\<Name>\AppData\Roaming\Microsoft\AddIns\`)
2. Open PowerPoint
3. Go to **File → Options → Add-ins**
4. At the bottom, set **Manage:** to **PowerPoint Add-ins** → click **Go...**
5. Click **Add New...** → browse to the `.ppam` file → **OK**
6. Check the box next to "SectionNavBar" → **Close**

The **"Section Nav"** tab now appears in the ribbon and will load
automatically every time PowerPoint opens.

### To Uninstall

1. **File → Options → Add-ins → PowerPoint Add-ins → Go...**
2. Uncheck or Remove the add-in

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Macros are disabled | File → Options → Trust Center → Trust Center Settings → Macro Settings → Enable all macros (or "with notification") |
| Ribbon tab doesn't appear | Close and reopen PowerPoint. If still missing, re-inject the customUI XML |
| "Sub or Function not defined" | Make sure the .bas module was imported correctly and the Sub names match exactly |
| Partners can't run macros | They need to enable macros in Trust Center, and the .ppam must be in a Trusted Location |

### Adding a Trusted Location (Recommended for Partners)

1. **File → Options → Trust Center → Trust Center Settings → Trusted Locations**
2. Click **Add new location...**
3. Browse to the folder containing the `.ppam` file
4. Check **Subfolders of this location are also trusted**
5. Click **OK**

This avoids the security warning every time PowerPoint opens.

---

## Files Included

| File | Purpose |
|------|---------|
| `SectionNavBar_Addin.bas` | VBA module with ribbon callbacks |
| `customUI14.xml` | Ribbon tab definition (2 buttons) |
| `inject_ribbon.py` | Python script to inject ribbon into .pptm → .ppam |
| `HOWTO_Create_Addin.md` | This guide |
