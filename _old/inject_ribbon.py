#!/usr/bin/env python3
"""
inject_ribbon.py
================
Injects a customUI14.xml ribbon definition into a PowerPoint .pptm or .ppam file.

Usage:
    python inject_ribbon.py  input.pptm  customUI14.xml  output.ppam

How it works:
    .pptm/.ppam files are ZIP archives. This script:
    1. Copies every existing entry from the input file
    2. Adds customUI/customUI14.xml inside the archive
    3. Updates [Content_Types].xml to register the customUI part
    4. Updates _rels/.rels to link the customUI to the package

Requirements:
    Python 3.6+  (no external packages needed)
"""

import sys
import zipfile
import shutil
import os
import xml.etree.ElementTree as ET
from io import BytesIO


def inject_custom_ui(input_path: str, custom_ui_path: str, output_path: str):
    """Inject customUI14.xml into a .pptm/.ppam file."""

    # Read the customUI XML content
    with open(custom_ui_path, "r", encoding="utf-8") as f:
        custom_ui_content = f.read()

    # Read all entries from the input ZIP
    with zipfile.ZipFile(input_path, "r") as zin:
        entries = {}
        for item in zin.infolist():
            entries[item.filename] = zin.read(item.filename)

    # ── 1. Add customUI/customUI14.xml ──────────────────────────────
    custom_ui_zip_path = "customUI/customUI14.xml"
    entries[custom_ui_zip_path] = custom_ui_content.encode("utf-8")

    # ── 2. Update [Content_Types].xml ───────────────────────────────
    ct_xml = entries["[Content_Types].xml"].decode("utf-8")
    ct_tree = ET.fromstring(ct_xml)
    ns = "http://schemas.openxmlformats.org/package/2006/content-types"

    # Check if customUI content type already exists
    existing = [
        el for el in ct_tree
        if el.get("PartName", "") == "/customUI/customUI14.xml"
    ]
    if not existing:
        ET.SubElement(ct_tree, f"{{{ns}}}Override", {
            "PartName": "/customUI/customUI14.xml",
            "ContentType": "application/xml"
        })

    # Re-serialize
    ET.register_namespace("", ns)
    entries["[Content_Types].xml"] = ET.tostring(ct_tree, encoding="unicode", xml_declaration=True).encode("utf-8")

    # ── 3. Update _rels/.rels ───────────────────────────────────────
    rels_xml = entries["_rels/.rels"].decode("utf-8")
    rels_tree = ET.fromstring(rels_xml)
    rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"

    # Check if relationship already exists
    custom_ui_rel_type = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
    existing_rels = [
        el for el in rels_tree
        if el.get("Type", "") == custom_ui_rel_type
    ]
    if not existing_rels:
        # Find max rId number to create a new unique one
        max_id = 0
        for el in rels_tree:
            rid = el.get("Id", "")
            if rid.startswith("rId"):
                try:
                    num = int(rid[3:])
                    max_id = max(max_id, num)
                except ValueError:
                    pass
        new_id = f"rId{max_id + 1}"

        ET.SubElement(rels_tree, f"{{{rels_ns}}}Relationship", {
            "Id": new_id,
            "Type": custom_ui_rel_type,
            "Target": "customUI/customUI14.xml"
        })

    ET.register_namespace("", rels_ns)
    entries["_rels/.rels"] = ET.tostring(rels_tree, encoding="unicode", xml_declaration=True).encode("utf-8")

    # ── 4. Write the output file ────────────────────────────────────
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)

    print(f"✅ Ribbon injected successfully!")
    print(f"   Input:  {input_path}")
    print(f"   Output: {output_path}")
    print()
    print(f"Next steps:")
    print(f"  1. Open PowerPoint")
    print(f"  2. Go to File > Options > Add-ins")
    print(f"  3. At the bottom: Manage = 'PowerPoint Add-ins' > Go...")
    print(f"  4. Click 'Add New...' and select: {output_path}")
    print(f"  5. The 'Section Nav' tab should appear in the ribbon!")


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python inject_ribbon.py <input.pptm> <customUI14.xml> <output.ppam>")
        print()
        print("  input.pptm      - A macro-enabled PowerPoint file with your VBA code")
        print("  customUI14.xml  - The ribbon definition XML file")
        print("  output.ppam     - Output path for the add-in file")
        sys.exit(1)

    inject_custom_ui(sys.argv[1], sys.argv[2], sys.argv[3])
