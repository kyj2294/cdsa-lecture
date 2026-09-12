"""Remove embedded font binaries from a PPTX while preserving slides and layouts."""
from __future__ import annotations

import argparse
import pathlib
import zipfile
import xml.etree.ElementTree as ET


PML = "http://schemas.openxmlformats.org/presentationml/2006/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def sanitize(source: pathlib.Path, output: pathlib.Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source, "r") as src, zipfile.ZipFile(
        output, "w", compression=zipfile.ZIP_DEFLATED
    ) as dst:
        for item in src.infolist():
            name = item.filename
            if name.startswith("ppt/fonts/"):
                continue

            data = src.read(name)
            if name == "ppt/presentation.xml":
                root = ET.fromstring(data)
                embedded = root.find(f"{{{PML}}}embeddedFontLst")
                if embedded is not None:
                    root.remove(embedded)
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            elif name == "ppt/_rels/presentation.xml.rels":
                root = ET.fromstring(data)
                for rel in list(root):
                    if rel.get("Type", "").endswith("/font"):
                        root.remove(rel)
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            elif name == "[Content_Types].xml":
                root = ET.fromstring(data)
                for override in list(root):
                    if override.get("PartName", "").startswith("/ppt/fonts/"):
                        root.remove(override)
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)

            dst.writestr(item, data)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=pathlib.Path)
    parser.add_argument("output", type=pathlib.Path)
    args = parser.parse_args()
    sanitize(args.source, args.output)


if __name__ == "__main__":
    main()
