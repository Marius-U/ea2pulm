#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# ============================================================
# EA → PlantUML Exporter (Packages + Lollipop + Block Notes, v2)
# ------------------------------------------------------------
# - Groups elements inside package blocks using on-diagram geometry:
#   center-in-rect OR area-overlap >= 50%
# - Preserves package colors
# - Lollipop interfaces by default (configurable)
# - Block notes with real newlines; tagged values from Value or Notes
# - Alias modes: uuid | human (default) | name
# - Edge labels default: both (name + stereotype) to match EA figure
# - Packages ordered by top-to-bottom, then left-to-right (closer to EA)
# - Robust connector endpoint resolution (InstanceUID fallback via ElementID)
# - Default output folder: ./output
# ============================================================

import argparse
import os
import re
from typing import Dict, Tuple, Optional, List

try:
    import win32com.client.dynamic  # pywin32
except Exception:
    print("ERROR: pywin32 is required. Install with: pip install pywin32")
    raise

# ---------------- Helpers ----------------

def puml_escape_inline(text: str) -> str:
    if text is None:
        return ""
    return str(text).replace("\\", "\\\\").replace('"', r'\"')

def sanitize_alias(guid: str) -> str:
    if not guid:
        return "E_" + str(abs(hash(os.urandom(4))))
    g = re.sub(r"[^A-Za-z0-9]", "", guid)
    if not g:
        g = "E" + str(abs(hash(guid)) % (10**8))
    return "E_" + g[:16]

def ea_color_long_to_hex(ea_long: int) -> Optional[str]:
    if ea_long is None or ea_long == -1:
        return None
    hx = hex(ea_long).replace("0x", "")
    hx = ("0" * (6 - len(hx))) + hx[-6:]
    b, g, r = hx[0:2], hx[2:4], hx[4:6]
    return f"#{r}{g}{b}".upper()

def slugify_name(name: str) -> str:
    s = re.sub(r"[^A-Za-z0-9_]+", "_", (name or "").strip())
    if not s or s[0].isdigit():
        s = "E_" + s
    return s[:40] if len(s) > 40 else s

class AliasFactory:
    def __init__(self, mode: str):
        self.mode = mode  # uuid | human | name
        self.counts: Dict[str, int] = {}
    def make(self, guid: str, name: str) -> str:
        if self.mode == "uuid":
            return sanitize_alias(guid)
        if self.mode == "human":
            base = slugify_name(name) or "E"
            n = self.counts.get(base, 0)
            self.counts[base] = n + 1
            return base if n == 0 else f"{base}_{n+1}"
        return ""  # name mode → no alias

# ---------------- Mappings ----------------

ELEMENT_TYPE_TO_PUML = {
    "Class": "class",
    "Interface": "interface",     # (renderer can switch to lollipop)
    "Enumeration": "enum",
    "Component": "component",
    "Node": "node",
    "Artifact": "artifact",
    "Actor": "actor",
    "UseCase": "usecase",
    "Package": "package",
    "State": "state",
    "Device": "node",
    "Text": "rectangle",
    "Note": "note",
    "Boundary": "rectangle",
    "Activity": "rectangle",
    "Requirement": "rectangle",
    "Environment": "frame",
}

def relation_for_type(conn_type: str) -> Tuple[str, str]:
    t = (conn_type or "").lower()
    if t in ("association",):
        return ("--", ">")
    if t in ("dependency",):
        return ("..", ">")
    if t in ("realization", "realisation"):
        return ("..", "|>")
    if t in ("generalization", "inheritance"):
        return ("--", "|>")
    if t in ("aggregation",):
        return ("o-", ">")
    if t in ("composition",):
        return ("*-", ">")
    if t in ("informationflow", "information flow"):
        return ("..", ">")
    if t in ("controlflow", "control flow", "flow"):
        return ("-", "->")
    return ("--", ">")

# ---------------- Exporter ----------------

class EAPlantUMLExporter:
    def __init__(self,
                 include_tags: bool = False,
                 include_colors: bool = True,
                 element_stereo: str = "off",       # off | on | inname
                 edge_labels: str = "both",         # stereotype | name | both | none
                 direction: Optional[str] = None,
                 skin: bool = False,
                 autolayout: bool = False,
                 explore: bool = False,
                 alias_mode: str = "human",
                 interface_style: str = "lollipop"  # lollipop | class
                 ):
        self.include_tags = include_tags
        self.include_colors = include_colors
        self.element_stereo = element_stereo
        self.edge_labels = edge_labels
        self.direction = direction
        self.skin = skin
        self.autolayout = autolayout
        self.explore = explore
        self.alias_mode = alias_mode
        self.interface_style = interface_style

        self.alias_factory = AliasFactory(self.alias_mode)
        self.eapp = None
        self.repo = None

        # Collected data
        self.elements: Dict[str, dict] = {}     # by DiagramObject.InstanceGUID
        self.connectors: Dict[str, dict] = {}   # by DiagramLink.ConnectorID
        self.alias_by_guid: Dict[str, str] = {}
        self.packages: Dict[str, dict] = {}     # pkg rect + color
        self.members_in_pkg: Dict[str, List[str]] = {}  # pkg_guid -> [element guids]
        self.elemid_to_instuid: Dict[int, str] = {}     # model ElementID -> InstanceUID on diagram

    # -------- COM ----------
    def connect_ea(self):
        print("Connecting to Enterprise Architect (EA)...")
        self.eapp = win32com.client.dynamic.Dispatch("EA.App")
        self.repo = self.eapp.Repository
        print("Connected to EA repository.")

    def get_selected_diagram(self):
        itemType, item = self.repo.GetTreeSelectedItem()
        if itemType != 8:
            raise RuntimeError("Please select a Diagram in EA's Project Browser before running.")
        return item

    # -------- Geometry helpers ----------
    @staticmethod
    def rect(dobj, cornerX=0, cornerY=0):
        x = dobj.Left + cornerX
        y = -dobj.Top + cornerY
        w = dobj.Right - dobj.Left
        h = -(dobj.Bottom - dobj.Top)
        return {"x": x, "y": y, "w": w, "h": h}

    @staticmethod
    def center(rect: dict) -> Tuple[int, int]:
        return (rect["x"] + rect["w"] // 2, rect["y"] + rect["h"] // 2)

    @staticmethod
    def area(rect: dict) -> int:
        return max(0, rect["w"]) * max(0, rect["h"])

    @staticmethod
    def overlap_area(a: dict, b: dict) -> int:
        x1 = max(a["x"], b["x"])
        y1 = max(a["y"], b["y"])
        x2 = min(a["x"] + a["w"], b["x"] + b["w"])
        y2 = min(a["y"] + a["h"], b["y"] + b["h"])
        if x2 <= x1 or y2 <= y1:
            return 0
        return (x2 - x1) * (y2 - y1)

    @staticmethod
    def center_inside(inner_rect: dict, outer_rect: dict) -> bool:
        cx, cy = EAPlantUMLExporter.center(inner_rect)
        return (cx >= outer_rect["x"] and
                cy >= outer_rect["y"] and
                cx <= outer_rect["x"] + outer_rect["w"] and
                cy <= outer_rect["y"] + outer_rect["h"])

    # -------- Endpoint resolution ----------
    def resolve_endpoint_uid(self, dlnk, dconn, which: str) -> Optional[str]:
        """Return a diagram InstanceUID for connector endpoint ('source' or 'target')."""
        uid = dlnk.SourceInstanceUID if which == "source" else dlnk.TargetInstanceUID
        if uid and uid in self.elements:
            return uid
        elem_id = int(dconn.ClientID) if which == "source" else int(dconn.SupplierID)
        return self.elemid_to_instuid.get(elem_id)

    # -------- Gather ----------
    def gather(self, diagram, cornerX=0, cornerY=0):
        # First pass: collect elements with geometry
        for dObj in diagram.DiagramObjects:
            dElem = self.repo.GetElementByID(dObj.ElementID)
            inst_guid = dObj.InstanceGUID
            # index model ElementID -> this diagram InstanceUID (first seen)
            self.elemid_to_instuid.setdefault(int(dObj.ElementID), inst_guid)
            alias = self.alias_factory.make(inst_guid, dElem.Name or "")
            self.alias_by_guid[inst_guid] = alias

            rect = self.rect(dObj, cornerX, cornerY)
            bg = ea_color_long_to_hex(getattr(dObj, "BackgroundColor", -1)) if self.include_colors else None

            el = {
                "guid": inst_guid,
                "alias": alias,
                "name": dElem.Name or "",
                "type": dElem.Type or "",
                "stereotype": dElem.Stereotype or "",
                "notes": dElem.Notes or "",
                "show_tags": bool(getattr(dObj, "ShowTags", False)),
                "bg_hex": bg,
                "raw": dElem,
                "dobj": dObj,
                "rect": rect,
            }
            self.elements[inst_guid] = el

            if el["type"] == "Package":
                self.packages[inst_guid] = {
                    "guid": inst_guid,
                    "name": el["name"],
                    "alias": alias,
                    "color": bg,
                    "rect": rect,
                }

            if self.explore:
                print(f"ELEMENT: {el['type']} '{el['name']}' alias={alias or '(name)'} bg={bg} rect={rect}")

            if el["type"] == "UMLDiagram":
                diagID = dElem.MiscData(0)
                sub = self.repo.GetDiagramByID(diagID)
                self.gather(sub, dObj.Left, -dObj.Top)

        # Second pass: geometry-based package membership
        for pkg_guid, pkg in self.packages.items():
            members: List[str] = []
            preg = pkg["rect"]
            for guid, el in self.elements.items():
                if guid == pkg_guid:
                    continue
                erect = el["rect"]
                # center-in-rect OR overlap >= 50% of element area
                ov = self.overlap_area(erect, preg)
                inside = self.center_inside(erect, preg) or (ov >= 0.5 * self.area(erect))
                if inside:
                    members.append(guid)
            self.members_in_pkg[pkg_guid] = members

        # Connectors
        for dLnk in diagram.DiagramLinks:
            if dLnk.IsHidden:
                continue
            dConn = self.repo.GetConnectorByID(dLnk.ConnectorID)

            source_uid = self.resolve_endpoint_uid(dLnk, dConn, "source")
            target_uid = self.resolve_endpoint_uid(dLnk, dConn, "target")
            if not source_uid or not target_uid:
                # endpoint not on this diagram (or hidden) -> skip
                continue

            conn = {
                "id": str(dLnk.ConnectorID),
                "name": dConn.Name or "",
                "type": dConn.Type or "",
                "stereotype": dConn.Stereotype or "",
                "direction": dConn.Direction or "Unspecified",
                "source_uid": source_uid,
                "target_uid": target_uid,
                "line_hex": ea_color_long_to_hex(getattr(dLnk, "LineColor", -1)) if self.include_colors else None,
                "line_width": getattr(dLnk, "LineWidth", 1),
                "hidden_labels": bool(getattr(dLnk, "HiddenLabels", False)),
                "raw": dConn,
                "dlnk": dLnk,
            }
            self.connectors[conn["id"]] = conn

            if self.explore:
                print(f"CONNECTOR: {conn['type']} '{conn['name']}' <<{conn['stereotype']}>> "
                      f"{conn['source_uid']} -> {conn['target_uid']} dir={conn['direction']} color={conn['line_hex']}")

    # -------- Render ----------
    def render_header(self, diagram) -> str:
        lines = []
        lines.append("@startuml")
        lines.append(f'title "{puml_escape_inline(diagram.Name)}"')
        if self.direction == "LR":
            lines.append("left to right direction")
        elif self.direction == "TB":
            lines.append("' top to bottom direction (PlantUML default)")
        if self.skin:
            lines.extend([
                "",
                "' --- optional skinparams ---",
                "skinparam defaultFontName Arial",
                "skinparam shadowing false",
                "skinparam classAttributeIconSize 0",
                "skinparam wrapWidth 200",
                "skinparam dpi 144",
            ])
        if self.autolayout:
            lines.append("")
            lines.append("autolayout")
        lines.append("")
        return "\n".join(lines)

    def render_footer(self) -> str:
        return "\n@enduml\n"

    def ref_token(self, el: dict) -> str:
        if self.alias_mode == "name":
            return f"\"{puml_escape_inline(el['name'])}\""
        return el["alias"]

    def element_decl(self, el: dict) -> str:
        name = el["name"] or ""
        st = el["stereotype"] or ""
        etype = el["type"] or ""
        alias = el["alias"]
        bg = el["bg_hex"]
        keyword = ELEMENT_TYPE_TO_PUML.get(etype, "rectangle")

        # Block note elements (Note/Text)
        if etype in ("Note", "Text"):
            txt = (el["notes"] or name)
            lines = [f"note as {alias or (slugify_name(name) or 'Note_1')}"]
            if txt:
                lines.extend(txt.splitlines())
            lines.append("end note")
            return "\n".join(lines)

        # Element stereotypes
        name_disp = name
        stereo_suffix = ""
        if self.element_stereo == "on" and st:
            stereo_suffix = f" <<{st}>>"
        elif self.element_stereo == "inname" and st:
            name_disp = f"{name} <<{st}>>"

        disp = f'"{puml_escape_inline(name_disp)}"'
        color = f' {bg}' if bg else ""

        # Interface style (lollipop)
        if etype == "Interface" and self.interface_style == "lollipop":
            if self.alias_mode == "name":
                return f'() {disp}{color}'
            else:
                return f'() {disp} as {alias}{color}'

        # Default element
        if self.alias_mode == "name":
            return f'{keyword}{stereo_suffix} {disp}{color}'
        else:
            return f'{keyword}{stereo_suffix} {disp} as {alias}{color}'

    def tag_block_for_element(self, el: dict) -> Optional[str]:
        try:
            tvs = el["raw"].TaggedValues
        except Exception:
            return None
        if not tvs or tvs.Count == 0:
            return None
        if not (self.include_tags or el.get("show_tags")):
            return None

        lines = []
        for tv in tvs:
            val = getattr(tv, "Value", None)
            if val is None or str(val).strip() == "":
                notes = getattr(tv, "Notes", None)
                if notes and str(notes).strip() != "":
                    val = notes
            if val is not None and str(val).strip() != "":
                lines.append(f"{tv.Name}={val}")

        if not lines:
            return None

        ref = self.ref_token(el)
        return f"note right of {ref}\n" + "\n".join(lines) + "\nend note"

    def render_elements_grouped(self) -> str:
        """Emit packages (ordered TL->BR) with members inside; then leftovers."""
        emitted: set = set()
        out: List[str] = []

        # sort packages by Y, then X (top-to-bottom then left-to-right)
        def pkg_key(p):
            r = p["rect"]
            return (r["y"], r["x"])
        sorted_pkgs = sorted(self.packages.values(), key=pkg_key)

        for pkg in sorted_pkgs:
            name = pkg["name"]
            color = pkg["color"] or ""
            hdr = f'package "{puml_escape_inline(name)}"'
            hdr += f" {color}" if color else ""
            hdr += " {"
            out.append(hdr)

            for guid in self.members_in_pkg.get(pkg["guid"], []):
                el = self.elements[guid]
                if el["type"] == "Package":
                    continue
                out.append(self.element_decl(el))
                tagb = self.tag_block_for_element(el)
                if tagb:
                    out.append(tagb)
                emitted.add(guid)

            out.append("}")
            out.append("")

        # leftovers (not contained by any package region)
        for guid, el in self.elements.items():
            if el["type"] == "Package" or guid in emitted:
                continue
            out.append(self.element_decl(el))
            tagb = self.tag_block_for_element(el)
            if tagb:
                out.append(tagb)

        out.append("")
        return "\n".join(out)

    def render_connectors(self) -> str:
        out = []
        for cid, c in self.connectors.items():
            src_el = self.elements.get(c["source_uid"])
            dst_el = self.elements.get(c["target_uid"])
            if not src_el or not dst_el:
                continue

            src = self.ref_token(src_el)
            dst = self.ref_token(dst_el)

            style, head = relation_for_type(c["type"])
            if head == "|>":
                rel = style + head
            elif head == "->":
                rel = style + "->"
            else:
                if "Source" in c["direction"] or "Bi-Directional" in c["direction"]:
                    rel = style + head
                else:
                    rel = style + "-"

            if self.include_colors and c.get("line_hex"):
                if rel.startswith("--"):
                    rel = f'-[{c["line_hex"]}]-' + rel[2:]
                elif rel.startswith(".."):
                    rel = f'.[{c["line_hex"]}].' + rel[2:]
                elif rel.startswith("*-"):
                    rel = f'*[{c["line_hex"]}]-' + rel[2:]
                elif rel.startswith("o-"):
                    rel = f'o[{c["line_hex"]}]-' + rel[2:]
                elif rel.startswith("-"):
                    rel = f'-[{c["line_hex"]}]' + rel[1:]

            parts = []
            if self.edge_labels in ("name", "both") and (not c["hidden_labels"]) and c["name"]:
                parts.append(c["name"])
            if self.edge_labels in ("stereotype", "both") and c["stereotype"]:
                parts.append(f'<<{c["stereotype"]}>>')
            label = " : " + " ".join(parts) if parts else ""

            out.append(f"{src} {rel} {dst}{label}")
        out.append("")
        return "\n".join(out)

    # -------- Sequence (basic) --------
    def is_sequence_diagram(self, diagram) -> bool:
        try:
            return (diagram.Type or "").lower() == "sequence"
        except Exception:
            return False

    def render_sequence(self, diagram) -> str:
        out = []
        for guid, el in self.elements.items():
            et = (el["type"] or "").lower()
            name = puml_escape_inline(el["name"])
            ref = self.ref_token(el)
            if et in ("actor",):
                if self.alias_mode == "name":
                    out.append(f'actor "{name}"')
                else:
                    out.append(f'actor "{name}" as {ref}')
            else:
                if self.alias_mode == "name":
                    out.append(f'participant "{name}"')
                else:
                    out.append(f'participant "{name}" as {ref}')
        out.append("")
        for cid, c in self.connectors.items():
            src_el = self.elements.get(c["source_uid"])
            dst_el = self.elements.get(c["target_uid"])
            if not src_el or not dst_el:
                continue
            t = (c["type"] or "").lower()
            if t in ("message", "sequence"):
                src = self.ref_token(src_el)
                dst = self.ref_token(dst_el)
                parts = []
                if self.edge_labels in ("name", "both") and c["name"]:
                    parts.append(c["name"])
                if self.edge_labels in ("stereotype", "both") and c["stereotype"]:
                    parts.append(f'<<{c["stereotype"]}>>')
                label = " : " + " ".join(parts) if parts else ""
                out.append(f"{src} -> {dst}{label}")
        out.append("")
        return "\n".join(out)

    # -------- Main --------
    def export(self, outdir: str, filename: Optional[str] = None):
        self.connect_ea()
        diagram = self.get_selected_diagram()
        self.gather(diagram)

        lines = [self.render_header(diagram)]
        if self.is_sequence_diagram(diagram):
            lines.append(self.render_sequence(diagram))
        else:
            lines.append(self.render_elements_grouped())
            lines.append(self.render_connectors())
        lines.append(self.render_footer())

        puml = "\n".join(lines)
        base = re.sub(r"[^A-Za-z0-9._-]+", "_", filename or (diagram.Name or "diagram"))
        os.makedirs(outdir, exist_ok=True)
        path = os.path.join(outdir, base + ".puml")
        with open(path, "w", encoding="utf-8") as f:
            f.write(puml)
        print(f"Exported PlantUML to: {path}")

# ---------------- CLI ----------------

def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Export the currently selected EA diagram to PlantUML (.puml)")
    p.add_argument("-o", "--outdir", default="output", help="Output directory (default: output)")
    p.add_argument("-f", "--filename", default=None, help="Override output file name (without extension)")
    p.add_argument("-t", "--include-tags", action="store_true", default=False, help="Include Tagged Values as block notes")
    p.add_argument("-c", "--no-colors", action="store_true", default=False, help="Disable colors on elements/connectors")
    p.add_argument("-s", "--element-stereo", choices=["off", "on", "inname"], default="off",
                   help="Element stereotypes: off (default), on (keyword suffix), inname (append to name)")
    p.add_argument("--edge-labels", choices=["stereotype", "name", "both", "none"], default="both",
                   help="Connector label policy (default: both)")
    p.add_argument("-d", "--direction", choices=["LR", "TB"], default=None, help="Layout direction (LR/TB)")
    p.add_argument("-e", "--explore", action="store_true", default=False, help="Print diagnostic listing")
    p.add_argument("--skin", action="store_true", default=False, help="Include minimal skinparam block")
    p.add_argument("--autolayout", action="store_true", default=False, help="Add 'autolayout' directive")
    p.add_argument("--alias-mode", choices=["uuid", "human", "name"], default="human",
                   help="Identifier mode: uuid (stable), human (from names, deduped), name (no alias; quoted names).")
    p.add_argument("--interface-style", choices=["lollipop", "class"], default="lollipop",
                   help="Interface rendering style: lollipop (circle) or class-like box (default: lollipop)")
    return p.parse_args(argv)

def main(argv=None):
    args = parse_args(argv)
    exporter = EAPlantUMLExporter(
        include_tags=args.include_tags,
        include_colors=not args.no_colors,
        element_stereo=args.element_stereo,
        edge_labels=args.edge_labels,
        direction=args.direction,
        skin=args.skin,
        autolayout=args.autolayout,
        explore=args.explore,
        alias_mode=args.alias_mode,
        interface_style=args.interface_style,
    )
    exporter.export(outdir=args.outdir, filename=args.filename)

if __name__ == "__main__":
    main()
