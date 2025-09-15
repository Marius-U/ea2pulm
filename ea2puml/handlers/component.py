from __future__ import annotations
from typing import Dict

from ..handler_registry import register
from ..models import Diagram, Element
from ..renderer import PlantUMLWriter
from ..utils import puml_escape_inline

def _index(diagram: Diagram) -> Dict[int, Element]:
    return {e.id: e for e in diagram.elements}

def _ref(el: Element, alias_mode: str) -> str:
    if alias_mode == "name":
        return f"\"{puml_escape_inline(el.name)}\""
    return el.alias or f"\"{puml_escape_inline(el.name)}\""

@register("Component")
class ComponentDiagramHandler:
    @staticmethod
    def render(diagram: Diagram, out: PlantUMLWriter, cfg) -> None:
        if cfg.direction:
            out.writeln(cfg.direction)

        # Elements (simple version; keep your richer styles if you already ported them)
        for el in diagram.elements:
            stereo = f" <<{el.stereotype}>>" if (el.stereotype and cfg.element_stereo) else ""
            color = ""  # fill when you port color logic
            alias = f" as {el.alias}" if (el.alias and cfg.alias_mode != "name") else ""
            out.writeln(f'component "{puml_escape_inline(el.name)}"{stereo}{alias}{color}')

        # Connectors
        idx = _index(diagram)
        for c in diagram.connectors:
            src_el = idx.get(c.source.element_id)
            dst_el = idx.get(c.target.element_id)
            if not src_el or not dst_el:
                continue
            src = _ref(src_el, cfg.alias_mode)
            dst = _ref(dst_el, cfg.alias_mode)

            # label policy (simplified; extend to match your monolith exactly)
            label = ""
            parts = []
            if cfg.edge_labels in ("name", "both"):
                nm = c.labels.get("name")
                if nm:
                    parts.append(nm)
            if cfg.edge_labels in ("stereotype", "both") and c.stereotype:
                parts.append(f"<<{c.stereotype}>>")
            if parts:
                label = " : " + " / ".join(parts)

            out.writeln(f"{src} --> {dst}{label}")

        # Notes
        for n in diagram.notes:
            out.writeln(f'note "{n.text}" as N{n.id}')
