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

@register("Sequence")
class SequenceDiagramHandler:
    @staticmethod
    def render(diagram: Diagram, out: PlantUMLWriter, cfg) -> None:
        idx = _index(diagram)

        # Participants
        for el in diagram.elements:
            name = puml_escape_inline(el.name)
            if (el.stereotype or "").lower() == "actor" or el.type.lower() == "actor":
                if cfg.alias_mode == "name":
                    out.writeln(f'actor "{name}"')
                else:
                    out.writeln(f'actor "{name}" as {el.alias}')
            else:
                if cfg.alias_mode == "name":
                    out.writeln(f'participant "{name}"')
                else:
                    out.writeln(f'participant "{name}" as {el.alias}')
        out.writeln("")

        # Messages
        for c in diagram.connectors:
            src_el = idx.get(c.source.element_id)
            dst_el = idx.get(c.target.element_id)
            if not src_el or not dst_el:
                continue
            src = _ref(src_el, cfg.alias_mode)
            dst = _ref(dst_el, cfg.alias_mode)
            label = c.labels.get("name") or ""
            # (If you distinguish sync/async by c.type in monolith, apply here.)
            out.writeln(f"{src} -> {dst}" + (f" : {label}" if label else ""))
        out.writeln("")
