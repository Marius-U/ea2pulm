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

@register("Use Case")
class UseCaseDiagramHandler:
    @staticmethod
    def render(diagram: Diagram, out: PlantUMLWriter, cfg) -> None:
        for el in diagram.elements:
            alias = f" as {el.alias}" if (el.alias and cfg.alias_mode != "name") else ""
            if (el.stereotype or "").lower() == "actor" or el.type.lower() == "actor":
                out.writeln(f'actor "{puml_escape_inline(el.name)}"{alias}')
            else:
                out.writeln(f'usecase "{puml_escape_inline(el.name)}"{alias}')

        idx = _index(diagram)
        for c in diagram.connectors:
            src_el = idx.get(c.source.element_id)
            dst_el = idx.get(c.target.element_id)
            if not src_el or not dst_el:
                continue
            src = _ref(src_el, cfg.alias_mode)
            dst = _ref(dst_el, cfg.alias_mode)
            label = c.labels.get("name") or ""
            out.writeln(f"{src} --> {dst}" + (f" : {label}" if label else ""))
