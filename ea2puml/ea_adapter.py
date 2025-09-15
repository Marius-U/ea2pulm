# ea2puml/ea_adapter.py
from __future__ import annotations
from typing import Dict, Optional

try:
    import win32com.client  # type: ignore
except Exception:
    win32com = None  # type: ignore

from .config import Config
from .models import Diagram, Element, Connector, Note, Geometry, ConnectorEnd
from .utils import AliasFactory

class EAAdapter:
    """EA COM adapter → neutral models. All pywin32 lives here."""
    def __init__(self, cfg: Config) -> None:
        self._repo = None
        self._cfg = cfg
        self._alias_factory = AliasFactory(cfg.alias_mode or "human")

    def _ensure_repo(self):
        if self._repo is not None:
            return
        if win32com is None:
            raise RuntimeError("pywin32 is required. Install with: pip install pywin32")
        app = win32com.client.Dispatch("EA.App")
        self._repo = app.Repository

    def get_selected_diagram(self) -> Diagram:
        self._ensure_repo()
        repo = self._repo
        dia = repo.GetCurrentDiagram()
        if dia is None:
            raise RuntimeError("No current diagram selected in EA.")

        diagram = Diagram(id=dia.DiagramID, name=dia.Name, type=dia.Type)

        elements_by_id: Dict[int, Element] = {}
        for d_obj in dia.DiagramObjects:
            el = repo.GetElementByID(d_obj.ElementID)
            inst_guid = d_obj.InstanceGUID
            alias = self._alias_factory.make(inst_guid, el.Name or "")

            geom = self._parse_geometry(d_obj)
            elem = Element(
                id=el.ElementID,
                guid=str(inst_guid),
                name=el.Name or "",
                type=el.Type or "",
                stereotype=getattr(el, "Stereotype", None) or None,
                color=None,  # fill if your monolith emitted background color from d_obj
                tags=self._collect_tags(el),
                geometry=geom,
                alias=alias if (self._cfg.alias_mode != "name") else "",  # name-mode uses quoted names
            )
            elements_by_id[elem.id] = elem
            diagram.elements.append(elem)

            # Notes as elements? keep here or in a separate pass if needed

        # Notes (optional): if your monolith treated notes specially, mirror it
        for d_obj in dia.DiagramObjects:
            el = repo.GetElementByID(d_obj.ElementID)
            if (el.Type or "").lower() in ("note", "text"):
                geom = self._parse_geometry(d_obj)
                diagram.notes.append(
                    Note(
                        id=el.ElementID,
                        text=el.Notes or el.Name or "",
                        geometry=geom,
                        tags=self._collect_tags(el),
                    )
                )

        for d_link in dia.DiagramLinks:
            conn = repo.GetConnectorByID(d_link.ConnectorID)
            c = Connector(
                id=conn.ConnectorID,
                type=conn.Type or "",
                stereotype=getattr(conn, "Stereotype", None) or None,
                color=None,  # fill if your monolith emitted connector color
                source=ConnectorEnd(element_id=int(conn.ClientID)),
                target=ConnectorEnd(element_id=int(conn.SupplierID)),
                labels={"name": getattr(conn, "Name", "") or ""},
                geometry=self._parse_link_geometry(d_link),
            )
            diagram.connectors.append(c)

        # Package membership (copy your exact algorithm here if needed)
        diagram.packages = {}

        return diagram

    # ---------- helpers ----------
    def _parse_geometry(self, d_obj) -> Geometry:
        left = getattr(d_obj, "left", None) or getattr(d_obj, "Left", 0)
        right = getattr(d_obj, "right", None) or getattr(d_obj, "Right", 0)
        top = getattr(d_obj, "top", None) or getattr(d_obj, "Top", 0)
        bottom = getattr(d_obj, "bottom", None) or getattr(d_obj, "Bottom", 0)
        return Geometry(left=left, right=right, top=top, bottom=bottom)

    def _parse_link_geometry(self, d_link) -> Optional[Geometry]:
        return None

    def _collect_tags(self, el) -> Dict[str, str]:
        tags: Dict[str, str] = {}
        try:
            for tv in el.TaggedValues:
                name = getattr(tv, "Name", None)
                val = getattr(tv, "Value", None) or getattr(tv, "Notes", None)
                if name:
                    tags[str(name)] = "" if val is None else str(val)
        except Exception:
            pass
        return tags
