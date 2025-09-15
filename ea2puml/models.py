from __future__ import annotations
from dataclasses import dataclass, field
from typing import Dict, List, Optional

Id = int  # EA model ElementID

@dataclass
class Geometry:
    left: int
    right: int
    top: int
    bottom: int

@dataclass
class Element:
    id: Id                  # EA model ElementID
    guid: str               # DiagramObject.InstanceGUID (for stable aliasing)
    name: str
    type: str               # EA.Element.Type
    stereotype: Optional[str]
    color: Optional[str]
    tags: Dict[str, str] = field(default_factory=dict)
    geometry: Optional[Geometry] = None
    alias: Optional[str] = None   # computed per alias mode ("" when alias_mode=name)

@dataclass
class Note:
    id: Id
    text: str
    geometry: Optional[Geometry] = None
    tags: Dict[str, str] = field(default_factory=dict)

@dataclass
class ConnectorEnd:
    element_id: Id
    role: Optional[str] = None
    label: Optional[str] = None

@dataclass
class Connector:
    id: Id
    type: str
    stereotype: Optional[str]
    color: Optional[str]
    source: ConnectorEnd
    target: ConnectorEnd
    labels: Dict[str, str] = field(default_factory=dict)
    geometry: Optional[Geometry] = None

@dataclass
class Diagram:
    id: Id
    name: str
    type: str
    elements: List[Element] = field(default_factory=list)
    connectors: List[Connector] = field(default_factory=list)
    notes: List[Note] = field(default_factory=list)
    packages: Dict[Id, List[Id]] = field(default_factory=dict)
    direction: Optional[str] = None
