from __future__ import annotations
from typing import Callable, Dict

# Handler protocol (duck-typed): class with
#   render(diagram, out, cfg) -> None

_REGISTRY: Dict[str, type] = {}


def register(ea_diagram_type: str) -> Callable[[type], type]:
    def deco(cls: type) -> type:
        _REGISTRY[ea_diagram_type.strip().lower()] = cls
        return cls
    return deco


def resolve(diagram_type: str) -> type:
    key = (diagram_type or "").strip().lower()
    try:
        return _REGISTRY[key]
    except KeyError:
        raise KeyError(f"No handler registered for diagram type: {diagram_type!r}")


def registered_types() -> Dict[str, type]:
    return dict(_REGISTRY)
