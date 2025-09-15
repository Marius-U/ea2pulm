# ea2puml/main.py
from __future__ import annotations
import re
from pathlib import Path

from .config import Config
from .ea_adapter import EAAdapter
from .handler_registry import resolve
from .handlers import component, sequence, usecase  # noqa: F401
from .renderer import PlantUMLWriter

_SANITIZE = re.compile(r"[^A-Za-z0-9._-]+")

def _sanitize_filename(name: str) -> str:
    name = name.strip()
    name = _SANITIZE.sub("_", name)
    return name or "diagram"

def run(cfg: Config) -> Path:
    adapter = EAAdapter(cfg)  # <<< was EAAdapter()
    diagram = adapter.get_selected_diagram()

    if cfg.direction:
        diagram.direction = cfg.direction

    Handler = resolve(diagram.type)

    out = PlantUMLWriter(skin=cfg.skin, autolayout=cfg.autolayout)
    out.start(diagram.name)
    Handler.render(diagram, out, cfg)
    out.end()

    fname = cfg.filename or _sanitize_filename(diagram.name)
    return out.save(cfg.outdir, fname)
