from __future__ import annotations
import argparse
from pathlib import Path
from typing import List

from .config import Config
from .main import run


def _csv_or_multi(values: List[str]) -> List[str]:
    out: List[str] = []
    for v in values:
        if "," in v:
            out.extend(x.strip() for x in v.split(",") if x.strip())
        else:
            s = v.strip()
            if s:
                out.append(s)
    return out


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="ea2puml",
        description="EA → PlantUML exporter (refactored, behavior-compatible)",
    )
    # CLI compatibility: keep identical flag names & meanings.
    p.add_argument("-o", "--outdir", type=Path, required=True,
                   help="Output directory for generated .puml files.")
    p.add_argument("-f", "--filename", type=str, default=None,
                   help="Output filename (without extension). If omitted, diagram name is used.")
    p.add_argument("--include-tags", action="append", default=[],
                   metavar="TAG[,TAG...]", help="Filter/include by EA tagged values (repeat or CSV).")
    p.add_argument("--no-colors", action="store_true",
                   help="Disable diagram/element color output.")
    p.add_argument("--element-stereo", type=str, default=None,
                   help="Element stereotype label policy (passes through unchanged).")
    p.add_argument("--edge-labels", type=str, default=None,
                   help="Edge label policy (e.g., name|stereotype|both).")
    p.add_argument("--direction", type=str, default=None,
                   help="Layout direction hint (e.g., 'left to right direction').")
    p.add_argument("--explore", action="store_true",
                   help="Exploration mode (kept as in monolith).")
    p.add_argument("--skin", type=str, default=None,
                   help="PlantUML skin/theme passthrough.")
    p.add_argument("--autolayout", type=str, default=None,
                   help="PlantUML autolayout control (string passthrough).")
    p.add_argument("--alias-mode", type=str, default=None,
                   help="Alias mode (uuid|human|name) – unchanged semantics.")
    p.add_argument("--interface-style", type=str, default=None,
                   help="Interface rendering style (e.g., lollipop or box).")
    return p


def main() -> None:
    parser = build_parser()
    ns = parser.parse_args()
    cfg = Config(
        outdir=ns.outdir,
        filename=ns.filename,
        include_tags=_csv_or_multi(ns.include_tags),
        no_colors=ns.no_colors,
        element_stereo=ns.element_stereo,
        edge_labels=ns.edge_labels,
        direction=ns.direction,
        explore=ns.explore,
        skin=ns.skin,
        autolayout=ns.autolayout,
        alias_mode=ns.alias_mode,
        interface_style=ns.interface_style,
        _raw_args=[],
    )
    run(cfg)
