from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Optional
from pathlib import Path


@dataclass
class Config:
    # Mirrors existing flags/semantics (names kept identical)
    outdir: Path                         # -o / --outdir
    filename: Optional[str]              # -f / --filename
    include_tags: List[str]              # --include-tags (comma- or multi-use)
    no_colors: bool                      # --no-colors
    element_stereo: Optional[str]        # --element-stereo
    edge_labels: Optional[str]           # --edge-labels
    direction: Optional[str]             # --direction
    explore: bool                        # --explore
    skin: Optional[str]                  # --skin
    autolayout: Optional[str]            # --autolayout (string to allow passthrough)
    alias_mode: Optional[str]            # --alias-mode  (uuid|human|name)
    interface_style: Optional[str]       # --interface-style (e.g., lollipop/box)

    # Internal passthrough of original argv if you need it
    _raw_args: List[str] = field(default_factory=list, repr=False)
