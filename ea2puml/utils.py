# ea2puml/utils.py
from __future__ import annotations
import os, re
from typing import Dict

def puml_escape_inline(text: str) -> str:
    if text is None:
        return ""
    return str(text).replace("\\", "\\\\").replace('"', r"\"")

def sanitize_alias(guid: str) -> str:
    if not guid:
        return "E_" + str(abs(hash(os.urandom(4))))
    g = re.sub(r"[^A-Za-z0-9]", "", guid)
    if not g:
        g = "E" + str(abs(hash(guid)) % (10**8))
    return "E_" + g[:16]

def slugify_name(name: str) -> str:
    s = re.sub(r"[^A-Za-z0-9_]+", "_", (name or "").strip())
    if not s or s[0].isdigit():
        s = "E_" + s
    return s[:40] if len(s) > 40 else s

class AliasFactory:
    """Matches the monolith behavior: uuid | human | name."""
    def __init__(self, mode: str):
        self.mode = mode
        self.counts: Dict[str, int] = {}
    def make(self, instance_guid: str, name: str) -> str:
        if self.mode == "uuid":
            return sanitize_alias(instance_guid)
        if self.mode == "human":
            base = slugify_name(name) or "E"
            n = self.counts.get(base, 0)
            self.counts[base] = n + 1
            return base if n == 0 else f"{base}_{n+1}"
        return ""  # name mode → no alias (use quoted name everywhere)
