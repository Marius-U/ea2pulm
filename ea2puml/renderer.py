from __future__ import annotations
from pathlib import Path
from typing import Optional, Iterable


class PlantUMLWriter:
    """
    Pure string builder. Only `save()` touches the filesystem.
    Keep semantics identical to monolith: where it wrote, how it wrapped, etc.
    """
    def __init__(self, skin: Optional[str] = None, autolayout: Optional[str] = None):
        self._buf: list[str] = []
        self._skin = skin
        self._autolayout = autolayout
        self._started = False
        self._ended = False

    def start(self, title: Optional[str] = None) -> None:
        if self._started:
            return
        self._started = True
        self._buf.append("@startuml")
        if self._skin:
            self._buf.append(self._skin)
        if self._autolayout:
            self._buf.append(self._autolayout)
        if title:
            from .utils import puml_escape_inline
            self._buf.append(f'title "{puml_escape_inline(title)}"')

    def writeln(self, line: str = "") -> None:
        self._buf.append(line)

    def extend(self, lines: Iterable[str]) -> None:
        self._buf.extend(lines)

    def end(self) -> None:
        if self._ended:
            return
        self._ended = True
        self._buf.append("@enduml")

    def text(self) -> str:
        return "\n".join(self._buf) + ("\n" if self._buf and not self._buf[-1].endswith("\n") else "")

    def save(self, outdir: Path, filename: str) -> Path:
        outdir.mkdir(parents=True, exist_ok=True)
        path = outdir / f"{filename}.puml"
        with path.open("w", encoding="utf-8") as f:
            f.write(self.text())
        return path
