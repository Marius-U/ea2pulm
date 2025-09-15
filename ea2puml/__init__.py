"""
ea2puml package – EA → PlantUML exporter (refactor scaffold).

This package keeps the CLI and behavior stable while splitting responsibilities:
 - COM access isolated in ea_adapter.py
 - Pure data models in models.py
 - PlantUML string builder in renderer.py
 - Diagram-type-specific rendering in handlers/*
 - Registry/decorator for handler lookup in handler_registry.py
 - CLI wiring in cli.py, main orchestration in main.py
"""
__all__ = [
    "cli",
    "config",
    "ea_adapter",
    "handler_registry",
    "handlers",
    "main",
    "models",
    "renderer",
]
