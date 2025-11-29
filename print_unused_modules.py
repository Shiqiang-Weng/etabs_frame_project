#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Utility: print modules that are not imported when loading the canonical pipeline entry points.
Safe to run; does not delete or modify anything.
"""
import importlib
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def discover_modules():
    modules = {}
    for path in ROOT.rglob('*.py'):
        if '.idea' in path.parts:
            continue
        rel = path.relative_to(ROOT)
        parts = rel.with_suffix('').parts
        if parts[-1] == '__init__':
            import_name = '.'.join(parts[:-1])
        else:
            import_name = '.'.join(parts)
        modules[import_name] = str(rel)
    return modules


def try_import_targets(targets):
    errors = {}
    for name in targets:
        try:
            importlib.import_module(name)
        except Exception as exc:  # noqa: BLE001
            errors[name] = str(exc)
    return errors


def main():
    modules = discover_modules()
    targets = ['main', 'geometry_modeling', 'load_module', 'analysis', 'results_extraction']
    errors = try_import_targets(targets)

    imported = set(sys.modules.keys())
    unused = []
    for mod_name, path in modules.items():
        # consider used if module itself or any submodule is loaded
        if mod_name in imported or any(m.startswith(mod_name + '.') for m in imported):
            continue
        unused.append((mod_name, path))

    print("=== Import attempts ===")
    if errors:
        for name, msg in errors.items():
            print(f"[FAIL] {name}: {msg}")
    else:
        print("All target imports succeeded.")

    print("\n=== Unused modules (not loaded via main/geometry_modeling/load_module/analysis/results_extraction) ===")
    if not unused:
        print("None")
        return

    for mod_name, path in sorted(unused):
        print(f"- {mod_name}  (path: {path})")


if __name__ == '__main__':
    main()
