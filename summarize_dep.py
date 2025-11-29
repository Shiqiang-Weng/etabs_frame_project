import json
from pathlib import Path

data = json.loads(Path('dep_report.json').read_text(encoding='utf-8'))
project_mods = set(data['modules'].keys())
lines = []

for mod, info in sorted(data['modules'].items()):
    deps = set()
    for s in info['imports'] + info['calls']:
        for pm in project_mods:
            if s == pm or s.startswith(pm + '.'):
                deps.add(pm)
    usage = data['usage'].get(mod, {})
    used_by = set(usage.get('functions', [])) | set(usage.get('classes', [])) | set(usage.get('methods', []))
    lines.append(f"Module: {mod} ({info['file']})")
    if info['functions']:
        lines.append("  Functions: " + ', '.join(info['functions']))
    if info['classes']:
        lines.append("  Classes: " + ', '.join(info['classes']))
    if info['methods']:
        lines.append("  Methods: " + ', '.join(info['methods']))
    if deps:
        lines.append("  Project deps: " + ', '.join(sorted(deps)))
    if used_by:
        lines.append("  Used by: " + ', '.join(sorted(used_by)))
    lines.append('')

Path('dep_summary.txt').write_text('\n'.join(lines), encoding='utf-8')
