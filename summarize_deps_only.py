import json
from pathlib import Path

data = json.loads(Path('dep_report.json').read_text(encoding='utf-8'))
project_mods = set(data['modules'].keys())
lines = []
for mod, info in sorted(data['modules'].items()):
    deps = set()
    for s in info['imports'] + info['calls']:
        for pm in project_mods:
            if pm == mod:
                continue
            if s == pm or s.startswith(pm + '.'):
                deps.add(pm)
    if deps:
        lines.append(f"{mod}: {', '.join(sorted(deps))}")
Path('dep_deps.txt').write_text('\n'.join(lines), encoding='utf-8')
