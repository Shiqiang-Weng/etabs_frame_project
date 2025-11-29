import json
from pathlib import Path

data=json.loads(Path('dep_report.json').read_text())
project_mods=set(data['modules'].keys())
reverse={mod:set() for mod in project_mods}
for mod, info in data['modules'].items():
    for s in info['imports']+info['calls']:
        for pm in project_mods:
            if pm!=mod and (s==pm or s.startswith(pm+'.')):
                reverse[pm].add(mod)
lines=[]
for mod, users in sorted(reverse.items()):
    if users:
        lines.append(f"{mod}: {', '.join(sorted(users))}")
Path('dep_reverse.txt').write_text('\n'.join(lines), encoding='utf-8')
