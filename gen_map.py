import ast, json
from pathlib import Path

root = Path('.')
files = [p for p in root.rglob('*.py') if '.idea' not in p.parts]
modules = {}
for path in files:
    rel = path.relative_to(root)
    parts = rel.with_suffix('').parts
    if parts[-1] == '__init__':
        import_name = '.'.join(parts[:-1])
    else:
        import_name = '.'.join(parts)
    modules[import_name] = {'path': str(rel), 'imports': []}

for import_name, info in modules.items():
    path = Path(info['path'])
    try:
        src = path.read_text(encoding='utf-8-sig')
    except UnicodeDecodeError:
        src = path.read_text(encoding='utf-8', errors='ignore')
    try:
        tree = ast.parse(src)
    except SyntaxError:
        continue
    imports=set()
    parent_parts = import_name.split('.') if import_name else []
    if path.name != '__init__.py' and parent_parts:
        parent_parts = parent_parts[:-1]

    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                imports.add(alias.name)
        elif isinstance(node, ast.ImportFrom):
            base_module = node.module or ''
            if node.level:
                prefix = parent_parts[: len(parent_parts) - (node.level - 1)] if node.level > 0 else parent_parts
                full_parts = prefix + (base_module.split('.') if base_module else [])
                base = '.'.join([p for p in full_parts if p])
            else:
                base = base_module
            if base:
                imports.add(base)
            for alias in node.names:
                name = alias.name
                imports.add(base + ('.' if base else '') + name)
    info['imports'] = sorted(imports)

imported_by = {name: [] for name in modules}
for mod, info in modules.items():
    for target in info['imports']:
        for candidate in modules:
            if target == candidate or target.startswith(candidate + '.'):
                imported_by[candidate].append(mod)

for mod in modules:
    modules[mod]['imported_by'] = sorted(set(imported_by[mod]))

Path('module_map.json').write_text(json.dumps(modules, indent=2), encoding='utf-8')
print('modules', len(modules))
