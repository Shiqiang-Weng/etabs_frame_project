import ast, json
from pathlib import Path

files = [Path(p) for p in [
"_tmp_generate_design_results.py","_tmp_edit.py","utility_functions.py","tmp_script.py","section_diagnostic.py","rewrite_main_block.py","file_operations.py","etabs_setup_old.py","etabs_setup.py","etabs_api_loader.py","design_module_section.py","design_module_column.py","design_module.py","config.py","Concrete_Frame_Detail_Data.py","analysis_module.py","results_extraction/__init__.py","results_extraction/member_forces.py","results_extraction/design_workflow.py","results_extraction/design_results.py","results_extraction/design_forces.py","results_extraction/core_results_module.py","results_extraction/analysis_results_module.py","response_spectrum.py","materials_sections.py","main_old.py","main.py","frame_geometry.py","frame_geometry_old.py","load_assignment.py","load_cases.py","analysis/__init__.py","analysis/status.py","analysis/runner.py","load_module/__init__.py","load_module/response_spectrum.py","load_module/cases.py","load_module/assignment.py","geometry_modeling/__init__.py","geometry_modeling/model_builder.py","geometry_modeling/layout.py","geometry_modeling/geometry_utils.py","geometry_modeling/base_constraints.py","geometry_modeling/api_compat.py",
] if Path(p).exists()]

def resolve_call_name(node):
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Attribute):
        base = resolve_call_name(node.value)
        if base:
            return base + '.' + node.attr
        else:
            return node.attr
    return None

def resolve_import_base(module_name, level, module):
    if level == 0:
        return module or ''
    parts = module_name.split('.')
    if parts and parts[-1] == '__init__':
        parts = parts[:-1]
    parts = parts[:-level]
    if module:
        parts += module.split('.')
    return '.'.join([p for p in parts if p])

modules = {}
def_to_module = {}
call_edges = []

for path in files:
    module_name = path.with_suffix('').as_posix().replace('/', '.')
    text = path.read_text(encoding='utf-8-sig', errors='ignore')
    try:
        tree = ast.parse(text)
    except SyntaxError as e:
        print(f"Failed parsing {path}: {e}")
        continue

    functions = []
    classes = []
    methods = []
    alias_map = {}
    imports = set()

    class ImportVisitor(ast.NodeVisitor):
        def visit_Import(self, node):
            for alias in node.names:
                alias_map[alias.asname or alias.name] = alias.name
                imports.add(alias.name)
        def visit_ImportFrom(self, node):
            base = resolve_import_base(module_name, node.level, node.module)
            for alias in node.names:
                name = alias.name
                asname = alias.asname or name
                full = base + ('.' if base else '') + name
                alias_map[asname] = full
                imports.add(full)
    ImportVisitor().visit(tree)

    class DefVisitor(ast.NodeVisitor):
        def __init__(self):
            self.class_stack = []
        def visit_FunctionDef(self, node):
            if self.class_stack:
                cname = self.class_stack[-1]
                methods.append(f"{cname}.{node.name}")
            else:
                functions.append(node.name)
            self.generic_visit(node)
        visit_AsyncFunctionDef = visit_FunctionDef
        def visit_ClassDef(self, node):
            classes.append(node.name)
            self.class_stack.append(node.name)
            self.generic_visit(node)
            self.class_stack.pop()
    DefVisitor().visit(tree)

    calls = []
    class CallVisitor(ast.NodeVisitor):
        def visit_Call(self, node):
            target = resolve_call_name(node.func)
            if target:
                parts = target.split('.')
                base = parts[0]
                mapped_base = alias_map.get(base, base)
                mapped = '.'.join([mapped_base] + parts[1:]) if parts[1:] else mapped_base
                calls.append(mapped)
            self.generic_visit(node)
    CallVisitor().visit(tree)

    modules[module_name] = {
        'file': str(path),
        'functions': sorted(set(functions)),
        'classes': sorted(set(classes)),
        'methods': sorted(set(methods)),
        'imports': sorted(imports),
        'calls': calls,
    }

    for fn in functions:
        def_to_module.setdefault(fn, set()).add(module_name)
    for cls in classes:
        def_to_module.setdefault(cls, set()).add(module_name)
    for m in methods:
        def_to_module.setdefault(m, set()).add(module_name)

for caller, info in modules.items():
    for call in info['calls']:
        call_edges.append((caller, call))

usage = {mod: {'functions': set(), 'classes': set(), 'methods': set()} for mod in modules}
for caller, call in call_edges:
    target = call.split('.')[0]
    for name, mods in def_to_module.items():
        if target == name or call == name:
            for m in mods:
                if '.' in name:
                    usage[m]['methods'].add(caller)
                elif name and name[0].isupper():
                    usage[m]['classes'].add(caller)
                else:
                    usage[m]['functions'].add(caller)

report = {
    'modules': modules,
    'usage': {k:{'functions':sorted(v['functions']), 'classes': sorted(v['classes']), 'methods': sorted(v['methods'])} for k,v in usage.items()},
}
Path('dep_report.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
