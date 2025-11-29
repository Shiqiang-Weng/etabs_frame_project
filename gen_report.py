import json
from pathlib import Path
from collections import defaultdict

data = json.loads(Path('module_map.json').read_text())

shim_names = {
    'config','etabs_api_loader','etabs_setup','utility_functions','file_operations',
    'materials_sections','Concrete_Frame_Detail_Data','section_diagnostic','frame_geometry',
    'results_extraction.design_workflow','analysis_module','design_module','design_module_column','design_module_section',
    'load_cases','load_assignment','response_spectrum'
}
legacy_names = set()
for name, info in data.items():
    stem = Path(info['path']).stem
    fname = Path(info['path']).name
    if stem.endswith('_old') or stem.startswith('_tmp') or fname in {'tmp_script.py','rewrite_main_block.py','main_old.py','etabs_setup_old.py','frame_geometry_old.py','_tmp_edit.py','_tmp_generate_design_results.py'}:
        legacy_names.add(name)

core_prefixes = ['geometry_modeling','load_module','analysis','results_extraction','common']
core_names = set()
for name in data:
    if name=='main' or any(name==pref or name.startswith(pref+'.') for pref in core_prefixes):
        core_names.add(name)

role_map = {}
for name,info in data.items():
    if name in legacy_names:
        role='legacy_marked'
    elif name in shim_names:
        role='shim'
    elif name in core_names:
        if info['path'].startswith('results_extraction/concrete_frame_detail_data') or info['path'].startswith('results_extraction/section_diagnostic'):
            role='test_or_tool'
        else:
            role='core_pipeline'
    else:
        role='uncertain'
    role_map[name]=role

by_role=defaultdict(list)
for name,info in data.items():
    by_role[role_map[name]].append((name,info))

def short(lst):
    if not lst:
        return '-'
    return ', '.join(lst) if len(lst)<=4 else ', '.join(lst[:3])+f', +{len(lst)-3}'

lines=[]
for role in ['core_pipeline','shim','legacy_marked','test_or_tool','uncertain']:
    items=by_role.get(role,[])
    if not items:
        continue
    lines.append(f"## {role} ({len(items)})\n")
    lines.append("| module_path | import_name | imported_by | imports |")
    lines.append("| --- | --- | --- | --- |")
    for name,info in sorted(items, key=lambda x: x[0]):
        impby = short(info['imported_by'])
        imps = short(info['imports'])
        lines.append(f"| {info['path']} | {name} | {impby} | {imps} |")
    lines.append("")

lines.append('## Cleanup Candidates\n')
lines.append('| module_path | role | imported_by | candidate_for_deletion | reasoning |')
lines.append('| --- | --- | --- | --- | --- |')

core_roots={'main','geometry_modeling','load_module','analysis','results_extraction','common'}

for name,info in sorted(data.items()):
    role = role_map[name]
    imported_by = set(info['imported_by'])
    candidate = False
    reason = ''
    if role in {'legacy_marked'}:
        candidate=True; reason='Legacy/temp file, not in pipeline.'
    elif role=='shim':
        if not imported_by or all(r not in core_roots and not any(r.startswith(cr+'.') for cr in core_roots) for r in imported_by):
            candidate='safe_if_no_external_users'
            reason='Shim re-export; not used by pipeline internally.'
    elif role=='test_or_tool':
        if not imported_by or imported_by.issubset(legacy_names|shim_names):
            candidate=True; reason='Tool/diagnostic not referenced by pipeline.'
    else:
        candidate=False; reason=''
    if candidate:
        impby = short(sorted(imported_by)) or 'none'
        lines.append(f"| {info['path']} | {role} | {impby} | {candidate} | {reason} |")

Path('module_report.md').write_text('\n'.join(lines), encoding='utf-8')
print('report written')
