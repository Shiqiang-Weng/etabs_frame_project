## core_pipeline (30)

| module_path | import_name | imported_by | imports |
| --- | --- | --- | --- |
| analysis\__init__.py | analysis | analysis, analysis_module, design_module, +4 | analysis.design_workflow, analysis.design_workflow.perform_concrete_design_and_extract_results, analysis.runner, +4 |
| analysis\design_workflow.py | analysis.design_workflow | analysis, design_module, design_module_column, +2 | ETABSv1, System, clr, +15 |
| analysis\runner.py | analysis.runner | analysis | __future__, __future__.annotations, common.config, +13 |
| analysis\status.py | analysis.status | analysis | __future__, __future__.annotations, etabs_setup, etabs_setup.get_etabs_objects |
| common\__init__.py | common | analysis.design_workflow, analysis.runner, config, +12 | - |
| common\config.py | common.config | analysis.design_workflow, analysis.runner, config, +8 | dataclasses, dataclasses.dataclass, os, +2 |
| common\etabs_api_loader.py | common.etabs_api_loader | analysis.runner, etabs_api_loader, main | ETABSv1, System, System.Runtime.InteropServices, +7 |
| common\etabs_setup.py | common.etabs_setup | analysis.design_workflow, analysis.runner, etabs_setup, +7 | config, config.ATTACH_TO_INSTANCE, config.BOTTOM_STORY_HEIGHT, +16 |
| common\file_operations.py | common.file_operations | file_operations, main | config, config.ATTACH_TO_INSTANCE, config.MODEL_PATH, +9 |
| common\utility_functions.py | common.utility_functions | analysis.design_workflow, analysis.runner, geometry_modeling.materials_sections, +5 | etabs_api_loader, etabs_api_loader.get_api_objects, sys, +5 |
| geometry_modeling\__init__.py | geometry_modeling | frame_geometry, geometry_modeling, geometry_modeling.base_constraints, +4 | geometry_modeling.api_compat, geometry_modeling.api_compat._get_all_points_safe, geometry_modeling.api_compat._get_name_list_safe, +33 |
| geometry_modeling\api_compat.py | geometry_modeling.api_compat | geometry_modeling, geometry_modeling.base_constraints, geometry_modeling.model_builder | etabs_api_loader, etabs_api_loader.get_api_objects, etabs_setup, +5 |
| geometry_modeling\base_constraints.py | geometry_modeling.base_constraints | geometry_modeling, geometry_modeling.model_builder | geometry_modeling.api_compat, geometry_modeling.api_compat._get_all_points_safe, geometry_modeling.api_compat._get_name_list_safe, +11 |
| geometry_modeling\geometry_utils.py | geometry_modeling.geometry_utils | geometry_modeling.model_builder | geometry_modeling.layout, geometry_modeling.layout.GridConfig, logging, +6 |
| geometry_modeling\layout.py | geometry_modeling.layout | geometry_modeling, geometry_modeling.base_constraints, geometry_modeling.geometry_utils, geometry_modeling.model_builder | config, config.BOTTOM_STORY_HEIGHT, config.FRAME_BEAM_HEIGHT, +13 |
| geometry_modeling\materials_sections.py | geometry_modeling.materials_sections | geometry_modeling, materials_sections | common.config, common.config.CONCRETE_E_MODULUS, common.config.CONCRETE_MATERIAL_NAME, +17 |
| geometry_modeling\model_builder.py | geometry_modeling.model_builder | geometry_modeling | common.config, common.config.FRAME_BEAM_SECTION_NAME, common.config.FRAME_COLUMN_SECTION_NAME, +28 |
| load_module\__init__.py | load_module | load_assignment, load_cases, load_module, +2 | load_module.assignment, load_module.assignment.assign_all_loads_to_frame_structure, load_module.assignment.assign_column_loads_fixed, +16 |
| load_module\assignment.py | load_module.assignment | load_assignment, load_module | config, config.DEFAULT_DEAD_SUPER_SLAB, config.DEFAULT_FINISH_LOAD_BEAM, +9 |
| load_module\cases.py | load_module.cases | load_cases, load_module | common.config, common.config.GENERATE_RS_COMBOS, common.config.GRAVITY_ACCEL, +10 |
| load_module\response_spectrum.py | load_module.response_spectrum | load_module, response_spectrum | config, config.GRAVITY_ACCEL, config.RS_BASE_ACCEL_G, +11 |
| main.py | main | - | analysis, common, common.config, +12 |
| results_extraction\__init__.py | results_extraction | Concrete_Frame_Detail_Data, analysis.design_workflow, main, +3 | __future__, __future__.annotations, common.config, +33 |
| results_extraction\analysis_results_module.py | results_extraction.analysis_results_module | results_extraction, results_extraction.core_results_module | __future__, __future__.annotations, common.config, +18 |
| results_extraction\concrete_frame_detail_data.py | results_extraction.concrete_frame_detail_data | Concrete_Frame_Detail_Data, results_extraction | common.config, common.config.*, common.etabs_setup, +13 |
| results_extraction\core_results_module.py | results_extraction.core_results_module | results_extraction | __future__, __future__.annotations, config, +12 |
| results_extraction\design_forces.py | results_extraction.design_forces | results_extraction, results_extraction.core_results_module | config, config.*, csv, +12 |
| results_extraction\design_results.py | results_extraction.design_results | analysis.design_workflow, results_extraction | __future__, __future__.annotations, csv, +13 |
| results_extraction\member_forces.py | results_extraction.member_forces | results_extraction | __future__, __future__.annotations, config, +14 |
| results_extraction\section_diagnostic.py | results_extraction.section_diagnostic | results_extraction, section_diagnostic | System, clr, common.config, +4 |

## shim (17)

| module_path | import_name | imported_by | imports |
| --- | --- | --- | --- |
| Concrete_Frame_Detail_Data.py | Concrete_Frame_Detail_Data | - | results_extraction.concrete_frame_detail_data, results_extraction.concrete_frame_detail_data.* |
| analysis_module.py | analysis_module | - | analysis, analysis.check_analysis_completion, analysis.safe_run_analysis, analysis.wait_and_run_analysis |
| config.py | config | analysis.design_workflow, common.etabs_api_loader, common.etabs_setup, +9 | common.config, common.config.* |
| design_module.py | design_module | - | analysis.design_workflow, analysis.design_workflow.perform_concrete_design_and_extract_results |
| design_module_column.py | design_module_column | - | analysis.design_workflow, analysis.design_workflow.perform_concrete_design_and_extract_results |
| design_module_section.py | design_module_section | - | analysis.design_workflow, analysis.design_workflow.perform_concrete_design_and_extract_results |
| etabs_api_loader.py | etabs_api_loader | common.etabs_setup, common.file_operations, common.utility_functions, +12 | common.etabs_api_loader, common.etabs_api_loader.* |
| etabs_setup.py | etabs_setup | analysis.status, common.file_operations, frame_geometry_old, +6 | common.etabs_setup, common.etabs_setup.* |
| file_operations.py | file_operations | - | common.file_operations, common.file_operations.* |
| frame_geometry.py | frame_geometry | - | geometry_modeling, geometry_modeling._get_all_points_safe, geometry_modeling._get_name_list_safe, +14 |
| load_assignment.py | load_assignment | - | load_module.assignment, load_module.assignment.* |
| load_cases.py | load_cases | - | load_module.cases, load_module.cases.* |
| materials_sections.py | materials_sections | - | geometry_modeling.materials_sections, geometry_modeling.materials_sections.* |
| response_spectrum.py | response_spectrum | - | load_module.response_spectrum, load_module.response_spectrum.* |
| results_extraction\design_workflow.py | results_extraction.design_workflow | - | analysis.design_workflow, analysis.design_workflow.* |
| section_diagnostic.py | section_diagnostic | - | results_extraction.section_diagnostic, results_extraction.section_diagnostic.* |
| utility_functions.py | utility_functions | common.etabs_setup, common.file_operations, etabs_setup_old, +7 | common.utility_functions, common.utility_functions.* |

## legacy_marked (7)

| module_path | import_name | imported_by | imports |
| --- | --- | --- | --- |
| _tmp_edit.py | _tmp_edit | - | pathlib, pathlib.Path |
| _tmp_generate_design_results.py | _tmp_generate_design_results | - | - |
| etabs_setup_old.py | etabs_setup_old | - | config, config.ATTACH_TO_INSTANCE, config.BOTTOM_STORY_HEIGHT, +16 |
| frame_geometry_old.py | frame_geometry_old | - | config, config.BOTTOM_STORY_HEIGHT, config.FRAME_BEAM_HEIGHT, +23 |
| main_old.py | main_old | - | - |
| rewrite_main_block.py | rewrite_main_block | - | pathlib, pathlib.Path |
| tmp_script.py | tmp_script | - | - |

## uncertain (7)

| module_path | import_name | imported_by | imports |
| --- | --- | --- | --- |
| gen_map.py | gen_map | - | ast, json, pathlib, pathlib.Path |
| gen_report.py | gen_report | - | collections, collections.defaultdict, json, +2 |
| print_unused_modules.py | print_unused_modules | - | importlib, pathlib, pathlib.Path, sys |
| summarize_dep.py | summarize_dep | - | json, pathlib, pathlib.Path |
| summarize_deps_only.py | summarize_deps_only | - | json, pathlib, pathlib.Path |
| summarize_reverse.py | summarize_reverse | - | json, pathlib, pathlib.Path |
| tmp_dep_scan.py | tmp_dep_scan | - | ast, json, pathlib, pathlib.Path |

## Cleanup Candidates

| module_path | role | imported_by | candidate_for_deletion | reasoning |
| --- | --- | --- | --- | --- |
| Concrete_Frame_Detail_Data.py | shim | - | safe_if_no_external_users | Shim re-export; not used by pipeline internally. |
| _tmp_edit.py | legacy_marked | - | True | Legacy/temp file, not in pipeline. |
| _tmp_generate_design_results.py | legacy_marked | - | True | Legacy/temp file, not in pipeline. |
| analysis_module.py | shim | - | safe_if_no_external_users | Shim re-export; not used by pipeline internally. |
| design_module.py | shim | - | safe_if_no_external_users | Shim re-export; not used by pipeline internally. |
| design_module_column.py | shim | - | safe_if_no_external_users | Shim re-export; not used by pipeline internally. |
| design_module_section.py | shim | - | safe_if_no_external_users | Shim re-export; not used by pipeline internally. |
| etabs_setup_old.py | legacy_marked | - | True | Legacy/temp file, not in pipeline. |
| file_operations.py | shim | - | safe_if_no_external_users | Shim re-export; not used by pipeline internally. |
| frame_geometry.py | shim | - | safe_if_no_external_users | Shim re-export; not used by pipeline internally. |
| frame_geometry_old.py | legacy_marked | - | True | Legacy/temp file, not in pipeline. |
| load_assignment.py | shim | - | safe_if_no_external_users | Shim re-export; not used by pipeline internally. |
| load_cases.py | shim | - | safe_if_no_external_users | Shim re-export; not used by pipeline internally. |
| main_old.py | legacy_marked | - | True | Legacy/temp file, not in pipeline. |
| materials_sections.py | shim | - | safe_if_no_external_users | Shim re-export; not used by pipeline internally. |
| response_spectrum.py | shim | - | safe_if_no_external_users | Shim re-export; not used by pipeline internally. |
| results_extraction\design_workflow.py | shim | - | safe_if_no_external_users | Shim re-export; not used by pipeline internally. |
| rewrite_main_block.py | legacy_marked | - | True | Legacy/temp file, not in pipeline. |
| section_diagnostic.py | shim | - | safe_if_no_external_users | Shim re-export; not used by pipeline internally. |
| tmp_script.py | legacy_marked | - | True | Legacy/temp file, not in pipeline. |