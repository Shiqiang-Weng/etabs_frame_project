# ETABS Frame Project Pipeline

Canonical four-stage flow (主流程):
1. 几何建模 (`geometry_modeling`)
   - `define_all_materials_and_sections`
   - `create_frame_structure` (columns, beams, slabs, base restraints)
2. 荷载定义与施加 (`load_module`)
   - `setup_response_spectrum`, `define_all_load_cases`
   - `assign_all_loads_to_frame_structure`
3. 分析 / 设计求解 (`analysis`)
   - `wait_and_run_analysis`, `check_analysis_completion`
   - `perform_concrete_design_and_extract_results` (可选，根据 `config.PERFORM_CONCRETE_DESIGN`)
4. 结果提取 / 后处理 (`results_extraction`)
   - `extract_modal_and_drift`, `extract_and_save_frame_forces`
   - `export_core_results`; optional `extract_design_forces_and_summary`

Entry point: `main.py` calls `run_pipeline()` to execute the four stages using the canonical packages and the shared `common` layer.

Legacy/compat files: `*_old.py`, `_tmp*`, `tmp_script.py`, `rewrite_main_block.py`, and root shims (load_cases.py, materials_sections.py, etc.) remain for backward compatibility but are not part of the main pipeline.

## Cleanup / Legacy Files
- Removed legacy/temp files (not part of the pipeline; recoverable from VCS history if needed): `_tmp_edit.py`, `_tmp_generate_design_results.py`, `etabs_setup_old.py`, `frame_geometry_old.py`, `main_old.py`, `rewrite_main_block.py`, `tmp_script.py`.
- Remaining shims kept for backward compatibility (safe to remove later if no external scripts depend on them): `config.py`, `etabs_api_loader.py`, `etabs_setup.py`, `utility_functions.py`, `file_operations.py`, `materials_sections.py`, `frame_geometry.py`, `load_cases.py`, `load_assignment.py`, `response_spectrum.py`, `analysis_module.py`, `design_module*.py`, `Concrete_Frame_Detail_Data.py`, `section_diagnostic.py`, `results_extraction/design_workflow.py`.
