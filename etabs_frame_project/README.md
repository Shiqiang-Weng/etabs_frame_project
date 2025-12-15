# ETABS Frame Project Pipeline

Canonical five-stage flow (主流程，可开关):
1. 几何建模 (`geometry_modeling`)
   - `define_all_materials_and_sections`
   - `create_frame_structure` (columns, beams, slabs, base restraints)
2. 荷载定义与施加 (`load_module`)
   - `setup_response_spectrum`, `define_all_load_cases`
   - `assign_all_loads_to_frame_structure`
3. 结构分析 (`analysis`)
   - `wait_and_run_analysis`, `check_analysis_completion`
4. 分析结果提取 (`results_extraction`)
   - `extract_modal_and_drift`, `export_beam_and_column_element_forces`
   - `export_frame_member_forces`, `export_other_output_items_tables`
5. 构件设计（可选，默认关闭） (`analysis.perform_concrete_design_and_extract_results`)
   - 设计成功后可按配置导出设计内力/报表

Entry point: `main.py` calls `run_pipeline()` to execute the stages using the canonical packages and the shared `common` layer.

### PipelineOptions 示例

`PipelineOptions` 控制每个阶段是否执行，默认关闭阶段 5（不创建 `design_data`）。示例：

```python
from main import PipelineOptions, run_pipeline

if __name__ == "__main__":
    options = PipelineOptions(
        run_stage1_geometry=True,
        run_stage2_loads=True,
        run_stage3_analysis=True,
        run_stage4_results=True,
        run_stage5_design=False,  # 默认 False：跳过构件设计/设计内力提取
    )
    run_pipeline(options)
```

Legacy/compat files: `*_old.py`, `_tmp*`, `tmp_script.py`, `rewrite_main_block.py`, and root shims (load_cases.py, materials_sections.py, etc.) remain for backward compatibility but are not part of the main pipeline.

## Cleanup / Legacy Files
- Removed legacy/temp files (not part of the pipeline; recoverable from VCS history if needed): `_tmp_edit.py`, `_tmp_generate_design_results.py`, `etabs_setup_old.py`, `frame_geometry_old.py`, `main_old.py`, `rewrite_main_block.py`, `tmp_script.py`.
- Remaining shims kept for backward compatibility (safe to remove later if no external scripts depend on them): `config.py`, `etabs_api_loader.py`, `etabs_setup.py`, `utility_functions.py`, `file_operations.py`, `materials_sections.py`, `frame_geometry.py`, `load_cases.py`, `load_assignment.py`, `response_spectrum.py`, `analysis_module.py`, `design_module*.py`, `Concrete_Frame_Detail_Data.py`, `section_diagnostic.py`, `results_extraction/design_workflow.py`.
