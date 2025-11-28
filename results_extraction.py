"""
Compatibility shim to expose results_extraction package when file deletion is restricted.
"""
from pathlib import Path
import importlib

# Treat this module as a package by providing a __path__ to the folder.
__path__ = [str(Path(__file__).parent / "results_extraction")]

_analysis = importlib.import_module(__name__ + ".analysis_results_module")
_core = importlib.import_module(__name__ + ".core_results_module")
_design_forces = importlib.import_module(__name__ + ".design_forces")
_design_results = importlib.import_module(__name__ + ".design_results")
_member_forces = importlib.import_module(__name__ + ".member_forces")

extract_modal_and_mass_info = _analysis.extract_modal_and_mass_info
extract_story_drifts_improved = _analysis.extract_story_drifts_improved
extract_modal_and_drift = _analysis.extract_modal_and_drift
export_core_results = _core.export_core_results
extract_design_forces_and_summary = _design_forces.extract_design_forces_and_summary
extract_design_results_enhanced = _design_results.extract_design_results_enhanced
save_design_results_enhanced = _design_results.save_design_results_enhanced
generate_enhanced_summary_report = _design_results.generate_enhanced_summary_report
extract_and_save_beam_results = _design_results.extract_and_save_beam_results
extract_and_save_column_results = _design_results.extract_and_save_column_results
extract_frame_forces = _member_forces.extract_frame_forces
save_forces_to_csv = _member_forces.save_forces_to_csv
extract_and_save_frame_forces = _member_forces.extract_and_save_frame_forces

__all__ = [
    "extract_modal_and_mass_info",
    "extract_story_drifts_improved",
    "extract_modal_and_drift",
    "export_core_results",
    "extract_design_forces_and_summary",
    "extract_design_results_enhanced",
    "save_design_results_enhanced",
    "generate_enhanced_summary_report",
    "extract_and_save_beam_results",
    "extract_and_save_column_results",
    "extract_frame_forces",
    "save_forces_to_csv",
    "extract_and_save_frame_forces",
]
