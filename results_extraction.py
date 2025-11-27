"""
Compatibility shim to expose results_extraction package when file deletion is restricted.
"""
from pathlib import Path
import importlib

# Treat this module as a package by providing a __path__ to the folder.
__path__ = [str(Path(__file__).parent / "results_extraction")]

_analysis = importlib.import_module(__name__ + ".analysis_results_module")
_core = importlib.import_module(__name__ + ".core_results_module")

extract_modal_and_mass_info = _analysis.extract_modal_and_mass_info
extract_story_drifts_improved = _analysis.extract_story_drifts_improved
extract_modal_and_drift = _analysis.extract_modal_and_drift
export_core_results = _core.export_core_results

__all__ = [
    "extract_modal_and_mass_info",
    "extract_story_drifts_improved",
    "extract_modal_and_drift",
    "export_core_results",
]
