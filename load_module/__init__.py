"""
Unified entry point for load definition, assignment, and response spectrum setup.
Expose clear APIs for the main workflow while keeping original helpers available.
"""

from .cases import (
    ensure_dead_pattern,
    ensure_live_pattern,
    define_static_load_cases,
    define_modal_case,
    define_response_spectrum_cases,
    define_response_spectrum_combinations,
    define_mass_source_simple,
    define_all_load_cases,
)
from .assignment import (
    assign_dead_and_live_loads_to_slabs,
    assign_finish_loads_to_beams,
    assign_column_loads_fixed,
    assign_seismic_mass_to_structure,
    assign_all_loads_to_frame_structure,
)
from .response_spectrum import (
    china_response_spectrum,
    generate_response_spectrum_data,
    define_response_spectrum_functions_in_etabs,
)

# Friendly aliases for the core steps in the main workflow
define_load_cases = define_all_load_cases
assign_loads_to_model = assign_all_loads_to_frame_structure
setup_response_spectrum = define_response_spectrum_functions_in_etabs

__all__ = [
    # Core workflow aliases
    "define_load_cases",
    "assign_loads_to_model",
    "setup_response_spectrum",
    # Case definitions
    "ensure_dead_pattern",
    "ensure_live_pattern",
    "define_static_load_cases",
    "define_modal_case",
    "define_response_spectrum_cases",
    "define_response_spectrum_combinations",
    "define_mass_source_simple",
    "define_all_load_cases",
    # Load assignment
    "assign_dead_and_live_loads_to_slabs",
    "assign_finish_loads_to_beams",
    "assign_column_loads_fixed",
    "assign_seismic_mass_to_structure",
    "assign_all_loads_to_frame_structure",
    # Response spectrum helpers
    "china_response_spectrum",
    "generate_response_spectrum_data",
    "define_response_spectrum_functions_in_etabs",
]
