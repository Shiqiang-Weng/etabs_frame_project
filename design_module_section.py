"""
Legacy shim – concrete design workflow now lives in analysis.design_workflow.
"""
from analysis.design_workflow import perform_concrete_design_and_extract_results  # noqa: F401

__all__ = ["perform_concrete_design_and_extract_results"]

