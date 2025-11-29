"""
Legacy shim – load assignment now lives in load_module.assignment.
New code should import from load_module instead of this module.
"""
from load_module.assignment import *  # noqa: F401,F403

