"""
Legacy shim – load case definitions now live in load_module.cases.
New code should import from load_module instead of this module.
"""
from load_module.cases import *  # noqa: F401,F403

