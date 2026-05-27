# -*- coding: utf-8 -*-
"""
Pytest configuration for Phase 1 tests
"""
import sys
import os

print("CONFTEST: Running conftest.py")
print("CONFTEST: __file__ would be:", __file__)

# Add shipping_helper to path
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
print("CONFTEST: Adding to path:", project_root)
sys.path.insert(0, project_root)
print("CONFTEST: sys.path now:", sys.path[:3])