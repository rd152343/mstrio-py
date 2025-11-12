"""
Zebra Configuration Package

This package provides centralized configuration management for Zebra MicroStrategy scripts.
"""

from .zebra_config_manager import ZebraConfig, get_zebra_config

__all__ = ['ZebraConfig', 'get_zebra_config']