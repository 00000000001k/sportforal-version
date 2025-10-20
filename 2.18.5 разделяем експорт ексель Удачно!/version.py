# version.py
"""
Файл версії програми
"""
from globals import version

__version__ = version
__build__ = "29.05.2025"

def get_version():
    """Повертає поточну версію програми"""
    return __version__

def get_build():
    """Повертає дату збірки"""
    return __build__