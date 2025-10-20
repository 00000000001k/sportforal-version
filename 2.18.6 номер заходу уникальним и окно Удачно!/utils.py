# utils.py
# -*- coding: utf-8 -*-

import os
import sys

def get_executable_dir():
    """
    Повертає директорію де знаходиться виконуваний файл або скрипт
    """
    if getattr(sys, 'frozen', False):
        # Якщо це exe файл створений через PyInstaller
        return os.path.dirname(sys.executable)
    else:
        # Якщо це звичайний Python скрипт
        return os.path.dirname(os.path.abspath(__file__))