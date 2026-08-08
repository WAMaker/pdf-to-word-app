#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""应用入口(供 PyInstaller 打包)"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.main import main  # noqa: E402

if __name__ == "__main__":
    main()
