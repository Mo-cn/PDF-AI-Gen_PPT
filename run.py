#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""PDF-AI-GEN_PPT 命令行入口"""

import sys
from pathlib import Path

src_path = Path(__file__).parent
if str(src_path) not in sys.path:
    sys.path.insert(0, str(src_path))

from src.main import main

if __name__ == '__main__':
    main()
