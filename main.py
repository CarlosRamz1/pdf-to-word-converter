#!/usr/bin/env python3
"""
PDF to Word Converter - Main Entry Point
A professional tool for converting PDF files to Word documents with high fidelity.
"""

import sys
import os
from pathlib import Path

# Add src to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from cli.interface import main

if __name__ == "__main__":
    main()
