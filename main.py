#!/usr/bin/env python
"""
CATIA Parametric Pipe Converter — Entry Point.

Launches the Tkinter GUI for interactive centerline + pipe building.
"""
import os
import sys

# Ensure project root is on the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.gui_app import main as gui_main


if __name__ == "__main__":
    gui_main()
