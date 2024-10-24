#!/usr/bin/env python3

import sys
import subprocess
import argparse

parser = argparse.ArgumentParser(
        prog="KiPFG",
        description="Generate production files for KiCad"
)

parser.add_argument('filename')

args = parser.parse_args()

subprocess.run(
    [
        "kicad-cli",
        "pcb",
        "export",
        "pdf",
        "-l",
        "F.Cu",
        args.filename
    ]
)
