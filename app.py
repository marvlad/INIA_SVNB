# update_colorimetric_report.py

from pathlib import Path
from shutil import copyfile
import re
import unicodedata
import warnings
from datetime import datetime
import argparse
import sys

from openpyxl import load_workbook
from openpyxl.drawing.image import Image


# ============================================================
# Windows UTF-8 console fix
# ============================================================

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
