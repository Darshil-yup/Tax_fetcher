import sys
from unittest.mock import MagicMock

# Block ALL windows-only and Excel modules before anything else imports them
_win_modules = [
    'win32com',
    'win32com.client',
    'win32api',
    'win32con',
    'pythoncom',
    'pywintypes',
    'xlwings',
    'xlwings.constants',
    'xlwings.utils',
]

for mod in _win_modules:
    sys.modules[mod] = MagicMock()
