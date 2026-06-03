import sys
from unittest.mock import MagicMock

# Mock xlwings and win32 modules before importing tax_scraper
sys.modules['xlwings'] = MagicMock()
sys.modules['win32com'] = MagicMock()
sys.modules['win32com.client'] = MagicMock()
sys.modules['win32com.client.dynamic'] = MagicMock()
