"""pytest configuration — adds project root to sys.path."""

import sys
from pathlib import Path

# Add project root to Python path so `import tokens` etc. work
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
