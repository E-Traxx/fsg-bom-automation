"""FSG CCBOM Automation — e-traxx entry point.

Env vars live in `src/env.py`; logic lives in `src/etraxx.py`.
Usage: python e_traxx.py
"""

from src.etraxx import main

if __name__ == "__main__":
    main()
