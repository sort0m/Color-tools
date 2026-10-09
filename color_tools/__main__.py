"""
__main__.py — Käynnistysmoduuli.
Aja: python -m color_tools
"""

import traceback

try:
    from .update_loop import main
    main()
except Exception:
    _msg = traceback.format_exc()
    from .bootstrap import _fatal   # jos tämäkin epäonnistuu, Python näyttää alkuperäisen virheen
    _fatal(_msg)
