"""
run_app.py — Käynnistin (esim. PyInstaller-pakettia varten).

Sama virheenkäsittely kuin __main__.py:ssä: käynnistysvirhe kirjataan lokiin
ja näytetään dialogina sen sijaan että ikkunaton prosessi vain katoaa.
"""

import traceback

if __name__ == "__main__":
    try:
        from color_tools.update_loop import main
        main()
    except Exception:
        _msg = traceback.format_exc()
        from color_tools.bootstrap import _fatal   # jos tämäkin epäonnistuu, Python näyttää alkuperäisen virheen
        _fatal(_msg)
