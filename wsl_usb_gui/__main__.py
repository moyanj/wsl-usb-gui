from .logger import log

try:
    from . import gui

    gui.main()
except:
    log.exception("APPCRASH")
