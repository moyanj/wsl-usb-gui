from .logger import log

try:
    from . import gui

    gui.main()
except:
    log.exception("应用程序崩溃")
