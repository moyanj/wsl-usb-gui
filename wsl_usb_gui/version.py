from importlib.metadata import version


try:
    __version__ = version("wsl_usb_gui")
except:
    # package is probably not installed
    try:
        import __version__ as git_versioner
        __version__ = git_versioner.version_py_short
    except ImportError:
        # git_versioner not yet installed
        __version__ = ""
