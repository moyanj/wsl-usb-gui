# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

WSL USB GUI is a Windows desktop application that provides a GUI for managing USB device connections between Windows and WSL (Windows Subsystem for Linux). It's a wxPython-based application that wraps the `usbipd-win` command-line tools.

## Development Commands

### Running the Application
```bash
# Run directly from source
python -m wsl_usb_gui

# Or run the main module
python wsl_usb_gui/__main__.py
```

### Building
```bash
# Build executable with PyOxidizer
pyoxidizer build

# Build release version
pyoxidizer build --release

# Build MSI installer (Windows only)
pyoxidizer build msi_installer --release
```

### Dependency Management
```bash
# Install dependencies (uses Rye)
rye sync

# Install development dependencies
rye sync --dev

# Update dependencies
rye add <package_name>
```

### Version Management
```bash
# Generate version from git tags
rye run git-versioner --python --short --save
```

## Architecture

### Core Components

- **`gui.py`** - Main wxPython GUI application with async support via wxasync
- **`usb_monitor.py`** - USB device monitoring and Windows API integration via ctypes
- **`install.py`** - Automatic dependency installation and WSL configuration
- **`logger.py`** - Centralized logging configuration with file and console output
- **`win_usb_inspect/`** - Windows USB device inspection utilities

### Key Dependencies

- **wxPython 4.2.1** - GUI framework
- **wxasync** - Async support for wxPython
- **PySerial** - USB device enumeration
- **appdirs** - Cross-platform app directories
- **git-versioner** - Automatic version generation

### External Integrations

- **usbipd-win** - Core USB-over-IP functionality (bundled MSI installer)
- **Windows API** - Device notifications, registry access, admin privileges
- **WSL** - Linux kernel integration for USB device forwarding

### Build System

- **PyOxidizer** - Packages Python app into standalone Windows executables
- **WiX Toolset** - MSI installer generation
- **Rye** - Modern Python dependency management
- **Code Signing** - Via SignPath.io (OSS program)

## Code Conventions

### Style Guidelines
- **PEP 8** compliance with 4-space indentation
- **Type hints** using the `typing` module
- **PascalCase** for classes (`WslUsbGui`, `ListCtrl`)
- **snake_case** for functions/variables (`load_config`, `usb_devices`)
- **UPPER_SNAKE_CASE** for constants (`CONFIG_FILE`, `DEVICE_COLUMNS`)

### Error Handling
- Use `wx.MessageBox` for user-facing errors
- Use `logging` module for system errors and debugging
- Wrap external process calls in try/except blocks

### Async Operations
- Use `asyncio` for non-blocking operations
- External commands via `asyncio.create_subprocess_exec`
- GUI updates via `wxasync` to maintain responsiveness

### Configuration
- User data stored in `appdirs.user_data_dir()` location
- Configuration in `config.json`
- Logging to rotating `log.txt` file

## Development Notes

### Testing
- No automated test suite - manual testing required
- Test with actual USB devices and WSL environments
- Verify admin privilege handling and device binding

### Windows-Specific Features
- Deep Windows API integration via ctypes
- Registry access for device management
- Admin privilege elevation for USB binding
- System tray integration for background operation

### Deployment
- MSI installer with code signing
- Bundles official usbipd-win installer (unmodified)
- No telemetry or network communications
- Automated CI/CD via AppVeyor and GitLab CI