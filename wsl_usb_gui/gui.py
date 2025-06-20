import appdirs
import asyncio
import json
import logging
import logging.handlers
import os
import serial.tools.list_ports
import re
import sys
from collections import namedtuple
from dataclasses import dataclass, astuple
from functools import partial
from pathlib import Path
from typing import *
import webbrowser
import winreg
import argparse

import wx
import wx.adv
import wxasync
from wx.lib.agw.ultimatelistctrl import (
    UltimateListCtrl,
    UltimateListItem,
    ULC_REPORT,
    ULC_NO_SORT_HEADER,
    ULC_SINGLE_SEL,
    ULC_NO_ITEM_DRAG,
    ULC_MASK_KIND,
    ULC_MASK_TEXT,
    EVT_LIST_ITEM_CHECKED,
    UltimateListMainWindow,
    ULC_HAS_VARIABLE_ROW_HEIGHT
)

from .version import __version__
from .usb_monitor import registerDeviceNotification, unregisterDeviceNotification, WM_SHOW_EXISTING
from .win_usb_inspect import InspectUsbDevices, gDeviceList
from .logger import log, APP_DIR
from .install import MSI_VERS

# High DPI Support.
import ctypes

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(True)
except:
    pass

DEVICE_COLUMNS = ["bus_id", "description", "bound", "forced"]
ATTACHED_COLUMNS = ["bus_id", "description", "forced"]  # , "client"]
PROFILES_COLUMNS = ["bus_id", "description", "enabled"]

CONFIG_FILE = APP_DIR / "config.json"

ProcResult = namedtuple("ProcResult", ("stdout", "stderr", "returncode"))

@dataclass
class Device:
    BusId: str
    Description: str
    bound: bool
    forced: bool
    InstanceId: str
    Attached: str
    OrigDescription: str

    def __eq__(self, __value: object) -> bool:
        if not isinstance(__value, Device):
            return False
        return (
            self.BusId == __value.BusId and self.InstanceId == __value.InstanceId
        )

    def __hash__(self) -> int:
        return hash((self.BusId, self.InstanceId))

@dataclass
class Profile:
    BusId: Optional[str] = None
    Description: Optional[str] = None
    InstanceId: Optional[str] = None
    enabled: Optional[bool] = None

    def __post_init__(self):
        if self.enabled is None:
            self.enabled = True

gui: Optional["WslUsbGui"] = None
loop = None
INSTALLED_DEPS = False

USBIPD_default = Path("C:\\Program Files\\usbipd-win\\usbipd.exe")
USBIPD_VERSION: Tuple[int, ...] = (0, 0, 0)
if USBIPD_default.exists():
    USBIPD = USBIPD_default
else:
    # try to run from anywhere on path, will try to install later if needed
    USBIPD = "usbipd"


async def run(args, decode=True):
    CREATE_NO_WINDOW = 0x08000000
    if isinstance(args, str):
        proc = await asyncio.create_subprocess_shell(
            args,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
            creationflags=CREATE_NO_WINDOW,
        )
    else:
        proc = await asyncio.create_subprocess_exec(
            *args,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
            creationflags=CREATE_NO_WINDOW,
        )

    stdout, stderr = await proc.communicate()
    # Recreate a basic "process results" object to return.
    if decode:
        return ProcResult(stdout.decode(), stderr.decode(), proc.returncode)
    else:
        return ProcResult(stdout, stderr, proc.returncode)


def get_resource(name):
    fname = Path(sys.executable).parent / name
    if not fname.exists():
        try:
            fname = Path(__file__).parent / name
        except:
            fname = None
    return fname


def get_icon(name="usb.ico"):
    icon = get_resource(name)
    if icon is not None:
        return wx.Icon(str(icon), wx.BITMAP_TYPE_ICO)


class WslUsbGui(wx.Frame):
    def __init__(self, minimised=False):
        wx.Frame.__init__(self, None, title=f"WSL USB Manager {__version__}")

        self.icon = get_icon()

        if self.icon:
            self.SetIcon(self.icon)

        # On first run after an upgrade, allow certain post-install tasks
        self.first_run = True

        # On first close, alert the user it's being minimised to tray
        self.informed_about_tray = False

        self.usb_devices: List[Device] = []
        self.pinned_profiles: List[Profile] = []
        self.name_mapping = dict()
        self.hidden_devices = list()
        self.show_hidden = False
        self.refreshing = False
        self.refreshing_delay = False

        self.auto_start_at_boot = False
        self.close_to_tray = True


        self.load_config()

        self.taskbar = TaskBarIcon(self, self.icon)

        self.filemenu = wx.Menu()
        # wx.ID_ABOUT and wx.ID_EXIT are standard IDs provided by wxWidgets.
        settings_menu = self.filemenu.Append(wx.ID_ANY, "&Settings"," Show Settings")
        self.Bind(wx.EVT_MENU, self.show_settings_window, settings_menu)
        about_menu = self.filemenu.Append(wx.ID_ABOUT, "&About"," Information about this program")
        self.Bind(wx.EVT_MENU, self._go_to_about, about_menu)
        logs_menu = self.filemenu.Append(wx.ID_ANY, "&Logs"," Open logs folder in explorer")
        self.Bind(wx.EVT_MENU, self._open_logs_folder, logs_menu)
        hide_menu = self.filemenu.Append(wx.ID_ANY, "Mi&nimise"," Minimise to Tray")
        self.Bind(wx.EVT_MENU, self.minimise, hide_menu)
        udev_menu = self.filemenu.Append(wx.ID_ANY, "Udev: allow all"," Add user permissions for all devices udev rule")
        self.Bind(wx.EVT_MENU, bg_af(self.udev_permissive_all), udev_menu)
        self.filemenu.AppendSeparator()
        exit_menu = self.filemenu.Append(wx.ID_EXIT, "E&xit"," Terminate the program")
        self.Bind(wx.EVT_MENU, self.taskbar.OnExit, exit_menu)

        self.menuBar = wx.MenuBar()
        self.menuBar.Append(self.filemenu,"&File")
        self.SetMenuBar(self.menuBar)

        self.statusbar = self.CreateStatusBar()

        # Intercept window close.
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        splitter_bottom = ProportionalSplitter(self, proportion=0.66, style=wx.SP_LIVE_UPDATE)
        splitter_top = ProportionalSplitter(
            splitter_bottom, proportion=0.33, style=wx.SP_LIVE_UPDATE
        )

        splitter_top.SetMinimumPaneSize(6)
        splitter_bottom.SetMinimumPaneSize(6)

        headingFont = wx.Font(
            16, wx.FONTFAMILY_DECORATIVE, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD
        )

        ## TOP SECTION - Available USB Devices

        top_panel = wx.Panel(splitter_top, style=wx.SUNKEN_BORDER)
        top_sizer = wx.BoxSizer(wx.VERTICAL)
        top_panel.SetSizerAndFit(top_sizer)

        available_list_label = wx.StaticText(top_panel, label="Windows USB Devices")
        available_list_label.SetFont(headingFont)

        self.busy_icon = wx.adv.AnimationCtrl(top_panel, wx.ID_ANY)
        self.busy_icon.LoadFile(str(get_resource("busy.gif")))
        refresh_button = self.Button(top_panel, "Refresh", command=self.refresh)

        top_controls = wx.BoxSizer(wx.HORIZONTAL)
        top_controls.Add(available_list_label, 2, wx.EXPAND | wx.TOP | wx.LEFT, border=6)
        top_controls.AddStretchSpacer(1)
        top_controls.Add(self.busy_icon, 0, wx.TOP, border=6)
        top_controls.AddSpacer(6)
        top_controls.Add(refresh_button, 1, wx.TOP, border=6)

        self.available_listbox = ListCtrl(top_panel, type=Device)
        self.available_listbox.InsertColumns(DEVICE_COLUMNS)

        async def available_menu(event):
            x, y = event.GetX(), event.GetY()
            popupmenu = wx.Menu()
            device = self.get_selected_device(verbose=False)
            clicked_on_background = False
            if device:
                item = self.available_listbox.GetFirstSelected()
                _mainWin: UltimateListMainWindow = self.available_listbox._mainWin
                x, y = _mainWin.CalcUnscrolledPosition(x, y)
                if _mainWin.InReportView():
                    if not _mainWin.HasAGWFlag(ULC_HAS_VARIABLE_ROW_HEIGHT):
                        current = y // _mainWin.GetLineHeight()
                        if current != item:
                            # clicked on background
                            clicked_on_background = True

            if not device or clicked_on_background:
                if self.show_hidden:
                    entries = [
                        ("Mask hidden devices", self.mask_hidden_devices,),
                    ]
                else:
                    entries = [
                        ("Show hidden devices", self.show_hidden_devices,),
                    ]
            else:
                entries = [
                    ("Attach to WSL", bg_af(self.attach_wsl)),
                    ("Auto-Attach Device", self.auto_attach_wsl),
                    ("Rename Device", self.rename_device),

                ]
                if not device.bound:
                    entries.extend([
                        ("Bind", bg_af(self.bind)),
                        ("Force Bind", bg_af(self.force_bind)),
                    ])
                else:
                    entries.extend([
                        ("Force Bind", bg_af(self.force_bind)),
                        ("Unbind", bg_af(self.unbind)),
                    ])
                if device.InstanceId not in self.hidden_devices:
                    entries.extend([
                        ("Hide", self.hide_device),
                    ])
                else:
                    entries.extend([
                        ("Unhide", self.unhide_device),
                    ])
                entries.extend([
                    ("Create Custom Auto Attach Profile", self.add_custom_profile_context_menu),
                ])

            for entry, fn in entries:
                menuItem = popupmenu.Append(-1, entry)
                self.Bind(wx.EVT_MENU, fn, menuItem)
            # Show menu
            self.PopupMenu(popupmenu)

        wxasync.AsyncBind(wx.EVT_RIGHT_UP, available_menu, self.available_listbox)

        async def available_checked(event):
            available_listbox = event.EventObject
            device: Device = available_listbox.devices[event.Index]
            bound = available_listbox.GetItem(event.Index, col=2).IsChecked()
            forced = available_listbox.GetItem(event.Index, col=3).IsChecked()
            if device.forced and not forced:
                await self.unbind_bus_id(device.BusId)
            elif forced and not device.forced:
                await self.bind_bus_id(device.BusId, forced=True)
            elif device.bound and not bound:
                await self.unbind_bus_id(device.BusId)
            elif bound and not device.bound:
                await self.bind_bus_id(device.BusId, forced=False)

        # self.available_listbox.Bind(EVT_LIST_ITEM_CHECKED, available_checked)
        wxasync.AsyncBind(EVT_LIST_ITEM_CHECKED, available_checked, self.available_listbox)

        top_sizer.Add(top_controls, flag=wx.EXPAND | wx.ALL, border=6)
        top_sizer.Add(self.available_listbox, proportion=1, flag=wx.EXPAND | wx.ALL, border=6)

        ## MIDDLE SECTION - USB devices currently attached

        middle_panel = wx.Panel(splitter_top, style=wx.SUNKEN_BORDER)
        middle_sizer = wx.BoxSizer(wx.VERTICAL)
        middle_panel.SetSizerAndFit(middle_sizer)

        attached_list_label = wx.StaticText(middle_panel, label="Forwarded Devices")
        attached_list_label.SetFont(headingFont)

        attach_button = self.Button(middle_panel, "Attach ↓", acommand=self.attach_wsl)

        detach_button = self.Button(middle_panel, "↑ Detach", acommand=self.detach_wsl)
        auto_attach_button = self.Button(middle_panel, "Auto-Attach", command=self.auto_attach_wsl)
        rename_button = self.Button(middle_panel, "Rename", command=self.rename_device)

        middle_controls = wx.BoxSizer(wx.HORIZONTAL)
        middle_controls.Add(attached_list_label, 2, wx.EXPAND | wx.TOP | wx.LEFT, border=6)
        middle_controls.AddStretchSpacer(1)
        middle_controls.Add(attach_button, 1, wx.TOP, border=6)
        middle_controls.Add(detach_button, 1, wx.TOP, border=6)
        middle_controls.Add(auto_attach_button, 1, wx.TOP, border=6)
        middle_controls.Add(rename_button, 1, wx.TOP, border=6)

        self.attached_listbox = ListCtrl(middle_panel, type=Device)
        self.attached_listbox.InsertColumns(ATTACHED_COLUMNS)

        async def attached_menu(event):
            popupmenu = wx.Menu()
            entries = [
                ("Detach Device", bg_af(self.detach_wsl)),
                ("Auto-Attach Device", self.auto_attach_wsl),
                ("Rename Device", self.rename_device),
                ("WSL: Grant User Permissions", bg_af(self.udev_permissive)),
            ]
            for entry, fn in entries:
                menuItem = popupmenu.Append(-1, entry)
                self.Bind(wx.EVT_MENU, fn, menuItem)
            # Show menu
            self.PopupMenu(popupmenu)

        wxasync.AsyncBind(wx.EVT_LIST_ITEM_RIGHT_CLICK, attached_menu, self.attached_listbox)


        async def attached_checked(event):
            attached_listbox = event.EventObject
            device: Device = attached_listbox.devices[event.Index]
            forced = attached_listbox.GetItem(event.Index, col=2).IsChecked()
            if device.forced and not forced:
                await self.unbind_bus_id(device.BusId)
            elif forced and not device.forced:
                await self.bind_bus_id(device.BusId, forced=True)
            return True

        wxasync.AsyncBind(EVT_LIST_ITEM_CHECKED, attached_checked, self.attached_listbox)

        middle_sizer.Add(middle_controls, flag=wx.EXPAND | wx.ALL, border=6)
        middle_sizer.Add(self.attached_listbox, proportion=1, flag=wx.EXPAND | wx.ALL, border=6)

        ## BOTTOM SECTION - saved profiles for auto-attach

        bottom_panel = wx.Panel(splitter_bottom, style=wx.SUNKEN_BORDER)
        bottom_sizer = wx.BoxSizer(wx.VERTICAL)
        bottom_panel.SetSizerAndFit(bottom_sizer)

        pinned_list_label = wx.StaticText(bottom_panel, label="Auto-Attach Profiles")
        pinned_list_label.SetFont(headingFont)
        pinned_list_delete_button = self.Button(
            bottom_panel, "Delete Profile", command=self.delete_profile
        )
        pinned_list_add_custom_button = self.Button(
            bottom_panel, "Custom Profile", command=self.add_custom_profile
        )

        bottom_controls = wx.BoxSizer(wx.HORIZONTAL)
        bottom_controls.Add(pinned_list_label, 2, wx.EXPAND | wx.TOP | wx.LEFT, border=6)
        bottom_controls.AddStretchSpacer(1)
        bottom_controls.Add(pinned_list_add_custom_button, 1, wx.TOP, border=6)
        bottom_controls.Add(pinned_list_delete_button, 1, wx.TOP, border=6)

        self.pinned_listbox = ListCtrl(bottom_panel, type=Profile)
        self.pinned_listbox.InsertColumns(PROFILES_COLUMNS)

        def profile_menu(event):
            popupmenu = wx.Menu()
            entries = [
                ("Delete Profile", self.delete_profile),
                ("Edit Profile", self.edit_profile),
            ]
            for entry, fn in entries:
                menuItem = popupmenu.Append(-1, entry)
                self.Bind(wx.EVT_MENU, fn, menuItem)

            self.PopupMenu(popupmenu)  # , position)

        def _profile_menu(event):
            wx.CallAfter(profile_menu, event)

        self.pinned_listbox.Bind(wx.EVT_LIST_ITEM_RIGHT_CLICK, _profile_menu)

        async def profile_checked(event):
            pinned_listbox = event.EventObject
            # print(f"aaaaaaa: {pinned_listbox.__dict__=}")
            # await asyncio.sleep(0.1)  # Ensure the event is processed after the item is checked
            profile: Profile = pinned_listbox.devices[event.Index]
            enabled = pinned_listbox.GetItem(event.Index, col=2).IsChecked()
            if profile.enabled != enabled:
                profile.enabled = enabled
                self.save_config()
                self.refresh(delay=1.0)
            return True

        wxasync.AsyncBind(EVT_LIST_ITEM_CHECKED, profile_checked, self.pinned_listbox)

        bottom_sizer.Add(bottom_controls, flag=wx.EXPAND | wx.ALL, border=6)
        bottom_sizer.Add(self.pinned_listbox, proportion=1, flag=wx.EXPAND | wx.ALL, border=6)

        ## Window Configure

        # Ensure only one device can be selected at a time
        self.available_listbox.Bind(
            wx.EVT_LIST_ITEM_SELECTED,
            partial(self.deselect_other_treeviews, treeview=self.available_listbox),
        )
        self.attached_listbox.Bind(
            wx.EVT_LIST_ITEM_SELECTED,
            partial(self.deselect_other_treeviews, treeview=self.attached_listbox),
        )
        self.pinned_listbox.Bind(
            wx.EVT_LIST_ITEM_SELECTED,
            partial(self.deselect_other_treeviews, treeview=self.pinned_listbox),
        )

        splitter_top.SplitHorizontally(top_panel, middle_panel)
        splitter_bottom.SplitHorizontally(splitter_top, bottom_panel)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(splitter_bottom, proportion=1, flag=wx.EXPAND)

        sizer.SetSizeHints(self)
        self.SetSizerAndFit(sizer)

        self.Show(True)
        self.SetSize(self.FromDIP(wx.Size(600, 800)))
        self.update_pinned_listbox()
        if minimised:
            self.Hide()

    # # Override MSWWindowProc to handle the custom message
    # def MSWWindowProc(self, msg, wparam, lparam):
    #     if msg == WM_SHOW_EXISTING:
    #         if self.IsIconized():
    #             self.Iconize(False)
    #         if not self.IsShown():
    #             self.Show(True)
    #         self.Restore() # Ensure it's not minimized
    #         self.Raise()   # Bring to front
    #         log.info("Received WM_SHOW_EXISTING message, brought window to front.")
    #         return 0 # Indicate message was handled

    #     # Call the default handler for other messages
    #     try:
    #         # Attempt to call super's MSWWindowProc if it exists
    #         return super().MSWWindowProc(msg, wparam, lparam)
    #     except AttributeError:
    #          # If super() doesn't have MSWWindowProc, use DefWindowProc
    #          return ctypes.windll.user32.DefWindowProcW(self.GetHandle(), msg, wparam, lparam)

    def Button(self, parent, button_text, command=None, acommand=None):
        btn = wx.Button(parent, label=button_text)
        btn.SetMaxSize(parent.FromDIP(wx.Size(90, 30)))
        if command:
            self.Bind(wx.EVT_BUTTON, lambda event: command(), btn)
        elif acommand:
            async def awrap(event):
                await acommand()
            wxasync.AsyncBind(wx.EVT_BUTTON, awrap, btn)

        return btn

    def OnClose(self, event):
        if event.CanVeto():
            if self.close_to_tray:
                if not self.informed_about_tray:
                    wx.MessageBox(
                        caption="Minimising to tray", message=f"This will stay running in background.\nCan be restored/exited from system tray icon.", style=wx.OK | wx.ICON_INFORMATION
                    )
                    self.informed_about_tray = True
                    self.save_config()

                # Hide the window instead of closing it
                self.Hide()
                event.Veto()
            else:
                wx.CallAfter(self.taskbar.OnExit, True)

        else:
            event.Skip()
            self.Destroy()

    def minimise(self, _):
        self.Hide()

    def _open_logs_folder(self, _event):
        os.startfile(APP_DIR)

    def _go_to_about(self, _event):
        webbrowser.open_new("https://gitlab.com/alelec/wsl-usb-gui")

    @staticmethod
    def create_profile(busid, description, instanceid, enabled=None):
        return Profile(*(None if a == "None" else a for a in (busid, description, instanceid, enabled)))

    def show_settings_window(self, _):
        settings_window = SettingsWindow(self)
        settings_window.ShowModal()

    def load_config(self):
        try:
            log.info(f"Loading config from: {CONFIG_FILE}")
            config = json.loads(CONFIG_FILE.read_text())
            if isinstance(config, list):
                self.pinned_profiles = [self.create_profile(*c) for c in config]
            else:
                self.pinned_profiles = [self.create_profile(*c) for c in config["pinned_profiles"]]
                self.name_mapping = config["name_mapping"]
                self.hidden_devices = config.get("hidden_devices", [])
                self.informed_about_tray = config.get("informed_about_tray", False)
                self.first_run = config.get("first_run", True)
                self.auto_start_at_boot = config.get("auto_start_at_boot", self.auto_start_at_boot)
                self.close_to_tray = config.get("close_to_tray", self.close_to_tray)

        except Exception as ex:
            pass

    def save_config(self):
        config = dict(
            pinned_profiles=[astuple(p) for p in self.pinned_profiles],
            name_mapping=self.name_mapping,
            hidden_devices=self.hidden_devices,
            informed_about_tray=self.informed_about_tray,
            first_run=__version__,
            auto_start_at_boot=self.auto_start_at_boot,
            close_to_tray=self.close_to_tray,

        )
        CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
        CONFIG_FILE.write_text(json.dumps(config, indent=4, sort_keys=True))

    def update_auto_start_at_boot(self):
        exe = sys.executable
        if Path(exe).name.lower().startswith("py"):
            log.warning("Ignore auto-start, running from source")
            return
        registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        runKey = winreg.OpenKey(
            registry,
            "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run",
            access=winreg.KEY_SET_VALUE | winreg.KEY_QUERY_VALUE
        )
        key = "WSL-USB-GUI"
        if self.auto_start_at_boot:
            winreg.SetValueEx(runKey, key, 0, winreg.REG_SZ, f'"{exe}" --minimised')
        else:
            winreg.DeleteKey(runKey, key)

    def parse_state(self, text) -> List[Device]:
        rows = []
        devices = json.loads(text)

        for device in devices["Devices"]:
            bus_info = device["BusId"]
            if bus_info:
                instanceId = device["InstanceId"]
                orig_description = device["Description"]
                description = self.name_mapping.get(instanceId, device["Description"])
                # bind = "☒" if device["PersistedGuid"] else "☐"
                # forced = "☒" if device["IsForced"] else "☐"
                bind = True if device["PersistedGuid"] else False
                forced = True if device["IsForced"] else False
                attached = device["ClientIPAddress"]
                rows.append(Device(str(bus_info), description, bind, forced, instanceId, attached, orig_description))
        return rows

    def deselect_other_treeviews(self, *args, treeview: UltimateListCtrl):
        if not treeview.GetSelectedItemCount():
            return

        for tv in (
            self.available_listbox,
            self.attached_listbox,
            self.pinned_listbox,
        ):
            if tv is treeview:
                continue
            while (i := tv.GetFirstSelected()) != -1:
                tv.Select(i, on=False)

    async def list_wsl_usb(self) -> List[Device]:
        try:
            result = await run([USBIPD, "state"], decode=False)
            return self.parse_state(result.stdout)
        except Exception as ex:
            if isinstance(ex, FileNotFoundError):
                log.warning("Failed to run usbipd state, install deps")
                install_deps()
            else:
                log.exception("list_wsl_usb")
            return []

    @staticmethod
    async def usbipd_run_admin_if_needed(command, msg=None):
        result = await run(command)
        stderr = result.stderr.lower()
        if "error:" in stderr and "administrator" in stderr:
            if msg:
                wx.MessageBox(
                    caption="Administrator Privileges",
                    message=msg,
                    style=wx.OK | wx.ICON_INFORMATION,
                )
            args_str = ", ".join(f'\\"{arg}\\"' for arg in command[1:])

            result = await run(
                r'''Powershell -Command "& { Start-Process \"%s\" -ArgumentList @(%s) -Verb RunAs } "'''
                % (USBIPD, args_str)
            )
        return result

    async def bind_bus_id(self, bus_id, forced, msg=None):
        command = [USBIPD, "bind", f"--busid={bus_id}"]
        if forced:
            command.append("--force")
        result = await WslUsbGui.usbipd_run_admin_if_needed(command, msg=msg)
        if result.stdout:
            log.info(result.stdout)
        if result.stderr:
            log.error(result.stderr)
        log.info(f"Bind {bus_id}: {'Success' if not result.returncode else 'Failed'}")
        self.refresh(delay=1.0)
        return result

    async def unbind_bus_id(self, bus_id):
        command = [USBIPD, "unbind", f"--busid={bus_id}"]
        result = await WslUsbGui.usbipd_run_admin_if_needed(command)
        if result.stdout:
            log.info(result.stdout)
        if result.stderr:
            log.error(result.stderr)
        log.info(f"Unbind {bus_id}: {'Success' if not result.returncode else 'Failed'}")
        self.refresh(delay=1.0)
        return result

    async def attach_wsl_usb(self, device: Device):
        msg = (
            "The first time attaching a device to WSL requires elevated privileges; "
            "subsequent attaches use standard user privileges."
        )
        if not device.bound:
            result = await self.bind_bus_id(device.BusId, forced=False, msg=msg)
            await asyncio.sleep(3)
            msg = None

        result = None
        if USBIPD_VERSION < (4, 0, 0):
            command = [USBIPD, "wsl", "attach", "--busid=" + device.BusId]
            result = await WslUsbGui.usbipd_run_admin_if_needed(command, msg)

        if result is None or (result.returncode != 0):
            command = [USBIPD, "attach", "--wsl", "--busid=" + device.BusId]
            result = await WslUsbGui.usbipd_run_admin_if_needed(command, msg)

        status = f"Attached: {device.Description}"
        if result.stdout:
            log.info(result.stdout)
            status = result.stdout
        if result.stderr:
            log.error(result.stderr)
            stderr_lower = result.stderr.lower()
            for stderr_line in result.stderr.strip().split("\n"):
                stderr_lower_line = stderr_line.lower()
                if "usbipd: info:" in stderr_lower_line and "error" not in stderr_lower_line:
                    pass
                else:
                    if "client not correctly installed" in stderr_lower_line:
                        status = "Client not correctly installed, installing dependencies..."
                        log.warning(status)
                        install_deps()

                    elif "is already attached to a client." in stderr_lower_line:
                        # Not an error, we've just tried to attach twice.
                        return result

                    elif "device busy (exported)" in stderr_lower_line or "the device appears to be used by windows" in stderr_lower_line:
                        status = "Error: device in use; stop the software using it, or force bind the device."
                        break

                    else:
                        status = stderr_line

        self.SetStatusText("  " + status)
        return result

    @staticmethod
    async def detach_wsl_usb(bus_id):
        result = None
        if USBIPD_VERSION < (4, 0, 0):
            result = await run([USBIPD, "wsl", "detach", "--busid=" + str(bus_id)])
        if not result or (result.returncode != 0):
            result = await run([USBIPD, "detach", "--busid=" + str(bus_id)])

        if result.stdout:
            log.info(result.stdout)
        if result.stderr:
            log.error(result.stderr)

    def update_pinned_listbox(self, focus=None):
        self.pinned_listbox.DeleteAllItems()
        for profile in self.pinned_profiles:
            if not profile.Description:
                profile.Description = self.lookup_description(profile.InstanceId)
            highlight = focus is not None and profile == focus
            self.pinned_listbox.Append(profile, highlight=highlight)

    # Define a function to implement choice function
    def auto_attach_wsl_choice(self, profile: Profile):
        self.pinned_profiles.append(profile)
        self.save_config()
        self.update_pinned_listbox(focus=profile)

    def update_pinned_profile(self, instanceid, new_description):
        for p in list(self.pinned_profiles):
            if p.InstanceId and p.InstanceId == instanceid:
                p.Description = new_description
                self.save_config()

    def remove_pinned_profile(self, busid, description, instanceid):
        profile = self.create_profile(busid, description, instanceid)
        for p in list(self.pinned_profiles):
            if (p.BusId and p.BusId == profile.BusId) or (
                p.InstanceId and p.InstanceId == profile.InstanceId
            ):
                self.pinned_profiles.remove(p)
                self.save_config()

    def delete_profile(self, event=None):
        selection = self.pinned_listbox.GetFirstSelected()
        if selection == -1:
            log.error("no selection to delete")
            return  # no selected item
        profile: Profile = self.pinned_listbox.devices[selection]
        # self.pinned_listbox.DeleteItem(selection)
        self.remove_pinned_profile(profile.BusId, profile.Description, profile.InstanceId)
        self.save_config()
        self.update_pinned_listbox()

    def add_custom_profile_context_menu(self, event):
        device = self.get_selection()
        if not device:
            log.error("no selection to create profile for")
            return
        self.add_custom_profile(device.BusId, device.Description, device.InstanceId)
        self.refresh(delay=3)

    def edit_profile(self, event=None):
        """
        Edit a profile from the pinned listbox.
        Opens a dialog to input bus ID, description, and instance ID.
        """
        selection = self.pinned_listbox.GetFirstSelected()
        if selection == -1:
            log.error("no selection to edit")
            return
        profile: Profile = self.pinned_listbox.devices[selection]
        dlg = CustomProfileDialog(self, "Edit", profile.BusId, profile.Description, profile.InstanceId)
        if dlg.ShowModal() == wx.ID_OK:
            busid, description, instanceid = dlg.get_values()
            if not (busid or (instanceid and description)):
                wx.MessageBox(
                    "Please provide either Bus ID or both Instance ID and Description.",
                    "Invalid Input",
                    wx.OK | wx.ICON_ERROR,
                )
                return
            self.remove_pinned_profile(profile.BusId, profile.Description, profile.InstanceId)
            profile = self.create_profile(busid, description, instanceid)
            self.auto_attach_wsl_choice(profile)
        dlg.Destroy()

    def add_custom_profile_button_handler(self, event):
        """
        Handler for the button to add a custom profile.
        Opens a dialog to input bus ID, description, and instance ID.
        """
        self.add_custom_profile()

    def add_custom_profile(self, busid=None, description=None, instanceid=None):
        """
        Opens a dialog to add a custom profile.
        """
        dlg = CustomProfileDialog(self, "Add", busid, description, instanceid)
        if dlg.ShowModal() == wx.ID_OK:
            busid, description, instanceid = dlg.get_values()
            if not (busid or (instanceid and description)):
                wx.MessageBox(
                    "Please provide either Bus ID or both Instance ID and Description.",
                    "Invalid Input",
                    wx.OK | wx.ICON_ERROR,
                )
                return
            profile = self.create_profile(busid, description, instanceid)
            self.auto_attach_wsl_choice(profile)
        dlg.Destroy()

    def get_selection(self, available=False, attached=False, verbose=True) -> Optional[Device]:
        if not available or attached:
            # If nether specified, return either
            attached = available = True
        if available:
            item = self.available_listbox.GetFirstSelected()
            if item != -1:
                return self.available_listbox.devices[item]
        if attached:
            item = self.attached_listbox.GetFirstSelected()
            if item != -1:
                return self.attached_listbox.devices[item]
        return None

    def get_selected_device(self, available=False, attached=False, verbose=True) -> Optional[Device]:
        selection = self.get_selection(available, attached, verbose=verbose)
        if not selection:
            if verbose:
                log.error("No device selected")
            return None

        device = [
            d
            for d in self.usb_devices
            if d.BusId == selection.BusId and d.Description == selection.Description
        ][0]
        return device

    def unhide_device(self, event=None):
        device = self.get_selected_device()
        if not device:
            return

        if device.InstanceId in self.hidden_devices:
            self.hidden_devices.remove(device.InstanceId)

        self.save_config()
        self.refresh()

    def show_hidden_devices(self, event=None):
        self.show_hidden = True
        self.refresh()

    def mask_hidden_devices(self, event=None):
        self.show_hidden = False
        self.refresh()

    def hide_device(self, event=None):
        device = self.get_selected_device()
        if not device:
            return

        self.hidden_devices.append(device.InstanceId)

        self.save_config()
        self.refresh()


    def rename_device(self, event=None):
        device = self.get_selected_device()
        if not device:
            return
        popupRename(self, device).ShowModal()

    @staticmethod
    def device_ident(device: Device):
        vid, pid, serial = re.search(r"\\VID_([0-9A-Fa-f]+)&PID_([0-9A-Fa-f]+)\\([&0-9A-Za-z]+)$", device.InstanceId).groups()
        return vid, pid, serial


    async def udev_permissive_all(self, event=None):
        udev_rule = (
            f'SUBSYSTEM=="usb|hidraw",MODE="0666"\n'
            f'SUBSYSTEM=="tty",MODE="0666"'
        )
        rules_file = "/etc/udev/rules.d/99-wsl-usb-gui.rules"
        await run([
            "wsl", "--user", "root", "sh", "-c",
            f"echo '{udev_rule}' >> {rules_file}; sudo udevadm control --reload-rules; sudo udevadm trigger",
        ])
        log.info(f"udev all rule added: {udev_rule}")
        wx.MessageBox(
            caption="WSL: Grant User Permissions",
            message=f"WSL udev user permissions rule added for all usb devices",
            style=wx.OK | wx.ICON_INFORMATION,
        )

    async def udev_permissive(self, event=None):
        device = self.get_selected_device()
        if not device:
            return

        try:
            vid, pid, serial = self.device_ident(device)
            name = device.Description.replace(" ", "_")
            # VID and PID must be hex/lowercase. serial on the otherhand must not be converted to lowercase, but copied verbatim.
            vid = vid.lower()
            pid = pid.lower()
            udev_rule_match = f'/SUBSYSTEM=="usb.*ATTRS{{idVendor}}=="{vid}".*ATTRS{{idProduct}}=="{pid}".*ENV{{ID_SERIAL_SHORT}}=="{serial}"/d'
            udev_rule = (
                f'SUBSYSTEM=="usb|hidraw",ATTRS{{idVendor}}=="{vid}",ATTRS{{idProduct}}=="{pid}",ENV{{ID_SERIAL_SHORT}}=="{serial}",MODE="0666",SYMLINK+="usb/{name}"\n'
                f'SUBSYSTEM=="tty",ATTRS{{idVendor}}=="{vid}",ATTRS{{idProduct}}=="{pid}",ENV{{ID_SERIAL_SHORT}}=="{serial}",MODE="0666",SYMLINK+="tty/{name}"'
            )
            rules_file = "/etc/udev/rules.d/99-wsl-usb-gui.rules"
            await run([
                "wsl", "--user", "root", "sh", "-c",
                f"sed -i '{udev_rule_match}' {rules_file}; echo '{udev_rule}' >> {rules_file}; sudo udevadm control --reload-rules; sudo udevadm trigger",
            ])
            # log.info(udev_settings)
            log.info(f"udev rule added: {udev_rule}")
            wx.MessageBox(
                caption="WSL: Grant User Permissions",
                message=f"WSL udev user permissions rule added for VID:{vid} PID:{pid}.",
                style=wx.OK | wx.ICON_INFORMATION,
            )

        except AttributeError as ex:
            log.error(f"Could not get device information for udev: {ex}")
            wx.MessageBox(
                caption="WSL: Grant User Permissions",
                message=f"ERROR: Failed to add udev rule.",
                style=wx.OK | wx.ICON_WARNING,
            )

    @staticmethod
    def window_is_focussed():
        return wx.GetActiveWindow() is not None

    def generic_device_name(self, name):
        if name == "USB Input Device":
            return True
        if re.match(r"USB Serial Device \(COM\d+\)", name):
            return True

    async def refresh_task(self, delay: float = 0):
        try:
            if self.refreshing_delay:
                # There's another refresh about to start
                return
            if delay:
                self.refreshing_delay = True
                await asyncio.sleep(delay)
                self.refreshing_delay = False

            if self.refreshing:
                # Busy right now, might have missed the trigger though to re-run soon
                asyncio.create_task(self.refresh_task(delay=5))
                return

            self.refreshing = True
            self.busy_icon.Show()
            self.busy_icon.Play()

            log.info("Refresh USB")

            task = asyncio.create_task(self.list_wsl_usb())

            comports = {}
            for c in serial.tools.list_ports.comports():
                if getattr(c, "vid", None) and getattr(c, "pid", None):
                    comports[str((f"{c.vid:04X}", f"{c.pid:04X}", c.serial_number.upper()))] = c.name
                await asyncio.sleep(0.01)

            try:
                raw_devices, tree = InspectUsbDevices()
            except:
                log.exception("Failures in InspectUsbDevices")
                raw_devices = {}

            usb_devices = await task

            new_devices = []
            if self.usb_devices:
                # Don't report new device on first run at startup.
                new_devices = set(usb_devices) - set(self.usb_devices)

            self.usb_devices = usb_devices

            self.attached_listbox.DeleteAllItems()
            self.available_listbox.DeleteAllItems()

            if not self.usb_devices:
                return

            tasks = []
            for device in sorted(self.usb_devices, key=lambda d: d.BusId):
                if device.InstanceId in self.hidden_devices:
                    if self.show_hidden:
                        self.available_listbox.Append(device, shade=True)
                    continue

                if self.generic_device_name(device.Description):
                    if device.InstanceId not in self.name_mapping:
                        if details := raw_devices.get(device.InstanceId):
                            if details.Manufacturer and details.Product:
                                device.OrigDescription = device.Description
                                device.Description = f"{details.Manufacturer.strip()} {details.Product.strip()}"

                try:
                    devid = str(self.device_ident(device)).upper()  # (vid, pid, sernum)
                    if str(devid) in comports:
                        windows_com_port = comports[str(devid)]
                        if windows_com_port not in device.Description:
                            device.Description = re.sub(r" ?\(COM\d+\) ?", "", device.Description)
                            device.Description += f" ({windows_com_port})"
                except:
                    pass

                new = device in new_devices
                if device.Attached:
                    self.attached_listbox.Append(device, highlight=new)
                else:
                    task = asyncio.create_task(self.attach_if_pinned(device, highlight=new))
                    tasks.append(task)

            if new_devices and not self.window_is_focussed():
                self.RequestUserAttention()
            await asyncio.gather(*tasks)
        finally:
            self.busy_icon.Stop()
            self.busy_icon.Hide()
            self.refreshing = False
            self.refreshing_delay = False

    def refresh(self, delay=0.0):
        asyncio.get_running_loop().call_soon_threadsafe(
            asyncio.ensure_future, self.refresh_task(delay)
        )

    async def check_wsl_udev(self):
        # Autostart WSL udev service if needed
        udev_start = (await run(
            [
                "wsl",
                "--user",
                "root",
                "sh",
                "-c",
                "pgrep udev || (echo 'starting udev'; service udev restart)",
            ]
        )).stdout.strip()
        udev_start = udev_start.replace("\n", ", ")
        log.info(f"udev: {udev_start}")

    def lookup_description(self, instanceId):
        if not instanceId:
            return None

        if instanceId == "None":
            return None

        for device in self.usb_devices:
            if device.InstanceId == instanceId:
                return device.Description

    @staticmethod
    def compare_profile_value(profileString:str, deviceStr:str) -> bool:
        """
        Compare a profile string with a device string.
        Handles None values and empty strings.
        """
        checkIsRegex = profileString.startswith("re:")
        if checkIsRegex:
            profileString = profileString[3:]  # Remove 're:' prefix for regex matching
        regexForComparison = re.compile(profileString, re.IGNORECASE)
        return bool(regexForComparison.search(deviceStr)) if checkIsRegex else profileString == deviceStr

    async def attach_if_pinned(self, device, highlight):
        enabledProfiles = [p for p in self.pinned_profiles if p.enabled]
        for profile in enabledProfiles:
            desc = profile.Description
            if profile.InstanceId or profile.BusId:
                # Only fallback to description if no other filter set
                desc = None

            if profile.BusId and not self.compare_profile_value(profile.BusId.strip(), device.BusId.strip()):
                continue
            if profile.InstanceId and not self.compare_profile_value(profile.InstanceId, device.InstanceId):
                continue
            if desc and not self.compare_profile_value(desc.strip(), device.Description.strip()):
                continue
            for __retry in reversed(range(3)):
                ret = await self.attach_wsl_usb(device)
                if ret.returncode == 0:
                    self.attached_listbox.Append(device, highlight)
                    return
                await asyncio.sleep(0.5)
            break

        self.available_listbox.Append(device, highlight=highlight)

    async def force_bind(self, event=None):
        device = self.get_selected_device(available=True)
        if not device:
            return
        result = await self.bind_bus_id(device.BusId, forced=True)
        log.info(f"Bind (forced) {device.BusId}: {'Success' if not result.returncode else 'Failed'}")
        self.refresh(delay=3)

    async def bind(self, event=None, refresh=True):
        device = self.get_selected_device(available=True)
        if not device:
            return
        result = await self.bind_bus_id(device.BusId, forced=False)
        self.refresh(delay=3)
        return result

    async def unbind(self, event=None):
        device = self.get_selected_device(available=True)
        if not device:
            return
        await self.unbind_bus_id(device.BusId)
        self.refresh(delay=3)

    async def attach_wsl(self, event=None):
        device = self.get_selected_device(available=True)
        if not device:
            return
        result = await self.attach_wsl_usb(device)
        log.info(f"Attach {device.BusId}: {'Success' if not result.returncode else 'Failed'}")
        self.refresh(delay=3)

    async def detach_wsl(self, event=None):
        device = self.get_selection(attached=True)
        if not device:
            log.error("no selection to detach")
            return  # no selected item
        log.info(f"Detach {device.BusId} {device.Description}")

        self.remove_pinned_profile(device.BusId, device.Description, device.InstanceId)

        await self.detach_wsl_usb(device.BusId)
        self.refresh(delay=2)

    def auto_attach_wsl(self, event=None):
        global pop
        device = self.get_selection()
        if not device:
            log.error("no selection to create profile for")
            return

        popup = popupAutoAttach(
            self, device.BusId, device.Description, device.InstanceId, self.icon
        )
        popup.ShowModal()
        self.refresh(delay=3)

    def raise_window(self):
        if self.IsIconized():
            self.Iconize(False)
        if not self.IsShown():
            self.Show(True)
        self.Restore() # Ensure it's not minimized
        self.Raise()   # Bring to front
        log.info("Received WM_SHOW_EXISTING message, brought window to front.")
        return 0 # Indicate message was handled


class ListCtrl(UltimateListCtrl):
    def __init__(self, parent, *args, type: Type, **kw):
        UltimateListCtrl.__init__(self, parent, wx.ID_ANY, agwStyle=ULC_REPORT|ULC_NO_SORT_HEADER|ULC_SINGLE_SEL|ULC_NO_ITEM_DRAG)

        self.devices: List[type] = []
        self.columns = []

    def IsVisible(self):
        cdc = wx.ClientDC(self)
        # Check some points inside the widget to determine if
        # this window is visible at those points
        widget_size = cdc.Size
        offset = 16
        widget_size.DecBy(offset)
        top_left = self.ClientToScreen((offset, offset))
        bottom_left = self.ClientToScreen((offset, widget_size.height))
        bottom_right = self.ClientToScreen(widget_size)
        tl_visible = wx.FindWindowAtPoint(top_left) is not None
        bl_visible = wx.FindWindowAtPoint(bottom_left) is not None
        br_visible = wx.FindWindowAtPoint(bottom_right) is not None
        return (tl_visible and bl_visible) or (tl_visible and br_visible)

    def HighlightRowReset(self, device, reset):
        try:
            row = self.devices.index(device)
            self.SetItemBackgroundColour(row, reset)
        except ValueError:
            pass  # device has likely been unplugged

    def HighlightRowMaintain(self, device, reset):
        if self.IsVisible():
            wx.CallLater(2000, self.HighlightRowReset, device, reset)
        else:
            wx.CallLater(1000, self.HighlightRowMaintain, device, reset)

    def EnsureVisible(self, item):
        # This can be removed once this is in a release:
        # https://github.com/wxWidgets/Phoenix/commit/54981636c5679e5fb4247cb5b094a7ae29dc545f
        try:
            if item >= self.GetItemCount():
                item = self.GetItemCount() - 1

            _mainWin: UltimateListMainWindow = self._mainWin

            rect = _mainWin.GetLineRect(item)
            client_w, client_h = _mainWin.GetClientSize()
            hLine = _mainWin.GetLineHeight(item)
            view_y = hLine*self.GetScrollPos(wx.VERTICAL)

            if _mainWin.InReportView():
                _mainWin.ResetVisibleLinesRange()
                if not _mainWin.HasAGWFlag(ULC_HAS_VARIABLE_ROW_HEIGHT):
                    if rect.y < view_y:
                        _mainWin.Scroll(-1, rect.y/hLine)
                    if rect.y+rect.height+5 > view_y+client_h:
                        _mainWin.Scroll(-1, (rect.y+rect.height-client_h+hLine)//hLine)
                    _mainWin._dirty = True
        except:
            pass

    def HighlightRow(self, device, index):
        original = self.GetItemBackgroundColour(index)
        self.EnsureVisible(index)  # scroll into view if needed  # needs wx update
        self.SetItemBackgroundColour(index, wx.YELLOW)
        if self.IsVisible():
            wx.CallLater(2000, self.HighlightRowReset, device, original)
        else:
            wx.CallLater(1000, self.HighlightRowMaintain, device, original)

    def InsertColumns(self, names):
        for i, name in enumerate(names):
            info = UltimateListItem()
            info._mask = wx.LIST_MASK_TEXT
            info._image = []
            info._format = 0
            info._kind = 1
            info._text = name
            self.InsertColumnInfo(i, info)
            if i != 1:
                self.SetColumnWidth(i, self.FromDIP(50))
            else:
                LIST_AUTOSIZE_FILL = -3
                self.SetColumnWidth(i, LIST_AUTOSIZE_FILL)
        self.columns = names

    def DeleteAllItems(self):
        pass
        UltimateListCtrl.DeleteAllItems(self)
        self.devices.clear()

    def Append(self, device, highlight=False, shade=False):
        if device in self.devices:
            pos = self.devices.index(device)
        else:
            details = astuple(device)[0 : UltimateListCtrl.GetColumnCount(self)]

            pos = self.GetItemCount()
            self.InsertStringItem(pos, str(details[0] or "---"))
            for i in range(1, len(details)):
                if i == 1:
                    self.SetStringItem(pos, i, str(details[i]))
                else:
                    # Checkbox column
                    column = self.columns[i]
                    value = getattr(device, column)
                    info = UltimateListItem()
                    info._text = ""
                    info._mask = ULC_MASK_TEXT | ULC_MASK_KIND
                    info._kind = 1
                    info.Check(value)

                    info._itemId = pos
                    info._col = i
                    self.SetItem(info)

            self.devices.append(device)
            if len(self.devices) != pos + 1:
                raise ValueError(f"{self} devices out of sync")
            if highlight:
                self.HighlightRow(device, pos)
            elif shade:
                self.SetItemBackgroundColour(pos, wx.LIGHT_GREY)
        return pos


class ProportionalSplitter(wx.SplitterWindow):
    def __init__(self, parent, id=-1, proportion=0.66, size=wx.DefaultSize, **kwargs):
        wx.SplitterWindow.__init__(self, parent, id, wx.Point(0, 0), size, **kwargs)
        self.SetMinimumPaneSize(50)  # the minimum size of a pane.
        self.proportion = proportion
        if not 0 < self.proportion < 1:
            raise ValueError("proportion value for ProportionalSplitter must be between 0 and 1.")
        self.ResetSash()
        self.Bind(wx.EVT_SIZE, self.OnReSize)
        self.Bind(wx.EVT_SPLITTER_SASH_POS_CHANGED, self.OnSashChanged, id=id)
        # hack to set sizes on first paint event
        self.Bind(wx.EVT_PAINT, self.OnPaint)
        self.firstpaint = True

    def SplitHorizontally(self, win1, win2):
        if self.GetParent() is None:
            return False
        return wx.SplitterWindow.SplitHorizontally(
            self, win1, win2, int(round(self.GetParent().GetSize().GetHeight() * self.proportion))
        )

    def SplitVertically(self, win1, win2):
        if self.GetParent() is None:
            return False
        return wx.SplitterWindow.SplitVertically(
            self, win1, win2, int(round(self.GetParent().GetSize().GetWidth() * self.proportion))
        )

    def GetExpectedSashPosition(self):
        if self.GetSplitMode() == wx.SPLIT_HORIZONTAL:
            tot = max(self.GetMinimumPaneSize(), self.GetParent().GetClientSize().height)
        else:
            tot = max(self.GetMinimumPaneSize(), self.GetParent().GetClientSize().width)
        return int(round(tot * self.proportion))

    def ResetSash(self):
        self.SetSashPosition(self.GetExpectedSashPosition())

    def OnReSize(self, event):
        "Window has been resized, so we need to adjust the sash based on self.proportion."
        self.ResetSash()
        event.Skip()

    def OnSashChanged(self, event):
        "We'll change self.proportion now based on where user dragged the sash."
        pos = float(self.GetSashPosition())
        if self.GetSplitMode() == wx.SPLIT_HORIZONTAL:
            tot = max(self.GetMinimumPaneSize(), self.GetParent().GetClientSize().height)
        else:
            tot = max(self.GetMinimumPaneSize(), self.GetParent().GetClientSize().width)
        self.proportion = pos / tot
        event.Skip()

    def OnPaint(self, event):
        if self.firstpaint:
            if self.GetSashPosition() != self.GetExpectedSashPosition():
                self.ResetSash()
            self.firstpaint = False
        event.Skip()


class popupAutoAttach(wx.Dialog):
    def __init__(self, parent, bus_id, description, instanceId, icon):
        super().__init__(parent, title="New Auto-Attach Profile")

        if icon:
            self.SetIcon(icon)

        top_sizer = wx.BoxSizer(wx.VERTICAL)

        message = wx.StaticText(self, label=f"Create Auto-Attach profile for:")
        top_sizer.Add(message, proportion=0, flag=wx.TOP | wx.LEFT, border=12)

        sizer = wx.FlexGridSizer(3, 2, 2, 2)
        device_btn = wx.Button(self, label="Device", size=self.FromDIP(wx.Size(60, 24)))
        device_txt = wx.StaticText(self, label=f"USB Device: {description}")
        port_btn = wx.Button(self, label="Port", size=self.FromDIP(wx.Size(60, 24)))
        port_txt = wx.StaticText(self, label=f"USB Port: {bus_id}")
        both_btn = wx.Button(self, label="Both", size=self.FromDIP(wx.Size(60, 24)))
        both_txt = wx.StaticText(self, label=f"This device only when plugged into same port.")

        sizer.AddMany(
            [
                (device_btn, 1, wx.ALL, 4),
                (device_txt, 1, wx.ALL, 8),
                (port_btn, 1, wx.ALL, 4),
                (port_txt, 1, wx.ALL, 8),
                (both_btn, 1, wx.ALL, 4),
                (both_txt, 1, wx.ALL, 8),
            ]
        )

        self.Bind(
            wx.EVT_BUTTON,
            partial(self.choice, parent, Profile(None, description, instanceId)),
            device_btn,
        )
        self.Bind(
            wx.EVT_BUTTON, partial(self.choice, parent, Profile(bus_id, None, None)), port_btn
        )
        self.Bind(
            wx.EVT_BUTTON,
            partial(self.choice, parent, Profile(bus_id, description, instanceId)),
            both_btn,
        )

        top_sizer.Add(sizer, proportion=0, flag=wx.ALL, border=8)
        self.SetSizerAndFit(top_sizer)

    def Button(self, parent, button_text, command):
        btn = wx.Button(parent, label=button_text)
        btn.SetMaxSize(parent.FromDIP(wx.Size(90, 30)))
        self.Bind(wx.EVT_BUTTON, lambda event: command(), btn)
        return btn

    def choice(self, parent: WslUsbGui, profile, event):
        parent.auto_attach_wsl_choice(profile)
        self.Close()


class popupRename(wx.Dialog):
    def __init__(self, parent: WslUsbGui, device: Device):
        super().__init__(parent, title="Rename Device")

        self.gui = parent
        self.device = device

        self.instanceId = device.InstanceId
        current = self.gui.name_mapping.get(self.instanceId, device.Description)

        top_sizer = wx.BoxSizer(wx.VERTICAL)

        label = f"Enter new label for device on port: {device.BusId}"
        message = wx.StaticText(self, label=label)
        font = message.GetFont()
        font.SetWeight(wx.BOLD)
        message.SetFont(font)
        top_sizer.Add(message, flag=wx.TOP | wx.LEFT | wx.RIGHT, border=12)

        top_sizer.Add(wx.StaticText(self, label="  Click buttons below to fill from defaults,"), flag=wx.LEFT | wx.RIGHT, border=12)
        top_sizer.Add(wx.StaticText(self, label="  Type your own label,"), flag=wx.LEFT | wx.RIGHT, border=12)
        top_sizer.Add(wx.StaticText(self, label="  Or leave blank to reset to default."), flag=wx.LEFT | wx.RIGHT, border=12)

        top_sizer.Add(wx.StaticText(self, label="Driver Default:"), flag=wx.TOP | wx.LEFT | wx.RIGHT, border=12)
        default_btn = wx.Button(self, label=device.OrigDescription)
        top_sizer.Add(default_btn, flag=wx.RIGHT | wx.LEFT | wx.EXPAND, border=20)

        try:
            raw_devices, tree = InspectUsbDevices()
        except:
            log.exception("Failures in InspectUsbDevices")
            raw_devices = {}

        device_btn = None
        if details := raw_devices.get(self.instanceId):
            if details.Manufacturer and details.Product:
                dev_label = f"{details.Manufacturer} {details.Product}"

                top_sizer.Add(wx.StaticText(self, label="Device Details:"), flag=wx.TOP | wx.LEFT | wx.RIGHT, border=12)
                device_btn = wx.Button(self, label=dev_label)
                top_sizer.Add(device_btn, flag=wx.RIGHT | wx.LEFT | wx.EXPAND, border=20)


        top_sizer.Add(wx.StaticLine(self, style=wx.LI_HORIZONTAL | wx.EXPAND), flag=wx.TOP | wx.BOTTOM | wx.LEFT | wx.RIGHT | wx.EXPAND, border=8)

        top_sizer.Add(wx.StaticText(self, label="New Name:"), flag=wx.LEFT | wx.RIGHT, border=12)

        self.textEntry = wx.TextCtrl(self, -1, current, size=wx.Size(-1, self.FromDIP(26)))
        top_sizer.Add(self.textEntry, flag=wx.RIGHT | wx.LEFT | wx.EXPAND, border=12)

        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        ok_btn = wx.Button(self, label="OK", size=self.FromDIP(wx.Size(60, 24)))
        reset_btn = wx.Button(self, label="Reset", size=self.FromDIP(wx.Size(60, 24)))
        cancel_btn = wx.Button(self, label="Cancel", size=self.FromDIP(wx.Size(60, 24)))
        btn_sizer.Add(ok_btn, border=6)
        btn_sizer.Add(reset_btn, border=6)
        btn_sizer.Add(cancel_btn, border=6)
        top_sizer.Add(btn_sizer, flag=wx.BOTTOM | wx.TOP | wx.RIGHT | wx.LEFT | wx.EXPAND, border=12)

        self.Bind(wx.EVT_BUTTON, self.set_text, default_btn)
        if device_btn:
            self.Bind(wx.EVT_BUTTON, self.set_text, device_btn)

        self.Bind(wx.EVT_BUTTON, self.ok, ok_btn)
        self.Bind(wx.EVT_BUTTON, self.reset, reset_btn)
        self.Bind(wx.EVT_BUTTON, self.cancel, cancel_btn)

        ok_btn.SetDefault()

        self.SetSizerAndFit(top_sizer)

    def set_text(self, event: wx.CommandEvent):
        btn: wx.Button = event.EventObject
        self.textEntry.SetValue(btn.GetLabelText())

    def cancel(self, event):
        self.Close()

    def reset(self, event):
        self.textEntry.SetValue("")
        self.ok(event)

    def ok(self, event):
        newname = self.textEntry.GetValue()

        if newname:
            self.gui.update_pinned_profile(self.instanceId, newname)
            self.gui.name_mapping[self.instanceId] = newname
        else:
            try:
                self.gui.name_mapping.pop(self.instanceId)
            except:
                pass
        self.gui.save_config()
        self.gui.refresh()

        self.Close()

def bg_af(fn):
    """
    Call async function in background.
    """
    def wrap(ev=None):
        asyncio.get_running_loop().call_soon_threadsafe(
            asyncio.ensure_future, fn()
        )
    return wrap


def windows_events_callback(event):
    delay = 1.0
    if event == "attach":
        log.info(f"USB device attached")
    elif event == "detach":
        log.info(f"USB device detached")
        delay = 0
    elif event == "show":
        if gui:
            gui.raise_window()
    else:
        log.info(f"unknown windows event")

    if gui:
        gui.refresh(delay=delay)


def install_deps():
    global INSTALLED_DEPS
    if not INSTALLED_DEPS:
        asyncio.get_event_loop().call_soon_threadsafe(_install_deps)
        INSTALLED_DEPS = True


def _install_deps():
    global USBIPD, gui
    rsp = wx.MessageBox(
        caption="Install Dependencies",
        message=(
            "Some of the dependencies are missing or outdated, install them now?\n"
            "Note: All WSL instances may need to be restarted."
        ),
        style=wx.YES_NO | wx.ICON_WARNING,
    )
    if rsp == wx.YES:
        from .install import install_task

        rsp = install_task(gui)
        wx.MessageBox(
            caption="Finished", message="Finished Installation.", style=wx.OK | wx.ICON_INFORMATION
        )

        if rsp:
            USBIPD = USBIPD_default
            if gui:
                gui.refresh()


class TaskBarIcon(wx.adv.TaskBarIcon):
    def __init__(self, frame, icon):
        super(TaskBarIcon, self).__init__()

        self.frame = frame

        # Set the icon for the system tray
        if icon:
            self.SetIcon(icon)

        # Bind the left-click event to the OnLeftClick function
        self.Bind(wx.adv.EVT_TASKBAR_LEFT_DOWN, self.OnLeftClick)

    def CreatePopupMenu(self):
        menu = wx.Menu()

        # Add an option to restore the window when clicked
        restore = menu.Append(wx.ID_ANY, 'Restore')
        self.Bind(wx.EVT_MENU, self.OnRestore, restore)

        # Add an option to exit the application when clicked
        exit_app = menu.Append(wx.ID_EXIT, 'Exit')
        self.Bind(wx.EVT_MENU, self.OnExit, exit_app)

        return menu

    def OnLeftClick(self, event=None):
        self.OnRestore()

    def OnRestore(self, event=None):
        self.frame.Show()
        self.frame.Restore()
        self.frame.Raise()

    def Close(self, event=None):
        self.RemoveIcon()
        self.Destroy()

    def OnExit(self, event=None):
        wx.CallAfter(self.frame.Close, True)
        self.RemoveIcon()
        self.Destroy()
        # self.Close()

async def check_usbipd_version():
    global USBIPD, USBIPD_VERSION
    try:
        vers_str = (await run([USBIPD, "--version"])).stdout
        vers_parts: re.Match[str] = re.search(r'(\d+)\.(\d+)\.(\d+)(?:[-+](\d+))?', vers_str) # type: ignore
        version = tuple((int(v) for v in vers_parts.groups()))
        USBIPD_VERSION = version
    except Exception as ex:
        log.error(f"Could not read usbipd version: {ex}")
        install_deps()


class CustomProfileDialog(wx.Dialog):
    def __init__(self, parent, prefix="Add", bus_id=None, description=None, instance_id=None):
        super().__init__(parent, title=f"{prefix} Custom Profile", style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        # All controls must use self as parent
        sizer = wx.BoxSizer(wx.VERTICAL)

        label = wx.StaticText(self, label=f"{prefix} custom auto-attach profile:")
        sizer.Add(label, flag=wx.ALL, border=8)

        grid = wx.FlexGridSizer(3, 2, 8, 8)
        grid.AddGrowableCol(1, 1)

        busid_label = wx.StaticText(self, label="Bus ID:")
        self.busid_ctrl = wx.TextCtrl(self, value=bus_id or "")
        grid.Add(busid_label, flag=wx.ALIGN_CENTER_VERTICAL)
        grid.Add(self.busid_ctrl, flag=wx.EXPAND)

        description_label = wx.StaticText(self, label="Description:")
        self.description_ctrl = wx.TextCtrl(self, value=description or "")
        grid.Add(description_label, flag=wx.ALIGN_CENTER_VERTICAL)
        grid.Add(self.description_ctrl, flag=wx.EXPAND)

        instanceid_label = wx.StaticText(self, label="Instance ID:")
        self.instanceid_ctrl = wx.TextCtrl(self, value=instance_id or "")
        grid.Add(instanceid_label, flag=wx.ALIGN_CENTER_VERTICAL)
        grid.Add(self.instanceid_ctrl, flag=wx.EXPAND)

        sizer.Add(grid, flag=wx.ALL | wx.EXPAND, border=8)

        btn_sizer = wx.StdDialogButtonSizer()
        ok_btn = wx.Button(self, wx.ID_OK)
        cancel_btn = wx.Button(self, wx.ID_CANCEL)
        btn_sizer.AddButton(ok_btn)
        btn_sizer.AddButton(cancel_btn)
        btn_sizer.Realize()
        sizer.Add(btn_sizer, flag=wx.ALL | wx.ALIGN_CENTER, border=8)

        self.SetSizerAndFit(sizer)

    def get_values(self):
        bus_id = self.busid_ctrl.GetValue().strip()
        if not bus_id:
            bus_id = None
        description = self.description_ctrl.GetValue().strip()
        if not description:
            description = None
        instance_id = self.instanceid_ctrl.GetValue().strip()
        if not instance_id:
            instance_id = None
        return (
            bus_id,
            description,
            instance_id,
        )


class SettingsWindow(wx.Dialog):
    def __init__(self, parent: WslUsbGui):
        super().__init__(parent, title="Settings")
        self.parent = parent
        outer_panel = wx.Panel(self)

        # Create a sizer for layout
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Minimise to tray checkbox
        self.minimize_tray_checkbox = wx.CheckBox(outer_panel, label="Close to tray (run in background)")
        self.minimize_tray_checkbox.SetValue(parent.close_to_tray)  # Example

        # Auto-start checkbox
        self.auto_start_checkbox = wx.CheckBox(outer_panel, label="Auto-start with Windows")
        self.auto_start_checkbox.SetValue(parent.auto_start_at_boot)

        # Add checkboxes to sizer
        sizer.Add(self.minimize_tray_checkbox, 0, wx.ALL | wx.ALIGN_LEFT, border=8)
        sizer.Add(self.auto_start_checkbox, 0, wx.ALL | wx.ALIGN_LEFT, border=8)

        close_button = wx.Button(outer_panel, label="Close")
        sizer.AddSpacer(8)
        sizer.Add(close_button, 0, wx.ALIGN_CENTRE | wx.ALL, border=8)
        self.Bind(wx.EVT_BUTTON, self.Save, close_button)

        # Outer panel with border
        outer_sizer = wx.BoxSizer(wx.VERTICAL)
        outer_sizer.Add(outer_panel, 1, wx.ALL | wx.EXPAND, border=8)

        outer_panel.SetSizerAndFit(sizer)

        # Set the dialog sizer
        self.SetSizerAndFit(outer_sizer)


    def Save(self, _):
        need_save = False
        if self.minimize_tray_checkbox.Value != self.parent.close_to_tray:
            self.parent.close_to_tray = self.minimize_tray_checkbox.Value
            need_save = True

        if self.auto_start_checkbox.Value != self.parent.auto_start_at_boot:
            self.parent.auto_start_at_boot = self.auto_start_checkbox.Value
            self.parent.update_auto_start_at_boot()
            need_save = True

        if need_save:
            self.parent.save_config()

        self.Close()

async def amain():
    global gui

    if sys.argv[0] is None:
        # Compiled app, need to inject app path for argparse
        sys.argv[0] = sys.executable

    parser = argparse.ArgumentParser()
    parser.add_argument("--minimised", action="store_true", help="Start app minimised to tray")
    args = parser.parse_args()

    app = wxasync.WxAsyncApp(False)

    # Generate a unique instance name based on user ID and app dir path hash
    import hashlib
    path_hash = hashlib.md5(str(APP_DIR.resolve()).encode()).hexdigest()
    instance_name = f"wsl_usb_gui_{wx.GetUserId()}_{path_hash}"
    instance = wx.SingleInstanceChecker(instance_name)

    if instance.IsAnotherRunning():
        # Find the existing window by title and send a custom message to restore it.
        window_title = f"WSL USB Manager {__version__}"
        hwnd = ctypes.windll.user32.FindWindowW(None, window_title)
        if hwnd:
            ctypes.windll.user32.PostMessageW(hwnd, WM_SHOW_EXISTING, 0, 0)
            log.info(f"Found existing instance (HWND: {hwnd}), sent WM_SHOW_EXISTING message.")
        else:
            # Fallback if window not found
            log.warning(f"Another instance is running, but could not find window with title: {window_title}")
            wx.MessageBox(caption="Already running", message="Another instance of the app is already running,\ncould not bring it to front.", style=wx.OK | wx.ICON_WARNING)
        return # Exit the new instance

    gui = WslUsbGui(minimised=args.minimised)
    app.SetTopWindow(gui)

    await check_usbipd_version()

    bundled_ver = MSI_VERS if MSI_VERS > (0, 0, 0,) else (4, 0, 0)

    if gui.first_run != __version__ and USBIPD_VERSION < bundled_ver:
        log.warning("version upgrade detected, install deps")
        install_deps()
        await check_usbipd_version()

    gui.refresh(delay=0.5)

    # TODO
    devNotifyHandle = registerDeviceNotification(handle=gui.GetHandle(), callback=windows_events_callback)

    await gui.check_wsl_udev()

    await app.MainLoop()

    try:
        unregisterDeviceNotification(devNotifyHandle)
    except:
        pass


def main():
    asyncio.run(amain())


if __name__ == "__main__":
    main()
