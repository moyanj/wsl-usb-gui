import appdirs
import asyncio
import json
import re
import subprocess
import sys
import time
from collections import namedtuple
from functools import partial
from pathlib import Path
from typing import *

import wx
import wx.adv
import wxasync
from wx.lib.agw.ultimatelistctrl import (
    UltimateListCtrl,
    UltimateListItem,
    ULC_REPORT,
    ULC_MASK_KIND,
    ULC_MASK_TEXT,
    EVT_LIST_ITEM_CHECKED,
)

from .usb_monitor import registerDeviceNotification, unregisterDeviceNotification

# High DPI Support.
import ctypes

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(True)
except:
    pass


DEVICE_COLUMNS = ["bus_id", "description", "bound", "forced"]
ATTACHED_COLUMNS = ["bus_id", "description", "forced"]  # , "client"]
PROFILES_COLUMNS = ["bus_id", "description"]

USBIPD_PORT = 3240
APP_DIR = Path(appdirs.user_data_dir("wsl-usb-gui", False))
APP_DIR.mkdir(exist_ok=True)
CONFIG_FILE = APP_DIR / "config.json"

Device = namedtuple("Device", "BusId Description bound forced InstanceId Attached")
Profile = namedtuple("Profile", "BusId Description InstanceId", defaults=(None, None, None))

gui: Optional["WslUsbGui"] = None
loop = None

USBIPD_default = Path("C:\\Program Files\\usbipd-win\\usbipd.exe")
if USBIPD_default.exists():
    USBIPD = USBIPD_default
else:
    # try to run from anywhere on path, will try to install later if needed
    USBIPD = "usbipd"


def run(args):
    CREATE_NO_WINDOW = 0x08000000
    return subprocess.run(
        args,
        capture_output=True,
        encoding="UTF-8",
        creationflags=CREATE_NO_WINDOW,
        shell=(isinstance(args, str)),
    )


def get_icon():
    icon = Path(sys.executable).parent / "usb.ico"
    if not icon.exists():
        try:
            icon = Path(__file__).parent / "usb.ico"
        except:
            icon = None
    if icon is not None:
        return wx.Icon(str(icon), wx.BITMAP_TYPE_ICO)


class WslUsbGui(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, title="WSL USB Manager")

        self.icon = get_icon()

        if self.icon:
            self.SetIcon(self.icon)

        self.Bind(wx.EVT_CLOSE, self.OnClose)

        splitter_bottom = ProportionalSplitter(self, proportion=0.66, style=wx.SP_LIVE_UPDATE)
        splitter_top = ProportionalSplitter(
            splitter_bottom, proportion=0.33, style=wx.SP_LIVE_UPDATE
        )

        splitter_top.SetMinimumPaneSize(6)
        splitter_bottom.SetMinimumPaneSize(6)

        # On first close, alert the user it's being minimised to tray
        self.informed_about_tray = False

        self.usb_devices: Set[Device] = set()
        self.pinned_profiles: List[Profile] = []
        self.name_mapping = dict()
        self.refreshing = False

        headingFont = wx.Font(
            16, wx.FONTFAMILY_DECORATIVE, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD
        )

        ## TOP SECTION - Available USB Devices

        top_panel = wx.Panel(splitter_top, style=wx.SUNKEN_BORDER)
        top_sizer = wx.BoxSizer(wx.VERTICAL)
        top_panel.SetSizerAndFit(top_sizer)

        available_list_label = wx.StaticText(top_panel, label="Windows USB Devices")
        available_list_label.SetFont(headingFont)
        refresh_button = self.Button(top_panel, "Refresh", command=self.refresh)

        top_controls = wx.BoxSizer(wx.HORIZONTAL)
        top_controls.Add(available_list_label, 2, wx.EXPAND | wx.TOP | wx.LEFT, border=6)
        top_controls.AddStretchSpacer(1)
        top_controls.Add(refresh_button, 1, wx.TOP, border=6)

        self.available_listbox = ListCtrl(top_panel)
        self.available_listbox.InsertColumns(DEVICE_COLUMNS)

        def available_menu(event):
            popupmenu = wx.Menu()
            entries = [
                ("Attach to WSL", self.attach_wsl),
                ("Auto-Attach Device", self.auto_attach_wsl),
                ("Rename Device", self.rename_device),
                ("Bind", self.bind),
                ("Force Bind", self.force_bind),
                ("Unbind", self.unbind),
            ]
            for entry, fn in entries:
                menuItem = popupmenu.Append(-1, entry)
                self.Bind(wx.EVT_MENU, fn, menuItem)
            # Show menu
            self.PopupMenu(popupmenu)

        def _available_menu(event):
            wx.CallAfter(available_menu, event)

        self.available_listbox.Bind(wx.EVT_LIST_ITEM_RIGHT_CLICK, _available_menu)

        def available_checked(event):
            available_listbox = event.EventObject
            device: Device = available_listbox.devices[event.Index]
            bound = available_listbox.GetItem(event.Index, col=2).IsChecked()
            forced = available_listbox.GetItem(event.Index, col=3).IsChecked()
            if device.forced and not forced:
                self.unbind_bus_id(device.BusId)
            elif forced and not device.forced:
                self.bind_bus_id(device.BusId, forced=True)
            elif device.bound and not bound:
                self.unbind_bus_id(device.BusId)
            elif bound and not device.bound:
                self.bind_bus_id(device.BusId, forced=False)

        self.available_listbox.Bind(EVT_LIST_ITEM_CHECKED, available_checked)

        top_sizer.Add(top_controls, flag=wx.EXPAND | wx.ALL, border=6)
        top_sizer.Add(self.available_listbox, proportion=1, flag=wx.EXPAND | wx.ALL, border=6)

        ## MIDDLE SECTION - USB devices currently attached

        middle_panel = wx.Panel(splitter_top, style=wx.SUNKEN_BORDER)
        middle_sizer = wx.BoxSizer(wx.VERTICAL)
        middle_panel.SetSizerAndFit(middle_sizer)

        attached_list_label = wx.StaticText(middle_panel, label="Forwarded Devices")
        attached_list_label.SetFont(headingFont)

        attach_button = self.Button(middle_panel, "Attach ↓", command=self.attach_wsl)

        detach_button = self.Button(middle_panel, "↑ Detach", command=self.detach_wsl)
        auto_attach_button = self.Button(middle_panel, "Auto-Attach", command=self.auto_attach_wsl)
        rename_button = self.Button(middle_panel, "Rename", command=self.rename_device)

        middle_controls = wx.BoxSizer(wx.HORIZONTAL)
        middle_controls.Add(attached_list_label, 2, wx.EXPAND | wx.TOP | wx.LEFT, border=6)
        middle_controls.AddStretchSpacer(1)
        middle_controls.Add(attach_button, 1, wx.TOP, border=6)
        middle_controls.Add(detach_button, 1, wx.TOP, border=6)
        middle_controls.Add(auto_attach_button, 1, wx.TOP, border=6)
        middle_controls.Add(rename_button, 1, wx.TOP, border=6)

        self.attached_listbox = ListCtrl(middle_panel)
        self.attached_listbox.InsertColumns(ATTACHED_COLUMNS)

        def attached_menu(event):
            popupmenu = wx.Menu()
            entries = [
                ("Detach Device", self.detach_wsl),
                ("Auto-Attach Device", self.auto_attach_wsl),
                ("Rename Device", self.rename_device),
                ("WSL: Grant User Permissions", self.udev_permissive),
            ]
            for entry, fn in entries:
                menuItem = popupmenu.Append(-1, entry)
                self.Bind(wx.EVT_MENU, fn, menuItem)
            # Show menu
            self.PopupMenu(popupmenu)

        def _attached_menu(event):
            wx.CallAfter(attached_menu, event)

        self.attached_listbox.Bind(wx.EVT_LIST_ITEM_RIGHT_CLICK, _attached_menu)

        def attached_checked(event):
            attached_listbox = event.EventObject
            device: Device = attached_listbox.devices[event.Index]
            forced = attached_listbox.GetItem(event.Index, col=2).IsChecked()
            if device.forced and not forced:
                self.unbind_bus_id(device.BusId)
            elif forced and not device.forced:
                self.bind_bus_id(device.BusId, forced=True)
            return True

        self.attached_listbox.Bind(EVT_LIST_ITEM_CHECKED, attached_checked)

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

        bottom_controls = wx.BoxSizer(wx.HORIZONTAL)
        bottom_controls.Add(pinned_list_label, 2, wx.EXPAND | wx.TOP | wx.LEFT, border=6)
        bottom_controls.AddStretchSpacer(1)
        bottom_controls.Add(pinned_list_delete_button, 1, wx.TOP, border=6)

        self.pinned_listbox = ListCtrl(bottom_panel)
        self.pinned_listbox.InsertColumns(PROFILES_COLUMNS)

        def profile_menu(event):
            popupmenu = wx.Menu()
            entries = [
                ("Delete Profile", self.delete_profile),
            ]
            for entry, fn in entries:
                menuItem = popupmenu.Append(-1, entry)
                self.Bind(wx.EVT_MENU, fn, menuItem)

            self.PopupMenu(popupmenu)  # , position)

        def _profile_menu(event):
            wx.CallAfter(profile_menu, event)

        self.pinned_listbox.Bind(wx.EVT_LIST_ITEM_RIGHT_CLICK, _profile_menu)

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

        self.load_config()
        self.refresh()

        self.Show(True)
        self.SetSize(self.FromDIP(wx.Size(600, 800)))

    def Button(self, parent, button_text, command):
        btn = wx.Button(parent, label=button_text)
        btn.SetMaxSize(parent.FromDIP(wx.Size(90, 30)))
        self.Bind(wx.EVT_BUTTON, lambda event: command(), btn)
        return btn

    def OnClose(self, event):
        if event.CanVeto():
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
            event.Skip()
            self.Destroy()

    @staticmethod
    def create_profile(busid, description, instanceid):
        return Profile(*(None if a == "None" else a for a in (busid, description, instanceid)))

    def load_config(self):
        try:
            print(f"Loading config from: {CONFIG_FILE}")
            config = json.loads(CONFIG_FILE.read_text())
            if isinstance(config, list):
                self.pinned_profiles = [self.create_profile(*c) for c in config]
            else:
                self.pinned_profiles = [self.create_profile(*c) for c in config["pinned_profiles"]]
                self.name_mapping = config["name_mapping"]
                self.informed_about_tray = config.get("informed_about_tray", False)

        except Exception as ex:
            pass

    def save_config(self):
        config = dict(
            pinned_profiles=self.pinned_profiles,
            name_mapping=self.name_mapping,
            informed_about_tray=self.informed_about_tray,
        )
        CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
        CONFIG_FILE.write_text(json.dumps(config, indent=4, sort_keys=True))

    def parse_state(self, text) -> List[Device]:
        rows = []
        devices = json.loads(text)

        for device in devices["Devices"]:
            bus_info = device["BusId"]
            if bus_info:
                instanceId = device["InstanceId"]
                description = self.name_mapping.get(instanceId, device["Description"])
                # bind = "☒" if device["PersistedGuid"] else "☐"
                # forced = "☒" if device["IsForced"] else "☐"
                bind = True if device["PersistedGuid"] else False
                forced = True if device["IsForced"] else False
                attached = device["ClientIPAddress"]
                rows.append(Device(str(bus_info), description, bind, forced, instanceId, attached))
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

    def list_wsl_usb(self) -> List[Device]:
        global loop
        try:
            result = run([USBIPD, "state"])
            return self.parse_state(result.stdout)
        except Exception as ex:
            if isinstance(ex, FileNotFoundError):
                install_deps()
            return []

    @staticmethod
    def usb_ipd_run_admin_if_needed(command, msg=None):
        result = run(command)
        if "error:" in result.stderr and "administrator privileges" in result.stderr:
            if msg:
                wx.MessageBox(
                    caption="Administrator Privileges",
                    message=msg,
                    style=wx.OK | wx.ICON_INFORMATION,
                )
            args_str = ", ".join(f'\\"{arg}\\"' for arg in command[1:])

            result = run(
                r'''Powershell -Command "& { Start-Process \"%s\" -ArgumentList @(%s) -Verb RunAs } "'''
                % (USBIPD, args_str)
            )
        return result

    @staticmethod
    def bind_bus_id(bus_id, forced):
        command = [USBIPD, "bind", f"--busid={bus_id}"]
        if forced:
            command.append("--force")
        result = WslUsbGui.usb_ipd_run_admin_if_needed(command)
        if result.stdout:
            print(result.stdout)
        if result.stderr:
            print(result.stderr)
        print(f"Bind {bus_id}: {'Success' if not result.returncode else 'Failed'}")
        return result

    @staticmethod
    def unbind_bus_id(bus_id):
        command = [USBIPD, "unbind", f"--busid={bus_id}"]
        result = WslUsbGui.usb_ipd_run_admin_if_needed(command)
        if result.stdout:
            print(result.stdout)
        if result.stderr:
            print(result.stderr)
        print(f"Unbind {bus_id}: {'Success' if not result.returncode else 'Failed'}")
        return result

    @staticmethod
    def attach_wsl_usb(bus_id):
        command = [USBIPD, "wsl", "attach", "--busid=" + bus_id]
        msg = (
            "The first time attaching a device to WSL requires elevated privileges; "
            "subsequent attaches will succeed with standard user privileges."
        )
        result = WslUsbGui.usb_ipd_run_admin_if_needed(command, msg)
        if result.stdout:
            print(result.stdout)
        if result.stderr:
            print(result.stderr)

        if "client not correctly installed" in result.stderr.lower():
            install_deps()

        elif "is already attached to a client." in result.stderr.lower():
            # Not an error, we've just tried to attach twice.
            pass
        elif "error:" in result.stderr.lower():
            err = [l for l in result.stderr.lower().split("\n") if "error:" in l][0].strip()

            wx.MessageBox(caption="Failed to attach", message=err, style=wx.OK | wx.ICON_WARNING)

        return result

    @staticmethod
    def detach_wsl_usb(bus_id):
        result = run([USBIPD, "wsl", "detach", "--busid=" + str(bus_id)])
        if result.stdout:
            print(result.stdout)
        if result.stderr:
            print(result.stderr)

    def update_pinned_listbox(self):
        self.pinned_listbox.DeleteAllItems()
        for profile in self.pinned_profiles:
            busid = str(profile.BusId)
            desc = str(profile.Description or self.lookup_description(profile.InstanceId))
            instanceId = str(profile.InstanceId)
            self.pinned_listbox.Append((busid, desc, instanceId))

    # Define a function to implement choice function
    def auto_attach_wsl_choice(self, profile):
        self.pinned_profiles.append(profile)
        self.save_config()
        self.update_pinned_listbox()

    def remove_pinned_profile(self, busid, description, instanceid):
        profile = self.create_profile(busid, description, instanceid)
        for i, p in enumerate(list(self.pinned_profiles)):
            if (p.BusId and p.BusId == profile.BusId) or (
                p.InstanceId and p.InstanceId == profile.InstanceId
            ):
                self.pinned_profiles.remove(p)

    def delete_profile(self, event=None):
        selection = self.pinned_listbox.GetFirstSelected()
        if selection == -1:
            print("no selection to delete")
            return  # no selected item
        device = self.pinned_listbox.devices[selection]
        self.pinned_listbox.DeleteItem(selection)
        busid, description, instanceid = device  # type: ignore
        self.remove_pinned_profile(busid, description, instanceid)
        self.save_config()

    def get_selection(self, available=False, attached=False) -> Optional[Device]:
        if not available or attached:
            # If nether specified, return both
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

    def get_selected_device(self, available=False, attached=False):
        selection = self.get_selection(available, attached)
        if not selection:
            print("no selection to rename")
            return

        device = [
            d
            for d in self.usb_devices
            if d.BusId == selection.BusId and d.Description == selection.Description
        ][0]
        return device

    def rename_device(self, event=None):
        device = self.get_selected_device
        if not device:
            return

        instanceId = device.InstanceId

        current = self.name_mapping.get(instanceId, device.Description)
        caption = "Rename"
        message = f"Enter new label for port: {device.BusId}\nOr leave blank to reset to default."
        dlg = wx.TextEntryDialog(self, message, caption, current)
        newname = None
        if dlg.ShowModal() == wx.ID_OK:
            newname = dlg.GetValue()

        if newname is None:
            # Cancel
            return

        if newname:
            self.name_mapping[instanceId] = newname
        else:
            try:
                self.name_mapping.pop(instanceId)
            except:
                pass
        self.save_config()
        self.refresh()

    def udev_permissive(self, event=None):
        device = self.get_selected_device()
        if not device:
            return

        try:
            vid = re.search("vid_([0-9a-f]+)&", device.InstanceId.lower()).group(1)
            pid = re.search("pid_([0-9a-f]+)\\\\", device.InstanceId.lower()).group(1)
            udev_rule = f'SUBSYSTEM=="usb", ATTRS{{idVendor}}=="{vid}", ATTRS{{idProduct}}=="{pid}", MODE="0666"'
            rules_file = "/etc/udev/rules.d/99-wsl-usb-gui.rules"
            udev_settings = run(
                [
                    "wsl",
                    "--user",
                    "root",
                    "sh",
                    "-c",
                    f"grep -q '{udev_rule}' {rules_file} || (echo '{udev_rule}' >> {rules_file}; sudo udevadm control --reload-rules; sudo udevadm trigger)",
                ]
            ).stdout.strip()
            print(f"udev rule added: {udev_rule}")
            wx.MessageBox(
                caption="WSL: Grant User Permissions",
                message=f"WSL udev rule added for VID:{vid} PID:{pid}.",
                style=wx.OK | wx.ICON_INFORMATION,
            )
        except AttributeError as ex:
            print("Could not get device information for udev: ", ex)
            wx.MessageBox(
                caption="WSL: Grant User Permissions",
                message=f"ERROR: Failed to add udev rule.",
                style=wx.OK | wx.ICON_WARNING,
            )

    async def refresh_task(self, delay: float = 0):
        try:
            if self.refreshing:
                return
            self.refreshing = True
            if delay:
                await asyncio.sleep(delay)

            print("Refresh USB")

            usb_devices = set(
                await asyncio.get_running_loop().run_in_executor(None, self.list_wsl_usb)
            )

            new_devices = usb_devices - self.usb_devices
            self.usb_devices = usb_devices

            if not self.usb_devices:
                return

            self.attached_listbox.DeleteAllItems()
            self.available_listbox.DeleteAllItems()
            for device in sorted(self.usb_devices, key=lambda d: d.BusId):
                new = device in new_devices
                if device.Attached:
                    row = self.attached_listbox.Append(device)
                    if new:
                        self.attached_listbox.HighlightRow(row)
                else:
                    if self.attach_if_pinned(device):
                        row = self.attached_listbox.Append(device)
                        if new:
                            self.attached_listbox.HighlightRow(row)
                    else:
                        row = self.available_listbox.Append(device)
                        if new:
                            self.available_listbox.HighlightRow(row)

            self.update_pinned_listbox()
        finally:
            self.refreshing = False

    def highlight_row(self, listbox: "ListCtrl", row: int):
        pass

    def refresh(self, delay=0.0):
        asyncio.get_running_loop().call_soon_threadsafe(
            asyncio.ensure_future, self.refresh_task(delay)
        )

    def check_wsl_udev(self):
        # Autostart WSL udev service if needed
        udev_start = run(
            [
                "wsl",
                "--user",
                "root",
                "sh",
                "-c",
                "pgrep udev || (echo 'starting udev'; service udev restart)",
            ]
        ).stdout.strip()
        udev_start = udev_start.replace("\n", ", ")
        print(f"udev: {udev_start}")

    def lookup_description(self, instanceId):
        if not instanceId:
            return None

        if instanceId == "None":
            return None

        for device in self.usb_devices:
            if device.InstanceId == instanceId:
                return device.Description

    def attach_if_pinned(self, device):
        for busid, desc, instanceId in self.pinned_profiles:
            if instanceId or busid:
                # Only fallback to description if no other filter set
                desc = None

            if busid and device.BusId.strip() != busid.strip():
                continue
            if instanceId and device.InstanceId != instanceId:
                continue
            if desc and device.Description.strip() != desc.strip():
                continue
            self.attach_wsl_usb(device.BusId)
            return True
        return False

    def _attach_selection_busid(self):
        device = self.get_selection(available=True)
        if not device:
            print("no selection to attach")
            return
        return device.BusId

    def force_bind(self, event=None):
        bus_id = self._attach_selection_busid()
        print(f"Bind (forced) {bus_id}")
        result = self.bind_bus_id(bus_id, forced=True)
        print(f"Bind (forced) {bus_id}: {'Success' if not result.returncode else 'Failed'}")
        self.refresh()
        self.refresh()

    def bind(self, event=None):
        bus_id = self._attach_selection_busid()
        result = self.bind_bus_id(bus_id, forced=False)

    def unbind(self, event=None):
        bus_id = self._attach_selection_busid()
        self.unbind_bus_id(bus_id)

    def attach_wsl(self, event=None):
        bus_id = self._attach_selection_busid()
        result = self.attach_wsl_usb(bus_id)
        print(f"Attach {bus_id}: {'Success' if not result.returncode else 'Failed'}")
        time.sleep(0.5)
        self.refresh()

    def detach_wsl(self, event=None):
        device = self.get_selection(attached=True)
        if not device:
            print("no selection to detach")
            return  # no selected item
        print(f"Detach {device.BusId} {device.Description}")

        self.remove_pinned_profile(device.BusId, device.Description, device.InstanceId)

        self.detach_wsl_usb(device.BusId)

        time.sleep(0.5)
        self.refresh()

    def auto_attach_wsl(self, event=None):
        global pop
        device = self.get_selection()
        if not device:
            print("no selection to create profile for")
            return

        popup = popupAutoAttach(
            self, device.BusId, device.Description, device.InstanceId, self.icon
        )
        popup.ShowModal()

        self.refresh()


class ListCtrl(UltimateListCtrl):
    def __init__(self, parent, *args, **kw):
        UltimateListCtrl.__init__(self, parent, wx.ID_ANY, agwStyle=ULC_REPORT)

        self.devices: List[Device] = []
        self.columns = []

    def HighlightRow(self, index):
        original = self.GetItemBackgroundColour(index)
        self.SetItemBackgroundColour(index, wx.YELLOW)
        wx.CallLater(2000, self.SetItemBackgroundColour, index, original)

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

    def Append(self, device):
        details = device[0 : UltimateListCtrl.GetColumnCount(self)]

        pos = self.GetItemCount()
        self.InsertStringItem(pos, details[0])
        for i in range(1, len(details)):
            if i == 1:
                self.SetStringItem(pos, i, str(details[i]))
            else:
                # Checkbox column
                column = self.columns[i]
                value = device._asdict()[column]
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

    def choice(self, parent, profile, event):
        parent.auto_attach_wsl_choice(profile)
        self.Close()


def usb_callback(attach):
    if attach:
        print(f"USB device attached")
    else:
        print(f"USB device detached")
    if gui:
        gui.refresh(0.5)


def install_deps():
    asyncio.get_event_loop().call_soon_threadsafe(_install_deps)


def _install_deps():
    global USBIPD, gui
    rsp = wx.MessageBox(
        caption="Install Dependencies",
        message=(
            "Some of the dependencies are missing, install them now?\n"
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

    def OnExit(self, event):
        wx.CallAfter(self.frame.Close, True)
        self.RemoveIcon()
        self.Destroy()



async def amain():
    global gui

    app = wxasync.WxAsyncApp(False)

    instance = wx.SingleInstanceChecker(f"wsl_usb_gui_{wx.GetUserId()}", str(APP_DIR.resolve()))
    if instance.IsAnotherRunning():
        wx.MessageBox(caption="Already running", message="Another instance of the app is already running,\ncheck system tray icon to restore instance.", style=wx.OK | wx.ICON_WARNING)
        return

    gui = WslUsbGui()
    app.SetTopWindow(gui)
    taskbar = TaskBarIcon(gui, gui.icon)

    # TODO
    devNotifyHandle = registerDeviceNotification(handle=gui.GetHandle(), callback=usb_callback)

    gui.check_wsl_udev()

    await app.MainLoop()

    try:
        unregisterDeviceNotification(devNotifyHandle)
    except:
        pass


def main():
    asyncio.run(amain())


if __name__ == "__main__":
    main()
