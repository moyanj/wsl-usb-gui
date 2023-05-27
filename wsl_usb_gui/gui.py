# requires python 3.8+ may work on 3.6, 3.7 definitely broken on <= 3.5 due to subprocess args (text=True)
import asyncio

# import PySimpleGUIWx as sg

import wx
import wxasync

# from wx.lib.splitter import MultiSplitterWindow
# from wx.lib.mixins.listctrl import ListCtrlAutoWidthMixin
from wx.lib.agw.ultimatelistctrl import (
    UltimateListCtrl,
    UltimateListItem,
    ULC_REPORT,
    ULC_MASK_KIND,
    ULC_MASK_TEXT,
    EVT_LIST_ITEM_CHECKED,
)

# from tkinter import *
# from tkinter.ttk import *
# from tkinter import simpledialog, PanedWindow
import json
import subprocess
import time
from collections import namedtuple
from pathlib import Path
from typing import *
import appdirs
from functools import partial
from .usb_monitor import registerDeviceNotification, unregisterDeviceNotification

import ctypes

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(True)
except:
    pass


# sg.theme('SystemDefaultForReal')


mod_dir = Path(__file__).parent
ICON_PATH = str(mod_dir / "usb.ico")

DEVICE_COLUMNS = ["bus_id", "description", "shared", "forced"]
ATTACHED_COLUMNS = ["bus_id", "description", "forced"]  # , "client"]
PROFILES_COLUMNS = ["bus_id", "description"]

USBIPD_PORT = 3240
CONFIG_FILE = Path(appdirs.user_data_dir("wsl-usb-gui", "")) / "config.json"


Device = namedtuple("Device", "BusId Description shared forced InstanceId Attached time")
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


class WslUsbGui(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, title="WSL USB Manager")
        self.SetIcon(wx.Icon(ICON_PATH, wx.BITMAP_TYPE_ICO))

        # splitter = MultiSplitterWindow(self, style=wx.SP_THIN_SASH|wx.SP_BORDER)
        # splitter.SetOrientation(wx.VERTICAL)
        # splitter.SetBackgroundColour(wx.LIGHT_GREY)

        splitter_bottom = ProportionalSplitter(self, proportion=0.66, style=wx.SP_LIVE_UPDATE)
        splitter_top = ProportionalSplitter(
            splitter_bottom, proportion=0.33, style=wx.SP_LIVE_UPDATE
        )

        splitter_top.SetMinimumPaneSize(6)
        splitter_bottom.SetMinimumPaneSize(6)

        # splitter_top.SetSashGravity(0.5)
        # splitter_top.SetSashSize
        # self.tkroot = Tk()
        # self.tkroot.wm_title("WSL USB Manager")
        # self.tkroot.geometry("600x800")
        # self.tkroot.iconbitmap(ICON_PATH)

        # self.pw = PanedWindow(orient="vertical", showhandle=False, sashwidth=6, sashrelief='groove')

        self.usb_devices: Set[Device] = set()
        self.pinned_profiles: List[Profile] = []
        self.name_mapping = dict()

        headingFont = wx.Font(
            16, wx.FONTFAMILY_DECORATIVE, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD
        )
        # listCtrlFont = wx.Font(11, wx.FONTFAMILY_DECORATIVE, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)

        ## TOP SECTION - Available USB Devices
        # top_frame = Frame(self.pw)
        # available_control_frame = Frame(top_frame)
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

        # available_listbox_frame = Frame(top_frame)

        # available_listbox_scroll = Scrollbar(available_listbox_frame)
        # available_listbox_scroll.configure(command=self.available_listbox.yview)
        # self.available_listbox.configure(yscrollcommand=available_listbox_scroll.set)

        self.available_listbox = ListCtrl(top_panel)
        self.available_listbox.InsertColumns(DEVICE_COLUMNS)

        # self.available_listbox.setResizeColumn(2)
        # self.available_listbox.SetFont(listCtrlFont)

        def available_menu(event):
            # item = event.GetItem()
            # itemData = self.available_listbox.GetItemData(item).GetData()
            # itemData = self.available_listbox.devices[item._itemId]

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
            # popupmenu.UpdateUI()
            # self.attached_listbox.Refresh()
            # self.attached_listbox.Update()
            self.PopupMenu(popupmenu)  # , self.FromDIP(event.GetPoint()))

        def _available_menu(event):
            wx.CallAfter(available_menu, event)

        self.available_listbox.Bind(wx.EVT_LIST_ITEM_RIGHT_CLICK, _available_menu)

        def available_checked(event):
            available_listbox = event.EventObject
            device: Device = available_listbox.devices[event.Index]
            shared = available_listbox.GetItem(event.Index, col=2).IsChecked()
            forced = available_listbox.GetItem(event.Index, col=3).IsChecked()
            if device.forced and not forced:
                self.unbind_bus_id(device.BusId)
            elif forced and not device.forced:
                self.bind_bus_id(device.BusId, forced=True)
            elif device.shared and not shared:
                self.unbind_bus_id(device.BusId)
            elif shared and not device.shared:
                self.bind_bus_id(device.BusId, forced=False)

        self.available_listbox.Bind(EVT_LIST_ITEM_CHECKED, available_checked)

        top_sizer.Add(top_controls, flag=wx.EXPAND | wx.ALL, border=6)
        top_sizer.Add(self.available_listbox, proportion=1, flag=wx.EXPAND | wx.ALL, border=6)

        # self.available_listbox.bind("<Button-3>", partial(self.do_listbox_menu, listbox=self.available_listbox, menu=available_menu))

        # for i, col in enumerate(DEVICE_COLUMNS):
        #     self.available_listbox.heading(col, text=col.title())
        #     if i < 2:
        #         self.available_listbox.column(
        #             col, minwidth=40, width=50, anchor=W if i else CENTER, stretch=TRUE if i else FALSE
        #         )
        #     else:
        #         self.available_listbox.column(
        #             col, minwidth=50, width=50, anchor=CENTER, stretch=FALSE
        #         )

        # available_list_label.grid(column=0, row=0, padx=5)
        # refresh_button.grid(column=2, row=0, sticky=E, padx=5)

        # available_control_frame.rowconfigure(0, weight=1)
        # available_control_frame.columnconfigure(1, weight=1)
        # available_control_frame.grid(column=0, row=0, sticky=N + W + E, pady=10, padx=10)

        # available_listbox_frame.grid(column=0, row=1, sticky=W + E + N + S, pady=10, padx=10)
        # available_listbox_frame.rowconfigure(0, weight=1)
        # available_listbox_frame.columnconfigure(0, weight=1)
        # self.available_listbox.grid(column=0, row=0, sticky=W + E + N + S)
        # available_listbox_scroll.grid(column=1, row=0, sticky=W + N + S)

        ## MIDDLE SECTION - USB devices currently attached
        middle_panel = wx.Panel(splitter_top, style=wx.SUNKEN_BORDER)
        middle_sizer = wx.BoxSizer(wx.VERTICAL)
        middle_panel.SetSizerAndFit(middle_sizer)
        # middle_frame = Frame(self.pw)

        # control_frame = Frame(middle_frame)
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

        # attached_listbox_frame = Frame(middle_frame)

        # attached_listbox_scroll = Scrollbar(attached_listbox_frame)
        # attached_listbox_scroll.configure(command=self.attached_listbox.yview)
        # self.attached_listbox.configure(yscrollcommand=attached_listbox_scroll.set)

        self.attached_listbox = ListCtrl(middle_panel)
        self.attached_listbox.InsertColumns(ATTACHED_COLUMNS)

        def attached_menu(event):
            # item = event.GetItem()
            # itemData = self.attached_listbox.GetItemData(item).GetData()
            # itemData = self.attached_listbox.devices[item._itemId]

            popupmenu = wx.Menu()
            entries = [
                ("Detach from WSL", self.detach_wsl),
                ("Auto-Attach Device", self.auto_attach_wsl),
                ("Rename Device", self.rename_device),
            ]
            for entry, fn in entries:
                menuItem = popupmenu.Append(-1, entry)
                self.Bind(wx.EVT_MENU, fn, menuItem)

            # Show menu
            # event.Skip()
            # popupmenu.UpdateUI()
            # self.attached_listbox.Refresh()
            # self.attached_listbox.Update()
            self.PopupMenu(popupmenu)  # , self.FromDIP(event.GetPoint()))

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

        # for i, col in enumerate(ATTACHED_COLUMNS):
        #     self.attached_listbox.heading(col, text=col.title())
        #     self.attached_listbox.column(
        #         col, minwidth=40, width=50, anchor=W if i else CENTER, stretch=TRUE if i else FALSE
        #     )

        # attached_list_label.grid(column=0, row=0, padx=5)

        # attach_button.grid(column=2, row=0, padx=5)
        # detach_button.grid(column=3, row=0, padx=5)
        # auto_attach_button.grid(column=4, row=0, padx=5)
        # rename_button.grid(column=5, row=0, padx=5)

        # control_frame.rowconfigure(0, weight=1)
        # control_frame.columnconfigure(1, weight=1)
        # control_frame.grid(column=0, row=0, sticky=N + E + W, pady=10, padx=10)

        # attached_listbox_frame.grid(column=0, row=1, sticky=W + E + N + S, pady=10, padx=10)
        # attached_listbox_frame.rowconfigure(0, weight=1)
        # attached_listbox_frame.columnconfigure(0, weight=1)
        # self.attached_listbox.grid(column=0, row=0, sticky=W + E + N + S)
        # attached_listbox_scroll.grid(column=1, row=0, sticky=W + N + S)

        ## BOTTOM SECTION - saved profiles for auto-attach
        bottom_panel = wx.Panel(splitter_bottom, style=wx.SUNKEN_BORDER)
        bottom_sizer = wx.BoxSizer(wx.VERTICAL)
        bottom_panel.SetSizerAndFit(bottom_sizer)
        # bottom_frame = Frame(self.pw)

        # pinned_control_frame = Frame(bottom_frame)
        pinned_list_label = wx.StaticText(bottom_panel, label="Auto-Attach Profiles")
        pinned_list_label.SetFont(headingFont)
        pinned_list_delete_button = self.Button(
            bottom_panel, "Delete Profile", command=self.delete_profile
        )

        bottom_controls = wx.BoxSizer(wx.HORIZONTAL)
        bottom_controls.Add(pinned_list_label, 2, wx.EXPAND | wx.TOP | wx.LEFT, border=6)
        bottom_controls.AddStretchSpacer(1)
        bottom_controls.Add(pinned_list_delete_button, 1, wx.TOP, border=6)

        # pinned_listbox_frame = Frame(bottom_frame)

        # pinned_listbox_scroll = Scrollbar(pinned_listbox_frame)
        # pinned_listbox_scroll.configure(command=self.pinned_listbox.yview)
        # self.pinned_listbox.configure(yscrollcommand=pinned_listbox_scroll.set)

        # pinned_menu = [
        #     '&Right', [
        #         self.Menu(label="Delete Profile", command=self.delete_profile)
        #     ]
        # ]

        self.pinned_listbox = ListCtrl(bottom_panel)
        self.pinned_listbox.InsertColumns(PROFILES_COLUMNS)

        def profile_menu(event):
            # item = event.GetItem()

            popupmenu = wx.Menu()
            entries = [
                ("Delete Profile", self.delete_profile),
            ]
            for entry, fn in entries:
                menuItem = popupmenu.Append(-1, entry)
                self.Bind(wx.EVT_MENU, fn, menuItem)

            # Show menu
            # position = self.FromDIP(self.pinned_listbox.GetScreenPosition() + event.GetPoint())
            # popupmenu.UpdateUI()
            # self.attached_listbox.Refresh()
            # self.attached_listbox.Update()
            self.PopupMenu(popupmenu)  # , position)

        def _profile_menu(event):
            wx.CallAfter(profile_menu, event)

        self.pinned_listbox.Bind(wx.EVT_LIST_ITEM_RIGHT_CLICK, _profile_menu)

        bottom_sizer.Add(bottom_controls, flag=wx.EXPAND | wx.ALL, border=6)
        bottom_sizer.Add(self.pinned_listbox, proportion=1, flag=wx.EXPAND | wx.ALL, border=6)

        # setup column names
        # for i, col in enumerate(PROFILES_COLUMNS):
        #     self.pinned_listbox.heading(col, text=col.title())
        #     self.pinned_listbox.column(
        #         col, minwidth=40, width=50, anchor=W if i else CENTER, stretch=TRUE if i else FALSE
        #     )

        # pinned_list_label.grid(column=0, row=0, padx=5)
        # pinned_list_delete_button.grid(column=2, row=0, padx=5)

        # pinned_control_frame.rowconfigure(0, weight=1)
        # pinned_control_frame.columnconfigure(1, weight=1)
        # pinned_control_frame.grid(column=0, row=0, sticky=N + W + E, pady=10, padx=10)

        # pinned_listbox_frame.grid(column=0, row=1, sticky=W + E + N + S, pady=10, padx=10)
        # pinned_listbox_frame.rowconfigure(0, weight=1)
        # pinned_listbox_frame.columnconfigure(0, weight=1)
        # self.pinned_listbox.grid(column=0, row=0, sticky=W + E + N + S)
        # pinned_listbox_scroll.grid(column=1, row=0, sticky=W + N + S)

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
        ## Window Configure

        # top_frame.pack(fill="both", expand=True)
        # top_frame.add(available_control_frame)
        # top_frame.add(available_listbox_frame)

        # top_frame.columnconfigure(0, weight=1)
        # top_frame.rowconfigure(1, weight=1)
        # middle_frame.columnconfigure(0, weight=1)
        # middle_frame.rowconfigure(1, weight=1)
        # bottom_frame.columnconfigure(0, weight=1)
        # bottom_frame.rowconfigure(1, weight=1)

        # self.pw.pack(fill="both", expand=True)
        # self.pw.add(top_frame)
        # self.pw.add(middle_frame)
        # self.pw.add(bottom_frame)

        # self.pw.columnconfigure(0, weight=1)
        # self.pw.rowconfigure(0, weight=1)
        # self.pw.rowconfigure(1, weight=1)
        # self.pw.rowconfigure(2, weight=1)

        # self = sg.Window(
        #     title='WSL USB Manager',
        #     layout=top_layout + middle_layout + bottom_layout,
        #     size=(600,800),
        #     icon=ICON_PATH,
        #     resizable=True,
        # )

        self.load_config()

        self.refresh()

        # splitter.AppendWindow(top_panel)
        # splitter.AppendWindow(middle_panel)
        # splitter.AppendWindow(bottom_panel)

        splitter_top.SplitHorizontally(top_panel, middle_panel)
        splitter_bottom.SplitHorizontally(splitter_top, bottom_panel)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(splitter_bottom, proportion=1, flag=wx.EXPAND)

        sizer.SetSizeHints(self)
        self.SetSizerAndFit(sizer)

        self.Show(True)
        self.SetSize(self.FromDIP(wx.Size(600, 800)))

    def Button(self, parent, button_text, command):
        btn = wx.Button(parent, label=button_text)
        btn.SetMaxSize(parent.FromDIP(wx.Size(90, 30)))
        self.Bind(wx.EVT_BUTTON, lambda event: command(), btn)
        return btn

    def Menu(self, label, command):
        key = f"_menu_{label.replace(' ', '_')}"
        return label

    # def _CreateCheckBoxBitmap(self, flag=0, size=(16, 16)):
    #     """Create a bitmap of the platforms native checkbox. The flag
    #     is used to determine the checkboxes state (see wx.CONTROL_*)

    #     """
    #     bmp = wx.Bitmap(*size)
    #     dc = wx.MemoryDC(bmp)
    #     dc.SetBackground(wx.WHITE_BRUSH)
    #     dc.Clear()
    #     wx.RendererNative.Get().DrawCheckBox(self, dc,
    #                                          (0, 0, size[0], size[1]), flag)
    #     dc.SelectObject(wx.NullBitmap)
    #     return bmp

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

        except Exception as ex:
            pass

    def save_config(self):
        config = dict(
            pinned_profiles=self.pinned_profiles,
            name_mapping=self.name_mapping,
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
                rows.append(
                    Device(
                        str(bus_info), description, bind, forced, instanceId, attached, time.time()
                    )
                )
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
                loop.call_soon_threadsafe(install_deps)
            return None

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
            loop.call_soon_threadsafe(install_deps)

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
        self.remove_pinned_profile(device.BusId, device.Description, device.InstanceId)
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

    def rename_device(self, event=None):
        selection = self.get_selection()
        if not selection:
            print("no selection to rename")
            return

        # busid, description, *args = selection["values"]

        device = [
            d
            for d in self.usb_devices
            if d.BusId == selection.BusId and d.Description == selection.Description
        ][0]

        instanceId = device.InstanceId

        current = self.name_mapping.get(instanceId, selection.Description)
        caption = "Rename"
        message = f"Enter new label for port: {device.BusId}\nOr leave blank to reset to default."
        dlg = wx.TextEntryDialog(self, message, caption, current)
        newname = None
        if dlg.ShowModal() == wx.ID_OK:
            newname = dlg.GetValue()

        # getnewname = popupTextEntry(self, device.BusId, current)
        # self.tkroot.wait_window(getnewname.root)
        # newname = getnewname.value

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

    async def refresh_task(self, delay=0):
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
            # self.refresh(delay=3000)  # Gets refreshed on USB change
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
        # self.refresh()

    def unbind(self, event=None):
        bus_id = self._attach_selection_busid()
        self.unbind_bus_id(bus_id)
        # self.refresh()

    def attach_wsl(self, event=None):
        bus_id = self._attach_selection_busid()
        result = self.attach_wsl_usb(bus_id)
        print(f"Attach {bus_id}: {'Success' if not result.returncode else 'Failed'}")
        """if result.returncode == 0:
            attached_devices[bus_id] = {
                'bus_id' : bus_id,
                'port' : len(attached_devices),
                'description' : description
            }
        print(attached_devices)
        """
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

        popup = popupAutoAttach(self, device.BusId, device.Description, device.InstanceId)
        popup.ShowModal()
        # self.tkroot.wait_window(popup.root)

        self.refresh()

    # def do_listbox_menu(self, event, listbox, menu):
    #     try:
    #         listbox.selection_clear()
    #         iid = listbox.identify_row(event.y)
    #         if iid:
    #             listbox.selection_set(iid)
    #             menu.tk_popup(event.x_root, event.y_root)
    #     finally:
    #         menu.grab_release()


class ListCtrl(UltimateListCtrl):
    def __init__(self, parent, *args, **kw):
        # wx.ListCtrl.__init__(self, parent, wx.ID_ANY, style=wx.LC_REPORT)
        # ListCtrlAutoWidthMixin.__init__(self)
        UltimateListCtrl.__init__(self, parent, wx.ID_ANY, agwStyle=ULC_REPORT)

        self.devices: List[Device] = []
        self.columns = []

    def HighlightRow(self, index):
        original = self.GetItemBackgroundColour(index)
        self.SetItemBackgroundColour(index, wx.YELLOW)
        wx.FutureCall(2000, self.SetItemBackgroundColour, index, original)

    def InsertColumns(self, names):
        for i, name in enumerate(names):
            info = UltimateListItem()
            # if i > 1:
            info._mask = (
                wx.LIST_MASK_TEXT
            )  # | wx.LIST_MASK_IMAGE | wx.LIST_MASK_FORMAT | ULC_MASK_CHECK
            # else:
            # info._mask = wx.LIST_MASK_TEXT
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
        ##hack to set sizes on first paint event
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
    def __init__(self, parent, bus_id, description, instanceId):
        super().__init__(parent, title="New Auto-Attach Profile")
        self.SetIcon(wx.Icon(ICON_PATH, wx.BITMAP_TYPE_ICO))

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
    global USBIPD
    rsp = wx.MessageBox(
        caption="Install Dependencies",
        message=(
            "Some of the dependencies are missing, install them now?\n"
            "Note: All WSL instances may need to be restarted."
        ),
        style=wx.YES_NO | wx.ICON_WARNING,
    )
    if rsp == wx.ID_YES:
        from .install import install_task

        rsp = install_task()
        wx.MessageBox(
            caption="Finished", message="Finished Installation.", style=wx.OK | wx.ICON_INFORMATION
        )

        if rsp:
            USBIPD = USBIPD_default
            if gui:
                gui.refresh()


async def amain():
    global gui

    app = wxasync.WxAsyncApp(False)
    gui = WslUsbGui()
    app.SetTopWindow(gui)

    # TODO
    devNotifyHandle = registerDeviceNotification(handle=gui.GetHandle(), callback=usb_callback)

    gui.check_wsl_udev()

    await app.MainLoop()

    try:
        unregisterDeviceNotification(devNotifyHandle)
    except:
        pass


def main():
    global loop
    loop = asyncio.get_event_loop()
    loop.run_until_complete(amain())


if __name__ == "__main__":
    main()
