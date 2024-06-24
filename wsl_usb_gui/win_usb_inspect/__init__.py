import ctypes
import ctypes.wintypes as wintypes
from typing import Optional, Dict, List, Tuple

from .winusbclasses import *

import logging

log = logging.getLogger("usb inspect")

DEBUG = False

didd_cb_sizes = (8, 6, 5)  # different on 64 bit / 32 bit etc
MAX_DEVICE_PROP = 200

# NULL = 0
FALSE = wintypes.BOOL(0)
TRUE = wintypes.BOOL(1)
S_OK = HRESULT(0)
E_FAIL = HRESULT(0x80004005)


# enum USBDEVICEINFOTYPE
class USBDEVICEINFOTYPE:
    HostControllerInfo = 0
    RootHubInfo = 1
    ExternalHubInfo = 2
    DeviceInfo = 3


class DEVICE_INFO_NODE:
    def __init__(self):
        self.DeviceInfo = HDEVINFO()
        # self.ListEntry = LIST_ENTRY()
        self.DeviceInfoData = SP_DEVINFO_DATA()
        self.DeviceInterfaceData = SP_DEVICE_INTERFACE_DATA()
        self.DeviceDetailData = SP_DEVICE_INTERFACE_DETAIL_DATA()
        self.DeviceDescName = ""
        self.DeviceDescNameLength = ULONG()
        self.DeviceDriverName = ""
        self.DeviceDriverNameLength = ULONG()
        self.InstanceId = ""
        # self.LatestDevicePowerState = DEVICE_POWER_STATE()


class USB_DEVICE_PNP_STRINGS:
    def __init__(self):
        self.DeviceId = ""
        self.DeviceDesc = ""
        self.HwId = ""
        self.Service = ""
        self.DeviceClass = ""
        self.PowerState = ""


class USBHOSTCONTROLLERINFO:
    def __init__(self, name):
        self.Name = name
        self.DeviceInfoType = -1
        # ListEntry = LIST_ENTRY()
        self.DriverKey = ""
        self.VendorID = ULONG()
        self.DeviceID = ULONG()
        self.SubSysID = ULONG()
        self.Revision = ULONG()
        # USBPowerInfo = USB_POWER_INFO()
        self.BusDeviceFunctionValid: bool = False
        self.BusNumber = ULONG()
        self.BusDevice = 0
        self.BusFunction = 0
        # ControllerInfo = USB_CONTROLLER_INFO_0()
        self.UsbDeviceProperties = USB_DEVICE_PNP_STRINGS()


class USBROOTHUBINFO:
    def __init__(self, name=""):
        self.DeviceInfoType = USBDEVICEINFOTYPE()
        self.HubInfo = USB_NODE_INFORMATION()
        # self.HubInfoEx = USB_HUB_INFORMATION_EX()
        self.HubName = name
        # self.PortConnectorProps = USB_PORT_CONNECTOR_PROPERTIES()
        self.UsbDeviceProperties = USB_DEVICE_PNP_STRINGS()
        # self.DeviceInfoNode = DEVICE_INFO_NODE()
        # self.HubCapabilityEx = USB_HUB_CAPABILITIES_EX()


class USBEXTERNALHUBINFO:
    def __init__(self, name=""):
        self.DeviceInfoType = USBDEVICEINFOTYPE()
        self.HubInfo = USB_NODE_INFORMATION()
        # self.HubInfoEx = USB_HUB_INFORMATION_EX()
        self.HubName = name
        self.ConnectionInfo = USB_NODE_CONNECTION_INFORMATION_EX()
        # self.PortConnectorProps = USB_PORT_CONNECTOR_PROPERTIES()
        self.ConfigDesc = USB_DESCRIPTOR_REQUEST()
        # self.BosDesc = USB_DESCRIPTOR_REQUEST()
        self.StringDescs = STRING_DESCRIPTOR_NODE()
        # self.ConnectionInfoV2 = USB_NODE_CONNECTION_INFORMATION_EX_V2()
        self.UsbDeviceProperties = USB_DEVICE_PNP_STRINGS()
        # self.DeviceInfoNode = DEVICE_INFO_NODE()
        # self.HubCapabilityEx = USB_HUB_CAPABILITIES_EX()


class USBDEVICEINFO:
    def __init__(self):
        self.SerialNumber = ""
        self.Manufacturer = ""
        self.Product = ""

        self.DeviceInfoType: Optional[USBDEVICEINFOTYPE] = USBDEVICEINFOTYPE()
        self.HubInfo: Optional[USB_NODE_INFORMATION] = (
            USB_NODE_INFORMATION()
        )  # NULL if not a HUB
        # self.HubInfoEx: Optional[USB_HUB_INFORMATION_EX] = USB_HUB_INFORMATION_EX()        # NULL if not a HUB
        self.HubName: Optional[PCHAR] = PCHAR()  # NULL if not a HUB
        self.ConnectionInfo: Optional[USB_NODE_CONNECTION_INFORMATION_EX] = (
            USB_NODE_CONNECTION_INFORMATION_EX()
        )  # NULL if root HUB
        # self.PortConnectorProps: Optional[USB_PORT_CONNECTOR_PROPERTIES] = USB_PORT_CONNECTOR_PROPERTIES()
        self.ConfigDesc: Optional[USB_DESCRIPTOR_REQUEST] = (
            USB_DESCRIPTOR_REQUEST()
        )  # NULL if root HUB
        # self.BosDesc: Optional[USB_DESCRIPTOR_REQUEST] = USB_DESCRIPTOR_REQUEST()          # NULL if root HUB
        self.StringDescs: List[STRING_DESCRIPTOR_NODE] = []
        self.Strings = {}
        # self.ConnectionInfoV2: Optional[USB_NODE_CONNECTION_INFORMATION_EX_V2] = USB_NODE_CONNECTION_INFORMATION_EX_V2() # NULL if root HUB
        self.UsbDeviceProperties: Optional[USB_DEVICE_PNP_STRINGS] = (
            USB_DEVICE_PNP_STRINGS()
        )
        self.DeviceInfoNode: Optional[DEVICE_INFO_NODE] = DEVICE_INFO_NODE()
        # self.HubCapabilityEx: Optional[USB_HUB_CAPABILITIES_EX] = USB_HUB_CAPABILITIES_EX()  # NULL if not a HUB


class STRING_DESCRIPTOR_NODE:
    def __init__(self):
        # self.Next: Optional["STRING_DESCRIPTOR_NODE"] = None
        self.HubIsBusPowered = False
        self.DescriptorIndex = 0
        self.LanguageID = 0
        self.StringDescriptor: Optional[USB_STRING_DESCRIPTOR] = None


gHubList: List[DEVICE_INFO_NODE] = []
gDeviceList: List[DEVICE_INFO_NODE] = []
gHostControllerList: List[USBHOSTCONTROLLERINFO] = []
TotalHubs = 0
TotalDevicesConnected = 0
gDoConfigDesc = True

parsed_devices = {}


def EnumerateAllDevices():
    global gHubList, gDeviceList
    gDeviceList = EnumerateAllDevicesWithGuid(GUID_DEVINTERFACE_USB_DEVICE)
    gHubList = EnumerateAllDevicesWithGuid(GUID_DEVINTERFACE_USB_HUB)


def EnumerateAllDevicesWithGuid(Guid) -> List[DEVICE_INFO_NODE]:
    DeviceList: List[DEVICE_INFO_NODE] = []

    DeviceInfo: HDEVINFO = SetupDiGetClassDevs(
        byref(Guid), None, None, DWORD(DIGCF_PRESENT | DIGCF_DEVICEINTERFACE)
    )

    index = ULONG(0)
    error = 0

    error = 0
    index = 0

    while error != ERROR_NO_MORE_ITEMS:
        success = BOOL(False)
        pNode = DEVICE_INFO_NODE()

        pNode.DeviceInfo = DeviceInfo
        pNode.DeviceInterfaceData.cbSize = sizeof(pNode.DeviceInterfaceData)
        pNode.DeviceInfoData.cbSize = sizeof(pNode.DeviceInfoData)

        success = SetupDiEnumDeviceInfo(
            DeviceInfo, DWORD(index), byref(pNode.DeviceInfoData)
        )

        index += 1

        if not success:
            error = GetLastError()

            if error != ERROR_NO_MORE_ITEMS:
                raise Exception("OOPS")

            # FreeDeviceInfoNode(byref(pNode))
            pNode = None

        else:
            bResult = BOOL(0)
            requiredLength = ULONG(0)

            bResult, pNode.DeviceDescName = GetDeviceProperty(
                DeviceInfo, pNode.DeviceInfoData, SPDRP_DEVICEDESC
            )
            if not bResult:
                # FreeDeviceInfoNode(byref(pNode))
                pNode = None
                raise Exception("OOPS")
                break

            bResult, pNode.DeviceDriverName = GetDeviceProperty(
                DeviceInfo, pNode.DeviceInfoData, SPDRP_DRIVER
            )
            if not bResult:
                # FreeDeviceInfoNode(byref(pNode))
                pNode = None
                raise Exception("OOPS")
                break

            pNode.DeviceInterfaceData.cbSize = sizeof(SP_DEVICE_INTERFACE_DATA)

            success = SetupDiEnumDeviceInterfaces(
                DeviceInfo,
                None,
                byref(Guid),
                index - 1,
                byref(pNode.DeviceInterfaceData),
            )
            if not success:
                # FreeDeviceInfoNode(byref(pNode))
                pNode = None
                raise Exception("OOPS")
                break

            success = SetupDiGetDeviceInterfaceDetail(
                DeviceInfo,
                byref(pNode.DeviceInterfaceData),
                NULL,
                0,
                byref(requiredLength),
                NULL,
            )

            error = GetLastError()

            if not success and error != ERROR_INSUFFICIENT_BUFFER:
                # FreeDeviceInfoNode(byref(pNode))
                pNode = None
                raise Exception("OOPS")
                break

            resize(pNode.DeviceDetailData, requiredLength.value)
            # pNode.DeviceDetailData = ALLOC(requiredLength)

            if pNode.DeviceDetailData == NULL:
                # FreeDeviceInfoNode(byref(pNode))
                pNode = None
                raise Exception("OOPS")
                break

            success = False
            global didd_cb_sizes
            for cb_size in didd_cb_sizes:
                pNode.DeviceDetailData.cbSize = (
                    cb_size  # sizeof(SP_DEVICE_INTERFACE_DETAIL_DATA) # cb_size
                )

                success = SetupDiGetDeviceInterfaceDetail(
                    DeviceInfo,
                    byref(pNode.DeviceInterfaceData),
                    byref(pNode.DeviceDetailData),
                    requiredLength,
                    byref(requiredLength),
                    byref(pNode.DeviceInfoData),
                )
                if not success:
                    error = GetLastError()
                    if DEBUG:
                        log.error(error)
                if success:
                    didd_cb_sizes = (cb_size,)
                    pNode.DeviceDetailData.cbSize = requiredLength.value
                    break
                # define ERROR_INVALID_USER_BUFFER        1784L
                # define ERROR_INSUFFICIENT_BUFFER        122L    # dderror

            if not success:
                # FreeDeviceInfoNode(byref(pNode))
                pNode = None
                raise Exception("OOPS")
                break

            DeviceList.append(pNode)

    return DeviceList


def InspectUsbDevices() -> Tuple[Dict[str, USBDEVICEINFO], List]:
    hHCDev: HANDLE = HANDLE()
    deviceInfo = HDEVINFO()
    deviceInfoData = SP_DEVINFO_DATA()
    deviceInterfaceData = SP_DEVICE_INTERFACE_DATA()
    deviceDetailData = SP_DEVICE_INTERFACE_DETAIL_DATA()
    index = ULONG(0)
    requiredLength = ULONG(0)
    success = False

    DevicesConnected = 0
    global TotalDevicesConnected
    TotalDevicesConnected = 0
    global TotalHubs
    TotalHubs = 0

    EnumerateAllDevices()

    # Iterate over host controllers using the new GUID based interface
    #
    deviceInfo = SetupDiGetClassDevs(
        byref(GUID_CLASS_USB_HOST_CONTROLLER),
        NULL,
        NULL,
        (DIGCF_PRESENT | DIGCF_DEVICEINTERFACE),
    )

    deviceInfoData.cbSize = sizeof(SP_DEVINFO_DATA)

    full_tree = []
    index = 0
    while SetupDiEnumDeviceInfo(deviceInfo, index, byref(deviceInfoData)):
        index += 1
        deviceInterfaceData.cbSize = sizeof(SP_DEVICE_INTERFACE_DATA)

        devInterfaceIndex = 0
        while SetupDiEnumDeviceInterfaces(
            deviceInfo,
            byref(deviceInfoData),
            byref(GUID_CLASS_USB_HOST_CONTROLLER),
            devInterfaceIndex,
            byref(deviceInterfaceData),
        ):
            devInterfaceIndex += 1
            success = SetupDiGetDeviceInterfaceDetail(
                deviceInfo,
                byref(deviceInterfaceData),
                NULL,
                0,
                byref(requiredLength),
                NULL,
            )

            if not success and GetLastError() != ERROR_INSUFFICIENT_BUFFER:
                raise Exception("OOPS")
                break

            resize(deviceDetailData, requiredLength.value)
            # deviceDetailData = ALLOC(requiredLength)
            if deviceDetailData == NULL:
                raise Exception("OOPS")
                break

            global didd_cb_sizes
            deviceDetailData.cbSize = didd_cb_sizes[0]

            success = SetupDiGetDeviceInterfaceDetail(
                deviceInfo,
                byref(deviceInterfaceData),
                byref(deviceDetailData),
                requiredLength,
                byref(requiredLength),
                NULL,
            )

            if not success:
                error = GetLastError()
                raise Exception("OOPS")
                break

            hHCDev = CreateFile(
                str(deviceDetailData),
                GENERIC_WRITE,
                FILE_SHARE_WRITE,
                NULL,
                OPEN_EXISTING,
                0,
                NULL,
            )

            # If the handle is valid, then we've successfully opened a Host
            # Controller.  Display some info about the Host Controller itself,
            # then enumerate the Root Hub attached to the Host Controller.
            #

            if hHCDev != INVALID_HANDLE_VALUE:
                items = EnumerateHostController(
                    hHCDev, str(deviceDetailData), deviceInfo, deviceInfoData
                )

                CloseHandle(hHCDev)
                full_tree.append(items)

            # FREE(deviceDetailData)

    SetupDiDestroyDeviceInfoList(deviceInfo)

    # *DevicesConnected = TotalDevicesConnected

    return parsed_devices, full_tree


def EnumerateHostController(
    hHCDev: HANDLE, leafName: str, deviceInfo: HANDLE, deviceInfoData: SP_DEVINFO_DATA
):
    driverKeyName = PCHAR()
    hHCItem = list()
    rootHubName = ""
    # listEntry = PLIST_ENTRY()
    hcInfo = USBHOSTCONTROLLERINFO(leafName)
    # hcInfoInList = USBHOSTCONTROLLERINFO()
    dwSuccess = DWORD()
    success = BOOL()
    deviceAndFunction = ULONG()
    DevProps = USB_DEVICE_PNP_STRINGS()

    # Allocate a structure to hold information about this host controller.
    #
    # hcInfo = (USBHOSTCONTROLLERINFO)ALLOC(sizeof(USBHOSTCONTROLLERINFO))

    # just return if could not alloc memory
    if NULL == hcInfo:
        return

    hcInfo.DeviceInfoType = USBDEVICEINFOTYPE.HostControllerInfo

    # Obtain the driver key name for this host controller.
    #
    driverKeyName = GetHCDDriverKeyName(hHCDev)

    if NULL == driverKeyName:
        # Failure obtaining driver key name.
        raise Exception("OOPS")
        FREE(hcInfo)
        return

    # Don't enumerate this host controller again if it already
    # on the list of enumerated host controllers.
    #
    # listEntry = EnumeratedHCListHead.Flink

    # for hcInfoInList in gHostControllerList:
    # # while (hcInfoInList != byref(EnumeratedHCListHead)):
    #     # hcInfoInList = CONTAINING_RECORD(listEntry,
    #                                     # USBHOSTCONTROLLERINFO,
    #                                     # ListEntry)

    #     if (driverKeyName == hcInfoInList.DriverKey) :
    #         # Already on the list, exit
    #         #
    #         # FREE(driverKeyName)
    #         # FREE(hcInfo)
    #         return

    # listEntry = listEntry.Flink

    # Obtain host controller device properties:
    cbDriverName = len(driverKeyName)
    # hr = S_OK

    # hr = StringCbLength(driverKeyName, MAX_DRIVER_KEY_NAME, byref(cbDriverName))
    # if (SUCCEEDED(hr)):
    DevProps = DriverNameToDeviceProperties(driverKeyName, cbDriverName)

    hcInfo.DriverKey = driverKeyName

    # if (DevProps):
    #     ULONG   ven, dev, subsys, rev
    #     ven = dev = subsys = rev = 0

    #     if (sscanf_s(DevProps.DeviceId,
    #             "PCI\\VEN_%x&DEV_%x&SUBSYS_%x&REV_%x",
    #             byref(ven), byref(dev), byref(subsys), byref(rev)) != 4):
    #         raise Exception("OOPS")

    #     hcInfo.VendorID = ven
    #     hcInfo.DeviceID = dev
    #     hcInfo.SubSysID = subsys
    #     hcInfo.Revision = rev
    #     hcInfo.UsbDeviceProperties = DevProps
    # else:
    #     raise Exception("OOPS")

    # if (DevProps != NULL and DevProps.DeviceDesc != NULL):
    #     leafName = DevProps.DeviceDesc
    # else:
    #     raise Exception("OOPS")

    # Get the USB Host Controller power map
    # dwSuccess = GetHostControllerPowerMap(hHCDev, hcInfo)

    # if (ERROR_SUCCESS != dwSuccess):
    #     raise Exception("OOPS")

    # Get bus, device, and function
    #
    hcInfo.BusDeviceFunctionValid = False

    success = SetupDiGetDeviceRegistryProperty(
        deviceInfo,
        deviceInfoData,
        SPDRP_BUSNUMBER,
        NULL,
        byref(hcInfo.BusNumber),
        sizeof(hcInfo.BusNumber),
        NULL,
    )

    if success:
        success = SetupDiGetDeviceRegistryProperty(
            deviceInfo,
            deviceInfoData,
            SPDRP_ADDRESS,
            NULL,
            byref(deviceAndFunction),
            sizeof(deviceAndFunction),
            NULL,
        )

    if success:
        hcInfo.BusDevice = deviceAndFunction.value >> 16
        hcInfo.BusFunction = deviceAndFunction.value & 0xFFFF
        hcInfo.BusDeviceFunctionValid = True

    # Get the USB Host Controller info
    # dwSuccess = GetHostControllerInfo(hHCDev, hcInfo)

    # if (ERROR_SUCCESS != dwSuccess):
    #     raise Exception("OOPS")

    # Add this host controller to the USB device tree view.
    #
    # hHCItem = AddLeaf(hTreeParent,
    #                 (LPARAM)hcInfo,
    #                 leafName,
    #                 GoodSsDeviceIcon if (hcInfo.Revision == UsbSuperSpeed) else GoodDeviceIcon)

    # if (NULL == hHCItem):
    #     # Failure adding host controller to USB device tree
    #     # view.

    #     raise Exception("OOPS")
    #     FREE(driverKeyName)
    #     FREE(hcInfo)
    #     return

    # Add this host controller to the list of enumerated
    # host controllers.
    #
    # InsertTailList(byref(EnumeratedHCListHead),
    #             &hcInfo.ListEntry)
    gHostControllerList.append(hcInfo)

    # Get the name of the root hub for this host
    # controller and then enumerate the root hub.
    #
    rootHubName = GetRootHubName(hHCDev)

    if rootHubName:
        cbHubName = 0
        # HRESULT hr = S_OK

        # hr = StringCbLength(rootHubName, MAX_DRIVER_KEY_NAME, byref(cbHubName))
        cbHubName = len(rootHubName)
        # if (SUCCEEDED(hr)):
        EnumerateHub(
            hHCItem,
            rootHubName,
            cbHubName,
            NULL,  # ConnectionInfo
            # NULL,       # ConnectionInfoV2
            # NULL,       # PortConnectorProps
            NULL,  # ConfigDesc
            # NULL,       # BosDesc
            NULL,  # StringDescs
            NULL,  # Strings
            NULL,
        )  # We do not pass DevProps for RootHub
    # else:
    #     # Failure obtaining root hub name.

    # raise Exception("OOPS")

    return hHCItem


def DriverNameToDeviceProperties(
    DriverName: str, cbDriverName: int
) -> USB_DEVICE_PNP_STRINGS:

    # deviceInfo = HDEVINFO()
    # deviceInfoData = SP_DEVINFO_DATA()
    length = ULONG()
    status = BOOL()
    DevProps = USB_DEVICE_PNP_STRINGS()
    lastError = DWORD()

    # Allocate device propeties structure
    # DevProps = (USB_DEVICE_PNP_STRINGS) ALLOC(sizeof(USB_DEVICE_PNP_STRINGS))

    if NULL == DevProps:
        status = FALSE
        error = GetLastError()
        if DEBUG:
            log.error(WinError(get_last_error()))

    # Get device instance
    status, deviceInfo, deviceInfoData = DriverNameToDeviceInst(
        DriverName, cbDriverName
    )  # , byref(deviceInfo), byref(deviceInfoData))
    if not status:
        # goto Done
        error = GetLastError()
        if DEBUG:
            log.error(WinError(get_last_error()))

    # When device is not attached this matches the usbipd InstanceID
    # Once attached however the VID/PID elements change to relate to the
    # "filter driver" ?? so can no longer be used to match usbipd ids.
    DevProps.DeviceId = GetInstanceId(deviceInfo, deviceInfoData)

    # status = GetDeviceProperty(deviceInfo,
    #                            byref(deviceInfoData),
    #                            SPDRP_DEVICEDESC,
    #                            byref(DevProps.DeviceDesc))

    # if (not status):
    #     #goto Done
    #     error = GetLastError()
    #     log.error(WinError(get_last_error()))

    #     #
    #     # We don't fail if the following registry query fails as these fields are additional information only
    #     #

    #     GetDeviceProperty(deviceInfo,
    #                       byref(deviceInfoData),
    #                       SPDRP_HARDWAREID,
    #                       byref(DevProps.HwId))

    #     GetDeviceProperty(deviceInfo,
    #                       byref(deviceInfoData),
    #                       SPDRP_SERVICE,
    #                       byref(DevProps.Service))

    #     GetDeviceProperty(deviceInfo,
    #                        byref(deviceInfoData),
    #                        SPDRP_CLASS,
    #                        byref(DevProps.DeviceClass))
    # Done:

    #     if (deviceInfo != INVALID_HANDLE_VALUE):
    #         SetupDiDestroyDeviceInfoList(deviceInfo)

    #     if (not status):
    #         if (DevProps != NULL):
    #             FreeDeviceProperties(byref(DevProps))

    return DevProps


# def GetHostControllerInfo(
#     hHCDev: HANDLE,
#     hcInfo: USBHOSTCONTROLLERINFO):

#     UsbControllerInfo: USBUSER_CONTROLLER_INFO_0
#     dwError: DWORD = 0
#     dwBytes: DWORD = 0
#     bSuccess: BOOL = FALSE

#     memset(byref(UsbControllerInfo), 0, sizeof(UsbControllerInfo))

#     # set the header and request sizes
#     UsbControllerInfo.Header.UsbUserRequest = USBUSER_GET_CONTROLLER_INFO_0
#     UsbControllerInfo.Header.RequestBufferLength = sizeof(UsbControllerInfo)

#     #
#     # Query for the USB_CONTROLLER_INFO_0 structure
#     #
#     bSuccess = DeviceIoControl(hHCDev,
#             IOCTL_USB_USER_REQUEST,
#             byref(UsbControllerInfo),
#             sizeof(UsbControllerInfo),
#             byref(UsbControllerInfo),
#             sizeof(UsbControllerInfo),
#             byref(dwBytes),
#             NULL)

#     if (NOT bSuccess):
#         dwError = GetLastError()
#         raise Exception("OOPS")

#     else:
#         hcInfo.ControllerInfo = (USB_CONTROLLER_INFO_0) ALLOC(sizeof(USB_CONTROLLER_INFO_0))
#         if(NULL == hcInfo.ControllerInfo):
#             dwError = GetLastError()
#             raise Exception("OOPS")

#         else:
#             # copy the data into our USB Host Controller's info structure
#             memcpy(hcInfo.ControllerInfo, byref(UsbControllerInfo.Info0), sizeof(USB_CONTROLLER_INFO_0))


#     return dwError


def GetRootHubName(HostController: HANDLE):
    success = BOOL(0)
    nBytes = ULONG(0)
    rootHubName = USB_ROOT_HUB_NAME()  #
    rootHubNameW = USB_ROOT_HUB_NAME()  # = NULL
    rootHubNameA = PCHAR()  # = NULL

    # Get the length of the name of the Root Hub attached to the
    # Host Controller
    #
    success = DeviceIoControl(
        HostController,
        IOCTL_USB_GET_ROOT_HUB_NAME,
        0,
        0,
        byref(rootHubName),
        sizeof(rootHubName),
        byref(nBytes),
        NULL,
    )

    if not success:
        raise Exception("OOPS")
        # goto GetRootHubNameError

    # Allocate space to hold the Root Hub name
    #
    nBytes = DWORD(rootHubName.ActualLength)

    # rootHubNameW = ALLOC(nBytes)
    resize(rootHubNameW, nBytes.value)
    if rootHubNameW == NULL:
        raise Exception("OOPS")
        # goto GetRootHubNameError

    # Get the name of the Root Hub attached to the Host Controller
    #
    success = DeviceIoControl(
        HostController,
        IOCTL_USB_GET_ROOT_HUB_NAME,
        NULL,
        0,
        byref(rootHubNameW),
        nBytes,
        byref(nBytes),
        NULL,
    )
    if not success:
        raise Exception("OOPS")
        # goto GetRootHubNameError

    return str(rootHubNameW)
    # # Convert the Root Hub name
    # #
    # rootHubNameA = WideStrToMultiStr(rootHubNameW.RootHubName, nBytes - sizeof(USB_ROOT_HUB_NAME) + sizeof(WCHAR))

    # # All done, free the uncoverted Root Hub name and return the
    # # converted Root Hub name
    # #
    # # FREE(rootHubNameW)

    # return rootHubNameA


# GetRootHubNameError:
#     # There was an error, free anything that was allocated
#     #
#     if (rootHubNameW != NULL):
#         FREE(rootHubNameW)
#         rootHubNameW = NULL

#     return NULL


def EnumerateHub(
    hTreeParent: List,
    HubName: str,
    cbHubName: int,
    ConnectionInfo: USB_NODE_CONNECTION_INFORMATION_EX,
    # ConnectionInfoV2: USB_NODE_CONNECTION_INFORMATION_EX_V2,
    # PortConnectorProps: USB_PORT_CONNECTOR_PROPERTIES,
    ConfigDesc: USB_DESCRIPTOR_REQUEST,
    # BosDesc: USB_DESCRIPTOR_REQUEST,
    StringDescs: List[STRING_DESCRIPTOR_NODE],
    Strings: dict,
    DevProps: USB_DEVICE_PNP_STRINGS,
):
    # Initialize locals to not allocated state so the error cleanup routine
    # only tries to cleanup things that were successfully allocated.
    #
    hubInfo = USB_NODE_INFORMATION()
    # hubInfoEx = USB_HUB_INFORMATION_EX()
    # hubCapabilityEx = USB_HUB_CAPABILITIES_EX()
    hHubDevice = INVALID_HANDLE_VALUE
    # hItem = HTREEITEM()
    info_ex = USBEXTERNALHUBINFO()
    info_root = USBROOTHUBINFO()
    deviceName = PCHAR()
    nBytes = ULONG(0)
    success = BOOL(0)

    hr = S_OK
    cchHeader = 0
    cchFullHubName = 0

    # Allocate some space for a USBDEVICEINFO structure to hold the
    # hub info, hub name, and connection info pointers.  GPTR zero
    # initializes the structure for us.
    #
    # info = ALLOC(sizeof(USBEXTERNALHUBINFO))
    # if (info == NULL):
    #         raise Exception("OOPS")
    #         goto EnumerateHubError

    # Allocate some space for a USB_NODE_INFORMATION structure for this Hub
    #
    # hubInfo = (USB_NODE_INFORMATION)ALLOC(sizeof(USB_NODE_INFORMATION))
    # if (hubInfo == NULL):
    #     raise Exception("OOPS")
    #     goto EnumerateHubError

    # hubInfoEx = (USB_HUB_INFORMATION_EX)ALLOC(sizeof(USB_HUB_INFORMATION_EX))
    # if (hubInfoEx == NULL):
    #     raise Exception("OOPS")
    #     goto EnumerateHubError

    # hubCapabilityEx = (USB_HUB_CAPABILITIES_EX)ALLOC(sizeof(USB_HUB_CAPABILITIES_EX))
    # if(hubCapabilityEx == NULL):
    #     raise Exception("OOPS")
    #     goto EnumerateHubError

    # Keep copies of the Hub Name, Connection Info, and Configuration
    # Descriptor pointers
    #
    info_ex.HubInfo = hubInfo
    info_ex.HubName = HubName

    info_root.HubInfo = hubInfo
    info_root.HubName = HubName

    if ConnectionInfo != NULL:
        info_ex.DeviceInfoType = USBDEVICEINFOTYPE.ExternalHubInfo
        info_ex.ConnectionInfo = ConnectionInfo
        info_ex.ConfigDesc = ConfigDesc
        info_ex.StringDescs = StringDescs
        # info_ex.PortConnectorProps = PortConnectorProps
        # info_ex.HubInfoEx = hubInfoEx
        # info_ex.HubCapabilityEx = hubCapabilityEx
        # info_ex.BosDesc = BosDesc
        # info_ex.ConnectionInfoV2 = ConnectionInfoV2
        info_ex.UsbDeviceProperties = DevProps
        info = info_ex

    else:
        info_root.DeviceInfoType = USBDEVICEINFOTYPE.RootHubInfo
        # info_root.HubInfoEx = hubInfoEx
        # info_root.HubCapabilityEx = hubCapabilityEx
        # info_root.PortConnectorProps = PortConnectorProps
        info_root.UsbDeviceProperties = DevProps
        info = info_root

    # Allocate a temp buffer for the full hub device name.
    #
    # cchHeader = len("\\\\.\\") + MAX_DEVICE_PROP
    # hr = StringCbLength("\\\\.\\", MAX_DEVICE_PROP, byref(cchHeader))
    # if (FAILED(hr)):
    #     goto EnumerateHubError

    # cchFullHubName = cchHeader + cbHubName + 1
    # # deviceName = (PCHAR)ALLOC((DWORD) cchFullHubName)
    # deviceName = ctypes.create_unicode_buffer(cchFullHubName)
    # if (deviceName == NULL):
    #     raise Exception("OOPS")
    #     # goto EnumerateHubError

    # Create the full hub device name
    #
    # deviceName = "\\\\.\\" + HubName
    # hr = StringCchCopyN(deviceName, cchFullHubName, "\\\\.\\", cchHeader)
    # if (FAILED(hr)):
    #     goto EnumerateHubError

    # hr = StringCchCatN(deviceName, cchFullHubName, HubName, cbHubName)
    # if (FAILED(hr)):
    #     goto EnumerateHubError
    deviceName = "\\\\.\\" + HubName

    # Try to hub the open device
    #
    hHubDevice = CreateFile(
        deviceName, GENERIC_WRITE, FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL
    )

    # Done with temp buffer for full hub device name
    #
    # FREE(deviceName)

    if hHubDevice == INVALID_HANDLE_VALUE:
        raise Exception("OOPS")
        # goto EnumerateHubError

    #
    # Now query USBHUB for the USB_NODE_INFORMATION structure for this hub.
    # This will tell us the number of downstream ports to enumerate, among
    # other things.
    #
    success = DeviceIoControl(
        hHubDevice,
        IOCTL_USB_GET_NODE_INFORMATION,
        byref(hubInfo),
        sizeof(USB_NODE_INFORMATION),
        byref(hubInfo),
        sizeof(USB_NODE_INFORMATION),
        byref(nBytes),
        NULL,
    )

    if not success:
        error = GetLastError()
        if DEBUG:
            log.error(WinError(get_last_error()))
        raise Exception("OOPS")
        # goto EnumerateHubError

    # success = DeviceIoControl(hHubDevice,
    #                           IOCTL_USB_GET_HUB_INFORMATION_EX,
    #                           hubInfoEx,
    #                           sizeof(USB_HUB_INFORMATION_EX),
    #                           hubInfoEx,
    #                           sizeof(USB_HUB_INFORMATION_EX),
    #                           byref(nBytes),
    #                           NULL)

    #
    # Fail gracefully for downlevel OS's from Win8
    #
    # if (!success || nBytes < sizeof(USB_HUB_INFORMATION_EX)):
    #     FREE(hubInfoEx)
    #     hubInfoEx = NULL
    #     if (ConnectionInfo != NULL):
    #         ((USBEXTERNALHUBINFO)info).HubInfoEx = NULL

    #     else:
    #         ((USBROOTHUBINFO)info).HubInfoEx = NULL

    #
    # Obtain Hub Capabilities
    #
    # success = DeviceIoControl(hHubDevice,
    #                           IOCTL_USB_GET_HUB_CAPABILITIES_EX,
    #                           hubCapabilityEx,
    #                           sizeof(USB_HUB_CAPABILITIES_EX),
    #                           hubCapabilityEx,
    #                           sizeof(USB_HUB_CAPABILITIES_EX),
    #                           byref(nBytes),
    #                           NULL)

    #
    # Fail gracefully
    #
    # if (!success || nBytes < sizeof(USB_HUB_CAPABILITIES_EX)):
    #     FREE(hubCapabilityEx)
    #     hubCapabilityEx = NULL
    #     if (ConnectionInfo != NULL):
    #         ((USBEXTERNALHUBINFO)info).HubCapabilityEx = NULL

    #     else:
    #         ((USBROOTHUBINFO)info).HubCapabilityEx = NULL

    # Build the leaf name from the port number and the device description
    #
    if ConnectionInfo:
        leafName = f"[Port{ConnectionInfo.ConnectionIndex}]"
    else:
        leafName = ""

    # dwSizeOfLeafName = sizeof(leafName)
    # if (ConnectionInfo):
    #     StringCchPrintf(leafName, dwSizeOfLeafName, "[Port%d] ", ConnectionInfo.ConnectionIndex)
    #     StringCchCat(leafName,
    #         dwSizeOfLeafName,
    #         ConnectionStatuses[ConnectionInfo.ConnectionStatus])
    #     StringCchCatN(leafName,
    #         dwSizeOfLeafName,
    #         " :  ",
    #         sizeof(" :  "))

    if DevProps and DevProps.DeviceDesc:
        # size_t cbDeviceDesc = 0
        # hr = StringCbLength(DevProps.DeviceDesc, MAX_DRIVER_KEY_NAME, byref(cbDeviceDesc))
        # if(SUCCEEDED(hr)):
        #     StringCchCatN(leafName,
        #             dwSizeOfLeafName,
        #             DevProps.DeviceDesc,
        #             cbDeviceDesc)
        leafName += str(DevProps.DeviceDesc)

    else:
        if ConnectionInfo != NULL:
            # External hub
            leafName += "ExternalHub"

        else:
            # Root hub
            leafName += "RootHub"
            # StringCchCatN(leafName,
            #         dwSizeOfLeafName,
            #         "RootHub",
            #         sizeof("RootHub"))

    # Now add an item to the TreeView with the USBDEVICEINFO pointer info
    # as the LPARAM reference value containing everything we know about the
    # hub.
    #
    # hItem = AddLeaf(hTreeParent,
    #                 (LPARAM)info,
    #                 leafName,
    #                 HubIcon)

    # if (hItem == NULL):
    #     raise Exception("OOPS")
    #     goto EnumerateHubError
    children = []
    hTreeParent.append((leafName, info, children))

    # Now recursively enumerate the ports of this hub.
    #
    EnumerateHubPorts(
        children, hHubDevice, hubInfo.u.HubInformation.HubDescriptor.bNumberOfPorts
    )

    CloseHandle(hHubDevice)
    return


def EnumerateHubPorts(hTreeParent: List, hHubDevice: HANDLE, bNumPorts: bytes):
    index = ULONG(0)
    success = BOOL(0)
    NumPorts = ord(bNumPorts)
    hr = S_OK
    driverKeyName = ""
    DevProps = USB_DEVICE_PNP_STRINGS()
    dwSizeOfLeafName = DWORD(0)
    leafName = ""
    icon = 0
    connectionInfoEx = USB_NODE_CONNECTION_INFORMATION_EX()
    # pPortConnectorProps = USB_PORT_CONNECTOR_PROPERTIES()
    # portConnectorProps = USB_PORT_CONNECTOR_PROPERTIES()
    configDescReq = USB_DESCRIPTOR_REQUEST()
    # bosDesc = USB_DESCRIPTOR_REQUEST()
    stringDescs: List[STRING_DESCRIPTOR_NODE] = []  # STRING_DESCRIPTOR_NODE()
    info = USBDEVICEINFO()
    # connectionInfoExV2 = USB_NODE_CONNECTION_INFORMATION_EX_V2()
    pNode = DEVICE_INFO_NODE()

    # Loop over all ports of the hub.
    #
    # Port indices are 1 based, not 0 based.
    #
    for index in range(1, NumPorts + 1):
        nBytesEx = 0
        nBytes = 0

        connectionInfoEx = NULL
        pPortConnectorProps = NULL
        # ZeroMemory(byref(portConnectorProps), sizeof(portConnectorProps))
        configDescReq: PUSB_DESCRIPTOR_REQUEST = NULL
        configDescBuff = NULL
        # bosDesc = NULL
        stringDescs = []
        info = NULL
        connectionInfoExV2 = NULL
        pNode = NULL
        DevProps = NULL
        leafName = ""
        # ZeroMemory(leafName, sizeof(leafName))

        #
        # Allocate space to hold the connection info for this port.
        # For now, allocate it big enough to hold info for 30 pipes.
        #
        # Endpoint numbers are 0-15.  Endpoint number 0 is the standard
        # control endpoint which is not explicitly listed in the Configuration
        # Descriptor.  There can be an IN endpoint and an OUT endpoint at
        # endpoint numbers 1-15 so there can be a maximum of 30 endpoints
        # per device configuration.
        #
        # Should probably size this dynamically at some point.
        #

        nBytesEx = DWORD(
            sizeof(USB_NODE_CONNECTION_INFORMATION_EX) + (sizeof(USB_PIPE_INFO) * 30)
        )

        connectionInfoEx = USB_NODE_CONNECTION_INFORMATION_EX()
        resize(connectionInfoEx, nBytesEx.value)

        if connectionInfoEx == NULL:
            raise Exception("OOPS")
            break

        # connectionInfoExV2 = (USB_NODE_CONNECTION_INFORMATION_EX_V2)
        #                             ALLOC(sizeof(USB_NODE_CONNECTION_INFORMATION_EX_V2))

        # if (connectionInfoExV2 == NULL):
        #     raise Exception("OOPS")
        #     FREE(connectionInfoEx)
        #     break

        #
        # Now query USBHUB for the structures
        # for this port.  This will tell us if a device is attached to this
        # port, among other things.
        # The fault tolerate code is executed first.
        #

        # portConnectorProps.ConnectionIndex = index

        # success = DeviceIoControl(hHubDevice,
        #                           IOCTL_USB_GET_PORT_CONNECTOR_PROPERTIES,
        #                           byref(portConnectorProps),
        #                           sizeof(USB_PORT_CONNECTOR_PROPERTIES),
        #                           byref(portConnectorProps),
        #                           sizeof(USB_PORT_CONNECTOR_PROPERTIES),
        #                           byref(nBytes),
        #                           NULL)

        # if (success && nBytes == sizeof(USB_PORT_CONNECTOR_PROPERTIES)):
        #     pPortConnectorProps = (USB_PORT_CONNECTOR_PROPERTIES)
        #                                 ALLOC(portConnectorProps.ActualLength)

        #     if (pPortConnectorProps != NULL):
        #         pPortConnectorProps.ConnectionIndex = index

        #         success = DeviceIoControl(hHubDevice,
        #                                   IOCTL_USB_GET_PORT_CONNECTOR_PROPERTIES,
        #                                   pPortConnectorProps,
        #                                   portConnectorProps.ActualLength,
        #                                   pPortConnectorProps,
        #                                   portConnectorProps.ActualLength,
        #                                   byref(nBytes),
        #                                   NULL)

        #         if (not success or nBytes < portConnectorProps.ActualLength):
        #             FREE(pPortConnectorProps)
        #             pPortConnectorProps = NULL

        # connectionInfoExV2.ConnectionIndex = index
        # connectionInfoExV2.Length = sizeof(USB_NODE_CONNECTION_INFORMATION_EX_V2)
        # connectionInfoExV2.SupportedUsbProtocols.Usb300 = 1

        # success = DeviceIoControl(hHubDevice,
        #                           IOCTL_USB_GET_NODE_CONNECTION_INFORMATION_EX_V2,
        #                           connectionInfoExV2,
        #                           sizeof(USB_NODE_CONNECTION_INFORMATION_EX_V2),
        #                           connectionInfoExV2,
        #                           sizeof(USB_NODE_CONNECTION_INFORMATION_EX_V2),
        #                           byref(nBytes),
        #                           NULL)

        # if (!success || nBytes < sizeof(USB_NODE_CONNECTION_INFORMATION_EX_V2)):
        #     FREE(connectionInfoExV2)
        #     connectionInfoExV2 = NULL

        connectionInfoEx.ConnectionIndex = DWORD(index)

        success = DeviceIoControl(
            hHubDevice,
            IOCTL_USB_GET_NODE_CONNECTION_INFORMATION_EX,
            byref(connectionInfoEx),
            nBytesEx,
            byref(connectionInfoEx),
            nBytesEx,
            byref(nBytesEx),
            NULL,
        )

        # if (success):
        #
        # Since the USB_NODE_CONNECTION_INFORMATION_EX is used to display
        # the device speed, but the hub driver doesn't support indication
        # of superspeed, we overwrite the value if the super speed
        # data structures are available and indicate the device is operating
        # at SuperSpeed.
        #

        # if (connectionInfoEx.Speed == UsbHighSpeed
        #     && connectionInfoExV2 != NULL
        #     && (connectionInfoExV2.Flags.DeviceIsOperatingAtSuperSpeedOrHigher ||
        #         connectionInfoExV2.Flags.DeviceIsOperatingAtSuperSpeedPlusOrHigher)):
        #     connectionInfoEx.Speed = UsbSuperSpeed

        if not success:
            connectionInfo = USB_NODE_CONNECTION_INFORMATION()

            # Try using IOCTL_USB_GET_NODE_CONNECTION_INFORMATION
            # instead of IOCTL_USB_GET_NODE_CONNECTION_INFORMATION_EX
            #

            nBytes = DWORD(
                sizeof(USB_NODE_CONNECTION_INFORMATION) + sizeof(USB_PIPE_INFO) * 30
            )

            resize(connectionInfo, nBytes.value)

            if connectionInfo == NULL:
                raise Exception("OOPS")

                # FREE(connectionInfoEx)
                # if (pPortConnectorProps != NULL):
                #     FREE(pPortConnectorProps)

                # if (connectionInfoExV2 != NULL):
                #     FREE(connectionInfoExV2)

                # continue

            connectionInfo.ConnectionIndex = index

            success = DeviceIoControl(
                hHubDevice,
                IOCTL_USB_GET_NODE_CONNECTION_INFORMATION,
                byref(connectionInfo),
                nBytes,
                byref(connectionInfo),
                nBytes,
                byref(nBytes),
                NULL,
            )

            if not success:
                raise Exception("OOPS")

                # FREE(connectionInfo)
                # FREE(connectionInfoEx)
                # if (pPortConnectorProps != NULL):
                #     FREE(pPortConnectorProps)

                # if (connectionInfoExV2 != NULL):
                #     FREE(connectionInfoExV2)

                # continue

            # Copy IOCTL_USB_GET_NODE_CONNECTION_INFORMATION into
            # IOCTL_USB_GET_NODE_CONNECTION_INFORMATION_EX structure.
            #
            connectionInfoEx.ConnectionIndex = connectionInfo.ConnectionIndex
            connectionInfoEx.DeviceDescriptor = connectionInfo.DeviceDescriptor
            connectionInfoEx.CurrentConfigurationValue = (
                connectionInfo.CurrentConfigurationValue
            )
            # connectionInfoEx.Speed = connectionInfo.LowSpeed ? UsbLowSpeed : UsbFullSpeed
            connectionInfoEx.DeviceIsHub = connectionInfo.DeviceIsHub
            connectionInfoEx.DeviceAddress = connectionInfo.DeviceAddress
            connectionInfoEx.NumberOfOpenPipes = connectionInfo.NumberOfOpenPipes
            connectionInfoEx.ConnectionStatus = connectionInfo.ConnectionStatus

            # memcpy(&connectionInfoEx.PipeList[0],
            #    &connectionInfo.PipeList[0],
            #    sizeof(USB_PIPE_INFO) * 30)
            connectionInfoEx.PipeList = connectionInfo.PipeList

            # FREE(connectionInfo)

        # Update the count of connected devices
        #
        if connectionInfoEx.ConnectionStatus == USB_CONNECTION_STATUS.DeviceConnected:
            global TotalDevicesConnected
            TotalDevicesConnected += 1

        if connectionInfoEx.DeviceIsHub:
            global TotalHubs
            TotalHubs += 1

        # If there is a device connected, get the Device Description
        #
        if connectionInfoEx.ConnectionStatus != USB_CONNECTION_STATUS.NoDeviceConnected:
            try:
                driverKeyName = GetDriverKeyName(hHubDevice, index)
            except OSError as err:
                name = f"{connectionInfoEx.DeviceDescriptor.idVendor:04X}:{connectionInfoEx.DeviceDescriptor.idProduct:04X}"
                log.debug(f"Failed to query {name}: {err}")
                driverKeyName = ""

            if driverKeyName:
                cbDriverName = len(driverKeyName)

                # hr = StringCbLength(driverKeyName, MAX_DRIVER_KEY_NAME, byref(cbDriverName))
                # if (SUCCEEDED(hr)):
                DevProps = DriverNameToDeviceProperties(driverKeyName, cbDriverName)
                pNode = FindMatchingDeviceNodeForDriverName(
                    driverKeyName, connectionInfoEx.DeviceIsHub
                )

            #     FREE(driverKeyName)

        # If there is a device connected to the port, try to retrieve the
        # Configuration Descriptor from the device.
        #
        if (  # gDoConfigDesc &&
            connectionInfoEx.ConnectionStatus == USB_CONNECTION_STATUS.DeviceConnected
        ):
            configDescBuff = GetConfigDescriptor(hHubDevice, index, 0)
            configDescReq = cast(
                byref(configDescBuff), PUSB_DESCRIPTOR_REQUEST
            ).contents

        else:
            configDescReq = NULL

        # if (configDesc != NULL &&
        #     connectionInfoEx.DeviceDescriptor.bcdUSB > 0x0200):
        #     bosDesc = GetBOSDescriptor(hHubDevice,
        #                                index)

        # else:
        #     bosDesc = NULL
        stringDescs = []
        Strings = {}
        if configDescReq:
            configDesc = cast(
                byref(configDescReq.Data), PUSB_CONFIGURATION_DESCRIPTOR
            ).contents

        if (
            not connectionInfoEx.DeviceIsHub
            and configDescReq
            and AreThereStringDescriptors(connectionInfoEx.DeviceDescriptor, configDesc)
        ):
            try:
                stringDescs, Strings = GetAllStringDescriptors(
                    hHubDevice,
                    index,
                    connectionInfoEx.DeviceDescriptor,
                    # byref(connectionInfoEx, type(connectionInfoEx).DeviceDescriptor.offset),
                    configDesc,
                )
            except AttributeError as ex:
                name = (
                    str(pNode.DeviceDetailData)
                    if pNode
                    else f"{connectionInfoEx.DeviceDescriptor.idVendor:04X}:{connectionInfoEx.DeviceDescriptor.idProduct:04X}"
                )
                log.debug(f"{ex}: {name}")

        # If the device connected to the port is an external hub, get the
        # name of the external hub and recursively enumerate it.
        #
        if connectionInfoEx.DeviceIsHub:
            extHubName = ""
            cbHubName = 0

            extHubName = GetExternalHubName(hHubDevice, index)
            # extHubName = ""
            if extHubName:
                # hr = StringCbLength(extHubName, MAX_DRIVER_KEY_NAME, byref(cbHubName))
                # if (SUCCEEDED(hr)):
                # cbHubName = len(extHubName)
                EnumerateHub(
                    hTreeParent,  # hPortItem,
                    extHubName,
                    cbHubName,
                    connectionInfoEx,
                    # connectionInfoExV2,
                    # pPortConnectorProps,
                    configDescReq,
                    # bosDesc,
                    stringDescs,
                    Strings,
                    DevProps,
                )

        else:
            # Allocate some space for a USBDEVICEINFO structure to hold the
            # hub info, hub name, and connection info pointers.  GPTR zero
            # initializes the structure for us.
            #
            # info = (USBDEVICEINFO) ALLOC(sizeof(USBDEVICEINFO))
            info = USBDEVICEINFO()

            if info == NULL:
                raise Exception("OOPS")
                # if (configDesc != NULL):
                #     FREE(configDesc)

                # if (bosDesc != NULL):
                #     FREE(bosDesc)

                # FREE(connectionInfoEx)

                # if (pPortConnectorProps != NULL):
                #     FREE(pPortConnectorProps)

                # if (connectionInfoExV2 != NULL):
                #     FREE(connectionInfoExV2)

                # break

            info.DeviceInfoType = USBDEVICEINFOTYPE.DeviceInfo
            info.ConnectionInfo = connectionInfoEx
            # info.PortConnectorProps = pPortConnectorProps
            info.ConfigDesc = configDescReq
            info.StringDescs = stringDescs
            info.Strings = Strings
            # info.BosDesc = bosDesc
            # info.ConnectionInfoV2 = connectionInfoExV2
            info.UsbDeviceProperties = DevProps
            info.DeviceInfoNode = pNode

            # StringCchPrintf(leafName, sizeof(leafName), "[Port%d] ", index)
            leafName = f"[Port{index}] "

            # Add error description if ConnectionStatus is other than NoDeviceConnected / DeviceConnected
            # StringCchCat(leafName,
            #     sizeof(leafName),
            #     ConnectionStatuses[connectionInfoEx.ConnectionStatus])

            if DevProps:
                leafName += DevProps.DeviceDesc

                # size_t cchDeviceDesc = 0

                # hr = StringCbLength(DevProps.DeviceDesc, MAX_DEVICE_PROP, byref(cchDeviceDesc))
                # if (FAILED(hr)):
                #     raise Exception("OOPS")

                # dwSizeOfLeafName = sizeof(leafName)
                # StringCchCatN(leafName,
                #     dwSizeOfLeafName - 1,
                #     " :  ",
                #     sizeof(" :  "))
                # StringCchCatN(leafName,
                #     dwSizeOfLeafName - 1,
                #     DevProps.DeviceDesc,
                #     cchDeviceDesc )

            # if (connectionInfoEx.ConnectionStatus == NoDeviceConnected):
            #     if (connectionInfoExV2 != NULL &&
            #         connectionInfoExV2.SupportedUsbProtocols.Usb300 == 1):
            #         icon = NoSsDeviceIcon

            #     else:
            #         icon = NoDeviceIcon

            # else if (connectionInfoEx.CurrentConfigurationValue):
            #     if (connectionInfoEx.Speed == UsbSuperSpeed):
            #         icon = GoodSsDeviceIcon

            #     else:
            #         icon = GoodDeviceIcon

            # else:
            #     icon = BadDeviceIcon

            if info.UsbDeviceProperties and info.UsbDeviceProperties.DeviceId:
                UsbipdInstanceId = info.UsbDeviceProperties.DeviceId
                if connectionInfoEx.DeviceDescriptor:
                    DeviceDesc = connectionInfoEx.DeviceDescriptor
                    info.Manufacturer = info.Strings.get(DeviceDesc.iManufacturer, "").replace("\x00", "")
                    info.Product = info.Strings.get(DeviceDesc.iProduct, "").replace("\x00", "")
                    if DeviceDesc.iSerialNumber:
                        info.SerialNumber = info.Strings.get(
                            DeviceDesc.iSerialNumber, ""
                        ).replace("\x00", "")

                    # When device is not attached this matches the usbipd InstanceID
                    # Once attached however the VID/PID elements change to relate to the
                    # "filter driver" ?? so can no longer be used to match usbipd ids.
                    vid = DeviceDesc.idVendor
                    pid = DeviceDesc.idProduct
                    unique = info.UsbDeviceProperties.DeviceId.split("\\")[2]
                    UsbipdInstanceId = f"USB\\VID_{vid:04X}&PID_{pid:04X}\\{unique}"

                parsed_devices[UsbipdInstanceId] = info

            hTreeParent.append((leafName, info))
            # AddLeaf(hTreeParent, #hPortItem,
            #                 (LPARAM)info,
            #                 leafName,
            #                 icon)


# for


def GetConfigDescriptor(hHubDevice: HANDLE, ConnectionIndex: int, DescriptorIndex: int):
    success = False
    nBytes = 0
    nBytesReturned = DWORD(0)

    # configDescReqBuf = UCHAR * (sizeof(USB_DESCRIPTOR_REQUEST) +
    #                          sizeof(USB_CONFIGURATION_DESCRIPTOR))
    nBytes = sizeof(USB_DESCRIPTOR_REQUEST) - 1 + sizeof(USB_CONFIGURATION_DESCRIPTOR)

    configDescReq = USB_DESCRIPTOR_REQUEST()

    # Request the Configuration Descriptor the first time using our
    # local buffer, which is just big enough for the Cofiguration
    # Descriptor itself.
    #
    resize(configDescReq, nBytes)

    # configDescReq = (USB_DESCRIPTOR_REQUEST)configDescReqBuf
    # configDesc = (USB_CONFIGURATION_DESCRIPTOR)(configDescReq+1)

    offset = sizeof(USB_DESCRIPTOR_REQUEST) - 1
    configDesc = cast(
        byref(configDescReq, offset), PUSB_CONFIGURATION_DESCRIPTOR
    ).contents

    # Zero fill the entire request structure
    #
    # memset(configDescReq, 0, nBytes)

    # Indicate the port from which the descriptor will be requested
    #
    configDescReq.ConnectionIndex = ConnectionIndex

    #
    # USBHUB uses URB_FUNCTION_GET_DESCRIPTOR_FROM_DEVICE to process this
    # IOCTL_USB_GET_DESCRIPTOR_FROM_NODE_CONNECTION request.
    #
    # USBD will automatically initialize these fields:
    #     bmRequest = 0x80
    #     bRequest  = 0x06
    #
    # We must inititialize these fields:
    #     wValue    = Descriptor Type (high) and Descriptor Index (low byte)
    #     wIndex    = Zero (or Language ID for String Descriptors)
    #     wLength   = Length of descriptor buffer
    #
    configDescReq.SetupPacket.wValue = (
        USB_CONFIGURATION_DESCRIPTOR_TYPE << 8
    ) | DescriptorIndex

    configDescReq.SetupPacket.wLength = (USHORT)(
        nBytes - sizeof(USB_DESCRIPTOR_REQUEST) + 1
    )

    # Now issue the get descriptor request.
    #
    success = DeviceIoControl(
        hHubDevice,
        IOCTL_USB_GET_DESCRIPTOR_FROM_NODE_CONNECTION,
        byref(configDescReq),
        nBytes,
        byref(configDescReq),
        nBytes,
        byref(nBytesReturned),
        NULL,
    )

    if not success:
        raise Exception("OOPS")
        return NULL

    if nBytes != nBytesReturned.value:
        raise Exception("OOPS")
        return NULL

    if configDesc.wTotalLength < sizeof(USB_CONFIGURATION_DESCRIPTOR):
        raise Exception("OOPS")
        return NULL

    # Now request the entire Configuration Descriptor using a dynamically
    # allocated buffer which is sized big enough to hold the entire descriptor
    #
    nBytes = sizeof(USB_DESCRIPTOR_REQUEST) - 1 + configDesc.wTotalLength

    buffer = ctypes.create_string_buffer(b"", nBytes)
    configDescReq = cast(byref(buffer), PUSB_DESCRIPTOR_REQUEST).contents

    if not configDescReq:
        raise Exception("OOPS")
        return NULL

    configDesc = cast(byref(configDescReq.Data), PUSB_CONFIGURATION_DESCRIPTOR).contents

    # Indicate the port from which the descriptor will be requested
    #
    configDescReq.ConnectionIndex = ConnectionIndex

    #
    # USBHUB uses URB_FUNCTION_GET_DESCRIPTOR_FROM_DEVICE to process this
    # IOCTL_USB_GET_DESCRIPTOR_FROM_NODE_CONNECTION request.
    #
    # USBD will automatically initialize these fields:
    #     bmRequest = 0x80
    #     bRequest  = 0x06
    #
    # We must inititialize these fields:
    #     wValue    = Descriptor Type (high) and Descriptor Index (low byte)
    #     wIndex    = Zero (or Language ID for String Descriptors)
    #     wLength   = Length of descriptor buffer
    #
    configDescReq.SetupPacket.wValue = (
        USB_CONFIGURATION_DESCRIPTOR_TYPE << 8
    ) | DescriptorIndex

    configDescReq.SetupPacket.wLength = (USHORT)(
        nBytes - sizeof(USB_DESCRIPTOR_REQUEST) + 1
    )

    # Now issue the get descriptor request.
    #

    success = DeviceIoControl(
        hHubDevice,
        IOCTL_USB_GET_DESCRIPTOR_FROM_NODE_CONNECTION,
        byref(configDescReq),
        nBytes,
        byref(configDescReq),
        nBytes,
        byref(nBytesReturned),
        NULL,
    )

    if not success:
        raise Exception("OOPS")
        FREE(configDescReq)
        return NULL

    if nBytes != nBytesReturned.value:
        raise Exception("OOPS")
        FREE(configDescReq)
        return NULL

    if configDesc.wTotalLength != (nBytes - sizeof(USB_DESCRIPTOR_REQUEST) + 1):
        raise Exception("OOPS")
        FREE(configDescReq)
        return NULL

    return buffer


def AreThereStringDescriptors(DeviceDesc: USB_DEVICE_DESCRIPTOR, allDesc):
    # descEnd = NULL
    # configDesc = cast(byref(commonDesc), PUSB_CONFIGURATION_DESCRIPTOR).contents

    #
    # Check Device Descriptor strings
    #

    if DeviceDesc.iManufacturer or DeviceDesc.iProduct or DeviceDesc.iSerialNumber:
        return True

    #
    # Check the Configuration and Interface Descriptor strings
    #

    descEnd = allDesc.wTotalLength

    # commonDecommonDescON_DESCRIPTOR)ConfigDesc
    offset = 0
    commonDesc = cast(byref(allDesc), PUSB_COMMON_DESCRIPTOR).contents

    while (
        offset + sizeof(USB_COMMON_DESCRIPTOR) < descEnd
        and offset + commonDesc.bLength <= descEnd
    ):

        if commonDesc.bDescriptorType in (
            USB_CONFIGURATION_DESCRIPTOR_TYPE,
            USB_OTHER_SPEED_CONFIGURATION_DESCRIPTOR_TYPE,
        ):
            if commonDesc.bLength != sizeof(USB_CONFIGURATION_DESCRIPTOR):
                raise Exception("OOPS")
                break

            configDesc = cast(byref(commonDesc), PUSB_CONFIGURATION_DESCRIPTOR).contents
            if configDesc.iConfiguration:
                return True

            # commonDesc = (USB_COMMON_DESCRIPTOR) ((PUCHAR) commonDesc + commonDesc.bLength)
            offset += commonDesc.bLength
            commonDesc = cast(byref(allDesc, offset), PUSB_COMMON_DESCRIPTOR).contents
            continue

        elif commonDesc.bDescriptorType == USB_INTERFACE_DESCRIPTOR_TYPE:
            if commonDesc.bLength != sizeof(
                USB_INTERFACE_DESCRIPTOR
            ) and commonDesc.bLength != sizeof(USB_INTERFACE_DESCRIPTOR2):
                raise Exception("OOPS")
                break

            configDesc = cast(byref(commonDesc), PUSB_INTERFACE_DESCRIPTOR).contents
            if configDesc.iInterface:
                return True

            # commonDesc = (USB_COMMON_DESCRIPTOR) ((PUCHAR) commonDesc + commonDesc.bLength)
            offset += commonDesc.bLength
            commonDesc = cast(byref(allDesc, offset), PUSB_COMMON_DESCRIPTOR).contents

            continue

        else:
            # commonDesc = (USB_COMMON_DESCRIPTOR) ((PUCHAR) commonDesc + commonDesc.bLength)
            offset += commonDesc.bLength
            commonDesc = cast(byref(allDesc, offset), PUSB_COMMON_DESCRIPTOR).contents
            continue
        break

    return False


def GetAllStringDescriptors(
    hHubDevice: HANDLE,
    ConnectionIndex: int,
    DeviceDesc: USB_DEVICE_DESCRIPTOR,
    configDesc,
) -> Tuple[List[STRING_DESCRIPTOR_NODE], dict]:

    stringDescs: List[STRING_DESCRIPTOR_NODE] = []
    Strings = {}
    numLanguageIDs = ULONG()  # = 0
    languageIDs = USHORT()  # = NULL

    descEnd = None  # = NULL
    commonDesc = USB_COMMON_DESCRIPTOR()  # = NULL
    uIndex = UCHAR(1)  # = 1
    bInterfaceClass = UCHAR()  # = 0
    getMoreStrings = False
    hr = S_OK

    #
    # Get the array of supported Language IDs, which is returned
    # in String Descriptor 0
    #
    supportedLanguagesString = GetStringDescriptor(hHubDevice, ConnectionIndex, 0, 0)

    if not supportedLanguagesString:
        raise AttributeError("Couldn't read languages string - possibly in sleep state")

    stringDescs.append(supportedLanguagesString)

    numLanguageIDs = (supportedLanguagesString.StringDescriptor.bLength - 2) // 2

    languageIDs = [ord(supportedLanguagesString.StringDescriptor.bString)]

    #
    # Get the Device Descriptor strings
    #

    if DeviceDesc.iManufacturer:
        Strings.update(
            GetStringDescriptors(
                hHubDevice,
                ConnectionIndex,
                DeviceDesc.iManufacturer,
                numLanguageIDs,
                languageIDs,
                stringDescs,
            )
        )

    if DeviceDesc.iProduct:
        Strings.update(
            GetStringDescriptors(
                hHubDevice,
                ConnectionIndex,
                DeviceDesc.iProduct,
                numLanguageIDs,
                languageIDs,
                stringDescs,
            )
        )

    if DeviceDesc.iSerialNumber:
        Strings.update(
            GetStringDescriptors(
                hHubDevice,
                ConnectionIndex,
                DeviceDesc.iSerialNumber,
                numLanguageIDs,
                languageIDs,
                stringDescs,
            )
        )

    #
    # Get the Configuration and Interface Descriptor strings
    #

    # descEnd = (PUCHAR)ConfigDesc + ConfigDesc.wTotalLength
    descEnd = configDesc.wTotalLength
    offset = 0
    commonDesc = cast(byref(configDesc), PUSB_COMMON_DESCRIPTOR).contents
    # commonDesc = (USB_COMMON_DESCRIPTOR)ConfigDesc

    # while ((PUCHAR)commonDesc + sizeof(USB_COMMON_DESCRIPTOR) < descEnd &&
    #        (PUCHAR)commonDesc + commonDesc.bLength <= descEnd):
    while (
        offset + sizeof(USB_COMMON_DESCRIPTOR) < descEnd
        and offset + commonDesc.bLength <= descEnd
    ):

        if commonDesc.bDescriptorType == USB_CONFIGURATION_DESCRIPTOR_TYPE:
            if commonDesc.bLength != sizeof(USB_CONFIGURATION_DESCRIPTOR):
                raise Exception("OOPS")
                break

            confDesc = cast(byref(commonDesc), PUSB_CONFIGURATION_DESCRIPTOR).contents
            if confDesc.iConfiguration:
                GetStringDescriptors(
                    hHubDevice,
                    ConnectionIndex,
                    confDesc.iConfiguration,
                    numLanguageIDs,
                    languageIDs,
                    stringDescs,
                )

            offset += commonDesc.bLength
            commonDesc = cast(
                byref(configDesc, offset), PUSB_COMMON_DESCRIPTOR
            ).contents

            continue

        elif commonDesc.bDescriptorType == USB_IAD_DESCRIPTOR_TYPE:
            if commonDesc.bLength < sizeof(USB_IAD_DESCRIPTOR):
                raise Exception("OOPS")
                break

            iadDesc = cast(byref(commonDesc), POINTER(USB_IAD_DESCRIPTOR)).contents
            if iadDesc.iFunction:
                GetStringDescriptors(
                    hHubDevice,
                    ConnectionIndex,
                    iadDesc.iFunction,
                    numLanguageIDs,
                    languageIDs,
                    stringDescs,
                )

            offset += commonDesc.bLength
            commonDesc = cast(
                byref(configDesc, offset), PUSB_COMMON_DESCRIPTOR
            ).contents
            continue

        elif commonDesc.bDescriptorType == USB_INTERFACE_DESCRIPTOR_TYPE:
            if commonDesc.bLength != sizeof(
                USB_INTERFACE_DESCRIPTOR
            ) and commonDesc.bLength != sizeof(USB_INTERFACE_DESCRIPTOR2):
                raise Exception("OOPS")
                break

            intDesc = cast(byref(commonDesc), PUSB_INTERFACE_DESCRIPTOR).contents
            if intDesc.iInterface:
                GetStringDescriptors(
                    hHubDevice,
                    ConnectionIndex,
                    intDesc.iInterface,
                    numLanguageIDs,
                    languageIDs,
                    stringDescs,
                )

            #
            # We need to display more string descriptors for the following
            # interface classes
            #
            bInterfaceClass = intDesc.bInterfaceClass
            if bInterfaceClass == USB_DEVICE_CLASS_VIDEO:
                getMoreStrings = True

            offset += commonDesc.bLength
            commonDesc = cast(
                byref(configDesc, offset), PUSB_COMMON_DESCRIPTOR
            ).contents
            # commonDesc = (USB_COMMON_DESCRIPTOR) ((PUCHAR) commonDesc + commonDesc.bLength)
            continue

        else:
            offset += commonDesc.bLength
            commonDesc = cast(
                byref(configDesc, offset), PUSB_COMMON_DESCRIPTOR
            ).contents
            # commonDesc = (USB_COMMON_DESCRIPTOR) ((PUCHAR) commonDesc + commonDesc.bLength)
            continue

        break

    # if (getMoreStrings):
    #     #
    #     # We might need to display strings later that are referenced only in
    #     # class-specific descriptors. Get String Descriptors 1 through 32 (an
    #     # arbitrary upper limit for Strings needed due to "bad devices"
    #     # returning an infinite repeat of Strings 0 through 4) until one is not
    #     # found.
    #     #
    #     # There are also "bad devices" that have issues even querying 1-32, but
    #     # historically USBView made this query, so the query should be safe for
    #     # video devices.
    #     #
    #     for (uIndex = 1 SUCCEEDED(hr) && (uIndex < NUM_STRING_DESC_TO_GET) uIndex++):
    #         hr = GetStringDescriptors(hHubDevice,
    #                                   ConnectionIndex,
    #                                   uIndex,
    #                                   numLanguageIDs,
    #                                   languageIDs,
    #                                   supportedLanguagesString)

    return stringDescs, Strings


# *****************************************************************************
#
# GetStringDescriptor()
#
# hHubDevice - Handle of the hub device containing the port from which the
# String Descriptor will be requested.
#
# ConnectionIndex - Identifies the port on the hub to which a device is
# attached from which the String Descriptor will be requested.
#
# DescriptorIndex - String Descriptor index.
#
# LanguageID - Language in which the string should be requested.
#
# *****************************************************************************


def GetStringDescriptor(
    hHubDevice: HANDLE, ConnectionIndex: int, DescriptorIndex: int, LanguageID: int
) -> STRING_DESCRIPTOR_NODE:
    success = 0
    nBytes = 0
    nBytesReturned = DWORD(0)

    stringDescReqBuf = ctypes.create_string_buffer(
        b"", sizeof(USB_DESCRIPTOR_REQUEST) - 1 + MAXIMUM_USB_STRING_LENGTH
    )
    # UCHAR   stringDescReqBuf[sizeof(USB_DESCRIPTOR_REQUEST) +
    #  MAXIMUM_USB_STRING_LENGTH]

    stringDescReq = USB_DESCRIPTOR_REQUEST()
    stringDesc = USB_STRING_DESCRIPTOR()
    stringDescNode = STRING_DESCRIPTOR_NODE()

    nBytes = sizeof(stringDescReqBuf)

    # stringDescReq = (USB_DESCRIPTOR_REQUEST)stringDescReqBuf
    # stringDesc = (USB_STRING_DESCRIPTOR)(stringDescReq+1)
    stringDescReq = cast(byref(stringDescReqBuf), PUSB_DESCRIPTOR_REQUEST).contents
    stringDesc = cast(byref(stringDescReq.Data), PUSB_STRING_DESCRIPTOR).contents

    # Zero fill the entire request structure
    #
    # memset(stringDescReq, 0, nBytes)

    # Indicate the port from which the descriptor will be requested
    #
    stringDescReq.ConnectionIndex = ConnectionIndex

    #
    # USBHUB uses URB_FUNCTION_GET_DESCRIPTOR_FROM_DEVICE to process this
    # IOCTL_USB_GET_DESCRIPTOR_FROM_NODE_CONNECTION request.
    #
    # USBD will automatically initialize these fields:
    #     bmRequest = 0x80
    #     bRequest  = 0x06
    #
    # We must inititialize these fields:
    #     wValue    = Descriptor Type (high) and Descriptor Index (low byte)
    #     wIndex    = Zero (or Language ID for String Descriptors)
    #     wLength   = Length of descriptor buffer
    #
    stringDescReq.SetupPacket.wValue = (
        USB_STRING_DESCRIPTOR_TYPE << 8
    ) | DescriptorIndex

    stringDescReq.SetupPacket.wIndex = LanguageID

    stringDescReq.SetupPacket.wLength = nBytes - sizeof(USB_DESCRIPTOR_REQUEST) + 1

    # Now issue the get descriptor request.
    #
    success = DeviceIoControl(
        hHubDevice,
        IOCTL_USB_GET_DESCRIPTOR_FROM_NODE_CONNECTION,
        byref(stringDescReq),
        nBytes,
        byref(stringDescReq),
        nBytes,
        byref(nBytesReturned),
        NULL,
    )

    #
    # Do some sanity checks on the return from the get descriptor request.
    #

    if not success:
        error = GetLastError()
        if DEBUG:
            log.error(WinError(error))
        if error == 31:
            # PermissionError(13, 'A device attached to the system is not functioning.', None, 31)
            # Device likely in sleep state: https://stackoverflow.com/a/60017122
            raise AttributeError(
                "Couldn't read descriptor - device likely in sleep state"
            )
        return None

    if nBytesReturned.value < 2:
        raise Exception("OOPS")
        return NULL

    if stringDesc.bDescriptorType != USB_STRING_DESCRIPTOR_TYPE:
        raise Exception("OOPS")
        return NULL

    if stringDesc.bLength != nBytesReturned.value - sizeof(USB_DESCRIPTOR_REQUEST) + 1:
        raise Exception("OOPS")
        return NULL

    if stringDesc.bLength % 2 != 0:
        raise Exception("OOPS")
        return NULL

    #
    # Looks good, allocate some (zero filled) space for the string descriptor
    # node and copy the string descriptor to it.
    #

    # stringDescNode = (PSTRING_DESCRIPTOR_NODE)ALLOC(sizeof(STRING_DESCRIPTOR_NODE) +
    # stringDesc.bLength)
    stringDescNode = STRING_DESCRIPTOR_NODE()
    # resize(stringDescNode, sizeof(STRING_DESCRIPTOR_NODE) + stringDesc.bLength)

    if stringDescNode == NULL:
        raise Exception("OOPS")
        return NULL

    stringDescNode.DescriptorIndex = DescriptorIndex
    stringDescNode.LanguageID = LanguageID

    # tBuff = c_char * (sizeof(USB_STRING_DESCRIPTOR) - 1 + stringDesc.bLength)
    # USB_STRING_DESCRIPTOR(bytearray(cast(byref(stringDesc), POINTER(tBuff))))

    stringDescLen = stringDesc.bLength
    stringDescS = USB_STRING_DESCRIPTOR()
    resize(stringDescS, stringDescLen)

    desc_offset = USB_DESCRIPTOR_REQUEST.Data.offset
    raw = stringDescReqBuf[desc_offset : desc_offset + stringDescLen]
    cast(byref(stringDescS), POINTER(c_char * stringDescLen)).contents[:] = raw

    stringDescNode.StringDescriptor = stringDescS

    return stringDescNode


# *****************************************************************************
#
# GetStringDescriptors()
#
# hHubDevice - Handle of the hub device containing the port from which the
# String Descriptor will be requested.
#
# ConnectionIndex - Identifies the port on the hub to which a device is
# attached from which the String Descriptor will be requested.
#
# DescriptorIndex - String Descriptor index.
#
# NumLanguageIDs -  Number of languages in which the string should be
# requested.
#
# LanguageIDs - Languages in which the string should be requested.
#
# StringDescNodeHead - First node in linked list of device's string descriptors
#
# Return Value: HRESULT indicating whether the string is on the list
#
# *****************************************************************************


def GetStringDescriptors(
    hHubDevice: HANDLE,
    ConnectionIndex: int,
    DescriptorIndex: int,
    NumLanguageIDs: int,
    LanguageIDs: List[int],
    StringDescList: List[STRING_DESCRIPTOR_NODE],
) -> dict:

    # tail = StringDescList
    # trailing: STRING_DESCRIPTOR_NODE = None
    i = 0

    #
    # Go to the end of the linked list, searching for the requested index to
    # see if we've already retrieved it
    #
    for node in StringDescList:
        # for (tail = StringDescNodeHead; tail != NULL; tail = tail.Next):
        if node.DescriptorIndex == DescriptorIndex:
            return {node.DescriptorIndex: str(node.StringDescriptor)}

        # trailing = tail
        # tail = tail.Next

    # tail = trailing

    #
    # Get the next String Descriptor. If this is NULL, then we're done (return)
    # Otherwise, loop through all Language IDs
    #
    # for (i = 0 (tail != NULL) and (i < NumLanguageIDs) i++):

    if NumLanguageIDs > 1:
        if DEBUG:
            log.debug("Multiple language strings not fully supported")

    found = {}
    for i in range(NumLanguageIDs):
        # if not tail:
        #     break
        node = GetStringDescriptor(
            hHubDevice, ConnectionIndex, DescriptorIndex, LanguageIDs[i]
        )

        StringDescList.append(node)
        if node.DescriptorIndex not in found:
            found[node.DescriptorIndex] = str(node.StringDescriptor)

    # if (tail == NULL):
    # return E_FAIL
    # else:
    return found


def GetDriverKeyName(Hub: HANDLE, ConnectionIndex: int):
    success = False
    nBytes = ULONG(0)
    driverKeyName = USB_NODE_CONNECTION_DRIVERKEY_NAME()
    # driverKeyNameW = USB_NODE_CONNECTION_DRIVERKEY_NAME()
    driverKeyNameW = driverKeyName
    # driverKeyNameA = ""

    # Get the length of the name of the driver key of the device attached to
    # the specified port.
    #
    driverKeyName.ConnectionIndex = ConnectionIndex

    success = DeviceIoControl(
        Hub,
        IOCTL_USB_GET_NODE_CONNECTION_DRIVERKEY_NAME,
        byref(driverKeyName),
        sizeof(driverKeyName),
        byref(driverKeyName),
        sizeof(driverKeyName),
        byref(nBytes),
        NULL,
    )

    if not success:
        error = GetLastError()
        raise WinError(error)
        # goto GetDriverKeyNameError;

    # Allocate space to hold the driver key name
    #
    nBytes = ULONG(driverKeyName.ActualLength)

    if nBytes.value <= sizeof(driverKeyName):
        error = GetLastError()
        raise Exception(f"OOPS: {error}")
        # goto GetDriverKeyNameError;

    # driverKeyNameW = ALLOC(nBytes);
    resize(driverKeyNameW, nBytes.value)
    if not driverKeyNameW:
        error = GetLastError()
        raise Exception(f"OOPS: {error}")
        # goto GetDriverKeyNameError;

    # Get the name of the driver key of the device attached to
    # the specified port.
    #
    driverKeyNameW.ConnectionIndex = ConnectionIndex

    success = DeviceIoControl(
        Hub,
        IOCTL_USB_GET_NODE_CONNECTION_DRIVERKEY_NAME,
        byref(driverKeyNameW),
        nBytes,
        byref(driverKeyNameW),
        nBytes,
        byref(nBytes),
        NULL,
    )

    if not success:
        error = GetLastError()
        raise WinError(get_last_error())
        # goto GetDriverKeyNameError;

    # Convert the driver key name
    #
    return str(driverKeyNameW)
    # driverKeyNameA = WideStrToMultiStr(driverKeyNameW->DriverKeyName, nBytes - sizeof(USB_NODE_CONNECTION_DRIVERKEY_NAME) + sizeof(WCHAR));

    # # All done, free the uncoverted driver key name and return the
    # # converted driver key name
    # #
    # FREE(driverKeyNameW)

    # return driverKeyNameA


# GetDriverKeyNameError:
#     # There was an error, free anything that was allocated
#     #
#     if (driverKeyNameW != NULL):
#         FREE(driverKeyNameW);
#         driverKeyNameW = NULL;

#     return NULL


def GetHCDDriverKeyName(HCD: HANDLE):
    nBytes = ULONG(0)
    driverKeyName = USB_HCD_DRIVERKEY_NAME()
    driverKeyNameW = USB_HCD_DRIVERKEY_NAME()
    # driverKeyNameA = PCHAR()
    # ZeroMemory(byref(driverKeyName), sizeof(driverKeyName))

    # Get the length of the name of the driver key of the HCD
    #
    success = DeviceIoControl(
        HCD,
        IOCTL_GET_HCD_DRIVERKEY_NAME,
        byref(driverKeyName),
        sizeof(driverKeyName),
        byref(driverKeyName),
        sizeof(driverKeyName),
        byref(nBytes),
        NULL,
    )

    error = GetLastError()
    if not success:
        error = GetLastError()
        raise Exception(f"OOPS: {error}")
        # goto GetHCDDriverKeyNameError

    # Allocate space to hold the driver key name
    #
    nBytes = DWORD(driverKeyName.ActualLength)
    if nBytes.value <= sizeof(driverKeyName):
        error = GetLastError()
        raise Exception(f"OOPS: {error}")
        # goto GetHCDDriverKeyNameError

    # driverKeyNameW = ALLOC(nBytes)
    # driverKeyNameW = ctypes.create_unicode_buffer(nBytes.value)
    resize(driverKeyNameW, nBytes.value)
    if not driverKeyNameW:
        error = GetLastError()
        raise Exception(f"OOPS: {error}")
        # goto GetHCDDriverKeyNameError

    # Get the name of the driver key of the device attached to
    # the specified port.
    #

    success = DeviceIoControl(
        HCD,
        IOCTL_GET_HCD_DRIVERKEY_NAME,
        byref(driverKeyNameW),
        nBytes,
        byref(driverKeyNameW),
        nBytes,
        byref(nBytes),
        NULL,
    )
    if not success:
        error = GetLastError()
        raise Exception(f"OOPS: {error}")
        # goto GetHCDDriverKeyNameError

    return str(driverKeyNameW)
    # return driverKeyNameW.DriverKeyName

    # #
    # # Convert the driver key name
    # # Pass the length of the DriverKeyName string
    # #

    # driverKeyNameA = WideStrToMultiStr(driverKeyNameW.DriverKeyName, nBytes - sizeof(USB_HCD_DRIVERKEY_NAME) + sizeof(WCHAR))

    # # All done, free the uncoverted driver key name and return the
    # # converted driver key name
    # #
    # FREE(driverKeyNameW)

    # return driverKeyNameA


def GetExternalHubName(Hub: HANDLE, ConnectionIndex: int):
    nBytes = ULONG(0)
    extHubName = USB_NODE_CONNECTION_NAME()
    # extHubNameW = PUSB_NODE_CONNECTION_NAME() = NULL
    # extHubNameA = PCHAR() = NULL

    # Get the length of the name of the external hub attached to the
    # specified port.
    #
    extHubName.ConnectionIndex = ConnectionIndex

    success = DeviceIoControl(
        Hub,
        IOCTL_USB_GET_NODE_CONNECTION_NAME,
        byref(extHubName),
        sizeof(extHubName),
        byref(extHubName),
        sizeof(extHubName),
        byref(nBytes),
        NULL,
    )

    if not success:
        error = GetLastError()
        if DEBUG:
            log.error(WinError(get_last_error()))
        raise Exception("OOPS")

    # Allocate space to hold the external hub name
    #
    nBytes = DWORD(extHubName.ActualLength)

    if nBytes.value <= sizeof(extHubName):
        raise Exception("OOPS")

    # extHubNameW = ALLOC(nBytes)
    resize(extHubName, nBytes.value)

    if not extHubName:
        raise Exception("OOPS")

    # Get the name of the external hub attached to the specified port
    #
    extHubName.ConnectionIndex = ConnectionIndex

    success = DeviceIoControl(
        Hub,
        IOCTL_USB_GET_NODE_CONNECTION_NAME,
        byref(extHubName),
        nBytes,
        byref(extHubName),
        nBytes,
        byref(nBytes),
        NULL,
    )

    if not success:
        error = GetLastError()
        if DEBUG:
            log.error(WinError(get_last_error()))
        raise Exception("OOPS")

    # Convert the External Hub name
    #
    # extHubNameA = WideStrToMultiStr(extHubNameW->NodeName, nBytes - sizeof(USB_NODE_CONNECTION_NAME) + sizeof(WCHAR))

    # All done, free the uncoverted external hub name and return the
    # converted external hub name
    #
    # FREE(extHubNameW)

    return str(extHubName)


def DriverNameToDeviceInst(
    DriverName: str,
    cbDriverName: int,
    # pDevInfo: POINTER(HDEVINFO),
    # pDevInfoData: PSP_DEVINFO_DATA
):

    deviceInfo: HDEVINFO = INVALID_HANDLE_VALUE
    status: BOOL = True
    deviceIndex = ULONG(0)
    deviceInfoData = SP_DEVINFO_DATA()
    bResult = False
    pDriverName = ""
    buf: PSTR = NULL
    done: BOOL = FALSE

    # if (pDevInfo == NULL):
    #     return FALSE;

    # if (pDevInfoData == NULL):
    #     return FALSE;

    # memset(pDevInfoData, 0, sizeof(SP_DEVINFO_DATA))

    # pDevInfo = INVALID_HANDLE_VALUE

    # Use local string to guarantee zero termination
    # pDriverName = (PCHAR) ALLOC((DWORD) cbDriverName + 1)
    # pDriverName = ctypes.create_string_buffer(DriverName.encode(), cbDriverName + 1)
    # if (NULL == pDriverName):
    #     status = FALSE
    #     error = GetLastError()
    #     log.error(WinError(get_last_error()))
    # StringCbCopyN(pDriverName, cbDriverName + 1, DriverName, cbDriverName)

    #
    # We cannot walk the device tree with CM_Get_Sibling etc. unless we assume
    # the device tree will stabilize. Any devnode removal (even outside of USB)
    # would force us to retry. Instead we use Setup API to snapshot all
    # devices.
    #

    # Examine all present devices to see if any match the given DriverName
    #
    deviceInfo = SetupDiGetClassDevs(NULL, NULL, NULL, DIGCF_ALLCLASSES | DIGCF_PRESENT)

    if deviceInfo == INVALID_HANDLE_VALUE:
        status = False
        error = GetLastError()
        if DEBUG:
            log.error(WinError(get_last_error()))

    deviceIndex = 0
    deviceInfoData.cbSize = sizeof(deviceInfoData)

    while not done:
        #
        # Get devinst of the next device
        #

        status = SetupDiEnumDeviceInfo(deviceInfo, deviceIndex, byref(deviceInfoData))

        deviceIndex += 1

        if not status:
            #
            # This could be an error, or indication that all devices have been
            # processed. Either way the desired device was not found.
            #

            done = True
            break

        #
        # Get the DriverName value
        #

        bResult, buf = GetDeviceProperty(deviceInfo, deviceInfoData, SPDRP_DRIVER)
        # byref(buf))

        # If the DriverName value matches, return the DeviceInstance
        #
        if bResult and buf and DriverName == buf:
            done = True
            # pDevInfo = deviceInfo
            # CopyMemory(pDevInfoData, &deviceInfoData, sizeof(deviceInfoData))
            # FREE(buf)
            break

        if buf != NULL:
            # FREE(buf)
            buf = NULL

    # Done:

    #     if (bResult == FALSE):
    #         if (deviceInfo != INVALID_HANDLE_VALUE):
    #             SetupDiDestroyDeviceInfoList(deviceInfo)

    #     if (pDriverName != NULL):
    #         FREE(pDriverName)

    return status, deviceInfo, deviceInfoData


def FindMatchingDeviceNodeForDriverName(
    DriverKeyName: str, IsHub: bool
) -> Optional[DEVICE_INFO_NODE]:
    pNode: PDEVICE_INFO_NODE = NULL
    pList: PDEVICE_GUID_LIST = NULL
    pEntry: PLIST_ENTRY = NULL

    pList = gHubList if IsHub else gDeviceList

    for pEntry in pList:
        pNode: DEVICE_INFO_NODE = pEntry
        if DriverKeyName == pNode.DeviceDriverName:
            return pNode

    return None


def GetDeviceProperty(
    DeviceInfoSet: HDEVINFO, DeviceInfoData: SP_DEVINFO_DATA, Property: DWORD
):
    # ) - > Tuple[bool, str]:

    bResult = BOOL(0)
    requiredLength = DWORD(0)
    lastError = DWORD(0)

    ppBuffer = ""

    bResult = SetupDiGetDeviceRegistryProperty(
        DeviceInfoSet,
        byref(DeviceInfoData),
        Property,
        NULL,
        NULL,
        0,
        byref(requiredLength),
    )
    lastError = GetLastError()

    if (requiredLength.value == 0) or (
        bResult != FALSE and lastError != ERROR_INSUFFICIENT_BUFFER
    ):
        return False, ppBuffer

    ppBuffer = ctypes.create_unicode_buffer(requiredLength.value)

    bResult = SetupDiGetDeviceRegistryProperty(
        DeviceInfoSet,
        byref(DeviceInfoData),
        Property,
        NULL,
        byref(ppBuffer),
        requiredLength,
        byref(requiredLength),
    )
    if not bResult:
        ppBuffer = ""
        return False, ppBuffer

    return True, ppBuffer.value


def GetInstanceId(deviceInfo: HDEVINFO, deviceInfoData: SP_DEVINFO_DATA):
    length = DWORD(0)
    status = SetupDiGetDeviceInstanceId(
        deviceInfo, byref(deviceInfoData), NULL, 0, byref(length)
    )
    lastError = GetLastError()

    if status != FALSE and lastError != ERROR_INSUFFICIENT_BUFFER:
        status = FALSE
        # goto Done
        error = GetLastError()
        log.error(WinError(get_last_error()))

    #
    # An extra byte is required for the terminating character
    #

    length.value += 1
    buffer = ctypes.create_unicode_buffer("", length.value)
    # DevProps.DeviceId = ctypes.create_string_buffer(b"", length.value)

    # if (DevProps.DeviceId == NULL):
    #     status = FALSE
    #     #goto Done
    #     error = GetLastError()
    #     log.error(WinError(get_last_error()))

    status = SetupDiGetDeviceInstanceId(
        deviceInfo, byref(deviceInfoData), byref(buffer), length, byref(length)
    )
    if not status:
        # goto Done
        error = GetLastError()
        log.error(WinError(get_last_error()))

    return buffer.value


def main():
    import time

    start = time.time()
    devices, tree = InspectUsbDevices()
    finished = time.time() - start
    return
