import ctypes
from ctypes import *
from ctypes.wintypes import *

_ole32 = oledll.ole32

_StringFromCLSID = _ole32.StringFromCLSID
_CoTaskMemFree = windll.ole32.CoTaskMemFree

NULL = None
# HDEVINFO = ctypes.c_int
HDEVINFO = c_void_p
BOOL = ctypes.c_int
CHAR = ctypes.c_char
PCTSTR = ctypes.c_char_p
HWND = ctypes.c_uint
DWORD = ctypes.c_ulong
PDWORD = ctypes.POINTER(DWORD)
UCHAR = ctypes.c_uint8
ULONG = ctypes.c_ulong
ULONG_PTR = ctypes.POINTER(ULONG)
#~ PBYTE = ctypes.c_char_p
PBYTE = ctypes.c_void_p
LPSECURITY_ATTRIBUTES = LPVOID
LPOVERLAPPED = LPVOID

USB_DEVICE_CLASS_VIDEO = 0x0E

class USB_HUB_NODE:
    UsbHub = 0
    UsbMIParent = 1

def ValidHandle(value):
    if value == 0:
        raise ctypes.WinError()
    return value


"""Flags controlling what is included in the device information set built by SetupDiGetClassDevs"""
DIGCF_DEFAULT = 0x00000001
DIGCF_PRESENT = 0x00000002
DIGCF_ALLCLASSES = 0x00000004
DIGCF_PROFILE = 0x00000008
DIGCF_DEVICE_INTERFACE = 0x00000010
DIGCF_DEVICEINTERFACE  = 0x00000010

"""Flags controlling File acccess"""
GENERIC_WRITE = (1073741824)
GENERIC_READ = (-2147483648)
FILE_SHARE_READ = 1
FILE_SHARE_WRITE = 2
OPEN_EXISTING = 3
OPEN_ALWAYS = 4
FILE_ATTRIBUTE_NORMAL = 128
FILE_FLAG_OVERLAPPED = 1073741824

INVALID_HANDLE_VALUE = HANDLE(-1)

""" USB PIPE TYPE """
PIPE_TYPE_CONTROL = 0
PIPE_TYPE_ISO = 1
PIPE_TYPE_BULK = 2
PIPE_TYPE_INTERRUPT = 3

""" Device registry property codes """
SPDRP_DEVICEDESC                  = (0x00000000)  # DeviceDesc (R/W)
SPDRP_HARDWAREID                  = (0x00000001)  # HardwareID (R/W)
SPDRP_COMPATIBLEIDS               = (0x00000002)  # CompatibleIDs (R/W)
SPDRP_SERVICE                     = (0x00000004)  # Service (R/W)
SPDRP_CLASS                       = (0x00000007)  # Class (R--tied to ClassGUID)
SPDRP_CLASSGUID                   = (0x00000008)  # ClassGUID (R/W)
SPDRP_DRIVER                      = (0x00000009)  # Driver (R/W)
SPDRP_CONFIGFLAGS                 = (0x0000000A)  # ConfigFlags (R/W)
SPDRP_MFG                         = (0x0000000B)  # Mfg (R/W)
SPDRP_FRIENDLYNAME                = (0x0000000C)  # FriendlyName (R/W)
SPDRP_LOCATION_INFORMATION        = (0x0000000D)  # LocationInformation (R/W)
SPDRP_PHYSICAL_DEVICE_OBJECT_NAME = (0x0000000E)  # PhysicalDeviceObjectName (R)
SPDRP_CAPABILITIES                = (0x0000000F)  # Capabilities (R)
SPDRP_UI_NUMBER                   = (0x00000010)  # UiNumber (R)
SPDRP_UPPERFILTERS                = (0x00000011)  # UpperFilters (R/W)
SPDRP_LOWERFILTERS                = (0x00000012)  # LowerFilters (R/W)
SPDRP_BUSTYPEGUID                 = (0x00000013)  # BusTypeGUID (R)
SPDRP_LEGACYBUSTYPE               = (0x00000014)  # LegacyBusType (R)
SPDRP_BUSNUMBER                   = (0x00000015)  # BusNumber (R)
SPDRP_ENUMERATOR_NAME             = (0x00000016)  # Enumerator Name (R)
SPDRP_SECURITY                    = (0x00000017)  # Security (R/W, binary form)
SPDRP_SECURITY_SDS                = (0x00000018)  # Security (W, SDS form)
SPDRP_DEVTYPE                     = (0x00000019)  # Device Type (R/W)
SPDRP_EXCLUSIVE                   = (0x0000001A)  # Device is exclusive-access (R/W)
SPDRP_CHARACTERISTICS             = (0x0000001B)  # Device Characteristics (R/W)
SPDRP_ADDRESS                     = (0x0000001C)  # Device Address (R)
SPDRP_UI_NUMBER_DESC_FORMAT       = (0X0000001D)  # UiNumberDescFormat (R/W)
SPDRP_DEVICE_POWER_DATA           = (0x0000001E)  # Device Power Data (R)
SPDRP_REMOVAL_POLICY              = (0x0000001F)  # Removal Policy (R)
SPDRP_REMOVAL_POLICY_HW_DEFAULT   = (0x00000020)  # Hardware Removal Policy (R)
SPDRP_REMOVAL_POLICY_OVERRIDE     = (0x00000021)  # Removal Policy Override (RW)
SPDRP_INSTALL_STATE               = (0x00000022)  # Device Install State (R)
SPDRP_LOCATION_PATHS              = (0x00000023)  # Device Location Paths (R)
SPDRP_BASE_CONTAINERID            = (0x00000024)  # Base ContainerID (R)
SPDRP_MAXIMUM_PROPERTY            = (0x00000025)  # Upper bound on ordinals


""" Errors """
ERROR_IO_INCOMPLETE = 996
ERROR_IO_PENDING = 997
ERROR_NO_MORE_ITEMS = 259
ERROR_INSUFFICIENT_BUFFER = 122

""" Function codes for user mode IOCTLs """
HCD_GET_STATS_1                    = 255
HCD_DIAGNOSTIC_MODE_ON             = 256
HCD_DIAGNOSTIC_MODE_OFF            = 257
HCD_GET_ROOT_HUB_NAME              = 258
HCD_GET_DRIVERKEY_NAME             = 265
HCD_GET_STATS_2                    = 266
HCD_DISABLE_PORT                   = 268
HCD_ENABLE_PORT                    = 269
HCD_USER_REQUEST                   = 270
HCD_TRACE_READ_REQUEST             = 275

USB_GET_NODE_INFORMATION                  = 258
USB_GET_NODE_CONNECTION_INFORMATION       = 259
USB_GET_DESCRIPTOR_FROM_NODE_CONNECTION   = 260
USB_GET_NODE_CONNECTION_NAME              = 261
USB_GET_NODE_CONNECTION_DRIVERKEY_NAME    = 264
USB_GET_HUB_CAPABILITIES                  = 271
USB_GET_NODE_CONNECTION_ATTRIBUTES        = 272
USB_HUB_CYCLE_PORT                        = 273
USB_GET_NODE_CONNECTION_INFORMATION_EX    = 274
USB_RESET_HUB                             = 275
USB_GET_HUB_CAPABILITIES_EX               = 276
USB_GET_HUB_INFORMATION_EX                = 277
USB_GET_PORT_CONNECTOR_PROPERTIES         = 278
USB_GET_NODE_CONNECTION_INFORMATION_EX_V2 = 279

""" Descriptor Types """
USB_DEVICE_DESCRIPTOR_TYPE                          = 0x01
USB_CONFIGURATION_DESCRIPTOR_TYPE                   = 0x02
USB_STRING_DESCRIPTOR_TYPE                          = 0x03
USB_INTERFACE_DESCRIPTOR_TYPE                       = 0x04
USB_ENDPOINT_DESCRIPTOR_TYPE                        = 0x05
USB_DEVICE_QUALIFIER_DESCRIPTOR_TYPE                = 0x06
USB_OTHER_SPEED_CONFIGURATION_DESCRIPTOR_TYPE       = 0x07
USB_INTERFACE_POWER_DESCRIPTOR_TYPE                 = 0x08
USB_OTG_DESCRIPTOR_TYPE                                     = 0x09
USB_DEBUG_DESCRIPTOR_TYPE                                   = 0x0A
USB_INTERFACE_ASSOCIATION_DESCRIPTOR_TYPE                   = 0x0B
USB_BOS_DESCRIPTOR_TYPE                                     = 0x0F
USB_DEVICE_CAPABILITY_DESCRIPTOR_TYPE                       = 0x10
USB_SUPERSPEED_ENDPOINT_COMPANION_DESCRIPTOR_TYPE           = 0x30
USB_SUPERSPEEDPLUS_ISOCH_ENDPOINT_COMPANION_DESCRIPTOR_TYPE = 0x31

USB_IAD_DESCRIPTOR_TYPE                       = 0x0B

""" Define the method codes for how buffers are passed for I/O and FS controls """
METHOD_BUFFERED                = 0
METHOD_IN_DIRECT               = 1
METHOD_OUT_DIRECT              = 2
METHOD_NEITHER                 = 3

""" Define the access check value for any access """
FILE_ANY_ACCESS               =  0
FILE_SPECIAL_ACCESS   = (FILE_ANY_ACCESS)
FILE_READ_ACCESS        =  ( 0x0001 )    # file & pipe
FILE_WRITE_ACCESS       =  ( 0x0002 )    # file & pipe

def CTL_CODE(DeviceType, Function, Method, Access):
    return (DeviceType << 16) | (Access << 14) | (Function << 2) | Method

def USB_CTL(id):
   # CTL_CODE(FILE_DEVICE_USB, (id), METHOD_BUFFERED, FILE_ANY_ACCESS)
    return CTL_CODE(0x22, id, 0, 0)

FILE_DEVICE_UNKNOWN             = 0x00000022
FILE_DEVICE_USB        = FILE_DEVICE_UNKNOWN

IOCTL_USB_GET_NODE_CONNECTION_DRIVERKEY_NAME = CTL_CODE(FILE_DEVICE_USB,
                                USB_GET_NODE_CONNECTION_DRIVERKEY_NAME,
                                METHOD_BUFFERED,
                                FILE_ANY_ACCESS)

IOCTL_GET_HCD_DRIVERKEY_NAME = CTL_CODE(FILE_DEVICE_USB,
                                                HCD_GET_DRIVERKEY_NAME,
                                                METHOD_BUFFERED,
                                                FILE_ANY_ACCESS)

IOCTL_USB_GET_ROOT_HUB_NAME = CTL_CODE(FILE_DEVICE_USB,
                                                HCD_GET_ROOT_HUB_NAME,
                                                METHOD_BUFFERED,
                                                FILE_ANY_ACCESS)

IOCTL_USB_GET_NODE_INFORMATION = CTL_CODE(FILE_DEVICE_USB,
                                USB_GET_NODE_INFORMATION,
                                METHOD_BUFFERED,
                                FILE_ANY_ACCESS)

IOCTL_USB_GET_HUB_INFORMATION_EX  \
                             = CTL_CODE(FILE_DEVICE_USB,
                                   USB_GET_HUB_INFORMATION_EX,
                                   METHOD_BUFFERED,
                                   FILE_ANY_ACCESS)

IOCTL_USB_GET_HUB_CAPABILITIES_EX  \
                             = CTL_CODE(FILE_DEVICE_USB,
                                    USB_GET_HUB_CAPABILITIES_EX,
                                    METHOD_BUFFERED,
                                    FILE_ANY_ACCESS)

IOCTL_USB_GET_PORT_CONNECTOR_PROPERTIES  \
                             = CTL_CODE(FILE_DEVICE_USB,
                                    USB_GET_PORT_CONNECTOR_PROPERTIES,
                                    METHOD_BUFFERED,
                                    FILE_ANY_ACCESS)

IOCTL_USB_GET_NODE_CONNECTION_INFORMATION_EX_V2  \
                             = CTL_CODE(FILE_DEVICE_USB,
                                    USB_GET_NODE_CONNECTION_INFORMATION_EX_V2,
                                    METHOD_BUFFERED,
                                    FILE_ANY_ACCESS)

IOCTL_USB_GET_NODE_CONNECTION_INFORMATION_EX  \
                             = CTL_CODE(FILE_DEVICE_USB,
                                USB_GET_NODE_CONNECTION_INFORMATION_EX,
                                METHOD_BUFFERED,
                                FILE_ANY_ACCESS)

IOCTL_USB_GET_NODE_CONNECTION_INFORMATION  \
                             = CTL_CODE(FILE_DEVICE_USB,
                                USB_GET_NODE_CONNECTION_INFORMATION,
                                METHOD_BUFFERED,
                                FILE_ANY_ACCESS)

IOCTL_USB_GET_DESCRIPTOR_FROM_NODE_CONNECTION   \
                             = CTL_CODE(FILE_DEVICE_USB,
                                USB_GET_DESCRIPTOR_FROM_NODE_CONNECTION,
                                METHOD_BUFFERED,
                                FILE_ANY_ACCESS)

IOCTL_USB_GET_NODE_CONNECTION_NAME    \
                             = CTL_CODE(FILE_DEVICE_USB,
                                USB_GET_NODE_CONNECTION_NAME,
                                METHOD_BUFFERED,
                                FILE_ANY_ACCESS)

class USB_CONNECTION_STATUS:
    NoDeviceConnected = 0
    DeviceConnected = 1

    # failure codes, these map to fail reasons
    DeviceFailedEnumeration = 2
    DeviceGeneralFailure = 3
    DeviceCausedOvercurrent = 4
    DeviceNotEnoughPower = 5
    DeviceNotEnoughBandwidth = 6
    DeviceHubNestedTooDeeply = 7
    DeviceInLegacyHub = 8
    DeviceEnumerating = 9
    DeviceReset = 10

class UsbSetupPacket(Structure):
    _fields_ = [("request_type", c_ubyte), ("request", c_ubyte),
                ("value", c_ushort), ("index", c_ushort), ("length", c_ushort)]


class Overlapped(Structure):
    _fields_ = [('Internal', LPVOID),
                ('InternalHigh', LPVOID),
                ('Offset', DWORD),
                ('OffsetHigh', DWORD),
                ('Pointer', LPVOID),
                ('hEvent', HANDLE),]


class UsbInterfaceDescriptor(Structure):
    _fields_ = [("b_length", c_ubyte), ("b_descriptor_type", c_ubyte),
                ("b_interface_number", c_ubyte), ("b_alternate_setting", c_ubyte),
                ("b_num_endpoints", c_ubyte), ("b_interface_class", c_ubyte),
                ("b_interface_sub_class", c_ubyte), ("b_interface_protocol", c_ubyte),
                ("i_interface", c_ubyte)]


class PipeInfo(Structure):
    _fields_ = [("pipe_type", c_ulong,), ("pipe_id", c_ubyte),
                ("maximum_packet_size", c_ushort), ("interval", c_ubyte)]


class LpSecurityAttributes(Structure):
    _fields_ = [("n_length", DWORD), ("lp_security_descriptor", c_void_p),
                ("b_Inherit_handle", BOOL)]


class GUID(Structure):
    _fields_ = [("data1", DWORD), ("data2", WORD),
                ("data3", WORD), ("data4", c_byte * 8)]

    def __repr__(self):
        return u'GUID("%s")' % str(self)

    def __str__(self):
        p = c_wchar_p()
        _StringFromCLSID(byref(self), byref(p))
        result = p.value
        _CoTaskMemFree(p)
        return result

    def __cmp__(self, other):
        if isinstance(other, GUID):
            a = bytes(self)
            b = bytes(other)
            return (a > b) - (a < b)
        return -1

    def __nonzero__(self):
        return self != GUID_null

    def __eq__(self, other):
        return isinstance(other, GUID) and \
               bytes(self) == bytes(other)

    def __hash__(self):
        # We make GUID instances hashable, although they are mutable.
        return hash(bytes(self))

GUID_null = GUID()

byte_array_8 = c_byte * 8
# f18a0e88-c30c-11d0-8815-00a0c906bed8
GUID_DEVINTERFACE_USB_HUB = GUID(0xf18a0e88, 0xc30c, 0x11d0, byte_array_8(0x88, 0x15, 0x00, 0xa0, 0xc9, 0x06, 0xbe, 0xd8))

#5e9adaef-f879-473f-b807-4e5ea77d1b1
GUID_DEVINTERFACE_USB_BILLBOARD = GUID(0x5e9adaef, 0xf879, 0x473f, byte_array_8(0xb8, 0x07, 0x4e, 0x5e, 0xa7, 0x7d, 0x1b, 0x1c))

# A5DCBF10-6530-11D2-901F-00C04FB951ED
GUID_DEVINTERFACE_USB_DEVICE = GUID(0xA5DCBF10, 0x6530, 0x11D2, byte_array_8(0x90, 0x1F, 0x00, 0xC0, 0x4F, 0xB9, 0x51, 0xED))

# 3ABF6F2D-71C4-462a-8A92-1E6861E6AF27
GUID_DEVINTERFACE_USB_HOST_CONTROLLER = GUID(0x3abf6f2d, 0x71c4, 0x462a, byte_array_8(0x8a, 0x92, 0x1e, 0x68, 0x61, 0xe6, 0xaf, 0x27))

GUID_CLASS_USBHUB               = GUID_DEVINTERFACE_USB_HUB
GUID_CLASS_USB_DEVICE           = GUID_DEVINTERFACE_USB_DEVICE
GUID_CLASS_USB_HOST_CONTROLLER  = GUID_DEVINTERFACE_USB_HOST_CONTROLLER


# class SpDevinfoData(Structure):
#     _fields_ = [("cb_size", DWORD), ("class_guid", GUID),
#                 ("dev_inst", DWORD), ("reserved", POINTER(c_ulong))]


# class SpDeviceInterfaceData(Structure):
#     _fields_ = [("cb_size", DWORD), ("interface_class_guid", GUID),
#                 ("flags", DWORD), ("reserved", POINTER(c_ulong))]


# class SpDeviceInterfaceDetailData(Structure):
#     _fields_ = [("cb_size", DWORD), ("device_path", WCHAR * 1)]  # devicePath array!!!

class USB_COMMON_DESCRIPTOR(ctypes.Structure):
    _fields_ = [
        ('bLength', UCHAR),
        ('bDescriptorType', UCHAR),
    ]
PUSB_COMMON_DESCRIPTOR = POINTER(USB_COMMON_DESCRIPTOR)


class USB_DESCRIPTOR_REQUEST_Setup(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('bmRequest', UCHAR),
        ('bRequest', UCHAR),
        ('wValue', USHORT),
        ('wIndex', USHORT),
        ('wLength', USHORT),
    ]

class USB_DESCRIPTOR_REQUEST(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('ConnectionIndex', ULONG),
        ('SetupPacket', USB_DESCRIPTOR_REQUEST_Setup),
        ('Data', UCHAR * 1),
    ]
PUSB_DESCRIPTOR_REQUEST = POINTER(USB_DESCRIPTOR_REQUEST)

class USB_HUB_DESCRIPTOR(ctypes.Structure):
    _fields_ = [
        ('bDescriptorLength', CHAR),
        ('bDescriptorType', CHAR),
        ('bNumberOfPorts', CHAR),
        ('wHubCharacteristics', USHORT),
        ('bPowerOnToPowerGood', CHAR),
        ('bHubControlCurrent', CHAR),
        ('bRemoveAndPowerMask', CHAR * 64),
    ]

class USB_MI_PARENT_INFORMATION(ctypes.Structure):
    _fields_ = [
        ('NumberOfInterfaces', ULONG),  # USB_HUB_NODE
    ]

class USB_HUB_INFORMATION(ctypes.Structure):
    _fields_ = [
        ('HubDescriptor', USB_HUB_DESCRIPTOR),  # USB_HUB_NODE
        ('HubIsBusPowered', BOOLEAN),
    ]

class USB_STRING_DESCRIPTOR(ctypes.Structure):
    _fields_ = [
        ('bLength', UCHAR),
        ('bDescriptorType', UCHAR),
        ('bString', WCHAR * 1),
    ]
    def __repr__(self):
        return f"USB_STRING_DESCRIPTOR<{str(self)}>"
    def __str__(self) -> str:
        return wstring_at(byref(self, USB_STRING_DESCRIPTOR.bString.offset), (self.bLength-2)//2)

PUSB_STRING_DESCRIPTOR = POINTER(USB_STRING_DESCRIPTOR)
MAXIMUM_USB_STRING_LENGTH  = 255

class _USB_NODE_INFORMATION_u(Union):
    _fields_ = [
        ('HubInformation', USB_HUB_INFORMATION),
        ('MiParentInformation', USB_MI_PARENT_INFORMATION),
    ]

class USB_NODE_INFORMATION(ctypes.Structure):
    _fields_ = [
        ('NodeType', DWORD),  # USB_HUB_NODE
        ('u', _USB_NODE_INFORMATION_u),
    ]
    def __str__(self) -> str:
        return wstring_at(byref(self, sizeof(ULONG)))


class USB_NODE_CONNECTION_DRIVERKEY_NAME(ctypes.Structure):
    _fields_ = [
        ('ConnectionIndex', ULONG),
        ('ActualLength', ULONG),
        ('DriverKeyName', WCHAR * 1),
    ]
    def __str__(self) -> str:
        return wstring_at(byref(self, sizeof(ULONG) * 2))
PUSB_NODE_CONNECTION_DRIVERKEY_NAME = ctypes.POINTER(USB_NODE_CONNECTION_DRIVERKEY_NAME)

class USB_HCD_DRIVERKEY_NAME(ctypes.Structure):
    _fields_ = [
        ('ActualLength', ULONG),
        ('DriverKeyName', WCHAR * 1),
    ]
    def __str__(self) -> str:
        return wstring_at(byref(self, sizeof(ULONG)))
PUSB_HCD_DRIVERKEY_NAME = ctypes.POINTER(USB_HCD_DRIVERKEY_NAME)

class USB_NODE_CONNECTION_NAME(ctypes.Structure):
    _fields_ = [
        ('ConnectionIndex', ULONG),
        ('ActualLength', ULONG),
        ('NodeName', WCHAR * 1),
    ]
    def __str__(self) -> str:
        return wstring_at(byref(self, USB_NODE_CONNECTION_NAME.NodeName.offset))
PUSB_NODE_CONNECTION_NAME = ctypes.POINTER(USB_NODE_CONNECTION_NAME)

class USB_ROOT_HUB_NAME(ctypes.Structure):
    _fields_ = [
        ('ActualLength', ULONG),
        ('RootHubName', WCHAR * 1),
    ]
    def __str__(self) -> str:
        return wstring_at(byref(self, sizeof(ULONG)))
PUSB_ROOT_HUB_NAME = ctypes.POINTER(USB_ROOT_HUB_NAME)

class USB_CONFIGURATION_DESCRIPTOR(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('bLength', UCHAR),
        ('bDescriptorType', UCHAR),
        ('wTotalLength', USHORT),
        ('bNumInterfaces', UCHAR),
        ('bConfigurationValue', UCHAR),
        ('iConfiguration', UCHAR),
        ('bmAttributes', UCHAR),
        ('MaxPower', UCHAR),
    ]
PUSB_CONFIGURATION_DESCRIPTOR = POINTER(USB_CONFIGURATION_DESCRIPTOR)

class USB_DEVICE_DESCRIPTOR(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('bLength', UCHAR),
        ('bDescriptorType', UCHAR),
        ('bcdUSB', USHORT),
        ('bDeviceClass', UCHAR),
        ('bDeviceSubClass', UCHAR),
        ('bDeviceProtocol', UCHAR),
        ('bMaxPacketSize0', UCHAR),
        ('idVendor', USHORT),
        ('idProduct', USHORT),
        ('bcdDevice', USHORT),
        ('iManufacturer', UCHAR),
        ('iProduct', UCHAR),
        ('iSerialNumber', UCHAR),
        ('bNumConfigurations', UCHAR),
    ]
assert sizeof(USB_DEVICE_DESCRIPTOR) == 18

class USB_INTERFACE_DESCRIPTOR(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('bLength', UCHAR),
        ('bDescriptorType', UCHAR),
        ('bInterfaceNumber', UCHAR),
        ('bAlternateSetting', UCHAR),
        ('bNumEndpoints', UCHAR),
        ('bInterfaceClass', UCHAR),
        ('bInterfaceSubClass', UCHAR),
        ('bInterfaceProtocol', UCHAR),
        ('iInterface', UCHAR),
    ]
assert(sizeof(USB_INTERFACE_DESCRIPTOR) == 9)
PUSB_INTERFACE_DESCRIPTOR = POINTER(USB_INTERFACE_DESCRIPTOR)

class USB_INTERFACE_DESCRIPTOR2(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('bLength', UCHAR),
        ('bDescriptorType', UCHAR),
        ('bInterfaceNumber', UCHAR),
        ('bAlternateSetting', UCHAR),
        ('bNumEndpoints', UCHAR),
        ('bInterfaceClass', UCHAR),
        ('bInterfaceSubClass', UCHAR),
        ('bInterfaceProtocol', UCHAR),
        ('iInterface', UCHAR),
        ('wNumClasses', USHORT),
    ]

class USB_IAD_DESCRIPTOR(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('bLength', UCHAR),
        ('bDescriptorType', UCHAR),
        ('bFirstInterface', UCHAR),
        ('bInterfaceCount', UCHAR),
        ('bFunctionClass', UCHAR),
        ('bFunctionSubClass', UCHAR),
        ('bFunctionProtocol', UCHAR),
        ('iFunction', UCHAR),
    ]

class USB_ENDPOINT_DESCRIPTOR(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('bLength', UCHAR),
        ('bDescriptorType', UCHAR),
        ('bEndpointAddress', UCHAR),
        ('bmAttributes', UCHAR),
        ('wMaxPacketSize', USHORT),
        ('bInterval', UCHAR),
    ]

class USB_PIPE_INFO(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('EndpointDescriptor', USB_ENDPOINT_DESCRIPTOR),
        ('ScheduleOffset', ULONG),
    ]

class USB_NODE_CONNECTION_INFORMATION_EX(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('ConnectionIndex', ULONG),
        ('DeviceDescriptor', USB_DEVICE_DESCRIPTOR),
        ('CurrentConfigurationValue', UCHAR),
        ('Speed', UCHAR),
        ('DeviceIsHub', BOOLEAN),
        ('DeviceAddress', USHORT),
        ('NumberOfOpenPipes', ULONG),
        ('ConnectionStatus', DWORD), # USB_CONNECTION_STATUS
        ('PipeList', USB_PIPE_INFO * 1),
    ]

class USB_NODE_CONNECTION_INFORMATION(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('ConnectionIndex', ULONG),
        ('DeviceDescriptor', USB_DEVICE_DESCRIPTOR),
        ('CurrentConfigurationValue', UCHAR),
        ('LowSpeed', BOOL),
        ('DeviceIsHub', BOOLEAN),
        ('DeviceAddress', USHORT),
        ('NumberOfOpenPipes', ULONG),
        ('ConnectionStatus', DWORD), # USB_CONNECTION_STATUS
        ('PipeList', USB_PIPE_INFO * 1),
    ]

class SP_DEVINFO_DATA(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('cbSize', DWORD),
        ('ClassGuid', GUID),
        ('DevInst', DWORD),
        ('Reserved', ULONG_PTR),
    ]
    def __str__(self):
        return "ClassGuid:%s DevInst:%s" % (self.ClassGuid, self.DevInst)
PSP_DEVINFO_DATA = ctypes.POINTER(SP_DEVINFO_DATA)

class SP_DEVICE_INTERFACE_DATA(ctypes.Structure):
    _fields_ = [
        ('cbSize', DWORD),
        ('InterfaceClassGuid', GUID),
        ('Flags', DWORD),
        ('Reserved', ULONG_PTR),
    ]
    def __str__(self):
        return "InterfaceClassGuid:%s Flags:%s" % (self.InterfaceClassGuid, self.Flags)

PSP_DEVICE_INTERFACE_DATA = ctypes.POINTER(SP_DEVICE_INTERFACE_DATA)

class SP_DEVICE_INTERFACE_DETAIL_DATA(Structure):
    _fields_ = [("cbSize", DWORD), ("DevicePath", WCHAR * 1)]  # devicePath array!!!

    def __repr__(self) -> str:
        return str(self)

    def __str__(self) -> str:
        return wstring_at(byref(self, SP_DEVICE_INTERFACE_DETAIL_DATA.DevicePath.offset))

    # def get_device_path(self):
    #     return wstring_at(byref(self, sizeof(DWORD)))

PSP_DEVICE_INTERFACE_DETAIL_DATA = ctypes.POINTER(SP_DEVICE_INTERFACE_DETAIL_DATA)

# class dummy(ctypes.Structure):
#     _fields_=[("d1", DWORD), ("d2", CHAR)]
#     _pack_ = 1
# SIZEOF_SP_DEVICE_INTERFACE_DETAIL_DATA_A = ctypes.sizeof(dummy)

SetupDiDestroyDeviceInfoList = ctypes.windll.setupapi.SetupDiDestroyDeviceInfoList
SetupDiDestroyDeviceInfoList.argtypes = [HDEVINFO]
SetupDiDestroyDeviceInfoList.restype = BOOL

SetupDiGetClassDevs = ctypes.windll.setupapi.SetupDiGetClassDevsW
SetupDiGetClassDevs.argtypes = [ctypes.POINTER(GUID), c_wchar_p, HANDLE, DWORD]
SetupDiGetClassDevs.restype = HANDLE #ValidHandle # HDEVINFO

SetupDiEnumDeviceInterfaces = ctypes.windll.setupapi.SetupDiEnumDeviceInterfaces
SetupDiEnumDeviceInterfaces.argtypes = [HDEVINFO, PSP_DEVINFO_DATA, ctypes.POINTER(GUID), DWORD, PSP_DEVICE_INTERFACE_DATA]
SetupDiEnumDeviceInterfaces.restype = BOOL

SetupDiGetDeviceInterfaceDetail = ctypes.windll.setupapi.SetupDiGetDeviceInterfaceDetailW
SetupDiGetDeviceInterfaceDetail.argtypes = [HDEVINFO, PSP_DEVICE_INTERFACE_DATA, PSP_DEVICE_INTERFACE_DETAIL_DATA, DWORD, PDWORD, PSP_DEVINFO_DATA]
SetupDiGetDeviceInterfaceDetail.restype = BOOL

SetupDiGetDeviceRegistryProperty = ctypes.windll.setupapi.SetupDiGetDeviceRegistryPropertyW
SetupDiGetDeviceRegistryProperty.argtypes = [HDEVINFO, PSP_DEVINFO_DATA, DWORD, PDWORD, PBYTE, DWORD, PDWORD]
SetupDiGetDeviceRegistryProperty.restype = BOOL

SetupDiGetDeviceInstanceId = ctypes.windll.setupapi.SetupDiGetDeviceInstanceIdW
SetupDiGetDeviceInstanceId.argtypes = [HDEVINFO, PSP_DEVINFO_DATA, PBYTE, DWORD, PDWORD]
SetupDiGetDeviceInstanceId.restype = BOOL

SetupDiEnumDeviceInfo = ctypes.windll.setupapi.SetupDiEnumDeviceInfo
SetupDiEnumDeviceInfo.argtypes = [c_void_p, DWORD, POINTER(SP_DEVINFO_DATA)]
SetupDiEnumDeviceInfo.restype = BOOL

# CMAPI CONFIGRET CM_Get_Parent([out] PDEVINST pdnDevInst, [in]  DEVINST  dnDevInst, [in]  ULONG    ulFlags);
CM_Get_Parent = ctypes.windll.setupapi.CM_Get_Parent
CM_Get_Parent.argtypes = [POINTER(DWORD), DWORD, c_ulong]
CM_Get_Parent.restype = c_ulong

CreateFile = ctypes.windll.kernel32.CreateFileW
CreateFile.argtypes = [
            LPWSTR,                    # _In_          LPCTSTR lpFileName
            DWORD,                     # _In_          DWORD dwDesiredAccess
            DWORD,                     # _In_          DWORD dwShareMode
            LPSECURITY_ATTRIBUTES,     # _In_opt_      LPSECURITY_ATTRIBUTES lpSecurityAttributes
            DWORD,                     # _In_          DWORD dwCreationDisposition
            DWORD,                     # _In_          DWORD dwFlagsAndAttributes
            HANDLE]
CreateFile.restype = HANDLE

CloseHandle = windll.kernel32.CloseHandle



DeviceIoControl = windll.kernel32.DeviceIoControl
DeviceIoControl.argtypes = [
        HANDLE,                    # _In_          HANDLE hDevice
        DWORD,                     # _In_          DWORD dwIoControlCode
        LPVOID,                    # _In_opt_      LPVOID lpInBuffer
        DWORD,                     # _In_          DWORD nInBufferSize
        LPVOID,                    # _Out_opt_     LPVOID lpOutBuffer
        DWORD,                     # _In_          DWORD nOutBufferSize
        LPDWORD,                            # _Out_opt_     LPDWORD lpBytesReturned
        LPOVERLAPPED]                       # _Inout_opt_   LPOVERLAPPED lpOverlapped
DeviceIoControl.restype = BOOL

