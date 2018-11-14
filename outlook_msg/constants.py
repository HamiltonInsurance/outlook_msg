import enum

PROPERTIES_NAME = '__properties_version1.0'

ATTACHMENTS_PREFIX = '__attach_version1.0_#'
SUBSTG_PREFIX = '__substg1.0_'

PROPERTY_IDS = {
    '0x0C1F': 'PidTagSenderEmailAddress',
    '0x0037': 'PidTagSubject',
    '0x1000': 'PidTagBody',
    '0x1013': 'PidTagBodyHtml',
    '0x1009': 'PidTagRtfCompressed',
    '0x0E1B': 'PidTagHasAttachments',
    '0x0E13': 'PidTagMessageAttachments',

    # Attachments
    '0x3701': 'PidTagAttachDataBinary',
    '0x3705': 'PidTagAttachMethod',
    '0x3707': 'PidTagAttachLongFilename'
}

PROPERTY_TYPES = {
    '0x001F': 'PtypeString',  # Null-terminated String in UTF-16LE
    '0x0003': 'PtypInteger32',
    '0x0102': 'PtypBinary',
    '0x000B': 'PtypBoolean',  # 1 or 0
    # 8 bytes; a 64-bit integer representing the number of 100-nanosecond intervals since January 1, 1601
    '0x0040': 'PtypTime',
    '0x0048': 'PtypGuid',  # 16 bytes; a GUID with Data1, Data2, and Data3 fields in little-endian format
    '0x0001': 'PtypNull',  # Null/Placeholder
    # '0x0000': '', # Special, ROP, to be handled specially
}


class HeaderFormat(enum.Enum):
    TOP_LEVEL = enum.auto()
    EMBEDDED_MESSAGE_OBJECT = enum.auto()
    ATTACHMENT_OBJECT = enum.auto()
    RECIPIENT_OBJECT = enum.auto()


class AttachMethod(enum.Enum):
    # 2.2.2.9 PidTagAttachMethod Property
    JustCreated = 0
    ByValue = 1
    ByReference = 2
    Undefined = 3
    ByReferenceOnly = 4
    EmbeddedMessage = 5
    Storage = 6
    WebReference = 7
