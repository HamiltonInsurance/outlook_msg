import itertools
import struct

import compoundfiles

from outlook_msg import constants
from outlook_msg.property import parse_property


def grouper(iterable, n, fillvalue=None):
    """Collect data into fixed-length chunks or blocks"""
    # grouper('ABCDEFG', 3, 'x') --> ABC DEF Gxx
    args = [iter(iterable)] * n
    return itertools.zip_longest(fillvalue=fillvalue, *args)


def parse_header(s, header_format):
    header_size = get_header_size(header_format)
    header = itertools.islice(s, header_size)
    return list(header)


def get_header_size(header_format):
    if header_format == constants.HeaderFormat.TOP_LEVEL:
        header_size = 32  # 2.4.1.1 Top Level
    elif header_format == constants.HeaderFormat.EMBEDDED_MESSAGE_OBJECT:
        header_size = 24  # 2.4.1.2 Embedded Message object Storage
    elif header_format in (constants.HeaderFormat.ATTACHMENT_OBJECT, constants.HeaderFormat.RECIPIENT_OBJECT):
        header_size = 8  # 2.4.1.3 Attachment Object Storage or Recipient Object Storage
    else:
        raise NotImplementedError(f"Can't handle {header_format}")
    return header_size


def parse_property_stream(ps, header_format):
    s = iter(ps)
    # per https://msdn.microsoft.com/en-us/library/ee178759(v=exchg.80).aspx
    # MUST consist of a header, followed by an array of 16 byte entries

    # Parse the header
    # 2.4.1 The header of the property stream differs depending on which storage this property stream belongs to.
    header = parse_header(s, header_format)
    # 2.4.2 The data inside the property stream MUST be an array of 16-byte entries.
    properties = [parse_property(p) for p in grouper(s, 16)]

    return header, properties


def header_format_from_storage_name(name: str) -> constants.HeaderFormat:
    """Given the name of a storage determine the header format"""
    if name.startswith(constants.ATTACHMENTS_PREFIX):
        return constants.HeaderFormat.ATTACHMENT_OBJECT

    raise NotImplementedError(f"Can't guess the header format for a {name}")


def decode(data, type_name):
    if type_name == 'PtypeString':
        return data.decode('utf-16')
    elif type_name == 'PtypBinary':
        return data  # Just binary so we can return it straight up
    elif type_name == 'PtypBoolean':
        return bool(data[0])
    elif type_name == 'PtypInteger32':
        return struct.unpack('Q', bytes(data))[0]

    raise KeyError(f"Unknown type {type_name}")


class MessageFileStorage:
    """Handles common operations shared between Message, Attachment and Recipient"""

    def __init__(self, root_document, storage, header_format=constants.HeaderFormat.TOP_LEVEL):
        self.storage = storage if storage else root_document.root
        self.document = root_document

        properties_stream = self.read_storage(constants.PROPERTIES_NAME)
        self._header, self._properties = parse_property_stream(properties_stream, header_format)

    def read_storage(self, storage_name):
        with self.document.open(self.storage[storage_name]) as f:
            properties_stream = f.read()
        return properties_stream

    def child(self, storage, header_format):
        return MessageFileStorage(self.document, storage, header_format=header_format)

    @classmethod
    def from_file(cls, fp):
        doc = compoundfiles.CompoundFileReader(fp)
        return cls(doc, doc.root)

    def __getitem__(self, item):
        for p in self._properties:
            type_name = p.property_tag.prop_type_name
            if p.name == item:
                if type_name in ('PtypeString', 'PtypBinary'):
                    binary_data = self.document.open(self.storage[p.property_tag.substg]).read()
                    return decode(binary_data, type_name)
                elif type_name in ('PtypBoolean', 'PtypInteger32'):
                    return decode(p.value, type_name)
                else:
                    raise TypeError(f"Can't decode a {type_name}")
        raise KeyError(item)

    def numbered_storage_names(self, prefix):
        for number in itertools.count():
            proposed_name = prefix + f'{number:08d}'
            if proposed_name in self.storage:
                yield proposed_name
            else:
                break

    def dir(self, name):
        if name in self.storage:
            o = self.storage[name]
            if o.isdir:
                return self.child(o, header_format_from_storage_name(name))