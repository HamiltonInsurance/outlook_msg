import struct

from outlook_msg import constants


class PropertyTag:
    def __init__(self, prop_id: int, prop_type: int):
        self.prop_id: int = prop_id
        self.prop_type: int = prop_type

    def __eq__(self, o: object) -> bool:
        if not isinstance(o, PropertyTag):
            return False
        return (self.prop_id == o.prop_id) and (self.prop_type == o.prop_type)

    @property
    def full_tag(self):
        return f'0x{self.prop_id:04X}{self.prop_type:04X}'

    @property
    def prop_id_hex(self):
        return f'0x{self.prop_id:04X}'

    @property
    def prop_type_hex(self):
        return f'0x{self.prop_type:04X}'

    @property
    def prop_id_name(self):
        return constants.PROPERTY_IDS.get(self.prop_id_hex, f'Unk: {self.prop_id_hex}')

    @property
    def prop_type_name(self):
        return constants.PROPERTY_TYPES.get(self.prop_type_hex, f'Unk: {self.prop_id_hex}')

    @property
    def substg(self):
        return constants.SUBSTG_PREFIX + f'{self.prop_id:04X}{self.prop_type:04X}'

    def __repr__(self):
        return (f"PropertyTag("
                f"prop_id='{self.prop_id_name}', "
                f"prop_type='{self.prop_type_name}', "
                f"full_tag='{self.full_tag}')")

    @classmethod
    def from_bytes(cls, t):
        """Parse a PropertyTag from a len(4) bytes or bytes-like iterable where each digit is 0-255"""
        # property tag: A 32-bit value that contains a property type and a property ID.
        # The low-order 16 bits represent the property type. The high-order 16 bits represent the property ID.

        b = bytes(t)
        pt_dec, pid_dec = struct.unpack('HH', b)
        pid = pid_dec
        pt = pt_dec
        return PropertyTag(prop_id=pid, prop_type=pt)