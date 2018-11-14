from outlook_msg.property_tag import PropertyTag


class Property:
    def __init__(self, property_tag: PropertyTag, flags: tuple, value: tuple):
        self.property_tag: PropertyTag = property_tag
        self.flags: tuple = flags
        self.value: tuple = value

    def __eq__(self, o: object) -> bool:
        if not isinstance(o, Property):
            return False

        return (self.property_tag == o.property_tag) and (self.flags == o.flags) and (self.value == o.value)

    @property
    def name(self) -> str:
        return self.property_tag.prop_id_name


def parse_property(p) -> Property:
    property_tag = PropertyTag.from_bytes(p[0:4])
    flags = p[4:8]
    value = p[8:16]
    return Property(property_tag, flags, value)