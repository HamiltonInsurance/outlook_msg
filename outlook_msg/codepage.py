import struct

from outlook_msg import constants


class Codepage:
    def __init__(self, id: int):
        self.id: int = id

    def __repr__(self):
        return constants.CODEPAGES.get(self.id, f'CP{self.id}')
