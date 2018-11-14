import io

from outlook_msg import constants


class Attachment:
    def __init__(self, mfs):
        self.mfs = mfs

    def open(self):
        if self.attachment_method == constants.AttachMethod.ByValue:
            # The PidTagAttachDataBinary property (section 2.2.2.7) contains the attachment data.
            return io.BytesIO(self.mfs['PidTagAttachDataBinary'])
        else:
            raise NotImplementedError(f"Unable to open attachments stored as: {self.attachment_method.name}")

    @property
    def attachment_method(self):
        return constants.AttachMethod(self.mfs['PidTagAttachMethod'])

    @property
    def filename(self):
        return self.mfs['PidTagAttachLongFilename']