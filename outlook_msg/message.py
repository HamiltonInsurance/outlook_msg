from outlook_msg import constants
from outlook_msg.message_file_storage import MessageFileStorage
from outlook_msg.attachment import Attachment


class Message:
    def __init__(self, fp):
        self.mfs = MessageFileStorage.from_file(fp)

    @property
    def sender_email(self):
        return self.mfs['PidTagSenderEmailAddress']

    @property
    def subject(self):
        return self.mfs['PidTagSubject']

    @property
    def body(self):
        return self.mfs['PidTagBody']

    @property
    def has_attachments(self):
        try:
            return self.mfs['PidTagHasAttachments']
        except KeyError:
            storage_names = self.mfs.numbered_storage_names(constants.ATTACHMENTS_PREFIX)
            return len(list(storage_names)) > 0

    @property
    def attachments(self):
        if not self.has_attachments:
            return []

        return [Attachment(self.mfs.dir(storage_name))
                for storage_name in self.attachment_storage_names()]

    def attachment_storage_names(self):
        return self.mfs.numbered_storage_names(constants.ATTACHMENTS_PREFIX)