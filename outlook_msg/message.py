from outlook_msg import constants
from outlook_msg.codepage import Codepage
from outlook_msg.message_file_storage import MessageFileStorage
from outlook_msg.attachment import Attachment
from email import message_from_string
import compressed_rtf
import email.policy


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

    def email_message(self, policy=email.policy.default):
        if self.mfs.get('PidTagCodepage'):
            charset = str(Codepage(self.mfs.get('PidTagCodepage')))
        email_message = message_from_string(self.mfs.get('PidTagHeader'), policy=policy)
        email_message.clear_content()
        if self.mfs.get('PidTagBody'):
            email_message.add_alternative(
                self.mfs.get('PidTagBody'),
                charset=charset,
                subtype='plain')
        if self.mfs.get('PidTagBodyHtml'):
            email_message.add_alternative(
                self.mfs.get('PidTagBodyHtml').encode('utf-8'),
                maintype='text',
                subtype='html')
        if self.mfs.get('PidTagRtfCompressed'):
            email_message.add_alternative(
                compressed_rtf.decompress(self.mfs.get('PidTagRtfCompressed')),
                maintype='application',
                subtype='rtf')
        for attachment in self.attachments:
            with attachment.open() as fh:
                email_message.add_attachment(fh.read(),
                    maintype='application',
                    subtype='octet-stream',
                    filename=attachment.filename)
        return email_message
