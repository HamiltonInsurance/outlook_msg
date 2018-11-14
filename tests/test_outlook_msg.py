"""End to end tests of the outlook message logic"""

import pytest

from outlook_msg import Message


@pytest.fixture(scope='session')
def message() -> Message:
    with open('test_data/Test.msg') as f:
        return Message(f)


@pytest.fixture(scope='session')
def message_no_attachment() -> Message:
    with open('test_data/No attachment.msg') as f:
        return Message(f)


def test_message_subject(message):
    assert message.subject == 'Test'


def test_message_body(message):
    assert message.body.strip() == 'This is a test body with a single line.'


def test_message_sender(message):
    assert message.sender_email == "/O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)/CN=RECIPIENTS/CN=17A93D132C634197AFD774F396AFFB26-ELLIOT HUGH"


def test_message_attachments(message):
    assert message.has_attachments
    assert len(message.attachments) == 1
    a = message.attachments[0]
    assert a.filename == 'attachment.txt'
    with a.open() as f:
        contents = f.read().decode()
        assert contents == 'Test Content'


def test_message_no_attachments(message_no_attachment):
    assert not message_no_attachment.has_attachments
    assert len(message_no_attachment.attachments) == 0
