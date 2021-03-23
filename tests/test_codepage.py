from outlook_msg.codepage import Codepage


def test_codepage_repr():
    assert str(Codepage(20127)) == 'US-ASCII'
    assert str(Codepage(31337)) == 'CP31337'
