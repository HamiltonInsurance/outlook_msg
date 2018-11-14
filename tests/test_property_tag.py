from outlook_msg.property_tag import PropertyTag


def test_from_tuple():
    tuple_form = (31, 0, 55, 0)
    tag1 = PropertyTag.from_bytes(tuple_form)
    tag2 = PropertyTag.from_bytes(bytes(tuple_form))

    assert tag1 == tag2

    assert tag1.prop_id_name == 'PidTagSubject'
    assert tag1.prop_type_name == 'PtypeString'