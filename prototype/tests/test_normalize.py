from prototype.compare_avito_xml_with_1c_stock import normalize_part, split_part_numbers


def test_normalize_hyphenated_number():
    assert normalize_part("270-5323") == "2705323"


def test_normalize_plain_number():
    assert normalize_part("2705323") == "2705323"


def test_split_removes_parentheses_text():
    assert split_part_numbers("1W5009 (0.25)") == ["1W5009"]


def test_split_comma_separated_numbers():
    assert split_part_numbers("1261373, 1261372") == ["1261373", "1261372"]


def test_split_slash_separated_numbers():
    assert split_part_numbers("4881675/2674631") == ["4881675", "2674631"]


def test_split_rn_suffix_adds_alternative_key():
    assert split_part_numbers("3918776RN") == ["3918776RN", "3918776"]


def test_split_leading_zeroes_adds_alternative_key():
    assert split_part_numbers("0406405520") == ["0406405520", "406405520"]
