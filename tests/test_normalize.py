from src.normalize import normalize_identifier


def test_strips_trailing_excel_float_suffix():
    assert normalize_identifier("10017617.0") == "10017617"


def test_leaves_plain_string_untouched():
    assert normalize_identifier("10017617") == "10017617"


def test_trims_surrounding_whitespace():
    assert normalize_identifier("  10017617  ") == "10017617"


def test_preserves_leading_zeros():
    assert normalize_identifier("0042") == "0042"


def test_preserves_leading_zeros_on_batch_like_values():
    assert normalize_identifier("00123.0") == "00123"


def test_accepts_numeric_input():
    assert normalize_identifier(10017617) == "10017617"
    assert normalize_identifier(10017617.0) == "10017617"


def test_none_and_nan_return_none():
    assert normalize_identifier(None) is None
    assert normalize_identifier(float("nan")) is None


def test_blank_string_returns_none():
    assert normalize_identifier("   ") is None


def test_does_not_mangle_non_numeric_text():
    assert normalize_identifier("2910") == "2910"
    assert normalize_identifier("TOL Mississauga") == "TOL Mississauga"
