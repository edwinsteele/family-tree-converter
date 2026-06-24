"""Unit tests for date-string normalisation into GEDCOM date phrases."""

import pytest

from family_tree_converter.reader import _date_precision_note, _parse_approx_string


@pytest.mark.parametrize(
    "raw, expected",
    [
        # ISO dates → "D MON YYYY" / "MON YYYY"
        ("1908-09-05", "5 SEP 1908"),
        ("1944-09-22", "22 SEP 1944"),
        ("1908-09", "SEP 1908"),
        # uncertain single year → ABT
        ("1828 ??", "ABT 1828"),
        ("1993 (?)", "ABT 1993"),
        # uncertain decade (ISO and plain) → BET range
        ("194?-02-28", "BET 1940 AND 1949"),
        ("1800s", "BET 1800 AND 1809"),
        ("mid.1730s", "BET 1733 AND 1737"),
        ("late 1790s", "BET 1796 AND 1799"),
        ("early 1760s", "BET 1760 AND 1764"),
        # slash uncertain year
        ("1862/3", "BET 1862 AND 1863"),
        ("1889/90", "BET 1889 AND 1890"),
        # month names (full and abbreviated, with/without day)
        ("April 1888", "APR 1888"),
        ("Feb. 1948", "FEB 1948"),
        ("June 1789", "JUN 1789"),
        # qualifiers
        ("pre 1911", "BEF 1911"),
        ("c. 1920", "ABT 1920"),
        ("approx.1761", "ABT 1761"),
        ("v.approx.1925", "ABT 1925"),
        # month + apostrophe-year, with/without approximate qualifier
        ("Dec'91", "DEC 1991"),
        ("Approx Dec'91", "ABT DEC 1991"),
        ("Approx Dec’91", "ABT DEC 1991"),  # typographic apostrophe
        ("c Jan'05", "ABT JAN 2005"),
        ("Mar'29", "MAR 2029"),  # windowed: '00-'29 → 2000s
        ("Mar'30", "MAR 1930"),  # windowed: '30-'99 → 1900s
        # unknown → dropped
        ("?", None),
    ],
)
def test_parse_approx_string(raw, expected):
    assert _parse_approx_string(raw) == expected


def test_decade_uncertain_date_yields_precision_note():
    # The headline date loses the day/month; a note must preserve it.
    note = _date_precision_note("194?-02-28", "Birth")
    assert note == (
        'Birth date recorded as "194?-02-28": 28 FEB is known, but the year '
        "is uncertain within the 1940s."
    )
    # Month-only variant.
    assert "FEB is known" in _date_precision_note("194?-02", "Death")


@pytest.mark.parametrize("val", ["1908-09-05", "1800s", "1828 ??", 17550420.0, "?"])
def test_no_precision_note_when_nothing_is_lost(val):
    assert _date_precision_note(val, "Birth") is None
