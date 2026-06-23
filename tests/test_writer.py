"""Tests for GEDCOM writer output."""

from pathlib import Path

import pytest

from family_tree_converter.reader import (
    Family,
    Individual,
    _code_base,
    _code_role,
    _code_self,
    _parse_approx_string,
    _parse_date,
)
from family_tree_converter.writer import write_gedcom


@pytest.mark.parametrize("code, role, base, self_", [
    # Top-level family head and its spouse / children.
    ("HntJm",            "husband", "HntJm",        "HntJm"),
    ("HntJm-Ca",         "wife",    "HntJm",        "HntJm"),
    ("HntJm/Jn",         "child",   "HntJm",        "HntJm/Jn"),
    # Deep nesting: a child belongs to its *immediate* parent, not the root.
    ("BelAl/Cl/Ar",      "child",   "BelAl/Cl",     "BelAl/Cl/Ar"),
    ("BelAl/Cl/Ar/Dv",   "child",   "BelAl/Cl/Ar",  "BelAl/Cl/Ar/Dv"),
    ("BelAl/Cl/Ar/Dv/T", "child",   "BelAl/Cl/Ar/Dv", "BelAl/Cl/Ar/Dv/T"),
    # An intermediate spouse married into a deep family, not a child of the root.
    ("BelAl/Cl/Ar-Dp",   "wife",    "BelAl/Cl/Ar",  "BelAl/Cl/Ar"),
    ("LivAx/Ar-L",       "wife",    "LivAx/Ar",     "LivAx/Ar"),
    # Multi-marriage chain keeps its base.
    ("GreJeAds-Ada-EmGe","wife",    "GreJeAds",     "GreJeAds"),
])
def test_code_classification(code, role, base, self_):
    assert _code_role(code) == role
    assert _code_base(code) == base
    assert _code_self(code) == self_


@pytest.fixture
def sample_individual():
    return Individual(
        id="I1",
        given_name="John",
        surname="Smith",
        birth_date="1 JAN 1900",
        birth_place="Sydney, NSW, Australia",
        sex="M",
    )


@pytest.fixture
def sample_family(sample_individual):
    wife = Individual(id="I2", given_name="Jane", surname="Smith", sex="F")
    child = Individual(id="I3", given_name="Bob", surname="Smith", sex="M")
    family = Family(
        id="F1",
        husband_id="I1",
        wife_id="I2",
        marriage_date="15 JUN 1925",
        child_ids=["I3"],
    )
    return [sample_individual, wife, child], [family]


@pytest.mark.parametrize("raw, expected", [
    ("1930/1",         "BET 1930 AND 1931"),
    ("1831/2",         "BET 1831 AND 1832"),
    ("1862/3",         "BET 1862 AND 1863"),
    ("approx.1886",    "ABT 1886"),
    ("approx.1893",    "ABT 1893"),
    ("v.approx.1945",  "ABT 1945"),
    ("v.approx.1925",  "ABT 1925"),
    ("1900s",          "BET 1900 AND 1909"),
    ("late 1980s",     "BET 1986 AND 1989"),
    ("late 1860s",     "BET 1866 AND 1869"),
    ("mid.1950s",      "BET 1953 AND 1957"),
    ("mid.1750s",      "BET 1753 AND 1757"),
    ("mid.1780s",      "BET 1783 AND 1787"),
    ("?",              None),
])
def test_parse_approx_string(raw, expected):
    assert _parse_approx_string(raw) == expected


def test_parse_date_excel_serial():
    assert _parse_date(24166.0) == "28 FEB 1966"


def test_parse_date_year_only_not_confused_with_serial():
    assert _parse_date(1760.0) == "1760"


def test_parse_date_clamps_impossible_day():
    # 29 Feb 1978 is not a real date (1978 is not a leap year) -> clamp to 28.
    assert _parse_date(19780229.0) == "28 FEB 1978"
    # A valid leap-year 29 Feb is left untouched.
    assert _parse_date(19800229.0) == "29 FEB 1980"


def test_write_gedcom_produces_valid_structure(sample_family, tmp_path):
    individuals, families = sample_family
    output = tmp_path / "test.ged"
    write_gedcom(individuals, families, output)

    content = output.read_text(encoding="utf-8")
    assert "0 HEAD" in content
    assert "0 TRLR" in content
    assert "0 @I1@ INDI" in content
    assert "1 NAME John /Smith/" in content
    assert "2 DATE 1 JAN 1900" in content
    assert "0 @F1@ FAM" in content
    assert "1 HUSB @I1@" in content
    assert "1 CHIL @I3@" in content
