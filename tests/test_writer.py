"""Tests for GEDCOM writer output."""

import pytest

from family_tree_converter.reader import (
    Family,
    Individual,
    _code_base,
    _code_role,
    _code_self,
    _maiden_name,
    _married_surnames,
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
    # A spouse suffix may follow a digit-ending segment ('P2-J') or be an
    # unknown given name ('L-?', 'An-?'). These must still classify as a married-
    # in spouse, not a child of the grandparent (the old letter-letter regex bug
    # turned John Nolen, the Hynard husband and Annette's husband into children).
    ("PontHe/P2-J",      "wife",    "PontHe/P2",    "PontHe/P2"),
    ("PontHe/L-?",       "wife",    "PontHe/L",     "PontHe/L"),
    ("BelLeAl/An-?",     "wife",    "BelLeAl/An",   "BelLeAl/An"),
    # A prior-marriage chain (two trailing hyphen suffixes) is the linking
    # spouse's *earlier* partner, who heads their own family — not a spouse of
    # the later family.
    ("GreJeAds-Ada-EmGe","ex_spouse", "GreJeAds-Ada-EmGe", "GreJeAds-Ada-EmGe"),
])
def test_code_classification(code, role, base, self_):
    assert _code_role(code) == role
    assert _code_base(code) == base
    assert _code_self(code) == self_


def test_partner_code_for_prior_marriage_chain():
    from family_tree_converter.reader import _partner_code
    # George Emberson's chain points back to the linking spouse Ada he married.
    assert _partner_code("GreJeAds-Ada-EmGe") == "GreJeAds-Ada"


@pytest.mark.parametrize("raw, expected", [
    ("Sydney, N.S.W.", "Sydney, New South Wales"),
    ("Herston, Brisbane, QLD", "Herston, Brisbane, Queensland"),
    ("Creswick, VIC.", "Creswick, Victoria"),
    ("Canberra, A.C.T.", "Canberra, Australian Capital Territory"),
    ("Perth, W.A.", "Perth, Western Australia"),
    ("Flinders Street, Syd.", "Flinders Street, Sydney"),
    ("County Down, IRL.", "County Down, Ireland"),
    ("Nth. Carlton, Melbourne", "North Carlton, Melbourne"),
    ("Maroubra Jnct., Syd.", "Maroubra Junction, Sydney"),
    ("St Peters, Sydney, N.S.W.", "Saint Peters, Sydney, New South Wales"),
    ("Blacktown , Sydney, N.S.W.", "Blacktown, Sydney, New South Wales"),
    # A full street name must NOT be touched, and uncertainty markers survive.
    ("Flinders Street", "Flinders Street"),
    ("Africa (?)", "Africa (?)"),
    (None, None),
])
def test_standardise_place(raw, expected):
    from family_tree_converter.reader import _standardise_place
    assert _standardise_place(raw) == expected


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


@pytest.mark.parametrize("surname, maiden", [
    ("LIVINGSTONE [née Hunter]", "Hunter"),
    ("PONTING then PETTY [née Richey]", "Richey"),
    ("STEELE [née Belshaw/Green]", "Belshaw/Green"),
    ("CLOAK [née  ? ]", None),       # maiden unknown
    ("HUNTER", None),                # no née clause
])
def test_maiden_name(surname, maiden):
    assert _maiden_name(surname) == maiden


@pytest.mark.parametrize("base, parts", [
    ("PONTING then PETTY", ["PONTING", "PETTY"]),
    ("JOB then TAYLOR", ["JOB", "TAYLOR"]),
    ("PONTING", ["PONTING"]),        # single surname → one element
    ("", []),
])
def test_married_surnames(base, parts):
    assert _married_surnames(base) == parts


def test_marriage_from_spouse_row_emitted(tmp_path):
    """A marriage recorded on spouse rows (no 'M'-flag) must reach the GEDCOM."""
    husband = Individual(id="I1", given_name="Daniel", surname="LIVINGSTONE", sex="M")
    wife = Individual(id="I2", given_name="Isobel", surname="LIVINGSTONE", sex="F")
    fam = Family(id="F1", husband_id="I1", wife_id="I2",
                 marriage_date="9 APR 1786", marriage_place="Barony, Scotland")
    out = tmp_path / "t.ged"
    write_gedcom([husband, wife], [fam], out)
    text = out.read_text()
    assert "1 MARR" in text
    assert "2 DATE 9 APR 1786" in text
    assert "2 PLAC Barony, Scotland" in text


def test_lineage_lines_emitted_as_typed_facts(tmp_path):
    """Lineage-chart membership is written as standard FACT/TYPE attributes (one
    per line, sorted) so importers like MacFamilyTree parse them natively —
    not as a custom _GROUP tag and not as a freeform NOTE."""
    ind = Individual(
        id="I1", given_name="Daniel", surname="LIVINGSTONE", sex="M",
        lineage_lines={"Williams", "Livingstone"},
    )
    out = tmp_path / "t.ged"
    write_gedcom([ind], [], out)
    text = out.read_text()
    assert "1 FACT Livingstone\n2 TYPE Lineage" in text
    assert "1 FACT Williams\n2 TYPE Lineage" in text
    # Sorted: Livingstone precedes Williams.
    assert text.index("FACT Livingstone") < text.index("FACT Williams")
    # No custom underscore tag and no freeform "Family lines:" note remain.
    assert "_GROUP" not in text
    assert "Family lines:" not in text


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
    # A single given name yields GIVN + SURN but no SECG (no middle names).
    assert "2 GIVN John" in content
    assert "2 SURN Smith" in content
    assert "2 DATE 1 JAN 1900" in content
    assert "0 @F1@ FAM" in content
    assert "1 HUSB @I1@" in content
    assert "1 CHIL @I3@" in content


def test_name_pieces_are_standard_no_middle_split():
    # GEDCOM 5.5.1 has no separate middle-name field: all given names stay
    # together in GIVN (the full given name), with no proprietary SECG split.
    ind = Individual(id="I1", given_name="Edwin Richard Paul", surname="STEELE")
    from family_tree_converter.writer import render_gedcom
    out = render_gedcom([ind], [])
    assert "1 NAME Edwin Richard Paul /STEELE/" in out
    assert "2 GIVN Edwin Richard Paul" in out
    assert "2 SURN STEELE" in out
    assert "SECG" not in out
