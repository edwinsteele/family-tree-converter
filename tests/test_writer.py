"""Tests for GEDCOM writer output."""

from pathlib import Path

import pytest

from family_tree_converter.reader import Family, Individual
from family_tree_converter.writer import write_gedcom


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
