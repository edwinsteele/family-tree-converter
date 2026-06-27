"""Read individuals and relationships from the source Excel spreadsheet."""

from __future__ import annotations

import calendar
import datetime
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import xlrd

# Column indices
_C_GENERATION = 4
_C_CODE = 15
_C_FATHER = 16
_C_MOTHER = 17
_C_SURNAME = 18
_C_GIVEN = 19
_C_DATE1 = 20       # birth date (individual rows) or marriage date (marriage rows)
_C_FLAG = 21        # 'C'=christening, 'M'=marriage row, ''=birth
_C_TOWN = 22
_C_COUNTY = 23
_C_DEATH_DATE = 24
_C_BURIED = 25
_C_LONGEVITY = 29      # recorded age at death (not emitted; cross-check only)
_C_MARRIAGE = 31       # marriage date on an individual (spouse) row
_C_MARRIED_PLACE = 32  # where married, on an individual (spouse) row
_C_OCCUPATION = 33
_C_NOTES = 34
_C_LINE_FIRST = 38     # first of the principal-lineage membership columns
_C_LINE_LAST = 42      # last of the principal-lineage membership columns

DATA_START_ROW = 17


@dataclass(frozen=True)
class FormatProfile:
    """Per-file layout for a genealogist horizontal-tree spreadsheet.

    The reference file (`BlsGrnLivMcCl`) is described by ``BLSGRN_PROFILE`` below;
    the additional trees lay their columns out differently (see scripts/profile.py
    and the per-file maps in project memory), so each gets its own profile and
    ``read_spreadsheet`` is parameterised on it. Column attributes are 0-based
    sheet column indices; set one to ``None`` when the file lacks that column.
    """
    name: str
    data_start_row: int
    generation: int
    code: int
    father: int
    mother: int
    surname: int
    given: int
    date1: int          # birth date (person rows) / marriage date (marriage rows)
    flag: int | None    # 'C'=christening, 'M'=marriage row, 'B'/'D' block rows;
    #                     None when the file has no such flag column
    town: int
    county: int | None  # None when the file uses a single combined place column
    death_date: int
    buried: int
    longevity: int | None
    marriage: int
    married_place: int
    occupation: int
    notes: int
    line_first: int     # first lineage-membership column; >= ncols ⇒ no lineage cols
    # How a data row is recognised as an individual. The reference file numbers
    # every person's generation in col 4, so a non-empty generation marks a
    # person row. Some files leave generation blank and rely on the path code
    # instead — set ``person_row_by_code`` so a non-empty code marks a person.
    person_row_by_code: bool = False
    # Optional column of per-person status markers. In the reference/C&A files
    # this holds Dv=divorced, Df=previously divorced, Tw=twin; in the no-code
    # files it holds Sp=married-in spouse, X=appears as both sibling and married.
    # None when the file has no such column.
    marker: int | None = None
    # A second status-marker column, when the file splits the roles across two
    # columns — J&DJ Steele keeps Sp (married-in) in col 13 but Dv (divorced) in
    # col 11. Holds Dv / Df / Tw; merged with ``marker`` at read time. None when
    # the file carries every status in the single ``marker`` column.
    status_marker: int | None = None
    # Columns whose Excel cell comments are "not for publication" and must be
    # excluded from the harvested notes. Empty when the file has no such column.
    private_note_cols: tuple[int, ...] = ()
    # Secondary "SURNAME"/"Other Names" display columns (cols 26/27 in
    # Stiff:Taylor). Normally an exact copy of the primary surname/given columns,
    # but a handful of rows fill ONLY these (the primary cells left blank) — fall
    # back to them so the person is not left nameless. None when the file has no
    # such pair.
    surname_alt: int | None = None
    given_alt: int | None = None
    # Column of an "Approximate (A)" flag: when it holds 'A' the genealogist has
    # marked a date on the row as uncertain in its year (Stiff:Taylor col 32).
    # A year-only date on such a row is wrapped 'ABT'; a full day/month/year date
    # is left exact. None when the file has no such column.
    approx_flag: int | None = None
    # Structural convention used to derive family relationships:
    #   "alpha" – col-15-style path codes (HntJm / HntJm-Ca / HntJm/Jn)
    #   "none"  – no path code; link by generation + parent names + role markers
    code_convention: str = "alpha"
    # When True, person rows that carry no path code (role "unknown") are linked
    # to their parents by the Father/Mother NAME columns, the way Pass 4 already
    # links coded family heads. Needed by the half-coded files (e.g. Hcks) whose
    # later generations are recorded without codes. Off for the fully-coded
    # reference/C&A files so their output is unchanged.
    name_link_uncoded: bool = False
    # Last data row + 1 (exclusive end). None means read to the sheet end. Set it
    # to stop before a trailing appendix that is not part of the tree — e.g.
    # Brc:Stl ends with a block of speculative "alternative" candidate rows
    # ("BELOW ARE ALTERNATIVES FOR JANE THOMSON", "IS THIS A GOER?") that must
    # not become individuals.
    data_end_row: int | None = None
    # Mid-file row ranges to exclude entirely, as (start, end-exclusive) pairs.
    # Unlike ``data_end_row`` these sit *between* valid regions, so a trailing cut
    # cannot express them — e.g. J&DJ Steele records a block of 13 PRICE entries
    # the genealogist explicitly disavowed ("The following 13 entries are now
    # known to have nothing to do with ... family.") followed by blank rows and a
    # second embedded header, with the real PRICE in-married family beginning
    # afterwards. The range covers the disavowal note, the disavowed entries, the
    # blanks, and the embedded header in one span.
    skip_rows: tuple[tuple[int, int], ...] = ()


# The reference layout, built from the module constants so the two never drift.
BLSGRN_PROFILE = FormatProfile(
    name="BlsGrnLivMcCl",
    data_start_row=DATA_START_ROW,
    generation=_C_GENERATION, code=_C_CODE, father=_C_FATHER, mother=_C_MOTHER,
    surname=_C_SURNAME, given=_C_GIVEN, date1=_C_DATE1, flag=_C_FLAG,
    town=_C_TOWN, county=_C_COUNTY, death_date=_C_DEATH_DATE, buried=_C_BURIED,
    longevity=_C_LONGEVITY, marriage=_C_MARRIAGE, married_place=_C_MARRIED_PLACE,
    occupation=_C_OCCUPATION, notes=_C_NOTES, line_first=_C_LINE_FIRST,
    code_convention="alpha",
)

# "C & A Stl H.Tree #81" — same alpha-code convention as the reference file but a
# different, more compact column layout: the code lives in col 6, the name/event
# block is shifted, birthplace is a single column (no town/county split), there is
# no christening/marriage flag column, and the generation column is left blank
# (so person rows are recognised by their code). Marriage dates are per-person
# (col 18), with no separate 'M' rows. Column map derived from the file's own
# embedded header row (see scripts/profile.py).
CASTL_PROFILE = FormatProfile(
    name="C & A Stl",
    data_start_row=13,
    generation=4, code=6, father=8, mother=9, surname=10, given=11,
    date1=12, flag=None, town=13, county=None, death_date=14, buried=15,
    longevity=16, marriage=18, married_place=19, occupation=20, notes=23,
    line_first=999,  # no lineage-membership columns
    person_row_by_code=True,
    marker=5,  # col 5: Dv (divorced) / Df (prev. divorced) / Tw (twin)
    code_convention="alpha",
)

# "Hcks:Thos:Krsl H.Tr.#120" — same alpha-code convention as the reference file,
# but only HALF its rows are coded: the rest are generation-numbered people who
# name their parents in the Father/Mother columns (linked by name, not code) and
# a tail of no-generation "compact descendant" rows. The column block is shifted
# (code@17 vs @15) with the death/marriage/occupation columns landing back in
# their reference positions. There are no lineage columns; col 16 carries an
# Sp/X marker (spouse / appears-as-both); cols 35-36 are "not for publication".
# Column map derived from the file's own row-16 header legend.
HCKS_PROFILE = FormatProfile(
    name="Hcks:Thos:Krsl",
    data_start_row=17,
    generation=4, code=17, father=18, mother=19, surname=20, given=21,
    date1=22, flag=23, town=24, county=25, death_date=26, buried=27,
    longevity=29, marriage=31, married_place=32, occupation=33, notes=34,
    line_first=999,  # no lineage-membership columns
    marker=16,  # col 16: Sp (married-in spouse) / X (sibling-and-married)
    private_note_cols=(35, 36),  # "Not for publication"
    code_convention="alpha",
    name_link_uncoded=True,  # half the rows are coded; the rest link by name
)

# "Brc:Stl H.Tree #46" — the Steele line's Scottish (Bruce/Steel) ancestry, drawn
# as a descending horizontal tree with NO path codes at all: every person is
# generation-numbered (col 4, counting *down* 11→0 from the oldest ancestor) and
# names their parents in the Father/Mother columns, so the whole file links by
# name (Pass 4b/6) the way the later, uncoded generations of Hcks do. There is a
# real SURNAME column (col 22, with "[née X]"), a B/Sp role marker (col 19;
# B = bloodline, Sp = married-in spouse) and an X "appears as both child and head"
# marker (col 18) — the X-row and the following non-X row are the same person and
# dedup to one individual. Marriage is recorded both on rarer 'M'-flag rows and
# per-person (cols 35/36). col 15 is empty throughout, so it doubles as the
# (unused) code column. The file ends with a speculative "alternatives" appendix
# (rows 181+) excluded via data_end_row. Map derived from the file's own row-20
# header legend.
BRCSTL_PROFILE = FormatProfile(
    name="Brc:Stl",
    data_start_row=21,
    generation=4, code=15, father=20, mother=21, surname=22, given=23,
    date1=24, flag=25, town=26, county=27, death_date=30, buried=31,
    longevity=33, marriage=35, married_place=36, occupation=37, notes=38,
    line_first=999,  # no lineage-membership columns
    marker=19,  # col 19: B (bloodline) / Sp (married-in spouse)
    approx_flag=34,  # 'A' = year uncertain (legend row 19); wrap year-only dates ABT
    code_convention="none",  # no path codes; link by generation + parent names
    name_link_uncoded=True,
    data_end_row=181,  # rows 181+ are a speculative "alternatives" appendix
)

# "J & D J Steele H.Tree #30" — another Steele branch, drawn like Brc:Stl with NO
# path codes: every person in the main tree (rows 14-178) is generation-numbered
# (col 4) and names their parents in the Father/Mother columns, so the file links
# by name (Pass 4b). col 13 carries the Sp marker (married-in spouse); col 19 is
# the B/C/M date-type flag (here only 'B' birth and one 'M' marriage row). col 15
# is empty throughout, so it doubles as the (unused) code column. The file ends
# with a separate, *legitimate* in-married PRICE family block (rows 185-209) under
# its own embedded header (rows 183-184): John Price + Mary Muldoon and their
# 22 children, listed as B-flag event rows under a single 'M' marriage row, with
# the children grouped by position (they name no parents). Dorothy Jane Price in
# that block is the same person as the née-Price spouse at row 15. Map derived
# from the file's own row-13 header legend.
JDJSTEELE_PROFILE = FormatProfile(
    name="J & D J Steele",
    data_start_row=14,
    # col 5 is empty throughout the data, so it serves as the (unused) code
    # column — note col 15 here is the *Mother* column, NOT empty (unlike
    # Brc:Stl, where col 15 was the spare). Father/Mother are cols 14/15.
    generation=4, code=5, father=14, mother=15, surname=16, given=17,
    date1=18, flag=19, town=20, county=21, death_date=24, buried=25,
    longevity=27, marriage=29, married_place=30, occupation=31, notes=32,
    line_first=999,  # no lineage-membership columns
    marker=13,  # col 13: Sp (married-in spouse)
    status_marker=11,  # col 11: Dv (this marriage ended in divorce)
    code_convention="none",  # no path codes; link by generation + parent names
    name_link_uncoded=True,
    # Rows 165-184: the disavowal note, the 13 disavowed PRICE entries the
    # genealogist later found unrelated, the trailing blanks, and the embedded
    # second header — all excluded so only the real PRICE block (185+) is read.
    skip_rows=((165, 185),),
)

# "Stiff:Taylor H.Tree #275" — the most entwined tree (16 people shared with the
# other files), but structurally the cleanest of the no-code files: every one of
# its ~280 rows is generation-numbered (col 4) and names its parents in the
# Father/Mother columns, so the whole file links by name (Pass 4b) like Brc:Stl.
# There are NO path codes, NO 'M'-flag marriage rows (marriage is per-person in
# cols 33/34), no embedded sub-headers and no disavowed appendix. The role
# markers are spread across three columns: col 16 carries Sp (married-in spouse),
# col 15 carries X (a sibling carrying the blood-line who re-appears later as a
# parent entry — handled by dedup, as in Brc) plus 1/2 (1st/2nd husband/wife
# disambiguators), and col 13 carries the single Dv (divorced). col 5 is empty
# throughout, so it doubles as the (unused) code column. cols 36/37 are
# "not for publication". Map derived from the file's own row-16 header legend.
STIFFTAYLOR_PROFILE = FormatProfile(
    name="Stiff:Taylor",
    data_start_row=17,
    generation=4, code=5, father=18, mother=19, surname=20, given=21,
    date1=22, flag=23, town=24, county=25, death_date=28, buried=29,
    longevity=31, marriage=33, married_place=34, occupation=35, notes=17,
    line_first=999,  # cols 38/39 hold bare 1s with no legend → no lineage facts
    marker=16,  # col 16: Sp (married-in spouse)
    status_marker=13,  # col 13: Dv (this marriage ended in divorce)
    private_note_cols=(36, 37),  # "not for publication"
    surname_alt=26, given_alt=27,  # dup name cols; used when primary is blank
    approx_flag=32,  # 'A' = year uncertain; wrap year-only dates as ABT
    code_convention="none",  # no path codes; link by generation + parent names
    name_link_uncoded=True,
)

# Registry of known per-file profiles, keyed by a substring of the file name.
PROFILES: dict[str, FormatProfile] = {
    "BlsGrnLivMcCl": BLSGRN_PROFILE,
    "C & A Stl": CASTL_PROFILE,
    "Hcks:Thos:Krsl": HCKS_PROFILE,
    "Brc:Stl": BRCSTL_PROFILE,
    "J & D J Steele": JDJSTEELE_PROFILE,
    "Stiff:Taylor": STIFFTAYLOR_PROFILE,
}


def profile_for(path: Path) -> FormatProfile:
    """Pick the layout profile for a source file by matching its name."""
    name = Path(path).name
    for key, prof in PROFILES.items():
        if key in name:
            return prof
    return BLSGRN_PROFILE

_MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN",
           "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]

# Month names (full and common abbreviations) → GEDCOM 3-letter code.
_MONTH_NAMES = {
    "january": "JAN", "jan": "JAN", "february": "FEB", "feb": "FEB",
    "march": "MAR", "mar": "MAR", "april": "APR", "apr": "APR",
    "may": "MAY", "june": "JUN", "jun": "JUN", "july": "JUL", "jul": "JUL",
    "august": "AUG", "aug": "AUG", "september": "SEP", "sept": "SEP",
    "sep": "SEP", "october": "OCT", "oct": "OCT", "november": "NOV",
    "nov": "NOV", "december": "DEC", "dec": "DEC",
}


@dataclass
class Individual:
    id: str
    given_name: str
    surname: str
    nickname: str | None = None
    birth_date: str | None = None
    birth_place: str | None = None
    birth_is_christening: bool = False
    death_date: str | None = None
    death_place: str | None = None
    sex: str | None = None
    occupation: str | None = None
    notes: str | None = None
    # Free-text annotations harvested from Excel cell comments.
    note_list: list[str] = field(default_factory=list)
    # Principal lineage charts this person belongs to (e.g. {"Belshaw"}).
    lineage_lines: set[str] = field(default_factory=set)
    # Successive married surnames for a woman recorded as "X then Y".
    married_surnames: list[str] = field(default_factory=list)
    # Family ids this person is linked to as an ADOPTED child (FAMC PEDI adopted).
    adopted_famc: set[str] = field(default_factory=set)


@dataclass
class Family:
    id: str
    husband_id: str | None = None
    wife_id: str | None = None
    marriage_date: str | None = None
    marriage_place: str | None = None
    divorced: bool = False
    child_ids: list[str] = field(default_factory=list)
    # Free-text annotations harvested from Excel cell comments on marriage rows.
    note_list: list[str] = field(default_factory=list)
    # Internal: base code of the husband row, used during construction
    _husband_base: str = field(default="", repr=False, compare=False)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_DECADE_RE = re.compile(
    r"^(?P<qualifier>early|mid|late|v\.approx\.|approx\.)?\s*(?P<year>\d{4})s?$",
    re.IGNORECASE,
)


def _parse_approx_string(s: str) -> str | None:
    """Convert informal approximate date strings to GEDCOM date phrases."""
    if s in ("?", "??", "unknown"):
        return None

    # A "date" with no digit at all is prose, not a date — e.g. the marriage
    # cell of a de-facto couple holds "Joseph & Isabella never married." Drop it
    # so it is never emitted as a GEDCOM DATE; the relationship is preserved via
    # the surname's [De-facto] annotation and the genealogist's cell comments.
    if not re.search(r"\d", s):
        return None

    # Normalise separators so "mid.1950s" and "mid 1950s" both match, and fold
    # a typographic apostrophe to a plain one so "Dec'91" / "Dec’91" both match.
    normalised = s.replace(".", " ").replace("’", "'").strip().lower()

    # "v approx YYYY" → "ABT YYYY"
    m = re.match(r"^v\s+approx\s+(\d{4})$", normalised)
    if m:
        return f"ABT {m.group(1)}"

    # "approx YYYY" → "ABT YYYY"
    m = re.match(r"^approx\s+(\d{4})$", normalised)
    if m:
        return f"ABT {m.group(1)}"

    # Decade patterns: "1900s", "mid 1950s", "late 1980s", "early 1900s"
    m = re.match(r"^(?P<qual>early|mid|late)?\s*(?P<decade>\d{3})0s?$", normalised)
    if m:
        base = int(m.group("decade")) * 10
        qual = m.group("qual")
        if qual == "early":
            return f"BET {base} AND {base + 4}"
        if qual == "mid":
            return f"BET {base + 3} AND {base + 7}"
        if qual == "late":
            return f"BET {base + 6} AND {base + 9}"
        # bare "1900s"
        return f"BET {base} AND {base + 9}"

    # "1931/2/3" → year-first numeric date (Y/M/D), → "3 FEB 1931".
    m = re.match(r"^(\d{4})/(\d{1,2})/(\d{1,2})$", s.strip())
    if m:
        year, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 1 <= month <= 12 and 1 <= day <= 31:
            return f"{day} {_MONTHS[month - 1]} {year}"

    # "1930/1", "1831/2" → uncertain year, convert to BET range
    m = re.match(r"^(\d{4})/\d{1,2}$", s.strip())
    if m:
        year = int(m.group(1))
        return f"BET {year} AND {year + 1}"

    # "1828 ??" / "1993 (?)" → uncertain single year → ABT
    m = re.match(r"^(\d{4})\s*(?:\?+|\(\?\))$", s.strip())
    if m:
        return f"ABT {m.group(1)}"

    # Year range with a two-digit end year, "1832-37" → "BET 1832 AND 1837".
    # Distinguished from an ISO "1908-09" date by the end value being > 12 (a
    # month never is); the end completes the start's century, bumping forward if
    # it would otherwise precede the start.
    m = re.match(r"^(\d{4})-(\d{2})$", s.strip())
    if m and int(m.group(2)) > 12:
        start = int(m.group(1))
        end = (start // 100) * 100 + int(m.group(2))
        if end < start:
            end += 100
        return f"BET {start} AND {end}"

    # ISO "1908-09-05" → "5 SEP 1908"; "1908-09" → "SEP 1908"
    m = re.match(r"^(\d{4})-(\d{2})(?:-(\d{2}))?$", s.strip())
    if m:
        year, month = int(m.group(1)), int(m.group(2))
        if 1 <= month <= 12:
            if m.group(3):
                return f"{int(m.group(3))} {_MONTHS[month - 1]} {year}"
            return f"{_MONTHS[month - 1]} {year}"

    # Uncertain-decade ISO "194?-02-28" → "BET 1940 AND 1949"
    m = re.match(r"^(\d{3})\?-\d{2}(?:-\d{2})?$", s.strip())
    if m:
        base = int(m.group(1)) * 10
        return f"BET {base} AND {base + 9}"

    # "pre 1911" → "BEF 1911"; "post/after 1911" → "AFT 1911"
    m = re.match(r"^(pre|before)\s+(\d{4})$", normalised)
    if m:
        return f"BEF {m.group(2)}"
    m = re.match(r"^(post|after)\s+(\d{4})$", normalised)
    if m:
        return f"AFT {m.group(2)}"

    # "c 1920", "ca 1920", "circa 1920" → "ABT 1920"
    m = re.match(r"^(c|ca|circa)\s+(\d{4})$", normalised)
    if m:
        return f"ABT {m.group(2)}"

    # Month + (optionally apostrophe-prefixed) year, with an optional approximate
    # qualifier: "Dec'91" → "DEC 1991", "Approx Dec'91" → "ABT DEC 1991",
    # "c Jan'05" → "ABT JAN 2005". A two-digit year is windowed at 30
    # ('00-'29 → 2000s, '30-'99 → 1900s).
    m = re.match(
        r"^(?P<qual>approx|circa|ca|c|v\s+approx)?\s*"
        r"(?P<mon>[a-z]+)\s*'?(?P<yy>\d{2}|\d{4})$",
        normalised,
    )
    if m and m.group("mon") in _MONTH_NAMES:
        mon = _MONTH_NAMES[m.group("mon")]
        yy = m.group("yy")
        year = yy if len(yy) == 4 else str((2000 if int(yy) <= 29 else 1900) + int(yy))
        prefix = "ABT " if m.group("qual") else ""
        return f"{prefix}{mon} {year}"

    # Month-name forms: "April 1888", "Feb. 1948", "5 June 1789"
    m = re.match(r"^(?:(\d{1,2})\s+)?([a-z]+)\s+(\d{4})$", normalised)
    if m and m.group(2) in _MONTH_NAMES:
        mon = _MONTH_NAMES[m.group(2)]
        if m.group(1):
            return f"{int(m.group(1))} {mon} {m.group(3)}"
        return f"{mon} {m.group(3)}"

    # Decade *range* "1940s/1950s" → "BET 1940 AND 1959" (spans both decades).
    m = re.match(r"^(\d{3})0s?\s*/\s*(\d{3})0s?$", normalised)
    if m:
        return f"BET {int(m.group(1)) * 10} AND {int(m.group(2)) * 10 + 9}"

    # "early/mid/late YYYY" with a *specific* year (not a decade, handled above):
    # the year is certain, the qualifier only narrows within it, which GEDCOM
    # cannot express — so emit the bare year ("early 1869" → "1869").
    m = re.match(r"^(?:early|mid|late)\s+(\d{4})$", normalised)
    if m:
        return m.group(1)

    # "early/mid/late MONTH YYYY" → "MONTH YYYY": same as above, the month and
    # year are certain and only the intra-month timing (inexpressible in GEDCOM)
    # is qualified — "late Feb 1994" → "FEB 1994".
    m = re.match(r"^(?:early|mid|late)\s+([a-z]+)\s+(\d{4})$", normalised)
    if m and m.group(1) in _MONTH_NAMES:
        return f"{_MONTH_NAMES[m.group(1)]} {m.group(2)}"

    return s  # pass through unchanged; writer will emit as-is


def _parse_date(val: Any) -> str | None:
    """Convert YYYYMMDD float or approximate string to a GEDCOM-friendly date."""
    if val == "" or val is None:
        return None
    if isinstance(val, float):
        date_int = int(val)
        year = date_int // 10000
        remainder = date_int % 10000
        month = remainder // 100
        day = remainder % 100
        if year < 1000:
            if date_int > 2200:
                # Excel date serial (e.g. 24166 → 28 FEB 1966)
                d = datetime.date(1899, 12, 30) + datetime.timedelta(days=date_int)
                return f"{d.day} {_MONTHS[d.month - 1]} {d.year}"
            # Year-only stored as plain integer (e.g. 1760.0)
            return str(date_int)
        if 1 <= month <= 12:
            if day > 0:
                # Clamp impossible days (e.g. 29 FEB 1978, a data error) to
                # the last valid day of the month rather than emitting junk.
                last_day = calendar.monthrange(year, month)[1]
                day = min(day, last_day)
                return f"{day} {_MONTHS[month - 1]} {year}"
            return f"{_MONTHS[month - 1]} {year}"
        return str(year)
    s = str(val).strip()
    if not s:
        return None
    return _parse_approx_string(s)


def _date_precision_note(val: Any, label: str) -> str | None:
    """Return a note preserving precision lost when a date is degraded.

    e.g. "194?-02-28" has a known day and month but a decade-uncertain year,
    so it is emitted as "BET 1940 AND 1949". The headline range hides that the
    event was a 28 Feb; this note records it so nothing is silently lost.
    """
    if not isinstance(val, str):
        return None
    m = re.match(r"^(\d{3})\?-(\d{2})(?:-(\d{2}))?$", val.strip())
    if not m:
        return None
    month = int(m.group(2))
    if not 1 <= month <= 12:
        return None
    known = _MONTHS[month - 1]
    if m.group(3):
        known = f"{int(m.group(3))} {known}"
    decade = int(m.group(1)) * 10
    return (
        f"{label} date recorded as \"{val.strip()}\": {known} is known, but "
        f"the year is uncertain within the {decade}s."
    )


def _longevity_discrepancy_note(
    birth: str | None, death: str | None, longevity: Any
) -> str | None:
    """Flag a contradiction between the recorded age at death (col 29) and the
    age implied by the emitted birth and death dates.

    Surfaces a likely transcription error for other researchers without
    asserting *which* fact is wrong (and without inventing a corrected date).
    Only fires for exact years on both dates and a numeric longevity, and only
    when they disagree by more than a year — so approximate dates and
    infant-death 'Months/Days' longevities never produce noise.
    """
    if not isinstance(longevity, (int, float)) or not birth or not death:
        return None
    if any(q in birth or q in death for q in ("ABT", "BET", "BEF", "AFT", "EST")):
        return None
    bm = re.search(r"\b(\d{4})\b", birth)
    dm = re.search(r"\b(\d{4})\b", death)
    if not (bm and dm):
        return None
    computed = int(dm.group(1)) - int(bm.group(1))
    age = int(longevity)
    if abs(computed - age) <= 1:
        return None
    return (
        f"Recorded age at death ({age}) does not match the {computed} years "
        f"between the recorded birth and death dates — one of these is "
        f"likely a transcription error."
    )


def _impossible_death_note(
    birth: str | None, death: str | None, death_place: str | None
) -> str | None:
    """When the recorded death year precedes the recorded birth year the source
    has an outright contradiction (e.g. Brc:Stl's Mary Ann Costigan, born 1917
    but with a death recorded as 1904). Emitting that as a structured death event
    asserts an impossibility that conformant readers reject, so the caller drops
    the death event and keeps this verbatim note instead — preserving the figure
    without inventing a correction, per the project's preserve-don't-assert rule.
    Returns None unless both dates carry a concrete year and death < birth."""
    if not birth or not death:
        return None
    bm = re.search(r"\b(\d{4})\b", birth)
    dm = re.search(r"\b(\d{4})\b", death)
    if not (bm and dm) or int(dm.group(1)) >= int(bm.group(1)):
        return None
    where = f" at {death_place}" if death_place else ""
    return (
        f'Source records a death date of "{death}"{where}, which precedes the '
        f'recorded birth — retained here as a note rather than a (chronologically '
        f"impossible) death event."
    )


def _is_marriage_label(s: str) -> bool:
    """A "<X> married <Y>" cell is a marriage row, even when its 'M' flag is
    missing (e.g. Brc:Stl's Henry-Steel/Agnes-Anderson row). Used only by the
    no-code files so it can never reclassify a real person row elsewhere."""
    return bool(re.search(r"\bmarried\b", s, re.IGNORECASE))


def _parse_parent_name(raw: str) -> tuple[str | None, str | None] | None:
    """Parse a parent-column value into (given, surname).

    Returns None when the entry is entirely unknown.
    Partial names are returned as (None, surname) or (given, None).

    Examples:
        'James Steele'  → ('James', 'Steele')
        '?  Hunter'     → (None, 'Hunter')
        'Joanna  ?'     → ('Joanna', None)
        '?'             → None
        'Harry(?)'      → ('Harry', None)
    """
    # Strip editorial annotations that are not part of a real ancestor's name:
    # a parenthetical uncertainty marker, and a square-bracket tag such as
    # "[Adopted]" (a child recorded as adopted names no real parent — treating
    # "[Adopted]" as a surname otherwise builds a bogus self-married family).
    raw = re.sub(r"\s*\[.*?\]", "", raw)
    raw = re.sub(r"\s*\(\s*\?\s*\)", "", raw)
    s = " ".join(raw.split())
    if not s or s == "?":
        return None
    parts = s.split()
    if parts[0] == "?":
        rest = " ".join(p for p in parts[1:] if p != "?").strip()
        return (None, rest or None)
    if parts[-1] == "?":
        given = " ".join(p for p in parts[:-1] if p != "?").strip()
        return (given or None, None)
    return (" ".join(parts[:-1]) or None, parts[-1])


# Abbreviation expansions for place names. Modern genealogy apps (e.g.
# MacFamilyTree) geocode and group by place, which works far better on full
# names. Matched only against a whole comma-separated segment (after stripping
# dots/spaces and lower-casing), so they can never substring-corrupt a real
# town name. The source uses these dotted forms consistently.
_PLACE_SEGMENT = {
    "nsw": "New South Wales", "vic": "Victoria", "qld": "Queensland",
    "wa": "Western Australia", "sa": "South Australia",
    "nt": "Northern Territory", "act": "Australian Capital Territory",
    "tas": "Tasmania", "syd": "Sydney", "irl": "Ireland", "eng": "England",
    "scot": "Scotland", "nz": "New Zealand", "png": "Papua New Guinea",
}
# Whole-word expansions applied within a segment.
_PLACE_WORD = {"nth": "North", "sth": "South", "jnct": "Junction"}


def _standardise_place(place: str | None) -> str | None:
    """Expand unambiguous abbreviations and tidy spacing in a place string.

    'Sydney, N.S.W.' → 'Sydney, New South Wales'; 'Nth. Carlton' → 'North
    Carlton'; 'St Peters' → 'Saint Peters'. Street names ('Flinders Street')
    are already spelled out in full in the source, so 'St' is only ever Saint.
    """
    if not place:
        return place
    out: list[str] = []
    for seg in place.split(","):
        seg = seg.strip()
        key = seg.lower().replace(".", "").replace(" ", "")
        if key in _PLACE_SEGMENT:
            out.append(_PLACE_SEGMENT[key])
            continue
        words = seg.split()
        rebuilt = []
        for wi, w in enumerate(words):
            wkey = w.lower().rstrip(".")
            if wi == 0 and wkey == "st":
                rebuilt.append("Saint")
            elif wkey in _PLACE_WORD:
                rebuilt.append(_PLACE_WORD[wkey])
            else:
                rebuilt.append(w)
        out.append(" ".join(rebuilt))
    return ", ".join(s for s in out if s) or None


def _build_place(*parts: str) -> str | None:
    cleaned = [p.strip() for p in parts if str(p).strip() and str(p).strip() != "?"]
    return _standardise_place(", ".join(cleaned)) or None


def _strip_nee(surname: str) -> str:
    """'SMITH [née JONES]' → 'SMITH'; 'STEELE (Maiden Name)' → 'STEELE'.

    Square-bracket née clauses and parenthetical annotations like '(Maiden Name)'
    are editorial, not part of the surname.
    """
    s = re.sub(r"\s*\[.*?\]", "", surname)
    s = re.sub(r"\s*\(.*?\)", "", s)
    return s.strip()


def _surname_annotation(surname: str) -> str | None:
    """Return a square-bracket surname annotation that is NOT a née clause.

    Stiff:Taylor marks a partner in an unmarried union with '[De-facto]' on the
    surname (e.g. 'TAYLOR  [De-facto]'); _strip_nee removes it, so capture it
    here to preserve the relationship status as a NOTE. née clauses are handled
    by _maiden_name and are not returned.
    """
    m = re.search(r"\[([^\]]+)\]", surname)
    if not m:
        return None
    content = m.group(1).strip()
    if re.match(r"n[ée]e\b", content, re.IGNORECASE):
        return None
    return content


def _clean_given(given: str) -> tuple[str, str | None, str | None]:
    """Split editorial annotations and nicknames out of a given name.

    Square-bracket tags like '[Infant death]' or '[MISSIONARY]' are the
    genealogist's notes, not part of the name; strip them and return the
    annotation so it can be preserved as a NOTE.

    A trailing parenthetical that is a *name* is a nickname ('Edith Rosetta
    (Edie or Cissy)', 'James (Jim)') — extract it for a structured GEDCOM NICK.
    A purely numeric or 'No.' parenthetical is a disambiguator ('(No.2)', '(1)')
    and STAYS inline. A bare uncertainty marker '(?)' is stripped.

    Returns (clean_given, annotation_or_None, nickname_or_None).
    """
    annotation = None
    m = re.search(r"\[(.+?)\]", given)
    if m:
        annotation = m.group(1).strip()
        given = re.sub(r"\s*\[.+?\]", "", given).strip()

    nickname = None
    pm = re.search(r"\(([^)]*)\)", given)
    if pm:
        content = pm.group(1).strip()
        if content in ("?", ""):
            # Uncertainty marker — strip it, keep no nickname.
            given = re.sub(r"\s*\([^)]*\)", "", given).strip()
        elif re.fullmatch(r"(no\.?\s*)?\d+", content, re.IGNORECASE):
            pass  # disambiguator — keep inline
        elif re.search(r"[A-Za-z]", content):
            stripped = re.sub(r"\s*\([^)]*\)", "", given).strip()
            # Only treat it as a nickname if a real given name remains; otherwise
            # the parenthetical *is* the name (e.g. "(unnamed child)") — keep it.
            if stripped:
                nickname = content
                given = stripped

    return given, annotation, nickname


def _maiden_name(surname: str) -> str | None:
    """Extract the maiden surname from 'MARRIED [née Maiden]'.

    Returns None when no née clause is present or the maiden name is unknown
    ('[née  ? ]').
    """
    m = re.search(r"\[\s*n[ée]e\s+(.*?)\s*\]", surname, re.IGNORECASE)
    if not m:
        return None
    val = " ".join(m.group(1).split())
    if not val or val == "?":
        return None
    return val


def _married_surnames(surname_base: str) -> list[str]:
    """Split a 'PONTING then PETTY' progression into ['PONTING', 'PETTY'].

    A single surname returns a one-element list.
    """
    parts = [p.strip() for p in re.split(r"\s+then\s+", surname_base, flags=re.IGNORECASE)]
    return [p for p in parts if p]


def _similar_surname(a: str, b: str) -> bool:
    """True when two surnames are equal or one transcription typo apart.

    Used only to reconnect an orphaned married-in spouse to the wife whose
    recorded married surname names him, where the source spelled the surname
    two slightly different ways (e.g. HARTELY vs HARTLEY). Uses Damerau-
    Levenshtein distance ≤ 1 (a single insert/delete/substitute/adjacent
    transposition) and requires a length of at least 5 so short surnames can't
    coincide.
    """
    if a == b:
        return True
    if min(len(a), len(b)) < 5 or abs(len(a) - len(b)) > 1:
        return False
    la, lb = len(a), len(b)
    prev2: list[int] = []
    prev = list(range(lb + 1))
    for i in range(1, la + 1):
        cur = [i] + [0] * lb
        for j in range(1, lb + 1):
            cost = 0 if a[i - 1] == b[j - 1] else 1
            cur[j] = min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost)
            if (i > 1 and j > 1 and a[i - 1] == b[j - 2]
                    and a[i - 2] == b[j - 1]):
                cur[j] = min(cur[j], prev2[j - 2] + 1)
        prev2, prev = prev, cur
    return prev[lb] <= 1


def _infer_sex(surname: str) -> str | None:
    if "née" in surname:
        return "F"
    return None


def _code_role(code: str) -> str:
    """Classify a col-15 code as 'husband', 'wife', 'child', or 'ex_spouse'.

    Codes are slash-separated paths into the tree; the role is decided by the
    *final* path segment so nesting depth doesn't matter:
      'BelAl/Cl/Ar'     → child     (a child of family 'BelAl/Cl')
      'BelAl/Cl/Ar-Dp'  → wife      (spouse who married into family 'BelAl/Cl/Ar')
      'BelAl'           → husband   (a family head)
      'GreJeAds-Ada-EmGe' → ex_spouse (Ada's *prior* husband before Jesse Green)
    A '-suffix' on the last segment marks a married-in spouse and wins over the
    '/' that merely signals depth. A *second* trailing '-suffix' (two hyphens in
    the last segment) marks a prior-marriage chain: the deepest token is the
    earlier spouse of the person named by the rest of the segment. The hyphen
    suffix can follow a digit ('P2-J') or be unknown ('L-?'), so split on the
    hyphen rather than requiring letters on both sides.
    """
    if not code:
        return "unknown"
    last = code.rsplit("/", 1)[-1]
    hyphens = last.count("-")
    if hyphens >= 2:
        return "ex_spouse"
    if hyphens == 1:
        return "wife"
    if "/" in code:
        return "child"
    return "husband"


def _partner_code(code: str) -> str:
    """For a prior-spouse chain, the code of the linking spouse they married.

    'GreJeAds-Ada-EmGe' (George, Ada's prior husband) → 'GreJeAds-Ada' (Ada).
    """
    path, _, last = code.rpartition("/")
    last = last.rsplit("-", 1)[0]
    return f"{path}/{last}" if path else last


def _code_base(code: str) -> str:
    """Return the base code of the family this person *belongs to*.

    A child of 'A/B/C' belongs to the family headed by its immediate parent
    'A/B' (not the top ancestor 'A'). A spouse 'A/B-X' married into family
    'A/B'. Examples:
      'HntJm/Jn'     → 'HntJm'      (child)
      'BelAl/Cl/Ar'  → 'BelAl/Cl'   (child, immediate parent)
      'HntJm-Ca'     → 'HntJm'      (wife)
      'BelAl/Cl/Ar-Dp' → 'BelAl/Cl/Ar' (wife)
    """
    if _code_role(code) == "wife":
        # Strip the spouse suffix from the final path segment.
        path, _, last = code.rpartition("/")
        last_base = last.split("-", 1)[0]
        return f"{path}/{last_base}" if path else last_base
    if "/" in code:
        # Child: belongs to the family headed by its immediate parent.
        return code.rsplit("/", 1)[0]
    return code


def _code_self(code: str) -> str:
    """Return the base code of the family this person *heads*.

    Husbands and children head a family identified by their own code path
    ('BelAl/Cl/Ar' for the child Arthur, whose own children are 'BelAl/Cl/Ar/*').
    Spouses don't head a family of their own; they join their partner's, so
    their self-base is the family they married into.
    """
    if _code_role(code) == "wife":
        return _code_base(code)
    return code


def _dedup_key(given: str, surname_base: str, father_raw: str, birth_raw: Any) -> tuple:
    """Stable key identifying the same person across multiple rows."""
    given_norm = given.strip().lower().split("(")[0].strip()
    surname_norm = surname_base.strip().upper()

    # Father's first name (or '?' if unknown)
    father_first = str(father_raw).strip().split()[0].lower() if str(father_raw).strip() else "?"
    if father_first in ("?", ""):
        father_first = "?"

    # Birth year
    birth_year: int | None = None
    if isinstance(birth_raw, float):
        full = int(birth_raw)
        if full > 9999:
            birth_year = full // 10000
        else:
            birth_year = full
    elif isinstance(birth_raw, str):
        m = re.search(r"\b(\d{4})", birth_raw)
        if m:
            birth_year = int(m.group(1))

    return (surname_norm, given_norm, father_first, birth_year)


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def read_spreadsheet(
    path: Path, profile: FormatProfile = BLSGRN_PROFILE,
    diagnostics: dict | None = None,
) -> tuple[list[Individual], list[Family]]:
    """Parse the spreadsheet into individuals and families.

    When ``diagnostics`` is a dict it is populated (never read) with row-level
    provenance and minted-record markers for the hardening reports in
    ``checks.py``. Collecting them only writes to that dict, so the GEDCOM output
    is byte-for-byte identical whether or not diagnostics are requested.
    """
    wb = xlrd.open_workbook(str(path))
    ws = wb.sheet_by_index(0)

    # Bind the profile's column indices to the historic _C_* / DATA_START_ROW
    # names as locals, so the parsing body below (and its nested closures) reads
    # exactly as before but is now driven by the per-file profile. For the
    # reference profile these equal the module constants, so output is unchanged.
    p = profile
    DATA_START_ROW = p.data_start_row
    DATA_END_ROW = ws.nrows if p.data_end_row is None else p.data_end_row
    # Mid-file rows to exclude entirely (disavowed appendices, embedded headers).
    _excluded_rows = {r for a, b in p.skip_rows for r in range(a, b)}
    _C_GENERATION, _C_CODE, _C_FATHER, _C_MOTHER = (
        p.generation, p.code, p.father, p.mother)
    _C_SURNAME, _C_GIVEN, _C_DATE1, _C_FLAG = (
        p.surname, p.given, p.date1, p.flag)
    _C_TOWN, _C_COUNTY, _C_DEATH_DATE, _C_BURIED = (
        p.town, p.county, p.death_date, p.buried)
    _C_LONGEVITY, _C_MARRIAGE, _C_MARRIED_PLACE = (
        p.longevity, p.marriage, p.married_place)
    _C_OCCUPATION, _C_NOTES, _C_LINE_FIRST = (
        p.occupation, p.notes, p.line_first)
    _C_MARKER = p.marker
    _C_STATUS = p.status_marker
    _C_SURNAME_ALT, _C_GIVEN_ALT = p.surname_alt, p.given_alt
    _C_APPROX = p.approx_flag

    def v(r: int, col: int) -> Any:
        return ws.cell(r, col).value

    def fv(r: int) -> str:
        """Row flag ('C'/'M'/'B'/'D'), or '' when the file has no flag column."""
        return "" if _C_FLAG is None else str(v(r, _C_FLAG)).strip()

    def _birth_place(r: int) -> str | None:
        """Birth/marriage place, honouring files with a single combined place
        column (county is None) instead of separate town and county columns."""
        town = str(v(r, _C_TOWN)) if _C_TOWN is not None else ""
        county = str(v(r, _C_COUNTY)) if _C_COUNTY is not None else ""
        return _build_place(town, county)

    def _longevity(r: int) -> Any:
        return v(r, _C_LONGEVITY) if _C_LONGEVITY is not None else ""

    def mv(r: int) -> str:
        """Per-person role/status marker (Sp/Dv/Df/Tw), or '' when no column."""
        return "" if _C_MARKER is None else str(v(r, _C_MARKER)).strip()

    def smv(r: int) -> str:
        """Secondary status marker (Dv/Df/Tw) when the file splits it into its
        own column, or '' when there is no such column."""
        return "" if _C_STATUS is None else str(v(r, _C_STATUS)).strip()

    def _approx(r: int, date: str | None) -> str | None:
        """Wrap a year-only date as 'ABT' when the row's Approximate (A) flag is
        set; leave full dates and already-qualified dates unchanged."""
        if (date and _C_APPROX is not None
                and str(v(r, _C_APPROX)).strip() == "A"
                and re.fullmatch(r"\d{3,4}", date)):
            return f"ABT {date}"
        return date

    def _is_person_row(r: int, code: str, generation: Any) -> bool:
        """A data row describes an individual. Most files number the generation
        (col 4); some leave it blank and mark people by their path code."""
        return bool(code) if p.person_row_by_code else generation != ""

    # Excel cell comments hold the genealogist's research notes (alternate
    # spellings, birth-vs-christening clarifications, sources). They are not
    # part of any cell's value, so gather them per row, ordered by column.
    row_notes: dict[int, list[str]] = {}
    for (nr, nc), note in sorted(ws.cell_note_map.items()):
        if nc in p.private_note_cols:
            continue  # "not for publication"
        text = (note.text or "").strip()
        if text:
            row_notes.setdefault(nr, []).append(text)

    # The header legend (rows above the data) maps the single-letter lineage
    # codes in cols 38-42 to the principal family-line surnames, e.g.
    # "B=Belshaw", "L=Livingstone". Build that lookup so per-person membership
    # codes can be rendered as readable names.
    line_map: dict[str, str] = {}
    for r in range(0, DATA_START_ROW):
        for c in range(_C_LINE_FIRST, ws.ncols):
            m = re.match(r"^([A-Za-z]{1,2})\s*=\s*(.+)$", str(v(r, c)).strip())
            if m:
                line_map.setdefault(m.group(1), m.group(2).strip())

    def _lines_for(r: int) -> set[str]:
        names: set[str] = set()
        for c in range(_C_LINE_FIRST, ws.ncols):
            code = str(v(r, c)).strip()
            if code:
                names.add(line_map.get(code, code))
        return names

    # Pre-scan every path code so a cross-tree "bridge spouse" can be told apart
    # from an ordinary descendant. A code like 'TmJo/P-HckC/Es' means the spouse
    # of family 'TmJo/P' who is *also* 'HckC/Es' in another subtree (Esther Hicks,
    # who married Pharaoh Thomas and so bridges the Hicks and Thomas lines). Its
    # hyphen sits before the final segment, which the generic role parser would
    # read as a child; here we recognise it because the part before the first
    # hyphen is itself a known family code and the part after matches a known
    # person code. Only enabled for files that use the alpha-code convention.
    all_codes: set[str] = set()
    if p.code_convention == "alpha":
        for r in range(DATA_START_ROW, DATA_END_ROW):
            c = str(v(r, _C_CODE)).strip().replace("|", "/")
            if c:
                all_codes.add(c)

    def _bridge_base(code: str) -> str | None:
        """If ``code`` is a cross-tree bridge spouse, return the base family it
        married into; otherwise None. See the pre-scan comment above."""
        if "-" not in code:
            return None
        a, _, b = code.partition("-")
        if not a or not b or a not in all_codes:
            return None
        if any(c == b or c.startswith(b) or b.startswith(c) for c in all_codes):
            return a
        return None

    def _classify_code(code: str) -> tuple[str, str]:
        """(role, base) for a path code, recognising bridge spouses first."""
        bb = _bridge_base(code)
        if bb is not None:
            return "wife", bb
        return _code_role(code), _code_base(code)

    # ------------------------------------------------------------------
    # Diagnostics – row-coverage accounting and minted-record markers.
    #
    # Each data row is classified into exactly one bucket so a later report can
    # flag any name-bearing row that no pass consumed (silent data loss). The
    # ``consumed_rows`` set is filled by the passes below as they turn rows into
    # output. Everything here only writes to the ``diagnostics`` dict.
    # ------------------------------------------------------------------
    diag = diagnostics

    def _classify_row(r: int) -> str:
        flag = fv(r)
        generation = v(r, _C_GENERATION)
        code = str(v(r, _C_CODE)).strip().replace("|", "/")
        surname = str(v(r, _C_SURNAME)).strip()
        given = str(v(r, _C_GIVEN)).strip()
        if flag == "M":
            return "marriage"
        # A no-code file can carry a generation-numbered marginal annotation with
        # no real name (e.g. Brc:Stl's "{2nd wife - ??}") — layout, not a person.
        if (profile.code_convention == "none" and not _strip_nee(surname)
                and bool(re.fullmatch(r"\{.*\}", given))):
            return "blank/layout"
        if _is_person_row(r, code, generation):
            return "coded-person" if code else "gen-person"
        if flag in ("B", "D") and surname:
            return "block"
        if given and not surname and not code:
            return "loose"
        if surname or given:
            return "no-gen"
        return "blank/layout"

    if diag is not None:
        diag.setdefault("row_class", {})
        diag.setdefault("row_text", {})
        diag.setdefault("consumed_rows", set())
        diag.setdefault("synthetic_ids", set())
        diag.setdefault("placeholder_ids", set())
        diag.setdefault("generation_by_id", {})
        diag.setdefault("name_linked_family_ids", set())
        for r in range(DATA_START_ROW, DATA_END_ROW):
            if r in _excluded_rows:
                continue
            diag["row_class"][r] = _classify_row(r)
            surname = str(v(r, _C_SURNAME)).strip()
            given = str(v(r, _C_GIVEN)).strip()
            if surname or given:
                diag["row_text"][r] = (surname, given)

    # ------------------------------------------------------------------
    # Pass 1 – collect raw person and marriage rows
    # ------------------------------------------------------------------
    person_rows: list[dict] = []
    marriage_rows: list[dict] = []
    # Given-name-only rows (no generation/code/surname/flag) list extra children
    # under the family above them — e.g. Alexander Livingstone's daughters
    # Stella and Maud. Captured with the base of the preceding coded row.
    loose_child_rows: list[tuple[int, str, str | None]] = []
    last_child_base: str | None = None

    for r in range(DATA_START_ROW, DATA_END_ROW):
        if r in _excluded_rows:
            continue
        flag = fv(r)
        generation = v(r, _C_GENERATION)
        code = str(v(r, _C_CODE)).strip().replace("|", "/")
        # No-code files (Brc:Stl) carry a few marriage rows that lack the 'M'
        # flag; recognise them by the "<X> married <Y>" text so they neither
        # become a person nor are mistaken for a loose given-name-only child.
        is_marr = flag == "M" or (
            p.code_convention == "none"
            and _is_marriage_label(str(v(r, _C_GIVEN)))
        )

        if is_marr:
            combined = str(v(r, _C_GIVEN)).strip()
            marr_notes = list(row_notes.get(r, []))
            mn = _date_precision_note(v(r, _C_DATE1), "Marriage")
            if mn:
                marr_notes.append(mn)
            marriage_rows.append({
                "row": r,
                "combined": combined,
                "date": _parse_date(v(r, _C_DATE1)),
                "place": _birth_place(r),
                "cell_notes": marr_notes,
            })
        elif _is_person_row(r, code, generation):
            surname = str(v(r, _C_SURNAME)).strip()
            given_raw = str(v(r, _C_GIVEN)).strip()
            # A few rows fill ONLY the secondary "SURNAME"/"Other Names" display
            # columns, leaving the primary name cells blank (e.g. Stiff:Taylor's
            # Frederick James Giles); fall back to them so the person is named.
            if (not surname and not given_raw
                    and _C_SURNAME_ALT is not None):
                surname = str(v(r, _C_SURNAME_ALT)).strip()
                given_raw = str(v(r, _C_GIVEN_ALT)).strip()
            # A generation-numbered marginal annotation with no real name (e.g.
            # Brc:Stl's "{2nd wife - ??}" placeholder) is not an individual.
            if (profile.code_convention == "none" and not _strip_nee(surname)
                    and re.fullmatch(r"\{.*\}", given_raw)):
                continue
            given, given_annotation, nickname = _clean_given(given_raw)
            father = str(v(r, _C_FATHER)).strip()
            surname_base = _strip_nee(surname)
            marker = mv(r)
            status = smv(r)  # secondary Dv/Df/Tw column, when the file has one
            all_markers = {marker, status}

            person_notes = list(row_notes.get(r, []))
            if given_annotation:
                person_notes.append(f"Name annotation: [{given_annotation}].")
            # A non-née square-bracket surname tag (Stiff:Taylor's '[De-facto]')
            # marks the relationship status and is stripped from the surname;
            # preserve it as a note. Gated to the no-code files so the coded
            # reference/C&A output (whose '(Maiden Name)' tags are paren, not
            # bracket) is unchanged.
            if profile.name_link_uncoded:
                sa = _surname_annotation(surname)
                if sa:
                    person_notes.append(f'Relationship recorded as "{sa}".')
            # A parent column holding only an editorial tag (e.g. "[Adopted]")
            # names no real ancestor; the tag is dropped from parent linking, so
            # preserve its meaning here as a note on the child. Gated to the
            # no-code files so the fully-coded reference/C&A output is unchanged
            # (their adopted children are coded and were never mislinked).
            if profile.name_link_uncoded:
                for praw in (father, str(v(r, _C_MOTHER)).strip()):
                    pm = re.fullmatch(r"\[(.+)\]", praw)
                    note = f'Parentage recorded as "{praw}".'
                    if pm and note not in person_notes:
                        person_notes.append(note)
            # Status markers: twin / previously-divorced become NOTEs on the
            # person; 'Dv' (this marriage divorced) is handled at family level.
            if "Tw" in all_markers:
                person_notes.append("Recorded as a twin.")
            elif "Df" in all_markers:
                person_notes.append("Recorded as previously divorced.")
            for cell, lbl in ((_C_DATE1, "Birth"), (_C_DEATH_DATE, "Death"),
                              (_C_MARRIAGE, "Marriage")):
                pn = _date_precision_note(v(r, cell), lbl)
                if pn:
                    person_notes.append(pn)
            birth_date = _approx(r, _parse_date(v(r, _C_DATE1)))
            death_date = _approx(r, _parse_date(v(r, _C_DEATH_DATE)))
            death_place = _standardise_place(str(v(r, _C_BURIED)).strip()) or None
            ld = _longevity_discrepancy_note(birth_date, death_date, _longevity(r))
            if ld:
                person_notes.append(ld)
            # A death recorded before the birth is a source contradiction; keep
            # the figure as a note and drop the impossible structured event.
            idn = _impossible_death_note(birth_date, death_date, death_place)
            if idn:
                person_notes.append(idn)
                death_date = death_place = None

            role, base = _classify_code(code)
            person_rows.append({
                "row": r,
                "code": code,
                "role": role,
                "base": base,
                "father_raw": father,
                "mother_raw": str(v(r, _C_MOTHER)).strip(),
                "surname": surname,
                "surname_base": surname_base,
                "maiden": _maiden_name(surname),
                "given": given,
                "nickname": nickname,
                "marker": marker,
                "status": status,
                "birth_date": birth_date,
                "birth_is_chr": flag == "C",
                "birth_place": _birth_place(r),
                "death_date": death_date,
                "death_place": death_place,
                # Marriage date/place are recorded on each spouse's own row
                # (cols 31/32), not only on the rarer 'M'-flag rows.
                "marr_date": _approx(r, _parse_date(v(r, _C_MARRIAGE))),
                "marr_place": _standardise_place(str(v(r, _C_MARRIED_PLACE)).strip()) or None,
                "sex": _infer_sex(surname),
                "occupation": str(v(r, _C_OCCUPATION)).strip() or None,
                "notes": str(v(r, _C_NOTES)).strip() or None,
                "cell_notes": person_notes,
                "lines": _lines_for(r),
                "dedup_key": _dedup_key(given, surname_base, father, v(r, _C_DATE1)),
            })
            last_child_base = _code_base(code)
        else:
            given = str(v(r, _C_GIVEN)).strip()
            if (
                given
                and not str(v(r, _C_SURNAME)).strip()
                and not str(v(r, _C_CODE)).strip()
            ):
                loose_child_rows.append((r, given, last_child_base))

    if diag is not None:
        # Every person row becomes (or merges into) an individual via Pass 2,
        # and every marriage row is offered to Pass 3 — all accounted for.
        diag["consumed_rows"].update(pr["row"] for pr in person_rows)
        diag["consumed_rows"].update(m["row"] for m in marriage_rows)

    # ------------------------------------------------------------------
    # Pass 2 – deduplicate individuals
    # ------------------------------------------------------------------
    dedup_map: dict[tuple, Individual] = {}
    _id_counter = 0

    def _get_or_create(p: dict) -> Individual:
        nonlocal _id_counter
        dk = p["dedup_key"]
        if dk in dedup_map:
            existing = dedup_map[dk]
            # Merge: fill gaps with data from the second appearance
            if not existing.birth_date and p["birth_date"]:
                existing.birth_date = p["birth_date"]
            if not existing.birth_place and p["birth_place"]:
                existing.birth_place = p["birth_place"]
            if not existing.death_date and p["death_date"]:
                existing.death_date = p["death_date"]
            if not existing.death_place and p["death_place"]:
                existing.death_place = p["death_place"]
            if not existing.occupation and p["occupation"]:
                existing.occupation = p["occupation"]
            if existing.sex is None and p["sex"]:
                existing.sex = p["sex"]
            if not existing.nickname and p.get("nickname"):
                existing.nickname = p["nickname"]
            for n in p["cell_notes"]:
                if n not in existing.note_list:
                    existing.note_list.append(n)
            existing.lineage_lines |= p["lines"]
            return existing
        _id_counter += 1
        # A woman recorded as "PONTING then PETTY [née Richey]" carries two
        # successive married surnames in one cell. Genealogically she is filed
        # under her maiden (birth) surname — which also matches her position in
        # her own family's chart — with the married names preserved separately.
        married = _married_surnames(p["surname_base"])
        if len(married) > 1 and p["maiden"]:
            display_surname = p["maiden"]
            married_surnames = married
        else:
            display_surname = p["surname_base"]
            married_surnames = []
        ind = Individual(
            id=f"I{_id_counter}",
            given_name=p["given"],
            surname=display_surname,
            nickname=p.get("nickname"),
            birth_date=p["birth_date"],
            birth_place=p["birth_place"],
            birth_is_christening=p["birth_is_chr"],
            death_date=p["death_date"],
            death_place=p["death_place"],
            sex=p["sex"],
            occupation=p["occupation"],
            notes=p["notes"],
            note_list=list(p["cell_notes"]),
            lineage_lines=set(p["lines"]),
            married_surnames=married_surnames,
        )
        dedup_map[dk] = ind
        return ind

    for p in person_rows:
        _get_or_create(p)

    # ------------------------------------------------------------------
    # Pass 2b – reconcile the same person recorded in multiple lineage
    # charts. Someone who bridges several family lines (e.g. appears as a
    # child in two different ancestral charts) gets distinct dedup keys
    # because their parent names differ between charts. When the name AND
    # an exact birth date (day + month + year) match, they are certainly
    # the same individual: merge them onto one record and repoint the
    # dedup keys so every chart contributes a parent family. This only
    # ever merges — a shared birth *year* alone is not enough, so same-
    # named cousins are never collapsed.
    # ------------------------------------------------------------------
    def _has_full_date(d: str | None) -> bool:
        return bool(d) and any(mon in d for mon in _MONTHS)

    def _merge_into(dst: Individual, src: Individual) -> None:
        if not dst.birth_place and src.birth_place:
            dst.birth_place = src.birth_place
        if not dst.death_date and src.death_date:
            dst.death_date = src.death_date
        if not dst.death_place and src.death_place:
            dst.death_place = src.death_place
        if not dst.occupation and src.occupation:
            dst.occupation = src.occupation
        if dst.sex is None and src.sex:
            dst.sex = src.sex
        if not dst.notes and src.notes:
            dst.notes = src.notes
        for n in src.note_list:
            if n not in dst.note_list:
                dst.note_list.append(n)

    strong_canon: dict[tuple, Individual] = {}
    for dk, ind in list(dedup_map.items()):
        if not _has_full_date(ind.birth_date):
            continue
        given_norm = (ind.given_name or "").lower().split("(")[0].strip()
        sk = ((ind.surname or "").upper(), given_norm, ind.birth_date)
        canon = strong_canon.get(sk)
        if canon is None:
            strong_canon[sk] = ind
        elif canon is not ind:
            _merge_into(canon, ind)
            dedup_map[dk] = canon

    # ------------------------------------------------------------------
    # Pass 3 – build families using code-based role classification
    #
    # The col-15 code encodes family structure:
    #   'HntJm'    → husband (family head)
    #   'HntJm-Ca' → wife of HntJm family
    #   'HntJm/Jn' → child of HntJm family
    #
    # We scan rows in spreadsheet order. A new 'husband' code always
    # starts a new family. Wives and children are matched by base code
    # to the most recently opened family with that base.
    # ------------------------------------------------------------------
    families: list[Family] = []
    _fam_counter = 0
    # Maps base code → the most recently opened Family with that base
    active_family: dict[str, Family] = {}

    # Maps a code to the individual who heads the family of that code, so an
    # intermediate child (e.g. 'BelAl/Cl/Ar') can anchor a family that its own
    # children ('BelAl/Cl/Ar/Dv') and spouse ('BelAl/Cl/Ar-Dp') attach to.
    code_to_head: dict[str, Individual] = {}
    for p in person_rows:
        if p["role"] in ("husband", "child"):
            code_to_head.setdefault(_code_self(p["code"]), dedup_map[p["dedup_key"]])

    # Literal code → individual, so a prior-spouse chain ('GreJeAds-Ada-EmGe')
    # can find the linking spouse it married ('GreJeAds-Ada').
    full_code_to_ind: dict[str, Individual] = {}
    for p in person_rows:
        full_code_to_ind.setdefault(p["code"], dedup_map[p["dedup_key"]])

    def _open_family(base: str, husband_id: str | None = None) -> Family:
        nonlocal _fam_counter
        if husband_id is None:
            head = code_to_head.get(base)
            husband_id = head.id if head else None
        _fam_counter += 1
        fam = Family(id=f"F{_fam_counter}", husband_id=husband_id, _husband_base=base)
        families.append(fam)
        active_family[base] = fam
        return fam

    def _find_head_code(base: str) -> str | None:
        """Map a family base to a known head code, tolerating abbreviation.

        Codes are sometimes abbreviated inconsistently — Barbara's own code is
        'BelAl/Rg/Do/Ba' but her spouses reference 'BelAl/Rg/Do/B'. Match within
        the *same parent path* and accept a final segment that is a prefix of
        (or prefixed by) the head's, preferring the longest (most specific) one.
        """
        if base in code_to_head:
            return base
        bpath, _, blast = base.rpartition("/")
        best: str | None = None
        for hc in code_to_head:
            hpath, _, hlast = hc.rpartition("/")
            if hpath != bpath:
                continue
            if hlast.startswith(blast) or blast.startswith(hlast):
                if best is None or len(hc) > len(best):
                    best = hc
        return best

    def _family_for(base: str, create: bool = True) -> Family | None:
        """Resolve (and optionally lazily anchor) the family for a base code."""
        fam = active_family.get(base)
        if fam is not None:
            return fam
        hc = _find_head_code(base)
        if hc is None:
            return None
        if hc in active_family:
            active_family[base] = active_family[hc]  # alias the abbreviation
            return active_family[hc]
        if not create:
            return None
        head = code_to_head.get(hc)
        fam = _open_family(hc, husband_id=head.id if head else None)
        if hc != base:
            active_family[base] = fam  # alias so sibling codes resolve here too
        return fam

    # Spouse-id → Individual, for matching marriage rows to families by name.
    id_to_ind = {ind.id: ind for ind in dedup_map.values()}

    def _marriage_tokens(combined: str) -> set[str]:
        return set(re.findall(r"[a-z]+", combined.lower()))

    # Interleave person and marriage rows in row order
    all_events: list[tuple[int, str, dict]] = (
        [(p["row"], "person", p) for p in person_rows]
        + [(m["row"], "marriage", m) for m in marriage_rows]
    )
    all_events.sort(key=lambda x: x[0])

    for _, etype, data in all_events:
        if etype == "marriage":
            # Attach marriage info to the most recently opened family whose
            # husband or wife surname appears in the combined "X m. Y" string.
            # The surname guard prevents a stray marriage row (e.g. from a
            # note block whose people were never read as individuals) from
            # corrupting an unrelated family that merely happens to lack a
            # marriage date.
            tokens = _marriage_tokens(data["combined"])
            for fam in reversed(families):
                if fam.marriage_date is not None:
                    continue
                h = id_to_ind.get(fam.husband_id)
                w = id_to_ind.get(fam.wife_id)
                h_sn = (h.surname or "").lower().strip() if h else ""
                w_sn = (w.surname or "").lower().strip() if w else ""
                if (h_sn and h_sn in tokens) or (w_sn and w_sn in tokens):
                    fam.marriage_date = data["date"]
                    fam.marriage_place = data["place"]
                    fam.note_list.extend(data["cell_notes"])
                    data["attached"] = True
                    break
            continue

        # Person row
        ind = dedup_map[data["dedup_key"]]
        role = data["role"]
        base = data["base"]

        if role == "husband":
            existing = active_family.get(base)
            if existing is not None and existing.husband_id in (None, ind.id):
                # Family already opened — either a spouse appeared first
                # (female-headed entry, husband_id still None) or it was
                # lazily anchored with this same person as head. Reuse it
                # rather than creating a duplicate.
                existing.husband_id = ind.id
            else:
                _open_family(base, husband_id=ind.id)

        elif role == "wife":
            # The family she married into may not have been opened yet — an
            # intermediate head (e.g. child 'BelAl/Cl/Ar') only anchors a family
            # lazily, when its first spouse or child appears. Fall back to a
            # headless family only if no head can be resolved at all.
            fam = _family_for(base) or _open_family(base)
            if fam.wife_id is not None and fam.wife_id != ind.id:
                # A second spouse on the same base is a remarriage — open a
                # parallel family sharing the same head so it isn't dropped.
                fam = _open_family(base, husband_id=fam.husband_id)
            if fam.wife_id is None:
                fam.wife_id = ind.id

        elif role == "child":
            # Anchor the immediate-parent family lazily; abbreviated child codes
            # (e.g. 'PettL' vs family 'PettLuGe') resolve via _find_head_code.
            fam = _family_for(base)
            if fam is not None:
                fam.child_ids.append(ind.id)
            # If no matching family, the child is an unresolved reference — skip.

        elif role == "ex_spouse":
            # A prior-marriage chain: this person married the linking spouse
            # ('GreJeAds-Ada') in an earlier marriage, so they head their own
            # family with that spouse — NOT a member of the later family. Place
            # the partner opposite by sex (a née-female partner ⇒ this is the
            # husband) and let the final sweep fill the remaining sex.
            partner = full_code_to_ind.get(_partner_code(data["code"]))
            if partner is not None:
                if partner.sex == "M":
                    h_id, w_id = ind.id, partner.id
                else:
                    h_id, w_id = partner.id, ind.id
                _fam_counter += 1
                families.append(Family(id=f"F{_fam_counter}",
                                       husband_id=h_id, wife_id=w_id))

    # ------------------------------------------------------------------
    # Post-process: swap husband/wife when the assigned "husband" is
    # demonstrably female (has née in surname). This corrects for families
    # where the female descendant is the tree's connecting person and
    # therefore gets the base code rather than the -suffix code.
    # ------------------------------------------------------------------
    id_to_ind = {ind.id: ind for ind in dedup_map.values()}
    for fam in families:
        if fam.husband_id and fam.wife_id:
            h = id_to_ind[fam.husband_id]
            if h.sex == "F":
                fam.husband_id, fam.wife_id = fam.wife_id, fam.husband_id

    # A 'Dv' status marker sits on a divorced spouse — its family is marked ended
    # in divorce (emitted as 1 DIV) in a final sweep below, once the name-linked
    # families (Pass 4b/6) that some of these spouses head have all been built.
    divorced_spouse_ids = {
        dedup_map[p["dedup_key"]].id
        for p in person_rows if "Dv" in (p.get("marker"), p.get("status"))
    }

    # Attach marriage date/place recorded on the spouses' own rows (cols 31/32).
    # The 'M'-flag marriage rows cover only a fraction of couples; the rest
    # carry their marriage details on each individual's row. Index by spouse id
    # and fill any family still lacking a marriage. Keying on the husband's row
    # keeps remarriages distinct (each husband row belongs to one family); the
    # wife's row is a fallback when the husband has no recorded date.
    row_marr: dict[str, dict] = {}
    for p in person_rows:
        if p["role"] not in ("husband", "wife", "ex_spouse"):
            continue
        if not (p["marr_date"] or p["marr_place"]):
            continue
        row_marr.setdefault(dedup_map[p["dedup_key"]].id, p)
    # A person who married more than once carries a single col-31 value that
    # describes only one of their marriages, so it must not be copied onto every
    # family they belong to (e.g. Henry Ponting's 1884 date leaking onto his
    # first wife Maud's family). Prefer the spouse who appears in the *fewest*
    # families — their col-31 unambiguously refers to this marriage — and break
    # ties husband-first.
    spouse_fam_count: dict[str, int] = {}
    for fam in families:
        for sid in (fam.husband_id, fam.wife_id):
            if sid:
                spouse_fam_count[sid] = spouse_fam_count.get(sid, 0) + 1
    for fam in families:
        if fam.marriage_date or fam.marriage_place:
            continue
        cands = [s for s in (fam.husband_id, fam.wife_id) if s and s in row_marr]
        cands.sort(key=lambda s: (spouse_fam_count.get(s, 0),
                                  0 if s == fam.husband_id else 1))
        if cands:
            src = row_marr[cands[0]]
            fam.marriage_date = src["marr_date"]
            fam.marriage_place = src["marr_place"]

    # ------------------------------------------------------------------
    # Pass 3.5 – loose given-name-only children
    #
    # Attach each bare given-name row to the family it sits under, taking the
    # father's surname (their maiden/paternal name). Skip names that merely
    # repeat an already-coded child (e.g. "Mabel" == "Mabel Annie", LivAx/Mb)
    # to avoid duplicating that individual.
    # ------------------------------------------------------------------
    for row, given, ctx_base in loose_child_rows:
        if not ctx_base:
            continue
        fam = active_family.get(ctx_base)
        if fam is None:
            continue
        # The row resolves to a family — whether it is added as a new child or
        # skipped as a duplicate of one already present, its data is accounted.
        if diag is not None:
            diag["consumed_rows"].add(row)
        existing_givens = {
            (id_to_ind[c].given_name or "").lower().split()[0]
            for c in fam.child_ids
            if c in id_to_ind and id_to_ind[c].given_name
        }
        if given.lower().split()[0] in existing_givens:
            continue
        father = id_to_ind.get(fam.husband_id)
        surname = (father.surname if father else "") or ""
        _id_counter += 1
        child = Individual(id=f"I{_id_counter}", given_name=given, surname=surname)
        dedup_map[("__loose__", row)] = child
        id_to_ind[child.id] = child
        fam.child_ids.append(child.id)

    # ------------------------------------------------------------------
    # Pass 4 – Synthetic parents for family-head rows
    #
    # Husband and wife rows (no '/' in code) carry their parents' names
    # in columns Q/R but are never assigned FAMC links by the code-based
    # Pass 3. Here we parse those names, match to existing individuals
    # where possible, create synthetic Individual records where not, and
    # wire up parent Family records accordingly.
    # ------------------------------------------------------------------
    def _ind_name_key(given: str | None, surname: str | None) -> tuple[str, str]:
        g_str = (given or "").lower().split("(")[0].strip()
        g = g_str.split()[0] if g_str else ""
        return ((surname or "").upper(), g)

    # Name-keyed lookup across all real individuals (first match wins)
    existing_by_name: dict[tuple[str, str], Individual] = {}
    for ind in dedup_map.values():
        k = _ind_name_key(ind.given_name or None, ind.surname or None)
        if k != ("", ""):
            existing_by_name.setdefault(k, ind)

    synth_by_name: dict[tuple, Individual] = {}
    # Synthetics minted from a GIVEN-NAME-ONLY parent reference. Such a name is
    # too weak an identity to share across families, so we track these to stop a
    # later given-only reference collapsing onto one belonging to another couple.
    surnameless_synth: set[str] = set()
    parent_fams: dict[tuple[str | None, str | None], Family] = {}
    already_child: set[str] = {cid for fam in families for cid in fam.child_ids}

    def _descendants(root_id: str | None) -> set[str]:
        """Ids reachable as descendants of root_id through current families."""
        if root_id is None:
            return set()
        seen: set[str] = set()
        stack = [root_id]
        while stack:
            cur = stack.pop()
            for fam in families:
                if cur in (fam.husband_id, fam.wife_id):
                    for c in fam.child_ids:
                        if c not in seen:
                            seen.add(c)
                            stack.append(c)
        return seen

    def _year_of(date: str | None) -> int | None:
        m = re.search(r"\b(\d{4})\b", date or "")
        return int(m.group(1)) if m else None

    def _resolve_parent(
        parsed: tuple[str | None, str | None] | None,
        sex: str,
        avoid: frozenset[str] = frozenset(),
        child_birth: str | None = None,
        co_parent: Individual | None = None,
    ) -> Individual | None:
        nonlocal _id_counter
        if parsed is None or (parsed[0] is None and parsed[1] is None):
            return None
        key = _ind_name_key(parsed[0], parsed[1])
        has_surname = bool(parsed[1])
        # A parent named by GIVEN NAME ONLY (no surname, e.g. a mother recorded
        # just as "Alice") is a weak identity. It may legitimately reuse a real
        # or block person, or a sibling's already-minted parent (scoped to the
        # same co-parent), but it must NEVER collapse onto another family's
        # given-only synthetic — that would merge two unrelated women who happen
        # to share a first name (e.g. Edward Bryan's "Alice" vs John Burrowes's
        # "Alice"). So for a given-only reference, first try the co-parent-scoped
        # key, then fall back to the plain key but reject a given-only synthetic.
        # Scoped only for the name-linked files: the reference file's output
        # depends on the historic plain-key sharing (incl. its block aliases),
        # which has no given-only over-merge in practice, so it is left as-is.
        gate = profile.name_link_uncoded
        parent = None
        if gate and not has_surname and co_parent is not None:
            parent = synth_by_name.get(("", co_parent.id, key[1]))
        if parent is None:
            cand = existing_by_name.get(key) or synth_by_name.get(key)
            if (gate and cand is not None and not has_surname
                    and cand.id in surnameless_synth):
                cand = None
            parent = cand
        # A loose name match (surname + first given word) can collide with a
        # descendant of the very person we are giving parents to — e.g. Bruce
        # Dallas's father "William H. Dallas" matching his grandson "William
        # John Peter Dallas". Such a match would invert the tree, so reject it
        # and mint a distinct ancestor instead.
        if parent is not None and parent.id in avoid:
            parent = None
        # A widened name match (nickname / middle name / maiden surname) can also
        # land on a same-named person of the wrong generation — e.g. "Dorothy
        # Julian", mother of a child born 1915, matching Dorothy Williams née
        # Julian (born 1925). A parent cannot be born on or after their child,
        # so reject such a match and mint a distinct ancestor instead.
        if parent is not None and child_birth is not None:
            py, cy = _year_of(parent.birth_date), _year_of(child_birth)
            if py is not None and cy is not None and py >= cy:
                parent = None
        if parent is None:
            _id_counter += 1
            parent = Individual(
                id=f"I{_id_counter}",
                given_name=parsed[0] or "",
                surname=parsed[1] or "",
                sex=sex,
            )
            if has_surname or not gate:
                synth_by_name[key] = parent
            else:
                # Register only under the co-parent-scoped key so siblings reuse
                # this parent while other couples don't; track it as given-only.
                surnameless_synth.add(parent.id)
                if co_parent is not None:
                    synth_by_name[("", co_parent.id, key[1])] = parent
                else:
                    # No co-parent (the other parent is unknown), so there is no
                    # shared key to register under — but the record is still
                    # referenced by the child's family and must reach `ordered`
                    # or it becomes a dangling pointer. Register under a unique,
                    # never-matched key so it is emitted without collapsing onto
                    # another given-only parent.
                    synth_by_name[("__unlinked__", parent.id)] = parent
        return parent

    for _, etype, data in all_events:
        if etype != "person" or data["role"] not in ("husband", "wife"):
            continue
        ind = dedup_map[data["dedup_key"]]
        if ind.id in already_child:
            continue

        fp = _parse_parent_name(data["father_raw"])
        mp = _parse_parent_name(data["mother_raw"])
        if fp is None and mp is None:
            continue

        avoid = frozenset(_descendants(ind.id) | {ind.id})
        father_ind = _resolve_parent(fp, "M", avoid)
        mother_ind = _resolve_parent(mp, "F", avoid, co_parent=father_ind)

        if father_ind is None and mother_ind is None:
            continue

        pkey = (
            father_ind.id if father_ind else None,
            mother_ind.id if mother_ind else None,
        )
        if pkey not in parent_fams:
            _fam_counter += 1
            pf = Family(id=f"F{_fam_counter}", husband_id=pkey[0], wife_id=pkey[1])
            families.append(pf)
            parent_fams[pkey] = pf
        parent_fams[pkey].child_ids.append(ind.id)
        already_child.add(ind.id)

    # ------------------------------------------------------------------
    # Pass 4b – link uncoded person rows by parent name
    #
    # Half-coded files record their later generations without path codes: each
    # such person (role "unknown") names their parents in the Father/Mother
    # columns instead. Reuse Pass 4's resolution to give them FAMC links, but
    # first widen the name index so the informal references these files use
    # actually resolve to the real person rather than minting a duplicate:
    #   * mothers are named by MAIDEN surname though filed under the married one
    #     ("Clara Haydon" = Clara Luke née Haydon)
    #   * people are referenced by NICKNAME ("Dick Julian" = Richard Julian) or
    #     initials ("R.H. Luke" = Richard Haydon Luke), both held in the NICK
    #   * or by a middle name ("Leslie Luke" = Alfred Leslie Luke)
    # A token (given word / nickname) is only indexed when it identifies exactly
    # one person under that surname, so a shared middle name like "Mary" can
    # never mislink. References given only by surname ("? Pearce") still miss
    # and become a mother-only or synthetic parent, reconciled in the audit.
    # ------------------------------------------------------------------
    if profile.name_link_uncoded:
        def _first_word(s: str | None) -> str:
            s = (s or "").lower().split("(")[0].strip()
            return s.split()[0] if s.split() else ""

        token_owners: dict[tuple[str, str], set[str]] = {}
        row_aliases: list[tuple[Individual, set[str], set[str]]] = []
        for prow in person_rows:
            ind = dedup_map[prow["dedup_key"]]
            surnames = {prow["surname_base"].upper()}
            if prow["maiden"]:
                surnames.add(prow["maiden"].upper())
            tokens = set(
                (prow["given"] or "").lower()
                .replace("(", " ").replace(")", " ").split()
            )
            if prow.get("nickname"):
                tokens.add(_first_word(prow["nickname"]))
            tokens.discard("")
            row_aliases.append((ind, surnames, tokens))
            for sn in surnames:
                for tk in tokens:
                    token_owners.setdefault((sn, tk), set()).add(ind.id)
        for ind, surnames, tokens in row_aliases:
            for sn in surnames:
                for tk in tokens:
                    if len(token_owners[(sn, tk)]) == 1:
                        existing_by_name.setdefault((sn, tk), ind)

        # The genealogist's generation numbers disambiguate a parent name that
        # is otherwise ambiguous within its surname — a parent sits exactly one
        # generation above their child.
        gen_by_id: dict[str, float] = {}
        for prow in person_rows:
            g = v(prow["row"], _C_GENERATION)
            if isinstance(g, float):
                gen_by_id.setdefault(dedup_map[prow["dedup_key"]].id, g)
        uncoded_ind_by_id = {i.id: i for i in dedup_map.values()}

        # All individuals sharing each surname, for the initials-aware matcher.
        surname_people: dict[str, set[str]] = {}
        for ind, surnames, _tokens in row_aliases:
            for sn in surnames:
                surname_people.setdefault(sn, set()).add(ind.id)

        def _is_initial(tok: str) -> bool:
            """A single-letter token (with optional dot): 'j', 'j.' — too weak
            to key a match on, and prone to collide with others' middle initials
            (the father ref "J. A. Bruce" must not match "Hazel J.")."""
            return bool(re.fullmatch(r"[a-z]\.?", tok))

        def _given_tokens(g: str | None) -> list[str]:
            return (g or "").lower().replace("(", " ").replace(")", " ").split()

        def _tok_compat(a: str, b: str) -> bool:
            """Two given-name tokens are compatible if equal, one is an initial
            matching the other's first letter ('j.'~'james'), or one is a
            ≥3-char abbreviation prefix of the other ('fred.'~'frederick')."""
            a, b = a.rstrip("."), b.rstrip(".")
            if a == b:
                return True
            if len(a) == 1 and b.startswith(a):
                return True
            if len(b) == 1 and a.startswith(b):
                return True
            return (len(a) >= 3 and b.startswith(a)) or (len(b) >= 3 and a.startswith(b))

        def _initials_match(
            parsed: tuple[str | None, str | None] | None, child_gen: float | None
        ) -> Individual | None:
            """Resolve a parent referenced by initials to their full-name record
            — e.g. the children of "J. A. Bruce Steele" mean James Alexander
            Bruce Steele. Among same-surname people one generation up, keep those
            whose full given name is initial-compatible with the reference, token
            for token; accept the unique one. No-code files only."""
            if not parsed or not parsed[0] or child_gen is None or not parsed[1]:
                return None
            ref = _given_tokens(parsed[0])
            if not ref:
                return None
            hits = set()
            for iid in surname_people.get(parsed[1].upper(), set()):
                if gen_by_id.get(iid) != child_gen + 1:
                    continue
                cand = _given_tokens(uncoded_ind_by_id[iid].given_name)
                if (len(ref) <= len(cand)
                        and all(_tok_compat(ref[i], cand[i]) for i in range(len(ref)))):
                    hits.add(iid)
            return uncoded_ind_by_id.get(next(iter(hits))) if len(hits) == 1 else None

        def _gen_disambig(
            parsed: tuple[str | None, str | None] | None, child_gen: float | None
        ) -> Individual | None:
            """Pick the generation-correct individual when a parent name collides
            with another of the same surname (e.g. "Ernest Luke" = both Alfred
            Ernest, nicknamed Ernest, and his nephew Ernest George). Only acts
            when the generation singles out exactly one candidate."""
            if not parsed or not parsed[0] or child_gen is None or not parsed[1]:
                return None
            token = _first_word(parsed[0])
            if _is_initial(token):
                return None  # an initial is too weak to disambiguate on
            cands = token_owners.get((parsed[1].upper(), token), set())
            if len(cands) <= 1:
                return None
            up = [c for c in cands if gen_by_id.get(c) == child_gen + 1]
            if len(up) == 1:
                return uncoded_ind_by_id.get(up[0])
            # Several same-generation candidates: the referenced name collides
            # with another's *middle* name (e.g. "James Steele" matches both
            # "James" b.1928 and his brother "Matthew James Bruce"). Prefer the
            # candidate whose FIRST given word is the referenced name.
            first = [c for c in up
                     if _first_word(uncoded_ind_by_id[c].given_name) == token]
            return uncoded_ind_by_id.get(first[0]) if len(first) == 1 else None

        def _gen_variant_match(
            parsed: tuple[str | None, str | None] | None, child_gen: float | None
        ) -> Individual | None:
            """Resolve a parent whose surname the source spells slightly
            differently from the parent's own row, using the generation to pick.

            Brc:Stl charts the Scottish ancestors as "STEEL" but their Australian
            descendants (and the parent references in those rows) as "STEELE", so
            James Steel b.1829's children name their father "James Steele" and the
            exact-surname index misses him — leaving them to mint a fresh synthetic
            or wrongly attach to his same-named son (b.1861). Gather candidates
            across the exact surname *and* any one-edit spelling variant present in
            the index, then accept the single one that sits exactly one generation
            above the child. Only used by the no-code files."""
            if not parsed or not parsed[0] or child_gen is None or not parsed[1]:
                return None
            token = _first_word(parsed[0])
            if _is_initial(token):
                return None
            target_sn = parsed[1].upper()
            cands: set[str] = set()
            for (sn, tk), owners in token_owners.items():
                if tk == token and _similar_surname(sn, target_sn):
                    cands |= owners
            up = [c for c in cands if gen_by_id.get(c) == child_gen + 1]
            return uncoded_ind_by_id.get(up[0]) if len(up) == 1 else None

        # People whose surname is the genealogist's literal "?" (unknown
        # surname), indexed by (first given word, generation). A chart references
        # such a person as a parent by given name plus a "?" surname
        # ("Samantha ?"), which _parse_parent_name reduces to a given-only
        # reference — so without this the reference mints a duplicate synthetic
        # instead of reusing the real row (e.g. three Samanthas). Because "?" is
        # the genealogist's explicit "unknown", a given-only reference may unify
        # with the *unique* same-given "?"-surname person exactly one generation
        # above the child. No-code files only.
        qmark_by_given_gen: dict[tuple[str, float], set[str]] = {}
        for _iid, _indiv in uncoded_ind_by_id.items():
            if (_indiv.surname or "").strip() == "?":
                _g = gen_by_id.get(_iid)
                _tok = _first_word(_indiv.given_name)
                if _g is not None and _tok and _tok != "?":
                    qmark_by_given_gen.setdefault((_tok, _g), set()).add(_iid)

        def _qmark_match(
            parsed: tuple[str | None, str | None] | None, child_gen: float | None
        ) -> Individual | None:
            """Link a given-only parent reference (its surname was the
            genealogist's "?" unknown marker) to the unique same-given
            "?"-surname person one generation above the child, collapsing what
            would otherwise be a duplicate synthetic. Given-only refs only."""
            if not parsed or not parsed[0] or parsed[1] or child_gen is None:
                return None
            tok = _first_word(parsed[0])
            if not tok or _is_initial(tok):
                return None
            cands = qmark_by_given_gen.get((tok, child_gen + 1), set())
            return uncoded_ind_by_id.get(next(iter(cands))) if len(cands) == 1 else None

        def _parent_family_uncoded(pkey: tuple[str | None, str | None]) -> Family:
            """Family for a name-linked child's parents, reusing one a path code
            already built for the same couple so children don't spawn a phantom
            duplicate of a real (e.g. Arthur Nagle + Lorna Kerslake) family."""
            nonlocal _fam_counter
            if pkey in parent_fams:
                return parent_fams[pkey]
            if pkey[0] and pkey[1]:
                for f in families:
                    if f.husband_id == pkey[0] and f.wife_id == pkey[1]:
                        parent_fams[pkey] = f
                        return f
            _fam_counter += 1
            pf = Family(id=f"F{_fam_counter}", husband_id=pkey[0], wife_id=pkey[1])
            families.append(pf)
            parent_fams[pkey] = pf
            return pf

        for prow in person_rows:
            if prow["role"] != "unknown":
                continue
            ind = dedup_map[prow["dedup_key"]]
            if ind.id in already_child:
                continue
            fp = _parse_parent_name(prow["father_raw"])
            mp = _parse_parent_name(prow["mother_raw"])
            # A parent known only by surname (or wholly unknown) is unknown
            # ancestry — don't mint a nameless placeholder for it.
            if not (fp and fp[0]):
                fp = None
            if not (mp and mp[0]):
                mp = None
            if fp is None and mp is None:
                continue
            child_gen = gen_by_id.get(ind.id)
            avoid = frozenset(_descendants(ind.id) | {ind.id})
            spell_tolerant = profile.code_convention == "none"
            father_ind = (
                _gen_disambig(fp, child_gen)
                or (_gen_variant_match(fp, child_gen) if spell_tolerant else None)
                or _initials_match(fp, child_gen)
                or (_qmark_match(fp, child_gen) if spell_tolerant else None)
                or _resolve_parent(fp, "M", avoid, ind.birth_date))
            mother_ind = (
                _gen_disambig(mp, child_gen)
                or (_gen_variant_match(mp, child_gen) if spell_tolerant else None)
                or _initials_match(mp, child_gen)
                or (_qmark_match(mp, child_gen) if spell_tolerant else None)
                or _resolve_parent(mp, "F", avoid, ind.birth_date,
                                   co_parent=father_ind))
            if father_ind is None and mother_ind is None:
                continue
            pkey = (father_ind.id if father_ind else None,
                    mother_ind.id if mother_ind else None)
            pf = _parent_family_uncoded(pkey)
            pf.child_ids.append(ind.id)
            already_child.add(ind.id)
            if diag is not None:
                diag["name_linked_family_ids"].add(pf.id)

    # ------------------------------------------------------------------
    # Pass 5 – flag-based event blocks
    #
    # Some ancestry is recorded not with generation numbers and codes but as
    # free-standing B (birth) / D (death) / M (marriage) event rows — e.g. the
    # Forster line behind Maud Ponting (née Forster), who is already in the
    # main tree. Merge each person's B and D rows, recover birth/death/notes,
    # and link them via the parent-name columns, reusing Pass-4 resolution so
    # block people unify with the synthetic parents the main tree already
    # implied (e.g. Maud's father "John Forster") instead of duplicating them.
    # ------------------------------------------------------------------
    def _parent_family(fid: str | None, mid: str | None) -> Family:
        nonlocal _fam_counter
        # Reuse a half-known couple (father-only or mother-only) rather than
        # splitting a sibling group across two families.
        for cand in ((fid, mid), (fid, None), (None, mid)):
            if cand == (None, None):
                continue
            fam = parent_fams.get(cand)
            if fam is not None:
                if fid and fam.husband_id is None:
                    fam.husband_id = fid
                if mid and fam.wife_id is None:
                    fam.wife_id = mid
                parent_fams.pop(cand, None)
                parent_fams[(fam.husband_id, fam.wife_id)] = fam
                return fam
        _fam_counter += 1
        fam = Family(id=f"F{_fam_counter}", husband_id=fid, wife_id=mid)
        families.append(fam)
        parent_fams[(fid, mid)] = fam
        return fam

    block_people: dict[tuple, dict] = {}
    block_order: list[tuple] = []
    # A block's marriage row ("John PRICE married Mary MULDOON") establishes the
    # couple for the children listed after it. Those children carry no parent
    # names of their own — the genealogist grouped them by position under the
    # marriage — so they inherit the marriage row's spouses as their parents.
    # Only the no-code files (which position rather than code) reach this; file
    # #1's blocks name each child's parents explicitly and never set a section.
    section_father = section_mother = ""
    for r in range(DATA_START_ROW, DATA_END_ROW):
        if r in _excluded_rows:
            continue
        if profile.name_link_uncoded and fv(r) == "M":
            parts = re.split(r"\bmarried\b", str(v(r, _C_GIVEN)),
                             maxsplit=1, flags=re.IGNORECASE)
            if len(parts) == 2:
                section_father, section_mother = parts[0].strip(), parts[1].strip()
            continue
        if fv(r) not in ("B", "D"):
            continue
        if v(r, _C_GENERATION) != "":
            continue
        surname = str(v(r, _C_SURNAME)).strip()
        if not surname:
            continue
        if diag is not None:
            diag["consumed_rows"].add(r)
        flag = fv(r)
        # Strip clarifying annotations like "John {Maud's father}".
        given = re.sub(r"\s*\{.*?\}", "", str(v(r, _C_GIVEN)).strip()).strip()
        given, given_annotation, nickname = _clean_given(given)
        surname_base = _strip_nee(surname)
        father_raw = str(v(r, _C_FATHER)).strip()
        mother_raw = str(v(r, _C_MOTHER)).strip()
        # A child listed under a block marriage row names no parents of its own;
        # adopt that couple so it links to them (and unifies with the synthetic
        # parents the main tree already implied for them).
        if (profile.name_link_uncoded and not father_raw and not mother_raw
                and (section_father or section_mother)):
            father_raw, mother_raw = section_father, section_mother
        f_first = father_raw.split()[0].lower() if father_raw and father_raw != "?" else "?"
        m_first = mother_raw.split()[0].lower() if mother_raw and mother_raw != "?" else "?"
        # Parents distinguish same-named people (e.g. two different "John"s).
        key = (surname_base.upper(), given.lower(), f_first, m_first)
        rec = block_people.get(key)
        if rec is None:
            rec = {
                "given": given, "surname_base": surname_base,
                "maiden": _maiden_name(surname),
                "nickname": nickname,
                "father_raw": father_raw, "mother_raw": mother_raw,
                "birth": None, "birth_place": None,
                "death": None, "death_place": None,
                "sex": _infer_sex(surname), "notes": [],
            }
            block_people[key] = rec
            block_order.append(key)
        elif nickname and not rec.get("nickname"):
            rec["nickname"] = nickname
        block_notes = list(row_notes.get(r, []))
        if given_annotation:
            block_notes.append(f"Name annotation: [{given_annotation}].")
        pn = _date_precision_note(v(r, _C_DATE1), "Birth" if flag == "B" else "Death")
        if pn:
            block_notes.append(pn)
        for n in block_notes:
            if n not in rec["notes"]:
                rec["notes"].append(n)
        date = _parse_date(v(r, _C_DATE1))
        if flag == "B":
            rec["birth"] = rec["birth"] or date
            # Capture the birth place too — only for the no-code files, so file
            # #1's existing block output stays byte-identical.
            if profile.name_link_uncoded and not rec["birth_place"]:
                rec["birth_place"] = _birth_place(r)
        else:
            rec["death"] = rec["death"] or date
            place = _standardise_place(str(v(r, _C_TOWN)).strip())
            if place and not rec["death_place"]:
                rec["death_place"] = place

    # Full-given-name index over every individual built so far, used to unify
    # a block person with the record the main tree already implied (e.g. Maud,
    # or her father "John Forster"). Keyed on the *whole* given name so
    # "John T." stays distinct from "John".
    def _full_key(given: str | None, surname: str | None) -> tuple[str, str]:
        return ((surname or "").upper(), (given or "").lower().split("(")[0].strip())

    full_index: dict[tuple[str, str], Individual] = {}
    for ind in list(dedup_map.values()) + list(synth_by_name.values()):
        full_index.setdefault(_full_key(ind.given_name, ind.surname), ind)
    # For the no-code files, a block person can be the same individual the main
    # tree already records under a *married* surname — e.g. block "Dorothy Jane
    # PRICE" is main-tree "Dorothy Jane STEELE née Price". Register née people
    # under their maiden surname too so the block unifies with them rather than
    # minting a duplicate.
    if profile.name_link_uncoded:
        for pr in person_rows:
            if pr["maiden"]:
                ind = dedup_map[pr["dedup_key"]]
                full_index.setdefault(_full_key(ind.given_name, pr["maiden"]), ind)

    # Materialise each block person, unifying with existing/synthetic records.
    # Same full name + different parents (e.g. John Forster the father vs John
    # Forster the infant brother) must stay distinct: only the first occurrence
    # of a given name unifies with a pre-existing record; later ones are new.
    block_created: list[Individual] = []
    block_fullname_seen: set[tuple[str, str]] = set()
    for key in block_order:
        rec = block_people[key]
        nk = _ind_name_key(rec["given"], rec["surname_base"])
        fk = _full_key(rec["given"], rec["surname_base"])
        ind = None if fk in block_fullname_seen else full_index.get(fk)
        # A née block woman is filed under her married surname here but the main
        # tree referenced her by maiden name (e.g. block "Mary Ann PRICE née
        # Muldoon" = the synthetic "Mary Muldoon" implied as Dorothy's mother).
        # Fall back to her maiden key so the two unify. No-code files only.
        if ind is None and profile.name_link_uncoded and rec.get("maiden"):
            mk = _ind_name_key(rec["given"], rec["maiden"])
            ind = synth_by_name.get(mk) or existing_by_name.get(mk)
        block_fullname_seen.add(fk)
        if ind is None:
            _id_counter += 1
            ind = Individual(
                id=f"I{_id_counter}", given_name=rec["given"],
                surname=rec["surname_base"], nickname=rec.get("nickname"),
                sex=rec["sex"],
            )
            block_created.append(ind)
        # Pin the loose name keys to the first (primary) holder so a later
        # same-named block person can't redirect parent lookups to itself.
        synth_by_name.setdefault(nk, ind)
        existing_by_name.setdefault(nk, ind)
        full_index.setdefault(fk, ind)
        if not ind.birth_date and rec["birth"]:
            ind.birth_date = rec["birth"]
        if not ind.birth_place and rec.get("birth_place"):
            ind.birth_place = rec["birth_place"]
        if not ind.death_date and rec["death"]:
            ind.death_date = rec["death"]
        if not ind.death_place and rec["death_place"]:
            ind.death_place = rec["death_place"]
        if ind.sex is None and rec["sex"]:
            ind.sex = rec["sex"]
        for n in rec["notes"]:
            if n not in ind.note_list:
                ind.note_list.append(n)
        rec["ind"] = ind
        # Married women are referenced as a mother by given name only; register
        # a given-name alias so those parent lookups unify with this record.
        if rec["sex"] == "F" and nk[1]:
            synth_by_name.setdefault(("", nk[1]), ind)

    # Link block people to their parents.
    for key in block_order:
        rec = block_people[key]
        ind = rec["ind"]
        if ind.id in already_child:
            continue
        fp = _parse_parent_name(rec["father_raw"])
        mp = _parse_parent_name(rec["mother_raw"])
        if fp is None and mp is None:
            continue
        avoid = frozenset(_descendants(ind.id) | {ind.id})
        father_ind = _resolve_parent(fp, "M", avoid)
        mother_ind = _resolve_parent(mp, "F", avoid, co_parent=father_ind)
        if father_ind is None and mother_ind is None:
            continue
        fam = _parent_family(
            father_ind.id if father_ind else None,
            mother_ind.id if mother_ind else None,
        )
        fam.child_ids.append(ind.id)
        already_child.add(ind.id)

    # Attach the block's marriage rows (skipped by Pass 3 as no family existed
    # for them then) to the now-built couples, matching both spouses by name.
    all_ind_by_id = {
        i.id: i for i in list(dedup_map.values()) + list(synth_by_name.values())
    }

    def _spouse_matches(ind: Individual | None, tokens: set[str]) -> bool:
        if ind is None:
            return False
        g = (ind.given_name or "").lower().split("(")[0].split()
        s = (ind.surname or "").lower().strip()
        first = g[0] if g else ""
        if s and s not in tokens:
            return False
        if first and first not in tokens:
            return False
        return bool(s or first)

    for m in marriage_rows:
        if m.get("attached") or (m["date"] is None and m["place"] is None):
            continue
        tokens = set(re.findall(r"[a-z]+", m["combined"].lower()))
        for fam in families:
            if fam.marriage_date is not None:
                continue
            h = all_ind_by_id.get(fam.husband_id)
            w = all_ind_by_id.get(fam.wife_id)
            if h and w and _spouse_matches(h, tokens) and _spouse_matches(w, tokens):
                fam.marriage_date = m["date"]
                fam.marriage_place = m["place"]
                fam.note_list.extend(m["cell_notes"])
                m["attached"] = True
                break

    # ------------------------------------------------------------------
    # Reconcile a coded child whose recorded mother NAME names a *different*
    # wife of the same father than the one the path code placed them under.
    # Charles Nicholls's children are coded under his (second) marriage to Lydia
    # Ann Thomas, but each names its mother as "Mary Thomas" — his first wife
    # Mary Louisa (née Thomas), whose own (childless) family already exists via
    # the ex-spouse path. Move such children onto the parent couple the source
    # actually names. General, but only fires when the father has two recorded
    # wives and the child's mother name matches the other one.
    # ------------------------------------------------------------------
    recon_by_id = {i.id: i for i in dedup_map.values()}
    maiden_by_id = {
        dedup_map[pr["dedup_key"]].id: pr["maiden"]
        for pr in person_rows if pr["maiden"]
    }
    child_mother: dict[str, str] = {}
    for pr in person_rows:
        if pr["role"] == "child" and pr["mother_raw"]:
            child_mother.setdefault(dedup_map[pr["dedup_key"]].id, pr["mother_raw"])

    def _names_wife(parsed: tuple[str | None, str | None] | None,
                    wife: Individual | None) -> bool:
        if parsed is None or wife is None:
            return False
        g, s = parsed
        surnames = {(wife.surname or "").upper()}
        m = maiden_by_id.get(wife.id)
        if m:
            surnames.add(m.upper())
        if s and s.upper() not in surnames:
            return False
        if g:
            gw = g.lower().split()[0]
            tokens = (wife.given_name or "").lower().replace("(", " ").split()
            if gw not in tokens and gw != (wife.nickname or "").lower():
                return False
        return bool(s or g)

    fams_by_husband: dict[str, list[Family]] = {}
    for f in families:
        if f.husband_id:
            fams_by_husband.setdefault(f.husband_id, []).append(f)
    for f in families:
        if not (f.husband_id and f.wife_id):
            continue
        co_wives = [g for g in fams_by_husband[f.husband_id]
                    if g is not f and g.wife_id]
        if not co_wives:
            continue
        wife = recon_by_id.get(f.wife_id)
        for cid in list(f.child_ids):
            parsed = _parse_parent_name(child_mother.get(cid, ""))
            if not (parsed and parsed[0]) or _names_wife(parsed, wife):
                continue
            target = next(
                (g for g in co_wives
                 if _names_wife(parsed, recon_by_id.get(g.wife_id))),
                None,
            )
            if target is not None and cid not in target.child_ids:
                f.child_ids.remove(cid)
                target.child_ids.append(cid)

    # ------------------------------------------------------------------
    # Pass 6 – no-generation "compact descendant" adjacency
    #
    # A tail of rows carry neither a generation number nor a path code: the
    # genealogist drew them as a compact descendant chart, positioned rather
    # than coded — a married-in spouse on the row directly ABOVE a née-female
    # anchor, that couple's children on the rows directly BELOW, several of them
    # nameless (surname only). Walk the surname-bearing rows in sheet order and:
    #   * pair a née-female with the immediately preceding row when its surname
    #     equals her *married* surname — that row is her husband, be it a no-gen
    #     orphan, a gen-numbered Sp spouse, or her own bloodline husband;
    #   * treat a no-gen née-female as a married daughter, filing her under the
    #     family (parents' or current couple's) whose surname matches her
    #     *maiden* name;
    #   * attach a plain no-gen row as a child of the current couple, giving a
    #     nameless one a placeholder name and an explanatory note.
    # Only the half-coded files set name_link_uncoded; the others never enter.
    # ------------------------------------------------------------------
    nogen_created: list[Individual] = []
    if profile.name_link_uncoded:
        child_to_fam: dict[str, Family] = {}
        for fam in families:
            for cid in fam.child_ids:
                child_to_fam.setdefault(cid, fam)
        row_to_ind = {pr["row"]: dedup_map[pr["dedup_key"]] for pr in person_rows}
        all_by_id: dict[str, Individual] = {
            i.id: i for i in list(dedup_map.values())
            + list(synth_by_name.values()) + block_created
        }

        def _new_ind(given: str, surname: str, row: int,
                     extra_note: str | None = None) -> Individual:
            nonlocal _id_counter
            _id_counter += 1
            ind = Individual(id=f"I{_id_counter}", given_name=given, surname=surname,
                             note_list=list(row_notes.get(row, [])))
            if extra_note:
                ind.note_list.append(extra_note)
            nogen_created.append(ind)
            all_by_id[ind.id] = ind
            if diag is not None:
                diag["consumed_rows"].add(row)
            return ind

        def _fam_surname(fam: Family | None) -> str:
            if fam is None:
                return ""
            for sid in (fam.husband_id, fam.wife_id):
                s = all_by_id.get(sid) if sid else None
                if s and s.surname:
                    return s.surname.upper()
            return ""

        def _marriage_family(wife: Individual, husband: Individual | None) -> Family:
            nonlocal _fam_counter
            for fam in families:
                if fam.wife_id == wife.id:
                    if husband is not None and fam.husband_id is None:
                        fam.husband_id = husband.id
                    return fam
            # Complete a husband-only parent family — children recorded with an
            # unnamed mother (e.g. Alfred Ernest Luke's daughter Lillian, whose
            # mother column is "?") — with the wife the compact chart pairs him
            # with, rather than splitting his marriage across two families.
            if husband is not None:
                for fam in families:
                    if (fam.husband_id == husband.id and fam.wife_id is None
                            and fam.child_ids):
                        fam.wife_id = wife.id
                        return fam
            _fam_counter += 1
            fam = Family(id=f"F{_fam_counter}",
                         husband_id=husband.id if husband else None,
                         wife_id=wife.id)
            families.append(fam)
            return fam

        # State carried down the compact chart.
        cur_family: Family | None = None     # couple whose plain no-gen kids attach
        cur_head: Individual | None = None    # a male head awaiting his own children
        parent_family: Family | None = None   # sibling-group parents (married-daughter)
        pending: Individual | None = None     # a deferred plain no-gen row
        pending_surname = ""
        pending_nameless = False
        prev_ind: Individual | None = None
        prev_surname = ""
        prev_nee = False

        def _flush() -> None:
            """Place a deferred plain no-gen row as a child of the current couple."""
            nonlocal pending, pending_surname, pending_nameless, cur_family, _fam_counter
            if pending is None:
                return
            fam = cur_family
            if fam is None and cur_head is not None:
                _fam_counter += 1
                fam = Family(id=f"F{_fam_counter}", husband_id=cur_head.id)
                families.append(fam)
                cur_family = fam
            if fam is not None:
                fam.child_ids.append(pending.id)
                if diag is not None:
                    diag["name_linked_family_ids"].add(fam.id)
                if pending_nameless:
                    pending.note_list.append(
                        "Recorded in the source as an unnamed child of this family.")
            pending = None
            pending_surname = ""
            pending_nameless = False

        def _claim_husband(married_surname: str) -> Individual | None:
            """A née-female's husband is the row immediately above with the same
            (married) surname — either the deferred no-gen row or the previous
            bloodline person."""
            nonlocal pending, pending_surname, pending_nameless
            if married_surname in ("", "?"):
                return None
            if pending is not None and pending_surname == married_surname:
                h = pending
                if pending_nameless:
                    h.note_list.append("Recorded in the source as an unnamed spouse.")
                pending = None
                pending_surname = ""
                pending_nameless = False
                return h
            if (prev_ind is not None and prev_surname == married_surname
                    and not prev_nee):
                return prev_ind
            return None

        for r in range(DATA_START_ROW, DATA_END_ROW):
            if r in _excluded_rows:
                continue
            code = str(v(r, _C_CODE)).strip().replace("|", "/")
            surname_raw = str(v(r, _C_SURNAME)).strip()
            given_raw = str(v(r, _C_GIVEN)).strip()
            if (fv(r) in ("B", "D", "M")
                    or (profile.code_convention == "none"
                        and _is_marriage_label(given_raw))
                    or (not surname_raw and not given_raw and not code)):
                # B/D flag rows are free-standing event blocks — Pass 5's domain;
                # handling them here too would mint a duplicate of every block
                # person. 'M' rows and pure-layout rows are likewise not people.
                continue
            if code:
                # Coded rows are Pass 3's work; they also break a compact block.
                _flush()
                cur_family = cur_head = parent_family = None
                prev_ind, prev_surname, prev_nee = None, "", False
                continue

            is_gen = v(r, _C_GENERATION) != ""
            surname_base = _strip_nee(surname_raw)
            maiden = _maiden_name(surname_raw)
            is_nee = bool(maiden) or "née" in surname_raw.lower()
            given, _ann, _nick = _clean_given(given_raw)
            nameless = given in ("", "?")
            married_up = surname_base.upper()

            if is_gen:
                ind = row_to_ind.get(r)
                if ind is None:
                    continue
                if is_nee:
                    husband = _claim_husband(married_up)
                    if husband is None:
                        _flush()
                    cur_family = _marriage_family(ind, husband)
                    cur_head = None
                else:
                    _flush()
                    cur_head = ind if ind.sex != "F" else None
                    cur_family = None
                famc = child_to_fam.get(ind.id)
                if famc is not None:
                    parent_family = famc
                prev_ind, prev_surname, prev_nee = ind, married_up, is_nee
                continue

            # --- no-generation row ---
            if is_nee:
                # A married daughter: file under her maiden (or married) surname,
                # pair a husband if the row above matches her married surname, and
                # attach her to whichever known family matches her maiden name.
                surname = (surname_base if surname_base not in ("", "?")
                           else (maiden or "?").upper())
                ind = _new_ind("[Unnamed]" if nameless else given, surname, r)
                husband = _claim_husband(married_up)
                if husband is None:
                    _flush()
                if husband is not None:
                    _marriage_family(ind, husband)
                maiden_up = (maiden or "").upper()
                if parent_family is not None and _fam_surname(parent_family) == maiden_up:
                    target = parent_family
                elif cur_family is not None and _fam_surname(cur_family) == maiden_up:
                    target = cur_family
                else:
                    target = parent_family or cur_family
                if target is not None:
                    target.child_ids.append(ind.id)
                    if diag is not None:
                        diag["name_linked_family_ids"].add(target.id)
                if surname_base in ("", "?"):
                    ind.note_list.append(
                        "Recorded as a married daughter; her married surname is "
                        "not given in the source.")
                elif nameless:
                    ind.note_list.append(
                        "Recorded in the source without a given name.")
                prev_ind, prev_surname, prev_nee = ind, married_up, True
            else:
                # A plain no-gen row: defer one step so a following née-female can
                # claim it as her husband; otherwise it is a child of the couple.
                _flush()
                pending = _new_ind("[Unnamed]" if nameless else given,
                                   surname_base or "?", r)
                pending_surname = married_up
                pending_nameless = nameless
        _flush()

        # Pair any still-orphaned married-in spouse (a gen Sp row in no family) to
        # a née-woman whose recorded successive married surnames name his surname —
        # e.g. George Hartely is Lorna Kerslake's second husband ("NAGLE then
        # HARTLEY"), listed far from her among her descendants, so adjacency cannot
        # reach him. A one-edit Damerau distance tolerates the source's spelling
        # wobble (HARTELY/HARTLEY) without risking a loose match.
        in_family = {sid for fam in families
                     for sid in (fam.husband_id, fam.wife_id) if sid}
        in_family |= {c for fam in families for c in fam.child_ids}
        married_women = [w for w in dedup_map.values() if w.married_surnames]
        for prow in person_rows:
            if prow.get("marker") != "Sp":
                continue
            ind = dedup_map[prow["dedup_key"]]
            if ind.id in in_family:
                continue
            sb = (ind.surname or "").upper()
            if not sb:
                continue
            match = next(
                (w for w in married_women
                 if any(_similar_surname(sb, ms.upper()) for ms in w.married_surnames)),
                None,
            )
            if match is None:
                continue
            _fam_counter += 1
            families.append(Family(id=f"F{_fam_counter}",
                                   husband_id=ind.id, wife_id=match.id))
            in_family.add(ind.id)

        # Per-person marriage details attach to the name-linked families built
        # above. The earlier per-person pass ran before these families existed and
        # only indexed coded husband/wife rows, so a name-linked couple whose
        # marriage sits on their own rows rather than on a separate 'M'-flag row
        # would lose it (e.g. Henry Joseph Costigan + Jane Steel, married 1871, or
        # most of the deep Bruce/Steel couples whose dates are per-person). Index
        # every person row by spouse id and fill any family still lacking a
        # marriage, preferring the spouse in the fewest families so a remarriage's
        # single recorded date cannot leak across both of their marriages.
        row_marr2: dict[str, dict] = {}
        for prow in person_rows:
            if not (prow["marr_date"] or prow["marr_place"]):
                continue
            row_marr2.setdefault(dedup_map[prow["dedup_key"]].id, prow)
        if row_marr2:
            fam_count: dict[str, int] = {}
            for fam in families:
                for sid in (fam.husband_id, fam.wife_id):
                    if sid:
                        fam_count[sid] = fam_count.get(sid, 0) + 1
            for fam in families:
                if fam.marriage_date or fam.marriage_place:
                    continue
                cands = [s for s in (fam.husband_id, fam.wife_id)
                         if s and s in row_marr2]
                cands.sort(key=lambda s: (fam_count.get(s, 0),
                                          0 if s == fam.husband_id else 1))
                if cands:
                    src = row_marr2[cands[0]]
                    fam.marriage_date = src["marr_date"]
                    fam.marriage_place = src["marr_place"]

    # Link a child recorded as adopted (father AND mother columns both literally
    # "[Adopted]") to the family it is drawn under in the chart. Such a child
    # names no real parents, so the parent-name passes leave it orphaned even
    # though the genealogist positioned it directly beneath an adoptive couple.
    # Attach it to the nearest preceding spouse's family one generation up, with
    # a "PEDI adopted" linkage so the relationship is recorded as adoptive, not
    # biological. The verbatim "[Adopted]" note stays on the child. No-code
    # files only (the only files with this gen-numbered, name-linked layout).
    if profile.name_link_uncoded:
        adopt_gen_by_id: dict[str, float] = {}
        for prow in person_rows:
            g = v(prow["row"], _C_GENERATION)
            if isinstance(g, float):
                adopt_gen_by_id.setdefault(dedup_map[prow["dedup_key"]].id, g)
        linked_children = {cid for fam in families for cid in fam.child_ids}
        fam_by_spouse: dict[str, Family] = {}
        for fam in families:
            for sid in (fam.husband_id, fam.wife_id):
                if sid:
                    fam_by_spouse.setdefault(sid, fam)
        rows_sorted = sorted(person_rows, key=lambda pr: pr["row"])
        for i, prow in enumerate(rows_sorted):
            fr = (prow.get("father_raw") or "").lower()
            mr = (prow.get("mother_raw") or "").lower()
            if "[adopt" not in fr or "[adopt" not in mr:
                continue
            child = dedup_map[prow["dedup_key"]]
            if child.id in linked_children:
                continue
            cg = adopt_gen_by_id.get(child.id)
            if cg is None:
                continue
            target = None
            for prev in reversed(rows_sorted[:i]):
                pid = dedup_map[prev["dedup_key"]].id
                if adopt_gen_by_id.get(pid) == cg + 1 and pid in fam_by_spouse:
                    target = fam_by_spouse[pid]
                    break
            if target is not None:
                target.child_ids.append(child.id)
                linked_children.add(child.id)
                child.adopted_famc.add(target.id)

    # Mark each divorced spouse's family as ended in divorce (1 DIV). Done here,
    # after every family is built, so the name-linked families (Pass 4b/6) that
    # a divorced spouse may head are covered too, not only the Pass-3 ones.
    if divorced_spouse_ids:
        for fam in families:
            if (fam.husband_id in divorced_spouse_ids
                    or fam.wife_id in divorced_spouse_ids):
                fam.divorced = True

    # Maiden-aware same-person merge (no-code files). A née woman recorded once
    # under her birth surname and once under a married surname keeps two dedup
    # keys, so the surname-keyed Pass 2b misses her and she builds two separate
    # records — each correctly married into its own family. When her given name,
    # BOTH parent names and the exact birth date match another individual she is
    # certainly one person (the genealogist flags the re-appearance with the 'X'
    # marker): merge the married-name record onto the birth-name one, repoint her
    # families, and keep the married surname. Run AFTER all families are built so
    # BOTH of her marriages survive — e.g. Stiff:Taylor's Isabella Taylor,
    # de-facto partner of Joseph Bowden and later wife of Charles Mullins, born
    # 19 Feb 1851 to John Taylor & Eliza Conway.
    if profile.code_convention == "none":
        def _fw(s: str | None) -> str:
            s = (s or "").lower().split("(")[0].strip()
            return s.split()[0] if s.split() else "?"

        def _ident(surname_up: str, prow: dict) -> tuple:
            return (surname_up, _fw(prow["given"]), _fw(prow["father_raw"]),
                    _fw(prow["mother_raw"]), prow["birth_date"])

        birth_identity: dict[tuple, Individual] = {}
        for prow in person_rows:
            if _has_full_date(prow["birth_date"]):
                ind = dedup_map.get(prow["dedup_key"])
                if ind is not None:
                    birth_identity.setdefault(
                        _ident(prow["surname_base"].upper(), prow), ind)
        merged_any = False
        for prow in person_rows:
            maiden = prow["maiden"]
            if not maiden or not _has_full_date(prow["birth_date"]):
                continue
            ind = dedup_map.get(prow["dedup_key"])
            canon = birth_identity.get(_ident(maiden.upper(), prow))
            if ind is None or canon is None or canon is ind:
                continue
            married = ind.surname
            if (married and married.upper() != (canon.surname or "").upper()
                    and married not in canon.married_surnames):
                canon.married_surnames.append(married)
            _merge_into(canon, ind)
            for fam in families:
                if fam.husband_id == ind.id:
                    fam.husband_id = canon.id
                if fam.wife_id == ind.id:
                    fam.wife_id = canon.id
                if ind.id in fam.child_ids:
                    fam.child_ids = [canon.id if c == ind.id else c
                                     for c in fam.child_ids]
            for dk2, v2 in list(dedup_map.items()):
                if v2 is ind:
                    dedup_map[dk2] = canon
            merged_any = True
        if merged_any:
            # The two appearances often share a parent family, so the merged
            # person can land in one child list twice — collapse duplicates.
            for fam in families:
                seen: set[str] = set()
                fam.child_ids = [c for c in fam.child_ids
                                 if not (c in seen or seen.add(c))]

    # A partner recorded with a non-née '[De-facto]' surname tag never married
    # into that union, so the family must carry no marriage event — yet the other
    # spouse's per-person date may legitimately belong to a *different*, legal
    # marriage and have been copied onto this family too (Stiff:Taylor's Henry
    # Joseph Bowden was de-facto with Isabella Taylor but married Elizabeth
    # Goscomb in 1885; that date wrongly reached the Isabella family). The de-
    # facto mate is the nearest preceding person at the same generation; clear
    # that family's marriage date/place and note the de-facto status. No-code
    # files only.
    if profile.name_link_uncoded:
        rows_by_num = sorted(person_rows, key=lambda pr: pr["row"])
        ind_gen: dict[str, float] = {}
        for pr in person_rows:
            g = v(pr["row"], _C_GENERATION)
            if isinstance(g, float):
                ind_gen.setdefault(dedup_map[pr["dedup_key"]].id, g)
        for idx, pr in enumerate(rows_by_num):
            sa = _surname_annotation(pr["surname"])
            if not sa or "de-facto" not in sa.lower():
                continue
            partner = dedup_map[pr["dedup_key"]]
            pgen = ind_gen.get(partner.id)
            # The mate is the nearest preceding bloodline person (not another
            # married-in 'Sp' spouse — that would be a co-wife) at the same
            # generation as the de-facto partner.
            mate_id = None
            for prev in reversed(rows_by_num[:idx]):
                pid = dedup_map[prev["dedup_key"]].id
                if (pid != partner.id and ind_gen.get(pid) == pgen
                        and prev.get("marker") != "Sp"):
                    mate_id = pid
                    break
            if mate_id is None:
                continue
            for fam in families:
                if {partner.id, mate_id} <= {fam.husband_id, fam.wife_id}:
                    fam.marriage_date = fam.marriage_place = None
                    note = ("Recorded as a de-facto relationship; "
                            "the couple never married.")
                    if note not in fam.note_list:
                        fam.note_list.append(note)

    # Consolidate a wife recorded twice under one husband. When children name
    # their mother by given name only ("Jean") or by her maiden surname, the
    # reference can mint a given-only synthetic OR match a same-given namesake
    # who married into a different family — leaving the children under the wrong
    # mother while the real wife (recorded on her own row, having taken the
    # husband's surname) heads a separate, childless family. Re-home the children
    # onto the wife whose surname matches the husband's and drop the duplicate.
    # Fixes e.g. Clive Steele's children, recorded under "Edna Stiff", landing on
    # Edna Mary Stiff née Tracey instead of his wife Edna May Steele née Stiff;
    # and the seven given-only synthetic mothers (Jean Evans, Madge Lock, ...).
    # No-code files only.
    remove_inds: set[str] = set()
    if profile.code_convention == "none":
        all_inds = {i.id: i for i in (list(dedup_map.values())
                    + list(synth_by_name.values()) + block_created
                    + nogen_created)}
        # Synthetic parents are the name-keyed individuals minted from a parent
        # reference (never from a data row); a mother named only by a reference
        # is one of these and, once her children move to the real wife, is a
        # placeholder to discard. A real wife recorded under a different surname
        # (Aunty Mary, née Tracey) has her own row and other families, so she is
        # only detached from this husband, not removed.
        real_ids = ({i.id for i in dedup_map.values()}
                    | {i.id for i in block_created}
                    | {i.id for i in nogen_created})
        synthetic_ids = {i.id for i in synth_by_name.values()
                         if i.id not in real_ids}

        def _gtoks(ind: Individual) -> set[str]:
            return set((ind.given_name or "").lower().replace(".", " ").split())

        by_husband: dict[str, list[Family]] = {}
        for f in families:
            if f.husband_id:
                by_husband.setdefault(f.husband_id, []).append(f)
        removed_fams: set[str] = set()
        candidate_remove: set[str] = set()
        for hid, fl in by_husband.items():
            husband = all_inds.get(hid)
            if husband is None or not husband.surname:
                continue
            hs = husband.surname.upper()
            # The "correct" wife took the husband's surname (recorded on her own
            # row as his Sp spouse). Act only when exactly one family qualifies.
            correct = [f for f in fl
                       if all_inds.get(f.wife_id) is not None
                       and (all_inds[f.wife_id].surname or "").upper() == hs]
            if len(correct) != 1:
                continue
            cf = correct[0]
            cwife = all_inds[cf.wife_id]
            ctoks = _gtoks(cwife)
            for f in fl:
                owife = all_inds.get(f.wife_id)
                if (f is cf or f.id in removed_fams or owife is None
                        or f.wife_id == cf.wife_id):
                    continue
                # A duplicate only if the other wife did NOT take the husband's
                # surname yet shares a given token with the correct wife (the
                # same first name recorded two ways / a mis-resolved namesake).
                if (owife.surname or "").upper() == hs or not (_gtoks(owife) & ctoks):
                    continue
                for c in f.child_ids:
                    if c not in cf.child_ids:
                        cf.child_ids.append(c)
                if not cf.marriage_date and f.marriage_date:
                    cf.marriage_date = f.marriage_date
                if not cf.marriage_place and f.marriage_place:
                    cf.marriage_place = f.marriage_place
                for n in f.note_list:
                    if n not in cf.note_list:
                        cf.note_list.append(n)
                # A synthetic placeholder mother IS the correct wife recorded
                # weakly — discard her once her family is gone. A real,
                # differently-surnamed namesake (Aunty Mary, née Tracey) is a
                # distinct person and is only detached from this husband.
                if owife.id in synthetic_ids:
                    candidate_remove.add(owife.id)
                removed_fams.add(f.id)
        if removed_fams:
            families[:] = [f for f in families if f.id not in removed_fams]
        # Discard a synthetic mother once nothing references her any more (her
        # children moved to the real wife, her placeholder family was dropped).
        if candidate_remove:
            referenced = {x for f in families
                          for x in (f.husband_id, f.wife_id, *f.child_ids) if x}
            remove_inds = {i for i in candidate_remove if i not in referenced}

    # Pass 2b can point several dedup keys at one merged individual, so
    # de-duplicate the final list by identity while preserving order.
    seen_ids: set[str] = set()
    ordered: list[Individual] = []
    for ind in (list(dedup_map.values()) + list(synth_by_name.values())
                + block_created + nogen_created):
        if ind.id not in seen_ids and ind.id not in remove_inds:
            seen_ids.add(ind.id)
            ordered.append(ind)

    # A person's role in a family establishes their sex: the husband is male,
    # the wife female. Previously sex was only inferred from a "née" surname,
    # leaving every family-head husband with no SEX. Fill the gap across all
    # families (Pass 3 plus the synthetic/block parents from Passes 4-5),
    # without overriding an explicit inference.
    final_by_id = {ind.id: ind for ind in ordered}
    for fam in families:
        h = final_by_id.get(fam.husband_id)
        if h is not None and h.sex is None:
            h.sex = "M"
        w = final_by_id.get(fam.wife_id)
        if w is not None and w.sex is None:
            w.sex = "F"

    # Emit the married-surname progression as a note once per person (after all
    # merges have unioned the data). Lineage-chart membership stays on
    # Individual.lineage_lines and is written as custom _GROUP tags by the
    # writer, rather than as a repetitive freeform note.
    for ind in ordered:
        if ind.married_surnames:
            ind.note_list.insert(
                0, "Married surname" + ("s" if len(ind.married_surnames) > 1 else "")
                + ": " + ", then ".join(ind.married_surnames) + ".")

    if diag is not None:
        # Minted-record markers: synthetic parents are the name-keyed individuals
        # that never came from a data row (so are absent from every "real" set);
        # placeholders are the unnamed Pass-6 people.
        real_ids = ({i.id for i in dedup_map.values()}
                    | {i.id for i in block_created}
                    | {i.id for i in nogen_created})
        diag["synthetic_ids"] = {
            i.id for i in synth_by_name.values() if i.id not in real_ids
        }
        diag["placeholder_ids"] = {
            i.id for i in nogen_created if i.given_name == "[Unnamed]"
        }
        # Generation number (col 4) per individual, for the name-linked
        # generation-consistency check (a parent sits one generation above).
        for pr in person_rows:
            g = v(pr["row"], _C_GENERATION)
            if isinstance(g, float):
                diag["generation_by_id"].setdefault(
                    dedup_map[pr["dedup_key"]].id, g)

    return ordered, families
