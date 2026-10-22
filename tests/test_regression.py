"""Regression tests against the sample files (implementation plan section 9).

Run from the CODES folder:  python -m pytest tests -q
"""
import sys
from pathlib import Path

import pytest

CODES = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(CODES / "automation"))

from common.keys import metro_key, oracle_invoice_text, r12  # noqa: E402
from common.references import RefRegistry  # noqa: E402
from costco.engine import digits_region, is_deduction, run_costco  # noqa: E402
from config import load_params  # noqa: E402
from metro.engine import _within, run_metro  # noqa: E402

METRO_CSV = CODES / "CAN_RMT_METRODE-CA6-RMT_2252_202608271334.csv"
METRO_OI = CODES / "Open Item Reports" / "Open Item report Metro.xlsx"
COSTCO_PDF = CODES / "Costco Payment.PDF"
COSTCO_OI = CODES / "Open Item Reports" / "open item report costco.xlsx"


@pytest.fixture(scope="module")
def metro(tmp_path_factory):
    return run_metro(METRO_CSV, METRO_OI, tmp_path_factory.mktemp("out"))[0]


@pytest.fixture(scope="module")
def costco(tmp_path_factory):
    return run_costco(COSTCO_PDF, COSTCO_OI, tmp_path_factory.mktemp("out"))


def line(costco, invoice):
    return next(ln for ln in costco["lines"] if ln["invoice"] == invoice)


# ---- keys / references ----------------------------------------------------------
def test_keys():
    assert metro_key("026593522226") == metro_key(26593522226)
    assert r12("26866221906") == r12("026866221906") == "026866221906"
    assert r12("1261732219883") == "261732219883"
    assert oracle_invoice_text(26593522226) == "026593522226"


def test_reference_suffix():
    reg = RefRegistry({"SB INV-26921620976", "656-080326-SP"})
    assert reg.make("SB INV-26921620976") == ("SB INV-26921620976-1", True)
    assert reg.make("656-080326-SP")[0] == "656-080326-SP-1"
    assert reg.make("656-080326-SP")[0] == "656-080326-SP-2"


def test_write_off_boundaries():
    P = load_params()
    assert _within(10.00, P["METRO"]["WO_TOLERANCE"], True) and not _within(10.01, 10.0, True)
    assert _within(-11.00, P["COSTCO"]["WO_TOLERANCE"], True) and not _within(11.01, 11.0, True)


# ---- Metro -------------------------------------------------------------------------
def test_metro_totals(metro):
    assert metro["total"] == 164986.02
    assert metro["controls"].all_pass, metro["controls"].failed()


def test_metro_classification(metro):
    f = metro["formatted"]
    assert metro["status_counts"] == {"Matched": 1380, "Deduction": 41, "NF": 6, "SB": 1}
    assert round(f.loc[f.Status == "Matched", "Gross"].sum(), 2) == 736853.60
    ded = f[f.Category == "Deduction"]
    assert (ded.Net < 0).sum() == 36 and (ded.Net > 0).sum() == 5
    assert round(ded.Net.sum(), 2) == -572219.34
    assert f.loc[f.Invoice == "26921620976", "Reference"].item() == "SB INV-26921620976-1"


def test_metro_nac_and_receipt(metro):
    assert metro["nac_ref"] == "NAC0238539" and metro["nac"] == -571867.58
    assert round(metro["upload"]["Transaction Line Amount"].sum(), 2) == 0
    assert round(metro["receipt"]["Reference Amount"].sum(), 2) == 164986.02


# ---- Costco ------------------------------------------------------------------------
def test_costco_totals(costco):
    assert costco["total"] == 898467.62
    assert len(costco["lines"]) == 242
    assert costco["controls"].all_pass, costco["controls"].failed()


def test_costco_combine_and_match(costco):
    ln = line(costco, "226860021923")
    assert ln["paid"] == 250.32 and ln["status"] == "Matched"
    assert line(costco, "65024")["status"] == "Matched"


def test_costco_656_short_pay(costco):
    ln = line(costco, "226307921597")
    assert [d["trx"] for d in ln["added"]] == ["226307921596"]
    assert ln["status"] == "SP" and ln["L"] == 100.28 and ln["ref"].startswith("656-080326-SP")


def test_costco_698_two_invoices(costco):
    ln = line(costco, "226860221501")
    assert ln["status"] == "Resolved" and len(ln["added"]) == 2


def test_costco_write_offs(costco):
    assert line(costco, "226603721412")["status"] == "WO"
    assert line(costco, "226860621945")["status"] == "WO"
    assert line(costco, "026867021534")["status"] == "WO"


def test_costco_deductions(costco):
    rv = [ln for ln in costco["lines"] if ln["invoice"].startswith("RV")]
    assert rv and all(ln["region"] == "Ontario" and ln["status"] == "Deduction" for ln in rv)
    assert all(ln["upload"] == -ln["paid"] for ln in rv)            # never written off
    nf = line(costco, "6518087993")
    assert nf["ref"] == "NF-6518087993" and nf["upload"] == 43591.80
    assert line(costco, "1261732219883")["ref"] == "NF-261732219883"


def test_costco_regions(costco):
    assert digits_region("026866321582") == "Ontario"
    assert digits_region("226603721414") == "Western"
    assert line(costco, "226603721414")["status"] == "NF"
    assert is_deduction("0140801900", load_params()["COSTCO"])


def test_costco_upload_balances_per_region(costco):
    up = costco["upload"]
    for group, amt in up.groupby("NAC Group")["Transaction Line Amount"].sum().items():
        assert abs(amt) < 0.01, group
    assert "WO" not in set(up["Category"])
