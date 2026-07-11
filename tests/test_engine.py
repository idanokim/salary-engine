"""
Engine regression tests — synthetic data only (no real worker records).

Covers the pieces that broke or almost broke on real files:
  - vetek interpolation + per-track caps
  - the four-status slip classification
  - plus-grade resolution (dropped-'+' labels in old מנהלי dumps)
  - חוקה percent rules with per-file trust calibration

Run:  python -m pytest tests/ -q
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import main as engine


def lookups():
    return engine.get_lookups()


def make_worker(rows):
    """rows: list of (comp_code, comp_name, amount, pensionable) — one worker."""
    return rows


def combined_row(darga, vatek, track=1, pct=1.0, amount=None, kod=0,
                 code=engine.CODE_COMBINED_BASE):
    lk = lookups()
    if amount is None:
        gb = engine.get_grade_base(lk, darga)
        m = engine.get_vatek_multiplier(lk, vatek, track)
        if code == engine.CODE_COMBINED_BASE:
            amount = round(gb * m * pct, 2)
        elif code == engine.CODE_YESOD:
            amount = round(gb * pct, 2)
        else:
            amount = round(gb * (m - 1.0) * pct, 2)
    return (1, "משרד", track, pct, kod, darga, vatek, code, "בסיס", "כן", amount)


def run(workers_raw):
    return engine.run_engine_full(workers_raw, lookups())


# --- lookups ---------------------------------------------------------------

def test_vetek_interpolation_between_grid_points():
    lk = lookups()
    lo = engine.get_vatek_multiplier(lk, 10.0, 1)
    hi = engine.get_vatek_multiplier(lk, 10.25, 1)
    mid = engine.get_vatek_multiplier(lk, 10.1, 1)
    assert lo < mid < hi
    expected = lo + (hi - lo) * (10.1 - 10.0) / 0.25
    assert abs(mid - expected) < 1e-9


def test_vetek_capped_per_track():
    lk = lookups()
    for track, cap in lk["track_max"].items():
        at_cap = engine.get_vatek_multiplier(lk, cap, track)
        beyond = engine.get_vatek_multiplier(lk, cap + 5, track)
        assert at_cap == beyond


# --- classification ----------------------------------------------------------

def test_valid_slip():
    entries = run({1: [combined_row("18", 10.0)]})
    assert entries[0]["result"].status == "valid"


def test_invalid_slip():
    entries = run({1: [combined_row("18", 10.0, amount=9999.99)]})
    assert entries[0]["result"].status == "invalid"


def test_no_base_slip():
    entries = run({1: [combined_row("18", 10.0, amount=0.0)]})
    assert entries[0]["result"].status == "no_base"


def test_multi_period_slip():
    row = combined_row("18", 10.0)
    entries = run({1: [row, row]})
    assert entries[0]["result"].status == "multi_period"


def test_split_base_valid():
    """Pension-Authority form: יסוד משולב (1) + תוספת ותק (2)."""
    rows = [combined_row("42+", 40.5, track=11, code=engine.CODE_YESOD),
            combined_row("42+", 40.5, track=11, code=engine.CODE_VETEK_TOSEFET)]
    entries = run({1: rows})
    assert entries[0]["result"].status == "valid"


def test_seniority_rounding_window_accepts_off_grid_base():
    """A base computed off the true (unrounded) seniority must still validate
    when the file carries the rounded ותק column value."""
    lk = lookups()
    gb = engine.get_grade_base(lk, "18")
    true_mult = engine.get_vatek_multiplier(lk, 10.05, 1)  # true seniority 10.05
    slip = round(gb * true_mult, 2)
    entries = run({1: [combined_row("18", 10.0, amount=slip)]})  # rounded to 10.0
    assert entries[0]["result"].status == "valid"


# --- plus-grade resolution ---------------------------------------------------

def population(darga_paid, darga_stated, kod, n=20):
    """n active workers whose slip was paid at darga_paid but stated darga_stated."""
    paid_amount = combined_row(darga_paid, 10.0)[10]
    return {100 + i: [(1, "משרד", 1, 1.0, kod, darga_stated, 10.0,
                       engine.CODE_COMBINED_BASE, "שכר משולב", "כן", paid_amount)]
            for i in range(n)}


def test_dropped_plus_grade_is_resolved():
    workers = population("18+", "18", kod=202)
    remap = engine.resolve_plus_grades(workers, lookups())
    assert remap == {(202, "18"): "18+"}
    entries = run(workers)
    assert all(e["result"].status == "valid" for e in entries)
    assert all(e["result"].darga_label == "18+" for e in entries)


def test_plain_grade_population_is_not_remapped():
    workers = population("18", "18", kod=200)
    assert engine.resolve_plus_grades(workers, lookups()) == {}
    entries = run(workers)
    assert all(e["result"].status == "valid" for e in entries)


def test_labels_already_carrying_plus_pass_through():
    workers = {i: [combined_row("18+", 10.0, kod=202)] for i in range(20)}
    assert engine.resolve_plus_grades(workers, lookups()) == {}
    entries = run(workers)
    assert all(e["result"].status == "valid" for e in entries)


def test_small_or_split_vote_is_not_remapped():
    """Below PLUS_MIN_VOTES deciding slips → stay conservative, no remap."""
    workers = population("18+", "18", kod=202, n=engine.PLUS_MIN_VOTES - 1)
    assert engine.resolve_plus_grades(workers, lookups()) == {}


# --- חוקה percent rules -------------------------------------------------------

def test_percent_rule_trust_calibration():
    """A rule flags a bad slip only when it holds on ≥97% of carriers (≥20)."""
    rules = {4544: {"codes": [4544], "name": "3.6%", "type": "percent",
                    "rates": [0.036], "base_codes": [10002, 1, 2]}}
    base_amt = 5000.0
    all_checks = []
    for i in range(40):
        amount = round(base_amt * 0.036, 2) if i else 999.0  # worker 0 is wrong
        comps = [(10002, "שכר משולב", base_amt, "כן"), (4544, "3.6%", amount, "כן")]
        all_checks.append(engine.check_worker_components(comps, 1.0, rules))
    trusted = engine.trusted_rule_codes(all_checks)
    assert trusted == {4544}          # 39/40 ok ≥ 97%
    assert all_checks[0][4544]["ok"] is False
    assert all_checks[1][4544]["ok"] is True


def test_percent_rule_suppressed_when_rule_fails_file_wide():
    rules = {4544: {"codes": [4544], "name": "3.6%", "type": "percent",
                    "rates": [0.036], "base_codes": [10002, 1, 2]}}
    all_checks = []
    for i in range(30):
        amount = 111.11  # nobody matches → era/track variant → suppress
        comps = [(10002, "שכר משולב", 5000.0, "כן"), (4544, "3.6%", amount, "כן")]
        all_checks.append(engine.check_worker_components(comps, 1.0, rules))
    assert engine.trusted_rule_codes(all_checks) == set()


def test_2024_agreement_accepts_the_2pct_phase():
    rules = engine.get_rules()
    assert 0.02 in rules[5533]["rates"]
