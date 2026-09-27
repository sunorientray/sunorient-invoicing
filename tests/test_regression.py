"""Regression test: the engine must reproduce real issued invoices to the cent."""
import os, sys, tempfile, warnings
warnings.filterwarnings("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "engine"))
from invoicegen import parse_kline_report, kline_lines, ecl_blocks, amount_in_words  # noqa: E402

F = lambda n: os.path.join(HERE, "fixtures", n)
tmp = tempfile.mkdtemp()
fails = []


def total(blocks):
    return str(sum(l["amount"] for b in blocks for l in b["lines"]))


# K Line 26-0901 Bishu Highway V134Z: weekday, no extras
rep = parse_kline_report(F("Bishu_Highway_V134Z_Confirmation_Report.xls"), tmp)
t = total(kline_lines(rep, {}))
if t != "17584.43": fails.append(f"Bishu 26-0901 expected 17584.43 got {t}")

# K Line 26-0906 American Highway V236A: Sunday 2nd/3rd shifts + MAFI stuffing block
rep = parse_kline_report(F("American_Highway_V236A_Confirmation_Report.xls"), tmp)
job = {"ph_shifts": ["2D1", "3D1"], "deduct": {"LASH_1": 1, "SUP_1": 1, "TM_OP": 1},
       "extra_blocks": [{"title": "MAFI", "items": [{"key": "LASH_1", "qty": 1}, {"key": "SUP_1", "qty": 1},
                                                    {"key": "TM_OP", "qty": 1}, {"key": "SIGNALMAN_MAFI", "qty": 1}]}]}
t = total(kline_lines(rep, job))
if t != "17904.76": fails.append(f"American 26-0906 expected 17904.76 got {t}")

# ECL Malaysia Grace V53 (hand-checked)
job = {"cargo": {"DIS": {"LT2": 761, "2TO5": 38, "GT5": 32, "GC": 4}, "LOD": {"GT5": 4}, "LNR": {"GT5": 1}},
       "supply": [{"shift": 2, "crew": {"SHIFTER": 2, "SIGNALLER": 1, "RAMP": 1, "WARDEN": 1, "LABOURER": 2}},
                  {"shift": 3, "crew": {"SHIFTER": 2, "SIGNALLER": 1, "RAMP": 1, "WARDEN": 1, "LABOURER": 2}}]}
t = total(ecl_blocks(job))
if t != "13981.80": fails.append(f"ECL Malaysia Grace expected 13981.80 got {t}")

w = amount_in_words("11124.50")
if w != "DOLLARS : ELEVEN THOUSAND ONE HUNDRED TWENTY FOUR AND CENTS FIFTY ONLY": fails.append("words: " + w)

print("ALL PASS" if not fails else "FAIL:\n  " + "\n  ".join(fails))
sys.exit(1 if fails else 0)
