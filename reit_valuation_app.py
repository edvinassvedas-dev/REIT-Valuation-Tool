import json
import os
import glob
import subprocess
import sys
import FreeSimpleGUI as sg
from datetime import date


# ── Local database setup ───────────────────────────────────────────────────────

DB_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reit_db")
os.makedirs(DB_DIR, exist_ok=True)


def _analysis_path(name: str) -> str:
    """Return the file path for a given analysis name."""
    safe = "".join(c if c.isalnum() or c in "-_ " else "_" for c in name)
    return os.path.join(DB_DIR, f"{safe}.json")


# ── Model logic ────────────────────────────────────────────────────────────────

def ddm_two_stage(dps, growth1, years1, growth2, discount_rate):
    """Two-stage DDM. Stage 1: growth1 for years1. Stage 2: growth2 perpetual."""
    if discount_rate <= growth2:
        raise ValueError("Discount rate must be greater than Stage 2 growth rate.")
    if years1 < 1:
        raise ValueError("Stage 1 years must be at least 1.")
    if dps <= 0:
        raise ValueError("Dividend per share must be greater than zero.")
    pv, d = 0.0, dps
    for t in range(1, int(years1) + 1):
        d *= (1 + growth1)
        pv += d / (1 + discount_rate) ** t
    tv = d * (1 + growth2) / (discount_rate - growth2)
    pv += tv / (1 + discount_rate) ** int(years1)
    return pv


def affo_dcf_calculate(affo, debt, cash, shares, years, growth_rate, wacc, terminal_growth):
    """AFFO-based DCF."""
    if wacc <= terminal_growth:
        raise ValueError("WACC must be greater than terminal growth rate.")
    if shares <= 0:
        raise ValueError("Shares outstanding must be greater than zero.")
    projected = [affo * (1 + growth_rate) ** i for i in range(1, years + 1)]
    discounted = [v / (1 + wacc) ** i for i, v in enumerate(projected, 1)]
    tv = projected[-1] * (1 + terminal_growth) / (wacc - terminal_growth)
    dtv = tv / (1 + wacc) ** years
    ev = sum(discounted) + dtv
    return (ev - debt + cash) / shares


def nav_calculate(gross_asset_value, total_debt, other_liabilities, shares):
    """NAV per share."""
    if shares <= 0:
        raise ValueError("Shares outstanding must be greater than zero.")
    return (gross_asset_value - total_debt - other_liabilities) / shares


def nav_from_cap_rate(noi, cap_rate, total_debt, other_liabilities, shares):
    """NAV per share derived from NOI / cap_rate as GAV."""
    if cap_rate <= 0:
        raise ValueError("Cap rate must be greater than zero.")
    gav = noi / cap_rate
    return nav_calculate(gav, total_debt, other_liabilities, shares)


def nav_sensitivity(gross_asset_value, total_debt, other_liabilities, shares):
    steps = [-0.20, -0.10, 0.0, 0.10, 0.20]
    rows, labels = [], []
    for s in steps:
        adj_gav = gross_asset_value * (1 + s)
        rows.append(nav_calculate(adj_gav, total_debt, other_liabilities, shares))
        labels.append(f"GAV {s*100:+.0f}%")
    return labels, rows


def mos(intrinsic, market):
    if intrinsic <= 0:
        return None
    return (intrinsic - market) / intrinsic * 100


def upside_pct(intrinsic, market):
    return (intrinsic - market) / market * 100


def weighted_avg(ddm_p, affo_p, nav_p, w_ddm, w_affo, w_nav):
    """Weighted average of available prices. Skips None values proportionally."""
    pairs = [(p, w) for p, w in [(ddm_p, w_ddm), (affo_p, w_affo), (nav_p, w_nav)]
             if p is not None]
    total_w = sum(w for _, w in pairs)
    if total_w == 0:
        return None
    return sum(p * w for p, w in pairs) / total_w


# ── Database helpers ───────────────────────────────────────────────────────────

_META = {
    "_meta": {
        "analysis_name":       "User-defined label for this analysis",
        "shares":              "Diluted shares outstanding in millions",
        "market_price":        "Current market price per share in dollars — optional, enables MoS and upside display",
        "dps":                 "Annual dividend per share in dollars",
        "ddm_stage1_years":    "Number of years for Stage 1 high-growth phase (typically 5–10)",
        "ddm_*_growth":        "Stage 1 dividend growth rate as a percentage (e.g. 4 means 4%)",
        "ddm_*_terminal":      "Stage 2 perpetual growth rate as a percentage — must be less than discount rate",
        "ddm_*_rate":          "Discount rate (required return) as a percentage",
        "affo":                "Adjusted Funds From Operations in millions",
        "affo_debt":           "Total debt in millions (for AFFO DCF equity bridge)",
        "affo_cash":           "Cash and equivalents in millions (for AFFO DCF equity bridge)",
        "affo_years":          "Projection horizon in years for AFFO DCF (typically 10)",
        "affo_*_growth":       "AFFO growth rate as a percentage",
        "affo_*_wacc":         "Weighted average cost of capital as a percentage — must exceed terminal growth",
        "affo_*_terminal":     "Terminal growth rate as a percentage",
        "gav":                 "Gross Asset Value in millions — total property portfolio at market value",
        "nav_debt":            "Total debt in millions (for NAV calculation)",
        "nav_other":           "Other liabilities in millions — use 0 if none",
        "noi":                 "Net Operating Income in millions — optional, enables cap rate sensitivity table",
        "w_ddm":               "Weight for DDM in weighted average — w_ddm + w_affo + w_nav should sum to 100",
        "w_affo":              "Weight for AFFO DCF in weighted average",
        "w_nav":               "Weight for NAV in weighted average",
        "notes":               "Free-text notes",
        "analysis_date":       "Date the analysis was saved — set automatically, format YYYY-MM-DD",
    }
}


def load_database():
    """Load all analyses from the reit_db directory. Keys starting with '_' are ignored."""
    files = sorted(glob.glob(os.path.join(DB_DIR, "*.json")))
    database = []
    for filepath in files:
        try:
            with open(filepath, "r", encoding="utf-8") as fh:
                raw = json.load(fh)
            # Strip meta/comment keys so they never reach the GUI
            record = {k: v for k, v in raw.items() if not k.startswith("_")}
            database.append(record)
        except Exception:
            pass  # skip corrupt or unreadable files silently
    names = [a.get("analysis_name", "N/A") for a in database]
    return database, names


def save_analysis(analysis_name, values):
    """Save a single analysis as an individual JSON file in reit_db/."""
    def g(k): return values.get(k, "")
    record = {
        "analysis_name":       analysis_name,
        "shares":              g("-SHARES-"),
        "market_price":        g("-MARKET_PRICE-"),
        "dps":                 g("-DPS-"),
        "ddm_stage1_years":    g("-DDM_STAGE1_YEARS-"),
        "ddm_worst_growth":    g("-DDM_WORST_GROWTH-"),
        "ddm_worst_terminal":  g("-DDM_WORST_TERMINAL-"),
        "ddm_worst_rate":      g("-DDM_WORST_RATE-"),
        "ddm_base_growth":     g("-DDM_BASE_GROWTH-"),
        "ddm_base_terminal":   g("-DDM_BASE_TERMINAL-"),
        "ddm_base_rate":       g("-DDM_BASE_RATE-"),
        "ddm_best_growth":     g("-DDM_BEST_GROWTH-"),
        "ddm_best_terminal":   g("-DDM_BEST_TERMINAL-"),
        "ddm_best_rate":       g("-DDM_BEST_RATE-"),
        "affo":                g("-AFFO-"),
        "affo_debt":           g("-AFFO_DEBT-"),
        "affo_cash":           g("-AFFO_CASH-"),
        "affo_years":          g("-AFFO_YEARS-"),
        "affo_worst_growth":   g("-AFFO_WORST_GROWTH-"),
        "affo_worst_wacc":     g("-AFFO_WORST_WACC-"),
        "affo_worst_terminal": g("-AFFO_WORST_TERMINAL-"),
        "affo_base_growth":    g("-AFFO_BASE_GROWTH-"),
        "affo_base_wacc":      g("-AFFO_BASE_WACC-"),
        "affo_base_terminal":  g("-AFFO_BASE_TERMINAL-"),
        "affo_best_growth":    g("-AFFO_BEST_GROWTH-"),
        "affo_best_wacc":      g("-AFFO_BEST_WACC-"),
        "affo_best_terminal":  g("-AFFO_BEST_TERMINAL-"),
        "gav":                 g("-GAV-"),
        "nav_debt":            g("-NAV_DEBT-"),
        "nav_other":           g("-NAV_OTHER-"),
        "noi":                 g("-NOI-"),
        "w_ddm":               g("-W_DDM-"),
        "w_affo":              g("-W_AFFO-"),
        "w_nav":               g("-W_NAV-"),
        "notes":               g("-NOTES-").strip(),
        "analysis_date":       date.today().strftime("%Y-%m-%d"),
        **_META,
    }
    try:
        with open(_analysis_path(analysis_name), "w", encoding="utf-8") as fh:
            json.dump(record, fh, indent=2)
        sg.popup("Analysis saved successfully!")
    except Exception as e:
        sg.popup_error(f"Error saving analysis: {e}")


def delete_analysis(analysis_name):
    """Delete the JSON file for a given analysis name."""
    path = _analysis_path(analysis_name)
    if os.path.exists(path):
        try:
            os.remove(path)
            return True
        except Exception as e:
            sg.popup_error(f"Error deleting analysis: {e}")
            return False
    sg.popup_error(f"Analysis '{analysis_name}' not found.")
    return False


# ── Default values ─────────────────────────────────────────────────────────────

DEFAULTS = {
    "-ANALYSIS_NAME-": "", "-SHARES-": "", "-MARKET_PRICE-": "",
    "-DPS-": "", "-DDM_STAGE1_YEARS-": "5",
    "-DDM_WORST_GROWTH-": "2",  "-DDM_WORST_TERMINAL-": "1.5", "-DDM_WORST_RATE-": "9",
    "-DDM_BASE_GROWTH-":  "4",  "-DDM_BASE_TERMINAL-":  "2",   "-DDM_BASE_RATE-":  "8",
    "-DDM_BEST_GROWTH-":  "6",  "-DDM_BEST_TERMINAL-":  "2.5", "-DDM_BEST_RATE-":  "7",
    "-AFFO-": "", "-AFFO_DEBT-": "", "-AFFO_CASH-": "", "-AFFO_YEARS-": "10",
    "-AFFO_WORST_GROWTH-": "1",   "-AFFO_WORST_WACC-": "9",   "-AFFO_WORST_TERMINAL-": "1.5",
    "-AFFO_BASE_GROWTH-":  "3",   "-AFFO_BASE_WACC-":  "8",   "-AFFO_BASE_TERMINAL-":  "2",
    "-AFFO_BEST_GROWTH-":  "5",   "-AFFO_BEST_WACC-":  "7",   "-AFFO_BEST_TERMINAL-":  "2.5",
    "-GAV-": "", "-NAV_DEBT-": "", "-NAV_OTHER-": "0", "-NOI-": "",
    "-W_DDM-": "33", "-W_AFFO-": "34", "-W_NAV-": "33",
    "-NOTES-": "",
}

RESULT_KEYS = [
    "-DDM_WORST_PRICE-", "-DDM_WORST_MOS-", "-DDM_WORST_UPSIDE-",
    "-DDM_BASE_PRICE-",  "-DDM_BASE_MOS-",  "-DDM_BASE_UPSIDE-",
    "-DDM_BEST_PRICE-",  "-DDM_BEST_MOS-",  "-DDM_BEST_UPSIDE-",
    "-DDM_COV-",
    "-AFFO_WORST_PRICE-", "-AFFO_WORST_MOS-", "-AFFO_WORST_UPSIDE-",
    "-AFFO_BASE_PRICE-",  "-AFFO_BASE_MOS-",  "-AFFO_BASE_UPSIDE-",
    "-AFFO_BEST_PRICE-",  "-AFFO_BEST_MOS-",  "-AFFO_BEST_UPSIDE-",
    "-NAV_PRICE-", "-NAV_PREMIUM-", "-NAV_CAP_RATE-",
    "-NAV_S1-", "-NAV_S1U-", "-NAV_S2-", "-NAV_S2U-",
    "-NAV_S3-", "-NAV_S3U-", "-NAV_S4-", "-NAV_S4U-",
    "-NAV_S5-", "-NAV_S5U-",
    "-CAP_S1-", "-CAP_S1U-", "-CAP_S2-", "-CAP_S2U-",
    "-CAP_S3-", "-CAP_S3U-", "-CAP_S4-", "-CAP_S4U-",
    "-CAP_S5-", "-CAP_S5U-", "-CAP_S6-", "-CAP_S6U-",
    "-CAP_S7-", "-CAP_S7U-",
    "-SUM_WORST_DDM-",  "-SUM_WORST_AFFO-",  "-SUM_WORST_NAV-",  "-SUM_WORST_WAVG-",
    "-SUM_BASE_DDM-",   "-SUM_BASE_AFFO-",   "-SUM_BASE_NAV-",   "-SUM_BASE_WAVG-",
    "-SUM_BEST_DDM-",   "-SUM_BEST_AFFO-",   "-SUM_BEST_NAV-",   "-SUM_BEST_WAVG-",
    "-SUM_WORST_DDM_U-","-SUM_WORST_AFFO_U-","-SUM_WORST_NAV_U-","-SUM_WORST_WAVG_U-",
    "-SUM_BASE_DDM_U-", "-SUM_BASE_AFFO_U-", "-SUM_BASE_NAV_U-", "-SUM_BASE_WAVG_U-",
    "-SUM_BEST_DDM_U-", "-SUM_BEST_AFFO_U-", "-SUM_BEST_NAV_U-", "-SUM_BEST_WAVG_U-",
]

CAP_RATE_STEPS = [0.035, 0.04, 0.045, 0.05, 0.055, 0.06, 0.065]
CAP_SENS_KEYS  = [
    ("-CAP_S1-", "-CAP_S1U-"), ("-CAP_S2-", "-CAP_S2U-"),
    ("-CAP_S3-", "-CAP_S3U-"), ("-CAP_S4-", "-CAP_S4U-"),
    ("-CAP_S5-", "-CAP_S5U-"), ("-CAP_S6-", "-CAP_S6U-"),
    ("-CAP_S7-", "-CAP_S7U-"),
]


# ── GUI layout ─────────────────────────────────────────────────────────────────

sg.theme("Reddit")

LBL = 22
INP = 10
INPpct = 5
RES = 10


def result_row(label, res_key, mos_key, upside_key):
    return [
        sg.Text(label, size=(10, 1)),
        sg.Text("—", key=res_key,    size=(RES, 1)),
        sg.Text("—", key=mos_key,    size=(8, 1)),
        sg.Text("—", key=upside_key, size=(8, 1)),
    ]


def col_header():
    return [
        sg.Text("",       size=(10, 1)),
        sg.Text("Price",  size=(RES, 1), font=("Helvetica", 9, "bold")),
        sg.Text("MoS",    size=(8,  1),  font=("Helvetica", 9, "bold")),
        sg.Text("Upside", size=(8,  1),  font=("Helvetica", 9, "bold")),
    ]


# ── Shared inputs (single top row) ──
shared_top = [
    [sg.Text("REIT Valuation", font=("Helvetica", 16, "bold"), text_color="#0079d3")],
    [sg.Text("Analysis Name:", size=(14, 1)),
     sg.InputText(key="-ANALYSIS_NAME-", size=(45, 1)),
     sg.VerticalSeparator(),
     sg.Text(" Shares Outstanding (M):", size=(20, 1)),
     sg.InputText("", key="-SHARES-", size=(INP, 1)),
     sg.Text("  Market Price:", size=(12, 1)),
     sg.InputText("", key="-MARKET_PRICE-", size=(INP, 1))],
]

# ── Notes ──
notes_col = [
    [sg.Text("Notes", font=("Helvetica", 11, "bold"))],
    [sg.Multiline(key="-NOTES-", size=(37, 15))],
]

# ── Model 1: Two-stage DDM ──
ddm_col = [
    [sg.Text("1 — Dividend Discount Model (DDM)", font=("Helvetica", 11, "bold"))],
    [sg.Text("Annual Dividend / Share:",  size=(LBL, 1)), sg.InputText("", key="-DPS-", size=(INP, 1))],
    [sg.Text("Stage 1 Years:",            size=(LBL, 1)), sg.InputText("5", key="-DDM_STAGE1_YEARS-", size=(INP, 1))],
    [sg.Text("", size=(LBL, 1)),
     sg.Text("Worst", size=(INPpct, 1), font=("Helvetica", 10, "bold"), text_color="#C0392B"),
     sg.Text("Base",  size=(INPpct, 1), font=("Helvetica", 10, "bold"), text_color="#7F8C8D"),
     sg.Text("Best",  size=(INPpct, 1), font=("Helvetica", 10, "bold"), text_color="#27AE60")],
    [sg.Text("Stage 1 Growth (%):",       size=(LBL, 1)),
     sg.InputText("2", key="-DDM_WORST_GROWTH-",    size=(INPpct, 1)),
     sg.InputText("4", key="-DDM_BASE_GROWTH-",     size=(INPpct, 1)),
     sg.InputText("6", key="-DDM_BEST_GROWTH-",     size=(INPpct, 1))],
    [sg.Text("Stage 2 Terminal Gr. (%):", size=(LBL, 1)),
     sg.InputText("1.5", key="-DDM_WORST_TERMINAL-", size=(INPpct, 1)),
     sg.InputText("2",   key="-DDM_BASE_TERMINAL-",  size=(INPpct, 1)),
     sg.InputText("2.5", key="-DDM_BEST_TERMINAL-",  size=(INPpct, 1))],
    [sg.Text("Discount Rate (%):",        size=(LBL, 1)),
     sg.InputText("9", key="-DDM_WORST_RATE-",      size=(INPpct, 1)),
     sg.InputText("8", key="-DDM_BASE_RATE-",       size=(INPpct, 1)),
     sg.InputText("7", key="-DDM_BEST_RATE-",       size=(INPpct, 1))],
    [sg.HorizontalSeparator()],
    col_header(),
    result_row("Worst:", "-DDM_WORST_PRICE-", "-DDM_WORST_MOS-", "-DDM_WORST_UPSIDE-"),
    result_row("Base:",  "-DDM_BASE_PRICE-",  "-DDM_BASE_MOS-",  "-DDM_BASE_UPSIDE-"),
    result_row("Best:",  "-DDM_BEST_PRICE-",  "-DDM_BEST_MOS-",  "-DDM_BEST_UPSIDE-"),
    [sg.HorizontalSeparator()],
    [sg.Text("Div. Coverage (AFFO/DPS):", size=(26, 1)),
     sg.Text("—", key="-DDM_COV-", size=(8, 1))],
    [sg.Text("Two-stage DDM: Stage 1 high growth, Stage 2 perpetual terminal growth",
             font=("Helvetica", 8), text_color="#888888")],
]

# ── Model 2: AFFO DCF ──
affo_col = [
    [sg.Text("2 — AFFO-Based DCF", font=("Helvetica", 11, "bold"))],
    [sg.Text("AFFO (millions):",       size=(LBL, 1)), sg.InputText("", key="-AFFO-",       size=(INP, 1))],
    [sg.Text("Total Debt (millions):", size=(LBL, 1)), sg.InputText("", key="-AFFO_DEBT-",  size=(INP, 1))],
    [sg.Text("Cash (millions):",       size=(LBL, 1)), sg.InputText("", key="-AFFO_CASH-",  size=(INP, 1))],
    [sg.Text("Years to Project:",      size=(LBL, 1)), sg.InputText("10", key="-AFFO_YEARS-", size=(INP, 1))],
    [sg.Text("", size=(LBL, 1)),
     sg.Text("Worst", size=(INPpct, 1), font=("Helvetica", 10, "bold"), text_color="#C0392B"),
     sg.Text("Base",  size=(INPpct, 1), font=("Helvetica", 10, "bold"), text_color="#7F8C8D"),
     sg.Text("Best",  size=(INPpct, 1), font=("Helvetica", 10, "bold"), text_color="#27AE60")],
    [sg.Text("AFFO Growth Rate (%):",  size=(LBL, 1)),
     sg.InputText("1",   key="-AFFO_WORST_GROWTH-",   size=(INPpct, 1)),
     sg.InputText("3",   key="-AFFO_BASE_GROWTH-",    size=(INPpct, 1)),
     sg.InputText("5",   key="-AFFO_BEST_GROWTH-",    size=(INPpct, 1))],
    [sg.Text("WACC (%):",              size=(LBL, 1)),
     sg.InputText("9",   key="-AFFO_WORST_WACC-",     size=(INPpct, 1)),
     sg.InputText("8",   key="-AFFO_BASE_WACC-",      size=(INPpct, 1)),
     sg.InputText("7",   key="-AFFO_BEST_WACC-",      size=(INPpct, 1))],
    [sg.Text("Terminal Growth (%):",   size=(LBL, 1)),
     sg.InputText("1.5", key="-AFFO_WORST_TERMINAL-", size=(INPpct, 1)),
     sg.InputText("2",   key="-AFFO_BASE_TERMINAL-",  size=(INPpct, 1)),
     sg.InputText("2.5", key="-AFFO_BEST_TERMINAL-",  size=(INPpct, 1))],
    [sg.HorizontalSeparator()],
    col_header(),
    result_row("Worst:", "-AFFO_WORST_PRICE-", "-AFFO_WORST_MOS-", "-AFFO_WORST_UPSIDE-"),
    result_row("Base:",  "-AFFO_BASE_PRICE-",  "-AFFO_BASE_MOS-",  "-AFFO_BASE_UPSIDE-"),
    result_row("Best:",  "-AFFO_BEST_PRICE-",  "-AFFO_BEST_MOS-",  "-AFFO_BEST_UPSIDE-"),
]

# ── Model 3: NAV — two internal sub-columns ──
_nav_left = [
    [sg.Text("Inputs", font=("Helvetica", 9, "bold"))],
    [sg.Text("Gross Asset Value (M):", size=(20, 1)), sg.InputText("", key="-GAV-",       size=(INP, 1))],
    [sg.Text("Total Debt (M):",        size=(20, 1)), sg.InputText("", key="-NAV_DEBT-",  size=(INP, 1))],
    [sg.Text("Other Liabilities (M):", size=(20, 1)), sg.InputText("0", key="-NAV_OTHER-", size=(INP, 1))],
    [sg.Text("NOI (M):",               size=(20, 1)), sg.InputText("", key="-NOI-",       size=(INP, 1))],
    [sg.HorizontalSeparator()],
    [sg.Text("NAV / Share:",      size=(16, 1)), sg.Text("—", key="-NAV_PRICE-",    size=(10, 1))],
    [sg.Text("Prem / Disc:",      size=(16, 1)), sg.Text("—", key="-NAV_PREMIUM-",  size=(10, 1))],
    [sg.Text("Implied Cap Rate:", size=(16, 1)), sg.Text("—", key="-NAV_CAP_RATE-", size=(10, 1))],
    [sg.HorizontalSeparator()],
    [sg.Text("GAV Sensitivity", font=("Helvetica", 10, "bold"))],
    [sg.Text("GAV ±%",    size=(10, 1), font=("Helvetica", 9, "bold")),
     sg.Text("NAV/Share", size=(10, 1), font=("Helvetica", 9, "bold")),
     sg.Text("vs Market", size=(10, 1), font=("Helvetica", 9, "bold"))],
    [sg.Text("GAV -20%", size=(10,1)), sg.Text("—", key="-NAV_S1-", size=(10,1)), sg.Text("—", key="-NAV_S1U-", size=(10,1))],
    [sg.Text("GAV -10%", size=(10,1)), sg.Text("—", key="-NAV_S2-", size=(10,1)), sg.Text("—", key="-NAV_S2U-", size=(10,1))],
    [sg.Text("GAV  0%",  size=(10,1)), sg.Text("—", key="-NAV_S3-", size=(10,1)), sg.Text("—", key="-NAV_S3U-", size=(10,1))],
    [sg.Text("GAV +10%", size=(10,1)), sg.Text("—", key="-NAV_S4-", size=(10,1)), sg.Text("—", key="-NAV_S4U-", size=(10,1))],
    [sg.Text("GAV +20%", size=(10,1)), sg.Text("—", key="-NAV_S5-", size=(10,1)), sg.Text("—", key="-NAV_S5U-", size=(10,1))],
]

_nav_right = [
    [sg.Text("Cap Rate Sensitivity", font=("Helvetica", 10, "bold"))],
    [sg.Text("Cap Rate",  size=(10, 1), font=("Helvetica", 9, "bold")),
     sg.Text("NAV/Share", size=(10, 1), font=("Helvetica", 9, "bold")),
     sg.Text("vs Market", size=(10, 1), font=("Helvetica", 9, "bold"))],
    [sg.Text("3.5%", size=(10,1)), sg.Text("—", key="-CAP_S1-", size=(10,1)), sg.Text("—", key="-CAP_S1U-", size=(10,1))],
    [sg.Text("4.0%", size=(10,1)), sg.Text("—", key="-CAP_S2-", size=(10,1)), sg.Text("—", key="-CAP_S2U-", size=(10,1))],
    [sg.Text("4.5%", size=(10,1)), sg.Text("—", key="-CAP_S3-", size=(10,1)), sg.Text("—", key="-CAP_S3U-", size=(10,1))],
    [sg.Text("5.0%", size=(10,1)), sg.Text("—", key="-CAP_S4-", size=(10,1)), sg.Text("—", key="-CAP_S4U-", size=(10,1))],
    [sg.Text("5.5%", size=(10,1)), sg.Text("—", key="-CAP_S5-", size=(10,1)), sg.Text("—", key="-CAP_S5U-", size=(10,1))],
    [sg.Text("6.0%", size=(10,1)), sg.Text("—", key="-CAP_S6-", size=(10,1)), sg.Text("—", key="-CAP_S6U-", size=(10,1))],
    [sg.Text("6.5%", size=(10,1)), sg.Text("—", key="-CAP_S7-", size=(10,1)), sg.Text("—", key="-CAP_S7U-", size=(10,1))],
    [sg.Text("Requires NOI input", font=("Helvetica", 8), text_color="#888888")],
]

nav_col = [
    [sg.Text("3 — Net Asset Value (NAV)", font=("Helvetica", 11, "bold"))],
    [
        sg.Column(_nav_left,  vertical_alignment="top"),
        sg.VerticalSeparator(),
        sg.Column(_nav_right, vertical_alignment="top", pad=((10, 0), 0)),
    ],
]

NAV_SENS_KEYS = [
    ("-NAV_S1-", "-NAV_S1U-"), ("-NAV_S2-", "-NAV_S2U-"), ("-NAV_S3-", "-NAV_S3U-"),
    ("-NAV_S4-", "-NAV_S4U-"), ("-NAV_S5-", "-NAV_S5U-"),
]

# ── Summary: models as rows, scenarios as columns ──
_SW = 12
_ML = 10


def sc_header_row():
    return [
        sg.Text("",      size=(_ML, 1)),
        sg.Text("Worst", size=(_SW, 1), font=("Helvetica", 9, "bold"), text_color="#C0392B"),
        sg.Text("Base",  size=(_SW, 1), font=("Helvetica", 9, "bold"), text_color="#7F8C8D"),
        sg.Text("Best",  size=(_SW, 1), font=("Helvetica", 9, "bold"), text_color="#27AE60"),
    ]


def model_price_row(label, kw, kb, kbest):
    return [
        sg.Text(label, size=(_ML, 1), font=("Helvetica", 9, "bold")),
        sg.Text("—", key=kw,    size=(_SW, 1)),
        sg.Text("—", key=kb,    size=(_SW, 1)),
        sg.Text("—", key=kbest, size=(_SW, 1)),
    ]


def model_upside_row(kw, kb, kbest):
    return [
        sg.Text("upside:", size=(_ML, 1), font=("Helvetica", 8), text_color="#888888"),
        sg.Text("—", key=kw,    size=(_SW, 1)),
        sg.Text("—", key=kb,    size=(_SW, 1)),
        sg.Text("—", key=kbest, size=(_SW, 1)),
    ]


summary_row_layout = [
    [sg.Text("Summary", font=("Helvetica", 11, "bold"))],
    # Weights row
    [sg.Text("Weights (%):", size=(_ML, 1), font=("Helvetica", 9, "bold")),
     sg.Text("DDM",  size=(6, 1), font=("Helvetica", 8)),
     sg.InputText("33", key="-W_DDM-",  size=(5, 1)),
     sg.Text("AFFO", size=(6, 1), font=("Helvetica", 8)),
     sg.InputText("34", key="-W_AFFO-", size=(5, 1)),
     sg.Text("NAV",  size=(5, 1), font=("Helvetica", 8)),
     sg.InputText("33", key="-W_NAV-",  size=(5, 1))],
    sc_header_row(),
    model_price_row("DDM:",      "-SUM_WORST_DDM-",  "-SUM_BASE_DDM-",  "-SUM_BEST_DDM-"),
    model_upside_row("-SUM_WORST_DDM_U-",  "-SUM_BASE_DDM_U-",  "-SUM_BEST_DDM_U-"),
    model_price_row("AFFO DCF:", "-SUM_WORST_AFFO-", "-SUM_BASE_AFFO-", "-SUM_BEST_AFFO-"),
    model_upside_row("-SUM_WORST_AFFO_U-", "-SUM_BASE_AFFO_U-", "-SUM_BEST_AFFO_U-"),
    model_price_row("NAV:",      "-SUM_WORST_NAV-",  "-SUM_BASE_NAV-",  "-SUM_BEST_NAV-"),
    model_upside_row("-SUM_WORST_NAV_U-",  "-SUM_BASE_NAV_U-",  "-SUM_BEST_NAV_U-"),
    [sg.HorizontalSeparator()],
    model_price_row("Wtd Avg:",  "-SUM_WORST_WAVG-", "-SUM_BASE_WAVG-", "-SUM_BEST_WAVG-"),
    model_upside_row("-SUM_WORST_WAVG_U-", "-SUM_BASE_WAVG_U-", "-SUM_BEST_WAVG_U-"),
]

# ── Action buttons ──
action_col = [
    [sg.Button("Calculate", button_color=("white", "#27AE60"))],
    [sg.Button("Save Analysis")],
    [sg.Button("Reset",      button_color=("white", "#999999"))],
]

# ── Saved Analyses ──
saved_col = [
    [sg.Text("Saved Analyses", font=("Helvetica", 11, "bold"))],
    [sg.Listbox(values=[], key="-ANALYSIS_LIST-", size=(50, 12), enable_events=True)],
    [
        sg.Button("Load Selected",   disabled=True, key="-LOAD_SELECTED-"),
        sg.Button("Delete Selected", button_color=("white", "#C0392B"), disabled=True, key="-DELETE_SELECTED-"), 
        sg.Button("Reload DB"),
    ],
    [
        sg.Button("Open File",       disabled=True, key="-OPEN_FILE-"),
        sg.Button("Open Folder",     key="-OPEN_DIR-"),
    ],
]

layout = [
    [
        sg.Column(shared_top, vertical_alignment="top"),
    ],
    [sg.HorizontalSeparator()],
    [
        sg.Column(ddm_col,  vertical_alignment="top"),
        sg.VerticalSeparator(),
        sg.Column(affo_col, vertical_alignment="top", pad=((12, 12), 0)),
        sg.VerticalSeparator(),
        sg.Column(nav_col,  vertical_alignment="top", pad=((12, 0), 0)),
    ],
    [sg.HorizontalSeparator()],
    [
        sg.Column(action_col,         vertical_alignment="top"),
        sg.VerticalSeparator(),
        sg.Column(summary_row_layout, vertical_alignment="top", pad=((12, 12), 0)),
        sg.VerticalSeparator(),
        sg.Column(notes_col,          vertical_alignment="top", pad=((12, 12), 0)),
        sg.VerticalSeparator(),
        sg.Column(saved_col,          vertical_alignment="top", pad=((12, 0), 0)),
    ],
]

window = sg.Window("REIT Valuation", layout, size=(1300, 790), resizable=True, finalize=True)

loaded_database, analysis_names = load_database()
window["-ANALYSIS_LIST-"].update(values=analysis_names)
has_items = bool(analysis_names)
window["-LOAD_SELECTED-"].update(disabled=not has_items)
window["-DELETE_SELECTED-"].update(disabled=not has_items)
window["-OPEN_FILE-"].update(disabled=not has_items)

# ── Event loop ─────────────────────────────────────────────────────────────────

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED:
        break

    # ── Reset ──────────────────────────────────────────────────────────────────
    if event == "Reset":
        for key, val in DEFAULTS.items():
            window[key].update(val)
        for key in RESULT_KEYS:
            window[key].update("—")
            try:
                window[key].update(text_color=sg.theme_text_color())
            except Exception:
                pass

    # ── Calculate ──────────────────────────────────────────────────────────────
    elif event == "Calculate":
        try:
            def fp(k): return float(values[k])

            mp_raw = values["-MARKET_PRICE-"].strip()
            market_price = float(mp_raw) if mp_raw else None
            shares       = fp("-SHARES-")
            stage1_years = int(fp("-DDM_STAGE1_YEARS-"))

            # Weights
            try:
                w_ddm  = fp("-W_DDM-")
                w_affo = fp("-W_AFFO-")
                w_nav  = fp("-W_NAV-")
            except Exception:
                w_ddm = w_affo = w_nav = 1.0

            # ── DDM ──
            dps = fp("-DPS-")
            ddm_scenarios = {
                "WORST": (fp("-DDM_WORST_GROWTH-") / 100, fp("-DDM_WORST_TERMINAL-") / 100, fp("-DDM_WORST_RATE-") / 100),
                "BASE":  (fp("-DDM_BASE_GROWTH-")  / 100, fp("-DDM_BASE_TERMINAL-")  / 100, fp("-DDM_BASE_RATE-")  / 100),
                "BEST":  (fp("-DDM_BEST_GROWTH-")  / 100, fp("-DDM_BEST_TERMINAL-")  / 100, fp("-DDM_BEST_RATE-")  / 100),
            }
            ddm_keys = {
                "WORST": ("-DDM_WORST_PRICE-", "-DDM_WORST_MOS-", "-DDM_WORST_UPSIDE-"),
                "BASE":  ("-DDM_BASE_PRICE-",  "-DDM_BASE_MOS-",  "-DDM_BASE_UPSIDE-"),
                "BEST":  ("-DDM_BEST_PRICE-",  "-DDM_BEST_MOS-",  "-DDM_BEST_UPSIDE-"),
            }
            ddm_worst_price = ddm_base_price = ddm_best_price = None
            for sc, (g1, g2, r) in ddm_scenarios.items():
                price = ddm_two_stage(dps, g1, stage1_years, g2, r)
                pk, mk, uk = ddm_keys[sc]
                window[pk].update(f"${price:.2f}")
                if sc == "WORST": ddm_worst_price = price
                elif sc == "BASE": ddm_base_price = price
                elif sc == "BEST": ddm_best_price = price
                if market_price:
                    m = mos(price, market_price)
                    u = upside_pct(price, market_price)
                    c = "green" if u > 0 else "red"
                    window[mk].update(f"{m:+.1f}%" if m is not None else "—", text_color=c)
                    window[uk].update(f"{u:+.1f}%", text_color=c)
                else:
                    window[mk].update("—"); window[uk].update("—")

            # Dividend coverage ratio: AFFO per share / DPS
            affo_raw = values["-AFFO-"].strip()
            if affo_raw and dps > 0:
                affo_ps = float(affo_raw) / shares
                cov = affo_ps / dps
                cov_color = "green" if cov >= 1.0 else "red"
                window["-DDM_COV-"].update(f"{cov:.2f}x", text_color=cov_color)
            else:
                window["-DDM_COV-"].update("—")

            # ── AFFO DCF ──
            affo  = fp("-AFFO-")
            debt  = fp("-AFFO_DEBT-")
            cash  = fp("-AFFO_CASH-")
            years = int(fp("-AFFO_YEARS-"))
            affo_scenarios = {
                "WORST": (fp("-AFFO_WORST_GROWTH-") / 100, fp("-AFFO_WORST_WACC-") / 100, fp("-AFFO_WORST_TERMINAL-") / 100),
                "BASE":  (fp("-AFFO_BASE_GROWTH-")  / 100, fp("-AFFO_BASE_WACC-")  / 100, fp("-AFFO_BASE_TERMINAL-")  / 100),
                "BEST":  (fp("-AFFO_BEST_GROWTH-")  / 100, fp("-AFFO_BEST_WACC-")  / 100, fp("-AFFO_BEST_TERMINAL-")  / 100),
            }
            affo_keys = {
                "WORST": ("-AFFO_WORST_PRICE-", "-AFFO_WORST_MOS-", "-AFFO_WORST_UPSIDE-"),
                "BASE":  ("-AFFO_BASE_PRICE-",  "-AFFO_BASE_MOS-",  "-AFFO_BASE_UPSIDE-"),
                "BEST":  ("-AFFO_BEST_PRICE-",  "-AFFO_BEST_MOS-",  "-AFFO_BEST_UPSIDE-"),
            }
            affo_worst_price = affo_base_price = affo_best_price = None
            for sc, (g, w, t) in affo_scenarios.items():
                price = affo_dcf_calculate(affo, debt, cash, shares, years, g, w, t)
                pk, mk, uk = affo_keys[sc]
                window[pk].update(f"${price:.2f}")
                if sc == "WORST": affo_worst_price = price
                elif sc == "BASE": affo_base_price = price
                elif sc == "BEST": affo_best_price = price
                if market_price:
                    m = mos(price, market_price)
                    u = upside_pct(price, market_price)
                    c = "green" if u > 0 else "red"
                    window[mk].update(f"{m:+.1f}%" if m is not None else "—", text_color=c)
                    window[uk].update(f"{u:+.1f}%", text_color=c)
                else:
                    window[mk].update("—"); window[uk].update("—")

            # ── NAV ──
            gav       = fp("-GAV-")
            nav_debt  = fp("-NAV_DEBT-")
            nav_other = fp("-NAV_OTHER-")
            nav_price = nav_calculate(gav, nav_debt, nav_other, shares)
            window["-NAV_PRICE-"].update(f"${nav_price:.2f}")

            noi_raw = values["-NOI-"].strip()
            if noi_raw:
                noi = float(noi_raw)
                cap_rate = noi / gav * 100
                window["-NAV_CAP_RATE-"].update(f"{cap_rate:.2f}%")
                # Cap rate sensitivity
                for cr, (nk, uk) in zip(CAP_RATE_STEPS, CAP_SENS_KEYS):
                    nav_cr = nav_from_cap_rate(noi, cr, nav_debt, nav_other, shares)
                    window[nk].update(f"${nav_cr:.2f}")
                    if market_price:
                        u = upside_pct(nav_cr, market_price)
                        c = "green" if u > 0 else "red"
                        window[uk].update(f"{u:+.1f}%", text_color=c)
                    else:
                        window[uk].update("—")
            else:
                window["-NAV_CAP_RATE-"].update("—")
                for _, (nk, uk) in zip(CAP_RATE_STEPS, CAP_SENS_KEYS):
                    window[nk].update("—"); window[uk].update("—")

            if market_price:
                premium = (market_price - nav_price) / nav_price * 100
                c = "red" if premium > 0 else "green"
                window["-NAV_PREMIUM-"].update(f"{premium:+.1f}%", text_color=c)
                _, navs = nav_sensitivity(gav, nav_debt, nav_other, shares)
                for (nk, uk), nav_val in zip(NAV_SENS_KEYS, navs):
                    u = upside_pct(nav_val, market_price)
                    window[nk].update(f"${nav_val:.2f}")
                    window[uk].update(f"{u:+.1f}%", text_color="green" if u > 0 else "red")
            else:
                window["-NAV_PREMIUM-"].update("—")
                _, navs = nav_sensitivity(gav, nav_debt, nav_other, shares)
                for (nk, uk), nav_val in zip(NAV_SENS_KEYS, navs):
                    window[nk].update(f"${nav_val:.2f}"); window[uk].update("—")

            # ── Summary + Weighted Average ──
            def update_sum_cell(pk, uk, price):
                if price is not None:
                    window[pk].update(f"${price:.2f}")
                    if market_price:
                        u = upside_pct(price, market_price)
                        c = "green" if u > 0 else "red"
                        window[uk].update(f"{u:+.1f}%", text_color=c)
                    else:
                        window[uk].update("—")
                else:
                    window[pk].update("—"); window[uk].update("—")

            for sc_label, ddm_p, affo_p in [
                ("WORST", ddm_worst_price, affo_worst_price),
                ("BASE",  ddm_base_price,  affo_base_price),
                ("BEST",  ddm_best_price,  affo_best_price),
            ]:
                nav_p  = nav_price
                wavg_p = weighted_avg(ddm_p, affo_p, nav_p, w_ddm, w_affo, w_nav)
                update_sum_cell(f"-SUM_{sc_label}_DDM-",  f"-SUM_{sc_label}_DDM_U-",  ddm_p)
                update_sum_cell(f"-SUM_{sc_label}_AFFO-", f"-SUM_{sc_label}_AFFO_U-", affo_p)
                update_sum_cell(f"-SUM_{sc_label}_NAV-",  f"-SUM_{sc_label}_NAV_U-",  nav_p)
                update_sum_cell(f"-SUM_{sc_label}_WAVG-", f"-SUM_{sc_label}_WAVG_U-", wavg_p)

        except ValueError as e:
            sg.popup_error(f"Input error: {e}")

    # ── Save ───────────────────────────────────────────────────────────────────
    elif event == "Save Analysis":
        name = values["-ANALYSIS_NAME-"].strip()
        if not name:
            sg.popup_error("Please enter an analysis name.")
        else:
            exists = os.path.exists(_analysis_path(name))
            if exists:
                new_name = sg.popup_get_text(
                    f"'{name}' already exists. Enter a new name:", title="Rename")
                if new_name and new_name != name:
                    values["-ANALYSIS_NAME-"] = new_name
                    save_analysis(new_name, values)
                    loaded_database, analysis_names = load_database()
                    window["-ANALYSIS_LIST-"].update(values=analysis_names)
            else:
                save_analysis(name, values)
                loaded_database, analysis_names = load_database()
                window["-ANALYSIS_LIST-"].update(values=analysis_names)

    # ── Reload ─────────────────────────────────────────────────────────────────
    elif event == "Reload DB":
        loaded_database, analysis_names = load_database()
        window["-ANALYSIS_LIST-"].update(values=analysis_names)
        has_items = bool(analysis_names)
        window["-LOAD_SELECTED-"].update(disabled=not has_items)
        window["-DELETE_SELECTED-"].update(disabled=not has_items)
        window["-OPEN_FILE-"].update(disabled=not has_items)

    # ── List selection ─────────────────────────────────────────────────────────
    elif event == "-ANALYSIS_LIST-":
        selected = bool(values["-ANALYSIS_LIST-"])
        window["-LOAD_SELECTED-"].update(disabled=not selected)
        window["-DELETE_SELECTED-"].update(disabled=not selected)
        window["-OPEN_FILE-"].update(disabled=not selected)

    # ── Load selected ──────────────────────────────────────────────────────────
    elif event == "-LOAD_SELECTED-":
        if values["-ANALYSIS_LIST-"]:
            sel_name = values["-ANALYSIS_LIST-"][0]
            for a in loaded_database:
                if a.get("analysis_name") == sel_name:
                    field_map = {
                        "-ANALYSIS_NAME-":        "analysis_name",
                        "-SHARES-":               "shares",
                        "-MARKET_PRICE-":         "market_price",
                        "-DPS-":                  "dps",
                        "-DDM_STAGE1_YEARS-":     "ddm_stage1_years",
                        "-DDM_WORST_GROWTH-":     "ddm_worst_growth",
                        "-DDM_WORST_TERMINAL-":   "ddm_worst_terminal",
                        "-DDM_WORST_RATE-":       "ddm_worst_rate",
                        "-DDM_BASE_GROWTH-":      "ddm_base_growth",
                        "-DDM_BASE_TERMINAL-":    "ddm_base_terminal",
                        "-DDM_BASE_RATE-":        "ddm_base_rate",
                        "-DDM_BEST_GROWTH-":      "ddm_best_growth",
                        "-DDM_BEST_TERMINAL-":    "ddm_best_terminal",
                        "-DDM_BEST_RATE-":        "ddm_best_rate",
                        "-AFFO-":                 "affo",
                        "-AFFO_DEBT-":            "affo_debt",
                        "-AFFO_CASH-":            "affo_cash",
                        "-AFFO_YEARS-":           "affo_years",
                        "-AFFO_WORST_GROWTH-":    "affo_worst_growth",
                        "-AFFO_WORST_WACC-":      "affo_worst_wacc",
                        "-AFFO_WORST_TERMINAL-":  "affo_worst_terminal",
                        "-AFFO_BASE_GROWTH-":     "affo_base_growth",
                        "-AFFO_BASE_WACC-":       "affo_base_wacc",
                        "-AFFO_BASE_TERMINAL-":   "affo_base_terminal",
                        "-AFFO_BEST_GROWTH-":     "affo_best_growth",
                        "-AFFO_BEST_WACC-":       "affo_best_wacc",
                        "-AFFO_BEST_TERMINAL-":   "affo_best_terminal",
                        "-GAV-":                  "gav",
                        "-NAV_DEBT-":             "nav_debt",
                        "-NAV_OTHER-":            "nav_other",
                        "-NOI-":                  "noi",
                        "-W_DDM-":                "w_ddm",
                        "-W_AFFO-":               "w_affo",
                        "-W_NAV-":                "w_nav",
                        "-NOTES-":                "notes",
                    }
                    for gui_key, db_key in field_map.items():
                        window[gui_key].update(a.get(db_key, ""))
                    sg.popup(f"'{sel_name}' loaded.")
                    break

    # ── Delete selected ────────────────────────────────────────────────────────
    elif event == "-DELETE_SELECTED-":
        if values["-ANALYSIS_LIST-"]:
            sel_name = values["-ANALYSIS_LIST-"][0]
            confirm = sg.popup_yes_no(
                f"Delete analysis '{sel_name}'?", title="Confirm Delete")
            if confirm == "Yes":
                if delete_analysis(sel_name):
                    loaded_database, analysis_names = load_database()
                    window["-ANALYSIS_LIST-"].update(values=analysis_names)
                    has_items = bool(analysis_names)
                    window["-LOAD_SELECTED-"].update(disabled=not has_items)
                    window["-DELETE_SELECTED-"].update(disabled=not has_items)
                    window["-OPEN_FILE-"].update(disabled=not has_items)
                    sg.popup(f"'{sel_name}' deleted.")

    # ── Open File ──────────────────────────────────────────────────────────────
    elif event == "-OPEN_FILE-":
        if values["-ANALYSIS_LIST-"]:
            sel_name = values["-ANALYSIS_LIST-"][0]
            path = _analysis_path(sel_name)
            if os.path.exists(path):
                try:
                    if sys.platform == "win32":
                        os.startfile(path)
                    elif sys.platform == "darwin":
                        subprocess.Popen(["open", path])
                    else:
                        subprocess.Popen(["xdg-open", path])
                except Exception as e:
                    sg.popup_error(f"Could not open file: {e}")
            else:
                sg.popup_error(f"File not found: {path}")
    # ── Open Folder ────────────────────────────────────────────────────────────
    elif event == "-OPEN_DIR-":
        try:
            if sys.platform == "win32":
                os.startfile(DB_DIR)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", DB_DIR])
            else:
                subprocess.Popen(["xdg-open", DB_DIR])
        except Exception as e:
            sg.popup_error(f"Could not open folder: {e}")

window.close()