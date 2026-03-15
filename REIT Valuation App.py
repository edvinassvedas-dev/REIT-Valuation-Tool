import FreeSimpleGUI as sg
from google.oauth2.service_account import Credentials
import gspread
from datetime import date

# ── Google Sheets setup ────────────────────────────────────────────────────────

CREDENTIALS_PATH = 'credentials.json' #Google service account credentials file
SPREADSHEET_NAME = 'DCF DB' #Replace with your spreadsheet file name
SHEET_NAME = 'REIT DB' #Replace with your sheet name

scope = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=scope)
gc = gspread.authorize(creds)


def get_worksheet():
    try:
        return gc.open(SPREADSHEET_NAME).worksheet(SHEET_NAME)
    except gspread.SpreadsheetNotFound:
        sg.popup_error(f"Spreadsheet '{SPREADSHEET_NAME}' not found!")
    except gspread.WorksheetNotFound:
        sg.popup_error(f"Sheet '{SHEET_NAME}' not found in '{SPREADSHEET_NAME}'!")
    except Exception as e:
        sg.popup_error(f"Error accessing spreadsheet: {e}")
    return None


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

HEADER = [
    "analysis_name", "shares", "market_price",
    "dps", "ddm_stage1_years",
    "ddm_worst_growth",    "ddm_worst_terminal",    "ddm_worst_rate",
    "ddm_base_growth",     "ddm_base_terminal",     "ddm_base_rate",
    "ddm_best_growth",     "ddm_best_terminal",     "ddm_best_rate",
    "affo", "affo_debt", "affo_cash", "affo_years",
    "affo_worst_growth",  "affo_worst_wacc",  "affo_worst_terminal",
    "affo_base_growth",   "affo_base_wacc",   "affo_base_terminal",
    "affo_best_growth",   "affo_best_wacc",   "affo_best_terminal",
    "gav", "nav_debt", "nav_other", "noi",
    "w_ddm", "w_affo", "w_nav",
    "notes", "analysis_date",
]


def load_database():
    worksheet = get_worksheet()
    if not worksheet:
        return [], []
    try:
        data = worksheet.get_all_values()
        if not data or len(data) < 2:
            return [], []
        header = data[0]
        database = [dict(zip(header, row)) for row in data[1:]]
        names = [a.get("analysis_name", "N/A") for a in database]
        return database, names
    except Exception as e:
        sg.popup_error(f"Error loading database: {e}")
        return [], []


def save_analysis(analysis_name, values):
    worksheet = get_worksheet()
    if not worksheet:
        return
    try:
        all_rows = worksheet.get_all_values()
        if not all_rows:
            worksheet.update(values=[HEADER], range_name="A1")
            next_row = 2
        else:
            if all_rows[0] != HEADER:
                worksheet.update(values=[HEADER], range_name="A1")
            next_row = len(all_rows) + 1

        def g(k): return values.get(k, "")

        row = [
            analysis_name,            g("-SHARES-"),            g("-MARKET_PRICE-"),
            g("-DPS-"),               g("-DDM_STAGE1_YEARS-"),
            g("-DDM_WORST_GROWTH-"),  g("-DDM_WORST_TERMINAL-"), g("-DDM_WORST_RATE-"),
            g("-DDM_BASE_GROWTH-"),   g("-DDM_BASE_TERMINAL-"),  g("-DDM_BASE_RATE-"),
            g("-DDM_BEST_GROWTH-"),   g("-DDM_BEST_TERMINAL-"),  g("-DDM_BEST_RATE-"),
            g("-AFFO-"),              g("-AFFO_DEBT-"),          g("-AFFO_CASH-"),   g("-AFFO_YEARS-"),
            g("-AFFO_WORST_GROWTH-"), g("-AFFO_WORST_WACC-"),   g("-AFFO_WORST_TERMINAL-"),
            g("-AFFO_BASE_GROWTH-"),  g("-AFFO_BASE_WACC-"),    g("-AFFO_BASE_TERMINAL-"),
            g("-AFFO_BEST_GROWTH-"),  g("-AFFO_BEST_WACC-"),    g("-AFFO_BEST_TERMINAL-"),
            g("-GAV-"),               g("-NAV_DEBT-"),           g("-NAV_OTHER-"),   g("-NOI-"),
            g("-W_DDM-"),             g("-W_AFFO-"),             g("-W_NAV-"),
            g("-NOTES-").strip(),
            date.today().strftime("%Y-%m-%d"),
        ]
        worksheet.add_rows(1)
        worksheet.update(values=[row], range_name=f"A{next_row}")
        sg.popup("Analysis saved successfully!")
    except Exception as e:
        sg.popup_error(f"Error saving: {e}")


def delete_analysis(analysis_name):
    worksheet = get_worksheet()
    if not worksheet:
        return False
    try:
        cell = worksheet.find(analysis_name)
        if cell:
            worksheet.delete_rows(cell.row)
            return True
        sg.popup_error(f"Analysis '{analysis_name}' not found.")
        return False
    except Exception as e:
        sg.popup_error(f"Error deleting analysis: {e}")
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


# ── Shared inputs ──
shared_left = [
    [sg.Text("REIT Valuation", font=("Helvetica", 16, "bold"), text_color="#0079d3")],
    [sg.Text("Analysis Name:",                size=(28, 1)), sg.InputText(key="-ANALYSIS_NAME-", size=(35, 1))],
    [sg.Text("Shares Outstanding (millions):", size=(28, 1)), sg.InputText("", key="-SHARES-",       size=(INP, 1)),
     sg.Text("  Market Price:",               size=(14, 1)), sg.InputText("", key="-MARKET_PRICE-", size=(INP, 1))],
]

# ── Notes ──
notes_col = [
    [sg.Text("Notes", font=("Helvetica", 11, "bold"))],
    [sg.Multiline(key="-NOTES-", size=(80, 4))],
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
    [sg.Button("Calculate")],
    [sg.Button("Save Analysis")],
    [sg.Button("Reset",      button_color=("white", "#999999"))],
]

# ── Weighting notes ──
weighting_notes_col = [
    [sg.Text("Weighting Guide", font=("Helvetica", 11, "bold"))],
    [sg.Text("Weights control the Wtd Avg in Summary.", font=("Helvetica", 9), text_color="#444444")],
    [sg.Text("Adjust based on REIT type:", font=("Helvetica", 9), text_color="#444444")],
    [sg.Text("")],
    [sg.Text("Property-heavy REITs", font=("Helvetica", 9, "bold"))],
    [sg.Text("NAV should carry more weight.", font=("Helvetica", 9), text_color="#444444")],
    [sg.Text("GAV reflects underlying asset value.", font=("Helvetica", 9), text_color="#444444")],
    [sg.Text("")],
    [sg.Text("Dividend-focused REITs", font=("Helvetica", 9, "bold"))],
    [sg.Text("DDM deserves more weight.", font=("Helvetica", 9), text_color="#444444")],
    [sg.Text("Dividend sustainability and growth", font=("Helvetica", 9), text_color="#444444")],
    [sg.Text("are the primary value drivers.", font=("Helvetica", 9), text_color="#444444")],
    [sg.Text("")],
    [sg.Text("Balanced approach", font=("Helvetica", 9, "bold"))],
    [sg.Text("Equal weights (33/34/33) work when", font=("Helvetica", 9), text_color="#444444")],
    [sg.Text("all three models are well supported.", font=("Helvetica", 9), text_color="#444444")],
]

# ── Saved Analyses ──
saved_col = [
    [sg.Text("Saved Analyses", font=("Helvetica", 11, "bold"))],
    [sg.Listbox(values=[], key="-ANALYSIS_LIST-", size=(70, 11), enable_events=True)],
    [
        sg.Button("Load Selected",   disabled=True, key="-LOAD_SELECTED-"),
        sg.Button("Delete Selected", disabled=True, key="-DELETE_SELECTED-"),
        sg.Button("Reload Database"),
    ],
]

layout = [
    [
        sg.Column(shared_left, vertical_alignment="top"),
        sg.VerticalSeparator(),
        sg.Column(notes_col, vertical_alignment="top", pad=((12, 0), 0)),
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
        sg.Column(weighting_notes_col, vertical_alignment="top", pad=((12, 12), 0)),
        sg.VerticalSeparator(),
        sg.Column(saved_col,          vertical_alignment="top", pad=((12, 0), 0)),
    ],
]

window = sg.Window("REIT Valuation", layout, size=(1300, 800), resizable=True, finalize=True)

loaded_database, analysis_names = load_database()
window["-ANALYSIS_LIST-"].update(values=analysis_names)
has_items = bool(analysis_names)
window["-LOAD_SELECTED-"].update(disabled=not has_items)
window["-DELETE_SELECTED-"].update(disabled=not has_items)

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
                        m = mos(price, market_price)
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
            exists = any(a.get("analysis_name") == name for a in loaded_database)
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
    elif event == "Reload Database":
        loaded_database, analysis_names = load_database()
        window["-ANALYSIS_LIST-"].update(values=analysis_names)
        has_items = bool(analysis_names)
        window["-LOAD_SELECTED-"].update(disabled=not has_items)
        window["-DELETE_SELECTED-"].update(disabled=not has_items)

    # ── List selection ─────────────────────────────────────────────────────────
    elif event == "-ANALYSIS_LIST-":
        selected = bool(values["-ANALYSIS_LIST-"])
        window["-LOAD_SELECTED-"].update(disabled=not selected)
        window["-DELETE_SELECTED-"].update(disabled=not selected)

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
                    sg.popup(f"'{sel_name}' deleted.")

window.close()