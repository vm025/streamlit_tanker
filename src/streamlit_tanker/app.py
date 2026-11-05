"""
Water Tanker Tracker
--------------------
Streamlit app to log water tanker deliveries per block.

* TESTING = True  -> reads/writes  <project_root>/data/tanker.xlsx  (local machine)
* TESTING = False -> reads/writes  Google Sheet "tanker" (cloud), credentials from secrets.toml

Run locally :  uv run streamlit run src/streamlit_tanker/app.py   (from the project root)
"""

import hmac
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import streamlit as st

# ----------------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------------
def _read_testing_flag() -> bool:
    """TESTING comes from secrets.toml (bool or string). Missing -> True (local Excel, the default.)."""
    try:
        val = st.secrets.get("TESTING", True)
    except Exception:  # no secrets file found
        return True
    return val if isinstance(val, bool) else str(val).strip().lower() == "true"


TESTING = _read_testing_flag()  # set TESTING = false in secrets.toml on the cloud

BLOCKS = ["A-block", "B-block", "C-block", "D-block", "E-block", "F-block"]
HEADERS = ["serial_no.", "date_str", "time_str", "block", "quantity", "date_obj", "time_obj"]
MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
FORMULA_LAST_ROW = 100  # formulas in columns F & G are pre-filled up to this row
MAX_ATTEMPTS = 5

IST = timezone(timedelta(hours=5, minutes=30))  # IST has no DST, fixed offset is safe
# app.py lives in <root>/src/streamlit_tanker/, so the project root is 2 levels up
PROJECT_ROOT = Path(__file__).resolve().parents[2]
LOCAL_FILE = PROJECT_ROOT / "data" / "tanker.xlsx"
GSHEET_NAME = "tanker"


def date_formula(r: int) -> str:
    return f'=IF(B{r}="","",DATE(VALUE(RIGHT(B{r},4)),VALUE(MID(B{r},4,2)),VALUE(LEFT(B{r},2))))'


def time_formula(r: int) -> str:
    return f'=IF(C{r}="","",TIME(VALUE(LEFT(C{r},2)),VALUE(RIGHT(C{r},2)),0))'


def sheet_title(now: datetime) -> str:
    """e.g. 'nov-2026' (always lower-case; lookups are case-insensitive)."""
    return f"{MONTHS[now.month - 1]}-{now.year}"


# ----------------------------------------------------------------------------
# LOCAL EXCEL BACKEND (TESTING = True)
# ----------------------------------------------------------------------------
def _save_local(now: datetime, block: str, qty: int) -> int:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font

    LOCAL_FILE.parent.mkdir(parents=True, exist_ok=True)
    if LOCAL_FILE.exists():
        wb = load_workbook(LOCAL_FILE)
    else:
        wb = Workbook()
        wb.remove(wb.active)

    title = sheet_title(now)
    ws = next((s for s in wb.worksheets if s.title.lower() == title.lower()), None)

    if ws is None:  # new month -> new worksheet
        ws = wb.create_sheet(title)
        for c, h in enumerate(HEADERS, start=1):
            ws.cell(row=1, column=c, value=h).font = Font(bold=True)
        for r in range(2, FORMULA_LAST_ROW + 1):
            ws.cell(row=r, column=2).number_format = "@"
            ws.cell(row=r, column=3).number_format = "@"
            ws.cell(row=r, column=6, value=date_formula(r)).number_format = "dd-mm-yyyy"
            ws.cell(row=r, column=7, value=time_formula(r)).number_format = "hh:mm"
        for col, w in zip("ABCDEFG", (11, 12, 10, 10, 10, 12, 10)):
            ws.column_dimensions[col].width = w

    # first empty row based on column B (F & G already hold formulas)
    r = 1
    while ws.cell(row=r + 1, column=2).value not in (None, ""):
        r += 1
    row = r + 1
    serial = row - 1

    ws.cell(row=row, column=1, value=serial)
    ws.cell(row=row, column=2, value=now.strftime("%d-%m-%Y")).number_format = "@"
    ws.cell(row=row, column=3, value=now.strftime("%H:%M")).number_format = "@"
    ws.cell(row=row, column=4, value=block)
    ws.cell(row=row, column=5, value=qty)
    if row > FORMULA_LAST_ROW:  # beyond pre-filled rows
        ws.cell(row=row, column=6, value=date_formula(row)).number_format = "dd-mm-yyyy"
        ws.cell(row=row, column=7, value=time_formula(row)).number_format = "hh:mm"

    wb.save(LOCAL_FILE)
    return serial


# ----------------------------------------------------------------------------
# GOOGLE SHEETS BACKEND (TESTING = False)
# ----------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def _gs_spreadsheet():
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=scopes
    )
    client = gspread.authorize(creds)
    sheet_id = st.secrets.get("GSHEET_ID", "")
    return client.open_by_key(sheet_id) if sheet_id else client.open(GSHEET_NAME)


def _gs_init_worksheet(ws):
    ws.update(range_name="A1:G1", values=[HEADERS], value_input_option="RAW")
    ws.format("A1:G1", {"textFormat": {"bold": True}})
    ws.format("B2:C1000", {"numberFormat": {"type": "TEXT"}})
    ws.format("F2:F1000", {"numberFormat": {"type": "DATE", "pattern": "dd-mm-yyyy"}})
    ws.format("G2:G1000", {"numberFormat": {"type": "TIME", "pattern": "hh:mm"}})
    formulas = [[date_formula(r), time_formula(r)] for r in range(2, FORMULA_LAST_ROW + 1)]
    ws.update(
        range_name=f"F2:G{FORMULA_LAST_ROW}",
        values=formulas,
        value_input_option="USER_ENTERED",
    )


def _gs_get_or_create_ws(sh, title: str):
    import gspread

    def find():
        return next((w for w in sh.worksheets() if w.title.lower() == title.lower()), None)

    ws = find()
    if ws is None:
        try:
            ws = sh.add_worksheet(title=title, rows=1000, cols=len(HEADERS))
            _gs_init_worksheet(ws)
        except gspread.exceptions.APIError:
            ws = find()  # created concurrently by another request
            if ws is None:
                raise
    return ws


def _save_gsheet(now: datetime, block: str, qty: int) -> int:
    sh = _gs_spreadsheet()
    ws = _gs_get_or_create_ws(sh, sheet_title(now))

    row = len(ws.col_values(2)) + 1  # col B holds header + data only
    serial = row - 1

    ws.update(
        range_name=f"A{row}:E{row}",
        values=[[serial, now.strftime("%d-%m-%Y"), now.strftime("%H:%M"), block, qty]],
        value_input_option="RAW",  # keeps date/time as plain strings
    )
    if row > FORMULA_LAST_ROW:
        ws.update(
            range_name=f"F{row}:G{row}",
            values=[[date_formula(row), time_formula(row)]],
            value_input_option="USER_ENTERED",
        )
    return serial


def save_entry(block: str, qty: int):
    now = datetime.now(IST)
    serial = _save_local(now, block, qty) if TESTING else _save_gsheet(now, block, qty)
    return serial, now


# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------
st.set_page_config(page_title="Water Tanker Tracker", page_icon="💧", layout="centered")

st.markdown(
    """
    <style>
    .stApp { background: #FFF8B5; }
    .block-container { max-width: 520px; padding-top: 2rem; padding-bottom: 2rem; }
    header[data-testid="stHeader"] { background: transparent; }
    [data-testid="stAppViewContainer"], [data-testid="stMain"] { background: #FFF8B5; }
    #MainMenu, footer { visibility: hidden; }

    .app-title {
        text-align: center; color: #2F5D6B; font-weight: 700;
        font-size: 2rem; margin: 0; letter-spacing: .3px;
    }
    .app-sub { text-align: center; color: #5E7A80; font-size: .95rem; margin: .2rem 0 1.4rem 0; }

    [data-testid="stForm"] {
        background: #FFFFFF; border: 1px solid #DCE6E9; border-radius: 16px;
        padding: 1.4rem 1.2rem; box-shadow: 0 4px 14px rgba(47,93,107,.07);
    }
    label, .stRadio label p { color: #3E5A63 !important; font-weight: 500; }

    /* block selector card */
    .st-key-blocks {
        background: #FFFFFF; border: 1px solid #DCE6E9; border-radius: 16px;
        padding: 1.2rem 1.2rem .8rem 1.2rem; box-shadow: 0 4px 14px rgba(47,93,107,.07);
        margin-bottom: 1rem;
    }
    .block-label { color: #3E5A63; font-weight: 500; font-size: .95rem; margin-bottom: .4rem; }
    /* keep the 3 buttons per row even on phones */
    .st-key-blocks [data-testid="stHorizontalBlock"] { flex-direction: row !important; flex-wrap: wrap; gap: .6rem; }
    .st-key-blocks [data-testid="stColumn"] { min-width: calc(33% - .6rem) !important; flex: 1 1 calc(33% - .6rem) !important; }

    /* block buttons: dark teal */
    .st-key-blocks .stButton > button {
        background: #0F5C61; border: 2px solid #0F5C61; color: #fff;
        border-radius: 10px; padding: .6rem .4rem; font-weight: 600;
        box-shadow: 0 2px 6px rgba(15,92,97,.25);
    }
    .st-key-blocks .stButton > button p { color: #fff !important; font-weight: 600; }
    .st-key-blocks .stButton > button:hover { background: #0B4A4F; border-color: #0B4A4F; color: #fff; }
    /* selected block */
    .st-key-blocks .stButton > button[kind="primary"],
    .st-key-blocks .stButton > [data-testid="stBaseButton-primary"] {
        background: #07363A; border-color: #F2C94C; box-shadow: 0 0 0 3px rgba(242,201,76,.55);
    }

    .stTextInput input { border-radius: 10px; border: 1px solid #CBDADE; background: #FBFDFD; }
    .stTextInput input:focus { border-color: #4F8A9A; box-shadow: 0 0 0 1px #4F8A9A; }

    .stButton > button, [data-testid="stFormSubmitButton"] > button {
        width: 100%; background: #4F8A9A; color: #fff; border: none;
        border-radius: 10px; padding: .6rem 1rem; font-weight: 600;
    }
    .stButton > button:hover, [data-testid="stFormSubmitButton"] > button:hover {
        background: #3F7584; color: #fff; border: none;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<h1 class="app-title">💧 Water Tanker Tracker</h1>', unsafe_allow_html=True)
st.markdown('<p class="app-sub">Log tanker deliveries block-wise</p>', unsafe_allow_html=True)

# ---- session state
ss = st.session_state
ss.setdefault("authed", False)
ss.setdefault("fails", 0)
ss.setdefault("form_id", 0)
ss.setdefault("flash", None)

try:
    PASSCODE = str(st.secrets["PASSCODE"]).strip()
except Exception:
    st.error("PASSCODE is not configured in secrets.toml.")
    st.stop()


def passcode_ok(entered: str) -> bool:
    entered = entered.strip()
    if not (entered.isdigit() and 6 <= len(entered) <= 10):
        return False
    return hmac.compare_digest(entered.encode(), PASSCODE.encode())


# ---- Landing page: passcode gate
if not ss.authed:
    with st.form("passcode_form"):
        code = st.text_input(
            "Passcode",
            type="password",
            max_chars=10,
            placeholder="••••",  # looks like a 4-digit code on purpose
        )
        go = st.form_submit_button("Enter")

    if go:
        if ss.fails >= MAX_ATTEMPTS:
            st.error("Too many attempts. Please refresh the page and try again later.")
        elif passcode_ok(code):
            ss.authed = True
            ss.fails = 0
            st.rerun()
        else:
            ss.fails += 1
            time.sleep(1)  # slow down guessing
            st.error("Incorrect passcode.")
    st.stop()

# ---- Entry form
if ss.flash:
    st.success(ss.flash)
    ss.flash = None

fid = ss.form_id
ss.setdefault("block", None)


def pick_block(name: str):
    ss.block = name


with st.container(key="blocks"):
    st.markdown('<div class="block-label">Block</div>', unsafe_allow_html=True)
    for row in (BLOCKS[:3], BLOCKS[3:]):
        cols = st.columns(3)
        for col, name in zip(cols, row):
            col.button(
                name,
                key=f"blk_{name}",
                on_click=pick_block,
                args=(name,),
                type="primary" if ss.block == name else "secondary",
                use_container_width=True,
            )
    st.caption(f"Selected: **{ss.block}**" if ss.block else "Tap a block to select it.")

with st.form("entry_form"):
    qty_txt = st.text_input("Quantity", value="1", max_chars=6, key=f"qty_{fid}")
    submitted = st.form_submit_button("Submit")

if submitted:
    qty_txt = qty_txt.strip()
    block = ss.block
    if block is None:
        st.error("Please select a block.")
    elif not qty_txt:
        st.error("Please enter the quantity.")
    elif not qty_txt.isdigit() or int(qty_txt) < 1:
        st.error("Quantity must be a whole number of 1 or more.")
    else:
        try:
            with st.spinner("Saving..."):
                serial, now = save_entry(block, int(qty_txt))
            ss.flash = (
                f"✅ Saved #{serial}: {block} × {int(qty_txt)} "
                f"on {now.strftime('%d-%m-%Y')} at {now.strftime('%H:%M')} IST"
            )
            ss.form_id += 1  # resets the quantity field
            ss.block = None  # clears the block selection
            st.rerun()
        except Exception as e:
            st.error(f"Could not save the entry: {e}")

st.write("")
if st.button("Lock", key="lock_btn"):
    ss.authed = False
    st.rerun()