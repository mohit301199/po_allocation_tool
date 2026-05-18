import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, text
import bcrypt
from datetime import datetime
from io import BytesIO

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment

from barcode import Code128
from barcode.writer import ImageWriter

from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import PageBreak

import tempfile
import os


# =====================================================
# PAGE CONFIG
# =====================================================

st.set_page_config(
    page_title="PO Allocation",
    page_icon="package",
    layout="wide"
)


# =====================================================
# CSS
# =====================================================

st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg,#f8fafc,#eef2ff,#fff7ed);
    color: #0f172a;
}

.page-header {
    background: rgba(255,255,255,0.86);
    border: 1px solid rgba(148,163,184,0.35);
    border-left: 5px solid #2563eb;
    padding: 18px 22px;
    border-radius: 14px;
    margin-bottom: 22px;
    box-shadow: 0 10px 28px rgba(15,23,42,0.06);
}

.page-title {
    font-size: 28px;
    line-height: 1.2;
    font-weight: 900;
    color: #0f172a;
}

.page-subtitle {
    font-size: 14px;
    color: #475569;
    margin-top: 5px;
    font-weight: 600;
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg,#0f172a,#1e1b4b,#312e81);
}

[data-testid="stSidebar"] * {
    color: white !important;
}

div.stButton > button {
    background: linear-gradient(135deg,#2563eb,#7c3aed);
    color: white !important;
    border: none;
    border-radius: 14px;
    font-weight: 800;
    padding: 12px 22px;
    font-size: 15px;
}

div.stDownloadButton > button {
    background: linear-gradient(135deg,#2563eb,#7c3aed) !important;
    color: white !important;
    border: none !important;
    border-radius: 14px !important;
    font-weight: 800 !important;
    padding: 12px 22px !important;
    font-size: 15px !important;
}

div.stDownloadButton > button:hover {
    background: linear-gradient(135deg,#1d4ed8,#6d28d9) !important;
    color: white !important;
}

div[data-testid="stFileUploader"] {
    background: rgba(255,255,255,0.9);
    border: 2px dashed #2563eb;
    border-radius: 18px;
    padding: 18px;
}

div[data-testid="stFileUploader"] button {
    background: #2563eb !important;
    color: white !important;
    font-weight: 800 !important;
}

div[data-testid="stFileUploader"] button * {
    color: white !important;
}

label,
div[data-testid="stWidgetLabel"] p {
    color: #0f172a !important;
    font-weight: 800 !important;
    font-size: 16px !important;
}

.metric-card {
    padding: 24px;
    border-radius: 22px;
    color: white;
    text-align: center;
}

.blue { background: linear-gradient(135deg,#2563eb,#7c3aed); }
.green { background: linear-gradient(135deg,#059669,#22c55e); }
.orange { background: linear-gradient(135deg,#f97316,#f59e0b); }
.red { background: linear-gradient(135deg,#dc2626,#f43f5e); }

.metric-title {
    font-size: 15px;
    font-weight: 700;
}

.metric-value {
    font-size: 34px;
    font-weight: 900;
}

.success-box {
    background: #dcfce7;
    border-left: 6px solid #22c55e;
    padding: 15px;
    border-radius: 14px;
    color: #166534;
    font-weight: 800;
}
</style>
""", unsafe_allow_html=True)


def render_page_header(title, subtitle):
    st.markdown(f"""
    <div class="page-header">
        <div class="page-title">{title}</div>
        <div class="page-subtitle">{subtitle}</div>
    </div>
    """, unsafe_allow_html=True)


# =====================================================
# DATABASE
# =====================================================

@st.cache_resource
def get_engine():
    if "DATABASE_URL" not in st.secrets:
        st.error(
            "DATABASE_URL is missing. Please add your Supabase connection string "
            "in .streamlit/secrets.toml."
        )
        st.stop()

    return create_engine(st.secrets["DATABASE_URL"], pool_pre_ping=True)


engine = get_engine()


def make_params_key(params):
    if not params:
        return tuple()

    return tuple(sorted(params.items()))


@st.cache_data(ttl=20, show_spinner=False)
def db_read_cached(query, params_key):
    params = dict(params_key)

    with engine.connect() as connection:
        return pd.read_sql(text(query), connection, params=params)


def db_read(query, params=None, use_cache=True):
    params_key = make_params_key(params)

    if use_cache:
        return db_read_cached(query, params_key).copy()

    with engine.connect() as connection:
        return pd.read_sql(text(query), connection, params=params or {})


def db_execute(query, params=None, clear_cache=True):
    with engine.begin() as connection:
        connection.execute(text(query), params or {})

    if clear_cache:
        st.cache_data.clear()


# =====================================================
# AUTHENTICATION
# =====================================================

def hash_password(password):
    return bcrypt.hashpw(
        password.encode("utf-8"),
        bcrypt.gensalt()
    ).decode("utf-8")


def verify_password(password, password_hash):
    try:
        return bcrypt.checkpw(
            password.encode("utf-8"),
            password_hash.encode("utf-8")
        )
    except Exception:
        return False


def auth_clean_text(x):
    if x is None:
        return ""
    return str(x).replace("\xa0", " ").strip()


def get_user_count():
    result = db_read("SELECT COUNT(*) AS user_count FROM app_users")
    return int(result.iloc[0]["user_count"])


def log_activity(action, details=""):
    username = st.session_state.get("username", "")

    if not username:
        return

    db_execute("""
        INSERT INTO activity_log (username, action, details)
        VALUES (:username, :action, :details)
    """, {
        "username": username,
        "action": action,
        "details": details
    }, clear_cache=False)


def create_first_admin_screen():
    render_page_header(
        "Create Admin User",
        "Set up the first login for this cloud tracker."
    )

    st.info("No app users exist yet. Create the first Admin account to continue.")

    username = st.text_input("Admin Username")
    password = st.text_input("Admin Password", type="password")
    confirm_password = st.text_input("Confirm Password", type="password")

    if st.button("Create Admin"):
        username = auth_clean_text(username)

        if username == "":
            st.error("Please enter an admin username.")

        elif len(password) < 6:
            st.error("Password should be at least 6 characters.")

        elif password != confirm_password:
            st.error("Passwords do not match.")

        else:
            db_execute("""
                INSERT INTO app_users (username, password_hash, role, active)
                VALUES (:username, :password_hash, :role, TRUE)
            """, {
                "username": username,
                "password_hash": hash_password(password),
                "role": "Admin"
            })

            st.success("Admin user created. Please log in.")
            st.rerun()


def login_screen():
    render_page_header(
        "Login",
        "Sign in to access the PO allocation tracker."
    )

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        username = auth_clean_text(username)

        users = db_read("""
            SELECT username, password_hash, role
            FROM app_users
            WHERE username = :username
            AND active = TRUE
        """, {"username": username})

        if users.empty:
            st.error("Invalid username or password.")
            return

        user = users.iloc[0]

        if not verify_password(password, user["password_hash"]):
            st.error("Invalid username or password.")
            return

        st.session_state.logged_in = True
        st.session_state.username = user["username"]
        st.session_state.role = user["role"]
        log_activity("login", "User logged in")
        st.rerun()


def get_allowed_screens(role):
    if role == "Admin":
        return [
            "Dashboard Summary",
            "Upload & Allocate",
            "Allocation Tracker",
            "Billing Summary",
            "Open Allocation Qty",
            "User Management"
        ]

    if role == "Ops":
        return [
            "Dashboard Summary",
            "Upload & Allocate",
            "Allocation Tracker",
            "Billing Summary",
            "Open Allocation Qty"
        ]

    if role == "Billing":
        return [
            "Dashboard Summary",
            "Allocation Tracker",
            "Billing Summary",
            "Open Allocation Qty"
        ]

    return [
        "Dashboard Summary",
        "Billing Summary",
        "Open Allocation Qty"
    ]


if "logged_in" not in st.session_state:
    st.session_state.logged_in = False


if get_user_count() == 0:
    create_first_admin_screen()
    st.stop()


if not st.session_state.logged_in:
    login_screen()
    st.stop()


# =====================================================
# SESSION STATE
# =====================================================

if "stock_df" not in st.session_state:
    st.session_state.stock_df = None

if "pending_df" not in st.session_state:
    st.session_state.pending_df = None

if "allocation_df" not in st.session_state:
    st.session_state.allocation_df = None


# =====================================================
# HELPER FUNCTIONS
# =====================================================

def clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).replace("\xa0", " ").strip()


def clean_number(x):
    if pd.isna(x):
        return 0
    try:
        return float(str(x).replace(",", "").strip())
    except:
        return 0


def normalize_header(col):
    col = str(col)
    col = col.replace("\xa0", " ")
    col = col.replace("\n", " ")
    col = col.replace("\r", " ")
    col = col.replace("\t", " ")
    col = " ".join(col.split())
    return col.strip()


def normalize_columns(df):
    df = df.copy()
    df.columns = [normalize_header(c) for c in df.columns]
    return df


ALLOWED_STOCK_SITES = [
    "WBKOL02",
    "KABNG01",
    "TSHYD01",
    "HRAMB01",
    "MHBHW02",
    "PBTEP01",
    "HRAMB02",
    "HRAMB03",
    "MHBHW01",
    "KABNG02",
]

ALLOWED_STOCK_LOCATIONS = ["FGI", "ECOM"]


def first_existing_column(df, options):
    for option in options:
        if option in df.columns:
            return option
    return None


def standardize_stock_file(stock_df, show_messages=False):
    stock_df = normalize_columns(stock_df)

    old_format_cols = ["FSN", "RR Warehouse", "SAP Code", "Stock"]

    if all(col in stock_df.columns for col in old_format_cols):
        return stock_df

    item_code_col = first_existing_column(stock_df, ["Item Code", "Item code", "ItemCode"])
    available_qty_col = first_existing_column(stock_df, ["Available Qty", "Available Qty.", "AvailableQty"])
    site_col = first_existing_column(stock_df, ["Site", "SITE"])
    custom_location_col = first_existing_column(
        stock_df,
        ["Custom Location", "CustomLocation", "CUSTOM LOCATION", "CUSTOMLOCATION"]
    )
    fsn_col = first_existing_column(stock_df, ["FSN", "fsn"])

    system_stock_cols = [
        item_code_col,
        available_qty_col,
        site_col,
        custom_location_col,
        fsn_col,
    ]

    if not all(system_stock_cols):
        return stock_df

    original_rows = len(stock_df)

    stock_df = stock_df.copy()
    stock_df[item_code_col] = stock_df[item_code_col].apply(clean_text)
    stock_df[site_col] = stock_df[site_col].apply(clean_text).str.upper()
    stock_df[custom_location_col] = stock_df[custom_location_col].apply(clean_text).str.upper()
    stock_df[fsn_col] = stock_df[fsn_col].apply(clean_text)
    stock_df[available_qty_col] = stock_df[available_qty_col].apply(clean_number)

    stock_df = stock_df[
        stock_df[custom_location_col].isin(ALLOWED_STOCK_LOCATIONS) &
        stock_df[site_col].isin(ALLOWED_STOCK_SITES) &
        (stock_df[available_qty_col] > 0) &
        (stock_df[fsn_col] != "")
    ]

    standardized_df = stock_df.rename(columns={
        fsn_col: "FSN",
        site_col: "RR Warehouse",
        item_code_col: "SAP Code",
        available_qty_col: "Stock",
    })[["FSN", "RR Warehouse", "SAP Code", "Stock"]]

    standardized_df = standardized_df.groupby(
        ["FSN", "RR Warehouse", "SAP Code"],
        as_index=False
    )["Stock"].sum()

    if show_messages:
        st.info(
            "System stock file detected. "
            f"Filtered {original_rows:,} rows to {len(standardized_df):,} usable stock rows "
            "using CustomLocation FGI/ECOM and allowed SITE list."
        )

    return standardized_df


def standardize_pending_file(pending_df, show_messages=False):
    pending_df = normalize_columns(pending_df)

    old_format_cols = ["PO No.", "FSN", "Title", "RR Warehouse", "FK Warehouse", "Pending Qty."]

    if all(col in pending_df.columns for col in old_format_cols):
        if "Order ID" not in pending_df.columns:
            pending_df["Order ID"] = ""
        if "Buyer Code" not in pending_df.columns:
            pending_df["Buyer Code"] = ""
        if "Pending Amount" not in pending_df.columns:
            pending_df["Pending Amount"] = 0
        return pending_df

    po_col = first_existing_column(pending_df, ["PONo", "PO No.", "PO No", "PO Number", "PO"])
    order_id_col = first_existing_column(pending_df, ["Order Id", "Order ID", "OrderId"])
    buyer_code_col = first_existing_column(pending_df, ["Buyer Code", "BuyerCode"])
    fsn_col = first_existing_column(pending_df, ["FSN", "fsn"])
    title_col = first_existing_column(pending_df, ["Item Description", "Title", "Product Title"])
    rr_col = first_existing_column(pending_df, ["Site Id", "Site ID", "SITE", "Site"])
    fk_col = first_existing_column(pending_df, ["FK FC", "FK Warehouse", "FK WH"])
    pending_qty_col = first_existing_column(pending_df, ["Pending Qty", "Pending Qty.", "Pending Quantity"])
    pending_amount_col = first_existing_column(pending_df, ["Pending Amt", "Pending Amount", "Pending Amt."])
    status_col = first_existing_column(pending_df, ["Status", "Order Status"])

    system_pending_cols = [
        po_col,
        order_id_col,
        buyer_code_col,
        fsn_col,
        title_col,
        rr_col,
        fk_col,
        pending_qty_col,
        pending_amount_col,
        status_col,
    ]

    if not all(system_pending_cols):
        return pending_df

    original_rows = len(pending_df)

    pending_df = pending_df.copy()
    pending_df[status_col] = pending_df[status_col].apply(clean_text)
    pending_df = pending_df[
        pending_df[status_col].str.lower().isin(["hold", "pending"])
    ]

    for col in [po_col, order_id_col, buyer_code_col, fsn_col, title_col, rr_col, fk_col]:
        pending_df[col] = pending_df[col].apply(clean_text)

    pending_df[rr_col] = pending_df[rr_col].str.upper()
    pending_df[pending_qty_col] = pending_df[pending_qty_col].apply(clean_number)
    pending_df[pending_amount_col] = pending_df[pending_amount_col].apply(clean_number)

    blank_fsn_rows = len(pending_df[pending_df[fsn_col] == ""])

    pending_df = pending_df[
        (pending_df[fsn_col] != "") &
        (pending_df[pending_qty_col] > 0)
    ]

    standardized_df = pending_df.rename(columns={
        po_col: "PO No.",
        order_id_col: "Order ID",
        buyer_code_col: "Buyer Code",
        fsn_col: "FSN",
        title_col: "Title",
        rr_col: "RR Warehouse",
        fk_col: "FK Warehouse",
        pending_qty_col: "Pending Qty.",
        pending_amount_col: "Pending Amount",
    })[
        [
            "PO No.",
            "Order ID",
            "Buyer Code",
            "FSN",
            "Title",
            "RR Warehouse",
            "FK Warehouse",
            "Pending Qty.",
            "Pending Amount",
        ]
    ]

    standardized_df = standardized_df.groupby(
        ["PO No.", "Order ID", "Buyer Code", "FSN", "Title", "RR Warehouse", "FK Warehouse"],
        as_index=False
    ).agg({
        "Pending Qty.": "sum",
        "Pending Amount": "sum",
    })

    if show_messages:
        st.info(
            "System pending file detected. "
            f"Included Hold/Pending rows and converted {original_rows:,} source rows "
            f"to {len(standardized_df):,} usable pending rows. "
            f"Skipped {blank_fsn_rows:,} rows with blank FSN."
        )

    return standardized_df


def get_allocation_tracker_columns():
    columns_df = db_read("""
        SELECT column_name
        FROM information_schema.columns
        WHERE table_schema = 'public'
        AND table_name = 'allocation_tracker'
    """)

    return set(columns_df["column_name"].tolist())


def get_tracker_df():
    return db_read("SELECT * FROM allocation_tracker ORDER BY id DESC")


def get_open_allocation_qty():
    tracker_columns = get_allocation_tracker_columns()

    if "billed_qty" in tracker_columns:
        query = """
        SELECT
            fsn,
            rr_warehouse,
            sap_code,
            SUM(GREATEST(allocated_qty - COALESCE(billed_qty, 0), 0)) AS open_alloc_qty
        FROM allocation_tracker
        WHERE sent_for_billing='Yes'
        AND GREATEST(allocated_qty - COALESCE(billed_qty, 0), 0) > 0
        GROUP BY fsn, rr_warehouse, sap_code
        """

    else:
        query = """
        SELECT
            fsn,
            rr_warehouse,
            sap_code,
            SUM(allocated_qty) AS open_alloc_qty
        FROM allocation_tracker
        WHERE sent_for_billing='Yes'
        AND (billing_done IS NULL OR billing_done!='Yes')
        GROUP BY fsn, rr_warehouse, sap_code
        """

    return db_read(query)


def get_existing_po_allocation_qty():
    query = """
    SELECT
        po_no,
        fsn,
        rr_warehouse,
        fk_warehouse,
        SUM(allocated_qty) AS already_allocated_qty
    FROM allocation_tracker
    GROUP BY po_no, fsn, rr_warehouse, fk_warehouse
    """

    return db_read(query)


# =========================
# BILLING SUMMARY
# =========================

def get_billing_summary():
    tracker_columns = get_allocation_tracker_columns()
    qty_expr = "COALESCE(billed_qty, allocated_qty)" if "billed_qty" in tracker_columns else "allocated_qty"

    query = f"""
    SELECT
        invoice_no,
        po_no,
        fsn,
        title,
        rr_warehouse,
        fk_warehouse,
        sap_code,
        {qty_expr} AS allocated_qty
    FROM allocation_tracker
    WHERE sent_for_billing = 'Yes'
    AND invoice_no IS NOT NULL
    AND TRIM(invoice_no) <> ''
    ORDER BY invoice_no, po_no, fsn, sap_code
    """

    return db_read(query)


def get_sent_for_billing_download_df(tracker_df):
    if tracker_df.empty or "sent_for_billing" not in tracker_df.columns:
        return pd.DataFrame()

    export_df = tracker_df[
        tracker_df["sent_for_billing"].fillna("").str.strip().str.lower() == "yes"
    ].copy()

    if export_df.empty:
        return pd.DataFrame()

    export_columns = [
        "id",
        "sent_date",
        "po_no",
        "order_id",
        "buyer_code",
        "fsn",
        "title",
        "rr_warehouse",
        "fk_warehouse",
        "sap_code",
        "pending_amount",
        "allocated_qty",
        "billed_qty",
        "balance_to_bill",
        "invoice_no",
        "billing_date",
        "billing_done",
        "remark",
    ]

    export_columns = [col for col in export_columns if col in export_df.columns]
    export_df = export_df[export_columns]

    return export_df.rename(columns={
        "id": "Tracker ID",
        "sent_date": "Sent Date",
        "po_no": "PO No.",
        "order_id": "Order ID",
        "buyer_code": "Buyer Code",
        "fsn": "FSN",
        "title": "Title",
        "rr_warehouse": "RR Warehouse",
        "fk_warehouse": "FK Warehouse",
        "sap_code": "SAP Code",
        "pending_amount": "Pending Amount",
        "allocated_qty": "Allocated Qty.",
        "billed_qty": "Billed Qty.",
        "balance_to_bill": "Balance To Bill",
        "invoice_no": "Invoice No.",
        "billing_date": "Billing Date",
        "billing_done": "Billing Done",
        "remark": "Remark",
    })


# =====================================================
# FSN BARCODE GENERATION
# =====================================================


def generate_fsn_barcode_bytes(fsn):
    fsn = clean_text(fsn)

    if not fsn:
        return None

    output = BytesIO()
    barcode = Code128(fsn, writer=ImageWriter())
    barcode.write(
        output,
        options={
            "module_width": 0.32,
            "module_height": 18,
            "quiet_zone": 3,
            "write_text": False,
            "dpi": 300,
        }
    )
    output.seek(0)

    return output.getvalue()


def get_fsn_barcode_bytes(fsn, barcode_cache):
    fsn = clean_text(fsn)

    if not fsn:
        return None

    if fsn not in barcode_cache:
        barcode_cache[fsn] = generate_fsn_barcode_bytes(fsn)

    return barcode_cache[fsn]


def to_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    return output.getvalue()


def read_uploaded_excel(uploaded_file, file_label):
    try:
        return pd.read_excel(uploaded_file), ""

    except Exception as exc:
        return pd.DataFrame(), (
            f"Could not read the {file_label}. Please upload a normal Excel workbook "
            f"with at least one visible sheet. Original error: {exc}"
        )


def prepare_direct_billing_df(uploaded_df):
    direct_df = normalize_columns(uploaded_df)

    column_aliases = {
        "Invoice No.": ["Invoice No.", "Invoice No", "Invoice", "Invoice Number", "invoice_no"],
        "PO No.": ["PO No.", "PO No", "PO", "PO Number", "po_no"],
        "FSN": ["FSN", "fsn"],
        "Title": ["Title", "Product Title", "Item Title", "title"],
        "RR Warehouse": ["RR Warehouse", "RR WH", "RR Warehouse Code", "rr_warehouse"],
        "FK Warehouse": ["FK Warehouse", "FK WH", "FK Warehouse Code", "fk_warehouse"],
        "SAP Code": ["SAP Code", "SAP", "SAP SKU", "sap_code"],
        "Qty.": ["Qty.", "Qty", "Quantity", "Billing Qty", "Allocated Qty.", "allocated_qty"],
    }

    rename_map = {}

    for required_col, aliases in column_aliases.items():
        for alias in aliases:
            if alias in direct_df.columns:
                rename_map[alias] = required_col
                break

    direct_df = direct_df.rename(columns=rename_map)

    required_cols = list(column_aliases.keys())
    missing_cols = [col for col in required_cols if col not in direct_df.columns]

    if missing_cols:
        return pd.DataFrame(), missing_cols

    direct_df = direct_df[required_cols].copy()

    for col in [
        "Invoice No.",
        "PO No.",
        "FSN",
        "Title",
        "RR Warehouse",
        "FK Warehouse",
        "SAP Code"
    ]:
        direct_df[col] = direct_df[col].apply(clean_text)

    direct_df["Qty."] = direct_df["Qty."].apply(clean_number)
    direct_df = direct_df[direct_df["FSN"] != ""]

    return direct_df, []


# =====================================================
# BILLING SUMMARY WITH EXACT BARCODES
# =====================================================

def to_excel_billing_with_exact_barcodes(billing_df, fsn_barcode_cache):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Billing Summary"

    headers = [
        "Invoice No.",
        "PO No.",
        "FSN",
        "FSN Barcode",
        "Title",
        "RR Warehouse",
        "FK Warehouse",
        "SAP Code",
        "Qty."
    ]

    ws.append(headers)

    for col in range(1, len(headers) + 1):
        ws.cell(row=1, column=col).font = Font(
            bold=True,
            color="FFFFFF"
        )

        ws.cell(row=1, column=col).fill = PatternFill(
            "solid",
            fgColor="1F4E78"
        )

        ws.cell(row=1, column=col).alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 55
    ws.column_dimensions["E"].width = 45
    ws.column_dimensions["F"].width = 18
    ws.column_dimensions["G"].width = 22
    ws.column_dimensions["H"].width = 18
    ws.column_dimensions["I"].width = 12

    temp_files = []

    for idx, row in billing_df.iterrows():
        excel_row = idx + 2
        fsn = clean_text(row["FSN"])

        ws.cell(excel_row, 1).value = row["Invoice No."]
        ws.cell(excel_row, 2).value = row["PO No."]
        ws.cell(excel_row, 3).value = fsn
        ws.cell(excel_row, 5).value = row["Title"]
        ws.cell(excel_row, 6).value = row["RR Warehouse"]
        ws.cell(excel_row, 7).value = row["FK Warehouse"]
        ws.cell(excel_row, 8).value = row["SAP Code"]
        ws.cell(excel_row, 9).value = row["Qty."]

        ws.row_dimensions[excel_row].height = 115

        barcode_bytes = get_fsn_barcode_bytes(fsn, fsn_barcode_cache)

        if barcode_bytes:
            temp_file = tempfile.NamedTemporaryFile(
                delete=False,
                suffix=".png"
            )

            temp_file.write(barcode_bytes)
            temp_file.close()
            temp_files.append(temp_file.name)

            xl_img = XLImage(temp_file.name)
            xl_img.width = 340
            xl_img.height = 120

            ws.add_image(
                xl_img,
                f"D{excel_row}"
            )

        else:
            ws.cell(excel_row, 4).value = "Barcode Not Found"

        for col in range(1, 10):
            ws.cell(excel_row, col).alignment = Alignment(
                vertical="center",
                horizontal="center",
                wrap_text=True
            )

    wb.save(output)

    for file in temp_files:
        try:
            os.remove(file)
        except:
            pass

    output.seek(0)

    return output.getvalue()


# =====================================================
# STICKER PDF GENERATOR
# =====================================================

def generate_sticker_pdf(billing_df, fsn_barcode_cache):
    pdf_buffer = BytesIO()

    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=A4,
        rightMargin=8,
        leftMargin=8,
        topMargin=8,
        bottomMargin=8
    )

    elements = []
    temp_files = []

    for _, row in billing_df.iterrows():
        fsn = clean_text(row["FSN"])
        title = str(row["Title"])
        invoice_text = f"Invoice No. - {row['Invoice No.']}"
        po_text = f"PO No. - {row['PO No.']}"

        sticker_data = []

        for i in range(8):
            barcode_bytes = get_fsn_barcode_bytes(fsn, fsn_barcode_cache)

            if barcode_bytes:
                temp_file = tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=".png"
                )

                temp_file.write(barcode_bytes)
                temp_file.close()
                temp_files.append(temp_file.name)

                barcode_img = Image(
                    temp_file.name,
                    width=68 * mm,
                    height=22 * mm
                )

            else:
                barcode_img = Table(
                    [["Barcode Not Found"]],
                    colWidths=[75 * mm]
                )

            sticker_inner = Table(
                [
                    [title],
                    [barcode_img],
                    [fsn],
                    [invoice_text],
                    [po_text]
                ],
                colWidths=[82 * mm],
                rowHeights=[
                    10 * mm,
                    25 * mm,
                    7 * mm,
                    7 * mm,
                    7 * mm
                ]
            )

            sticker_inner.setStyle(TableStyle([
                ("BOX", (0, 0), (-1, -1), 1, colors.black),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTNAME", (0, 0), (0, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (0, 0), 7),
                ("FONTSIZE", (0, 2), (0, 4), 7),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]))

            sticker_data.append(sticker_inner)

        rows = []

        for i in range(0, len(sticker_data), 2):
            rows.append(sticker_data[i:i + 2])

        final_table = Table(
            rows,
            colWidths=[90 * mm, 90 * mm],
            rowHeights=[65 * mm] * 4
        )

        final_table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))

        elements.append(final_table)
        elements.append(PageBreak())

    if elements and isinstance(elements[-1], PageBreak):
        elements.pop()

    doc.build(elements)

    for file in temp_files:
        try:
            os.remove(file)
        except:
            pass

    pdf_buffer.seek(0)

    return pdf_buffer.getvalue()


# =====================================================
# ALLOCATION ENGINE
# =====================================================

def run_allocation(pending_df, stock_df):
    pending_df = standardize_pending_file(pending_df)
    stock_df = standardize_stock_file(stock_df)

    required_pending = [
        "PO No.", "FSN", "Title", "RR Warehouse", "FK Warehouse", "Pending Qty."
    ]

    required_stock = [
        "FSN", "RR Warehouse", "SAP Code", "Stock"
    ]

    missing_pending = [c for c in required_pending if c not in pending_df.columns]
    missing_stock = [c for c in required_stock if c not in stock_df.columns]

    if missing_pending:
        st.error(f"Missing columns in Pending Qty file: {missing_pending}")
        st.write("Pending file columns found:", list(pending_df.columns))
        return pd.DataFrame()

    if missing_stock:
        st.error(f"Missing columns in Stock file: {missing_stock}")
        st.write("Stock file columns found:", list(stock_df.columns))
        return pd.DataFrame()

    result = []

    for col in ["PO No.", "Order ID", "Buyer Code", "FSN", "Title", "RR Warehouse", "FK Warehouse"]:
        if col not in pending_df.columns:
            pending_df[col] = ""
        pending_df[col] = pending_df[col].apply(clean_text)

    pending_df["Pending Qty."] = pending_df["Pending Qty."].apply(clean_number)

    if "Pending Amount" not in pending_df.columns:
        pending_df["Pending Amount"] = 0

    pending_df["Pending Amount"] = pending_df["Pending Amount"].apply(clean_number)

    existing_po_alloc = get_existing_po_allocation_qty()

    if not existing_po_alloc.empty:
        existing_po_alloc["po_no"] = existing_po_alloc["po_no"].apply(clean_text)
        existing_po_alloc["fsn"] = existing_po_alloc["fsn"].apply(clean_text)
        existing_po_alloc["rr_warehouse"] = existing_po_alloc["rr_warehouse"].apply(clean_text)
        existing_po_alloc["fk_warehouse"] = existing_po_alloc["fk_warehouse"].apply(clean_text)

        pending_df = pending_df.merge(
            existing_po_alloc,
            how="left",
            left_on=["PO No.", "FSN", "RR Warehouse", "FK Warehouse"],
            right_on=["po_no", "fsn", "rr_warehouse", "fk_warehouse"]
        )

        pending_df["already_allocated_qty"] = pending_df["already_allocated_qty"].fillna(0)

    else:
        pending_df["already_allocated_qty"] = 0

    pending_df["Allocatable Pending Qty."] = (
        pending_df["Pending Qty."] - pending_df["already_allocated_qty"]
    )

    pending_df["Allocatable Pending Qty."] = pending_df["Allocatable Pending Qty."].apply(
        lambda x: max(x, 0)
    )

    for col in ["FSN", "RR Warehouse", "SAP Code"]:
        stock_df[col] = stock_df[col].apply(clean_text)

    stock_df["Stock"] = stock_df["Stock"].apply(clean_number)

    open_alloc = get_open_allocation_qty()

    if not open_alloc.empty:
        open_alloc["fsn"] = open_alloc["fsn"].apply(clean_text)
        open_alloc["rr_warehouse"] = open_alloc["rr_warehouse"].apply(clean_text)
        open_alloc["sap_code"] = open_alloc["sap_code"].apply(clean_text)

        stock_df = stock_df.merge(
            open_alloc,
            how="left",
            left_on=["FSN", "RR Warehouse", "SAP Code"],
            right_on=["fsn", "rr_warehouse", "sap_code"]
        )

        stock_df["open_alloc_qty"] = stock_df["open_alloc_qty"].fillna(0)

    else:
        stock_df["open_alloc_qty"] = 0

    stock_df["usable_stock"] = stock_df["Stock"] - stock_df["open_alloc_qty"]
    stock_df["usable_stock"] = stock_df["usable_stock"].apply(lambda x: max(x, 0))

    for _, p in pending_df.iterrows():
        po = clean_text(p["PO No."])
        order_id = clean_text(p["Order ID"])
        buyer_code = clean_text(p["Buyer Code"])
        fsn = clean_text(p["FSN"])
        title = clean_text(p["Title"])
        rr = clean_text(p["RR Warehouse"])
        fk = clean_text(p["FK Warehouse"])
        uploaded_pending_qty = clean_number(p["Pending Qty."])
        pending_amount = clean_number(p["Pending Amount"])
        already_allocated_qty = clean_number(p["already_allocated_qty"])
        pending_qty = clean_number(p["Allocatable Pending Qty."])
        pending_unit_amount = pending_amount / uploaded_pending_qty if uploaded_pending_qty > 0 else 0

        remaining = pending_qty
        allocated_anything = False

        if pending_qty <= 0:
            result.append({
                "PO No.": po,
                "Order ID": order_id,
                "Buyer Code": buyer_code,
                "FSN": fsn,
                "Title": title,
                "RR Warehouse": rr,
                "FK Warehouse": fk,
                "SAP Code": "",
                "Pending Qty.": pending_qty,
                "Pending Amount": 0,
                "Allocated Qty.": 0,
                "Balance Pending Qty.": 0,
                "Current Stock": 0,
                "Open Allocation Qty": 0,
                "Usable Stock Before": 0,
                "Usable Stock After": 0,
                "Status": "Already Allocated"
            })
            continue

        matching = stock_df[
            (stock_df["FSN"] == fsn) &
            (stock_df["RR Warehouse"] == rr) &
            (stock_df["usable_stock"] > 0)
        ]

        for idx, s in matching.iterrows():
            if remaining <= 0:
                break

            usable = clean_number(stock_df.loc[idx, "usable_stock"])
            alloc = min(usable, remaining)

            if alloc > 0:
                pending_qty_before_allocation = remaining
                pending_amount_before_allocation = pending_qty_before_allocation * pending_unit_amount

                allocated_anything = True
                stock_df.loc[idx, "usable_stock"] = usable - alloc
                remaining -= alloc

                result.append({
                    "PO No.": po,
                    "Order ID": order_id,
                    "Buyer Code": buyer_code,
                    "FSN": fsn,
                    "Title": title,
                    "RR Warehouse": rr,
                    "FK Warehouse": fk,
                    "SAP Code": s["SAP Code"],
                    "Pending Qty.": pending_qty_before_allocation,
                    "Pending Amount": pending_amount_before_allocation,
                    "Allocated Qty.": alloc,
                    "Balance Pending Qty.": remaining,
                    "Current Stock": s["Stock"],
                    "Open Allocation Qty": s["open_alloc_qty"],
                    "Usable Stock Before": usable,
                    "Usable Stock After": usable - alloc,
                    "Status": "Allocated"
                })

        if not allocated_anything:
            result.append({
                "PO No.": po,
                "Order ID": order_id,
                "Buyer Code": buyer_code,
                "FSN": fsn,
                "Title": title,
                "RR Warehouse": rr,
                "FK Warehouse": fk,
                "SAP Code": "",
                "Pending Qty.": pending_qty,
                "Pending Amount": pending_qty * pending_unit_amount,
                "Allocated Qty.": 0,
                "Balance Pending Qty.": remaining,
                "Current Stock": 0,
                "Open Allocation Qty": 0,
                "Usable Stock Before": 0,
                "Usable Stock After": 0,
                "Status": "No Stock"
            })

    return pd.DataFrame(result)


# =====================================================
# SAVE ALLOCATION TO TRACKER
# =====================================================

def save_allocation(df):
    saved_count = 0
    tracker_columns = get_allocation_tracker_columns()
    has_order_id = "order_id" in tracker_columns
    has_buyer_code = "buyer_code" in tracker_columns
    has_pending_amount = "pending_amount" in tracker_columns
    has_billed_qty = "billed_qty" in tracker_columns
    has_balance_to_bill = "balance_to_bill" in tracker_columns
    has_billing_source = "billing_source" in tracker_columns

    for _, r in df.iterrows():
        if clean_number(r["Allocated Qty."]) <= 0:
            continue

        insert_columns = [
            "allocation_date",
            "po_no",
            "fsn",
            "title",
            "rr_warehouse",
            "fk_warehouse",
            "sap_code",
            "allocated_qty",
            "sent_for_billing",
            "billing_done",
            "remark"
        ]

        value_keys = [
            ":allocation_date",
            ":po_no",
            ":fsn",
            ":title",
            ":rr_warehouse",
            ":fk_warehouse",
            ":sap_code",
            ":allocated_qty",
            ":sent_for_billing",
            ":billing_done",
            ":remark"
        ]

        params = {
            "allocation_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "po_no": r["PO No."],
            "fsn": r["FSN"],
            "title": r["Title"],
            "rr_warehouse": r["RR Warehouse"],
            "fk_warehouse": r["FK Warehouse"],
            "sap_code": r["SAP Code"],
            "allocated_qty": clean_number(r["Allocated Qty."]),
            "sent_for_billing": "No",
            "billing_done": "No",
            "remark": "Fresh Allocation"
        }

        if has_order_id:
            insert_columns.insert(2, "order_id")
            value_keys.insert(2, ":order_id")
            params["order_id"] = r.get("Order ID", "")

        if has_buyer_code:
            insert_at = 3 if has_order_id else 2
            insert_columns.insert(insert_at, "buyer_code")
            value_keys.insert(insert_at, ":buyer_code")
            params["buyer_code"] = r.get("Buyer Code", "")

        if has_pending_amount:
            insert_at = 9
            if has_order_id:
                insert_at += 1
            if has_buyer_code:
                insert_at += 1
            insert_columns.insert(insert_at, "pending_amount")
            value_keys.insert(insert_at, ":pending_amount")
            params["pending_amount"] = clean_number(r.get("Pending Amount", 0))

        if has_billed_qty:
            insert_columns.insert(-3, "billed_qty")
            value_keys.insert(-3, ":billed_qty")
            params["billed_qty"] = 0

        if has_balance_to_bill:
            insert_columns.insert(-3, "balance_to_bill")
            value_keys.insert(-3, ":balance_to_bill")
            params["balance_to_bill"] = clean_number(r["Allocated Qty."])

        if has_billing_source:
            insert_columns.insert(-1, "billing_source")
            value_keys.insert(-1, ":billing_source")
            params["billing_source"] = "Allocation"

        db_execute(f"""
        INSERT INTO allocation_tracker (
            {", ".join(insert_columns)}
        )
        VALUES (
            {", ".join(value_keys)}
        )
        """, params)

        saved_count += 1

    return saved_count


def reset_tracker_data():
    db_execute("DELETE FROM allocation_tracker")


# =====================================================
# SIDEBAR
# =====================================================

menu = st.sidebar.radio(
    "Select Screen",
    get_allowed_screens(st.session_state.role)
)

st.sidebar.markdown("---")
st.sidebar.write(f"Logged in as: {st.session_state.username}")
st.sidebar.write(f"Role: {st.session_state.role}")

if st.sidebar.button("Logout"):
    log_activity("logout", "User logged out")
    st.session_state.clear()
    st.rerun()


# =====================================================
# DASHBOARD
# =====================================================

if menu == "Dashboard Summary":
    render_page_header(
        "Dashboard Summary",
        "A quick view of allocation, billing, and open stock movement."
    )

    tracker = get_tracker_df()
    open_alloc = get_open_allocation_qty()

    total_alloc = tracker["allocated_qty"].sum() if not tracker.empty else 0
    sent_qty = tracker[tracker["sent_for_billing"] == "Yes"]["allocated_qty"].sum() if not tracker.empty else 0
    if not tracker.empty and "billed_qty" in tracker.columns:
        billed_qty = tracker["billed_qty"].fillna(0).sum()
    else:
        billed_qty = tracker[tracker["billing_done"] == "Yes"]["allocated_qty"].sum() if not tracker.empty else 0
    open_qty = open_alloc["open_alloc_qty"].sum() if not open_alloc.empty else 0

    c1, c2, c3, c4 = st.columns(4)

    c1.markdown(f"""
    <div class="metric-card blue">
        <div class="metric-title">Total Allocated</div>
        <div class="metric-value">{total_alloc:,.0f}</div>
    </div>
    """, unsafe_allow_html=True)

    c2.markdown(f"""
    <div class="metric-card orange">
        <div class="metric-title">Sent for Billing</div>
        <div class="metric-value">{sent_qty:,.0f}</div>
    </div>
    """, unsafe_allow_html=True)

    c3.markdown(f"""
    <div class="metric-card green">
        <div class="metric-title">Billed Qty</div>
        <div class="metric-value">{billed_qty:,.0f}</div>
    </div>
    """, unsafe_allow_html=True)

    c4.markdown(f"""
    <div class="metric-card red">
        <div class="metric-title">Open Allocation</div>
        <div class="metric-value">{open_qty:,.0f}</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    if not tracker.empty:
        st.subheader("Date-wise Billing Summary")

        summary_c1, summary_c2 = st.columns(2)

        with summary_c1:
            sent_summary = tracker.copy()
            sent_summary["sent_date"] = pd.to_datetime(
                sent_summary["sent_date"],
                errors="coerce"
            ).dt.date
            sent_summary = sent_summary[
                (sent_summary["sent_for_billing"] == "Yes") &
                (sent_summary["sent_date"].notna())
            ]

            if sent_summary.empty:
                st.info("No date-wise sent for billing data yet.")
            else:
                sent_summary = sent_summary.groupby(
                    "sent_date",
                    as_index=False
                )["allocated_qty"].sum().rename(columns={
                    "sent_date": "Sent Date",
                    "allocated_qty": "Sent Qty"
                })
                st.markdown("#### Sent for Billing")
                st.dataframe(sent_summary, use_container_width=True)

        with summary_c2:
            billed_summary = tracker.copy()
            billed_summary["billing_date"] = pd.to_datetime(
                billed_summary["billing_date"],
                errors="coerce"
            ).dt.date
            billed_summary = billed_summary[
                (billed_summary["billing_done"] == "Yes") &
                (billed_summary["billing_date"].notna())
            ]

            if billed_summary.empty:
                st.info("No date-wise billing done data yet.")
            else:
                billed_qty_column = "billed_qty" if "billed_qty" in billed_summary.columns else "allocated_qty"
                billed_summary[billed_qty_column] = billed_summary[billed_qty_column].fillna(0)
                billed_summary = billed_summary.groupby(
                    "billing_date",
                    as_index=False
                )[billed_qty_column].sum().rename(columns={
                    "billing_date": "Billing Date",
                    billed_qty_column: "Billed Qty"
                })
                st.markdown("#### Billing Done")
                st.dataframe(billed_summary, use_container_width=True)

        st.markdown("---")
        st.subheader("Recent Allocation Records")
        st.dataframe(tracker, use_container_width=True)
    else:
        st.info("No allocation records yet.")

    if st.session_state.role == "Admin":
        st.markdown("---")
        st.subheader("System Reset")

        with st.expander("Reset Tracker Data"):
            st.markdown("""
<div style="
background: linear-gradient(135deg,#fee2e2,#fecaca);
border-left: 6px solid #dc2626;
padding: 16px;
border-radius: 14px;
font-weight: 700;
color: #7f1d1d;
margin-top: 10px;
margin-bottom: 10px;
">
This will permanently delete all saved allocation tracker data.
</div>
""", unsafe_allow_html=True)
            confirm_reset = st.checkbox(
                "I understand and want to reset tracker data"
            )

            if confirm_reset:
                if st.button("Reset All Tracker Data"):
                    reset_tracker_data()
                    log_activity("reset_tracker", "All tracker data reset")
                    st.success("Tracker data reset successfully")
                    st.rerun()


# =====================================================
# UPLOAD & ALLOCATE
# =====================================================

elif menu == "Upload & Allocate":
    render_page_header(
        "Upload & Allocate",
        "Upload pending PO demand and stock sheets, then generate fresh allocation."
    )

    st.markdown("#### Required Upload Formats")

    sample1, sample2 = st.columns(2)

    with sample1:
        st.markdown("#### Pending Qty File Format")
        pending_sample = pd.DataFrame({
            "PO No.": ["PO123"],
            "Order ID": ["SO-123"],
            "Buyer Code": ["CUS001"],
            "FSN": ["FSN001"],
            "Title": ["Fan"],
            "RR Warehouse": ["WH1"],
            "FK Warehouse": ["FK1"],
            "Pending Qty.": [100],
            "Pending Amount": [25000]
        })
        st.dataframe(pending_sample, use_container_width=True)
        st.caption(
            "System pending files are also accepted. The app maps PONo, Order Id, "
            "Buyer Code, Item Description, Site Id, FK FC, Pending Qty and Pending Amt. "
            "Both Hold and Pending status rows are included."
        )

    with sample2:
        st.markdown("#### Stock File Format")
        stock_sample = pd.DataFrame({
            "FSN": ["FSN001"],
            "RR Warehouse": ["WH1"],
            "SAP Code": ["SAP001"],
            "Stock": [200]
        })
        st.dataframe(stock_sample, use_container_width=True)
        st.caption(
            "System stock files are also accepted. The app maps Item Code to SAP Code, "
            "Available Qty to Stock, Site to RR Warehouse, and filters CustomLocation "
            "to FGI/ECOM plus the approved SITE list."
        )

    st.markdown("---")
    st.subheader("Upload Files")

    c1, c2 = st.columns(2)

    with c1:
        stock_file = st.file_uploader(
            "Upload Stock File",
            type=["xlsx"],
            key="stock_file"
        )

    with c2:
        pending_file = st.file_uploader(
            "Upload Pending Qty File",
            type=["xlsx"],
            key="pending_file"
        )

    if stock_file is not None:
        st.session_state.stock_df = standardize_stock_file(
            pd.read_excel(stock_file),
            show_messages=True
        )

    if pending_file is not None:
        st.session_state.pending_df = standardize_pending_file(
            pd.read_excel(pending_file),
            show_messages=True
        )

    if st.session_state.stock_df is not None and st.session_state.pending_df is not None:
        st.markdown("""
        <div class="success-box">
        Files uploaded successfully
        </div>
        """, unsafe_allow_html=True)

        p1, p2 = st.columns(2)

        with p1:
            st.subheader("Stock Preview")
            st.write(f"Rows uploaded: {len(st.session_state.stock_df)}")
            st.dataframe(st.session_state.stock_df, use_container_width=True)

        with p2:
            st.subheader("Pending Qty Preview")
            st.write(f"Rows uploaded: {len(st.session_state.pending_df)}")
            st.dataframe(st.session_state.pending_df, use_container_width=True)

        st.markdown("---")

        if st.button("Run Allocation"):
            allocation = run_allocation(
                st.session_state.pending_df,
                st.session_state.stock_df
            )
            st.session_state.allocation_df = allocation

        if st.session_state.allocation_df is not None:
            st.subheader("Allocation Output")
            st.dataframe(st.session_state.allocation_df, use_container_width=True)

            b1, b2, b3 = st.columns(3)

            with b1:
                if st.button("Save Allocation to Tracker"):
                    saved = save_allocation(st.session_state.allocation_df)
                    st.success(f"{saved} allocation rows saved successfully")

            with b2:
                st.download_button(
                    "Download Allocation Output",
                    data=to_excel(st.session_state.allocation_df),
                    file_name="allocation_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with b3:
                if st.button("Clear Uploaded Files"):
                    st.session_state.stock_df = None
                    st.session_state.pending_df = None
                    st.session_state.allocation_df = None
                    st.rerun()


# =====================================================
# ALLOCATION TRACKER
# =====================================================

elif menu == "Allocation Tracker":
    render_page_header(
        "Allocation Tracker",
        "Update billing status, invoice details, dates, and remarks for saved allocations."
    )

    tracker = get_tracker_df()

    if tracker.empty:
        st.info("No allocation records found.")

    else:
        sent_for_billing_export = get_sent_for_billing_download_df(tracker)

        st.markdown("---")
        st.subheader("Sent For Billing Download")

        if sent_for_billing_export.empty:
            st.info("No rows are currently marked as Sent for Billing.")
        else:
            st.download_button(
                "Download Sent For Billing Rows",
                data=to_excel(sent_for_billing_export),
                file_name="sent_for_billing_rows.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        editable_tracker = tracker.copy()

        editable_tracker["sent_for_billing"] = editable_tracker["sent_for_billing"].fillna("No")
        editable_tracker["billing_done"] = editable_tracker["billing_done"].fillna("No")
        editable_tracker["invoice_no"] = editable_tracker["invoice_no"].fillna("")
        editable_tracker["remark"] = editable_tracker["remark"].fillna("")

        if "billed_qty" in editable_tracker.columns:
            editable_tracker["billed_qty"] = editable_tracker["billed_qty"].fillna(0)

        if "balance_to_bill" in editable_tracker.columns:
            editable_tracker["balance_to_bill"] = (
                editable_tracker["allocated_qty"].fillna(0) -
                editable_tracker.get("billed_qty", 0)
            ).apply(lambda x: max(clean_number(x), 0))

        editable_tracker["sent_for_billing_tick"] = editable_tracker["sent_for_billing"].apply(
            lambda x: True if str(x).strip().lower() == "yes" else False
        )

        editable_tracker["billing_done_tick"] = editable_tracker["billing_done"].apply(
            lambda x: True if str(x).strip().lower() == "yes" else False
        )

        editable_tracker["sent_date"] = pd.to_datetime(
            editable_tracker["sent_date"],
            errors="coerce"
        ).dt.date

        editable_tracker["billing_date"] = pd.to_datetime(
            editable_tracker["billing_date"],
            errors="coerce"
        ).dt.date

        editable_tracker = editable_tracker.drop(
            columns=["sent_for_billing", "billing_done"],
            errors="ignore"
        )

        tracker_display_columns = [
            "id",
            "allocation_date",
            "po_no",
            "order_id",
            "buyer_code",
            "fsn",
            "title",
            "rr_warehouse",
            "fk_warehouse",
            "sap_code",
            "pending_amount",
            "allocated_qty",
            "billed_qty",
            "balance_to_bill",
            "sent_date",
            "billing_date",
            "invoice_no",
            "remark",
            "sent_for_billing_tick",
            "billing_done_tick",
        ]

        tracker_display_columns = [
            col for col in tracker_display_columns if col in editable_tracker.columns
        ]

        editable_tracker = editable_tracker[tracker_display_columns]

        st.markdown("---")
        st.subheader("Bulk Update")

        select_all = st.checkbox("Select All Allocation Rows")

        if select_all:
            selected_ids = editable_tracker["id"].tolist()
            st.info(f"{len(selected_ids)} rows selected.")
        else:
            selected_ids = st.multiselect(
                "Select Allocation IDs",
                options=editable_tracker["id"].tolist()
            )

        b1, b2, b3, b4 = st.columns(4)

        with b1:
            bulk_sent_tick = st.checkbox("Mark Sent for Billing")

        with b2:
            bulk_sent_date = st.date_input("Sent Date")

        with b3:
            bulk_billing_tick = st.checkbox("Mark Billing Done")

        with b4:
            bulk_billing_date = st.date_input("Billing Date")

        b5, b6 = st.columns(2)

        with b5:
            bulk_invoice_no = st.text_input("Invoice No. Optional")

        with b6:
            bulk_remark = st.text_input("Remark Optional")

        if st.button("Apply Bulk Update"):
            if len(selected_ids) == 0:
                st.error("Please select at least one allocation row.")

            elif not bulk_sent_tick and not bulk_billing_tick and bulk_invoice_no == "" and bulk_remark == "":
                st.error("Please choose at least one update action.")

            else:
                for allocation_id in selected_ids:
                    if bulk_sent_tick:
                        db_execute("""
                        UPDATE allocation_tracker
                        SET
                            sent_for_billing = :sent_for_billing,
                            sent_date = :sent_date
                        WHERE id = :id
                        """, {
                            "sent_for_billing": "Yes",
                            "sent_date": str(bulk_sent_date),
                            "id": int(allocation_id)
                        })

                    if bulk_billing_tick:
                        if "billed_qty" in tracker.columns and "balance_to_bill" in tracker.columns:
                            db_execute("""
                            UPDATE allocation_tracker
                            SET
                                billing_done = :billing_done,
                                billing_date = :billing_date,
                                billed_qty = allocated_qty,
                                balance_to_bill = 0
                            WHERE id = :id
                            """, {
                                "billing_done": "Yes",
                                "billing_date": str(bulk_billing_date),
                                "id": int(allocation_id)
                            })

                        else:
                            db_execute("""
                            UPDATE allocation_tracker
                            SET
                                billing_done = :billing_done,
                                billing_date = :billing_date
                            WHERE id = :id
                            """, {
                                "billing_done": "Yes",
                                "billing_date": str(bulk_billing_date),
                                "id": int(allocation_id)
                            })

                    if bulk_invoice_no.strip() != "":
                        db_execute("""
                        UPDATE allocation_tracker
                        SET invoice_no = :invoice_no
                        WHERE id = :id
                        """, {
                            "invoice_no": bulk_invoice_no,
                            "id": int(allocation_id)
                        })

                    if bulk_remark.strip() != "":
                        db_execute("""
                        UPDATE allocation_tracker
                        SET remark = :remark
                        WHERE id = :id
                        """, {
                            "remark": bulk_remark,
                            "id": int(allocation_id)
                        })

                st.success("Bulk update applied successfully")
                st.rerun()

        st.markdown("---")

        edited_df = st.data_editor(
            editable_tracker,
            use_container_width=True,
            hide_index=True,
            column_config={
                "id": st.column_config.NumberColumn("ID", disabled=True),
                "allocation_date": st.column_config.TextColumn("Allocation Date", disabled=True),
                "po_no": st.column_config.TextColumn("PO No.", disabled=True),
                "order_id": st.column_config.TextColumn("Order ID", disabled=True),
                "buyer_code": st.column_config.TextColumn("Buyer Code", disabled=True),
                "fsn": st.column_config.TextColumn("FSN", disabled=True),
                "title": st.column_config.TextColumn("Title", disabled=True),
                "rr_warehouse": st.column_config.TextColumn("RR Warehouse", disabled=True),
                "fk_warehouse": st.column_config.TextColumn("FK Warehouse", disabled=True),
                "sap_code": st.column_config.TextColumn("SAP Code", disabled=True),
                "pending_amount": st.column_config.NumberColumn("Pending Amount", disabled=True),
                "allocated_qty": st.column_config.NumberColumn("Allocated Qty.", disabled=True),
                "billed_qty": st.column_config.NumberColumn("Billed Qty."),
                "balance_to_bill": st.column_config.NumberColumn("Balance To Bill", disabled=True),
                "sent_date": st.column_config.DateColumn("Sent Date", format="DD-MM-YYYY"),
                "billing_date": st.column_config.DateColumn("Billing Date", format="DD-MM-YYYY"),
                "invoice_no": st.column_config.TextColumn("Invoice No."),
                "remark": st.column_config.TextColumn("Remark"),
                "sent_for_billing_tick": st.column_config.CheckboxColumn("Sent for Billing?"),
                "billing_done_tick": st.column_config.CheckboxColumn("Billing Done?")
            }
        )

        if st.button("Save Manual Table Updates"):
            for _, row in edited_df.iterrows():
                sent_value = "Yes" if row["sent_for_billing_tick"] else "No"
                billed_value = "Yes" if row["billing_done_tick"] else "No"
                manual_billed_qty = clean_number(row.get("billed_qty", 0))
                manual_allocated_qty = clean_number(row.get("allocated_qty", 0))
                manual_balance_to_bill = max(manual_allocated_qty - manual_billed_qty, 0)

                if "billed_qty" in edited_df.columns:
                    billed_value = "Yes" if manual_balance_to_bill <= 0 and manual_allocated_qty > 0 else "No"

                db_execute("""
                UPDATE allocation_tracker
                SET
                    sent_for_billing = :sent_for_billing,
                    sent_date = :sent_date,
                    billing_done = :billing_done,
                    billing_date = :billing_date,
                    invoice_no = :invoice_no,
                    remark = :remark
                WHERE id = :id
                """, {
                    "sent_for_billing": sent_value,
                    "sent_date": str(row["sent_date"]) if pd.notna(row["sent_date"]) else "",
                    "billing_done": billed_value,
                    "billing_date": str(row["billing_date"]) if pd.notna(row["billing_date"]) else "",
                    "invoice_no": row["invoice_no"],
                    "remark": row["remark"],
                    "id": int(row["id"])
                })

                if "billed_qty" in edited_df.columns and "balance_to_bill" in tracker.columns:
                    db_execute("""
                    UPDATE allocation_tracker
                    SET
                        billed_qty = :billed_qty,
                        balance_to_bill = :balance_to_bill
                    WHERE id = :id
                    """, {
                        "billed_qty": manual_billed_qty,
                        "balance_to_bill": manual_balance_to_bill,
                        "id": int(row["id"])
                    })

            st.success("Manual tracker updates saved successfully")
            st.rerun()


# =====================================================
# BILLING SUMMARY
# =====================================================

elif menu == "Billing Summary":
    render_page_header(
        "Billing Summary",
        "Generate FSN barcode Excel and sticker PDFs from tracker data or urgent direct uploads."
    )

    fsn_barcode_cache = {}
    st.info("Barcode images are generated automatically from FSN. No barcode image upload is required.")

    st.markdown("### Urgent Direct Barcode Download")
    st.caption(
        "Use this when billing needs to happen quickly without updating allocation tracker rows first."
    )

    direct_sample = pd.DataFrame({
        "Invoice No.": ["INV001"],
        "PO No.": ["PO1001"],
        "FSN": ["FSN001"],
        "Title": ["Ceiling Fan 1200mm"],
        "RR Warehouse": ["WH1"],
        "FK Warehouse": ["FK-BLR"],
        "SAP Code": ["SAP-A1"],
        "Qty.": [10]
    })

    with st.expander("Direct billing upload format"):
        st.dataframe(direct_sample, use_container_width=True)
        st.download_button(
            "Download Direct Billing Template",
            data=to_excel(direct_sample),
            file_name="direct_billing_upload_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    direct_billing_file = st.file_uploader(
        "Upload Direct Billing Details File",
        type=["xlsx"],
        key="direct_billing_file"
    )

    if direct_billing_file is not None:
        try:
            direct_uploaded_df = pd.read_excel(direct_billing_file)
        except Exception as e:
            direct_uploaded_df = None
            st.error(
                "Could not read the Direct Billing Details file. "
                "Please upload a normal Excel workbook with at least one visible sheet. "
                f"Original error: {e}"
            )

        if direct_uploaded_df is not None:
            direct_billing_df, missing_direct_cols = prepare_direct_billing_df(
                direct_uploaded_df
            )

            if missing_direct_cols:
                st.error(f"Missing columns in Direct Billing file: {missing_direct_cols}")

            elif direct_billing_df.empty:
                st.warning("No valid FSN rows found in the direct billing file.")

            else:
                d1, d2, d3 = st.columns(3)
                d1.metric("Rows Uploaded", len(direct_billing_df))
                d2.metric("Total Qty", f"{direct_billing_df['Qty.'].sum():,.0f}")
                d3.metric("Generated Barcodes", direct_billing_df["FSN"].nunique())

                st.dataframe(direct_billing_df, use_container_width=True)

                q1, q2 = st.columns(2)

                with q1:
                    st.download_button(
                        "Download Direct Barcode Excel",
                        data=to_excel_billing_with_exact_barcodes(
                            direct_billing_df,
                            fsn_barcode_cache
                        ),
                        file_name="direct_billing_barcode_sheet.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                with q2:
                    st.download_button(
                        "Download Direct Sticker PDF",
                        data=generate_sticker_pdf(
                            direct_billing_df,
                            fsn_barcode_cache
                        ),
                        file_name="direct_warehouse_sticker_sheet.pdf",
                        mime="application/pdf"
                    )

    st.markdown("---")
    st.markdown("### Tracker Billing Workflow")

    billing_df = get_billing_summary()

    if billing_df.empty:
        st.info(
            "No rows found. Mark 'Sent for Billing' as Yes and enter Invoice No. in Allocation Tracker."
        )

    else:
        billing_df = billing_df.rename(columns={
            "invoice_no": "Invoice No.",
            "po_no": "PO No.",
            "fsn": "FSN",
            "title": "Title",
            "rr_warehouse": "RR Warehouse",
            "fk_warehouse": "FK Warehouse",
            "sap_code": "SAP Code",
            "allocated_qty": "Qty."
        })

        st.markdown("### Invoice-wise Billing Lines")

        st.dataframe(
            billing_df,
            use_container_width=True
        )

        st.markdown("---")
        st.markdown("### Invoice-wise Billing Lines With Barcode")

        for _, row in billing_df.iterrows():
            fsn = clean_text(row["FSN"])

            with st.container():
                c1, c2, c3, c4, c5, c6 = st.columns(
                    [2, 2, 3, 4, 2, 2]
                )

                with c1:
                    st.markdown("**Invoice No.**")
                    st.write(row["Invoice No."])

                with c2:
                    st.markdown("**PO No.**")
                    st.write(row["PO No."])

                with c3:
                    st.markdown("**FSN**")
                    st.write(row["FSN"])

                with c4:
                    st.markdown("**FSN Barcode**")

                    barcode_bytes = get_fsn_barcode_bytes(fsn, fsn_barcode_cache)

                    if barcode_bytes:
                        st.image(
                            barcode_bytes,
                            width=300
                        )
                    else:
                        st.warning("Barcode Not Found")

                with c5:
                    st.markdown("**SAP Code**")
                    st.write(row["SAP Code"])

                with c6:
                    st.markdown("**Qty.**")
                    st.write(row["Qty."])

                st.markdown("---")

        invoice_summary = billing_df.groupby(
            "Invoice No."
        )["Qty."].sum().reset_index()

        invoice_summary = invoice_summary.rename(
            columns={"Qty.": "Total Qty."}
        )

        st.markdown("### Invoice-wise Total Qty")

        st.dataframe(
            invoice_summary,
            use_container_width=True
        )

        st.download_button(
            "Download Billing Summary With Exact Barcode",
            data=to_excel_billing_with_exact_barcodes(
                billing_df,
                fsn_barcode_cache
            ),
            file_name="billing_summary_with_exact_barcode.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            "Download Warehouse Sticker Sheet PDF",
            data=generate_sticker_pdf(
                billing_df,
                fsn_barcode_cache
            ),
            file_name="warehouse_sticker_sheet.pdf",
            mime="application/pdf"
        )


# =====================================================
# USER MANAGEMENT
# =====================================================

elif menu == "User Management":
    render_page_header(
        "User Management",
        "Create team logins, change roles, and manage account access."
    )

    if st.session_state.role != "Admin":
        st.error("Only Admin users can access this screen.")
        st.stop()

    st.markdown("### Create User")

    u1, u2, u3 = st.columns(3)

    with u1:
        new_username = st.text_input("New Username")

    with u2:
        new_password = st.text_input("New Password", type="password")

    with u3:
        new_role = st.selectbox(
            "Role",
            ["Ops", "Billing", "Viewer", "Admin"]
        )

    if st.button("Create User"):
        new_username = clean_text(new_username)

        if new_username == "":
            st.error("Please enter a username.")

        elif len(new_password) < 6:
            st.error("Password should be at least 6 characters.")

        else:
            try:
                db_execute("""
                    INSERT INTO app_users (username, password_hash, role, active)
                    VALUES (:username, :password_hash, :role, TRUE)
                """, {
                    "username": new_username,
                    "password_hash": hash_password(new_password),
                    "role": new_role
                })

                log_activity(
                    "create_user",
                    f"Created user {new_username} with role {new_role}"
                )
                st.success(f"User {new_username} created successfully.")
                st.rerun()

            except Exception as e:
                st.error(f"Could not create user. It may already exist. Error: {e}")

    st.markdown("---")
    st.markdown("### Manage Existing Users")

    users_df = db_read("""
        SELECT
            id,
            username,
            role,
            active,
            created_at
        FROM app_users
        ORDER BY id
    """)

    st.markdown("#### Change User Access")

    selected_username = st.selectbox(
        "Select User",
        options=users_df["username"].tolist()
    )

    selected_user_row = users_df[users_df["username"] == selected_username].iloc[0]

    a1, a2 = st.columns(2)

    with a1:
        selected_role = st.selectbox(
            "New Role",
            ["Admin", "Ops", "Billing", "Viewer"],
            index=["Admin", "Ops", "Billing", "Viewer"].index(selected_user_row["role"])
        )

    with a2:
        selected_active = st.checkbox(
            "Active User",
            value=bool(selected_user_row["active"])
        )

    if st.button("Update Selected User"):
        would_remove_self_admin = (
            selected_username == st.session_state.username and
            (selected_role != "Admin" or not selected_active)
        )

        other_active_admins = users_df[
            (users_df["username"] != selected_username) &
            (users_df["role"] == "Admin") &
            (users_df["active"] == True)
        ]

        if would_remove_self_admin:
            st.error("You cannot remove your own active Admin access while logged in.")

        elif selected_role != "Admin" and other_active_admins.empty:
            st.error("At least one active Admin user is required.")

        elif not selected_active and selected_role == "Admin" and other_active_admins.empty:
            st.error("At least one active Admin user is required.")

        else:
            db_execute("""
                UPDATE app_users
                SET
                    role = :role,
                    active = :active
                WHERE id = :id
            """, {
                "role": selected_role,
                "active": selected_active,
                "id": int(selected_user_row["id"])
            })

            log_activity(
                "update_user",
                f"Updated {selected_username} to role {selected_role}, active={selected_active}"
            )
            st.success(f"{selected_username} updated successfully.")
            st.rerun()

    st.markdown("#### User Table")

    edited_users_df = st.data_editor(
        users_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "id": st.column_config.NumberColumn("ID", disabled=True),
            "username": st.column_config.TextColumn("Username", disabled=True),
            "role": st.column_config.SelectboxColumn(
                "Role",
                options=["Admin", "Ops", "Billing", "Viewer"],
                required=True
            ),
            "active": st.column_config.CheckboxColumn("Active"),
            "created_at": st.column_config.TextColumn("Created At", disabled=True),
        }
    )

    if st.button("Save User Changes"):
        active_admin_count = len(
            edited_users_df[
                (edited_users_df["role"] == "Admin") &
                (edited_users_df["active"] == True)
            ]
        )

        current_user_row = edited_users_df[
            edited_users_df["username"] == st.session_state.username
        ]

        if active_admin_count == 0:
            st.error("At least one active Admin user is required.")

        elif (
            not current_user_row.empty and
            (
                current_user_row.iloc[0]["role"] != "Admin" or
                not bool(current_user_row.iloc[0]["active"])
            )
        ):
            st.error("You cannot remove your own active Admin access while logged in.")

        else:
            for _, user_row in edited_users_df.iterrows():
                db_execute("""
                    UPDATE app_users
                    SET
                        role = :role,
                        active = :active
                    WHERE id = :id
                """, {
                    "role": user_row["role"],
                    "active": bool(user_row["active"]),
                    "id": int(user_row["id"])
                })

            log_activity("update_users", "Updated user roles or active status")
            st.success("User changes saved successfully.")
            st.rerun()


# =====================================================
# OPEN ALLOCATION
# =====================================================

elif menu == "Open Allocation Qty":
    render_page_header(
        "Open Allocation Qty",
        "Stock already sent for billing but not yet marked as billed."
    )

    open_alloc = get_open_allocation_qty()

    if open_alloc.empty:
        st.info("No open allocations found.")

    else:
        st.dataframe(
            open_alloc,
            use_container_width=True
        )
