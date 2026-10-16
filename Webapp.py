import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Import Shipment Report", layout="wide")
st.title("Import Shipment Report")

# keyword -> ชื่อสั้นที่ใช้ในโค้ด (หาแบบ contains, ไม่สนตัวพิมพ์)
COLMAP = {
    "job": "JobID", "eta": "ETA", "ata": "ATA", "invoice": "InvoiceNo",
    "origin": "Origin", "release port": "ReleasePort", "carrier": "Carrier",
    "ctnr no. count": "CtnrSeq", "ctnr no.": "CtnrNo", "seal": "Seal",
    "container size": "Size", "cartons": "Cartons", "cbm": "CBM",
    "filenamedate": "FileNameDate", "do release": "DORelease",
    "shipping document receive": "DocReceive",
    "issue reason": "DocIssue", "shipping document complete": "DocComplete",
    "status declaration": "Status", "(0209)": "CustomsRelease0209",
    "delivery plan": "DeliveryPlan", "(0309)": "DeclRelease0309",
    "delivery date": "DeliveryDate", "delivery place": "DeliveryPlace",
}
DATE_COLS = ["ETA", "ATA", "DORelease", "CustomsRelease0209",
             "DeliveryPlan", "DeclRelease0309", "DeliveryDate"]

def rename_cols(df):
    new = {}
    for col in df.columns:
        c = " ".join(str(col).split()).lower()
        for key, name in COLMAP.items():          # ลำดับสำคัญ: "ctnr no. count" ก่อน "ctnr no."
            if key in c and name not in new.values():
                new[col] = name
                break
    return df.rename(columns=new)

def to_dt(s, fmt):
    s = s.astype(str).str.strip()
    a = pd.to_datetime(s, format=fmt, errors="coerce")                   # text เช่น 09/01/2026
    b = pd.to_datetime(s, format="%Y-%m-%d %H:%M:%S", errors="coerce")   # Excel date จริง
    c = pd.to_datetime(s, format="%Y-%m-%d", errors="coerce")
    return a.fillna(b).fillna(c)

@st.cache_data
def load(f):
    df = rename_cols(pd.read_excel(f, dtype=str))
    # st.write(df[["ETA", "ATA", "DeliveryDate", "FileNameDate"]].head())
    df = df[df["JobID"].str.match(r"^I\d+", na=False)].copy()   # ตัดแถวขยะ

    for c in DATE_COLS:
        if c in df:
            df[c] = to_dt(df[c], "%m/%d/%Y")
    df["FileNameDate"] = to_dt(df["FileNameDate"].str.replace(r"\.0$", "", regex=True), "%Y%m%d")
    df["DocComplete"] = to_dt(df["DocComplete"], "%m/%d/%Y %H:%M:%S")

    # "2026-08-20 16:44:37 2026-09-10 16:34:10" -> first / last
    ts = df["DocReceive"].fillna("").str.findall(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}")
    df["DocReceiveFirst"] = pd.to_datetime(ts.str[0], errors="coerce")
    df["DocReceiveLast"] = pd.to_datetime(ts.str[-1], errors="coerce")

    for c in ("Cartons", "CBM"):
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df["Declared"] = df["Status"].fillna("").str.strip().eq("409")
    df["HasDocIssue"] = df["DocIssue"].fillna("").str.strip().ne("")
    df["Carrier"] = df["Carrier"].str.strip().str.rstrip(".")      # "CNC LINE." -> "CNC LINE"
    df["LeadTimeDays"] = (df["DeliveryDate"] - df["ATA"]).dt.days  # ATA -> ส่งของ
    return df

file = st.file_uploader("Import Excel", type=["xlsx"])
if not file:
    st.stop()
df = load(file)

# ---------- Filter ----------
with st.sidebar:
    st.header("Filter")
    date_col = st.selectbox("ใช้วันที่จาก", ["ETA", "ATA", "DORelease", "DeliveryDate"])
    s = df[date_col].dropna()
    if s.empty:
        st.warning(f"คอลัมน์ {date_col} ไม่มีวันที่ที่อ่านได้")
        dr = None
    else:
        dr = st.date_input("ช่วงวันที่", (s.min().date(), s.max().date()))
    q_job = st.text_input("Job ID")
    q_inv = st.text_input("Invoice No")
    q_ctnr = st.text_input("Container No")

    def ms(label, col):
        return st.multiselect(label, sorted(df[col].dropna().unique()))
    f_origin = ms("Origin", "Origin")
    f_port = ms("Release Port", "ReleasePort")
    f_carrier = ms("Carrier", "Carrier")
    f_size = ms("Container Size", "Size")
    f_status = st.radio("Declaration", ["All", "Declared (409)", "Pending"], horizontal=True)
    only_issue = st.checkbox("เฉพาะที่มีปัญหาเอกสาร")

f = df.copy()
if isinstance(dr, tuple) and len(dr) == 2:
    f = f[f[date_col].dt.date.between(dr[0], dr[1])]
for q, col in ((q_job, "JobID"), (q_inv, "InvoiceNo"), (q_ctnr, "CtnrNo")):
    if q:
        f = f[f[col].str.contains(q, case=False, na=False)]
for sel, col in ((f_origin, "Origin"), (f_port, "ReleasePort"),
                 (f_carrier, "Carrier"), (f_size, "Size")):
    if sel:
        f = f[f[col].isin(sel)]
if f_status == "Declared (409)":
    f = f[f["Declared"]]
elif f_status == "Pending":
    f = f[~f["Declared"]]
if only_issue:
    f = f[f["HasDocIssue"]]

# ---------- KPI (CBM/Cartons dedupe ตาม Invoice) ----------
inv = f.drop_duplicates("InvoiceNo")
k = st.columns(6)
k[0].metric("Jobs", f["JobID"].nunique())
k[1].metric("Invoices", f["InvoiceNo"].nunique())
k[2].metric("Containers", f["CtnrNo"].nunique())
k[3].metric("CBM", f"{inv['CBM'].sum():,.2f}")
k[4].metric("Pending declaration", int((~f["Declared"]).sum()))
k[5].metric("Avg ATA→Delivery (days)", f"{f['LeadTimeDays'].mean():.1f}")

# ---------- Charts ----------
c1, c2 = st.columns(2)
c1.subheader("Containers by Origin")
c1.bar_chart(f.groupby("Origin")["CtnrNo"].nunique())
c2.subheader("Containers by Carrier")
c2.bar_chart(f.groupby("Carrier")["CtnrNo"].nunique())

c3, c4 = st.columns(2)
c3.subheader(f"Containers per day ({date_col})")
c3.bar_chart(f.groupby(f[date_col].dt.date)["CtnrNo"].nunique())
c4.subheader("Container Size")
c4.bar_chart(f["Size"].value_counts())

st.subheader("Data")
st.dataframe(f, use_container_width=True)

# ---------- Export ----------
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as w:
    f.to_excel(w, sheet_name="Data", index=False)
    inv.groupby("Origin").agg(Invoices=("InvoiceNo", "nunique"),
                              CBM=("CBM", "sum"),
                              Cartons=("Cartons", "sum")).to_excel(w, sheet_name="By Origin")
    f.groupby("Carrier")["CtnrNo"].nunique().to_frame("Containers").to_excel(w, sheet_name="By Carrier")
    f[~f["Declared"]].to_excel(w, sheet_name="Pending", index=False)

st.download_button(
    "Export Report (Excel)",
    data=buf.getvalue(),
    file_name="export_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)