import re
import pandas as pd
from io import StringIO
import streamlit as st
from fuzzywuzzy import fuzz, process

st.set_page_config(page_title="Petty Cash → QuickBooks IIF", layout="wide")
st.title("💵 Petty Cash → QuickBooks IIF")

# ---------- Helpers ----------
def norm(s: str) -> str:
    """lower, strip, collapse non-alnum to single space"""
    return re.sub(r"[^a-z0-9]+", " ", str(s).strip().lower()).strip()

def qb_date(x):
    try:
        return pd.to_datetime(x, errors="coerce").strftime("%m/%d/%Y")
    except Exception:
        return str(x)

def clean_text(x):
    return str(x).replace('"', "").replace("\n", " ").strip()

# ---------- Vendor Master ----------
VENDORS = [
    "254 Brewing Company", "A.S.W Enterprises Limited", "A.W Water Boozer Services",
    "AAA Growers LTD", "Alyemda Enterprise Ltd", "Araali Limited",
    "Assmazab General Stores", "Baraka Israel Enterprises Limited",
    "Benchmark Distributors Limited", "Best Buy Distributors",
    "Beyond Fruits Limited", "Bio Food Products Limited", "Boos Ventures",
    "Bowip Agencies Ltd", "Branded Fine Foods Ltd", "Brookside Dairy Ltd",
    "Brown Bags", "Cafesserie Bread Store", "Casks And Barrels Ltd",
    "Chandaria Industries Ltd", "CHIRAG AFRICA LIMITED",
    "Coastal Bottlers Limited", "Crystal Frozen & Chilled Foods Ltd",
    "De Vries Africa Ventures", "Debenham & Fear Ltd", "Dekow Wholesale",
    "Deliveries", "Diamond Trust Bank", "Dilawers",
    "Dion Wine And Spirits East Africa Limited", "Disney Wines & Spirits",
    "Domaine Kenya Ltd", "Dormans Coffee Ltd", "Eco-Essentials Limited",
    "Ewca Marketing (Kenya) Limited", "Exotics Africa Ltd", "Express Shop Bofa",
    "Ezzi Traders Limited", "Farmers Choice Limited", "Fayaz Bakers Ltd",
    "Finsbury Trading Ltd", "Fratres Malindi", "FRAWAROSE LIMITED",
    "Galina Agencies", "Gilani's Distributors LTD", "Glacier Products Ltd",
    "Global  Slacker Enterprises Ltd", "Handas Juice Ltd", "Hasbah Kenya Limited",
    "Healthy U Two Thousand Ltd", "HOME BEST HEALTH FOOD LIMITED",
    "House of Booch Ltd", "Ice Hub Limited", "Isinya Feeds Limited",
    "Jetlak Limited", "Kalon Foods Limited", "Karen Fork", "Kenchic Limited",
    "Kenya Commercial Bank", "Kenya Nut Company",
    "Kenya Power and Lighting Company", "Kenya Revenue Authority",
    "Khosal Wholesale Kilifi", "Kioko Enterprises",
    "Lakhani General Suppliers Lilmited", "Laki Laki Ltd", "Landlord",
    "LEXO ENERGY KENYA LIMITED", "Lindas Nut Butter",
    "Linkbizz E-Hub Commerce Ltd", "Loki Ventures Limited", "Malachite Limited",
    "Malindi Industries Limited", "Mill Bakers", "Mini Bakeries (NRB) ltd",
    "Mjengo Limited", "Mnarani Pens", "Mnarani Water Refil",
    "MohanS Oysterbay Drinks K Ltd", "Moonsun Picture International Limited",
    "Mudee Concepts Limited", "Mwanza Kambi Tsuma", "Mzuri Sweets Limited",
    "Naaman Muses & Co. Ltd", "Nairobi Java House Limited", "Naji Superstores",
    "National Social Security Fund", "Neema Stores Kilifi", "Njugu Supplier",
    "Nyali Air Conditioning & Refrigeration Se", "Pasagot Limited", "Plastic Cups",
    "Pride Industries Ltd", "Radbone-Clark Kenya limited",
    "Raisons Distributors Ltd", "Rehema Jumwa Muli", "RK'S Products",
    "S.H.I.F Payable", "Safaricom", "Savannah Brands Company Ltd",
    "SEA HARVEST (K) LTD", "Shiva Mombasa Limited", "SIDR Distributors Limited",
    "Slater", "Sliquor Limited", "Social Health Insurance Fund",
    "Soko (Market)", "Sol O Vino Limited", "South Lemon LTD", "Soy's Limited",
    "Sun Power Products Limited", "Supreme Filing Station", "Takataka",
    "Tandaa Networks Limited", "Taraji", "Tawakal Store Company Ltd",
    "The Happy Lamb Butchery", "The Standard Group Plc",
    "Thomas Mwachongo Mwangala", "Three Spears Limited", "TOP IT UP DISTRIBUTOR",
    "Towfiq Kenya Limited", "Traderoots Limited",
    "Under the Influence East Africa", "UvA Wines", "VEGAN WORLD LIMITED",
    "Vyema Eggs", "Water Refil", "Wine and More Limited", "Wingu Box Ltd",
    "Zabach", "Zabach Enterprises Limited", "Zen Mahitaji Ltd",
    "Zenko Kenya Limited", "Zuri Central"
]

def match_supplier(detail: str, suppliers: list, threshold: int = 85) -> str | None:
    """Match petty cash detail text to a known supplier by fuzzy match."""
    if not detail:
        return None
    detail_norm = norm(detail)

    # Exact word matches first
    for sup in suppliers:
        sup_norm = norm(sup)
        for word in sup_norm.split():
            if word and word in detail_norm.split():
                return sup  # strong match on word

    # Fuzzy best match
    best, score = process.extractOne(detail_norm, suppliers, scorer=fuzz.token_set_ratio)
    if score >= threshold:
        return best

    return None

# ---------- Column Mapping ----------
EXPECTED = {
    "pay type": ["pay type","paytype","type","pay_type"],
    "till no": ["till no","till","till number","till_no","tillno"],
    "transaction date": ["transaction date","date","txn date","trans date","txndate"],
    "detail": ["detail","details","description","memo","narration"],
    "transacted amount": ["transacted amount","amount","amt","value","transaction amount"],
    "user name": ["user name","username","user","cashier","handled by","handledby"]
}

def find_columns(df_cols):
    normalized = {norm(c): c for c in df_cols}
    mapping = {}
    for target, alts in EXPECTED.items():
        found = None
        for alt in alts:
            if alt in normalized:
                found = normalized[alt]
                break
        if found is None:
            for key, orig in normalized.items():
                if all(w in key for w in norm(alts[0]).split()):
                    found = orig
                    break
        mapping[target] = found
    return mapping

# ---------- Classification ----------
def classify_and_rows(row, seq):
    pay_type = norm(row["pay type"])
    details_n = norm(row["detail"])
    details_raw = clean_text(row["detail"])
    user = clean_text(row["user name"])
    date_str = qb_date(row["transaction date"])
    amt = abs(float(row["transacted amount"]))
    clear = "N"
    docnum = f"{pd.to_datetime(row['transaction date']).strftime('%Y%m%d')}{seq:03d}"

    memo_petty = f"Petty cash by {user} for {details_raw} on {date_str}"
    memo_deliv = f"Delivery expense on behalf of customer on {date_str}"
    memo_trans = f"Interbranch transport by {user} on {date_str}"
    memo_pick  = f"Cash pick up for deposit on {date_str}"

    # Cash Pickup → TRANSFER Cash in Drawer -> DTB
    if "cash" in pay_type and "pickup" in pay_type:
        return [
            ["TRNS", "TRANSFER", date_str, "Cash in Drawer", "Diamond Trust Bank", -amt, memo_pick, docnum, clear],
            ["SPL",  "TRANSFER", date_str, "Diamond Trust Bank", "Diamond Trust Bank",  amt, memo_pick, docnum, clear],
        ]

    # Delivery
    if "deliv" in details_n:
        return [
            ["TRNS", "CHECK", date_str, "Cash in Drawer", user, -amt, memo_deliv, docnum, clear],
            ["SPL",  "CHECK", date_str, "Customer Deliveries", user,  amt, memo_deliv, docnum, clear],
        ]

    # Transport
    if any(k in details_n for k in ["fare", "fair", "transport", "trasport"]):
        return [
            ["TRNS", "CHECK", date_str, "Cash in Drawer", user, -amt, memo_trans, docnum, clear],
            ["SPL",  "CHECK", date_str, "Interbranch Transport Cost", user,  amt, memo_trans, docnum, clear],
        ]

    # Vendor match
    matched = match_supplier(details_raw)
    if matched:
        vendor_name = matched
    else:
        vendor_name = details_raw.title() if details_raw else "Vendor"

    return [
        ["TRNS", "CHECK", date_str, "Cash in Drawer", vendor_name, -amt, memo_petty, docnum, clear],
        ["SPL",  "CHECK", date_str, "Accounts Payable", vendor_name,  amt, memo_petty, docnum, clear],
    ]

# ---------- IIF Builder ----------
def build_iif(df):
    out = StringIO()
    out.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tMEMO\tDOCNUM\tCLEAR\n")
    out.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tMEMO\tDOCNUM\tCLEAR\n")
    out.write("!ENDTRNS\n")

    seq = 1
    for _, r in df.iterrows():
        lines = classify_and_rows(r, seq)
        for line in lines:
            out.write("\t".join([str(x) for x in line]) + "\n")
        out.write("ENDTRNS\n")
        seq += 1
    return out.getvalue()

# ---------- UI ----------
uploaded = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])

if uploaded:
    if uploaded.name.lower().endswith(".csv"):
        df_raw = pd.read_csv(uploaded)
    else:
        df_raw = pd.read_excel(uploaded)

    st.subheader("📄 Raw preview")
    st.dataframe(df_raw.head(20), use_container_width=True)

    colmap = find_columns(df_raw.columns)
    missing = [k for k, v in colmap.items() if v is None]
    if missing:
        st.error("Missing required columns after normalization: " + ", ".join(missing))
        st.stop()

    df = pd.DataFrame({
        "pay type": df_raw[colmap["pay type"]],
        "till no": df_raw[colmap["till no"]],
        "transaction date": df_raw[colmap["transaction date"]],
        "detail": df_raw[colmap["detail"]],
        "transacted amount": df_raw[colmap["transacted amount"]],
        "user name": df_raw[colmap["user name"]],
    })

    for c in ["pay type","till no","detail","user name"]:
        df[c] = df[c].astype(str).fillna("").map(clean_text)
    df["transacted amount"] = pd.to_numeric(df["transacted amount"], errors="coerce").fillna(0)

    st.subheader("✅ Normalized preview")
    st.dataframe(df.head(30), use_container_width=True)

    # Warn on new vendors
    unmatched = []
    for _, r in df.iterrows():
        if not match_supplier(r["detail"]):
            unmatched.append(r["detail"])
    if unmatched:
        st.warning("⚠️ New vendors detected: " + ", ".join(set(unmatched)))

    if st.button("Generate QuickBooks IIF"):
        iif_txt = build_iif(df)
        st.download_button(
            "📥 Download petty_cash.iif",
            data=iif_txt.encode("utf-8"),
            file_name="petty_cash.iif",
            mime="text/plain",
        )

else:
    st.info("Upload your petty cash file (CSV/XLSX) with columns like: Pay Type, Till No, Transaction Date, Detail, Transacted Amount, User Name.")

