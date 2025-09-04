import re
import pandas as pd
from io import StringIO
import streamlit as st
from fuzzywuzzy import fuzz, process

# ==============================
# Streamlit setup
# ==============================
st.set_page_config(page_title="Petty Cash → QuickBooks IIF", layout="wide")
st.title("💵 Petty Cash → QuickBooks IIF")

# ==============================
# Helpers
# ==============================
def norm(s: str) -> str:
    """Lower, strip, collapse non-alnum to single space."""
    return re.sub(r"[^a-z0-9]+", " ", str(s).strip().lower()).strip()

def qb_date(x):
    """Return MM/DD/YYYY or original string if unparsable."""
    try:
        dt = pd.to_datetime(x, errors="coerce", dayfirst=False)
        if pd.isna(dt):
            return str(x)
        return dt.strftime("%m/%d/%Y")
    except Exception:
        return str(x)

def clean_text(x):
    return str(x).replace('"', "").replace("\n", " ").strip()

def make_docnum(date_val, seq: int) -> str:
    """DOCNUM as YYYYMMDD + seq; if date invalid, use 00000000."""
    try:
        dt = pd.to_datetime(date_val, errors="coerce")
        if pd.isna(dt):
            return f"00000000{seq:03d}"
        return f"{dt.strftime('%Y%m%d')}{seq:03d}"
    except Exception:
        return f"00000000{seq:03d}"

# ==============================
# Vendor Master
# ==============================
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

# Common noise words that shouldn't decide a match (kept minimal and localizable)
STOPWORDS = {
    "ltd", "limited", "enterprises", "enterprise", "company", "kenya", "plc", "group",
    "east", "africa", "and", "the", "store", "stores", "trading", "distributor",
    "distributors", "foods", "food", "products", "general", "suppliers", "supplier",
    "ventures", "agencies", "agency", "industries", "wholesale", "bakers", "services",
    "two", "thousand", "k", "co", "lilmited", "ltd.", "limited.", "ltds"
}

def tokens(s: str) -> list[str]:
    return [t for t in norm(s).split() if t and t not in STOPWORDS]

# Build an alias map: unique significant words map to a vendor
def build_alias_map(vendors: list[str]) -> dict[str, str]:
    word_to_vendors = {}
    for v in vendors:
        for t in set(tokens(v)):
            word_to_vendors.setdefault(t, set()).add(v)
    # Only keep words that uniquely identify a vendor (avoid words shared by many)
    alias = {}
    for w, vs in word_to_vendors.items():
        if len(vs) == 1:
            alias[w] = list(vs)[0]
    # also add some handpicked short aliases that are common in details
    manual = {
        "brookside": "Brookside Dairy Ltd",
        "fayaz": "Fayaz Bakers Ltd",
        "benchmark": "Benchmark Distributors Limited",
        "takataka": "Takataka",
        "zuri": "Zuri Central",
        "lexo": "LEXO ENERGY KENYA LIMITED",
        "safaricom": "Safaricom",
        "kcb": "Kenya Commercial Bank",
        "dtb": "Diamond Trust Bank",
        "kenchic": "Kenchic Limited",
    }
    alias.update(manual)
    return alias

ALIAS_MAP = build_alias_map(VENDORS)

def strip_stopwords(s: str) -> str:
    return " ".join(tokens(s))

def match_supplier(detail: str, suppliers: list[str], threshold: int = 86) -> str | None:
    """
    Robust supplier matching:
    1) Alias/keyword hit (unique token) -> immediate match
    2) Exact token subset (>=1 significant token overlap) -> candidate
    3) Fuzzy best match using token_set_ratio over stopword-stripped strings
    """
    if not detail:
        return None

    detail_tokens = set(tokens(detail))

    # 1) alias direct hit
    for t in detail_tokens:
        if t in ALIAS_MAP:
            return ALIAS_MAP[t]

    # 2) simple token intersection heuristic (prefer vendors with overlap on significant tokens)
    best_overlap = (None, 0)
    for sup in suppliers:
        sup_tokens = set(tokens(sup))
        overlap = len(detail_tokens & sup_tokens)
        if overlap > best_overlap[1] and overlap > 0:
            best_overlap = (sup, overlap)
    # Note: don't return immediately; still run fuzzy to avoid false positives from generic words

    # 3) fuzzy fallback on stripped strings
    stripped_detail = strip_stopwords(detail)
    stripped_suppliers = [strip_stopwords(s) for s in suppliers]
    # Keep index to map back to original supplier name
    best_match_idx, best_score = None, -1
    for i, ss in enumerate(stripped_suppliers):
        score = fuzz.token_set_ratio(stripped_detail, ss)
        if score > best_score:
            best_score, best_match_idx = score, i

    fuzzy_winner = suppliers[best_match_idx] if best_match_idx is not None else None

    # Decision: prefer fuzzy if it clears threshold; otherwise, if we had overlap, use that
    if best_score >= threshold and fuzzy_winner:
        return fuzzy_winner
    if best_overlap[0]:
        return best_overlap[0]

    return None

# ==============================
# Column Mapping
# ==============================
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

# ==============================
# Classification
# ==============================
def classify_and_rows(row, seq, threshold):
    pay_type = norm(row["pay type"])
    details_n = norm(row["detail"])
    details_raw = clean_text(row["detail"])
    user = clean_text(row["user name"])
    date_str = qb_date(row["transaction date"])
    amt = abs(float(row["transacted amount"]))
    clear = "N"
    docnum = make_docnum(row["transaction date"], seq)

    memo_petty = f"Petty cash by {user} for {details_raw} on {date_str}"
    memo_deliv = f"Delivery expense on behalf of customer on {date_str}"
    memo_trans = f"Interbranch transport by {user} on {date_str}"
    memo_pick  = f"Cash pick up for deposit on {date_str}"

    # 1) Cash Pickup => TRANSFER Cash in Drawer -> Diamond Trust Bank
    if "cash" in pay_type and "pickup" in pay_type:
        return [
            ["TRNS", "TRANSFER", date_str, "Cash in Drawer", "Diamond Trust Bank", -amt, memo_pick, docnum, clear],
            ["SPL",  "TRANSFER", date_str, "Diamond Trust Bank", "Diamond Trust Bank",  amt, memo_pick, docnum, clear],
        ]

    # 2) Deliveries
    if "deliv" in details_n:
        return [
            ["TRNS", "CHECK", date_str, "Cash in Drawer", user, -amt, memo_deliv, docnum, clear],
            ["SPL",  "CHECK", date_str, "Customer Deliveries", user,  amt, memo_deliv, docnum, clear],
        ]
    # 3) Bank Charges - Mpesa
    if "transaction cost" in details_n:
        memo_bank = f"MPESA transaction cost on {date_str}"
        return [
            ["TRNS", "CHECK", date_str, "Cash in Drawer", "Safaricom", -amt, memo_bank, docnum, clear],
            ["SPL",  "CHECK", date_str, "Bank Charges - Mpesa", "Safaricom",  amt, memo_bank, docnum, clear],
        ]


    # 4) Staff transport (fare/fair/transport/trasport)
    if any(k in details_n for k in ["fare", "fair", "transport", "trasport"]):
        return [
            ["TRNS", "CHECK", date_str, "Cash in Drawer", user, -amt, memo_trans, docnum, clear],
            ["SPL",  "CHECK", date_str, "Interbranch Transport Cost", user,  amt, memo_trans, docnum, clear],
        ]

    # 4) Vendor purchase
    matched = match_supplier(details_raw, VENDORS, threshold=threshold)
    vendor_name = matched if matched else (details_raw.title() if details_raw else "Vendor")

    return [
        ["TRNS", "CHECK", date_str, "Cash in Drawer", vendor_name, -amt, memo_petty, docnum, clear],
        ["SPL",  "CHECK", date_str, "Accounts Payable", vendor_name,  amt, memo_petty, docnum, clear],
    ]

# ==============================
# IIF Builder
# ==============================
def build_iif(df, threshold):
    out = StringIO()
    out.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tMEMO\tDOCNUM\tCLEAR\n")
    out.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tMEMO\tDOCNUM\tCLEAR\n")
    out.write("!ENDTRNS\n")

    seq = 1
    for _, r in df.iterrows():
        lines = classify_and_rows(r, seq, threshold)
        for tag, trnstype, date_str, accnt, name, amount, memo, docnum, clear in lines:
            out.write("\t".join([
                tag, trnstype, date_str, accnt, name, str(amount), memo, docnum, clear
            ]) + "\n")
        out.write("ENDTRNS\n")
        seq += 1
    return out.getvalue()

# ==============================
# UI
# ==============================
with st.sidebar:
    st.header("Settings")
    fuzzy_threshold = st.slider("Fuzzy match threshold", min_value=70, max_value=98, value=86, step=1,
                                help="Higher = stricter vendor matching. If no match meets the threshold, a new vendor will be created.")

uploaded = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])

if uploaded:
    # Read file
    if uploaded.name.lower().endswith(".csv"):
        df_raw = pd.read_csv(uploaded)
    else:
        # openpyxl required for .xlsx
        df_raw = pd.read_excel(uploaded)

    st.subheader("📄 Raw preview")
    st.dataframe(df_raw.head(20), use_container_width=True)

    # Map & validate columns
    colmap = find_columns(df_raw.columns)
    missing = [k for k, v in colmap.items() if v is None]
    if missing:
        st.error(
            "Missing required columns after normalization: "
            + ", ".join(missing)
            + ".\nFound columns: "
            + ", ".join(df_raw.columns.astype(str))
        )
        st.stop()

    # Normalize working frame
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

    # Detect unmatched/new vendors
    unmatched = []
    suggested = []
    for _, r in df.iterrows():
        d = r["detail"]
        m = match_supplier(d, VENDORS, threshold=fuzzy_threshold)
        if m:
            suggested.append((d, m))
        else:
            unmatched.append(d)

    if unmatched:
        st.warning("⚠️ New/Unmatched vendors detected (examples): " + ", ".join(list(dict.fromkeys(unmatched[:10]))))
    if suggested:
        with st.expander("📎 Suggested vendor matches (sample)"):
            samp = pd.DataFrame(suggested[:20], columns=["Detail", "Matched Vendor"])
            st.dataframe(samp, use_container_width=True)

    # Build IIF
    if st.button("Generate QuickBooks IIF"):
        iif_txt = build_iif(df, threshold=fuzzy_threshold)
        st.download_button(
            "📥 Download petty_cash.iif",
            data=iif_txt.encode("utf-8"),
            file_name="petty_cash.iif",
            mime="text/plain",
        )
else:
    st.info("Upload your petty cash file (CSV/XLSX) with columns like: Pay Type, Till No, Transaction Date, Detail, Transacted Amount, User Name.")

