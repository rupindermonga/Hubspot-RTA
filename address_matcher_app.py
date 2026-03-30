import streamlit as st
import pandas as pd
import re
import io
import hashlib
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ── Page config ──
st.set_page_config(page_title="Address Matcher", page_icon="📍", layout="wide")

# ── Authentication ──
# Credentials stored as SHA-256 hashes in .streamlit/secrets.toml
# To add a user:
#   1. Generate hash: python -c "import hashlib; print(hashlib.sha256('PASSWORD'.encode()).hexdigest())"
#   2. Add to .streamlit/secrets.toml under [users]: username = "hash"
#   3. On Streamlit Cloud: add the same in the app's Secrets settings

def verify_password(password, hashed):
    """Check password against SHA-256 hash."""
    return hashlib.sha256(password.encode()).hexdigest() == hashed

def login():
    """Show login form and validate credentials."""
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
        st.session_state.username = ''

    if st.session_state.authenticated:
        return True

    users = st.secrets.get("users", {})

    st.title("🔐 Address Matcher — Login")
    st.markdown("Please log in to continue.")

    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Log in", use_container_width=True)

        if submitted:
            if username in users and verify_password(password, users[username]):
                st.session_state.authenticated = True
                st.session_state.username = username
                st.rerun()
            else:
                st.error("Invalid username or password.")
    return False

if not login():
    st.stop()

# ── Logged in: show user + logout ──
st.sidebar.markdown(f"**Logged in as:** {st.session_state.username}")
if st.sidebar.button("Logout"):
    st.session_state.authenticated = False
    st.session_state.username = ''
    st.rerun()

st.title("📍 Address Matcher")
st.markdown("Upload **Hubspot** and **RTA** files separately. The app matches addresses and appends RTA Address + RTA Status to the Hubspot file.")

# ── Normalization engine ──
ABBREVS = {
    'STREET': 'ST', 'ROAD': 'RD', 'DRIVE': 'DR', 'AVENUE': 'AVE',
    'BOULEVARD': 'BLVD', 'CRESCENT': 'CRES', 'COURT': 'CRT', 'PLACE': 'PL',
    'LANE': 'LN', 'CIRCLE': 'CIR', 'TERRACE': 'TERR', 'HIGHWAY': 'HWY',
    'TRAIL': 'TRL', 'SQUARE': 'SQ', 'PARKWAY': 'PKY', 'WAY': 'WAY',
    'CLOSE': 'CL', 'GROVE': 'GRV', 'HEIGHTS': 'HTS', 'RIDGE': 'RDG',
    'NORTH': 'N', 'SOUTH': 'S', 'EAST': 'E', 'WEST': 'W',
    'REG': 'REGIONAL',
}

def clean_address(s):
    """Strip PO Box, RR, Suite, Apt, Unit noise."""
    s = str(s).strip()
    s = re.sub(r'^(?:P[\s.]?O[\s.]?\s*)?BOX\s*#?\s*\d*[\s,/]*', '', s, flags=re.IGNORECASE)
    s = re.sub(r'^SUITE\s+\d+\s*', '', s, flags=re.IGNORECASE)
    s = re.sub(r'^RR\s*#?\s*\d+[\s,]*', '', s, flags=re.IGNORECASE)
    s = re.sub(r'[\s,]+(?:P[\s.]?O[\s.]?\s*)?BOX\s*#?\s*\d*.*$', '', s, flags=re.IGNORECASE)
    s = re.sub(r'[\s,]+RR\s*#?\s*\d+.*$', '', s, flags=re.IGNORECASE)
    s = re.sub(r'[\s,]+(?:SUITE|APT|UNIT)\s*#?\s*\w*.*$', '', s, flags=re.IGNORECASE)
    return s.strip(' ,/')

def normalize(s):
    s = clean_address(s)
    s = s.upper().replace('.', '')
    s = re.sub(r'[^A-Z0-9 ]', '', s)
    s = re.sub(r' +', ' ', s).strip()
    words = s.split()
    words = [ABBREVS.get(w, w) for w in words]
    return ' '.join(words)

def strip_direction(s):
    return re.sub(r'\s+[NSEW]$', '', s)

def norm_pc(s):
    s = str(s).strip().upper().replace(' ', '')
    corrected = []
    for i, c in enumerate(s[:6]):
        if i in (1, 3, 5):
            if c == 'O': c = '0'
            elif c == 'I': c = '1'
        elif i in (0, 2, 4):
            if c == '0': c = 'O'
        corrected.append(c)
    s = ''.join(corrected)
    return s[:3]

def apply_canonical(street_full, canonical_map):
    m = re.match(r'^(\d+[A-Z]?\s+)(.*)', street_full)
    if m:
        hnum, sname = m.group(1), m.group(2)
    else:
        hnum, sname = '', street_full
    sname = canonical_map.get(sname, sname)
    return hnum + sname


# ── Sidebar: Aliases ──
st.sidebar.header("Street Name Aliases")
st.sidebar.markdown("Add aliases for streets that are the same road but named differently.")

if 'aliases' not in st.session_state:
    st.session_state.aliases = [
        ('PANACHE N SHR RD', 'PANACHE NSHORE RD'),
        ('PANACHE N SHORE RD', 'PANACHE NSHORE RD'),
        ('PANACHE NORTHSHORE RD', 'PANACHE NSHORE RD'),
        ('A PANACHE N SHORE RD', 'PANACHE NSHORE RD'),
        ('A PANACHE N SHR RD', 'PANACHE NSHORE RD'),
        ('PANACHE SHORE RD', 'PANACHE NSHORE RD'),
        ('NORTHSHORE RD', 'PANACHE NSHORE RD'),
        ('N SHORE RD', 'PANACHE NSHORE RD'),
        ('PENACHE NORTHSHORE RD', 'PANACHE NSHORE RD'),
        ('PENACHE N SHORE RD', 'PANACHE NSHORE RD'),
        ('HENNESSY RD', 'HENNESSEY RD'),
        ('OLD SYLVAIN VALLEY HILL RD', 'OLD SLYVAN VALLEY HILL RD'),
        ('LITTLE PENAGE LAKE RD', 'LITTLE PANACHE RD'),
        ('REGIONAL 10 RD', 'REGIONAL RD 10'),
        ('FINDLAY HILL RD', 'FINDLAY RD'),
        ('FINDLAY HILL RD E', 'FINDLAY RD E'),
        ('FINDLAY HILL RD W', 'FINDLAY RD W'),
    ]

with st.sidebar.expander("Current aliases", expanded=False):
    for frm, to in st.session_state.aliases:
        st.text(f"{frm}  →  {to}")

with st.sidebar.form("add_alias"):
    st.markdown("**Add new alias** (use normalized form: uppercase, abbreviated)")
    new_from = st.text_input("From (variant name)")
    new_to = st.text_input("To (canonical name)")
    if st.form_submit_button("Add alias"):
        if new_from and new_to:
            st.session_state.aliases.append((new_from.upper().strip(), new_to.upper().strip()))
            st.success(f"Added: {new_from} → {new_to}")


# ── File uploads ──
st.markdown("---")
upload_col1, upload_col2 = st.columns(2)

with upload_col1:
    st.subheader("1. Hubspot File")
    hub_file = st.file_uploader("Upload Hubspot file (.xlsx / .csv)", type=['xlsx', 'csv'], key='hub')

with upload_col2:
    st.subheader("2. RTA File")
    rta_file = st.file_uploader("Upload RTA file (.xlsx / .csv)", type=['xlsx', 'csv'], key='rta')


def load_file(uploaded_file):
    """Load uploaded file as DataFrame, handling xlsx (with sheet selection) and csv."""
    if uploaded_file.name.endswith('.csv'):
        return pd.read_csv(uploaded_file), None
    else:
        xl = pd.ExcelFile(uploaded_file)
        return xl, xl.sheet_names


if hub_file and rta_file:

    # ── Load Hubspot ──
    hub_result = load_file(hub_file)
    if isinstance(hub_result[0], pd.ExcelFile):
        hub_xl = hub_result[0]
        hub_sheets = hub_result[1]
        hub_sheet = st.selectbox("Hubspot sheet", hub_sheets, key='hub_sheet') if len(hub_sheets) > 1 else hub_sheets[0]
        df_hub = pd.read_excel(hub_xl, sheet_name=hub_sheet)
    else:
        df_hub = hub_result[0]

    # ── Load RTA ──
    rta_result = load_file(rta_file)
    if isinstance(rta_result[0], pd.ExcelFile):
        rta_xl = rta_result[0]
        rta_sheets = rta_result[1]
        rta_sheet = st.selectbox("RTA sheet", rta_sheets, key='rta_sheet') if len(rta_sheets) > 1 else rta_sheets[0]
        df_rta = pd.read_excel(rta_xl, sheet_name=rta_sheet)
    else:
        df_rta = rta_result[0]

    st.markdown("---")
    cfg1, cfg2 = st.columns(2)

    # ── Hubspot column config ──
    with cfg1:
        st.subheader("Hubspot Columns")
        st.dataframe(df_hub.head(5), use_container_width=True)
        hub_cols = df_hub.columns.tolist()

        hub_street_col = st.selectbox(
            "Street Address column",
            hub_cols,
            index=hub_cols.index('Street Address') if 'Street Address' in hub_cols else 0,
            key='hub_street'
        )
        hub_pc_col = st.selectbox(
            "Postal Code column",
            hub_cols,
            index=hub_cols.index('Postal Code') if 'Postal Code' in hub_cols else 0,
            key='hub_pc'
        )

    # ── RTA column config ──
    with cfg2:
        st.subheader("RTA Columns")
        st.dataframe(df_rta.head(5), use_container_width=True)
        rta_cols = df_rta.columns.tolist()

        st.markdown("**Address components** (will be combined for matching)")
        rta_addr_no_col = st.selectbox(
            "Address Number column",
            rta_cols,
            index=rta_cols.index('AddressNo') if 'AddressNo' in rta_cols else 0,
            key='rta_addr_no'
        )
        rta_street_col = st.selectbox(
            "Street Name column",
            rta_cols,
            index=rta_cols.index('StreetName') if 'StreetName' in rta_cols else 0,
            key='rta_street'
        )
        rta_locality_col = st.selectbox(
            "Locality / City column",
            rta_cols,
            index=rta_cols.index('Locality') if 'Locality' in rta_cols else 0,
            key='rta_locality'
        )
        rta_pc_col = st.selectbox(
            "Postal Code column",
            rta_cols,
            index=rta_cols.index('PostalCode') if 'PostalCode' in rta_cols else 0,
            key='rta_pc'
        )
        rta_status_col = st.selectbox(
            "RTA Status column (to bring into output)",
            rta_cols,
            index=rta_cols.index('RTA Status') if 'RTA Status' in rta_cols else len(rta_cols)-1,
            key='rta_status'
        )

    # ── Run matching ──
    st.markdown("---")
    if st.button("🔍 Run Address Matching", type="primary", use_container_width=True):
        with st.spinner("Matching addresses..."):

            canonical_map = dict(st.session_state.aliases)

            # Build RTA combined full address: "AddressNo StreetName Locality PostalCode"
            df_rta['_rta_full'] = (
                df_rta[rta_addr_no_col].fillna('').astype(str).str.strip() + ' ' +
                df_rta[rta_street_col].fillna('').astype(str).str.strip() + ' ' +
                df_rta[rta_locality_col].fillna('').astype(str).str.strip() + ' ' +
                df_rta[rta_pc_col].fillna('').astype(str).str.strip()
            ).str.strip()

            # Normalize Hubspot
            df_hub['_street'] = df_hub[hub_street_col].fillna('').apply(normalize)
            df_hub['_street_canon'] = df_hub['_street'].apply(lambda s: apply_canonical(s, canonical_map))
            df_hub['_pc'] = df_hub[hub_pc_col].fillna('').apply(norm_pc)

            # Normalize RTA: combine AddressNo + StreetName for matching key
            df_rta['_street'] = (
                df_rta[rta_addr_no_col].fillna('').astype(str) + ' ' +
                df_rta[rta_street_col].fillna('')
            ).apply(normalize)
            df_rta['_street_canon'] = df_rta['_street'].apply(lambda s: apply_canonical(s, canonical_map))
            df_rta['_pc'] = df_rta[rta_pc_col].fillna('').apply(norm_pc)

            # Build key variants
            for df in [df_hub, df_rta]:
                df['_k_exact']     = df['_street']       + '|' + df['_pc']
                df['_k_dir']       = df['_street'].apply(strip_direction) + '|' + df['_pc']
                df['_k_canon']     = df['_street_canon']  + '|' + df['_pc']
                df['_k_canon_dir'] = df['_street_canon'].apply(strip_direction) + '|' + df['_pc']

            # Build lookups: key -> (rta_full_address, rta_status)
            # We need to return TWO columns, so build a combined lookup
            df_rta['_lookup_val'] = df_rta['_rta_full'] + '|||' + df_rta[rta_status_col].fillna('').astype(str)

            lookups = {}
            for key_col in ['_k_exact', '_k_dir', '_k_canon', '_k_canon_dir']:
                lookups[key_col] = df_rta.drop_duplicates(subset=key_col).set_index(key_col)['_lookup_val']

            # Initialize output columns
            df_hub['RTA Address'] = pd.Series(dtype='object')
            df_hub['RTA Status'] = pd.Series(dtype='object')
            df_hub['_match_type'] = ''

            passes = [
                ('_k_exact',     'exact'),
                ('_k_dir',       'direction_strip'),
                ('_k_canon',     'fuzzy'),
                ('_k_canon_dir', 'fuzzy'),
            ]

            for key_col, mtype in passes:
                unmatched = df_hub['RTA Address'].isna()
                mapped = df_hub.loc[unmatched, key_col].map(lookups[key_col])
                matched_mask = mapped.notna()
                if matched_mask.any():
                    split = mapped[matched_mask].str.split('|||', n=1, expand=True)
                    df_hub.loc[split.index, 'RTA Address'] = split[0].values
                    df_hub.loc[split.index, 'RTA Status'] = split[1].values
                    newly_matched = unmatched & df_hub['RTA Address'].notna() & (df_hub['_match_type'] == '')
                    df_hub.loc[newly_matched, '_match_type'] = mtype

            # Pass 5: street-only (no postal code) → RED
            r_lookup_street = {}
            r_lookup_street_stripped = {}
            for i in range(len(df_rta)):
                val = df_rta.iloc[i]['_lookup_val']
                for st_key in [df_rta.iloc[i]['_street'], df_rta.iloc[i]['_street_canon']]:
                    if st_key and st_key not in r_lookup_street:
                        r_lookup_street[st_key] = val
                for st_key in [strip_direction(df_rta.iloc[i]['_street']), strip_direction(df_rta.iloc[i]['_street_canon'])]:
                    if st_key and st_key not in r_lookup_street_stripped:
                        r_lookup_street_stripped[st_key] = val

            unmatched = df_hub['RTA Address'].isna()
            for idx in df_hub[unmatched].index:
                h_st = df_hub.loc[idx, '_street']
                h_st_canon = df_hub.loc[idx, '_street_canon']
                val = (r_lookup_street.get(h_st) or r_lookup_street.get(h_st_canon)
                       or r_lookup_street_stripped.get(strip_direction(h_st))
                       or r_lookup_street_stripped.get(strip_direction(h_st_canon)))
                if val:
                    parts = val.split('|||', 1)
                    df_hub.loc[idx, 'RTA Address'] = parts[0]
                    df_hub.loc[idx, 'RTA Status'] = parts[1] if len(parts) > 1 else ''
                    df_hub.loc[idx, '_match_type'] = 'no_pc'

            # Stats
            exact_count = (df_hub['_match_type'] == 'exact').sum()
            yellow_count = df_hub['_match_type'].isin(['fuzzy', 'direction_strip']).sum()
            red_count = (df_hub['_match_type'] == 'no_pc').sum()
            total_matched = df_hub['RTA Address'].notna().sum()
            total_rows = len(df_hub)

            st.markdown("---")
            st.subheader("Results")

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Total rows", total_rows)
            m2.metric("Total matched", total_matched)
            m3.metric("Exact (white)", exact_count)
            m4.metric("Fuzzy (yellow)", yellow_count)
            m5.metric("Risky (red)", red_count)

            # Show special matches
            special = df_hub[df_hub['_match_type'].isin(['fuzzy', 'direction_strip', 'no_pc'])][
                [hub_street_col, hub_pc_col, 'RTA Address', 'RTA Status', '_match_type']
            ].copy()
            special.columns = ['Street Address', 'Postal Code', 'RTA Address', 'RTA Status', 'Match Type']

            if len(special) > 0:
                st.markdown("**Flagged matches for review:**")

                def highlight_match_type(row):
                    if row['Match Type'] == 'no_pc':
                        return ['background-color: #FF6666'] * len(row)
                    else:
                        return ['background-color: #FFFF00'] * len(row)

                st.dataframe(special.style.apply(highlight_match_type, axis=1), use_container_width=True)

            # Preview output
            preview = df_hub[[hub_street_col, hub_pc_col, 'RTA Address', 'RTA Status']].head(20)
            st.markdown("**Output preview (first 20 rows):**")
            st.dataframe(preview, use_container_width=True)

            # Save to Excel with colors
            match_type = df_hub['_match_type'].copy()
            df_out = df_hub.drop(columns=[c for c in df_hub.columns if c.startswith('_')])

            buffer = io.BytesIO()
            df_out.to_excel(buffer, index=False, engine='openpyxl')
            buffer.seek(0)

            wb = load_workbook(buffer)
            ws = wb.active

            # Find RTA Address column index
            rta_addr_col_idx = None
            rta_status_col_idx = None
            for cell in ws[1]:
                if cell.value == 'RTA Address':
                    rta_addr_col_idx = cell.column
                elif cell.value == 'RTA Status':
                    rta_status_col_idx = cell.column

            yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
            red_fill = PatternFill(start_color='FF6666', end_color='FF6666', fill_type='solid')

            for i, mt in enumerate(match_type):
                if mt in ('fuzzy', 'direction_strip'):
                    fill = yellow_fill
                elif mt == 'no_pc':
                    fill = red_fill
                else:
                    continue
                if rta_addr_col_idx:
                    ws.cell(row=i+2, column=rta_addr_col_idx).fill = fill
                if rta_status_col_idx:
                    ws.cell(row=i+2, column=rta_status_col_idx).fill = fill

            out_buffer = io.BytesIO()
            wb.save(out_buffer)
            out_buffer.seek(0)

            st.download_button(
                label="📥 Download color-coded Excel",
                data=out_buffer,
                file_name="hubspot_matched_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            st.markdown("""
            ---
            **Color legend:**
            - ⬜ **White** — Exact match (street + postal code)
            - 🟨 **Yellow** — Fuzzy match (name alias, direction stripped, spelling variant) — same postal code area
            - 🟥 **Red** — Street matched but postal codes differ — manual verification needed
            """)

elif hub_file or rta_file:
    st.info("Please upload both files to proceed.")
