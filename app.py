import os
import json
import streamlit as st
import pandas as pd
from PIL import Image

# Page Configuration
st.set_page_config(
    page_title="Sabzi Mandi to Table — Omnichannel Vegetable Business Platform",
    page_icon="🥦",
    layout="wide",
    initial_sidebar_state="expanded"
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'Vegetable_Omnichannel_Business_Model.xlsx')
ORG_IMG_DIR = os.path.join(BASE_DIR, 'organized_shops')
AUTH_CONFIG_FILE = os.path.join(BASE_DIR, 'auth_config.json')

# Authentication Configuration & Helpers
DEFAULT_USERS = {
    "admin": "admin123",
    "manager": "mandi2024"
}

def load_auth_config():
    if not os.path.exists(AUTH_CONFIG_FILE):
        try:
            with open(AUTH_CONFIG_FILE, 'w') as f:
                json.dump({"users": DEFAULT_USERS}, f, indent=4)
        except Exception:
            pass
        return {"users": DEFAULT_USERS}
    try:
        with open(AUTH_CONFIG_FILE, 'r') as f:
            return json.load(f)
    except Exception:
        return {"users": DEFAULT_USERS}

def save_auth_config(config):
    with open(AUTH_CONFIG_FILE, 'w') as f:
        json.dump(config, f, indent=4)

def verify_credentials(username, password):
    if not password:
        return False, None
    # Environment variable override if specified
    env_pass = os.environ.get("APP_PASSWORD") or os.environ.get("ADMIN_PASSWORD")
    if env_pass and password == env_pass:
        return True, username.strip() if username.strip() else "admin"

    config = load_auth_config()
    users = config.get("users", DEFAULT_USERS)
    u_clean = username.strip() if username else ""

    if u_clean:
        for u, p in users.items():
            if u.lower() == u_clean.lower() and str(p) == str(password):
                return True, u
    else:
        if str(users.get("admin")) == str(password):
            return True, "admin"
        for u, p in users.items():
            if str(p) == str(password):
                return True, u

    return False, None

# Session State Initialization
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "username" not in st.session_state:
    st.session_state["username"] = ""

# Custom CSS for polished green theme
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 800;
        color: #1E4620;
        margin-bottom: 0px;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #2D6A4F;
        margin-bottom: 20px;
    }
    .metric-card {
        background: linear-gradient(135deg, #F4F9F4, #E8F5E9);
        border: 1px solid #C8E6C9;
        border-radius: 10px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 2px 5px rgba(0,0,0,0.04);
    }
    .vendor-card {
        background: #ffffff;
        border: 1px solid #E0E0E0;
        border-radius: 12px;
        padding: 18px;
        margin-bottom: 20px;
        box-shadow: 0 3px 8px rgba(0,0,0,0.06);
        transition: transform 0.2s ease-in-out;
    }
    .vendor-card:hover {
        transform: translateY(-2px);
        border-color: #2D6A4F;
    }
    .badge-shop {
        background-color: #1E4620;
        color: white;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 0.85rem;
        display: inline-block;
    }
    .badge-cat {
        background-color: #E8F5E9;
        color: #2D6A4F;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.85rem;
        display: inline-block;
        margin-left: 6px;
    }
    .badge-lic {
        background-color: #FFF3E0;
        color: #E65100;
        padding: 4px 8px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.8rem;
        display: inline-block;
        margin-left: 6px;
    }
    .sop-box {
        background-color: #F8FAF8;
        border-left: 4px solid #2D6A4F;
        padding: 10px 14px;
        border-radius: 0 8px 8px 0;
        margin-top: 8px;
        font-size: 0.92rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #F1F8F3;
        border-radius: 8px 8px 0 0;
        padding: 10px 18px;
        font-weight: 600;
        color: #2D6A4F;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1E4620 !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# AUTHENTICATION GATE (HOME SCREEN LOGIN)
# -------------------------------------------------------------
if not st.session_state.get("authenticated", False):
    # Minimal locked sidebar
    with st.sidebar:
        st.image("https://img.icons8.com/fluency/96/vegetables.png", width=70)
        st.markdown("### **Sabzi Mandi Platform**")
        st.markdown("**Omnichannel Operating Model**")
        st.markdown("---")
        st.info("🔒 **Access Protected**\nPlease enter your login password on the home screen to unlock the platform.")
        st.markdown("---")
        st.caption("📍 Anchored at New Sabzi Mandi, Ghazipur, Delhi-110096")

    # Centered login form on home screen
    col_l1, col_l2, col_l3 = st.columns([1, 1.8, 1])
    with col_l2:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #F4F9F4 0%, #E8F5E9 100%);
                    border: 2px solid #2D6A4F;
                    border-radius: 16px;
                    padding: 28px 24px 20px 24px;
                    text-align: center;
                    margin-top: 25px;
                    margin-bottom: 25px;
                    box-shadow: 0 6px 18px rgba(0,0,0,0.06);">
            <div style="font-size: 3.2rem; line-height: 1.1; margin-bottom: 6px;">🥦</div>
            <h2 style="color: #1E4620; margin: 0 0 4px 0; font-weight: 800; font-size: 1.8rem;">Sabzi Mandi to Table</h2>
            <p style="color: #2D6A4F; font-size: 0.95rem; font-weight: 600; margin: 0 0 14px 0;">Omnichannel Vegetable Business Platform</p>
            <span style="background-color: #1E4620; color: white; padding: 5px 14px; border-radius: 20px; font-size: 0.82rem; font-weight: 600; letter-spacing: 0.5px;">
                🔒 HOME SCREEN LOGIN REQUIRED
            </span>
        </div>
        """, unsafe_allow_html=True)

        with st.form("home_screen_login_form"):
            st.markdown("#### 🔐 **Enter Credentials**")
            login_user = st.text_input("Username", value="admin", placeholder="e.g. admin", help="Enter username (default: admin)")
            login_pass = st.text_input("Password", type="password", placeholder="Enter your password", help="Enter login password")
            
            submit_login = st.form_submit_button("🔓 Log In to Platform", use_container_width=True)
            if submit_login:
                if not login_pass:
                    st.warning("⚠️ Please enter a password.")
                else:
                    is_valid, user_name = verify_credentials(login_user, login_pass)
                    if is_valid:
                        st.session_state["authenticated"] = True
                        st.session_state["username"] = user_name
                        st.success(f"✅ Access granted! Welcome, **{user_name}**.")
                        st.rerun()
                    else:
                        st.error("❌ Invalid password or username. Please check and try again.")

        st.markdown("""
        <div style="background-color: #FFFFFF; border: 1px dashed #81C784; border-radius: 10px; padding: 14px 18px; margin-top: 15px; box-shadow: 0 2px 6px rgba(0,0,0,0.03);">
            <p style="margin: 0; font-size: 0.88rem; color: #1E4620;">
                🔑 <strong>Default Login Credentials:</strong><br>
                • <strong>Username:</strong> <code>admin</code><br>
                • <strong>Password:</strong> <code>admin123</code><br>
                <small style="color: #4B6B4E;">(You can log in directly by keeping username as <code>admin</code> and entering <code>admin123</code>)</small>
            </p>
        </div>
        """, unsafe_allow_html=True)

    # Halt execution so protected data and tabs are not loaded or shown
    st.stop()

# Helper function to load data
@st.cache_data
def load_vendor_data():
    df = pd.read_excel(EXCEL_PATH, sheet_name='Vendor_Directory', skiprows=2)
    return df

@st.cache_data
def load_pricing_data():
    df = pd.read_excel(EXCEL_PATH, sheet_name='Unit_Economics_&_Pricing', skiprows=2)
    # only rows 0 to 24 are the 25 SKUs
    skus_df = df.iloc[:25].copy()
    return skus_df

vendors_df = load_vendor_data()
pricing_df = load_pricing_data()

# Sidebar for Authenticated Users
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/vegetables.png", width=70)
    st.markdown("### **Sabzi Mandi Platform**")
    st.markdown("**Omnichannel Operating Model**")
    st.markdown("---")
    
    # User Profile & Logout
    st.markdown(f"""
    <div style="background: #E8F5E9; border: 1px solid #C8E6C9; border-radius: 8px; padding: 10px 12px; margin-bottom: 10px;">
        <span style="font-size: 0.78rem; color: #2D6A4F; font-weight: 700; text-transform: uppercase;">👤 Active Session</span><br>
        <strong style="color: #1E4620; font-size: 1.05rem;">{st.session_state.get('username', 'Admin')}</strong>
    </div>
    """, unsafe_allow_html=True)
    
    if st.button("🚪 Sign Out", key="sidebar_logout_btn", use_container_width=True):
        st.session_state["authenticated"] = False
        st.session_state["username"] = ""
        st.rerun()

    st.markdown("---")
    st.markdown("#### 📊 **Key System Metrics**")
    st.metric(label="Verified Mandi Shops", value=len(vendors_df), delta="58 Signboards")
    st.metric(label="Modeled Produce SKUs", value=len(pricing_df), delta="Daily Staples & Exotics")
    st.metric(label="Target Shrinkage %", value="< 4.5%", delta="-10% vs Market", delta_color="inverse")
    st.metric(label="Projected Monthly Net", value="₹2,42,781", delta="19.6% EBITDA")
    
    st.markdown("---")
    st.markdown("#### 📥 **Download Model**")
    with open(EXCEL_PATH, "rb") as f:
        st.download_button(
            label="📄 Download Master Excel File",
            data=f,
            file_name="Vegetable_Omnichannel_Business_Model.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.markdown("---")
    with st.expander("⚙️ Account & Change Password"):
        with st.form("change_password_form"):
            st.markdown("<small><strong>Update Login Password</strong></small>", unsafe_allow_html=True)
            curr_p = st.text_input("Current Password", type="password", key="cp_curr")
            new_p = st.text_input("New Password", type="password", key="cp_new")
            conf_p = st.text_input("Confirm New Password", type="password", key="cp_conf")
            change_submitted = st.form_submit_button("Update Password", use_container_width=True)
            if change_submitted:
                cfg = load_auth_config()
                curr_u = st.session_state.get("username", "admin")
                curr_stored = cfg.get("users", {}).get(curr_u, DEFAULT_USERS.get(curr_u, ""))
                if curr_p != curr_stored:
                    st.error("❌ Current password does not match.")
                elif not new_p:
                    st.error("❌ New password cannot be empty.")
                elif new_p != conf_p:
                    st.error("❌ New passwords do not match.")
                else:
                    cfg.setdefault("users", {})[curr_u] = new_p
                    save_auth_config(cfg)
                    st.success("✅ Password updated successfully!")

    st.markdown("---")
    st.caption("📍 Anchored at New Sabzi Mandi, Ghazipur, Delhi-110096")

# Main Content Header for Authenticated Users
col_h1, col_h2, col_h3 = st.columns([3.5, 0.9, 0.9])
with col_h1:
    st.markdown('<div class="main-header">🥦 Sabzi Mandi to Table — Operating Platform</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Ghazipur Wholesale Sourcing • Offline Retail Blueprint • Hyperlocal WhatsApp Delivery • Dynamic P&L</div>', unsafe_allow_html=True)
with col_h2:
    st.markdown("<br>", unsafe_allow_html=True)
    st.button("🔄 Refresh Data", on_click=st.cache_data.clear, use_container_width=True)
with col_h3:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Sign Out", key="top_logout_btn", use_container_width=True):
        st.session_state["authenticated"] = False
        st.session_state["username"] = ""
        st.rerun()

# Tabs
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🏪 Mandi Shop Directory",
    "🚚 Procurement & Mandi SOP",
    "🏬 Offline Store Operations",
    "📱 Online WhatsApp Delivery",
    "💰 Dynamic P&L & Pricing",
    "🗓️ 30-60-90 Day Roadmap"
])

# -------------------------------------------------------------
# TAB 1: MANDI SHOP DIRECTORY
# -------------------------------------------------------------
with tab1:
    st.markdown("### 🏪 Verified Ghazipur APMC Mandi Vendor Directory")
    st.markdown("Browse verified wholesale shops extracted directly from signboards at New Sabzi Mandi Ghazipur with phone numbers, photos, procurement guidelines, and selling strategies.")
    
    col_f1, col_f2, col_f3 = st.columns([2, 2, 2])
    with col_f1:
        categories = ["All Categories"] + sorted(vendors_df['Produce Category'].dropna().unique().tolist())
        sel_cat = st.selectbox("Filter by Category", categories)
    with col_f2:
        search_query = st.text_input("🔍 Search by Shop No, Name or Keyword", "")
    with col_f3:
        blocks = ["All Blocks"] + sorted(vendors_df['Block'].dropna().unique().tolist())
        sel_block = st.selectbox("Filter by Mandi Block", blocks)
        
    filtered_df = vendors_df.copy()
    if sel_cat != "All Categories":
        filtered_df = filtered_df[filtered_df['Produce Category'] == sel_cat]
    if sel_block != "All Blocks":
        filtered_df = filtered_df[filtered_df['Block'] == sel_block]
    if search_query:
        q = search_query.lower()
        filtered_df = filtered_df[
            filtered_df['Shop No'].astype(str).str.lower().str.contains(q) |
            filtered_df['Trade / Firm Name'].astype(str).str.lower().str.contains(q) |
            filtered_df['Proprietor / Aadhati Name'].astype(str).str.lower().str.contains(q) |
            filtered_df['Specialist Produce / Varieties'].astype(str).str.lower().str.contains(q)
        ]
        
    st.markdown(f"**Showing {len(filtered_df)} verified shops matching criteria**")
    
    # Grid display
    for idx, row in filtered_df.iterrows():
        st.markdown(f"""
        <div class="vendor-card">
            <span class="badge-shop">{row['Shop No']}</span>
            <span class="badge-cat">{row['Produce Category']}</span>
            <span class="badge-lic">{row['APMC License No']}</span>
            <span style="float:right; font-weight:600; color:#666;">{row['Block']}</span>
            <h3 style="margin-top:10px; margin-bottom:4px; color:#1E4620;">{row['Trade / Firm Name']}</h3>
            <p style="margin-bottom:6px; color:#444;"><strong>Aadhati / Contacts:</strong> {row['Proprietor / Aadhati Name']}</p>
        </div>
        """, unsafe_allow_html=True)
        
        c_img, c_details = st.columns([1, 2])
        
        with c_img:
            rel_photo = str(row['Organized Photo Reference'])
            img_path = os.path.join(ORG_IMG_DIR, rel_photo)
            if os.path.exists(img_path):
                img = Image.open(img_path)
                st.image(img, caption=f"Shop {row['Shop No']} Signboard", use_column_width=True)
            else:
                st.info("Photo available in original archive.")
                
        with c_details:
            p1 = str(row['Primary Phone']).replace('.0', '')
            p2 = str(row['Alternate Phone'])
            
            call_link = f"tel:+91{p1}" if len(p1) >= 10 else "#"
            wa_link = f"https://wa.me/91{p1}" if len(p1) >= 10 else "#"
            
            st.markdown(f"""
            - 📞 **Primary Phone:** `{p1}` | 📱 **Alternate:** `{p2}`
            - 📦 **Specialist Commodities:** {row['Specialist Produce / Varieties']}
            - ⚖️ **Standard Packing Unit:** `{row['Standard Packaging Unit']}`
            """)
            
            c_btn1, c_btn2 = st.columns(2)
            with c_btn1:
                st.markdown(f'<a href="{call_link}" target="_blank" style="display:inline-block; width:100%; text-align:center; background-color:#2D6A4F; color:white; padding:8px 12px; border-radius:6px; text-decoration:none; font-weight:600;">📞 Call Vendor ({p1})</a>', unsafe_allow_html=True)
            with c_btn2:
                st.markdown(f'<a href="{wa_link}" target="_blank" style="display:inline-block; width:100%; text-align:center; background-color:#25D366; color:white; padding:8px 12px; border-radius:6px; text-decoration:none; font-weight:600;">💬 WhatsApp Order</a>', unsafe_allow_html=True)
            
            with st.expander("🛒 How We Buy From Them (Procurement SOP)", expanded=False):
                st.markdown(f'<div class="sop-box"><strong>Procurement Checklist:</strong><br>{row["How We Buy From Them (Procurement SOP)"]}</div>', unsafe_allow_html=True)
                
            with st.expander("🏪 How We Sell In Store & Online (Retail Strategy)", expanded=False):
                st.markdown(f'<div class="sop-box"><strong>Selling Strategy:</strong><br>{row["How We Sell In Store & Online (Retail Strategy)"]}</div>', unsafe_allow_html=True)
                
        st.markdown("<hr style='margin-top:10px; margin-bottom:20px; border-top:1px dashed #DDD;'>", unsafe_allow_html=True)

# -------------------------------------------------------------
# TAB 2: PROCUREMENT & MANDI SOP
# -------------------------------------------------------------
with tab2:
    st.markdown("### 🚚 Ghazipur Subzi Mandi — Daily Procurement SOP & Calculator")
    
    st.markdown("""
    Daily procurement operations require strict adherence to auction arrival windows, statutory market fees, crate float balances, and physical defect inspection.
    """)
    
    col_p1, col_p2 = st.columns([1, 1])
    
    with col_p1:
        st.markdown("#### ⏰ **Daily Mandi Procurement Timeline**")
        timeline_items = [
            ("03:00 AM - 03:30 AM", "Transit Dispatch", "Tata Ace vehicle departs store dock with 40 clean plastic crates & driver allowance."),
            ("03:30 AM - 04:30 AM", "Auction Phase 1: High Perishables", "Herbs & Tomatoes sourced from Shop B-211 (Dhaniya Wale), B-81 (Haji Habib), B-105. Check turgidity & firmness."),
            ("04:30 AM - 05:15 AM", "Auction Phase 2: Gourds & Exotics", "Bhindi (Vedram Baba B-242 snap test), Button Mushrooms (JMD A-171 cold box), Capsicum (SRTC B-125)."),
            ("05:15 AM - 06:15 AM", "Auction Phase 3: Brassicas & Tubers", "Cauliflower (B-180), Sambhal Potatoes (B-56 truck tailgate), Nashik Onions (B-17)."),
            ("06:15 AM - 07:00 AM", "Consolidation & Palledari Loading", "Tubers on vehicle floor, delicate greens on top; settle palledari (₹5/crate); exchange empty crates."),
            ("07:00 AM - 07:45 AM", "Return Transit & Intake Shelving", "Unload at retail store; weigh verification; mist greens; tag prices before 8:00 AM opening.")
        ]
        for t_time, t_title, t_desc in timeline_items:
            st.markdown(f"""
            <div style="background:#F9FBF9; border-left:4px solid #1E4620; padding:10px; margin-bottom:10px; border-radius:0 6px 6px 0;">
                <span style="font-weight:700; color:#1E4620;">{t_time}</span> — <strong>{t_title}</strong><br>
                <span style="font-size:0.9rem; color:#555;">{t_desc}</span>
            </div>
            """, unsafe_allow_html=True)
            
    with col_p2:
        st.markdown("#### 🧮 **Interactive Daily Mandi Cost & Fee Calculator**")
        st.markdown("Simulate true landed cost including all mandi statutory and handling surcharges:")
        
        mandi_gross = st.slider("Estimated Daily Mandi Gross Purchases (₹)", 10000, 50000, 25000, step=1000)
        crates_count = st.slider("Plastic Crates Sourced (25kg each)", 10, 40, 22)
        sacks_count = st.slider("50kg Jute Sacks Sourced (Potato/Onion)", 5, 25, 12)
        vehicle_freight = st.number_input("Vehicle Freight per Run (₹)", value=1000, step=100)
        
        # Calculations
        apmc_fee = mandi_gross * 0.0125 # 1.25%
        aadhati_dami = mandi_gross * 0.055 # 5.5%
        palledari_crates = crates_count * 5.0 # ₹5 per crate
        palledari_sacks = sacks_count * 8.0 # ₹8 per sack
        total_palledari = palledari_crates + palledari_sacks
        crate_deposit_float = crates_count * 50.0 # ₹50 refundable
        
        total_outbound_cash = mandi_gross + apmc_fee + aadhati_dami + total_palledari + vehicle_freight + crate_deposit_float
        net_landed_inventory_cost = mandi_gross + apmc_fee + aadhati_dami + total_palledari + vehicle_freight
        effective_overhead_pct = ((net_landed_inventory_cost - mandi_gross) / mandi_gross) * 100
        
        st.markdown(f"""
        <div style="background:#E8F5E9; border:1px solid #A5D6A7; border-radius:10px; padding:15px; margin-top:10px;">
            <h4 style="margin:0; color:#1E4620;">Procurement Cost Summary</h4>
            <table style="width:100%; margin-top:10px; font-size:0.95rem;">
                <tr><td>Raw Auction Purchases:</td><td style="text-align:right; font-weight:600;">₹{mandi_gross:,.2f}</td></tr>
                <tr><td>APMC Cess (1.25%):</td><td style="text-align:right;">₹{apmc_fee:,.2f}</td></tr>
                <tr><td>Aadhati Dami (5.5%):</td><td style="text-align:right;">₹{aadhati_dami:,.2f}</td></tr>
                <tr><td>Palledari (Loading/Unloading):</td><td style="text-align:right;">₹{total_palledari:,.2f}</td></tr>
                <tr><td>Inbound Freight (Tata Ace):</td><td style="text-align:right;">₹{vehicle_freight:,.2f}</td></tr>
                <tr style="border-top:1px solid #81C784; font-weight:700; color:#1E4620;">
                    <td>True Landed Cost of Produce:</td><td style="text-align:right;">₹{net_landed_inventory_cost:,.2f}</td>
                </tr>
                <tr><td>Refundable Crate Security Float:</td><td style="text-align:right; color:#666;">₹{crate_deposit_float:,.2f}</td></tr>
                <tr style="border-top:2px solid #1E4620; font-weight:800; font-size:1.05rem; color:#1E4620;">
                    <td>Total Cash Needed for Mandi Run:</td><td style="text-align:right;">₹{total_outbound_cash:,.2f}</td>
                </tr>
            </table>
            <p style="margin-top:10px; margin-bottom:0; font-size:0.88rem; color:#2E7D32;">
                <strong>Effective Inbound Overhead:</strong> +{effective_overhead_pct:.1f}% on auction prices.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown("---")
    st.markdown("#### 🔬 **Commodity Quality Defect Acceptance Matrix**")
    defects_data = [
        {"Category": "Tomatoes", "Grade A Acceptance": "Firm, 70-80% red, intact calyx, 50-60mm", "Zero Tolerance Defects": "Blossom end rot, squishing, puncture holes", "Action": "Reject crate or negotiate 25% discount"},
        {"Category": "Potatoes", "Grade A Acceptance": "Smooth unwashed skin, dormant eyes, 45-60mm", "Zero Tolerance Defects": "Green sunburn (solanine), watery rot, sprouts", "Action": "Refuse wet-bottom or sprouted sacks"},
        {"Category": "Onions", "Grade A Acceptance": "Dry crackling skin, tight necks, deep purple", "Zero Tolerance Defects": "Wet spongy neck (neck rot), black mold, sprouts", "Action": "Slice open 2 test onions per 3 sacks"},
        {"Category": "Leafy Greens", "Grade A Acceptance": "Crisp turgid stems, vibrant green leaves", "Zero Tolerance Defects": "Yellowing leaves, slimy root rot, fermentation smell", "Action": "Reject fermented or yellowed bundles"},
        {"Category": "Okra (Bhindi)", "Grade A Acceptance": "2-3 inch bright green, clean tip snap", "Zero Tolerance Defects": "Tough bending tips (over-mature), borer holes", "Action": "Snap 5 sample pods; if >1 bends, downgrade"},
        {"Category": "Mushrooms", "Grade A Acceptance": "Clean closed white button caps, tight veil", "Zero Tolerance Defects": "Dark open gills, slime patches, condensation", "Action": "Reject punnets with black gill spores"}
    ]
    st.dataframe(pd.DataFrame(defects_data), use_container_width=True)

# -------------------------------------------------------------
# TAB 3: OFFLINE STORE OPERATIONS
# -------------------------------------------------------------
with tab3:
    st.markdown("### 🏬 Physical Retail Store Blueprint & Zero-Wastage Model")
    
    col_s1, col_s2 = st.columns([1, 1])
    
    with col_s1:
        st.markdown("#### 📐 **Store Merchandising Architecture (350–500 sq ft)**")
        st.markdown("""
        ```
        ┌────────────────────────────────────────────────────────┐
        │ [ ENTRANCE ]   Seasonal Promo Island (Watermelon/Mango)│
        ├──────────────────────────┬─────────────────────────────┤
        │                          │                             │
        │   HYDRATION GREEN WALL   │   ALLIUM & TUBER BUNKERS    │
        │   (Ultrasonic Misters)   │   (Sloped Wooden Bins)      │
        │   • Dhaniya, Pudina      │   • Potatoes (Sambhal/Agra) │
        │   • Palak, Methi, Salads │   • Onions (Nashik Red)     │
        │   • English Cucumbers    │   • Garlic & Ginger Roots   │
        │                          │                             │
        ├──────────────────────────┴─────────────────────────────┤
        │                CENTER GONDOLA AISLE                    │
        │       Modular Crates on Astro-Turf (Color Block)       │
        │       Tomatoes • Capsicum • Bhindi • Gourds • Beans    │
        ├──────────────────────────┬─────────────────────────────┤
        │   COMMERCIAL CHILLER     │   CHECKOUT & WEIGHING POS   │
        │   (4°C - 8°C Glass Door) │   • Barcode Weighing Scale  │
        │   • Button Mushrooms     │   • Touch Billing Terminal  │
        │   • Sprouts, Paneer, Corn│   • Instant UPI QR Stand    │
        │   • Ready-to-Cook Packs  │   • Eco Paper Bag Packing   │
        └──────────────────────────┴─────────────────────────────┘
        ```
        """)
        
    with col_s2:
        st.markdown("#### ♻️ **4-Tier Zero-Wastage & Shrinkage Protocol**")
        st.markdown("Unorganized retail suffers 15–20% wastage. Our structured 4-tier model controls physical shrinkage to **under 4.5%**:")
        
        st.markdown("""
        - 🥇 **Tier 1: Fresh Premium Retail (75–80% Volume)**
          - Pristine Grade A produce washed, misted, and sold at full markup (**35%–45% Gross Margin**).
        - 🥈 **Tier 2: Ready-to-Cook Processing (10–12% Volume)**
          - Odd-shaped or superficially bruised items peeled, diced, sanitized, and vacuum-sealed in prep room (**40%–60% Value-Add Margin**).
          - *Examples:* Peeled garlic, diced pumpkin, chopped kathal, sambar mix vegetable kits.
        - 🥉 **Tier 3: 7:30 PM 'Happy Hour' Markdown (6–8% Volume)**
          - Delicate greens and ripe tomatoes sold at **25%–30% discount** after 7:30 PM to achieve near-zero overnight carryover.
        - ♻️ **Tier 4: B2B Liquidation & Compost (2–3% Volume)**
          - Over-ripe tomatoes liquidated to local cloud kitchens/street carts at cost; peels sent to organic composting. **Zero landfill waste!**
        """)
        
    st.markdown("---")
    st.markdown("#### ⏰ **Daily Hourly Store Operations Rhythm**")
    rhythm_df = pd.DataFrame([
        {"Time Window": "07:00 AM - 08:00 AM", "Phase": "Intake & Shelving", "Action": "Unload vehicle, verify weigh slips, mist greens, display Grade A, update POS prices"},
        {"Time Window": "08:00 AM - 12:30 PM", "Phase": "Peak Morning Rush", "Action": "High footfall (homemakers & seniors); continuous weighing; complimentary dhaniya-mirchi"},
        {"Time Window": "12:30 PM - 04:00 PM", "Phase": "Midday Maintenance & Online Prep", "Action": "Restock displays; second misting; trim greens; pack online pre-orders; prep Tier 2 packs"},
        {"Time Window": "04:00 PM - 07:30 PM", "Phase": "Peak Evening Rush", "Action": "Working professionals & dinner cooks; promote salad boxes, ready-to-cook meal kits"},
        {"Time Window": "07:30 PM - 09:15 PM", "Phase": "Happy Hour Clearance", "Action": "Activate 25-30% discount on delicate greens & ripe tomatoes; clear perishable stock"},
        {"Time Window": "09:15 PM - 10:00 PM", "Phase": "Close & Reconciliation", "Action": "Deep washdown; misting purge; weigh waste (< 4.5% target); transmit morning indent"}
    ])
    st.table(rhythm_df)

# -------------------------------------------------------------
# TAB 4: ONLINE DELIVERY MODEL
# -------------------------------------------------------------
with tab4:
    st.markdown("### 📱 Online Hyperlocal Delivery — Store-as-a-Hub Model")
    st.markdown("Expanding into online delivery using the physical store as a fulfillment micro-hub, avoiding redundant dark store leases.")
    
    col_w1, col_w2 = st.columns([1, 1])
    
    with col_w1:
        st.markdown("#### 📲 **Interactive WhatsApp Catalog Simulator**")
        st.markdown("Experience how residential apartment customers order via WhatsApp:")
        
        st.markdown("""
        <div style="background:#E5DDD5; border-radius:12px; padding:15px; border:1px solid #CCC; max-width:420px; font-family:sans-serif;">
            <div style="background:#075E54; color:white; padding:8px 12px; border-radius:8px 8px 0 0; font-weight:700;">
                🥦 Fresh Mandi Direct — WhatsApp Store
            </div>
            <div style="background:white; border-radius:8px; padding:10px; margin-top:8px; box-shadow:0 1px 2px rgba(0,0,0,0.1);">
                <p style="margin:0; font-size:0.9rem; color:#333;"><strong>Namaste Sharmaji! 🙏</strong><br>Fresh morning harvest arrived from Ghazipur Mandi at 7:30 AM!</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        sim_cart = {}
        st.markdown("**Select Customer Basket:**")
        c1, c2 = st.columns(2)
        with c1:
            q_pot = st.number_input("🥔 Potato (1kg pack) @ ₹29", 0, 5, 2)
            q_oni = st.number_input("🧅 Nashik Onion (1kg pack) @ ₹49", 0, 5, 1)
            q_tom = st.number_input("🍅 Hybrid Tomato (1kg pack) @ ₹48", 0, 5, 1)
        with c2:
            q_dha = st.number_input("🌿 Fresh Dhaniya (100g) @ ₹12", 0, 5, 2)
            q_bhi = st.number_input("🥒 Bhindi / Okra (500g) @ ₹34", 0, 5, 1)
            q_msh = st.number_input("🍄 Button Mushroom (200g) @ ₹60", 0, 5, 1)
            
        slot = st.selectbox("Preferred Delivery Slot", ["🌅 Morning Slot (7:30 AM - 9:30 AM)", "🌇 Evening Slot (5:30 PM - 7:30 PM)", "⚡ 60-Minute Express (₹35 fee)"])
        
        cart_total = (q_pot * 29) + (q_oni * 49) + (q_tom * 48) + (q_dha * 12) + (q_bhi * 34) + (q_msh * 60)
        delivery_fee = 0 if (cart_total >= 199 or "Express" in slot) else 25
        if "Express" in slot:
            delivery_fee = 35
        final_bill = cart_total + delivery_fee
        
        st.markdown(f"""
        <div style="background:#DCF8C6; border-radius:8px; padding:12px; margin-top:12px; font-size:0.92rem;">
            <strong>Order Receipt:</strong><br>
            • Subtotal: ₹{cart_total:.2f}<br>
            • Delivery Fee: {'FREE (Order >= ₹199)' if delivery_fee==0 else f'₹{delivery_fee:.2f}'}<br>
            <strong>Total Payable: ₹{final_bill:.2f}</strong><br>
            <span style="color:#075E54; font-weight:600;">Slot: {slot}</span>
        </div>
        """, unsafe_allow_html=True)
        st.button("📲 Send Order via WhatsApp (Simulate)")
        
    with col_w2:
        st.markdown("#### 📦 **Sustainable Packaging Specifications**")
        st.markdown("""
        We use 100% plastic-free, breathable packaging to maintain produce freshness:
        """)
        pack_specs = [
            ("Potatoes & Onions (1-5 kg)", "Micro-perforated brown Kraft paper sack", "₹2.50 / bag", "Maintains breathability, stops moisture buildup"),
            ("Tomatoes & Fruiting Veg", "Ventilated RPET clamshell punnets", "₹3.00 / punnet", "Stops squishing and skin puncture"),
            ("Fresh Herbs (Dhaniya/Mint)", "Parchment moisture-lock wrap", "₹1.50 / wrap", "Prevents leaf slime; keeps stems turgid"),
            ("Bhindi & Gourds (500g)", "Kraft pouch with glassine window", "₹2.20 / bag", "Allows visual inspection of freshness"),
            ("Mushrooms & Exotics", "Molded pulp tray with perforated film", "₹3.50 / tray", "Maintains 4°C-8°C cold chain protection")
        ]
        for p_name, p_mat, p_cost, p_ben in pack_specs:
            st.markdown(f"""
            <div style="background:#F4F8F4; border:1px solid #C8E6C9; border-radius:8px; padding:10px; margin-bottom:8px;">
                <strong>{p_name}</strong> — <span style="color:#1E4620; font-weight:600;">{p_cost}</span><br>
                <span style="font-size:0.9rem; color:#444;">Material: {p_mat}</span><br>
                <span style="font-size:0.85rem; color:#666;">Benefit: {p_ben}</span>
            </div>
            """, unsafe_allow_html=True)
            
        st.markdown("#### 🛵 **Delivery Fleet Economics**")
        st.markdown("""
        - **Vehicle Fleet:** 2 Electric Scooters (EVs) on lease (₹4,500/mo/scooter).
        - **Delivery Radius:** 3.5 km primary radius (expandable to 5 km).
        - **Batching:** 6 to 8 orders delivered per delivery run.
        - **Cost per Drop:** **₹24 to ₹28 per order** (significantly lower than on-demand 3rd party couriers).
        """)

# -------------------------------------------------------------
# TAB 5: DYNAMIC P&L & PRICING
# -------------------------------------------------------------
with tab5:
    st.markdown("### 💰 Dynamic SKU Pricing Engine & Monthly Store P&L")
    st.markdown("Simulate retail markups, mandi prices, shrinkage adjustments, and monthly financial performance.")
    
    st.markdown("#### ⚙️ **Financial Sensitivity Sliders**")
    c_p1, c_p2, c_p3 = st.columns(3)
    with c_p1:
        daily_volume_kg = st.slider("Daily Volume Handled (kg/day)", 400, 1500, 850, step=50)
        markup_modifier = st.slider("Retail Price Markup Adjustment", -15, +25, 0, step=5, format="%+d%%")
    with c_p2:
        store_rent = st.number_input("Monthly Store Rent (₹)", value=45000, step=5000)
        store_staff = st.number_input("Staff Salaries (4 Full-time) (₹)", value=62000, step=2000)
    with c_p3:
        vehicle_expense = st.number_input("Daily Mandi Vehicle (₹/mo)", value=30000, step=2000)
        elec_chiller = st.number_input("Electricity & Cold Storage (₹/mo)", value=18000, step=1000)
        
    # Calculate responsive financials
    base_avg_selling_price = 48.50 * (1.0 + (markup_modifier / 100.0))
    daily_gross_revenue = daily_volume_kg * base_avg_selling_price
    monthly_gross_revenue = daily_gross_revenue * 30.0
    
    blended_gross_margin_pct = 0.375 + (markup_modifier * 0.003)
    monthly_gross_profit = monthly_gross_revenue * blended_gross_margin_pct
    
    delivery_fleet_exp = 32000
    packaging_exp = (daily_volume_kg * 0.85) * 30.0 # ~₹25k
    mktg_software = 12000
    
    total_monthly_opex = store_rent + store_staff + vehicle_expense + elec_chiller + delivery_fleet_exp + packaging_exp + mktg_software
    net_operating_profit = monthly_gross_profit - total_monthly_opex
    net_profit_margin = (net_operating_profit / monthly_gross_revenue) * 100.0
    annual_net_profit = net_operating_profit * 12.0
    
    # Financial KPI Cards
    col_k1, col_k2, col_k3, col_k4 = st.columns(4)
    with col_k1:
        st.markdown(f"""
        <div class="metric-card">
            <span style="font-size:0.9rem; color:#666;">Daily Gross Revenue</span>
            <h2 style="margin:5px 0; color:#1E4620;">₹{daily_gross_revenue:,.0f}</h2>
            <span style="font-size:0.85rem; color:#2E7D32;">{daily_volume_kg} kg @ ₹{base_avg_selling_price:.1f}/kg</span>
        </div>
        """, unsafe_allow_html=True)
    with col_k2:
        st.markdown(f"""
        <div class="metric-card">
            <span style="font-size:0.9rem; color:#666;">Monthly Gross Profit</span>
            <h2 style="margin:5px 0; color:#1E4620;">₹{monthly_gross_profit:,.0f}</h2>
            <span style="font-size:0.85rem; color:#2E7D32;">{blended_gross_margin_pct*100:.1f}% Gross Margin</span>
        </div>
        """, unsafe_allow_html=True)
    with col_k3:
        st.markdown(f"""
        <div class="metric-card">
            <span style="font-size:0.9rem; color:#666;">Total Monthly OPEX</span>
            <h2 style="margin:5px 0; color:#D32F2F;">₹{total_monthly_opex:,.0f}</h2>
            <span style="font-size:0.85rem; color:#C62828;">Fixed & Operating Costs</span>
        </div>
        """, unsafe_allow_html=True)
    with col_k4:
        st.markdown(f"""
        <div class="metric-card" style="background:linear-gradient(135deg, #E8F5E9, #C8E6C9); border:2px solid #2D6A4F;">
            <span style="font-size:0.9rem; color:#1E4620; font-weight:700;">NET MONTHLY EBITDA</span>
            <h2 style="margin:5px 0; color:#1E4620;">₹{net_operating_profit:,.0f}</h2>
            <span style="font-size:0.85rem; color:#1B5E20; font-weight:700;">{net_profit_margin:.1f}% Net Margin!</span>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown("<br>", unsafe_allow_html=True)
    st.info(f"💡 **Projected Annual Net Profit:** **₹{annual_net_profit:,.0f} per year** at stabilized Month 3 run rate.")
    
    st.markdown("#### 📋 **25 Core SKU Economics Table**")
    st.dataframe(pricing_df, use_container_width=True)

# -------------------------------------------------------------
# TAB 6: ROADMAP 30-60-90
# -------------------------------------------------------------
with tab6:
    st.markdown("### 🗓️ 30-60-90 Day Phased Omnichannel Implementation Plan")
    
    phases = [
        {
            "Phase": "Phase 1: Days 1 - 30",
            "Title": "Offline Physical Store Foundation & Launch",
            "Budget": "₹3,60,000",
            "Target": "Footfall > 250 daily; Revenue > ₹25,000/day",
            "Tasks": [
                "Days 1-10: Mandi Onboarding — Visit Ghazipur at 4:00 AM; open accounts with B-238 (Deepak), B-81 (Haji Habib), B-211 (Naresh Kumar), B-17 (Shyam Sunder); contract Tata Ace vehicle driver.",
                "Days 11-20: Retail Lease & Fit-Out — Lease 350-500 sq ft high-street space; install ultrasonic hydration rack, wooden bunkers, commercial chiller, and barcode POS scale.",
                "Days 21-30: Soft Launch & Grand Opening — Hire 3 staff; distribute 1,500 flyers; launch 'Free 1kg Potato on ₹199' promo; register 100% customer phone numbers at checkout."
            ]
        },
        {
            "Phase": "Phase 2: Days 31 - 60",
            "Title": "Offline Stabilization & WhatsApp Hyperlocal Rollout",
            "Budget": "₹40,000",
            "Target": "Shrinkage < 4.5%; Online orders 25-35/day",
            "Tasks": [
                "Days 31-45: Shrinkage Calibration — Fine-tune 4-tier zero-wastage system; roll out 7:30 PM 'Happy Hour' 30% discount; launch Ready-to-Cook processed packs.",
                "Days 46-60: WhatsApp Delivery Rollout — Launch WhatsApp Business Catalog with daily fresh harvest broadcasts; onboard 1 dedicated EV scooter runner; start morning (7:30-9:30 AM) and evening delivery slots."
            ]
        },
        {
            "Phase": "Phase 3: Days 61 - 90",
            "Title": "Full Omnichannel Scale & Subscription Expansion",
            "Budget": "₹50,000",
            "Target": "Daily volume 850 kg; Net Profit > ₹2.4 Lakhs/month",
            "Tasks": [
                "Days 61-75: PWA Web Storefront & Subscriptions — Deploy progressive web app; launch Weekly Kitchen Box Subscriptions (Standard ₹499/week, Family ₹799/week); expand to 2 EV scooters.",
                "Days 76-90: Institutional B2B Supply & Scale — Onboard 5 local cafes/restaurants for daily vegetable supply; achieve break-even and sustainable 19.6% net EBITDA profit."
            ]
        }
    ]
    
    for p in phases:
        st.markdown(f"""
        <div style="background:#FFFFFF; border:1px solid #D0E1D4; border-radius:10px; padding:18px; margin-bottom:15px; box-shadow:0 2px 6px rgba(0,0,0,0.04);">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="background:#1E4620; color:white; padding:4px 10px; border-radius:6px; font-weight:700;">{p['Phase']}</span>
                <span style="font-weight:700; color:#2D6A4F;">Budget: {p['Budget']} | Target: {p['Target']}</span>
            </div>
            <h3 style="margin-top:10px; margin-bottom:8px; color:#1E4620;">{p['Title']}</h3>
            <ul style="margin-bottom:0; color:#333;">
                {''.join([f'<li style="margin-bottom:6px;">{t}</li>' for t in p['Tasks']])}
            </ul>
        </div>
        """, unsafe_allow_html=True)
