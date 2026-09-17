import streamlit as st
import requests
import pandas as pd
import plotly.express as px
from io import BytesIO

# ---------------------------------------------------------
# 1. Ρυθμίσεις Σελίδας & Design
# ---------------------------------------------------------
st.set_page_config(
    page_title="Metalstorm Stats Scout",
    page_icon="✈️",
    layout="wide"
)

# Custom CSS για πιο καθαρό & ευανάγνωστο UI
st.markdown("""
    <style>
    .stApp {
        background-color: #0f172a;
        color: #f8fafc;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.8rem !important;
        color: #38bdf8 !important;
    }
    .main-header {
        font-size: 2rem;
        font-weight: 700;
        color: #38bdf8;
        margin-bottom: 0.2rem;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">✈️ Metalstorm Scout — Stats Tracker</div>', unsafe_allow_html=True)
st.caption("Αναλυτική παρακολούθηση στατιστικών & εξαγωγή δεδομένων σε Excel")

# ---------------------------------------------------------
# 2. Sidebar - Εισαγωγή Startoken & Οδηγίες
# ---------------------------------------------------------
with st.sidebar:
    st.header("🔑 Σύνδεση")
    token_input = st.text_input("Εισάγετε το Startoken σας:", type="password", help="Επικολλήστε το token από το Network Tab")
    fetch_btn = st.button("🔄 Φόρτωση Στατιστικών", type="primary", use_container_width=True)
    
    st.divider()
    st.markdown("### ❓ Πώς να βρείτε το Token σας")
    st.markdown("""
    1. Ανοίξτε το **playmetalstorm.com/stats** στον PC browser.
    2. Πατήστε **F12** ➔ καρτέλα **Network**.
    3. Βρείτε το αίτημα **`my`**.
    4. Αντιγράψτε την τιμή του **`Authorization`** (μετά το `startoken `).
    """)

# ---------------------------------------------------------
# 3. Συνάρτηση Ανάκτησης Δεδομένων (API Call)
# ---------------------------------------------------------
def fetch_user_stats(token):
    url = "https://playmetalstorm.com/stats/users/my"
    clean_token = token.replace("startoken ", "").strip()
    
    headers = {
        "Authorization": f"startoken {clean_token}",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/json"
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            return response.json(), None
        elif response.status_code == 401:
            return None, "Μη έγκυρο ή ληγμένο Token."
        else:
            return None, f"Σφάλμα διακομιστή ({response.status_code})."
    except Exception as e:
        return None, f"Σφάλμα σύνδεσης: {str(e)}"

# ---------------------------------------------------------
# 4. Συνάρτηση Δημιουργίας Αρχείου Excel
# ---------------------------------------------------------
def generate_excel(player_info_df, planes_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        player_info_df.to_excel(writer, sheet_name='Προφίλ & Σύνολα', index=False)
        planes_df.to_excel(writer, sheet_name='Στατιστικά Αεροσκαφών', index=False)
    return output.getvalue()

# Execution
if fetch_btn and token_input:
    with st.spinner("Ανάκτηση δεδομένων..."):
        data, error = fetch_user_stats(token_input)
        if error:
            st.error(error)
        else:
            st.session_state['raw_data'] = data

# ---------------------------------------------------------
# 5. Επεξεργασία & Προβολή Στατιστικών
# ---------------------------------------------------------
if 'raw_data' in st.session_state:
    raw = st.session_state['raw_data']
    
    # Εντοπισμός του playerData στο JSON
    player_data = {}
    try:
        accounts = raw.get("accounts", {})
        first_account_id = list(accounts.keys())[0]
        player_data = accounts[first_account_id].get("playerData", {})
    except Exception:
        player_data = raw

    # Βασικά Στοιχεία
    player_name = player_data.get("playerName", "Unknown")
    friend_code = player_data.get("friendCode", "-")
    trophies = player_data.get("cachedTotalVehicleRankingPoints", 0)
    peak_trophies = player_data.get("highestTotalVehicleRankingPoints", 0)
    player_xp = player_data.get("playerXp", 0)
    
    # Επεξεργασία Αεροσκαφών
    career_counters = player_data.get("careerCounters", {}).get("countersByVehicleConfig", {})
    
    total_pvp_matches = 0
    total_pvp_wins = 0
    total_mvps = 0
    total_deaths = 0
    planes_list = []
    
    for plane_id, stats in career_counters.items():
        clean_name = plane_id.replace("aircraft:", "").replace("-v2", "").replace("-v3", "").replace("-", " ").title()
        
        matches = stats.get("pvpMatchesPlayed", 0)
        wins = stats.get("pvpMatchesWon", 0)
        deaths = stats.get("pvpDeaths", 0)
        score = stats.get("pvpScore", 0)
        mvps = stats.get("pvpMvpCount", 0)
        
        win_rate = round((wins / matches) * 100, 1) if matches > 0 else 0.0
        
        total_pvp_matches += matches
        total_pvp_wins += wins
        total_mvps += mvps
        total_deaths += deaths
        
        planes_list.append({
            "Αεροσκάφος": clean_name,
            "Αγώνες": matches,
            "Νίκες": wins,
            "Win Rate (%)": win_rate,
            "MVPs": mvps,
            "Θάνατοι": deaths,
            "Σκορ": score
        })
    
    overall_win_rate = round((total_pvp_wins / total_pvp_matches) * 100, 1) if total_pvp_matches > 0 else 0.0
    
    # --- HEADER & EXPORT BUTTON ---
    col_title, col_export = st.columns([3, 1])
    with col_title:
        st.subheader(f"👤 Παίκτης: **{player_name}** | Tag: `{friend_code}`")
    
    # Δημιουργία DataFrames για Εξαγωγή
    df_planes = pd.DataFrame(planes_list)
    df_planes_sorted = df_planes.sort_values(by="Αγώνες", ascending=False) if not df_planes.empty else df_planes
    
    df_profile = pd.DataFrame([{
        "Όνομα Παίκτη": player_name,
        "Friend Code": friend_code,
        "Τρόπαια": trophies,
        "Peak Τρόπαια": peak_trophies,
        "Player XP": player_xp,
        "Συνολικό Win Rate (%)": overall_win_rate,
        "Σύνολο Αγώνων": total_pvp_matches,
        "Σύνολο Νικών": total_pvp_wins,
        "Σύνολο MVPs": total_mvps,
        "Σύνολο Θανάτων": total_deaths
    }])
    
    # Κουμπί Κατεβάσματος Excel
    with col_export:
        excel_data = generate_excel(df_profile, df_planes_sorted)
        st.download_button(
            label="📥 Εξαγωγή σε Excel (.xlsx)",
            data=excel_data,
            file_name=f"Metalstorm_Stats_{player_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
    
    st.divider()

    # --- TOP CARDS (METRICS) ---
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Τρόπαια 🏆", f"{trophies:,}")
    c2.metric("Peak Τρόπαια 🔝", f"{peak_trophies:,}")
    c3.metric("Win Rate 🎯", f"{overall_win_rate}%")
    c4.metric("Αγώνες ⚔️", f"{total_pvp_matches:,}")
    c5.metric("MVPs 🎖️", f"{total_mvps:,}")
    
    st.divider()

    # --- GRAPH & TABLE ---
    if not df_planes.empty:
        # Φιλτράρισμα αεροσκαφών με >= 5 αγώνες για καθαρό γράφημα
        df_filtered = df_planes[df_planes["Αγώνες"] >= 5].sort_values(by="Win Rate (%)", ascending=False)
        
        st.subheader("📊 Win Rate ανά Μαχητικό (min. 5 αγώνες)")
        
        fig = px.bar(
            df_filtered,
            x="Αεροσκάφος",
            y="Win Rate (%)",
            text="Win Rate (%)",
            color="Win Rate (%)",
            color_continuous_scale="Blues",
            hover_data=["Αγώνες", "Νίκες", "MVPs"]
        )
        fig.update_traces(texttemplate='%{text}%', textposition='outside')
        fig.update_layout(template="plotly_dark", xaxis_tickangle=-45, height=450)
        st.plotly_chart(fig, use_container_width=True)
        
        st.subheader("📋 Αναλυτικά Στατιστικά Αεροσκαφών")
        st.dataframe(df_planes_sorted, use_container_width=True)