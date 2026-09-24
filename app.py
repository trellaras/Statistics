import streamlit as st
import requests
import pandas as pd
import plotly.express as px
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

# ---------------------------------------------------------
# 1. Ρυθμίσεις Σελίδας & Design
# ---------------------------------------------------------
st.set_page_config(
    page_title="Metalstorm Stats Scout",
    page_icon="✈️",
    layout="wide"
)

# Custom CSS
st.markdown("""
    <style>
    .stApp {
        background-color: #0f172a;
        color: #f8fafc;
    }
    div[data-testid="stMetric"] {
        background-color: #1e293b;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #334155;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.7rem !important;
        color: #38bdf8 !important;
    }
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #38bdf8;
        margin-bottom: 0.2rem;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">✈️ Metalstorm Scout — Lifetime Stats</div>', unsafe_allow_html=True)
st.caption("Πλήρης ανάλυση προφίλ, αεροσκαφών, Game Modes & μορφοποιημένη εξαγωγή Excel")

# ---------------------------------------------------------
# 2. Sidebar - Εισαγωγή Startoken & Οδηγίες
# ---------------------------------------------------------
with st.sidebar:
    st.header("🔑 Σύνδεση")
    token_input = st.text_input("Εισάγετε το Startoken σας:", type="password", help="Επικολλήστε το token")
    fetch_btn = st.button("🔄 Φόρτωση Στατιστικών", type="primary", use_container_width=True)
    
    st.divider()
    st.markdown("### ❓ Πώς να βρείτε το Token σας")
    
    tab_pc, tab_mobile = st.tabs(["💻 Υπολογιστής", "📱 Κινητό"])
    
    with tab_pc:
        st.markdown("""
        1. Ανοίξτε το **playmetalstorm.com/stats** στον Chrome/Edge.
        2. Πατήστε **F12** ➔ καρτέλα **Network**.
        3. Επιλέξτε το φίλτρο **Fetch/XHR**.
        4. Βρείτε το αίτημα **`my`**.
        5. Στο **Headers ➔ Authorization**, αντιγράψτε το κείμενο μετά το `startoken `.
        """)
        
    with tab_mobile:
        st.markdown("""
        **Android (Kiwi Browser):**
        1. Ανοίξτε το **playmetalstorm.com/stats** στο Kiwi Browser.
        2. Πατήστε 3 τελείες ➔ **Developer Tools ➔ Network**.
        3. Βρείτε το `my` και αντιγράψτε το token.

        **iPhone (iOS):**
        * Χρησιμοποιήστε εφαρμογή όπως το **Web Inspector** για να δείτε τα Network Headers.
        """)

# ---------------------------------------------------------
# 3. Συνάρτηση Ανάκτησης Δεδομένων
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
# 4. Συνάρτηση Δημιουργίας Styled Excel
# ---------------------------------------------------------
def generate_styled_excel(df_profile, df_planes, df_modes):
    output = BytesIO()
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Αφαίρεση default sheet
    
    # Στυλ Επικεφαλίδων & Δεδομένων
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Σκούρο Μπλε
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    
    # Στυλ Γραμμής Totals
    summary_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid") # Ανοιχτό Γκρι/Σλατ
    summary_font = Font(name="Calibri", size=10, bold=True)
    
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    
    totals_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='1E3A8A'),
        bottom=Side(style='double', color='1E3A8A')
    )

    sheets_data = [
        ('Profile', df_profile),
        ('Statistics', df_planes),
        ('Game Modes', df_modes)
    ]
    
    created_sheets = {}
    
    for sheet_name, df in sheets_data:
        ws = wb.create_sheet(title=sheet_name)
        created_sheets[sheet_name] = ws
        
        # Headers
        headers = list(df.columns)
        ws.append(headers)
        
        # Format Headers
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        # Data
        for row in df.itertuples(index=False):
            ws.append(list(row))
            
        # Format Data Rows
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.font = data_font
                cell.border = thin_border
                if isinstance(cell.value, float):
                    cell.number_format = '0.0'
                elif isinstance(cell.value, int):
                    cell.number_format = '#,##0'

        # --- ΠΡΟΣΘΗΚΗ ΓΡΑΜΜΗΣ TOTALS ΣΤΗΝ ΚΑΡΤΕΛΑ STATISTICS ---
        if sheet_name == 'Statistics':
            last_data_row = ws.max_row
            totals_row_idx = last_data_row + 1
            
            # Συναρτήσεις Excel για τη γραμμή Totals
            totals_formulas = [
                'Totals',
                f'=SUM(B2:B{last_data_row})',
                f'=SUM(C2:C{last_data_row})',
                f'=AVERAGE(D2:D{last_data_row})',
                f'=SUM(E2:E{last_data_row})',
                f'=AVERAGE(F2:F{last_data_row})',
                f'=SUM(G2:G{last_data_row})',
                f'=AVERAGE(H2:H{last_data_row})'
            ]
            
            ws.append(totals_formulas)
            
            # Μορφοποίηση Γραμμής Totals
            for col_num in range(1, len(totals_formulas) + 1):
                cell = ws.cell(row=totals_row_idx, column=col_num)
                cell.font = summary_font
                cell.fill = summary_fill
                cell.border = totals_border
                
                # Format αριθμών στις συναρτήσεις
                if col_num in [4, 6, 8]:  # Win Rate, Deaths/Match, Avg Score
                    cell.number_format = '0.0'
                elif col_num in [2, 3, 5, 7]:  # Matches, Wins, Deaths, MVPs
                    cell.number_format = '#,##0'

        # Auto adjust column width
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # --- ΠΡΟΣΘΗΚΗ ΓΡΑΦΗΜΑΤΟΣ GAME MODES ΣΤΗΝ ΚΑΡΤΕΛΑ STATISTICS ---
    ws_stats = created_sheets['Statistics']
    ws_modes = created_sheets['Game Modes']
    
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Αποδοτικότητα (Win Rate %) ανά Game Mode"
    chart.y_axis.title = "Win Rate (%)"
    chart.x_axis.title = "Game Mode"
    chart.legend = None
    chart.width = 16
    chart.height = 11
    
    # Αναφορά στα δεδομένα της καρτέλας Game Modes (Στήλη 4: Win Rate %, Στήλη 1: Game Mode)
    data_ref = Reference(ws_modes, min_col=4, min_row=1, max_row=ws_modes.max_row)
    cats_ref = Reference(ws_modes, min_col=1, min_row=2, max_row=ws_modes.max_row)
    
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    
    # Τοποθέτηση του γραφήματος στα δεξιά του πίνακα Statistics (στη θέση J2)
    ws_stats.add_chart(chart, "J2")

    wb.save(output)
    return output.getvalue()

# Execution Event
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
    
    try:
        accounts = raw.get("accounts", {})
        first_account_id = list(accounts.keys())[0]
        account_obj = accounts[first_account_id]
        player_data = account_obj.get("playerData", {})
    except Exception:
        player_data = raw

    # Profile Attributes
    player_name = player_data.get("playerName", "Unknown")
    friend_code = player_data.get("friendCode", "-")
    trophies = player_data.get("cachedTotalVehicleRankingPoints", 0)
    peak_trophies = player_data.get("highestTotalVehicleRankingPoints", 0)
    player_xp = player_data.get("playerXp", 0)
    hangar_points = player_data.get("cachedHangerPointsSum", 0)
    equipped_plane_raw = player_data.get("equippedVehicleConfig", "-")
    equipped_plane = equipped_plane_raw.replace("aircraft:", "").replace("-v2", "").replace("-v3", "").replace("-v4", "").replace("-", " ").title()
    pilot_banner = player_data.get("pilotBanner", "-").replace("pilot-banner:", "").replace("-", " ").title()
    
    # Career Counters
    career_counters = player_data.get("careerCounters", {})
    counters_by_plane = career_counters.get("countersByVehicleConfig", {})
    overall_stats = career_counters.get("overall", {})
    
    total_pvp_matches = overall_stats.get("pvpMatchesPlayed", 0)
    total_pvp_wins = overall_stats.get("pvpMatchesWon", 0)
    total_mvps = overall_stats.get("pvpMvpCount", 0)
    total_deaths = overall_stats.get("pvpDeaths", 0)
    total_score = overall_stats.get("pvpScore", 0)
    
    # ---------------------------------------------------------
    # A. Επεξεργασία Στατιστικών Αεροσκαφών (Lifetime Stats)
    # ---------------------------------------------------------
    planes_list = []

    for plane_id, stats in counters_by_plane.items():
        clean_name = plane_id.replace("aircraft:", "").replace("-v2", "").replace("-v3", "").replace("-v4", "").replace("-", " ").title()
        
        matches = stats.get("pvpMatchesPlayed", 0)
        wins = stats.get("pvpMatchesWon", 0)
        deaths = stats.get("pvpDeaths", 0)
        score = stats.get("pvpScore", 0)
        mvps = stats.get("pvpMvpCount", 0)
        
        win_rate = round((wins / matches) * 100, 1) if matches > 0 else 0.0
        avg_score = round(score / matches, 1) if matches > 0 else 0.0
        deaths_per_match = round(deaths / matches, 2) if matches > 0 else 0.0

        planes_list.append({
            "Aircraft": clean_name,
            "Matches": matches,
            "Wins": wins,
            "Win Rate (%)": win_rate,
            "Deaths": deaths,
            "Deaths/Match": deaths_per_match,
            "MVPs": mvps,
            "Avg Score": avg_score
        })

    df_planes = pd.DataFrame(planes_list)
    df_planes_sorted = df_planes.sort_values(by="Matches", ascending=False) if not df_planes.empty else df_planes

    # Overall Ratios
    overall_win_rate = round((total_pvp_wins / total_pvp_matches) * 100, 1) if total_pvp_matches > 0 else 0.0
    overall_d_m = round(total_deaths / total_pvp_matches, 2) if total_pvp_matches > 0 else 0.0

    # ---------------------------------------------------------
    # B. Επεξεργασία Game Modes
    # ---------------------------------------------------------
    game_modes_data = career_counters.get("countersByGameModeCategory", {})
    modes_list = []
    
    for mode_key, mode_stats in game_modes_data.items():
        clean_mode = mode_key.replace("GameModeCategory", "").replace("GameMode", "")
        mode_names_map = {
            "ControlPoints": "Air Superiority",
            "Deathmatch": "Deathmatch",
            "PriorityTarget": "Priority Target",
            "CaptureTheFlag": "Capture The Flag",
            "Elimination": "Elimination"
        }
        display_mode = mode_names_map.get(clean_mode, clean_mode)
        
        m_matches = mode_stats.get("pvpMatchesPlayed", 0)
        m_wins = mode_stats.get("pvpMatchesWon", 0)
        m_mvps = mode_stats.get("pvpMvpCount", 0)
        m_deaths = mode_stats.get("pvpDeaths", 0)
        m_score = mode_stats.get("pvpScore", 0)
        
        m_wr = round((m_wins / m_matches) * 100, 1) if m_matches > 0 else 0.0
        m_avg_score = round(m_score / m_matches, 1) if m_matches > 0 else 0.0
        
        modes_list.append({
            "Game Mode": display_mode,
            "Matches": m_matches,
            "Wins": m_wins,
            "Win Rate (%)": m_wr,
            "MVPs": m_mvps,
            "Deaths": m_deaths,
            "Avg Score": m_avg_score
        })
        
    df_modes = pd.DataFrame(modes_list).sort_values(by="Matches", ascending=False) if modes_list else pd.DataFrame()

    # ---------------------------------------------------------
    # C. Δημιουργία Dataframe για Excel Export
    # ---------------------------------------------------------
    profile_export = {
        "Player": player_name,
        "Friend Code": friend_code,
        "Trophies": trophies,
        "Max Trophies": peak_trophies,
        "Player XP": player_xp,
        "Hangar Points": hangar_points,
        "Total Win Rate (%)": overall_win_rate,
        "Total Matches": total_pvp_matches,
        "Total Wins": total_pvp_wins,
        "Deaths": total_deaths,
        "Deaths/Match": overall_d_m,
        "Total MVPs": total_mvps
    }
    df_profile_export = pd.DataFrame([profile_export])

    # --- HEADER & EXPORT BUTTON ---
    col_title, col_export = st.columns([3, 1])
    with col_title:
        st.subheader(f"👤 Παίκτης: **{player_name}** | Tag: `{friend_code}`")
        st.caption(f"Εξοπλισμένο Aircraft: **{equipped_plane}** | Banner: `{pilot_banner}`")
    
    with col_export:
        excel_bytes = generate_styled_excel(df_profile_export, df_planes_sorted, df_modes)
        st.download_button(
            label="📥 Εξαγωγή σε Excel (.xlsx)",
            data=excel_bytes,
            file_name=f"Metalstorm_Stats_{player_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
    
    st.divider()

    # --- TOP METRIC CARDS ---
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Τρόπαια 🏆", f"{trophies:,}")
    c2.metric("Peak Τρόπαια 🔝", f"{peak_trophies:,}")
    c3.metric("Win Rate 🎯", f"{overall_win_rate}%")
    c4.metric("Player XP ⭐", f"{player_xp:,}")
    c5.metric("Hangar Points 🏭", f"{hangar_points:,}")
    c6.metric("MVPs 🎖️", f"{total_mvps:,}")
    
    st.divider()

    # --- MAIN TABS ---
    tab_planes, tab_modes, tab_raw = st.tabs(["✈️ Αεροσκάφη", "🎮 Ανάλυση Game Modes", "📋 Πλήρεις Πίνακες"])

    # ---------------------------------------------------------
    # TAB 1: ΑΕΡΟΣΚΑΦΗ & GRAPHICS
    # ---------------------------------------------------------
    with tab_planes:
        if not df_planes.empty:
            df_filtered = df_planes[df_planes["Matches"] >= 5].sort_values(by="Win Rate (%)", ascending=False)
            
            sub1, sub2 = st.tabs(["📊 Γράφημα Win Rate", "📈 Avg Score ανά Αγώνα"])
            
            with sub1:
                st.subheader("Win Rate (%) ανά Aircraft (min. 5 Matches)")
                fig1 = px.bar(
                    df_filtered,
                    x="Aircraft",
                    y="Win Rate (%)",
                    text="Win Rate (%)",
                    color="Win Rate (%)",
                    color_continuous_scale="Blues",
                    hover_data=["Matches", "Wins", "Deaths", "MVPs"]
                )
                fig1.update_traces(texttemplate='%{text}%', textposition='outside')
                fig1.update_layout(template="plotly_dark", xaxis_tickangle=-45, height=450)
                st.plotly_chart(fig1, use_container_width=True)
                
            with sub2:
                st.subheader("Avg Score ανά Αγώνα")
                df_score_sorted = df_planes[df_planes["Matches"] >= 5].sort_values(by="Avg Score", ascending=False)
                fig2 = px.bar(
                    df_score_sorted,
                    x="Aircraft",
                    y="Avg Score",
                    text="Avg Score",
                    color="Avg Score",
                    color_continuous_scale="Viridis",
                    hover_data=["Matches", "MVPs"]
                )
                fig2.update_traces(texttemplate='%{text}', textposition='outside')
                fig2.update_layout(template="plotly_dark", xaxis_tickangle=-45, height=450)
                st.plotly_chart(fig2, use_container_width=True)

            st.subheader("📋 Αναλυτικός Πίνακας Στατιστικών Αεροσκαφών")
            st.dataframe(df_planes_sorted, use_container_width=True)

    # ---------------------------------------------------------
    # TAB 2: GAME MODES ANALYSIS
    # ---------------------------------------------------------
    with tab_modes:
        if not df_modes.empty:
            col_gm_chart, col_gm_data = st.columns([3, 2])
            
            with col_gm_chart:
                st.subheader("Αποδοτικότητα (Win Rate %) ανά Game Mode")
                fig_gm = px.bar(
                    df_modes,
                    x="Game Mode",
                    y="Win Rate (%)",
                    text="Win Rate (%)",
                    color="Win Rate (%)",
                    color_continuous_scale="Tealgrn",
                    hover_data=["Matches", "Wins", "MVPs"]
                )
                fig_gm.update_traces(texttemplate='%{text}%', textposition='outside')
                fig_gm.update_layout(template="plotly_dark", height=400)
                st.plotly_chart(fig_gm, use_container_width=True)
                
            with col_gm_data:
                st.subheader("Κατανομή Αγώνων")
                fig_pie = px.pie(
                    df_modes,
                    names="Game Mode",
                    values="Matches",
                    hole=0.4,
                    color_discrete_sequence=px.colors.qualitative.Pastel
                )
                fig_pie.update_layout(template="plotly_dark", height=400)
                st.plotly_chart(fig_pie, use_container_width=True)

            st.subheader("📊 Πίνακας Game Modes")
            st.dataframe(df_modes, use_container_width=True)

    # ---------------------------------------------------------
    # TAB 3: RAW TABLES & EXPORT PREVIEW
    # ---------------------------------------------------------
    with tab_raw:
        st.subheader("Προφίλ Παίκτη")
        st.dataframe(df_profile_export, use_container_width=True)
