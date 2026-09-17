import streamlit as st
import requests
import pandas as pd
import plotly.express as px

# ---------------------------------------------------------
# 1. Ρυθμίσεις Σελίδας
# ---------------------------------------------------------
st.set_page_config(
    page_title="Metalstorm Scout App",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS για πιο σκοτεινό / gaming αισθητική
st.markdown("""
    <style>
    .stApp {
        background-color: #0f172a;
        color: #f8fafc;
    }
    .metric-card {
        background-color: #1e293b;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #334155;
        text-align: center;
    }
    </style>
""", unsafe_allow_html=True)

st.title("✈️ Metalstorm Scout — Cross-Platform")
st.caption("Παρακολουθήστε τα στατιστικά σας από οποιαδήποτε συσκευή (Mobile & Desktop)")

# Sidebar για εισαγωγή Token & Οδηγίες
with st.sidebar:
    st.header("🔑 Σύνδεση Λογαριασμού")
    token_input = st.text_input("Εισάγετε το Startoken:", type="password", help="Το token που βρήκατε από το Network tab (Header Authorization)")
    
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
# 2. Συνάρτηση Ανάκτησης Δεδομένων από το Metalstorm API
# ---------------------------------------------------------
def fetch_user_stats(token):
    url = "https://playmetalstorm.com/stats/users/my"
    
    # Αφαίρεση τυχόν διπλού 'startoken' αν το επικολλούσε ο χρήστης
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
            return None, "Μη έγκυρο ή ληγμένο Token. Παρακαλώ ανανεώστε το Token σας."
        else:
            return None, f"Σφάλμα διακομιστή ({response.status_code})."
    except Exception as e:
        return None, f"Σφάλμα σύνδεσης: {str(e)}"

# ---------------------------------------------------------
# 3. Εμφάνιση Δεδομένων & Διαγραμμάτων
# ---------------------------------------------------------
if fetch_btn:
    if not token_input:
        st.warning("Παρακαλώ εισάγετε το Startoken στη δεξιά/αριστερή στήλη.")
    else:
        with st.spinner("Ανάκτηση δεδομένων από το Metalstorm..."):
            data, error = fetch_user_stats(token_input)
            
            if error:
                st.error(error)
            else:
                st.success("Τα στατιστικά φορτώθηκαν επιτυχώς!")
                
                # Αποθήκευση στο Session State για διατήρηση των δεδομένων κατά την αλληλεπίδραση
                st.session_state['user_data'] = data

# Αν υπάρχουν φορτωμένα δεδομένα στο session, τα εμφανίζουμε
if 'user_data' in st.session_state:
    data = st.session_state['user_data']
    
    # -----------------------------------------------------
    # A. Βασικές Κάρτες Στατιστικών (Top Metrics)
    # -----------------------------------------------------
    # Προσαρμόζουμε τα κλειδιά ανάλογα με τη δομή του JSON που επιστρέφει το Metalstorm
    username = data.get("username", data.get("name", "Player"))
    trophies = data.get("trophies", data.get("currentTrophies", 0))
    peak_trophies = data.get("peakTrophies", data.get("maxTrophies", trophies))
    win_rate = data.get("winRate", 0)
    total_matches = data.get("totalMatches", data.get("matchesPlayed", 0))
    
    st.subheader(f"👤 Προφίλ Παίκτη: **{username}**")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Τρόπαια 🏆", f"{trophies:,}")
    with col2:
        st.metric("Peak Τρόπαια 🔝", f"{peak_trophies:,}")
    with col3:
        st.metric("Win Rate 🎯", f"{win_rate}%" if isinstance(win_rate, (int, float)) else win_rate)
    with col4:
        st.metric("Σύνολο Αγώνων ⚔️", f"{total_matches:,}")
        
    st.divider()

    # -----------------------------------------------------
    # B. Αναλυτική Απόδοση Αεροσκαφών (Aircraft Stats)
    # -----------------------------------------------------
    aircraft_list = data.get("aircraft", data.get("planes", []))
    
    if aircraft_list:
        st.subheader("✈️ Στατιστικά ανά Αεροσκάφος")
        
        df_planes = pd.DataFrame(aircraft_list)
        
        # Καθαρισμός/Ταξινόμηση αν υπάρχουν τα αντίστοιχα πεδία
        if "winRate" in df_planes.columns:
            df_planes = df_planes.sort_values(by="winRate", ascending=False)
        
        # Γράφημα Win Rate ανά Αεροσκάφος με το Plotly
        if "name" in df_planes.columns and "winRate" in df_planes.columns:
            fig = px.bar(
                df_planes,
                x="name",
                y="winRate",
                text="winRate",
                title="Ποσοστό Νικών (Win Rate %) ανά Μαχητικό",
                labels={"name": "Αεροσκάφος", "winRate": "Win Rate (%)"},
                color="winRate",
                color_continuous_scale="Blues"
            )
            fig.update_traces(texttemplate='%{text}%', textposition='outside')
            fig.update_layout(template="plotly_dark", xaxis_tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)
        
        # Πίνακας Δεδομένων
        st.dataframe(df_planes, use_container_width=True)
    else:
        # Αν η λίστα αεροσκαφών είναι σε διαφορετικό κλειδί, εμφανίζουμε το raw JSON για έλεγχο
        st.info("💡 Αναλυτική προβολή δεδομένων JSON:")
        st.json(data)