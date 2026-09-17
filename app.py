import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd
import plotly.express as px

# 1. Ρύθμιση σελίδας για κινητά και PC
st.set_page_config(
    page_title="Metalstorm Stats Scout",
    page_icon="✈️",
    layout="centered"
)

st.title("✈️ Metalstorm Stats Scout")
st.caption("Παρακολουθήστε τα στατιστικά σας από οποιαδήποτε συσκευή!")

# 2. Input για το Player ID ή Profile URL
player_id = st.text_input("Εισάγετε το Player ID ή το URL προφίλ:", placeholder="π.χ. 12345678")

# 3. Συνάρτηση για scraping/fetching δεδομένων
def fetch_metalstorm_stats(pid):
    url = f"https://playmetalstorm.com/stats/{pid}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code != 200:
            return None, "Δεν βρέθηκε παίκτης με αυτό το ID."
        
        soup = BeautifulSoup(response.text, "html.parser")
        
        # ΕΔΩ: Προσαρμόζουμε το scraping ανάλογα με τη δομή του HTML
        # Παράδειγμα άντλησης βασικών δεδομένων:
        stats = {
            "Username": "Player_One", # Αντικαθίσταται με soup.find(...)
            "Trophies": 1250,
            "Win Rate": "68%",
            "Total Matches": 450,
            "Aircraft": [
                {"Name": "F-22 Raptor", "Win Rate": 79, "Matches": 120},
                {"Name": "JAS 39 Gripen", "Win Rate": 76, "Matches": 95},
                {"Name": "Su-22 Fitter", "Win Rate": 70, "Matches": 60},
            ]
        }
        return stats, None
    except Exception as e:
        return None, f"Σφάλμα κατά τη σύνδεση: {str(e)}"

# 4. Εμφάνιση Αποτελεσμάτων
if st.button("Αναζήτηση Στατιστικών", type="primary"):
    if not player_id:
        st.warning("Παρακαλώ συμπληρώστε το Player ID.")
    else:
        with st.spinner("Ανάκτηση δεδομένων από το Metalstorm..."):
            data, error = fetch_metalstorm_stats(player_id)
            
            if error:
                st.error(error)
            else:
                st.success(f"Βρέθηκαν στατιστικά για τον παίκτη: **{data['Username']}**")
                
                # Metrics (Κάρτες Στατιστικών)
                col1, col2, col3 = st.columns(3)
                col1.metric("Τρόπαια 🏆", data["Trophies"])
                col2.metric("Win Rate 🎯", data["Win Rate"])
                col3.metric("Σύνολο Αγώνων ⚔️", data["Total Matches"])
                
                st.divider()
                st.subheader("📊 Απόδοση ανά Αεροσκάφος")
                
                # Δημιουργία Dataframe & Γραφήματος
                df = pd.DataFrame(data["Aircraft"])
                
                fig = px.bar(
                    df, 
                    x="Name", 
                    y="Win Rate", 
                    text="Win Rate",
                    labels={"Name": "Αεροσκάφος", "Win Rate": "Win Rate (%)"},
                    color="Win Rate",
                    color_continuous_scale="Viridis"
                )
                fig.update_traces(texttemplate='%{text}%', textposition='outside')
                
                st.plotly_chart(fig, use_container_width=True)
                
                # Πίνακας Δεδομένων
                st.dataframe(df, use_container_width=True)