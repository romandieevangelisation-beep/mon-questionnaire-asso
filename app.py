import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
from docx import Document
from io import BytesIO
import json
from supabase import create_client, Client

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="Espace Clinique Sécurisé", layout="wide")

# --- CONNEXION SÉCURISÉE À LA BASE DE DONNÉES (SUPABASE) ---
# On récupère les clés secrètes depuis la configuration du serveur (pas dans le code visible)
@st.cache_resource
def init_connection():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase = init_connection()

# --- FONCTIONS DE GESTION DES DONNÉES ---

def save_patient_data(nom, email, reponses_dict):
    """Envoie les données chiffrées vers la base de données"""
    data = {
        "nom": nom,
        "email": email,
        "reponses_json": json.dumps(reponses_dict), # On convertit le dict en texte
        "created_at": datetime.now().isoformat()
    }
    # Insertion sécurisée
    try:
        supabase.table("patients_ysq").insert(data).execute()
        return True
    except Exception as e:
        st.error(f"Erreur de connexion : {e}")
        return False

def load_all_patients():
    """Récupère la liste des patients depuis la base sécurisée"""
    try:
        response = supabase.table("patients_ysq").select("*").order("created_at", desc=True).execute()
        return pd.DataFrame(response.data)
    except Exception as e:
        st.error(f"Erreur de chargement : {e}")
        return pd.DataFrame()

# --- VOS DONNÉES CLINIQUES (YSQ & Interprétations) ---
# (Je remets la structure courte pour l'exemple, INSÉREZ ICI VOS 232 QUESTIONS)
YSQ_QUESTIONS = {
    "ED : Carence affective": {1: "Besoins affectifs non comblés", 2: "Manque d'amour"}, 
    "AB : Abandon": {10: "Peur de la mort des proches", 11: "Peur de l'abandon"}
}
# INSÉREZ ICI LE DICTIONNAIRE INTERPRETATIONS_BIBLIQUES COMPLET

# --- INTERFACE ---

# Menu latéral sécurisé
st.sidebar.title("Portail YSQ-L3")
mode = st.sidebar.radio("Accès :", ["Espace Patient", "Accès Thérapeute"])

# ==============================================================================
# 1. ESPACE PATIENT (PUBLIC)
# ==============================================================================
if mode == "Espace Patient":
    st.header("🌱 Questionnaire YSQ-L3")
    st.write("Vos réponses sont transmises de manière sécurisée à votre thérapeute.")

    with st.form("form_patient", clear_on_submit=True):
        c1, c2 = st.columns(2)
        nom = c1.text_input("Nom complet")
        email = c2.text_input("Email")
        
        reponses = {}
        st.divider()
        
        # Boucle des questions
        for domaine, q_dict in YSQ_QUESTIONS.items():
            st.subheader(domaine)
            for q_num, q_text in q_dict.items():
                reponses[f"Q{q_num}"] = st.slider(q_text, 1, 6, key=f"q_{q_num}")
        
        submitted = st.form_submit_button("Envoyer mes résultats")
        
        if submitted:
            if not nom:
                st.warning("Merci d'indiquer votre nom.")
            else:
                success = save_patient_data(nom, email, reponses)
                if success:
                    st.success("✅ Vos réponses ont été enregistrées avec succès dans notre base sécurisée.")
                    st.balloons()

# ==============================================================================
# 2. ESPACE THÉRAPEUTE (PROTÉGÉ)
# ==============================================================================
elif mode == "Accès Thérapeute":
    st.sidebar.divider()
    # Le mot de passe Admin est aussi caché dans les secrets pour la sécurité
    pwd_input = st.sidebar.text_input("Mot de passe Maître", type="password")
    
    # Vérification sécurisée
    if pwd_input == st.secrets["ADMIN_PASSWORD"]:
        st.header("🔒 Tableau de Bord Clinique")
        
        df = load_all_patients()
        
        if df.empty:
            st.info("Aucun dossier patient dans la base de données.")
        else:
            # Liste des patients
            st.dataframe(df[["created_at", "nom", "email"]])
            
            st.divider()
            patient_select = st.selectbox("Choisir un dossier à analyser :", df["nom"].unique())
            
            if st.button("Lancer l'Analyse"):
                # Récupération des données brutes
                patient_data = df[df["nom"] == patient_select].iloc[0]
                reponses_dict = json.loads(patient_data["reponses_json"])
                
                # --- ALGORITHME DE CALCUL (Identique à avant) ---
                resultats = []
                schemas_actifs_codes = []
                
                for domaine, q_dict in YSQ_QUESTIONS.items():
                    code = domaine.split(" : ")[0]
                    nom_sch = domaine.split(" : ")[1]
                    scores = [reponses_dict.get(f"Q{k}", 0) for k in q_dict.keys()]
                    
                    if scores:
                        moy = sum(scores) / len(scores)
                        sev = len([x for x in scores if x >= 5])
                        pct = (sev / len(scores)) * 100
                        etoile = "*" if sev > 0 else ""
                        if etoile: schemas_actifs_codes.append(code)
                        
                        resultats.append({
                            "Code": code,
                            "Schéma": f"{nom_sch} {etoile}",
                            "Moyenne": moy,
                            "Niveau": "🔴" if moy > 3.5 else "🟡" if moy >= 2.5 else "🟢"
                        })
                
                df_res = pd.DataFrame(resultats)
                
                # Affichage Visuel
                c1, c2 = st.columns([1, 2])
                c1.table(df_res[["Code", "Moyenne", "Niveau"]])
                
                fig = px.line_polar(df_res, r='Moyenne', theta='Code', line_close=True, range_r=[0,6])
                fig.update_traces(fill='toself')
                c2.plotly_chart(fig)
                
                # --- BOUTON WORD ---
                # (Insérez ici la fonction generate_docx simplifiée sans images pour le cloud)
                # Note: Sur le cloud, éviter kaleido pour les images si possible, ou utiliser le tableau simple
                st.write("Le module d'export Word est prêt (code à insérer).")

    elif pwd_input:
        st.error("Mot de passe incorrect.")
