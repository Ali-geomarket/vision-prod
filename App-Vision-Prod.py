import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

FICHIER_SUIVI = "Suivi_demandes_AUTOMATISATION.xlsx"
USERS = {"sg": "dri", "ps": "dri"}

# -------------------------
# État initial
# -------------------------
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "page" not in st.session_state:
    st.session_state["page"] = "formulaire"
if "mode_modif" not in st.session_state:
    st.session_state["mode_modif"] = False
if "modif_index" not in st.session_state:
    st.session_state["modif_index"] = None
if "ligne_temporaire" not in st.session_state:
    st.session_state["ligne_temporaire"] = None

# -------------------------
# Authentification
# -------------------------
if not st.session_state["authenticated"]:
    st.title("Connexion")
    with st.form("login_form"):
        username = st.text_input("Nom d'utilisateur")
        password = st.text_input("Mot de passe", type="password")
        if st.form_submit_button("Se connecter"):
            if USERS.get(username) == password:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Identifiants incorrects")
    st.stop()

# -------------------------
# Fonctions
# -------------------------
def charger_df():
    return pd.read_excel(FICHIER_SUIVI)

def enregistrer_df(df):
    with pd.ExcelWriter(FICHIER_SUIVI, engine="openpyxl", mode='w') as writer:
        df.to_excel(writer, index=False)

def safe_number(val):
    try:
        return float(val)
    except:
        return None

colonnes_excel = [
    "DATE RECEPTION", "RESEAU", "RESPONSABLE PROD", "COMMERCIAL", "PROJET", "TYPE DE DEMANDE",
    "COUT EXTENSION", "COUT GLOBAL PROJET", "OPERATEUR", "TIRAGE TOTAL",
    "GAIN DRI", "ROI", "NB CLIENTS AMORTISSEMENT", "COMMANDE", "DATE TRAITEMENT",
    "DELAI TRAITEMENT", "ETAT GEOMARKETING", "RESP GEOMARKET", "CONCLUSION", "COMMENTAIRE"
]

df = charger_df()

# -------------------------
# Barre de navigation
# -------------------------
st.set_page_config(page_title="Vision Prod", layout="centered")
st.title("Vision Prod – Création de commande")

col1, col2, col3 = st.columns(3)
with col1:
    if st.button("Créer une commande"):
        st.session_state["page"] = "formulaire"
        st.session_state["mode_modif"] = False
        st.session_state["modif_index"] = None
        st.session_state["ligne_temporaire"] = None
        st.rerun()
with col2:
    if st.button("Visualiser Excel"):
        st.session_state["page"] = "visualiser"
        st.rerun()
with col3:
    if st.button("Modifier une commande"):
        st.session_state["page"] = "modifier"
        st.rerun()

# -------------------------
# PAGE : FORMULAIRE
# -------------------------
if st.session_state["page"] == "formulaire":
    st.header("Création de commande")

    modif_data = {}
    if st.session_state["mode_modif"]:
        ligne = df.loc[st.session_state["modif_index"]]
        modif_data = {
            "cout_ext": ligne.get("COUT EXTENSION", None),
            "cout_global": ligne.get("COUT GLOBAL PROJET", None),
            "tirage": ligne.get("TIRAGE TOTAL", None),
            "reseau": ligne.get("RESEAU", ""),
            "commande": ligne.get("COMMANDE", "")
        }

    with st.form("formulaire_commande"):
        cout_ext_val = safe_number(modif_data.get("cout_ext"))
        cout_global_val = safe_number(modif_data.get("cout_global"))
        tirage_val = safe_number(modif_data.get("tirage"))

        cout_ext = st.number_input("Coût de l'extension (€)", min_value=0.0, value=cout_ext_val if cout_ext_val is not None else 0.0, step=100.0, format="%.2f")
        if cout_ext and (cout_ext < 100 or cout_ext > 100000):
            st.warning("Coût de l'extension hors limites (100€ - 100 000€)")

        cout_global = st.number_input("Coût global du projet (€)", min_value=0.0, value=cout_global_val if cout_global_val is not None else 0.0, step=100.0, format="%.2f")
        if cout_global and (cout_global < 100 or cout_global > 100000):
            st.warning("Coût global hors limites (100€ - 100 000€)")

        tirage = st.number_input("Tirage total (ml)", min_value=0.0, value=tirage_val if tirage_val is not None else 0.0, step=100.0, format="%.2f")
        if tirage and tirage > 50000:
            st.warning("Tirage supérieur à 50 000 ml")

        reseau = st.text_input("Réseau", value=modif_data.get("reseau", ""))
        fichier_bpe = st.file_uploader("Fichier BPE à poser (KMZ/KML/SHP)", type=["kmz", "kml", "shp"])
        commande = modif_data.get("commande") or f"CMD_X_{datetime.now().strftime('%Y%m%d%H%M%S')}"

        submit = st.form_submit_button("Modifier la commande" if st.session_state["mode_modif"] else "Envoyer")

    if st.session_state["mode_modif"]:
        if st.button("Annuler la modification"):
            st.session_state["mode_modif"] = False
            st.session_state["modif_index"] = None
            st.rerun()

    if submit:
        if not fichier_bpe:
            st.error("Le fichier BPE est obligatoire.")
            st.stop()

        nouvelle_ligne = {
            "DATE RECEPTION": datetime.today().strftime("%d/%m/%Y"),
            "RESEAU": reseau,
            "RESPONSABLE PROD": "",
            "COMMERCIAL": "",
            "PROJET": "",
            "TYPE DE DEMANDE": "",
            "COUT EXTENSION": cout_ext,
            "COUT GLOBAL PROJET": cout_global,
            "OPERATEUR": "",
            "TIRAGE TOTAL": tirage,
            "GAIN DRI": "",
            "ROI": "",
            "NB CLIENTS AMORTISSEMENT": "",
            "COMMANDE": commande,
            "DATE TRAITEMENT": datetime.today().strftime("%d/%m/%Y"),
            "DELAI TRAITEMENT": "",
            "ETAT GEOMARKETING": "",
            "RESP GEOMARKET": "",
            "CONCLUSION": "",
            "COMMENTAIRE": ""
        }

        if st.session_state["mode_modif"]:
            df.loc[st.session_state["modif_index"]] = nouvelle_ligne
            enregistrer_df(df)
            st.success("Commande modifiée avec succès.")
            st.dataframe(pd.DataFrame([nouvelle_ligne]))
            if st.button("Retour au formulaire vierge"):
                st.session_state["mode_modif"] = False
                st.session_state["modif_index"] = None
                st.rerun()
        else:
            st.session_state["ligne_temporaire"] = nouvelle_ligne
            st.success("Commande ajoutée avec succès.")
            st.rerun()

    if st.session_state["ligne_temporaire"] is not None:
        st.subheader("Ligne à valider")
        df_temp = pd.DataFrame([st.session_state["ligne_temporaire"]])[colonnes_excel]
        edited_row = st.data_editor(df_temp, num_rows="fixed", use_container_width=True)

        if st.button("Valider l'enregistrement"):
            df = pd.concat([df, edited_row], ignore_index=True)
            enregistrer_df(df)
            st.session_state["ligne_temporaire"] = None
            st.success("Ligne enregistrée avec les modifications.")
            if st.button("Nouvelle commande"):
                st.rerun()

# -------------------------
# PAGE : VISUALISER
# -------------------------
elif st.session_state["page"] == "visualiser":
    st.header("Visualisation du fichier Excel")
    try:
        df_viz = pd.read_excel(FICHIER_SUIVI)
        st.dataframe(df_viz, use_container_width=True)
        with open(FICHIER_SUIVI, "rb") as file:
            st.download_button("Télécharger le fichier Excel", data=file, file_name="Suivi_demandes_AUTOMATISATION.xlsx")
    except Exception as e:
        st.error(f"Erreur de lecture : {e}")
    if st.button("⬅ Retour au formulaire"):
        st.session_state["page"] = "formulaire"
        st.rerun()

# -------------------------
# PAGE : MODIFIER
# -------------------------
elif st.session_state["page"] == "modifier":
    st.header("Modifier une commande existante")
    commande_rech = st.text_input("Nom de la commande à modifier")
    if st.button("Chercher la commande"):
        match = df[df["COMMANDE"] == commande_rech]
        if not match.empty:
            st.session_state["mode_modif"] = True
            st.session_state["modif_index"] = match.index[0]
            st.session_state["page"] = "formulaire"
            st.rerun()
        else:
            st.error("Commande non trouvée.")
    if st.button("⬅ Retour au formulaire"):
        st.session_state["page"] = "formulaire"
        st.rerun()
