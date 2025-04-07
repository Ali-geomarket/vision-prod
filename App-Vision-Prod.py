import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

FICHIER_SUIVI = "Suivi_demandes_AUTOMATISATION.xlsx"
USERS = {"sg": "dri", "ps": "dri"}

st.set_page_config(page_title="Vision Prod", layout="centered")

# Authentification simple
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
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

def charger_df():
    return pd.read_excel(FICHIER_SUIVI)

def enregistrer_df(df):
    with pd.ExcelWriter(FICHIER_SUIVI, engine="openpyxl", mode='w') as writer:
        df.to_excel(writer, index=False)

# Pour sécuriser les valeurs numériques
def safe_number(val):
    try:
        return float(val)
    except:
        return None

st.title("Vision Prod – Création de commande")

if "mode_modif" not in st.session_state:
    st.session_state["mode_modif"] = False
if "modif_index" not in st.session_state:
    st.session_state["modif_index"] = None

df = charger_df()

# Bouton discret "Modifier une commande"
with st.expander("Modifier une commande"):
    commande_rech = st.text_input("Entrer le nom exact de la commande à modifier")
    if st.button("Chercher la commande"):
        match = df[df["COMMANDE"] == commande_rech]
        if not match.empty:
            st.session_state["mode_modif"] = True
            st.session_state["modif_index"] = match.index[0]
        else:
            st.error("Commande non trouvée.")

# Pré-remplir le formulaire si modification
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
    st.subheader("Formulaire de saisie")

    cout_ext_val = safe_number(modif_data.get("cout_ext"))
    cout_global_val = safe_number(modif_data.get("cout_global"))
    tirage_val = safe_number(modif_data.get("tirage"))

    cout_ext = st.number_input(
        "Coût de l'extension (€)",
        min_value=0.0,
        value=cout_ext_val if cout_ext_val is not None else 0.0,
        step=100.0,
        format="%.2f"
    )
    if cout_ext and (cout_ext < 100 or cout_ext > 100000):
        st.warning("Coût de l'extension hors limites (100€ - 100 000€)")

    cout_global = st.number_input(
        "Coût global du projet (€)",
        min_value=0.0,
        value=cout_global_val if cout_global_val is not None else 0.0,
        step=100.0,
        format="%.2f"
    )
    if cout_global and (cout_global < 100 or cout_global > 100000):
        st.warning("Coût global hors limites (100€ - 100 000€)")

    tirage = st.number_input(
        "Tirage total (ml)",
        min_value=0.0,
        value=tirage_val if tirage_val is not None else 0.0,
        step=100.0,
        format="%.2f"
    )
    if tirage and tirage > 50000:
        st.warning("Tirage supérieur à 50 000 ml")

    reseau = st.text_input("Réseau", value=modif_data.get("reseau", ""))
    fichier_bpe = st.file_uploader("Fichier BPE à poser (KMZ/KML/SHP)", type=["kmz", "kml", "shp"])

    commande = modif_data.get("commande") or f"CMD_X_{datetime.now().strftime('%Y%m%d%H%M%S')}"

    submit = st.form_submit_button("Modifier la commande" if st.session_state["mode_modif"] else "Envoyer")

if submit:
    if not fichier_bpe:
        st.error("Le fichier BPE est obligatoire.")
        st.stop()

    nouvelle_ligne = {
        "COMMERCIAL": "",
        "PROJET": "",
        "TYPE DE DEMANDE": "",
        "COUT EXTENSION": cout_ext,
        "COUT GLOBAL PROJET": cout_global,
        "OPERATEUR": "",
        "TIRAGE TOTAL": tirage,
        "GAIN DRI": "",
        "ROI": "",
        "CLIENTS AMORTISSEMENT": "",
        "COMMANDE": commande,
        "DATE TRAITEMENT": datetime.today().strftime("%d/%m/%Y"),
        "DELAI TRAITEMENT": "",
        "ETAT GEOMARKETING": "",
        "RESP GEOMARKET": "",
        "CONCLUSION": "",
        "COMMENTAIRE": "",
        "RESEAU": reseau
    }

    if st.session_state["mode_modif"]:
        df.loc[st.session_state["modif_index"]] = nouvelle_ligne
        st.success("Commande modifiée avec succès.")
        st.session_state["mode_modif"] = False
    else:
        df = pd.concat([df, pd.DataFrame([nouvelle_ligne])], ignore_index=True)
        st.success("Commande ajoutée avec succès.")

    enregistrer_df(df)
    st.dataframe(pd.DataFrame([nouvelle_ligne]))

    st.button("Regénérer MA (à venir)")

