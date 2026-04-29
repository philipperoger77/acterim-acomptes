import streamlit as st
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from datetime import datetime, date
import json

# Connexion Google Sheet
def get_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes
    )
    client = gspread.authorize(creds)
    return client.open_by_key("1M9lOBhJrc50ts8g4aYBCNEOfGuFmbEbEhUUNgEKCA7E")
  # Configuration
st.set_page_config(
    page_title="Acterim - Acomptes",
    page_icon="💶",
    layout="centered"
)

BUREAUX = [
    "BORDEAUX", "GRENOBLE", "LILLE", "LYON", "MARSEILLE",
    "MONTPELLIER", "NANTES", "RENNES", "ROUEN",
    "SAINT MALO", "STRASBOURG", "TOULOUSE"
]

# Détection du mode
params = st.query_params
bureau = params.get("bureau", "").upper().strip()
mode_agence = bureau in BUREAUX
mode_admin = not mode_agence
# =====================
# VUE AGENCE
# =====================
if mode_agence:
    st.title(f"💶 Demande d'acompte — {bureau}")

    # Chargement des salariés du bureau
    try:
        sheet = get_sheet()
        ws_salaries = sheet.worksheet("SALARIES")
        data = ws_salaries.get_all_records()
        df = pd.DataFrame(data)
        df_bureau = df[df["BUREAU"].str.upper() == bureau].copy()

        if df_bureau.empty:
            st.warning("Aucun salarié actif pour ce bureau.")
            st.stop()

        # Menu déroulant
        df_bureau["AFFICHAGE"] = (
            "[" + df_bureau["CODE AGENCE"].astype(str) + "] " +
            df_bureau["NOM"] + " " +
            df_bureau["PRENOM"] + " — " +
            df_bureau["DATE FIN THEORIQUE DERNIERE MISSION"].astype(str)
        )

        choix = st.selectbox("Salarié", df_bureau["AFFICHAGE"].tolist())
        ligne = df_bureau[df_bureau["AFFICHAGE"] == choix].iloc[0]

        # Saisie
        montant = st.number_input("Montant de l'acompte (€)", min_value=0.0, step=10.0)
        commentaire = st.text_input("Commentaire (optionnel)")

        # Bouton valider
        if st.button("✅ Valider la demande"):
            if montant <= 0:
                st.error("Le montant doit être supérieur à 0.")
            else:
                ws_demandes = sheet.worksheet("DEMANDES")
                nouvelle_ligne = [
                    datetime.now().strftime("%d/%m/%Y %H:%M"),
                    ligne["BUREAU"],
                    str(ligne["CODE AGENCE"]),
                    str(ligne["MATRICULE"]),
                    ligne["NOM"],
                    ligne["PRENOM"],
                    str(ligne["DATE FIN THEORIQUE DERNIERE MISSION"]),
                    str(ligne["MATRICULE DERNIERE MISSION"]),
                    montant,
                    commentaire,
                    "EN ATTENTE"
                ]
                ws_demandes.append_row(nouvelle_ligne)
                st.success(f"Demande enregistrée pour {ligne['NOM']} {ligne['PRENOM']} — {montant} €")

    except Exception as e:
        st.error(f"Erreur de connexion : {e}")
      # =====================
# VUE GESTIONNAIRE
# =====================
if mode_admin:
    st.title("🔐 Acterim — Gestion des acomptes")

    # Login
    if "authentifie" not in st.session_state:
        st.session_state.authentifie = False

    if not st.session_state.authentifie:
        pwd = st.text_input("Mot de passe", type="password")
        if st.button("Se connecter"):
            if pwd == st.secrets["admin_password"]:
                st.session_state.authentifie = True
                st.rerun()
            else:
                st.error("Mot de passe incorrect.")
        st.stop()

    # Interface gestionnaire
    st.success("Connecté ✅")
    if st.button("Se déconnecter"):
        st.session_state.authentifie = False
        st.rerun()

    try:
        sheet = get_sheet()
        ws_demandes = sheet.worksheet("DEMANDES")
        data = ws_demandes.get_all_records()
        df = pd.DataFrame(data)

        if df.empty:
            st.info("Aucune demande enregistrée.")
            st.stop()

        df_attente = df[df["STATUT"] == "EN ATTENTE"].copy()

        if df_attente.empty:
            st.info("Aucune demande en attente.")
        else:
            st.subheader(f"{len(df_attente)} demande(s) EN ATTENTE")
            for idx, row in df_attente.iterrows():
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.write(
                        f"**{row['NOM']} {row['PRENOM']}** — "
                        f"{row['BUREAU']} [{row['CODE AGENCE']}] — "
                        f"**{row['MONTANT']} €** — "
                        f"saisi le {row['DATE SAISIE']}"
                    )
                    if row["COMMENTAIRE"]:
                        st.caption(f"💬 {row['COMMENTAIRE']}")
                with col2:
                    if st.button("✅ Traité", key=f"traite_{idx}"):
                        # Trouver la ligne dans le Sheet (idx + 2 car ligne 1 = en-têtes)
                        col_statut = df.columns.tolist().index("STATUT") + 1
                        ws_demandes.update_cell(idx + 2, col_statut, "TRAITE")
                        st.rerun()

    except Exception as e:
        st.error(f"Erreur de connexion : {e}")
