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
st.image("logo_acterim.png", width=200)

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

 # Menu déroulant CLIENT
        clients = sorted(df_bureau["CLIENT"].dropna().unique().tolist())
        client_choisi = st.selectbox("Client", clients)
        df_client = df_bureau[df_bureau["CLIENT"] == client_choisi].copy()

        if df_client.empty:
            st.warning("Aucun salarié actif pour ce client.")
            st.stop()

        # Chargement des demandes EN ATTENTE pour vérif doublons
        ws_demandes = sheet.worksheet("DEMANDES")
        demandes_data = ws_demandes.get_all_records()
        df_demandes = pd.DataFrame(demandes_data)
        en_attente = []
        if not df_demandes.empty:
            en_attente = df_demandes[df_demandes["STATUT"] == "EN ATTENTE"]["MATRICULE"].astype(str).tolist()

        # Affichage salariés en masse
        st.markdown("---")
        montants = {}
        commentaires = {}
        for _, row in df_client.iterrows():
            mat = str(row["MATRICULE"])
            label = (
                f"[{row['CODE AGENCE']}] {row['NOM']} {row['PRENOM']} "
                f"— Mission {row['MATRICULE DERNIERE MISSION']} "
                f"— {row['DATE FIN THEORIQUE DERNIERE MISSION']}"
            )
            col1, col2, col3 = st.columns([3, 1, 2])
            with col1:
                st.markdown(f"**{label}**")
                if mat in en_attente:
                    st.warning("⚠️ Demande déjà en attente")
            with col2:
                montants[mat] = st.number_input(
                    "Montant (€)", min_value=0.0, step=10.0,
                    key=f"montant_{mat}"
                )
            with col3:
                commentaires[mat] = st.text_input(
                    "Commentaire", key=f"commentaire_{mat}"
                )
            st.markdown("---")

        # Bouton valider tout
        if st.button("✅ Valider toutes les demandes"):
            erreurs = []
            succes = []
            for _, row in df_client.iterrows():
                mat = str(row["MATRICULE"])
                montant = montants[mat]
                if montant <= 0:
                    continue
                if mat in en_attente:
                    erreurs.append(f"{row['NOM']} {row['PRENOM']} — demande déjà en attente")
                    continue
                nouvelle_ligne = [
                    datetime.now().strftime("%d/%m/%Y %H:%M"),
                    row["BUREAU"],
                    str(row["CODE AGENCE"]),
                    mat,
                    row["NOM"],
                    row["PRENOM"],
                    str(row["DATE FIN THEORIQUE DERNIERE MISSION"]),
                    str(row["MATRICULE DERNIERE MISSION"]),
                    montant,
                    commentaires[mat],
                    "EN ATTENTE"
                ]
                ws_demandes.append_row(nouvelle_ligne)
                succes.append(f"{row['NOM']} {row['PRENOM']} — {montant} €")

            if succes:
                st.success("Demandes enregistrées :\n" + "\n".join(succes))
            if erreurs:
                st.error("Ignorées (déjà en attente) :\n" + "\n".join(erreurs))       

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
