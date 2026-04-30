# -*- coding: utf-8 -*-
import streamlit as st
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from datetime import datetime, date, timedelta
import io
import calendar

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

def parse_date(date_str):
    """Parse une date DD/MM/YYYY ou YYYY-MM-DD"""
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(date_str).strip(), fmt)
        except:
            continue
    return None

def calculer_lundi(date_fin_str, fin_mois):
    """Calcule le lundi de la semaine selon la règle métier"""
    date_fin = parse_date(date_fin_str)
    if not date_fin:
        return ""
    date_ref = fin_mois if date_fin > fin_mois else date_fin
    lundi = date_ref - timedelta(days=date_ref.weekday())
    return lundi.strftime("%d/%m/%Y")

def calculer_lundi_avec_fallback(code_mission, date_fin_str, date_debut_str, fin_mois, df_salaries):
    """
    Calcule le lundi en vérifiant qu'il est dans la plage de la mission.
    Si non, remonte à la mission N-1 du même salarié/agence.
    Retourne (code_mission_final, date_lundi_final)
    """
    date_fin = parse_date(date_fin_str)
    date_debut = parse_date(date_debut_str)

    # Calcul du lundi candidat
    lundi_str = calculer_lundi(date_fin_str, fin_mois)
    lundi = parse_date(lundi_str)

    # Vérification : le lundi est-il dans la plage de la mission ?
    if lundi and date_debut and date_fin:
        if date_debut <= lundi <= date_fin:
            return code_mission, lundi_str

    # Fallback : chercher la mission N-1 dans SALARIES
    # Trouver le salarié correspondant à ce code mission (RANG MISSION = 1)
    mission_rang1 = df_salaries[df_salaries["MATRICULE MISSION"].astype(str) == str(code_mission)]
    if mission_rang1.empty:
        return code_mission, lundi_str  # pas de fallback possible

    matricule = mission_rang1.iloc[0]["MATRICULE"]
    code_agence = mission_rang1.iloc[0]["CODE AGENCE"]

    # Chercher la mission N-1 (RANG MISSION = 2) pour ce salarié/agence
    mission_n1 = df_salaries[
        (df_salaries["MATRICULE"].astype(str) == str(matricule)) &
        (df_salaries["CODE AGENCE"].astype(str) == str(code_agence)) &
        (df_salaries["RANG MISSION"].astype(str) == "2")
    ]

    if mission_n1.empty:
        return code_mission, lundi_str  # pas de N-1 disponible

    row_n1 = mission_n1.iloc[0]
    date_fin_n1_str = str(row_n1["DATE FIN MISSION"])
    date_debut_n1_str = str(row_n1["DATE DEBUT MISSION"])
    code_mission_n1 = str(row_n1["MATRICULE MISSION"])

    # Calcul du lundi sur la mission N-1
    lundi_n1_str = calculer_lundi(date_fin_n1_str, fin_mois)
    lundi_n1 = parse_date(lundi_n1_str)
    date_debut_n1 = parse_date(date_debut_n1_str)
    date_fin_n1 = parse_date(date_fin_n1_str)

    # Vérifier que le lundi N-1 est dans la plage ET dans le mois choisi
    fin_mois_dt = fin_mois
    debut_mois = datetime(fin_mois_dt.year, fin_mois_dt.month, 1)

    if lundi_n1 and date_debut_n1 and date_fin_n1:
        if date_debut_n1 <= lundi_n1 <= date_fin_n1 and debut_mois <= lundi_n1 <= fin_mois_dt:
            return code_mission_n1, lundi_n1_str

    # Si N-1 ne convient pas non plus, on reste sur la mission d'origine
    return code_mission, lundi_str

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

MOIS_FR = {
    1: "JANVIER", 2: "FEVRIER", 3: "MARS", 4: "AVRIL",
    5: "MAI", 6: "JUIN", 7: "JUILLET", 8: "AOUT",
    9: "SEPTEMBRE", 10: "OCTOBRE", 11: "NOVEMBRE", 12: "DECEMBRE"
}

# Génération liste mois : Avril 2026 + 24 mois
mois_options = []
mois_depart = 4
annee_depart = 2026
for i in range(24):
    m = (mois_depart + i - 1) % 12 + 1
    y = annee_depart + (mois_depart + i - 1) // 12
    mois_options.append((m, y, f"{MOIS_FR[m]} {y}"))

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

    try:
        sheet = get_sheet()
        ws_salaries = sheet.worksheet("SALARIES")
        data = ws_salaries.get_all_records()
        df = pd.DataFrame(data)
        df_bureau = df[df["BUREAU"].str.upper() == bureau].copy()

        if df_bureau.empty:
            st.warning("Aucun salarié actif pour ce bureau.")
            st.stop()

        # Filtrer uniquement RANG MISSION = 1 pour l'affichage agence
        df_bureau = df_bureau[df_bureau["RANG MISSION"].astype(str) == "1"].copy()

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
            df_en_attente = df_demandes[df_demandes["STATUT"] == "EN ATTENTE"].copy()
            en_attente = (
                df_en_attente["MATRICULE"].astype(str) + "_" +
                df_en_attente["CODE AGENCE"].astype(str) + "_" +
                df_en_attente["MATRICULE MISSION"].astype(str)
            ).tolist()

        # Affichage salariés en masse
        st.markdown("---")
        montants = {}
        commentaires = {}
        for _, row in df_client.iterrows():
            mat = str(row["MATRICULE"])
            miss = str(row["MATRICULE MISSION"])
            label = (
                f"[{row['CODE AGENCE']}] {row['NOM']} {row['PRENOM']} "
                f"— Mission {miss} "
                f"— {row['DATE FIN MISSION']}"
            )
            col1, col2, col3 = st.columns([3, 1, 2])
            with col1:
                st.markdown(f"**{label}**")
                cle = mat + "_" + str(row["CODE AGENCE"]) + "_" + miss
                if cle in en_attente:
                    st.warning("⚠️ Demande déjà en attente")
            with col2:
                montants[miss] = st.number_input(
                    "Montant (€)", min_value=0.0, step=10.0,
                    key=f"montant_{miss}"
                )
            with col3:
                commentaires[miss] = st.text_input(
                    "Commentaire", key=f"commentaire_{miss}"
                )
            st.markdown("---")

        # Bouton valider tout
        if st.button("✅ Valider toutes les demandes"):
            erreurs = []
            succes = []
            for _, row in df_client.iterrows():
                mat = str(row["MATRICULE"])
                miss = str(row["MATRICULE MISSION"])
                montant = montants[miss]
                if montant <= 0:
                    continue
                cle = mat + "_" + str(row["CODE AGENCE"]) + "_" + miss
                if cle in en_attente:
                    erreurs.append(f"{row['NOM']} {row['PRENOM']} — demande déjà en attente")
                    continue
                nouvelle_ligne = [
                    (datetime.now() + timedelta(hours=2)).strftime("%d/%m/%Y %H:%M"),
                    row["BUREAU"],
                    str(row["CODE AGENCE"]),
                    mat,
                    row["NOM"],
                    row["PRENOM"],
                    str(row["DATE FIN MISSION"]),
                    miss,
                    montant,
                    commentaires[miss],
                    "EN ATTENTE",
                    row["CLIENT"]
                ]
                ws_demandes.append_row(nouvelle_ligne)
                succes.append(f"{row['NOM']} {row['PRENOM']} — {montant} €")

            if succes:
                st.success("Demandes enregistrées :\n" + "\n".join(succes))
            if erreurs:
                st.error("Ignorées (déjà en attente) :\n" + "\n".join(erreurs))

        # Historique des demandes du bureau
        st.markdown("---")
        st.subheader("📋 Historique des demandes — " + bureau)
        try:
            ws_hist = sheet.worksheet("DEMANDES")
            hist_data = ws_hist.get_all_records()
            df_hist = pd.DataFrame(hist_data)
            if not df_hist.empty:
                df_hist_bureau = df_hist[df_hist["BUREAU"].str.upper() == bureau].copy()
                if df_hist_bureau.empty:
                    st.info("Aucune demande enregistrée pour ce bureau.")
                else:
                    df_hist_bureau = df_hist_bureau.sort_values("DATE SAISIE", ascending=False)
                    df_hist_bureau = df_hist_bureau[[
                        "DATE SAISIE", "CODE AGENCE", "NOM", "PRENOM",
                        "CLIENT", "MATRICULE MISSION",
                        "MONTANT", "COMMENTAIRE", "STATUT"
                    ]]
                    st.dataframe(df_hist_bureau, use_container_width=True)
            else:
                st.info("Aucune demande enregistrée.")
        except Exception as e:
            st.error(f"Erreur historique : {e}")

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
        ws_import = sheet.worksheet("IMPORT")
        data = ws_demandes.get_all_records()
        df = pd.DataFrame(data)

        if df.empty:
            st.info("Aucune demande enregistrée.")
            st.stop()

        # Filtres
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            bureaux_dispo = ["Tous"] + sorted(df["BUREAU"].dropna().unique().tolist())
            filtre_bureau = st.selectbox("Bureau", bureaux_dispo)
        with col2:
            if filtre_bureau != "Tous":
                agences_dispo = ["Toutes"] + sorted(df[df["BUREAU"] == filtre_bureau]["CODE AGENCE"].astype(str).dropna().unique().tolist())
            else:
                agences_dispo = ["Toutes"] + sorted(df["CODE AGENCE"].astype(str).dropna().unique().tolist())
            filtre_agence = st.selectbox("Agence", agences_dispo)
        with col3:
            if filtre_bureau != "Tous":
                clients_dispo = ["Tous"] + sorted(df[df["BUREAU"] == filtre_bureau]["CLIENT"].dropna().unique().tolist())
            else:
                clients_dispo = ["Tous"] + sorted(df["CLIENT"].dropna().unique().tolist())
            filtre_client = st.selectbox("Client", clients_dispo)
        with col4:
            dates_dispo = ["Toutes"] + sorted(df["DATE SAISIE"].str[:10].dropna().unique().tolist(), reverse=True)
            filtre_date = st.selectbox("Date", dates_dispo)

        # Application des filtres
        df_attente = df[df["STATUT"] == "EN ATTENTE"].copy()
        if filtre_bureau != "Tous":
            df_attente = df_attente[df_attente["BUREAU"] == filtre_bureau]
        if filtre_agence != "Toutes":
            df_attente = df_attente[df_attente["CODE AGENCE"].astype(str) == filtre_agence]
        if filtre_client != "Tous":
            df_attente = df_attente[df_attente["CLIENT"] == filtre_client]
        if filtre_date != "Toutes":
            df_attente = df_attente[df_attente["DATE SAISIE"].str[:10] == filtre_date]

        if df_attente.empty:
            st.info("Aucune demande en attente pour ces critères.")
        else:
            st.subheader(f"{len(df_attente)} demande(s) EN ATTENTE")
            for idx, row in df_attente.iterrows():
                col1, col2, col3 = st.columns([4, 1, 1])
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
                    if st.button("🔴 À traiter", key=f"traite_{idx}"):
                        col_statut = df.columns.tolist().index("STATUT") + 1
                        ws_demandes.update_cell(idx + 2, col_statut, "TRAITE")
                        # Stocker aussi DATE FIN MISSION et DATE DEBUT MISSION pour le fallback
                        ws_import.append_row([
                            str(row["MATRICULE MISSION"]),
                            "",
                            "",
                            1,
                            row["MONTANT"],
                            1,
                            "",
                            "",
                            "",
                            str(row["DATE FIN THEORIQUE DERNIERE MISSION"]),
                            str(row.get("DATE DEBUT MISSION", ""))
                        ])
                        st.rerun()
                with col3:
                    if st.button("⬛ Annuler", key=f"annuler_{idx}"):
                        col_statut = df.columns.tolist().index("STATUT") + 1
                        ws_demandes.update_cell(idx + 2, col_statut, "ANNULE")
                        st.rerun()

        # =====================
        # EXPORT FICHIER IMPORT
        # =====================
        st.markdown("---")
        st.subheader("📤 Export fichier d'import Evolia")

        mois_labels = [o[2] for o in mois_options]
        mois_choisi_label = st.selectbox("Mois de paie", mois_labels, index=0)
        mois_choisi = next(o for o in mois_options if o[2] == mois_choisi_label)
        mois_num, annee_num = mois_choisi[0], mois_choisi[1]

        dernier_jour = calendar.monthrange(annee_num, mois_num)[1]
        fin_mois = datetime(annee_num, mois_num, dernier_jour)

        nb_en_attente = len(df[df["STATUT"] == "EN ATTENTE"])
        if nb_en_attente > 0:
            st.warning(f"⚠️ Attention : {nb_en_attente} demande(s) sont encore EN ATTENTE et ne seront pas incluses dans l'export.")

        import_data = ws_import.get_all_records()
        df_import = pd.DataFrame(import_data)

        # Charger l'onglet SALARIES pour le fallback
        ws_sal = sheet.worksheet("SALARIES")
        df_salaries = pd.DataFrame(ws_sal.get_all_records())

        if st.button("📥 Exporter le fichier d'import"):
            if df_import.empty:
                st.error("Aucune ligne dans l'onglet IMPORT à exporter.")
            else:
                lignes = []
                header = "code Mission;rubrique;Libellé de la rubrique;base payé;taux payé;base facturé;taux facturé;date (choix de la semaine);Commentaire rubrique"
                lignes.append(header)

                for _, row_imp in df_import.iterrows():
                    code_mission = str(row_imp["code Mission"])
                    montant = row_imp["taux payé"]
                    date_fin_str = str(row_imp.get("DATE FIN MISSION", ""))
                    date_debut_str = str(row_imp.get("DATE DEBUT MISSION", ""))

                    # Calcul avec fallback
                    code_mission_final, date_lundi = calculer_lundi_avec_fallback(
                        code_mission, date_fin_str, date_debut_str, fin_mois, df_salaries
                    )

                    ligne = f"{code_mission_final};;;1;{montant};1;;{date_lundi};"
                    lignes.append(ligne)

                csv_content = "\n".join(lignes)
                csv_bytes = csv_content.encode("latin-1", errors="replace")
                nom_fichier = f"import_acomptes_{MOIS_FR[mois_num]}_{annee_num}.csv"

                st.download_button(
                    label="⬇️ Télécharger",
                    data=csv_bytes,
                    file_name=nom_fichier,
                    mime="text/csv"
                )

                col_statut = df.columns.tolist().index("STATUT") + 1
                for i, row in df.iterrows():
                    if row["STATUT"] == "TRAITE":
                        ws_demandes.update_cell(i + 2, col_statut, "IMPORTE")

                ws_import.clear()
                ws_import.append_row([
                    "code Mission", "rubrique", "Libellé de la rubrique",
                    "base payé", "taux payé", "base facturé",
                    "taux facturé", "date (choix de la semaine)", "Commentaire rubrique",
                    "DATE FIN MISSION", "DATE DEBUT MISSION"
                ])

                st.success(f"✅ Export généré : {nom_fichier} — onglet IMPORT vidé — statuts mis à jour")

        # Historique complet gestionnaire
        st.markdown("---")
        st.subheader("📋 Historique complet des demandes")
        try:
            ws_hist = sheet.worksheet("DEMANDES")
            hist_data = ws_hist.get_all_records()
            df_hist = pd.DataFrame(hist_data)
            if not df_hist.empty:
                df_hist = df_hist.sort_values("DATE SAISIE", ascending=False)
                df_hist = df_hist[[
                    "DATE SAISIE", "BUREAU", "CODE AGENCE", "NOM", "PRENOM",
                    "CLIENT", "MATRICULE MISSION",
                    "MONTANT", "COMMENTAIRE", "STATUT"
                ]]
                st.dataframe(df_hist, use_container_width=True)
            else:
                st.info("Aucune demande enregistrée.")
        except Exception as e:
            st.error(f"Erreur historique : {e}")

    except Exception as e:
        st.error(f"Erreur de connexion : {e}")
