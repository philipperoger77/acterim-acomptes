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

def calculer_lundi(date_fin_str, fin_mois):
    """Calcule le lundi de la semaine selon la règle métier"""
    try:
        date_fin = datetime.strptime(date_fin_str, "%d/%m/%Y")
    except:
        try:
            date_fin = datetime.strptime(date_fin_str, "%Y-%m-%d")
        except:
            return ""
    
    # Si date fin > fin du mois choisi, on prend la fin du mois comme référence
    if date_fin > fin_mois:
        date_ref = fin_mois
    else:
        date_ref = date_fin
    
    # Lundi de la semaine de date_ref (weekday: lundi=0)
    lundi = date_ref - timedelta(days=date_ref.weekday())
    return lundi.strftime("%d/%m/%Y")

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
                df_en_attente["MATRICULE DERNIERE MISSION"].astype(str)
            ).tolist()

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
                cle = mat + "_" + str(row["CODE AGENCE"]) + "_" + str(row["MATRICULE DERNIERE MISSION"])
                if cle in en_attente:
                    st.warning("⚠️ Demande déjà en attente")
            with col2:
                montants[mat] = st.number_input(
                    "Montant (€)", min_value=0.0, step=10.0,
                    key=f"montant_{mat}_{row['MATRICULE DERNIERE MISSION']}"
                )
            with col3:
                commentaires[mat] = st.text_input(
                    "Commentaire", key=f"commentaire_{mat}_{row['MATRICULE DERNIERE MISSION']}"
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
                cle = mat + "_" + str(row["CODE AGENCE"]) + "_" + str(row["MATRICULE DERNIERE MISSION"])
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
                    str(row["DATE FIN THEORIQUE DERNIERE MISSION"]),
                    str(row["MATRICULE DERNIERE MISSION"]),
                    montant,
                    commentaires[mat],
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
                        "CLIENT", "MATRICULE DERNIERE MISSION",
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
                        # Mise à jour STATUT -> TRAITE dans DEMANDES
                        col_statut = df.columns.tolist().index("STATUT") + 1
                        ws_demandes.update_cell(idx + 2, col_statut, "TRAITE")
                        # Ajout dans onglet IMPORT
                        ws_import.append_row([
                            str(row["MATRICULE DERNIERE MISSION"]),  # code Mission
                            "",                                        # rubrique
                            "",                                        # Libellé de la rubrique
                            1,                                         # base payé
                            row["MONTANT"],                            # taux payé
                            1,                                         # base facturé
                            "",                                        # taux facturé
                            "",                                        # date (sera calculée à l'export)
                            ""                                         # Commentaire rubrique
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

        # Menu déroulant mois/année
        now = datetime.now()
        mois_options = []
        for i in range(-2, 4):
            m = (now.month + i - 1) % 12 + 1
            y = now.year + (now.month + i - 1) // 12
            mois_options.append((m, y, f"{MOIS_FR[m]} {y}"))

        mois_labels = [o[2] for o in mois_options]
        mois_idx_defaut = 2
        mois_choisi_label = st.selectbox("Mois de paie", mois_labels, index=mois_idx_defaut)
        mois_choisi = next(o for o in mois_options if o[2] == mois_choisi_label)
        mois_num, annee_num = mois_choisi[0], mois_choisi[1]

        # Fin du mois choisi
        dernier_jour = calendar.monthrange(annee_num, mois_num)[1]
        fin_mois = datetime(annee_num, mois_num, dernier_jour)

        # Vérification EN ATTENTE
        nb_en_attente = len(df[df["STATUT"] == "EN ATTENTE"])
        if nb_en_attente > 0:
            st.warning(f"⚠️ Attention : {nb_en_attente} demande(s) sont encore EN ATTENTE et ne seront pas incluses dans l'export.")

        # Lecture onglet IMPORT
        import_data = ws_import.get_all_records()
        df_import = pd.DataFrame(import_data)

        if st.button("📥 Exporter le fichier d'import"):
            if df_import.empty:
                st.error("Aucune ligne dans l'onglet IMPORT à exporter.")
            else:
                # Récupérer les dates fin mission depuis DEMANDES pour calcul
                df_traite = df[df["STATUT"] == "TRAITE"].copy()

                # Construire le CSV ligne par ligne
                lignes = []
                header = "code Mission;rubrique;Libellé de la rubrique;base payé;taux payé;base facturé;taux facturé;date (choix de la semaine);Commentaire rubrique"
                lignes.append(header)

                for _, row_imp in df_import.iterrows():
                    code_mission = str(row_imp["code Mission"])
                    montant = row_imp["taux payé"]

                    # Trouver la date fin mission correspondante dans DEMANDES
                    match = df_traite[df_traite["MATRICULE DERNIERE MISSION"].astype(str) == code_mission]
                    if not match.empty:
                        date_fin_str = str(match.iloc[0]["DATE FIN THEORIQUE DERNIERE MISSION"])
                        date_lundi = calculer_lundi(date_fin_str, fin_mois)
                    else:
                        date_lundi = calculer_lundi("", fin_mois)

                    ligne = f"{code_mission};;;1;{montant};1;;{date_lundi};"
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

                # Passage TRAITE -> IMPORTE dans DEMANDES
                col_statut = df.columns.tolist().index("STATUT") + 1
                for i, row in df.iterrows():
                    if row["STATUT"] == "TRAITE":
                        ws_demandes.update_cell(i + 2, col_statut, "IMPORTE")

                # Vidage onglet IMPORT
                ws_import.clear()
                ws_import.append_row([
                    "code Mission", "rubrique", "Libellé de la rubrique",
                    "base payé", "taux payé", "base facturé",
                    "taux facturé", "date (choix de la semaine)", "Commentaire rubrique"
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
                    "CLIENT", "MATRICULE DERNIERE MISSION",
                    "MONTANT", "COMMENTAIRE", "STATUT"
                ]]
                st.dataframe(df_hist, use_container_width=True)
            else:
                st.info("Aucune demande enregistrée.")
        except Exception as e:
            st.error(f"Erreur historique : {e}")

    except Exception as e:
        st.error(f"Erreur de connexion : {e}")
