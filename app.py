"""
Registre Labo — Personnel & Équipements
Version connectée à Google Sheets, avec mot de passe.
"""
import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Registre Labo", page_icon="🧪", layout="wide")

PERSONNEL_COLS = [
    "NOMS ET PRENOMS", "FONCTION", "Parcours scolaire", "DATE D'AMBAUCHE",
    "Département", "Type EPI", "Taille EPI",
    "Date de récupération EPI", "Date prochaine récupération EPI",
]
EQUIP_COLS = [
    "Identifiant", "Nom / type", "Constructeur", "Numéro de série", "Criticité",
    "Emplacement", "Date d'achat",
    "Suivi requis (Étalonnage / Maintenance / Vérification)",
    "Type étalonnage (Interne / Externe / Les deux)",
    "Fréquence étalonnage interne (jours)", "Dernier étalonnage interne",
    "Fréquence étalonnage externe (jours)", "Dernier étalonnage externe",
    "Dernière maintenance", "Dernière vérification",
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ---------- Mot de passe ----------

def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["app_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct", False):
        return True

    st.title("🧪 Registre Labo")
    st.text_input("Mot de passe", type="password", on_change=password_entered, key="password")
    if "password_correct" in st.session_state and not st.session_state["password_correct"]:
        st.error("Mot de passe incorrect.")
    return False


if not check_password():
    st.stop()

# ---------- Connexion Google Sheets ----------

@st.cache_resource
def get_client():
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
    return gspread.authorize(creds)


def get_sheet():
    return get_client().open(st.secrets["sheet_name"])


def load_data():
    sh = get_sheet()
    personnel_ws = sh.worksheet("Personnel")
    equip_ws = sh.worksheet("Equipements")
    p_records = personnel_ws.get_all_records()
    e_records = equip_ws.get_all_records()
    personnel = pd.DataFrame(p_records) if p_records else pd.DataFrame(columns=PERSONNEL_COLS)
    equip = pd.DataFrame(e_records) if e_records else pd.DataFrame(columns=EQUIP_COLS)
    return personnel, equip


def save_personnel(df):
    ws = get_sheet().worksheet("Personnel")
    ws.clear()
    clean = df.fillna("").astype(str)
    ws.update([clean.columns.tolist()] + clean.values.tolist())


def save_equip(df):
    ws = get_sheet().worksheet("Equipements")
    ws.clear()
    clean = df.fillna("").astype(str)
    ws.update([clean.columns.tolist()] + clean.values.tolist())


if "personnel" not in st.session_state or "equip" not in st.session_state:
    p, e = load_data()
    st.session_state.personnel = p
    st.session_state.equip = e

# ---------- Statuts calculés ----------

def make_calib_status(last_col, freq_col):
    def _status(row):
        try:
            last = pd.to_datetime(row[last_col])
            freq = float(row[freq_col])
        except Exception:
            return "⚪ Non renseigné"
        due = last + pd.Timedelta(days=freq)
        days_left = (due - pd.Timestamp.now()).days
        if days_left < 0:
            return f"🔴 En retard ({abs(days_left)} j)"
        elif days_left <= 1:
            return "🟠 À faire aujourd'hui/demain"
        return f"🟢 OK ({days_left} j restants)"
    return _status


calib_status_interne = make_calib_status("Dernier étalonnage interne", "Fréquence étalonnage interne (jours)")
calib_status_externe = make_calib_status("Dernier étalonnage externe", "Fréquence étalonnage externe (jours)")


def epi_status(row):
    try:
        expiry = pd.to_datetime(row["Date prochaine récupération EPI"])
    except Exception:
        return "⚪ Non renseigné"
    days_left = (expiry - pd.Timestamp.now()).days
    if days_left < 0:
        return f"🔴 Expiré ({abs(days_left)} j)"
    elif days_left <= 30:
        return f"🟠 À renouveler ({days_left} j)"
    return f"🟢 OK ({days_left} j)"


# ---------- Navigation ----------

st.sidebar.title("🧪 Registre Labo")
if st.sidebar.button("🔓 Se déconnecter"):
    st.session_state["password_correct"] = False
    st.rerun()

page = st.sidebar.radio("Navigation", ["Tableau de bord", "Personnel", "Équipements", "Importer un Excel"])

# ---------- Tableau de bord ----------

if page == "Tableau de bord":
    st.title("Tableau de bord")
    col1, col2 = st.columns(2)
    col1.metric("Employés enregistrés", len(st.session_state.personnel))
    col2.metric("Équipements enregistrés", len(st.session_state.equip))

    if not st.session_state.equip.empty:
        eq = st.session_state.equip.copy()
        eq["Statut interne"] = eq.apply(calib_status_interne, axis=1)
        eq["Statut externe"] = eq.apply(calib_status_externe, axis=1)

        st.subheader("Alertes étalonnage interne")
        alertes_int = eq[eq["Statut interne"].str.contains("🔴|🟠", na=False)]
        if alertes_int.empty:
            st.success("Aucune alerte d'étalonnage interne en cours.")
        else:
            st.dataframe(alertes_int[["Identifiant", "Nom / type", "Statut interne"]], use_container_width=True)

        st.subheader("Alertes étalonnage externe")
        alertes_ext = eq[eq["Statut externe"].str.contains("🔴|🟠", na=False)]
        if alertes_ext.empty:
            st.success("Aucune alerte d'étalonnage externe en cours.")
        else:
            st.dataframe(alertes_ext[["Identifiant", "Nom / type", "Statut externe"]], use_container_width=True)

    if not st.session_state.personnel.empty:
        st.subheader("Alertes EPI")
        pe = st.session_state.personnel.copy()
        pe["Statut EPI"] = pe.apply(epi_status, axis=1)
        alertes_epi = pe[pe["Statut EPI"].str.contains("🔴|🟠", na=False)]
        if alertes_epi.empty:
            st.success("Aucune alerte EPI en cours.")
        else:
            st.dataframe(alertes_epi[["NOMS ET PRENOMS", "Type EPI", "Statut EPI"]], use_container_width=True)

# ---------- Personnel ----------

elif page == "Personnel":
    st.title("Personnel")
    st.caption(
        "Modifie directement le tableau comme dans Excel, puis clique sur Sauvegarder. "
        "Une ligne = une personne + un EPI. Quand quelqu'un reçoit un **nouvel** EPI, "
        "ajoute une **nouvelle ligne** avec son nom (ne modifie pas l'ancienne) — l'ancienne "
        "ligne reste comme historique/archive. Types d'EPI courants : Chaussures, Combinaison "
        "ou Blouse, Masque, Lunettes, Casque antibruit. Dates au format AAAA-MM-JJ (ex. 2026-07-17)."
    )

    st.subheader("🔍 Historique par employé")
    recherche = st.text_input("Rechercher un nom (archive complète de ses EPI)")
    if recherche.strip():
        resultats = st.session_state.personnel[
            st.session_state.personnel["NOMS ET PRENOMS"].astype(str).str.contains(
                recherche, case=False, na=False
            )
        ].copy()
        if resultats.empty:
            st.info("Aucun employé trouvé avec ce nom.")
        else:
            resultats["Statut EPI"] = resultats.apply(epi_status, axis=1)
            resultats = resultats.sort_values("Date de récupération EPI", ascending=False)
            st.dataframe(
                resultats[["NOMS ET PRENOMS", "FONCTION", "Type EPI", "Taille EPI",
                           "Date de récupération EPI", "Date prochaine récupération EPI", "Statut EPI"]],
                use_container_width=True,
            )

    st.subheader("Tableau complet")
    edited = st.data_editor(
        st.session_state.personnel,
        num_rows="dynamic",
        use_container_width=True,
        key="personnel_editor",
    )
    st.session_state.personnel = edited
    if st.button("💾 Sauvegarder Personnel", type="primary"):
        save_personnel(st.session_state.personnel)
        st.success("Données sauvegardées dans Google Sheets.")

# ---------- Équipements ----------

elif page == "Équipements":
    st.title("Équipements")
    st.caption(
        "Modifie directement le tableau, puis clique sur Sauvegarder. Dates au format AAAA-MM-JJ."
    )
    with st.expander("ℹ️ Comment remplir les colonnes étalonnage / maintenance / vérification"):
        st.markdown(
            "- **Suivi requis** : écris ce qui s'applique pour cet équipement, ex. `Étalonnage, "
            "Maintenance` ou `Étalonnage, Maintenance, Vérification`.\n"
            "- **Type étalonnage** : `Interne`, `Externe`, ou `Interne et Externe` si les deux "
            "s'appliquent (l'interne est fait par le service qualité, l'externe par un audit).\n"
            "- **Fréquences (en jours)** : 1 = quotidien, 3-4 = 2x/semaine, 7 = hebdomadaire, "
            "30 = mensuel, 365 = annuel, 730 = tous les 2 ans. Exemple typique : étalonnage interne "
            "quotidien (1) pour une balance, étalonnage externe annuel (365) fait par l'audit.\n"
            "- Si un suivi ne s'applique pas à un équipement (ex. pas d'étalonnage externe), "
            "laisse les cases correspondantes vides."
        )
    edited = st.data_editor(
        st.session_state.equip,
        num_rows="dynamic",
        use_container_width=True,
        key="equip_editor",
    )
    st.session_state.equip = edited
    if st.button("💾 Sauvegarder Équipements", type="primary"):
        save_equip(st.session_state.equip)
        st.success("Données sauvegardées dans Google Sheets.")

# ---------- Import depuis l'Excel existant ----------

elif page == "Importer un Excel":
    st.title("Importer un fichier Excel existant")
    st.caption(
        "Charge le fichier que RQ t'a donné. Si le fichier a un en-tête de document "
        "(titre, référence qualité...) au-dessus du vrai tableau, indique la ligne où "
        "commencent les vrais titres de colonnes."
    )
    cible = st.radio("Ce fichier concerne :", ["Personnel", "Équipements"], horizontal=True)
    attendu = PERSONNEL_COLS if cible == "Personnel" else EQUIP_COLS
    st.write("Colonnes attendues par l'application :", ", ".join(attendu))

    uploaded = st.file_uploader("Fichier Excel (.xlsx)", type=["xlsx"])
    if uploaded is not None:
        header_row = st.number_input(
            "Numéro de la ligne contenant les vrais titres de colonnes (1 = première ligne du fichier)",
            min_value=1, max_value=30, value=1, step=1,
        )
        df_new = pd.read_excel(uploaded, header=header_row - 1)
        df_new = df_new.dropna(axis=1, how="all")  # retire les colonnes totalement vides
        df_new = df_new.dropna(how="all")  # retire les lignes totalement vides
        noms = [str(c).strip() for c in df_new.columns]
        vus = {}
        noms_uniques = []
        for n in noms:
            if n in vus:
                vus[n] += 1
                noms_uniques.append(f"{n} ({vus[n]})")
            else:
                vus[n] = 0
                noms_uniques.append(n)
        df_new.columns = noms_uniques

        st.write("Aperçu avec cette ligne d'en-tête :")
        st.dataframe(df_new.head(10), use_container_width=True)

        st.subheader("Faire correspondre les colonnes")

        mode_multi_epi = False
        if cible == "Personnel":
            mode_multi_epi = st.checkbox(
                "Mon fichier a une colonne séparée pour chaque type d'EPI "
                "(ex. « Taille chaussures », « Date récup lunettes »...) plutôt qu'une seule "
                "colonne « Type EPI »"
            )

        if cible == "Personnel" and mode_multi_epi:
            # ---- Mode spécial : une personne = plusieurs colonnes EPI (une par type) ----
            base_cols = ["NOMS ET PRENOMS", "FONCTION", "Parcours scolaire", "DATE D'AMBAUCHE", "Département"]
            st.markdown("**1. Informations générales de la personne**")
            options_base = ["-- Ignorer --"] + base_cols
            mapping_base = {}
            for col in df_new.columns:
                default_idx = options_base.index(col) if col in options_base else 0
                choix = st.selectbox(f"« {col} » →", options_base, index=default_idx, key=f"mapb_{col}")
                if choix != "-- Ignorer --":
                    mapping_base[col] = choix

            valeurs_base = list(mapping_base.values())
            doublons_base = sorted({v for v in valeurs_base if valeurs_base.count(v) > 1})
            if doublons_base:
                st.error(f"⚠️ Choisies plusieurs fois : {', '.join(doublons_base)}. Corrige avant de continuer.")

            st.markdown("**2. Colonnes EPI, une par type**")
            st.caption("Pour chaque type d'EPI présent dans ton fichier, indique ses 3 colonnes (ou « Aucune » si le type n'existe pas chez toi).")
            epi_types_liste = ["Chaussures", "Combinaison ou Blouse", "Masque", "Lunettes", "Casque antibruit"]
            colonnes_dispo = ["-- Aucune --"] + list(df_new.columns)
            epi_mapping = {}
            for epi_type in epi_types_liste:
                with st.expander(f"EPI : {epi_type}"):
                    c1, c2, c3 = st.columns(3)
                    col_taille = c1.selectbox("Colonne Taille", colonnes_dispo, key=f"epi_taille_{epi_type}")
                    col_recup = c2.selectbox("Colonne Date de récupération", colonnes_dispo, key=f"epi_recup_{epi_type}")
                    col_next = c3.selectbox("Colonne Date prochaine récupération", colonnes_dispo, key=f"epi_next_{epi_type}")
                    epi_mapping[epi_type] = (col_taille, col_recup, col_next)

            if st.button("Ajouter ces lignes", disabled=bool(doublons_base)):
                if not mapping_base:
                    st.error("Fais correspondre au moins une colonne d'information générale.")
                else:
                    base_df = df_new[list(mapping_base.keys())].rename(columns=mapping_base)
                    for c in base_cols:
                        if c not in base_df.columns:
                            base_df[c] = ""
                    lignes = []
                    for i in range(len(df_new)):
                        base_row = base_df.iloc[i]
                        for epi_type, (col_taille, col_recup, col_next) in epi_mapping.items():
                            if col_taille == "-- Aucune --" and col_recup == "-- Aucune --" and col_next == "-- Aucune --":
                                continue
                            taille_val = df_new.iloc[i][col_taille] if col_taille != "-- Aucune --" else ""
                            recup_val = df_new.iloc[i][col_recup] if col_recup != "-- Aucune --" else ""
                            next_val = df_new.iloc[i][col_next] if col_next != "-- Aucune --" else ""
                            if pd.isna(taille_val) and pd.isna(recup_val) and pd.isna(next_val):
                                continue
                            lignes.append({
                                "NOMS ET PRENOMS": base_row.get("NOMS ET PRENOMS", ""),
                                "FONCTION": base_row.get("FONCTION", ""),
                                "Parcours scolaire": base_row.get("Parcours scolaire", ""),
                                "DATE D'AMBAUCHE": base_row.get("DATE D'AMBAUCHE", ""),
                                "Département": base_row.get("Département", ""),
                                "Type EPI": epi_type,
                                "Taille EPI": "" if pd.isna(taille_val) else taille_val,
                                "Date de récupération EPI": "" if pd.isna(recup_val) else recup_val,
                                "Date prochaine récupération EPI": "" if pd.isna(next_val) else next_val,
                            })
                    if not lignes:
                        st.warning("Aucune ligne EPI détectée avec ce mappage — vérifie tes choix ci-dessus.")
                    else:
                        df_clean = pd.DataFrame(lignes)[PERSONNEL_COLS]
                        st.session_state.personnel = pd.concat(
                            [st.session_state.personnel, df_clean], ignore_index=True
                        )
                        save_personnel(st.session_state.personnel)
                        st.success(f"{len(df_clean)} ligne(s) ajoutée(s) et sauvegardée(s) dans Google Sheets.")

        else:
            # ---- Mode simple : correspondance directe colonne par colonne ----
            st.caption(
                "Pour chaque colonne de ton fichier, choisis à quelle colonne de l'application "
                "elle correspond, ou laisse « Ignorer » si elle n'est pas utile."
            )
            options = ["-- Ignorer --"] + attendu
            mapping = {}
            for col in df_new.columns:
                default_idx = options.index(col) if col in options else 0
                choix = st.selectbox(f"« {col} » →", options, index=default_idx, key=f"map_{cible}_{col}")
                if choix != "-- Ignorer --":
                    mapping[col] = choix

            colonnes_manquantes = [c for c in attendu if c not in mapping.values()]
            if colonnes_manquantes:
                st.info(f"Colonnes de l'app non renseignées (resteront vides) : {', '.join(colonnes_manquantes)}")

            valeurs = list(mapping.values())
            doublons = sorted({v for v in valeurs if valeurs.count(v) > 1})
            if doublons:
                st.error(
                    f"⚠️ Ces colonnes de l'application ont été choisies plusieurs fois : "
                    f"{', '.join(doublons)}. Chaque colonne de l'app ne peut être utilisée qu'une "
                    f"seule fois — corrige les menus déroulants ci-dessus avant de continuer."
                )

            if st.button("Ajouter ces lignes", disabled=bool(doublons)):
                if not mapping:
                    st.error("Fais correspondre au moins une colonne avant d'ajouter.")
                else:
                    df_clean = df_new[list(mapping.keys())].rename(columns=mapping)
                    for c in attendu:
                        if c not in df_clean.columns:
                            df_clean[c] = ""
                    df_clean = df_clean[attendu]
                    if cible == "Personnel":
                        st.session_state.personnel = pd.concat(
                            [st.session_state.personnel, df_clean], ignore_index=True
                        )
                        save_personnel(st.session_state.personnel)
                    else:
                        st.session_state.equip = pd.concat(
                            [st.session_state.equip, df_clean], ignore_index=True
                        )
                        save_equip(st.session_state.equip)
                    st.success(f"{len(df_clean)} ligne(s) ajoutée(s) et sauvegardée(s) dans Google Sheets.")

st.sidebar.markdown("---")
st.sidebar.caption(f"Session ouverte : {datetime.now().strftime('%d/%m/%Y %H:%M')}")
