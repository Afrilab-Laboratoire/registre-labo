"""
Registre Labo — Personnel & Équipements
Version connectée à Google Sheets, avec mot de passe.
Tableau Personnel en format large : une ligne par personne.
"""
import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Registre Labo", page_icon="🧪", layout="wide")

EPI_TYPES = ["Chaussures", "Combinaison ou Blouse", "Masque", "Lunettes", "Casque antibruit"]

PERSONNEL_BASE_COLS = ["NOMS ET PRENOMS", "FONCTION", "Parcours scolaire", "DATE D'AMBAUCHE", "Département"]
PERSONNEL_EPI_COLS = []
for _t in EPI_TYPES:
    PERSONNEL_EPI_COLS += [f"Taille {_t}", f"Date récupération {_t}", f"Date prochaine récupération {_t}"]
PERSONNEL_COLS = PERSONNEL_BASE_COLS + PERSONNEL_EPI_COLS

EQUIP_COLS = [
    "Identifiant", "Nom / type", "Constructeur", "Numéro de série", "Criticité",
    "Emplacement", "Date d'achat",
    "Suivi requis (Étalonnage / Maintenance / Vérification)",
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


def log_epi_changes(old_df, new_df):
    """Archive dans l'onglet Historique_EPI tout EPI dont les infos changent entre old_df et new_df."""
    try:
        ws = get_sheet().worksheet("Historique_EPI")
    except gspread.exceptions.WorksheetNotFound:
        return None  # onglet pas encore créé

    today = datetime.now().strftime("%Y-%m-%d %H:%M")
    rows_to_log = []
    for i in range(min(len(old_df), len(new_df))):
        old_row = old_df.iloc[i]
        new_row = new_df.iloc[i]
        name = str(old_row.get("NOMS ET PRENOMS", ""))
        for t in EPI_TYPES:
            cols = [f"Taille {t}", f"Date récupération {t}", f"Date prochaine récupération {t}"]
            old_vals = [str(old_row.get(c, "")).strip() for c in cols]
            new_vals = [str(new_row.get(c, "")).strip() for c in cols]
            if old_vals != new_vals and any(old_vals):
                rows_to_log.append([name, t] + old_vals + [today])
    # Lignes supprimées entièrement (personne retirée du tableau) : archive aussi leurs EPI
    for i in range(len(new_df), len(old_df)):
        old_row = old_df.iloc[i]
        name = str(old_row.get("NOMS ET PRENOMS", ""))
        for t in EPI_TYPES:
            cols = [f"Taille {t}", f"Date récupération {t}", f"Date prochaine récupération {t}"]
            old_vals = [str(old_row.get(c, "")).strip() for c in cols]
            if any(old_vals):
                rows_to_log.append([name, t] + old_vals + [today])

    if rows_to_log:
        ws.append_rows(rows_to_log)
    return len(rows_to_log)


def load_historique_epi():
    try:
        ws = get_sheet().worksheet("Historique_EPI")
    except gspread.exceptions.WorksheetNotFound:
        return None
    records = ws.get_all_records()
    cols = ["NOMS ET PRENOMS", "Type EPI", "Taille EPI (ancienne)",
            "Date de récupération EPI (ancienne)", "Date prochaine récupération EPI (ancienne)", "Archivé le"]
    return pd.DataFrame(records) if records else pd.DataFrame(columns=cols)


if "personnel" not in st.session_state or "equip" not in st.session_state:
    p, e = load_data()
    st.session_state.personnel = p
    st.session_state.personnel_baseline = p.copy()
    st.session_state.equip = e

# ---------- Conversion des dates Excel (numéros de série -> AAAA-MM-JJ) ----------

def convert_excel_date_value(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or v == "":
        return ""
    if isinstance(v, pd.Timestamp):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(v))).strftime("%Y-%m-%d")
        except Exception:
            return str(v)
    try:
        parsed = pd.to_datetime(v, errors="raise", dayfirst=True)
        return parsed.strftime("%Y-%m-%d")
    except Exception:
        return str(v)


def fix_date_columns(df):
    df = df.copy()
    for col in df.columns:
        if "date" in col.lower():
            df[col] = df[col].apply(convert_excel_date_value)
    return df


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


def epi_type_status(row, epi_type):
    """Renvoie un message d'alerte pour ce type d'EPI, ou None si tout va bien / non renseigné."""
    col = f"Date prochaine récupération {epi_type}"
    val = row.get(col, "")
    if val in ("", None) or (isinstance(val, float) and pd.isna(val)):
        return None
    try:
        expiry = pd.to_datetime(val)
    except Exception:
        return None
    days_left = (expiry - pd.Timestamp.now()).days
    if days_left < 0:
        return f"🔴 {epi_type} expiré depuis {abs(days_left)} j"
    elif days_left <= 30:
        return f"🟠 {epi_type} à renouveler dans {days_left} j"
    return None


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
        alertes_liste = []
        for _, prow in st.session_state.personnel.iterrows():
            for t in EPI_TYPES:
                s = epi_type_status(prow, t)
                if s:
                    alertes_liste.append({"NOMS ET PRENOMS": prow.get("NOMS ET PRENOMS", ""), "Alerte": s})
        if not alertes_liste:
            st.success("Aucune alerte EPI en cours.")
        else:
            st.dataframe(pd.DataFrame(alertes_liste), use_container_width=True)

# ---------- Personnel ----------

elif page == "Personnel":
    st.title("Personnel")
    st.caption(
        "Une ligne = une personne, avec une colonne Taille / Date de récupération / Date "
        "prochaine récupération pour chaque type d'EPI. Modifie directement le tableau comme "
        "dans Excel, puis clique sur Sauvegarder. Dates au format AAAA-MM-JJ (ex. 2026-07-17)."
    )

    st.subheader("🔍 Rechercher un employé")
    recherche = st.text_input("Nom de l'employé")
    if recherche.strip():
        resultats = st.session_state.personnel[
            st.session_state.personnel["NOMS ET PRENOMS"].astype(str).str.contains(
                recherche, case=False, na=False
            )
        ]
        if resultats.empty:
            st.info("Aucun employé trouvé avec ce nom.")
        else:
            for _, prow in resultats.iterrows():
                with st.container(border=True):
                    st.markdown(f"**{prow.get('NOMS ET PRENOMS', '')}** — {prow.get('FONCTION', '')}")
                    for t in EPI_TYPES:
                        taille = prow.get(f"Taille {t}", "")
                        recup = prow.get(f"Date récupération {t}", "")
                        prochaine = prow.get(f"Date prochaine récupération {t}", "")
                        if not taille and not recup and not prochaine:
                            continue
                        alerte = epi_type_status(prow, t)
                        statut_txt = alerte if alerte else "🟢 OK"
                        st.write(
                            f"- **{t}** — taille : {taille or '—'} · récupéré le : {recup or '—'} "
                            f"· prochaine : {prochaine or '—'} · {statut_txt}"
                        )

            hist = load_historique_epi()
            if hist is None:
                st.info(
                    "Pour voir l'archive des anciens EPI ici, crée l'onglet « Historique_EPI » "
                    "dans ton Google Sheet (voir les instructions données à côté)."
                )
            elif not hist.empty:
                hist_personne = hist[
                    hist["NOMS ET PRENOMS"].astype(str).str.contains(recherche, case=False, na=False)
                ]
                if not hist_personne.empty:
                    st.markdown("**📜 Archive — anciens EPI remplacés**")
                    st.dataframe(hist_personne, use_container_width=True)

    st.subheader("Tableau complet")
    edited = st.data_editor(
        st.session_state.personnel,
        num_rows="dynamic",
        use_container_width=True,
        key="personnel_editor",
    )
    st.session_state.personnel = edited
    if st.button("💾 Sauvegarder Personnel", type="primary"):
        nb_archives = log_epi_changes(st.session_state.personnel_baseline, edited)
        save_personnel(edited)
        st.session_state.personnel_baseline = edited.copy()
        if nb_archives is None:
            st.success(
                "Données sauvegardées dans Google Sheets. (Archive non activée : crée l'onglet "
                "« Historique_EPI » pour garder une trace des anciens EPI.)"
            )
        elif nb_archives > 0:
            st.success(f"Données sauvegardées. {nb_archives} ancien(s) EPI archivé(s) dans l'historique.")
        else:
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
            "- Si l'étalonnage interne (service qualité) s'applique, remplis sa fréquence et sa "
            "dernière date. Si l'étalonnage externe (audit) s'applique aussi, remplis les siennes. "
            "Si un seul des deux s'applique, laisse l'autre vide — pas besoin d'un champ séparé "
            "pour préciser lequel.\n"
            "- **Fréquences (en jours)** : 1 = quotidien, 3-4 = 2x/semaine, 7 = hebdomadaire, "
            "30 = mensuel, 365 = annuel, 730 = tous les 2 ans.\n"
            "- Si un suivi ne s'applique pas à un équipement, laisse les cases correspondantes vides."
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
        "commencent les vrais titres de colonnes. Les dates sont converties automatiquement, "
        "même si Excel les affiche comme de simples nombres."
    )
    cible = st.radio("Ce fichier concerne :", ["Personnel", "Équipements"], horizontal=True)
    attendu = PERSONNEL_COLS if cible == "Personnel" else EQUIP_COLS
    with st.expander("Voir les colonnes attendues par l'application"):
        st.write(", ".join(attendu))

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
        st.caption(
            "Pour chaque colonne de ton fichier, choisis à quelle colonne de l'application "
            "elle correspond, ou laisse « Ignorer » si elle n'est pas utile. Pour le Personnel, "
            "il y a une case Taille / Date récupération / Date prochaine récupération par type "
            "d'EPI (Chaussures, Lunettes...) — choisis la bonne pour chaque colonne de ton fichier."
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
            with st.expander(f"Colonnes de l'app non renseignées ({len(colonnes_manquantes)}) — resteront vides"):
                st.write(", ".join(colonnes_manquantes))

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
                df_clean = fix_date_columns(df_clean)
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
