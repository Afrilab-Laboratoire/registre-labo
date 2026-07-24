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

st.markdown("""
<style>
[data-testid="stDataFrame"] thead tr th,
[data-testid="stDataEditor"] thead tr th {
    background-color: #CDEEFB !important;
    color: #0B3554 !important;
}
</style>
""", unsafe_allow_html=True)

# ---------- Structure des données ----------

FIXED_CYCLE_EPI = ["Chaussures", "Combinaison ou Blouse", "Masque", "Casque antibruit"]
ON_DEMAND_EPI = ["Gants", "Lunettes"]
EPI_TYPES = FIXED_CYCLE_EPI + ON_DEMAND_EPI

PERSONNEL_BASE_COLS = ["NOMS ET PRENOMS", "FONCTION", "Parcours scolaire", "DATE D'AMBAUCHE", "Département"]
PERSONNEL_EPI_COLS = []
for _t in FIXED_CYCLE_EPI:
    PERSONNEL_EPI_COLS += [
        f"Taille {_t}", f"Date récupération {_t}", f"Date prochaine récupération {_t}",
        f"Cas exceptionnel {_t}", f"Confirmé {_t}",
    ]
for _t in ON_DEMAND_EPI:
    PERSONNEL_EPI_COLS += [
        f"Taille {_t}", f"Date récupération {_t}", f"Date prochaine récupération {_t}", f"Confirmé {_t}",
    ]
PERSONNEL_COLS = PERSONNEL_BASE_COLS + PERSONNEL_EPI_COLS

EQUIP_COLS = ["Identifiant", "Nom", "Constructeur", "Type", "Numéro de série",
              "Criticité", "Emplacement", "Date d'achat"]

NATURE_OPTIONS = ["Étalonnage externe", "Maintenance", "Vérification"]
FREQ_EXT_OPTIONS = ["Annuel", "Semestriel"]
FREQ_EXT_DAYS = {"Annuel": 365, "Semestriel": 182}
RESULTAT_OPTIONS = ["Conforme", "Non conforme"]

INTERVENTION_COLS = ["Identifiant", "Nom / Type", "Nature d'intervention",
                      "Fréquence prévue", "Date de l'intervention", "Résultat"]
BALANCE_COLS = ["Identifiant", "Nom / Type", "Date", "Résultat", "Remarque"]

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


def get_or_none(name):
    try:
        return get_sheet().worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        return None


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


def load_journal(sheet_name, cols):
    ws = get_or_none(sheet_name)
    if ws is None:
        return None
    records = ws.get_all_records()
    return pd.DataFrame(records) if records else pd.DataFrame(columns=cols)


def append_journal_row(sheet_name, row_values):
    ws = get_or_none(sheet_name)
    if ws is None:
        return False
    ws.append_row(row_values)
    return True


def log_epi_changes(old_df, new_df):
    try:
        ws = get_sheet().worksheet("Historique_EPI")
    except gspread.exceptions.WorksheetNotFound:
        return None
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
                exceptionnel = str(new_row.get(f"Cas exceptionnel {t}", "")).strip() if t in FIXED_CYCLE_EPI else ""
                rows_to_log.append([name, t] + old_vals + [exceptionnel, today])
    for i in range(len(new_df), len(old_df)):
        old_row = old_df.iloc[i]
        name = str(old_row.get("NOMS ET PRENOMS", ""))
        for t in EPI_TYPES:
            cols = [f"Taille {t}", f"Date récupération {t}", f"Date prochaine récupération {t}"]
            old_vals = [str(old_row.get(c, "")).strip() for c in cols]
            if any(old_vals):
                rows_to_log.append([name, t] + old_vals + ["", today])
    if rows_to_log:
        ws.append_rows(rows_to_log)
    return len(rows_to_log)


def load_historique_epi():
    ws = get_or_none("Historique_EPI")
    if ws is None:
        return None
    records = ws.get_all_records()
    cols = ["NOMS ET PRENOMS", "Type EPI", "Taille EPI (ancienne)",
            "Date de récupération EPI (ancienne)", "Date prochaine récupération EPI (ancienne)",
            "Changement exceptionnel", "Archivé le"]
    return pd.DataFrame(records) if records else pd.DataFrame(columns=cols)


if "personnel" not in st.session_state or "equip" not in st.session_state:
    p, e = load_data()
    st.session_state.personnel = p
    st.session_state.personnel_baseline = p.copy()
    st.session_state.equip = e

# ---------- Conversion des dates Excel ----------

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


def auto_fill_epi_expiration(df):
    df = df.copy()
    for t in FIXED_CYCLE_EPI:
        col_recup = f"Date récupération {t}"
        col_next = f"Date prochaine récupération {t}"
        if col_recup not in df.columns or col_next not in df.columns:
            continue
        for i in df.index:
            val = df.at[i, col_recup]
            if val in ("", None) or (isinstance(val, float) and pd.isna(val)):
                continue
            try:
                d = pd.to_datetime(val)
                df.at[i, col_next] = (d + pd.DateOffset(months=6)).strftime("%Y-%m-%d")
            except Exception:
                pass
    return df


# ---------- Statuts calculés ----------

def epi_type_status(row, epi_type):
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
        return f"🔴 Expiré depuis {abs(days_left)} j"
    elif days_left <= 30:
        return f"🟠 À renouveler dans {days_left} j"
    return f"🟢 OK ({days_left} j)"


def dernieres_interventions(interventions_df):
    """Renvoie, pour chaque (Identifiant, Nature), la ligne la plus récente."""
    if interventions_df is None or interventions_df.empty:
        return pd.DataFrame(columns=INTERVENTION_COLS)
    df = interventions_df.copy()
    df["_date"] = pd.to_datetime(df["Date de l'intervention"], errors="coerce")
    df = df.sort_values("_date")
    return df.groupby(["Identifiant", "Nature d'intervention"], as_index=False).last()


def intervention_status(row):
    freq = row.get("Fréquence prévue", "")
    date_val = row.get("Date de l'intervention", "")
    resultat = row.get("Résultat", "")
    if resultat == "Non conforme":
        return "🔴 Dernier résultat non conforme"
    if freq not in FREQ_EXT_DAYS or not date_val:
        return "⚪ Non renseigné"
    try:
        last_d = pd.to_datetime(date_val)
    except Exception:
        return "⚪ Non renseigné"
    due = last_d + pd.Timedelta(days=FREQ_EXT_DAYS[freq])
    days_left = (due - pd.Timestamp.now()).days
    if days_left < 0:
        return f"🔴 En retard de {abs(days_left)} j"
    elif days_left <= 15:
        return f"🟠 À prévoir dans {days_left} j"
    return f"🟢 OK ({days_left} j restants)"


def style_status_cell(val):
    if isinstance(val, str):
        if val.startswith("🔴"):
            return "background-color:#FBEAE7;color:#B0392E;"
        if val.startswith("🟠"):
            return "background-color:#FBEEDA;color:#B8710A;"
        if val.startswith("🟢"):
            return "background-color:#E7F3EB;color:#2F7A4F;"
    return ""


# ---------- Navigation ----------

st.sidebar.title("🧪 Registre Labo")
if st.sidebar.button("🔓 Se déconnecter"):
    st.session_state["password_correct"] = False
    st.rerun()

page = st.sidebar.radio(
    "Navigation",
    ["Tableau de bord", "Personnel", "Équipements", "Contrôle interne balances", "Importer un Excel"],
)

# ---------- Tableau de bord ----------

if page == "Tableau de bord":
    st.title("Tableau de bord")
    col1, col2 = st.columns(2)
    col1.metric("Employés enregistrés", len(st.session_state.personnel))
    col2.metric("Équipements enregistrés", len(st.session_state.equip))

    interventions = load_journal("Interventions_Externes", INTERVENTION_COLS)
    if interventions is None:
        st.info("Crée l'onglet « Interventions_Externes » dans Google Sheet pour activer ce suivi.")
    elif not interventions.empty:
        derniers = dernieres_interventions(interventions)
        derniers["Statut"] = derniers.apply(intervention_status, axis=1)
        alertes = derniers[derniers["Statut"].str.contains("🔴|🟠", na=False)]
        st.subheader("Alertes interventions externes")
        if alertes.empty:
            st.success("Aucune alerte en cours.")
        else:
            st.dataframe(alertes[["Identifiant", "Nom / Type", "Nature d'intervention", "Statut"]],
                         use_container_width=True)

    balances = load_journal("Controle_Interne_Balances", BALANCE_COLS)
    if balances is not None and not balances.empty:
        st.subheader("Contrôle interne balances — non-conformités récentes")
        non_conf = balances[balances["Résultat"] == "Non conforme"]
        if non_conf.empty:
            st.success("Aucune non-conformité enregistrée.")
        else:
            st.dataframe(non_conf.tail(10), use_container_width=True)

    if not st.session_state.personnel.empty:
        st.subheader("Alertes EPI")
        alertes_liste = []
        for _, prow in st.session_state.personnel.iterrows():
            for t in EPI_TYPES:
                s = epi_type_status(prow, t)
                if s and not s.startswith("🟢"):
                    alertes_liste.append({"NOMS ET PRENOMS": prow.get("NOMS ET PRENOMS", ""), "Alerte": s})
        if not alertes_liste:
            st.success("Aucune alerte EPI en cours.")
        else:
            st.dataframe(pd.DataFrame(alertes_liste), use_container_width=True)

# ---------- Personnel ----------

elif page == "Personnel":
    st.title("Personnel")
    st.caption(
        "Chaussures, Combinaison ou Blouse, Masque et Casque antibruit ont un cycle de 6 mois "
        "calculé automatiquement. Gants et Lunettes se changent à la demande."
    )

    tab_vue, tab_form, tab_couleur = st.tabs(["📋 Tableau complet", "➕ Ajouter / ✏️ Modifier", "🎨 Vue colorée"])

    with tab_vue:
        st.subheader("🔍 Rechercher un employé")
        recherche = st.text_input("Nom de l'employé", key="recherche_personnel")
        if recherche.strip():
            resultats = st.session_state.personnel[
                st.session_state.personnel["NOMS ET PRENOMS"].astype(str).str.contains(recherche, case=False, na=False)
            ]
            if resultats.empty:
                st.info("Aucun employé trouvé.")
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
                            alerte = epi_type_status(prow, t) or "⚪ Non renseigné"
                            st.write(f"- **{t}** — taille : {taille or '—'} · récupéré le : {recup or '—'} "
                                     f"· prochaine : {prochaine or '—'} · {alerte}")
                hist = load_historique_epi()
                if hist is not None and not hist.empty:
                    hist_p = hist[hist["NOMS ET PRENOMS"].astype(str).str.contains(recherche, case=False, na=False)]
                    if not hist_p.empty:
                        st.markdown("**📜 Archive**")
                        st.dataframe(hist_p, use_container_width=True)

        st.subheader("Tableau (édition libre, pour import en masse)")
        edited = st.data_editor(st.session_state.personnel, num_rows="dynamic", use_container_width=True, key="personnel_editor")
        st.session_state.personnel = edited
        if st.button("💾 Sauvegarder Personnel", type="primary"):
            edited_calc = auto_fill_epi_expiration(edited)
            nb_archives = log_epi_changes(st.session_state.personnel_baseline, edited_calc)
            save_personnel(edited_calc)
            st.session_state.personnel = edited_calc
            st.session_state.personnel_baseline = edited_calc.copy()
            st.success("Sauvegardé." + (f" {nb_archives} ancien(s) EPI archivé(s)." if nb_archives else ""))

    with tab_form:
        noms_existants = ["-- Nouvel employé --"] + [
            f"{i} — {row.get('NOMS ET PRENOMS', '')}" for i, row in st.session_state.personnel.iterrows()
        ]
        choix = st.selectbox("Choisir un employé à modifier, ou créer un nouveau", noms_existants)
        idx_edit = None
        if choix != "-- Nouvel employé --":
            idx_edit = int(choix.split(" — ")[0])
            row_data = st.session_state.personnel.loc[idx_edit]
        else:
            row_data = pd.Series({c: "" for c in PERSONNEL_COLS})

        with st.form("form_personnel"):
            st.markdown("**Informations générales**")
            nom = st.text_input("Noms et prénoms", value=row_data.get("NOMS ET PRENOMS", ""))
            fonction = st.text_input("Fonction", value=row_data.get("FONCTION", ""))
            parcours = st.text_input("Parcours scolaire", value=row_data.get("Parcours scolaire", ""))
            embauche = st.text_input("Date d'embauche (AAAA-MM-JJ)", value=row_data.get("DATE D'AMBAUCHE", ""))
            dept = st.text_input("Département", value=row_data.get("Département", ""))

            st.markdown("**EPI à cycle fixe (6 mois)**")
            valeurs_fixed = {}
            for t in FIXED_CYCLE_EPI:
                c1, c2, c3, c4 = st.columns(4)
                taille = c1.text_input(f"Taille {t}", value=row_data.get(f"Taille {t}", ""), key=f"f_taille_{t}")
                recup = c2.text_input(f"Récup. {t} (AAAA-MM-JJ)", value=row_data.get(f"Date récupération {t}", ""), key=f"f_recup_{t}")
                exceptionnel = c3.checkbox("Cas exceptionnel", value=str(row_data.get(f"Cas exceptionnel {t}", "")).strip().lower() == "oui", key=f"f_except_{t}")
                confirme = c4.checkbox("Confirmé", value=str(row_data.get(f"Confirmé {t}", "")).strip().lower() == "oui", key=f"f_confirm_{t}")
                valeurs_fixed[t] = (taille, recup, exceptionnel, confirme)

            st.markdown("**EPI sur demande**")
            valeurs_demande = {}
            for t in ON_DEMAND_EPI:
                c1, c2, c3 = st.columns(3)
                taille = c1.text_input(f"Taille {t}", value=row_data.get(f"Taille {t}", ""), key=f"f_taille_{t}")
                recup = c2.text_input(f"Récup. {t} (AAAA-MM-JJ)", value=row_data.get(f"Date récupération {t}", ""), key=f"f_recup_{t}")
                confirme = c3.checkbox("Confirmé", value=str(row_data.get(f"Confirmé {t}", "")).strip().lower() == "oui", key=f"f_confirm_{t}")
                valeurs_demande[t] = (taille, recup, confirme)

            col_save, col_delete = st.columns(2)
            submitted = col_save.form_submit_button("💾 Enregistrer", type="primary")
            deleted = col_delete.form_submit_button("🗑️ Supprimer cet employé") if idx_edit is not None else False

        if submitted:
            new_row = {
                "NOMS ET PRENOMS": nom, "FONCTION": fonction, "Parcours scolaire": parcours,
                "DATE D'AMBAUCHE": embauche, "Département": dept,
            }
            for t, (taille, recup, exceptionnel, confirme) in valeurs_fixed.items():
                new_row[f"Taille {t}"] = taille
                new_row[f"Date récupération {t}"] = recup
                new_row[f"Cas exceptionnel {t}"] = "Oui" if exceptionnel else ""
                new_row[f"Confirmé {t}"] = "Oui" if confirme else ""
            for t, (taille, recup, confirme) in valeurs_demande.items():
                new_row[f"Taille {t}"] = taille
                new_row[f"Date récupération {t}"] = recup
                new_row[f"Confirmé {t}"] = "Oui" if confirme else ""

            df_courant = st.session_state.personnel.copy()
            if idx_edit is not None:
                for k, v in new_row.items():
                    df_courant.at[idx_edit, k] = v
            else:
                df_courant = pd.concat([df_courant, pd.DataFrame([new_row])], ignore_index=True)

            df_courant = auto_fill_epi_expiration(df_courant)
            nb_archives = log_epi_changes(st.session_state.personnel_baseline, df_courant)
            save_personnel(df_courant)
            st.session_state.personnel = df_courant
            st.session_state.personnel_baseline = df_courant.copy()
            st.success("Employé enregistré." + (f" {nb_archives} ancien(s) EPI archivé(s)." if nb_archives else ""))
            st.rerun()

        if deleted and idx_edit is not None:
            df_courant = st.session_state.personnel.drop(index=idx_edit).reset_index(drop=True)
            log_epi_changes(st.session_state.personnel_baseline, df_courant)
            save_personnel(df_courant)
            st.session_state.personnel = df_courant
            st.session_state.personnel_baseline = df_courant.copy()
            st.success("Employé supprimé.")
            st.rerun()

    with tab_couleur:
        if st.session_state.personnel.empty:
            st.info("Aucune donnée à afficher.")
        else:
            aff = st.session_state.personnel[["NOMS ET PRENOMS", "FONCTION"]].copy()
            for t in EPI_TYPES:
                aff[t] = st.session_state.personnel.apply(lambda r: epi_type_status(r, t) or "⚪", axis=1)
            styled = aff.style.map(style_status_cell, subset=EPI_TYPES)
            st.dataframe(styled, use_container_width=True)

# ---------- Équipements ----------

elif page == "Équipements":
    st.title("Équipements")

    tab_inv, tab_interv, tab_couleur = st.tabs(["📋 Inventaire", "📝 Interventions externes", "🎨 Vue colorée"])

    with tab_inv:
        edited = st.data_editor(st.session_state.equip, num_rows="dynamic", use_container_width=True, key="equip_editor")
        st.session_state.equip = edited
        if st.button("💾 Sauvegarder l'inventaire", type="primary"):
            save_equip(st.session_state.equip)
            st.success("Inventaire sauvegardé dans Google Sheets.")

        st.markdown("---")
        st.subheader("➕ Ajouter / ✏️ Modifier un équipement")
        ids_existants = ["-- Nouvel équipement --"] + [
            f"{i} — {row.get('Identifiant', '')} ({row.get('Nom', '')})"
            for i, row in st.session_state.equip.iterrows()
        ]
        choix = st.selectbox("Choisir un équipement", ids_existants)
        idx_edit = None
        if choix != "-- Nouvel équipement --":
            idx_edit = int(choix.split(" — ")[0])
            row_data = st.session_state.equip.loc[idx_edit]
        else:
            row_data = pd.Series({c: "" for c in EQUIP_COLS})

        with st.form("form_equip"):
            identifiant = st.text_input("Identifiant", value=row_data.get("Identifiant", ""))
            nom = st.text_input("Nom", value=row_data.get("Nom", ""))
            constructeur = st.text_input("Constructeur", value=row_data.get("Constructeur", ""))
            type_eq = st.text_input("Type", value=row_data.get("Type", ""))
            num_serie = st.text_input("Numéro de série", value=row_data.get("Numéro de série", ""))
            criticite = st.text_input("Criticité", value=row_data.get("Criticité", ""))
            emplacement = st.text_input("Emplacement", value=row_data.get("Emplacement", ""))
            date_achat = st.text_input("Date d'achat (AAAA-MM-JJ)", value=row_data.get("Date d'achat", ""))
            col_save, col_delete = st.columns(2)
            submitted = col_save.form_submit_button("💾 Enregistrer", type="primary")
            deleted = col_delete.form_submit_button("🗑️ Supprimer") if idx_edit is not None else False

        if submitted:
            new_row = {
                "Identifiant": identifiant, "Nom": nom, "Constructeur": constructeur, "Type": type_eq,
                "Numéro de série": num_serie, "Criticité": criticite, "Emplacement": emplacement,
                "Date d'achat": date_achat,
            }
            df_courant = st.session_state.equip.copy()
            if idx_edit is not None:
                for k, v in new_row.items():
                    df_courant.at[idx_edit, k] = v
            else:
                df_courant = pd.concat([df_courant, pd.DataFrame([new_row])], ignore_index=True)
            save_equip(df_courant)
            st.session_state.equip = df_courant
            st.success("Équipement enregistré.")
            st.rerun()

        if deleted and idx_edit is not None:
            df_courant = st.session_state.equip.drop(index=idx_edit).reset_index(drop=True)
            save_equip(df_courant)
            st.session_state.equip = df_courant
            st.success("Équipement supprimé.")
            st.rerun()

    with tab_interv:
        st.caption(
            "Chaque étalonnage externe, maintenance ou vérification est enregistré ici comme une "
            "nouvelle ligne — rien n'est jamais écrasé, c'est ton archive complète."
        )
        if st.session_state.equip.empty:
            st.info("Ajoute d'abord des équipements dans l'onglet Inventaire.")
        else:
            with st.form("form_intervention"):
                options_eq = [f"{row.get('Identifiant', '')} — {row.get('Nom', '')}" for _, row in st.session_state.equip.iterrows()]
                choix_eq = st.selectbox("Équipement concerné", options_eq)
                nature = st.selectbox("Nature d'intervention", NATURE_OPTIONS)
                freq = st.selectbox("Fréquence prévue", FREQ_EXT_OPTIONS)
                date_interv = st.text_input("Date de l'intervention (AAAA-MM-JJ)", value=datetime.now().strftime("%Y-%m-%d"))
                resultat = st.selectbox("Résultat", RESULTAT_OPTIONS)
                envoye = st.form_submit_button("📌 Enregistrer cette intervention", type="primary")

            if envoye:
                identifiant_sel = choix_eq.split(" — ")[0]
                nom_sel = choix_eq.split(" — ", 1)[1] if " — " in choix_eq else ""
                ok = append_journal_row(
                    "Interventions_Externes",
                    [identifiant_sel, nom_sel, nature, freq, date_interv, resultat],
                )
                if ok:
                    st.success("Intervention enregistrée dans l'archive.")
                else:
                    st.error("Crée d'abord l'onglet « Interventions_Externes » dans Google Sheet.")

            st.subheader("Journal des interventions")
            interventions = load_journal("Interventions_Externes", INTERVENTION_COLS)
            if interventions is None:
                st.info("Onglet « Interventions_Externes » introuvable dans Google Sheet.")
            elif interventions.empty:
                st.info("Aucune intervention enregistrée pour le moment.")
            else:
                st.dataframe(interventions.sort_values("Date de l'intervention", ascending=False), use_container_width=True)

    with tab_couleur:
        interventions = load_journal("Interventions_Externes", INTERVENTION_COLS)
        if interventions is None or interventions.empty:
            st.info("Aucune donnée d'intervention à afficher.")
        else:
            derniers = dernieres_interventions(interventions)
            derniers["Statut"] = derniers.apply(intervention_status, axis=1)
            aff = derniers[["Identifiant", "Nom / Type", "Nature d'intervention", "Statut"]]
            styled = aff.style.map(style_status_cell, subset=["Statut"])
            st.dataframe(styled, use_container_width=True)

# ---------- Contrôle interne balances ----------

elif page == "Contrôle interne balances":
    st.title("Contrôle interne des balances")
    st.caption(
        "Saisie quotidienne via la carte de contrôle. Un résumé mensuel et annuel est calculé "
        "automatiquement à partir de ces entrées — rien n'est jamais écrasé, c'est ton archive."
    )

    balances_eq = st.session_state.equip[
        st.session_state.equip["Type"].astype(str).str.contains("balance", case=False, na=False)
    ]
    if balances_eq.empty:
        st.warning(
            "Aucun équipement avec « Balance » dans son Type. Ajoute-le dans Équipements → Inventaire "
            "(mets « Balance » dans la colonne Type), ou choisis directement dans la liste ci-dessous."
        )
        balances_eq = st.session_state.equip

    if balances_eq.empty:
        st.info("Ajoute d'abord des équipements dans Équipements → Inventaire.")
    else:
        with st.form("form_balance"):
            options_bal = [f"{row.get('Identifiant', '')} — {row.get('Nom', '')}" for _, row in balances_eq.iterrows()]
            choix_bal = st.selectbox("Balance concernée", options_bal)
            date_ctrl = st.text_input("Date (AAAA-MM-JJ)", value=datetime.now().strftime("%Y-%m-%d"))
            resultat = st.selectbox("Résultat du jour", RESULTAT_OPTIONS)
            remarque = st.text_input("Remarque (optionnel)")
            envoye = st.form_submit_button("📌 Enregistrer le contrôle du jour", type="primary")

        if envoye:
            identifiant_sel = choix_bal.split(" — ")[0]
            nom_sel = choix_bal.split(" — ", 1)[1] if " — " in choix_bal else ""
            ok = append_journal_row(
                "Controle_Interne_Balances",
                [identifiant_sel, nom_sel, date_ctrl, resultat, remarque],
            )
            if ok:
                st.success("Contrôle du jour enregistré.")
            else:
                st.error("Crée d'abord l'onglet « Controle_Interne_Balances » dans Google Sheet.")

        balances_log = load_journal("Controle_Interne_Balances", BALANCE_COLS)
        if balances_log is None:
            st.info("Onglet « Controle_Interne_Balances » introuvable dans Google Sheet.")
        elif balances_log.empty:
            st.info("Aucun contrôle enregistré pour le moment.")
        else:
            balances_log = balances_log.copy()
            balances_log["_date"] = pd.to_datetime(balances_log["Date"], errors="coerce")
            balances_log["Année"] = balances_log["_date"].dt.year
            balances_log["Mois"] = balances_log["_date"].dt.to_period("M").astype(str)

            st.subheader("Historique quotidien")
            st.dataframe(balances_log.drop(columns=["_date"]).sort_values("Date", ascending=False), use_container_width=True)

            st.subheader("📆 Résumé mensuel")
            resume_mois = balances_log.groupby(["Identifiant", "Mois"]).agg(
                Nb_controles=("Résultat", "count"),
                Non_conformites=("Résultat", lambda x: (x == "Non conforme").sum()),
            ).reset_index()
            resume_mois["Statut du mois"] = resume_mois["Non_conformites"].apply(
                lambda n: "🔴 Non conforme" if n > 0 else "🟢 Conforme"
            )
            st.dataframe(resume_mois, use_container_width=True)

            st.subheader("📅 Résumé annuel")
            resume_annee = balances_log.groupby(["Identifiant", "Année"]).agg(
                Nb_controles=("Résultat", "count"),
                Non_conformites=("Résultat", lambda x: (x == "Non conforme").sum()),
            ).reset_index()
            resume_annee["Statut de l'année"] = resume_annee["Non_conformites"].apply(
                lambda n: "🔴 Non conforme" if n > 0 else "🟢 Conforme"
            )
            st.dataframe(resume_annee, use_container_width=True)

# ---------- Import depuis l'Excel existant ----------

elif page == "Importer un Excel":
    st.title("Importer un fichier Excel existant")
    st.caption(
        "Charge le fichier que RQ t'a donné. Si le fichier a un en-tête de document au-dessus "
        "du vrai tableau, indique la ligne où commencent les vrais titres de colonnes."
    )
    cible = st.radio("Ce fichier concerne :", ["Personnel", "Équipements"], horizontal=True)
    attendu = PERSONNEL_COLS if cible == "Personnel" else EQUIP_COLS
    with st.expander("Voir les colonnes attendues par l'application"):
        st.write(", ".join(attendu))

    uploaded = st.file_uploader("Fichier Excel (.xlsx)", type=["xlsx"])
    if uploaded is not None:
        header_row = st.number_input(
            "Numéro de la ligne contenant les vrais titres de colonnes", min_value=1, max_value=30, value=1, step=1,
        )
        df_new = pd.read_excel(uploaded, header=header_row - 1)
        df_new = df_new.dropna(axis=1, how="all")
        df_new = df_new.dropna(how="all")
        noms = [str(c).strip() for c in df_new.columns]
        vus, noms_uniques = {}, []
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
        options = ["-- Ignorer --"] + attendu
        mapping = {}
        for col in df_new.columns:
            default_idx = options.index(col) if col in options else 0
            choix = st.selectbox(f"« {col} » →", options, index=default_idx, key=f"map_{cible}_{col}")
            if choix != "-- Ignorer --":
                mapping[col] = choix

        colonnes_manquantes = [c for c in attendu if c not in mapping.values()]
        if colonnes_manquantes:
            with st.expander(f"Colonnes non renseignées ({len(colonnes_manquantes)})"):
                st.write(", ".join(colonnes_manquantes))

        valeurs = list(mapping.values())
        doublons = sorted({v for v in valeurs if valeurs.count(v) > 1})
        if doublons:
            st.error(f"⚠️ Choisies plusieurs fois : {', '.join(doublons)}. Corrige avant de continuer.")

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
                    df_clean = auto_fill_epi_expiration(df_clean)
                    st.session_state.personnel = pd.concat([st.session_state.personnel, df_clean], ignore_index=True)
                    save_personnel(st.session_state.personnel)
                    st.session_state.personnel_baseline = st.session_state.personnel.copy()
                else:
                    st.session_state.equip = pd.concat([st.session_state.equip, df_clean], ignore_index=True)
                    save_equip(st.session_state.equip)
                st.success(f"{len(df_clean)} ligne(s) ajoutée(s) et sauvegardée(s) dans Google Sheets.")

st.sidebar.markdown("---")
st.sidebar.caption(f"Session ouverte : {datetime.now().strftime('%d/%m/%Y %H:%M')}")
