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
    "Nom complet", "Rôle", "Parcours scolaire", "Date de recrutement",
    "Département", "Type EPI", "Taille EPI",
    "Date attribution EPI", "Date renouvellement EPI",
]
EQUIP_COLS = [
    "Identifiant", "Nom / type", "Emplacement", "Date d'achat",
    "Fréquence étalonnage (jours)", "Dernier étalonnage", "Dernière maintenance",
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
    ws.update([df.columns.tolist()] + df.astype(str).values.tolist())


def save_equip(df):
    ws = get_sheet().worksheet("Equipements")
    ws.clear()
    ws.update([df.columns.tolist()] + df.astype(str).values.tolist())


if "personnel" not in st.session_state or "equip" not in st.session_state:
    p, e = load_data()
    st.session_state.personnel = p
    st.session_state.equip = e

# ---------- Statuts calculés ----------

def calib_status(row):
    try:
        last = pd.to_datetime(row["Dernier étalonnage"])
        freq = float(row["Fréquence étalonnage (jours)"])
    except Exception:
        return "⚪ Non renseigné"
    due = last + pd.Timedelta(days=freq)
    days_left = (due - pd.Timestamp.now()).days
    if days_left < 0:
        return f"🔴 En retard ({abs(days_left)} j)"
    elif days_left <= 1:
        return "🟠 À faire aujourd'hui/demain"
    return f"🟢 OK ({days_left} j restants)"


def epi_status(row):
    try:
        expiry = pd.to_datetime(row["Date renouvellement EPI"])
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
        st.subheader("Alertes étalonnage")
        eq = st.session_state.equip.copy()
        eq["Statut"] = eq.apply(calib_status, axis=1)
        alertes = eq[eq["Statut"].str.contains("🔴|🟠", na=False)]
        if alertes.empty:
            st.success("Aucune alerte d'étalonnage en cours.")
        else:
            st.dataframe(alertes[["Identifiant", "Nom / type", "Statut"]], use_container_width=True)

    if not st.session_state.personnel.empty:
        st.subheader("Alertes EPI")
        pe = st.session_state.personnel.copy()
        pe["Statut EPI"] = pe.apply(epi_status, axis=1)
        alertes_epi = pe[pe["Statut EPI"].str.contains("🔴|🟠", na=False)]
        if alertes_epi.empty:
            st.success("Aucune alerte EPI en cours.")
        else:
            st.dataframe(alertes_epi[["Nom complet", "Type EPI", "Statut EPI"]], use_container_width=True)

# ---------- Personnel ----------

elif page == "Personnel":
    st.title("Personnel")
    st.caption("Modifie directement le tableau comme dans Excel, puis clique sur Sauvegarder.")
    edited = st.data_editor(
        st.session_state.personnel,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Date de recrutement": st.column_config.DateColumn(),
            "Date attribution EPI": st.column_config.DateColumn(),
            "Date renouvellement EPI": st.column_config.DateColumn(),
        },
        key="personnel_editor",
    )
    st.session_state.personnel = edited
    if st.button("💾 Sauvegarder Personnel", type="primary"):
        save_personnel(st.session_state.personnel)
        st.success("Données sauvegardées dans Google Sheets.")

# ---------- Équipements ----------

elif page == "Équipements":
    st.title("Équipements")
    st.caption("Fréquence d'étalonnage en jours (1 = quotidien, 3-4 = 2x/semaine, 30 = mensuel...).")
    edited = st.data_editor(
        st.session_state.equip,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Date d'achat": st.column_config.DateColumn(),
            "Dernier étalonnage": st.column_config.DateColumn(),
            "Dernière maintenance": st.column_config.DateColumn(),
            "Fréquence étalonnage (jours)": st.column_config.NumberColumn(min_value=1),
        },
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
        df_new.columns = [str(c).strip() for c in df_new.columns]

        st.write("Aperçu avec cette ligne d'en-tête :")
        st.dataframe(df_new.head(10), use_container_width=True)

        st.subheader("Faire correspondre les colonnes")
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

        if st.button("Ajouter ces lignes"):
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
