"""
Registre Labo — Personnel & Équipements
Application Streamlit simple, basée sur un fichier Excel comme "base de données".
Lancer avec : streamlit run app.py
"""
import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime

st.set_page_config(page_title="Registre Labo", page_icon="🧪", layout="wide")

DATA_FILE = Path("data/registre_labo.xlsx")
DATA_FILE.parent.mkdir(parents=True, exist_ok=True)

PERSONNEL_COLS = [
    "Nom complet", "Rôle", "Parcours scolaire", "Date de recrutement",
    "Département", "Type EPI", "Taille EPI",
    "Date attribution EPI", "Date renouvellement EPI",
]
EQUIP_COLS = [
    "Identifiant", "Nom / type", "Emplacement", "Date d'achat",
    "Fréquence étalonnage (jours)", "Dernier étalonnage", "Dernière maintenance",
]

# ---------- Chargement / sauvegarde ----------

def load_data():
    if DATA_FILE.exists():
        xls = pd.ExcelFile(DATA_FILE)
        personnel = (
            pd.read_excel(xls, "Personnel") if "Personnel" in xls.sheet_names
            else pd.DataFrame(columns=PERSONNEL_COLS)
        )
        equip = (
            pd.read_excel(xls, "Equipements") if "Equipements" in xls.sheet_names
            else pd.DataFrame(columns=EQUIP_COLS)
        )
    else:
        personnel = pd.DataFrame(columns=PERSONNEL_COLS)
        equip = pd.DataFrame(columns=EQUIP_COLS)
    return personnel, equip


def save_data(personnel_df, equip_df):
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as writer:
        personnel_df.to_excel(writer, sheet_name="Personnel", index=False)
        equip_df.to_excel(writer, sheet_name="Equipements", index=False)


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
    st.caption("Modifie directement le tableau comme dans Excel. Ajoute une ligne en bas, coche pour supprimer.")
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
        save_data(st.session_state.personnel, st.session_state.equip)
        st.success("Données sauvegardées dans data/registre_labo.xlsx")

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
        save_data(st.session_state.personnel, st.session_state.equip)
        st.success("Données sauvegardées dans data/registre_labo.xlsx")

# ---------- Import depuis l'Excel existant ----------

elif page == "Importer un Excel":
    st.title("Importer un fichier Excel existant")
    st.caption(
        "Charge le fichier que RQ t'a donné. Si les noms de colonnes correspondent à ceux "
        "attendus ci-dessous, les lignes seront ajoutées automatiquement."
    )
    cible = st.radio("Ce fichier concerne :", ["Personnel", "Équipements"], horizontal=True)
    attendu = PERSONNEL_COLS if cible == "Personnel" else EQUIP_COLS
    st.write("Colonnes attendues :", ", ".join(attendu))

    uploaded = st.file_uploader("Fichier Excel (.xlsx)", type=["xlsx"])
    if uploaded is not None:
        df_new = pd.read_excel(uploaded)
        st.write("Aperçu du fichier importé :")
        st.dataframe(df_new.head(), use_container_width=True)

        colonnes_ok = [c for c in df_new.columns if c in attendu]
        colonnes_inconnues = [c for c in df_new.columns if c not in attendu]
        if colonnes_inconnues:
            st.warning(
                f"Colonnes non reconnues (ignorées) : {', '.join(colonnes_inconnues)}. "
                "Renomme-les dans ton Excel pour qu'elles correspondent, puis réimporte."
            )

        if st.button("Ajouter ces lignes"):
            df_clean = df_new[colonnes_ok]
            if cible == "Personnel":
                st.session_state.personnel = pd.concat(
                    [st.session_state.personnel, df_clean], ignore_index=True
                )
            else:
                st.session_state.equip = pd.concat(
                    [st.session_state.equip, df_clean], ignore_index=True
                )
            save_data(st.session_state.personnel, st.session_state.equip)
            st.success(f"{len(df_clean)} ligne(s) ajoutée(s) et sauvegardée(s).")

st.sidebar.markdown("---")
st.sidebar.caption(f"Session ouverte : {datetime.now().strftime('%d/%m/%Y %H:%M')}")
