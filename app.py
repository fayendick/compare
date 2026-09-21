# -*- coding: utf-8 -*-
"""
Interface Streamlit — Comparaison des soldes
Report Finance  vs  Balance Détaillée

Ce fichier ne fait QUE de l'affichage : tous les calculs sont dans comparaison_soldes.py

Lancement :
    pip install streamlit pandas openpyxl
    streamlit run app.py
"""

import io

import pandas as pd
import streamlit as st

from comparaison_soldes import (
    ORIGINE_EXCLU,
    ColonnesManquantes,
    comparer,
    detecter_colonnes,
    diagnostiquer_compte,
    fmt,
    preparer_balance,
    preparer_report_finance,
    read_file,
)

st.set_page_config(page_title="Comparaison des soldes", layout="wide", page_icon="🔍")

st.title("🔍 Comparaison des soldes — Report Finance vs Balance Détaillée")

st.markdown(
    """
Cette application liste **uniquement les comptes qui expliquent l'écart** entre le
*Report Finance* et la *Balance Détaillée*, pour le solde **début de mois** et le solde **fin de mois**.

- Seules les lignes du *Report Finance* **sans MATRICULE_CLIENT** entrent dans la comparaison.
- Les lignes **avec MATRICULE_CLIENT** ne sont pas nécessaires et n'apparaissent pas dans le rapport.
- Un compte présent dans un seul fichier est compté comme **0** de l'autre côté.
"""
)

col_up1, col_up2 = st.columns(2)
with col_up1:
    file1 = st.file_uploader("📄 Report Finance", type=["xlsx", "xls", "csv"], key="f1")
with col_up2:
    file2 = st.file_uploader("📄 Balance Détaillée", type=["xlsx", "xls", "csv"], key="f2")

tolerance = st.sidebar.number_input(
    "Tolérance (écart ignoré en dessous de cette valeur absolue)",
    min_value=0.0, value=1.0, step=0.5,
)
afficher_exclus = st.sidebar.checkbox(
    "Inclure les comptes avec matricule dans les tableaux", value=False,
    help="Ces comptes existent dans le Report Finance mais ne sont pas nécessaires au rapport. "
         "Ils pèsent tout de même sur l'écart total : décoche pour les sortir des tableaux, "
         "leur montant reste affiché à part.",
)

st.sidebar.markdown("---")
st.sidebar.subheader("🔥 Top N des plus gros écarts")
top_n_choix = st.sidebar.selectbox(
    "Nombre de comptes à afficher",
    options=["10", "20", "50", "Personnalisé", "Tout"],
    index=1,
    help="Limite l'affichage aux N comptes ayant le plus gros écart. "
         "Choisis « Tout » pour ne rien filtrer.",
)
if top_n_choix == "Personnalisé":
    top_n = st.sidebar.number_input("Valeur de N", min_value=1, value=15, step=1)
elif top_n_choix == "Tout":
    top_n = None
else:
    top_n = int(top_n_choix)

if not file1 or not file2:
    st.info("⬆️ Merci d'uploader les deux fichiers pour lancer la comparaison.")
    st.stop()

# ---------------------------------------------------------------- lecture / préparation
try:
    df1_raw = read_file(file1)
    df2_raw = read_file(file2)
except Exception as e:
    st.error(f"Erreur de lecture des fichiers : {e}")
    st.stop()

try:
    cols = detecter_colonnes(df1_raw, df2_raw)
except ColonnesManquantes as e:
    st.error(f"❌ Colonnes manquantes dans {e.fichier} : {e.colonnes}")
    st.write("Colonnes détectées :", e.disponibles)
    st.stop()

with st.expander("🔎 Colonnes détectées"):
    st.write(
        f"**Report Finance** → ACCOUNT_NO=`{cols.account1}` | MATRICULE_CLIENT=`{cols.matricule}` "
        f"| Ouverture=`{cols.ouverture}` | Final=`{cols.final1}`"
    )
    st.write(
        f"**Balance Détaillée** → Numéro de compte=`{cols.account2}` "
        f"| Début=`{cols.debut2}` | Fin=`{cols.fin2}`"
    )

r1, r1_exclus, nb_ignorees, dup1 = preparer_report_finance(df1_raw, cols)
r2, dup2 = preparer_balance(df2_raw, cols)
res = comparer(r1, r2, tolerance, nb_ignorees, (dup1, dup2), r1_exclus)

if dup1 or dup2:
    st.warning(
        f"⚠️ Doublons de comptes détectés (Report Finance : {dup1}, Balance Détaillée : {dup2}). "
        "Les montants ont été additionnés par compte."
    )

st.success(
    f"✅ Report Finance : {res.nb_comptes_f1} comptes retenus (sans matricule client) "
    f"— {res.nb_lignes_ignorees} lignes avec matricule non nécessaires au rapport. "
    f"Balance Détaillée : {res.nb_comptes_f2} comptes."
)

# ------------------------------------------------------------------------- indicateurs
st.header("📊 Écarts globaux")

c1, c2, c3 = st.columns(3)
c1.metric("Total début — Report Finance", fmt(res.total_debut_f1))
c2.metric("Total début — Balance Détaillée", fmt(res.total_debut_f2))
c3.metric("Écart début de mois", fmt(res.ecart_total_debut))

c4, c5, c6 = st.columns(3)
c4.metric("Total fin — Report Finance", fmt(res.total_fin_f1))
c5.metric("Total fin — Balance Détaillée", fmt(res.total_fin_f2))
c6.metric("Écart fin de mois", fmt(res.ecart_total_fin))

# ------------------------------------------------------------- comptes à l'origine
st.header("🚨 Comptes à l'origine de l'écart")

RENOM = {
    "ACCOUNT_NO": "Compte",
    "ACCOUNT_KEY": "Clé de rapprochement",
    "SOLDE_F1": "Report Finance",
    "SOLDE_F2": "Balance Détaillée",
    "ECART": "Écart absolu",
    "ORIGINE": "Origine",
}


def _to_excel_bytes(df) -> bytes:
    """Convertit un DataFrame en bytes .xlsx (une seule feuille)."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Ecarts")
    return buffer.getvalue()


def _afficher(df, ecart_total, libelle, top_n):
    exclus = df[df["ORIGINE"] == ORIGINE_EXCLU]
    visible_complet = df if afficher_exclus else df[df["ORIGINE"] != ORIGINE_EXCLU]

    part_exclus = float(exclus["ECART"].sum()) if len(exclus) else 0.0

    st.caption(f"{len(visible_complet)} compte(s) au total.")
    if len(exclus) and not afficher_exclus:
        st.caption(
            f"➕ {len(exclus)} compte(s) avec matricule client, non nécessaires au rapport, "
            f"pèsent **{fmt(part_exclus)}** dans l'écart total."
        )

    # ---- Application du Top N (déjà trié par écart décroissant en amont) ----
    if top_n is not None and len(visible_complet) > top_n:
        visible = visible_complet.head(top_n)
        st.info(
            f"🔥 Top {top_n} affiché ({len(visible)} compte(s) sur {len(visible_complet)})."
        )
    else:
        visible = visible_complet

    if visible.empty:
        st.success(f"Aucun écart affiché sur le {libelle} au-delà de la tolérance.")
    else:
        # Graphique des écarts (comptes les plus impactants, sur la sélection affichée)
        chart_df = visible[["ACCOUNT_NO", "ECART"]].set_index("ACCOUNT_NO")
        st.bar_chart(chart_df, y="ECART", height=320)

        df_renomme = visible.rename(columns=RENOM)
        st.dataframe(df_renomme, width='stretch', hide_index=True)

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                f"⬇️ CSV — {libelle}",
                data=df_renomme.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"ecarts_{libelle.replace(' ', '_')}.csv",
                mime="text/csv",
                key=f"dl_csv_{libelle}",
                width='stretch',
            )
        with col_dl2:
            st.download_button(
                f"⬇️ Excel — {libelle}",
                data=_to_excel_bytes(df_renomme),
                file_name=f"ecarts_{libelle.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_xlsx_{libelle}",
                width='stretch',
            )

        if top_n is not None and len(visible_complet) > top_n:
            with st.expander(f"📋 Voir les {len(visible_complet)} comptes (liste complète) — {libelle}"):
                df_complet_renomme = visible_complet.rename(columns=RENOM)
                st.dataframe(df_complet_renomme, width='stretch', hide_index=True)

                col_dlc1, col_dlc2 = st.columns(2)
                with col_dlc1:
                    st.download_button(
                        f"⬇️ CSV — {libelle} (complet)",
                        data=df_complet_renomme.to_csv(index=False).encode("utf-8-sig"),
                        file_name=f"ecarts_{libelle.replace(' ', '_')}_complet.csv",
                        mime="text/csv",
                        key=f"dl_csv_complet_{libelle}",
                        width='stretch',
                    )
                with col_dlc2:
                    st.download_button(
                        f"⬇️ Excel — {libelle} (complet)",
                        data=_to_excel_bytes(df_complet_renomme),
                        file_name=f"ecarts_{libelle.replace(' ', '_')}_complet.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_xlsx_complet_{libelle}",
                        width='stretch',
                    )

    if len(exclus):
        with st.expander(f"🚫 Comptes avec matricule client (non nécessaires) — {libelle} ({len(exclus)})"):
            exclus_renomme = exclus.rename(columns=RENOM)
            st.dataframe(exclus_renomme, width='stretch', hide_index=True)

            col_dle1, col_dle2 = st.columns(2)
            with col_dle1:
                st.download_button(
                    f"⬇️ CSV — exclus {libelle}",
                    data=exclus_renomme.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"exclus_{libelle.replace(' ', '_')}.csv",
                    mime="text/csv",
                    key=f"dl_csv_exclus_{libelle}",
                    width='stretch',
                )
            with col_dle2:
                st.download_button(
                    f"⬇️ Excel — exclus {libelle}",
                    data=_to_excel_bytes(exclus_renomme),
                    file_name=f"exclus_{libelle.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_xlsx_exclus_{libelle}",
                    width='stretch',
                )


tab1, tab2 = st.tabs(["Solde début de mois", "Solde fin de mois"])

with tab1:
    _afficher(res.ecarts_debut, res.ecart_total_debut, "solde debut", top_n)

with tab2:
    _afficher(res.ecarts_fin, res.ecart_total_fin, "solde fin", top_n)

# ------------------------------------------------------------------ diagnostic compte
st.header("🩺 Vérifier un compte")

numero = st.text_input("Numéro de compte à diagnostiquer", placeholder="ex. 203110000001")
if numero:
    diag = diagnostiquer_compte(numero, df1_raw, df2_raw, cols)
    st.write(
        f"Clé de rapprochement : `{diag['cle_normalisee']}` — **{diag['origine']}**"
    )
    st.write(
        f"Report Finance : {diag['nb_lignes_report_finance']} ligne(s) trouvée(s), "
        f"dont {diag['nb_lignes_retenues']} retenue(s) (sans matricule client). "
        f"Balance Détaillée : {diag['nb_lignes_balance']} ligne(s)."
    )
    if diag["nb_lignes_report_finance"] and not diag["nb_lignes_retenues"]:
        st.warning("Ce compte existe bien dans le Report Finance mais toutes ses lignes ont un matricule client, donc non nécessaires au rapport.")
    if diag["nb_lignes_report_finance"]:
        st.write("**Lignes Report Finance**")
        st.dataframe(diag["lignes_report_finance"], width='stretch')
    if diag["nb_lignes_balance"]:
        st.write("**Lignes Balance Détaillée**")
        st.dataframe(diag["lignes_balance"], width='stretch')