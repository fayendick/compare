# -*- coding: utf-8 -*-
"""
Logique métier — Comparaison des soldes
Report Finance  vs  Balance Détaillée

Aucun import Streamlit ici : ce module est testable et réutilisable seul.

Règles :
  - Les lignes du Report Finance SANS MATRICULE_CLIENT sont exclues AVANT tout calcul
    (elles n'apparaissent donc jamais dans les écarts).
  - On ne sort que les comptes qui expliquent l'écart entre les deux fichiers :
      * présents des deux côtés avec solde différent
      * présents uniquement dans un fichier (l'autre côté vaut alors 0)
  - Un écart est calculé séparément pour le solde début de mois et le solde fin de mois.
  - Les deux fichiers peuvent utiliser des conventions de signe différentes
    (ex. débiteur/créditeur inversés) : tous les écarts sont donc calculés
    comme la différence des VALEURS ABSOLUES des soldes, et non la différence
    des soldes signés (sinon deux montants opposés en signe mais égaux en
    valeur s'additionnent au lieu de s'annuler).
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation

import numpy as np
import pandas as pd

# --------------------------------------------------------------------------------------
# Utilitaires
# --------------------------------------------------------------------------------------


def normalize(s: str) -> str:
    """Enlève les accents, met en minuscule et nettoie les espaces."""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.replace("_", " ")
    return re.sub(r"\s+", " ", s).strip().lower()


def read_file(uploaded_file) -> pd.DataFrame:
    """Lit un CSV ou un Excel (objet fichier Streamlit ou chemin)."""
    name = getattr(uploaded_file, "name", str(uploaded_file)).lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file, sep=None, engine="python")
    return pd.read_excel(uploaded_file)


def clean_numeric(series: pd.Series) -> pd.Series:
    """Convertit une colonne de soldes (texte ou nombre, format FR ou EN) en float."""

    def conv(v):
        if pd.isna(v):
            return np.nan
        if isinstance(v, (int, float, np.integer, np.floating)):
            return float(v)
        s = str(v).strip().replace("\xa0", "").replace(" ", "")
        if s == "":
            return np.nan
        neg = False
        if s.startswith("(") and s.endswith(")"):
            neg = True
            s = s[1:-1]
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")   # 1.234,56 -> 1234.56
        elif "," in s:
            s = s.replace(",", ".")
        s = re.sub(r"[^0-9\.\-]", "", s)
        if s in ("", "-", "."):
            return np.nan
        try:
            val = float(s)
            return -val if neg else val
        except ValueError:
            return np.nan

    return series.apply(conv)


def clean_account_key(series: pd.Series) -> pd.Series:
    """Normalise les numéros de compte pour permettre le rapprochement entre les 2 fichiers."""
    return series.apply(_key_one)


def _key_one(v) -> str:
    if pd.isna(v):
        return ""
    s = str(v).strip()
    if s.lower() in ("nan", "none", "null", ""):
        return ""
    # Excel transforme parfois un long numéro en flottant : 2.0311e+11 -> 203110000001
    if re.fullmatch(r"[+-]?\d+(\.\d+)?[eE][+-]?\d+", s):
        try:
            s = f"{Decimal(s):f}"
        except (InvalidOperation, ValueError):
            pass
    s = re.sub(r"\.0+$", "", s)                  # 12345.0 -> 12345
    s = re.sub(r"[^0-9A-Za-z]", "", s).upper()   # enlève espaces, tirets, points...
    s = s.lstrip("0")
    return s or "0"


def fmt(x) -> str:
    """Formatage FR : 1 234 567,89"""
    try:
        return f"{float(x):,.2f}".replace(",", " ").replace(".", ",")
    except (TypeError, ValueError):
        return str(x)


# --------------------------------------------------------------------------------------
# Détection des colonnes
# --------------------------------------------------------------------------------------


class ColonnesManquantes(Exception):
    def __init__(self, fichier: str, colonnes: list, disponibles: list):
        self.fichier = fichier
        self.colonnes = colonnes
        self.disponibles = disponibles
        super().__init__(f"Colonnes manquantes dans {fichier} : {colonnes}")


@dataclass
class Colonnes:
    account1: str
    matricule: str
    ouverture: str
    final1: str
    account2: str
    debut2: str
    fin2: str


def _exact(colmap: dict, target: str):
    return colmap.get(normalize(target))


def _startswith(colmap: dict, prefix: str):
    p = normalize(prefix)
    for norm, orig in colmap.items():
        if norm.startswith(p):
            return orig
    return None


def detecter_colonnes(df1: pd.DataFrame, df2: pd.DataFrame) -> Colonnes:
    colmap1 = {normalize(c): c for c in df1.columns}
    colmap2 = {normalize(c): c for c in df2.columns}

    account1 = _exact(colmap1, "ACCOUNT_NO")
    matricule = _exact(colmap1, "MATRICULE_CLIENT")
    ouverture = _startswith(colmap1, "SOLDE_OUVERTURE")
    final1 = _startswith(colmap1, "SOLDE_FINAL")

    account2 = _exact(colmap2, "Numéro de compte")
    debut2 = _exact(colmap2, "Solde debut de mois")
    fin2 = _exact(colmap2, "Solde fin de mois")

    manquantes1 = [n for n, v in [
        ("ACCOUNT_NO", account1),
        ("MATRICULE_CLIENT", matricule),
        ("SOLDE_OUVERTURE <date>", ouverture),
        ("SOLDE_FINAL <date>", final1),
    ] if v is None]
    if manquantes1:
        raise ColonnesManquantes("Report Finance", manquantes1, list(df1.columns))

    manquantes2 = [n for n, v in [
        ("Numéro de compte", account2),
        ("Solde debut de mois", debut2),
        ("Solde fin de mois", fin2),
    ] if v is None]
    if manquantes2:
        raise ColonnesManquantes("Balance Détaillée", manquantes2, list(df2.columns))

    return Colonnes(account1, matricule, ouverture, final1, account2, debut2, fin2)


# --------------------------------------------------------------------------------------
# Préparation des données
# --------------------------------------------------------------------------------------


def _mask_matricule(df1_raw: pd.DataFrame, cols: Colonnes) -> pd.Series:
    """True quand la ligne porte un matricule client exploitable."""
    mat = df1_raw[cols.matricule].astype(str).str.strip().str.lower()
    return df1_raw[cols.matricule].notna() & ~mat.isin(["", "nan", "none", "null", "0"])


def _mise_en_forme_f1(df: pd.DataFrame, cols: Colonnes) -> pd.DataFrame:
    r = df[[cols.account1, cols.ouverture, cols.final1]].copy()
    r.columns = ["ACCOUNT_NO", "SOLDE_DEBUT_F1", "SOLDE_FIN_F1"]
    r["SOLDE_DEBUT_F1"] = clean_numeric(r["SOLDE_DEBUT_F1"])
    r["SOLDE_FIN_F1"] = clean_numeric(r["SOLDE_FIN_F1"])
    r["ACCOUNT_KEY"] = clean_account_key(r["ACCOUNT_NO"])
    return r.groupby("ACCOUNT_KEY", as_index=False).agg(
        ACCOUNT_NO=("ACCOUNT_NO", "first"),
        SOLDE_DEBUT_F1=("SOLDE_DEBUT_F1", "sum"),
        SOLDE_FIN_F1=("SOLDE_FIN_F1", "sum"),
    )


def preparer_report_finance(df1_raw: pd.DataFrame, cols: Colonnes):
    """Sépare les lignes du Report Finance selon la présence d'un matricule client.

    Seules les lignes SANS matricule client entrent dans la comparaison : ce sont elles
    qui doivent être rapprochées de la Balance Détaillée. Les lignes AVEC matricule ne
    sont pas nécessaires et n'apparaissent jamais dans le rapport.

    Retourne (r1_retenus, r1_exclus, nb_lignes_ignorees, nb_doublons) où :
      - r1_retenus = comptes sans matricule (utilisés pour la comparaison)
      - r1_exclus  = comptes avec matricule (ignorés du rapport)

    r1_exclus sert uniquement à savoir qu'un compte EXISTE bien dans le Report Finance
    même s'il a été écarté : sans cette information, un compte avec matricule serait
    étiqueté à tort « Absent du Report Finance ».
    """
    avec_matricule = _mask_matricule(df1_raw, cols)
    df1 = df1_raw[~avec_matricule].copy()      # sans matricule -> retenu
    df1_exclu = df1_raw[avec_matricule].copy()  # avec matricule -> exclu
    nb_ignorees = len(df1_exclu)

    r1 = _mise_en_forme_f1(df1, cols)
    r1_exclus = _mise_en_forme_f1(df1_exclu, cols) if nb_ignorees else _mise_en_forme_f1(df1.iloc[0:0], cols)

    nb_doublons = int(len(df1) - len(r1))
    return r1, r1_exclus, nb_ignorees, nb_doublons


def preparer_balance(df2_raw: pd.DataFrame, cols: Colonnes):
    """Nettoie et regroupe la Balance Détaillée par compte. Retourne (DataFrame, nb_doublons)."""
    r2 = df2_raw[[cols.account2, cols.debut2, cols.fin2]].copy()
    r2.columns = ["NUM_COMPTE", "SOLDE_DEBUT_F2", "SOLDE_FIN_F2"]
    r2["SOLDE_DEBUT_F2"] = clean_numeric(r2["SOLDE_DEBUT_F2"])
    r2["SOLDE_FIN_F2"] = clean_numeric(r2["SOLDE_FIN_F2"])
    r2["ACCOUNT_KEY"] = clean_account_key(r2["NUM_COMPTE"])

    nb_doublons = int(r2["ACCOUNT_KEY"].duplicated().sum())
    r2 = r2.groupby("ACCOUNT_KEY", as_index=False).agg(
        NUM_COMPTE=("NUM_COMPTE", "first"),
        SOLDE_DEBUT_F2=("SOLDE_DEBUT_F2", "sum"),
        SOLDE_FIN_F2=("SOLDE_FIN_F2", "sum"),
    )
    return r2, nb_doublons


# --------------------------------------------------------------------------------------
# Comparaison
# --------------------------------------------------------------------------------------


@dataclass
class Resultat:
    ecarts_debut: pd.DataFrame          # uniquement les comptes qui causent l'écart (début)
    ecarts_fin: pd.DataFrame            # uniquement les comptes qui causent l'écart (fin)
    total_debut_f1: float = 0.0
    total_debut_f2: float = 0.0
    total_fin_f1: float = 0.0
    total_fin_f2: float = 0.0
    nb_comptes_f1: int = 0
    nb_comptes_f2: int = 0
    nb_lignes_ignorees: int = 0
    doublons: tuple = field(default=(0, 0))

    @property
    def ecart_total_debut(self) -> float:
        # Diff des valeurs ABSOLUES : évite qu'un signe inversé entre les deux
        # fichiers (débiteur/créditeur) ne fasse s'additionner deux montants
        # égaux au lieu de s'annuler.
        return abs(abs(self.total_debut_f1) - abs(self.total_debut_f2))

    @property
    def ecart_total_fin(self) -> float:
        return abs(abs(self.total_fin_f1) - abs(self.total_fin_f2))


ORIGINE_DEUX = "Présent dans les deux"
ORIGINE_ABS_F2 = "Absent de la Balance Détaillée"
ORIGINE_ABS_F1 = "Absent du Report Finance"
ORIGINE_EXCLU = "Exclu du Report Finance (avec matricule)"


def _construire_ecarts(merged: pd.DataFrame, col_f1: str, col_f2: str, tolerance: float) -> pd.DataFrame:
    """Ne garde que les comptes dont le solde diffère entre les deux fichiers.

    ECART est la différence des VALEURS ABSOLUES des soldes des deux fichiers
    (et non la différence des soldes signés), car les deux fichiers peuvent
    utiliser des conventions de signe opposées pour un même compte.
    Les colonnes SOLDE_F1 / SOLDE_F2 affichées restent les valeurs absolues
    utilisées pour le calcul, afin que la colonne ECART soit cohérente avec
    ce qui est montré.
    """
    d = merged.copy()
    d["SOLDE_F1"] = d[col_f1].fillna(0.0).abs()
    d["SOLDE_F2"] = d[col_f2].fillna(0.0).abs()
    d["ECART"] = (d["SOLDE_F1"] - d["SOLDE_F2"]).abs()

    # Un compte absent de r1 peut l'être pour deux raisons très différentes :
    # il n'existe pas du tout dans le Report Finance, ou il y existe mais a été
    # écarté faute de matricule client. On ne les confond plus.
    absent_f1 = d["_merge"] == "right_only"
    d["ORIGINE"] = np.select(
        [d["_merge"] == "left_only", absent_f1 & d["EXCLU_F1"], absent_f1],
        [ORIGINE_ABS_F2, ORIGINE_EXCLU, ORIGINE_ABS_F1],
        default=ORIGINE_DEUX,
    )
    d["ACCOUNT_NO"] = d["ACCOUNT_NO"].astype(str).str.strip().str.replace(r"\.0+$", "", regex=True)
    d = d[d["ECART"] > tolerance]
    d = d.sort_values("ECART", ascending=False)
    return d[["ACCOUNT_NO", "ACCOUNT_KEY", "SOLDE_F1", "SOLDE_F2", "ECART", "ORIGINE"]].reset_index(drop=True)


def comparer(r1: pd.DataFrame, r2: pd.DataFrame, tolerance: float = 1.0,
             nb_lignes_ignorees: int = 0, doublons: tuple = (0, 0),
             r1_exclus: pd.DataFrame | None = None) -> Resultat:
    merged = pd.merge(r1, r2, on="ACCOUNT_KEY", how="outer", indicator=True)
    merged["ACCOUNT_NO"] = merged["ACCOUNT_NO"].fillna(merged["NUM_COMPTE"])

    cles_exclues = set(r1_exclus["ACCOUNT_KEY"]) if r1_exclus is not None else set()
    merged["EXCLU_F1"] = merged["ACCOUNT_KEY"].isin(cles_exclues)

    ecarts_debut = _construire_ecarts(merged, "SOLDE_DEBUT_F1", "SOLDE_DEBUT_F2", tolerance)
    ecarts_fin = _construire_ecarts(merged, "SOLDE_FIN_F1", "SOLDE_FIN_F2", tolerance)

    return Resultat(
        ecarts_debut=ecarts_debut,
        ecarts_fin=ecarts_fin,
        total_debut_f1=float(r1["SOLDE_DEBUT_F1"].sum()),
        total_debut_f2=float(r2["SOLDE_DEBUT_F2"].sum()),
        total_fin_f1=float(r1["SOLDE_FIN_F1"].sum()),
        total_fin_f2=float(r2["SOLDE_FIN_F2"].sum()),
        nb_comptes_f1=len(r1),
        nb_comptes_f2=len(r2),
        nb_lignes_ignorees=nb_lignes_ignorees,
        doublons=doublons,
    )


def analyser(fichier_report, fichier_balance, tolerance: float = 1.0) -> Resultat:
    """Point d'entrée unique : deux fichiers en entrée, un Resultat en sortie."""
    df1_raw = read_file(fichier_report)
    df2_raw = read_file(fichier_balance)
    cols = detecter_colonnes(df1_raw, df2_raw)
    r1, r1_exclus, nb_ignorees, dup1 = preparer_report_finance(df1_raw, cols)
    r2, dup2 = preparer_balance(df2_raw, cols)
    return comparer(r1, r2, tolerance, nb_ignorees, (dup1, dup2), r1_exclus)


# --------------------------------------------------------------------------------------
# Diagnostic d'un compte précis
# --------------------------------------------------------------------------------------


def diagnostiquer_compte(numero, df1_raw: pd.DataFrame, df2_raw: pd.DataFrame, cols: Colonnes) -> dict:
    """Explique, pour un numéro de compte donné, où il est trouvé et pourquoi il est classé ainsi."""
    cle = _key_one(numero)

    k1 = clean_account_key(df1_raw[cols.account1])
    k2 = clean_account_key(df2_raw[cols.account2])
    sans_matricule = ~_mask_matricule(df1_raw, cols)  # ce sont ces lignes qui sont retenues

    lignes_f1 = df1_raw[k1 == cle]
    lignes_f1_retenues = df1_raw[(k1 == cle) & sans_matricule]
    lignes_f2 = df2_raw[k2 == cle]

    if len(lignes_f1) and not len(lignes_f1_retenues):
        origine = ORIGINE_EXCLU
    elif not len(lignes_f1):
        origine = ORIGINE_ABS_F1
    elif not len(lignes_f2):
        origine = ORIGINE_ABS_F2
    else:
        origine = ORIGINE_DEUX

    return {
        "cle_normalisee": cle,
        "origine": origine,
        "nb_lignes_report_finance": len(lignes_f1),
        "nb_lignes_retenues": len(lignes_f1_retenues),
        "nb_lignes_balance": len(lignes_f2),
        "lignes_report_finance": lignes_f1,
        "lignes_balance": lignes_f2,
        "cles_report_finance": sorted(set(df1_raw.loc[k1 == cle, cols.account1].astype(str))),
        "cles_balance": sorted(set(df2_raw.loc[k2 == cle, cols.account2].astype(str))),
    }