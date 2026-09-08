# -*- coding: utf-8 -*-
"""Chargement et normalisation des exports pp_XX_STOCK.csv / pm_XX_STOCK.csv.

Ces fichiers sont produits par Script_V3.r : ils sont déjà dédoublonnés
(agreger_doublons) et purgés des clients non fiabilisables. Deux conséquences
traitées ici :
  - la devise locale a été vidée par le script R, donc DEVISE non vide
    signifie « compte en devise étrangère » ;
  - les colonnes des clients multi-lignes ont été concaténées par une virgule
    ("EUR,XOF", "O,N"), d'où la lecture par jetons pour DEVISE et RESID.
"""
import csv, re, sys
from collections import Counter
from datetime import datetime

from .config import DEVISES_LOCALES, RISQUE_ELEVE

csv.field_size_limit(10_000_000)

_ESPACES = re.compile("[\t    ]+")


def nettoyer(v):
    """Équivalent de clean_spaces() du script R."""
    if not v:
        return ""
    v = v.replace('"', "")
    v = _ESPACES.sub(" ", v)
    return v.strip()


def parse_date(v):
    """Accepte les formats de l'export : ISO, JJ/MM/AAAA, JJ-MM-AAAA, JJ/MM/AA."""
    v = (v or "").strip()
    if not v:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y%m%d",
                "%d/%m/%y", "%d-%m-%y"):
        try:
            return datetime.strptime(v, fmt).date()
        except ValueError:
            continue
    return None


def nombre(v):
    """Parse un numérique en tolérant la notation R ('1e+06') et la virgule décimale."""
    if v is None:
        return None
    v = str(v).strip().replace(" ", "")
    if not v:
        return None
    try:
        return float(v.replace(",", "."))
    except ValueError:
        return None


def _ouvrir(path):
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            f = open(path, "r", encoding=enc, newline="", errors="strict")
            f.readline()
            f.seek(0)
            return f, enc
        except (UnicodeDecodeError, LookupError):
            try:
                f.close()
            except Exception:
                pass
    return open(path, "r", encoding="latin-1", newline=""), "latin-1"


class Table:
    """Table en mémoire, accès colonne par nom. Les alias plateforme sont résolus."""

    ALIAS = {"ORIGINE_REV": "ORIGINE_REVENU", "DATEREV": "DATREV"}

    def __init__(self, header, rows, typologie):
        self.header = header
        self.rows = rows
        self.typologie = typologie
        self._idx = {h: i for i, h in enumerate(header)}

    def col(self, name):
        name = self.ALIAS.get(name, name)
        return self._idx.get(name)

    def get(self, row, name):
        i = self.col(name)
        return row[i] if i is not None and i < len(row) else ""

    def valeurs(self, name):
        i = self.col(name)
        return [] if i is None else [r[i] for r in self.rows]

    def sous_table(self, rows):
        return Table(self.header, rows, self.typologie)

    def filtrer(self, predicat):
        return self.sous_table([r for r in self.rows if predicat(r)])

    def __len__(self):
        return len(self.rows)

    @staticmethod
    def _jetons(v):
        return [t.strip().upper() for t in (v or "").split(",") if t.strip()]

    def est_devise_etrangere(self, row):
        """Vrai si au moins un jeton de DEVISE n'est pas une devise locale."""
        return any(t not in DEVISES_LOCALES for t in self._jetons(self.get(row, "DEVISE")))

    def est_non_resident(self, row):
        return "N" in self._jetons(self.get(row, "RESID"))

    def est_ppe(self, row):
        return "O" in self._jetons(self.get(row, "PPE"))

    def segment_devise(self):
        return self.filtrer(self.est_devise_etrangere)

    def segment_non_resident(self):
        return self.filtrer(self.est_non_resident)

    def segment_ppe(self):
        return self.filtrer(self.est_ppe)

    def est_risque_eleve(self, row):
        return self.get(row, "RISQUE").strip() == RISQUE_ELEVE

    def segment_risque_eleve(self):
        return self.filtrer(self.est_risque_eleve)

    def a_colonne(self, name):
        return self.col(name) is not None


def charger(path, typologie, verbose=True):
    """Lit un export ; retourne (Table, diagnostic)."""
    f, enc = _ouvrir(path)
    rd = csv.reader(f, delimiter=";", quotechar='"')
    header = [nettoyer(h).upper() for h in next(rd)]
    rows, anormales = [], 0
    for r in rd:
        if not r or all(not x.strip() for x in r):
            continue
        if len(r) != len(header):
            anormales += 1
            r = (r + [""] * len(header))[:len(header)]
        rows.append([nettoyer(x) for x in r])
    f.close()

    fantomes = [h for i, h in enumerate(header)
                if re.fullmatch(r"V\d+", h) and not any(r[i] for r in rows)]

    t = Table(header, rows, typologie)
    diag = {
        "fichier": path,
        "encodage": enc,
        "lignes": len(rows),
        "colonnes": header,
        "lignes_anormales": anormales,
        "colonnes_fantomes": fantomes,
    }
    ic = t.col("CLIENT")
    if ic is not None:
        c = Counter(r[ic] for r in rows)
        diag["clients_uniques"] = len(c)
        diag["clients_multi_lignes"] = sum(1 for v in c.values() if v > 1)
    if verbose:
        print(f"[charger] {path} : {len(rows)} lignes, {len(header)} colonnes "
              f"({enc}){', ' + str(anormales) + ' anormales' if anormales else ''}")
    return t, diag
