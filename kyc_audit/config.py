# -*- coding: utf-8 -*-
"""Référentiel de l'audit KYC : champs, libellés, filiales, seuils.

Toute adaptation (nouveau champ, nouveau libellé, nouvelle filiale) se fait ici.
"""
from datetime import date

PP_FIELDS = ["PAYNAIS", "PROFESSION", "SALAIRE", "NUMID", "CODAPE", "TEL",
             "DATNAIS", "ADRESSE", "DATVALID", "ORIGINE_REVENU", "PAYS_RESID"]
PM_FIELDS = ["CODAPE", "AGEC", "CAPITAL", "CA", "RESULTAT", "RCSNO",
             "ORIGINE_REVENU", "TEL"]

PP_LABELS = {
    "PAYNAIS": "Lieu de naissance",
    "PROFESSION": "Profession",
    "SALAIRE": "Revenu",
    "NUMID": "NIN",
    "CODAPE": "Code agent économique",
    "TEL": "Téléphone",
    "ADRESSE": "Adresse",
    "DATNAIS": "Date de naissance",
    "DATVALID": "Date de validité CIN",
    "ORIGINE_REVENU": "Origine du revenu",
    "PAYS_RESID": "Pays de résidence",
}
PM_LABELS = {
    "CODAPE": "Secteur d'activité",
    "AGEC": "Code agent économique",
    "CAPITAL": "Capital social",
    "CA": "Chiffre d'affaires",
    "RESULTAT": "Résultat net",
    "RCSNO": "Numéro de registre de commerce",
    "ORIGINE_REVENU": "Origine du revenu",
    "TEL": "Téléphone",
}

AUTRES_LABELS = {
    "DATREV": "Date de révision KYC",
    "RISQUE": "Niveau de risque",
    "PPE": "Personne politiquement exposée",
    "RESID": "Résidence",
    "DEVISE": "Devise du compte",
    "BOITE_POSTALE": "Boîte postale",
    "EMPLOYEUR": "Employeur",
    "ACTIONNAIRE": "Actionnaire",
    "MANDATAIRE": "Mandataire",
    "NUMERO_FISCAL": "Numéro fiscal",
    "ADRESSE_SOCIALE": "Adresse sociale",
    "DATOUV": "Date d'ouverture",
    "LIEU_DELIVRANCE_CIN": "Lieu de délivrance CIN",
    "CONSENT_BIC": "Consentement BIC",
    "NOM_PAYS": "Pays de naissance",
    "PAYS_JUR": "Pays juridique",
    "INTITULE_COMPTE": "Intitulé du compte",
}


def label(field, typologie):
    src = PP_LABELS if typologie == "PP" else PM_LABELS
    return src.get(field) or AUTRES_LABELS.get(field) or field


FILIALES = {
    "SN": ("BOA SN", "Sénégal"),
    "CI": ("BOA CI", "Côte d'Ivoire"),
    "BJ": ("BOA BJ", "Bénin"),
    "BF": ("BOA BF", "Burkina Faso"),
    "ML": ("BOA ML", "Mali"),
    "NE": ("BOA NE", "Niger"),
    "TG": ("BOA TG", "Togo"),
    "MG": ("BOA MG", "Madagascar"),
    "MR": ("BOA MR", "Mauritanie"),
    "CD": ("BOA CD", "RD Congo"),
    "RDC": ("BOA CD", "RD Congo"),  # ancien code, conserve pour les donnees historiques
    "KE": ("BOA KE", "Kenya"),
    "TZ": ("BOA TZ", "Tanzanie"),
    "UG": ("BOA UG", "Ouganda"),
    "RW": ("BOA RW", "Rwanda"),
    "GH": ("BOA GH", "Ghana"),
    "KM": ("BOA KM", "Comores"),
    "FR": ("BOA FR", "France"),
    "CG": ("CG", "Congo"),
    "BCB": ("BCB", "Burundi"),
}


def filiale_info(iso2):
    iso2 = iso2.strip().upper()
    if iso2 in FILIALES:
        nom, pays = FILIALES[iso2]
    else:
        nom, pays = f"BOA {iso2}", iso2
    return {"iso2": iso2, "nom_regles": nom, "pays": pays,
            "libelle": f"BOA {iso2}"}


SEUIL_FAIBLE = 80
SEUIL_MOYEN = 100

RISQUE_ELEVE = "Risque eleve"

DEVISES_LOCALES = {"XOF", "XAF", "FCFA", "CFA", "F CFA", "MGA", "MRU", "CDF",
                   "KES", "TZS", "UGX", "RWF", "GHS", "KMF", "BIF", "EUR_LOCAL"}

ECHEANCES = [(92, "À échoir < 3 mois"),
             (183, "À échoir 3–6 mois"),
             (365, "À échoir 6–12 mois")]
ECHEANCE_LOINTAINE = "À échoir > 12 mois"
RETARDS = [(182, "< 6 mois"), (365, "6 à 12 mois"), (730, "1 à 2 ans")]
RETARD_LOINTAIN = "> 2 ans"

REGLES_JSON = "quality_rules_export.json"
TEMPLATE_PPTX = "exemple rapport.pptx"
PIED_DE_PAGE = "Conformité BOA Group | Fiabilisation KYC"


def fichiers_entree(iso2):
    return f"pp_{iso2.upper()}_STOCK.csv", f"pm_{iso2.upper()}_STOCK.csv"


def aujourd_hui():
    return date.today()
