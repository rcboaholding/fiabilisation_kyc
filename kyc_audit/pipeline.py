# -*- coding: utf-8 -*-
"""Orchestration de l'audit : charge les exports et produit le dictionnaire de
résultats consommé par le générateur de rapport (et sérialisable en JSON)."""
import os
from collections import Counter

from . import completude, qualite, scoring
from .config import (PM_FIELDS, PP_FIELDS, REGLES_JSON, aujourd_hui,
                     fichiers_entree, filiale_info)
from .dataset import charger, nombre


def _segments(t, champs, regles, filiale, typologie, today):
    """Complétude / qualité / scoring pour chaque segment sensible."""
    defs = [("Comptes en devise étrangère", t.segment_devise()),
            ("Clients non-résidents", t.segment_non_resident())]
    if t.a_colonne("PPE"):
        defs.insert(0, ("Clients PPE", t.segment_ppe()))

    vus, union = set(), []
    for _, s in defs:
        for r in s.rows:
            if id(r) not in vus:
                vus.add(id(r))
                union.append(r)
    defs.append(("Périmètre sensible consolidé", t.sous_table(union)))

    out = {}
    for libelle, s in defs:
        out[libelle] = {
            "n": len(s),
            "part": round(len(s) / len(t) * 100, 2) if len(t) else 0.0,
            "completude": completude.analyser(s, champs),
            "qualite": qualite.evaluer(s, regles, filiale, typologie),
            "scoring": scoring.analyser(s, today) if len(s) else None,
        }

    from .config import DEVISES_LOCALES
    devises = Counter()
    for r in t.rows:
        for jeton in (t.get(r, "DEVISE") or "").split(","):
            jeton = jeton.strip().upper()
            if jeton and jeton not in DEVISES_LOCALES:
                devises[jeton] += 1
    out["_devises"] = dict(devises.most_common(12))

    out["_croisements"] = {
        "Non-résident et devise étrangère":
            sum(1 for r in t.rows if t.est_non_resident(r) and t.est_devise_etrangere(r)),
        "Risque élevé et non-résident":
            sum(1 for r in t.rows if t.est_risque_eleve(r) and t.est_non_resident(r)),
        "Risque élevé et devise étrangère":
            sum(1 for r in t.rows if t.est_risque_eleve(r) and t.est_devise_etrangere(r)),
    }
    if t.a_colonne("PPE"):
        out["_croisements"]["PPE et non-résident"] = \
            sum(1 for r in t.rows if t.est_ppe(r) and t.est_non_resident(r))
        out["_croisements"]["PPE et devise étrangère"] = \
            sum(1 for r in t.rows if t.est_ppe(r) and t.est_devise_etrangere(r))

    for libelle, s in defs:
        out[libelle]["risque_eleve"] = sum(1 for r in s.rows if s.est_risque_eleve(r))
    return out


def _profil_risque_eleve(t, champs, regles, filiale, typologie, today, qualite_globale):
    """Complétude, qualité et exposition des seuls clients de niveau « risque élevé ».

    qualite_globale est réutilisée telle quelle : réévaluer les règles sur tout le
    portefeuille coûterait une seconde passe complète pour rien.
    """
    s = t.segment_risque_eleve()
    n = len(s)
    bloc = {
        "n": n,
        "part": round(n / len(t) * 100, 2) if len(t) else 0.0,
        "completude": completude.analyser(s, champs) if n else None,
        "qualite": qualite.evaluer(s, regles, filiale, typologie) if n else None,
        "scoring": scoring.analyser(s, today) if n else None,
        "non_residents": sum(1 for r in s.rows if s.est_non_resident(r)),
        "devise_etrangere": sum(1 for r in s.rows if s.est_devise_etrangere(r)),
        "ppe": sum(1 for r in s.rows if s.est_ppe(r)) if t.a_colonne("PPE") else None,
    }
    if n:
        global_par_nom = {r["nom"]: r for r in qualite_globale["regles"]}
        for r in bloc["qualite"]["regles"]:
            g = global_par_nom.get(r["nom"])
            r["part_globale"] = g["part"] if g else None
            r["anomalies_globales"] = g["anomalies"] if g else None
    return bloc


def _valeurs_par_defaut(t, champs):
    """Valeurs de complaisance : renseignées au sens du taux, non informatives."""
    out = {}
    for c in champs:
        i = t.col(c)
        if i is None:
            continue
        vals = [r[i] for r in t.rows if r[i]]
        cnt = Counter(vals)
        nc = sum(n for v, n in cnt.items()
                 if v.strip().lower().startswith(("non communiqu", "non renseign")))
        zeros = sum(n for v, n in cnt.items() if nombre(v) == 0)
        out[c] = {
            "renseigne": len(vals),
            "distinctes": len(cnt),
            "top": cnt.most_common(3),
            "non_communiquee": nc,
            "part_non_communiquee": round(nc / len(t) * 100, 1) if len(t) else 0.0,
            "valeur_nulle": zeros,
            "part_valeur_nulle": round(zeros / len(t) * 100, 1) if len(t) else 0.0,
        }
    return out


def _champs_vides(t, candidats):
    """Champs présents dans l'export mais jamais alimentés."""
    out = {}
    for c in candidats:
        i = t.col(c)
        if i is None:
            out[c] = None
            continue
        remplis = sum(1 for r in t.rows if r[i])
        out[c] = round(remplis / len(t) * 100, 1) if len(t) else 0.0
    return out


def executer(iso2, dossier=".", regles_path=None, today=None,
             pp_path=None, pm_path=None):
    fil = filiale_info(iso2)
    today = today or aujourd_hui()
    qualite.fixer_date_reference(today)

    f_pp, f_pm = fichiers_entree(fil["iso2"])
    p_pp = pp_path or os.path.join(dossier, f_pp)
    p_pm = pm_path or os.path.join(dossier, f_pm)
    for p in (p_pp, p_pm):
        if not os.path.exists(p):
            raise FileNotFoundError(
                f"Fichier introuvable : {p}\n"
                f"Attendus dans {os.path.abspath(dossier)} : {f_pp} et {f_pm}")

    regles = qualite.charger_regles(regles_path or os.path.join(dossier, REGLES_JSON))
    nom_regles = fil["nom_regles"]

    pp, diag_pp = charger(p_pp, "PP")
    pm, diag_pm = charger(p_pm, "PM")

    q_pp = qualite.evaluer(pp, regles, nom_regles, "PP")
    q_pm = qualite.evaluer(pm, regles, nom_regles, "PM")

    res = {
        "filiale": fil,
        "date_audit": today.isoformat(),
        "source": {"pp": diag_pp, "pm": diag_pm},
        "referentiel": {
            "regles_totales": len(regles),
            "regles_pp": len(qualite.selectionner(regles, nom_regles, "PP")),
            "regles_pm": len(qualite.selectionner(regles, nom_regles, "PM")),
            "regles_filiale": len(qualite.selectionner(regles, nom_regles, "PP"))
                              + len(qualite.selectionner(regles, nom_regles, "PM")),
            "champs_pp": PP_FIELDS,
            "champs_pm": PM_FIELDS,
        },
        "completude": {
            "pp": completude.analyser(pp, PP_FIELDS),
            "pm": completude.analyser(pm, PM_FIELDS),
            "pp_agences": completude.par_agence(pp, PP_FIELDS),
            "pm_agences": completude.par_agence(pm, PM_FIELDS),
        },
        "qualite": {"pp": q_pp, "pm": q_pm},
        "scoring": {
            "pp": scoring.analyser(pp, today),
            "pm": scoring.analyser(pm, today),
            "ppe_pp": scoring.coherence_ppe(pp),
            "ppe_pm": scoring.coherence_ppe(pm),
        },
        "risque_eleve": {
            "pp": _profil_risque_eleve(pp, PP_FIELDS, regles, nom_regles, "PP",
                                       today, q_pp),
            "pm": _profil_risque_eleve(pm, PM_FIELDS, regles, nom_regles, "PM",
                                       today, q_pm),
        },
        "segments": {
            "pp": _segments(pp, PP_FIELDS, regles, nom_regles, "PP", today),
            "pm": _segments(pm, PM_FIELDS, regles, nom_regles, "PM", today),
        },
        "valeurs": {
            "pp": _valeurs_par_defaut(pp, PP_FIELDS),
            "pm": _valeurs_par_defaut(pm, PM_FIELDS),
        },
        "champs_non_alimentes": {
            "pp": _champs_vides(pp, ["BOITE_POSTALE", "EMPLOYEUR", "LIEU_DELIVRANCE_CIN",
                                     "PPE", "DATREV"]),
            "pm": _champs_vides(pm, ["BOITE_POSTALE", "ACTIONNAIRE", "MANDATAIRE",
                                     "NUMERO_FISCAL", "ADRESSE_SOCIALE", "PPE", "DATREV"]),
        },
    }

    res["qualite"]["pp"]["doublons_regles"] = qualite.doublons_de_regles(
        res["qualite"]["pp"]["regles"])
    res["qualite"]["pm"]["doublons_regles"] = qualite.doublons_de_regles(
        res["qualite"]["pm"]["regles"])
    return res
