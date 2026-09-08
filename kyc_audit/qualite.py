# -*- coding: utf-8 -*-
"""Qualité — réplique du moteur de règles de la plateforme KYC.

Sources répliquées (kyc/views.py) :
  - _dq_eval_condition   : sémantique de chaque opérateur ;
  - _dq_conditions_flag  : combinaison en forme normale disjonctive
                           (« OU » ouvre un groupe, « ET » à l'intérieur) ;
  - compute_quality_rate_by_typology : taux = Σ conformes / Σ évaluations.

Une règle dont la liste `filiale` ne contient pas la filiale auditée est
écartée (comme evaluate_data_quality_rule qui renvoie alors total = 0).
"""
import json
import re
from collections import Counter

from .dataset import parse_date


def _age(valeur, today):
    d = parse_date(valeur)
    if not d:
        return None
    return today.year - d.year - ((today.month, today.day) < (d.month, d.day))


def evaluer_condition(op, valeur, cible, today):
    """True si la condition « matche », c.-à-d. contribue à l'anomalie."""
    val = str(valeur or "").strip()
    tgt = str(cible or "").strip()

    if op == "=":
        return val == tgt
    if op == "!=":
        return val != tgt
    if op in (">", "<", ">=", "<="):
        try:
            a = float(val.replace(",", "."))
            b = float(tgt.replace(",", "."))
        except ValueError:
            return False
        return {">": a > b, "<": a < b, ">=": a >= b, "<=": a <= b}[op]
    if op == "contains":
        return tgt.lower() in val.lower()
    if op == "not_contains":
        return tgt.lower() not in val.lower()
    if op == "word_contains":
        return bool(tgt and re.search(rf"(?<!\w){re.escape(tgt)}(?!\w)", val, re.I))
    if op == "word_not_contains":
        return not bool(tgt and re.search(rf"(?<!\w){re.escape(tgt)}(?!\w)", val, re.I))
    if op == "contains_alpha":
        return any(c.isalpha() for c in val)
    if op == "contains_digit":
        return any(c.isdigit() for c in val)
    if op == "is_empty":
        return not val
    if op == "is_not_empty":
        return bool(val)
    if op == "expired":
        d = parse_date(val)
        return bool(d and d < today)
    if op == "age_gt":
        a = _age(val, today)
        try:
            return a is not None and a > int(tgt)
        except ValueError:
            return False
    if op == "age_lt":
        a = _age(val, today)
        try:
            return a is not None and a < int(tgt)
        except ValueError:
            return False
    if op == "min_length":
        try:
            return len(val) < int(tgt)
        except ValueError:
            return False
    if op == "max_length":
        try:
            return len(val) > int(tgt)
        except ValueError:
            return False
    if op == "regex":
        try:
            return re.search(tgt, val) is not None
        except re.error:
            return False
    return False


def combiner(conditions, lire, today):
    """Forme normale disjonctive ; le connecteur de la 1re condition est ignoré."""
    resultat, courant, demarre = False, True, False
    for c in conditions:
        m = evaluer_condition(c["operator"], lire(c["field_name"]), c["value"], today)
        if not demarre:
            courant, demarre = m, True
        elif (c.get("logic") or "AND") == "OR":
            resultat = resultat or courant
            courant = m
        else:
            courant = courant and m
    return (resultat or courant) if demarre else False


def charger_regles(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)["rules"]


def regle_applicable(regle, filiale):
    f = (regle.get("filiale") or "").strip()
    if not f:
        return True
    return filiale in [p.strip() for p in f.split("|") if p.strip()]


def selectionner(regles, filiale, typologie):
    return [r for r in regles
            if r.get("active")
            and r["applicability"] == typologie
            and regle_applicable(r, filiale)]


def evaluer(t, regles, filiale, typologie, exemples=0):
    """Évalue toutes les règles applicables ; retourne le détail + le taux agrégé."""
    retenues = selectionner(regles, filiale, typologie)
    n = len(t)
    detail, total_ok, total_eval = [], 0, 0
    touches = set()  # index des clients portant au moins une anomalie

    for regle in retenues:
        conds = regle["conditions"]
        absents = sorted({c["field_name"] for c in conds
                          if t.col(c["field_name"]) is None})
        anomalies, echantillon = 0, []
        for i, r in enumerate(t.rows):
            if combiner(conds, lambda f: t.get(r, f), _TODAY[0]):
                anomalies += 1
                touches.add(i)
                if len(echantillon) < exemples:
                    echantillon.append({
                        "client": t.get(r, "CLIENT"),
                        "agence": t.get(r, "AGENCE"),
                        "expl": t.get(r, "EXPL"),
                    })
        total_ok += n - anomalies
        total_eval += n
        detail.append({
            "nom": regle["name"],
            "description": (regle.get("description") or "").strip(),
            "total": n,
            "anomalies": anomalies,
            "part": round(anomalies / n * 100, 1) if n else 0.0,
            "taux_conformite": round((n - anomalies) / n * 100, 1) if n else 0.0,
            "evaluable": not absents,
            "champs_absents": absents,
            "conditions": [f'{c["logic"]} {c["field_name"]} {c["operator"]} {c["value"]}'
                           for c in conds],
            "exemples": echantillon,
        })

    detail.sort(key=lambda x: -x["anomalies"])
    return {
        "typologie": typologie,
        "taux_qualite": round(total_ok / total_eval * 100, 1) if total_eval else 0.0,
        "nb_regles": len(detail),
        "nb_regles_evaluables": sum(1 for d in detail if d["evaluable"]),
        "total_anomalies": sum(d["anomalies"] for d in detail),
        "clients_anomalie": len(touches),
        "part_clients_anomalie": round(len(touches) / n * 100, 1) if n else 0.0,
        "regles": detail,
    }


_TODAY = [None]


def fixer_date_reference(d):
    _TODAY[0] = d


def doublons_de_regles(detail):
    """Repère les règles portant exactement les mêmes conditions (double comptage)."""
    par_cond = {}
    for d in detail:
        cle = tuple(d["conditions"])
        par_cond.setdefault(cle, []).append(d["nom"])
    return [noms for noms in par_cond.values() if len(noms) > 1]
