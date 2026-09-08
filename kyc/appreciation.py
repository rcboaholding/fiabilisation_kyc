"""
Matrice d'appréciation globale — reproduction fidèle du script R
`appreciation_globale.r`.

Liaison des modèles :
  - Taux_qualite_agent : taux de conformité des règles qualité par agent (filiale + expl)
  - Notation           : dernière notation de l'agent (flux ou stock selon AppreciationConfig.methode_taux)
  - TauxEvolution      : taux d'évolution flux ou stock (moyenne PP + PM, dernier point)
  - trimestre_actuel   : trimestres écoulés depuis la date de démarrage (admin), plafonné 1..4

Notations utilisées (libellés Django, identiques au script R mis à jour) :
  "Très Bien", "Bien", "Passable", "Insuffisant".
Une notation absente/vide est traitée favorablement (comme dans le script R).
"""
from datetime import date


def appreciation_qualite(taux_qualite, notation):
    """Appréciation qualité = f(taux qualité agent, notation flux).

    Matrice (script R `appreciation_globale.r`) :
      Taux <= 80           : tous -> Insatisfaisant
      80 < Taux < 100      : Très Bien / Bien / (vide) -> Satisfaisant ;
                             Passable / Insuffisant     -> Insatisfaisant
      Taux == 100          : Très Bien / (vide)         -> Très satisfaisant ;
                             Bien / Passable            -> Satisfaisant ;
                             Insuffisant                -> Insatisfaisant
    """
    note = (notation or "").strip()
    tq = taux_qualite if taux_qualite is not None else 0.0
    empty = note == ""

    if tq <= 80:
        return "Insatisfaisant"

    if tq < 100:                   
        if empty or note in ("Très Bien", "Bien"):
            return "Satisfaisant"
                               
        return "Insatisfaisant"

                          
    if empty or note == "Très Bien":
        return "Très satisfaisant"
    if note in ("Bien", "Passable"):
        return "Satisfaisant"
                 
    return "Insatisfaisant"


                                                                        
QUARTER_BANDS = {1: (5, 30), 2: (30, 60), 3: (60, 90), 4: (65, 95)}


def appreciation_globale(taux_evolution, appr_qualite, trimestre):
    """Appréciation globale = f(taux d'évolution, appréciation qualité, trimestre)."""
    try:
        t = int(trimestre or 1)
    except (TypeError, ValueError):
        t = 1
    low, high = QUARTER_BANDS.get(t, QUARTER_BANDS[1])

    if taux_evolution is None:
        # Taux d'évolution absent = traité favorablement (comme la notation vide côté qualité).
        prefix = "Bon"
    elif taux_evolution <= low:
        prefix = "Faible"
    elif taux_evolution <= high:
        prefix = "Moyen"
    else:
        prefix = "Bon"

    suffix = {"Très satisfaisant": "++", "Satisfaisant": "+", "Insatisfaisant": "-"}.get(appr_qualite, "")
    return prefix + suffix


def mesure_from_appreciation(appr_globale):
    """Mesure (action RH) déduite de l'appréciation globale."""
    a = appr_globale or ""
    if a in ("Faible++", "Faible+", "Moyen++", "Moyen+"):
        return "Demande d'explication et décote sur le bonus si motif infondé"
    if a in ("Faible-", "Moyen-", "Bon-"):
        return "Demande d'explication avec sanction forte et décote sur le bonus si motif infondé"
    if a in ("Bon++", "Bon+"):
        return "Impact positif sur le bonus"
    return ""


def current_trimestre(start_date, today=None):
    """Trimestre courant = nb de trimestres (3 mois) écoulés depuis la date de
    démarrage + 1, plafonné à [1, 4]."""
    if not start_date:
        return 1
    today = today or date.today()
    if today < start_date:
        return 1
    months = (today.year - start_date.year) * 12 + (today.month - start_date.month)
    if today.day < start_date.day:
        months -= 1
    return max(1, min(4, months // 3 + 1))
