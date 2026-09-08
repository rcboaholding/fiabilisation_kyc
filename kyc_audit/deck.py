# -*- coding: utf-8 -*-
"""Construction du rapport PPTX d'audit KYC (charte « exemple rapport.pptx »)."""
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from .config import PIED_DE_PAGE, PM_FIELDS, PP_FIELDS, label
from .theme import (AMBER, AMBER_PALE, BODY, CARD, CYAN, DIVIDER, GREEN,
                    GREEN_OK, GREEN_PALE, H, KPI_H, LEFT, MUTED, NAVY, RED,
                    RED_PALE, W, WHITE, anneau, barres, bloc, cloture,
                    couleur_taux, couverture, encart, fr, kpi, note, page, pc,
                    rect, sommaire, table, txt)
from .scoring import NIVEAUX, NIVEAUX_LIB, STATUTS

SECTIONS = [
    ("I", "Périmètre et méthodologie"),
    ("II", "Complétude des données"),
    ("III", "Qualité des données"),
    ("IV", "PPE, comptes en devise et non-résidents"),
    ("V", "Revue Scoring AML"),
    ("VI", "Synthèse et plan d'action"),
]
BANDEAU = {rom: f"{rom}.  {lib.upper()}" for rom, lib in SECTIONS}

MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet",
        "août", "septembre", "octobre", "novembre", "décembre"]


def _date_fr(d):
    return f"{d.day} {MOIS[d.month - 1]} {d.year}"


def _couverture(prs, R, d):
    couverture(
        prs,
        [("RAPPORT D'AUDIT", 22, True, None),
         ("COMPLÉTUDE ET QUALITÉ DES DONNÉES KYC", 28, True, None)],
        f"{R['filiale']['libelle']} — {R['filiale']['pays']}",
        f"Stock au {_date_fr(d)}  |  {fr(R['source']['pp']['lignes'])} particuliers  |  "
        f"{fr(R['source']['pm']['lignes'])} entreprises",
        PIED_DE_PAGE)


def _sommaire(prs):
    sommaire(prs, [(rom, lib) for rom, lib in SECTIONS], PIED_DE_PAGE)


def _perimetre(prs, R, d):
    s = page(prs, BANDEAU["I"], "Base de l'audit", PIED_DE_PAGE)
    spp, spm = R["source"]["pp"], R["source"]["pm"]
    ref = R["referentiel"]

    for i, (val, lib, nt, col) in enumerate([
        (fr(spp["lignes"]), "Clients particuliers",
         f"{fr(spp.get('clients_uniques', spp['lignes']))} identifiants uniques", GREEN),
        (fr(spm["lignes"]), "Clients entreprises",
         f"{fr(spm.get('clients_uniques', spm['lignes']))} identifiants uniques", GREEN),
        (str(len(PP_FIELDS) + len(PM_FIELDS)), "Champs de complétude suivis",
         f"{len(PP_FIELDS)} champs PP · {len(PM_FIELDS)} champs PM", NAVY),
        (str(ref["regles_filiale"]), "Règles qualité applicables",
         f"{ref['regles_pp']} règles PP · {ref['regles_pm']} règles PM", NAVY),
    ]):
        kpi(s, 0.4 + i * 2.32, 1.22, 2.24, KPI_H, val, lib, nt, col, vsize=20)

    y = bloc(s, 2.20, "Méthodes de calcul appliquées")
    table(s, LEFT, y, 4.5, [
        ["Complétude — Script_V3.r", "Définition"],
        ["Taux par champ", "floor(100 − vides / effectif × 100)"],
        ["Taux global", "floor(100 × (1 − vides / (champs × effectif)))"],
        ["Valeur « vide »", "chaîne vide après nettoyage des espaces"],
        ["Appréciation", "Faible < 80 % · Moyen 80–99 % · Bon = 100 %"],
    ], col_w=[38, 62], font_size=7.5, row_h=0.30, head_h=0.26)

    table(s, 5.1, y, 4.5, [
        ["Qualité — plateforme KYC", "Définition"],
        ["Référentiel", f"{ref['regles_totales']} règles, {ref['regles_filiale']} applicables"],
        ["Combinaison", "forme normale disjonctive (OU ouvre un groupe)"],
        ["Taux qualité", "Σ conformes / Σ évaluations"],
        ["Date de référence", f"{_date_fr(d)} (opérateurs expired / age_gt)"],
    ], col_w=[38, 62], font_size=7.5, row_h=0.30, head_h=0.26)

    reserves = [
        "Les exports pp/pm_STOCK sont produits par Script_V3.r : clients déjà regroupés "
        "par identifiant et clients non fiabilisables (code OI) déjà exclus.",
        "La devise locale ayant été vidée par le script R, tout compte à devise renseignée "
        "est traité comme compte en devise étrangère.",
    ]
    absents = [c for c, v in R["champs_non_alimentes"]["pm"].items() if v is None]
    if absents:
        reserves.append("Champs absents de l'export entreprises : "
                        + ", ".join(absents) + " — les règles associées sont inévaluables.")
    encart(s, LEFT, 4.32, 9.2, 0.96, "Réserves méthodologiques", reserves,
           ton="amber", size=7.5)


def _synthese(prs, R):
    s = page(prs, BANDEAU["I"], "Synthèse exécutive", PIED_DE_PAGE)
    cpp, cpm = R["completude"]["pp"], R["completude"]["pm"]
    qpp, qpm = R["qualite"]["pp"], R["qualite"]["pm"]
    npp, npm = cpp["n"], cpm["n"]
    scpp, scpm = R["scoring"]["pp"], R["scoring"]["pm"]

    for i, (val, lib, nt, col) in enumerate([
        (f"{cpp['taux_global']} %", "Complétude particuliers (par champ)",
         f"{len(PP_FIELDS)} champs · {cpp['part_clients_complets']} % de clients complets",
         couleur_taux(cpp["taux_global"])),
        (f"{cpm['taux_global']} %", "Complétude entreprises (par champ)",
         f"{len(PM_FIELDS)} champs · {cpm['part_clients_complets']} % de clients complets",
         couleur_taux(cpm["taux_global"])),
        (f"{qpp['taux_qualite']} %", "Qualité particuliers",
         f"{fr(qpp['clients_anomalie'])} clients en anomalie "
         f"({qpp['part_clients_anomalie']} %)",
         couleur_taux(qpp["taux_qualite"], 85, 95)),
        (f"{qpm['taux_qualite']} %", "Qualité entreprises",
         f"{fr(qpm['clients_anomalie'])} clients en anomalie "
         f"({qpm['part_clients_anomalie']} %)",
         couleur_taux(qpm["taux_qualite"], 85, 95)),
    ]):
        kpi(s, 0.4 + i * 2.32, 1.22, 2.24, KPI_H, val, lib, nt, col, vsize=20)

    vpp, vpm = R["valeurs"]["pp"], R["valeurs"]["pm"]
    pire_pp = min(cpp["par_champ"].items(), key=lambda kv: (kv[1] is None, kv[1]))
    pire_pm = min(cpm["par_champ"].items(), key=lambda kv: (kv[1] is None, kv[1]))
    seg_pp = R["segments"]["pp"]["Clients non-résidents"]

    cartes = [
        ("1", "Complétude affichée et information exploitable divergent", RED, [
            f"{label('ORIGINE_REVENU','PP')} : {vpp['ORIGINE_REVENU']['part_non_communiquee']} % "
            f"des particuliers portent « Non communiquée », comptée comme renseignée.",
            f"Côté entreprises, la proportion atteint "
            f"{vpm['ORIGINE_REVENU']['part_non_communiquee']} %.",
            f"Le champ le plus faible reste {label(pire_pp[0],'PP')} ({pire_pp[1]} %).",
        ]),
        ("2", "Le bloc financier entreprises est peu alimenté", RED, [
            f"{label(pire_pm[0],'PM')} : {pire_pm[1]} % de complétude "
            f"({fr(cpm['vides_par_champ'][pire_pm[0]])} dossiers vides).",
            f"{label('CAPITAL','PM')} {cpm['par_champ']['CAPITAL']} % · "
            f"{label('CA','PM')} {cpm['par_champ']['CA']} % · "
            f"{label('RCSNO','PM')} {cpm['par_champ']['RCSNO']} %.",
            f"Seuls {cpm['part_clients_complets']} % des clients entreprises sont "
            f"intégralement renseignés.",
        ]),
        ("3", "La revue scoring présente des défaillances", RED, [
            f"{scpp['part_sans_revision']} % des particuliers ont une date de revue scoring "
            f"échue ou absente ({fr(scpp['sans_revision'])} clients).",
            f"{scpm['part_sans_revision']} % côté entreprises "
            f"({fr(scpm['sans_revision'])} sociétés).",
            f"{fr(scpp['risque_eleve']['sans_revision'])} clients à risque élevé sont concernés, "
            f"soit {scpp['risque_eleve']['part']} % des clients de cette catégorie.",
        ]),
        ("4", "Les segments sensibles concentrent les écarts", AMBER, [
            f"Non-résidents : complétude {seg_pp['completude']['taux_global']} % et qualité "
            f"{seg_pp['qualite']['taux_qualite']} % (global {cpp['taux_global']} % / "
            f"{qpp['taux_qualite']} %).",
            f"{fr(R['segments']['pm']['Comptes en devise étrangère']['n'])} entreprises en devise "
            f"étrangère ({R['segments']['pm']['Comptes en devise étrangère']['part']} % du portefeuille).",
            f"Périmètre sensible consolidé : "
            f"{fr(R['segments']['pp']['Périmètre sensible consolidé']['n'])} particuliers.",
        ]),
    ]
    for i, (num, titre, col, lignes) in enumerate(cartes):
        x = 0.4 + i * 2.32
        rect(s, x, 2.18, 2.24, 2.72, fill=CARD)
        rect(s, x, 2.18, 2.24, 0.05, fill=col)
        rect(s, x + 0.10, 2.30, 0.28, 0.28, fill=col)
        txt(s, x + 0.10, 2.30, 0.28, 0.28, num, size=10, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER)
        txt(s, x + 0.46, 2.30, 1.70, 0.56, titre, size=8.5, bold=True, color=col)
        txt(s, x + 0.12, 2.94, 2.00, 1.88, [f"•  {l}" for l in lignes], size=7,
            color=BODY, spacing=4)

    note(s, 4.98, "Les taux de complétude affichés sont des taux par champ (cellules "
                  "renseignées / cellules attendues) ; la part de clients complets compte les "
                  "clients sans aucun champ vide. Qualité : moteur de règles de la plateforme "
                  "KYC restreint aux règles applicables à la filiale.")


def _completude(prs, R, typo):
    c = R["completude"]["pp" if typo == "PP" else "pm"]
    champs = PP_FIELDS if typo == "PP" else PM_FIELDS
    nom = "Particuliers" if typo == "PP" else "Entreprises"
    s = page(prs, BANDEAU["II"], f"{nom} — {len(champs)} champs du référentiel",
             PIED_DE_PAGE)

    sous = sum(1 for v in c["par_champ"].values() if v is not None and v < 100)
    vides = sum(v for v in c["vides_par_champ"].values() if v)
    for i, (val, lib, nt, col) in enumerate([
        (f"{c['taux_global']} %", "Taux de complétude par champ",
         f"Appréciation : {c['appreciation']}", couleur_taux(c["taux_global"])),
        (f"{c['part_clients_complets']} %", "Clients intégralement renseignés",
         f"{fr(c['clients_complets'])} sur {fr(c['n'])} clients à fiabiliser",
         couleur_taux(c["part_clients_complets"], 50, 90)),
        (f"{sous} / {len(champs)}", "Champs sous les 100 %",
         f"{fr(vides)} cellules vides au total", NAVY),
    ]):
        kpi(s, 0.4 + i * 3.1, 1.22, 3.0, KPI_H, val, lib, nt, col, vsize=20)

    ordre = sorted([(f, v) for f, v in c["par_champ"].items() if v is not None],
                   key=lambda kv: kv[1])
    cats = [label(f, typo) for f, _ in ordre]
    vals = [v for _, v in ordre]
    txt(s, LEFT, 2.14, 5.4, 0.22, "Taux de complétude par champ (%)", size=9,
        bold=True, color=NAVY)
    plancher = max(0, (min(vals) - 10) // 5 * 5) if vals else 0
    barres(s, LEFT, 2.36, 5.4, 2.92, cats, [("Complétude", vals)],
           couleurs=[[couleur_taux(v) for v in vals]], horizontal=True,
           maxval=100, minval=plancher, size=7, gap=40)

    rows = [["Champ incomplet", "Taux de complétude", "Cellules vides"]]
    fills = {}
    for i, (f, v) in enumerate(ordre):
        if v < 100:
            rows.append([label(f, typo), f"{v} %", fr(c["vides_par_champ"][f])])
            if v < 80:
                fills[(len(rows) - 1, 1)] = RED_PALE
            elif v < 99:
                fills[(len(rows) - 1, 1)] = AMBER_PALE
    if len(rows) == 1:
        rows.append(["Aucun", "—", "—"])
    txt(s, 6.0, 2.14, 3.6, 0.22, "Champs ayant le taux de complétude le plus faible",
        size=9, bold=True, color=NAVY)
    rows = rows[:9]
    ytab = 2.36
    table(s, 6.0, ytab, 3.6, rows, col_w=[52, 22, 26], font_size=7.5,
          row_h=0.245, head_h=0.26, fills=fills)

    complets = [label(f, typo) for f, v in c["par_champ"].items() if v == 100]
    lignes = []
    if ordre:
        lignes.append(f"Champ le plus en retrait : {label(ordre[0][0], typo)} "
                      f"({ordre[0][1]} %, {fr(c['vides_par_champ'][ordre[0][0]])} vides).")
    if complets:
        lignes.append(f"Champs à 100 % : {', '.join(complets[:4])}"
                      f"{' …' if len(complets) > 4 else ''}.")
    lignes.append(f"{c['part_clients_complets']} % des clients à fiabiliser ne présentent "
                  f"aucun champ vide ({fr(c['clients_complets'])} sur {fr(c['n'])}).")
    ybas = ytab + 0.26 + 0.245 * (len(rows) - 1) + 0.14
    encart(s, 6.0, ybas, 3.6, max(0.6, 5.28 - ybas), None, lignes,
           ton="amber" if c["taux_global"] and c["taux_global"] < 95 else "green", size=7)


def _agences(prs, R):
    s = page(prs, BANDEAU["II"],
             "Répartition du taux de complétude par champ, par agence", PIED_DE_PAGE)
    app, apm = R["completude"]["pp_agences"], R["completude"]["pm_agences"]

    def bloc_table(x, titre, agences, seuil):
        txt(s, x, 1.22, 4.5, 0.22, titre, size=9, bold=True, color=NAVY)
        rows = [["Code", "Agence", "Clients", "Taux par champ"]]
        fills = {}
        for i, a in enumerate(agences[:9]):
            rows.append([a["agence"], a["libelle"][:26], fr(a["n"]), f"{a['taux']} %"])
            fills[(i + 1, 3)] = RED_PALE if a["taux"] < seuil else AMBER_PALE
        table(s, x, 1.46, 4.5, rows, col_w=[12, 50, 18, 20], font_size=7.5,
              row_h=0.26, head_h=0.26, fills=fills)

    bloc_table(LEFT, "Particuliers — 9 agences les plus en retrait (taux par champ)",
               app, 95)
    bloc_table(5.1, "Entreprises — 9 agences les plus en retrait (taux par champ)",
               apm, 75)

    sub = [a for a in apm if a["n"] >= 100][:12]
    if sub:
        txt(s, LEFT, 3.90, 9.2, 0.22,
            "Entreprises — agences les plus faibles (effectif ≥ 100 clients)",
            size=9, bold=True, color=NAVY)
        barres(s, LEFT, 4.12, 9.2, 1.16,
               [a["libelle"].replace("AGENCE ", "")[:14] for a in sub],
               [("Complétude", [a["taux"] for a in sub])],
               couleurs=[[couleur_taux(a["taux"], 75, 90) for a in sub]],
               maxval=100, size=6.5, gap=40)
    if app and apm:
        note(s, 3.62, f"Particuliers : de {app[0]['taux']} % ({app[0]['libelle']}) à "
                      f"{app[-1]['taux']} %.  Entreprises : de {apm[0]['taux']} % à "
                      f"{apm[-1]['taux']} %, soit {apm[-1]['taux'] - apm[0]['taux']} points d'écart.")


def _completude_reelle(prs, R):
    s = page(prs, BANDEAU["II"],
             "Complétude apparente et information réellement exploitable", PIED_DE_PAGE)
    encart(s, LEFT, 1.22, 9.2, 0.56, None,
           ["Le taux de complétude mesure la présence d'une valeur, pas sa portée "
            "informative : les valeurs par défaut et les champs jamais alimentés "
            "élèvent mécaniquement les taux affichés."], ton="red", size=8)

    y = bloc(s, 1.92, "Valeurs par défaut comptées comme renseignées")
    rows = [["Champ", "Typologie", "Valeur dominante", "Effectif", "% du portefeuille",
             "Taux affiché"]]
    fills = {}
    for typo, cle in (("PP", "pp"), ("PM", "pm")):
        v = R["valeurs"][cle]
        c = R["completude"][cle]
        for f, info in sorted(v.items(), key=lambda kv: -kv[1]["part_non_communiquee"])[:2]:
            if info["part_non_communiquee"] <= 0 and info["part_valeur_nulle"] <= 0:
                continue
            if info["part_non_communiquee"] >= info["part_valeur_nulle"]:
                dom, eff, part = "« Non communiquée »", info["non_communiquee"], info["part_non_communiquee"]
            else:
                dom, eff, part = "0 (valeur nulle)", info["valeur_nulle"], info["part_valeur_nulle"]
            rows.append([label(f, typo), "Particuliers" if typo == "PP" else "Entreprises",
                         dom, fr(eff), f"{part} %", f"{c['par_champ'][f]} %"])
            if part >= 50:
                fills[(len(rows) - 1, 4)] = RED_PALE
            elif part >= 20:
                fills[(len(rows) - 1, 4)] = AMBER_PALE
    table(s, LEFT, y, 9.2, rows, col_w=[22, 15, 22, 13, 15, 13], font_size=7.5,
          row_h=0.26, head_h=0.30, fills=fills)

    y2 = bloc(s, y + 0.30 + 0.26 * (len(rows) - 1) + 0.16,
              "Champs présents dans l'export mais non alimentés")
    rows2 = [["Champ", "Typologie", "Taux de remplissage", "Conséquence"]]
    fills2 = {}
    for typo, cle in (("PP", "pp"), ("PM", "pm")):
        for f, taux in R["champs_non_alimentes"][cle].items():
            if taux is None:
                rows2.append([label(f, typo), "Particuliers" if typo == "PP" else "Entreprises",
                              "absent de l'export", "règles associées inévaluables"])
                fills2[(len(rows2) - 1, 2)] = RED_PALE
            elif taux < 20:
                rows2.append([label(f, typo), "Particuliers" if typo == "PP" else "Entreprises",
                              f"{taux} %", "contrôle peu ou pas opérant"])
                fills2[(len(rows2) - 1, 2)] = RED_PALE if taux < 5 else AMBER_PALE
    if len(rows2) == 1:
        rows2.append(["Aucun", "—", "—", "—"])
    rows2 = rows2[:5]
    table(s, LEFT, y2, 9.2, rows2, col_w=[26, 16, 20, 38], font_size=7.5,
          row_h=0.24, head_h=0.26, fills=fills2)


def _dispositif(prs, R):
    s = page(prs, BANDEAU["III"], "Dispositif de contrôle et taux de conformité", PIED_DE_PAGE)
    qpp, qpm = R["qualite"]["pp"], R["qualite"]["pm"]

    for i, (val, lib, nt, col) in enumerate([
        (f"{qpp['taux_qualite']} %", "Taux de qualité particuliers",
         f"{qpp['nb_regles']} règles appliquées", couleur_taux(qpp["taux_qualite"], 85, 95)),
        (f"{qpm['taux_qualite']} %", "Taux de qualité entreprises",
         f"{qpm['nb_regles']} règles appliquées", couleur_taux(qpm["taux_qualite"], 85, 95)),
        (fr(qpp["clients_anomalie"]), "Particuliers en anomalie",
         f"{qpp['part_clients_anomalie']} % du portefeuille · "
         f"{fr(qpp['total_anomalies'])} anomalies", RED),
        (fr(qpm["clients_anomalie"]), "Entreprises en anomalie",
         f"{qpm['part_clients_anomalie']} % du portefeuille · "
         f"{fr(qpm['total_anomalies'])} anomalies", RED),
    ]):
        kpi(s, 0.4 + i * 2.32, 1.22, 2.24, KPI_H, val, lib, nt, col, vsize=20)

    y = bloc(s, 2.20, "Règles produisant le plus d'anomalies, toutes typologies")
    rows = [["Règle de contrôle", "Typologie", "Anomalies", "% du portefeuille", "Conformité"]]
    fusion = ([(r, "Particuliers") for r in qpp["regles"]]
              + [(r, "Entreprises") for r in qpm["regles"]])
    fusion.sort(key=lambda x: -x[0]["anomalies"])
    fills = {}
    for r, typ in fusion[:6]:
        rows.append([r["nom"], typ, fr(r["anomalies"]), f"{r['part']} %",
                     f"{r['taux_conformite']} %"])
        if r["part"] >= 20:
            fills[(len(rows) - 1, 3)] = RED_PALE
        elif r["part"] >= 5:
            fills[(len(rows) - 1, 3)] = AMBER_PALE
    table(s, LEFT, y, 9.2, rows, col_w=[42, 15, 14, 15, 14], font_size=7.5,
          row_h=0.27, head_h=0.28, fills=fills)

    lignes = ["Un client en anomalie n'est compté qu'une fois, quel que soit le nombre de "
              "règles déclenchées : ce compte est inférieur au total des anomalies.",
              "Le taux agrégé dilue les écarts : une règle défaillante sur 40 % du "
              "portefeuille est compensée par les règles sans anomalie."]
    for typ, q in (("particuliers", qpp), ("entreprises", qpm)):
        for doublon in q.get("doublons_regles", []):
            lignes.append(f"Règles {typ} au libellé différent mais aux conditions identiques : "
                          + " / ".join(f"« {n} »" for n in doublon)
                          + " — anomalies comptées deux fois.")
    inev = [r["nom"] for r in qpp["regles"] + qpm["regles"] if not r["evaluable"]]
    if inev:
        lignes.append("Règles inévaluables faute de champ dans l'export : "
                      + ", ".join(f"« {n} »" for n in inev[:3]) + ".")
    encart(s, LEFT, 4.44, 9.2, 0.84, "Précautions de lecture", lignes[:3],
           ton="amber", size=7.5)


def _anomalies(prs, R, typo):
    q = R["qualite"]["pp" if typo == "PP" else "pm"]
    nom = "Particuliers" if typo == "PP" else "Entreprises"
    n = R["completude"]["pp" if typo == "PP" else "pm"]["n"]
    s = page(prs, BANDEAU["III"], f"{nom} — anomalies détectées par règle", PIED_DE_PAGE)

    re_bloc = R["risque_eleve"]["pp" if typo == "PP" else "pm"]
    re_par_nom = {r["nom"]: r for r in (re_bloc["qualite"]["regles"] if re_bloc["n"] else [])}

    rows = [["Règle de contrôle", "Anomalies", "% portef.", "dont risque élevé", "Conformité"]]
    fills = {}
    for r in q["regles"]:
        nom_r = r["nom"] + ("" if r["evaluable"] else " (inévaluable)")
        re_r = re_par_nom.get(r["nom"])
        re_txt = (f"{fr(re_r['anomalies'])} ({re_r['part']} %)"
                  if re_r and re_r["anomalies"] else ("—" if re_bloc["n"] else "n. d."))
        rows.append([nom_r, fr(r["anomalies"]), f"{r['part']} %", re_txt,
                     f"{r['taux_conformite']} %"])
        if r["part"] >= 20:
            fills[(len(rows) - 1, 1)] = RED_PALE
        elif r["part"] >= 1:
            fills[(len(rows) - 1, 1)] = AMBER_PALE
        if re_r and re_r["part"] >= 30:
            fills[(len(rows) - 1, 3)] = RED_PALE
    nlig = len(rows)
    rh = 0.205 if nlig > 15 else (0.235 if nlig > 12 else 0.27)
    table(s, LEFT, 1.24, 5.75, rows, col_w=[44, 13, 12, 19, 12],
          font_size=6.5 if nlig > 15 else 7, row_h=rh, head_h=0.30, fills=fills)

    top = [r for r in q["regles"] if r["anomalies"]][:5]
    if top:
        txt(s, 6.30, 1.24, 3.3, 0.22, "Règles les plus impactantes", size=9,
            bold=True, color=NAVY)
        barres(s, 6.25, 1.48, 3.35, 1.85, [r["nom"][:20] for r in top],
               [("Anomalies", [r["anomalies"] for r in top])],
               couleurs=[[RED if r["part"] >= 20 else AMBER for r in top]],
               horizontal=True, size=6.5, num_fmt='#,##0', gap=40)

    lignes = []
    for r in top[:3]:
        lignes.append(f"« {r['nom']} » : {fr(r['anomalies'])} clients, soit {r['part']} % "
                      f"du portefeuille.")
    inev = [r["nom"] for r in q["regles"] if not r["evaluable"]]
    if inev:
        lignes.append("Inévaluable faute de champ : " + ", ".join(f"« {x} »" for x in inev) + ".")
    if not lignes:
        lignes = ["Aucune anomalie détectée sur les règles applicables."]
    encart(s, 6.25, 3.45, 3.35, 1.83, "Constats", lignes[:4], ton="red", size=7)


def _segments_carto(prs, R):
    s = page(prs, BANDEAU["IV"], "Cartographie des segments sensibles", PIED_DE_PAGE)
    spp, spm = R["segments"]["pp"], R["segments"]["pm"]
    npp, npm = R["completude"]["pp"]["n"], R["completude"]["pm"]["n"]

    ppe = spp.get("Clients PPE")
    cartes = [
        (fr(ppe["n"]) if ppe else "n. d.", "Clients PPE",
         f"{ppe['part']} % du portefeuille particuliers" if ppe else "champ absent des exports",
         NAVY),
        (fr(spp["Comptes en devise étrangère"]["n"] + spm["Comptes en devise étrangère"]["n"]),
         "Comptes en devise étrangère",
         f"{fr(spp['Comptes en devise étrangère']['n'])} PP · "
         f"{fr(spm['Comptes en devise étrangère']['n'])} PM", AMBER),
        (fr(spp["Clients non-résidents"]["n"] + spm["Clients non-résidents"]["n"]),
         "Clients non-résidents",
         f"{fr(spp['Clients non-résidents']['n'])} PP · "
         f"{fr(spm['Clients non-résidents']['n'])} PM", AMBER),
        (fr(spp["Périmètre sensible consolidé"]["n"] + spm["Périmètre sensible consolidé"]["n"]),
         "Périmètre sensible consolidé", "union des segments, sans double compte",
         RED),
    ]
    for i, (val, lib, nt, col) in enumerate(cartes):
        kpi(s, 0.4 + i * 2.32, 1.22, 2.24, KPI_H, val, lib, nt, col, vsize=20)

    y = bloc(s, 2.20, "Poids des segments dans chaque portefeuille")
    rows = [["Segment", "PP", "% PP", "dont risque élevé", "PM", "dont risque élevé"]]
    fills = {}
    for lib in ("Clients PPE", "Comptes en devise étrangère", "Clients non-résidents",
                "Périmètre sensible consolidé"):
        a, b = spp.get(lib), spm.get(lib)
        rea = f"{fr(a['risque_eleve'])} ({pc(a['risque_eleve'], a['n'])} %)" \
            if a and a["n"] else "—"
        reb = f"{fr(b['risque_eleve'])} ({pc(b['risque_eleve'], b['n'])} %)" \
            if b and b["n"] else "—"
        rows.append([lib,
                     fr(a["n"]) if a else "—", f"{a['part']} %" if a else "—", rea,
                     fr(b["n"]) if b else "champ absent", reb])
        if a and a["n"] and pc(a["risque_eleve"], a["n"]) >= 10:
            fills[(len(rows) - 1, 3)] = RED_PALE
        if b and b["n"] and pc(b["risque_eleve"], b["n"]) >= 10:
            fills[(len(rows) - 1, 5)] = RED_PALE
    table(s, LEFT, y, 4.5, rows, col_w=[30, 12, 11, 19, 12, 16], font_size=6.5,
          row_h=0.28, head_h=0.34, fills=fills)

    dev_pp, dev_pm = spp["_devises"], spm["_devises"]
    cats = list(dict.fromkeys(list(dev_pm)[:6] + list(dev_pp)[:6]))[:6]
    if cats:
        txt(s, 5.1, y - 0.02, 4.5, 0.22, "Devises étrangères rencontrées", size=9,
            bold=True, color=NAVY)
        barres(s, 5.1, y + 0.22, 4.5, 1.40, cats,
               [("Entreprises", [dev_pm.get(c, 0) for c in cats]),
                ("Particuliers", [dev_pp.get(c, 0) for c in cats])],
               couleurs=[GREEN, GREEN_OK], size=6.5, num_fmt='#,##0',
               legende=True, gap=45)

    cpp_x, cpm_x = spp["_croisements"], spm["_croisements"]
    lignes = []
    for k in dict.fromkeys(list(cpp_x) + list(cpm_x)):
        pp_v = f"{fr(cpp_x[k])} PP" if k in cpp_x else "— PP"
        pm_v = f"{fr(cpm_x[k])} PM" if k in cpm_x else "— PM (champ absent)"
        lignes.append(f"{k} : {pp_v} · {pm_v}.")
    encart(s, LEFT, 4.26, 9.2, 1.02, "Croisements entre segments", lignes[:5],
           ton="navy", size=7)


def _segments_completude(prs, R):
    s = page(prs, BANDEAU["IV"], "Complétude comparée des segments", PIED_DE_PAGE)
    spp, spm = R["segments"]["pp"], R["segments"]["pm"]
    cpp, cpm = R["completude"]["pp"], R["completude"]["pm"]

    rows = [["Segment", "Eff. PP", "Complétude PP",
             "Écart PP / portefeuille global", "Eff. PM", "Complétude PM",
             "Écart PM / portefeuille global"]]
    fills = {}
    rows.append(["Portefeuille global", fr(cpp["n"]), f"{cpp['taux_global']} %", "référence",
                 fr(cpm["n"]), f"{cpm['taux_global']} %", "référence"])
    for lib in ("Clients PPE", "Comptes en devise étrangère", "Clients non-résidents"):
        a, b = spp.get(lib), spm.get(lib)
        ta = a["completude"]["taux_global"] if a and a["n"] else None
        tb = b["completude"]["taux_global"] if b and b["n"] else None
        rows.append([
            lib,
            fr(a["n"]) if a else "—", f"{ta} %" if ta is not None else "—",
            f"{ta - cpp['taux_global']:+d} pt" if ta is not None else "—",
            fr(b["n"]) if b else "—", f"{tb} %" if tb is not None else "—",
            f"{tb - cpm['taux_global']:+d} pt" if tb is not None else "—",
        ])
        if ta is not None and ta < cpp["taux_global"]:
            fills[(len(rows) - 1, 3)] = RED_PALE
        if tb is not None and tb < cpm["taux_global"]:
            fills[(len(rows) - 1, 6)] = RED_PALE
    table(s, LEFT, 1.24, 9.2, rows, col_w=[22, 11, 14, 17, 11, 14, 17],
          font_size=7.5, row_h=0.27, head_h=0.44, fills=fills)

    def graphe(x, typo, cle, champs_cle, titre):
        c = R["completude"][cle]
        segs = R["segments"][cle]
        dispo = [(l, segs[l]) for l in ("Clients PPE", "Comptes en devise étrangère",
                                        "Clients non-résidents")
                 if l in segs and segs[l]["n"]]
        court = {"Clients PPE": "PPE",
                 "Comptes en devise étrangère": "Devise étrangère",
                 "Clients non-résidents": "Non-résidents"}
        series = [("Global", [c["par_champ"][f] for f in champs_cle])]
        for l, sg in dispo:
            series.append((court.get(l, l),
                           [sg["completude"]["par_champ"][f] for f in champs_cle]))
        couleurs = [MUTED, NAVY, AMBER, RED][:len(series)]
        txt(s, x, 2.86, 4.5, 0.22, titre, size=9, bold=True, color=NAVY)
        barres(s, x, 3.08, 4.5, 2.14, [label(f, typo) for f in champs_cle], series,
               couleurs=couleurs, maxval=100, size=6.5, legende=True, gap=40)

    pires_pp = sorted(cpp["par_champ"].items(), key=lambda kv: (kv[1] is None, kv[1]))[:4]
    pires_pm = sorted(cpm["par_champ"].items(), key=lambda kv: (kv[1] is None, kv[1]))[:4]
    graphe(LEFT, "PP", "pp", [f for f, _ in pires_pp],
           "Particuliers — champs les plus faibles, par segment (%)")
    graphe(5.1, "PM", "pm", [f for f, _ in pires_pm],
           "Entreprises — champs les plus faibles, par segment (%)")


def _segments_qualite(prs, R):
    s = page(prs, BANDEAU["IV"], "Qualité et anomalies des segments", PIED_DE_PAGE)
    spp, spm = R["segments"]["pp"], R["segments"]["pm"]
    qpp, qpm = R["qualite"]["pp"], R["qualite"]["pm"]

    rows = [["Segment", "Eff. PP", "Qualité PP",
             "Écart PP / portefeuille global", "Eff. PM", "Qualité PM",
             "Écart PM / portefeuille global"]]
    fills = {}
    rows.append(["Portefeuille global", fr(qpp["regles"][0]["total"]) if qpp["regles"] else "—",
                 f"{qpp['taux_qualite']} %", "référence",
                 fr(qpm["regles"][0]["total"]) if qpm["regles"] else "—",
                 f"{qpm['taux_qualite']} %", "référence"])
    for lib in ("Clients PPE", "Comptes en devise étrangère", "Clients non-résidents"):
        a, b = spp.get(lib), spm.get(lib)
        ta = a["qualite"]["taux_qualite"] if a and a["n"] else None
        tb = b["qualite"]["taux_qualite"] if b and b["n"] else None
        rows.append([
            lib,
            fr(a["n"]) if a else "—", f"{ta} %" if ta is not None else "—",
            f"{ta - qpp['taux_qualite']:+.1f} pt" if ta is not None else "—",
            fr(b["n"]) if b else "—", f"{tb} %" if tb is not None else "—",
            f"{tb - qpm['taux_qualite']:+.1f} pt" if tb is not None else "—",
        ])
        if ta is not None and ta < qpp["taux_qualite"]:
            fills[(len(rows) - 1, 3)] = RED_PALE
        if tb is not None and tb < qpm["taux_qualite"]:
            fills[(len(rows) - 1, 6)] = RED_PALE
    table(s, LEFT, 1.24, 9.2, rows, col_w=[22, 11, 14, 17, 11, 14, 17],
          font_size=7.5, row_h=0.27, head_h=0.44, fills=fills)

    def mini(x, titre, seg):
        txt(s, x, 2.86, 2.95, 0.22, titre, size=8.5, bold=True, color=NAVY)
        rows = [["Anomalie", "Clients", "%"]]
        top = [r for r in seg["qualite"]["regles"] if r["anomalies"]][:5] if seg else []
        for r in top:
            rows.append([r["nom"][:34], fr(r["anomalies"]), f"{pc(r['anomalies'], seg['n'])} %"])
        while len(rows) < 4:
            rows.append(["—", "—", "—"])
        table(s, x, 3.08, 2.95, rows, col_w=[58, 21, 21], font_size=6.5,
              row_h=0.235, head_h=0.24)

    trio = [("Non-résidents PP", spp.get("Clients non-résidents")),
            ("Devise étrangère PP", spp.get("Comptes en devise étrangère")),
            ("Devise étrangère PM", spm.get("Comptes en devise étrangère"))]
    for i, (titre, seg) in enumerate(trio):
        mini(0.4 + i * 3.10, titre, seg)

    nr = spp.get("Clients non-résidents")
    lignes = []
    if nr and nr["n"]:
        lignes.append(f"Non-résidents particuliers : qualité {nr['qualite']['taux_qualite']} % "
                      f"contre {qpp['taux_qualite']} % en global.")
    dv = spm.get("Comptes en devise étrangère")
    if dv and dv["n"]:
        lignes.append(f"Entreprises en devise étrangère : complétude "
                      f"{dv['completude']['taux_global']} % et qualité "
                      f"{dv['qualite']['taux_qualite']} %.")
    ppe = spp.get("Clients PPE")
    if ppe and ppe["n"]:
        coh = R["scoring"]["ppe_pp"]
        lignes.append(f"PPE particuliers : {fr(coh['ppe'])} clients, dont "
                      f"{fr(coh['sans_risque_eleve'])} sans classement en risque élevé.")
    encart(s, LEFT, 4.48, 9.2, 0.80, "Constats sur le périmètre sensible",
           lignes[:3] or ["Aucun segment sensible identifié."], ton="amber", size=7.5)


def _risque(prs, R):
    s = page(prs, BANDEAU["V"], "Cartographie du risque client (champ RISQUE)", PIED_DE_PAGE)
    a, b = R["scoring"]["pp"], R["scoring"]["pm"]

    el_pp = a["risque"].get("Risque eleve", 0)
    el_pm = b["risque"].get("Risque eleve", 0)
    for i, (val, lib, nt, col) in enumerate([
        (fr(el_pp), "Particuliers à risque élevé", f"{pc(el_pp, a['n'])} % du portefeuille", RED),
        (fr(el_pm), "Entreprises à risque élevé", f"{pc(el_pm, b['n'])} % du portefeuille", RED),
        (fr(a["risque"].get("Risque moyen faible", 0)), "Particuliers en « moyen faible »",
         f"{pc(a['risque'].get('Risque moyen faible', 0), a['n'])} % — concentration", AMBER),
        (fr(a["risque_non_renseigne"] + b["risque_non_renseigne"]), "Clients sans notation",
         f"{fr(a['risque_non_renseigne'])} PP · {fr(b['risque_non_renseigne'])} PM", NAVY),
    ]):
        kpi(s, 0.4 + i * 2.32, 1.22, 2.24, KPI_H, val, lib, nt, col, vsize=20)

    couleurs = [GREEN_OK, GREEN, AMBER, RED, MUTED]
    cats = [NIVEAUX_LIB[n] for n in NIVEAUX]
    txt(s, 0.5, 2.18, 3.0, 0.22, "Particuliers", size=9, bold=True, color=NAVY,
        align=PP_ALIGN.CENTER)
    anneau(s, 0.4, 2.40, 3.0, 2.55, cats, [a["risque"].get(n, 0) for n in NIVEAUX],
           couleurs, size=6.5, labels=False)
    txt(s, 3.6, 2.18, 3.0, 0.22, "Entreprises", size=9, bold=True, color=NAVY,
        align=PP_ALIGN.CENTER)
    anneau(s, 3.5, 2.40, 3.0, 2.55, cats, [b["risque"].get(n, 0) for n in NIVEAUX],
           couleurs, size=6.5, labels=False)

    txt(s, 6.75, 2.18, 2.85, 0.22, "Effectifs par niveau", size=9, bold=True, color=NAVY)
    rows = [["Niveau", "PP", "% PP", "PM", "% PM"]]
    fills = {}
    for n in NIVEAUX:
        x, y = a["risque"].get(n, 0), b["risque"].get(n, 0)
        if not (x or y):
            continue
        rows.append([NIVEAUX_LIB[n], fr(x), f"{pc(x, a['n'])} %", fr(y), f"{pc(y, b['n'])} %"])
        if n == "Risque eleve":
            fills[(len(rows) - 1, 1)] = RED_PALE
            fills[(len(rows) - 1, 3)] = RED_PALE
    table(s, 6.75, 2.40, 2.85, rows, col_w=[30, 18, 17, 18, 17], font_size=6.5,
          row_h=0.26, head_h=0.26, fills=fills)

    lignes = [f"{pc(a['risque'].get('Risque moyen faible', 0), a['n'])} % des particuliers "
              "relèvent d'un seul niveau : la segmentation discrimine peu."]
    if a["risque_non_renseigne"] + b["risque_non_renseigne"]:
        lignes.append(f"{fr(a['risque_non_renseigne'] + b['risque_non_renseigne'])} clients "
                      "restent sans notation de risque.")
    ybas = 2.40 + 0.26 + 0.26 * (len(rows) - 1) + 0.14
    encart(s, 6.75, ybas, 2.85, max(0.6, 5.28 - ybas), None, lignes, ton="amber", size=7)


def _daterev(prs, R):
    s = page(prs, BANDEAU["V"], "État des échéances de revue scoring (champ DATREV)", PIED_DE_PAGE)
    a, b = R["scoring"]["pp"], R["scoring"]["pm"]

    for i, (val, lib, nt, col) in enumerate([
        (f"{a['part_sans_revision']} %",
         "Particuliers avec date de revue scoring échue ou absente",
         f"{fr(a['sans_revision'])} clients : échue, absente ou invalide", RED),
        (f"{b['part_sans_revision']} %",
         "Entreprises avec date de revue scoring échue ou absente",
         f"{fr(b['sans_revision'])} sociétés", RED),
        (fr(a["statuts"].get("Échue", 0)), "Revues scoring particuliers échues",
         f"{pc(a['statuts'].get('Échue', 0), a['n'])} % du portefeuille", RED),
        (fr(b["statuts"].get("Non renseignée", 0)),
         "Date de revue scoring entreprises absente",
         f"{pc(b['statuts'].get('Non renseignée', 0), b['n'])} % du portefeuille", RED),
    ]):
        kpi(s, 0.4 + i * 2.32, 1.22, 2.24, KPI_H, val, lib, nt, col, vsize=18)

    presents = [st for st in STATUTS if a["statuts"].get(st) or b["statuts"].get(st)]
    txt(s, LEFT, 2.18, 5.6, 0.22, "Répartition des échéances de revue scoring", size=9,
        bold=True, color=NAVY)
    barres(s, LEFT, 2.40, 5.6, 1.88, [p.replace("À échoir ", "") for p in presents],
           [("Particuliers", [a["statuts"].get(p, 0) for p in presents]),
            ("Entreprises", [b["statuts"].get(p, 0) for p in presents])],
           couleurs=[GREEN, AMBER], size=6.5, num_fmt='#,##0', legende=True, gap=45)

    txt(s, 6.20, 2.18, 3.4, 0.22, "Détail des statuts", size=9, bold=True, color=NAVY)
    rows = [["Statut", "PP", "% PP", "PM", "% PM"]]
    fills = {}
    for st in presents:
        x, y = a["statuts"].get(st, 0), b["statuts"].get(st, 0)
        libelle = "Non renseignée *" if st == "Non renseignée" else st
        rows.append([libelle, fr(x), f"{pc(x, a['n'])} %", fr(y), f"{pc(y, b['n'])} %"])
        if st in ("Échue", "Non renseignée", "Format invalide"):
            fills[(len(rows) - 1, 0)] = RED_PALE
    table(s, 6.20, 2.40, 3.4, rows, col_w=[34, 17, 16, 17, 16], font_size=6.5,
          row_h=0.245, head_h=0.26, fills=fills)

    anc = ["< 6 mois", "6 à 12 mois", "1 à 2 ans", "> 2 ans"]
    txt(s, LEFT, 4.44, 5.6, 0.22, "Ancienneté des revues scoring échues", size=9,
        bold=True, color=NAVY)
    rows = [["Retard"] + anc, ["Particuliers"] + [fr(a["anciennete_echues"].get(k, 0)) for k in anc],
            ["Entreprises"] + [fr(b["anciennete_echues"].get(k, 0)) for k in anc]]
    table(s, LEFT, 4.66, 5.6, rows, col_w=[28, 18, 18, 18, 18],
          font_size=7, row_h=0.24, head_h=0.26)
    note(s, 5.44, "* « Non renseignée » inclut des clients dont le scoring est en cours.")


def _zone_critique(prs, R):
    s = page(prs, BANDEAU["V"],
             "Croisement risque et échéance — zone critique", PIED_DE_PAGE)
    a, b = R["scoring"]["pp"], R["scoring"]["pm"]

    def bloc_crois(y, titre, sc):
        txt(s, LEFT, y, 9.2, 0.22, titre, size=9, bold=True, color=NAVY)
        rows = [["Niveau de risque", "Effectif", "Clients avec date de revue échue",
                 "Clients avec date de revue non renseignée",
                 "Total (date de revue scoring échue ou absente)", "% du niveau"]]
        fills = {}
        for niv in NIVEAUX:
            d = sc["croise"].get(niv)
            if not d:
                continue
            tot = sum(d.values())
            ec, nr = d.get("Échue", 0), d.get("Non renseignée", 0)
            ko = ec + nr + d.get("Format invalide", 0)
            rows.append([NIVEAUX_LIB[niv], fr(tot), fr(ec), fr(nr), fr(ko),
                         f"{pc(ko, tot)} %"])
            if pc(ko, tot) >= 50:
                fills[(len(rows) - 1, 5)] = RED_PALE
            elif pc(ko, tot) >= 30:
                fills[(len(rows) - 1, 5)] = AMBER_PALE
        table(s, LEFT, y + 0.22, 9.2, rows, col_w=[20, 12, 17, 19, 21, 11],
              font_size=7, row_h=0.21, head_h=0.46, fills=fills)
        return y + 0.22 + 0.46 + 0.21 * (len(rows) - 1)

    y = bloc_crois(1.24, "Particuliers — état de la revue scoring par niveau de risque", a)
    y = bloc_crois(y + 0.14,
                   "Entreprises — état de la revue scoring par niveau de risque", b)

    ra, rb = a["risque_eleve"], b["risque_eleve"]
    lignes = [
        f"Particuliers : {fr(ra['sans_revision'])} clients à risque élevé sur {fr(ra['total'])} "
        f"({ra['part']} %) ont une date de revue scoring échue ou absente — "
        f"{fr(ra['echue'])} échues, {fr(ra['non_renseignee'])} absentes.",
        f"Entreprises : {fr(rb['sans_revision'])} sociétés sur {fr(rb['total'])} ({rb['part']} %) "
        f"— {fr(rb['echue'])} échues, {fr(rb['non_renseignee'])} absentes.",
        f"Périmètre prioritaire de régularisation : "
        f"{fr(ra['sans_revision'] + rb['sans_revision'])} clients à risque élevé.",
    ]
    ye = y + 0.14
    encart(s, LEFT, ye, 9.2, max(0.62, 5.28 - ye),
           "Zone critique — risque élevé avec date de revue scoring échue ou absente",
           lignes,
           ton="red", size=7)


def _profil_risque_eleve(prs, R):
    """Profil des clients de niveau « risque élevé » : complétude, qualité, exposition."""
    s = page(prs, BANDEAU["V"], "Profil des clients à risque élevé", PIED_DE_PAGE)
    a, b = R["risque_eleve"]["pp"], R["risque_eleve"]["pm"]
    cpp, cpm = R["completude"]["pp"], R["completude"]["pm"]
    qpp, qpm = R["qualite"]["pp"], R["qualite"]["pm"]
    scpp, scpm = R["scoring"]["pp"], R["scoring"]["pm"]

    for i, (val, lib, nt, col) in enumerate([
        (fr(a["n"]), "Particuliers à risque élevé",
         f"{a['part']} % du portefeuille particuliers", RED),
        (fr(b["n"]), "Entreprises à risque élevé",
         f"{b['part']} % du portefeuille entreprises", RED),
        (f"{scpp['risque_eleve']['part']} %",
         "Date de revue scoring échue ou absente (PP)",
         f"{fr(scpp['risque_eleve']['sans_revision'])} clients concernés", RED),
        (f"{scpm['risque_eleve']['part']} %",
         "Date de revue scoring échue ou absente (PM)",
         f"{fr(scpm['risque_eleve']['sans_revision'])} sociétés concernées", RED),
    ]):
        kpi(s, 0.4 + i * 2.32, 1.22, 2.24, KPI_H, val, lib, nt, col, vsize=20)

    y = bloc(s, 2.20, "Complétude et qualité comparées au portefeuille global")
    rows = [["Indicateur", "Particuliers global", "PP risque élevé",
             "Entreprises global", "PM risque élevé"]]
    fills = {}

    def _cmp(glob, loc, suffixe=" %"):
        return (f"{glob}{suffixe}", f"{loc}{suffixe}" if loc is not None else "—")

    lignes_cmp = [
        ("Taux de complétude", cpp["taux_global"],
         a["completude"]["taux_global"] if a["n"] else None,
         cpm["taux_global"], b["completude"]["taux_global"] if b["n"] else None),
        ("Clients intégralement renseignés", cpp["part_clients_complets"],
         a["completude"]["part_clients_complets"] if a["n"] else None,
         cpm["part_clients_complets"],
         b["completude"]["part_clients_complets"] if b["n"] else None),
        ("Taux de qualité", qpp["taux_qualite"],
         a["qualite"]["taux_qualite"] if a["n"] else None,
         qpm["taux_qualite"], b["qualite"]["taux_qualite"] if b["n"] else None),
        ("Date de revue scoring échue ou absente", scpp["part_sans_revision"],
         a["scoring"]["part_sans_revision"] if a["n"] else None,
         scpm["part_sans_revision"],
         b["scoring"]["part_sans_revision"] if b["n"] else None),
    ]
    for lib, gp, rp, gm, rm in lignes_cmp:
        rows.append([lib, f"{gp} %", f"{rp} %" if rp is not None else "—",
                     f"{gm} %", f"{rm} %" if rm is not None else "—"])
        pire = "revue scoring" in lib
        if rp is not None and ((rp > gp) if pire else (rp < gp)):
            fills[(len(rows) - 1, 2)] = RED_PALE
        if rm is not None and ((rm > gm) if pire else (rm < gm)):
            fills[(len(rows) - 1, 4)] = RED_PALE
    table(s, LEFT, y, 4.5, rows, col_w=[36, 16, 16, 16, 16], font_size=7,
          row_h=0.30, head_h=0.36, fills=fills)

    txt(s, 5.1, y - 0.02, 4.5, 0.22, "Anomalies les plus fréquentes sur ce niveau",
        size=9, bold=True, color=NAVY)
    rows2 = [["Règle", "Clients", "% du niveau", "% global"]]
    top = ([r for r in a["qualite"]["regles"] if r["anomalies"]][:4] if a["n"] else [])
    fills2 = {}
    for r in top:
        rows2.append([r["nom"][:38], fr(r["anomalies"]), f"{r['part']} %",
                      f"{r['part_globale']} %" if r.get("part_globale") is not None else "—"])
        if r.get("part_globale") is not None and r["part"] > r["part_globale"]:
            fills2[(len(rows2) - 1, 2)] = RED_PALE
    while len(rows2) < 3:
        rows2.append(["—", "—", "—", "—"])
    table(s, 5.1, y + 0.22, 4.5, rows2, col_w=[48, 17, 18, 17], font_size=6.5,
          row_h=0.26, head_h=0.30, fills=fills2)

    exp = [f"Non-résidents : {fr(a['non_residents'])} PP · {fr(b['non_residents'])} PM.",
           f"Comptes en devise étrangère : {fr(a['devise_etrangere'])} PP · "
           f"{fr(b['devise_etrangere'])} PM."]
    if a["ppe"] is not None:
        exp.append(f"Clients PPE classés à risque élevé : {fr(a['ppe'])} "
                   f"sur {fr(R['scoring']['ppe_pp']['ppe'])} PPE identifiés.")
    encart(s, LEFT, 4.32, 9.2, 0.96,
           "Exposition du niveau « risque élevé » aux segments sensibles", exp,
           ton="red", size=7.5)


def _plan(R):
    """Construit le plan d'action à partir des résultats (valable pour toute filiale)."""
    cpp, cpm = R["completude"]["pp"], R["completude"]["pm"]
    qpp, qpm = R["qualite"]["pp"], R["qualite"]["pm"]
    spp, spm = R["scoring"]["pp"], R["scoring"]["pm"]
    segpp, segpm = R["segments"]["pp"], R["segments"]["pm"]
    vpp, vpm = R["valeurs"]["pp"], R["valeurs"]["pm"]
    plan = []

    ko = spp["risque_eleve"]["sans_revision"] + spm["risque_eleve"]["sans_revision"]
    if ko:
        plan.append(("Clients à risque élevé sans date de revue scoring AML",
                     f"{fr(ko)} clients (PP + PM)",
                     "Campagne de révision périodique", "Critique"))
    ech = spp["statuts"].get("Échue", 0) + spm["statuts"].get("Échue", 0)
    if ech:
        plan.append(("Revues scoring échues sur l'ensemble du portefeuille",
                     f"{fr(ech)} clients",
                     "Plan de rattrapage avec priorisation des clients à risque élevé "
                     "et des retards de plus de 12 mois", "Élevée"))
    nr = spm["statuts"].get("Non renseignée", 0)
    if pc(nr, spm["n"]) >= 20:
        plan.append(("Date de revue scoring absente pour les entreprises",
                     f"{fr(nr)} sociétés ({pc(nr, spm['n'])} %)",
                     "Initialisation automatique de la date de revue lors du scoring AML",
                     "Élevée"))

    for typo, c, cle in (("PP", cpp, "pp"), ("PM", cpm, "pm")):
        faibles = sorted([(f, v) for f, v in c["par_champ"].items()
                          if v is not None and v < 90], key=lambda kv: kv[1])[:2]
        for f, v in faibles:
            plan.append((f"Champ « {label(f, typo)} » peu alimenté "
                         f"({'particuliers' if typo == 'PP' else 'entreprises'})",
                         f"{fr(c['vides_par_champ'][f])} dossiers vides ({v} %)",
                         "Campagne de fiabilisation des données KYC",
                         "Élevée" if v < 60 else "Moyenne"))

    for typo, v, cle in (("PP", vpp, "pp"), ("PM", vpm, "pm")):
        for f, info in v.items():
            if info["part_non_communiquee"] >= 50:
                plan.append((f"« Non communiquée » dominante sur {label(f, typo)} "
                             f"({'particuliers' if typo == 'PP' else 'entreprises'})",
                             f"{fr(info['non_communiquee'])} dossiers "
                             f"({info['part_non_communiquee']} %)",
                             "Retirer la valeur par défaut du référentiel ou la traiter "
                             "comme une non-réponse", "Élevée"))

    for typo, q in (("particuliers", qpp), ("entreprises", qpm)):
        for r in q["regles"][:2]:
            if r["part"] >= 20:
                plan.append((f"Règle « {r['nom']} » ({typo})",
                             f"{fr(r['anomalies'])} anomalies ({r['part']} %)",
                             "Campagne de fiabilisation des données KYC", "Élevée"))
        for doublon in q.get("doublons_regles", []):
            plan.append(("Règles au libellé différent mais aux conditions identiques",
                         " / ".join(doublon),
                         "Fusionner les règles pour supprimer le double comptage", "Moyenne"))

    nrpp = segpp.get("Clients non-résidents")
    if nrpp and nrpp["n"] and nrpp["qualite"]["taux_qualite"] < qpp["taux_qualite"]:
        plan.append(("Non-résidents : complétude et qualité en retrait",
                     f"{fr(nrpp['n'])} PP · {fr(segpm['Clients non-résidents']['n'])} PM",
                     "Revue dossier par dossier du segment", "Moyenne"))
    devpm = segpm.get("Comptes en devise étrangère")
    if devpm and devpm["part"] >= 5:
        plan.append(("Comptes en devise étrangère (entreprises)",
                     f"{fr(devpm['n'])} sociétés ({devpm['part']} %)",
                     "Contrôle renforcé au titre de la réglementation des changes", "Moyenne"))

    for cle, typo in (("pp", "PP"), ("pm", "PM")):
        for f, taux in R["champs_non_alimentes"][cle].items():
            if taux is None:
                plan.append((f"Champ {label(f, typo)} absent de l'export "
                             f"{'particuliers' if cle == 'pp' else 'entreprises'}",
                             "règles associées inévaluables",
                             "Ajouter la colonne à l'export ou désactiver les règles",
                             "Moyenne"))
            elif taux == 0.0:
                plan.append((f"Champ {label(f, typo)} jamais alimenté",
                             "100 % du portefeuille",
                             "Arbitrer : alimenter le champ ou désactiver les règles associées",
                             "Moyenne"))

    ordre = {"Critique": 0, "Élevée": 1, "Moyenne": 2}
    plan.sort(key=lambda p: ordre[p[3]])
    return plan


def _plan_slides(prs, R):
    plan = _plan(R)
    par_page = 11
    pages = [plan[i:i + par_page] for i in range(0, len(plan), par_page)] or [[]]
    for idx, lot in enumerate(pages, 1):
        titre = "Chantiers de fiabilisation proposés"
        if len(pages) > 1:
            titre += f" ({idx}/{len(pages)})"
        s = page(prs, BANDEAU["VI"], titre, PIED_DE_PAGE)
        rows = [["#", "Constat", "Volumétrie", "Action proposée", "Priorité"]]
        fills = {}
        for i, (constat, volum, action, prio) in enumerate(lot):
            rows.append([str((idx - 1) * par_page + i + 1), constat, volum, action, prio])
            fills[(len(rows) - 1, 4)] = (RED_PALE if prio == "Critique"
                                         else AMBER_PALE if prio == "Élevée" else CARD)
        if len(rows) == 1:
            rows.append(["—", "Aucun chantier prioritaire identifié", "—", "—", "—"])
        table(s, LEFT, 1.24, 9.2, rows, col_w=[4, 30, 20, 36, 10], font_size=7,
              row_h=0.335, head_h=0.28, fills=fills,
              align=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT,
                     PP_ALIGN.CENTER])
        if idx == len(pages):
            note(s, 5.06, "Priorité « Critique » : impact direct sur la conformité LBC/FT. "
                          "« Élevée » : fiabilité des indicateurs publiés. "
                          "« Moyenne » : dette de qualité à résorber.")


def construire(R, sortie):
    from datetime import date
    d = date.fromisoformat(R["date_audit"])
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)

    _couverture(prs, R, d)
    _sommaire(prs)
    _perimetre(prs, R, d)
    _synthese(prs, R)
    _completude(prs, R, "PP")
    _completude(prs, R, "PM")
    _agences(prs, R)
    _completude_reelle(prs, R)
    _dispositif(prs, R)
    _anomalies(prs, R, "PP")
    _anomalies(prs, R, "PM")
    _segments_carto(prs, R)
    _segments_completude(prs, R)
    _segments_qualite(prs, R)
    _risque(prs, R)
    _daterev(prs, R)
    _zone_critique(prs, R)
    _profil_risque_eleve(prs, R)
    _plan_slides(prs, R)
    cloture(prs, "FIN DU DOCUMENT",
            f"Audit KYC {R['filiale']['libelle']} — {PIED_DE_PAGE}", _date_fr(d))

    prs.save(sortie)
    return sortie, len(prs.slides._sldIdLst)
