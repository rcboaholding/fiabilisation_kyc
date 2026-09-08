# -*- coding: utf-8 -*-
"""Génère le rapport d'audit KYC d'une filiale, en une commande.

    python generate_report.py --filiale SN
    python generate_report.py --filiale CI --dossier D:/exports --json
    python generate_report.py --filiale CI --dossier-csv D:/exports/csv

Entrées attendues dans le dossier de travail :
    pp_XX_STOCK.csv          export particuliers  (produit par Script_V3.r)
    pm_XX_STOCK.csv          export entreprises   (produit par Script_V3.r)
    quality_rules_export.json référentiel des règles qualité de la plateforme
    assets_tpl/tpl_image1.png logo repris du modèle (facultatif)

Sortie :
    Audit_KYC_BOA_XX_AAAAMMJJ.pptx  (+ .json avec --json)
"""
import argparse
import json
import os
import sys
from datetime import date


def main(argv=None):
    p = argparse.ArgumentParser(
        description="Audit KYC : complétude, qualité, segments sensibles, scoring AML.")
    p.add_argument("--filiale", "-f", required=True,
                   help="Code ISO 2 de la filiale (SN, CI, BJ, ML, BF, ...)")
    p.add_argument("--dossier", "-d", default=".",
                   help="Dossier contenant les exports (défaut : dossier courant)")
    p.add_argument("--dossier-csv", dest="dossier_csv", default=None,
                   help="Dossier des exports pp_XX_STOCK.csv / pm_XX_STOCK.csv "
                        "(défaut : --dossier)")
    p.add_argument("--pp", default=None,
                   help="Chemin complet du CSV particuliers (prioritaire sur --dossier-csv)")
    p.add_argument("--pm", default=None,
                   help="Chemin complet du CSV entreprises (prioritaire sur --dossier-csv)")
    p.add_argument("--sortie", "-o", default=None,
                   help="Chemin du PPTX (défaut : Audit_KYC_BOA_XX_AAAAMMJJ.pptx)")
    p.add_argument("--regles", default=None,
                   help="Chemin du référentiel de règles (défaut : quality_rules_export.json)")
    p.add_argument("--date", default=None,
                   help="Date de référence AAAA-MM-JJ (défaut : aujourd'hui). "
                        "Pilote les opérateurs expired / age_gt.")
    p.add_argument("--json", action="store_true",
                   help="Écrire également les résultats bruts en JSON")
    args = p.parse_args(argv)

    sys.stdout.reconfigure(encoding="utf-8")
    from kyc_audit import deck, pipeline

    jour = date.fromisoformat(args.date) if args.date else date.today()
    iso = args.filiale.strip().upper()

    print(f"=== Audit KYC BOA {iso} — référence {jour.isoformat()} ===")
    dossier_csv = args.dossier_csv or args.dossier
    pp_path = args.pp or os.path.join(dossier_csv, f"pp_{iso}_STOCK.csv")
    pm_path = args.pm or os.path.join(dossier_csv, f"pm_{iso}_STOCK.csv")

    R = pipeline.executer(iso, dossier=args.dossier, regles_path=args.regles, today=jour,
                          pp_path=pp_path, pm_path=pm_path)

    sortie = args.sortie or os.path.join(
        args.dossier, f"Audit_KYC_BOA_{iso}_{jour.strftime('%Y%m%d')}.pptx")
    chemin, n = deck.construire(R, sortie)

    if args.json:
        cj = os.path.splitext(chemin)[0] + ".json"
        with open(cj, "w", encoding="utf-8") as f:
            json.dump(R, f, ensure_ascii=False, indent=1)
        print(f"  résultats bruts : {cj}")

    c, q = R["completude"], R["qualite"]
    print(f"  PP : {R['source']['pp']['lignes']:>8} lignes | complétude "
          f"{c['pp']['taux_global']} % | qualité {q['pp']['taux_qualite']} %")
    print(f"  PM : {R['source']['pm']['lignes']:>8} lignes | complétude "
          f"{c['pm']['taux_global']} % | qualité {q['pm']['taux_qualite']} %")
    print(f"=== {n} diapositives -> {chemin} ===")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
