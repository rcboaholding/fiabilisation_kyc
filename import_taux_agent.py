import csv
import os
import sys
import logging
from datetime import datetime
import argparse

import django
from django.db import transaction

                                                       
                                                 
                                                       
                                              
DEFAULT_DATA_DIR = r"C:\Users\mamsylla\OneDrive - BANK OF AFRICA(1)\Documents\Projets\2025\Plateforme notatio kyc v2\data"
                                                        
                                                                           
                                            
DEFAULT_FILIALES = []

                                                                     
DEFAULT_TAUX_FILENAME = "taux_{code}.csv"

                                  
BULK_SIZE = 5000

                                                       
                  
                                                       
                                                         
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "Fiabilisation_kyc.settings")
django.setup()

from kyc.models import TauxEvolution, Filiales as FILIALES_CHOICES

                                                       
            
                                                       
def setup_logging():
    logger = logging.getLogger("KYC_IMPORT")
    logger.setLevel(logging.INFO)

    if not logger.handlers:
        console = logging.StreamHandler(sys.stdout)
        formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        console.setFormatter(formatter)
        logger.addHandler(console)

    return logger

logger = setup_logging()

                                                       
                
                                                       
def normalize_header(name):
    return (name or "").strip().upper().replace("\ufeff", "")

def pick_value(row_norm, candidates):
    for col in candidates:
        val = row_norm.get(col, "")
        if val is not None and str(val).strip() != "":
            return str(val).strip()
    if candidates:
        return (row_norm.get(candidates[0], "") or "").strip()
    return ""

def parse_date_multi(val):
    if not val:
        return None
    val = str(val).strip()
    try:
        if "-" in val:
            return datetime.strptime(val[:10], "%Y-%m-%d").date()
        elif "/" in val:
            return datetime.strptime(val[:10], "%d/%m/%Y").date()
    except:
        return None
    return None

def parse_taux(val):
    # 'N/A', vide ou non numerique => None (pas 0) pour faire un saut sur la courbe
    if val is None:
        return None
    s = str(val).strip()
    if s == "" or s.lower() in ("n/a", "na", "nan", "null", "none", "-"):
        return None
    try:
        return float(s.replace("%", "").replace(",", "."))
    except ValueError:
        return None

def detect_encoding(path):
    try:
        import chardet
        with open(path, "rb") as f:
            raw = f.read(50000)
        return chardet.detect(raw)["encoding"] or "utf-8"
    except:
        return "utf-8"

                                                       
                         
                                                       
def build_filiales_list(args_filiales):
                                                           
    if args_filiales:
        return [f.strip().upper() for f in args_filiales.split(",")]

                                                          
    if DEFAULT_FILIALES:
        return [f.upper() for f in DEFAULT_FILIALES]

                                                                            
    return [
        val.replace("BOA ", "").strip()
        for val, _ in FILIALES_CHOICES
        if val.startswith("BOA ")
    ]

                                                       
                         
                                                       
def importer_taux(filiales, taux_pattern, clear=False):
    logger.info("===== DÉBUT IMPORT TAUX EVOLUTION =====")

    for code in filiales:
        filepath = taux_pattern.format(code=code)
        nom_filiale = f"BOA {code}"

        if not os.path.exists(filepath):
            logger.warning(f"Fichier ignoré (introuvable) : {filepath}")
            continue

                                          
        if clear:
            deleted = TauxEvolution.objects.filter(filiale=nom_filiale).delete()[0]
            logger.info(f"{code} : {deleted} anciennes lignes supprimées")

        encoding = detect_encoding(filepath)
        objects_to_create = []
        inserted_count = 0
        unique_keys = set()

        try:
            with open(filepath, encoding=encoding, errors="replace") as f:
                reader = csv.DictReader(f, delimiter=";")
                if not reader.fieldnames:
                    logger.error(f"{code} : en-têtes vides dans {filepath}")
                    continue
                reader.fieldnames = [normalize_header(h) for h in reader.fieldnames]

                                                
                agent_cols = ["EXPL", "AGENT", "AGENTS", "CODE_EXPL", "CODE_AGENT"]
                date_cols = ["DATE", "DATE_NOTATION", "DATEREV", "DATE_REV", "DATE_REVISION"]
                taux_cols = ["TAUX", "TAUX_AGENT", "SCORE"]
                flux_stock_cols = ["FLUX_STOCK", "FLUX/STOCK", "FLUXSTOCK", "FS"]
                pp_pm_cols = ["PP_PM", "PP/PM", "PPPM"]

                header_set = set(reader.fieldnames)
                if not any(c in header_set for c in agent_cols):
                    logger.error(f"{code} : colonne agent introuvable. En-têtes: {reader.fieldnames}")
                    continue
                if not any(c in header_set for c in date_cols):
                    logger.error(f"{code} : colonne date introuvable. En-têtes: {reader.fieldnames}")
                    continue

                for row in reader:
                    row_norm = {normalize_header(k): (v or "") for k, v in row.items()}
                    agent = pick_value(row_norm, agent_cols)
                    date_val = parse_date_multi(pick_value(row_norm, date_cols))
                    taux_val = parse_taux(pick_value(row_norm, taux_cols))
                    flux_stock = pick_value(row_norm, flux_stock_cols)
                    pp_pm = pick_value(row_norm, pp_pm_cols)

                    if not agent or not date_val:
                        continue

                                                                       
                    key = (agent, date_val, nom_filiale)
                    if key in unique_keys:
                        continue
                    unique_keys.add(key)

                    obj = TauxEvolution(
                        filiale=nom_filiale,
                        agence=None,
                        expl=agent,
                        date=date_val,
                        taux=taux_val,
                        flux_stock=flux_stock,
                        pp_pm=pp_pm,
                    )
                    objects_to_create.append(obj)

                                                   
                    if len(objects_to_create) >= BULK_SIZE:
                        TauxEvolution.objects.bulk_create(objects_to_create)
                        inserted_count += len(objects_to_create)
                        objects_to_create.clear()
                        logger.info(f"{code} en cours... ({inserted_count} lignes)")

                                       
                if objects_to_create:
                    TauxEvolution.objects.bulk_create(objects_to_create)
                    inserted_count += len(objects_to_create)

            logger.info(f"SUCCÈS {code} : {inserted_count} lignes importées au total")

        except Exception as e:
            logger.error(f"ERREUR sur la filiale {code} : {e}")

    logger.info("===== FIN DE L'IMPORT =====")

                                                       
              
                                                       
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Script d'importation des Taux d'Évolution KYC")
    parser.add_argument("--filiales", help="Codes filiales séparés par virgule (ex: SN,CI)")
    parser.add_argument("--data-dir", help="Dossier contenant les fichiers", default=DEFAULT_DATA_DIR)
    parser.add_argument("--clear", action="store_true", help="Vider la table pour les filiales avant l'import")
    args = parser.parse_args()

                                                   
    final_pattern = os.path.join(args.data_dir, DEFAULT_TAUX_FILENAME)

                                          
    target_filiales = build_filiales_list(args.filiales)

                                            
    for code in target_filiales:
        test_path = final_pattern.format(code=code)
        if not os.path.exists(test_path):
            logger.warning(f"Fichier manquant pour {code} : {test_path}")

                             
    print("-" * 50)
    logger.info(f"Dossier source  : {args.data_dir}")
    logger.info(f"Filiales cibles : {target_filiales}")
    logger.info(f"Action 'Clear'  : {'OUI' if args.clear else 'NON'}")
    print("-" * 50)

                           
    importer_taux(
        filiales=target_filiales,
        taux_pattern=final_pattern,
        clear=args.clear
    )
