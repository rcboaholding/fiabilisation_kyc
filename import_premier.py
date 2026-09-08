import csv
import os
import re
import sys
import logging
from datetime import datetime
from logging.handlers import RotatingFileHandler
import argparse

import django
from django.db import transaction
from django.db.models import Model

                                                       
                  
                                                       
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "Fiabilisation_kyc.settings")
django.setup()

from kyc.models import  TauxEvolution_filiale, Filiales as FILIALES_CHOICES

                                                       
                       
                                                       

CHEMIN_BASE = os.environ.get(
    "KYC_DATA_DIR",
   r"C:\Users\mamsylla\OneDrive - BANK OF AFRICA(1)\Documents\Projets\2025\Plateforme notatio kyc v2\data",
)
SUIVI_PATTERN = os.environ.get(
    "KYC_SUIVI_PATTERN",
    os.path.join(CHEMIN_BASE, "suivi_fiabilisation_{code}.csv"),
)
TAUX_PATTERN = os.environ.get(
    "KYC_TAUX_PATTERN",
    os.path.join(CHEMIN_BASE, "taux_{code}.csv"),
)

DELIMITEUR = ";"
BULK_SIZE_TAUX = 5000
BULK_SIZE_TAUX_FILIALE = 500
LOG_STEP = 100000

                                                       
            
                                                       
def setup_logging():
    log_dir = os.path.join(os.getcwd(), "logs")
    os.makedirs(log_dir, exist_ok=True)
    logger = logging.getLogger("KYC_Importer")
    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    if not logger.handlers:
        console = logging.StreamHandler(sys.stdout)
        console.setLevel(logging.INFO)
        console.setFormatter(formatter)
        file_handler = RotatingFileHandler(
            os.path.join(log_dir, "import_kyc.log"),
            maxBytes=5 * 1024 * 1024,
            backupCount=5,
            encoding="utf-8",
        )
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(formatter)
        debug_handler = RotatingFileHandler(
            os.path.join(log_dir, "import_premier_debug.log"),
            maxBytes=10 * 1024 * 1024,
            backupCount=5,
            encoding="utf-8",
        )
        debug_handler.setLevel(logging.DEBUG)
        debug_handler.setFormatter(formatter)
        logger.addHandler(console)
        logger.addHandler(file_handler)
        logger.addHandler(debug_handler)
    return logger

logger = setup_logging()

                                                       
          
                                                       
def parse_date_multi(val):
    if not val or str(val).lower() in ("nan", "null", "", "none"):
        return None

    s = str(val).strip()
    if len(s) == 10:
        if s[4] == "-" and s[7] == "-":
            try:
                return datetime(int(s[0:4]), int(s[5:7]), int(s[8:10])).date()
            except ValueError:
                pass
        if s[2] == "/" and s[5] == "/":
            try:
                return datetime(int(s[6:10]), int(s[3:5]), int(s[0:2])).date()
            except ValueError:
                pass

    clean_val = re.sub(r"[^0-9/-]", "", s)
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y", "%Y/%m/%d"):
        try:
            return datetime.strptime(clean_val, fmt).date()
        except ValueError:
            continue
    return None

def parse_taux(val):
    # 'N/A', vide ou non numerique => None (pas 0) pour faire un saut sur la courbe
    if val is None:
        return None
    s = str(val).strip()
    if s == "" or s.lower() in ("n/a", "na", "nan", "null", "none", "-"):
        return None
    try:
        return float(s.replace("%", "").replace(",", ".").strip())
    except ValueError:
        return None

def bulk_create_fallback(rows: list[Model], model: Model):
    inserted = 0
    for obj in rows:
        try:
            obj.save()
            inserted += 1
        except Exception as e:
            pk_val = getattr(obj, "expl", getattr(obj, "filiale", "N/A"))
            logger.error(f"fallback insert failed (key={pk_val}): {e}")
    return inserted

def normalize_header(name):
    return (name or "").strip().upper()

def pick_value(row_norm, candidates):
    for col in candidates:
        val = row_norm.get(col, "")
        if val is not None and str(val).strip() != "":
            return str(val).strip()
    if candidates:
        return (row_norm.get(candidates[0], "") or "").strip()
    return ""

try:
    import chardet                
except Exception:
    chardet = None

def detect_encoding(path):
    if chardet is None:
        return "utf-8-sig"
    try:
        with open(path, "rb") as f:
            raw = f.read(100000)
        enc = chardet.detect(raw).get("encoding")
        return enc or "utf-8-sig"
    except Exception:
        return "utf-8-sig"

def resolve_path(pattern, code):
    return pattern.format(code=code)

                                                       
                        
                                                       
def build_filiales_codes(cli_filiales=None):
                  
    if cli_filiales:
        parts = [p.strip() for p in cli_filiales.replace(";", ",").split(",")]
        return [p for p in parts if p]

                                    
    override = os.environ.get("KYC_FILIALES")
    if override:
        parts = [p.strip() for p in override.replace(";", ",").split(",")]
        return [p for p in parts if p]

                     
    codes = []
    for val, _ in FILIALES_CHOICES:
        if val.startswith("BOA "):
            codes.append(val.replace("BOA ", "").strip())
    return codes

                                                       
                         
                                                       
def import_taux_filiales():
    logger.info("START TAUX_FILIALES")
    for code in FILIALES:
        nom_filiale_complet = f"BOA {code}"
        fichier = resolve_path(SUIVI_PATTERN, code)

        if not os.path.exists(fichier):
            logger.warning(f"missing taux file: {fichier}")
            continue

        TauxEvolution_filiale.objects.filter(filiale=nom_filiale_complet).delete()
        buffer, dates_vues, inserted = [], set(), 0

        encodings = [detect_encoding(fichier), "utf-8-sig", "latin-1", "cp1252"]
        seen = set()
        encodings = [e for e in encodings if not (e in seen or seen.add(e))]

        try:
            with open(fichier, "r", encoding=encodings[0], errors="replace") as f:
                reader = csv.DictReader(f, delimiter=DELIMITEUR)
                orig_fields = reader.fieldnames if reader.fieldnames else []
                norm_fields = [normalize_header(fld) for fld in orig_fields]
                reader.fieldnames = norm_fields
                logger.debug(f"[{code}] taux_filiales encoding={encodings[0]} headers={norm_fields}")

                                                            
                date_key = next((k for k in norm_fields if "DATE" in k), None)
                flux_pm_key = next((k for k in norm_fields if "FLUX" in k and "PM" in k), None)
                flux_pp_key = next((k for k in norm_fields if "FLUX" in k and "PP" in k), None)
                stock_pm_key = next((k for k in norm_fields if "STOCK" in k and "PM" in k), None)
                stock_pp_key = next((k for k in norm_fields if "STOCK" in k and "PP" in k), None)

                if not date_key:
                    logger.error(f"no Date column detected for {code}")
                    continue
                if not any([flux_pm_key, flux_pp_key, stock_pm_key, stock_pp_key]):
                    logger.error(f"no FLUX/STOCK columns detected for {code}")
                    continue

                for row in reader:
                    try:
                        date_str = row.get(date_key)
                        date_obj = parse_date_multi(date_str)
                        if not date_obj or date_obj in dates_vues:
                            continue

                        val_flux_pm = row.get(flux_pm_key) if flux_pm_key else None
                        val_flux_pp = row.get(flux_pp_key) if flux_pp_key else None
                        val_stock_pm = row.get(stock_pm_key) if stock_pm_key else None
                        val_stock_pp = row.get(stock_pp_key) if stock_pp_key else None

                        taux = TauxEvolution_filiale(
                            filiale=nom_filiale_complet,
                            date=date_obj,
                            flux_PM=parse_taux(val_flux_pm),
                            flux_PP=parse_taux(val_flux_pp),
                            stock_PM=parse_taux(val_stock_pm),
                            stock_PP=parse_taux(val_stock_pp),
                        )
                        buffer.append(taux)
                        dates_vues.add(date_obj)

                        if len(buffer) >= BULK_SIZE_TAUX_FILIALE:
                            TauxEvolution_filiale.objects.bulk_create(buffer)
                            inserted += len(buffer)
                            buffer.clear()
                    except Exception as e:
                        logger.error(f"row error: {e}")
                        continue

                if buffer:
                    TauxEvolution_filiale.objects.bulk_create(buffer)
                    inserted += len(buffer)

            logger.info(f"{code}: {inserted} rows imported")
        except Exception as e:
            logger.error(f"{code} critical error: {e}")

    logger.info("END TAUX_FILIALES")

                                                       
                   
                                                       
if __name__ == "__main__":
    try:
        parser = argparse.ArgumentParser(description="KYC import")

        parser.add_argument("--filiales", help="Liste des filiales: SN,CI,BF")
        parser.add_argument("--data-dir", dest="data_dir", help="Base directory for CSV files")
        parser.add_argument("--suivi-pattern", dest="suivi_pattern")
        parser.add_argument("--taux-pattern", dest="taux_pattern")

        args = parser.parse_args()

                                      
                          
                                      
        if args.data_dir:
            globals()["CHEMIN_BASE"] = args.data_dir

        if args.suivi_pattern:
            globals()["SUIVI_PATTERN"] = args.suivi_pattern
        if args.taux_pattern:
            globals()["TAUX_PATTERN"] = args.taux_pattern

        if args.data_dir:
            if not args.suivi_pattern:
                globals()["SUIVI_PATTERN"] = os.path.join(CHEMIN_BASE, "suivi_fiabilisation_{code}.csv")
            if not args.taux_pattern:
                globals()["TAUX_PATTERN"] = os.path.join(CHEMIN_BASE, "taux_{code}.csv")

                                      
                             
                                      
        FILIALES = build_filiales_codes(args.filiales)
        logger.info(f"Filiales utilisées: {FILIALES}")
        logger.info(f"CHEMIN_BASE: {CHEMIN_BASE}")

                                      
                              
                                      
        only_raw = os.environ.get("KYC_ONLY", "").strip()
        only = []
        if only_raw:
            only = [p.strip().lower() for p in only_raw.replace(";", ",").split(",") if p.strip()]

        if not only or "taux_filiales" in only:
            with transaction.atomic():
                import_taux_filiales()

        logger.info("IMPORT DONE")

    except Exception as e:
        logger.critical(f"STOP - FATAL ERROR: {e}")
