import csv
import logging
import os
import re
import sys
import time
import random
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path

import django
from django.conf import settings
from django.db import connection, transaction

try:
    import chardet
except Exception:
    chardet = None


os.environ.setdefault("DJANGO_SETTINGS_MODULE", "Fiabilisation_kyc.settings")
django.setup()

from kyc.models import Kyc_pm, Kyc_pp, Filiales as FILIALES_CHOICES


                          
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_DATA_DIR = r"C:\Users\mamsylla\OneDrive - BANK OF AFRICA(1)\Documents\Projets\2025\Plateforme notatio kyc v2\data"



STRICT_FIELD_LIMIT = int(os.environ.get("KYC_FIELD_LIMIT", "131072"))
csv.field_size_limit(STRICT_FIELD_LIMIT)

CHEMIN_BASE = os.environ.get("KYC_DATA_DIR", DEFAULT_DATA_DIR)
DELIMITEUR_CSV = os.environ.get("KYC_DELIMITER", ";")
BULK_SIZE = int(os.environ.get("KYC_BULK_SIZE", "10000"))
LOG_STEP = int(os.environ.get("KYC_LOG_STEP", "100000"))
IMPORT_METHOD = os.environ.get("KYC_IMPORT_METHOD", "auto").strip().lower()


def build_filiales_codes():
                                            
    override = os.environ.get("KYC_FILIALES")
    if override:
        return [part.strip() for part in override.replace(";", ",").split(",") if part.strip()]

                                                                     
    codes = []
    for val, _ in FILIALES_CHOICES:
        if val.startswith("BOA "):
            codes.append(val.replace("BOA ", "").strip())
    return codes


FILIALES = build_filiales_codes()
KYC_ONLY = os.environ.get("KYC_ONLY", "all").strip().lower()


                    
def setup_logging():
    log_dir = BASE_DIR / "logs"
    log_dir.mkdir(exist_ok=True)
    logger = logging.getLogger("KYC_Importer")
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    if not logger.handlers:
        console = logging.StreamHandler(sys.stdout)
        console.setFormatter(formatter)
        logger.addHandler(console)

        file_handler = RotatingFileHandler(
            log_dir / "import_kyc.log",
            maxBytes=5 * 1024 * 1024,
            backupCount=5,
            encoding="utf-8",
        )
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    return logger


logger = setup_logging()


                   
def log_rejected_line(filename, line_num, error_msg):
    reject_file = Path(CHEMIN_BASE) / "lignes_ignorees.txt"
    reject_file.parent.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%H:%M:%S")
    with open(reject_file, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] Fichier: {filename} | Ligne: {line_num} | Erreur: {error_msg}\n")


                        
def normalize_header(name):
    return (name or "").replace("\ufeff", "").strip().upper()


def unique_values(values):
    seen = set()
    unique = []
    for value in values:
        normalized = normalize_header(value)
        if normalized and normalized not in seen:
            seen.add(normalized)
            unique.append(value)
    return unique


FIELD_ALIASES = {
    "INTITULE_COMPTE": ["INTITULE", "INTITULE COMPTE", "INTITULE_COMPTE", "NOM_COMPTE", "NOM COMPTE"],
    "EMPLOYEUR": ["EMPLOYEUR", "EMPLOYER"],
    "LIB_AGENCE": ["AGENCELIB", "AGENCE_LIB", "LIBELLE_AGENCE", "LIB AGENCE"],
    "ADRESSE_SOCIALE": ["ADRESSE_SOCIALE", "ADRESSE SOCIALE", "SIEGE_SOCIAL", "SIEGE SOCIAL"],
    "NUMERO_FISCAL": ["NUMERO_FISCAL", "NUMERO FISCAL", "NIF", "IFU", "TIN"],
    "PAYS_JUR": ["PAYS_JUR", "PAYS JUR", "PAYS_JURIDICTION", "PAYS JURIDICTION"],
    "ACTIONNAIRE": ["ACTIONNAIRE", "ACTIONNAIRES"],
    "MANDATAIRE": ["MANDATAIRE", "MANDATAIRES"],
    "BOITE_POSTALE": ["BOITE_POSTALE", "BOITE POSTALE", "BP", "B.P."],
    "CONSENT_BIC": ["CONSENT_BIC", "CONSENT BIC", "CONSENTEMENT_BIC", "CONSENTEMENT BIC"],
    "LIEU_DELIVRANCE_CIN": ["LIEU_DELIVRANCE_CIN", "LIEU DELIVRANCE CIN", "LIEU_DELIVRANCE", "LIEU DELIVRANCE"],
    "ORIGINE_REV": ["ORIGINE_REVENU", "ORIGINE_REVENUS", "ORIGINE REVENU", "ORIGINE REVENUS"],
    "DATEREV": ["DATREV", "DATEREV"],
    "RISQUE": ["RISQUE", "CLASSE", "CLASSE_RISQUE", "CLASSE RISQUE", "RISK"],
    "PAYS_RESID": ["PAYS_RESIDENCE", "PAYS_RESID", "PAYS RESIDENCE", "PAYS RESID"],
    "NUMID": ["NUM_ID", "NUMERO_ID", "NUMERO_IDENTITE", "NUMERO DOCUMENT", "NUMERO_DOCUMENT"],
    "DATVALID": ["DATE_VALIDITE", "DAT_VALID", "DATE VALIDITE", "DATE_EXPIRATION", "DATE EXPIRATION"],
    "DATNAIS": ["DATE_NAISSANCE", "DAT_NAIS", "DATE NAISSANCE", "DN"],
    "DATOUV": ["DATE_OUVERTURE", "DAT_OUV", "DATE OUVERTURE"],
    "PAYNAIS": ["PAYS_NAISSANCE", "PAYS NAISSANCE", "PAYS_NAIS"],
    "SALAIRE": ["REVENU", "REVENUS", "REVENU_MENSUEL", "SALAIRE_MENSUEL"],
    "TEL": ["TELEPHONE", "PHONE", "MOBILE"],
    "IDP": ["IDP", "IDENTIFIANT_PP", "IDENTIFIANT CLIENT", "ID_CLIENT"],
    "IDM": ["IDM", "IDENTIFIANT_PM", "IDENTIFIANT CLIENT", "ID_CLIENT"],
    "RCSNO": ["RCCM", "RCS", "NUMERO_RCS", "RCS NO"],
}

NUMERIC_TEXT_FIELDS = {"CAPITAL", "CA", "RESULTAT"}

                                                                           
                                                                                  
DATE_TEXT_FIELDS = {"DATEREV"}

_DATE_RE_ISO = re.compile(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})")
_DATE_RE_FR = re.compile(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})")
_DATE_RE_FR_YY = re.compile(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{2})$")


def normalize_date_text(value):
    """Ramène 'dd/mm/yyyy', 'yyyy/mm/dd', 'dd-mm-yyyy', 'dd/mm/yy'… au format ISO
    'YYYY-MM-DD'. Année à 2 chiffres : pivot 50 (29 -> 2029, 87 -> 1987).
    Renvoie la valeur d'origine si le format n'est pas reconnu ou la date invalide."""
    m = _DATE_RE_ISO.match(value)
    if m:
        y, mo, d = m.groups()
    else:
        m = _DATE_RE_FR.match(value)
        if m:
            d, mo, y = m.groups()
        else:
            m = _DATE_RE_FR_YY.match(value)
            if not m:
                return value
            d, mo, y = m.groups()
            y = str((2000 if int(y) < 50 else 1900) + int(y))
    try:
        return date(int(y), int(mo), int(d)).isoformat()
    except ValueError:
        return value


def detect_encoding(path):
    if chardet is None:
        return "utf-8-sig"
    try:
        with open(path, "rb") as f:
            raw = f.read(100000)
        return chardet.detect(raw).get("encoding") or "utf-8-sig"
    except Exception:
        return "utf-8-sig"


def build_model_mapping(model):
    mapping = {}
    for field in get_insert_fields(model):
        name = field.name
        mapping[name] = unique_values([name, name.upper(), *FIELD_ALIASES.get(name, [])])
    return mapping


def get_insert_fields(model):
    return [
        field
        for field in model._meta.concrete_fields
        if not field.primary_key and not field.auto_created
    ]


def build_column_plan(headers, fields, mapping):
    header_positions = {}
    for idx, header in enumerate(headers):
        header_positions.setdefault(normalize_header(header), idx)

    plan = []
    missing = []
    for field in fields:
        max_length = getattr(field, "max_length", None)

        if field.name == "FILIALE":
            plan.append((field.name, None, max_length))
            continue

        candidates = [normalize_header(candidate) for candidate in mapping.get(field.name, [field.name])]
        idx = next((header_positions[candidate] for candidate in candidates if candidate in header_positions), None)
        plan.append((field.name, idx, max_length))
        if idx is None:
            missing.append(field.name)

    return plan, missing


def row_to_values(row, plan, filiale_value):
    values = []
    for field_name, idx, max_length in plan:
        if field_name == "FILIALE":
            value = filiale_value
        elif idx is None or idx >= len(row):
            value = ""
        else:
            value = (row[idx] or "").strip()

        if field_name in NUMERIC_TEXT_FIELDS and value:
            value = value.replace(",", ".").replace(" ", "")

        if field_name in DATE_TEXT_FIELDS and value:
            value = normalize_date_text(value)

        if max_length and len(value) > max_length:
            logger.warning(
                f"Valeur tronquee pour {field_name}: {len(value)} caracteres > {max_length}."
            )
            value = value[:max_length]

        values.append(value)
    return tuple(values)


def quoted_name(name):
    return connection.ops.quote_name(name)


def is_mssql_database():
    engine = settings.DATABASES["default"]["ENGINE"].lower()
    vendor = getattr(connection, "vendor", "").lower()
    return "mssql" in engine or vendor in {"microsoft", "mssql"}


def should_use_direct_insert():
    if IMPORT_METHOD in {"direct", "sql", "fast"}:
        return True
    if IMPORT_METHOD in {"orm", "django"}:
        return False
    return is_mssql_database()


                                                 
DEADLOCK_MAX_RETRIES = int(os.environ.get("KYC_DEADLOCK_RETRIES", "5"))


def is_deadlock_error(exc):
    text = str(exc)
    return "1205" in text or "deadlock" in text.lower()


def run_with_deadlock_retry(func, description):
    """Rejoue func() si SQL Server tue la transaction sur deadlock (1205)."""
    attempt = 0
    while True:
        try:
            return func()
        except Exception as exc:
            if not is_deadlock_error(exc) or attempt >= DEADLOCK_MAX_RETRIES:
                raise
            attempt += 1
                                                                             
            wait = min(2 ** attempt, 30) * 0.5 + random.uniform(0, 0.5)
            logger.warning(
                f"Deadlock sur {description} (tentative {attempt}/{DEADLOCK_MAX_RETRIES}), "
                f"nouvelle tentative dans {wait:.1f}s."
            )
            time.sleep(wait)


                      
def build_insert_sql(model, columns):
    table = quoted_name(model._meta.db_table)
    column_sql = ", ".join(quoted_name(column) for column in columns)
    placeholders = ", ".join(["%s"] * len(columns))
    return f"INSERT INTO {table} ({column_sql}) VALUES ({placeholders})"


_FAST_EXECUTEMANY_LOGGED = False


def enable_fast_executemany(cursor):
    """Active fast_executemany sur le VRAI curseur pyodbc.

    Les wrappers Django puis mssql-django *lisent* fast_executemany depuis
    pyodbc via __getattr__ (donc hasattr renvoie True) mais ne proxifient pas
    __setattr__ et ne repropagent pas le flag : le poser sur un wrapper est
    inerte et pyodbc reste a 1 aller-retour reseau par ligne. Il faut
    descendre jusqu'au curseur pyodbc lui-meme (.cursor.cursor).
    """
    global _FAST_EXECUTEMANY_LOGGED
    if not is_mssql_database():
        return False

    inner = cursor
    while getattr(inner, "cursor", None) is not None:
        inner = inner.cursor

    try:
        inner.fast_executemany = True
    except AttributeError:
        if not _FAST_EXECUTEMANY_LOGGED:
            logger.warning(
                "fast_executemany INTROUVABLE (%s) : insertions ligne par ligne.",
                type(inner).__name__)
            _FAST_EXECUTEMANY_LOGGED = True
        return False

    if not _FAST_EXECUTEMANY_LOGGED:
        logger.info("fast_executemany ACTIF sur le curseur pyodbc (%s).", type(inner).__name__)
        _FAST_EXECUTEMANY_LOGGED = True
    return True


def direct_insert_batch(model, columns, rows_with_lines, filename):
    if not rows_with_lines:
        return 0

    sql = build_insert_sql(model, columns)
    rows = [values for _, values in rows_with_lines]

    def _insert_all():
        with transaction.atomic():
            with connection.cursor() as cursor:
                enable_fast_executemany(cursor)
                cursor.executemany(sql, rows)
        return len(rows)

    try:
        return run_with_deadlock_retry(_insert_all, f"INSERT {model._meta.db_table} ({len(rows)} lignes)")
    except Exception as batch_error:
        logger.error(f"Insertion batch refusee ({len(rows)} lignes) dans {model._meta.db_table}: {batch_error}")

    inserted = 0
    for line_num, values in rows_with_lines:
        try:
            with transaction.atomic():
                with connection.cursor() as cursor:
                    cursor.execute(sql, values)
            inserted += 1
        except Exception as row_error:
            log_rejected_line(filename, line_num, f"Insertion SQL: {row_error}")

    return inserted


def orm_insert_batch(model, fields, rows_with_lines, filename):
    if not rows_with_lines:
        return 0

    field_names = [field.name for field in fields]
    prepared = []
    for line_num, values in rows_with_lines:
        try:
            prepared.append((line_num, model(**dict(zip(field_names, values)))))
        except Exception as obj_error:
            log_rejected_line(filename, line_num, f"Preparation ORM: {obj_error}")

    if not prepared:
        return 0

    objects = [obj for _, obj in prepared]
    try:
        model.objects.bulk_create(objects, batch_size=BULK_SIZE)
        return len(objects)
    except Exception as batch_error:
        logger.error(f"bulk_create refuse ({len(objects)} lignes) dans {model._meta.db_table}: {batch_error}")

    inserted = 0
    for line_num, obj in prepared:
        try:
            obj.save(force_insert=True)
            inserted += 1
        except Exception as row_error:
            log_rejected_line(filename, line_num, f"Insertion ORM: {row_error}")

    return inserted


def insert_batch(model, fields, rows_with_lines, filename, direct_mode):
    columns = [field.name for field in fields]
    if direct_mode:
        return direct_insert_batch(model, columns, rows_with_lines, filename)
    return orm_insert_batch(model, fields, rows_with_lines, filename)


def delete_filiale_rows(model, filiale_value):
    table = quoted_name(model._meta.db_table)
    column = quoted_name("FILIALE")
    sql = f"DELETE FROM {table} WHERE {column} = %s"

    def _delete():
        with transaction.atomic():
            with connection.cursor() as cursor:
                cursor.execute(sql, [filiale_value])
                return cursor.rowcount

    return run_with_deadlock_retry(_delete, f"DELETE {model._meta.db_table} ({filiale_value})")


# Desactivation des index secondaires pendant le chargement : SQL Server
# maintient chaque index a chaque INSERT. On les eteint le temps du bulk load
# puis on les reconstruit en une passe. L'index de scope (1re colonne FILIALE)
# reste actif pour que le DELETE par filiale ne parte pas en scan complet.
TOGGLE_INDEXES = os.environ.get("KYC_TOGGLE_INDEXES", "1") == "1"
_INDEX_TABLES = tuple(sorted({Kyc_pp._meta.db_table, Kyc_pm._meta.db_table}))
_SECONDARY_INDEX_SQL = """
    SELECT t.name, i.name
    FROM sys.indexes i
    JOIN sys.tables t ON t.object_id = i.object_id
    WHERE t.name IN (%s, %s)
      AND i.type_desc = 'NONCLUSTERED'
      AND i.is_primary_key = 0
      AND i.is_unique_constraint = 0
      AND i.name IS NOT NULL
      AND NOT EXISTS (
          SELECT 1 FROM sys.index_columns ic
          JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
          WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id
            AND ic.key_ordinal = 1 AND c.name = 'FILIALE'
      )
"""


def toggle_secondary_indexes(disable):
    """DISABLE (avant import) ou REBUILD (apres) les index secondaires MSSQL."""
    if not TOGGLE_INDEXES or not is_mssql_database():
        return
    action = "DISABLE" if disable else "REBUILD"
    try:
        with connection.cursor() as cursor:
            cursor.execute(_SECONDARY_INDEX_SQL, list(_INDEX_TABLES))
            targets = cursor.fetchall()
            for table, index in targets:
                cursor.execute(
                    f"ALTER INDEX {quoted_name(index)} ON {quoted_name(table)} {action}"
                )
        logger.info("Index secondaires : %s sur %d index.", action, len(targets))
    except Exception as exc:
        logger.warning("Bascule des index (%s) impossible : %s", action, exc)



def importer_csv_optimise(path, model, mapping, code_filiale):
    inserted = 0
    read_rows = 0
    filename = os.path.basename(path)
    filiale_value = f"BOA {code_filiale}"
    direct_mode = should_use_direct_insert()
    fields = get_insert_fields(model)

    logger.info(f"--- Lecture de {filename} vers {model._meta.db_table} ({'SQL direct' if direct_mode else 'ORM'}) ---")

    started_at = time.time()
    enc = detect_encoding(path)
    pending_rows = []

    try:
        with open(path, newline="", encoding=enc, errors="replace") as f:
            reader = csv.reader(f, delimiter=DELIMITEUR_CSV)
            headers = next(reader, None)
            if not headers:
                logger.warning(f"{filename}: fichier vide.")
                return 0

            plan, missing = build_column_plan(headers, fields, mapping)
            if missing:
                logger.warning(f"{filename}: colonnes absentes ou non reconnues: {', '.join(missing)}")

            for line_counter, row in enumerate(reader, start=2):
                try:
                    pending_rows.append((line_counter, row_to_values(row, plan, filiale_value)))
                    read_rows += 1

                    if len(pending_rows) >= BULK_SIZE:
                        inserted += insert_batch(model, fields, pending_rows, filename, direct_mode)
                        pending_rows = []
                        if inserted and inserted % LOG_STEP == 0:
                            logger.info(f"    -> {inserted} lignes inserees...")

                except Exception as row_error:
                    log_rejected_line(filename, line_counter, f"Lecture CSV: {row_error}")

            if pending_rows:
                inserted += insert_batch(model, fields, pending_rows, filename, direct_mode)

    except Exception as file_error:
        logger.critical(f"Erreur fatale sur {filename}: {file_error}")
        return inserted
    finally:
        connection.close_if_unusable_or_obsolete()

    duree = max(time.time() - started_at, 0.001)
    logger.info(f"{filename}: {inserted}/{read_rows} lignes inserees "
                f"en {duree:.0f}s ({inserted / duree:.0f} lignes/s).")
    return inserted


                     
MAPPING_PM = build_model_mapping(Kyc_pm)
MAPPING_PP = build_model_mapping(Kyc_pp)


                                   
def should_import(kind):
    if KYC_ONLY in {"", "all", "both", "*"}:
        return True
    normalized = KYC_ONLY.replace(" ", "").replace("-", "_")
    return kind in {part.strip().lower() for part in normalized.replace(";", ",").split(",")}


def resolve_import_path(code, kind):
    env_key = f"KYC_{kind.upper()}_PATTERN"
    pattern = os.environ.get(env_key)
    if pattern:
        return pattern.format(code=code)

    candidates = [
        os.path.join(CHEMIN_BASE, f"{kind}_{code}_STOCK.csv"),
        os.path.join(CHEMIN_BASE, f"{kind}_{code}_STOCK_F.csv"),
    ]
    for candidate in candidates:
        if os.path.exists(candidate):
            return candidate
    return candidates[0]


def traiter_fichier_si_present(code, kind, model, mapping):
    path = resolve_import_path(code, kind)
    if not os.path.exists(path):
        logger.warning(f"Fichier absent, import ignore: {path}")
        return 0

    filiale_value = f"BOA {code}"
    deleted = delete_filiale_rows(model, filiale_value)
    logger.info(f"{model._meta.db_table}: {deleted} anciennes lignes supprimees pour {filiale_value}.")
    return importer_csv_optimise(path, model, mapping, code)


def traiter_filiale(code):
    """PM et PP d'une meme filiale ne partagent aucune donnee : le goulot
    d'etranglement etant le reseau/MSSQL (pas le CPU), les importer sur deux
    threads distincts (chacun avec sa propre connexion Django thread-locale)
    permet de les chevaucher au lieu de les serialiser."""
    logger.info(f"\n>>> FILIALE {code}")

    jobs = []
    if should_import("pm"):
        jobs.append(("pm", Kyc_pm, MAPPING_PM))
    if should_import("pp"):
        jobs.append(("pp", Kyc_pp, MAPPING_PP))

    total = 0
    if len(jobs) > 1:
        with ThreadPoolExecutor(max_workers=len(jobs)) as executor:
            futures = {
                executor.submit(traiter_fichier_si_present, code, kind, model, mapping): kind
                for kind, model, mapping in jobs
            }
            for future in as_completed(futures):
                kind = futures[future]
                try:
                    total += future.result()
                except Exception as exc:
                    logger.error(f"Erreur {kind.upper()} filiale {code}: {exc}")
    else:
        for kind, model, mapping in jobs:
            total += traiter_fichier_si_present(code, kind, model, mapping)

    logger.info(f"FILIALE {code}: {total} lignes inserees au total.")
    return total


                 
if __name__ == "__main__":
    Path(CHEMIN_BASE).mkdir(parents=True, exist_ok=True)
    with open(Path(CHEMIN_BASE) / "lignes_ignorees.txt", "w", encoding="utf-8") as f:
        f.write(f"--- REJETS DU {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n")

    logger.info(
        "Demarrage import KYC | filiales=%s | data_dir=%s | bulk_size=%s | only=%s | method=%s",
        ",".join(FILIALES),
        CHEMIN_BASE,
        BULK_SIZE,
        KYC_ONLY,
        IMPORT_METHOD,
    )

                                                                           
                                                                           
                                                                         
                                                                             
    toggle_secondary_indexes(disable=True)
    try:
        if os.environ.get("KYC_PARALLEL", "0") == "1" and len(FILIALES) > 1:
            # Threads, pas processus : le goulot est reseau/MSSQL (pas CPU), et
            # ProcessPoolExecutor plante sous tache planifiee Windows sans stdio
            # valide ("OSError: [WinError 6] The handle is invalid" au spawn).
            # Chaque thread obtient sa propre connexion Django thread-locale.
            workers = min(len(FILIALES), int(os.environ.get("KYC_MAX_WORKERS", "6")))
            with ThreadPoolExecutor(max_workers=workers) as executor:
                futures = {executor.submit(traiter_filiale, code): code for code in FILIALES}
                for future in as_completed(futures):
                    code = futures[future]
                    try:
                        future.result()
                    except Exception as exc:
                        logger.error(f"Erreur filiale {code}: {exc}")
        else:
            for code in FILIALES:
                try:
                    traiter_filiale(code)
                except Exception as exc:
                    logger.error(f"Erreur filiale {code}: {exc}")
    finally:
        toggle_secondary_indexes(disable=False)

    logger.info("\nProcessus termine.")
