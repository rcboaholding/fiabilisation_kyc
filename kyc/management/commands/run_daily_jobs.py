"""
Tâche planifiée UNIQUE regroupant tous les traitements quotidiens :
  0a. Worker OCR documents (process_document_ocr --loop, processus détaché)
  0. Imports (import_kyc.py, import_premier.py, import_taux_agent.py)
  1. Calcul des taux de qualité complets (compute_quality_rates)
  2. Préchauffe des caches (warm_ui_caches)
  3. Calcul des appréciations globales (compute_appreciation_globale)
  4. Envoi des rappels DATEREV (filiales payées uniquement)

À la fin, envoie un rapport d'exécution (OK / ERREURS) par email à la liste
de supervision (champ « Emails de supervision » de la Configuration Rappel DATEREV),
via la même configuration SMTP. Le rapport comporte :
  - le tableau des traitements globaux,
  - le détail des étapes de Script_V3.r **par filiale** (lu dans le journal CSV
    que le script alimente, cf. `kyc_journal` dans Script_V3.r),
  - le fichier log complet de l'exécution en pièce jointe.

La fiabilisation R (Script_V3.r) N'EST PLUS lancée ici : elle est exécutée en
amont par la tâche planifiée (app.txt, étape 1) depuis « C:/Fiabilisation KYC/R ».
Cette commande se contente de relire le journal CSV produit par cette exécution
pour en reporter le détail par filiale (voir --r-journal).

Planifier UNE seule tâche, APRÈS le script R :
    python manage.py run_daily_jobs
"""
import csv
import io
import os
import time
import traceback

from django.core.management.base import BaseCommand, CommandError
from django.utils import timezone



JOURNAL_HEADER = ["horodatage", "filiale", "etape", "statut", "duree_s", "message"]



MAX_ATTACHMENT_BYTES = 5 * 1024 * 1024


class Command(BaseCommand):
    help = "Exécute tous les traitements quotidiens et envoie un rapport par email."

    def add_arguments(self, parser):
        parser.add_argument("--no-mail", action="store_true",
                            help="N'envoie pas le rapport par email (exécution seule).")
        parser.add_argument("--skip", default="",
                            help="Jobs à ignorer, séparés par des virgules : "
                                 "ocr_worker,import_kyc,import_premier,import_taux,"
                                 "quality,warm,appreciation,daterev")
        parser.add_argument("--script-timeout", type=int, default=7200,
                            help="Délai max (secondes) par script d'import (défaut 7200).")
        parser.add_argument("--r-journal", default="",
                            help="Chemin du journal d'étapes de Script_V3.r produit par "
                                 "l'exécution externe du script R. Vide = "
                                 "<BASE_DIR>/logs/journal_script_v3.csv.")
        parser.add_argument("--parallel-workers", type=int, default=6,
                            help="Nombre de workers parallèles pour warm_ui_caches et "
                                 "compute_appreciation_globale (défaut 6). Chaque worker est un "
                                 "processus Django complet : en cas de MemoryError, réduire à 2-3 "
                                 "(les slices en échec sont de toute façon repris en séquentiel).")
        parser.add_argument("--daterev-filiale", default="",
                            help="Limite l'envoi des rappels DATEREV à cette filiale (ex. « BOA SN »). "
                                 "Vide = toutes les filiales payées.")

    def handle(self, *args, **options):
        import sys as _sys

        from kyc.models import EmailReminderConfig

                                                                                
                                                                             
        for _stream in (_sys.stdout, _sys.stderr):
            try:
                _stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass

        skip = {s.strip().lower() for s in (options.get("skip") or "").split(",") if s.strip()}
        results = []
        started = timezone.localtime()

        import subprocess
        from django.conf import settings



        logs_dir = os.path.join(str(settings.BASE_DIR), "logs")
        os.makedirs(logs_dir, exist_ok=True)
        self.log_path = os.path.join(
            logs_dir, f"run_daily_jobs_{started.strftime('%Y%m%d_%H%M%S')}.log")
        self.journal_path = (options.get("r_journal") or "").strip() or \
            os.path.join(logs_dir, "journal_script_v3.csv")

        self._log(f"=== Traitements quotidiens — démarrage {started.strftime('%d/%m/%Y %H:%M:%S')} ===")
        self._log(f"Journal Script_V3 : {self.journal_path}")

        def run(key, label, fn):
            if key in skip:
                results.append({"label": label, "status": "ignoré", "ok": True, "detail": "", "duration": 0})
                self.stdout.write(f"[SKIP] {label}")
                self._log(f"[SKIP] {label}")
                return
            t0 = time.time()
            self.stdout.write(f"[START] {label}")
            self._log(f"[START] {label}")
            try:
                detail = fn() or ""
                dur = round(time.time() - t0, 1)
                results.append({"label": label, "status": "OK", "ok": True, "detail": str(detail), "duration": dur})
                self.stdout.write(self.style.SUCCESS(f"[OK] {label} ({dur}s) {detail}"))
                self._log(f"[OK] {label} ({dur}s) {detail}")
            except Exception as e:
                dur = round(time.time() - t0, 1)
                tb = traceback.format_exc()
                results.append({"label": label, "status": "ÉCHEC", "ok": False,
                                "detail": f"{type(e).__name__}: {e}", "duration": dur, "trace": tb[-1500:]})
                self.stderr.write(f"[ÉCHEC] {label} ({dur}s) : {e}")
                self._log(f"[ÉCHEC] {label} ({dur}s) : {e}\n{tb}")


        import sys as _sys

        script_timeout = options.get("script_timeout") or 7200

        def _run_script(filename, extra_env=None):
            path = os.path.join(str(settings.BASE_DIR), filename)
            if not os.path.exists(path):
                raise FileNotFoundError(f"Script introuvable : {filename}")
            script_env = dict(os.environ)
            if extra_env:
                script_env.update(extra_env)
            proc = subprocess.run([_sys.executable, path], cwd=str(settings.BASE_DIR),
                                  capture_output=True, text=True, encoding="utf-8",
                                  errors="replace", timeout=script_timeout, env=script_env)
            out = ((proc.stdout or "") + "\n" + (proc.stderr or "")).strip()
            self._log(f"----- Sortie de {filename} -----\n{out}\n----- fin {filename} -----")
            tail = out.splitlines()[-1].strip() if out else ""
            if proc.returncode != 0:
                raise RuntimeError(f"code retour {proc.returncode} — {tail[:200]}")
            return (tail[:200] or "Terminé.")

                                                                                
                                                                                    
                                                                                  
                                                                                      
        def _ocr_worker():
            pid_file = os.path.join(str(settings.BASE_DIR), "ocr_worker.pid")
            if os.path.exists(pid_file):
                try:
                    pid = int(open(pid_file).read().strip())
                    check = subprocess.run(["tasklist", "/FI", f"PID eq {pid}"],
                                           capture_output=True, timeout=30)
                    if str(pid).encode() in (check.stdout or b""):
                        return f"Worker OCR déjà actif (PID {pid}) — rien à faire."
                except Exception:
                    pass                                                        
            log_handle = open(os.path.join(str(settings.BASE_DIR), "ocr_worker.log"), "ab")
            creationflags = 0
            if os.name == "nt":
                creationflags = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
            proc = subprocess.Popen(
                [_sys.executable, "-X", "utf8", "manage.py", "process_document_ocr",
                 "--loop", "--interval", "5", "--workers", "3"],
                cwd=str(settings.BASE_DIR), stdout=log_handle, stderr=log_handle,
                stdin=subprocess.DEVNULL, creationflags=creationflags, close_fds=True,
            )
            with open(pid_file, "w") as fh:
                fh.write(str(proc.pid))
            return f"Worker OCR lancé en arrière-plan (PID {proc.pid}, log : ocr_worker.log)."
        run("ocr_worker", "Worker OCR documents (process_document_ocr --loop)", _ocr_worker)






        run("import_premier", "Importation évolution filiales (import_premier.py)", lambda: _run_script("import_premier.py"))
        run("import_taux", "Importation taux agents (import_taux_agent.py)", lambda: _run_script("import_taux_agent.py"))
        run("import_kyc", "Importation KYC (import_kyc.py)", lambda: _run_script(
            "import_kyc.py",
            extra_env={"KYC_PARALLEL": "1", "KYC_BULK_SIZE": "30000"},
        ))

                                                                                
                                                                            
                                                                                
                                                                            
                                                                              
                                                                         
        parallel_workers = max(1, options.get("parallel_workers") or 6)

                                                                               
                                                                            
        child_env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}

        def _run_sliced(command, extra_args=()):
            # Jobs lies a MSSQL (agregats sur 1,1 M lignes, aller-retours reseau),
            # pas au CPU : on ne plafonne pas sur os.cpu_count(). Sur la machine de
            # prod (4 vCPU) l'ancien "cpu_count() - 2" bloquait a 2 slices.
            workers = max(1, int(os.environ.get("KYC_SLICE_WORKERS", str(parallel_workers))))

            def _spawn(slice_str=None):
                argv = [_sys.executable, "manage.py", command, *extra_args]
                if slice_str:
                    argv += ["--slice", slice_str]
                return subprocess.Popen(
                    argv,
                    cwd=str(settings.BASE_DIR),
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding="utf-8", errors="replace", env=child_env,
                )

            def _wait(proc, label):
                try:
                    out, _ = proc.communicate(timeout=script_timeout)
                except subprocess.TimeoutExpired:
                    proc.kill()
                    proc.communicate()
                    return False, f"{label} (timeout)", ""
                tail = out.strip().splitlines()[-1].strip() if out and out.strip() else ""
                if proc.returncode != 0:
                    return False, f"{label} (code {proc.returncode} — {tail[:120]})", ""
                return True, "", tail

                                                                            
                                                                             
                                                                             
                                                                      
            help_proc = subprocess.run(
                [_sys.executable, "manage.py", command, "--help"],
                cwd=str(settings.BASE_DIR), capture_output=True,
                text=True, encoding="utf-8", errors="replace", timeout=300,
                env=child_env,
            )
            if "--slice" not in (help_proc.stdout or ""):
                self.stdout.write(f"  {command} ne supporte pas --slice : exécution simple sans parallélisme.")
                ok, err, tail = _wait(_spawn(), command)
                if not ok:
                    raise RuntimeError(err)
                return f"exécution simple (--slice non supporté). {tail[:150]}"

            procs = {i: _spawn(f"{i}/{workers}") for i in range(1, workers + 1)}
            failed_slices, last_line = [], ""
            for i, proc in procs.items():
                ok, err, tail = _wait(proc, f"w{i}")
                if ok:
                    last_line = tail or last_line
                else:
                    failed_slices.append((i, err))

                                                                              
                                                 
            retried_errors = []
            for i, first_err in failed_slices:
                self.stdout.write(f"  Reprise séquentielle du slice {i}/{workers} après : {first_err}")
                ok, err, tail = _wait(_spawn(f"{i}/{workers}"), f"w{i}-retry")
                if ok:
                    last_line = tail or last_line
                else:
                    retried_errors.append(err)
            if retried_errors:
                raise RuntimeError(
                    f"{len(retried_errors)}/{workers} worker(s) en échec après reprise : "
                    f"{'; '.join(retried_errors)}")

            note = f", {len(failed_slices)} slice(s) repris en séquentiel" if failed_slices else ""
            return f"{workers} workers parallèles{note}. {last_line[:150]}"

                                                                                
                                                                               
                                                                                
                                                                          
        run("quality", "Calcul des taux de qualité (compute_quality_rates, parallèle)",
            lambda: _run_sliced("compute_quality_rates"))

                                                                                
                                                                         
                                                                            
        run("warm", "Préchauffe des caches (warm_ui_caches, parallèle)",
            lambda: _run_sliced("warm_ui_caches", ("--skip-snapshot",)))

                                                                                
        run("appreciation", "Calcul des appréciations globales (parallèle)",
            lambda: _run_sliced("compute_appreciation_globale"))

                                                                                
        daterev_filiale = (options.get("daterev_filiale") or "").strip()

        def _daterev():
            from kyc.daterev_mailer import send_daterev_reminders_core
            cfg = EmailReminderConfig.objects.filter(active=True).order_by("-updated_at").first()
            if not cfg:
                return "Aucune configuration SMTP active — envoi ignoré."
                                                                                
                                                                                         
            today = timezone.localdate()
            freq = (cfg.frequency or "manual").lower()
            if freq == "manual":
                return "Fréquence « Manuel » — envoi automatique ignoré."
            if freq == "weekly" and today.weekday() != 0:
                return "Fréquence « Hebdomadaire » — envoi uniquement le lundi, ignoré aujourd'hui."
            if freq == "monthly" and today.day != 1:
                return "Fréquence « Mensuel » — envoi uniquement le 1er du mois, ignoré aujourd'hui."
            sent, skipped = send_daterev_reminders_core(cfg, filiale=daterev_filiale or None, only_paid=True)
            cible = f" pour la filiale « {daterev_filiale} »" if daterev_filiale else ""
            return f"{sent} mail(s) envoyé(s), {skipped} ignoré(s) (email manquant){cible}."
        label_daterev = (f"Envoi des rappels DATEREV — filiale « {daterev_filiale} »"
                         if daterev_filiale else "Envoi des rappels DATEREV (filiales payées)")
        run("daterev", label_daterev, _daterev)

        ended = timezone.localtime()
        all_ok = all(r["ok"] for r in results)
        self.stdout.write(self.style.SUCCESS("=== Traitements quotidiens terminés ===")
                          if all_ok else self.style.ERROR("=== Traitements quotidiens : DES ERREURS ==="))
        self._log(f"=== Terminé {ended.strftime('%d/%m/%Y %H:%M:%S')} — "
                  f"{'OK' if all_ok else 'ERREURS'} ===")



        par_filiale = self._read_journal()

        if options.get("no_mail"):
            self.stdout.write(f"Log de l'exécution : {self.log_path}")
        else:
            try:
                self._send_report(results, started, ended, all_ok, par_filiale)
            except Exception as e:
                self.stderr.write(f"Impossible d'envoyer le rapport de supervision : {e}")



        if not all_ok:
            echecs = ", ".join(r["label"] for r in results if not r["ok"])
            raise CommandError(f"Traitements en échec : {echecs}")




    def _log(self, message):
        """Ajoute une ligne horodatée au fichier log de l'exécution (pièce jointe du rapport)."""
        path = getattr(self, "log_path", None)
        if not path:
            return
        stamp = timezone.localtime().strftime("%H:%M:%S")
        try:
            with open(path, "a", encoding="utf-8", errors="replace") as fh:
                fh.write(f"{stamp} {message}\n")
        except OSError:
            pass

    def _log_tail_since(self, offset, max_chars=4000):
        """Renvoie la fin du log écrite depuis `offset` octets (sortie d'un sous-processus)."""
        try:
            with open(self.log_path, "rb") as fh:
                fh.seek(offset)
                data = fh.read()
        except OSError:
            return ""
        texte = data.decode("utf-8", errors="replace").strip()
        lignes = [l.strip() for l in texte.splitlines() if l.strip()]
        return " | ".join(lignes[-6:])[:max_chars]

    def _read_journal(self):
        """Lit le journal d'étapes de Script_V3.r et le regroupe par filiale.

        Retourne un dict {filiale: [ {etape, statut, duree, message, horodatage}, ... ]}
        dans l'ordre de traitement. Dict vide si le journal est absent (script ignoré
        ou jamais exécuté)."""
        path = getattr(self, "journal_path", "")
        if not path or not os.path.exists(path):
            return {}
        par_filiale = {}
        try:
            with io.open(path, "r", encoding="utf-8", errors="replace", newline="") as fh:
                for ligne in csv.DictReader(fh, delimiter=";"):
                    fil = (ligne.get("filiale") or "").strip()
                    etape = (ligne.get("etape") or "").strip()
                    if not fil or not etape:
                        continue
                    statut = (ligne.get("statut") or "").strip().upper()
                    if statut == "DEBUT":
                        par_filiale.setdefault(fil, [])
                        continue
                    par_filiale.setdefault(fil, []).append({
                        "horodatage": (ligne.get("horodatage") or "").strip(),
                        "etape": etape,
                        "statut": statut,
                        "duree": (ligne.get("duree_s") or "").strip(),
                        "message": (ligne.get("message") or "").strip(),
                    })
        except OSError as e:
            self.stderr.write(f"Journal Script_V3 illisible : {e}")
            return {}
        return par_filiale

    def _attachment(self):
        """Chemin du log à joindre. Au-delà de MAX_ATTACHMENT_BYTES, une copie
        tronquée (début + fin) est produite pour ne pas faire échouer l'envoi."""
        path = getattr(self, "log_path", None)
        if not path or not os.path.exists(path):
            return None
        try:
            taille = os.path.getsize(path)
            if taille <= MAX_ATTACHMENT_BYTES:
                return path
            garde = MAX_ATTACHMENT_BYTES // 2
            with open(path, "rb") as fh:
                debut = fh.read(garde)
                fh.seek(taille - garde)
                fin = fh.read(garde)
            court = path.replace(".log", "_tronque.log")
            with open(court, "wb") as fh:
                fh.write(debut)
                fh.write(f"\n\n[... {taille - 2 * garde} octets omis — "
                         f"log complet : {path} ...]\n\n".encode("utf-8"))
                fh.write(fin)
            return court
        except OSError:
            return path

    def _send_report(self, results, started, ended, all_ok, par_filiale=None):
        from kyc.models import EmailReminderConfig
        from kyc.daterev_mailer import parse_recipients, send_html_email

        cfg = EmailReminderConfig.objects.filter(active=True).order_by("-updated_at").first()
        if not cfg:
            self.stderr.write("Aucune configuration SMTP : rapport non envoyé.")
            return
        recipients = parse_recipients(cfg.notify_emails)
        if not recipients:
            self.stderr.write("Aucun email de supervision configuré : rapport non envoyé.")
            return

        date_str = started.strftime("%d/%m/%Y %H:%M")
        statut_global = "" if all_ok else " ERREURS DÉTECTÉES"
        head_color = "#0a3d2e" if all_ok else "#9a3412"

        rows = ""
        for r in results:
            if r["status"] == "ignoré":
                color, badge = "#94a3b8", "Ignoré"
            elif r["ok"]:
                color, badge = "#059669", "OK"
            else:
                color, badge = "#dc2626", "Échec"
            detail = (r.get("detail") or "").replace("<", "&lt;").replace(">", "&gt;")
            rows += (
                f'<tr>'
                f'<td style="padding:8px 12px;border-bottom:1px solid #eee;font-weight:600;color:#334155;">{r["label"]}</td>'
                f'<td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;">'
                f'<span style="background:{color};color:#fff;font-size:11px;font-weight:700;padding:2px 10px;border-radius:100px;">{badge}</span></td>'
                f'<td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;color:#64748b;">{r["duration"]}s</td>'
                f'<td style="padding:8px 12px;border-bottom:1px solid #eee;color:#475569;font-size:12px;">{detail}</td>'
                f'</tr>'
            )

        bloc_filiales = self._html_par_filiale(par_filiale or {})

        piece_jointe = self._attachment()
        mention_pj = (f'Log complet de l\'exécution en pièce jointe '
                      f'(<b>{os.path.basename(piece_jointe)}</b>).'
                      if piece_jointe else
                      'Aucun fichier log disponible pour cette exécution.')

        html = f"""
        <div style="font-family:Segoe UI,Arial,sans-serif;max-width:760px;margin:auto;background:#fff;border-radius:12px;overflow:hidden;border:1px solid #e2e8f0;">
          <div style="background:{head_color};padding:22px 28px;">
            <h1 style="color:#fff;font-size:17px;margin:0 0 4px;">Rapport des traitements quotidiens — KYC BOA</h1>
            <p style="color:rgba(255,255,255,.7);font-size:12px;margin:0;">Exécution du {date_str} · {statut_global}</p>
          </div>
          <div style="padding:20px 28px;">
            <h2 style="font-size:13px;color:#0f172a;margin:0 0 10px;text-transform:uppercase;letter-spacing:.04em;">1. Traitements globaux</h2>
            <table style="width:100%;border-collapse:collapse;font-size:13px;">
              <thead><tr style="background:#065f46;color:#fff;">
                <th style="padding:8px 12px;text-align:left;">Traitement</th>
                <th style="padding:8px 12px;">Statut</th>
                <th style="padding:8px 12px;text-align:right;">Durée</th>
                <th style="padding:8px 12px;text-align:left;">Détail</th>
              </tr></thead>
              <tbody>{rows}</tbody>
            </table>
          </div>
          {bloc_filiales}
          <div style="padding:0 28px 20px;font-size:12px;color:#475569;">{mention_pj}</div>
          <div style="background:#f8fafc;padding:14px 28px;border-top:1px solid #e2e8f0;font-size:11px;color:#94a3b8;">
            Message automatique · Plateforme Fiabilisation KYC v3.0 · BOA Group
          </div>
        </div>
        """
        subject = f"Traitements quotidiens — {'OK' if all_ok else 'PROBLEMES'} ({started.strftime('%d/%m/%Y')})"
        n = send_html_email(cfg, recipients, subject, html,
                            attachments=[piece_jointe] if piece_jointe else None)
        self.stdout.write(self.style.SUCCESS(f"Rapport de supervision envoyé à {n} destinataire(s)."))
        self.stdout.write(f"Log de l'exécution : {self.log_path}")

    @staticmethod
    def _fraicheur(etapes):
        """Extrait l'étape « Derniere date traitee » d'une filiale.

        Le message est produit par `kyc_controle_fraicheur` dans Script_V3.r au
        format : « PP : jj/mm/aaaa | PM : jj/mm/aaaa | mode X | attendu au plus
        tard le jj/mm/aaaa [| ANOMALIE : ...] ». Retourne None si l'étape est
        absente (script d'une version antérieure, ou filiale arrêtée avant)."""
        import re

        ligne = next((e for e in etapes
                      if e["etape"].lower().startswith("derniere date traitee")), None)
        if not ligne:
            return None
        msg = ligne["message"]

        def champ(motif, defaut="?"):
            m = re.search(motif, msg)
            return m.group(1).strip() if m else defaut

        return {
            "pp": champ(r"PP\s*:\s*([^|]+)"),
            "pm": champ(r"PM\s*:\s*([^|]+)"),
            "mode": champ(r"\bmode\s+([^|]+)"),
            "attendu": champ(r"attendu au plus tard le\s+([^|]+)"),
            "anomalie": champ(r"ANOMALIE\s*:\s*(.+)$", ""),
            "statut": ligne["statut"],
        }

    @staticmethod
    def _esc(texte):
        return (str(texte or "").replace("&", "&amp;")
                .replace("<", "&lt;").replace(">", "&gt;"))

    def _html_par_filiale(self, par_filiale):
        """Section « détail par filiale » : une carte par filiale avec ses étapes
        Script_V3.r (statut, durée, message)."""
        if not par_filiale:
            return ('<div style="padding:0 28px 20px;font-size:12px;color:#94a3b8;">'
                    'Aucun détail par filiale : le journal de Script_V3.r est absent '
                    '(script ignoré ou non exécuté).</div>')

        couleurs = {"OK": "#059669", "ATTENTION": "#d97706", "ANOMALIE": "#b91c1c",
                    "ÉCHEC": "#dc2626", "ECHEC": "#dc2626", "IGNORE": "#94a3b8"}
        cartes = ""
        for fil, etapes in par_filiale.items():


            resume, detail = [], []
            for e in etapes:
                (resume if e["etape"].lower().startswith("traitement complet")
                 else detail).append(e)
            statut_fil = resume[-1]["statut"] if resume else (
                "ECHEC" if any(e["statut"] in ("ECHEC", "ÉCHEC") for e in etapes) else "OK")
            duree_fil = resume[-1]["duree"] if resume else ""
            nb_ko = sum(1 for e in detail if e["statut"] in ("ECHEC", "ÉCHEC"))
            nb_warn = sum(1 for e in detail if e["statut"] == "ATTENTION")
            nb_anomalie = sum(1 for e in detail if e["statut"] == "ANOMALIE")


            if statut_fil == "OK" and nb_anomalie:
                statut_fil = "ANOMALIE"
            c_fil = couleurs.get(statut_fil, "#64748b")

            lignes = ""
            for e in detail:
                c = couleurs.get(e["statut"], "#64748b")
                duree = f'{e["duree"]}s' if e["duree"] else "—"
                lignes += (
                    f'<tr>'
                    f'<td style="padding:6px 10px;border-bottom:1px solid #f1f5f9;color:#334155;">{self._esc(e["etape"])}</td>'
                    f'<td style="padding:6px 10px;border-bottom:1px solid #f1f5f9;text-align:center;">'
                    f'<span style="background:{c};color:#fff;font-size:10px;font-weight:700;'
                    f'padding:2px 8px;border-radius:100px;">{self._esc(e["statut"])}</span></td>'
                    f'<td style="padding:6px 10px;border-bottom:1px solid #f1f5f9;text-align:right;color:#64748b;">{duree}</td>'
                    f'<td style="padding:6px 10px;border-bottom:1px solid #f1f5f9;color:#475569;font-size:11px;">{self._esc(e["message"])}</td>'
                    f'</tr>'
                )
            if not lignes:
                lignes = ('<tr><td colspan="4" style="padding:8px 10px;color:#94a3b8;font-size:12px;">'
                          'Aucune étape journalisée.</td></tr>')

            compteurs = []
            if nb_ko:
                compteurs.append(f"{nb_ko} étape(s) en échec")
            if nb_anomalie:
                compteurs.append(f"{nb_anomalie} anomalie(s) de fraîcheur")
            if nb_warn:
                compteurs.append(f"{nb_warn} avertissement(s)")
            sous_titre = " · ".join(compteurs) or f"{len(detail)} étape(s) exécutée(s)"
            if duree_fil:
                sous_titre += f" · {duree_fil}s au total"



            fr = self._fraicheur(detail)
            if fr:
                c_fr = couleurs.get(fr["statut"], "#64748b")
                bandeau_dates = (
                    f'<div style="font-size:11px;color:#0f172a;margin-top:5px;">'
                    f'Dernière date traitée — <b>pp_stock : {self._esc(fr["pp"])}</b> · '
                    f'<b>pm_stock : {self._esc(fr["pm"])}</b> · mode « {self._esc(fr["mode"])} » '
                    f'(attendu au plus tard le {self._esc(fr["attendu"])})'
                    + (f'<div style="color:{c_fr};font-weight:700;margin-top:3px;">'
                       f'⚠ Anomalie : {self._esc(fr["anomalie"])}</div>' if fr["anomalie"] else '')
                    + '</div>')
            else:
                bandeau_dates = ('<div style="font-size:11px;color:#94a3b8;margin-top:5px;">'
                                 'Dernière date traitée non journalisée.</div>')

            cartes += f"""
            <div style="border:1px solid #e2e8f0;border-radius:10px;margin-bottom:14px;overflow:hidden;">
              <div style="background:#f8fafc;padding:10px 14px;border-bottom:1px solid #e2e8f0;">
                <span style="background:{c_fil};color:#fff;font-size:11px;font-weight:700;padding:2px 10px;border-radius:100px;">{self._esc(statut_fil)}</span>
                <span style="font-size:13px;font-weight:700;color:#0f172a;margin-left:8px;">BOA {self._esc(fil)}</span>
                <div style="font-size:11px;color:#64748b;margin-top:3px;">{self._esc(sous_titre)}</div>
                {bandeau_dates}
              </div>
              <table style="width:100%;border-collapse:collapse;font-size:12px;">
                <thead><tr style="background:#0f172a;color:#fff;">
                  <th style="padding:6px 10px;text-align:left;">Étape</th>
                  <th style="padding:6px 10px;">Statut</th>
                  <th style="padding:6px 10px;text-align:right;">Durée</th>
                  <th style="padding:6px 10px;text-align:left;">Message</th>
                </tr></thead>
                <tbody>{lignes}</tbody>
              </table>
            </div>
            """

        return f"""
        <div style="padding:0 28px 8px;">
          <h2 style="font-size:13px;color:#0f172a;margin:0 0 10px;text-transform:uppercase;letter-spacing:.04em;">
            2. Détail par filiale — Fiabilisation (Script_V3.r)</h2>
          {cartes}
        </div>
        """
