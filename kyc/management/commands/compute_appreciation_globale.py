"""
Calcule et stocke l'appréciation globale par agent (filiale + exploitant)
selon la matrice du script R `appreciation_globale.r`.

À planifier (cron / tâche planifiée), p.ex. chaque nuit :
    python manage.py compute_appreciation_globale
"""
from collections import defaultdict

from django.core.management.base import BaseCommand
from django.db.models import Max

from kyc.models import (
    Appreciation_globale, AppreciationConfig, DataQualityRule,
    Notation, TauxEvolution,
)
from kyc.appreciation import (
    appreciation_qualite, appreciation_globale, mesure_from_appreciation,
    current_trimestre,
)


class Command(BaseCommand):
    help = "Calcule l'appréciation globale de chaque agent (matrice appreciation_globale.r)."

    def add_arguments(self, parser):
        parser.add_argument("--filiale", type=str, default=None,
                            help="Limiter à une filiale donnée.")
        parser.add_argument("--verbose", action="store_true", help="Affiche le détail par agent.")
        parser.add_argument("--slice", type=str, default="1/1",
                            help="Découpage parallèle des agents, format i/N (ex. 2/6 = worker 2 sur 6). "
                                 "Chaque worker traite un sous-ensemble disjoint des agents.")

    def handle(self, *args, **options):
        from accounts.models import ProfileV                                    

                                                                       
        trimestre_par_filiale = {}
        methode_par_filiale = {}
        for cfg in AppreciationConfig.objects.filter(active=True).exclude(filiale=""):
            key = cfg.filiale.strip().upper()
            trimestre_par_filiale[key] = current_trimestre(cfg.date_demarrage)
            methode_par_filiale[key] = (cfg.methode_taux or "flux").strip().lower()
        if trimestre_par_filiale:
            apercu = ", ".join(f"{f}=T{t}/{methode_par_filiale.get(f, 'flux')}"
                               for f, t in sorted(trimestre_par_filiale.items()))
            self.stdout.write(f"Config par filiale : {apercu}")
        else:
            self.stdout.write("Aucune configuration par filiale — trimestre 1 / méthode flux par défaut.")

        def trimestre_for(fil):
            return trimestre_par_filiale.get((fil or "").strip().upper(), 1)

        def methode_for(fil):
            return methode_par_filiale.get((fil or "").strip().upper(), "flux")

        from kyc.views import flux_datouv_window
        flux_start, flux_end = flux_datouv_window()

        filiale_filter = options.get("filiale")

                                                                                        
        evo_qs = TauxEvolution.objects.filter(flux_stock__in=("F", "S")).exclude(expl="")
        if filiale_filter:
            evo_qs = evo_qs.filter(filiale=filiale_filter)
        agents = set(evo_qs.values_list("filiale", "expl").distinct())

        if not agents:
            self.stdout.write(self.style.WARNING("Aucun agent (TauxEvolution) trouvé."))
            return

                                                                                 
                                                                                 
        try:
            slice_idx, slice_total = (int(x) for x in options["slice"].split("/"))
        except (ValueError, AttributeError):
            slice_idx, slice_total = 1, 1
        if slice_total < 1 or not (1 <= slice_idx <= slice_total):
            slice_idx, slice_total = 1, 1
        agents = sorted(agents)
        if slice_total > 1:
            agents = agents[slice_idx - 1::slice_total]
        if not agents:
            self.stdout.write(self.style.WARNING("Aucun agent pour ce slice."))
            return

                                                                         
        def taux_evolution_agent(fil, expl, methode):
            code = "S" if methode == "stock" else "F"
            vals = []
            for pp in ("P", "M"):
                rec = (TauxEvolution.objects
                       .filter(filiale=fil, expl=expl, flux_stock=code, pp_pm=pp)
                       .order_by("-date").first())
                if rec and rec.taux is not None:
                    vals.append(rec.taux)
            return round(sum(vals) / len(vals), 1) if vals else None

                                                                               
        def _latest_notation_by_agent(flux_stock):
            notes = Notation.objects.filter(flux_stock=flux_stock)
            latest = notes.values("agent").annotate(d=Max("date_notation"))
            latest_dates = {n["agent"]: n["d"] for n in latest}
            out = {}
            for n in notes.filter(date_notation__in=latest_dates.values()).select_related("agent"):
                if latest_dates.get(n.agent_id) == n.date_notation:
                    out[n.agent_id] = n.note
            return out

        notation_by_methode = {
            "flux": _latest_notation_by_agent("Flux"),
            "stock": _latest_notation_by_agent("Stock"),
        }
                                                                    
        profile_idx = {}
        for p in ProfileV.objects.exclude(code_expl="").exclude(code_expl__isnull=True):
            key = ((p.filiale or "").strip().upper(), (p.code_expl or "").strip().upper())
            profile_idx[key] = p

                                                                                        
        from kyc.views import evaluate_data_quality_rule
        rules = list(DataQualityRule.objects.filter(active=True))

        def taux_qualite_agent(fil, expl, methode):
            ds, de = (flux_start, flux_end) if methode == "flux" else (None, None)
            ok = tot = 0
            for rule in rules:
                stat = evaluate_data_quality_rule(rule, filiale=fil, expl=expl,
                                                  datouv_start=ds, datouv_end=de)
                tot += stat.get("total", 0)
                ok += stat.get("ok_count", 0)
            if tot == 0:
                                                                               
                return 100.0
                                                                                           
            import math
            rate = math.floor(100.0 * ok / tot)
                                                                               
                                                       
            if ok < tot:
                rate = min(rate, 99)
            return float(rate)

                                                             
                                                                                
                                                                             
                                                                                 
        rows = []
        for fil, expl in agents:
            trimestre = trimestre_for(fil)
            methode = methode_for(fil)
            te = taux_evolution_agent(fil, expl, methode)
            tq = taux_qualite_agent(fil, expl, methode)
            prof = profile_idx.get((fil.strip().upper(), expl.strip().upper()))
            note = notation_by_methode[methode].get(prof.id, "") if prof else ""

            appr_q = appreciation_qualite(tq, note)
            appr_g = appreciation_globale(te, appr_q, trimestre)
            mesure = mesure_from_appreciation(appr_g)

            rows.append((fil, expl, {
                "agent": prof,
                "trimestre": trimestre,
                "methode_taux": methode,
                "taux_evolution": te,
                "taux_qualite": tq,
                "notation": note or "",
                "appreciation_qualite": appr_q,
                "appreciation_globale": appr_g,
                "mesure": mesure,
            }))
            if options.get("verbose"):
                self.stdout.write(f"  {fil} / {expl} (T{trimestre}, {methode}): évo={te} qual={tq} "
                                  f"note={note or '—'} -> {appr_q} / {appr_g}")

                                                                               
        import time
        from django.db import transaction
        from django.db.utils import OperationalError

        created = updated = 0
        for attempt in range(5):
            try:
                with transaction.atomic():
                    created = updated = 0
                    for fil, expl, defaults in rows:
                        _, is_created = Appreciation_globale.objects.update_or_create(
                            filiale=fil, expl=expl, defaults=defaults,
                        )
                        created += int(is_created)
                        updated += int(not is_created)
                break
            except OperationalError as e:
                if "locked" not in str(e).lower() or attempt == 4:
                    raise
                wait = 5 * (attempt + 1)
                self.stdout.write(f"Base verrouillée, nouvel essai dans {wait}s "
                                  f"(tentative {attempt + 2}/5)...")
                time.sleep(wait)

        self.stdout.write(self.style.SUCCESS(
            f"Appréciation globale calculée : {created} créées, {updated} mises à jour "
            f"({created + updated} agents)."))
