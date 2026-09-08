from django.contrib import admin
from django.core.management import call_command

from .models import (KycDocumentExtraction, KycDocumentMatchJob, KycDocumentMatchSettings,
                     KycExpiredDocumentScanMatch, FilialeModuleConfig, EmailReminderConfig,
                     AppreciationConfig, Appreciation_globale, TermTranslation, KycDocumentOcrJob,
                     KycMatchValidatorRole, KycMatchDecision, KycScreeningAccess,
                     SidebarAccess, DataQualityRule, DataQualityCondition, QualityFluxConfig, TauxQualite)


@admin.register(QualityFluxConfig)
class QualityFluxConfigAdmin(admin.ModelAdmin):
    list_display = ("flux_window", "active", "updated_at")
    list_editable = ("active",)
    list_display_links = ("flux_window",)


@admin.register(TauxQualite)
class TauxQualiteAdmin(admin.ModelAdmin):
    list_display = ("date", "flux_stock", "applicability", "filiale", "agence", "expl",
                    "rate", "ok_count", "total")
    list_filter = ("flux_stock", "applicability", "date", "filiale")
    search_fields = ("filiale", "agence", "expl")
    date_hierarchy = "date"
    ordering = ("-date", "filiale")


class DataQualityConditionInline(admin.TabularInline):
    model = DataQualityCondition
    extra = 0
    fields = ("logic", "field_name", "operator", "value")


@admin.register(DataQualityRule)
class DataQualityRuleAdmin(admin.ModelAdmin):
    list_display = ("id", "name", "applicability", "control_type", "active",
                    "filiale", "created_by", "created_at")
    list_filter = ("applicability", "control_type", "active")
    search_fields = ("name", "field_name", "filiale")
    list_editable = ("active",)
    readonly_fields = ("created_at",)
    inlines = [DataQualityConditionInline]


@admin.register(DataQualityCondition)
class DataQualityConditionAdmin(admin.ModelAdmin):
    list_display = ("id", "rule", "logic", "field_name", "operator", "value")
    list_filter = ("logic", "operator")
    search_fields = ("field_name", "value", "rule__name")


@admin.register(KycMatchValidatorRole)
class KycMatchValidatorRoleAdmin(admin.ModelAdmin):
    list_display = ("organe", "can_validate", "can_reject", "updated_at")
    list_editable = ("can_validate", "can_reject")
    list_filter = ("can_validate", "can_reject")


@admin.register(KycScreeningAccess)
class KycScreeningAccessAdmin(admin.ModelAdmin):
    list_display = ("organe", "tab_charger", "tab_suivi", "tab_resultats", "tab_sources",
                    "tab_documents", "can_upload_batches", "can_run_matching", "updated_at")
    list_editable = ("tab_charger", "tab_suivi", "tab_resultats", "tab_sources",
                     "tab_documents", "can_upload_batches", "can_run_matching")
    list_filter = ("can_upload_batches", "can_run_matching")


@admin.register(SidebarAccess)
class SidebarAccessAdmin(admin.ModelAdmin):
    list_display = ("organe", "dashboard", "agents_notes", "champs_non_renseignes",
                    "clients_anomalie", "scoring_clients", "screening_kyc",
                    "nouvelle_notation", "historique_notation", "ppe",
                    "comptes_specifiques", "parametrage_utilisateurs",
                    "regles_qualite", "champs_kyc", "documents_screening",
                    "rappels_scoring", "pilotage", "audit", "updated_at")
    list_editable = ("dashboard", "agents_notes", "champs_non_renseignes",
                     "clients_anomalie", "scoring_clients", "screening_kyc",
                     "nouvelle_notation", "historique_notation", "ppe",
                     "comptes_specifiques", "parametrage_utilisateurs",
                     "regles_qualite", "champs_kyc", "documents_screening",
                     "rappels_scoring", "pilotage", "audit")
    list_filter = ("dashboard", "screening_kyc", "pilotage", "audit")
    search_fields = ("organe",)


@admin.register(KycMatchDecision)
class KycMatchDecisionAdmin(admin.ModelAdmin):
    list_display = ("client_code", "client_type", "status", "match_rate", "filiale", "agence", "decided_by", "decided_at")
    list_filter = ("status", "client_type", "filiale")
    search_fields = ("client_code", "filiale", "agence")
    readonly_fields = ("created_at", "updated_at")


@admin.register(KycDocumentOcrJob)
class KycDocumentOcrJobAdmin(admin.ModelAdmin):
    list_display = ('id', 'import_batch', 'mode', 'status', 'progress_current', 'progress_total',
                    'done_count', 'failed_count', 'created_by', 'created_at', 'completed_at')
    list_filter = ('status', 'mode')
    search_fields = ('import_batch',)
    readonly_fields = ('created_at', 'started_at', 'completed_at', 'updated_at')
    ordering = ('-created_at',)


@admin.register(TermTranslation)
class TermTranslationAdmin(admin.ModelAdmin):
    list_display = ('terme_fr', 'terme_en', 'note', 'updated_at')
    list_editable = ('terme_en',)
    search_fields = ('terme_fr', 'terme_en', 'note')
    ordering = ('terme_fr',)


@admin.register(AppreciationConfig)
class AppreciationConfigAdmin(admin.ModelAdmin):
    list_display = ('filiale', 'date_demarrage', 'trimestre_actuel', 'methode_taux', 'active', 'updated_at')
    list_editable = ('date_demarrage', 'methode_taux', 'active')
    search_fields = ('filiale',)
    list_filter = ('active',)
    ordering = ('filiale',)
    readonly_fields = ('trimestre_actuel', 'updated_at')

    @admin.display(description="Trimestre courant")
    def trimestre_actuel(self, obj):
        return obj.trimestre_actuel() if obj and obj.date_demarrage else "—"


@admin.register(Appreciation_globale)
class AppreciationGlobaleAdmin(admin.ModelAdmin):
    list_display = ('filiale', 'expl', 'trimestre', 'methode_taux', 'taux_evolution', 'taux_qualite',
                    'notation', 'appreciation_qualite', 'appreciation_globale', 'computed_at')
    list_filter = ('filiale', 'trimestre', 'methode_taux', 'appreciation_globale', 'appreciation_qualite')
    search_fields = ('filiale', 'expl')
    readonly_fields = ('computed_at',)
    actions = ['recalculer']

    @admin.action(description="Recalculer l'appréciation globale (tous les agents)")
    def recalculer(self, request, queryset):
        call_command('compute_appreciation_globale')
        self.message_user(request, "Appréciation globale recalculée pour tous les agents.")


@admin.register(FilialeModuleConfig)
class FilialeModuleConfigAdmin(admin.ModelAdmin):
    list_display = ('filiale', 'screening_kyc_paye_active', 'daterev_reminder_paye_active')
    list_editable = ('screening_kyc_paye_active', 'daterev_reminder_paye_active')
    search_fields = ('filiale',)
    list_filter = ('screening_kyc_paye_active', 'daterev_reminder_paye_active')


@admin.register(KycDocumentExtraction)
class KycDocumentExtractionAdmin(admin.ModelAdmin):
    list_display = (
        "created_at",
        "import_batch",
        "document_type",
        "original_filename",
        "page_range",
        "prenom",
        "nom",
        "numero_document",
        "nationalite",
        "uploaded_by",
    )
    list_filter = ("document_type", "created_at", "nationalite", "import_batch")
    search_fields = (
        "import_batch",
        "original_filename",
        "source_filename",
        "prenom",
        "nom",
        "numero_document",
        "date_naissance",
        "date_expiration",
        "nationalite",
        "pays_naissance",
        "numero_identification_nationale",
        "lieu_naissance",
        "adresse",
        "origine_revenu",
        "extracted_text",
    )
    readonly_fields = ("created_at",)


@admin.register(KycDocumentMatchSettings)
class KycDocumentMatchSettingsAdmin(admin.ModelAdmin):
    list_display = (
        "name",
        "pp_fullname_weight", "pp_birth_date_weight", "pp_birth_place_weight", "pp_birth_country_weight",
        "pm_fullname_weight", "pm_fiscal_weight", "pm_address_weight", "pm_country_weight",
        "combination_threshold", "min_display_score", "active", "updated_at",
    )
    readonly_fields = ("updated_at",)
    fieldsets = (
        (None, {"fields": ("name", "active")}),
        ("Poids PP — Personnes Physiques (N° identification nationale = 100 %, somme des poids <= 100)", {
            "fields": (
                "pp_fullname_weight",
                "pp_birth_date_weight",
                "pp_birth_place_weight",
                "pp_birth_country_weight",
            ),
        }),
        ("Poids PM — Personnes Morales (Registre commerce RCSNO = 100 %, somme des poids <= 100)", {
            "fields": (
                "pm_fullname_weight",
                "pm_fiscal_weight",
                "pm_address_weight",
                "pm_country_weight",
            ),
        }),
        ("Seuils communs", {
            "fields": ("combination_threshold", "min_display_score"),
        }),
        ("Equivalence Nom & Prenom / Raison sociale dans les modeles KYC", {
            "description": "Champ des modeles KYC compare au nom & prenom (PP) / a la raison sociale (PM) "
                           "extraits du document. Par defaut INTITULE_COMPTE.",
            "fields": (
                "pp_fullname_field",
                "pm_fullname_field",
            ),
        }),
        ("Suivi", {"fields": ("updated_at",)}),
    )


@admin.register(KycExpiredDocumentScanMatch)
class KycExpiredDocumentScanMatchAdmin(admin.ModelAdmin):
    list_display = (
        "scan_date",
        "status",
        "client_code",
        "idp",
        "filiale",
        "agence",
        "old_validity_date",
        "document_validity_date",
        "match_rate",
        "document",
    )
    list_filter = ("status", "scan_date", "filiale", "agence")
    search_fields = (
        "client_code",
        "idp",
        "filiale",
        "agence",
        "old_validity_date",
        "document_validity_date",
        "document__original_filename",
        "document__import_batch",
    )
    readonly_fields = ("scan_date", "updated_at")
    list_editable = ("status",)


@admin.register(KycDocumentMatchJob)
class KycDocumentMatchJobAdmin(admin.ModelAdmin):
    list_display = (
        "created_at",
        "status",
        "progress_current",
        "progress_total",
        "message",
        "created_by",
    )
    list_filter = ("status", "created_at")
    search_fields = ("message", "error")
    readonly_fields = ("created_at", "started_at", "completed_at", "updated_at")


@admin.register(EmailReminderConfig)
class EmailReminderConfigAdmin(admin.ModelAdmin):
    list_display = ('smtp_host', 'smtp_port', 'smtp_mode_display', 'smtp_user', 'from_email', 'frequency', 'days_before', 'active', 'updated_at')
    list_editable = ('active',)
    readonly_fields = ('smtp_mode_display',)
    fieldsets = (
        ('Configuration SMTP', {
            'description': (
                '<div style="background:#fef3c7;border-left:4px solid #f59e0b;padding:10px 14px;border-radius:6px;margin-bottom:12px;font-size:12px;">'
                '<strong>Combinaisons valides :</strong><br>'
                '&bull; Port <strong>587</strong> → cocher <em>Utiliser TLS</em>, décocher <em>Utiliser SSL</em><br>'
                '&bull; Port <strong>465</strong> → cocher <em>Utiliser SSL</em>, décocher <em>Utiliser TLS</em><br>'
                '&bull; Port <strong>25</strong> → décocher les deux'
                '</div>'
            ),
            'fields': ('smtp_host', 'smtp_port', 'smtp_use_tls', 'smtp_use_ssl', 'smtp_mode_display', 'smtp_user', 'smtp_password', 'from_email', 'from_name'),
        }),
        ('Paramètres de rappel', {
            'fields': ('frequency', 'days_before', 'active'),
        }),
        ('Supervision des tâches quotidiennes', {
            'fields': ('notify_emails',),
        }),
    )

    @admin.display(description='Mode détecté')
    def smtp_mode_display(self, obj):
        if not obj.pk:
            return '—'
        if obj.smtp_use_ssl and obj.smtp_use_tls:
            return '⚠️ Invalide — SSL et TLS simultanés'
        if obj.smtp_use_ssl:
            return f'🔒 SSL direct (port {obj.smtp_port})'
        if obj.smtp_use_tls:
            return f'🔐 STARTTLS (port {obj.smtp_port})'
        return f'⚠️ Non chiffré (port {obj.smtp_port})'

    def save_model(self, request, obj, form, change):
        from django.contrib import messages as dj_messages
        if obj.smtp_use_ssl and obj.smtp_use_tls:
            self.message_user(request, "⚠️ SSL et TLS ne peuvent pas être activés simultanément. Désactivez l'un des deux.", level='warning')
        if obj.smtp_port == 465 and obj.smtp_use_tls and not obj.smtp_use_ssl:
            self.message_user(request, "⚠️ Port 465 détecté avec TLS — vous devriez utiliser SSL à la place.", level='warning')
        if obj.smtp_port == 587 and obj.smtp_use_ssl and not obj.smtp_use_tls:
            self.message_user(request, "⚠️ Port 587 détecté avec SSL — vous devriez utiliser TLS (STARTTLS) à la place.", level='warning')
        super().save_model(request, obj, form, change)
