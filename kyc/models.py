from django.db import models
from django.core.exceptions import ValidationError
from django.core.validators import MaxValueValidator
from datetime import datetime
from django.contrib.auth.models import User
from django.utils import timezone
from django.db import models
from django.utils import timezone
from django.conf import settings

from PIL import Image


Type = (
    ('Agent', 'Agent'),
    ('Contrôleur', 'Contrôleur'),
)
FS = (
    ('Flux', 'Flux'),
    ('Stock', 'Stock'),
)

NoteChoices = (
    ('Très Bien', 'Très Bien'),
    ('Bien', 'Bien'),
    ('Passable', 'Passable'),
    ('Insuffisant','Insuffisant'),
)

APPLICABILITY_CHOICES = (
    ('PP', 'Client PP'),
    ('PM', 'Client PM'),
)

DATA_QUALITY_CONTROL_TYPE_CHOICES = (
    ('simple', 'Contrôle simple (Existence / Valeur)'),
    ('composite', 'Règle multi-critères (Composite)'),
)

OPERATOR_CHOICES = (
    ('=', 'Égal à (=)'),
    ('!=', 'Différent de (!=)'),
    ('>', 'Supérieur à (>)'),
    ('<', 'Inférieur à (<)'),
    ('>=', 'Supérieur ou égal (>=)'),
    ('<=', 'Inférieur ou égal (<=)'),
    ('contains', 'Contient'),
    ('not_contains', 'Ne contient pas'),
    ('word_contains', 'Contient le mot exact'),
    ('word_not_contains', 'Ne contient pas le mot exact'),
    ('contains_alpha', 'Contient des lettres'),
    ('contains_digit', 'Contient des chiffres'),
    ('regex', 'Expression régulière'),
    ('is_empty', 'Est vide'),
    ('is_not_empty', 'N\'est pas vide'),
    ('expired', 'Est expiré (Date < Aujourd\'hui)'),
    ('age_gt', 'Âge supérieur à'),
    ('age_lt', 'Âge inférieur à'),
    ('min_length', 'Longueur minimum'),
    ('max_length', 'Longueur maximum'),
)

LOGIC_CHOICES = (
    ('AND', 'ET'),
    ('OR', 'OU'),
)

PP_FIELD_CHOICES = (
    ('FILIALE', 'FILIALE'),
    ('AGENCE', 'AGENCE'),
    ('LIB_AGENCE', 'LIB_AGENCE'),
    ('EXPL', 'EXPL'),
    ('CLIENT', 'CLIENT'),
    ('CODAPE', 'CODAPE'),
    ('IDP', 'IDP'),
    ('PAYNAIS', 'PAYNAIS'),
    ('PROFESSION', 'PROFESSION'),
    ('ADRESSE', 'ADRESSE'),
    ('PAYS_RESID', 'PAYS_RESID'),
    ('NUMID', 'NUMID'),
    ('SALAIRE', 'SALAIRE'),
    ('ORIGINE_REV', 'ORIGINE_REV'),
    ('DATVALID', 'DATVALID'),
    ('DATNAIS', 'DATNAIS'),
    ('TEL', 'TEL'),
    ('DATOUV', 'DATOUV'),
    ('PPE', 'PPE'),
    ('DEVISE', 'DEVISE'),
    ('RESID', 'RESID'),
    ('DATEREV', 'DATEREV'),
    ('RISQUE', 'RISQUE'),
    ('BOITE_POSTALE', 'BOITE_POSTALE'),
    ('CONSENT_BIC', 'CONSENT_BIC'),
    ('EMPLOYEUR', 'EMPLOYEUR'),
    ('INTITULE_COMPTE', 'INTITULE_COMPTE'),
    ('LIEU_DELIVRANCE_CIN', 'LIEU_DELIVRANCE_CIN'),
)

PM_FIELD_CHOICES = (
    ('FILIALE', 'FILIALE'),
    ('AGENCE', 'AGENCE'),
    ('LIB_AGENCE', 'LIB_AGENCE'),
    ('EXPL', 'EXPL'),
    ('CLIENT', 'CLIENT'),
    ('AGEC', 'AGEC'),
    ('CODAPE', 'CODAPE'),
    ('IDM', 'IDM'),
    ('RCSNO', 'RCSNO'),
    ('CAPITAL', 'CAPITAL'),
    ('CA', 'CA'),
    ('RESULTAT', 'RESULTAT'),
    ('ORIGINE_REV', 'ORIGINE_REV'),
    ('DATOUV', 'DATOUV'),
    ('TEL', 'TEL'),
    ('DEVISE', 'DEVISE'),
    ('RESID', 'RESID'),
    ('DATEREV', 'DATEREV'),
    ('PPE', 'PPE'),
    ('RISQUE', 'RISQUE'),
    ('ACTIONNAIRE', 'ACTIONNAIRE'),
    ('ADRESSE_SOCIALE', 'ADRESSE_SOCIALE'),
    ('BOITE_POSTALE', 'BOITE_POSTALE'),
    ('CONSENT_BIC', 'CONSENT_BIC'),
    ('INTITULE_COMPTE', 'INTITULE_COMPTE'),
    ('MANDATAIRE', 'MANDATAIRE'),
    ('NUMERO_FISCAL', 'NUMERO_FISCAL'),
    ('PAYS_JUR', 'PAYS_JUR'),
)

DATA_QUALITY_FIELD_CHOICES = PP_FIELD_CHOICES + PM_FIELD_CHOICES

class DataQualityRule(models.Model):
    name = models.CharField(max_length=200)
    applicability = models.CharField(max_length=3, choices=APPLICABILITY_CHOICES)
    field_name = models.CharField(max_length=100, choices=DATA_QUALITY_FIELD_CHOICES)
    control_type = models.CharField(max_length=50, choices=DATA_QUALITY_CONTROL_TYPE_CHOICES)
    parameter = models.CharField(blank=True, max_length=200, help_text='Seuil, longueur ou valeur de référence')
    description = models.TextField(blank=True, null=True)
    active = models.BooleanField(default=True)
    filiale = models.TextField(blank=True)
    created_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.name} [{self.applicability}]"

    @property
    def rule_number(self):
        if hasattr(self, '_rule_number') and self._rule_number is not None:
            return self._rule_number
        rule_map = get_rule_number_map()
        return rule_map.get((self.name or '').strip(), self.id)

def get_rule_number_map():
    """
    Retourne un dictionnaire { nom_regle.strip(): numero_unique }.
    Les règles sont numérotées de façon unique selon leur intitulé (name).
    Deux règles ayant le même intitulé ont exactement le même numéro unique,
    même si elles s'appliquent à des filiales différentes.
    L'ordre d'attribution (1, 2, 3...) est basé sur le min(id) de chaque intitulé.
    """
    from django.db.models import Min
    rule_names = (
        DataQualityRule.objects
        .values('name')
        .annotate(min_id=Min('id'))
        .order_by('min_id')
    )
    mapping = {}
    idx = 1
    for item in rule_names:
        name = (item['name'] or '').strip()
        if name and name not in mapping:
            mapping[name] = idx
            idx += 1
    return mapping


class DataQualityCondition(models.Model):
    rule = models.ForeignKey(DataQualityRule, on_delete=models.CASCADE, related_name='conditions')
    field_name = models.CharField(max_length=100, choices=DATA_QUALITY_FIELD_CHOICES)
    operator = models.CharField(max_length=20, choices=OPERATOR_CHOICES)
    value = models.CharField(max_length=255, blank=True, null=True, help_text="Valeur fixe")
    logic = models.CharField(
        max_length=3, choices=LOGIC_CHOICES, default='AND',
        help_text="Connecteur logique avec la condition précédente (ET / OU). "
                  "OU démarre un nouveau groupe ; ignoré pour la 1re condition.")

    class Meta:
        ordering = ['id']

    def __str__(self):
        return f"{self.logic} {self.field_name} {self.operator} {self.value}"

class DataQualityRuleAudit(models.Model):
    rule_name = models.CharField(max_length=200)
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True)
    action = models.CharField(max_length=50)                                      
    details = models.TextField(blank=True, null=True)
    timestamp = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-timestamp']


Filiales = (
    ('BOA SN', 'BOA SN'),
)

class FilialeModuleConfig(models.Model):
    filiale = models.CharField(max_length=15, choices=Filiales, unique=True, verbose_name="Filiale/Pays")
    screening_kyc_paye_active = models.BooleanField(default=False, verbose_name="Module Screening KYC PAYE actif")
    daterev_reminder_paye_active = models.BooleanField(
        default=False, verbose_name="Module Rappels Scoring PAYE actif",
        help_text="Si décoché, les chargés de cette filiale ne reçoivent pas les rappels Scoring automatiques.")

    class Meta:
        verbose_name = "Configuration Module Filiale"
        verbose_name_plural = "Configurations Modules Filiales"

    def __str__(self):
        return f"Config {self.filiale} - Screening KYC: {'Oui' if self.screening_kyc_paye_active else 'Non'}"


DOCUMENT_EXTRACTION_TYPE_CHOICES = (
    ('piece_identite', "Piece d'identite"),
    ('passeport', 'Passeport'),
)

CLIENT_TYPE_CHOICES = (
    ('pp', 'Particuliers (PP)'),
    ('pm', 'Entreprises (PM)'),
)


class KycDocumentType(models.Model):
    code = models.CharField(max_length=50, verbose_name="Code technique")
    label = models.CharField(max_length=150, verbose_name="Libellé affiché")
    filiale = models.CharField(max_length=15, choices=Filiales, blank=True, default="", verbose_name="Filiale/Pays")
    client_type = models.CharField(max_length=10, choices=CLIENT_TYPE_CHOICES, default='pp', verbose_name="Type de client")
    keywords = models.TextField(blank=True, help_text="Mots-clés séparés par des virgules")
    min_score = models.FloatField(default=2.0, help_text="Score minimum pour être détecté")
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['label']
        unique_together = ('code', 'filiale', 'client_type')

    def __str__(self):
        return self.label

    def get_keyword_list(self):
        return [k.strip().upper() for k in self.keywords.split(',') if k.strip()]


class KycFieldVisibilityConfig(models.Model):
    CLIENT_TYPE_CHOICES = [
        ("pp", "Particuliers (PP)"),
        ("pm", "Entreprises (PM)"),
    ]
    client_type = models.CharField(max_length=2, choices=CLIENT_TYPE_CHOICES)
    empty_check_fields = models.JSONField(blank=True, default=list)
    display_fields = models.JSONField(blank=True, default=list)
    filiales = models.JSONField(blank=True, default=list)
    field_sources = models.JSONField(blank=True, default=dict)
                                                                                   
                                                                             
                                                             
    field_labels = models.JSONField(blank=True, default=dict)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Configuration champs KYC"
        verbose_name_plural = "Configurations champs KYC"

    def __str__(self):
        return f"{self.get_client_type_display()} - {len(self.filiales)} filiales"


class KycDocumentExtraction(models.Model):
    EXTRACTION_STATUS_CHOICES = (
        ("pending", "En attente OCR"),
        ("processing", "OCR en cours"),
        ("done", "Traite"),
        ("failed", "Echec OCR"),
    )

    document_type = models.CharField(max_length=30, choices=DOCUMENT_EXTRACTION_TYPE_CHOICES)
    uploaded_file = models.FileField(upload_to='document_extraction/')
    original_filename = models.CharField(max_length=255, blank=True)
    source_filename = models.CharField(max_length=255, blank=True)
    import_batch = models.CharField(max_length=120, blank=True)
    extraction_status = models.CharField(max_length=20, choices=EXTRACTION_STATUS_CHOICES, default="done", db_index=True)
    file_hash = models.CharField(max_length=64, blank=True, default="", db_index=True)
    client_type = models.CharField(max_length=10, choices=CLIENT_TYPE_CHOICES, default='pp', verbose_name="Type de client")
    page_number = models.PositiveIntegerField(null=True, blank=True)
    page_range = models.CharField(max_length=30, blank=True)
    nom = models.CharField(max_length=120, blank=True)
    prenom = models.CharField(max_length=120, blank=True)
    numero_document = models.CharField(max_length=120, blank=True)
    date_naissance = models.CharField(max_length=120, blank=True)
    date_expiration = models.CharField(max_length=120, blank=True)
    nationalite = models.CharField(max_length=120, blank=True)
    pays_naissance = models.CharField(max_length=120, blank=True)
    pays_delivrance = models.CharField(max_length=120, blank=True)
    numero_identification_nationale = models.CharField(max_length=120, blank=True)
    lieu_naissance = models.CharField(max_length=120, blank=True)
    adresse = models.CharField(max_length=255, blank=True)
    origine_revenu = models.CharField(max_length=120, blank=True)
    extracted_text = models.TextField(blank=True)
    extraction_warnings = models.TextField(blank=True)
    uploaded_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at']
        indexes = [
            models.Index(fields=['document_type']),
            models.Index(fields=['numero_document']),
            models.Index(fields=['nom']),
            models.Index(fields=['prenom']),
            models.Index(fields=['date_naissance']),
            models.Index(fields=['date_expiration']),
            models.Index(fields=['nationalite']),
            models.Index(fields=['pays_delivrance']),
            models.Index(fields=['import_batch']),
            models.Index(fields=['client_type']),
        ]

    def __str__(self):
        reference = self.numero_document or self.original_filename or str(self.pk)
        return f"{self.get_document_type_display()} - {reference}"


                                                                                              
                                                                                         
KYC_PP_NAME_FIELD_CHOICES = (
    ('INTITULE_COMPTE', "INTITULE_COMPTE (Nom & Prenom)"),
    ('EMPLOYEUR', "EMPLOYEUR (Employeur)"),
)
KYC_PM_NAME_FIELD_CHOICES = (
    ('INTITULE_COMPTE', "INTITULE_COMPTE (Raison sociale / Denomination)"),
    ('ACTIONNAIRE', "ACTIONNAIRE (Actionnaire)"),
    ('MANDATAIRE', "MANDATAIRE (Mandataire)"),
)


class KycDocumentMatchSettings(models.Model):
    name = models.CharField(max_length=80, default="Parametrage standard", unique=True)

                                                                                        
    pp_fullname_weight = models.PositiveSmallIntegerField(default=35, validators=[MaxValueValidator(100)], verbose_name="PP · Poids nom & prenom")
    pp_birth_date_weight = models.PositiveSmallIntegerField(default=35, validators=[MaxValueValidator(100)], verbose_name="PP · Poids date de naissance")
    pp_birth_place_weight = models.PositiveSmallIntegerField(default=15, validators=[MaxValueValidator(100)], verbose_name="PP · Poids lieu de naissance")
    pp_birth_country_weight = models.PositiveSmallIntegerField(default=15, validators=[MaxValueValidator(100)], verbose_name="PP · Poids pays de naissance")

                                                                            
    pm_fullname_weight = models.PositiveSmallIntegerField(default=35, validators=[MaxValueValidator(100)], verbose_name="PM · Poids raison sociale")
    pm_fiscal_weight = models.PositiveSmallIntegerField(default=35, validators=[MaxValueValidator(100)], verbose_name="PM · Poids numero fiscal")
    pm_address_weight = models.PositiveSmallIntegerField(default=15, validators=[MaxValueValidator(100)], verbose_name="PM · Poids adresse sociale")
    pm_country_weight = models.PositiveSmallIntegerField(default=15, validators=[MaxValueValidator(100)], verbose_name="PM · Poids pays de creation (PAYS_JUR)")

                                                                                
    pp_fullname_field = models.CharField(max_length=50, choices=KYC_PP_NAME_FIELD_CHOICES, default="INTITULE_COMPTE",
                                         verbose_name="Champ nom & prenom (KYC PP)")
    pm_fullname_field = models.CharField(max_length=50, choices=KYC_PM_NAME_FIELD_CHOICES, default="INTITULE_COMPTE",
                                         verbose_name="Champ raison sociale (KYC PM)")
    combination_threshold = models.PositiveSmallIntegerField(default=65, validators=[MaxValueValidator(100)], verbose_name="Seuil de correspondance combinee")
    min_display_score = models.PositiveSmallIntegerField(default=30, validators=[MaxValueValidator(100)],
                                                         verbose_name="Score minimum affiche",
                                                         help_text="Une correspondance dont le score est inferieur n'est pas proposee.")
    active = models.BooleanField(default=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Parametrage de correspondance KYC ID"
        verbose_name_plural = "Parametrage des correspondances KYC ID"

    def clean(self):
        pp_total = (self.pp_fullname_weight + self.pp_birth_date_weight
                    + self.pp_birth_place_weight + self.pp_birth_country_weight)
        pm_total = (self.pm_fullname_weight + self.pm_fiscal_weight
                    + self.pm_address_weight + self.pm_country_weight)
        if pp_total > 100:
            raise ValidationError("La somme des poids PP (nom, naissance, lieu, pays) ne doit pas depasser 100.")
        if pm_total > 100:
            raise ValidationError("La somme des poids PM (raison sociale, fiscal, adresse, pays) ne doit pas depasser 100.")

    def save(self, *args, **kwargs):
        self.clean()
        super().save(*args, **kwargs)

    def __str__(self):
        return self.name

    @classmethod
    def get_active(cls):
        settings = cls.objects.filter(active=True).order_by("-updated_at").first()
        if settings:
            return settings
        return cls(
            name="Parametrage standard",
            pp_fullname_weight=35, pp_birth_date_weight=35,
            pp_birth_place_weight=15, pp_birth_country_weight=15,
            pm_fullname_weight=35, pm_fiscal_weight=35,
            pm_address_weight=15, pm_country_weight=15,
            combination_threshold=65,
            active=True,
        )


class KycExpiredDocumentScanMatch(models.Model):
    STATUS_CHOICES = (
        ("a_valider", "A valider"),
        ("valide", "Valide"),
        ("rejete", "Rejete"),
    )

    client = models.ForeignKey("Kyc_pp", on_delete=models.CASCADE, related_name="expired_document_matches")
    document = models.ForeignKey(KycDocumentExtraction, on_delete=models.CASCADE, related_name="expired_kyc_matches")
    client_code = models.CharField(max_length=200, blank=True)
    idp = models.CharField(max_length=200, blank=True)
    filiale = models.CharField(max_length=200, blank=True)
    agence = models.CharField(max_length=200, blank=True)
    old_validity_date = models.CharField(max_length=120, blank=True)
    document_validity_date = models.CharField(max_length=120, blank=True)
    match_rate = models.PositiveSmallIntegerField(default=0, validators=[MaxValueValidator(100)])
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="a_valider")
    scan_date = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    notes = models.TextField(blank=True)

    class Meta:
        ordering = ["-scan_date"]
        verbose_name = "Correspondance document expire KYC"
        verbose_name_plural = "Correspondances documents expires KYC"
        constraints = [
            models.UniqueConstraint(
                fields=["client", "document"],
                name="unique_expired_kyc_document_match",
            )
        ]
        indexes = [
            models.Index(fields=["status"]),
            models.Index(fields=["client_code"]),
            models.Index(fields=["filiale", "agence"]),
            models.Index(fields=["scan_date"]),
        ]

    def __str__(self):
        return f"{self.client_code or self.client_id} - {self.document_validity_date}"


class KycDocumentMatchJob(models.Model):
    STATUS_CHOICES = (
        ("pending", "En attente"),
        ("running", "En cours"),
        ("completed", "Termine"),
        ("failed", "Echec"),
    )

    created_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True)
    scope_params = models.JSONField(default=dict, blank=True)
    result = models.JSONField(default=dict, blank=True)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="pending")
    progress_current = models.PositiveIntegerField(default=0)
    progress_total = models.PositiveIntegerField(default=0)
    message = models.CharField(max_length=255, blank=True)
    error = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    started_at = models.DateTimeField(null=True, blank=True)
    completed_at = models.DateTimeField(null=True, blank=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["-created_at"]
        indexes = [
            models.Index(fields=["status", "created_at"]),
            models.Index(fields=["created_by", "created_at"]),
        ]

    @property
    def progress_percent(self):
        if not self.progress_total:
            return 0
        return min(100, int(self.progress_current / self.progress_total * 100))

    def __str__(self):
        return f"Rapprochement documents #{self.pk} - {self.get_status_display()}"


class KycDocumentOcrJob(models.Model):
    """File d'attente OCR d'un lot de documents, traitee par la commande
    `python manage.py process_document_ocr` (cron ou manuelle)."""

    STATUS_CHOICES = (
        ("pending", "En attente"),
        ("running", "En cours"),
        ("completed", "Termine"),
        ("failed", "Echec"),
    )
    MODE_CHOICES = (
        ("files", "Fichiers individuels"),
        ("grouped_pdf", "PDF groupe"),
    )

    import_batch = models.CharField(max_length=120, db_index=True)
    mode = models.CharField(max_length=20, choices=MODE_CHOICES, default="files")
    client_type = models.CharField(max_length=10, choices=CLIENT_TYPE_CHOICES, default="pp")
    document_type = models.CharField(max_length=50, blank=True, default="")
    pages_per_document = models.PositiveSmallIntegerField(default=1)
    grouped_source_file = models.CharField(max_length=255, blank=True, default="",
                                           help_text="Chemin relatif du PDF groupe (mode grouped_pdf).")
    grouped_original_name = models.CharField(max_length=255, blank=True, default="")
    created_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="pending", db_index=True)
    progress_current = models.PositiveIntegerField(default=0)
    progress_total = models.PositiveIntegerField(default=0)
    done_count = models.PositiveIntegerField(default=0)
    failed_count = models.PositiveIntegerField(default=0)
    message = models.CharField(max_length=255, blank=True)
    error = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    started_at = models.DateTimeField(null=True, blank=True)
    completed_at = models.DateTimeField(null=True, blank=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["-created_at"]
        verbose_name = "Job OCR documents KYC"
        verbose_name_plural = "Jobs OCR documents KYC"
        indexes = [
            models.Index(fields=["status", "created_at"]),
            models.Index(fields=["import_batch"]),
        ]

    @property
    def progress_percent(self):
        if not self.progress_total:
            return 100 if self.status == "completed" else 0
        return min(100, int(self.progress_current / self.progress_total * 100))

    def __str__(self):
        return f"OCR lot {self.import_batch} - {self.get_status_display()}"


class KycMatchValidatorRole(models.Model):
    """Organe (profil) autorise a valider / rejeter les correspondances KYC ID."""
    from accounts.models import Organe as _ORGANE_CHOICES
    organe = models.CharField(max_length=50, choices=_ORGANE_CHOICES, unique=True,
                              verbose_name="Organe autorise")
    can_validate = models.BooleanField(default=True, verbose_name="Peut valider")
    can_reject = models.BooleanField(default=True, verbose_name="Peut rejeter")
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Profil validateur (correspondances KYC ID)"
        verbose_name_plural = "Profils validateurs (correspondances KYC ID)"
        ordering = ["organe"]

    def __str__(self):
        return self.organe

    @classmethod
    def user_can_validate(cls, user):
        if not user or not getattr(user, "is_authenticated", False):
            return False
        if user.is_superuser:
            return True
        return cls.objects.filter(organe=getattr(user, "organe", ""), can_validate=True).exists()

    @classmethod
    def user_can_reject(cls, user):
        if not user or not getattr(user, "is_authenticated", False):
            return False
        if user.is_superuser:
            return True
        return cls.objects.filter(organe=getattr(user, "organe", ""), can_reject=True).exists()


class KycScreeningAccess(models.Model):
    """Habilitations par organe sur le module Screening KYC ID (/document-extraction).
    Meme logique que KycMatchValidatorRole : une ligne par organe, editable en admin.
    Si aucune ligne n'existe pour l'organe d'un utilisateur, repli sur le comportement
    historique (DSI / PASS groupe pour le chargement, onglets Resultats + Documents pour tous)."""
    from accounts.models import Organe as _ORGANE_CHOICES
    organe = models.CharField(max_length=50, choices=_ORGANE_CHOICES, unique=True,
                              verbose_name="Organe")
    tab_charger = models.BooleanField(default=False, verbose_name="Onglet Charger")
    tab_suivi = models.BooleanField(default=False, verbose_name="Onglet Suivi")
    tab_resultats = models.BooleanField(default=True, verbose_name="Onglet Resultats")
    tab_sources = models.BooleanField(default=False, verbose_name="Onglet Sources par champ")
    tab_documents = models.BooleanField(default=True, verbose_name="Onglet Documents extraits")
    can_upload_batches = models.BooleanField(default=False, verbose_name="Peut charger des lots")
    can_run_matching = models.BooleanField(default=False, verbose_name="Peut lancer le rapprochement")
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Acces Screening KYC ID (par organe)"
        verbose_name_plural = "Acces Screening KYC ID (par organe)"
        ordering = ["organe"]

    def __str__(self):
        return self.organe

    ALL_PERMS = ("tab_charger", "tab_suivi", "tab_resultats", "tab_sources",
                 "tab_documents", "can_upload_batches", "can_run_matching")

    @classmethod
    def perms_for(cls, user, legacy_can_insert=False):
        """Renvoie {perm: bool} pour l'utilisateur.
        Superuser : tout autorise. Organe sans ligne : repli historique."""
        if user and getattr(user, "is_superuser", False):
            return {p: True for p in cls.ALL_PERMS}
        row = None
        if user and getattr(user, "is_authenticated", False):
            row = cls.objects.filter(organe=getattr(user, "organe", "")).first()
        if row is None:
            return {
                "tab_charger": legacy_can_insert,
                "tab_suivi": legacy_can_insert,
                "tab_resultats": True,
                "tab_sources": legacy_can_insert,
                "tab_documents": True,
                "can_upload_batches": legacy_can_insert,
                "can_run_matching": legacy_can_insert,
            }
        return {p: getattr(row, p) for p in cls.ALL_PERMS}


class SidebarAccess(models.Model):
    from accounts.models import Organe as _ORGANE_CHOICES
    organe = models.CharField(max_length=50, choices=_ORGANE_CHOICES, unique=True,
                              verbose_name="Organe")
    dashboard = models.BooleanField(default=True, verbose_name="Tableau de bord")
    agents_notes = models.BooleanField(default=False, verbose_name="Agents notes")
    champs_non_renseignes = models.BooleanField(default=True, verbose_name="Champs non-renseignes")
    clients_anomalie = models.BooleanField(default=True, verbose_name="Clients en anomalie")
    scoring_clients = models.BooleanField(default=True, verbose_name="Scoring clients")
    screening_kyc = models.BooleanField(default=True, verbose_name="Screening KYC ID")
    nouvelle_notation = models.BooleanField(default=False, verbose_name="Nouvelle notation")
    historique_notation = models.BooleanField(default=False, verbose_name="Historique notation")
    ppe = models.BooleanField(default=False, verbose_name="PPE")
    comptes_specifiques = models.BooleanField(default=False, verbose_name="Comptes specifiques")
    parametrage_utilisateurs = models.BooleanField(default=False, verbose_name="Parametrage utilisateurs")
    regles_qualite = models.BooleanField(default=False, verbose_name="Regles de qualite")
    champs_kyc = models.BooleanField(default=False, verbose_name="Champs KYC")
    documents_screening = models.BooleanField(default=False, verbose_name="Documents Screening KYC ID")
    rappels_scoring = models.BooleanField(default=False, verbose_name="Rappels Scoring")
    pilotage = models.BooleanField(default=False, verbose_name="Pilotage")
    audit = models.BooleanField(default=False, verbose_name="Audit")
    updated_at = models.DateTimeField(auto_now=True)

    ALL_PERMS = (
        "dashboard", "agents_notes", "champs_non_renseignes", "clients_anomalie",
        "scoring_clients", "screening_kyc", "nouvelle_notation", "historique_notation",
        "ppe", "comptes_specifiques", "parametrage_utilisateurs", "regles_qualite",
        "champs_kyc", "documents_screening", "rappels_scoring", "pilotage", "audit",
    )

    class Meta:
        verbose_name = "Acces sidebar (par organe)"
        verbose_name_plural = "Acces sidebar (par organe)"
        ordering = ["organe"]

    def __str__(self):
        return self.organe

    @classmethod
    def legacy_perms_for_organe(cls, organe):
        user_stat = ["Directeur Agence", "Chargé Client"]
        user_stat += ["Chargé Client"]
        notation_org = ["Contrôle Permanent", "PASS"]
        notation_org += ["Contrôle Permanent"]
        conformite_org = ["Conformité", "Conformité Groupe", "PASS"]
        conformite_org += ["Conformité", "Conformité Groupe"]
        admin_org = ["PASS", "DSI", "Conformité", "Qualité", "Contrôle Permanent",
                     "Conformité Groupe", "Contrôle Permanent Groupe"]
        admin_org += ["Conformité", "Qualité", "Contrôle Permanent",
                      "Conformité Groupe", "Contrôle Permanent Groupe"]
        return {
            "dashboard": True,
            "agents_notes": organe not in user_stat,
            "champs_non_renseignes": True,
            "clients_anomalie": True,
            "scoring_clients": True,
            "screening_kyc": True,
            "nouvelle_notation": organe in notation_org,
            "historique_notation": organe in notation_org,
            "ppe": organe in conformite_org,
            "comptes_specifiques": organe in conformite_org,
            "parametrage_utilisateurs": organe in ["PASS", "DSI"],
            "regles_qualite": organe == "PASS",
            "champs_kyc": organe == "PASS",
            "documents_screening": organe == "PASS",
            "rappels_scoring": organe == "PASS",
            "pilotage": organe in admin_org,
            "audit": organe in admin_org,
        }

    @classmethod
    def perms_for(cls, user):
        if user and getattr(user, "is_superuser", False):
            return {p: True for p in cls.ALL_PERMS}
        organe = getattr(user, "organe", "") if user and getattr(user, "is_authenticated", False) else ""
        row = None
        if organe:
            row = cls.objects.filter(organe=organe).first()
            if row is None:
                try:
                    mojibake_organe = organe.encode("utf-8").decode("cp1252")
                except UnicodeError:
                    mojibake_organe = organe
                if mojibake_organe != organe:
                    row = cls.objects.filter(organe=mojibake_organe).first()
        if row is None:
            return cls.legacy_perms_for_organe(organe)
        return {p: getattr(row, p) for p in cls.ALL_PERMS}


class KycMatchDecision(models.Model):
    """Decision de validation d'une correspondance document <-> client KYC.
    Superpose un statut persistant (tracable) sur le resultat de rapprochement."""
    STATUS_CHOICES = (
        ("pending", "A valider"),
        ("validated", "Valide"),
        ("rejected", "Rejete"),
    )

    document = models.ForeignKey(KycDocumentExtraction, on_delete=models.CASCADE, related_name="match_decisions")
    client_type = models.CharField(max_length=10, choices=CLIENT_TYPE_CHOICES, default="pp")
    client_id = models.PositiveIntegerField(help_text="PK du client Kyc_pp / Kyc_pm concerne.")
    client_code = models.CharField(max_length=200, blank=True, default="")
    filiale = models.CharField(max_length=200, blank=True, default="")
    agence = models.CharField(max_length=200, blank=True, default="")
    match_rate = models.PositiveSmallIntegerField(default=0)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="pending", db_index=True)
    decided_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True)
    decided_at = models.DateTimeField(null=True, blank=True)
    notes = models.TextField(blank=True, default="")
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Decision de correspondance KYC ID"
        verbose_name_plural = "Decisions de correspondance KYC ID"
        ordering = ["-updated_at"]
        constraints = [
            models.UniqueConstraint(fields=["document", "client_type", "client_id"],
                                    name="unique_kyc_match_decision"),
        ]
        indexes = [
            models.Index(fields=["status"]),
            models.Index(fields=["client_type", "client_id"]),
            models.Index(fields=["filiale", "agence"]),
        ]

    def __str__(self):
        return f"{self.client_code or self.client_id} - {self.get_status_display()}"







class Person(models.Model):
    username = models.EmailField(blank=True,max_length=30, default='')
    first_name = models.CharField(blank=True,max_length=30, default='')
    last_name = models.CharField(blank=True,max_length=30,  default='')
    email = models.EmailField()
    telephone = models.CharField(blank=True,max_length=20)
    password = models.CharField(blank=True,max_length=32,  default='')
    Photo_profil = models.ImageField(upload_to="media", blank=True, null=True)
    def __str__(self):
        return self.first_name +" "+ self.last_name


class Notation(models.Model):
    agent = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name="notations")
    note = models.CharField(choices=NoteChoices, default="Bien", max_length=15)
    flux_stock = models.CharField(choices=FS, max_length=15, default="")
                                
    recommandation = models.TextField(blank=True, null=True)
    date_notation = models.DateTimeField(default=timezone.now)
    note_par = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE)

    def __str__(self):
        code = getattr(self.agent, "code_expl", "") or getattr(self.agent, "username", "")
        return f"{code} - {self.note} - {(self.recommandation or '')[:20]}..."

class Historique(models.Model):
      agent= models.ForeignKey(settings.AUTH_USER_MODEL,on_delete=models.CASCADE, related_name="historiques")
      notation= models.ForeignKey(Notation,on_delete=models.CASCADE)


class Compte(models.Model):
    TYPE_COMPTE_CHOICES = [
        ('PPE', 'Personne Politiquement Exposée'),
        ('DEV', 'Compte en Devise'),
        ('NON_RES', 'Non Résident'),
        ('scoring', 'Compte scoring'),
    ]
    agent = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name="comptes")
    type_compte = models.CharField(blank=True,max_length=10, choices=TYPE_COMPTE_CHOICES)
    solde = models.DecimalField(max_digits=15, decimal_places=2)
    devise = models.CharField(blank=True,max_length=3)
    date_ouverture = models.DateField()

    def __str__(self):
        return f"{self.agent.email} - {self.type_compte}"


class Kyc_pm(models.Model):
    FILIALE = models.CharField(blank=True,max_length=200)
    AGENCE = models.CharField(blank=True,max_length=200)
    LIB_AGENCE = models.CharField(blank=True,max_length=50)
    EXPL = models.CharField(blank=True,max_length=200)
    CLIENT = models.CharField(blank=True,max_length=200, db_index=True)
    AGEC = models.CharField(blank=True,max_length=200)
    CODAPE = models.CharField(blank=True,max_length=200)
    IDM = models.CharField(blank=True,max_length=200)
    RCSNO = models.CharField(blank=True,max_length=200)
    CAPITAL = models.CharField(blank=True,max_length=200)
    CA = models.CharField(blank=True,max_length=200)
    RESULTAT = models.CharField(blank=True,max_length=200)
    ORIGINE_REV = models.CharField(blank=True,max_length=200)
    DATOUV = models.CharField(blank=True, max_length=200)
    TEL = models.CharField(blank=True,max_length=200)
    DEVISE = models.CharField(blank=True, max_length=200)
    RESID = models.CharField(blank=True, max_length=200)
    DATEREV = models.CharField(blank=True, max_length=200)
    PPE = models.CharField(blank=True, max_length=200)
    RISQUE = models.CharField(blank=True, max_length=200)
    ACTIONNAIRE = models.CharField(blank=True, max_length=4000)
    ADRESSE_SOCIALE = models.CharField(blank=True, max_length=200)
    BOITE_POSTALE = models.CharField(blank=True, max_length=200)
    CONSENT_BIC = models.CharField(blank=True, max_length=200)
    INTITULE_COMPTE = models.CharField(blank=True, max_length=200)
    MANDATAIRE = models.CharField(blank=True, max_length=4000)
    NUMERO_FISCAL = models.CharField(blank=True, max_length=200)
    PAYS_JUR = models.CharField(blank=True, max_length=200)

    class Meta:
                                                                              
                                                                                
        indexes = [
            models.Index(fields=["FILIALE", "AGENCE", "EXPL"], name="kyc_pm_scope_idx"),
            models.Index(fields=["DATEREV"], name="kyc_pm_daterev_idx"),
            models.Index(fields=["RISQUE"], name="kyc_pm_risque_idx"),
                                                                  
            models.Index(fields=["RESID"], name="kyc_pm_resid_idx"),
            models.Index(fields=["DEVISE"], name="kyc_pm_devise_idx"),
        ]

    def __str__(self):
        return f"{self.CLIENT} - {self.FILIALE}"

class Kyc_pp(models.Model):
    FILIALE = models.CharField(blank=True,max_length=200)
    AGENCE = models.CharField(blank=True,max_length=200)
    LIB_AGENCE = models.CharField(blank=True,max_length=50)
    EXPL = models.CharField(blank=True,max_length=200)
    CLIENT = models.CharField(blank=True,max_length=200, db_index=True)
    CODAPE = models.CharField(blank=True,max_length=200)
    IDP = models.CharField(blank=True,max_length=200)
    PAYNAIS = models.CharField(blank=True,max_length=200)
    PROFESSION = models.CharField(blank=True,max_length=200)
    ADRESSE = models.CharField(blank=True,max_length=200)
    PAYS_RESID = models.CharField(blank=True,max_length=200)
    NUMID = models.CharField(blank=True,max_length=200)
    SALAIRE = models.CharField(blank=True,max_length=200)
    ORIGINE_REV = models.CharField(blank=True,max_length=200)
    DATVALID = models.CharField(blank=True,max_length=200)
    DATNAIS = models.CharField(blank=True,max_length=200)
    TEL = models.CharField(blank=True,max_length=200)
    DATOUV = models.CharField(blank=True, max_length=200)
    PPE = models.CharField(blank=True,max_length=200)
    DEVISE = models.CharField(blank=True, max_length=200)
    RESID = models.CharField(blank=True, max_length=200)
    DATEREV = models.CharField(blank=True, max_length=200)
    RISQUE = models.CharField(blank=True, max_length=200)
    BOITE_POSTALE = models.CharField(blank=True, max_length=200)
    CONSENT_BIC = models.CharField(blank=True, max_length=200)
    EMPLOYEUR = models.CharField(blank=True, max_length=200)
    INTITULE_COMPTE = models.CharField(blank=True, max_length=200)
    LIEU_DELIVRANCE_CIN = models.CharField(blank=True, max_length=200)

    class Meta:
                                                                             
                                                                                 
        indexes = [
            models.Index(fields=["FILIALE", "AGENCE", "EXPL"], name="kyc_pp_scope_idx"),
                                                                               
                                 
            models.Index(fields=["DATEREV"], name="kyc_pp_daterev_idx"),
            models.Index(fields=["RISQUE"], name="kyc_pp_risque_idx"),
                                                                  
            models.Index(fields=["PPE"], name="kyc_pp_ppe_idx"),
            models.Index(fields=["RESID"], name="kyc_pp_resid_idx"),
            models.Index(fields=["DEVISE"], name="kyc_pp_devise_idx"),
        ]

    def __str__(self):
        return f"{self.CLIENT} - {self.FILIALE}"

                            
                                                                 

                      
                        

class TauxEvolution(models.Model):
    filiale = models.CharField(blank=True,max_length=10)
    agence   = models.CharField(blank=True,max_length=50,null=True)
    expl     = models.CharField(blank=True,max_length=50)
    date     = models.DateField(blank=True)
    taux     = models.FloatField(blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    flux_stock = models.CharField(blank=True,max_length=50)
    pp_pm = models.CharField(blank=True,max_length=50)

    class Meta:
        ordering = ['filiale', 'expl', 'date']
        indexes = [
                                                                                   
                                                                              
            models.Index(fields=["date"], name="tauxevo_date_idx"),
                                                                         
            models.Index(fields=["filiale", "flux_stock", "expl"], name="tauxevo_scope_idx"),
        ]
    def __str__(self):
        return f"{self.filiale} - {self.expl}"

from django.db import models

class TauxEvolution_filiale(models.Model):
    filiale = models.CharField(blank=True, max_length=10)
    flux_PM = models.FloatField(null=True, blank=True)
    flux_PP = models.FloatField(null=True, blank=True)
    stock_PM = models.FloatField(null=True, blank=True)
    stock_PP = models.FloatField(null=True, blank=True)
    date = models.DateField()
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
                                                                        
        indexes = [
            models.Index(fields=["filiale", "date"], name="tauxevofil_scope_idx"),
            models.Index(fields=["date"], name="tauxevofil_date_idx"),
        ]

    def __str__(self):
        return f"{self.filiale} – {self.date}"


class TauxQualite(models.Model):
    """Table de synthèse du taux de qualité par scope (comme TauxEvolution_filiale
    pour les taux d'évolution). Remplie chaque matin par la commande
    `compute_quality_rates`, lue par le dashboard pour éviter de rescanner Kyc_pp.

    Un scope est identifié par (filiale, agence, expl) ; les champs à None
    représentent un périmètre plus large (expl=None => niveau agence ; agence=None
    => niveau filiale ; filiale=None => niveau Groupe). `applicability` vaut 'PP'
    ou 'PM'. Le taux est calculé exactement comme dans la vue statistiques :
    somme(ok_count) / somme(total) * 100 sur toutes les règles actives du scope.
    """
    FLUX_STOCK_CHOICES = [
        ("stock", "Stock (toute la base)"),
        ("flux", "Flux (nouveaux clients selon DATOUV)"),
    ]

    filiale = models.CharField(max_length=50, null=True, blank=True)
    agence = models.CharField(max_length=50, null=True, blank=True)
    expl = models.CharField(max_length=50, null=True, blank=True)
    applicability = models.CharField(max_length=2)                
                                                                  
                                                                           
                                                                
    flux_stock = models.CharField(max_length=5, choices=FLUX_STOCK_CHOICES, default="stock")
    rate = models.FloatField(default=0)
    ok_count = models.IntegerField(default=0)
    total = models.IntegerField(default=0)
    date = models.DateField()
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
                                                                           
        constraints = [
            models.UniqueConstraint(
                fields=["filiale", "agence", "expl", "applicability", "flux_stock", "date"],
                name="uniq_tauxqualite_scope_jour",
            )
        ]
        indexes = [
            models.Index(fields=["date", "applicability", "filiale", "agence", "expl"]),
            models.Index(fields=["flux_stock", "date"], name="tauxqualite_flux_date_idx"),
        ]

    def __str__(self):
        return f"{self.filiale or 'GROUPE'}/{self.agence or '-'}/{self.expl or '-'} {self.applicability} {self.date}: {self.rate}%"


class QualityFluxConfig(models.Model):
    """Configuration (admin Django) de la fenêtre « flux » du taux de qualité.

    Définit quels clients PP/PM constituent le flux, via leur DATOUV (format ISO
    YYYY-MM-DD) : la dernière journée (veille du calcul) ou le mois calendaire
    précédent. Utilisée par compute_quality_rates pour historiser le taux
    qualité flux dans TauxQualite (flux_stock='flux'), comme les taux de
    complétude le sont dans TauxEvolution.
    """
    WINDOW_CHOICES = [
        ("veille", "Dernière journée (DATOUV = hier)"),
        ("mois", "Mois précédent (DATOUV dans le mois calendaire précédent)"),
    ]

    flux_window = models.CharField(
        max_length=10, choices=WINDOW_CHOICES, default="veille",
        verbose_name="Fenêtre du flux",
        help_text="Clients comptés dans le flux : DATOUV de la veille, ou du mois calendaire précédent.",
    )
    active = models.BooleanField(default=True, verbose_name="Active")
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Configuration flux qualité"
        verbose_name_plural = "Configuration flux qualité"

    def __str__(self):
        return f"Flux qualité : {self.get_flux_window_display()}"


class DATEREV(models.Model):
    FILIALE = models.CharField(blank=True, max_length=10)
    AGENCE = models.CharField(blank=True, max_length=10)
    LIB_AGENCE = models.CharField(blank=True, max_length=50)
    EXPL = models.CharField(blank=True, max_length=10)
    CLIENT = models.CharField(blank=True, max_length=10)
    DATEREV = models.DateField(blank=True, null=True)
    PPE = models.CharField(blank=True, max_length=20)
    RISQUE = models.CharField(blank=True, max_length=20)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=["CLIENT"],
                condition=models.Q(CLIENT__gt=""),
                name="uniq_daterev_client_non_empty",
            ),
        ]

    def __str__(self):
        return f"{self.FILIALE} - {self.EXPL}"



class Profile(models.Model):
    user = models.OneToOneField(settings.AUTH_USER_MODEL, on_delete=models.CASCADE)
    avatar = models.ImageField(default='default.jpg', upload_to='profile_avatars/')
    updated_at = models.DateTimeField(auto_now=True)            

    def __str__(self):
        return f'{self.user.username} Profile'

    def save(self, *args, **kwargs):
        super().save(*args, **kwargs)
        try:
            img = Image.open(self.avatar.path)
            if img.height > 300 or img.width > 300:
                output_size = (300, 300)
                img.thumbnail(output_size)
                img.save(self.avatar.path)
        except Exception:
            pass

class Devise(models.Model):

    filiale =models.CharField(choices=Filiales, max_length=10, default='')
    devise = models.CharField(max_length=4, default='')

    def __str__(self):
        return f"{self.filiale} - {self.devise}"


class EmailReminderConfig(models.Model):
    FREQUENCY_CHOICES = [
        ('manual', 'Manuel uniquement'),
        ('daily',  'Quotidien'),
        ('weekly', 'Hebdomadaire'),
        ('monthly', 'Mensuel'),
    ]

                                                            
    smtp_host     = models.CharField(max_length=200, default='smtp.gmail.com', verbose_name='Serveur SMTP')
    smtp_port     = models.IntegerField(default=587, verbose_name='Port SMTP')
    smtp_user     = models.EmailField(verbose_name='Utilisateur SMTP')
    smtp_password = models.CharField(max_length=300, verbose_name='Mot de passe SMTP')
    smtp_use_tls  = models.BooleanField(default=True, verbose_name='Utiliser TLS')
    smtp_use_ssl  = models.BooleanField(default=False, verbose_name='Utiliser SSL')
    from_email    = models.EmailField(verbose_name='Email expéditeur')
    from_name     = models.CharField(max_length=100, default='KYC Portal BOA', verbose_name='Nom expéditeur')

                                                            
    frequency    = models.CharField(max_length=20, choices=FREQUENCY_CHOICES, default='manual', verbose_name='Fréquence')
    days_before  = models.IntegerField(default=30, verbose_name='Jours avant expiration', help_text='Inclure les clients dont la date de revue (Scoring) expire dans ce nombre de jours')
    active       = models.BooleanField(default=True, verbose_name='Actif')

                                                                           
    notify_emails = models.TextField(
        blank=True, default='', verbose_name="Emails de supervision (rapport quotidien)",
        help_text="Destinataires du rapport d'exécution des tâches quotidiennes, séparés par des virgules / points-virgules / sauts de ligne.")

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Configuration Rappel Scoring"
        verbose_name_plural = "Configurations Rappel Scoring"

    def __str__(self):
        return f"SMTP {self.smtp_host}:{self.smtp_port} — {self.get_frequency_display()}"


class AppreciationConfig(models.Model):
    """Configuration de la campagne d'appréciation globale, propre à chaque filiale
    (une seule et unique configuration possible par filiale).
    La date de démarrage détermine le trimestre courant (trimestres écoulés + 1)."""
    filiale = models.CharField(
        max_length=50, unique=True, default="", verbose_name="Filiale",
        help_text="Code de la filiale (ex. « BOA TG »). Une seule configuration possible par filiale.")
    date_demarrage = models.DateField(
        verbose_name="Date de démarrage de la campagne",
        help_text="Le trimestre courant = nombre de trimestres écoulés depuis cette date + 1 (plafonné à 4).")
    active = models.BooleanField(default=True, verbose_name="Active")
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Configuration Appréciation globale"
        verbose_name_plural = "Configurations Appréciation globale (par filiale)"

    def __str__(self):
        return f"{self.filiale} — démarrée le {self.date_demarrage}"

    def trimestre_actuel(self):
        from kyc.appreciation import current_trimestre
        return current_trimestre(self.date_demarrage)


class Appreciation_globale(models.Model):
    """Appréciation globale calculée par agent (filiale + exploitant), selon la
    matrice du script R appreciation_globale.r."""
    agent = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL,
                              null=True, blank=True, related_name="appreciations")
    filiale = models.CharField(max_length=50, blank=True)
    expl = models.CharField(max_length=50, blank=True)

    trimestre = models.IntegerField(default=1)
    taux_evolution = models.FloatField(null=True, blank=True,
                                       verbose_name="Taux d'évolution (flux, moy. PP+PM)")
    taux_qualite = models.FloatField(null=True, blank=True,
                                     verbose_name="Taux de qualité agent")
    notation = models.CharField(max_length=20, blank=True,
                                verbose_name="Notation flux retenue")
    appreciation_qualite = models.CharField(max_length=30, blank=True)
    appreciation_globale = models.CharField(max_length=20, blank=True)
    mesure = models.TextField(blank=True)
    computed_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Appréciation globale"
        verbose_name_plural = "Appréciations globales"
        unique_together = ("filiale", "expl")
        ordering = ["filiale", "expl"]

    def __str__(self):
        return f"{self.filiale} - {self.expl} : {self.appreciation_globale or '—'}"


class TermTranslation(models.Model):
    """Glossaire de traduction éditable en admin : équivalence FR -> EN d'un terme.
    Prioritaire sur le catalogue gettext lors du rendu en anglais."""
    terme_fr = models.CharField(max_length=300, unique=True, verbose_name="Terme (français)",
                                help_text="Texte français exact tel qu'affiché (ex. « Taux de complétude »).")
    terme_en = models.CharField(max_length=300, verbose_name="Équivalence (anglais)")
    note = models.CharField(max_length=300, blank=True, default='', verbose_name="Note / contexte")
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Traduction de terme (glossaire)"
        verbose_name_plural = "Glossaire des traductions FR → EN"
        ordering = ['terme_fr']

    def __str__(self):
        return f"{self.terme_fr} → {self.terme_en}"

    def save(self, *args, **kwargs):
        super().save(*args, **kwargs)
        from django.core.cache import cache
        cache.delete('term_glossary_en')

    def delete(self, *args, **kwargs):
        super().delete(*args, **kwargs)
        from django.core.cache import cache
        cache.delete('term_glossary_en')

