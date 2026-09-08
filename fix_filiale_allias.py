"""
Corrige la filiale des utilisateurs ayant un alias @boa.mg mais rattachés à
« BOA Group » au lieu de « BOA MG ».

Usage (toujours via le venv du projet) :

    venv\\Scripts\\python.exe fix_filiale_boa_mg.py            # aperçu (dry-run)
    venv\\Scripts\\python.exe fix_filiale_boa_mg.py --apply    # applique les changements
"""
import os
import sys

import django

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "Fiabilisation_kyc.settings")
django.setup()

from django.contrib.auth import get_user_model  # noqa: E402

DOMAINE = "@boa.mg"
FILIALE_ERRONEE = "BOA Group"
FILIALE_CIBLE = "BOA MG"


def main(apply_changes: bool) -> None:
    User = get_user_model()

    qs = User.objects.filter(
        email__iendswith=DOMAINE,
        filiale=FILIALE_ERRONEE,
    ).order_by("email")

    total = qs.count()
    if not total:
        print("Aucun utilisateur à corriger.")
        return

    print(f"{total} utilisateur(s) concerné(s) :")
    for user in qs:
        print(f"  - {user.username or user.pk:<20} {user.email:<40} "
              f"{user.filiale} -> {FILIALE_CIBLE}")

    if not apply_changes:
        print("\nMode aperçu : aucun changement écrit. "
              "Relancer avec --apply pour appliquer.")
        return

    updated = qs.update(filiale=FILIALE_CIBLE)
    print(f"\n{updated} utilisateur(s) mis à jour.")


if __name__ == "__main__":
    main(apply_changes="--apply" in sys.argv)
