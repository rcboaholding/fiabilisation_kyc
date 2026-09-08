import csv
import chardet
from django.core.management.base import BaseCommand
from django.contrib.auth import get_user_model

class Command(BaseCommand):
    help = 'Crée des utilisateurs en masse depuis un fichier CSV'

    def add_arguments(self, parser):
        parser.add_argument('bulk_users', type=str, help='/management/commands')

    def handle(self, *args, **kwargs):
        csv_file = kwargs['bulk_users']
        User = get_user_model()

                                             
        with open(csv_file, 'rb') as f:
            result = chardet.detect(f.read())
        encoding = result['encoding']

                                                        
        with open(csv_file, mode='r', encoding=encoding) as file:
            reader = csv.DictReader(file, delimiter=';')
            users_created = 0

            for row in reader:
                                          
                email = (row.get('email') or '').strip()
                username = (row.get('username') or email).strip()
                first_name = (row.get('first_name') or '').strip()
                last_name = (row.get('last_name') or '').strip()
                agence = (row.get('code_agence') or row.get('agence') or '').strip()
                code_expl = (row.get('code_expl') or '').strip()
                organe = (row.get('organe') or '').strip()
                filiale = (row.get('filiale') or '').strip()
                téléphone = (row.get('téléphone') or '').strip()
                password1 = row.get('password1')

                                        
                if not email or not password1:
                    self.stdout.write(self.style.WARNING(f"Utilisateur ignoré. Email ou mot de passe manquant pour l'entrée : {row}"))
                    continue

                                                                   
                user, created = User.objects.get_or_create(
                    email=email,
                    defaults={
                        'username': username,
                        'first_name': first_name,
                        'last_name': last_name,
                        'agence': agence,
                        'code_expl': code_expl,
                        'filiale': filiale,
                        'organe': organe,
                        'téléphone': téléphone,
                    }
                )
                
                if created:
                    user.set_password(password1)                           
                    user.save()
                    users_created += 1

        self.stdout.write(self.style.SUCCESS(f"{users_created} utilisateurs créés avec succès."))