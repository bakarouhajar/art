import os
import re
import pandas as pd
import numpy as np

# Définition des répertoires
repertoire_entree = './majority_vote'
repertoire_sortie = './preprocessed_majority_vote'

# Création du répertoire de sortie s'il n'existe pas
if not os.path.exists(repertoire_sortie):
    os.makedirs(repertoire_sortie)

# Définition des termes liés aux rôles à supprimer des noms
termes_roles = [
    'victime', 'intimidateur', 'victim_support', 'bully_support', 'conciliateur',
    'Soutien', 'Harceleur', 'Conciliateur', 'Defenseur', 'victim', 'concil'
]

# Vérification du format d'heure
def est_format_heure(s):
    try:
        pd.to_datetime(s, format='%H:%M:%S')
        return True
    except (ValueError, TypeError):
        return False

# Vérification du format de nom
def est_format_nom(s):
    return isinstance(s, str) and not est_format_heure(s)

# Nettoyage de la colonne TIME
def nettoyer_colonne_temps(time_str):
    if isinstance(time_str, str):
        time_str = re.sub(r'\[|\]', '', time_str)  # Suppression des crochets
        time_str = re.sub(r'\s*AM|\s*PM', '', time_str, flags=re.IGNORECASE)  # Suppression de AM/PM
    return time_str

# Nettoyage de la colonne NAME en supprimant les termes liés aux rôles
def nettoyer_nom(name_str):
    if isinstance(name_str, str):
        # Suppression des termes liés aux rôles et des traits d'union ou underscores environnants
        for terme in termes_roles:
            name_str = re.sub(f'[-_]*{terme}[-_]*', '', name_str, flags=re.IGNORECASE)
            name_str = re.sub(f'{terme}', '', name_str, flags=re.IGNORECASE)
        # Suppression des underscores et caractères non alphanumériques en début/fin
        name_str = re.sub(r'[_\s]+$', '', name_str)  # Fin
        name_str = re.sub(r'^[^a-zA-Z0-9]+|[^a-zA-Z0-9]+$', '', name_str)  # Début/Fin non alphanumérique
        return name_str.strip('_')
    return name_str

# Extraction de la partie pertinente du nom de fichier et comptage des occurrences
def renommer_fichiers(repertoire, repertoire_sortie):
    fichiers = [f for f in os.listdir(repertoire) if 'scenario' in f and f.endswith('.xlsx')]
    compteurs = {}

    for fichier in fichiers:
        match = re.match(r'scenario_(.*?)_', fichier)
        if match:
            cle = match.group(1)
            compteurs[cle] = compteurs.get(cle, 0) + 1
            nouveau_nom = f'{cle}_{compteurs[cle]}.xlsx'
            ancien_chemin = os.path.join(repertoire, fichier)
            nouveau_chemin = os.path.join(repertoire_sortie, nouveau_nom)

            # Traitement du fichier
            data = pd.read_excel(ancien_chemin)

            # Remplacement de la colonne ID par des numéros séquentiels commençant à 1
            if 'ID' in data.columns:
                data['ID'] = range(1, len(data) + 1)

            # Nettoyage de la colonne TIME
            data['TIME'] = data['TIME'].apply(nettoyer_colonne_temps)

            # Extraction de la date de TIME et création de la colonne DATE
            def extraire_date(time_str):
                if isinstance(time_str, str) and ' ' in time_str:
                    date, time = time_str.split(' ')
                    return date
                return np.nan

            def extraire_temps(time_str):
                if isinstance(time_str, str) and ' ' in time_str:
                    date, time = time_str.split(' ')
                    return time
                return time_str

            data['DATE'] = data['TIME'].apply(extraire_date)
            data['TIME'] = data['TIME'].apply(extraire_temps)

            # S'assurer que DATE est avant TIME
            cols = list(data.columns)
            index_temps = cols.index('TIME')
            cols.insert(index_temps, cols.pop(cols.index('DATE')))
            data = data[cols]

            # Prétraitement des colonnes TIME et NAME
            temps_dans_nom_count = data['NAME'].apply(est_format_heure).sum()
            nom_dans_temps_count = data['TIME'].apply(est_format_nom).sum()

            if temps_dans_nom_count > len(data) / 2 and nom_dans_temps_count > len(data) / 2:
                data[['TIME', 'NAME']] = data[['NAME', 'TIME']]
            
            data['TIME'] = data['TIME'].apply(lambda x: x.replace('[', '').replace(']', '') if pd.notna(x) else x)
            data['NAME'] = data['NAME'].apply(lambda x: x.replace('@', '').replace('<', '').replace('>', '').strip('_') if pd.notna(x) else x)
            
            # Nettoyage de la colonne NAME en supprimant les termes liés aux rôles
            data['NAME'] = data['NAME'].apply(nettoyer_nom)
            
            data = data.replace(r'^\s*$', np.nan, regex=True)  # Remplacement des chaînes vides par NaN

            # Remplacement de NaN par 'NULL'
            data = data.fillna('NULL')

            # Enregistrement du DataFrame prétraité dans le répertoire de sortie
            data.to_excel(nouveau_chemin, index=False)

            print(f'Renommé : {fichier} en {nouveau_nom}')

renommer_fichiers(repertoire_entree, repertoire_sortie)
print(f"Les fichiers prétraités ont été enregistrés dans le répertoire : {repertoire_sortie}")
