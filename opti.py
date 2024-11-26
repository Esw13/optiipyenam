import csv
import locale
from decimal import Decimal
import os
import gspread
import requests
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


locale.setlocale(locale.LC_ALL, '')

gains_combinaisons = {}
nbreF = int(input("Combien y a-t-il de fichiers ? \n"))
dossier = input("Comment s'appelle le dossier dans lequel sont placés les fichiers ? \n")
nbreP = int(input("Combien y a-t-il de variables ?\n"))

for i in range(1, nbreF + 1):
    filename = os.path.join(dossier, f'a{i}.csv')

    with open(filename, 'r', encoding='ISO-8859-1') as file:
        reader = csv.reader(file, delimiter='\t')
        next(reader)

        for row in reader:
            if len(row) >= nbreP + 9:
                gain = Decimal(row[2].replace('†', '').replace('\xa0', '').replace('$', '').replace(',', ''))
                variables = tuple(float(row[9 + j]) for j in range(nbreP))

                if variables in gains_combinaisons:
                    gains_combinaisons[variables] += gain
                else:
                    gains_combinaisons[variables] = gain


sorted_gains_combinaisons = sorted(gains_combinaisons.items(), key=lambda x: x[1], reverse=True)

top_combinaison = sorted_gains_combinaisons[0][0]

gains_top_combinaison = []

print("Combinaison avec le gain le plus élevé :", sorted_gains_combinaisons[0], "\n")

for i in range(1, nbreF + 1):
    filename = os.path.join(dossier, f'a{i}.csv')
    with open(filename, 'r', encoding='ISO-8859-1') as file:
        reader = csv.reader(file, delimiter='\t')
        next(reader)

        for row in reader:
            if len(row) >= 9 + nbreP:
                cours = row[0]
                gain = Decimal(row[2].replace('†', '').replace('\xa0', '').replace('$', '').replace(',', ''))
                variables = tuple(float(row[9 + j]) for j in range(nbreP))

                if variables == top_combinaison:
                    gains_top_combinaison.append(str(gain/100).replace('.',','))
                    print(f"{cours}, Gain: {gain}, Combinaison: {variables}")

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CREDENTIALS_FILE = 'credentials.json'  # Fichier de crédentiels OAuth 2.0

def column_to_number(column):
    """Convertir une colonne Excel de lettres en un numéro de colonne."""
    column = column.upper()
    col_num = 0
    for char in column:
        col_num = col_num * 26 + (ord(char) - ord('A') + 1)
    return col_num

def authenticate():
    """ Authentification à l'aide d'OAuth 2.0 et gestion des jetons """
    creds = None
    # Charger les jetons à partir du fichier 'token.json', s'il existe
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    # Si les jetons n'existent pas ou sont invalides, effectuer une nouvelle authentification
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Enregistrer les jetons pour les futures exécutions
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds

try:
    # Authentification et création du client gspread
    creds = authenticate()
    client = gspread.authorize(creds)

    # Récupération de l'ID du Google Sheet
    SPREADSHEET_ID = '1dNwO2CyM_qYdZ-SV75AOAx7XOiW4Ldi-AEKltgEV0PU'

    # Demander à l'utilisateur quelle feuille et quelle colonne utiliser
    sheet_name = input("Entrez le nom de la feuille: ")
    column = input("Entrez la colonne (par exemple, A, B, C, ...): ")

    try:
        # Accéder à la feuille spécifiée
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        print(f"La feuille '{sheet_name}' n'a pas été trouvée.")
        print("Voici les feuilles disponibles :")
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        sheets = spreadsheet.worksheets()
        for s in sheets:
            print(s.title)
        exit(1)

    best_gains = gains_top_combinaison

    # Calculer le numéro de colonne à partir de la lettre
    col_num = column_to_number(column)

    # Remplir les données dans la Google Sheet
    for j in range(0, 50): 
        row_index = j + 15 
        sheet.update_cell(row_index, col_num, best_gains[j])
    
    for j in range(50, 98): 
        row_index = j + 17
        sheet.update_cell(row_index, col_num, best_gains[j])

    for j in range(100, 144): 
        row_index = j + 19
        sheet.update_cell(row_index, col_num, best_gains[j])

    print("Données mises à jour dans Google Sheets avec succès.")
except gspread.exceptions.APIError as e:
    print(f"Erreur de l'API Google Sheets : {e}")
except requests.exceptions.RequestException as e:
    print(f"Erreur de requête HTTP : {e}")
except Exception as e:
    print(f"Erreur : {e}")
