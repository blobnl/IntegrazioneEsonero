from math import nan
import pandas as pd
import os
import math


# nNOTA: Conversione dei numeri da formato italiano a float, per il formato di moodle (che a volte è str a volte float)...
def convert_to_float(value):
    try:
        if isinstance(value, str):
            value = value.strip().replace(',', '.')
        return float(value)
    except (ValueError, AttributeError, TypeError):  # se la stringa non è un float, ritorna 0
        return 0.0

# risultati esonero
def dati_esonero(filename):
    
    # Carica i dati dal file Excel -> sheet esonero
    data = pd.ExcelFile(filename)
    # Prendo automaticamente il primo (e unico) foglio
    sheet_name = data.sheet_names[0]
    df = data.parse(sheet_name)
    
    # Dizionario (matricola, dati)
    risultato = {}
    
    for _, row in df.iterrows():
        username = row['Username']
        if username == 'nan' or (isinstance(username, (int,float)) and math.isnan(username) ):
            continue
        risultato[username] = {
            'cognome': row['Cognome'],
            'nome': row['Nome'],
            'ritirato': 'r' in str(row['T1']).lower(),
            'teoria': sum(convert_to_float(row[col]) for col in ['T1', 'T2', 'T3']),
            'prog': convert_to_float(row['Prog'])
        }
    
    return risultato

def is_float(value):
    try:
        float(value.replace(',', '.'))
        return True
    except ValueError:
        return False

# risultati dell'esame
def dati_esame(file_risultati, file_risposte):
    
    # Caricare i risultati dal file Excel -> calcola voto teoria e voto programmazione
    data = pd.ExcelFile(file_risultati)
    # Prendo automaticamente il primo (e unico) foglio
    sheet_name = data.sheet_names[0]
    df_risultati = data.parse(sheet_name)
    
    # Dizionario risultato -> (username, (voto teoria, voto programmazione))
    risultato = {}
    for _, row in df_risultati.iterrows():
        username = row['Username']
        if username == 'nan' or (isinstance(username, (int,float)) and math.isnan(username) ):
            continue
        risultato[username] = {
            'cognome': row['Cognome'],
            'nome': row['Nome'],
            'teoria': sum(convert_to_float(row[col]) for col in ['D. 3 /2,00', 'D. 4 /2,00', 'D. 5 /2,00']),
            'prog': convert_to_float(row['D. 6 /26,00']),
            'ritirato' : False,
            'usa_vi' : False
        }
        
    
    
    # Caricare le risposte dal file Excel -> verifica ritiro e uso voto intermedio
    data = pd.ExcelFile(file_risposte)
    # Prendo automaticamente il primo (e unico) foglio
    sheet_name = data.sheet_names[0]
    df_risposte = data.parse(sheet_name)
    
    #### completare con lettura ritiro e uso voto intermedio
    for _, row in df_risposte.iterrows():
        username = row['Username']
        if username not in risultato:
            print(f'Errore in file risposte: {username} non presente nei risultati')
            continue
        
        risposta_data = str(row.get('Risposta data 1', '')).strip()
        # Verificare se la risposta   diversa da '-' o vuota
        if risposta_data and risposta_data != '-' and (isinstance(risposta_data, (int,float)) and math.isnan(username) ):
            if username in risultato:
                risultato[username]['ritirato'] = True
                print(f"Ritirato: {username} {risultato[username]['nome']} {risultato[username]['cognome']} --> {risposta_data}")
               
     
        # idem per usa valutazione intermedia...
        risposta_data = str(row.get('Risposta data 2', '')).strip().lower()
        # Verificare se la risposta   diversa da '-' o vuota
        if risposta_data == 'vero' or risposta_data == 'true':
            if username in risultato:
                risultato[username]['usa_vi'] = True
                print(f"{username} {username} {risultato[username]['nome']} {risultato[username]['cognome']} usa voto intermedio")
        
    
    return risultato

def crea_file_registrazione(filename, esonero, esame, data_esame):
    # Lista delle righe da salvare -> create elaborando i due dict in ingresso
    righe = []
    for username, dati in esame.items():
        # se username, usa VI, recuperalo da esonero, altrimenti usa quello dell'esame. Se ha voti esonero prog, sommali
        voto = dati['prog'] 
        note = ''
        if not dati['ritirato']:
            voto_teoria = dati['teoria'] 
            if dati['usa_vi']:
                if not username in esonero or esonero[username]['ritirato']:
                    print(f'Errore: {username} {dati["nome"]} {dati["cognome"]} chiede di usare VI ma non ha voto')
                else:
                    voto_teoria = esonero[username]['teoria']
                    note += f'teoria da prova intermedia ({voto_teoria}). '
            voto += voto_teoria
            if username in esonero:
                if esonero[username]['prog'] > 0:
                    print(f'Prog: {username} {dati["nome"]} {dati["cognome"]} punti extra programmazione: {esonero[username]["prog"]}')
                    voto += esonero[username]['prog']
                    note += f'{esonero[username]["prog"]} punti extra prog.'

        if not dati['ritirato'] and round(voto) >= 18:
            voto = round(voto)
            if voto > 31:  # lode per voto > 31, se arrotonda a 31 ma minore è solo 30
                voto = "30L"
            elif voto == 31:
                voto = 30
        else:
            voto = ''
                
        righe.append({
            'Matricola': username,
            'Cognome': dati['cognome'],
            'Nome': dati['nome'],
            'Voto': voto,  
            'Respinto': 'S' if not voto or not dati['ritirato'] and isinstance(voto, (int,float)) and round(voto) < 18 else '',
            'Assente': '',  
            'Ritirato': 'S' if dati['ritirato'] else '',
            'Data esame': data_esame,  
            'Note': note
        })
        
    df = pd.DataFrame(righe)
    df.to_excel(filename, index=False, sheet_name='Sheet1')
    return

'''
ISTRUZIONI:
- salvare in un file xlsx i risultati della prova intermedia (il mio file usa il formato xlsx dei risultati di moodle...)
- correggere tutti i compiti, salvare i risultati in file xlsx
MOODLE -> esame -> risultati -> impostare "valutazioni" nel primo box in alto, selezionare tutti gli studenti con checkbox in alto, scaricare come excel
- esportare il file delle risposte date, sempre in xlsx
MOODLE -> esame -> risultati -> impostare "risposte dettagliate" nel primo box in alto, selezionare tutti gli studenti con checkbox in alto, scaricare come excel

'''


def main():
    # dir da cui recuperare i dati prova intermedia
    DIR_ESONERO = r'G:\Didattica\Informatica\Slides 2024\Esami\Prova intermedia\risultati'
    # dir da cui recuperare i dati esame (e in cui salvare excel per registrazione)
    DIR_ESAME = r'G:\Didattica\Informatica\Slides 2024\Esami\2024.12.19'
    DATA_ESAME = '19/12/2024' # data da stampare in excel registrazione esami


    # Creazione del dizionario per esonero
    file_esonero = os.path.join(DIR_ESONERO, 'esonero.xlsx')
    esonero = dati_esonero(file_esonero)
    
    # Creazione del dizionario per esame
    file_esame_risultati = os.path.join(DIR_ESAME, 'risultati.xlsx')
    file_esame_risposte = os.path.join(DIR_ESAME, 'risposte.xlsx')
    esame = dati_esame(file_esame_risultati, file_esame_risposte)
    
    # genera file registrazione
    file_registrazione = os.path.join(DIR_ESAME, 'registrazione_esame.xlsx')
    crea_file_registrazione(file_registrazione, esonero, esame, DATA_ESAME)

    


main()
