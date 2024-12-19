## INFORMATICA - Utility per integrazione esami e prova intermedia

**DISCLAIMER:** l'uso del codice è a vostro rischio e pericolo (non garantisco che sia privo di errori e perfettamente adatto ad ogni scenario)



### Istruzioni generali

Il codice permette di combinare un file xlsx con i dati della prova intermedia, con i file xlsx generati da moodle con la correzione e le risposte dell'esame, per generare un file finale nel formato adatto per le registrazioni (con l'unica pecca che per la registrazione serve un file xls, e io riesco a generare solo un xlsx, quindi c'è uno step di conversione manuale)

Il codice funziona con questa logica:

- l'esame contiene una domanda per il ritiro e una per l'uso dei voti della valutazione intermedia (PI)
- nel mio caso, nella VI gli studenti avevano un voto di teoria (su 3 domande) e uno di programmazione. I dati sono contenuti in un file excel, con colonne T1, T2, T3 per le domande di programmazione, e Prog per programmazione
- il codice carica il file con i risultati dell'esame (post correzioni in moodle) e quello delle risposte e poi:
  - verifica se lo studente si è ritirato; in caso contrario:
  - verifica se vuole usare il voto VI di teoria (se disponibile per quello studente) o quello del compito
  - somma il voto di programmazione del compito + eventuali punti extra da VI
  - genera il file nel formato corretto per l'import nel verbale, che deve però essere salvato a mano in un file xls

### Esecuzione

Fondamentalmente, sapendo dove sono i vari file, basta sostituire le constanti nel main...

Per l'estrazione dei file da moodle:

- risultati: MOODLE -> esame -> risultati -> impostare "valutazioni" nel primo box in alto, selezionare tutti gli studenti con checkbox in alto, scaricare come excel
- risposte: MOODLE -> esame -> risultati -> impostare "risposte dettagliate" nel primo box in alto, selezionare tutti gli studenti con checkbox in alto, scaricare come excel

```
    def main():
        # dir da cui recuperare i dati prova intermedia
        DIR_ESONERO = r'xx\Informatica\2024\Esami\Prova intermedia\risultati'
        # dir da cui recuperare i dati esame (e in cui salvare excel per registrazione)
        DIR_ESAME = r'xx\Informatica\2024\Esami\2024.12.19'
        DATA_ESAME = '19/12/2024' # data da stampare in excel registrazione esami
    
        # Creazione del dizionario per esonero
        file_esonero = os.path.join(DIR_ESONERO, 'esonero.xlsx')
        esonero = dati_esonero(file_esonero)
        # Mostro i primi elementi del dizionario per verificarne il contenuto
    
        # Creazione del dizionario per esame
        file_esame_risultati = os.path.join(DIR_ESAME, 'risultati.xlsx')
        file_esame_risposte = os.path.join(DIR_ESAME, 'risposte.xlsx')
        esame = dati_esame(file_esame_risultati, file_esame_risposte)
    
        # genera file registrazione
        file_registrazione = os.path.join(DIR_ESAME, 'registrazione_esame.xlsx')
        crea_file_registrazione(file_registrazione, esonero, esame, DATA_ESAME)
```