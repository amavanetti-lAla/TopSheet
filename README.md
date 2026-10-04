# Topsheet (Streamlit + Google Sheet)

L'app legge e scrive un Google Sheet. Dalla scheda **Importa** carichi il topsheet .xlsx, dalla scheda **PDF** scarichi il topsheet aggiornato con la stessa struttura del foglio Excel.

## Setup

1. **Google Sheet**: crea un foglio vuoto e copia l'ID dall'indirizzo (la parte tra `/d/` e `/edit`). I fogli Ordini, Brand, Stagioni, Clienti e Log vengono creati dall'app.
2. **Account di servizio**: su console.cloud.google.com crea un progetto, attiva *Google Sheets API* e *Google Drive API*, crea un *Account di servizio* e scarica la chiave JSON.
3. **Condivisione**: condividi il Google Sheet con l'email dell'account di servizio (`...@...iam.gserviceaccount.com`) come *Editor*.
4. **GitHub**: carica questi file in un repository.
5. **Streamlit**: su share.streamlit.io scegli *New app*, seleziona il repository e `app.py`. In *Advanced settings > Secrets* incolla:

```toml
sheet_id = "ID_DEL_TUO_GOOGLE_SHEET"

[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "..."
client_id = "..."
token_uri = "https://oauth2.googleapis.com/token"
```
(copia i valori dal file JSON scaricato)

## Note

- **Senza password**: l'autore delle modifiche è il nome scritto nella barra laterale. Chiunque abbia il link può modificare, quindi imposta l'app come privata (*Share*, inviti per email) finché non c'è un login.
- Le modifiche fatte a mano direttamente nel Google Sheet non vengono registrate nel Log.
- L'import salta gli ordini già presenti (stesso brand, stagione, trade name e città).
- Il PDF replica il layout del foglio per i dati in euro (sezione "EUR ACCOUNTS").
