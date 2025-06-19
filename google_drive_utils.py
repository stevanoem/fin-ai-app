import os
import streamlit as st
import logging

#from google.oauth2.credentials import Credentials
#from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
#from google.auth.transport.requests import Request
from google.oauth2 import service_account

# --- GOOGLE DRIVE SETUP ---

SCOPES = ['https://www.googleapis.com/auth/drive.file']

def google_drive_auth(logger):
    """
    Autentifikacija pomoću Google Service Account kredencijala iz Streamlit Secrets.
    Ovo je ispravan i preporučen način za server-side aplikacije.
    """
    logger = logging.getLogger(__name__) # Dobijamo logger za svaki slučaj

    try:
        # Učitava SVE kredencijale iz secrets sekcije koju smo nazvali [google_service_account]
        # st.secrets vraća rečnik (dictionary) koji savršeno odgovara onome što funkcija očekuje
        creds_json = st.secrets["google_service_account"]
        
        # Kreiramo objekat sa kredencijalima iz učitanog rečnika
        creds = service_account.Credentials.from_service_account_info(
            creds_json, 
            scopes=SCOPES
        )
        
        logger.info("Uspešno kreirani kredencijali pomoću Service Account-a.")
        return creds

    except KeyError:
        # ako sekcija [google_service_account] uopšte ne postoji u secrets
        error_msg = "Greška: Sekcija [google_service_account] nije pronađena u Streamlit Secrets."
        logger.error(error_msg)
        st.error(error_msg)
        st.error("Molimo vas, proverite da li ste ispravno uneli tajne na Streamlit Cloud-u i u lokalnom 'secrets.toml' fajlu.")
        return None

    except Exception as e:
    
        error_msg = f"Greška prilikom učitavanja ili parsiranja Service Account kredencijala: {e}"
        logger.error(error_msg)
        st.error("Došlo je do neočekivane greške pri povezivanju sa Google nalogom.")
        st.error(error_msg)
        return None

def upload_drive(file_path, creds, folder_id):
    """Postavlja fajl sa date putanje na Google Drive."""
    try:
        service = build('drive', 'v3', credentials=creds)
        file_metadata = {'name': os.path.basename(file_path) ,
            'parents': [folder_id] }
        media = MediaFileUpload(file_path)
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id',
            supportsAllDrives=True
        ).execute()
        return file.get('id')
    except Exception as e:
        
        logging.getLogger(__name__).error(f"Greška prilikom upload-a na Google Drive: {e}")
        st.error(f"Greška prilikom upload-a na Google Drive: {e}")
        return None