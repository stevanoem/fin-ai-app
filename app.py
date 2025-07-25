import os
import sys
import shutil
from datetime import datetime
import pandas as pd
import json
import hashlib

import streamlit as st
import logging

from openai import OpenAI
from openai import OpenAIError

from excel_processor import to_JSON, generate_AIcomment, shorter_text, generate_PDF
from google_drive_utils import upload_drive, google_drive_auth



LOCAL_OUTPUT_BASE_DIR = "output"
LOG_PATH = os.path.join(LOCAL_OUTPUT_BASE_DIR, "app.log")
os.makedirs(LOCAL_OUTPUT_BASE_DIR, exist_ok=True)
API_KEY = st.secrets["api_keys"]["openai"]

def hesiraj_lozinku(lozinka: str) -> str:
    # Pretvaramo lozinku u bajtove
    lozinka_bytes = lozinka.encode('utf-8')
    # Pravimo SHA-256 heš objekat
    sha256 = hashlib.sha256()
    # Dodajemo bajtove lozinke u heš objekat
    sha256.update(lozinka_bytes)
    # Vraćamo heš u heksadecimalnom obliku (string)
    return sha256.hexdigest()



def check_password():
    """ True ako je šifra tačna."""
    st.title("Prijava")
    password = st.text_input("Unesite pristupnu šifru:", type="password")

    if st.button("Potvrdi"):
        if password:
            try:
                # u hex
                entered_password_hex = hesiraj_lozinku(password)
                # Preuzmi tačnu hex šifru iz st.secrets
                correct_password_hex1 = st.secrets["auth1"]["password_hex"]
                correct_password_hex2 = st.secrets["auth2"]["password_hex"]
                correct_password_hex3 = st.secrets["auth3"]["password_hex"]
                correct_password_hex4 = st.secrets["auth4"]["password_hex"]
                correct_password_hex5 = st.secrets["auth5"]["password_hex"]
                correct_password_hex6 = st.secrets["auth6"]["password_hex"]
                correct_password_hex7 = st.secrets["auth7"]["password_hex"]

                if entered_password_hex == correct_password_hex1:
                    st.session_state["authenticated"] = True
                    st.session_state['user'] = 'Emilija'
                    # st.rerun() - stranica se odmah osvezi i prikaze app
                    st.rerun()
                elif entered_password_hex == correct_password_hex2:
                    st.session_state["authenticated"] = True
                    st.session_state['user'] = 'Stefan'
                    # st.rerun() - stranica se odmah osvezi i prikaze app
                    st.rerun()
                elif entered_password_hex == correct_password_hex3:
                    st.session_state["authenticated"] = True
                    st.session_state['user'] = 'Daliborka'
                    # st.rerun() - stranica se odmah osvezi i prikaze app
                    st.rerun()
                elif entered_password_hex == correct_password_hex4:
                    st.session_state["authenticated"] = True
                    st.session_state['user'] = 'Tina'
                    # st.rerun() - stranica se odmah osvezi i prikaze app
                    st.rerun()
                elif entered_password_hex == correct_password_hex5:
                    st.session_state["authenticated"] = True
                    st.session_state['user'] = 'Jovan'
                    # st.rerun() - stranica se odmah osvezi i prikaze app
                    st.rerun()
                elif entered_password_hex == correct_password_hex6:
                    st.session_state["authenticated"] = True
                    st.session_state['user'] = 'Novica'
                    # st.rerun() - stranica se odmah osvezi i prikaze app
                    st.rerun()
                elif entered_password_hex == correct_password_hex7:
                    st.session_state["authenticated"] = True
                    st.session_state['user'] = 'Marko'
                    # st.rerun() - stranica se odmah osvezi i prikaze app
                    st.rerun()
                else:
                    st.error("Pristupna šifra nije tačna.")
            except KeyError:
                st.error("Greška u konfiguraciji: 'auth.password_hex' nije pronađen u secrets.toml.")
                return False
        else:
            st.warning("Molimo unesite šifru.")
    return False

# --- GLAVNI DEO APLIKACIJE ---

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
    st.session_state['user'] = ''

# Ako korisnik nije autorizovan, prikaži ekran za prijavu
if not st.session_state["authenticated"]:
    check_password()
else:
    # --- LOGGING SETTINGS ---

    def initialize_logger():
        logger = logging.getLogger("FinAiApp")

        if not logger.handlers:
            logger.setLevel(logging.INFO)
            logger.propagate = False

            log_formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")

            # Stream handler za prikaz u Streamlit Cloud logovima
            stream_handler = logging.StreamHandler(sys.stdout)
            stream_handler.setFormatter(log_formatter)
            logger.addHandler(stream_handler)

            # File handler za lokalni log koji ćeš slati na Google Drive
            file_handler = logging.FileHandler(LOG_PATH, encoding="utf-8")
            file_handler.setFormatter(log_formatter)
            logger.addHandler(file_handler)

            logger.info("--- Aplikacija pokrenuta, loger konfigurisan ---")

        return logger

    # --- Initialization ---
    if 'logger' not in st.session_state:
        st.session_state['logger'] = initialize_logger()

    logger = st.session_state['logger']
    st.title('Analiza kreditnog rizika')

    # Initialization session state
    if 'current_stage' not in st.session_state:
        st.session_state['current_stage'] = 'waiting_for_file'
        st.session_state['ai_comment'] = ''
        st.session_state['ai_comment_path'] = ''
        st.session_state['pdf_path'] = ''
        st.session_state['client_name'] = ''
        st.session_state['uploaded_file_path'] = ''
        st.session_state['original_file_name'] = ''
        st.session_state['json_content_for_display'] = ''
        st.session_state['timestamp'] = ''
        st.session_state['log_uploaded'] = False
        st.session_state['file_error'] =''
        st.session_state['openai_error']=''
        st.session_state['upload_in_progress'] = False
        logger.info("Session state inicijalizovan. Aplikacija čeka fajl.")

    # --- KONTROLA TOKA APLIKACIJE ---

    # --- FAZA 1: ČEKANJE FAJLA ---
    if st.session_state['current_stage'] == 'waiting_for_file':

        if st.session_state.get('file_error'):
            st.error(st.session_state['file_error'])
            st.session_state['file_error'] = ''
            
        uploaded_file = st.file_uploader(
            "Izaberi Excel fajl",
            type=["xls", "xlsx", "xlsm"]
        )
        if uploaded_file is not None:
            temp_dir = 'temp_uploaded_files'
            st.session_state['timestamp'] = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            os.makedirs(temp_dir, exist_ok=True)
            temp_file_path = os.path.join(temp_dir, st.session_state['timestamp'] +'_'+ st.session_state['user'] + '_' + uploaded_file.name)

            with open(temp_file_path, 'wb') as f:
                f.write(uploaded_file.getbuffer())

            # Update session state
            st.session_state['uploaded_file_path'] = temp_file_path
            st.session_state['original_file_name'] = uploaded_file.name
            st.session_state['current_stage'] = 'file_uploaded'
            
            logger.info(f"Fajl uspešno sačuvan: {uploaded_file.name} na putanji {temp_file_path}")
            
            st.rerun()

# TODO iskljuci dugme za pokretanje analize dok traje ubacivanje na drive
    # --- FAZA 2: FAJL UBAČEN, ČEKA SE ANALIZA ---
    elif st.session_state['current_stage'] == 'file_uploaded':

        if st.session_state.get('openai_error'):
            st.error(st.session_state['openai_error'])
            st.session_state['openai_error'] = ''

        st.success(f"Fajl '{st.session_state['original_file_name']}' je spreman za analizu.")

        if not st.session_state.get('upload_in_progress', False):
            if st.button('Pokreni analizu'):
                st.session_state['upload_in_progress'] = True
                st.rerun()
        else:
            #  upload_in_progress postavljen
            creds = google_drive_auth(logger)
            if creds:
                drive_folder_id = st.secrets["google_drive_folder"]["folder_id"]
                file_id = upload_drive(st.session_state['uploaded_file_path'], creds, drive_folder_id, logger)
                if file_id:
                    st.success(f"Fajl uspešno uploadovan! ID: {file_id}")
                    st.session_state['current_stage'] = 'analysis_in_progress'
                    st.rerun()
                else:
                    st.error("Upload fajla nije uspeo.")
                    st.session_state['upload_in_progress'] = False  # Resetuj
            else:
                st.error("Nije uspela autentifikacija za Google Drive.")
                st.session_state['upload_in_progress'] = False
                

    # --- FAZA 3: ANALIZA U TOKU ---
    elif st.session_state['current_stage'] == 'analysis_in_progress':
        with st.spinner("Analiziram podatke i generišem izveštaj..."):
            try:
                excel_file_path = st.session_state['uploaded_file_path']
                client_name = os.path.basename(st.session_state['original_file_name']).split("_")[3]
                st.session_state['client_name'] = client_name

                logger.info(f"Pokrenuta analiza za klijenta: {client_name}, fajl: {excel_file_path}")

                # --- OVDE IDE VAŠA LOGIKA ZA ANALIZU ---
                json_content_for_ai = to_JSON(excel_file_path)
                logger.info("JSON sadržaj uspešno generisan.")
                #st.write("Prikaz JSON sadržaja:")
                #st.json(json_content_for_ai)
                st.session_state['json_content_for_display'] = json_content_for_ai

                try:
                    client_name_from_json = json_content_for_ai.get("osnovne_informacije", [{}])[0].get("Vrednost", client_name)
                    if isinstance(client_name_from_json, list) and client_name_from_json: # If Naziv Kupca is a list
                        client_name_from_json = client_name_from_json[0]
                    elif not isinstance(client_name_from_json, str) : # if it's not a string (e.g. NaN)
                        client_name_from_json = client_name # fallback to filename derived
                except:
                    client_name_from_json = client_name # fallback

                prompt_text = f"""
                You are an expert Credit Risk Analyst AI. Your task is to analyze the provided JSON data for a client and generate a concise "AI Comment" **in Serbian** for a human credit risk analyst. This comment should highlight key insights, potential risks, positive indicators, and any anomalies relevant to a credit decision. Your language should be professional and direct, avoiding unnecessary jargon explanations or raw data markers in the final comment unless specifically instructed.

                **Input Data:**
                You will receive a JSON object containing various details about the client. Key sections include:
                - `osnovne_informacije`: Basic company information (name, establishment date, ownership, representative).
                - `prometRSD`: Turnover data in RSD (annual, quarterly), planned monthly turnover, current debt ("Dug na dan obrade zahteva" - note: a negative value here, like "29,990.04-", indicates a credit balance or overpayment by the client, which is positive), and average payment delay.
                - `predlogRSD`: Proposed credit limit, existing limit, and justification for changes.
                - `ocena_rizika`: Risk assessment data (NBS blockage, risky persons, disputes, PPL (This stands for "Povezana pravna lica" (related legal entities). If the value is "Ima" (Has/Yes), it indicates the client is linked to such entities; "Nema" (None/No) indicates no such flagged connections. `Status PPL-ova` (e.g., "Aktivan", "u blokadi", "u stečaju") provides crucial context: if the status of PPLs indicates a direct risk (e.g., "u blokadi", "u stečaju", "neaktivan sa dugovanjima"), this should be treated as a key risk factor and listed within the "Ključni faktori rizika" section, prioritized accordingly; if PPLs exist and their status does not indicate a direct risk (e.g., "Aktivan"), their existence should be noted as the last bullet point in the "Ključni faktori rizika" section, ideally prefixed with "Napomena:" (e.g., "Napomena: Klijent ima povezana pravna lica (status: Aktivan).") to distinguish it from direct risk factors.)).
                - `bonitetna_ocena`: Creditworthiness scores (e.g., "E1 - Preduzeće posluje loše i ima veliku verovatnoću neuspeha u budućnosti." indicates poor performance and high failure probability. "E2", "E3" are progressively worse than 'A' or 'B' ratings if those existed). **When referring to the creditworthiness score in the AI Comment, use only the rating code (e.g., 'E1'). However, you MUST use the full description provided with the rating for your internal analysis to understand its implications and severity. Provide the full credit score history.** The "Ocena rizika" here is a numerical score where higher might mean higher risk; interpret this based on context if available.
                - `finansijska_analizaEUR`: Financial analysis data in EUR (Capital, Total Revenue, EBITDA, Net Working Capital, Cash, Receivables, Liabilities, Liquidity Ratios, etc.). Pay close attention to trends (22/21 %, 23/22 %), negative values (e.g., EBITDA, Net Working Capital), and key ratios. ** The year keys in finansijska_analizaEUR  are fixed and do not reflect the actual fiscal years of the financial data. The true (valid) years must be taken from rezimeEUR, where actual fiscal years are explicitly stated.**
                - `rezimeEUR`: Summary of financial data in EUR over several years. **Unlike `finansijska_analizaEUR`, the year labels in `rezimeEUR` represent actual fiscal years. These years must be used to determine the correct temporal context for the data.**
                - `sudski_sporovi`: History of legal disputes, containing information about past and ongoing court cases.

                **Analysis Guidelines:**
                1.  **Overall Assessment:** Start with a brief overall sentiment (e.g., nizak rizik, srednji rizik, visok rizik, značajne zabrinutosti). This assessment should reflect the most critical findings.
                2.  **Positive Indicators:** Identify strengths (e.g., consistent revenue growth, positive net working capital if present, no blockages, overpayment of dues, strong justification for credit if supported by data).
                3.  **Key Risk Factors & Concerns:** Pinpoint weaknesses or areas of concern. Quantify where possible. **Prioritize the risks you list, starting with the most critical ones. Factors like NBS blockages, severe creditworthiness ratings (e.g., 'E' categories), significant negative financial trends (e.g., declining revenue, negative EBITDA, poor liquidity), and substantial legal disputes should generally be considered high priority.** Examples:
                    *   Poor creditworthiness rating (e.g., "E1", "E2", "E3").
                    *   Negative EBITDA or declining profitability.
                    *   Negative or very low net working capital.
                    *   High or increasing debt.
                    *   Significant payment delays.
                    *   Discrepancies between requested credit limit and financial capacity (e.g., large increase requested with poor financials).
                    *   Presence of NBS blockages or legal disputes.
                    *   Low liquidity ratios.
                4.  **Red Flags/Anomalies:** Highlight any unusual data points, inconsistencies, or information that requires immediate attention (e.g., a very high credit limit request despite clear indicators of financial distress, recent establishment with high turnover/requests, missing critical data if observable).
                5.  **Specific Data Points to Consider:**
                    *   Evaluate the `Tražena korekcija kredit limita` (requested credit limit) against the `Postojeća visina kredit limta` (existing limit) and the company's financial health (EBITDA, `bonitetna_ocena` code, `Neto radni kapital`, `Ukupni prihodi`, `Dug na dan obrade zahteva`).
                    *   Comment on the implications of the `bonitetna_ocena` code in your overall analysis, even if only the code is stated in the risk factors.
                    *   Analyze trends in `prometRSD` (turnover) and `finansijska_analizaEUR` (financials like revenue, EBITDA).
                    *   Note the `Dug na dan obrade zahteva` and `Prosečan broj dana kašnjenja`. Compare the latter with the allowed delay tolerance—if it exceeds the tolerance, state so explicitly; if no tolerance is provided ('-'), indicate that the comparison is not applicable.
                    *   Check `ocena_rizika` for blockages (`NBS blokada`) or disputes (`Sporovi u poslednje 3 godine`).
                6.  **Formulating the Recommendation:** The recommendation should be **actionable and provide clear guidance** to the human analyst. Based on the overall risk assessment, suggest concrete next steps, such as approval (with or without conditions like a reduced limit or additional collateral/guarantees), rejection, or the need for specific further investigation (e.g., requesting additional documents, clarifying specific financial items) before a decision can be made.

                **Output Format ("AI Comment"):**
                **IMPORTANT: The entire output comment MUST be in Serbian.**
                Structure your comment clearly:

                **AI komentar kreditnog rizika za [{client_name_from_json}]**

                *   **Ukupna procena:** (e.g., Visok rizik zbog loše bonitetne ocene i negativnog EBITDA...)
                *   **Pozitivni indikatori:**
                    *   (Tačka 1, sa kratkom referencom na podatke, npr., "Nema prijavljenih NBS blokada.")
                    *   (Tačka 2, npr., "Klijent ima preplatu/kreditni saldo od [iznos].")
                *   **Ključni faktori rizika:** (Navesti počevši od najkritičnijih. **Ako PPL postoji i ne predstavlja direktan rizik, navesti ga kao poslednju stavku sa napomenom.**)
                    *   (Tačka 1, npr., "Registrovana NBS blokada u poslednjih godinu dana.")
                    *   (Tačka 2, npr., "Bonitetna ocena: [oznaka_ocene].")
                    *   (Tačka 3, npr., "Značajno negativan EBITDA od [iznos] EUR u 2023.")
                    *   (Tačka 4, npr., "Zahtevano povećanje kreditnog limita sa [postojeći] na [zahtevani] RSD deluje veoma visoko s obzirom na finansijske pokazatelje.")
                    *   (Tačka 5, npr., "Sudski spor [tuženi] [strana] [datum] [iznos]")
                    ** *   (Primer ako je PPL rizik: "Povezano pravno lice [Naziv PPL-a ako je dostupan] je u blokadi/stečaju.")**
                    ** *   ...**
                    ** *   (Poslednja tačka, ako PPL postoji i nije sam po sebi rizik): "Napomena: Klijent ima povezana pravna lica (status: [status_PPL-a, npr. Aktivan])."**
                *   **Crvene zastavice / anomalije:**
                    *   (Tačka 1, ako postoji)
                *   **Preporuka:** (Primeri akcionih preporuka, prilagoditi na osnovu analize)
                    *   (Za visok rizik): "Preporučuje se odbijanje zahteva zbog [ključni razlog 1] i [ključni razlog 2]. Alternativno, ukoliko se izuzetno razmatra odobrenje, neophodno je obezbediti [vrsta dodatne garancije/kolaterala] i smanjiti traženi limit na maksimalno [iznos] RSD."
                    *   (Za visok/srednji rizik sa potrebom za daljom analizom): "Visok/Srednji rizik. Preporučuje se detaljna provera [specifična oblast, npr. strukture potraživanja ili obaveza] i zahtevanje [dodatni dokument, npr. najnovijeg preseka stanja ili biznis plana za naredni period] pre konačne odluke. Razmotriti odobrenje samo uz limit ne veći od [iznos] RSD i pojačan monitoring."
                    *   (Za srednji rizik): "Srednji rizik. Moguće je razmotriti odobrenje traženog limita, ali se savetuje oprez. Predlaže se odobrenje limita od [nešto niži iznos od traženog ili traženi iznos] RSD uz obavezan kvartalni monitoring [ključnog pokazatelja, npr. EBITDA ili likvidnosti]."
                    *   (Za nizak rizik): "Nizak rizik. Finansijski pokazatelji i istorija poslovanja podržavaju zahtev. Preporučuje se odobrenje traženog kreditnog limita od [iznos] RSD uz standardne uslove praćenja."

                    Uz opštu preporuku, navesti i dodatne procene:
                    - **Na osnovu kombinovanog kreditnog limita:** da li traženi iznos premašuje, odgovara ili je značajno ispod vrednosti `Kombinovanog kreditnog limita`.
                    - **Na osnovu prometa:** da li je predlog kreditnog limita opravdan u odnosu na ostvareni godišnji promet (poslednja dostupna godina ili prosek).

                Be factual, objective, and derive your insights directly from the provided JSON data.
                The financial data in `finansijska_analizaEUR` and `rezimeEUR` is in EUR, while `prometRSD` and `predlogRSD` are in RSD. Be mindful of this but focus on the qualitative interpretation and trends unless direct comparison is essential and possible.

                --- START OF CLIENT JSON DATA ---
                {json.dumps(json_content_for_ai, indent=2, ensure_ascii=False)}
                --- END OF CLIENT JSON DATA ---
                """

                ai_comment = generate_AIcomment(prompt_text, API_KEY)
                logger.info("AI komentar uspešno generisan.")

                ai_comment_output_base_dir = os.path.join(LOCAL_OUTPUT_BASE_DIR, "komentari")
                ai_comment_firm_specific_dir = os.path.join(ai_comment_output_base_dir, client_name)
                os.makedirs(ai_comment_firm_specific_dir, exist_ok=True)
                ai_comment_local_file = os.path.join(ai_comment_firm_specific_dir, f'{st.session_state['timestamp'] +'_'+ st.session_state['user'] + '_' +  client_name}_ai_comment.txt')

                with open(ai_comment_local_file, 'w', encoding='utf-8') as f_comment:
                    f_comment.write(ai_comment)

                os.makedirs(os.path.join(LOCAL_OUTPUT_BASE_DIR, 'json'), exist_ok=True)
                json_output_path = os.path.join(LOCAL_OUTPUT_BASE_DIR, 'json', f'{st.session_state['timestamp'] +'_'+ st.session_state['user'] + '_' + client_name}_data_for_ai.json')
                with open(json_output_path, 'w', encoding='utf-8') as json_file:
                    json.dump(json_content_for_ai, json_file, ensure_ascii=False, indent=4)

                st.session_state['ai_comment_path'] = ai_comment_local_file

                
                # --- Upload JSON i AI komentar na Google Drive ---
                creds = google_drive_auth(logger)
                if creds:
                        # Upload JSON (.json)
                        drive_folder_id = st.secrets["google_drive_folder"]["folder_id"]
                        file_id = upload_drive(json_output_path, creds, drive_folder_id, logger)
                        if file_id:
                            st.success(f"Fajl uspešno uploadovan! ID: {file_id}")
                            logger.info(f"JSON uspešno uploadovan na Google Drive. ID: {drive_folder_id }")
                        else:
                            st.error("Upload fajla nije uspeo.")
                        
                        # Upload AI komentara (.txt)
                        drive_folder_id = st.secrets["google_drive_folder"]["folder_id"]
                        file_id = upload_drive(ai_comment_local_file, creds, drive_folder_id, logger)
                        if file_id:
                            st.success(f"Fajl ai kom uspešno uploadovan! ID: {file_id}")
                            logger.info(f"AI komentar uspešno uploadovan na Google Drive. ID: {drive_folder_id }")
                        else:
                            st.error("Upload fajla nije uspeo.")

                            
                else:
                    st.error("Autentifikacija za Google Drive nije uspela. Fajlovi nisu uploadovani.")
                    logger.error("Google Drive autentifikacija nije uspela.")
                    

                #short_ai_text_for_pdf = shorter_text(ai_comment)
                #print(short_ai_text_for_pdf)
                
                #pdf_output_dir = os.path.join('output', 'pdf')
                #os.makedirs(pdf_output_dir, exist_ok=True)
                #pdf_file_path = os.path.join(pdf_output_dir, f'{st.session_state["client_name"]}_kreditna_analiza.pdf')

                #generate_PDF(pdf_file_path, excel_file_path, short_ai_text_for_pdf)
                logger.info(f"TXT uspešno generisan: {ai_comment_local_file}")

                # Saving result
                st.session_state['ai_comment'] = ai_comment
                #st.session_state['pdf_path'] = pdf_file_path
                st.session_state['current_stage'] = 'analysis_done'
            
                st.rerun()

            
            except OpenAIError as oe:
                logger.error(f"Greška prilikom poziva OpenAI API-ja: {oe}")
                st.error("Došlo je do problema sa AI servisom (OpenAI). Pokušajte ponovo kasnije.")
                st.session_state['current_stage'] = 'file_uploaded'
                st.session_state['openai_error'] = 'Došlo je do problema sa AI servisom (OpenAI). Pokušajte ponovo kasnije.'
                st.rerun()

            except (ValueError, KeyError, AttributeError, TypeError, IndexError) as ex:
                logger.error(f"Greška prilikom čitanja fajla: {ex}")
                st.error("Fajl nije u ispravnom formatu. Molimo izaberite ispravan fajl.")
                st.session_state['current_stage'] = 'waiting_for_file'
                st.session_state['file_error'] = "Fajl nije u ispravnom formatu. Molimo izaberite ispravan fajl."
                st.rerun()

            except Exception as e:
                logger.error(f"Neočekivana greška tokom analize: {e}")
                st.error("Došlo je do neočekivane greške tokom analize. Pokušajte ponovo.")
                st.session_state['current_stage'] = 'waiting_for_file'
                st.rerun()

    # --- FAZA 4: ANALIZA ZAVRŠENA, PRIKAZ REZULTATA ---
    elif st.session_state['current_stage'] == 'analysis_done':
        st.header("Rezultati analize")
        st.success("Analiza je uspešno završena!")
        # Ovde prikažite rezultate koje ste sačuvali u session_state
        st.subheader("AI Komentar:")
        st.text_area("Generisani AI Komentar:", st.session_state['ai_comment'], height=300, key="ai_comment_display")

        if not st.session_state.get('log_uploaded'):

            if st.session_state.get('ai_comment_path'):
                creds = google_drive_auth(logger)
                if creds:
                    try:
                        DRIVE_FOLDER_ID = st.secrets["google_drive_folder"]["folder_id"]
                    except KeyError:
                        st.error("Nije pronađen ID Google Drive foldera u secrets.toml!")
                        DRIVE_FOLDER_ID = None

                    if DRIVE_FOLDER_ID:
                        log_temp_path = f"{st.session_state['timestamp'] +'_'+ st.session_state['user']}_app.log"

                    
                        shutil.copy(LOG_PATH, log_temp_path)

                        log_drive_id = upload_drive(log_temp_path, creds, DRIVE_FOLDER_ID, logger)
                        if log_drive_id:
                            logger.info(f"Log fajl uspešno uploadovan na Google Drive. ID: {log_drive_id}")
                        else:
                            st.error("Došlo je do greške prilikom upload-a log fajla na Google Drive.")

                        os.remove(log_temp_path)
                        st.session_state['log_uploaded'] = True

            #Ciscenje log fajla
            open(LOG_PATH, 'w').close()
            logger.info("Log fajl uspešno ispražnjen.")
            
            try:
                # Open the generated PDF file and provide download button
                with open(st.session_state['ai_comment_path'], "rb") as file:
                    btn = st.download_button(
                        label="Preuzmi TXT",
                        data=file,
                        file_name=os.path.basename(st.session_state['ai_comment_path']),
                        mime="application/pdf"
                    )
                
            except FileNotFoundError:
                st.error("PDF fajl nije pronađen. Molimo pokrenite analizu ponovo.")
                logger.error(f"TXT fajl nije pronađen na putanji: {st.session_state.get('ai_comment_path')}")

        st.write(f"Klijent: {st.session_state['client_name']}")
        # st.write(f"Komentar AI: {st.session_state['ai_comment']}")

    
        if st.button("Pokreni novu analizu"):

            #Resetovanje stanja
            st.session_state['current_stage'] = 'waiting_for_file'
            st.session_state['log_uploaded'] = False
            st.session_state['upload_in_progress'] = False 
            logger.info("Pokretanje nove analize.")
            st.rerun()