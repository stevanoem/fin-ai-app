import pandas as pd
import matplotlib.pyplot as plt
import os
import json
import re
from datetime import date, datetime # datetime was imported but not used directly, date is.

from openpyxl.utils import column_index_from_string # get_column_letter was imported but not used

from openai import OpenAI

# ReportLab imports
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle # SimpleDocTemplate not used
# from reportlab.lib.units import mm # mm not used
from reportlab.lib.utils import ImageReader

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
font_path = os.path.join(BASE_DIR, "assets", "LibreBaskerville-Regular.ttf")
pdfmetrics.registerFont(TTFont('LibreBaskerville-Regular', font_path, 'UTF-8'))
font_path = os.path.join(BASE_DIR, "assets", "LibreBaskerville-Bold.ttf")
pdfmetrics.registerFont(TTFont('LibreBaskerville-Bold', font_path, 'UTF-8'))
font_path = os.path.join(BASE_DIR, "assets", "LibreBaskerville-Italic.ttf")
pdfmetrics.registerFont(TTFont('LibreBaskerville-Italic', font_path, 'UTF-8'))

pdfmetrics.registerFontFamily(
    'LibreBaskerville',
    normal='LibreBaskerville-Regular',
    bold='LibreBaskerville-Bold',
    italic='LibreBaskerville-Italic',
    boldItalic='LibreBaskerville-Bold'
)

# --- START OF USER CONFIGURATION ---
# !!! IMPORTANT: SET THESE PATHS TO YOUR LOCAL DIRECTORIES !!!

# Directory where your input Excel files are located
LOCAL_EXCEL_DIR = "input"

# Base directory where output folders (pdf, slike, komentari) will be created
LOCAL_OUTPUT_BASE_DIR = "output"

# Full path to your logo image file
LOCAL_LOGO_FILE = os.path.join(BASE_DIR, "assets", "dmd_logo.png")

# Flag to control AI comment regeneration. Set to True to always regenerate.
REGENERATE_AI_COMMENT = True
# --- END OF USER CONFIGURATION ---


def get_cell_value(df, cell_address):
    col_letter = ''.join(filter(str.isalpha, cell_address))
    row_number = int(''.join(filter(str.isdigit, cell_address))) - 1 # excel krece numeraciju od 1
    col_index = column_index_from_string(col_letter) - 1 # excel krece numeraciju od 1
    return df.iloc[row_number, col_index]

def to_JSON(file_path):
  excel_path = file_path
  sheet_name_kupac = "Kupac" # Assuming the first sheet is "Kupac" or index 0
  
  # Extract filename part for reporting
  filename_part = os.path.basename(file_path)
  try:
    company_name_from_file = filename_part.split("_")[3]
    print(f"Obrada fajla za: {company_name_from_file}")
  except IndexError:
    print(f"Obrada fajla: {filename_part} (Could not extract company name from filename)")


  all_data = {}

  # Tabela 1: Osnovne informacije (E5:F16)
  df1 = pd.read_excel(excel_path, sheet_name=sheet_name_kupac, usecols="E:F", skiprows=3, nrows=12, engine='openpyxl')
  df1.columns = ['Atribut', 'Vrednost']
  all_data["osnovne_informacije"] = df1.to_dict(orient='records')

  # Tabela 2: Promet RSD bez PDV (E19:F47)
  df2 = pd.read_excel(excel_path, sheet_name=sheet_name_kupac, usecols="E:F", skiprows=17, nrows=29, engine='openpyxl')
  df2.columns = ['Atribut', 'Vrednost']
  df2 = df2.dropna(how='all') # ukloni redove gde je sve NaN
  all_data["prometRSD"] = df2.to_dict(orient='records')

  #Tabela 3: Predlog RSD
  df3 = pd.read_excel(excel_path, sheet_name=sheet_name_kupac, usecols="E:F", skiprows=49, nrows=6, engine='openpyxl')
  df3.columns = ['Atribut', 'Vrednost']
  all_data["predlogRSD"] = df3.to_dict(orient='records')

  #Tabela 4: Ocena rizika
  df4 = pd.read_excel(excel_path, sheet_name=sheet_name_kupac, usecols="I:J", skiprows=8, nrows=10, engine='openpyxl')
  df4.columns = ['Atribut', 'Vrednost']
  all_data["ocena_rizika"] = df4.to_dict(orient='records')

  #Tabela 5: Bonitetna ocena
  df5_1 = pd.read_excel(excel_path, sheet_name=sheet_name_kupac, usecols="L:O", skiprows=7, nrows=1, header=1, engine='openpyxl')
  prefix = df5_1.columns[0]
  df5_1 = df5_1.drop(columns=[prefix])
  df5_1.columns = [f"{prefix} {col}" for col in df5_1.columns]

  df5_2 = pd.read_excel(excel_path, sheet_name=sheet_name_kupac, usecols="L:M", skiprows=10, nrows=1, header=None, engine='openpyxl')
  col_name = df5_2.iloc[0,0]
  value = df5_2.iloc[0,1]
  df5_2= pd.DataFrame({col_name: [value]})

  df5 = pd.concat([df5_1, df5_2], axis=1)
  all_data["bonitetna_ocena"] = df5.to_dict(orient='records')

  #Tabela 6: Finansijska analiza
  df6 = pd.read_excel(
      excel_path,
      sheet_name=sheet_name_kupac,
      usecols="I:N",
      skiprows=26,
      header=0,
      nrows=22,
      engine='openpyxl'
  )
  df6.columns.values[0] = "Atribut"
  all_data["finansijska_analizaEUR"] = df6.to_dict(orient='records')

  #Tabela 7: Istorija izmene KL
  df7 = pd.read_excel(
      excel_path,
      sheet_name=sheet_name_kupac,
      usecols="I:K",
      skiprows=52,
      header=0,
      engine='openpyxl'
  )
  if df7.dropna(how='all').empty:
      print("Tabela kreditne istorije je prazna.")
      df7 = df7.dropna(how='all')
  all_data["istorijaKL"] = df7.to_dict(orient='records')

  sheet_name_rezime = "Rezime (EUR)"
  try:
    df8 = pd.read_excel(excel_path, sheet_name=sheet_name_rezime, usecols="B:G", skiprows=3, nrows=30, header=0, engine='openpyxl')
    all_data["rezimeEUR"] = df8.to_dict(orient='records')
  except Exception as e:
    print(f"Nije moguće pročitati list '{sheet_name_rezime}': {e}. Preskačem.")
    all_data["rezimeEUR"] = []


  sheet_name_sporovi = "Sudski sporovi"
  try:
    df9 = pd.read_excel(excel_path, sheet_name=sheet_name_sporovi, engine='openpyxl')
    if df9.empty:
        print("Tabela sudskih sporova je prazna.")
    all_data["Sudski sporovi"] = df9.to_dict(orient='records')
  except Exception as e:
    print(f"Nije moguće pročitati list '{sheet_name_sporovi}': {e}. Preskačem.")
    all_data["Sudski sporovi"] = []

  res = json.loads(json.dumps(all_data, default=str))
  return res

def generate_AIcomment(prompt, key):
  client = OpenAI(api_key=key)
  response = client.chat.completions.create(
      model="gpt-4.1",
      messages=[
          {"role": "user", "content": prompt}
      ],
      temperature=0.0
  )
  return response.choices[0].message.content

def map_cells(god_str):
  # Ensure 'god' is an integer before arithmetic operations
  try:
    god = int(str(god_str).split('.')[0]) # Take integer part if it's like "2023.0"
  except ValueError:
    print(f"Greška: Nije moguće konvertovati godinu '{god_str}' u broj. Koristim 2023 kao podrazumevanu.")
    god = 2023 # Default or raise error

  kupac = {
    "Šifra" : "F5", "PIB" : "F8", "Valuta": "F15", "Tolerancija": "F16",
    "NBS blokada": "J10", "Rizicna lica": "J11",
    "Puštene blanko menice od 2018.godine": "J19", "Broj blanko menica" : "J13",
    "Povlašćenost": "J14", "Bonitetna ocena": "M10", "Ocena rizika": "M11",
    "Sudski sporovi": "J17",
    f"APR Kapital {god-2}": "J28", f"APR Kapital {god-1}": "K28", f"APR Kapital {god}": "L28",
    f"EBITDA {god-2}": "J31", f"EBITDA {god-1}": "K31", f"EBITDA {god}": "L31",
    f"WC {god-2}": "J39", f"WC {god-1}": "K39", f"WC {god}": "L39",
    f"Broj zaposlenih {god-2}": "J48", f"Broj zaposlenih {god-1}": "K48", f"Broj zaposlenih {god}": "L48",
    f"RRL {god-2}": "J41", f"RRL {god-1}": "K41", f"RRL {god}": "L41",
    "Promenjen KL": "J54", "Dug na dan obrade zahteva": "F44", "Dospeli dug": "F45",
    "Broj dana kašnjenja": "F46", "Prosečan broj dana kasnjenja u poslednjih 12m": "F47",
    "Kombinovani kreditni limit": "F53", "Postojeća visina kreditnog limita": "F54",
    "Tražena korekcija kreditnog limita": "F55", "Sporovi": "J17",
    f"Promet {god+1}": "F22", # Assuming current year is god+1 if god is last financial year
    f"Promet {god}": "F21",
    f"{god+1}_Q1": "F31", f"{god+1}_Q2": "F32", f"{god+1}_Q3": "F33", f"{god+1}_Q4": "F34"
  }
  rezime = {
    f"Neto dobitak {god-2}": "E6", f"Neto dobitak {god-1}": "F6", f"Neto dobitak {god}": "G6"
  }
  return (kupac, rezime)

def make_df(file_path):
  firma = os.path.basename(file_path).split("_")[3]
  print(f"Kreiranje df za {firma}...")
  data = {"Filename": firma}

  sheet_names = ["Kupac", "Rezime (EUR)", "Sudski sporovi"]
  df_kupac = pd.read_excel(file_path, sheet_name=sheet_names[0], header=None, engine='openpyxl')
  
  god_val = get_cell_value(df_kupac, "L27") # This is the latest financial year available
  cell_mapping_Kupac, cell_mapping_Rezime = map_cells(god_val)
  
  # kurs
  c_val = get_cell_value(df_kupac, "J4")
  try:
      c = float(c_val)
  except (ValueError, TypeError):
      print(f"Upozorenje: Nije moguće konvertovati kurs '{c_val}' u broj. Koristim 1.0.")
      c = 1.0


  ocena_rizika = ["visok", "umeren", "nizak"]

  #KUPAC
  for feature, cell in cell_mapping_Kupac.items():
        try:
            val = get_cell_value(df_kupac, cell)
            data[feature] = val

            if feature == "Rizicna lica":
                data[feature] = 0 if str(data[feature]).lower() == "nema" else 1
            
            # Ensure 'god' is defined for these checks, using the one from map_cells logic
            god_num = int(str(god_val).split('.')[0]) # Use the same 'god' as in map_cells

            if feature in [f"APR Kapital {god_num-2}", f"APR Kapital {god_num-1}", f"APR Kapital {god_num}",
                           f"EBITDA {god_num-2}", f"EBITDA {god_num-1}", f"EBITDA {god_num}",
                           f"WC {god_num-2}", f"WC {god_num-1}", f"WC {god_num}"]:
                if isinstance(data[feature], (int, float)):
                    data[feature] *= c
                else: # Try to convert if not numeric, else assign 0 or handle error
                    try: data[feature] = float(data[feature]) * c
                    except: data[feature] = 0 


            if feature == "Broj dana kasnjenja":
                if data[feature] == 0 or pd.isna(data[feature]) or str(data[feature]).lower() == "nan":
                    data[feature] = 0

            if feature == "Kombinovani kreditni limit":
                if isinstance(data[feature], (int, float)):
                    data[feature] *= 1.0 # No change, but ensures it's float
                else:
                    try: data[feature] = float(data[feature])
                    except: data[feature] = 0.0


            if feature == "NBS blokada":
                data[feature] = 0 if str(data[feature]).lower() == "nema" else 1
            
            if feature == "Broj blanko menica":
                v_str = str(data[feature])
                if v_str and not pd.isna(data[feature]) and v_str.lower() != 'nan' and len(v_str) > 0 and v_str.strip() != "": # Check if not empty or NaN
                    try: data[feature] = int(float(v_str)) # Convert to float first for "1.0" then int
                    except ValueError: data[feature] = 0 # Or some other default / error handling
                else: data[feature] = 0


            if feature == "Povlascenost":
                data[feature] = 1 if "Povlasceni" in str(data[feature]) else 0

            if feature == "Bonitetna ocena" and not pd.isna(data[feature]):
                data[feature] = str(data[feature])[:2]
            elif feature == "Bonitetna ocena" and pd.isna(data[feature]):
                data[feature] = "N/A"


            if feature == "Ocena rizika" and not pd.isna(data[feature]):
                found_level = "nepoznat"
                for level in ocena_rizika:
                    if level in str(data[feature]).lower():
                        found_level = level
                        break
                data[feature] = found_level
            elif feature == "Ocena rizika" and pd.isna(data[feature]):
                 data[feature] = "nepoznat"

        except Exception as e:
            print(f"Greška pri obradi atributa '{feature}' (ćelija {cell}): {e}. Postavljam na NaN.")
            data[feature] = pd.NA


  #REZIME
  try:
    df_rezime = pd.read_excel(file_path, sheet_name=sheet_names[1], header=None, engine='openpyxl')
    god_num = int(str(god_val).split('.')[0])
    for feature, cell in cell_mapping_Rezime.items():
        try:
            val = get_cell_value(df_rezime, cell)
            data[feature] = val
            if feature in [f"Neto dobitak {god_num-2}", f"Neto dobitak {god_num-1}", f"Neto dobitak {god_num}"]:
                if isinstance(data[feature], (int, float)):
                    data[feature] *= c
                else:
                    try: data[feature] = float(data[feature]) * c
                    except: data[feature] = 0.0
        except Exception as e:
            print(f"Greška pri obradi atributa '{feature}' (ćelija {cell}) u listu Rezime: {e}. Postavljam na NaN.")
            data[feature] = pd.NA
  except Exception as e:
      print(f"Nije moguće pročitati list '{sheet_names[1]}': {e}. Podaci iz rezimea će nedostajati.")
      # Add keys from cell_mapping_Rezime with NA if sheet is missing
      for feature in cell_mapping_Rezime.keys():
          if feature not in data: data[feature] = pd.NA


  #SPOR
  try:
    df_spor = pd.read_excel(file_path, sheet_name=sheet_names[2], engine='openpyxl')
    if df_spor.empty:
        data["Sporovi"] = ""
    else:
        sporovi_texts = []
        for i, r in df_spor.iterrows():
            sporovi_texts.append(f'Ucesnik: {r.get("Učesnik","N/A")}, Datum {r.get("Datum","N/A")}, u iznosu {r.get("Iznos u RSD","N/A")}; ')
        data["Sporovi"] = "".join(sporovi_texts)
  except Exception as e:
      print(f"Nije moguće pročitati list '{sheet_names[2]}': {e}. Podaci o sporovima će nedostajati.")
      if "Sporovi" not in data: # Ensure the key exists from Kupac sheet mapping
          data["Sporovi"] = "N/A zbog greške u čitanju lista"


  df = pd.DataFrame([data])
  return df

def generate_plots(naziv_firme, godine, atribut_values, atribut_naziv, dir_path):
    # Ensure godine and atribut_values are lists of numbers
    godine_num = [int(g) for g in godine]
    atribut_values_num = []
    for v in atribut_values:
        try:
            atribut_values_num.append(float(v))
        except (ValueError, TypeError):
            atribut_values_num.append(0) # Default to 0 if conversion fails

    fig, ax = plt.subplots(figsize=(8, 5))
    for label in ax.get_yticklabels():
        label.set_fontweight('bold')

    ax.plot(godine_num, atribut_values_num, marker='o', label=atribut_naziv)
    ax.set_xlabel('Godina')
    ax.set_ylabel('Iznos (RSD)')
    ax.set_title(f'{atribut_naziv} {naziv_firme}') # Removed slicing from naziv_firme
    ax.legend()
    ax.grid(True)
    ax.set_xticks(godine_num) # Set x-ticks to be the actual years

    plot_filename = f'{atribut_naziv.replace(" ", "_")}_{naziv_firme.replace(" ", "_")}.png'
    full_plot_path = os.path.join(dir_path, plot_filename)
    plt.savefig(full_plot_path, bbox_inches='tight')
    plt.close(fig)
    return full_plot_path

def create_img(dir_path, df, naziv_firme, base_financial_year):
  print(f"Kreiranje grafika za {naziv_firme}...")
  img_paths = []
  
  # base_financial_year is the latest year from APR, e.g., 2023 if data is for 2021, 2022, 2023
  # tekuca godina za kvartale je base_financial_year + 1
  current_processing_year = base_financial_year + 1


  # Financial years for trends
  year1 = base_financial_year - 2 # e.g. 2021
  year2 = base_financial_year - 1 # e.g. 2022
  year3 = base_financial_year     # e.g. 2023
  godine = [year1, year2, year3]

  def get_df_val(key_pattern, year):
      return df.iloc[0].get(key_pattern.format(year), 0) # Default to 0 if key missing

  kapital = [get_df_val("APR Kapital {}", y) for y in godine]
  ebitda = [get_df_val("EBITDA {}", y) for y in godine]
  wc = [get_df_val("WC {}", y) for y in godine]
  # Neto dobitak might have different year naming in map_cells, adjust if needed
  neto_dobitak = [get_df_val("Neto dobitak {}", y) for y in godine]


  img_paths.append(generate_plots(naziv_firme, godine, kapital, "ARP kapital", dir_path))
  img_paths.append(generate_plots(naziv_firme, godine, ebitda, "EBITDA", dir_path))
  img_paths.append(generate_plots(naziv_firme, godine, wc, "WC", dir_path))
  img_paths.append(generate_plots(naziv_firme, godine, neto_dobitak, "Neto Dobitak", dir_path))


  labels = ["Q1", "Q2", "Q3", "Q4"]
  values_kvartali = [
      df.iloc[0].get(f"{current_processing_year}_Q1", 0),
      df.iloc[0].get(f"{current_processing_year}_Q2", 0),
      df.iloc[0].get(f"{current_processing_year}_Q3", 0),
      df.iloc[0].get(f"{current_processing_year}_Q4", 0)
  ]
  values_kvartali_num = []
  for v in values_kvartali:
      try: values_kvartali_num.append(float(v))
      except: values_kvartali_num.append(0)


  plt.figure(figsize=(8, 5))
  plt.bar(labels, values_kvartali_num, color=['blue', 'green', 'orange', 'red'])
  plt.xlabel("Kvartal")
  plt.ylabel("Vrednost")
  plt.title(f"Kvartalne vrednosti za {current_processing_year} ({naziv_firme})")
  plt.grid(axis='y', linestyle='--', alpha=0.7)
  plt.tick_params(axis='y', labelsize=10)
  for label_obj in plt.gca().get_yticklabels(): # Renamed variable to avoid conflict
    label_obj.set_fontweight('bold')

  kvartali_plot_filename = f'Kvartali_{naziv_firme.replace(" ", "_")}.png'
  full_kvartali_plot_path = os.path.join(dir_path, kvartali_plot_filename)
  plt.savefig(full_kvartali_plot_path, bbox_inches='tight')
  img_paths.append(full_kvartali_plot_path)
  plt.close()

  return img_paths

def formatiraj(x):
  if pd.isna(x) or x is None:
      return "N/A"
  try:
    x_num = float(x)
    # Format with Serbian locale (dot for thousands, comma for decimal)
    return f"{x_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
  except (ValueError, TypeError):
    return str(x) # Return original if not number

def shorter_text(com):
    # Lokacija početka "Ukupna procena"
    start_procena = com.find("**Ukupna procena:**")
    # Lokacija početka "Preporuka"
    end_procena = com.find("**Pozitivni indikatori:**")
    start_preporuka = com.find("**Crvene zastavice / anomalije:**")
    # Lokacija kraja "Preporuka" — uzmi prvi naredni pasus nakon toga
    #end_preporuka = com.find("- **Na osnovu kombinovanog", start_preporuka)
    
    # Ako neki od ključnih delova fali, fallback
    if start_procena == -1 or start_preporuka == -1:
        return com  # možeš i `return ""` ako hoćeš prazan string
    
    #if end_preporuka == -1:
     #   preporuka_part = com[start_preporuka:].strip()
    #else:
    #    preporuka_part = com[start_preporuka:end_preporuka].strip()

    procena_part = com[start_procena:end_procena].strip()
    preporuka_part = com[start_preporuka:].strip()

    # Spoji oba dela
    return procena_part + "\n\n" + preporuka_part

def create_pdf(title_firm_name, output_filename, items, kl, image_paths, content_ai, table_data):
    print(f"Formiranje PDF-a za {title_firm_name}")
    c = canvas.Canvas(output_filename, pagesize=letter)
    width, height = letter

    # Logo
    if os.path.exists(LOCAL_LOGO_FILE):
        logo = ImageReader(LOCAL_LOGO_FILE)
        logo_width = 80
        logo_height = 40
        x_pos_logo = width - logo_width - 10
        y_pos_logo = 10
        c.drawImage(logo, x_pos_logo, y_pos_logo, width=logo_width, height=logo_height, mask='auto')
    else:
        print(f"Upozorenje: Logo fajl nije pronađen na putanji: {LOCAL_LOGO_FILE}")


    # Title and Date
    doc_title = "Kreditna analiza za kupca: " + title_firm_name
    c.setFont("LibreBaskerville-Bold", 16)
    c.drawCentredString(width // 2, height - 30, doc_title) # Adjusted y for more space
    today_str = date.today().strftime("%d.%m.%Y")
    c.setFont("LibreBaskerville-Regular", 9)
    c.drawRightString(width - 40, height - 45, f"Datum: {today_str}")

    left_x = 30 # Increased left margin
    right_x = width // 2 + 70 # Adjusted for new left_x
    current_y = height - 70 # Start content lower

    # --- Prva kolona (Opšte informacije, KL, AI Komentar) ---
    
    # Opšte informacije
    c.setFont("LibreBaskerville-Bold", 12)
    c.drawString(left_x, current_y, "Opšte informacije")
    current_y -= 15 # Space after title

    box1_width = (width // 2) + 45 # Adjust width
    box1_height = 110 # Can be dynamic if needed
    
    c.setFillColor(colors.whitesmoke)
    c.rect(left_x, current_y - box1_height, box1_width, box1_height, stroke=0, fill=1)
    c.setStrokeColor(colors.black)
    c.rect(left_x, current_y - box1_height, box1_width, box1_height, stroke=1, fill=0)

    c.setFont("LibreBaskerville-Regular", 9) # Smaller font for items
    c.setFillColor(colors.black)
    text_x_key = left_x + 10
    text_x_val_anchor = left_x + box1_width - 10 # Anchor for right alignment
    text_y_item = current_y - 14

    for k, v_item in items.items():
        v_str = str(v_item if not pd.isna(v_item) else "N/A")
        if k == "Tolerancija" and v_str == "-": continue
        if text_y_item < (current_y - box1_height + 10): break # Don't overflow box
        c.drawString(text_x_key, text_y_item, f"{k}:")
        c.drawRightString(text_x_val_anchor, text_y_item, v_str)
        text_y_item -= 14 # Smaller line height
    current_y -= (box1_height + 20) # Space after box

    # Kreditni limit
    c.setFont("LibreBaskerville-Bold", 12)
    c.drawString(left_x, current_y, "Kreditni limit")
    current_y -= 20
    
    text_x_kl = left_x + 5
    kl_idx = 0
    for k, v_kl in kl.items():
        c.setFillColor(colors.black)
        c.setFont("LibreBaskerville-Bold", 10)
        c.drawString(text_x_kl, current_y, f"{k}")
        c.setFillColor(colors.darkolivegreen if v_kl >=0 else colors.indianred) # Color based on value
        c.setFont("LibreBaskerville-Bold", 11) # Slightly larger
        c.drawString(text_x_kl + 120, current_y, f": {formatiraj(v_kl)} RSD") # Increased offset
        current_y -= 15
        if kl_idx == 0 and len(kl) > 1: # Draw arrow only if there's a next item
            c.setFillColor(colors.black)
            c.drawString(text_x_kl + 10, current_y, "↓")
            current_y -= 5
        kl_idx +=1
    current_y -= 10 # Space after KL

    # AI Komentar
    c.setFont("LibreBaskerville-Bold", 12)
    c.setFillColor(colors.black)
    c.drawString(left_x, current_y, "AI komentar")
    current_y -= 5

    box2_width = box1_width # Same width as info box
    # Calculate height needed for AI comment, or set max
    ai_content_styled = content_ai.replace('\n','<br/>')
    ai_content_styled = ai_content_styled.replace('**Ukupna procena:**', '<b>Ukupna procena:</b>')
    ai_content_styled = ai_content_styled.replace('* **Pozitivni indikatori:**', '<b>Pozitivni indikatori:</b>')
    ai_content_styled = ai_content_styled.replace('* **Ključni faktori rizika:**', '<b>Ključni faktori rizika:</b>')
    ai_content_styled = ai_content_styled.replace('* **Crvene zastavice / anomalije:**', '<b>Crvene zastavice / anomalije:</b>')
    ai_content_styled = ai_content_styled.replace('* **Preporuka:**', '<b>Preporuka:</b>')
    ai_content_styled = ai_content_styled.replace('*','-')
    #print(f"ai_content_styled: /n {ai_content_styled}")

    styles = getSampleStyleSheet()
    ai_comment_style = ParagraphStyle(
        'AIComment', parent=styles['BodyText'], fontSize=9, leading=11,
        textColor=colors.darkslategray, fontName='LibreBaskerville-Regular',
        leftIndent=5, rightIndent=5, spaceBefore=3, spaceAfter=3,
    )
    p_ai = Paragraph(ai_content_styled, ai_comment_style)
    p_ai_w, p_ai_h = p_ai.wrapOn(c, box2_width - 10, 500) # Max height 500, adjust
    
    box2_height = p_ai_h + 10 # Add padding
    max_ai_box_height = 340 # Max height for AI comment box
    box2_height = min(box2_height, max_ai_box_height)


    comment_rect_y = current_y - box2_height
    c.setFillColor(colors.lightgrey)
    c.rect(left_x, comment_rect_y, box2_width, box2_height, fill=1, stroke=0)
    c.setStrokeColor(colors.black)
    c.rect(left_x, comment_rect_y, box2_width, box2_height, stroke=1, fill=0)
    
    p_ai.drawOn(c, left_x + 5, comment_rect_y + box2_height - p_ai_h - 5) # Position from top of box
    current_y = comment_rect_y - 20 # Space after AI comment box


    # --- Druga kolona (Grafikoni) ---
    current_y_col2 = height - 70 # Reset Y for second column
    c.setFont("LibreBaskerville-Bold", 12)
    c.setFillColor(colors.black)
    c.drawString(right_x, current_y_col2, "Grafikoni")
    current_y_col2 -= 5

    img_width_pdf = (width - right_x) - 30 # Available width for images
    img_height_pdf = 130 # Fixed height for consistency
    
    for img_path in image_paths:
        if not os.path.exists(img_path):
            print(f"Upozorenje: Slika nije pronađena: {img_path}")
            # Optionally draw a placeholder text
            c.setFont("LibreBaskerville-Regular", 8)
            c.setFillColor(colors.red)
            c.drawString(right_x + 5, current_y_col2 - (img_height_pdf/2), f"Slika nedostaje: {os.path.basename(img_path)}")
            current_y_col2 -= (img_height_pdf + 10)
            continue

        current_y_col2 -= (img_height_pdf + 10) # Space then image
        if current_y_col2 < 50: # Check if space runs out
             print("Nema dovoljno mesta za sve grafikone na prvoj strani.")
             break # Stop adding images if no space
        try:
            c.drawImage(img_path, right_x + 5, current_y_col2, width=img_width_pdf, height=img_height_pdf, preserveAspectRatio=True)
        except Exception as e:
            print(f"Greška pri crtanju slike {img_path}: {e}")
            c.setFont("LibreBaskerville-Regular", 8)
            c.setFillColor(colors.red)
            c.drawString(right_x + 5, current_y_col2 + (img_height_pdf/2), f"Greška pri crtanju slike.")


    # --- Tabela pozicija (na dnu stranice) ---
    c.setFont("LibreBaskerville-Bold", 12)
    c.setFillColor(colors.black)
    # Position table title above where table will be drawn
    # Table height needs to be estimated or fixed. Let's assume fixed for now.
    table_title_y = 30 + (len(table_data) * 15) + 10 # Estimate: 30 margin + rows*row_height + title_space
    if table_title_y > current_y : # If table overlaps with AI comment, move it down
        table_title_y = current_y - 20 # Place it below the lowest element of col1

    #c.drawString(left_x, table_title_y, "Tabela pozicija")

    # Convert all table data to string, format numbers
    formatted_table_data = []
    for i, row_data in enumerate(table_data):
        new_row = []
        for j, cell_data in enumerate(row_data):
            if i > 0 and j > 0 and j < 4 : # Numeric columns (EUR values)
                 new_row.append(formatiraj(cell_data))
            elif i > 0 and j >=4 : # Percentage columns
                 new_row.append(f"{formatiraj(cell_data)}%")
            else:
                 new_row.append(str(cell_data if not pd.isna(cell_data) else "N/A"))
        formatted_table_data.append(new_row)


    table = Table(formatted_table_data, colWidths=[100, 70, 70, 70, 60, 60]) # Adjust colWidths as needed

    table_styles_list = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'LibreBaskerville-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black)
    ]

    for i_row, row_val in enumerate(table_data):
      if i_row == 0: continue # Skip header
      # Color for percentage change columns
      for j_col_idx in [4, 5]: # Indices of percentage columns
          try:
              val_percent = float(str(row_val[j_col_idx]).replace('%',''))
              cell_color = colors.lightgreen if val_percent > 0 else colors.lightpink if val_percent < 0 else colors.white
              table_styles_list.append(("BACKGROUND", (j_col_idx, i_row), (j_col_idx, i_row), cell_color))
          except ValueError: # Handle if value is not a number
              pass
              
    table.setStyle(TableStyle(table_styles_list))
    
    table_width_pdf, table_height_pdf = table.wrapOn(c, width - 60, 0) # Available width, height is dynamic
    table_x_pdf = (width - table_width_pdf) / 2 # Center table
    table_y_pdf = 30 # Margin from bottom

    # Adjust table_title_y based on actual table_height_pdf
    c.setFillColor(colors.black) # Reset color for title
    c.setFont("LibreBaskerville-Bold", 12)
    c.drawString(left_x, table_y_pdf + table_height_pdf + 5, "Tabela pozicija") # Redraw title correctly

    table.drawOn(c, table_x_pdf, table_y_pdf)
    c.save()

def generate_PDF(output_pdf_path, file_path, ai_com_text):
  df_data = make_df(file_path)
  
  # Extract company name for images and PDF title
  # Ensure this matches how it's used elsewhere if it's a global concept
  naziv_firme_pdf = os.path.basename(file_path).split("_")[3]

  # Determine base financial year from DataFrame (e.g., from a column name like "APR Kapital 2023")
  # This is a bit heuristic; assumes column names like "APR Kapital YYYY" exist
  base_fin_year = 2023 # Default
  for col in df_data.columns:
      if "APR Kapital" in col:
          try:
              year_from_col = int(col.split()[-1])
              base_fin_year = max(base_fin_year, year_from_col)
          except:
              pass
  print(f"Utvrđena bazna finansijska godina za grafikone: {base_fin_year}")


  img_output_base_dir = os.path.join(LOCAL_OUTPUT_BASE_DIR, "slike")
  img_dir_for_firm = os.path.join(img_output_base_dir, naziv_firme_pdf)
  os.makedirs(img_dir_for_firm, exist_ok=True)
  
  img_paths_list = create_img(img_dir_for_firm, df_data, naziv_firme_pdf, base_fin_year)

  row_data = df_data.iloc[0]

  items_dict = {
    "Šifra" : row_data.get("Šifra", "N/A"),
    "PIB" : row_data.get("PIB", "N/A"),
    "Valuta": row_data.get("Valuta", "N/A"),
    "Tolerancija" : row_data.get("Tolerancija", "N/A"),
    "Preduzeće u blokadi" : "Da" if row_data.get("NBS blokada", 0) else "Ne",
    "Bonitetna ocena" : row_data.get("Bonitetna ocena", "N/A"),
    "Ocena rizika" :  row_data.get("Ocena rizika", "N/A") ,
    "Broj menica" : row_data.get("Broj blanko menica", "N/A")
  }

  kl_dict = {
      "Postojeći KL" : row_data.get("Postojeća visina kreditnog limita", 0.0),
      "Traženi KL" : row_data.get("Tražena korekcija kreditnog limita", 0.0)
  }
  
  # Years for table data, relative to base_fin_year
  y1, y2, y3 = base_fin_year - 2, base_fin_year - 1, base_fin_year

  # Exchange rate for EUR conversion (example, should be dynamic or configured)
  eur_rate = 117.0 

  table_data_list = [
      ["Pozicije", f"{y1}", f"{y2}", f"{y3}", f"{str(y2)[-2:]}/{str(y1)[-2:]} (%)", f"{str(y3)[-2:]}/{str(y2)[-2:]} (%)"],
      ["ARP kapital (EUR)", 
       row_data.get(f"APR Kapital {y1}", 0) / eur_rate, 
       row_data.get(f"APR Kapital {y2}", 0) / eur_rate, 
       row_data.get(f"APR Kapital {y3}", 0) / eur_rate, None, None],
      ["EBITDA (EUR)", 
       row_data.get(f"EBITDA {y1}", 0) / eur_rate, 
       row_data.get(f"EBITDA {y2}", 0) / eur_rate, 
       row_data.get(f"EBITDA {y3}", 0) / eur_rate, None, None],
      ["WC (EUR)", 
       row_data.get(f"WC {y1}", 0) / eur_rate, 
       row_data.get(f"WC {y2}", 0) / eur_rate, 
       row_data.get(f"WC {y3}", 0) / eur_rate, None, None],
      ["Racio red. likvidnosti", 
       row_data.get(f"RRL {y1}", 0.0), 
       row_data.get(f"RRL {y2}", 0.0), 
       row_data.get(f"RRL {y3}", 0.0), None, None]
  ]

  for i in range(1, len(table_data_list)):
      x_row = table_data_list[i]
      # Ensure values are numeric before division
      val_y1 = x_row[1] if isinstance(x_row[1], (int, float)) else 0
      val_y2 = x_row[2] if isinstance(x_row[2], (int, float)) else 0
      val_y3 = x_row[3] if isinstance(x_row[3], (int, float)) else 0

      v22_21 = ((val_y2 / val_y1 - 1) * 100) if val_y1 != 0 else (100 if val_y2 > 0 else (-100 if val_y2 < 0 else 0))
      v23_22 = ((val_y3 / val_y2 - 1) * 100) if val_y2 != 0 else (100 if val_y3 > 0 else (-100 if val_y3 < 0 else 0))
      x_row[4] = round(v22_21, 2)
      x_row[5] = round(v23_22, 2)

  create_pdf(naziv_firme_pdf, output_pdf_path, items_dict, kl_dict, img_paths_list, ai_com_text, table_data_list)