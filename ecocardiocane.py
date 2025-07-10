import glob
import os
import pandas as pd
import datetime
import re


SRC_GLOB    = "C:/Users/cross/Downloads/ecocheat/sources/*.xls"      
DB_PATH     = "C:/Users/cross/Downloads/ecocheat/database/Hyper_DB.xlsx"  
DB_SHEET    = "Sheet1"                  
SRC_SHEET   = 0 

#transformers
def format_date_only(x):   
    if pd.isna(x):
        return None
    if isinstance(x, (pd.Timestamp, datetime.datetime)):
        return x.strftime("%d/%m/%Y")
    # otherwise try parsing
    return pd.to_datetime(x).strftime("%d/%m/%Y")

def clean_name(x):
    # strip  whitespace e maiuscole
    if pd.isna(x):
        return None
    return str(x).strip().title()

def to_int(x):
    # convert floats to int
    if pd.isna(x):
        return None
    return int(round(x))

def aorta_ascendente(x):
    if pd.isna(x):
        return None
    s = str(x)
    m = re.search(r":\s*([0-9]+,[0-9]+)", s)
    return m.group(1) if m else s

def extract_emed(x):
    if pd.isna(x):
        return None
    s = str(x)
    m = re.search(r"E' *med\.{0,2}[:\.]?\s*([0-9]+(?:,[0-9]+)?)", s, flags=re.IGNORECASE)
    return m.group(1) if m else s

def extract_elat(x):
    if pd.isna(x):
        return None
    s = str(x)
    m = re.search(r"E' *lat\.{0,2}[:\.]?\s*([0-9]+(?:,[0-9]+)?)", s, flags=re.IGNORECASE)
    return m.group(1) if m else s

def extract_tapse(x):
    if pd.isna(x):
        return None
    s = str(x)
    return s.split(":", 1)[-1].strip()

def extract_hypertrophy_block(df):
    for r in range(51, 60):
        raw = df.iat[r, 0]
        if pd.isna(raw):
            continue
        s = str(raw).lower()
        if "eccentrico" in s:
            return 2
        if "concentrico" in s:
            return 1
    return 0

#coordinate              
CELL_COORDS = [(8, 7),  # data di nascita
               (7, 2),  # cognome nome
                # eta
               (10, 4),  # height
               (10, 2),  # weight
               (6, 7),  # data di oggi
               (15, 2),  # bulbo dimensioni
               (16, 2),  # flusso sistolico
               (15, 4),  # aorta ascendente dimensioni cm
               (22, 5),  # LA area cm2
               (21, 5),  # LAV ml
               (24, 3),  # EDD cm
               (25, 3),  # IVS cm
               (26, 3),  # PW cm
               (24, 5),  # dimensioni V sx sist cm
               (25, 5),  # IVS sist cm
               (26, 5),  # PW sist cm
               (38, 3),  # E cm/s
               (38, 4),  # A cm/s
               (38, 7),  # e' MED CM/S
               (38, 5),  # e' LAT CM/S
               (38, 6),  # e' medio
               (28, 7),  # RWT
               (27, 5),  # LVMI g/mq
               (27, 3),  # LVM (g)
                 # ipertrofia
               (34, 4),  # TAPSE
               # MAPSE
               # GLS
               # ENERGY DISPERSION VFM
               # WALL SHARE STRESS VFM
               (30, 3),  # EF
               # PAPS mmHg
               (21, 3),  # atrio sn dimensioni cm
               (45, 5),  # jet reg tricuspide m/s
               (39, 7),  # E/E'
               # VTI
               # MAPSE/VTI
               # TAPSE/VTI
               (28, 3),  # End diastolic volume
               (29, 3),  # End systolic volume
               # Stenosi aortica
               # AVA cm2
               (16,5) # gradiente massimo mmhg
               # gradiente medio mmhg
               # stroke volume
               # doppio rapporto
]


DB_COLUMNS  = ["DoB (xx/xx/xxxx)", "Name", "Height (cm)", "weight (kg)","Visit_date",
                "Bulbo dimensioni (cm)", "Flusso sistolico (m/s)", "Aorta Ascendente dimensioni (cm)", 
                "LA area (cm2)", "LAV (mL)", "EDD (cm)", "IVS (cm)", "PW (cm)",
                "Diametro V sx sist (cm)", "IVS sist (cm)", "PW sist (cm)",
                "E (cm/s)", "A  (cm/s)", "E'med  (cm/s)", "E' lat  (cm/s)", "E' Medio (cm/s)",
                "RWT", "LVMI (g/mq)", "LVM (g)", "TAPSE (mm)",
                "EF", "Atrio Sx dimensioni (cm)", "Jet reg tricuspide (m/s)", "E/E'.1",
                "End Diastolic Volume", "End Systolic Volume", "Gradiente massimo(mm Hg)" 
               ]  # column names of db


TRANSFORMS = {
    "DoB":        format_date_only,
    "Name":       clean_name,
    "Height":     to_int,
    "weight":     to_int,
    "Aorta Ascendente dimensioni (cm)": aorta_ascendente,
    "E'med  (cm/s)": extract_emed,
    "E' lat  (cm/s)": extract_elat,
    "TAPSE (mm)": extract_tapse,
}

records = []
for src_path in glob.glob(SRC_GLOB):
    ext = os.path.splitext(src_path)[1].lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    df = pd.read_excel(src_path, sheet_name=SRC_SHEET, header=None, engine=engine)

    row = {}
    for out_col, (r, c) in zip(DB_COLUMNS, CELL_COORDS):
        raw = df.iat[r, c]
        #if transform is present use it, if not use raw value
        f = TRANSFORMS.get(out_col, lambda v: v)
        row[out_col] = f(raw)
        row["Hypertrophy (0 = No Hypertrophy; 1= Concentric; 2 = Eccentric)"] = \
        extract_hypertrophy_block(df)
    row["source_file"] = os.path.basename(src_path)
    records.append(row)

#errore
if not records:
    print("No source files found", SRC_GLOB)
    exit()

new_rows = pd.DataFrame.from_records(records)

#se db esiste usalo, senno crea db nuovo
if os.path.exists(DB_PATH):
    db_df = pd.read_excel(DB_PATH, sheet_name=DB_SHEET)
else:
    db_df = pd.DataFrame(columns=DB_COLUMNS + ["source_file"])

# elimina dupes
combined = pd.concat([db_df, new_rows], ignore_index=True)

#replace sheet
with pd.ExcelWriter(DB_PATH, engine="openpyxl", mode="w") as writer:
    combined.to_excel(writer, sheet_name=DB_SHEET, index=False)

print(f"Appended {len(new_rows)} rows to {DB_PATH}/{DB_SHEET}")
