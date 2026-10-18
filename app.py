import base64
import datetime
import os
import shutil
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# Ottiene esattamente la cartella in cui si trova app.py (D:\VOLONTARIATO)
cartella_script = os.path.dirname(os.path.abspath(__file__))

# Percorso completo e sicuro del file Excel principale
DB_FILE = os.path.join(cartella_script, "servizi anc dal 2003.xlsx")

# --- CREAZIONE AUTOMATICA DELLE CARTELLE PER I DOCUMENTI E FOTO ---
anni_cartelle = [2025, 2026, 2027, 2028, "DOCUMENTI VARI", "FOTO"]
for anno in anni_cartelle:
    os.makedirs(str(anno), exist_ok=True)

# Funzione per gestire lo sfondo e la luminosità/opacità
def imposta_sfondo(percorso_immagine, opacita_bianco):
    opacita_sidebar = min(1.0, opacita_bianco + 0.07)
    if os.path.exists(percorso_immagine):
        with open(percorso_immagine, "rb") as f:
            data = f.read()
        encoded = base64.b64encode(data).decode()
        ext = percorso_immagine.split(".")[-1].lower()
        mime = (
            "image/png"
            if ext == "png"
            else "image/jpeg"
            if ext in ["jpg", "jpeg"]
            else "image/webp"
        )
        
        css = f"""
        <style>
        .stApp {{
            background-image: linear-gradient(rgba(255, 255, 255, {opacita_bianco}), rgba(255, 255, 255, {opacita_bianco})), url("data:{mime};base64,{encoded}");
            background-size: cover;
            background-position: center;
            background-repeat: no-repeat;
            background-attachment: fixed;
        }}
        section[data-testid="stSidebar"] {{
            background-color: rgba(255, 255, 255, {opacita_sidebar}) !important;
        }}
        </style>
        """
        st.markdown(css, unsafe_allow_html=True)

# --- FUNZIONE DI CARICAMENTO DATI INTELLIGENTE ---
def carica_dati(sorgente_file=None):
    file_da_usare = sorgente_file if sorgente_file is not None else DB_FILE
    
    if file_da_usare != DB_FILE and not hasattr(file_da_usare, "read"):
        if not os.path.exists(file_da_usare):
            return pd.DataFrame(columns=["Data", "Ora Inizio", "Ora Fine", "Durata", "Tipologia", "Ente", "Descrizione del servizio"])
    
    try:
        xl = pd.ExcelFile(file_da_usare)
        fogli_disponibili = xl.sheet_names
        
        foglio_scelto = fogli_disponibili[0]
        for f in fogli_disponibili:
            if f.strip().lower() == "riepilogo":
                foglio_scelto = f
                break
        df = pd.read_excel(file_da_usare, sheet_name=foglio_scelto, header=0)
        df.columns = df.columns.astype(str).str.strip()
        
        mappa_colonne = {}
        for col in df.columns:
            c_up = col.upper()
            if "DATA" in c_up:
                mappa_colonne[col] = "Data"
            elif "ORA" in c_up and "INIZIO" in c_up:
                mappa_colonne[col] = "Ora Inizio"
            elif "ORA" in c_up and "FINE" in c_up:
                mappa_colonne[col] = "Ora Fine"
            elif "DURATA" in c_up:
                mappa_colonne[col] = "Durata"
            elif "TIPOLOGIA" in c_up:
                mappa_colonne[col] = "Tipologia"
            elif "ENTE" in c_up or "LUOGO" in c_up:
                mappa_colonne[col] = "Ente"
            elif "DESCRIZIONE" in c_up or "NOTE" in c_up:
                mappa_colonne[col] = "Descrizione del servizio"
        df = df.rename(columns=mappa_colonne)
        
        if "Data" not in df.columns and len(df.columns) > 0:
            df = df.rename(columns={df.columns[0]: "Data"})
            
        if "Data" in df.columns:
            def converte_data(val):
                if pd.isna(val):
                    return pd.NaT
                if isinstance(val, (datetime.datetime, datetime.date)):
                    return pd.to_datetime(val)
                if isinstance(val, (int, float)):
                    try:
                        return datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(val))
                    except:
                        return pd.NaT
                val_str = str(val).strip()
                if not val_str or val_str.lower() in ["nan", "nat", "none"]:
                    return pd.NaT
                return pd.to_datetime(val_str, errors='coerce', dayfirst=True)
            df["Data"] = df["Data"].apply(converte_data)
            df = df.dropna(subset=["Data"])
            
        return df
    except Exception as e:
        st.error(f"Errore nella lettura del file: {e}")
        return pd.DataFrame(columns=["Data", "Ora Inizio", "Ora Fine", "Durata", "Tipologia", "Ente", "Descrizione del servizio"])

def salva_dati(df):
    try:
        df.to_excel(DB_FILE, sheet_name="riepilogo", index=False)
    except Exception:
        df.to_csv("associazione_carabinieri_backup.csv", index=False)

def calcola_secondi_totali(durata_val):
    if pd.isna(durata_val):
        return 0
    if isinstance(durata_val, (int, float)):
        return int(durata_val * 3600)
        
    durata_str = str(durata_val).strip().lower()
    try:
        if "h" in durata_str:
            parti = durata_str.split("h")
            ore = int(parti[0].strip()) if parti[0].strip() else 0
            minuti = 0
            secondi = 0
            if len(parti) > 1 and "m" in parti[1]:
                sub_parti = parti[1].split("m")
                minuti = int(sub_parti[0].strip()) if sub_parti[0].strip() else 0
                if len(sub_parti) > 1 and "s" in sub_parti[1]:
                    s_str = sub_parti[1].replace("s", "").strip()
                    secondi = int(s_str) if s_str else 0
            return ore * 3600 + minuti * 60 + secondi
        elif ":" in durata_str:
            parti = durata_str.split(":")
            ore = int(parti[0].strip()) if parti[0].strip() else 0
            minuti = int(parti[1].strip()) if len(parti) > 1 and parti[1].strip() else 0
            secondi = int(parti[2].strip()) if len(parti) > 2 and parti[2].strip() else 0
            return ore * 3600 + minuti * 60 + secondi
        else:
            return int(float(durata_str) * 3600)
    except:
        return 0

def formato_hh_mm_ss(secondi_totali):
    secondi_totali = int(secondi_totali)
    ore = secondi_totali // 3600
    minuti = (secondi_totali % 3600) // 60
    secondi = secondi_totali % 60
    return f"{ore:02d}:{minuti:02d}:{secondi:02d}"

def genera_svg_grafico(riepilogo_anni):
    svg_width = 700
    svg_height = 250
    margin_left = 45
    margin_bottom = 35
    chart_width = svg_width - margin_left - 20
    chart_height = svg_height - margin_bottom - 20
    num_bars = len(riepilogo_anni)
    slot_width = chart_width / max(1, num_bars)
    bar_width = max(10, min(45, slot_width * 0.75))
    righe_svg = [
        f'<svg width="100%" height="{svg_height}" viewBox="0 0 {svg_width} {svg_height}" style="background-color: rgba(255,255,255,0.9); border-radius: 12px; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">'
    ]
    for val in [0, 25, 50, 75, 100]:
        y_pos = 10 + chart_height - (val / 100.0 * chart_height)
        righe_svg.append(
            f'<line x1="{margin_left}" y1="{y_pos}" x2="{svg_width - 15}" y2="{y_pos}" stroke="#f1f5f9" stroke-dasharray="3,3" />'
        )
        righe_svg.append(
            f'<text x="{margin_left - 8}" y="{y_pos + 4}" font-size="10" fill="#94a3b8" text-anchor="end" font-family="sans-serif">{val}</text>'
        )
    for idx, row in riepilogo_anni.reset_index(drop=True).iterrows():
        anno = str(int(row["Anno"]))
        valore = min(max(int(row["Servizi"]), 0), 100)
        x_pos = (
            margin_left
            + 10
            + (idx * slot_width)
            + (slot_width - bar_width) / 2
        )
        b_height = (valore / 100.0) * chart_height
        y_pos = 10 + chart_height - b_height
        righe_svg.append(
            f'<rect x="{x_pos}" y="{y_pos}" width="{bar_width}" height="{b_height}" fill="#3b82f6" rx="4" />'
        )
        righe_svg.append(
            f'<text x="{x_pos + bar_width/2}" y="{y_pos - 6}" font-size="10" font-weight="bold" fill="#1e3a8a" text-anchor="middle" font-family="sans-serif">{valore}</text>'
        )
        righe_svg.append(
            f'<text x="{x_pos + bar_width/2}" y="{svg_height - 10}" font-size="10" fill="#64748b" text-anchor="middle" font-family="sans-serif">{anno}</text>'
        )
    righe_svg.append("</svg>")
    return "\n".join(righe_svg)

st.set_page_config(
    page_title="Storico Servizi Volontariato", page_icon="🤝", layout="wide"
)

# --- BARRA LATERALE ---
try:
    st.sidebar.image("logo.png", width=120)
except:
    st.sidebar.info("💡 Inserisci 'logo.png' nella cartella per il logo.")
    
# --- GESTIONE PASSWORD AMMINISTRATORE & LOGOUT ---
PASSWORD_CORRETTA = "Carabinieri2026"
st.sidebar.markdown("### 🔐 Area Riservata")

# Inizializza lo stato di login se non esiste
if "is_logged_in" not in st.session_state:
    st.session_state["is_logged_in"] = False

# Se non è loggato, mostra il campo password
if not st.session_state["is_logged_in"]:
    password_inserita = st.sidebar.text_input("Password Amministratore", type="password", key="password_admin_input")
    if password_inserita == PASSWORD_CORRETTA:
        st.session_state["is_logged_in"] = True
        st.rerun()
    elif password_inserita != "":
        st.sidebar.error("Password errata!")

is_admin = st.session_state["is_logged_in"]

if is_admin:
    st.sidebar.success("Accesso Admin Effettuato!")
    if st.sidebar.button("🚪 Logout", use_container_width=True):
        st.session_state["is_logged_in"] = False
        st.rerun()

st.sidebar.markdown("### 🎨 Personalizzazione Sfondo")
luminosita_sfondo = st.sidebar.slider(
    "Luminosità / Velatura Sfondo",
    min_value=0.0,
    max_value=1.0,
    value=0.60,
    step=0.02,
)
sfondo_applicato = False
for formato in ["sfondo.jpg", "sfondo.jpeg", "sfondo.png", "sfondo.webp"]:
    percorso_img = os.path.join(cartella_script, formato)
    if os.path.exists(percorso_img):
        imposta_sfondo(percorso_img, luminosita_sfondo)
        sfondo_applicato = True
        break
if not sfondo_applicato:
    imposta_sfondo(os.path.join(cartella_script, "inesistente.jpg"), luminosita_sfondo)
if st.sidebar.button("⬆ Torna all'inizio", use_container_width=True):
    st.session_state["vista_corrente"] = "Panoramica Generale"
    st.rerun()
st.sidebar.divider()

# --- GESTIONE FILE EXCEL (RISERVATA A ADMIN) ---
st.sidebar.markdown("### 📁 Gestione File Excel")
if is_admin:
    file_caricato_utente = st.sidebar.file_uploader("Carica file Excel (.xlsx)", type=["xlsx", "xls", "csv"])
    if file_caricato_utente is not None:
        df_servizi = carica_dati(file_caricato_utente)
    else:
        df_servizi = carica_dati(DB_FILE)
else:
    st.sidebar.info("🔒 Inserisci la password per sbloccare il caricamento file.")
    df_servizi = carica_dati(DB_FILE)
st.sidebar.divider()

# --- INSERISCI NUOVO SERVIZIO (RISERVATO A ADMIN) ---
if is_admin:
    st.sidebar.markdown("### ➕ Inserisci Nuovo Servizio", unsafe_allow_html=True)
    with st.sidebar.form("form_servizio", clear_on_submit=True):
        st.markdown(
            "<p style='font-size: 13px; color: #64748b; margin-bottom: -10px;'>Formato data: <b>GG/MM/AAAA</b></p>",
            unsafe_allow_html=True,
        )
        data_servizio = st.date_input(
            "Data del servizio",
            value=datetime.date.today(),
            min_value=datetime.date(2000, 1, 1),
            max_value=datetime.date.today(),
            format="DD/MM/YYYY",
        )
        c1, c2 = st.columns(2)
        with c1:
            ora_inizio = st.time_input("Ora Inizio", value=datetime.time(8, 0))
        with c2:
            ora_fine = st.time_input("Ora Fine", value=datetime.time(12, 0))
        tipologia_scelta = st.selectbox(
            "Tipologia",
            [
                "Emergenza / Soccorso",
                "Trasporto Sanitario",
                "Attività d'Ufficio / Amministrativa",
                "Formazione",
                "Manifestazioni / Eventi",
                "BLU",
                "Altro",
            ],
        )
        ente = st.text_input("Ente *")
        descrizione = st.text_area("Descrizione del servizio")
        if st.form_submit_button("💾 Salva Servizio", use_container_width=True):
            if ente.strip() == "":
                st.sidebar.error("Inserisci l'Ente")
            else:
                dt_i = datetime.datetime.combine(data_servizio, ora_inizio)
                dt_f = datetime.datetime.combine(data_servizio, ora_fine)
                if dt_f < dt_i:
                    dt_f += datetime.timedelta(days=1)
                diff = (dt_f - dt_i).total_seconds()
                ore_t = int(diff // 3600)
                min_t = int((diff % 3600) // 60)
                nuova_riga = pd.DataFrame({
                    "Data": [pd.to_datetime(data_servizio)],
                    "Ora Inizio": [ora_inizio.strftime("%H:%M")],
                    "Ora Fine": [ora_fine.strftime("%H:%M")],
                    "Durata": [f"{ore_t}h {min_t}m"],
                    "Tipologia": [tipologia_scelta],
                    "Ente": [ente],
                    "Descrizione del servizio": [descrizione],
                })
                df_servizi = pd.concat([df_servizi, nuova_riga], ignore_index=True)
                salva_dati(df_servizi)
                st.sidebar.success("Salvato con successo!")

# --- INTESTAZIONE PRINCIPALE ---
col_logo, col_titolo = st.columns([1, 8])
with col_logo:
    try:
        st.image("logo1.png", width=400)
    except:
        st.markdown("### 🤝")
with col_titolo:
    st.markdown(
        "<h1 style='margin-bottom:0;'>IL MIO VOLONTARIATO IN ASSOCIAZIONE NAZIONALE CARABINIERI</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<p style='color: BLACK; font-size: 20px; margin-top:0;'>Gestione e monitoraggio attività (Dal 2000 in poi)</p>",
        unsafe_allow_html=True,
    )

# --- GESTIONE STATO NAVIGAZIONE ---
if "vista_corrente" not in st.session_state:
    st.session_state["vista_corrente"] = "Panoramica Generale"

if not df_servizi.empty:
    df_servizi["Anno"] = pd.to_datetime(df_servizi["Data"]).dt.year
    df_servizi["Secondi_Totali"] = df_servizi["Durata"].apply(calcola_secondi_totali)
    st.markdown("---")
    
    # --- VISTA PANORAMICA GENERALE ---
    if st.session_state["vista_corrente"] == "Panoramica Generale":
        tot_servizi = len(df_servizi)
        tot_secondi = df_servizi["Secondi_Totali"].sum()
        tot_ore_str = formato_hh_mm_ss(tot_secondi)
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.markdown(
                f"""
                <div style="background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%); padding: 20px; border-radius: 12px; color: white; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                    <h3 style="margin: 0; color: #93c5fd; font-size: 20px;">TOTALE SERVIZI REGISTRATI</h3>
                    <p style="font-size: 36px; font-weight: bold; margin: 5px 0 0 0;">{tot_servizi}</p>
                </div>
            """,
                unsafe_allow_html=True,
            )
        with col_m2:
            st.markdown(
                f"""
                <div style="background: linear-gradient(135deg, #991b1b 0%, #ef4444 100%); padding: 20px; border-radius: 12px; color: white; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                    <h3 style="margin: 0; color: #fecaca; font-size: 20px;">TOTALE ORE DI SERVIZIO</h3>
                    <p style="font-size: 36px; font-weight: bold; margin: 5px 0 0 0;">{tot_ore_str}</p>
                </div>
            """,
                unsafe_allow_html=True,
            )
        st.markdown("<br>", unsafe_allow_html=True)
        
        # --- SEZIONE CONSIGLIERE 2025 2026 2027 2028 E FOTO (RISERVATA ADMIN) ---
        if is_admin:
            st.markdown("### 🏛️ CONSIGLIERE 2025 2026 2027 2028")
            st.markdown(
                "<p style='color: BLACK; font-size: 14px;'>Seleziona un anno per gestire e caricare i documenti o la cartella Foto:</p>",
                unsafe_allow_html=True,
            )
            c_25, c_26, c_27, c_28, c_var, c_foto = st.columns(6)
            with c_25:
                if st.button("📁 Archivio 2025", use_container_width=True):
                    st.session_state["vista_corrente"] = "ARCHIVIO_2025"
                    st.rerun()
            with c_26:
                if st.button("📁 Archivio 2026", use_container_width=True):
                    st.session_state["vista_corrente"] = "ARCHIVIO_2026"
                    st.rerun()
            with c_27:
                if st.button("📁 Archivio 2027", use_container_width=True):
                    st.session_state["vista_corrente"] = "ARCHIVIO_2027"
                    st.rerun()
            with c_28:
                if st.button("📁 Archivio 2028", use_container_width=True):
                    st.session_state["vista_corrente"] = "ARCHIVIO_2028"
                    st.rerun()
            with c_var:
                if st.button("📁 Doc. Vari", use_container_width=True):
                    st.session_state["vista_corrente"] = "ARCHIVIO_DOCUMENTI VARI"
                    st.rerun()
            with c_foto:
                if st.button("📸 Gestione Foto", use_container_width=True):
                    st.session_state["vista_corrente"] = "ARCHIVIO_FOTO"
                    st.rerun()
           
            st.markdown("<br>", unsafe_allow_html=True)
            
        # --- GALLERIA FOTOGRAFICA ---
        st.markdown("<h3 style='text-align: center;'>📸 Galleria Fotografica</h3>", unsafe_allow_html=True)
        cartella_foto = "FOTO"
        estensioni_valide = [".png", ".jpg", ".jpeg", ".webp", ".jfif"]
        if os.path.exists(cartella_foto):
            lista_foto = [f for f in os.listdir(cartella_foto) if os.path.splitext(f)[1].lower() in estensioni_valide]
            
            if lista_foto:
                if "indice_foto" not in st.session_state:
                    st.session_state["indice_foto"] = 0
                    
                if st.session_state["indice_foto"] >= len(lista_foto):
                    st.session_state["indice_foto"] = 0
                    
                foto_corrente = lista_foto[st.session_state["indice_foto"]]
                percorso_foto = os.path.join(cartella_foto, foto_corrente)
                
                st.markdown(
                    """
                    <style>
                    div[data-testid="stColumn"] button {
                        padding: 6px 14px !important;
                        font-size: 13px !important;
                    }
                    </style>
                    """,
                    unsafe_allow_html=True
                )
                col_spazio_sx, col_f1, col_f2, col_f3 = st.columns([0.8, 1.2, 3, 1.2], vertical_alignment="center")
                
                with col_f1:
                    if st.button("◀️ Precedente", key="btn_foto_prev"):
                        st.session_state["indice_foto"] = (st.session_state["indice_foto"] - 1) % len(lista_foto)
                        st.rerun()
                        
                with col_f2:
                    st.image(percorso_foto, caption=f"Foto {st.session_state['indice_foto'] + 1} di {len(lista_foto)}: {foto_corrente}", width=380)
                    
                with col_f3:
                    if st.button("Successiva ▶️", key="btn_foto_next"):
                        st.session_state["indice_foto"] = (st.session_state["indice_foto"] + 1) % len(lista_foto)
                        st.rerun()
            else:
                st.info("💡 La cartella `FOTO` è attualmente vuota. Clicca su '📸 Gestione Foto' in alto per caricare le tue prime fotografie!")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # --- PRIMO EXPANDER: TABELLA RIEPILOGATIVA E GRAFICO PER ANNO ---
        with st.expander("📈 Visualizza / Nascondi Tabella Riepilogativa e Grafico per Anno", expanded=False):
            if is_admin:
                st.markdown("### 📈 Andamento e Riepilogo per Anno")
                riepilogo_anni = (
                    df_servizi.groupby("Anno")
                    .agg(
                        Servizi=("Data", "count"),
                        Secondi_Totali=("Secondi_Totali", "sum"),
                    )
                    .reset_index()
                    .sort_values(by="Anno", ascending=True)
                )
                if not riepilogo_anni.empty:
                    max_servizi_row = riepilogo_anni.loc[
                        riepilogo_anni["Servizi"].idxmax()
                    ]
                    min_servizi_row = riepilogo_anni.loc[
                        riepilogo_anni["Servizi"].idxmin()
                    ]
                    col_minmax1, col_minmax2 = st.columns(2)
                    with col_minmax1:
                        st.metric(
                            "🏆 Anno con più Servizi",
                            f"{int(max_servizi_row['Servizi'])} servizi",
                            f"Anno {int(max_servizi_row['Anno'])}",
                        )
                    with col_minmax2:
                        st.metric(
                            "📌 Anno con meno Servizi",
                            f"{int(min_servizi_row['Servizi'])} servizi",
                            f"Anno {int(min_servizi_row['Anno'])}",
                        )
                col_tab, col_graf = st.columns([1, 1.5])
                with col_tab:
                    st.markdown(
                        "##### Tabella Riepilogativa (Clicca l'anno per i dettagli)"
                    )
                    for _, row in riepilogo_anni.sort_values(
                        by="Anno", ascending=False
                    ).iterrows():
                        anno_val = int(row["Anno"])
                        ore_str = formato_hh_mm_ss(row['Secondi_Totali'])
                        col_btn, col_txt1, col_txt2 = st.columns([1.5, 1, 1])
                        with col_btn:
                            if st.button(
                                f"🔍 Anno {anno_val}",
                                key=f"btn_anno_tab_{anno_val}",
                                use_container_width=True,
                            ):
                                st.session_state["vista_corrente"] = f"ANNO_{anno_val}"
                                st.rerun()
                        with col_txt1:
                            st.markdown(
                                f"<p style='padding-top: 5px; color: #f2f061; font-size: 18px; font-weight: 500;'>{int(row['Servizi'])} serv.</p>",
                                unsafe_allow_html=True,
                            )
                        with col_txt2:
                            st.markdown(
                                f"<p style='padding-top: 5px; color: #f2f061; font-size: 18px; font-weight: 500;'>{ore_str}</p>",
                                unsafe_allow_html=True,
                            )
                with col_graf:
                    st.markdown("##### Grafico Numero Servizi per Anno")
                    svg_code = genera_svg_grafico(riepilogo_anni)
                    components.html(svg_code, height=270)
            else:
                st.info("🔒 *I dettagli completi di tutti i servizi sono riservati. Inserisci la password nella barra laterale per sbloccarli.*")
        st.markdown("---")
        
        # --- SECONDO EXPANDER: TUTTI I SERVIZI REGISTRATI ---
        with st.expander("📄 Visualizza / Nascondi Tutti i Servizi Registrati", expanded=False):
            if is_admin:
                df_mostra_tutti = (
                    df_servizi.drop(columns=["Anno", "Secondi_Totali"])
                    .sort_values(by="Data", ascending=False)
                    .copy()
                )
                df_mostra_tutti["Data"] = df_mostra_tutti["Data"].dt.strftime("%d/%m/%Y")
                st.dataframe(
                    df_mostra_tutti,
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.info("🔒 *I dettagli completi di tutti i servizi sono riservati. Inserisci la password nella barra laterale per sbloccarli.*")
            
    # --- VISTA GESTIONE ARCHIVIO DOCUMENTI E FOTO ---
    elif st.session_state["vista_corrente"].startswith("ARCHIVIO_"):
        anno_archivio = st.session_state["vista_corrente"].replace("ARCHIVIO_", "")
        
        if st.button("⬅️ Torna alla Panoramica Generale", key="ritorna_panoramica_archivio"):
            st.session_state["vista_corrente"] = "Panoramica Generale"
            st.rerun()
        if anno_archivio == "FOTO":
            st.markdown("## 📸 Archivio Fotografico")
            st.markdown("Carica e gestisci le immagini salvate nella cartella **`FOTO`**.")
        else:
            st.markdown(f"## 📁 Archivio Documenti - Anno {anno_archivio}")
            st.markdown(f"Gestione dei file e documenti salvati nella cartella **`{anno_archivio}`**.")
        
        cartella_path = anno_archivio
        
        st.markdown("### 📤 Carica nuovi file")
        files_caricati_doc = st.file_uploader(
            f"Seleziona file da archiviare in {anno_archivio}",
            key=f"uploader_{anno_archivio}",
            accept_multiple_files=True
        )
        
        if files_caricati_doc:
            if st.button("Salva file nella cartella", use_container_width=True):
                for f_caricato in files_caricati_doc:
                    percorso_salvataggio = os.path.join(cartella_path, f_caricato.name)
                    with open(percorso_salvataggio, "wb") as f:
                        f.write(f_caricato.getbuffer())
                st.success(f"🎉 {len(files_caricati_doc)} file salvati con successo nella cartella `{anno_archivio}/`!")
                st.rerun()
        st.markdown("---")
        st.markdown("### 📂 File attualmente presenti nella cartella:")
        
        if os.path.exists(cartella_path):
            file_presenti = os.listdir(cartella_path)
            if file_presenti:
                for nome_file in file_presenti:
                    file_completo = os.path.join(cartella_path, nome_file)
                    if os.path.isfile(file_completo):
                        ext_file = nome_file.split(".")[-1].lower()
                        
                        if anno_archivio == "FOTO" and ext_file in ["png", "jpg", "jpeg", "webp", "jfif"]:
                            col_img, col_nome, col_vedi, col_scarica, col_elimina, col_sposta = st.columns([0.8, 2, 0.8, 0.8, 0.8, 0.8])
                            with col_img:
                                st.image(file_completo, width=60)
                        else:
                            col_nome, col_vedi, col_scarica, col_elimina, col_sposta = st.columns([2.8, 0.8, 0.8, 0.8, 0.8])
                        
                        with col_nome:
                            if anno_archivio != "FOTO" or ext_file not in ["png", "jpg", "jpeg", "webp", "jfif"]:
                                st.markdown(f"📄 **{nome_file}**")
                            else:
                                st.markdown(f"**{nome_file}**")
                        
                        with col_vedi:
                            if ext_file in ["png", "jpg", "jpeg", "webp", "pdf", "txt", "jfif"]:
                                if st.button("👁️ Vedi", key=f"vedi_{anno_archivio}_{nome_file}", use_container_width=True):
                                    st.session_state[f"anteprima_{anno_archivio}_{nome_file}"] = not st.session_state.get(f"anteprima_{anno_archivio}_{nome_file}", False)
                            else:
                                st.markdown("<span style='font-size:12px; color:#94a3b8;'>No anteprima</span>", unsafe_allow_html=True)
                                
                        with col_scarica:
                            with open(file_completo, "rb") as f:
                                st.download_button(
                                    label="📥",
                                    data=f,
                                    file_name=nome_file,
                                    key=f"dl_{anno_archivio}_{nome_file}",
                                    use_container_width=True
                                )
                                
                        with col_elimina:
                            if st.button("🗑️", key=f"del_{anno_archivio}_{nome_file}", use_container_width=True):
                                st.session_state[f"conferma_del_{anno_archivio}_{nome_file}"] = True
                        with col_sposta:
                            with st.popover("🔄", use_container_width=True):
                                tutte_cartelle = [2025, 2026, 2027, 2028, "DOCUMENTI VARI", "FOTO"]
                                destinazioni_disponibili = [str(c) for c in tutte_cartelle if str(c) != str(anno_archivio)]
                                
                                nuova_destinazione = st.selectbox(
                                    "Destinazione:",
                                    destinazioni_disponibili,
                                    key=f"sel_dest_{anno_archivio}_{nome_file}"
                                )
                                
                                if st.button("Conferma", key=f"btn_sposta_{anno_archivio}_{nome_file}", use_container_width=True):
                                    try:
                                        percorso_origine = file_completo
                                        percorso_destinazione = os.path.join(nuova_destinazione, nome_file)
                                        
                                        shutil.move(percorso_origine, percorso_destinazione)
                                        
                                        st.session_state[f"anteprima_{anno_archivio}_{nome_file}"] = False
                                        st.success(f"Spostato in {nuova_destinazione}!")
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"Errore: {e}")
                                        
                        if st.session_state.get(f"conferma_del_{anno_archivio}_{nome_file}", False):
                            st.warning(f"Vuoi davvero eliminare '{nome_file}'?")
                            col_conf1, col_conf2 = st.columns(2)
                            with col_conf1:
                                if st.button("Sì, elimina", key=f"sì_del_{anno_archivio}_{nome_file}", use_container_width=True):
                                    try:
                                        os.remove(file_completo)
                                        st.session_state[f"anteprima_{anno_archivio}_{nome_file}"] = False
                                        st.session_state[f"conferma_del_{anno_archivio}_{nome_file}"] = False
                                        st.success(f"File '{nome_file}' eliminato!")
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"Errore: {e}")
                            with col_conf2:
                                if st.button("Annulla", key=f"no_del_{anno_archivio}_{nome_file}", use_container_width=True):
                                    st.session_state[f"conferma_del_{anno_archivio}_{nome_file}"] = False
                                    st.rerun()
                                    
                        if st.session_state.get(f"anteprima_{anno_archivio}_{nome_file}", False):
                            if st.button("❌ Chiudi anteprima", key=f"chiudi_{anno_archivio}_{nome_file}"):
                                st.session_state[f"anteprima_{anno_archivio}_{nome_file}"] = False
                                st.rerun()
                                
                            if ext_file in ["png", "jpg", "jpeg", "webp", "jfif"]:
                                st.image(file_completo, caption=nome_file, use_container_width=True)
                            elif ext_file == "pdf":
                                with open(file_completo, "rb") as f_pdf:
                                    base64_pdf = base64.b64encode(f_pdf.read()).decode('utf-8')
                                pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="600px" type="application/pdf"></iframe>'
                                st.markdown(pdf_display, unsafe_allow_html=True)
                            elif ext_file == "txt":
                                with open(file_completo, "r", encoding="utf-8", errors="ignore") as f_txt:
                                    st.text(f_txt.read())
            else:
                st.info(f"La cartella `{anno_archivio}` è attualmente vuota. Carica i tuoi primi file sopra!")
        else:
            st.warning("La cartella non esiste.")
            
    # --- VISTA DETTAGLIO ANNO SERVIZI ---
    elif st.session_state["vista_corrente"].startswith("ANNO_"):
        anno_selezionato = int(st.session_state["vista_corrente"].replace("ANNO_", ""))
        
        if st.button("⬅️ Torna alla Panoramica Generale", key="ritorna_panoramica_anno"):
            st.session_state["vista_corrente"] = "Panoramica Generale"
            st.rerun()
        st.markdown(f"## 📅 Dettaglio Servizi dell'Anno {anno_selezionato}")
        
        df_anno = df_servizi[df_servizi["Anno"] == anno_selezionato]
        s_anno = len(df_anno)
        o_anno_sec = df_anno["Secondi_Totali"].sum()
        o_anno_str = formato_hh_mm_ss(o_anno_sec)
        ma1, ma2 = st.columns(2)
        with ma1:
            st.metric(f"Servizi nel {anno_selezionato}", s_anno)
        with ma2:
            st.metric(f"Ore Totali nel {anno_selezionato}", o_anno_str)
        df_mostra_anno = (
            df_anno.drop(columns=["Anno", "Secondi_Totali"])
            .sort_values(by="Data", ascending=False)
            .copy()
        )
        df_mostra_anno["Data"] = df_mostra_anno["Data"].dt.strftime("%d/%m/%Y")
        st.dataframe(
            df_mostra_anno,
            use_container_width=True,
            hide_index=True,
        )
else:
    st.info("Nessun servizio presente o file non trovato. Carica un file dalla barra laterale.")
