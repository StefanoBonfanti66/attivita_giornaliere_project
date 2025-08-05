import streamlit as st
import pandas as pd
import plotly.express as px
import os
import datetime

# --- Configuration ---
st.set_page_config(
    page_title="Report Attività Giornaliere",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Dashboard Report Attività Giornaliere")

# --- File Uploader ---
st.sidebar.header("Carica il tuo file Excel")
uploaded_file = st.sidebar.file_uploader("Scegli un file Excel", type=["xlsx"])

# --- Load Data ---
@st.cache_data
def load_data(file):
    try:
        # Read all sheets from the Excel file
        xls = pd.ExcelFile(file)
        all_sheets_df = pd.DataFrame()
        
        for sheet_name in xls.sheet_names:
            # Skip the original sheet if it somehow persists or any other non-report sheets
            if "Foglio1" in sheet_name or "Sheet1" in sheet_name: # Adjust if original sheet name is different
                continue
            
            df_sheet = pd.read_excel(xls, sheet_name=sheet_name, skiprows=2) # Skip logo and title rows
            df_sheet['Report_Sheet'] = sheet_name # Add a column to identify the original sheet
            all_sheets_df = pd.concat([all_sheets_df, df_sheet], ignore_index=True)
            
        return all_sheets_df
    except Exception as e:
        st.error(f"Errore durante il caricamento del file Excel: {e}")
        return pd.DataFrame()

df = pd.DataFrame() # Initialize df as empty DataFrame
if uploaded_file is not None:
    df = load_data(uploaded_file)

if not df.empty:
    st.success(f"Dati caricati con successo dal file: {uploaded_file.name}")

    # --- Data Processing for Dashboard ---
    # Assuming 'Report_Sheet' contains 'Inseritore_Categoria' (e.g., 'Alessandra_Contatto Cliente')
    df['Inseritore'] = df['Report_Sheet'].apply(lambda x: x.split('_')[0] if '_' in x else 'Sconosciuto')
    df['Categoria'] = df['Report_Sheet'].apply(lambda x: x.split('_')[1] if '_' in x else 'Sconosciuto')

    # Convert 'dt.ins.' to datetime
    if 'dt.ins.' in df.columns:
        df['dt.ins.'] = pd.to_datetime(df['dt.ins.'], errors='coerce')

    # --- Filters ---
    st.sidebar.header("Filtri")

    # Inseritore filter
    all_inseritori = sorted(df['Inseritore'].unique())
    selected_inseritore = st.sidebar.multiselect(
        "Seleziona Inseritore:",
        options=all_inseritori,
        default=all_inseritori
    )

    # Date range filter
    st.sidebar.subheader("Filtro per Data")
    
    # Ensure 'dt.ins.' is datetime and handle NaT values
    if 'dt.ins.' in df.columns:
        df['dt.ins.'] = pd.to_datetime(df['dt.ins.'], errors='coerce')
        min_date = df['dt.ins.'].min()
        max_date = df['dt.ins.'].max()

        if pd.isna(min_date) or pd.isna(max_date):
            st.sidebar.warning("La colonna delle date ('dt.ins.') contiene valori non validi.")
            start_date = end_date = datetime.date.today()
        else:
            start_date = st.sidebar.date_input("Data di Inizio", value=min_date.date(), min_value=min_date.date(), max_value=max_date.date())
            end_date = st.sidebar.date_input("Data di Fine", value=max_date.date(), min_value=start_date, max_value=max_date.date())
    else:
        st.sidebar.warning("Colonna 'dt.ins.' non trovata. Impossibile filtrare per data.")
        start_date = end_date = datetime.date.today()

    # Apply filters
    df_filtered = df.copy() # Start with a copy of the original dataframe

    # Apply inseritore filter
    if selected_inseritore:
        df_filtered = df_filtered[df_filtered['Inseritore'].isin(selected_inseritore)]

    # Apply date filter
    if 'dt.ins.' in df.columns and not df_filtered.empty:
        # Convert start_date and end_date to datetime objects for comparison
        start_datetime = pd.to_datetime(start_date)
        end_datetime = pd.to_datetime(end_date)
        
        # Filter the dataframe
        df_filtered = df_filtered[
            (df_filtered['dt.ins.'].dt.normalize() >= start_datetime) & 
            (df_filtered['dt.ins.'].dt.normalize() <= end_datetime)
        ]

    # --- Key Metrics ---
    st.header("Riepilogo Generale")
    total_activities = len(df_filtered) # Use the filtered dataframe length
    st.metric("Numero Totale di Attività", total_activities)


    if df_filtered.empty:
        st.warning("Nessun dato disponibile per i filtri selezionati.")
    else:
        # --- Visualizations ---
        st.header("Visualizzazioni Dati")

        # 1. Attività per Inseritore
        st.subheader("Attività per Inseritore")
        activities_by_inseritore = df_filtered['Inseritore'].value_counts().reset_index()
        activities_by_inseritore.columns = ['Inseritore', 'Numero Attività']
        fig_inseritore = px.bar(
            activities_by_inseritore,
            x='Inseritore',
            y='Numero Attività',
            title='Numero di Attività per Inseritore',
            color='Inseritore'
        )
        st.plotly_chart(fig_inseritore, use_container_width=True)

        # 2. Distribuzione Contatto Cliente vs Azione Commerciale
        st.subheader("Distribuzione Attività per Categoria")
        activities_by_category = df_filtered['Categoria'].value_counts().reset_index()
        activities_by_category.columns = ['Categoria', 'Numero Attività']
        fig_category = px.pie(
            activities_by_category,
            names='Categoria',
            values='Numero Attività',
            title='Distribuzione Attività (Contatto Cliente vs Azione Commerciale)'
        )
        st.plotly_chart(fig_category, use_container_width=True)

        # 3. Attività per Inseritore e Categoria
        st.subheader("Attività per Inseritore e Categoria")
        activities_by_inseritore_category = df_filtered.groupby(['Inseritore', 'Categoria']).size().reset_index(name='Numero Attività')
        fig_stacked_bar = px.bar(
            activities_by_inseritore_category,
            x='Inseritore',
            y='Numero Attività',
            color='Categoria',
            title='Numero di Attività per Inseritore, suddivise per Categoria',
            barmode='group' # or 'stack'
        )
        st.plotly_chart(fig_stacked_bar, use_container_width=True)

        # --- Raw Data (Optional) ---
        st.header("Dati Dettagliati")
        columns_to_hide = ['dt.ins.', 'soggetto', 'contatto', 'report_sheet']
        df_display = df_filtered.drop(columns=[col for col in columns_to_hide if col in df_filtered.columns], errors='ignore')
        st.dataframe(df_display)

else:
    st.info("Carica un file Excel per visualizzare la dashboard.")