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

    # --- Key Metrics ---
    st.header("Riepilogo Generale")
    total_activities = len(df)
    st.metric("Numero Totale di Attività", total_activities)

    # --- Filters ---
    st.sidebar.header("Filtri")

    # Date Range Selector
    st.sidebar.subheader("Filtro per Data")
    min_date = df['dt.ins.'].min().date() if not df.empty and 'dt.ins.' in df.columns and not pd.isna(df['dt.ins.'].min()) else datetime.date.today()
    max_date = df['dt.ins.'].max().date() if not df.empty and 'dt.ins.' in df.columns and not pd.isna(df['dt.ins.'].max()) else datetime.date.today()

    start_date = st.sidebar.date_input("Data di Inizio", min_value=min_date, max_value=max_date, value=min_date)
    end_date = st.sidebar.date_input("Data di Fine", min_value=min_date, max_value=max_date, value=max_date)

    # Filter by date range
    df_filtered_by_date = df[(df['dt.ins.'].dt.date >= start_date) & (df['dt.ins.'].dt.date <= end_date)] if 'dt.ins.' in df.columns else df


    selected_inseritore = st.sidebar.multiselect(
        "Seleziona Inseritore:",
        options=df_filtered_by_date['Inseritore'].unique(),
        default=df_filtered_by_date['Inseritore'].unique()
    )

    df_filtered = df_filtered_by_date[df_filtered_by_date['Inseritore'].isin(selected_inseritore)]

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