import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from io import BytesIO

# Streamlit page configuration
st.set_page_config(page_title="Warranty Conversion Dashboard", layout="wide", initial_sidebar_state="expanded")

# --- Google Sheets Integration Configuration ---
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwx04lOl9zasbQUwCDeGVOR_kqzFJzfyvY1kS4oiYcjSHmHD1_b6Z3LWNSI9J1QRuws/exec"  # Replace with your Apps Script web app URL

# Configure requests with retry logic
session = requests.Session()
retries = Retry(total=3, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
session.mount('https://', HTTPAdapter(max_retries=retries))

# Available sheets
SHEETS = [
    "2024 December",
    "2025 JAN",
    "2025 FEB",
    "2025 MARCH",
    "2025 April",
    "2025 MAY",
    "2025 JUN",
    "2025 JULY",
    "2025 AUG",
    "2025 SEP"
]

# In sidebar
selected_sheet = st.selectbox("Select Month", SHEETS, index=0)

# Function to fetch data from Google Sheets for a specific sheet
@st.cache_data(ttl=60)  # Cache for 60 seconds
def fetch_data_from_sheets(sheet_name):
    try:
        response = session.get(APPS_SCRIPT_URL, params={"action": "read", "sheet": sheet_name}, timeout=30)
        response.raise_for_status()
        data = response.json()
        if data.get("status") == "success" and data.get("data"):
            df = pd.DataFrame(data["data"])
            return df
        else:
            st.error(f"Error fetching data from Google Sheets for {sheet_name}: {data.get('message', 'No data returned or invalid response')}")
            return None
    except requests.exceptions.RequestException as e:
        st.error(f"Failed to connect to Google Sheets for {sheet_name}: {str(e)}")
        return None

# Enhanced CSS for modern, attractive styling
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

        * {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        }
        
        .stApp {
            background: linear-gradient(135deg, #f0f4ff 0%, #e0e7ff 50%, #dbeafe 100%);
            background-attachment: fixed;
        }
        
        .main .block-container {
            padding: 2rem 3rem;
            max-width: 1400px;
        }
        
        /* Main Header */
        .main-header {
            background: linear-gradient(135deg, #1e40af 0%, #3b82f6 50%, #06b6d4 100%);
            color: white;
            font-size: 3em;
            font-weight: 900;
            text-align: center;
            padding: 40px;
            border-radius: 24px;
            margin-bottom: 40px;
            box-shadow: 0 25px 70px rgba(59, 130, 246, 0.5);
            letter-spacing: -1px;
            animation: slideDown 0.6s ease-out;
            position: relative;
            overflow: hidden;
        }
        
        .main-header::before {
            content: '';
            position: absolute;
            top: -50%;
            left: -50%;
            width: 200%;
            height: 200%;
            background: radial-gradient(circle, rgba(255,255,255,0.1) 0%, transparent 70%);
            animation: rotate 20s linear infinite;
        }
        
        @keyframes rotate {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        @keyframes slideDown {
            from {
                opacity: 0;
                transform: translateY(-50px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }
        
        /* Subheaders */
        .subheader {
            color: #1e293b;
            font-size: 1.75em;
            font-weight: 800;
            margin: 40px 0 25px 0;
            padding: 20px 25px;
            background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
            border-radius: 16px;
            border-left: 6px solid #3b82f6;
            box-shadow: 0 6px 20px rgba(59, 130, 246, 0.15);
            position: relative;
        }
        
        .subheader::after {
            content: '';
            position: absolute;
            bottom: 0;
            left: 25px;
            right: 25px;
            height: 3px;
            background: linear-gradient(90deg, #3b82f6 0%, #06b6d4 100%);
            border-radius: 2px;
        }
        
        /* Buttons */
        .stButton>button {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border-radius: 12px;
            padding: 14px 28px;
            font-weight: 600;
            font-size: 15px;
            border: none;
            box-shadow: 0 8px 20px rgba(102, 126, 234, 0.3);
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .stButton>button:hover {
            transform: translateY(-3px);
            box-shadow: 0 12px 28px rgba(102, 126, 234, 0.5);
        }
        
        .stButton>button:active {
            transform: translateY(-1px);
        }
        
        /* Download Button */
        .stDownloadButton>button {
            background: linear-gradient(135deg, #10b981 0%, #059669 100%);
            color: white;
            border-radius: 12px;
            padding: 12px 24px;
            font-weight: 600;
            border: none;
            box-shadow: 0 8px 20px rgba(16, 185, 129, 0.3);
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }
        
        .stDownloadButton>button:hover {
            transform: translateY(-3px);
            box-shadow: 0 12px 28px rgba(16, 185, 129, 0.5);
        }
        
        /* Input Fields */
        .stTextInput>div>div>input,
        .stSelectbox>div>div>select {
            border-radius: 10px;
            border: 2px solid #e2e8f0;
            padding: 12px 16px;
            background-color: white;
            transition: all 0.3s ease;
            font-size: 14px;
        }
        
        .stTextInput>div>div>input:focus,
        .stSelectbox>div>div>select:focus {
            border-color: #667eea;
            box-shadow: 0 0 0 4px rgba(102, 126, 234, 0.1);
            outline: none;
        }
        
        /* Sidebar */
        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
            border-right: none;
            box-shadow: 4px 0 20px rgba(0, 0, 0, 0.08);
        }
        
        [data-testid="stSidebar"] .stForm {
            background: linear-gradient(135deg, #f0f4ff 0%, #e0e7ff 100%);
            padding: 24px;
            border-radius: 16px;
            margin-bottom: 24px;
            border: 2px solid #e0e7ff;
            box-shadow: 0 4px 12px rgba(102, 126, 234, 0.1);
        }
        
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3 {
            color: #1e293b;
            font-weight: 700;
            margin-bottom: 16px;
        }
        
        /* Metrics */
        [data-testid="stMetricValue"] {
            font-size: 2.2em;
            font-weight: 800;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }
        
        [data-testid="stMetricLabel"] {
            font-size: 0.95em;
            font-weight: 600;
            color: #475569;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        
        div[data-testid="metric-container"] {
            background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
            padding: 28px;
            border-radius: 20px;
            box-shadow: 0 10px 30px rgba(102, 126, 234, 0.15);
            border: 2px solid #e0e7ff;
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
            position: relative;
            overflow: hidden;
        }
        
        div[data-testid="metric-container"]::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 4px;
            background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        }
        
        div[data-testid="metric-container"]:hover {
            transform: translateY(-8px);
            box-shadow: 0 20px 40px rgba(102, 126, 234, 0.25);
            border-color: #c7d2fe;
        }
        
        /* DataFrames */
        .stDataFrame {
            border-radius: 16px;
            overflow: hidden;
            background: white;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.08);
            border: 1px solid #e2e8f0;
        }
        
        .stDataFrame thead tr th {
            background: linear-gradient(135deg, #1e40af 0%, #3b82f6 100%) !important;
            color: white !important;
            font-weight: 700 !important;
            text-align: center !important;
            padding: 16px !important;
            font-size: 14px !important;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .stDataFrame tbody tr td {
            text-align: center !important;
            padding: 14px !important;
            font-size: 14px !important;
            border-bottom: 1px solid #f1f5f9 !important;
            color: #1e293b !important;
            font-weight: 500 !important;
        }
        
        .stDataFrame tbody tr:hover {
            background-color: #f1f5f9 !important;
        }
        
        .stDataFrame tbody tr:last-child {
            background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%) !important;
            font-weight: 700 !important;
            border-top: 3px solid #f59e0b !important;
        }
        
        .stDataFrame tbody tr:last-child td {
            color: #92400e !important;
            font-weight: 700 !important;
            font-size: 15px !important;
        }
        
        /* Charts */
        .stPlotlyChart {
            border-radius: 16px;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.08);
            background: white;
            padding: 20px;
            border: 1px solid #e2e8f0;
        }
        
        /* Alert Boxes */
        .stSuccess, .stWarning, .stInfo, .stError {
            border-radius: 12px;
            padding: 20px;
            margin: 16px 0;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
            font-weight: 500;
            border: none;
        }
        
        .stSuccess {
            background: linear-gradient(135deg, #d1fae5 0%, #a7f3d0 100%);
            color: #065f46;
        }
        
        .stWarning {
            background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%);
            color: #92400e;
        }
        
        .stInfo {
            background: linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%);
            color: #1e40af;
        }
        
        .stError {
            background: linear-gradient(135deg, #fee2e2 0%, #fecaca 100%);
            color: #991b1b;
        }
        
        /* Expander */
        .streamlit-expanderHeader {
            background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
            border-radius: 12px;
            padding: 16px;
            font-weight: 600;
            color: #1e293b;
            border: 1px solid #e2e8f0;
        }
        
        .streamlit-expanderHeader:hover {
            background: linear-gradient(135deg, #f1f5f9 0%, #e2e8f0 100%);
        }
        
        /* Checkbox */
        .stCheckbox label {
            color: #1e293b;
            font-weight: 500;
            font-size: 14px;
        }
        
        /* Slider */
        .stSlider [data-testid="stTickBar"] {
            background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        }
        
        /* Hide Streamlit branding */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        
        /* Scrollbar */
        ::-webkit-scrollbar {
            width: 10px;
            height: 10px;
        }
        
        ::-webkit-scrollbar-track {
            background: #f1f5f9;
            border-radius: 10px;
        }
        
        ::-webkit-scrollbar-thumb {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-radius: 10px;
        }
        
        ::-webkit-scrollbar-thumb:hover {
            background: linear-gradient(135deg, #764ba2 0%, #667eea 100%);
        }
        
        /* Separator */
        hr {
            border: none;
            height: 2px;
            background: linear-gradient(90deg, transparent, #667eea, transparent);
            margin: 30px 0;
        }
        
        /* Low conversion highlight */
        .low-value-conversion {
            background-color: #fee2e2 !important;
        }
    </style>
""", unsafe_allow_html=True)

# Function to convert DataFrame to Excel for download
def to_excel(df, sheet_name='Data'):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export = df.copy()
        
        percentage_columns = [col for col in df_export.columns if 'Conv (%)' in col]
        for col in percentage_columns:
            if col in df_export.columns:
                df_export[col] = df_export[col] / 100.0
        
        df_export.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#667eea',
            'font_color': 'white',
            'border': 1,
            'align': 'center'
        })
        
        for col_num, value in enumerate(df_export.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        low_conversion_format = workbook.add_format({'bg_color': '#fee2e2'})
        value_conv_col = df_export.columns.get_loc('Value Conv (%)') if 'Value Conv (%)' in df_export.columns else None
        
        if value_conv_col is not None:
            worksheet.conditional_format(1, value_conv_col, len(df_export), value_conv_col, {
                'type': 'cell',
                'criteria': '<',
                'value': 0.02,
                'format': low_conversion_format
            })
        
        num_format = workbook.add_format({'num_format': '#,##0.00', 'align': 'center'})
        percent_format = workbook.add_format({'num_format': '0.00%', 'align': 'center'})
        currency_format = workbook.add_format({'num_format': '₹#,##0.00', 'align': 'center'})
        integer_format = workbook.add_format({'num_format': '#,##0', 'align': 'center'})
        text_format = workbook.add_format({'align': 'center'})
        
        for col_num, col_name in enumerate(df_export.columns):
            if 'Conv (%)' in col_name:
                worksheet.set_column(col_num, col_num, 15, percent_format)
            elif 'AHSP' in col_name:
                worksheet.set_column(col_num, col_num, 15, currency_format)
            elif 'Sales' in col_name:
                worksheet.set_column(col_num, col_num, 18, currency_format)
            elif 'Units' in col_name:
                worksheet.set_column(col_num, col_num, 15, integer_format)
            elif df_export[col_name].dtype in ['float64']:
                worksheet.set_column(col_num, col_num, 15, num_format)
            elif df_export[col_name].dtype in ['int64']:
                worksheet.set_column(col_num, col_num, 15, integer_format)
            else:
                worksheet.set_column(col_num, col_num, 20, text_format)
        
        if 'Total' in df_export['Store'].values if 'Store' in df_export.columns else False:
            total_row_format = workbook.add_format({'bold': True, 'align': 'center'})
            total_row_index = df_export.index[df_export['Store'] == 'Total'][0] + 1
            worksheet.set_row(total_row_index, None, total_row_format)
    
    processed_data = output.getvalue()
    return processed_data

# Function to map item categories to replacement warranty categories
def map_to_replacement_category(item_category):
    fan_categories = ['CEILING FAN', 'PEDESTAL FAN', 'RECHARGABLE FAN', 'TABLE FAN', 'TOWER FAN', 'WALL FAN']
    steamer_categories = ['GARMENTS STEAMER', 'STEAMER']
    
    if any(fan in str(item_category).upper() for fan in fan_categories):
        return 'FAN'
    elif 'MIXER GRINDER' in str(item_category).upper():
        return 'MIXER GRINDER'
    elif 'IRON BOX' in str(item_category).upper():
        return 'IRON BOX'
    elif 'ELECTRIC KETTLE' in str(item_category).upper():
        return 'ELECTRIC KETTLE'
    elif 'OTG' in str(item_category).upper():
        return 'OTG'
    elif any(steamer in str(item_category).upper() for steamer in steamer_categories):
        return 'STEAMER'
    elif 'INDUCTION' in str(item_category).upper():
        return 'INDUCTION COOKER'
    else:
        return item_category

# Function to check if category is small appliance
def is_small_appliance(item_category):
    major_appliances = ['AC', 'TV', 'WASHING MACHINE', 'REFRIGERATOR', 'MICROWAVE OVEN', 'DISH WASHER', 'DRYER']
    return not any(appliance in str(item_category).upper() for appliance in major_appliances)

# Function to get appliance type
def get_appliance_type(item_category):
    major_appliances = ['AC', 'TV', 'WASHING MACHINE', 'REFRIGERATOR', 'MICROWAVE OVEN', 'DISH WASHER', 'DRYER']
    if any(appliance in str(item_category).upper() for appliance in major_appliances):
        return 'Large Appliance'
    else:
        return 'Small Appliance'

# Session state initialization
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'current_df' not in st.session_state:
    st.session_state.current_df = None
if 'selected_sheet' not in st.session_state:
    st.session_state.selected_sheet = SHEETS[0]  # Default to first sheet

# Sidebar content
with st.sidebar:
    st.markdown('<h2 style="color: #1e293b; font-weight: 700; margin-bottom: 20px;">🔍 Dashboard Controls</h2>', unsafe_allow_html=True)
    st.markdown('<hr>', unsafe_allow_html=True)
    
    # Month selection dropdown
    selected_sheet = st.selectbox("📅 Select Month", SHEETS, index=SHEETS.index(st.session_state.selected_sheet))
    
    # Update session state if month changes
    if selected_sheet != st.session_state.selected_sheet:
        st.session_state.selected_sheet = selected_sheet
        st.session_state.current_df = None
        st.session_state.data_loaded = False
        st.cache_data.clear()
    
    # Refresh Button
    if st.button("🔄 Refresh Data"):
        st.cache_data.clear()
        st.session_state.current_df = None
        st.session_state.data_loaded = False
        st.rerun()

# Main dashboard header
st.markdown(f'<div class="main-header">📊 Warranty Conversion Analysis Dashboard - {st.session_state.selected_sheet}</div>', unsafe_allow_html=True)

# Load data function
required_columns = ['Item Category', 'BDM', 'RBM', 'Store', 'Staff Name', 'TotalSoldPrice', 'WarrantyPrice', 'TotalCount', 'WarrantyCount']

@st.cache_data
def load_data(sheet_name):
    try:
        df = fetch_data_from_sheets(sheet_name)
        if df is None:
            return None
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"Missing columns in Google Sheets data for {sheet_name}: {', '.join(missing_columns)}")
            return None
        
        numeric_cols = ['TotalSoldPrice', 'WarrantyPrice', 'TotalCount', 'WarrantyCount']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        
        if df[numeric_cols].isna().any().any():
            st.warning(f"Missing or invalid values in numeric columns for {sheet_name}. Filling with 0.")
            df[numeric_cols] = df[numeric_cols].fillna(0)
        
        df['Replacement Category'] = df['Item Category'].apply(map_to_replacement_category)
        df['Appliance Type'] = df['Item Category'].apply(get_appliance_type)
        
        df['Conversion% (Count)'] = (df['WarrantyCount'] / df['TotalCount'] * 100).round(2)
        df['Conversion% (Price)'] = (df['WarrantyPrice'] / df['TotalSoldPrice'] * 100).where(df['TotalSoldPrice'] > 0, 0).round(2)
        df['AHSP'] = (df['WarrantyPrice'] / df['WarrantyCount']).where(df['WarrantyCount'] > 0, 0).round(2)
        df['Month'] = sheet_name
        
        return df
    except Exception as e:
        st.error(f"❌ Error loading data from Google Sheets for {sheet_name}: {str(e)}")
        return None

# Load data from Google Sheets
if st.session_state.current_df is None:
    df = load_data(st.session_state.selected_sheet)
    if df is not None:
        st.session_state.current_df = df
        st.session_state.data_loaded = True

# Now that we have data, set up the filters in sidebar
if st.session_state.data_loaded and st.session_state.current_df is not None:
    df = st.session_state.current_df
    
    with st.sidebar:
        # Sidebar filters
        st.markdown('<hr>', unsafe_allow_html=True)
        st.markdown('<h3 style="color: #1e293b; font-weight: 600; margin-top: 20px;">⚙️ Filters</h3>', unsafe_allow_html=True)
        
        replacement_filter = st.checkbox("🔄 Show Replacement Warranty Categories Only")
        speaker_filter = st.checkbox("🔊 Show Speaker Categories Only")
        future_filter = st.checkbox("🏢 Show only FUTURE stores")
        
        # Value Conversion range filter
        st.markdown('<hr>', unsafe_allow_html=True)
        st.markdown('<h4 style="color: #1e293b; font-weight: 600;">📊 Value Conversion Filter</h4>', unsafe_allow_html=True)
        
        store_summary_temp = df.groupby('Store').agg({
            'TotalSoldPrice': 'sum',
            'WarrantyPrice': 'sum'
        }).reset_index()
        store_summary_temp['Value Conv (%)'] = (store_summary_temp['WarrantyPrice'] / store_summary_temp['TotalSoldPrice'] * 100).where(store_summary_temp['TotalSoldPrice'] > 0, 0).round(2)
        
        min_conv = float(store_summary_temp['Value Conv (%)'].min())
        max_conv = float(store_summary_temp['Value Conv (%)'].max())
        
        if min_conv == max_conv:
            min_conv = max(0, min_conv - 0.1)
            max_conv = max_conv + 0.1
        
        value_conv_range = st.slider(
            "Filter by Value Conversion (%)",
            min_value=min_conv,
            max_value=max_conv,
            value=(min_conv, max_conv),
            step=0.1,
            help="Filter stores based on Value Conversion percentage range"
        )
        
        # Sorting option for all tables
        st.markdown('<hr>', unsafe_allow_html=True)
        sort_by = st.selectbox("📊 Sort All Tables By", ["Count Conv (%)", "Value Conv (%)", "AHSP (₹)", "Warranty Sales (₹)", "Warranty Units"], index=0)
        sort_order = st.selectbox("↕️ Sort Order", ["Descending", "Ascending"], index=0)

    # Apply replacement or speaker filter
    if replacement_filter:
        replacement_categories = ['FAN', 'MIXER GRINDER', 'IRON BOX', 'ELECTRIC KETTLE', 'OTG', 'STEAMER', 'INDUCTION COOKER']
        df = df[df['Replacement Category'].isin(replacement_categories)]
        category_column = 'Replacement Category'
    elif speaker_filter:
        speaker_categories = ['SOUND BAR', 'PARTY SPEAKER', 'BLUETOOTH SPEAKER', 'HOME THEATRE']
        df = df[df['Item Category'].isin(speaker_categories)]
        category_column = 'Item Category'
    else:
        category_column = 'Item Category'

    # Define filters after df is loaded
    with st.sidebar:
        bdm_options = ['All'] + sorted(df['BDM'].unique().tolist())
        rbm_options = ['All'] + sorted(df['RBM'].unique().tolist())
        store_options = ['All'] + sorted(df['Store'].unique().tolist())
        category_options = ['All'] + sorted(df[category_column].unique().tolist())

        selected_bdm = st.selectbox("👔 BDM", bdm_options, index=0)
        selected_rbm = st.selectbox("👤 RBM", rbm_options, index=0)
        selected_store = st.selectbox("🏪 Store", store_options, index=0)
        selected_category = st.selectbox(f"📦 {category_column}", category_options, index=0)

        if selected_rbm != 'All':
            staff_options = ['All'] + sorted(df[df['RBM'] == selected_rbm]['Staff Name'].unique().tolist())
        else:
            staff_options = ['All'] + sorted(df['Staff Name'].unique().tolist())
        selected_staff = st.selectbox("👨‍💼 Staff", staff_options, index=0)

    # Apply filters
    filtered_df = df.copy()
    if selected_bdm != 'All':
        filtered_df = filtered_df[filtered_df['BDM'] == selected_bdm]
    if selected_rbm != 'All':
        filtered_df = filtered_df[filtered_df['RBM'] == selected_rbm]
    if selected_store != 'All':
        filtered_df = filtered_df[filtered_df['Store'] == selected_store]
    if selected_category != 'All':
        filtered_df = filtered_df[filtered_df[category_column] == selected_category]
    if selected_staff != 'All':
        filtered_df = filtered_df[filtered_df['Staff Name'] == selected_staff]
    if future_filter:
        filtered_df = filtered_df[filtered_df['Store'].str.contains('FUTURE', case=True)]

    if filtered_df.empty:
        st.warning("⚠️ No data matches your filters.")
        st.stop()

    # Main dashboard
    st.markdown(f'<h2 class="subheader">📈 {st.session_state.selected_sheet} Performance Overview</h2>', unsafe_allow_html=True)

    # KPI metrics
    st.markdown('<h3 style="color: #1e293b; font-weight: 600; margin: 25px 0 15px 0;">🎯 Key Performance Indicators</h3>', unsafe_allow_html=True)
    col1, col2, col3, col4 = st.columns(4)

    total_warranty = filtered_df['WarrantyPrice'].sum()
    total_units = filtered_df['TotalCount'].sum()
    total_warranty_units = filtered_df['WarrantyCount'].sum()
    total_sales = filtered_df['TotalSoldPrice'].sum()
    count_conversion = (total_warranty_units / total_units * 100) if total_units > 0 else 0
    value_conversion = (total_warranty / total_sales * 100) if total_sales > 0 else 0
    ahsp = (total_warranty / total_warranty_units) if total_warranty_units > 0 else 0

    with col1:
        st.metric("💰 Warranty Sales", f"₹{total_warranty:,.0f}")
    with col2:
        st.metric("📊 Count Conversion", f"{count_conversion:.2f}%")
    with col3:
        st.metric("📈 Value Conversion", f"{value_conversion:.2f}%")
    with col4:
        st.metric("💵 AHSP", f"₹{ahsp:,.2f}")

    # Store Performance
    st.markdown(f'<h3 class="subheader">🏬 Store Performance Analysis - {st.session_state.selected_sheet}</h3>', unsafe_allow_html=True)

    store_summary = filtered_df.groupby('Store').agg({
        'TotalSoldPrice': 'sum',
        'WarrantyPrice': 'sum',
        'TotalCount': 'sum',
        'WarrantyCount': 'sum'
    }).reset_index()

    store_summary['Count Conv (%)'] = (store_summary['WarrantyCount'] / store_summary['TotalCount'] * 100).round(2)
    store_summary['Value Conv (%)'] = (store_summary['WarrantyPrice'] / store_summary['TotalSoldPrice'] * 100).round(2)
    store_summary['AHSP'] = (store_summary['WarrantyPrice'] / store_summary['WarrantyCount']).where(store_summary['WarrantyCount'] > 0, 0).round(2)

    store_summary['Count Conv (%)'] = store_summary['Count Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)
    store_summary['Value Conv (%)'] = store_summary['Value Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)

    store_display = store_summary[['Store', 'WarrantyPrice', 'WarrantyCount', 'Count Conv (%)', 'Value Conv (%)', 'AHSP']].copy()
    store_display.columns = ['Store', 'Warranty Sales (₹)', 'Warranty Units', 'Count Conv (%)', 'Value Conv (%)', 'AHSP (₹)']

    total_sales = store_summary['TotalSoldPrice'].sum()
    total_warranty_sales = store_summary['WarrantyPrice'].sum()
    total_units = store_summary['TotalCount'].sum()
    total_warranty_units = store_summary['TotalCount'].sum()
    
    total_count_conv = (total_warranty_units / total_units * 100) if total_units > 0 else 0
    total_value_conv = (total_warranty_sales / total_sales * 100) if total_sales > 0 else 0
    total_ahsp = (total_warranty_sales / total_warranty_units) if total_warranty_units > 0 else 0

    total_row = pd.DataFrame({
        'Store': ['Total'],
        'Warranty Sales (₹)': [total_warranty_sales],
        'Warranty Units': [total_warranty_units],
        'Count Conv (%)': [round(total_count_conv, 2)],
        'Value Conv (%)': [round(total_value_conv, 2)],
        'AHSP (₹)': [round(total_ahsp, 2)]
    })

    non_total_stores = store_display[store_display['Store'] != 'Total']

    if value_conv_range != (min_conv, max_conv):
        non_total_stores = non_total_stores[
            (non_total_stores['Value Conv (%)'] >= value_conv_range[0]) & 
            (non_total_stores['Value Conv (%)'] <= value_conv_range[1])
        ]

    sort_ascending = True if sort_order == "Ascending" else False

    sort_column_mapping = {
        "Count Conv (%)": "Count Conv (%)",
        "Value Conv (%)": "Value Conv (%)", 
        "AHSP (₹)": "AHSP (₹)",
        "Warranty Sales (₹)": "Warranty Sales (₹)",
        "Warranty Units": "Warranty Units"
    }

    sort_column = sort_column_mapping.get(sort_by, "Count Conv (%)")
    non_total_stores = non_total_stores.sort_values(sort_column, ascending=sort_ascending)

    final_store_display = pd.concat([non_total_stores, total_row], ignore_index=True)

    def highlight_low_value_conversion(row):
        if row['Value Conv (%)'] < 2.0 and row['Store'] != 'Total':
            return ['background-color: #fee2e2'] * len(row)
        return [''] * len(row)

    styled_final_display = final_store_display.style.format({
        'Warranty Sales (₹)': '₹{:,.0f}',
        'Warranty Units': '{:,.0f}',
        'Count Conv (%)': '{:.2f}%',
        'Value Conv (%)': '{:.2f}%',
        'AHSP (₹)': '₹{:.2f}'
    }).apply(highlight_low_value_conversion, axis=1)

    st.dataframe(styled_final_display, use_container_width=True)

    st.download_button(
        label="📥 Download Store Performance as Excel",
        data=to_excel(final_store_display, 'Store Performance'),
        file_name=f"store_performance_{st.session_state.selected_sheet.lower().replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    store_summary_for_chart = store_summary[store_summary['Store'].isin(non_total_stores['Store'])]

    if not store_summary_for_chart.empty:
        chart_sort_column_mapping = {
            "Count Conv (%)": "Count Conv (%)",
            "Value Conv (%)": "Value Conv (%)",
            "AHSP (₹)": "AHSP",
            "Warranty Sales (₹)": "WarrantyPrice", 
            "Warranty Units": "WarrantyCount"
        }
        
        chart_sort_column = chart_sort_column_mapping.get(sort_by, "Count Conv (%)")
        store_summary_for_chart = store_summary_for_chart.sort_values(chart_sort_column, ascending=sort_ascending)
        
        if sort_by == "Value Conv (%)":
            chart_y_column = 'Value Conv (%)'
            chart_title = f'📊 Store Value Conversion - {st.session_state.selected_sheet}'
            chart_text = 'Value Conv (%)'
        elif sort_by == "AHSP (₹)":
            chart_y_column = 'AHSP'
            chart_title = f'💵 Store AHSP - {st.session_state.selected_sheet}'
            chart_text = 'AHSP'
        elif sort_by == "Warranty Sales (₹)":
            chart_y_column = 'WarrantyPrice'
            chart_title = f'💰 Store Warranty Sales - {st.session_state.selected_sheet}'
            chart_text = 'WarrantyPrice'
        elif sort_by == "Warranty Units":
            chart_y_column = 'WarrantyCount'
            chart_title = f'📦 Store Warranty Units - {st.session_state.selected_sheet}'
            chart_text = 'WarrantyCount'
        else:
            chart_y_column = 'Count Conv (%)'
            chart_title = f'📊 Store Count Conversion - {st.session_state.selected_sheet}'
            chart_text = 'Count Conv (%)'

        fig_store = px.bar(store_summary_for_chart, 
                           x='Store', 
                           y=chart_y_column, 
                           title=chart_title,
                           template='plotly_white',
                           color=chart_y_column,
                           color_continuous_scale='Viridis',
                           text=chart_text)
        
        if chart_y_column in ['Count Conv (%)', 'Value Conv (%)']:
            fig_store.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
        elif chart_y_column == 'AHSP':
            fig_store.update_traces(texttemplate='₹%{text:.2f}', textposition='outside')
        elif chart_y_column == 'WarrantyPrice':
            fig_store.update_traces(texttemplate='₹%{text:,.0f}', textposition='outside')
        else:
            fig_store.update_traces(texttemplate='%{text:.0f}', textposition='outside')
        
        fig_store.update_layout(
            plot_bgcolor='rgba(255,255,255,1)',
            paper_bgcolor='rgba(255,255,255,1)',
            font=dict(family="Inter, sans-serif", size=12, color="#1e293b"),
            xaxis=dict(showgrid=False, tickangle=45, title_font=dict(size=14, color="#1e293b")),
            yaxis=dict(showgrid=True, gridcolor='#e2e8f0', title_font=dict(size=14, color="#1e293b")),
            title_font=dict(size=18, color="#1e293b", family="Inter"),
            height=550,
            margin=dict(t=80, b=80, l=60, r=60)
        )
        st.plotly_chart(fig_store, use_container_width=True)

    # Staff Performance Table
    st.markdown(f'<h3 class="subheader">👨‍💼 Staff Performance Analysis - {st.session_state.selected_sheet}</h3>', unsafe_allow_html=True)

    staff_summary = filtered_df.groupby(['Staff Name', 'Store']).agg({
        'TotalSoldPrice': 'sum',
        'WarrantyPrice': 'sum',
        'TotalCount': 'sum',
        'WarrantyCount': 'sum'
    }).reset_index()

    staff_summary['Count Conv (%)'] = (staff_summary['WarrantyCount'] / staff_summary['TotalCount'] * 100).round(2)
    staff_summary['Value Conv (%)'] = (staff_summary['WarrantyPrice'] / staff_summary['TotalSoldPrice'] * 100).round(2)
    staff_summary['AHSP'] = (staff_summary['WarrantyPrice'] / staff_summary['WarrantyCount']).where(staff_summary['WarrantyCount'] > 0, 0).round(2)

    staff_summary['Count Conv (%)'] = staff_summary['Count Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)
    staff_summary['Value Conv (%)'] = staff_summary['Value Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)

    staff_display = staff_summary[['Staff Name', 'Store', 'Value Conv (%)', 'Count Conv (%)', 'WarrantyPrice', 'WarrantyCount', 'AHSP']].copy()
    staff_display.columns = ['Staff Name', 'Store', 'Value Conv (%)', 'Count Conv (%)', 'Warranty Sales (₹)', 'Warranty Units', 'AHSP (₹)']
    
    # Apply sorting to staff performance table
    staff_sort_column_mapping = {
        "Count Conv (%)": "Count Conv (%)",
        "Value Conv (%)": "Value Conv (%)", 
        "AHSP (₹)": "AHSP (₹)",
        "Warranty Sales (₹)": "Warranty Sales (₹)",
        "Warranty Units": "Warranty Units"
    }
    staff_sort_column = staff_sort_column_mapping.get(sort_by, "Count Conv (%)")
    staff_display = staff_display.sort_values(staff_sort_column, ascending=sort_ascending)

    st.dataframe(staff_display.style.format({
        'Value Conv (%)': '{:.2f}%',
        'Count Conv (%)': '{:.2f}%',
        'Warranty Sales (₹)': '₹{:.0f}',
        'Warranty Units': '{:.0f}',
        'AHSP (₹)': '₹{:.2f}'
    }), use_container_width=True)

    st.download_button(
        label="📥 Download Staff Performance as Excel",
        data=to_excel(staff_display, 'Staff Performance'),
        file_name=f"staff_performance_{st.session_state.selected_sheet.lower().replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # RBM Performance
    st.markdown(f'<h3 class="subheader">👥 RBM Performance Analysis - {st.session_state.selected_sheet}</h3>', unsafe_allow_html=True)

    rbm_summary = filtered_df.groupby('RBM').agg({
        'TotalSoldPrice': 'sum',
        'WarrantyPrice': 'sum',
        'TotalCount': 'sum',
        'WarrantyCount': 'sum'
    }).reset_index()

    rbm_summary['Count Conv (%)'] = (rbm_summary['WarrantyCount'] / rbm_summary['TotalCount'] * 100).round(2)
    rbm_summary['Value Conv (%)'] = (rbm_summary['WarrantyPrice'] / rbm_summary['TotalSoldPrice'] * 100).round(2)
    rbm_summary['AHSP'] = (rbm_summary['WarrantyPrice'] / rbm_summary['WarrantyCount']).where(rbm_summary['WarrantyCount'] > 0, 0).round(2)

    rbm_summary['Count Conv (%)'] = rbm_summary['Count Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)
    rbm_summary['Value Conv (%)'] = rbm_summary['Value Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)

    rbm_display = rbm_summary[['RBM', 'Count Conv (%)', 'Value Conv (%)', 'AHSP', 'WarrantyPrice', 'WarrantyCount']]
    rbm_display.columns = ['RBM', 'Count Conv (%)', 'Value Conv (%)', 'AHSP (₹)', 'Warranty Sales (₹)', 'Warranty Units']
    
    # Apply sorting to RBM performance table
    rbm_sort_column_mapping = {
        "Count Conv (%)": "Count Conv (%)",
        "Value Conv (%)": "Value Conv (%)", 
        "AHSP (₹)": "AHSP (₹)",
        "Warranty Sales (₹)": "Warranty Sales (₹)",
        "Warranty Units": "Warranty Units"
    }
    rbm_sort_column = rbm_sort_column_mapping.get(sort_by, "Count Conv (%)")
    rbm_display = rbm_display.sort_values(rbm_sort_column, ascending=sort_ascending)

    st.dataframe(rbm_display.style.format({
        'Count Conv (%)': '{:.2f}%',
        'Value Conv (%)': '{:.2f}%',
        'AHSP (₹)': '₹{:.2f}',
        'Warranty Sales (₹)': '₹{:.0f}',
        'Warranty Units': '{:.0f}'
    }), use_container_width=True)

    st.download_button(
        label="📥 Download RBM Performance as Excel",
        data=to_excel(rbm_display, 'RBM Performance'),
        file_name=f"rbm_performance_{st.session_state.selected_sheet.lower().replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Product Category Performance with Small Appliance Grouping
    st.markdown(f'<h3 class="subheader">📦 Product Category Performance - {st.session_state.selected_sheet}</h3>', unsafe_allow_html=True)

    # Define major appliances that should stay as separate rows
    major_appliances = ['AC', 'TV', 'WASHING MACHINE', 'REFRIGERATOR', 'MICROWAVE OVEN', 'DISH WASHER', 'DRYER']

    # Create a new column for the grouped category
    filtered_df['Grouped Category'] = filtered_df['Item Category'].apply(
        lambda x: 'SMALL APPLIANCE' if x not in major_appliances else x
    )

    category_summary = filtered_df.groupby('Grouped Category').agg({
        'TotalSoldPrice': 'sum',
        'WarrantyPrice': 'sum',
        'TotalCount': 'sum',
        'WarrantyCount': 'sum'
    }).reset_index()

    category_summary['Count Conv (%)'] = (category_summary['WarrantyCount'] / category_summary['TotalCount'] * 100).round(2)
    category_summary['Value Conv (%)'] = (category_summary['WarrantyPrice'] / category_summary['TotalSoldPrice'] * 100).round(2)
    category_summary['AHSP'] = (category_summary['WarrantyPrice'] / category_summary['WarrantyCount']).where(category_summary['WarrantyCount'] > 0, 0).round(2)

    category_summary['Count Conv (%)'] = category_summary['Count Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)
    category_summary['Value Conv (%)'] = category_summary['Value Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)

    if not category_summary.empty:
        category_display = category_summary[['Grouped Category', 'Count Conv (%)', 'Value Conv (%)', 'AHSP', 'WarrantyPrice', 'WarrantyCount']]
        category_display.columns = ['Product Category', 'Count Conv (%)', 'Value Conv (%)', 'AHSP (₹)', 'Warranty Sales (₹)', 'Warranty Units']
        
        # Apply sorting to category performance table
        category_sort_column_mapping = {
            "Count Conv (%)": "Count Conv (%)",
            "Value Conv (%)": "Value Conv (%)", 
            "AHSP (₹)": "AHSP (₹)",
            "Warranty Sales (₹)": "Warranty Sales (₹)",
            "Warranty Units": "Warranty Units"
        }
        category_sort_column = category_sort_column_mapping.get(sort_by, "Count Conv (%)")
        category_display = category_display.sort_values(category_sort_column, ascending=sort_ascending)

        st.dataframe(category_display.style.format({
            'Count Conv (%)': '{:.2f}%',
            'Value Conv (%)': '{:.2f}%',
            'AHSP (₹)': '₹{:.2f}',
            'Warranty Sales (₹)': '₹{:.0f}',
            'Warranty Units': '{:.0f}'
        }), use_container_width=True)

        st.download_button(
            label="📥 Download Product Category Performance as Excel",
            data=to_excel(category_display, 'Product Category Performance'),
            file_name=f"product_category_performance_{st.session_state.selected_sheet.lower().replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Chart for the grouped categories
        category_summary_for_chart = category_summary.copy()
        category_summary_for_chart = category_summary_for_chart.sort_values(category_sort_column, ascending=sort_ascending)
        
        fig_category = px.bar(category_summary_for_chart, 
                              x='Grouped Category', 
                              y='Count Conv (%)', 
                              title=f'📊 Product Category Count Conversion - {st.session_state.selected_sheet}',
                              template='plotly_white',
                              color='Count Conv (%)',
                              color_continuous_scale='Plasma',
                              text='Count Conv (%)')
        
        fig_category.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
        fig_category.update_layout(
            plot_bgcolor='rgba(255,255,255,1)',
            paper_bgcolor='rgba(255,255,255,1)',
            font=dict(family="Inter, sans-serif", size=12, color="#1e293b"),
            xaxis=dict(showgrid=False, tickangle=45, title_font=dict(size=14, color="#1e293b")),
            yaxis=dict(showgrid=True, gridcolor='#e2e8f0', title_font=dict(size=14, color="#1e293b")),
            title_font=dict(size=18, color="#1e293b", family="Inter"),
            height=550,
            margin=dict(t=80, b=80, l=60, r=60)
        )
        st.plotly_chart(fig_category, use_container_width=True)
    else:
        st.warning("⚠️ No category data available with current filters.")

    # Item Category Performance (Full Product Breakdown)
    st.markdown(f'<h3 class="subheader">📋 Item Category Performance - Full Product Breakdown - {st.session_state.selected_sheet}</h3>', unsafe_allow_html=True)

    # Full item category performance without grouping
    item_category_summary = filtered_df.groupby('Item Category').agg({
        'TotalSoldPrice': 'sum',
        'WarrantyPrice': 'sum',
        'TotalCount': 'sum',
        'WarrantyCount': 'sum'
    }).reset_index()

    item_category_summary['Count Conv (%)'] = (item_category_summary['WarrantyCount'] / item_category_summary['TotalCount'] * 100).round(2)
    item_category_summary['Value Conv (%)'] = (item_category_summary['WarrantyPrice'] / item_category_summary['TotalSoldPrice'] * 100).round(2)
    item_category_summary['AHSP'] = (item_category_summary['WarrantyPrice'] / item_category_summary['WarrantyCount']).where(item_category_summary['WarrantyCount'] > 0, 0).round(2)

    item_category_summary['Count Conv (%)'] = item_category_summary['Count Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)
    item_category_summary['Value Conv (%)'] = item_category_summary['Value Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)

    if not item_category_summary.empty:
        item_category_display = item_category_summary[['Item Category', 'Count Conv (%)', 'Value Conv (%)', 'AHSP', 'WarrantyPrice', 'WarrantyCount']]
        item_category_display.columns = ['Item Category', 'Count Conv (%)', 'Value Conv (%)', 'AHSP (₹)', 'Warranty Sales (₹)', 'Warranty Units']
        
        # Apply sorting to item category performance table
        item_category_sort_column_mapping = {
            "Count Conv (%)": "Count Conv (%)",
            "Value Conv (%)": "Value Conv (%)", 
            "AHSP (₹)": "AHSP (₹)",
            "Warranty Sales (₹)": "Warranty Sales (₹)",
            "Warranty Units": "Warranty Units"
        }
        item_category_sort_column = item_category_sort_column_mapping.get(sort_by, "Count Conv (%)")
        item_category_display = item_category_display.sort_values(item_category_sort_column, ascending=sort_ascending)

        st.dataframe(item_category_display.style.format({
            'Count Conv (%)': '{:.2f}%',
            'Value Conv (%)': '{:.2f}%',
            'AHSP (₹)': '₹{:.2f}',
            'Warranty Sales (₹)': '₹{:.0f}',
            'Warranty Units': '{:.0f}'
        }), use_container_width=True)

        st.download_button(
            label="📥 Download Item Category Performance as Excel",
            data=to_excel(item_category_display, 'Item Category Performance'),
            file_name=f"item_category_performance_{st.session_state.selected_sheet.lower().replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("ℹ️ No item category data available with current filters.")

    # Insights Section
    st.markdown(f'<h3 class="subheader">💡 Performance Insights & Key Takeaways - {st.session_state.selected_sheet}</h3>', unsafe_allow_html=True)
    
    if not store_summary.empty:
        top_store = store_summary.loc[store_summary['Count Conv (%)'].idxmax()]
        bottom_store = store_summary.loc[store_summary['Count Conv (%)'].idxmin()]
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.success(f"🏆 **Top Performing Store**: {top_store['Store']} with **{top_store['Count Conv (%)']:.2f}%** count conversion")
        
        with col2:
            st.warning(f"📉 **Store Needing Improvement**: {bottom_store['Store']} with **{bottom_store['Count Conv (%)']:.2f}%** count conversion")
    
    if not rbm_summary.empty:
        top_rbm = rbm_summary.loc[rbm_summary['Count Conv (%)'].idxmax()]
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.info(f"👑 **Top Performing RBM**: {top_rbm['RBM']} with **{top_rbm['Count Conv (%)']:.2f}%** count conversion")
        
        with col2:
            st.info(f"💰 **Overall AHSP**: **₹{ahsp:,.2f}**")

else:
    st.error(f"❌ Failed to load data from Google Sheets for {st.session_state.selected_sheet}. Please check the Apps Script URL, Spreadsheet ID, or network connection.")
