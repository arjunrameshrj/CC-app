import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from io import BytesIO
import time

# Streamlit page configuration
st.set_page_config(page_title="Warranty Conversion Dashboard", layout="wide", initial_sidebar_state="expanded")

# --- Google Sheets Integration Configuration ---
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwqzSILXSQDecrzt7G_Y5uIKKYJSOTzo1EI9iiZa0hicYNJ42X6c6oDIQQo9iisbaPr8w/exec"  # Replace with your Apps Script web app URL

# Configure requests with retry logic
session = requests.Session()
retries = Retry(total=3, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
session.mount('https', HTTPAdapter(max_retries=retries))

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
    "2025 SEP",
    "2025 OCT",
    "2025 NOV",
    "2025 DECEMBER"
]

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

# Enhanced CSS for modern, attractive styling with loading animation
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
        
        /* Loading Animation */
        .loading-container {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            padding: 60px 20px;
            text-align: center;
        }
        
        .loading-spinner {
            width: 80px;
            height: 80px;
            border: 8px solid #f3f4f6;
            border-top: 8px solid #3b82f6;
            border-radius: 50%;
            animation: spin 1.5s linear infinite;
            margin-bottom: 20px;
        }
        
        .loading-spinner-small {
            width: 40px;
            height: 40px;
            border: 4px solid #f3f4f6;
            border-top: 4px solid #3b82f6;
            border-radius: 50%;
            animation: spin 1.5s linear infinite;
            margin: 0 auto;
        }
        
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        .loading-text {
            font-size: 1.5em;
            font-weight: 600;
            color: #1e293b;
            margin-bottom: 10px;
        }
        
        .loading-subtext {
            font-size: 1em;
            color: #64748b;
            max-width: 500px;
        }
        
        .pulse {
            animation: pulse 2s infinite;
        }
        
        @keyframes pulse {
            0% { opacity: 1; }
            50% { opacity: 0.7; }
            100% { opacity: 1; }
        }
        
        /* Progress Bar */
        .progress-container {
            width: 100%;
            background-color: #f1f5f9;
            border-radius: 10px;
            margin: 20px 0;
            overflow: hidden;
            height: 12px;
        }
        
        .progress-bar {
            height: 100%;
            background: linear-gradient(90deg, #3b82f6, #06b6d4);
            border-radius: 10px;
            transition: width 0.3s ease;
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
        
        /* Comparison Header */
        .comparison-header {
            color: #1e293b;
            font-size: 1.5em;
            font-weight: 800;
            margin: 30px 0 20px 0;
            padding: 18px 22px;
            background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%);
            border-radius: 14px;
            border-left: 6px solid #f59e0b;
            box-shadow: 0 6px 20px rgba(245, 158, 11, 0.15);
            position: relative;
        }
        
        .comparison-header::after {
            content: '';
            position: absolute;
            bottom: 0;
            left: 22px;
            right: 22px;
            height: 3px;
            background: linear-gradient(90deg, #f59e0b 0%, #d97706 100%);
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
        .stSelectbox>div>div>select,
        .stDateInput>div>div>input {
            border-radius: 10px;
            border: 2px solid #e2e8f0;
            padding: 12px 16px;
            background-color: white;
            transition: all 0.3s ease;
            font-size: 14px;
        }
        
        .stTextInput>div>div>input:focus,
        .stSelectbox>div>div>select:focus,
        .stDateInput>div>div>input:focus {
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
        
        /* Multi-select styling */
        .stMultiSelect [data-baseweb="tag"] {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border-radius: 8px;
        }
        
        /* Comparison metrics */
        .comparison-metric {
            background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%) !important;
            border: 2px solid #f59e0b !important;
        }
        
        .comparison-metric [data-testid="stMetricValue"] {
            background: linear-gradient(135deg, #92400e 0%, #b45309 100%) !important;
            -webkit-background-clip: text !important;
            -webkit-text-fill-color: transparent !important;
            background-clip: text !important;
        }
        
        .positive-change {
            color: #059669 !important;
            font-weight: 700;
        }
        
        .negative-change {
            color: #dc2626 !important;
            font-weight: 700;
        }
    </style>
""", unsafe_allow_html=True)

# Function to convert DataFrame to Excel for download
def to_excel(df, sheet_name='Data'):
    # Shorten sheet name if it's too long for Excel
    if len(sheet_name) > 31:
        sheet_name = sheet_name[:31]
    
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

# Function to display loading animation
def show_loading_animation(message="Loading Data", submessage="Please wait while we fetch your data..."):
    st.markdown(f"""
        <div class="loading-container">
            <div class="loading-spinner"></div>
            <div class="loading-text pulse">{message}</div>
            <div class="loading-subtext">{submessage}</div>
            <div class="progress-container">
                <div class="progress-bar" style="width: 70%;"></div>
            </div>
        </div>
    """, unsafe_allow_html=True)

# Function to format numbers in lakhs, crores, and thousands
def format_indian_currency(value):
    """Format numbers in Indian currency format (Cr, L, T)"""
    if pd.isna(value) or value == 0:
        return "0"
    
    value = float(value)
    
    # For crores (>= 1,00,00,000)
    if abs(value) >= 10000000:
        return f"{value/10000000:.1f}Cr"
    # For lakhs (>= 1,00,000)
    elif abs(value) >= 100000:
        return f"{value/100000:.1f}L"
    # For thousands (>= 1,000)
    elif abs(value) >= 1000:
        return f"{value/1000:.0f}T"
    else:
        return f"{value:.0f}"

# Function to create product-wise monthly summary table with filters applied
def create_product_monthly_summary(individual_data, filters, category_column, replacement_filter, speaker_filter):
    """Create a product-wise monthly summary table with formatted values and filters applied"""
    
    # Define the main product categories we want to track
    main_categories = ['AC', 'TV', 'REFRIGERATOR', 'WASHING MACHINE', 'MICROWAVE OVEN']
    
    # Initialize the summary dictionary
    summary_data = []
    
    # Get all available months
    available_months = list(individual_data.keys())
    
    for category in main_categories:
        category_data = {'Product': category}
        
        for month in available_months:
            month_df = individual_data[month]
            
            # Apply filters to the month data
            filtered_month_data = apply_comparison_filters(
                month_df, 
                filters, 
                category_column,
                replacement_filter,
                speaker_filter
            )
            
            # Filter data for this category
            category_sales = filtered_month_data[filtered_month_data['Item Category'] == category]
            
            # Calculate warranty sales for this category
            warranty_sales = category_sales['WarrantyPrice'].sum()
            
            # Format the value
            category_data[month] = format_indian_currency(warranty_sales)
        
        summary_data.append(category_data)
    
    # Add "OTHERS" category
    others_data = {'Product': 'OTHERS'}
    
    for month in available_months:
        month_df = individual_data[month]
        
        # Apply filters to the month data
        filtered_month_data = apply_comparison_filters(
            month_df, 
            filters, 
            category_column,
            replacement_filter,
            speaker_filter
        )
        
        # Filter data for non-main categories
        other_categories_sales = filtered_month_data[~filtered_month_data['Item Category'].isin(main_categories)]
        
        # Calculate warranty sales for other categories
        others_warranty_sales = other_categories_sales['WarrantyPrice'].sum()
        
        # Format the value
        others_data[month] = format_indian_currency(others_warranty_sales)
    
    summary_data.append(others_data)
    
    # Add "TOTAL" row
    total_data = {'Product': 'TOTAL'}
    
    for month in available_months:
        month_df = individual_data[month]
        
        # Apply filters to the month data
        filtered_month_data = apply_comparison_filters(
            month_df, 
            filters, 
            category_column,
            replacement_filter,
            speaker_filter
        )
        
        # Calculate total warranty sales for the month
        total_warranty_sales = filtered_month_data['WarrantyPrice'].sum()
        
        # Format the value
        total_data[month] = format_indian_currency(total_warranty_sales)
    
    summary_data.append(total_data)
    
    # Convert to DataFrame
    summary_df = pd.DataFrame(summary_data)
    
    return summary_df

# Function to create RBM-wise monthly summary table with filters applied
def create_rbm_monthly_summary(individual_data, filters, category_column, replacement_filter, speaker_filter):
    """Create an RBM-wise monthly summary table with formatted values and filters applied"""
    
    # Initialize the summary dictionary
    summary_data = []
    
    # Get all available months
    available_months = list(individual_data.keys())
    
    # Get all unique RBMs across all months
    all_rbms = set()
    for month_data in individual_data.values():
        filtered_month_data = apply_comparison_filters(
            month_data, 
            filters, 
            category_column,
            replacement_filter,
            speaker_filter
        )
        all_rbms.update(filtered_month_data['RBM'].unique())
    
    all_rbms = sorted(list(all_rbms))
    
    for rbm in all_rbms:
        rbm_data = {'RBM': rbm}
        
        for month in available_months:
            month_df = individual_data[month]
            
            # Apply filters to the month data
            filtered_month_data = apply_comparison_filters(
                month_df, 
                filters, 
                category_column,
                replacement_filter,
                speaker_filter
            )
            
            # Filter data for this RBM
            rbm_sales = filtered_month_data[filtered_month_data['RBM'] == rbm]
            
            # Calculate warranty sales for this RBM
            warranty_sales = rbm_sales['WarrantyPrice'].sum()
            
            # Format the value
            rbm_data[month] = format_indian_currency(warranty_sales)
        
        summary_data.append(rbm_data)
    
    # Add "TOTAL" row
    total_data = {'RBM': 'TOTAL'}
    
    for month in available_months:
        month_df = individual_data[month]
        
        # Apply filters to the month data
        filtered_month_data = apply_comparison_filters(
            month_df, 
            filters, 
            category_column,
            replacement_filter,
            speaker_filter
        )
        
        # Calculate total warranty sales for the month
        total_warranty_sales = filtered_month_data['WarrantyPrice'].sum()
        
        # Format the value
        total_data[month] = format_indian_currency(total_warranty_sales)
    
    summary_data.append(total_data)
    
    # Convert to DataFrame
    summary_df = pd.DataFrame(summary_data)
    
    return summary_df

# Function to create RBM-wise monthly value conversion summary table with filters applied
def create_rbm_monthly_value_conversion_summary(individual_data, filters, category_column, replacement_filter, speaker_filter):
    """Create an RBM-wise monthly value conversion summary table with filters applied"""
    
    # Initialize the summary dictionary
    summary_data = []
    
    # Get all available months
    available_months = list(individual_data.keys())
    
    # Get all unique RBMs across all months
    all_rbms = set()
    for month_data in individual_data.values():
        filtered_month_data = apply_comparison_filters(
            month_data, 
            filters, 
            category_column,
            replacement_filter,
            speaker_filter
        )
        all_rbms.update(filtered_month_data['RBM'].unique())
    
    all_rbms = sorted(list(all_rbms))
    
    for rbm in all_rbms:
        rbm_data = {'RBM': rbm}
        
        for month in available_months:
            month_df = individual_data[month]
            
            # Apply filters to the month data
            filtered_month_data = apply_comparison_filters(
                month_df, 
                filters, 
                category_column,
                replacement_filter,
                speaker_filter
            )
            
            # Filter data for this RBM
            rbm_filtered_data = filtered_month_data[filtered_month_data['RBM'] == rbm]
            
            # Calculate value conversion for this RBM
            total_sales = rbm_filtered_data['TotalSoldPrice'].sum()
            warranty_sales = rbm_filtered_data['WarrantyPrice'].sum()
            value_conversion = (warranty_sales / total_sales * 100) if total_sales > 0 else 0
            
            # Format the value as percentage
            rbm_data[month] = f"{value_conversion:.2f}%"
        
        summary_data.append(rbm_data)
    
    # Add "TOTAL" row
    total_data = {'RBM': 'TOTAL'}
    
    for month in available_months:
        month_df = individual_data[month]
        
        # Apply filters to the month data
        filtered_month_data = apply_comparison_filters(
            month_df, 
            filters, 
            category_column,
            replacement_filter,
            speaker_filter
        )
        
        # Calculate overall value conversion for the month
        total_sales_month = filtered_month_data['TotalSoldPrice'].sum()
        warranty_sales_month = filtered_month_data['WarrantyPrice'].sum()
        value_conversion_month = (warranty_sales_month / total_sales_month * 100) if total_sales_month > 0 else 0
        
        # Format the value as percentage
        total_data[month] = f"{value_conversion_month:.2f}%"
    
    summary_data.append(total_data)
    
    # Convert to DataFrame
    summary_df = pd.DataFrame(summary_data)
    
    return summary_df

# NEW FUNCTION: Create product-wise monthly value conversion summary table
def create_product_monthly_value_conversion_summary(individual_data, filters, category_column, replacement_filter, speaker_filter):
    """Create a product-wise monthly value conversion summary table with filters applied"""
    
    # Define the main product categories we want to track
    main_categories = ['AC', 'TV', 'REFRIGERATOR', 'WASHING MACHINE', 'MICROWAVE OVEN']
    
    # Initialize the summary dictionary
    summary_data = []
    
    # Get all available months
    available_months = list(individual_data.keys())
    
    for category in main_categories:
        category_data = {'Product': category}
        
        for month in available_months:
            month_df = individual_data[month]
            
            # Apply filters to the month data
            filtered_month_data = apply_comparison_filters(
                month_df, 
                filters, 
                category_column,
                replacement_filter,
                speaker_filter
            )
            
            # Filter data for this category
            category_data_filtered = filtered_month_data[filtered_month_data['Item Category'] == category]
            
            # Calculate value conversion for this category
            total_sales = category_data_filtered['TotalSoldPrice'].sum()
            warranty_sales = category_data_filtered['WarrantyPrice'].sum()
            value_conversion = (warranty_sales / total_sales * 100) if total_sales > 0 else 0
            
            # Format the value as percentage
            category_data[month] = f"{value_conversion:.2f}%"
        
        summary_data.append(category_data)
    
    # Add "OTHERS" category
    others_data = {'Product': 'OTHERS'}
    
    for month in available_months:
        month_df = individual_data[month]
        
        # Apply filters to the month data
        filtered_month_data = apply_comparison_filters(
            month_df, 
            filters, 
            category_column,
            replacement_filter,
            speaker_filter
        )
        
        # Filter data for non-main categories
        other_categories_data = filtered_month_data[~filtered_month_data['Item Category'].isin(main_categories)]
        
        # Calculate value conversion for other categories
        total_sales_others = other_categories_data['TotalSoldPrice'].sum()
        warranty_sales_others = other_categories_data['WarrantyPrice'].sum()
        value_conversion_others = (warranty_sales_others / total_sales_others * 100) if total_sales_others > 0 else 0
        
        # Format the value as percentage
        others_data[month] = f"{value_conversion_others:.2f}%"
    
    summary_data.append(others_data)
    
    # Add "TOTAL" row
    total_data = {'Product': 'TOTAL'}
    
    for month in available_months:
        month_df = individual_data[month]
        
        # Apply filters to the month data
        filtered_month_data = apply_comparison_filters(
            month_df, 
            filters, 
            category_column,
            replacement_filter,
            speaker_filter
        )
        
        # Calculate overall value conversion for the month
        total_sales_month = filtered_month_data['TotalSoldPrice'].sum()
        warranty_sales_month = filtered_month_data['WarrantyPrice'].sum()
        value_conversion_month = (warranty_sales_month / total_sales_month * 100) if total_sales_month > 0 else 0
        
        # Format the value as percentage
        total_data[month] = f"{value_conversion_month:.2f}%"
    
    summary_data.append(total_data)
    
    # Convert to DataFrame
    summary_df = pd.DataFrame(summary_data)
    
    return summary_df

# Function to calculate comparison metrics for all tables
def calculate_comparison(month1_data, month2_data, month1_name, month2_name):
    """Calculate comparison metrics between two months for all tables"""
    comparison_data = {}
    
    # Store Performance Comparison
    store_comp1 = month1_data.groupby('Store').agg({
        'WarrantyPrice': 'sum',
        'TotalSoldPrice': 'sum',
        'WarrantyCount': 'sum',
        'TotalCount': 'sum'
    }).reset_index()
    
    store_comp2 = month2_data.groupby('Store').agg({
        'WarrantyPrice': 'sum',
        'TotalSoldPrice': 'sum',
        'WarrantyCount': 'sum',
        'TotalCount': 'sum'
    }).reset_index()
    
    # Calculate metrics for both months - CORRECTED: Use WarrantyCount for warranty units and count conversion
    store_comp1['Value Conv (%)'] = (store_comp1['WarrantyPrice'] / store_comp1['TotalSoldPrice'] * 100).round(2)
    store_comp1['Count Conv (%)'] = (store_comp1['WarrantyCount'] / store_comp1['TotalCount'] * 100).round(2)
    store_comp1['AHSP'] = (store_comp1['WarrantyPrice'] / store_comp1['WarrantyCount']).where(store_comp1['WarrantyCount'] > 0, 0).round(2)
    
    store_comp2['Value Conv (%)'] = (store_comp2['WarrantyPrice'] / store_comp2['TotalSoldPrice'] * 100).round(2)
    store_comp2['Count Conv (%)'] = (store_comp2['WarrantyCount'] / store_comp2['TotalCount'] * 100).round(2)
    store_comp2['AHSP'] = (store_comp2['WarrantyPrice'] / store_comp2['WarrantyCount']).where(store_comp2['WarrantyCount'] > 0, 0).round(2)
    
    # Merge for comparison
    store_comparison = pd.merge(store_comp1, store_comp2, on='Store', suffixes=(f'_{month1_name}', f'_{month2_name}'))
    
    # Calculate changes
    store_comparison['Value Conv Change'] = store_comparison[f'Value Conv (%)_{month2_name}'] - store_comparison[f'Value Conv (%)_{month1_name}']
    store_comparison['Count Conv Change'] = store_comparison[f'Count Conv (%)_{month2_name}'] - store_comparison[f'Count Conv (%)_{month1_name}']
    store_comparison['AHSP Change'] = store_comparison[f'AHSP_{month2_name}'] - store_comparison[f'AHSP_{month1_name}']
    store_comparison['Warranty Sales Change'] = store_comparison[f'WarrantyPrice_{month2_name}'] - store_comparison[f'WarrantyPrice_{month1_name}']
    store_comparison['Warranty Units Change'] = store_comparison[f'WarrantyCount_{month2_name}'] - store_comparison[f'WarrantyCount_{month1_name}']
    
    # Calculate percentage changes
    store_comparison['Warranty Sales Change %'] = ((store_comparison[f'WarrantyPrice_{month2_name}'] - store_comparison[f'WarrantyPrice_{month1_name}']) / store_comparison[f'WarrantyPrice_{month1_name}'] * 100).where(store_comparison[f'WarrantyPrice_{month1_name}'] > 0, 0).round(2)
    store_comparison['Warranty Units Change %'] = ((store_comparison[f'WarrantyCount_{month2_name}'] - store_comparison[f'WarrantyCount_{month1_name}']) / store_comparison[f'WarrantyCount_{month1_name}'] * 100).where(store_comparison[f'WarrantyCount_{month1_name}'] > 0, 0).round(2)
    
    comparison_data['store_comparison'] = store_comparison
    
    # Staff Performance Comparison
    staff_comp1 = month1_data.groupby(['Staff Name', 'Store']).agg({
        'WarrantyPrice': 'sum',
        'TotalSoldPrice': 'sum',
        'WarrantyCount': 'sum',
        'TotalCount': 'sum'
    }).reset_index()
    
    staff_comp2 = month2_data.groupby(['Staff Name', 'Store']).agg({
        'WarrantyPrice': 'sum',
        'TotalSoldPrice': 'sum',
        'WarrantyCount': 'sum',
        'TotalCount': 'sum'
    }).reset_index()
    
    staff_comp1['Value Conv (%)'] = (staff_comp1['WarrantyPrice'] / staff_comp1['TotalSoldPrice'] * 100).round(2)
    staff_comp1['Count Conv (%)'] = (staff_comp1['WarrantyCount'] / staff_comp1['TotalCount'] * 100).round(2)
    staff_comp1['AHSP'] = (staff_comp1['WarrantyPrice'] / staff_comp1['WarrantyCount']).where(staff_comp1['WarrantyCount'] > 0, 0).round(2)
    
    staff_comp2['Value Conv (%)'] = (staff_comp2['WarrantyPrice'] / staff_comp2['TotalSoldPrice'] * 100).round(2)
    staff_comp2['Count Conv (%)'] = (staff_comp2['WarrantyCount'] / staff_comp2['TotalCount'] * 100).round(2)
    staff_comp2['AHSP'] = (staff_comp2['WarrantyPrice'] / staff_comp2['WarrantyCount']).where(staff_comp2['WarrantyCount'] > 0, 0).round(2)
    
    staff_comparison = pd.merge(staff_comp1, staff_comp2, on=['Staff Name', 'Store'], suffixes=(f'_{month1_name}', f'_{month2_name}'))
    
    staff_comparison['Value Conv Change'] = staff_comparison[f'Value Conv (%)_{month2_name}'] - staff_comparison[f'Value Conv (%)_{month1_name}']
    staff_comparison['Count Conv Change'] = staff_comparison[f'Count Conv (%)_{month2_name}'] - staff_comparison[f'Count Conv (%)_{month1_name}']
    staff_comparison['AHSP Change'] = staff_comparison[f'AHSP_{month2_name}'] - staff_comparison[f'AHSP_{month1_name}']
    staff_comparison['Warranty Sales Change'] = staff_comparison[f'WarrantyPrice_{month2_name}'] - staff_comparison[f'WarrantyPrice_{month1_name}']
    staff_comparison['Warranty Units Change'] = staff_comparison[f'WarrantyCount_{month2_name}'] - staff_comparison[f'WarrantyCount_{month1_name}']
    
    comparison_data['staff_comparison'] = staff_comparison
    
    # RBM Performance Comparison
    rbm_comp1 = month1_data.groupby('RBM').agg({
        'WarrantyPrice': 'sum',
        'TotalSoldPrice': 'sum',
        'WarrantyCount': 'sum',
        'TotalCount': 'sum'
    }).reset_index()
    
    rbm_comp2 = month2_data.groupby('RBM').agg({
        'WarrantyPrice': 'sum',
        'TotalSoldPrice': 'sum',
        'WarrantyCount': 'sum',
        'TotalCount': 'sum'
    }).reset_index()
    
    rbm_comp1['Value Conv (%)'] = (rbm_comp1['WarrantyPrice'] / rbm_comp1['TotalSoldPrice'] * 100).round(2)
    rbm_comp1['Count Conv (%)'] = (rbm_comp1['WarrantyCount'] / rbm_comp1['TotalCount'] * 100).round(2)
    rbm_comp1['AHSP'] = (rbm_comp1['WarrantyPrice'] / rbm_comp1['WarrantyCount']).where(rbm_comp1['WarrantyCount'] > 0, 0).round(2)
    
    rbm_comp2['Value Conv (%)'] = (rbm_comp2['WarrantyPrice'] / rbm_comp2['TotalSoldPrice'] * 100).round(2)
    rbm_comp2['Count Conv (%)'] = (rbm_comp2['WarrantyCount'] / rbm_comp2['TotalCount'] * 100).round(2)
    rbm_comp2['AHSP'] = (rbm_comp2['WarrantyPrice'] / rbm_comp2['WarrantyCount']).where(rbm_comp2['WarrantyCount'] > 0, 0).round(2)
    
    rbm_comparison = pd.merge(rbm_comp1, rbm_comp2, on='RBM', suffixes=(f'_{month1_name}', f'_{month2_name}'))
    
    rbm_comparison['Value Conv Change'] = rbm_comparison[f'Value Conv (%)_{month2_name}'] - rbm_comparison[f'Value Conv (%)_{month1_name}']
    rbm_comparison['Count Conv Change'] = rbm_comparison[f'Count Conv (%)_{month2_name}'] - rbm_comparison[f'Count Conv (%)_{month1_name}']
    rbm_comparison['AHSP Change'] = rbm_comparison[f'AHSP_{month2_name}'] - rbm_comparison[f'AHSP_{month1_name}']
    rbm_comparison['Warranty Sales Change'] = rbm_comparison[f'WarrantyPrice_{month2_name}'] - rbm_comparison[f'WarrantyPrice_{month1_name}']
    rbm_comparison['Warranty Units Change'] = rbm_comparison[f'WarrantyCount_{month2_name}'] - rbm_comparison[f'WarrantyCount_{month1_name}']
    
    comparison_data['rbm_comparison'] = rbm_comparison
    
    # Product Category Performance Comparison
    category_comp1 = month1_data.groupby('Item Category').agg({
        'WarrantyPrice': 'sum',
        'TotalSoldPrice': 'sum',
        'WarrantyCount': 'sum',
        'TotalCount': 'sum'
    }).reset_index()
    
    category_comp2 = month2_data.groupby('Item Category').agg({
        'WarrantyPrice': 'sum',
        'TotalSoldPrice': 'sum',
        'WarrantyCount': 'sum',
        'TotalCount': 'sum'
    }).reset_index()
    
    category_comp1['Value Conv (%)'] = (category_comp1['WarrantyPrice'] / category_comp1['TotalSoldPrice'] * 100).round(2)
    category_comp1['Count Conv (%)'] = (category_comp1['WarrantyCount'] / category_comp1['TotalCount'] * 100).round(2)
    category_comp1['AHSP'] = (category_comp1['WarrantyPrice'] / category_comp1['WarrantyCount']).where(category_comp1['WarrantyCount'] > 0, 0).round(2)
    
    category_comp2['Value Conv (%)'] = (category_comp2['WarrantyPrice'] / category_comp2['TotalSoldPrice'] * 100).round(2)
    category_comp2['Count Conv (%)'] = (category_comp2['WarrantyCount'] / category_comp2['TotalCount'] * 100).round(2)
    category_comp2['AHSP'] = (category_comp2['WarrantyPrice'] / category_comp2['WarrantyCount']).where(category_comp2['WarrantyCount'] > 0, 0).round(2)
    
    category_comparison = pd.merge(category_comp1, category_comp2, on='Item Category', suffixes=(f'_{month1_name}', f'_{month2_name}'))
    
    category_comparison['Value Conv Change'] = category_comparison[f'Value Conv (%)_{month2_name}'] - category_comparison[f'Value Conv (%)_{month1_name}']
    category_comparison['Count Conv Change'] = category_comparison[f'Count Conv (%)_{month2_name}'] - category_comparison[f'Count Conv (%)_{month1_name}']
    category_comparison['AHSP Change'] = category_comparison[f'AHSP_{month2_name}'] - category_comparison[f'AHSP_{month1_name}']
    category_comparison['Warranty Sales Change'] = category_comparison[f'WarrantyPrice_{month2_name}'] - category_comparison[f'WarrantyPrice_{month1_name}']
    category_comparison['Warranty Units Change'] = category_comparison[f'WarrantyCount_{month2_name}'] - category_comparison[f'WarrantyCount_{month1_name}']
    
    comparison_data['category_comparison'] = category_comparison
    
    # Overall KPIs comparison
    overall_kpis = {
        'total_warranty_1': month1_data['WarrantyPrice'].sum(),
        'total_warranty_2': month2_data['WarrantyPrice'].sum(),
        'total_units_1': month1_data['TotalCount'].sum(),
        'total_units_2': month2_data['TotalCount'].sum(),
        'warranty_units_1': month1_data['WarrantyCount'].sum(),
        'warranty_units_2': month2_data['WarrantyCount'].sum(),
        'count_conv_1': (month1_data['WarrantyCount'].sum() / month1_data['TotalCount'].sum() * 100) if month1_data['TotalCount'].sum() > 0 else 0,
        'count_conv_2': (month2_data['WarrantyCount'].sum() / month2_data['TotalCount'].sum() * 100) if month2_data['TotalCount'].sum() > 0 else 0,
        'value_conv_1': (month1_data['WarrantyPrice'].sum() / month1_data['TotalSoldPrice'].sum() * 100) if month1_data['TotalSoldPrice'].sum() > 0 else 0,
        'value_conv_2': (month2_data['WarrantyPrice'].sum() / month2_data['TotalSoldPrice'].sum() * 100) if month2_data['TotalSoldPrice'].sum() > 0 else 0,
        'ahsp_1': (month1_data['WarrantyPrice'].sum() / month1_data['WarrantyCount'].sum()) if month1_data['WarrantyCount'].sum() > 0 else 0,
        'ahsp_2': (month2_data['WarrantyPrice'].sum() / month2_data['WarrantyCount'].sum()) if month2_data['WarrantyCount'].sum() > 0 else 0,
    }
    
    # Calculate overall changes
    overall_kpis['warranty_change'] = overall_kpis['total_warranty_2'] - overall_kpis['total_warranty_1']
    overall_kpis['warranty_change_pct'] = (overall_kpis['warranty_change'] / overall_kpis['total_warranty_1'] * 100) if overall_kpis['total_warranty_1'] > 0 else 0
    overall_kpis['warranty_units_change'] = overall_kpis['warranty_units_2'] - overall_kpis['warranty_units_1']
    overall_kpis['warranty_units_change_pct'] = (overall_kpis['warranty_units_change'] / overall_kpis['warranty_units_1'] * 100) if overall_kpis['warranty_units_1'] > 0 else 0
    overall_kpis['count_conv_change'] = overall_kpis['count_conv_2'] - overall_kpis['count_conv_1']
    overall_kpis['value_conv_change'] = overall_kpis['value_conv_2'] - overall_kpis['value_conv_1']
    overall_kpis['ahsp_change'] = overall_kpis['ahsp_2'] - overall_kpis['ahsp_1']
    
    comparison_data['overall_kpis'] = overall_kpis
    
    return comparison_data

# Function to create monthly warranty sales trend chart
def create_monthly_trend_chart(individual_data):
    """Create a line chart showing warranty sales trend across all months"""
    monthly_trend = []
    
    for month_name, month_data in individual_data.items():
        total_warranty = month_data['WarrantyPrice'].sum()
        monthly_trend.append({
            'Month': month_name,
            'Warranty Sales': total_warranty
        })
    
    trend_df = pd.DataFrame(monthly_trend)
    
    if not trend_df.empty:
        fig = px.line(
            trend_df, 
            x='Month', 
            y='Warranty Sales',
            title='📈 Monthly Warranty Sales Trend',
            markers=True,
            line_shape='spline'
        )
        
        fig.update_layout(
            xaxis_title='Month',
            yaxis_title='Warranty Sales (₹)',
            hovermode='x unified',
            showlegend=False,
            plot_bgcolor='white',
            height=400
        )
        
        fig.update_traces(
            line=dict(color='#3b82f6', width=4),
            marker=dict(size=8, color='#1e40af')
        )
        
        return fig
    return None

# Function to add total row to any dataframe with numeric columns
def add_total_row(df, group_by_columns, numeric_columns):
    """Add a total row to a dataframe"""
    if df.empty:
        return df
    
    # Calculate totals for numeric columns
    total_values = {}
    for col in df.columns:
        if col in numeric_columns:
            total_values[col] = df[col].sum()
        elif col in group_by_columns:
            total_values[col] = 'Total'
        else:
            total_values[col] = ''
    
    # Create total row
    total_row = pd.DataFrame([total_values])
    
    # Concatenate with original dataframe
    result_df = pd.concat([df, total_row], ignore_index=True)
    
    return result_df

# Session state initialization
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'current_df' not in st.session_state:
    st.session_state.current_df = None
if 'selected_sheets' not in st.session_state:
    st.session_state.selected_sheets = [SHEETS[11]]  # Default to first sheet
if 'individual_month_data' not in st.session_state:
    st.session_state.individual_month_data = {}
if 'comparison_filters' not in st.session_state:
    st.session_state.comparison_filters = {
        'selected_bdm': 'All',
        'selected_rbm': 'All', 
        'selected_store': 'All',
        'selected_category': 'All',
        'selected_staff': 'All',
        'replacement_filter': False,
        'speaker_filter': False,
        'future_filter': False
    }

# Sidebar content
with st.sidebar:
    st.markdown('<h2 style="color: #1e293b; font-weight: 700; margin-bottom: 20px;">🔍 Dashboard Controls</h2>', unsafe_allow_html=True)
    st.markdown('<hr>', unsafe_allow_html=True)
    
    # Month selection - Multi-select for multiple sheets with "All" option
    month_options = ["All"] + SHEETS
    
    selected_sheets = st.multiselect(
        "📅 Select Month(s) for Comparison", 
        month_options, 
        default=st.session_state.selected_sheets,
        help="Select 'All' to combine all months data, or select specific months for comparison"
    )
    
    # Handle "All" selection logic
    if "All" in selected_sheets:
        # If "All" is selected, select all months and remove individual month selections
        selected_sheets = ["All"]
    elif len(selected_sheets) > 1 and "All" in selected_sheets:
        # If multiple selections including "All", keep only "All"
        selected_sheets = ["All"]
    
    # Update session state if selection changes
    if selected_sheets != st.session_state.selected_sheets:
        st.session_state.selected_sheets = selected_sheets
        st.session_state.current_df = None
        st.session_state.data_loaded = False
        st.cache_data.clear()
    
    # Refresh Button
    if st.button("🔄 Refresh Data"):
        st.cache_data.clear()
        st.session_state.current_df = None
        st.session_state.data_loaded = False
        st.session_state.individual_month_data = {}
        st.rerun()

# Generate dashboard title based on selected sheets
if len(st.session_state.selected_sheets) == 0:
    dashboard_title = "📊 Warranty Conversion Analysis Dashboard - No Month Selected"
elif "All" in st.session_state.selected_sheets:
    dashboard_title = "📊 Warranty Conversion Analysis Dashboard - All Months Combined"
elif len(st.session_state.selected_sheets) == 1:
    dashboard_title = f"📊 Warranty Conversion Analysis Dashboard - {st.session_state.selected_sheets[0]}"
else:
    if len(st.session_state.selected_sheets) == len(SHEETS):
        dashboard_title = "📊 Warranty Conversion Analysis Dashboard - All Months Comparison"
    else:
        dashboard_title = f"📊 Warranty Conversion Analysis Dashboard - {len(st.session_state.selected_sheets)} Months Comparison"

# Main dashboard header
st.markdown(f'<div class="main-header">{dashboard_title}</div>', unsafe_allow_html=True)

# Load data function
required_columns = ['Item Category', 'BDM', 'RBM', 'Store', 'Staff Name', 'TotalSoldPrice', 'WarrantyPrice', 'TotalCount', 'WarrantyCount']

@st.cache_data
def load_individual_month(sheet_name):
    """Load individual month data"""
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
        
        # CORRECTED: Use WarrantyCount for warranty units and count conversion
        df['Conversion% (Count)'] = (df['WarrantyCount'] / df['TotalCount'] * 100).round(2)
        df['Conversion% (Price)'] = (df['WarrantyPrice'] / df['TotalSoldPrice'] * 100).where(df['TotalSoldPrice'] > 0, 0).round(2)
        df['AHSP'] = (df['WarrantyPrice'] / df['WarrantyCount']).where(df['WarrantyCount'] > 0, 0).round(2)
        df['Month'] = sheet_name
        
        return df
    except Exception as e:
        st.error(f"❌ Error loading data for {sheet_name}: {str(e)}")
        return None

@st.cache_data
def load_all_data(sheet_names):
    """Load data for all selected sheets"""
    try:
        # Show loading animation
        with st.spinner(''):
            # Handle "All" selection
            if "All" in sheet_names:
                actual_sheets = SHEETS
                show_loading_animation(
                    "Loading All Months Data", 
                    "Combining data from all available months..."
                )
            elif len(sheet_names) > 1:
                show_loading_animation(
                    f"Loading {len(sheet_names)} Months Data", 
                    f"Loading data for comparison: {', '.join(sheet_names)}"
                )
            else:
                show_loading_animation(
                    f"Loading {sheet_names[0]} Data", 
                    "Fetching data from Google Sheets and processing..."
                )
            
            # Simulate loading time for better UX
            time.sleep(1.5)
            
            all_dfs = []
            individual_data = {}
            
            # Determine which sheets to load
            sheets_to_load = SHEETS if "All" in sheet_names else sheet_names
            
            for sheet_name in sheets_to_load:
                df = load_individual_month(sheet_name)
                if df is not None:
                    all_dfs.append(df)
                    individual_data[sheet_name] = df
                else:
                    st.error(f"Failed to load data for {sheet_name}")
            
            if not all_dfs:
                st.error("❌ No data could be loaded from any selected sheets.")
                return None, None
            
            # Combine all DataFrames
            combined_df = pd.concat(all_dfs, ignore_index=True)
            return combined_df, individual_data
            
    except Exception as e:
        st.error(f"❌ Error loading data from Google Sheets: {str(e)}")
        return None, None

# Load data from Google Sheets
if st.session_state.current_df is None and st.session_state.selected_sheets:
    combined_df, individual_data = load_all_data(st.session_state.selected_sheets)
    if combined_df is not None:
        st.session_state.current_df = combined_df
        st.session_state.individual_month_data = individual_data
        st.session_state.data_loaded = True
        if "All" in st.session_state.selected_sheets:
            st.success(f"✅ Data loaded successfully for all {len(SHEETS)} months combined!")
        elif len(st.session_state.selected_sheets) > 1:
            st.success(f"✅ Data loaded successfully for {len(st.session_state.selected_sheets)} months comparison!")
        else:
            st.success(f"✅ Data loaded successfully for {st.session_state.selected_sheets[0]}!")
        st.rerun()

# Now that we have data, set up the filters in sidebar
if st.session_state.data_loaded and st.session_state.current_df is not None:
    df = st.session_state.current_df
    individual_data = st.session_state.individual_month_data
    
    with st.sidebar:
        # Sidebar filters
        st.markdown('<hr>', unsafe_allow_html=True)
        st.markdown('<h3 style="color: #1e293b; font-weight: 600; margin-top: 20px;">⚙️ Filters</h3>', unsafe_allow_html=True)
        
        replacement_filter = st.checkbox("🔄 Show Replacement Warranty Categories Only", value=st.session_state.comparison_filters['replacement_filter'])
        speaker_filter = st.checkbox("🔊 Show Speaker Categories Only", value=st.session_state.comparison_filters['speaker_filter'])
        future_filter = st.checkbox("🏢 Show only FUTURE stores", value=st.session_state.comparison_filters['future_filter'])
        
        # Update session state for filters
        st.session_state.comparison_filters['replacement_filter'] = replacement_filter
        st.session_state.comparison_filters['speaker_filter'] = speaker_filter
        st.session_state.comparison_filters['future_filter'] = future_filter
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
        selected_bdm = st.selectbox("👔 BDM", bdm_options, index=bdm_options.index(st.session_state.comparison_filters['selected_bdm']) if st.session_state.comparison_filters['selected_bdm'] in bdm_options else 0)
        selected_rbm = st.selectbox("👤 RBM", rbm_options, index=rbm_options.index(st.session_state.comparison_filters['selected_rbm']) if st.session_state.comparison_filters['selected_rbm'] in rbm_options else 0)
        selected_store = st.selectbox("🏪 Store", store_options, index=store_options.index(st.session_state.comparison_filters['selected_store']) if st.session_state.comparison_filters['selected_store'] in store_options else 0)
        selected_category = st.selectbox(f"📦 {category_column}", category_options, index=category_options.index(st.session_state.comparison_filters['selected_category']) if st.session_state.comparison_filters['selected_category'] in category_options else 0)
        if selected_rbm != 'All':
            staff_options = ['All'] + sorted(df[df['RBM'] == selected_rbm]['Staff Name'].unique().tolist())
        else:
            staff_options = ['All'] + sorted(df['Staff Name'].unique().tolist())
        selected_staff = st.selectbox("👨‍💼 Staff", staff_options, index=staff_options.index(st.session_state.comparison_filters['selected_staff']) if st.session_state.comparison_filters['selected_staff'] in staff_options else 0)

        # Update session state for dropdown filters
        st.session_state.comparison_filters['selected_bdm'] = selected_bdm
        st.session_state.comparison_filters['selected_rbm'] = selected_rbm
        st.session_state.comparison_filters['selected_store'] = selected_store
        st.session_state.comparison_filters['selected_category'] = selected_category
        st.session_state.comparison_filters['selected_staff'] = selected_staff

    # Apply filters function for comparison data
    def apply_comparison_filters(data, filters, category_column, replacement_filter, speaker_filter):
        """Apply filters to comparison data"""
        filtered_data = data.copy()
        
        # Apply replacement or speaker filter first
        if replacement_filter:
            replacement_categories = ['FAN', 'MIXER GRINDER', 'IRON BOX', 'ELECTRIC KETTLE', 'OTG', 'STEAMER', 'INDUCTION COOKER']
            filtered_data = filtered_data[filtered_data['Replacement Category'].isin(replacement_categories)]
        elif speaker_filter:
            speaker_categories = ['SOUND BAR', 'PARTY SPEAKER', 'BLUETOOTH SPEAKER', 'HOME THEATRE']
            filtered_data = filtered_data[filtered_data['Item Category'].isin(speaker_categories)]
        
        # Apply other filters
        if filters['selected_bdm'] != 'All':
            filtered_data = filtered_data[filtered_data['BDM'] == filters['selected_bdm']]
        if filters['selected_rbm'] != 'All':
            filtered_data = filtered_data[filtered_data['RBM'] == filters['selected_rbm']]
        if filters['selected_store'] != 'All':
            filtered_data = filtered_data[filtered_data['Store'] == filters['selected_store']]
        if filters['selected_category'] != 'All':
            filtered_data = filtered_data[filtered_data[category_column] == filters['selected_category']]
        if filters['selected_staff'] != 'All':
            filtered_data = filtered_data[filtered_data['Staff Name'] == filters['selected_staff']]
        if filters['future_filter']:
            filtered_data = filtered_data[filtered_data['Store'].str.contains('FUTURE', case=True)]
            
        return filtered_data

    # Apply filters to main data
    filtered_df = apply_comparison_filters(
        df, 
        st.session_state.comparison_filters, 
        category_column,
        replacement_filter,
        speaker_filter
    )

    if filtered_df.empty:
        st.warning("⚠️ No data matches your filters.")
        st.stop()

    # Generate period text for subheaders
    if "All" in st.session_state.selected_sheets:
        period_text = "All Months Combined"
    elif len(st.session_state.selected_sheets) == 1:
        period_text = st.session_state.selected_sheets[0]
    else:
        period_text = f"{len(st.session_state.selected_sheets)} Selected Months"

    # COMPARISON SECTION - Show only when exactly 2 months are selected (and not "All")
    if len(st.session_state.selected_sheets) == 2 and "All" not in st.session_state.selected_sheets:
        st.markdown(f'<div class="comparison-header">📊 Monthly Comparison Analysis</div>', unsafe_allow_html=True)
        
        month1 = st.session_state.selected_sheets[0]
        month2 = st.session_state.selected_sheets[1]
        
        st.markdown(f'### 🔄 Comparison: {month1} vs {month2}')
        
        # Apply filters to individual month data for comparison
        month1_filtered = apply_comparison_filters(
            individual_data[month1], 
            st.session_state.comparison_filters, 
            category_column,
            replacement_filter,
            speaker_filter
        )
        month2_filtered = apply_comparison_filters(
            individual_data[month2], 
            st.session_state.comparison_filters, 
            category_column,
            replacement_filter,
            speaker_filter
        )
        
        # Calculate comparison metrics for all tables with filtered data
        comparison_data = calculate_comparison(
            month1_filtered, 
            month2_filtered, 
            month1, 
            month2
        )
        
        # Display overall KPI comparison - Show only changes, not combined totals
        kpis = comparison_data['overall_kpis']
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            st.metric(
                "💰 Warranty Sales Change", 
                f"₹{kpis['warranty_change']:+,.0f}",
                f"{kpis['warranty_change_pct']:+.1f}%",
                delta_color="normal"
            )
        
        with col2:
            st.metric(
                "📦 Warranty Units Change", 
                f"{kpis['warranty_units_change']:+,.0f}",
                f"{kpis['warranty_units_change_pct']:+.1f}%",
                delta_color="normal"
            )
        
        with col3:
            st.metric(
                "📊 Count Conversion Change", 
                f"{kpis['count_conv_change']:+.2f}%",
                delta_color="normal"
            )
        
        with col4:
            st.metric(
                "📈 Value Conversion Change", 
                f"{kpis['value_conv_change']:+.2f}%",
                delta_color="normal"
            )
        
        with col5:
            st.metric(
                "💵 AHSP Change", 
                f"₹{kpis['ahsp_change']:+.2f}",
                delta_color="normal"
            )
        
        # Store comparison table - Show only changes
        st.markdown(f'#### 🏬 Store Performance Changes')
        store_comp = comparison_data['store_comparison']
        
        # Create display table with only change columns
        display_cols = ['Store']
        # Add change columns only
        display_cols.extend([
            'Value Conv Change', 
            'Count Conv Change', 
            'AHSP Change', 
            'Warranty Sales Change',
            'Warranty Sales Change %',
            'Warranty Units Change',
            'Warranty Units Change %'
        ])
        
        # Check if all required columns exist
        missing_cols = [col for col in display_cols if col not in store_comp.columns]
        if missing_cols:
            st.warning(f"Some comparison data is missing: {missing_cols}")
            # Use only available columns
            display_cols = [col for col in display_cols if col in store_comp.columns]
        
        comparison_display = store_comp[display_cols].copy()
        
        # Rename columns for better display
        column_mapping = {
            'Value Conv Change': 'Value Conv Change (%)',
            'Count Conv Change': 'Count Conv Change (%)',
            'AHSP Change': 'AHSP Change (₹)',
            'Warranty Sales Change': 'Warranty Sales Change (₹)',
            'Warranty Sales Change %': 'Warranty Sales Change (%)',
            'Warranty Units Change': 'Warranty Units Change',
            'Warranty Units Change %': 'Warranty Units Change (%)'
        }
        
        comparison_display = comparison_display.rename(columns=column_mapping)
        
        # Format the display
        def format_comparison_row(row):
            formats = {}
            for col in comparison_display.columns:
                if 'Value Conv Change (%)' in col:
                    formats[col] = '{:+.2f}%'
                elif 'Count Conv Change (%)' in col:
                    formats[col] = '{:+.2f}%'
                elif 'AHSP Change (₹)' in col:
                    formats[col] = '₹{:+.2f}'
                elif 'Warranty Sales Change (₹)' in col:
                    formats[col] = '₹{:+,.0f}'
                elif 'Warranty Sales Change (%)' in col:
                    formats[col] = '{:+.2f}%'
                elif 'Warranty Units Change' in col:
                    formats[col] = '{:+,.0f}'
                elif 'Warranty Units Change (%)' in col:
                    formats[col] = '{:+.2f}%'
            return formats
        
        format_dict = format_comparison_row(comparison_display.iloc[0] if not comparison_display.empty else {})
        
        styled_comparison = comparison_display.style.format(format_dict)
        
        st.dataframe(styled_comparison, use_container_width=True)
        
        # Staff Performance Comparison
        st.markdown(f'#### 👨‍💼 Staff Performance Changes')
        staff_comp = comparison_data['staff_comparison']
        
        # Create display table with only change columns
        staff_display_cols = ['Staff Name', 'Store']
        staff_display_cols.extend([
            'Value Conv Change', 
            'Count Conv Change', 
            'AHSP Change', 
            'Warranty Sales Change',
            'Warranty Units Change'
        ])
        
        staff_comparison_display = staff_comp[staff_display_cols].copy()
        
        # Rename columns for better display
        staff_column_mapping = {
            'Value Conv Change': 'Value Conv Change (%)',
            'Count Conv Change': 'Count Conv Change (%)',
            'AHSP Change': 'AHSP Change (₹)',
            'Warranty Sales Change': 'Warranty Sales Change (₹)',
            'Warranty Units Change': 'Warranty Units Change'
        }
        
        staff_comparison_display = staff_comparison_display.rename(columns=staff_column_mapping)
        
        # Format the display
        def format_staff_comparison_row(row):
            formats = {}
            for col in staff_comparison_display.columns:
                if 'Value Conv Change (%)' in col:
                    formats[col] = '{:+.2f}%'
                elif 'Count Conv Change (%)' in col:
                    formats[col] = '{:+.2f}%'
                elif 'AHSP Change (₹)' in col:
                    formats[col] = '₹{:+.2f}'
                elif 'Warranty Sales Change (₹)' in col:
                    formats[col] = '₹{:+,.0f}'
                elif 'Warranty Units Change' in col:
                    formats[col] = '{:+,.0f}'
            return formats
        
        staff_format_dict = format_staff_comparison_row(staff_comparison_display.iloc[0] if not staff_comparison_display.empty else {})
        
        styled_staff_comparison = staff_comparison_display.style.format(staff_format_dict)
        
        st.dataframe(styled_staff_comparison, use_container_width=True)
        
        # RBM Performance Comparison
        st.markdown(f'#### 👥 RBM Performance Changes')
        rbm_comp = comparison_data['rbm_comparison']
        
        # Create display table with only change columns
        rbm_display_cols = ['RBM']
        rbm_display_cols.extend([
            'Value Conv Change', 
            'Count Conv Change', 
            'AHSP Change', 
            'Warranty Sales Change',
            'Warranty Units Change'
        ])
        
        rbm_comparison_display = rbm_comp[rbm_display_cols].copy()
        
        # Rename columns for better display
        rbm_column_mapping = {
            'Value Conv Change': 'Value Conv Change (%)',
            'Count Conv Change': 'Count Conv Change (%)',
            'AHSP Change': 'AHSP Change (₹)',
            'Warranty Sales Change': 'Warranty Sales Change (₹)',
            'Warranty Units Change': 'Warranty Units Change'
        }
        
        rbm_comparison_display = rbm_comparison_display.rename(columns=rbm_column_mapping)
        
        # Format the display
        def format_rbm_comparison_row(row):
            formats = {}
            for col in rbm_comparison_display.columns:
                if 'Value Conv Change (%)' in col:
                    formats[col] = '{:+.2f}%'
                elif 'Count Conv Change (%)' in col:
                    formats[col] = '{:+.2f}%'
                elif 'AHSP Change (₹)' in col:
                    formats[col] = '₹{:+.2f}'
                elif 'Warranty Sales Change (₹)' in col:
                    formats[col] = '₹{:+,.0f}'
                elif 'Warranty Units Change' in col:
                    formats[col] = '{:+,.0f}'
            return formats
        
        rbm_format_dict = format_rbm_comparison_row(rbm_comparison_display.iloc[0] if not rbm_comparison_display.empty else {})
        
        styled_rbm_comparison = rbm_comparison_display.style.format(rbm_format_dict)
        
        st.dataframe(styled_rbm_comparison, use_container_width=True)
        
        # Product Category Performance Comparison
        st.markdown(f'#### 📦 Product Category Performance Changes')
        category_comp = comparison_data['category_comparison']
        
        # Create display table with only change columns
        category_display_cols = ['Item Category']
        category_display_cols.extend([
            'Value Conv Change', 
            'Count Conv Change', 
            'AHSP Change', 
            'Warranty Sales Change',
            'Warranty Units Change'
        ])
        
        category_comparison_display = category_comp[category_display_cols].copy()
        
        # Rename columns for better display
        category_column_mapping = {
            'Value Conv Change': 'Value Conv Change (%)',
            'Count Conv Change': 'Count Conv Change (%)',
            'AHSP Change': 'AHSP Change (₹)',
            'Warranty Sales Change': 'Warranty Sales Change (₹)',
            'Warranty Units Change': 'Warranty Units Change'
        }
        
        category_comparison_display = category_comparison_display.rename(columns=category_column_mapping)
        
        # Format the display
        def format_category_comparison_row(row):
            formats = {}
            for col in category_comparison_display.columns:
                if 'Value Conv Change (%)' in col:
                    formats[col] = '{:+.2f}%'
                elif 'Count Conv Change (%)' in col:
                    formats[col] = '{:+.2f}%'
                elif 'AHSP Change (₹)' in col:
                    formats[col] = '₹{:+.2f}'
                elif 'Warranty Sales Change (₹)' in col:
                    formats[col] = '₹{:+,.0f}'
                elif 'Warranty Units Change' in col:
                    formats[col] = '{:+,.0f}'
            return formats
        
        category_format_dict = format_category_comparison_row(category_comparison_display.iloc[0] if not category_comparison_display.empty else {})
        
        styled_category_comparison = category_comparison_display.style.format(category_format_dict)
        
        st.dataframe(styled_category_comparison, use_container_width=True)
        
        # Download option for comparison data
        st.markdown("---")
        st.markdown("#### 📥 Download Comparison Data")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            store_comparison_download = store_comp[['Store'] + [col for col in store_comp.columns if 'Change' in col]]
            st.download_button(
                label="Download Store Comparison",
                data=to_excel(store_comparison_download, 'Store Comparison'),
                file_name=f"store_comparison_{month1}_vs_{month2}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col2:
            staff_comparison_download = staff_comp[['Staff Name', 'Store'] + [col for col in staff_comp.columns if 'Change' in col]]
            st.download_button(
                label="Download Staff Comparison",
                data=to_excel(staff_comparison_download, 'Staff Comparison'),
                file_name=f"staff_comparison_{month1}_vs_{month2}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col3:
            rbm_comparison_download = rbm_comp[['RBM'] + [col for col in rbm_comp.columns if 'Change' in col]]
            st.download_button(
                label="Download RBM Comparison",
                data=to_excel(rbm_comparison_download, 'RBM Comparison'),
                file_name=f"rbm_comparison_{month1}_vs_{month2}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col4:
            category_comparison_download = category_comp[['Item Category'] + [col for col in category_comp.columns if 'Change' in col]]
            st.download_button(
                label="Download Category Comparison",
                data=to_excel(category_comparison_download, 'Category Comparison'),
                file_name=f"category_comparison_{month1}_vs_{month2}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # Add some space between comparisons
        st.markdown("---")

    # MAIN DASHBOARD SECTION (Individual, Combined, or All data view)
    # Show this section for all cases except when exactly 2 months are selected for comparison
    if not (len(st.session_state.selected_sheets) == 2 and "All" not in st.session_state.selected_sheets):
        if "All" in st.session_state.selected_sheets:
            # All months combined view
            st.markdown(f'<h2 class="subheader">📈 All Months Combined Performance Overview</h2>', unsafe_allow_html=True)
            display_df = filtered_df
        elif len(st.session_state.selected_sheets) == 1:
            # Single month view
            current_month = st.session_state.selected_sheets[0]
            st.markdown(f'<h2 class="subheader">📈 {current_month} Performance Overview</h2>', unsafe_allow_html=True)
            
            # Use individual month data for single month view
            single_month_df = individual_data[current_month]
            
            # Apply filters to single month data
            single_month_filtered = apply_comparison_filters(
                single_month_df, 
                st.session_state.comparison_filters, 
                category_column,
                replacement_filter,
                speaker_filter
            )
            
            # Use filtered single month data for KPIs
            display_df = single_month_filtered
        else:
            # Multiple months view - show combined data
            st.markdown(f'<h2 class="subheader">📈 Combined Performance Overview ({len(st.session_state.selected_sheets)} Months)</h2>', unsafe_allow_html=True)
            display_df = filtered_df

        # KPI metrics
        st.markdown('<h3 style="color: #1e293b; font-weight: 600; margin: 25px 0 15px 0;">🎯 Key Performance Indicators</h3>', unsafe_allow_html=True)
        col1, col2, col3, col4, col5 = st.columns(5)
        total_warranty = display_df['WarrantyPrice'].sum()
        total_units = display_df['TotalCount'].sum()
        total_warranty_units = display_df['WarrantyCount'].sum()  # CORRECTED: Use WarrantyCount for warranty units
        total_sales = display_df['TotalSoldPrice'].sum()
        count_conversion = (total_warranty_units / total_units * 100) if total_units > 0 else 0  # CORRECTED: WarrantyCount/TotalCount
        value_conversion = (total_warranty / total_sales * 100) if total_sales > 0 else 0
        ahsp = (total_warranty / total_warranty_units) if total_warranty_units > 0 else 0  # CORRECTED: Use WarrantyCount for AHSP

        with col1:
            st.metric("💰 Warranty Sales", f"₹{total_warranty:,.0f}")
        with col2:
            st.metric("📦 Warranty Units", f"{total_warranty_units:,.0f}")  # CORRECTED: Show WarrantyCount
        with col3:
            st.metric("📊 Count Conversion", f"{count_conversion:.2f}%")
        with col4:
            st.metric("📈 Value Conversion", f"{value_conversion:.2f}%")
        with col5:
            st.metric("💵 AHSP", f"₹{ahsp:,.2f}")

        # Store Performance
        if "All" in st.session_state.selected_sheets:
            st.markdown(f'<h3 class="subheader">🏬 Store Performance Analysis - All Months Combined</h3>', unsafe_allow_html=True)
        elif len(st.session_state.selected_sheets) == 1:
            st.markdown(f'<h3 class="subheader">🏬 Store Performance Analysis - {current_month}</h3>', unsafe_allow_html=True)
        else:
            st.markdown(f'<h3 class="subheader">🏬 Store Performance Analysis - Combined View</h3>', unsafe_allow_html=True)

        store_summary = display_df.groupby('Store').agg({
            'TotalSoldPrice': 'sum',
            'WarrantyPrice': 'sum',
            'TotalCount': 'sum',
            'WarrantyCount': 'sum'  # CORRECTED: Include WarrantyCount
        }).reset_index()

        # CORRECTED: Use WarrantyCount for count conversion and warranty units
        store_summary['Count Conv (%)'] = (store_summary['WarrantyCount'] / store_summary['TotalCount'] * 100).round(2)
        store_summary['Value Conv (%)'] = (store_summary['WarrantyPrice'] / store_summary['TotalSoldPrice'] * 100).round(2)
        store_summary['AHSP'] = (store_summary['WarrantyPrice'] / store_summary['WarrantyCount']).where(store_summary['WarrantyCount'] > 0, 0).round(2)

        store_summary['Count Conv (%)'] = store_summary['Count Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)
        store_summary['Value Conv (%)'] = store_summary['Value Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)

        store_display = store_summary[['Store', 'WarrantyPrice', 'WarrantyCount', 'Count Conv (%)', 'Value Conv (%)', 'AHSP']].copy()  # CORRECTED: Use WarrantyCount
        store_display.columns = ['Store', 'Warranty Sales (₹)', 'Warranty Units', 'Count Conv (%)', 'Value Conv (%)', 'AHSP (₹)']

        total_sales = store_summary['TotalSoldPrice'].sum()
        total_warranty_sales = store_summary['WarrantyPrice'].sum()
        total_units = store_summary['TotalCount'].sum()
        total_warranty_units = store_summary['WarrantyCount'].sum()  # CORRECTED: Use WarrantyCount
        
        total_count_conv = (total_warranty_units / total_units * 100) if total_units > 0 else 0  # CORRECTED: WarrantyCount/TotalCount
        total_value_conv = (total_warranty_sales / total_sales * 100) if total_sales > 0 else 0
        total_ahsp = (total_warranty_sales / total_warranty_units) if total_warranty_units > 0 else 0  # CORRECTED: Use WarrantyCount

        total_row = pd.DataFrame({
            'Store': ['Total'],
            'Warranty Sales (₹)': [total_warranty_sales],
            'Warranty Units': [total_warranty_units],  # CORRECTED: Use WarrantyCount
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

        # Generate filename based on selected sheets
        if "All" in st.session_state.selected_sheets:
            file_suffix = "all_months_combined"
        elif len(st.session_state.selected_sheets) == 1:
            file_suffix = st.session_state.selected_sheets[0].lower().replace(' ', '_')
        else:
            file_suffix = f"{len(st.session_state.selected_sheets)}_months_combined"

        st.download_button(
            label="📥 Download Store Performance as Excel",
            data=to_excel(final_store_display, 'Store Performance'),
            file_name=f"store_performance_{file_suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Staff Performance Table
        if "All" in st.session_state.selected_sheets:
            st.markdown(f'<h3 class="subheader">👨‍💼 Staff Performance Analysis - All Months Combined</h3>', unsafe_allow_html=True)
        elif len(st.session_state.selected_sheets) == 1:
            st.markdown(f'<h3 class="subheader">👨‍💼 Staff Performance Analysis - {current_month}</h3>', unsafe_allow_html=True)
        else:
            st.markdown(f'<h3 class="subheader">👨‍💼 Staff Performance Analysis - Combined View</h3>', unsafe_allow_html=True)

        staff_summary = display_df.groupby(['Staff Name', 'Store']).agg({
            'TotalSoldPrice': 'sum',
            'WarrantyPrice': 'sum',
            'TotalCount': 'sum',
            'WarrantyCount': 'sum'  # CORRECTED: Include WarrantyCount
        }).reset_index()

        # CORRECTED: Use WarrantyCount for count conversion and warranty units
        staff_summary['Count Conv (%)'] = (staff_summary['WarrantyCount'] / staff_summary['TotalCount'] * 100).round(2)
        staff_summary['Value Conv (%)'] = (staff_summary['WarrantyPrice'] / staff_summary['TotalSoldPrice'] * 100).round(2)
        staff_summary['AHSP'] = (staff_summary['WarrantyPrice'] / staff_summary['WarrantyCount']).where(staff_summary['WarrantyCount'] > 0, 0).round(2)

        staff_summary['Count Conv (%)'] = staff_summary['Count Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)
        staff_summary['Value Conv (%)'] = staff_summary['Value Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)

        staff_display = staff_summary[['Staff Name', 'Store', 'Value Conv (%)', 'Count Conv (%)', 'WarrantyPrice', 'WarrantyCount', 'AHSP']].copy()  # CORRECTED: Use WarrantyCount
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
        
        # Add total row to staff performance table
        total_staff_row = pd.DataFrame({
            'Staff Name': ['Total'],
            'Store': [''],
            'Value Conv (%)': [round((staff_summary['WarrantyPrice'].sum() / staff_summary['TotalSoldPrice'].sum() * 100) if staff_summary['TotalSoldPrice'].sum() > 0 else 0, 2)],
            'Count Conv (%)': [round((staff_summary['WarrantyCount'].sum() / staff_summary['TotalCount'].sum() * 100) if staff_summary['TotalCount'].sum() > 0 else 0, 2)],
            'Warranty Sales (₹)': [staff_summary['WarrantyPrice'].sum()],
            'Warranty Units': [staff_summary['WarrantyCount'].sum()],
            'AHSP (₹)': [round((staff_summary['WarrantyPrice'].sum() / staff_summary['WarrantyCount'].sum()) if staff_summary['WarrantyCount'].sum() > 0 else 0, 2)]
        })
        
        staff_display_with_total = pd.concat([staff_display, total_staff_row], ignore_index=True)

        st.dataframe(staff_display_with_total.style.format({
            'Value Conv (%)': '{:.2f}%',
            'Count Conv (%)': '{:.2f}%',
            'Warranty Sales (₹)': '₹{:.0f}',
            'Warranty Units': '{:.0f}',
            'AHSP (₹)': '₹{:.2f}'
        }), use_container_width=True)

        st.download_button(
            label="📥 Download Staff Performance as Excel",
            data=to_excel(staff_display_with_total, 'Staff Performance'),
            file_name=f"staff_performance_{file_suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # RBM Performance
        if "All" in st.session_state.selected_sheets:
            st.markdown(f'<h3 class="subheader">👥 RBM Performance Analysis - All Months Combined</h3>', unsafe_allow_html=True)
        elif len(st.session_state.selected_sheets) == 1:
            st.markdown(f'<h3 class="subheader">👥 RBM Performance Analysis - {current_month}</h3>', unsafe_allow_html=True)
        else:
            st.markdown(f'<h3 class="subheader">👥 RBM Performance Analysis - Combined View</h3>', unsafe_allow_html=True)

        rbm_summary = display_df.groupby('RBM').agg({
            'TotalSoldPrice': 'sum',
            'WarrantyPrice': 'sum',
            'TotalCount': 'sum',
            'WarrantyCount': 'sum'  # CORRECTED: Include WarrantyCount
        }).reset_index()

        # CORRECTED: Use WarrantyCount for count conversion and warranty units
        rbm_summary['Count Conv (%)'] = (rbm_summary['WarrantyCount'] / rbm_summary['TotalCount'] * 100).round(2)
        rbm_summary['Value Conv (%)'] = (rbm_summary['WarrantyPrice'] / rbm_summary['TotalSoldPrice'] * 100).round(2)
        rbm_summary['AHSP'] = (rbm_summary['WarrantyPrice'] / rbm_summary['WarrantyCount']).where(rbm_summary['WarrantyCount'] > 0, 0).round(2)

        rbm_summary['Count Conv (%)'] = rbm_summary['Count Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)
        rbm_summary['Value Conv (%)'] = rbm_summary['Value Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)

        rbm_display = rbm_summary[['RBM', 'Count Conv (%)', 'Value Conv (%)', 'AHSP', 'WarrantyPrice', 'WarrantyCount']]  # CORRECTED: Use WarrantyCount
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
        
        # Add total row to RBM performance table
        total_rbm_row = pd.DataFrame({
            'RBM': ['Total'],
            'Count Conv (%)': [round((rbm_summary['WarrantyCount'].sum() / rbm_summary['TotalCount'].sum() * 100) if rbm_summary['TotalCount'].sum() > 0 else 0, 2)],
            'Value Conv (%)': [round((rbm_summary['WarrantyPrice'].sum() / rbm_summary['TotalSoldPrice'].sum() * 100) if rbm_summary['TotalSoldPrice'].sum() > 0 else 0, 2)],
            'AHSP (₹)': [round((rbm_summary['WarrantyPrice'].sum() / rbm_summary['WarrantyCount'].sum()) if rbm_summary['WarrantyCount'].sum() > 0 else 0, 2)],
            'Warranty Sales (₹)': [rbm_summary['WarrantyPrice'].sum()],
            'Warranty Units': [rbm_summary['WarrantyCount'].sum()]
        })
        
        rbm_display_with_total = pd.concat([rbm_display, total_rbm_row], ignore_index=True)

        st.dataframe(rbm_display_with_total.style.format({
            'Count Conv (%)': '{:.2f}%',
            'Value Conv (%)': '{:.2f}%',
            'AHSP (₹)': '₹{:.2f}',
            'Warranty Sales (₹)': '₹{:.0f}',
            'Warranty Units': '{:.0f}'
        }), use_container_width=True)

        st.download_button(
            label="📥 Download RBM Performance as Excel",
            data=to_excel(rbm_display_with_total, 'RBM Performance'),
            file_name=f"rbm_performance_{file_suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Product Category Performance with Small Appliance Grouping
        if "All" in st.session_state.selected_sheets:
            st.markdown(f'<h3 class="subheader">📦 Product Category Performance - All Months Combined</h3>', unsafe_allow_html=True)
        elif len(st.session_state.selected_sheets) == 1:
            st.markdown(f'<h3 class="subheader">📦 Product Category Performance - {current_month}</h3>', unsafe_allow_html=True)
        else:
            st.markdown(f'<h3 class="subheader">📦 Product Category Performance - Combined View</h3>', unsafe_allow_html=True)

        # Define major appliances that should stay as separate rows
        major_appliances = ['AC', 'TV', 'WASHING MACHINE', 'REFRIGERATOR', 'MICROWAVE OVEN', 'DISH WASHER', 'DRYER']

        # Create a new column for the grouped category
        display_df['Grouped Category'] = display_df['Item Category'].apply(
            lambda x: 'SMALL APPLIANCE' if x not in major_appliances else x
        )

        category_summary = display_df.groupby('Grouped Category').agg({
            'TotalSoldPrice': 'sum',
            'WarrantyPrice': 'sum',
            'TotalCount': 'sum',
            'WarrantyCount': 'sum'  # CORRECTED: Include WarrantyCount
        }).reset_index()

        # CORRECTED: Use WarrantyCount for count conversion and warranty units
        category_summary['Count Conv (%)'] = (category_summary['WarrantyCount'] / category_summary['TotalCount'] * 100).round(2)
        category_summary['Value Conv (%)'] = (category_summary['WarrantyPrice'] / category_summary['TotalSoldPrice'] * 100).round(2)
        category_summary['AHSP'] = (category_summary['WarrantyPrice'] / category_summary['WarrantyCount']).where(category_summary['WarrantyCount'] > 0, 0).round(2)

        category_summary['Count Conv (%)'] = category_summary['Count Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)
        category_summary['Value Conv (%)'] = category_summary['Value Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)

        if not category_summary.empty:
            category_display = category_summary[['Grouped Category', 'Count Conv (%)', 'Value Conv (%)', 'AHSP', 'WarrantyPrice', 'WarrantyCount']]  # CORRECTED: Use WarrantyCount
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
            
            # Add total row to category performance table
            total_category_row = pd.DataFrame({
                'Product Category': ['Total'],
                'Count Conv (%)': [round((category_summary['WarrantyCount'].sum() / category_summary['TotalCount'].sum() * 100) if category_summary['TotalCount'].sum() > 0 else 0, 2)],
                'Value Conv (%)': [round((category_summary['WarrantyPrice'].sum() / category_summary['TotalSoldPrice'].sum() * 100) if category_summary['TotalSoldPrice'].sum() > 0 else 0, 2)],
                'AHSP (₹)': [round((category_summary['WarrantyPrice'].sum() / category_summary['WarrantyCount'].sum()) if category_summary['WarrantyCount'].sum() > 0 else 0, 2)],
                'Warranty Sales (₹)': [category_summary['WarrantyPrice'].sum()],
                'Warranty Units': [category_summary['WarrantyCount'].sum()]
            })
            
            category_display_with_total = pd.concat([category_display, total_category_row], ignore_index=True)

            st.dataframe(category_display_with_total.style.format({
                'Count Conv (%)': '{:.2f}%',
                'Value Conv (%)': '{:.2f}%',
                'AHSP (₹)': '₹{:.2f}',
                'Warranty Sales (₹)': '₹{:.0f}',
                'Warranty Units': '{:.0f}'
            }), use_container_width=True)

            st.download_button(
                label="📥 Download Product Category Performance as Excel",
                data=to_excel(category_display_with_total, 'Product Category Performance'),
                file_name=f"product_category_performance_{file_suffix}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ No category data available with current filters.")

        # Item Category Performance (Full Product Breakdown)
        if "All" in st.session_state.selected_sheets:
            st.markdown(f'<h3 class="subheader">📋 Item Category Performance - Full Product Breakdown - All Months Combined</h3>', unsafe_allow_html=True)
        elif len(st.session_state.selected_sheets) == 1:
            st.markdown(f'<h3 class="subheader">📋 Item Category Performance - Full Product Breakdown - {current_month}</h3>', unsafe_allow_html=True)
        else:
            st.markdown(f'<h3 class="subheader">📋 Item Category Performance - Full Product Breakdown - Combined View</h3>', unsafe_allow_html=True)

        # Full item category performance without grouping
        item_category_summary = display_df.groupby('Item Category').agg({
            'TotalSoldPrice': 'sum',
            'WarrantyPrice': 'sum',
            'TotalCount': 'sum',
            'WarrantyCount': 'sum'  # CORRECTED: Include WarrantyCount
        }).reset_index()

        # CORRECTED: Use WarrantyCount for count conversion and warranty units
        item_category_summary['Count Conv (%)'] = (item_category_summary['WarrantyCount'] / item_category_summary['TotalCount'] * 100).round(2)
        item_category_summary['Value Conv (%)'] = (item_category_summary['WarrantyPrice'] / item_category_summary['TotalSoldPrice'] * 100).round(2)
        item_category_summary['AHSP'] = (item_category_summary['WarrantyPrice'] / item_category_summary['WarrantyCount']).where(item_category_summary['WarrantyCount'] > 0, 0).round(2)

        item_category_summary['Count Conv (%)'] = item_category_summary['Count Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)
        item_category_summary['Value Conv (%)'] = item_category_summary['Value Conv (%)'].replace([float('inf'), -float('inf')], 0).fillna(0)

        if not item_category_summary.empty:
            item_category_display = item_category_summary[['Item Category', 'Count Conv (%)', 'Value Conv (%)', 'AHSP', 'WarrantyPrice', 'WarrantyCount']]  # CORRECTED: Use WarrantyCount
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
            
            # Add total row to item category performance table
            total_item_category_row = pd.DataFrame({
                'Item Category': ['Total'],
                'Count Conv (%)': [round((item_category_summary['WarrantyCount'].sum() / item_category_summary['TotalCount'].sum() * 100) if item_category_summary['TotalCount'].sum() > 0 else 0, 2)],
                'Value Conv (%)': [round((item_category_summary['WarrantyPrice'].sum() / item_category_summary['TotalSoldPrice'].sum() * 100) if item_category_summary['TotalSoldPrice'].sum() > 0 else 0, 2)],
                'AHSP (₹)': [round((item_category_summary['WarrantyPrice'].sum() / item_category_summary['WarrantyCount'].sum()) if item_category_summary['WarrantyCount'].sum() > 0 else 0, 2)],
                'Warranty Sales (₹)': [item_category_summary['WarrantyPrice'].sum()],
                'Warranty Units': [item_category_summary['WarrantyCount'].sum()]
            })
            
            item_category_display_with_total = pd.concat([item_category_display, total_item_category_row], ignore_index=True)

            st.dataframe(item_category_display_with_total.style.format({
                'Count Conv (%)': '{:.2f}%',
                'Value Conv (%)': '{:.2f}%',
                'AHSP (₹)': '₹{:.2f}',
                'Warranty Sales (₹)': '₹{:.0f}',
                'Warranty Units': '{:.0f}'
            }), use_container_width=True)

            st.download_button(
                label="📥 Download Item Category Performance as Excel",
                data=to_excel(item_category_display_with_total, 'Item Category Performance'),
                file_name=f"item_category_performance_{file_suffix}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("ℹ️ No item category data available with current filters.")
    
    # NEW: RBM-WISE MONTHLY SUMMARY TABLE
    st.markdown(f'<h3 class="subheader">👥 RBM-wise Monthly Warranty Sales Summary</h3>', unsafe_allow_html=True)
    
    # Create the RBM-wise monthly summary table with filters applied
    rbm_summary_table = create_rbm_monthly_summary(
        individual_data, 
        st.session_state.comparison_filters, 
        category_column,
        replacement_filter,
        speaker_filter
    )
    
    if not rbm_summary_table.empty:
        # Display the table
        st.dataframe(rbm_summary_table, use_container_width=True)
        
        # Download button for the RBM summary
        st.download_button(
            label="📥 Download RBM Monthly Summary as Excel",
            data=to_excel(rbm_summary_table, 'RBM Monthly Summary'),
            file_name="rbm_monthly_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("No RBM summary data available with current filters.")

    # NEW: RBM-WISE MONTHLY VALUE CONVERSION SUMMARY TABLE
    st.markdown(f'<h3 class="subheader">👥 RBM-wise Monthly Value Conversion Summary</h3>', unsafe_allow_html=True)
    
    # Create the RBM-wise monthly value conversion summary table with filters applied
    rbm_value_conversion_table = create_rbm_monthly_value_conversion_summary(
        individual_data, 
        st.session_state.comparison_filters, 
        category_column,
        replacement_filter,
        speaker_filter
    )
    
    if not rbm_value_conversion_table.empty:
        # Display the table
        st.dataframe(rbm_value_conversion_table, use_container_width=True)
        
        # Download button for the RBM value conversion summary
        st.download_button(
            label="📥 Download RBM Value Conversion Summary as Excel",
            data=to_excel(rbm_value_conversion_table, 'RBM Value Conv Summary'),
            file_name="rbm_value_conversion_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("No RBM value conversion summary data available with current filters.")

    # MOVED: PRODUCT-WISE MONTHLY SUMMARY TABLE - Now at the bottom before the trend chart
    st.markdown(f'<h3 class="subheader">📊 Product-wise Monthly Warranty Sales Summary</h3>', unsafe_allow_html=True)
    
    # Create the product-wise monthly summary table with filters applied
    product_summary_table = create_product_monthly_summary(
        individual_data, 
        st.session_state.comparison_filters, 
        category_column,
        replacement_filter,
        speaker_filter
    )
    
    if not product_summary_table.empty:
        # Display the table
        st.dataframe(product_summary_table, use_container_width=True)
        
        # Download button for the product summary
        st.download_button(
            label="📥 Download Product Monthly Summary as Excel",
            data=to_excel(product_summary_table, 'Product Monthly Summary'),
            file_name="product_monthly_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("No product summary data available with current filters.")

    # NEW: PRODUCT-WISE MONTHLY VALUE CONVERSION SUMMARY TABLE
    st.markdown(f'<h3 class="subheader">📊 Product-wise Monthly Value Conversion Summary</h3>', unsafe_allow_html=True)
    
    # Create the product-wise monthly value conversion summary table with filters applied
    product_value_conversion_table = create_product_monthly_value_conversion_summary(
        individual_data, 
        st.session_state.comparison_filters, 
        category_column,
        replacement_filter,
        speaker_filter
    )
    
    if not product_value_conversion_table.empty:
        # Display the table
        st.dataframe(product_value_conversion_table, use_container_width=True)
        
        # Download button for the product value conversion summary
        st.download_button(
            label="📥 Download Product Value Conversion Summary as Excel",
            data=to_excel(product_value_conversion_table, 'Product Value Conv Summary'),
            file_name="product_value_conversion_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("No product value conversion summary data available with current filters.")

    # MONTHLY TREND CHART SECTION - Show for all scenarios (now at the very bottom)
    st.markdown(f'<h3 class="subheader">📈 Monthly Warranty Sales Trend</h3>', unsafe_allow_html=True)
    
    # Create and display the monthly trend chart
    trend_chart = create_monthly_trend_chart(st.session_state.individual_month_data)
    if trend_chart:
        st.plotly_chart(trend_chart, use_container_width=True)
    else:
        st.info("No trend data available to display.")

else:
    if not st.session_state.data_loaded and st.session_state.selected_sheets:
        show_loading_animation(
            "Initializing Dashboard", 
            "Setting up your warranty conversion analysis dashboard..."
        )
    elif not st.session_state.selected_sheets:
        st.warning("⚠️ Please select at least one month from the sidebar to view data.")
    else:
        st.error("❌ Failed to load data from Google Sheets. Please check the Apps Script URL, Spreadsheet ID, or network connection.")
