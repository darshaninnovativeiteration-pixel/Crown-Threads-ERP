import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import pandas as pd

# ૧. Google Sheets કનેક્શન
@st.cache_resource
def connect_to_google():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name("keys.json", scope)
        client = gspread.authorize(creds)
        return client.open("Crown_Threads_DB")
    except Exception as e:
        st.error(f"શીટ કનેક્ટ કરવામાં ભૂલ આવી: {e}")
        return None

spreadsheet = connect_to_google()

st.set_page_config(page_title="Crown Threads ERP", layout="wide")
st.sidebar.title("💎 Crown Threads ERP")
page = st.sidebar.radio("મેનુ પસંદ કરો", ["📊 Dashboard & Audit", "📤 Dispatch Scan"])

# --- PAGE 1: DASHBOARD ---
if page == "📊 Dashboard & Audit":
    st.title("📊 Live Audit Dashboard")
    if spreadsheet:
        master_sheet = spreadsheet.worksheet("Ajio_Master_List")
        dispatch_sheet = spreadsheet.worksheet("Sheet1")
        
        with st.expander("📁 સવારે નવી ફાઈલો અહીં અપલોડ કરો"):
            uploaded_files = st.file_uploader("Ajio Excel પસંદ કરો", type=['xlsx'], accept_multiple_files=True)
            if uploaded_files and st.button("ક્લાઉડમાં સેવ કરો"):
                all_new_rows = []
                today_date = datetime.now().strftime("%Y-%m-%d")
                for uploaded_file in uploaded_files:
                    df_raw = pd.read_excel(uploaded_file, sheet_name='Order Details', header=None)
                    header_row = next((i for i, row in df_raw.iterrows() if 'Customer Order Id' in row.values), None)
                    ajio_df = pd.read_excel(uploaded_file, sheet_name='Order Details', skiprows=header_row)
                    sku_col = 'Seller SKU ID' if 'Seller SKU ID' in ajio_df.columns else 'Seller SKU'
                    for _, row in ajio_df.iterrows():
                        if pd.notna(row['Customer Order Id']):
                            order_id = str(row['Customer Order Id']).split('.')[0]
                            sku = str(row[sku_col]).strip()
                            if sku and sku != "-" and sku != "nan":
                                all_new_rows.append([today_date, order_id, 'Pending', sku])
                if all_new_rows:
                    master_sheet.append_rows(all_new_rows)
                    st.success("ડેટા સેવ થયો!")
                    st.rerun()

        master_df = pd.DataFrame(master_sheet.get_all_records())
        user_df = pd.DataFrame(dispatch_sheet.get_all_records())
        if not master_df.empty:
            def clean_sku_join(x):
                skus = [str(i).strip() for i in x.unique() if str(i).strip() not in ["-", "nan", ""]]
                return ", ".join(skus)
            grouped_master = master_df.groupby('Customer Order Id').agg({'Upload_Date': 'first', 'SKU': clean_sku_join}).reset_index()
            status_dict = dict(zip(user_df['AWB'].astype(str), user_df['Status'])) if not user_df.empty else {}
            grouped_master['Live_Status'] = grouped_master['Customer Order Id'].astype(str).apply(lambda x: status_dict.get(x, '❌ Pending'))

            c1, c2, c3 = st.columns(3)
            c1.metric("📦 બાકી", len(grouped_master[grouped_master['Live_Status'] == '❌ Pending']))
            c2.metric("✅ ડિસ્પેચ", len(grouped_master[grouped_master['Live_Status'] == 'Dispatched']))
            c3.metric("🚩 સોફ્ટ ડેટા", len(grouped_master[grouped_master['Live_Status'] == 'Soft Data']))
            
            t1, t2, t3 = st.tabs(["❌ બાકી", "✅ ડિસ્પેચ", "🚩 સોફ્ટ ડેટા"])
            with t1: st.dataframe(grouped_master[grouped_master['Live_Status'] == '❌ Pending'].reset_index(drop=True), use_container_width=True)
            with t2: st.dataframe(grouped_master[grouped_master['Live_Status'] == 'Dispatched'].reset_index(drop=True), use_container_width=True)
            with t3: st.dataframe(grouped_master[grouped_master['Live_Status'] == 'Soft Data'].reset_index(drop=True), use_container_width=True)

# --- PAGE 2: DISPATCH SCAN ---
elif page == "📤 Dispatch Scan":
    st.title("📤 Dispatch Scanning")
    
    master_sheet = spreadsheet.worksheet("Ajio_Master_List")
    sheet1 = spreadsheet.worksheet("Sheet1")
    
    master_data = pd.DataFrame(master_sheet.get_all_records())
    user_data = pd.DataFrame(sheet1.get_all_records())
    
    valid_order_ids = set(master_data['Customer Order Id'].astype(str).unique()) if not master_data.empty else set()
    already_scanned_ids = set(user_data['AWB'].astype(str).unique()) if not user_data.empty else set()

    def process_scan():
        scanned_val = st.session_state.barcode_input.strip()
        if scanned_val:
            # ૧. ડુપ્લીકેટ ચેક
            if scanned_val in already_scanned_ids:
                st.error(f"🚨 RED ALERT: ઓર્ડર {scanned_val} પહેલાથી સ્કેન થઈ ગયો છે!")
            # ૨. લિસ્ટમાં છે કે નહીં તેનો ચેક
            elif scanned_val not in valid_order_ids:
                st.error(f"❌ ઓર્ડર ID {scanned_val} લિસ્ટમાં નથી!")
            # ૩. બધું બરાબર હોય તો સેવ કરો
            else:
                order_items = master_data[master_data['Customer Order Id'].astype(str) == scanned_val]
                skus = ", ".join([str(s).strip() for s in order_items['SKU'].unique() if str(s).strip() not in ["-", "nan", ""]])
                sheet1.append_row([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), scanned_val, "PC_Scanner", "Dispatched", skus])
                st.toast(f"✅ {scanned_val} Dispatched!")
        
        # ઇનપુટ બોક્સ ઓટોમેટિક ક્લીયર કરો
        st.session_state.barcode_input = ""

    st.text_input("Order ID સ્કેન કરો", key="barcode_input", on_change=process_scan)

    st.divider()
    
    with st.container():
        st.subheader("🚩 Soft Data Scan Section")
        soft_input = st.text_input("સોફ્ટ ડેટા માટે અહીં સ્કેન કરો", key="soft_input_box")
        if soft_input:
            soft_val = soft_input.strip()
            if soft_val in already_scanned_ids:
                st.warning(f"આ આઈડી {soft_val} પહેલાથી કોઈ સ્ટેટસમાં છે.")
            elif soft_val in valid_order_ids:
                if st.button(f"Confirm Soft Data for {soft_val}", use_container_width=True):
                    order_items = master_data[master_data['Customer Order Id'].astype(str) == soft_val]
                    skus = ", ".join([str(s).strip() for s in order_items['SKU'].unique() if str(s).strip() not in ["-", "nan", ""]])
                    sheet1.append_row([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), soft_val, "PC_Scanner", "Soft Data", skus])
                    st.success(f"{soft_val} સોફ્ટ ડેટામાં સેવ થયો!")
                    st.rerun()
            else:
                st.error("આ આઈડી માસ્ટર લિસ્ટમાં નથી.")