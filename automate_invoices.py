from fpdf import FPDF
import tempfile
import pandas as pd
from num2words import num2words
import warnings
import datetime

# Suppress warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Config
EXCEL_FILE = 'Baggage invoicing.xlsm'
OUTPUT_BASE_DIR = 'Invoices_Output_Final'
MEDIA_DIR = os.path.join(os.path.dirname(__file__), 'media')
UPLOAD_MEDIA_DIR = os.path.join(tempfile.gettempdir(), 'fly91_media')

LOGO_PATH = os.path.join(MEDIA_DIR, 'image1.png')
# Seal and Sign: Check UPLOAD_MEDIA_DIR (/tmp) first, then repo defaults
SEAL_PATH = os.path.join(UPLOAD_MEDIA_DIR, 'image2.png')
if not os.path.exists(SEAL_PATH): SEAL_PATH = os.path.join(MEDIA_DIR, 'image2.png')

SIGN_PATH = os.path.join(UPLOAD_MEDIA_DIR, 'image3.png')
if not os.path.exists(SIGN_PATH): SIGN_PATH = os.path.join(MEDIA_DIR, 'image3.png')

FOOTER_IMAGE_PATH = os.path.join(MEDIA_DIR, 'image4.jpg')

class ProfessionalInvoice(FPDF):
    def __init__(self, *args, **kwargs):
        # Orientation: 'L' for Landscape, Format: 'A4'
        super().__init__(orientation='L', unit='mm', format='A4', **kwargs)
        self.red_color = (237, 28, 36)
        self.grey_line = (200, 200, 200)
        self.supplier_address_str = ""
        self.supplier_gstin_str = ""

    def header(self):
        # Logo Top Left
        if os.path.exists(LOGO_PATH):
            self.image(LOGO_PATH, 10, 8, 45)
        
        # Center Title - Very Top Center
        self.set_y(15)
        self.set_font('helvetica', 'B', 11)
        self.set_text_color(0, 0, 0)
        self.cell(0, 5, 'TAX INVOICE', 0, 1, 'C')
        self.set_font('helvetica', '', 7)
        self.cell(0, 4, '(Original for Recipient)', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        # Image4 in the middle bottom (Landscape width 297mm)
        if os.path.exists(FOOTER_IMAGE_PATH):
            # Center it: (297 - 240) / 2 = 28.5mm
            self.image(FOOTER_IMAGE_PATH, 28.5, 175, 240)
        
        self.set_y(-10)
        self.set_font('helvetica', '', 7)
        self.set_text_color(150, 150, 150)
        # self.cell(0, 5, f'Page {self.page_no()}/{{nb}}', 0, 0, 'C')

def clean_filename(s):
    if not isinstance(s, str): return str(s)
    for c in r'\/:*?"<>|':
        s = s.replace(c, '')
    return s.strip()

def format_num(val):
    try:
        f = float(val)
        if f == 0: return "-"
        # Round to whole number and then format with .00
        return "{:,.2f}".format(float(round(f)))
    except:
        return "-"

def number_to_words_indian(num):
    try:
        val = float(num)
        if val == 0: return "Zero Only"
        return num2words(val, lang='en_IN').title().replace("-", " ").replace(",", "") + " Only"
    except:
        return ""

def generate_kind_pdf(data, output_path, seal_pos=None, sign_pos=None):
    pdf = ProfessionalInvoice()
    pdf.supplier_address_str = data.get('supplier_address', '')
    pdf.supplier_gstin_str = data.get('supplier_gstin', '')
    pdf.add_page()
    
    # --- Section: Row Positioning ---
    line_h = 4.5
    title_fs = 8
    val_fs = 7.5
    
    # --- ROW 1: Invoice No & Company Name ---
    pdf.set_y(32)
    pdf.set_font('helvetica', 'B', title_fs)
    pdf.cell(30, line_h, "Invoice No")
    pdf.set_font('helvetica', '', title_fs)
    pdf.cell(5, line_h, ":")
    pdf.set_font('helvetica', '', val_fs)
    pdf.cell(100, line_h, data['invoice_no'], 0, 0)
    
    # RIGHT: Company Name (Keep Bold)
    pdf.set_x(150)
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(0, line_h, 'JUST UDO AVIATION PVT LTD', 0, 1, 'R')
    
    # --- ROW 2: Invoice Date & Address ---
    pdf.set_font('helvetica', 'B', title_fs)
    pdf.cell(30, line_h, "Invoice Date")
    pdf.set_font('helvetica', '', title_fs)
    pdf.cell(5, line_h, ":")
    pdf.set_font('helvetica', '', val_fs)
    pdf.cell(100, line_h, data['invoice_date'], 0, 0)
    
    # RIGHT: Address wrap (Narrower width to force 3 rows)
    pdf.set_font('helvetica', '', 6.5)
    pdf.set_x(217) # Move X further right to keep aligned but narrow
    pdf.multi_cell(70, 2.5, pdf.supplier_address_str, 0, 'R')
    
    # GAP between Invoice Date and Passenger Name
    pdf.ln(4)

    # --- BLOCK 2: Details & Route ---
    # Row 1
    pdf.set_font('helvetica', 'B', title_fs)
    pdf.cell(30, line_h, "Passenger Name")
    pdf.set_font('helvetica', '', title_fs)
    pdf.cell(5, line_h, ":")
    pdf.set_font('helvetica', '', val_fs)
    pdf.cell(100, line_h, data['passenger_name'], 0, 0)
    
    # Right Col (Route Starts) - Move Left more
    route_x = 155 # From 180
    pdf.set_x(route_x) 
    pdf.set_font('helvetica', 'B', title_fs)
    pdf.cell(35, line_h, "From")
    pdf.set_font('helvetica', '', title_fs)
    pdf.cell(5, line_h, ":")
    pdf.set_font('helvetica', '', val_fs)
    pdf.cell(0, line_h, data['origin'], 0, 1)

    # Row 2
    pdf.set_font('helvetica', 'B', title_fs)
    pdf.cell(30, line_h, "PNR No")
    pdf.set_font('helvetica', '', title_fs)
    pdf.cell(5, line_h, ":")
    pdf.set_font('helvetica', '', val_fs)
    pdf.cell(50, line_h, data['pnr_no'], 0, 0)
    
    # MIDDLE Flight No
    pdf.set_x(100) 
    pdf.set_font('helvetica', 'B', title_fs)
    pdf.cell(20, line_h, "Flight No")
    pdf.set_font('helvetica', '', title_fs)
    pdf.cell(5, line_h, ":")
    pdf.set_font('helvetica', '', val_fs)
    pdf.cell(30, line_h, data['flight_no'], 0, 0)
    
    # Right Col
    pdf.set_x(route_x)
    pdf.set_font('helvetica', 'B', title_fs)
    pdf.cell(35, line_h, "To")
    pdf.set_font('helvetica', '', title_fs)
    pdf.cell(5, line_h, ":")
    pdf.set_font('helvetica', '', val_fs)
    pdf.cell(0, line_h, data['destination'], 0, 1)

    # Row 3 (Place of Supply & GSTIN of Supplier)
    pdf.set_x(route_x)
    pdf.set_font('helvetica', 'B', title_fs)
    pdf.cell(35, line_h, "Place of Supply")
    pdf.set_font('helvetica', '', title_fs)
    pdf.cell(5, line_h, ":")
    pdf.set_font('helvetica', '', val_fs)
    pdf.cell(40, line_h, data['place_of_supply'], 0, 0)
    
    # GSTIN of Supplier (Adjusted X - Moved Left with the rest)
    pdf.set_x(225) # From 250
    pdf.set_font('helvetica', 'B', title_fs)
    pdf.cell(30, line_h, "GSTIN of Supplier")
    pdf.set_font('helvetica', '', title_fs)
    pdf.cell(5, line_h, ":")
    pdf.set_font('helvetica', '', val_fs)
    pdf.cell(0, line_h, f"{pdf.supplier_gstin_str}", 0, 1)
    
    pdf.ln(3)

    # Section 2: Bill To
    pdf.set_font('helvetica', 'B', title_fs)
    pdf.cell(30, line_h, "GSTIN of Customer")
    pdf.set_font('helvetica', '', title_fs)
    pdf.cell(5, line_h, ":")
    pdf.set_font('helvetica', '', val_fs)
    pdf.cell(0, line_h, data['customer_gstin'], 0, 1)

    pdf.set_font('helvetica', 'B', title_fs)
    pdf.cell(30, line_h, "Customer Name")
    pdf.set_font('helvetica', '', title_fs)
    pdf.cell(5, line_h, ":")
    pdf.set_font('helvetica', '', val_fs)
    pdf.cell(0, line_h, data['customer_name'], 0, 1)

    pdf.set_font('helvetica', 'B', title_fs)
    pdf.cell(30, line_h, "Customer Address")
    pdf.set_font('helvetica', '', title_fs)
    pdf.cell(5, line_h, ":")
    pdf.set_font('helvetica', '', 7)
    pdf.multi_cell(0, 3.5, data['customer_address'])
    
    pdf.set_y(pdf.get_y() + 2)
    pdf.set_font('helvetica', 'B', 8)
    pdf.cell(275, 5, "Currency : INR", 0, 1, 'R')
    pdf.set_font('helvetica', '', 8)

    # Section 3: Complex Table (Landscape Width 277mm usable)
    # Widths: [60, 18, 22, 18, 18, 12, 15, 12, 15, 12, 15, 12, 15, 31]
    w = [60, 18, 22, 18, 18, 12, 15, 12, 15, 12, 15, 12, 15, 31]
    
    pdf.set_font('helvetica', 'B', 8)
    y_hdr = pdf.get_y()
    
    # Table Header Row 1 (DRAW EMPTY BOXES FOR MULTILINE HEADERS)
    pdf.set_font('helvetica', 'B', 8)
    pdf.cell(w[0], 12, "Description", 1, 0, 'C')
    pdf.cell(w[1], 12, "SAC Code", 1, 0, 'C')
    pdf.cell(w[2], 12, "Taxable Value", 1, 0, 'C')
    # Empty boxes for multiline
    pdf.cell(w[3], 12, "", 1, 0, 'C') 
    pdf.cell(w[4], 12, "Total", 1, 0, 'C')
    
    pdf.cell(w[5]+w[6], 6, "IGST", 1, 0, 'C')
    pdf.cell(w[7]+w[8], 6, "CGST", 1, 0, 'C')
    pdf.cell(w[9]+w[10], 6, "SGST/UTGST", 1, 0, 'C')
    pdf.cell(w[11]+w[12], 6, "CESS", 1, 0, 'C')
    
    # Last column Row 1: Total (Incl Taxes)
    pdf.set_font('helvetica', 'B', 8)
    pdf.cell(w[13], 6, "Total (Incl Taxes)", 1, 1, 'C')
    
    # Table Header Row 2
    pdf.set_xy(10 + w[0] + w[1] + w[2] + w[3] + w[4], y_hdr + 6)
    pdf.set_font('helvetica', 'B', 6)
    for _ in range(4):
        pdf.cell(w[5], 6, "TAX%", 1, 0, 'C')
        pdf.cell(w[6], 6, "Amount", 1, 0, 'C')
        
    # Last column Row 2: Amount
    pdf.set_font('helvetica', 'B', 7)
    pdf.cell(w[13], 6, "Amount", 1, 1, 'C')
    
    # PLACE MULTILINE TEXT FOR NON-TAXABLE PRECISELY
    pdf.set_font('helvetica', 'B', 6.5)
    # Box 4: Non Taxable / Exempted Value
    pdf.set_xy(10+w[0]+w[1]+w[2], y_hdr + 1)
    pdf.multi_cell(w[3], 3.2, "Non Taxable/\nExempted\nvalue", 0, 'C')
    
    pdf.set_y(y_hdr + 12)
    
    # --- Data Row 1 ---
    pdf.set_font('helvetica', '', 9)
    pdf.cell(w[0], 10, data['description'], 1, 0, 'L')
    pdf.cell(w[1], 10, data['sac_code'], 1, 0, 'C')
    pdf.cell(w[2], 10, format_num(data['taxable_value']), 1, 0, 'R')
    pdf.cell(w[3], 10, format_num(data.get('non_taxable', 0)), 1, 0, 'R')
    pdf.cell(w[4], 10, format_num(data['taxable_value']), 1, 0, 'R')
    
    if data['igst'] > 0:
        pdf.cell(w[5], 10, "5%", 1, 0, 'C')
        pdf.cell(w[6], 10, format_num(data['igst']), 1, 0, 'R')
        pdf.cell(w[7], 10, "-", 1, 0, 'C')
        pdf.cell(w[8], 10, "-", 1, 0, 'R')
        pdf.cell(w[9], 10, "-", 1, 0, 'C')
        pdf.cell(w[10], 10, "-", 1, 0, 'R')
    else:
        pdf.cell(w[5], 10, "0%", 1, 0, 'C')
        pdf.cell(w[6], 10, "-", 1, 0, 'R')
        pdf.cell(w[7], 10, "2.5%", 1, 0, 'C')
        pdf.cell(w[8], 10, format_num(data['cgst']), 1, 0, 'R')
        pdf.cell(w[9], 10, "2.5%", 1, 0, 'C')
        pdf.cell(w[10], 10, format_num(data['sgst']), 1, 0, 'R')
        
    pdf.cell(w[11], 10, "-", 1, 0, 'C')
    pdf.cell(w[12], 10, "-", 1, 0, 'R')
    pdf.cell(w[13], 10, format_num(data['total_amount']), 1, 1, 'R')
    
    # --- Row 2 (Airport Charges) ---
    pdf.cell(w[0], 8, "Airport Charges", 1, 0, 'L')
    pdf.cell(w[1], 8, "", 1, 0, 'C')
    pdf.cell(w[2], 8, "", 1, 0, 'C')
    pdf.cell(w[3], 8, "-", 1, 0, 'R')
    pdf.cell(w[4], 8, "-", 1, 0, 'R')
    for i in range(5, 13):
        pdf.cell(w[i], 8, "", 1, 0)
    pdf.cell(w[13], 8, "-", 1, 1, 'R')
    
    # --- Grand Total ---
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(w[0]+w[1], 8, "Grand Total", 1, 0, 'L')
    pdf.cell(w[2], 8, format_num(data['taxable_value']), 1, 0, 'R')
    pdf.cell(w[3], 8, "-", 1, 0, 'R')
    pdf.cell(w[4], 8, format_num(data['taxable_value']), 1, 0, 'R')
    
    pdf.cell(w[5], 8, "", 1, 0)
    pdf.cell(w[6], 8, format_num(data['igst']) if data['igst']>0 else "-", 1, 0, 'R')
    pdf.cell(w[7], 8, "", 1, 0)
    pdf.cell(w[8], 8, format_num(data['cgst']) if data['cgst']>0 else "-", 1, 0, 'R')
    pdf.cell(w[9], 8, "", 1, 0)
    pdf.cell(w[10], 8, format_num(data['sgst']) if data['sgst']>0 else "-", 1, 0, 'R')
    pdf.cell(w[11], 8, "", 1, 0)
    pdf.cell(w[12], 8, "-", 1, 0, 'R')
    pdf.cell(w[13], 8, format_num(data['total_amount']), 1, 1, 'R')
    
    # --- Amount in Words Row ---
    pdf.cell(sum(w[:5]), 8, "Total Invoice Amount (in figures)", 1, 0, 'L')
    pdf.set_font('helvetica', 'B', 9)
    # Remaining width sum(w[5:])
    pdf.cell(sum(w[5:]), 8, f" {data['amount_in_words']}", 1, 1, 'L')
    
    pdf.ln(5)
    
    # Section 4: Footer Notes
    pdf.set_font('helvetica', '', 8)
    notes = [
        "1. Air Travel And Related Charges :- Includes all charges related to air transportation of passengers.",
        "2. Airport Charges :- Includes ADF, UDF, PSF and other airport charges collected on behalf of Airport operator, as applicable",
        "3. Meal :- Includes all prepaid meals purchased before travel",
        "4. Amounts have been rounded off.",
        "5. I/We hereby declare that though our aggregate turnover in any preceding financial year from 2017-18 onwards is more than the aggregate turnover notified under sub-rule (4) of rule 48, we are not required to prepare an invoice in terms of the provisions of the said sub-rule"
    ]
    cur_y = pdf.get_y()
    for note in notes:
        pdf.set_x(10) # Explicitly reset X for every note
        pdf.multi_cell(180, 4, note)
        
    # Signature Right
    pdf.set_y(cur_y)
    pdf.set_font('helvetica', '', 10)
    pdf.cell(0, 5, "For Just Udo Aviation Pvt. Ltd.", 0, 1, 'R')
    
    y_sig = pdf.get_y()
    # Seal and Signature Images
    # Seal and Signature Images (ONLY if provided via UI)
    if seal_pos and os.path.exists(SEAL_PATH):
        pdf.image(SEAL_PATH, seal_pos.get('x', 248), seal_pos.get('y', y_sig - 5), seal_pos.get('w', 30))
        
    if sign_pos and os.path.exists(SIGN_PATH):
        pdf.image(SIGN_PATH, sign_pos.get('x', 250), sign_pos.get('y', y_sig + 1), sign_pos.get('w', 25))
        
    pdf.set_y(y_sig + 15)
    pdf.cell(0, 5, "Authorised Signatory", 0, 1, 'R')
    
    pdf.output(output_path)

def process_all_invoices():
    print("Reading Excel file...")
    df_data = pd.read_excel(EXCEL_FILE, sheet_name='Data')
    
    try:
        df_address = pd.read_excel(EXCEL_FILE, sheet_name='Address Master')
        customer_address_lookup = df_address.set_index(df_address.columns[0])[df_address.columns[2]].fillna('').to_dict()
    except:
        customer_address_lookup = {}
        
    try:
        df_fly91 = pd.read_excel(EXCEL_FILE, sheet_name='FLY91 Address Master')
        supplier_address_lookup = df_fly91.set_index(df_fly91.columns[0])[df_fly91.columns[1]].fillna('').to_dict()
    except:
        supplier_address_lookup = {}

    if not os.path.exists(OUTPUT_BASE_DIR):
        os.makedirs(OUTPUT_BASE_DIR)

    print(f"Generating invoices for {len(df_data)} rows...")
    
    # Process First 50
    for index, row in df_data.head(50).iterrows():
        invoice_no = str(row['Invoicenumber'])
        if pd.isna(invoice_no) or invoice_no == 'nan': continue
        
        supplier_gstin = str(row.get('FLY91 GSTIN', ''))
        supplier_address = supplier_address_lookup.get(supplier_gstin, "Address not found")
        
        customer_gstin = str(row.get('GSTIN', '-'))
        customer_address = customer_address_lookup.get(customer_gstin, "-")
        
        # inv_date = str(row.get('Invoice Date', ''))[:10]
        try:
            raw_date = row.get('Invoice Date', '')
            if pd.isna(raw_date):
                inv_date = ""
            else:
                inv_date = pd.to_datetime(raw_date).strftime('%d-%m-%Y')
        except:
            inv_date = str(row.get('Invoice Date', ''))[:10]
        
        total_val = float(row.get('Invoice Value', 0))
        # Round the total value for amount in words as well
        rounded_total_val = round(total_val)
        
        # Round Flight No
        flight_no_raw = row.get('Flight Number', '')
        try:
            flight_no = str(int(round(float(flight_no_raw))))
        except:
            flight_no = str(flight_no_raw)

        # Round SAC Code
        sac_code_raw = row.get('HSN', '996425')
        try:
            sac_code = str(int(round(float(sac_code_raw))))
        except:
            sac_code = str(sac_code_raw)
        
        data = {
            'supplier_address': supplier_address,
            'supplier_gstin': supplier_gstin,
            'invoice_no': invoice_no,
            'invoice_date': inv_date,
            'passenger_name': str(row.get('Passenger Name', '')).upper(),
            'pnr_no': str(row.get('PNRNumber', '')),
            'flight_no': flight_no,
            'origin': str(row.get('Origin', '')),
            'destination': str(row.get('Destination', '')),
            'place_of_supply': str(row.get('Place of supply - State', '')).upper(),
            'customer_name': str(row.get('Customer Name ', '')).upper(),
            'customer_address': customer_address,
            'customer_gstin': customer_gstin,
            'description': str(row.get('DESCRIPTION ON INVOICE', 'Airport travel and related charges')),
            'sac_code': sac_code,
            'taxable_value': float(row.get('Taxable Value', 0)),
            'igst': float(row.get('IGST', 0)),
            'cgst': float(row.get('CGST', 0)),
            'sgst': float(row.get('SGST', 0)),
            'total_amount': total_val,
            'amount_in_words': number_to_words_indian(rounded_total_val)
        }
        
        folder_name = clean_filename(str(row.get('Folder bifurcation', 'Unknown')))
        target_dir = os.path.join(OUTPUT_BASE_DIR, folder_name)
        if not os.path.exists(target_dir):
            os.makedirs(target_dir)
            
        output_pdf = os.path.join(target_dir, f"{clean_filename(invoice_no)}.pdf")
        generate_kind_pdf(data, output_pdf)
        
        if (index + 1) % 10 == 0:
            print(f"Generated {index+1} invoices...")

def get_excel_data_rows(excel_path):
    df_data = pd.read_excel(excel_path, sheet_name='Data')
    # Drop rows where 'Invoicenumber' is NaN - critical for .xlsm files
    # which report 1,048,575 rows even when most are empty
    df_data = df_data.dropna(subset=['Invoicenumber'])
    df_data = df_data.reset_index(drop=True)
    return df_data

def get_invoicing_data(df_data, index, excel_path):
    try:
        df_address = pd.read_excel(excel_path, sheet_name='Address Master')
        customer_address_lookup = df_address.set_index(df_address.columns[0])[df_address.columns[2]].fillna('').to_dict()
    except:
        customer_address_lookup = {}
        
    try:
        df_fly91 = pd.read_excel(excel_path, sheet_name='FLY91 Address Master')
        supplier_address_lookup = df_fly91.set_index(df_fly91.columns[0])[df_fly91.columns[1]].fillna('').to_dict()
    except:
        supplier_address_lookup = {}

    row = df_data.iloc[index]
    invoice_no = str(row['Invoicenumber'])
    
    supplier_gstin = str(row.get('FLY91 GSTIN', ''))
    supplier_address = supplier_address_lookup.get(supplier_gstin, "Address not found")
    
    customer_gstin = str(row.get('GSTIN', '-'))
    customer_address = customer_address_lookup.get(customer_gstin, "-")
    
    try:
        raw_date = row.get('Invoice Date', '')
        if pd.isna(raw_date):
            inv_date = ""
        else:
            inv_date = pd.to_datetime(raw_date).strftime('%d-%m-%Y')
    except:
        inv_date = str(row.get('Invoice Date', ''))[:10]
    
    total_val = float(row.get('Invoice Value', 0))
    if pd.isna(total_val): total_val = 0.0
    rounded_total_val = round(total_val)
    
    flight_no_raw = row.get('Flight Number', '')
    if pd.isna(flight_no_raw) or flight_no_raw == '':
        flight_no = ""
    else:
        try:
            flight_no = str(int(round(float(flight_no_raw))))
        except:
            flight_no = str(flight_no_raw)

    sac_code_raw = row.get('HSN', '996425')
    if pd.isna(sac_code_raw) or sac_code_raw == '':
        sac_code = "996425"
    else:
        try:
            sac_code = str(int(round(float(sac_code_raw))))
        except:
            sac_code = str(sac_code_raw)
    
    folder_name = clean_filename(str(row.get('Folder bifurcation', 'Unknown')))

    return {
        'supplier_address': supplier_address,
        'supplier_gstin': supplier_gstin,
        'invoice_no': invoice_no,
        'invoice_date': inv_date,
        'passenger_name': str(row.get('Passenger Name', '')).upper(),
        'pnr_no': str(row.get('PNRNumber', '')),
        'flight_no': flight_no,
        'origin': str(row.get('Origin', '')),
        'destination': str(row.get('Destination', '')),
        'place_of_supply': str(row.get('Place of supply - State', '')).upper(),
        'customer_name': str(row.get('Customer Name ', '')).upper(),
        'customer_address': customer_address,
        'customer_gstin': customer_gstin,
        'description': str(row.get('DESCRIPTION ON INVOICE', 'Airport travel and related charges')),
        'sac_code': sac_code,
        'taxable_value': float(row.get('Taxable Value', 0)) if not pd.isna(row.get('Taxable Value')) else 0.0,
        'igst': float(row.get('IGST', 0)) if not pd.isna(row.get('IGST')) else 0.0,
        'cgst': float(row.get('CGST', 0)) if not pd.isna(row.get('CGST')) else 0.0,
        'sgst': float(row.get('SGST', 0)) if not pd.isna(row.get('SGST')) else 0.0,
        'total_amount': total_val,
        'amount_in_words': number_to_words_indian(rounded_total_val),
        'folder_bifurcation': folder_name
    }

if __name__ == "__main__":
    process_all_invoices()
    print("Horizontal Layout complete.")
