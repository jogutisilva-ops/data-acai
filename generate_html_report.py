import pandas as pd
import re
import json
import sqlite3
import os

csv_path = "informe-productos-2026-02-01_2026-05-25.csv"
xlsx_path = "informe-ventas-2026-02-01_2026-05-25.xlsx"

df_csv = pd.read_csv(csv_path)
df_xl = pd.read_excel(xlsx_path, sheet_name='Sheet0')

# Convert types
df_xl['Cantidad'] = pd.to_numeric(df_xl['Cantidad'], errors='coerce')
df_xl['Precio (Bruto)'] = pd.to_numeric(df_xl['Precio (Bruto)'], errors='coerce')
df_xl['Precio (Neto)'] = pd.to_numeric(df_xl['Precio (Neto)'], errors='coerce')
df_xl['Datetime'] = pd.to_datetime(df_xl['Fecha'], format='%d-%m-%Y, %H:%M', errors='coerce')
if df_xl['Datetime'].isna().sum() > 0:
    df_xl['Datetime'] = df_xl['Datetime'].fillna(pd.to_datetime(df_xl['Fecha'], errors='coerce'))

# Clean names
def clean_product_name(name):
    if pd.isna(name):
        return "Sin Descripción"
    
    name_clean = str(name).strip().upper()
    name_clean = re.sub(r'\s+', ' ', name_clean)
    name_clean = re.sub(r'\(INCLUYE PAPAS FRITAS\)', '', name_clean)
    name_clean = re.sub(r'INCLUYE PAPAS FRITAS', '', name_clean)
    name_clean = name_clean.strip().strip('-').strip()
    
    name_clean = re.sub(r'\(?(\d+)\s*ML\)?', r'(\1 ML)', name_clean)
    name_clean = re.sub(r'\(?(\d+)\s*CC\)?', r'\1CC', name_clean)
    name_clean = re.sub(r'\bBEBIDAS\b', 'BEBIDA', name_clean)
    name_clean = re.sub(r'BEBIDA 350\s*CC', 'BEBIDA 350CC', name_clean)
    
    name_clean = re.sub(r'350CC\s*\-?\s*', '350CC - ', name_clean)
    name_clean = re.sub(r'600ML\s*\-?\s*', '600ML - ', name_clean)
    name_clean = re.sub(r'\((\d+)\s*ML\)\s*\-?\s*', r'(\1 ML) - ', name_clean)
    
    name_clean = re.sub(r'CAFETERÍA\s*\-?\s*', 'CAFETERÍA - ', name_clean)
    name_clean = re.sub(r'PAPAS FRITAS\s*\-?\s*FAMILIAR', 'PAPAS FRITAS - FAMILIAR', name_clean)
    name_clean = re.sub(r'PAPAS FRITAS\s*\-?\s*INDIVIDUAL', 'PAPAS FRITAS - INDIVIDUAL', name_clean)
    name_clean = re.sub(r'^JUGO[S]?\s+NATURAL(ES)?\b.*$', 'JUGOS NATURALES', name_clean)
    
    name_clean = re.sub(r'\s*-\s*', ' - ', name_clean)
    name_clean = re.sub(r'\s+', ' ', name_clean)
    name_clean = name_clean.strip().strip('-').strip()
    
    # Milky coffees
    if any(kw in name_clean for kw in ['CAPUCCINO', 'LATTE', 'CORTADO']):
        name_clean = 'CAFETERÍA - CAFÉ CON LECHE'
    elif 'ESPRESSO DOBLE' in name_clean:
        name_clean = 'CAFETERÍA - ESPRESSO DOBLE'
    elif 'ESPRESSO' in name_clean:
        name_clean = 'CAFETERÍA - ESPRESSO'
    elif 'AMERICANO' in name_clean:
        name_clean = 'CAFETERÍA - AMERICANO'
        
    # Waters
    water_keywords = ['AGUA VITAL', 'AGUA SIN GAS', 'AGUA CON GAS']
    if any(wk in name_clean for wk in water_keywords):
        if 'SIN GAS' in name_clean:
            name_clean = 'AGUA VITAL 600ML - SIN GAS'
        else:
            name_clean = 'AGUA VITAL 600ML - CON GAS'
            
    # Fanta Pomelo Zero
    if 'FANTA' in name_clean and 'POMELO' in name_clean and 'ZERO' in name_clean:
        name_clean = 'BEBIDA 350CC - FANTA ZERO POMELO'
        
    # Açaí + Dragon Fruit Mix
    if 'AÇAÍ + DRAGON FRUIT' in name_clean and 'MIXTO' not in name_clean:
        name_clean = name_clean.replace('AÇAÍ + DRAGON FRUIT', 'MIXTO AÇAÍ + DRAGON FRUIT')
        
    return name_clean

# Map categories from CSV
csv_map = {}
for idx, row in df_csv.iterrows():
    name = row['Nombre del producto o servicio']
    var = row['Nombre de la variante']
    cat = row['Categoría']
    if pd.notna(cat) and cat != '':
        if pd.notna(name) and name != '':
            csv_map[str(name).strip().upper()] = cat
            if pd.notna(var) and var != '':
                combined = f"{str(name).strip()} - {str(var).strip()}".upper()
                csv_map[combined] = cat

manual_map = {'PROPINA': 'Propina', 'IMPORTE PERSONALIZADO': 'Importe personalizado'}

def resolve_category(row):
    cat = row['Categoría']
    if pd.notna(cat) and str(cat).strip() != '':
        return cat
    desc = str(row['Descripción']).strip().upper() if pd.notna(row['Descripción']) else ''
    if desc in csv_map: return csv_map[desc]
    if desc in manual_map: return manual_map[desc]
    for key, value in csv_map.items():
        if key in desc or desc in key: return value
    if 'AÇAÍ' in desc or 'BOWL' in desc: return 'AÇAÍ PRIME'
    if 'BURGER' in desc or 'PAPAS' in desc or 'BRISKET' in desc or 'PORK' in desc: return 'AMERICAN PRIME BURGER'
    return 'Otros / Sin Categoría'

df_xl['Categoría_Clean'] = df_xl.apply(resolve_category, axis=1)
df_xl['Clean_Desc'] = df_xl['Descripción'].apply(clean_product_name)

# Commissions calculation
def get_fee_rate(payment_method):
    if pd.isna(payment_method):
        return 0.0
    pm = str(payment_method).upper()
    if 'DÉBITO' in pm or 'DEBITO' in pm:
        return 0.0155
    elif 'CRÉDITO' in pm or 'CREDITO' in pm or 'VISA' in pm or 'AMERICAN' in pm:
        return 0.0250
    elif 'EFECTIVO' in pm:
        return 0.0
    return 0.0

df_xl['Fee_Rate'] = df_xl['Forma de pago'].apply(get_fee_rate)
df_xl['Fee_Amount'] = df_xl['Precio (Bruto)'] * df_xl['Fee_Rate']
df_xl['Net_After_Fee'] = df_xl['Precio (Bruto)'] - df_xl['Fee_Amount']

# Group products
def get_primary_category(group):
    cat_counts = group.groupby('Categoría_Clean')['Cantidad'].sum()
    if cat_counts.empty:
        return 'Otros / Sin Categoría'
    return cat_counts.idxmax()

product_categories = df_xl.groupby('Clean_Desc').apply(get_primary_category).reset_index(name='Primary_Category')
grouped_metrics = df_xl.groupby('Clean_Desc').agg(
    Units_Sold=('Cantidad', 'sum'),
    Gross_Revenue=('Precio (Bruto)', 'sum')
).reset_index()

grouped = pd.merge(grouped_metrics, product_categories, on='Clean_Desc')
grouped = grouped.sort_values(by='Units_Sold', ascending=False)

# Compile raw values for JavaScript
product_list_js = []
for idx, row in enumerate(grouped.iterrows(), 1):
    r_idx, r_data = row
    product_list_js.append({
        'rank': idx,
        'name': r_data['Clean_Desc'],
        'category': r_data['Primary_Category'],
        'units': int(r_data['Units_Sold']),
        'revenue': int(r_data['Gross_Revenue'])
    })

# Compute general KPI metrics
total_gross = int(df_xl['Precio (Bruto)'].sum())
total_fees = int(df_xl['Fee_Amount'].sum())
total_net_real = int(df_xl['Net_After_Fee'].sum())
total_tx = int(df_xl['ID de transacción'].nunique())

# Compute sales by cleaned category
cat_sales = df_xl.groupby('Categoría_Clean').agg(
    Gross=('Precio (Bruto)', 'sum'),
    Units=('Cantidad', 'sum')
).reset_index()
cat_sales_js = []
for idx, row in cat_sales.iterrows():
    cat_sales_js.append({
        'category': row['Categoría_Clean'],
        'gross': int(row['Gross']),
        'units': int(row['Units'])
    })

# Compute Day of Week sales
df_xl['DayOfWeek'] = df_xl['Datetime'].dt.day_name()
df_xl['DayOfWeek_Num'] = df_xl['Datetime'].dt.dayofweek
day_mapping = {
    'Monday': 'Lunes', 'Tuesday': 'Martes', 'Wednesday': 'Miércoles',
    'Thursday': 'Jueves', 'Friday': 'Viernes', 'Saturday': 'Sábado', 'Sunday': 'Domingo'
}
df_xl['DayOfWeek_ES'] = df_xl['DayOfWeek'].map(day_mapping)

dow_sales = df_xl.groupby(['DayOfWeek_Num', 'DayOfWeek_ES']).agg(
    Gross=('Precio (Bruto)', 'sum'),
    Tx=('ID de transacción', 'nunique')
).reset_index().sort_values('DayOfWeek_Num')

dow_sales_js = []
for idx, row in dow_sales.iterrows():
    dow_sales_js.append({
        'day': row['DayOfWeek_ES'],
        'gross': int(row['Gross']),
        'tx': int(row['Tx'])
    })

# Compute Weekly sales
df_xl['Week'] = df_xl['Datetime'].dt.to_period('W').astype(str)
weekly_sales = df_xl.groupby('Week').agg(
    Gross=('Precio (Bruto)', 'sum'),
    Tx=('ID de transacción', 'nunique')
).reset_index().sort_values('Week')

weekly_sales_js = []
for idx, row in weekly_sales.iterrows():
    weekly_sales_js.append({
        'week': row['Week'],
        'gross': int(row['Gross']),
        'tx': int(row['Tx'])
    })

# Compute payment methods and fees
def get_payment_group(pm):
    if pd.isna(pm):
        return "No especificado"
    pm_str = str(pm).upper()
    if 'DÉBITO' in pm_str or 'DEBITO' in pm_str:
        return 'Tarjeta de Débito'
    elif 'CRÉDITO' in pm_str or 'CREDITO' in pm_str or 'VISA' in pm_str or 'AMERICAN' in pm_str:
        return 'Tarjeta de Crédito'
    elif 'EFECTIVO' in pm_str:
        return 'Efectivo'
    return 'No especificado'

df_xl['Forma_Pago_Group'] = df_xl['Forma de pago'].apply(get_payment_group)

payments_summary = df_xl.groupby('Forma_Pago_Group', dropna=False).agg(
    Gross=('Precio (Bruto)', 'sum'),
    Fee_Amount=('Fee_Amount', 'sum'),
    Net_After_Fee=('Net_After_Fee', 'sum'),
    Transactions=('ID de transacción', 'nunique')
).reset_index().sort_values(by='Gross', ascending=False)

payment_list_js = []
for idx, row in payments_summary.iterrows():
    pm_name = str(row['Forma_Pago_Group'])
    if pm_name == 'Tarjeta de Débito':
        rate_pct = 1.55
    elif pm_name == 'Tarjeta de Crédito':
        rate_pct = 2.50
    else:
        rate_pct = 0.00
    payment_list_js.append({
        'name': pm_name,
        'rate': rate_pct,
        'gross': int(row['Gross']),
        'fee': int(row['Fee_Amount']),
        'net': int(row['Net_After_Fee']),
        'tx': int(row['Transactions'])
    })

# Save data summary JSON for the Serverless chat function
data_summary_dict = {
    'total_gross': total_gross,
    'total_fees': total_fees,
    'total_net_real': total_net_real,
    'total_tx': total_tx,
    'category_sales': cat_sales_js,
    'day_of_week_sales': dow_sales_js,
    'weekly_sales': weekly_sales_js,
    'payment_methods': payment_list_js,
    'products': product_list_js,
    'unifications': {
        'Fanta': 'BEBIDA 350CC - FANTA ZERO POMELO (Consolida Fanta Pomelo Zero, Bebidas 350cc - Fanta Pomelo Zero, etc.)',
        'Common Drinks': 'BEBIDA 350CC - COCA COLA, BEBIDA 350CC - KEM (Une ventas de categorías Açaí y Burger)',
        'Water': 'AGUA VITAL 600ML - CON GAS y SIN GAS (Separados por gas, unifica variaciones tipográficas)',
        'Burgers': 'Agrupa por hamburguesa base, eliminando la frase (INCLUYE PAPAS FRITAS)',
        'Milky Coffee': 'CAFETERÍA - CAFÉ CON LECHE (Agrupa Capuccino, Latte, Cortado. Espresso, Espresso Doble y Americano permanecen separados)',
        'Natural Juices': 'JUGOS NATURALES (Agrupa todos los jugos de frutas)'
    }
}
with open('data_summary.json', 'w', encoding='utf-8') as f:
    json.dump(data_summary_dict, f, ensure_ascii=False, indent=2)

# Generar Base de Datos SQLite sales_data.db a partir de df_xl
db_path = "sales_data.db"
if os.path.exists(db_path):
    try:
        os.remove(db_path)
    except Exception:
        pass

try:
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE ventas (
        id_transaccion TEXT,
        fecha_original TEXT,
        fecha_datetime TEXT,
        semana TEXT,
        dia_semana TEXT,
        descripcion_original TEXT,
        producto_limpio TEXT,
        categoria TEXT,
        cantidad INTEGER,
        precio_sin_descuento REAL,
        descuento REAL,
        precio_bruto REAL,
        precio_neto REAL,
        forma_pago TEXT,
        fee_rate REAL,
        fee_amount REAL,
        net_after_fee REAL,
        sku TEXT,
        dispositivo TEXT
    )
    """)
    
    # Asegurar formato string de Datetime
    df_xl['Datetime_Str'] = df_xl['Datetime'].dt.strftime('%Y-%m-%d %H:%M:%S')
    
    rows_to_insert = []
    for idx, row in df_xl.iterrows():
        rows_to_insert.append((
            str(row['ID de transacción']),
            str(row['Fecha']),
            str(row['Datetime_Str']),
            str(row['Week']),
            str(row['DayOfWeek_ES']),
            str(row['Descripción']),
            str(row['Clean_Desc']),
            str(row['Categoría_Clean']),
            int(row['Cantidad']) if pd.notna(row['Cantidad']) else 0,
            float(row['Precio sin descuento']) if pd.notna(row['Precio sin descuento']) else 0.0,
            float(row['Descuento']) if pd.notna(row['Descuento']) else 0.0,
            float(row['Precio (Bruto)']) if pd.notna(row['Precio (Bruto)']) else 0.0,
            float(row['Precio (Neto)']) if pd.notna(row['Precio (Neto)']) else 0.0,
            str(row['Forma de pago']),
            float(row['Fee_Rate']),
            float(row['Fee_Amount']),
            float(row['Net_After_Fee']),
            str(row['SKU']) if pd.notna(row['SKU']) else '',
            str(row['Número de serie del dispositivo']) if pd.notna(row['Número de serie del dispositivo']) else ''
        ))
        
    cursor.executemany("""
    INSERT INTO ventas (
        id_transaccion, fecha_original, fecha_datetime, semana, dia_semana,
        descripcion_original, producto_limpio, categoria, cantidad,
        precio_sin_descuento, descuento, precio_bruto, precio_neto,
        forma_pago, fee_rate, fee_amount, net_after_fee, sku, dispositivo
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, rows_to_insert)
    
    conn.commit()
    conn.close()
    print(f"Base de datos SQLite generada con éxito en {db_path}.")
except Exception as e:
    print(f"Error al generar base de datos SQLite: {e}")


# HTML Construction
html_template = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Açaí Prime - Dashboard de Análisis de Ventas y Comisiones</title>
    <!-- Google Fonts: Outfit & Inter -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Outfit:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <!-- Chart.js CDN -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root {{
            --primary: #5b1a8f; /* Açaí Purple */
            --primary-dark: #3b0e60;
            --primary-light: #f3ecf8;
            --accent: #e91e63; /* Pink Accent */
            --accent-light: #fce4ec;
            --text-main: #2b213a;
            --text-muted: #6e5e82;
            --bg-body: #FAF8FC;
            --card-bg: #FFFFFF;
            --border-color: #E6DFEE;
            --success: #10b981;
            --danger: #ef4444;
            --shadow-sm: 0 4px 6px -1px rgba(91, 26, 143, 0.05), 0 2px 4px -1px rgba(91, 26, 143, 0.03);
            --shadow-md: 0 10px 15px -3px rgba(91, 26, 143, 0.08), 0 4px 6px -2px rgba(91, 26, 143, 0.04);
            --shadow-lg: 0 20px 25px -5px rgba(91, 26, 143, 0.12), 0 10px 10px -5px rgba(91, 26, 143, 0.06);
            --radius-sm: 8px;
            --radius-md: 16px;
            --radius-lg: 24px;
        }}

        @keyframes pulse-accent {{
            0% {{ box-shadow: 0 0 0 0 rgba(233, 30, 99, 0.4); }}
            70% {{ box-shadow: 0 0 0 8px rgba(233, 30, 99, 0); }}
            100% {{ box-shadow: 0 0 0 0 rgba(233, 30, 99, 0); }}
        }}

        .animate-pulse-accent {{
            border: 2px solid var(--accent) !important;
            animation: pulse-accent 2s infinite;
        }}

        .chat-wrapper {{
            display: flex;
            flex-direction: column;
            height: 500px;
            background-color: var(--card-bg);
            border-radius: var(--radius-md);
            overflow: hidden;
            border: 1px solid var(--border-color);
        }}

        .chat-history {{
            flex-grow: 1;
            padding: 24px;
            overflow-y: auto;
            display: flex;
            flex-direction: column;
            gap: 16px;
            background-color: #FAF9FC;
        }}

        .chat-message {{
            max-width: 80%;
            padding: 14px 18px;
            border-radius: var(--radius-md);
            font-size: 14px;
            line-height: 1.5;
            animation: chatSlideUp 0.3s ease;
        }}

        @keyframes chatSlideUp {{
            from {{ opacity: 0; transform: translateY(10px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}

        .message-assistant {{
            align-self: flex-start;
            background-color: white;
            border: 1px solid var(--border-color);
            border-top-left-radius: 4px;
            color: var(--text-main);
        }}

        .message-user {{
            align-self: flex-end;
            background-color: var(--primary);
            color: white;
            border-top-right-radius: 4px;
        }}

        .message-user * {{
            color: white;
        }}

        .chat-suggestions {{
            display: flex;
            gap: 10px;
            padding: 12px 24px;
            background-color: white;
            border-top: 1px solid var(--border-color);
            flex-wrap: wrap;
        }}

        .suggestion-chip {{
            background-color: var(--primary-light);
            color: var(--primary);
            padding: 8px 14px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
            border: 1px solid transparent;
        }}

        .suggestion-chip:hover {{
            background-color: var(--primary);
            color: white;
        }}

        .chat-input-bar {{
            display: flex;
            padding: 16px 24px;
            background-color: white;
            border-top: 1px solid var(--border-color);
            gap: 12px;
            align-items: center;
        }}

        .chat-input {{
            flex-grow: 1;
            padding: 12px 18px;
            border-radius: var(--radius-sm);
            border: 1px solid var(--border-color);
            font-size: 14px;
            outline: none;
            transition: border-color 0.3s ease;
            height: 46px;
        }}

        .chat-input:focus {{
            border-color: var(--primary);
        }}

        .chat-send-btn {{
            background-color: var(--primary);
            color: white;
            border: none;
            padding: 0 24px;
            border-radius: var(--radius-sm);
            font-weight: 600;
            cursor: pointer;
            height: 46px;
            transition: background-color 0.2s ease;
            display: flex;
            align-items: center;
            justify-content: center;
        }}

        .chat-send-btn:hover {{
            background-color: var(--primary-dark);
        }}

        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Inter', sans-serif;
            color: var(--text-main);
        }}

        body {{
            background-color: var(--bg-body);
            min-height: 100vh;
            padding-bottom: 50px;
        }}

        header {{
            background: linear-gradient(135deg, var(--primary) 0%, var(--primary-dark) 100%);
            padding: 24px 8%;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: var(--shadow-md);
            position: sticky;
            top: 0;
            z-index: 100;
        }}

        .logo-container {{
            display: flex;
            align-items: center;
            gap: 12px;
        }}

        .logo-symbol {{
            background: linear-gradient(135deg, var(--accent) 0%, #ff6090 100%);
            width: 40px;
            height: 40px;
            border-radius: 50%;
            display: flex;
            justify-content: center;
            align-items: center;
            font-weight: 800;
            color: white;
            font-size: 20px;
            font-family: 'Outfit', sans-serif;
            box-shadow: 0 4px 10px rgba(233, 30, 99, 0.4);
        }}

        .logo-text {{
            font-family: 'Outfit', sans-serif;
            font-weight: 800;
            font-size: 24px;
            color: white;
            letter-spacing: -0.5px;
        }}

        .logo-text span {{
            color: #ff6090;
        }}

        .header-tag {{
            background-color: rgba(255, 255, 255, 0.15);
            padding: 8px 16px;
            border-radius: 20px;
            color: white;
            font-size: 14px;
            font-weight: 500;
            backdrop-filter: blur(5px);
            border: 1px solid rgba(255, 255, 255, 0.1);
        }}

        .container {{
            max-width: 1400px;
            margin: 40px auto 0 auto;
            padding: 0 24px;
        }}

        /* Navigation Tabs */
        .tabs {{
            display: flex;
            gap: 12px;
            margin-bottom: 30px;
            border-bottom: 2px solid var(--border-color);
            padding-bottom: 12px;
        }}

        .tab-btn {{
            background: none;
            border: none;
            padding: 12px 24px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            border-radius: var(--radius-sm);
            transition: all 0.3s ease;
            color: var(--text-muted);
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        .tab-btn:hover {{
            background-color: var(--primary-light);
            color: var(--primary);
        }}

        .tab-btn.active {{
            background-color: var(--primary);
            color: white;
            box-shadow: 0 4px 12px rgba(91, 26, 143, 0.2);
        }}

        /* KPI Cards Grid */
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
            gap: 24px;
            margin-bottom: 40px;
        }}

        .kpi-card {{
            background-color: var(--card-bg);
            border-radius: var(--radius-md);
            padding: 24px;
            box-shadow: var(--shadow-sm);
            border: 1px solid var(--border-color);
            position: relative;
            overflow: hidden;
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }}

        .kpi-card:hover {{
            transform: translateY(-5px);
            box-shadow: var(--shadow-md);
        }}

        .kpi-card::before {{
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 4px;
            height: 100%;
            background-color: var(--primary);
        }}

        .kpi-card.accent::before {{
            background-color: var(--accent);
        }}

        .kpi-card.success::before {{
            background-color: var(--success);
        }}

        .kpi-card.danger::before {{
            background-color: var(--danger);
        }}

        .kpi-title {{
            font-size: 14px;
            color: var(--text-muted);
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 8px;
        }}

        .kpi-value {{
            font-family: 'Outfit', sans-serif;
            font-size: 28px;
            font-weight: 700;
            color: var(--text-main);
            margin-bottom: 4px;
        }}

        .kpi-desc {{
            font-size: 12px;
            color: var(--text-muted);
        }}

        /* Sections and Card layouts */
        .tab-content {{
            display: none;
        }}

        .tab-content.active {{
            display: block;
        }}

        .dashboard-row {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 30px;
            margin-bottom: 40px;
        }}

        @media (max-width: 1024px) {{
            .dashboard-row {{
                grid-template-columns: 1fr;
            }}
        }}

        .section-card {{
            background-color: var(--card-bg);
            border-radius: var(--radius-md);
            padding: 30px;
            box-shadow: var(--shadow-sm);
            border: 1px solid var(--border-color);
            margin-bottom: 30px;
        }}

        .section-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 24px;
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 16px;
        }}

        .section-title {{
            font-family: 'Outfit', sans-serif;
            font-size: 20px;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 10px;
            color: var(--primary-dark);
        }}

        .chart-container {{
            position: relative;
            height: 320px;
            width: 100%;
        }}

        /* Data Tables */
        table {{
            width: 100%;
            border-collapse: collapse;
            text-align: left;
            font-size: 14px;
        }}

        th {{
            background-color: var(--primary-light);
            color: var(--primary);
            font-weight: 700;
            padding: 14px 16px;
            border-bottom: 2px solid var(--border-color);
            text-transform: uppercase;
            font-size: 11px;
            letter-spacing: 0.5px;
        }}

        td {{
            padding: 14px 16px;
            border-bottom: 1px solid var(--border-color);
            color: var(--text-main);
            vertical-align: middle;
        }}

        tr:last-child td {{
            border-bottom: none;
        }}

        tr:hover td {{
            background-color: #FAF9FC;
        }}

        /* Category tags */
        .badge {{
            display: inline-block;
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
        }}

        .badge-acai {{
            background-color: #EDE7F6;
            color: #5E35B1;
        }}

        .badge-burger {{
            background-color: #FFF3E0;
            color: #E65100;
        }}

        .badge-tip {{
            background-color: #E8F5E9;
            color: #2E7D32;
        }}

        .badge-other {{
            background-color: #ECEFF1;
            color: #37474F;
        }}

        .badge-extra {{
            background-color: #FCE4EC;
            color: #C2185B;
        }}

        .badge-custom {{
            background-color: #E0F7FA;
            color: #00838F;
        }}

        /* Search & Filter Controls */
        .filter-bar {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 20px;
            margin-bottom: 24px;
            flex-wrap: wrap;
        }}

        .search-wrapper {{
            position: relative;
            flex-grow: 1;
            max-width: 450px;
        }}

        .search-input {{
            width: 100%;
            padding: 12px 16px 12px 42px;
            border-radius: var(--radius-sm);
            border: 1px solid var(--border-color);
            font-size: 14px;
            outline: none;
            transition: all 0.3s ease;
        }}

        .search-input:focus {{
            border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(91, 26, 143, 0.1);
        }}

        .search-icon {{
            position: absolute;
            left: 14px;
            top: 50%;
            transform: translateY(-50%);
            color: var(--text-muted);
            width: 18px;
            height: 18px;
            pointer-events: none;
        }}

        .select-filter {{
            padding: 12px 16px;
            border-radius: var(--radius-sm);
            border: 1px solid var(--border-color);
            font-size: 14px;
            background-color: white;
            outline: none;
            cursor: pointer;
            min-width: 180px;
            transition: all 0.3s ease;
        }}

        .select-filter:focus {{
            border-color: var(--primary);
        }}

        .product-rank {{
            font-weight: 700;
            color: var(--text-muted);
            width: 40px;
        }}

        .text-bold {{
            font-weight: 600;
        }}

        /* Consolidations detail grid */
        .consolidations-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
            gap: 24px;
        }}

        .consolidations-card {{
            background-color: var(--card-bg);
            border-radius: var(--radius-md);
            padding: 24px;
            border: 1px solid var(--border-color);
            box-shadow: var(--shadow-sm);
        }}

        .consolidations-card h3 {{
            font-family: 'Outfit', sans-serif;
            font-size: 16px;
            color: var(--primary);
            margin-bottom: 12px;
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        .consolidations-card ul {{
            list-style: none;
            padding-left: 0;
        }}

        .consolidations-card li {{
            font-size: 13px;
            padding: 8px 0;
            border-bottom: 1px dashed var(--border-color);
            color: var(--text-muted);
        }}

        .consolidations-card li:last-child {{
            border-bottom: none;
            padding-bottom: 0;
        }}

        .consolidations-card strong {{
            color: var(--text-main);
        }}

        /* Footer styling */
        footer {{
            text-align: center;
            padding: 40px 24px;
            color: var(--text-muted);
            font-size: 13px;
            border-top: 1px solid var(--border-color);
            margin-top: 60px;
        }}
    </style>
</head>
<body>

    <header>
        <div class="logo-container">
            <div class="logo-symbol">AP</div>
            <div class="logo-text">Açaí<span>Prime</span></div>
        </div>
        <div class="header-tag">📊 Dashboard de Análisis y Comisiones (Feb 11 - May 25, 2026)</div>
    </header>

    <div class="container">
        
        <!-- Navigation tabs -->
        <div class="tabs">
            <button class="tab-btn active" onclick="switchTab('tab-dashboard')">
                <svg class="size-4" style="fill: currentColor" viewBox="0 0 24 24"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-9 14H5v-2h5v2zm0-4H5v-2h5v2zm0-4H5V7h5v2zm9 8h-7v-2h7v2zm0-4h-7v-2h7v2zm0-4h-7V7h7v2z"/></svg>
                Dashboard y Ventas
            </button>
            <button class="tab-btn" onclick="switchTab('tab-fees')">
                <svg class="size-4" style="fill: currentColor" viewBox="0 0 24 24"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 17h-2v-2h2v2zm2.07-7.75l-.9.92C13.45 12.9 13 13.5 13 15h-2v-.5c0-1.1.45-2.1 1.17-2.83l1.24-1.26c.37-.36.59-.86.59-1.41 0-1.1-.9-2-2-2s-2 .9-2 2H7c0-2.76 2.24-5 5-5s5 2.24 5 5c0 1.04-.42 1.99-1.07 2.75z"/></svg>
                Trazabilidad de Comisiones
            </button>
            <button class="tab-btn" onclick="switchTab('tab-products')">
                <svg class="size-4" style="fill: currentColor" viewBox="0 0 24 24"><path d="M20 4H4c-1.1 0-1.99.9-1.99 2L2 18c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm-5 14H4v-4h11v4zm0-5H4V9h11v4zm5 5h-4V9h4v9z"/></svg>
                Ventas por Producto
            </button>
            <button class="tab-btn" onclick="switchTab('tab-unifications')">
                <svg class="size-4" style="fill: currentColor" viewBox="0 0 24 24"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-6h2v6zm0-8h-2V7h2v2z"/></svg>
                Consolidaciones
            </button>
            <button class="tab-btn animate-pulse-accent" onclick="switchTab('tab-assistant')">
                <svg class="size-4" style="fill: currentColor" viewBox="0 0 24 24"><path d="M20 2H4c-1.1 0-2 .9-2 2v18l4-4h14c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm0 14H6l-2 2V4h16v12zm-3-5h-2V9h2v2zm-4 0h-2V9h2v2zm-4 0H7V9h2v2z"/></svg>
                Asistente de Consultas (AI)
            </button>
        </div>

        <!-- KPI Summary Cards -->
        <div class="kpi-grid">
            <div class="kpi-card">
                <div class="kpi-title">Ventas Brutas</div>
                <div class="kpi-value">${total_gross:,.0f} CLP</div>
                <div class="kpi-desc">Total acumulado en el periodo</div>
            </div>
            <div class="kpi-card danger">
                <div class="kpi-title">Comisiones Totales</div>
                <div class="kpi-value">${total_fees:,.0f} CLP</div>
                <div class="kpi-desc">Débito: 1.55% | Crédito: 2.50% | Efectivo: 0%</div>
            </div>
            <div class="kpi-card success">
                <div class="kpi-title">Ingreso Neto Real</div>
                <div class="kpi-value">${total_net_real:,.0f} CLP</div>
                <div class="kpi-desc">Ingresos recibidos (Bruto - Comisión)</div>
            </div>
            <div class="kpi-card accent">
                <div class="kpi-title">Transacciones</div>
                <div class="kpi-value">{total_tx:,}</div>
                <div class="kpi-desc">Ventas totales procesadas</div>
            </div>
        </div>

        <!-- TAB 1: General Dashboard -->
        <div id="tab-dashboard" class="tab-content active">
            
            <div class="dashboard-row">
                <!-- Day of Week Chart & Table -->
                <div class="section-card">
                    <div class="section-header">
                        <div class="section-title">
                            <svg class="size-4" style="fill: var(--primary)" viewBox="0 0 24 24"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm.5-13H11v6l5.25 3.15.75-1.23-4.5-2.67z"/></svg>
                            Ventas por Día de la Semana
                        </div>
                    </div>
                    <div class="chart-container" style="margin-bottom: 24px;">
                        <canvas id="dowChart"></canvas>
                    </div>
                    <table>
                        <thead>
                            <tr>
                                <th>Día</th>
                                <th>Transacciones</th>
                                <th>Venta Bruta</th>
                                <th>% Participación</th>
                            </tr>
                        </thead>
                        <tbody>
                            """
for d in dow_sales_js:
    pct = (d['gross'] / total_gross) * 100
    html_template += f"""
                            <tr>
                                <td class="text-bold">{d['day']}</td>
                                <td>{d['tx']:,}</td>
                                <td class="text-bold">${d['gross']:,.0f} CLP</td>
                                <td>{pct:.1f}%</td>
                            </tr>"""

html_template += f"""
                        </tbody>
                    </table>
                </div>

                <!-- Category sales Chart & Table -->
                <div class="section-card">
                    <div class="section-header">
                        <div class="section-title">
                            <svg class="size-4" style="fill: var(--primary)" viewBox="0 0 24 24"><path d="M11 2v20c-5.07-.5-9-4.79-9-10s3.93-9.5 9-10zm2.03 0v8.99H22c-.47-4.74-4.24-8.52-8.97-8.99zm0 11.01V22c4.73-.47 8.5-4.25 8.97-8.99h-8.97z"/></svg>
                            Ventas por Categoría
                        </div>
                    </div>
                    <div class="chart-container" style="margin-bottom: 24px;">
                        <canvas id="categoryChart"></canvas>
                    </div>
                    <table>
                        <thead>
                            <tr>
                                <th>Categoría</th>
                                <th>Unidades</th>
                                <th>Monto Bruto</th>
                                <th>% Participación</th>
                            </tr>
                        </thead>
                        <tbody>
                            """
for c in cat_sales_js:
    pct = (c['gross'] / total_gross) * 100
    # Determine badge type
    badge_class = "badge-other"
    if c['category'] == 'AÇAÍ PRIME': badge_class = "badge-acai"
    elif c['category'] == 'AMERICAN PRIME BURGER': badge_class = "badge-burger"
    elif c['category'] == 'Propina': badge_class = "badge-tip"
    elif c['category'] == 'Extras': badge_class = "badge-extra"
    elif c['category'] == 'Importe personalizado': badge_class = "badge-custom"

    html_template += f"""
                            <tr>
                                <td><span class="badge {badge_class}">{c['category']}</span></td>
                                <td>{c['units']:,}</td>
                                <td class="text-bold">${c['gross']:,.0f} CLP</td>
                                <td>{pct:.1f}%</td>
                            </tr>"""

html_template += f"""
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- Weekly trend line chart -->
            <div class="section-card">
                <div class="section-header">
                    <div class="section-title">
                        <svg class="size-4" style="fill: var(--primary)" viewBox="0 0 24 24"><path d="M16 6l2.29 2.29-4.88 4.88-4-4L2 16.59 3.41 18l6-6 4 4 6.3-6.29L22 12V6z"/></svg>
                        Tendencia de Compras Semanales
                    </div>
                </div>
                <div class="chart-container" style="height: 250px;">
                    <canvas id="weeklyChart"></canvas>
                </div>
            </div>

        </div>

        <!-- TAB 2: Fee Traceability Dashboard -->
        <div id="tab-fees" class="tab-content">
            <div class="dashboard-row">
                <!-- Fee breakdown pie chart -->
                <div class="section-card">
                    <div class="section-header">
                        <div class="section-title">
                            <svg class="size-4" style="fill: var(--danger)" viewBox="0 0 24 24"><path d="M12 2c1.1 0 2 .9 2 2v6h6c1.1 0 2 .9 2 2s-.9 2-2 2h-6v6c0 1.1-.9 2-2 2s-2-.9-2-2v-6H4c-1.1 0-2-.9-2-2s.9-2 2-2h6V4c0-1.1.9-2 2-2z"/></svg>
                            Distribución de Comisiones de Pago
                        </div>
                    </div>
                    <div class="chart-container" style="margin-bottom: 24px;">
                        <canvas id="feesChart"></canvas>
                    </div>
                </div>

                <!-- Summary information text -->
                <div class="section-card" style="display: flex; flex-direction: column; justify-content: center;">
                    <div class="section-title" style="margin-bottom: 16px; color: var(--primary);">
                        🔒 Impacto de las Comisiones en Margen
                    </div>
                    <p style="font-size: 14px; line-height: 1.6; color: var(--text-muted); margin-bottom: 12px;">
                        De un total bruto de <strong>${total_gross:,.0f} CLP</strong>, un total de <strong>${total_fees:,.0f} CLP ({(total_fees/total_gross)*100:.2f}%)</strong> se destina al pago de comisiones del procesador de pagos. 
                    </p>
                    <p style="font-size: 14px; line-height: 1.6; color: var(--text-muted); margin-bottom: 12px;">
                        La comisión promedio ponderada en tus ventas con tarjeta es aproximadamente del <strong>1.84%</strong>, empujada fuertemente por la alta preferencia por tarjetas de débito.
                    </p>
                    <div style="background-color: var(--primary-light); padding: 16px; border-radius: var(--radius-sm); border-left: 4px solid var(--primary); font-size: 13px; font-weight: 500;">
                        💡 <strong>Tip Comercial:</strong> Dado que el 96.9% de tus ventas se realizan con tarjeta, optimizar tus contratos adquirentes (por ejemplo, negociar una rebaja del 1.55% al 1.35% en débito) podría ahorrarte más de <strong>$1.500.000 CLP</strong> anuales.
                    </div>
                </div>
            </div>

            <!-- Detailed fees table -->
            <div class="section-card">
                <div class="section-header">
                    <div class="section-title">
                        📊 Trazabilidad de Comisiones por Medio de Pago
                    </div>
                </div>
                <table>
                    <thead>
                        <tr>
                            <th>Método de Pago</th>
                            <th>Transacciones</th>
                            <th>Venta Bruta</th>
                            <th>Tasa de Comisión</th>
                            <th style="color: var(--danger);">Comisiones (CLP)</th>
                            <th style="color: var(--success);">Ingreso Neto (CLP)</th>
                            <th>% del Total</th>
                        </tr>
                    </thead>
                    <tbody>
                        """
for p in payment_list_js:
    pct = (p['gross'] / total_gross) * 100
    html_template += f"""
                        <tr>
                            <td class="text-bold">{p['name']}</td>
                            <td>{p['tx']:,}</td>
                            <td>${p['gross']:,.0f} CLP</td>
                            <td class="text-bold" style="color: var(--text-muted);">{p['rate']:.2f}%</td>
                            <td class="text-bold" style="color: var(--danger);">${p['fee']:,.0f} CLP</td>
                            <td class="text-bold" style="color: var(--success);">${p['net']:,.0f} CLP</td>
                            <td>{pct:.1f}%</td>
                        </tr>"""

html_template += f"""
                    </tbody>
                </table>
            </div>
        </div>

        <!-- TAB 3: Consolidated Products List -->
        <div id="tab-products" class="tab-content">
            <div class="section-card">
                <div class="section-header">
                    <div class="section-title">
                        <svg class="size-4" style="fill: var(--primary)" viewBox="0 0 24 24"><path d="M7 18c-1.1 0-1.99.9-1.99 2S5.9 22 7 22s2-.9 2-2-.9-2-2-2zM1 2v2h2l3.6 7.59-1.35 2.45c-.16.28-.25.61-.25.96 0 1.1.9 2 2 2h12v-2H7.42c-.14 0-.25-.11-.25-.25l.03-.12.9-1.63h7.45c.75 0 1.41-.41 1.75-1.03l3.58-6.49c.08-.14.12-.31.12-.48 0-.55-.45-1-1-1H5.21l-.94-2H1zm16 16c-1.1 0-1.99.9-1.99 2s.9 2 1.99 2 2-.9 2-2-.9-2-2-2z"/></svg>
                        Reporte de Ventas por Producto Base
                    </div>
                </div>

                <!-- NEW Filters Section: Category, Units, Amount -->
                <div class="filter-bar" style="background-color: var(--primary-light); padding: 20px; border-radius: var(--radius-md); margin-bottom: 20px; display: flex; flex-direction: column; gap: 16px; width: 100%;">
                    <div class="search-wrapper" style="width: 100%; max-width: 100%; position: relative;">
                        <!-- Search Icon -->
                        <svg class="search-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="position: absolute; left: 14px; top: 50%; transform: translateY(-50%); color: var(--text-muted); width: 18px; height: 18px; pointer-events: none;"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
                        <input type="text" id="productSearch" class="search-input" placeholder="Buscar producto por nombre..." onkeyup="filterProducts()" style="width: 100%; padding: 12px 16px 12px 42px; border-radius: var(--radius-sm); border: 1px solid var(--border-color); font-size: 14px; outline: none; transition: all 0.3s ease;">
                    </div>
                    
                    <div style="display: flex; gap: 16px; flex-wrap: wrap; align-items: flex-end; width: 100%;">
                        <div style="flex: 1; min-width: 200px;">
                            <label style="font-size: 11px; text-transform: uppercase; font-weight: 700; color: var(--text-muted); display: block; margin-bottom: 6px; letter-spacing: 0.5px;">Categoría</label>
                            <select id="categoryFilter" class="select-filter" onchange="filterProducts()" style="width: 100%; height: 46px; padding: 12px 16px; border-radius: var(--radius-sm); border: 1px solid var(--border-color); font-size: 14px; background-color: white; outline: none; cursor: pointer; transition: all 0.3s ease;">
                                <option value="ALL">Todas las Categorías</option>
                            </select>
                        </div>
                        
                        <div style="flex: 1; min-width: 140px;">
                            <label style="font-size: 11px; text-transform: uppercase; font-weight: 700; color: var(--text-muted); display: block; margin-bottom: 6px; letter-spacing: 0.5px;">Mínimo Unidades</label>
                            <input type="number" id="minUnits" class="select-filter" placeholder="Ej. 10" onkeyup="filterProducts()" onchange="filterProducts()" style="width: 100%; height: 46px; padding: 12px 16px; border-radius: var(--radius-sm); border: 1px solid var(--border-color); font-size: 14px; background-color: white; outline: none; transition: all 0.3s ease;">
                        </div>
                        
                        <div style="flex: 1; min-width: 160px;">
                            <label style="font-size: 11px; text-transform: uppercase; font-weight: 700; color: var(--text-muted); display: block; margin-bottom: 6px; letter-spacing: 0.5px;">Monto Mínimo (CLP)</label>
                            <input type="number" id="minAmount" class="select-filter" placeholder="Ej. 50000" onkeyup="filterProducts()" onchange="filterProducts()" style="width: 100%; height: 46px; padding: 12px 16px; border-radius: var(--radius-sm); border: 1px solid var(--border-color); font-size: 14px; background-color: white; outline: none; transition: all 0.3s ease;">
                        </div>
                        
                        <div>
                            <button onclick="clearFilters()" style="background-color: var(--accent); color: white; border: none; padding: 12px 20px; border-radius: var(--radius-sm); font-size: 14px; font-weight: 600; cursor: pointer; height: 46px; transition: all 0.3s ease; display: flex; align-items: center; gap: 8px;">
                                <svg style="width: 16px; height: 16px; fill: currentColor;" viewBox="0 0 24 24"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12 19 6.41z"/></svg>
                                Limpiar Filtros
                            </button>
                        </div>
                    </div>
                </div>

                <!-- Filter Summary Stats -->
                <div id="filterSummary" style="display: flex; gap: 24px; padding: 14px 20px; background-color: white; border-radius: var(--radius-sm); border: 1px solid var(--border-color); margin-bottom: 24px; font-size: 14px; color: var(--text-muted); align-items: center; flex-wrap: wrap;">
                    <div>Mostrando: <strong id="statsCount" style="color: var(--primary);">0</strong> productos</div>
                    <div style="width: 1px; height: 16px; background-color: var(--border-color); display: inline-block;"></div>
                    <div>Total Unidades: <strong id="statsUnits" style="color: var(--text-main);">0</strong></div>
                    <div style="width: 1px; height: 16px; background-color: var(--border-color); display: inline-block;"></div>
                    <div>Total Venta Bruta: <strong id="statsRevenue" style="color: var(--success);">$0 CLP</strong></div>
                </div>

                <!-- Product Table -->
                <div style="overflow-x: auto;">
                    <table id="productTable">
                        <thead>
                            <tr>
                                <th style="width: 60px;">Ránking</th>
                                <th>Producto</th>
                                <th>Categoría</th>
                                <th style="cursor: pointer;" onclick="sortProducts('units')">Unidades Vendidas ↕</th>
                                <th style="cursor: pointer;" onclick="sortProducts('revenue')">Monto Bruto (CLP) ↕</th>
                            </tr>
                        </thead>
                        <tbody id="productTableBody">
                            <!-- Populated by JS -->
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- TAB 4: Unifications Applied -->
        <div id="tab-unifications" class="tab-content">
            <div class="section-card">
                <div class="section-header">
                    <div class="section-title">
                        <svg class="size-4" style="fill: var(--primary)" viewBox="0 0 24 24"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-6h2v6zm0-8h-2V7h2v2z"/></svg>
                        Criterios de Consolidación de Nombres
                    </div>
                </div>
                <p style="margin-bottom: 24px; color: var(--text-muted); font-size: 14px; line-height: 1.6;">
                    Para evitar la dispersión de datos generada por diferencias menores en el registro de productos, agrupamos las transacciones que correspondían al mismo artículo bajo un único nombre base consolidado. A continuación se muestran los principales criterios aplicados:
                </p>

                <div class="consolidations-grid">
                    <div class="consolidations-card">
                        <h3>🥤 Bebidas de Fanta</h3>
                        <ul>
                            <li>Unificado bajo: <strong>BEBIDA 350CC - FANTA ZERO POMELO</strong></li>
                            <li>Nombres agrupados: <em>Fanta Pomelo Zero, Fanta Zero Pomelo, Bebidas 350cc - Fanta Pomelo Zero, Fanta Zero Pomelo</em></li>
                        </ul>
                    </div>
                    <div class="consolidations-card">
                        <h3>🥤 Bebidas Comunes</h3>
                        <ul>
                            <li>Unificado bajo: <strong>BEBIDA 350CC - KEM</strong> y <strong>BEBIDA 350CC - COCA COLA</strong></li>
                            <li>Resuelve duplicados generados por ventas registradas indistintamente en la categoría Açaí Prime o American Prime Burger.</li>
                        </ul>
                    </div>
                    <div class="consolidations-card">
                        <h3>💧 Aguas</h3>
                        <ul>
                            <li>Unificado bajo: <strong>AGUA VITAL 600ML - CON GAS</strong> y <strong>AGUA VITAL 600ML - SIN GAS</strong></li>
                            <li>Agrupa todas las variaciones tipográficas (con o sin paréntesis, mayúsculas y minúsculas) respetando la distinción del tipo de gas.</li>
                        </ul>
                    </div>
                    <div class="consolidations-card">
                        <h3>🍔 Hamburguesas</h3>
                        <ul>
                            <li>Unificado bajo: <strong>CHEESE BURGER</strong>, <strong>TEXAS BURGER</strong>, etc.</li>
                            <li>Agrupa las ventas registradas con y sin la anotación "(INCLUYE PAPAS FRITAS)", uniendo el volumen total de la hamburguesa base.</li>
                        </ul>
                    </div>
                    <div class="consolidations-card">
                        <h3>☕ Cafés con Leche</h3>
                        <ul>
                            <li>Unificado bajo: <strong>CAFETERÍA - CAFÉ CON LECHE</strong></li>
                            <li>Consolida los cafés preparados con leche (*Capuccino, Latte y Cortado*), mientras que se mantienen independientes los cafés negros (*Espresso, Espresso Doble y Americano*).</li>
                        </ul>
                    </div>
                    <div class="consolidations-card">
                        <h3>🍹 Jugos Naturales</h3>
                        <ul>
                            <li>Unificado bajo: <strong>JUGOS NATURALES</strong></li>
                            <li>Suma todas las ventas de jugos de frutas (Frambuesa, Mango, Chirimoya, Maracuyá, Frutos Rojos) para entender el volumen global de jugos.</li>
                        </ul>
                    </div>
                </div>
            </div>
        </div>

        <!-- TAB 5: Query Assistant (AI) -->
        <div id="tab-assistant" class="tab-content">
            <div class="section-card" style="padding: 0; overflow: hidden; border-radius: var(--radius-md);">
                <div style="padding: 24px; background: linear-gradient(135deg, var(--primary) 0%, var(--primary-dark) 100%); color: white;">
                    <h2 style="color: white; font-family: 'Outfit', sans-serif; font-size: 20px; font-weight: 700; margin-bottom: 4px; display: flex; align-items: center; gap: 8px;">
                        💬 Asistente Virtual Açaí Prime
                    </h2>
                    <p style="color: rgba(255, 255, 255, 0.8); font-size: 13px;">Haz preguntas en lenguaje natural sobre las ventas, comisiones y los criterios de consolidación de productos.</p>
                </div>
                
                <div class="chat-wrapper">
                    <!-- Chat Message History -->
                    <div id="chatHistory" class="chat-history">
                        <div class="chat-message message-assistant">
                            ¡Hola! Soy tu asistente de datos de <strong>Açaí Prime</strong>. Puedo darte respuestas y análisis inmediatos basados en las planillas procesadas. Pregúntame sobre el producto más vendido, comisiones de pago, ingresos netos o escribe el nombre de un producto para ver su detalle.<br><br>
                            ¿Qué te gustaría analizar hoy?
                        </div>
                    </div>
                    
                    <!-- Suggested Questions Chips -->
                    <div class="chat-suggestions">
                        <span class="suggestion-chip" onclick="askSuggested('dame el registro de venta por producto de la semana 2026-04-20/2026-04-26')">📊 Ventas del 20 de Abril</span>
                        <span class="suggestion-chip" onclick="askSuggested('¿Cuáles son los 5 productos con mayor ingreso neto después de comisiones?')">💰 Top 5 Productos Netos</span>
                        <span class="suggestion-chip" onclick="askSuggested('¿Cuál es la hora peak de transacciones los días domingo?')">⚡ Hora Peak Domingos</span>
                        <span class="suggestion-chip" onclick="askSuggested('¿Cuál es la participación y ticket promedio del Efectivo vs Débito?')">💳 Efectivo vs Débito</span>
                        <span class="suggestion-chip" onclick="askSuggested('¿Cuántas ventas con un 20% de descuento se registraron desde el 11 de mayo?')">🏷️ Descuentos 20%</span>
                    </div>

                    
                    <!-- Chat Input Field -->
                    <div class="chat-input-bar">
                        <input type="text" id="chatInput" class="chat-input" placeholder="Pregúntame algo, ej. ¿Cuánto vendió Cheese Burger?" onkeydown="checkChatEnter(event)">
                        <button onclick="sendChatMessage()" class="chat-send-btn">
                            Enviar
                        </button>
                    </div>
                </div>
            </div>
        </div>

    </div>

    <footer>
        <p>© 2026 Açaí Prime Chile • Reporte de Análisis Comercial</p>
        <p style="font-size: 11px; margin-top: 5px; opacity: 0.7;">Diseño y análisis generado de forma autónoma con datos unificados de ventas</p>
    </footer>

    <script>
        // Data injected from Python
        const productsData = {json.dumps(product_list_js)};
        const dowData = {json.dumps(dow_sales_js)};
        const weeklyData = {json.dumps(weekly_sales_js)};
        const catData = {json.dumps(cat_sales_js)};
        const paymentData = {json.dumps(payment_list_js)};
        
        // Totals injected from Python
        const totalGross = {total_gross};
        const totalFees = {total_fees};
        const totalNetReal = {total_net_real};
        const totalTx = {total_tx};

        let currentSort = {{ column: 'units', direction: 'desc' }};

        // Initialize App
        document.addEventListener('DOMContentLoaded', () => {{
            populateCategoryFilter();
            renderProductTable(productsData);
            updateStatsSummary(productsData);
            initCharts();
        }});

        // Tab Switcher
        function switchTab(tabId) {{
            document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
            document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
            
            document.getElementById(tabId).classList.add('active');
            
            // Find button matching tabId click
            const btnIdx = tabId === 'tab-dashboard' ? 0 : tabId === 'tab-fees' ? 1 : tabId === 'tab-products' ? 2 : tabId === 'tab-unifications' ? 3 : 4;
            document.querySelectorAll('.tab-btn')[btnIdx].classList.add('active');
        }}

        // Populate Category Filter Options Dynamically
        function populateCategoryFilter() {{
            const categories = [...new Set(productsData.map(p => p.category))].sort();
            const select = document.getElementById('categoryFilter');
            select.innerHTML = '<option value="ALL">Todas las Categorías</option>';
            categories.forEach(cat => {{
                const opt = document.createElement('option');
                opt.value = cat;
                opt.textContent = cat;
                select.appendChild(opt);
            }});
        }}

        // Update Statistics Bar Below Filters
        function updateStatsSummary(data) {{
            const count = data.length;
            const units = data.reduce((sum, p) => sum + p.units, 0);
            const revenue = data.reduce((sum, p) => sum + p.revenue, 0);
            
            document.getElementById('statsCount').textContent = count;
            document.getElementById('statsUnits').textContent = units.toLocaleString('es-CL');
            document.getElementById('statsRevenue').textContent = revenue.toLocaleString('es-CL', {{style: 'currency', currency: 'CLP', maximumFractionDigits: 0}});
        }}

        // Render Table Data
        function renderProductTable(data) {{
            const tbody = document.getElementById('productTableBody');
            tbody.innerHTML = '';
            
            data.forEach((p, idx) => {{
                let badgeClass = 'badge-other';
                if (p.category === 'AÇAÍ PRIME') badgeClass = 'badge-acai';
                else if (p.category === 'AMERICAN PRIME BURGER') badgeClass = 'badge-burger';
                else if (p.category === 'Propina') badgeClass = 'badge-tip';
                else if (p.category === 'Extras') badgeClass = 'badge-extra';
                else if (p.category === 'Importe personalizado') badgeClass = 'badge-custom';
                
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td class="product-rank">#${{p.rank}}</td>
                    <td class="text-bold">${{p.name}}</td>
                    <td><span class="badge ${{badgeClass}}">${{p.category}}</span></td>
                    <td class="text-bold">${{p.units.toLocaleString('es-CL')}}</td>
                    <td class="text-bold" style="color: var(--primary);">${{p.revenue.toLocaleString('es-CL', {{style: 'currency', currency: 'CLP', maximumFractionDigits: 0}})}}</td>
                `;
                tbody.appendChild(tr);
            }});
        }}

        // NEW FILTER FUNCTION: Filtering products by query, category, min units, and min amount
        function filterProducts() {{
            const query = document.getElementById('productSearch').value.toLowerCase();
            const catFilter = document.getElementById('categoryFilter').value;
            const minUnits = parseFloat(document.getElementById('minUnits').value) || 0;
            const minAmount = parseFloat(document.getElementById('minAmount').value) || 0;
            
            const filtered = productsData.filter(p => {{
                const matchesSearch = p.name.toLowerCase().includes(query);
                const matchesCategory = (catFilter === 'ALL' || p.category === catFilter);
                const matchesUnits = p.units >= minUnits;
                const matchesAmount = p.revenue >= minAmount;
                return matchesSearch && matchesCategory && matchesUnits && matchesAmount;
            }});
            
            renderProductTable(filtered);
            updateStatsSummary(filtered);
        }}

        // Clear all filter values and reload the table
        function clearFilters() {{
            document.getElementById('productSearch').value = '';
            document.getElementById('categoryFilter').value = 'ALL';
            document.getElementById('minUnits').value = '';
            document.getElementById('minAmount').value = '';
            filterProducts();
        }}

        // Chat bot functions
        function askSuggested(text) {{
            document.getElementById('chatInput').value = text;
            sendChatMessage();
        }}

        function checkChatEnter(event) {{
            if (event.key === 'Enter') {{
                sendChatMessage();
            }}
        }}

        function formatMarkdown(text) {{
            if (!text) return "";
            let html = text;
            // Bold
            html = html.replace(/\\*\\*([^\\*]+)\\*\\*/g, '<strong>$1</strong>');
            // Bullet points
            html = html.replace(/^\\s*-\\s+(.+)/gm, '<li style="margin-left: 20px; margin-bottom: 4px;">$1</li>');
            // Paragraph breaks
            html = html.replace(/\\n\\n/g, '<br><br>');
            html = html.replace(/\\n/g, '<br>');
            return html;
        }}

        function sendChatMessage() {{
            const input = document.getElementById('chatInput');
            const text = input.value.trim();
            if (!text) return;
            
            // Clear input
            input.value = '';
            
            // Append user message
            appendMessage(text, 'user');
            
            // Show typing indicator
            const history = document.getElementById('chatHistory');
            const typingDiv = document.createElement('div');
            typingDiv.className = 'chat-message message-assistant';
            typingDiv.id = 'typingIndicator';
            typingDiv.innerHTML = '<em>Consultando asistente...</em>';
            history.appendChild(typingDiv);
            history.scrollTop = history.scrollHeight;
            
            // Obtener fecha y hora actual en la zona horaria del cliente en español
            const clientOptions = {{ weekday: 'long', year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' }};
            const clientTime = new Date().toLocaleString('es-CL', clientOptions);

            // Fetch API from Vercel Serverless Function
            fetch('/api/chat', {{
                method: 'POST',
                headers: {{
                    'Content-Type': 'application/json'
                }},
                body: JSON.stringify({{ 
                    message: text,
                    client_time: clientTime
                }})
            }})

            .then(res => {{
                if (!res.ok) throw new Error('API server returned error');
                return res.json();
            }})
            .then(data => {{
                const indicator = document.getElementById('typingIndicator');
                if (indicator) indicator.remove();
                
                // Fallback locally if GEMINI_API_KEY is not set on Vercel yet
                if (data.error_no_key) {{
                    console.log('Gemini API key missing on Vercel. Falling back to local engine...');
                    const localAnswer = getAssistantAnswer(text);
                    appendMessage(localAnswer, 'assistant');
                    return;
                }}
                
                if (data.response) {{
                    let msg = formatMarkdown(data.response);
                    if (data.query_executed) {{
                        msg += `<div style="margin-top: 10px; font-size: 11px; font-family: 'Courier New', Courier, monospace; color: var(--text-muted); opacity: 0.85; background-color: var(--primary-light); padding: 8px 12px; border-radius: var(--radius-sm); border-left: 3px solid var(--primary); word-break: break-all; line-height: 1.4;">🔍 <strong>Consulta SQL ejecutada:</strong><br>${{data.query_executed}}</div>`;
                    }}


                    appendMessage(msg, 'assistant');
                }} else {{
                    throw new Error('Empty response');
                }}

            }})
            .catch(err => {{
                console.log('Error calling Vercel chat API, falling back to local engine:', err);
                const indicator = document.getElementById('typingIndicator');
                if (indicator) indicator.remove();
                
                // Fallback to local mathematical engine
                const localAnswer = getAssistantAnswer(text);
                appendMessage(localAnswer, 'assistant');
            }});
        }}

        function appendMessage(htmlContent, sender) {{
            const history = document.getElementById('chatHistory');
            const msgDiv = document.createElement('div');
            msgDiv.className = `chat-message message-${{sender}}`;
            msgDiv.innerHTML = htmlContent;
            history.appendChild(msgDiv);
            history.scrollTop = history.scrollHeight;
        }}

        function getAssistantAnswer(query) {{
            query = query.toLowerCase().trim();
            
            // Normalize accents
            const normalized = query.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            
            // Mas vendido / estrella
            if (normalized.includes("mas vendido") || normalized.includes("estrella") || normalized.includes("mejor producto") || normalized.includes("top producto") || normalized.includes("lider")) {{
                const top = productsData[0];
                return `El producto estrella es <strong>${{top.name}}</strong> con <strong>${{top.units.toLocaleString('es-CL')}} unidades</strong> vendidas y una recaudación bruta de <strong>${{top.revenue.toLocaleString('es-CL', {{style: 'currency', currency: 'CLP', maximumFractionDigits: 0}})}}</strong>.`;
            }}

            // Promedio semanal
            if (normalized.includes("promedio") && (normalized.includes("semanal") || normalized.includes("semana"))) {{
                const numWeeks = weeklyData.length;
                if (numWeeks === 0) return "No hay suficientes datos semanales para calcular el promedio.";
                const totalWeeklyGross = weeklyData.reduce((sum, w) => sum + w.gross, 0);
                const totalWeeklyTx = weeklyData.reduce((sum, w) => sum + w.tx, 0);
                const avgWeeklyGross = totalWeeklyGross / numWeeks;
                const avgWeeklyTx = totalWeeklyTx / numWeeks;
                return `El promedio de venta semanal desde el inicio de operaciones es de <strong>$${{avgWeeklyGross.toLocaleString('es-CL', {{maximumFractionDigits: 0}})}} CLP</strong> por semana, con una media de <strong>${{avgWeeklyTx.toFixed(1)}} transacciones</strong> semanales (calculado sobre un total de ${{numWeeks}} semanas de ventas registradas).`;
            }}
            
            // Comisiones / fees
            if (normalized.includes("comision") || normalized.includes("comisiones") || normalized.includes("cobro") || normalized.includes("descuento forma de pago") || normalized.includes("porcentaje total")) {{
                const pct = ((totalFees / totalGross) * 100).toFixed(2);
                let detail = `Se pagó un total de <strong>$${{totalFees.toLocaleString('es-CL')}} CLP</strong> en comisiones bancarias (representa el <strong>${{pct}}%</strong> de las ventas brutas totales).<br><br><strong>Detalle por medio de pago:</strong><ul>`;
                paymentData.forEach(p => {{
                    if (p.fee > 0) {{
                        detail += `<li style="margin-bottom: 4px;"><strong>${{p.name}}</strong> (Tasa: ${{p.rate}}%): $${{p.fee.toLocaleString('es-CL')}} CLP de comisión (Recaudación neta: $${{p.net.toLocaleString('es-CL')}} CLP)</li>`;
                    }}
                }});
                detail += `</ul>`;
                return detail;
            }}
            
            // Neto / real
            if (normalized.includes("neto") || normalized.includes("ingreso real") || normalized.includes("ganancia real") || normalized.includes("cuanto recibi") || normalized.includes("ingreso neto")) {{
                return `El ingreso neto real (recaudación bruta menos comisiones) es de <strong>$${{totalNetReal.toLocaleString('es-CL')}} CLP</strong> (sobre una venta bruta total de <strong>$${{totalGross.toLocaleString('es-CL')}} CLP</strong>).`;
            }}
            
            // Bruto / total
            if (normalized.includes("bruto") || normalized.includes("venta total") || normalized.includes("recaudacion total") || (normalized.includes("cuanto") && normalized.includes("vendio") && !normalized.includes("neto") && !normalized.includes("capuccino") && !normalized.includes("burger") && !normalized.includes("agua") && !normalized.includes("fanta"))) {{
                return `La venta bruta total en el periodo registrado es de <strong>$${{totalGross.toLocaleString('es-CL')}} CLP</strong> a través de <strong>${{totalTx.toLocaleString('es-CL')}} transacciones</strong>.`;
            }}
            
            // Dia de la semana
            if (normalized.includes("dia") && (normalized.includes("venta") || normalized.includes("fuerte") || normalized.includes("mejor") || normalized.includes("semana"))) {{
                const sortedDays = [...dowData].sort((a, b) => b.gross - a.gross);
                const topDay = sortedDays[0];
                return `El día con mayor facturación es el <strong>${{topDay.day}}</strong> con <strong>$${{topDay.gross.toLocaleString('es-CL')}} CLP</strong> en ventas brutas, representando el <strong>${{((topDay.gross / totalGross) * 100).toFixed(1)}}%</strong> de la semana.`;
            }}
            
            // Transacciones / boletas
            if (normalized.includes("transaccion") || normalized.includes("boleta") || normalized.includes("ventas hechas") || normalized.includes("operaciones") || normalized.includes("tickets")) {{
                return `Se procesaron un total de <strong>${{totalTx.toLocaleString('es-CL')}} transacciones</strong> en tu punto de venta.`;
            }}
            
            // Consolidaciones / agrupamiento
            if (normalized.includes("consolid") || normalized.includes("agrup") || normalized.includes("unific") || normalized.includes("limpi")) {{
                return `Se aplicaron las siguientes consolidaciones para limpiar el catálogo de ventas:<br><br>
                1. <strong>Bebidas Fanta</strong>: Unificadas bajo <em>BEBIDA 350CC - FANTA ZERO POMELO</em>.<br>
                2. <strong>Cafetería</strong>: Cafés preparados con leche (*Capuccino, Latte, Cortado*) consolidados bajo <em>CAFETERÍA - CAFÉ CON LECHE</em>. Espresso, Espresso Doble y Americano permanecen separados.<br>
                3. <strong>Aguas</strong>: Divididas estrictamente en <em>CON GAS</em> y <em>SIN GAS</em>.<br>
                4. <strong>Hamburguesas</strong>: Agrupadas por nombre base (se removió el texto "(INCLUYE PAPAS FRITAS)").<br>
                5. <strong>Jugos</strong>: Consolidado total en la etiqueta <em>JUGOS NATURALES</em> para todos los sabores.`;
            }}
            
            // Specific product search
            let foundProduct = null;
            for (const p of productsData) {{
                if (normalized.includes(p.name.toLowerCase()) || p.name.toLowerCase().includes(normalized)) {{
                    foundProduct = p;
                    break;
                }}
            }}
            if (foundProduct) {{
                return `Encontré información de ventas para el producto <strong>${{foundProduct.name}}</strong>:<br>
                - **Categoría**: ${{foundProduct.category}}<br>
                - **Ranking de Ventas**: #${{foundProduct.rank}}<br>
                - **Unidades Vendidas**: ${{foundProduct.units.toLocaleString('es-CL')}} unidades<br>
                - **Recaudación Bruta**: $${{foundProduct.revenue.toLocaleString('es-CL')}} CLP.`;
            }}
            
            // Specific category search
            let foundCat = null;
            const uniqueCats = [...new Set(productsData.map(p => p.category))];
            for (const c of uniqueCats) {{
                if (normalized.includes(c.toLowerCase()) || c.toLowerCase().includes(normalized)) {{
                    foundCat = c;
                    break;
                }}
            }}
            if (foundCat) {{
                const catProducts = productsData.filter(p => p.category === foundCat);
                const units = catProducts.reduce((s, p) => s + p.units, 0);
                const gross = catProducts.reduce((s, p) => s + p.revenue, 0);
                return `Para la categoría <strong>${{foundCat}}</strong> registramos:<br>
                - **Productos distintos**: ${{catProducts.length}}<br>
                - **Total unidades vendidas**: ${{units.toLocaleString('es-CL')}} unidades<br>
                - **Total venta bruta**: $${{gross.toLocaleString('es-CL')}} CLP.`;
            }}

            // Fallback
            return `No he podido encontrar una respuesta matemática para esa consulta. Prueba preguntándome:<br><br>
            - <em>"¿Cuál es el producto estrella?"</em><br>
            - <em>"¿Cuánto fue la comisión de débito?"</em><br>
            - <em>"¿Qué día se vende más?"</em><br>
            - <em>"¿Cuánto vendió American Prime Burger?"</em><br>
            - <em>"¿Cuál es la venta neta real?"</em>`;
        }}

        // Sort Products
        function sortProducts(column) {{
            if (currentSort.column === column) {{
                currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
            }} else {{
                currentSort.column = column;
                currentSort.direction = 'desc';
            }}
            
            const sorted = [...productsData].sort((a, b) => {{
                let valA = a[column];
                let valB = b[column];
                if (currentSort.direction === 'asc') {{
                    return valA > valB ? 1 : -1;
                }} else {{
                    return valA < valB ? 1 : -1;
                }}
            }});
            
            // Rewrite original dataset's rank order based on new sorting
            sorted.forEach((p, idx) => p.rank = idx + 1);
            
            // Re-inject sorted data into the array
            for (let i = 0; i < productsData.length; i++) {{
                productsData[i] = sorted[i];
            }}
            
            filterProducts(); // Apply filters on top of the sorted dataset
        }}

        // Initialize Charts (Chart.js)
        function initCharts() {{
            // 1. Day of Week Chart
            const dowLabels = dowData.map(d => d.day);
            const dowGross = dowData.map(d => d.gross);
            
            new Chart(document.getElementById('dowChart'), {{
                type: 'bar',
                data: {{
                    labels: dowLabels,
                    datasets: [{{
                        label: 'Ventas Brutas ($)',
                        data: dowGross,
                        backgroundColor: dowLabels.map(day => 
                            (day === 'Sábado' || day === 'Domingo') ? '#e91e63' : '#5b1a8f'
                        ),
                        borderRadius: 6,
                        borderWidth: 0
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{ display: false }},
                        tooltip: {{
                            callbacks: {{
                                label: function(context) {{
                                    return 'Ventas: $' + context.raw.toLocaleString('es-CL');
                                }}
                            }}
                        }}
                    }},
                    scales: {{
                        y: {{
                            beginAtZero: true,
                            grid: {{ color: '#f1ecf6' }},
                            ticks: {{
                                callback: function(value) {{
                                    return '$' + (value/1e6) + 'M';
                                }}
                            }}
                        }},
                        x: {{ grid: {{ display: false }} }}
                    }}
                }}
            }});

            // 2. Category Doughnut Chart
            const catLabels = catData.map(c => c.category);
            const catGross = catData.map(c => c.gross);
            
            new Chart(document.getElementById('categoryChart'), {{
                type: 'doughnut',
                data: {{
                    labels: catLabels,
                    datasets: [{{
                        data: catGross,
                        backgroundColor: [
                            '#5E35B1', // Purple
                            '#E65100', // Orange
                            '#2E7D32', // Green
                            '#00838F', // Cyan
                            '#C2185B', // Pink
                            '#546E7A'  // Slate
                        ],
                        borderWidth: 2,
                        borderColor: '#ffffff'
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{
                            position: 'right',
                            labels: {{
                                font: {{ family: 'Inter', size: 12 }},
                                boxWidth: 15
                            }}
                        }},
                        tooltip: {{
                            callbacks: {{
                                label: function(context) {{
                                    const val = context.raw;
                                    const sum = context.dataset.data.reduce((a,b)=>a+b, 0);
                                    const pct = ((val/sum)*100).toFixed(1);
                                    return context.label + ': $' + val.toLocaleString('es-CL') + ' (' + pct + '%)';
                                }}
                            }}
                        }}
                    }}
                }}
            }});

            // 3. Weekly Trend Chart
            const weekLabels = weeklyData.map(w => w.week.replace('/2026', ''));
            const weeklyGross = weeklyData.map(w => w.gross);
            
            new Chart(document.getElementById('weeklyChart'), {{
                type: 'line',
                data: {{
                    labels: weekLabels,
                    datasets: [{{
                        label: 'Ventas Semanales ($)',
                        data: weeklyGross,
                        borderColor: '#5b1a8f',
                        backgroundColor: 'rgba(91, 26, 143, 0.05)',
                        fill: true,
                        tension: 0.3,
                        borderWidth: 3,
                        pointBackgroundColor: '#e91e63',
                        pointBorderColor: '#ffffff',
                        pointHoverRadius: 6
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{ display: false }},
                        tooltip: {{
                            callbacks: {{
                                label: function(context) {{
                                    return 'Semana: $' + context.raw.toLocaleString('es-CL');
                                }}
                            }}
                        }}
                    }},
                    scales: {{
                        y: {{
                            beginAtZero: true,
                            grid: {{ color: '#f1ecf6' }},
                            ticks: {{
                                callback: function(value) {{
                                    return '$' + (value/1e6) + 'M';
                                }}
                            }}
                        }},
                        x: {{ grid: {{ display: false }} }}
                    }}
                }}
            }});

            // 4. Payment Fee Doughnut Chart
            const pmLabels = paymentData.filter(p => p.fee > 0).map(p => p.name);
            const pmFees = paymentData.filter(p => p.fee > 0).map(p => p.fee);

            new Chart(document.getElementById('feesChart'), {{
                type: 'doughnut',
                data: {{
                    labels: pmLabels,
                    datasets: [{{
                        data: pmFees,
                        backgroundColor: [
                            '#3f51b5', // Indigo
                            '#ff9800', // Orange
                            '#00bcd4', // Teal
                            '#f44336'  // Red
                        ],
                        borderWidth: 2,
                        borderColor: '#ffffff'
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{
                            position: 'bottom',
                            labels: {{ font: {{ family: 'Inter', size: 12 }}, boxWidth: 15 }}
                        }},
                        tooltip: {{
                            callbacks: {{
                                label: function(context) {{
                                    const val = context.raw;
                                    const sum = context.dataset.data.reduce((a,b)=>a+b, 0);
                                    const pct = ((val/sum)*100).toFixed(1);
                                    return context.label + ': $' + val.toLocaleString('es-CL') + ' (' + pct + '%)';
                                }}
                            }}
                        }}
                    }}
                }}
            }});
        }}
    </script>
</body>
</html>
"""

# Save output to workspace
output_file = "index.html"
with open(output_file, 'w', encoding='utf-8') as f:
    f.write(html_template)

print("SUCCESS: HTML Report successfully updated with interactive filters.")
