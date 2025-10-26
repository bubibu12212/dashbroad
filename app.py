import os
import glob 
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
from datetime import datetime
import openpyxl
from werkzeug.utils import secure_filename
import locale
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import json
import plotly

# --- Inisialisasi & Konfigurasi ---
app = Flask(__name__)
app.secret_key = "denso_secret_key_final_version"
app.config['UPLOAD_FOLDER'] = os.path.dirname(os.path.abspath(__file__))

# --- Mengatur Locale ke Bahasa Inggris ---
try:
    locale.setlocale(locale.LC_TIME, 'en_US.UTF-8')
except:
    locale.setlocale(locale.LC_ALL, '')

# --- Fungsi Bantuan (Revisi) ---
DATA_FOLDER_PATH = r"C:\Users\ASUS\Downloads\file"

def get_data_file_path(year):
    """Membuat path file lengkap ke file data Excel berdasarkan TAHUN."""
    file_name = f"data_{year}.xlsx"
    os.makedirs(DATA_FOLDER_PATH, exist_ok=True)
    return os.path.join(DATA_FOLDER_PATH, file_name)

def load_data(file_path):
    """Memuat data dari SATU path file Excel yang spesifik."""
    try:
        df = pd.read_excel(file_path)
        if df.empty:
            return pd.DataFrame()
        # Menambahkan 'row_id' unik berdasarkan hash dari file + index
        # Ini penting agar ID tidak bentrok antar file
        df.reset_index(inplace=True)
        df['row_id'] = df.apply(lambda row: hash(f"{os.path.basename(file_path)}-{row['index']}"), axis=1)
        df.drop(columns=['index'], inplace=True, errors='ignore')
        return df
    except FileNotFoundError:
        return pd.DataFrame()
    except Exception as e:
        print(f"Error loading data from {file_path}: {e}")
        return pd.DataFrame()

# --- PERUBAHAN BARU: Fungsi untuk memuat SEMUA data dari SEMUA file ---
def load_all_data():
    """Mencari semua file data_YYYY.xlsx, memuat, dan menggabungkannya."""
    all_files = glob.glob(os.path.join(DATA_FOLDER_PATH, "data_*.xlsx"))
    
    if not all_files:
        return pd.DataFrame() # Kembalikan DataFrame kosong jika tidak ada file sama sekali

    df_list = []
    for file_path in all_files:
        df_list.append(load_data(file_path))
    
    # Gabungkan semua DataFrame menjadi satu
    if not df_list:
        return pd.DataFrame()
        
    combined_df = pd.concat(df_list, ignore_index=True)
    return combined_df

def save_data(df, file_path):
    """Menyimpan seluruh DataFrame ke path file Excel yang spesifik."""
    try:
        # Selalu pastikan kolom 'row_id' tidak ikut tersimpan ke Excel
        if 'row_id' in df.columns:
            df_to_save = df.drop(columns=['row_id'])
        else:
            df_to_save = df
        df_to_save.to_excel(file_path, index=False)
        return True
    except Exception as e:
        flash(f"Failed to save Excel file to {file_path}. Error: {e}", "danger")
        return False

# --- Fungsi Grafik (Tidak Ada Perubahan) ---
def create_performance_chart(df):
    #Pastikan df tidak kosong sebelum membuat chart
    if df.empty or 'CLOSING MONTH' not in df.columns:
        return json.dumps({}) # Kembalikan chart kosong
    df['CLOSING MONTH'] = pd.to_datetime(df['CLOSING MONTH'])
    df = df.sort_values('CLOSING MONTH')
    month_labels = df['CLOSING MONTH'].dt.strftime('%B %Y')
    fig = go.Figure()
    fig.add_trace(go.Bar(x=month_labels, y=df['TOTAL DELIVERY ITEM'], name='Total Delivery (Bar)', marker_color='#C50000'))
    fig.add_trace(go.Bar(x=month_labels, y=df['ON TIME'], name='On Time (Bar)', marker_color='#28a745'))
    fig.add_trace(go.Scatter(x=month_labels, y=df['TOTAL DELIVERY ITEM'], name='Total Delivery (Line)', mode='lines+markers', line=dict(color="#ff6600", width=3, dash='dash')))
    fig.add_trace(go.Scatter(x=month_labels, y=df['ON TIME'], name='On Time (Line)', mode='lines+markers', line=dict(color='#72e08a', width=3, dash='dash')))
    fig.update_layout(title_text="<b>Delivery and On-Time Performance (All Years)</b>", title_font_color='black', barmode='group', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(color='black')), plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF", font=dict(family="Poppins, sans-serif", color='black'), hovermode='x unified', height=500, xaxis=dict(showgrid=True, gridcolor="#000000", tickfont=dict(color='black')), yaxis=dict(title_text="Jumlah (pcs)", gridcolor='#495057', tickfont=dict(color='black'), title_font_color='black'))
    return json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)

def create_purchasing_chart(df):
    # --- PERBAIKAN KECIL ---: Pastikan df tidak kosong
    if df.empty or 'CLOSING MONTH' not in df.columns:
        return json.dumps({})
    df['CLOSING MONTH'] = pd.to_datetime(df['CLOSING MONTH'])
    df = df.sort_values('CLOSING MONTH')
    month_labels = df['CLOSING MONTH'].dt.strftime('%B %Y')
    fig = go.Figure(go.Scatter(x=month_labels, y=df['Purchase Amount'], name='Purchasing Amount', mode='lines+markers', line=dict(color='#0d6efd', width=4, shape='spline'), fill='tozeroy', fillcolor='rgba(13, 110, 253, 0.2)'))
    fig.update_layout(title_text="<b>Monthly Purchasing Amount (All Years)</b>", title_font_color='black', plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF", font=dict(family="Poppins, sans-serif", color='black'), yaxis_title="<b>Purchase Amount (Rp)</b>", yaxis_title_font_color='black', hovermode='x unified', height=500)
    fig.update_yaxes(gridcolor="#000000", tickfont=dict(color='black'))
    return json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)


# --- Rute Aplikasi ---

# --- PERUBAHAN: Gunakan load_all_data() ---
@app.route('/')
def index():
    df = load_all_data() # Memuat SEMUA data
    suppliers = []
    if not df.empty and 'SUPPLIER NAME' in df.columns:
        suppliers = sorted(df['SUPPLIER NAME'].unique().tolist())
    return render_template('index.html', suppliers=suppliers)

@app.route('/search', methods=['POST'])
def search():
    supplier_name = request.form.get('supplier_name')
    return redirect(url_for('dashboard', supplier_name=supplier_name))

# --- PERUBAHAN: Gunakan load_all_data() ---
@app.route('/dashboard/<supplier_name>')
def dashboard(supplier_name):
    df = load_all_data() # Memuat SEMUA data

    if df.empty: 
        flash('No data files found in the data folder. Please add data to begin.', 'info')
        return redirect(url_for('index'))
    
    df['CLOSING MONTH'] = pd.to_datetime(df['CLOSING MONTH'])
    supplier_data = df[df['SUPPLIER NAME'] == supplier_name].sort_values('CLOSING MONTH')
    
    if supplier_data.empty:
        flash(f'Data for "{supplier_name}" not found.', 'warning')
        return redirect(url_for('index'))
        
    total_purchase = supplier_data['Purchase Amount'].sum()
    avg_achievement = supplier_data['ACHIEVEMENT'].mean() * 100
    total_delivery = supplier_data['TOTAL DELIVERY ITEM'].sum()
    kpi = {'total_purchase': f"Rp {total_purchase:,.0f}".replace(',', '.'), 'avg_achievement': f"{avg_achievement:.2f}%", 'total_delivery': f"{int(total_delivery)} pcs"}
    
    # Chart sekarang akan menampilkan data lintas tahun
    chart1_json = create_performance_chart(supplier_data)
    chart2_json = create_purchasing_chart(supplier_data)
    
    table_data_display = supplier_data.sort_values('CLOSING MONTH', ascending=False).copy()
    table_data_display['ACHIEVEMENT_FORMATTED'] = (table_data_display['ACHIEVEMENT'] * 100).map('{:.2f}%'.format)
    table_data_display['TARGET_FORMATTED'] = '90%'
    table_data_display['PURCHASE_FORMATTED'] = table_data_display['Purchase Amount'].map('Rp {:,.0f}'.format).str.replace(',', '.', regex=False)
    table_data = table_data_display.to_dict(orient='records')
    return render_template('dashboard.html', supplier_name=supplier_name, chart1_json=chart1_json, chart2_json=chart2_json, kpi=kpi, table_data=table_data)

# --- Rute untuk Aksi CRUD & Upload ---

# --- PERUBAHAN: Logika 'update' sudah benar (dari respons sebelumnya) ---
@app.route('/update', methods=['GET', 'POST'])
def update():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part selected.', 'warning')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No file selected.', 'warning')
            return redirect(request.url)
        
        if file and file.filename.endswith('.xlsx'):
            try:
                df_upload = pd.read_excel(file)
                # Pastikan kolom tanggal dikenali
                df_upload['CLOSING MONTH'] = pd.to_datetime(df_upload['CLOSING MONTH'])
                upload_years = df_upload['CLOSING MONTH'].dt.year.unique()
                processed_files = []
                
                for year in upload_years:
                    file_path = get_data_file_path(year)
                    df_existing = load_data(file_path) # Memuat data LAMA dari file spesifik
                    # Hapus 'row_id' lama sebelum digabung agar tidak bentrok
                    if 'row_id' in df_existing.columns:
                         df_existing = df_existing.drop(columns=['row_id'])

                    df_new_for_year = df_upload[df_upload['CLOSING MONTH'].dt.year == year]
                    df_combined = pd.concat([df_existing, df_new_for_year], ignore_index=True)
                    df_combined.drop_duplicates(subset=['SUPPLIER NAME', 'CLOSING MONTH'], keep='last', inplace=True)
                    
                    save_data(df_combined, file_path) # Simpan kembali ke file spesifik
                    processed_files.append(os.path.basename(file_path))

                flash(f'Data successfully processed for files: {", ".join(processed_files)}!', 'success')
                return redirect(url_for('index'))
            except Exception as e:
                flash(f"Failed to process file. Error: {e}", 'danger')
    return render_template('update.html')

# --- PERUBAHAN BESAR DI FUNGSI INI ---
@app.route('/supplier/add', methods=['POST'])
def add_supplier():
    form = request.form
    month_input = datetime.strptime(form['month'], '%Y-%m')
    target_year = month_input.year
    file_path = get_data_file_path(target_year)
    
    df = load_data(file_path) # Memuat data dari file spesifik
    # Hapus 'row_id' lama sebelum digabung
    if 'row_id' in df.columns:
        df = df.drop(columns=['row_id'])

    supplier_name = form['supplier_name']
    if not df.empty and supplier_name in df['SUPPLIER NAME'].unique():
        flash(f'Supplier "{supplier_name}" already exists in data for year {target_year}.', 'warning')
        return redirect(url_for('index'))
    
    # --- LOGIKA BARU DIMULAI ---
    total_delivery_val = int(float(form['total_delivery']))
    on_time_val = int(float(form['on_time']))
    
    achievement_val = 0.0
    if total_delivery_val > 0: # Hindari pembagian dengan nol
        # Perhitungan achievement sebagai rasio (misal: 0.95 untuk 95%)
        achievement_val = (on_time_val / total_delivery_val) 
    # --- LOGIKA BARU SELESAI ---

    new_row_data = {
        'CLOSING MONTH': month_input, 'SUPPLIER NAME': supplier_name,
        'TOTAL DELIVERY ITEM': total_delivery_val, # Menggunakan variabel
        'ON TIME': on_time_val, # Menggunakan variabel
        'MINUS': int(float(form['minus'])),
        'TARGET DELIVERY': float(form['target_delivery']) / 100,
        'ACHIEVEMENT': achievement_val, # Menggunakan variabel hasil perhitungan
        'Purchase Amount': float(form['purchase_amount']),
        'ITEM DELAY': form.get('item_delay', 0)
    }
    df_new = pd.concat([df, pd.DataFrame([new_row_data])], ignore_index=True)
    
    if save_data(df_new, file_path): 
        flash(f'New supplier "{supplier_name}" added successfully to data for {target_year}!', 'success')
    return redirect(url_for('index'))

# --- PERUBAHAN BESAR DI FUNGSI INI ---
@app.route('/data/add/<supplier_name>', methods=['POST'])
def add_monthly_entry(supplier_name):
    form = request.form
    form_month = datetime.strptime(form['month'], '%Y-%m')
    target_year = form_month.year
    file_path = get_data_file_path(target_year)
    df = load_data(file_path) # Memuat data dari file spesifik
    
    data_exists = False
    if not df.empty:
        df['CLOSING MONTH'] = pd.to_datetime(df['CLOSING MONTH'])
        mask = (df['SUPPLIER NAME'] == supplier_name) & (df['CLOSING MONTH'].dt.year == form_month.year) & (df['CLOSING MONTH'].dt.month == form_month.month)
        if not df[mask].empty:
            data_exists = True

    if data_exists:
        flash(f'Data for {form_month.strftime("%B %Y")} already exists in file data_{target_year}.xlsx. Use "Edit" button.', 'warning')
    else:
        # Hapus 'row_id' lama sebelum digabung
        if 'row_id' in df.columns:
            df = df.drop(columns=['row_id'])
        
        # --- LOGIKA BARU DIMULAI ---
        total_delivery_val = int(float(form['total_delivery']))
        on_time_val = int(float(form['on_time']))
        
        achievement_val = 0.0
        if total_delivery_val > 0: # Hindari pembagian dengan nol
            achievement_val = (on_time_val / total_delivery_val) 
        # --- LOGIKA BARU SELESAI ---

        new_row_data = {
            'CLOSING MONTH': form_month, 'SUPPLIER NAME': supplier_name,
            'TOTAL DELIVERY ITEM': total_delivery_val, # Menggunakan variabel
            'ON TIME': on_time_val, # Menggunakan variabel
            'MINUS': int(float(form['minus'])),
            'TARGET DELIVERY': float(form['target_delivery']) / 100,
            'ACHIEVEMENT': achievement_val, # Menggunakan variabel hasil perhitungan
            'Purchase Amount': float(form['purchase_amount']),
            'ITEM DELAY': form.get('item_delay', 0)
        }
        df_new = pd.concat([df, pd.DataFrame([new_row_data])], ignore_index=True)
        if save_data(df_new, file_path): 
            flash(f'Data for {form_month.strftime("%B %Y")} added successfully to data_{target_year}.xlsx!', 'success')
            
    return redirect(url_for('dashboard', supplier_name=supplier_name))

# --- PERUBAHAN BESAR DI FUNGSI INI ---
@app.route('/data/edit/<int:row_id>', methods=['POST'])
def edit_entry(row_id):
    form = request.form
    form_month = datetime.strptime(form['month'], '%Y-%m')
    target_year = form_month.year
    file_path = get_data_file_path(target_year)
    df = load_data(file_path) # Memuat data dari file spesifik
    
    supplier_name = ""
    if not df.empty and row_id in df['row_id'].values:
        supplier_name = df.loc[df['row_id'] == row_id, 'SUPPLIER NAME'].iloc[0]
        
        # --- LOGIKA BARU DIMULAI ---
        total_delivery_val = int(float(form['total_delivery']))
        on_time_val = int(float(form['on_time']))
        
        achievement_val = 0.0
        if total_delivery_val > 0: # Hindari pembagian dengan nol
            achievement_val = (on_time_val / total_delivery_val) 
        # --- LOGIKA BARU SELESAI ---

        # Update data di DataFrame
        df.loc[df['row_id'] == row_id, 'CLOSING MONTH'] = form_month
        df.loc[df['row_id'] == row_id, 'TOTAL DELIVERY ITEM'] = total_delivery_val # Menggunakan variabel
        df.loc[df['row_id'] == row_id, 'ON TIME'] = on_time_val # Menggunakan variabel
        df.loc[df['row_id'] == row_id, 'MINUS'] = int(float(form['minus']))
        df.loc[df['row_id'] == row_id, 'TARGET DELIVERY'] = float(form['target_delivery']) / 100
        df.loc[df['row_id'] == row_id, 'ACHIEVEMENT'] = achievement_val # Menggunakan variabel hasil perhitungan
        df.loc[df['row_id'] == row_id, 'Purchase Amount'] = float(form['purchase_amount'])
        
        if save_data(df, file_path): 
            flash('Data updated successfully!', 'success')
    else: 
        flash(f'Data with ID {row_id} not found in file for year {target_year}.', 'danger')
        if not supplier_name:
            return redirect(url_for('index'))
        
    return redirect(url_for('dashboard', supplier_name=supplier_name))

# --- PERUBAHAN BESAR: Logika 'delete_entry' yang lebih cerdas ---
@app.route('/data/delete/<int:row_id>', methods=['POST'])
def delete_entry(row_id):
    # 1. Muat SEMUA data untuk menemukan data yang akan dihapus
    df_all = load_all_data()
    if df_all.empty:
        flash('Data not found.', 'danger')
        return redirect(url_for('index'))

    supplier_name = ""
    target_year = None
    
    # 2. Cari baris yang sesuai dengan row_id
    target_row = df_all[df_all['row_id'] == row_id]
    
    if not target_row.empty:
        # 3. Dapatkan informasi dari baris tersebut
        supplier_name = target_row.iloc[0]['SUPPLIER NAME']
        closing_month = pd.to_datetime(target_row.iloc[0]['CLOSING MONTH'])
        target_year = closing_month.year
        
        # 4. Muat HANYA file spesifik tahun tersebut
        file_path = get_data_file_path(target_year)
        df_single_file = load_data(file_path) # Data dari file (misal: data_2026.xlsx)
        
        # 5. Hapus baris dari DataFrame file spesifik itu
        df_single_file = df_single_file[df_single_file['row_id'] != row_id]
        
        # 6. Simpan kembali file spesifik tersebut
        if save_data(df_single_file, file_path): 
            flash(f'Data for {closing_month.strftime("%B %Y")} deleted successfully from file data_{target_year}.xlsx!', 'success')
        else:
            flash('Failed to delete data.', 'danger')
    else: 
        flash(f'Data with ID {row_id} not found.', 'danger')
        return redirect(url_for('index'))

    return redirect(url_for('dashboard', supplier_name=supplier_name))

# --- PERUBAHAN: Gunakan load_all_data() ---
@app.route('/check_supplier', methods=['POST'])
def check_supplier():
    """Mengecek apakah nama supplier sudah ada di SEMUA data."""
    df = load_all_data()
    name_to_check = request.form.get('supplier_name', '').strip().lower()

    if df.empty or 'SUPPLIER NAME' not in df.columns:
        return jsonify({'exists': False})

    is_present = df['SUPPLIER NAME'].str.strip().str.lower().eq(name_to_check).any()
    return jsonify({'exists': bool(is_present)})

# --- PERUBAHAN: Gunakan load_all_data() ---
@app.context_processor
def inject_suppliers():
    """Membuat daftar SEMUA supplier dari SEMUA tahun tersedia di semua template."""
    df = load_all_data()
    supplier_list = []
    if not df.empty and 'SUPPLIER NAME' in df.columns:
        supplier_list = sorted(df['SUPPLIER NAME'].unique().tolist())
    return dict(all_suppliers=supplier_list)

# ...existing code...
if __name__ == '__main__':
    app.run(debug=True, host='127.0.0.1', port=5001)