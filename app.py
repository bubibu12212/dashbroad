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

# --- Fungsi Bantuan ---
DATA_FOLDER_PATH = r"C:\Users\ASUS\Downloads\file" # Sesuaikan path ini

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
        df.reset_index(inplace=True)
        df['row_id'] = df.apply(lambda row: hash(f"{os.path.basename(file_path)}-{row['index']}"), axis=1)
        df.drop(columns=['index'], inplace=True, errors='ignore')
        return df
    except FileNotFoundError:
        return pd.DataFrame()
    except Exception as e:
        print(f"Error loading data from {file_path}: {e}")
        return pd.DataFrame()

def load_all_data():
    """Mencari semua file data_YYYY.xlsx, memuat, dan menggabungkannya."""
    all_files = glob.glob(os.path.join(DATA_FOLDER_PATH, "data_*.xlsx"))
    if not all_files:
        return pd.DataFrame()
    df_list = [load_data(fp) for fp in all_files]
    if not df_list:
        return pd.DataFrame()
    return pd.concat(df_list, ignore_index=True)

def save_data(df, file_path):
    """Menyimpan seluruh DataFrame ke path file Excel yang spesifik."""
    try:
        df_to_save = df.drop(columns=['row_id'], errors='ignore')
        df_to_save.to_excel(file_path, index=False)
        return True
    except Exception as e:
        flash(f"Failed to save Excel file to {file_path}. Error: {e}", "danger")
        return False

# --- Fungsi Grafik ---
# (Tidak ada perubahan di fungsi grafik)
def create_performance_chart(df):
    if df.empty or 'CLOSING MONTH' not in df.columns: return json.dumps({})
    df['CLOSING MONTH'] = pd.to_datetime(df['CLOSING MONTH'])
    df = df.sort_values('CLOSING MONTH')
    month_labels = df['CLOSING MONTH'].dt.strftime('%B %Y')
    fig = go.Figure()
    fig.add_trace(go.Bar(x=month_labels, y=df['TOTAL DELIVERY ITEM'], name='Total Delivery (Bar)', marker_color='#C50000'))
    fig.add_trace(go.Bar(x=month_labels, y=df['ON TIME'], name='On Time (Bar)', marker_color='#28a745'))
    fig.add_trace(go.Scatter(x=month_labels, y=df['TOTAL DELIVERY ITEM'], name='Total Delivery (Line)', mode='lines+markers', line=dict(color="#ff6600", width=3, dash='dash')))
    fig.add_trace(go.Scatter(x=month_labels, y=df['ON TIME'], name='On Time (Line)', mode='lines+markers', line=dict(color='#72e08a', width=3, dash='dash')))
    fig.update_layout(title_text="<b>Delivery and On-Time Performance (All Years)</b>", title_font_color='black', barmode='group', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(color='black')), plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF", font=dict(family="Poppins, sans-serif", color='black'), hovermode='x unified', height=500, xaxis=dict(showgrid=True, gridcolor="#DDDDDD", tickfont=dict(color='black')), yaxis=dict(title_text="Jumlah (pcs)", gridcolor='#DDDDDD', tickfont=dict(color='black'), title_font_color='black')) # Grid color softer
    return json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)

def create_purchasing_chart(df):
    if df.empty or 'CLOSING MONTH' not in df.columns: return json.dumps({})
    df['CLOSING MONTH'] = pd.to_datetime(df['CLOSING MONTH'])
    df = df.sort_values('CLOSING MONTH')
    month_labels = df['CLOSING MONTH'].dt.strftime('%B %Y')
    fig = go.Figure(go.Scatter(x=month_labels, y=df['Purchase Amount'], name='Purchasing Amount', mode='lines+markers', line=dict(color='#0d6efd', width=4, shape='spline'), fill='tozeroy', fillcolor='rgba(13, 110, 253, 0.2)'))
    fig.update_layout(title_text="<b>Monthly Purchasing Amount (All Years)</b>", title_font_color='black', plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF", font=dict(family="Poppins, sans-serif", color='black'), yaxis_title="<b>Purchase Amount (Rp)</b>", yaxis_title_font_color='black', hovermode='x unified', height=500)
    fig.update_yaxes(gridcolor="#DDDDDD", tickfont=dict(color='black')) # Grid color softer
    return json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)

# --- Rute Aplikasi ---
@app.route('/')
def index():
    df = load_all_data()
    suppliers = []
    if not df.empty and 'SUPPLIER NAME' in df.columns:
        suppliers = sorted(df['SUPPLIER NAME'].unique().tolist())
    return render_template('index.html', suppliers=suppliers)

@app.route('/search', methods=['POST'])
def search():
    supplier_name = request.form.get('supplier_name')
    return redirect(url_for('dashboard', supplier_name=supplier_name))

@app.route('/dashboard/<supplier_name>')
def dashboard(supplier_name):
    df = load_all_data()
    if df.empty: 
        flash('No data files found. Please add data.', 'info')
        return redirect(url_for('index'))
    df['CLOSING MONTH'] = pd.to_datetime(df['CLOSING MONTH'])
    supplier_data = df[df['SUPPLIER NAME'] == supplier_name].sort_values('CLOSING MONTH')
    if supplier_data.empty:
        flash(f'Data for "{supplier_name}" not found.', 'warning')
        return redirect(url_for('index'))
    total_purchase = supplier_data['Purchase Amount'].sum()
    avg_achievement = supplier_data['ACHIEVEMENT'].mean() * 100
    total_delivery = supplier_data['TOTAL DELIVERY ITEM'].sum()
    kpi = {'total_purchase': f"Rp {total_purchase:,.0f}".replace(',', '.'), 
           'avg_achievement': f"{avg_achievement:.2f}%" if not pd.isna(avg_achievement) else "N/A", # Handle NaN
           'total_delivery': f"{int(total_delivery)} pcs"}
    chart1_json = create_performance_chart(supplier_data)
    chart2_json = create_purchasing_chart(supplier_data)
    table_data = supplier_data.sort_values('CLOSING MONTH', ascending=False).to_dict(orient='records')
    return render_template('dashboard.html', supplier_name=supplier_name, chart1_json=chart1_json, chart2_json=chart2_json, kpi=kpi, table_data=table_data)

@app.route('/update', methods=['GET', 'POST'])
def update():
    if request.method == 'POST':
        # ... (logika upload file tetap sama) ...
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
                if 'CLOSING MONTH' not in df_upload.columns or 'SUPPLIER NAME' not in df_upload.columns:
                     flash('Uploaded file must contain "CLOSING MONTH" and "SUPPLIER NAME" columns.', 'danger')
                     return redirect(request.url)

                df_upload['CLOSING MONTH'] = pd.to_datetime(df_upload['CLOSING MONTH'])
                upload_years = df_upload['CLOSING MONTH'].dt.year.unique()
                processed_files = []
                
                for year in upload_years:
                    file_path = get_data_file_path(year)
                    df_existing = load_data(file_path) 
                    if 'row_id' in df_existing.columns:
                         df_existing = df_existing.drop(columns=['row_id'])

                    df_new_for_year = df_upload[df_upload['CLOSING MONTH'].dt.year == year]
                    # Penting: Pastikan kolom tanggal di df_existing juga datetime
                    if not df_existing.empty and 'CLOSING MONTH' in df_existing.columns:
                       df_existing['CLOSING MONTH'] = pd.to_datetime(df_existing['CLOSING MONTH'])

                    df_combined = pd.concat([df_existing, df_new_for_year], ignore_index=True)
                    df_combined.drop_duplicates(subset=['SUPPLIER NAME', 'CLOSING MONTH'], keep='last', inplace=True)
                    
                    if save_data(df_combined, file_path):
                         processed_files.append(os.path.basename(file_path))

                if processed_files:
                    flash(f'Data successfully processed for files: {", ".join(processed_files)}!', 'success')
                else:
                    flash('No data was processed. Check file format or content.', 'warning')
                return redirect(url_for('index'))
            except Exception as e:
                flash(f"Failed to process file. Error: {e}", 'danger')
    return render_template('update.html')

# --- FUNGSI ADD SUPPLIER ---
@app.route('/supplier/add', methods=['POST'])
def add_supplier():
    form = request.form
    try:
        month_input = datetime.strptime(form['month'], '%Y-%m')
        target_year = month_input.year
        file_path = get_data_file_path(target_year)
        
        df = load_data(file_path)
        if 'row_id' in df.columns: df = df.drop(columns=['row_id'])

        supplier_name = form['supplier_name'].strip() # Trim whitespace
        if not supplier_name:
             flash('Supplier name cannot be empty.', 'warning')
             return redirect(url_for('index'))

        # Cek duplikat (case-insensitive)
        if not df.empty and supplier_name.lower() in df['SUPPLIER NAME'].str.lower().unique():
            flash(f'Supplier "{supplier_name}" already exists for year {target_year}.', 'warning')
            return redirect(url_for('index'))
            
        total_delivery_val = int(float(form['total_delivery']))
        on_time_val = int(float(form['on_time']))
        achievement_val = (on_time_val / total_delivery_val) if total_delivery_val > 0 else (1.0 if on_time_val == 0 else 0.0)

        new_row_data = {
            'CLOSING MONTH': month_input, 
            'SUPPLIER NAME': supplier_name,
            'TOTAL DELIVERY ITEM': total_delivery_val, 
            'ON TIME': on_time_val,
            'MINUS': int(float(form['minus'])),
            # --- PERUBAHAN DI SINI ---
            'TARGET DELIVERY': 0.9, # Nilai tetap 90%
            'ACHIEVEMENT': achievement_val,
            'Purchase Amount': float(form['purchase_amount']),
            # Handle potential missing 'ITEM DELAY' if your Excel doesn't always have it
            'ITEM DELAY': df['ITEM DELAY'].iloc[0] if 'ITEM DELAY' in df.columns and not df.empty else 0 
        }
        # Tambah kolom lain jika ada, dengan nilai default atau NaN
        # Pastikan kolom baru konsisten dengan file Excel yang ada
        expected_cols = df.columns.tolist() if not df.empty else list(new_row_data.keys())
        for col in expected_cols:
             if col not in new_row_data:
                  new_row_data[col] = pd.NA # Atau 0 atau '' sesuai tipe data kolom

        # Buat DataFrame baru dengan urutan kolom yang benar
        df_new_row = pd.DataFrame([new_row_data], columns=expected_cols)
        
        df_final = pd.concat([df, df_new_row], ignore_index=True)
        
        if save_data(df_final, file_path): 
            flash(f'New supplier "{supplier_name}" added successfully for {target_year}!', 'success')
        return redirect(url_for('index'))

    except ValueError:
        flash('Invalid number format entered.', 'danger')
        return redirect(url_for('index'))
    except Exception as e:
        flash(f'An error occurred: {e}', 'danger')
        return redirect(url_for('index'))


# --- FUNGSI ADD MONTHLY ENTRY ---
@app.route('/data/add/<supplier_name>', methods=['POST'])
def add_monthly_entry(supplier_name):
    form = request.form
    try:
        form_month = datetime.strptime(form['month'], '%Y-%m')
        target_year = form_month.year
        file_path = get_data_file_path(target_year)
        df = load_data(file_path)
        
        data_exists = False
        if not df.empty:
            df['CLOSING MONTH'] = pd.to_datetime(df['CLOSING MONTH'])
            mask = (df['SUPPLIER NAME'].str.lower() == supplier_name.lower()) & \
                   (df['CLOSING MONTH'].dt.year == form_month.year) & \
                   (df['CLOSING MONTH'].dt.month == form_month.month)
            if df[mask].any().any(): # Check if any row matches
                data_exists = True

        if data_exists:
            flash(f'Data for {form_month.strftime("%B %Y")} already exists in data_{target_year}.xlsx. Use "Edit".', 'warning')
        else:
            if 'row_id' in df.columns: df = df.drop(columns=['row_id'])
            
            total_delivery_val = int(float(form['total_delivery']))
            on_time_val = int(float(form['on_time']))
            achievement_val = (on_time_val / total_delivery_val) if total_delivery_val > 0 else (1.0 if on_time_val == 0 else 0.0)

            new_row_data = {
                'CLOSING MONTH': form_month, 
                'SUPPLIER NAME': supplier_name, # Gunakan nama dari URL
                'TOTAL DELIVERY ITEM': total_delivery_val, 
                'ON TIME': on_time_val,
                'MINUS': int(float(form['minus'])),
                # --- PERUBAHAN DI SINI ---
                'TARGET DELIVERY': 0.9, # Nilai tetap 90%
                'ACHIEVEMENT': achievement_val,
                'Purchase Amount': float(form['purchase_amount']),
                # Handle potential missing 'ITEM DELAY' 
                'ITEM DELAY': df['ITEM DELAY'].iloc[0] if 'ITEM DELAY' in df.columns and not df.empty else 0
            }
            # Tambah kolom lain jika ada
            expected_cols = df.columns.tolist() if not df.empty else list(new_row_data.keys())
            for col in expected_cols:
                 if col not in new_row_data:
                      new_row_data[col] = pd.NA 

            df_new_row = pd.DataFrame([new_row_data], columns=expected_cols)
            df_final = pd.concat([df, df_new_row], ignore_index=True)

            if save_data(df_final, file_path): 
                flash(f'Data for {form_month.strftime("%B %Y")} added successfully to data_{target_year}.xlsx!', 'success')
                
    except ValueError:
        flash('Invalid number format entered.', 'danger')
    except Exception as e:
        flash(f'An error occurred: {e}', 'danger')

    return redirect(url_for('dashboard', supplier_name=supplier_name))


# --- FUNGSI EDIT ENTRY ---
@app.route('/data/edit/<int:row_id>', methods=['POST'])
def edit_entry(row_id):
    form = request.form
    supplier_name_redirect = "" # Untuk redirect jika error
    try:
        form_month = datetime.strptime(form['month'], '%Y-%m')
        target_year = form_month.year
        file_path = get_data_file_path(target_year)
        df = load_data(file_path)
        
        if not df.empty and row_id in df['row_id'].values:
            # Dapatkan nama supplier SEBELUM update, untuk redirect
            supplier_name_redirect = df.loc[df['row_id'] == row_id, 'SUPPLIER NAME'].iloc[0]

            total_delivery_val = int(float(form['total_delivery']))
            on_time_val = int(float(form['on_time']))
            achievement_val = (on_time_val / total_delivery_val) if total_delivery_val > 0 else (1.0 if on_time_val == 0 else 0.0)

            # Update data di DataFrame
            df.loc[df['row_id'] == row_id, 'CLOSING MONTH'] = form_month
            df.loc[df['row_id'] == row_id, 'TOTAL DELIVERY ITEM'] = total_delivery_val
            df.loc[df['row_id'] == row_id, 'ON TIME'] = on_time_val
            df.loc[df['row_id'] == row_id, 'MINUS'] = int(float(form['minus']))
            # --- PERUBAHAN DI SINI ---
            df.loc[df['row_id'] == row_id, 'TARGET DELIVERY'] = 0.9 # Nilai tetap 90%
            df.loc[df['row_id'] == row_id, 'ACHIEVEMENT'] = achievement_val
            df.loc[df['row_id'] == row_id, 'Purchase Amount'] = float(form['purchase_amount'])
            # Update kolom lain jika perlu (misal: 'ITEM DELAY')
            
            if save_data(df, file_path): 
                flash('Data updated successfully!', 'success')
            # Redirect ke dashboard supplier yang diedit
            return redirect(url_for('dashboard', supplier_name=supplier_name_redirect))
        else: 
            flash(f'Data with ID {row_id} not found for year {target_year}.', 'danger')
            # Coba redirect ke index jika nama supplier tidak ditemukan
            return redirect(url_for('index'))
            
    except ValueError:
        flash('Invalid number format entered.', 'danger')
        # Redirect kembali ke supplier yang sedang diedit jika memungkinkan
        if supplier_name_redirect:
            return redirect(url_for('dashboard', supplier_name=supplier_name_redirect))
        else:
            return redirect(url_for('index')) # Fallback ke index
    except Exception as e:
        flash(f'An error occurred during edit: {e}', 'danger')
        if supplier_name_redirect:
            return redirect(url_for('dashboard', supplier_name=supplier_name_redirect))
        else:
            return redirect(url_for('index'))


# --- FUNGSI DELETE ENTRY ---
@app.route('/data/delete/<int:row_id>', methods=['POST'])
def delete_entry(row_id):
    df_all = load_all_data()
    supplier_name_redirect = "" # Untuk redirect
    if df_all.empty:
        flash('No data found to delete from.', 'danger')
        return redirect(url_for('index'))

    target_row = df_all[df_all['row_id'] == row_id]
    
    if not target_row.empty:
        supplier_name_redirect = target_row.iloc[0]['SUPPLIER NAME']
        closing_month = pd.to_datetime(target_row.iloc[0]['CLOSING MONTH'])
        target_year = closing_month.year
        file_path = get_data_file_path(target_year)
        
        try:
            df_single_file = load_data(file_path)
            # Pastikan row_id ada sebelum mencoba menghapus
            if not df_single_file.empty and row_id in df_single_file['row_id'].values:
                df_single_file = df_single_file[df_single_file['row_id'] != row_id]
                if save_data(df_single_file, file_path): 
                    flash(f'Data for {closing_month.strftime("%B %Y")} deleted successfully from data_{target_year}.xlsx!', 'success')
                else:
                    flash('Failed to save after deleting data.', 'danger') # Pesan error lebih spesifik
            else:
                 flash(f'Data with ID {row_id} not found in specific file data_{target_year}.xlsx.', 'warning')

        except Exception as e:
             flash(f'An error occurred during deletion: {e}', 'danger')
    else: 
        flash(f'Data with ID {row_id} not found in any file.', 'danger')
        return redirect(url_for('index')) # Redirect ke index jika data tidak ditemukan sama sekali

    # Redirect ke dashboard supplier yang datanya dihapus, jika nama supplier ada
    if supplier_name_redirect:
         return redirect(url_for('dashboard', supplier_name=supplier_name_redirect))
    else:
         return redirect(url_for('index')) # Fallback ke index jika nama supplier tidak ada


# --- Rute Lainnya ---
@app.route('/check_supplier', methods=['POST'])
def check_supplier():
    df = load_all_data()
    name_to_check = request.form.get('supplier_name', '').strip().lower()
    if df.empty or 'SUPPLIER NAME' not in df.columns:
        return jsonify({'exists': False})
    # Gunakan .any() untuk cek keberadaan
    is_present = df['SUPPLIER NAME'].str.strip().str.lower().eq(name_to_check).any()
    return jsonify({'exists': bool(is_present)})

@app.context_processor
def inject_suppliers():
    df = load_all_data()
    supplier_list = []
    if not df.empty and 'SUPPLIER NAME' in df.columns:
        supplier_list = sorted(df['SUPPLIER NAME'].dropna().unique().tolist()) # Tambah dropna()
    return dict(all_suppliers=supplier_list)

# --- Run App ---
if __name__ == '__main__':
    app.run(debug=True, host='127.0.0.1', port=5001)