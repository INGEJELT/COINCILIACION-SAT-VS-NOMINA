from flask import Flask, render_template, request, redirect, url_for, session
import pandas as pd
import os
import sqlite3
from werkzeug.utils import secure_filename
import warnings
from datetime import datetime

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

app = Flask(__name__)
app.secret_key = 'luca_rg_secure_enterprise_key'
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def init_db():
    conn = sqlite3.connect('conciliaciones.db')
    cursor = conn.cursor()
    # Crear tabla con columnas para montos originales
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS conciliaciones (
            uuid TEXT PRIMARY KEY,
            empresa TEXT,
            clave TEXT,
            nombre TEXT,
            rfc TEXT,
            curp TEXT,
            periodo TEXT,
            fecha_emision TEXT,
            estado_sat TEXT,
            estado_nomina TEXT,
            total_sat REAL,
            total_nomina REAL,
            total_ptu REAL,
            total_sat_original REAL,
            total_nomina_original REAL,
            total_ptu_original REAL,
            conciliacion TEXT,
            fecha_importacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    # Verificar si las columnas originales existen, si no, agregarlas
    cursor.execute("PRAGMA table_info(conciliaciones)")
    columns = [col[1] for col in cursor.fetchall()]
    if 'total_sat_original' not in columns:
        cursor.execute("ALTER TABLE conciliaciones ADD COLUMN total_sat_original REAL")
    if 'total_nomina_original' not in columns:
        cursor.execute("ALTER TABLE conciliaciones ADD COLUMN total_nomina_original REAL")
    if 'total_ptu_original' not in columns:
        cursor.execute("ALTER TABLE conciliaciones ADD COLUMN total_ptu_original REAL")
    conn.commit()
    conn.close()

init_db()

def safe_float(val):
    try:
        if pd.isna(val): return 0.0
        v = str(val).replace('$', '').replace(',', '').strip()
        if v.lower() in ['nan', 'none', '']: return 0.0
        return float(v)
    except:
        return 0.0

def get_best_val(row, cols):
    for c in cols:
        if c in row and pd.notna(row[c]):
            v = str(row[c]).strip()
            if v.lower() not in ['nan', 'none', 'nat', '<na>', 'n/a', '']:
                return v
    return 'N/A'

def limpiar_periodo(val):
    v = str(val).strip()
    if v.lower() in ['nan', 'none', 'nat', '<na>', 'n/a', '']: return 'N/A'
    return v.split('.')[0]

def obtener_semana(fecha):
    try:
        if isinstance(fecha, pd.Timestamp):
            fecha = fecha.to_pydatetime()
        elif isinstance(fecha, str):
            for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%Y/%m/%d']:
                try:
                    fecha = datetime.strptime(fecha, fmt)
                    break
                except:
                    continue
            if isinstance(fecha, str):
                return 'N/A'
        if isinstance(fecha, datetime):
            semana = fecha.isocalendar()[1]
            return str(semana)
        return 'N/A'
    except:
        return 'N/A'

def encontrar_columna_fecha(df):
    for col in df.columns:
        col_upper = col.upper().strip()
        if any(p in col_upper for p in ['FECHA EMISION', 'FECHA DE EMISION', 'FECHA DE EMISIÓN']):
            return col
    for col in df.columns:
        if 'FECHA' in col.upper():
            return col
    return None

def encontrar_columna_estado(df):
    for col in df.columns:
        if col.upper().strip() == 'ESTADO':
            return col
    for col in df.columns:
        if col.upper().strip() == 'STATUS':
            return col
    for col in df.columns:
        col_upper = col.upper()
        if ('ESTADO' in col_upper or 'ESTATUS' in col_upper) and 'CANCELACION' not in col_upper:
            return col
    return None

def leer_excel_info(ruta, es_sat=False):
    if not ruta or not os.path.exists(ruta):
        return pd.DataFrame(), 'N/A'
        
    df_temp = pd.read_excel(ruta, header=None, nrows=15, engine='openpyxl')
    header_row = 0
    empresa_extraida = 'N/A'
    
    if not es_sat:
        val = str(df_temp.iloc[0, 0]).strip()
        if val and val != 'nan':
            empresa_extraida = val

    for index, row in df_temp.iterrows():
        if row.astype(str).str.contains('UUID', case=False, na=False).any():
            header_row = index
            break
            
    df = pd.read_excel(ruta, header=header_row, engine='openpyxl')
    return df, empresa_extraida

def guardar_en_base_datos(datos):
    conn = sqlite3.connect('conciliaciones.db')
    cursor = conn.cursor()
    for fila in datos:
        cursor.execute('SELECT * FROM conciliaciones WHERE uuid = ?', (fila['UUID'],))
        existing = cursor.fetchone()
        
        estado_sat = str(fila.get('ESTADO_SAT', '')).upper().strip()
        estado_nom = str(fila.get('ESTADO_NOMINA', '')).upper().strip()
        
        # Montos originales (siempre se guardan)
        f_sat_orig = safe_float(fila.get('TOTAL_SAT'))
        f_nom_orig = safe_float(fila.get('TOTAL_NOMINA'))
        f_ptu_orig = safe_float(fila.get('TOTAL_PTU'))
        
        # Montos para sumar (0 si está cancelado)
        f_sat = f_sat_orig if estado_sat not in ['CANCELADO', 'CANCELADA'] else 0.0
        f_nom = f_nom_orig if estado_nom not in ['CANCELADO', 'CANCELADA'] else 0.0
        f_ptu = f_ptu_orig if estado_sat not in ['CANCELADO', 'CANCELADA'] and estado_nom not in ['CANCELADO', 'CANCELADA'] else 0.0
        
        sat_tiene = f_sat > 0 or estado_sat in ['VIGENTE', 'TIMBRADO']
        nom_tiene = f_nom > 0 or estado_nom in ['TIMBRADO', 'VIGENTE']
        sat_cancelado = estado_sat in ['CANCELADO', 'CANCELADA']
        nom_cancelado = estado_nom in ['CANCELADO', 'CANCELADA']
        
        if sat_cancelado and nom_cancelado:
            e_con = 'Coincide'
        elif sat_tiene and nom_tiene and abs(f_sat - f_nom) < 0.01:
            e_con = 'Coincide'
        elif sat_tiene and not nom_tiene and not nom_cancelado:
            e_con = 'Solo en SAT'
        elif nom_tiene and not sat_tiene and not sat_cancelado:
            e_con = 'Falta en SAT'
        else:
            e_con = 'Diferencia'

        if existing:
            is_nomina_o_ptu = f_nom_orig > 0 or f_ptu_orig > 0
            
            e_emp = fila.get('EMPRESA', 'N/A') if is_nomina_o_ptu and fila.get('EMPRESA') not in ['N/A', 'nan', ''] else existing[1]
            if e_emp in ['N/A', 'nan', '']: e_emp = existing[1]

            e_per = fila.get('PERIODO', 'N/A') if is_nomina_o_ptu and fila.get('PERIODO') not in ['N/A', 'nan', ''] else existing[6]
            if e_per in ['N/A', 'nan', '']: e_per = existing[6]

            e_cla = fila.get('CLAVE', 'N/A') if is_nomina_o_ptu and fila.get('CLAVE') not in ['N/A', 'nan', ''] else existing[2]
            if e_cla in ['N/A', 'nan', '']: e_cla = existing[2]
            
            e_nom = fila.get('NOMBRE', 'N/A') if fila.get('NOMBRE') not in ['N/A', 'nan', ''] else existing[3]
            e_rfc = fila.get('RFC', 'N/A') if fila.get('RFC') not in ['N/A', 'nan', ''] else existing[4]
            e_cur = fila.get('CURP', 'N/A') if fila.get('CURP') not in ['N/A', 'nan', ''] else existing[5]
            e_fec = fila.get('FECHA DE EMISIÓN', 'N/A') if fila.get('FECHA DE EMISIÓN') not in ['N/A', 'nan', ''] else existing[7]
            
            e_esat = fila.get('ESTADO_SAT', 'No en SAT') if fila.get('ESTADO_SAT') not in ['No en SAT', 'N/A', ''] else existing[8]
            e_enom = fila.get('ESTADO_NOMINA', 'N/A') if fila.get('ESTADO_NOMINA') not in ['N/A', ''] else existing[9]
            
            # Si el registro existente tiene valores originales, mantenerlos
            e_tsat_orig = f_sat_orig if f_sat_orig > 0 else (existing[13] if len(existing) > 13 else 0)
            e_tnom_orig = f_nom_orig if f_nom_orig > 0 else (existing[14] if len(existing) > 14 else 0)
            e_tptu_orig = f_ptu_orig if f_ptu_orig > 0 else (existing[15] if len(existing) > 15 else 0)
            
            # Los montos para sumar (0 si cancelado)
            e_tsat = f_sat if f_sat > 0 else (existing[10] if len(existing) > 10 else 0)
            e_tnom = f_nom if f_nom > 0 else (existing[11] if len(existing) > 11 else 0)
            e_tptu = f_ptu if f_ptu > 0 else (existing[12] if len(existing) > 12 else 0)

            cursor.execute('''
                UPDATE conciliaciones 
                SET empresa=?, clave=?, nombre=?, rfc=?, curp=?, periodo=?, fecha_emision=?, 
                    estado_sat=?, estado_nomina=?, total_sat=?, total_nomina=?, total_ptu=?,
                    total_sat_original=?, total_nomina_original=?, total_ptu_original=?, conciliacion=?
                WHERE uuid=?
            ''', (e_emp, e_cla, e_nom, e_rfc, e_cur, e_per, e_fec, e_esat, e_enom, 
                  e_tsat, e_tnom, e_tptu, e_tsat_orig, e_tnom_orig, e_tptu_orig, e_con, fila['UUID']))
            
        else:
            cursor.execute('''
                INSERT INTO conciliaciones 
                (uuid, empresa, clave, nombre, rfc, curp, periodo, fecha_emision, 
                 estado_sat, estado_nomina, total_sat, total_nomina, total_ptu,
                 total_sat_original, total_nomina_original, total_ptu_original, conciliacion)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                fila['UUID'], fila.get('EMPRESA', 'N/A'), fila.get('CLAVE', 'N/A'), fila.get('NOMBRE', 'N/A'), 
                fila.get('RFC', 'N/A'), fila.get('CURP', 'N/A'), fila.get('PERIODO', 'N/A'),
                fila.get('FECHA DE EMISIÓN', 'N/A'), fila.get('ESTADO_SAT', 'No en SAT'), fila.get('ESTADO_NOMINA', 'N/A'),
                f_sat, f_nom, f_ptu, f_sat_orig, f_nom_orig, f_ptu_orig, e_con
            ))
    conn.commit()
    conn.close()

def procesar_conciliacion(ruta_sat, ruta_nomina, ruta_ptu):
    try:
        df_sat, _ = leer_excel_info(ruta_sat, es_sat=True)
        df_nomina, emp_nom = leer_excel_info(ruta_nomina, es_sat=False)
        df_ptu, emp_ptu = leer_excel_info(ruta_ptu, es_sat=False)

        for df in [df_sat, df_nomina, df_ptu]:
            if not df.empty:
                df.columns = df.columns.str.upper().str.strip()
                if 'UUID' in df.columns:
                    df['UUID'] = df['UUID'].astype(str).str.strip().str.upper()

        if not df_sat.empty:
            if 'PERIODO' in df_sat.columns:
                df_sat['PERIODO_SAT'] = df_sat['PERIODO'].apply(limpiar_periodo)
            else:
                col_fecha = encontrar_columna_fecha(df_sat)
                if col_fecha:
                    try:
                        df_sat['PERIODO_SAT'] = df_sat[col_fecha].apply(obtener_semana)
                    except:
                        df_sat['PERIODO_SAT'] = 'N/A'
                else:
                    df_sat['PERIODO_SAT'] = 'N/A'
            
            df_sat['PERIODO_SAT'] = df_sat['PERIODO_SAT'].fillna('N/A')
            df_sat['EMPRESA_SAT'] = df_sat['NOMBRE EMISOR'] if 'NOMBRE EMISOR' in df_sat.columns else 'N/A'

            col_estado_sat = encontrar_columna_estado(df_sat)
            if col_estado_sat:
                df_sat['ESTADO_SAT'] = df_sat[col_estado_sat]
            else:
                df_sat['ESTADO_SAT'] = 'N/A'

            sat_df = df_sat[['UUID', 'TOTAL', 'ESTADO_SAT', 'NOMBRE RECEPTOR', 'RFC RECEPTOR', 'FECHA EMISION', 'PERIODO_SAT', 'EMPRESA_SAT']].copy()
            sat_df.rename(columns={'TOTAL': 'TOTAL_SAT', 'NOMBRE RECEPTOR': 'NOMBRE_SAT', 'RFC RECEPTOR': 'RFC_SAT', 'FECHA EMISION': 'FECHA_SAT'}, inplace=True)
        else:
            sat_df = pd.DataFrame(columns=['UUID', 'TOTAL_SAT', 'ESTADO_SAT', 'NOMBRE_SAT', 'RFC_SAT', 'FECHA_SAT', 'PERIODO_SAT', 'EMPRESA_SAT'])

        if not df_nomina.empty:
            if 'PERIODO' not in df_nomina.columns:
                col_fecha = encontrar_columna_fecha(df_nomina)
                if col_fecha:
                    try:
                        df_nomina['PERIODO'] = df_nomina[col_fecha].apply(obtener_semana)
                    except:
                        df_nomina['PERIODO'] = 'N/A'
                else:
                    df_nomina['PERIODO'] = 'N/A'
            df_nomina['PERIODO'] = df_nomina['PERIODO'].apply(limpiar_periodo)

            col_estado_nom = encontrar_columna_estado(df_nomina)
            if col_estado_nom:
                df_nomina['ESTADO_NOM'] = df_nomina[col_estado_nom]
            else:
                df_nomina['ESTADO_NOM'] = 'N/A'

            nom_df = df_nomina[['UUID', 'TOTAL', 'CLAVE', 'NOMBRE', 'RFC', 'CURP', 'PERIODO', 'FECHA DE EMISIÓN', 'ESTADO_NOM']].copy()
            nom_df['EMPRESA_NOM'] = emp_nom
            nom_df.rename(columns={'TOTAL': 'TOTAL_NOMINA', 'ESTADO_NOM': 'ESTADO_NOMINA', 'NOMBRE': 'NOMBRE_NOM', 'RFC': 'RFC_NOM', 'PERIODO': 'PERIODO_NOM', 'FECHA DE EMISIÓN': 'FECHA_NOM'}, inplace=True)
        else:
            nom_df = pd.DataFrame(columns=['UUID', 'TOTAL_NOMINA', 'ESTADO_NOMINA', 'CLAVE', 'NOMBRE_NOM', 'RFC_NOM', 'CURP', 'PERIODO_NOM', 'FECHA_NOM', 'EMPRESA_NOM'])

        if not df_ptu.empty:
            if 'PERIODO' not in df_ptu.columns:
                col_fecha = encontrar_columna_fecha(df_ptu)
                if col_fecha:
                    try:
                        df_ptu['PERIODO'] = df_ptu[col_fecha].apply(obtener_semana)
                    except:
                        df_ptu['PERIODO'] = 'N/A'
                else:
                    df_ptu['PERIODO'] = 'N/A'
            df_ptu['PERIODO'] = df_ptu['PERIODO'].apply(limpiar_periodo)

            col_estado_ptu = encontrar_columna_estado(df_ptu)
            if col_estado_ptu:
                df_ptu['ESTADO_PTU'] = df_ptu[col_estado_ptu]
            else:
                df_ptu['ESTADO_PTU'] = 'N/A'

            ptu_df = df_ptu[['UUID', 'TOTAL', 'CLAVE', 'NOMBRE', 'RFC', 'CURP', 'PERIODO', 'FECHA DE EMISIÓN', 'ESTADO_PTU']].copy()
            ptu_df['EMPRESA_PTU'] = emp_ptu
            ptu_df.rename(columns={'TOTAL': 'TOTAL_PTU', 'ESTADO_PTU': 'ESTADO_PTU', 'NOMBRE': 'NOMBRE_PTU', 'RFC': 'RFC_PTU', 'PERIODO': 'PERIODO_PTU', 'FECHA DE EMISIÓN': 'FECHA_PTU', 'CLAVE': 'CLAVE_PTU', 'CURP': 'CURP_PTU'}, inplace=True)
        else:
            ptu_df = pd.DataFrame(columns=['UUID', 'TOTAL_PTU', 'ESTADO_PTU', 'CLAVE_PTU', 'NOMBRE_PTU', 'RFC_PTU', 'CURP_PTU', 'PERIODO_PTU', 'FECHA_PTU', 'EMPRESA_PTU'])

        merged = pd.merge(sat_df, nom_df, on='UUID', how='outer')
        merged = pd.merge(merged, ptu_df, on='UUID', how='outer')

        merged['EMPRESA'] = merged.apply(lambda r: get_best_val(r, ['EMPRESA_NOM', 'EMPRESA_PTU', 'EMPRESA_SAT']), axis=1)
        merged['PERIODO'] = merged.apply(lambda r: get_best_val(r, ['PERIODO_NOM', 'PERIODO_PTU', 'PERIODO_SAT']), axis=1).apply(limpiar_periodo)
        merged['CLAVE'] = merged.apply(lambda r: get_best_val(r, ['CLAVE', 'CLAVE_PTU']), axis=1)
        merged['CURP'] = merged.apply(lambda r: get_best_val(r, ['CURP', 'CURP_PTU']), axis=1)
        merged['NOMBRE'] = merged.apply(lambda r: get_best_val(r, ['NOMBRE_NOM', 'NOMBRE_PTU', 'NOMBRE_SAT']), axis=1)
        merged['RFC'] = merged.apply(lambda r: get_best_val(r, ['RFC_NOM', 'RFC_PTU', 'RFC_SAT']), axis=1)
        merged['FECHA DE EMISIÓN'] = merged.apply(lambda r: get_best_val(r, ['FECHA_NOM', 'FECHA_PTU', 'FECHA_SAT']), axis=1)
        merged['ESTADO_NOMINA'] = merged.apply(lambda r: get_best_val(r, ['ESTADO_NOMINA', 'ESTADO_PTU']), axis=1)
        merged['ESTADO_SAT'] = merged.apply(lambda r: get_best_val(r, ['ESTADO_SAT']), axis=1)

        if not merged.empty:
            resultado_lista = merged.to_dict('records')
            guardar_en_base_datos(resultado_lista)
            
    except Exception as e:
        print(f"Error procesando archivos: {e}")

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        if request.form['username'] == 'admin' and request.form['password'] == 'luca2026':
            session['logged_in'] = True
            return redirect(url_for('index'))
        else:
            error = 'Credenciales inválidas. Intente de nuevo.'
    return render_template('login.html', error=error)

@app.route('/', methods=['GET', 'POST'])
def index():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
        
    datos_conciliados = None
    totales_globales = None
    
    if request.method == 'POST':
        sat_file = request.files.get('sat_file')
        nomina_file = request.files.get('nomina_file')
        ptu_file = request.files.get('ptu_file')
        
        if (sat_file and sat_file.filename) or (nomina_file and nomina_file.filename) or (ptu_file and ptu_file.filename):
            ruta_sat, ruta_nomina, ruta_ptu = None, None, None
            
            if sat_file and sat_file.filename:
                ruta_sat = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(sat_file.filename))
                sat_file.save(ruta_sat)
            if nomina_file and nomina_file.filename:
                ruta_nomina = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(nomina_file.filename))
                nomina_file.save(ruta_nomina)
            if ptu_file and ptu_file.filename:
                ruta_ptu = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(ptu_file.filename))
                ptu_file.save(ruta_ptu)

            procesar_conciliacion(ruta_sat, ruta_nomina, ruta_ptu)
            
        return redirect(url_for('index'))
        
    else:
        try:
            # Obtener todos los registros para la tabla (mostrar montos originales)
            conn = sqlite3.connect('conciliaciones.db')
            df_all = pd.read_sql_query("""
                SELECT 
                    uuid, empresa, clave, nombre, rfc, curp, periodo, fecha_emision,
                    estado_sat, estado_nomina,
                    COALESCE(total_sat_original, total_sat) as total_sat,
                    COALESCE(total_nomina_original, total_nomina) as total_nomina,
                    COALESCE(total_ptu_original, total_ptu) as total_ptu,
                    conciliacion
                FROM conciliaciones
            """, conn)
            conn.close()
            
            if not df_all.empty:
                df_all.rename(columns={
                    'uuid': 'UUID', 'empresa': 'EMPRESA', 'clave': 'CLAVE', 'nombre': 'NOMBRE',
                    'rfc': 'RFC', 'curp': 'CURP', 'periodo': 'PERIODO',
                    'fecha_emision': 'FECHA DE EMISIÓN', 'estado_sat': 'ESTADO_SAT',
                    'estado_nomina': 'ESTADO_NOMINA', 'total_sat': 'TOTAL_SAT',
                    'total_nomina': 'TOTAL_NOMINA', 'total_ptu': 'TOTAL_PTU', 'conciliacion': 'CONCILIACION'
                }, inplace=True)
                
                df_all['DIFERENCIA_SAT_NOMINA'] = df_all['TOTAL_SAT'] - df_all['TOTAL_NOMINA']
                
                # Agregar columna de Observaciones para CFDI cancelados
                df_all['OBSERVACIONES'] = df_all.apply(
                    lambda row: 'POSIBLE DIFERENCIA ENTRE TOTALES' 
                    if row.get('ESTADO_SAT') in ['Cancelado', 'Cancelada'] or row.get('ESTADO_NOMINA') in ['CANCELADO', 'CANCELADA'] 
                    else '', axis=1
                )
                
                datos_conciliados = df_all.to_dict('records')
            
            # Calcular totales excluyendo cancelados (usando total_sat y total_nomina que ya son 0 para cancelados)
            conn = sqlite3.connect('conciliaciones.db')
            query = """
                SELECT 
                    SUM(total_sat) as total_sat,
                    SUM(total_nomina) as total_nomina,
                    SUM(total_ptu) as total_ptu
                FROM conciliaciones
            """
            df_totales = pd.read_sql_query(query, conn)
            conn.close()
            
            totales_globales = {
                'sat': float(df_totales['total_sat'].iloc[0] if not df_totales.empty else 0),
                'nomina': float(df_totales['total_nomina'].iloc[0] if not df_totales.empty else 0),
                'ptu': float(df_totales['total_ptu'].iloc[0] if not df_totales.empty else 0)
            }
            
        except Exception as e:
            print(f"Error consultando persistencia en index: {e}")

    return render_template('index.html', data=datos_conciliados, totales=totales_globales)

@app.route('/periodos')
def periodos():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
        
    conn = sqlite3.connect('conciliaciones.db')
    
    # Usar total_sat y total_nomina (que ya son 0 para cancelados)
    query = '''
        SELECT 
            empresa as EMPRESA,
            periodo as PERIODO,
            SUM(total_sat) as TOTAL_SAT,
            SUM(total_nomina) as TOTAL_NOMINA,
            (SUM(total_sat) - SUM(total_nomina)) as DIFERENCIA,
            SUM(total_ptu) as TOTAL_PTU,
            COUNT(uuid) as TOTAL_REGISTROS,
            SUM(CASE WHEN CAST(total_sat AS REAL) > 0.01 THEN 1 ELSE 0 END) as REGISTROS_SAT,
            SUM(CASE WHEN CAST(total_nomina AS REAL) > 0.01 THEN 1 ELSE 0 END) as REGISTROS_NOMINA,
            SUM(CASE WHEN CAST(total_ptu AS REAL) > 0.01 THEN 1 ELSE 0 END) as REGISTROS_PTU
        FROM conciliaciones
        WHERE periodo IS NOT NULL AND periodo != 'N/A' AND periodo != ''
        GROUP BY empresa, periodo
        ORDER BY empresa ASC, CAST(periodo AS INTEGER) ASC
    '''
    try:
        df_periodos = pd.read_sql_query(query, conn)
        lista_periodos = df_periodos.to_dict('records')
    except:
        lista_periodos = []
    finally:
        conn.close()
        
    return render_template('periodos.html', periodos=lista_periodos)

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(debug=True, port=5032)