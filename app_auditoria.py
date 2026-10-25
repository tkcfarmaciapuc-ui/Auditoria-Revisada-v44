# -*- coding: utf-8 -*-
"""
🚀 AUDITORÍA UNIFICADA - VERSIÓN LOCAL (.EXE)
Adaptado para ejecución en PC local y compilación con PyInstaller.
"""

import os
import sys
import glob
import re
import pickle
import pandas as pd
from pypdf import PdfReader
import pdfplumber
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import PatternFill, Font

# ==========================================
# 📍 MANEJO DE RUTAS DINÁMICAS (COMPATIBLE CON PYINSTALLER .EXE)
# ==========================================
def obtener_ruta_base():
    """ Obtiene la ruta del directorio actual donde se está ejecutando el .exe o el script. """
    if getattr(sys, 'frozen', False):
        # Si el programa está ejecutándose como un archivo .exe compilado
        return os.path.dirname(sys.executable)
    else:
        # Si se está ejecutando como script .py normal
        return os.path.dirname(os.path.abspath(__file__))

base_path = obtener_ruta_base()
CACHE_FILE = os.path.join(base_path, 'cache_auditoria_estado.pkl')
UMBRAL_DOSIS = 6

# --- CONFIGURACIONES GLOBALES ---
codigos_especiales = {
    'FX0009', 'FX0010', 'FX0012', 'FX0013', 'FX0014',
    'FX0026', 'FX0030', 'FX0031', 'FX0033', 'FX0040',
    'FX0041', 'FX0050', 'FX9022'
}

# --- GRUPOS DE PRINCIPIO ACTIVO (P.A.) Y SUS FAMILIAS ---
grupos_pa = {
    "Tacrolimus": {'FX0012', 'FX0031', 'FX0033', 'FX0041'},
    "Vancomicina": {'FK3055', 'FK3065'},
    "TAZOBACTAM/PIPERACILINA": {'FK3063', 'FR1091'},
    "Meropenem": {'FK3085', 'FK3084'}
}

grupo_medicamentos = set().union(*grupos_pa.values())

codigos_confirmar_necesidad = {
    'FH9021', 'FH9000', 'FH9015', 'FH9005', 'FX9012', 'FK8069', 'FI9041',
    'FK8053', 'FH8005', 'FH8002', 'FH6006', 'FK0016', 'FX0001', 'FF7005',
    'FH8010', 'FK8100'
}

codigos_dosis_alta_especial = {
    'FA3026', 'FA3029', 'FA4001', 'FA4005', 'FA4006', 'FA4013', 'FA4014', 'FA4026',
    'FA6018', 'FA8012', 'FB0034', 'FB0046', 'FB0063', 'FB0065', 'FB0069', 'FB0072',
    'FB0075', 'FB0078', 'FB0082', 'FB4002', 'FC3008', 'FC4019', 'FC4022', 'FC5004',
    'FD4007', 'FD4012', 'FD4023', 'FD4024', 'FD4030', 'FD4033', 'FD9011', 'FD9015',
    'FE3003', 'FE3013', 'FE4003', 'FE5012', 'FE6003', 'FE6004', 'FF3011', 'FF4053',
    'FF5014', 'FF7005', 'FG0001', 'FG0002', 'FG1049', 'FG1062', 'FG1067', 'FG5000',
    'FG5037', 'FG5038', 'FG5044', 'FG6021', 'FG6035', 'FG6037', 'FG8001', 'FG8007',
    'FG9010', 'FI0002', 'FI3005', 'FI3006', 'FI3007', 'FI3008', 'FI3014', 'FI3023',
    'FI9002', 'FI9017', 'FI9030', 'FI9041', 'FI9050', 'FJ2002', 'FJ2011', 'FJ2034',
    'FJ2037', 'FJ3015', 'FJ3016', 'FJ3018', 'FJ9012', 'FJ9013', 'FJ9032', 'FJ9033',
    'FK0016', 'FK2004', 'FK2005', 'FK2009', 'FK3001', 'FK3039', 'FK3042', 'FK3048',
    'FK3074', 'FK3096', 'FK4000', 'FK4005', 'FK5001', 'FK5004', 'FK8008', 'FK8011',
    'FK8013', 'FK8016', 'FK8017', 'FK8019', 'FK8024', 'FK8025', 'FK8026', 'FK8027',
    'FK8050', 'FK8051', 'FK8052', 'FK8068', 'FK8101', 'FM0000', 'FM9001', 'FM9002',
    'FM9005', 'FM9012', 'FM9013', 'FM9014', 'FM9015', 'FM9016', 'FM9018', 'FM9019',
    'FM9021', 'FN7012', 'FN7021', 'FN7026', 'FN7031', 'FN7048', 'FX0001', 'FX0009',
    'FX0010', 'FX0012', 'FX0013', 'FX0014', 'FX0017', 'FX0018', 'FX0026', 'FX0030',
    'FX0031', 'FX0033', 'FX0038', 'FX0040', 'FX0041', 'FX0050', 'FX0052', 'FX9012',
    'FX9015', 'FX9022', 'FH7027',
}

sku_excluidos_upc = {'FD4023', 'FD4024', 'FD4012', 'FI9040'}

# ==========================================
# 🧱 PASO 1: LÓGICA DE AUDITORÍA CON FECHAS CSV
# ==========================================
def ejecutar_auditoria(pdf_files, csv_path, cache_data):
    fecha_limite_normal = datetime.now() - timedelta(days=5)
    fecha_limite_extendida = datetime.now() - timedelta(days=10)

    df_csv = pd.read_csv(csv_path, sep=';', dtype=str, encoding='latin-1')
    df_csv['Fecha_DT'] = pd.to_datetime(df_csv['Fe.contabilización'], dayfirst=True, format='mixed')

    df_5_dias = df_csv[df_csv['Fecha_DT'] >= fecha_limite_normal]
    df_10_dias = df_csv[df_csv['Fecha_DT'] >= fecha_limite_extendida]

    dict_5_dias = {}
    for _, row in df_5_dias.iterrows():
        llave = str(row.iloc[21]).strip()
        fecha = str(row['Fe.contabilización']).strip()
        almacen_receptor = str(row['Almacén Receptor Pedido']).strip() if 'Almacén Receptor Pedido' in df_csv.columns else ""
        dict_5_dias[llave] = {'fecha': fecha, 'almacen_receptor': almacen_receptor}

    dict_10_dias = {}
    for _, row in df_10_dias.iterrows():
        llave = str(row.iloc[21]).strip()
        fecha = str(row['Fe.contabilización']).strip()
        almacen_receptor = str(row['Almacén Receptor Pedido']).strip() if 'Almacén Receptor Pedido' in df_csv.columns else ""
        dict_10_dias[llave] = {'fecha': fecha, 'almacen_receptor': almacen_receptor}

    df_10_dias_grupo = df_10_dias[df_10_dias.iloc[:, 21].str.strip().str.endswith(tuple(grupo_medicamentos))]
    dict_grupo_10_dias = {}
    for _, row in df_10_dias_grupo.iterrows():
        llave = str(row.iloc[21]).strip()
        fecha = str(row['Fe.contabilización']).strip()
        dict_grupo_10_dias[llave] = fecha

    target_ola = os.path.basename(pdf_files[-1]).replace('.pdf', '')
    processed_olas = cache_data.get('processed_olas', set()) if cache_data else set()

    if cache_data and 'memoria_olas' in cache_data:
        memoria_olas = cache_data['memoria_olas']
        memoria_pedidos_unicos = cache_data['memoria_pedidos_unicos']
        memoria_olas_unicas = cache_data['memoria_olas_unicas']
        memoria_grupo = cache_data['memoria_grupo']
        resultados_totales = cache_data['resultados_totales']
    else:
        memoria_olas = {}
        memoria_pedidos_unicos = {}
        memoria_olas_unicas = {}
        memoria_grupo = {}
        resultados_totales = []

    code_pattern = re.compile(r' F[A-Z0-9] ox |\bF[A-Z0-9]\d{4}\b')
    nuevos_procesados = 0

    for pdf_path in pdf_files:
        nombre_ola = os.path.basename(pdf_path).replace('.pdf', '')

        if nombre_ola in processed_olas and nombre_ola != target_ola:
            continue

        nuevos_procesados += 1
        print(f"📄 Procesando PDF nuevo en Auditoría: {nombre_ola}...")

        try:
            reader = PdfReader(pdf_path)
        except Exception as e:
            print(f"⚠️ Error al leer el PDF '{pdf_path}': {e}. Se saltará.")
            continue

        current_pedido = "Desconocido"
        current_rut = "Desconocido"
        current_paciente = "Desconocido"
        current_almacen = "Desconocido"

        for page in reader.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split('\n'):
                clean_line = line.replace('"', '').replace(',', '').strip()
                if "Numero de pedido:" in clean_line:
                    current_pedido = clean_line.split("Numero de pedido:")[-1].strip()
                if "Rut del Paciente:" in clean_line:
                    current_rut = clean_line.split("Rut del Paciente:")[-1].strip()
                if "Nombre del paciente:" in clean_line:
                    current_paciente = clean_line.split("Nombre del paciente:")[-1].strip()
                if "Almacén:" in clean_line:
                    current_almacen = clean_line.split("Almacén:")[-1].strip()

                match = code_pattern.search(clean_line)
                if match:
                    sap_code = match.group(0)

                    linea_restante = clean_line.split(sap_code)[-1]
                    cantidad_match = re.findall(r'\b\d{1,3}(?:\.\d{3})*\b|\b\d+\b', linea_restante)

                    if len(cantidad_match) >= 2:
                        raw_cant = cantidad_match[-2]
                    elif cantidad_match:
                        raw_cant = cantidad_match[-1]
                    else:
                        raw_cant = "1"

                    cant_limpia = raw_cant.replace('.', '').replace(',', '')
                    cantidad = str(int(cant_limpia)) if cant_limpia.isdigit() else "1"

                    llave_unica = f"{current_rut}{sap_code}"
                    es_duplicado = False
                    origen = ""
                    pedido_orig = ""

                    familia_actual = None
                    for fam_nombre, fam_skus in grupos_pa.items():
                        if sap_code in fam_skus:
                            familia_actual = fam_nombre
                            break

                    if familia_actual:
                        ya_tiene_grupo_csv = False
                        fecha_grupo_csv = ""
                        for cod in grupos_pa[familia_actual]:
                            k_test = f"{current_rut}{cod}"
                            if k_test in dict_grupo_10_dias:
                                ya_tiene_grupo_csv = True
                                fecha_grupo_csv = dict_grupo_10_dias[k_test]
                                break

                        if ya_tiene_grupo_csv:
                            es_duplicado = True
                            origen = f"Duplicidad de {familia_actual} (CSV Fecha: {fecha_grupo_csv})"
                            pedido_orig = "Ver en SAP (Control)"
                        elif current_rut in memoria_grupo and memoria_grupo[current_rut]['Familia'] == familia_actual:
                            if memoria_grupo[current_rut]['Pedido'] != current_pedido:
                                es_duplicado = True
                                ola_origen = memoria_grupo[current_rut]['Ola']
                                sku_previo = memoria_grupo[current_rut]['SKU']
                                ped_previo = memoria_grupo[current_rut]['Pedido']
                                cant_previo = memoria_grupo[current_rut].get('Cant', '1')

                                if ola_origen == nombre_ola:
                                    origen = f"Duplicidad de {familia_actual} (Ola Actual)"
                                else:
                                    origen = f"Duplicidad de {familia_actual} (Ola Anterior: {ola_origen})"

                                pedido_orig = f"{ped_previo} | {sku_previo} ({cant_previo})"

                    if not es_duplicado:
                        if sap_code in codigos_especiales and llave_unica in dict_10_dias:
                            es_duplicado = True
                            fecha_prev = dict_10_dias[llave_unica]['fecha']
                            origen = f"Historial CSV (Fecha: {fecha_prev})"
                            pedido_orig = "Ver en SAP (Control)"
                        elif sap_code not in codigos_especiales and llave_unica in dict_5_dias:
                            es_duplicado = True
                            fecha_prev = dict_5_dias[llave_unica]['fecha']
                            origen = f"Historial CSV (Fecha: {fecha_prev})"
                            pedido_orig = "Ver en SAP (Control)"
                        elif llave_unica in memoria_pedidos_unicos:
                            pedidos_previos = list(memoria_pedidos_unicos[llave_unica])
                            if pedidos_previos:
                                es_duplicado = True
                                pedido_orig = " & ".join(pedidos_previos)
                                olas_encontradas = sorted(list(memoria_olas_unicas[llave_unica]))
                                olas_str = ", ".join(olas_encontradas)
                                origen = f"Repetición en Olas: {olas_str}"

                    if es_duplicado and sap_code in sku_excluidos_upc and "UPC" in current_almacen:
                        es_duplicado = False
                        estado = "🟢 APROBADO (Exento UPC)"
                    else:
                        if es_duplicado:
                            estado = "🚨 Med. Duplicado"
                        else:
                            estado = "🟢 APROBADO"

                    almacen_receptor_actual = ""
                    if llave_unica in dict_10_dias:
                        almacen_receptor_actual = dict_10_dias[llave_unica]['almacen_receptor']
                    elif llave_unica in dict_5_dias:
                        almacen_receptor_actual = dict_5_dias[llave_unica]['almacen_receptor']

                    if sap_code == 'FH6005':
                        if almacen_receptor_actual in ['HNEO', 'UPCP']:
                            if estado == "🟢 APROBADO":
                                estado = "🟦 REVISAR NECESIDAD"
                            else:
                                estado = f"{estado} | 🟦 REVISAR NECESIDAD"
                    elif sap_code == 'FH7027':
                        if almacen_receptor_actual not in ['HNEO', 'UPCP'] and almacen_receptor_actual != "":
                            if estado == "🟢 APROBADO":
                                estado = "🚨 NO ENVIAR"
                            else:
                                estado = f"{estado} | 🚨 NO ENVIAR"
                        else:
                            if estado == "🟢 APROBADO":
                                estado = "🟦 REVISAR NECESIDAD"
                            else:
                                estado = f"{estado} | 🟦 REVISAR NECESIDAD"

                    if sap_code in codigos_confirmar_necesidad and sap_code not in ['FH6005', 'FH7027']:
                        if estado == "🟢 APROBADO":
                            estado = "🟦 REVISAR NECESIDAD"
                        else:
                            estado = f"{estado} | 🟦 REVISAR NECESIDAD"

                    if llave_unica not in memoria_olas:
                        memoria_olas[llave_unica] = []
                        memoria_pedidos_unicos[llave_unica] = set()
                        memoria_olas_unicas[llave_unica] = set()

                    if current_pedido not in memoria_pedidos_unicos[llave_unica]:
                        memoria_pedidos_unicos[llave_unica].add(current_pedido)
                        memoria_olas[llave_unica].append({"Ola": nombre_ola, "Pedido": current_pedido})
                    else:
                        es_duplicado = True
                        if estado == "🟢 APROBADO":
                            estado = "🚨 Med. Duplicado"
                        elif "Med. Duplicado" not in estado:
                            estado = f"🚨 Med. Duplicado | {estado}"
                        if not origen:
                            origen = f"Duplicidad Interna (Mismo Pedido: {current_pedido})"
                            pedido_orig = current_pedido

                    memoria_olas_unicas[llave_unica].add(nombre_ola)

                    if familia_actual and current_rut not in memoria_grupo:
                        memoria_grupo[current_rut] = {
                            "Ola": nombre_ola,
                            "Pedido": current_pedido,
                            "SKU": sap_code,
                            "Cant": cantidad,
                            "Familia": familia_actual
                        }

                    resultados_totales.append({
                        "Ola": nombre_ola, "RUT": current_rut, "Nombre Paciente": current_paciente,
                        "SKU": sap_code, "Cant": cantidad, "Estado": estado,
                        "Pedido Original": pedido_orig, "Pedido Actual": current_pedido, "Origen": origen
                    })

        processed_olas.add(nombre_ola)

    return pd.DataFrame(resultados_totales), target_ola, processed_olas, {
        'memoria_olas': memoria_olas,
        'memoria_pedidos_unicos': memoria_pedidos_unicos,
        'memoria_olas_unicas': memoria_olas_unicas,
        'memoria_grupo': memoria_grupo,
        'resultados_totales': resultados_totales
    }, nuevos_procesados

# ==========================================
# 🧱 PASO 2: LÓGICA DE DOSIS ALTA
# ==========================================
def ejecutar_dosis_alta(ruta, umbral, codigos_dosis_especial):
    datos = []
    with pdfplumber.open(ruta) as pdf:
        texto_completo = "".join([page.extract_text() + "\n" for page in pdf.pages if page.extract_text()])

    texto_norm = texto_completo.replace('\n"', '"').replace('"\n"', '""').replace('",\n"', '","')
    pedidos = re.split(r'Numero de pedido:', texto_norm, flags=re.IGNORECASE)

    for pedido in pedidos[1:]:
        pedido_limpio = pedido.replace('"', '')
        num_pedido = re.search(r'^\s*(\d+)', pedido_limpio).group(1) if re.search(r'^\s*(\d+)', pedido_limpio) else "N/A"

        for linea in pedido_limpio.split('\n'):
            linea_str = linea.strip()
            sku_match = re.search(r'\b(F[A-Z0-9]\d{4})\b', linea_str)
            if sku_match:
                codigo = sku_match.group(1)
                linea_post_sku = linea_str.split(codigo)[-1]

                cant_match = re.findall(r'\b\d{1,3}(?:\.\d{3})*\b|\b\d+\b', linea_post_sku)

                if cant_match:
                    if len(cant_match) >= 2:
                        cant_str = cant_match[-2]
                    else:
                        cant_str = cant_match[-1]

                    cant_str_limpia = cant_str.replace('.', '').replace(',', '')
                    try:
                        cant = int(cant_str_limpia)
                    except ValueError:
                        cant = 0

                    if codigo in codigos_dosis_especial:
                        if cant > 1:
                            datos.append({'Número de Pedido': num_pedido, 'Código': codigo, 'Cantidad': cant, 'Tipo': 'Especial'})
                    else:
                        if cant >= umbral:
                            datos.append({'Número de Pedido': num_pedido, 'Código': codigo, 'Cantidad': cant, 'Tipo': 'Normal'})
    return pd.DataFrame(datos)

# ==========================================
# 🧱 PASO 4: EXTRACCIÓN DE CANTIDADES HISTÓRICAS
# ==========================================
def ejecutar_cantidades_alertas(pdf_files, cache_data):
    processed_cant_olas = cache_data.get('processed_cant_olas', set()) if cache_data else set()
    memoria_cantidades_pedidos = cache_data.get('memoria_cantidades_pedidos', {}) if cache_data else {}

    nuevos_procesados = 0

    for pdf_path in pdf_files:
        nombre_ola = os.path.basename(pdf_path).replace('.pdf', '')

        if nombre_ola in processed_cant_olas:
            continue

        nuevos_procesados += 1
        print(f"📦 Extrayendo cantidades de PDF nuevo: {nombre_ola}...")

        with pdfplumber.open(pdf_path) as pdf:
            texto_completo = "".join([page.extract_text() + "\n" for page in pdf.pages if page.extract_text()])

        texto_norm = texto_completo.replace('\n"', '"').replace('"\n"', '""').replace('",\n"', '","')
        pedidos = re.split(r'Numero de pedido:', texto_norm, flags=re.IGNORECASE)

        for pedido in pedidos[1:]:
            pedido_limpio = pedido.replace('"', '')
            num_pedido = re.search(r'^\s*(\d+)', pedido_limpio).group(1) if re.search(r'^\s*(\d+)', pedido_limpio) else "N/A"

            for linea in pedido_limpio.split('\n'):
                linea_str = linea.strip()
                sku_match = re.search(r'\b(F[A-Z0-9]\d{4})\b', linea_str)
                if sku_match:
                    codigo = sku_match.group(1)
                    linea_post_sku = linea_str.split(codigo)[-1]

                    cant_match = re.findall(r'\b\d{1,3}(?:\.\d{3})*\b|\b\d+\b', linea_post_sku)
                    if cant_match:
                        if len(cant_match) >= 2:
                            cant_str = cant_match[-2]
                        else:
                            cant_str = cant_match[-1]

                        cant_str_limpia = cant_str.replace('.', '').replace(',', '')
                        try:
                            cant = int(cant_str_limpia)
                        except ValueError:
                            cant = 0

                        clave_memoria = (str(num_pedido).strip(), str(codigo).strip())
                        memoria_cantidades_pedidos[clave_memoria] = cant

        processed_cant_olas.add(nombre_ola)

    return memoria_cantidades_pedidos, processed_cant_olas, nuevos_procesados

# ==========================================
# 🧠 PASO 3 Y 5: ORQUESTADOR PRINCIPAL
# ==========================================
def unificar_reportes():
    print(f"📁 Directorio de trabajo: {base_path}\n")

    csv_files = glob.glob(os.path.join(base_path, '*.csv'))
    pdf_files = sorted(glob.glob(os.path.join(base_path, 'P*.pdf')))

    if not csv_files:
        print("❌ ERROR: No se encontró ningún archivo CSV en la carpeta actual.")
        input("\nPresiona ENTER para salir...")
        return
    if not pdf_files:
        print("❌ ERROR: No se encontraron archivos PDF (P*.pdf) en la carpeta actual.")
        input("\nPresiona ENTER para salir...")
        return

    ultimo_pdf = pdf_files[-1]

    cache_data = None
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'rb') as f:
                cache_data = pickle.load(f)
            print("📦 Caché global cargado con éxito.")
        except Exception as e:
            print(f"⚠️ No se pudo leer el caché, se procesará desde cero: {e}")

    print("\n⏳ Paso 1: Ejecutando análisis de Auditoría (Con Fechas de CSV)...")
    df_auditoria, target_ola, processed_olas, audit_payload, nuevos_p1 = ejecutar_auditoria(pdf_files, csv_files[0], cache_data)

    print("⏳ Paso 2: Ejecutando análisis de Dosis Altas...")
    df_dosis = ejecutar_dosis_alta(ultimo_pdf, UMBRAL_DOSIS, codigos_dosis_alta_especial)

    print("🔄 Paso 3: Cruzando datos (Pedido + SKU + Cantidad)...")
    df_final = df_auditoria[df_auditoria["Ola"] == target_ola].copy()

    if not df_dosis.empty:
        for index, row_dosis in df_dosis.iterrows():
            pedido_buscar = str(row_dosis['Número de Pedido']).strip()
            sku_buscar = str(row_dosis['Código']).strip()
            cant_buscar = str(row_dosis['Cantidad']).strip()
            tipo_dosis = row_dosis.get('Tipo', 'Normal')
            etiqueta_dosis = "⚠️ ALTA DOSIS - SOLO 1" if tipo_dosis == 'Especial' else "⚠️ DOSIS ALTA"

            mascara = (df_final['Pedido Actual'].astype(str).str.strip() == pedido_buscar) & \
                      (df_final['SKU'].astype(str).str.strip() == sku_buscar) & \
                      (df_final['Cant'].astype(str).str.strip() == cant_buscar)

            for indice_coincidencia in df_final[mascara].index:
                estado_actual = df_final.at[indice_coincidencia, 'Estado']
                if "DOSIS ALTA" not in estado_actual and "ALTA DOSIS - SOLO 1" not in estado_actual:
                    if "APROBADO" in estado_actual:
                        df_final.at[indice_coincidencia, 'Estado'] = etiqueta_dosis
                    else:
                        df_final.at[indice_coincidencia, 'Estado'] = f"{estado_actual} | {etiqueta_dosis}"

    print("⏳ Paso 4: Extrayendo cantidades y cruzando memoria histórica...")
    memoria_cantidades, processed_cant_olas, nuevos_p4 = ejecutar_cantidades_alertas(pdf_files, cache_data)

    if nuevos_p1 > 0 or nuevos_p4 > 0:
        try:
            complete_cache_payload = {
                'processed_olas': processed_olas,
                'processed_cant_olas': processed_cant_olas,
                'memoria_cantidades_pedidos': memoria_cantidades,
                **audit_payload
            }
            with open(CACHE_FILE, 'wb') as f:
                pickle.dump(complete_cache_payload, f)
            print("💾 Caché global actualizado y guardado correctamente.")
        except Exception as e:
            print(f"⚠️ No se pudo guardar el archivo de caché: {e}")

    df_final['Cantidad'] = df_final['Cant']

    for idx, row in df_final.iterrows():
        ped_actual = str(row['Pedido Actual']).strip()
        sku_actual = str(row['SKU']).strip()
        ped_orig = str(row['Pedido Original']).strip()

        if ped_orig and ped_orig != "" and ped_orig != "Ver en SAP (Control)":
            if " & " in ped_orig:
                pedidos_individuales = [p.strip() for p in ped_orig.split(" & ")]
                pedidos_con_cant = []
                for p_item in pedidos_individuales:
                    clave_h = (p_item, sku_actual)
                    if clave_h in memoria_cantidades:
                        pedidos_con_cant.append(f"{p_item} ({memoria_cantidades[clave_h]})")
                    else:
                        pedidos_con_cant.append(p_item)
                df_final.at[idx, 'Pedido Original'] = " & ".join(pedidos_con_cant)
            else:
                clave_orig = (ped_orig, sku_actual)
                if clave_orig in memoria_cantidades:
                    cant_orig = memoria_cantidades[clave_orig]
                    df_final.at[idx, 'Pedido Original'] = f"{ped_orig} ({cant_orig})"

    columnas = list(df_final.columns)

    if 'Cant' in columnas:
        columnas.remove('Cant')
    if 'Cantidad' in columnas and 'SKU' in columnas:
        columnas.remove('Cantidad')
        idx_sku = columnas.index('SKU')
        columnas.insert(idx_sku + 1, 'Cantidad')

    if 'Pedido Actual' in columnas and 'Cantidad' in columnas:
        columnas.remove('Pedido Actual')
        idx_cant = columnas.index('Cantidad')
        columnas.insert(idx_cant + 1, 'Pedido Actual')

    df_final = df_final[columnas]

    # ==========================================
    # 🎨 PASO 5: GENERACIÓN DEL EXCEL CON FORMATO
    # ==========================================
    nombre_archivo = os.path.join(base_path, f"Reporte_Auditoria_Unificado_{target_ola}.xlsx")
    
    # Se genera directamente el archivo Excel local
    df_final.to_excel(nombre_archivo, index=False)

    try:
        wb = openpyxl.load_workbook(nombre_archivo)
        ws = wb.active

        fill_rojo_oscuro = PatternFill(start_color='800000', end_color='800000', fill_type='solid')
        font_blanco = Font(color='FFFFFF', bold=True)
        fill_morado = PatternFill(start_color='E6CCFF', end_color='E6CCFF', fill_type='solid')
        font_morado = Font(color='4B0082', bold=True)
        fill_amarillo = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
        font_amarillo = Font(color='806000', bold=True)
        fill_rojo_claro = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
        font_rojo_oscuro = Font(color='800000', bold=True)
        fill_azul = PatternFill(start_color='D0EBFF', end_color='D0EBFF', fill_type='solid')
        font_azul = Font(color='0B5ED7', bold=True)
        fill_gris = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
        font_negro = Font(color='000000', bold=True)

        header_row = [cell.value for cell in ws[1]]
        idx_estado = header_row.index('Estado') + 1 if 'Estado' in header_row else None

        if idx_estado:
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                estado_str = str(row[idx_estado - 1].value or "")

                es_duplicado = "Duplicado" in estado_str
                es_dosis = "DOSIS ALTA" in estado_str or "ALTA DOSIS - SOLO 1" in estado_str
                es_necesidad = "REVISAR NECESIDAD" in estado_str
                es_no_enviar = "NO ENVIAR" in estado_str

                total_alertas = sum([es_duplicado, es_dosis, es_necesidad, es_no_enviar])

                if es_no_enviar:
                    fill, font = fill_rojo_oscuro, font_blanco
                elif total_alertas >= 2:
                    fill, font = fill_morado, font_morado
                elif es_dosis:
                    fill, font = fill_amarillo, font_amarillo
                elif es_duplicado:
                    fill, font = fill_rojo_claro, font_rojo_oscuro
                elif es_necesidad:
                    fill, font = fill_azul, font_azul
                else:
                    fill, font = fill_gris, font_negro

                for cell in row:
                    cell.fill = fill
                    cell.font = font

        wb.save(nombre_archivo)
        print(f"\n📊 Archivo Excel generado con estilo y colores.")
    except Exception as e:
        print(f"⚠️ No se pudo aplicar formato completo al Excel: {e}")

    print("\n==================================================")
    print("✅ ¡PROCESO TERMINADO CON ÉXITO!")
    print(f"📁 Reporte guardado en: {nombre_archivo}")
    print("==================================================\n")
    
    input("Presiona ENTER para salir...")

# 🔥 PUNTO DE ENTRADA AL EJECUTAR EL EXECUTABLE
if __name__ == '__main__':
    unificar_reportes()
