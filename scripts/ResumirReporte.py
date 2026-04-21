# Script: ResumirReporte.py
# Autor: Fernando Cisneros Chavez (verdevenus23@gmail.com)
# Fecha: 21 de agosto de 2025
# Licencia: MIT

import subprocess
import sys

# Paquetes necesarios y su correspondencia con el nombre de importación
package_import_map = {
    "requests": "requests",
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "xlsxwriter": "xlsxwriter"
}

# Instalar silenciosamente si falta
for pip_name, import_name in package_import_map.items():
    try:
        __import__(import_name)
    except ImportError:
        print(f"📦 Instalando '{pip_name}'...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name, "--user"])

import requests
import pandas as pd
import time
from datetime import datetime
from datetime import timedelta
import sys
import glob

archivos = glob.glob(r"scripts\catalogo_empleados.xlsx")

if archivos:
    archivo_a_cargar = archivos[0]  # toma el primer archivo que encuentre
    df_empleados = pd.read_excel(archivo_a_cargar)
    print(f"📄 Archivo cargado: {archivo_a_cargar}")
else:
    archivo_a_cargar = "catalogo_empleados.xlsx"
    df_empleados = pd.DataFrame(columns=['_id'])
    print("❌ No se encontró ningún catalogo de empleados.")

# ============================ INTENTA CONECTAR Y ACTUALIZAR CATALOGO ============================
try:
    # login y obtención del token
    print("📡 Conectando a Kenjo...")
    url = "https://api.kenjo.io/api/v1/auth/login"
    payload = { "apiKey": "64efa6ae3b506b33130d07850b292be7131a50d6bbf1dccbe229a3b0" }
    headers = {
        "accept": "application/json",
        "content-type": "application/json"
    }
    acceso = requests.post(url, json=payload, headers=headers)
    acceso.raise_for_status()

    credenciales = acceso.json()
    token = credenciales["token"]
    headers = {"accept": "application/json", "Authorization": token}

    # Sucursales
    url = "https://api.kenjo.io/api/v1/offices"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    oficinas = response.json()
    df_oficinas = pd.DataFrame(oficinas)[['_id', 'name']].rename(columns={'_id': 'officeId'})

    # Departamentos
    url = "https://api.kenjo.io/api/v1/departments"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    departamentos = response.json()
    df_departamentos = pd.DataFrame(departamentos)[['_id', 'name']].rename(columns={'_id': 'departmentId'})

    # Empleados
    print(f"📥 Actualizando lista de empleados")
    url = "https://api.kenjo.io/api/v1/employees"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    registro_empleados = response.json()
    empleados_api = registro_empleados.get("data", [])

    ids_existentes = df_empleados['id'].tolist() if not df_empleados.empty else []
    if df_empleados.empty:
        df_empleados.rename(columns={"_id": "id"}, inplace=True)

    empleados_detalle = []
    for empleado in empleados_api:
        _id = empleado.get("_id")
        if _id not in ids_existentes:
            url_detalle = f"https://api.kenjo.io/api/v1/employees/{_id}"
            try:
                response = requests.get(url_detalle, headers=headers)
                response.raise_for_status()
                data = response.json()
                empleado_info = {
                    "id": data["account"].get("_id", None),
                    "Activo": data["account"].get("isActive", None),
                    "Nombre": data["personal"].get("displayName", None),
                    "Numero de colaborador": data["personal"].get("c_NumerodeColaborador", None),
                    "Puesto": data["work"].get("jobTitle", None),
                    "officeId": data["work"].get("officeId", None),
                    "departmentId": data["work"].get("departmentId", None)
                }
                empleados_detalle.append(empleado_info)
            except requests.exceptions.RequestException as e:
                print(f"⚠️ Error con el ID {_id}: {e}")
            time.sleep(0.3)

    df_nuevos = pd.DataFrame(empleados_detalle)
    if df_nuevos.empty or len(df_nuevos.columns) == 0:
        df_nuevos = pd.DataFrame(columns=['id','Activo','Nombre','Numero de colaborador','Puesto','officeId','departmentId'])

    # Enriquecer con oficinas y departamentos
    df_nuevos_completo = df_nuevos.merge(df_oficinas, on="officeId", how="left")
    df_nuevos_completo.drop("officeId", axis=1, inplace=True)
    df_nuevos_completo.rename(columns={"name": "Oficina"}, inplace=True)

    df_nuevos_completo = df_nuevos_completo.merge(df_departamentos, on="departmentId", how="left")
    df_nuevos_completo.drop("departmentId", axis=1, inplace=True)
    df_nuevos_completo.rename(columns={"name": "Departamento"}, inplace=True)

    # Concatenar con lo existente
    df_empleados_completo = pd.concat([df_empleados, df_nuevos_completo], ignore_index=True).drop_duplicates('id')

    # Guardar archivo actualizado
    print("\n📊 Procesando catalogo en Excel...")
    try:
        with pd.ExcelWriter(archivo_a_cargar, engine='xlsxwriter') as writer:
            df_empleados_completo.to_excel(writer, sheet_name='empleados', index=False)
        print(f"✅ Catálogo actualizado: {archivo_a_cargar}")
    except Exception as e:
        print("❌ Error al generar catalogo en Excel.")
        print(f"🛠️ Detalles del error: {e}")

except requests.exceptions.RequestException as e:
    print("⚠️ No se pudo conectar a Kenjo o hubo un error en la red.")
    print(f"🔌 Detalles del error: {e}")
    print("🔁 Se continuará usando el catálogo local existente.")

archivo = sys.argv[1]

if archivo:
    df_reporte = pd.read_excel(archivo)
    df_reporte2 = df_reporte.copy()
    print(f"Archivo cargado: {archivo}")
else:
    print("❌ No se selecciono ningun archivo valido.")
# Realizando operaciones entre tablas =============================================================
df_reporte['Fecha'] = pd.to_datetime(df_reporte['Fecha'])
df_reporte['Hora de inicio'] = pd.to_datetime(
    df_reporte['Hora de inicio'], format='%H:%M:%S', errors='coerce')
df_reporte['Hora de fin'] = pd.to_datetime(
    df_reporte['Hora de fin'], format='%H:%M:%S', errors='coerce')
# Calculamos diferencia en horas (fin - inicio)
df_reporte['horas_trabajadas'] = (df_reporte['Hora de fin'] - df_reporte['Hora de inicio']).dt.total_seconds() / 3600

# Si horas_trabajadas es NaN (porque fin era NaT), poner 0
df_reporte['horas_trabajadas'] = df_reporte['horas_trabajadas'].fillna(0)

# Aplicamos redondeo hacia abajo si la parte decimal es menor a 0.5
def redondear_horas(h):
    entero = int(h)
    decimal = h - entero
    if decimal < 0.5:
        return entero
    else:
        return h
    
df_reporte['horas_trabajadas'] = df_reporte['horas_trabajadas'].apply(redondear_horas).round(2)
df_reporte['Tiempo de pausa'] = df_reporte['Tiempo de pausa'] / 60
df_reporte['horas_trabajadas'] = df_reporte['horas_trabajadas'] - df_reporte['Tiempo de pausa']
df_reporte['horas_extra'] = (df_reporte['horas_trabajadas'] - df_reporte['Total turno']).clip(lower=0)
df_reporte['faltas'] = (df_reporte['Hora de inicio'].isnull()).astype(int)
df_reporte['asistencias'] = df_reporte['Hora de inicio'].notnull()

pd.options.display.max_columns = None

catalogo = glob.glob(r"scripts\catalogo_empleados.xlsx")

if catalogo:
    archivo_a_cargar = catalogo[0]  # toma el primer archivo que encuentre
    df_catalogo = pd.read_excel(archivo_a_cargar)
    print(f"Catalogo cargado: {archivo_a_cargar}")
else:
    archivo_a_cargar = "catalogo_empleados.xlsx"
    df_catalogo = pd.DataFrame(columns=['_id'])
    print("❌ No se encontró ningún catalogo de empleados.")

df_catalogo['Nombre_Oficina'] = df_catalogo['Nombre'] + ' - ' + df_catalogo['Oficina']
df_reporte['Nombre_Oficina'] = df_reporte['Nombre'] + ' - ' + df_reporte['Oficina']
df_reporte = df_reporte.merge(df_catalogo, on="Nombre_Oficina", how="left")

dias_laborados = (
    df_reporte[df_reporte['asistencias']]
    .drop_duplicates(subset=['Numero de colaborador', 'Fecha'])
    .groupby('Numero de colaborador')
    .size()
    .rename('dias_laborados')
)

df_resumen = df_reporte.groupby('Numero de colaborador').agg({
    'Nombre_x': 'first',
    'Puesto': 'first',
    'Oficina_x': 'first',
    'Departamento_x': 'first',
    'Fecha': ['min', 'max'],
    'Total turno': 'sum',
    'horas_trabajadas': 'sum',
    'horas_extra': 'sum',
    'faltas':'sum'
})
# Aplanamos columnas multi-índice generadas por agg
df_resumen.columns = ['_'.join(col).strip() if isinstance(col, tuple) else col for col in df_resumen.columns.values]
df_resumen = df_resumen.merge(dias_laborados.rename('dias_laborados'), on='Numero de colaborador', how='left')
df_resumen['dias_laborados'] = df_resumen['dias_laborados'].astype('Int64')
# Renombramos
df_resumen = df_resumen.rename(columns={
    'Nombre_x_first': 'Nombre',
    'Puesto_first': 'Puesto',
    'Oficina_x_first': 'Oficina',
    'Departamento_x_first': 'Departamento',
    'Fecha_min': 'fecha inicial',
    'Fecha_max': 'fecha final',
    'Total turno_sum': 'horas asignadas',
    'horas_trabajadas_sum': 'horas trabajadas',
    'horas_extra_sum': 'horas extra',
    'faltas_sum': 'faltas'
})

# Calculamos días del periodo (fecha_final - fecha_inicial)
fin_periodo = df_resumen['fecha final'].dropna().max()
inicio_periodo = df_resumen['fecha inicial'].dropna().min()
df_resumen['dias_periodo'] = (fin_periodo - inicio_periodo +timedelta(days=1)).days

#df_resumen['Fecha'] = df_resumen['Fecha'].dt.date

df_resumen['fecha inicial'] = df_resumen['fecha inicial'].dt.date
df_resumen['fecha final'] = df_resumen['fecha final'].dt.date
df_resumen['faltas'] = df_resumen['faltas'].replace(0, None)
df_resumen = df_resumen.reset_index()

# Creacion del excel ==============================================================================
print("\n📊 Procesando archivo Excel...")
inicio_str = inicio_periodo.strftime('%Y-%m-%d')
fin_str = fin_periodo.strftime('%Y-%m-%d')
try:
    archivo = f"cotejo_asistencias_{inicio_str}-{fin_str}.xlsx"
    with pd.ExcelWriter(archivo, engine='xlsxwriter') as writer:
        df_resumen.to_excel(writer, sheet_name='resumen', index=False)
        df_reporte2.to_excel(writer, sheet_name='registro horario', index=False)

    print(f"✅ Resumen horario generado: {archivo}")
except PermissionError:
    print("❌ Error: No se pudo guardar el archivo. ¿Está abierto en Excel?")

input("\nPresiona Enter para salir...")