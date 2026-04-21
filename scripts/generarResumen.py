# Script: GenerarResumen.py
# Autor: Fernando Cisneros Chavez (verdevenus23@gmail.com)
# Fecha: 20 de agosto de 2025
# Licencia: MIT
# Actualizacion 27 de octubre, se elimina el campo break de https://api.kenjo.io/api/v1/attendances

import subprocess
import sys
import math

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
        print(f"Instalando '{pip_name}'...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name, "--user"])

import requests
import pandas as pd
import time
from datetime import datetime
from datetime import timedelta
import glob
import tkinter as tk
from tkinter import ttk, messagebox
import threading

# Códigos de salida del proceso
EXIT_OK             = 0   # Éxito
EXIT_ARGS           = 1   # Argumentos inválidos
EXIT_RED            = 2   # Error de red / API
EXIT_PERMISO        = 3   # Archivo Excel abierto o sin permisos
EXIT_ERROR_GENERAL  = 4   # Cualquier otro error inesperado

# Validación de argumentos al inicio
def _error_y_salir(msg, codigo=EXIT_ARGS):
    """Muestra un messagebox de error y termina el proceso con el código dado."""
    _root = tk.Tk()
    _root.withdraw()
    messagebox.showerror("Error", msg)
    _root.destroy()
    sys.exit(codigo)

if len(sys.argv) < 3:
    _error_y_salir(
        "Uso correcto:\n  python GenerarResumen.py YYYY-MM-DD YYYY-MM-DD\n\n"
        "Ejemplo:\n  python GenerarResumen.py 2025-10-01 2025-10-31"
    )

try:
    fecha_inicio = datetime.strptime(sys.argv[1], "%Y-%m-%d").date()
    fecha_fin    = datetime.strptime(sys.argv[2], "%Y-%m-%d").date()
except ValueError:
    _error_y_salir("Formato de fecha inválido.\nUse YYYY-MM-DD  (ej: 2025-10-01).")

if fecha_fin < fecha_inicio:
    _error_y_salir("La fecha final no puede ser anterior a la fecha inicial.")

# Modo depuración — se activa pasando --debug como tercer argumento
# Ejemplo: python GenerarResumen.py 2025-10-01 2025-10-31 --debug
DEBUG = "--debug" in sys.argv

# Ventana de progreso ============================================================================
def crear_ventana_progreso(titulo, pasos):
    """Crea y devuelve una ventana de progreso con barra indeterminada."""
    raiz = tk.Tk()
    raiz.title(titulo)
    raiz.resizable(False, False)
    raiz.attributes("-topmost", True)

    ancho, alto = 380, 120
    x = (raiz.winfo_screenwidth()  - ancho) // 2
    y = (raiz.winfo_screenheight() - alto)  // 2
    raiz.geometry(f"{ancho}x{alto}+{x}+{y}")

    lbl_estado = tk.Label(raiz, text="Iniciando...", font=("Segoe UI", 10), pady=8)
    lbl_estado.pack()

    barra = ttk.Progressbar(raiz, mode="determinate", maximum=pasos, length=340)
    barra.pack(pady=4)

    lbl_detalle = tk.Label(raiz, text="", font=("Segoe UI", 8), fg="gray")
    lbl_detalle.pack()

    return raiz, lbl_estado, barra, lbl_detalle

def avanzar(estado, lbl_estado, barra, lbl_detalle, detalle=""):
    """Actualiza la barra de progreso. En modo DEBUG solo imprime en consola."""
    if DEBUG:
        print(f"  [{estado}] {detalle}".rstrip())
        return
    lbl_estado.config(text=estado)
    lbl_detalle.config(text=detalle)
    barra.step(1)
    barra.update()
    lbl_estado.update()
    lbl_detalle.update()

if DEBUG:
    # En modo debug no se crea ventana — la consola muestra todo
    ventana = lbl_estado = barra = lbl_detalle = None
    print("=" * 50)
    print("  MODO DEPURACIÓN ACTIVO")
    print(f"  Periodo: {sys.argv[1]} → {sys.argv[2]}")
    print("=" * 50)
else:
    PASOS_TOTALES = 8
    ventana, lbl_estado, barra, lbl_detalle = crear_ventana_progreso(
        f"Generando reporte de asistencias", PASOS_TOTALES
    )

try:
    hoy = datetime.today().strftime('%d-%m-%Y')

    archivos = glob.glob(r"scripts\catalogo_empleados.xlsx")

    if archivos:
        archivo_a_cargar = archivos[0]
        df_empleados = pd.read_excel(archivo_a_cargar)
        # Normalizamos el nombre de la columna independientemente de cómo venga en el archivo
        if '_id' in df_empleados.columns and 'id' not in df_empleados.columns:
            df_empleados.rename(columns={'_id': 'id'}, inplace=True)
        print(f"Archivo cargado: {archivo_a_cargar}")
    else:
        import os
        archivo_a_cargar = os.path.join(os.getcwd(), "catalogo_empleados.xlsx")
        df_empleados = pd.DataFrame(columns=['id', 'Activo', 'Nombre', 'Numero de colaborador',
                                              'Puesto', 'officeId', 'departmentId'])
        print("⚠️  No se encontró ningún catálogo de empleados. Se creará uno nuevo.")

    avanzar("Catálogo cargado", lbl_estado, barra, lbl_detalle)

    # login y obtencion del token =====================================================================
    avanzar("Conectando a Kenjo...", lbl_estado, barra, lbl_detalle)
    print("📡 Conectando a kenjo...")
    url = "https://api.kenjo.io/api/v1/auth/login"

    payload = { "apiKey": "1234567890qwertyuiopadfghjklñzxcvbnm" }
    headers = {
        "accept": "application/json",
        "content-type": "application/json"
    }

    try:
        acceso = requests.post(url, json=payload, headers=headers, timeout=15)
        acceso.raise_for_status()
        credenciales = acceso.json()
        if "token" not in credenciales:
            raise ValueError(f"La API no devolvió un token. Respuesta: {credenciales}")
        token = credenciales["token"]
    except requests.exceptions.ConnectionError:
        raise ConnectionError("No se pudo conectar con Kenjo.\n¿Hay conexión a internet?")
    except requests.exceptions.Timeout:
        raise TimeoutError("La conexión con Kenjo tardó demasiado.\nIntenta de nuevo en unos momentos.")
    except requests.exceptions.HTTPError as e:
        status = e.response.status_code if e.response is not None else "?"
        if status == 401:
            raise PermissionError("Credenciales de API inválidas (401).\nVerifica la apiKey en el script.")
        raise ConnectionError(f"Error HTTP {status} al autenticar con Kenjo:\n{e}")
    except ValueError as e:
        raise RuntimeError(str(e))

    headers = {"accept": "application/json", "Authorization": token}

    # sucursales ======================================================================================
    url = "https://api.kenjo.io/api/v1/offices"
    response = requests.get(url, headers=headers, timeout=15)
    response.raise_for_status()
    oficinas = response.json()

    df_oficinas = pd.DataFrame(oficinas)[['_id', 'name']]
    df_oficinas = df_oficinas.rename(columns={'_id': 'officeId'})

    # departamentos ===================================================================================
    url = "https://api.kenjo.io/api/v1/departments"
    response = requests.get(url, headers=headers, timeout=15)
    response.raise_for_status()
    departamentos = response.json()

    df_departamentos = pd.DataFrame(departamentos)[['_id', 'name']]
    df_departamentos = df_departamentos.rename(columns={'_id': 'departmentId'})

    # empleados =======================================================================================
    avanzar("Descargando empleados...", lbl_estado, barra, lbl_detalle)
    print(f"Actualizando lista de empleados")
    url = "https://api.kenjo.io/api/v1/employees"
    response = requests.get(url, headers=headers)
    registro_empleados = response.json()
    empleados_api = registro_empleados.get("data", [])

    # Sincronizar campo Activo en el catálogo ya guardado usando la lista general (sin llamadas extra)
    ids_activos = {e["_id"] for e in empleados_api if e.get("isActive", False)}
    if not df_empleados.empty and 'id' in df_empleados.columns:
        df_empleados["Activo"] = df_empleados["id"].isin(ids_activos)

    # Solo consultar empleados activos y que aún no estén en el catálogo
    ids_existentes = df_empleados['id'].tolist() if not df_empleados.empty else []
    empleados_nuevos_activos = [e for e in empleados_api
                                 if e.get("isActive", False) and e.get("_id") not in ids_existentes]

    total_nuevos = len(empleados_nuevos_activos)
    print(f"   → {total_nuevos} empleado(s) nuevo(s) activo(s) para descargar")

    empleados_detalle = []

    # Descargamos solo los empleados activos que NO estén en el catálogo
    for i, empleado in enumerate(empleados_nuevos_activos, 1):
        _id = empleado.get("_id")
        url_detalle = f"https://api.kenjo.io/api/v1/employees/{_id}"
        try:
            response = requests.get(url_detalle, headers=headers)
            response.raise_for_status()
            data = response.json()

            empleado_info = {
                "id": data["account"].get("_id", None),
                "Activo": True,
                "Nombre": data.get("financial", {}).get("accountHolderName") or data["personal"].get("displayName", None),
                "Numero de colaborador": data["personal"].get("c_NumerodeColaborador", None),
                "Puesto": data["work"].get("jobTitle", None),
                "officeId": data["work"].get("officeId", None),
                "departmentId": data["work"].get("departmentId", None)
            }
            empleados_detalle.append(empleado_info)
            print(f"   [{i}/{total_nuevos}] {empleado_info['Nombre']}")
        except requests.exceptions.RequestException as e:
            print(f"⚠️ Error con el ID {_id}: {e}")

        time.sleep(0.3)

    df_nuevos = pd.DataFrame(empleados_detalle)
    if df_nuevos.empty or len(df_nuevos.columns) == 0:
        df_nuevos = pd.DataFrame(columns=['id', 'Activo', 'Nombre', 'Numero de colaborador',
                                           'Puesto', 'officeId', 'departmentId'])

    # Realizando operaciones entre tablas =============================================================
    df_nuevos_completo = df_nuevos.merge(df_oficinas, on="officeId", how="left")
    df_nuevos_completo.drop("officeId", axis=1, inplace=True)
    df_nuevos_completo.rename(columns={"name": "Oficina"}, inplace=True)
    df_nuevos_completo = df_nuevos_completo.merge(df_departamentos, on="departmentId", how="left")
    df_nuevos_completo.drop("departmentId", axis=1, inplace=True)
    df_nuevos_completo.rename(columns={"name": "Departamento"}, inplace=True)

    df_empleados_completo = pd.concat([df_empleados, df_nuevos_completo], ignore_index=True).drop_duplicates('id')

    # Creacion del excel ==============================================================================
    print("\nProcesando catalogo en Excel...")
    try:
        with pd.ExcelWriter(archivo_a_cargar, engine='xlsxwriter') as writer:
            df_empleados_completo.to_excel(writer, sheet_name='empleados', index=False)
        print(f"✅ Catálogo actualizado: {archivo_a_cargar}")
    except PermissionError:
        # No abortamos el proceso completo — solo avisamos y continuamos con los datos en memoria
        aviso = (
            "No se pudo actualizar el catálogo de empleados.\n"
            "¿Está abierto el archivo en Excel?\n\n"
            f"Archivo: {archivo_a_cargar}\n\n"
            "El reporte de asistencias se generará de todas formas."
        )
        print(f"⚠️  {aviso}")
        if not DEBUG:
            messagebox.showwarning("Advertencia — catálogo", aviso)
    except Exception as e:
        print(f"⚠️  Error al guardar catálogo: {e}")

    # horas esperadas =================================================================================
    avanzar("Descargando horarios asignados...", lbl_estado, barra, lbl_detalle)
    print(f"Descargando horarios asignados")
    url_base = "https://api.kenjo.io/api/v1/attendances/expected-time"

    params = {
        "from": fecha_inicio,
        "to": fecha_fin,
        "offset": 1,
        "limit": 100
    }

    all_results = []

    while True:
        response = requests.get(url_base, headers=headers, params=params, timeout=20)
        if response.status_code == 400:
            break
        response.raise_for_status()
        data = response.json()
        results = data.get("data") or data
        all_results.extend(results)
        if len(results) < params["limit"]:
            break
        params["offset"] += 1

    filas = []

    for usuario in all_results:
        user_id = usuario["userId"]
        for dia in usuario["days"]:
            filas.append({
                "id": user_id,
                "fecha": dia["date"],
                "horas_asignadas": dia["expectedHours"],
                "minutos": dia["expectedMinutes"]
            })

    df_horas_esperadas = pd.DataFrame(filas)
    df_horas_esperadas['fecha'] = pd.to_datetime(df_horas_esperadas['fecha']).dt.date

    # asistencias registradas =========================================================================
    avanzar("Descargando asistencias registradas...", lbl_estado, barra, lbl_detalle)
    print(f"Descargando asistencias registradas")

    url = f"https://api.kenjo.io/api/v1/attendances?from={fecha_inicio}&to={fecha_fin}"
    response = requests.get(url, headers=headers)
    horas_registradas = response.json()

    COLUMNAS_ASISTENCIA = ["userId", "startTime", "endTime", "breakTime", "comment"]
    df_asistencias = pd.DataFrame(horas_registradas).reindex(columns=COLUMNAS_ASISTENCIA)
    df_asistencias.rename(columns={"userId": "id"}, inplace=True)
    df_asistencias['startTime'] = pd.to_datetime(df_asistencias['startTime'], errors='coerce')
    df_asistencias['fecha'] = df_asistencias['startTime'].dt.date

    # uniendo horas esperadas y registradas ===========================================================
    df_unido = pd.merge(df_horas_esperadas, df_asistencias, on=['id', 'fecha'], how='outer')
    df_unido['horas_asignadas'] = df_unido['horas_asignadas'].fillna(0)

    df_unido['fecha'] = df_unido['fecha'].combine_first(
        pd.to_datetime(df_unido['startTime'], errors='coerce').dt.date
            .apply(lambda x: pd.NaT if pd.isnull(x) else x)
    )

    for col in ['startTime', 'endTime']:
        df_unido[col] = pd.to_datetime(df_unido[col], utc=True, errors='coerce')
        df_unido[col] = df_unido[col].dt.tz_localize(None)
        df_unido[col] = df_unido[col].astype(str)

    #    Una falta = empleado con horas asignadas pero sin registro de entrada
    df_unido['faltas'] = (
        (df_unido['horas_asignadas'] > 0) & df_unido['startTime'].isnull()
    ).astype(int)

    # Ahora sí filtramos filas completamente vacías (sin fecha ni registro)
    df_unido = df_unido[df_unido['fecha'].notnull()]
    df_unido.reset_index(drop=True, inplace=True)
    df_unido.rename(columns={'startTime': 'inicio', 'endTime': 'fin', 'breakTime': 'descanso'}, inplace=True)
    df_unido = df_unido.fillna(0)

    # Realizando operaciones entre tablas =============================================================
    df_unido = df_unido.merge(df_empleados_completo, on="id", how="left")

    # Filtrar filas sin datos utiles:
    # Se excluyen unicamente registros sin nombre en catalogo Y sin horas asignadas Y sin entrada.
    # Esto preserva empleados dados de baja que registraron horas o periodos en que estaban activos.
    sin_nombre    = df_unido['Nombre'].isnull() | (df_unido['Nombre'] == 0)
    sin_asignadas = df_unido['horas_asignadas'] == 0
    sin_entrada   = df_unido['inicio'].isnull() | (df_unido['inicio'] == 0)
    ruido         = sin_nombre & sin_asignadas & sin_entrada

    n_excluidos = df_unido[ruido]['id'].nunique()
    if n_excluidos > 0:
        print(f"   -> Excluyendo {n_excluidos} empleado(s) sin datos en el periodo")
    df_unido = df_unido[~ruido]

    df_unido = df_unido[['Numero de colaborador', 'Nombre', 'Puesto', 'Oficina', 'Departamento',
                          'fecha', 'horas_asignadas', 'inicio', 'fin', 'descanso', 'faltas']]

    df_unido['fecha']  = pd.to_datetime(df_unido['fecha'])
    df_unido['inicio'] = pd.to_datetime(df_unido['inicio'], errors='coerce')
    df_unido['fin']    = pd.to_datetime(df_unido['fin'],    errors='coerce')

    def redondear_horas(h):
        """Trunca hacia abajo si la parte decimal es < 0.5, de lo contrario conserva dos decimales."""
        if pd.isna(h):
            return 0.0
        entero  = math.floor(h)
        decimal = h - entero
        return float(entero) if decimal < 0.5 else round(h, 2)

    df_unido['horas_trabajadas'] = (df_unido['fin'] - df_unido['inicio']).dt.total_seconds() / 3600
    df_unido['horas_trabajadas'] = df_unido['horas_trabajadas'].fillna(0).apply(redondear_horas)

    df_unido['descanso']        = pd.to_numeric(df_unido['descanso'], errors='coerce').fillna(0) / 60
    df_unido['horas_trabajadas'] = df_unido['horas_trabajadas'] - df_unido['descanso']
    df_unido['horas_extra']      = (df_unido['horas_trabajadas'] - df_unido['horas_asignadas']).clip(lower=0)
    df_unido['asistencias']      = df_unido['inicio'].notnull()

    dias_laborados = (
        df_unido[df_unido['asistencias']]
        .drop_duplicates(subset=['Numero de colaborador', 'fecha'])
        .groupby('Numero de colaborador')
        .size()
        .rename('dias_laborados')
    )

    df_resumen = df_unido.groupby('Numero de colaborador').agg({
        'Nombre':          'first',
        'Puesto':          'first',
        'Oficina':         'first',
        'Departamento':    'first',
        'fecha':           ['min', 'max'],
        'horas_asignadas': 'sum',
        'horas_trabajadas':'sum',
        'horas_extra':     'sum',
        'faltas':          'sum'
    })

    df_resumen.columns = ['_'.join(col).strip() if isinstance(col, tuple) else col
                          for col in df_resumen.columns.values]
    df_resumen = df_resumen.merge(dias_laborados.rename('dias_laborados'), on='Numero de colaborador', how='left')
    df_resumen['dias_laborados'] = df_resumen['dias_laborados'].astype('Int64')

    df_resumen = df_resumen.rename(columns={
        'Nombre_first':          'Nombre',
        'Puesto_first':          'Puesto',
        'Oficina_first':         'Oficina',
        'Departamento_first':    'Departamento',
        'fecha_min':             'fecha inicial',
        'fecha_max':             'fecha final',
        'horas_asignadas_sum':   'horas asignadas',
        'horas_trabajadas_sum':  'horas trabajadas',
        'horas_extra_sum':       'horas extra',
        'faltas_sum':            'faltas'
    })

    fin_periodo    = df_unido['fin'].dropna().max()
    inicio_periodo = df_unido['inicio'].dropna().min()
    df_resumen['dias_periodo'] = (fin_periodo - inicio_periodo + timedelta(days=1)).days

    df_unido['inicio'] = df_unido['inicio'].dt.time
    df_unido['fin']    = df_unido['fin'].dt.time
    df_unido['fecha']  = df_unido['fecha'].dt.date

    df_resumen['fecha inicial'] = df_resumen['fecha inicial'].dt.date
    df_resumen['fecha final']   = df_resumen['fecha final'].dt.date

    df_resumen = df_resumen.reset_index()

    # Creacion del excel ==============================================================================
    avanzar("Generando archivo Excel...", lbl_estado, barra, lbl_detalle)
    print("\nProcesando archivo Excel...")
    try:
        archivo = f"cotejo_asistencias_{fecha_inicio}-{fecha_fin}.xlsx"
        with pd.ExcelWriter(archivo, engine='xlsxwriter') as writer:
            df_resumen.to_excel(writer, sheet_name='resumen', index=False)
            df_unido.to_excel(writer, sheet_name='registro horario', index=False)
        avanzar("¡Listo!", lbl_estado, barra, lbl_detalle, archivo)
        print(f"✅ Resumen horario generado: {archivo}")
    except PermissionError:
        raise PermissionError(
            f"No se pudo guardar el archivo de reporte:\n  {archivo}\n\n"
            "¿Está abierto en Excel?\n"
            "Ciérralo e intenta de nuevo."
        )
except PermissionError as e:
    codigo_salida = EXIT_PERMISO
    if DEBUG:
        import traceback; traceback.print_exc()
        input("\nPresiona Enter para cerrar...")
    else:
        ventana.withdraw()
        messagebox.showerror("Error al guardar", str(e))
    sys.exit(codigo_salida)
except (ConnectionError, TimeoutError) as e:
    codigo_salida = EXIT_RED
    if DEBUG:
        import traceback; traceback.print_exc()
        input("\nPresiona Enter para cerrar...")
    else:
        ventana.withdraw()
        messagebox.showerror("Error de conexión", str(e))
    sys.exit(codigo_salida)
except Exception as e:
    codigo_salida = EXIT_ERROR_GENERAL
    import traceback
    if DEBUG:
        print("\n❌ Error inesperado:")
        traceback.print_exc()
        input("\nPresiona Enter para cerrar...")
    else:
        ventana.withdraw()
        tb = traceback.format_exc()
        messagebox.showerror(
            "Error inesperado",
            f"Ocurrió un error durante el proceso:\n\n{tb}"
        )
    sys.exit(codigo_salida)
else:
    if DEBUG:
        print("\n✅ Proceso completado:", archivo)
        input("\nPresiona Enter para cerrar...")
    else:
        ventana.withdraw()
        messagebox.showinfo(
            "Proceso completado",
            f"Reporte generado exitosamente:\n\n{archivo}"
        )
    sys.exit(EXIT_OK)
finally:
    if not DEBUG and ventana:
        ventana.destroy()
