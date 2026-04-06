import pyautogui # interacts with the UI
import time # controls timing between actions
import subprocess # launches apps
import pandas as pd # processes business data (likely invoices/movements)
import pyperclip # clipboard
import io # move structured data between systems
import traceback # robust error debugging
import getpass # secure login handling
from datetime import datetime # logging + audit trail

# ===== CONFIG =====
IMAGES_PATH = "images/"
CONFIDENCE = 0.6

# ===== APP AUTH =====
APP_USER = "admin"
APP_PASSWORD = "1914"

USUARIO = "mauricio"
PASSWORD = "1234"

pyautogui.FAILSAFE = True

# ===== HELPERS =====
def get_date_input():
    print("\n📅 Ingrese el rango de fechas (formato: DD-MM-YYYY)")
    print("Ejemplo: 01-01-2026")
    while True:
        fecha_inicio_input = input("Fecha inicio: ")
        fecha_fin_input = input("Fecha fin: ")
        try:
            # Validate using hyphens
            inicio = datetime.strptime(fecha_inicio_input, "%d-%m-%Y")
            fin = datetime.strptime(fecha_fin_input, "%d-%m-%Y")
            if inicio > fin:
                print("❌ 'Fecha inicio' no puede ser mayor que 'Fecha fin'\n")
                continue
            # Convert to required format with slashes
            fecha_inicio = inicio.strftime("%d/%m/%Y")
            fecha_fin = fin.strftime("%d/%m/%Y")
            print("✅ Fechas válidas\n")
            return fecha_inicio, fecha_fin
        except ValueError:
            print("❌ Formato inválido. Use DD-MM-YYYY\n")

def wait_and_click(image, timeout=10, confidence=CONFIDENCE, fallback=None):
    start = time.time()
    while time.time() - start < timeout:
        try:
            location = pyautogui.locateCenterOnScreen(
                IMAGES_PATH + image,
                confidence=confidence,
                grayscale=True
            )
            if location:
                pyautogui.click(location)
                print(f"✅ Click realizado (imagen): {image}")
                return True
        except:
            pass
        time.sleep(1)
    if fallback:
        x, y = fallback
        pyautogui.click(x, y)
        print(f"⚠️ Usando coordenadas para {image} en ({x},{y})")
        return True
    raise Exception(f"❌ Elemento no encontrado: {image}")

def wait_until_visible(image, timeout=60, confidence=CONFIDENCE):
    start = time.time()
    while time.time() - start < timeout:
        if pyautogui.locateOnScreen(IMAGES_PATH + image, confidence=confidence):
            print(f"👁️ Elemento visible: {image}")
            return True
        time.sleep(1)
    raise Exception(f"⏱️ Tiempo de espera agotado para: {image}")

def wait_for_clipboard_change(timeout=30):
    old = pyperclip.paste()
    start = time.time()
    while time.time() - start < timeout:
        new = pyperclip.paste()
        if new != old and new.strip():
            print("📋 Portapapeles actualizado")
            return new
        time.sleep(1)
    raise Exception("⏱️ El portapapeles no se actualizó")

def app_login():
    print("=== 🔐 INICIO DE SESIÓN ===")
    for intento in range(3):  # máximo 3 intentos
        usuario = input("Usuario: ")
        password = getpass.getpass("Contraseña: ")
        if usuario == APP_USER and password == APP_PASSWORD:
            print("✅ Acceso concedido\n")
            return True
        else:
            print(f"❌ Credenciales inválidas ({intento + 1}/3)\n")
    raise Exception("🚫 Demasiados intentos fallidos")

def show_menu():
    print("=== 📋 MENÚ PRINCIPAL ===")
    print("1. Facturas y movimiento de inventario - Documentos con artículos")
    print("0. Salir")
    while True:
        opcion = input("Seleccione una opción: ")
        if opcion in ["1", "0"]:
            return opcion
        else:
            print("❌ Opción inválida. Intente nuevamente.\n")

def login_to_dapas():
    print("🚀 Abriendo DAPAS...")

    subprocess.Popen(["mstsc", r"C:\Users\ASUS\OneDrive\Escritorio\DAPAS.rdp"])
    
    # ===== 1. ESPERAR CARGA =====
    print("⏳ Esperando que SAG cargue (10s)...")
    time.sleep(10)

    # ===== 2. SELECCIONAR BASE DE DATOS =====
    print("🗂️ Seleccionando base de datos...")
    for _ in range(2): 
        pyautogui.press("down")
        time.sleep(0.1)

    # ===== 3. IR A USUARIO =====
    print("👤 Navegando al campo de usuario...")
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.1)

    pyautogui.write(USUARIO)
    time.sleep(0.2)

    # ===== 4. IR A CONTRASEÑA =====
    print("🔑 Ingresando contraseña...")
    pyautogui.press("tab")
    time.sleep(0.1)

    pyautogui.write(PASSWORD)
    time.sleep(0.2)

    # ===== 5. ENVIAR LOGIN =====
    print("📤 Enviando credenciales...")
    pyautogui.press("enter")

    # ===== 6. MANEJAR MODAL =====
    print("⚠️ Gestionando mensaje de sesión activa...")
    # wait_until_visible("ok.png")
    time.sleep(5)

    pyautogui.press("enter")
    time.sleep(0.1)
    pyautogui.press("enter")

    print("✅ Inicio de sesión completado")

# ===== NAVIGATION =====
def navigate_to_report():
    print("📊 Navegando al módulo de informes...")

    wait_and_click("menu_informes.png", confidence=0.8, fallback=(150, 33))
    time.sleep(0.5)

    print("📂 Accediendo a 'Varios'...")
    pyautogui.press("down")   # Varios
    time.sleep(0.1)

    pyautogui.press("right")
    time.sleep(0.1)

    print("📄 Seleccionando '05 - Facturas y movimientos'...")
    for _ in range(4):  # 05
        pyautogui.press("down")
        time.sleep(0.1)

    pyautogui.press("enter")

    print("✅ Módulo de informe abierto")

# ===== REPORT CONFIG =====
# ===== CONFIGURACIÓN DEL INFORME =====
def configure_report(fecha_inicio, fecha_fin):
    print("⚙️ Configurando el informe...")

    # Seleccionar "05-B. OTROS"
    print("📄 Seleccionando tipo de informe: 05-B. OTROS...")
    wait_and_click("otros_05b.png", confidence=0.8, fallback=(70, 200))
    time.sleep(1)

    # Abrir dropdown de Bodega
    print("🏭 Abriendo selección de bodega...")
    wait_and_click("bodega_dropdown.png", confidence=0.8, fallback=(995, 340))
    time.sleep(1)

    # Seleccionar Bodega 04
    print("📦 Seleccionando bodega: 04 - Mangueras...")
    wait_and_click("bodega_04.png", confidence=0.8, fallback=(880, 420))
    time.sleep(1)

    # ---- NAVEGAR A CAMPOS DE FECHA ----
    print("📅 Navegando a campos de fecha...")
    for _ in range(3):
        pyautogui.press("tab")
        time.sleep(0.1)

    # Ingresar Fecha Inicio
    print(f"🟢 Fecha inicio: {fecha_inicio}")
    pyautogui.write(fecha_inicio)
    time.sleep(0.1)

    # Ingresar Fecha Fin
    pyautogui.press("tab")
    print(f"🔴 Fecha fin: {fecha_fin}")
    pyautogui.write(fecha_fin)
    time.sleep(0.1)

    print("✅ Informe configurado correctamente")

# ===== CLIPBOARD =====
def read_clipboard_to_df(raw_data):
    try:
        df = pd.read_csv(io.StringIO(raw_data), sep='\t')
    except:
        df = pd.read_csv(io.StringIO(raw_data), sep=';')

    df = df.dropna(how='all')
    df = df.drop_duplicates()

    return df

def flow_facturas_movimientos():
    print("\n🚀 Ejecutando flujo: Facturas y movimiento de inventario\n")

    # ===== INPUT DE FECHAS =====
    fecha_inicio, fecha_fin = get_date_input()

    # ===== LOGIN =====
    print("🔐 Iniciando sesión en DAPAS...")
    login_to_dapas()
    time.sleep(2.5)

    # ===== NAVEGACIÓN =====
    print("📊 Navegando al módulo de informes...")
    navigate_to_report()
    time.sleep(2.5)

    # ===== CONFIGURACIÓN =====
    configure_report(fecha_inicio, fecha_fin)

    # ===== GENERAR INFORME =====
    print("⚙️ Generando informe...")
    wait_and_click("generar.png", confidence=0.8, fallback=(100, 100))

    # ===== ESPERAR PROCESAMIENTO =====
    print("⏳ Esperando que el informe esté listo...")
    wait_until_visible("wait_exportar.png")

    # Cerrar posible modal
    print("🪟 Cerrando ventana de confirmación...")
    wait_and_click("ok.png", confidence=0.8, fallback=(745, 435))

    # ===== EXPORTAR =====
    print("📤 Exportando datos...")
    wait_and_click("exportar.png", confidence=0.8, fallback=(390, 100))
    time.sleep(0.2)

    wait_and_click("clipboard_option.png", confidence=0.8, fallback=(300, 120))

    time.sleep(2)
    pyautogui.press("enter")

    # ===== LECTURA DE DATOS =====
    print("📋 Leyendo datos desde el portapapeles...")
    raw_data = wait_for_clipboard_change()

    df = read_clipboard_to_df(raw_data)

    # ===== FILTRADO =====
    print("🧹 Procesando datos...")
    df = df[[
        "d_fecha_documento",
        "k_sc_codigo_fuente",
        "sc_nombre"
    ]]

    # ===== RESULTADO =====
    print("✅ Datos obtenidos correctamente\n")
    print(df.head())

    return df

# ===== EJECUCIÓN DE FLUJOS =====
def run_selected_flow(option):
    if option == "1":
        return flow_facturas_movimientos()
    elif option == "0":
        print("👋 Saliendo del sistema...")
        exit()
    else:
        print("❌ Opción inválida\n")


# ===== FLUJO PRINCIPAL =====
def run_flow():
    print("🚀 Iniciando automatización...\n")

    app_login()

    while True:
        opcion = show_menu()
        resultado = run_selected_flow(opcion)

        if resultado is not None:
            print("\n✅ Flujo ejecutado correctamente\n")
        else:
            print("\n⚠️ No se ejecutó ningún flujo\n")


# ===== PUNTO DE ENTRADA =====
if __name__ == "__main__":
    try:
        run_flow()
    except Exception as e:
        print("\n❌ ERROR EN LA EJECUCIÓN:")
        traceback.print_exc()
        input("\nPresione ENTER para salir...")