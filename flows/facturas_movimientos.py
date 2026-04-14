import time
import pyautogui
import subprocess

from helpers import (
    get_date_input,
    reset_checkboxes,
    scroll_until_image,
    wait_and_click,
    wait_until_visible,
    wait_for_clipboard_change,
    read_clipboard_to_df,
    IMAGES_PATH,
)

# ===== APP AUTH =====
USUARIO = "mauricio"
PASSWORD = "1234"


# ===== NAVIGATION =====
def navigate_to_report():
    print("📊 Navegando al módulo de informes...")
    wait_and_click("menu_informes.png", confidence=0.8, fallback=(150, 33))
    time.sleep(0.25)
    print("📂 Accediendo a 'Varios'...")
    pyautogui.press("down")
    time.sleep(0.1)
    pyautogui.press("right")
    time.sleep(0.1)
    print("📄 Seleccionando '05 - Facturas y movimientos'...")
    for _ in range(4):
        pyautogui.press("down")
        time.sleep(0.1)
    pyautogui.press("enter")
    print("✅ Módulo de informe abierto")


# ===== LOGIN =====
def login_to_dapas():
    print("🚀 Abriendo DAPAS...")
    subprocess.Popen(["mstsc", r"C:\Users\ASUS\OneDrive\Escritorio\DAPAS.rdp"])
    print("⏳ Esperando que la interfaz de SAG cargue...")
    wait_until_visible("sag_main_screen.png", timeout=120, confidence=0.7)
    print("🗂️ Seleccionando base de datos...")
    for _ in range(2):
        pyautogui.press("down")
        time.sleep(0.1)
    print("👤 Navegando al campo de usuario...")
    for _ in range(2):
        pyautogui.press("tab")
        time.sleep(0.1)
    pyautogui.write(USUARIO)
    print("🔑 Ingresando contraseña...")
    pyautogui.press("tab")
    pyautogui.write(PASSWORD)
    print("📤 Enviando credenciales...")
    pyautogui.press("enter")
    print("⚠️ Gestionando mensaje de sesión activa...")
    wait_until_visible("sag_welcome.png", timeout=120, confidence=0.7)
    pyautogui.press("enter")
    pyautogui.press("enter")
    wait_until_visible("sag_loaded.png", timeout=120, confidence=0.7)
    print("✅ Inicio de sesión completado")


def select_bodega_04():
    print("🏭 Seleccionando bodega: 04 - Mangueras...")
    # Abrir modal
    wait_and_click("seleccionar_bodegas.png", confidence=0.8)
    time.sleep(1)
    reset_checkboxes("marcar_bodega.png", "desmarcar_bodega.png")
    # Seleccionar checkbox 04 (click preciso)
    wait_and_click("seleccionar_bodegas.png", confidence=0.8)
    time.sleep(1)
    print("🔍 Buscando checkbox de bodega 04...")
    location = wait_until_visible("bodega_04_checkbox.png", confidence=0.8)
    x, y = pyautogui.center(location)
    # Ajuste opcional hacia la izquierda (checkbox suele estar ahí)
    pyautogui.click(x - 115, y)
    print(f"🖱️ Coordenadas: ({x-115}, {y})")
    print("🖱️ Checkbox de bodega 04 seleccionado")
    time.sleep(0.5)
    # Confirmar
    print("✅ Confirmando selección de bodega...")
    wait_and_click("aceptar.png", confidence=0.8)
    time.sleep(0.2)


def select_fuentes():
    print("📌 Seleccionando fuentes (CM, FM, NE)...")
    wait_and_click("seleccionar_fuentes.png", confidence=0.8)
    time.sleep(1)
    reset_checkboxes("marcar_fuentes.png", "desmarcar_fuentes.png")
    wait_and_click("seleccionar_fuentes.png", confidence=0.8)
    time.sleep(1)
    fuentes = ["CM.png", "FM.png", "NE.png"]
    arrow_down_pos = (930, 522)
    for fuente in fuentes:
        print(f"🔍 Buscando fuente: {fuente}")
        location = scroll_until_image(fuente, arrow_down_pos)
        x, y = pyautogui.center(location)
        pyautogui.click(x - 115, y)
        print(f"🖱️ Coordenadas: ({x-115}, {y})")
        print(f"🖱️ Fuente seleccionada: {fuente}")
        time.sleep(0.5)
    print("✅ Confirmando selección...")
    wait_and_click("aceptar.png", confidence=0.8)
    time.sleep(0.2)


# ===== CONFIG =====
def configure_report(fecha_inicio, fecha_fin):
    print("⚙️ Configurando el informe...")
    wait_and_click("otros_05b.png", confidence=0.8, fallback=(70, 200))
    time.sleep(1)
    select_bodega_04()
    select_fuentes()
    # ---- NAVEGAR A CAMPOS DE FECHA ----
    print("📅 Navegando a campos de fecha...")
    for _ in range(3):
        pyautogui.press("tab")
        time.sleep(0.1)
    print(f"🟢 Fecha inicio: {fecha_inicio}")
    pyautogui.write(fecha_inicio)
    pyautogui.press("tab")
    print(f"🔴 Fecha fin: {fecha_fin}")
    pyautogui.write(fecha_fin)
    print("✅ Informe configurado correctamente")


# ===== FLOW =====
def flow_facturas_movimientos():
    print("\n🚀 Ejecutando flujo: Facturas y movimiento de inventario\n")
    fecha_inicio, fecha_fin = get_date_input()
    print("🔐 Iniciando sesión en DAPAS...")
    login_to_dapas()
    navigate_to_report()
    time.sleep(2.5)
    configure_report(fecha_inicio, fecha_fin)
    print("⚙️ Generando informe...")
    wait_and_click("generar.png", confidence=0.8, fallback=(100, 100))
    print("⏳ Esperando que el informe esté listo...")
    wait_until_visible("wait_exportar.png")
    print("🪟 Cerrando ventana de confirmación...")
    wait_and_click("ok.png", confidence=0.8, fallback=(745, 435))
    print("📤 Exportando datos...")
    wait_and_click("exportar.png", confidence=0.8, fallback=(390, 100))
    time.sleep(0.2)
    wait_and_click("clipboard_option.png", confidence=0.8, fallback=(385, 167))
    time.sleep(2)
    pyautogui.press("enter")
    print("📋 Leyendo datos desde el portapapeles...")
    raw_data = wait_for_clipboard_change()
    df = read_clipboard_to_df(raw_data)
    print("🧹 Procesando datos...")
    df = df[["d_fecha_documento", "k_sc_codigo_fuente", "sc_nombre"]]
    print("✅ Datos obtenidos correctamente\n")
    print(df.head())
    return df
