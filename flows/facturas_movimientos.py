import time
import pyautogui
import subprocess

from helpers import (
    get_date_input,
    wait_and_click,
    wait_until_visible,
    wait_for_clipboard_change,
    read_clipboard_to_df
)

# ===== APP AUTH =====
USUARIO = "mauricio"
PASSWORD = "1234"

# ===== NAVIGATION =====
def navigate_to_report():
    print("📊 Navegando al módulo de informes...")

    wait_and_click("menu_informes.png", confidence=0.8, fallback=(150, 33))
    time.sleep(0.5)

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

    print("⏳ Esperando que SAG cargue (10s)...")
    time.sleep(10)

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
    time.sleep(5)

    pyautogui.press("enter")
    pyautogui.press("enter")

    print("✅ Inicio de sesión completado")

# ===== CONFIG =====
def configure_report(fecha_inicio, fecha_fin):
    print("⚙️ Configurando el informe...")

    wait_and_click("otros_05b.png", confidence=0.8, fallback=(70, 200))
    time.sleep(1)

    wait_and_click("bodega_dropdown.png", confidence=0.8, fallback=(995, 340))
    time.sleep(1)

    wait_and_click("bodega_04.png", confidence=0.8, fallback=(880, 420))
    time.sleep(1)

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
    time.sleep(2.5)

    print("📊 Navegando al módulo de informes...")
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

    wait_and_click("clipboard_option.png", confidence=0.8, fallback=(300, 120))

    time.sleep(2)
    pyautogui.press("enter")

    print("📋 Leyendo datos desde el portapapeles...")
    raw_data = wait_for_clipboard_change()

    df = read_clipboard_to_df(raw_data)

    print("🧹 Procesando datos...")
    df = df[[
        "d_fecha_documento",
        "k_sc_codigo_fuente",
        "sc_nombre"
    ]]

    print("✅ Datos obtenidos correctamente\n")
    print(df.head())

    return df