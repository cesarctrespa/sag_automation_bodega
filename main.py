import pyautogui
import traceback
import getpass

from flows.facturas_movimientos import flow_facturas_movimientos

# ===== APP AUTH =====
APP_USER = "admin"
APP_PASSWORD = "1914"

pyautogui.FAILSAFE = True

def app_login():
    print("=== 🔐 INICIO DE SESIÓN ===")
    for intento in range(3):
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

def run_selected_flow(option):
    if option == "1":
        return flow_facturas_movimientos()
    elif option == "0":
        print("👋 Saliendo del sistema...")
        exit()
    else:
        print("❌ Opción inválida\n")

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

if __name__ == "__main__":
    try:
        run_flow()
    except Exception:
        print("\n❌ ERROR EN LA EJECUCIÓN:")
        traceback.print_exc()
        input("\nPresione ENTER para salir...")