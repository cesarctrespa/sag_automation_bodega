import pyautogui
import time
import pyperclip
import pandas as pd
import io
from datetime import datetime
import os
import sys


# ===== CONFIG =====
def get_base_path():
    if hasattr(sys, "_MEIPASS"):
        return sys._MEIPASS
    return os.path.abspath(".")


IMAGES_PATH = os.path.join(get_base_path(), "images") + "\\"
CONFIDENCE = 0.8


# ===== HELPERS =====
def get_date_input():
    print("\n📅 Ingrese el rango de fechas (formato: DD-MM-YYYY)")
    print("Ejemplo: 01-01-2026")
    while True:
        fecha_inicio_input = input("Fecha inicio: ")
        fecha_fin_input = input("Fecha fin: ")
        try:
            inicio = datetime.strptime(fecha_inicio_input, "%d-%m-%Y")
            fin = datetime.strptime(fecha_fin_input, "%d-%m-%Y")
            if inicio > fin:
                print("❌ 'Fecha inicio' no puede ser mayor que 'Fecha fin'\n")
                continue
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
                IMAGES_PATH + image, confidence=confidence, grayscale=True
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


def wait_until_visible(image, timeout=600, confidence=CONFIDENCE, region=None):
    full_path = IMAGES_PATH + image
    start = time.time()
    while True:
        try:
            location = pyautogui.locateOnScreen(
                full_path, confidence=confidence, grayscale=True, region=region
            )
            if location:
                print(f"👁️ Elemento visible: {image}")
                return location
        except Exception as e:
            print(f"⚠️ Esperando elemento visible: {image}")
        if time.time() - start > timeout:
            raise TimeoutError(f"⏱️ Tiempo de espera agotado para: {image}")
        time.sleep(2)


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


def read_clipboard_to_df(raw_data):
    try:
        df = pd.read_csv(io.StringIO(raw_data), sep="\t")
    except:
        df = pd.read_csv(io.StringIO(raw_data), sep=";")

    df = df.dropna(how="all")
    df = df.drop_duplicates()
    return df


def scroll_until_image(image, arrow_pos, confidence=0.9, max_clicks=30):
    print(f"🔍 Buscando imagen con navegación: {image}")
    full_path = os.path.join(IMAGES_PATH, image)
    for i in range(max_clicks):
        try:
            location = pyautogui.locateOnScreen(
                full_path, confidence=confidence, grayscale=True
            )

            if location:
                print(f"✅ Encontrado: {image}")
                return location
        except pyautogui.ImageNotFoundException:
            pass
        pyautogui.click(arrow_pos[0], arrow_pos[1])
        print(
            f"🔽 Click en flecha ({arrow_pos[0]}, {arrow_pos[1]}) [{i+1}/{max_clicks}]"
        )
        time.sleep(0.2)
    raise Exception(f"❌ No se encontró la imagen: {image}")


def reset_checkboxes(marcar_img, desmarcar_img):
    print("🧹 Reiniciando selección (Marcar → Desmarcar)...")
    location_marcar = wait_until_visible(marcar_img, confidence=0.8)
    x1, y1 = pyautogui.center(location_marcar)
    pyautogui.click(x1, y1)
    time.sleep(1)
    location_desmarcar = wait_until_visible(desmarcar_img, confidence=0.8)
    x2, y2 = pyautogui.center(location_desmarcar)
    pyautogui.click(x2, y2)
    time.sleep(1)
    print("✅ Selección reiniciada")
    wait_and_click("aceptar.png", confidence=0.8)
    time.sleep(1)
