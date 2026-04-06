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
CONFIDENCE = 0.6


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


def read_clipboard_to_df(raw_data):
    try:
        df = pd.read_csv(io.StringIO(raw_data), sep="\t")
    except:
        df = pd.read_csv(io.StringIO(raw_data), sep=";")

    df = df.dropna(how="all")
    df = df.drop_duplicates()
    return df
