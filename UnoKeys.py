import yaml
import serial
import threading
import pystray
from pystray import MenuItem as item
from PIL import Image
import time
import subprocess
import os
import sys
from pynput.keyboard import Controller, Key
from pynput.mouse import Button, Controller as MouseController
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
import pythoncom

# Inicializar COM en el hilo principal
pythoncom.CoInitialize()

# Función para leer la configuración desde el archivo YAML
def load_config():
    try:
        with open('config.yaml', 'r') as f:
            config = yaml.safe_load(f)
        return config
    except Exception as e:
        print(f"Error al leer el archivo de configuración: {e}")
        return None

# Función para convertir cadenas de texto a objetos Key de pynput
def str_to_key(key_str):
    try:
        if key_str.startswith("Key."):
            return getattr(Key, key_str.split('.')[1])
        return key_str
    except Exception as e:
        print(f"Error al convertir la cadena de texto a objeto Key: {e}")
        return None

# Funciones para controlar el micrófono y el sonido del sistema
def mute_system_sound():
    pythoncom.CoInitialize()  # Inicializar COM
    try:
        devices = AudioUtilities.GetSpeakers()
        volume = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = volume.QueryInterface(IAudioEndpointVolume)
        volume.SetMute(True, None)
    except Exception as e:
        print(f"Error al silenciar el sistema: {e}")
    finally:
        pythoncom.CoUninitialize()  # Finalizar COM

def unmute_system_sound():
    pythoncom.CoInitialize()  # Inicializar COM
    try:
        devices = AudioUtilities.GetSpeakers()
        volume = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = volume.QueryInterface(IAudioEndpointVolume)
        volume.SetMute(False, None)
    except Exception as e:
        print(f"Error al activar el sonido del sistema: {e}")
    finally:
        pythoncom.CoUninitialize()  # Finalizar COM

def mute_microphone():
    pythoncom.CoInitialize()  # Inicializar COM
    try:
        devices = AudioUtilities.GetMicrophone()
        volume = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = volume.QueryInterface(IAudioEndpointVolume)
        volume.SetMute(True, None)
    except Exception as e:
        print(f"Error al silenciar el micrófono: {e}")
    finally:
        pythoncom.CoUninitialize()  # Finalizar COM

def unmute_microphone():
    pythoncom.CoInitialize()  # Inicializar COM
    try:
        devices = AudioUtilities.GetMicrophone()
        volume = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = volume.QueryInterface(IAudioEndpointVolume)
        volume.SetMute(False, None)
    except Exception as e:
        print(f"Error al activar el micrófono: {e}")
    finally:
        pythoncom.CoUninitialize()  # Finalizar COM

# Cargar la configuración
config = load_config()
if config is None:
    print("Error: No se pudo cargar la configuración.")
    exit(1)

port = config['serial']['port']
baudrate = config['serial']['baudrate']
key_combinations = {int(k): v for k, v in config['key_combinations'].items()}
hold_time = config['repeat']['hold_time']
repeat_rate = config['repeat']['repeat_rate']

# Configura el puerto serial
try:
    ser = serial.Serial(port, baudrate)
except Exception as e:
    print(f"Error al configurar el puerto serial: {e}")
    exit(1)

# Inicializa el controlador de teclado y mouse
keyboard = Controller()
mouse = MouseController()

# Lista para mantener las teclas presionadas actualmente
current_keys_pressed = set()

# Variables para el mantenimiento de clics
click_hold_state = {}

# Función para manejar la presión de teclas y acciones del mouse con delay
def press_keys_with_delay(keys, button_number=None):
    global current_keys_pressed
    for key_info in keys:
        if 'action' in key_info:
            action = key_info['action']
            if action == "delay":
                time.sleep(key_info['duration'] / 1000.0)  # Convertir a segundos
            elif action == "click_left":
                mouse.click(Button.left)
            elif action == "click_right":
                mouse.click(Button.right)
            elif action == "click_middle":
                mouse.click(Button.middle)
            elif action == "double_click_left":
                mouse.click(Button.left)
                time.sleep(0.1)  # Pequeña pausa entre clics
                mouse.click(Button.left)
            elif action == "move":
                x = key_info.get('x', 0)
                y = key_info.get('y', 0)
                mouse.position = (x, y)
            elif action == "scroll":
                direction = key_info.get('direction', 'down')
                amount = key_info.get('amount', 1)
                if direction == "up":
                    mouse.scroll(0, amount)
                elif direction == "down":
                    mouse.scroll(0, -amount)
                elif direction == "left":
                    mouse.scroll(-amount, 0)
                elif direction == "right":
                    mouse.scroll(amount, 0)
            elif action == "mute_system":
                mute_system_sound()
            elif action == "unmute_system":
                unmute_system_sound()
            elif action == "mute_microphone":
                mute_microphone()
            elif action == "unmute_microphone":
                unmute_microphone()
            elif action == "holdclick_left" and button_number is not None:
                if button_number not in click_hold_state:
                    mouse.press(Button.left)
                    click_hold_state[button_number] = 'left'
            elif action == "holdclick_right" and button_number is not None:
                if button_number not in click_hold_state:
                    mouse.press(Button.right)
                    click_hold_state[button_number] = 'right'
            else:
                try:
                    key = str_to_key(action)
                    if isinstance(key, str):
                        keyboard.press(key)
                        current_keys_pressed.add(key)
                    else:
                        keyboard.press(key)
                        current_keys_pressed.add(key)
                except Exception as e:
                    print(f"Error al presionar la tecla {action}: {e}")

# Función para manejar la liberación de teclas (vacío para que no haga nada)
def release_keys(keys):
    pass

# Variable para mantener el estado de los botones
button_states = {}

# Función para manejar la repetición de teclas
def repeat_keys(button_number):
    while button_number in button_states:
        pythoncom.CoInitialize()  # Inicializar COM en el hilo
        start_time = time.time()
        press_keys_with_delay(key_combinations[button_number], button_number)
        elapsed_time = (time.time() - start_time) * 1000  # Convertir a milisegundos

        if any('delay' in action for action in key_combinations[button_number]):
            time.sleep(max(0, repeat_rate - elapsed_time) / 1000.0)  # Asegurar que no haya delay adicional
        else:
            time.sleep(repeat_rate / 1000.0)  # Esperar el tiempo de repetición
        pythoncom.CoUninitialize()  # Finalizar COM en el hilo

# Función principal para manejar los eventos desde Arduino
def main():
    print("Esperando datos del Arduino...")

    # Variable para controlar la ejecución del hilo
    running = True

    def on_quit_callback(icon, item):
        nonlocal running
        running = False
        icon.stop()

    def on_config_callback(icon, item):
        # Obtener el directorio del ejecutable
        executable_dir = os.path.dirname(sys.executable)
        subprocess.Popen(['explorer', executable_dir])  # Para Windows
        # subprocess.Popen(['open', executable_dir])  # Para macOS
        # subprocess.Popen(['xdg-open', executable_dir])  # Para Linux

    # Crear icono de la bandeja
    icon = pystray.Icon("nombre-app")
    icon.icon = Image.open("icon.ico")
    
    # Crear menú para el icono
    menu = (item('Config', on_config_callback), item('Exit', on_quit_callback))
    icon.menu = menu

    # Función para el hilo
    def thread_func():
        while running:
            try:
                if ser.in_waiting > 0:
                    line = ser.readline().decode('utf-8').strip()
                    print(f"Recibido desde Arduino: {line}")  # Mensaje de depuración
                    parts = line.split()
                    button_number = int(parts[1])
                    action = parts[2]
                    
                    if action == "Pressed" and button_number not in button_states:
                        button_states[button_number] = True
                        if button_number in key_combinations:
                            print(f"Ejecutando acciones para botón {button_number}")
                            press_keys_with_delay(key_combinations[button_number], button_number)
                            # Iniciar la cuenta del tiempo de repetición
                            threading.Timer(hold_time / 1000.0, repeat_keys, [button_number]).start()
                    elif action == "Released" and button_number in button_states:
                        if button_number in click_hold_state:
                            if click_hold_state[button_number] == 'left':
                                mouse.release(Button.left)
                            elif click_hold_state[button_number] == 'right':
                                mouse.release(Button.right)
                            del click_hold_state[button_number]
                        del button_states[button_number]
            except Exception as e:
                print(f"Error durante la lectura del puerto serial: {e}")

    # Iniciar el hilo para manejar los eventos desde Arduino
    thread = threading.Thread(target=thread_func)
    thread.start()

    # Agregar el icono a la bandeja del sistema
    icon.run()

if __name__ == "__main__":
    main()
