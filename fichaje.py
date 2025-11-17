"""
SISTEMA DE FICHAJE AUTOMATIZADO - VERSIÓN COMPLETA CON NOTIFICACIONES
Incluye: Script de fichaje + Interfaz gráfica + Programador de horarios + Notificaciones (Telegram y Email)
TODO EN UN SOLO ARCHIVO
"""

import os
import sys
import time
import base64
import requests
import pandas as pd
import logging
import threading
import schedule
import json
import configparser
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from pathlib import Path

# Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

# Tkinter para GUI
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog

# ==================== CONFIGURACIÓN GLOBAL ====================
CONFIG = {
    'url': "http://172.22.0.132/",
    'csv_file': "datos.csv",
    'log_file': f"fichajes_{datetime.now().strftime('%Y%m%d')}.log",
    'results_file': f"resultados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
    'screenshots_dir': "screenshots",
    'config_file': "horarios_config.json",
    'notifications_file': "notificaciones.ini",
    'headless': False,
    'api_key_2captcha': os.getenv("API_KEY_2CAPTCHA", "41c8f96621747395fd9731ebd83a746c"),
    'timeout_short': 5,
    'timeout_medium': 10,
    'timeout_long': 30
}

# Crear directorios necesarios
os.makedirs(CONFIG['screenshots_dir'], exist_ok=True)

# ==================== CONFIGURACIÓN DE LOGS ====================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(CONFIG['log_file'], encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


# ==================== GESTOR DE NOTIFICACIONES ====================
class NotificationManager:
    """Gestor de notificaciones por Telegram y Email"""

    def __init__(self, config_file):
        self.config_file = config_file
        self.telegram_enabled = False
        self.email_enabled = False
        self.telegram_token = None
        self.telegram_chat_id = None
        self.email_config = {}

        self.load_config()

    def load_config(self):
        """Carga la configuración desde notificaciones.ini"""
        try:
            if not os.path.exists(self.config_file):
                logger.warning(f"⚠️ No existe {self.config_file} - Notificaciones desactivadas")
                return

            config = configparser.ConfigParser()
            config.read(self.config_file, encoding='utf-8')

            # Cargar configuración de Telegram
            if 'telegram' in config:
                token = config['telegram'].get('token', '').strip()
                chat_id = config['telegram'].get('chat_id', '').strip()

                if token and chat_id and token != 'AQUI_TU_TOKEN' and chat_id != 'AQUI_TU_CHAT_ID':
                    self.telegram_token = token
                    self.telegram_chat_id = chat_id
                    self.telegram_enabled = True
                    logger.info("✅ Notificaciones de Telegram configuradas")

            # Cargar configuración de Email
            if 'email' in config:
                smtp_server = config['email'].get('smtp_server', '').strip()
                smtp_port = config['email'].get('smtp_port', '').strip()
                email_from = config['email'].get('email_from', '').strip()
                email_password = config['email'].get('email_password', '').strip()
                email_to = config['email'].get('email_to', '').strip()

                if (smtp_server and smtp_port and email_from and email_password and email_to and
                    smtp_server != 'AQUI_TU_SERVIDOR_SMTP' and email_from != 'AQUI_TU_EMAIL'):
                    self.email_config = {
                        'smtp_server': smtp_server,
                        'smtp_port': int(smtp_port),
                        'email_from': email_from,
                        'email_password': email_password,
                        'email_to': email_to
                    }
                    self.email_enabled = True
                    logger.info("✅ Notificaciones de Email configuradas")

        except Exception as e:
            logger.error(f"❌ Error cargando configuración de notificaciones: {e}")

    def send_telegram(self, mensaje):
        """Envía notificación por Telegram"""
        if not self.telegram_enabled:
            return False

        try:
            url = f"https://api.telegram.org/bot{self.telegram_token}/sendMessage"
            data = {
                'chat_id': self.telegram_chat_id,
                'text': mensaje,
                'parse_mode': 'HTML'
            }
            response = requests.post(url, data=data, timeout=10)

            if response.status_code == 200:
                logger.info("📱 Notificación de Telegram enviada")
                return True
            else:
                logger.error(f"❌ Error enviando Telegram: {response.status_code}")
                return False

        except Exception as e:
            logger.error(f"❌ Error enviando notificación de Telegram: {e}")
            return False

    def send_email(self, asunto, mensaje):
        """Envía notificación por Email"""
        if not self.email_enabled:
            return False

        try:
            msg = MIMEMultipart()
            msg['From'] = self.email_config['email_from']
            msg['To'] = self.email_config['email_to']
            msg['Subject'] = asunto

            msg.attach(MIMEText(mensaje, 'html'))

            with smtplib.SMTP(self.email_config['smtp_server'], self.email_config['smtp_port']) as server:
                server.starttls()
                server.login(self.email_config['email_from'], self.email_config['email_password'])
                server.send_message(msg)

            logger.info("📧 Notificación de Email enviada")
            return True

        except Exception as e:
            logger.error(f"❌ Error enviando notificación de Email: {e}")
            return False

    def notify(self, titulo, mensaje, tipo="info"):
        """Envía notificación por todos los canales configurados"""
        if not self.telegram_enabled and not self.email_enabled:
            return

        # Emojis según el tipo
        emojis = {
            'success': '✅',
            'error': '❌',
            'warning': '⚠️',
            'info': 'ℹ️'
        }
        emoji = emojis.get(tipo, 'ℹ️')

        # Formato para Telegram (HTML)
        telegram_msg = f"{emoji} <b>{titulo}</b>\n\n{mensaje}"

        # Formato para Email (HTML)
        email_html = f"""
        <html>
            <body style="font-family: Arial, sans-serif;">
                <h2 style="color: #2c3e50;">{emoji} {titulo}</h2>
                <div style="background-color: #ecf0f1; padding: 15px; border-radius: 5px;">
                    <pre style="white-space: pre-wrap;">{mensaje}</pre>
                </div>
                <hr>
                <p style="color: #7f8c8d; font-size: 12px;">
                    Sistema de Fichaje Automatizado v2.1<br>
                    {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
                </p>
            </body>
        </html>
        """

        # Enviar por ambos canales
        if self.telegram_enabled:
            self.send_telegram(telegram_msg)

        if self.email_enabled:
            self.send_email(f"[Fichaje] {titulo}", email_html)


# ==================== CLASE PARA EL MOTOR DE FICHAJE ====================
class FichajeEngine:
    """Motor de fichaje con todas las funciones necesarias"""

    def __init__(self, config):
        self.config = config
        self.notifier = NotificationManager(config['notifications_file'])

    def start_driver(self, headless=False):
        """Inicia el driver de Chrome con configuración optimizada"""
        try:
            options = webdriver.ChromeOptions()
            if headless:
                options.add_argument("--headless=new")

            # Opciones para mayor estabilidad
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-gpu")
            options.add_argument("--disable-software-rasterizer")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)

            prefs = {
                "profile.default_content_setting_values.notifications": 2,
                "profile.default_content_settings.popups": 0
            }
            options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(
                service=ChromeService(ChromeDriverManager().install()),
                options=options
            )
            driver.set_window_size(1200, 900)
            driver.set_page_load_timeout(60)

            logger.info("✅ Driver de Chrome iniciado correctamente")
            return driver
        except Exception as e:
            logger.error(f"❌ Error iniciando driver: {e}")
            raise

    def solve_captcha_2captcha(self, image_path, api_key, timeout=120):
        """Resuelve captcha usando API de 2Captcha"""
        try:
            with open(image_path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()

            logger.info("📤 Enviando captcha a 2Captcha...")
            r = requests.post("http://2captcha.com/in.php", data={
                "method": "base64",
                "key": api_key,
                "body": b64,
                "json": 1
            }, timeout=10).json()

            if r.get("status") != 1:
                logger.error(f"❌ Error enviando captcha: {r}")
                return None

            captcha_id = r["request"]
            logger.info(f"⏳ Esperando respuesta de captcha...")

            for _ in range(timeout // 5):
                time.sleep(5)
                res = requests.get("http://2captcha.com/res.php", params={
                    "key": api_key,
                    "action": "get",
                    "id": captcha_id,
                    "json": 1
                }, timeout=10).json()

                if res.get("status") == 1:
                    logger.info(f"✅ Captcha resuelto: {res['request']}")
                    return res["request"]
                if res.get("request") != "CAPCHA_NOT_READY":
                    logger.error(f"❌ Error resolviendo captcha: {res}")
                    return None

            logger.warning("⏱ Timeout esperando respuesta de 2Captcha")
            return None
        except Exception as e:
            logger.error(f"❌ Error en 2Captcha: {e}")
            return None

    def find_captcha_image(self, driver):
        """Busca la imagen del captcha"""
        try:
            imgs = driver.find_elements(By.TAG_NAME, "img")
            for img in imgs:
                src = (img.get_attribute("src") or "").lower()
                if "captcha" in src or "codigo" in src:
                    logger.info("✅ Imagen de captcha encontrada")
                    return img
            logger.warning("⚠️ No se encontró imagen de captcha")
            return None
        except Exception as e:
            logger.error(f"❌ Error buscando captcha: {e}")
            return None

    def take_screenshot(self, driver, filename):
        """Captura screenshot con timestamp"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filepath = os.path.join(self.config['screenshots_dir'], f"{timestamp}_{filename}")
            driver.save_screenshot(filepath)
            logger.info(f"📸 Screenshot guardado: {filepath}")
            return filepath
        except Exception as e:
            logger.error(f"❌ Error guardando screenshot: {e}")
            return ""

    def guardar_resultado(self, usuario, estado, mensaje, screenshot_path=""):
        """Guarda el resultado del fichaje en CSV"""
        try:
            resultado = {
                'fecha_hora': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'usuario': usuario,
                'estado': estado,
                'mensaje': mensaje,
                'screenshot': screenshot_path
            }

            df = pd.DataFrame([resultado])

            if os.path.exists(self.config['results_file']):
                df.to_csv(self.config['results_file'], mode='a', header=False, index=False, encoding='utf-8')
            else:
                df.to_csv(self.config['results_file'], mode='w', header=True, index=False, encoding='utf-8')

            logger.info(f"📝 Resultado guardado")
        except Exception as e:
            logger.error(f"❌ Error guardando resultado: {e}")

    def safe_click(self, driver, element, description="elemento"):
        """Hace clic de forma segura con múltiples intentos"""
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
            time.sleep(0.3)
            element.click()
            logger.info(f"✅ Clic en {description}")
            return True
        except Exception as e1:
            try:
                driver.execute_script("arguments[0].click();", element)
                logger.info(f"✅ Clic con JS en {description}")
                return True
            except Exception as e2:
                logger.error(f"❌ No se pudo hacer clic en {description}")
                return False

    def realizar_fichaje(self, usuario, password, driver, callback=None):
        """Realiza el proceso completo de fichaje"""
        logger.info(f"\n{'=' * 70}")
        logger.info(f"🚀 FICHAJE PARA: {usuario}")
        logger.info(f"{'=' * 70}")

        if callback:
            callback(f"Iniciando fichaje para {usuario}...")

        screenshot_path = ""

        try:
            # 1. Cargar página
            logger.info("📍 Cargando página de login...")
            if callback:
                callback("Cargando página de login...")
            driver.get(self.config['url'])
            time.sleep(3)

            # 2. Entrar a frames
            logger.info("🔀 Entrando a frames...")
            if callback:
                callback("Accediendo al sistema...")
            driver.switch_to.frame("cuerpo_WCRONOS")
            time.sleep(1)
            driver.switch_to.frame("principal_wcronos")
            time.sleep(1)

            # 3. Localizar campos
            logger.info("🔍 Localizando campos del formulario...")
            if callback:
                callback("Localizando formulario de login...")

            tarjeta_field = WebDriverWait(driver, self.config['timeout_medium']).until(
                EC.presence_of_element_located((By.ID, "USUARIO"))
            )
            contrasena_field = driver.find_element(By.ID, "CONTRASENA")
            captcha_field = driver.find_element(By.NAME, "codigo_captcha")

            # 4. Captcha
            captcha_value = None
            captcha_img = self.find_captcha_image(driver)

            if captcha_img:
                if callback:
                    callback("Resolviendo captcha...")
                captcha_path = f"captcha_{usuario}_{datetime.now().strftime('%H%M%S')}.png"
                try:
                    captcha_img.screenshot(captcha_path)
                    captcha_value = self.solve_captcha_2captcha(captcha_path, self.config['api_key_2captcha'])
                    os.remove(captcha_path)
                except:
                    pass

            if not captcha_value:
                captcha_value = "0000"

            # 5. Rellenar formulario
            logger.info("📝 Rellenando formulario...")
            if callback:
                callback("Ingresando credenciales...")

            tarjeta_field.clear()
            tarjeta_field.send_keys(usuario)
            time.sleep(0.5)

            contrasena_field.clear()
            contrasena_field.send_keys(password)
            time.sleep(0.5)

            captcha_field.clear()
            captcha_field.send_keys(captcha_value)
            time.sleep(0.5)

            screenshot_path = self.take_screenshot(driver, f"antes_login_{usuario}.png")

            # 6. Hacer clic en ENTRAR
            logger.info("🔍 Buscando botón ENTRAR...")
            if callback:
                callback("Haciendo login...")

            botones = driver.find_elements(By.XPATH, "//button | //input[@type='submit']")
            entrar_clicked = False
            for btn in botones:
                texto = (btn.text or "").lower()
                value = (btn.get_attribute("value") or "").lower()
                if "entrar" in texto or "entrar" in value:
                    if self.safe_click(driver, btn, "botón ENTRAR"):
                        entrar_clicked = True
                        break

            if not entrar_clicked:
                raise Exception("No se pudo hacer clic en ENTRAR")

            # 7. Esperar redirección
            time.sleep(5)
            logger.info(f"📍 Login completado")

            # 8. Ir a Punto de Fichaje
            if callback:
                callback("Navegando a punto de fichaje...")
            time.sleep(2)

            boton_fichaje = driver.find_element(
                By.XPATH,
                "//button[contains(@onclick, 'form_pfichaje.submit')]"
            )
            onclick = boton_fichaje.get_attribute("onclick")
            driver.execute_script(onclick)
            logger.info("✅ Navegando a Punto De Fichaje")
            time.sleep(4)

            # 9. Volver a entrar en frames
            driver.switch_to.default_content()
            time.sleep(1)
            driver.switch_to.frame("cuerpo_WCRONOS")
            time.sleep(1)
            driver.switch_to.frame("principal_wcronos")
            time.sleep(1)

            # 10. REALIZAR FICHAJE
            logger.info("🔍 Buscando botón 'Realizar Fichaje'...")
            if callback:
                callback("Realizando fichaje...")

            fichaje_realizado = False

            # Intento 1: Por ID
            try:
                boton = driver.find_element(By.ID, "btnEnviarForm")
                if self.safe_click(driver, boton, "Realizar Fichaje"):
                    fichaje_realizado = True
            except:
                pass

            # Intento 2: Por texto
            if not fichaje_realizado:
                try:
                    botones = driver.find_elements(By.XPATH,
                                                   "//button[contains(., 'Realizar Fichaje')] | //button[contains(., 'Fichar')]")
                    if botones and self.safe_click(driver, botones[0], "Realizar Fichaje"):
                        fichaje_realizado = True
                except:
                    pass

            # Intento 3: Submit
            if not fichaje_realizado:
                driver.execute_script("""
                    let btn = document.getElementById('btnEnviarForm');
                    if (btn && btn.form) btn.form.submit();
                """)
                fichaje_realizado = True

            if not fichaje_realizado:
                raise Exception("No se pudo realizar el fichaje")

            # 11. Verificar resultado
            time.sleep(3)
            screenshot_path = self.take_screenshot(driver, f"resultado_{usuario}.png")

            # GUARDAR HTML COMPLETO PARA DEBUG
            try:
                page_html = driver.page_source
                html_file = os.path.join(self.config['screenshots_dir'],
                                         f"html_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{usuario}.html")
                with open(html_file, 'w', encoding='utf-8') as f:
                    f.write(page_html)
                logger.info(f"📄 HTML guardado en: {html_file}")

                # Imprimir HTML en consola también
                print("\n" + "=" * 80)
                print("🔍 HTML COMPLETO DE LA PÁGINA RESULTADO:")
                print("=" * 80)
                print(page_html)
                print("=" * 80)
                print("FIN DEL HTML")
                print("=" * 80 + "\n")
            except Exception as e:
                logger.error(f"Error guardando HTML: {e}")

            # Obtener el texto de la página
            page_text = driver.page_source.lower()

            # INDICADORES DE ÉXITO - Por prioridad (más específicos primero)
            success_indicators_alta_prioridad = [
                "el fichaje se a realizado correctamente",  # Mensaje EXACTO
                "el fichaje se ha realizado correctamente",  # Variante con H
                "fichaje se a realizado correctamente",
                "fichaje se ha realizado correctamente",
                "se a realizado correctamente",
                "se ha realizado correctamente",
            ]

            success_indicators_media_prioridad = [
                "fichaje realizado",
                "fichaje añadido",
                "fichaje registrado",
                "añadido con éxito",
                "registrado con éxito",
                "realizado con éxito",
                "fichaje correcto",
                "operación exitosa",
                "guardado correctamente",
            ]

            success_indicators_baja_prioridad = [
                "éxito",
                "exitoso",
                "correctamente",
                "confirmado",
                "completado"
            ]

            # INDICADORES DE ERROR - Solo mensajes claros de error
            error_indicators_especificos = [
                "error al realizar",
                "error en el fichaje",
                "fichaje incorrecto",
                "fichaje fallido",
                "no se pudo realizar",
                "operación fallida",
                "fichaje rechazado",
                "no se ha podido",
                "ha ocurrido un error"
            ]

            # PASO 1: Buscar indicadores de ALTA PRIORIDAD (mensajes específicos de éxito)
            for indicator in success_indicators_alta_prioridad:
                if indicator in page_text:
                    logger.info(f"✅✅✅ FICHAJE EXITOSO para {usuario}")
                    logger.info(f"   Mensaje detectado (ALTA PRIORIDAD): '{indicator}'")
                    if callback:
                        callback(f"✅ Fichaje exitoso para {usuario}")
                    self.guardar_resultado(usuario, "ÉXITO", f"Fichaje completado: {indicator}", screenshot_path)

                    # Enviar notificación de éxito
                    self.notifier.notify(
                        "Fichaje Exitoso",
                        f"Usuario: {usuario}\nEstado: ÉXITO\nMensaje: {indicator}\nFecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
                        tipo="success"
                    )
                    return True

            # PASO 2: Buscar errores ESPECÍFICOS (solo si no encontramos éxito de alta prioridad)
            error_especifico_found = False
            error_message = ""
            for indicator in error_indicators_especificos:
                if indicator in page_text:
                    error_especifico_found = True
                    error_message = indicator
                    break

            if error_especifico_found:
                logger.error(f"❌ FICHAJE FALLIDO para {usuario}")
                logger.error(f"   Error ESPECÍFICO detectado: '{error_message}'")
                if callback:
                    callback(f"❌ Error en fichaje para {usuario}")
                self.guardar_resultado(usuario, "ERROR", f"Error específico: {error_message}", screenshot_path)

                # Enviar notificación de error
                self.notifier.notify(
                    "Fichaje Fallido",
                    f"Usuario: {usuario}\nEstado: ERROR\nMensaje: {error_message}\nFecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
                    tipo="error"
                )
                return False

            # PASO 3: Buscar indicadores de MEDIA PRIORIDAD
            for indicator in success_indicators_media_prioridad:
                if indicator in page_text:
                    logger.info(f"✅✅✅ FICHAJE EXITOSO para {usuario}")
                    logger.info(f"   Mensaje detectado (MEDIA PRIORIDAD): '{indicator}'")
                    if callback:
                        callback(f"✅ Fichaje exitoso para {usuario}")
                    self.guardar_resultado(usuario, "ÉXITO", f"Fichaje completado: {indicator}", screenshot_path)

                    # Enviar notificación de éxito
                    self.notifier.notify(
                        "Fichaje Exitoso",
                        f"Usuario: {usuario}\nEstado: ÉXITO\nMensaje: {indicator}\nFecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
                        tipo="success"
                    )
                    return True

            # PASO 4: Buscar indicadores de BAJA PRIORIDAD
            for indicator in success_indicators_baja_prioridad:
                if indicator in page_text:
                    logger.info(f"✅✅✅ FICHAJE EXITOSO para {usuario}")
                    logger.info(f"   Mensaje detectado (BAJA PRIORIDAD): '{indicator}'")
                    if callback:
                        callback(f"✅ Fichaje exitoso para {usuario}")
                    self.guardar_resultado(usuario, "ÉXITO", f"Fichaje completado: {indicator}", screenshot_path)

                    # Enviar notificación de éxito
                    self.notifier.notify(
                        "Fichaje Exitoso",
                        f"Usuario: {usuario}\nEstado: ÉXITO\nMensaje: {indicator}\nFecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
                        tipo="success"
                    )
                    return True

            # PASO 5: Si no encontramos nada claro, intentar con el título
            try:
                page_title = driver.title.lower()
                if any(word in page_title for word in ["éxito", "correcto", "confirmado", "realizado"]):
                    logger.info(f"✅✅✅ FICHAJE EXITOSO para {usuario} (detectado en título)")
                    if callback:
                        callback(f"✅ Fichaje exitoso para {usuario}")
                    self.guardar_resultado(usuario, "ÉXITO", "Fichaje completado (título)", screenshot_path)

                    # Enviar notificación de éxito
                    self.notifier.notify(
                        "Fichaje Exitoso",
                        f"Usuario: {usuario}\nEstado: ÉXITO\nMensaje: Detectado en título de página\nFecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
                        tipo="success"
                    )
                    return True
            except:
                pass

            # PASO 6: Si llegamos aquí, no pudimos determinar el estado
            logger.warning(f"⚠️ ESTADO DESCONOCIDO para {usuario}")
            logger.warning(f"   No se encontraron indicadores claros de éxito o error")
            logger.warning(f"   Revisa el screenshot: {screenshot_path}")
            if callback:
                callback(f"⚠️ Estado desconocido para {usuario}")
            self.guardar_resultado(usuario, "DESCONOCIDO", "Estado no determinado - revisar screenshot",
                                   screenshot_path)

            # Enviar notificación de estado desconocido
            self.notifier.notify(
                "Estado Desconocido",
                f"Usuario: {usuario}\nEstado: DESCONOCIDO\nMensaje: No se pudo determinar el resultado\nRevisa el screenshot: {screenshot_path}\nFecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
                tipo="warning"
            )
            return None

        except WebDriverException as e:
            logger.error(f"❌ ERROR DE CHROMEDRIVER para {usuario}")
            if callback:
                callback(f"❌ Error de Chrome para {usuario}")
            screenshot_path = self.take_screenshot(driver, f"error_{usuario}.png")
            self.guardar_resultado(usuario, "ERROR", "Chrome crash", screenshot_path)

            # Enviar notificación de error
            self.notifier.notify(
                "Error de Chrome",
                f"Usuario: {usuario}\nEstado: ERROR\nMensaje: Chrome crash\nFecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
                tipo="error"
            )
            return False
        except Exception as e:
            logger.error(f"❌ ERROR para {usuario}: {e}")
            if callback:
                callback(f"❌ Error: {str(e)[:50]}")
            screenshot_path = self.take_screenshot(driver, f"error_{usuario}.png")
            self.guardar_resultado(usuario, "ERROR", str(e)[:200], screenshot_path)

            # Enviar notificación de error
            self.notifier.notify(
                "Error en Fichaje",
                f"Usuario: {usuario}\nEstado: ERROR\nMensaje: {str(e)[:200]}\nFecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
                tipo="error"
            )
            return False

    def procesar_usuarios(self, callback=None):
        """Procesa todos los usuarios del CSV"""
        logger.info("\n" + "=" * 80)
        logger.info("🚀 INICIANDO SISTEMA DE FICHAJE")
        logger.info("=" * 80 + "\n")

        if callback:
            callback("Iniciando sistema de fichaje...")

        # Verificar CSV
        if not os.path.exists(self.config['csv_file']):
            msg = f"❌ No existe {self.config['csv_file']}"
            logger.error(msg)
            if callback:
                callback(msg)
            return {'exitos': 0, 'fallos': 0, 'desconocidos': 0, 'total': 0}

        # Cargar datos
        try:
            df = pd.read_csv(self.config['csv_file'])
            if df.empty:
                logger.error("❌ El CSV está vacío")
                if callback:
                    callback("❌ El CSV está vacío")
                return {'exitos': 0, 'fallos': 0, 'desconocidos': 0, 'total': 0}
            logger.info(f"✅ Cargados {len(df)} usuarios")
            if callback:
                callback(f"✅ Cargados {len(df)} usuarios")
        except Exception as e:
            logger.error(f"❌ Error leyendo CSV: {e}")
            if callback:
                callback(f"❌ Error leyendo CSV: {e}")
            return {'exitos': 0, 'fallos': 0, 'desconocidos': 0, 'total': 0}

        # Enviar notificación de inicio
        self.notifier.notify(
            "Inicio de Proceso de Fichaje",
            f"Iniciando proceso para {len(df)} usuarios\nFecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
            tipo="info"
        )

        # Procesar
        driver = None
        exitos = 0
        fallos = 0
        desconocidos = 0

        for i, row in df.iterrows():
            usuario = str(row["tarjeta"]).strip()
            password = str(row["contrasena"]).strip()

            logger.info(f"\n{'=' * 80}")
            logger.info(f"📋 USUARIO {i + 1}/{len(df)}: {usuario}")
            logger.info(f"{'=' * 80}")

            if callback:
                callback(f"\n{'=' * 60}\n📋 Procesando {i + 1}/{len(df)}: {usuario}\n{'=' * 60}")

            try:
                if driver:
                    try:
                        driver.quit()
                    except:
                        pass
                    time.sleep(2)

                driver = self.start_driver(self.config['headless'])
                resultado = self.realizar_fichaje(usuario, password, driver, callback)

                if resultado is True:
                    exitos += 1
                elif resultado is False:
                    fallos += 1
                else:
                    desconocidos += 1

            except Exception as e:
                logger.error(f"❌ Error crítico procesando {usuario}: {e}")
                if callback:
                    callback(f"❌ Error crítico: {str(e)[:50]}")
                fallos += 1
                self.guardar_resultado(usuario, "ERROR", f"Error crítico: {str(e)[:100]}", "")

            if i < len(df) - 1:
                logger.info("⏸ Pausa de 3 segundos...")
                time.sleep(3)

        # Cerrar driver
        if driver:
            try:
                driver.quit()
                logger.info("🔚 Driver cerrado")
            except:
                pass

        # Resumen
        logger.info("\n" + "=" * 80)
        logger.info("📊 RESUMEN FINAL")
        logger.info("=" * 80)
        logger.info(f"✅ Exitosos: {exitos}")
        logger.info(f"❌ Fallidos: {fallos}")
        logger.info(f"⚠️ Desconocidos: {desconocidos}")
        logger.info(f"📁 Total: {len(df)}")
        logger.info("=" * 80 + "\n")

        if callback:
            callback(f"\n{'=' * 60}\n📊 RESUMEN FINAL\n{'=' * 60}")
            callback(f"✅ Exitosos: {exitos}")
            callback(f"❌ Fallidos: {fallos}")
            callback(f"⚠️ Desconocidos: {desconocidos}")
            callback(f"📁 Total procesados: {len(df)}")
            callback(f"📄 Resultados en: {self.config['results_file']}")
            callback(f"{'=' * 60}\n")

        # Enviar notificación de resumen
        resumen_msg = f"""Total procesados: {len(df)}
✅ Exitosos: {exitos}
❌ Fallidos: {fallos}
⚠️ Desconocidos: {desconocidos}

Archivo de resultados: {self.config['results_file']}
Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"""

        tipo_resumen = "success" if fallos == 0 else "warning" if exitos > 0 else "error"
        self.notifier.notify("Resumen de Fichajes", resumen_msg, tipo=tipo_resumen)

        return {'exitos': exitos, 'fallos': fallos, 'desconocidos': desconocidos, 'total': len(df)}


# ==================== INTERFAZ GRÁFICA ====================
class FichajeGUI:
    """Interfaz gráfica con programador de horarios"""

    def __init__(self, root, engine):
        self.root = root
        self.engine = engine
        self.root.title("Sistema de Fichaje Automatizado v2.1 - Con Notificaciones")
        self.root.geometry("950x750")  # Aumentado para acomodar más elementos

        # Variables
        self.horarios = []
        self.scheduler_running = False
        self.scheduler_thread = None

        # Cargar configuración
        self.cargar_configuracion()

        # Crear interfaz
        self.crear_interfaz()

        # Protocolo de cierre
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def crear_interfaz(self):
        """Crea todos los elementos de la interfaz"""

        # TÍTULO
        frame_titulo = tk.Frame(self.root, bg="#2c3e50", pady=15)
        frame_titulo.pack(fill=tk.X)

        tk.Label(frame_titulo,
                 text="🕒 SISTEMA DE FICHAJE AUTOMATIZADO",
                 font=("Arial", 16, "bold"),
                 bg="#2c3e50",
                 fg="white").pack()

        tk.Label(frame_titulo,
                 text="Versión 2.1 - Con Notificaciones (Telegram & Email)",
                 font=("Arial", 10),
                 bg="#2c3e50",
                 fg="#ecf0f1").pack()

        # PANEL SUPERIOR: Configuración de archivo
        frame_config = tk.LabelFrame(self.root,
                                     text="⚙️ Configuración",
                                     font=("Arial", 10, "bold"),
                                     padx=10, pady=10)
        frame_config.pack(fill=tk.X, padx=10, pady=5)

        frame_csv = tk.Frame(frame_config)
        frame_csv.pack(fill=tk.X, pady=5)

        tk.Label(frame_csv, text="Archivo CSV:", font=("Arial", 9)).pack(side=tk.LEFT, padx=5)
        self.csv_entry = tk.Entry(frame_csv, width=40, font=("Arial", 9))
        self.csv_entry.insert(0, CONFIG['csv_file'])
        self.csv_entry.pack(side=tk.LEFT, padx=5)

        tk.Button(frame_csv, text="📁 Examinar", command=self.seleccionar_csv,
                  bg="#3498db", fg="white", font=("Arial", 8), padx=10).pack(side=tk.LEFT)

        # Checkbox para modo Headless
        frame_headless = tk.Frame(frame_config)
        frame_headless.pack(fill=tk.X, pady=5)

        self.headless_var = tk.BooleanVar(value=CONFIG['headless'])
        tk.Checkbutton(frame_headless,
                       text="🔇 Modo Headless (Chrome invisible - recomendado para producción)",
                       variable=self.headless_var,
                       font=("Arial", 9),
                       command=self.toggle_headless).pack(side=tk.LEFT, padx=5)

        # Estado de notificaciones
        notif_status = []
        if self.engine.notifier.telegram_enabled:
            notif_status.append("📱 Telegram")
        if self.engine.notifier.email_enabled:
            notif_status.append("📧 Email")

        if notif_status:
            notif_text = "Notificaciones activas: " + " + ".join(notif_status)
            color = "#27ae60"
        else:
            notif_text = "⚠️ Notificaciones desactivadas (edita notificaciones.ini)"
            color = "#e67e22"

        tk.Label(frame_config,
                 text=notif_text,
                 font=("Arial", 9, "bold"),
                 fg=color).pack(pady=5)

        # HORARIOS
        frame_horarios = tk.LabelFrame(self.root,
                                       text="⏰ Programación de Horarios y Días",
                                       font=("Arial", 11, "bold"),
                                       padx=15, pady=15)
        frame_horarios.pack(fill=tk.X, padx=10, pady=10)

        # Selector de hora
        frame_selector = tk.Frame(frame_horarios)
        frame_selector.pack(fill=tk.X, pady=5)

        tk.Label(frame_selector,
                 text="Hora:",
                 font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)

        self.hora_var = tk.StringVar(value="07")
        tk.Spinbox(frame_selector,
                   from_=0, to=23,
                   textvariable=self.hora_var,
                   width=5,
                   font=("Arial", 11, "bold"),
                   format="%02.0f").pack(side=tk.LEFT, padx=2)

        tk.Label(frame_selector, text=":", font=("Arial", 14, "bold")).pack(side=tk.LEFT)

        self.minuto_var = tk.StringVar(value="00")
        tk.Spinbox(frame_selector,
                   from_=0, to=59,
                   textvariable=self.minuto_var,
                   width=5,
                   font=("Arial", 11, "bold"),
                   format="%02.0f").pack(side=tk.LEFT, padx=2)

        # Selector de días de la semana
        frame_dias = tk.Frame(frame_horarios)
        frame_dias.pack(fill=tk.X, pady=10)

        tk.Label(frame_dias,
                 text="Días de la semana:",
                 font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)

        # Variables para cada día
        self.dias_vars = {
            'L': tk.BooleanVar(value=True),  # Lunes
            'M': tk.BooleanVar(value=True),  # Martes
            'X': tk.BooleanVar(value=True),  # Miércoles
            'J': tk.BooleanVar(value=True),  # Jueves
            'V': tk.BooleanVar(value=True),  # Viernes
            'S': tk.BooleanVar(value=False),  # Sábado
            'D': tk.BooleanVar(value=False)  # Domingo
        }

        dias_nombres = {
            'L': 'Lun',
            'M': 'Mar',
            'X': 'Mié',
            'J': 'Jue',
            'V': 'Vie',
            'S': 'Sáb',
            'D': 'Dom'
        }

        for key, nombre in dias_nombres.items():
            tk.Checkbutton(frame_dias,
                           text=nombre,
                           variable=self.dias_vars[key],
                           font=("Arial", 9)).pack(side=tk.LEFT, padx=3)

        # Botón para seleccionar/deseleccionar todos
        tk.Button(frame_dias,
                  text="Todos",
                  command=lambda: self.seleccionar_todos_dias(True),
                  bg="#95a5a6",
                  fg="white",
                  font=("Arial", 8),
                  padx=5).pack(side=tk.LEFT, padx=5)

        tk.Button(frame_dias,
                  text="Ninguno",
                  command=lambda: self.seleccionar_todos_dias(False),
                  bg="#95a5a6",
                  fg="white",
                  font=("Arial", 8),
                  padx=5).pack(side=tk.LEFT, padx=2)

        # Botón añadir horario
        frame_boton_add = tk.Frame(frame_horarios)
        frame_boton_add.pack(pady=10)

        tk.Button(frame_boton_add,
                  text="➕ Añadir Horario con Días Seleccionados",
                  command=self.anadir_horario,
                  bg="#27ae60",
                  fg="white",
                  font=("Arial", 10, "bold"),
                  cursor="hand2",
                  padx=15,
                  pady=8).pack()

        # Lista de horarios
        frame_lista = tk.Frame(frame_horarios)
        frame_lista.pack(fill=tk.BOTH, expand=True, pady=10)

        tk.Label(frame_lista,
                 text="📋 Horarios programados:",
                 font=("Arial", 10, "bold")).pack(anchor=tk.W)

        # Listbox para horarios
        frame_listbox = tk.Frame(frame_lista)
        frame_listbox.pack(fill=tk.BOTH, expand=True)

        self.lista_horarios = tk.Listbox(frame_listbox,
                                         height=5,
                                         font=("Courier", 10),
                                         selectmode=tk.SINGLE)
        self.lista_horarios.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(frame_listbox, orient=tk.VERTICAL,
                                 command=self.lista_horarios.yview)
        self.lista_horarios.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        tk.Button(frame_horarios,
                  text="🗑 Eliminar Horario Seleccionado",
                  command=self.eliminar_horario,
                  bg="#e74c3c",
                  fg="white",
                  font=("Arial", 9, "bold"),
                  cursor="hand2",
                  padx=10).pack(pady=5)

        # CONTROL
        frame_control = tk.LabelFrame(self.root,
                                      text="🎮 Control de Ejecución",
                                      font=("Arial", 11, "bold"),
                                      padx=15, pady=15)
        frame_control.pack(fill=tk.X, padx=10, pady=10)

        frame_botones = tk.Frame(frame_control)
        frame_botones.pack(pady=10)

        self.btn_ejecutar = tk.Button(frame_botones,
                                      text="▶️ Ejecutar Fichaje AHORA",
                                      command=self.ejecutar_ahora,
                                      bg="#3498db",
                                      fg="white",
                                      font=("Arial", 11, "bold"),
                                      cursor="hand2",
                                      padx=20,
                                      pady=10,
                                      width=22)
        self.btn_ejecutar.pack(side=tk.LEFT, padx=5)

        self.btn_scheduler = tk.Button(frame_botones,
                                       text="🚀 Activar Programador",
                                       command=self.toggle_scheduler,
                                       bg="#f39c12",
                                       fg="white",
                                       font=("Arial", 11, "bold"),
                                       cursor="hand2",
                                       padx=20,
                                       pady=10,
                                       width=22)
        self.btn_scheduler.pack(side=tk.LEFT, padx=5)

        # Estado
        self.label_estado = tk.Label(frame_control,
                                     text="⏸ Estado: PROGRAMADOR INACTIVO",
                                     font=("Arial", 11, "bold"),
                                     fg="#e74c3c")
        self.label_estado.pack(pady=10)

        self.label_proxima = tk.Label(frame_control,
                                      text="",
                                      font=("Arial", 9),
                                      fg="#7f8c8d")
        self.label_proxima.pack()

        # CONSOLA
        frame_consola = tk.LabelFrame(self.root,
                                      text="📝 Consola de Ejecución en Tiempo Real",
                                      font=("Arial", 11, "bold"),
                                      padx=10, pady=10)
        frame_consola.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.consola = scrolledtext.ScrolledText(frame_consola,
                                                 wrap=tk.WORD,
                                                 width=100,
                                                 height=10,
                                                 font=("Consolas", 9),
                                                 bg="#1e1e1e",
                                                 fg="#00ff00",
                                                 insertbackground="white")
        self.consola.pack(fill=tk.BOTH, expand=True)

        frame_botones_consola = tk.Frame(frame_consola)
        frame_botones_consola.pack(pady=5)

        tk.Button(frame_botones_consola,
                  text="🧹 Limpiar Consola",
                  command=self.limpiar_consola,
                  bg="#95a5a6",
                  fg="white",
                  font=("Arial", 9),
                  cursor="hand2",
                  padx=10).pack(side=tk.LEFT, padx=5)

        tk.Button(frame_botones_consola,
                  text="📁 Abrir Carpeta de Screenshots",
                  command=self.abrir_screenshots,
                  bg="#9b59b6",
                  fg="white",
                  font=("Arial", 9),
                  cursor="hand2",
                  padx=10).pack(side=tk.LEFT, padx=5)

        # Mensaje inicial
        self.log_consola("=" * 80)
        self.log_consola("🚀 Sistema de Fichaje Automatizado v2.1 iniciado")
        self.log_consola(f"📅 Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        self.log_consola(f"📂 Carpeta de trabajo: {os.getcwd()}")
        self.log_consola("=" * 80)
        self.log_consola("")
        self.log_consola("📋 INSTRUCCIONES:")
        self.log_consola("1. Verifica que existe el archivo 'datos.csv'")
        self.log_consola("2. Configura notificaciones en 'notificaciones.ini'")
        self.log_consola("3. Selecciona hora, minutos y días de la semana")
        self.log_consola("4. Añade horarios de fichaje automático")
        self.log_consola("5. Activa el programador o ejecuta manualmente")
        self.log_consola("")
        self.log_consola("✨ CARACTERÍSTICAS:")
        self.log_consola("   • Selección de días específicos (L-D)")
        self.log_consola("   • Modo Headless (Chrome invisible)")
        self.log_consola("   • Notificaciones por Telegram")
        self.log_consola("   • Notificaciones por Email")
        self.log_consola("=" * 80)

        # Cargar horarios
        self.actualizar_lista_horarios()

    def log_consola(self, mensaje):
        """Añade mensaje a la consola"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.consola.insert(tk.END, f"[{timestamp}] {mensaje}\n")
        self.consola.see(tk.END)
        self.root.update_idletasks()

    def limpiar_consola(self):
        """Limpia la consola"""
        self.consola.delete(1.0, tk.END)
        self.log_consola("Consola limpiada")

    def seleccionar_csv(self):
        """Abre diálogo para seleccionar archivo CSV"""
        filename = filedialog.askopenfilename(
            title="Seleccionar archivo CSV",
            filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")]
        )
        if filename:
            self.csv_entry.delete(0, tk.END)
            self.csv_entry.insert(0, filename)
            CONFIG['csv_file'] = filename
            self.log_consola(f"✅ Archivo CSV seleccionado: {filename}")

    def abrir_screenshots(self):
        """Abre la carpeta de screenshots"""
        try:
            screenshot_path = os.path.abspath(CONFIG['screenshots_dir'])
            if sys.platform == 'win32':
                os.startfile(screenshot_path)
            elif sys.platform == 'darwin':
                os.system(f'open "{screenshot_path}"')
            else:
                os.system(f'xdg-open "{screenshot_path}"')
            self.log_consola(f"📁 Abriendo carpeta: {screenshot_path}")
        except Exception as e:
            self.log_consola(f"❌ Error abriendo carpeta: {e}")

    def toggle_headless(self):
        """Activa/desactiva el modo headless"""
        CONFIG['headless'] = self.headless_var.get()
        estado = "activado" if CONFIG['headless'] else "desactivado"
        self.log_consola(f"🔇 Modo Headless {estado}")

    def seleccionar_todos_dias(self, seleccionar):
        """Selecciona o deselecciona todos los días"""
        for var in self.dias_vars.values():
            var.set(seleccionar)

    def anadir_horario(self):
        """Añade un horario a la lista con días específicos"""
        hora = self.hora_var.get().zfill(2)
        minuto = self.minuto_var.get().zfill(2)
        horario = f"{hora}:{minuto}"

        # Obtener días seleccionados
        dias_seleccionados = [dia for dia, var in self.dias_vars.items() if var.get()]

        if not dias_seleccionados:
            messagebox.showwarning("Sin Días Seleccionados",
                                   "Debes seleccionar al menos un día de la semana")
            return

        # Crear identificador único con horario y días
        dias_str = ''.join(dias_seleccionados)
        id_horario = f"{horario}_{dias_str}"

        # Verificar que no existe ya esta combinación exacta
        if any(h["horario"] == horario and h["dias_str"] == dias_str for h in self.horarios):
            messagebox.showwarning("Horario Duplicado",
                                   f"El horario {horario} con estos días ya existe")
            return

        nuevo_horario = {
            "horario": horario,
            "dias": dias_seleccionados,
            "dias_str": dias_str,
            "activo": True,
            "ultima_ejecucion": "Nunca"
        }
        self.horarios.append(nuevo_horario)

        dias_texto = ', '.join(dias_seleccionados)
        self.log_consola(f"✅ Horario añadido: {horario} - Días: {dias_texto}")
        self.actualizar_lista_horarios()
        self.guardar_configuracion()

        if self.scheduler_running:
            self.programar_tareas()

    def eliminar_horario(self):
        """Elimina el horario seleccionado"""
        seleccion = self.lista_horarios.curselection()
        if not seleccion:
            messagebox.showinfo("Sin Selección", "Selecciona un horario para eliminar")
            return

        index = seleccion[0]
        horario = self.horarios[index]["horario"]

        if messagebox.askyesno("Confirmar Eliminación",
                               f"¿Eliminar el horario {horario}?"):
            self.horarios.pop(index)
            self.log_consola(f"🗑 Horario eliminado: {horario}")
            self.actualizar_lista_horarios()
            self.guardar_configuracion()

            if self.scheduler_running:
                self.programar_tareas()

    def actualizar_lista_horarios(self):
        """Actualiza la lista visual de horarios con días"""
        self.lista_horarios.delete(0, tk.END)

        for h in sorted(self.horarios, key=lambda x: x["horario"]):
            estado = "✅ ACTIVO" if h["activo"] else "❌ INACTIVO"
            dias = ', '.join(h.get("dias", ["L", "M", "X", "J", "V", "S", "D"]))
            texto = f"{h['horario']} - Días:[{dias}] - {estado} - Última:{h['ultima_ejecucion']}"
            self.lista_horarios.insert(tk.END, texto)

    def ejecutar_ahora(self):
        """Ejecuta el fichaje inmediatamente"""
        # Actualizar CSV si cambió
        CONFIG['csv_file'] = self.csv_entry.get()
        # Actualizar headless desde checkbox
        CONFIG['headless'] = self.headless_var.get()

        self.log_consola("=" * 80)
        self.log_consola("▶️ EJECUTANDO FICHAJE MANUAL...")
        if CONFIG['headless']:
            self.log_consola("🔇 Modo: Headless (Chrome invisible)")
        else:
            self.log_consola("👁️ Modo: Visible (Chrome con ventana)")
        self.log_consola("=" * 80)

        self.btn_ejecutar.config(state=tk.DISABLED, text="⏳ Ejecutando...")

        thread = threading.Thread(target=self._ejecutar_thread)
        thread.daemon = True
        thread.start()

    def _ejecutar_thread(self):
        """Ejecuta el fichaje en thread separado"""
        try:
            resultado = self.engine.procesar_usuarios(callback=self.log_consola)
            self.log_consola("✅ Ejecución completada")
        except Exception as e:
            self.log_consola(f"❌ ERROR: {e}")
        finally:
            self.btn_ejecutar.config(state=tk.NORMAL, text="▶️ Ejecutar Fichaje AHORA")

    def toggle_scheduler(self):
        """Activa/desactiva el programador"""
        if self.scheduler_running:
            self.detener_scheduler()
        else:
            self.iniciar_scheduler()

    def iniciar_scheduler(self):
        """Inicia el programador de tareas"""
        if not self.horarios:
            messagebox.showwarning("Sin Horarios",
                                   "Añade al menos un horario antes de activar el programador")
            return

        self.scheduler_running = True
        self.btn_scheduler.config(text="⏸ Detener Programador", bg="#e74c3c")
        self.label_estado.config(text="✅ Estado: PROGRAMADOR ACTIVO", fg="#27ae60")

        self.log_consola("=" * 80)
        self.log_consola("🚀 PROGRAMADOR DE TAREAS ACTIVADO")
        self.log_consola("=" * 80)

        # Enviar notificación de activación
        self.engine.notifier.notify(
            "Programador Activado",
            f"El programador de tareas ha sido activado\nHorarios configurados: {len(self.horarios)}\nFecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
            tipo="info"
        )

        self.programar_tareas()

        self.scheduler_thread = threading.Thread(target=self._run_scheduler)
        self.scheduler_thread.daemon = True
        self.scheduler_thread.start()

    def detener_scheduler(self):
        """Detiene el programador de tareas"""
        self.scheduler_running = False
        schedule.clear()

        self.btn_scheduler.config(text="🚀 Activar Programador", bg="#f39c12")
        self.label_estado.config(text="⏸ Estado: PROGRAMADOR INACTIVO", fg="#e74c3c")
        self.label_proxima.config(text="")

        self.log_consola("=" * 80)
        self.log_consola("⏸ PROGRAMADOR DE TAREAS DETENIDO")
        self.log_consola("=" * 80)

        # Enviar notificación de desactivación
        self.engine.notifier.notify(
            "Programador Detenido",
            f"El programador de tareas ha sido detenido\nFecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
            tipo="info"
        )

    def programar_tareas(self):
        """Programa todas las tareas con días específicos"""
        schedule.clear()

        # Mapeo de días a funciones de schedule
        dias_schedule = {
            'L': 'monday',
            'M': 'tuesday',
            'X': 'wednesday',
            'J': 'thursday',
            'V': 'friday',
            'S': 'saturday',
            'D': 'sunday'
        }

        for h in self.horarios:
            if h["activo"]:
                horario = h["horario"]
                dias = h.get("dias", ["L", "M", "X", "J", "V", "S", "D"])

                for dia_letra in dias:
                    dia_schedule = dias_schedule.get(dia_letra)
                    if dia_schedule:
                        # Programar para cada día específico
                        getattr(schedule.every(), dia_schedule).at(horario).do(
                            self._tarea_programada, horario, h
                        )

                dias_texto = ', '.join(dias)
                self.log_consola(f"⏰ Programado: {horario} - Días: {dias_texto}")

        if schedule.jobs:
            proxima = schedule.next_run()
            if proxima:
                texto_proxima = proxima.strftime('%d/%m/%Y %H:%M:%S')
                dia_semana = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'][
                    proxima.weekday()]
                self.log_consola(f"📅 Próxima ejecución: {dia_semana} {texto_proxima}")
                self.label_proxima.config(text=f"Próxima: {dia_semana} {texto_proxima}")

    def _tarea_programada(self, horario_str, horario_obj):
        """Función que se ejecuta cuando llega la hora"""
        dias_texto = ', '.join(horario_obj.get("dias", []))
        self.log_consola("=" * 80)
        self.log_consola(f"⏰ TAREA PROGRAMADA - {horario_str} - Días: [{dias_texto}]")
        self.log_consola("=" * 80)

        horario_obj["ultima_ejecucion"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        self.actualizar_lista_horarios()
        self.guardar_configuracion()

        self._ejecutar_thread()

    def _run_scheduler(self):
        """Loop del scheduler"""
        while self.scheduler_running:
            schedule.run_pending()
            time.sleep(1)

    def cargar_configuracion(self):
        """Carga la configuración guardada"""
        try:
            config_file = CONFIG['config_file']
            if os.path.exists(config_file):
                with open(config_file, 'r') as f:
                    data = json.load(f)
                    self.horarios = data.get("horarios", [])
        except Exception as e:
            self.horarios = []

    def guardar_configuracion(self):
        """Guarda la configuración"""
        try:
            with open(CONFIG['config_file'], 'w') as f:
                json.dump({"horarios": self.horarios}, f, indent=4)
        except Exception as e:
            print(f"Error guardando config: {e}")

    def on_closing(self):
        """Maneja el cierre de la aplicación"""
        if self.scheduler_running:
            respuesta = messagebox.askyesno(
                "Programador Activo",
                "El programador está activo.\n\n"
                "¿Seguro que quieres cerrar?\n"
                "Las tareas programadas se detendrán."
            )
            if not respuesta:
                return
            self.detener_scheduler()

        self.guardar_configuracion()
        self.root.destroy()


# ==================== FUNCIÓN PRINCIPAL ====================
def main():
    """Función principal que inicia la aplicación"""

    # Crear motor de fichaje
    engine = FichajeEngine(CONFIG)

    # Crear ventana principal
    root = tk.Tk()

    # Crear GUI
    app = FichajeGUI(root, engine)

    # Iniciar bucle de eventos
    root.mainloop()


if __name__ == "__main__":
    main()
