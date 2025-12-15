# app.py
import streamlit as st
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from datetime import datetime, timedelta, date
from openpyxl import Workbook
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
import time
import re

# =========================
# CONFIGURACIÓN BASE
# =========================
LOGIN_URL = "https://app.citytroops.com/"
TRACKING_URL_TEMPLATE = "https://app.citytroops.com/tracking?user={user_id}&date={date}"

DEFAULT_USER_IDS = [19901, 19904, 19907, 19906, 19900, 19905, 19895, 19896, 19903, 19908]

DEFAULT_COLABORADORES = {
    19901: "Carlos Heredia",
    19904: "Cinthia Palacio",
    19907: "Erick Picazo",
    19906: "Flor Llanas",
    19900: "Horacio Benitez",
    19905: "Julissa Anguiano",
    19895: "Laura Leyja",
    19896: "Luis Villafuerte",
    19903: "Maria Trinidad Hernandez",
    19908: "Yessica Rodriguez",
}


# =========================
# SELENIUM HELPERS
# =========================
def crear_driver(headless: bool, chrome_binary: str | None = None, chromedriver_path: str | None = None):
    options = webdriver.ChromeOptions()
    options.page_load_strategy = "eager"

    # Recomendado en entornos “server”
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    if headless:
        # Chrome moderno
        options.add_argument("--headless=new")

    if chrome_binary:
        options.binary_location = chrome_binary

    if chromedriver_path:
        driver = webdriver.Chrome(service=webdriver.chrome.service.Service(chromedriver_path), options=options)
    else:
        driver = webdriver.Chrome(options=options)

    driver.set_page_load_timeout(25)
    return driver


def login(driver, email: str, password: str):
    wait = WebDriverWait(driver, 25)
    driver.get(LOGIN_URL)

    wait.until(EC.presence_of_element_located((By.NAME, "email"))).send_keys(email)
    driver.find_element(By.NAME, "password").send_keys(password)
    driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

    wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//*[contains(text(),'Team tracking') or contains(text(),'Dashboard')]")
        )
    )


def scrap_user(user_id: int, colaborador: str, fecha_inicial: date, fecha_final: date,
              email: str, password: str, headless: bool,
              chrome_binary: str | None, chromedriver_path: str | None):

    resultados = []
    driver = crear_driver(headless=headless, chrome_binary=chrome_binary, chromedriver_path=chromedriver_path)
    wait = WebDriverWait(driver, 8)

    try:
        login(driver, email, password)

        fecha = fecha_inicial
        while fecha <= fecha_final:
            fecha_str = fecha.strftime("%Y-%m-%d")
            url = TRACKING_URL_TEMPLATE.format(user_id=user_id, date=fecha_str)

            try:
                driver.get(url)
            except Exception:
                resultados.append((user_id, colaborador, fecha_str, 0.0))
                fecha += timedelta(days=1)
                continue

            try:
                km_elem = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//*[contains(., 'Distance traveled')]"))
                )
                km_text = km_elem.text
                match = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*Km", km_text)
                km_value = float(match.group(1)) if match else 0.0
                resultados.append((user_id, colaborador, fecha_str, km_value))

            except TimeoutException:
                resultados.append((user_id, colaborador, fecha_str, 0.0))

            fecha += timedelta(days=1)

    finally:
        driver.quit()

    return resultados


def construir_excel_bytes(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "KM Recorridos"
    ws.append(["Usuario", "Colaborador", "Fecha", "Kilómetros"])

    for row in sorted(rows, key=lambda x: (x[0], x[2])):
        ws.append(list(row))

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# =========================
# UI STREAMLIT
# =========================
st.set_page_config(page_title="CityTroops KM Scraper", layout="wide")
st.title("📍 CityTroops | KM recorridos (Selenium)")

with st.sidebar:
    st.header("⚙️ Configuración")

    # Usa secrets si están disponibles
    default_email = st.secrets.get("CITYTROOPS_EMAIL", "")
    default_pass = st.secrets.get("CITYTROOPS_PASSWORD", "")

    email = st.text_input("Email", value=default_email)
    password = st.text_input("Password", value=default_pass, type="password")

    headless = st.toggle("Headless (recomendado)", value=True)

    st.caption("Opcional (si tu entorno lo requiere):")
    chrome_binary = st.text_input("Ruta Chrome binary (opcional)", value="")
    chromedriver_path = st.text_input("Ruta chromedriver (opcional)", value="")

    max_workers = st.slider("Paralelismo (drivers)", min_value=1, max_value=6, value=3)

st.subheader("👥 Usuarios y fechas")

col1, col2 = st.columns(2)
with col1:
    user_ids_txt = st.text_area(
        "USER_IDS (uno por línea)",
        value="\n".join(str(x) for x in DEFAULT_USER_IDS),
        height=180
    )
with col2:
    colaboradores_txt = st.text_area(
        "Mapa COLABORADORES (formato: user_id=Nombre, uno por línea)",
        value="\n".join([f"{k}={v}" for k, v in DEFAULT_COLABORADORES.items()]),
        height=180
    )

dcol1, dcol2 = st.columns(2)
with dcol1:
    fecha_inicial = st.date_input("Fecha inicial", value=date(2025, 11, 11))
with dcol2:
    fecha_final = st.date_input("Fecha final", value=date(2025, 12, 6))

run = st.button("🚀 Ejecutar scraping", type="primary")

# =========================
# RUN
# =========================
if run:
    if not email or not password:
        st.error("Falta email o password.")
        st.stop()

    # Parse USER_IDS
    try:
        user_ids = [int(x.strip()) for x in user_ids_txt.splitlines() if x.strip()]
    except Exception:
        st.error("USER_IDS inválidos. Asegúrate de que cada línea sea un número.")
        st.stop()

    # Parse COLABORADORES
    colaboradores = {}
    try:
        for line in colaboradores_txt.splitlines():
            line = line.strip()
            if not line:
                continue
            if "=" not in line:
                continue
            k, v = line.split("=", 1)
            colaboradores[int(k.strip())] = v.strip()
    except Exception:
        st.error("Mapa COLABORADORES inválido. Usa formato: 19901=Nombre")
        st.stop()

    if fecha_final < fecha_inicial:
        st.error("La fecha final no puede ser menor que la fecha inicial.")
        st.stop()

    total_dias = (fecha_final - fecha_inicial).days + 1
    total_trabajo = len(user_ids) * total_dias

    st.info(f"Trabajo estimado: **{len(user_ids)} usuarios × {total_dias} días = {total_trabajo} registros**")

    progress = st.progress(0)
    log_box = st.empty()
    logs = []
    all_rows = []

    inicio = time.time()

    # Para actualizar progreso “por usuario” (aprox).
    completados = 0

    try:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = []
            for uid in user_ids:
                colaborador = colaboradores.get(uid, "Desconocido")
                futures.append(
                    executor.submit(
                        scrap_user,
                        uid, colaborador,
                        fecha_inicial, fecha_final,
                        email, password,
                        headless,
                        chrome_binary.strip() or None,
                        chromedriver_path.strip() or None
                    )
                )

            for f in as_completed(futures):
                try:
                    filas = f.result()
                    all_rows.extend(filas)
                    completados += 1
                    pct = completados / max(1, len(user_ids))
                    progress.progress(min(1.0, pct))

                    logs.append(f"✔ Usuario completado: {filas[0][0] if filas else 'N/A'} → {len(filas)} filas")
                    log_box.code("\n".join(logs[-20:]))  # últimos 20 logs

                except Exception as e:
                    completados += 1
                    pct = completados / max(1, len(user_ids))
                    progress.progress(min(1.0, pct))
                    logs.append(f"❌ Error en un usuario: {e}")
                    log_box.code("\n".join(logs[-20:]))

    except Exception as e:
        st.error(f"Error general: {e}")
        st.stop()

    fin = time.time()
    st.success(f"✅ Listo. Tiempo total: {fin - inicio:.1f} segundos | Filas: {len(all_rows)}")

    # Excel en memoria
    xlsx_bytes = construir_excel_bytes(all_rows)

    filename = f"km_recorridos_{fecha_inicial.strftime('%Y%m%d')}_a_{fecha_final.strftime('%Y%m%d')}.xlsx"
    st.download_button(
        label="⬇️ Descargar Excel",
        data=xlsx_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Vista rápida
    st.subheader("🔎 Vista rápida (primeras 50 filas)")
    st.dataframe(all_rows[:50], use_container_width=True)
