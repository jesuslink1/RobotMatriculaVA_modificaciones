import os
from dotenv import load_dotenv

from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

load_dotenv()

usuario = os.getenv("SISE_USUARIO")
contrasena = os.getenv("SISE_PASSWORD")

URL_ACADEMICO = "https://academico.sise.edu.pe/psp/CS90PRD/?cmd=login"


def iniciar_sesion_academica(driver, timeout=10):
    """Abre el portal académico e inicia sesión con credenciales del entorno."""

    usuario = os.getenv("SISE_USUARIO")
    contrasena = os.getenv("SISE_PASSWORD")

    driver.get(URL_ACADEMICO)

    campo_usuario = WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located((By.ID, "userid"))
    )
    campo_usuario.clear()
    campo_usuario.send_keys(usuario)

    campo_contrasena = driver.find_element(By.ID, "pwd")
    campo_contrasena.clear()
    campo_contrasena.send_keys(contrasena)

    boton_conectar = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button.btn-danger-sise")
        )
    )
    boton_conectar.click()


def abrir_centro_servicios(driver, timeout=10):
    """Navega hasta Centro Servicios al Alumnado y entra en su iframe."""

    WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "a#fldra_MYFAVORITES"))
    ).click()

    WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable(
            (By.LINK_TEXT, "Centro Servicios al Alumnado")
        )
    ).click()

    WebDriverWait(driver, timeout).until(
        EC.frame_to_be_available_and_switch_to_it("ptifrmtgtframe")
    )
