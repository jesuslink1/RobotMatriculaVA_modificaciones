import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException


# Leer los códigos de alumno desde un archivo Excel
df = pd.read_excel('C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\Cambio_plan.xlsx', sheet_name='Hoja1')
codigos_alumno = df['Codigo'].tolist()

# Configura el objeto Service para el controlador de Chrome
chrome_service = webdriver.chrome.service.Service('C:\\Users\\jramosc\\Downloads\\chromedriver-win64\\Chromedriver.exe')

# Inicia el controlador de Chrome
chrome_service.start()

# Configura las opciones del navegador
chrome_options = webdriver.ChromeOptions()
# Opcional: Puedes agregar opciones adicionales según tus necesidades, como la ejecución en segundo plano o la omisión de notificaciones.

# Crea una instancia del controlador de Chrome
driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

# Abre la página web
driver.get('https://academico.sise.edu.pe/psp/CS90PRD/?cmd=login')

#################################################################
# LOGUEO DE USUARIO

# Encontrar el elemento de entrada de usuario por su atributo id
input_usuario = driver.find_element(By.ID, "userid")
input_usuario.clear()
input_usuario.send_keys("JVITORM")

# Encontrar el elemento de entrada de contraseña por su atributo id
input_contrasena = driver.find_element(By.ID, "pwd")
input_contrasena.clear()
input_contrasena.send_keys("FevV1511")

# Encontrar el botón de "CONECTAR" por su selector CSS
boton_conectar = driver.find_element(By.CSS_SELECTOR, 'button.btn-danger-sise')
boton_conectar.click()

time.sleep(2)

###################################################################
# INGRESAR A PROGRAMA

# Esperar hasta que se cargue el enlace "Mis Favoritos"
enlace_favoritos = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a#fldra_MYFAVORITES')))
enlace_favoritos.click()

# Esperar hasta que se cargue el enlace "Inscripción Rápida"
enlace_inscripcion = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, "Programa/Plan de Alumno")))
enlace_inscripcion.click()

time.sleep(1)

# Cambiar el contexto al iframe "ptifrmtgtframe"
driver.switch_to.frame("ptifrmtgtframe")

time.sleep(2)

#####################################################
# BUSQUEDA DE CÓDIGOS

# Iterar sobre cada fila del DataFrame
for index, row in df.iterrows():
    # Obtener el código de alumno y el programa académico de la fila actual
    codigo_alumno = row['Codigo']
    programa_academico = row['Programa Académico']

    # Encontrar el elemento de entrada de código de alumno por su atributo id
    input_codigo_alumno = driver.find_element(By.ID, "STDNT_CAR_SRCH_EMPLID")

    # Limpiar el campo de entrada por si acaso ya hay algo escrito
    input_codigo_alumno.clear()

    # Ingresar el código de alumno en el campo de entrada
    input_codigo_alumno.send_keys(codigo_alumno)

    # Encontrar el botón de búsqueda por su ID
    boton_buscar = driver.find_element(By.ID, "#ICSearch")

    # Hacer clic en el botón de búsqueda
    boton_buscar.click()

    # Esperar un tiempo para que la página se cargue completamente antes de buscar el enlace del programa académico
    time.sleep(2)

    programas_academicos = {
        "ADHYR": "ADM. DE HO",
        "ADMBF": "ADM BF",
        "ADMLO": "ADMLOG",
        "ADMNE": "ADMEMPRE",
        "ADMNI": "ADMNI",
        "COMAU": "COAU",
        "CONTA": "CONTABILID",
        "DG": "DG",
        "DYGMO": "DISEÑO Y G",
        "FISIO": "FISIOTERAP",
        "MARKT": "MARKETING",
        "PA": "PER AUDIO",
        "REDES": "REDES Y SI",
        "SALUD": "ENFERTECN",
        "SECRE": "SECRET",
        "SWSIS": "SW Y SIS",
        "TECFA": "TECFARMAC"
    }

    try:
        # Obtener el nombre del programa correspondiente
        nombre_programa = programas_academicos.get(programa_academico)

        # Encuentra todos los elementos 'a' dentro de la tabla con el texto del programa académico
        enlaces_programa = driver.find_elements(By.XPATH,
                                                f"//table[@id='PTSRCHRESULTS']//a[text()='{nombre_programa}']")

        if enlaces_programa:
            # Haz clic en el primer enlace encontrado
            enlaces_programa[0].click()
        else:
            print(f"No se encontró ningún enlace para el programa académico: {nombre_programa}")
    except NoSuchElementException as e:
        print(f"No se encontró el elemento de la tabla: {e}")

    # Esperar un tiempo para que la página se cargue completamente antes de hacer clic en el enlace del plan del alumno
    time.sleep(2)

    # Encontrar el elemento <a> con el id "ICTAB_1" que contiene el texto "Plan del Alumno" y hacer clic en él
    enlace_plan_alumno = driver.find_element(By.ID, "ICTAB_1")
    enlace_plan_alumno.click()

    time.sleep(1)

    # Encontrar el botón "Corregir Historial" por su ID
    boton_corregir_historial = driver.find_element(By.ID, "#ICCorrection")

    # Hacer clic en el botón "Corregir Historial"
    boton_corregir_historial.click()

    time.sleep(2)

    # Mapeo de programas académicos del Excel a los valores correspondientes para el cuadro de texto
    mapeo_programas = {
        'ADHYR': 'ADHYR2024',
        'ADMBF': 'ADBFI2024',
        'ADMLO': 'GELOG2024',
        'ADMNE': 'ADEMP2024',
        'ADMNI': 'ADNIN2024',
        'COMAU': 'COAUD2024',
        'CONTA': 'CONTA2024',
        'DG': 'DISGR2024',
        'DYGMO': 'DYGMO2024',
        'FISIO': 'FISIO2024',
        'GYGEC': 'GYGEC2024',
        'MARKT': 'MARKT2024',
        'PA': 'PEAUD2024',
        'REDES': 'REDES2024',
        'SALUD': 'ENFTE2024',
        'SECRE': 'SECRE2024',
        'SWSIS': 'DSIST2024',
        'TECFA': 'TECFA2024'
    }

    # Obtener el programa académico del DataFrame y encontrar su valor correspondiente en el mapeo
    programa_academico = row['Programa Académico']
    valor_cuadro_texto = mapeo_programas.get(programa_academico)

    # Encontrar el elemento de entrada de texto por su ID
    input_cuadro_texto = driver.find_element(By.ID, "ACAD_PLAN_ACAD_PLAN$0")

    # Limpiar el campo de entrada por si acaso ya hay algo escrito
    input_cuadro_texto.clear()

    # Ingresar el valor correspondiente al programa académico en el campo de entrada
    input_cuadro_texto.send_keys(valor_cuadro_texto)

    # Encontrar el botón "Guardar" por su ID
    boton_guardar = driver.find_element(By.ID, "#ICSave")

    # Hacer clic en el botón "Guardar"
    boton_guardar.click()

    # Esperar un tiempo para que la página se cargue completamente después de guardar
    time.sleep(2)  # Puedes ajustar este tiempo según sea necesario

    # Encontrar el botón "Volver a Buscar" por su ID
    boton_volver_buscar = driver.find_element(By.ID, "#ICList")

    # Hacer clic en el botón "Volver a Buscar"
    boton_volver_buscar.click()

    # Esperar un tiempo para que la página se cargue completamente después de volver a buscar
    time.sleep(2)  # Puedes ajustar este tiempo según sea necesario

    # Encontrar el botón "Borrar" por su ID
    boton_borrar = driver.find_element(By.ID, "#ICClear")

    # Hacer clic en el botón "Borrar"
    boton_borrar.click()

    # Esperar un tiempo para que la página se cargue completamente después de borrar
    time.sleep(2)  # Puedes ajustar este tiempo según sea necesario

######################################################################

# Cerrar el navegador
driver.quit()

# Detener el servicio del controlador de Chrome
chrome_service.stop()