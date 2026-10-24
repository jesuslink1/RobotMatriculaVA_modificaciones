import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from robot_matricula.core.navegador import crear_navegador_chrome


# Leer los códigos de alumno desde un archivo Excel
df = pd.read_excel('C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\Extensión\\CodigoMatExt.xlsx', sheet_name='Hoja1')
codigos_alumno = df['Codigo'].tolist()

# Configura el objeto Service para el controlador de Chrome
chrome_service = webdriver.chrome.service.Service(
        'C:\\Users\\jramosc\\Downloads\\chromedriver-win64\\Chromedriver.exe')

# Crear el navegador usando la función compartida
driver = crear_navegador_chrome()

# Abre la página web
driver.get('https://academico.sise.edu.pe/psp/CS90PRD/?cmd=login')

#################################################################
# LOGUEO DE USUARIO

# Encontrar el elemento de entrada de usuario por su atributo id
input_usuario = driver.find_element(By.ID, "userid")
input_usuario.clear()
input_usuario.send_keys("JRAMOSC")

# Encontrar el elemento de entrada de contraseña por su atributo id
input_contrasena = driver.find_element(By.ID, "pwd")
input_contrasena.clear()
input_contrasena.send_keys("jOLgEt6n")

# Encontrar el botón de "CONECTAR" por su selector CSS
boton_conectar = driver.find_element(By.CSS_SELECTOR, 'button.btn-danger-sise')
boton_conectar.click()

time.sleep(2)

###################################################################
# INGRESAR A INSCRIPCIÓN RAPIDA

# Esperar hasta que se cargue el enlace "Mis Favoritos"
enlace_favoritos = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a#fldra_MYFAVORITES')))
enlace_favoritos.click()

# Esperar hasta que se cargue el enlace "Inscripción Rápida"
enlace_inscripcion = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, "Inscripción Rápida")))
enlace_inscripcion.click()

#####################################################
#INGRESAR ALUMNO

# Cambiar el contexto al iframe "ptifrmtgtframe"
driver.switch_to.frame("ptifrmtgtframe")

# Iterar sobre los códigos de alumno

for codigo in codigos_alumno:

    # Esperar hasta que el elemento de búsqueda de alumno esté presente y sea interactivo
    input_alumno = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[name="QUIK_E_ADD_SRCH_EMPLID"]')))

    # Limpiar el campo de texto
    input_alumno.clear()

    # Ingresar el código de alumno en el campo de texto
    input_alumno.send_keys(str(codigo))

    time.sleep(2)

    # Encontrar el elemento de búsqueda de carrera académica por su atributo id
    input_carrera = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "QUIK_E_ADD_SRCH_ACAD_CAREER")))
    input_carrera.clear()
    input_carrera.send_keys("VIRT")

    # Encontrar el elemento de búsqueda de institución por su atributo id
    input_institucion = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "QUIK_E_ADD_SRCH_INSTITUTION")))
    input_institucion.clear()
    input_institucion.send_keys("SISE")

    # Encontrar el elemento de búsqueda de período académico por su atributo id
    input_periodo = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "QUIK_E_ADD_SRCH_STRM")))
    input_periodo.clear()
    input_periodo.send_keys("2206")

    time.sleep(2)

    # Encontrar el botón "Añadir" por su atributo id
    boton_anadir = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "#ICSearch")))
    boton_anadir.click()

    time.sleep(4)

    # Verificar si el elemento de deuda existe
    try:
        elemento_deuda = driver.find_element(By.ID, "DERIVED_CS_SRVC_IND_NEG")
        if elemento_deuda:
            print(f"Código: {codigo} - Alumno tiene deuda")

            # Volver al iframe normal
            driver.switch_to.default_content()
            time.sleep(1)

            # Hacer clic en "Inscripción Rápida"
            enlace_inscripcion = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.LINK_TEXT, "Inscripción Rápida"))
            )
            enlace_inscripcion.click()

            time.sleep(1)

            # Volver al iframe
            driver.switch_to.frame("ptifrmtgtframe")

            # Continuar con el siguiente código del for
            continue
    except:
        pass  # Si el elemento no existe, continúa con el flujo normal

    ##############################################
    # ingresar números de clase

    # Encontrar el elemento de búsqueda de clase por su atributo id
    input_clase = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ENRL_REQ_DETAIL_CLASS_NBR$0")))
    input_clase.clear()

    # Obtener el código de clase correspondiente
    codigo_clase = df.loc[df['Codigo'] == codigo, 'NºClase1']

    if not codigo_clase.empty:
        codigo_clase = codigo_clase.iloc[0]  # Tomar solo el primer valor
        print(f"Código: {codigo} -> Clase: {codigo_clase}")  # Verificar el valor obtenido
        input_clase.send_keys(str(codigo_clase))
    else:
        print(f"⚠ No se encontró una clase para el código {codigo}")
        continue  # Saltar al siguiente código

    time.sleep(3)

    """# Encontrar el elemento del botón por su atributo id
    boton_add_fila = driver.find_element(By.ID, "ENRL_REQ_DETAIL$new$0$$0")
    boton_add_fila.click()

    time.sleep(3)

    ## Numero de clase 2
    # Obtener el texto de la columna "NºClase2" para la fila actual
    texto_columna_3 = df.loc[df['Codigo'] == codigos_alumno, 'NºClase2'].item()

    # Encontrar el elemento del cuadro de texto por su atributo id
    input_cuadro_texto = driver.find_element(By.ID, "ENRL_REQ_DETAIL_CLASS_NBR$1")

    # Ingresar el valor deseado en el campo de texto
    input_cuadro_texto.send_keys(texto_columna_3)

    # Agregar fila
    boton_add_fila = driver.find_element(By.ID, "ENRL_REQ_DETAIL$new$0$$0")
    boton_add_fila.click()

    time.sleep(2)

    ## Numero de clase 3
    # Obtener el texto de la columna "NºClase3" para la fila actual
    texto_columna_4 = df.loc[df['Codigo'] == codigos_alumno, 'NºClase3'].item()

    # Encontrar el elemento del cuadro de texto por su atributo id
    input_cuadro_texto = driver.find_element(By.ID, "ENRL_REQ_DETAIL_CLASS_NBR$1")

    # Ingresar el valor deseado en el campo de texto
    input_cuadro_texto.send_keys(texto_columna_4)

    # Agregar fila
    boton_add_fila = driver.find_element(By.ID, "ENRL_REQ_DETAIL$new$0$$0")
    boton_add_fila.click()

    time.sleep(2)

    ## Numero de clase 4
    # Obtener el texto de la columna "NºClase4" para la fila actual
    texto_columna_5 = df.loc[df['Codigo'] == codigos_alumno, 'NºClase4'].item()

    # Encontrar el elemento del cuadro de texto por su atributo id
    input_cuadro_texto = driver.find_element(By.ID, "ENRL_REQ_DETAIL_CLASS_NBR$1")

    # Ingresar el valor deseado en el campo de texto
    input_cuadro_texto.send_keys(texto_columna_5)

    time.sleep(3)"""

    # ###### Boton enviar######## En desarrollo------------

    # Localizar el botón por ID
    boton_enviar = driver.find_element(By.ID, "ENRL_LIST_WRK_ENRL_REQ_RC_PB")

    # Hacer clic en el botón
    boton_enviar.click()

    time.sleep(3)

    driver.find_element(By.ID, "#ICList").click()
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "#ICList"))
    ).click()

    time.sleep(3)

    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "ICTAB_1"))
    ).click()

    time.sleep(3)

######################################################################

# Cerrar el navegador
driver.quit()