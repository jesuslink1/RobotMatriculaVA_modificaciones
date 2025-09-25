from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import re

def navegar_a_centro_servicios():
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
    # INGRESAR A Centro Servicios al Alumnado

    # Esperar hasta que se cargue el enlace "Mis Favoritos"
    enlace_favoritos = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'a#fldra_MYFAVORITES')))
    enlace_favoritos.click()

    # Esperar hasta que se cargue el enlace "Inscripción Rápida"
    enlace_inscripcion = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Centro Servicios al Alumnado")))
    enlace_inscripcion.click()

    #####################################################
    # VERIFICACIÓN

    # Cambiar el contexto al iframe "ptifrmtgtframe"
    driver.switch_to.frame("ptifrmtgtframe")

    # Encontrar el botón "Borrar" por su ID
    boton_borrar = driver.find_element(By.ID, "#ICClear")
    boton_borrar.click()

    time.sleep(3)


# Leer los códigos de alumno desde un archivo Excel
df = pd.read_excel('C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\CodigosVerificarMalla.xlsx', sheet_name='Hoja1')
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
# INGRESAR A Centro Servicios al Alumnado

# Esperar hasta que se cargue el enlace "Mis Favoritos"
enlace_favoritos = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a#fldra_MYFAVORITES')))
enlace_favoritos.click()

# Esperar hasta que se cargue el enlace "Inscripción Rápida"
enlace_inscripcion = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, "Centro Servicios al Alumnado")))
enlace_inscripcion.click()

#####################################################
# VERIFICACIÓN

# Cambiar el contexto al iframe "ptifrmtgtframe"
driver.switch_to.frame("ptifrmtgtframe")

# Encontrar el botón "Borrar" por su ID
boton_borrar = driver.find_element(By.ID, "#ICClear")
boton_borrar.click()

time.sleep(2)

# Lista para almacenar los resultados
resultados = []
# Cantidad de tablas a verificar
numero_de_tablas = 2
nota_aprobatoria = 13

# Iterar sobre cada fila del DataFrame
for index, row in df.iterrows():
    tiene_notas_aprobatorias = False
    esta_llevando_cursos = False
    row_index = 0
    # Obtener el código de alumno y el programa académico de la fila actual
    codigo_alumno = str(row['Codigo'])  # Convertir a cadena

    # Encontrar el campo de entrada por su ID y limpiar cualquier texto previo
    input_codigo = driver.find_element(By.ID, "PEOPLE_SRCH_EMPLID")
    input_codigo.clear()

    # Ingresar el código del alumno en el campo de entrada
    input_codigo.send_keys(codigo_alumno)

    time.sleep(2)

    # Encontrar el botón "Buscar" por su ID y hacer clic en él
    boton_buscar = driver.find_element(By.ID, "#ICSearch")
    boton_buscar.click()

    time.sleep(3)

    # Encontrar el campo de selección de "Más Académicos"
    dropdown = Select(driver.find_element(By.ID, "DERIVED_SSS_SCL_SSS_MORE_ACADEMICS"))
    dropdown.select_by_value("3010")  # Seleccionar "Condiciones Académicas"

    time.sleep(1)

    # Encontrar el botón de "Ir" por su nombre
    boton_ir = driver.find_element(By.NAME, "DERIVED_SSS_SCL_SSS_GO_1$IMG")
    boton_ir.click()

    time.sleep(2)  # Esperar a que cargue la página siguiente

    try:
        # Verificar si el combobox está presente
        combobox = driver.find_element(By.ID, "DERIVED_SAA_DPR_SSR_STDNT_INSTCAR$63$")

        # Seleccionar la opción "SISE > Profesional SISE"
        select = Select(combobox)
        select.select_by_value("SISE PRGS")

        # Hacer clic en el enlace "Cambiar"
        change_link = driver.find_element(By.ID, "DERIVED_SAA_DPR_SSS_PB_CHANGE")
        change_link.click()

    except:
        pass

    time.sleep(4)

    # Localizar el div que contiene la malla
    try:
        malla_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "win0divDERIVED_SAA_DPR_GROUPBOX1GP$0"))
        )
        malla_text = malla_element.text
        # Extraer el último elemento del texto de la malla
        malla_last_part = malla_text.split()[-1]  # Obtener la última palabra

        # Verificar si el último elemento es un número
        if re.match(r'^\d+$', malla_last_part):  # Verificar si es un número
            malla_year = int(malla_last_part)
            if malla_year != 2024:
                resultado = 'Malla antigua'
            else:
                resultado = 'Malla 2024'
        else:
            resultado = 'Malla antigua'
    except TimeoutException:
        print(f'Elemento malla no encontrado para el código {codigo_alumno}')
        resultado = 'Elemento malla no encontrado'
        notas_resultado = 'Ninguno'
        # Agregar el resultado a la lista con la nueva columna
        resultados.append((codigo_alumno, resultado, notas_resultado))
        navegar_a_centro_servicios()
        continue

    # Bandera para verificar si hubo un error en el bucle de tablas
    error_en_tablas = False

    # Iterar a través de las tablas (desde ACE_DERIVED_SAA_DPR_GROUPBOX4$0 hasta ACE_DERIVED_SAA_DPR_GROUPBOX4$11)
    for i in range(numero_de_tablas):
        try:
            table_id = f"ACE_DERIVED_SAA_DPR_GROUPBOX4${i}"
            table = driver.find_element(By.ID, table_id)

            # Dentro de cada tabla, iterar a través de las posibles filas de cursos
            while True:
                try:
                    grade_id = f"DERIVED_SAA_DPR_CRSE_GRADE_OFF${row_index}"
                    grade_element = table.find_element(By.ID, grade_id)
                    grade_text = grade_element.text.strip()

                    if grade_text:
                        if grade_text.isdigit():
                            grade = int(grade_text)
                            if grade >= 13:
                                tiene_notas_aprobatorias = True
                                break

                    # Verificar si hay una imagen en la celda correspondiente al row_index
                    image_container_id = f"win0divCRSE_STAT${row_index}"
                    image_container = table.find_element(By.ID, image_container_id)
                    image_element = image_container.find_elements(By.TAG_NAME, "img")

                    # Verificar el atributo alt de la imagen
                    for img in image_element:
                        if img.get_attribute("alt") == "En Curso":
                            esta_llevando_cursos = True

                    row_index += 1
                except NoSuchElementException:
                    break
                except Exception as e:
                    print(f"Error inesperado: {str(e)}")
                    break

            if tiene_notas_aprobatorias or esta_llevando_cursos:
                break

        except TimeoutException:
            notas_resultado = 'Error al conseguir Notas'
            print(f'{codigo_alumno} - {notas_resultado}')
            # Agregar el resultado a la lista con la nueva columna
            resultados.append((codigo_alumno, resultado, notas_resultado))
            navegar_a_centro_servicios()
            error_en_tablas = True
            break

        except Exception as e:
            print(f"Error inesperado al buscar tabla {i}: {str(e)}")
            notas_resultado = 'Error al conseguir Notas'
            print(f'{codigo_alumno} - {notas_resultado}')
            # Agregar el resultado a la lista con la nueva columna
            resultados.append((codigo_alumno, resultado, notas_resultado))
            navegar_a_centro_servicios()
            error_en_tablas = True
            break

    if error_en_tablas:
        continue

    # Determinar si está llevando cursos o tiene notas aprobatorias
    if esta_llevando_cursos:
        notas_resultado = "Llevando cursos"
    elif tiene_notas_aprobatorias:
        notas_resultado = "Tiene notas aprobatorias"
    else:
        notas_resultado = "No tiene notas aprobatorias"

    # Agregar el resultado a la lista con la nueva columna
    resultados.append((codigo_alumno, resultado, notas_resultado))

    print(f'{codigo_alumno} - {resultado} y {notas_resultado}')

    # Hacer clic en el botón "Cancelar"
    cancel_button = driver.find_element(By.ID, "DERIVED_SCC_XFR_CANCEL_BTN")
    cancel_button.click()

    time.sleep(3)

    # Hacer clic en el botón "Volver a Búsqueda"
    return_button = driver.find_element(By.ID, "DERIVED_SSS_SCL_RETURN_SEARCH$282$")
    return_button.click()

    time.sleep(3)

    # Encontrar el botón "Borrar" por su ID
    boton_borrar = driver.find_element(By.ID, "#ICClear")
    boton_borrar.click()

    time.sleep(4)

# Convertir la lista de resultados en un DataFrame
resultados_df = pd.DataFrame(resultados, columns=['Codigo', 'Resultado', 'Notas'])

# Intentar guardar el DataFrame en un archivo CSV con reintentos
for attempt in range(3):
    try:
        resultados_df.to_csv('resultados_mallas.csv', index=False)
        print("Archivo CSV guardado correctamente.")
        break
    except Exception as e:
        print(f"Error al guardar el archivo CSV: {str(e)}")
        if attempt < 2:  # Si no es el último intento
            retry = input(f"Desea intentarlo nuevamente? (Sí/N): ").strip().lower()
            if retry == 'sí' or retry == 'si' or retry == 's':
                print(f"Reintentando... (Intento {attempt + 2}/3)")
                continue
            else:
                print("No se intentará guardar el archivo CSV nuevamente.")
                break
        else:
            print("No se pudo guardar el archivo CSV después de 3 intentos.")

# Cerrar el navegador
driver.quit()
