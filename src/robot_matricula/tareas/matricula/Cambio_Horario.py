from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
import time
import pandas as pd
import csv
from selenium.common.exceptions import TimeoutException, NoSuchElementException, NoSuchWindowException


def click_home_button():

    try:
        # Intenta ejecutar un script de JavaScript para simular el clic en el botón
        driver.execute_script("document.querySelector('i.fa.fa-home').click();")
    except Exception as e:
        print(f"No se pudo simular el clic en el botón 'Volver a la página principal': {str(e)}")
        print(f"Código: {codigo}")


try:
    ###############################
    # INGRESO A PORTAL

    # Configura el objeto Service para el controlador de Chrome
    chrome_service = webdriver.chrome.service.Service(
        'C:\\Users\\jramosc\\Downloads\\chromedriver-win64\\Chromedriver.exe')

    # Inicia el controlador de Chrome
    chrome_service.start()

    # Configura las opciones del navegador
    chrome_options = webdriver.ChromeOptions()

    # Crea una instancia del controlador de Chrome
    driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

    # Abre la página web
    driver.get('https://matricula.sise.edu.pe/admin')

    # Localiza el campo de texto por su atributo "placeholder"
    campo_texto = driver.find_element(By.CSS_SELECTOR, "input[placeholder='Usuario']")

    # Ingresa el texto "JRAMOSC" en el campo de texto
    campo_texto.send_keys("JRAMOSC")

    # Localiza el campo de texto para la contraseña por su atributo "placeholder"
    campo_contraseña = driver.find_element(By.CSS_SELECTOR, "input[placeholder='Contraseña']")
    campo_contraseña.send_keys("jOLgEt6n")

    boton_ingresar = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-black"))
    )
    boton_ingresar.click()

    time.sleep(3)

    ################################################################################################
    # INGRESO DE CÓDIGOS:
    # Cargar los datos del archivo Excel
    df_hoja1 = pd.read_excel(
        'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\En Construcion\\CodigosCambioHorario.xlsx',
        sheet_name='Hoja1')

    # Obtener los códigos de la primera columna
    codigos = df_hoja1['Codigo'].tolist()

    # Lista para almacenar los registros de errores
    errores = []

    ######################################################
    # VERIFICACION DE COINCIDENCIAS

    # Cargar las hojas del archivo Excel
    hoja2 = pd.read_excel(
        'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\En Construcion\\CodigosCambioHorario.xlsx',
        sheet_name='Hoja2'
    )
    hoja3 = pd.read_excel(
        'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\En Construcion\\CodigosCambioHorario.xlsx',
        sheet_name='Hoja3'
    )

    # Crear un diccionario con los datos de Hoja3
    dict_cursos = dict(zip(hoja3['Nº Clase'], hoja3['Nombre Curso']))

    # Crear listas para registrar inconsistencias
    inconsistencias_ncambiar = []
    inconsistencias_nnuevo = []

    # Crear conjuntos para verificar coincidencias entre nombres
    cambiar_nombres = set(dict_cursos[n] for n in hoja2['Ncambiar'] if n in dict_cursos)
    nuevo_nombres = set(dict_cursos[n] for n in hoja2['Nnuevo'] if n in dict_cursos)

    # Identificar inconsistencias específicas
    for n in hoja2['Ncambiar']:
        if n in dict_cursos and dict_cursos[n] not in nuevo_nombres:
            inconsistencias_ncambiar.append((n, dict_cursos[n]))

    for n in hoja2['Nnuevo']:
        if n in dict_cursos and dict_cursos[n] not in cambiar_nombres:
            inconsistencias_nnuevo.append((n, dict_cursos[n]))

    # Verificar si los conjuntos coinciden
    if not inconsistencias_ncambiar and not inconsistencias_nnuevo:
        print("Todos los números de clase coinciden con los nombres de curso. Procediendo...")
    else:
        print("No todos los nombres de curso coinciden entre Ncambiar y Nnuevo.")

        if inconsistencias_ncambiar:
            print("Inconsistencias en Ncambiar:")
            for n, curso in inconsistencias_ncambiar:
                print(f"- Nº Clase: {n}, Nombre Curso: {curso}")

        if inconsistencias_nnuevo:
            print("Inconsistencias en Nnuevo:")
            for n, curso in inconsistencias_nnuevo:
                print(f"- Nº Clase: {n}, Nombre Curso: {curso}")

        respuesta = input("¿Desea continuar de todos modos? (Si/No): ")
        if respuesta.lower() != 'si':
            print("Proceso terminado debido a inconsistencias.")
            exit()

    # Iterar sobre los códigos
    for codigo in codigos:

        try:
            # Localiza el elemento del cuadro de texto
            cuadro_codigo = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'code'))
            )

            # Ingresa el código en el cuadro de texto
            cuadro_codigo.clear()
            cuadro_codigo.send_keys(str(codigo))

            # Esperar un tiempo antes de hacer clic en el botón "Buscar"
            time.sleep(2)

            # Buscar y hacer clic en el botón "Buscar"
            boton_buscar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH,
                                            "//button[contains(@class, 'btn') and contains(@class, 'rounded-1') and contains(@class, 'btn-sise') and contains(@class, 'ml-2') and text()='Buscar']"))
            )
            boton_buscar.click()

            # Esperar un tiempo adicional después de hacer clic en el botón "Buscar"
            time.sleep(3)

            # Buscar y hacer clic en el botón "Cargar"
            boton_cargar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Carga del Alumno')]"))
            )
            boton_cargar.click()

            time.sleep(5)

            # Buscar y hacer clic en el botón "Continuar"
            boton_continuar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Continuar')]"))
            )
            boton_continuar.click()

            time.sleep(6)

            # Buscar y hacer clic en el botón "Seleccionar"
            boton_seleccionar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Seleccionar')]"))
            )
            boton_seleccionar.click()

            time.sleep(8)

            # Declarar la variable overlay_container
            overlay_container = None

            # Buscar el elemento <div class="overlay-container">
            try:
                 overlay_container = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "overlay-container"))
                )
            except NoSuchElementException:
                # No se encontró el elemento, puedes manejarlo de acuerdo a tus necesidades
                pass

            # Comprobar si el mensaje de alerta está presente en el elemento
            if overlay_container:
                try:
                    alert_message = overlay_container.find_element(By.CLASS_NAME, "toast-message")
                    mensaje_alerta = alert_message.text
                    # Verificar si el mensaje contiene "El alumno no tiene ciclo activo"
                    if "El alumno no tiene ciclo activo" in mensaje_alerta:
                        print(f"{codigo} {mensaje_alerta}")
                        print("------------------------------------------------------------------------")
                        print(" ")
                        errores.append((codigo, "El alumno no tiene ciclo activo"))
                        time.sleep(2)
                        continue  # Saltar al siguiente código de alumno

                except NoSuchElementException:
                    # No se encontró el mensaje de alerta, es posible que no aparezca
                    pass

            # Buscar y hacer clic en el botón "Unidades Didácticas Matriculadas"
            boton_ud_matriculadas = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Unidades Didácticas Matriculadas')]"))
            )
            boton_ud_matriculadas.click()

            time.sleep(2)

            try:
                # Verificar si la tabla contiene "Cursos a Cargo"
                cursos_a_cargo = driver.find_elements(By.XPATH,
                                                      "//tr[contains(@class, 'inCharge_bg') and contains(., 'Cursos a Cargo')]")
                if cursos_a_cargo:
                    print(f"{codigo} - Tiene cursos a cargo matriculados. No se realizará la desmatricula.")
                    errores.append((codigo, "Tiene cursos a cargo matriculados"))
                    click_home_button()
                    time.sleep(2)
                else:
                    time.sleep(2)

                    # Verificar si hay tachos al inicio
                    tachos = driver.find_elements(By.XPATH, "//i[contains(@class, 'fa-trash')]")
                    if not tachos:
                        # No hay cursos matriculados desde el inicio
                        print(f"{codigo} - No se encuentra matriculado en ningún curso")
                        errores.append((codigo, "No se encuentra matriculado en ningún curso"))
                        click_home_button()

                    else:
                        # Hay tachos disponibles, iniciar desmatriculación

                        while True:
                            # Buscar todos los íconos de tachos (fa-trash)
                            tachos = driver.find_elements(By.XPATH, "//i[contains(@class, 'fa-trash')]")

                            if not tachos:
                                print(f"{codigo} - Desmatrícula completa")
                                errores.append((codigo, "Desmatricula completa"))
                                click_home_button()
                                break  # Salir del bucle si no hay tachos

                            try:
                                # Hacer clic en el primer tacho de la lista
                                tacho = tachos[0]
                                time.sleep(2)
                                tacho.click()  # Hacer clic en el tacho para desmatricular
                                time.sleep(2)  # Breve pausa para evitar conflictos de sincronización

                                # Esperar a que aparezca el botón "Confirmar"
                                boton_confirmar = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Confirmar')]"))
                                )
                                boton_confirmar.click()  # Hacer clic en "Confirmar"
                                time.sleep(3)  # Esperar un momento para que el proceso se complete

                            except TimeoutException:
                                print("No apareció el botón de confirmación después de hacer clic en un tacho.")
                                error_msg = "No apareció el botón de confirmación después de hacer clic en un tacho."
                                errores.append((codigo, error_msg))
                                click_home_button()
                                time.sleep(4)
                            except Exception as e:
                                print(f"No se pudo procesar el tacho: {e}")
            except Exception as e:
                print(f"Error al procesar la tabla de cursos: {e}")

        except TimeoutException:
            # Si se produce un TimeoutException, determina si no se encontraron cursos o si la matrícula es incompleta
            print(f"{codigo} - TimeoutException para el código")
            error_msg = "TimeoutException para el código"
            errores.append((codigo, error_msg))
            click_home_button()


except Exception as e:
    # Si se produce un error en cualquier parte del proceso, agrega una tupla con el código erróneo y la descripción del error a la lista de errores
    errores.append((codigo, f"Error en el proceso: {str(e)}"))
    click_home_button()  # Aquí se hace clic al botón en caso de error
    time.sleep(4)

finally:
    # Cierra la ventana actual del navegador
    if 'driver' in locals():
        driver.close()

    # Detiene el servicio
    if 'chrome_service' in locals():
        chrome_service.stop()

# Escribir los registros de errores en un archivo CSV
intentos = 3  # Número máximo de intentos permitidos

while intentos > 0:
    try:
        # Escribir los registros de errores en un archivo CSV
        if errores:
            with open('../../../../data/matricula/salida/reporteDesmatriculados.csv', 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(['Codigo', 'Estado'])
                for codigo, error_msg in errores:
                    writer.writerow([codigo, error_msg])

        if errores:
            print("========================================================================")
            print("Se generó un reporte.")
            print("Códigos de alumnos atendidos:")
            for codigo, error_msg in errores:
                print(f"Código: {codigo}, {error_msg}")

        # Si todo fue exitoso, salir del bucle while
        break
    except Exception as e:
        print(f"Error al generar el reporte: {e}")
        intentos -= 1  # Reducir el número de intentos disponibles

        if intentos > 0:
            respuesta = input("¿Desea volver a intentar? (Si/No): ")
            if respuesta.lower() != 'si':
                break
        else:
            print("No se encontraron errores en el proceso.")
