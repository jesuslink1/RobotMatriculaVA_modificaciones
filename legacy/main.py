from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
import time
import pandas as pd
import csv
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from VN import matricular_alumno_vn_si
from VA import matricular_alumno_vn_no

def click_home_button():
    try:
        # Intenta ejecutar un script de JavaScript para simular el clic en el botón
        driver.execute_script("document.querySelector('i.fa.fa-home').click();")
    except Exception as e:
        print(f"No se pudo simular el clic en el botón 'Volver a la página principal': {str(e)}")
        print(f"Código: {codigo}")


try:
    # Configura el objeto Service para el controlador de Chrome
    chrome_service = webdriver.chrome.service.Service('C:\\Users\\jramosc\\Downloads\\chromedriver-win64\\Chromedriver.exe')

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
    df_hoja1 = pd.read_excel('C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\ingresarCodigos.xlsx', sheet_name='Hoja1')

    # Obtener los códigos de la primera columna
    codigos = df_hoja1['Codigo'].tolist()

    # Lista para almacenar los registros de errores
    errores = []

    # Iterar sobre los códigos
    for codigo in codigos:

        matricula_completa = True  # Variable de bandera para rastrear si la matrícula se completa
        error_msg = None  # Variable para almacenar el mensaje de error final
        matriculado_algun_curso = False  # Variable para rastrear si se matriculó en al menos un curso

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
            time.sleep(2)

            # Buscar y hacer clic en el botón "Cargar"
            boton_cargar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Carga del Alumno')]"))
            )
            boton_cargar.click()

            time.sleep(2)

            # Buscar y hacer clic en el botón "Continuar"
            boton_continuar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Continuar')]"))
            )
            boton_continuar.click()

            time.sleep(2)

            # Buscar y hacer clic en el botón "Seleccionar"
            boton_seleccionar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Seleccionar')]"))
            )
            boton_seleccionar.click()

            time.sleep(4)

            matriculado_algun_curso = False

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
                        errores.append((codigo, "El alumno no tiene ciclo activo"))
                        continue # Saltar al siguiente código de alumno
                        time.sleep(4)

                except NoSuchElementException:
                    # No se encontró el mensaje de alerta, es posible que no aparezca
                    pass

            time.sleep(4)

            ################################################################################################
            # DENTRO DE LA MATRÍCULA: ESCOGER CICLOS Y CURSOS

            # Obtener el Programa, Campus y Turno correspondientes al código de alumno actual
            codigo_row = df_hoja1[df_hoja1['Codigo'] == codigo].iloc[0]
            programa = codigo_row['Programa']
            campus = codigo_row['Campus']
            turno = codigo_row['Turno']
            vn = codigo_row['Venta Nueva']
            modalidad_alumno = codigo_row['Modalidad']

            # Inicializamos las variables para el siguiente ciclo y bloque de matriculación
            siguiente_ciclo_matricula = None
            siguiente_bloque_matricula = None

            if vn == 'SI':
                print(f"Código de Alumno {codigo} - Venta nueva")
                # Llama a Def en hoja VN
                matricular_alumno_vn_si(codigo, campus, turno, vn, programa, modalidad_alumno, click_home_button)

            else:
                print(f"Código de Alumno {codigo} - Venta antigua")
                #Llama a Def en hoja VA
                matricular_alumno_vn_no(codigo, campus, turno, vn, programa, modalidad_alumno, driver, errores,
                                        click_home_button)

        except TimeoutException:
            # Si se produce un TimeoutException, determina si no se encontraron cursos o si la matrícula es incompleta
            print(f"TimeoutException para el código {codigo}: Curso '{nombre_curso}' (Nº Clase: {numero_clase})")

            if not matriculado_algun_curso:
                error_msg = f"No se encontraron cursos para matricular Seccion: {secciones_disponibles[0]} Programa {programa}"
                print(f"Código: {codigo}, Error: {error_msg} Seccion: {secciones_disponibles[0]} Programa {programa}")
            else:
                error_msg = f"MATRICULA INCOMPLETA - Sección: {secciones_disponibles[0]}, Número de Clase: {numero_clase}, Nombre del Curso: {nombre_curso}, Sede: {campus}, Turno: {turno}, Programa: {programa}"
                # Imprimir información adicional si hay una MATRÍCULA INCOMPLETA
                print(
                    f" MATRICULA INCOMPLETA = Código: {codigo}, Sección: {secciones_disponibles}, Número de Clase: {numero_clase}, Nombre del Curso: {nombre_curso}, Sede: {campus}, Turno: {turno}, Programa: {programa}")
            errores.append((codigo, error_msg))
            click_home_button()
            time.sleep(8)
            #continue  # Salta al siguiente código de alumno


        except Exception as e:
            error_msg = str(e)
            if "Message:" in error_msg and "Stacktrace:" in error_msg:
                error_msg = f"Código: {codigo}, REVISAR hubo un error en el proceso"
                errores.append((codigo, error_msg))
                click_home_button()
                time.sleep(4)
            else:
                matricula_completa = False
                click_home_button()
                time.sleep(4)

        print(f"Código: {codigo}, Error: {error_msg}")

            # Verificar si la matrícula se completó para todos los cursos
        #if not matricula_completa:
        #   if error_msg is None:
        #      error_msg = "MATRICULA INCOMPLETA"
        #if error_msg is not None:
        #    errores.append((codigo, error_msg))

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
            with open('reporte.csv', 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(['Codigo', 'Estado'])
                for codigo, error_msg in errores:
                    writer.writerow([codigo, error_msg])

        if errores:
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
