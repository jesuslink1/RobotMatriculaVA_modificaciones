from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import re
import time


# Leer los códigos de alumno desde un archivo Excel
df = pd.read_excel('C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\En Construcion\\CodigosVerificarCurso.xlsx', sheet_name='Hoja1')
codigos_alumno = df['Codigo'].tolist()

chrome_service = webdriver.chrome.service.Service('C:\\Users\\jramosc\\Downloads\\chromedriver-win64\\Chromedriver.exe')
chrome_service.start()

chrome_options = webdriver.ChromeOptions()
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

time.sleep(3)

# Inicializar un DataFrame vacío para asegurarnos de que 'resultados_df' siempre exista
resultados_df = pd.DataFrame()

# Solicitar al usuario que ingrese el número de Ciclo Lvo.
curso_buscar = input("Ingrese el nombre del curso a verificar nota: ")

nota_aprobatoria = 13

# Lista para almacenar los resultados
resultados = []

# Iterar sobre cada código de alumno
try:
    for codigo_alumno in codigos_alumno:
        input_codigo = driver.find_element(By.ID, "PEOPLE_SRCH_EMPLID")
        input_codigo.clear()
        input_codigo.send_keys(codigo_alumno)
        time.sleep(2)
        boton_buscar = driver.find_element(By.ID, "#ICSearch")
        boton_buscar.click()

        time.sleep(3)

        # Selecciona "Historial de Cursos" en el dropdown
        dropdown = Select(driver.find_element(By.ID, "DERIVED_SSS_SCL_SSS_MORE_ACADEMICS"))
        dropdown.select_by_value("2050")  # Valor para Historial de Cursos
        time.sleep(3)

        # Haz clic en el botón "Ir"
        boton_ir = driver.find_element(By.NAME, "DERIVED_SSS_SCL_SSS_GO_1$IMG")
        boton_ir.click()
        time.sleep(3)

        # Busca todas las filas de la tabla
        filas_tabla = driver.find_elements(By.CSS_SELECTOR, '#CRSE_HIST\\$scroll\\$0 tr[id^="trCRSE_HIST$0_row"]')

        time.sleep(2)

        # Inicializa las variables de verificación
        curso_encontrado = False
        nota_obtenida = None
        curso_aprobado = False  # Para saber si se aprobó alguna vez

        # Itera sobre las filas para buscar el curso
        for fila in filas_tabla:
            try:
                nombre_curso = fila.find_element(By.CSS_SELECTOR, 'td:nth-child(2) span a').text.strip()
                if nombre_curso == curso_buscar:
                    curso_encontrado = True  # El curso fue encontrado
                    # Extrae la nota obtenida del curso
                    nota_obtenida = fila.find_element(By.CSS_SELECTOR, 'td:nth-child(4) span').text.strip()
                    # Convierte la nota a un número para compararla
                    nota_obtenida = int(nota_obtenida)

                    # Verifica si la nota es aprobatoria (por ejemplo, 13 o más)
                    if nota_obtenida >= 13:
                        curso_aprobado = True  # Si la nota es aprobatoria, marca el curso como aprobado
                        break  # Se hace break si encuentra el curso aprobado
            except Exception as e:
                print(f"Error al procesar fila: {e}")

        # Maneja los dos casos
        if not curso_encontrado:
            print(f"{codigo_alumno} - El curso '{curso_buscar}' no fue encontrado.")
        elif not curso_aprobado:
            print(f"{codigo_alumno} - El curso '{curso_buscar}' fue desaprobado en todas las ocasiones.")

        # Almacena el resultado
        if curso_encontrado:
            aprobado = "SI" if int(nota_obtenida) >= nota_aprobatoria else "No"
            print(f"{codigo_alumno} Curso encontrado - {nota_obtenida} - {aprobado}")
            resultados.append({
                "Codigo": codigo_alumno,
                "Curso": curso_buscar,
                "Nota": nota_obtenida,
                "Aprobado": aprobado
            })
        else:
            resultados.append({
                "Codigo": codigo_alumno,
                "Curso": curso_buscar,
                "Nota": "No llevado",
                "Aprobado": "NO"
            })

        time.sleep(2)
        # Hacer clic en el botón por su atributo "id"
        button = driver.find_element(By.ID, "DERIVED_SCC_XFR_CANCEL_BTN")
        button.click()

        time.sleep(2)

        # Esperar hasta que el botón esté disponible y hacer clic
        button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "DERIVED_SSS_SCL_RETURN_SEARCH$282$"))
        )
        button.click()

        time.sleep(2)

        # Hacer clic en el botón por su atributo "id"
        button = driver.find_element(By.ID, "#ICClear")
        button.click()

        time.sleep(2)

except Exception as e:
    print(f"Se ha producido un error")

finally:
    # Crea un DataFrame con los resultados y guárdalo en Excel
    resultados_df = pd.DataFrame(resultados)

    # Intentar guardar el archivo Excel hasta 3 veces si ocurre un error
    intentos = 0
    exito = False

    while intentos < 3 and not exito:
        try:
            # Guardar el archivo como Excel
            resultados_df.to_excel("resultadosVerificarNotaCurso.xlsx", index=False)
            print("Resultados guardados en 'resultadosVerificarNotaCurso.xlsx'.")

            exito = True  # Indicar que la operación fue exitosa

        except Exception as e:
            intentos += 1
            print(f"Ocurrió un error al guardar el archivo Excel: Intento {intentos} de 3.")

            if intentos < 3:  # Preguntar si quiere volver a intentar si no ha alcanzado los 3 intentos
                respuesta = input("¿Deseas intentar guardar el archivo nuevamente? (si/no): ").lower()
                if respuesta != 'si':
                    print("No se volverá a intentar crear el archivo Excel.")
                    break  # Salir del bucle si el usuario no quiere volver a intentar
            else:
                print("Se alcanzó el número máximo de intentos. No se pudo crear el archivo Excel.")

    # Cerrar el navegador
    driver.quit()


