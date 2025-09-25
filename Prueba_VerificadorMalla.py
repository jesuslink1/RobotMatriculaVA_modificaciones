from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import re
import time

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

# Lista para almacenar los resultados
resultados = []
nota_aprobatoria = 13

try:

    # Iterar sobre cada fila del DataFrame
    for index, row in df.iterrows():
        tiene_notas_aprobatorias = False
        esta_llevando_cursos = False
        row_index = 0
        periodo_llevando_cursos = "No aplica"
        robot_decision = "Usuario debe verificar"

        # Obtener el código de alumno
        codigo_alumno = str(row['Codigo'])

        input_codigo = driver.find_element(By.ID, "PEOPLE_SRCH_EMPLID")
        input_codigo.clear()
        input_codigo.send_keys(codigo_alumno)
        time.sleep(2)
        boton_buscar = driver.find_element(By.ID, "#ICSearch")
        boton_buscar.click()

        time.sleep(3)

        # Encontrar el campo de selección de "Más Académicos"
        dropdown = Select(driver.find_element(By.ID, "DERIVED_SSS_SCL_SSS_MORE_ACADEMICS"))
        dropdown.select_by_value("3010")  # Seleccionar "Condiciones Académicas"

        time.sleep(3)

        # Encontrar el botón de "Ir" por su nombre
        boton_ir = driver.find_element(By.NAME, "DERIVED_SSS_SCL_SSS_GO_1$IMG")
        boton_ir.click()

        time.sleep(3)  # Esperar a que cargue la página siguiente

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

        time.sleep(3)

        # Verificar si tiene dos programas activos
        programas_activos = []
        # Establecer un índice manualmente para controlar el bucle
        programa_index = 0  # Empezamos desde el primer programa

        #####################################################
        # VERIFICAR PROGRAMAS

        try:
            while True:
                programa_id = f"win0divDERIVED_SAA_DPR_GROUPBOX1GP${programa_index}"
                try:
                    programa_element = driver.find_element(By.ID, programa_id)
                    #print(programa_element)
                    # Extraer el nombre completo del programa de forma segura
                    programa_text = programa_element.text.strip()  # Captura todo el texto del programa
                    #print(programa_text)
                    if programa_text:
                        programas_activos.append(programa_text)  # Agrega el nombre completo a la lista
                    else:
                        programas_activos.append("Programa desconocido")  # Si está vacío, marcar como desconocido

                except NoSuchElementException:
                    break  # Romper el bucle si no se encuentran más programas activos

                # Incrementar el índice para buscar el siguiente programa
                programa_index += 1

        except NoSuchElementException:
            pass

        #print(programas_activos)

        # Indicar el número de programas activos dinámicamente
        if len(programas_activos) == 1:
            resultado_programas = "Un programa activo"
        else:
            resultado_programas = f"{len(programas_activos)} programas activos"

        time.sleep(1)

        #####################################################
        # OBTENER MALLAS

        # Localizar los divs que contienen las mallas para cada programa activo
        resultados_malla = []

        try:
            malla_index = 0  # Comienza con el primer programa
            while True:
                # Buscar la malla del programa actual
                malla_id = f"win0divDERIVED_SAA_DPR_GROUPBOX1GP${malla_index}"
                malla_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, malla_id))
                )
                malla_text = malla_element.text

                # Verificar si el texto contiene "CPEX"
                if "CPEX" in malla_text:
                    resultados_malla.append('CPEX')
                else:
                    malla_last_part = malla_text.split()[-1]

                    if re.match(r'^\d+$', malla_last_part):
                        malla_year = int(malla_last_part)
                        if malla_year != 2024:
                            resultados_malla.append('Malla antigua')
                        else:
                            resultados_malla.append('Malla 2024')
                    else:
                        resultados_malla.append('Malla antigua')

                # Incrementar el índice para buscar el siguiente programa
                malla_index += 1

        except TimeoutException:
            # Si no se encuentra más mallas, terminar el bucle
            if not resultados_malla:  # Si no encontró ninguna malla
                # print(f'Elemento malla no encontrado para el código {codigo_alumno}')
                resultado = 'Elemento malla no encontrado'
                notas_resultado = 'Ninguno'
                resultados.append((codigo_alumno, resultado, notas_resultado, resultado_programas))
                navegar_a_centro_servicios()
                continue
        # print(f"resultados_malla : {resultados_malla}")
        # Concatenar los resultados de las mallas en caso de que haya más de una

        if len(resultados_malla) > 1:
            # print("Resultado malla + de 1")
            resultado = ' y '.join(set(resultados_malla))  # Eliminar duplicados y concatenar
            # print(resultado)
        else:
            resultado = resultados_malla[0]

        #####################################################
        # "Programa a baja"

        programas_a_baja = []  # Almacenar los programas que deben darse de baja

        # Caso 1: Malla antigua y malla 2024
        if 'Malla antigua' in resultados_malla and 'Malla 2024' in resultados_malla:
            for i, malla in enumerate(resultados_malla):
                if malla == 'Malla antigua':
                    programas_a_baja.append(programas_activos[i])  # Agregar programas de malla antigua a la lista de baja

        # Caso 2: Todos los programas son de malla antigua
        elif 'Malla antigua' in resultados_malla and 'Malla 2024' not in resultados_malla:
            programas_a_baja.append("Verificar programa a baja")  # Indicar que se necesita verificar

        # Caso 3: Todos los programas son de malla 2024
        elif 'Malla 2024' in resultados_malla and 'Malla antigua' not in resultados_malla:
            # Dar de baja a todos los programas menos el más reciente
            programas_a_baja = programas_activos[:-1]  # Baja todos excepto el programa con el índice 0 (el más reciente)

        # Añadir la información de programas a baja a la columna del DataFrame
        programa_a_baja_resultado = ', '.join(programas_a_baja) if programas_a_baja else 'No aplica'

        time.sleep(1)

        #####################################################
        # VERIFICADOR DE NOTAS

        # Inicializar el índice de las tablas y filas
        row_index = 0
        tabla_index = 0
        ultimo_ciclo = 0
        ultimo_bloque = 0

        # Agregar una lista para rastrear los ciclos y bloques de notas aprobatorias
        cursos_aprobados = 0
        ultimo_ciclo_aprobatorio = 0
        ultimo_bloque_aprobatorio = 0

        # Variables para controlar si se encontró una nota aprobatoria o si está llevando cursos
        tiene_notas_aprobatorias = False
        esta_llevando_cursos = False
        ultimo_periodo_estudio = None  # Variable para el último período con nota aprobatoria
        todas_desaprobatorias = True  # Asumimos que todas las notas son desaprobatorias inicialmente
        tiene_notas_calificadas_desaprobatorias = False

        # Bucle para buscar todas las tablas dinámicamente
        while True:
            try:
                # Identificar el ciclo y bloque
                ciclo_bloque_id = f"win0divDERIVED_SAA_DPR_GROUPBOX3${tabla_index}"
                ciclo_bloque_element = driver.find_element(By.ID, ciclo_bloque_id)
                ciclo_bloque_text = ciclo_bloque_element.text.strip()

                # Actualización de la expresión regular para detectar tanto "C1 B1" como "C1B1"
                ciclo_bloque_match = re.search(r'C(\d+)\s?B(\d+)', ciclo_bloque_text)
                if ciclo_bloque_match:
                    ciclo = ciclo_bloque_match.group(1)
                    bloque = ciclo_bloque_match.group(2)
                    print(f"Ciclo: {ciclo} Bloque: {bloque}")

                # Ahora procesar las notas de este ciclo/bloque
                table_id = f"ACE_DERIVED_SAA_DPR_GROUPBOX4${tabla_index}"
                table = driver.find_element(By.ID, table_id)

                # Bucle para buscar filas dentro de la tabla
                while True:
                    try:
                        # Identificar la nota de la fila
                        grade_id = f"DERIVED_SAA_DPR_CRSE_GRADE_OFF${row_index}"
                        grade_element = table.find_element(By.ID, grade_id)
                        grade_text = grade_element.text.strip()

                        # Obtener el nombre del curso en la fila
                        course_descr_id = f"CRSE_DESCR${row_index}"
                        course_element = table.find_element(By.ID, course_descr_id)
                        course_name = course_element.text.strip()

                        # Obtener el período del curso en la fila
                        periodo_id = f"CRSE_WHEN${row_index}"
                        periodo_element = table.find_element(By.ID, periodo_id)
                        periodo_text = periodo_element.text.strip()

                        # Ignorar cursos de inglés ("ELEMENTARY")
                        if "ELEMENTARY" in course_name.upper():
                            print(f"Ignorando curso de inglés: {course_name}")
                            row_index += 1
                            continue  # Ignorar este curso

                        # Mostrar la información de la fila (curso, nota, y período)
                        print(f"  Curso: {course_name} - Nota: {grade_text} - Período: {periodo_text}")

                        # Verificar si la nota es un número
                        if grade_text and grade_text.isdigit():
                            grade = int(grade_text)

                            # Verificar si la nota es aprobatoria (13 o más)
                            if grade >= nota_aprobatoria:
                                tiene_notas_aprobatorias = True
                                todas_desaprobatorias = False  # Ya no es cierto que todas las notas sean desaprobatorias

                                # Actualizar el contador de cursos aprobados
                                cursos_aprobados += 1

                                # Actualizar el último ciclo y bloque de nota aprobatoria
                                ultimo_ciclo_aprobatorio = ciclo
                                ultimo_bloque_aprobatorio = bloque

                                # Solo actualizar el último período si `periodo_text` no está vacío
                                if periodo_text:
                                    ultimo_periodo_estudio = periodo_text
                                    ultimo_ciclo = ciclo
                                    ultimo_bloque = bloque
                                    print(f"Nota aprobatoria para el curso {course_name}: {grade}")
                                else:
                                    print(
                                        f"Periodo vacío encontrado para el curso {course_name}, no se actualiza ultimo_periodo_estudio")
                                    # Verificar si la nota es desaprobatoria (entre 1 y 12)

                            elif 1 <= grade < nota_aprobatoria:
                                print(f"Nota desaprobatoria para el curso {course_name}: {grade}")
                                tiene_notas_calificadas_desaprobatorias = True

                            else:
                                todas_desaprobatorias = False  # Si la nota no es desaprobatoria (ej. vacía o no válida)

                        # Verificar si hay imágenes que indican que está llevando cursos
                        image_container_id = f"win0divCRSE_STAT${row_index}"
                        image_container = table.find_element(By.ID, image_container_id)
                        image_element = image_container.find_elements(By.TAG_NAME, "img")

                        for img in image_element:
                            if img.get_attribute("alt") == "En Curso":
                                esta_llevando_cursos = True
                                periodo_llevando_cursos = periodo_text

                        # Incrementar el índice de fila para buscar la siguiente
                        row_index += 1

                    except NoSuchElementException:
                        # No hay más filas en esta tabla
                        break  # Salimos del bucle interno de filas

                # Incrementar el índice de la tabla para buscar la siguiente
                tabla_index += 1

            except NoSuchElementException:
                # No hay más tablas
                break  # Salimos del bucle externo de tablas

        # Al final del ciclo, imprimir el último período con nota aprobatoria y ciclo/bloque
        if ultimo_periodo_estudio:
            print(f"Último período en el que se estudió con nota aprobatoria: {ultimo_periodo_estudio}")
            print(
                f"Último ciclo y bloque en el que se estudió con nota aprobatoria: Ciclo {ultimo_ciclo}, Bloque {ultimo_bloque}")
        else:
            print("No se encontraron notas aprobatorias.")

        # Verificación si el ciclo y bloque son 0
        if ultimo_ciclo == 0 and ultimo_bloque == 0:
            ultimo_ciclo_bloque = "No aplica"
        else:
            ultimo_ciclo_bloque = f"Ciclo {ultimo_ciclo}, Bloque {ultimo_bloque}"

        # Verificación si el último período de estudio está vacío
        if not ultimo_periodo_estudio:
            ultimo_periodo_estudio = "No aplica"

        # Al final del ciclo, revisamos si se encontraron notas aprobatorias o cursos en curso
        if tiene_notas_aprobatorias and esta_llevando_cursos:
            notas_resultado = "Tiene notas aprobatorias y esta llevando cursos"
        elif tiene_notas_aprobatorias:
            notas_resultado = "Tiene notas aprobatorias"
        elif tiene_notas_calificadas_desaprobatorias and not tiene_notas_aprobatorias:
            notas_resultado = "Tiene solo notas calificadas desaprobatorias"
        elif esta_llevando_cursos:
            notas_resultado = "Llevando cursos"
        else:
            notas_resultado = "No tiene notas aprobatorias"

        time.sleep(1)

        ##########################
        # VERIFICACIONES ROBOT

        # Verificar si en resultados_malla figura "CPEX"
        if 'CPEX' in resultados_malla:
            robot_decision = "Programa CPEX"

        # Verificar si es "Venta Nueva"
        elif (resultado == "Malla 2024" and
                notas_resultado == "No tiene notas aprobatorias" and
                resultado_programas == "Un programa activo" and
                programa_a_baja_resultado == "No aplica" and
                ultimo_periodo_estudio == "No aplica" and
                ultimo_ciclo_bloque == "No aplica" and
                periodo_llevando_cursos == "No aplica"):
            robot_decision = "Venta Nueva"

        elif (resultado == "Malla 2024" and
                notas_resultado == "Tiene solo notas calificadas desaprobatorias" and
                resultado_programas == "Un programa activo" and
                programa_a_baja_resultado == "No aplica" and
                ultimo_periodo_estudio == "No aplica" and
                ultimo_ciclo_bloque == "No aplica" and
                periodo_llevando_cursos == "No aplica"):
            robot_decision = "Venta nueva, se debe eliminar notas desaprobatorias"

        # Verificar si es "Venta nueva, se debe actualizar malla"
        elif (resultado == "Malla antigua" and
              notas_resultado in ["No tiene notas aprobatorias", "Tiene solo notas calificadas desaprobatorias"] and
              resultado_programas == "Un programa activo"):
            robot_decision = "Venta nueva, se debe actualizar malla"

        # Verificar si es "Venta nueva, se debe dar baja a programa extra"
        elif (notas_resultado in ["No tiene notas aprobatorias", "Tiene solo notas calificadas desaprobatorias"] and
              len(programas_activos) >= 2):  # Verifica si tiene 2 o más programas activos
            # Verificar si las mallas son antiguas
            if all(malla == "Malla antigua" for malla in resultados_malla):
                robot_decision = "Venta nueva, se debe dar de baja a programa y actualizar malla"
            else:
                robot_decision = "Venta nueva, se debe dar baja a programa adicional"

        # Verificar si es "Venta nueva, se debe eliminar notas desaprobatorias"
        elif resultado == "Malla 2024" and todas_desaprobatorias:
            # Si también está matriculado en un periodo anterior, ajustar la decisión
            if esta_llevando_cursos:
                robot_decision = f"Venta nueva, matriculado en periodo anterior: {periodo_llevando_cursos}"
            else:
                robot_decision = "Venta nueva, se debe eliminar notas desaprobatorias"

        # Verificar si tiene notas aprobatorias
        if tiene_notas_aprobatorias:
            # Convertir ultimo_ciclo_aprobatorio a entero para la comparación
            if ultimo_ciclo_aprobatorio.isdigit():
                ciclo_aprobatorio_int = int(ultimo_ciclo_aprobatorio)
            else:
                ciclo_aprobatorio_int = 0  # Si no es un número válido, asignamos 0 por defecto

            # Si tiene 2 o más programas activos y ha completado el ciclo 6 o superior (egresado)
            if len(programas_activos) >= 2 and ciclo_aprobatorio_int >= 6:
                # Marcamos el programa más antiguo como "Egresado"
                programas_activos[0] += "(Egresado)"
                programa_a_baja_resultado = "Egresado"
                ultimo_ciclo_bloque = "Egresado"

                # Identificar la segunda carrera (la más reciente o de malla 2024)
                segunda_carrera = programas_activos[1] if len(programas_activos) > 1 else "Desconocida"

                # Si la segunda carrera no tiene notas aprobatorias (verificar si las notas de la segunda carrera son todas vacías o desaprobadas)
                if 'Malla 2024' in resultados_malla or segunda_carrera in programas_activos:
                    # Aquí se puede agregar una lógica más detallada para verificar las notas de la segunda carrera en particular
                    # Verificar si la segunda carrera tiene "No tiene notas aprobatorias"
                    print(notas_resultado)
                    print(resultado_programas)
                    if notas_resultado in ["No tiene notas aprobatorias", "Tiene solo notas calificadas desaprobatorias"] and resultado_programas == "2 programas activos":
                        robot_decision = "Venta nueva, tienen una carrera como egresado"
                    else:
                        robot_decision = f"Notas aprobatorias hasta: Ciclo {ultimo_ciclo_aprobatorio}, Bloque {ultimo_bloque_aprobatorio}"

            # Si no es egresado, verificar el número de cursos aprobados
            elif cursos_aprobados == 1:
                robot_decision = f"Notas aprobatorias hasta: Ciclo {ultimo_ciclo_aprobatorio}, Bloque {ultimo_bloque_aprobatorio}. Solo 1 curso aprobado"
            elif cursos_aprobados == 2:
                robot_decision = f"Notas aprobatorias hasta: Ciclo {ultimo_ciclo_aprobatorio}, Bloque {ultimo_bloque_aprobatorio}. Solo 2 cursos aprobados"
            else:
                robot_decision = f"Notas aprobatorias hasta: Ciclo {ultimo_ciclo_aprobatorio}, Bloque {ultimo_bloque_aprobatorio}"

        ######################################
        # DATAFRAME

        # Agregar resultados al DataFrame
        resultados.append((codigo_alumno, resultado, notas_resultado, resultado_programas, ', '.join(programas_activos),
                           programa_a_baja_resultado, ultimo_periodo_estudio, ultimo_ciclo_bloque, periodo_llevando_cursos,
                           robot_decision))

        print(f'{codigo_alumno} - {resultado}, {notas_resultado}, {resultado_programas}, {programas_activos}, Programa a baja: {programa_a_baja_resultado}, Ultimo periodo: {ultimo_periodo_estudio}, Ultimo ciclo y bloque: Ciclo {ultimo_ciclo}, Bloque {ultimo_bloque}, Periodo llevando cursos: {periodo_llevando_cursos}, Robot: {robot_decision}')

        cancel_button = driver.find_element(By.ID, "DERIVED_SCC_XFR_CANCEL_BTN")
        cancel_button.click()

        time.sleep(3)

        return_button = driver.find_element(By.ID, "DERIVED_SSS_SCL_RETURN_SEARCH$282$")
        return_button.click()

        time.sleep(3)

        boton_borrar = driver.find_element(By.ID, "#ICClear")
        boton_borrar.click()

        time.sleep(3)

except Exception as e:
    print(f"Se ha producido un error: {e}")

finally:
    # Convertir la lista de resultados en un DataFrame
    resultados_df = pd.DataFrame(resultados, columns=['Codigo', 'Malla', 'Notas', 'Programas Activos', 'Carrera actual',
                                                      'Programa a baja', 'Ultimo periodo que estudio',
                                                      'Ultimo ciclo y bloque', 'Periodo llevando cursos', 'Robot'])

    # Intentar guardar el archivo Excel hasta 3 veces si ocurre un error
    intentos = 0
    exito = False

    while intentos < 3 and not exito:
        try:
            # Guardar el archivo como Excel
            resultados_df.to_excel('resultados_mallas.xlsx', index=False, engine='openpyxl')
            print("Archivo 'resultados_mallas.xlsx' creado con éxito.")

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

