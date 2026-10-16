from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
import time
import pandas as pd
import csv
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException

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
    df_hoja1 = pd.read_excel(
        'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\ingresarCodigos.xlsx',
        sheet_name='Hoja1')

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
            time.sleep(3)

            # Buscar y hacer clic en el botón "Cargar"
            boton_cargar = WebDriverWait(driver, 17).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Carga del Alumno')]"))
            )
            boton_cargar.click()

            time.sleep(5)

            # Buscar y hacer clic en el botón "Continuar"
            boton_continuar = WebDriverWait(driver, 17).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Continuar')]"))
            )
            boton_continuar.click()

            time.sleep(6)

            # Buscar y hacer clic en el botón "Seleccionar"
            boton_seleccionar = WebDriverWait(driver, 17).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Seleccionar')]"))
            )
            boton_seleccionar.click()

            time.sleep(6)

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
                        print("------------------------------------------------------------------------")
                        print(" ")
                        errores.append((codigo, "El alumno no tiene ciclo activo"))
                        continue  # Saltar al siguiente código de alumno
                        time.sleep(4)

                except NoSuchElementException:
                    # No se encontró el mensaje de alerta, es posible que no aparezca
                    pass

            time.sleep(6)

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
            filtro = None

            if vn == 'SI':
                print(f"Código de Alumno {codigo} - Venta nueva")

                # Agregar condiciones para ajustar el programa si es uno de los programas especificados
                # if programa in ["ADMNI", "MARKT"]:
                #    programa = "ADMNE"
                # elif programa in ["ADMBF", "CONTA"]:
                #    programa = "ADMBF"
                if programa in ["SALUD", "TECFA", "FISIO"]:
                    programa = "SALUD"
                elif programa in ["SWSIS", "REDES"]:
                    programa = "SWSIS"
                # elif programa in ["COMAU", "PA"]:
                #    programa = "DG"

                df_sheet1 = pd.read_excel(
                    'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\ingresarCodigos.xlsx',
                    sheet_name='sheet1')

                df_sheet1['Modalidad'] = df_sheet1['Modalidad'].fillna('').astype(str)

                # Verificar la modalidad del alumno
                if modalidad_alumno == "Virtual":
                    # Si el alumno es "Virtual", no aplicar filtro de sede
                    filtro = (df_sheet1['Programa'].str.contains(programa)) & (
                        df_sheet1['Venta Nueva'].str.contains(vn))
                else:
                    # Aplicar el filtro de sede si la modalidad no es "Virtual"
                    filtro = (df_sheet1['Programa'].str.contains(programa)) & (df_sheet1['Campus'] == campus) & (
                        df_sheet1['Venta Nueva'].str.contains(vn))

                print(
                    f"Filtro aplicado: (Campus: {campus}, Turno: {turno}, Venta Nueva: {vn}, Programa: {programa}, Modalidad: {modalidad_alumno})")

                # Obtener todas las secciones disponibles para el primer filtro
                secciones_disponibles = df_sheet1[filtro]['Sección'].unique()

                print(f"Secciones disponibles: {secciones_disponibles}")

                # Definir las modalidades a buscar según las consideraciones
                modalidades_a_buscar = []

                # Caso 1: Si la modalidad del alumno está vacía o es NaN, buscar primero Presencial y luego Semipresencial
                if pd.isna(modalidad_alumno) or modalidad_alumno == "":
                    modalidades_a_buscar = ["Presencial", "Semipresencial"]
                else:
                    # Caso 2: Si la modalidad del alumno es Presencial, buscar primero Presencial y luego Semipresencial
                    if modalidad_alumno == "Presencial":
                        modalidades_a_buscar = ["Presencial", "Semipresencial"]
                    # Caso 3: Si la modalidad del alumno es Semipresencial, buscar primero Semipresencial y luego Presencial
                    elif modalidad_alumno == "Semipresencial":
                        modalidades_a_buscar = ["Semipresencial", "Presencial"]
                    # Caso 4: Si la modalidad del alumno es Virtual, buscar solo Virtual
                    elif modalidad_alumno == "Virtual":
                        modalidades_a_buscar = ["Virtual"]

                # Definir el orden de búsqueda de turnos según las consideraciones proporcionadas
                orden_busqueda = []

                if turno == 'A':
                    orden_busqueda = ['A', 'M', 'C']
                elif turno == 'C':
                    orden_busqueda = ['C', 'M', 'A']
                elif turno == 'M':
                    orden_busqueda = ['M', 'A', 'C']
                    # Verificar si no hay secciones en 'M', pero hay al menos una en 'A' o 'C'
                    if modalidad_alumno == "Virtual":
                        # Si el alumno es "Virtual", buscar solo las secciones con modalidad "Virtual" en el turno 'M'
                        filtro_modalidad_turno = filtro & (df_sheet1['Modalidad'].str.strip() == "Virtual") & (
                                df_sheet1['Turno'] == 'M')
                        secciones_disponibles_modalidad_turno = df_sheet1[filtro_modalidad_turno]['Sección'].unique()

                    else:
                        # Verificar si hay secciones disponibles en el turno 'M' para la modalidad del alumno
                        filtro_modalidad_turno = filtro & (df_sheet1['Turno'] == 'M')
                        secciones_disponibles_modalidad_turno = df_sheet1[filtro_modalidad_turno]['Sección'].unique()

                        if len(secciones_disponibles_modalidad_turno) == 0:
                            # No hay secciones disponibles en el turno 'M' para la modalidad del alumno
                            # Verificar si hay al menos una sección en los turnos 'A' o 'C'
                            if not df_sheet1[(df_sheet1['Turno'].isin(['A', 'C'])) & filtro].empty:
                                # Marcar un mensaje de error y detener la matrícula
                                error_msg_consultar = f"Se debe consultar turno de matricula a 'A' o 'C'. Programa: {programa}, Sede: {campus}"
                                print(error_msg_consultar)
                                print("------------------------------------------------------------------------")
                                print(" ")
                                errores.append((codigo, error_msg_consultar))
                                click_home_button()
                                time.sleep(4)
                                # Salir del bucle actual
                                continue
                elif turno == 'N':
                    orden_busqueda = ['N']

                seccion_elegida = None
                for turno_busqueda in orden_busqueda:
                    for modalidad in modalidades_a_buscar:
                        # Filtrar las secciones que coinciden con la modalidad, el turno de búsqueda y la sección
                        filtro_modalidad_turno = filtro & (df_sheet1['Modalidad'].str.strip() == modalidad.strip()) & (
                                df_sheet1['Turno'] == turno_busqueda)
                        secciones_disponibles_modalidad_turno = df_sheet1[filtro_modalidad_turno]['Sección'].unique()

                        print(
                            f"Secciones disponibles para {modalidad} en turno {turno_busqueda}: {secciones_disponibles_modalidad_turno}")

                        if len(secciones_disponibles_modalidad_turno) > 0:
                            # Se encontraron secciones para la modalidad, turno y sección actual, proceder con la primera
                            seccion_elegida = secciones_disponibles_modalidad_turno[0]
                            break

                    if seccion_elegida is not None:
                        break

                print(secciones_disponibles_modalidad_turno)

                if seccion_elegida is not None:
                    # Aplicar el filtro para obtener la fila correspondiente a la sección seleccionada
                    filtro_seccion_elegida = df_sheet1['Sección'] == seccion_elegida
                    fila_seccion_elegida = df_sheet1[filtro & filtro_seccion_elegida]

                    if not fila_seccion_elegida.empty:
                        # Obtener la modalidad de la sección seleccionada
                        modalidad_seccion = fila_seccion_elegida['Modalidad'].iloc[0]

                        # Obtener el estado de la sección
                        estado_seccion = fila_seccion_elegida['Estado'].iloc[0]

                        # Verificar el estado de la sección
                        if estado_seccion == "" or pd.isna(estado_seccion):
                            # Si la columna Estado está vacía o es NaN, considerar como "Aperturado"
                            print(f"Sección seleccionada: {seccion_elegida}")

                        if estado_seccion == "Aperturado":
                            print(f"Sección seleccionada: {seccion_elegida}")

                        elif estado_seccion == "Cerrado":
                            # Marcar un mensaje de error y continuar con la siguiente sección si hay disponible
                            print(f"Sección cerrada para matrícula: {seccion_elegida} Sede: {campus} Programa: {programa}")
                            # Convertir secciones_disponibles a lista
                            secciones_disponibles_modalidad_turno = list(secciones_disponibles_modalidad_turno)

                            # Continuar con la siguiente sección si hay disponible
                            secciones_disponibles_modalidad_turno.remove(
                                seccion_elegida)  # Remover la sección actual de las disponibles
                            if len(secciones_disponibles_modalidad_turno) > 0:
                                seccion_elegida = secciones_disponibles_modalidad_turno[0]
                                print(f"Continuando con la siguiente sección disponible: {seccion_elegida}")
                            else:
                                # No hay más secciones disponibles
                                error_msg = f"Seccion cerrada para matricula: {seccion_elegida} Sede: {campus} Programa: {programa}"
                                print(
                                    f"Sección cerrada para matrícula: {seccion_elegida} Sede: {campus} Programa: {programa}")
                                print("------------------------------------------------------------------------")
                                print(" ")
                                errores.append((codigo, error_msg))
                                click_home_button()
                                time.sleep(4)
                                continue

                        elif estado_seccion == "Sin cupo":
                            # Marcar un mensaje de error y continuar con la siguiente sección si hay disponible
                            print(f"Sección sin cupo para matrícula: {seccion_elegida}")
                            # Convertir secciones_disponibles a lista
                            secciones_disponibles_modalidad_turno = list(secciones_disponibles_modalidad_turno)

                            # Continuar con la siguiente sección si hay disponible
                            secciones_disponibles_modalidad_turno.remove(
                                seccion_elegida)  # Remover la sección actual de las disponibles
                            if len(secciones_disponibles_modalidad_turno) > 0:
                                seccion_elegida = secciones_disponibles_modalidad_turno[0]
                                print(f"Continuando con la siguiente sección disponible: {seccion_elegida}")
                            else:
                                # No hay más secciones disponibles
                                error_msg = f"Sección sin cupo para matrícula: {seccion_elegida} Programa: {programa} Sede: {campus}."
                                print(
                                    f"Codigo de Alumno {codigo} - Sección sin cupo para matrícula: {seccion_elegida} Programa: {programa} Sede: {campus}.")
                                print("------------------------------------------------------------------------")
                                print(" ")
                                errores.append((codigo, error_msg))
                                click_home_button()
                                time.sleep(4)
                                continue

                    # print(f"Sección seleccionada: {seccion_elegida}")

                    # Verificar si hay cambio de modalidad
                    if modalidad_alumno != modalidad_seccion:
                        mensaje_cambio_modalidad = f"Se cambia de {modalidad_alumno} a modalidad {modalidad_seccion}"
                        print(mensaje_cambio_modalidad)

                    # Filtrar los cursos para que coincidan con la sección seleccionada
                    filtro_seccion = (df_sheet1['Sección'] == seccion_elegida)
                    cursos_seleccionados = df_sheet1[filtro & filtro_seccion]
                    cursos_matriculados_exitosamente = True  # Variable para rastrear si el estudiante se matriculó en todos los cursos

                    # Inicializar un conjunto para rastrear los números de clase matriculados en esta sección
                    numeros_de_clase_matriculados = set()

                    for index, row in cursos_seleccionados.iterrows():
                        nombre_curso = row['Nombre Curso']
                        numero_clase = int(row['Nº Clase'])

                        # Verifica si el número de clase ya ha sido matriculado en esta sección
                        if numero_clase in numeros_de_clase_matriculados:
                            print(f"El curso con Nº Clase {numero_clase} ya ha sido matriculado en esta sección.")
                            continue  # Salta este curso y pasa al siguiente

                        # Hacer clic en el curso solo si hay cupo disponible
                        enlace_curso = WebDriverWait(driver, 16).until(
                            EC.element_to_be_clickable((By.XPATH, f"//label[contains(text(), '{nombre_curso}')]"))
                        )
                        enlace_curso.click()

                        time.sleep(2)

                        # Buscar el elemento de la fila que contiene el número de clase
                        fila_curso = WebDriverWait(driver, 16).until(
                            EC.presence_of_element_located((By.XPATH, f"//tr[contains(.//label, '{numero_clase}')]"))
                        )

                        time.sleep(2)

                        # Inicializar cupo_disponible con un valor predeterminado de 0
                        cupo_disponible = 0

                        # Obtener el elemento que contiene el cupo disponible
                        cupo_element = fila_curso.find_element(by=By.CSS_SELECTOR,
                                                               value="td.icon_container.p-1.text-center")

                        # Obtener el cupo disponible
                        cupo_disponible = int(cupo_element.text)

                        if cupo_disponible > 0:
                            # Hacer clic en el enlace del turno
                            enlace_turno = WebDriverWait(driver, 16).until(
                                EC.element_to_be_clickable((By.XPATH, f"//label[contains(text(), '{numero_clase}')]"))
                            )
                            enlace_turno.click()

                            time.sleep(2)

                            # Hacer clic en el botón matricular
                            boton_matricular = WebDriverWait(driver, 16).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-orange.m-2"))
                            )
                            boton_matricular.click()

                            time.sleep(2)

                            # Esperar a que aparezca el mensaje "Curso matriculado"
                            try:
                                mensaje_curso_matriculado = WebDriverWait(driver, 5).until(
                                    EC.visibility_of_element_located((By.XPATH,
                                                                      "//div[@aria-live='polite' and contains(text(), 'Curso matriculado')]"))
                                )
                                print(f"{nombre_curso} - {numero_clase} : Curso matriculado")
                                matriculado_algun_curso = True
                                # Marcar el número de clase como matriculado en esta sección
                                numeros_de_clase_matriculados.add(numero_clase)
                            except TimeoutException:
                                print(f"{nombre_curso} - {numero_clase} No se pudo matricular al curso")
                                cursos_matriculados_exitosamente = False

                            time.sleep(4)

                        else:
                            # Mensaje de error en la consola
                            error_msg = f"Código de Alumno {codigo} - No hay cupo disponible para el curso: {nombre_curso} (Nº Clase: {numero_clase}) Sección: {seccion_elegida}, Programa: {programa} "
                            print(error_msg)

                            # Registrar el error en el archivo CSV
                            errores.append((codigo, error_msg))
                            cursos_matriculados_exitosamente = False
                            # Si no hay cupo, volver a la página principal y registrar el error
                            click_home_button()
                            time.sleep(4)

                    if cursos_matriculados_exitosamente:
                        # Imprimir mensaje de matrícula completa solo si todos los cursos de la sección se han matriculado con éxito
                        print(f"{codigo} Matricula completa")
                        errores.append((codigo, "Matricula completa"))

                    # Volver a la página principal después de matricular todos los cursos
                    click_home_button()
                    time.sleep(4)

                else:
                    # No se encontraron secciones disponibles para el programa, campus y turno, agregar un mensaje de error
                    error_msg = f"No se encontraron secciones disponibles para la matricula. Sede: {campus}"
                    print(
                        f"Código de Alumno {codigo} - No se encontraron secciones disponibles para la matricula. Sede: {campus}")
                    errores.append((codigo, error_msg))
                    click_home_button()
                    time.sleep(4)

            else:
                ########################## VENTA ANTIGUA #############################

                print(f"Código de Alumno {codigo} - Venta antigua")
                # Buscar al alumno en sheet3 y obtener su información de ciclos y bloques
                df_sheet1 = pd.read_excel(
                    'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\ingresarCodigos.xlsx',
                    sheet_name='sheet1')

                df_sheet3 = pd.read_excel(
                    'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\ingresarCodigos.xlsx',
                    sheet_name='sheet3')

                alumno_sheet3 = df_sheet3[df_sheet3['Id_Alumno'] == codigo]
                if alumno_sheet3.empty:
                    error_message = "Probable Reingreso - Debe proceder una matricula administrativa"
                    errores.append((codigo, error_message))
                    print(f"Código de Alumno {codigo} - {error_message}")
                else:
                    alumno_sheet3 = alumno_sheet3.iloc[0]

                # Lista de las columnas a verificar en alumno_sheet3
                columnas_a_verificar = ['C1B1', 'C1B2', 'C2B1', 'C2B2', 'C3B1', 'C3B2', 'C4B1', 'C4B2', 'C5B1', 'C5B2',
                                        'C6B1', 'C6B2']

                # Determinar internamente el ciclo y bloque de matriculación en sheet1
                ciclo_actual = None
                bloque_actual = None
                ciclo_actual_int = None  # Declarar ciclo_actual_int
                encontrado = False  # Variable para rastrear si se encontraron valores 'OK' en las columnas

                for columna in columnas_a_verificar:
                    value = alumno_sheet3[columna]

                    print(f"Ciclo/Bloque: {columna}, Valor: {value}")  # Agregar esta línea para depurar

                    if 'OK' in str(value) or value == 1:  # Verifica si la cadena "OK" está contenida en el valor
                        ciclo_actual, bloque_actual = columna.split('B')
                        ciclo_actual_int = int(ciclo_actual.split('C')[1])  # Asignar valor a ciclo_actual_int
                        encontrado = True
                        # No uses break aquí si deseas verificar todas las columnas

                if encontrado:
                    # Ciclos anteriores al ciclo actual
                    ciclos_anteriores = [f'C{i}' for i in range(1, int(ciclo_actual.split('C')[1]))]

                    # Verificar si todos los bloques de los ciclos anteriores están completos
                    if (alumno_sheet3[f'{ciclo_actual}B1'] == 'OK' or alumno_sheet3[f'{ciclo_actual}B1'] == 1) and (
                            alumno_sheet3[f'{ciclo_actual}B2'] == 'OK' or alumno_sheet3[f'{ciclo_actual}B2'] == 1):
                        caso = 1  # Caso 1: Puede matricular en B1 o B2
                    elif alumno_sheet3[f'{ciclo_actual}B1'] == 'OK' or alumno_sheet3[f'{ciclo_actual}B1'] == 1:
                        caso = 2  # Caso 2: Completo B1 de un ciclo, solo OK en B1.
                    elif alumno_sheet3[f'{ciclo_actual}B2'] == 'OK' or alumno_sheet3[f'{ciclo_actual}B2'] == 1:
                        caso = 3  # Caso 3: Completo B2 de un ciclo, solo OK en B2.

                    if caso == 1:
                        # Caso 1: Completo ciclos y ambos bloques: B1 y B2 con OK.
                        ciclo_actual_int = int(ciclo_actual.split('C')[1])
                        siguiente_ciclo_matricula = f'C{ciclo_actual_int + 1}'

                        # Comprobar si todos los ciclos anteriores están completos
                        ciclos_anteriores = [f'C{i}' for i in range(1, ciclo_actual_int)]
                        if all(
                                (alumno_sheet3[f'{ciclo}B1'] in ['OK', 1]) and
                                (alumno_sheet3[f'{ciclo}B2'] in ['OK', 1])
                                for ciclo in ciclos_anteriores
                        ):
                            df_sheet1 = pd.read_excel(
                                'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\ingresarCodigos.xlsx',
                                sheet_name='sheet1')
                            # Verificar si es matriculado en 5to o 6to ciclo
                            if int(siguiente_ciclo_matricula[1]) in [5, 6]:
                                # Lógica para determinar las opciones de sede según la sede actual
                                opciones_sede = {
                                    'ATE': ['ATE', 'SJL', 'STABE'],
                                    'SJL': ['SJL', 'ATE', 'STABE'],
                                    'STABE': ['STABE', 'MIRAF'],
                                    'SURCO': ['SURCO', 'SJM', 'MIRAF', 'STABE'],
                                    'SJM': ['SJM', 'SURCO', 'MIRAF', 'STABE'],
                                    'MIRAF': ['MIRAF', 'SURCO', 'STABE', 'SJM'],
                                    'INDEP': ['INDEP', 'VENTN', 'STABE'],
                                    'VENTN': ['VENTN', 'INDEP', 'STABE'],
                                    'VILLA': ['VILLA', 'SJM', 'SURCO']
                                }

                                # Obtener la sede actual del alumno
                                sede_actual = campus

                                # Verificar si la sede actual está en las opciones
                                if sede_actual in opciones_sede:
                                    # Filtrar las sedes según la sede actual
                                    sedes_permitidas = opciones_sede[sede_actual]

                                    for sede in sedes_permitidas:
                                        filtro = (
                                                (df_sheet1['Campus'] == sede) & (df_sheet1['Turno'] == turno) & (
                                            df_sheet1['Venta Nueva'].str.contains(vn)) & (
                                                    df_sheet1['Programa'].str.contains(programa)) & (
                                                        (df_sheet1['Ciclo'] == siguiente_ciclo_matricula) | (
                                                        df_sheet1['Ciclo'] == f"{siguiente_ciclo_matricula}B1") | (
                                                                df_sheet1['Ciclo'] == f"{siguiente_ciclo_matricula}B2"))
                                        )
                                        # Verificar si hay alguna fila que cumple con el filtro
                                        if any(filtro):
                                            if sede_actual != sede:
                                                print(
                                                    f"No apertura en sede del alumno en {sede_actual}, probando siguientes sedes...")
                                                print(f"Se cambia sede: De {sede_actual} a sede {sede}")

                                            break  # Salir del bucle si se encuentra una sede válida
                                    else:
                                        print(f"No hay sedes disponibles para matrícula en {sedes_permitidas}")
                                else:
                                    print(f"Sede actual no reconocida: {sede_actual}")

                            else:
                                # En caso de 1, 2, 3, 4, aplicar el filtro original sin modificaciones
                                filtro = (df_sheet1['Campus'] == campus) & (df_sheet1['Turno'] == turno) & (
                                    df_sheet1['Venta Nueva'].str.contains(vn)) & (
                                             df_sheet1['Programa'].str.contains(programa)) & (
                                                 (df_sheet1['Ciclo'] == siguiente_ciclo_matricula) | (
                                                 df_sheet1['Ciclo'] == f"{siguiente_ciclo_matricula}B1") | (
                                                         df_sheet1['Ciclo'] == f"{siguiente_ciclo_matricula}B2"))
                        else:
                            caso = 4
                            print(
                                f"{codigo} Caso 4: El alumno ha adelantado cursos en ciclos a los que no le pertenecen. No se permite la matrícula.")
                            print("------------------------------------------------------------------------")
                            print(" ")
                            # Caso 4: Adelanto irregular de cursos, no se permite la matriculación
                            error_message = "Caso 4: El alumno ha adelantado cursos en ciclos a los que no le pertenecen. Debe proceder una matricula administrativa."
                            # Agregar el error a la lista de errores
                            errores.append((codigo, error_message))
                            # Continuar con el siguiente código de alumno
                            click_home_button()
                            time.sleep(1)
                            continue

                    elif caso == 2:
                        # Caso 2: Completo B1 de un ciclo, solo OK en B1, debe verificar completitud de ciclos anteriores.
                        ciclos_anteriores = [f'C{i}' for i in range(1, ciclo_actual_int)]
                        if all(
                                (alumno_sheet3[f'{ciclo}B1'] in ['OK', 1]) and
                                (alumno_sheet3[f'{ciclo}B2'] in ['OK', 1])
                                for ciclo in ciclos_anteriores
                        ):
                            siguiente_ciclo_matricula = f'{ciclo_actual}'  # Sigue en el mismo ciclo
                            siguiente_bloque_matricula = 'B2'  # Puede matricular en B2 del mismo ciclo
                            df_sheet1 = pd.read_excel(
                                'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\ingresarCodigos.xlsx',
                                sheet_name='sheet1')
                            # Verificar si es matriculado en 5to o 6to ciclo
                            if int(siguiente_ciclo_matricula[1]) in [5, 6]:
                                # Lógica para determinar las opciones de sede según la sede actual
                                opciones_sede = {
                                    'ATE': ['ATE', 'SJL', 'STABE'],
                                    'SJL': ['SJL', 'ATE', 'STABE'],
                                    'STABE': ['STABE', 'MIRAF'],
                                    'SURCO': ['SURCO', 'SJM', 'MIRAF', 'STABE'],
                                    'SJM': ['SJM', 'SURCO', 'MIRAF', 'STABE'],
                                    'MIRAF': ['MIRAF', 'SURCO', 'STABE', 'SJM'],
                                    'INDEP': ['INDEP', 'VENTN', 'STABE'],
                                    'VENTN': ['VENTN', 'INDEP', 'STABE'],
                                    'VILLA': ['VILLA', 'SJM', 'SURCO']
                                }

                                # Obtener la sede actual del alumno
                                sede_actual = campus

                                # Verificar si la sede actual está en las opciones
                                if sede_actual in opciones_sede:
                                    # Filtrar las sedes según la sede actual
                                    sedes_permitidas = opciones_sede[sede_actual]

                                    for sede in sedes_permitidas:
                                        filtro = (df_sheet1['Campus'] == sede) & (df_sheet1['Turno'] == turno) & (
                                            df_sheet1['Venta Nueva'].str.contains(vn)) & (
                                                     df_sheet1['Programa'].str.contains(programa)) & (
                                                         df_sheet1['Ciclo'] == f'{ciclo_actual}B2')

                                        # Verificar si hay alguna fila que cumple con el filtro
                                        if any(filtro):
                                            if sede_actual != sede:
                                                print(
                                                    f"No apertura en sede del alumno en {sede_actual}, probando siguientes sedes...")
                                                print(f"Se cambia sede: De {sede_actual} a sede {sede}")

                                            break  # Salir del bucle si se encuentra una sede válida
                                    else:
                                        print(f"No hay sedes disponibles para matrícula en {sedes_permitidas}")
                                else:
                                    print(f"Sede actual no reconocida: {sede_actual}")
                            else:
                                filtro = (df_sheet1['Campus'] == campus) & (df_sheet1['Turno'] == turno) & (
                                    df_sheet1['Venta Nueva'].str.contains(vn)) & (
                                             df_sheet1['Programa'].str.contains(programa)) & (
                                                 df_sheet1['Ciclo'] == f'{ciclo_actual}B2')

                        else:
                            caso = 4
                            print(
                                f"{codigo} Caso 4: El alumno ha adelantado cursos en ciclos a los que no le pertenecen. No se permite la matrícula.")
                            print("------------------------------------------------------------------------")
                            print(" ")
                            # Caso 4: Adelanto irregular de cursos, no se permite la matriculación
                            error_message = "Caso 4: El alumno ha adelantado cursos en ciclos a los que no le pertenecen. Debe proceder una matricula administrativa."
                            # Agregar el error a la lista de errores
                            errores.append((codigo, error_message))
                            # Continuar con el siguiente código de alumno
                            click_home_button()
                            time.sleep(1)
                            continue

                    elif caso == 3:
                        # Caso 3: Completo B2 de un ciclo, solo OK en B1, debe verificar completitud de ciclos anteriores.
                        ciclos_anteriores = [f'C{i}' for i in range(1, ciclo_actual_int)]
                        if all(
                                (alumno_sheet3[f'{ciclo}B1'] in ['OK', 1]) and
                                (alumno_sheet3[f'{ciclo}B2'] in ['OK', 1])
                                for ciclo in ciclos_anteriores
                        ):
                            siguiente_ciclo_matricula = f'{ciclo_actual}'
                            siguiente_bloque_matricula = 'B1'
                            df_sheet1 = pd.read_excel(
                                'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\ingresarCodigos.xlsx',
                                sheet_name='sheet1')
                            # Verificar si es matriculado en 5to o 6to ciclo
                            if int(siguiente_ciclo_matricula[1]) in [5, 6]:
                                # Lógica para determinar las opciones de sede según la sede actual
                                opciones_sede = {
                                    'ATE': ['ATE', 'SJL', 'STABE'],
                                    'SJL': ['SJL', 'ATE', 'STABE'],
                                    'STABE': ['STABE', 'MIRAF'],
                                    'SURCO': ['SURCO', 'SJM', 'MIRAF', 'STABE'],
                                    'SJM': ['SJM', 'SURCO', 'MIRAF', 'STABE'],
                                    'MIRAF': ['MIRAF', 'SURCO', 'STABE', 'SJM'],
                                    'INDEP': ['INDEP', 'VENTN', 'STABE'],
                                    'VENTN': ['VENTN', 'INDEP', 'STABE'],
                                    'VILLA': ['VILLA', 'SJM', 'SURCO']
                                }

                                # Obtener la sede actual del alumno
                                sede_actual = campus

                                # Verificar si la sede actual está en las opciones
                                if sede_actual in opciones_sede:
                                    # Filtrar las sedes según la sede actual
                                    sedes_permitidas = opciones_sede[sede_actual]

                                    for sede in sedes_permitidas:

                                        filtro = (df_sheet1['Campus'] == sede) & (df_sheet1['Turno'] == turno) & (
                                            df_sheet1['Venta Nueva'].str.contains(vn)) & (
                                                     df_sheet1['Programa'].str.contains(programa)) & (
                                                         df_sheet1['Ciclo'] == f'{ciclo_actual}B1')

                                        # Verificar si hay alguna fila que cumple con el filtro
                                        if any(filtro):
                                            if sede_actual != sede:
                                                print(
                                                    f"No apertura en sede del alumno en {sede_actual}, probando siguientes sedes...")
                                                print(f"Se cambia sede: De {sede_actual} a sede {sede}")

                                            break  # Salir del bucle si se encuentra una sede válida
                                    else:
                                        print(f"No hay sedes disponibles para matrícula en {sedes_permitidas}")
                                else:
                                    print(f"Sede actual no reconocida: {sede_actual}")
                            else:
                                filtro = (df_sheet1['Campus'] == campus) & (df_sheet1['Turno'] == turno) & (
                                    df_sheet1['Venta Nueva'].str.contains(vn)) & (
                                             df_sheet1['Programa'].str.contains(programa)) & (
                                                 df_sheet1['Ciclo'] == f'{ciclo_actual}B1')
                        else:
                            caso = 4
                            print(
                                f"{codigo} Caso 4: El alumno ha adelantado cursos en ciclos a los que no le pertenecen. No se permite la matrícula.")
                            print("------------------------------------------------------------------------")
                            print(" ")
                            # Caso 4: Adelanto irregular de cursos, no se permite la matriculación
                            error_message = "Caso 4: El alumno ha adelantado cursos en ciclos a los que no le pertenecen. Debe proceder una matricula administrativa."
                            # Agregar el error a la lista de errores
                            errores.append((codigo, error_message))
                            # Continuar con el siguiente código de alumno
                            click_home_button()
                            time.sleep(1)
                            continue

                    print(f"Caso identificado: {caso}")
                    print(
                        f"Ciclo de Matriculación: {siguiente_ciclo_matricula}, Bloque de Matriculación: {siguiente_bloque_matricula}")

                else:
                    print(f"No se encontraron valores 'OK' en las columnas específicas para el alumno {codigo}")

                print(
                    f"Filtro aplicado: (Campus: {campus}, Turno: {turno}, Venta Nueva: {vn}, Programa: {programa}, Ciclo/Bloque: {siguiente_ciclo_matricula}{siguiente_bloque_matricula})")

                # Obtener todas las secciones disponibles para el programa, campus y turno dados
                secciones_disponibles = df_sheet1[filtro]['Sección'].unique()

                print(f"Secciones disponibles: {secciones_disponibles}")

                print(f"Sección seleccionada: {secciones_disponibles[0]}" if len(
                    secciones_disponibles) > 0 else "No se encontraron secciones disponibles")

                # Procesamiento previo para depurar las secciones y eliminar duplicaciones de números de clase
                secciones_limpia = []  # Lista para rastrear las secciones depuradas

                for seccion in secciones_disponibles:
                    numeros_de_clase_en_seccion = set()  # Conjunto para rastrear los números de clase en esta sección
                    cursos_de_seccion = df_sheet1[filtro & (df_sheet1['Sección'] == seccion)]

                    # Verificar cada curso de la sección
                    for index, row in cursos_de_seccion.iterrows():
                        numero_clase = row['Nº Clase']

                        # Verificar si el número de clase ya existe en esta sección
                        if numero_clase not in numeros_de_clase_en_seccion:
                            numeros_de_clase_en_seccion.add(numero_clase)

                    # Agregar la sección depurada solo si contiene cursos con números de clase únicos
                    if len(numeros_de_clase_en_seccion) > 0:
                        secciones_limpia.append(seccion)

                # Seleccionar la primera sección encontrada
                primera_seccion = None
                for seccion in secciones_disponibles:
                    primera_seccion = seccion
                    break  # Romper el bucle después de encontrar la primera sección
                    print(f"Primera sección seleccionada: {primera_seccion}")

                if primera_seccion is not None:
                    # Filtrar los cursos para que coincidan con la sección seleccionada
                    filtro_seccion = (df_sheet1['Sección'] == primera_seccion)
                    cursos_seleccionados = df_sheet1[filtro & filtro_seccion]
                    cursos_matriculados_exitosamente = True  # Variable para rastrear si el estudiante se matriculó en todos los cursos

                    # Inicializa un conjunto para rastrear los números de clase matriculados en esta sección
                    numeros_de_clase_matriculados = set()

                    for index, row in cursos_seleccionados.iterrows():
                        nombre_curso = row['Nombre Curso']
                        numero_clase = int(row['Nº Clase'])

                        # Verifica si el número de clase ya ha sido matriculado en esta sección
                        if numero_clase in numeros_de_clase_matriculados:
                            print(f"El curso con Nº Clase {numero_clase} ya ha sido matriculado en esta sección.")
                            continue  # Salta este curso y pasa al siguiente

                        # Hacer clic en el curso solo si hay cupo disponible
                        enlace_curso = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, f"//label[contains(text(), '{nombre_curso}')]"))
                        )
                        enlace_curso.click()

                        time.sleep(2)

                        # Buscar el elemento de la fila que contiene el número de clase
                        fila_curso = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, f"//tr[contains(.//label, '{numero_clase}')]"))
                        )

                        time.sleep(2)

                        # Inicializar cupo_disponible con un valor predeterminado de 0
                        cupo_disponible = 0

                        # Obtener el elemento que contiene el cupo disponible
                        cupo_element = fila_curso.find_element(by=By.CSS_SELECTOR,
                                                               value="td.icon_container.p-1.text-center")

                        # Obtener el cupo disponible
                        cupo_disponible = int(cupo_element.text)

                        if cupo_disponible > 0:

                            # Hacer clic en el enlace del turno
                            enlace_turno = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, f"//label[contains(text(), '{numero_clase}')]"))
                            )
                            enlace_turno.click()

                            time.sleep(2)

                            # Hacer clic en el botón matricular
                            boton_matricular = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-orange.m-2"))
                            )
                            boton_matricular.click()

                            # Esperar a que aparezca el mensaje "Curso matriculado"
                            try:
                                mensaje_curso_matriculado = WebDriverWait(driver, 5).until(
                                    EC.visibility_of_element_located((By.XPATH,
                                                                      "//div[@aria-live='polite' and contains(text(), 'Curso matriculado')]"))
                                )
                                print("Curso matriculado")
                                matriculado_algun_curso = True
                                # Marcar el número de clase como matriculado en esta sección
                                numeros_de_clase_matriculados.add(numero_clase)
                            except TimeoutException:
                                print("No se pudo matricular al curso")
                                cursos_matriculados_exitosamente = False

                            time.sleep(4)

                        else:
                            # Mensaje de error en la consola
                            error_msg = f"Código de Alumno {codigo} - No hay cupo disponible para el curso: {nombre_curso} (Nº Clase: {numero_clase}) Sección: {secciones_disponibles}, Programa: {programa} "
                            print(error_msg)

                            # Registrar el error en el archivo CSV
                            errores.append((codigo, error_msg))
                            cursos_matriculados_exitosamente = False
                            # Si no hay cupo, volver a la página principal y registrar el error
                            click_home_button()
                            time.sleep(4)

                    if cursos_matriculados_exitosamente:
                        # Imprimir mensaje de matrícula completa solo si todos los cursos de la sección se han matriculado con éxito
                        print(f"{codigo} Matricula completa")
                        errores.append((codigo, "Matricula completa"))

                    # Volver a la página principal después de matricular todos los cursos
                    click_home_button()
                    time.sleep(4)

                else:
                    # No se encontraron secciones disponibles para el programa, campus y turno, agregar un mensaje de error
                    error_msg = f"No se encontraron secciones disponibles para la matricula."
                    print(f"Código de Alumno {codigo} - No se encontraron secciones disponibles para la matricula.")
                    errores.append((codigo, error_msg))
                    click_home_button()
                    time.sleep(4)

        except TimeoutException:
            # Si se produce un TimeoutException, determina si no se encontraron cursos o si la matrícula es incompleta
            print(f"TimeoutException para el código {codigo}: Curso '{nombre_curso}' (Nº Clase: {numero_clase})")

            if vn == 'SI':
                if not matriculado_algun_curso:
                    error_msg = f"No se encontraron cursos para matricular Seccion: {seccion_elegida} Programa {programa}"
                    print(
                        f"Código: {codigo}, Error: No se encontraron cursos para matricular Seccion: {seccion_elegida} Programa {programa}")
                else:
                    error_msg = f"MATRICULA INCOMPLETA - Seccion: {seccion_elegida}, Numero de Clase: {numero_clase}, Nombre del Curso: {nombre_curso}, Sede: {campus}, Turno: {turno}, Programa: {programa}"
                    # Imprimir información adicional si hay una MATRÍCULA INCOMPLETA
                    print(
                        f" MATRICULA INCOMPLETA = Código: {codigo}, Sección: {seccion_elegida}, Número de Clase: {numero_clase}, Nombre del Curso: {nombre_curso}, Sede: {campus}, Turno: {turno}, Programa: {programa}")
                errores.append((codigo, error_msg))
                click_home_button()
                time.sleep(8)

            else:
                if not matriculado_algun_curso:
                    error_msg = f"No se encontraron cursos para matricular Seccion: {secciones_disponibles[0]} Programa {programa}"
                    print(
                        f"Código: {codigo}, Error: No se encontraron cursos para matricular Seccion: {secciones_disponibles[0]} Programa {programa}")
                else:
                    error_msg = f"MATRICULA INCOMPLETA - Seccion: {secciones_disponibles[0]}, Numero de Clase: {numero_clase}, Nombre del Curso: {nombre_curso}, Sede: {campus}, Turno: {turno}, Programa: {programa}"
                    # Imprimir información adicional si hay una MATRÍCULA INCOMPLETA
                    print(
                        f" MATRICULA INCOMPLETA = Código: {codigo}, Sección: {secciones_disponibles}, Número de Clase: {numero_clase}, Nombre del Curso: {nombre_curso}, Sede: {campus}, Turno: {turno}, Programa: {programa}")
                errores.append((codigo, error_msg))
                click_home_button()
                time.sleep(8)


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
        print("------------------------------------------------------------------------")
        print(" ")

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
            with open('../../../../data/matricula/salida/reporte.csv', 'w', newline='', encoding='utf-8') as file:
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

