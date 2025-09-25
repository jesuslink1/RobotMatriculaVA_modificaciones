import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def matricular_alumno_vn_si(codigo, campus, turno, vn, programa, modalidad_alumno, driver, errores, click_home_button):
    # Agregar condiciones para ajustar el programa si es uno de los programas especificados
    if programa in ["ADMNI", "MARKT"]:
        programa = "ADMNE"
    elif programa in ["ADMBF", "CONTA"]:
        programa = "ADMBF"
    elif programa in ["SALUD", "TECFA", "FISIO"]:
        programa = "SALUD"
    elif programa in ["SWSIS", "REDES"]:
        programa = "SWSIS"
    elif programa in ["COMAU", "PA"]:
        programa = "DG"

    df_sheet1 = pd.read_excel('C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\ingresarCodigos.xlsx',
                              sheet_name='sheet1')

    # Filtrar cursos por sede, programa y turno
    filtro = (df_sheet1['Programa'].str.contains(programa)) & (df_sheet1['Campus'] == campus) & (
        df_sheet1['Venta Nueva'].str.contains(vn)) & (df_sheet1['Turno'] == turno) & (
                 df_sheet1['Modalidad'].str.contains(modalidad_alumno))

    print(
        f"Filtro aplicado: (Campus: {campus}, Turno: {turno}, Venta Nueva: {vn}, Programa: {programa})")

    # Obtener todas las secciones disponibles para el programa, campus y turno dados
    secciones_disponibles = df_sheet1[filtro]['Sección'].unique()

    print(f"Secciones disponibles: {secciones_disponibles}")

    print(f"Sección seleccionada: {secciones_disponibles[0]}" if len(
        secciones_disponibles) > 0 else "No se encontraron secciones disponibles")

    # Definir las modalidades a buscar según las consideraciones
    modalidades_a_buscar = []

    # Caso 1: Si la modalidad del alumno está vacía, buscar primero Presencial y luego Semipresencial
    if modalidad_alumno == "":
        modalidades_a_buscar = ["Presencial", "Semipresencial"]
    else:
        # Caso 2: Si la modalidad del alumno es Presencial, buscar primero Presencial y luego Semipresencial
        if modalidad_alumno == "Presencial":
            modalidades_a_buscar = ["Presencial", "Semipresencial"]
        # Caso 3: Si la modalidad del alumno es Semipresencial, buscar primero Semipresencial y luego Presencial
        elif modalidad_alumno == "Semipresencial":
            modalidades_a_buscar = ["Semipresencial", "Presencial"]

    # Iterar sobre las modalidades a buscar en el orden especificado
    seccion_elegida = None
    for modalidad in modalidades_a_buscar:
        # Agregar la condición de modalidad al filtro
        filtro_modalidad = filtro & (df_sheet1['Modalidad'] == modalidad)

        # Obtener todas las secciones disponibles para el programa, campus, turno y modalidad dados
        secciones_disponibles = df_sheet1[filtro_modalidad]['Sección'].unique()

        if len(secciones_disponibles) > 0:
            # Se encontraron secciones para la modalidad actual, proceder con la primera
            seccion_elegida = secciones_disponibles[0]
            break

    if seccion_elegida is not None:
        print(f"Sección seleccionada: {seccion_elegida}")

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

                time.sleep(2)

                # Marcar que se ha matriculado en este curso
                matriculado_algun_curso = True
                # Marca el número de clase como matriculado en esta sección
                numeros_de_clase_matriculados.add(numero_clase)

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
        error_msg = f"Código de Alumno {codigo} - No se encontraron secciones disponibles para la matricula."
        print(error_msg)
        errores.append((codigo, error_msg))
        click_home_button()
        time.sleep(4)