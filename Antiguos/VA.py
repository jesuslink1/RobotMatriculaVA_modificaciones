import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def matricular_alumno_vn_no(codigo, campus, turno, vn, programa, modalidad_alumno, driver, errores, click_home_button):
    # Buscar al alumno en sheet3 y obtener su información de ciclos y bloques
    df_sheet3 = pd.read_excel('C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\ingresarCodigos.xlsx',
                              sheet_name='sheet3')

    alumno_sheet3 = df_sheet3[df_sheet3['Id_Alumno'] == codigo]
    if alumno_sheet3.empty:
        error_message = "Probable Reingreso - Debe proceder una matrícula administrativa"
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
        # Agregar una condición para verificar si el valor es igual a 1
        if value == 1:
            value = 'OK'  # Establece el valor en 'OK' si es igual a 1

        print(f"Ciclo/Bloque: {columna}, Valor: {value}")  # Agregar esta línea para depurar

        if 'OK' in str(value):  # Verifica si la cadena "OK" está contenida en el valor
            ciclo_actual, bloque_actual = columna.split('B')
            ciclo_actual_int = int(ciclo_actual.split('C')[1])  # Asignar valor a ciclo_actual_int
            print(
                f"Ciclo actual antes de conversión: {ciclo_actual}, Bloque actual antes de conversión: {bloque_actual}")
            encontrado = True
            # No uses break aquí si deseas verificar todas las columnas

    if encontrado:
        # Ciclos anteriores al ciclo actual
        ciclos_anteriores = [f'C{i}' for i in range(1, int(ciclo_actual.split('C')[1]))]

        # Verificar si todos los bloques de los ciclos anteriores están completos
        if alumno_sheet3[f'{ciclo_actual}B1'] == 'OK' and alumno_sheet3[f'{ciclo_actual}B2'] == 'OK':
            caso = 1  # Caso 1: Puede matricular en B1 o B2
        elif alumno_sheet3[f'{ciclo_actual}B1'] == 'OK':
            caso = 2  # Caso 2: Completo B1 de un ciclo, solo OK en B1.
        elif alumno_sheet3[f'{ciclo_actual}B2'] == 'OK':
            caso = 3  # Caso 3: Completo B2 de un ciclo, solo OK en B2.
        else:
            caso = 4  # Caso 4: Adelanto irregular de cursos, no se permite la matriculación

        if caso == 1:
            # Caso 1: Completo ciclos y ambos bloques: B1 y B2 con OK.
            ciclo_actual_int = int(ciclo_actual.split('C')[1])
            siguiente_ciclo_matricula = f'C{ciclo_actual_int + 1}'

            # Comprobar si todos los ciclos anteriores están completos
            ciclos_anteriores = [f'C{i}' for i in range(1, ciclo_actual_int)]
            if all(alumno_sheet3[f'{ciclo}B1'] == 'OK' and alumno_sheet3[f'{ciclo}B2'] == 'OK' for ciclo in
                   ciclos_anteriores):
                df_sheet1 = pd.read_excel(
                    'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\ingresarCodigos.xlsx',
                    sheet_name='sheet1')
                filtro = (df_sheet1['Campus'] == campus) & (df_sheet1['Turno'] == turno) & (
                    df_sheet1['Venta Nueva'].str.contains(vn)) & (df_sheet1['Programa'].str.contains(programa)) & (
                                 (df_sheet1['Ciclo'] == siguiente_ciclo_matricula) | (
                                 df_sheet1['Ciclo'] == f"{siguiente_ciclo_matricula}B1") | (
                                         df_sheet1['Ciclo'] == f"{siguiente_ciclo_matricula}B2"))
            else:
                caso = 4
                print(
                    f"{codigo} Caso 4: El alumno ha adelantado cursos en ciclos a los que no le pertenecen. No se permite la matrícula.")
                # Caso 4: Adelanto irregular de cursos, no se permite la matriculación
                error_message = "Caso 4: El alumno ha adelantado cursos en ciclos a los que no le pertenecen. Debe proceder una matrícula administrativa."
                # Agregar el error a la lista de errores
                errores.append((codigo, error_message))
                # Continuar con el siguiente código de alumno
                click_home_button()
                time.sleep(1)

        elif caso == 2:
            # Caso 2: Completo B1 de un ciclo, solo OK en B1, debe verificar completitud de ciclos anteriores.
            ciclos_anteriores = [f'C{i}' for i in range(1, ciclo_actual_int)]
            if all(alumno_sheet3[f'{ciclo}B1'] == 'OK' and alumno_sheet3[f'{ciclo}B2'] == 'OK' for ciclo in
                   ciclos_anteriores):
                siguiente_ciclo_matricula = f'{ciclo_actual}'  # Sigue en el mismo ciclo
                siguiente_bloque_matricula = 'B2'  # Puede matricular en B2 del mismo ciclo
                df_sheet1 = pd.read_excel(
                    'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\ingresarCodigos.xlsx',
                    sheet_name='sheet1')
                filtro = (df_sheet1['Campus'] == campus) & (df_sheet1['Turno'] == turno) & (
                    df_sheet1['Venta Nueva'].str.contains(vn)) & (df_sheet1['Programa'].str.contains(programa)) & (
                                     df_sheet1['Ciclo'] == f'{ciclo_actual}B2')

            else:
                caso = 4
                print(
                    f"{codigo} Caso 4: El alumno ha adelantado cursos en ciclos a los que no le pertenecen. No se permite la matrícula.")
                # Caso 4: Adelanto irregular de cursos, no se permite la matriculación
                error_message = "Caso 4: El alumno ha adelantado cursos en ciclos a los que no le pertenecen. Debe proceder una matrícula administrativa."
                # Agregar el error a la lista de errores
                errores.append((codigo, error_message))
                # Continuar con el siguiente código de alumno
                click_home_button()
                time.sleep(1)

        elif caso == 3:
            # Caso 3: Completo B2 de un ciclo, solo OK en B1, debe verificar completitud de ciclos anteriores.
            ciclos_anteriores = [f'C{i}' for i in range(1, ciclo_actual_int)]
            if all(alumno_sheet3[f'{ciclo}B1'] == 'OK' and alumno_sheet3[f'{ciclo}B2'] == 'OK' for ciclo in
                   ciclos_anteriores):
                siguiente_ciclo_matricula = f'{ciclo_actual}'
                siguiente_bloque_matricula = 'B1'
                df_sheet1 = pd.read_excel(
                    'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\ingresarCodigos.xlsx',
                    sheet_name='sheet1')
                filtro = (df_sheet1['Campus'] == campus) & (df_sheet1['Turno'] == turno) & (
                    df_sheet1['Venta Nueva'].str.contains(vn)) & (df_sheet1['Programa'].str.contains(programa)) & (
                                     df_sheet1['Ciclo'] == f'{ciclo_actual}B1')
            else:
                caso = 4
                print(
                    f"{codigo} Caso 4: El alumno ha adelantado cursos en ciclos a los que no le pertenecen. No se permite la matrícula.")
                # Caso 4: Adelanto irregular de cursos, no se permite la matriculación
                error_message = "Caso 4: El alumno ha adelantado cursos en ciclos a los que no le pertenecen. Debe proceder una matrícula administrativa."
                # Agregar el error a la lista de errores
                errores.append((codigo, error_message))
                # Continuar con el siguiente código de alumno
                click_home_button()
                time.sleep(1)

        elif caso == 4:
            # Caso 4: Adelanto irregular de cursos, no se permite la matriculación
            print(
                f"{codigo} Caso 4: El alumno ha adelantado cursos en ciclos a los que no le pertenecen. No se permite la matrícula.")
            # Caso 4: Adelanto irregular de cursos, no se permite la matriculación
            error_message = "Caso 4: El alumno ha adelantado cursos en ciclos a los que no le pertenecen. Debe proceder una matricula administrativa."
            # Agregar el error a la lista de errores
            errores.append((codigo, error_message))
            # Continuar con el siguiente código de alumno
            click_home_button()
            time.sleep(1)

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

                time.sleep(2)

                # Marcar que se ha matriculado en este curso
                matriculado_algun_curso = True
                # Marca el número de clase como matriculado en esta sección
                numeros_de_clase_matriculados.add(numero_clase)

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
        error_msg = f"Código de Alumno {codigo} - No se encontraron secciones disponibles para la matricula."
        print(error_msg)
        errores.append((codigo, error_msg))
        click_home_button()
        time.sleep(4)