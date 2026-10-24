import pandas as pd
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from datetime import datetime, timedelta
from robot_matricula.core.navegador import crear_navegador_chrome


def convertir_a_hora(valor):
    """Convierte un valor en formato 'HH.MM.SS.mmmmmm' a 'HH:MM'."""
    partes = valor.split(".")
    if len(partes) >= 2:
        return f"{partes[0]}:{partes[1]}"
    return valor  # Devuelve el valor original si no tiene el formato esperado


def si_modal(new_driver):
    try:
        # Volver al contexto principal para asegurarse de que podemos interactuar con el modal
        new_driver.switch_to.default_content()

        time.sleep(1)

        # Verificar si aparece un modal con el ID dinámico "ptModTable_"
        modal = WebDriverWait(new_driver, 3).until(
            EC.presence_of_element_located((By.XPATH, "//div[starts-with(@id, 'ptModTable_')]"))
        )

        #print("Modal encontrado.")

        # Obtener el contenido del modal
        modal_content = modal.find_element(By.XPATH, ".//div[@id='alertmsg']").text
        print(f"Contenido del modal: {modal_content}")

        # Determinar acción en función del contenido del modal
        if "guardar" in modal_content.lower():
            # Presionar botón "No"
            no_btn = WebDriverWait(new_driver, 3).until(
                EC.element_to_be_clickable((By.ID, "ptpopupmsgbtn2"))
            )
            no_btn.click()
            print("Botón 'No' clickeado.")
        else:
            # Presionar botón "Aceptar"
            aceptar_btn = WebDriverWait(new_driver, 3).until(
                EC.element_to_be_clickable((By.ID, "#ICOK"))
            )
            aceptar_btn.click()
            print("Botón 'Aceptar' clickeado.")

        # Volver al iframe principal para otras operaciones
        new_driver.switch_to.frame("ptifrmtgtframe")

    except TimeoutException:
        # Si el modal no aparece, continuar sin problemas
        print("No se encontró modal. Continuando...")


def volver_al_inicio(new_driver):

    try:
        # Volver al contexto principal
        new_driver.switch_to.default_content()
        time.sleep(1)

        # Buscar y hacer clic en el enlace "Programa de Curso Nuevo"
        enlace_inscripcion = WebDriverWait(new_driver, 10).until(
            EC.presence_of_element_located((By.LINK_TEXT, "Programa de Curso Nuevo"))
        )
        enlace_inscripcion.click()

        time.sleep(1)

        # Intentar encontrar el botón "No" y hacer clic
        try:

            # Esperar a que el botón "No" esté clickeable usando su ID correcto
            no_btn = WebDriverWait(new_driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ptpopupmsgbtn2"))
            )
            no_btn.click()

        except TimeoutException:
            # El botón "No" no apareció
            print("Botón 'No' no encontrado. Continuando con el flujo.")

        # Volver al iframe principal para otras operaciones
        new_driver.switch_to.frame("ptifrmtgtframe")

    except Exception as e:
        print(f"Error durante la ejecución de los pasos del programa: {str(e)}")


# Solicitar al usuario que ingrese el número de Ciclo Lvo.
# ciclo_lvo = int(input("Ingrese el número de Ciclo Lvo.: "))
ciclo_lvo = 3133

# Ruta del archivo Excel
excel_file = 'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\En Construcion\\ProgNClasesVirtPresencial.xlsx'

# Leer las hojas del archivo Excel
df_hoja1 = pd.read_excel(excel_file, sheet_name='Hoja1')
df_hoja2 = pd.read_excel(excel_file, sheet_name='Hoja2')

# Leer el archivo Excel
df = pd.read_excel(excel_file)

# Crear el navegador usando la función compartida
new_driver = crear_navegador_chrome()

# Lista para almacenar los resultados
resultados = []

try:
    # Abre la página web
    new_driver.get('https://academico.sise.edu.pe/psp/CS90PRD/?cmd=login')

    # REGISTRO DE USUARIO
    input_usuario = new_driver.find_element(By.ID, "userid")
    input_usuario.clear()
    input_usuario.send_keys("ZVERGARAL")

    input_contrasena = new_driver.find_element(By.ID, "pwd")
    input_contrasena.clear()
    input_contrasena.send_keys("Sofia0603")

    boton_conectar = new_driver.find_element(By.CSS_SELECTOR, 'button.btn-danger-sise')
    boton_conectar.click()

    time.sleep(2)

    # INGRESAR A INSCRIPCIÓN RÁPIDA
    enlace_favoritos = WebDriverWait(new_driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'a#fldra_MYFAVORITES')))
    enlace_favoritos.click()

    enlace_inscripcion = WebDriverWait(new_driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Programa de Curso Nuevo")))
    enlace_inscripcion.click()

    time.sleep(1)

    new_driver.switch_to.frame("ptifrmtgtframe")

    for numero_clase in df_hoja1['Nº Clase']:

        try:

            reiniciar_for = False

            print("=====================================")
            print(" ")
            print(f"Procesando clase número: {numero_clase}")

            # Buscar el número de clase en la Hoja2
            fila = df_hoja2[df_hoja2['Nº Clase'] == numero_clase]
            id_curso = fila['ID Curso'].values[0] if not fila.empty else None
            nombre_curso = fila['Nombre Curso'].values[0] if not fila.empty else None
            seccion = fila['Sección'].values[0] if not fila.empty else None

            if id_curso is None:
                print(f"No se encontró la clase {numero_clase} en el archivo Excel.")
                continue

            print(f"El Id del curso es: {id_curso}")

            boton_borrar = new_driver.find_element(By.ID, "#ICClear")
            boton_borrar.click()
            time.sleep(2)

            campo_texto = new_driver.find_element(By.ID, "CRSE_OFFER_SCTY_INSTITUTION")
            campo_texto.clear()
            campo_texto.send_keys("SISE")
            time.sleep(1)

            campo_ciclo = new_driver.find_element(By.ID, "CRSE_OFFER_SCTY_STRM")
            campo_ciclo.clear()
            campo_ciclo.send_keys(f"{ciclo_lvo}")
            time.sleep(1)

            combobox_grado = new_driver.find_element(By.ID, "CRSE_OFFER_SCTY_ACAD_CAREER")
            select_grado = Select(combobox_grado)
            select_grado.select_by_value("PRGS")
            time.sleep(1)

            campo_numero = new_driver.find_element(By.ID, "CRSE_OFFER_SCTY_CRSE_ID")
            campo_numero.clear()
            campo_numero.send_keys(f"{id_curso}")

            boton_buscar = new_driver.find_element(By.ID, "#ICSearch")
            boton_buscar.click()

            time.sleep(2)

            # INGRESO A CLASES POR PROFESOR

            enlace_clases_profesor = WebDriverWait(new_driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ICTAB_1"))
            )
            time.sleep(1)
            enlace_clases_profesor.click()
            # print("Se hizo clic en el enlace 'Clases por Profesor'.")

            # BÚSQUEDA DE NUMERO DE CLASE EN ALERT

            try:

                # SE DEBE TRABAJAR EN VOLVERLO DINÁMICO
                boton_buscar = WebDriverWait(new_driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "$ICField4$hfind$0"))
                )
                time.sleep(2)
                boton_buscar.click()
                # print("Se hizo clic en el botón 'Buscar'.")

                # Esperar a que aparezca el alert
                WebDriverWait(new_driver, 10).until(EC.alert_is_present())
                alert = Alert(new_driver)  # Crear un objeto de alerta

                time.sleep(2)

                # Ingresar el número en el alert si es posible (solo si es un prompt)
                numero_a_ingresar = f"{numero_clase}"
                alert.send_keys(numero_a_ingresar)

                time.sleep(1)
                # print(f"Se ingresó el número '{numero_a_ingresar}' en el alert.")

                # Aceptar el alert
                alert.accept()
                # print("Se aceptó el alert.")

            except TimeoutException:
                print("No se encontró el botón 'Buscar' o el alert no apareció.")
                resultados.append((numero_clase, f'Error procesando el codigo'))
                si_modal(new_driver)
                volver_al_inicio(new_driver)
                continue

            time.sleep(2)

            # COMPROBACIÓN EN "VER TODO" En caso existan numeros de clase con mas días entonces se deberá corregir con correlativos a trabajar

            try:
                # Esperar a que aparezca el enlace "Ver Todo"
                enlace_ver_todo = WebDriverWait(new_driver, 6).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//a[contains(@class, 'PSLEVEL2SCROLLAREAHEADER') and contains(@id, 'hviewall')]")
                    )
                )
                # Si se encuentra el enlace, se considera un problema
                print("Se encontró más espacios en Modelo Reunión.")
                resultados.append((numero_clase, f"Se encontro más espacios en Modelo Reunión."))
                si_modal(new_driver)
                volver_al_inicio(new_driver)
                time.sleep(2)
                continue  # Salta al siguiente número de clase en el bucle principal

            except TimeoutException:
                pass
                # print("El botón 'Ver Todo' no apareció. Continuando sin hacer clic.")

            time.sleep(2)

            # ############### VERIFICACIÓN DE NÚMERO DE CLASE ###############

            try:

                # SE DEBE MEJORAR EL CÓDIGO PARA VER POSIBLES ERRORES CON CURSOS DE VARIOS HORARIOS

                correlativo = 0  # Correlativo inicial

                # Paso 1: Localizar el select según el correlativo
                select_id = f"UCS_DERIVED_SR_UCS_MODO${correlativo}"
                select_element = WebDriverWait(new_driver, 10).until(
                    EC.presence_of_element_located((By.ID, select_id))
                )

                # Paso 2: Verificar el valor seleccionado en el select
                selected_option = select_element.find_element(By.XPATH, "./option[@selected='selected']")
                selected_value = selected_option.get_attribute("value")

                if selected_value != "V":
                    print(f"El curso {numero_clase} ya es presencial o semipresencial.")
                    resultados.append((numero_clase, f"Ya es presencial o semipresencial."))
                    volver_al_inicio(new_driver)
                    time.sleep(2)
                    continue
                else:
                    # print(f"El curso con correlativo {correlativo} está marcado como Virtual. Procediendo...")
                    time.sleep(2)
                    # Paso 3: Hacer clic en "Añadir Fila"
                    add_row_id = f"$ICField19$new$0$${correlativo}"
                    add_row_button = WebDriverWait(new_driver, 10).until(
                        EC.element_to_be_clickable((By.ID, add_row_id))
                    )
                    add_row_button.click()
                    # print(f"Se añadió una nueva fila.")

                    time.sleep(3)

                    # AQUÍ DEBERÍA SER DINÁMICO EN CASO DE VARIOS HORARIOS
                    nuevo_correlativo = correlativo + 1

                    # Paso 6: Hacer clic en "Ver Todo" si es necesario
                    try:
                        ver_todo_button = WebDriverWait(new_driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH,
                                                        "//a[contains(@class, 'PSLEVEL2SCROLLAREAHEADER') and contains(@id, 'hviewall')]"))
                        )
                        ver_todo_button.click()
                        # print("Se hizo clic en 'Ver Todo'.")
                    except TimeoutException:
                        print("No se encontró el botón 'Ver Todo 2'.")
                        resultados.append((numero_clase, f"No se encontró el botón 'Ver Todo 2'."))
                        si_modal(new_driver)
                        volver_al_inicio(new_driver)
                        time.sleep(2)
                        continue

                    time.sleep(2)

            except TimeoutException as e:
                print("Error en procesar VERIFICACION NUMERO DE CLASE")
                resultados.append((numero_clase, f"Error en procesar VERIFICACION NUMERO DE CLASE"))
                si_modal(new_driver)
                volver_al_inicio(new_driver)
                time.sleep(2)
                continue

            # ############### COPIA DE VALORES  ###############

            # ####### IDInstalacion AQUI SE EN LA HOJA1 SE OBTENDRÁ EL ID ##########

            # Localizar el valor del campo en la primera tabla (correlativo 0)
            campo_origen = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, f"CLASS_MTG_PAT_FACILITY_ID${correlativo}"))
            )
            valor_a_copiar = campo_origen.get_attribute("value")
            # print(f"Valor a copiar: {valor_a_copiar}")

            # Localizar el campo en la segunda tabla (correlativo 1)
            campo_destino = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, f"CLASS_MTG_PAT_FACILITY_ID${nuevo_correlativo}"))
            )
            campo_destino.clear()  # Limpiar el campo destino por si tiene algún valor
            campo_destino.send_keys(valor_a_copiar)  # Copiar el valor al campo destino
            # print("Valor copiado correctamente en la segunda tabla.")

            time.sleep(1)

            # ############ HORAS ###############

            # Localizar el valor de la hora de inicio en la primera tabla
            campo_hora_inicio_origen = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, "UCS_HORPATE_WRK_UCS_START_HRS$0"))
            )
            hora_inicio_a_copiar = convertir_a_hora(campo_hora_inicio_origen.get_attribute("value"))
            # print(f"Hora de inicio a copiar: {hora_inicio_a_copiar}")

            # Localizar el campo de destino para la hora de inicio (corregido)
            campo_hora_inicio_destino = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, "UCS_HORPATE_WRK_UCS_START_HRS$1"))
            )
            campo_hora_inicio_destino.clear()
            campo_hora_inicio_destino.send_keys(hora_inicio_a_copiar)
            # print("Hora de inicio copiada correctamente.")

            # Localizar el valor de la hora de fin en la primera tabla
            campo_hora_fin_origen = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, "UCS_HORPATE_WRK_UCS_END_HRS$0"))
            )
            hora_fin_a_copiar = convertir_a_hora(campo_hora_fin_origen.get_attribute("value"))
            # print(f"Hora de fin a copiar: {hora_fin_a_copiar}")

            # Localizar el campo de destino para la hora de fin
            campo_hora_fin_destino = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, "UCS_HORPATE_WRK_UCS_END_HRS$1"))
            )
            campo_hora_fin_destino.clear()
            campo_hora_fin_destino.send_keys(hora_fin_a_copiar)
            # print("Hora de fin copiada correctamente.")

            # ############## DÍAS DE LA SEMANA:

            dias_semana = [
                ("CLASS_MTG_PAT_MON", "Lunes"),
                ("CLASS_MTG_PAT_TUES", "Martes"),
                ("CLASS_MTG_PAT_WED", "Miércoles"),
                ("CLASS_MTG_PAT_THURS", "Jueves"),
                ("CLASS_MTG_PAT_FRI", "Viernes"),
                ("CLASS_MTG_PAT_SAT", "Sábado"),
                ("CLASS_MTG_PAT_SUN", "Domingo"),
            ]

            time.sleep(1)

            try:
                # Identificar qué día está marcado en correlativo 0
                dia_marcado = None
                for dia_id, nombre_dia in dias_semana:
                    checkbox_origen = WebDriverWait(new_driver, 10).until(
                        EC.presence_of_element_located((By.ID, f"{dia_id}$0"))
                    )
                    if checkbox_origen.get_attribute("checked") == "true":
                        dia_marcado = dia_id
                        # print(f"Día marcado en correlativo 0: {nombre_dia}")
                        break

                if dia_marcado:
                    # Marcar el día correspondiente en correlativo 1
                    checkbox_destino = WebDriverWait(new_driver, 10).until(
                        EC.presence_of_element_located((By.ID, f"{dia_marcado}$1"))
                    )
                    if not checkbox_destino.is_selected():
                        checkbox_destino.click()
                        # print(f"Día marcado correctamente en correlativo 1: {nombre_dia}")
                else:
                    print("No se encontró ningún día marcado en correlativo Virtual.")

            except TimeoutException:
                print("No se pudo localizar alguno de los checkboxes.")
                resultados.append((numero_clase, f"No se pudo localizar alguno de los checkboxes."))
                si_modal(new_driver)
                volver_al_inicio(new_driver)
                time.sleep(2)
                continue



            # ############### FECHAS INICIO Y FIN ###############

            time.sleep(2)

            # PARTE 1: Cambiar el estado a "Cerrado" en correlativo 0
            estado_selector = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, "UCS_HORPATE_WRK_UCS_SCHEDULESTATUS$0"))
            )
            # Encontrar la opción dentro del select
            opcion_cerrado = estado_selector.find_element(By.XPATH, "//option[@value='2']")

            time.sleep(1)

            # Hacer clic en la opción
            opcion_cerrado.click()
            # print("Opción 'Cerrado' seleccionada mediante clic.")

            # Disparar el evento onchange manualmente, por si aún es necesario
            new_driver.execute_script(
                "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", estado_selector
            )
            # print("Evento onchange disparado manualmente después del clic.")

            time.sleep(2)

            # PARTE 2: Obtener la fecha fin general del correlativo 0
            fecha_fin_input = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, "CLASS_MTG_PAT_END_DT$0"))
            )
            fecha_fin_general = fecha_fin_input.get_attribute("value")
            print(f"Fecha fin general obtenida: {fecha_fin_general}")

            # PARTE 3: Obtener la fecha de inicio del correlativo 0 y calcular la nueva fecha
            fecha_inicio_input = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, "CLASS_MTG_PAT_START_DT$0"))
            )
            fecha_inicio_str = fecha_inicio_input.get_attribute("value")
            # print(f"Fecha de inicio obtenida: {fecha_inicio_str}")

            # Convertir fecha_inicio_str a formato de fecha y calcular la nueva fecha fin
            fecha_inicio = datetime.strptime(fecha_inicio_str, "%m/%d/%Y")
            nueva_fecha_fin = fecha_inicio + timedelta(days=7)  # Acá se agrega 7 días
            nueva_fecha_fin_str = nueva_fecha_fin.strftime("%m/%d/%Y")
            # print(f"Nueva fecha calculada para correlativo 0: {nueva_fecha_fin_str}")

            time.sleep(1)

            # PARTE 4: Termino de fecha de la primera semana virtual.

            # Obtener la fecha de inicio del campo de entrada
            fecha_inicio_input = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, "CLASS_MTG_PAT_START_DT$0"))
            )
            fecha_inicio_str = fecha_inicio_input.get_attribute("value")
            # print(f"Fecha de inicio obtenida: {fecha_inicio_str}")

            # Calcular la nueva fecha fin (sumando 6 días)
            fecha_inicio_dt = datetime.strptime(fecha_inicio_str, "%m/%d/%Y")  # Convertir a objeto datetime
            fecha_fin_calculada = fecha_inicio_dt + timedelta(days=6)  # Sumar 6 días
            fecha_fin_calculada_str = fecha_fin_calculada.strftime("%m/%d/%Y")  # Convertir de nuevo a cadena
            # print(f"Nueva fecha calculada: {fecha_fin_calculada_str}")

            time.sleep(1)

            # Escribir la nueva fecha fin en el campo correspondiente
            fecha_fin_input = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, "CLASS_MTG_PAT_END_DT$0"))
            )

            # Limpiar el campo y escribir la nueva fecha
            fecha_fin_input.clear()  # Limpiar cualquier valor previo
            fecha_fin_input.send_keys(fecha_fin_calculada_str)  # Escribir la nueva fecha
            # print("Nueva fecha fin establecida en el campo.")

            # PARTE 5: Asignar la nueva fecha calculada en el correlativo 1

            # Localizar el campo de fecha de inicio del correlativo 1 y asignar la nueva fecha calculada
            fecha_inicio_correlativo_1_input = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, "CLASS_MTG_PAT_START_DT$1"))
            )
            fecha_inicio_correlativo_1_input.clear()  # Limpiar el campo de entrada antes de asignar
            fecha_inicio_correlativo_1_input.send_keys(nueva_fecha_fin_str)  # Asignar la nueva fecha calculada
            # print(f"Nueva fecha de inicio asignada en correlativo 1: {nueva_fecha_fin_str}")

            # Localizar el campo de fecha fin del correlativo 1 y asignar la fecha fin general
            fecha_fin_correlativo_1_input = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, "CLASS_MTG_PAT_END_DT$1"))
            )
            fecha_fin_correlativo_1_input.clear()  # Limpiar el campo de entrada antes de asignar
            print(fecha_fin_general)
            fecha_fin_correlativo_1_input.send_keys(fecha_fin_general)  # Asignar la fecha fin general
            # print(f"Fecha fin general asignada en correlativo 1: {fecha_fin_general}")

            time.sleep(1)

            # PARTE 6: Seleccionar "Presencial" en el combobox del correlativo 1 y Cerrar correlativo 0 y 1
            select_modo_input = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, "UCS_DERIVED_SR_UCS_MODO$1"))
            )
            select_modo = Select(select_modo_input)  # Utilizar Select para interactuar con el combobox
            select_modo.select_by_value("P")  # Seleccionar "Presencial" (value="P")
            # print("Opción 'Presencial' seleccionada en el combobox de correlativo 1.")

            time.sleep(2)

            # Activo a correlativo 0

            # Obtener el elemento select
            estado_selector = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, "UCS_HORPATE_WRK_UCS_SCHEDULESTATUS$0"))
            )

            # Cambiar el valor a "Activo" (value="1") usando JavaScript
            new_driver.execute_script("""
                var select = arguments[0];
                select.value = '1';  // Cambiar el valor a 'Activo'
                select.dispatchEvent(new Event('change', { bubbles: true }));  // Disparar el evento onchange
            """, estado_selector)

            # print("Estado cambiado a 'Activo' mediante JavaScript.")

            time.sleep(1)

            # ############### CAMBIO Profesores p/Modelo Reunión VIRTUAL ###############

            # Obtener el elemento select de Virtual
            estado_selector = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, "UCS_DOCINST_WRK_UCS_SCHEDULESTATUS$0"))
            )

            # Cambiar el valor a "Cerrado" (value="2") usando JavaScript
            new_driver.execute_script("""
                var select = arguments[0];
                select.value = '2';  // Cambiar el valor a 'Cerrado'
                select.dispatchEvent(new Event('change', { bubbles: true }));  // Disparar el evento onchange
            """, estado_selector)

            # print("Estado cambiado a 'Cerrado' mediante JavaScript.")

            time.sleep(2)

            # Localizar el contenedor del campo de texto
            contenedor_fecha = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, "win0divUCS_DOCINST_WRK_END_DT$0"))
            )

            time.sleep(1)

            # Localizar el campo de texto dentro del contenedor
            fecha_fin_input = contenedor_fecha.find_element(By.ID, "UCS_DOCINST_WRK_END_DT$0")

            # Limpiar el contenido del campo de texto
            fecha_fin_input.clear()

            time.sleep(1)

            # Ingresar el nuevo valor de fecha
            fecha_fin_input.send_keys(fecha_fin_calculada_str)

            # print(f"Valor '{fecha_fin_calculada_str}' introducido correctamente en el campo de texto.")

            time.sleep(2)

            # SIMULACIÓN A PONER A PRUEBA

            try:
                # Esperar que el botón "Cerrar Horario" esté presente
                cerrar_horario_btn = WebDriverWait(new_driver, 10).until(
                    EC.presence_of_element_located((By.ID, "UCS_BTNINSTRWRK_UCS_BTN_CERRAR_HOR$0"))
                )

                # Hacer clic en el botón
                cerrar_horario_btn.click()
                # print("Botón 'Cerrar Horario' clickeado.")

                time.sleep(1)

                # Cambiar al contexto principal para verificar si aparece un modal
                new_driver.switch_to.default_content()
                # print("Cambiado al contexto principal para verificar modales.")

                while True:
                    try:
                        # Esperar a que aparezca un modal
                        modal = WebDriverWait(new_driver, 3).until(
                            EC.presence_of_element_located((By.XPATH, "//div[contains(@id, 'ptModTable_')]"))
                        )
                        # print(f"Modal detectado con ID: {modal.get_attribute('id')}")

                        # Leer el contenido del modal
                        modal_content = modal.find_element(By.ID, "alertmsg").text
                        # print(f"Contenido del modal: {modal_content}")

                        time.sleep(1)

                        # Verificar el contenido del mensaje del modal
                        if "Eliminara las Marcaciones del Docente" in modal_content:
                            print("Error detectado: Eliminara las Marcaciones del Docente")
                            time.sleep(1)
                            aceptar_btn = modal.find_element(By.XPATH, "//input[@id='#ICOK']")
                            aceptar_btn.click()
                            # print("Botón 'Aceptar' del modal clickeado después de detectar error.")

                            time.sleep(2)

                            resultados.append((numero_clase,f"Eliminara las Marcaciones del Docente."))
                            volver_al_inicio(new_driver)
                            reiniciar_for = True
                            break

                        # Si no es un error, buscar y hacer clic en el botón "Aceptar"
                        aceptar_btn = modal.find_element(By.XPATH, "//input[@id='#ICOK']")
                        aceptar_btn.click()
                        # print("Botón 'Aceptar' del modal clickeado.")
                        time.sleep(2)

                    except TimeoutException:
                        # No hay más modales, continuar con el flujo
                        # print("No se detectaron más modales. Continuando con el flujo.")
                        break

                    if reiniciar_for:
                        break

                if reiniciar_for:
                    continue  # Reiniciamos el for

            except Exception as e:
                print(f"Error al intentar cerrar el horario Profesores p/Modelo Reunión VIRTUAL ")
                resultados.append((numero_clase, f"Error al intentar cerrar el horario Profesores p/Modelo Reunión VIRTUAL."))
                si_modal(new_driver)
                volver_al_inicio(new_driver)
                time.sleep(2)
                continue

            # ############### CAMBIO Profesores p/Modelo Reunión PRESENCIAL ###############

            # Obtener el elemento select de Virtual
            estado_selector = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, "UCS_DOCINST_WRK_UCS_SCHEDULESTATUS$1"))
            )

            # Cambiar el valor a "Cerrado" (value="2") usando JavaScript
            new_driver.execute_script("""
                                            var select = arguments[0];
                                            select.value = '2';  // Cambiar el valor a 'Cerrado'
                                            select.dispatchEvent(new Event('change', { bubbles: true }));  // Disparar el evento onchange
                                        """, estado_selector)

            print("Estado cambiado a 'Cerrado' mediante JavaScript.")

            time.sleep(2)

            # Localizar el contenedor y el campo de texto dentro de él
            contenedor_fecha = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, "win0divUCS_DOCINST_WRK_END_DT$1"))
            )
            fecha_fin_input = WebDriverWait(contenedor_fecha, 10).until(
                EC.presence_of_element_located((By.ID, "UCS_DOCINST_WRK_END_DT$1"))
            )

            # Limpiar el contenido del campo de texto
            #fecha_fin_input.clear()
            print("Campo de fecha limpiado ya no.")

            time.sleep(1)

            # Esperar a que el campo de texto esté presente y sea interactuable
            campo_fecha = WebDriverWait(new_driver, 10).until(
                EC.presence_of_element_located((By.ID, "UCS_DOCINST_WRK_START_DT$1"))
            )

            # Limpiar cualquier texto preexistente en el campo
            campo_fecha.clear()

            # Ingresar el nuevo valor en el campo
            campo_fecha.send_keys(nueva_fecha_fin_str)

            print(f"Se ingresó la fecha '{nueva_fecha_fin_str}' en el campo de texto.")

            print(f"fecha fin general {fecha_fin_general}")
            time.sleep(1)

            # Ingresar el nuevo valor de fecha
            fecha_fin_input.send_keys(fecha_fin_general)

            print(f"Valor '{fecha_fin_general}' introducido correctamente en el campo de texto Presencial.")

            time.sleep(3)

            # SIMULACIÓN A PONER A PRUEBA

            try:
                # Esperar que el botón "Cerrar Horario" esté presente
                cerrar_horario_btn = WebDriverWait(new_driver, 10).until(
                    EC.presence_of_element_located((By.ID, "UCS_BTNINSTRWRK_UCS_BTN_CERRAR_HOR$1"))
                )

                # Hacer clic en el botón
                cerrar_horario_btn.click()
                print("Botón 'Cerrar Horario' clickeado.")

                time.sleep(1)

                # Cambiar al contexto principal para verificar si aparece un modal
                new_driver.switch_to.default_content()
                print("Cambiado al contexto principal para verificar modales.")

                while True:
                    try:
                        # Esperar a que aparezca un modal
                        modal = WebDriverWait(new_driver, 3).until(
                            EC.presence_of_element_located((By.XPATH, "//div[contains(@id, 'ptModTable_')]"))
                        )
                        print(f"Modal detectado con ID: {modal.get_attribute('id')}")

                        # Leer el contenido del modal
                        modal_content = modal.find_element(By.ID, "alertmsg").text
                        print(f"Contenido del modal: {modal_content}")

                        # Verificar el contenido del mensaje del modal
                        if "Eliminara las Marcaciones del Docente" in modal_content:
                            print("Error detectado: Eliminara las Marcaciones del Docente")
                            time.sleep(1)
                            aceptar_btn = modal.find_element(By.XPATH, "//input[@id='#ICOK']")
                            aceptar_btn.click()
                            print("Botón 'Aceptar' del modal clickeado después de detectar error.")

                            time.sleep(2)
                            volver_al_inicio(new_driver)
                            reiniciar_for = True
                            break

                        # Si no es un error, buscar y hacer clic en el botón "Aceptar"
                        aceptar_btn = modal.find_element(By.XPATH, "//input[@id='#ICOK']")
                        aceptar_btn.click()
                        print("Botón 'Aceptar' del modal clickeado.")

                        time.sleep(2)

                    except TimeoutException:
                        # No hay más modales, continuar con el flujo
                        print("No se detectaron más modales. Continuando con el flujo.")
                        break

                    if reiniciar_for:
                        break

                if reiniciar_for:
                    continue  # Reiniciamos el for
            except Exception as e:
                print(f"Error al intentar cerrar el horario Profesores p/Modelo Reunión PRESENCIAL")
                resultados.append(
                    (numero_clase, f"Error al intentar cerrar el horario Profesores p/Modelo Reunión PRESENCIAL."))
                si_modal(new_driver)
                volver_al_inicio(new_driver)
                time.sleep(2)
                continue

            time.sleep(1)

            # Esperar hasta que el botón "Guardar" esté presente y clickeable
            guardar_btn = WebDriverWait(new_driver, 10).until(
                EC.element_to_be_clickable((By.ID, "#ICSave"))
            )
            # Hacer clic en el botón
            guardar_btn.click()
            print("Botón 'Guardar' clickeado correctamente.")

            time.sleep(1)

            print(f"{numero_clase} - Procesado correctamente")
            resultados.append((numero_clase, f'Procesado correctamente'))

            time.sleep(2)

            # Hacer clic en el botón "Volver a Buscar"
            boton_volver_a_buscar = new_driver.find_element(By.ID, "#ICList")
            boton_volver_a_buscar.click()
            time.sleep(3)

            # Hacer clic en el botón "Borrar"
            boton_borrar = new_driver.find_element(By.ID, "#ICClear")
            boton_borrar.click()
            time.sleep(2)

        except Exception:
            print(f'{numero_clase} : Error procesando el código')
            resultados.append((numero_clase, f'Error procesando el codigo'))
            si_modal(new_driver)
            volver_al_inicio(new_driver)
            continue

except Exception as e:
    print(f"Se ha producido un error")

finally:
    # Convertir la lista de resultados en un DataFrame
    resultados_df = pd.DataFrame(resultados, columns=['Codigo', 'Resultado'])

    # Intentar guardar el DataFrame en un archivo CSV con reintentos
    for attempt in range(3):
        try:
            resultados_df.to_csv('resultados_clases_virt_pres.csv', index=False)
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
    new_driver.quit()
