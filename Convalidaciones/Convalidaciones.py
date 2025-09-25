from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time

# Solicitar al usuario que ingrese el número de Ciclo Lvo.
ciclo_lvo = int(input("Ingrese el número de Ciclo Lvo.: "))

# Leer los códigos de alumno desde un archivo Excel
df = pd.read_excel('C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\Convalidaciones\\Convalidaciones.xlsx', sheet_name='Hoja1')
codigos_alumno = df['IdAlumno'].tolist()

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
enlace_inscripcion = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, "Créditos de Curso - Manual")))
enlace_inscripcion.click()

#####################################################
# INGRESO A CONVALIDACIÓN

# Cambiar el contexto al iframe "ptifrmtgtframe"
driver.switch_to.frame("ptifrmtgtframe")

# Esperar hasta que el enlace "Añadir un Valor" esté presente y hacer clic
anadir_valor = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "ICTAB_1"))
)
anadir_valor.click()

time.sleep(2)

# Lista para almacenar los resultados de convalidación
convalidaciones_resultado = []

try:
    for index, row in df.iterrows():
        codigo = row['IdAlumno']
        programa = row['Prog Acad']
        plan = row['Plan Acad']
        anio = row['AÑO']
        materia = row['MATERIA']
        numero_curso = row['N°CURSO']
        curso_entrante = row['CURSO ENTRANTE']
        creditos = row['CREDITOS']
        nota = row['NOTA']
        id_curso = row["IdCurso"]
        creditos_trnsfr = row["Créditos"]
        calificacion = row["Calificación"]

        # Esperar y llenar el campo del código de alumno
        input_codigo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "TRNSFR_CRS_SRCH_EMPLID"))
        )
        input_codigo.clear()
        input_codigo.send_keys(codigo)

        # Esperar y seleccionar "Profesional SISE"
        select_career = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "TRNSFR_CRS_SRCH_ACAD_CAREER"))
        )
        Select(select_career).select_by_visible_text("Profesional SISE")

        # Esperar y llenar el campo "SISE"
        input_institution = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "TRNSFR_CRS_SRCH_INSTITUTION"))
        )
        input_institution.clear()
        input_institution.send_keys("SISE")

        # Hacer clic en el botón "Añadir"
        boton_añadir = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "#ICSearch"))
        )
        boton_añadir.click()

        time.sleep(2)

        # Intentar detectar el mensaje de que el valor ya existe
        try:
            mensaje_alerta = WebDriverWait(driver, 2).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//div[@role='alert' and contains(text(), 'El valor que ha intentado añadir ya existe')]")
                )
            )
            print("Ya existe una convalidación previa. Entrando a editar...")

            # Hacer clic en el enlace de resultado para editar
            enlace_editar = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "SEARCH_RESULT1"))
            )
            enlace_editar.click()

            time.sleep(8)

            #####################################################
            # Entrada a Curso Convalidado existente

            # Aqui poner lógica específica para editar convalidación



        except TimeoutException:
            print("No se detectó convalidación previa. Procediendo con el ingreso nuevo.")
            #####################################################
            # Entrada a Nuevo Curso Convalidado

            # Campo de programa académico
            input_programa = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "TRNS_CRSE_SCH_ACAD_PROG$0"))
            )
            input_programa.clear()
            input_programa.send_keys(programa)

            # Simular TAB para que el sistema dispare el evento onchange
            input_programa.send_keys(Keys.TAB)

            time.sleep(1.5)

            # Esperar a que el campo de plan esté listo y vacío
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "TRNS_CRSE_SCH_ACAD_PLAN$0"))
            )

            input_plan = driver.find_element(By.ID, "TRNS_CRSE_SCH_ACAD_PLAN$0")
            input_plan.send_keys(plan)

            input_plan.send_keys(Keys.TAB)

            time.sleep(1)

            input_origen = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "TRNS_CRSE_SCH_SRC_ORG_NAME$1$$0"))
            )
            input_origen.clear()
            input_origen.send_keys("CAMBIO INTERNO - SISE")

            # Campo de ciclo lectivo (term)
            input_term = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "TRNS_CRSE_TERM_ARTICULATION_TERM$0"))
            )
            input_term.clear()
            input_term.send_keys(ciclo_lvo)

            time.sleep(1.5)

            #####################################################
            # CURSO ENTRANTE

            # Campo año

            input_anio = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "TRNS_CRSE_DTL_TERM_YEAR$0"))
            )
            input_anio.clear()
            input_anio.send_keys(str(anio))

            # Campo MATERIA
            input_materia = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "TRNS_CRSE_DTL_SCHOOL_SUBJECT$0"))
            )
            input_materia.clear()
            input_materia.send_keys(materia)

            # Campo N°CURSO
            input_numero_curso = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "TRNS_CRSE_DTL_SCHOOL_CRSE_NBR$0"))
            )
            input_numero_curso.clear()
            input_numero_curso.send_keys(str(numero_curso))

            # Campo CURSO ENTRANTE (Descripción)
            input_curso_entrante = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "TRNS_CRSE_DTL_DESCR$0"))
            )
            input_curso_entrante.clear()
            input_curso_entrante.send_keys(curso_entrante)

            # Campo CRÉDITOS
            input_creditos = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "TRNS_CRSE_DTL_UNT_TAKEN$0"))
            )
            input_creditos.clear()
            input_creditos.send_keys(str(creditos))

            # Campo NOTA
            input_nota = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "TRNS_CRSE_DTL_CRSE_GRADE_INPUT$0"))
            )
            input_nota.clear()
            input_nota.send_keys(str(nota))

            time.sleep(1.5)

            #####################################################
            # CURSO SALIENTE

            # Ingresar IdCurso
            input_id_curso = driver.find_element(By.ID, "TRNS_CRSE_DTL_CRSE_ID$0")
            input_id_curso.clear()
            input_id_curso.send_keys(str(id_curso))
            input_id_curso.send_keys(Keys.TAB)

            # Ingresar valor de la columna Créditos
            input_creditos_trnsfr = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "TRNS_CRSE_DTL_UNT_TRNSFR$0"))
            )
            input_creditos_trnsfr.clear()
            input_creditos_trnsfr.send_keys(str(creditos_trnsfr))

            # Ingresar valor de la columna Calificación
            input_calificacion = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "TRNS_CRSE_DTL_CRSE_GRADE_OFF$0"))
            )
            input_calificacion.clear()
            input_calificacion.send_keys(str(calificacion))

            input_calificacion.send_keys(Keys.TAB)
            input_calificacion.send_keys(Keys.TAB)

            # Espera 1.5 segundos después de ingresar IdCurso
            time.sleep(1.5)

            #####################################################
            # Créditos de Curso p/Ciclo Lvo

            # Esperar que el enlace esté presente y clickeable
            tab_creditos = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ICTAB_1"))
            )
            tab_creditos.click()

            time.sleep(1.5)

            # Esperar a que el botón 'Publicar' esté disponible y hacer clic
            btn_publicar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "DERIVED_TRCR_POST_PB$0"))
            )
            btn_publicar.click()
            time.sleep(1.5)

            # Esperar el mensaje de "Guardado"
            try:
                WebDriverWait(driver, 10).until(
                    EC.text_to_be_present_in_element((By.ID, "ptStatusText_win0"), "Guardado")
                )
                print("✅ Datos guardados correctamente.")
                convalidaciones_resultado.append((codigo, "Convalidación correcta"))
            except TimeoutException:
                print("⚠️ No se detectó el mensaje de 'Guardado'. Verificar si se publicó correctamente.")
                convalidaciones_resultado.append((codigo, "Fallo al guardar"))

            #######################################################
            # Pasar al siguiente código

            # Cambiar al frame principal (salir del iframe)
            driver.switch_to.default_content()

            # Esperar hasta que se cargue el enlace "Créditos de Curso - Manual"
            enlace_inscripcion = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.LINK_TEXT, "Créditos de Curso - Manual"))
            )
            enlace_inscripcion.click()

            # Volver al iframe "ptifrmtgtframe" después de hacer clic
            driver.switch_to.frame("ptifrmtgtframe")

            # Esperar y hacer clic en "Añadir un Valor"
            anadir_valor = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ICTAB_1"))
            )
            anadir_valor.click()

            time.sleep(2)

except Exception as e:
    print(f"Se ha producido un error general durante la ejecución: {e}")

finally:
    # Convertir a DataFrame y guardar resultados
    convalidaciones_df = pd.DataFrame(convalidaciones_resultado, columns=['Codigo', 'Convalidación'])

    intentos = 0
    exito = False

    while intentos < 3 and not exito:
        try:
            convalidaciones_df.to_excel('convalidaciones_resultado.xlsx', index=False, engine='openpyxl')
            print("✅ Archivo 'convalidaciones_resultado.xlsx' creado con éxito.")
            exito = True
        except Exception as e:
            intentos += 1
            print(f"❌ Error al guardar el archivo Excel: Intento {intentos} de 3. Error: {e}")
            if intentos < 3:
                respuesta = input("¿Deseas intentar guardar el archivo nuevamente? (si/no): ").lower()
                if respuesta != 'si':
                    print("No se volverá a intentar crear el archivo Excel.")
                    break
            else:
                print("Se alcanzó el número máximo de intentos. No se pudo crear el archivo Excel.")

    driver.quit()
