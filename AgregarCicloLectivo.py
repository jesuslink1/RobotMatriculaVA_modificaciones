import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Solicitar al usuario que ingrese el número de Ciclo Lvo.
ciclo_lvo = int(input("Ingrese el número de Ciclo Lvo a agregar: "))

# Lista para almacenar los errores
errores = []

# Leer los códigos de alumno desde un archivo Excel
df = pd.read_excel('C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\CodigosEliminarCicloLectivo.xlsx', sheet_name='Hoja1')
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
# INGRESAR A CICLO LECTIVO

# Esperar hasta que se cargue el enlace "Mis Favoritos"
enlace_favoritos = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a#fldra_MYFAVORITES')))
enlace_favoritos.click()

# CAMBIAR AL NUEVO ENLACE DE ACTIVACIÓN DE CICLO LECTIVO
enlace_activacion_ciclo_lectivo = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'li#crefli_fav_HC_STDNT_ACTIVATION_GBL > a'))
)
enlace_activacion_ciclo_lectivo.click()

time.sleep(2)

#####################################################
#EN ACTIVACION DE CICLO LECTIVO

# Cambiar el contexto al iframe "ptifrmtgtframe"
driver.switch_to.frame("ptifrmtgtframe")


#####################################################
# ITERAR SOBRE LOS CODIGOS DE ALUMNO

for codigo in codigos_alumno:
    try:
        # Esperar hasta que el campo de entrada esté disponible y limpiarlo
        campo_codigo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="STDNT_SRCH_EMPLID"]'))
        )
        campo_codigo.clear()
        campo_codigo.send_keys(codigo)

        # Esperar hasta que el botón de búsqueda esté disponible y hacer clic en él
        boton_buscar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, '#ICSearch'))
        )
        boton_buscar.click()

        time.sleep(2)  # Ajustar el tiempo según sea necesario

        try:
            boton_ver_todo = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable((By.ID, '$ICField4$hviewall$0'))
            )
            boton_ver_todo.click()
            time.sleep(2)
        except TimeoutException:
            pass

        """numero_final = None
        for i in range(20):
            try:
                elemento = driver.find_element(By.ID, f'PSXLATITEM_XLATLONGNAME${i}')
                if "Profesional SISE" in elemento.text:
                    numero_final = i
                    break
            except NoSuchElementException:
                continue

        if numero_final is not None:
            try:
                boton_profesional_sise = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.ID, f'$ICField9$hviewall${numero_final}'))
                )
                boton_profesional_sise.click()
                time.sleep(3)
            except TimeoutException:
                pass
        else:
            print(f'{codigo} : Elemento "Profesional SISE" no encontrado para el código.')
            errores.append((codigo, f'Posible código de extensión'))"""

        # Buscar el campo del ciclo y eliminar la fila correspondiente
        ciclo_existe = False

        # Verificar si el ciclo ya está registrado
        for i in range(20):
            try:
                campo_ciclo = driver.find_element(By.ID, f'STDNT_CAR_TERM_STRM${i}')
                valor = campo_ciclo.get_attribute("value")
                if valor == str(ciclo_lvo):
                    print(f'{codigo} : ⚠️ Ciclo lectivo {ciclo_lvo} ya existe. No se agregará nuevamente.')
                    errores.append((codigo, f'Ciclo {ciclo_lvo} ya existe'))
                    ciclo_existe = True
                    break
            except NoSuchElementException:
                continue

        # ✅ Ahora este bloque solo se ejecuta una vez si el ciclo no existe
        if not ciclo_existe:
            try:
                time.sleep(2)
                # Verificar si el campo STDNT_CAR_TERM_STRM$0 está vacío
                try:
                    campo_primero = driver.find_element(By.ID, "STDNT_CAR_TERM_STRM$0")
                    valor_primero = campo_primero.get_attribute("value")
                    if valor_primero == "":
                        print("⚠️ El campo está vacío")
                        valor_primero = None
                    else:
                        print(valor_primero)
                        print("✅ El campo tiene valor")
                except NoSuchElementException:
                    valor_primero = None
                    print("❌ No existe el campo")

                # Si tiene valor, agregamos una nueva fila
                if valor_primero is not None:
                    boton_agregar = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "$ICField9$new$0$$0"))
                    )
                    boton_agregar.click()
                    time.sleep(2)
                    print("aqui 1")

                # Si no tiene valor, ingresamos los datos en la fila actual
                if valor_primero is None:
                    campo_institucion = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.ID, "STDNT_CAR_TERM_INSTITUTION$0"))
                    )
                    campo_institucion.clear()
                    campo_institucion.send_keys("SISE")
                    time.sleep(0.5)
                    campo_institucion.send_keys(Keys.TAB)
                    time.sleep(1)

                # Ingresar el ciclo
                campo_nuevo_ciclo = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, "STDNT_CAR_TERM_STRM$0"))
                )
                campo_nuevo_ciclo.clear()
                campo_nuevo_ciclo.send_keys(str(ciclo_lvo))
                time.sleep(1)
                campo_nuevo_ciclo.send_keys(Keys.TAB)
                time.sleep(2)

                print(f'{codigo} : ✅ Ciclo lectivo {ciclo_lvo} agregado.')
                errores.append((codigo, f'Ciclo {ciclo_lvo} agregado'))

            except Exception as e:
                print(f'{codigo} : ❌ Error al agregar ciclo {ciclo_lvo} → {e}')
                errores.append((codigo, f'❌ Error al agregar ciclo {ciclo_lvo}: {e}'))

        # Hacer clic en el botón Guardar
        boton_guardar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, '#ICSave'))
        )
        boton_guardar.click()

        time.sleep(3)

        # Hacer clic en el botón Volver a Buscar
        boton_volver_a_buscar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, '#ICList'))
        )
        boton_volver_a_buscar.click()

        time.sleep(2)


    except TimeoutException as e:
        print(f'{codigo} : Error procesando el código')
        errores.append((codigo, f'TimeoutException'))
        continue
    except NoSuchElementException as e:
        print(f'{codigo} : Error procesando el código')
        errores.append((codigo, f'NoSuchElementException'))
        continue


######################################################################

# Intentar guardar el DataFrame en un archivo CSV con reintentos
for attempt in range(3):
    try:
        # Crear un DataFrame con los errores capturados
        df_reporte = pd.DataFrame(errores, columns=['Código', 'Motivo'])

        # Escribir el DataFrame en un archivo CSV
        df_reporte.to_csv('reporteCicloLectivo.csv', index=False)
        print("Archivo CSV guardado correctamente.")
        break  # Salir del bucle si el archivo se guarda correctamente

    except Exception as e:
        print(f"Error al guardar el archivo CSV: {str(e)}")
        if attempt < 2:  # Si no es el último intento
            retry = input("¿Desea intentarlo nuevamente? (Sí/No): ").strip().lower()
            if retry in ('sí', 'si', 's'):
                print(f"Reintentando... (Intento {attempt + 2}/3)")
                continue
            else:
                print("No se intentará guardar el archivo CSV nuevamente.")
                break
        else:
            print("No se pudo guardar el archivo CSV después de 3 intentos.")

# Detener el servicio del controlador de Chrome
chrome_service.stop()