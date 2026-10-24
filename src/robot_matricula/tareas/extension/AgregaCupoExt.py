from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from robot_matricula.core.navegador import crear_navegador_chrome

# Número de Ciclo Lvo
ciclo_lvo = 2206

# Ruta del archivo Excel
excel_file = 'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\Extensión\\CursosCupoID.xlsx'

# Leer el archivo Excel
df = pd.read_excel(excel_file)

# Crear el navegador usando la función compartida
new_driver = crear_navegador_chrome()

try:
    # Abrir la página web
    new_driver.get('https://academico.sise.edu.pe/psp/CS90PRD/?cmd=login')

    # Iniciar sesión
    new_driver.find_element(By.ID, "userid").send_keys("JRAMOSC")
    new_driver.find_element(By.ID, "pwd").send_keys("jOLgEt6n")
    new_driver.find_element(By.CSS_SELECTOR, 'button.btn-danger-sise').click()
    time.sleep(2)

    # Acceder a inscripción rápida
    WebDriverWait(new_driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'a#fldra_MYFAVORITES'))).click()
    WebDriverWait(new_driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Programa de Curso Nuevo"))).click()
    time.sleep(2)

    new_driver.switch_to.frame("ptifrmtgtframe")

    for index, row in df.iterrows():
        numero_clase = row['NroClase']
        id_curso = row['IdCurso']
        tot_inscritos = row['Tot_Inscr']
        nueva_capacidad = tot_inscritos + 12

        print(f"Procesando clase {numero_clase} - Curso ID {id_curso}")
        time.sleep(2)
        new_driver.find_element(By.ID, "#ICClear").click()
        time.sleep(3)

        new_driver.find_element(By.ID, "CRSE_OFFER_SCTY_INSTITUTION").send_keys("SISE")
        new_driver.find_element(By.ID, "CRSE_OFFER_SCTY_STRM").send_keys(f"{ciclo_lvo}")
        Select(new_driver.find_element(By.ID, "CRSE_OFFER_SCTY_ACAD_CAREER")).select_by_value("VIRT")
        time.sleep(1)
        new_driver.find_element(By.ID, "CRSE_OFFER_SCTY_CRSE_ID").send_keys(f"{id_curso}")
        new_driver.find_element(By.ID, "#ICSearch").click()
        time.sleep(9)

        try:
            new_driver.find_element(By.ID, "$ICField3$hviewall$0").click()
            time.sleep(3)
        except NoSuchElementException:
            print("Botón 'Ver Todo' no disponible")

        numero_clase_correlativo = None
        for i in range(200):
            try:
                new_driver.find_element(By.XPATH, f"//span[@id='CLASS_TBL_CLASS_NBR${i}' and text()='{numero_clase}']")
                numero_clase_correlativo = i
                break
            except NoSuchElementException:
                pass

        if numero_clase_correlativo is not None:
            capacidad_input = new_driver.find_element(By.ID, f"CLASS_TBL_ENRL_CAP${numero_clase_correlativo}")
            capacidad_input.clear()
            time.sleep(2)
            capacidad_input.send_keys(str(nueva_capacidad))
            print(f"Se actualiza capacidad de clase {numero_clase} a {nueva_capacidad}")

            time.sleep(2)

            new_driver.find_element(By.ID, "#ICSave").click()
            time.sleep(3)

            try:
                guardado = WebDriverWait(new_driver, 5).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[contains(text(),'Guardado')]"))
                )
                print(f"Se ha agregado cupo correctamente a {nueva_capacidad}")

            except TimeoutException:
                while True:
                    try:
                        estado = new_driver.find_element(By.XPATH,
                                                         "//*[contains(text(),'Guardando...')]")
                        new_driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ENTER)
                    except NoSuchElementException:
                        pass

                    try:
                        guardado = new_driver.find_element(By.XPATH,
                                                           "//*[contains(text(),'Guardado')]")
                        print(f"Se ha agregado cupo correctamente a {nueva_capacidad}")
                        break
                    except NoSuchElementException:
                        pass

            time.sleep(2)

            # Hacer clic en el botón "Volver a Buscar"
            boton_volver_a_buscar = new_driver.find_element(By.ID, "#ICList")
            boton_volver_a_buscar.click()
            time.sleep(3)

            # Hacer clic en el botón "Borrar"
            boton_borrar = new_driver.find_element(By.ID, "#ICClear")
            boton_borrar.click()
            time.sleep(2)

        else:
            print(f"No se encontró la clase {numero_clase} en la página.")

        time.sleep(3)

finally:
    new_driver.quit()
