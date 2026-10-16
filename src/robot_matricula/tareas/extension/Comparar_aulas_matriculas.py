from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
import time
import pandas as pd
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
import csv

# Solicitar al usuario que ingrese el número de Ciclo Lvo.
ciclo_lvo = input("Ingrese el número de Ciclo Lvo.: ")  # Cambiado a string para evitar problemas de comparación

try:
    # Configurar el servicio de Chrome
    chrome_service = webdriver.chrome.service.Service(
        'C:\\Users\\jramosc\\Downloads\\chromedriver-win64\\Chromedriver.exe')
    chrome_service.start()

    # Configurar opciones del navegador
    chrome_options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

    # Iniciar sesión
    driver.get("https://aulavirtualfcid.sise.edu.pe/login/canvas")
    wait = WebDriverWait(driver, 10)
    boton = wait.until(EC.element_to_be_clickable((By.ID, "btn-adm")))
    boton.click()
    time.sleep(3)

    usuario_input = driver.find_element(By.ID, "pseudonym_session_unique_id")
    usuario_input.send_keys("echamorro@sise.edu.pe")

    password_input = driver.find_element(By.ID, "pseudonym_session_password")
    password_input.send_keys("12345678")

    login_button = driver.find_element(By.CLASS_NAME, "Button--login")
    login_button.click()
    time.sleep(2)

    admin_link = wait.until(EC.presence_of_element_located((By.XPATH, "//a[@id='global_nav_accounts_link']")))
    admin_link.click()

    instituto_sise_link = wait.until(
        EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'SISE Extensión')]")))
    instituto_sise_link.click()

    # Leer los números de clase desde el archivo Excel
    excel_file = 'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\Extensión\\numerosClase.xlsx'
    df_codigos = pd.read_excel(excel_file, sheet_name='Hoja1')

    # Lista para almacenar los resultados y errores
    resultados = []
    errores = []

    for index_excel, row in df_codigos.iterrows():  # Recorrer filas del Excel
        nro_clase = row['NroClase']
        Tot_Inscr = row['Tot_Inscr']

        try:
            Tot_Inscr = int(Tot_Inscr)  # Convertir a entero
        except (ValueError, TypeError):
            print(f"❌ Error: Tot_Inscr inválido en la fila {index_excel + 2}. Saltando...")
            continue  # Saltar esta iteración si Tot_Inscr no es válido
        time.sleep(1)
        search_input = wait.until(EC.presence_of_element_located((By.ID, "TextInput_2")))
        search_input.send_keys(Keys.CONTROL + "a")  # Seleccionar todo
        search_input.send_keys(Keys.DELETE)  # Borrar
        time.sleep(0.5)
        search_input.send_keys(str(nro_clase))
        time.sleep(2)

        filas = driver.find_elements(By.XPATH, "//tr[contains(@class, 'css-uysa3a-view-row')]")


        try:
            for index, fila in enumerate(filas, start=1):
                try:
                    time.sleep(1)
                    celda_ciclo = fila.find_element(By.XPATH, "./td[3]")

                    if ciclo_lvo in celda_ciclo.text:
                        celda_td6 = fila.find_element(By.XPATH, "./td[6]")
                        num_estudiantes = int(celda_td6.text.strip())  # Convertir a entero

                        print(f"🔹 Numero de clase: {nro_clase} (Estudiantes: {num_estudiantes})")

                        # Comparar con Tot_Inscr y guardar en errores
                        if num_estudiantes > Tot_Inscr:
                            errores.append((nro_clase, "Verificar"))
                            print(
                                f"⚠️ Estudiantes en Canvas de {nro_clase} ({num_estudiantes}) es mayor que Tot_Inscr ({Tot_Inscr}).")
                            time.sleep(2)
                        else:
                            errores.append((nro_clase, "Correcto"))
                            print(
                                f"✅ CORRECTO: {num_estudiantes} estudiantes en fila {index}, coincide o es menor que {Tot_Inscr}")
                    else:
                        errores.append((nro_clase, "No se encontró curso en Canvas"))
                        print(f"❌ No se encontró el curso {nro_clase} en Canvas para el ciclo {ciclo_lvo}")

                except NoSuchElementException:
                    continue  # Si no encuentra la celda, simplemente pasa a la siguiente fila

                except Exception as e:
                    print(f"❌ Error en fila {index}: {e}")
                    errores.append((nro_clase, f"Error: {str(e)}"))
                    continue

        except Exception as e:
            print(f"❌ Error en la búsqueda de filas para la clase {nro_clase}: {e}")
            errores.append((nro_clase, f"Error: {str(e)}"))
            continue

finally:
    # Cierra la ventana actual del navegador
    if 'driver' in locals():
        driver.close()

    # Detiene el servicio
    if 'chrome_service' in locals():
        chrome_service.stop()

    # Guardar resultados en un archivo CSV (Correcto y Verificar incluidos)
    intentos = 3  # Número máximo de intentos permitidos

    while intentos > 0:
        try:
            error_csv_path = "C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\Extensión\\reporte.csv"
            with open(error_csv_path, 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(['Codigo', 'Estado'])
                for codigo, error_msg in errores:
                    writer.writerow([codigo, error_msg])

            print("========================================================================")
            print("📢 Se generó un reporte con TODOS los estados (Correcto y Verificar).")
            for codigo, error_msg in errores:
                print(f"⚠️ Código: {codigo}, Estado: {error_msg}")

            # Si todo fue exitoso, salir del bucle while
            break
        except Exception as e:
            print(f"❌ Error al generar el reporte: {e}")
            intentos -= 1  # Reducir intentos disponibles

            if intentos > 0:
                respuesta = input("¿Desea volver a intentar? (Si/No): ")
                if respuesta.lower() != 'si':
                    break
            else:
                print("❌ No se pudieron guardar los resultados en el reporte.")
