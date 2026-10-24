from selenium import webdriver


def crear_navegador_chrome(
    opciones: webdriver.ChromeOptions | None = None,
    maximizar: bool = True,
) -> webdriver.Chrome:
    """
    Crea una instancia de Chrome utilizando Selenium Manager.

    Args:
        opciones: Opciones específicas de Chrome para cada tarea.
        maximizar: Indica si se maximiza la ventana al iniciar.

    Returns:
        Una instancia de Chrome lista para utilizarse.
    """
    if opciones is None:
        opciones = webdriver.ChromeOptions()

    # Selenium Manager gestiona ChromeDriver automáticamente.
    driver = webdriver.Chrome(options=opciones)

    if maximizar:
        driver.maximize_window()

    return driver