import pandas as pd
import re

def extraer_datos_por_secciones(archivo_excel):
    # Leer todas las hojas del archivo Excel
    hojas = pd.read_excel(archivo_excel, sheet_name=None)

    # Leer la hoja "Cursos" y crear un mapeo de Nº Clase a Nombre Curso
    hoja_cursos = hojas.get("Cursos")
    if hoja_cursos is None:
        raise ValueError("La hoja 'Cursos' no existe en el archivo Excel.")

    mapeo_cursos = hoja_cursos.set_index("Nº Clase")["Nombre Curso"].to_dict()

    datos_finales = []  # Aquí guardaremos los datos procesados

    # Delimitadores de las secciones
    inicio_seccion = r'HORARIOS ENERO 2025-\s*.*?TURNO MAÑANA'
    fin_seccion = r'HORARIO SUJETO A MODIFICACIONES\s*'

    for nombre_hoja, hoja in hojas.items():
        print(f"Procesando hoja: {nombre_hoja}")

        if nombre_hoja == "Cursos":
            continue  # Saltar la hoja "Cursos", ya que ya la procesamos

        # Limpiar y normalizar los datos
        hoja = hoja.dropna(how="all")  # Eliminar filas completamente vacías
        hoja.columns = hoja.columns.str.strip().str.upper()  # Normalizamos los nombres de las columnas
        hoja = hoja.fillna('').astype(str)  # Convertir valores vacíos a cadenas vacías

        # Combinar todas las filas en un texto limpio
        contenido_hoja = "\n".join(
            [" ".join(row).replace("nan", "").strip() for row in hoja.values]
        )

        # Normalizar el contenido del texto
        contenido_hoja = re.sub(r'\s+', ' ', contenido_hoja)

        # Antes de dividir en secciones, verifica si falta el delimitador inicial
        if not contenido_hoja.strip().startswith("HORARIOS ENERO 2025"):
            contenido_hoja = f"HORARIOS ENERO 2025-TURNO MAÑANA {contenido_hoja}"

        print(contenido_hoja)
        # Dividir el contenido en secciones
        secciones = dividir_en_secciones(contenido_hoja, inicio_seccion, fin_seccion)

        # Procesar cada sección encontrada
        for idx, contenido_seccion in enumerate(secciones, 1):
            #print(f"Procesando Sección {idx} (primeros 500 caracteres): {contenido_seccion[:500]}")

            # Extraer sede, modalidad, sección (desde el encabezado)
            sede_match = re.search(r'(\b[^\d,]+?)(?=\s+(SEMIPRESENCIAL|PRESENCIAL|VIRTUAL))', contenido_seccion)
            modalidad_match = re.search(r'(SEMIPRESENCIAL|PRESENCIAL|VIRTUAL)', contenido_seccion)
            seccion_match = re.search(r'GESTION Y NEGOCIOS.*?(\b\d+[A-Z]{1,2}\b)', contenido_seccion)

            seccion = seccion_match.group(1).strip() if seccion_match else 'Sección no identificada'
            sede = sede_match.group(1).strip() if sede_match else 'Sede no identificada'
            print(sede)
            modalidad = modalidad_match.group(1).strip() if modalidad_match else 'Modalidad no identificada'
            #print(modalidad)

            # Procesar filas individuales dentro de esta sección
            for linea in contenido_seccion.split("\n"):
                # Normalizar la línea
                linea = re.sub(r'\s+', ' ', linea).strip()

                # Buscar horarios
                hora_match = re.findall(r'\b\d{2}:\d{2}\b', linea)

                # Filtrar cualquier valor vacío de hora_match
                if not hora_match:
                    continue

                hora = f"{hora_match[0]} - {hora_match[1]}" if len(hora_match) >= 2 else hora_match[0]

                # Buscar números de clase de 4 cifras
                numeros_clase = re.findall(r'\b\d{4}\b', linea)

                # Buscar nombres de cursos usando el mapeo
                cursos = [mapeo_cursos.get(int(numero), "Curso no encontrado") for numero in numeros_clase]

                # Agregar datos al resultado
                for curso, numero_clase in zip(cursos, numeros_clase):
                    datos_finales.append([sede, seccion, hora, curso, numero_clase, modalidad])


    # Crear un DataFrame con los resultados
    df_resultado = pd.DataFrame(datos_finales,
                                columns=['Sede', 'Sección', 'Hora', 'Curso', 'Número de clase', 'Modalidad'])

    # Eliminar el prefijo "SEDE: " de la columna "Sede"
    df_resultado["Sede"] = df_resultado["Sede"].str.replace(r"^SEDE:\s*", "", regex=True)

    # Mapeo de horas a identificadores
    mapeo_horas = {
        "07:30 - 09:10": "A",
        "09:10 - 10:50": "M",
        "10:50 - 12:30": "C"
    }

    # Función para determinar el identificador de una sección
    def asignar_identificador(grupo):
        # Encuentra la primera hora válida en el grupo
        primera_hora = grupo["Hora"].iloc[0]
        identificador = mapeo_horas.get(primera_hora, "Horario no definido")
        # Asigna el identificador a toda la sección
        grupo["Hora"] = identificador
        return grupo

    # Aplicar la función a cada grupo de sección
    df_resultado = df_resultado.groupby("Sección").apply(asignar_identificador)

    return df_resultado


def dividir_en_secciones(texto, inicio, fin):
    secciones = []
    for match in re.finditer(f"{inicio}(.*?){fin}", texto, re.DOTALL):
        secciones.append(match.group(1))
    return secciones


# Ruta del archivo de ejemplo
archivo_excel = 'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\En Construcion\\horarios.xlsx'


# Llamar a la función y guardar los resultados
resultado = extraer_datos_por_secciones(archivo_excel)

# Guardar el resultado en un nuevo archivo Excel
resultado.to_excel('resultado_horarios_secciones.xlsx', index=False)
print("Procesamiento completo. Los datos se guardaron en 'resultado_horarios_secciones.xlsx'.")