import pandas as pd

def calcular_ocupabilidad(archivo_excel, hoja=0):
    # Cargar el archivo Excel
    df = pd.read_excel(archivo_excel, sheet_name=hoja)

    # Suponemos que la primera columna es 'AULA' y las demás son franjas horarias
    columnas_horarios = df.columns[1:]  # Tomamos todas las columnas después de 'AULA'

    # Normalizar los nombres de las columnas eliminando sufijos numéricos
    horarios_normalizados = [col.split('.')[0] for col in columnas_horarios]
    df.columns = ['AULA'] + horarios_normalizados

    # Subintervalos
    sub_intervalos = {
        "07:30 A 09:10": ["07:30 A 08:20", "08:20 A 09:10"],
        "09:10 A 10:50": ["09:10 A 10:00", "10:00 A 10:50"],
        "10:50 A 12:30": ["10:50 A 11:40", "11:40 A 12:30"],
        "12:30 A 02:10": ["12:30 A 01:20", "01:20 A 02:10"]
    }

    # Tabla de ocupabilidad
    ocupabilidad_map = {
        0: "0%",
        1: "17%",
        2: "33%",
        3: "50%",
        4: "67%",
        5: "83%",
        6: "100%"
    }

    # Crear un nuevo DataFrame con horarios consolidados
    df_consolidado = pd.DataFrame()
    df_consolidado['AULA'] = df['AULA']

    # Función para verificar si la celda tiene contenido significativo
    def tiene_contenido(valor):
        if pd.isna(valor) or valor == '':
            return False
        return True

    # Calcular ocupabilidad para cada bloque horario
    for horario in sub_intervalos.keys():
        print(f"\nProcesando horario: {horario}")

        # Filtrar las columnas que corresponden al horario actual
        columnas_horario = [col for col in df.columns if col == horario]
        print(f"Columnas para {horario}: {columnas_horario}")

        # Eliminar duplicados en las columnas de horarios para evitar contar varias veces las mismas columnas
        columnas_horario_unicas = list(set(columnas_horario))
        print(f"Columnas únicas para {horario}: {columnas_horario_unicas}")

        # Si no hay columnas para este horario, continuar
        if not columnas_horario_unicas:
            print(f"No se encontraron columnas para el horario {horario}.")
            continue

        # Agrupar todas las columnas del mismo horario y sumar celdas no vacías
        df_horario = df[columnas_horario_unicas].applymap(tiene_contenido)
        conteo_por_fila = df_horario.sum(axis=1)  # Sumar celdas no vacías por fila
        print(f"Conteo de celdas no vacías por fila para {horario}: \n{conteo_por_fila}")

        # Limitar el conteo de celdas al rango válido del mapeo
        conteo_por_fila = conteo_por_fila.clip(upper=6)
        print(f"Conteo recortado por fila para {horario}: \n{conteo_por_fila}")

        # Calcular el porcentaje de ocupabilidad según el conteo de celdas no vacías
        porcentaje_ocupabilidad = conteo_por_fila.apply(lambda x: ocupabilidad_map.get(x, "100%"))
        print(f"Porcentaje de ocupabilidad calculado para {horario}: \n{porcentaje_ocupabilidad}")

        # Asignar el porcentaje a las columnas de ocupabilidad correspondientes al bloque completo
        df_consolidado[horario] = porcentaje_ocupabilidad
        print(f"Asignación de ocupabilidad para el bloque completo ({horario}): \n{df_consolidado[horario]}")

        # Luego, asignar el mismo porcentaje a los subintervalos correspondientes
        for sub_horario in sub_intervalos[horario]:
            print(f"Asignando porcentaje a sub-horario: {sub_horario}")
            df_consolidado[sub_horario] = porcentaje_ocupabilidad
            print(f"Asignación de ocupabilidad para sub-horario ({sub_horario}): \n{df_consolidado[sub_horario]}")

    return df_consolidado


# Uso del código
archivo_excel = "ocupabilidad.xlsx"  # Reemplaza con el nombre real de tu archivo
resultado = calcular_ocupabilidad(archivo_excel)
print("\nResultado final:\n", resultado)

# Guardar en un nuevo archivo
resultado.to_excel("ocupabilidad_resultado.xlsx", index=False)