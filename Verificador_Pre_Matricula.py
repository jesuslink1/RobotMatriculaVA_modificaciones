import pandas as pd
import re

# Cargar el archivo Excel con las tres pestañas y la lista de alumnos
file_path = 'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\PreMatricula.xlsx'
presencial = pd.read_excel(file_path, sheet_name='Presencial')
semipresencial = pd.read_excel(file_path, sheet_name='Semipresencial')
virtual = pd.read_excel(file_path, sheet_name='Virtual')
sheet1 = pd.read_excel(file_path, sheet_name='sheet1')  # Lista de alumnos

# Mapear las carreras a escuelas profesionales
escuelas = {
    'Banca': ['ADMBF', 'CONTA'],
    'Gestión': ['ADMLO', 'ADMNE', 'ADMNI', 'MARKT'],
    'Salud': ['SALUD', 'TECFA', 'FISIO'],
    'Discom': ['DG', 'COMAU', 'PA'],
    'ETI': ['SWSIS', 'REDES'],
    'Modas': ['DYGMO'],
    'Gastronomia': ['GYGEC'],
    'Hoteles': ['ADYRH'],
    'Secretariado': ['SECRET'],
    'Dis. Interiores': ['DISIN'],
    'PUMD': ['PUMD'],  # Aquí se cambia el nombre de la columna para que entienda la lectura.
    'Ciberseguridad': ['CIBER']
}

# Ajuste específico para modalidad virtual: Solo Diseño Gráfico (DG) dentro de Discom
escuelas_virtual = {
    'Diseño Gráfico': ['DG'],  # Solo DG de Discom para modalidad virtual
    'Banca': ['ADMBF', 'CONTA'],
    'Gestión': ['ADMLO', 'ADMNE', 'ADMNI', 'MARKT'],
    'ETI': ['SWSIS', 'REDES'],
    'Secretariado': ['SECRET']
}

# Carreras que no aperturan
carreras_no_aperturan = []

# Mapear los turnos
nombre_a_codigo = {
    'Madrugador': 'A',
    'Mañana01': 'M',
    'Mañana02': 'C',
    'Tarde': 'T',
    'Noche': 'N'
}

codigo_a_nombre = {v: k for k, v in nombre_a_codigo.items()}

# Mapear los campus
campus_map = {
    'Ate': 'ATE',
    'Santa Beatriz': 'STABE',
    'Surco': 'SURCO',
    'SJM': 'SJM',
    'SJL': 'SJL',
    'Independencia': 'INDEP',
    'VES': 'VILLA',
    'Ventanilla': 'VENTN',
    'Miraflores': 'MIRAF',
    'Cantuarias': 'CANTU',
}


# Función para limpiar y extraer turnos de las celdas
def extraer_turnos(celda):
    if pd.isna(celda):
        return []
    # Reemplazar guiones por espacios para separar los turnos
    celda = re.sub(r'[-–]', ' ', celda)  # Consideramos también el guion largo
    # Eliminar cualquier otro carácter no deseado, excepto letras, números y espacios
    celda = re.sub(r'[^\w\s]', '', celda)
    turnos_limpios = re.split(r'\s+', celda.strip())  # Dividir por espacios
    turnos_disponibles = [nombre_a_codigo.get(turno.strip(), None) for turno in turnos_limpios]
    return [t for t in turnos_disponibles if t is not None]


# Función para transformar los datos de una hoja de modalidad (Presencial, Semipresencial, Virtual)
def transformar_hoja_aperturas(hoja, modalidad):
    aperturas = {}

    # Seleccionar el mapeo de escuelas basado en la modalidad
    if modalidad == 'Virtual':
        escuelas_actuales = escuelas_virtual
    else:
        escuelas_actuales = escuelas

    # Recorrer filas y columnas para extraer datos
    for idx, fila in hoja.iterrows():
        campus = fila['Sedes']
        campus = campus_map.get(campus, campus)  # Mapear el nombre del campus

        for escuela in escuelas_actuales:
            if escuela in fila:
                turnos_disponibles = extraer_turnos(fila[escuela])

                # Imprimir los turnos que encuentra por fila y escuela
                # print(f"Modalidad: {modalidad}, Campus: {campus}, Escuela: {escuela}, Turnos disponibles: {turnos_disponibles}")

                if turnos_disponibles:
                    if campus not in aperturas:
                        aperturas[campus] = {}
                    aperturas[campus][escuela] = turnos_disponibles
    return aperturas

# Obtener aperturas de las tres hojas


aperturas_presencial = transformar_hoja_aperturas(presencial, 'Presencial')
aperturas_semipresencial = transformar_hoja_aperturas(semipresencial, 'Semipresencial')
aperturas_virtual = transformar_hoja_aperturas(virtual, 'Virtual')
print(f"aperturas_virtual {aperturas_virtual}")


# Función para verificar la coincidencia y guardar la discrepancia
def verificar_discrepancia(alumno, aperturas_presencial, aperturas_semipresencial, aperturas_virtual):
    campus_alumno = alumno['Campus']
    turno_alumno = alumno['Turno Académico'].strip()
    modalidad_alumno = alumno['Modalidad']
    carrera_alumno = alumno['Programa Académico']

    discrepancia = {
        'ID': alumno['ID'],
        'Carrera': carrera_alumno,
        'Turno': turno_alumno,
        'Modalidad': modalidad_alumno,
        'Campus': campus_alumno,
        'Razon': []
    }

    # Imprimir el progreso del alumno que está siendo procesado
    print(f"\n--- Procesando ID: {alumno['ID']} ---")
    print(f"Carrera: {carrera_alumno}, Turno: {turno_alumno}, Modalidad: {modalidad_alumno}, Campus: {campus_alumno}")

    # Si el alumno no tiene modalidad
    if pd.isna(modalidad_alumno) or modalidad_alumno.strip() == "":
        if carrera_alumno in escuelas['Salud'] or carrera_alumno in escuelas['Gastronomia']:
            discrepancia['Razon'].append('Se le debe asignar modalidad Presencial')
        else:
            discrepancia['Razon'].append('El alumno no tiene modalidad')
        return discrepancia

    # Ajustar modalidad si es Presencial o Semipresencial
    if modalidad_alumno in ['Presencial', 'Semipresencial']:
        modalidad_alumno = 'Presencial'
    elif modalidad_alumno == 'Virtual':
        # Verificar si la carrera soporta modalidad virtual
        # Se comprueba si la carrera pertenece a alguna de las escuelas que no soportan modalidad virtual
        escuelas_que_no_aperturan_virtual = ['Salud', 'Gastronomia', 'Hoteles', 'Modas']
        carreras_no_virtual = sum([escuelas[escuela] for escuela in escuelas_que_no_aperturan_virtual], [])

        if carrera_alumno in carreras_no_virtual:
            discrepancia['Razon'].append(f'La carrera {carrera_alumno} no tiene modalidad Virtual')
            return discrepancia

    aperturas = aperturas_presencial if modalidad_alumno == 'Presencial' else aperturas_virtual
    print("---------------------")
    print(aperturas)
    print(modalidad_alumno)
    print(turno_alumno)
    print(carrera_alumno)

    # Definir el diccionario de escuelas adecuado
    escuelas_diccionario = escuelas if modalidad_alumno != 'Virtual' else escuelas_virtual

    for escuela, carreras in escuelas_diccionario.items():
        if carrera_alumno in carreras:
            print(f"Modalidad alumno: {modalidad_alumno}, Carrera: {carrera_alumno}, Escuela: {escuela}")

            # Verificar si es modalidad "Presencial"
            if modalidad_alumno == 'Presencial':
                print("Modalidad: Presencial")

                # Buscar primero en Presencial
                if campus_alumno in aperturas_presencial and escuela in aperturas_presencial[campus_alumno]:
                    turnos_disponibles = aperturas_presencial[campus_alumno][escuela]
                    print(f"Turnos disponibles en Presencial: {turnos_disponibles}")
                    print(f"Turno del alumno: {turno_alumno}")

                    if turno_alumno in turnos_disponibles:
                        print("Turno del alumno encontrado en Presencial.")
                        return None  # No hay discrepancia
                else:
                    print("Turno no encontrado en Presencial. Verificando en Semipresencial...")

                # Si no está en Presencial, buscar en Semipresencial
                if campus_alumno in aperturas_semipresencial and escuela in aperturas_semipresencial[campus_alumno]:
                    turnos_disponibles = aperturas_semipresencial[campus_alumno][escuela]
                    print(f"Turnos disponibles en Semipresencial: {turnos_disponibles}")

                    if turno_alumno in turnos_disponibles:
                        print("Turno del alumno encontrado en Semipresencial.")
                        return None  # No hay discrepancia
                else:
                    print("Turno no encontrado en Semipresencial.")

                # Si no se encuentra ni en Presencial ni en Semipresencial
                print("Añadiendo discrepancia por no encontrar el turno en ninguna modalidad.")
                discrepancia['Razon'].append(
                    f'Turno {turno_alumno} no aperturado en la carrera {carrera_alumno} en modalidad Presencial/Semipresencial sede: {campus_alumno}'
                )
                print("Discrepancia añadida.")

            elif modalidad_alumno == 'Virtual':

                print("Modalidad: Virtual")
                # Verificar si la sede es distinta a "STABE" y generar una alerta
                if campus_alumno != 'STABE':
                    discrepancia['Razon'].append(
                        'Alerta: Modalidad virtual con sede distinta a STABE. '
                    )
                    print("Alerta de sede distinta a STABE para alumno virtual.")
                    # Asumir que el campus ahora es "STABE"
                    campus_alumno = 'STABE'
                    print("Campus reasignado a STABE para continuar la verificación.")

                # Verificar si el campus y la escuela están en aperturas_virtual
                if campus_alumno in aperturas_virtual:
                    print(f"Campus encontrado en Virtual: {campus_alumno}")
                    print(f"Escuela: {escuela}")
                    print(f"aperturas_virtual[campus_alumno] {aperturas_virtual[campus_alumno]}")
                    if escuela in aperturas_virtual[campus_alumno]:
                        print(f"Escuela encontrada en Virtual: {escuela}")
                        turnos_disponibles = aperturas_virtual[campus_alumno][escuela]
                        print(f"Turnos disponibles en Virtual: {turnos_disponibles}")
                        print(f"Turno del alumno: {turno_alumno}")
                        if turno_alumno in turnos_disponibles:
                            print("Turno del alumno encontrado en Virtual.")

                            return discrepancia if discrepancia['Razon'] else None

                        else:
                            print("Turno del alumno no encontrado en Virtual.")
                    else:
                        print(
                            f"Escuela {escuela} no encontrada en aperturas de Virtual para el campus {campus_alumno}.")
                else:
                    print(f"Campus {campus_alumno} no encontrado en aperturas de Virtual.")
                # Si no se encuentra el turno en Virtual, añadir discrepancia
                discrepancia['Razon'].append(
                    f'Turno {turno_alumno} no aperturado en la carrera {carrera_alumno} en modalidad Virtual sede: {campus_alumno}'
                )
                print("Discrepancia añadida.")

            break  # Termina la búsqueda para esta carrera

    if carrera_alumno in carreras_no_aperturan:
        discrepancia['Razon'].append(f'Carrera {carrera_alumno} no se aperturó este ciclo')

    return discrepancia if discrepancia['Razon'] else None


# Lista para guardar discrepancias
discrepancias = []

# Recorrer cada alumno y verificar si hay discrepancia
for index, alumno in sheet1.iterrows():
    resultado = verificar_discrepancia(alumno, aperturas_presencial, aperturas_semipresencial, aperturas_virtual)
    if resultado:
        discrepancias.append(resultado)
        # Imprimir la discrepancia encontrada
        print(f"Discrepancia encontrada para ID {resultado['ID']}: {resultado['Razon']}")

# Crear un DataFrame con las discrepancias y exportar a Excel
discrepancias_df = pd.DataFrame(discrepancias)
discrepancias_df.to_excel('resultadoPreMatricula.xlsx', index=False)

print("Archivo 'resultadoPreMatricula.xlsx' creado con éxito.")
