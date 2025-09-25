import pandas as pd
import re

# Cargar el archivo Excel
file_path = r"C:\Users\jramosc\Desktop\Proyectos\RobotMatriculaVA_modificaciones\En Construcion\VerificadorTurnoSede.xlsx"
df = pd.read_excel(file_path)

# Función para obtener el primer número de la sección
def extraer_ciclo(seccion):
    match = re.match(r'\d+', str(seccion))
    return int(match.group()) if match else None

df['CicloSeccion'] = df['Sección'].apply(extraer_ciclo)


def deducir_turno(row_group):
    turnos = row_group['TurnoClase']
    idturnos = row_group['IdTurno']

    conteo = turnos.value_counts()
    max_frecuencia = conteo.max()
    candidatos = conteo[conteo == max_frecuencia].index.tolist()

    if len(candidatos) == 1:
        return candidatos[0]

    if "M" in candidatos and len(candidatos) > 1:
        candidatos = [t for t in candidatos if t != "M"]

    if sorted(candidatos) == ['A', 'C']:
        valores_idturno = idturnos.unique()
        for val in valores_idturno:
            if val in ['A', 'C']:
                return val
        return "Se debe verificar turno A o C"

    return sorted(candidatos)[0]

# Normalizar la modalidad del alumno
def normalizar_modalidad(modalidad):
    return 'Presencial' if modalidad in ['Presencial', 'Semipresencial'] else ('Virtual' if modalidad == 'Virtual' else '')

df['Modalidad'] = df['Modalidad'].apply(normalizar_modalidad)

# Determinar la modalidad de cada curso (NroClase)
modalidad_curso = df.groupby('NroClase')['ModalidadReunión'].apply(
    lambda modalidades: 'Presencial' if 'Presencial' in modalidades.values else 'Virtual'
).to_dict()
df['ModalidadCurso'] = df['NroClase'].map(modalidad_curso)

# Determinar la modalidad predominante del alumno
modalidad_alumno = df.groupby('IdAlumno')['ModalidadCurso'].apply(
    lambda modalidades: 'Presencial' if 'Presencial' in modalidades.values else 'Virtual'
).to_dict()
df['ModalidadDeducida'] = df['IdAlumno'].map(modalidad_alumno)

# Determinar el campus de cada curso (NroClase), ignorando cursos virtuales
campus_curso = df[df['ModalidadReunión'] != 'Virtual'].groupby('NroClase')['IdCampus'].agg(lambda x: x.mode()[0]).to_dict()
df['CampusCurso'] = df['NroClase'].map(campus_curso)

# Determinar el campus predominante del alumno (solo si es presencial)
campus_alumno = df[df['ModalidadCurso'] == 'Presencial'].groupby('IdAlumno')['CampusCurso'].agg(lambda x: x.mode()[0]).to_dict()
df['CampusDeducido'] = df['IdAlumno'].map(campus_alumno)

prioridad_turno = ['C', 'A', 'M']

# Filtrar para mantener solo una fila por NroClase y alumno
df_unique = df.drop_duplicates(subset=['IdAlumno', 'NroClase'])

# Aplicar deducción de turno sobre los datos únicos
turno_alumno = df_unique.groupby('IdAlumno').apply(deducir_turno).to_dict()

# Mapear el resultado al DataFrame original
df['TurnoDeducido'] = df['IdAlumno'].map(turno_alumno)


# Identificar alumnos sin modalidad
df['AlumnoSinModalidad'] = df['Modalidad'] == ''

# Reemplazar 'CANTU' por 'MIRAF' en las columnas 'IdCampusAlu' y 'IdCampus'
df['IdCampusAlu'] = df['IdCampusAlu'].replace({'CANTU': 'MIRAF'})
df['CampusDeducido'] = df['CampusDeducido'].replace({'CANTU': 'MIRAF'})
df['CampusCurso'] = df['CampusCurso'].replace({'CANTU': 'MIRAF'})

# Comparar modalidad y turno (Sede solo si el alumno es presencial)
df['DiscrepanciaModalidad'] = (df['Modalidad'] != df['ModalidadDeducida']) & ~df['AlumnoSinModalidad']
df['DiscrepanciaTurno'] = df['IdTurno'] != df['TurnoDeducido']
df['DiscrepanciaSede'] = (df['ModalidadDeducida'] == 'Presencial') & (df['IdCampusAlu'] != df['CampusDeducido'])

# Contar cuántos cursos distintos tiene cada alumno
cursos_por_alumno = df.groupby('IdAlumno')['NroClase'].nunique().to_dict()
df['CursosDistintos'] = df['IdAlumno'].map(cursos_por_alumno)

# Filtrar discrepancias y agrupar por alumno y tipo de discrepancia
discrepancias_lista = []

discrepancias = df[df[['DiscrepanciaModalidad', 'DiscrepanciaTurno', 'DiscrepanciaSede', 'AlumnoSinModalidad']].any(axis=1)]

if not discrepancias.empty:
    print("\nResumen de Discrepancias:")

    # Agrupar por IdAlumno y UsuarioInscr, tomando la primera aparición para evitar duplicados
    alumnos_con_discrepancias = discrepancias.groupby(['IdAlumno', 'UsuarioInscr']).first().reset_index()

    for _, row in alumnos_con_discrepancias.iterrows():
        id_alumno = row['IdAlumno']
        usuario_inscr = row['UsuarioInscr']
        secciones_alumno = df[df['IdAlumno'] == id_alumno]['Sección'].unique()
        ciclos_seccion = df[df['IdAlumno'] == id_alumno]['CicloSeccion'].unique()

        if row['CursosDistintos'] <= 2:
            razon = "Matriculado a cursos a cargo."
            discrepancias_lista.append({'IdAlumno': id_alumno, 'Discrepancia': razon, 'UsuarioInscr': usuario_inscr})
            print(f"IdAlumno: {id_alumno} - {razon} Inscrito por: {usuario_inscr}")
            continue  # Saltamos la verificación de otras discrepancias

        """
        EN CONSTRUCCIÓN
        # Si los primeros números de la sección son distintos en todos o la mayoría
        if len(set(ciclos_seccion)) > 1:
                razon = "Cursos a cargo: Ciclos de secciones diferentes."
                discrepancias_lista.append(
                    {'IdAlumno': id_alumno, 'Discrepancia': razon, 'UsuarioInscr': usuario_inscr})
                print(f"IdAlumno: {id_alumno} - {razon} Inscrito por: {usuario_inscr}")
                continue

        # Si todas las secciones son distintas dentro del mismo ciclo
        if len(set(secciones_alumno)) == len(secciones_alumno) and len(set(ciclos_seccion)) == 1:
            razon = "Cursos a cargo: Secciones diferentes del mismo ciclo."
            discrepancias_lista.append(
                {'IdAlumno': id_alumno, 'Discrepancia': razon, 'UsuarioInscr': usuario_inscr}
            )
            print(f"IdAlumno: {id_alumno} - {razon} Inscrito por: {usuario_inscr}")
            continue"""

        if row['AlumnoSinModalidad']:
            razon = f"Alumno sin modalidad asignada. Se debería asignar modalidad: ({row['ModalidadDeducida']})."
            discrepancias_lista.append({'IdAlumno': id_alumno, 'Discrepancia': razon, 'UsuarioInscr': usuario_inscr})
            print(f"IdAlumno: {id_alumno} - {razon} Inscrito por: {usuario_inscr}")

        # Lista de programas que deben ignorar cambio Presencial -> Virtual
        programas_ignorar_modalidad = ['GYGEC', 'GASTRO', 'SALUD', 'FISIO', 'TECFA', 'ADHYR', 'DYGMO']

        if row['DiscrepanciaModalidad']:
            ignorar_cambio = (
                    row['Modalidad'] == 'Presencial' and
                    row['ModalidadDeducida'] == 'Virtual' and
                    row['IdPrograma'] in programas_ignorar_modalidad
            )

            if not ignorar_cambio:
                razon = f"Modalidad alumno: ({row['Modalidad']}), modalidad deducida: ({row['ModalidadDeducida']})."
                discrepancias_lista.append(
                    {'IdAlumno': id_alumno, 'Discrepancia': razon, 'UsuarioInscr': usuario_inscr})
                print(f"IdAlumno: {id_alumno} - {razon} Inscrito por: {usuario_inscr}")

        if row['DiscrepanciaTurno']:
            razon = f"Turno alumno: ({row['IdTurno']}), turno deducido: ({row['TurnoDeducido']})."
            discrepancias_lista.append({'IdAlumno': id_alumno, 'Discrepancia': razon, 'UsuarioInscr': usuario_inscr})
            print(f"IdAlumno: {id_alumno} - {razon} Inscrito por: {usuario_inscr}")

        if row['DiscrepanciaSede']:
            razon = f"Sede alumno: ({row['IdCampusAlu']}), sede deducida: ({row['CampusDeducido']})."
            discrepancias_lista.append({'IdAlumno': id_alumno, 'Discrepancia': razon, 'UsuarioInscr': usuario_inscr})
            print(f"IdAlumno: {id_alumno} - {razon} Inscrito por: {usuario_inscr}")

    # Guardar discrepancias en un archivo Excel
    discrepancias_df = pd.DataFrame(discrepancias_lista)
    output_file = 'resultadoVerificadorSedeTurno.xlsx'
    discrepancias_df.to_excel(output_file, index=False)
    print(f"\nArchivo '{output_file}' creado con éxito.")
else:
    print("\nNo se encontraron discrepancias.")