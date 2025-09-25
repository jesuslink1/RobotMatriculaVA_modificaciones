import pandas as pd

# Cargar archivo Excel
file_path = r"C:\Users\jramosc\Desktop\Proyectos\RobotMatriculaVA_modificaciones\Lista Alumnos Matricula Masiva\ListaAlumnos.xlsx"
query_df = pd.read_excel(file_path, sheet_name='Query')
prog_df = pd.read_excel(file_path, sheet_name='Programacion')

# Programas válidos (sin equivalencias)
programas = [
    'ADMNE', 'CONTA', 'SALUD', 'ADMBF', 'ADMNI', 'PUMD', 'SWSIS', 'COMAU',
    'SECRE', 'REDES', 'FISIO', 'DG', 'MARKT', 'TECFA', 'CIBER', 'ADMLO',
    'DISIN', 'GYGEC', 'ADHYR', 'DYGMO', 'COC24', 'PA', 'GASTR'
]
equivalencias_programa = {prog: [prog] for prog in programas}

# Funciones
def traducir_modalidad(mod):
    mod = str(mod).strip().lower()
    if mod in ['presencial', 'semipresencial', '', 'nan']:
        return 'PRE'
    elif mod == 'virtual':
        return 'VIRT'
    return mod.upper()

def normalizar_sede(sede):
    sede = sede.strip().upper()
    return 'MIRAF' if sede in ['CANTU', 'MIRAF'] else sede

def expandir_ciclo(siguiente_ciclo):
    if '|' in siguiente_ciclo:
        base = siguiente_ciclo.split('|')[0]
        sufijos = siguiente_ciclo.split('|')[1:]
        if len(sufijos) == 1 and sufijos[0] in ['B1', 'B2']:
            return [base, base.replace('B1', 'B2') if 'B1' in base else base.replace('B2', 'B1')]
    return [siguiente_ciclo]

# Resultados
resultados = []

# Proceso por alumno
for _, alumno in query_df.iterrows():
    id_alumno = alumno['ID']
    raw_modalidad = str(alumno['Modalidad']).strip()
    raw_sede = str(alumno['Campus']).strip().upper()
    prog_alu = alumno['Programa Académico']
    turno_alu = str(alumno['Turno Académico']).strip()
    siguiente_ciclo = str(alumno.get('SiguienteCiclo', '')).strip().upper()

    print(f"\nProcesando alumno ID: {id_alumno}, Programa: {prog_alu}, Ciclo: {siguiente_ciclo}")

    observacion_extra = ""
    modalidad_alu = ''

    if raw_sede == 'VIRTU':
        if raw_modalidad == '' or raw_modalidad.lower() == 'nan':
            modalidad_alu = 'VIRT'
            observacion_extra = 'Se debe agregar modalidad: Virtual'
        elif traducir_modalidad(raw_modalidad) == 'PRE':
            modalidad_alu = 'VIRT'
            observacion_extra = 'Modalidad debería ser Virtual por Campus VIRTU'
        else:
            modalidad_alu = traducir_modalidad(raw_modalidad)
    else:
        modalidad_alu = 'PRE' if raw_modalidad == '' or raw_modalidad.lower() == 'nan' else traducir_modalidad(raw_modalidad)

    if siguiente_ciclo == 'EGRESADO':
        resultados.append({
            'INSTITUCION': 'SISE',
            'GRUPO_INSCRIPCION': '',
            'DESCRIPCION': '',
            'ID_ALUMNO': id_alumno,
            'GRADO_ACADEMICO': 'PRGS',
            'PROGRAMA_ACADEMICO': prog_alu,
            'OBSERVACION': 'EGRESADO'
        })
        continue

    asignado = False
    grupos_descartados = []
    clases_sin_cupo_total = []
    secciones_sin_cupo_actuales = []

    while not asignado:
        mejor_grupo = None
        mejor_filas = None
        menor_cupo_grupo = None
        clases_sin_cupo_grupo = []

        # 🔍 Filtro principal antes de agrupar
        filtro1 = prog_df[
            (prog_df['Programa que admite'].isin(equivalencias_programa.get(prog_alu, [prog_alu]))) &
            (prog_df['Turno'].astype(str).str.strip() == turno_alu) &
            (prog_df['Modalidad'] == modalidad_alu)
            ]

        if modalidad_alu != 'VIRT':
            filtro1 = filtro1[filtro1['Sede'].apply(normalizar_sede) == normalizar_sede(raw_sede)]

        if siguiente_ciclo == 'VENTA NUEVA':
            ciclos_validos = ['C1B1', 'C1B2']
        else:
            ciclos_validos = expandir_ciclo(siguiente_ciclo)

        filtro1 = filtro1[filtro1['CicloBloque'].astype(str).str.strip().str.upper().isin(ciclos_validos)]

        # ❌ Caso: Alumno no se relaciona con ninguna sección
        if filtro1.empty:
            print("✘ No asignado. Alumno no se relaciona con ninguna sección disponible.")
            resultados.append({
                'INSTITUCION': 'SISE',
                'GRUPO_INSCRIPCION': '',
                'DESCRIPCION': '',
                'ID_ALUMNO': id_alumno,
                'GRADO_ACADEMICO': 'PRGS',
                'PROGRAMA_ACADEMICO': prog_alu,
                'OBSERVACION': 'Alumno no se relaciona con ninguna sección disponible.'
            })
            break

        grupos = filtro1.groupby('GRUPO_INSCRIPCION')

        for grupo_inscripcion, grupo_df in grupos:
            if grupo_inscripcion in grupos_descartados:
                continue

            filas_validas = []
            cupos_clase = []

            for _, fila in grupo_df.iterrows():
                modalidad_prog = fila['Modalidad']
                turno_prog = str(fila['Turno']).strip()
                sede_prog = str(fila['Sede']).strip()
                programa_admite = fila['Programa que admite']
                ciclo_prog = str(fila.get('CicloBloque', '')).strip().upper()
                clase_id = str(fila.get('CLASE', '')).strip()

                equivalentes = equivalencias_programa.get(programa_admite, [programa_admite])

                # Ciclos válidos
                if siguiente_ciclo == 'VENTA NUEVA':
                    ciclos_validos = ['C1B1', 'C1B2']
                else:
                    ciclos_validos = expandir_ciclo(siguiente_ciclo)

                # Filtros
                if modalidad_alu != modalidad_prog:
                    continue
                if prog_alu not in equivalentes:
                    continue
                if turno_alu != turno_prog:
                    continue
                if modalidad_alu != 'VIRT' and normalizar_sede(raw_sede) != normalizar_sede(sede_prog):
                    continue
                if ciclo_prog not in ciclos_validos:
                    continue

                # Cupo
                try:
                    capcd = int(str(fila.get('CapcdInscr', '0')).strip())
                except:
                    capcd = 0
                try:
                    totinscr = int(str(fila.get('Tot_Inscr', '0')).strip())
                except:
                    totinscr = 0

                cupo_disp = capcd - totinscr

                if cupo_disp <= 0:
                    clases_sin_cupo_grupo.append(clase_id)
                    secciones_sin_cupo_actuales.append((grupo_inscripcion, clase_id))

                if cupo_disp < 0:
                    continue
                filas_validas.append(fila)
                cupos_clase.append(cupo_disp)

            if filas_validas:
                capcd_vals = []
                totinscr_vals = []

                for _, fila_comp in grupo_df.iterrows():
                    modalidad_prog = fila_comp['Modalidad']
                    turno_prog = str(fila_comp['Turno']).strip()
                    sede_prog = str(fila_comp['Sede']).strip()
                    programa_admite = fila_comp['Programa que admite']
                    ciclo_prog = str(fila_comp.get('CicloBloque', '')).strip().upper()

                    equivalentes = equivalencias_programa.get(programa_admite, [programa_admite])
                    ciclos_validos = ['C1B1', 'C1B2'] if siguiente_ciclo == 'VENTA NUEVA' else expandir_ciclo(
                        siguiente_ciclo)

                    # filtros principales
                    if modalidad_alu != modalidad_prog: continue
                    if prog_alu not in equivalentes: continue
                    if turno_alu != turno_prog: continue
                    if modalidad_alu != 'VIRT' and normalizar_sede(raw_sede) != normalizar_sede(sede_prog): continue
                    if ciclo_prog not in ciclos_validos: continue

                    try:
                        capcd = int(str(fila_comp.get('CapcdInscr', '0')).strip())
                        totinscr = int(str(fila_comp.get('Tot_Inscr', '0')).strip())
                        clase_id = str(fila_comp.get('CLASE', '')).strip()
                        capcd_vals.append(capcd)
                        totinscr_vals.append(totinscr)
                        print(
                            f"Grupo: {grupo_inscripcion}, Clase: {clase_id}, Capacidad: {capcd}, Inscritos: {totinscr}, Cupo: {capcd - totinscr}")
                    except Exception as e:
                        print(f"Error parsing fila: {e}")
                        continue

                if not capcd_vals or not totinscr_vals:
                    print(f"❌ Grupo descartado por falta de datos: {grupo_inscripcion}")
                    grupos_descartados.append(grupo_inscripcion)
                    continue

                cupo_disponible = min(cupos_clase)
                print(f"🧮 Grupo: {grupo_inscripcion} → cupos por clase: {cupos_clase} ⇒ cupo_disponible: {min(cupos_clase)}")


                if cupo_disponible <= 0:
                    print(f"⚠️ Grupo {grupo_inscripcion} descartado por cupo insuficiente.")
                    clases_sin_cupo_total.extend([str(f.get('CLASE', '')).strip() for f in filas_validas])
                    grupos_descartados.append(grupo_inscripcion)
                    continue

                if mejor_grupo is None or cupo_disponible < menor_cupo_grupo:
                    mejor_grupo = grupo_inscripcion
                    mejor_filas = filas_validas
                    menor_cupo_grupo = cupo_disponible
                    print(f"✅ Grupo {grupo_inscripcion} es actualmente el mejor con cupo: {cupo_disponible}")

            elif clases_sin_cupo_grupo:
                clases_sin_cupo_total.extend(clases_sin_cupo_grupo)
                grupos_descartados.append(grupo_inscripcion)

        if mejor_grupo:
            fila_final = mejor_filas[0]
            obs_final = observacion_extra + (f" | Cupo disponible: {menor_cupo_grupo}" if menor_cupo_grupo is not None else "")
            print(f"✔ Asignado a grupo: {mejor_grupo} con cupo mínimo: {menor_cupo_grupo}")
            resultados.append({
                'INSTITUCION': 'SISE',
                'GRUPO_INSCRIPCION': fila_final['GRUPO_INSCRIPCION'],
                'DESCRIPCION': fila_final['DESCRIPCION'],
                'ID_ALUMNO': id_alumno,
                'GRADO_ACADEMICO': 'PRGS',
                'PROGRAMA_ACADEMICO': prog_alu,
                'OBSERVACION': obs_final.strip(" |")
            })
            asignado = True
        else:
            if not asignado:
                print("✘ No asignado. Sección sin cupo. Clases sin cupo:")
                clases_por_grupo = {}

                for grupo_inscripcion, clase_id in secciones_sin_cupo_actuales:
                    if grupo_inscripcion not in clases_por_grupo:
                        clases_por_grupo[grupo_inscripcion] = set()
                    clases_por_grupo[grupo_inscripcion].add(clase_id)

                obs_final = "Sección sin cupo. Clases sin cupo:\n"
                for grupo, clases in clases_por_grupo.items():
                    clases_str = ', '.join(sorted(clases))
                    print(f"- {grupo}: {clases_str}")
                    obs_final += f"- {grupo}: {clases_str}\n"

                obs_final = obs_final.strip()

                primer_grupo = next(iter(clases_por_grupo), '')
                descripcion_final = ''
                if primer_grupo:
                    descripcion_match = prog_df[prog_df['GRUPO_INSCRIPCION'] == primer_grupo]
                    if not descripcion_match.empty:
                        descripcion_final = descripcion_match.iloc[0]['DESCRIPCION']

                resultados.append({
                    'INSTITUCION': 'SISE',
                    'GRUPO_INSCRIPCION': primer_grupo,
                    'DESCRIPCION': descripcion_final,
                    'ID_ALUMNO': id_alumno,
                    'GRADO_ACADEMICO': 'PRGS',
                    'PROGRAMA_ACADEMICO': prog_alu,
                    'OBSERVACION': obs_final
                })
                break

# Guardar resultados
resultado_df = pd.DataFrame(resultados)
resultado_df.to_excel("Alumnos_Asignados.xlsx", index=False)
print("\n✅ Proceso finalizado. Archivo generado: Alumnos_Asignados.xlsx")