import pandas as pd
import os

def determinar_siguiente_ciclo(ciclo_data):
    bloques_ok = {}
    jalados = 0

    ciclos = ['C1B1', 'C1B2', 'C2B1', 'C2B2', 'C3B1', 'C3B2',
              'C4B1', 'C4B2', 'C5B1', 'C5B2', 'C6B1', 'C6B2']

    for bloque in ciclos:
        val = str(ciclo_data.get(bloque, "")).strip().upper()
        if val == "OK":
            ciclo_str, b = bloque.split('B')
            ciclo_num = int(ciclo_str.replace('C', ''))
            if ciclo_num not in bloques_ok:
                bloques_ok[ciclo_num] = []
            bloques_ok[ciclo_num].append(f'B{b}')
        elif val.isdigit() and int(val) in [1, 2]:
            jalados += int(val)
            ciclo_str, b = bloque.split('B')
            ciclo_num = int(ciclo_str.replace('C', ''))
            if ciclo_num not in bloques_ok:
                bloques_ok[ciclo_num] = []
            bloques_ok[ciclo_num].append(f'B{b}')

    if not bloques_ok:
        return "Venta nueva", jalados

    ciclo_actual_int = max(bloques_ok.keys())
    bloques = bloques_ok[ciclo_actual_int]
    ciclo_actual = f'C{ciclo_actual_int}'

    if 'B1' in bloques and 'B2' in bloques:
        if ciclo_actual_int < 6:
            return f'C{ciclo_actual_int + 1}B1|B2', jalados
        else:
            return "EGRESADO", jalados
    elif 'B1' in bloques:
        return f'{ciclo_actual}B2', jalados
    elif 'B2' in bloques:
        return f'{ciclo_actual}B1', jalados
    else:
        return f'{ciclo_actual}B1', jalados

# Rutas
ruta_excel_principal = 'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\Reportes\\ingresoDeterminarCicloBloque.xlsx'
ruta_carpeta_programas = 'C:\\Users\\jramosc\\Downloads\\140725'

# Cargar archivo principal
df_principal = pd.read_excel(ruta_excel_principal, sheet_name='SISE_RPT_ALU_DEU_WV', dtype={'ID': str})
bloques_academicos = ['C1B1', 'C1B2', 'C2B1', 'C2B2', 'C3B1', 'C3B2', 'C4B1', 'C4B2', 'C5B1', 'C5B2', 'C6B1', 'C6B2']

df_matriculados = pd.read_excel(ruta_excel_principal, sheet_name='SISE_ALUMNOSMATRICULADOS', dtype={'IdAlumno': str, 'IdCurso': str})
df_notas = pd.read_excel(ruta_excel_principal, sheet_name='SISE_NOTAS_PROMED_QRY_V02', dtype={'IdAlumno': str, 'IdCurso': str})

# Resultado final
resultados = []

# Agrupar por programa
grupos_por_programa = df_principal.groupby('Programa Académico')

for programa, alumnos in grupos_por_programa:
    archivo_programa = os.path.join(ruta_carpeta_programas, f"{programa}.xlsx")

    if not os.path.exists(archivo_programa):
        for _, fila in alumnos.iterrows():
            resultados.append({
                'ID': fila['ID'],
                'Programa': programa,
                'PlanAcadémico': fila.get('Plan Académico', ''),
                'TurnoAcadémico': fila.get('Turno Académico', ''),
                'Campus': fila.get('Campus', ''),
                'Modalidad': fila.get('Modalidad', ''),
                'SiguienteCiclo': 'Archivo no encontrado',
                'TotalJalados': '',
                'Observacion': f"No se encontró archivo {programa}.xlsx"
            })
        continue

    try:
        df_programa = pd.read_excel(archivo_programa, header=7, dtype={'IdAlumno': str})
        df_programa['IdCurso'] = df_programa['IdCurso'].astype(str)

    except Exception as e:
        for _, fila in alumnos.iterrows():
            resultados.append({
                'ID': fila['ID'],
                'Programa': programa,
                'PlanAcadémico': fila.get('Plan Académico', ''),
                'TurnoAcadémico': fila.get('Turno Académico', ''),
                'Campus': fila.get('Campus', ''),
                'Modalidad': fila.get('Modalidad', ''),
                'SiguienteCiclo': 'Error al leer archivo',
                'TotalJalados': '',
                'Observacion': str(e)
            })
        continue

    for _, fila_alumno in alumnos.iterrows():
        id_alumno = fila_alumno['ID']
        df_cursos = df_programa[df_programa['IdAlumno'] == id_alumno]

        # Agregar cursos actuales del ciclo más reciente (de hoja de alumnos matriculados)
        df_actual_matriculado = df_matriculados[df_matriculados['IdAlumno'] == id_alumno]

        # Si hay cursos actuales
        if not df_actual_matriculado.empty:
            registros_actuales = []

            for _, fila_curso in df_actual_matriculado.iterrows():
                id_curso = str(fila_curso['IdCurso'])
                ciclo = str(fila_curso['Ciclo']).strip()
                bloque = str(fila_curso['Bloque']).strip()

                # Buscar calificación en la hoja de notas
                fila_nota = df_notas[(df_notas['IdAlumno'] == id_alumno) & (df_notas['IdCurso'] == id_curso)]
                calificacion = fila_nota['Calif'].values[0] if not fila_nota.empty else ''

                # Determinar estado y estado nota según la calificación
                if isinstance(calificacion, (int, float)):
                    if calificacion >= 13:
                        estado = 'Matriculado'
                        estado_nota = 'Aprobado'
                    else:
                        estado = 'Matriculado'
                        estado_nota = 'Desaprobado'
                else:
                    estado = 'Pendiente'
                    estado_nota = 'Pendiente'

                # 🖨️ Mostrar información de depuración
                print(f"🧾 Alumno {id_alumno} | Curso {id_curso} | Ciclo {ciclo}{bloque} | "
                      f"Nota: {calificacion} | Estado: {estado_nota}")

                # ✅ Agregar a registros incluyendo IdCurso
                registros_actuales.append({
                    'Ciclo': ciclo,
                    'Bloque': bloque,
                    'IdCurso': id_curso,  # 🔧 Este campo es crucial
                    'Estado': estado,
                    'EstadoNota': estado_nota,
                    'IdCicloLectivo': fila_curso.get('IdCicloLectivo', '')
                })

            # Asegúrate de que cada registro tenga Estado y EstadoNota como Aprobado
            """for r in registros_actuales:
                r['Estado'] = 'Matriculado'
                r['EstadoNota'] = 'Aprobado'"""

            # Convertir a DataFrame los cursos actuales
            df_actual = pd.DataFrame(registros_actuales)

            # Actualizar estado en df_cursos si ya estaba como Pendiente
            for _, fila_nueva in df_actual.iterrows():
                ciclo = fila_nueva['Ciclo']
                bloque = fila_nueva['Bloque']
                estado = fila_nueva['Estado']
                estado_nota = fila_nueva['EstadoNota']

                # Buscar coincidencias en df_cursos existentes
                mask = (
                        (df_cursos['IdAlumno'] == id_alumno) &
                        (df_cursos['Ciclo'] == ciclo) &
                        (df_cursos['Bloque'] == bloque) &
                        (df_cursos['Estado'] == 'Pendiente')
                )

                if mask.any():
                    df_cursos.loc[mask, 'Estado'] = estado
                    df_cursos.loc[mask, 'EstadoNota'] = estado_nota
                    print(
                        f"🔁 Actualizado {ciclo}{bloque} del alumno {id_alumno} a Estado={estado}, EstadoNota={estado_nota}")


            # Eliminar duplicados antes de agregar
            ids_actuales = set((r['Ciclo'], r['Bloque'], r['IdCurso']) for r in registros_actuales if r.get('IdCurso'))

            df_cursos = df_cursos[
                ~df_cursos.apply(lambda row: (row['Ciclo'], row['Bloque'], row['IdCurso']) in ids_actuales, axis=1)
            ]

            print("df_actual.columns:", df_actual.columns)
            print("df_cursos.columns:", df_cursos.columns)

            # Agregar los registros actualizados una vez
            df_cursos = pd.concat([df_cursos, df_actual], ignore_index=True)

            if not df_cursos.empty and 'IdAlumno' in df_cursos.columns:
                subset = df_cursos[df_cursos['IdAlumno'] == id_alumno]
                print("\n📋 DataFrame final df_cursos del alumno", id_alumno)
                print(subset[['Ciclo', 'Bloque', 'IdCurso', 'Estado', 'EstadoNota']])
            else:
                print(f"⚠️ No se encontraron filas en df_cursos para IdAlumno {id_alumno}, salto la impresión.")

        if df_cursos.empty:
            resultados.append({
                'ID': id_alumno,
                'Programa': programa,
                'PlanAcadémico': fila_alumno.get('Plan Académico', ''),
                'TurnoAcadémico': fila_alumno.get('Turno Académico', ''),
                'Campus': fila_alumno.get('Campus', ''),
                'Modalidad': fila_alumno.get('Modalidad', ''),
                'SiguienteCiclo': 'Venta nueva',
                'TotalJalados': '',
                'Observacion': 'No se encontraron cursos para el alumno'
            })
            continue

        estado_de_ciclos = {}

        for bloque in bloques_academicos:
            ciclo = bloque[:2]
            sub_bloque = bloque[2:]

            df_bloque = df_cursos[(df_cursos['Ciclo'] == ciclo) & (df_cursos['Bloque'] == sub_bloque)].copy()

            if df_bloque.empty:
                continue

            df_bloque['IdCurso'] = df_bloque['IdCurso'].astype(str)

            # Contar cursos problemáticos: pendientes o desaprobados
            # Agrupar por curso para evitar duplicados
            df_bloque_unico = df_bloque.drop_duplicates(subset=['Ciclo', 'Bloque', 'IdCurso'])
            df_bloque_unico = df_bloque_unico[df_bloque_unico['IdCurso'].notna()]  # ← FILTRAR NaN

            pendientes = df_bloque_unico[df_bloque_unico['Estado'] == 'Pendiente']
            desaprobados = df_bloque_unico[
                (df_bloque_unico['Estado'] == 'Matriculado') & (df_bloque_unico['EstadoNota'] != 'Aprobado')
                ]

            print("Cursos únicos en", bloque, df_bloque_unico['IdCurso'].tolist())
            print(f"Len Pendientes: {len(pendientes)}")
            print(f"Len desaprobados: {len(desaprobados)}")
            total_problematicos = len(pendientes) + len(desaprobados)

            if total_problematicos >= 3:
                estado_de_ciclos[bloque] = str(total_problematicos)
            elif 1 <= total_problematicos <= 2:
                estado_de_ciclos[bloque] = str(total_problematicos)
            else:
                estado_de_ciclos[bloque] = "OK"

        print(f'{id_alumno} : {estado_de_ciclos}')

        siguiente_bloque, total_jalados = determinar_siguiente_ciclo(estado_de_ciclos)

        # Determinar el ciclo lectivo más reciente
        ultimo_ciclo_lectivo = df_cursos['IdCicloLectivo'].max() if 'IdCicloLectivo' in df_cursos.columns else ''

        # 👁️ Detectar salto de ciclo (adelanto probable)
        # 🧠 Comprobación de adelanto de ciclo completo
        observacion = ''
        bloques_presentes = set(estado_de_ciclos.keys())
        ciclo_indices = {bloque: i for i, bloque in enumerate(bloques_academicos)}

        for i in range(1, len(bloques_academicos) - 1):
            bloque_anterior = bloques_academicos[i]
            bloque_siguiente = bloques_academicos[i + 1]

            if (bloque_anterior not in bloques_presentes and
                    estado_de_ciclos.get(bloque_siguiente, '') == 'OK'):
                observacion = f'Posible adelanto de ciclo (bloque {bloque_anterior} ausente)'
                break

        # 🧠 Comprobación de adelanto de ciclo completo
        for n in range(2, 7):
            prev_b1 = f'C{n - 1}B1'
            prev_b2 = f'C{n - 1}B2'
            cur_b1 = f'C{n}B1'
            cur_b2 = f'C{n}B2'

            val_prev1 = estado_de_ciclos.get(prev_b1)
            val_prev2 = estado_de_ciclos.get(prev_b2)
            val_cur1 = estado_de_ciclos.get(cur_b1)
            val_cur2 = estado_de_ciclos.get(cur_b2)

            if (val_prev1 not in ('OK', None) and val_prev2 not in ('OK', None)) \
                    and (val_cur1 == 'OK' or val_cur2 == 'OK'):
                observacion = f'Posible adelanto de ciclo completo: saltado C{n - 1}'
                break

        # 🔚 Resultado final
        resultados.append({
            'ID': id_alumno,
            'Programa': programa,
            'PlanAcadémico': fila_alumno.get('Plan Académico', ''),
            'TurnoAcadémico': fila_alumno.get('Turno Académico', ''),
            'Campus': fila_alumno.get('Campus', ''),
            'Modalidad': fila_alumno.get('Modalidad', ''),
            'SiguienteCiclo': siguiente_bloque,
            'TotalJalados': total_jalados,
            'UltCicloLectivo': ultimo_ciclo_lectivo,
            'Observacion': observacion
        })

# Exportar resultados
df_resultado = pd.DataFrame(resultados)
ruta_salida = 'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\Reportes\\resultado_ciclos.xlsx'
df_resultado.to_excel(ruta_salida, index=False)
print(f"\n✅ Proceso finalizado. Resultados guardados con nombre: resultado_ciclos.xlsx")