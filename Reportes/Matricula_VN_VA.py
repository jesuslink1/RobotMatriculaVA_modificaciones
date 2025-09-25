import pandas as pd

# Ruta del archivo Excel
archivo = 'C:\\Users\\jramosc\\Desktop\\Proyectos\\RobotMatriculaVA_modificaciones\\Reportes\\ingresoData.xlsx'

# Cargar hojas del archivo
df_principal = pd.read_excel(archivo, sheet_name='SISE_RPT_ALU_DEU_WV', dtype={'ID': str})
df_egresados = pd.read_excel(archivo, sheet_name='SISE_SR_EGRESADOS', dtype={'Id_Alumno': str})

# Lista ordenada de ciclos académicos
ciclos = ['C1B1', 'C1B2', 'C2B1', 'C2B2', 'C3B1', 'C3B2',
          'C4B1', 'C4B2', 'C5B1', 'C5B2', 'C6B1', 'C6B2']

# Nueva función basada en última B1/B2 con estado OK
def determinar_siguiente_ciclo(ciclo_data):
    bloques_ok = {}
    jalados = 0

    print("\n--- Análisis de alumno ---")
    print("Datos de ciclo recibidos:", ciclo_data)

    for bloque in ciclos:
        val = str(ciclo_data.get(bloque, "")).strip().upper()
        print(f"Bloque: {bloque} => Valor: {val}")

        if val == "OK":
            ciclo_str, b = bloque.split('B')
            ciclo_num = int(ciclo_str.replace('C', ''))
            if ciclo_num not in bloques_ok:
                bloques_ok[ciclo_num] = []
            bloques_ok[ciclo_num].append(f'B{b}')
            print(f"✔️ Aprobado: Ciclo {ciclo_num}, Bloque B{b}")
        elif val.isdigit() and int(val) in [1, 2]:
            jalados += int(val)
            print(f"⚠️ Jalado tolerado en {bloque}. Se suma {val}. Total jalados: {jalados}")
        # elif val.isdigit() and int(val) >= 3:
            # print(f"⛔ Curso pendiente detectado en {bloque}, se detiene el análisis.")
            # return bloque, jalados

    print("📊 Bloques OK detectados:", bloques_ok)
    print("🔢 Total jalados:", jalados)

    if not bloques_ok:
        print("🆕 Alumno nuevo sin ciclos cursados.")
        return "C1B1 o C1B2", jalados

    ciclo_actual_int = max(bloques_ok.keys())
    bloques = bloques_ok[ciclo_actual_int]
    ciclo_actual = f'C{ciclo_actual_int}'
    print(f"📍 Último ciclo con bloques OK: {ciclo_actual}, bloques: {bloques}")

    if 'B1' in bloques and 'B2' in bloques:
        if ciclo_actual_int < 6:
            print("✅ Ciclo completo. Se pasa al siguiente ciclo.")
            return f'C{ciclo_actual_int + 1}B1|B2', jalados
        else:
            print("🎓 EGRESADO")
            return "EGRESADO", jalados
    elif 'B1' in bloques:
        print("➡️ Solo B1 cursado, debe cursar B2")
        return f'{ciclo_actual}B2', jalados
    elif 'B2' in bloques:
        print("➡️ Solo B2 cursado, debe cursar B1")
        return f'{ciclo_actual}B1', jalados
    else:
        print("⚠️ Ciclo con bloques OK, pero sin B1 ni B2 (caso raro)")
        return f'{ciclo_actual}B1', jalados

# Simulación: actas cerradas
actas_cerradas = False

if not actas_cerradas:
    print("\nRESULTADOS DE ANÁLISIS DE CICLO A CURSAR:\n")
    for _, row in df_principal.iterrows():
        id_alumno = row['ID']
        nombre = row['Nombre']
        registros = df_egresados[df_egresados['Id_Alumno'] == id_alumno]

        if registros.empty:
            print(f"{nombre} (ID {id_alumno}) - No se encontró avance académico")
        else:
            avance = registros.iloc[0][ciclos].to_dict()
            siguiente_ciclo, cantidad_jalados = determinar_siguiente_ciclo(avance)
            print(f"{nombre} (ID {id_alumno}) debe estudiar: {siguiente_ciclo} | Cursos jalados: {cantidad_jalados}")