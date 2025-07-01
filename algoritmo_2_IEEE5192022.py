''' Prueba preliminar de cumplimiento de IEEE-519-2022 para un PCC bajo analisis
Estudiante: Ana Sofia Barrantes Mena, C01012
Curso: Calidad de la energia
Semestre I, 2025 '''

import pandas as pd
import matplotlib.pyplot as plt
import re
from typing import List

# Declaracion de variables, editar acorde PCC en analisis

I_L = 1120
I_SC = 56000
scr = I_SC/I_L
bus_voltage_PCC = 277

# Editar paths de archivos a como sea requerido
paths_armonicos_voltaje = [ 
        "reportes/testReporte Tabular_ARMÓNICAS VOLTAJE FASE 1.xls",
        "reportes/testReporte Tabular_ARMÓNICAS VOLTAJE FASE 2.xls",
        "reportes/testReporte Tabular_ARMÓNICAS VOLTAJE FASE 3.xls"]
paths_armonicos_corriente = [
        "reportes/testReporte Tabular_ARMÓNICAS CORRIENTE FASE 1.xls",
        "reportes/testReporte Tabular_ARMÓNICAS CORRIENTE FASE 2.xls",
        "reportes/testReporte Tabular_ARMÓNICAS CORRIENTE FASE 3.xls"]

path_thd_reporte = "reportes/testInforme Tabular_THD CORRIENTE Y VOLTAJE.xls"


def extraer_numero_armonico(nombre_columna: str) -> int:
    """Extrae el número de armónico desde el nombre de la columna."""
    match = re.search(r"Harm\s+(\d+)", nombre_columna)
    return int(match.group(1)) if match else None


def procesar_archivo_armonicos(path_archivo: str):
    """
    Procesa un archivo individual de armónicos
    """
    df = pd.read_excel(path_archivo, sheet_name=0)
    df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], dayfirst=True)
    col_fecha = df.columns[0]

    # Detectar la columna de la fundamental (armónico 1)
    col_fundamental = [col for col in df.columns if re.search(r"Harm\s+1", str(col))][0]

    # Filtrar columnas con "Harm" en el nombre
    columnas_armonicos = [col for col in df.columns if "Harm" in str(col)]
    armonico_data = []

    for col in columnas_armonicos:
        num_arm = extraer_numero_armonico(str(col))
        if num_arm and num_arm != 1:
            df[col] = pd.to_numeric(df[col], errors='coerce')
            porcentaje = (df[col] / df[col_fundamental]) * 100
            armonico_data.append((num_arm, porcentaje))

    result_df = pd.DataFrame({col_fecha: df[col_fecha]})
    for num_arm, porcentaje_series in armonico_data:
        result_df[f"H{num_arm} (%)"] = porcentaje_series

    return result_df


def analizar_tres_fases(paths: List[str], nombres_fases: List[str]):
    """
    Ejecuta el análisis de 3 archivos (uno por fase), grafica los porcentajes
    de armónicos individuales respecto a la fundamental.
    """
    dfs_resultado = []
    for path, nombre_fase in zip(paths, nombres_fases):
        df_analizado = procesar_archivo_armonicos(path)
        df_analizado['Fase'] = nombre_fase
        dfs_resultado.append(df_analizado)

    df_total = pd.concat(dfs_resultado)
    return df_total


def graficar_armonicos(df_total):
    columnas_armonicos = [col for col in df_total.columns if col.startswith("H")]
    col_fecha = df_total.columns[0]
    fases = df_total['Fase'].unique()

    fig, axs = plt.subplots(1, 3, figsize=(22, 6), sharey=True)

    for i, fase in enumerate(fases):
        df_fase = df_total[df_total['Fase'] == fase]
        ax = axs[i]
        for armonico in columnas_armonicos:
            ax.plot(df_fase[col_fecha], df_fase[armonico], label=armonico, linewidth=1)
        ax.axhline(3, color='red', linestyle='--', label='Límite 3%')
        ax.set_title(f"Armónicos - {fase}")
        ax.set_xlabel("Fecha")
        if i == 0:
            ax.set_ylabel("Porcentaje respecto a H1 (%)")
        ax.legend(fontsize=7, ncol=2)
        ax.grid(True)

    plt.tight_layout()
    plt.show()


def calcular_TDD_limit(scr):
    ''' Calcula el limite de TDD basado en el SRC en el PCC.
    Limites obtenidos de Tabla 2 de Std 2022 '''

    tdd_limit = 0

    if scr <= 20:
        tdd_limit = 5
    elif scr <= 50:
        tdd_limit = 8
    elif scr <= 100:
        tdd_limit = 12
    elif scr <= 1000:
        tdd_limit = 15
    else:
        tdd_limit = 20

    return tdd_limit


def calcular_THDV_limit(bus_voltage_PCC):
    ''' Calcula el limite de THD-V basado en el bus voltage en el PCC. 
    Limites obtenidos de Tabla 1 de Std 2022 '''

    thdv_limit = 0

    if bus_voltage_PCC <= 1000:
        thdv_limit = 8
    elif bus_voltage_PCC <= 69000:
        thdv_limit = 5
    elif bus_voltage_PCC <= 161000:
        thdv_limit = 2.5
    else:
        thdv_limit = 1.5

    return thdv_limit


def leer_THDV_file(path_archivo, thdv_limit):
    df = pd.read_excel(path_archivo, sheet_name=0)

    # Limpiar nombres de columnas
    df.columns = [col.strip().replace('\n', ' ') for col in df.columns]

    # Detectar columna de fecha/hora (primera)
    col_fecha = df.columns[0]
    df[col_fecha] = pd.to_datetime(df[col_fecha], dayfirst=True)

    # Detectar todas las columnas relevantes de THD
    columnas_thd = [col for col in df.columns if "THD de tensión en V" in col]

    # Crea una columna que diga si el límite fue excedido en cualquier fase
    df['Excede THD-V'] = df[columnas_thd].gt(thdv_limit).any(axis=1)

    # Filtrar filas con violación
    df_excede = df[df['Excede THD-V']]

    # Total de violaciones
    cantidad_excedidas = df_excede.shape[0]
    total_intervalos = 0

    # Imprimir intervalos de violaciones siempre que existan
    if cantidad_excedidas != 0:
        print("❗ Se ha excedido el limite de THD-V en " + str(cantidad_excedidas) + " intervalos.")
        fechas_excedidas = df_excede[col_fecha].tolist()
        print("Fecha y hora de violaciones al limite de THD-V:\n")
        print(fechas_excedidas)
        # Calcular total de intervalos posibles
        total_intervalos = df.shape[0] - 1
        # Porcentaje
        porcentaje_violado = round((cantidad_excedidas / total_intervalos) * 100, 2)
        print("\nEl porcentaje de intervalos excediendo el limite para THD-V es: " + str(porcentaje_violado))
        if porcentaje_violado <= 1:
            print("\n✅ El criterio de THD-V cumple con la norma IEEE-519-2022.")
            cumple = True
        else:
            print("\n❌ El criterio de THD-V NO cumple con la norma IEEE-519-2022.")
            cumple = False
        
        columna_fecha = "Fecha y hora"
        # Editar dependiendo del reporte
        columnas_thdv = [
            "SUBESTACION.PRINCIPAL THD de tensión en V1 alta (%)",
            "SUBESTACION.PRINCIPAL THD de tensión en V1 media (%)",
            "SUBESTACION.PRINCIPAL THD de tensión en V2 alta (%)",
            "SUBESTACION.PRINCIPAL THD de tensión en V2 media (%)",
            "SUBESTACION.PRINCIPAL THD de tensión en V3 alta (%)",
            "SUBESTACION.PRINCIPAL THD de tensión en V3 media (%)"
        ]
        print("Graficando THD-V durante el periodo de medicion...")
        graficar_thdv(df, columna_fecha, columnas_thdv, thdv_limit)
    else:
        print("No hay intervalos excediendo limite de THD-V.\n Por ende, el criterio de THD-V cumple con la norma IEEE-519-2022")
        cumple = True

    return cumple


def graficar_thdv(df, col_fecha, columnas_thdv, thdv_limit):
    df[col_fecha] = pd.to_datetime(df[col_fecha], dayfirst=True)
    plt.figure(figsize=(16, 7))

    colores = ['blue', 'green', 'orange', 'purple', 'brown', 'cyan']

    for i, col in enumerate(columnas_thdv):
        plt.plot(df[col_fecha], df[col], label=col, color=colores[i % len(colores)], linewidth=1)

        # Puntos de exceso
        excesos = df[df[col] > thdv_limit]
        plt.scatter(excesos[col_fecha], excesos[col], color='red', s=10, zorder=5)

    plt.axhline(thdv_limit, color='gray', linestyle='--', label=f'Límite {thdv_limit}%')
    plt.title("THD-V durante periodo de medicion")
    plt.xlabel("Fecha")
    plt.ylabel("THD-V (%)")
    plt.legend(loc='upper right', fontsize=8)
    plt.grid(True)
    plt.tight_layout()
    plt.show()


def armonicos_voltage_ind(path_archivo):
    """
    Verifica si algún valor de armónicos supera el 3% de la magnitud base (columna 2).
    """
    # Leer archivo
    df = pd.read_excel(path_archivo, sheet_name=0)

    # primera columna es fecha y la segunda es la fundamental
    col_fecha = df.columns[0]
    col_fundamental = df.columns[1]
    columnas_armonicos = df.columns[2:]

    # Convertir todas las columnas a numérico
    for col in df.columns[1:]:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # Asegurarse de que la fecha esté en formato datetime
    df[col_fecha] = pd.to_datetime(df[col_fecha], dayfirst=True)

    # Verificar si algún armónico excede el 3% de la fundamental
    # Para eso, comparamos cada columna de armónicos con 0.03 * fundamental
    limites = df[col_fundamental] * 0.03
    mask = (df[columnas_armonicos].gt(limites, axis=0)).any(axis=1)

    # Filtrar violaciones
    df_excede = df[mask]

    # Total de violaciones
    cantidad_excedidas = df_excede.shape[0]

    # Imprimir resultados
    if cantidad_excedidas == 0:
        print("✅ No se encontraron violaciones. Todos los armónicos están dentro del 3% permitido.")
        return True
    else:
        print(f"⚠ Se han encontrado {cantidad_excedidas} intervalos con violaciones.")

        fechas_excedidas = df_excede[col_fecha].tolist()
        print("\n🕒 Fechas y horas de violaciones:")
        for f in fechas_excedidas:
            print(f"- {f}")

        # Cálculo del total de intervalos estimado
        total_intervalos = df.shape[0] - 1

        # Porcentaje de violaciones
        porcentaje_violado = round((cantidad_excedidas / total_intervalos) * 100, 2)
        print(f"\n📊 Porcentaje de intervalos con violaciones: {porcentaje_violado}%")

        if porcentaje_violado <= 1:
            print("✅ Cumple con la norma IEEE-519-2022 para armónicos individuales de voltaje.")
            return True
        else:
            print("❌ NO cumple con la norma IEEE-519-2022 para armónicos individuales de voltaje.")
            return False


def calcular_tdd_corriente(path_archivo: str, tdd_limit: float, nombre_fase: str):
    df = pd.read_excel(path_archivo)
    df.columns = [col.strip().replace('\n', ' ') for col in df.columns]

    col_fecha = df.columns[0]
    df[col_fecha] = pd.to_datetime(df[col_fecha], dayfirst=True)

    col_fund = [col for col in df.columns if "Harm 1" in col][0]
    cols_arm = [col for col in df.columns if "Harm" in col and "Harm 1" not in col]

    df[col_fund] = pd.to_numeric(df[col_fund], errors='coerce')
    df[cols_arm] = df[cols_arm].apply(pd.to_numeric, errors='coerce')

    df['TDD'] = (df[cols_arm]**2).sum(axis=1).pow(0.5) / I_L * 100
    df['Excede TDD'] = df['TDD'] > tdd_limit
    df_excede = df[df['Excede TDD']]

    cantidad_excedidas = df_excede.shape[0]
    fechas_excedidas = df_excede[col_fecha].tolist()
    total_intervalos = df.shape[0] - 1
    porcentaje_violado = round((cantidad_excedidas / total_intervalos) * 100, 2)
    cumple = porcentaje_violado <= 1

    if cantidad_excedidas:
        print(f"❗ Se ha excedido el límite de TDD en {cantidad_excedidas} intervalos.")
        print("Fechas y horas de violaciones:")
        print(fechas_excedidas)
        print(f"Porcentaje de intervalos violados: {porcentaje_violado}%")
        if cumple:
            print("✅ Cumple con la norma IEEE-519-2022.")
        else:
            print("❌ NO cumple con la norma IEEE-519-2022.")
    else:
        print("✅ No hay violaciones del límite de TDD. Cumple con la norma.")

    df['Fase'] = nombre_fase
    return df[[col_fecha, 'TDD', 'Fase']], cumple


def graficar_tdd_tres_fases(df_tdd_total):
    plt.figure(figsize=(15, 6))
    for fase in df_tdd_total['Fase'].unique():
        df_fase = df_tdd_total[df_tdd_total['Fase'] == fase]
        plt.plot(df_fase.iloc[:, 0], df_fase['TDD'], label=fase)
    plt.axhline(8, color='red', linestyle='--', label='Límite 8%')
    plt.title("Distorsión total de demanda durante período de medición")
    plt.xlabel("Fecha")
    plt.ylabel("TDD (%)")
    plt.legend()
    plt.grid(True)
    plt.tight_layout()
    plt.show()


def criterio1(path):
    # Se obtiene el limite de THD V permitido para el voltaje de bus en el PCC
    print("\n🔎 PRIMER CRITERIO: ANÁLISIS DE THD-V \n")
    thdv_limit = calcular_THDV_limit(bus_voltage_PCC)
    print("\nEl limite de THD-V para este PCC es: "+ str(thdv_limit) + "%")
    cumple_primer_criterio = leer_THDV_file(path, thdv_limit)
    return cumple_primer_criterio


def criterio2(paths):
    print("\n🔎 SEGUNDO CRITERIO: ANÁLISIS DE ARMÓNICOS INDIVIDUALES DE VOLTAJE\n")
    cumplimiento_por_fase = []
    fases_incumplidas = []
    cont = 1
    for path in paths:
        print("\n💠 Analisis para Fase " + str(cont) + "\n")
        cumple_fase = armonicos_voltage_ind(path)
        cumplimiento_por_fase.append(cumple_fase)
        if not cumple_fase:
            fases_incumplidas.append(f"Fase {cont}")
        cont += 1

    nombres_fases = ["Fase 1", "Fase 2", "Fase 3"] 
    df_final = analizar_tres_fases(paths, nombres_fases)
    graficar_armonicos(df_final)

    # Verificación general
    if all(cumplimiento_por_fase):
        print("\n✅ Todas las fases cumplen con la norma IEEE-519-2022 para armónicos individuales de voltaje. ✅")
        return True
    else:
        print("\n❌ Las siguientes fases NO cumplen con la norma IEEE-519-2022 para armónicos individuales de voltaje:")
        for fase in fases_incumplidas:
            print(f"- {fase}")
            return False


def criterio3(paths):
    print("\n🔎 TERCER CRITERIO: ANÁLISIS DE TDD\n")
    tdd_limit = calcular_TDD_limit(scr)
    print("\nEl limite de TDD para este PCC es: "+ str(tdd_limit) + "%")
    cumplimiento_por_fase = []
    fases_incumplidas = []
    cont = 1
    df_total = pd.DataFrame()
    for path in paths:
        print("\n💠 Analisis para Fase " + str(cont) + "\n")
        nombre_fase = "Fase" + str(cont)
        df, cumple_fase = calcular_tdd_corriente(path, tdd_limit, nombre_fase)
        df_total = pd.concat([df_total, df])
        cumplimiento_por_fase.append(cumple_fase)
        if not cumple_fase:
            fases_incumplidas.append(f"Fase {cont}")
        cont += 1

    graficar_tdd_tres_fases(df_total)

    # Verificación general
    if all(cumplimiento_por_fase):
        print("\n✅ Todas las fases cumplen con la norma IEEE-519-2022 para distorsión total de demanda.")
        return True
    else:
        print("\n❌ Las siguientes fases NO cumplen con la norma IEEE-519-2022 para distorsión total de demanda:")
        for fase in fases_incumplidas:
            print(f"- {fase}")
            return False


def main():
    cumple_1 = criterio1(path_thd_reporte)
    cumple_2 = criterio2(paths_armonicos_voltaje)
    cumple_3 = criterio3(paths_armonicos_corriente)

    if (cumple_1 and cumple_2 and cumple_3):
        print("\n✅ El PCC cumple preliminarmente con la norma IEEE-519-2022.")
    else:
        print("\n❌ El PCC NO cumple con la norma IEEE-519-2022.")

    return


main()
