# Script para extraer y simplificar datos de POA desde archivos CSV o Excel

import pandas as pd
import re
import argparse

def extraer_poa_simple(path, tipo='excel', hoja=None):
    if tipo == 'csv':
        df = pd.read_csv(path, header=None, encoding='latin1', delimiter=';')
    elif tipo == 'excel':
        df = pd.read_excel(path, sheet_name=hoja, header=None)
    else:
        raise ValueError("Tipo de archivo no soportado")

    # Detectar subtemas y actividades con su índice de fila
    subtema_map = {}
    actividad_map = {}
    total_por_actividad_map = {}

    for i in df.index:
        celda = str(df.at[i, 3]) if pd.notna(df.at[i, 3]) else ""
        if re.match(r"\(\d+\)", celda):
            actividad_map[i] = celda
            total_por_actividad_map[celda] = df.at[i, 9] if pd.notna(df.at[i, 9]) else None
        elif re.match(r"\d\.\d\s", celda):
            subtema_map[i] = celda

    # Combinar subtema + actividad por fila
    contexto = {}
    last_subtema = ""
    last_subtema_id = ""
    last_actividad = ""

    for i in df.index:
        if i in actividad_map:
            last_actividad = actividad_map[i]
        if i in subtema_map:
            last_subtema = subtema_map[i]
            match = re.match(r"(\d\.\d)", last_subtema)
            last_subtema_id = match.group(1) if match else ""

        contexto[i] = {
            "Actividad": last_actividad,
            "Subtema": last_subtema,
            "ID Subtema": last_subtema_id
        }

    # Detectar encabezados de fechas
    fecha_row_idx = df[df.apply(lambda r: r.astype(str).str.contains("PROGRAMACIÓN DE EJECUCIÓN", na=False)).any(axis=1)].index
    headers = df.loc[fecha_row_idx[0] + 1, 13:22].tolist() if not fecha_row_idx.empty else []

    # Filtrar filas válidas con datos reales
    patrones_invalidos = ["DETALLE", "ITEM", "CANTIDAD", "UNITARIO", "TOTAL", "TAREA", "RUBROS"]
    resultados = []
    subtema_seen = set()

    for i in df.index:
        row = df.loc[i]
        if (
            pd.notna(row[4]) and
            pd.notna(row[6]) and
            pd.notna(row[7]) and
            pd.notna(row[8]) and
            not any(pat in str(row[4]).upper() for pat in patrones_invalidos)
        ):
            fechas_dict = {
                str(h).strip(): str(v).strip() for h, v in zip(headers, row[13:13+len(headers)].tolist())
                if pd.notna(v) and pd.notna(h)
                and str(v).strip().lower() != 'nan'
                and not str(v).strip().lower().startswith('suman')
                and not str(v).strip().lower().startswith('total')
            }
            fechas_str = ', '.join([f"{k}: {v}" for k, v in fechas_dict.items()])

            actividad_actual = contexto[i]["Actividad"]
            subtema_actual = contexto[i]["Subtema"]

            mostrar_total_actividad = subtema_actual not in subtema_seen
            subtema_seen.add(subtema_actual)
            total_por_actividad = total_por_actividad_map.get(actividad_actual) if mostrar_total_actividad else None

            fila = {
                "Actividad": actividad_actual,
                "ID Subtema": contexto[i]["ID Subtema"],
                "Subtema": subtema_actual,
                "Descripción o Detalle": row[4],
                "Cantidad (Meses de contrato)": row[6],
                "Precio Unitario": row[7],
                "Total": row[8],
                "Total por Actividad": total_por_actividad,
                "Programación Ejecución": fechas_str
            }
            resultados.append(fila)

    columnas_ordenadas = [
        "Actividad", "ID Subtema", "Subtema", "Descripción o Detalle",
        "Cantidad (Meses de contrato)", "Precio Unitario", "Total",
        "Total por Actividad", "Programación Ejecución"
    ]

    df_resultado = pd.DataFrame(resultados)[columnas_ordenadas]
    return df_resultado

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extraer y simplificar POA desde archivo")
    parser.add_argument("archivo", help="Ruta del archivo .xlsx o .csv")
    parser.add_argument("--tipo", default="excel", choices=["excel", "csv"], help="Tipo de archivo")
    parser.add_argument("--sheet", help="Nombre de la hoja (solo si es Excel)")
    parser.add_argument("--out", default="resultado_simplificado.csv", help="Nombre del CSV de salida")
    args = parser.parse_args()

    df_final = extraer_poa_simple(args.archivo, tipo=args.tipo, hoja=args.sheet)
    df_final.to_csv(args.out, index=False)
    print(f"Archivo generado: {args.out}")
