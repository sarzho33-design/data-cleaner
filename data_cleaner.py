"""
Data Cleaner - Limpia y analiza archivos CSV/Excel automáticamente.

Uso:
    python data_cleaner.py ventas.csv
    python data_cleaner.py datos.xlsx
"""

import pandas as pd
import sys
import os
from datetime import datetime


def cargar_archivo(ruta):
    """Carga un CSV o Excel automáticamente según la extensión."""
    extension = os.path.splitext(ruta)[1].lower()
    if extension == ".csv":
        return pd.read_csv(ruta)
    elif extension in [".xlsx", ".xls"]:
        return pd.read_excel(ruta)
    else:
        raise ValueError(f"Formato no soportado: {extension}. Usa CSV o Excel.")


def limpiar_datos(df):
    """Limpia el dataframe: duplicados, espacios, valores vacíos."""
    filas_antes = len(df)

    # Eliminar filas completamente vacías
    df = df.dropna(how="all")

    # Eliminar duplicados
    df = df.drop_duplicates()

    # Limpiar espacios en columnas de texto
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].str.strip()

    # Estandarizar nombres de columnas
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]

    filas_despues = len(df)
    filas_eliminadas = filas_antes - filas_despues

    return df, filas_eliminadas


def analizar_datos(df):
    """Genera estadísticas básicas del dataframe."""
    stats = {}

    stats["total_filas"] = len(df)
    stats["total_columnas"] = len(df.columns)
    stats["columnas"] = list(df.columns)
    stats["valores_nulos"] = df.isnull().sum().to_dict()

    # Estadísticas de columnas numéricas
    numericas = df.select_dtypes(include="number")
    if not numericas.empty:
        stats["numericas"] = numericas.describe().to_dict()

    return stats


def generar_informe(nombre_archivo, df, stats, filas_eliminadas):
    """Genera un informe HTML con los resultados."""

    # Tabla de valores nulos
    nulos_html = ""
    for col, nulos in stats["valores_nulos"].items():
        color = "#e74c3c" if nulos > 0 else "#2ecc71"
        nulos_html += f"<tr><td>{col}</td><td style='color:{color}'>{nulos}</td></tr>"

    # Tabla de estadísticas numéricas
    stats_html = ""
    if "numericas" in stats:
        for col, valores in stats["numericas"].items():
            stats_html += f"""
            <tr>
                <td><b>{col}</b></td>
                <td>{valores.get('mean', 'N/A'):.2f}</td>
                <td>{valores.get('min', 'N/A'):.2f}</td>
                <td>{valores.get('max', 'N/A'):.2f}</td>
                <td>{valores.get('std', 'N/A'):.2f}</td>
            </tr>
            """

    # Preview de los primeros 10 registros
    preview_html = df.head(10).to_html(
        classes="preview-table", index=False, border=0
    )

    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")

    html = f"""
    <html>
    <head>
        <title>Informe - {nombre_archivo}</title>
        <meta charset="utf-8">
        <style>
            * {{ box-sizing: border-box; margin: 0; padding: 0; }}
            body {{
                font-family: 'Segoe UI', sans-serif;
                background: #f4f6f9;
                color: #333;
                padding: 30px;
            }}
            h1 {{ color: #2c3e50; margin-bottom: 5px; }}
            .fecha {{ color: #95a5a6; margin-bottom: 30px; font-size: 0.9em; }}
            .cards {{
                display: flex;
                gap: 20px;
                margin-bottom: 30px;
                flex-wrap: wrap;
            }}
            .card {{
                background: white;
                border-radius: 10px;
                padding: 20px 30px;
                flex: 1;
                min-width: 150px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.08);
                text-align: center;
            }}
            .card .numero {{
                font-size: 2.5em;
                font-weight: bold;
                color: #3498db;
            }}
            .card .etiqueta {{
                color: #95a5a6;
                font-size: 0.85em;
                margin-top: 5px;
            }}
            .card.alerta .numero {{ color: #e74c3c; }}
            .card.ok .numero {{ color: #2ecc71; }}
            .seccion {{
                background: white;
                border-radius: 10px;
                padding: 24px;
                margin-bottom: 24px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            }}
            .seccion h2 {{
                color: #2c3e50;
                margin-bottom: 16px;
                font-size: 1.1em;
                border-bottom: 2px solid #f0f0f0;
                padding-bottom: 10px;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
            }}
            th {{
                background: #f8f9fa;
                padding: 10px 14px;
                text-align: left;
                font-size: 0.85em;
                color: #666;
                text-transform: uppercase;
            }}
            td {{
                padding: 10px 14px;
                border-bottom: 1px solid #f0f0f0;
                font-size: 0.92em;
            }}
            .preview-table {{ font-size: 0.85em; }}
            .preview-table th {{ background: #3498db; color: white; }}
        </style>
    </head>
    <body>
        <h1>📊 Informe de datos: {nombre_archivo}</h1>
        <p class="fecha">Generado el {fecha}</p>

        <div class="cards">
            <div class="card">
                <div class="numero">{stats['total_filas']}</div>
                <div class="etiqueta">Filas limpias</div>
            </div>
            <div class="card">
                <div class="numero">{stats['total_columnas']}</div>
                <div class="etiqueta">Columnas</div>
            </div>
            <div class="card {'alerta' if filas_eliminadas > 0 else 'ok'}">
                <div class="numero">{filas_eliminadas}</div>
                <div class="etiqueta">Filas eliminadas</div>
            </div>
        </div>

        <div class="seccion">
            <h2>Valores nulos por columna</h2>
            <table>
                <tr><th>Columna</th><th>Valores nulos</th></tr>
                {nulos_html}
            </table>
        </div>

        {'<div class="seccion"><h2>Estadísticas numéricas</h2><table><tr><th>Columna</th><th>Media</th><th>Mínimo</th><th>Máximo</th><th>Desv. estándar</th></tr>' + stats_html + '</table></div>' if stats_html else ''}

        <div class="seccion">
            <h2>Vista previa (primeros 10 registros)</h2>
            {preview_html}
        </div>
    </body>
    </html>
    """
    return html


def main():
    if len(sys.argv) < 2:
        print("Uso: python data_cleaner.py archivo.csv")
        print("     python data_cleaner.py archivo.xlsx")
        sys.exit(1)

    ruta = sys.argv[1]

    if not os.path.exists(ruta):
        print(f"Error: no se encuentra el archivo '{ruta}'")
        sys.exit(1)

    print(f"Cargando {ruta}...")
    df = cargar_archivo(ruta)

    print("Limpiando datos...")
    df, filas_eliminadas = limpiar_datos(df)

    print("Analizando datos...")
    stats = analizar_datos(df)

    # Guardar CSV limpio
    nombre_base = os.path.splitext(os.path.basename(ruta))[0]
    csv_salida = f"{nombre_base}_limpio.csv"
    df.to_csv(csv_salida, index=False)
    print(f"✅ CSV limpio guardado: {csv_salida}")

    # Guardar informe HTML
    html_salida = f"{nombre_base}_informe.html"
    informe = generar_informe(os.path.basename(ruta), df, stats, filas_eliminadas)
    with open(html_salida, "w", encoding="utf-8") as f:
        f.write(informe)
    print(f"✅ Informe HTML guardado: {html_salida}")

    print(f"\nResumen:")
    print(f"  Filas procesadas: {stats['total_filas']}")
    print(f"  Columnas: {stats['total_columnas']}")
    print(f"  Filas eliminadas: {filas_eliminadas}")


if __name__ == "__main__":
    main()