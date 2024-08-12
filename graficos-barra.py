import pandas as pd
import xlsxwriter

# Definir colores específicos para cada mineral
element_colors = {
    "Actinolite": "#E2EFDA",
    "Albite": "#CCCC00",
    "Alunite": "#00FFCC",
    "Andesine": "#C65911",
    "Andradite": "#9999FF",
    "Anhydrite": "#FF66CC",
    "Ankerite": "#3366CC",
    "Amphibole": "#8C941E",
    "Barite": "#666699",
    "Biotite": "#984807",
    "Calcite": "#66FFFF",
    "Apatite": "#3333FF",
    "Carbonate-fluorapatite": "#33CCCC",
    "Chlorite": "#548235",
    "Clay_Ca": "#9BC2E6",
    "Clay_Mg": "#9BC2E6",
    "Clay": "#9BC2E6",
    "Delafossite": "#DC47E7",
    "Diopside": "#00AA7F",
    "Dolomite": "#0099FF",
    "Dravite": "#9933FF",
    "Epidote": "#00FF00",
    "Gree-Gray Sericite": "#F8CBAD",
    "Gypsum": "#FFFFCC",
    "Halloysite": "#66FFFF",
    "Illite": "#8DADF5",
    "Ilmenite": "#808080",
    "Indeterminate Clay": "#203764",
    "Kaolinite": "#99CCFF",
    "K-Feldespar": "#CB2198",
    "Montmorillonite": "#B9CDE5",
    "Muscovite": "#FFFF00",
    "NoData": "#FFFFFF",
    "NoData2": "#FFFFFF",
    "NoData3": "#FFFFFF",
    "Orthoclase": "#CB2198",
    "Plagioclase": "#A50021",
    "Plagioclase Ca": "#A50021",
    "Phengite": "#F4B084",
    "Quartz": "#FFC000",
    "Qz-Fe": "#FFD966",
    "Qz-Ser": "#FF9900",
    "Rutile": "#757171",
    "Sericite": "#DFDA00",
    "Sericite Ti": "#DFDA00",
    "SiNa": "#BF8F00",
    "Smithsonite": "#5F5F5F",
    "Smectite": "#65CDA9",
    "Titanite": "#C0C0C0",
    "Tremolite": "#92D050",
    "Wollastonita": "#CAAFFF",
    "AgCu": "#BFBFBF",
    "Atacamite": "#00FFFF",
    "Black-Copper": "#003300",
    "Bornite": "#FF66CC",
    "Brochantite": "#269960",
    "Chalcocite": "#0066FF",
    "Chalcopyrite": "#FF9933",
    "Chenevixite": "#C3E900",
    "Chrysocolla": "#53C9C7",
    "Copper-pitch": "#AEAAAA",
    "Copper-wad": "#757171",
    "Fe-Hidroxide": "#FFF2CC",
    "Galena": "#8EA9DB",
    "Goethite": "#996633",
    "Goethite-Cu": "#CC9900",
    "Hematite": "#996600",
    "Jarosite": "#663300",
    "Magnetite": "#000000",
    "Malachite": "#339933",
    "Molybdenite": "#8EA9DB",
    "Pyrite": "#FFFF00",
    "Pyrolusite": "#C0C0C0",
    "Sphalerite": "#666699",
    "Stromeyerite": "#FF0000",
    "Tennantite": "#7030A0",
    "Enargite": "#7030A0",
    "Tetrahedrite": "#9966FF",
    "Rutilo": "#747474",
    "Ferrimolybdite": "#8EA9DB",
    "Turquoise": "#9CE0DE",
    "Malachite-Azurite": "#339933",
}

def generar_grafico_barra(df, nombre_archivo, filtrar_ceros=False):
    # Generar lista de metros desde el inicio hasta el final
    min_metros = int(df["From"].min())
    max_metros = int(df["To"].max())
    all_metros = list(range(min_metros, max_metros + 1, 2))

    # Crear DataFrame completo con todos los metros y rellenar con 0 donde falte
    complete_df = pd.DataFrame({"From": all_metros})
    complete_df = pd.merge(complete_df, df, on="From", how="left").fillna(0)

    # Excluir las columnas no deseadas
    columnas = complete_df.columns[3:]
    columnas = [
        col for col in columnas if col not in ["TOTAL", "Litología", "Zona Mineral"]
    ]

    nombres = {
        col: f"'{col}'" if " " in col else col for col in columnas
    }  # Asegurarse de que los nombres de hoja están entre comillas si contienen espacios

    profundidad = complete_df["From"].values
    with xlsxwriter.Workbook(nombre_archivo) as workbook:
        bold_centered = workbook.add_format(
            {"bold": True, "align": "center", "valign": "vcenter", "font_size": 12}
        )
        centered = workbook.add_format(
            {"align": "center", "valign": "vcenter", "font_size": 12}
        )
        title_format = workbook.add_format({"font_size": 20, "bold": True})

        resumen_worksheet = workbook.add_worksheet("Resumen Gráficos")
        resumen_worksheet.set_column("A:A", 30)
        resumen_row = 1

        for idx, columna in enumerate (columnas, start=1):
            # Filtrar las columnas con todos los valores iguales a 0 si se solicita
            if filtrar_ceros and complete_df[columna].sum() == 0:
                continue

            valores = complete_df[columna].values
            unidad = "%"

            worksheet = workbook.add_worksheet(
                nombres[columna].strip("'")
            )  # Eliminar las comillas para el nombre real de la hoja
            worksheet.write("A1", "From (m)", bold_centered)
            worksheet.write("B1", unidad, bold_centered)
            worksheet.write_column("A2", profundidad, centered)
            worksheet.write_column("B2", valores, centered)

            chart = workbook.add_chart({"type": "column"})
            chart_color = element_colors.get(
                columna, "#000000"
            )  # Obtener color o negro por defecto
            chart.add_series(
                {
                    "categories": f"={nombres[columna]}!$A$2:$A${len(profundidad)+1}",
                    "values": f"={nombres[columna]}!$B$2:$B${len(profundidad)+1}",
                    "name": nombres[columna],
                    "marker": {"type": "circle", "size": 6},
                    "fill": {"color": chart_color},
                    "line": {"color": chart_color},
                }
            )
            chart.set_title({"name": columna, "name_font": {"size": 20, "bold": True}})
            chart.set_x_axis(
                {
                    "name": "From (m)",
                    "num_font": {"rotation": -90},
                    "label_position": "low",
                }
            )
            chart.set_y_axis({"name": unidad, "name_font": {"size": 12}})
            chart.set_legend({"none": True})
            chart.set_size({"width": 1280, "height": 274})
            worksheet.insert_chart("D2", chart)

            # Crear una copia del gráfico para la hoja de resumen
            resumen_chart = workbook.add_chart({"type": "column"})
            resumen_chart.add_series(
                {
                    "categories": f"={nombres[columna]}!$A$2:$A${len(profundidad)+1}",
                    "values": f"={nombres[columna]}!$B$2:$B${len(profundidad)+1}",
                    "name": nombres[columna],
                    "marker": {"type": "circle", "size": 6},
                    "fill": {"color": chart_color},
                    "line": {"color": chart_color},
                }
            )
            resumen_chart.set_title(
                {"name": columna, "name_font": {"size": 20, "bold": True}}
            )
            resumen_chart.set_x_axis(
                {
                    "name": "From (m)",
                    "num_font": {"rotation": -90},
                    "label_position": "low",
                }
            )
            resumen_chart.set_y_axis({"name": unidad, "name_font": {"size": 12}})
            resumen_chart.set_legend({"none": True})
            resumen_chart.set_size({"width": 1280, "height": 274})

            # Insertar el gráfico en la hoja de resumen
            resumen_worksheet.write(resumen_row, 0, columna, title_format)
            resumen_worksheet.insert_chart(resumen_row, 1, resumen_chart)
            resumen_row += 15  # Espacio para el siguiente gráfico en la hoja de resumen


# Cargar datos desde el archivo Excel
archivo_excel = (
    "./Entregable Proyecto.xlsx"  # Ajusta esta ruta según la ubicación de tu archivo
)
df_geoquimica = pd.read_excel(archivo_excel, sheet_name="Geoquímica")
df_mineralogia = pd.read_excel(archivo_excel, sheet_name="Mineralogía")

# Generar gráficos para Geoquímica, excluyendo elementos con todos sus valores en 0
generar_grafico_barra(
    df_geoquimica, "Grafico_Barra_Geoquimica.xlsx", filtrar_ceros=True
)

# Generar gráficos para Mineralogía, sin filtrar valores 0
generar_grafico_barra(df_mineralogia, "Grafico_Barra_Mineralogia.xlsx")
