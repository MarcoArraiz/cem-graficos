import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

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
    "Clay": "#9BC2E6",
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
    "Phengite": "#F4B084",
    "Quartz": "#FFC000",
    "Qz-Fe": "#FFD966",
    "Qz-Ser": "#FF9900",
    "Rutile": "#757171",
    "Sericite": "#DFDA00",
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
}


def load_data():
    root = tk.Tk()
    root.withdraw()
    file_path_curva = filedialog.askopenfilename(
        title="Seleccione el archivo para Entregable Cara Curva"
    )
    file_path_plana = filedialog.askopenfilename(
        title="Seleccione el archivo para Entregable Cara Plana"
    )

    # Listar las hojas disponibles en cada archivo
    xls_curva = pd.ExcelFile(file_path_curva)
    xls_plana = pd.ExcelFile(file_path_plana)

    print(f"Hojas en {file_path_curva}: {xls_curva.sheet_names}")
    print(f"Hojas en {file_path_plana}: {xls_plana.sheet_names}")

    # Cargar las hojas correctas
    df_curva_mineralogia = pd.read_excel(file_path_curva, sheet_name="Mineralogía")
    df_curva_geoquimica = pd.read_excel(file_path_curva, sheet_name="Geoquímica")
    df_plana_mineralogia = pd.read_excel(file_path_plana, sheet_name="Mineralogía")
    df_plana_geoquimica = pd.read_excel(file_path_plana, sheet_name="Geoquímica")

    common_columns_mineralogia = df_curva_mineralogia.columns.intersection(
        df_plana_mineralogia.columns
    )
    common_columns_geoquimica = df_curva_geoquimica.columns.intersection(
        df_plana_geoquimica.columns
    )

    print("Columnas comunes Mineralogía:", common_columns_mineralogia)
    print("Columnas comunes Geoquímica:", common_columns_geoquimica)

    # Determinar cuál columna "From" usar (la más larga)
    from_curva = (
        df_curva_mineralogia["From"]
        if len(df_curva_mineralogia["From"]) >= len(df_plana_mineralogia["From"])
        else df_plana_mineralogia["From"]
    )

    return (
        (df_curva_mineralogia, df_plana_mineralogia, common_columns_mineralogia),
        (df_curva_geoquimica, df_plana_geoquimica, common_columns_geoquimica),
        from_curva,
    )


def generate_excel_graphs(
    selected_items, df_curva, df_plana, sheet_name, writer, from_column
):
    for item in selected_items:
        df = pd.DataFrame(
            {
                "From": from_column,
                "Cara Curva": df_curva[item],
                "Cara Plana": df_plana[item],
            }
        )
        df.to_excel(
            writer, sheet_name=f"{sheet_name}_{item}", index=False, startrow=0
        )  # Empezar a escribir desde la fila 0

        workbook = writer.book
        worksheet = writer.sheets[f"{sheet_name}_{item}"]
        chart = workbook.add_chart({"type": "line"})

        max_row = len(df) + 1
        chart.add_series(
            {
                "name": f"{item} plana",
                "categories": f"={sheet_name}_{item}!$A$2:$A${max_row}",
                "values": f"={sheet_name}_{item}!$C$2:$C${max_row}",
                "line": {"color": "blue"},
            }
        )
        chart.add_series(
            {
                "name": f"{item} curvo",
                "categories": f"={sheet_name}_{item}!$A$2:$A${max_row}",
                "values": f"={sheet_name}_{item}!$B$2:$B${max_row}",
                "line": {
                    "color": element_colors.get(item, "#FFA500")
                },  # Usar color definido o naranja por defecto
            }
        )

        chart.set_title({"name": " " + item})
        chart.set_x_axis(
            {
                "name": "From (m)",
                "num_font": {"rotation": -90},
                "label_position": "low",
            }
        )
        chart.set_y_axis(
            {
                "name": "%",
                "major_gridlines": {"visible": True},
                "num_format": "0.000",  # Formato para mostrar números con tres decimales, sin multiplicar por 100
            }
        )
        chart.set_legend({"position": "bottom"})

        chart.set_size({"width": 1280, "height": 274})  # Tamaño específico del gráfico

        worksheet.insert_chart(
            "F1", chart
        )  # Insertar el gráfico en la parte superior de la hoja


def select_items(
    df_curva_mineralogia,
    df_plana_mineralogia,
    common_columns_mineralogia,
    df_curva_geoquimica,
    df_plana_geoquimica,
    common_columns_geoquimica,
    from_column,
):
    master = tk.Tk()
    master.title("Selección de Minerales y Elementos")
    listbox = tk.Listbox(master, selectmode="multiple", width=50)

    # Excluir las columnas "Hole Id", "From" y "To"
    excluded_columns = {"Hole ID", "From", "To"}
    columns_to_display_mineralogia = [
        column
        for column in common_columns_mineralogia
        if column not in excluded_columns
    ]
    columns_to_display_geoquimica = [
        column for column in common_columns_geoquimica if column not in excluded_columns
    ]

    # Mostrar opciones para seleccionar desde Mineralogía
    tk.Label(master, text="Mineralogía").pack()
    for column in columns_to_display_mineralogia:
        listbox.insert("end", f"Mineralogía: {column}")
    # Mostrar opciones para seleccionar desde Geoquímica
    tk.Label(master, text="Geoquímica").pack()
    for column in columns_to_display_geoquimica:
        listbox.insert("end", f"Geoquímica: {column}")

    listbox.pack()

    submit_btn = tk.Button(
        master,
        text="Comparar y Generar",
        command=lambda: on_submit(
            master,
            listbox,
            df_curva_mineralogia,
            df_plana_mineralogia,
            common_columns_mineralogia,
            df_curva_geoquimica,
            df_plana_geoquimica,
            common_columns_geoquimica,
            from_column,
        ),
    )
    submit_btn.pack()
    master.mainloop()


def on_submit(
    master,
    listbox,
    df_curva_mineralogia,
    df_plana_mineralogia,
    common_columns_mineralogia,
    df_curva_geoquimica,
    df_plana_geoquimica,
    common_columns_geoquimica,
    from_column,
):
    selections = [listbox.get(i) for i in listbox.curselection()]
    master.destroy()

    # Separar selecciones para Mineralogía y Geoquímica
    mineralogia_selections = [
        s.split(": ")[1] for s in selections if s.startswith("Mineralogía")
    ]
    geoquimica_selections = [
        s.split(": ")[1] for s in selections if s.startswith("Geoquímica")
    ]

    current_dir = os.getcwd()
    output_path = os.path.join(current_dir, "Entregable Cara Plana y Cara Curva.xlsx")
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        if mineralogia_selections:
            generate_excel_graphs(
                mineralogia_selections,
                df_curva_mineralogia,
                df_plana_mineralogia,
                "Mineralogía",
                writer,
                from_column,
            )
        if geoquimica_selections:
            generate_excel_graphs(
                geoquimica_selections,
                df_curva_geoquimica,
                df_plana_geoquimica,
                "Geoquímica",
                writer,
                from_column,
            )

    messagebox.showinfo(
        "Proceso completado",
        "Los gráficos han sido guardados en el archivo Excel 'Entregable Cara Plana y Cara Curva'.",
    )


(
    (df_curva_mineralogia, df_plana_mineralogia, common_columns_mineralogia),
    (df_curva_geoquimica, df_plana_geoquimica, common_columns_geoquimica),
    from_column,
) = load_data()
select_items(
    df_curva_mineralogia,
    df_plana_mineralogia,
    common_columns_mineralogia,
    df_curva_geoquimica,
    df_plana_geoquimica,
    common_columns_geoquimica,
    from_column,
)
