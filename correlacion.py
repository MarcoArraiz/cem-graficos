import tkinter as tk
from tkinter import filedialog
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np


def load_and_process_data(file_path, sheet_name):
    data = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=2)
    data = data.loc[:, ~data.columns.str.contains("Unnamed")]
    data = data.loc[:, ~data.columns.str.contains("TOTAL")]
    data = data.loc[:, ~data.columns.str.contains("Metros")]
    data.replace("", pd.NA, inplace=True)
    data = data.apply(pd.to_numeric, errors="coerce")
    return data


def analyze_correlations(correlation_matrix):
    upper = correlation_matrix.where(
        np.triu(np.ones(correlation_matrix.shape), k=1).astype(np.bool_)
    )
    high_corr = upper.stack().loc[lambda x: x > 0.75]
    mid_corr = upper.stack().loc[lambda x: (x > 0.5) & (x <= 0.75)]
    low_corr = upper.stack().loc[lambda x: x <= 0.5]
    return high_corr, mid_corr, low_corr


def plot_and_save_correlation_matrix(data, output_excel_path, sheet_name):
    correlation_matrix = data.corr()
    high_corr, mid_corr, low_corr = analyze_correlations(correlation_matrix)

    with pd.ExcelWriter(output_excel_path, engine="openpyxl", mode="w") as writer:
        correlation_matrix.to_excel(writer, sheet_name=sheet_name)
        workbook = writer.book

        def save_and_format_sheet(data_frame, sheet_title, workbook):
            data_frame.to_excel(writer, sheet_name=sheet_title, index=True)
            worksheet = workbook[sheet_title]
            for col in ["A", "B", "C"]:
                worksheet.column_dimensions[col].width = 20

        save_and_format_sheet(
            pd.DataFrame({"Correlación Alta ( > 0.75)": high_corr}),
            "Correlación Alta",
            workbook,
        )
        save_and_format_sheet(
            pd.DataFrame({"Correlación Media (0.5 - 0.75)": mid_corr}),
            "Correlación Media",
            workbook,
        )
        save_and_format_sheet(
            pd.DataFrame({"Correlación Baja ( <= 0.5)": low_corr}),
            "Correlación Baja",
            workbook,
        )


def open_file_dialog():
    file_path = filedialog.askopenfilename(
        title="Select a file",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")),
    )
    if file_path:
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, file_path)


def process_file():
    file_path = file_path_entry.get()
    if file_path:
        mineralogia_data = load_and_process_data(file_path, "Mineralogía")
        plot_and_save_correlation_matrix(
            mineralogia_data, "Correlación Mineralogía Detallada.xlsx", "Mineralogía"
        )
        geoquimica_data = load_and_process_data(file_path, "Geoquímica")
        plot_and_save_correlation_matrix(
            geoquimica_data, "Correlacion Geoquímica Detallada.xlsx", "Geoquímica"
        )
        print("Correlation analysis completed and saved.")


root = tk.Tk()
root.title("Correlation Matrix Processor")

file_path_label = tk.Label(root, text="Ruta de Archivo:")
file_path_label.pack()

file_path_entry = tk.Entry(root, width=100)
file_path_entry.pack()

browse_button = tk.Button(root, text="Navegar", command=open_file_dialog)
browse_button.pack()

process_button = tk.Button(root, text="Procesar", command=process_file)
process_button.pack()

root.mainloop()
