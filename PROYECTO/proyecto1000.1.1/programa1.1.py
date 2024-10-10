import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd

def importar_excel(archivo):
    """Importa un archivo Excel y devuelve un DataFrame de Pandas."""
    try:
        df = pd.read_excel(archivo)
        return df
    except Exception as e:
        print(f"Error al importar el archivo: {e}")
        return None

def exportar_excel(df, archivo):
    """Exporta un DataFrame de Pandas a un archivo Excel."""
    try:
        df.to_excel(archivo, index=False)
        print(f"Archivo exportado correctamente a: {archivo}")
    except Exception as e:
        print(f"Error al exportar el archivo: {e}")

def determinar_estado(disponibilidad):
    """Devuelve el estado del stock basado en la disponibilidad."""
    if disponibilidad < 2:
        return "CRITICO"
    elif 2 <= disponibilidad < 3:
        return "SUB STOCK"
    elif 3 <= disponibilidad <= 6:
        return "NORMO STOCK"
    else:
        return "SOBRE STOCK"

def redistribuir_stock(df, meses):
    """Redistribuye el stock entre los puestos considerando la disponibilidad y las salidas en los meses proporcionados."""
    try:
        df['stock'] = pd.to_numeric(df['stock'], errors='coerce')
        df['precio'] = pd.to_numeric(df['precio'], errors='coerce')
        df['disponibilidad'] = pd.to_numeric(df['disponibilidad'], errors='coerce')

        df = df[(df['disponibilidad'] >= 3) & (df['disponibilidad'] <= 6)]
        redistribucion = []
        
        for micro_red in df['micro red'].unique():
            df_micro_red = df[df['micro red'] == micro_red]

            for index, row in df_micro_red.iterrows():
                stock_actual = row['stock']
                stock_a_dar = 0
                stock_a_recibir = 0
                establecimiento_origen = row['establecimiento']
                establecimiento_destino = row['establecimiento']  # Se inicializa con el mismo establecimiento

                if stock_actual > 0:
                    for mes in meses:
                        if mes in df.columns:
                            salidas = pd.to_numeric(row[mes], errors='coerce')
                            stock_a_dar += salidas
                            stock_a_dar = min(stock_a_dar, stock_actual)  # No exceder el stock actual
                else:
                    medicamento = row['codigo']
                    otros_establecimientos = df[(df['codigo'] == medicamento) & (df['stock'] > 0)]

                    for _, otro_row in otros_establecimientos.iterrows():
                        stock_disponible = otro_row['stock']
                        for mes in meses:
                            if mes in df.columns:
                                salidas = pd.to_numeric(row[mes], errors='coerce')
                                if stock_disponible > 0:
                                    cantidad_dada = min(salidas, stock_disponible)
                                    stock_a_recibir += cantidad_dada  # Acumular lo recibido
                                    stock_disponible -= cantidad_dada
                                    establecimiento_origen = otro_row['establecimiento']
                                    establecimiento_destino = row['establecimiento']  # Asignar el destino
                                    break  # Salir del bucle si se encontró el stock

                stock_final = stock_actual + stock_a_recibir - stock_a_dar
                disponibilidad = row['disponibilidad']
                estado = determinar_estado(disponibilidad)

                if stock_a_recibir > 0 and pd.notna(row['precio']):
                    total = stock_a_recibir * row['precio']
                else:
                    total = 0

                redistribucion.append({
                    'MICRO RED': micro_red,
                    'ESTABLECIMIENTO': row['establecimiento'],
                    'COD-MEDICAMENTO': row['codigo'],
                    'MEDICAMENTO': row['medicamentos'],
                    'PRECIO': row['precio'],
                    'STOCK ACTUAL': stock_actual,
                    'STOCK A DAR': stock_a_dar,
                    'STOCK A RECIBIR': stock_a_recibir,
                    'STOCK FINAL': stock_final,
                    'TOTAL': total,
                    'ESTABLECIMIENTO DE DONDE SE EXTRAE EL STOCK': establecimiento_origen,
                    'ESTABLECIMIENTO A DONDE SE TRASPASA EL STOCK': establecimiento_destino,
                    'DISPONIBILIDAD': disponibilidad,
                    'ESTADO': estado
                })

        df_redistribuido = pd.DataFrame(redistribucion)
        df_redistribuido = df_redistribuido.sort_values(by=['MICRO RED', 'ESTABLECIMIENTO'])
        
        return df_redistribuido.reset_index(drop=True)
    except Exception as e:
        print(f"Error al redistribuir el stock: {e}")
        return None

class App:
    def __init__(self, master):
        self.master = master
        master.title("Gestor de Tablas Excel")

        self.frame = tk.Frame(master)
        self.frame.pack(padx=10, pady=10)

        self.boton_importar = tk.Button(self.frame, text="Importar Excel", command=self.importar_archivo)
        self.boton_importar.grid(row=0, column=0, padx=5, pady=5)

        self.boton_exportar = tk.Button(self.frame, text="Exportar Excel", command=self.exportar_archivo)
        self.boton_exportar.grid(row=0, column=1, padx=5, pady=5)

        self.boton_redistribuir = tk.Button(self.frame, text="Redistribuir Stock", command=self.redistribuir_columna)
        self.boton_redistribuir.grid(row=0, column=2, padx=5, pady=5)

        # Filtros de búsqueda
        self.label_buscar_micro_red = tk.Label(master, text="Buscar Micro Red:")
        self.label_buscar_micro_red.pack()
        self.entry_buscar_micro_red = tk.Entry(master)
        self.entry_buscar_micro_red.pack()
        self.boton_buscar_micro_red = tk.Button(master, text="Buscar", command=self.filtrar_micro_red)
        self.boton_buscar_micro_red.pack()

        self.label_buscar_establecimiento = tk.Label(master, text="Buscar Establecimiento:")
        self.label_buscar_establecimiento.pack()
        self.entry_buscar_establecimiento = tk.Entry(master)
        self.entry_buscar_establecimiento.pack()
        self.boton_buscar_establecimiento = tk.Button(master, text="Buscar", command=self.filtrar_establecimiento)
        self.boton_buscar_establecimiento.pack()

        self.label_buscar_medicamento = tk.Label(master, text="Buscar Medicamento:")
        self.label_buscar_medicamento.pack()
        self.entry_buscar_medicamento = tk.Entry(master)
        self.entry_buscar_medicamento.pack()
        self.boton_buscar_medicamento = tk.Button(master, text="Buscar", command=self.filtrar_medicamento)
        self.boton_buscar_medicamento.pack()

        self.label_info = tk.Label(master, text="")
        self.label_info.pack(pady=5)

        self.tree = ttk.Treeview(master, columns=('MICRO RED', 'ESTABLECIMIENTO', 'COD-MEDICAMENTO', 
                                                    'MEDICAMENTO', 'PRECIO', 'STOCK ACTUAL', 
                                                    'STOCK A DAR', 'STOCK A RECIBIR', 'STOCK FINAL', 
                                                    'TOTAL', 'ESTABLECIMIENTO DE DONDE SE EXTRAE EL STOCK', 
                                                    'ESTABLECIMIENTO A DONDE SE TRASPASA EL STOCK', 
                                                    'DISPONIBILIDAD', 'ESTADO'), show='headings')
        self.tree.pack(padx=10, pady=10)

        for col in self.tree['columns']:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor='center')

        self.df = None
        self.df_redistribuido = None

    def importar_archivo(self):
        archivo = filedialog.askopenfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
        )
        if archivo:
            try:
                self.df = importar_excel(archivo)
                if self.df is not None:
                    self.df.columns = self.df.columns.str.lower()

                    required_columns = ['micro red', 'codigo_est', 'establecimiento', 'codigo', 
                                        'medicamentos', 'precio', 'siga', 'tipo', 
                                        'petitorio', 'estrategico', 'stock', 
                                        'total', 'cant_sin_ceros', 'cpa', 'disponibilidad']
                    missing_columns = [col for col in required_columns if col not in self.df.columns]

                    if not missing_columns:
                        self.label_info.config(text=f"Archivo '{archivo}' importado correctamente.")
                    else:
                        missing_columns_str = ', '.join(missing_columns)
                        self.label_info.config(text=f"Faltan columnas requeridas: {missing_columns_str}.")
                        self.df = None
                else:
                    self.label_info.config(text="Error al importar el archivo.")
            except Exception as e:
                self.label_info.config(text=f"Error al importar el archivo: {e}")

    def exportar_archivo(self):
        if self.df_redistribuido is not None:
            archivo = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
            )
            if archivo:
                exportar_excel(self.df_redistribuido, archivo)
                self.label_info.config(text=f"Archivo exportado correctamente a: {archivo}")
        else:
            self.label_info.config(text="No hay ningún DataFrame de redistribución cargado.")

    def redistribuir_columna(self):
        if self.df is not None:
            meses = ['setiembre', 'octubre', 'noviembre', 'diciembre', 
                     'enero', 'febrero', 'marzo', 'abril', 
                     'mayo', 'junio', 'julio', 'agosto']
            existing_months = [mes for mes in meses if mes in self.df.columns]
            if existing_months:
                self.df_redistribuido = redistribuir_stock(self.df, existing_months)
                if self.df_redistribuido is not None:
                    self.label_info.config(text="Stock redistribuido correctamente.")
                    self.tree.delete(*self.tree.get_children())

                    for _, row in self.df_redistribuido.iterrows():
                        self.tree.insert("", "end", values=tuple(row))

                else:
                    self.label_info.config(text="Error al redistribuir el stock.")
            else:
                self.label_info.config(text="No hay meses válidos para redistribuir el stock.")

    def filtrar_micro_red(self):
        self.filtrar_tabla('MICRO RED', self.entry_buscar_micro_red.get().upper())

    def filtrar_establecimiento(self):
        self.filtrar_tabla('ESTABLECIMIENTO', self.entry_buscar_establecimiento.get().upper())

    def filtrar_medicamento(self):
        self.filtrar_tabla('MEDICAMENTO', self.entry_buscar_medicamento.get().upper())

    def filtrar_tabla(self, columna, valor):
        if self.df_redistribuido is not None:
            filtrado = self.df_redistribuido[self.df_redistribuido[columna].str.contains(valor, na=False, case=False)]
            if filtrado.empty:
                messagebox.showinfo("Resultado de búsqueda", "No se encontró.")
                self.tree.delete(*self.tree.get_children())
            else:
                self.tree.delete(*self.tree.get_children())
                for _, row in filtrado.iterrows():
                    self.tree.insert("", "end", values=tuple(row))
        else:
            self.label_info.config(text="No hay datos disponibles para filtrar.")

# Ejecución de la aplicación
root = tk.Tk()
app = App(root)
root.mainloop()


