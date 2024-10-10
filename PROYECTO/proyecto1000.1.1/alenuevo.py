import sys
import pandas as pd
import matplotlib.pyplot as plt
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, 
                             QLineEdit, QVBoxLayout, QWidget, 
                             QFileDialog, QTableWidget, QTableWidgetItem, 
                             QMessageBox, QGridLayout, QMenuBar, QMenu, QAction, QProgressBar, QComboBox)
from PyQt5.QtCore import Qt, QTimer
import qdarkstyle

def importar_excel(archivo):
    try:
        df = pd.read_excel(archivo)
        return df
    except Exception as e:
        print(f"Error al importar el archivo: {e}")
        return None

def exportar_excel(df, archivo):
    try:
        df.to_excel(archivo, index=False)
        print(f"Archivo exportado correctamente a: {archivo}")
    except Exception as e:
        print(f"Error al exportar el archivo: {e}")

def determinar_estado(disponibilidad):
    if disponibilidad < 2:
        return "CRITICO"
    elif 2 <= disponibilidad < 3:
        return "SUB STOCK"
    elif 3 <= disponibilidad <= 6:
        return "NORMO STOCK"
    else:
        return "SOBRE STOCK"

def calcular_porcentaje_cpa(cpa):
    try:
        cpa['cpa'] = pd.to_numeric(cpa['cpa'], errors='coerce')
        cpa['total'] = pd.to_numeric(cpa['total'], errors='coerce')
        cpa['ABASTECIMIENTO'] = (cpa['cpa'] / cpa['total']) * 100
        return cpa
    except Exception as e:
        print(f"Error al calcular el porcentaje de CPA: {e}")
        return cpa

def redistribuir_stock(df, meses, progress_callback):
    try:
        df['stock'] = pd.to_numeric(df['stock'], errors='coerce')
        df['precio'] = pd.to_numeric(df['precio'], errors='coerce')
        df['disponibilidad'] = pd.to_numeric(df['disponibilidad'], errors='coerce')

        df['original_index'] = df.index

        redistribucion = []
        total_rows = len(df)
        processed_rows = 0
        
        for micro_red in df['micro red'].unique():
            df_micro_red = df[df['micro red'] == micro_red]

            for index, row in df_micro_red.iterrows():
                stock_actual = row['stock']
                abastecimiento = 0
                stock_a_recibir = 0
                stock_reutilizado = 0
                establecimiento_origen = "NO SE EXTRAE STOCK"
                establecimiento_destino = "NO SE TRASPASAN STOCK"
                tipo_medicamento = row['tipo']

                if stock_actual > 0:
                    for mes in meses:
                        if mes in df.columns:
                            salidas = pd.to_numeric(row[mes], errors='coerce')
                            abastecimiento += salidas
                            abastecimiento = min(abastecimiento, stock_actual)

                medicamento = row['codigo']
                otros_establecimientos = df[(
                    df['codigo'] == medicamento) & 
                    (df['stock'] > 0) & 
                    (df['establecimiento'] != row['establecimiento']) & 
                    (df['tipo'] == tipo_medicamento)
                ]

                if abastecimiento > 0 and row['establecimiento'] in otros_establecimientos['establecimiento'].values:
                    stock_reutilizado = min(abastecimiento, stock_actual)
                    abastecimiento -= stock_reutilizado

                if abastecimiento > 0:
                    establecimiento_origen = row['establecimiento']
                    for _, otro_row in otros_establecimientos.iterrows():
                        stock_disponible = otro_row['stock']
                        if abastecimiento > 0:
                            for mes in meses:
                                if mes in df.columns:
                                    salidas = pd.to_numeric(row[mes], errors='coerce')
                                    cantidad_a_recibir = min(salidas, stock_disponible, abastecimiento)

                                    stock_a_recibir += cantidad_a_recibir
                                    abastecimiento -= cantidad_a_recibir
                                    establecimiento_destino = otro_row['establecimiento']
                                    stock_disponible -= cantidad_a_recibir

                stock_total = stock_reutilizado + stock_a_recibir
                stock_final = stock_actual + stock_a_recibir - abastecimiento
                disponibilidad = row['disponibilidad']
                estado = determinar_estado(disponibilidad)

                total = stock_a_recibir * row['precio'] if stock_a_recibir > 0 and pd.notna(row['precio']) else 0

                redistribucion.append({
                    'MICRO RED': micro_red,
                    'ESTABLECIMIENTO': row['establecimiento'],
                    'COD-MEDICAMENTO': row['codigo'],
                    'MEDICAMENTO': row['medicamentos'],
                    'PRECIO': row['precio'],
                    'STOCK ACTUAL': stock_actual if total != 0 else "SC",
                    'ABASTECIMIENTO': row['cpa'],
                    'STOCK A RECIBIR': stock_a_recibir if stock_a_recibir > 0 else ("SC" if total == 0 else ""),
                    'STOCK FINAL': (row['cpa'] + stock_a_recibir) if total != 0 else "SC",
                    'TOTAL': total,
                    'ESTABLECIMIENTO DE DONDE SE EXTRAE EL STOCK': establecimiento_origen,
                    'ESTABLECIMIENTO A DONDE SE TRASPASA EL STOCK': establecimiento_destino if abastecimiento > 0 else "NO SE TRASPASAN STOCK",
                    'DISPONIBILIDAD': disponibilidad,
                    'ESTADO': estado,
                    'original_index': row['original_index']
                })

                processed_rows += 1
                progress_callback(int((processed_rows / total_rows) * 100))

        df_redistribuido = pd.DataFrame(redistribucion)
        df_redistribuido = df_redistribuido.sort_values(by='original_index').reset_index(drop=True)
        df_redistribuido.drop(columns=['original_index'], inplace=True)
        
        return df_redistribuido
    except Exception as e:
        print(f"Error al redistribuir el stock: {e}")
        return None

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gestor de Tablas Excel")
        self.setGeometry(100, 100, 1000, 600)

        self.setStyleSheet(""" 
            QMainWindow {
                background-color: #f0f0f0;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px;
                font-size: 14px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QLabel {
                font-size: 14px;
                margin: 5px 0;
            }
            QLineEdit {
                padding: 8px;
                border: 1px solid #ccc;
                border-radius: 4px;
                font-size: 14px;
            }
            QTableWidget {
                border: 1px solid #ccc;
                border-radius: 4px;
                font-size: 14px;
            }
        """)

        self.layout = QGridLayout()

        self.boton_importar = QPushButton("Importar Excel")
        self.boton_importar.clicked.connect(self.importar_archivo)
        self.layout.addWidget(self.boton_importar, 0, 0)

        self.boton_exportar = QPushButton("Exportar Excel")
        self.boton_exportar.clicked.connect(self.exportar_archivo)
        self.layout.addWidget(self.boton_exportar, 0, 1)

        self.boton_redistribuir = QPushButton("Redistribuir Stock")
        self.boton_redistribuir.clicked.connect(self.redistribuir_columna)
        self.layout.addWidget(self.boton_redistribuir, 0, 2)
        
        self.label_buscar_micro_red = QLabel("Buscar Micro Red:")
        self.layout.addWidget(self.label_buscar_micro_red, 1, 0)
        self.combo_buscar_micro_red = QComboBox()
        self.combo_buscar_micro_red.addItems(["", "09 DE OCTUBRE", "IPARIA", "MASISEA", "PURUS", "SAN FERNANDO"])
        self.layout.addWidget(self.combo_buscar_micro_red, 1, 1)
        self.boton_buscar_micro_red = QPushButton("Buscar Micro Red")
        self.boton_buscar_micro_red.clicked.connect(self.filtrar_micro_red)
        self.layout.addWidget(self.boton_buscar_micro_red, 1, 2)
        
        self.label_buscar_establecimiento = QLabel("Buscar Establecimiento:")
        self.layout.addWidget(self.label_buscar_establecimiento, 2, 0)
        self.entry_buscar_establecimiento = QLineEdit()
        self.layout.addWidget(self.entry_buscar_establecimiento, 2, 1)
        self.boton_buscar_establecimiento = QPushButton("Buscar Establecimiento")
        self.boton_buscar_establecimiento.clicked.connect(self.filtrar_establecimiento)
        self.layout.addWidget(self.boton_buscar_establecimiento, 2, 2)

        self.label_buscar_medicamento = QLabel("Buscar Medicamento:")
        self.layout.addWidget(self.label_buscar_medicamento, 3, 0)
        self.entry_buscar_medicamento = QLineEdit()
        self.layout.addWidget(self.entry_buscar_medicamento, 3, 1)
        self.boton_buscar_medicamento = QPushButton("Buscar Medicamento")
        self.boton_buscar_medicamento.clicked.connect(self.filtrar_medicamento)
        self.layout.addWidget(self.boton_buscar_medicamento, 3, 2)

        self.label_buscar_disponibilidad = QLabel("Buscar Disponibilidad:")
        self.layout.addWidget(self.label_buscar_disponibilidad, 4, 0)
        self.combo_buscar_disponibilidad = QComboBox()
        self.combo_buscar_disponibilidad.addItems(["", "CRITICO", "SUB STOCK", "NORMO STOCK", "SOBRE STOCK"])
        self.layout.addWidget(self.combo_buscar_disponibilidad, 4, 1)
        self.label_rango_disponibilidad = QLabel("Rango de Disponibilidad:")
        self.layout.addWidget(self.label_rango_disponibilidad, 5, 0)
        self.entry_rango_min = QLineEdit()
        self.entry_rango_min.setPlaceholderText("Min")
        self.layout.addWidget(self.entry_rango_min, 5, 1)
        self.entry_rango_max = QLineEdit()
        self.entry_rango_max.setPlaceholderText("Max")
        self.layout.addWidget(self.entry_rango_max, 5, 2)
        self.boton_buscar_disponibilidad = QPushButton("Buscar Disponibilidad")
        self.boton_buscar_disponibilidad.clicked.connect(self.filtrar_disponibilidad)
        self.layout.addWidget(self.boton_buscar_disponibilidad, 6, 0, 1, 3)

        self.label_info = QLabel("")
        self.layout.addWidget(self.label_info, 7, 0, 1, 3)

        self.table_widget = QTableWidget()
        self.layout.addWidget(self.table_widget, 8, 0, 1, 3)

        self.progress_bar = QProgressBar(self)
        self.layout.addWidget(self.progress_bar, 9, 0, 1, 3)

        container = QWidget()
        container.setLayout(self.layout)
        self.setCentralWidget(container)

        self.df = None
        self.df_redistribuido = None

        self.crear_menu()

    def crear_menu(self):
        menubar = self.menuBar()
        archivo_menu = menubar.addMenu('Archivo')

        importar_action = QAction('Importar Excel', self)
        importar_action.triggered.connect(self.importar_archivo)
        archivo_menu.addAction(importar_action)

        exportar_action = QAction('Exportar Excel', self)
        exportar_action.triggered.connect(self.exportar_archivo)
        archivo_menu.addAction(exportar_action)

    def importar_archivo(self):
        archivo, _ = QFileDialog.getOpenFileName(self, "Abrir Archivo Excel", "", "Archivos Excel (*.xlsx);;Todos los archivos (*)")
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
                    extra_columns = [col for col in self.df.columns if col not in required_columns]

                    if not missing_columns:
                        self.label_info.setText(f"Archivo '{archivo}' importado correctamente.")
                    else:
                        missing_columns_str = ', '.join(missing_columns)
                        self.label_info.setText(f"Faltan las columnas: {missing_columns_str}")

                    if extra_columns:
                        extra_columns_str = ', '.join(extra_columns)
                        print(f"Columnas adicionales encontradas: {extra_columns_str}")

                    # Sugerencias de columnas adicionales
                    suggested_columns = ['fecha_actualizacion', 'proveedor', 'categoria', 'ubicacion_almacen', 'cantidad_minima_requerida']
                    suggested_columns_str = ', '.join(suggested_columns)
                    print(f"Sugerencias de columnas adicionales: {suggested_columns_str}")

                    # Calcular porcentaje de CPA
                    self.df = calcular_porcentaje_cpa(self.df)

            except Exception as e:
                self.label_info.setText("Error al importar el archivo.")
                print(f"Error: {e}")

    def exportar_archivo(self):
        if self.df_redistribuido is not None:
            archivo, _ = QFileDialog.getSaveFileName(self, "Guardar Archivo Excel", "", "Archivos Excel (*.xlsx);;Todos los archivos (*)")
            if archivo:
                exportar_excel(self.df_redistribuido, archivo)
                self.label_info.setText(f"Archivo exportado correctamente a: {archivo}")
        else:
            self.label_info.setText("No hay ningún DataFrame de redistribución cargado.")

    def redistribuir_columna(self):
        if self.df is not None:
            meses = ['setiembre', 'octubre', 'noviembre', 'diciembre', 
                     'enero', 'febrero', 'marzo', 'abril', 
                     'mayo', 'junio', 'julio', 'agosto']
            existing_months = [mes for mes in meses if mes in self.df.columns]
            if existing_months:
                self.df_redistribuido = redistribuir_stock(self.df, existing_months, self.update_progress)
                if self.df_redistribuido is not None:
                    self.label_info.setText("Stock redistribuido correctamente.")
                    self.table_widget.setRowCount(0)
                    self.table_widget.setColumnCount(len(self.df_redistribuido.columns))
                    self.table_widget.setHorizontalHeaderLabels(self.df_redistribuido.columns)

                    for _, row in self.df_redistribuido.iterrows():
                        row_position = self.table_widget.rowCount()
                        self.table_widget.insertRow(row_position)
                        for column, value in enumerate(row):
                            self.table_widget.setItem(row_position, column, QTableWidgetItem(str(value)))

                else:
                    self.label_info.setText("Error al redistribuir el stock.")
            else:
                self.label_info.setText("No hay meses válidos para redistribuir el stock.")
    
    def filtrar_micro_red(self):
        self.filtrar_tabla('MICRO RED', self.combo_buscar_micro_red.currentText())
        
    def filtrar_establecimiento(self):
        self.filtrar_tabla('ESTABLECIMIENTO', self.entry_buscar_establecimiento.text().upper())

    def filtrar_medicamento(self):
        self.filtrar_tabla('MEDICAMENTO', self.entry_buscar_medicamento.text().upper())

    def filtrar_disponibilidad(self):
        estado = self.combo_buscar_disponibilidad.currentText()
        rango_min = self.entry_rango_min.text()
        rango_max = self.entry_rango_max.text()

        if estado:
            self.filtrar_tabla('ESTADO', estado)
        elif rango_min and rango_max:
            try:
                rango_min = float(rango_min)
                rango_max = float(rango_max)
                self.filtrar_rango_disponibilidad(rango_min, rango_max)
            except ValueError:
                QMessageBox.warning(self, "Error de entrada", "Por favor, ingrese valores numéricos válidos para el rango de disponibilidad.")

    def filtrar_rango_disponibilidad(self, rango_min, rango_max):
        if self.df_redistribuido is not None:
            filtrado = self.df_redistribuido[(self.df_redistribuido['DISPONIBILIDAD'] >= rango_min) & (self.df_redistribuido['DISPONIBILIDAD'] <= rango_max)]

            if filtrado.empty:
                QMessageBox.information(self, "Resultado de búsqueda", "No se encontró.")
                self.table_widget.setRowCount(0)
            else:
                self.table_widget.setRowCount(0)
                self.table_widget.setColumnCount(len(filtrado.columns))
                self.table_widget.setHorizontalHeaderLabels(filtrado.columns)

                for _, row in filtrado.iterrows():
                    row_position = self.table_widget.rowCount()
                    self.table_widget.insertRow(row_position)
                    for column, value in enumerate(row):
                        self.table_widget.setItem(row_position, column, QTableWidgetItem(str(value)))
        else:
            self.label_info.setText("No hay datos disponibles para filtrar.")

    def filtrar_tabla(self, columna, valor):
        if self.df_redistribuido is not None:
            if isinstance(valor, str):
                filtrado = self.df_redistribuido[self.df_redistribuido[columna].str.contains(valor, na=False, case=False)]
            else:
                filtrado = self.df_redistribuido[self.df_redistribuido[columna] == valor]

            if filtrado.empty:
                QMessageBox.information(self, "Resultado de búsqueda", "No se encontró.")
                self.table_widget.setRowCount(0)
            else:
                self.table_widget.setRowCount(0)
                self.table_widget.setColumnCount(len(filtrado.columns))
                self.table_widget.setHorizontalHeaderLabels(filtrado.columns)

                for _, row in filtrado.iterrows():
                    row_position = self.table_widget.rowCount()
                    self.table_widget.insertRow(row_position)
                    for column, value in enumerate(row):
                        self.table_widget.setItem(row_position, column, QTableWidgetItem(str(value)))
        else:
            self.label_info.setText("No hay datos disponibles para filtrar.")

    def update_progress(self, value):
        self.progress_bar.setValue(value)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    sys.exit(app.exec_())