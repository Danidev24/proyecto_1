import time
import pandas as pd
from modbus_tk import modbus_tcp
import modbus_tk.defines as cst
from datetime import datetime
from openpyxl import load_workbook
import os

class PLCClient:
    def __init__(self, ip, port, excel_path):
        self.master = modbus_tcp.TcpMaster(host=ip, port=port)
        self.master.set_timeout(2)
        self.excel_path = excel_path
        self.lista_datos = []
        self.difficulties = {
            1: "Muy Fácil",
            2: "Fácil",
            3: "Medio",
            4: "Difícil",
            5: "Muy Difícil"
        }

    def save_to_bigdata(self, promedio, dificultad, valor, dosificacion, rango):
        try:
            sheet_name = "BigData"
            new_row = {
                "Promedio": promedio,
                "Dificultad": dificultad,
                "Valor": valor,
                "Dosificación": dosificacion,
                "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Rango": rango
            }

            # Cargar archivo Excel
            if os.path.exists(self.excel_path):
                with pd.ExcelFile(self.excel_path) as xls:
                    # Si existe BigData, cargarla y agregar nueva fila
                    if sheet_name in xls.sheet_names:
                        df_existing = pd.read_excel(self.excel_path, sheet_name=sheet_name)
                        df_final = pd.concat([df_existing, pd.DataFrame([new_row])], ignore_index=True)
                    else:
                        df_final = pd.DataFrame([new_row])
            else:
                df_final = pd.DataFrame([new_row])

            # Guardar usando openpyxl sin afectar otras hojas
            with pd.ExcelWriter(self.excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df_final.to_excel(writer, sheet_name=sheet_name, index=False)

            print("✅ Datos guardados correctamente en la hoja 'BigData'.")

        except Exception as e:
            print(f"❌ Error al guardar los datos en Excel: {e}")


    def read_excel_value(self, sheet_name, column, row):
        try:
            df = pd.read_excel(self.excel_path, sheet_name=sheet_name)
            column_index = ord(column.upper()) - ord('A')
            value = df.iloc[row - 1, column_index]

            print(f"Leyendo hoja: '{sheet_name}', columna: '{column}' (índice {column_index}), fila: {row}")
            print(f"Valor encontrado en {column}{row}: {value}")

            if isinstance(value, (float, int)):
                return float(value)
            else:
                return None
        except Exception as e:
            print(f"Error al leer el archivo Excel: {e}")
            return None

    def calcular_dosificacion_desde_excel(self, turbidez, promedio, dificultad):
        difficulty_sheet = self.difficulties.get(dificultad)
        if not difficulty_sheet:
            print(f"Dificultad no válida: {dificultad}")
            return None, None
            
        print(f"DEBUG: Hoja de dificultad seleccionada: {difficulty_sheet}")
        print(f"DEBUG: Valores recibidos - Promedio: {promedio}, Turbidez: {turbidez}")

        column_ranges = [(1, 15), (16, 25), (26, 50), (51, 75), (76, 100), (101, 125), (126, 150),
                         (151, 175), (176, 200), (201, 225), (226, 250), (251, 275), (276, 300),
                         (301, 350), (351, 400), (401, 450), (451, 500), (501, 550), (551, 600),
                         (601, 650), (651, 700), (701, 750), (751, 800), (801, 850), (851, 900),
                         (901, 950), (951, 1000)]

        column = None
        for idx, (low, high) in enumerate(column_ranges):
            if low <= promedio <= high:
                column = chr(ord('C') + idx)
                break

        row_number = None
        rango_turbidez = None
        for idx, (low, high) in enumerate(column_ranges):
            if low <= turbidez <= high:
                row_number = idx + 4
                rango_turbidez = f"{low}-{high}"
                break

        if column is None or row_number is None:
            print("No se encontró un rango válido para turbidez o promedio.")
            return None, None

        valor = self.read_excel_value(difficulty_sheet, column, row_number)
        return valor, rango_turbidez



    def ejecutar(self):
        while True:
            try:
                registro_estado = self.master.execute(1, cst.READ_HOLDING_REGISTERS, 3, 1)[0]

                if registro_estado == 1:
                    print("PLC está enviando datos...")
                    turbidez = self.master.execute(1, cst.READ_HOLDING_REGISTERS, 0, 1)[0]
                    promedio1 = self.master.execute(1, cst.READ_HOLDING_REGISTERS, 1, 1)[0]
                    promedio = promedio1/1000
                    print(f"el promedio leido es ------------- {promedio}")
                    dificultad = self.master.execute(1, cst.READ_HOLDING_REGISTERS, 2, 1)[0]
                    flag = self.master.execute(1, cst.READ_HOLDING_REGISTERS, 3, 1)[0]

                    if dificultad in self.difficulties:
                        dificultad_texto = self.difficulties[dificultad]
                    else:
                        dificultad_texto = "Desconocido"

                    self.lista_datos.append([turbidez, promedio, dificultad_texto, flag])
                    print(f"Datos recibidos: {self.lista_datos[-1]}")
                    time.sleep(2)

                else:
                    print("PLC dejó de enviar. Preparando dosificación...")

                    if self.lista_datos:
                        ultimo = self.lista_datos[-1]
                        turbidez_final = ultimo[0]
                        promedio_final = ultimo[1]
                        dificultadGuardar = ultimo[2]
                        dificultad_final = list(self.difficulties.keys())[list(self.difficulties.values()).index(ultimo[2])]

                        dosificacion, rango = self.calcular_dosificacion_desde_excel(
                            turbidez=turbidez_final,
                            promedio=promedio_final,
                            dificultad=dificultad_final
                        )

                        if dosificacion is not None:
                            self.save_to_bigdata(promedio_final, dificultadGuardar, turbidez_final, dosificacion, rango)
                            dosificacion_entera = int(float(dosificacion) * 100)
                            print(f"Enviando dosificación al PLC: {dosificacion_entera}")
                            self.master.execute(1, cst.WRITE_MULTIPLE_REGISTERS, 4, output_value=[dosificacion_entera, 1])
                        else:
                            print("No se pudo calcular una dosificación válida desde Excel.")


                    self.lista_datos.clear()
                    print("Lista de datos limpiada. Esperando 2 segundos para volver a leer...")
                    time.sleep(2)

            except Exception as e:
                print(f"Error en la comunicación Modbus: {e}")
                time.sleep(2)

# Ejemplo de uso:
if __name__ == "__main__":
    cliente = PLCClient(ip="127.0.0.1", port=502, excel_path="data/datos.xlsx")
    cliente.ejecutar()
