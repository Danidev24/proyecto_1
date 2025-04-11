import modbus_tk
import modbus_tk.modbus_tcp as modbus_tcp
import modbus_tk.defines as cst
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import time
import threading
import os

class PLCExcelDataProcessor:
    def __init__(self, plc_ip, plc_port, excel_path):
        # Configuración de conexión Modbus
        print('ingresamos aquí...')
        try:
            self.modbus_master = modbus_tcp.TcpMaster(host=plc_ip, port=plc_port)
            self.modbus_master.open()
            print("Conexión establecida con el servidor Modbus")
        except Exception as e:
            print(f"Error al conectar con el servidor Modbus: {e}")
            self.modbus_master = None  # Evita errores posteriores si no se establece conexión
        
        # Configuración de Excel
        self.excel_path = excel_path
        self.difficulty_sheets = self.select_difficulty()  # Aquí defines directamente la hoja de Excel a usar
        
        # Variables de estado
        self.last_30_values = []
        self.last_update_time = datetime.now()
        
        # Variables de configuración
        self.MAX_VALUES = 10
        self.MAX_WAIT_TIME = timedelta(minutes=30)

    # el usuario escoje la dificultad mediante la consola con el rango de número de 1 al 5
    def select_difficulty(self):
        """Muestra un menú para que el usuario seleccione la dificultad y devuelve la hoja correspondiente"""
        difficulties = {
            1: "Muy Fácil",
            2: "Fácil",
            3: "Medio",
            4: "Difícil",
            5: "Muy Difícil"
        }
        
        # Mostrar las opciones al usuario
        print("Seleccione la dificultad:")
        for key, value in difficulties.items():
            print(f"{key} - {value}")
        
        # Solicitar la opción al usuario
        try:
            choice = int(input("Escriba el número de la dificultad que desea elegir: "))
            if choice in difficulties:
                print(f"Ha seleccionado: {difficulties[choice]}")
                return [difficulties[choice]]  # Devolver la hoja correspondiente
            else:
                print("Opción no válida, se seleccionará por defecto 'Muy Difícil'.")
                return ["Muy Difícil"]  # Valor por defecto
        except ValueError:
            print("Entrada no válida, se seleccionará por defecto 'Muy Difícil'.")
            return ["Muy Difícil"]  # Valor por defecto
        

    # se dedica a leer un dato de excel
    def read_excel_value(self, sheet_name, column, row):
        try:
            # Leer el archivo Excel y la hoja específica
            df = pd.read_excel(self.excel_path, sheet_name=sheet_name)
            
            # Convertir la columna en índice numérico (A=0, B=1, ..., H=7)
            column_index = ord(column.upper()) - ord('A')
            
            # Leer el valor de la celda específica
            value = df.iloc[row - 1, column_index]  # Ajuste por índices basados en 0
            print(f"Valor leído de la celda ({column}{row}): {value}")
            
            # Verificar si es un número decimal
            if isinstance(value, (float, int)):
                return float(value)
            else:
                print(f"El valor en la celda no es un número válido: {value}")
                return None
        except Exception as e:
            print(f"Error al leer el archivo Excel: {e}")
            return None
    
    def read_plc_value(self, register_address):
        """Lee un valor específico del PLC via Modbus"""
        if not self.modbus_master:
            print("No se puede leer porque la conexión Modbus no está establecida.")
            return None
        
        try:
            value = self.modbus_master.execute(1, modbus_tk.defines.READ_HOLDING_REGISTERS, 0, 1)[0]
            print(f"Valor leído del PLC en la dirección {register_address}: {value}")
            return value
        except Exception as e:
            print(f"Error al leer PLC: {e}")
            return None
    
    def write_plc_value(self, register_address, value):
        """Escribe un valor en el PLC via Modbus"""
        if not self.modbus_master:
            print("No se puede escribir porque la conexión Modbus no está establecida.")
            return
        
        try:
            print(f"Escribiendo el valor {value} en el registro {register_address}")
            self.modbus_master.execute(1, modbus_tk.defines.WRITE_SINGLE_REGISTER, register_address, output_value=value)
        except Exception as e:
            print(f"Error al escribir en PLC: {e}")


    # guardar la información en la hoja de cálculo llamada BigData
    def save_to_bigdata(self, promedio, dificultad, valor, dosificacion):
        try:
            sheet_name = "BigData"
            new_data = {
                "Promedio": [promedio],
                "Dificultad": [dificultad],
                "Valor": [valor],
                "Dosificación": [dosificacion],
                "Fecha": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
            }

            df_new = pd.DataFrame(new_data)

            if os.path.exists(self.excel_path):
                with pd.ExcelFile(self.excel_path) as xls:
                    if sheet_name in xls.sheet_names:
                        df_existing = pd.read_excel(self.excel_path, sheet_name=sheet_name)
                        df_final = pd.concat([df_existing, df_new], ignore_index=True)
                    else:
                        df_final = df_new
            else:
                df_final = df_new

            with pd.ExcelWriter(self.excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df_final.to_excel(writer, sheet_name=sheet_name, index=False)

            print("Datos guardados en 'BigData' correctamente.")
        except Exception as e:
            print(f"Error al guardar los datos en Excel: {e}")


    def calculate_weighted_average(self, num_readings=2):
        """Calcula promedio ponderado de los últimos 5 valores y reinicia los valores después de cada cálculo"""
        curren_time = datetime.now()

        if curren_time - self.last_update_time < timedelta(minutes=2):
            return None
        
        # Tomamos los últimos 'num_readings' valores
        recent_values = self.last_30_values[-num_readings:]
        
        # Calcular el promedio ponderado
        weights = np.linspace(1, num_readings, num_readings)
        weighted_avg = np.average(recent_values, weights=weights)
        
        # Imprimir el resultado del promedio ponderado
        print(f"Promedio ponderado calculado: {weighted_avg}")
        
        # Reiniciar los valores después de calcular el promedio
        self.last_30_values = []  # Vuelve a su estado inicial
        
        # Devolver el promedio como entero
        return int(weighted_avg)

    
    def process_data(self, row_register, output_register):
        """Procesa los datos completos"""
        # Leer valor de fila del PLC
        # print(f"Procesando datos para la fila del registro {row_register}")
        row = self.read_plc_value(row_register)
        if row is None:
            print('ingresamos acá')
            print("No se pudo leer valor del PLC.")
            return
        
        # Actualizar lista de valores y tiempo
        self.last_30_values.append(row)
        if len(self.last_30_values) > self.MAX_VALUES:
            self.last_30_values.pop(0)
        print(f"Valores leídos por el PLC: {self.last_30_values}")
        
        self.last_update_time = datetime.now()
        
        # Calcular promedio ponderado
        weighted_avg = self.calculate_weighted_average()
        
        if weighted_avg is not None:
            # Aquí usas directamente la dificultad que has definido
            difficulty_sheet = self.difficulty_sheets[0]  # 'Muy Difícil' en este caso
            print(f"Hoja de dificultad determinada: {difficulty_sheet}")
            print(f"Último valor leído del PLC {row} ")
            
            # Definir los rangos de las columnas (esto puede ser una lista con los valores que corresponden a cada columna)
            column_ranges = [
                (1, 15), (16, 25), (26, 50), (51, 75), (76, 100), (101, 125), (126, 150), 
                (151, 175), (176, 200), (201, 225), (226, 250), (251, 275), (276, 300), 
                (301, 350), (351, 400), (401, 450), (451, 500), (501, 550), (551, 600), 
                (601, 650), (651, 700), (701, 750), (751, 800), (801, 850), (851, 900), 
                (901, 950), (951, 1000)
            ]
            
            # Calcular la columna correspondiente al promedio ponderado
            column = None
            for idx, (low, high) in enumerate(column_ranges):
                if low <= weighted_avg <= high:
                    column = chr(ord('C') + idx)  # La columna empieza en 'C'
                    break
            
            if column is None:
                print(f"Promedio ponderado {weighted_avg} no cae en ningún rango.")
                return
            
            # print(f"Promedio ponderado {weighted_avg} corresponde a la columna {column}")
            
            # Calcular la fila correspondiente al valor del PLC
            row_number = None
            for idx, (low, high) in enumerate(column_ranges):
                if low <= row <= high:
                    row_number = idx + 4  # La fila empieza en la 5
                    break
            
            if row_number is None:
                print(f"El valor del PLC {row} no cae en ningún rango.")
                return
            
            # print(f"El valor del PLC {row} corresponde a la fila {row_number}")
            
            # Leer valor de Excel usando la función read_excel_value
            try:
                excel_value = self.read_excel_value(difficulty_sheet, column, row_number)
                if excel_value is not None:
                    print(f"El valor leído de Excel en la celda es: {excel_value}")
                    print(f"La inteligencia artificial ha tomado el control y ha escrito el valor de: {int(excel_value)} *********************")
                    self.save_to_bigdata(weighted_avg, difficulty_sheet, row, excel_value)
                    self.write_plc_value(300, int(excel_value))
                else:
                    print(f"No se pudo leer un valor válido de la celda ({column}{row_number}) en la hoja {difficulty_sheet}.")
            except Exception as e:
                print(f"Error procesando Excel: {e}")
        
        # Lógica de pendiente de respaldo
        if datetime.now() - self.last_update_time > self.MAX_WAIT_TIME:
            trend = self.calculate_trend()
            if trend is not None:
                predicted_row = self.predict_next_row(trend)
                print(f"Predicción de la siguiente fila: {predicted_row}")
                # Procesar con fila predicha
                self.process_data(predicted_row, output_register)
    
    def continuous_processing(self, row_register, output_register, read_interval=5):
        """Procesamiento continuo en hilo separado"""

        last_difficulty_selection = time.time()
        
        while True:

            current_time = time.time()

            # Verificar si ha pasado al menos un minuto (30 seg)  desde la última selección de dificultad
            if current_time - last_difficulty_selection >= 30:
                self.select_difficulty()
                last_difficulty_selection = current_time  # Actualiza el tiempo de última ejecución
                # print("Iniciando procesamiento continuo...")

            # procesar los datos del PLC
            self.process_data(row_register, output_register)

            # esperar el intervalo deficnido (1 minuto, cada minuto lee datos del PLC)
            time.sleep(read_interval)  # Intervalo de procesamiento

def main():
    # Configuración
    plc_processor = PLCExcelDataProcessor(
        # plc_ip='127.0.0.1',  # IP del PLC3
        plc_ip='127.0.0.1',  # IP del PLC3
        plc_port=502,         # Puerto Modbus
        excel_path='data/datos.xlsx'  # Ruta del archivo Excel
    )
    
    # Iniciar procesamiento continuo en un hilo
    processing_thread = threading.Thread(target=plc_processor.continuous_processing,
                                         args=(0, 1))  # Registros de ejemplo
    processing_thread.start()

if __name__ == "__main__":
    main()