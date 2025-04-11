from pymodbus.client import ModbusTcpClient

# Configura el cliente Modbus TCP con la dirección IP y el puerto de tu PLC
cliente = ModbusTcpClient('192.168.1.200', port=502)

# Intenta conectar con el PLC
if cliente.connect():
    print("✅ Conectado al PLC HNC")

    # Lee 2 registros desde la dirección 0
    respuesta = cliente.read_holding_registers(address=0, count=2, unit=1)

    if not respuesta.isError():
        print("📥 Datos recibidos:", respuesta.registers)
    else:
        print("⚠️ Error en la respuesta del PLC:", respuesta)

    # Cierra la conexión
    cliente.close()
    print("🔌 Conexión cerrada")
else:
    print("❌ No se pudo conectar al PLC HNC")
