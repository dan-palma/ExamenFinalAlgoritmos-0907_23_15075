import openpyxl

libro = openpyxl.load_workbook("vehiculos.xlsx")

hoja = libro['listado']

print (hoja['A1'].value)

def leer_ventas():
    ventas = []
    for numero_fila in range (2, hoja.max_row + 1):
        venta_actual = {}
        venta_actual ['Código']= hoja[f"A{numero_fila}"].value
        venta_actual ['Marca']= hoja[f"B{numero_fila}"].Value
        venta_actual ['Modelo']= hoja[f"C{numero_fila}"].value
        venta_actual ['Precio']= hoja[f"D{numero_fila}"].value
        venta_actual ['Kilometraje']= hoja[f"E{numero_fila}"].value
        ventas.append(venta_actual)
        print(ventas)

leer_ventas()