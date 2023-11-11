import openpyxl

libro = openpyxl.load_workbook("vehiculos.xlsx")

hoja = libro['listado']

hoja['A1'].value = "Código"
hoja['B1'].value = "Marca"
hoja['C1'].value = "Modelo"
hoja['D1'].value = "Precio"
hoja['E1'].value = "Kilometraje"

datos_vehiculos = [
   {
      "Código":"CITY01",
      "Marca":"HONDA",
      "Modelo":"2020",
      "Precio":"80000",
      "Kilometraje":"600"
   },
   {
      "Código":"CIVIC01",
      "Marca":"HONDA",
      "Modelo":"2021",
      "Precio":"90000",
      "Kilometraje":"0"
   },
   {
      "Código":"PILOT01",
      "Marca":"HONDA",
      "Modelo":"2021",
      "Precio":"40000",
      "Kilometraje":"1300"
   },
    {
      "Código":"BT50",
      "Marca":"MAZDA",
      "Modelo":"2021",
      "Precio":"50000",
      "Kilometraje":"600"
   },
   {
      "Código":"BALENO1",
      "Marca":"SUZUKI",
      "Modelo":"2021",
      "Precio":"60000",
      "Kilometraje":"2000"
   },
   {
      "Código":"XL71",
      "Marca":"SUZUKI",
      "Modelo":"2021",
      "Precio":"70000",
      "Kilometraje":"1500"
   }
]

proxima_fila = hoja.max_row + 1

    
for venta in datos_vehiculos :
    hoja[f'A{proxima_fila}'].value = venta["Código"]
    hoja[f'B{proxima_fila}'].value = venta["Marca"]
    hoja[f'C{proxima_fila}'].value = venta["Modelo"]
    hoja[f'D{proxima_fila}'].value = venta["Precio"]
    hoja[f'E{proxima_fila}'].value = venta["Kilometraje"]
    proxima_fila +=1

    libro.save("vehiculos.xlsx")