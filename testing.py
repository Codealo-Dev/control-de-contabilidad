from openpyxl import workbook
from openpyxl import load_workbook
from openpyxl import cell

global run
run = True #Crea una variable para que el script siga corriendo siempre que el usuario quiera

while run == True:
    ruta1 = str(input("Ubicacion del libro : ")) 
    archivo = load_workbook(ruta1) #Abre el archivo Excel del libro a utilizar
    hoja = archivo.active
    mes = int(input('Mes a cargar (insertar nº del 1-12): '))
    año = str(input('Año del libro (ejemplo: 2020): '))
        
    global totalmes #Crea la variable donde se alojaran los totales de los meses, para asi encontrar mas facilmente los datos necesarios
    if mes == 1:
        totalmes = f"TOTALES AL 31/01/{año}"
    elif mes == 2:
        totalmes = f"TOTALES AL 28/02/{año}"
    elif mes == 3:
        totalmes = f"TOTALES AL 31/03/{año}"
    elif mes == 4:
        totalmes = f"TOTALES AL 30/04/{año}"
    elif mes == 5:
        totalmes = f"TOTALES AL 31/05/{año}"
    elif mes == 6:
        totalmes = f"TOTALES AL 30/06/{año}"
    elif mes == 7:
        totalmes = f"TOTALES AL 31/07{año}"
    elif mes == 8:
        totalmes = f"TOTALES AL 31/08/{año}"
    elif mes == 9:
        totalmes = f"TOTALES AL 30/09/{año}"
    elif mes == 10:
        totalmes = f"TOTALES AL 31/10/{año}"
    elif mes == 11:
        totalmes = f"TOTALES AL 30/11/{año}"
    elif mes == 12:
        totalmes = f"TOTALES AL 31/12/{año}"
    else:
        print("Error en el mes (linea 12-37)")
        

    for celda in hoja['A']: #Recorre la columna A en busca de la fila donde estan los valores, gracias a la variable totalmes
        if celda.value == totalmes: 
            fila = celda.row 
            netogravado = hoja[f"E{fila}"].value #Busca los valores del neto gravado, el IVA, y el total en base a la fila donde encontro a la variable totalmes
            iva = hoja[f"F{fila}"].value
            importetotal = hoja[f"G{fila}"].value
        else: 
            if celda.value == f"TOTALES AL 29/02/{año}": #Evita que el script se rompa, en el caso de que un libro sea de un año bisiesto
                fila = celda.row
                netogravado = hoja[f"E{fila}"].value
                iva = hoja[f"F{fila}"].value
                importetotal = hoja[f"G{fila}"].value


    ruta2 = input("Ubicacion del archivo DDJJ Ganancias: ")
    archivo2 = load_workbook(ruta2) #Busca y abre el archivo donde se alojaran los datos del libro
    global notError1
    notError1 = False #Evita errores en el caso de que el usuario inserte mal el numero
    while notError1 == False:
        libro = int(input('Libro (ventas=1, compras=2): ')) #El usuario decide si va a cargar un libro de Ventas o de compras)
        if libro == 1:
            hoja2 = archivo2['VENTAS']
            if mes == 1:
                hoja2.cell(row=8, column=5).value = netogravado
                archivo2.save(ruta2)
                notError1 = True
        elif libro == 2:
            hoja2 = archivo2['COMPRAS']

    global notError2
    notError2 = False
    while notError2 == False:
        sigLibro = int(input('¿Desea ingresar otro libro?(Si=1, No=2): '))
        if sigLibro == 1:
            run = True
            notError2 = True
        elif sigLibro == 2:
            run = False
            notError2 = True
        else:
            notError2 = False
