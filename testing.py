from openpyxl import workbook, cell, load_workbook

global run
run = True #Crea una variable para que el script siga corriendo siempre que el usuario quiera
ruta2 = input("Ubicacion del archivo DDJJ Ganancias: ")


def ventas(hojaV, mes): #Crea una funcion que busca dentro del archivo DDJJ Ganancias la fila y la columna en la cual debera ubicar los datos
    global fila1
    global columna
    for celda in hojaV['A']:
        if celda.value == mes:
            fila1 = celda.row
            for celda2 in hojaV['C']:
                if celda2.value == "$": #La diferencia entre "$" y "gravado" es que en algunos archivos aparece una, y en otros aparecen ambas. Por eso es una solucion global.
                    fila2 = celda2.row
                    if hojaV.cell(row=fila2, column=5).value == "gravado":
                        columna = 5
                        return columna
                        break
                    else:
                        columna = celda2.column
                        return columna
                        break
            return fila1 

gastos = ["BCO.CREDICOOP - ARGENCARD", "BANCO CREDICOOP - CABAL", "EDENOR SA", "TELEFONICA",
 "AMERICAN EXPRESS SA", "REDGUARD SA"]

def compras(hojaC, hojaG, mes, neto):
    global valor1
    global valor2
    global gastoTotal
    global rowC
    for celda in hojaC['D']:
        for gasto in gastos: #Busca cada uno de los gastos definidos en la lista de la linea 27
            if celda.value == gasto: #Compara el valor de la celda con los gastos
                row = celda.row + 1 #Al tener uno o dos valores (ya que hay importes con un iva de 10,5% y otros con 21%)
                row2 = celda.row + 2 #Busca los valores de las 2 filas siguientes a donde se encuentra el concepto gasto 
                valor1 = hojaC.cell(row=row, column=1).value #Como los valores siempre estan en la columna A (o column=1), toma esos valores
                valor2 = hojaC.cell(row=row2, column=1).value
                if valor2 == None: #Evita el error por NoneType, ya que si encuentra que el segundo valor es nulo, solo suma el valor1
                    gastoTotal =+ valor1
                elif type(valor1) and type(valor2) == float:
                    gastoTotal =+ valor1 + valor2
    for celda2 in hojaG['A']:
        if celda2.value == mes:
            rowC = celda2.row
            return gastoTotal and rowC

while run == True:
    ruta1 = str(input("Ubicacion del libro : ")) 
    archivo = load_workbook(ruta1) #Abre el archivo Excel del libro a utilizar
    hoja = archivo.active
    mes = int(input('Mes a cargar (insertar nº del 1-12): '))
    año = str(input('Año del libro (ejemplo: 2020): '))

    global nombremes #Crea la variable donde se alojara el nombre del mes, para asi encontrar en el archivo de DDJJ la fila donde debera ubicar los datos   
    global totalmes #Crea la variable donde se alojaran los totales de los meses, para asi encontrar mas facilmente los datos necesarios
    if mes == 1:
        nombremes = "ENERO"
        totalmes = f"TOTALES AL 31/01/{año}"
    elif mes == 2:
        nombremes = "FEBRERO"
        totalmes = f"TOTALES AL 28/02/{año}"
    elif mes == 3:
        nombremes = "MARZO"
        totalmes = f"TOTALES AL 31/03/{año}"
    elif mes == 4:
        nombremes = "ABRIL"
        totalmes = f"TOTALES AL 30/04/{año}"
    elif mes == 5:
        nombremes = "MAYO"
        totalmes = f"TOTALES AL 31/05/{año}"
    elif mes == 6:
        nombremes = "JUNIO"
        totalmes = f"TOTALES AL 30/06/{año}"
    elif mes == 7:
        nombremes = "JULIO"
        totalmes = f"TOTALES AL 31/07{año}"
    elif mes == 8:
        nombremes = "AGOSTO"
        totalmes = f"TOTALES AL 31/08/{año}"
    elif mes == 9:
        nombremes = "SEPTIEMBRE"
        totalmes = f"TOTALES AL 30/09/{año}"
    elif mes == 10:
        nombremes = "OCTUBRE"
        totalmes = f"TOTALES AL 31/10/{año}"
    elif mes == 11:
        nombremes = "NOVIEMBRE"
        totalmes = f"TOTALES AL 30/11/{año}"
    elif mes == 12:
        nombremes = "DICIEMBRE"
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

    archivo2 = load_workbook(ruta2) #Busca y abre el archivo donde se alojaran los datos del libro

    global notError1
    notError1 = False #Evita errores en el caso de que el usuario inserte mal el numero

    while notError1 == False:
        libro = int(input('Libro (ventas=1, compras=2): ')) #El usuario decide si va a cargar un libro de Ventas o de compras)
        if libro == 1:
            hoja2 = archivo2['VENTAS']
            ventas(hoja2, nombremes)
            hoja2.cell(row=fila1, column=columna).value = netogravado #Carga el valor obtenido en la fila y la columna segun la funcion ventas()
            archivo2.save(ruta2) #Sobreescribe el archivo original, con los datos obtenidos
            notError1 = True
        elif libro == 2:
            hoja2 = archivo2['COMPRAS']
            compras(hoja, hoja2, nombremes, netogravado)
            hoja2.cell(row=rowC, column= 7).value = gastoTotal
            hoja2.cell(row=rowC, column= 5).value = netogravado - gastoTotal
            archivo2.save(ruta2)
            notError1 = True

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
