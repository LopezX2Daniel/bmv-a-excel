import os
import subprocess 
import requests
import json
import io
import zipfile
from bs4 import BeautifulSoup as bs
import xlsxwriter

d_links = {}

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36'}

def descarga_links():
    print("Descargado desde web lista de links... ", end="")
    url = "https://www.bmv.com.mx/es/emisoras/archivos-estadar-xbrl"
    
    respuesta = requests.get(url, headers=headers)
    sopadecoditos = bs(respuesta.content, "html.parser")
    tabla_zips = sopadecoditos.find("tbody").findAll("tr")

    print("¡Hecho!")
    print("Preparando la información... ", end="")

    for row in tabla_zips:
        texto_tabla = row.find_all("td")
        emisora = texto_tabla[0].get_text()

        link_tabla = row.find_all("a")

        link = str(link_tabla[0].get("href")).replace("/docs-pub/ifrsxbrl/../visor/visorXbrl.html?docins=../","https://www.bmv.com.mx/docs-pub/")
        periodo = str(link_tabla[0].get_text()).replace("\n","")

        lista_hija = []
        lista_padre = []

        if "anexon" not in link:
            lista_hija.append(link)
            lista_hija.append(periodo)
            lista_padre.append(lista_hija)

            if emisora in d_links:
                d_links[emisora].append(lista_padre)
            else:
                d_links[emisora] = [lista_padre]
    print("¡Hecho!")
    print("\n")


def Descargar(recarga):
    
    if recarga == 1:
        global d_links
        d_links = {}
        descarga_links()

    elif recarga == 0:
        if len(d_links) == 0:
            descarga_links()

    while True:
        emisora_select = str(input("Escoger emisora:")).upper()
        try:
            print("_"*60)
            for i,item in enumerate(d_links[emisora_select]):
                print(f"{i+1}{'# '} {item[0][1]} {' #'}{i+1}")
            print("_"*60)
            print("\n")
            while True:
                try:
                    periodo_select = input("¿Qué periodo?:")
                    try:
                        periodo_select = int(periodo_select)
                    except ValueError:
                        print("Selección vacía. Reintentar.")
                        continue
                    url = d_links[emisora_select][periodo_select-1][0][0]
                    break
                except IndexError:
                    print(f"No tengo registros con el número '{periodo_select}'. Reintentar.")
                    print("\n")
                    continue
            break
        except KeyError:
            print(f"No tengo registros con la emisora '{emisora_select}'. Reintentar.")
            print("\n")
            continue

    print("Descargando y procesando Estados Financieros a Excel... ")

    if str(url[-3:]).lower() == "zip":
        try:
            respuesta = requests.get(url, stream=True)
        except requests.exceptions.ConnectionError:
            raise ConnectionError("El servidor rechaza solicitud, reintentar en 2 minutos.")
        archivo_zip = zipfile.ZipFile(io.BytesIO(respuesta.content), 'r')
        nombre = archivo_zip.namelist()
        enbruto = archivo_zip.read(nombre[0])
        final = json.loads(enbruto)
    elif str(url[-4:]).lower() == "json":
        try:
            respuesta = requests.get(url)
        except requests.exceptions.ConnectionError:
            raise ConnectionError("El servidor rechaza solicitud, reintentar en 2 minutos.")
        soupadecoditos = bs(respuesta.content, "html.parser")
        final = json.loads(soupadecoditos.prettify())
    else:
        raise ValueError(f"La URL '{url}' no termina en .zip o .json, verificar lista de links")

    for hechos in final["HechosPorIdConcepto"].keys():
        if "DateOfEndOfReportingPeriod" in hechos:
            id_periodo_reporte = hechos
            break

    for ronda in range(2):
        fecha_cierre = final["HechosPorId"][final["HechosPorIdConcepto"][id_periodo_reporte][0]]["Valor"]

        if ronda == 1:
            fecha_cierre = str(int(fecha_cierre[0:4])-1)+fecha_cierre[4:10]

        fecha_inicio = str(fecha_cierre[0:4])+"-01-01"
        fecha_inicio_cierre = fecha_inicio + "_" + fecha_cierre
        titulo_resultados_flujo = f"Del {fecha_inicio} Al {fecha_cierre}"
        fecha_cierre_anterior = str(int(fecha_cierre[0:4])-1)+"-12-31"
        anio_actual = str(fecha_cierre[0:4])
        fecha_cierre_actual = fecha_cierre

        if ronda == 1:
            fecha_cierre = str(fecha_cierre[0:4])+"-12-31"

        if int(fecha_cierre[0:4]) <= 2015:
            id_moneda = "ifrs_DescriptionOfPresentationCurrency"
            if "12" in fecha_cierre[5:7]:
                trimestre = "4"
            elif "9" in fecha_cierre[5:7]:
                trimestre = "3"
            elif "6" in fecha_cierre[5:7]:
                trimestre = "2"
            elif "3" in fecha_cierre[5:7]:
                trimestre = "1"
            else:
                trimestre = "ND"
                id_clave_pizarra = "mx-ifrs-ics_ClaveCotizacion"
            id_nombre_emisora = "ifrs_NameOfReportingEntityOrOtherMeansOfIdentification"
            id_efectivo = "ifrs_CashAndCashEquivalents"
        else:
            id_moneda = "ifrs-full_DescriptionOfPresentationCurrency"
            id_trimestre = "ifrs_mx-cor_20141205_NumeroDeTrimestre"
            id_clave_pizarra = "ifrs_mx-cor_20141205_ClaveDeCotizacionBloqueDeTexto"
            id_nombre_emisora = "ifrs-full_NameOfReportingEntityOrOtherMeansOfIdentification"
            id_efectivo = "ifrs-full_CashAndCashEquivalents"
            trimestre = str(final["HechosPorId"][final["HechosPorIdConcepto"][id_trimestre][0]]["Valor"]).upper()

        moneda = str(final["HechosPorId"][final["HechosPorIdConcepto"][id_moneda][0]]["Valor"]).upper()
        clave_pizarra = str(final["HechosPorId"][final["HechosPorIdConcepto"][id_clave_pizarra][0]]["Valor"]).upper()
        nombre_emisora = str(final["HechosPorId"][final["HechosPorIdConcepto"][id_nombre_emisora][0]]["Valor"]).upper()

        l_contextos_unicos = []

        for concepto in final["ContextosPorId"]:
            if final["ContextosPorId"][concepto]["ValoresDimension"] == None:
                l_contextos_unicos.append(final["ContextosPorId"][concepto]["Id"])

        d_balance = {}
        d_resultados = {}
        d_flujo = {}

        def edosfinnombres():
            l_codigos = {"[210000]":"balance","[310000]":"resultados","[520000]":"flujo"}
            no_incluir = ["Table", "SharesAxis", "Member", "LineItems","Explanatory"]
            d_quitar = {"[":"","]":"","{":"","}":"","'":""," ":""}

            for (llave, valor) in l_codigos.items():
                texto_final = None
                for rol in final["Taxonomia"]["RolesPresentacion"]:
                    if llave in rol["Nombre"]:
                        texto_final = str(rol["Estructuras"][0])
                        break

                for (quita, pone) in d_quitar.items():
                    texto_final = texto_final.replace(quita,pone)
                texto_final_split = texto_final.split(",")

                l_conceptos = []

                for nombre in texto_final_split:
                    if "IdConcepto:" in nombre and not any(x in nombre for x in no_incluir):
                        temp_agregar = None
                        temp_agregar = nombre.replace("IdConcepto:","")
                        temp_agregar = temp_agregar.replace("SubEstructuras:","")
                        l_conceptos.append(temp_agregar)

                if valor == l_codigos["[210000]"]:
                    for item in l_conceptos:
                        d_balance[item] = [str(final["Taxonomia"]["ConceptosPorId"][item]["Etiquetas"]["es"]\
                                                    [list(final["Taxonomia"]["ConceptosPorId"][item]["Etiquetas"]["es"].keys())[0]]["Valor"]).replace("[sinopsis]","")]
                elif valor == l_codigos["[310000]"]:
                    for item in l_conceptos:
                        d_resultados[item] = [str(final["Taxonomia"]["ConceptosPorId"][item]["Etiquetas"]["es"]\
                                                       [list(final["Taxonomia"]["ConceptosPorId"][item]["Etiquetas"]["es"].keys())[0]]["Valor"]).replace("[sinopsis]","")]
                elif valor == l_codigos["[520000]"]:
                    for item in l_conceptos:
                        d_flujo[item] = [str(final["Taxonomia"]["ConceptosPorId"][item]["Etiquetas"]["es"]\
                                                  [list(final["Taxonomia"]["ConceptosPorId"][item]["Etiquetas"]["es"].keys())[0]]["Valor"]).replace("[sinopsis]","")]   
                    d_flujo[id_efectivo] = [final["Taxonomia"]["ConceptosPorId"][id_efectivo]["Etiquetas"]["es"]\
                                                 ["http://www.xbrl.org/2003/role/periodStartLabel"]["Valor"]]                 
                    d_flujo[f"{id_efectivo}_Ending"] = [final["Taxonomia"]["ConceptosPorId"][id_efectivo]["Etiquetas"]["es"]\
                                                             ["http://www.xbrl.org/2003/role/periodEndLabel"]["Valor"]]                 
        edosfinnombres()


        for unicontexto in l_contextos_unicos:
            if unicontexto in final["ContextosPorFecha"][fecha_cierre]: uuid_balance = unicontexto
            if unicontexto in final["ContextosPorFecha"][fecha_inicio_cierre]: uuid_resultados = uuid_flujo = unicontexto
            if unicontexto in final["ContextosPorFecha"][fecha_cierre_anterior]: uuid_inicio_efectivo = unicontexto
            if unicontexto in final["ContextosPorFecha"][fecha_cierre_actual]: uuid_cierre_efectivo = unicontexto            

        for hecho in final["HechosPorId"].keys():
            if final["HechosPorId"][hecho]["IdContexto"] == uuid_balance and final["HechosPorId"][hecho]["IdConcepto"] in d_balance:
                d_balance[final["HechosPorId"][hecho]["IdConcepto"]].append(final["HechosPorId"][hecho]["ValorNumerico"])

            if final["HechosPorId"][hecho]["IdContexto"] == uuid_resultados and final["HechosPorId"][hecho]["IdConcepto"] in d_resultados:
                d_resultados[final["HechosPorId"][hecho]["IdConcepto"]].append(final["HechosPorId"][hecho]["ValorNumerico"])

            if final["HechosPorId"][hecho]["IdContexto"] == uuid_flujo and final["HechosPorId"][hecho]["IdConcepto"] in d_flujo:
                d_flujo[final["HechosPorId"][hecho]["IdConcepto"]].append(final["HechosPorId"][hecho]["ValorNumerico"])

            if final["HechosPorId"][hecho]["IdContexto"] == uuid_inicio_efectivo and final["HechosPorId"][hecho]["IdConcepto"] == id_efectivo:
                d_flujo[id_efectivo].append(final["HechosPorId"][hecho]["ValorNumerico"])

            if final["HechosPorId"][hecho]["IdContexto"] == uuid_cierre_efectivo and final["HechosPorId"][hecho]["IdConcepto"] == id_efectivo:
                d_flujo[f"{id_efectivo}_Ending"].append(final["HechosPorId"][hecho]["ValorNumerico"])


        """Inicio de maquila Excel"""
        if ronda == 0:
            nombre_archivo = f"EstadosFinancieros_{clave_pizarra}_{anio_actual}_{trimestre}T.xlsx"
            workbook = xlsxwriter.Workbook(nombre_archivo)
            worksheet_balance = workbook.add_worksheet('Balance General')
            worksheet_resultados = workbook.add_worksheet('Estado de Resultados')
            worksheet_flujo = workbook.add_worksheet('Flujo de Efectivo')

            negritas = workbook.add_format({'bold':True})
            negritas_centrado = workbook.add_format({'bold':True})
            centrado = negritas_centrado.set_align('center')

            cashbaby = workbook.add_format({'num_format':'#,##0.00;[Red](#,##0.00)'})
            nigga_cashbaby = workbook.add_format({'num_format':'#,##0.00;[Red](#,##0.00)','bold':True})

            columna_numeros = 1
        elif ronda == 1:
            columna_numeros = 2


        """Balance General"""
        worksheet_balance.write(0,0,nombre_emisora,negritas_centrado)
        worksheet_balance.write(0,1,f'Cifras en {moneda}',centrado)
        worksheet_balance.write(2,columna_numeros,f"Al {fecha_cierre}",negritas_centrado)

        for i, (llave, valor) in enumerate(d_balance.items()):
            worksheet_balance.write(2+i,0,valor[0])
            try:
                worksheet_balance.write(2+i,columna_numeros,valor[1],cashbaby)
            except IndexError:
                pass

        """Estado de Resultados"""
        worksheet_resultados.write(0,0,nombre_emisora, negritas_centrado)
        worksheet_resultados.write(0,1,f'Cifras en {moneda}')
        worksheet_resultados.write(2,columna_numeros,titulo_resultados_flujo,negritas_centrado)

        for i, (llave, valor) in enumerate(d_resultados.items()):
            worksheet_resultados.write(2+i,0,valor[0])
            try:
                worksheet_resultados.write(2+i,columna_numeros,valor[1],cashbaby)
            except IndexError:
                pass

        """Flujo de Efectivo"""
        worksheet_flujo.write(0,0,nombre_emisora,negritas_centrado)
        worksheet_flujo.write(0,1,'Cifras en {}'.format(moneda))
        worksheet_flujo.write(2,columna_numeros,titulo_resultados_flujo,negritas_centrado)

        for i, (llave, valor) in enumerate(d_flujo.items()):
            worksheet_flujo.write(2+i,0,valor[0])
            try:
                worksheet_flujo.write(2+i,columna_numeros,valor[1],cashbaby)
            except IndexError:
                pass

    worksheet_balance.set_column('A:A',116)
    worksheet_balance.set_column('B:C',18)
    worksheet_resultados.set_column('A:A',62)
    worksheet_resultados.set_column('B:C',25)
    worksheet_flujo.set_column('A:A',108)
    worksheet_flujo.set_column('B:C',25)

    workbook.close()

    try:
        ruta_excel = (f"{os.path.abspath('')}\\{nombre_archivo}")
        subprocess.Popen([ruta_excel],shell=True)
        print("\n")
        print(f"Archivo creado. Abriendo excel: '{nombre_archivo}'")
    except:
        print("\n")
        print(f"Excel '{nombre_archivo}' creado en la misma ruta en donde se encuentre este archivo de Python.")