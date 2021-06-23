# Estados Financieros de Emisoras en la BMV a Excel.



Los Estados Financieros que la Bolsa Mexicana de Valores ofrece para su consulta están en formato XBRL:
```
https://www.bmv.com.mx/es/emisoras/archivos-estadar-xbrl
```
Sin embargo, no es posible realizar análisis financiero de forma directa dada la imposibilidad de descargar la informacion a excel u otro formato más flexible para su manipulación. El objetivo de este script es que los pequeños inversionistas, estudiantes, profesores o interesados en el mundo financiero en México puedan acceder de forma directa a los estados financieros de empresas que cotizan en la BMV.

** **
**Requisitos**

El script utiliza las siguientes librerías:
```
os
subprocess
requests
json
io
zipfile
BeautifulSoup
xlsxwriter
```


**Funcionamiento**

Para ejecutar, se tiene dos opciones:

```
Descargar(0)
```
```
Descargar(1)
```

La única diferencia entre ambos es que ***Descargar(0)*** permite descargar una sola vez la lista de links, se guardará como variable y evita conectarse a la pagina de la BMV cada vez que se vuelva a descargar estados financieros. Mientras que ***Descargar(1)*** se conecta y actualiza la lista de links cada vez que se ejecute, de forma que siempre tendrá la versión más actualizada. 
Generalmente, la mayoría de los usuarios les será util usar ***Descargar(0)***.

Después, se deberá escribir el nombre de la emisora (sin serie). Es indiferente si se hace en mayúsculas o minúsculas.

```
Descargado desde web lista de links... ¡Hecho!
Preparando la información... ¡Hecho!
```

Una vez escogida, se mostrará el listado de fechas de estados financieros que se tiene disponible en la BMV. El periodo a selecionar es el número que se encuentra en la parte izquieda o derecha.

```
Escoger emisora:
____________________________________________________________
1#  Descargar Información Del Trimestre 1 Del Año 2020   #1
2#  Descargar Información Del Trimestre 4 Del Año 2019   #2
3#  Descargar Información Del Trimestre 3 Del Año 2019   #3
4#  Descargar Información Del Trimestre 2 Del Año 2019   #4
5#  Descargar Información Del Trimestre 1 Del Año 2019   #5
6#  Descargar Información Del Trimestre 4 Del Año 2018   #6
7#  Descargar Información Del Trimestre 3 Del Año 2018   #7
8#  Descargar Información Del Trimestre 2 Del Año 2018   #8
9#  Descargar Información Del Trimestre 1 Del Año 2018   #9
10#  Descargar Información Del Trimestre 4 Del Año 2017   #10
11#  Descargar Información Del Trimestre 3 Del Año 2017   #11
12#  Descargar Información Del Trimestre 2 Del Año 2017   #12
13#  Descargar Información Del Trimestre 1 Del Año 2017   #13
14#  Descargar Información Del Trimestre 4 Del Año 2016   #14
15#  Descargar Información Del Trimestre 3 Del Año 2016   #15
16#  Descargar Información Del Trimestre 2 Del Año 2016   #16
17#  Descargar Información Del Trimestre 1 Del Año 2016   #17
18#  Descargar Información Del Trimestre 4 Del Año 2015   #18
19#  Descargar Información Del Trimestre 3 Del Año 2015  #19
20#  Descargar Información Del Trimestre 2 Del Año 2015  #20
____________________________________________________________

¿Qué periodo?: 
```
Posteriomente, el script descargará y procesará la informacion a formato Excel, que se guardará en la mista ruta en donde Python este corriendo.


```
Descargando y procesando Estados Financieros a Excel... 


Archivo creado. Abriendo excel: 'EstadosFinancieros_PE&OLES_2020_1T.xlsx'
```

El nombre del archivo de Excel esta compuesto de "EstadosFinancieros_" "Emisora" "Año" "Trimestre".xlsx

Dado que la BMV solo tiene estados financieros trimestrales disponibles, si se desea descargar anuales, de deberá descargar el 4to trimestre de cada emisora de cada año.

Espero que este script sea de utilidad y ayude a difundir más nuestro mercado bursátil en México.

@author: Daniel Eduardo López López
@email: e.lopezlopezdaniel@protonmail.com
@linkedin: linkedin.com/in/lopezeduardodaniel
