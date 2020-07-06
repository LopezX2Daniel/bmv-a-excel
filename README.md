# bmv-a-excel

Los Estados Financieros que la BMV ofrece para su consulta estan en formato XBRL (https://www.bmv.com.mx/es/emisoras/archivos-estadar-xbrl), sin embargo no es posible realizar análisis financiero de forma directa dada la imposibilidad de descargar la informacion a excel u otro formato más flexible para su manipulación. El objetivo de este script es que los pequeños inversionistas, estudiantes, profesores o interesados en el mundo financiero en México puedan acceder de forma directa a los estados financieros de empresas que cotizan en la BMV.

**Requisitos:**

El scrip utiliza las siguientes librerías: os, subprocess, requests, json, io, zipfile, BeautifulSoup y xlsxwriter.

Una vez ejecutado, se tendrán dos opciones para descargar estados financieros.

Descargar(0) & Descargar(1)

La unica diferencia entre ambos es que 0 permite descargar una sola vez la lista de links, se guardará como variable y permite evitar la conexion a la pagina de la BMV para solicitad la lista actualizada de links. Mientras que 1 se conecta cada vez y actualiza la lista de links, de forma que siempre tendrá la version mas actualizada.

