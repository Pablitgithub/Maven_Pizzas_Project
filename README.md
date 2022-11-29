# Maven_Pizzas_Project
Con el uso de csv files sobre la pizzería maven(pedidos, pizzas, tipos, ingredientes, precio),  me he encargado de, realizar una estimación sobre la cantidad de ingredientes demanda semanalmente, acompañada de una serie de gráficas y reportes.

## Guía

Es necesario tener instaladas las librerías del requirements.txt con el comando pip install -r requirements.txt, así como los csv no pertenecientes al output, y en los datos del 2016 haría falta tener descargadas las imágnes de las gráficas o descomentar las líneas con el comando savefig y en ellas la dirección en la que se está ejecutanto el script.

# Datos del 2015
Con los csv's de este año al encontrarse en perfectas condiciones, mediante una ETL, extraemos de los csv la información que me ha parecido importante, tanto la cantidad como el tamaño de las pizzas, aplicando un multiplicador dependiendo de su tamaño(s,m,l.xl,xxl), y de esta forma agrupar las pizzas por tipo en lugar de por tamaño.

Finalmente una vez hemos aplicado los multiplicadores, calculamos los ingradientes y las raciones de cada uno de estos a lo largo del año, hacemos la estimación y devolvemos el resultado tanto en un csv con los ingredientes y sus cantidades como en un archivo de excel en el que se observa la calidad del dato. Quedará adjuntado el resultado de la ejecución del programa. Serán necesarios para esta los csv y las librearías que aparecen en el requirements.txt de cada carpeta.


# Datos del 2016

Se trata de realizar lo mismo que para el año 2015 pero en este caso los csv no se encontraban en perfecto estado si no que hemos tenido que limpiarlos hasta el punto de dejarlos en un estado óptimo para nuestra estimación.

Además de la estimación y su posterior output en un csv, hemos querido añadir un archivo xml con dicha estimación, un archivo deexcel con distintas gráficas, relacionadas con las pizzas más vendidas, los ingresos, y la cantidad de pizzas vendidas anualmente. Finalmente un reporte pdf en el que se encuentran dichas gráficas y una serie de conclusiones que se toman de estas. Para la ejecución de este programa, nuevamente será necesaria la instalación de la librerías utilizadas, además de la descarga de los csv y las imágenes de las gráficas para el reporte en el archivo .pdf, de igual manera se encontrará un ejemplo del output en caso de cualquier tipo de fallo.
