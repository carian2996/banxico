# Ian Castillo Rosales
# 23072014
# Identifica los paquetes para subyacente en divisas 
# utilizando un algoritmo de discriminacion en base
# a la difinicion de paquete de derivados.

# Un paquete de derivados cumple con las siguientes
# caracteristicas:

# = Institucion
# = Fecha de concertacion
# = Importe base
# = Precio de ejercicio
# = Subyacente 

# Ademas cumple con lo siguiente:
# La fecha de vencimiento para los derivados del paquete
# tienen una fecha de vencimiento consecutiva, es decir, 
# La diferencia entre las fechas de vencimiento de cada
# derivado son iguales.

find_packages <- function(ruta, file_name) {

	# ENTRADA
	# ruta = Especifica la ruta del directorio donde se encuentran los datos
	# file_name = Especifica el nombre del archivo .dbf

	# SALIDA
	# paquetes = Lista que contiene los paquetes identificados


	# ===== Librerias y directorios =====
      setwd(paste(ruta)) # Donde estan mis datos?
      library(foreign) # Libreria necesaria para cargar los datos
      options(scipen=999, digits=8) # Opciones de formato
      
      # ===== Carga de datos =====
      data <- read.dbf(file_name, as.is=T) # Lee los datos
      data <- data[data$, ] # Quita los paquetes que ya fueron identificados
      gc() # Libera espacio en la memoria

      # Busca en que columna hay por lo menos un valor vacio
      apply(data, 2, function(x) any(is.na(x)))
}
