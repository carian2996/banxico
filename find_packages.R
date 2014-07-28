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

find_packages <- function(ruta, file_name, n_records, n_fields) {
      ruta = "/Volumes/IAN"
      file_name = "base_paquetes_divisas.csv"
      
	# ENTRADA
	# ruta = Especifica la ruta del directorio donde se encuentran los datos
	# file_name = Especifica el nombre del archivo .dbf

	# SALIDA
	# paquetes = Lista que contiene los paquetes identificados


	# ===== Librerias y directorios =====
      setwd(ruta) # Donde estan mis datos?
      
      # ===== Carga de datos =====
      data <- read.csv(file_name, quote = "", as.is = T)
      data$INSTI <- sprintf("%06d", data[, "INSTI"])
      data$MDA_IMP <- sprintf("%02d", data[, "MDA_IMP"])
      data$SUBY <- sprintf("%03d", data[, "SUBY"])
      data$CONT <- sprintf("%06d", data[, "CONT"])
      data <- cbind(seq(nrow(data)), data)
      names(data)[1] <- "ID"
      
      # Busca en que columna hay por lo menos un valor vacio
      if(any(apply(data, 2, function(x) any(is.na(x))))){
            message("Existen registros incompletos la base")
      }
      
      precios <- subset(count(data, "PREC_EJE"), freq >= 2)
      new_data <- match_df(data, precios, on="PREC_EJE")

      importe_base <- subset(count(new_data, "IMP_BASE"), freq >= 2)
      new_data <- match_df(new_data, importe_base, on="IMP_BASE")
      
      institucion <- subset(count(new_data, "INSTI"), freq >= 2)
      new_data <- match_df(new_data, institucion, on="INSTI")
      
      fecha_con <- subset(count(new_data, "FE_CON_OPE"), freq >= 2)
      new_data <- match_df(new_data, fecha_con, on="FE_CON_OPE")
      
      subyacentes <- sort(unique(data$SUBY))
      list_packages <- list()
      
      for(i in seq_along(subyacentes)){
            
            new_data <- data[data$SUBY == subyacentes[i], ]
            precios <- sort(unique(new_data$PREC_EJE))
            precios <- precios[table(new_data$PREC_EJE) >= 2]
            new_data <- new_data[new_data$PREC_EJE %in% precios, ]
            
            imp_base <- sort(unique(new_data$IMP_BASE))
            imp_base <- imp_base[table(new_data$IMP_BASE) >= 2]
            new_data <- new_data[new_data$IMP_BASE %in% imp_base, ]
            
            insti <- sort(unique(new_data$INSTI))
            insti <- insti[table(new_data$INSTI) >= 2]
            new_data <- new_data[new_data$INSTI %in% insti, ]
            
            fecha_con <- sort(unique(new_data$FE_CON_OPE))
            fecha_con <- fecha_con[table(new_data$FE_CON_OPE) >= 2]
            new_data <- new_data[new_data$FE_CON_OPE %in% fecha_con, ]
            
            list_packages[[i]] <- new_data
      }
      
      names(list_packages) <- subyacentes
      
      new_data <- list_packages[subyacentes[8]][[1]]
      dlply(new_data, .(new_data$PREC_EJE), nrow)

}
