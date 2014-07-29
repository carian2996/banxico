# Ian Castillo Rosales
# 23072014
# Identifica los paquetes para subyacente en divisas 
# utilizando un algoritmo de discriminacion en base
# a la difinicion de paquete de derivados.

# Un paquete de derivados cumple con las siguientes
# caracteristicas:

# = Institucion
# = Fecha de concertacion
# = Tipo de operacion
# = Tipo de opcion
# = Importe base
# = Precio de ejercicio
# = Subyacente
# Si MDO != E => = MDO
# Si MDO == E => = CONT


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
      
      # ========== Eliminar los registros que no pueden ser paquete ==========      
      features <- c("INSTI", "FE_CON_OPE", "TIP_OPE", "OPC", "PREC_EJE", "SUBY")
      conteos <- subset(count(data, features), freq >= 2)
      new_data <- match_df(data, conteos, on = features)
      
      conteos <- subset(count(new_data[new_data$MDO == "E", ], "CONT"), freq >= 2)
      new_data1 <- match_df(new_data[new_data$MDO == "E", ], conteos, on = "CONT")
      
      conteos <- subset(count(new_data[new_data$MDO != "E", ], "CONT"), freq >= 2)
      new_data2 <- match_df(new_data[new_data$MDO != "E", ], conteos, on = "CONT")
      
      new_data <- rbind(new_data1, new_data2)
      
      # ========== Separar por subyacente ==========
      data_suby <- dlply(new_data, .(new_data$SUBY), function(x) x)
      
      ordena <- function(x){
            x <- arrange(x, x$PREC_EJE, x$IMP_BASE, x$FE_VEN_OPE)
            x
      }
      
      data_suby <- llply(data_suby, ordena)
      
      paquetes <- function(x){
            x <- data_suby[[8]]
            subyacente <- unique(x$SUBY)
            x$diff <- c(0, diff(as.Date(x$FE_VEN_OPE, "%Y/%m/%d")))
            
            consec <- x[x$ID %in% x[x$diff > 0, ]$ID, ]
            consec_ant <- x[x$ID %in% (x[x$diff > 0, ]$ID - 1) & x$SUBY == subyacente, ]
            
            x <- rbind(consec, consec_ant)
            x <- x[!duplicated(x), ]
      }
      
      data_suby <- llply(data_suby, paquetes)
      data_suby <- llply(data_suby, ordena)
}

x <- data_suby[[8]]
