# Ian Castillo Rosales (BANXICO\T41348)
# Gerencia de Información del Sistema Financiero
# Subgerencia de Información de Moneda Extranjera y Derivados
# 
# Validación de información para operaciones con swaps (plazo)
# 040614 - 300414

swaps_plazo_junta <- function(ruta){
      
      # ENTRADA
      # ruta = Ruta donde se encuentran los archivos .R para los cálculos
      
      # SALIDA
      # swaps_plazo_[fecha].dbf - Archivo tipo .dbf con los resultados
      
      # ===== Librerias y directorios =====
      library(foreign) # Libreria necesaria para cargar los datos
      options(scipen=999, digits=8) # Quita la notación exp y trunca a 4 decimales
      
      # ===== Funciones =====
      source(paste(ruta, "R/swaps1_plazo.R", sep=""))
      source(paste(ruta, "R/swaps2_plazo.R", sep=""))
      source(paste(ruta, "R/swaps3_plazo.R", sep=""))
      source(paste(ruta, "R/swaps4_plazo.R", sep=""))
      
      if(class(swaps1_plazo(ruta))=="data.frame"){
            s1 <- swaps1_plazo(ruta)
      } else{
            s1 <- swaps1_plazo(ruta)$"cuadro"
      }
      
      if(class(swaps1_plazo(ruta))=="data.frame"){
            s2 <- swaps2_plazo(ruta)
      } else{
            s2 <- swaps2_plazo(ruta)$"cuadro"
      }
      
      if(class(swaps1_plazo(ruta))=="data.frame"){
            s3 <- swaps3_plazo(ruta)
      } else{
            s3 <- swaps3_plazo(ruta)$"cuadro"
      }
      
      if(class(swaps1_plazo(ruta))=="data.frame"){
            s4 <- swaps4_plazo(ruta)
      } else{
            s4 <- swaps4_plazo(ruta)$"cuadro"
      }
      
      # ===== JUNTA =====
      s1$SECCION <- "I"
      s2$SECCION <- "II"
      s3$SECCION <- "III"
      s4$SECCION <- "IV"
      
      # Selecciona las columnas de interés
      columnas <- c("INSTI", "FE_CON_OPE", "TIPO_INST", "IMPORTE", "BANDA", "SECCION")
      swaps <- data.frame(matrix(NA, nrow=sum(nrow(s1), nrow(s2), nrow(s3), 
                                              nrow(s4)), ncol=13))
      
      colnames(swaps) <- c("INSTI", "FE_CON_OPE", "C_IMP_BASE", "MDA_IMP", 
                           "FE_VEN_ORI", "FE_LIQ_ORI", "TIPO_INST", "UDIS", "FIX", 
                           "IMPORTE", "PLAZO", "BANDA", "SECCION")
      
      swaps$SECCION <- c(s1$SECCION, s2$SECCION, s3$SECCION, s4$SECCION)
      
      swaps[swaps$SECCION=="I", ] <- s1[, 1:13]
      swaps[swaps$SECCION=="II", columnas] <- s2[, columnas]
      swaps[swaps$SECCION=="III", columnas] <- s3[, columnas]
      swaps[swaps$SECCION=="IV", columnas] <- s4[, columnas]
      
      # Escribe el cuadro (.xlsx) en el directorio de trabajo
      write.dbf(swaps, paste("swaps_plazo_", format(Sys.Date()[1], "%d%m%Y"), ".dbf", sep=""))
      
      swaps
      # apply(swaps, 2, function(x) any(is.na(x)))
}