# Ian Castillo Rosales (BANXICO\T41348)
# Gerencia de Informacion del Sistema Financiero
# Subgerencia de Informacion de Moneda Extranjera y Derivados
# 
# Validacion de informacion para operaciones con swaps (contra)
# 230614 - 300414

swaps_contra_junta <- function(ruta){
      
      # ENTRADA
      # ruta = Ruta donde se encuentran los archivos .R para los calculos
      
      # SALIDA
      # swaps_contra_[fecha].dbf - Archivo tipo .dbf con los resultados
      
      # ===== Librerias y directorios =====
      library(foreign) # Libreria necesaria para cargar los datos
      options(scipen=999, digits=8) # Quita la notacion exp y trunca a 4 decimales
      
      # ===== Funciones =====
      source(paste(ruta, "R/swaps1_contra.R", sep=""))
      source(paste(ruta, "R/swaps2_contra.R", sep=""))
      source(paste(ruta, "R/swaps3_contra.R", sep=""))
      source(paste(ruta, "R/swaps4_contra.R", sep=""))
      
      if(class(swaps1_contra(ruta))=="data.frame"){
            s1 <- swaps1_contra(ruta)
      } else{
            s1 <- swaps1_contra(ruta)$"cuadro"
      }
      
      if(class(swaps1_contra(ruta))=="data.frame"){
            s2 <- swaps2_contra(ruta)
      } else{
            s2 <- swaps2_contra(ruta)$"cuadro"
      }
      
      if(class(swaps1_contra(ruta))=="data.frame"){
            s3 <- swaps3_contra(ruta)
      } else{
            s3 <- swaps3_contra(ruta)$"cuadro"
      }
      
      if(class(swaps1_contra(ruta))=="data.frame"){
            s4 <- swaps4_contra(ruta)
      } else{
            s4 <- swaps4_contra(ruta)$"cuadro"
      }
      
      s1$SECCION <- "I"
      s2$SECCION <- "II"
      s3$SECCION <- "III"
      s4$SECCION <- "IV"
      
      columnas <- c("INSTI", "FE_CON_OPE", "TIPO_INST", "IMPORTE", "SECTOR", "SECCION")
      swaps <- data.frame(matrix(NA, nrow=sum(nrow(s1), nrow(s2), nrow(s3), 
                                              nrow(s4)), ncol=12))
      
      colnames(swaps) <- c("INSTI", "FE_CON_OPE", "C_IMP_BASE", "MDA_IMP", 
                           "FE_VEN_ORI", "FE_LIQ_ORI", "TIPO_INST", "UDIS", "FIX", 
                           "IMPORTE", "SECTOR", "SECCION")
      
      swaps$SECCION <- c(s1$SECCION, s2$SECCION, s3$SECCION, s4$SECCION)
      
      swaps[swaps$SECCION=="I", ] <- s1[, 1:12]
      swaps[swaps$SECCION=="II", columnas] <- s2[, columnas]
      swaps[swaps$SECCION=="III", columnas] <- s3[, columnas]
      swaps[swaps$SECCION=="IV", columnas] <- s4[, columnas]
      
      # Escribe el cuadro (.xlsx) en el directorio de trabajo
      write.dbf(swaps, paste("swaps_contra_", format(Sys.Date()[1], "%d%m%Y"), ".dbf", sep=""))
      
      swaps
}