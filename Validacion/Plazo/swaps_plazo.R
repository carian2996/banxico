# Ian Castillo Rosales
# 26062014

swapsp_junta <- function(ruta){
      
      options(scipen=999)
      options(encoding="UTF-8")
      
      source(paste(ruta, "swaps1_plazo.R", sep=""))
      source(paste(ruta, "swaps2_plazo.R", sep=""))
      source(paste(ruta, "swaps3_plazo.R", sep=""))
      source(paste(ruta, "swaps4_plazo.R", sep=""))
      
      s1 <- swaps1_plazo()
      s2 <- swaps2_plazo()
      s3 <- swaps3_plazo()
      s4 <- swaps4_plazo()
      
      s1$SECCION <- "I"
      s2$SECCION <- "II"
      s3$SECCION <- "III"
      s4$SECCION <- "VI"
      
      columnas <- c("INSTI", "FE_CON_OPE", "TIPO_INST", "IMPORTE", "BANDA", "SECCION")
      swaps <- data.frame(matrix(NA, nrow=sum(nrow(s1), nrow(s2), nrow(s3), 
                                              nrow(s4)), ncol=12))
      
      colnames(swaps) <- c("INSTI", "FE_CON_OPE", "C_IMP_BASE", "MDA_IMP", 
                           "FE_VEN_ORI", "FE_LIQ_ORI", "TIPO_INST", "UDIS", "FIX", 
                           "IMPORTE", "BANDA", "SECCION")
      
      swaps$SECCION <- c(s1$SECCION, s2$SECCION, s3$SECCION, s4$SECCION)
      
      swaps[swaps$SECCION=="I", ] <- s1[, 1:12]
      swaps[swaps$SECCION=="II", columnas] <- s2[, columnas]
      swaps[swaps$SECCION=="III", columnas] <- s3[, columnas]
      swaps[swaps$SECCION=="VI", columnas] <- s4[, columnas]
      
      write.dbf(swaps, paste("swaps_plazo_", format(Sys.Date()[1], "%d_%m_%Y"), ".dbf", sep=""))
}

swapsp_junta(ruta="/Volumes/IAN/Estadisticas/Plazo/R/")
