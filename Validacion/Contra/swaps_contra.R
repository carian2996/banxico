# Ian Castillo Rosales
# 26062014

swapsc_junta <- function(ruta1, ruta2, ruta3, ruta4){
      
      options(scipen=999)
      options(encoding="UTF-8")
      
      source(ruta1)
      source(ruta2)
      source(ruta3)
      source(ruta4)
      
      s1 <- swaps1_contra()
      s2 <- swaps2_contra()
      s3 <- swaps3_contra()
      s4 <- swaps4_contra()
      
      s1$SECCION <- "I"
      s2$SECCION <- "II"
      s3$SECCION <- "III"
      s4$SECCION <- "VI"
      
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
      swaps[swaps$SECCION=="VI", columnas] <- s4[, columnas]
      
      write.dbf(swaps, paste("swaps_contra_", format(Sys.Date()[1], "%d_%m_%Y"), ".dbf", sep=""))
}