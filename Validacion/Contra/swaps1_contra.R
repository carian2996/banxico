# Ian Castillo Rosales
# 26062014

swaps1_contra <- function(){
      # ===== SWAPS 1 =====
      
      # SALIDA
      # cuadro_swaps1_[fecha].dbf - Archivo tipo .dbf con los resultados
      
      # ===== Librerias y directorios =====
      setwd("/Volumes/IAN/Estadisticas/Contraparte/SWAPS") # ¿Dónde están mis datos?
      library(foreign) # Libreria necesaria para cargar los datos
      options(scipen=999)
      options(encoding="UTF-8")
      
      # ===== Carga de datos =====
      data <- read.dbf("swaps1.dbf", as.is=T)
      gc()
      udis <- read.dbf("udi2013.dbf", as.is=T)
      fix <- read.dbf("tcfix.dbf", as.is=T)

      clave_deri <- read.csv("clave_deri.csv", as.is=T)
      
      # ===== Código =====
      data <- data[complete.cases(data$FE_CON_OPE, data$FE_VEN_OPE, data$MDO, 
                                  data$C_IMP_BASE, data$MDA_IMP), ]
      raros <- data[!complete.cases(data$FE_CON_OPE, data$FE_VEN_OPE, data$MDO, 
                                    data$C_IMP_BASE, data$MDA_IMP), ]
      
      if(nrow(raros)!=0){
            message("Existen registros incompletos")
      }
      
      data$FE_CON_OPE <- as.Date(data$FE_CON_OPE) # Cambiar tipo caractér a tipo fecha
      
      # ===== Tipo de Institucion =====
      data$TIPO_INST <- NA
      data$TIPO_INST[substr(data$INSTI, 1, 3) == "013"] <- "CB"
      data$TIPO_INST[substr(data$INSTI, 1, 3) != "013"] <- "BM_BD"
      
      # ===== UDIS y FIX =====
      data$UDIS <- udis$CIERRE[match(data$FE_CON_OPE, as.Date(udis$FE_PUBLI))] # Buscar UDIS y unir con datos 
      data$FIX <- fix$CIERRE[match(data$FE_CON_OPE, as.Date(fix$FE_PUBLI))] # Buscar FIX y unir con datos
      
      # ===== IMPORTE =====
      data$IMPORTE <- NA
      
      data$IMPORTE[data$MDA_IMP=="MXP"] <- data$C_IMP_BASE[data$MDA_IMP=="MXP"]/2000
      data$IMPORTE[data$MDA_IMP=="UDI"] <- data$C_IMP_BASE[data$MDA_IMP=="UDI"]*data$UDIS[data$MDA_IMP=="UDI"]/2000
      data$IMPORTE[data$MDA_IMP=="USD"] <- data$C_IMP_BASE[data$MDA_IMP=="USD"]*data$FIX[data$MDA_IMP=="USD"]/2000
      
      # ===== Sector ====
      data$SECTOR <- "Mercado Mexicano de Derivados"
      
      # ===== WRITE =====
      # Escribe el cuadro (.dbf) en el directorio de trabajo
      write.dbf(data, paste("swaps1_contra_", format(Sys.Date()[1], "%d_%m_%Y"), "dbf", sep=""))
      
      data
}