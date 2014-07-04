# Ian Castillo Rosales (BANXICO\T41348)
# Gerencia de Informacion del Sistema Financiero
# Subgerencia de Informacion de Moneda Extranjera y Derivados
# 
# Validacion de informacion para operaciones con swaps (contraparte)
# 090614 - 010714

swaps1_contra <- function(ruta){
      
      # ENTRADA
      # ruta = Ruta donde se encuentran los datos para los calculos
            # swaps4.dbf
            # udi2013.dbf
            # fix.dbf
      
      # SALIDA
      # swaps1_contra_[fecha].dbf - Archivo tipo .dbf con los resultados
      
      # ===== Librerias y directorios =====
      setwd(paste(ruta, "SWAPS/", sep="")) # Donde estan mis datos?
      library(foreign) # Libreria necesaria para cargar los datos
      options(scipen=999, digits=8)
      
      # ===== Carga de datos =====
      data <- read.dbf("swaps1.dbf", as.is=T)
      gc()
      udis <- read.dbf("udi2013.dbf", as.is=T)
      fix <- read.dbf("tcfix.dbf", as.is=T)

      clave_deri <- read.csv("clave_deri.csv", as.is=T)
      
      # ===== Codigo =====
      # apply(data, 2, function(x) any(is.na(x)))
      
      # Colocamos los registros que contengan casos incompletos (con NA's)
      raros <- data[!complete.cases(data[, 1:6]), ]
      # Escogemos los registros que contengan casos completos (sin NA's)
      data <- data[complete.cases(data[, 1:6]), ]
      # Arrojamos un mensaje en caso de que existan casos incompletos
      if(nrow(raros)!=0){
            message("Existen registros incompletos para Swaps1 Contra")
      }
      
      # ===== Tipo de Institucion =====
      data$TIPO_INST <- NA
      data$TIPO_INST[substr(data$INSTI, 1, 3) == "013"] <- "CB"
      data$TIPO_INST[substr(data$INSTI, 1, 3) != "013"] <- "BM_BD"
      
      # ===== UDIS y FIX =====
      data$UDIS <- udis$CIERRE[match(as.Date(data$FE_CON_OPE), as.Date(udis$FE_PUBLI))] # Buscar UDIS y unir con datos 
      data$FIX <- fix$CIERRE[match(as.Date(data$FE_CON_OPE), as.Date(fix$FE_PUBLI))] # Buscar FIX y unir con datos
      
      # ===== IMPORTE =====
      data$IMPORTE <- NA
      
      data$IMPORTE[data$MDA_IMP=="MXP"] <- data$C_IMP_BASE[data$MDA_IMP=="MXP"]/2000
      data$IMPORTE[data$MDA_IMP=="UDI"] <- data$C_IMP_BASE[data$MDA_IMP=="UDI"]*data$UDIS[data$MDA_IMP=="UDI"]/2000
      data$IMPORTE[data$MDA_IMP=="USD"] <- data$C_IMP_BASE[data$MDA_IMP=="USD"]*data$FIX[data$MDA_IMP=="USD"]/2000
      
      # ===== Sector ====
      data$SECTOR <- "Mercado Mexicano de Derivados"
      
      # ===== WRITE =====
      # Escribe el cuadro (.xlsx) en el directorio de trabajo
      write.dbf(data, paste("swaps1_contra_", format(Sys.Date()[1], "%d%m%Y"), ".dbf", sep=""))
      
      if(nrow(raros)!=0){
            resultado <- list(cuadro=data, raros=raros)
            invisible(resultado)
      } else{
            invisible(data)
      }
}