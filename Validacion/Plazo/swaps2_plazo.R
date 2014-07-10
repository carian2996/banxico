# Ian Castillo Rosales (BANXICO\T41348)
# Gerencia de Información del Sistema Financiero
# Subgerencia de Información de Moneda Extranjera y Derivados
# 
# Validación de información para operaciones con swaps (plazo)
# 040614 - 010714

swaps2_plazo <- function(ruta){
      
      # ENTRADA
      # ruta = Ruta donde se encuentran los datos para los cálculos
            # swaps2.dbf
            # udi2013.dbf
            # fix.dbf
      
      # SALIDA
      # swaps2_plazo_[fecha].dbf - Archivo tipo .dbf con los resultados
      
      # ===== Librerias y directorios =====
      setwd(paste(ruta, "SWAPS/", sep="")) # ¿Dónde están mis datos?
      library(foreign) # Librería necesaria para cargar los datos
      options(scipen=999, digits=8) # Quita la notación exp y trunca a 4 decimales
      
      # ===== Carga de datos =====
      data <- read.dbf("swaps2.dbf", as.is=T)
      gc()
      udis <- read.dbf("udi2013.dbf", as.is=T)
      fix <- read.dbf("tcfix.dbf", as.is=T)
      
      # ===== Código =====
      # apply(data, 2, function(x) any(is.na(x)))

      # Colocamos los registros que contengan casos incompletos (con NA's)
      raros <- data[!complete.cases(data[, -c(6, 7)]), ]
      # Escogemos los registros que contengan casos completos (sin NA's)
      data <- data[complete.cases(data[, -c(6, 7)]), ]
      # Arrojamos un mensaje en caso de que existan casos incompletos
      if(nrow(raros)!=0){
            message("Existen registros incompletos para Swaps2 Plazo")
      }
      
      # ===== Tipo de Institución =====
      data$TIPO_INST <- NA
      data$TIPO_INST[substr(data$INSTI, 1, 3) == "013"] <- "CB"
      data$TIPO_INST[substr(data$INSTI, 1, 3) != "013"] <- "BM_BD"
      
      # ===== UDIS y FIX =====
      data$UDIS <- udis$CIERRE[match(as.Date(data$FE_CON_OPE), as.Date(udis$FE_PUBLI))]
      data$FIX <- fix$CIERRE[match(as.Date(data$FE_CON_OPE), as.Date(fix$FE_PUBLI))]
      
      # ===== IMPORTE =====
      data$IMPORTE <- NA
      
      # ===== Importe =====
      data$IMPORTE[data$MDA_IMP=="MXP" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BASE[data$MDA_IMP=="MXP" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/2000
      data$IMPORTE[data$MDA_IMP=="UDI" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BASE[data$MDA_IMP=="UDI" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$UDIS[data$MDA_IMP=="UDI" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/2000
      data$IMPORTE[data$MDA_IMP=="USD" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BASE[data$MDA_IMP=="USD" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$FIX[data$MDA_IMP=="USD" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/2000
      
      if(any(is.na(match(unique(data$MDA_IMP), c("MXP", "UDI", "USD"))))){
            data$IMPORTE[!(data$MDA_IMP=="MXP" | data$MDA_IMP=="UDI" | data$MDA_IMP=="USD") & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BA_D[!(data$MDA_IMP=="MXP" | data$MDA_IMP=="UDI" | data$MDA_IMP=="USD") & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$FIX[!(data$MDA_IMP=="MXP" | data$MDA_IMP=="UDI" | data$MDA_IMP=="USD") & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/2000
      }
      
      data$IMPORTE[data$MDA_IMP=="MXP" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BASE[data$MDA_IMP=="MXP" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/1000
      data$IMPORTE[data$MDA_IMP=="UDI" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BASE[data$MDA_IMP=="UDI" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$UDIS[data$MDA_IMP=="UDI" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/1000
      data$IMPORTE[data$MDA_IMP=="USD" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BASE[data$MDA_IMP=="USD" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$FIX[data$MDA_IMP=="USD" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/1000
      
      if(any(is.na(match(unique(data$MDA_IMP), c("MXP", "UDI", "USD"))))){
            data$IMPORTE[!(data$MDA_IMP=="MXP" | data$MDA_IMP=="UDI" | data$MDA_IMP=="USD") & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BA_D[!(data$MDA_IMP=="MXP" | data$MDA_IMP=="UDI" | data$MDA_IMP=="USD") & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$FIX[!(data$MDA_IMP=="MXP" | data$MDA_IMP=="UDI" | data$MDA_IMP=="USD") & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/1000
      }
      
      # ===== Plazo =====
      # Calcula máximo entre fecha de entrega y fecha de recibo
      fechas <- data.frame(FE_ORI_RE=as.Date(data$FE_ORI_RE), FE_ORI_EN=as.Date(data$FE_ORI_EN))
      nueva_fecha <- apply(fechas, 1, max)
      # Realiza la diferencia entre fechas, excepto cuando no haya fecha de liquidación
      data$PLAZO <- as.numeric(as.Date(nueva_fecha) - as.Date(data$FE_CON_OPE))
      
      if(any(data$PLAZO < 0)){
            data$PLAZO[data$PLAZO < 0] <- 0
            message("Existen plazos negativos en SWAPS2 Plazo")
      }
      
      # ===== BANDA =====
      # Matriz de bandas
      bandas <- matrix(0, nrow=14, ncol=2)
      bandas[, 1] <- c(0, 8, 32, 93, 185, 367, 732, 1097, 1462, 1828, 2558, 
                       3654, 5480, 7306)
      bandas[, 2] <- c("1 a 7", "8 a 31", "32 a 92", "93 a 184", "185 a 366", 
                       "367 a 731", "732 a 1096", "1097 a 1461", "1462 a 1827",
                       "1828 a 2557", "2558 a 3653", "3654 a 5479", "5480 a 7305", 
                       "mas de 7306")
      
      data$BANDA <- NA
      data$BANDA <- bandas[, 2][findInterval(data$PLAZO, as.numeric(bandas[, 1]))]
      
      # ===== WRITE =====
      # Escribe el cuadro (.xlsx) en el directorio de trabajo
      write.dbf(data, paste("swaps2_plazo_", format(Sys.Date()[1], "%d%m%Y"), ".dbf", sep=""))
      
      if(nrow(raros)!=0){
            resultado <- list(cuadro=data, raros=raros)
            invisible(resultado)
      } else{
            invisible(data)
      }
}