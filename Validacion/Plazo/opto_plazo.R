# Ian Castillo Rosales (BANXICO\T41348)
# Gerencia de Información del Sistema Financiero
# Subgerencia de Información de Moneda Extranjera y Derivados
# 
# Validación de información para operaciones con opciones
# 230614 - 020714

opto_plazo <- function(ruta){
      
      # ENTRADA
      # ruta = Ruta donde se encuentran los datos para los calculos
            # opto.dbf
            # udi2013.dbf
            # fix.dbf
      
      # SALIDA
      # opto_plazo_[fecha].dbf - Archivo tipo .dbf con los resultados
      
      # ===== Librerias y directorios =====
      setwd(paste(ruta, "OPTO/", sep="")) # ¿Dónde están mis datos?
      library(foreign) # Libreria necesaria para cargar los datos
      options(scipen=999, digits=8)

      # ===== Carga de datos =====
      data <- read.dbf("opto.dbf", as.is=T)
      gc()
      udis <- read.dbf("udi2013.dbf", as.is=T)
      fix <- read.dbf("tcfix.dbf", as.is=T)
      
      # ===== Codigo ====
      # apply(data, 2, function(x) any(is.na(x)))
      
      # Colocamos los registros que contengan casos incompletos (con NA's)
      raros <- data[!complete.cases(data[, -c(3, 13, 15, 17)]), ]
      # Escogemos los registros que contengan casos completos (sin NA's)
      data <- data[complete.cases(data[, -c(3, 13, 15, 17)]), ]
      # Arrojamos un mensaje en caso de que existan casos incompletos
      if(nrow(raros)!=0){
            message("¡Existen registros incompletos en OPTO Plazo!")
      }
      
      # ===== Tipo de Institucion =====
      data$TIPO_INST <- NA
      data$TIPO_INST[substr(data$INSTI, 1, 3) == "013"] <- "CB"
      data$TIPO_INST[substr(data$INSTI, 1, 3) != "013"] <- "BM_BD"
      
      # ===== UDIS =====
      data$UDIS <- udis$CIERRE[match(as.Date(data$FE_CON_OPE), as.Date(udis$FE_PUBLI))]
      data$FIX <- fix$CIERRE[match(as.Date(data$FE_CON_OPE), as.Date(fix$FE_PUBLI))]
      
      # ===== Importe =====
      data$IMPORTE <- NA
      
      data$IMPORTE[data$MDO!="E" & data$MDA_IMP=="01"] <- data$C_IMP_BASE[data$MDO!="E" & data$MDA_IMP=="01"]/2000
      data$IMPORTE[data$MDO!="E" & data$MDA_IMP=="02"] <- data$C_IMP_BASE[data$MDO!="E" & data$MDA_IMP=="02"]*data$UDIS[data$MDO!="E" & data$MDA_IMP=="02"]/2000
      data$IMPORTE[data$MDO!="E" & data$MDA_IMP=="10"] <- data$C_IMP_BASE[data$MDO!="E" & data$MDA_IMP=="10"]*data$FIX[data$MDO!="E" & data$MDA_IMP=="10"]/2000
      
      data$IMPORTE[data$MDO=="E" & data$MDA_IMP=="01" & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")] <- data$C_IMP_BASE[data$MDO=="E" & data$MDA_IMP=="01" & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]/2000
      data$IMPORTE[data$MDO=="E" & data$MDA_IMP=="02" & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")] <- data$C_IMP_BASE[data$MDO=="E" & data$MDA_IMP=="02" & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]*data$UDIS[data$MDO=="E" & data$MDA_IMP=="02" & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]/2000
      data$IMPORTE[data$MDO=="E" & data$MDA_IMP=="10" & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")] <- data$C_IMP_BASE[data$MDO=="E" & data$MDA_IMP=="10" & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]*data$FIX[data$MDO=="E" & data$MDA_IMP=="10" & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]/2000
      if(any(is.na(match(unique(data$MDA_IMP), c("01", "02", "10"))))){
            data$IMPORTE[!(data$MDA_IMP=="01" | data$MDA_IMP=="02" | data$MDA_IMP=="10") & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BA_D[!(data$MDA_IMP=="01" | data$MDA_IMP=="02" | data$MDA_IMP=="10") & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$FIX[!(data$MDA_IMP=="01" | data$MDA_IMP=="02" | data$MDA_IMP=="10") & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/2000
      }
      
      data$IMPORTE[data$MDO=="E" & data$MDA_IMP=="01" & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")] <- data$C_IMP_BASE[data$MDO=="E" & data$MDA_IMP=="01" & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]/1000
      data$IMPORTE[data$MDO=="E" & data$MDA_IMP=="02" & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")] <- data$C_IMP_BASE[data$MDO=="E" & data$MDA_IMP=="02" & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]*data$UDIS[data$MDO=="E" & data$MDA_IMP=="02" & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]/1000
      data$IMPORTE[data$MDO=="E" & data$MDA_IMP=="10" & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")] <- data$C_IMP_BASE[data$MDO=="E" & data$MDA_IMP=="10" & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]*data$FIX[data$MDO=="E" & data$MDA_IMP=="10" & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]/1000
      if(any(is.na(match(unique(data$MDA_IMP), c("01", "02", "10"))))){
            data$IMPORTE[!(data$MDA_IMP=="01" | data$MDA_IMP=="02" | data$MDA_IMP=="10") & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BA_D[!(data$MDA_IMP=="01" | data$MDA_IMP=="02" | data$MDA_IMP=="10") & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$FIX[!(data$MDA_IMP=="01" | data$MDA_IMP=="02" | data$MDA_IMP=="10") & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/1000
      }
      
      data$IMPORTE[data$INDICA_CF=="S" & data$CONSEC_CF!=1] <- 0

      # ===== Clase de Operaciones =====
      data$CLASE_OPER <- "Exotica"
      
      data$CLASE_OPER[data$OPC=="A" & data$INDICA_CF=="N" & is.na(data$OPE_SUBY)] <- "Estandar"
      data$CLASE_OPER[data$OPC=="E" & data$INDICA_CF=="N" & is.na(data$OPE_SUBY)] <- "Estandar"
      data$CLASE_OPER[data$OPC=="A" & is.na(data$INDICA_CF) & is.na(data$OPE_SUBY)] <- "Estandar"
      data$CLASE_OPER[data$OPC=="E" & is.na(data$INDICA_CF) & is.na(data$OPE_SUBY)] <- "Estandar"
      
      data$CLASE_OPER[data$CLASE_OPE == "WARRANT"] <- "Warrant"
      data$CLASE_OPER[data$CLASE_OPE == "SWAPTION"] <- "Exotica"
      # ===== Plazo =====
      data$PLAZO <- NA # Crear columna de plazo
      
      data$PLAZO[data$INDICA_CF=="S" & !is.na(data$INDICA_CF) & data$CONSEC_CF==1] <- as.numeric(as.Date(data$FE_VEN_CF[data$INDICA_CF=="S" & !is.na(data$INDICA_CF) & data$CONSEC_CF==1]) - as.Date(data$FE_CON_OPE[data$INDICA_CF=="S" & !is.na(data$INDICA_CF) & data$CONSEC_CF==1]))
      data$PLAZO[data$INDICA_CF!="S" & !is.na(data$INDICA_CF)] <- as.numeric(as.Date(data$FE_VEN_ORI[data$INDICA_CF!="S" & !is.na(data$INDICA_CF)]) - as.Date(data$FE_CON_OPE[data$INDICA_CF!="S" & !is.na(data$INDICA_CF)]))
      data$PLAZO[is.na(data$INDICA_CF)] <- as.numeric(as.Date(data$FE_VEN_ORI[is.na(data$INDICA_CF)]) - as.Date(data$FE_CON_OPE[is.na(data$INDICA_CF)]))
      
      if(sum(data$PLAZO[!is.na(data$PLAZO)] < 0)!=0){
            data$PLAZO[data$PLAZO < 0] <- 0
            message("¡Existen plazos negativos en OPTO Plazo!")
      }
      
      # ===== BANDA =====
      bandas <- matrix(0, nrow=14, ncol=2)
      bandas[, 1] <- c(0, 8, 32, 93, 185, 367, 732, 1097, 1462, 1828, 2558, 
                       3654, 5480, 7306)
      bandas[, 2] <- c("1 a 7", "8 a 31", "32 a 92", "93 a 184", "185 a 366", 
                       "367 a 731", "732 a 1096", "1097 a 1461", "1462 a 1827", 
                       "1828 a 2557", "2558 a 3653", "3654 a 5479", "5480 a 7305", 
                       "Más de 7306")
      
      # Crear y renombrar columna de bandas
      data$BANDA <- NA
      data$BANDA <- bandas[, 2][findInterval(data$PLAZO, as.numeric(bandas[, 1]))]
      
      # ===== WRITE =====
      # Escribe el cuadro en el directorio de trabajo
      write.dbf(data, paste("opto_plazo_", format(Sys.Date()[1], "%d%m%Y"), ".dbf", sep=""))
      
      if(nrow(raros)!=0){
            resultado <- list(cuadro=data, raros=raros)
            invisible(resultado)
      } else{
            invisible(data)
      }
}