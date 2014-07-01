# Ian Castillo Rosales (BANXICO\T41348)
# Gerencia de Información del Sistema Financiero
# Subgerencia de Información de Moneda Extranjera y Derivados
# 
# Validación de información para operaciones con opciones
# 090614 - 010714

off_plazo <- function(ruta){
      
      # ENTRADA
      # ruta = Ruta donde se encuentran los datos para los calculos
            # derivado.dbf
            # udi2013.dbf
            # fix.dbf
      
      # SALIDA
      # off_plazo_[fecha].dbf - Archivo tipo .dbf con los resultados
      
      # ===== Librerias y directorios =====
      setwd(paste(ruta, "/OFF/", sep="")) # ¿Dónde están mis datos?
      library(foreign) # Libreria necesaria para cargar los datos
      options(scipen=999, digits=8)
      
      # ===== Carga de datos =====
      data <- read.dbf("derivado.dbf", as.is=T)
      gc()
      udis <- read.dbf("udi2013.dbf", as.is=T)
      fix <- read.dbf("tcfix.dbf", as.is=T)
      
      # ===== Código =====
      # apply(data, 2, function(x) any(is.na(x)))
      
      raros <- data[!complete.cases(data[, names(data)[1:8]]), ]
      data <- data[complete.cases(data[, names(data)[1:8]]), ]
      
      if(nrow(raros)!=0){
            message("Existen registros incompletos en OFF Plazo")
      }
      
      # ===== UDIS y FIX =====
      data$UDIS <- udis$CIERRE[match(as.Date(data$FE_CON_OPE), as.Date(udis$FE_PUBLI))] # Buscar UDIS y unir con datos
      data$FIX <- fix$CIERRE[match(as.Date(data$FE_CON_OPE), as.Date(fix$FE_PUBLI))] # Buscar FIX y unir con datos
      
      # ===== Plazo =====
      data$PLAZO <- NA # Crear columna de plazo
      # Realiza la diferencia entre fechas, excepto cuando no haya fecha de liquidación
      data$PLAZO <- as.numeric(as.Date(data$FE_LIQ_ORI) - as.Date(data$FE_CON_OPE))
      
      if(any(data$PLAZO < 0)){
            data$PLAZO[data$PLAZO < 0] <- 0
            message("Existen plazos negativos en OFF Plazo")
      }
      
      # ===== IMPORTE =====
      data$IMPORTE <- NA
      
      # ===== Importe FUT =====
      data$IMPORTE[data$MDO!="E" & data$MDA_IMP=="01"] <- data$C_IMP_BASE[data$MDO!="E" & data$MDA_IMP=="01"]/2000
      data$IMPORTE[data$MDO!="E" & data$MDA_IMP=="02"] <- data$C_IMP_BASE[data$MDO!="E" & data$MDA_IMP=="02"]*data$UDIS[data$MDO!="E" & data$MDA_IMP=="02"]/2000
      data$IMPORTE[data$MDO!="E" & data$MDA_IMP=="10"] <- data$C_IMP_BASE[data$MDO!="E" & data$MDA_IMP=="10"]*data$FIX[data$MDO!="E" & data$MDA_IMP=="10"]/2000
      
      # ===== Importe FWD =====
      data$IMPORTE[data$MDO=="E" & data$MDA_IMP=="01" & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")] <- data$C_IMP_BASE[data$MDO=="E" & data$MDA_IMP=="01" & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]/2000
      data$IMPORTE[data$MDO=="E" & data$MDA_IMP=="02" & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")] <- data$C_IMP_BASE[data$MDO=="E" & data$MDA_IMP=="02" & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]*data$UDIS[data$MDO=="E" & data$MDA_IMP=="02" & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]/2000
      data$IMPORTE[data$MDO=="E" & data$MDA_IMP=="10" & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")] <- data$C_IMP_BASE[data$MDO=="E" & data$MDA_IMP=="10" & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]*data$FIX[data$MDO=="E" & data$MDA_IMP=="10" & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]/2000
      data$IMPORTE[data$MDO=="E" & !(data$MDA_IMP=="01" & data$MDA_IMP=="02" & data$MDA_IMP=="10") & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")] <- data$C_IMP_BASE[data$MDO=="E" & !(data$MDA_IMP=="01" & data$MDA_IMP=="02" & data$MDA_IMP=="10") & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]*data$FIX[data$MDO=="E" & !(data$MDA_IMP=="01" & data$MDA_IMP=="02" & data$MDA_IMP=="10") & (data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]/2000
      
      data$IMPORTE[data$MDO=="E" & data$MDA_IMP=="01" & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")] <- data$C_IMP_BASE[data$MDO=="E" & data$MDA_IMP=="01" & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]/1000
      data$IMPORTE[data$MDO=="E" & data$MDA_IMP=="02" & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")] <- data$C_IMP_BASE[data$MDO=="E" & data$MDA_IMP=="02" & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]*data$UDIS[data$MDO=="E" & data$MDA_IMP=="02" & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]/1000
      data$IMPORTE[data$MDO=="E" & data$MDA_IMP=="10" & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")] <- data$C_IMP_BASE[data$MDO=="E" & data$MDA_IMP=="10" & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]*data$FIX[data$MDO=="E" & data$MDA_IMP=="10" & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]/1000
      data$IMPORTE[data$MDO=="E" & !(data$MDA_IMP=="01" & data$MDA_IMP=="02" & data$MDA_IMP=="10") & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")] <- data$C_IMP_BASE[data$MDO=="E" & !(data$MDA_IMP=="01" & data$MDA_IMP=="02" & data$MDA_IMP=="10") & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]*data$FIX[data$MDO=="E" & !(data$MDA_IMP=="01" & data$MDA_IMP=="02" & data$MDA_IMP=="10") & !(data$TIP_CONT=="RIC" | data$TIP_CONT=="RCB")]/1000
      
      data$IMPORTE[data$ID_SPOT=="S"] <- 0
      
      # ===== BANDA =====
      # Matriz de bandas
      bandas <- matrix(0, nrow=14, ncol=2)
      bandas[, 1] <- c(0, 8, 32, 93, 185, 367, 732, 1097, 1462, 1828, 2558, 
                       3654, 5480, 7306)
      bandas[, 2] <- c("1 a 7", "8 a 31", "32 a 92", "93 a 184", "185 a 366", 
                       "367 a 731", "732 a 1096", "1097 a 1461", "1462 a 1827", 
                       "1828 a 2557", "2558 a 3653", "3654 a 5479", "5480 a 7305", 
                       "Más de 7306")
      
      # Crear y renombrar columna de bandas
      data$BANDA <- NA
      # Encuentra el intervalo y pone la banda
      data$BANDA <- bandas[, 2][findInterval(data$PLAZO, as.numeric(bandas[, 1]))]
      
      # ===== WRITE =====
      # Escribe el cuadro (.dbf) en el directorio de trabajo
      write.dbf(data, paste("off_plazo_", format(Sys.Date()[1], "%d%m%Y"), ".dbf", sep=""))
      
      if(nrow(raros)!=0){
            resultado <- list(cuadro=data, raros=raros)
            invisible(resultado)
      } else{
            invisible(data)
      }
}
