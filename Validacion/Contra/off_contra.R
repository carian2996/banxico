# Ian Castillo Rosales (BANXICO\T41348)
# Gerencia de Información del Sistema Financiero
# Subgerencia de Información de Moneda Extranjera y Derivados
# 
# Validación de información para operaciones con opciones
# 090614 - 010714

off_contra <- function(ruta){
      
      # ENTRADA
      # ruta = Ruta donde se encuentran los datos para los calculos
            # derivado.dbf
            # udi2013.dbf
            # fix.dbf
      
      # SALIDA
      # off_contra_[fecha].dbf - Archivo tipo .dbf con los resultados
      
      # ===== Librerias y directorios =====
      setwd(paste(ruta, "/OFF/", sep="")) # ¿Dónde están mis datos?
      library(foreign) # Libreria necesaria para cargar los datos
      options(scipen=999, digits=8)
      
      # ===== Carga de datos =====
      data <- read.dbf("derivado.dbf", as.is=T)
      gc()
      udis <- read.dbf("udi2013.dbf", as.is=T)
      fix <- read.dbf("tcfix.dbf", as.is=T)
      
      clave_deri <- read.csv("clave_deri.csv", as.is=T)
      
      # ===== Código =====
      # apply(data, 2, function(x) any(is.na(x)))
      
      raros <- data[!complete.cases(data[, names(data)[c(1:8, 12)]]), ]
      data <- data[complete.cases(data[, names(data)[c(1:8, 12)]]), ]
      
      if(nrow(raros)!=0){
            message("Existen registros incompletos en OFF Contraparte")
      }
      
      # ===== UDIS y FIX =====
      data$UDIS <- udis$CIERRE[match(as.Date(data$FE_CON_OPE), as.Date(udis$FE_PUBLI))] # Buscar UDIS y unir con datos
      data$FIX <- fix$CIERRE[match(as.Date(data$FE_CON_OPE), as.Date(fix$FE_PUBLI))] # Buscar FIX y unir con datos
      
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
      
      # ===== Tipo de Entidad y Residencia ====
      data$TIPO_ENTE <- NA
      data$RESI <- NA
      
      data$TIPO_ENTE[!is.na(data$CONT)] <- clave_deri$tipo_ente[match(data$CONT[!is.na(data$CONT)], clave_deri$clave_deri)]
      data$RESI[!is.na(data$CONT)] <- clave_deri$residencia[match(data$CONT[!is.na(data$CONT)], clave_deri$clave_deri)]
      
      # ===== Sector =====
      data$SECTOR <- NA
      
      # ===== FORWARDS =====
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & data$TIPO_ENTE==4 &data$RESI=="MX" ] <- "Bancos Múltiples"
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & data$TIPO_ENTE==5 &data$RESI=="MX" ] <- "Bancos de Desarrollo"
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & data$TIPO_ENTE==6 &data$RESI=="MX" ] <- "Casas de Bolsa"
      
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & (data$TIPO_ENTE==19  | data$TIPO_ENTE==48) & data$RESI=="MX" ] <- "Sociedades y Fondos de Inversión"
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & (data$TIPO_ENTE==20  | data$TIPO_ENTE==37) & data$RESI=="MX" ] <- "SIEFORES"
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & (data$TIPO_ENTE==38  | data$TIPO_ENTE==46) & data$RESI=="MX" ] <- "SOFOMES"
      
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & is.na(match(data$TIPO_ENTE, c(1, 29, 30, 34, 44, 54, 43, 49, 25, 17, 26, 45, 4, 5, 6, 19, 48, 20, 37, 38, 46))) & data$RESI=="MX" ] <- "Otras entidades financieras nacionales"
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & !is.na(match(data$TIPO_ENTE, c(1, 29, 30, 34, 44, 54))) & data$RESI=="MX" ] <- "Gobierno Federal y Entidades Descentralizadas"
      
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & (data$TIPO_ENTE==43  | data$TIPO_ENTE==49) & data$RESI=="MX" ] <- "Gobiernos y Entidades Estatales o Municipales"
      
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & data$TIPO_ENTE==25 & data$RESI=="MX" ] <- "Personas físicas residentes en México"
      
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & (data$TIPO_ENTE==17  | data$TIPO_ENTE==26 | data$TIPO_ENTE==45) & data$RESI=="MX" ] <- "Empresas Privadas No Financieras"
      
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & (data$TIPO_ENTE==4  | data$TIPO_ENTE==31) & data$RESI=="US" ] <- "Bancos comerciales en E.U.A."
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & (data$TIPO_ENTE==4  | data$TIPO_ENTE==31) & !is.na(match(data$RESI, c("AT", "BE", "BG", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE", "GB", "VG"))) ] <- "Bancos comerciales en la Unión Europea"
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & (data$TIPO_ENTE==4  | data$TIPO_ENTE==31) & !is.na(match(data$RESI, c("AR", "BO", "BR", "CL", "CO", "CU", "DO", "EC", "SV", "GT", "HN", "NI", "PA", "PY", "PE", "PR", "UY", "VE"))) ] <- "Bancos comerciales en América Latina"
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & (data$TIPO_ENTE==4  | data$TIPO_ENTE==31) & !is.na(match(data$RESI, c("BS", "CA", "CH", "ND", "KY", "AN", "AU", "HK", "SG"))) ] <- "Bancos comerciales en otras regiones y países"
      
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & !is.na(match(data$TIPO_ENTE, c(10,11,50,51,53))) & data$RESI=="US" ] <- "Otras entidades financieras en E.U.A."
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & !is.na(match(data$TIPO_ENTE, c(10,11,50,51,53))) & !is.na(match(data$RESI, c("AT", "BE", "BG", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE", "GB", "VG"))) ] <- "Otras entidades financieras en la Unión Europea"
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & !is.na(match(data$TIPO_ENTE, c(10,11,50,51,53))) & !is.na(match(data$RESI, c("AR", "BO", "BR", "CL", "CO", "CU", "DO", "EC", "SV", "GT", "HN", "NI", "PA", "PY", "PE", "PR", "UY", "VE"))) ] <- "Otras entidades financieras en América Latina"
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & !is.na(match(data$TIPO_ENTE, c(10,11,50,51,53))) & !is.na(match(data$RESI, c("BS", "CA", "CH", "ND", "KY", "AN", "AU", "HK", "SG"))) ] <- "Otras entidades financieras extranjeras"
      
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & data$TIPO_ENTE==47 & data$RESI=="US" ] <- "Gobiernos Federal y Local de los E.U.A."
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & data$TIPO_ENTE==47 & !is.na(match(data$RESI, c("AT", "BE", "BG", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE", "GB", "VG"))) ] <- "Gobiernos Federal y Local de la Unión Europea"
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & data$TIPO_ENTE==47 & !is.na(match(data$RESI, c("AR", "BO", "BR", "CL", "CO", "CU", "DO", "EC", "SV", "GT", "HN", "NI", "PA", "PY", "PE", "PR", "UY", "VE"))) ] <- "Gobiernos Federal y Local en América Latina"
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & data$TIPO_ENTE==47 & !is.na(match(data$RESI, c("BS", "CA", "CH", "ND", "KY", "AN", "AU", "HK", "SG"))) ] <- "Gobiernos Federal y Local en otras regiones y países"
      
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & data$TIPO_ENTE==28 & data$RESI=="US" ] <- "Empresas privadas no financieras en E.U.A."
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & data$TIPO_ENTE==28 & !is.na(match(data$RESI, c("AT", "BE", "BG", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE", "GB", "VG"))) ] <- "Empresas privadas no financieras en la Unión Europea"
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & data$TIPO_ENTE==28 & !is.na(match(data$RESI, c("AR", "BO", "BR", "CL", "CO", "CU", "DO", "EC", "SV", "GT", "HN", "NI", "PA", "PY", "PE", "PR", "UY", "VE"))) ] <- "Empresas privadas no financieras en América Latina"
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & data$TIPO_ENTE==28 & !is.na(match(data$RESI, c("BS", "CA", "CH", "ND", "KY", "AN", "AU", "HK", "SG"))) ] <- "Empresas privadas no financieras en otras regiones y países"
      
      data$SECTOR[data$CLASE_OPE=="FORWARD" & (data$ID_SPOT=="N" | is.na(data$ID_SPOT)) & data$TIPO_ENTE==27] <- "Personas físicas residentes en el extranjero"
      
      # ===== FUTUROS =====
      data$SECTOR[data$CLASE_OPE=="FUTURO" & data$MDO=="MEX" ] <- "Mercado Mexicano de Derivados"
      data$SECTOR[data$CLASE_OPE=="FUTURO" & data$MDO=="CME" ] <- "Chicago Mercantile Exchange"
      data$SECTOR[data$CLASE_OPE=="FUTURO" & data$MDO=="CBE" ] <- "Chicago Board of Trade"
      data$SECTOR[data$CLASE_OPE=="FUTURO" & data$MDO=="ICE" ] <- "Intercontinental Futures and Options Exchange"
      data$SECTOR[data$CLASE_OPE=="FUTURO" & data$MDO=="NYM" ] <- "New York Mercantile Exchange"
      data$SECTOR[data$CLASE_OPE=="FUTURO" & data$MDO=="BMF" ] <- "Bolsa de Mercaderias y Futuros"
      data$SECTOR[data$CLASE_OPE=="FUTURO" & data$MDO=="LFF" ] <- "London International Financial Futures and Options Exchange"
      data$SECTOR[data$CLASE_OPE=="FUTURO" & data$MDO=="ASE" ] <- "Athens Stock Exchange"
      data$SECTOR[data$CLASE_OPE=="FUTURO" & data$MDO=="MEF" ] <- "Sociedad Rectora del Mercado de Productos Derivados"
      data$SECTOR[data$CLASE_OPE=="FUTURO" & data$MDO=="EUX" ] <- "Europe Global Financial Marketplace"
      data$SECTOR[data$CLASE_OPE=="FUTURO" & is.na(match(data$MDO, c("MEX","CME","CBE","ICE","NYM","LFF","ASE","MEF","EUX"))) ] <- "Mercados organizados en otras regiones y países"
      
      # ===== WRITE =====
      # Escribe el cuadro en el directorio de trabajo
      write.dbf(data, paste("off_contra_", format(Sys.Date()[1], "%d%m%Y"), ".dbf", sep=""))
      
      if(nrow(raros)!=0){
            resultado <- list(cuadro=data, raros=raros)
            invisible(resultado)
      } else{
            invisible(data)
      }
}