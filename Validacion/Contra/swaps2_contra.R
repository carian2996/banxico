# Ian Castillo Rosales
# 26062014

swaps2_contra <- function(){
      # ===== SWAPS 2 =====
      # SALIDA
      # swaps2_[fecha].dbf - Archivo tipo .dbf con los resultados
      
      # ===== Librerias y directorios =====
      setwd("/Volumes/IAN/Estadisticas/Contraparte/SWAPS") # ¿Dónde están mis datos?
      library(foreign) # Libreria necesaria para cargar los datos
      options(scipen=999)
      options(encoding="UTF-8")
      
      # ===== Carga de datos =====
      data <- read.dbf("swaps2.dbf", as.is=T)
      gc()
      udis <- read.dbf("udi2013.dbf", as.is=T)
      fix <- read.dbf("tcfix.dbf", as.is=T)

      clave_deri <- read.csv("clave_deri.csv", as.is=T)
      
      # ===== Código =====
      data <- data[complete.cases(data$FE_CON_OPE, data$C_IMP_BASE, data$MDA_IMP, 
                                  data$C_IMP_BA_D, data$CONT, data$TIP_CONT), ]
      raros <- data[!complete.cases(data$FE_CON_OPE, data$C_IMP_BASE, data$MDA_IMP, 
                                   data$C_IMP_BA_D, data$CONT, data$TIP_CONT), ]
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
      
      # ===== Importe =====
      data$IMPORTE[data$MDA_IMP=="MXP" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BASE[data$MDA_IMP=="MXP" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/2000
      data$IMPORTE[data$MDA_IMP=="UDI" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BASE[data$MDA_IMP=="UDI" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$UDIS[data$MDA_IMP=="UDI" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/2000
      data$IMPORTE[data$MDA_IMP=="USD" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BASE[data$MDA_IMP=="USD" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$FIX[data$MDA_IMP=="USD" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/2000
      data$IMPORTE[!(data$MDA_IMP=="MXP" | data$MDA_IMP=="UDI" | data$MDA_IMP=="USD") & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BA_D[!(data$MDA_IMP=="MXP" | data$MDA_IMP=="UDI" | data$MDA_IMP=="USD") & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$FIX[!(data$MDA_IMP=="MXP" | data$MDA_IMP=="UDI" | data$MDA_IMP=="USD") & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/2000
      
      data$IMPORTE[data$MDA_IMP=="MXP" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BASE[data$MDA_IMP=="MXP" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/1000
      data$IMPORTE[data$MDA_IMP=="UDI" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BASE[data$MDA_IMP=="UDI" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$UDIS[data$MDA_IMP=="UDI" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/1000
      data$IMPORTE[data$MDA_IMP=="USD" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BASE[data$MDA_IMP=="USD" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$FIX[data$MDA_IMP=="USD" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/1000
      data$IMPORTE[!(data$MDA_IMP=="MXP" | data$MDA_IMP=="UDI" | data$MDA_IMP=="USD") & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BA_D[!(data$MDA_IMP=="MXP" | data$MDA_IMP=="UDI" | data$MDA_IMP=="USD") & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$FIX[!(data$MDA_IMP=="MXP" | data$MDA_IMP=="UDI" | data$MDA_IMP=="USD") & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/1000
      
      # ===== Tipo de Entidad y Residencia ====
      data$TIPO_ENTE <- NA
      data$RESI <- NA
      
      data$TIPO_ENTE <- clave_deri$tipo_ente[match(data$CONT, clave_deri$clave_deri)]
      data$RESI <- clave_deri$residencia[match(data$CONT, clave_deri$clave_deri)]
      
      # ===== Sector =====
      data$SECTOR[data$TIPO_ENTE==4 &data$RESI=="MX" ] <- "Bancos Múltiples"
      data$SECTOR[data$TIPO_ENTE==5 &data$RESI=="MX" ] <- "Bancos de Desarrollo"
      data$SECTOR[data$TIPO_ENTE==6 &data$RESI=="MX" ] <- "Casas de Bolsa"
      
      data$SECTOR[(data$TIPO_ENTE==19  | data$TIPO_ENTE==48) & data$RESI=="MX" ] <- "Sociedades y Fondos de Inversión"
      data$SECTOR[(data$TIPO_ENTE==20  | data$TIPO_ENTE==37) & data$RESI=="MX" ] <- "SIEFORES"
      data$SECTOR[(data$TIPO_ENTE==38  | data$TIPO_ENTE==46) & data$RESI=="MX" ] <- "SOFOMES"
      
      data$SECTOR[is.na(match(data$TIPO_ENTE, c(1, 29, 30, 34, 44, 54, 43, 49, 25, 17, 26, 45, 4, 5, 6, 19, 48, 20, 37, 38, 46))) & data$RESI=="MX" ] <- "Otras entidades financieras nacionales"
      data$SECTOR[!is.na(match(data$TIPO_ENTE, c(1, 29, 30, 34, 44, 54))) & data$RESI=="MX" ] <- "Gobierno Federal y Entidades Descentralizadas"
      
      data$SECTOR[(data$TIPO_ENTE==43  | data$TIPO_ENTE==49) & data$RESI=="MX" ] <- "Gobiernos y Entidades Estatales o Municipales"
      
      data$SECTOR[data$TIPO_ENTE==25 &data$RESI=="MX" ] <- "Personas físicas residentes en México"
      
      data$SECTOR[(data$TIPO_ENTE==17  | data$TIPO_ENTE==26 | data$TIPO_ENTE==45) & data$RESI=="MX" ] <- "Empresas Privadas No Financieras"
      
      data$SECTOR[(data$TIPO_ENTE==4  | data$TIPO_ENTE==31) & data$RESI=="US" ] <- "Bancos comerciales en E.U.A."
      data$SECTOR[(data$TIPO_ENTE==4  | data$TIPO_ENTE==31) & !is.na(match(data$RESI, c("AT", "BE", "BG", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE", "GB", "VG"))) ] <- "Bancos comerciales en la Unión Europea"
      data$SECTOR[(data$TIPO_ENTE==4  | data$TIPO_ENTE==31) & !is.na(match(data$RESI, c("AR", "BO", "BR", "CL", "CO", "CU", "DO", "EC", "SV", "GT", "HN", "NI", "PA", "PY", "PE", "PR", "UY", "VE"))) ] <- "Bancos comerciales en América Latina"
      data$SECTOR[(data$TIPO_ENTE==4  | data$TIPO_ENTE==31) & !is.na(match(data$RESI, c("BS", "CA", "CH", "ND", "KY", "AN", "AU", "HK", "SG"))) ] <- "Bancos comerciales en otras regiones y países"
      
      data$SECTOR[!is.na(match(data$TIPO_ENTE, c(10,11,50,51,53))) & data$RESI=="US" ] <- "Otras entidades financieras en E.U.A."
      data$SECTOR[!is.na(match(data$TIPO_ENTE, c(10,11,50,51,53))) & !is.na(match(data$RESI, c("AT", "BE", "BG", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE", "GB", "VG"))) ] <- "Otras entidades financieras en la Unión Europea"
      data$SECTOR[!is.na(match(data$TIPO_ENTE, c(10,11,50,51,53))) & !is.na(match(data$RESI, c("AR", "BO", "BR", "CL", "CO", "CU", "DO", "EC", "SV", "GT", "HN", "NI", "PA", "PY", "PE", "PR", "UY", "VE"))) ] <- "Otras entidades financieras en América Latina"
      data$SECTOR[!is.na(match(data$TIPO_ENTE, c(10,11,50,51,53))) & !is.na(match(data$RESI, c("BS", "CA", "CH", "ND", "KY", "AN", "AU", "HK", "SG"))) ] <- "Otras entidades financieras extranjeras"
      
      data$SECTOR[data$TIPO_ENTE==47 & data$RESI=="US" ] <- "Gobiernos Federal y Local de los E.U.A."
      data$SECTOR[data$TIPO_ENTE==47 & !is.na(match(data$RESI, c("AT", "BE", "BG", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE", "GB", "VG"))) ] <- "Gobiernos Federal y Local de la Unión Europea"
      data$SECTOR[data$TIPO_ENTE==47 & !is.na(match(data$RESI, c("AR", "BO", "BR", "CL", "CO", "CU", "DO", "EC", "SV", "GT", "HN", "NI", "PA", "PY", "PE", "PR", "UY", "VE"))) ] <- "Gobiernos Federal y Local en América Latina"
      data$SECTOR[data$TIPO_ENTE==47 & !is.na(match(data$RESI, c("BS", "CA", "CH", "ND", "KY", "AN", "AU", "HK", "SG"))) ] <- "Gobiernos Federal y Local en otras regiones y países"
      
      data$SECTOR[data$TIPO_ENTE==28 & data$RESI=="US" ] <- "Empresas privadas no financieras en E.U.A."
      data$SECTOR[data$TIPO_ENTE==28 & !is.na(match(data$RESI, c("AT", "BE", "BG", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE", "GB", "VG"))) ] <- "Empresas privadas no financieras en la Unión Europea"
      data$SECTOR[data$TIPO_ENTE==28 & !is.na(match(data$RESI, c("AR", "BO", "BR", "CL", "CO", "CU", "DO", "EC", "SV", "GT", "HN", "NI", "PA", "PY", "PE", "PR", "UY", "VE"))) ] <- "Empresas privadas no financieras en América Latina"
      data$SECTOR[data$TIPO_ENTE==28 & !is.na(match(data$RESI, c("BS", "CA", "CH", "ND", "KY", "AN", "AU", "HK", "SG"))) ] <- "Empresas privadas no financieras en otras regiones y países"
      
      data$SECTOR[data$TIPO_ENTE==27] <- "Personas físicas residentes en el extranjero"
      
      # ===== WRITE =====
      # Escribe el cuadro (.dbf) en el directorio de trabajo
      write.dbf(data, paste("swaps2_contra_", format(Sys.Date()[1], "%d_%m_%Y"), ".dbf", sep=""))
      
      data
}