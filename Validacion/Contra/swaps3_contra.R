# Ian Castillo Rosales (BANXICO\T41348)
# Gerencia de Informacion del Sistema Financiero
# Subgerencia de Informacion de Moneda Extranjera y Derivados
# 
# Validacion de informacion para operaciones con swaps (contraparte)
# 110614 - 300414

swaps3_contra <- function(ruta){
      
      # ENTRADA
      # ruta = Ruta donde se encuentran los datos para los calculos
            # swaps3.dbf
            # udi2013.dbf
            # fix.dbf
      
      # SALIDA
      # swaps3_contra_[fecha].dbf - Archivo tipo .dbf con los resultados
      
      # ===== Librerias y directorios =====
      setwd(paste(ruta, "/SWAPS/", sep="")) # ?D?nde est?n mis datos?
      library(foreign) # Libreria necesaria para cargar los datos
      options(scipen=999, digits=8) # Quita la notaci?n exp y trunca a 4 decimales
      
      # ===== Carga de datos =====
      data <- read.dbf("swaps3.dbf", as.is=T)
      gc()
      udis <- read.dbf("udi2013.dbf", as.is=T)
      fix <- read.dbf("tcfix.dbf", as.is=T)
      
      clave_deri <- read.csv("clave_deri.csv", as.is=T)
      clave_deri$clave_deri <- sprintf("%06d", as.numeric(clave_deri$clave_deri)) # Saca warning por NA's
      
      # ===== Codigo =====
      # apply(data, 2, function(x) any(is.na(x)))
      
      # Colocamos los registros que contengan casos incompletos (con NA's)
      raros <- data[!complete.cases(data[, -c(6, 10)]), ]
      # Escogemos los registros que contengan casos completos (sin NA's)
      data <- data[complete.cases(data[, -c(6, 10)]), ]
      # Arrojamos un mensaje en caso de que existan casos incompletos
      if(nrow(raros)!=0){
            message("Existen registros incompletos para Swaps3 Contra")
      }
      
      # ===== Tipo de Institucion =====
      data$TIPO_INST <- NA
      data$TIPO_INST[substr(data$INSTI, 1, 3) == "013"] <- "CB"
      data$TIPO_INST[substr(data$INSTI, 1, 3) != "013"] <- "BM_BD"
      
      # ===== UDIS y FIX =====
      data$UDIS <- udis$CIERRE[match(as.Date(data$FE_CON_OPE), as.Date(udis$FE_PUBLI))]
      data$FIX <- fix$CIERRE[match(as.Date(data$FE_CON_OPE), as.Date(fix$FE_PUBLI))]
      
      # ===== IMPORTE =====
      data$IMPORTE_R <- NA
      data$IMPORTE_E <- NA
      data$IMPORTE <- NA
      
      # ===== Importe R =====
      data$IMPORTE_R[data$MDA_IMP_RE=="MXP" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BA_R[data$MDA_IMP_RE=="MXP" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/2000
      data$IMPORTE_R[data$MDA_IMP_RE=="UDI" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BA_R[data$MDA_IMP_RE=="UDI" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$UDIS[data$MDA_IMP_RE=="UDI" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/2000
      data$IMPORTE_R[data$MDA_IMP_RE=="USD" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BA_R[data$MDA_IMP_RE=="USD" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$FIX[data$MDA_IMP_RE=="USD" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/2000
      if(any(is.na(match(unique(data$MDA_IMP_RE), c("MXP", "UDI", "USD"))))){
            data$IMPORTE_R[!(data$MDA_IMP_RE=="MXP" | data$MDA_IMP_RE=="UDI" | data$MDA_IMP_RE=="USD") & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_RE_D[!(data$MDA_IMP_RE=="MXP" | data$MDA_IMP_RE=="UDI" | data$MDA_IMP_RE=="USD") & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$FIX[!(data$MDA_IMP_RE=="MXP" | data$MDA_IMP_RE=="UDI" | data$MDA_IMP_RE=="USD") & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/2000
      }
      
      data$IMPORTE_R[data$MDA_IMP_RE=="MXP" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BA_R[data$MDA_IMP_RE=="MXP" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/1000
      data$IMPORTE_R[data$MDA_IMP_RE=="UDI" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BA_R[data$MDA_IMP_RE=="UDI" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$UDIS[data$MDA_IMP_RE=="UDI" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/1000
      data$IMPORTE_R[data$MDA_IMP_RE=="USD" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BA_R[data$MDA_IMP_RE=="USD" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$FIX[data$MDA_IMP_RE=="USD" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/1000
      if(any(is.na(match(unique(data$MDA_IMP_RE), c("MXP", "UDI", "USD"))))){
            data$IMPORTE_R[!(data$MDA_IMP_RE=="MXP" | data$MDA_IMP_RE=="UDI" | data$MDA_IMP_RE=="USD") & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_RE_D[!(data$MDA_IMP_RE=="MXP" | data$MDA_IMP_RE=="UDI" | data$MDA_IMP_RE=="USD") & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$FIX[!(data$MDA_IMP_RE=="MXP" | data$MDA_IMP_RE=="UDI" | data$MDA_IMP_RE=="USD") & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/1000
      }
      # ===== Importe E =====
      data$IMPORTE_E[data$MDA_IMP_EN=="MXP" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BA_E[data$MDA_IMP_EN=="MXP" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/2000
      data$IMPORTE_E[data$MDA_IMP_EN=="UDI" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BA_E[data$MDA_IMP_EN=="UDI" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$UDIS[data$MDA_IMP_EN=="UDI" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/2000
      data$IMPORTE_E[data$MDA_IMP_EN=="USD" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BA_E[data$MDA_IMP_EN=="USD" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$FIX[data$MDA_IMP_EN=="USD" & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/2000
      if(any(is.na(match(unique(data$MDA_IMP_EN), c("MXP", "UDI", "USD"))))){
            data$IMPORTE_E[!(data$MDA_IMP_EN=="MXP" | data$MDA_IMP_EN=="UDI" | data$MDA_IMP_EN=="USD") & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_EN_D[!(data$MDA_IMP_EN=="MXP" | data$MDA_IMP_EN=="UDI" | data$MDA_IMP_EN=="USD") & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$FIX[!(data$MDA_IMP_EN=="MXP" | data$MDA_IMP_EN=="UDI" | data$MDA_IMP_EN=="USD") & (data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/2000
      }
      
      data$IMPORTE_E[data$MDA_IMP_EN=="MXP" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BA_E[data$MDA_IMP_EN=="MXP" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/1000
      data$IMPORTE_E[data$MDA_IMP_EN=="UDI" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BA_E[data$MDA_IMP_EN=="UDI" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$UDIS[data$MDA_IMP_EN=="UDI" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/1000
      data$IMPORTE_E[data$MDA_IMP_EN=="USD" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_BA_E[data$MDA_IMP_EN=="USD" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$FIX[data$MDA_IMP_EN=="USD" & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/1000
      if(any(is.na(match(unique(data$MDA_IMP_EN), c("MXP", "UDI", "USD"))))){
            data$IMPORTE_E[!(data$MDA_IMP_EN=="MXP" | data$MDA_IMP_EN=="UDI" | data$MDA_IMP_EN=="USD") & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")] <- data$C_IMP_EN_D[!(data$MDA_IMP_EN=="MXP" | data$MDA_IMP_EN=="UDI" | data$MDA_IMP_EN=="USD") & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]*data$FIX[!(data$MDA_IMP_EN=="MXP" | data$MDA_IMP_EN=="UDI" | data$MDA_IMP_EN=="USD") & !(data$TIP_CONT=="RCB" | data$TIP_CONT=="RIC")]/1000
      }
      
      # ===== Importe =====
      data$IMPORTE <- apply(data[, c("IMPORTE_R", "IMPORTE_E")], 1, max)
      
      # ===== Tipo de Entidad y Residencia ====
      data$TIPO_ENTE <- NA
      data$RESI <- NA
      
      data$TIPO_ENTE <- clave_deri$tipo_ente[match(data$CONT, clave_deri$clave_deri)]
      data$RESI <- clave_deri$residencia[match(data$CONT, clave_deri$clave_deri)]
      
      # ===== Sector =====
      data$SECTOR[data$TIPO_ENTE==4 &data$RESI=="MX" ] <- "Bancos Multiples"
      data$SECTOR[data$TIPO_ENTE==5 &data$RESI=="MX" ] <- "Bancos de Desarrollo"
      data$SECTOR[data$TIPO_ENTE==6 &data$RESI=="MX" ] <- "Casas de Bolsa"
      
      data$SECTOR[(data$TIPO_ENTE==19  | data$TIPO_ENTE==48) & data$RESI=="MX" ] <- "Sociedades y Fondos de Inversion"
      data$SECTOR[(data$TIPO_ENTE==20  | data$TIPO_ENTE==37) & data$RESI=="MX" ] <- "SIEFORES"
      data$SECTOR[(data$TIPO_ENTE==38  | data$TIPO_ENTE==46) & data$RESI=="MX" ] <- "SOFOMES"
      
      data$SECTOR[is.na(match(data$TIPO_ENTE, c(1, 29, 30, 34, 44, 54, 43, 49, 25, 17, 26, 45, 4, 5, 6, 19, 48, 20, 37, 38, 46))) & data$RESI=="MX" ] <- "Otras entidades financieras nacionales"
      data$SECTOR[!is.na(match(data$TIPO_ENTE, c(1, 29, 30, 34, 44, 54))) & data$RESI=="MX" ] <- "Gobierno Federal y Entidades Descentralizadas"
      
      data$SECTOR[(data$TIPO_ENTE==43  | data$TIPO_ENTE==49) & data$RESI=="MX" ] <- "Gobiernos y Entidades Estatales o Municipales"
      
      data$SECTOR[data$TIPO_ENTE==25 &data$RESI=="MX" ] <- "Personas fisicas residentes en Mexico"
      
      data$SECTOR[(data$TIPO_ENTE==17  | data$TIPO_ENTE==26 | data$TIPO_ENTE==45) & data$RESI=="MX" ] <- "Empresas Privadas No Financieras"
      
      data$SECTOR[(data$TIPO_ENTE==4  | data$TIPO_ENTE==31) & data$RESI=="US" ] <- "Bancos comerciales en E.U.A."
      data$SECTOR[(data$TIPO_ENTE==4  | data$TIPO_ENTE==31) & !is.na(match(data$RESI, c("AT", "BE", "BG", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE", "GB", "VG"))) ] <- "Bancos comerciales en la Union Europea"
      data$SECTOR[(data$TIPO_ENTE==4  | data$TIPO_ENTE==31) & !is.na(match(data$RESI, c("AR", "BO", "BR", "CL", "CO", "CU", "DO", "EC", "SV", "GT", "HN", "NI", "PA", "PY", "PE", "PR", "UY", "VE"))) ] <- "Bancos comerciales en America Latina"
      data$SECTOR[(data$TIPO_ENTE==4  | data$TIPO_ENTE==31) & !is.na(match(data$RESI, c("BS", "CA", "CH", "ND", "KY", "AN", "AU", "HK", "SG"))) ] <- "Bancos comerciales en otras regiones y paises"
      
      data$SECTOR[!is.na(match(data$TIPO_ENTE, c(10,11,50,51,53))) & data$RESI=="US" ] <- "Otras entidades financieras en los E.U.A."
      data$SECTOR[!is.na(match(data$TIPO_ENTE, c(10,11,50,51,53))) & !is.na(match(data$RESI, c("AT", "BE", "BG", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE", "GB", "VG"))) ] <- "Otras entidades financieras en la Union Europea"
      data$SECTOR[!is.na(match(data$TIPO_ENTE, c(10,11,50,51,53))) & !is.na(match(data$RESI, c("AR", "BO", "BR", "CL", "CO", "CU", "DO", "EC", "SV", "GT", "HN", "NI", "PA", "PY", "PE", "PR", "UY", "VE"))) ] <- "Otras entidades financieras en America Latina"
      data$SECTOR[!is.na(match(data$TIPO_ENTE, c(10,11,50,51,53))) & !is.na(match(data$RESI, c("BS", "CA", "CH", "ND", "KY", "AN", "AU", "HK", "SG"))) ] <- "Otras entidades financieras extranjeras"
      
      data$SECTOR[data$TIPO_ENTE==47 & data$RESI=="US" ] <- "Gobiernos Federal y Local de los E.U.A."
      data$SECTOR[data$TIPO_ENTE==47 & !is.na(match(data$RESI, c("AT", "BE", "BG", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE", "GB", "VG"))) ] <- "Gobiernos Federal y Local de la Union Europea"
      data$SECTOR[data$TIPO_ENTE==47 & !is.na(match(data$RESI, c("AR", "BO", "BR", "CL", "CO", "CU", "DO", "EC", "SV", "GT", "HN", "NI", "PA", "PY", "PE", "PR", "UY", "VE"))) ] <- "Gobiernos Federal y Local en America Latina"
      data$SECTOR[data$TIPO_ENTE==47 & !is.na(match(data$RESI, c("BS", "CA", "CH", "ND", "KY", "AN", "AU", "HK", "SG"))) ] <- "Gobiernos Federal y Local en otras regiones y paises"
      
      data$SECTOR[data$TIPO_ENTE==28 & data$RESI=="US" ] <- "Empresas privadas no financieras en E.U.A."
      data$SECTOR[data$TIPO_ENTE==28 & !is.na(match(data$RESI, c("AT", "BE", "BG", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE", "GB", "VG"))) ] <- "Empresas privadas no financieras en la Union Europea"
      data$SECTOR[data$TIPO_ENTE==28 & !is.na(match(data$RESI, c("AR", "BO", "BR", "CL", "CO", "CU", "DO", "EC", "SV", "GT", "HN", "NI", "PA", "PY", "PE", "PR", "UY", "VE"))) ] <- "Empresas privadas no financieras en America Latina"
      data$SECTOR[data$TIPO_ENTE==28 & !is.na(match(data$RESI, c("BS", "CA", "CH", "ND", "KY", "AN", "AU", "HK", "SG"))) ] <- "Empresas privadas no financieras de otras regiones y paises"
      
      data$SECTOR[data$TIPO_ENTE==27] <- "Personas fisicas residentes en el extranjero"
      
      # ===== WRITE =====
      # Escribe el cuadro (.xlsx) en el directorio de trabajo
      write.dbf(data, paste("swaps3_contra_", format(Sys.Date()[1], "%d%m%Y"), ".dbf", sep=""))
      
      if(nrow(raros)!=0){
            resultado <- list(cuadro=data, raros=raros)
            invisible(resultado)
      } else{
            invisible(data)
      }
}
