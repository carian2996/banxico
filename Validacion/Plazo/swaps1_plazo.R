# Ian Castillo Rosales (BANXICO\T41348)
# Gerencia de Informaci?n del Sistema Financiero
# Subgerencia de Informaci?n de Moneda Extranjera y Derivados
# 
# Validación de información para operaciones con swaps (plazo)
# 040614 - 010714

swaps1_plazo <- function(ruta){
      
      # ENTRADA
      # ruta = Ruta donde se encuentran los datos para los calculos
            # swaps1.dbf
            # udi2013.dbf
            # fix.dbf
      
      # SALIDA
      # swaps1_plazo_[fecha].dbf - Archivo tipo .dbf con los resultados
      
      # ===== Librerias y directorios =====
      setwd(paste(ruta, "SWAPS/", sep="")) # ¿Dónde están mis datos?
      library(foreign) # Libreria necesaria para cargar los datos
      options(scipen=999, digits=8) # Quita la notación exp y trunca a 4 decimales
      
      # ===== Carga de datos =====
      data <- read.dbf("swaps1.dbf", as.is=T)
      gc()
      udis <- read.dbf("udi2013.dbf", as.is=T)
      fix <- read.dbf("tcfix.dbf", as.is=T)
      
      # ===== Código =====
      # apply(data, 2, function(x) any(is.na(x)))
      
      # Colocamos los registros que contengan casos incompletos (con NA's)
      raros <- data[!complete.cases(data[, 1:6]), ]
      # Escogemos los registros que contengan casos completos (sin NA's)
      data <- data[complete.cases(data[, 1:6]), ]
      # Arrojamos un mensaje en caso de que existan casos incompletos
      if(nrow(raros)!=0){
            message("Existen registros incompletos para Swaps1 Plazo")
      }
      
      # ===== Tipo de Institucion =====
      data$TIPO_INST <- NA
      data$TIPO_INST[substr(data$INSTI, 1, 3) == "013"] <- "CB"
      data$TIPO_INST[substr(data$INSTI, 1, 3) != "013"] <- "BM_BD"
      
      # ===== UDIS y FIX =====
      # Buscamos UDIS y unir con datos
      data$UDIS <- udis$CIERRE[match(as.Date(data$FE_CON_OPE), as.Date(udis$FE_PUBLI))] 
      # Buscamos FIX y unir con datos
      data$FIX <- fix$CIERRE[match(as.Date(data$FE_CON_OPE), as.Date(fix$FE_PUBLI))] 
      
      # ===== IMPORTE =====
      data$IMPORTE <- NA
      
      data$IMPORTE[data$MDA_IMP=="MXP"] <- data$C_IMP_BASE[data$MDA_IMP=="MXP"]/2000
      data$IMPORTE[data$MDA_IMP=="UDI"] <- data$C_IMP_BASE[data$MDA_IMP=="UDI"]*data$UDIS[data$MDA_IMP=="UDI"]/2000
      data$IMPORTE[data$MDA_IMP=="USD"] <- data$C_IMP_BASE[data$MDA_IMP=="USD"]*data$FIX[data$MDA_IMP=="USD"]/2000
      
      # ===== Plazo =====
      data$PLAZO <- NA # Crear columna de plazo
      # Realiza la diferencia entre fechas
      data$PLAZO <- as.numeric(as.Date(data$FE_LIQ_ORI) - as.Date(data$FE_CON_OPE))
      
      if(any(data$PLAZO < 0)){
            data$PLAZO[data$PLAZO < 0] <- 0
            message("Existen plazos negativos en SWAPS1 Plazo")
      }
      
      # ===== BANDA =====
      # Matriz de bandas
      bandas <- matrix(0, nrow=14, ncol=2)
      bandas[, 1] <- c(0, 8, 32, 93, 185, 367, 732, 1097, 1462, 1828, 2558, 
                       3654, 5480, 7306)
      bandas[, 2] <- c("1 a 7", "8 a 31", "32 a 92", "93 a 184", "185 a 366", 
                       "367 a 731", "732 a 1096", "1097 a 1461", "1462 a 1827", 
                       "1828 a 2557", "2558 a 3653", "3654 a 5479", 
                       "5480 a 7305", "mas de 7306")
      
      # Crear y renombrar columna de bandas
      data$BANDA <- NA
      # Encuentra el intervalo y pone la banda
      data$BANDA <- bandas[, 2][findInterval(data$PLAZO, as.numeric(bandas[, 1]))]
      
      # ===== WRITE =====
      # Escribe el cuadro (.xlsx) en el directorio de trabajo
      write.dbf(data, paste("swaps1_plazo_", format(Sys.Date()[1], "%d%m%Y"), ".dbf", sep=""))
      
      if(nrow(raros)!=0){
            resultado <- list(cuadro=data, raros=raros)
            invisible(resultado)
      } else{
            invisible(data)
      }
}