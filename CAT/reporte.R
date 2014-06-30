# Ian Castillo Rosales (BANXICO\T41348)
# Gerencia de Información del Sistema Financiero
# Subgerencia de Información de Moneda Extranjera y Derivados
# 
# Matriz de condiciones para calculo de CAT
# 180214 - 110414

reporte <- function(archivo, datos=F, toge=F){
  
  # Genera un reporte con la matriz de condiciones para el calculo del CAT e información consistente del CCT-H
  # 
  # ENTRADA
  # archivo = (Caracter) Nombre del archivo .xlsx con los datos a importar.
  # datos = (Logico) FALSE Carga los datos con la función "readWorksheetFromFile". TRUE Carga los datos de un archivo .RData.
  # toge = (Logico) FALSE Genera un archivo por cada institucion. TRUE Genera un archivo con todas las instituciones.
  #
  # SALIDA
  # Serie de archivos .xslx por cada institución (con página web) con sus respectivos productos (vigentes) y la 
  # matriz de condiciones para el calculo del CAT e información consistente del CCT-H. Otra opcion es la generacion
  # de un archivo llamado 'Reporte' con todas las instituciones (con página web) con sus respectivos productos 
  # (vigentes) y la matriz de condiciones para el calculo del CAT e información consistente del CCT-H.
  
  # ============================== Opciones y librerias ===================================
    # Carga la libreria 'XLConnect'
    library(XLConnect)
    # Determinamos la ruta donde se guardaran nuestros archivos
    setwd("/Users/iancastillo/Google Drive/Banxico/Workspace(Banxico)/Archivos por institucion")
    # 'scipen' nos permite penalizar la impresion de un numero con notacion exponencial y se asegura que se imprima con notacion regular
    options(scipen=100)
    # 'econding' nos permite cambiar el uso de caracteres. ISO 8859-1 representa los usos para el alfabeto latino
    options(encoding="ISO 8859-1")
    # 'java.parameters' nos permite aumentar la memoria para poder cargar los datos adecuadamente
    options(java.parameters = "-Xmx2048m")
    options(verbose=F)

  # CONDICIONAL. Si el usuario tiene datos precargados o no.
  if(datos==T){
    # Se carga los datos previamente guardados. 'file.choose' nos permite abrir el explorador y buscar nuestro archivo
    print("Seleccione los datos precargados")
    load(file.choose()) 
  } else{
    # Se leen los datos del 'archivo' para cargar los datos a R
    gc(reset=T)
    cat.inst.cct.h <- readWorksheetFromFile(paste("/Users/iancastillo/Google Drive/Banxico/Workspace(Banxico)/", archivo, sep=""),"cat_inst_cct_h", header=T, startRow=8, startCol=3)
    gc(reset=T)
    cat.prod.h <- readWorksheetFromFile(paste("/Users/iancastillo/Google Drive/Banxico/Workspace(Banxico)/", archivo, sep=""),"cat_prod_h", header=T, startRow=6, startCol=3)
    gc(reset=T)
    cat.noweb <- readWorksheetFromFile(paste("/Users/iancastillo/Google Drive/Banxico/Workspace(Banxico)/", archivo, sep=""),"Cat_noweb", header=T)
    gc(reset=T)
    especificaciones <- readWorksheetFromFile(paste("/Users/iancastillo/Google Drive/Banxico/Workspace(Banxico)/", archivo, sep=""),"Especificaciones", header=T)
    gc(reset=T)
    cat.etiquetas <- readWorksheetFromFile(paste("/Users/iancastillo/Google Drive/Banxico/Workspace(Banxico)/", archivo, sep=""),"cat_etiquetas", startRow = 3, startCol = 2, header=T)
    gc(reset=T)
    aforos <- readWorksheetFromFile(paste("/Users/iancastillo/Google Drive/Banxico/Workspace(Banxico)/", archivo, sep=""),"Aforos", header=T)
    gc(reset=T)
    hechos.genera <- readWorksheetFromFile(paste("/Users/iancastillo/Google Drive/Banxico/Workspace(Banxico)/", archivo, sep=""),"Hechos generadores", header=T)
    gc(reset=T)
    cat.avaluo <- readWorksheetFromFile(paste("/Users/iancastillo/Google Drive/Banxico/Workspace(Banxico)/", archivo, sep=""),"Cat_avaluo", startRow = 2, startCol = 2, header=T)
    gc(reset=T)
    pagos <- readWorksheetFromFile(paste("/Users/iancastillo/Google Drive/Banxico/Workspace(Banxico)/", archivo, sep=""),"pagos", header=T)
    gc(reset=T)
    tasas <- readWorksheetFromFile(paste("/Users/iancastillo/Google Drive/Banxico/Workspace(Banxico)/", archivo, sep=""),"tasas", header=T)
    gc(reset=T)
    seguros <- readWorksheetFromFile(paste("/Users/iancastillo/Google Drive/Banxico/Workspace(Banxico)/", archivo, sep=""),"seguros", header=T)
    gc(reset=T)
    segurosg <- readWorksheetFromFile(paste("/Users/iancastillo/Google Drive/Banxico/Workspace(Banxico)/", archivo, sep=""),"Cat_seguro_gracia", header=T)
    gc(reset=T)
    bonis <- readWorksheetFromFile(paste("/Users/iancastillo/Google Drive/Banxico/Workspace(Banxico)/", archivo, sep=""),"bonificaciones_especiales", header=T)
    gc(reset=T)
    cat <- readWorksheetFromFile(paste("/Users/iancastillo/Google Drive/Banxico/Workspace(Banxico)/", archivo, sep=""), "CAT", header=T)
    gc(reset=T)
  }
  
  # CONDICIONAL. Si el usuario quiere un reporte con todas instituciones se genera un workbook antes del loop.
  if (toge==T){
    # 'loadWolrkbook' crea un archivo .xlsx con el nombre 'Reporte_Fecha_De_Ejecucion'. Si el workbook ya existe, lo carga en el ambiente de trabajo.
    # 'formart' nos permite adecuar un formato especifico para nuestra fecha
    # 'Sys.Date' nos proporciona la fecha a la hora de su ejecucion
    wb <- loadWorkbook(paste("Reporte_", format(Sys.Date()[1], "%Y_%m_%d"), ".xlsx", sep=""), create = T)
  }
  
  # LOOP. Se realiza una iteracion por cada institución donde la variable 'web' sea VERDADERO.
  for(i in which(cat.inst.cct.h$web)){
    # ====================================== Clave :) ======================================
    {
      # Se genera el 'nombre' de la institucion.
      clave_inst <- cat.inst.cct.h$clave[i]
      # CONDICIONAL. Si la opcion de reporte conjunto es falsa, se genera un workbook por cada iteracion. Si el workbook
      # ya existe, lo carga en el ambiente de trabajo.
      if(toge==F){
        wb <- loadWorkbook(paste(clave_inst,"_", format(Sys.Date()[1], "%Y_%m_%d"), ".xlsx", sep=""), create = T)
      }
      # Se crea una hoja en el workbook 'wb' con el nombre de la institucion
      createSheet(wb, name=clave_inst)
    }
    # ============================== Productos vigentes :) =================================
    {
      # Genera la variable 'prods.vigentes' con todos las claves de los productos vigentes, con las siguientes condiciones:
        # 1. Los productos deben corresponder a la institucion 'nombre'
        # 2. Los productos deben cumplir variable 'fecha_can' igual a 1900-01-02
      # La funcion 'sort' ordena la clave de los productos de menor a mayor
      prods.vigentes <- sort(cat.prod.h[cat.prod.h$inst==clave_inst & as.Date(cat.prod.h$fecha_can)=="1900-01-02",1])
      # Creamos un copia de los datos para trabajar con ellos
      cat.noweb.inst <- cat.noweb
      cat.noweb.inst <- cat.noweb[cat.noweb$inst_id==clave_inst,] # Dejamos los registros que coincidan con la institución correspondiente
      prods.vigentes <- prods.vigentes[is.na(match(prods.vigentes, cat.noweb.inst$prod_id))] # Dejamos los productos que estén vigentes y que correspondan a la institución
      # Se genera la variable 'names.prods' con los nombres de los productos vigentes. La funcion 'match' permite 
      # saber en que posicion están los productos vigentes sobre la variable que contiene el nombre de los productos
      names.prods <- cat.prod.h$clave_inst[match(prods.vigentes, cat.prod.h$prod_id)]
    }
    # ============================== Especificaciones :) ===================================
    {
      esp.inst <- especificaciones # Copia de datos
      esp.inst <- esp.inst[!is.na(match(esp.inst$INST,clave_inst)),] # Se desechan registros que no correspondan a institución
      esp.inst <- esp.inst[!is.na(match(esp.inst$PROD_ID, prods.vigentes)),] # Se desechan registros que no correspondan a productos vigentes
      
      claves <- as.character(c(51030, 51031, 51040, 51140, 51200)) # Se asignan claves
      names.claves <- cat.etiquetas[match(claves, cat.etiquetas$Clave),4] # Se asignan nombres de claves
      
      matriz.esp <- matrix(NA, nrow=length(claves), ncol=length(prods.vigentes)) # Se crea la matriz de especificaciones
      colnames(matriz.esp) <- prods.vigentes # Se da nombre a las columnas de la matriz (los nombres son los productos vigentes)
      
      matriz.esp[1:5,is.na(match(prods.vigentes, levels(as.factor(esp.inst$PROD_ID))))] <- NA # Se asigna NA a los productos que se reportaron vigentes pero no se encuentran en los registros
      # Se llena la matriz de especificaciones
      for(j in 1:5){
        matriz.esp[j,!is.na(match(prods.vigentes, levels(as.factor(esp.inst$PROD_ID))))] <- esp.inst$DATO[match(prods.vigentes[!is.na(match(prods.vigentes, levels(as.factor(esp.inst$PROD_ID))))], esp.inst$PROD_ID)+j-1]
      }
    }
    # ============================== Aforos :) =============================================
    {
      # Creamos la matriz de aforos en primer lugar, ya que pueden existir diferentes aforos en una sola institucion.
      # Se genera un data.frame 'aforos.inst' donde solo contenga la informacion de la institucion correspondiente
      aforos.inst <- aforos
      aforos.inst <- aforos.inst[!is.na(match(aforos$INST,clave_inst)),]
      aforos.inst <- data.frame(aforos.inst[,-9], aforos.inst$NO_RANGOS)
      # La variable 'n.rangos.aforos' nos permite determinar los rangos que tienen los productos con respecto sus aforos
      # Ejemplo: Si n.rangos.aforos = 1, 2, 3. Quiere decir que existen productos con un aforo, productos con dos aforos
      # y también existen productos con tres aforos diferentes.
      n.rangos.aforos <- levels(factor(aforos.inst$aforos.inst.NO_RANGOS))
      # Creamos la variable 'aforos.eti' que contiene las etiquetas de los aforos, se toman del catalogos de etiquetas
      aforos.eti <- cat.etiquetas[cat.etiquetas$Concepto=="Aforo",5]
      if(length(n.rangos.aforos)!=0){
        n.aforos <- sort(rep(paste("Aforo", 1:max(as.numeric(n.rangos.aforos)), sep=" "), length(aforos.eti)))
      } else{
        n.aforos <- c()
      }
      
      # CONDICIONAL. Si el numero de rangos es un numero mayor que 0 (es decir, existen por lo menos un producto con al
      # menos un aforo) se realizan las acciones
      if(length(n.rangos.aforos)!=0){
        # Crea la matriz de aforos que se anexara al reporte. Tienen una longitud por fila dependiendo del numero de
        # etiquetas que necesite y una longitud por columna dependiendo de cuandos productos vigentes totales se tenga.
        # Nota: Recordar que productos vigentes totales representa a los productos vigentes con sus repeticiones dependiendo
        # de su rango.
        matriz.aforos <- matrix(NA, nrow=(length(aforos.eti)*max(as.numeric(n.rangos.aforos))), ncol=length(prods.vigentes))
        rownames(matriz.aforos) <- n.aforos
        
        # Asignamos a las columnas de la matriz un NA (not available) si existen productos vigentes pero no están
        # en el data.frame de aforos  
        matriz.aforos[1:nrow(matriz.aforos), is.na(match(prods.vigentes, levels(as.factor(aforos.inst$PROD_ID))))] <- NA
        
        for(j in 1:length(aforos.eti)){
          conta <- c()
          for(k in as.numeric(n.rangos.aforos)){
            conta <- j+3*(1:max(as.numeric(n.rangos.aforos))-1)
            matriz.aforos[conta[1:k], !is.na(match(prods.vigentes, aforos.inst$PROD_ID[aforos.inst$aforos.inst.NO_RANGOS==k]))] <- aforos.inst[aforos.inst$aforos.inst.NO_RANGOS==k & !is.na(match(aforos.inst$PROD_ID, prods.vigentes)), j+6]
          }
        }
      } else{
        matriz.aforos <- matrix(NA, nrow=length(aforos.eti), ncol=length(prods.vigentes))
      }
    }
    # ============================== Hechos generadores de comisiones :) ===================
    {
      comisiones.clave <- as.character(c(53010, 53012, 53013, 52025))
      
      hechos.inst <- hechos.genera
      hechos.inst <- hechos.inst[!is.na(match(hechos.genera$INST, clave_inst)),]
      hechos.inst <- hechos.inst[!is.na(match(hechos.inst$PROD_ID, prods.vigentes)),]
      hechos.inst <- hechos.inst[!is.na(match(hechos.inst$CVE_COMIS, comisiones.clave)),]
      # hechos.inst <- rbind(hechos.inst[!is.na(match(hechos.inst$CVE_COMIS, comisiones.clave[1])), ], 
      #                     hechos.inst[!is.na(match(hechos.inst$CVE_COMIS, comisiones.clave[2])), ], 
      #                     hechos.inst[!is.na(match(hechos.inst$CVE_COMIS, comisiones.clave[3])), ], 
      #                     hechos.inst[!is.na(match(hechos.inst$CVE_COMIS, comisiones.clave[4])), ])     
      
      names.comisiones <- cat.etiquetas$Concepto[match(comisiones.clave, cat.etiquetas$Clave)]
      comisiones.etiq <- cat.etiquetas$Etiqueta[!is.na(match(cat.etiquetas$Clave, comisiones.clave))][1:length(comisiones.clave)]
      
        ## ======= Matriz de Moneda Base ========
        {
          matriz.mda.base <- matrix(NA, nrow=length(comisiones.clave), ncol=length(prods.vigentes))
          colnames(matriz.mda.base) <- prods.vigentes
          
          for(k in 1:length(comisiones.clave)){
            k=1
            matriz.mda.base[k, is.na(match(prods.vigentes, hechos.inst$PROD_ID[hechos.inst$CVE_COMIS==comisiones.clave[k]]))] <- NA
          }
          for(k in 1:length(comisiones.clave)){
            matriz.mda.base[k, !is.na(match(prods.vigentes, hechos.inst$PROD_ID[hechos.inst$CVE_COMIS==comisiones.clave[k]]))] <- na.omit(hechos.inst$MDA_BASE[hechos.inst$CVE_COMIS==comisiones.clave[k]][match(prods.vigentes, hechos.inst$PROD_ID[hechos.inst$CVE_COMIS==comisiones.clave[k]])])
          }
        }
        ## ======= Matriz de Importe FI ========
        {
          matriz.imp.fi <- matrix(NA, nrow=length(comisiones.clave), ncol=length(prods.vigentes))
          colnames(matriz.imp.fi) <- prods.vigentes
          for(k in 1:length(comisiones.clave)){
            matriz.imp.fi[k, is.na(match(prods.vigentes, hechos.inst$PROD_ID[hechos.inst$CVE_COMIS==comisiones.clave[k]]))] <- NA
          }
          for(k in 1:length(comisiones.clave)){
            matriz.imp.fi[k, !is.na(match(prods.vigentes, hechos.inst$PROD_ID[hechos.inst$CVE_COMIS==comisiones.clave[k]]))] <- na.omit(hechos.inst$IMPORTE_FI[hechos.inst$CVE_COMIS==comisiones.clave[k]][match(prods.vigentes, hechos.inst$PROD_ID[hechos.inst$CVE_COMIS==comisiones.clave[k]])])
          }
        }
        ## ======= Matriz de Importe VA ========
        {
          matriz.imp.va <- matrix(NA, nrow=length(comisiones.clave), ncol=length(prods.vigentes))
          colnames(matriz.imp.va) <- prods.vigentes
          for(k in 1:length(comisiones.clave)){
            matriz.imp.va[k, is.na(match(prods.vigentes, hechos.inst$PROD_ID[hechos.inst$CVE_COMIS==comisiones.clave[k]]))] <- NA
          }
          for(k in 1:length(comisiones.clave)){
            matriz.imp.va[k, !is.na(match(prods.vigentes, hechos.inst$PROD_ID[hechos.inst$CVE_COMIS==comisiones.clave[k]]))] <- na.omit(hechos.inst$IMPORTE_VA[hechos.inst$CVE_COMIS==comisiones.clave[k]][match(prods.vigentes, hechos.inst$PROD_ID[hechos.inst$CVE_COMIS==comisiones.clave[k]])])
          }
        }
        ## ======= Matriz de Importe MA ========
        {
          matriz.imp.ma <- matrix(NA, nrow=length(comisiones.clave), ncol=length(prods.vigentes))
          colnames(matriz.imp.ma) <- prods.vigentes
          for(k in 1:length(comisiones.clave)){
            matriz.imp.ma[k, is.na(match(prods.vigentes, hechos.inst$PROD_ID[hechos.inst$CVE_COMIS==comisiones.clave[k]]))] <- NA
          }
          for(k in 1:length(comisiones.clave)){
            matriz.imp.ma[k, !is.na(match(prods.vigentes, hechos.inst$PROD_ID[hechos.inst$CVE_COMIS==comisiones.clave[k]]))] <- na.omit(hechos.inst$IMPORTE_MA[hechos.inst$CVE_COMIS==comisiones.clave[k]][match(prods.vigentes, hechos.inst$PROD_ID[hechos.inst$CVE_COMIS==comisiones.clave[k]])])
          }
        }
    }
    # ============================== Hechos generadores de comisiones 2 :) =================
    {
      hechos2.claves <- c("52000", "52020", "52040")
      hechos2.concep <- cat.etiquetas$Concepto[match(hechos2.claves, cat.etiquetas$Clave)]
      hechos2.eti <- cat.etiquetas$Etiqueta[!is.na(match(cat.etiquetas$Clave, hechos2.claves))]
      
      hechos2.inst <- hechos.genera
      hechos2.inst <- hechos2.inst[!is.na(match(hechos2.inst$INST, clave_inst)),]
      hechos2.inst <- hechos2.inst[!is.na(match(hechos2.inst$PROD_ID, prods.vigentes)),]
      hechos2.inst <- hechos2.inst[!is.na(match(hechos2.inst$CVE_COMIS, hechos2.claves)),]
      hechos2.inst <- rbind(hechos2.inst[!is.na(match(hechos2.inst$CVE_COMIS, hechos2.claves[1])), ], 
                           hechos2.inst[!is.na(match(hechos2.inst$CVE_COMIS, hechos2.claves[2])), ], 
                           hechos2.inst[!is.na(match(hechos2.inst$CVE_COMIS, hechos2.claves[3])), ])  
      
        ## ======= Matriz de Moneda Base ========
        {
          matriz.mda.base2 <- matrix(NA, nrow=length(hechos2.claves), ncol=length(prods.vigentes))
          colnames(matriz.mda.base2) <- prods.vigentes
          
          for(k in 1:length(hechos2.claves)){
            matriz.mda.base2[k, is.na(match(prods.vigentes, hechos2.inst$PROD_ID[hechos2.inst$CVE_COMIS==hechos2.claves[k]]))] <- NA
          }
          for(k in 1:length(hechos2.claves)){
            matriz.mda.base2[k, !is.na(match(prods.vigentes, hechos2.inst$PROD_ID[hechos2.inst$CVE_COMIS==hechos2.claves[k]]))] <- na.omit(hechos2.inst$MDA_BASE[hechos2.inst$CVE_COMIS==hechos2.claves[k]][match(prods.vigentes, hechos2.inst$PROD_ID[hechos2.inst$CVE_COMIS==hechos2.claves[k]])])
          }
        }
        ## ======= Matriz de Importe FI ========
        {
          matriz.imp.fi2 <- matrix(NA, nrow=length(hechos2.claves), ncol=length(prods.vigentes))
          colnames(matriz.imp.fi2) <- prods.vigentes
          for(k in 1:length(hechos2.claves)){
            matriz.imp.fi2[k, is.na(match(prods.vigentes, hechos2.inst$PROD_ID[hechos2.inst$CVE_COMIS==hechos2.claves[k]]))] <- NA
          }
          for(k in 1:length(hechos2.claves)){
            matriz.imp.fi2[k, !is.na(match(prods.vigentes, hechos2.inst$PROD_ID[hechos2.inst$CVE_COMIS==hechos2.claves[k]]))] <- na.omit(hechos2.inst$IMPORTE_FI[hechos2.inst$CVE_COMIS==hechos2.claves[k]][match(prods.vigentes, hechos2.inst$PROD_ID[hechos2.inst$CVE_COMIS==hechos2.claves[k]])])
          }
        }
        ## ======= Matriz de Importe VA ========
        {
          matriz.imp.va2 <- matrix(NA, nrow=length(hechos2.claves), ncol=length(prods.vigentes))
          colnames(matriz.imp.va2) <- prods.vigentes
          
          for(k in 1:length(hechos2.claves)){
            matriz.imp.va2[k, is.na(match(prods.vigentes, hechos2.inst$PROD_ID[hechos2.inst$CVE_COMIS==hechos2.claves[1]]))] <- NA
          }
          for(k in 1:length(hechos2.claves)){
            matriz.imp.va2[k, !is.na(match(prods.vigentes, hechos2.inst$PROD_ID[hechos2.inst$CVE_COMIS==hechos2.claves[k]]))] <- na.omit(hechos2.inst$IMPORTE_VA[hechos2.inst$CVE_COMIS==hechos2.claves[k]][match(prods.vigentes, hechos2.inst$PROD_ID[hechos2.inst$CVE_COMIS==hechos2.claves[k]])])
          }
        }
        ## ======= Matriz de Importe MA ========
        {
          matriz.imp.ma2 <- matrix(NA, nrow=length(hechos2.claves), ncol=length(prods.vigentes))
          colnames(matriz.imp.ma2) <- prods.vigentes
          for(k in 1:length(hechos2.claves)){
            matriz.imp.ma2[k, is.na(match(prods.vigentes, hechos2.inst$PROD_ID[hechos2.inst$CVE_COMIS==hechos2.claves[k]]))] <- NA
          }
          for(k in 1:length(hechos2.claves)){
            matriz.imp.ma2[k, !is.na(match(prods.vigentes, hechos2.inst$PROD_ID[hechos2.inst$CVE_COMIS==hechos2.claves[k]]))] <- na.omit(hechos2.inst$IMPORTE_MA[hechos2.inst$CVE_COMIS==hechos2.claves[k]][match(prods.vigentes, hechos2.inst$PROD_ID[hechos2.inst$CVE_COMIS==hechos2.claves[k]])])
          }
        }
    }
    # ============================== Avaluos :) ============================================
    {
      # Agregamos las etiquetas y creamos un data.frame para los avaluos de la institucion
      avaluos.eti <- c("ur_id", cat.etiquetas$Etiqueta[cat.etiquetas$Origen=="Cat_avaluo"])
      cat.avaluo.inst <- cat.avaluo
      cat.avaluo.inst <- cat.avaluo.inst[!is.na(match(cat.avaluo.inst$inst,clave_inst)), ]
      ur_id <- levels(factor(cat.avaluo.inst$ur_id))
      n.rangos.avaluos <- levels(factor(cat.avaluo.inst$ur_id))
      
      if(length(n.rangos.avaluos)!=0){
        n.avaluos <- sort(rep(paste("Avaluo", 1:max(as.numeric(n.rangos.avaluos)), sep=" "), length(avaluos.eti)))
        # Crea la matriz de avaluos que se anexara al reporte. Tienen una longitud por fila dependiendo del numero de
        # etiquetas que necesite y una longitud por columna dependiendo de cuandos productos vigentes totales se tenga.
        matriz.avaluo <- matrix(NA, nrow=(length(avaluos.eti)*max(as.numeric(n.rangos.avaluos))), ncol=length(prods.vigentes))
        rownames(matriz.avaluo) <- n.avaluos
        
        # Asignamos a las columnas de la matriz un NA si existen productos vigentes pero no están en el data.frame de aforos  
        matriz.avaluo[1:(length(avaluos.eti)*max(as.numeric(n.rangos.avaluos))), is.na(match(prods.vigentes, levels(as.factor(cat.avaluo.inst$prod_id))))] <- NA
        
        for(j in 1:length(avaluos.eti)){
          for(k in as.numeric(n.rangos.avaluos)){matriz.avaluo[(j+length(avaluos.eti)*(k-1)), !is.na(match(prods.vigentes, cat.avaluo.inst$prod_id[cat.avaluo.inst$ur_id==ur_id[k]]))] <- cat.avaluo.inst[cat.avaluo.inst$ur_id==ur_id[k] & !is.na(match(cat.avaluo.inst$prod_id, prods.vigentes)), j+3]
          }
        }
      } else{
        n.rangos.avaluos <- 0
        n.avaluos <- sort(rep(paste("Avaluo", 1:max(as.numeric(n.rangos.avaluos)), sep=" "), length(avaluos.eti)))
        matriz.avaluo <- matrix(NA, nrow=length(avaluos.eti), ncol=length(prods.vigentes))
      }
    }
    # ============================== Rangos de plazos :) ===================================
    {
      pagos.inst <- pagos
      pagos.inst <- pagos.inst[!is.na(match(pagos.inst$inst,clave_inst)),]
      pagos.inst <- pagos.inst[!is.na(match(pagos.inst$prod_id, prods.vigentes)),] 
      
      plazos.eti <- cat.etiquetas$Etiqueta[59:60]
      plazos.concep <- cat.etiquetas$Concepto[59:60]
      
      matriz.plazos <- matrix(NA, nrow=length(plazos.eti), ncol=length(prods.vigentes))
      colnames(matriz.plazos) <- prods.vigentes
      
      matriz.plazos[1:length(plazos.eti), is.na(match(prods.vigentes, levels(as.factor(pagos.inst$prod_id))))] <- NA     
      matriz.plazos[1, !is.na(match(prods.vigentes, levels(as.factor(pagos.inst$prod_id))))] <- pagos.inst$li_pzocred[match(prods.vigentes[!is.na(match(prods.vigentes, levels(as.factor(pagos.inst$prod_id))))], pagos.inst$prod_id)]
      matriz.plazos[2, !is.na(match(prods.vigentes, levels(as.factor(pagos.inst$prod_id))))] <- pagos.inst$ls_pzocred[match(prods.vigentes[!is.na(match(prods.vigentes, levels(as.factor(pagos.inst$prod_id))))], pagos.inst$prod_id)]
    }
    # ============================== Pagos al millar fijos :) ==============================
    {
      pagos.eti <- cat.etiquetas$Etiqueta[61:88]
      pagos.concep <- cat.etiquetas$Concepto[61:88]
      
      matriz.pagosxmil <- matrix(NA, nrow=length(pagos.eti), ncol=length(prods.vigentes))
      colnames(matriz.pagosxmil) <- prods.vigentes
      rownames(matriz.pagosxmil) <- pagos.eti
      
      matriz.pagosxmil[1:length(pagos.eti), is.na(match(prods.vigentes, levels(as.factor(pagos.inst$prod_id))))] <- NA     
      for(k in 1:length(pagos.eti)){
        matriz.pagosxmil[k, !is.na(match(prods.vigentes, levels(as.factor(pagos.inst$prod_id))))] <- pagos.inst[match(prods.vigentes[!is.na(match(prods.vigentes, levels(as.factor(pagos.inst$prod_id))))], pagos.inst$prod_id), which(colnames(pagos.inst)==pagos.eti[k])]
      } 
    }
    # ============================== Pagos al millar no uniformes :)) ======================
    {
      n.plazos <- c()
      for(k in 1:length(prods.vigentes)){
        n.plazos[k] <- length(which(matriz.pagosxmil[,k]!=0))
      }
      
      plazos.millar <- na.omit(rownames(matriz.pagosxmil)[matriz.pagosxmil[, which(n.plazos==max(n.plazos))]!=0])
      
      pagosnounif.eti <- cat.etiquetas$Etiqueta[89:128]
      pagosnounif.concep <- cat.etiquetas$Concepto[89:128]
      
      if(max(n.plazos)!=0){
        matriz.pagosnounif <- matrix(NA, nrow=(length(pagosnounif.eti)*max(n.plazos)), ncol=length(prods.vigentes))
        colnames(matriz.pagosnounif) <- prods.vigentes
        filas.pagosnounif <- c()
        for(k in 1:max(n.plazos)) filas.pagosnounif <- c(filas.pagosnounif, rep(plazos.millar[k], length(pagosnounif.eti)))
        rownames(matriz.pagosnounif) <- filas.pagosnounif
        
        matriz.pagosnounif[1:(length(pagosnounif.eti)*max(n.plazos)), is.na(match(prods.vigentes, levels(as.factor(pagos.inst$prod_id))))] <- NA
        for(k in 1:length(plazos.millar)){
          matriz.pagosnounif[rownames(matriz.pagosnounif)==plazos.millar[k], which(matriz.pagosxmil[plazos.millar[k], ]!=0)]<- t(pagos.inst[match(prods.vigentes[which(matriz.pagosxmil[plazos.millar[k], ]!=0)], pagos.inst$prod_id),40:79])
        }
      } else{
        matriz.pagosnounif <- matrix(NA, nrow=length(pagosnounif.eti), ncol=length(prods.vigentes))
      }
    }
    # ============================== Tasas de interés :)) ==================================
    {
      tasas.eti <- cat.etiquetas$Etiqueta[134:145]
      tasas.eti2 <- c()
      for(k in 1:4) tasas.eti2 <- c(tasas.eti2, tasas.eti[k], tasas.eti[(3+2*k):(4+2*k)])
      
      tasas.concep <- cat.etiquetas$Concepto[134:145]
      tasas.concep2 <- c()
      for(k in 1:4) tasas.concep2 <- c(tasas.concep2, tasas.concep[k], tasas.concep[(3+2*k):(4+2*k)])
      
      tasas.inst <- tasas
      tasas.inst <- tasas.inst[!is.na(match(tasas.inst$INST, clave_inst)), ]
      tasas.inst <- tasas.inst[!is.na(match(tasas.inst$PROD_ID, prods.vigentes)), ]
      
      tasasxprods <- c()
      for(k in 1:length(levels(factor(tasas.inst$LS_PZOCRED)))){
        tasasxprods <- c(tasasxprods, levels(factor(tasas.inst$PROD_ID[tasas.inst$LS_PZOCRED==levels(factor(tasas.inst$LS_PZOCRED))[k]])))
      }
      
      if(length(levels(factor(tasas.inst$LS_PZOCRED)))!=0){
        matriz.tasas <- matrix(NA, nrow=(length(levels(factor(tasas.inst$LS_PZOCRED)))*length(tasas.eti2)), ncol=length(prods.vigentes))
        colnames(matriz.tasas) <- prods.vigentes
        rownames(matriz.tasas) <- rep(tasas.eti, length(levels(factor(tasas.inst$LS_PZOCRED))))
        
        niveles <- rep(paste("Plazo", sort(as.numeric(levels(factor(tasas.inst$LS_PZOCRED)))), sep=" "), length(tasas.eti))
        
        matriz.tasas[1:(length(levels(factor(tasas.inst$LS_PZOCRED)))*length(tasas.eti)), is.na(match(prods.vigentes, levels(as.factor(tasas.inst$PROD_ID))))] <- NA
        
        for(j in 1:4){ 
          for(k in 1:length(levels(factor(tasas.inst$LS_PZOCRED)))){
            matriz.tasas[(1+3*(j-1))+(4*(k-1)*3), !is.na(match(prods.vigentes, tasas.inst$PROD_ID[tasas.inst$LS_PZOCRED==levels(factor(tasas.inst$LS_PZOCRED))[k]]))] <- tasas.inst[tasas.inst$LS_PZOCRED==levels(factor(tasas.inst$LS_PZOCRED))[k], (12+3*(j-1))][na.omit(match(prods.vigentes, tasas.inst$PROD_ID[tasas.inst$LS_PZOCRED==levels(factor(tasas.inst$LS_PZOCRED))[k]]))]
          }
        }
        
        for(j in 1:4){ 
          for(k in 1:length(levels(factor(tasas.inst$LS_PZOCRED)))){
            matriz.tasas[(2+3*(j-1))+(4*(k-1)*3), !is.na(match(prods.vigentes, tasas.inst$PROD_ID[tasas.inst$LS_PZOCRED==levels(factor(tasas.inst$LS_PZOCRED))[k]]))] <- tasas.inst[tasas.inst$LS_PZOCRED==levels(factor(tasas.inst$LS_PZOCRED))[k], (10+3*(j-1))][na.omit(match(prods.vigentes, tasas.inst$PROD_ID[tasas.inst$LS_PZOCRED==levels(factor(tasas.inst$LS_PZOCRED))[k]]))]
          }
        }
        
        for(j in 1:4){ 
          for(k in 1:length(levels(factor(tasas.inst$LS_PZOCRED)))){
            matriz.tasas[(3+3*(j-1))+(4*(k-1)*3), !is.na(match(prods.vigentes, tasas.inst$PROD_ID[tasas.inst$LS_PZOCRED==levels(factor(tasas.inst$LS_PZOCRED))[k]]))] <- tasas.inst[tasas.inst$LS_PZOCRED==levels(factor(tasas.inst$LS_PZOCRED))[k], (11+3*(j-1))][na.omit(match(prods.vigentes, tasas.inst$PROD_ID[tasas.inst$LS_PZOCRED==levels(factor(tasas.inst$LS_PZOCRED))[k]]))]
          }
        }
        
      } else{
        matriz.tasas <- matrix(NA, nrow=length(tasas.eti[1:4]), ncol=length(prods.vigentes))
        niveles <- c()
        }
    }
    # ============================== Seguros :) ============================================
    {
      claves.seguros <- c("65515", "66000")
      
      seguros.inst <- seguros
      seguros.inst <- seguros.inst[!is.na(match(seguros.inst$INST,clave_inst)), ]
      seguros.inst <- seguros.inst[!is.na(match(seguros.inst$PROD_ID, prods.vigentes)),]
      seguros.inst <- rbind(seguros.inst[!is.na(match(seguros.inst$CVE_SEG, claves.seguros[1])), ], seguros.inst[!is.na(match(seguros.inst$CVE_SEG, claves.seguros[2])), ])
      
      seguros.eti <- cat.etiquetas[!is.na(match(cat.etiquetas$Clave, claves.seguros)) & cat.etiquetas$tipo=="Tabla", 5]
      seguros.concep <- rev(levels(factor(cat.etiquetas[!is.na(match(cat.etiquetas$Clave, claves.seguros)) & cat.etiquetas$tipo=="Tabla", 4])))
      
      matriz.seguros <- matrix(NA, nrow=length(seguros.eti), ncol=length(prods.vigentes))
      colnames(matriz.seguros) <- prods.vigentes
      rownames(matriz.seguros) <- seguros.eti
      
      matriz.seguros[1:length(seguros.eti), is.na(match(prods.vigentes, levels(as.factor(seguros.inst$PROD_ID))))] <- NA 
      for(k in 1:length(levels(factor(seguros.eti)))){
        matriz.seguros[k, !is.na(match(prods.vigentes, levels(as.factor(seguros.inst$PROD_ID))))] <- seguros.inst[seguros.inst$CVE_SEG==levels(factor(seguros.inst$CVE_SEG))[1] , 9+(k-1)][na.omit(match(prods.vigentes, seguros.inst$PROD_ID[seguros.inst$CVE_SEG==levels(factor(seguros.inst$CVE_SEG))[1]]))]
      }
      for(k in 1:length(levels(factor(seguros.eti)))){
        matriz.seguros[k+3, !is.na(match(prods.vigentes, levels(as.factor(seguros.inst$PROD_ID))))] <- seguros.inst[seguros.inst$CVE_SEG==levels(factor(seguros.inst$CVE_SEG))[2], 9+(k-1)][na.omit(match(prods.vigentes, seguros.inst$PROD_ID[seguros.inst$CVE_SEG==levels(factor(seguros.inst$CVE_SEG))[2]]))]
      }
      
    }
    # ================ Seguros gracia y bonificaciones especiales (vida) :) ================
    {
      claves.segurosg <- c("65515", "66000")
      segurosg.eti.v <- cat.etiquetas$Etiqueta[!is.na(match(cat.etiquetas$Clave, claves.segurosg[1])) & cat.etiquetas$Origen=="Cat_seguro_gracia"]
      segurosg.concep.v <- levels(factor(cat.etiquetas$Concepto[!is.na(match(cat.etiquetas$Clave, claves.segurosg[1])) & cat.etiquetas$Origen=="Cat_seguro_gracia"]))
      bonis.eti.v <- cat.etiquetas$Etiqueta[!is.na(match(cat.etiquetas$Clave, claves.segurosg[1])) & cat.etiquetas$Origen=="bonificaciones_especiales"]
      bonis.concep.v <- levels(factor(cat.etiquetas$Concepto[!is.na(match(cat.etiquetas$Clave, claves.segurosg[1])) & cat.etiquetas$Origen=="bonificaciones_especiales"]))
      
      claves.segurosg2 <- c("65516", "66000")
      segurosg.inst.v <- segurosg
      segurosg.inst.v <- segurosg.inst.v[!is.na(match(segurosg.inst.v$prod_id, prods.vigentes)) & segurosg.inst.v$tipo_seg==claves.segurosg2[1],]
      niveles.segurosg.v <- factor(table(segurosg.inst.v$prod_id[segurosg.inst.v$tipo_seg==claves.segurosg2[1]]))
      
      bonis.inst.v <- bonis
      bonis.inst.v <- bonis.inst.v[bonis.inst.v$inst==clave_inst,]
      bonis.inst.v <- bonis.inst.v[!is.na(match(bonis.inst.v$prod_id, prods.vigentes)),]
      niveles.bonis.v <- factor(table(bonis.inst.v$prod_id))
      
      if(length(niveles.segurosg.v)!=0){
        matriz.segurosg.v <- matrix(NA, nrow=(length(segurosg.eti.v)+length(bonis.eti.v))*as.numeric(levels(niveles.segurosg.v)), ncol=length(prods.vigentes))
        colnames(matriz.segurosg.v) <- prods.vigentes
        rownames(matriz.segurosg.v) <- rep(c(segurosg.eti.v, bonis.eti.v), as.numeric(levels(niveles.segurosg.v)))
        
        matriz.segurosg.v[, is.na(match(prods.vigentes, levels(as.factor(segurosg.inst.v$prod_id))))] <- NA
        matriz.segurosg.v[, is.na(match(prods.vigentes, levels(as.factor(bonis.inst.v$prod_id))))] <- NA
        
        matriz.segurosg.v[1:length(segurosg.eti.v), match(levels(factor(segurosg.inst.v$prod_id))[niveles.segurosg.v==levels(niveles.segurosg.v)[1]], prods.vigentes)] <- t(segurosg.inst.v[match(levels(factor(segurosg.inst.v$prod_id))[niveles.segurosg.v==levels(niveles.segurosg.v)[1]], bonis.inst.v$prod_id), 5:8])
        matriz.segurosg.v[(length(segurosg.eti.v)+1):(length(segurosg.eti.v)+length(bonis.eti.v)), match(levels(factor(segurosg.inst.v$prod_id))[niveles.segurosg.v==levels(niveles.segurosg.v)[1]], prods.vigentes)] <- t(bonis.inst.v[match(levels(factor(segurosg.inst.v$prod_id))[niveles.segurosg.v==levels(niveles.segurosg.v)[1]], bonis.inst.v$prod_id), 5:11])
      } else{
        matriz.segurosg.v <- matrix(NA, nrow=length(segurosg.eti.v)+length(bonis.eti.v), ncol=length(prods.vigentes))
        
        matriz.segurosg.v[, is.na(match(prods.vigentes, levels(as.factor(segurosg.inst.v$prod_id))))] <- NA
        matriz.segurosg.v[, is.na(match(prods.vigentes, levels(as.factor(bonis.inst.v$prod_id))))] <- NA
        
        matriz.segurosg.v[1:length(segurosg.eti.v), match(levels(factor(segurosg.inst.v$prod_id))[niveles.segurosg.v==levels(niveles.segurosg.v)[1]], prods.vigentes)] <- t(segurosg.inst.v[match(levels(factor(segurosg.inst.v$prod_id))[niveles.segurosg.v==levels(niveles.segurosg.v)[1]], bonis.inst.v$prod_id), 5:8])
        matriz.segurosg.v[(length(segurosg.eti.v)+1):(length(segurosg.eti.v)+length(bonis.eti.v)), match(levels(factor(segurosg.inst.v$prod_id))[niveles.segurosg.v==levels(niveles.segurosg.v)[1]], prods.vigentes)] <- t(bonis.inst.v[match(levels(factor(segurosg.inst.v$prod_id))[niveles.segurosg.v==levels(niveles.segurosg.v)[1]], bonis.inst.v$prod_id), 5:11])
      }
    }
    # ================ Seguros gracia y bonificaciones especiales (daños) :) ===============
    {
      segurosg.eti.d <- cat.etiquetas$Etiqueta[!is.na(match(cat.etiquetas$Clave, claves.segurosg[2])) & cat.etiquetas$Origen=="Cat_seguro_gracia"]
      segurosg.concep.d <- levels(factor(cat.etiquetas$Concepto[!is.na(match(cat.etiquetas$Clave, claves.segurosg[2])) & cat.etiquetas$Origen=="Cat_seguro_gracia"]))
      bonis.eti.d <- cat.etiquetas$Etiqueta[!is.na(match(cat.etiquetas$Clave, claves.segurosg[2])) & cat.etiquetas$Origen=="bonificaciones_especiales"]
      bonis.concep.d <- levels(factor(cat.etiquetas$Concepto[!is.na(match(cat.etiquetas$Clave, claves.segurosg[2])) & cat.etiquetas$Origen=="bonificaciones_especiales"]))
      
      segurosg.inst.d <- segurosg
      segurosg.inst.d <- segurosg.inst.d[!is.na(match(segurosg.inst.d$prod_id, prods.vigentes)) & segurosg.inst.d$tipo_seg==claves.segurosg[2],]
      niveles.segurosg.d <- factor(table(segurosg.inst.d$prod_id))
      
      bonis.inst.d <- bonis
      bonis.inst.d <- bonis.inst.d[bonis.inst.v$inst==clave_inst,]
      bonis.inst.d <- bonis.inst.d[!is.na(match(bonis.inst.d$prod_id, prods.vigentes)),]
      
      
      if(length(niveles.segurosg.d)!=0){
        matriz.segurosg.d <- matrix(NA, nrow=(length(segurosg.eti.v)+length(bonis.eti.v))*as.numeric(levels(niveles.segurosg.d)), ncol=length(prods.vigentes))
        colnames(matriz.segurosg.d) <- prods.vigentes
        rownames(matriz.segurosg.d) <- rep(c(segurosg.eti.d, bonis.eti.d), as.numeric(levels(niveles.segurosg.d)))
        
        matriz.segurosg.d[, is.na(match(prods.vigentes, levels(as.factor(segurosg.inst.d$prod_id))))] <- NA
        matriz.segurosg.d[, is.na(match(prods.vigentes, levels(as.factor(bonis.inst.d$prod_id))))] <- NA
       
        for(k in 1:levels(niveles.segurosg.d)){
          matriz.segurosg.d[1:length(segurosg.eti.d)+(k-1)*(length(segurosg.eti.d)+length(bonis.eti.d)), match(levels(factor(segurosg.inst.d$prod_id))[niveles.segurosg.d==levels(niveles.segurosg.d)], prods.vigentes)] <- t(segurosg.inst.d[match(levels(factor(segurosg.inst.d$prod_id))[niveles.segurosg.d==levels(niveles.segurosg.d)], segurosg.inst.d$prod_id)+(k-1), 5:8])
          matriz.segurosg.d[((length(segurosg.eti.d)+1):(length(segurosg.eti.d)+length(bonis.eti.d)))+(k-1)*(length(segurosg.eti.d)+length(bonis.eti.d)), match(levels(factor(segurosg.inst.d$prod_id))[niveles.segurosg.d==levels(niveles.segurosg.d)], prods.vigentes)] <- t(bonis.inst.d[match(levels(factor(segurosg.inst.d$prod_id))[niveles.segurosg.d==levels(niveles.segurosg.d)], bonis.inst.v$prod_id), 5:11])
        } 
      } else{
        matriz.segurosg.d <- matrix(NA, nrow=(length(segurosg.eti.v)+length(bonis.eti.v)), ncol=length(prods.vigentes)) 
      }
    }
    # ====================================== CAT :) ========================================
    {
      cat.inst <- cat
      cat.inst <- cat.inst[!is.na(match(cat.inst$inst,clave_inst)),]
      cat.inst <- cat.inst[!is.na(match(cat.inst$prod_id, prods.vigentes)),]
      
      cat.eti <- cat.etiquetas$Etiqueta[cat.etiquetas$Origen=="CAT"]
      cat.concep <- cat.etiquetas$Concepto[cat.etiquetas$Origen=="CAT"]
      
      matriz.cat.vv <- matrix(NA, nrow=length(cat.eti[2:7]), ncol=length(prods.vigentes))
      rownames(matriz.cat.vv) <- cat.eti[2:7]
      
      matriz.cat.vv[, is.na(match(prods.vigentes, levels(as.factor(cat.inst$prod_id))))] <- NA
      matriz.cat.vv[, !is.na(match(prods.vigentes, levels(as.factor(cat.inst$prod_id))))] <- t(cat.inst[na.omit(match(prods.vigentes, levels(as.factor(cat.inst$prod_id)))), 9:14])
      
      niveles.cat <- levels(factor(cat.inst$subprod_id))
      
      if(length(niveles.cat)!=0){
        matriz.cat.subprods <- matrix(NA, nrow=length(c(cat.eti[1], cat.eti[8:14]))*length(niveles.cat), ncol=length(prods.vigentes))
        rownames(matriz.cat.subprods) <- rep(c(cat.eti[1], cat.eti[8:14]), length(niveles.cat))
        
        matriz.cat.subprods[, is.na(match(prods.vigentes, levels(as.factor(cat.inst$prod_id))))] <- NA
        for(k in 1:length(niveles.cat)){
          matriz.cat.subprods[1+(k-1)*length(c(cat.eti[1], cat.eti[8:14])), !is.na(match(prods.vigentes, levels(as.factor(cat.inst$prod_id)))) & !is.na(match(prods.vigentes, cat.inst$prod_id[cat.inst$subprod_id==niveles.cat[k]]))] <- niveles.cat[k]
        }
        for(k in 1:length(niveles.cat)){
          matriz.cat.subprods[2+(k-1)*length(c(cat.eti[1], cat.eti[8:14])), !is.na(match(prods.vigentes, levels(as.factor(cat.inst$prod_id)))) & !is.na(match(prods.vigentes, cat.inst$prod_id[cat.inst$subprod_id==niveles.cat[k]]))] <- cat.inst$cat_prod_min[!is.na(match(cat.inst$prod_id, prods.vigentes)) & cat.inst$subprod_id==niveles.cat[k]]
        }
        for(k in 1:length(niveles.cat)){
          matriz.cat.subprods[((3:8)+(k-1)*length(c(cat.eti[1], cat.eti[8:14]))), !is.na(match(prods.vigentes, levels(as.factor(cat.inst$prod_id))))  & !is.na(match(prods.vigentes, cat.inst$prod_id[cat.inst$subprod_id==niveles.cat[k]]))] <- t(cat.inst[!is.na(match(cat.inst$prod_id, prods.vigentes)) & cat.inst$subprod_id==niveles.cat[k], 15:20])
        }
      } else{
        matriz.cat.subprods <- matrix(NA, nrow=length(c(cat.eti[1], cat.eti[8:14])), ncol=length(prods.vigentes))
      }
    }
    # ================================== Escritura :) ======================================
    {
      writeWorksheet(wb, data = "Matriz de condiciones mínimas para cálculo de CAT e información consistente del CCT-H", sheet = clave_inst, startRow = 1, startCol = 2, header = F)
      writeWorksheet(wb, data = "Fecha de datos", sheet = clave_inst, startRow = 3, startCol = 2, header = F)
      writeWorksheet(wb, data = "Fecha de ejecución", sheet = clave_inst, startRow = 4, startCol = 2, header = F)
      writeWorksheet(wb, data = "Productos vigentes", sheet = clave_inst, startRow = 5, startCol = 2, header = F)
    
      writeWorksheet(wb, data = as.POSIXct(especificaciones$FECHA_DAT[1]), sheet = clave_inst, startRow = 3, startCol = 3, header = F)
      writeWorksheet(wb, data = Sys.time()[1], sheet = clave_inst, startRow = 4, startCol = 3, header = F)
      writeWorksheet(wb, data = length(prods.vigentes) , sheet = clave_inst, startRow = 5, startCol = 3, header = F)
      
      writeWorksheet(wb, data = "Institución", sheet = clave_inst, startRow = 7, startCol = 2 , header = F)
      writeWorksheet(wb, data = cat.inst.cct.h[i,1:2], sheet = clave_inst, startRow = 7, startCol = 3 , header = F)
      writeWorksheet(wb, data = "Producto", sheet = clave_inst, startRow = 7, startCol = 5 , header = F)
      writeWorksheet(wb, data = "Clave", sheet = clave_inst, startRow = 8, startCol = 3 , header = F)
      writeWorksheet(wb, data = "Concepto", sheet = clave_inst, startRow = 8, startCol = 4 , header = F)
      writeWorksheet(wb, data = "Etiqueta", sheet = clave_inst, startRow = 8, startCol = 5 , header = F)
      writeWorksheet(wb, data = t(prods.vigentes), sheet = clave_inst, startRow = 7, startCol = 6 , header = F)
      writeWorksheet(wb, data = t(cat.prod.h$nombre[match(prods.vigentes, cat.prod.h$prod_id)]), sheet = clave_inst, startRow = 8, startCol = 6 , header = F)
      
      lim.inf <- 9
      linea1 <- c()
      linea2 <- c()
      
      writeWorksheet(wb, data = "Especificaciones", sheet = clave_inst, startRow = lim.inf, startCol = 2 , header = F)
      writeWorksheet(wb, data = claves, sheet = clave_inst, startRow = lim.inf, startCol = 3 , header = F)
      writeWorksheet(wb, data = names.claves, sheet=clave_inst, startRow=lim.inf, startCol=4, header=F)
      writeWorksheet(wb, data = rep("DATO", 5), sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
      writeWorksheet(wb, data = matriz.esp, sheet = clave_inst, startRow = 9, startCol = 6 , header = F)
      lim.inf <- lim.inf+nrow(matriz.esp)
      linea1 <- c(linea1, lim.inf)
      if(length(n.rangos.aforos)!=0) for(k in 1:(max(as.numeric(n.rangos.aforos))-1)) linea2 <- c(linea2, lim.inf+length(aforos.eti)*k)
      
      writeWorksheet(wb, data = "Aforo", sheet = clave_inst, startRow = lim.inf, startCol = 2 , header = F)
      if(length(n.rangos.aforos)!=0){  
        for(k in 1:max(as.numeric(n.rangos.aforos))){
          writeWorksheet(wb, sheet=clave_inst, data=levels(factor(n.aforos))[k], startRow=(lim.inf+(k-1)*length(aforos.eti)), startCol= 4, header=F)
        }
      }
      if(length(n.rangos.aforos)!=0) {
        writeWorksheet(wb, data = rep(aforos.eti, max(as.numeric(n.rangos.aforos))), sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
      } else{
        writeWorksheet(wb, data = aforos.eti, sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
      }
      writeWorksheet(wb, data = matriz.aforos, sheet = clave_inst, startRow = lim.inf, startCol = 6 , header = F)
      lim.inf <- lim.inf+nrow(matriz.aforos)
      linea1 <- c(linea1, lim.inf)
      for(k in 1:(length(hechos2.claves)-1)) linea2 <- c(linea2, lim.inf+length(levels(factor(hechos2.eti)))*k)
      
      writeWorksheet(wb, data = "Hechos generadores de comisiones", sheet = clave_inst, startRow = lim.inf, startCol = 2 , header = F)
      for(k in 1:length(hechos2.concep)){
        writeWorksheet(wb, data = hechos2.concep[k], sheet = clave_inst, startRow = lim.inf + (k-1)*(length(levels(factor(hechos2.eti)))), startCol = 4 , header = F)
      }
      for(k in 1:length(hechos2.claves)){
        writeWorksheet(wb, data = hechos2.claves[k], sheet = clave_inst, startRow = lim.inf + (k-1)*(length(levels(factor(hechos2.eti)))), startCol = 3 , header = F)
      }
      writeWorksheet(wb, data = hechos2.eti, sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
      for(k in 1:length(hechos2.claves)){
        writeWorksheet(wb, data = t(matriz.mda.base2[k,]), sheet = clave_inst, startRow = lim.inf+(k-1)*length(levels(factor(hechos2.eti))), startCol = 6 , header = F)
      }
      for(k in 1:length(hechos2.claves)){
        writeWorksheet(wb, data = t(matriz.imp.fi2[k,]), sheet = clave_inst, startRow = lim.inf+1+(k-1)*length(levels(factor(hechos2.eti))), startCol = 6 , header = F)
      }
      for(k in 1:length(hechos2.claves)){
        writeWorksheet(wb, data = t(matriz.imp.va2[k,]), sheet = clave_inst, startRow = lim.inf+2+(k-1)*length(levels(factor(hechos2.eti))), startCol = 6 , header = F)
      }
      for(k in 1:length(hechos2.claves)){
        writeWorksheet(wb, data = t(matriz.imp.ma2[k,]), sheet = clave_inst, startRow = lim.inf+3+(k-1)*length(levels(factor(hechos2.eti))), startCol = 6 , header = F)
      }
      lim.inf <- lim.inf + length(hechos2.eti)
      if(length(n.rangos.avaluos)!=0) for(k in 1:(length(n.rangos.avaluos))) linea2 <- c(linea2, lim.inf+length(avaluos.eti)*(k-1))
      
      writeWorksheet(wb, data = "Avaluo", sheet = clave_inst, startRow = lim.inf, startCol = 3 , header = F)
      if(length(n.rangos.avaluos)!=0) {
          for(k in 1:length(n.rangos.avaluos)){
            writeWorksheet(wb, data = levels(factor(n.avaluos))[k], sheet = clave_inst, startRow = (8*(k-1)+lim.inf), startCol = 4 , header = F)
          }
      }
      writeWorksheet(wb, data = rep(avaluos.eti, length(n.rangos.avaluos)), sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
      writeWorksheet(wb, data = matriz.avaluo, sheet = clave_inst, startRow = lim.inf, startCol = 6 , header = F)
      lim.inf <- lim.inf + nrow(matriz.avaluo)
      for(k in 1:length(comisiones.clave)) linea2 <- c(linea2, lim.inf+length(comisiones.etiq)*(k-1))
      
      
      for(k in 1:length(comisiones.clave)){
        writeWorksheet(wb, data = comisiones.clave[k], sheet = clave_inst, startRow = lim.inf+(4*(k-1)), startCol = 3 , header = F)
      }
      for(k in 0:(length(names.comisiones)-1)){
        writeWorksheet(wb, data = names.comisiones[k+1], sheet = clave_inst, startRow = lim.inf+(4*k), startCol = 4 , header = F)
      }
      writeWorksheet(wb, data = rep(comisiones.etiq, 4), sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
      for(k in 1:length(comisiones.clave)){
        writeWorksheet(wb, data = t(matriz.mda.base[k,]), sheet = clave_inst, startRow = lim.inf+(4*(k-1)), startCol = 6 , header = F)
      }
      for(k in 1:length(comisiones.clave)){
        writeWorksheet(wb, data = t(matriz.imp.fi[k,]), sheet = clave_inst, startRow = lim.inf+1+(4*(k-1)), startCol = 6 , header = F)
      }
      for(k in 1:length(comisiones.clave)){
        writeWorksheet(wb, data = t(matriz.imp.va[k,]), sheet = clave_inst, startRow = lim.inf+2+(4*(k-1)), startCol = 6 , header = F)
      }
      for(k in 1:length(comisiones.clave)){
        writeWorksheet(wb, data = t(matriz.imp.ma[k,]), sheet = clave_inst, startRow = lim.inf+3+(4*(k-1)), startCol = 6 , header = F)
      }
      lim.inf <- lim.inf + length(comisiones.etiq)*length(comisiones.clave)
      linea1 <- c(linea1, lim.inf)
      
      
      writeWorksheet(wb, data = "Rango de plazos", sheet = clave_inst, startRow = lim.inf, startCol = 2 , header = F)
      writeWorksheet(wb, data = plazos.concep, sheet = clave_inst, startRow = lim.inf, startCol = 4 , header = F)
      writeWorksheet(wb, data = plazos.eti, sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
      writeWorksheet(wb, data = matriz.plazos, sheet = clave_inst, startRow = lim.inf, startCol = 6 , header = F)
      lim.inf <- lim.inf+nrow(matriz.plazos)
      linea1 <- c(linea1, lim.inf)
      
      writeWorksheet(wb, data = "Pagos al millar fijos y/o en su caso", sheet = clave_inst, startRow = lim.inf, startCol = 2 , header = F)
      writeWorksheet(wb, data = "el inicial cuando el pago no es uniforme", sheet = clave_inst, startRow = lim.inf+1, startCol = 2 , header = F)
      writeWorksheet(wb, data = pagos.concep, sheet = clave_inst, startRow = lim.inf, startCol = 4 , header = F)
      writeWorksheet(wb, data = pagos.eti, sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
      writeWorksheet(wb, data = matriz.pagosxmil, sheet = clave_inst, startRow = lim.inf, startCol = 6 , header = F)
      lim.inf <- lim.inf+nrow(matriz.pagosxmil)
      linea1 <- c(linea1, lim.inf)
      if(length(n.plazos)!=0) for(k in 1:(max(n.plazos)-1)) linea2 <- c(linea2, lim.inf+length(pagosnounif.eti)*k)
      
      writeWorksheet(wb, data = "Pago por mil no uniforme (incremento o decremento)", sheet = clave_inst, startRow = lim.inf, startCol = 2 , header = F)
      if(max(n.plazos)!=0){
        writeWorksheet(wb, data = rep(pagosnounif.concep, max(n.plazos)), sheet = clave_inst, startRow = lim.inf, startCol = 4 , header = F)
        writeWorksheet(wb, data = rep(pagosnounif.eti, max(n.plazos)), sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
      } else{
        writeWorksheet(wb, data = pagosnounif.concep, sheet = clave_inst, startRow = lim.inf, startCol = 4 , header = F)
        writeWorksheet(wb, data = pagosnounif.eti, sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
      }
      writeWorksheet(wb, data = matriz.pagosnounif, sheet = clave_inst, startRow = lim.inf, startCol = 6 , header = F)
      
      w <- lim.inf
      if(max(n.plazos)!=0){
        for(k in 2:(max(n.plazos)+1)){
          w[k] <- w[k-1]+length(pagosnounif.eti)
        }
        for (k in 1:max(n.plazos)){
          writeWorksheet(wb, data = rownames(matriz.pagosxmil)[matriz.pagosxmil[, which(n.plazos==max(n.plazos))[1]]!=0][k], sheet = clave_inst, startRow = w[1:max(n.plazos)][k], startCol = 3 , header = F)
        }
      }
      lim.inf <- lim.inf+nrow(matriz.pagosnounif)
      linea1 <- c(linea1, lim.inf)
      

      if(length(niveles)!=0){
        for(k in 1:(length(levels(factor(niveles)))-1)) linea2 <- c(linea2, lim.inf+length(tasas.eti)*k)
        
        writeWorksheet(wb, data = "Tasas de interes", sheet = clave_inst, startRow = lim.inf, startCol = 2 , header = F)
        for(k in 1:length(levels(factor(niveles)))){
          writeWorksheet(wb, data = niveles[k], sheet = clave_inst, startRow = (lim.inf + (k-1)*length(tasas.concep)), startCol = 3 , header = F)
        }
        writeWorksheet(wb, data = rep(tasas.concep2, length(levels(factor(niveles)))), sheet = clave_inst, startRow = lim.inf, startCol = 4 , header = F)
        writeWorksheet(wb, data = rep(tasas.eti2, length(levels(factor(niveles)))), sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
        writeWorksheet(wb, data = matriz.tasas, sheet = clave_inst, startRow = lim.inf, startCol = 6 , header = F)
        lim.inf <- lim.inf + nrow(matriz.tasas)
        linea1 <- c(linea1, lim.inf)
        for(k in 1:(length(claves.seguros)-1)) linea2 <- c(linea2, lim.inf+length(levels(factor(seguros.eti)))*k)
        
        writeWorksheet(wb, data = "Seguros", sheet = clave_inst, startRow = lim.inf, startCol = 2 , header = F)
        for(k in 1:length(claves.seguros)){
          writeWorksheet(wb, data = claves.seguros[k], sheet = clave_inst, startRow = (lim.inf+(k-1)*length(levels(factor(seguros.eti)))), startCol = 3 , header = F)
        }
        for(k in 1:length(levels(factor(seguros.eti)))){
          writeWorksheet(wb, data = seguros.concep[k], sheet = clave_inst, startRow = lim.inf + length(levels(factor(seguros.eti)))*(k-1), startCol = 4 , header = F)
        }
        writeWorksheet(wb, data = seguros.eti, sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
        writeWorksheet(wb, data = matriz.seguros, sheet = clave_inst, startRow = lim.inf, startCol = 6 , header = F)
        lim.inf <- lim.inf+nrow(matriz.seguros)
        linea1 <- c(linea1, lim.inf)
        for(k in 1:(length(claves.segurosg)-1)) linea2 <- c(linea2, lim.inf+length(c(segurosg.eti.v,bonis.eti.v))*k)
        
        
        for(k in 1:length(claves.segurosg)){
          writeWorksheet(wb, data = claves.segurosg[1], sheet = clave_inst, startRow = lim.inf + (k-1)*length(levels(factor(segurosg.eti.v))), startCol = 3 , header = F)
        }
        writeWorksheet(wb, data = segurosg.concep.v, sheet = clave_inst, startRow = lim.inf, startCol = 4 , header = F)
        writeWorksheet(wb, data = bonis.concep.v, sheet = clave_inst, startRow = lim.inf + length(levels(factor(segurosg.eti.v))), startCol = 4 , header = F)
        writeWorksheet(wb, data = segurosg.eti.v, sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
        writeWorksheet(wb, data = bonis.eti.v, sheet = clave_inst, startRow = lim.inf + length(levels(factor(segurosg.eti.v))), startCol = 5 , header = F)
        writeWorksheet(wb, data = matriz.segurosg.v, sheet = clave_inst, startRow = lim.inf, startCol = 6 , header = F)
        
        lim.inf <- lim.inf+nrow(matriz.segurosg.v)
        linea1 <- c(linea1, lim.inf)
        for(k in 1:(length(claves.segurosg)-1)) linea2 <- c(linea2, lim.inf+length(c(segurosg.eti.d,bonis.eti.d))*k)
        
        
        for(k in 1:length(claves.segurosg)){
          writeWorksheet(wb, data = claves.segurosg[2], sheet = clave_inst, startRow = lim.inf + (k-1)*length(levels(factor(segurosg.eti.d))), startCol = 3 , header = F)
        }
        writeWorksheet(wb, data = segurosg.concep.d, sheet = clave_inst, startRow = lim.inf, startCol = 4 , header = F)
        writeWorksheet(wb, data = bonis.concep.d, sheet = clave_inst, startRow = lim.inf + length(levels(factor(segurosg.eti.d))), startCol = 4 , header = F)
        if(length(levels(niveles.segurosg.d))!=0){
          writeWorksheet(wb, data = rep(c(segurosg.eti.d, bonis.eti.d), levels(niveles.segurosg.d)), sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
        } else{
          writeWorksheet(wb, data = c(segurosg.eti.d, bonis.eti.d), sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
        }
        writeWorksheet(wb, data = matriz.segurosg.d, sheet = clave_inst, startRow = lim.inf, startCol = 6 , header = F)
        lim.inf <- lim.inf+nrow(matriz.segurosg.d)
        linea1 <- c(linea1, lim.inf)
        for(k in 1:length(niveles.cat)) linea2 <- c(linea2, lim.inf+length(c(cat.eti[1], cat.eti[8:14]))*k-2)

        writeWorksheet(wb, data = "CAT", sheet = clave_inst, startRow = lim.inf, startCol = 2 , header = F)
        writeWorksheet(wb, data = "Valores de vivienda", sheet = clave_inst, startRow = lim.inf, startCol = 3 , header = F)
        writeWorksheet(wb, data = "Resultado de CAT", sheet = clave_inst, startRow = lim.inf+length(cat.eti[2:7]), startCol = 3 , header = F)
        writeWorksheet(wb, data = cat.concep[2:7], sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
        writeWorksheet(wb, data = cat.eti[2:7], sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
        writeWorksheet(wb, data = rep(c(cat.eti[1], cat.eti[8:14]), length(niveles.cat)), sheet = clave_inst, startRow = lim.inf+length(cat.eti[2:7]), startCol = 5 , header = F)
        writeWorksheet(wb, data = rep(c(cat.concep[1], cat.concep[8:14]), length(niveles.cat)), sheet = clave_inst, startRow = lim.inf+length(cat.eti[2:7]), startCol = 5 , header = F)
        writeWorksheet(wb, data = matriz.cat.vv, sheet = clave_inst, startRow = lim.inf, startCol = 6 , header = F)
        writeWorksheet(wb, data = matriz.cat.subprods, sheet = clave_inst, startRow = lim.inf+length(cat.eti[2:7]), startCol = 6 , header = F)
        lim.inf <- lim.inf+nrow(matriz.cat.vv)+nrow(matriz.cat.subprods)
        
      } else{
        writeWorksheet(wb, data = "Tasas de interes", sheet = clave_inst, startRow = lim.inf, startCol = 2 , header = F)
        writeWorksheet(wb, data = tasas.concep[1:4], sheet = clave_inst, startRow = lim.inf, startCol = 4 , header = F)
        writeWorksheet(wb, data = tasas.eti[1:4], sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
        writeWorksheet(wb, data = matriz.tasas, sheet = clave_inst, startRow = lim.inf, startCol = 6 , header = F)
        lim.inf <- lim.inf + nrow(matriz.tasas)
        linea1 <- c(linea1, lim.inf)
        
        writeWorksheet(wb, data = "Seguros", sheet = clave_inst, startRow =lim.inf, startCol = 2 , header = F)
        for(k in 1:length(claves.seguros)){
          writeWorksheet(wb, data = claves.seguros[k], sheet = clave_inst, startRow = lim.inf+(k-1)*length(levels(factor(seguros.eti))), startCol = 3 , header = F)
        }
        for(k in 1:length(levels(factor(seguros.eti)))){
          writeWorksheet(wb, data = seguros.concep[k], sheet = clave_inst, startRow = lim.inf + length(levels(factor(seguros.eti)))*(k-1), startCol = 4 , header = F)
        }
        writeWorksheet(wb, data = seguros.eti, sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
        writeWorksheet(wb, data = matriz.seguros, sheet = clave_inst, startRow = lim.inf, startCol = 6 , header = F)
        lim.inf <- lim.inf+nrow(matriz.seguros)
        linea1 <- c(linea1, lim.inf)
        
        for(k in 1:length(claves.segurosg)){
          writeWorksheet(wb, data = claves.segurosg[1], sheet = clave_inst, startRow = lim.inf + (k-1)*length(levels(factor(segurosg.eti.v))), startCol = 3 , header = F)
        }
        writeWorksheet(wb, data = segurosg.concep.v, sheet = clave_inst, startRow = lim.inf, startCol = 4 , header = F)
        writeWorksheet(wb, data = bonis.concep.v, sheet = clave_inst, startRow = lim.inf + length(levels(factor(segurosg.eti.v))), startCol = 4 , header = F)
        writeWorksheet(wb, data = segurosg.eti.v, sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
        writeWorksheet(wb, data = bonis.eti.v, sheet = clave_inst, startRow = lim.inf + length(levels(factor(segurosg.eti.v))), startCol = 5 , header = F)
        writeWorksheet(wb, data = matriz.segurosg.v, sheet = clave_inst, startRow = lim.inf, startCol = 6 , header = F)
        
        lim.inf <- lim.inf+nrow(matriz.segurosg.v)
        linea1 <- c(linea1, lim.inf)
        
        for(k in 1:length(claves.segurosg)){
          writeWorksheet(wb, data = claves.segurosg[2], sheet = clave_inst, startRow = lim.inf + (k-1)*length(levels(factor(segurosg.eti.d))), startCol = 3 , header = F)
        }
        writeWorksheet(wb, data = segurosg.concep.d, sheet = clave_inst, startRow = lim.inf, startCol = 4 , header = F)
        writeWorksheet(wb, data = bonis.concep.d, sheet = clave_inst, startRow = lim.inf + length(levels(factor(segurosg.eti.d))), startCol = 4 , header = F)
        writeWorksheet(wb, data = c(segurosg.eti.d, bonis.eti.d), sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
        writeWorksheet(wb, data = matriz.segurosg.d, sheet = clave_inst, startRow = lim.inf, startCol = 6 , header = F)
        lim.inf <- lim.inf+nrow(matriz.segurosg.d)
        linea1 <- c(linea1, lim.inf)
        
        writeWorksheet(wb, data = "CAT", sheet = clave_inst, startRow = lim.inf, startCol = 2 , header = F)
        writeWorksheet(wb, data = "Valores de vivienda", sheet = clave_inst, startRow = lim.inf, startCol = 3 , header = F)
        writeWorksheet(wb, data = "Resultado de CAT", sheet = clave_inst, startRow = lim.inf+length(cat.eti[2:7]), startCol = 3 , header = F)
        writeWorksheet(wb, data = cat.concep[2:7], sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
        writeWorksheet(wb, data = cat.eti[2:7], sheet = clave_inst, startRow = lim.inf, startCol = 5 , header = F)
        writeWorksheet(wb, data = c(cat.eti[1], cat.eti[8:14]), sheet = clave_inst, startRow = lim.inf+length(cat.eti[2:7]), startCol = 5 , header = F)
        writeWorksheet(wb, data = c(cat.concep[1], cat.concep[8:14]), sheet = clave_inst, startRow = lim.inf+length(cat.eti[2:7]), startCol = 5 , header = F)
        writeWorksheet(wb, data = matriz.cat.vv, sheet = clave_inst, startRow = lim.inf, startCol = 6 , header = F)
        writeWorksheet(wb, data = matriz.cat.subprods, sheet = clave_inst, startRow = lim.inf+length(cat.eti[2:7]), startCol = 6 , header = F)
        lim.inf <- lim.inf+nrow(matriz.cat.vv)+nrow(matriz.cat.subprods)
      }
    }
    # ================================== Formato de celdas :] ==============================
    {
      setColumnWidth(wb, sheet = clave_inst, column = c(3,5:(6+length(c(prods.vigentes)))), width = -1)
      setColumnWidth(wb, sheet=clave_inst, column=c(2,4), width=12000)
      
      BordeInferior <- createCellStyle(wb)
      setBorder(BordeInferior, side=c("bottom"), type=XLC$BORDER.MEDIUM, color=XLC$COLOR.BLACK)
      BordeSuperior <- createCellStyle(wb)
      setBorder(BordeSuperior, side=c("top"), type=XLC$BORDER.MEDIUM, color=XLC$COLOR.BLACK)
      BordeDerecho <- createCellStyle(wb)
      setBorder(BordeDerecho, side=c("left"), type=XLC$BORDER.MEDIUM, color=XLC$COLOR.BLACK)
      BordeIzquierdo <- createCellStyle(wb)
      setBorder(BordeIzquierdo, side=c("right"), type=XLC$BORDER.MEDIUM, color=XLC$COLOR.BLACK)
      
      Linea <- createCellStyle(wb)
      setBorder(Linea, side="top", type=XLC$BORDER.THIN, color=XLC$COLOR.BLACK)
      
      ColorAzul <- createCellStyle(wb)
      setFillPattern(ColorAzul, fill=XLC$FILL.SOLID_FOREGROUND)
      setFillForegroundColor(ColorAzul, color=XLC$COLOR.PALE_BLUE)
      
      ColorGris <- createCellStyle(wb)
      setFillPattern(ColorGris, fill=XLC$FILL.SOLID_FOREGROUND)
      setFillForegroundColor(ColorGris, color=XLC$COLOR.GREY_25_PERCENT)
      
      ColorGrisLinea <- createCellStyle(wb)
      setFillPattern(ColorGrisLinea, fill=XLC$FILL.SOLID_FOREGROUND)
      setFillForegroundColor(ColorGrisLinea, color=XLC$COLOR.GREY_25_PERCENT)
      setBorder(ColorGrisLinea, side="top", type=XLC$BORDER.THIN, color=XLC$COLOR.BLACK)
      
      ColorAzulLinea <- createCellStyle(wb)
      setFillPattern(ColorAzulLinea, fill=XLC$FILL.SOLID_FOREGROUND)
      setFillForegroundColor(ColorAzulLinea, color=XLC$COLOR.PALE_BLUE)
      setBorder(ColorAzulLinea, side="top", type=XLC$BORDER.THIN, color=XLC$COLOR.BLACK)
      
      Fecha <- createCellStyle(wb)
      setDataFormat(Fecha, format="d-m-yyyy")
      
      createFreezePane(wb, sheet=clave_inst, 6,9)
      
      setCellStyle(wb, sheet = clave_inst, row=6, col=(2:(5+length(prods.vigentes))), cellstyle=BordeInferior)
      
      setCellStyle(wb, sheet = clave_inst, row=lim.inf, col=(2:(5+length(prods.vigentes))), cellstyle=BordeSuperior)
      setCellStyle(wb, sheet = clave_inst, row=7:(lim.inf-1), col=1, cellstyle=BordeIzquierdo)
      setCellStyle(wb, sheet = clave_inst, row=7:(lim.inf-1), col=6+length(prods.vigentes), cellstyle=BordeDerecho)
      
      z <- 9:(lim.inf-1)
      setCellStyle(wb, sheet = clave_inst, row=9:(lim.inf-1), col=3, cellstyle=ColorGris)
      setCellStyle(wb, sheet = clave_inst, row=9:(lim.inf-1), col=4, cellstyle=ColorAzul)
      setCellStyle(wb, sheet = clave_inst, row=9:(lim.inf-1), col=5, cellstyle=ColorGris)
      
      for(k in linea1){
        setCellStyle(wb, sheet = clave_inst, row=k, col=(2:(5+length(prods.vigentes))), cellstyle=Linea)
      }
      
      for(k in linea2){
        setCellStyle(wb, sheet = clave_inst, row=k, col=(3:(5+length(prods.vigentes))), cellstyle=Linea)
      }
      for(k in c(linea1,linea2)){
        setCellStyle(wb, sheet=clave_inst, row=k, col=c(3,5), cellstyle=ColorGrisLinea)
      }
      for(k in c(linea1,linea2)){
        setCellStyle(wb, sheet=clave_inst, row=k, col=4, cellstyle=ColorAzulLinea)
      }
    }
    # ================================== Guardar :] ========================================
    if(toge==F){
      saveWorkbook(wb)
    }
  }
  if(toge==T){
    saveWorkbook(wb)
  }
}

# ======================================== Otros ========================================  
system.time(reporte(archivo="Carga de información 10042014_1.xlsx"))
system.time(reporte(datos=T, toge=F))