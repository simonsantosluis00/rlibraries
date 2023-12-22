#' Function to scrape the newest data
#'
#' @param x just run the function
#' @return The newest data available
#' @export
scraper <- function() {
  get_desktop_path <- function() {
    if (Sys.info()['sysname'] == 'Windows') {
      command <- 'cmd.exe /c echo %USERPROFILE%\\Documents'
      desktop_path <- shell(command, intern = TRUE)
      # Remove any trailing whitespace or newline characters
      desktop_path <- gsub("[\r\n]", "", desktop_path)
      
      if (file.exists(desktop_path)) {
        return(desktop_path)
      } else {
        # If Documents folder doesn't exist, try "Documentos" for Spanish systems
        command_spanish <- 'cmd.exe /c echo %USERPROFILE%\\Documentos'
        desktop_path_spanish <- shell(command_spanish, intern = TRUE)
        desktop_path_spanish <- gsub("[\r\n]", "", desktop_path_spanish)
        
        if (file.exists(desktop_path_spanish)) {
          return(desktop_path_spanish)
        } else {
          return(NULL)  # Neither Documents nor Documentos folder found
        }
      }
    } else {
      return(NULL)  # For non-Windows systems
    }
  }
  
  
  # Get the modified desktop path
  Documents_path <- get_desktop_path()
  download_directory <- paste0(Documents_path, "\\Integration\\DOWNLOADS")
  move_directory <- paste0(Documents_path, "\\Integration\\DATA\\MODELO")
  
  download_directory
  move_directory
  
  library(RSelenium)
  library(wdman)
  library(netstat)
  library(tidyverse)
  library(readxl)
  library(openxlsx)
  
  system("taskkill /im java.exe /f", intern=FALSE, ignore.stdout=FALSE)
  
  # Set the download directory path
  download_directory <- download_directory
  
  # Set Chrome preferences
  eCaps <- list(
    chromeOptions = 
      list(prefs = list(
        "download.default_directory" = download_directory
      )
      )
  )
  
  eCaps 
  
  rs_driver_object <- rsDriver( browser = "chrome", chromever = "119.0.6045.105", extraCapabilities = eCaps)
  
  remDr <- rs_driver_object$client
  
  
  remDr$navigate("https://flexicredit.vortem.com/web#action=189&model=vor.cre_prestamos&view_type=kanban&cids=1&menu_id=124")
  
  Sys.sleep(.5)
  
  remDr$maxWindowSize()
  
  username <- remDr$findElement(using = 'xpath', '//*[@id="login"]') 
  username$sendKeysToElement(list("finanzas3@flexicredit.mx"))
  
  password <- remDr$findElement(using = 'xpath', '//*[@id="password"]') 
  password$sendKeysToElement(list("Tioflexi!12345"))
  
  submit_button <- remDr$findElement(using = 'xpath', '/html/body/div[1]/main/div/form/div[3]/button')
  submit_button$clickElement()
  
  Sys.sleep(4)
  
  Cartera_de_credito <- remDr$findElement(using = 'link text', 'Cartera de Crédito')
  Cartera_de_credito$clickElement()
  
  Ver_prestamos <- remDr$findElement(using = 'link text', 'Ver Préstamos')
  Ver_prestamos$clickElement()
  
  Ver_lista <- remDr$findElement(using = 'xpath', '//button[@aria-label="Ver list"]')
  Ver_lista$clickElement()
  
  Sys.sleep(1)
  
  Checkbox_todo <- remDr$findElement(using = 'xpath', '//input[@class="custom-control-input"]') 
  Checkbox_todo$clickElement()
  
  Seleccionar_todos <- remDr$findElement(using = 'xpath', '//a[@class="o_list_select_domain"]')
  Seleccionar_todos$clickElement()
  
  Accion <- remDr$findElement(using = 'xpath', '//button[@aria-expanded="false"]')
  Accion$clickElement()
  
  Accion2 <- remDr$findElement(using = 'xpath', '//button[@aria-expanded="false"]')
  Accion2$clickElement()
  
  Exportar <- remDr$findElement(using = 'xpath', '//a[@role="menuitemcheckbox"]')
  Exportar$clickElement()
  
  Sys.sleep(2)
  
  Plantilla <- remDr$findElement(using = 'xpath', '//select[@class="form-control ml-4 o_exported_lists_select"]')
  Plantilla$clickElement()
  
  Plantilla_BaseDeDatosModelov1<- remDr$findElement(using = 'xpath', '//option[@value="85"]')
  Plantilla_BaseDeDatosModelov1$clickElement()
  
  Sys.sleep(1)
  
  ExportarArchivo <- remDr$findElement(using = 'xpath', '//button[@class="btn btn-primary"]')
  ExportarArchivo$clickElement()
  
  Sys.sleep(60)
  
  downloads_folder <- download_directory
  destination_folder <- move_directory
  downloaded_files <- list.files(downloads_folder, full.names = TRUE, recursive = FALSE)
  downloaded_files <- downloaded_files[file.info(downloaded_files)$isdir == FALSE]  # Exclude directories
  downloaded_files <- downloaded_files[order(file.info(downloaded_files)$mtime, decreasing = TRUE)]
  new_file_name <- "BaseDeDatosModelo.xlsx"
  
  if (length(downloaded_files) > 0) {
    # Get the most recent file (the first one in the sorted list)
    most_recent_file <- downloaded_files[1]
    
    # Define the destination path for the most recent file
    destination_path <- file.path(destination_folder, new_file_name)
    
    # Move the most recent file to the destination folder
    file.rename(most_recent_file, destination_path)
    cat("Moved file:", most_recent_file, "\nTo destination:", destination_path, "\n")
  } else {
    cat("No files to move in the Downloads folder.\n")
  }
  
  download_directory <- paste0(Documents_path, "\\Integration\\DOWNLOADS")
  move_directory <- paste0(Documents_path, "\\Integration\\DATA\\MODELO\\PLAZOSINPAGO")
  
  Sys.sleep(2)
  
  Plantilla <- remDr$findElement(using = 'xpath', '//select[@class="form-control ml-4 o_exported_lists_select"]')
  Plantilla$clickElement()
  
  Plantilla_BaseDeDatosModelov1<- remDr$findElement(using = 'xpath', '//option[@value="91"]')
  Plantilla_BaseDeDatosModelov1$clickElement()
  
  Sys.sleep(1)
  
  ExportarArchivo <- remDr$findElement(using = 'xpath', '//button[@class="btn btn-primary"]')
  ExportarArchivo$clickElement()
  
  Sys.sleep(60)
  
  downloads_folder <- download_directory
  destination_folder <- move_directory
  downloaded_files <- list.files(downloads_folder, full.names = TRUE, recursive = FALSE)
  downloaded_files <- downloaded_files[file.info(downloaded_files)$isdir == FALSE]  # Exclude directories
  downloaded_files <- downloaded_files[order(file.info(downloaded_files)$mtime, decreasing = TRUE)]
  new_file_name <- "Flexicredit - MAC Ultimo Pago.xlsx"
  
  if (length(downloaded_files) > 0) {
    # Get the most recent file (the first one in the sorted list)
    most_recent_file <- downloaded_files[1]
    
    # Define the destination path for the most recent file
    destination_path <- file.path(destination_folder, new_file_name)
    
    # Move the most recent file to the destination folder
    file.rename(most_recent_file, destination_path)
    cat("Moved file:", most_recent_file, "\nTo destination:", destination_path, "\n")
  } else {
    cat("No files to move in the Downloads folder.\n")
  }
  
  remDr$close()
  
  Sys.sleep(0)
  
}

#' Function to send last credits available for each business
#'
#' @param empresa tell which business to send latest data 
#' @return send emails
#' @export
business <- function(empresa) {
  library(readxl)
  library(openxlsx)
  library(dplyr)
  # Pedir al usuario que ingrese un número
  empresa <- readline(prompt = "Ingresa un nombre de empresa: ")
  
  
  get_desktop_path <- function() {
    if (Sys.info()['sysname'] == 'Windows') {
      command <- 'cmd.exe /c echo %USERPROFILE%\\Documents'
      desktop_path <- shell(command, intern = TRUE)
      # Remove any trailing whitespace or newline characters
      desktop_path <- gsub("[\r\n]", "", desktop_path)
      
      if (file.exists(desktop_path)) {
        return(desktop_path)
      } else {
        # If Documents folder doesn't exist, try "Documentos" for Spanish systems
        command_spanish <- 'cmd.exe /c echo %USERPROFILE%\\Documentos'
        desktop_path_spanish <- shell(command_spanish, intern = TRUE)
        desktop_path_spanish <- gsub("[\r\n]", "", desktop_path_spanish)
        
        if (file.exists(desktop_path_spanish)) {
          return(desktop_path_spanish)
        } else {
          return(NULL)  # Neither Documents nor Documentos folder found
        }
      }
    } else {
      return(NULL)  # For non-Windows systems
    }
  }
  
  # Get the modified desktop path
  Documents_path <- get_desktop_path()
  empresa_directory <- paste0(Documents_path, "\\Integration\\DATA\\FINANCIEROS\\", empresa, "\\FORMATS")
  move_directory <- paste0(Documents_path, "\\Integration\\DATA\\FINANCIEROS\\", empresa,"\\ENTREGABLE")
  sent_directory <- paste0(Documents_path, "\\Integration\\DATA\\FINANCIEROS\\", empresa,"\\ENVIADOS")
  
  
  sent_files <- list.files(sent_directory, full.names = TRUE, recursive = FALSE)
  sent_files <- sent_files[file.info(sent_files)$isdir == FALSE]
  sent_files <- sent_files[order(file.info(sent_files)$mtime, decreasing = TRUE)]
  most_recent_file_sent <- sent_files[1]
  most_recent_file_sent
  excel_path_basededatos <- paste0(Documents_path,"\\Integration\\DATA\\MODELO\\BaseDeDatosModelo.xlsx")
  
  excel_file_basededatos <- excel_path_basededatos
  wb <- loadWorkbook(excel_file_basededatos)
  sheet_basededatos <- 1
  col_types <- c("text", "text", "date", 
                 "text", "numeric", "date", "date", 
                 "numeric", "numeric", "numeric", 
                 "numeric", "text", "text", "numeric", 
                 "text", "text", "date", "numeric", 
                 "numeric", "date", "text", "text", 
                 "text", "text", "date", "numeric", 
                 "numeric", "text")
  
  base_datos1 <- read_excel(excel_path_basededatos, sheet = "Sheet1", col_types = col_types)
  #base_datos <- base_datos1[,1:13]
  
  columns_to_find <- c("Fecha de Desembolso", "Numero de Contrato", "Solicitud/Numero Nómina", "Cliente",
                       "Empresa donde Labora", "Periodicidad", "Pago Periodico", "Primera Amortización",
                       "Numero de Periodos", "Monto Desembolsado", "Suma Vencimientos", "State", "Fecha Final")
  
  column_numbers <- numeric(length(columns_to_find))
  
  for (i in seq_along(columns_to_find)) {
    column_numbers[i] <- which(names(base_datos1) == columns_to_find[i])
  }
  
  base_datos_pagosV3 <- base_datos1[,column_numbers]
  base_datos_pagosV3
  
  
  # Parte del nombre del archivo que conoces
  #partial_filename <- paste0("../DATA/FINANCIEROS/",empresa,"/ENVIADOS/", "FlexiCredit - Listado de Descuentos ", empresa)
  
  # Encontrar archivos que coincidan con la parte conocida del nombre
  #matching_files <- list.files(path = dirname(partial_filename), pattern = basename(partial_filename))
  
  # Mostrar los archivos coincidentes encontrados
  #print(matching_files)
  
  
  #path <- paste0("../DATA/FINANCIEROS/",empresa,"/ENVIADOS/", matching_files)
  
  denso_descuento <- read_excel(most_recent_file_sent, 
                                skip = 4)
  
  # Identify rows where `Empresa donde Labora` matches certain values and replace them
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("ALLENDE VALLE DEL NORTE SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Allende"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("ARFINSA SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Arfinsa"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("BRITE LITE DE MEXICO SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Britelite"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("BURO JURIDICO INTEGRAL DE COBRANZA SC")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Buro Juridico"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("PRODUCTOS DE CALIZA SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Caliza"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("CARMEN PATRICIA SAENZ TREVIÑO")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Carmen Patricia Saenz"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("CARNES SAN MIGUEL DE SALTILLO SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Carnes San Miguel"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("COMEDORES Y BANQUETES SA DE CV", "Comedores Y Banquetes")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Comedores Y Banquetes"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("DENSO MEXICO SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Denso"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("DOGA CNC MAQUINADOS SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Doga"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("EDER ALEJANDRO TORRES PEREZ")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Eder Alejandro Torres"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("ELOISA YAMILET PEREZ DELGADILLO")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Eloisa Yamilet"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("EDIFIKA MANTENIMIENTOS PLAZAS Y PARQUES SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Edifika"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("FLEXICREDIT SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "FlexiCredit"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("GRANIX SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Granix"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("INTERCONSTRUCTORA SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Interconstructora"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("LADRILLERA MECANIZADA, SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Ladrillera"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("PRODUCTOS LA TRADICIONAL SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "La Tradicional"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("MAQUINARIA Y FLETES SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Maquinaria Y Fletes"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("MARIA DE LA LUZ FIERROS GALLEGOS")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Maria De La Luz Fierros"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("MAURICIO ORDOÑEZ GARZA")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Mauricio Ordonez"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("MAXIWELD SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Maxiweld"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("MIGUEL DANIEL CHAVARRIA ALVAREZ")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Miguel Daniel Chavarria"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("MONDELEZ MEXICO, SA DE RL DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Mondelez"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("MUEBLES KRILL SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Muebles Krill"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("NATIONAL UNITY SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "National Unity"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("PLANTASFALTOS SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Plantasfaltos"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("CORPORACION TATSUMI DE MEXICO SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Tatsumi"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("LOGISTICA Y SOLUCIONES ROAN SC")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Roan"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("SERVICIO INTEGRAL DE SEGURIDAD PRIVADA MG SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Seguridad MG"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("SEGURIDAD PRIVADA INTEGRAL PISCIS SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Seguridad Piscis"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("SCORE EVENTOS SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Score"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("TALENTO EJECUTIVO SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Talento Ejecutivo"
  
  replace_indices <- base_datos_pagosV3$`Empresa donde Labora` %in% c("COMERCIALIZADORA TRENT SA DE CV")
  base_datos_pagosV3$`Empresa donde Labora`[replace_indices] <- "Trent"
  
  
  
  denso_descuento$`Fecha de Pago` <- as.Date(denso_descuento$`Fecha de Pago`)
  
  denso_descuento_last <- tail(denso_descuento, 1)
  denso_descuento_last <- denso_descuento_last[,1]
  
  denso_descunto_last <- as.numeric(denso_descuento_last)
  indices <- which(base_datos_pagosV3$`Solicitud/Numero Nómina` == denso_descunto_last)
  if (length(indices) > 0){
    index_max <- indices[which.max(indices)] + 1 
    
  } else {
    
  }
  
  denso_descuento_last
  
  base_datos_pagosV3 <- base_datos_pagosV3[index_max:nrow(base_datos_pagosV3),]
  base_datos_pagosV3
  
  
  naindex <- which(is.na(base_datos_pagosV3$`Monto Desembolsado`))
  base_datos_pagosV3[-naindex,]
  
  base_datos_pagosV3 <- base_datos_pagosV3[(base_datos_pagosV3$`Empresa donde Labora` == empresa),]
  base_datos_pagosV3 <-base_datos_pagosV3[1:nrow(base_datos_pagosV3),]
  
  length <- length(base_datos_pagosV3) - 2
  base_datos_pagosV3 <- base_datos_pagosV3[,3:length]
  
  #
  base_datos_pagosV3$`Primera Amortización` <- as.Date(base_datos_pagosV3$`Primera Amortización`)
  
  library(lubridate)
  
  # ymd
  base_datos_pagosV3$`Primera Amortización` <- ymd(base_datos_pagosV3$`Primera Amortización`)
  
  #
  base_datos_pagosV3$`Primera Amortización` <- if_else(
    wday(base_datos_pagosV3$`Primera Amortización`) == 6,
    base_datos_pagosV3$`Primera Amortización`,
    base_datos_pagosV3$`Primera Amortización` + days(6 - wday(base_datos_pagosV3$`Primera Amortización`) + ifelse(wday(base_datos_pagosV3$`Primera Amortización`) > 6, 7, 0))
  )
  
  library(openxlsx)
  
  
  # Parte del nombre del archivo que conoces
  partial_filename <- paste0(Documents_path, "\\Integration\\DATA\\FINANCIEROS\\", empresa, "\\FORMATS\\", "FlexiCredit - Listado de Descuentos ", empresa)
  
  # Encontrar archivos que coincidan con la parte conocida del nombre
  matching_files <- list.files(path = dirname(partial_filename), pattern = basename(partial_filename))
  
  # Mostrar los archivos coincidentes encontrados
  matching_files
  
  file_path <- paste0(Documents_path, "\\Integration\\DATA\\FINANCIEROS\\", empresa, "\\FORMATS\\", matching_files)
  
  file_path
  
  
  if (!file.exists(file_path)) {
    # Create a new workbook
    wb <- createWorkbook()
    
    # Add a new sheet to the newly created workbook
    addWorksheet(wb, "MySheet")
  } else {
    # Load the existing workbook
    wb <- loadWorkbook(file_path)
    
    # Extract the file name without the extension
    file_name <- tools::file_path_sans_ext(basename(file_path))
    
    # Replace the date part in the file name with the current system date
    new_date <- format(Sys.Date(), "%Y-%m-%d")
    new_file_name <- gsub("\\d{4}-\\d{2}-\\d{2}", new_date, file_name)
    
    # Generate the new file path
    new_file_path <- file.path(dirname(file_path), paste0(new_file_name, ".xlsx"))
    
    # Save the workbook with the updated file name
    saveWorkbook(wb, new_file_path)
    
    # Load the modified workbook
    wb <- loadWorkbook(new_file_path)
  }
  
  # Rest of your code remains unchanged from here onwards
  
  # Starting row and column to write the data
  start_row <- 6
  start_col <- 2
  start_row2 <- start_row + 1
  
  # Get dimensions of the existing data
  existing_data <- readWorkbook(wb, sheet = "MySheet", startRow = start_row)
  existing_rows <- nrow(existing_data)
  
  # Clear existing content from the specified range
  for (i in start_row2:(start_row2 + existing_rows - 1)) {
    for (j in start_col:(start_col + ncol(existing_data) - 1)) {
      writeData(wb, "MySheet", "", startRow = i, startCol = j)
    }
  }
  
  # Write data to the first sheet
  writeData(wb, "MySheet", base_datos_pagosV3, startRow = start_row, startCol = start_col)
  
  # Save the workbook to the specified location
  saveWorkbook(wb, new_file_path, overwrite = TRUE)
  
  
  # Get today's date
  today <- Sys.Date()
  
  if (wday(today) == 2) {  # Check if today is Monday (wday = 2 for Monday)
    nearest_monday <- today  # Keep today's date as Monday
  } else {
    # Calculate the nearest Monday
    days_to_monday <- ifelse(wday(today) < 2, 2 - wday(today), 9 - wday(today))
    nearest_monday <- today + days(days_to_monday)
  }
  
  empresas_folder <- empresa_directory
  destination_folder <- move_directory
  
  empresas_files <- list.files(empresas_folder, full.names = TRUE, recursive = FALSE)
  empresas_files <- empresas_files[file.info(empresas_files)$isdir == FALSE]
  empresas_files <- empresas_files[order(file.info(empresas_files)$mtime, decreasing = TRUE)]
  new_file_name <- paste0("FlexiCredit - Listado de Descuentos ", empresa, " ", nearest_monday, ".xlsx")
  
  
  if (length(empresas_files) > 0) {
    # Get the most recent file (the first one in the sorted list)
    most_recent_file <- empresas_files[1]
    
    # Define the destination path for the most recent file
    destination_path <- file.path(destination_folder, new_file_name)
    
    # Move the most recent file to the destination folder
    file.rename(most_recent_file, destination_path)
    cat("Moved file:", most_recent_file, "\nTo destination:", destination_path, "\n")
  } else {
    cat("No files to move in the Downloads folder.\n")
  }
  
  
  library(RDCOMClient)
  empresas <- c("Denso", "Mondelez", "Ladrillera", "Muebles Krill", "La Tradicional", "Tatsumi")
  empleados <- c("Karla De La Rosa", "Juan vazquez", "Ricardo Cazares", "Elizabeth Estrella", "Aurora Idolina", "Alejandra Ramirez") 
  emails <- list(c("karla.delarosa@na.denso.com"), c("jvazkez_9093@hotmail.com", "oscarop1912950@gmail.com"), c("rcazares@ladrillera.mx", "ybustos@ladrillera.mx", "acastillo@ladrillera.mx"),c("eli_estrella@muebleskrill.com", "krill.rh2@gmail.com"), c("recursosh@tostadasdelicias.com"), c("laura-martinez@tdm.mitsuba-gr.com", "alejandra-rmz@tdm.mitsuba-gr.com"))
  
  # Create a list to associate empresas, empleados, and emails
  company_data <- list()
  
  # Loop through each empresa and associate empleados and emails
  for (i in seq_along(empresas)) {
    company <- empresas[i]
    employee <- empleados[i]
    email <- emails[[i]]
    
    company_data[[company]] <- list(Empleado = employee, Email = email)
  }
  
  # Load the lubridate package
  library(lubridate)
  
  # Get today's date
  today <- Sys.Date()
  
  # Calculate the nearest Friday
  days_to_friday <- ifelse(wday(today) < 6, 6 - wday(today), 13 - wday(today))
  nearest_friday <- today + days(days_to_friday)
  
  # Print the nearest Friday
  formatted_friday <- format(nearest_friday, "%A, %d de %B %Y")
  
  htmlbody <- paste0(
    "<html>
    <head>
      <style>
        body {
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        line-height: 1.6;
        margin: 20px;
        background-color: #f4f4f4;
        display: flex;
        flex-direction: column;
        min-height: 100vh; 
        }
        h1 {
          color: #007bff;
        font-size: 35px;
        margin-bottom: 5px;
        }
        p {
          color: #555;
        font-size: 15px;
        margin-bottom: 10px;
        }
        .message {
        background-color: #f4f4f4;
        padding: 20px;
        border-radius: 8px;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        } 
        footer {
        margin-top: auto; 
        padding: 15px;
        background-color: #f9f9f9;
        border-radius: 5px;
        font-size: 12px; /* Adjust the font size for the footer */
        }
      footer p {
        font-size: 10px !important; 
        margin: 5px 0; 
      }
      address {
        margin-left: auto; /* Push address to the right */
        text-align: left; /* Align text within address to the right */
      }
      </style>
    </head>
    <body>
      <h1>Listado de Descuentos ", empresa, "</h1> <br>
      <img src='https://flexicredit.mx/wp-content/uploads/2022/06/Flexicredit-02-02.png' alt='Logo de la empresa' width='100'>
      <div class='message'>
        <p>Buenos días ", company_data[[empresa]]$Empleado , ",</p>
        <p>Espero te encuentres muy bien. Te compartimos el listado con los nuevos descuentos a realizar a partir de este <em><strong> ", formatted_friday ,". </strong></em>
        <br>Adjunto encontraras el archivo con el detalle de la información. Quedamos al pendiente para cualquier duda o aclaración.</p>
        <p>Saludos y excelente inicio de semana!</p>
      
     <footer>
     <br>     
     <hr>
      <p style='font-size: 14px'><strong>Simon Santos Luis</strong><br>
      Finanzas<br>
      +52(81)50009060<br>
      finanzas3@flexicredit.mx</p>
    <hr>
    <br>
    
      <address>
        <p style='font-size: 13px'>250 Calzada San Pedro. 3er Piso.<br>
        Miravalle<br>
        Monterrey, Nuevo Leon. Mexico 64660</p>
      </address>
    </footer>
    </div>
    </body>
  </html>"
  ) 
  
  
  # Open Outlook
  Outlook <- COMCreate("Outlook.Application")
  
  # Create a new message
  Email = Outlook$CreateItem(0)
  
  # Set the recipient, subject, and body
  Email[["to"]] = paste(company_data[[empresa]]$Email, collapse = ";") 
  Email[["cc"]] = ""
  Email[["bcc"]] = "dir.general@flexicredit.mx;finanzas2@flexicredit.mx"
  Email[["subject"]] = paste0("Listado de Descuentos ", empresa)
  Email[["htmlbody"]] = htmlbody
  
  # Attach the file
  Email[["attachments"]]$Add(paste0("C:/Users/SimonSantos/Documents/Integration/DATA/FINANCIEROS/",empresa,"/ENTREGABLE/FlexiCredit - Listado de Descuentos ", empresa, " ", nearest_monday, ".xlsx"))
  
  # Send the message
  trueorfalse <- Email$Send()
  trueorfalse
  
  if (trueorfalse) {
    move_directory <- paste0(Documents_path, "\\Integration\\DATA\\FINANCIEROS\\", empresa,"\\ENTREGABLE")
    sent_directory <- paste0(Documents_path, "\\Integration\\DATA\\FINANCIEROS\\", empresa,"\\ENVIADOS")
    
    move_files <- list.files(move_directory, full.names = TRUE, recursive = FALSE)
    move_files <- move_files[file.info(move_files)$isdir == FALSE]
    move_files <- move_files[order(file.info(move_files)$mtime, decreasing = TRUE)]
    new_file_name <- paste0("FlexiCredit - Listado de Descuentos ", empresa, " ", nearest_monday, ".xlsx")
    
    destination_folder <- sent_directory
    if (length(empresas_files) > 0) {
      # Get the most recent file (the first one in the sorted list)
      most_recent_file <- move_files[1]
      
      # Define the destination path for the most recent file
      destination_path <- file.path(destination_folder, new_file_name)
      
      # Move the most recent file to the destination folder
      file.rename(most_recent_file, destination_path)
      cat("Moved file:", most_recent_file, "\nTo destination:", destination_path, "\n")
    } else {
      cat("No files to move in the Downloads folder.\n")
    }
  }
  
}