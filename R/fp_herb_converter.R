#' Fill empty herbarium spreadsheet from forestplots-formatted data
#'
#' @author Giulia Ottino & Domingos Cardoso
#'
#' @description Function to process data collected in the \href{https://forestplots.net/}{Forestplots format}
#' and convert it into a herbarium format.
#' The default herbarium format is \href{https://rb.jbrj.gov.br/v2/validarplanilha_externo.php}{JABOT},
#' but it also supports conversion to the \href{https://herbaria.plants.ox.ac.uk/bol/brahms/support/conifers}{Brahms}
#' and custom herbarium formats, where the user can specify the column names in
#' column_map list. The function performs the following steps: (i) Reads the input
#' forestplots field sheet template and herbarium datasets; (ii) Ensures all
#' required metadata (e.g., majorarea, minorarea, lat, long) is correctly
#' provided;(iii) Processes taxonomic, spatial, and additional metadata to fill
#' the herbarium format sheet; (iv) Generates a new xlsx file with the converted
#' data, saved in a directory named after the current date.
#' This function automatically converts field codes from the Rainfor protocol
#' (e.g., tree condition codes and light exposure codes) into their full
#' descriptive meanings, and incorporates them into the final herbarium notes.
#' For example, the code "A" is converted to "alive normal" (or "árvore viva
#' normal" in Portuguese). These translations are integrated into a standardized
#' descriptive text alongside other information extracted directly from the
#' forestplots sheet, such as DBH, field notes, and subplot coordinates.
#' For instance, a final note in English might read:
#'
#' "Tree, 24.83cm DBH, with peeling bark, small buttress roots, bark with tiny
#'  reddish plates, crown completely exposed to vertical and lateral light in a
#'  45 degree curve. Individual #3 in the subplot 1, x = 1.5m, y = 2.5m."
#'
#' The phrases "with peeling bark" and "crown completely exposed to vertical
#' and lateral light..." are automatically generated from the Rainfor codes,
#' while phrases such as "small buttress roots" and "bark with  reddish plates"
#' come from the original field notes recorded in the Forestplots sheet.The
#' conversion of Rainfor codes and the construction of this descriptive text
#' are supported in three languages: English, Portuguese, and Spanish, depending
#' on the value set in the `language` argument.

#' @usage
#' fp_herb_converter(forestplots_file_path = NULL,
#'                   herb_file_path = NULL,
#'                   language = "en",
#'                   herbarium_format = "jabot",
#'                   country = "Brazil",
#'                   majorarea = NULL,
#'                   minorarea = NULL,
#'                   protectedarea = NULL,
#'                   locnotes = NULL,
#'                   project = NULL,
#'                   collector = NULL,
#'                   addcoll = NULL,
#'                   lat = NULL,
#'                   long = NULL,
#'                   alt = NULL,
#'                   dir = "Results_rainfor_herb",
#'                   filename = "rainfor_to_herb")
#'
#' @param forestplots_file_path File path to the forestplots dataset.
#'
#' @param herb_file_path File path of the herbarium empty template dataset.
#'
#' @param language Language for output metadata. Default is "en", options are
#' "en", "pt" and "es".
#'
#' @param herbarium_format Format of the herbarium dataset. Default is "jabot",
#' additional options is "brahms", "custom" herbarium format, where column names
#' must be specified by the user.
#'
#' @param country Country where the data was collected. Default is "Brazil".
#'
#' @param majorarea State where the data was collected.
#'
#' @param minorarea Municipality where the data was collected.
#'
#' @param protectedarea Name of the protected area (if applicable).
#'
#' @param locnotes Notes of the locality where the data was collected.
#'
#' @param project Project name or description (if applicable).
#'
#' @param collector Main collector name. If \code{'NULL'}, extract the 'Team'
#' information of forestplots dataset. the first member of the team will be
#' considered as the main collector.
#'
#' @param addcoll Character. Additional collector(s) name(s).If \code{'NULL'},
#' extract the additional members of the 'Team' in forestplots dataset.
#'
#' @param lat Latitude in decimal degrees of the collection site.
#'
#' @param long Longitude in decimal degrees of the collection site.
#'
#' @param alt Altitude in meters of the collection site. If \code{'NULL'},
#' extract the altitude from geographic coordinates dataset.
#'
#' @param add_census_notes If \code{'TRUE'}, census notes will be included in the
#' herbarium  sheet's notes. If \code{'FALSE'}, census notes will be omitted.
#'
#' @param dir Pathway to the computer's directory, where the file will be saved.
#' The default is to create a directory named **Results_filled_herbarium_sheet**
#' and the search results will be saved within a subfolder named after the current
#' date.
#'
#' @param filename Name of the output file. Default is **forestplot_to_herbarium_sheet**.
#'
#' @return A data frame in herbarium format with the processed data from
#' forestplots dataset.
#'
#'
#' @examples
#' fp_herb_converter <- function(fp_file_path = NULL,
#'                               herb_file_path = NULL,
#'                               language = "en",
#'                               herbarium_format = "jabot",
#'                               country = "Brazil",
#'                               majorarea = NULL,
#'                               minorarea = NULL,
#'                               protectedarea = NULL,
#'                               locnotes = NULL,
#'                               project = NULL,
#'                               collector = NULL,
#'                               addcoll = NULL,
#'                               lat = NULL,
#'                               long = NULL,
#'                               alt = NULL,
#'                               add_census_notes = TRUE,
#'                               dir = "Results_filled_herbarium_sheet",
#'                               filename =  "forestplot_to_herbarium_sheet")
#'
#' @importFrom readxl read_excel
#' @importFrom openxlsx write.xlsx
#' @importFrom geodata elevation_global
#' @importFrom terra vect extract
#' @importFrom stringi stri_trans_general
#'
#' @export
#'

fp_herb_converter <- function(fp_file_path = NULL,
                              herb_file_path = NULL,
                              language = "en",
                              herbarium_format = "jabot",
                              country = "Brazil",
                              majorarea = NULL,
                              minorarea = NULL,
                              protectedarea = NULL,
                              locnotes = NULL,
                              project = NULL,
                              collector = NULL,
                              addcoll = NULL,
                              lat = NULL,
                              long = NULL,
                              alt = NULL,
                              add_census_notes = TRUE,
                              dir = "Results_filled_herbarium_sheet",
                              filename =  "forestplot_to_herbarium_sheet") {


  # Apply defensive functions. to verify if required arguments are present

  # dir check
  dir <- .arg_check_dir(dir)

  # majorarea check
  majorarea <- .arg_check_majorarea(majorarea, country)

  # minorarea check
  .arg_check_minorarea(minorarea)

  # lat check
  lat <- .arg_check_lat(lat)

  # long check
  long <- .arg_check_long(long)

  # alt check
  alt <- .arg_check_alt(alt)

  # check if a point falls within a given administrative division
  .check_point_in_admin(lat = lat,
                        long = long,
                        majorarea = majorarea,
                        country = country)

  # Creating the directory to save the file based on the current date
  # Make folder name to save search results  e.g.: 05Dec2024
  foldername <- paste0(dir, "/", format(Sys.time(), "%d%b%Y"))

  # Create the main 'Results_filled_herbarium_sheet' folder and subfolders with
  # the date,  if they don't exist
  if (!dir.exists(dir)) dir.create(dir)
  if (!dir.exists(foldername)) dir.create(foldername)

  # Create the file path name for the spreadsheet to be saved in .xlsx format
  fullname <- paste0(foldername, "/", filename, ".xlsx")

  #### Load Data ####
  # Load forestplot spreadsheet (adjust file name as needed)
  fp_sheet <- suppressMessages({readxl::read_excel(fp_file_path, sheet = 1)
  })

  # Load the empty herbarium spreadsheet (adjust file name as needed)
  herb_sheet <- readxl::read_excel(herb_file_path)

  # Extract information for 'collector' and 'addcoll' before modifying the dataset
  if (is.null(collector) | is.null(addcoll)) {

    team_col <- which(grepl("^Team[:]", names(fp_sheet)))

    team_info <- strsplit(as.character(names(fp_sheet)[team_col]), ",")

    if (is.null(collector)) {
      collector <- sapply(team_info, function(x) trimws(strsplit(x[1], ":")[[1]][2]))
    }

    if (is.null(addcoll)) {
      addcoll <- sapply(team_info, function(x) {
        if (length(x) > 1) {
          trimws(paste(x[-1], collapse = ","))
        } else {
          NA
        }
      })
    }
  }

  # Check if alt is NULL and use the user-provided value (if available)
  if (is.null(alt)) {
    # If alt is NULL, proceed with elevation extraction logic
    # Download global elevation data at 2.5 arc-second resolution (~75m at equator)
    elev <- geodata::elevation_global(res = 2.5, path = tempdir())

    # Create spatial points object from coordinates
    points <- data.frame(lon = long, lat = lat)
    points_sp <- terra::vect(points,
                             geom = c("lon", "lat"),
                             crs = "EPSG:4326")  # WGS84 coordinate system

    # Extract elevation values for the points
    elev_values <- terra::extract(elev, points_sp)

    # Get elevation values from the second column (first column is ID)
    alt <- elev_values[, 2]

    # Convert to numeric
    alt <- as.numeric(alt)

    # Remove the elevation data file from the disk after extraction
    unlink(elev)
  }

  # Extract the date from the 11th column name (format: "Date: dd/mm/yy")
  date_col <- grep("^Date: \\d{2}/\\d{2}/\\d{2}$", names(fp_sheet))
  date_col <- grep("^Date: \\d{2}/\\d{2}/\\d{4}$", names(fp_sheet))
  date_info <- strsplit(names(fp_sheet)[date_col], "Date: ")[[1]][2]

  # Split the date into day, month, and year
  date_parts <- strsplit(date_info, "/")[[1]]
  colldd <- date_parts[1]
  collmm <- date_parts[2]
  collyy <- date_parts[3]

  # Remove the first row as it's not relevant
  colnames(fp_sheet) <- fp_sheet[1, ]
  fp_sheet <- fp_sheet[-1, ]

  # Edit original info in the forestplot sheet
  # Filtrar apenas as linhas com alguma informação na coluna 'Collected'
  fp_sheet <- fp_sheet[!is.na(fp_sheet$Collected) & fp_sheet$Collected != "", ]
  fp_sheet$D <- as.numeric(fp_sheet$D)/10
  fp_sheet$'Census Notes' <- gsub("[.]$", "", fp_sheet$'Census Notes')
  fp_sheet$'Census Notes' <- .convert_to_lowercase(string=fp_sheet$'Census Notes')

  # Prepare the herb_sheet DataFrame structure
  n <- length(unique(na.omit(fp_sheet$Voucher)))
  mtemp <- as.data.frame(matrix(NA, n, ncol(herb_sheet)))
  names(mtemp) <- names(herb_sheet)
  herb_sheet <- mtemp

  if (herbarium_format == "jabot"){
    # Fill the herb_sheet columns with necessary information
    herb_sheet$family <- ifelse(grepl("(i|I)ndet", fp_sheet$Family), NA, fp_sheet$Family)
    herb_sheet$number <- gsub("[^0-9]", "", fp_sheet$Voucher)
    herb_sheet$collector <- collector
    herb_sheet$addcoll <- addcoll
    herb_sheet$colldd <- colldd
    herb_sheet$collmm <- collmm
    herb_sheet$collyy <- collyy
    herb_sheet$country <- country
    herb_sheet$majorarea <- majorarea
    herb_sheet$minorarea <- minorarea
    herb_sheet$uc <- protectedarea
    herb_sheet$projeto <- project
    herb_sheet$locnotes <- locnotes
    herb_sheet$latitude <- lat
    herb_sheet$longitude <- long
    herb_sheet$altprof <- alt
    notes <- "notes"

    # Update genus and sp1 based on 'Original determination' column
    for (i in 1:nrow(fp_sheet)) {
      # Split the 'Original determination' into components (gênero e espécie)
      determination <- strsplit(as.character(fp_sheet$`Original determination`)[i], " ")[[1]]

      # Fill genus (first word), and replace "indet" with ""
      herb_sheet$genus[i] <- ifelse(length(determination) > 0,
                                    gsub("(i|I)ndet", NA, determination[1]), NA)

      # Fill sp1 (second word if available), and replace "indet" with ""
      herb_sheet$sp1[i] <- ifelse(length(determination) > 1,
                                  gsub("indet", NA, determination[2]), NA)

      # Ensure sp1 is different from genus
      herb_sheet$sp1[i] <- ifelse(is.na(herb_sheet$sp1[i]) | herb_sheet$sp1[i] == herb_sheet$genus[i],
                                  NA, herb_sheet$sp1[i])
    }


    # Build the final notes entry
    herb_sheet <- .build_notes(fp_sheet, herb_sheet, i, notes = notes,
                               language = language,
                               add_census_notes = add_census_notes)

    # Saving the file in the correct directory
    message(paste0("Writing spreadsheet '", filename, ".xlsx' within '",
                   dir, "' on disk."))

    openxlsx::write.xlsx(herb_sheet, fullname)

    # Return spreadsheet to R environment
    return(herb_sheet)
  }
  else if (herbarium_format == "brahms") {
    # Fill the herb_sheet columns with necessary information
    herb_sheet$FamilyName <- ifelse(grepl("(i|I)ndet", fp_sheet$Family),
                                    NA, fp_sheet$Family)
    herb_sheet$FieldNumber <- gsub("[^0-9]", "", fp_sheet$Voucher)
    herb_sheet$Collectors <- collector
    herb_sheet$AdditionalCollectors <- addcoll
    herb_sheet$CollectionDay <- colldd
    herb_sheet$CollectionMonth <- collmm
    herb_sheet$CollectionYear <- collyy
    herb_sheet$CountryName <- country
    herb_sheet$MajorAdminName <- majorarea
    herb_sheet$MinorAdminName<- minorarea
    herb_sheet$LocalityName <- protectedarea
    herb_sheet$LocalityNotes <- locnotes
    herb_sheet$Latitude <- lat
    herb_sheet$Longitude <- long
    herb_sheet$Elevation <- alt
    notes <- "DescriptionText"

    # Update GenusName and SpeciesName based on 'Original determination' column
    for (i in 1:nrow(fp_sheet)) {
      # Split the 'Original determination' into components (gênero e espécie)
      determination <- strsplit(as.character(fp_sheet$`Original determination`)[i], " ")[[1]]
      # Fill GenusName (first word), and replace "indet" with ""
      herb_sheet$GenusName[i] <- ifelse(length(determination) > 0,
                                        gsub("(i|I)ndet", NA, determination[i]), NA)
      # Fill GenusName (first word), and replace "indet" with ""
      herb_sheet$GenusName[i] <- ifelse(length(determination) > 0,
                                        gsub("(i|I)ndet", NA, determination[1]), NA)

      # Fill SpeciesName (second word if available), and replace "indet" with ""
      herb_sheet$SpeciesName[i] <- ifelse(length(determination) > 1,
                                          gsub("indet", NA, determination[2]), NA)

      # Ensure SpeciesName is different from GenusName
      herb_sheet$SpeciesName[i] <- ifelse(is.na(herb_sheet$SpeciesName[i]) | herb_sheet$SpeciesName[i] == herb_sheet$GenusName[i],
                                          NA, herb_sheet$SpeciesName[i])
    }

    # Build the final notes entry
    herb_sheet <- .build_notes(fp_sheet, herb_sheet, i, notes = notes,
                               language = language,
                               add_census_notes = add_census_notes)

    # Saving the file in the correct directory
    message(paste0("Writing spreadsheet '", filename, ".xlsx' within '",
                   dir, "' on disk."))

    openxlsx::write.xlsx(herb_sheet, fullname)

    # Return spreadsheet to R environment
    return(herb_sheet)
  }
  else if (herbarium_format == "custom") {
    # Fill the herb_sheet columns with necessary information
    herb_sheet[[column_map$family]] <- ifelse(grepl("(i|I)ndet", fp_sheet$Family),
                                              NA, fp_sheet$Family)
    herb_sheet[[column_map$number]] <- gsub("[^0-9]", "", fp_sheet$Voucher)
    herb_sheet[[column_map$collector]] <- collector
    herb_sheet[[column_map$addcoll]] <- addcoll
    herb_sheet[[column_map$colldd]] <- colldd
    herb_sheet[[column_map$collmm]] <- collmm
    herb_sheet[[column_map$collyy]] <- collyy
    herb_sheet[[column_map$country]] <- country
    herb_sheet[[column_map$majorarea]] <- majorarea
    herb_sheet[[column_map$minorarea]]<- minorarea
    herb_sheet[[column_map$protectedarea]] <- protectedarea
    herb_sheet[[column_map$locnotes]] <- locnotes
    herb_sheet[[column_map$project]] <- project
    herb_sheet[[column_map$lat]] <- lat
    herb_sheet[[column_map$long]] <- long
    herb_sheet[[column_map$alt]] <- alt
    notes <- column_map[["notes"]]

    # Update GenusName and SpeciesName based on 'Original determination' column
    for (i in 1:nrow(fp_sheet)) {
      # Split the 'Original determination' into components (gênero e espécie)
      determination <- strsplit(as.character(fp_sheet$`Original determination`)[i], " ")[[1]]
      # Fill GenusName (first word), and replace "indet" with ""
      herb_sheet[[column_map$genus]][i] <- ifelse(length(determination) > 0,
                                                  gsub("(i|I)ndet", NA, determination[i]), NA)
      # Fill GenusName (first word), and replace "indet" with ""
      herb_sheet[[column_map$genus]][i] <- ifelse(length(determination) > 0,
                                                  gsub("(i|I)ndet", NA, determination[1]), NA)

      # Fill SpeciesName (second word if available), and replace "indet" with ""
      herb_sheet[[column_map$sp1]][i] <- ifelse(length(determination) > 1,
                                                gsub("indet", NA, determination[2]), NA)

      # Ensure SpeciesName is different from GenusName
      herb_sheet[[column_map$sp1]][i] <- ifelse(is.na(herb_sheet[[column_map$sp1]][i]) | herb_sheet[[column_map$sp1]][i] == herb_sheet[[column_map$genus]][i],
                                                NA, herb_sheet[[column_map$sp1]][i])
    }

    # Build the final notes entry
    herb_sheet <- .build_notes(fp_sheet, herb_sheet, i, notes = notes,
                               language = language,
                               add_census_notes = add_census_notes)

    # Saving the file in the correct directory
    message(paste0("Writing spreadsheet '", filename, ".xlsx' within '",
                   dir, "' on disk."))

    openxlsx::write.xlsx(herb_sheet, fullname)

    # Return spreadsheet to R environment
    return(herb_sheet)
  }
  else if (herbarium_format == "custom") {
    # Fill the herb_sheet columns with necessary information
    herb_sheet[[column_map$family]] <- ifelse(grepl("(i|I)ndet", fp_sheet$Family),
                                              NA, fp_sheet$Family)
    herb_sheet[[column_map$number]] <- gsub("[^0-9]", "", fp_sheet$Voucher)
    herb_sheet[[column_map$collector]] <- collector
    herb_sheet[[column_map$addcoll]] <- addcoll
    herb_sheet[[column_map$colldd]] <- colldd
    herb_sheet[[column_map$collmm]] <- collmm
    herb_sheet[[column_map$collyy]] <- collyy
    herb_sheet[[column_map$country]] <- country
    herb_sheet[[column_map$majorarea]] <- majorarea
    herb_sheet[[column_map$minorarea]]<- minorarea
    herb_sheet[[column_map$protectedarea]] <- protectedarea
    herb_sheet[[column_map$locnotes]] <- locnotes
    herb_sheet[[column_map$project]] <- project
    herb_sheet[[column_map$lat]] <- lat
    herb_sheet[[column_map$long]] <- long
    herb_sheet[[column_map$alt]] <- alt
    notes <- column_map[["notes"]]

    # Update GenusName and SpeciesName based on 'Original determination' column
    for (i in 1:nrow(fp_sheet)) {
      # Split the 'Original determination' into components (gênero e espécie)
      determination <- strsplit(as.character(fp_sheet$`Original determination`)[i], " ")[[1]]
      # Fill GenusName (first word), and replace "indet" with ""
      herb_sheet[[column_map$genus]][i] <- ifelse(length(determination) > 0,
                                                  gsub("(i|I)ndet", NA, determination[i]), NA)
      # Fill GenusName (first word), and replace "indet" with ""
      herb_sheet[[column_map$genus]][i] <- ifelse(length(determination) > 0,
                                                  gsub("(i|I)ndet", NA, determination[1]), NA)

      # Fill SpeciesName (second word if available), and replace "indet" with ""
      herb_sheet[[column_map$sp1]][i] <- ifelse(length(determination) > 1,
                                                gsub("indet", NA, determination[2]), NA)

      # Ensure SpeciesName is different from GenusName
      herb_sheet[[column_map$sp1]][i] <- ifelse(is.na(herb_sheet[[column_map$sp1]][i]) | herb_sheet[[column_map$sp1]][i] == herb_sheet[[column_map$genus]][i],
                                                NA, herb_sheet[[column_map$sp1]][i])
    }

    # Build the final notes entry
    herb_sheet <- .build_notes(fp_sheet, herb_sheet, i, notes = notes,
                               language = language,
                               add_census_notes = add_census_notes)

    # Saving the file in the correct directory
    message(paste0("Writing spreadsheet '", filename, ".xlsx' within '",
                   dir, "' on disk."))

    openxlsx::write.xlsx(herb_sheet, fullname)
    # Return spreadsheet to R environment
    return(herb_sheet)
  }
}
#_______________________________________________________________________________
# Auxiliary functions

# Function to retrieve descriptions based on flags
.get_descriptions <- function(language = language) {
  descriptions <- list(
    if (language == "pt"){
      return(data.frame(
        letters = c("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k",
                    "l", "m", "o", "p", "q", "s", "w", "x", "y", "z"),
        description = c(
          "viva normal",
          "com fuste/copa quebrado/a e com rebroto, ou pelo menos com floema/xilema vivo",
          "inclinada ≥10%",
          "caída (por ex. no chão)",
          "acanalada e/ou fenestrada",
          "com caule oco",
          "com caule podre e/ou com presença de fungos de prateleira (ou de suporte)",
          "com múltiplos fustes",
          "sem folhas/com poucas folhas",
          "com caule queimado",
          "com caule quebrado <1,3m",
          "com liana ≥ 10cm de diâmetro no caule ou na copa",
          "coberta por lianas",
          "que sofreu danos causados por raio",
          "cortada",
          "com casca solta/descamando",
          "com estrangulador",
          "com uma ferida e/ou câmbio exposto",
          "danificada por elefantes",
          "danificada por cupins",
          "com baixa produtividade (quase morta)"
        ))
      )},
    if (language == "en") {
      return(data.frame(
        letters = c("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k",
                    "l", "m", "o", "p", "q", "s", "w", "x", "y", "z"),
        description = c(
          "alive normal",
          "with broken stem/top & resprouting, or at least live phloem/xylem",
          "leaning by ≥10%",
          "fallen (e.g. on ground)",
          "fluted or/fenestrated",
          "with hollow",
          "with rotten and or presence of bracket fungus",
          "as multiple stemmed individual",
          "no leaves, few leaves",
          "burnt",
          "snapped < 1.3m",
          "with liana ≥10cm diameter on stem or in canopy",
          "covered by lianas",
          "with lightning damage",
          "cut",
          "with peeling bark",
          "with a strangler",
          "with wound and/or cambium exposed",
          "with elephant damage",
          "with termite damage",
          "declining productivity (nearing death)"
        ))
      )},
    if (language == "es") {
      return(data.frame(
        letters = c("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k",
                    "l", "m", "o", "p", "q", "s", "w", "x", "y", "z"),
        description = c(
          "vivo normal",
          "con tallo partido y con rebrotes, o por lo menos hay floema/xilema vivo",
          "inclinado ≥10%",
          "caído (por ejemplo: sobre el suelo)",
          "acanalado y/o fenestrado",
          "con tallo hueco",
          "con tallo podrido",
          "con tallos múltiples",
          "sin o con pocas hojas",
          "con tallo quemado",
          "con tallo partido <1.3m",
          "con liana ≥ 10cm de diámetro en el tronco o en la copa del árbol",
          "cubierto por lianas",
          "dañado por un rayo",
          "cortado",
          "con corteza pelada",
          "con un estrangulador",
          "con una herida o cambio expuesto",
          "dañado por elefantes",
          "dañado por termitas",
          "con baja productividad (casi muerto)"
        ))
      )
    }
  )
  return(descriptions[[language]])
}

# Function to retrieve descriptions based on luminosity
.get_luminosity <- function(language = language) {
  luminosity <- list(
    if (language == "pt"){
      return(data.frame(
        code = c("5", "4", "3b", "3a", "2c", "2b", "2a", "1"),
        description = c(
          "copa totalmente exposta à luz vertical e lateral numa curva de 45 graus",
          "copa totalmente exposta à luz vertical, mas a luz lateral está bloqueada por alguns ou todos os cones invertidos de 90 graus que englobam a copa",
          "sob muita luz vertical (>50%)",
          "sob pouca luz vertical (menos de 50% da copa está exposta à luz vertical)",
          "sob muita luz lateral",
          "sob média luz lateral",
          "sob pouca luz lateral",
          "sem luz direta"
        ))
      )
    },
    if (language == "en"){
      return(data.frame(
        code = c("5", "4", "3b", "3a", "2c", "2b", "2a", "1"),
        description = c(
          "crown completely exposed to vertical and lateral light in a 45 degree curve",
          "crown completely exposed to vertical light, but lateral light blocked",
          "under high vertical illumination (>50%)",
          "under some vertical light (<50%)",
          "under high lateral light",
          "under medium lateral light",
          "under low lateral light",
          "not under direct light")))
    },
    if (language == "es"){
      return(data.frame(
        code = c("5", "4", "3b", "3a", "2c", "2b", "2a", "1"),
        description = c(
          "copa completamente expuesta a luz vertical y horizontal en una curva de 45 grados",
          "copa completamente expuesta a luz vertical, pero la luz lateral está bloqueada",
          "sob iluminación vertical alta (más que 50%)",
          "sob poca luz vertical (menos que 50%)",
          "sob luz lateral alta",
          "sob luz lateral media",
          "sob luz lateral baja",
          "no hay luz directa")))
    }
  )
  return(luminosity[[language]])
}

# Function to convert the first word and words after commas to lowercase
.convert_to_lowercase <- function(string) {
  for (l in 1:length(string)) {
    if (!is.na(string[l])) {

      # Split the text by spaces to get individual words
      string_temp <- strsplit(string[l], " ")[[1]]

      # Convert all words to lowercase
      string_temp <- tolower(string_temp)

      # Rejoin the words into a single string
      modified_string <- paste(string_temp, collapse = " ")

      # Replace any period (.) with a comma (,)
      modified_string <- gsub("\\.", ",", modified_string)

      # Assign the modified string back
      string[l] <- modified_string
    } else {
      # Replace NA with an empty string
      string[l] <- ""
    }
  }

  # Return the modified string
  return(string)
}

# Function to build herbarium notes
.build_notes <- function(fp_sheet, herb_sheet, i, notes = notes,
                         language = "en",
                         add_census_notes = TRUE) {

  # Loop through each row in the fp_sheet data
  for (i in 1:nrow(fp_sheet)) {

    # Process Flag1 descriptions
    flag_letters <- strsplit(tolower(fp_sheet$Flag1[i]), split = "")[[1]]
    converted_descriptions <- .get_descriptions(language = language)$description[match(flag_letters,                                                                                   .get_descriptions(language = language)$letters)]
    description_flag <- paste(converted_descriptions, collapse = ", ")

    # Process LI (luminosity) descriptions
    li_code <- as.character(fp_sheet$LI[i])
    description_li <- .get_luminosity(language = language)$description[match(li_code,
                                                                             .get_luminosity(language = language)$code)]

    # Constructing the herbarium notes based on the language
    # English version
    if (language == "en") {
      if (add_census_notes) {
        herb_sheet[[notes]][i] <- paste0("Tree, ",
                                         ifelse(!is.na(fp_sheet$D[i]) & fp_sheet$D[i] != "",
                                                paste0(fp_sheet$D[i], "cm DBH, "), ""),
                                         description_flag, ", ",
                                         ifelse(!is.na(fp_sheet$'Census Notes'[i]) & fp_sheet$'Census Notes'[i] != "",
                                                paste0(fp_sheet$'Census Notes'[i], ", "), ""),
                                         description_li, ". ",
                                         ifelse(!is.na(fp_sheet$'New Tag No'[i]) & fp_sheet$'New Tag No'[i] != "",
                                                paste0("Individual #", fp_sheet$'New Tag No'[i]), ""),
                                         ifelse(!is.na(fp_sheet$T1[i]) & fp_sheet$T1[i] != "",
                                                paste0(" in the subplot ", fp_sheet$T1, ", x = ", fp_sheet$X, "m, y = ", fp_sheet$Y, "m."), ""))

      } else {
        herb_sheet[[notes]][i] <- paste0("Tree, ",
                                         ifelse(!is.na(fp_sheet$D[i]) & fp_sheet$D[i] != "",
                                                paste0(fp_sheet$D[i], "cm DBH, "), ""),
                                         description_flag, ", ",
                                         description_li, ". ",
                                         ifelse(!is.na(fp_sheet$'New Tag No'[i]) & fp_sheet$'New Tag No'[i] != "",
                                                paste0("Individual #", fp_sheet$'New Tag No'[i]), ""),
                                         ifelse(!is.na(fp_sheet$T1[i]) & fp_sheet$T1[i] != "",
                                                paste0(" in the subplot ", fp_sheet$T1, ", x = ", fp_sheet$X, "m, y = ", fp_sheet$Y, "m."), ""))
      }

    } else if (language == "pt") {

      # Portuguese version
      if (add_census_notes) {
        herb_sheet[[notes]][i] <- paste0("Árvore, ",
                                         ifelse(!is.na(fp_sheet$D[i]) & fp_sheet$D[i] != "",
                                                paste0(fp_sheet$D[i], "cm de diâmetro à altura do peito (DAP), "), ""),
                                         description_flag, ", ",
                                         ifelse(!is.na(fp_sheet$'Census Notes'[i]) & fp_sheet$'Census Notes'[i] != "",
                                                paste0(fp_sheet$'Census Notes'[i], ", "), ""),
                                         description_li, ". ",
                                         ifelse(!is.na(fp_sheet$'New Tag No'[i]) & fp_sheet$'New Tag No'[i] != "",
                                                paste0("Indivíduo #", fp_sheet$'New Tag No'[i]), ""),
                                         ifelse(!is.na(fp_sheet$T1[i]) & fp_sheet$T1[i] != "",
                                                paste0(" na subparcela ", fp_sheet$T1, ", x = ", fp_sheet$X, "m, y = ", fp_sheet$Y, "m."), ""))

      } else {
        herb_sheet[[notes]][i] <- paste0("Árvore, ",
                                         ifelse(!is.na(fp_sheet$D[i]) & fp_sheet$D[i] != "",
                                                paste0(fp_sheet$D[i], "cm de diâmetro à altura do peito (DAP), "), ""),
                                         description_flag, ", ",
                                         description_li, ". ",
                                         ifelse(!is.na(fp_sheet$'New Tag No'[i]) & fp_sheet$'New Tag No'[i] != "",
                                                paste0("Indivíduo #", fp_sheet$'New Tag No'[i]), ""),
                                         ifelse(!is.na(fp_sheet$T1[i]) & fp_sheet$T1[i] != "",
                                                paste0(" na subparcela ", fp_sheet$T1, ", x = ", fp_sheet$X, "m, y = ", fp_sheet$Y, "m."), ""))
      }

    } else if (language == "es") {

      # Spanish version
      if (add_census_notes) {
        herb_sheet[[notes]][i] <- paste0("Árbol, ",
                                         ifelse(!is.na(fp_sheet$D[i]) & fp_sheet$D[i] != "",
                                                paste0(fp_sheet$D[i], "cm de diámetro a la altura del pecho (DAP), "), ""),
                                         description_flag, ", ",
                                         ifelse(!is.na(fp_sheet$'Census Notes'[i]) & fp_sheet$'Census Notes'[i] != "",
                                                paste0(fp_sheet$'Census Notes'[i], ", "), ""),
                                         description_li, ". ",
                                         ifelse(!is.na(fp_sheet$'New Tag No'[i]) & fp_sheet$'New Tag No'[i] != "",
                                                paste0("Individuo #", fp_sheet$'New Tag No'[i]), ""),
                                         ifelse(!is.na(fp_sheet$T1[i]) & fp_sheet$T1[i] != "",
                                                paste0(" en la subparcelas ", fp_sheet$T1, ", x = ", fp_sheet$X, "m, y = ", fp_sheet$Y, "m."), ""))

      } else {
        herb_sheet[[notes]][i] <- paste0("Árbol, ",
                                         ifelse(!is.na(fp_sheet$D[i]) & fp_sheet$D[i] != "",
                                                paste0(fp_sheet$D[i], "cm de diámetro a la altura del pecho (DAP), "), ""),
                                         description_flag, ", ",
                                         description_li, ". ",
                                         ifelse(!is.na(fp_sheet$'New Tag No'[i]) & fp_sheet$'New Tag No'[i] != "",
                                                paste0("Individuo #", fp_sheet$'New Tag No'[i]), ""),
                                         ifelse(!is.na(fp_sheet$T1[i]) & fp_sheet$T1[i] != "",
                                                paste0(" en la subparcelas ", fp_sheet$T1, ", x = ", fp_sheet$X, "m, y = ", fp_sheet$Y, "m."), ""))
      }
    }

  }

  return(herb_sheet)
}

