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
#' provided. When \code{input_type = "monitora"}, the function first converts the
#' MONITORA spreadsheet into a field-sheet-like data frame and infers plot-level
#' metadata (e.g., station, protected area, team, and census year) from the
#' corresponding columns in the input dataset; (iii) Processes taxonomic,
#' spatial, and additional metadata to fill the herbarium format sheet;
#' (iv) Generates a new xlsx file with the converted data, saved in a directory
#' named after the current date.
#' This function automatically converts field codes from the Rainfor protocol
#' (e.g., tree condition codes and light exposure codes) into their full
#' descriptive meanings, and incorporates them into the final herbarium notes.
#' For example, the code "A" is converted to "alive normal" (or "arvore viva
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
#' while phrases such as "small buttress roots" and "bark with reddish plates"
#' come from the original field notes recorded in the Forestplots sheet. The
#' conversion of Rainfor codes and the construction of this descriptive text
#' are supported in three languages: English, Portuguese, and Spanish, depending
#' on the value set in the `language` argument.
#'
#' @usage
#' fp_herb_converter(fp_file_path = NULL,
#'                   herb_file_path = NULL,
#'                   input_type = c("forestplots", "monitora"),
#'                   sheet = 1,
#'                   station_name = NULL,
#'                   language = "en",
#'                   herbarium_format = "jabot",
#'                   column_map = NULL,
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
#'                   add_census_notes = TRUE,
#'                   validate_locality = TRUE,
#'                   dir = "Results_herb_label",
#'                   filename = "plot_to_herb_sheet")
#'
#' @param fp_file_path File path to the input dataset. This can be either a
#' ForestPlots field sheet or a MONITORA spreadsheet, depending on the value of
#' \code{input_type}.
#'
#' @param herb_file_path File path of the herbarium empty template dataset.
#'
#' @param input_type Type of input dataset. Options are "forestplots" (default)
#' and "monitora". For \code{"monitora"}, the function converts the input table
#' to a forestplots-like structure before filling the herbarium template.
#'
#' @param sheet Sheet index or name to read from the input Excel file.
#'
#' @param station_name Optional station name or station number used to filter
#' MONITORA data before conversion. For \code{input_type = "monitora"}, this
#' value is matched against station-related columns in the input dataset.
#'
#' @param language Language for output metadata. Default is "en", options are
#' "en", "pt" and "es".
#'
#' @param herbarium_format Format of the herbarium dataset. Default is "jabot",
#' additional options is "brahms", "custom" herbarium format, where column names
#' must be specified by the user.
#'
#' @param column_map Named list used when \code{herbarium_format = "custom"}.
#' It must map the output herbarium fields to the corresponding columns in the
#' custom herbarium template.
#'
#' @param country Country where the data was collected. Default is "Brazil".
#'
#' @param majorarea State where the data was collected.
#'
#' @param minorarea Municipality where the data was collected.
#'
#' @param protectedarea Name of the protected area (if applicable). For
#' \code{input_type = "monitora"}, if \code{NULL}, the function attempts to
#' recover this value from the corresponding protected-area / plot-name columns
#' in the input dataset.
#'
#' @param locnotes Notes of the locality where the data was collected.
#'
#' @param project Project name or description (if applicable). For
#' \code{input_type = "monitora"}, if \code{NULL}, the function attempts to
#' recover this value from the corresponding station / plot-code columns in the
#' input dataset.
#'
#' @param collector Main collector name. If \code{NULL}, the function attempts
#' to extract this information from the input dataset metadata. For
#' \code{input_type = "monitora"}, this information is inferred preferentially
#' from voucher-collector columns, and falls back to team-related columns when
#' needed.
#'
#' @param addcoll Character. Additional collector(s) name(s). If \code{NULL},
#' the function attempts to extract this information from the input dataset
#' metadata. For \code{input_type = "monitora"}, this information is inferred
#' from team-related columns when available.
#'
#' @param lat Latitude in decimal degrees of the collection site.
#'
#' @param long Longitude in decimal degrees of the collection site.
#'
#' @param alt Altitude in meters of the collection site. If \code{NULL},
#' extract the altitude from geographic coordinates dataset.
#'
#' @param add_census_notes If \code{TRUE}, census notes will be included in the
#' herbarium sheet's notes. If \code{FALSE}, census notes will be omitted.
#'
#' @param validate_locality If \code{TRUE} (default), validates whether the
#' coordinates fall within the provided administrative division (country/majorarea).
#'
#' @param dir Pathway to the computer's directory, where the file will be saved.
#' The default is to create a directory named **Results_herb_label**
#' and the search results will be saved within a subfolder named after the current
#' date.
#'
#' @param filename Name of the output file. Default is **plot_to_herb_sheet**.
#'
#' @return A data frame in herbarium format with the processed data from the
#' input dataset.
#'
#' @examples
#' \dontrun{
#' fp_herb_converter(
#'   fp_file_path = "your_forestplots_file.xlsx",
#'   herb_file_path = "your_herbarium_template.xlsx",
#'   input_type = "forestplots"
#' )
#'
#' fp_herb_converter(
#'   fp_file_path = "your_monitora_file.xlsx",
#'   herb_file_path = "your_herbarium_template.xlsx",
#'   input_type = "monitora",
#'   sheet = 1,
#'   station_name = "MTQ-01"
#' )
#' }
#'
#' @importFrom readxl read_excel
#' @importFrom openxlsx write.xlsx
#' @importFrom geodata elevation_global
#' @importFrom terra vect extract
#' @importFrom stringi stri_trans_general
#' @importFrom stats na.omit
#'
#' @export
#'

fp_herb_converter <- function(fp_file_path = NULL,
                              herb_file_path = NULL,
                              input_type = c("forestplots", "monitora"),
                              sheet = 1,
                              station_name = NULL,
                              language = "en",
                              herbarium_format = "jabot",
                              column_map = NULL,
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
                              validate_locality = TRUE,
                              dir = "Results_herb_label",
                              filename =  "plot_to_herb_sheet") {

  # early file checks ----
  if (is.null(fp_file_path) || !is.character(fp_file_path) || length(fp_file_path) != 1 || !nzchar(fp_file_path)) {
    stop("Please provide a valid path in 'fp_file_path'.", call. = FALSE)
  }

  if (!file.exists(fp_file_path)) {
    stop("The provided 'fp_file_path' does not exist.", call. = FALSE)
  }

  if (is.null(herb_file_path) || !is.character(herb_file_path) || length(herb_file_path) != 1 || !nzchar(herb_file_path)) {
    stop("Please provide a valid path in 'herb_file_path'.", call. = FALSE)
  }

  if (!file.exists(herb_file_path)) {
    stop("The provided 'herb_file_path' does not exist.", call. = FALSE)
  }



  input_type <- match.arg(input_type)

    if (identical(herbarium_format, "custom") && is.null(column_map)) {
    stop(
      "When 'herbarium_format = \"custom\"', you must provide 'column_map'.",
      call. = FALSE
    )
    }

  # Apply defensive functions
  dir <- .arg_check_dir(dir)
  majorarea <- .arg_check_majorarea(majorarea, country)
  .arg_check_minorarea(minorarea)
  lat <- .arg_check_lat(lat)
  long <- .arg_check_long(long)
  alt <- .arg_check_alt(alt)

  #### Load Data ####
  if (identical(input_type, "forestplots")) {
    fp_sheet <- suppressMessages(
      readxl::read_excel(fp_file_path, sheet = sheet)
    )
  } else if (identical(input_type, "monitora")) {
    fp_sheet <- .monitora_to_field_sheet_df(
      path = fp_file_path,
      sheet = sheet,
      station_name = station_name
    )
  }

  herb_sheet <- readxl::read_excel(herb_file_path)

  # check if a point falls within a given administrative division (OPTIONAL)
  if (isTRUE(validate_locality)) {
    .check_point_in_admin(
      lat = lat,
      long = long,
      majorarea = majorarea,
      country = country
    )
  }

  # Creating the directory to save the file based on the current date
  foldername <- paste0(dir, "/", format(Sys.time(), "%d%b%Y"))

  if (!dir.exists(dir)) dir.create(dir)
  if (!dir.exists(foldername)) dir.create(foldername)

  fullname <- paste0(foldername, "/", filename, ".xlsx")

   # Extract information for 'collector' and 'addcoll' before modifying the dataset
  if (identical(input_type, "forestplots")) {

    if (is.null(collector) || is.null(addcoll)) {

      team_col <- which(grepl("^Team[:]", names(fp_sheet)))

      if ((is.null(collector) || is.null(addcoll)) && length(team_col) == 0) {
        warning("Could not infer collector metadata from ForestPlots header.")
      } else {
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
    }

  }

  else if (identical(input_type, "monitora")) {
    plot_meta <- attr(fp_sheet, "plot_meta")
    if (is.null(plot_meta)) plot_meta <- list()

    if (is.null(plot_meta$plot_name) && is.null(protectedarea)) {
      warning("Protected area metadata could not be inferred from MONITORA columns.")
    }

    if (is.null(plot_meta$plot_code) && is.null(project)) {
      warning("Project/station metadata could not be inferred from MONITORA columns.")
    }

    team_val <- if (!is.null(plot_meta$team)) plot_meta$team else NULL
    plot_name_val <- if (!is.null(plot_meta$plot_name)) plot_meta$plot_name else NULL
    plot_code_val <- if (!is.null(plot_meta$plot_code)) plot_meta$plot_code else NULL

    voucher_collector_vec <- if ("Voucher Collector" %in% names(fp_sheet)) {
      as.character(fp_sheet[["Voucher Collector"]])
    } else {
      rep(NA_character_, nrow(fp_sheet))
    }

    voucher_collector_vec[is.na(voucher_collector_vec)] <- ""
    voucher_collector_vec <- trimws(voucher_collector_vec)

    team_first <- NA_character_
    if (!is.null(team_val) && nzchar(team_val)) {
      team_first <- trimws(strsplit(team_val, ";|,")[[1]][1])
    }

    has_voucher_collector <- voucher_collector_vec != ""

    if (any(has_voucher_collector)) {
      # If voucher/coletor exists anywhere, use only that information.
      # Do NOT complete empty rows with team or argument.
      collector <- voucher_collector_vec
      collector[collector == ""] <- NA_character_

    } else if (!is.na(team_first) && nzchar(team_first)) {
      collector <- rep(team_first, nrow(fp_sheet))

    } else if (!is.null(collector) && length(collector) == 1L && nzchar(collector)) {
      collector <- rep(collector, nrow(fp_sheet))

    } else {
      collector <- rep(NA_character_, nrow(fp_sheet))
      warning("Collector could not be inferred from MONITORA voucher/coletor column, team metadata, or function argument.")
    }

    if (is.null(addcoll)) {
      if (!is.null(team_val) && nzchar(team_val)) {
        team_parts <- trimws(strsplit(team_val, ";|,")[[1]])
        addcoll <- if (length(team_parts) > 1) paste(team_parts[-1], collapse = ", ") else NA_character_
      } else {
        addcoll <- NA_character_
      }
    }

    if (is.null(protectedarea) && !is.null(plot_name_val) && nzchar(plot_name_val)) {
      protectedarea <- plot_name_val
    }

    if (is.null(project) && !is.null(plot_code_val) && nzchar(plot_code_val)) {
      project <- plot_code_val
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

  if (identical(input_type, "forestplots")) {
    date_col <- grep("^Date: \\d{2}/\\d{2}/(\\d{2}|\\d{4})$", names(fp_sheet))

    if (length(date_col) == 0) {
      stop(
        "Could not find a column name with collection date in the format 'Date: dd/mm/yy' or 'Date: dd/mm/yyyy'.",
        call. = FALSE
      )
    }

    date_info <- strsplit(names(fp_sheet)[date_col], "Date: ")[[1]][2]
    date_parts <- strsplit(date_info, "/")[[1]]
    colldd <- date_parts[1]
    collmm <- date_parts[2]
    collyy <- date_parts[3]

    colnames(fp_sheet) <- fp_sheet[1, ]
    fp_sheet <- fp_sheet[-1, ]
  } else if (identical(input_type, "monitora")) {
    plot_meta <- attr(fp_sheet, "plot_meta")
    if (is.null(plot_meta)) plot_meta <- list()
    monitora_years <- plot_meta$census_years

    colldd <- NA_character_
    collmm <- NA_character_
    collyy <- if (length(monitora_years)) as.character(max(monitora_years, na.rm = TRUE)) else NA_character_
  }

  # Edit original info in the forestplot sheet
  if (identical(input_type, "monitora")) {
    voucher_collector_vec <- if ("Voucher Collector" %in% names(fp_sheet)) {
      as.character(fp_sheet[["Voucher Collector"]])
    } else {
      rep(NA_character_, nrow(fp_sheet))
    }

    voucher_collector_vec[is.na(voucher_collector_vec)] <- ""
    voucher_collector_vec <- trimws(voucher_collector_vec)

    collected_vec <- as.character(fp_sheet$Collected)
    collected_vec[is.na(collected_vec)] <- ""
    collected_vec <- trimws(collected_vec)

    keep_rows <- voucher_collector_vec != "" | collected_vec != ""

    fp_sheet <- fp_sheet[keep_rows, , drop = FALSE]

    if (length(collector) > 1L) {
      collector <- collector[keep_rows]
    }

    if (length(addcoll) > 1L) {
      addcoll <- addcoll[keep_rows]
    }

  }  else {
    fp_sheet <- fp_sheet[!is.na(fp_sheet$Collected) & fp_sheet$Collected != "", , drop = FALSE]
  }
  fp_sheet$D <- as.numeric(fp_sheet$D)/10
  fp_sheet$'Census Notes' <- gsub("[.]$", "", fp_sheet$'Census Notes')
  fp_sheet$'Census Notes' <- .convert_to_lowercase(string=fp_sheet$'Census Notes')

  # Prepare the herb_sheet DataFrame structure
  n <- nrow(fp_sheet)
  mtemp <- as.data.frame(matrix(NA, n, ncol(herb_sheet)))
  names(mtemp) <- names(herb_sheet)
  herb_sheet <- mtemp
  if (length(collector) == 1L) collector <- rep(collector, nrow(fp_sheet))
  if (length(addcoll) == 1L) addcoll <- rep(addcoll, nrow(fp_sheet))
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
      # Split the 'Original determination' into components (genero e especie)
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
      # Split the 'Original determination' into components (genero e especie)
      determination <- strsplit(as.character(fp_sheet$`Original determination`)[i], " ")[[1]]
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
      # Split the 'Original determination' into components (genero e especie)
      determination <- strsplit(as.character(fp_sheet$`Original determination`)[i], " ")[[1]]
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

.clean_text <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- ""
  trimws(x)
}
# Retrieve descriptions based on flags
.get_descriptions <- function(language = language) {
  descriptions <- list(
    if (language == "pt"){
      return(data.frame(
        letters = c("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k",
                    "l", "m", "o", "p", "q", "s", "w", "x", "y", "z"),
        description = c(
          "viva normal",
          "com fuste/copa quebrado/a e com rebroto, ou pelo menos com floema/xilema vivo",
          "inclinada \u226510%",
          "ca\u00edda (por ex. no ch\u00e3o)",
          "acanalada e/ou fenestrada",
          "com caule oco",
          "com caule podre e/ou com presen\u00e7a de fungos de prateleira (ou de suporte)",
          "com m\u00faltiplos fustes",
          "sem folhas/com poucas folhas",
          "com caule queimado",
          "com caule quebrado <1,3m",
          "com liana \u2265 10cm de di\u00e2metro no caule ou na copa",
          "coberta por lianas",
          "que sofreu danos causados por raio",
          "cortada",
          "com casca solta/descamando",
          "com estrangulador",
          "com uma ferida e/ou c\u00e2mbio exposto",
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
          "leaning by \u226510%",
          "fallen (e.g. on ground)",
          "fluted or/fenestrated",
          "with hollow",
          "with rotten and or presence of bracket fungus",
          "as multiple stemmed individual",
          "no leaves, few leaves",
          "burnt",
          "snapped < 1.3m",
          "with liana \u226510cm diameter on stem or in canopy",
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
          "inclinado \u226510%",
          "ca\u00eddo (por ejemplo: sobre el suelo)",
          "acanalado y/o fenestrado",
          "con tallo hueco",
          "con tallo podrido",
          "con tallos m\u00faltiples",
          "sin o con pocas hojas",
          "con tallo quemado",
          "con tallo partido <1.3m",
          "con liana \u2265 10cm de di\u00e1metro en el tronco o en la copa del \u00e1rbol",
          "cubierto por lianas",
          "da\u00f1ado por un rayo",
          "cortado",
          "con corteza pelada",
          "con un estrangulador",
          "con una herida o cambio expuesto",
          "da\u00f1ado por elefantes",
          "da\u00f1ado por termitas",
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
          "copa totalmente exposta \u00e0 luz vertical e lateral numa curva de 45 graus",
          "copa totalmente exposta \u00e0 luz vertical, mas a luz lateral est\u00e1 bloqueada por alguns ou todos os cones invertidos de 90 graus que englobam a copa",
          "sob muita luz vertical (>50%)",
          "sob pouca luz vertical (menos de 50% da copa est\u00e1 exposta \u00e0 luz vertical)",
          "sob muita luz lateral",
          "sob m\u00e9dia luz lateral",
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
          "copa completamente expuesta a luz vertical, pero la luz lateral est\u00e1 bloqueada",
          "sob iluminaci\u00f3n vertical alta (m\u00e1s que 50%)",
          "sob poca luz vertical (menos que 50%)",
          "sob luz lateral alta",
          "sob luz lateral media",
          "sob luz lateral baja",
          "no hay luz directa")))
    }
  )
  return(luminosity[[language]])
}

# Convert the first word and words after commas to lowercase
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

# Build herbarium notes
.build_notes <- function(fp_sheet, herb_sheet, i, notes = notes,
                         language = "en",
                         add_census_notes = TRUE) {

  # Loop through each row in the fp_sheet data
  for (i in 1:nrow(fp_sheet)) {

    flag_value <- .clean_text(fp_sheet$Flag1[i])

    if (nzchar(flag_value)) {
      flag_letters <- strsplit(tolower(flag_value), split = "")[[1]]
      converted_descriptions <- .get_descriptions(language = language)$description[
        match(flag_letters, .get_descriptions(language = language)$letters)
      ]
      converted_descriptions <- converted_descriptions[
        !is.na(converted_descriptions) & nzchar(converted_descriptions)
      ]
      description_flag <- paste(converted_descriptions, collapse = ", ")
    } else {
      description_flag <- ""
    }

    li_code <- .clean_text(fp_sheet$LI[i])

    if (nzchar(li_code)) {
      description_li <- .get_luminosity(language = language)$description[
        match(li_code, .get_luminosity(language = language)$code)
      ]
      description_li <- .clean_text(description_li)
    } else {
      description_li <- ""
    }

    census_note_i <- .clean_text(fp_sheet$`Census Notes`[i])
    tag_i <- .clean_text(fp_sheet$`New Tag No`[i])
    t1_i <- .clean_text(fp_sheet$T1[i])
    x_i <- .clean_text(fp_sheet$X[i])
    y_i <- .clean_text(fp_sheet$Y[i])
    d_i <- .clean_text(fp_sheet$D[i])

    # English version
    if (language == "en") {
      if (add_census_notes) {
        herb_sheet[[notes]][i] <- paste0(
          "Tree, ",
          ifelse(d_i != "", paste0(d_i, "cm DBH, "), ""),
          ifelse(description_flag != "", paste0(description_flag, ", "), ""),
          ifelse(census_note_i != "", paste0(census_note_i, ", "), ""),
          description_li,
          ifelse(description_li != "" | tag_i != "" | t1_i != "", ". ", ""),
          ifelse(tag_i != "", paste0("Individual #", tag_i), ""),
          ifelse(t1_i != "",
                 paste0(" in the subplot ", t1_i, ", x = ", x_i, "m, y = ", y_i, "m."),
                 "")
        )
      } else {
        herb_sheet[[notes]][i] <- paste0(
          "Tree, ",
          ifelse(d_i != "", paste0(d_i, "cm DBH, "), ""),
          ifelse(description_flag != "", paste0(description_flag, ", "), ""),
          description_li,
          ifelse(description_li != "" | tag_i != "" | t1_i != "", ". ", ""),
          ifelse(tag_i != "", paste0("Individual #", tag_i), ""),
          ifelse(t1_i != "",
                 paste0(" in the subplot ", t1_i, ", x = ", x_i, "m, y = ", y_i, "m."),
                 "")
        )
      }

    } else if (language == "pt") {

      # Portuguese version
      if (add_census_notes) {
        herb_sheet[[notes]][i] <- paste0(
          "Árvore, ",
          ifelse(d_i != "", paste0(d_i, "cm de diâmetro à altura do peito (DAP), "), ""),
          ifelse(description_flag != "", paste0(description_flag, ", "), ""),
          ifelse(census_note_i != "", paste0(census_note_i, ", "), ""),
          description_li,
          ifelse(description_li != "" | tag_i != "" | t1_i != "", ". ", ""),
          ifelse(tag_i != "", paste0("Indivíduo #", tag_i), ""),
          ifelse(t1_i != "",
                 paste0(" na subparcela ", t1_i, ", x = ", x_i, "m, y = ", y_i, "m."),
                 "")
        )
      } else {
        herb_sheet[[notes]][i] <- paste0(
          "Árvore, ",
          ifelse(d_i != "", paste0(d_i, "cm de diâmetro à altura do peito (DAP), "), ""),
          ifelse(description_flag != "", paste0(description_flag, ", "), ""),
          description_li,
          ifelse(description_li != "" | tag_i != "" | t1_i != "", ". ", ""),
          ifelse(tag_i != "", paste0("Indivíduo #", tag_i), ""),
          ifelse(t1_i != "",
                 paste0(" na subparcela ", t1_i, ", x = ", x_i, "m, y = ", y_i, "m."),
                 "")
        )
      }

    } else if (language == "es") {

      # Spanish version
      if (add_census_notes) {
        herb_sheet[[notes]][i] <- paste0(
          "Árbol, ",
          ifelse(d_i != "", paste0(d_i, "cm de diámetro a la altura del pecho (DAP), "), ""),
          ifelse(description_flag != "", paste0(description_flag, ", "), ""),
          ifelse(census_note_i != "", paste0(census_note_i, ", "), ""),
          description_li,
          ifelse(description_li != "" | tag_i != "" | t1_i != "", ". ", ""),
          ifelse(tag_i != "", paste0("Individuo #", tag_i), ""),
          ifelse(t1_i != "",
                 paste0(" en la subparcelas ", t1_i, ", x = ", x_i, "m, y = ", y_i, "m."),
                 "")
        )
      } else {
        herb_sheet[[notes]][i] <- paste0(
          "Árbol, ",
          ifelse(d_i != "", paste0(d_i, "cm de diámetro a la altura del pecho (DAP), "), ""),
          ifelse(description_flag != "", paste0(description_flag, ", "), ""),
          description_li,
          ifelse(description_li != "" | tag_i != "" | t1_i != "", ". ", ""),
          ifelse(tag_i != "", paste0("Individuo #", tag_i), ""),
          ifelse(t1_i != "",
                 paste0(" en la subparcelas ", t1_i, ", x = ", x_i, "m, y = ", y_i, "m."),
                 "")
        )
      }
    }

    herb_sheet[[notes]][i] <- gsub(",\\s*,", ", ", herb_sheet[[notes]][i])
    herb_sheet[[notes]][i] <- gsub(",\\s*\\.", ".", herb_sheet[[notes]][i])
    herb_sheet[[notes]][i] <- gsub("\\.\\s*\\.", ".", herb_sheet[[notes]][i])
    herb_sheet[[notes]][i] <- gsub("\\s+", " ", herb_sheet[[notes]][i])
    herb_sheet[[notes]][i] <- trimws(herb_sheet[[notes]][i])

    herb_sheet[[notes]][i] <- gsub(",\\s*$", "", herb_sheet[[notes]][i])
  }

  return(herb_sheet)
}
