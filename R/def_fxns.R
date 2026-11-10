# Small functions to evaluate user input for the main functions and
# return meaningful errors.
# Author: Giulia Ottino & Domingos Cardoso


#_______________________________________________________________________________
# Check if the dir input is "character" type and if it has a "/" in the end ####

.arg_check_dir <- function(x) {
  # Check classes
  class_x <- class(x)
  if (!"character" %in% class_x) {
    stop(paste0("The argument dir should be a character, not '", class_x, "'."),
         call. = FALSE)
  }
  if (grepl("[/]$", x)) {
    x <- gsub("[/]$", "", x)
  }
  return(x)
}


#_______________________________________________________________________________

.arg_check_majorarea <- function(majorarea, country) {

  if (missing(majorarea)) {
    stop("'majorarea' is required.")
  }

  if (country == "Brasil") {
    valid_states <- c("Acre" = "AC", "Alagoas" = "AL", "Amap\u00e1" = "AP", "Amazonas" = "AM",
                      "Bahia" = "BA", "Cear\u00e1" = "CE", "Distrito Federal" = "DF",
                      "Esp\u00edrito Santo" = "ES", "Goi\u00e1s" = "GO", "Maranh\u00e3o" = "MA",
                      "Mato Grosso" = "MT", "Mato Grosso do Sul" = "MS", "Minas Gerais" = "MG",
                      "Par\u00e1" = "PA", "Para\u00edba" = "PB", "Paran\u00e1" = "PR", "Pernambuco" = "PE",
                      "Piau\u00ed" = "PI", "Rio de Janeiro" = "RJ", "Rio Grande do Norte" = "RN",
                      "Rio Grande do Sul" = "RS", "Rond\u00f4nia" = "RO", "Roraima" = "RR",
                      "Santa Catarina" = "SC", "S\u00e3o Paulo" = "SP", "Sergipe" = "SE",
                      "Tocantins" = "TO")

    valid_states_full <- names(valid_states)
    valid_states_acronyms <- unname(valid_states)

    states_no_diacritics <- stringi::stri_trans_general(majorarea, "Latin-ASCII")
    valid_states_full_no_diacritics <- stringi::stri_trans_general(valid_states_full, "Latin-ASCII")
    valid_states_acronyms_no_diacritics <- stringi::stri_trans_general(valid_states_acronyms, "Latin-ASCII")

    match_full <- match(states_no_diacritics, valid_states_full_no_diacritics)
    match_acronym <- match(states_no_diacritics, valid_states_acronyms_no_diacritics)

    if (!is.na(match_full)) {
      return(valid_states_full[match_full])
    } else if (!is.na(match_acronym)) {
      return(valid_states_full[match_acronym])
    } else {
      stop(paste0("'", majorarea, "' is not a Brazilian state."))
    }
  }

  # Retorna o valor original se nao for Brasil
  return(majorarea)
}


#_______________________________________________________________________________
# Check if the minorarea input is present ####

.arg_check_minorarea <- function(minorarea) {
  if (missing(minorarea)) {
    stop("'minorarea' is required.")
  }
}


#_______________________________________________________________________________
# Check if the lat input is present or correct ####

.arg_check_lat <- function(lat) {
  if (missing(lat)) {
    stop("'lat' is required.")
  }
  class_x <- class(lat)
  if ("character" %in% class_x) {
    if (grepl("[A-Za-z]", lat)) {
      stop("'", lat, "' must be a valid numeric latitude.")
    } else {
      lat <- as.numeric(lat)
    }
  }
  return(lat)
}

#_______________________________________________________________________________
# Check if the long input is present or correct ####
.arg_check_long <- function(long) {
  if (missing(long)) {
    stop("'long' is required.")
  }
  class_x <- class(long)
  if ("character" %in% class_x) {
    if (grepl("[A-Za-z]", long)) {
      stop("'", long, "' must be a valid numeric longitude.")
    } else {
      long <- as.numeric(long)
    }
  }
  return(long)
}


#_______________________________________________________________________________
# Check if the alt input is present ####

.arg_check_alt <- function(alt) {
  class_x <- class(alt)
  if ("character" %in% class_x) {
    if (grepl("[A-Za-z]", alt)) {
      stop("'", alt, "' must be a number.")
    } else {
      alt <- as.numeric(alt)
    }
  }
  return(alt)
}


#_______________________________________________________________________________

.translate_country <- function(user_country_input) {
  # Lista de nomes padronizados
  standard_names <- c("Brazil", "Argentina", "Colombia", "Peru", "France",
                      "Spain", "Mexico", "Chile", "Paraguay", "Uruguay", "Bolivia")

  # Mapeamento de variantes conhecidas
  country_map <- list(
    "Brazil" = c("Brasil", "Br\u00e9sil"),
    "Argentina" = c("Argentine"),
    "Colombia" = c("Col\u00f4mbia", "Colombie"),
    "Peru" = c("P\u00e9rou"),
    "France" = c("Fran\u00e7a"),
    "Spain" = c("Espa\u00f1a", "Espanha", "Espagne"),
    "Mexico" = c("M\u00e9xico", "Mexique"),
    "Paraguay" = c("Paraguai"),
    "Uruguay" = c("Uruguai"),
    "Bolivia" = c("Bol\u00edvia", "Bolivie")
  )

  # Normaliza o nome de entrada removendo acentos
  normalized_input <- stringi::stri_trans_general(user_country_input, "Latin-ASCII")

  # Caso ja seja um nome padrao, retorna direto
  if (normalized_input %in% standard_names) {
    return(normalized_input)
  }

  # Verifica se corresponde a alguma variante
  for (country in names(country_map)) {
    normalized_variants <- stringi::stri_trans_general(country_map[[country]], "Latin-ASCII")
    if (normalized_input %in% normalized_variants) {
      return(country)
    }
  }

  stop(paste0("Unknown country: ", user_country_input))
}


#_______________________________________________________________________________
# Main function to validate if coordinates fall within the specified administrative division ####

.check_point_in_admin <- function(lat, long, majorarea, country, level = 1) {

  # Translate country name to standardized format
  standardized_country <- .translate_country(country)

  # Brazilian state abbreviations to full names
  state_name_map <- c(
    "AC" = "Acre", "AL" = "Alagoas", "AP" = "Amap\u00e1", "AM" = "Amazonas",
    "BA" = "Bahia", "CE" = "Cear\u00e1", "DF" = "Distrito Federal", "ES" = "Esp\u00edrito Santo",
    "GO" = "Goi\u00e1s", "MA" = "Maranh\u00e3o", "MT" = "Mato Grosso", "MS" = "Mato Grosso do Sul",
    "MG" = "Minas Gerais", "PA" = "Par\u00e1", "PB" = "Para\u00edba", "PR" = "Paran\u00e1",
    "PE" = "Pernambuco", "PI" = "Piau\u00ed", "RJ" = "Rio de Janeiro", "RN" = "Rio Grande do Norte",
    "RS" = "Rio Grande do Sul", "RO" = "Rond\u00f4nia", "RR" = "Roraima",
    "SC" = "Santa Catarina", "SP" = "S\u00e3o Paulo", "SE" = "Sergipe", "TO" = "Tocantins"
  )

  # Convert abbreviation to full state name if applicable
  if (majorarea %in% names(state_name_map)) {
    majorarea <- state_name_map[majorarea]
  }

  # Download administrative boundaries
  boundaries <- geodata::gadm(country = standardized_country, level = level, path = tempdir())

  if (is.null(boundaries) || nrow(boundaries) == 0) {
    stop("\u274c No administrative boundaries found for the selected country. Check the country name.")
  }

  # Standardize admin names by removing accents and trimming
  majorarea_clean <- stri_trim_both(stri_trans_general(majorarea, "Latin-ASCII"))
  boundaries$NAME_1 <- as.character(boundaries$NAME_1)
  boundaries$NAME_CLEAN <- stri_trim_both(stri_trans_general(boundaries$NAME_1, "Latin-ASCII"))

  # Match by exact name
  match_idx <- which(boundaries$NAME_CLEAN == majorarea_clean)

  # If no exact match, try partial match
  if (length(match_idx) == 0) {
    match_idx <- grep(majorarea_clean, boundaries$NAME_CLEAN, ignore.case = TRUE)
  }

  if (length(match_idx) == 0) {
    stop(paste0("\u274c Could not match '", majorarea, "' to any administrative region.\n",
                "Please check spelling, abbreviation, or administrative level."))
  }

  # Extract matched geometry
  admin_geom <- boundaries[match_idx, ]

  # Create point geometry
  point <- vect(matrix(c(long, lat), ncol = 2), type = "points", crs = crs(boundaries))

  # Check if point lies within administrative polygon
  inside <- any(relate(point, admin_geom, "intersects"))

  if (!inside) {
    stop("\u274c The latitude and longitude DO NOT fall within the specified administrative region.\n",
         "Please verify coordinates and the provided administrative name.")
  }
}


#_______________________________________________________________________________
# Check plot size when plotting forest balance ####

.validate_plot_size <- function(plot_size) {
  ok <- is.numeric(plot_size) && length(plot_size) == 1L && is.finite(plot_size) && plot_size > 0
  if (!ok) {
    stop("`plot_size` must be a single positive numeric value in hectares.", call. = FALSE)
  }
  invisible(TRUE)
}

#_______________________________________________________________________________
# Check subplot size when plotting forest balance ####

.validate_subplot_size <- function(subplot_size) {
  ok <- is.numeric(subplot_size) && length(subplot_size) == 1L && is.finite(subplot_size) && subplot_size > 0
  if (!ok) {
    stop("`subplot_size` must be a single positive numeric value in meters.", call. = FALSE)
  }
  invisible(TRUE)
}

#_______________________________________________________________________________
# Validate herbarium codes ####

#' Validate herbarium codes
#'
#' @param x Character vector of herbarium codes.
#'
#' @return Invisibly TRUE.
#'
#' @keywords internal
#' @noRd
#'
.arg_check_herbarium <- function(x) {
  if (is.null(x) || !length(x)) {
    stop("`herbaria` must be a non-empty character vector.", call. = FALSE)
  }

  x <- toupper(trimws(as.character(x)))
  x <- x[nzchar(x)]

  if (!length(x)) {
    stop("`herbaria` must be a non-empty character vector.", call. = FALSE)
  }

  bad <- x[!grepl("^[A-Z0-9]+$", x)]
  if (length(bad)) {
    stop(
      "Invalid herbarium code(s): ",
      paste(unique(bad), collapse = ", "),
      call. = FALSE
    )
  }

  invisible(TRUE)
}

#_______________________________________________________________________________
# Column mapping for forplot_balance ####

#' Registry of user-facing colname_* arguments and their canonical targets
#'
#' Keys starting with "." are not canonical columns; they are only used by
#' the MONITORA reader (station filter and census year).
#'
#' @keywords internal
#' @noRd
.forplot_colname_registry <- function() {
  c(
    colname_tagno       = "New Tag No",
    colname_stemgroup   = "New Stem Grouping",
    colname_subplot     = "T1",
    colname_subplot_t2  = "T2",
    colname_x           = "X",
    colname_y           = "Y",
    colname_family      = "Family",
    colname_species     = "Original determination",
    colname_dbh         = "D",
    colname_cap         = "CAP",
    colname_voucher     = "Voucher",
    colname_collected   = "Collected",
    colname_station     = ".station",
    colname_census_year = ".census_year"
  )
}


#' Validate one colname_* argument
#'
#' @keywords internal
#' @noRd
.arg_check_colname <- function(x, arg) {
  if (is.null(x)) return(invisible(TRUE))
  if (!is.character(x) || length(x) != 1L || is.na(x) || !nzchar(trimws(x))) {
    stop("`", arg, "` must be NULL or a single non-empty character string.",
         call. = FALSE)
  }
  invisible(TRUE)
}


#' Build the column map (canonical name -> user column name)
#'
#' @param ... colname_* arguments, named as in .forplot_colname_registry().
#'
#' @return Named character vector (names = canonical keys) with attribute
#'   `arg_names` (canonical key -> argument name), or NULL if nothing mapped.
#'
#' @keywords internal
#' @noRd
.build_col_map <- function(...) {
  args <- list(...)
  reg <- .forplot_colname_registry()

  unknown <- setdiff(names(args), names(reg))
  if (length(unknown)) {
    stop("Unknown column-mapping argument(s): ",
         paste(unknown, collapse = ", "), call. = FALSE)
  }

  for (nm in names(args)) .arg_check_colname(args[[nm]], nm)

  args <- args[!vapply(args, is.null, logical(1))]
  if (!length(args)) return(NULL)

  keys <- unname(reg[names(args)])
  out <- stats::setNames(trimws(unlist(args, use.names = FALSE)), keys)
  attr(out, "arg_names") <- stats::setNames(names(args), keys)
  out
}


#' Resolve a user-supplied column name against a data frame
#'
#' Tries exact match, then trimmed match, then normalized match
#' (case-, accent- and punctuation-insensitive).
#'
#' @keywords internal
#' @noRd
.resolve_user_col <- function(df, user_col, arg_name = NULL) {
  nm <- names(df)

  if (user_col %in% nm) return(user_col)

  hit <- which(trimws(nm) == trimws(user_col))
  if (length(hit) == 1L) return(nm[hit])

  hit <- which(.norm_nm(nm) == .norm_nm(user_col))
  if (length(hit) == 1L) return(nm[hit])

  arg_txt <- if (!is.null(arg_name)) paste0(" (from `", arg_name, "`)") else ""

  if (length(hit) > 1L) {
    stop(sprintf("Column '%s'%s matches more than one input column: %s",
                 user_col, arg_txt, paste(nm[hit], collapse = " | ")),
         call. = FALSE)
  }

  stop(sprintf("Column '%s'%s was not found in the input data. Available columns: %s",
               user_col, arg_txt, paste(nm, collapse = " | ")),
       call. = FALSE)
}


#' Return the mapped column for a key, or a fallback when not mapped
#'
#' @keywords internal
#' @noRd
.mapped_col <- function(df, col_map, key, fallback = NA_character_) {
  if (is.null(col_map) || !(key %in% names(col_map))) return(fallback)
  arg_names <- attr(col_map, "arg_names")
  .resolve_user_col(
    df,
    col_map[[key]],
    if (!is.null(arg_names)) arg_names[[key]] else NULL
  )
}


#' Copy user columns into their canonical names
#'
#' Values are snapshotted before assignment, so swaps (e.g. X <-> Y) and
#' one source column feeding two targets both work.
#'
#' @keywords internal
#' @noRd
.apply_col_map <- function(df, col_map) {
  if (is.null(col_map) || !length(col_map)) return(df)

  keys <- names(col_map)
  keys <- keys[!startsWith(keys, ".")]
  if (!length(keys)) return(df)

  src <- vapply(keys, function(k) .mapped_col(df, col_map, k), character(1))
  vals <- lapply(src, function(s) df[[s]])

  for (i in seq_along(keys)) {
    df[[keys[i]]] <- vals[[i]]
  }
  df
}
