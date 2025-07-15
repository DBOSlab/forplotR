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
    valid_states <- c("Acre" = "AC", "Alagoas" = "AL", "Amapá" = "AP", "Amazonas" = "AM",
                      "Bahia" = "BA", "Ceará" = "CE", "Distrito Federal" = "DF",
                      "Espírito Santo" = "ES", "Goiás" = "GO", "Maranhão" = "MA",
                      "Mato Grosso" = "MT", "Mato Grosso do Sul" = "MS", "Minas Gerais" = "MG",
                      "Pará" = "PA", "Paraíba" = "PB", "Paraná" = "PR", "Pernambuco" = "PE",
                      "Piauí" = "PI", "Rio de Janeiro" = "RJ", "Rio Grande do Norte" = "RN",
                      "Rio Grande do Sul" = "RS", "Rondônia" = "RO", "Roraima" = "RR",
                      "Santa Catarina" = "SC", "São Paulo" = "SP", "Sergipe" = "SE",
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
  
  # Retorna o valor original se não for Brasil
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
    "Brazil" = c("Brasil", "Brésil"),
    "Argentina" = c("Argentine"),
    "Colombia" = c("Colômbia", "Colombie"),
    "Peru" = c("Pérou"),
    "France" = c("França"),
    "Spain" = c("España", "Espanha", "Espagne"),
    "Mexico" = c("México", "Mexique"),
    "Paraguay" = c("Paraguai"),
    "Uruguay" = c("Uruguai"),
    "Bolivia" = c("Bolívia", "Bolivie")
  )
  
  # Normaliza o nome de entrada removendo acentos
  normalized_input <- stringi::stri_trans_general(user_country_input, "Latin-ASCII")
  
  # Caso já seja um nome padrão, retorna direto
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
  library(terra)
  library(geodata)
  library(stringi)
  
  # Translate country name to standardized format
  standardized_country <- .translate_country(country)
  
  # Brazilian state abbreviations to full names
  state_name_map <- c(
    "AC" = "Acre", "AL" = "Alagoas", "AP" = "Amapá", "AM" = "Amazonas",
    "BA" = "Bahia", "CE" = "Ceará", "DF" = "Distrito Federal", "ES" = "Espírito Santo",
    "GO" = "Goiás", "MA" = "Maranhão", "MT" = "Mato Grosso", "MS" = "Mato Grosso do Sul",
    "MG" = "Minas Gerais", "PA" = "Pará", "PB" = "Paraíba", "PR" = "Paraná",
    "PE" = "Pernambuco", "PI" = "Piauí", "RJ" = "Rio de Janeiro", "RN" = "Rio Grande do Norte",
    "RS" = "Rio Grande do Sul", "RO" = "Rondônia", "RR" = "Roraima",
    "SC" = "Santa Catarina", "SP" = "São Paulo", "SE" = "Sergipe", "TO" = "Tocantins"
  )
  
  # Convert abbreviation to full state name if applicable
  if (majorarea %in% names(state_name_map)) {
    majorarea <- state_name_map[majorarea]
  }
  
  # Download administrative boundaries
  boundaries <- geodata::gadm(country = standardized_country, level = level, path = tempdir())
  
  if (is.null(boundaries) || nrow(boundaries) == 0) {
    stop("❌ No administrative boundaries found for the selected country. Check the country name.")
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
    stop(paste0("❌ Could not match '", majorarea, "' to any administrative region.\n",
                "Please check spelling, abbreviation, or administrative level."))
  }
  
  # Extract matched geometry
  admin_geom <- boundaries[match_idx, ]
  
  # Create point geometry
  point <- vect(matrix(c(long, lat), ncol = 2), type = "points", crs = crs(boundaries))
  
  # Check if point lies within administrative polygon
  inside <- any(relate(point, admin_geom, "intersects"))
  
  if (!inside) {
    stop("❌ The latitude and longitude DO NOT fall within the specified administrative region.\n",
         "Please verify coordinates and the provided administrative name.")
  }
}


#_______________________________________________________________________________
# Check plot size when plotting forest balance ####

.validate_plot_size <- function(plot_size) {
  if (!plot_size %in% c(0.2, 0.5, 1)) {
    stop("`plot_size` must be one of: 0.2, 0.5, or 1 hectare.", call. = FALSE)
  }
}


#_______________________________________________________________________________
# Check subplot size when plotting forest balance ####

.validate_subplot_size <- function(subplot_size) {
  if (!subplot_size %in% c(10, 20, 25)) {
    stop("`subplot_size` must be one of: 10, 20, or 25 meters.", call. = FALSE)
  }
}
