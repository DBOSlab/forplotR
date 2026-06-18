#' Cross-check inventory determinations against one selected JABOT herbarium
#'
#' @author
#' Giulia Ottino & Domingos Cardoso
#'
#' @description
#' Cross-checks voucher determinations from an inventory spreadsheet against
#' specimen records from one selected herbarium available in JABOT Darwin Core
#' Archive files.
#'
#' The function is intended for inventory workflows already supported in the
#' package, namely ForestPlots-style spreadsheets and MONITORA spreadsheets.
#' These input files are first converted to a common internal schema using the
#' package helper functions, and then matched to herbarium records using a
#' normalized collector-name + collector-number identity key.
#'
#' Matching is restricted to the herbarium informed in \code{herbaria}.
#' Therefore, unlike broader duplicate-based determination comparison routines,
#' this function focuses on checking whether an inventory voucher is already
#' present in one selected herbarium and, if so, whether the determination in
#' the inventory agrees with the determination currently available in JABOT for
#' that herbarium.
#'
#' @details
#' The matching workflow is:
#' \enumerate{
#'   \item Read and standardize the inventory spreadsheet into the package's
#'   common field-sheet-like structure.
#'   \item Extract and normalize voucher identity from the inventory using the
#'   existing package helper \code{.parse_census_identity()}.
#'   \item Resolve local \code{occurrence.txt} files under \code{path}, or,
#'   if needed, automatically discover and download the corresponding JABOT
#'   DwC-A resource for the selected herbarium into a cache directory.
#'   \item Load the selected herbarium records into DuckDB.
#'   \item Build normalized collector-number keys for herbarium records.
#'   \item Match inventory vouchers to herbarium records by primary or fallback
#'   identity keys.
#'   \item Normalize taxon names on both sides before comparison.
#'   \item Classify each match into a diagnostic category.
#' }
#'
#' Determination comparison uses simplified canonical taxon strings. On both
#' inventory and herbarium sides, authorship is ignored for routine comparison.
#' Family-only records are treated as \code{"Indet indet"}. Genus-only records,
#' or records such as \code{"Inga"} and \code{"Inga indet"}, are normalized to
#' \code{"Inga indet"}. Cases with qualifiers such as \code{cf.},
#' \code{aff.}, \code{subsp.}, \code{ssp.}, and \code{var.} are preserved in
#' the normalized comparison string when possible.
#'
#' On the herbarium side, canonical names are built preferentially from
#' structured Darwin Core fields (\code{genus}, \code{specificEpithet},
#' \code{infraspecificEpithet}), with fallback to the best available scientific
#' name field. Family names are never reused as genus names during
#' normalization.
#'
#' The main diagnostic categories are:
#' \itemize{
#'   \item \code{inventory_indet_herbarium_det}: the inventory is indeterminate
#'   but the herbarium has a usable determination;
#'   \item \code{inventory_det_herbarium_indet}: the inventory is determined but
#'   the herbarium record is indeterminate or lacks a usable determination;
#'   \item \code{same_determination}: inventory and herbarium agree after
#'   normalization;
#'   \item \code{different_determination}: inventory and herbarium disagree
#'   after normalization.
#' }
#'
#' Records classified as \code{same_determination} are counted in the
#' diagnostics table but are omitted from the \code{matches} output sheet.
#'
#' This function works offline after download. If local DwC-A files are not
#' available, the function can automatically download the required JABOT archive
#' for the selected herbarium and then proceed locally.
#'
#' @usage
#' forplot_crossdet(fp_file_path,
#'                  input_type = c("field_sheet", "fp_query_sheet", "monitora"),
#'                  herbaria,
#'                  collector_fallback = NULL,
#'                  collector_codes = NULL,
#'                  station_name = NULL,
#'                  sheet = NULL,
#'                  path = NULL,
#'                  only_collected = TRUE,
#'                  save = TRUE,
#'                  dir = "crossdet_plot",
#'                  filename = "crossdet_plot",
#'                  dbdir = NULL,
#'                  force_refresh = FALSE,
#'                  verbose = FALSE,
#'                  resource_map = NULL)
#'
#' @param fp_file_path Character scalar. Path to the inventory Excel file.
#'
#' @param input_type Character scalar. One of \code{"field_sheet"},
#'   \code{"fp_query_sheet"}, or \code{"monitora"}.
#'
#' @param herbaria Character scalar. One herbarium code to be checked in JABOT,
#'   for example \code{"RB"}.
#'
#' @param collector_fallback Optional character scalar. Collector name used when
#'   voucher strings contain only the collection number.
#'
#' @param collector_codes Optional named character vector mapping compact
#'   collector prefixes to full collector names. This is passed to
#'   \code{.parse_census_identity()} and is useful when vouchers use compact
#'   forms such as \code{"DC1234"}.
#'
#' @param station_name Optional station filter passed to the MONITORA helper.
#'   Ignored for non-MONITORA inputs.
#'
#' @param sheet Optional worksheet selector passed to the spreadsheet reader
#'   helper.
#'
#' @param path Optional character scalar. Directory containing local JABOT
#'   DwC-A folders with \code{occurrence.txt} files. If \code{NULL}, the
#'   function uses the package cache directory and can automatically download
#'   the required archive when needed.
#'
#' @param only_collected Logical. If \code{TRUE}, restricts the inventory side
#'   to rows marked as collected and/or with non-empty voucher values.
#'
#' @param save Logical. If \code{TRUE}, saves an Excel workbook with the result
#'   sheets.
#'
#' @param dir Character scalar. Output directory used when \code{save = TRUE}.
#'
#' @param filename Character scalar. Output file name without extension.
#'
#' @param dbdir Character scalar or \code{NULL}. DuckDB file path. Use
#'   \code{NULL} for an in-memory database. When \code{NULL}, the function may
#'   also use the package cache DuckDB path if available.
#'
#' @param force_refresh Logical. If \code{TRUE}, forces a fresh download of the
#'   JABOT DwC-A archive instead of reusing cached files.
#'
#' @param verbose Logical. If \code{TRUE}, prints progress messages during
#'   archive discovery, download, loading, and matching.
#'
#' @param resource_map Optional named character vector overriding automatic IPT
#'   resource discovery. Names should follow the pattern
#'   \code{"HERBARIUM::source"}, for example \code{"RB::jabot"}.
#'
#' @return A list with three elements:
#' \describe{
#'   \item{matches}{A data frame with one row per matched inventory-herbarium
#'   record that remains diagnostically relevant after normalization. Records
#'   classified as \code{same_determination} are omitted from this table.}
#'   \item{diagnostics}{A summary table with counts and human-readable
#'   interpretations of the diagnostic classes.}
#'   \item{unmatched_inventory}{Inventory rows with a usable voucher identity
#'   that were not found in the selected herbarium.}
#' }
#'
#' @seealso \code{\link{jabot_crossdet}}
#'
#' @examples
#' \dontrun{
#' library(forplotR)
#' library(openxlsx)
#'
#' data(RUS01)
#' openxlsx::write.xlsx(RUS01, "RUS_plot.xlsx")
#'
#' out1 <- forplot_crossdet(
#'   fp_file_path = "RUS_plot.xlsx",
#'   input_type = "field_sheet",
#'   herbaria = "RB",
#'   collector_fallback = "G. Ottino",
#'   save = FALSE
#' )
#'
#' data(TUM)
#' openxlsx::write.xlsx(TUM, "monitora_station.xlsx")
#'
#' out2 <- forplot_crossdet(
#'   fp_file_path = "monitora_station.xlsx",
#'   input_type = "monitora",
#'   herbaria = "RB",
#'   station_name = "Estacao 2",
#'   save = FALSE
#' )
#'
#' out3 <- forplot_crossdet(
#'   fp_file_path = "monitora_station.xlsx",
#'   input_type = "monitora",
#'   herbaria = "RB",
#'   station_name = 1,
#'   collector_fallback = "H. Medeiros",
#'   force_refresh = TRUE,
#'   verbose = TRUE,
#'   save = FALSE
#' )
#' }
#'
#' @importFrom DBI dbConnect dbDisconnect dbExecute dbGetQuery
#' @importFrom duckdb duckdb
#'
#' @export
#'
forplot_crossdet <- function(fp_file_path,
                             input_type = c("field_sheet", "fp_query_sheet", "monitora"),
                             herbaria,
                             collector_fallback = NULL,
                             collector_codes = NULL,
                             station_name = NULL,
                             sheet = NULL,
                             path = NULL,
                             only_collected = TRUE,
                             save = TRUE,
                             dir = "crossdet_plot",
                             filename = "crossdet_plot",
                             dbdir = NULL,
                             force_refresh = FALSE,
                             verbose = FALSE,
                             resource_map = NULL) {

  `%||%` <- function(x, y) if (is.null(x)) y else x

  # Save the three output tables to an xlsx workbook.
  .save_xlsx_three_sheets <- function(matches, diagnostics, unmatched, dir, filename) {
    if (!requireNamespace("writexl", quietly = TRUE)) {
      stop(
        "Package 'writexl' is required when save = TRUE. Install it with install.packages('writexl').",
        call. = FALSE
      )
    }

    if (!dir.exists(dir)) {
      dir.create(dir, recursive = TRUE, showWarnings = FALSE)
    }

    out <- file.path(dir, paste0(filename, ".xlsx"))

    writexl::write_xlsx(
      x = list(
        Matches = matches,
        Diagnostics = diagnostics,
        Unmatched_inventory = unmatched
      ),
      path = out
    )

    message("Saved: ", out)
    invisible(out)
  }

  # Clean strings, remove accents, and normalize whitespace.
  .clean_chr_local <- function(x) {
    x <- as.character(x)
    x <- iconv(x, to = "ASCII//TRANSLIT", sub = "")
    x[is.na(x)] <- ""
    x <- gsub("\\s+", " ", x)
    trimws(x)
  }

  # Convert a string to simple title case for the first word.
  .title_case_first <- function(x) {
    x <- .clean_chr_local(x)
    ifelse(
      !nzchar(x),
      "",
      paste0(toupper(substr(x, 1, 1)), substr(tolower(x), 2, nchar(x)))
    )
  }

  # Detect family-like strings so they are not reused as genus names.
  .is_family_like <- function(x) {
    x <- tolower(.clean_chr_local(x))
    !nzchar(x) |
      grepl("aceae$", x) |
      x %in% c("leguminosae", "compositae", "gramineae", "labiatae",
               "guttiferae", "umbelliferae", "palmae")
  }

  # Normalize a free-text taxon name for comparison.
  # Rules:
  # - family-only -> "Indet indet"
  # - genus-only -> "Genus indet"
  # - genus + indet/sp -> "Genus indet"
  # - genus + epithet + author -> "Genus epithet"
  # - genus + cf./aff./subsp./ssp./var. + epithet -> keep qualifier when possible
  .canon_taxon_compare <- function(x) {
    x <- as.character(x)
    x[is.na(x)] <- ""
    x <- iconv(x, to = "ASCII//TRANSLIT", sub = "")
    x <- trimws(gsub("\\s+", " ", x))

    out <- rep("Indet indet", length(x))

    qualifiers <- c("cf", "cf.", "aff", "aff.", "subsp", "subsp.", "ssp", "ssp.", "var", "var.")
    indet_terms <- c("indet", "indet.", "sp", "sp.", "spp", "spp.", "indeterminado", "indeterminada")

    for (i in seq_along(x)) {
      xi <- tolower(trimws(x[i]))
      if (!nzchar(xi)) next

      parts <- unlist(strsplit(xi, "\\s+"))
      parts <- parts[nzchar(parts)]
      if (!length(parts)) next

      genus_raw <- parts[1]

      # Treat family-like first terms as fully indeterminate.
      if (.is_family_like(genus_raw)) {
        out[i] <- "Indet indet"
        next
      }

      genus <- .title_case_first(genus_raw)

      # Genus only.
      if (length(parts) == 1L) {
        out[i] <- paste(genus, "indet")
        next
      }

      second <- parts[2]

      # Genus + indeterminate marker.
      if (second %in% indet_terms) {
        out[i] <- paste(genus, "indet")
        next
      }

      # Genus + qualifier + next token.
      if (second %in% qualifiers) {
        if (length(parts) >= 3L) {
          third <- parts[3]

          if (!nzchar(third) || third %in% indet_terms || .is_family_like(third) || third == genus_raw) {
            out[i] <- paste(genus, "indet")
          } else {
            out[i] <- paste(genus, second, tolower(third))
          }
        } else {
          out[i] <- paste(genus, "indet")
        }
        next
      }

      # If the second token behaves like a family or repeats the genus, keep genus as indeterminate.
      if (.is_family_like(second) || second == genus_raw) {
        out[i] <- paste(genus, "indet")
        next
      }

      # Standard genus + specific epithet form.
      if (grepl("^[a-z-]+$", second)) {
        out[i] <- paste(genus, tolower(second))
      } else {
        out[i] <- paste(genus, "indet")
      }
    }

    out
  }

  # Build a normalized herbarium-side taxon name from DwC fields.
  # Structured fields are preferred over free-text name fields.
  .canon_taxon_from_dwc <- function(genus,
                                    specificEpithet,
                                    infraspecificEpithet,
                                    scientificName,
                                    acceptedScientificName,
                                    verbatimScientificName) {
    genus <- .clean_chr_local(genus)
    specificEpithet <- .clean_chr_local(specificEpithet)
    infraspecificEpithet <- .clean_chr_local(infraspecificEpithet)

    best_name <- dplyr::coalesce(
      ifelse(nzchar(.clean_chr_local(acceptedScientificName)), .clean_chr_local(acceptedScientificName), NA_character_),
      ifelse(nzchar(.clean_chr_local(scientificName)), .clean_chr_local(scientificName), NA_character_),
      ifelse(nzchar(.clean_chr_local(verbatimScientificName)), .clean_chr_local(verbatimScientificName), NA_character_)
    )

    out <- rep("Indet indet", length(genus))

    genus_ok <- !is.na(genus) & nzchar(genus) & !.is_family_like(genus)
    epithet_ok <- !is.na(specificEpithet) & nzchar(specificEpithet) & !.is_family_like(specificEpithet)

    # Genus + valid epithet, excluding cases where epithet repeats the genus.
    keep_species <- genus_ok & epithet_ok &
      (tolower(specificEpithet) != tolower(genus))

    out[keep_species] <- paste(
      .title_case_first(genus[keep_species]),
      tolower(specificEpithet[keep_species])
    )

    # Valid genus without a useful epithet.
    keep_genus_only <- genus_ok & !keep_species
    out[keep_genus_only] <- paste(
      .title_case_first(genus[keep_genus_only]),
      "indet"
    )

    # Fallback to the best available free-text scientific name when genus is unusable.
    need_fallback <- !genus_ok & !is.na(best_name) & nzchar(best_name)
    if (any(need_fallback)) {
      out[need_fallback] <- .canon_taxon_compare(best_name[need_fallback])
    }

    out
  }

  # Check whether a normalized name should be treated as indeterminate.
  .is_indet_canon <- function(canon_det) {
    canon_det <- .clean_chr_local(canon_det)
    canon_det[!nzchar(canon_det)] <- "Indet indet"
    canon_low <- tolower(canon_det)

    canon_low == "indet indet" |
      grepl("^[a-z-]+\\s+indet$", canon_low) |
      grepl("^[a-z-]+\\s+(cf|aff|subsp|ssp|var)\\.?$", canon_low)
  }

  # Diagnose the comparison outcome between normalized inventory and herbarium names.
  .diagnose_case <- function(inv_canon, herb_canon) {
    inv_indet <- .is_indet_canon(inv_canon)
    herb_indet <- .is_indet_canon(herb_canon)

    same_det <- !is.na(inv_canon) & !is.na(herb_canon) &
      trimws(tolower(inv_canon)) == trimws(tolower(herb_canon))

    cls <- ifelse(
      same_det,
      "same_determination",
      ifelse(
        inv_indet & !herb_indet,
        "inventory_indet_herbarium_det",
        ifelse(
          !inv_indet & herb_indet,
          "inventory_det_herbarium_indet",
          "different_determination"
        )
      )
    )

    msg <- ifelse(
      cls == "same_determination",
      "Same determination in inventory and herbarium after taxonomic normalization.",
      ifelse(
        cls == "inventory_indet_herbarium_det",
        "Inventory indeterminate; herbarium determined. This is a priority case for updating the inventory identification.",
        ifelse(
          cls == "inventory_det_herbarium_indet",
          "Inventory determined; herbarium indeterminate. The herbarium record may need updating or checking.",
          "Different determination between inventory and herbarium. Taxonomic review is recommended."
        )
      )
    )

    data.frame(
      diagnostic_class = cls,
      diagnostic_message = msg,
      stringsAsFactors = FALSE
    )
  }

  # Standardize the input inventory spreadsheet to the common internal schema.
  .inventory_to_common_df <- function(fp_file_path,
                                      input_type,
                                      sheet = NULL,
                                      station_name = NULL) {
    input_type <- match.arg(tolower(input_type), c("field_sheet", "fp_query_sheet", "monitora"))

    if (identical(input_type, "monitora")) {
      if (!exists(".monitora_to_field_sheet_df", mode = "function")) {
        stop("Required helper `.monitora_to_field_sheet_df()` was not found.", call. = FALSE)
      }

      out <- .monitora_to_field_sheet_df(
        path = fp_file_path,
        sheet = sheet %||% 1,
        station_name = station_name
      )

      return(as.data.frame(out, stringsAsFactors = FALSE, check.names = FALSE))
    }

    if (!exists(".fp_query_to_field_sheet_df", mode = "function")) {
      stop("Required helper `.fp_query_to_field_sheet_df()` was not found.", call. = FALSE)
    }

    out <- .fp_query_to_field_sheet_df(
      path = fp_file_path,
      sheet = sheet
    )

    as.data.frame(out, stringsAsFactors = FALSE, check.names = FALSE)
  }

  # Resolve local occurrence.txt files or automatically download the required JABOT archive.
  .resolve_occurrence_files <- function(herbarium,
                                        path = NULL,
                                        force_refresh = FALSE,
                                        verbose = FALSE,
                                        resource_map = NULL) {
    if (!exists(".get_ipt_info", mode = "function")) {
      stop("Required helper `.get_ipt_info()` was not found.", call. = FALSE)
    }

    if (!exists(".download_dwca_one", mode = "function")) {
      stop("Required helper `.download_dwca_one()` was not found.", call. = FALSE)
    }

    if (!exists(".herbaria_cache_paths", mode = "function")) {
      stop("Required helper `.herbaria_cache_paths()` was not found.", call. = FALSE)
    }

    if (!is.null(path)) {
      files_existing <- list.files(
        path,
        pattern = "^occurrence\\.txt$",
        recursive = TRUE,
        full.names = TRUE
      )

      if (length(files_existing)) {
        if (isTRUE(verbose)) {
          message("Using existing local occurrence.txt under: ", path)
        }
        return(unique(files_existing))
      }
    }

    cache_info <- .herbaria_cache_paths()
    base_dir <- path %||% file.path(cache_info$cache_dir, "jabot_download")

    if (!dir.exists(base_dir)) {
      dir.create(base_dir, recursive = TRUE, showWarnings = FALSE)
    }

    ipt_info <- .get_ipt_info(
      herbarium = herbarium,
      ipt = "jabot",
      resource_map = resource_map
    )

    if (!nrow(ipt_info)) {
      stop("No JABOT IPT resource was found for herbarium '", herbarium, "'.", call. = FALSE)
    }

    out_dirs <- character(0)

    for (i in seq_len(nrow(ipt_info))) {
      one_dir <- .download_dwca_one(
        info_row = ipt_info[i, , drop = FALSE],
        dir = base_dir,
        verbose = verbose,
        force_refresh = force_refresh
      )

      if (!is.null(one_dir) && dir.exists(one_dir)) {
        out_dirs <- c(out_dirs, one_dir)
      }
    }

    files_downloaded <- unique(unlist(lapply(
      out_dirs,
      function(d) list.files(d, pattern = "^occurrence\\.txt$", recursive = TRUE, full.names = TRUE)
    )))

    if (!length(files_downloaded)) {
      stop(
        "Automatic download ran, but no 'occurrence.txt' files were obtained for herbarium '",
        herbarium, "'.",
        call. = FALSE
      )
    }

    files_downloaded
  }

  input_type <- match.arg(tolower(input_type), c("field_sheet", "fp_query_sheet", "monitora"))

  if (!is.character(fp_file_path) || length(fp_file_path) != 1L || !file.exists(fp_file_path)) {
    stop("`fp_file_path` must be a valid existing file.", call. = FALSE)
  }

  if (!exists(".arg_check_herbarium", mode = "function")) {
    stop("Required helper `.arg_check_herbarium()` was not found.", call. = FALSE)
  }

  .arg_check_herbarium(herbaria)

  herbaria <- unique(toupper(trimws(as.character(herbaria))))
  if (length(herbaria) != 1L) {
    stop("`herbaria` must contain exactly one herbarium code for this function.", call. = FALSE)
  }

  herbarium <- herbaria[1]

  for (pkg in c("DBI", "duckdb")) {
    if (!requireNamespace(pkg, quietly = TRUE)) {
      stop(
        "Package '", pkg, "' is required. Install it with install.packages('", pkg, "').",
        call. = FALSE
      )
    }
  }

  # Resolve or download local JABOT occurrence files.
  files <- .resolve_occurrence_files(
    herbarium = herbarium,
    path = path,
    force_refresh = force_refresh,
    verbose = verbose,
    resource_map = resource_map
  )

  # Read and standardize the inventory.
  inv <- .inventory_to_common_df(
    fp_file_path = fp_file_path,
    input_type = input_type,
    sheet = sheet,
    station_name = station_name
  )

  req_cols <- c("Voucher", "Family", "Original determination")
  missing_cols <- setdiff(req_cols, names(inv))
  if (length(missing_cols)) {
    stop(
      "The standardized inventory is missing required columns: ",
      paste(missing_cols, collapse = ", "),
      call. = FALSE
    )
  }

  # Keep collected rows only when requested.
  if (isTRUE(only_collected) && "Collected" %in% names(inv)) {
    collected_chr <- tolower(trimws(as.character(inv$Collected)))
    voucher_chr <- trimws(as.character(inv$Voucher))
    voucher_chr[is.na(voucher_chr)] <- ""

    keep <- collected_chr %in% c("sim", "s", "yes", "y", "true", "1", "coletado", "coletada") |
      nzchar(voucher_chr)

    inv <- inv[keep, , drop = FALSE]
  }

  if (!nrow(inv)) {
    stop("No inventory rows remained after filtering.", call. = FALSE)
  }

  if (!exists(".parse_census_identity", mode = "function")) {
    stop("Required helper `.parse_census_identity()` was not found.", call. = FALSE)
  }

  # Parse voucher identity into collector and number keys.
  id_parsed <- .parse_census_identity(
    voucher = inv$Voucher,
    collector_fallback = collector_fallback,
    collector_codes = collector_codes
  )

  inv$inventory_family <- .clean_chr_local(inv$Family)
  inv$inventory_species <- .clean_chr_local(inv$`Original determination`)
  inv$inventory_det_canon <- .canon_taxon_compare(inv$inventory_species)
  inv$voucher_is_numeric_only <- grepl("^\\s*\\d+\\s*$", as.character(inv$Voucher))

  inv$collector_raw <- id_parsed$collector_raw
  inv$number_raw <- id_parsed$number_raw
  inv$number_clean <- id_parsed$number_clean
  inv$primary_key <- id_parsed$primary_key
  inv$fallback_key <- id_parsed$fallback_key

  inv_ok <- inv[
    (!is.na(inv$primary_key) & nzchar(inv$primary_key)) |
      (!is.na(inv$fallback_key) & nzchar(inv$fallback_key)),
    ,
    drop = FALSE
  ]

  if (!nrow(inv_ok)) {
    stop(
      "No usable voucher identity keys were found in the inventory. Check voucher formatting or provide `collector_fallback` / `collector_codes`.",
      call. = FALSE
    )
  }

  # Restrict herbarium loading to candidate record numbers only.
  need_primary <- unique(stats::na.omit(inv_ok$primary_key))
  need_fallback <- unique(stats::na.omit(inv_ok$fallback_key))
  all_numbers <- unique(c(
    sub("^.*\\|\\|", "", need_primary),
    sub("^.*\\|\\|", "", need_fallback)
  ))
  all_numbers <- all_numbers[nzchar(all_numbers)]

  dbdir_use <- dbdir
  if (is.null(dbdir_use) && exists(".herbaria_cache_paths", mode = "function")) {
    dbdir_use <- .herbaria_cache_paths()$duckdb_path
  }

  con <- tryCatch(
    DBI::dbConnect(duckdb::duckdb(), dbdir = dbdir_use %||% ":memory:"),
    error = function(e) DBI::dbConnect(duckdb::duckdb(), dbdir = ":memory:")
  )
  on.exit(try(DBI::dbDisconnect(con, shutdown = TRUE), silent = TRUE), add = TRUE)

  try(DBI::dbExecute(con, "PRAGMA enable_progress_bar=false;"), silent = TRUE)

  std <- c(
    "collectionCode","catalogNumber","occurrenceID",
    "recordedBy","recordNumber",
    "family","genus","specificEpithet","infraspecificEpithet","taxonRank",
    "scientificName","acceptedScientificName","verbatimScientificName",
    "identifiedBy","dateIdentified"
  )

  DBI::dbExecute(con, paste0(
    "CREATE OR REPLACE TABLE occ AS SELECT ",
    paste0("CAST(NULL AS VARCHAR) AS ", std, collapse = ", "),
    " WHERE 1=0;"
  ))

  # Detect the best delimiter from the first line.
  delim1 <- function(f) {
    l <- tryCatch(readLines(f, n = 1, warn = FALSE), error = function(e) "")
    if (!nzchar(l)) return("\t")
    ct <- c(
      tab   = lengths(regmatches(l, gregexpr("\t", l, fixed = TRUE))),
      pipe  = lengths(regmatches(l, gregexpr("|",  l, fixed = TRUE))),
      semi  = lengths(regmatches(l, gregexpr(";",  l, fixed = TRUE))),
      comma = lengths(regmatches(l, gregexpr(",",  l, fixed = TRUE)))
    )
    k <- names(ct)[which.max(ct)]
    switch(k, tab = "\t", pipe = "|", semi = ";", comma = ",")
  }

  # Read the header row using UTF-8 first, then latin1.
  header_fields <- function(f, d) {
    x <- tryCatch(readLines(f, n = 1, warn = FALSE, encoding = "UTF-8"),
                  error = function(e) character(0))
    if (!length(x)) {
      x <- tryCatch(readLines(f, n = 1, warn = FALSE, encoding = "latin1"),
                    error = function(e) character(0))
    }
    if (!length(x)) return(character(0))
    strsplit(x, d, fixed = TRUE)[[1]]
  }

  # Select only the standard columns that exist in the current file.
  sel <- function(f, d) {
    h <- header_fields(f, d)
    lk <- setNames(h, tolower(h))
    hs <- names(lk)

    vapply(std, function(x) {
      k <- tolower(x)
      if (k %in% hs) paste0(lk[[k]], " AS ", x) else paste0("NULL AS ", x)
    }, character(1))
  }

  # Insert one occurrence.txt file into DuckDB using permissive CSV settings.
  ins <- function(f, d, fn, encoding, quote, escape) {
    f2 <- gsub("'", "''", normalizePath(f, winslash = "/", mustWork = FALSE))

    common <- paste0(
      "delim='", gsub("\\\\", "\\\\\\\\", d), "', header=TRUE, nullstr='', ",
      "all_varchar=true, strict_mode=false, ignore_errors=true, null_padding=true, ",
      "max_line_size=10000000, encoding='", encoding, "', quote='", quote, "', escape='", escape, "'"
    )

    DBI::dbExecute(con, paste0(
      "INSERT INTO occ SELECT ",
      paste(sel(f, d), collapse = ", "),
      " FROM ", fn, "('", f2, "', ", common,
      if (identical(fn, "read_csv")) ", auto_detect=false" else "",
      ");"
    ))
  }

  # Load all available occurrence files into the temporary DuckDB table.
  for (f in files) {
    d <- delim1(f)
    try(ins(f, d, "read_csv_auto", "utf-8", "\"", "\""), silent = TRUE)
    try(ins(f, d, "read_csv_auto", "latin-1", "\"", "\""), silent = TRUE)
    try(ins(f, d, "read_csv", "utf-8", "\"", "\""), silent = TRUE)
    try(ins(f, d, "read_csv", "latin-1", "\"", "\""), silent = TRUE)
    try(ins(f, d, "read_csv_auto", "utf-8", "", ""), silent = TRUE)
    try(ins(f, d, "read_csv", "utf-8", "", ""), silent = TRUE)
  }

  herb_q <- gsub("'", "''", herbarium)

  # Restrict the working herbarium table to the selected herbarium only.
  DBI::dbExecute(con, paste0("
    CREATE OR REPLACE TABLE occ_selected AS
    SELECT
      collectionCode,
      catalogNumber,
      occurrenceID,
      recordedBy,
      recordNumber,
      family,
      genus,
      specificEpithet,
      infraspecificEpithet,
      taxonRank,
      scientificName,
      acceptedScientificName,
      verbatimScientificName,
      identifiedBy,
      dateIdentified,
      lower(regexp_replace(trim(COALESCE(recordNumber,'')), '[^0-9]', '', 'g')) AS recordNumber_clean
    FROM occ
    WHERE upper(collectionCode) = '", herb_q, "';
  "))

  n_loaded <- DBI::dbGetQuery(con, "SELECT COUNT(*) AS n FROM occ_selected;")$n
  if (!isTRUE(n_loaded > 0)) {
    stop("No records were loaded for herbarium '", herbarium, "'.", call. = FALSE)
  }

  nums_sql <- paste(sprintf("'%s'", gsub("'", "''", all_numbers)), collapse = ", ")

  herb_cand <- DBI::dbGetQuery(con, paste0("
    SELECT
      collectionCode,
      catalogNumber,
      occurrenceID,
      recordedBy,
      recordNumber,
      recordNumber_clean,
      family,
      genus,
      specificEpithet,
      infraspecificEpithet,
      scientificName,
      acceptedScientificName,
      verbatimScientificName,
      identifiedBy,
      dateIdentified
    FROM occ_selected
    WHERE recordNumber_clean IN (", nums_sql, ");
  "))

  if (!nrow(herb_cand)) {
    matches <- data.frame()
    diagnostics <- data.frame(
      diagnostic_class = "no_matches_found",
      n = 0L,
      summary_message = paste0("No inventory voucher keys were found in herbarium ", herbarium, "."),
      stringsAsFactors = FALSE
    )

    unmatched <- inv_ok[, c("Voucher", "inventory_family", "inventory_species", "collector_raw", "number_raw"), drop = FALSE]
    names(unmatched) <- c("voucher", "inventory_family", "inventory_species", "collector", "record_number")

    if (isTRUE(save)) {
      .save_xlsx_three_sheets(matches, diagnostics, unmatched, dir, filename)
    }

    return(invisible(list(
      matches = matches,
      diagnostics = diagnostics,
      unmatched_inventory = unmatched
    )))
  }

  if (!exists(".split_recordedby_people", mode = "function")) {
    stop("Required helper `.split_recordedby_people()` was not found.", call. = FALSE)
  }
  if (!exists(".collector_tokens_one", mode = "function")) {
    stop("Required helper `.collector_tokens_one()` was not found.", call. = FALSE)
  }

  herb_idx_list <- vector("list", nrow(herb_cand))
  pos <- 0L

  # Build herbarium-side primary and fallback collector-number keys.
  for (i in seq_len(nrow(herb_cand))) {
    people <- .split_recordedby_people(herb_cand$recordedBy[i])
    if (!length(people)) next

    tok <- t(vapply(
      people,
      .collector_tokens_one,
      FUN.VALUE = c(primary = NA_character_, fallback = NA_character_)
    ))

    primary_keys <- ifelse(
      !is.na(tok[, "primary"]) & nzchar(tok[, "primary"]) &
        !is.na(herb_cand$recordNumber_clean[i]) & nzchar(herb_cand$recordNumber_clean[i]),
      paste0(tok[, "primary"], "||", herb_cand$recordNumber_clean[i]),
      NA_character_
    )

    fallback_keys <- ifelse(
      !is.na(tok[, "fallback"]) & nzchar(tok[, "fallback"]) &
        !is.na(herb_cand$recordNumber_clean[i]) & nzchar(herb_cand$recordNumber_clean[i]),
      paste0(tok[, "fallback"], "||", herb_cand$recordNumber_clean[i]),
      NA_character_
    )

    one <- rbind(
      data.frame(
        match_key = primary_keys,
        key_type = "primary",
        row_id = i,
        stringsAsFactors = FALSE
      ),
      data.frame(
        match_key = fallback_keys,
        key_type = "fallback",
        row_id = i,
        stringsAsFactors = FALSE
      )
    )

    one <- one[!is.na(one$match_key) & nzchar(one$match_key), , drop = FALSE]
    if (!nrow(one)) next

    pos <- pos + 1L
    herb_idx_list[[pos]] <- unique(one)
  }

  if (!pos) {
    matches <- data.frame()
    diagnostics <- data.frame(
      diagnostic_class = "no_matches_found",
      n = 0L,
      summary_message = paste0("No candidate herbarium records generated usable identity keys in ", herbarium, "."),
      stringsAsFactors = FALSE
    )

    unmatched <- inv_ok[, c("Voucher", "inventory_family", "inventory_species", "collector_raw", "number_raw"), drop = FALSE]
    names(unmatched) <- c("voucher", "inventory_family", "inventory_species", "collector", "record_number")

    if (isTRUE(save)) {
      .save_xlsx_three_sheets(matches, diagnostics, unmatched, dir, filename)
    }

    return(invisible(list(
      matches = matches,
      diagnostics = diagnostics,
      unmatched_inventory = unmatched
    )))
  }

  herb_idx <- do.call(rbind, herb_idx_list[seq_len(pos)])
  herb_idx <- merge(
    herb_idx,
    cbind(row_id = seq_len(nrow(herb_cand)), herb_cand, stringsAsFactors = FALSE),
    by = "row_id",
    all.x = TRUE
  )

  # Use primary keys for all rows and fallback keys only when voucher strings are weak.
  inv_primary <- inv_ok[, c("Voucher", "inventory_family", "inventory_species", "inventory_det_canon",
                            "collector_raw", "number_raw", "primary_key"), drop = FALSE]
  names(inv_primary)[names(inv_primary) == "primary_key"] <- "match_key"
  inv_primary$key_type <- "primary"

  inv_fallback <- inv_ok[
    inv_ok$voucher_is_numeric_only | is.na(inv_ok$collector_raw) | !nzchar(trimws(inv_ok$collector_raw)),
    c("Voucher", "inventory_family", "inventory_species", "inventory_det_canon",
      "collector_raw", "number_raw", "fallback_key"),
    drop = FALSE
  ]
  names(inv_fallback)[names(inv_fallback) == "fallback_key"] <- "match_key"
  inv_fallback$key_type <- "fallback"

  inv_keys <- rbind(inv_primary, inv_fallback)
  inv_keys <- inv_keys[!is.na(inv_keys$match_key) & nzchar(inv_keys$match_key), , drop = FALSE]
  inv_keys <- unique(inv_keys)

  all_matches <- merge(
    inv_keys,
    herb_idx,
    by = c("match_key", "key_type"),
    all = FALSE
  )

  if (nrow(all_matches)) {
    # Prefer primary matches over fallback matches for the same voucher.
    all_matches$key_priority <- ifelse(all_matches$key_type == "primary", 1L, 2L)
    all_matches <- all_matches[order(all_matches$Voucher, all_matches$key_priority), , drop = FALSE]
    all_matches <- all_matches[!duplicated(paste(
      all_matches$Voucher,
      all_matches$catalogNumber,
      all_matches$occurrenceID,
      sep = "||"
    )), , drop = FALSE]

    all_matches$herbarium_family <- .clean_chr_local(all_matches$family)

    # Normalize herbarium-side taxon names from DwC fields.
    all_matches$herbarium_species <- .canon_taxon_from_dwc(
      genus = all_matches$genus,
      specificEpithet = all_matches$specificEpithet,
      infraspecificEpithet = all_matches$infraspecificEpithet,
      scientificName = all_matches$scientificName,
      acceptedScientificName = all_matches$acceptedScientificName,
      verbatimScientificName = all_matches$verbatimScientificName
    )

    all_matches$detby <- .clean_chr_local(all_matches$identifiedBy)
    all_matches$detdate <- .clean_chr_local(all_matches$dateIdentified)

    diag_df <- .diagnose_case(
      inv_canon = all_matches$inventory_det_canon,
      herb_canon = all_matches$herbarium_species
    )

    all_matches <- cbind(all_matches, diag_df)

    # Keep only diagnostically relevant differences in the Matches sheet.
    matches <- all_matches[all_matches$diagnostic_class != "same_determination", , drop = FALSE]

    if (nrow(matches)) {
      matches <- matches[, c(
        "Voucher",
        "inventory_family",
        "inventory_species",
        "herbarium_family",
        "herbarium_species",
        "detby",
        "detdate",
        "diagnostic_class",
        "diagnostic_message",
        "collector_raw",
        "number_raw",
        "collectionCode",
        "catalogNumber",
        "occurrenceID",
        "recordedBy",
        "recordNumber"
      )]

      names(matches) <- c(
        "voucher",
        "inventory_family",
        "inventory_species",
        "herbarium_family",
        "herbarium_species",
        "detby",
        "detdate",
        "diagnostic_class",
        "diagnostic_message",
        "inventory_collector",
        "inventory_record_number",
        "herbarium",
        "catalog_number",
        "occurrence_id",
        "herbarium_recordedBy",
        "herbarium_record_number"
      )

      matches <- unique(matches)
    } else {
      matches <- data.frame(
        voucher = character(0),
        inventory_family = character(0),
        inventory_species = character(0),
        herbarium_family = character(0),
        herbarium_species = character(0),
        detby = character(0),
        detdate = character(0),
        diagnostic_class = character(0),
        diagnostic_message = character(0),
        stringsAsFactors = FALSE
      )
    }
  } else {
    all_matches <- data.frame()
    matches <- data.frame(
      voucher = character(0),
      inventory_family = character(0),
      inventory_species = character(0),
      herbarium_family = character(0),
      herbarium_species = character(0),
      detby = character(0),
      detdate = character(0),
      diagnostic_class = character(0),
      diagnostic_message = character(0),
      stringsAsFactors = FALSE
    )
  }

  # Inventory vouchers that were not matched in the selected herbarium.
  matched_vouchers <- if (nrow(all_matches)) unique(all_matches$Voucher) else character(0)

  unmatched <- inv_ok[!(inv_ok$Voucher %in% matched_vouchers), c(
    "Voucher", "inventory_family", "inventory_species", "collector_raw", "number_raw"
  ), drop = FALSE]

  names(unmatched) <- c(
    "voucher", "inventory_family", "inventory_species", "collector", "record_number"
  )

  # Summarize all match outcomes, including same_determination cases.
  if (nrow(all_matches)) {
    diagnostics <- as.data.frame(table(all_matches$diagnostic_class), stringsAsFactors = FALSE)
    names(diagnostics) <- c("diagnostic_class", "n")

    diagnostics$summary_message <- ifelse(
      diagnostics$diagnostic_class == "inventory_indet_herbarium_det",
      "These are the main target cases: the inventory is still indeterminate, but the selected herbarium already has a determination.",
      ifelse(
        diagnostics$diagnostic_class == "inventory_det_herbarium_indet",
        "The inventory has a determination, but the selected herbarium record is indeterminate or lacks a usable determination.",
        ifelse(
          diagnostics$diagnostic_class == "same_determination",
          "The inventory and herbarium agree after normalization and therefore do not appear in the Matches sheet.",
          "The inventory and herbarium disagree on the determination and should be checked manually."
        )
      )
    )
  } else {
    diagnostics <- data.frame(
      diagnostic_class = "no_matches_found",
      n = 0L,
      summary_message = paste0("No inventory voucher keys were found in herbarium ", herbarium, "."),
      stringsAsFactors = FALSE
    )
  }

  # Save workbook if requested.
  if (isTRUE(save)) {
    .save_xlsx_three_sheets(
      matches = matches,
      diagnostics = diagnostics,
      unmatched = unmatched,
      dir = dir,
      filename = filename
    )
  }

  invisible(list(
    matches = matches,
    diagnostics = diagnostics,
    unmatched_inventory = unmatched
  ))
}
