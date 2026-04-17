#' Organize voucher image directories by updated identification
#'
#' @author Giulia Ottino & Domingos Cardoso
#'
#' @description Organizes image folders of voucher specimens based on their
#' updated taxonomic identification provided in spreadsheet files compatible
#' with the input layouts already supported in the package workflow. The
#' function is designed to be safely rerun multiple times, adapting its
#' behavior according to the existing directory structure and to the selected
#' input type.
#'
#' Internally, the input is harmonized into the same canonical field-sheet
#' schema used elsewhere in the package. For `input_type = "fp_query_sheet"`
#' and `input_type = "monitora"`, the function relies on the existing internal
#' converters. For `input_type = "field_sheet_ti"`, it applies the same direct
#' positional remapping used in the package workflow for Indigenous Land /
#' Panará-style field sheets.
#'
#' On the **first execution**, the function reads voucher information from the
#' input file, (i) creates a main folder (default: `"voucher_imgs"`) if it does
#' not already exist, and (ii) for each voucher treated as collected, creates a
#' nested directory structure following the format `Family/Genus/Voucher`, where
#' voucher images can be stored.
#'
#' On **subsequent executions**, the function performs a reorganization step.
#' It scans the existing voucher directories inside `output_dir`, compares them
#' with the currently accepted taxonomic placement inferred from the harmonized
#' data, and identifies folders that may reflect outdated or incorrect genus
#' assignments. If a voucher is found inside a taxonomically outdated path, the
#' function attempts to move the files to the correct `Family/Genus/Voucher`
#' directory. If the destination folder already contains files, the function
#' avoids overwriting existing content and only moves files that are not already
#' present.
#'
#' Once files are moved (or verified), the function removes empty outdated
#' voucher directories and, when possible, also removes empty parent taxonomic
#' folders left behind by the reorganization. This ensures that the directory
#' tree remains synchronized with the most up-to-date taxonomic identification,
#' making voucher image organization robust, repeatable, and easy to maintain
#' over time.
#'
#' A voucher is treated as collected when the canonical `Collected` field
#' indicates a positive collection status (for example `"Sim"`, `"Yes"`,
#' `"Collected"`, `"Coletado"`, `"1"`, `"x"`, or similar values). As an
#' additional fallback, the function also treats a record as collected when a
#' non-empty voucher code is present after input harmonization.
#'
#' @param fp_file_path Character. Path to the input Excel file containing data
#'   in one of the supported plot-data layouts.
#' @param output_dir Character. Root directory in which voucher image folders
#'   will be organized. Default is `"voucher_imgs"`.
#' @param input_type Character. Input layout. Must be one of `"field_sheet"`,
#'   `"field_sheet_ti"`, `"fp_query_sheet"`, or `"monitora"`.
#' @param station_name Optional character or numeric. Station name or station
#'   identifier used to filter MONITORA spreadsheets when applicable. Ignored
#'   for the other input types.
#'
#' @return Invisibly returns the harmonized field-sheet data frame used
#'   internally to organize the voucher directories.
#'
#' @examples
#' \dontrun{
#' # Standard field sheet
#' mk_voucher_dirs(
#'   fp_file_path = "data/RUS_plot.xlsx",
#'   input_type = "field_sheet",
#'   output_dir = "voucher_imgs"
#' )
#'
#' # Panará / TI field sheet
#' mk_voucher_dirs(
#'   fp_file_path = "data/SKR_01.xlsx",
#'   input_type = "field_sheet_ti",
#'   output_dir = "voucher_imgs_skr"
#' )
#'
#' # ForestPlots Query export
#' mk_voucher_dirs(
#'   fp_file_path = "data/query_export.xlsx",
#'   input_type = "fp_query_sheet",
#'   output_dir = "voucher_imgs_query"
#' )
#'
#' # MONITORA spreadsheet
#' mk_voucher_dirs(
#'   fp_file_path = "data/Monitora_Plantas_PND_2024_validado_subir_no_processo.xlsx",
#'   input_type = "monitora",
#'   output_dir = "voucher_imgs_descobrimento",
#'   station_name = 1
#' )
#' }
#'
#' @export

mk_voucher_dirs <- function(fp_file_path = NULL,
                            output_dir = "voucher_imgs",
                            input_type = c("field_sheet",
                                           "field_sheet_ti",
                                           "fp_query_sheet",
                                           "monitora"),
                            station_name = NULL) {

  input_type <- match.arg(input_type)

  if (is.null(fp_file_path) || !file.exists(fp_file_path)) {
    stop("Please provide a valid Excel file path.", call. = FALSE)
  }

  output_dir <- .arg_check_dir(output_dir)

  fp_sheet <- switch(
    input_type,
    "monitora" = .monitora_to_field_sheet_df(
      path = fp_file_path,
      station_name = station_name
    ),

    "fp_query_sheet" = .fp_query_to_field_sheet_df(
      path = fp_file_path
    ),

    "field_sheet" = {
      raw <- suppressMessages(
        readxl::read_excel(
          fp_file_path,
          sheet = 1,
          col_names = FALSE,
          .name_repair = "minimal"
        )
      )

      raw <- as.data.frame(raw, stringsAsFactors = FALSE)

      if (nrow(raw) < 2) {
        stop("The field sheet must contain at least two rows.", call. = FALSE)
      }

      header <- as.character(unlist(raw[2, ]))
      header <- trimws(header)
      header[is.na(header) | header == ""] <- paste0(
        "unnamed_",
        seq_along(header)
      )[is.na(header) | header == ""]

      names(raw) <- make.unique(header)

      out <- raw[-c(1, 2), , drop = FALSE]
      out <- as.data.frame(out, stringsAsFactors = FALSE)
      names(out) <- trimws(names(out))

      out
    },

    "field_sheet_ti" = {
      raw <- suppressMessages(
        readxl::read_excel(
          fp_file_path,
          sheet = 1,
          col_names = FALSE,
          .name_repair = "minimal"
        )
      )

      raw <- as.data.frame(raw, stringsAsFactors = FALSE)

      if (nrow(raw) < 3) {
        stop("The field_sheet_ti input must contain metadata, header, and data rows.",
             call. = FALSE)
      }

      raw_data <- raw[-c(1, 2), , drop = FALSE]

      dest_cols <- .field_sheet_cols()

      temp <- as.data.frame(
        matrix(NA, nrow = nrow(raw_data), ncol = length(dest_cols)),
        stringsAsFactors = FALSE
      )
      names(temp) <- dest_cols

      temp$`New Tag No` <- as.character(raw_data[[1]])
      temp$`New Stem Grouping` <- as.character(raw_data[[2]])
      temp$T1 <- suppressWarnings(as.numeric(raw_data[[3]]))
      temp$T2 <- suppressWarnings(as.numeric(raw_data[[4]]))
      temp$X <- suppressWarnings(as.numeric(raw_data[[5]]))
      temp$Y <- suppressWarnings(as.numeric(raw_data[[6]]))
      temp$Family <- as.character(raw_data[[7]])

      genus_part <- as.character(raw_data[[8]])
      species_part <- as.character(raw_data[[9]])
      genus_part[is.na(genus_part)] <- ""
      species_part[is.na(species_part)] <- ""

      temp$`Original determination` <- trimws(
        ifelse(
          nzchar(species_part),
          paste(genus_part, species_part),
          genus_part
        )
      )

      temp$Morphospecies <- as.character(raw_data[[10]])
      temp$D <- suppressWarnings(as.numeric(raw_data[[11]]))
      temp$POM <- suppressWarnings(as.numeric(raw_data[[12]]))
      temp$ExtraD <- NA_character_
      temp$ExtraPOM <- NA_character_
      temp$Flag1 <- as.character(raw_data[[13]])
      temp$Flag2 <- NA_character_
      temp$Flag3 <- NA_character_
      temp$LI <- NA_character_
      temp$CI <- NA_character_
      temp$CF <- NA_character_
      temp$CD1 <- NA_character_
      temp$nrdups <- as.character(raw_data[[14]])
      temp$Height <- suppressWarnings(as.numeric(raw_data[[15]]))
      temp$Voucher <- as.character(raw_data[[16]])
      temp$Silica <- as.character(raw_data[[17]])
      temp$Collected <- as.character(raw_data[[18]])
      temp$`Census Notes` <- as.character(raw_data[[19]])
      temp$CAP <- suppressWarnings(as.numeric(raw_data[[21]]))
      temp$`Basal Area` <- suppressWarnings(as.numeric(raw_data[[22]]))

      temp
    }
  )

  fp_sheet <- as.data.frame(fp_sheet, stringsAsFactors = FALSE)

  needed_cols <- c("Family", "Original determination", "Voucher", "Collected")
  missing_cols <- setdiff(needed_cols, names(fp_sheet))
  if (length(missing_cols)) {
    stop(
      "The harmonized input is missing required columns: ",
      paste(missing_cols, collapse = ", "),
      call. = FALSE
    )
  }

  fp_sheet$Family <- as.character(fp_sheet$Family)
  fp_sheet$Voucher <- as.character(fp_sheet$Voucher)
  fp_sheet$Collected <- as.character(fp_sheet$Collected)
  fp_sheet$`Original determination` <- as.character(fp_sheet$`Original determination`)

  has_tag <- !is.na(fp_sheet$`New Tag No`)
  is_yes <- !is.na(fp_sheet$Collected) & tolower(trimws(fp_sheet$Collected)) == "yes"
  has_voucher <- !is.na(fp_sheet$Voucher) & trimws(fp_sheet$Voucher) != ""
  rows <- has_tag & is_yes
  fp_sheet$Voucher[rows] <- ifelse(
    has_voucher[rows],
    paste0(fp_sheet$Voucher[rows], " Tree#", fp_sheet$`New Tag No`[rows]),
    paste0("Tree#", fp_sheet$`New Tag No`[rows])
  )

  if (!dir.exists(output_dir)) {
    dir.create(output_dir, recursive = TRUE)
    message("Created main output directory: ", output_dir)
  }

  existing_dirs <- list.dirs(output_dir, recursive = TRUE, full.names = TRUE)
  existing_voucher_dirs <- existing_dirs[basename(existing_dirs) != basename(output_dir)]

  n_created <- 0L
  n_moved <- 0L
  n_skipped_not_collected <- 0L
  n_skipped_missing_data <- 0L
  old_voucher_dirs_checked <- character(0)

  for (i in seq_len(nrow(fp_sheet))) {
    voucher <- .clean_path_component(fp_sheet$Voucher[i])
    collected_raw <- fp_sheet$Collected[i]
    family <- .clean_path_component(fp_sheet$Family[i])
    genus_correct <- .clean_path_component(.extract_genus(fp_sheet$`Original determination`[i]))

    if (!nzchar(voucher)) {
      n_skipped_missing_data <- n_skipped_missing_data + 1L
      next
    }

    if (!.is_collected_value(collected_raw, voucher = voucher)) {
      n_skipped_not_collected <- n_skipped_not_collected + 1L
      next
    }

    if (!nzchar(family) || !nzchar(genus_correct)) {
      n_skipped_missing_data <- n_skipped_missing_data + 1L
      next
    }

    correct_path <- file.path(output_dir, family, genus_correct, voucher)
    correct_path_norm <- normalizePath(correct_path, winslash = "/", mustWork = FALSE)

    voucher_matches <- existing_voucher_dirs[basename(existing_voucher_dirs) == voucher]

    if (length(voucher_matches) > 0) {
      for (vm in voucher_matches) {
        vm_norm <- normalizePath(vm, winslash = "/", mustWork = FALSE)

        if (!identical(vm_norm, correct_path_norm)) {
          photos <- list.files(vm, full.names = TRUE, recursive = FALSE, all.files = FALSE)

          if (length(photos) > 0) {
            if (!dir.exists(correct_path)) {
              dir.create(correct_path, recursive = TRUE)
            }

            existing_dest_files <- basename(list.files(correct_path, full.names = FALSE))

            for (photo in photos) {
              dest_file <- file.path(correct_path, basename(photo))

              if (basename(photo) %in% existing_dest_files) {
                next
              }

              moved <- file.rename(photo, dest_file)

              if (!moved) {
                copied <- file.copy(photo, dest_file, overwrite = FALSE)
                if (isTRUE(copied)) {
                  unlink(photo)
                  n_moved <- n_moved + 1L
                }
              } else {
                n_moved <- n_moved + 1L
              }
            }
          }

          old_voucher_dirs_checked <- c(old_voucher_dirs_checked, vm)
        }
      }
    }

    if (!dir.exists(correct_path)) {
      dir.create(correct_path, recursive = TRUE)
      n_created <- n_created + 1L
    }
  }

  for (old_voucher_dir in unique(old_voucher_dirs_checked)) {
    if (dir.exists(old_voucher_dir) &&
        length(list.files(old_voucher_dir, all.files = FALSE, no.. = TRUE)) == 0) {
      unlink(old_voucher_dir, recursive = TRUE)
      .delete_empty_parents(dirname(old_voucher_dir), stop_at = output_dir)
    }
  }

  message("Folders created: ", n_created)
  message("Files moved: ", n_moved)
  message("Rows skipped (not collected): ", n_skipped_not_collected)
  message("Rows skipped (missing voucher/family/species): ", n_skipped_missing_data)

  invisible(fp_sheet)
}

.clean_path_component <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- ""
  x <- trimws(x)
  x <- gsub("[/\\:*?\"<>|]", "_", x)
  x
}

.norm_txt <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- ""
  x2 <- suppressWarnings(iconv(x, from = "", to = "ASCII//TRANSLIT", sub = ""))
  bad <- is.na(x2) | !nzchar(x2)
  x2[bad] <- x[bad]
  tolower(trimws(x2))
}

.extract_genus <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- ""
  x <- trimws(x)
  x[x %in% c("", "NA", "Na", "N/A")] <- NA_character_
  genus <- sub("\\s+.*$", "", x)
  genus[is.na(x)] <- NA_character_
  genus
}

.is_collected_value <- function(x, voucher = NULL) {
  z <- .norm_txt(x)

  out <- z %in% c(
    "sim", "s", "yes", "y", "true", "1",
    "x", "coletado", "coletada", "collected"
  )

  if (!is.null(voucher)) {
    v <- as.character(voucher)
    v[is.na(v)] <- ""
    out <- out | nzchar(trimws(v))
  }

  out
}

.delete_empty_parents <- function(path, stop_at) {
  cur <- normalizePath(path, winslash = "/", mustWork = FALSE)
  stop_at <- normalizePath(stop_at, winslash = "/", mustWork = FALSE)

  while (!identical(cur, stop_at) && dir.exists(cur)) {
    if (length(list.files(cur, all.files = FALSE, no.. = TRUE)) == 0) {
      unlink(cur, recursive = TRUE)
      cur <- dirname(cur)
    } else {
      break
    }
  }
}

