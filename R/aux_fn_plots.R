# =============================================================================
# Plot ingestion and geometry helpers
# Author: Giulia Ottino & Domingos Cardoso
# Revised: 08/Mar/2026
# =============================================================================

#' Normalize names for matching
#'
#' @param x Character vector.
#'
#' @return Character vector normalized to lowercase ASCII alphanumeric text.
#'
#' @keywords internal
#' @noRd
.norm_nm <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- ""

  x2 <- suppressWarnings(iconv(x, from = "", to = "ASCII//TRANSLIT", sub = ""))
  bad <- is.na(x2) | !nzchar(x2)
  x2[bad] <- x[bad]

  x2 <- tolower(x2)
  gsub("[^a-z0-9]+", "", x2)
}


#' Pick the first matching column name from candidate aliases
#'
#' @param df Data frame.
#' @param candidates Character vector of candidate names.
#'
#' @return Character scalar or `NA_character_`.
#'
#' @keywords internal
#' @noRd
.pick_colname <- function(df, candidates) {
  cn_raw <- names(df)
  cn <- .norm_nm(cn_raw)
  cand <- .norm_nm(candidates)

  idx <- match(cand, cn, nomatch = 0L)
  if (any(idx > 0L)) {
    return(cn_raw[idx[idx > 0L][1]])
  }

  for (z in cand) {
    if (!nzchar(z) || nchar(z) < 2L) next
    hit <- which(grepl(z, cn, fixed = TRUE))
    if (length(hit) == 1L) {
      return(cn_raw[hit[1]])
    }
  }

  NA_character_
}


#' Test whether any alias exists in a data frame
#'
#' @param df Data frame.
#' @param candidates Character vector.
#'
#' @return Logical scalar.
#'
#' @keywords internal
#' @noRd
.has_any <- function(df, candidates) {
  !is.na(.pick_colname(df, candidates))
}


#' Parse numeric text robustly
#'
#' @param x Vector.
#'
#' @return Numeric vector.
#'
#' @keywords internal
#' @noRd
.parse_num <- function(x) {
  z <- as.character(x)
  z[is.na(z)] <- ""

  z <- gsub(",", ".", z, fixed = TRUE)
  z <- trimws(z)

  z <- gsub("\\s+", "", z)

  z[grepl("/", z, fixed = TRUE)] <- NA_character_

  suppressWarnings(as.numeric(z))
}


#' Choose the best numeric candidate column by completeness
#'
#' @param df Data frame.
#' @param candidates Character vector or list of character vectors.
#' @param default Default numeric value.
#'
#' @return Vector of length `nrow(df)`.
#'
#' @keywords internal
#' @noRd
.best_numeric_vec <- function(df, candidates, default = NA_real_) {
  cn <- unique(stats::na.omit(vapply(
    candidates,
    function(cc) .pick_colname(df, cc),
    FUN.VALUE = character(1)
  )))

  if (!length(cn)) {
    return(rep(default, nrow(df)))
  }

  best <- cn[1]
  best_n <- -1L

  for (cl in cn) {
    v <- .parse_num(df[[cl]])
    n_ok <- sum(is.finite(v), na.rm = TRUE)
    if (n_ok > best_n) {
      best_n <- n_ok
      best <- cl
    }
  }

  .parse_num(df[[best]])
}


#' Return a cleaned text vector
#'
#' @param x Vector.
#'
#' @return Character vector.
#'
#' @keywords internal
#' @noRd
.clean_chr <- function(x) {
  x <- as.character(x)
  x <- iconv(x, to = "ASCII//TRANSLIT", sub = "")
  x[is.na(x)] <- ""
  x <- gsub(intToUtf8(0x00A0), " ", x, fixed = TRUE)
  x <- gsub("\\s+", " ", x)
  trimws(x)
}


#' Parse year values from mixed text
#'
#' @param x Vector.
#'
#' @return Integer vector.
#'
#' @keywords internal
#' @noRd
.parse_year <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- ""
  m <- regexpr("(?:19|20)\\d{2}", x, perl = TRUE)
  out <- rep(NA_integer_, length(x))
  ok <- m > 0L
  out[ok] <- suppressWarnings(as.integer(substr(
    x[ok],
    m[ok],
    m[ok] + attr(m, "match.length")[ok] - 1L
  )))
  out
}


#' Normalize station names
#'
#' @param x Vector.
#'
#' @return Character vector.
#'
#' @keywords internal
#' @noRd
.normalize_station_name <- function(x) {
  x <- .clean_chr(x)
  tolower(x)
}


#' Normalize station numeric identifiers
#'
#' @param x Vector.
#'
#' @return Character vector.
#'
#' @keywords internal
#' @noRd
.normalize_station_number <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- ""
  x <- trimws(x)
  x <- gsub("\\s+", "", x)
  x <- sub("^0+([0-9])", "\\1", x)
  x[!nzchar(x)] <- NA_character_
  x
}

#' Standard field-sheet schema used internally by plot functions
#'
#' @return Character vector of canonical column names.
#'
#' @keywords internal
#' @noRd
.field_sheet_cols <- function() {
  c(
    "New Tag No", "New Stem Grouping", "T1", "T2", "X", "Y", "Family",
    "Original determination", "Morphospecies", "D", "POM", "ExtraD", "ExtraPOM",
    "Flag1", "Flag2", "Flag3", "LI", "CI", "CF", "CD1", "nrdups", "Height",
    "Voucher", "Silica", "Collected", "Census Notes", "CAP", "Basal Area"
  )
}


#' Score workbook sheets for plot data detection
#'
#' @param df Data frame.
#'
#' @return Integer score.
#'
#' @keywords internal
#' @noRd
.score_plot_sheet <- function(df) {
  dest_cols <- .field_sheet_cols()
  s <- 0L

  if (all(dest_cols %in% names(df))) {
    s <- s + 50L
  }

  if (.has_any(df, c("Tag No", "T1", "X", "Y")) &&
      .has_any(df, c("Voucher Code")) &&
      .has_any(df, c("Recommended Species", "Recommended Voucher Species"))) {
    s <- s + 20L
  }

  if (.has_any(df, c("Voucher Code")) && .has_any(df, c("TagNumber"))) {
    s <- s + 10L
  }

  if (.has_any(df, c("Plot Code")) &&
      .has_any(df, c("Tag No")) &&
      .has_any(df, c("Sub Plot T1", "Standardised SubPlot T1", "Standardized SubPlot T1")) &&
      .has_any(df, c("X", "Y", "Standardised X", "Standardised Y", "Standardized X", "Standardized Y"))) {
    s <- s + 40L
  }

  if (.has_any(df, c("Nome_estacao", "Nome Estacao", "Estacao", "N_estacao")) &&
      .has_any(df, c("N_arvore", "Tag", "Tag No")) &&
      .has_any(df, c("X(m)", "Y(m)", "X", "Y"))) {
    s <- s + 40L
  }

  s + as.integer(min(ncol(df), 200) / 10)
}


#' Read the best worksheet from a workbook
#'
#' @param path File path.
#' @param sheet Optional preferred sheet name or index.
#'
#' @return Data frame with `sheet_name` attribute.
#'
#' @keywords internal
#' @noRd
.read_best_plot_sheet <- function(path, sheet = NULL) {
  if (!grepl("\\.xlsx?$", path, ignore.case = TRUE)) {
    stop("This helper expects an Excel workbook path.", call. = FALSE)
  }

  sheets <- readxl::excel_sheets(path)

  if (!is.null(sheet) && nzchar(trimws(sheet))) {
    if (!sheet %in% sheets) {
      stop(
        paste0("Requested sheet '", sheet, "' was not found in workbook."),
        call. = FALSE
      )
    }
    pref <- sheet
  } else {
    pref <- unique(c("Data", "Plot Dump", sheets))
    pref <- pref[pref %in% sheets]
    if (!length(pref)) pref <- sheets
  }

  best_df <- NULL
  best_sheet <- NULL
  best_score <- -Inf

  for (s in pref) {
    temp0 <- suppressMessages(
      readxl::read_excel(
        path,
        sheet = s,
        col_names = FALSE,
        .name_repair = "minimal"
      )
    )

    if (!is.data.frame(temp0) || !nrow(temp0) || ncol(temp0) < 3L) next

    max_r <- min(15L, nrow(temp0))

    known_headers <- c(
      "Tag No", "New Tag No", "Tree ID", "TreeID", "TagNumber",
      "T1", "Sub Plot T1", "SubPlotT1", "Standardised SubPlot T1", "Standardized SubPlot T1",
      "T2", "Sub Plot T2", "SubPlotT2",
      "X", "Y", "Standardised X", "Standardized X", "Standardised Y", "Standardized Y",
      "Voucher Code", "Voucher Collected",
      "Recommended Voucher Species", "Recommended Species",
      "Family", "Recommended Family", "Recommended Voucher Family",
      "D", "D1", "DBH", "CAP", "Basal Area"
    )

    scores <- vapply(seq_len(max_r), function(i) {
      row_i <- as.character(unlist(temp0[i, , drop = TRUE]))
      row_i[is.na(row_i)] <- ""
      row_i <- trimws(row_i)
      sum(known_headers %in% row_i)
    }, numeric(1))

    header_row_idx <- which.max(scores)

    if (!length(header_row_idx) || scores[header_row_idx] == 0) next

    header_bottom <- as.character(unlist(temp0[header_row_idx, , drop = TRUE]))
    header_bottom[is.na(header_bottom)] <- ""
    header_bottom <- trimws(header_bottom)

    header_top <- rep("", length(header_bottom))
    if (header_row_idx > 1) {
      header_top <- as.character(unlist(temp0[header_row_idx - 1, , drop = TRUE]))
      header_top[is.na(header_top)] <- ""
      header_top <- trimws(header_top)
    }

    # combina cabecalho de duas linhas sempre que a linha de cima tiver conteudo util
    use_combined_header <- sum(nzchar(header_top)) >= 2L

    header <- if (use_combined_header) {
      ifelse(
        nzchar(header_top) & nzchar(header_bottom),
        paste(header_top, header_bottom),
        ifelse(nzchar(header_bottom), header_bottom, header_top)
      )
    } else {
      header_bottom
    }

    header[is.na(header)] <- ""
    header <- trimws(header)
    header[header == ""] <- paste0("unnamed_", seq_along(header))[header == ""]

    temp <- temp0[-seq_len(header_row_idx), , drop = FALSE]
    names(temp) <- make.unique(header)
    temp <- as.data.frame(temp, stringsAsFactors = FALSE)

    keep_cols <- colSums(!is.na(temp) & trimws(as.character(as.matrix(temp))) != "") > 0
    temp <- temp[, keep_cols, drop = FALSE]
    names(temp) <- trimws(names(temp))

    sc <- .score_plot_sheet(temp)

    if (sc > best_score) {
      best_score <- sc
      best_df <- temp
      best_sheet <- s
    }
  }

  if (is.null(best_df)) {
    stop("No data found in any worksheet.", call. = FALSE)
  }

  attr(best_df, "sheet_name") <- best_sheet
  best_df
}


#' Convert FP Query / field export to canonical field-sheet-like object
#'
#' @param path Excel file path.
#' @param sheet Optional preferred sheet.
#' @param plot_code Optional plot code filter for workbooks with multiple plots.
#' @param split_plots Logical; if TRUE and multiple plots are found, return a named list.
#'
#' @return Tibble in field-sheet-like format, or a named list of such tibbles.
#'
#' @keywords internal
#' @noRd
.fp_query_to_field_sheet_df <- function(path,
                                        sheet = NULL,
                                        plot_code = NULL,
                                        split_plots = FALSE) {
  if (!file.exists(path)) {
    stop("Input file does not exist.", call. = FALSE)
  }

  dest_cols <- c(
    "New Tag No","New Stem Grouping","T1","T2","X","Y","Family",
    "Original determination","Morphospecies","D","POM","ExtraD","ExtraPOM",
    "Flag1","Flag2","Flag3","LI","CI","CF","CD1","nrdups","Height",
    "Voucher","Silica","Collected","Census Notes","CAP","Basal Area"
  )

  in_dat <- .read_best_plot_sheet(path, sheet = sheet)
  sheet_used <- attr(in_dat, "sheet_name")
  names(in_dat) <- trimws(names(in_dat))

  plot_meta <- list(
    plot_code = .meta_value(in_dat, c("Plot Code", "PlotCode", "Plotcode")),
    plot_name = .meta_value(in_dat, c("Plot Name", "PlotName")),
    team = .meta_value(in_dat, c("PI", "Team", "Collectors", "Collector")),
    census_no_fp = .meta_value(in_dat, c("Census No"))
  )

  if (nzchar(plot_meta$plot_code) && !grepl("-", plot_meta$plot_code)) {
    plot_meta$plot_code <- gsub("(?<=[A-Z])(?=\\d)", "-", plot_meta$plot_code, perl = TRUE)
  }

  is_field_sheet <- all(dest_cols %in% names(in_dat))
  if (is_field_sheet) {
    out <- in_dat[, dest_cols, drop = FALSE]
    attr(out, "coord_mode") <- "local"
    attr(out, "sheet_name") <- sheet_used
    attr(out, "plot_meta") <- plot_meta
    return(out)
  }

  plot_code_col <- .pick_first_existing(in_dat, c("Plot Code", "PlotCode", "Plotcode"))
  if (is.na(plot_code_col)) {
    plot_code_col <- .pick_colname(in_dat, c("Plot Code", "PlotCode", "Plotcode"))
  }

  if (!is.na(plot_code_col)) {
    pc <- trimws(as.character(in_dat[[plot_code_col]]))
    pc <- pc[!is.na(pc) & nzchar(pc)]
    plot_codes <- unique(pc)
  } else {
    plot_codes <- character(0)
  }

  if (!is.null(plot_code) && nzchar(trimws(plot_code)) && !is.na(plot_code_col)) {
    keep <- trimws(as.character(in_dat[[plot_code_col]])) == trimws(as.character(plot_code))
    in_dat <- in_dat[keep, , drop = FALSE]
    if (!nrow(in_dat)) {
      stop("Requested `plot_code` was not found in the detected sheet.", call. = FALSE)
    }
  } else if (isTRUE(split_plots) && length(plot_codes) > 1L && !is.na(plot_code_col)) {
    out_list <- setNames(vector("list", length(plot_codes)), plot_codes)
    for (pc in plot_codes) {
      out_list[[pc]] <- .fp_query_to_field_sheet_df(
        path = path,
        sheet = sheet_used,
        plot_code = pc,
        split_plots = FALSE
      )
    }
    return(out_list)
  }

  tag_col <- .pick_first_existing(in_dat, c(
    "New Tag No", "Tag No", "TagNo", "TagNumber",
    "Main Stem Tag", "Pv. Tag No", "Tree ID", "TreeID"
  ))
  if (is.na(tag_col)) {
    tag_col <- .pick_colname(in_dat, c(
      "New Tag No", "Tag No", "TagNo", "TagNumber",
      "Main Stem Tag", "Pv. Tag No", "Tree ID", "TreeID"
    ))
  }

  stem_group_col <- .pick_first_existing(in_dat, c(
    "New Stem Grouping", "Stem Group ID", "StemGroupID"
  ))
  if (is.na(stem_group_col)) {
    stem_group_col <- .pick_colname(in_dat, c(
      "New Stem Grouping", "Stem Group ID", "StemGroupID"
    ))
  }
  t1_col <- .pick_first_existing(in_dat, c("Sub Plot T1", "T1", "SubPlotT1"))
  t2_col <- .pick_first_existing(in_dat, c("Sub Plot T2", "T2", "SubPlotT2"))
  x_col  <- .pick_first_existing(in_dat, c("X"))
  y_col  <- .pick_first_existing(in_dat, c("Y"))

  std_t1_col <- .pick_first_existing(in_dat, c("Standardised SubPlot T1", "Standardized SubPlot T1"))
  std_x_col  <- .pick_first_existing(in_dat, c("Standardised X", "Standardized X"))
  std_y_col  <- .pick_first_existing(in_dat, c("Standardised Y", "Standardized Y"))

  if (is.na(t1_col))  t1_col  <- .pick_colname(in_dat, c("Sub Plot T1", "T1", "SubPlotT1"))
  if (is.na(t2_col))  t2_col  <- .pick_colname(in_dat, c("Sub Plot T2", "T2", "SubPlotT2"))
  if (is.na(x_col))   x_col   <- .pick_colname(in_dat, c("X"))
  if (is.na(y_col))   y_col   <- .pick_colname(in_dat, c("Y"))
  if (is.na(std_t1_col)) std_t1_col <- .pick_colname(in_dat, c("Standardised SubPlot T1", "Standardized SubPlot T1"))
  if (is.na(std_x_col))  std_x_col  <- .pick_colname(in_dat, c("Standardised X", "Standardized X"))
  if (is.na(std_y_col))  std_y_col  <- .pick_colname(in_dat, c("Standardised Y", "Standardized Y"))

  has_local <- !is.na(tag_col) && !is.na(t1_col) && !is.na(x_col) && !is.na(y_col)
  has_std   <- !is.na(tag_col) && !is.na(std_t1_col) && !is.na(std_x_col) && !is.na(std_y_col)

  if (!(has_local || has_std)) {
    stop(
      paste0(
        "Could not detect FP query coordinates correctly. ",
        "tag_col=", tag_col,
        "; t1_col=", t1_col,
        "; t2_col=", t2_col,
        "; x_col=", x_col,
        "; y_col=", y_col,
        "; std_t1_col=", std_t1_col,
        "; std_x_col=", std_x_col,
        "; std_y_col=", std_y_col
      ),
      call. = FALSE
    )
  }

  if (has_local) {
    t1 <- .parse_num_safe(in_dat[[t1_col]])
    t2 <- if (!is.na(t2_col)) .parse_num_safe(in_dat[[t2_col]]) else rep(NA_real_, nrow(in_dat))
    x  <- .parse_num_safe(in_dat[[x_col]])
    y  <- .parse_num_safe(in_dat[[y_col]])
    coord_mode <- "local"
  } else {
    t1 <- .parse_num_safe(in_dat[[std_t1_col]])
    t2 <- if (!is.na(t2_col)) .parse_num_safe(in_dat[[t2_col]]) else rep(NA_real_, nrow(in_dat))
    x  <- .parse_num_safe(in_dat[[std_x_col]])
    y  <- .parse_num_safe(in_dat[[std_y_col]])
    coord_mode <- "standardised"
  }

  # se T2 vier todo vazio no query export, deixa NA
  if (all(is.na(t2))) {
    t2 <- rep(NA_real_, length(t1))
  }

  if (all(is.na(x)) || all(is.na(y))) {
    stop(
      paste0(
        "FP query conversion produced X/Y entirely as NA. ",
        "Detected columns: X='", ifelse(coord_mode == "local", x_col, std_x_col),
        "', Y='", ifelse(coord_mode == "local", y_col, std_y_col), "'."
      ),
      call. = FALSE
    )
  }

  same_xy <- mean(is.finite(x) & is.finite(y) & x == y, na.rm = TRUE)
  if (is.finite(same_xy) && same_xy > 0.95) {
    warning(
      paste0(
        "The selected X/Y columns are suspiciously identical. ",
        "Chosen columns: T1='", ifelse(coord_mode == "local", t1_col, std_t1_col),
        "', T2='", ifelse(is.na(t2_col), "NA", t2_col),
        "', X='", ifelse(coord_mode == "local", x_col, std_x_col),
        "', Y='", ifelse(coord_mode == "local", y_col, std_y_col), "'."
      ),
      call. = FALSE
    )
  }

  voucher_vals <- .get_col(in_dat, c("Voucher", "Voucher Code"))
  collected_vals <- .get_col(in_dat, c("Collected", "Voucher Collected"))

  voucher_chr <- trimws(as.character(voucher_vals))
  collected_chr <- trimws(as.character(collected_vals))

  voucher_chr[is.na(voucher_chr)] <- ""
  collected_chr[is.na(collected_chr)] <- ""

  collected_final <- collected_chr
  collected_final[voucher_chr != "" & collected_chr == ""] <- "yes"

  out <- tibble::tibble(
    `New Tag No` = as.character(in_dat[[tag_col]]),
    `New Stem Grouping` = if (!is.na(stem_group_col)) as.character(in_dat[[stem_group_col]]) else NA_character_,
    T1 = t1,
    T2 = t2,
    X = x,
    Y = y,
    Family = .get_col(in_dat, c("Recommended Voucher Family", "Recommended Family", "Family")),
    `Original determination` = .get_col(in_dat, c(
      "Recommended Voucher Species",
      "Recommended Species",
      "Original Identification",
      "Species"
    )),
    Morphospecies = .get_col(in_dat, c("Morphospecies", "MorphoSpecies", "Morpho"), default = NA),
    D = .get_col(in_dat, c("D", "D1", "DBH", "Dbh", "D0")),
    POM = .get_col(in_dat, c("POM", "POM0", "DPOMtMinus1")),
    ExtraD = .get_col(in_dat, c("Extra D", "ExtraD", "Extra D0")),
    ExtraPOM = .get_col(in_dat, c("Extra POM", "ExtraPOM", "Extra POM0")),
    Flag1 = .get_col(in_dat, c("Flag1", "F1")),
    Flag2 = .get_col(in_dat, c("Flag2", "F2")),
    Flag3 = .get_col(in_dat, c("Flag3", "F3")),
    LI = .get_col(in_dat, c("LI")),
    CI = .get_col(in_dat, c("CI")),
    CF = .get_col(in_dat, c("CF")),
    CD1 = .get_col(in_dat, c("CD1")),
    nrdups = .get_col(in_dat, c("nrdups", "Stem Count")),
    Height = .get_col(in_dat, c("Height", "Ht")),
    Voucher = voucher_vals,
    Silica = .get_col(in_dat, c("Silica")),
    Collected = collected_final,
    `Census Notes` = .get_col(in_dat, c("Census Notes", "Tree Notes", "Determination Comments", "Comments")),
    CAP = .get_col(in_dat, c("CAP")),
    `Basal Area` = .get_col(in_dat, c("Basal Area", "BA"))
  )
  for (col in dest_cols) {
    if (!col %in% names(out)) out[[col]] <- NA
  }
  out <- out[, dest_cols]

  attr(out, "coord_mode") <- coord_mode
  attr(out, "sheet_name") <- sheet_used
  attr(out, "plot_meta") <- plot_meta

  out
}

.pick_first_existing <- function(df, candidates) {
  nm <- names(df)
  hit <- candidates[candidates %in% nm]
  if (length(hit)) hit[1] else NA_character_
}

.get_col <- function(df, candidates, default = NA) {
  cl <- .pick_first_existing(df, candidates)
  if (is.na(cl)) cl <- .pick_colname(df, candidates)
  if (!is.na(cl) && cl %in% names(df)) return(df[[cl]])
  rep(default, nrow(df))
}

.parse_num_safe <- function(x) {
  z <- as.character(x)
  z[is.na(z)] <- ""
  z <- trimws(z)
  z <- gsub(",", ".", z, fixed = TRUE)
  suppressWarnings(as.numeric(z))
}

.meta_value <- function(df, candidates) {
  cl <- .pick_first_existing(df, candidates)
  if (is.na(cl)) cl <- .pick_colname(df, candidates)
  if (is.na(cl) || !(cl %in% names(df))) return("")

  x <- as.character(df[[cl]])
  x[is.na(x)] <- ""
  x <- trimws(x)
  x <- x[nzchar(x)]
  if (!length(x)) return("")
  unique(x)[1]
}


#' Consolidate multistemmed trees from New Stem Grouping
#'
#' Identifies trees with multiple stems (same New Stem Grouping),
#' calculates equivalent diameter using the area-based method (sqrt(sum of squares)),
#' and keeps only the main stem row (matching New Tag No).
#'
#' @param df Data frame with columns: New Tag No, New Stem Grouping, D, and optionally ExtraD
#' @param min_diameter Numeric. Minimum diameter to include (default = 5 cm)
#'
#' @return Data frame with multistemmed trees consolidated to single rows
#'
#' @keywords internal
#' @noRd
#'
.consolidate_multistem_trees <- function(df, min_diameter = 5) {

  # Check if required columns exist
  if (!all(c("New Tag No", "New Stem Grouping", "D") %in% names(df))) {
    message("Required columns for multistem consolidation not found. Skipping.")
    return(df)
  }

  # Convert columns to appropriate types
  df <- df %>%
    dplyr::mutate(
      `New Stem Grouping` = as.character(`New Stem Grouping`),
      `New Tag No` = as.character(`New Tag No`),
      D = suppressWarnings(as.numeric(D)),
      ExtraD = suppressWarnings(as.numeric(ExtraD))
    )

  # Find groups with multiple stems (duplicate New Stem Grouping)
  stem_group_counts <- df %>%
    dplyr::filter(!is.na(`New Stem Grouping`) & nzchar(`New Stem Grouping`)) %>%
    dplyr::group_by(`New Stem Grouping`) %>%
    dplyr::summarise(n_stems = dplyr::n(), .groups = "drop") %>%
    dplyr::filter(n_stems > 1)

  if (nrow(stem_group_counts) == 0) {
    message("No multistemmed trees found (no duplicate New Stem Grouping values)")
    return(df)
  }

  message("Found ", nrow(stem_group_counts), " multistemmed tree groups")

  # Process each multistem group
  groups_to_keep <- character()
  updated_rows <- list()

  for (group_id in stem_group_counts$`New Stem Grouping`) {

    # Get all rows for this stem group
    group_rows <- df %>% dplyr::filter(`New Stem Grouping` == group_id)

    # Identify ExtraD# Identify the main stem (should match New Tag No within the group)
    # Usually the main stem has the same number as New Tag No
    main_stem_tag <- unique(group_rows$`New Tag No`[!is.na(group_rows$`New Tag No`)])
    main_stem_tag <- main_stem_tag[!is.na(main_stem_tag) & nzchar(main_stem_tag)]

    if (length(main_stem_tag) == 0) {
      warning("No valid New Tag No found for stem group: ", group_id)
      next
    }

    # Use the first tag as main if multiple (shouldn't happen)
    main_tag <- main_stem_tag[1]

    # Collect all diameters from all stems in the group
    all_diameters <- group_rows$D[!is.na(group_rows$D)]

    # Filter valid diameters (>= minimum diameter)
    valid_diameters <- all_diameters[!is.na(all_diameters) & all_diameters >= min_diameter]

    if (length(valid_diameters) <= 1) {
      message("  Group ", group_id, ": Only ", length(valid_diameters),
              " valid stems, no consolidation needed")
      groups_to_keep <- c(groups_to_keep, group_rows$`New Tag No`)
      next
    }

    # Calculate equivalent diameter (sqrt sum of squares)
    D_eq <- sqrt(sum(valid_diameters^2))

    message("  Group ", group_id, ": ", length(valid_diameters),
            " stems -> Equivalent diameter: ", round(D_eq, 1), " mm")

    # Find the main stem row (where New Tag No matches the group's main tag)
    main_row_idx <- which(group_rows$`New Tag No` == main_tag)

    if (length(main_row_idx) == 0) {
      # If main tag not found, use the first row
      main_row_idx <- 1
      warning("Main tag not found in group ", group_id, ", using first row")
    }

    # Get the main row data
    main_row <- group_rows[main_row_idx[1], , drop = FALSE]

    # Update the D value with equivalent diameter
    main_row$D <- D_eq

    # Also check for ExtraD column if available
    if ("ExtraD" %in% names(group_rows)) {
      all_extra_diameters <- group_rows$ExtraD[!is.na(group_rows$ExtraD)]
      valid_extra_diameters <- all_extra_diameters[!is.na(all_extra_diameters) & all_extra_diameters >= min_diameter]
      D_eq <- sqrt(sum(valid_extra_diameters^2))
      main_row$ExtraD <- D_eq
    }

    # Store the updated row and mark which tags to keep
    updated_rows[[length(updated_rows) + 1]] <- main_row
    groups_to_keep <- c(groups_to_keep, main_row$`New Tag No`)
  }

  # Combine updated rows
  updated_df <- dplyr::bind_rows(updated_rows)

  # Remove all rows that belong to processed groups
  rows_to_remove <- df %>%
    dplyr::filter(`New Stem Grouping` %in% stem_group_counts$`New Stem Grouping`)

  # Keep rows that are not in multistem groups
  result <- df %>%
    dplyr::filter(!(`New Stem Grouping` %in% stem_group_counts$`New Stem Grouping`))

  # Add back the consolidated rows
  result <- dplyr::bind_rows(result, updated_df)

  message("Consolidated ", nrow(rows_to_remove), " stem rows into ",
          nrow(updated_df), " tree rows")

  return(result)
}


#' Convert a MONITORA spreadsheet into field-sheet schema
#'
#' @param path Character path to input file.
#' @param sheet Optional preferred worksheet name or index.
#' @param station_name Optional station filter.
#'
#' @return Tibble in the intermediate format expected by plot_html_map(),
#' with metadata row, header row and data rows.
#'
#' @keywords internal
#' @noRd
.monitora_to_field_sheet_df <- function(path,
                                        sheet = 1,
                                        station_name = NULL) {
  if (!file.exists(path)) {
    stop("The provided MONITORA file does not exist.", call. = FALSE)
  }

  if (!requireNamespace("readxl", quietly = TRUE)) {
    stop("Package 'readxl' is required to read MONITORA files.", call. = FALSE)
  }

  .norm <- function(x) {
    x <- as.character(x)
    x[is.na(x)] <- ""
    x2 <- suppressWarnings(iconv(x, from = "", to = "ASCII//TRANSLIT", sub = ""))
    x2[is.na(x2) | !nzchar(x2)] <- x[is.na(x2) | !nzchar(x2)]
    tolower(gsub("[^a-z0-9]+", "", x2))
  }

  .find_best_col <- function(df, aliases) {
    if (ncol(df) == 0) return(NA_character_)

    cn_raw <- names(df)
    cn_norm <- .norm(cn_raw)
    ali_norm <- .norm(aliases)

    hit_idx <- match(ali_norm, cn_norm, nomatch = 0L)
    if (any(hit_idx > 0L)) {
      return(cn_raw[hit_idx[which(hit_idx > 0L)[1]]])
    }

    for (a in ali_norm) {
      if (!nzchar(a) || nchar(a) < 2L) next
      hit <- which(grepl(a, cn_norm, fixed = TRUE))
      if (length(hit) == 1L) {
        return(cn_raw[hit[1]])
      }
    }

    NA_character_
  }

  .find_best_numeric_col <- function(df, aliases) {
    cand <- unique(stats::na.omit(vapply(
      aliases,
      function(a) .find_best_col(df, a),
      FUN.VALUE = character(1)
    )))

    if (!length(cand)) return(NA_character_)

    best <- cand[1]
    best_n <- -1L
    for (cl in cand) {
      v <- .to_numeric(df[[cl]])
      n_ok <- sum(is.finite(v), na.rm = TRUE)
      if (n_ok > best_n) {
        best_n <- n_ok
        best <- cl
      }
    }
    best
  }

  clean_spaces <- function(x) {
    x <- as.character(x)
    x[is.na(x)] <- ""
    x <- gsub(intToUtf8(0x00A0), " ", x, fixed = TRUE)
    x <- gsub(intToUtf8(0x2007), " ", x, fixed = TRUE)
    x <- gsub(intToUtf8(0x202F), " ", x, fixed = TRUE)
    x <- gsub("\\s+", " ", x)
    trimws(x)
  }

  .to_numeric <- function(x) {
    if (is.factor(x)) x <- as.character(x)
    x <- as.character(x)
    x[is.na(x)] <- ""
    x <- clean_spaces(x)
    x <- sub("^([^;/|]+)[;/|].*$", "\\1", x, perl = TRUE)
    x <- gsub("\\.(?=\\d{3}(\\D|$))", "", x, perl = TRUE)
    x <- gsub(",", ".", x, fixed = TRUE)

    pat <- "[-+]?(?:\\d+\\.?\\d*|\\d*\\.?\\d+)"
    has_num <- grepl(pat, x, perl = TRUE)

    y <- rep(NA_character_, length(x))
    y[has_num] <- regmatches(x[has_num], regexpr(pat, x[has_num], perl = TRUE))
    suppressWarnings(as.numeric(y))
  }

  .to_year <- function(x) {
    x <- as.character(x)
    x[is.na(x)] <- ""
    x <- gsub("\\s+", "", x)
    x <- gsub(",", "", x, fixed = TRUE)
    m <- regexpr("(?:19|20)\\d{2}", x, perl = TRUE)
    y <- rep(NA_integer_, length(x))
    hit <- m > 0
    y[hit] <- suppressWarnings(as.integer(substr(
      x[hit],
      m[hit],
      m[hit] + attr(m, "match.length")[hit] - 1
    )))
    y
  }

  .station_norm_name <- function(v) {
    v <- clean_spaces(v)
    v2 <- suppressWarnings(iconv(v, from = "", to = "ASCII//TRANSLIT", sub = ""))
    v2[is.na(v2) | !nzchar(v2)] <- v[is.na(v2) | !nzchar(v2)]
    tolower(v2)
  }

  .station_norm_num <- function(v) {
    v <- clean_spaces(v)
    v <- gsub("\\s+", "", v)
    v <- sub("^0+([0-9])", "\\1", v)
    v[v == ""] <- NA_character_
    v
  }

  first_nonempty <- function(df, col, lixo = character(0)) {
    if (is.na(col) || !(col %in% names(df)) || nrow(df) == 0) return("")
    v <- clean_spaces(df[[col]])
    ok <- !is.na(v) & nzchar(v) & !(tolower(v) %in% tolower(lixo))
    v <- v[ok]
    if (length(v) > 0) v[1] else ""
  }

  map_sub <- function(v) {
    v <- toupper(clean_spaces(v))
    dplyr::case_when(
      v %in% c("N", "NORTE") ~ 1,
      v %in% c("S", "SUL", "SOUTH") ~ 2,
      v %in% c("L", "LESTE", "E", "EAST") ~ 3,
      v %in% c("O", "OESTE", "W", "WEST") ~ 4,
      TRUE ~ NA_real_
    )
  }

  df <- suppressMessages(
    readxl::read_excel(
      path,
      sheet = sheet,
      .name_repair = "minimal",
      col_types = "text"
    )
  )

  if (!is.data.frame(df) || !nrow(df)) {
    stop("Empty MONITORA worksheet.", call. = FALSE)
  }

  names(df) <- clean_spaces(names(df))
  nm <- names(df)

  col_subunidade <- if ("subunidade" %in% nm) "subunidade" else
    if ("Subunidades" %in% nm) "Subunidades" else
      .find_best_col(df, c(
        "subunidade", "subunidades", "sub_unidade", "sub_unidades",
        "orientacao", "orientação", "unidade", "ul"
      ))

  col_nparcela <- if ("N_parcela" %in% nm) "N_parcela" else
    .find_best_col(df, c(
      "n_parcela", "nparcela", "numero_parcela", "número_parcela",
      "parcela", "subplot", "subparcela", "t2"
    ))

  col_tag <- if ("N_arvore" %in% nm) "N_arvore" else
    .find_best_col(df, c(
      "n_arvore", "n árvore", "numero_arvore", "número_arvore",
      "narvore", "arvore", "árvore", "tag", "newtag", "numarvore"
    ))

  col_coletores <- if ("nome_coletores" %in% nm) "nome_coletores" else
    .find_best_col(df, c("nome_coletores", "coletores", "equipe", "team"))

  col_uc <- if ("NOMEUC" %in% nm) "NOMEUC" else
    if ("CDUC" %in% nm) "CDUC" else
      .find_best_col(df, c("nomeuc", "nome_uc", "uc", "unidadeconservacao", "unidade_conservacao", "cduc"))

  col_estacao <- if ("Nome_estacao" %in% nm) "Nome_estacao" else
    .find_best_col(df, c("nome_estacao", "nome estação", "nome_estação", "estacao", "estação", "nomeestacao"))

  col_estacao_n <- if ("N_estacao" %in% nm) "N_estacao" else
    .find_best_col(df, c("n_estacao", "nestacao", "num_estacao", "numero_estacao", "número_estacao", "nºestacao", "n°estacao"))

  col_familia <- if ("Família" %in% nm) "Família" else
    if ("Familia" %in% nm) "Familia" else
      .find_best_col(df, c("familia", "família", "family"))

  col_genero <- if ("Gênero" %in% nm) "Gênero" else
    if ("Genero" %in% nm) "Genero" else
      .find_best_col(df, c("genero", "gênero", "genus"))

  col_especie <- if ("Espécie" %in% nm) "Espécie" else
    if ("Especie" %in% nm) "Especie" else
      .find_best_col(df, c("especie", "espécie", "species", "sp"))

  col_nomecomum <- if ("Nome comum" %in% nm) "Nome comum" else
    .find_best_col(df, c("nome comum", "nome_comum", "nomecomum", "popular", "morphospecies"))

  col_coletado <- if ("individuo coletado" %in% nm) "individuo coletado" else
    if ("individuo coletado ?" %in% nm) "individuo coletado ?" else
      if ("indivíduo coletado" %in% nm) "indivíduo coletado" else
        if ("indivíduo coletado ?" %in% nm) "indivíduo coletado ?" else
          .find_best_col(df, c(
            "individuo coletado", "individuo coletado ?",
            "indivíduo coletado", "indivíduo coletado ?",
            "individuo_coletado", "indivíduo_coletado",
            "coletado", "collected"
          ))

  # prioridade total para a coluna específica de voucher/coletor
  col_voucher_c <- if ("voucher/coletor" %in% .norm(nm)) {
    nm[match("vouchercoletor", .norm(nm))]
  } else {
    .find_best_col(df, c(
      "voucher/coletor", "voucher / coletor", "voucher_coletor",
      "vouchercoletor", "voucher-coletor", "collector voucher", "coletor voucher",
      "coletor do voucher", "voucher collector"
    ))
  }

  col_voucher_n <- if ("voucher/número" %in% nm) "voucher/número" else
    if ("voucher/numero" %in% nm) "voucher/numero" else
      .find_best_col(df, c(
        "voucher/número", "voucher/numero", "voucher_numero", "voucher_número",
        "vouchernumero", "voucher_n", "numero voucher", "número voucher",
        "voucher number"
      ))

  col_x <- if ("X" %in% nm) "X" else
    if ("X(m)" %in% nm) "X(m)" else
      if ("X (m)" %in% nm) "X (m)" else
        .find_best_numeric_col(df, list(
          c("x"), c("x(m)"), c("x (m)"), c("x_m"), c("xm"),
          c("coordx"), c("coord_x"), c("coordenadax"), c("coordenada_x"),
          c("xlocal"), c("x_local")
        ))

  col_y <- if ("Y" %in% nm) "Y" else
    if ("Y(m)" %in% nm) "Y(m)" else
      if ("Y (m)" %in% nm) "Y (m)" else
        .find_best_numeric_col(df, list(
          c("y"), c("y(m)"), c("y (m)"), c("y_m"), c("ym"),
          c("coordy"), c("coord_y"), c("coordenaday"), c("coordenada_y"),
          c("ylocal"), c("y_local")
        ))

  col_cap <- if ("cap_tot" %in% nm) "cap_tot" else
    if ("circ total" %in% nm) "circ total" else
      .find_best_col(df, c("cap_tot", "captot", "cap total", "cap_total", "cap", "circ total", "circunferencia", "circunferência"))

  col_ano <- if ("Ano" %in% nm) "Ano" else
    if ("Data" %in% nm) "Data" else
      .find_best_col(df, c("ano", "censo", "year", "data", "date"))

  col_dead <- if ("arvore_morta" %in% nm) "arvore_morta" else
    if ("árvore_morta" %in% nm) "árvore_morta" else
      .find_best_col(df, c("arvore_morta", "árvore_morta", "arvore morta", "morta", "dead"))

  col_obs <- if ("observação" %in% nm) "observação" else
    if ("observacao" %in% nm) "observacao" else
      .find_best_col(df, c("observação", "observacao", "obs", "census notes", "comentarios", "comentários"))

  col_basal_area <- if ("AB" %in% nm) "AB" else
    if ("ABcap" %in% nm) "ABcap" else
      .find_best_col(df, c("ab", "abcap", "basal area", "area basal", "área basal"))

  col_pom <- if ("POM(m)" %in% nm) "POM(m)" else
    if ("POM" %in% nm) "POM" else
      .find_best_col(df, c("pom(m)", "pom", "pom_m"))

  col_altura <- if ("altura" %in% nm) "altura" else
    .find_best_col(df, c("altura", "height"))

  if (is.na(col_tag) || is.na(col_subunidade) || is.na(col_nparcela)) {
    stop(
      paste0(
        "MONITORA schema could not be detected correctly. ",
        "Detected columns: ", paste(names(df), collapse = " | ")
      ),
      call. = FALSE
    )
  }

  if (is.na(col_x) || is.na(col_y)) {
    stop(
      paste0(
        "MONITORA coordinate columns could not be detected. ",
        "col_x=", col_x, "; col_y=", col_y,
        ". Available columns: ", paste(names(df), collapse = " | ")
      ),
      call. = FALSE
    )
  }

  ano_vec <- if (!is.na(col_ano) && col_ano %in% names(df)) .to_year(df[[col_ano]]) else rep(NA_integer_, nrow(df))
  ano_max <- suppressWarnings(max(ano_vec, na.rm = TRUE))
  df_use <- if (is.finite(ano_max)) df[ano_vec == ano_max, , drop = FALSE] else df
  df_all <- df

  if (!is.na(col_estacao) && col_estacao %in% names(df_use)) df_use$.__st_name <- .station_norm_name(df_use[[col_estacao]]) else df_use$.__st_name <- NA_character_
  if (!is.na(col_estacao_n) && col_estacao_n %in% names(df_use)) df_use$.__st_num <- .station_norm_num(df_use[[col_estacao_n]]) else df_use$.__st_num <- NA_character_
  if (!is.na(col_estacao) && col_estacao %in% names(df_all)) df_all$.__st_name <- .station_norm_name(df_all[[col_estacao]]) else df_all$.__st_name <- NA_character_
  if (!is.na(col_estacao_n) && col_estacao_n %in% names(df_all)) df_all$.__st_num <- .station_norm_num(df_all[[col_estacao_n]]) else df_all$.__st_num <- NA_character_

  if (!is.null(station_name) && length(station_name)) {
    want_raw <- trimws(as.character(station_name))
    want_name <- .station_norm_name(want_raw)
    want_num  <- .station_norm_num(want_raw)

    keep_recent <- (!is.na(df_use$.__st_name) & df_use$.__st_name %in% want_name) |
      (!is.na(df_use$.__st_num) & df_use$.__st_num %in% want_num)

    if (!any(keep_recent)) {
      stop("Requested `station_name` not found in the most recent census.", call. = FALSE)
    }

    df_use <- df_use[keep_recent, , drop = FALSE]

    keep_all <- (!is.na(df_all$.__st_name) & df_all$.__st_name %in% want_name) |
      (!is.na(df_all$.__st_num) & df_all$.__st_num %in% want_num)

    if (any(keep_all)) {
      df_all <- df_all[keep_all, , drop = FALSE]
    }
  }

  x_recent <- .to_numeric(df_use[[col_x]])
  y_recent <- .to_numeric(df_use[[col_y]])

  if (!is.na(col_tag) && !is.na(col_ano) &&
      all(c(col_x, col_y, col_tag, col_ano) %in% names(df_all))) {

    need_xy <- !is.finite(x_recent) | !is.finite(y_recent)

    if (any(need_xy)) {
      year_all <- .to_year(df_all[[col_ano]])
      older <- df_all[
        year_all != suppressWarnings(max(year_all, na.rm = TRUE)) & !is.na(year_all),
        ,
        drop = FALSE
      ]

      if (nrow(older)) {
        old_tag <- clean_spaces(older[[col_tag]])
        old_x <- .to_numeric(older[[col_x]])
        old_y <- .to_numeric(older[[col_y]])
        old_year <- .to_year(older[[col_ano]])

        ord <- order(old_year, decreasing = TRUE, na.last = TRUE)
        old_tag <- old_tag[ord]
        old_x <- old_x[ord]
        old_y <- old_y[ord]

        tag_use <- clean_spaces(df_use[[col_tag]])

        for (i in which(need_xy)) {
          tg <- tag_use[i]
          if (!nzchar(tg)) next
          hit <- which(old_tag == tg & (is.finite(old_x) | is.finite(old_y)))
          if (!length(hit)) next
          j <- hit[1]
          if (!is.finite(x_recent[i]) && is.finite(old_x[j])) x_recent[i] <- old_x[j]
          if (!is.finite(y_recent[i]) && is.finite(old_y[j])) y_recent[i] <- old_y[j]
        }
      }
    }
  }

  voucher_collector_raw <- if (!is.na(col_voucher_c) && col_voucher_c %in% names(df_use)) {
    clean_spaces(df_use[[col_voucher_c]])
  } else {
    rep("", nrow(df_use))
  }

  voucher_number_raw <- if (!is.na(col_voucher_n) && col_voucher_n %in% names(df_use)) {
    clean_spaces(df_use[[col_voucher_n]])
  } else {
    rep("", nrow(df_use))
  }

  voucher_vec <- ifelse(
    (!nzchar(voucher_collector_raw)) & (!nzchar(voucher_number_raw)),
    NA_character_,
    trimws(paste(voucher_collector_raw, voucher_number_raw))
  )

  voucher_vec[!nzchar(voucher_vec)] <- NA_character_
  voucher_collector_raw[!nzchar(voucher_collector_raw)] <- NA_character_

  gen <- if (!is.na(col_genero) && col_genero %in% names(df_use)) {
    clean_spaces(df_use[[col_genero]])
  } else {
    rep(NA_character_, nrow(df_use))
  }

  esp <- if (!is.na(col_especie) && col_especie %in% names(df_use)) {
    clean_spaces(df_use[[col_especie]])
  } else {
    rep(NA_character_, nrow(df_use))
  }

  family_vec <- if (!is.na(col_familia) && col_familia %in% names(df_use)) {
    clean_spaces(df_use[[col_familia]])
  } else {
    rep(NA_character_, nrow(df_use))
  }

  family_norm <- suppressWarnings(iconv(family_vec, from = "", to = "ASCII//TRANSLIT", sub = ""))
  family_norm[is.na(family_norm) | !nzchar(family_norm)] <- family_vec[is.na(family_norm) | !nzchar(family_norm)]
  family_norm <- tolower(trimws(family_norm))

  family_bad <- is.na(family_vec) |
    !nzchar(family_vec) |
    family_norm %in% c(
      "indet", "indet.", "indeterminada", "indeterminado",
      "indeterminada.", "indeterminado.", "indeterm.", "na", "n/a"
    )

  family_vec[family_bad] <- "Indet"

  gen_norm <- suppressWarnings(iconv(gen, from = "", to = "ASCII//TRANSLIT", sub = ""))
  gen_norm[is.na(gen_norm) | !nzchar(gen_norm)] <- gen[is.na(gen_norm) | !nzchar(gen_norm)]
  gen_norm <- tolower(trimws(gen_norm))

  gen_bad <- is.na(gen) |
    !nzchar(gen) |
    gen_norm %in% c(
      "indet", "indet.", "indeterminada", "indeterminado",
      "indeterminada.", "indeterminado.", "indeterm.", "na", "n/a"
    )

  gen[gen_bad] <- NA_character_

  esp_norm <- suppressWarnings(iconv(esp, from = "", to = "ASCII//TRANSLIT", sub = ""))
  esp_norm[is.na(esp_norm) | !nzchar(esp_norm)] <- esp[is.na(esp_norm) | !nzchar(esp_norm)]
  esp_norm <- tolower(trimws(esp_norm))

  esp_bad <- is.na(esp) |
    !nzchar(esp) |
    esp_norm %in% c(
      "indet", "indet.", "indeterminada", "indeterminado",
      "indeterminada.", "indeterminado.", "indeterm.", "na", "n/a"
    ) |
    grepl("^sp\\.?\\s*\\d*$", esp_norm, perl = TRUE)

  esp[esp_bad] <- "indet"

  od <- ifelse(
    is.na(gen) | !nzchar(gen),
    "Indet indet",
    paste(gen, esp)
  )

  od[is.na(od) | !nzchar(od)] <- "Indet indet"

  morpho_vec <- if (!is.na(col_nomecomum) && col_nomecomum %in% names(df_use)) clean_spaces(df_use[[col_nomecomum]]) else rep(NA_character_, nrow(df_use))
  morpho_vec[!nzchar(morpho_vec)] <- NA_character_

  collected_raw <- if (!is.na(col_coletado) && col_coletado %in% names(df_use)) {
    clean_spaces(df_use[[col_coletado]])
  } else {
    rep("", nrow(df_use))
  }

  collected_norm <- tolower(collected_raw)

  is_collected <- collected_norm %in% c(
    "sim", "s", "yes", "y", "true", "1", "coletado", "coletada"
  ) | nzchar(as.character(voucher_vec))

  collected_vec <- ifelse(is_collected, "Sim", NA_character_)

  notes_vec <- if (!is.na(col_obs) && col_obs %in% names(df_use)) clean_spaces(df_use[[col_obs]]) else rep(NA_character_, nrow(df_use))
  notes_vec[!nzchar(notes_vec)] <- NA_character_

  cap_vec <- if (!is.na(col_cap) && col_cap %in% names(df_use)) .to_numeric(df_use[[col_cap]]) else rep(NA_real_, nrow(df_use))
  d_mm <- (cap_vec / pi) * 10

  pom_vec <- if (!is.na(col_pom) && col_pom %in% names(df_use)) .to_numeric(df_use[[col_pom]]) else rep(NA_real_, nrow(df_use))
  ba_vec <- if (!is.na(col_basal_area) && col_basal_area %in% names(df_use)) .to_numeric(df_use[[col_basal_area]]) else rep(NA_real_, nrow(df_use))
  height_vec <- if (!is.na(col_altura) && col_altura %in% names(df_use)) .to_numeric(df_use[[col_altura]]) else rep(NA_real_, nrow(df_use))

  out <- data.frame(
    `New Tag No` = as.character(df_use[[col_tag]]),
    `New Stem Grouping` = NA_character_,
    T1 = map_sub(df_use[[col_subunidade]]),
    T2 = .to_numeric(df_use[[col_nparcela]]),
    X = x_recent,
    Y = y_recent,
    Family = family_vec,
    `Original determination` = od,
    Morphospecies = morpho_vec,
    D = ifelse(is.finite(d_mm), round(d_mm, 2), NA_real_),
    POM = pom_vec,
    ExtraD = NA_character_,
    ExtraPOM = NA_character_,
    Flag1 = NA_character_,
    Flag2 = NA_character_,
    Flag3 = NA_character_,
    LI = NA_character_,
    CI = NA_character_,
    CF = NA_character_,
    CD1 = NA_character_,
    nrdups = NA_character_,
    Height = height_vec,
    Voucher = voucher_vec,
    `Voucher Collector` = voucher_collector_raw,
    Silica = NA_character_,
    Collected = collected_vec,
    `Census Notes` = notes_vec,
    CAP = cap_vec,
    `Basal Area` = ba_vec,
    check.names = FALSE,
    stringsAsFactors = FALSE
  )

  if (all(is.na(out$X)) || all(is.na(out$Y))) {
    stop(
      paste0(
        "MONITORA conversion produced X/Y entirely as NA. ",
        "Detected columns: X='", col_x, "', Y='", col_y, "'."
      ),
      call. = FALSE
    )
  }

  lixo_uc <- c("", "na", "n/a", "-", "--", ".", "s/n", "sn")
  lixo_estacao <- c("", "na", "n/a", "-", "--", ".")

  plot_name <- first_nonempty(df_use, col_uc, lixo_uc)
  plot_code <- first_nonempty(df_use, col_estacao, lixo_estacao)

  if (!nzchar(plot_code) && !is.na(col_estacao_n) && col_estacao_n %in% names(df_use)) {
    plot_code <- first_nonempty(df_use, col_estacao_n, lixo_estacao)
  }

  team_val <- ""
  if (!is.na(col_coletores) && col_coletores %in% names(df_use)) {
    team_vals <- unique(clean_spaces(df_use[[col_coletores]]))
    team_vals <- team_vals[nzchar(team_vals)]
    if (length(team_vals)) {
      team_val <- paste(team_vals, collapse = "; ")
    }
  }

  census_years <- if (!is.na(col_ano) && col_ano %in% names(df_all)) sort(unique(.to_year(df_all[[col_ano]]))) else integer(0)
  dead_since_first <- NA_integer_

  if (!is.na(col_dead) && col_dead %in% names(df_all)) {
    v <- tolower(clean_spaces(df_all[[col_dead]]))
    dead_since_first <- sum(v %in% c("sim", "s", "yes", "y", "true", "1"), na.rm = TRUE)
  }

  recruits_since_first <- NA_integer_
  if (!is.na(col_tag) && col_tag %in% names(df_all) && length(census_years) >= 2 && !is.na(col_ano)) {
    year_all <- .to_year(df_all[[col_ano]])
    y0 <- min(census_years, na.rm = TRUE)
    y1 <- max(census_years, na.rm = TRUE)
    tag_first <- unique(clean_spaces(df_all[[col_tag]][year_all == y0]))
    tag_last  <- unique(clean_spaces(df_all[[col_tag]][year_all == y1]))
    tag_first <- tag_first[!is.na(tag_first) & nzchar(tag_first)]
    tag_last  <- tag_last[!is.na(tag_last) & nzchar(tag_last)]
    recruits_since_first <- length(setdiff(tag_last, tag_first))
  }

  try(assign("plot_name", plot_name, envir = parent.frame()), silent = TRUE)
  try(assign("plot_code", plot_code, envir = parent.frame()), silent = TRUE)
  try(assign("team", team_val, envir = parent.frame()), silent = TRUE)
  try(assign("monitora_census_years", census_years, envir = parent.frame()), silent = TRUE)
  try(assign("monitora_census_n", length(census_years), envir = parent.frame()), silent = TRUE)
  try(assign("monitora_dead_since_first", dead_since_first, envir = parent.frame()), silent = TRUE)
  try(assign("monitora_recruits_since_first", recruits_since_first, envir = parent.frame()), silent = TRUE)

  attr(out, "coord_mode") <- "local"
  attr(out, "sheet_name") <- sheet
  attr(out, "plot_meta") <- list(
    plot_name = plot_name,
    plot_code = plot_code,
    team = team_val,
    census_years = census_years
  )

  out
}


#' Compute global coordinates for serpentine ForestPlots layouts
#'
#' @param fp_clean Canonical field-sheet data frame.
#' @param plot_size Plot size in hectares.
#' @param subplot_size Subplot side in meters.
#'
#' @return Data frame with draw/global coordinates.
#'
#' @keywords internal
#' @noRd
.compute_global_coordinates <- function(fp_clean,
                                        subplot_size,
                                        plot_width_m,
                                        plot_length_m) {
  n_rows <- floor(plot_length_m / subplot_size)
  n_cols <- floor(plot_width_m / subplot_size)

  if (n_rows <= 0 || n_cols <= 0) {
    stop("Invalid plot dimensions relative to `subplot_size`.", call. = FALSE)
  }

  max_x <- n_cols * subplot_size
  max_y <- n_rows * subplot_size

  fp_clean %>%
    dplyr::mutate(
      T1 = suppressWarnings(as.numeric(T1)),
      X = suppressWarnings(as.numeric(X)),
      Y = suppressWarnings(as.numeric(Y)),
      col = floor((T1 - 1) / n_rows),
      row = (T1 - 1) %% n_rows,
      global_x = col * subplot_size + X,
      global_y = dplyr::if_else(
        col %% 2 == 0,
        row * subplot_size + Y,
        (n_rows - row - 1) * subplot_size + Y
      )
    ) %>%
    dplyr::filter(
      global_x >= 0,
      global_y >= 0,
      global_x <= max_x,
      global_y <= max_y
    )
}


#' Compute MONITORA geometry for both full-plot and individual-subplot views
#'
#' @param fp_df Canonical field-sheet-like data frame containing at least
#'   `T1`, `T2`, `X`, and `Y`.
#' @param keep_only_cell Logical; if TRUE, keep only points falling inside the
#'   expected 10 x 10 m subplot cell. Use FALSE for the full MONITORA layout
#'   and TRUE for individual subplot pages.
#'
#' @return A data frame with MONITORA geometry columns added:
#'   `subunit_letter`, `arm`, `X_loc`, `Y_loc`, `along_m`,
#'   `draw_x`, `draw_y`, `subplot_name`, `x10`, and `y10`.
#'
#' @details
#' This helper centralizes all MONITORA geometry logic in one place.
#' It replaces the previous split workflow where one function computed
#' global drawing coordinates and another converted points to 10 x 10 m
#' subplot cell coordinates.
#'
#' The main bug fixed here is that the old implementation discarded points
#' too early based on the sign of `X_loc` (left half vs right half). That
#' could artificially empty valid subplots when the input MONITORA file did
#' not follow the exact assumed sign convention. Here, `x10` is derived
#' directly from the expected subplot row, and filtering happens only at the
#' end, after the local subplot coordinates have been computed.
#'
#' @keywords internal
#' @noRd
#'
.compute_monitora_geometry <- function(fp_df, keep_only_cell = TRUE) {

  if (is.null(fp_df) || !nrow(fp_df)) {
    return(fp_df)
  }

  out <- fp_df

  if (!all(c("T1", "T2", "X", "Y") %in% names(out))) {
    stop(
      "MONITORA geometry requires columns `T1`, `T2`, `X`, and `Y`.",
      call. = FALSE
    )
  }

  t1_raw <- toupper(trimws(as.character(out$T1)))
  t1_raw[t1_raw %in% "1"] <- "N"
  t1_raw[t1_raw %in% "2"] <- "S"
  t1_raw[t1_raw %in% "3"] <- "L"
  t1_raw[t1_raw %in% "4"] <- "O"
  out$T1 <- t1_raw

  out$T2 <- suppressWarnings(as.integer(out$T2))
  out$X  <- suppressWarnings(as.numeric(out$X))
  out$Y  <- suppressWarnings(as.numeric(out$Y))

  out <- out %>%
    dplyr::mutate(
      subunit_letter = T1,
      arm = T1,
      X_loc = pmax(pmin(X, 10), -10),
      Y_loc = pmax(pmin(Y, 50), 0)
    )

  out <- out %>%
    dplyr::mutate(
      along_m = Y_loc
    )

  out <- out %>%
    dplyr::mutate(
      draw_x = dplyr::case_when(
        arm %in% c("N", "S") ~ X_loc,
        arm == "L" ~ 50 + along_m,
        arm == "O" ~ -100 + along_m,
        TRUE ~ NA_real_
      ),
      draw_y = dplyr::case_when(
        arm == "N" ~ 50 + along_m,
        arm == "S" ~ -100 + along_m,
        arm %in% c("L", "O") ~ X_loc,
        TRUE ~ NA_real_
      ),
      subplot_name = paste0(arm, T2)
    ) %>%
    dplyr::filter(is.finite(draw_x), is.finite(draw_y))

  col_idx <- ceiling(out$T2 / 2)
  base_num <- (col_idx - 1L) * 2L + 1L
  off <- out$T2 - base_num
  is_even_col <- (col_idx %% 2L) == 0L
  row_in_col <- ifelse(is_even_col, 1L - off, off)
  col_from_center <- col_idx - 1L

  out <- out %>%
    dplyr::mutate(
      y10 = along_m - col_from_center * 10,
      x10 = dplyr::if_else(row_in_col == 0L, X_loc + 10, X_loc)
    )

  if (isTRUE(keep_only_cell)) {
    out <- out %>%
      dplyr::filter(
        is.finite(x10), is.finite(y10),
        x10 >= 0, x10 <= 10,
        y10 >= 0, y10 <= 10
      )
  }

  out
}


#' Calculate phytosociological metrics from canonical field-sheet data
#'
#' @param fp_sheet Canonical field-sheet data frame.
#' @param plot_size_ha Plot size in hectares.
#' @param subplot_size_m Subplot side in meters.
#'
#' @return Named list with species, family and diversity metrics.
#'
#' @keywords internal
#' @noRd
#'
.calculate_phytosociological_metrics <- function(fp_sheet,
                                                 plot_size_ha = 1,
                                                 subplot_size_m = 10) {
  required_cols <- c("Family", "Original determination", "T1")
  miss <- setdiff(required_cols, names(fp_sheet))
  if (length(miss)) {
    stop(
      "Missing required canonical column(s) in `fp_sheet`: ",
      paste(miss, collapse = ", "),
      call. = FALSE
    )
  }

  df_clean <- fp_sheet %>%
    dplyr::mutate(
      Family = trimws(as.character(Family)),
      Family = dplyr::if_else(is.na(Family) | Family == "", "Indet", Family),
      Species = trimws(as.character(`Original determination`)),
      Species = dplyr::if_else(is.na(Species) | Species == "", "indet", Species),
      T1 = as.character(T1),
      D = suppressWarnings(as.numeric(D))
    ) %>%
    dplyr::filter(!is.na(Family) & Family != "") %>%
    dplyr::filter(!is.na(Species) & Species != "")

  if (nrow(df_clean) == 0) {
    warning("No valid species data found after filtering", call. = FALSE)
    return(list(
      species_metrics = data.frame(),
      family_metrics = data.frame(),
      diversity_metrics = list(
        total_species = 0,
        total_families = 0,
        total_individuals = 0,
        shannon_index = NA_real_,
        simpson_index = NA_real_
      )
    ))
  }

  n_subplots_total <- dplyr::n_distinct(df_clean$T1)

  species_stats <- df_clean %>%
    dplyr::group_by(species = Species) %>%
    dplyr::summarise(
      family = dplyr::first(stats::na.omit(Family)),
      abundance = dplyr::n(),
      n_subplots = dplyr::n_distinct(T1),
      .groups = "drop"
    ) %>%
    dplyr::mutate(
      relative_density = round(100 * abundance / sum(abundance), 2),
      relative_frequency = round(100 * n_subplots / n_subplots_total, 2),
      importance_value = round(relative_density + relative_frequency, 2)
    ) %>%
    dplyr::arrange(dplyr::desc(abundance), species)

  family_stats <- df_clean %>%
    dplyr::group_by(family = Family) %>%
    dplyr::summarise(
      abundance = dplyr::n(),
      n_species = dplyr::n_distinct(Species),
      n_subplots = dplyr::n_distinct(T1),
      .groups = "drop"
    ) %>%
    dplyr::mutate(
      relative_density = round(100 * abundance / sum(abundance), 2),
      relative_frequency = round(100 * n_subplots / n_subplots_total, 2),
      importance_value = round(relative_density + relative_frequency, 2)
    ) %>%
    dplyr::arrange(dplyr::desc(abundance), family)

  abundances <- species_stats$abundance
  proportions <- abundances / sum(abundances)
  shannon_index <- -sum(proportions * log(proportions))
  simpson_index <- 1 - sum(proportions ^ 2)

  list(
    species_metrics = species_stats,
    family_metrics = family_stats,
    diversity_metrics = list(
      total_species = nrow(species_stats),
      total_families = nrow(family_stats),
      total_individuals = sum(species_stats$abundance),
      shannon_index = round(shannon_index, 3),
      simpson_index = round(simpson_index, 3)
    )
  )
}


#' Prepare taxonomic dashboard tables and plots
#'
#' @param fp_sheet Canonical field-sheet data frame.
#' @param plot_size_ha Plot size in hectares.
#' @param subplot_size_m Subplot size in meters.
#' @param language Language code.
#'
#' @return Named list with summary tables and plots.
#'
#' @keywords internal
#' @noRd
#'
.prepare_report_dashboard <- function(fp_sheet,
                                      plot_size_ha = 1,
                                      subplot_size_m = 10,
                                      language = "en") {
  language <- tolower(trimws(as.character(language)[1]))
  if (!language %in% c("en", "pt", "es", "fr", "ma", "pa")) {
    language <- "en"
  }

  lab <- switch(
    language,
    pt = list(
      metric = "M\u00e9trica",
      value = "Valor",
      total_individuals = "Total de indiv\u00edduos",
      collected = "Coletados",
      uncollected = "N\u00e3o coletados",
      palms = "Palmeiras (Arecaceae)",
      families = "Fam\u00edlias distintas",
      species = "Esp\u00e9cies distintas",
      genera = "G\u00eaneros distintos",
      shannon = "\u00cdndice de Shannon",
      simpson = "\u00cdndice de Simpson",
      family_plot = "Fam\u00edlias mais comuns",
      species_plot = "Esp\u00e9cies mais abundantes",
      subplot_plot = "Percentual de coleta por subparcela",
      dbh_plot = "Classes de DAP (cm)",
      dbh_x = "Classe de DAP (cm)",
      dbh_y = "Numero de indiv\u00edduos",
      x_ind = "Indiv\u00edduos",
      x_pct = "%",
      subplot = "Subparcela",
      species_tbl = c(
        "Esp\u00e9cie", "Fam\u00edlia", "Abund\u00e2ncia", "Subparcelas",
        "Dens. rel. (%)", "Freq. rel. (%)", "VI"
      ),
      family_tbl = c(
        "Fam\u00edlia", "Abund\u00e2ncia", "Riqueza", "Subparcelas",
        "Dens. rel. (%)", "Freq. rel. (%)", "VI"
      )
    ),
    es = list(
      metric = "M\u00e9trica",
      value = "Valor",
      total_individuals = "Total de individuos",
      collected = "Espec\u00edmenes colectados",
      uncollected = "Espec\u00edmenes no colectados",
      palms = "Palmas (Arecaceae)",
      families = "Familias distintas",
      species = "Especies distintas",
      genera = "G\u00e9neros distintos",
      shannon = "\u00cdndice de Shannon",
      simpson = "\u00cdndice de Simpson",
      family_plot = "Familias m\u00e1s comunes",
      species_plot = "Especies m\u00e1s abundantes",
      subplot_plot = "Porcentaje de recolecci\u00f3n por subparcela",
      dbh_plot = "Clases de DAP (cm)",
      dbh_x = "Clase de DAP (cm)",
      dbh_y = "N\u00famero de individuos",
      x_ind = "Individuos",
      x_pct = "%",
      subplot = "Subparcela",
      species_tbl = c(
        "Especie", "Familia", "Abundancia", "Subparcelas",
        "Dens. rel. (%)", "Freq. rel. (%)", "VI"
      ),
      family_tbl = c(
        "Familia", "Abundancia", "Riqueza", "Subparcelas",
        "Dens. rel. (%)", "Freq. rel. (%)", "VI"
      )
    ),
    fr = list(
      metric = "M\u00e9trique",
      value = "Valeur",
      total_individuals = "Nombre total d'individus",
      collected = "Sp\u00e9cimens collect\u00e9s",
      uncollected = "Sp\u00e9cimens non collect\u00e9s",
      palms = "Palmiers (Arecaceae)",
      families = "Familles distinctes",
      species = "Esp\u00e8ces distinctes",
      genera = "Genres distincts",
      shannon = "Indice de Shannon",
      simpson = "Indice de Simpson",
      family_plot = "Familles les plus communes",
      species_plot = "Esp\u00e8ces les plus abondantes",
      subplot_plot = "Pourcentage de collecte par sous-parcelle",
      dbh_plot = "Classes de DHP (cm)",
      dbh_x = "Classe de DHP (cm)",
      dbh_y = "Nombre d'individus",
      x_ind = "Individus",
      x_pct = "%",
      subplot = "Sous-parcelle",
      species_tbl = c(
        "Esp\u00e8ce", "Famille", "Abondance", "Sous-parcelles",
        "Densit\u00e9 rel. (%)", "Fr\u00e9q. rel. (%)", "IV"
      ),
      family_tbl = c(
        "Famille", "Abondance", "Richesse", "Sous-parcelles",
        "Densit\u00e9 rel. (%)", "Fr\u00e9q. rel. (%)", "IV"
      )
    ),
    ma = list(
      metric = "\u6307\u6807",
      value = "\u6570\u503c",
      total_individuals = "\u4e2a\u4f53\u603b\u6570",
      collected = "\u5df2\u91c7\u96c6\u4e2a\u4f53",
      uncollected = "\u672a\u91c7\u96c6\u4e2a\u4f53",
      palms = "\u68d5\u6988\u79d1\u4e2a\u4f53 (Arecaceae)",
      families = "\u4e0d\u540c\u79d1\u6570",
      species = "\u4e0d\u540c\u7269\u79cd\u6570",
      genera = "\u4e0d\u540c\u5c5e\u6570",
      shannon = "Shannon \u6307\u6570",
      simpson = "Simpson \u6307\u6570",
      family_plot = "\u6700\u5e38\u89c1\u79d1",
      species_plot = "\u6700\u4e30\u5bcc\u7269\u79cd",
      subplot_plot = "\u5404\u5b50\u6837\u5730\u91c7\u96c6\u767e\u5206\u6bd4",
      dbh_plot = "\u80f8\u5f84\u7b49\u7ea7 (cm)",
      dbh_x = "\u80f8\u5f84\u7b49\u7ea7 (cm)",
      dbh_y = "\u4e2a\u4f53\u6570\u91cf",
      x_ind = "\u4e2a\u4f53\u6570",
      x_pct = "%",
      subplot = "\u5b50\u6837\u5730",
      species_tbl = c(
        "\u7269\u79cd", "\u79d1", "\u4e30\u5ea6", "\u5b50\u6837\u5730",
        "\u76f8\u5bf9\u5bc6\u5ea6 (%)", "\u76f8\u5bf9\u9891\u7387 (%)", "\u91cd\u8981\u503c"
      ),
      family_tbl = c(
        "\u79d1", "\u4e30\u5ea6", "\u4e30\u5bcc\u5ea6", "\u5b50\u6837\u5730",
        "\u76f8\u5bf9\u5bc6\u5ea6 (%)", "\u76f8\u5bf9\u9891\u7387 (%)", "\u91cd\u8981\u503c"
      )
    ),
    pa = list(
      metric = "Hopjon",
      value = "Kukr\u1ebd",
      total_individuals = "P\u00e3p\u00e3 p\u00e2ri",
      collected = "P\u00e2ri sonswa",
      uncollected = "P\u00e2ri r\u00f5r\u0129",
      palms = "Kwatis\u00f4m\u1ebdra (Arecaceae)",
      families = "Kyapi\u00e2hapi\u00e2ra j\u00e0ri",
      species = "P\u0129rak\u00e2ri j\u00e0ri",
      genera = "G\u00eaneros distintos",
      shannon = "Indice de Shannon",
      simpson = "Indice de Simpson",
      family_plot = "Inkj\u00eati tip\u0129njakjura",
      species_plot = "Sotinkj\u00eati sop\u00e2ri m\u1ebdra",
      subplot_plot = "Percentual de Coleta por Kuk\u00e2ra krep\u00e3\u00e3 s\u00e2\u00e2",
      dbh_plot = "Classes de DAP (cm)",
      dbh_x = "Classe de DAP (cm)",
      dbh_y = "N\u00famero de P\u00e2ri",
      x_ind = "P\u00e2ri",
      x_pct = "%",
      subplot = "Kuk\u00e2ra krep\u00e3\u00e3 s\u00e2\u00e2",
      species_tbl = c(
        "P\u0129rak\u00e2ri", "Kyapi\u00e2hapi\u00e2ra", "Abund\u00e2ncia", "Kuk\u00e2ra krep\u00e3\u00e3 s\u00e2\u00e2",
        "Dens. rel. (%)", "Freq. rel. (%)", "VI"
      ),
      family_tbl = c(
        "Kyapi\u00e2hapi\u00e2ra", "Abund\u00e2ncia", "Riqueza", "Kuk\u00e2ra krep\u00e3\u00e3 s\u00e2\u00e2",
        "Dens. rel. (%)", "Freq. rel. (%)", "VI"
      )
    ),
    list(
      metric = "Metric",
      value = "Value",
      total_individuals = "Total individuals",
      collected = "Collected specimens",
      uncollected = "Uncollected specimens",
      palms = "Palms (Arecaceae)",
      families = "Distinct families",
      species = "Distinct species",
      genera = "Distinct genera",
      shannon = "Shannon index",
      simpson = "Simpson index",
      family_plot = "Most common families",
      species_plot = "Most abundant species",
      subplot_plot = "Collection percentage by subplot",
      dbh_plot = "DBH classes (cm)",
      dbh_x = "DBH class (cm)",
      dbh_y = "Number of individuals",
      x_ind = "Individuals",
      x_pct = "%",
      subplot = "Subplot",
      species_tbl = c(
        "Species", "Family", "Abundance", "Subplots",
        "Rel. density (%)", "Rel. frequency (%)", "IV"
      ),
      family_tbl = c(
        "Family", "Abundance", "Richness", "Subplots",
        "Rel. density (%)", "Rel. frequency (%)", "IV"
      )
    )
  )

  fmt_int <- function(x) {
    if (length(x) == 0 || is.na(x)) return(NA_character_)
    as.character(as.integer(round(x)))
  }

  fmt_dec <- function(x, digits = 3) {
    if (length(x) == 0 || is.na(x)) return(NA_character_)
    sprintf(paste0("%.", digits, "f"), x)
  }

  df <- fp_sheet %>%
    dplyr::mutate(
      Family = trimws(as.character(Family)),
      Family = dplyr::if_else(is.na(Family) | Family == "", "Indet", Family),
      Species = trimws(as.character(`Original determination`)),
      Species = dplyr::if_else(is.na(Species) | Species == "", "indet", Species),
      Collected = trimws(as.character(Collected)),
      T1 = as.character(T1),
      D = suppressWarnings(as.numeric(D))
    )

  total_individuals <- nrow(df)
  collected_count <- sum(!is.na(df$Collected) & df$Collected != "", na.rm = TRUE)
  uncollected_count <- sum(
    (is.na(df$Collected) | df$Collected == "") & df$Family != "Arecaceae",
    na.rm = TRUE
  )
  palms_count <- sum(df$Family == "Arecaceae", na.rm = TRUE)

  distinct_families <- dplyr::n_distinct(df$Family[df$Family != ""])
  distinct_species <- dplyr::n_distinct(df$Species[df$Species != ""])
  distinct_genera <- dplyr::n_distinct(sub(" .*", "", df$Species[df$Species != ""]))

  phytosoc <- .calculate_phytosociological_metrics(
    fp_sheet = df,
    plot_size_ha = plot_size_ha,
    subplot_size_m = subplot_size_m
  )

  metrics_tbl <- tibble::tibble(
    !!lab$metric := c(
      lab$total_individuals,
      lab$collected,
      lab$uncollected,
      lab$palms,
      lab$families,
      lab$species,
      lab$genera,
      lab$shannon,
      lab$simpson
    ),
    !!lab$value := c(
      fmt_int(total_individuals),
      fmt_int(collected_count),
      fmt_int(uncollected_count),
      fmt_int(palms_count),
      fmt_int(distinct_families),
      fmt_int(distinct_species),
      fmt_int(distinct_genera),
      fmt_dec(phytosoc$diversity_metrics$shannon_index, 3),
      fmt_dec(phytosoc$diversity_metrics$simpson_index, 3)
    )
  )

  family_counts <- df %>%
    dplyr::count(Family, name = "count", sort = TRUE) %>%
    dplyr::slice_head(n = 10)

  species_counts <- df %>%
    dplyr::count(Species, name = "count", sort = TRUE) %>%
    dplyr::slice_head(n = 10)

  subplot_stats <- df %>%
    dplyr::group_by(subplot = as.character(T1)) %>%
    dplyr::summarise(
      total = dplyr::n(),
      collected = sum(!is.na(Collected) & Collected != "", na.rm = TRUE),
      .groups = "drop"
    ) %>%
    dplyr::mutate(
      pct = ifelse(total > 0, round(100 * collected / total, 1), 0)
    )

  ord <- suppressWarnings(as.numeric(subplot_stats$subplot))
  subplot_stats$.ord <- ifelse(is.na(ord), Inf, ord)

  subplot_stats <- subplot_stats %>%
    dplyr::arrange(.ord, subplot) %>%
    dplyr::select(-.ord)

  subplot_stats$subplot <- factor(subplot_stats$subplot, levels = subplot_stats$subplot)

  family_plot <- ggplot2::ggplot(
    family_counts,
    ggplot2::aes(x = stats::reorder(Family, count), y = count)
  ) +
    ggplot2::geom_col(fill = "#3d5941") +
    ggplot2::coord_flip() +
    ggplot2::theme_bw() +
    ggplot2::labs(
      title = lab$family_plot,
      x = NULL,
      y = lab$x_ind
    )

  species_plot <- ggplot2::ggplot(
    species_counts,
    ggplot2::aes(x = stats::reorder(Species, count), y = count)
  ) +
    ggplot2::geom_col(fill = "darkolivegreen") +
    ggplot2::coord_flip() +
    ggplot2::theme_bw() +
    ggplot2::labs(
      title = lab$species_plot,
      x = NULL,
      y = lab$x_ind
    )

  subplot_plot <- ggplot2::ggplot(
    subplot_stats,
    ggplot2::aes(x = subplot, y = pct)
  ) +
    ggplot2::geom_col(fill = "olivedrab4") +
    ggplot2::theme_bw() +
    ggplot2::labs(
      title = lab$subplot_plot,
      x = lab$subplot,
      y = lab$x_pct
    ) +
    ggplot2::theme(
      axis.text.x = ggplot2::element_text(angle = 90, vjust = 0.5, hjust = 1)
    )

  d_num <- suppressWarnings(as.numeric(df$D)) / 10
  d_num <- d_num[is.finite(d_num) & d_num > 0]

  if (length(d_num) > 0) {
    min_d <- floor(min(d_num, na.rm = TRUE))
    max_d <- ceiling(max(d_num, na.rm = TRUE))
    rng_d <- max_d - min_d

    # largura adaptativa para evitar classes demais
    class_width <- dplyr::case_when(
      rng_d <= 100 ~ 10,
      rng_d <= 200 ~ 20,
      rng_d <= 400 ~ 25,
      TRUE ~ 50
    )

    start_break <- floor(min_d / class_width) * class_width
    end_break <- ceiling(max_d / class_width) * class_width
    dbh_breaks <- seq(start_break, end_break + class_width, by = class_width)

    if (length(dbh_breaks) < 2) {
      dbh_breaks <- c(start_break, start_break + class_width)
    }

    dbh_labels <- paste0(
      "[",
      format(dbh_breaks[-length(dbh_breaks)], trim = TRUE, scientific = FALSE),
      ",",
      format(dbh_breaks[-1], trim = TRUE, scientific = FALSE),
      ifelse(seq_along(dbh_breaks[-1]) == length(dbh_breaks[-1]), "]", ")")
    )

    dbh_df <- data.frame(D_cm = d_num) %>%
      dplyr::mutate(
        dbh_class = cut(
          D_cm,
          breaks = dbh_breaks,
          right = FALSE,
          include.lowest = TRUE,
          labels = dbh_labels
        )
      ) %>%
      dplyr::count(dbh_class, name = "count", .drop = FALSE)

    dbh_df$dbh_class <- factor(dbh_df$dbh_class, levels = dbh_labels)

    dbh_plot <- ggplot2::ggplot(
      dbh_df,
      ggplot2::aes(x = dbh_class, y = count)
    ) +
      ggplot2::geom_col(fill = "#543005") +
      ggplot2::theme_bw() +
      ggplot2::labs(
        x = lab$dbh_x,
        y = lab$dbh_y
      ) +
      ggplot2::theme(
        axis.text.x = ggplot2::element_text(
          angle = 90,
          vjust = 0.5,
          hjust = 1,
          size = 8
        )
      )
  } else {
    dbh_plot <- ggplot2::ggplot() +
      ggplot2::theme_void() +
      ggplot2::labs(title = lab$dbh_plot)
  }

  if (is.data.frame(phytosoc$species_metrics) && nrow(phytosoc$species_metrics) > 0) {
    species_txt <- as.character(phytosoc$species_metrics$species)
    species_txt[is.na(species_txt) | !nzchar(trimws(species_txt))] <- "indet"

    species_metrics_tbl <- phytosoc$species_metrics %>%
      dplyr::transmute(
        col1 = vapply(
          species_txt,
          function(x) {
            if (nchar(x) > 35) paste0(substr(x, 1, 35), "...") else x
          },
          character(1)
        ),
        col2 = family,
        col3 = abundance,
        col4 = n_subplots,
        col5 = relative_density,
        col6 = relative_frequency,
        col7 = importance_value
      ) %>%
      utils::head(20)

    names(species_metrics_tbl) <- lab$species_tbl
  } else {
    species_metrics_tbl <- data.frame(matrix(
      nrow = 0,
      ncol = length(lab$species_tbl)
    ))
    names(species_metrics_tbl) <- lab$species_tbl
  }

  if (is.data.frame(phytosoc$family_metrics) && nrow(phytosoc$family_metrics) > 0) {
    family_metrics_tbl <- phytosoc$family_metrics %>%
      dplyr::transmute(
        col1 = family,
        col2 = abundance,
        col3 = n_species,
        col4 = n_subplots,
        col5 = relative_density,
        col6 = relative_frequency,
        col7 = importance_value
      ) %>%
      utils::head(20)

    names(family_metrics_tbl) <- lab$family_tbl
  } else {
    family_metrics_tbl <- data.frame(matrix(
      nrow = 0,
      ncol = length(lab$family_tbl)
    ))
    names(family_metrics_tbl) <- lab$family_tbl
  }

  # Build interactive plotly versions for HTML output
  family_plotly <- plotly::plot_ly(
    data = family_counts,
    y = stats::reorder(family_counts$Family, family_counts$count),
    x = family_counts$count,
    type = "bar", orientation = "h",
    marker = list(color = "#3d5941"),
    hovertext = paste0(family_counts$Family, ": ", family_counts$count),
    hoverinfo = "text"
  ) %>%
    plotly::layout(
      title = list(text = paste0("<b>", lab$family_plot, "</b>"), x = 0.5),
      xaxis = list(title = lab$x_ind),
      yaxis = list(title = ""),
      margin = list(l = 120, r = 10, t = 40, b = 40),
      plot_bgcolor = "white", paper_bgcolor = "white"
    ) %>%
    plotly::config(displaylogo = FALSE)

  species_plotly <- plotly::plot_ly(
    data = species_counts,
    y = stats::reorder(species_counts$Species, species_counts$count),
    x = species_counts$count,
    type = "bar", orientation = "h",
    marker = list(color = "darkolivegreen"),
    hovertext = paste0(species_counts$Species, ": ", species_counts$count),
    hoverinfo = "text"
  ) %>%
    plotly::layout(
      title = list(text = paste0("<b>", lab$species_plot, "</b>"), x = 0.5),
      xaxis = list(title = lab$x_ind),
      yaxis = list(title = ""),
      margin = list(l = 180, r = 10, t = 40, b = 40),
      plot_bgcolor = "white", paper_bgcolor = "white"
    ) %>%
    plotly::config(displaylogo = FALSE)

  subplot_plotly <- plotly::plot_ly(
    data = subplot_stats,
    x = subplot_stats$subplot,
    y = subplot_stats$pct,
    type = "bar",
    marker = list(color = "olivedrab4"),
    hovertext = paste0(lab$subplot, " ", subplot_stats$subplot, ": ", subplot_stats$pct, "%"),
    hoverinfo = "text"
  ) %>%
    plotly::layout(
      title = list(text = paste0("<b>", lab$subplot_plot, "</b>"), x = 0.5),
      xaxis = list(title = lab$subplot, tickangle = -90),
      yaxis = list(title = lab$x_pct),
      margin = list(l = 40, r = 10, t = 40, b = 60),
      plot_bgcolor = "white", paper_bgcolor = "white"
    ) %>%
    plotly::config(displaylogo = FALSE)

  if (exists("dbh_df") && is.data.frame(dbh_df) && nrow(dbh_df) > 0) {
    dbh_plotly <- plotly::plot_ly(
      data = dbh_df,
      x = dbh_df$dbh_class,
      y = dbh_df$count,
      type = "bar",
      marker = list(color = "#543005"),
      hovertext = paste0(dbh_df$dbh_class, ": ", dbh_df$count),
      hoverinfo = "text"
    ) %>%
      plotly::layout(
        title = list(text = paste0("<b>", lab$dbh_plot, "</b>"), x = 0.5),
        xaxis = list(title = lab$dbh_x, tickangle = -90),
        yaxis = list(title = lab$dbh_y),
        margin = list(l = 40, r = 10, t = 40, b = 80),
        plot_bgcolor = "white", paper_bgcolor = "white"
      ) %>%
      plotly::config(displaylogo = FALSE)
  } else {
    dbh_plotly <- NULL
  }

  list(
    metrics_tbl = metrics_tbl,
    species_metrics_tbl = species_metrics_tbl,
    family_metrics_tbl = family_metrics_tbl,
    family_plot = family_plot,
    species_plot = species_plot,
    subplot_plot = subplot_plot,
    dbh_plot = dbh_plot,
    family_plotly = family_plotly,
    species_plotly = species_plotly,
    subplot_plotly = subplot_plotly,
    dbh_plotly = dbh_plotly,
    phytosoc = phytosoc
  )
}


#' @keywords internal
#' @noRd
#'
`%||%` <- function(x, y) {
  if (is.null(x)) y else x
}


#' Create the R Markdown content for the plot PDF report
#'
#' @param subplot_plots List of subplot plot objects.
#' @param tf_col Logical vector indicating whether collected individuals exist.
#' @param tf_uncol Logical vector indicating whether uncollected individuals exist.
#' @param tf_palm Logical vector indicating whether uncollected palms exist.
#' @param plot_name Character plot name.
#' @param plot_code Character plot code.
#' @param spec_df Data frame used to build the checklist section.
#' @param language Language code. One of `"en"`, `"pt"`, `"es"` or `"ma"`.
#'
#' @return Character vector containing the full `.Rmd` content.
#'
#' @keywords internal
#' @noRd
#'
.create_rmd_content <- function(subplot_plots,
                                tf_col,
                                tf_uncol,
                                tf_palm,
                                plot_name,
                                plot_code,
                                spec_df,
                                dashboard = NULL,
                                language = "en") {
  if (missing(language) || is.null(language) || !length(language)) {
    language <- "en"
  }

  language <- tolower(trimws(as.character(language)[1]))
  if (!language %in% c("en", "pt", "es", "fr", "ma", "pa")) {
    warning(sprintf("Invalid language '%s'. Using 'en'.", language), call. = FALSE)
    language <- "en"
  }

  dict <- tibble::tibble(
    en = c(
      "Full Plot Report",
      "MONITORA Program",
      "Back to Contents",
      "Contents",
      "Metadata",
      "Plot Name",
      "Plot Code",
      "Census No",
      "Team",
      "Total Specimens",
      "Collected (excluding palms)",
      "Not Collected (excluding palms)",
      "Palms (Arecaceae)",
      "Dead trees since first census",
      "Recruits since first census",
      "Dashboard",
      "Metric Summary",
      "Most Common Families",
      "Most Abundant Species",
      "Collection Percentage by Subplot",
      "DBH Classes",
      "DBH Class",
      "Number of individuals",
      "Species Metrics",
      "Family Metrics",
      "Species",
      "Family",
      "Abundance",
      "Subplots",
      "Rel. density (%)",
      "Rel. frequency (%)",
      "IV",
      "Richness",
      "General Plot",
      "Collected Only",
      "Not Collected Palms",
      "Not Collected",
      "Palms",
      "Subplot Index",
      "Subplots",
      "Subplot ",
      "Checklist",
      "Individual Subplots"
    ),

    pt = c(
      "Relat\u00f3rio Completo da Parcela = Full Plot Report",
      "Programa MONITORA = MONITORA Program",
      "Voltar ao Sum\u00e1rio = Back to Contents",
      "Sum\u00e1rio = Contents",
      "Metadados = Metadata",
      "Nome da Parcela = Plot Name",
      "C\u00f3digo da Parcela = Plot Code",
      "N\u00famero de Censos = Census No",
      "Equipe = Team",
      "Total de Indiv\u00edduos = Total Specimens",
      "Coletados (excluindo palmeiras) = Collected (excluding palms)",
      "N\u00e3o Coletados (excluindo palmeiras) = Not Collected (excluding palms)",
      "Palmeiras (Arecaceae) = Palms (Arecaceae)",
      "\u00c1rvores mortas desde o primeiro censo = Dead trees since first census",
      "Recrutas desde o primeiro censo = Recruits since first census",
      "Dashboard = Dashboard",
      "Resumo de M\u00e9tricas = Metric Summary",
      "Fam\u00edlias Mais Comuns = Most Common Families",
      "Esp\u00e9cies Mais Abundantes = Most Abundant Species",
      "Percentual de Coleta por Subparcela = Collection Percentage by Subplot",
      "Classes de DAP = DBH Classes",
      "Classe de DAP = DBH Class",
      "N\u00famero de indiv\u00edduos = Number of individuals",
      "M\u00e9tricas de Esp\u00e9cies = Species Metrics",
      "M\u00e9tricas de Fam\u00edlias = Family Metrics",
      "Esp\u00e9cies = Species",
      "Fam\u00edlia = Family",
      "Abund\u00e2ncia = Abundance",
      "Subparcelas = Subplots",
      "Dens. rel. (%) = Rel. density (%)",
      "Freq. rel. (%) = Rel. frequency (%)",
      "VI = IV",
      "Riqueza = Richness",
      "Mapa Geral da Parcela = General Plot",
      "Apenas Coletados = Collected Only",
      "Palmeiras N\u00e3o Coletadas = Not Collected Palms",
      "N\u00e3o Coletados = Not Collected",
      "Palmeiras = Palms",
      "\u00cdndice de Subparcelas = Subplot Index",
      "Subparcelas = Subplots",
      "Subparcela  = Subplot  ",
      "Lista de Esp\u00e9cies = Checklist",
      "Subparcelas Individuais = Individual Subplots"
    ),

    es = c(
      "Informe completo de la parcela = Full Plot Report",
      "Programa MONITORA = MONITORA Program",
      "Volver al Contenido = Back to Contents",
      "Contenido = Contents",
      "Metadatos = Metadata",
      "Nombre de la Parcela = Plot Name",
      "C\u00f3digo de la Parcela = Plot Code",
      "N\u00famero de Censos = Census No",
      "Equipo = Team",
      "Total de Individuos = Total Specimens",
      "Colectados (excluyendo palmas) = Collected (excluding palms)",
      "No Colectados (excluyendo palmas) = Not Collected (excluding palms)",
      "Palmas (Arecaceae) = Palms (Arecaceae)",
      "\u00c1rboles muertos desde el primer censo = Dead trees since first census",
      "Reclutas desde el primer censo = Recruits since first census",
      "Dashboard = Dashboard",
      "Resumen de M\u00e9tricas = Metric Summary",
      "Familias M\u00e1s Comunes = Most Common Families",
      "Especies M\u00e1s Abundantes = Most Abundant Species",
      "Porcentaje de Recolecci\u00f3n por Subparcela = Collection Percentage by Subplot",
      "Clases de DAP = DBH Classes",
      "Clase de DAP = DBH Class",
      "N\u00famero de individuos = Number of individuals",
      "M\u00e9tricas de Especies = Species Metrics",
      "M\u00e9tricas de Familias = Family Metrics",
      "Especies = Species",
      "Familia = Family",
      "Abundancia = Abundance",
      "Subparcelas = Subplots",
      "Dens. rel. (%) = Rel. density (%)",
      "Freq. rel. (%) = Rel. frequency (%)",
      "VI = IV",
      "Riqueza = Richness",
      "Mapa General de la Parcela = General Plot",
      "Solo Colectados = Collected Only",
      "Palmas No Colectadas = Not Collected Palms",
      "No Colectados = Not Collected",
      "Palmas = Palms",
      "\u00cdndice de Subparcelas = Subplot Index",
      "Subparcelas = Subplots",
      "Subparcela  = Subplot ",
      "Lista de Especies = Checklist",
      "Subparcelas Individuales = Individual Subplots"
    ),

    fr = c(
      "Rapport Complet de la Parcelle = Full Plot Report",
      "Programme MONITORA = MONITORA Program",
      "Retour au Sommaire = Back to Contents",
      "Sommaire = Contents",
      "M\u00e9tadonn\u00e9es = Metadata",
      "Nom de la Parcelle = Plot Name",
      "Code de la Parcelle = Plot Code",
      "Num\u00e9ro de Recensement = Census No",
      "\u00c9quipe = Team",
      "Total des Individus = Total Specimens",
      "Collect\u00e9s (hors palmiers) = Collected (excluding palms)",
      "Non Collect\u00e9s (hors palmiers) = Not Collected (excluding palms)",
      "Palmiers (Arecaceae) = Palms (Arecaceae)",
      "Arbres morts depuis le premier recensement = Dead trees since first census",
      "Recrues depuis le premier recensement = Recruits since first census",
      "Tableau de bord = Dashboard",
      "R\u00e9sum\u00e9 des M\u00e9triques = Metric Summary",
      "Familles les Plus Communes = Most Common Families",
      "Esp\u00e8ces les Plus Abondantes = Most Abundant Species",
      "Pourcentage de Collecte par Sous-parcelle = Collection Percentage by Subplot",
      "Classes de DHP = DBH Classes",
      "Classe de DHP = DBH Class",
      "Nombre d'individus = Number of individuals",
      "M\u00e9triques des Esp\u00e8ces = Species Metrics",
      "M\u00e9triques des Familles = Family Metrics",
      "Esp\u00e8ces = Species",
      "Famille = Family",
      "Abondance = Abundance",
      "Sous-parcelles = Subplots",
      "Densit\u00e9 rel. (%) = Rel. density (%)",
      "Fr\u00e9q. rel. (%) = Rel. frequency (%)",
      "VI = IV",
      "Richesse = Richness",
      "Carte G\u00e9n\u00e9rale de la Parcelle = General Plot",
      "Collect\u00e9s Seulement = Collected Only",
      "Palmiers Non Collect\u00e9s = Not Collected Palms",
      "Non Collect\u00e9s = Not Collected",
      "Palmiers = Palms",
      "Index des Sous-parcelles = Subplot Index",
      "Sous-parcelles = Subplots",
      "Sous-parcelle  = Subplot ",
      "Liste des Esp\u00e8ces = Checklist",
      "Sous-parcelles Individuelles = Individual Subplots"
    ),

    ma = c(
      "\u6837\u5730\u5b8c\u6574\u62a5\u544a = Full Plot Report",
      "MONITORA \u9879\u76ee = MONITORA Program",
      "\u8fd4\u56de\u76ee\u5f55 = Back to Contents",
      "\u76ee\u5f55 = Contents",
      "\u5143\u6570\u636e = Metadata",
      "\u6837\u5730\u540d\u79f0 = Plot Name",
      "\u6837\u5730\u4ee3\u7801 = Plot Code",
      "\u666e\u67e5\u7f16\u53f7 = Census No",
      "\u56e2\u961f = Team",
      "\u4e2a\u4f53\u603b\u6570 = Total Specimens",
      "\u5df2\u91c7\u96c6\uff08\u4e0d\u542b\u68d5\u6988\u79d1\uff09 = Collected (excluding palms)",
      "\u672a\u91c7\u96c6\uff08\u4e0d\u542b\u68d5\u6988\u79d1\uff09 = Not Collected (excluding palms)",
      "\u68d5\u6988\u79d1\u4e2a\u4f53 (Arecaceae) = Palms (Arecaceae)",
      "\u81ea\u7b2c\u4e00\u6b21\u666e\u67e5\u4ee5\u6765\u6b7b\u4ea1\u7684\u6811\u6728 = Dead trees since first census",
      "\u81ea\u7b2c\u4e00\u6b21\u666e\u67e5\u4ee5\u6765\u65b0\u589e\u4e2a\u4f53 = Recruits since first census",
      "\u4eea\u8868\u677f = Dashboard",
      "\u6307\u6807\u6c47\u603b = Metric Summary",
      "\u6700\u5e38\u89c1\u79d1 = Most Common Families",
      "\u6700\u4e30\u5bcc\u7269\u79cd = Most Abundant Species",
      "\u5404\u5b50\u6837\u5730\u91c7\u96c6\u767e\u5206\u6bd4 = Collection Percentage by Subplot",
      "\u80f8\u5f84\u7b49\u7ea7 = DBH Classes",
      "\u80f8\u5f84\u7b49\u7ea7 = DBH Class",
      "\u4e2a\u4f53\u6570\u91cf = Number of individuals",
      "\u7269\u79cd\u6307\u6807 = Species Metrics",
      "\u79d1\u6307\u6807 = Family Metrics",
      "\u7269\u79cd = Species",
      "\u79d1 = Family",
      "\u4e30\u5ea6 = Abundance",
      "\u5b50\u6837\u5730 = Subplots",
      "\u76f8\u5bf9\u5bc6\u5ea6 (%) = Rel. density (%)",
      "\u76f8\u5bf9\u9891\u7387 (%) = Rel. frequency (%)",
      "\u91cd\u8981\u503c = IV",
      "\u4e30\u5bcc\u5ea6 = Richness",
      "\u6837\u5730\u603b\u4f53\u5730\u56fe = General Plot",
      "\u4ec5\u5df2\u91c7\u96c6\u4e2a\u4f53 = Collected Only",
      "\u672a\u91c7\u96c6\u68d5\u6988\u79d1\u4e2a\u4f53 = Not Collected Palms",
      "\u672a\u91c7\u96c6\u4e2a\u4f53 = Not Collected",
      "\u68d5\u6988\u79d1 = Palms",
      "\u5b50\u6837\u5730\u7d22\u5f15 = Subplot Index",
      "\u5b50\u6837\u5730 = Subplots",
      "\u5b50\u6837\u5730  = Subplot ",
      "\u7269\u79cd\u6e05\u5355 = Checklist",
      "\u5355\u4e2a\u5b50\u6837\u5730 = Individual Subplots"
    ),

    pa = c(
      "Hokjya r\u00ea t\u00e3waj\u00e3ri p\u00e2p\u00e3\u00e3 r\u00eata kur\u00e2ri p\u00e2ri h\u00e3 = Full Plot Report",
      "Programa MONITORA = MONITORA Program",
      "Pikjatit\u00e3 t\u00e4 sokkjaraa = Back to Contents",
      "T\u00e4 Sokkjaraa = Contents",
      "R\u00ea raa san r\u00ea kuk\u00e2ri = Metadata",
      "Issi r\u00ea t\u00e3 kuk\u00e2ri = Plot Name",
      "Kypa kuk\u00e2ri = Plot Code",
      "Junti h\u1ebd r\u00f5 s\u00ean p\u00e2rikran = Census No",
      "S\u00e2p\u00ear\u00e3t\u00ea = Team",
      "P\u00e3p\u00e3 p\u00e2ri = Total Specimens",
      "P\u00e2ri sonswa = Collected (excluding palms)",
      "P\u00e2ri r\u00f5r\u0129 = Not Collected (excluding palms)",
      "Kwatis\u00f4m\u1ebdra = Palms (Arecaceae)",
      "P\u00e2ri m\u00e3m\u00e3 jy ty = Dead trees since first census",
      "P\u00e2rituem~era krep\u00e2\u00e2 s\u00e2\u00e2 = Recruits since first census",
      "P\u00e3p\u00e3\u00e3 kja skreeha kuk\u00e2ra krep\u00e3\u00e3 s\u00e2\u00e2 = Dashboard",
      "Swan kukrek\u00e2ra = Metric Summary",
      "Inkj\u00eati tip\u0129njakjura = Most Common Families",
      "Sotinkj\u00eati sop\u00e2ri m\u1ebdra = Most Abundant Species",
      "Jyy ti he p\u00e2ri m\u1ebdra kar\u00ear\u00e2kjan kuk\u00e2ri s\u00e2\u00e2  = Collection Percentage by Subplot",
      "Classes de DAP = DBH Classes",
      "Classe de DAP = DBH Class",
      "N\u00famero de p\u00e2ri = Number of individuals",
      "P\u0129rak\u00e2ri p\u00e2ri m\u1ebdra = Species Metrics",
      "P\u0129rak\u00e2ri p\u00e2ri kyapi\u00e2hapi\u00e2ra = Family Metrics",
      "P\u0129rak\u00e2ri = Species",
      "P\u0129rak\u00e2ri kyapi\u00e2hapi\u00e2ra = Family",
      "Abund\u00e2ncia = Abundance",
      "Kuk\u00e2ra krep\u00e3\u00e3 s\u00e2\u00e2 = Subplots",
      "Dens. rel. (%) = Rel. density (%)",
      "Freq. rel. (%) = Rel. frequency (%)",
      "VI = IV",
      "Sop\u00e2ri m\u1ebdra = Richness",
      "P\u00e3p\u00e3 kypapr\u1ebdpi kuk\u00e2ri = General Plot",
      "P\u00e2ri sonswa kypa kuk\u00e2ri kran = Collected Only",
      "R\u00f5\u00f5rin kwatis\u00f4m\u00eara kuk\u00e2ri kran = Not Collected Palms",
      "P\u00e2ri r\u00f5r\u0129 kypa kuk\u00e2ri kran = Not Collected",
      "Kwatis\u00f4m\u1ebdra = Palms",
      "T\u00e4 sokkjaraa r\u00ea t\u00e2 kuk\u00e2ra kran = Subplot Index",
      "Kuk\u00e2ra krep\u00e3\u00e3 s\u00e2\u00e2 = Subplots",
      "Kuk\u00e2ra krep\u00e3\u00e3 s\u00e2\u00e2   = Subplot ",
      "Issi pyr\u00e3h\u00e3 p\u00e2rijnsim\u1ebdra = Checklist",
      "Kuk\u00e2ra krep\u00e3\u00e3 s\u00e2\u00e2 pyti = Individual Subplots"
    )
  )

  .tr <- function(en_key) {
    if (language == "en") return(en_key)
    idx <- match(en_key, dict$en)
    if (is.na(idx)) return(en_key)
    raw <- dict[[language]][idx]
    if (is.na(raw) || !nzchar(raw)) return(en_key)
    trimws(sub("\\s*=\\s*.*$", "", raw))
  }

  .translate_lines <- function(lines, dict, lang = c("en", "pt", "es", "fr", "ma", "pa")) {
    lang <- match.arg(lang)

    if (lang == "en") {
      return(lines)
    }

    clean_translation <- function(x) {
      sub("=\\s+.*$", "", x)
    }

    dict_translated <- dict
    dict_translated[[lang]] <- vapply(dict_translated[[lang]], clean_translation, character(1))

    out <- lines
    in_r_chunk <- FALSE
    ord <- order(nchar(dict_translated$en), decreasing = TRUE)

    for (j in seq_along(lines)) {
      line <- lines[j]

      if (grepl("^```\\{r", line)) {
        in_r_chunk <- TRUE
        out[j] <- line
        next
      }

      if (grepl("^```\\s*$", line) && in_r_chunk) {
        in_r_chunk <- FALSE
        out[j] <- line
        next
      }

      if (in_r_chunk) {
        out[j] <- line
        next
      }

      for (i in ord) {
        line <- gsub(dict_translated$en[[i]], dict_translated[[lang]][[i]], line, fixed = TRUE)
      }

      out[j] <- line
    }

    out
  }
  .chunk_id_factory <- local({
    i <- 0L
    function(prefix = "chunk") {
      i <<- i + 1L
      paste0(prefix, "-", i)
    }
  })

  .pagebreak_block <- function(prefix = "pagebreak") {
    id <- .chunk_id_factory(prefix)
    c(
      sprintf("```{r %s, echo=FALSE, results='asis'}", id),
      "if (knitr::is_latex_output()) {",
      "  cat('\\\\newpage\\n')",
      "} else if (knitr::is_html_output()) {",
      "  cat('<div style=\"page-break-after: always;\"></div>\\n')",
      "}",
      "```"
    )
  }

  .separator_block <- function(prefix = "separator") {
    id <- .chunk_id_factory(prefix)
    c(
      sprintf("```{r %s, echo=FALSE, results='asis'}", id),
      "if (knitr::is_latex_output()) {",
      "  cat('\\\\vspace{0.6\\\\baselineskip}\\n')",
      "  cat('\\\\hrule\\n')",
      "} else if (knitr::is_html_output()) {",
      "  cat('<hr style=\"margin-top: 1rem; margin-bottom: 1rem;\">\\n')",
      "}",
      "```"
    )
  }

  .contents_block <- function(prefix = "contentsfmt") {
    id <- .chunk_id_factory(prefix)
    c(
      sprintf("```{r %s, echo=FALSE, results='asis'}", id),
      "if (knitr::is_latex_output()) {",
      "  cat('\\\\vspace*{-1cm}\\n')",
      "  cat('\\\\thispagestyle{plain}\\n')",
      "  cat('\\\\tableofcontents\\n')",
      "  cat('\\\\newpage\\n')",
      "} else if (knitr::is_html_output()) {",
      "  cat('<div style=\"margin-top:-0.5rem;\"></div>\\n')",
      "}",
      "```"
    )
  }

  .tiny_nav_block <- function(nav_targets, prefix = "navlinks") {
    id <- .chunk_id_factory(prefix)
    nav_text <- paste(nav_targets, collapse = " | ")
    c(
      sprintf("```{r %s, echo=FALSE, results='asis'}", id),
      sprintf("nav_text <- %s", deparse(nav_text)),
      "if (knitr::is_latex_output()) {",
      "  cat('\\\\begingroup\\\\tiny\\\\color{gray}\\n')",
      "  cat(paste0('\u00ab ', nav_text, '\\n'))",
      "  cat('\\\\endgroup\\n')",
      "} else if (knitr::is_html_output()) {",
      "  cat(sprintf(",
      "    '<div style=\"font-size:0.8em; color:#666; margin-top:0.5rem; margin-bottom:0.5rem;\">&laquo; %s</div>\\n',",
      "    nav_text",
      "  ))",
      "}",
      "```"
    )
  }

  title_txt <- dict[[language]][match("Full Plot Report", dict$en)]
  if (is.na(title_txt) || !nzchar(title_txt)) title_txt <- "Full Plot Report"
  title_txt <- trimws(sub("\\s+=\\s+.*$", "", title_txt))
  title_txt <- gsub("'", "\\\\'", title_txt)

  yaml_head <- c(
    "---",
    "output:",
    "  html_document:",
    "    toc: true",
    "    toc_depth: 2",
    "    number_sections: true",
    "    df_print: paged",
    "  pdf_document:",
    "    toc: true",
    "    toc_depth: 2",
    "    number_sections: true",
    "fontsize: 12pt",
    "params:",
    "  input_type: \"forestplots\"",
    "  language: \"en\"",
    "  metadata: NULL",
    "  main_plot: NULL",
    "  interactive_main_plot: NULL",
    "  collected_plot: NULL",
    "  interactive_collected_plot: NULL",
    "  uncollected_plot: NULL",
    "  interactive_uncollected_plot: NULL",
    "  uncollected_palm_plot: NULL",
    "  interactive_palm_plot: NULL",
    "  subplots_list: NULL",
    "  subplot_size: 70",
    "  stats: NULL",
    "  dashboard: NULL",
    "  tag_list: NULL",
    "  tag_to_subplot: NULL",
    "---"
  )

  mid_section <- c(
    "```{r setup, include=FALSE}",
    "knitr::opts_chunk$set(echo = FALSE, warning = FALSE, message = FALSE)",
    "library(ggplot2)",
    "library(dplyr)",
    "library(knitr)",
    "library(plotly)",
    "library(htmlwidgets)",
    "options(plotly_autotypenumbers = FALSE)",
    "```",
    "",
    "```{r html-style, echo=FALSE, results='asis'}",
    "if (knitr::is_html_output()) {",
    "  cat('",
    "<style>",
    "  * { -webkit-tap-highlight-color: rgba(0,0,0,0); }",
    "  body { margin: 0; padding: 8px; font-size: 14px; }",
    "  .main-container { max-width: 100% !important; padding: 0 !important; margin: 0 !important; }",
    "  .plotly, .js-plotly-plot { ",
    "    width: 100% !important; ",
    "    height: auto !important; ",
    "    min-height: 700px !important; ",
    "    max-width: 100vw !important; ",
    "    margin-bottom: 20px !important; ",
    "  }",
    "  .plotly .main-svg, .js-plotly-plot .main-svg {",
    "    width: 100% !important;",
    "    height: auto !important;",
    "    min-height: 700px !important;",
    "  }",
    "  .svg-container { ",
    "    width: 100% !important; ",
    "    height: auto !important;",
    "    min-height: 700px !important;",
    "    overflow-x: auto !important; ",
    "    -webkit-overflow-scrolling: touch; ",
    "  }",
    "  .plot-container {",
    "    width: 100%;",
    "    overflow-x: auto;",
    "    overflow-y: hidden;",
    "    margin: 20px 0;",
    "  }",
    "  table { display: block; overflow-x: auto; white-space: nowrap; -webkit-overflow-scrolling: touch; }",
    "  a, button, .plotly .modebar-btn { min-height: 44px; min-width: 44px; }",
    "  @media (max-width: 768px) {",
    "    h1 { font-size: 1.4rem !important; }",
    "    h2 { font-size: 1.2rem !important; }",
    "    h3 { font-size: 1.0rem !important; }",
    "    body, p, li, .table, .kable { font-size: 12px !important; }",
    "    pre, code { font-size: 10px !important; white-space: pre-wrap !important; }",
    "    .nav-links, .tiny-nav { font-size: 10px !important; }",
    "    .plotly, .js-plotly-plot { min-height: 550px !important; }",
    "    .svg-container { min-height: 550px !important; }",
    "  }",
    "</style>",
    "')",
    "}",
    "```",
    "",
    "```{r setup2, include=FALSE}",
    "spec_df <- params$stats$spec_df",
    "safe_id <- function(x) {",
    "  x <- trimws(as.character(x))",
    "  x <- gsub('\\\\s*\\\\(.*\\\\)$', '', x)",
    "  x <- gsub('[^A-Za-z0-9]', '', x)",
    "  tolower(x)",
    "}",
    "safe_num <- function(x) suppressWarnings(as.numeric(as.character(x)))",
    "norm_path <- function(p) {",
    "  p <- as.character(p)",
    "  p[is.na(p)] <- ''",
    "  tryCatch(normalizePath(p, winslash = '/', mustWork = FALSE), error = function(e) p)",
    "}",
    "find_logo <- function(fname) {",
    "  cand <- c(",
    "    system.file('figures', fname, package = 'forplotR'),",
    "    file.path('figures', fname),",
    "    file.path(getwd(), 'figures', fname)",
    "  )",
    "  cand <- norm_path(cand)",
    "  cand <- cand[nzchar(cand) & file.exists(cand)]",
    "  if (length(cand)) cand[1] else ''",
    "}",
    "```",
    "",
    "```{r title-block, echo=FALSE, results='asis'}",
    "if (knitr::is_latex_output()) {",
    "  cat('\\\\begin{center}\\n')",
    sprintf("  cat('\\\\Huge\\\\textbf{%s} \\\\\\\\ \\n')", title_txt),
    "  cat('\\\\vspace{0.5em}\\n')",
    "  cat(paste0('\\\\normalsize ', params$metadata$plot_name, ' | ', params$metadata$plot_code, ' \\\\\\\\ \\n'))",
    "  cat('\\\\end{center}\\n')",
    "} else if (knitr::is_html_output()) {",
    "  cat('<div style=\"text-align:center; margin-top:1rem; margin-bottom:1.5rem;\">\\n')",
    sprintf("  cat('<h1 style=\"margin-bottom:0.3rem;\">%s</h1>\\n')", title_txt),
    "  cat(sprintf('<p style=\"margin-top:0;\">%s | %s</p>\\n', params$metadata$plot_name, params$metadata$plot_code, params$metadata$plot_census_no_fp))",
    "  cat('</div>\\n')",
    "}",
    "```",
    "",
    "```{r logo, echo=FALSE, results='asis'}",
    "forplotr_logo_path <- find_logo('forplotR_hex_sticker.png')",
    "forestplots_logo_path <- find_logo('forestplotsnet_logo.png')",
    "monitora_logo_path <- find_logo('monitora_logo.png')",
    "",
    "raw_it <- if (!is.null(params$input_type)) {",
    "  params$input_type",
    "} else if (!is.null(params$metadata$input_type)) {",
    "  params$metadata$input_type",
    "} else {",
    "  ''",
    "}",
    "",
    "input_type <- tolower(trimws(as.character(raw_it)))",
    "lang <- if (!is.null(params$language)) tolower(trimws(as.character(params$language))) else ''",
    "is_monitora <- identical(input_type, 'monitora')",
    "is_pa <- identical(input_type, 'field_sheet_ti')",
    "show_partner_logo <- !is_pa",
    "",
    "pa_logo_paths <- c(",
    "  find_logo('jbrj.png'),",
    "  find_logo('iakio.png'),",
    "  find_logo('ci.png'),",
    "  find_logo('isa.png')",
    ")",
    "",
    "pa_logo_hrefs <- c(",
    "  'https://www.gov.br/jbrj/pt-br',",
    "  'https://www.instagram.com/associacao.iakio/#',",
    "  'https://brasil.conservation.org/',",
    "  'https://www.socioambiental.org/'",
    ")",
    "",
    "partner_logo_path <- if (is_monitora) monitora_logo_path else forestplots_logo_path",
    "partner_href <- if (is_monitora) {",
    "  'https://www.gov.br/icmbio/pt-br/assuntos/monitoramento/programa-monitora'",
    "} else {",
    "  'https://forestplots.net'",
    "}",
    "partner_width <- if (is_monitora) '0.24\\\\linewidth' else '0.32\\\\linewidth'",
    "",
    "if (knitr::is_latex_output()) {",
    "  cat('\\\\vspace{0.6cm}\\n')",
    "  cat('\\\\begin{center}\\n')",
    "",
    "  if (nzchar(forplotr_logo_path) && file.exists(forplotr_logo_path)) {",
    "    cat(sprintf(",
    "      '\\\\href{https://dboslab.github.io/forplotR-website/}{\\\\includegraphics[width=0.18\\\\linewidth]{%s}}',",
    "      forplotr_logo_path",
    "    ), '\\n')",
    "  } else {",
    "    cat('{\\\\Large\\\\textbf{forplotR}}\\n')",
    "  }",
    "",
    "  cat('\\\\vspace{0.25cm}\\n')",
    "  cat('\\\\begin{minipage}{\\\\textwidth}\\\\centering\\n')",
    "",
    "  if (show_partner_logo && nzchar(partner_logo_path) && file.exists(partner_logo_path)) {",
    "    cat(sprintf(",
    "      '\\\\href{%s}{\\\\includegraphics[width=%s,keepaspectratio]{%s}}',",
    "      partner_href, partner_width, partner_logo_path",
    "    ), '\\n')",
    "  } else if (show_partner_logo) {",
    "    cat(sprintf('{\\\\large %s}\\n',",
    "      if (is_monitora) 'MONITORA Program' else 'ForestPlots.net'))",
    "  }",
    "",
    "  if (is_pa && length(pa_logo_paths) > 0) {",
    "    cat('\\\\vspace{0.25cm}\\n')",
    "    for (i in seq_along(pa_logo_paths)) {",
    "      if (nzchar(pa_logo_paths[i]) && file.exists(pa_logo_paths[i])) {",
    "        cat(sprintf(",
    "          '\\\\href{%s}{\\\\includegraphics[width=0.14\\\\linewidth,keepaspectratio]{%s}}',",
    "          pa_logo_hrefs[i], pa_logo_paths[i]",
    "        ))",
    "        if (i < length(pa_logo_paths)) cat('\\\\hspace{0.25cm}\\n')",
    "      }",
    "    }",
    "    cat('\\n')",
    "  }",
    "",
    "  cat('\\\\end{minipage}\\n')",
    "  cat('\\\\end{center}\\n')",
    "",
    "} else if (knitr::is_html_output()) {",
    "  cat('<div style=\"text-align:center; margin-top:1rem; margin-bottom:1rem;\">\\n')",
    "",
    "  if (nzchar(forplotr_logo_path) && file.exists(forplotr_logo_path)) {",
    "    cat(sprintf(",
    "      '<a href=\"https://dboslab.github.io/forplotR-website/\"><img src=\"%s\" style=\"width:140px; margin-bottom:0.6rem;\"></a><br/>',",
    "      forplotr_logo_path",
    "    ))",
    "  } else {",
    "    cat('<div style=\"font-size:1.6rem; font-weight:bold; margin-bottom:0.6rem;\">forplotR</div>\\n')",
    "  }",
    "",
    "  cat('<div style=\"margin-top:0.2rem;\">\\n')",
    "",
    "  if (show_partner_logo && nzchar(partner_logo_path) && file.exists(partner_logo_path)) {",
    "    cat(sprintf(",
    "      '<a href=\"%s\"><img src=\"%s\" style=\"width:180px; margin:0.3rem 0.6rem; vertical-align:middle;\"></a>',",
    "      partner_href, partner_logo_path",
    "    ))",
    "  } else if (show_partner_logo) {",
    "    cat(sprintf('<span style=\"font-size:1.1rem; margin:0 0.8rem;\">%s</span>',",
    "      if (is_monitora) 'MONITORA Program' else 'ForestPlots.net'))",
    "  }",
    "",
    "  if (is_pa && length(pa_logo_paths) > 0) {",
    "    cat('<div style=\"margin-top:0.5rem;\">\\n')",
    "    for (i in seq_along(pa_logo_paths)) {",
    "      if (nzchar(pa_logo_paths[i]) && file.exists(pa_logo_paths[i])) {",
    "        cat(sprintf(",
    "          '<a href=\"%s\"><img src=\"%s\" style=\"width:95px; margin:0.2rem 0.4rem; vertical-align:middle;\"></a>',",
    "          pa_logo_hrefs[i], pa_logo_paths[i]",
    "        ))",
    "      }",
    "    }",
    "    cat('</div>\\n')",
    "  }",
    "",
    "  cat('</div>\\n')",
    "  cat('</div>\\n')",
    "}",
    "```",
    "",
    "```{r all-tag-anchors, results='asis', echo=FALSE}",
    "t2s_df <- params$tag_to_subplot",
    "if (!is.null(t2s_df) && nrow(t2s_df) > 0) {",
    "  tag_vec <- as.character(t2s_df$`New Tag No`)",
    "  tag_vec <- trimws(tag_vec)",
    "  tag_vec <- tag_vec[!is.na(tag_vec) & nzchar(tag_vec)]",
    "  subplot_vec <- t2s_df$T1[match(tag_vec, t2s_df$`New Tag No`)]",
    "  ids <- sprintf('tag-%s-S%s', tag_vec, subplot_vec)",
    "  ids <- unique(ids)",
    "  if (knitr::is_latex_output()) {",
    "    for (id in ids) cat(sprintf('\\\\hypertarget{%s}{}', id), '\\n')",
    "  } else if (knitr::is_html_output()) {",
    "    for (id in ids) cat(sprintf('<div id=\"%s\"></div>', id), '\\n')",
    "  }",
    "}",
    "```",
    "",
    "## Contents {#contents}",
    .contents_block(),
    "",
    "## Metadata {#metadata}",
    "",
    paste0("**", .tr("Plot Name"), ":** ", "`r params$metadata$plot_name`"),
    "",
    paste0("**", .tr("Plot Code"), ":** ", "`r params$metadata$plot_code`"),
    "",
    paste0("**", .tr("Census No"), ":** ", "`r params$metadata$plot_census_no_fp`"),
    "",
    paste0("**", .tr("Team"), ":** ", "`r params$metadata$team`"),
    "",
    .separator_block(),
    "",
    "```{r meta-census, echo=FALSE, results='asis'}",
    "yrs <- params$metadata$census_years",
    "yrs <- yrs[is.finite(yrs)]",
    "if (length(yrs) > 1) {",
    "  cat(paste0('**Number of Census:** ', length(yrs), '\\n\\n'))",
    "  cat(paste0('**Dates of Census:** ', paste(yrs, collapse = ' | '), '\\n\\n'))",
    "}",
    "```",
    "",
    .pagebreak_block(),
    "",
    "## Dashboard {#dashboard}",
    "",
    "### Metric Summary",
    "",
    "```{r dashboard-metrics, echo=FALSE}",
    "if (!is.null(params$dashboard$metrics_tbl)) knitr::kable(params$dashboard$metrics_tbl, align = c('l', 'r'))",
    "```",
    "",
    "### Most Common Families",
    "",
    "```{r dashboard-family-plot, fig.width=10, fig.height=6, fig.align='center', out.width='95%'}",
    "if (knitr::is_html_output() && !is.null(params$dashboard$family_plotly)) {",
    "  params$dashboard$family_plotly",
    "} else if (!is.null(params$dashboard$family_plot)) {",
    "  params$dashboard$family_plot",
    "}",
    "```",
    "",
    "### Most Abundant Species",
    "",
    "```{r dashboard-species-plot, fig.width=10, fig.height=6, fig.align='center', out.width='95%'}",
    "if (knitr::is_html_output() && !is.null(params$dashboard$species_plotly)) {",
    "  params$dashboard$species_plotly",
    "} else if (!is.null(params$dashboard$species_plot)) {",
    "  params$dashboard$species_plot",
    "}",
    "```",
    "",
    "### Collection Percentage by Subplot",
    "",
    "```{r dashboard-subplot-plot, fig.width=10, fig.height=6, fig.align='center', out.width='95%'}",
    "if (knitr::is_html_output() && !is.null(params$dashboard$subplot_plotly)) {",
    "  params$dashboard$subplot_plotly",
    "} else if (!is.null(params$dashboard$subplot_plot)) {",
    "  params$dashboard$subplot_plot",
    "}",
    "```",
    "",
    "### DBH Classes",
    "",
    "```{r dashboard-dbh-plot, fig.width=10, fig.height=6, fig.align='center', out.width='95%'}",
    "if (knitr::is_html_output() && !is.null(params$dashboard$dbh_plotly)) {",
    "  params$dashboard$dbh_plotly",
    "} else if (!is.null(params$dashboard$dbh_plot)) {",
    "  params$dashboard$dbh_plot",
    "}",
    "```",
    "",
    "### Species Metrics",
    "",
    "```{r dashboard-species-table, echo=FALSE, results='asis'}",
    "if (!is.null(params$dashboard$species_metrics_tbl)) {",
    "  if (knitr::is_latex_output()) {",
    "    cat('\\\\begingroup\\\\fontsize{6}{7}\\\\selectfont\\n')",
    "    print(knitr::kable(",
    "      params$dashboard$species_metrics_tbl,",
    "      align = 'l',",
    "      format = 'latex',",
    "      booktabs = TRUE,",
    "      longtable = FALSE",
    "    ))",
    "    cat('\\\\endgroup\\n')",
    "  } else {",
    "    knitr::kable(params$dashboard$species_metrics_tbl, align = 'l')",
    "  }",
    "}",
    "```",
    "",
    "### Family Metrics",
    "",
    "```{r dashboard-family-table, echo=FALSE, results='asis'}",
    "if (!is.null(params$dashboard$family_metrics_tbl)) {",
    "  if (knitr::is_latex_output()) {",
    "    cat('\\\\begingroup\\\\fontsize{6}{7}\\\\selectfont\\n')",
    "    print(knitr::kable(",
    "      params$dashboard$family_metrics_tbl,",
    "      align = 'l',",
    "      format = 'latex',",
    "      booktabs = TRUE,",
    "      longtable = FALSE",
    "    ))",
    "    cat('\\\\endgroup\\n')",
    "  } else {",
    "    knitr::kable(params$dashboard$family_metrics_tbl, align = 'l')",
    "  }",
    "}",
    "```",
    "",
    .pagebreak_block()
  )

  nav_targets <- c(
    sprintf("[%s](#contents)", .tr("Back to Contents")),
    sprintf("[%s](#metadata)", .tr("Metadata")),
    sprintf("[%s](#dashboard)", .tr("Dashboard")),
    sprintf("[%s](#general-plot)", .tr("General Plot")),
    sprintf("[%s](#subplot-index)", .tr("Subplot Index")),
    sprintf("[%s](#checklist)", .tr("Checklist"))
  )

  gencol_section <- c(
    "",
    "## General Plot {#general-plot}",
    "",
    "```{r general-plot-html, echo=FALSE}",
    "if (knitr::is_html_output()) {",
    "  if (!is.null(params$interactive_main_plot)) {",
    "    params$interactive_main_plot",
    "  } else if (!is.null(params$main_plot)) {",
    "    params$main_plot",
    "  }",
    "}",
    "```",
    "",
    "```{r general-plot-pdf, echo=FALSE, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "if (knitr::is_latex_output() && !is.null(params$main_plot)) {",
    "  print(params$main_plot)",
    "}",
    "```",
    "",
    .tiny_nav_block(nav_targets),
    "",
    .pagebreak_block()
  )

  col_section <- c(
    "",
    "## Collected Only {#collected-only}",
    "",
    "```{r collected-only-html, echo=FALSE}",
    "if (knitr::is_html_output()) {",
    "  if (!is.null(params$interactive_collected_plot)) {",
    "    params$interactive_collected_plot",
    "  } else if (!is.null(params$collected_plot)) {",
    "    params$collected_plot",
    "  }",
    "}",
    "```",
    "",
    "```{r collected-only-pdf, echo=FALSE, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "if (knitr::is_latex_output() && !is.null(params$collected_plot)) {",
    "  print(params$collected_plot)",
    "}",
    "```",
    "",
    .tiny_nav_block(nav_targets),
    "",
    .pagebreak_block()
  )

  uncol_section <- c(
    "",
    "## Not Collected {#uncollected}",
    "",
    "```{r uncollected-html, echo=FALSE}",
    "if (knitr::is_html_output()) {",
    "  if (!is.null(params$interactive_uncollected_plot)) {",
    "    params$interactive_uncollected_plot",
    "  } else if (!is.null(params$uncollected_plot)) {",
    "    params$uncollected_plot",
    "  }",
    "}",
    "```",
    "",
    "```{r uncollected-pdf, echo=FALSE, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "if (knitr::is_latex_output() && !is.null(params$uncollected_plot)) {",
    "  print(params$uncollected_plot)",
    "}",
    "```",
    "",
    .tiny_nav_block(nav_targets),
    "",
    .pagebreak_block()
  )

  palm_section <- c(
    "",
    "## Not Collected Palms {#uncollected-palm}",
    "",
    "```{r uncollected-palm-html, echo=FALSE}",
    "if (knitr::is_html_output()) {",
    "  if (!is.null(params$interactive_palm_plot)) {",
    "    params$interactive_palm_plot",
    "  } else if (!is.null(params$uncollected_palm_plot)) {",
    "    params$uncollected_palm_plot",
    "  }",
    "}",
    "```",
    "",
    "```{r uncollected-palm-pdf, echo=FALSE, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "if (knitr::is_latex_output() && !is.null(params$uncollected_palm_plot)) {",
    "  print(params$uncollected_palm_plot)",
    "}",
    "```",
    "",
    .tiny_nav_block(nav_targets),
    "",
    .pagebreak_block()
  )


  index_section <- c(
    "",
    "## Subplot Index {#subplot-index}",
    "",
    "```{r toc, results='asis', echo=FALSE}",
    "cols <- 5",
    "n <- length(params$subplots_list)",
    "per_col <- ceiling(n / cols)",
    "",
    "# Get translations inside the chunk",
    "lang <- params$language",
    "subplots_label <- switch(lang,",
    "  'pt' = 'Subparcelas',",
    "  'es' = 'Subparcelas',",
    "  'fr' = 'Sous-parcelles',",
    "  'ma' = '子样地',",
    "  'pa' = 'Kukâra krepãã sââ',",
    "  'Subplots'",
    ")",
    "",
    "# Create header with proper Markdown table format",
    "header_cells <- rep(subplots_label, cols)",
    "header <- paste(header_cells, collapse = ' | ')",
    "",
    "# Separator line (must use only dashes and pipes, no translated text)",
    "separator <- paste(rep(':---', cols), collapse = ' | ')",
    "",
    "toc_lines <- c(header, separator)",
    "",
    "# Create subplot links with translated prefix",
    "subplot_prefix <- switch(lang,",
    "  'pt' = 'Subparcela ',",
    "  'es' = 'Subparcela ',",
    "  'fr' = 'Sous-parcelle ',",
    "  'ma' = '子样地 ',",
    "  'pa' = 'Kukâra krepãã sââ ',",
    "  'Subplot '",
    ")",
    "",
    "for (i in seq_len(per_col)) {",
    "  row <- character(cols)",
    "  for (j in 0:(cols - 1)) {",
    "    idx <- i + j * per_col",
    "    if (idx <= n) {",
    "      row[j + 1] <- paste0('[<small>', subplot_prefix, '<small>**', idx, '**](#subplot-', idx, ')')",
    "    } else {",
    "      row[j + 1] <- ' '",
    "    }",
    "  }",
    "  toc_lines <- c(toc_lines, paste(row, collapse = ' | '))",
    "}",
    "cat(paste(toc_lines, collapse = '\\n'))",
    "```",
    "",
    .tiny_nav_block(nav_targets),
    "",
    .pagebreak_block()
  )


  subplot_sections <- unlist(lapply(seq_along(subplot_plots), function(i) {
    anchor_id <- .chunk_id_factory("anchors")
    pagebreak_i <- if (i < length(subplot_plots)) .pagebreak_block("pagebreak-subplot") else character(0)

    c(
      sprintf("```{r %s, results='asis', echo=FALSE}", anchor_id),
      sprintf("tags_i <- params$subplots_list[[%d]]$data$`New Tag No`", i),
      "tags_i <- trimws(as.character(tags_i))",
      "tags_i <- tags_i[!is.na(tags_i) & nzchar(tags_i)]",
      sprintf("sp_number <- unique(params$subplots_list[[%d]]$data$T1)[1]", i),
      "if (knitr::is_latex_output()) {",
      "  anchors <- sprintf('\\\\hypertarget{tag-%s-S%s}{}', tags_i, sp_number)",
      "} else if (knitr::is_html_output()) {",
      "  anchors <- sprintf('<div id=\"tag-%s-S%s\"></div>', tags_i, sp_number)",
      "} else {",
      "  anchors <- character(0)",
      "}",
      "if (length(anchors)) cat(paste(anchors, collapse = '\n'), '\n')",
      "```",
      "",
      sprintf("### Subplot %d {#subplot-%d .unlisted .unnumbered}", i, i),
      "",
      sprintf("```{r subplot-%d, echo=FALSE, fig.width=12, fig.height=9}", i),
      sprintf("sp_data <- params$subplots_list[[%d]]$data", i),
      sprintf("if (knitr::is_html_output() && !is.null(params$subplots_list[[%d]]$data)) {", i),
      "  if (nrow(sp_data) == 0L) {",
      sprintf("    if (!is.null(params$subplots_list[[%d]]$plot)) print(params$subplots_list[[%d]]$plot)", i, i),
      "  } else {",
      "    x_col <- if ('x10' %in% names(sp_data)) 'x10' else 'X'",
      "    y_col <- if ('x10' %in% names(sp_data)) 'y10' else 'Y'",
      "    sz <- if ('x10' %in% names(sp_data)) 10 else params$subplot_size",
      "    status_vec <- if ('Status' %in% names(sp_data)) as.integer(sp_data$Status) else rep(NA_integer_, nrow(sp_data))",
      "    diam_vec <- if ('diameter' %in% names(sp_data)) suppressWarnings(as.numeric(sp_data$diameter)) else suppressWarnings(as.numeric(sp_data$D) / 10)",
      "    diam_vec[!is.finite(diam_vec)] <- 0",
      "    pt_colors <- dplyr::case_when(",
      "      status_vec == 1L ~ 'gray80',",
      "      status_vec == 2L ~ '#EF4444',",
      "      status_vec == 3L ~ 'gold',",
      "      TRUE ~ 'gray80'",
      "    )",
      "    pt_sizes <- pmax(1 + diam_vec * 1.5, 6)",
      "    tag_ids <- gsub('[^A-Za-z0-9-]', '', trimws(as.character(sp_data[['New Tag No']])))",
      "    hover <- paste0(",
      "      '<b>Tag:</b> ', sp_data[['New Tag No']], '<br>',",
      "      '<b>Species:</b> ', sp_data[['Original determination']], '<br>',",
      "      '<b>DBH:</b> ', round(suppressWarnings(as.numeric(sp_data$D)) / 10, 1), ' cm<br>'",
      "    )",
      "    sp_ticks <- seq(0, sz, by = 2.5)",
      "    sp_tick_labels <- paste0(sp_ticks, ' m')",
      "    p <- plotly::plot_ly(",
      "      x = sp_data[[x_col]],",
      "      y = sp_data[[y_col]],",
      "      type = 'scatter',",
      "      mode = 'markers+text',",
      "      marker = list(",
      "        color = pt_colors,",
      "        size = pt_sizes,",
      "        line = list(color = 'black', width = 0.8),",
      "        symbol = 'circle'",
      "      ),",
      "      text = as.character(sp_data[['New Tag No']]),",
      "      textposition = 'middle center',",
      "      textfont = list(size = 8, color = 'black'),",
      "      hovertext = hover,",
      "      hoverinfo = 'text',",
      "      customdata = tag_ids,",
      "      showlegend = FALSE",
      "    ) %>%",
      "      plotly::layout(",
      "        xaxis = list(title = list(text = 'X (m)', standoff = 5), range = c(-0.5, sz + 0.5), scaleanchor = 'y', scaleratio = 1, showgrid = TRUE, gridcolor = 'gray90', zeroline = FALSE, tickmode = 'array', tickvals = sp_ticks, ticktext = sp_tick_labels),",
      "        yaxis = list(title = list(text = 'Y (m)', standoff = 5), range = c(-0.5, sz + 0.5), showgrid = TRUE, gridcolor = 'gray90', zeroline = FALSE, tickmode = 'array', tickvals = sp_ticks, ticktext = sp_tick_labels),",
      "        shapes = list(list(type = 'rect', x0 = 0, x1 = sz, y0 = 0, y1 = sz, line = list(color = 'darkolivegreen', width = 2), fillcolor = 'rgba(0,0,0,0)', layer = 'below')),",
      "        plot_bgcolor = 'white',",
      "        paper_bgcolor = 'white',",
      "        margin = list(l = 5, r = 10, t = 60, b = 30)",
      "      ) %>%",
      "      plotly::config(displaylogo = FALSE, responsive = TRUE, displayModeBar = TRUE, modeBarButtonsToRemove = c('lasso2d', 'select2d'), toImageButtonOptions = list(format = 'png', scale = 2)) %>%",
      "      htmlwidgets::onRender(",
      "        'function(el) {",
      "          el.on(\"plotly_click\", function(d) {",
      "            if (!d.points || !d.points.length) return;",
      "            var t = d.points[0].customdata;",
      "            if (!t) return;",
      "            var a = document.getElementById(\"checklist-tag-\" + t);",
      "            if (a) a.scrollIntoView({behavior: \"smooth\", block: \"center\"});",
      "          });",
      "        }'",
      "      )",
      "    p",
      "  }",
      "} else {",
      sprintf("  if (!is.null(params$subplots_list[[%d]]$plot)) print(params$subplots_list[[%d]]$plot)", i, i),
      "}",
      "```",
      "",
      .tiny_nav_block(nav_targets, prefix = "navlinks-subplot"),
      "",
      pagebreak_i
    )
  }))


  checklist_section <- c(
    "",
    .pagebreak_block("pagebreak-checklist-intro"),
    "",
    "## Checklist {#checklist}",
    "",
    "```{r checklist, results='asis', echo=FALSE}",
    "library(dplyr)",
    "cat('\\n')",
    "for (fam in unique(spec_df$Family)) {",
    "  cat('\\n\\n### ', fam, '\\n\\n', sep = '')",
    "  fam_df <- spec_df %>% filter(Family == fam)",
    "  for (i in seq_len(nrow(fam_df))) {",
    "    sp <- fam_df$Species_fmt[i]",
    "    tag_vec <- fam_df$tag_vec[[i]]",
    "    tag_vec <- as.character(tag_vec)",
    "    tag_vec <- trimws(tag_vec)",
    "    tag_vec <- tag_vec[!is.na(tag_vec) & nzchar(tag_vec)]",
    "    if (length(tag_vec) == 0) next",
    "    ids <- safe_id(tag_vec)",
    "    t2s_df <- params$tag_to_subplot",
    "    tag_df <- data.frame(tag = tag_vec, id = ids, stringsAsFactors = FALSE)",
    "    is_monitora <- tolower(trimws(as.character(params$input_type))) == 'monitora'",
    "    tag_df <- merge(tag_df, t2s_df, by.x = 'tag', by.y = 'New Tag No', all.x = TRUE, sort = FALSE)",
    "    if (is_monitora) {",
    "      tag_df$target_idx <- ifelse(!is.na(tag_df$subplot_index), tag_df$subplot_index, safe_num(tag_df$T1))",
    "      has_fields <- !is.na(tag_df$subunit_letter) & !is.na(tag_df$T2)",
    "      tag_df$subplot_code <- ifelse(has_fields, paste0(tag_df$subunit_letter, tag_df$T2), paste0('S', tag_df$target_idx))",
    "      o <- order(tag_df$target_idx, safe_num(tag_df$tag))",
    "      tag_df <- tag_df[o, , drop = FALSE]",
    "      tag_df$anchor_id <- gsub('[^A-Za-z0-9-]', '', tag_df$tag)",
    "      if (knitr::is_html_output()) {",
    "        tag_links <- paste(sprintf('<span id=\"checklist-tag-%s\"></span>[%s \u2192 %s](#subplot-%s)', tag_df$anchor_id, tag_df$tag, tag_df$subplot_code, tag_df$target_idx), collapse = ' | ')",
    "      } else {",
    "        tag_links <- paste(sprintf('[ %s \u2192 %s ](#subplot-%s)', tag_df$tag, tag_df$subplot_code, tag_df$target_idx), collapse = ' | ')",
    "      }",
    "    } else {",
    "      tag_df$target_idx <- safe_num(tag_df$T1)",
    "      o <- order(tag_df$target_idx, safe_num(tag_df$tag))",
    "      tag_df <- tag_df[o, , drop = FALSE]",
    "      tag_df$anchor_id <- gsub('[^A-Za-z0-9-]', '', tag_df$tag)",
    "      if (knitr::is_html_output()) {",
    "        tag_links <- paste(sprintf('<span id=\"checklist-tag-%s\"></span>[%s - S%s](#subplot-%s)', tag_df$anchor_id, tag_df$tag, tag_df$target_idx, tag_df$target_idx), collapse = ' | ')",
    "      } else {",
    "        tag_links <- paste(sprintf('[ %s - S%s ](#subplot-%s)', tag_df$tag, tag_df$target_idx, tag_df$target_idx), collapse = ' | ')",
    "      }",
    "    }",
    "    cat('* ', sp, ' : ', tag_links, '\\n', sep = '')",
    "  }",
    "}",
    "```",
    "",
    .tiny_nav_block(nav_targets, prefix = "navlinks-checklist"),
    "",
    "```{r force-plotly-resize, echo=FALSE, results='asis'}",
    "if (knitr::is_html_output()) {",
    "  cat('",
    "<script>",
    "  setTimeout(function() {",
    "    var plots = document.querySelectorAll(\".js-plotly-plot, .plotly\");",
    "    plots.forEach(function(plot) {",
    "      if (plot._fullLayout && plot._fullLayout._size) {",
    "        Plotly.relayout(plot, {",
    "          \"xaxis.autorange\": true,",
    "          \"yaxis.autorange\": true",
    "        });",
    "      }",
    "    });",
    "  }, 1000);",
    "  window.dispatchEvent(new Event(\"resize\"));",
    "  setTimeout(function() {",
    "    window.dispatchEvent(new Event(\"resize\"));",
    "  }, 1500);",
    "</script>",
    "')",
    "}",
    "```",
    ""
  )


  rmd_content <- yaml_head

  .add_body <- function(...) {
    rmd_content <<- c(rmd_content, ...)
  }

  if (any(tf_col) && !any(tf_uncol) && !any(tf_palm)) {
    .add_body(mid_section, gencol_section, col_section, index_section)
  } else if (!any(tf_col) && any(tf_uncol) && !any(tf_palm)) {
    .add_body(mid_section, gencol_section, uncol_section, index_section)
  } else if (any(tf_col) && any(tf_uncol) && any(tf_palm)) {
    .add_body(mid_section, gencol_section, col_section, uncol_section, palm_section, index_section)
  } else if (any(tf_col) && any(tf_uncol) && !any(tf_palm)) {
    .add_body(mid_section, gencol_section, col_section, uncol_section, index_section)
  } else if (any(tf_col) && !any(tf_uncol) && any(tf_palm)) {
    .add_body(mid_section, gencol_section, col_section, palm_section, index_section)
  } else if (!any(tf_col) && any(tf_uncol) && any(tf_palm)) {
    .add_body(mid_section, gencol_section, uncol_section, palm_section, index_section)
  } else {
    .add_body(mid_section, gencol_section, index_section)
  }

  .add_body(
    "## Individual Subplots {#individual-subplots}",
    subplot_sections,
    checklist_section
  )

  rmd_content <- .translate_lines(rmd_content, dict, language)
  rmd_content
}

