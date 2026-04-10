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

  pick_first_existing <- function(df, candidates) {
    nm <- names(df)
    hit <- candidates[candidates %in% nm]
    if (length(hit)) hit[1] else NA_character_
  }

  get_col <- function(df, candidates, default = NA) {
    cl <- pick_first_existing(df, candidates)
    if (is.na(cl)) cl <- .pick_colname(df, candidates)
    if (!is.na(cl) && cl %in% names(df)) return(df[[cl]])
    rep(default, nrow(df))
  }

  parse_num_safe <- function(x) {
    z <- as.character(x)
    z[is.na(z)] <- ""
    z <- trimws(z)
    z <- gsub(",", ".", z, fixed = TRUE)
    suppressWarnings(as.numeric(z))
  }

  meta_value <- function(df, candidates) {
    cl <- pick_first_existing(df, candidates)
    if (is.na(cl)) cl <- .pick_colname(df, candidates)
    if (is.na(cl) || !(cl %in% names(df))) return("")

    x <- as.character(df[[cl]])
    x[is.na(x)] <- ""
    x <- trimws(x)
    x <- x[nzchar(x)]
    if (!length(x)) return("")
    unique(x)[1]
  }

  plot_meta <- list(
    plot_code = meta_value(in_dat, c("Plot Code", "PlotCode", "Plotcode")),
    plot_name = meta_value(in_dat, c("Plot Name", "PlotName")),
    team      = meta_value(in_dat, c("PI", "Team", "Collectors", "Collector"))
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

  plot_code_col <- pick_first_existing(in_dat, c("Plot Code", "PlotCode", "Plotcode"))
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

  tag_col <- pick_first_existing(in_dat, c(
    "New Tag No", "Tag No", "TagNo", "TagNumber",
    "Main Stem Tag", "Pv. Tag No", "Tree ID", "TreeID"
  ))
  if (is.na(tag_col)) {
    tag_col <- .pick_colname(in_dat, c(
      "New Tag No", "Tag No", "TagNo", "TagNumber",
      "Main Stem Tag", "Pv. Tag No", "Tree ID", "TreeID"
    ))
  }

  stem_group_col <- pick_first_existing(in_dat, c(
    "New Stem Grouping", "Stem Group ID", "StemGroupID"
  ))
  if (is.na(stem_group_col)) {
    stem_group_col <- .pick_colname(in_dat, c(
      "New Stem Grouping", "Stem Group ID", "StemGroupID"
    ))
  }
  t1_col <- pick_first_existing(in_dat, c("Sub Plot T1", "T1", "SubPlotT1"))
  t2_col <- pick_first_existing(in_dat, c("Sub Plot T2", "T2", "SubPlotT2"))
  x_col  <- pick_first_existing(in_dat, c("X"))
  y_col  <- pick_first_existing(in_dat, c("Y"))

  std_t1_col <- pick_first_existing(in_dat, c("Standardised SubPlot T1", "Standardized SubPlot T1"))
  std_x_col  <- pick_first_existing(in_dat, c("Standardised X", "Standardized X"))
  std_y_col  <- pick_first_existing(in_dat, c("Standardised Y", "Standardized Y"))

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
    t1 <- parse_num_safe(in_dat[[t1_col]])
    t2 <- if (!is.na(t2_col)) parse_num_safe(in_dat[[t2_col]]) else rep(NA_real_, nrow(in_dat))
    x  <- parse_num_safe(in_dat[[x_col]])
    y  <- parse_num_safe(in_dat[[y_col]])
    coord_mode <- "local"
  } else {
    t1 <- parse_num_safe(in_dat[[std_t1_col]])
    t2 <- if (!is.na(t2_col)) parse_num_safe(in_dat[[t2_col]]) else rep(NA_real_, nrow(in_dat))
    x  <- parse_num_safe(in_dat[[std_x_col]])
    y  <- parse_num_safe(in_dat[[std_y_col]])
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

  voucher_vals <- get_col(in_dat, c("Voucher", "Voucher Code"))
  collected_vals <- get_col(in_dat, c("Collected", "Voucher Collected"))

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
    Family = get_col(in_dat, c("Recommended Voucher Family", "Recommended Family", "Family")),
    `Original determination` = get_col(in_dat, c(
      "Recommended Voucher Species",
      "Recommended Species",
      "Original Identification",
      "Species"
    )),
    Morphospecies = get_col(in_dat, c("Morphospecies", "MorphoSpecies", "Morpho"), default = NA),
    D = get_col(in_dat, c("D", "D1", "DBH", "Dbh", "D0")),
    POM = get_col(in_dat, c("POM", "POM0", "DPOMtMinus1")),
    ExtraD = get_col(in_dat, c("Extra D", "ExtraD", "Extra D0")),
    ExtraPOM = get_col(in_dat, c("Extra POM", "ExtraPOM", "Extra POM0")),
    Flag1 = get_col(in_dat, c("Flag1", "F1")),
    Flag2 = get_col(in_dat, c("Flag2", "F2")),
    Flag3 = get_col(in_dat, c("Flag3", "F3")),
    LI = get_col(in_dat, c("LI")),
    CI = get_col(in_dat, c("CI")),
    CF = get_col(in_dat, c("CF")),
    CD1 = get_col(in_dat, c("CD1")),
    nrdups = get_col(in_dat, c("nrdups", "Stem Count")),
    Height = get_col(in_dat, c("Height", "Ht")),
    Voucher = voucher_vals,
    Silica = get_col(in_dat, c("Silica")),
    Collected = collected_final,
    `Census Notes` = get_col(in_dat, c("Census Notes", "Tree Notes", "Determination Comments", "Comments")),
    CAP = get_col(in_dat, c("CAP")),
    `Basal Area` = get_col(in_dat, c("Basal Area", "BA"))
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

    # 1) exact match first
    hit_idx <- match(ali_norm, cn_norm, nomatch = 0L)
    if (any(hit_idx > 0L)) {
      return(cn_raw[hit_idx[which(hit_idx > 0L)[1]]])
    }

    # 2) partial fallback only for safer aliases
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

  # exact-first, then normalized matcher
  col_subunidade <- if ("subunidade" %in% nm) "subunidade" else
    if ("Subunidades" %in% nm) "Subunidades" else
      .find_best_col(df, c(
        "subunidade", "subunidades", "sub_unidade", "sub_unidades",
        "orientacao", "orienta\u00e7\u00e3o", "unidade", "ul"
      ))

  col_nparcela <- if ("N_parcela" %in% nm) "N_parcela" else
    .find_best_col(df, c(
      "n_parcela", "nparcela", "numero_parcela", "n\u00famero_parcela",
      "parcela", "subplot", "subparcela", "t2"
    ))

  col_tag <- if ("N_arvore" %in% nm) "N_arvore" else
    .find_best_col(df, c(
      "n_arvore", "n \u00e1rvore", "numero_arvore", "n\u00famero_arvore",
      "narvore", "arvore", "\u00e1rvore", "tag", "newtag", "numarvore"
    ))

  col_coletores <- if ("nome_coletores" %in% nm) "nome_coletores" else
    .find_best_col(df, c("nome_coletores", "coletores", "equipe", "team"))

  col_uc <- if ("NOMEUC" %in% nm) "NOMEUC" else
    if ("CDUC" %in% nm) "CDUC" else
      .find_best_col(df, c("nomeuc", "nome_uc", "uc", "unidadeconservacao", "unidade_conservacao", "cduc"))

  col_estacao <- if ("Nome_estacao" %in% nm) "Nome_estacao" else
    .find_best_col(df, c("nome_estacao", "nome esta\u00e7\u00e3o", "nome_esta\u00e7\u00e3o", "estacao", "esta\u00e7\u00e3o", "nomeestacao"))

  col_estacao_n <- if ("N_estacao" %in% nm) "N_estacao" else
    .find_best_col(df, c("n_estacao", "nestacao", "num_estacao", "numero_estacao", "n\u00famero_estacao", "n\u00baestacao", "n\u00b0estacao"))

  col_familia <- if ("Fam\u00edlia" %in% nm) "Fam\u00edlia" else
    if ("Familia" %in% nm) "Familia" else
      .find_best_col(df, c("familia", "fam\u00edlia", "family"))

  col_genero <- if ("G\u00eanero" %in% nm) "G\u00eanero" else
    if ("Genero" %in% nm) "Genero" else
      .find_best_col(df, c("genero", "g\u00eanero", "genus"))

  col_especie <- if ("Esp\u00e9cie" %in% nm) "Esp\u00e9cie" else
    if ("Especie" %in% nm) "Especie" else
      .find_best_col(df, c("especie", "esp\u00e9cie", "species", "sp"))

  col_nomecomum <- if ("Nome comum" %in% nm) "Nome comum" else
    .find_best_col(df, c("nome comum", "nome_comum", "nomecomum", "popular", "morphospecies"))

  col_coletado <- if ("individuo coletado" %in% nm) "individuo coletado" else
    if ("individuo coletado ?" %in% nm) "individuo coletado ?" else
      if ("indiv\u00edduo coletado" %in% nm) "indiv\u00edduo coletado" else
        if ("indiv\u00edduo coletado ?" %in% nm) "indiv\u00edduo coletado ?" else
          .find_best_col(df, c(
            "individuo coletado", "individuo coletado ?",
            "indiv\u00edduo coletado", "indiv\u00edduo coletado ?",
            "individuo_coletado", "indiv\u00edduo_coletado",
            "coletado", "collected"
          ))

  col_voucher_c <- .find_best_col(df, c(
    "voucher/coletor", "voucher_coletor", "vouchercoletor",
    "coletor", "collector"
  ))

  col_voucher_n <- if ("voucher/n\u00famero" %in% nm) "voucher/n\u00famero" else
    if ("voucher/numero" %in% nm) "voucher/numero" else
      .find_best_col(df, c("voucher/n\u00famero", "voucher/numero", "voucher_numero", "voucher_n\u00famero", "vouchernumero", "voucher_n", "voucher"))

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
      .find_best_col(df, c("cap_tot", "captot", "cap total", "cap_total", "cap", "circ total", "circunferencia", "circunfer\u00eancia"))

  col_ano <- if ("Ano" %in% nm) "Ano" else
    if ("Data" %in% nm) "Data" else
      .find_best_col(df, c("ano", "censo", "year", "data", "date"))

  col_dead <- if ("arvore_morta" %in% nm) "arvore_morta" else
    if ("\u00e1rvore_morta" %in% nm) "\u00e1rvore_morta" else
      .find_best_col(df, c("arvore_morta", "\u00e1rvore_morta", "arvore morta", "morta", "dead"))

  col_obs <- if ("observa\u00e7\u00e3o" %in% nm) "observa\u00e7\u00e3o" else
    if ("observacao" %in% nm) "observacao" else
      .find_best_col(df, c("observa\u00e7\u00e3o", "observacao", "obs", "census notes", "comentarios", "coment\u00e1rios"))

  col_basal_area <- if ("AB" %in% nm) "AB" else
    if ("ABcap" %in% nm) "ABcap" else
      .find_best_col(df, c("ab", "abcap", "basal area", "area basal", "\u00e1rea basal"))

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

  vc <- if (!is.na(col_voucher_c) && col_voucher_c %in% names(df_use)) clean_spaces(df_use[[col_voucher_c]]) else rep("", nrow(df_use))
  vnr <- if (!is.na(col_voucher_n) && col_voucher_n %in% names(df_use)) clean_spaces(df_use[[col_voucher_n]]) else rep("", nrow(df_use))
  voucher_vec <- ifelse((!nzchar(vc)) & (!nzchar(vnr)), NA_character_, trimws(paste(vc, vnr)))

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

  voucher_collector_raw <- if (!is.na(col_voucher_c) && col_voucher_c %in% names(df_use)) {
    clean_spaces(df_use[[col_voucher_c]])
  } else {
    rep("", nrow(df_use))
  }

  collected_norm <- tolower(collected_raw)

  is_collected <- collected_norm %in% c(
    "sim", "s", "yes", "y", "true", "1", "coletado", "coletada"
  ) | nzchar(voucher_collector_raw)

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
      global_x < max_x,
      global_y < max_y
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
      metric = "Metrica",
      value = "Valor",
      total_individuals = "Total de individuos",
      collected = "Coletados",
      uncollected = "Nao coletados",
      palms = "Palmeiras (Arecaceae)",
      families = "Familias distintas",
      species = "Especies distintas",
      genera = "Generos distintos",
      shannon = "Indice de Shannon",
      simpson = "Indice de Simpson",
      family_plot = "Familias mais comuns",
      species_plot = "Especies mais abundantes",
      subplot_plot = "Percentual de coleta por subparcela",
      dbh_plot = "Classes de DAP (cm)",
      dbh_x = "Classe de DAP (cm)",
      dbh_y = "Numero de individuos",
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
      subplot_plot = "Percentual de Coleta por Subparcela",
      dbh_plot = "Classes de DAP (cm)",
      dbh_x = "Classe de DAP (cm)",
      dbh_y = "N\u00famero de indiv\u00edduos",
      x_ind = "Individuos",
      x_pct = "%",
      subplot = "Subparcela",
      species_tbl = c(
        "P\u0129rak\u00e2ri", "Kyapi\u00e2hapi\u00e2ra", "Abund\u00e2ncia", "Subparcelas",
        "Dens. rel. (%)", "Freq. rel. (%)", "VI"
      ),
      family_tbl = c(
        "Kyapi\u00e2hapi\u00e2ra", "Abund\u00e2ncia", "Riqueza", "Subparcelas",
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

  list(
    metrics_tbl = metrics_tbl,
    species_metrics_tbl = species_metrics_tbl,
    family_metrics_tbl = family_metrics_tbl,
    family_plot = family_plot,
    species_plot = species_plot,
    subplot_plot = subplot_plot,
    dbh_plot = dbh_plot,
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
      "Team",
      "Number of Census",
      "Dates of Census",
      "Specimen Counts",
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
      "Equipe = Team",
      "N\u00famero de Censos = Number of Census",
      "Datas dos Censos = Dates of Census",
      "Contagem de Indiv\u00edduos = Specimen Counts",
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
      "Subparcela   = Subplot  ",
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
      "Equipo = Team",
      "N\u00famero de Censos = Number of Census",
      "Fechas de los Censos = Dates of Census",
      "Conteo de Individuos = Specimen Counts",
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
      "\u00c9quipe = Team",
      "Nombre de Recensements = Number of Census",
      "Dates des Recensements = Dates of Census",
      "Comptage des Individus = Specimen Counts",
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
      "\u56e2\u961f = Team",
      "\u666e\u67e5\u6b21\u6570 = Number of Census",
      "\u666e\u67e5\u65e5\u671f = Dates of Census",
      "\u4e2a\u4f53\u7edf\u8ba1 = Specimen Counts",
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
      "Кypa kuk\u00e2ri = Plot Code",
      "S\u00e2p\u00ear\u00e4nt\u00ea = Team",
      "Junti h\u1ebd r\u00f5 s\u00ean p\u00e2rikran = Number of Census",
      "Pj\u00e3n haka h\u00e3 r\u00ea war\u0129 p\u00e2ri = Dates of Census",
      "Junti h\u1ebd si m\u1ebdra = Specimen Counts",
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
      "Jyy ti he P\u00e2ri m\u1ebdra karêrãkjan kukâri s\u00e2\u00e2  = Collection Percentage by Subplot",
      "Classes de DAP = DBH Classes",
      "Classe de DAP = DBH Class",
      "N\u00famero de indiv\u00edduos = Number of individuals",
      "P\u0129rak\u00e2ri p\u00e2ri m\u1ebdra = Species Metrics",
      "P\u0129rak\u00e2ri p\u00e2ri kyapi\u00e2hapi\u00e2ra = Family Metrics",
      "P\u0129rak\u00e2ri = Species",
      "P\u0129rak\u00e2ri kyapi\u00e2hapi\u00e2ra = Family",
      "Abund\u00e2ncia = Abundance",
      "Kuk\u00e2ra krep\u00e3\u00e3 S\u00e2\u00e2 = Subplots",
      "Dens. rel. (%) = Rel. density (%)",
      "Freq. rel. (%) = Rel. frequency (%)",
      "VI = IV",
      "sop\u00e2ri m\u1ebdra = Richness",
      "P\u00e3p\u00e3 kypapr\u1ebdpi kuk\u00e2ri = General Plot",
      "P\u00e2ri sonswa kypa kuk\u00e2ri kran = Collected Only",
      "R\u00f5\u00f5rin kwatis\u00f4m\u00eara kuk\u00e2ri kran = Not Collected Palms",
      "P\u00e2ri r\u00f5r\u0129 kypa kuk\u00e2ri kran = Not Collected",
      "Kwatis\u00f4m\u1ebdra = Palms",
      "T\u00e4 sokkjaraa r\u00ea t\u00e2 kuk\u00e2ra kran = Subplot Index",
      "Kuk\u00e2ra krep\u00e3\u00e3 S\u00e2\u00e2 = Subplots",
      "Kuk\u00e2ra krep\u00e3\u00e3 S\u00e2\u00e2   = Subplot ",
      "Issi pyr\u00e3h\u00e3 p\u00e2rijnsim\u1ebdra = Checklist",
      "Kuk\u00e2ra krep\u00e3\u00e3 S\u00e2\u00e2 pyti = Individual Subplots"
    )
  )

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
    "  collected_plot: NULL",
    "  uncollected_plot: NULL",
    "  uncollected_palm_plot: NULL",
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
    "  cat(sprintf('<p style=\"margin-top:0;\">%s | %s</p>\\n', params$metadata$plot_name, params$metadata$plot_code))",
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
    "is_pa <- identical(lang, 'pa')",
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
    "**Plot Name:** `r params$metadata$plot_name`",
    "",
    "**Plot Code:** `r params$metadata$plot_code`",
    "",
    "**Team:** `r params$metadata$team`",
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
    "if (!is.null(params$dashboard$family_plot)) print(params$dashboard$family_plot)",
    "```",
    "",
    "### Most Abundant Species",
    "",
    "```{r dashboard-species-plot, fig.width=10, fig.height=6, fig.align='center', out.width='95%'}",
    "if (!is.null(params$dashboard$species_plot)) print(params$dashboard$species_plot)",
    "```",
    "",
    "### Collection Percentage by Subplot",
    "",
    "```{r dashboard-subplot-plot, fig.width=10, fig.height=6, fig.align='center', out.width='95%'}",
    "if (!is.null(params$dashboard$subplot_plot)) print(params$dashboard$subplot_plot)",
    "```",
    "",
    "### DBH Classes",
    "",
    "```{r dashboard-dbh-plot, fig.width=10, fig.height=6, fig.align='center', out.width='95%'}",
    "if (!is.null(params$dashboard$dbh_plot)) print(params$dashboard$dbh_plot)",
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
    "[Back to Contents](#contents)",
    "[Metadata](#metadata)",
    "[Specimen Counts](#counts)",
    "[Dashboard](#dashboard)",
    "[General Plot](#general-plot)",
    "[Subplot Index](#subplot-index)",
    "[Checklist](#checklist)"
  )

  gencol_section <- c(
    "",
    "## General Plot {#general-plot}",
    "```{r general-plot, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$main_plot)",
    "```",
    "",
    .tiny_nav_block(nav_targets),
    "",
    .pagebreak_block()
  )

  col_section <- c(
    "## Collected Only {#collected-only}",
    "```{r collected-only, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$collected_plot)",
    "```",
    "",
    .tiny_nav_block(nav_targets),
    "",
    .pagebreak_block()
  )

  uncol_section <- c(
    "## Not Collected {#uncollected}",
    "```{r uncollected, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$uncollected_plot)",
    "```",
    "",
    .tiny_nav_block(nav_targets),
    "",
    .pagebreak_block()
  )

  palm_section <- c(
    "## Not Collected Palms {#uncollected-palm}",
    "```{r uncollected-palm, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$uncollected_palm_plot)",
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
    "header <- paste(rep('Subplots', cols), collapse = ' | ')",
    "separator <- paste(rep('---', cols), collapse = ' | ')",
    "toc_lines <- c(header, separator)",
    "for (i in seq_len(per_col)) {",
    "  row <- character(cols)",
    "  for (j in 0:(cols - 1)) {",
    "    idx <- i + j * per_col",
    "    if (idx <= n) {",
    "      row[j + 1] <- paste0('[Subplot ', idx, '](#subplot-', idx, ')')",
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
      sprintf("```{r subplot-%d, fig.width=12, fig.height=9}", i),
      sprintf("print(params$subplots_list[[%d]]$plot)", i),
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
    "      tag_links <- paste(sprintf('[ %s \u2192 %s ](#subplot-%s)', tag_df$tag, tag_df$subplot_code, tag_df$target_idx), collapse = ' | ')",
    "    } else {",
    "      tag_df$target_idx <- safe_num(tag_df$T1)",
    "      o <- order(tag_df$target_idx, safe_num(tag_df$tag))",
    "      tag_df <- tag_df[o, , drop = FALSE]",
    "      tag_links <- paste(sprintf('[ %s - S%s ](#subplot-%s)', tag_df$tag, tag_df$target_idx, tag_df$target_idx), collapse = ' | ')",
    "    }",
    "    cat('* ', sp, ' : ', tag_links, '\\n', sep = '')",
    "  }",
    "}",
    "```",
    "",
    .tiny_nav_block(nav_targets, prefix = "navlinks-checklist"),
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

# =============================================================================
# Herbarium lookup helpers
# =============================================================================
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

#' Resolve cache paths for herbarium lookup
#'
#' @return Named list with cache directory and DuckDB path.
#'
#' @keywords internal
#' @noRd
#'
.herbaria_cache_paths <- function() {
  cache_dir <- getOption("forplotR.herbaria_cache_dir")

  if (!(is.character(cache_dir) && length(cache_dir) == 1L && nzchar(cache_dir))) {
    cache_dir <- tryCatch(
      tools::R_user_dir("forplotR", which = "cache"),
      error = function(e) NULL
    )
  }

  if (is.null(cache_dir) || !nzchar(cache_dir)) {
    cache_dir <- file.path(getwd(), ".forplotR_cache")
  }

  if (!dir.exists(cache_dir)) {
    dir.create(cache_dir, recursive = TRUE, showWarnings = FALSE)
  }

  list(
    cache_dir = cache_dir,
    duckdb_path = file.path(cache_dir, "herbaria_index.duckdb")
  )
}

#' Extract a collector number token from free text
#'
#' @param x Character vector.
#'
#' @return Character vector with extracted number token.
#'
#' @keywords internal
#' @noRd
.extract_number_token <- function(x) {
  if (is.na(x) || !nzchar(trimws(x))) {
    return(NA_character_)
  }

  x <- as.character(x)

  m <- regexec("(\\d+)\\s*$", x, perl = TRUE)
  if (m[[1]][1] > 0L) {
    hit <- regmatches(x, m)[[1]]
    if (length(hit) > 1L && nzchar(hit[2])) {
      return(trimws(hit[2]))
    }
  }

  nums <- regmatches(x, gregexpr("\\d+", x, perl = TRUE))[[1]]
  if (length(nums)) {
    return(nums[which.max(nchar(nums))])
  }

  NA_character_
}

#' Keep only digits in collector numbers
#'
#' @param x Character vector.
#'
#' @return Character vector with only digits.
#'
#' @keywords internal
#' @noRd
.normalize_number_token <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- ""
  x <- gsub("[^0-9]", "", x)
  x[!nzchar(x)] <- NA_character_
  x
}

#' Split recordedBy into candidate people
#'
#' @param recordedBy Character scalar.
#'
#' @return Character vector of candidate person strings.
#'
#' @keywords internal
#' @noRd
.split_recordedby_people <- function(recordedBy) {
  if (is.na(recordedBy) || !nzchar(trimws(recordedBy))) {
    return(character(0))
  }

  x <- as.character(recordedBy)
  x <- iconv(x, to = "ASCII//TRANSLIT", sub = "")
  x[is.na(x)] <- ""
  x <- trimws(x)

  if (!nzchar(x)) {
    return(character(0))
  }

  x <- gsub("\\bet\\s+al\\b\\.?", "", x, ignore.case = TRUE)
  x <- gsub("\\s+", " ", x)

  parts <- unlist(strsplit(
    x,
    "\\s*;\\s*|\\s*\\|\\s*|\\s*&\\s*|\\s+and\\s+",
    perl = TRUE
  ))
  parts <- trimws(parts)
  parts <- parts[nzchar(parts)]

  out <- character(0)

  for (p in parts) {
    comma_n <- lengths(regmatches(p, gregexpr(",", p, fixed = TRUE)))

    if (comma_n <= 1L && grepl("^[^,]+,[^,]+$", p)) {
      out <- c(out, p)
    } else if (comma_n >= 1L) {
      pp <- unlist(strsplit(p, "\\s*,\\s*", perl = TRUE))
      pp <- trimws(pp)
      pp <- pp[nzchar(pp)]
      out <- c(out, pp)
    } else {
      out <- c(out, p)
    }
  }

  out <- trimws(out)
  unique(out[nzchar(out)])
}

#' Build collector tokens from one person name
#'
#' @param name_raw Character scalar.
#'
#' @return Named character vector with primary and fallback tokens.
#'
#' @keywords internal
#' @noRd
.collector_tokens_one <- function(name_raw) {
  if (is.na(name_raw) || !nzchar(trimws(name_raw))) {
    return(c(primary = NA_character_, fallback = NA_character_))
  }

  x <- as.character(name_raw)
  x <- iconv(x, to = "ASCII//TRANSLIT", sub = "")
  x[is.na(x)] <- ""
  x <- trimws(tolower(x))
  x <- gsub("\\bet\\s+al\\b\\.?", "", x)
  x <- gsub("\\s+", " ", x)

  if (!nzchar(x)) {
    return(c(primary = NA_character_, fallback = NA_character_))
  }

  particles <- c("de", "da", "do", "das", "dos", "del", "della", "van", "von")

  words_clean <- function(z) {
    z <- gsub("[^a-z ]", " ", z)
    z <- unlist(strsplit(z, "\\s+"))
    z <- z[nzchar(z)]
    z <- z[!(z %in% particles)]
    z
  }

  surname <- ""
  initials <- ""

  if (grepl(",", x, fixed = TRUE)) {
    parts <- strsplit(x, ",", fixed = TRUE)[[1]]
    left <- words_clean(parts[1])
    right <- if (length(parts) >= 2L) words_clean(parts[2]) else character(0)

    if (length(left)) {
      surname <- left[length(left)]
    }
    if (length(right)) {
      initials <- paste(substr(right, 1L, 1L), collapse = "")
    }
  } else {
    w <- words_clean(x)
    if (length(w)) {
      surname <- w[length(w)]
      if (length(w) > 1L) {
        initials <- paste(substr(w[-length(w)], 1L, 1L), collapse = "")
      }
    }
  }

  surname <- gsub("[^a-z0-9]", "", surname)
  initials <- gsub("[^a-z0-9]", "", initials)

  fallback <- if (nzchar(surname)) surname else NA_character_
  primary <- if (nzchar(surname)) paste0(initials, surname) else NA_character_

  if (!nzchar(primary)) {
    primary <- fallback
  }

  c(primary = primary, fallback = fallback)
}

#' Expand compact collector code prefixes
#'
#' @param prefix Character scalar prefix extracted from a compact voucher.
#' @param collector_codes Optional named character vector mapping prefixes to collector names.
#'
#' @return Character scalar or NA.
#'
#' @keywords internal
#' @noRd
.expand_collector_code <- function(prefix, collector_codes = NULL) {
  if (is.null(collector_codes) || !length(collector_codes)) {
    return(NA_character_)
  }

  if (is.null(names(collector_codes)) || !length(names(collector_codes))) {
    stop(
      "`collector_codes` must be a named character vector. ",
      "Example: c('DC' = 'D. Cardoso', 'PWM' = 'P. W. Moonlight').",
      call. = FALSE
    )
  }

  nm <- toupper(trimws(names(collector_codes)))
  val <- as.character(collector_codes)
  key <- toupper(trimws(as.character(prefix)))

  hit <- match(key, nm, nomatch = 0L)
  if (hit > 0L) {
    out <- trimws(val[hit])
    if (nzchar(out)) {
      return(out)
    }
  }

  NA_character_
}

#' Parse one voucher string into collector and number
#'
#' @param voucher Character scalar.
#' @param collector_fallback Optional fallback collector name.
#' @param collector_codes Optional named character vector mapping compact prefixes to collector names.
#'
#' @return Named character vector with collector and number.
#'
#' @keywords internal
#' @noRd
.parse_voucher_one <- function(voucher,
                               collector_fallback = NULL,
                               collector_codes = NULL) {
  v <- as.character(voucher)
  v[is.na(v)] <- ""
  v <- trimws(v)

  if (!nzchar(v)) {
    return(c(collector = NA_character_, number = NA_character_))
  }

  if (grepl("^\\d+$", v)) {
    coll <- if (!is.null(collector_fallback) && nzchar(trimws(collector_fallback))) {
      trimws(collector_fallback)
    } else {
      NA_character_
    }
    return(c(collector = coll, number = v))
  }

  m_compact <- regexec("^([A-Za-z]+)\\s*([0-9]+)$", v, perl = TRUE)
  if (m_compact[[1]][1] > 0L) {
    hit <- regmatches(v, m_compact)[[1]]
    coll <- .expand_collector_code(hit[2], collector_codes = collector_codes)

    if (is.na(coll) && !is.null(collector_fallback) && nzchar(trimws(collector_fallback))) {
      coll <- trimws(collector_fallback)
    }

    return(c(collector = coll, number = hit[3]))
  }

  v2 <- gsub("[/;|]+", " ", v)
  v2 <- gsub("\\bet\\s+al\\b\\.?", "", v2, ignore.case = TRUE)
  v2 <- gsub("\\s+", " ", trimws(v2))

  m <- regexec("^(.*?)[[:space:]]*(\\d+)\\s*$", v2, perl = TRUE)
  if (m[[1]][1] > 0L) {
    hit <- regmatches(v2, m)[[1]]
    coll <- trimws(hit[2])
    num <- trimws(hit[3])

    if (!nzchar(coll) && !is.null(collector_fallback) && nzchar(trimws(collector_fallback))) {
      coll <- trimws(collector_fallback)
    }

    return(c(
      collector = if (nzchar(coll)) coll else NA_character_,
      number = if (nzchar(num)) num else NA_character_
    ))
  }

  c(
    collector = if (!is.null(collector_fallback) && nzchar(trimws(collector_fallback))) {
      trimws(collector_fallback)
    } else {
      NA_character_
    },
    number = .extract_number_token(v2)
  )
}

#' Parse census vouchers into normalized match keys
#'
#' @param voucher Character vector of voucher values.
#' @param collector_fallback Optional collector fallback.
#' @param collector_codes Optional named character vector mapping compact prefixes to collector names.
#'
#' @return Data frame with raw and normalized voucher identity fields.
#'
#' @keywords internal
#' @noRd
.parse_census_identity <- function(voucher,
                                   collector_fallback = NULL,
                                   collector_codes = NULL) {
  voucher <- as.character(voucher)
  voucher[is.na(voucher)] <- ""

  parsed <- lapply(
    voucher,
    .parse_voucher_one,
    collector_fallback = collector_fallback,
    collector_codes = collector_codes
  )

  collector_raw <- vapply(parsed, `[[`, character(1), "collector")
  number_raw <- vapply(parsed, `[[`, character(1), "number")
  number_clean <- .normalize_number_token(number_raw)

  tok <- t(vapply(
    collector_raw,
    .collector_tokens_one,
    FUN.VALUE = c(primary = NA_character_, fallback = NA_character_)
  ))

  primary_key <- ifelse(
    !is.na(tok[, "primary"]) & nzchar(tok[, "primary"]) &
      !is.na(number_clean) & nzchar(number_clean),
    paste0(tok[, "primary"], "||", number_clean),
    NA_character_
  )

  fallback_key <- ifelse(
    !is.na(tok[, "fallback"]) & nzchar(tok[, "fallback"]) &
      !is.na(number_clean) & nzchar(number_clean),
    paste0(tok[, "fallback"], "||", number_clean),
    NA_character_
  )

  data.frame(
    collector_raw = collector_raw,
    number_raw = number_raw,
    number_clean = number_clean,
    primary_key = primary_key,
    fallback_key = fallback_key,
    stringsAsFactors = FALSE
  )
}

#' Convert one IPT resource id into a guessed herbarium code
#'
#' @param resource_id Character scalar.
#'
#' @return Character scalar.
#'
#' @keywords internal
#' @noRd
#'
.resource_guess_herbarium <- function(resource_id) {
  rid <- tolower(trimws(as.character(resource_id)))
  rid <- sub("^.*[?&]r=", "", rid)
  rid <- sub("&.*$", "", rid)

  out <- sub("^jbrj_", "", rid)
  out <- sub("_herbarium$", "", out)
  out <- sub("_reflora$", "", out)
  out <- sub("_jabot$", "", out)
  out <- sub("_hv$", "", out)
  out <- sub("_.*$", "", out)
  out <- toupper(out)

  if (grepl("(^|_)rb($|_)", rid)) {
    out <- "RB"
  }

  out
}

#' Read latest version info from an IPT resource page
#'
#' @param resource_url Character scalar.
#'
#' @return Named list with version, published_on and records.
#'
#' @keywords internal
#' @noRd
.get_latest_version_info <- function(resource_url) {
  x <- tryCatch(
    suppressWarnings(readLines(resource_url, encoding = "UTF-8", warn = FALSE)),
    error = function(e) character(0)
  )

  out <- list(
    version = NA_character_,
    published_on = NA_character_,
    records = NA_character_
  )

  if (!length(x)) {
    return(out)
  }

  txt <- paste(x, collapse = " ")

  ver <- sub(
    ".*latestVersion[^0-9]*([0-9]+(?:\\.[0-9]+)+).*",
    "\\1",
    txt,
    perl = TRUE
  )
  if (!identical(ver, txt)) {
    out$version <- ver
  }

  dates <- regmatches(txt, gregexpr("\\d{4}-\\d{2}-\\d{2}", txt, perl = TRUE))[[1]]
  if (length(dates)) {
    out$published_on <- dates[1]
  }

  recs <- regmatches(txt, gregexpr("\\b[0-9][0-9,]*\\b", txt, perl = TRUE))[[1]]
  if (length(recs)) {
    out$records <- recs[1]
  }

  out
}

#' Discover IPT resources for one or more herbaria
#'
#' @param herbarium Character vector of herbarium codes.
#' @param ipt One of "jabot" or "reflora".
#' @param resource_map Optional named character vector mapping "HERBARIUM::source" to resource_id.
#'
#' @return Data frame with discovered resources.
#'
#' @keywords internal
#' @noRd
.get_ipt_info <- function(herbarium,
                          ipt = c("jabot", "reflora"),
                          resource_map = NULL) {
  ipt <- match.arg(ipt)
  .arg_check_herbarium(herbarium)

  herbarium <- unique(toupper(trimws(as.character(herbarium))))

  empty <- data.frame(
    ipt = character(0),
    herbarium = character(0),
    resource_id = character(0),
    archive_base = character(0),
    resource_url = character(0),
    stringsAsFactors = FALSE
  )

  dcat_roots <- if (ipt == "jabot") {
    c("jabot", "jbrj")
  } else {
    "reflora"
  }

  out_list <- vector("list", length(dcat_roots))
  pos <- 0L

  for (root in dcat_roots) {
    dcat_url <- paste0("https://ipt.jbrj.gov.br/", root, "/dcat")

    lines <- tryCatch(
      suppressWarnings(readLines(dcat_url, encoding = "UTF-8", warn = FALSE)),
      error = function(e) character(0)
    )

    if (!length(lines)) {
      next
    }

    txt <- paste(lines, collapse = "\n")

    dl_pat <- paste0(
      "https://ipt\\.jbrj\\.gov\\.br/", root,
      "/archive\\.do\\?r=[A-Za-z0-9_\\-]+(?:&v=[0-9.]+)?"
    )

    download_urls <- unique(unlist(regmatches(txt, gregexpr(dl_pat, txt, perl = TRUE))))

    if (!length(download_urls)) {
      next
    }

    resource_ids <- sub("^.*[?&]r=", "", download_urls)
    resource_ids <- sub("&.*$", "", resource_ids)

    herb_guess <- vapply(resource_ids, .resource_guess_herbarium, character(1))

    one <- data.frame(
      ipt = rep(ipt, length(resource_ids)),
      herbarium = herb_guess,
      resource_id = resource_ids,
      archive_base = paste0("https://ipt.jbrj.gov.br/", root, "/archive.do?r="),
      resource_url = paste0("https://ipt.jbrj.gov.br/", root, "/resource?r=", resource_ids),
      stringsAsFactors = FALSE
    )

    one <- one[one$herbarium %in% herbarium, , drop = FALSE]

    if (nrow(one)) {
      pos <- pos + 1L
      out_list[[pos]] <- one
    }
  }

  out <- if (pos) {
    do.call(rbind, out_list[seq_len(pos)])
  } else {
    empty
  }

  if (!is.null(resource_map) && length(resource_map)) {
    if (is.null(names(resource_map)) || !length(names(resource_map))) {
      stop(
        "`herbaria_resource_map` must be a named character vector. ",
        "Example: c('RB::jabot' = 'jbrj_rb').",
        call. = FALSE
      )
    }

    wanted_names <- paste0(herbarium, "::", ipt)
    hit <- names(resource_map) %in% wanted_names

    if (any(hit)) {
      rid <- as.character(resource_map[hit])
      root <- ifelse(grepl("^jbrj_", rid), "jbrj", ipt)

      override_df <- data.frame(
        ipt = ipt,
        herbarium = sub("::.*$", "", names(resource_map)[hit]),
        resource_id = rid,
        archive_base = paste0("https://ipt.jbrj.gov.br/", root, "/archive.do?r="),
        resource_url = paste0("https://ipt.jbrj.gov.br/", root, "/resource?r=", rid),
        stringsAsFactors = FALSE
      )

      out <- rbind(out, override_df)
    }
  }

  if (!nrow(out)) {
    return(empty)
  }

  out <- out[!duplicated(paste(out$archive_base, out$resource_id, sep = "::")), , drop = FALSE]

  if (ipt == "jabot") {
    prio <- ifelse(grepl("/jbrj/", out$archive_base, fixed = FALSE), 0L, 1L)
    out <- out[order(out$herbarium, prio, out$resource_id), , drop = FALSE]
    out <- out[!duplicated(out$herbarium), , drop = FALSE]
  }

  rownames(out) <- NULL
  out
}

#' Download one DwC-A archive and extract occurrence.txt
#'
#' @param info_row One-row data frame returned by .get_ipt_info().
#' @param dir Output directory.
#' @param verbose Logical.
#' @param force_refresh Logical.
#'
#' @return Character path to extracted archive directory, or NULL.
#'
#' @keywords internal
#' @noRd
.download_dwca_one <- function(info_row,
                               dir,
                               verbose = FALSE,
                               force_refresh = FALSE) {
  stopifnot(is.data.frame(info_row), nrow(info_row) == 1L)

  if (!dir.exists(dir)) {
    dir.create(dir, recursive = TRUE, showWarnings = FALSE)
  }

  safe_id <- gsub("[^A-Za-z0-9]+", "_", info_row$resource_id)
  out_dir <- file.path(
    dir,
    paste0("dwca_", info_row$ipt, "_", info_row$herbarium, "_", safe_id)
  )

  if (isTRUE(force_refresh) && dir.exists(out_dir)) {
    unlink(out_dir, recursive = TRUE, force = TRUE)
  }

  if (dir.exists(out_dir) && file.exists(file.path(out_dir, "occurrence.txt"))) {
    return(out_dir)
  }

  if (!dir.exists(out_dir)) {
    dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)
  }

  vinfo <- .get_latest_version_info(info_row$resource_url)
  archive_url <- paste0(info_row$archive_base, info_row$resource_id)

  if (!is.na(vinfo$version) && nzchar(vinfo$version)) {
    archive_url <- paste0(
      archive_url,
      "&v=",
      utils::URLencode(vinfo$version, reserved = TRUE)
    )
  }

  zipfile <- tempfile(pattern = "dwca_", tmpdir = dir, fileext = ".zip")
  old_timeout <- getOption("timeout")
  options(timeout = max(600, old_timeout))
  on.exit(options(timeout = old_timeout), add = TRUE)

  if (isTRUE(verbose)) {
    message("Downloading DwC-A: ", info_row$herbarium, " [", info_row$ipt, "]")
  }

  ok <- tryCatch({
    suppressWarnings(
      utils::download.file(
        archive_url,
        zipfile,
        mode = "wb",
        quiet = !isTRUE(verbose)
      )
    )
    TRUE
  }, error = function(e) FALSE)

  if (!ok && grepl("&v=", archive_url, fixed = TRUE)) {
    ok <- tryCatch({
      suppressWarnings(
        utils::download.file(
          sub("&v=.*$", "", archive_url),
          zipfile,
          mode = "wb",
          quiet = !isTRUE(verbose)
        )
      )
      TRUE
    }, error = function(e) FALSE)
  }

  if (!ok || !file.exists(zipfile)) {
    unlink(out_dir, recursive = TRUE, force = TRUE)
    return(NULL)
  }

  unzip_ok <- tryCatch({
    utils::unzip(zipfile, exdir = out_dir)
    TRUE
  }, error = function(e) FALSE)
  unlink(zipfile)

  if (!unzip_ok || !file.exists(file.path(out_dir, "occurrence.txt"))) {
    unlink(out_dir, recursive = TRUE, force = TRUE)
    return(NULL)
  }

  out_dir
}

#' Connect to herbarium DuckDB cache
#'
#' @return DuckDB connection.
#'
#' @keywords internal
#' @noRd
.herbaria_db_connect <- function() {
  for (pkg in c("DBI", "duckdb")) {
    if (!requireNamespace(pkg, quietly = TRUE)) {
      stop(
        "Package '", pkg, "' is required for herbarium lookup.",
        call. = FALSE
      )
    }
  }

  db_path <- .herbaria_cache_paths()$duckdb_path
  con <- DBI::dbConnect(duckdb::duckdb(), dbdir = db_path)

  DBI::dbExecute(con, "
    CREATE TABLE IF NOT EXISTS occ_raw (
      catalogNumber VARCHAR,
      occurrenceID VARCHAR,
      recordedBy VARCHAR,
      recordNumber VARCHAR,
      recordNumber_clean VARCHAR,
      herbarium VARCHAR,
      source VARCHAR,
      resource_id VARCHAR
    );
  ")

  DBI::dbExecute(con, "
    CREATE TABLE IF NOT EXISTS herbaria_index (
      key VARCHAR,
      key_type VARCHAR,
      catalogNumber VARCHAR,
      occurrenceID VARCHAR,
      herbarium VARCHAR,
      source VARCHAR,
      resource_id VARCHAR
    );
  ")

  try(
    DBI::dbExecute(
      con,
      "CREATE INDEX IF NOT EXISTS idx_occ_num ON occ_raw(herbarium, source, resource_id, recordNumber_clean);"
    ),
    silent = TRUE
  )

  try(
    DBI::dbExecute(
      con,
      "CREATE INDEX IF NOT EXISTS idx_idx_key ON herbaria_index(key_type, key);"
    ),
    silent = TRUE
  )

  con
}

#' Load occurrence.txt into DuckDB
#'
#' @param con DuckDB connection.
#' @param occurrence_path Character path to occurrence.txt.
#' @param herbarium Herbarium code.
#' @param source One of "jabot" or "reflora".
#' @param resource_id IPT resource id.
#' @param verbose Logical.
#'
#' @return Logical.
#'
#' @keywords internal
#' @noRd
.duckdb_load_occurrence <- function(con,
                                    occurrence_path,
                                    herbarium,
                                    source,
                                    resource_id,
                                    verbose = FALSE) {
  if (!file.exists(occurrence_path)) {
    return(FALSE)
  }

  herb_q <- as.character(DBI::dbQuoteString(con, herbarium))
  src_q  <- as.character(DBI::dbQuoteString(con, source))
  rid_q  <- as.character(DBI::dbQuoteString(con, resource_id))
  occ_q  <- as.character(DBI::dbQuoteString(
    con,
    normalizePath(occurrence_path, winslash = "/", mustWork = FALSE)
  ))

  DBI::dbExecute(con, paste0(
    "DELETE FROM occ_raw WHERE herbarium = ", herb_q,
    " AND source = ", src_q,
    " AND resource_id = ", rid_q, ";"
  ))

  DBI::dbExecute(con, paste0(
    "DELETE FROM herbaria_index WHERE herbarium = ", herb_q,
    " AND source = ", src_q,
    " AND resource_id = ", rid_q, ";"
  ))

  probe <- tryCatch(
    DBI::dbGetQuery(
      con,
      paste0(
        "SELECT * FROM read_csv_auto(",
        occ_q,
        ", delim = '\\t', header = TRUE, quote = '', sample_size = 20000, ignore_errors = TRUE) LIMIT 0;"
      )
    ),
    error = function(e) NULL
  )

  if (is.null(probe)) {
    return(FALSE)
  }

  nms_raw <- names(probe)
  nms <- .norm_nm(nms_raw)

  pick_first <- function(cands) {
    cands <- .norm_nm(cands)
    hit <- match(cands, nms, nomatch = 0L)
    if (any(hit > 0L)) {
      return(nms_raw[hit[hit > 0L][1]])
    }
    for (z in cands) {
      ii <- which(grepl(z, nms, fixed = TRUE))
      if (length(ii)) {
        return(nms_raw[ii[1]])
      }
    }
    NA_character_
  }

  catalog_col <- pick_first(c("catalogNumber", "specimen", "barcode", "accession"))
  occid_col <- pick_first(c("occurrenceID", "id"))
  recorded_col <- pick_first(c("recordedBy", "collector", "coletor", "coletores"))
  record_col <- pick_first(c(
    "recordNumber", "collectorNumber", "voucherNumber",
    "verbatimCollectorNumber", "fieldNumber", "otherCatalogNumbers", "numero"
  ))

  sel <- function(col, alias) {
    if (is.na(col)) {
      paste0("NULL::VARCHAR AS ", alias)
    } else {
      paste0('CAST("', gsub('"', '""', col, fixed = TRUE), '" AS VARCHAR) AS ', alias)
    }
  }

  rec_expr <- if (is.na(record_col)) {
    "NULL::VARCHAR"
  } else {
    paste0('CAST("', gsub('"', '""', record_col, fixed = TRUE), '" AS VARCHAR)')
  }

  scan_sql <- paste0(
    "INSERT INTO occ_raw ",
    "SELECT ",
    sel(catalog_col, "catalogNumber"), ", ",
    sel(occid_col, "occurrenceID"), ", ",
    sel(recorded_col, "recordedBy"), ", ",
    sel(record_col, "recordNumber"), ", ",
    "regexp_replace(coalesce(", rec_expr, ", ''), '[^0-9]', '', 'g') AS recordNumber_clean, ",
    herb_q, " AS herbarium, ",
    src_q, " AS source, ",
    rid_q, " AS resource_id ",
    "FROM read_csv_auto(",
    occ_q,
    ", delim = '\\t', header = TRUE, quote = '', sample_size = 20000, ignore_errors = TRUE);"
  )

  ok <- tryCatch({
    DBI::dbExecute(con, scan_sql)
    TRUE
  }, error = function(e) FALSE)

  if (isTRUE(verbose) && isTRUE(ok)) {
    n_ins <- DBI::dbGetQuery(con, paste0(
      "SELECT COUNT(*) AS n FROM occ_raw WHERE herbarium = ", herb_q,
      " AND source = ", src_q,
      " AND resource_id = ", rid_q, ";"
    ))$n
    message("Inserted ", n_ins, " records into database")
  }

  ok
}

#' Match one resource against requested collector-number keys
#'
#' @param con DuckDB connection.
#' @param herbarium Herbarium code.
#' @param source Source string.
#' @param resource_id IPT resource id.
#' @param need_primary Character vector of requested primary keys.
#' @param need_fallback Character vector of requested fallback keys.
#' @param verbose Logical.
#'
#' @return Logical.
#'
#' @keywords internal
#' @noRd
.duckdb_match_resource <- function(con,
                                   herbarium,
                                   source,
                                   resource_id,
                                   need_primary,
                                   need_fallback,
                                   verbose = FALSE) {
  need_primary <- unique(stats::na.omit(as.character(need_primary)))
  need_fallback <- unique(stats::na.omit(as.character(need_fallback)))

  all_numbers <- unique(c(
    sub("^.*\\|\\|", "", need_primary),
    sub("^.*\\|\\|", "", need_fallback)
  ))
  all_numbers <- all_numbers[nzchar(all_numbers)]

  if (!length(all_numbers)) {
    return(FALSE)
  }

  herb_q <- as.character(DBI::dbQuoteString(con, herbarium))
  src_q  <- as.character(DBI::dbQuoteString(con, source))
  rid_q  <- as.character(DBI::dbQuoteString(con, resource_id))
  nums_sql <- paste(sprintf("'%s'", gsub("'", "''", all_numbers)), collapse = ", ")

  qry <- paste0(
    "SELECT catalogNumber, occurrenceID, recordedBy, recordNumber, recordNumber_clean ",
    "FROM occ_raw ",
    "WHERE herbarium = ", herb_q,
    " AND source = ", src_q,
    " AND resource_id = ", rid_q,
    " AND recordNumber_clean IN (", nums_sql, ");"
  )

  cand <- tryCatch(DBI::dbGetQuery(con, qry), error = function(e) NULL)

  if (is.null(cand) || !nrow(cand)) {
    if (isTRUE(verbose)) {
      message("No candidate records found for resource ", resource_id)
    }
    return(FALSE)
  }

  if (isTRUE(verbose)) {
    message("Candidate records found: ", nrow(cand))
  }

  idx_list <- vector("list", nrow(cand))
  pos <- 0L

  for (i in seq_len(nrow(cand))) {
    people <- .split_recordedby_people(cand$recordedBy[i])
    if (!length(people)) {
      next
    }

    tok <- t(vapply(
      people,
      .collector_tokens_one,
      FUN.VALUE = c(primary = NA_character_, fallback = NA_character_)
    ))

    primary_keys <- ifelse(
      !is.na(tok[, "primary"]) & nzchar(tok[, "primary"]),
      paste0(tok[, "primary"], "||", cand$recordNumber_clean[i]),
      NA_character_
    )

    fallback_keys <- ifelse(
      !is.na(tok[, "fallback"]) & nzchar(tok[, "fallback"]),
      paste0(tok[, "fallback"], "||", cand$recordNumber_clean[i]),
      NA_character_
    )

    one <- rbind(
      data.frame(
        key = primary_keys,
        key_type = "primary",
        catalogNumber = cand$catalogNumber[i],
        occurrenceID = cand$occurrenceID[i],
        herbarium = herbarium,
        source = source,
        resource_id = resource_id,
        stringsAsFactors = FALSE
      ),
      data.frame(
        key = fallback_keys,
        key_type = "fallback",
        catalogNumber = cand$catalogNumber[i],
        occurrenceID = cand$occurrenceID[i],
        herbarium = herbarium,
        source = source,
        resource_id = resource_id,
        stringsAsFactors = FALSE
      )
    )

    one <- one[!is.na(one$key) & nzchar(one$key), , drop = FALSE]
    if (!nrow(one)) {
      next
    }

    pos <- pos + 1L
    idx_list[[pos]] <- one
  }

  if (!pos) {
    return(FALSE)
  }

  idx <- do.call(rbind, idx_list[seq_len(pos)])
  idx <- unique(idx)

  keep <- rbind(
    if (length(need_primary)) {
      data.frame(key = need_primary, key_type = "primary", stringsAsFactors = FALSE)
    },
    if (length(need_fallback)) {
      data.frame(key = need_fallback, key_type = "fallback", stringsAsFactors = FALSE)
    }
  )

  idx <- merge(idx, keep, by = c("key", "key_type"), all = FALSE)
  idx <- idx[!duplicated(paste(
    idx$key, idx$key_type, idx$catalogNumber, idx$occurrenceID, idx$source
  )), , drop = FALSE]

  if (!nrow(idx)) {
    if (isTRUE(verbose)) {
      message("Candidates found, but none matched requested keys")
    }
    return(FALSE)
  }

  DBI::dbWriteTable(con, "herbaria_index", idx, append = TRUE)

  if (isTRUE(verbose)) {
    message("Index built with ", nrow(idx), " matched rows")
  }

  TRUE
}

#' Build a JABOT specimen URL
#'
#' @param catalogNumber Character scalar.
#' @param herbarium Character scalar.
#'
#' @return Character scalar URL or NA.
#'
#' @keywords internal
#' @noRd
.make_jabot_url <- function(catalogNumber, herbarium) {
  cat_str <- trimws(as.character(catalogNumber))
  if (is.na(catalogNumber) || !nzchar(cat_str)) {
    return(NA_character_)
  }

  herb_str <- toupper(trimws(as.character(herbarium)))
  if (is.na(herbarium) || !nzchar(herb_str)) {
    return(NA_character_)
  }

  paste0(
    "https://catalogo-ucs-brasil.jbrj.gov.br/regua/visualizador.php",
    "?codtestemunho=",
    utils::URLencode(cat_str, reserved = TRUE),
    "&colbot=",
    herb_str,
    "&db=Jabot&r=true"
  )
}

#' Build a REFLORA image URL
#'
#' @param occurrenceID Character scalar. Aceita URI completa ou ID numerico puro.
#'
#' @return Character scalar URL or NA.
#'
#' @keywords internal
#' @noRd
.make_reflora_url <- function(occurrenceID) {
  if (is.na(occurrenceID) || !nzchar(trimws(as.character(occurrenceID)))) {
    return(NA_character_)
  }

  id_raw <- trimws(as.character(occurrenceID))

  # Caso 1: URI com segmento numerico final  (ex: ".../ExibeFiguraFSIUC/12345")
  m_uri <- regexec("/([0-9]+)\\s*$", id_raw, perl = TRUE)
  if (m_uri[[1]][1] > 0L) {
    id <- regmatches(id_raw, m_uri)[[1]][2]
    if (nzchar(id)) {
      return(paste0(
        "https://reflora.jbrj.gov.br/reflora/geral/ExibeFiguraFSIUC/ExibeFiguraFSIUC.do?idFigura=",
        id
      ))
    }
  }

  # Caso 2: ID puramente numerico (sem nenhum outro caractere)
  if (grepl("^[0-9]+$", id_raw)) {
    return(paste0(
      "https://reflora.jbrj.gov.br/reflora/geral/ExibeFiguraFSIUC/ExibeFiguraFSIUC.do?idFigura=",
      id_raw
    ))
  }

  # Caso 3: fallback - maior bloco numerico continuo (minimo 3 digitos)
  # Cobre formatos como "RB00012345" onde nao ha separador "/"
  nums <- regmatches(id_raw, gregexpr("[0-9]{3,}", id_raw, perl = TRUE))[[1]]
  if (length(nums)) {
    id <- nums[which.max(nchar(nums))]
    return(paste0(
      "https://reflora.jbrj.gov.br/reflora/geral/ExibeFiguraFSIUC/ExibeFiguraFSIUC.do?idFigura=",
      id
    ))
  }

  NA_character_
}

#' Build herbarium links for voucher records using JABOT first and REFLORA only if needed
#'
#' @description
#' Matches voucher records from a field data frame against locally indexed
#' Darwin Core Archive (DwC-A) occurrence data and returns HTML fragments
#' linking to herbarium specimen pages.
#'
#' The lookup workflow is:
#' \enumerate{
#'   \item Parse voucher strings into normalized collector-number keys.
#'   \item Download and index JABOT resources for the requested herbaria.
#'   \item Resolve as many vouchers as possible using JABOT records.
#'   \item Only for still unresolved vouchers, download and index REFLORA
#'         resources and try again.
#' }
#'
#' JABOT links are built from `catalogNumber`, while REFLORA links are built
#' from `occurrenceID`.
#'
#' @param fp_df A data frame containing at least a `Voucher` column.
#' @param herbariums Character vector of herbarium codes, for example `"RB"`.
#' @param force_refresh Logical. If `TRUE`, forces redownload and reindexing
#'   of DwC-A resources.
#' @param keep_downloads Logical. If `FALSE`, temporary downloaded DwC-A folders
#'   are removed at the end of the function.
#' @param verbose Logical. If `TRUE`, prints detailed progress messages.
#' @param collector_fallback Optional character scalar. Collector name used when
#'   the voucher string does not explicitly include one.
#' @param collector_codes Optional named character vector mapping compact
#'   collector prefixes to full collector names.
#' @param resource_map Optional named character vector overriding IPT resource
#'   ids, using names such as `"RB::jabot"` or `"RB::reflora"`.
#'
#' @return A character vector of length `nrow(fp_df)`. Each element is either
#'   `NA_character_` or an HTML fragment containing a herbarium link.
#'
#' @keywords internal
#' @noRd
.herbaria_lookup_links <- function(fp_df,
                                   herbariums,
                                   force_refresh = FALSE,
                                   keep_downloads = FALSE,
                                   verbose = FALSE,
                                   collector_fallback = NULL,
                                   collector_codes = NULL,
                                   resource_map = NULL) {
  .arg_check_herbarium(herbariums)
  herbariums <- unique(toupper(trimws(as.character(herbariums))))

  if (!("Voucher" %in% names(fp_df))) {
    return(rep(NA_character_, nrow(fp_df)))
  }

  # Parse voucher strings into normalized identity keys.
  parsed <- .parse_census_identity(
    voucher = fp_df$Voucher,
    collector_fallback = collector_fallback,
    collector_codes = collector_codes
  )

  need_primary <- unique(stats::na.omit(parsed$primary_key))
  need_fallback <- unique(stats::na.omit(parsed$fallback_key))

  # Nothing to match.
  if (!length(need_primary) && !length(need_fallback)) {
    return(rep(NA_character_, nrow(fp_df)))
  }

  if (isTRUE(verbose)) {
    message("HERBARIA LOOKUP")
    message("Herbaria: ", paste(herbariums, collapse = ", "))
    message("Records: ", nrow(fp_df))
    message("Valid primary keys: ", length(need_primary))
    message("Valid fallback keys: ", length(need_fallback))
  }

  con <- .herbaria_db_connect()
  on.exit(try(DBI::dbDisconnect(con, shutdown = TRUE), silent = TRUE), add = TRUE)

  # Remove download folders at the end unless the user wants to keep them.
  if (!isTRUE(keep_downloads)) {
    on.exit({
      for (d in c("jabot_download", "reflora_download")) {
        if (dir.exists(d)) {
          unlink(d, recursive = TRUE, force = TRUE)
        }
      }
    }, add = TRUE)
  }

  already_loaded <- character(0)

  # Download and load one resource only once.
  load_resource_once <- function(rowi, dir_name) {
    rid_key <- paste(rowi$herbarium, rowi$source, rowi$resource_id, sep = "::")

    if (rid_key %in% already_loaded) {
      return(TRUE)
    }

    folder <- .download_dwca_one(
      info_row = rowi,
      dir = dir_name,
      verbose = verbose,
      force_refresh = force_refresh
    )

    if (is.null(folder)) {
      return(FALSE)
    }

    occ_path <- file.path(folder, "occurrence.txt")

    ok_load <- .duckdb_load_occurrence(
      con = con,
      occurrence_path = occ_path,
      herbarium = rowi$herbarium,
      source = rowi$source,
      resource_id = rowi$resource_id,
      verbose = verbose
    )

    if (isTRUE(ok_load)) {
      already_loaded <<- c(already_loaded, rid_key)
      return(TRUE)
    }

    FALSE
  }

  # Resolve links from an index table already stored in DuckDB.
  resolve_links_from_index <- function(index_df, rows_to_check, source_label) {
    out_local <- rep(NA_character_, nrow(fp_df))
    resolved_local <- rep(FALSE, nrow(fp_df))

    if (!nrow(index_df) || !length(rows_to_check)) {
      return(list(links = out_local, resolved = resolved_local))
    }

    for (i in rows_to_check) {
      pk <- parsed$primary_key[i]
      fk <- parsed$fallback_key[i]

      hit <- index_df[0, , drop = FALSE]

      # Try primary key first.
      if (!is.na(pk) && nzchar(pk)) {
        hit <- index_df[
          index_df$key == pk & index_df$key_type == "primary",
          ,
          drop = FALSE
        ]
      }

      # Fall back to surname + number if needed.
      if (!nrow(hit) && !is.na(fk) && nzchar(fk)) {
        hit <- index_df[
          index_df$key == fk & index_df$key_type == "fallback",
          ,
          drop = FALSE
        ]
      }

      if (!nrow(hit)) {
        next
      }

      if (identical(source_label, "jabot")) {
        # JABOT URLs require catalogNumber and herbarium.
        hit_valid <- hit[
          !is.na(hit$catalogNumber) & nzchar(trimws(hit$catalogNumber)) &
            !is.na(hit$herbarium) & nzchar(trimws(hit$herbarium)),
          ,
          drop = FALSE
        ]

        if (!nrow(hit_valid)) {
          next
        }

        urls <- unique(mapply(
          .make_jabot_url,
          hit_valid$catalogNumber,
          hit_valid$herbarium,
          USE.NAMES = FALSE
        ))

        urls <- urls[!is.na(urls) & nzchar(urls)]

        if (length(urls)) {
          out_local[i] <- paste0(
            "<br/><b>Herbarium image:</b> ",
            sprintf("<a href='%s' target='_blank' title='JABOT'>JABOT</a>", urls[1])
          )
          resolved_local[i] <- TRUE
        }

      } else if (identical(source_label, "reflora")) {
        # REFLORA URLs require occurrenceID.
        hit_valid <- hit[
          !is.na(hit$occurrenceID) & nzchar(trimws(hit$occurrenceID)),
          ,
          drop = FALSE
        ]

        if (!nrow(hit_valid)) {
          next
        }

        urls <- unique(vapply(
          hit_valid$occurrenceID,
          .make_reflora_url,
          character(1)
        ))

        urls <- urls[!is.na(urls) & nzchar(urls)]

        if (length(urls)) {
          out_local[i] <- paste0(
            "<br/><b>Herbarium image:</b> ",
            sprintf("<a href='%s' target='_blank' title='REFLORA'>REFLORA</a>", urls[1])
          )
          resolved_local[i] <- TRUE
        }
      }
    }

    list(links = out_local, resolved = resolved_local)
  }

  out <- rep(NA_character_, nrow(fp_df))
  unresolved <- rep(TRUE, nrow(fp_df))

  herb_sql <- paste(sprintf("'%s'", gsub("'", "''", herbariums)), collapse = ", ")

  # ============================================================================
  # Step 1: Load and match JABOT only
  # ============================================================================
  for (h in herbariums) {
    jabot_res <- .get_ipt_info(
      herbarium = h,
      ipt = "jabot",
      resource_map = resource_map
    )

    if (!nrow(jabot_res)) {
      if (isTRUE(verbose)) {
        message("No JABOT resource found for ", h)
      }
      next
    }

    jabot_res$source <- "jabot"

    if (isTRUE(verbose)) {
      message("JABOT resources found for ", h, ": ", nrow(jabot_res))
      print(jabot_res[, c("ipt", "herbarium", "resource_id", "source")], row.names = FALSE)
    }

    for (i in seq_len(nrow(jabot_res))) {
      rowi <- jabot_res[i, , drop = FALSE]

      ok_load <- load_resource_once(rowi, "jabot_download")
      if (!ok_load) {
        next
      }

      .duckdb_match_resource(
        con = con,
        herbarium = rowi$herbarium,
        source = rowi$source,
        resource_id = rowi$resource_id,
        need_primary = need_primary,
        need_fallback = need_fallback,
        verbose = verbose
      )
    }
  }

  # Query JABOT matches from the local index.
  idx_jabot <- DBI::dbGetQuery(
    con,
    paste0(
      "SELECT key, key_type, catalogNumber, occurrenceID, herbarium, source ",
      "FROM herbaria_index ",
      "WHERE herbarium IN (", herb_sql, ") AND source = 'jabot';"
    )
  )

  jabot_resolved <- resolve_links_from_index(
    index_df = idx_jabot,
    rows_to_check = which(unresolved),
    source_label = "jabot"
  )

  if (any(jabot_resolved$resolved)) {
    out[which(jabot_resolved$resolved)] <- jabot_resolved$links[which(jabot_resolved$resolved)]
    unresolved[which(jabot_resolved$resolved)] <- FALSE
  }

  # ============================================================================
  # Step 2: Only if needed, load and match REFLORA for unresolved records
  # ============================================================================
  if (any(unresolved)) {
    need_primary_rf <- unique(stats::na.omit(parsed$primary_key[unresolved]))
    need_fallback_rf <- unique(stats::na.omit(parsed$fallback_key[unresolved]))

    if (length(need_primary_rf) || length(need_fallback_rf)) {
      for (h in herbariums) {
        reflora_res <- .get_ipt_info(
          herbarium = h,
          ipt = "reflora",
          resource_map = resource_map
        )

        if (!nrow(reflora_res)) {
          if (isTRUE(verbose)) {
            message("No REFLORA resource found for ", h)
          }
          next
        }

        reflora_res$source <- "reflora"

        if (isTRUE(verbose)) {
          message("REFLORA resources found for ", h, ": ", nrow(reflora_res))
          print(reflora_res[, c("ipt", "herbarium", "resource_id", "source")], row.names = FALSE)
        }

        for (i in seq_len(nrow(reflora_res))) {
          rowi <- reflora_res[i, , drop = FALSE]

          ok_load <- load_resource_once(rowi, "reflora_download")
          if (!ok_load) {
            next
          }

          .duckdb_match_resource(
            con = con,
            herbarium = rowi$herbarium,
            source = rowi$source,
            resource_id = rowi$resource_id,
            need_primary = need_primary_rf,
            need_fallback = need_fallback_rf,
            verbose = verbose
          )
        }
      }

      # Query REFLORA matches from the local index.
      idx_reflora <- DBI::dbGetQuery(
        con,
        paste0(
          "SELECT key, key_type, catalogNumber, occurrenceID, herbarium, source ",
          "FROM herbaria_index ",
          "WHERE herbarium IN (", herb_sql, ") AND source = 'reflora';"
        )
      )

      reflora_resolved <- resolve_links_from_index(
        index_df = idx_reflora,
        rows_to_check = which(unresolved),
        source_label = "reflora"
      )

      if (any(reflora_resolved$resolved)) {
        out[which(reflora_resolved$resolved)] <- reflora_resolved$links[which(reflora_resolved$resolved)]
        unresolved[which(reflora_resolved$resolved)] <- FALSE
      }
    }
  }

  out
}
