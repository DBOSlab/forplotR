# Header-row detection, plot metadata and bilingual column names ####
# Authors: Giulia Ottino & Domingos Cardoso
#
# Some field sheets (e.g. the Panara Indigenous Land template) start with a
# block of plot metadata ("Plotcode | Kypa kukari: ABC01") and only then bring
# the header row, with bilingual names ("New Tag No | Pjankaproo paa ha hokjy").
# Everything that needs to know where the table starts, or which column is
# which, goes through the helpers in this file.


#' Normalize a header or label for matching
#'
#' Lower case, accents folded to ASCII, trimmed, inner whitespace collapsed.
#'
#' @keywords internal
#' @noRd
.hdr_norm <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- ""
  x <- gsub(intToUtf8(0x00A0), " ", x, fixed = TRUE)
  x2 <- suppressWarnings(stringi::stri_trans_general(x, "Latin-ASCII"))
  bad <- is.na(x2)
  x2[bad] <- x[bad]
  trimws(gsub("[[:space:]]+", " ", tolower(x2)))
}


#' Candidate names of one header cell: every side of "A | B"
#'
#' @return Character vector of normalized, non-empty candidates.
#' @keywords internal
#' @noRd
.hdr_candidates <- function(cell) {
  if (is.na(cell) || !nzchar(trimws(cell))) return(character())
  parts <- .hdr_norm(strsplit(as.character(cell), "|", fixed = TRUE)[[1]])
  # also without units and punctuation: "X (m)" -> "x m", "Tag_No" -> "tag no"
  loose <- trimws(gsub("[^a-z0-9]+", " ", parts))
  unique(c(parts, loose)[nzchar(c(parts, loose))])
}


#' Column-name aliases, by canonical field-sheet column
#'
#' Aliases are stored already normalized (see `.hdr_norm()`). English,
#' Portuguese and Panara labels are accepted.
#'
#' @keywords internal
#' @noRd
.header_alias_dict <- function() {
  list(
    "New Tag No" = c("new tag no", "tag no", "tag", "tag number", "tagnumber",
                     "tree id", "treeid", "placa", "plaqueta", "n arvore",
                     "numero da arvore", "pjankaproo paa ha hokjy"),
    "New Stem Grouping" = c("new stem grouping", "stem grouping", "stem group",
                            "grupo de fustes", "parinkra ha hokjy"),
    "T1" = c("t1", "sub plot t1", "subplot t1", "subplott1",
             "standardised subplot t1", "standardized subplot t1", "subplot",
             "sub plot", "subparcela", "reta kukara krepaa saa"),
    "T2" = c("t2", "sub plot t2", "subplot t2", "subplott2"),
    "X" = c("x", "x m", "standardised x", "standardized x", "coord x",
            "coordenada x"),
    "Y" = c("y", "y m", "standardised y", "standardized y", "coord y",
            "coordenada y"),
    "Family" = c("family", "familia", "familia botanica", "botanical family",
                 "recommended family", "recommended voucher family",
                 "ksyapiahapiara parikia"),
    "Original determination" = c("original determination", "species", "especie",
                                 "especie botanica", "botanical species",
                                 "scientific name", "nome cientifico",
                                 "recommended species",
                                 "recommended voucher species",
                                 "determinacao", "isi cientifico"),
    "Local name" = c("local name", "especie panara", "panara species",
                     "nome panara", "panara pee jama issi"),
    "Morphospecies" = c("morphospecies", "morfoespecies", "morfoespecie",
                        "morfoespecie", "morphospecie"),
    "D" = c("d", "d1", "dbh", "dap", "dap mm", "d mm", "diameter", "diametro",
            "wy"),
    "POM" = c("pom", "ponto de medida", "maman ni"),
    "Flag1" = c("flag1", "flag 1", "flag", "hokjyha pari hakja"),
    "Height" = c("height", "altura", "pja"),
    "Voucher" = c("voucher", "voucher code", "numero de coleta",
                  "soton pjon ha hokjy"),
    "Silica" = c("silica"),
    "Collected" = c("collected", "coletado", "voucher collected",
                    "individuo coletado", "hoppjon"),
    "Census Notes" = c("census notes", "notas de campo", "field notes",
                       "notes", "notas", "mipypari jinsa"),
    "Vernacular name" = c("vernacular name", "nome vernacular", "nome popular",
                          "common name", "issi pepejama"),
    "CAP" = c("cap", "cap cm", "circumference"),
    "Basal Area" = c("basal area", "area basal")
  )
}


#' Header names that identify a header row but are not canonical columns
#'
#' (ForestPlots.net Query Library and MONITORA layouts.)
#'
#' @keywords internal
#' @noRd
.header_extra_known <- function() {
  c("plot code", "plotcode", "plot name", "census no", "census date",
    "nome estacao", "n estacao", "estacao", "n parcela", "subunidades",
    "cap tot", "genero", "genus", "flag2", "flag3", "li", "ci", "cf", "cd1",
    "extrad", "extrapom", "nrdups")
}


#' Regular expression of a metadata cell: "Label: value" or "Label | Label: value"
#'
#' The value may be empty, as in a blank template ("Plotcode | Kypa kukari:").
#'
#' @keywords internal
#' @noRd
.metadata_cell_pattern <- function() "^[^:|]+(\\|[^:]+)?:\\s*.*$"


#' Canonical column of every header cell
#'
#' @param header Character vector with the header cells.
#' @return Character vector, `NA` where the cell is not a known column.
#' @keywords internal
#' @noRd
.match_header_cells <- function(header) {
  dict <- .header_alias_dict()
  alias <- unlist(dict, use.names = FALSE)
  canon <- rep(names(dict), lengths(dict))
  vapply(as.character(header), function(cell) {
    hit <- match(.hdr_candidates(cell), alias, nomatch = 0L)
    hit <- hit[hit > 0L]
    if (length(hit)) canon[hit[1L]] else NA_character_
  }, character(1), USE.NAMES = FALSE)
}


#' Score every scanned row as a candidate header row
#'
#' A cell scores one point when any side of its "A | B" label is a known
#' column name; a cell that looks like metadata ("Label: value") scores minus
#' one.
#'
#' @param raw Data frame read without column names.
#' @param extra_targets Extra column names to accept (e.g. user-mapped names).
#' @return List with `row` (best row, or `NA` when no row reaches `min_hits`),
#'   `scores` and `hits`.
#' @keywords internal
#' @noRd
.score_header_rows <- function(raw, max_scan = 20L, extra_targets = NULL,
                               min_hits = 3L) {
  n <- min(NROW(raw), max_scan)
  if (n < 1L) return(list(row = NA_integer_, scores = numeric(), hits = integer()))

  known <- unique(c(unlist(.header_alias_dict(), use.names = FALSE),
                    .header_extra_known(),
                    .hdr_norm(extra_targets)))
  known <- known[nzchar(known)]
  meta_re <- .metadata_cell_pattern()

  hits <- integer(n)
  scores <- numeric(n)
  for (i in seq_len(n)) {
    cells <- as.character(unlist(raw[i, , drop = TRUE], use.names = FALSE))
    cells <- cells[!is.na(cells) & nzchar(trimws(cells))]
    if (!length(cells)) next
    is_meta <- grepl(meta_re, trimws(cells), perl = TRUE)
    is_hit <- vapply(cells, function(cl) any(.hdr_candidates(cl) %in% known),
                     logical(1), USE.NAMES = FALSE) & !is_meta
    hits[i] <- sum(is_hit)
    scores[i] <- hits[i] - sum(is_meta)
  }

  ok <- which(hits >= min_hits)
  best <- if (length(ok)) ok[which.max(scores[ok])] else NA_integer_
  list(row = as.integer(best), scores = scores, hits = hits)
}


#' Read the first rows of a worksheet as text, at their true row numbers
#'
#' @keywords internal
#' @noRd
.read_raw_rows <- function(path, sheet = 1, n_max = Inf) {
  rng <- readxl::cell_limits(c(1L, 1L), c(if (is.finite(n_max)) as.integer(n_max) else NA_integer_, NA_integer_))
  raw <- suppressMessages(readxl::read_excel(
    path, sheet = sheet, range = rng, col_names = FALSE, col_types = "text",
    .name_repair = "minimal"
  ))
  as.data.frame(raw, stringsAsFactors = FALSE, check.names = FALSE)
}


#' Validate a user-supplied `header_row`
#'
#' @keywords internal
#' @noRd
.arg_check_header_row <- function(header_row) {
  if (is.null(header_row)) return(NULL)
  if (!is.numeric(header_row) || length(header_row) != 1L || is.na(header_row) ||
      header_row < 1 || header_row != round(header_row)) {
    stop("`header_row` must be NULL (automatic) or a single positive whole number.",
         call. = FALSE)
  }
  as.integer(header_row)
}


#' Header row of a raw table: user value, detected row, or a fallback
#'
#' @param fallback Row used when nothing is detected (`NULL` warns and uses 1).
#' @keywords internal
#' @noRd
.resolve_header_row <- function(raw, header_row = NULL, extra_targets = NULL,
                                max_scan = 20L, fallback = NULL) {
  header_row <- .arg_check_header_row(header_row)
  if (!is.null(header_row)) {
    cells <- if (header_row <= NROW(raw)) {
      as.character(unlist(raw[header_row, , drop = TRUE], use.names = FALSE))
    } else {
      character()
    }
    if (!any(!is.na(cells) & nzchar(trimws(cells)))) {
      stop("`header_row` (", header_row, ") is an empty row or is beyond the last row ",
           "of the sheet.", call. = FALSE)
    }
    return(header_row)
  }
  best <- .score_header_rows(raw, max_scan = max_scan, extra_targets = extra_targets)$row
  if (!is.na(best)) return(best)
  if (!is.null(fallback)) return(as.integer(fallback))
  warning(
    "Could not find a header row in the first ", min(NROW(raw), max_scan),
    " rows (fewer than 3 known column names in any of them); using row 1. ",
    "Set `header_row` to the row that holds the column names.",
    call. = FALSE
  )
  1L
}


#' Detect the header row of a plot spreadsheet
#'
#' @description
#' Finds the row that holds the column names of a field sheet. Some templates,
#' such as the Indigenous Land (Panará) field sheet, start with one or more rows
#' of plot metadata (\code{"Plotcode | Kypa kukâri: ABC01"}) before the header.
#' The first \code{max_scan} rows are read as text and each one is scored by the
#' number of cells that match a column name known to the package (tag, stem
#' grouping, subplot, X, Y, family, species, diameter, POM, flag, height,
#' voucher, CAP and so on). Matching ignores case and accents, and accepts either
#' side of a bilingual label written as \code{"Name | Translation"}. Cells that
#' look like metadata (\code{"Label: value"}) count against the row.
#'
#' The row with the highest score wins, provided it has at least three known
#' column names; ties go to the first row. When no row qualifies, row 1 is
#' returned with a warning.
#'
#' @param path Path to an Excel file (\code{.xlsx} or \code{.xls}).
#' @param sheet Worksheet name or index.
#' @param max_scan Number of rows, from the top of the sheet, to examine.
#'
#' @return A single integer: the row number of the header, counted from the
#' first row of the worksheet.
#'
#' @seealso [read_plot_header()], [read_plot_metadata()], [read_plot_sheet()]
#'
#' @examples
#' # Blank template of the Panará field sheet: metadata labels on row 1,
#' # bilingual header on row 2
#' f <- system.file("extdata", "panara_field_sheet_template.xlsx",
#'                  package = "forplotR")
#' detect_header_row(f)
#'
#' @export
detect_header_row <- function(path, sheet = 1, max_scan = 20) {
  .check_xlsx_path(path)
  if (!is.numeric(max_scan) || length(max_scan) != 1L || is.na(max_scan) || max_scan < 1) {
    stop("`max_scan` must be a single positive number.", call. = FALSE)
  }
  raw <- .read_raw_rows(path, sheet = sheet, n_max = max_scan)
  .resolve_header_row(raw, max_scan = as.integer(max_scan))
}


#' @keywords internal
#' @noRd
.check_xlsx_path <- function(path) {
  if (!is.character(path) || length(path) != 1L || is.na(path) || !file.exists(path)) {
    stop("`path` must be the path to an existing Excel file.", call. = FALSE)
  }
  invisible(TRUE)
}


#' Metadata labels, by standard key
#'
#' @keywords internal
#' @noRd
.metadata_label_dict <- function() {
  list(
    plot_code = c("plotcode", "plot code", "codigo da parcela", "codigo",
                  "kypa kukari"),
    plot_name = c("plot name", "plotname", "nome da parcela", "parcela",
                  "issi reta kukari"),
    date = c("date", "data", "census date", "data do censo", "haka"),
    team = c("team", "equipe", "pi", "saperate"),
    census_no = c("census no", "census", "censo", "numero do censo")
  )
}


#' Parse a date written as text, day first
#'
#' @keywords internal
#' @noRd
.parse_meta_date <- function(x) {
  x <- trimws(as.character(x))
  if (!length(x) || is.na(x) || !nzchar(x)) return(as.Date(NA))
  # a date cell read as text arrives as an Excel serial number
  if (grepl("^[0-9]{4,6}(\\.[0-9]+)?$", x)) {
    return(as.Date(floor(as.numeric(x)), origin = "1899-12-30"))
  }
  x <- sub("[ T][0-9]{1,2}:[0-9]{2}.*$", "", x)
  for (fmt in c("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d", "%Y/%m/%d",
                "%m/%d/%Y", "%d/%m/%y", "%d %b %Y", "%d %B %Y")) {
    d <- suppressWarnings(as.Date(x, format = fmt))
    if (is.na(d)) next
    yr <- as.integer(format(d, "%Y"))
    if (yr >= 1900L && yr <= 2200L) return(d)
  }
  as.Date(NA)
}


#' Split a team string into people
#'
#' @keywords internal
#' @noRd
.split_team <- function(x) {
  x <- trimws(as.character(x))
  if (!length(x) || is.na(x) || !nzchar(x)) return(character())
  out <- trimws(strsplit(x, "\\s*[,;&]\\s*")[[1]])
  out[nzchar(out)]
}


#' Parse the metadata rows that sit above the header
#'
#' @param rows Data frame (or character vector) with the rows above the header.
#' @return List with `plot_code`, `plot_name`, `date`, `team`, `census_no` and
#'   `other` (named character vector of labels that are not recognised).
#' @keywords internal
#' @noRd
.parse_plot_metadata <- function(rows) {
  out <- list(plot_code = NA_character_, plot_name = NA_character_,
              date = as.Date(NA), team = character(),
              census_no = NA_character_, other = character())

  cells <- if (is.data.frame(rows)) {
    as.character(unlist(t(as.matrix(rows)), use.names = FALSE))
  } else {
    as.character(rows)
  }
  cells <- trimws(cells[!is.na(cells)])
  cells <- cells[grepl(.metadata_cell_pattern(), cells, perl = TRUE)]
  if (!length(cells)) return(out)

  dict <- .metadata_label_dict()
  lab <- unlist(dict, use.names = FALSE)
  key_of <- rep(names(dict), lengths(dict))

  for (cl in cells) {
    label <- sub(":.*$", "", cl)
    value <- trimws(sub("^[^:]*:", "", cl))
    if (!nzchar(value)) next
    hit <- match(.hdr_candidates(label), lab, nomatch = 0L)
    hit <- hit[hit > 0L]
    if (!length(hit)) {
      out$other[[trimws(strsplit(label, "|", fixed = TRUE)[[1]][1])]] <- value
      next
    }
    key <- key_of[hit[1L]]
    if (key == "date") {
      if (is.na(out$date)) out$date <- .parse_meta_date(value)
    } else if (key == "team") {
      if (!length(out$team)) out$team <- .split_team(value)
    } else if (is.na(out[[key]])) {
      out[[key]] <- value
    }
  }
  out
}


#' Number of metadata fields that were actually read
#'
#' @keywords internal
#' @noRd
.metadata_fields <- function(meta) {
  c(
    plot_code = if (!is.na(meta$plot_code)) meta$plot_code,
    plot_name = if (!is.na(meta$plot_name)) meta$plot_name,
    date = if (!is.na(meta$date)) format(meta$date, "%Y-%m-%d"),
    team = if (length(meta$team)) paste0(length(meta$team), if (length(meta$team) == 1L) " person" else " people"),
    census_no = if (!is.na(meta$census_no)) meta$census_no
  )
}


#' One-line, user-facing summary of what was detected
#'
#' @keywords internal
#' @noRd
.header_message <- function(header_row, meta) {
  f <- .metadata_fields(meta)
  n <- length(f) + length(meta$other)
  if (!n) {
    return(paste0("Header detected on row ", header_row, "; no metadata rows above it."))
  }
  paste0(
    "Header detected on row ", header_row, "; ", n, " metadata field",
    if (n == 1L) "" else "s", " read (",
    paste(paste(names(f), f, sep = " = "), collapse = ", "),
    if (length(meta$other)) paste0(if (length(f)) ", " else "", paste(names(meta$other), collapse = ", ")) else "",
    ")."
  )
}


#' Read the plot metadata written above the header of a field sheet
#'
#' @description
#' Reads every row above the header row and turns the cells written as
#' \code{"Label: value"} or \code{"Label | Translation: value"} into a list
#' with standard names. Labels are accepted in English, Portuguese or Panará
#' (for example \code{"Plotcode"}, \code{"Código da parcela"} or
#' \code{"Kypa kukâri"}); matching ignores case and accents.
#'
#' @param path Path to an Excel file (\code{.xlsx} or \code{.xls}).
#' @param sheet Worksheet name or index.
#' @param header_row Row that holds the column names. The default \code{NULL}
#' finds it with [detect_header_row()].
#'
#' @return A list with:
#' \describe{
#'   \item{plot_code}{Plot code, or \code{NA}.}
#'   \item{plot_name}{Plot name, or \code{NA}.}
#'   \item{date}{Census date as a \code{Date} (day/month/year is tried first),
#'     or \code{NA}.}
#'   \item{team}{Character vector with one element per person; names separated
#'     by commas, semicolons or \code{&} are split.}
#'   \item{census_no}{Census number, or \code{NA}.}
#'   \item{other}{Named character vector with any other \code{"Label: value"}
#'     cell found.}
#'   \item{header_row}{The header row that was used.}
#' }
#'
#' @seealso [detect_header_row()], [read_plot_sheet()]
#'
#' @examples
#' # A sheet with one metadata row above the header
#' f <- tempfile(fileext = ".xlsx")
#' openxlsx::write.xlsx(
#'   data.frame(
#'     a = c("Plotcode | Kypa kukâri: ABC01", "New Tag No", "1"),
#'     b = c("Date | Haka: 26/11/2014", "T1", "1"),
#'     c = c("Team | Sâpêrãtê: Ana Souza, Bia Lima & Caio Reis", "X", "2.5"),
#'     d = c(NA, "Y", "3")
#'   ),
#'   f, colNames = FALSE
#' )
#' meta <- read_plot_metadata(f)
#' meta$plot_code
#' meta$date
#' meta$team
#' meta$header_row
#'
#' @export
read_plot_metadata <- function(path, sheet = 1, header_row = NULL) {
  .check_xlsx_path(path)
  header_row <- .arg_check_header_row(header_row)
  raw <- .read_raw_rows(path, sheet = sheet,
                        n_max = if (is.null(header_row)) 20L else header_row)
  hr <- .resolve_header_row(raw, header_row = header_row)
  meta <- .parse_plot_metadata(raw[seq_len(hr - 1L), , drop = FALSE])
  meta$header_row <- hr
  meta
}


#' Final column names for a header row
#'
#' Known columns get their canonical name; other columns keep the first side
#' of their "A | B" label. Names that repeat after that (e.g. two columns that
#' both read as "D") are made unique in sheet order: "D", "D_2", "D_3".
#'
#' @param header Character vector with the header cells.
#' @return Data frame with one row per column: `position`, `header` (as
#'   written), `name` (the name the column has when the sheet is read, which is
#'   what a `colname_*` argument must give: the header as written, `NA_col_<n>`
#'   for an empty header cell, `.1`, `.2` added to a repeated header),
#'   `column` (final name), `canonical` (the canonical field-sheet column it
#'   stands for, or `NA`) and `duplicate_of`.
#' @keywords internal
#' @noRd
.canonical_header <- function(header) {
  header <- as.character(header)
  header[is.na(header)] <- ""
  name <- .header_read_names(header)
  header <- trimws(gsub("[[:space:]]+", " ", header))
  canon <- .match_header_cells(header)

  short <- trimws(vapply(strsplit(header, "|", fixed = TRUE),
                         function(p) if (length(p)) p[1L] else "", character(1)))
  base <- ifelse(is.na(canon), short, canon)
  empty <- !nzchar(base)
  base[empty] <- paste0("NA_col_", seq_along(base))[empty]

  key <- .hdr_norm(base)
  nth <- stats::ave(seq_along(key), key, FUN = seq_along)
  column <- ifelse(nth > 1L, paste0(base, "_", nth), base)
  duplicate_of <- ifelse(nth > 1L, base, NA_character_)
  canon[nth > 1L] <- NA_character_

  data.frame(position = seq_along(header), header = header, name = name,
             column = column, canonical = canon, duplicate_of = duplicate_of,
             stringsAsFactors = FALSE)
}


#' Column names of a sheet as the readers assign them
#'
#' One rule for every reader and for [read_plot_header()]: the header cell as
#' written (trimmed), `NA_col_<position>` for an empty cell, and `make.unique()`
#' for repeated cells (`"D"`, `"D.1"`).
#'
#' @keywords internal
#' @noRd
.header_read_names <- function(header) {
  header <- trimws(as.character(header))
  empty <- is.na(header) | !nzchar(header)
  header[empty] <- paste0("NA_col_", seq_along(header))[empty]
  make.unique(header)
}


#' Read the column names of a plot spreadsheet from its real header row
#'
#' @description
#' Returns the columns of a field sheet as the package reads them: the header
#' row is found with [detect_header_row()] (the same detection every reader
#' uses), so rows of plot metadata above the header never show up as columns.
#' Use it to see which names are available for the \code{colname_*} arguments
#' of [forplot_balance()] and [forplot_map()], or to fill a column picker.
#'
#' @param path Path to an Excel file (\code{.xlsx} or \code{.xls}).
#' @param sheet Worksheet name or index.
#' @param header_row Row that holds the column names. The default \code{NULL}
#' finds it automatically; give a number to force it.
#'
#' @return A data frame with one row per column, in sheet order, and the
#' attribute \code{header_row}:
#' \describe{
#'   \item{position}{Column position in the sheet.}
#'   \item{header}{The header cell as written (\code{""} when empty).}
#'   \item{name}{The column name to give to a \code{colname_*} argument: the
#'     header as written; \code{NA_col_<position>} for an empty header cell;
#'     \code{.1}, \code{.2} added when the same header is repeated.}
#'   \item{canonical}{The column the package recognises it as (for example
#'     \code{"New Tag No"}, \code{"T1"}, \code{"D"}), or \code{NA}.}
#'   \item{has_data}{Whether any cell below the header is filled.}
#' }
#'
#' @seealso [detect_header_row()], [read_plot_sheet()]
#'
#' @examples
#' f <- system.file("extdata", "panara_field_sheet_template.xlsx",
#'                  package = "forplotR")
#' cols <- read_plot_header(f)
#' attr(cols, "header_row")
#' cols[, c("position", "name", "canonical")]
#'
#' @export
read_plot_header <- function(path, sheet = 1, header_row = NULL) {
  .check_xlsx_path(path)
  header_row <- .arg_check_header_row(header_row)
  raw <- .read_raw_rows(path, sheet = sheet)
  if (!nrow(raw)) stop("The worksheet is empty.", call. = FALSE)
  hr <- .resolve_header_row(raw, header_row = header_row)

  cols <- .canonical_header(as.character(unlist(raw[hr, , drop = TRUE])))
  body <- raw[-seq_len(hr), , drop = FALSE]
  cols$has_data <- vapply(seq_len(nrow(cols)), function(i) {
    v <- body[[i]]
    any(!is.na(v) & nzchar(trimws(v)))
  }, logical(1))

  out <- cols[, c("position", "header", "name", "canonical", "has_data")]
  # an empty header over an empty column is not a column
  out <- out[nzchar(out$header) | out$has_data, , drop = FALSE]
  rownames(out) <- NULL
  attr(out, "header_row") <- hr
  out
}


#' Read a field sheet from its real header row
#'
#' @description
#' Reads a plot spreadsheet whose header may not be on the first row and whose
#' column names may be bilingual (\code{"Name | Translation"}). The header row
#' is found with [detect_header_row()], the rows above it are read as plot
#' metadata with [read_plot_metadata()], and the columns are renamed to the
#' names used throughout the package (\code{"New Tag No"}, \code{"T1"},
#' \code{"X"}, \code{"Y"}, \code{"Family"}, \code{"Original determination"},
#' \code{"D"}, \code{"POM"}, \code{"Flag1"}, \code{"Voucher"} and so on).
#' Columns that are not recognised keep the first side of their label.
#'
#' When two columns end up with the same name (for example \code{"D | Wy"} and
#' a helper column \code{"D"}), the first one in the sheet keeps the name and
#' the others get a numeric suffix (\code{"D_2"}); a message says so. Cells
#' holding formulas are read by their stored value.
#'
#' @param path Path to an Excel file (\code{.xlsx} or \code{.xls}).
#' @param sheet Worksheet name or index.
#' @param header_row Row that holds the column names. The default \code{NULL}
#' finds it automatically; give a number to force it.
#' @param verbose Logical. If \code{TRUE}, reports the header row, the metadata
#' that were read and any column that had to be renamed.
#'
#' @return A data frame with one row per record (empty rows are dropped) and
#' three attributes:
#' \describe{
#'   \item{header_row}{The header row that was used.}
#'   \item{metadata}{The list returned by [read_plot_metadata()].}
#'   \item{columns}{A data frame describing each column: its position in the
#'     sheet, the header as written, the name it received, the canonical
#'     column it was recognised as (\code{NA} when not recognised) and, for a
#'     repeated name, the column it duplicates.}
#' }
#'
#' @seealso [detect_header_row()], [read_plot_metadata()]
#'
#' @examples
#' # The blank Panará template has no trees yet, but shows how each
#' # bilingual header is named after reading
#' f <- system.file("extdata", "panara_field_sheet_template.xlsx",
#'                  package = "forplotR")
#' tpl <- read_plot_sheet(f)
#' attr(tpl, "header_row")
#' attr(tpl, "columns")[, c("position", "header", "column")]
#'
#' @export
read_plot_sheet <- function(path, sheet = 1, header_row = NULL, verbose = TRUE) {
  .check_xlsx_path(path)
  header_row <- .arg_check_header_row(header_row)

  top <- .read_raw_rows(path, sheet = sheet,
                        n_max = if (is.null(header_row)) 20L else header_row)
  if (!nrow(top)) stop("The worksheet is empty.", call. = FALSE)
  hr <- .resolve_header_row(top, header_row = header_row)
  meta <- .parse_plot_metadata(top[seq_len(hr - 1L), , drop = FALSE])
  meta$header_row <- hr

  dat <- suppressMessages(readxl::read_excel(
    path, sheet = sheet,
    range = readxl::cell_limits(c(hr, 1L), c(NA_integer_, NA_integer_)),
    col_names = FALSE, .name_repair = "minimal", guess_max = 1048576L,
    col_types = "list"
  ))
  dat <- as.data.frame(dat, stringsAsFactors = FALSE, check.names = FALSE)

  header <- vapply(dat[1L, , drop = TRUE], function(v) {
    v <- v[[1L]]
    if (is.null(v) || !length(v) || is.na(v)) "" else as.character(v)
  }, character(1), USE.NAMES = FALSE)
  cols <- .canonical_header(header)

  body <- dat[-1L, , drop = FALSE]
  body[] <- lapply(body, .simplify_list_column)
  names(body) <- cols$column

  blank <- vapply(body, function(v) all(is.na(v) | !nzchar(trimws(as.character(v)))), logical(1))
  keep_col <- !(blank & grepl("^NA_col_[0-9]+$", cols$column))
  body <- body[, keep_col, drop = FALSE]
  cols <- cols[keep_col, , drop = FALSE]

  if (ncol(body)) {
    empty_row <- Reduce(`&`, lapply(body, function(v) is.na(v) | !nzchar(trimws(as.character(v)))))
    body <- body[!empty_row, , drop = FALSE]
  }
  rownames(body) <- NULL
  rownames(cols) <- NULL

  if (isTRUE(verbose)) {
    message(.header_message(hr, meta))
    dup <- cols[!is.na(cols$duplicate_of), , drop = FALSE]
    for (i in seq_len(nrow(dup))) {
      first <- dup$duplicate_of[i]
      message("Column ", dup$position[i], " ('", dup$header[i], "') repeats the name '",
              first, "' and was read as '", dup$column[i], "'; '", first,
              "' is column ", cols$position[match(first, cols$column)], " ('",
              cols$header[match(first, cols$column)], "').")
    }
  }

  attr(body, "header_row") <- hr
  attr(body, "metadata") <- meta
  attr(body, "columns") <- cols
  body
}


#' Collapse a list column read by readxl into the simplest atomic vector
#'
#' Numbers stay numeric, dates become ISO text only when mixed with text, and
#' a column with both numbers and text becomes character.
#'
#' @keywords internal
#' @noRd
.simplify_list_column <- function(v) {
  if (!is.list(v)) return(v)
  n <- length(v)
  is_null <- vapply(v, function(z) is.null(z) || !length(z) || (length(z) == 1L && is.na(z)), logical(1))
  if (all(is_null)) return(rep(NA, n))
  cls <- vapply(v[!is_null], function(z) class(z)[1L], character(1))

  if (all(cls %in% c("numeric", "integer"))) {
    out <- rep(NA_real_, n)
    out[!is_null] <- vapply(v[!is_null], function(z) as.numeric(z[[1L]]), numeric(1))
    return(out)
  }
  if (all(cls == "logical")) {
    out <- rep(NA, n)
    out[!is_null] <- vapply(v[!is_null], function(z) as.logical(z[[1L]]), logical(1))
    return(out)
  }
  out <- rep(NA_character_, n)
  out[!is_null] <- vapply(v[!is_null], function(z) {
    z <- z[[1L]]
    if (inherits(z, c("POSIXct", "Date"))) return(format(z, "%Y-%m-%d"))
    if (is.numeric(z)) return(format(z, digits = 15, trim = TRUE, scientific = FALSE))
    as.character(z)
  }, character(1))
  out
}
