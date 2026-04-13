#' Generate forest plot specimen map and collection balance
#'
#' @author Giulia Ottino & Domingos Cardoso
#'
#' @description Processes field data collected using the
#' \href{https://forestplots.net/}{ForestPlots.net} format (field sheets or
#' Query Library output) or
#' \href{https://www.gov.br/icmbio/pt-br/assuntos/monitoramento/programa-monitora}{Monitora program}
#' layouts, and generates a specimen map with collection status and spatial
#' Distribution of individuals across subplots. It creates a PDF report with
#' plot-level and subplot-level maps and an Excel spreadsheet summarizing the
#' percentage of collected specimens per subplot. The function performs the
#' following steps: (i) validates all input arguments and checks/creates output
#' folders; (ii) reads the input sheet, extracts metadata (team, plot name, plot
#' code), and cleans the data; (iii) normalizes and cleans coordinate and
#' diameter values, computing global plot coordinates; (iv) builds a PDF report
#' including plot metadata, the main map, collected and uncollected specimen
#' maps, an optional palm map and navigable subplot maps; (v) optionally
#' generates a spreadsheet with the collection percentage per subplot (including
#' totals), separating palms from non-palms.
#'
#' @usage
#' plot_for_balance(fp_file_path = NULL,
#'                  language = c("en", "pt", "es", "fr", "ma", "pa"),
#'                  input_type = c("field_sheet", "field_sheet_ti", "fp_query_sheet", "monitora"),
#'                  plot_size = 1,
#'                  subplot_size = 10,
#'                  plot_width_m = 100,
#'                  plot_length_m = NULL,
#'                  highlight_palms = TRUE,
#'                  station_name = NULL,
#'                  plot_name = NULL,
#'                  plot_code = NULL,
#'                  plot_census_no = NULL,
#'                  team = NULL,
#'                  render_html = TRUE,
#'                  render_pdf = TRUE,
#'                  write_xlsx = TRUE,
#'                  dir = "Results_map_plot",
#'                  filename = "plot_specimen")
#'
#' @param fp_file_path Path to the Excel file (field or query sheet) in
#' ForestPlots format.
#'
#' @param language Character. One of en (english), pt (portuguese), es
#' (spanish), fr (french), ma (mandarin), pa (panara) (default = en)
#'
#' @param input_type Character. One of `"field_sheet"`, `"field_sheet_ti"`,
#' `"fp_query_sheet"` or `"monitora"`. Specifies the type/layout of the input file.
#'
#' @param plot_size Total plot size in hectares. Used for `"field_sheet"` and
#' `"fp_query_sheet"` inputs. Ignored when `input_type = "monitora"`.
#'
#' @param subplot_size Side length of each subplot in meters (default = 10).
#'
#' @param plot_width_m Plot width in meters. Used for `"field_sheet"` and
#' `"fp_query_sheet"` inputs. Ignored when `input_type = "monitora"`.
#'
#' @param plot_length_m Plot length in meters. Used for `"field_sheet"` and
#' `"fp_query_sheet"` inputs. Ignored when `input_type = "monitora"`.
#'
#' @param highlight_palms Logical. If `TRUE`, highlights Arecaceae individuals
#' in the plots.
#'
#' @param station_name Optional station identifier used when `input_type = "monitora"`.
#' If a vector of length > 1, the function will recursively generate a report
#' and summary spreadsheet for each station.
#'
#' @param plot_name Optional plot name. If provided, overrides the plot name
#' extracted from the input file metadata.
#'
#' @param plot_code Optional plot code. If provided, overrides the plot code
#' extracted from the input file metadata.
#'
#' @param plot_census_no Optional plot census no. Default is 1, overrides the plot
#' census number extract from a fp_query_sheet input sheet.
#'
#' @param team Optional team or PI name. If provided, overrides the team
#' extracted from the input file metadata.
#'
#' @param render_html Logical. If `TRUE`, renders the HTML report.
#'
#' @param render_pdf Logical. If `TRUE`, renders the PDF report.
#'
#' @param write_xlsx Logical. If `TRUE`, writes the Excel summary workbook.
#'
#' @param dir Directory path where output will be saved
#' (default is `"Results_map_plot"`). A date-stamped subfolder will be created
#' inside this directory.
#'
#' @param filename Basename for output files (without extension).
#'
#' @return Invisibly returns \code{NULL}. Called for its side effects of
#' generating a PDF report and an Excel file summarizing specimen collection
#' statistics per subplot.
#'
#' @import ggplot2
#' @importFrom readxl read_excel
#' @importFrom dplyr mutate if_else group_by arrange select distinct filter bind_rows any_of case_when rename_with rename n
#' @importFrom magrittr "%>%"
#' @importFrom grDevices pdf dev.off colorRampPalette
#' @importFrom utils capture.output download.file read.csv
#' @importFrom rmarkdown render pdf_document
#' @importFrom openxlsx read.xlsx createWorkbook addWorksheet writeData createStyle addStyle saveWorkbook
#' @importFrom tinytex latexmk
#' @importFrom readr locale read_delim
#' @importFrom tibble as_tibble tibble
#' @importFrom here here
#' @importFrom stringr str_to_lower str_trim str_detect str_remove word
#' @importFrom tidyr replace_na
#' @importFrom grid unit
#' @importFrom writexl write_xlsx
#' @importFrom stats na.omit setNames
#' @importFrom plotly plot_ly layout config
#' @importFrom htmlwidgets onRender
#'
#' @examples
#' \dontrun{
#' # Basic usage with field sheet
#' plot_for_balance(
#'   fp_file_path = "data/forestplot.xlsx",
#'   language = "en",
#'   input_type = "field_sheet"
#' )
#'
#' # With Monitora data
#' plot_for_balance(
#'   fp_file_path = "data/monitora_data.xlsx",
#'   input_type = "monitora",
#'   station_name = "Station1"
#' )
#' }
#'
#' @export
#'

plot_for_balance <- function(fp_file_path = NULL,
                             language = c("en", "pt", "es", "fr", "ma", "pa"),
                             input_type = c("field_sheet", "field_sheet_ti", "fp_query_sheet", "monitora"),
                             plot_size = 1,
                             subplot_size = 10,
                             plot_width_m = 100,
                             plot_length_m = NULL,
                             highlight_palms = TRUE,
                             station_name = NULL,
                             plot_name = NULL,
                             plot_code = NULL,
                             plot_census_no = NULL,
                             team = NULL,
                             render_html = TRUE,
                             render_pdf = TRUE,
                             write_xlsx = TRUE,
                             dir = "Results_map_plot",
                             filename = "plot_specimen") {

  .validate_subplot_size(subplot_size)

  input_type <- tolower(trimws(as.character(input_type)))
  input_type <- match.arg(input_type, c("field_sheet", "field_sheet_ti", "fp_query_sheet", "monitora"))

  if (!is.logical(render_html) || length(render_html) != 1 || is.na(render_html)) {
    stop("`render_html` must be TRUE or FALSE.", call. = FALSE)
  }

  if (!is.logical(render_pdf) || length(render_pdf) != 1 || is.na(render_pdf)) {
    stop("`render_pdf` must be TRUE or FALSE.", call. = FALSE)
  }

  if (!is.logical(write_xlsx) || length(write_xlsx) != 1 || is.na(write_xlsx)) {
    stop("`write_xlsx` must be TRUE or FALSE.", call. = FALSE)
  }

  if (input_type %in% c("field_sheet", "field_sheet_ti", "fp_query_sheet")) {
    .validate_plot_size(plot_size)

    plot_area_m2 <- plot_size * 10000

    if (!is.null(plot_width_m)) {
      if (!is.numeric(plot_width_m) || length(plot_width_m) != 1 || is.na(plot_width_m) || plot_width_m <= 0) {
        stop("`plot_width_m` must be a single positive numeric value.", call. = FALSE)
      }
    }

    if (!is.null(plot_length_m)) {
      if (!is.numeric(plot_length_m) || length(plot_length_m) != 1 || is.na(plot_length_m) || plot_length_m <= 0) {
        stop("`plot_length_m` must be a single positive numeric value.", call. = FALSE)
      }
    }

    if (is.null(plot_width_m) && is.null(plot_length_m)) {
      plot_width_m <- 100
    }

    if (is.null(plot_length_m)) {
      plot_length_m <- plot_area_m2 / plot_width_m
    }

    if (is.null(plot_width_m)) {
      plot_width_m <- plot_area_m2 / plot_length_m
    }

    if (abs((plot_width_m * plot_length_m) - plot_area_m2) > 1e-6) {
      stop(
        "`plot_width_m * plot_length_m` must match the area implied by `plot_size`.",
        call. = FALSE
      )
    }

    if ((plot_width_m %% subplot_size) != 0 || (plot_length_m %% subplot_size) != 0) {
      stop(
        "`plot_width_m` and `plot_length_m` must be exact multiples of `subplot_size`.",
        call. = FALSE
      )
    }
  } else {
    plot_width_m <- NULL
    plot_length_m <- NULL
  }

  language <- match.arg(tolower(trimws(as.character(language))),
                        c("en", "pt", "es", "fr", "ma", "pa"))

  tr <- .plot_i18n(language)

  dir <- .arg_check_dir(dir)

  foldername <- file.path(dir, format(Sys.time(), "%d%b%Y"))
  if (!dir.exists(dir)) dir.create(dir, recursive = TRUE)
  if (!dir.exists(foldername)) dir.create(foldername, recursive = TRUE)

  arg_plot_name <- plot_name
  arg_plot_code <- plot_code
  arg_census_no_fp <- plot_census_no
  arg_team <- team

  if (input_type == "monitora" && !is.null(station_name) && length(station_name) > 1L) {
    for (st in unique(trimws(as.character(station_name)))) {
      plot_for_balance(
        fp_file_path = fp_file_path,
        input_type = "monitora",
        plot_size = plot_size,
        subplot_size = subplot_size,
        plot_width_m = plot_width_m,
        plot_length_m = plot_length_m,
        highlight_palms = highlight_palms,
        station_name = st,
        plot_name = arg_plot_name,
        plot_code = arg_plot_code,
        team = arg_team,
        render_html = render_html,
        render_pdf = render_pdf,
        write_xlsx = write_xlsx,
        dir = dir,
        filename = paste0(filename, "_station_", st),
        language = language
      )
    }
    return(invisible(TRUE))
  }

  raw <- switch(
    input_type,
    "monitora" = .monitora_to_field_sheet_df(fp_file_path, station_name = station_name),
    "fp_query_sheet" = .fp_query_to_field_sheet_df(fp_file_path),
    "field_sheet" = suppressMessages(
      readxl::read_excel(
        fp_file_path,
        sheet = 1,
        col_names = FALSE,
        .name_repair = "minimal"
      )
    ),
    "field_sheet_ti" = suppressMessages(
      readxl::read_excel(
        fp_file_path,
        sheet = 1,
        col_names = FALSE,
        .name_repair = "minimal"
      )
    )
  )

  message("RAW import check:")
  print(raw[1:15, 1:10])

  canonical_cols <- c(
    "New Tag No", "T1", "T2", "X", "Y",
    "Family", "Original determination", "D"
  )

  if (input_type %in% c("fp_query_sheet", "monitora")) {
    has_meta <- FALSE

    data <- as.data.frame(raw, stringsAsFactors = FALSE)
    raw_meta <- attr(raw, "plot_meta")

    team <- if (!is.null(arg_team) && nzchar(arg_team)) {
      arg_team
    } else if (!is.null(raw_meta) &&
               !is.null(raw_meta$team) &&
               nzchar(trimws(raw_meta$team))) {
      raw_meta$team
    } else {
      ""
    }

    plot_name <- if (!is.null(arg_plot_name) && nzchar(arg_plot_name)) {
      arg_plot_name
    } else if (!is.null(raw_meta) &&
               !is.null(raw_meta$plot_name) &&
               nzchar(trimws(raw_meta$plot_name))) {
      raw_meta$plot_name
    } else {
      "Unknown Plot"
    }

    plot_code <- if (!is.null(raw_meta) &&
                     !is.null(raw_meta$plot_code) &&
                     nzchar(trimws(raw_meta$plot_code))) {
      raw_meta$plot_code
    } else {
      ""
    }

    plot_census_no_fp <- if (!is.null(arg_census_no_fp) && nzchar(arg_census_no_fp)) {
      arg_census_no_fp
    } else if (!is.null(raw_meta) &&
               !is.null(raw_meta$census_no_fp) &&
               nzchar(trimws(raw_meta$census_no_fp))) {
      raw_meta$census_no_fp
    } else {
      ""
    }

    if (!all(canonical_cols %in% names(data))) {
      stop(
        "For input_type = '", input_type,
        "', the imported table must already contain canonical columns. Missing: ",
        paste(setdiff(canonical_cols, names(data)), collapse = ", "),
        call. = FALSE
      )
    }

    message("Input already converted to canonical schema; skipping header detection.")
    message("Detected columns: ", paste(names(data)[1:min(12, ncol(data))], collapse = " | "))
    message("Using metadata from raw/attributes:")
    message("plot_name = ", plot_name)
    message("plot_code = ", plot_code)
    message("team = ", team)
    message("DATA check:")
    print(data[1:min(20, nrow(data)), intersect(c("New Tag No", "T1", "T2", "X", "Y"), names(data))])

    if (!("Collected" %in% names(data))) data$Collected <- NA_character_

  } else {

    metadata_row <- .safe_char_row(raw)

    has_meta <- any(grepl("^\\s*(Plotcode:|Plot Code:|Plot Name:|Team:|PI:)",
                          metadata_row, ignore.case = TRUE))

    team <- if (!is.null(arg_team) && nzchar(arg_team)) {
      arg_team
    } else {
      hit <- metadata_row[grepl("^\\s*Team:", metadata_row, ignore.case = TRUE)]
      if (length(hit)) sub("^\\s*Team:\\s*", "", hit[1], ignore.case = TRUE) else ""
    }

    plot_name <- if (!is.null(arg_plot_name) && nzchar(arg_plot_name)) {
      arg_plot_name
    } else {
      hit <- metadata_row[grepl("^\\s*Plot Name:", metadata_row, ignore.case = TRUE)]
      if (length(hit)) sub("^\\s*Plot Name:\\s*", "", hit[1], ignore.case = TRUE) else "Unknown Plot"
    }

    plot_code <- if (!is.null(arg_plot_code) && nzchar(arg_plot_code)) {
      arg_plot_code
    } else {
      hit <- metadata_row[grepl("^\\s*(Plotcode:|Plot Code:)", metadata_row, ignore.case = TRUE)]
      if (length(hit)) sub("^\\s*(Plotcode:|Plot Code:)\\s*", "", hit[1], ignore.case = TRUE) else ""
    }

    plot_census_no_fp <- if (!is.null(arg_census_no_fp) && nzchar(arg_census_no_fp)) {
      arg_census_no_fp
    } else {
      "1"
    }

    if (input_type == "field_sheet_ti") {

      raw_data <- raw[-c(1, 2), , drop = FALSE]
      n <- nrow(raw_data)

      dest_cols <- c(
        "New Tag No","New Stem Grouping","T1","T2","X","Y","Family",
        "Original determination","Morphospecies","D","POM","ExtraD","ExtraPOM",
        "Flag1","Flag2","Flag3","LI","CI","CF","CD1","nrdups","Height",
        "Voucher","Silica","Collected","Census Notes","CAP","Basal Area"
      )

      data <- as.data.frame(
        setNames(
          replicate(length(dest_cols), rep(NA, n), simplify = FALSE),
          dest_cols
        ),
        stringsAsFactors = FALSE
      )

      data$`New Tag No` <- raw_data[[1]]
      data$`New Stem Grouping` <- raw_data[[2]]
      data$T1 <- raw_data[[3]]
      data$T2 <- raw_data[[4]]
      data$X <- raw_data[[5]]
      data$Y <- raw_data[[6]]
      data$Family <- raw_data[[7]]
      data$`Original determination` <- paste0(raw_data[[9]], " | ", raw_data[[8]])
      data$Morphospecies <- raw_data[[10]]
      data$D <- raw_data[[11]]
      data$POM <- raw_data[[12]]
      data$Flag1 <- raw_data[[13]]
      data$nrdups <- raw_data[[14]]
      data$Height <- raw_data[[15]]
      data$Voucher <- raw_data[[16]]
      data$Silica <- raw_data[[17]]
      data$Collected <- raw_data[[18]]
      data$`Census Notes` <- raw_data[[19]]

      message("Detected Indigenous Land layout; skipping header remap and using direct column mapping.")

    } else {

      header_row_idx <- .find_field_header_row(raw)

      header <- raw[header_row_idx, ] |> unlist() |> as.character()
      header[is.na(header) | header == ""] <- paste0("NA_col_", seq_along(header))[is.na(header) | header == ""]

      data <- raw[-seq_len(header_row_idx), , drop = FALSE]
      colnames(data) <- make.unique(header)
      data <- data[, !is.na(colnames(data)) & colnames(data) != "", drop = FALSE]

      message("Detected header row: ", header_row_idx)
    }

    message("Detected columns: ", paste(names(data)[1:min(12, ncol(data))], collapse = " | "))
    message("DATA check:")
    print(data[1:min(20, nrow(data)), intersect(c("New Tag No", "T1", "T2", "X", "Y"), names(data))])

    if (!("Collected" %in% names(data))) data$Collected <- NA_character_
  }
  data$Collected <- gsub("^$", NA, data$Collected)
  data <- .replace_empty_with_na(data)
  if (!has_meta && input_type %in% c("field_sheet", "field_sheet_ti")) {
    message("No metadata row found; proceeding without Plot Name / Plot Code / Team.")
  }

  fp_sheet <- as.data.frame(data, stringsAsFactors = FALSE)

  # CONSOLIDATE MULTISTEMMED TREES
  # It identifies trees with multiple stems (same New Stem Grouping),
  # calculates equivalent diameter, and keeps only the main stem row.

  if (!is.null(fp_sheet) && nrow(fp_sheet) > 0) {

    # Check if we have the necessary columns
    if (all(c("New Stem Grouping", "New Tag No") %in% names(fp_sheet))) {

      # Count rows before consolidation
      n_before <- nrow(fp_sheet)

      # Count unique stem groups
      n_groups <- fp_sheet %>%
        dplyr::filter(!is.na(`New Stem Grouping`) & nzchar(`New Stem Grouping`)) %>%
        dplyr::pull(`New Stem Grouping`) %>%
        unique() %>%
        length()

      message("\n=== MULTISTEM CONSOLIDATION ===")
      message("Rows before consolidation: ", n_before)
      message("Unique stem groups found: ", n_groups)

      # Apply consolidation
      fp_sheet <- .consolidate_multistem_trees(fp_sheet, min_diameter = 5)

      # Report results
      n_after <- nrow(fp_sheet)
      message("Rows after consolidation: ", n_after)
      message("Rows removed: ", n_before - n_after)
      message("================================\n")

    } else {
      message("Note: Columns 'New Stem Grouping' and/or 'New Tag No' not found. ")
      message("      Multistem consolidation skipped.")
    }
  }

  if (!("Collected" %in% names(fp_sheet))) {
    fp_sheet$Collected <- NA_character_
  }

  if (!("New Stem Grouping" %in% names(fp_sheet))) {
    fp_sheet$`New Stem Grouping` <- NA_character_
  }

  if (!("T2" %in% names(fp_sheet))) {
    fp_sheet$T2 <- NA_character_
  }

  required_cols <- c(
    "New Tag No", "T1", "X", "Y",
    "Family", "Original determination", "D"
  )

  missing_cols <- setdiff(required_cols, names(fp_sheet))

  if (length(missing_cols)) {
    stop(
      paste0(
        "Missing required columns after parsing input sheet: ",
        paste(missing_cols, collapse = ", ")
      ),
      call. = FALSE
    )
  }

  message("Canonical columns check:")
  print(utils::head(fp_sheet[, c("New Tag No", "T1", "X", "Y")], 20))

  fp_clean <- .clean_fp_data(fp_sheet)
  cat("\nCheck X/Y:\n")
  print(summary(fp_clean$X))
  print(summary(fp_clean$Y))

  cat("\nTop pares X,Y:\n")
  print(
    fp_clean %>%
      dplyr::count(X, Y, sort = TRUE) %>%
      head(20)
  )

  if (all(c("X", "Y") %in% names(fp_clean))) {
    same_xy <- sum(fp_clean$X == fp_clean$Y, na.rm = TRUE)
    total_xy <- sum(is.finite(fp_clean$X) & is.finite(fp_clean$Y))
    n_complete_xy <- sum(is.finite(fp_clean$X) & is.finite(fp_clean$Y))

    cor_xy <- if (n_complete_xy >= 2) {
      suppressWarnings(cor(fp_clean$X, fp_clean$Y, use = "complete.obs"))
    } else {
      NA_real_
    }
    message("X summary:")
    print(summary(fp_clean$X))
    message("Y summary:")
    print(summary(fp_clean$Y))
    message("Proportion X == Y: ", round(same_xy / max(total_xy, 1), 3))
    message("Correlation X~Y: ", round(cor_xy, 3))

    if (total_xy > 0 && same_xy / total_xy > 0.5) {
      warning(
        "More than 50% of individuals have X == Y. X/Y may have been misread before plotting.",
        call. = FALSE
      )
    }
  }

  message("Computing plot coordinates...")
  if (input_type == "monitora") {
    fp_coords <- .compute_monitora_geometry(fp_clean, keep_only_cell = FALSE)
  } else {
    fp_coords <- .compute_global_coordinates(
      fp_clean = fp_clean,
      subplot_size = subplot_size,
      plot_width_m = plot_width_m,
      plot_length_m = plot_length_m
    )
  }

  fp_coords <- fp_coords %>%
    dplyr::mutate(
      diameter = (D / 100) * 2,
      `New Tag No` = trimws(as.character(`New Tag No`))
    )

  spec_df <- fp_coords %>%
    dplyr::mutate(
      Family = as.character(Family),
      Family = dplyr::if_else(is.na(Family) | Family == "", "Indet", Family),
      `Original determination` = as.character(`Original determination`),
      Species = dplyr::if_else(
        is.na(`Original determination`) | `Original determination` == "",
        "indet",
        `Original determination`
      ),
      Species_fmt = paste0("*", Species, "*")
    ) %>%
    dplyr::group_by(Family, Species_fmt) %>%
    dplyr::summarise(
      tags = .collapse_sorted_tags(`New Tag No`, sep = " | "),
      tag_vec = list(unique(trimws(as.character(`New Tag No`)))),
      .groups = "drop"
    ) %>%
    dplyr::arrange(dplyr::if_else(Family == "Indet", 0, 1), Family)

  subplot_labels <- NULL

  if (input_type %in% c("field_sheet", "field_sheet_ti", "fp_query_sheet")) {
    subplot_labels <- tibble::tibble(T1 = sort(unique(fp_coords$T1))) %>%
      dplyr::mutate(
        T1 = suppressWarnings(as.numeric(T1)),
        n_subplots_per_col = floor(plot_length_m / subplot_size),
        col_index = (T1 - 1) %/% n_subplots_per_col,
        row_index = (T1 - 1) %% n_subplots_per_col,
        subplot_y = dplyr::if_else(
          col_index %% 2 == 0,
          row_index * subplot_size,
          (n_subplots_per_col - 1 - row_index) * subplot_size
        ),
        subplot_x = col_index * subplot_size,
        center_x = subplot_x + subplot_size / 2,
        center_y = subplot_y + subplot_size / 2
      ) %>%
      dplyr::select(T1, center_x, center_y)
  }

  message("Building plot maps...")

  base_plot_interactive <- NULL

  if (input_type %in% c("field_sheet", "field_sheet_ti", "fp_query_sheet")) {
    base_plot_static <- .build_fp_base_plot(
      fp_coords = fp_coords,
      subplot_labels = subplot_labels,
      subplot_size = subplot_size,
      plot_width_m = plot_width_m,
      plot_length_m = plot_length_m,
      plot_name = plot_name,
      plot_code = plot_code,
      highlight_palms = highlight_palms,
      language = language
    )
    if (isTRUE(render_html)) {
      base_plot_interactive <- .build_fp_base_plot_interactive(
        fp_coords = fp_coords,
        subplot_labels = subplot_labels,
        subplot_size = subplot_size,
        plot_width_m = plot_width_m,
        plot_length_m = plot_length_m,
        plot_name = plot_name,
        plot_code = plot_code,
        highlight_palms = highlight_palms,
        language = language
      )
    }
  } else {
    base_plot_static <- .build_monitora_base_plot(
      fp_coords = fp_coords,
      plot_name = plot_name,
      plot_code = plot_code,
      highlight_palms = highlight_palms,
      language = language
    )
    if (isTRUE(render_html)) {
      base_plot_interactive <- .build_monitora_base_plot_interactive(
        fp_coords = fp_coords,
        plot_name = plot_name,
        plot_code = plot_code,
        highlight_palms = highlight_palms,
        language = language
      )
    } else {
      base_plot_interactive <- NULL
    }
  }

  rmd_path <- file.path(foldername, paste0(filename, "_temp.Rmd"))
  final_html <- file.path(foldername, paste0(filename, "_full_report.html"))
  final_pdf  <- file.path(foldername, paste0(filename, "_full_report.pdf"))

  collected_plot <- NULL
  uncollected_plot <- NULL
  palms_plot <- NULL

  # Use static ggplot
  tf_col <- !is.na(fp_coords$Collected)
  if (any(tf_col)) {
    collected_plot <- base_plot_static + (fp_coords %>% dplyr::filter(!is.na(Collected) & Family != "Arecaceae"))
  }

  tf_uncol <- is.na(fp_coords$Collected)
  if (any(tf_uncol)) {
    uncollected_plot <- base_plot_static + (fp_coords %>% dplyr::filter(is.na(Collected) & Family != "Arecaceae"))
  }

  tf_palm <- fp_coords$Family %in% "Arecaceae"
  if (any(tf_palm)) {
    palms_plot <- base_plot_static + (fp_coords %>% dplyr::filter(is.na(Collected) & Family == "Arecaceae"))
  }

  message("Saving PNG maps...")
  ggplot2::ggsave(
    filename = file.path(foldername, paste0(filename, "_general.png")),
    plot = base_plot_static,
    width = 14, height = 11, units = "in", dpi = 300
  )

  if (!is.null(collected_plot)) {
    ggplot2::ggsave(
      filename = file.path(foldername, paste0(filename, "_collected.png")),
      plot = collected_plot,
      width = 14, height = 11, units = "in", dpi = 300
    )
  }

  if (!is.null(uncollected_plot)) {
    ggplot2::ggsave(
      filename = file.path(foldername, paste0(filename, "_uncollected.png")),
      plot = uncollected_plot,
      width = 14, height = 11, units = "in", dpi = 300
    )
  }

  if (!is.null(palms_plot)) {
    ggplot2::ggsave(
      filename = file.path(foldername, paste0(filename, "_uncollected_palms.png")),
      plot = palms_plot,
      width = 14, height = 11, units = "in", dpi = 300
    )
  }

  message("Building subplot maps...")
  subplot_plots <- list()

  if (input_type == "monitora") {
    arms_order <- c("N", "S", "L", "O")
    arm_id <- stats::setNames(1:4, arms_order)

    full_idx <- do.call(rbind, lapply(arms_order, function(a) {
      data.frame(subunit_letter = a, T2 = 1:10, stringsAsFactors = FALSE)
    }))
    full_idx$T1 <- (arm_id[full_idx$subunit_letter] - 1L) * 10L + full_idx$T2

    .safe_breaks <- function(x) {
      x <- x[is.finite(x)]
      if (!length(x)) return(c(1, 3, 7, 10))
      rng <- range(x)
      br <- round(seq(rng[1], rng[2], length.out = 4), 0)
      br[1] <- rng[1]
      br[4] <- rng[2]
      br
    }

    for (k in seq_len(nrow(full_idx))) {
      sub <- full_idx$subunit_letter[k]
      t2 <- full_idx$T2[k]

      sp_data <- fp_coords %>%
        dplyr::filter(subunit_letter == sub, T2 == t2) %>%
        dplyr::mutate(
          Status = dplyr::case_when(
            highlight_palms & Family == "Arecaceae" ~ tr["palms"],
            !is.na(Collected) & Collected != "" ~ tr["collected"],
            TRUE ~ tr["uncollected"]
          ),
          Status = factor(Status, levels = c(tr["collected"], tr["uncollected"], tr["palms"])),
          diameter = dplyr::if_else(is.finite(D / 10), D / 10, NA_real_)
        )

      sp_plot <- .compute_monitora_geometry(sp_data, keep_only_cell = TRUE)

      brk <- .safe_breaks(sp_plot$diameter)
      brk <- unique(stats::na.omit(brk))
      if (length(brk) < 2) brk <- c(brk, brk + 0.01)
      if (anyDuplicated(brk)) brk <- unique(round(brk, 2))
      brk <- sort(brk)

      p <- ggplot2::ggplot(sp_plot, ggplot2::aes(x = x10, y = y10)) +
        ggplot2::geom_rect(ggplot2::aes(xmin = 0, xmax = 10, ymin = 0, ymax = 10),
                           fill = NA, color = "black", linewidth = 0.6) +
        ggplot2::geom_point(ggplot2::aes(size = diameter, fill = Status),
                            shape = 21, color = "black", stroke = 0.2, alpha = 0.9) +
        ggplot2::geom_text(ggplot2::aes(label = `New Tag No`), size = 1.7) +
        ggplot2::scale_fill_manual(
          values = stats::setNames(
            c("gray", "red", "gold"),
            c(tr["collected"], tr["uncollected"], tr["palms"])
          ),
          name = tr["status"]
        ) +
        ggplot2::scale_size_area(
          name = tr["dbh"],
          breaks = brk,
          labels = sprintf("%.1f", brk),
          limits = c(0, max(brk, na.rm = TRUE)),
          max_size = 10,
          guide = "legend"
        ) +
        ggplot2::labs(
          title = paste(tr["subunit"], sub, "—", tr["subplot"], t2),
          subtitle = paste(tr["plot_name"], plot_name, "|", tr["plot_code"], plot_code),
          x = tr["local_x_m"],
          y = tr["local_y_m"]
        ) +
        ggplot2::coord_fixed() +
        ggplot2::theme_bw() +
        ggplot2::theme(
          panel.grid.minor = ggplot2::element_blank(),
          panel.grid.major = ggplot2::element_line(linewidth = 0.3, color = "gray80"),
          legend.position = "right"
        ) +
        ggplot2::guides(
          fill = ggplot2::guide_legend(override.aes = list(size = 5), order = 1),
          size = ggplot2::guide_legend(title = tr["dbh"], order = 2)
        )

      #subplot_plots[[length(subplot_plots) + 1]] <- list(plot = p, data = sp_plot)
      subplot_plots[[length(subplot_plots) + 1]] <- list(
        plot = p,
        data = sp_plot,
        plotly = if (isTRUE(render_html)) .build_monitora_subplot_plot_interactive(sp_plot,
                                                                                   subplot_size = subplot_size,
                                                                                   highlight_palms = highlight_palms,
                                                                                   language = language) else NULL
      )
    }

  } else {
    all_subplot_ids <- seq_len(
      floor(plot_width_m / subplot_size) * floor(plot_length_m / subplot_size)
    )
    for (i in all_subplot_ids) {
      sp_data <- fp_coords %>%
        dplyr::filter(T1 == i) %>%
        dplyr::mutate(
          Status = dplyr::case_when(
            highlight_palms & Family == "Arecaceae" ~ tr["palms"],
            !is.na(Collected) & Collected != "" ~ tr["collected"],
            TRUE ~ tr["uncollected"]
          ),
          Status = factor(Status, levels = c(tr["collected"], tr["uncollected"], tr["palms"])),
          diameter = dplyr::if_else(is.finite(D / 10), D / 10, NA_real_)
        )

      valid_d <- sp_data$diameter[is.finite(sp_data$diameter)]
      if (length(valid_d)) {
        min_d <- min(valid_d)
        max_d <- max(valid_d)
        breaks_d <- round(seq(min_d, max_d, length.out = 4), 0)
        breaks_d[1] <- min_d
        breaks_d[4] <- max_d
      } else {
        max_d <- 10
        breaks_d <- c(1, 3, 7, 10)
      }

      d_limit <- max_d
      breaks_d <- unique(stats::na.omit(round(breaks_d, 1)))
      if (length(breaks_d) < 2) breaks_d <- c(breaks_d, breaks_d + 0.01)

      p <- ggplot2::ggplot(sp_data, ggplot2::aes(x = X, y = Y)) +
        ggplot2::annotate("rect", xmin = 0, xmax = subplot_size, ymin = 0, ymax = subplot_size,
                          fill = NA, color = "darkolivegreen", linewidth = 0.6) +
        ggplot2::geom_point(ggplot2::aes(size = diameter, fill = Status),
                            shape = 21, color = "black", stroke = 0.2, alpha = 0.9) +
        ggplot2::geom_text(ggplot2::aes(label = `New Tag No`), size = 1.7) +
        ggplot2::scale_fill_manual(
          values = stats::setNames(
            c("gray", "red", "gold"),
            c(tr["collected"], tr["uncollected"], tr["palms"])
          ),
          name = tr["status"]
        ) +
        ggplot2::scale_size_area(
          name = tr["dbh"],
          breaks = breaks_d,
          labels = sprintf("%.1f", breaks_d),
          limits = c(0, d_limit),
          max_size = 10,
          guide = "legend"
        ) +
        ggplot2::labs(
          title = paste(tr["plot_name"], plot_name),
          subtitle = paste(tr["plot_code"], plot_code),
          x = tr["x_m"],
          y = tr["y_m"]
        ) +
        ggplot2::coord_fixed() +
        ggplot2::theme_bw() +
        ggplot2::theme(
          panel.grid.minor = ggplot2::element_blank(),
          panel.grid.major = ggplot2::element_line(linewidth = 0.3, color = "gray80"),
          panel.border = ggplot2::element_rect(color = "gray80", fill = NA, linewidth = 0.3),
          legend.position = "right",
        ) +
        ggplot2::guides(
          fill = ggplot2::guide_legend(override.aes = list(size = 5), order = 1),
          size = ggplot2::guide_legend(title = tr["dbh"], order = 2)
        )

      #subplot_plots[[length(subplot_plots) + 1]] <- list(plot = p, data = sp_data)
      subplot_plots[[length(subplot_plots) + 1]] <- list(
        plot = p,
        data = sp_data,
        plotly = if (isTRUE(render_html)) .build_monitora_subplot_plot_interactive(sp_data,
                                                                                   subplot_size = subplot_size,
                                                                                   highlight_palms = highlight_palms,
                                                                                   language = language) else NULL
      )
    }
  }




  total_specimens <- nrow(fp_coords)
  collected_count <- sum(!is.na(fp_coords$Collected) & fp_coords$Family != "Arecaceae", na.rm = TRUE)
  uncollected_count <- sum(is.na(fp_coords$Collected) & fp_coords$Family != "Arecaceae", na.rm = TRUE)
  palms_count <- sum(fp_coords$Family == "Arecaceae", na.rm = TRUE)

  if (tolower(trimws(as.character(input_type))) == "monitora") {
    tag_to_subplot <- fp_coords %>%
      dplyr::transmute(
        `New Tag No` = trimws(as.character(`New Tag No`)),
        subunit_letter = as.character(subunit_letter),
        T2 = suppressWarnings(as.integer(T2)),
        subplot_index = dplyr::case_when(
          subunit_letter == "N" ~ T2,
          subunit_letter == "S" ~ 10L + T2,
          subunit_letter == "L" ~ 20L + T2,
          subunit_letter == "O" ~ 30L + T2,
          TRUE ~ NA_integer_
        )
      ) %>%
      dplyr::distinct()
  } else {
    tag_to_subplot <- fp_coords %>%
      dplyr::select(`New Tag No`, T1) %>%
      dplyr::distinct()
  }

  dashboard_obj <- .prepare_report_dashboard(
    fp_sheet = fp_coords,
    plot_size_ha = plot_size,
    subplot_size_m = subplot_size,
    language = language
  )

  rmd_content <- .create_rmd_content(
    subplot_plots = subplot_plots,
    tf_col = tf_col,
    tf_uncol = tf_uncol,
    tf_palm = tf_palm,
    plot_name = plot_name,
    plot_code = plot_code,
    spec_df = spec_df,
    dashboard = dashboard_obj,
    language = language
  )
  writeLines(c(rmd_content, ""), rmd_path)

  options(tinytex.pdflatex.args = "--no-crop")

  param_list <- list(
    language = language,
    input_type = input_type,
    metadata = list(
      plot_name = plot_name,
      plot_census_no_fp = plot_census_no_fp,
      plot_code = plot_code,
      team = team,
      census_years = if (exists("monitora_census_years", inherits = TRUE)) get("monitora_census_years") else NULL,
      census_n = if (exists("monitora_census_n", inherits = TRUE)) get("monitora_census_n") else NULL
    ),
    main_plot = base_plot_static,
    interactive_main_plot = base_plot_interactive,
    subplots_list = subplot_plots,
    subplot_size = subplot_size,
    stats = list(
      total = total_specimens,
      collected = collected_count,
      uncollected = uncollected_count,
      palms = palms_count,
      spec_df = spec_df,
      dead_since_first = if (exists("monitora_dead_since_first", inherits = TRUE)) get("monitora_dead_since_first") else NULL,
      recruits_since_first = if (exists("monitora_recruits_since_first", inherits = TRUE)) get("monitora_recruits_since_first") else NULL
    ),
    dashboard = dashboard_obj,
    tag_list = unique(trimws(as.character(fp_coords$`New Tag No`))),
    tag_to_subplot = tag_to_subplot
  )

  if (!is.null(collected_plot)) param_list$collected_plot <- collected_plot
  if (!is.null(uncollected_plot)) param_list$uncollected_plot <- uncollected_plot
  if (!is.null(palms_plot)) param_list$uncollected_palm_plot <- palms_plot

  # Build interactive plotly versions of the filtered maps for HTML
  if (isTRUE(render_html)) {

    collected_data <- fp_coords %>% dplyr::filter(!is.na(Collected) & Collected != "" & Family != "Arecaceae")
    uncollected_data <- fp_coords %>% dplyr::filter((is.na(Collected) | Collected == "") & Family != "Arecaceae")
    palm_data <- fp_coords %>% dplyr::filter((is.na(Collected) | Collected == "") & Family == "Arecaceae")

    if (input_type %in% c("field_sheet", "field_sheet_ti", "fp_query_sheet")) {

      if (nrow(collected_data) > 0) {
        param_list$interactive_collected_plot <- .build_fp_base_plot_interactive(
          fp_coords = collected_data,
          subplot_labels = subplot_labels,
          subplot_size = subplot_size,
          plot_width_m = plot_width_m,
          plot_length_m = plot_length_m,
          plot_name = plot_name,
          plot_code = plot_code,
          highlight_palms = FALSE,
          language = language
        )
      }

      if (nrow(uncollected_data) > 0) {
        param_list$interactive_uncollected_plot <- .build_fp_base_plot_interactive(
          fp_coords = uncollected_data,
          subplot_labels = subplot_labels,
          subplot_size = subplot_size,
          plot_width_m = plot_width_m,
          plot_length_m = plot_length_m,
          plot_name = plot_name,
          plot_code = plot_code,
          highlight_palms = FALSE,
          language = language
        )
      }

      if (nrow(palm_data) > 0) {
        param_list$interactive_palm_plot <- .build_fp_base_plot_interactive(
          fp_coords = palm_data,
          subplot_labels = subplot_labels,
          subplot_size = subplot_size,
          plot_width_m = plot_width_m,
          plot_length_m = plot_length_m,
          plot_name = plot_name,
          plot_code = plot_code,
          highlight_palms = TRUE,
          language = language
        )
      }

    } else if (input_type == "monitora") {

      if (nrow(collected_data) > 0) {
        param_list$interactive_collected_plot <- .build_monitora_base_plot_interactive(
          fp_coords = collected_data,
          plot_name = plot_name,
          plot_code = plot_code,
          highlight_palms = FALSE,
          language = language
        )
      }

      if (nrow(uncollected_data) > 0) {
        param_list$interactive_uncollected_plot <- .build_monitora_base_plot_interactive(
          fp_coords = uncollected_data,
          plot_name = plot_name,
          plot_code = plot_code,
          highlight_palms = FALSE,
          language = language
        )
      }

      if (nrow(palm_data) > 0) {
        param_list$interactive_palm_plot <- .build_monitora_base_plot_interactive(
          fp_coords = palm_data,
          plot_name = plot_name,
          plot_code = plot_code,
          highlight_palms = TRUE,
          language = language
        )
      }
    }
  }

  latex_engine <- "xelatex"

  html_report <- NULL
  pdf_report <- NULL
  xlsx_report <- NULL

  if (isTRUE(render_html)) {
    message("Rendering HTML report (this may take a few minutes)...")
    html_report <- .render_plot_report(
      rmd_path = rmd_path,
      output_path = foldername,
      output_name = final_html,
      params = param_list,
      format = "html",
      latex_engine = latex_engine
    )
  }

  if (isTRUE(render_pdf)) {
    message("Rendering PDF report...")
    pdf_report <- tryCatch(
      .render_plot_report(
        rmd_path = rmd_path,
        output_path = foldername,
        output_name = final_pdf,
        params = param_list,
        format = "pdf",
        latex_engine = latex_engine
      ),
      error = function(e) {
        warning("PDF generation failed: ", conditionMessage(e), call. = FALSE)
        NULL
      }
    )
  }

  if (isTRUE(write_xlsx)) {
    if (input_type != "monitora") {
      xlsx_report <- .collection_percentual(
        fp_sheet = fp_coords,
        dir = foldername,
        plot_name = plot_name,
        plot_code = plot_code,
        plot_census_no_fp = plot_census_no_fp,
        team = team,
        plot_width_m = plot_width_m,
        plot_length_m = plot_length_m,
        subplot_size = subplot_size
      )
    }
  }

  unlink(rmd_path, force = TRUE)
  unlink(gsub("_temp[.]Rmd", "_temp.knit.md", rmd_path), force = TRUE)

  invisible(list(
    html_report = html_report,
    pdf_report = pdf_report,
    xlsx_report = xlsx_report,
    output_dir = foldername
  ))
}

#### Helpers ####
.collapse_sorted_tags <- function(x, sep = "|") {
  x <- trimws(as.character(x))
  x <- x[!is.na(x) & nzchar(x)]

  if (!length(x)) return(NA_character_)

  x <- unique(x)

  x_num <- suppressWarnings(as.numeric(x))
  ord <- order(is.na(x_num), x_num, x)

  paste(x[ord], collapse = sep)
}

# Replace empty cells with NA ####
.replace_empty_with_na <- function(df) {
  df[] <- lapply(df, function(col) {
    if (is.character(col)) {
      col[col == ""] <- NA
    }
    col
  })
  df
}

# Data Clean ####
.clean_fp_data <- function(fp_sheet) {
  fp_sheet %>%
    dplyr::mutate(
      T1 = suppressWarnings(as.integer(.parse_num(T1))),
      X  = .parse_num(X),
      Y  = .parse_num(Y),
      D  = .parse_num(D),
      Collected = as.character(Collected)
    )
}

.plot_i18n <- function(language = "en") {
  language <- tolower(trimws(as.character(language)[1]))
  if (!language %in% c("en", "pt", "es", "fr", "ma", "pa")) {
    language <- "en"
  }

  dict <- list(
    en = c(
      status = "Status",
      collected = "Collected",
      uncollected = "Uncollected",
      palms = "Palms",
      dbh = "DBH (cm)",
      plot_name = "Plot Name",
      plot_code = "Plot Code",
      plot_census_no_fp = "Census No",
      team = "Team:",
      x_m = "X (m)",
      y_m = "Y (m)",
      local_x_m = "Local X (m)",
      local_y_m = "Local Y (m)",
      collection_balance = "Collection Balance",
      subplot = "Subplot",
      subunit = "Subunit",
      hover_tag = "Tag",
      hover_species = "Species",
      hover_dbh = "DBH"
    ),
    pt = c(
      status = "Status",
      collected = "Coletados",
      uncollected = "Não Coletados",
      palms = "Palmeiras",
      dbh = "DAP (cm)",
      plot_name = "Nome da Parcela",
      plot_code = "Código da Parcela",
      plot_census_no_fp = "Número do Censo",
      team = "Equipe:",
      x_m = "X (m)",
      y_m = "Y (m)",
      local_x_m = "X local (m)",
      local_y_m = "Y local (m)",
      collection_balance = "Balanço de Coleta",
      subplot = "Subparcela",
      subunit = "Subunidade",
      hover_tag = "Placa",
      hover_species = "Espécie",
      hover_dbh = "DAP"
    ),
    es = c(
      status = "Estado",
      collected = "Colectados",
      uncollected = "No Colectados",
      plot_census_no_fp = "Número de Censo",
      palms = "Palmas",
      dbh = "DAP (cm)",
      plot_name = "Nombre de la Parcela",
      plot_code = "Código de la Parcela",
      team = "Equipo",
      x_m = "X (m)",
      y_m = "Y (m)",
      local_x_m = "X local (m)",
      local_y_m = "Y local (m)",
      collection_balance = "Balance de Recolección",
      subplot = "Subparcela",
      subunit = "Subunidad",
      hover_tag = "Etiqueta",
      hover_species = "Especie",
      hover_dbh = "DAP"
    ),
    fr = c(
      status = "Statut",
      collected = "Collectés",
      uncollected = "Non Collectés",
      palms = "Palmiers",
      dbh = "DHP (cm)",
      plot_name = "Nom de la Parcelle",
      plot_code = "Code de la Parcelle",
      plot_census_no_fp = "Numéro de Recensement",
      team = "Équipe",
      x_m = "X (m)",
      y_m = "Y (m)",
      local_x_m = "X local (m)",
      local_y_m = "Y local (m)",
      collection_balance = "Bilan de Collecte",
      subplot = "Sous-parcelle",
      subunit = "Sous-unité",
      hover_tag = "Étiquette",
      hover_species = "Espèce",
      hover_dbh = "DHP"
    ),
    ma = c(
      status = "状态",
      collected = "已采集",
      uncollected = "未采集",
      palms = "棕榈科",
      dbh = "胸径 (cm)",
      plot_name = "样地名称",
      plot_code = "样地代码",
      plot_census_no_fp = "普查编号",
      team = "团队",
      x_m = "X（米）",
      y_m = "Y（米）",
      local_x_m = "局部 X（米）",
      local_y_m = "局部 Y（米）",
      collection_balance = "采集平衡",
      subplot = "子样地",
      subunit = "子单元",
      hover_tag = "标签",
      hover_species = "物种",
      hover_dbh = "胸径"
    ),
    pa = c(
      status = "Junti hẽ si mẽra",
      collected = "Pâri sonswa",
      uncollected = "Pâri rõrĩ",
      palms = "Kwatisômẽra",
      dbh = "Classes de DAP",
      plot_name = "Issi rê tâ kukâri",
      plot_code = "Kypa kukâri",
      plot_census_no_fp = "Junti hẽ rõ sên pârikran",
      team = "Sâpêrãtê",
      x_m = "X (m)",
      y_m = "Y (m)",
      local_x_m = "X local (m)",
      local_y_m = "Y local (m)",
      collection_balance = "Kukâra krepãã sââ",
      subplot = "Kukâra krepãã sââ",
      subunit = "Subunit",
      hover_tag = "Placa",
      hover_species = "Pĩrakâri jàri",
      hover_dbh = "DAP"
    )
  )

  dict[[language]]
}

# Build interactive fp base plot using plotly (client-side rendering — no SVG pre-computation)
.build_fp_base_plot_interactive <- function(fp_coords,
                                            subplot_labels,
                                            subplot_size,
                                            plot_width_m,
                                            plot_length_m,
                                            plot_name,
                                            plot_code,
                                            highlight_palms,
                                            language = "en") {
  tr <- .plot_i18n(language)

  max_x <- floor(plot_width_m / subplot_size) * subplot_size
  max_y <- floor(plot_length_m / subplot_size) * subplot_size

  # Determine status and colour per individual
  status <- dplyr::case_when(
    highlight_palms & fp_coords$Family == "Arecaceae" ~ tr["palms"],
    !is.na(fp_coords$Collected) & fp_coords$Collected != "" ~ tr["collected"],
    TRUE ~ tr["uncollected"]
  )
  color_map <- stats::setNames(c("gray80", "#EF4444", "gold"),
                               c(tr["collected"], tr["uncollected"], tr["palms"]))
  point_colors <- unname(color_map[status])

  # Marker sizes proportional to diameter (cm); fallback to minimum when NA
  diams <- fp_coords$diameter
  diams[!is.finite(diams)] <- 0
  sizes <- 2 + diams * 3
  sizes <- pmax(sizes, 6)

  # Hover tooltip text
  status_label <- dplyr::case_when(
    !is.na(fp_coords$Collected) & fp_coords$Collected != "" ~ tr["collected"],
    fp_coords$Family == "Arecaceae" ~ tr["palms"],
    TRUE ~ tr["uncollected"]
  )
  hover_text <- paste0(
    "<b>", tr["hover_tag"], ":", "</b> ", fp_coords$`New Tag No`, "<br>",
    "<b>", tr["hover_species"], ":", "</b> ", fp_coords$`Original determination`, "<br>",
    "<b>", tr["hover_dbh"], ":", "</b> ", round(fp_coords$D/10, 1), " cm<br>",
    "<b>", tr["status"], ":", "</b> ", status_label
  )

  # Sanitised tag IDs matched to checklist anchors
  tag_ids <- gsub("[^A-Za-z0-9-]", "", trimws(as.character(fp_coords$`New Tag No`)))

  # Grid lines and border as plotly shapes
  v_lines <- lapply(seq(0, max_x, by = subplot_size), function(x) {
    list(type = "line", x0 = x, x1 = x, y0 = 0, y1 = max_y,
         line = list(color = "gray", width = 0.4), layer = "below")
  })
  h_lines <- lapply(seq(0, max_y, by = subplot_size), function(y) {
    list(type = "line", x0 = 0, x1 = max_x, y0 = y, y1 = y,
         line = list(color = "gray", width = 0.4), layer = "below")
  })
  border <- list(type = "rect", x0 = 0, x1 = max_x, y0 = 0, y1 = max_y,
                 line = list(color = "darkolivegreen", width = 2.0),
                 fillcolor = "rgba(0,0,0,0)", layer = "below")
  all_shapes <- c(v_lines, h_lines, list(border))

  # Subplot number labels as plotly annotations
  annotations <- list()
  if (!is.null(subplot_labels) && nrow(subplot_labels) > 0) {
    annotations <- lapply(seq_len(nrow(subplot_labels)), function(i) {
      list(x = subplot_labels$center_x[i],
           y = subplot_labels$center_y[i],
           text = as.character(subplot_labels$T1[i]),
           showarrow = FALSE,
           font = list(size = 9, color = "lightgray"),
           xref = "x", yref = "y")
    })
  }

  # Create tick values for every subplot_size (10 meters)
  tick_vals_x <- seq(0, max_x, by = subplot_size)
  tick_vals_y <- seq(0, max_y, by = subplot_size)

  # Build the plotly scatter
  p <- plotly::plot_ly(
    x = fp_coords$global_x,
    y = fp_coords$global_y,
    type = "scatter",
    mode = "markers+text",
    height = 1000,  # Increased from 850 to 1000
    width = 1200,  # Add explicit width
    marker = list(color = point_colors, size = sizes,
                  line = list(color = "black", width = 0.5),
                  symbol = "circle"),
    text = as.character(fp_coords$`New Tag No`),
    textposition = "middle center",
    textfont = list(size = 7, color = "black"),
    hovertext = hover_text,
    hoverinfo = "text",
    customdata = tag_ids,
    showlegend = FALSE
  ) %>%
    plotly::layout(
      title = list(
        text = paste0("<b>", tr["collection_balance"], ": ", plot_name, "</b><br>",
                      "<sup>", tr["plot_code"], ": ", plot_code, "</sup>"),
        x = 0.5, xanchor = "center"
      ),
      xaxis = list(
        title = list(text = tr["x_m"], standoff = 2),
        range = c(-1, max_x + 2),
        scaleanchor = "y",
        scaleratio = 1,
        showgrid = FALSE,
        zeroline = FALSE,
        tickmode = "array",
        tickvals = tick_vals_x,
        ticktext = as.character(tick_vals_x),
        automargin = TRUE
      ),
      yaxis = list(
        title = list(text = tr["y_m"], standoff = 2),
        range = c(-1, max_y + 2),
        showgrid = FALSE,
        zeroline = FALSE,
        tickmode = "array",
        tickvals = tick_vals_y,
        ticktext = as.character(tick_vals_y),
        automargin = TRUE
      ),
      shapes = all_shapes,
      annotations = annotations,
      plot_bgcolor = "white",
      paper_bgcolor = "white",
      margin = list(l = 60, r = 60, t = 80, b = 60),  # Increased margins
      autosize = FALSE  # Change to FALSE for fixed dimensions
    ) %>%
    plotly::config(responsive = TRUE,
                   displayModeBar = TRUE,
                   scrollZoom = TRUE,
                   doubleClick = "reset",
                   displaylogo = FALSE,
                   toImageButtonOptions = list(format = "png", scale = 2)
                   ) %>%
    htmlwidgets::onRender(
      "function(el) {
         el.on('plotly_click', function(eventData) {
           if (!eventData.points || eventData.points.length === 0) return;
           var tag = eventData.points[0].customdata;
           if (!tag) return;
           var anchor = document.getElementById('checklist-tag-' + tag);
           if (anchor) {
             anchor.scrollIntoView({behavior: 'smooth', block: 'center'});
           }
         });
       }"
    )

  p
}


# Build fp base plot
.build_fp_base_plot <- function(fp_coords,
                                subplot_labels,
                                subplot_size,
                                plot_width_m,
                                plot_length_m,
                                plot_name,
                                plot_code,
                                highlight_palms,
                                language = "en") {
  tr <- .plot_i18n(language)

  status_levels <- c(tr["collected"], tr["uncollected"], tr["palms"])
  status_values <- stats::setNames(
    c("gray", "red", "gold"),
    status_levels
  )

  max_x <- floor(plot_width_m / subplot_size) * subplot_size
  max_y <- floor(plot_length_m / subplot_size) * subplot_size

  base_plot <- ggplot2::ggplot(fp_coords, ggplot2::aes(x = global_x, y = global_y)) +
    ggplot2::geom_vline(
      xintercept = seq(0, max_x, by = subplot_size),
      color = "gray80",
      linewidth = 0.3
    ) +
    ggplot2::geom_hline(
      yintercept = seq(0, max_y, by = subplot_size),
      color = "gray80",
      linewidth = 0.3
    ) +
    ggplot2::geom_rect(
      ggplot2::aes(xmin = 0, xmax = max_x, ymin = 0, ymax = max_y),
      fill = NA,
      color = "darkolivegreen",
      linewidth = 0.6
    ) +
    ggplot2::geom_text(
      data = subplot_labels,
      ggplot2::aes(x = center_x, y = center_y, label = T1),
      color = "gray",
      size = 2,
      fontface = "bold",
      inherit.aes = FALSE
    ) +
    ggplot2::geom_point(
      ggplot2::aes(
        size = diameter,
        fill = factor(
          dplyr::case_when(
            highlight_palms & Family == "Arecaceae" ~ tr["palms"],
            !is.na(Collected) & Collected != "" ~ tr["collected"],
            TRUE ~ tr["uncollected"]
          ),
          levels = status_levels
        )
      ),
      shape = 21,
      stroke = 0.2,
      color = "black",
      alpha = 0.9
    ) +
    ggplot2::geom_text(
      ggplot2::aes(label = `New Tag No`),
      vjust = 0.5,
      hjust = 0.5,
      size = 0.6
    ) +
    ggplot2::scale_x_continuous(
      limits = c(0, max_x),
      breaks = seq(0, max_x, by = subplot_size)
    ) +
    ggplot2::scale_y_continuous(
      limits = c(0, max_y),
      breaks = seq(0, max_y, by = subplot_size)
    ) +
    ggplot2::coord_fixed(ratio = 1, clip = "off") +
    ggplot2::scale_fill_manual(
      values = status_values,
      name = tr["status"]
    ) +
    ggplot2::scale_size_continuous(range = c(2, 6), guide = "none") +
    ggplot2::labs(
      x = tr["x_m"],
      y = tr["y_m"],
      title = paste0(tr["collection_balance"], ": ", plot_name),
      subtitle = paste0(tr["plot_code"], ": ", plot_code)
    ) +
    ggplot2::theme_bw() +
    ggplot2::theme(
      legend.position = "right",
      legend.justification = c(0, 1),
      legend.box.margin = ggplot2::margin(0, 5, 0, 5),
      plot.margin = ggplot2::margin(5.5, 20, 5.5, 5.5),
      panel.border = ggplot2::element_rect(color = "gray80", fill = NA, linewidth = 0.3)
    ) +
    ggplot2::guides(
      fill = ggplot2::guide_legend(override.aes = list(size = 5))
    )

  base_plot
}


# Build monitora base plot
.build_monitora_base_plot <- function(
    fp_coords,
    plot_name,
    plot_code,
    highlight_palms,
    language = "en",
    arm_offset = 18,
    pad = 4
) {
  tr <- .plot_i18n(language)

  status_levels <- c(tr["collected"], tr["uncollected"], tr["palms"])
  status_values <- stats::setNames(
    c("gray", "red", "gold"),
    status_levels
  )
  stopifnot(all(c("draw_x","draw_y","subunit_letter","D","Collected","Family") %in% names(fp_coords)))

  # ---- constants ----
  arm_length <- 50
  cell_size <- 10
  gap_title <- 3.0
  gap_sub <- 2.5
  left_inset <- 0.8

  # grid per arm
  mk_cells <- function(arm) {
    base <- expand.grid(c = 0:4, r = 0:1)
    if (arm == "N") {
      dplyr::mutate(base, arm="N",
                    xmin=-cell_size+r*cell_size, xmax=-cell_size+r*cell_size+cell_size,
                    ymin= arm_offset+c*cell_size, ymax= arm_offset+c*cell_size+cell_size)
    } else if (arm == "S") {
      dplyr::mutate(base, arm="S",
                    xmin=-cell_size+r*cell_size, xmax=-cell_size+r*cell_size+cell_size,
                    ymin=-(arm_offset+arm_length)+c*cell_size,
                    ymax=-(arm_offset+arm_length)+c*cell_size+cell_size)
    } else if (arm == "L") {
      dplyr::mutate(base, arm="L",
                    xmin= arm_offset+c*cell_size, xmax= arm_offset+c*cell_size+cell_size,
                    ymin=-cell_size+r*cell_size, ymax=-cell_size+r*cell_size+cell_size)
    } else { # O
      dplyr::mutate(base, arm="O",
                    xmin=-(arm_offset+arm_length)+c*cell_size,
                    xmax=-(arm_offset+arm_length)+c*cell_size+cell_size,
                    ymin=-cell_size+r*cell_size, ymax=-cell_size+r*cell_size+cell_size)
    }
  }
  cells <- dplyr::bind_rows(mk_cells("N"), mk_cells("S"), mk_cells("L"), mk_cells("O"))

  centers <- cells %>%
    dplyr::mutate(
      cx=(xmin+xmax)/2, cy=(ymin+ymax)/2,
      col_from_center = dplyr::case_when(arm %in% c("N","L") ~ c, arm %in% c("S","O") ~ 4L - c, TRUE ~ c),
      col_idx  = col_from_center + 1L,
      base_num = (col_idx - 1L) * 2L + 1L,
      label = ifelse(col_idx %% 2L == 1L, base_num + r, base_num + (1L - r))
    ) %>% dplyr::select(-col_from_center, -col_idx, -base_num)

  # cross & limits
  to_arm <- arm_offset
  solid_half <- 10  # 10 m each side => total 20 m
  gx_min <- -(arm_offset + arm_length) - pad
  gx_max <- (arm_offset + arm_length) + pad
  gy_min <- gx_min
  gy_max <- gx_max
  y_max_ext <- gy_max + gap_title + gap_sub + 0.5

  edge_pad_in <- max(0.6, pad/2)
  arm_labels <- data.frame(
    lab = c("N","S","L","O"),
    lx = c(0, 0, gx_max - edge_pad_in, gx_min + edge_pad_in),
    ly = c(gy_max - edge_pad_in, gy_min + edge_pad_in, 0, 0),
    hjust = c(0.5, 0.5, 0.0, 1.0),
    vjust = c(0.0, 1.0, 0.5, 0.5)
  )

  # --- mini legend (define BEFORE building the plot) --------------------
  dashed_len <- max(to_arm - solid_half, 0)
  sx <- gx_max - 130   # more left  if you decrease
  sy <- gy_min + 10   # lower      if you decrease

  mini_legend_layers <- list(
    ggplot2::geom_segment(
      x = sx, y = sy, xend = sx + dashed_len, yend = sy,
      inherit.aes = FALSE, color = "gray50", linewidth = 0.6, linetype = "22",
      show.legend = FALSE
    ),
    ggplot2::annotate(
      "text", x = sx + dashed_len/2, y = sy - 1.6, label = "40 m",
      size = 3.4, hjust = 0.5, vjust = 1, color = "gray30"
    ),
    ggplot2::geom_segment(
      x = sx, y = sy - 5.0, xend = sx + 20, yend = sy - 5.0,
      inherit.aes = FALSE, color = "gray35", linewidth = 0.7,
      show.legend = FALSE
    ),
    ggplot2::annotate(
      "text", x = sx + 10, y = sy - 6.6, label = "20 m",
      size = 3.4, hjust = 0.5, vjust = 1, color = "gray30"
    )
  )

  # build plot -----------------------------------------------------------
  ggplot2::ggplot(
    fp_coords,
    ggplot2::aes(
      x = dplyr::case_when(
        subunit_letter %in% c("N","S") ~ draw_x,
        subunit_letter == "L" ~ draw_x - 50 + arm_offset,
        subunit_letter == "O" ~ draw_x + 100 - (arm_offset + arm_length),
        TRUE ~ draw_x
      ),
      y = dplyr::case_when(
        subunit_letter %in% c("L","O") ~ draw_y,
        subunit_letter == "N" ~ draw_y - 50 + arm_offset,
        subunit_letter == "S" ~ draw_y + 100 - (arm_offset + arm_length),
        TRUE ~ draw_y
      )
    )
  ) +
    ggplot2::geom_rect(
      data = cells,
      ggplot2::aes(xmin = xmin, xmax = xmax, ymin = ymin, ymax = ymax),
      inherit.aes = FALSE, fill = NA, color = "gray70", linewidth = 0.28, show.legend = FALSE
    ) +
    ggplot2::geom_text(
      data = centers,
      ggplot2::aes(x = cx, y = cy, label = label),
      inherit.aes = FALSE, size = 3.2, fontface = "bold", color = "gray35", show.legend = FALSE
    ) +
    ggplot2::geom_segment(x = -solid_half, xend = solid_half, y = 0, yend = 0,
                          inherit.aes = FALSE, color = "gray35", linewidth = 0.7, show.legend = FALSE) +
    ggplot2::geom_segment(y = -solid_half, yend = solid_half, x = 0, xend = 0,
                          inherit.aes = FALSE, color = "gray35", linewidth = 0.7, show.legend = FALSE) +
    ggplot2::geom_segment(x = -to_arm, xend = -solid_half, y = 0, yend = 0,
                          inherit.aes = FALSE, color = "gray50", linewidth = 0.6, linetype = "22", show.legend = FALSE) +
    ggplot2::geom_segment(x = solid_half, xend = to_arm, y = 0, yend = 0,
                          inherit.aes = FALSE, color = "gray50", linewidth = 0.6, linetype = "22", show.legend = FALSE) +
    ggplot2::geom_segment(y = -to_arm, yend = -solid_half, x = 0, xend = 0,
                          inherit.aes = FALSE, color = "gray50", linewidth = 0.6, linetype = "22", show.legend = FALSE) +
    ggplot2::geom_segment(y = solid_half, yend = to_arm, x = 0, xend = 0,
                          inherit.aes = FALSE, color = "gray50", linewidth = 0.6, linetype = "22", show.legend = FALSE) +
    ggplot2::geom_point(
      ggplot2::aes(
        size = (D/10),
        fill = factor(
          dplyr::case_when(
            highlight_palms & Family == "Arecaceae" ~ tr["palms"],
            !is.na(Collected) & Collected != "" ~ tr["collected"],
            TRUE ~ tr["uncollected"]
          ),
          levels = status_levels
        )
      ),
      shape = 21, stroke = 0.2, color = "black", alpha = 0.9
    ) +
    ggplot2::geom_text(ggplot2::aes(label = `New Tag No`), size = 0.65, show.legend = FALSE) +
    ggplot2::annotate("text",
                      x = gx_min + left_inset, y = gy_max + gap_title + gap_sub,
                      label = paste0(tr["collection_balance"], ": ", plot_name),
                      size = 5, hjust = 0, vjust = 1
    ) +
    ggplot2::annotate("text",
                      x = gx_min + left_inset, y = gy_max + gap_title,
                      label = paste0(tr["plot_code"], ": ", ifelse(is.null(plot_code), "", plot_code)),
                      size = 3.8, hjust = 0, vjust = 1
    ) +
    ggplot2::scale_x_continuous(limits = c(gx_min, gx_max),
                                expand = ggplot2::expansion(add = c(6, 0))) +
    ggplot2::scale_y_continuous(limits = c(gy_min, y_max_ext),
                                expand = ggplot2::expansion(add = c(8, 0))) +  # extra bottom room for "10 m"
    ggplot2::coord_fixed(clip = "off") +
    ggplot2::theme_void() +
    ggplot2::scale_fill_manual(
      values = status_values,
      name = tr["status"]
    ) +
    ggplot2::scale_size_continuous(range = c(2,6), guide = "none") +
    ggplot2::geom_text(
      data = arm_labels,
      ggplot2::aes(x = lx, y = ly, label = lab, hjust = hjust, vjust = vjust),
      inherit.aes = FALSE, size = 9, fontface = "bold", show.legend = FALSE
    ) +
    ggplot2::theme(
      legend.position = c(0.98, 0.10),
      legend.justification = c(1, 0),
      legend.background = ggplot2::element_rect(fill = "white", colour = NA),
      legend.title = ggplot2::element_text(size = 13, face = "plain"),
      legend.text = ggplot2::element_text(size = 12),
      legend.key.size = grid::unit(1.1, "lines")
    ) +
    ggplot2::guides(fill = ggplot2::guide_legend(override.aes = list(size = 6), order = 1)) +
    # <- add the pre-built list of layers here
    mini_legend_layers
}


.build_monitora_base_plot_interactive <- function(fp_coords,
                                                  plot_name,
                                                  plot_code,
                                                  highlight_palms,
                                                  language = "en",
                                                  arm_offset = 18,
                                                  pad = 4) {

  tr <- .plot_i18n(language)

  stopifnot(all(
    c("draw_x", "draw_y", "subunit_letter", "Family", "Collected", "D", "New Tag No") %in% names(fp_coords)
  ))

  arm_length <- 50
  cell_size <- 10
  gap_title <- 3.0
  gap_sub <- 2.5
  left_inset <- 0.8

  to_arm <- arm_offset
  solid_half <- 10
  gx_min <- -(arm_offset + arm_length) - pad
  gx_max <-  (arm_offset + arm_length) + pad
  gy_min <- gx_min
  gy_max <- gx_max
  y_max_ext <- gy_max + gap_title + gap_sub + 0.5

  # Rebuild the same transformed plotting coordinates used in the static MONITORA plot
  plot_df <- fp_coords %>%
    dplyr::mutate(
      plot_x = dplyr::case_when(
        subunit_letter %in% c("N", "S") ~ draw_x,
        subunit_letter == "L" ~ draw_x - 50 + arm_offset,
        subunit_letter == "O" ~ draw_x + 100 - (arm_offset + arm_length),
        TRUE ~ draw_x
      ),
      plot_y = dplyr::case_when(
        subunit_letter %in% c("L", "O") ~ draw_y,
        subunit_letter == "N" ~ draw_y - 50 + arm_offset,
        subunit_letter == "S" ~ draw_y + 100 - (arm_offset + arm_length),
        TRUE ~ draw_y
      )
    )

  status <- dplyr::case_when(
    highlight_palms & plot_df$Family == "Arecaceae" ~ tr["palms"],
    !is.na(plot_df$Collected) & plot_df$Collected != "" ~ tr["collected"],
    TRUE ~ tr["uncollected"]
  )

  color_map <- stats::setNames(
    c("gray80", "#EF4444", "gold"),
    c(tr["collected"], tr["uncollected"], tr["palms"])
  )
  point_colors <- unname(color_map[status])

  diams <- suppressWarnings(as.numeric(plot_df$D)) / 10
  diams[!is.finite(diams)] <- 0
  sizes <- pmax(4, sqrt(diams) * 3.5)

  hover_text <- paste0(
    "<b>", tr["hover_tag"], ":</b> ", plot_df$`New Tag No`, "<br>",
    "<b>", tr["hover_species"], ":</b> ", plot_df$`Original determination`, "<br>",
    "<b>", tr["hover_dbh"], ":</b> ", round(suppressWarnings(as.numeric(plot_df$D)) / 10, 1), " cm<br>",
    "<b>", tr["status"], ":</b> ", status
  )

  tag_ids <- gsub("[^A-Za-z0-9-]", "", trimws(as.character(plot_df$`New Tag No`)))

  # Build the same arm-cell architecture used in the static plot
  mk_cells <- function(arm) {
    base <- expand.grid(c = 0:4, r = 0:1)
    if (arm == "N") {
      dplyr::mutate(
        base, arm = "N",
        xmin = -cell_size + r * cell_size,
        xmax = -cell_size + r * cell_size + cell_size,
        ymin = arm_offset + c * cell_size,
        ymax = arm_offset + c * cell_size + cell_size
      )
    } else if (arm == "S") {
      dplyr::mutate(
        base, arm = "S",
        xmin = -cell_size + r * cell_size,
        xmax = -cell_size + r * cell_size + cell_size,
        ymin = -(arm_offset + arm_length) + c * cell_size,
        ymax = -(arm_offset + arm_length) + c * cell_size + cell_size
      )
    } else if (arm == "L") {
      dplyr::mutate(
        base, arm = "L",
        xmin = arm_offset + c * cell_size,
        xmax = arm_offset + c * cell_size + cell_size,
        ymin = -cell_size + r * cell_size,
        ymax = -cell_size + r * cell_size + cell_size
      )
    } else { # O
      dplyr::mutate(
        base, arm = "O",
        xmin = -(arm_offset + arm_length) + c * cell_size,
        xmax = -(arm_offset + arm_length) + c * cell_size + cell_size,
        ymin = -cell_size + r * cell_size,
        ymax = -cell_size + r * cell_size + cell_size
      )
    }
  }

  cells <- dplyr::bind_rows(
    mk_cells("N"),
    mk_cells("S"),
    mk_cells("L"),
    mk_cells("O")
  )

  centers <- cells %>%
    dplyr::mutate(
      cx = (xmin + xmax) / 2,
      cy = (ymin + ymax) / 2,
      col_from_center = dplyr::case_when(
        arm %in% c("N", "L") ~ c,
        arm %in% c("S", "O") ~ 4L - c,
        TRUE ~ c
      ),
      col_idx = col_from_center + 1L,
      base_num = (col_idx - 1L) * 2L + 1L,
      label = ifelse(col_idx %% 2L == 1L, base_num + r, base_num + (1L - r))
    ) %>%
    dplyr::select(-col_from_center, -col_idx, -base_num)

  edge_pad_in <- max(0.6, pad / 2)
  arm_labels <- data.frame(
    lab = c("N", "S", "L", "O"),
    lx = c(0, 0, gx_max - edge_pad_in, gx_min + edge_pad_in),
    ly = c(gy_max - edge_pad_in, gy_min + edge_pad_in, 0, 0),
    stringsAsFactors = FALSE
  )

  # Plotly shapes: cell grid
  cell_shapes <- lapply(seq_len(nrow(cells)), function(i) {
    list(
      type = "rect",
      x0 = cells$xmin[i],
      x1 = cells$xmax[i],
      y0 = cells$ymin[i],
      y1 = cells$ymax[i],
      line = list(color = "gray70", width = 0.8),
      fillcolor = "rgba(0,0,0,0)",
      layer = "below"
    )
  })

  # Plotly shapes: cross and dashed connectors
  cross_shapes <- list(
    list(type = "line", x0 = -solid_half, x1 =  solid_half, y0 = 0, y1 = 0,
         line = list(color = "gray35", width = 1.2), layer = "below"),
    list(type = "line", x0 = 0, x1 = 0, y0 = -solid_half, y1 =  solid_half,
         line = list(color = "gray35", width = 1.2), layer = "below"),

    list(type = "line", x0 = -to_arm, x1 = -solid_half, y0 = 0, y1 = 0,
         line = list(color = "gray50", width = 1, dash = "dot"), layer = "below"),
    list(type = "line", x0 =  solid_half, x1 =  to_arm, y0 = 0, y1 = 0,
         line = list(color = "gray50", width = 1, dash = "dot"), layer = "below"),
    list(type = "line", x0 = 0, x1 = 0, y0 = -to_arm, y1 = -solid_half,
         line = list(color = "gray50", width = 1, dash = "dot"), layer = "below"),
    list(type = "line", x0 = 0, x1 = 0, y0 =  solid_half, y1 =  to_arm,
         line = list(color = "gray50", width = 1, dash = "dot"), layer = "below")
  )

  all_shapes <- c(cell_shapes, cross_shapes)

  # Plotly annotations: subplot numbers
  subplot_annotations <- lapply(seq_len(nrow(centers)), function(i) {
    list(
      x = centers$cx[i],
      y = centers$cy[i],
      text = as.character(centers$label[i]),
      showarrow = FALSE,
      font = list(size = 12, color = "gray35"),
      xref = "x",
      yref = "y"
    )
  })

  # Plotly annotations: N/S/L/O labels
  arm_annotations <- lapply(seq_len(nrow(arm_labels)), function(i) {
    list(
      x = arm_labels$lx[i],
      y = arm_labels$ly[i],
      text = arm_labels$lab[i],
      showarrow = FALSE,
      font = list(size = 22, color = "black"),
      xref = "x",
      yref = "y"
    )
  })

  # Plot title annotations positioned like the static figure
  title_annotations <- list(
    list(
      x = gx_min + left_inset,
      y = gy_max + gap_title + gap_sub,
      text = paste0("<b>", tr["collection_balance"], ":</b> ", plot_name),
      showarrow = FALSE,
      xanchor = "left",
      yanchor = "top",
      font = list(size = 18, color = "black"),
      xref = "x",
      yref = "y"
    ),
    list(
      x = gx_min + left_inset,
      y = gy_max + gap_title,
      text = paste0("<b>", tr["plot_code"], ":</b> ", ifelse(is.null(plot_code), "", plot_code)),
      showarrow = FALSE,
      xanchor = "left",
      yanchor = "top",
      font = list(size = 13, color = "black"),
      xref = "x",
      yref = "y"
    )
  )

  # Optional scale annotations, matching the mini-legend spirit
  dashed_len <- max(to_arm - solid_half, 0)
  sx <- gx_max - 130
  sy <- gy_min + 10

  scale_shapes <- list(
    list(type = "line", x0 = sx, x1 = sx + dashed_len, y0 = sy, y1 = sy,
         line = list(color = "gray50", width = 1, dash = "dot")),
    list(type = "line", x0 = sx, x1 = sx + 20, y0 = sy - 5, y1 = sy - 5,
         line = list(color = "gray35", width = 1.2))
  )

  scale_annotations <- list(
    list(
      x = sx + dashed_len / 2,
      y = sy - 1.6,
      text = "40 m",
      showarrow = FALSE,
      xanchor = "center",
      yanchor = "top",
      font = list(size = 11, color = "gray30"),
      xref = "x",
      yref = "y"
    ),
    list(
      x = sx + 10,
      y = sy - 6.6,
      text = "20 m",
      showarrow = FALSE,
      xanchor = "center",
      yanchor = "top",
      font = list(size = 11, color = "gray30"),
      xref = "x",
      yref = "y"
    )
  )

  all_shapes <- c(all_shapes, scale_shapes)
  all_annotations <- c(subplot_annotations, arm_annotations, title_annotations, scale_annotations)

  p <- plotly::plot_ly(
    x = plot_df$plot_x,
    y = plot_df$plot_y,
    type = "scatter",
    mode = "markers+text",
    height = 950,
    width = 1100,
    marker = list(
      color = point_colors,
      size = sizes,
      line = list(color = "black", width = 0.5),
      symbol = "circle"
    ),
    text = as.character(plot_df$`New Tag No`),
    textposition = "middle center",
    textfont = list(size = 7, color = "black"),
    hovertext = hover_text,
    hoverinfo = "text",
    customdata = tag_ids,
    showlegend = FALSE
  ) %>%
    plotly::layout(
      xaxis = list(
        title = "",
        range = c(gx_min, gx_max),
        scaleanchor = "y",
        scaleratio = 1,
        showgrid = FALSE,
        zeroline = FALSE,
        showticklabels = FALSE
      ),
      yaxis = list(
        title = "",
        range = c(gy_min, y_max_ext),
        showgrid = FALSE,
        zeroline = FALSE,
        showticklabels = FALSE
      ),
      shapes = all_shapes,
      annotations = all_annotations,
      plot_bgcolor = "white",
      paper_bgcolor = "white",
      margin = list(l = 20, r = 20, t = 20, b = 20),
      autosize = FALSE
    ) %>%
    plotly::config(
      responsive = TRUE,
      displayModeBar = TRUE,
      scrollZoom = TRUE,
      doubleClick = "reset",
      displaylogo = FALSE,
      toImageButtonOptions = list(format = "png", scale = 2)
    ) %>%
    htmlwidgets::onRender(
      "function(el) {
         el.on('plotly_click', function(eventData) {
           if (!eventData.points || eventData.points.length === 0) return;
           var tag = eventData.points[0].customdata;
           if (!tag) return;
           var anchor = document.getElementById('checklist-tag-' + tag);
           if (anchor) {
             anchor.scrollIntoView({behavior: 'smooth', block: 'center'});
           }
         });
       }"
    )

  p
}


.build_monitora_subplot_plot_interactive <- function(sp_data,
                                                     subplot_size = 10,
                                                     highlight_palms,
                                                     language = "en") {
  if (is.null(sp_data) || !nrow(sp_data)) {
    return(NULL)
  }

  tr <- .plot_i18n(language)

  use_x10 <- all(c("x10", "y10") %in% names(sp_data))
  x_col <- if (use_x10) "x10" else "X"
  y_col <- if (use_x10) "y10" else "Y"
  size_limit <- if (use_x10) 10 else subplot_size

  if (!all(c(x_col, y_col, "New Tag No", "D") %in% names(sp_data))) {
    return(NULL)
  }

  pt_status <- as.character(sp_data$Status)
  pt_colors <- dplyr::case_when(
    grepl("Collected|Coletados|Collectés|已采集|Pâri sonswa", pt_status) ~ "gray80",
    grepl("Palms|Palmeiras|Palmiers|棕榈科|Kwatis", pt_status) ~ "gold",
    TRUE ~ "#EF4444"
  )

  diams <- suppressWarnings(as.numeric(sp_data$diameter))
  diams[!is.finite(diams)] <- 0
  pt_sizes <- pmax(4, sqrt(diams) * 3.5)

  tag_ids <- gsub("[^A-Za-z0-9-]", "", trimws(as.character(sp_data$`New Tag No`)))

  status_label <- dplyr::case_when(
    highlight_palms & sp_data$Family == "Arecaceae" ~ tr["palms"],
    !is.na(sp_data$Collected) & sp_data$Collected != "" ~ tr["collected"],
    TRUE ~ tr["uncollected"]
  )
  hover_text <- paste0(
    "<b>", tr["hover_tag"], ":", "</b> ", sp_data$`New Tag No`, "<br>",
    "<b>", tr["hover_species"], ":", "</b> ", sp_data$`Original determination`, "<br>",
    "<b>", tr["hover_dbh"], ":", "</b> ", round(sp_data$D/10, 1), " cm<br>",
    "<b>", tr["status"], ":", "</b> ", status_label
  )

  plotly::plot_ly(
    x = sp_data[[x_col]],
    y = sp_data[[y_col]],
    type = "scatter",
    mode = "markers+text",
    marker = list(
      color = pt_colors,
      size = pt_sizes,
      line = list(color = "black", width = 0.8),
      symbol = "circle"
    ),
    text = as.character(sp_data$`New Tag No`),
    textposition = "middle center",
    textfont = list(size = 8, color = "black"),
    hovertext = hover_text,
    hoverinfo = "text",
    customdata = tag_ids,
    showlegend = FALSE
  ) %>%
    plotly::layout(
      xaxis = list(
        title = "X (m)",
        range = c(-0.5, size_limit + 0.5),
        scaleanchor = "y",
        scaleratio = 1,
        showgrid = TRUE,
        gridcolor = "gray90",
        zeroline = FALSE
      ),
      yaxis = list(
        title = "Y (m)",
        range = c(-0.5, size_limit + 0.5),
        showgrid = TRUE,
        gridcolor = "gray90",
        zeroline = FALSE
      ),
      shapes = list(list(
        type = "rect",
        x0 = 0, x1 = size_limit, y0 = 0, y1 = size_limit,
        line = list(color = "darkolivegreen", width = 2),
        fillcolor = "rgba(0,0,0,0)",
        layer = "below"
      )),
      plot_bgcolor = "white",
      paper_bgcolor = "white",
      margin = list(l = 30, r = 10, t = 30, b = 30)
    ) %>%
    plotly::config(
      responsive = TRUE,
      displaylogo = FALSE
    ) %>%
    htmlwidgets::onRender(
      "function(el) {
         el.on('plotly_click', function(d) {
           if (!d.points || !d.points.length) return;
           var t = d.points[0].customdata;
           if (!t) return;
           var a = document.getElementById('checklist-tag-' + t);
           if (a) a.scrollIntoView({behavior: 'smooth', block: 'center'});
         });
       }"
    )
}


# Get percentual values ####
.collection_percentual <- function(fp_sheet, dir = getwd(),
                                   plot_name = "",
                                   plot_code = "",
                                   plot_census_no_fp = "",
                                   team = "",
                                   plot_width_m,
                                   plot_length_m,
                                   subplot_size = 10) {
  if (!dir.exists(dir)) dir.create(dir, recursive = TRUE)

  output_file <- file.path(dir, "collection_balance.xlsx")

  fp_sheet <- fp_sheet %>%
    dplyr::mutate(
      Family = trimws(as.character(Family)),
      Family = dplyr::if_else(is.na(Family) | Family == "", "Indet", Family),
      Collected = as.character(Collected),
      `New Tag No` = trimws(as.character(`New Tag No`)),
      T1 = as.character(T1)
    )

  n_cols <- floor(plot_width_m / subplot_size)
  n_rows <- floor(plot_length_m / subplot_size)
  total_subplots <- n_cols * n_rows
  all_subplots <- tibble::tibble(subplot = as.character(seq_len(total_subplots)))

  resume <- fp_sheet %>%
    dplyr::group_by(subplot = as.character(T1)) %>%
    dplyr::summarise(
      total_individuals = dplyr::n(),
      total_non_arecaceae = sum(Family != "Arecaceae", na.rm = TRUE),
      collected = sum(!is.na(Collected) & Collected != "", na.rm = TRUE),
      collected_non_arecaceae = sum(!is.na(Collected) & Collected != "" & Family != "Arecaceae", na.rm = TRUE),
      uncollected_non_arecaceae = sum((is.na(Collected) | Collected == "") & Family != "Arecaceae", na.rm = TRUE),
      arecaceae_count = sum(Family == "Arecaceae", na.rm = TRUE),
      .groups = "drop"
    )

  resume <- all_subplots %>%
    dplyr::left_join(resume, by = "subplot") %>%
    tidyr::replace_na(list(
      total_individuals = 0,
      total_non_arecaceae = 0,
      collected = 0,
      collected_non_arecaceae = 0,
      uncollected_non_arecaceae = 0,
      arecaceae_count = 0
    )) %>%
    dplyr::mutate(
      collected_percentual = dplyr::if_else(
        total_individuals > 0,
        round(100 * collected / total_individuals, 1),
        NA_real_
      ),
      collected_percentual_without_arecaceae = dplyr::if_else(
        total_non_arecaceae > 0,
        round(100 * collected_non_arecaceae / total_non_arecaceae, 1),
        NA_real_
      )
    )

  resume$subplot_num <- suppressWarnings(as.numeric(resume$subplot))
  resume <- resume %>%
    dplyr::arrange(is.na(subplot_num), subplot_num, subplot) %>%
    dplyr::select(-subplot_num)

  total <- resume %>%
    dplyr::summarise(
      subplot = "TOTAL",
      total_individuals = sum(total_individuals, na.rm = TRUE),
      total_non_arecaceae = sum(total_non_arecaceae, na.rm = TRUE),
      collected = sum(collected, na.rm = TRUE),
      collected_non_arecaceae = sum(collected_non_arecaceae, na.rm = TRUE),
      uncollected_non_arecaceae = sum(uncollected_non_arecaceae, na.rm = TRUE),
      arecaceae_count = sum(arecaceae_count, na.rm = TRUE)
    ) %>%
    dplyr::mutate(
      collected_percentual = dplyr::if_else(
        total_individuals > 0,
        round(100 * collected / total_individuals, 1),
        NA_real_
      ),
      collected_percentual_without_arecaceae = dplyr::if_else(
        total_non_arecaceae > 0,
        round(100 * collected_non_arecaceae / total_non_arecaceae, 1),
        NA_real_
      )
    )

  final_resume <- dplyr::bind_rows(resume, total)

  uncollected_df <- fp_sheet %>%
    dplyr::filter((is.na(Collected) | Collected == "") & Family != "Arecaceae") %>%
    dplyr::group_by(subplot = as.character(T1)) %>%
    dplyr::summarise(
      n_uncollected = n(),
      tagno_uncollected = .collapse_sorted_tags(`New Tag No`),
      .groups = "drop"
    )

  uncollected_df$subplot_num <- suppressWarnings(as.numeric(uncollected_df$subplot))
  uncollected_df <- uncollected_df %>%
    dplyr::arrange(is.na(subplot_num), subplot_num, subplot) %>%
    dplyr::select(-subplot_num)

  total_uncollected <- uncollected_df %>%
    dplyr::summarise(
      subplot = "TOTAL",
      n_uncollected = sum(n_uncollected, na.rm = TRUE),
      tagno_uncollected = .collapse_sorted_tags(unlist(strsplit(paste(tagno_uncollected, collapse = "|"), "\\|")))
    )

  final_uncollected <- dplyr::bind_rows(uncollected_df, total_uncollected)

  collected_df <- fp_sheet %>%
    dplyr::filter(!is.na(Collected) & Collected != "") %>%
    dplyr::group_by(subplot = as.character(T1)) %>%
    dplyr::summarise(
      n_collected = n(),
      tagno_collected = .collapse_sorted_tags(`New Tag No`),
      .groups = "drop"
    )

  collected_df$subplot_num <- suppressWarnings(as.numeric(collected_df$subplot))
  collected_df <- collected_df %>%
    dplyr::arrange(is.na(subplot_num), subplot_num, subplot) %>%
    dplyr::select(-subplot_num)

  total_collected <- collected_df %>%
    dplyr::summarise(
      subplot = "TOTAL",
      n_collected = sum(n_collected, na.rm = TRUE),
      tagno_collected = .collapse_sorted_tags(unlist(strsplit(paste(tagno_collected, collapse = "|"), "\\|")))
    )

  final_collected <- dplyr::bind_rows(collected_df, total_collected)

  wb <- openxlsx::createWorkbook()

  header_string <- paste(
    "Plot Name:", plot_name,
    "| Plot Code:", plot_code,
    "| Census No:", plot_census_no_fp,
    "| Team:", team
  )

  openxlsx::addWorksheet(wb, "COLLECTION_PERCENTUAL")
  openxlsx::writeData(wb, "COLLECTION_PERCENTUAL", header_string, startRow = 1, colNames = FALSE)
  openxlsx::writeData(wb, "COLLECTION_PERCENTUAL", final_resume, startRow = 3)

  openxlsx::addWorksheet(wb, "NOT_COLLECTED")
  openxlsx::writeData(wb, "NOT_COLLECTED", header_string, startRow = 1, colNames = FALSE)
  openxlsx::writeData(wb, "NOT_COLLECTED", final_uncollected, startRow = 3)

  openxlsx::addWorksheet(wb, "COLLECTED")
  openxlsx::writeData(wb, "COLLECTED", header_string, startRow = 1, colNames = FALSE)
  openxlsx::writeData(wb, "COLLECTED", final_collected, startRow = 3)

  color <- grDevices::colorRampPalette(c("red", "yellow", "green"))(101)

  for (i in seq_len(nrow(resume))) {
    for (col_name in c("collected_percentual", "collected_percentual_without_arecaceae")) {
      valor <- final_resume[[col_name]][i]
      if (!is.na(valor)) {
        cor_hex <- color[pmax(1, pmin(101, round(valor) + 1))]
        style <- openxlsx::createStyle(
          fgFill = cor_hex,
          halign = "CENTER",
          textDecoration = "bold",
          border = "TopBottomLeftRight"
        )
        openxlsx::addStyle(
          wb,
          sheet = "COLLECTION_PERCENTUAL",
          style = style,
          rows = i + 3,
          cols = which(names(final_resume) == col_name),
          gridExpand = FALSE,
          stack = TRUE
        )
      }
    }
  }

  openxlsx::saveWorkbook(wb, file = output_file, overwrite = TRUE)

  invisible(output_file)
}


# Extract a character row from a data.frame / matrix / vector
.safe_char_row <- function(x, row = 1L) {

  if (is.null(x) || NROW(x) < row) return(character())

  if (is.data.frame(x) || is.matrix(x)) {
    v <- x[row, , drop = TRUE]
  } else if (is.atomic(x)) {
    v <- x
  } else {
    v <- x[[row]]
  }

  v <- as.character(v)
  v[is.na(v)] <- ""
  trimws(v)
}


.render_plot_report <- function(
    rmd_path,
    output_path,
    output_name,
    params,
    format = c("pdf", "html"),
    latex_engine = "pdflatex") {

  format <- match.arg(format)

  if (!dir.exists(output_path)) {
    dir.create(output_path, recursive = TRUE)
  }

  final_out <- if (grepl("[/\\\\]", output_name)) {
    output_name
  } else {
    file.path(output_path, output_name)
  }

  render_env <- new.env(parent = globalenv())

  if (identical(format, "html")) {
    rmarkdown::render(
      input = rmd_path,
      output_file = basename(final_out),
      output_dir = dirname(final_out),
      params = params,
      envir = render_env,
      clean = TRUE,
      output_format = rmarkdown::html_document(
        toc = TRUE,
        toc_depth = 2,
        number_sections = TRUE,
        self_contained = TRUE
      )
    )

    message("Full HTML report saved to: ", final_out)
    return(invisible(final_out))
  }

  wd <- getwd()
  old_logs <- list.files(
    path = wd,
    pattern = "^file[0-9a-f]+\\.log$",
    full.names = TRUE
  )

  temp_pdf <- tempfile(fileext = ".pdf")

  extra_deps <- NULL
  if (identical(latex_engine, "xelatex")) {
    extra_deps <- rmarkdown::latex_dependency("ctex", options = "fontset=fandol")
  }

  rmarkdown::render(
    input = rmd_path,
    output_file = basename(temp_pdf),
    output_dir = tempdir(),
    params = params,
    envir = render_env,
    clean = FALSE,
    output_format = rmarkdown::pdf_document(
      keep_tex = TRUE,
      latex_engine = latex_engine,
      fig_caption = TRUE,
      extra_dependencies = extra_deps
    )
  )

  preamble_lines <- c(
    "\\makeatletter",
    "\\providecommand{\\vcenteredhbox}[1]{\\begingroup",
    "  \\setbox0=\\hbox{#1}\\parbox{\\wd0}{\\box0}\\endgroup}",
    "\\renewcommand{\\@dotsep}{4}",
    "\\renewcommand{\\contentsname}{\\hspace*{-1cm}}",
    "\\makeatother"
  )

  tex_candidates <- list.files(tempdir(), pattern = "\\.tex$", full.names = TRUE)
  tex_file <- tex_candidates[which.max(file.info(tex_candidates)$mtime)]

  if (file.exists(tex_file)) {
    tex_lines <- readLines(tex_file, warn = FALSE)
    insert_idx <- grep("\\\\begin\\{document\\}", tex_lines)[1]

    if (!is.na(insert_idx)) {
      new_tex_lines <- append(tex_lines, preamble_lines, after = insert_idx - 1)
      writeLines(new_tex_lines, tex_file)
    }

    tinytex::latexmk(tex_file, engine = latex_engine)
  }

  built_pdf <- sub("\\.tex$", ".pdf", tex_file)

  if (!file.exists(built_pdf)) {
    stop("PDF was not generated from the temporary .tex file.", call. = FALSE)
  }

  file.copy(built_pdf, final_out, overwrite = TRUE)

  new_logs <- list.files(
    path = wd,
    pattern = "^file[0-9a-f]+\\.log$",
    full.names = TRUE
  )
  extra_logs <- setdiff(new_logs, old_logs)
  if (length(extra_logs)) unlink(extra_logs, force = TRUE)

  log_files <- list.files(
    path = output_path,
    pattern = "^file[0-9a-f]+\\.log$",
    full.names = TRUE
  )
  if (length(log_files)) unlink(log_files, force = TRUE)

  message("Full PDF report saved to: ", final_out)

  invisible(final_out)
}


.detect_coordinate_mode <- function(df) {
  has_local_xy <- .has_any(df, c("X")) && .has_any(df, c("Y"))
  has_std_xy <- .has_any(df, c("Standardised X", "Standardized X")) &&
    .has_any(df, c("Standardised Y", "Standardized Y"))

  has_local_t1 <- .has_any(df, c("T1", "Sub Plot T1", "SubPlotT1"))
  has_std_t1 <- .has_any(df, c("Standardised SubPlot T1", "Standardized SubPlot T1"))

  local_score <- 0L
  std_score <- 0L

  if (has_local_xy && has_local_t1) {
    local_t1 <- .best_numeric_vec(df, c("T1", "Sub Plot T1", "SubPlotT1"), default = NA_real_)
    local_x  <- .best_numeric_vec(df, c("X"), default = NA_real_)
    local_y  <- .best_numeric_vec(df, c("Y"), default = NA_real_)
    local_score <- sum(is.finite(.parse_num(local_t1)) &
                         is.finite(.parse_num(local_x)) &
                         is.finite(.parse_num(local_y)))
  }

  if (has_std_xy && has_std_t1) {
    std_t1 <- .best_numeric_vec(df, c("Standardised SubPlot T1", "Standardized SubPlot T1"), default = NA_real_)
    std_x  <- .best_numeric_vec(df, c("Standardised X", "Standardized X"), default = NA_real_)
    std_y  <- .best_numeric_vec(df, c("Standardised Y", "Standardized Y"), default = NA_real_)
    std_score <- sum(is.finite(.parse_num(std_t1)) &
                       is.finite(.parse_num(std_x)) &
                       is.finite(.parse_num(std_y)))
  }

  if (local_score > 0 || std_score > 0) {
    if (std_score > local_score) {
      return("standardised")
    } else {
      return("local")
    }
  }

  if (has_local_xy) {
    return("local_xy_only")
  }

  "unknown"
}

.find_field_header_row <- function(raw, max_check = 6) {
  max_r <- min(nrow(raw), max_check)
  targets <- c("New Tag No", "T1", "T2", "X", "Y", "Family", "D")

  scores <- vapply(seq_len(max_r), function(i) {
    row_i <- as.character(unlist(raw[i, , drop = TRUE]))
    row_i[is.na(row_i)] <- ""
    row_i <- trimws(row_i)
    sum(targets %in% row_i)
  }, numeric(1))

  best <- which.max(scores)
  if (!length(best) || scores[best] == 0) {
    return(2L)
  }
  best
}
