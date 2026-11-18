#' Generate forest plot specimen map and collection balance
#'
#' @author Giulia Ottino & Domingos Cardoso
#'
#' @description
#' Generate specimen collection-balance maps and reports from
#' \href{https://forestplots.net/}{ForestPlots.net} or
#' \href{https://www.gov.br/icmbio/pt-br/assuntos/monitoramento/programa-monitora}{MONITORA program}
#' or \href{https://robson-lima2023.github.io/page-gigaflor/}{GigaFlor network}
#' field data. The function reads plot data, standardizes coordinates and
#' specimen metadata, builds plot and subplot maps, and optionally exports
#' HTML/PDF reports and an Excel summary workbook.
#'
#' @details
#' Supported inputs include ForestPlots.net field sheets, Indigenous Land field
#' sheets, Query Library exports, MONITORA layouts and GigaFlor field sheets. Depending on the selected
#' \code{input_type}, the function:
#' \enumerate{
#'   \item validates arguments and output directories;
#'   \item reads and harmonizes the input table into a canonical internal schema;
#'   \item extracts or overrides plot metadata;
#'   \item cleans coordinates and diameter values and computes plot geometry;
#'   \item generates whole-plot and subplot-level maps, with optional palm highlighting;
#'   \item writes the field report (HTML) and optionally prints it to PDF;
#'   \item optionally writes an Excel workbook summarizing collection balance per subplot.
#' }
#'
#' Collection status is inferred from the canonical \code{Collected} column
#' after input harmonization. For some input types, this field is derived from
#' source-specific voucher or collection-status columns.
#'
#' The function is primarily called for its side effects of writing reports and
#' summary files to disk.
#'
#' @section Input types:
#' \describe{
#'   \item{\code{"field_sheet"}}{Standard ForestPlots.net field sheet.}
#'   \item{\code{"field_sheet_ti"}}{Indigenous Land field-sheet layout: rows of
#'   plot metadata, then a header that may be bilingual
#'   (\code{"Name | Translation"}). Columns are recognised by name and, failing
#'   that, by their fixed position in the template.}
#'   \item{\code{"fp_query_sheet"}}{ForestPlots.net Query Library export, converted internally to field-sheet schema.}
#'   \item{\code{"monitora"}}{MONITORA layout, converted internally and plotted using MONITORA geometry.}
#'   \item{\code{"gigaflor"}}{GigaFlor network field sheet (see \emph{GigaFlor
#'   field sheets}).}
#' }
#'
#' @section GigaFlor field sheets:
#' The \href{https://robson-lima2023.github.io/page-gigaflor/}{GigaFlor}
#' spreadsheet is read from its \code{"Data"} worksheet (the first sheet when
#' there is none), with the plot metadata (\code{"Plotcode:"},
#' \code{"Plot Name:"}, \code{"Date:"}, \code{"Team:"}) above the header.
#' It follows GigaFlor protocol 1: a 1-ha plot of 25 subplots of
#' 20 x 20 m (\code{T1}, numbered in serpentine order as in ForestPlots), each
#' split into four 10 x 10 m quadrants (\code{T2}). With
#' \code{input_type = "gigaflor"}, \code{subplot_size} defaults to 20 m.
#' Differences from a ForestPlots field sheet are handled automatically:
#' \itemize{
#'   \item \code{"Tag Number"} is the tree tag (\code{"New Tag No"}),
#'   \code{"Species"} the determination (\code{"Original determination"})
#'   and \code{"Main Stem Tag"} groups the stems of a multi-stemmed tree;
#'   \item \code{D} is in centimetres (converted to millimetres internally;
#'   taken from \code{"CAP (cm)"} / pi when empty);
#'   \item \code{X} and \code{Y} are the geographic coordinates of each tree
#'   in decimal degrees (\code{X} = longitude, \code{Y} = latitude). They are
#'   projected to metres and fitted onto the subplot grid by a rotation and a
#'   translation that bring each tree closest to the centre of its own
#'   \code{T1} subplot, which recovers the orientation of the plot on the
#'   ground. Trees are then drawn at that position inside their subplot
#'   (kept inside it when GPS error places them just outside). A warning is
#'   given when the coordinates fit the grid poorly. Trees without
#'   coordinates are drawn at the centre of their subplot.
#' }
#'
#' @section Field report:
#' The report is a single self-contained HTML file designed to be used in the
#' field, on a phone (offline) or printed. It contains a summary of the
#' collection balance with footnotes on how each number is computed, a
#' clickable grid of subplots coloured by what remains to collect, whole-plot
#' maps with the tag of every stem (all stems, with voucher, pending and, when
#' \code{highlight_palms = TRUE}, palms), a priority route, one A4 page per
#' subplot (map, tree table with stems pending collection first, and blank
#' columns for new DBH, status, voucher number and notes), a floristic and
#' phytosociological overview and a checklist appendix. The first page
#' carries the logos chosen with \code{logos}. Maps are inline SVG, so the file opens without internet. In a
#' browser the report has a search box for tags and subplots; tapping a stem on
#' a subplot map highlights its row in the table. MONITORA plots are drawn as
#' the cross of four arms of the protocol.
#'
#' The floristic and phytosociological overview follows Mueller-Dombois &
#' Ellenberg (1974), with the subplot as the sampling unit for frequency. It
#' gives richness (species, genera, families), density and basal area per
#' hectare, Shannon \eqn{H'}, Simpson \eqn{1 - D} and Pielou \eqn{J'}; the
#' families and species with the highest importance value (IVI = relative
#' density + relative frequency + relative dominance) as stacked bar charts;
#' the distribution of stems per diameter class; and the full tables of
#' phytosociological parameters of every family and species. A name used in
#' more than one family (e.g. "Indet indet") is treated as one taxon per
#' family. The same charts are also saved as PNG files.
#'
#' The PDF report is the same HTML printed by a headless Chrome, Chromium or
#' Microsoft Edge, one A4 page per subplot. The browser is found automatically
#' or through the \code{FORPLOTR_CHROME} environment variable. When none is
#' available, the PDF is skipped with a warning; the HTML report can then be
#' printed to PDF from any browser.
#'
#' @section Side effects:
#' The function creates output files on disk, including map images, an optional
#' HTML report, an optional PDF report, and an optional Excel workbook.
#'
#' @param fp_file_path Path to the input Excel file in ForestPlots.net
#' field-sheet, Query Library, or MONITORA format.
#'
#' @param language Output language. One of \code{"en"} (English),
#' \code{"pt"} (Portuguese), \code{"es"} (Spanish), \code{"fr"} (French),
#' \code{"ma"} (Mandarin), or \code{"pa"} (Panará).
#'
#' @param input_type Input layout. One of \code{"field_sheet"},
#' \code{"field_sheet_ti"}, \code{"fp_query_sheet"}, \code{"monitora"}, or
#' \code{"gigaflor"}.
#'
#' @param plot_size Total plot size in hectares. Used for \code{"field_sheet"}
#' and \code{"fp_query_sheet"} inputs. Ignored when
#' \code{input_type = "monitora"}.
#'
#' @param subplot_size Side length of each subplot in meters. Defaults to 10,
#' or to 20 (GigaFlor protocol 1) when \code{input_type = "gigaflor"} and the
#' argument is not given.
#'
#' @param plot_width_m Plot width in meters. Used for \code{"field_sheet"} and
#' \code{"fp_query_sheet"} inputs. Ignored when
#' \code{input_type = "monitora"}.
#'
#' @param plot_length_m Plot length in meters. Used for \code{"field_sheet"} and
#' \code{"fp_query_sheet"} inputs. Ignored when
#' \code{input_type = "monitora"}.
#'
#' @param highlight_palms Logical. If \code{TRUE}, Arecaceae individuals are
#' highlighted in maps and treated as a separate visual status category.
#'
#' @param station_name Optional station identifier used when
#' \code{input_type = "monitora"}. If a character vector of length greater than
#' 1 is supplied, the function recursively generates one report per station,
#' appending the station identifier to \code{filename}.
#'
#' @param plot_name Optional plot name. If provided, overrides the plot name
#' extracted from the input file metadata.
#'
#' @param plot_code Optional plot code. If provided, overrides the plot code
#' extracted from the input file metadata.
#'
#' @param plot_census_no Optional census number. If provided, overrides the
#' census number extracted from a \code{"fp_query_sheet"} input. For field-sheet
#' inputs, the function falls back to \code{"1"} when no value is available.
#'
#' @param team Optional team or PI name. If provided, overrides the team
#' extracted from the input file metadata.
#'
#' @param render_html Logical. If \code{TRUE}, writes the field report as a
#' single self-contained HTML file (see \emph{Field report}).
#'
#' @param render_pdf Logical. If \code{TRUE}, prints the field report to PDF
#' (A4) with a headless Chrome, Chromium or Edge. Skipped with a warning when
#' no such browser is found.
#'
#' @param write_xlsx Logical. If \code{TRUE}, writes the Excel summary workbook.
#'
#' @param verbose Logical. If \code{TRUE}, prints progress and diagnostic
#' messages during input harmonization, coordinate computation, map generation,
#' and report rendering.
#'
#' @param dir Output directory. A date-stamped subfolder in the format
#' \code{ddMonYYYY} is created inside this directory. The default writes to a
#' session temporary directory (\code{tempdir()}), which is deleted when the R
#' session ends; supply a path to keep the outputs.
#'
#' @param filename Basename for generated output files, without extension.
#'
#' @param logos Logos shown at the top of the HTML and PDF report. A character
#' vector that may mix logo names, presets and paths to your own image files
#' (PNG, SVG or JPEG), in the order they should appear. Logos shipped with the
#' package: \code{"forplotR"}, \code{"jbrj"} (Jardim Botânico do Rio de
#' Janeiro), \code{"ci"} (Conservação Internacional), \code{"iakio"}
#' (Associação Iakiô Panará), \code{"monitora"} (Programa MONITORA, ICMBio)
#' and \code{"forestplots"} (ForestPlots.net). Presets: \code{"default"} is
#' forplotR + JBRJ; \code{"panara"} adds CI and Iakiô; \code{"monitora"} and
#' \code{"forestplots"}, given alone, add that partner logo to the two
#' standard ones. The default \code{NULL} uses \code{"default"}, or
#' \code{"panara"} when \code{input_type = "field_sheet_ti"}; pass
#' \code{c("forplotR", "jbrj")} to keep only the two standard logos there.
#' \code{FALSE} or \code{"none"} leaves the report without logos. Logos keep
#' their proportions and share one height.
#'
#' @param header_row Row of the spreadsheet that holds the column names. The
#' default \code{NULL} finds it automatically (see [detect_header_row()]),
#' which copes with sheets that start with rows of plot metadata. Give a
#' number to force it.
#'
#' @param colname_tagno,colname_stemgroup,colname_subplot,colname_x,colname_y,colname_family,colname_species,colname_dbh,colname_voucher,colname_collected
#' Optional column names in the input file to map to the internal schema,
#' respectively: tree tag (\code{"New Tag No"}), stem group
#' (\code{"New Stem Grouping"}), subplot (\code{"T1"}; for MONITORA, the
#' subunit N/S/L/O), local X and Y coordinates (m), family, species
#' (\code{"Original determination"}), diameter in mm (\code{"D"}; in cm for
#' \code{input_type = "gigaflor"}), voucher,
#' and collection status. Use them when your file uses different column names.
#' The default \code{NULL} keeps automatic detection. Matching ignores case,
#' accents, and punctuation. A supplied column always overrides automatic
#' detection.
#'
#' @param colname_subplot_t2 MONITORA only. Column with the subplot number
#' within the subunit (\code{"N_parcela"}).
#'
#' @param colname_cap MONITORA only. Column with circumference at breast height
#' in cm, which is converted to diameter. Ignored if \code{colname_dbh} is
#' supplied.
#'
#' @param colname_station MONITORA only. Column with the station name, used to
#' filter by \code{station_name} and as the plot code.
#'
#' @param colname_census_year MONITORA only. Column with the census year or
#' date. Only the most recent census is plotted.
#'
#' @return Invisibly returns a list with paths to generated outputs:
#' \describe{
#'   \item{html_report}{Path to the rendered HTML report, or \code{NULL} if not generated.}
#'   \item{pdf_report}{Path to the PDF report, or \code{NULL} if not generated
#'     (including when no Chrome, Chromium or Edge browser is available).}
#'   \item{xlsx_report}{Path to the Excel summary workbook, or \code{NULL} if not generated.}
#'   \item{output_dir}{Directory where outputs were written.}
#' }
#' When \code{input_type = "monitora"} and \code{station_name} has length greater
#' than 1, the function recursively processes each station and invisibly returns
#' \code{TRUE}.
#'
#' @import ggplot2
#' @importFrom readxl read_excel
#' @importFrom dplyr mutate if_else group_by arrange select distinct filter bind_rows any_of case_when rename_with rename n
#' @importFrom magrittr "%>%"
#' @importFrom grDevices pdf dev.off colorRampPalette
#' @importFrom utils capture.output download.file read.csv
#' @importFrom openxlsx read.xlsx createWorkbook addWorksheet writeData createStyle addStyle saveWorkbook
#' @importFrom readr locale read_delim
#' @importFrom tibble as_tibble tibble
#' @importFrom stringr str_to_lower str_trim str_detect str_remove word
#' @importFrom tidyr replace_na
#' @importFrom grid unit
#' @importFrom writexl write_xlsx
#' @importFrom stats na.omit setNames
#'
#' @examples
#' \donttest{
#' # Write the bundled RUS01 field sheet to a temporary Excel file
#' fp <- file.path(tempdir(), "RUS01.xlsx")
#' openxlsx::write.xlsx(RUS01, fp)
#'
#' # Field report (HTML) and Excel summary; set render_pdf = TRUE to also
#' # print the report to PDF with Chrome, Chromium or Edge
#' res <- forplot_balance(
#'   fp,
#'   input_type = "field_sheet",
#'   render_pdf = FALSE,
#'   dir = tempdir(),
#'   verbose = FALSE
#' )
#' res$html_report
#' res$xlsx_report
#'
#' # MONITORA data, one station
#' fm <- file.path(tempdir(), "TUM.xlsx")
#' openxlsx::write.xlsx(TUM, fm)
#' forplot_balance(
#'   fm,
#'   input_type = "monitora",
#'   station_name = "EA-1",
#'   render_html = FALSE,
#'   render_pdf = FALSE,
#'   dir = tempdir(),
#'   verbose = FALSE
#' )
#' }
#'
#' \dontrun{
#' # GigaFlor field sheet (protocol 1: 25 subplots of 20 x 20 m, X/Y in
#' # decimal degrees, D in cm)
#' forplot_balance(
#'   "PNMT_01.xlsx",
#'   input_type = "gigaflor",
#'   dir = "path/to/output"
#' )
#'
#' # Field sheet with non-standard column names
#' forplot_balance(
#'   "my_plot.xlsx",
#'   input_type = "field_sheet",
#'   colname_tagno = "Placa",
#'   colname_subplot = "Subparcela",
#'   colname_x = "Coord X",
#'   colname_y = "Coord Y",
#'   colname_family = "Familia",
#'   colname_species = "Especie",
#'   colname_dbh = "DAP_mm",
#'   colname_collected = "Coletado",
#'   dir = "path/to/output"
#' )
#' }
#'
#' @export

forplot_balance <- function(fp_file_path = NULL,
                            language = c("en", "pt", "es", "fr", "ma", "pa"),
                            input_type = c("field_sheet", "field_sheet_ti", "fp_query_sheet", "monitora", "gigaflor"),
                            plot_size = 1,
                            subplot_size = 10,
                            plot_width_m = NULL,
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
                            verbose = TRUE,
                            dir = file.path(tempdir(), "Results_map_plot"),
                            filename = "plot_specimen",
                            logos = NULL,
                            header_row = NULL,
                            colname_tagno = NULL,
                            colname_stemgroup = NULL,
                            colname_subplot = NULL,
                            colname_subplot_t2 = NULL,
                            colname_x = NULL,
                            colname_y = NULL,
                            colname_family = NULL,
                            colname_species = NULL,
                            colname_dbh = NULL,
                            colname_cap = NULL,
                            colname_voucher = NULL,
                            colname_collected = NULL,
                            colname_station = NULL,
                            colname_census_year = NULL) {

  header_row <- .arg_check_header_row(header_row)

  input_type <- tolower(trimws(as.character(input_type)))
  input_type <- match.arg(input_type, c("field_sheet", "field_sheet_ti", "fp_query_sheet", "monitora", "gigaflor"))

  # GigaFlor protocol 1: 20 x 20 m subplots unless told otherwise
  if (input_type == "gigaflor" && missing(subplot_size)) {
    subplot_size <- .gigaflor_protocols()[["1"]]$subplot_size
  }
  .validate_subplot_size(subplot_size)

  col_map <- .build_col_map(
    colname_tagno = colname_tagno,
    colname_stemgroup = colname_stemgroup,
    colname_subplot = colname_subplot,
    colname_subplot_t2 = colname_subplot_t2,
    colname_x = colname_x,
    colname_y = colname_y,
    colname_family = colname_family,
    colname_species = colname_species,
    colname_dbh = colname_dbh,
    colname_cap = colname_cap,
    colname_voucher = colname_voucher,
    colname_collected = colname_collected,
    colname_station = colname_station,
    colname_census_year = colname_census_year
  )

  monitora_only <- c("T2", "CAP", ".station", ".census_year")
  if (input_type != "monitora" && any(names(col_map) %in% monitora_only)) {
    warning(
      "`colname_subplot_t2`, `colname_cap`, `colname_station` and ",
      "`colname_census_year` are only used when input_type = 'monitora'; ignored.",
      call. = FALSE
    )
  }

  if (!is.logical(render_html) || length(render_html) != 1 || is.na(render_html)) {
    stop("`render_html` must be TRUE or FALSE.", call. = FALSE)
  }

  if (!is.logical(render_pdf) || length(render_pdf) != 1 || is.na(render_pdf)) {
    stop("`render_pdf` must be TRUE or FALSE.", call. = FALSE)
  }

  if (!is.logical(write_xlsx) || length(write_xlsx) != 1 || is.na(write_xlsx)) {
    stop("`write_xlsx` must be TRUE or FALSE.", call. = FALSE)
  }

  plot_geometry <- NULL

  if (input_type %in% c("field_sheet", "field_sheet_ti", "fp_query_sheet", "gigaflor")) {
    plot_geometry <- .resolve_plot_geometry(
      plot_size = plot_size,
      subplot_size = subplot_size,
      plot_width_m = plot_width_m,
      plot_length_m = plot_length_m,
      auto_message = isTRUE(verbose)
    )

    plot_width_m <- plot_geometry$plot_width_m
    plot_length_m <- plot_geometry$plot_length_m
  } else {
    plot_width_m <- NULL
    plot_length_m <- NULL
  }

  language <- match.arg(tolower(trimws(as.character(language))),
                        c("en", "pt", "es", "fr", "ma", "pa"))

  # validated before anything is read or written
  report_logos <- .resolve_report_logos(logos, input_type)

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
      forplot_balance(
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
        plot_census_no = arg_census_no_fp,
        team = arg_team,
        render_html = render_html,
        render_pdf = render_pdf,
        write_xlsx = write_xlsx,
        dir = dir,
        filename = paste0(filename, "_station_", st),
        language = language,
        verbose = verbose,
        logos = logos,
        header_row = header_row,
        colname_tagno = colname_tagno,
        colname_stemgroup = colname_stemgroup,
        colname_subplot = colname_subplot,
        colname_subplot_t2 = colname_subplot_t2,
        colname_x = colname_x,
        colname_y = colname_y,
        colname_family = colname_family,
        colname_species = colname_species,
        colname_dbh = colname_dbh,
        colname_cap = colname_cap,
        colname_voucher = colname_voucher,
        colname_collected = colname_collected,
        colname_station = colname_station,
        colname_census_year = colname_census_year
      )
    }
    return(invisible(TRUE))
  }

  # Harmonize the plot input data type
  fp_loaded <- .harmonize_plot_input(
    fp_file_path = fp_file_path,
    input_type = input_type,
    station_name = station_name,
    verbose = verbose,
    col_map = col_map,
    has_user_meta = .safe_nzchar(arg_plot_name) || .safe_nzchar(arg_plot_code),
    header_row = header_row
  )

  # GigaFlor: decimal-degree X/Y -> metres within the subplot; from here on
  # the sheet is an ordinary field sheet
  if (input_type == "gigaflor") {
    fp_loaded$fp_sheet <- .gigaflor_local_xy(
      fp_loaded$fp_sheet,
      subplot_size = subplot_size,
      plot_width_m = plot_width_m,
      plot_length_m = plot_length_m,
      verbose = verbose
    )
    input_type <- "field_sheet"
  }

  balance <- .prepare_balance_data(
    fp_loaded = fp_loaded,
    input_type = input_type,
    language = language,
    plot_size = plot_size,
    subplot_size = subplot_size,
    plot_width_m = plot_width_m,
    plot_length_m = plot_length_m,
    plot_geometry = plot_geometry,
    highlight_palms = highlight_palms,
    plot_name = arg_plot_name,
    plot_code = arg_plot_code,
    plot_census_no = arg_census_no_fp,
    team = arg_team,
    verbose = verbose
  )

  fp_coords <- balance$fp_coords
  plot_name <- balance$metadata$plot_name
  plot_code <- balance$metadata$plot_code
  plot_census_no_fp <- balance$metadata$plot_census_no_fp
  team <- balance$metadata$team

  final_html <- file.path(foldername, paste0(filename, "_full_report.html"))
  final_pdf  <- file.path(foldername, paste0(filename, "_full_report.pdf"))

  # PNG exports -------------------------------------------------------------
  if (verbose) {
    message("Saving PNG maps...")
  }

  export_size <- if (!is.null(plot_geometry)) {
    .plot_export_size(
      plot_width_m = plot_geometry$plot_width_m,
      plot_length_m = plot_geometry$plot_length_m,
      max_width_in = 12,
      max_height_in = 8
    )
  } else {
    list(width = 12, height = 8)
  }

  map_pngs <- list(
    general = balance$plots$base,
    collected = balance$plots$collected,
    uncollected = balance$plots$uncollected,
    uncollected_palms = balance$plots$palms,
    priority_uncollected_species = balance$plots$priority
  )

  for (nm in names(map_pngs)) {
    if (is.null(map_pngs[[nm]])) next
    ggplot2::ggsave(
      filename = file.path(foldername, paste0(filename, "_", nm, ".png")),
      plot = map_pngs[[nm]],
      width = export_size$width, height = export_size$height, units = "in", dpi = 300
    )
  }

  ggplot2::ggsave(
    filename = file.path(foldername, paste0(filename, "_species_plot.png")),
    plot = balance$dashboard$species_plot,
    width = 10, height = 7.5, units = "in", dpi = 300, bg = "white"
  )

  ggplot2::ggsave(
    filename = file.path(foldername, paste0(filename, "_family_plot.png")),
    plot = balance$dashboard$family_plot,
    width = 10, height = 6.5, units = "in", dpi = 300, bg = "white"
  )

  ggplot2::ggsave(
    filename = file.path(foldername, paste0(filename, "_dbh_classes.png")),
    plot = balance$dashboard$dbh_plot,
    width = 9, height = 5.5, units = "in", dpi = 300, bg = "white"
  )

  # Report ------------------------------------------------------------------
  # One self-contained HTML file (inline SVG maps); the PDF is the same file
  # printed by a headless Chrome/Chromium/Edge, so no pandoc or LaTeX is needed.
  html_report <- NULL
  pdf_report <- NULL
  xlsx_report <- NULL

  if (isTRUE(render_html) || isTRUE(render_pdf)) {
    if (verbose) {
      message("Writing the field report...")
    }
    report_file <- if (isTRUE(render_html)) final_html else tempfile(fileext = ".html")
    .write_field_report(
      file = report_file,
      fp_coords = fp_coords,
      input_type = input_type,
      subplot_size = subplot_size,
      plot_width_m = plot_width_m,
      plot_length_m = plot_length_m,
      highlight_palms = highlight_palms,
      meta = balance$metadata,
      language = language,
      plot_area_ha = if (!is.null(plot_geometry)) plot_geometry$plot_width_m * plot_geometry$plot_length_m / 10000 else NULL,
      phyto = balance$dashboard$phytosoc,
      logos = report_logos
    )
    if (isTRUE(render_html)) {
      html_report <- final_html
    }
    if (isTRUE(render_pdf)) {
      pdf_report <- .html_to_pdf(report_file, final_pdf, verbose = verbose)
      if (!isTRUE(render_html)) unlink(report_file)
    }
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

  invisible(list(
    html_report = html_report,
    pdf_report = pdf_report,
    xlsx_report = xlsx_report,
    output_dir = foldername
  ))
}


# Helpers ####
.find_field_header_row <- function(raw, max_check = 6, extra_targets = NULL) {
  max_r <- min(nrow(raw), max_check)
  targets <- c("New Tag No", "T1", "T2", "X", "Y", "Family", "D")
  targets <- unique(c(targets, trimws(as.character(extra_targets))))

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

.safe_nzchar <- function(x) {
  is.character(x) && length(x) == 1L && !is.na(x) && nzchar(trimws(x))
}

.collapse_sorted_tags <- function(x, sep = "|") {
  x <- trimws(as.character(x))
  x <- x[!is.na(x) & nzchar(x)]

  if (!length(x)) return(NA_character_)

  x <- unique(x)

  x_num <- suppressWarnings(as.numeric(x))
  ord <- order(is.na(x_num), x_num, x)

  paste(x[ord], collapse = sep)
}

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

  tr <- setNames(.tr_dict_vec(dict$key, language = language), dict$key)

  status_levels <- c(tr["collected"], tr["uncollected"], tr["palms"])
  status_values <- stats::setNames(
    c("gray", "red", "gold"),
    status_levels
  )

  max_x <- floor(plot_width_m / subplot_size) * subplot_size
  max_y <- floor(plot_length_m / subplot_size) * subplot_size
  symbol_scale <- .plot_symbol_scale(plot_width_m, plot_length_m)

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
      size = max(0.7, 2 * symbol_scale),
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
      size = max(0.2, 0.6 * symbol_scale)
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
    ggplot2::scale_size_continuous(range = c(max(0.8, 2 * symbol_scale), max(2, 6 * symbol_scale)), guide = "none") +
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
    pad = 4) {

  tr <- setNames(.tr_dict_vec(dict$key, language = language), dict$key)

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
