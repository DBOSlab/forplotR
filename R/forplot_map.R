#' Plot and save a vouchered forest plot map as HTML
#'
#' @author Giulia Ottino & Domingos Cardoso
#'
#' @description
#' Generate an interactive HTML map of vouchered individuals in a forest plot
#' using Leaflet. The function supports standard ForestPlots.net layouts defined
#' by plot vertices and MONITORA layouts defined by a single central coordinate.
#' It harmonizes specimen records, converts local plot coordinates to geographic
#' coordinates, draws plot geometry and subplot boundaries, and writes a
#' shareable HTML map to disk.
#'
#' @details
#' Supported inputs include standard ForestPlots.net field sheets, Indigenous
#' Land field sheets, Query Library exports, and MONITORA layouts. Depending on
#' \code{input_type}, the function:
#' \enumerate{
#'   \item validates required arguments and reads the input table;
#'   \item harmonizes the data into a canonical internal schema, optionally
#'   using the \code{colname_*} arguments to map non-standard column names;
#'   \item extracts plot metadata such as team, plot name, and plot code;
#'   \item converts local subplot coordinates to geographic coordinates using
#'   either plot vertices or a MONITORA center point;
#'   \item builds plot and subplot overlays;
#'   \item colors specimen markers by collection status, with palms optionally
#'   shown as a separate visual category;
#'   \item optionally embeds voucher image carousels in specimen popups;
#'   \item optionally performs a local herbarium lookup and adds herbarium image
#'   links when available; and
#'   \item saves the final map as an HTML file.
#' }
#'
#' Collection status is inferred from the canonical \code{Collected} column
#' after input harmonization. For some input types, this field is derived from
#' source-specific voucher or collection-status columns.
#'
#' @section Vertex order:
#' For non-MONITORA inputs, \code{vertex_coords} must list the four plot
#' corners in this order, using the plot's local coordinate system:
#' \enumerate{
#'   \item origin, \code{(X = 0, Y = 0)};
#'   \item \code{(X = 0, Y = plot_length_m)};
#'   \item \code{(X = plot_width_m, Y = 0)};
#'   \item \code{(X = plot_width_m, Y = plot_length_m)}.
#' }
#' A warning is issued when the distances between corners do not match the
#' plot dimensions, which usually means the corners are in a different order.
#' Coordinates are used exactly as given (decimal degrees, negative for south
#' and west).
#'
#' @section Input types:
#' \describe{
#'   \item{\code{"field_sheet"}}{Standard ForestPlots.net field-sheet layout.}
#'   \item{\code{"field_sheet_ti"}}{Indigenous Land field-sheet layout with direct column mapping.}
#'   \item{\code{"fp_query_sheet"}}{ForestPlots.net Query Library export converted internally to a field-sheet-like schema.}
#'   \item{\code{"monitora"}}{MONITORA layout converted internally and plotted using MONITORA geometry.}
#' }
#'
#' @section Side effects:
#' The function writes an HTML file named
#' \code{Results_<ddMonYYYY>_<filename>.html} in \code{dir} (the working
#' directory by default). When voucher images are linked rather than embedded, they are
#' copied to a \code{<filename>_vouchers} folder next to the HTML file; keep
#' both together when sharing the map. When \code{herbaria_lookup = TRUE}, it
#' may also build or refresh a local herbarium index from DwC-A downloads and
#' may remove downloaded archives at the end of the run depending on
#' \code{keep_herbaria_downloads}.
#'
#' @param fp_file_path Character scalar. Path to the Excel file containing plot
#' data in ForestPlots.net field-sheet, Query Library, or MONITORA format.
#'
#' @param input_type Input layout. One of \code{"field_sheet"},
#' \code{"field_sheet_ti"}, \code{"fp_query_sheet"}, or \code{"monitora"}.
#'
#' @param vertex_coords Plot reference coordinates. Accepted inputs are a data
#' frame, a path to an Excel or CSV file, or a numeric vector of length 2.
#' Latitude and longitude columns are detected by name (for example
#' \code{lat}, \code{latitude}, \code{lon}, \code{long}, \code{longitude}).
#' For \code{"field_sheet"}, \code{"field_sheet_ti"}, and
#' \code{"fp_query_sheet"}, provide the four plot corners in the order
#' described in the "Vertex order" section. For \code{"monitora"}, provide a
#' single central latitude/longitude. A numeric vector of length 2 is
#' interpreted as \code{c(lat, lon)} unless named with \code{lat}/\code{latitude}
#' and \code{lon}/\code{longitude}. When several MONITORA stations are
#' requested, supply one center per station: a named list, a data frame with
#' one row per station (in the order of \code{station_name}), or a data frame
#' with a station column.
#'
#' @param plot_size Numeric scalar. Total plot size in hectares. Used for
#' non-MONITORA inputs.
#'
#' @param subplot_size Numeric scalar. Side length of each subplot in meters.
#'
#' @param plot_width_m Optional numeric scalar. Plot width in meters for
#' non-MONITORA layouts. If omitted, dimensions are inferred from
#' \code{plot_size} and \code{subplot_size}.
#'
#' @param plot_length_m Optional numeric scalar. Plot length in meters for
#' non-MONITORA layouts. If one dimension is supplied, the other is derived
#' from \code{plot_size}.
#'
#' @param voucher_imgs Character scalar or \code{NULL}. Directory containing
#' voucher image subfolders. Each subfolder must be named after a voucher
#' (matching ignores case, spaces, and punctuation, so \code{"GCO 123"} matches
#' a folder named \code{"GCO_123"}) and contain one or more image files.
#'
#' @param filename Character scalar. Basename for the output HTML file, without
#' extension.
#'
#' @param station_name Character vector or \code{NULL}. For
#' \code{input_type = "monitora"}, the station(s) to map. When more than one
#' station is supplied, one HTML map is written per station.
#'
#' @param collector Character scalar or \code{NULL}. Optional fallback collector
#' name used when compact voucher codes are provided without an explicit
#' collector string.
#'
#' @param collector_map Named character vector or \code{NULL}. Optional mapping
#' used to expand voucher prefixes into full collector names before herbarium
#' matching, for example
#' \code{c("DC" = "D. Cardoso", "PWM" = "P. W. Moonlight")}.
#'
#' @param herbaria_lookup Logical. If \code{TRUE}, the function performs a local
#' herbarium lookup and adds herbarium image links to specimen popups when
#' available.
#'
#' @param herbaria Character vector or \code{NULL}. Herbarium codes to use in
#' the herbarium lookup, for example \code{c("RB", "UPCB")}. When
#' \code{herbaria_lookup = TRUE}, this should be a non-empty vector.
#'
#' @param herbaria_force_refresh Logical. If \code{TRUE}, forces rebuilding the
#' local herbarium index and ignores any existing cached index.
#'
#' @param keep_herbaria_downloads Logical. If \code{FALSE} (default), downloaded
#' DwC-A archives used during herbarium lookup are removed at the end of the run.
#' Set to \code{TRUE} to retain them on disk.
#'
#' @param verbose Logical. If \code{TRUE}, prints detailed messages during
#' harmonization, herbarium lookup and related processing.
#'
#' @param highlight_palms Logical. If \code{TRUE}, Arecaceae individuals are
#' shown as a separate status category.
#'
#' @param plot_name Optional plot name. If provided, overrides the plot name
#' extracted from the input file metadata.
#'
#' @param plot_code Optional plot code. If provided, overrides the plot code
#' extracted from the input file metadata.
#'
#' @param team Optional team or PI name. If provided, overrides the team
#' extracted from the input file metadata.
#'
#' @param language Map interface language. One of \code{"en"} (English),
#' \code{"pt"} (Portuguese), \code{"es"} (Spanish), \code{"fr"} (French),
#' \code{"ma"} (Mandarin), or \code{"pa"} (Panará).
#'
#' @param embed_voucher_imgs Logical. If \code{TRUE}, voucher images are
#' embedded in the HTML file, so the map is a single file that can be shared
#' on its own (but may become large). If \code{FALSE} (default), images are
#' copied next to the HTML file and linked.
#'
#' @param dir Optional output directory, created if needed. If \code{NULL}
#' (default), the HTML file is written to the working directory.
#'
#' @param colname_tagno,colname_stemgroup,colname_subplot,colname_x,colname_y,colname_family,colname_species,colname_dbh,colname_voucher,colname_collected
#' Optional column names in the input file to map to the internal schema,
#' respectively: tree tag (\code{"New Tag No"}), stem group
#' (\code{"New Stem Grouping"}), subplot (\code{"T1"}; for MONITORA, the
#' subunit N/S/L/O), local X and Y coordinates (m), family, species
#' (\code{"Original determination"}), diameter in mm (\code{"D"}), voucher,
#' and collection status. The default \code{NULL} keeps automatic detection.
#' Matching ignores case, accents, and punctuation. A supplied column always
#' overrides automatic detection.
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
#' @return Invisibly returns the path to the generated HTML file, or a
#' character vector of paths when several MONITORA stations are mapped.
#'
#' @examples
#' \dontrun{
#' # Corners in order: (0,0), (0,length), (width,0), (width,length)
#' vertex_df <- data.frame(
#'   Latitude = c(-3.123, -3.122, -3.123, -3.122),
#'   Longitude = c(-60.012, -60.012, -60.011, -60.011)
#' )
#'
#' forplot_map(
#'   fp_file_path = "data/tree_data.xlsx",
#'   input_type = "field_sheet",
#'   vertex_coords = vertex_df,
#'   voucher_imgs = "voucher_imgs",
#'   filename = "plot_map",
#'   language = "pt"
#' )
#'
#' forplot_map(
#'   fp_file_path = "data/RUS_plot.xlsx",
#'   input_type = "field_sheet",
#'   vertex_coords = "data/vertices.xlsx",
#'   voucher_imgs = "voucher_imgs",
#'   embed_voucher_imgs = TRUE,
#'   collector = "G. C. Ottino",
#'   herbaria_lookup = TRUE,
#'   herbaria = "RB",
#'   filename = "rus_plot_map"
#' )
#'
#' # Field sheet with non-standard column names
#' forplot_map(
#'   fp_file_path = "data/minha_parcela.xlsx",
#'   input_type = "field_sheet",
#'   vertex_coords = vertex_df,
#'   colname_tagno = "Placa",
#'   colname_subplot = "Subparcela",
#'   colname_x = "Coord X",
#'   colname_y = "Coord Y",
#'   colname_family = "Familia",
#'   colname_species = "Especie",
#'   colname_dbh = "DAP_mm",
#'   colname_voucher = "Voucher"
#' )
#'
#' # MONITORA, one map per station
#' forplot_map(
#'   fp_file_path = "data/monitora.xlsx",
#'   input_type = "monitora",
#'   vertex_coords = list(
#'     Station1 = c(lat = -3.10, lon = -60.02),
#'     Station2 = c(lat = -3.20, lon = -60.05)
#'   ),
#'   station_name = c("Station1", "Station2"),
#'   filename = "monitora_map"
#' )
#' }
#'
#' @importFrom leaflet leaflet addProviderTiles addPolygons addCircleMarkers fitBounds
#' @importFrom leaflet addPolylines addLabelOnlyMarkers setView addControl
#' @importFrom htmltools HTML tags htmlEscape
#' @importFrom htmlwidgets onRender prependContent saveWidget
#' @importFrom geosphere destPoint bearing distGeo
#' @importFrom readxl read_excel
#' @importFrom here here
#' @importFrom magrittr "%>%"
#' @importFrom openxlsx read.xlsx
#' @importFrom data.table fread :=
#' @importFrom stats na.omit setNames
#' @importFrom DBI dbConnect dbDisconnect dbExecute dbGetQuery dbWriteTable
#' @importFrom duckdb duckdb
#'
#' @export

forplot_map <- function(fp_file_path = NULL,
                        input_type = c("field_sheet", "field_sheet_ti", "fp_query_sheet", "monitora"),
                        vertex_coords = NULL,
                        plot_size = 1,
                        subplot_size = 10,
                        plot_width_m = NULL,
                        plot_length_m = NULL,
                        voucher_imgs = NULL,
                        filename = "plot_map",
                        station_name = NULL,
                        collector = NULL,
                        collector_map = NULL,
                        herbaria_lookup = FALSE,
                        herbaria = NULL,
                        herbaria_force_refresh = FALSE,
                        keep_herbaria_downloads = FALSE,
                        verbose = TRUE,
                        highlight_palms = TRUE,
                        plot_name = NULL,
                        plot_code = NULL,
                        team = NULL,
                        language = c("en", "pt", "es", "fr", "ma", "pa"),
                        embed_voucher_imgs = FALSE,
                        dir = NULL,
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

  # Arguments ####
  input_type <- tolower(trimws(as.character(input_type)))
  input_type <- match.arg(input_type, c("field_sheet", "field_sheet_ti", "fp_query_sheet", "monitora"))

  language <- match.arg(tolower(trimws(as.character(language))),
                        c("en", "pt", "es", "fr", "ma", "pa"))

  for (flag in c("herbaria_lookup", "herbaria_force_refresh", "keep_herbaria_downloads",
                 "verbose", "highlight_palms", "embed_voucher_imgs")) {
    v <- get(flag)
    if (!is.logical(v) || length(v) != 1L || is.na(v)) {
      stop("`", flag, "` must be TRUE or FALSE.", call. = FALSE)
    }
  }

  if (!is.character(filename) || length(filename) != 1L || !nzchar(trimws(filename))) {
    stop("`filename` must be a single non-empty character string.", call. = FALSE)
  }

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

  .validate_plot_size(plot_size)
  .validate_subplot_size(subplot_size)

  if (!is.character(fp_file_path) || length(fp_file_path) != 1L || !file.exists(fp_file_path)) {
    stop("The provided 'fp_file_path' does not exist.", call. = FALSE)
  }

  if (is.null(vertex_coords)) {
    stop("`vertex_coords` must be provided.", call. = FALSE)
  }

  # One map per MONITORA station ####
  if (input_type == "monitora" && !is.null(station_name) && length(station_name) > 1L) {
    stations <- unique(trimws(as.character(station_name)))
    paths <- character(0)

    for (i in seq_along(stations)) {
      paths[i] <- forplot_map(
        fp_file_path = fp_file_path,
        input_type = "monitora",
        vertex_coords = .vertex_for_station(vertex_coords, stations, i),
        voucher_imgs = voucher_imgs,
        filename = paste0(filename, "_station_", stations[i]),
        station_name = stations[i],
        collector = collector,
        collector_map = collector_map,
        herbaria_lookup = herbaria_lookup,
        herbaria = herbaria,
        herbaria_force_refresh = herbaria_force_refresh && i == 1L,
        keep_herbaria_downloads = keep_herbaria_downloads,
        verbose = verbose,
        highlight_palms = highlight_palms,
        plot_name = plot_name,
        plot_code = plot_code,
        team = team,
        language = language,
        embed_voucher_imgs = embed_voucher_imgs,
        dir = dir,
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

    return(invisible(paths))
  }

  # Plot geometry ####
  plot_geometry <- NULL
  if (input_type != "monitora") {
    plot_geometry <- .resolve_plot_geometry(
      plot_size = plot_size,
      subplot_size = subplot_size,
      plot_width_m = plot_width_m,
      plot_length_m = plot_length_m,
      auto_message = isTRUE(verbose)
    )
    plot_width_m <- plot_geometry$plot_width_m
    plot_length_m <- plot_geometry$plot_length_m
  }

  # Output paths ####
  output_file <- paste0("Results_", format(Sys.time(), "%d%b%Y_"), filename, ".html")

  if (is.null(dir)) {
    foldername <- "."
    output_path <- output_file
  } else {
    foldername <- .arg_check_dir(dir)
    if (!dir.exists(foldername)) dir.create(foldername, recursive = TRUE)
    output_path <- file.path(foldername, output_file)
  }

  # Labels ####
  tr <- stats::setNames(.tr_dict_vec(dict$key, language = language), dict$key)
  ui <- .map_ui_labels(language)

  # Harmonize the plot input data ####
  fp_loaded <- .harmonize_plot_input(
    fp_file_path = fp_file_path,
    input_type = input_type,
    station_name = station_name,
    verbose = verbose,
    col_map = col_map,
    has_user_meta = .safe_nzchar(plot_name) || .safe_nzchar(plot_code)
  )

  fp_sheet <- fp_loaded$fp_sheet

  team <- if (.safe_nzchar(team)) trimws(team) else
    if (.safe_nzchar(fp_loaded$team)) trimws(fp_loaded$team) else ""
  plot_name <- if (.safe_nzchar(plot_name)) trimws(plot_name) else
    if (.safe_nzchar(fp_loaded$plot_name)) trimws(fp_loaded$plot_name) else "Unknown Plot"
  plot_code <- if (.safe_nzchar(plot_code)) trimws(plot_code) else
    if (.safe_nzchar(fp_loaded$plot_code)) trimws(fp_loaded$plot_code) else ""

  # Local coordinates ####
  if (input_type == "monitora") {
    fp_clean <- fp_sheet
    fp_clean$T1 <- as.character(fp_clean$T1)
    fp_clean$T2 <- as.character(fp_clean$T2)
    fp_clean$X <- .coerce_xy(fp_clean$X)
    fp_clean$Y <- .coerce_xy(fp_clean$Y)
    fp_clean$D <- .numify(fp_clean$D)
    fp_clean <- fp_clean[is.finite(fp_clean$X) & is.finite(fp_clean$Y), , drop = FALSE]

    if (!nrow(fp_clean)) {
      stop("No valid points to plot after cleaning MONITORA X/Y values.", call. = FALSE)
    }

    fp_coords <- .compute_monitora_geometry(fp_df = fp_clean, keep_only_cell = FALSE)

    if (!all(c("draw_x", "draw_y") %in% names(fp_coords))) {
      stop("MONITORA helper must add 'draw_x' and 'draw_y' columns (meters).", call. = FALSE)
    }

    fp_coords$subplot_label <- paste0(fp_coords$subunit_letter, fp_coords$T2)

  } else {
    fp_clean <- fp_sheet
    fp_clean$T1 <- .numify(fp_clean$T1)
    fp_clean$X <- .coerce_xy(fp_clean$X)
    fp_clean$Y <- .coerce_xy(fp_clean$Y)
    fp_clean$D <- .numify(fp_clean$D)

    n_before <- nrow(fp_clean)
    fp_clean <- fp_clean[is.finite(fp_clean$T1) & is.finite(fp_clean$X) & is.finite(fp_clean$Y), , drop = FALSE]

    if (!nrow(fp_clean)) {
      stop("No valid points to plot: T1/X/Y could not be parsed from the input sheet.", call. = FALSE)
    }
    if (verbose && nrow(fp_clean) < n_before) {
      message(n_before - nrow(fp_clean), " row(s) without valid T1/X/Y were not mapped.")
    }

    fp_coords <- .compute_global_coordinates(
      fp_clean = fp_clean,
      subplot_size = subplot_size,
      plot_width_m = plot_width_m,
      plot_length_m = plot_length_m
    )

    fp_coords$subplot_label <- as.character(fp_coords$T1)
  }

  # Geographic coordinates and grid ####
  vertex_coords <- .parse_vertex_coords(vertex_coords)

  grid_polys <- list()
  grid_labels <- data.frame(subplot = character(0), lat = numeric(0), lon = numeric(0))

  if (input_type == "monitora") {
    if (nrow(vertex_coords) != 1L) {
      stop("For 'monitora', provide ONE central coordinate in 'vertex_coords'.", call. = FALSE)
    }

    center <- c(vertex_coords$longitude[1], vertex_coords$latitude[1])

    ll <- .get_latlon_from_center_mat(fp_coords$draw_x, fp_coords$draw_y, center)
    fp_coords$Latitude <- ll[, "lat"]
    fp_coords$Longitude <- ll[, "lon"]

    cells <- rbind(.mk_cells("N"), .mk_cells("S"), .mk_cells("L"), .mk_cells("O"))

    col_from_center <- ifelse(cells$arm %in% c("N", "L"), cells$c, 4L - cells$c)
    col_idx <- col_from_center + 1L
    base_num <- (col_idx - 1L) * 2L + 1L
    cell_label <- ifelse(col_idx %% 2L == 1L, base_num + cells$r, base_num + (1L - cells$r))

    for (k in seq_len(nrow(cells))) {
      xs <- c(cells$xmin[k], cells$xmax[k], cells$xmax[k], cells$xmin[k], cells$xmin[k])
      ys <- c(cells$ymin[k], cells$ymin[k], cells$ymax[k], cells$ymax[k], cells$ymin[k])
      llp <- .get_latlon_from_center_mat(xs, ys, center)
      grid_polys[[k]] <- list(lng = llp[, "lon"], lat = llp[, "lat"])
    }

    lab_ll <- .get_latlon_from_center_mat(
      (cells$xmin + cells$xmax) / 2,
      (cells$ymin + cells$ymax) / 2,
      center
    )

    grid_labels <- data.frame(
      subplot = paste0(cells$arm, cell_label),
      lat = lab_ll[, "lat"],
      lon = lab_ll[, "lon"],
      stringsAsFactors = FALSE
    )

  } else {
    if (nrow(vertex_coords) != 4L) {
      stop(
        "For 'field_sheet' and 'fp_query_sheet' you must provide FOUR plot corners in ",
        "'vertex_coords' (", nrow(vertex_coords), " given).",
        call. = FALSE
      )
    }

    p1 <- c(vertex_coords$longitude[1], vertex_coords$latitude[1])
    p2 <- c(vertex_coords$longitude[2], vertex_coords$latitude[2])
    p3 <- c(vertex_coords$longitude[3], vertex_coords$latitude[3])
    p4 <- c(vertex_coords$longitude[4], vertex_coords$latitude[4])

    .check_vertex_geometry(p1, p2, p3, p4, plot_width_m, plot_length_m)

    ll <- .get_latlon_mat(fp_coords$draw_x, fp_coords$draw_y, p1, p2, p3, p4)
    fp_coords$Latitude <- ll[, "lat"]
    fp_coords$Longitude <- ll[, "lon"]

    n_cols <- plot_geometry$n_cols
    n_rows <- plot_geometry$n_rows

    lab_sub <- character(0)
    lab_lat <- numeric(0)
    lab_lon <- numeric(0)

    for (col in 0:(n_cols - 1L)) {
      for (row in 0:(n_rows - 1L)) {
        local_poly <- .get_subplot_polygon(col, row, subplot_size)
        llp <- .get_latlon_mat(local_poly[, 1], local_poly[, 2], p1, p2, p3, p4)
        grid_polys[[length(grid_polys) + 1L]] <- list(lng = llp[, "lon"], lat = llp[, "lat"])

        center_ll <- .get_latlon_mat(mean(local_poly[1:4, 1]), mean(local_poly[1:4, 2]), p1, p2, p3, p4)

        t1_label <- if (col %% 2 == 0) {
          row + col * n_rows + 1
        } else {
          (n_rows - row - 1) + col * n_rows + 1
        }

        lab_sub <- c(lab_sub, as.character(t1_label))
        lab_lat <- c(lab_lat, center_ll[1, "lat"])
        lab_lon <- c(lab_lon, center_ll[1, "lon"])
      }
    }

    grid_labels <- data.frame(subplot = lab_sub, lat = lab_lat, lon = lab_lon,
                              stringsAsFactors = FALSE)
  }

  keep <- is.finite(fp_coords$Latitude) & is.finite(fp_coords$Longitude)
  fp_coords <- fp_coords[keep, , drop = FALSE]

  if (!nrow(fp_coords)) {
    stop("No valid points to plot: Latitude/Longitude are missing after conversion.",
         call. = FALSE)
  }

  # Status ####
  for (cl in c("New Tag No", "Family", "Original determination", "Collected", "Voucher")) {
    if (!cl %in% names(fp_coords)) fp_coords[[cl]] <- NA_character_
  }

  fp_coords$Family <- trimws(as.character(fp_coords$Family))
  fp_coords$Family[is.na(fp_coords$Family) | !nzchar(fp_coords$Family)] <- "Indet"

  sp <- trimws(as.character(fp_coords[["Original determination"]]))
  sp[is.na(sp) | !nzchar(sp)] <- "indet"
  fp_coords[["Original determination"]] <- sp

  is_palm <- isTRUE(highlight_palms) & fp_coords$Family == "Arecaceae"
  is_coll <- !is.na(fp_coords$Collected) & nzchar(trimws(as.character(fp_coords$Collected)))

  fp_coords$status <- ifelse(is_palm, "palm", ifelse(is_coll, "collected", "missing"))

  # Same palette as the forplot_balance maps
  status_colors <- c(collected = "gray", missing = "red", palm = "gold")
  status_labels <- c(collected = unname(tr["collected"]),
                     missing = unname(tr["uncollected"]),
                     palm = unname(tr["palms"]))

  fp_coords$color <- unname(status_colors[fp_coords$status])

  n_total <- nrow(fp_coords)
  n_by_status <- vapply(names(status_colors), function(s) sum(fp_coords$status == s), integer(1))

  # Herbarium lookup ####
  fp_coords$herbaria_link <- NA_character_

  if (isTRUE(herbaria_lookup)) {
    if (is.null(herbaria) || !length(herbaria)) {
      warning(
        "`herbaria_lookup = TRUE` but `herbaria` is NULL or empty. Skipping herbarium lookup.",
        call. = FALSE
      )
    } else {
      herbaria_links <- tryCatch(
        .herbaria_lookup_links(
          fp_df = fp_coords,
          herbaria = herbaria,
          force_refresh = herbaria_force_refresh,
          keep_downloads = keep_herbaria_downloads,
          collector_fallback = collector,
          collector_codes = collector_map,
          verbose = verbose
        ),
        error = function(e) {
          warning(paste0("Herbarium lookup failed: ", conditionMessage(e)), call. = FALSE)
          rep(NA_character_, nrow(fp_coords))
        }
      )

      if (is.character(herbaria_links) && length(herbaria_links) == nrow(fp_coords)) {
        fp_coords$herbaria_link <- herbaria_links
      } else {
        warning("Invalid result returned by herbarium lookup. Proceeding without herbarium links.",
                call. = FALSE)
      }
    }
  }

  # Voucher images ####
  fp_coords$Voucher <- as.character(fp_coords$Voucher)
  carousel <- rep("", nrow(fp_coords))

  if (!is.null(voucher_imgs)) {
    img_index <- .voucher_image_index(voucher_imgs)

    if (length(img_index)) {
      # forplot_vouchers() names folders "<Voucher>_Tree#<tag>" (or "Tree#<tag>"
      # without voucher) for collected trees, and "<Voucher>" otherwise.
      # Try the tree-specific name first, then the bare voucher.
      vch <- as.character(fp_coords$Voucher)
      vch[is.na(vch)] <- ""
      tag <- as.character(fp_coords[["New Tag No"]])
      tag[is.na(tag)] <- ""

      key_tree <- ifelse(nzchar(tag), .norm_nm(paste0(vch, " Tree#", tag)), "")
      key_vch <- .norm_nm(vch)
      key_vch[!nzchar(trimws(vch))] <- ""

      v_key <- ifelse(nzchar(key_tree) & key_tree %in% names(img_index), key_tree,
                      ifelse(nzchar(key_vch) & key_vch %in% names(img_index), key_vch, ""))
      used_keys <- unique(v_key[nzchar(v_key)])

      if (verbose) {
        message("Voucher images matched: ", length(used_keys), " of ",
                length(img_index), " image folder(s).")
      }

      srcs <- .voucher_image_sources(
        img_index[used_keys],
        embed = embed_voucher_imgs,
        out_dir = foldername,
        asset_dirname = paste0(filename, "_vouchers")
      )

      for (k in used_keys) {
        carousel[v_key == k] <- .carousel_html(srcs[[k]], embed = embed_voucher_imgs)
      }
    }
  }

  fp_coords$has_photo <- nzchar(carousel)

  # Popups and labels ####
  fp_coords$popup <- .build_specimen_popup(
    fp_coords = fp_coords,
    carousel = carousel,
    status_labels = status_labels,
    tr = tr,
    ui = ui
  )

  fp_coords$hover <- paste0(
    fp_coords[["New Tag No"]], " \u00b7 ", fp_coords[["Original determination"]]
  )

  fp_coords$radius <- .dbh_radius(fp_coords$D)

  # Header ####
  pct_coll <- if ((n_by_status[["collected"]] + n_by_status[["missing"]]) > 0) {
    100 * n_by_status[["collected"]] / (n_by_status[["collected"]] + n_by_status[["missing"]])
  } else {
    NA_real_
  }

  logo_uri <- .forplotr_logo_uri()
  esc <- function(x) htmltools::htmlEscape(as.character(x))

  header_html <- paste0(
    "<div class='fp-head'>",
    if (nzchar(logo_uri)) {
      paste0("<a href='https://dboslab.github.io/forplotR-website/' target='_blank' rel='noopener'>",
             "<img class='fp-logo' src='", logo_uri, "' alt='forplotR'></a>")
    } else "",
    "<div class='fp-head-txt'>",
    "<div class='fp-title'>", esc(plot_name), "</div>",
    "<div class='fp-sub'>",
    if (nzchar(plot_code)) paste0(esc(tr["plot_code"]), ": ", esc(plot_code)) else "",
    if (nzchar(team)) paste0(if (nzchar(plot_code)) " &middot; " else "", esc(ui$team), ": ", esc(team)) else "",
    "</div>",
    "<div class='fp-stats'><b>", n_total, "</b> ", esc(ui$individuals),
    if (is.finite(pct_coll)) {
      paste0(" &middot; <b>", round(pct_coll), "%</b> ", esc(tolower(tr["collected"])))
    } else "",
    "</div>",
    "</div></div>"
  )

  # Legend ####
  legend_keys <- names(status_colors)
  if (!isTRUE(highlight_palms)) legend_keys <- setdiff(legend_keys, "palm")

  # Map ####
  bounds_lat <- range(c(vertex_coords$latitude, fp_coords$Latitude), na.rm = TRUE)
  bounds_lon <- range(c(vertex_coords$longitude, fp_coords$Longitude), na.rm = TRUE)

  south <- bounds_lat[1]
  north <- bounds_lat[2]
  west <- bounds_lon[1]
  east <- bounds_lon[2]

  grp_specimens <- ui$specimens
  grp_grid <- ui$grid

  tile_opts <- function(native) leaflet::providerTileOptions(maxNativeZoom = native, maxZoom = 22)

  map <- leaflet::leaflet(
    fp_coords,
    options = leaflet::leafletOptions(minZoom = 2, maxZoom = 22)
  ) %>%
    leaflet::addProviderTiles(leaflet::providers$Esri.WorldImagery, group = "Satellite",
                              options = tile_opts(18)) %>%
    leaflet::addProviderTiles(leaflet::providers$OpenStreetMap, group = "OSM Street",
                              options = tile_opts(19)) %>%
    leaflet::addProviderTiles(leaflet::providers$Esri.WorldStreetMap, group = "ESRI Street",
                              options = tile_opts(18)) %>%
    leaflet::addProviderTiles(leaflet::providers$OpenTopoMap, group = "OpenTopo",
                              options = tile_opts(17)) %>%
    leaflet::addControl(html = header_html, position = "topleft", className = "fp-head-ctrl") %>%
    leaflet::addEasyButton(
      leaflet::easyButton(
        icon = "<span class='fp-ico'>&#9678;</span>",
        title = ui$back_to_plot,
        onClick = leaflet::JS(
          sprintf(
            "function(btn, map){ map.fitBounds([[%.8f, %.8f], [%.8f, %.8f]], {padding:[25,25]}); }",
            south, west, north, east
          )
        )
      )
    )

  if (input_type != "monitora") {
    map <- map %>%
      leaflet::addPolygons(
        lng = c(p1[1], p2[1], p4[1], p3[1], p1[1]),
        lat = c(p1[2], p2[2], p4[2], p3[2], p1[2]),
        color = "#2E7D32",
        weight = 2,
        fillColor = "#A5D6A7",
        fillOpacity = 0.12,
        options = leaflet::pathOptions(interactive = FALSE)
      )
  } else {
    lat0 <- vertex_coords$latitude[1]
    lon0 <- vertex_coords$longitude[1]
    arm <- 100

    eo <- .get_latlon_from_center_mat(c(-arm, arm), c(0, 0), c(lon0, lat0))
    ns <- .get_latlon_from_center_mat(c(0, 0), c(-arm, arm), c(lon0, lat0))

    map <- map %>%
      leaflet::addPolylines(
        lng = eo[, "lon"], lat = eo[, "lat"],
        color = "#FFFFFF", weight = 2, opacity = 0.8, dashArray = "6 6",
        options = leaflet::pathOptions(interactive = FALSE)
      ) %>%
      leaflet::addPolylines(
        lng = ns[, "lon"], lat = ns[, "lat"],
        color = "#FFFFFF", weight = 2, opacity = 0.8, dashArray = "6 6",
        options = leaflet::pathOptions(interactive = FALSE)
      )
  }

  for (i in seq_along(grid_polys)) {
    map <- map %>%
      leaflet::addPolygons(
        lng = grid_polys[[i]]$lng,
        lat = grid_polys[[i]]$lat,
        color = "#FFFFFF",
        weight = 0.6,
        opacity = 0.7,
        fillOpacity = 0,
        group = grp_grid,
        options = leaflet::pathOptions(interactive = FALSE)
      )
  }

  map <- map %>%
    leaflet::addLabelOnlyMarkers(
      data = grid_labels,
      lng = ~lon, lat = ~lat,
      label = ~as.character(subplot),
      labelOptions = leaflet::labelOptions(
        noHide = TRUE,
        direction = "center",
        textOnly = TRUE,
        className = "fp-grid-label"
      ),
      group = grp_grid
    ) %>%
    leaflet::addCircleMarkers(
      lng = ~Longitude,
      lat = ~Latitude,
      radius = ~radius,
      color = "#1F2937",
      weight = 0.8,
      fillColor = ~color,
      fillOpacity = 0.9,
      popup = ~popup,
      popupOptions = leaflet::popupOptions(maxWidth = 320, minWidth = 240, autoPanPadding = c(40, 40)),
      label = ~hover,
      labelOptions = leaflet::labelOptions(direction = "top", offset = c(0, -6), className = "fp-hover"),
      group = grp_specimens
    ) %>%
    leaflet::addLegend(
      position = "bottomright",
      colors = unname(status_colors[legend_keys]),
      labels = paste0(htmltools::htmlEscape(status_labels[legend_keys]),
                      " (", n_by_status[legend_keys], ")"),
      title = htmltools::htmlEscape(tr["status"]),
      opacity = 0.95,
      className = "info legend fp-legend"
    ) %>%
    leaflet::addScaleBar(
      position = "bottomleft",
      options = leaflet::scaleBarOptions(imperial = FALSE)
    ) %>%
    leaflet::addMeasure(
      position = "topleft",
      primaryLengthUnit = "meters",
      primaryAreaUnit = "sqmeters",
      activeColor = "#2E7D32",
      completedColor = "#1B5E20",
      localization = switch(language, pt = "pt_BR", es = "es", fr = "fr", "en")
    ) %>%
    leaflet::addLayersControl(
      baseGroups = c("Satellite", "OSM Street", "ESRI Street", "OpenTopo"),
      overlayGroups = c(grp_specimens, grp_grid),
      position = "bottomleft",
      options = leaflet::layersControlOptions(collapsed = TRUE, autoZIndex = TRUE)
    ) %>%
    leaflet::fitBounds(lng1 = west, lat1 = south, lng2 = east, lat2 = north) %>%
    htmlwidgets::prependContent(htmltools::tags$style(htmltools::HTML(.forplot_map_css()))) %>%
    htmlwidgets::onRender(
      .forplot_map_js(),
      data = list(
        labels = ui,
        statusLabels = as.list(status_labels),
        statusColors = as.list(status_colors),
        statusOrder = legend_keys,
        specimensGroup = grp_specimens,
        subplotLabel = unname(tr["subplot"]),
        familyLabel = unname(tr["hover_family"]),
        speciesLabel = unname(tr["hover_species"]),
        statusLabel = unname(tr["status"])
      )
    )

  message("Saving HTML map: '", output_path, "'")
  htmlwidgets::saveWidget(
    map,
    file = normalizePath(output_path, winslash = "/", mustWork = FALSE),
    selfcontained = TRUE,
    title = plot_name
  )

  invisible(output_path)
}


# Helpers ####

#' Map-specific interface labels
#'
#' Labels already present in the package dictionary (status, plot code, tag,
#' species, family, DBH, subplot) are taken from it; the strings below are the
#' ones only the HTML map needs. Mandarin and Panara fall back to English
#' until they are added to the dictionary.
#'
#' @keywords internal
#' @noRd
.map_ui_labels <- function(language = "en") {
  lab <- list(
    en = list(
      team = "Team", individuals = "individuals", voucher = "Voucher",
      unvouchered = "Unvouchered", herbarium = "Herbarium image",
      photo = "Has Photo", photos_only = "With photos only",
      search = "Search family, species, tag or voucher",
      filters = "Filters", showing = "Showing", of = "of",
      reset = "Reset filters", all = "All", none = "None",
      back_to_plot = "Back to plot", specimens = "Specimens",
      grid = "Subplot Grid", tag = "Tag", dbh = "DBH (cm)",
      no_results = "No individuals match the filters."
    ),
    pt = list(
      team = "Equipe", individuals = "indiv\u00edduos", voucher = "Voucher",
      unvouchered = "Sem voucher", herbarium = "Imagem de herb\u00e1rio",
      photo = "Foto", photos_only = "Somente com fotos",
      search = "Buscar fam\u00edlia, esp\u00e9cie, placa ou voucher",
      filters = "Filtros", showing = "Mostrando", of = "de",
      reset = "Limpar filtros", all = "Todos", none = "Nenhum",
      back_to_plot = "Voltar \u00e0 parcela", specimens = "Indiv\u00edduos",
      grid = "Grade de subparcelas", tag = "Placa", dbh = "DAP (cm)",
      no_results = "Nenhum indiv\u00edduo corresponde aos filtros."
    ),
    es = list(
      team = "Equipo", individuals = "individuos", voucher = "Voucher",
      unvouchered = "Sin voucher", herbarium = "Imagen de herbario",
      photo = "Foto", photos_only = "Solo con fotos",
      search = "Buscar familia, especie, placa o voucher",
      filters = "Filtros", showing = "Mostrando", of = "de",
      reset = "Restablecer filtros", all = "Todos", none = "Ninguno",
      back_to_plot = "Volver a la parcela", specimens = "Individuos",
      grid = "Cuadr\u00edcula de subparcelas", tag = "Placa", dbh = "DAP (cm)",
      no_results = "Ning\u00fan individuo coincide con los filtros."
    ),
    fr = list(
      team = "\u00c9quipe", individuals = "individus", voucher = "Voucher",
      unvouchered = "Sans voucher", herbarium = "Image d\u2019herbier",
      photo = "Photo", photos_only = "Avec photos uniquement",
      search = "Rechercher famille, esp\u00e8ce, \u00e9tiquette ou voucher",
      filters = "Filtres", showing = "Affichage", of = "sur",
      reset = "R\u00e9initialiser", all = "Tous", none = "Aucun",
      back_to_plot = "Retour \u00e0 la parcelle", specimens = "Individus",
      grid = "Grille des sous-parcelles", tag = "\u00c9tiquette", dbh = "DHP (cm)",
      no_results = "Aucun individu ne correspond aux filtres."
    )
  )

  if (language %in% names(lab)) lab[[language]] else lab$en
}


#' Parse vertex coordinates into a latitude/longitude data frame
#'
#' @keywords internal
#' @noRd
.parse_vertex_coords <- function(vertex_coords) {
  vc <- vertex_coords

  if (is.character(vc) && length(vc) == 1L) {
    vc <- if (grepl("\\.xlsx?$", vc, ignore.case = TRUE)) {
      as.data.frame(readxl::read_excel(vc), stringsAsFactors = FALSE)
    } else if (grepl("\\.csv$", vc, ignore.case = TRUE)) {
      utils::read.csv(vc, stringsAsFactors = FALSE, check.names = FALSE)
    } else {
      stop("`vertex_coords` file must be .xlsx, .xls or .csv.", call. = FALSE)
    }
  }

  if (!is.data.frame(vc)) {
    if (is.atomic(vc) && length(vc) == 2L) {
      nm <- tolower(names(vc))
      if (length(nm) == 2L && any(nm %in% c("lon", "long", "lng", "longitude")) &&
          any(nm %in% c("lat", "latitude"))) {
        lon <- as.numeric(vc[which(nm %in% c("lon", "long", "lng", "longitude"))][1])
        lat <- as.numeric(vc[which(nm %in% c("lat", "latitude"))][1])
      } else {
        lat <- as.numeric(vc[1])
        lon <- as.numeric(vc[2])
      }
      vc <- data.frame(latitude = lat, longitude = lon)
    } else {
      stop(
        "vertex_coords must be a 1-row data.frame or a length-2 numeric vector (lat, lon), ",
        "or a data frame / Excel or CSV path with the plot corners.",
        call. = FALSE
      )
    }
  }

  vc <- as.data.frame(vc, stringsAsFactors = FALSE)

  lat_col <- .pick_colname(vc, c("latitude", "lat", "decimalLatitude"))
  lon_col <- .pick_colname(vc, c("longitude", "lon", "long", "lng", "decimalLongitude"))

  if (is.na(lat_col) || is.na(lon_col)) {
    stop(
      "Could not find latitude/longitude columns in `vertex_coords`. ",
      "Available columns: ", paste(names(vc), collapse = " | "),
      call. = FALSE
    )
  }

  out <- data.frame(
    latitude = .numify(vc[[lat_col]]),
    longitude = .numify(vc[[lon_col]])
  )

  bad <- !is.finite(out$latitude) | !is.finite(out$longitude) |
    abs(out$latitude) > 90 | abs(out$longitude) > 180

  if (any(bad)) {
    stop(
      "Invalid coordinates in `vertex_coords` (row(s) ", paste(which(bad), collapse = ", "),
      "). Use decimal degrees, negative for south and west.",
      call. = FALSE
    )
  }

  out
}


#' Select the central coordinate for one MONITORA station
#'
#' @keywords internal
#' @noRd
.vertex_for_station <- function(vertex_coords, stations, i) {
  st <- stations[i]
  vc <- vertex_coords

  if (is.list(vc) && !is.data.frame(vc)) {
    if (!is.null(names(vc)) && st %in% names(vc)) return(vc[[st]])
    if (length(vc) == length(stations)) return(vc[[i]])
    stop("`vertex_coords` list must have one element per station, named or in the order of `station_name`.",
         call. = FALSE)
  }

  if (is.character(vc) && length(vc) == 1L && file.exists(vc)) {
    vc <- if (grepl("\\.csv$", vc, ignore.case = TRUE)) {
      utils::read.csv(vc, stringsAsFactors = FALSE, check.names = FALSE)
    } else {
      as.data.frame(readxl::read_excel(vc), stringsAsFactors = FALSE)
    }
  }

  if (is.data.frame(vc)) {
    st_col <- .pick_colname(vc, c("station_name", "station", "nome_estacao", "estacao"))
    if (!is.na(st_col)) {
      v_name <- .station_norm_name(vc[[st_col]])
      v_num <- .station_norm_num(vc[[st_col]])
      w_name <- .station_norm_name(st)
      w_num <- .station_norm_num(st)
      hit <- which((nzchar(v_name) & v_name == w_name) |
                     (!is.na(v_num) & !is.na(w_num) & v_num == w_num))
      if (length(hit) >= 1L) return(vc[hit[1], , drop = FALSE])
      stop("Station '", st, "' was not found in the station column of `vertex_coords`.",
           call. = FALSE)
    }
    if (nrow(vc) == length(stations)) return(vc[i, , drop = FALSE])
  }

  stop(
    "For several MONITORA stations, `vertex_coords` must give one center per station: ",
    "a named list, a data frame with one row per station, or a data frame with a station column.",
    call. = FALSE
  )
}


#' Warn when vertex distances do not match the plot dimensions
#'
#' @keywords internal
#' @noRd
.check_vertex_geometry <- function(p1, p2, p3, p4, plot_width_m, plot_length_m, tol = 0.25) {
  len_12 <- geosphere::distGeo(p1, p2)
  wid_13 <- geosphere::distGeo(p1, p3)

  off <- function(a, b) abs(a - b) / b

  if (off(len_12, plot_length_m) > tol || off(wid_13, plot_width_m) > tol) {
    warning(
      sprintf(
        paste0(
          "Vertex distances (corner 1-2 = %.0f m, corner 1-3 = %.0f m) do not match ",
          "plot_length_m = %s and plot_width_m = %s. Check the corner order: ",
          "1 = (0, 0), 2 = (0, length), 3 = (width, 0), 4 = (width, length)."
        ),
        len_12, wid_13, format(plot_length_m), format(plot_width_m)
      ),
      call. = FALSE
    )
  }

  invisible(TRUE)
}


#' Convert local plot coordinates to latitude/longitude (vectorized)
#'
#' @return Matrix with columns `lat` and `lon`.
#'
#' @keywords internal
#' @noRd
.get_latlon_mat <- function(x, y, p1, p2, p3, p4) {
  n <- max(length(x), length(y))
  x <- rep_len(as.numeric(x), n)
  y <- rep_len(as.numeric(y), n)

  left <- geosphere::destPoint(p1, geosphere::bearing(p1, p2), y)
  right <- geosphere::destPoint(p3, geosphere::bearing(p3, p4), y)
  interp <- geosphere::destPoint(left, geosphere::bearing(left, right), x)

  cbind(lat = unname(interp[, "lat"]), lon = unname(interp[, "lon"]))
}


#' Convert MONITORA offsets from the center to latitude/longitude (vectorized)
#'
#' @return Matrix with columns `lat` and `lon`.
#'
#' @keywords internal
#' @noRd
.get_latlon_from_center_mat <- function(x, y, center_lonlat) {
  n <- max(length(x), length(y))
  x <- rep_len(as.numeric(x), n)
  y <- rep_len(as.numeric(y), n)

  p1 <- geosphere::destPoint(center_lonlat, ifelse(x >= 0, 90, 270), abs(x))
  p2 <- geosphere::destPoint(p1, ifelse(y >= 0, 0, 180), abs(y))

  cbind(lat = unname(p2[, "lat"]), lon = unname(p2[, "lon"]))
}


#' Single-point wrappers kept for backwards compatibility
#'
#' Return a named vector `c(lat, lon)` for one point and a two-column matrix
#' for several points.
#'
#' @keywords internal
#' @noRd
.get_latlon <- function(x, y, p1, p2, p3, p4) {
  out <- .get_latlon_mat(x, y, p1, p2, p3, p4)
  if (nrow(out) == 1L) out[1, ] else out
}

.get_latlon_from_center <- function(x, y, center_lonlat) {
  out <- .get_latlon_from_center_mat(x, y, center_lonlat)
  if (nrow(out) == 1L) out[1, ] else out
}


.numify <- function(z) {
  if (is.numeric(z)) return(as.numeric(z))
  suppressWarnings(as.numeric(.normalize_decimal(z)))
}


.coerce_xy <- function(z) {
  if (is.numeric(z)) return(as.numeric(z))

  zc <- as.character(z)
  zc[is.na(zc)] <- ""
  zc <- gsub("[[:space:]]+", "", zc)
  # Keep only the first value of "3,5/9,0" before normalizing separators
  zc <- sub("^([^;/|]+)[;/|].*$", "\\1", zc, perl = TRUE)
  zc <- .normalize_decimal(zc)

  pat <- "[-+]?(?:\\d+\\.?\\d*|\\d*\\.?\\d+)"
  has_num <- grepl(pat, zc, perl = TRUE)

  out <- rep(NA_real_, length(zc))
  out[has_num] <- suppressWarnings(as.numeric(
    regmatches(zc[has_num], regexpr(pat, zc[has_num], perl = TRUE))
  ))
  out
}


.mk_cells <- function(arm) {
  base <- expand.grid(c = 0:4, r = 0:1)
  cs <- 10

  out <- if (arm == "N") {
    data.frame(xmin = -cs + base$r * cs, ymin = 50 + base$c * cs)
  } else if (arm == "S") {
    data.frame(xmin = -cs + base$r * cs, ymin = -100 + base$c * cs)
  } else if (arm == "L") {
    data.frame(xmin = 50 + base$c * cs, ymin = -cs + base$r * cs)
  } else {
    data.frame(xmin = -100 + base$c * cs, ymin = -cs + base$r * cs)
  }

  out$xmax <- out$xmin + cs
  out$ymax <- out$ymin + cs
  out$c <- base$c
  out$r <- base$r
  out$arm <- arm
  out
}


.get_subplot_polygon <- function(col, row, size) {
  x0 <- col * size
  y0 <- row * size
  matrix(
    c(
      x0, y0,
      x0 + size, y0,
      x0 + size, y0 + size,
      x0, y0 + size,
      x0, y0
    ),
    ncol = 2,
    byrow = TRUE
  )
}


#' Marker radius proportional to the square root of DBH
#'
#' @keywords internal
#' @noRd
.dbh_radius <- function(d, min_r = 3, max_r = 10) {
  d <- suppressWarnings(as.numeric(d))
  out <- rep(min_r + 1, length(d))
  ok <- is.finite(d) & d > 0
  if (!any(ok)) return(out)

  s <- sqrt(d[ok])
  rng <- range(s)
  out[ok] <- if (diff(rng) > 0) {
    min_r + (max_r - min_r) * (s - rng[1]) / diff(rng)
  } else {
    (min_r + max_r) / 2
  }
  out
}


#' Index voucher image folders by normalized voucher name
#'
#' @return Named list (normalized voucher -> image file paths).
#'
#' @keywords internal
#' @noRd
.voucher_image_index <- function(voucher_imgs) {
  base <- if (dir.exists(voucher_imgs)) voucher_imgs else here::here(voucher_imgs)

  if (!dir.exists(base)) {
    message("`voucher_imgs` folder not found: ", voucher_imgs, ". Skipping images.")
    return(list())
  }

  dirs <- list.dirs(base, recursive = TRUE, full.names = TRUE)
  dirs <- setdiff(dirs, base)

  out <- list()
  for (d in dirs) {
    imgs <- list.files(d, pattern = "\\.(jpe?g|png|gif|webp)$", ignore.case = TRUE,
                       full.names = TRUE)
    if (!length(imgs)) next
    key <- .norm_nm(basename(d))
    if (!nzchar(key)) next
    out[[key]] <- unique(c(out[[key]], sort(imgs)))
  }

  out
}


#' Turn image files into HTML sources (embedded data URIs or copied files)
#'
#' @keywords internal
#' @noRd
.voucher_image_sources <- function(img_index, embed, out_dir, asset_dirname) {
  if (!length(img_index)) return(list())

  if (isTRUE(embed)) {
    total <- sum(file.info(unlist(img_index))$size, na.rm = TRUE)
    if (total > 25 * 1024^2) {
      warning(sprintf("Embedding %.0f MB of voucher images; the HTML file will be large.",
                      total / 1024^2), call. = FALSE)
    }
    return(lapply(img_index, function(files) vapply(files, .file_data_uri, character(1),
                                                    USE.NAMES = FALSE)))
  }

  asset_root <- file.path(out_dir, asset_dirname)

  out <- list()
  for (key in names(img_index)) {
    dest_dir <- file.path(asset_root, key)
    if (!dir.exists(dest_dir)) dir.create(dest_dir, recursive = TRUE)
    files <- img_index[[key]]
    file.copy(files, file.path(dest_dir, basename(files)), overwrite = TRUE)
    out[[key]] <- vapply(
      basename(files),
      function(f) utils::URLencode(paste(asset_dirname, key, f, sep = "/")),
      character(1),
      USE.NAMES = FALSE
    )
  }
  out
}


.file_data_uri <- function(path) {
  ext <- tolower(tools::file_ext(path))
  mime <- switch(ext, jpg = , jpeg = "image/jpeg", png = "image/png",
                 gif = "image/gif", webp = "image/webp", "application/octet-stream")
  raw <- readBin(path, what = "raw", n = file.info(path)$size)
  paste0("data:", mime, ";base64,", jsonlite::base64_enc(raw))
}


.forplotr_logo_uri <- function() {
  f <- system.file("figures", "forplotR_hex_sticker.png", package = "forplotR")
  if (!nzchar(f) || !file.exists(f)) return("")
  .file_data_uri(f)
}


#' Build the HTML for a photo carousel
#'
#' @keywords internal
#' @noRd
.carousel_html <- function(srcs, embed = FALSE) {
  if (!length(srcs)) return("")

  slides <- vapply(srcs, function(s) {
    img <- paste0("<img loading='lazy' src='", s, "' alt=''>")
    if (isTRUE(embed)) {
      paste0("<div class='fp-slide'>", img, "</div>")
    } else {
      paste0("<a class='fp-slide' href='", s, "' target='_blank' rel='noopener'>", img, "</a>")
    }
  }, character(1), USE.NAMES = FALSE)

  nav <- if (length(srcs) > 1L) {
    paste0(
      "<div class='fp-nav'>",
      "<button type='button' class='fp-prev' aria-label='previous'>&#10094;</button>",
      "<span class='fp-count'>1 / ", length(srcs), "</span>",
      "<button type='button' class='fp-next' aria-label='next'>&#10095;</button>",
      "</div>"
    )
  } else ""

  paste0("<div class='fp-carousel slideshow-container' data-idx='0'>",
         paste(slides, collapse = ""), nav, "</div>")
}


#' Build the herbarium button from the lookup HTML fragment
#'
#' @keywords internal
#' @noRd
.herbarium_button <- function(frag, label) {
  out <- rep("", length(frag))
  ok <- !is.na(frag) & grepl("href='[^']+'", frag)
  if (!any(ok)) return(out)

  href <- sub(".*href='([^']+)'.*", "\\1", frag[ok])
  src <- ifelse(
    grepl("title='[^']+'", frag[ok]),
    sub(".*title='([^']+)'.*", "\\1", frag[ok]),
    ifelse(grepl(">[^<]+</a>", frag[ok]), sub(".*>([^<]+)</a>.*", "\\1", frag[ok]), "")
  )

  out[ok] <- paste0(
    "<a class='fp-btn' href='", htmltools::htmlEscape(href, attribute = TRUE),
    "' target='_blank' rel='noopener'>", htmltools::htmlEscape(label),
    ifelse(nzchar(src), paste0(" &middot; ", htmltools::htmlEscape(src)), ""),
    " &#8599;</a>"
  )
  out
}


#' Build specimen popups
#'
#' The `fp-pop` element carries the filter metadata as data attributes,
#' which the sidebar reads instead of parsing the visible text.
#'
#' @keywords internal
#' @noRd
.build_specimen_popup <- function(fp_coords, carousel, status_labels, tr, ui) {
  esc <- function(x) {
    x <- as.character(x)
    x[is.na(x)] <- ""
    htmltools::htmlEscape(x)
  }
  attr_esc <- function(x) {
    x <- as.character(x)
    x[is.na(x)] <- ""
    htmltools::htmlEscape(x, attribute = TRUE)
  }

  tag <- as.character(fp_coords[["New Tag No"]])
  fam <- fp_coords$Family
  sp <- fp_coords[["Original determination"]]
  sub <- fp_coords$subplot_label
  status <- fp_coords$status
  voucher <- as.character(fp_coords$Voucher)
  has_v <- !is.na(voucher) & nzchar(trimws(voucher))

  d_cm <- suppressWarnings(as.numeric(fp_coords$D)) / 10
  dbh_txt <- ifelse(is.finite(d_cm), formatC(d_cm, format = "f", digits = 1), "\u2014")

  herb <- .herbarium_button(fp_coords$herbaria_link, ui$herbarium)

  row_html <- function(label, value) {
    paste0("<tr><th>", esc(label), "</th><td>", value, "</td></tr>")
  }

  paste0(
    "<div class='fp-pop'",
    " data-tag=\"", attr_esc(tag), "\"",
    " data-fam=\"", attr_esc(fam), "\"",
    " data-sp=\"", attr_esc(sp), "\"",
    " data-sub=\"", attr_esc(sub), "\"",
    " data-status=\"", attr_esc(status), "\"",
    " data-voucher=\"", attr_esc(ifelse(has_v, voucher, "")), "\"",
    " data-photo=\"", ifelse(nzchar(carousel), "1", "0"), "\">",
    "<div class='fp-pop-head'>",
    "<div class='fp-pop-sp'>", esc(sp), "</div>",
    "<div class='fp-pop-fam'>", esc(fam), "</div>",
    "<span class='fp-badge fp-badge-", status, "'>", esc(status_labels[status]), "</span>",
    ifelse(nzchar(carousel), paste0(" <span class='fp-badge fp-badge-photo'>", esc(ui$photo), "</span>"), ""),
    "</div>",
    "<table class='fp-pop-tbl'>",
    row_html(ui$tag, esc(tag)),
    row_html(tr["subplot"], esc(sub)),
    row_html(ui$dbh, dbh_txt),
    row_html(ui$voucher, ifelse(has_v, esc(voucher),
                                paste0("<span class='fp-muted'>", esc(ui$unvouchered), "</span>"))),
    "</table>",
    ifelse(nzchar(herb), paste0("<div class='fp-pop-actions'>", herb, "</div>"), ""),
    carousel,
    "</div>"
  )
}


#' Stylesheet for the HTML map
#'
#' @keywords internal
#' @noRd
.forplot_map_css <- function() {
  "
.leaflet-container, .leaflet-popup-content, .fp-sidebar, .fp-head, .fp-legend {
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica Neue', Arial,
               'Noto Sans', sans-serif;
}
.fp-head-ctrl { background: none; box-shadow: none; }
.fp-head {
  display: flex; align-items: center; gap: 10px;
  background: rgba(255,255,255,0.95); border-radius: 12px;
  padding: 8px 14px 8px 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.18);
  max-width: min(440px, 70vw);
}
.fp-logo { width: 52px; height: auto; display: block; }
.fp-title { font-size: 18px; font-weight: 700; color: #1B2A1E; line-height: 1.2; }
.fp-sub { font-size: 13px; color: #4B5563; margin-top: 2px; }
.fp-stats { font-size: 12px; color: #374151; margin-top: 4px; }
.fp-ico { font-size: 18px; line-height: 26px; }

.fp-grid-label {
  color: rgba(255,255,255,0.85) !important; font-size: 10px !important;
  font-weight: 600 !important; text-shadow: 0 0 3px rgba(0,0,0,0.8);
}
.fp-hover { font-size: 12px; padding: 3px 7px; border-radius: 6px; }

.fp-legend { border-radius: 10px !important; box-shadow: 0 2px 10px rgba(0,0,0,0.18) !important;
             line-height: 20px; }
.fp-legend i { border-radius: 50%; border: 1px solid #1F2937; }

.leaflet-popup-content-wrapper { border-radius: 12px; box-shadow: 0 4px 16px rgba(0,0,0,0.22); }
.leaflet-popup-content { margin: 12px 14px; font-size: 13px; line-height: 1.35; }
.fp-pop-head { margin-bottom: 8px; }
.fp-pop-sp { font-size: 15px; font-weight: 600; font-style: italic; color: #111827; }
.fp-pop-fam { font-size: 12px; color: #6B7280; text-transform: uppercase; letter-spacing: 0.04em;
              margin: 2px 0 6px; }
.fp-badge { display: inline-block; font-size: 11px; font-weight: 600; padding: 2px 8px;
            border-radius: 999px; color: #fff; }
.fp-badge-collected { background: #6B7280; }
.fp-badge-missing { background: #E5484D; }
.fp-badge-palm { background: #C99700; }
.fp-badge-photo { background: #2E7D32; }
.fp-pop-tbl { width: 100%; border-collapse: collapse; }
.fp-pop-tbl th { text-align: left; font-weight: 500; color: #6B7280; padding: 3px 8px 3px 0;
                 white-space: nowrap; width: 1%; }
.fp-pop-tbl td { padding: 3px 0; color: #111827; }
.fp-pop-tbl tr + tr th, .fp-pop-tbl tr + tr td { border-top: 1px solid #F3F4F6; }
.fp-muted { color: #9CA3AF; font-style: italic; }
.fp-pop-actions { margin-top: 8px; }
.fp-btn { display: inline-block; padding: 5px 10px; border-radius: 8px; background: #2E7D32;
          color: #fff !important; text-decoration: none; font-size: 12px; font-weight: 600; }
.fp-btn:hover { background: #1B5E20; }

.fp-carousel { margin-top: 10px; }
.fp-slide { display: none; }
.fp-slide:first-child { display: block; }
.fp-slide img { width: 100%; max-height: 220px; object-fit: contain; border-radius: 8px;
                background: #F3F4F6; display: block; }
.fp-nav { display: flex; align-items: center; justify-content: center; gap: 12px; margin-top: 6px; }
.fp-nav button { border: none; background: #F3F4F6; border-radius: 50%; width: 28px; height: 28px;
                 cursor: pointer; font-size: 13px; color: #374151; }
.fp-nav button:hover { background: #E5E7EB; }
.fp-count { font-size: 12px; color: #6B7280; min-width: 40px; text-align: center; }

.fp-toggle {
  position: absolute; top: 10px; right: 10px; z-index: 1100;
  display: flex; align-items: center; gap: 6px;
  background: #fff; color: #1F2937; border: none; border-radius: 10px;
  padding: 7px 12px; font-size: 13px; font-weight: 600; cursor: pointer;
  box-shadow: 0 2px 10px rgba(0,0,0,0.2);
}
.fp-toggle .fp-dot { background: #2E7D32; color: #fff; border-radius: 999px;
                     padding: 0 7px; font-size: 11px; }
.fp-sidebar {
  position: absolute; top: 0; right: 0; height: 100%; width: min(320px, 88vw);
  background: rgba(255,255,255,0.98); box-shadow: -4px 0 16px rgba(0,0,0,0.2);
  overflow-y: auto; padding: 52px 14px 16px; z-index: 1050; box-sizing: border-box;
  transform: translateX(100%); transition: transform .25s ease-in-out; font-size: 13px;
}
.fp-sidebar.open { transform: translateX(0); }
.fp-search { width: 100%; box-sizing: border-box; padding: 8px 10px; border: 1px solid #D1D5DB;
             border-radius: 8px; font-size: 13px; }
.fp-search:focus { outline: 2px solid #A5D6A7; border-color: #2E7D32; }
.fp-summary { display: flex; justify-content: space-between; align-items: center;
              margin: 8px 0 4px; color: #4B5563; }
.fp-link { background: none; border: none; color: #2E7D32; cursor: pointer; font-size: 12px;
           padding: 0; font-weight: 600; }
.fp-empty { color: #E5484D; font-size: 12px; margin: 4px 0; display: none; }
.fp-sec { border-top: 1px solid #E5E7EB; padding-top: 8px; margin-top: 8px; }
.fp-sec > summary { cursor: pointer; font-weight: 700; color: #1F2937; list-style: none;
                    display: flex; justify-content: space-between; align-items: center; }
.fp-sec > summary::-webkit-details-marker { display: none; }
.fp-sec > summary::after { content: '+'; color: #9CA3AF; font-weight: 400; }
.fp-sec[open] > summary::after { content: '\\2212'; }
.fp-sec-tools { margin: 4px 0; display: flex; gap: 10px; }
.fp-list { max-height: 190px; overflow-y: auto; margin-top: 2px; }
.fp-opt { display: flex; align-items: center; gap: 6px; padding: 2px 0; cursor: pointer; }
.fp-opt span.fp-v { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.fp-opt span.fp-it { font-style: italic; }
.fp-opt span.fp-n { color: #9CA3AF; font-size: 11px; }
.fp-sw { width: 11px; height: 11px; border-radius: 50%; border: 1px solid #1F2937; display: inline-block; }
"
}


#' JavaScript for the sidebar filters and photo carousel
#'
#' Receives labels and settings through `htmlwidgets::onRender(data = ...)`.
#'
#' @keywords internal
#' @noRd
.forplot_map_js <- function() {
  "
function(el, x, data) {
  var map = this;
  var T = data.labels;

  function natSort(a, b) { return String(a).localeCompare(String(b), undefined, {numeric: true}); }
  function mk(tag, cls, txt) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (txt !== undefined) e.textContent = txt;
    return e;
  }

  /* Photo carousel */
  function showSlide(c, i) {
    var s = c.querySelectorAll('.fp-slide');
    if (!s.length) return;
    i = ((i % s.length) + s.length) % s.length;
    for (var k = 0; k < s.length; k++) s[k].style.display = (k === i) ? 'block' : 'none';
    c.setAttribute('data-idx', String(i));
    var ct = c.querySelector('.fp-count');
    if (ct) ct.textContent = (i + 1) + ' / ' + s.length;
  }

  map.on('popupopen', function(e) {
    var root = e.popup.getElement();
    if (!root) return;
    root.querySelectorAll('.fp-carousel').forEach(function(c) {
      showSlide(c, 0);
      if (c.getAttribute('data-bound') === '1') return;
      c.setAttribute('data-bound', '1');
      c.querySelectorAll('.fp-prev, .fp-next').forEach(function(b) {
        L.DomEvent.on(b, 'click', function(ev) {
          L.DomEvent.stop(ev);
          var step = b.classList.contains('fp-next') ? 1 : -1;
          showSlide(c, parseInt(c.getAttribute('data-idx') || '0', 10) + step);
        });
      });
    });
  });

  /* Markers and their metadata */
  var grp = null;
  if (map.layerManager && map.layerManager.getLayerGroup) {
    grp = map.layerManager.getLayerGroup(data.specimensGroup, false);
  }
  var host = grp || map;

  var markers = [];
  var collect = function(l) { if (l instanceof L.CircleMarker && l.getPopup && l.getPopup()) markers.push(l); };
  if (grp) grp.eachLayer(collect); else map.eachLayer(collect);
  if (!markers.length) return;

  var parser = new DOMParser();
  var famCount = {}, spCount = {}, subCount = {}, stCount = {}, sp2fam = {};
  var anyPhoto = false;

  markers.forEach(function(m) {
    var html = m.getPopup().getContent();
    var doc = parser.parseFromString(typeof html === 'string' ? html : '', 'text/html');
    var d = doc.querySelector('.fp-pop');
    var g = function(k) { return d ? (d.getAttribute('data-' + k) || '') : ''; };
    var f = { tag: g('tag'), fam: g('fam'), sp: g('sp'), sub: g('sub'),
              stat: g('status'), voucher: g('voucher'), photo: g('photo') === '1' };
    f.txt = (f.fam + ' ' + f.sp + ' ' + f.tag + ' ' + f.voucher).toLowerCase();
    m._fp = f;
    famCount[f.fam] = (famCount[f.fam] || 0) + 1;
    spCount[f.sp] = (spCount[f.sp] || 0) + 1;
    if (f.sub) subCount[f.sub] = (subCount[f.sub] || 0) + 1;
    stCount[f.stat] = (stCount[f.stat] || 0) + 1;
    sp2fam[f.sp] = f.fam;
    if (f.photo) anyPhoto = true;
  });

  /* Sidebar */
  var container = map.getContainer();
  var btn = mk('button', 'fp-toggle');
  btn.type = 'button';
  btn.innerHTML = '&#9776; ';
  btn.appendChild(document.createTextNode(T.filters + ' '));
  var dot = mk('span', 'fp-dot', String(markers.length));
  btn.appendChild(dot);

  var sb = mk('div', 'fp-sidebar');
  container.appendChild(sb);
  container.appendChild(btn);
  L.DomEvent.disableClickPropagation(sb);
  L.DomEvent.disableScrollPropagation(sb);
  L.DomEvent.disableClickPropagation(btn);
  btn.addEventListener('click', function() { sb.classList.toggle('open'); });

  var search = mk('input', 'fp-search');
  search.type = 'search';
  search.placeholder = T.search;
  sb.appendChild(search);

  var summary = mk('div', 'fp-summary');
  var countEl = mk('span');
  var resetBtn = mk('button', 'fp-link', T.reset);
  resetBtn.type = 'button';
  summary.appendChild(countEl);
  summary.appendChild(resetBtn);
  sb.appendChild(summary);

  var emptyEl = mk('div', 'fp-empty', T.no_results);
  sb.appendChild(emptyEl);

  var photoCb = null;
  if (anyPhoto) {
    var pl = mk('label', 'fp-opt');
    photoCb = mk('input');
    photoCb.type = 'checkbox';
    pl.appendChild(photoCb);
    pl.appendChild(mk('span', 'fp-v', T.photos_only));
    sb.appendChild(pl);
  }

  function section(title, open) {
    var det = mk('details', 'fp-sec');
    if (open) det.open = true;
    det.appendChild(mk('summary', '', title));
    sb.appendChild(det);
    return det;
  }

  function buildList(parent, counts, opts) {
    opts = opts || {};
    if (opts.tools) {
      var tools = mk('div', 'fp-sec-tools');
      var a = mk('button', 'fp-link', T.all); a.type = 'button';
      var n = mk('button', 'fp-link', T.none); n.type = 'button';
      tools.appendChild(a); tools.appendChild(n);
      parent.appendChild(tools);
    }
    var list = mk('div', 'fp-list');
    var keys = opts.order || Object.keys(counts).sort(natSort);
    keys.forEach(function(v) {
      if (!counts[v]) return;
      var lab = mk('label', 'fp-opt');
      var cb = mk('input'); cb.type = 'checkbox'; cb.value = v; cb.checked = true;
      lab.appendChild(cb);
      if (opts.colors) {
        var sw = mk('span', 'fp-sw'); sw.style.background = opts.colors[v];
        lab.appendChild(sw);
      }
      var txt = opts.names ? (opts.names[v] || v) : v;
      lab.appendChild(mk('span', 'fp-v' + (opts.italic ? ' fp-it' : ''), txt));
      lab.appendChild(mk('span', 'fp-n', String(counts[v])));
      list.appendChild(lab);
    });
    parent.appendChild(list);
    if (opts.tools) {
      a.addEventListener('click', function() { setAll(list, true); update(); });
      n.addEventListener('click', function() { setAll(list, false); update(); });
    }
    list.addEventListener('change', update);
    return list;
  }

  function setAll(list, val) {
    list.querySelectorAll('label').forEach(function(l) {
      if (l.style.display !== 'none') l.querySelector('input').checked = val;
    });
  }
  function checkedSet(list) {
    var s = {};
    list.querySelectorAll('input:checked').forEach(function(c) { s[c.value] = true; });
    return s;
  }

  var stList = buildList(section(data.statusLabel, true), stCount,
                         { order: data.statusOrder, colors: data.statusColors, names: data.statusLabels });
  var famList = buildList(section(data.familyLabel, false), famCount, { tools: true });
  var spList = buildList(section(data.speciesLabel, false), spCount, { tools: true, italic: true });
  var subList = Object.keys(subCount).length ?
    buildList(section(data.subplotLabel, false), subCount, { tools: true }) : null;

  function update() {
    var term = search.value.trim().toLowerCase();
    var st = checkedSet(stList), fs = checkedSet(famList), ss = checkedSet(spList);
    var sbs = subList ? checkedSet(subList) : null;
    var onlyPhoto = photoCb && photoCb.checked;
    var shown = 0;

    markers.forEach(function(m) {
      var f = m._fp;
      var ok = st[f.stat] && fs[f.fam] && ss[f.sp] &&
               (!sbs || !f.sub || sbs[f.sub]) &&
               (!onlyPhoto || f.photo) &&
               (!term || f.txt.indexOf(term) !== -1);
      if (ok) { shown++; if (!host.hasLayer(m)) host.addLayer(m); }
      else if (host.hasLayer(m)) host.removeLayer(m);
    });

    spList.querySelectorAll('label').forEach(function(l) {
      var v = l.querySelector('input').value;
      l.style.display = fs[sp2fam[v]] ? '' : 'none';
    });

    countEl.textContent = T.showing + ' ' + shown + ' ' + T.of + ' ' + markers.length;
    dot.textContent = String(shown);
    emptyEl.style.display = shown ? 'none' : 'block';
  }

  search.addEventListener('input', update);
  if (photoCb) photoCb.addEventListener('change', update);
  resetBtn.addEventListener('click', function() {
    search.value = '';
    if (photoCb) photoCb.checked = false;
    [stList, famList, spList, subList].forEach(function(l) {
      if (l) l.querySelectorAll('input').forEach(function(c) { c.checked = true; });
    });
    update();
  });

  update();
}
"
}
