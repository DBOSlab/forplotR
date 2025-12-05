#' Plot and save vouchered forest plot map as HTML
#'
#' @author Giulia Ottino & Domingos Cardoso
#'
#' @description
#' Generates an interactive and self-contained HTML map for forest plots
#' following either the \href{https://forestplots.net/}{ForestPlots plot protocol}
#' or the \href{https://www.gov.br/icmbio/pt-br/assuntos/monitoramento/programa-monitora}
#' {MONITORA monitoring program}, using Leaflet and based on vouchered tree data
#' and plot coordinates. The function handles both standard ForestPlots layouts
#' (1 ha or 0.5 ha plots defined by four corner vertices) and MONITORA
#' layouts (Maltese cross design defined by a single central coordinate).
#' It: (i) converts local subplot coordinates (X, Y) to geographic coordinates
#' using geospatial interpolation from plot vertices or a  central point;
#' (ii) extracts metadata such as team name, plot name, and plot code from the
#' input file; (iii) draws the plot boundary polygon or Maltese cross arms over
#' selectable basemaps; (iv) colors tree points by collection status and palms
#' (palms = yellow, collected = gray, missing = red); (v) embeds image carousels in
#' specimen popups when matching voucher images are found in the specified
#' folder; (vi) optionally performs a local herbarium lookup (e.g. JABOT DwC-A
#' downloads) to attach direct “Herbarium image” links to each voucher when
#' available; (vii) adds an interactive filter sidebar with checkboxes,
#' multi-select inputs, and a search box for filtering by family, species,
#' collection status, subplot, and presence of photos; (viii) displays
#' informative popups for each tree, showing tag number, subplot, family,
#' species, DBH, voucher code, herbarium image link (if found), and photos;
#' and (x) saves the resulting map as a date-stamped standalone HTML file with
#' multiple interactive Leaflet map options for easy sharing, archiving, or
#' field consultation.
#'
#' @usage plot_html_map(fp_file_path = NULL,
#'                      input_type = c("field_sheet", "fp_query_sheet", "monitora"),
#'                      vertex_coords = NULL,
#'                      plot_size = 1,
#'                      subplot_size = 10,
#'                      voucher_imgs = NULL,
#'                      filename = "plot_map",
#'                      station_name = NULL,
#'                      collector = NULL,
#'                      herbaria_lookup = FALSE,
#'                      herbaria = NULL,
#'                      jabot_cache_rds = "jabot_index.rds",
#'                      jabot_force_refresh = FALSE)
#'
#' @param fp_file_path Character. Path to the Excel file containing tree data.
#'
#' @param input_type One of \code{"field_sheet"}, \code{"fp_query_sheet"} or
#' \code{"monitora"}. For \code{"fp_query_sheet"}, the function expects a
#' ForestPlots Query Library export and converts it internally into a
#' field-sheet-like table. For \code{"monitora"}, it expects a Maltese-cross
#' layout.
#'
#' @param vertex_coords Data frame, path to an Excel file, or numeric vector.
#' For \code{"field_sheet"} and \code{"fp_query_sheet"}, provide four plot
#' corners (Latitude/Longitude). For \code{"monitora"}, provide a single
#' central Latitude/Longitude used as the cross center. A numeric vector of
#' length two is interpreted either as \code{c(lat, lon)} or as a named vector
#' with elements \code{lat}/\code{latitude} and \code{lon}/\code{longitude}.
#'
#' @param plot_size Numeric. Total plot size in hectares (default is 1).
#'
#' @param subplot_size Numeric. Subplot size in meters (default is 10).
#'
#' @param filename Character. Base name of the output HTML file (without
#' extension). The function will prepend a date-stamped prefix and append
#' \code{".html"} (default is \code{"plot_map"}).
#'
#' @param voucher_imgs Character. Directory path where voucher images are stored.
#' Subdirectories are assumed to be named after voucher IDs, each containing
#' one or more image files.
#'
#' @param station_name Character or \code{NULL}. For \code{"monitora"} inputs,
#' optionally filter or loop over one or more station names. When a vector of
#' station names is given, the function calls itself once per station.
#'
#' @param collector Character. Optional fallback collector name used when
#' compact voucher codes are provided (e.g. \code{"GCO1095"} without an explicit
#' collector string).
#'
#' @param herbaria_lookup Logical. If \code{TRUE}, the function performs a local
#' herbarium lookup and adds a \code{herbaria_link} column to the specimen data,
#' building a minimal index from JABOT DwC-A downloads (via \pkg{jabotR}).
#'
#' @param herbaria Character vector or \code{NULL}. Herbarium codes to be used
#' in the herbaria lookup, e.g. \code{c("RB","UPCB")}. When
#' \code{herbaria_lookup = TRUE} this argument must be a non-empty vector;
#' by default it is \code{NULL} and no herbaria search is performed unless the
#' user explicitly supplies the herbarium codes.
#'
#' @param jabot_cache_rds Character. Path to an \code{.rds} file used to cache
#' the JABOT-based index (keys and catalog numbers) between runs. Defaults to
#' \code{"jabot_index.rds"} in the working directory.
#'
#' @param jabot_force_refresh Logical. If \code{TRUE}, forces rebuilding the
#' JABOT-based index from the local DwC-A archives, ignoring any existing
#' cache file.
#'
#' @examples
#' \dontrun{
#' vertex_df <- data.frame(
#'   Latitude = c(-3.123, -3.123, -3.125, -3.125),
#'   Longitude = c(-60.012, -60.010, -60.012, -60.010)
#' )
#'
#' # Basic example (field_sheet)
#' plot_html_map(
#'   fp_file_path = "data/tree_data.xlsx",
#'   input_type = "field_sheet",
#'   vertex_coords = vertex_df,
#'   voucher_imgs = "voucher_imgs",
#'   filename = "plot_map"
#' )
#'
#' # Example with herbaria lookup
#' plot_html_map(
#'   fp_file_path = "data/RUS_plot.xlsx",
#'   input_type = "field_sheet",
#'   vertex_coords = "data/vertices.xlsx",
#'   voucher_imgs = "voucher_imgs",
#'   collector = "G. C. Ottino",
#'   herbaria_lookup = TRUE,
#'   herbaria = c("RB", "UPCB"),
#'   filename = "rus_plot_map"
#' )
#' }
#'
#' @importFrom leaflet leaflet addProviderTiles addPolygons addCircleMarkers setView addControl
#' @importFrom htmltools HTML tags htmlEscape save_html
#' @importFrom htmlwidgets onRender prependContent saveWidget
#' @importFrom geosphere destPoint bearing
#' @importFrom readxl read_excel
#' @importFrom here here
#' @importFrom dplyr select filter mutate case_when rename_with rename if_else
#' @importFrom magrittr "%>%"
#' @importFrom scales rescale
#' @importFrom openxlsx read.xlsx
#' @importFrom data.table fread
#'
#' @export
plot_html_map <- function(fp_file_path = NULL,
                          input_type = c("field_sheet", "fp_query_sheet", "monitora"),
                          vertex_coords = NULL,
                          plot_size = 1,
                          subplot_size = 10,
                          voucher_imgs = NULL,
                          filename = "plot_map",
                          station_name = NULL,
                          collector = NULL,
                          herbaria_lookup = FALSE,
                          herbaria = NULL,
                          jabot_cache_rds = "jabot_index.rds",
                          jabot_force_refresh = FALSE) {

  # --- Normalize / validate input type (now supports 'fp_query_sheet') ---
  input_type <- tolower(trimws(as.character(input_type)))
  input_type <- match.arg(input_type, c("field_sheet", "fp_query_sheet", "monitora"))

  # --- Validate input file path ---
  if (!file.exists(fp_file_path)) {
    stop("The provided 'fp_file_path' does not exist.", call. = FALSE)
  }

  # --- Optional: handle multiple MONITORA stations (same behavior as plot_for_balance) ---
  if (input_type == "monitora" && !is.null(station_name) && length(station_name) > 1L) {
    for (st in unique(trimws(as.character(station_name)))) {
      plot_html_map(fp_file_path = fp_file_path,
                    input_type = "monitora",
                    vertex_coords = vertex_coords,
                    plot_size = plot_size,
                    subplot_size = subplot_size,
                    voucher_imgs = voucher_imgs,
                    filename = paste0(filename, "_station_", st),
                    station_name = st,
                    collector = collector,
                    herbaria_lookup = herbaria_lookup,
                    herbaria = herbaria,
                    jabot_cache_rds = jabot_cache_rds,
                    jabot_force_refresh = jabot_force_refresh)
    }
    return(invisible(TRUE))
  }

  #  Read and convert the input to a standard dataset
  if (input_type == "monitora") {
    if (!exists(".monitora_to_field_sheet_df", mode = "function")) {
      stop("Missing helper `.monitora_to_field_sheet_df()` for MONITORA input.")
    }
    raw <- .monitora_to_field_sheet_df(fp_file_path, station_name = station_name)
  } else if (input_type == "fp_query_sheet") {
    if (!exists(".fp_query_to_field_sheet_df", mode = "function")) {
      stop("Missing helper `.fp_query_to_field_sheet_df()` for FP Query input.")
    }
    raw <- .fp_query_to_field_sheet_df(fp_file_path)
  } else {
    raw <- suppressMessages(openxlsx::read.xlsx(fp_file_path, sheet = 1, colNames = FALSE))
  }

  # --- Unify header and data extraction for all input types ---
  header_row <- as.character(raw[2, ])
  header_row[is.na(header_row)] <- paste0("NA_col_", seq_along(header_row))[is.na(header_row)]
  colnames(raw) <- make.unique(header_row)
  fp_sheet <- raw[-(1:2), , drop = FALSE]

  # --- Extract metadata (from row 1) ---
  meta_row_chr <- as.character(raw[1, ])
  meta_row_chr[is.na(meta_row_chr)] <- ""

  grab_meta <- function(key) {
    hit <- meta_row_chr[grepl(paste0("^", key, "[:]\\s*"), meta_row_chr, ignore.case = TRUE)]
    if (length(hit)) {
      sub(paste0("^", key, "[:]\\s*"), "", hit[1], ignore.case = TRUE)
    } else {
      ""
    }
  }

  team <- grab_meta("Team")
  plot_name <- grab_meta("Plot Name")
  plot_code <- grab_meta("Plotcode")

  # Ensure plot_code has hyphen between letters and digits (e.g. ABC-01)
  if (!grepl("-", plot_code)) {
    plot_code <- gsub("(?<=[A-Z])(?=\\d)", "-", plot_code, perl = TRUE)
  }

  # Clean and convert necessary columns
  numify <- function(z) suppressWarnings(as.numeric(gsub(",", ".", as.character(z))))
  coerce <- function(z) {
    zc <- as.character(z)
    out <- suppressWarnings(as.numeric(gsub(",", ".", zc)))
    bad <- is.na(out) & grepl("/", zc)
    out[bad] <- suppressWarnings(as.numeric(sub(".*/\\s*([0-9]+).*", "\\1", zc[bad])))
    out
  }

  fp_clean <- fp_sheet %>%
    dplyr::mutate(
      T1 = numify(T1),
      X = coerce(X),
      Y = coerce(Y),
      D = numify(D)
    ) %>%
    dplyr::filter(is.finite(T1), is.finite(X), is.finite(Y))

  if (!nrow(fp_clean)) {
    stop("No valid points to plot: T1/X/Y could not be parsed from the input sheet.", call. = FALSE)
  }

  # Geometry calculations
  if (input_type == "monitora") {
    if (!exists(".compute_global_coordinates_monitora", mode = "function")) {
      stop("Missing helper `.compute_global_coordinates_monitora()` for MONITORA layout.")
    }
    fp_coords <- .compute_global_coordinates_monitora(fp_clean)
    if (!all(c("draw_x", "draw_y") %in% names(fp_coords))) {
      stop("MONITORA helper must add 'draw_x' and 'draw_y' columns (meters).")
    }
  } else {
    max_coord <- plot_size * 100
    n_rows <- floor(max_coord / subplot_size)

    # Convert local to global coordinates
    fp_coords <- fp_clean %>%
      dplyr::mutate(
        col = floor((T1 - 1) / n_rows),
        row = (T1 - 1) %% n_rows,
        global_x = col * subplot_size + X,
        global_y = dplyr::if_else(
          col %% 2 == 0,
          row * subplot_size + Y,
          (n_rows - row - 1) * subplot_size + Y
        ),
        draw_x = global_x,
        draw_y = global_y
      )
  }

  # Read vertex_coords if given as path
  if (is.character(vertex_coords) && grepl("\\.xlsx?$", vertex_coords)) {
    vertex_coords <- readxl::read_excel(vertex_coords)
  }

  # --- Accept numeric vector for MONITORA center ---
  if (!is.null(vertex_coords) && !is.data.frame(vertex_coords)) {
    if (is.atomic(vertex_coords) && length(vertex_coords) == 2) {
      nm <- tolower(names(vertex_coords))
      if (length(nm) == 2 && any(nm %in% c("lon", "longitude"))) {
        lon <- as.numeric(vertex_coords[which(nm %in% c("lon", "longitude"))][1])
        lat <- as.numeric(vertex_coords[which(nm %in% c("lat", "latitude"))][1])
      } else {
        # assume c(lat, lon)
        lat <- as.numeric(vertex_coords[1])
        lon <- as.numeric(vertex_coords[2])
      }
      vertex_coords <- data.frame(latitude = lat, longitude = lon, check.names = FALSE)
    } else {
      stop("vertex_coords must be a 1-row data.frame or a length-2 numeric vector (lat, lon).",
           call. = FALSE)
    }
  }

  # Standardize and clean vertex coordinates
  vertex_coords <- vertex_coords %>%
    dplyr::rename_with(~tolower(gsub("\\s*\\(.*?\\)", "", .x)), dplyr::everything()) %>%
    dplyr::rename(latitude = latitude, longitude = longitude) %>%
    dplyr::mutate(
      longitude = as.numeric(gsub(",", ".", longitude)),
      latitude = as.numeric(gsub(",", ".", latitude))
    )

  # Correct hemisphere signs only for 4-corner ForestPlots inputs
  if (input_type != "monitora" && nrow(vertex_coords) >= 4) {
    vertex_coords <- vertex_coords %>%
      dplyr::mutate(
        longitude = ifelse(longitude > 0, -longitude, longitude),
        latitude = ifelse(latitude > 0, -latitude, latitude)
      )
  }

  if (input_type == "monitora") {
    if (nrow(vertex_coords) != 1) {
      stop("For 'monitora', provide ONE central coordinate in 'vertex_coords'.", call. = FALSE)
    }
    center <- c(vertex_coords$longitude[1], vertex_coords$latitude[1])
    coords_geo <- vapply(
      seq_len(nrow(fp_coords)),
      function(i) .get_latlon_from_center(fp_coords$draw_x[i], fp_coords$draw_y[i], center),
      FUN.VALUE = c(lat = NA_real_, lon = NA_real_)
    )

    fp_coords$Latitude <- coords_geo["lat", ]
    fp_coords$Longitude <- coords_geo["lon", ]
  } else {
    # Use coordinates after cleaning (four vertices)
    if (nrow(vertex_coords) < 4) {
      stop("For 'field_sheet' and 'fp_query_sheet' you must provide FOUR plot corners in 'vertex_coords'.",
           call. = FALSE)
    }
    p1 <- c(vertex_coords$longitude[1], vertex_coords$latitude[1])
    p2 <- c(vertex_coords$longitude[2], vertex_coords$latitude[2])
    p3 <- c(vertex_coords$longitude[3], vertex_coords$latitude[3])
    p4 <- c(vertex_coords$longitude[4], vertex_coords$latitude[4])

    # Robust to mapply simplification (always 2×N)
    coords_geo <- simplify2array(
      mapply(
        .get_latlon,
        x = fp_coords$draw_x,
        y = fp_coords$draw_y,
        MoreArgs = list(p1 = p1, p2 = p2, p3 = p3, p4 = p4),
        SIMPLIFY = FALSE
      )
    )
    if (is.null(dim(coords_geo))) {
      coords_geo <- matrix(coords_geo, nrow = 2)
    }
    rownames(coords_geo) <- c("lat", "lon")

    fp_coords$Latitude <- coords_geo[1, , drop = TRUE]
    fp_coords$Longitude <- coords_geo[2, , drop = TRUE]

    # Coerce to numeric and drop invalid points (avoids leaflet errors)
    fp_coords$Latitude <- suppressWarnings(as.numeric(fp_coords$Latitude))
    fp_coords$Longitude <- suppressWarnings(as.numeric(fp_coords$Longitude))
    keep <- is.finite(fp_coords$Latitude) & is.finite(fp_coords$Longitude)
    fp_coords <- fp_coords[keep, , drop = FALSE]
    if (!nrow(fp_coords)) {
      stop("No valid points to plot: Latitude/Longitude are missing or non-numeric after conversion.",
           call. = FALSE)
    }
  }

  # Generate subplot grid rectangles and labels
  grid_polys <- list()
  grid_labels <- data.frame()

  if (input_type == "monitora") {
    # Build 5x2 cells for each arm of the Maltese cross (10 m cells)
    mk_cells <- function(arm) {
      base <- expand.grid(c = 0:4, r = 0:1)   # c: along the arm (5 columns), r: width (2 rows)
      cs <- 10                                 # subplot size (10 m)

      if (arm == "N") {
        # from +50 to +100 m on the Y axis
        dplyr::mutate(
          base,
          xmin = -cs + r * cs, xmax = -cs + r * cs + cs,
          ymin = 50 + c * cs, ymax = 50 + c * cs + cs
        )
      } else if (arm == "S") {
        # from -100 to -50 m on the Y axis
        dplyr::mutate(
          base,
          xmin = -cs + r * cs, xmax = -cs + r * cs + cs,
          ymin = -100 + c * cs, ymax = -100 + c * cs + cs
        )
      } else if (arm == "L") {  # East
        # from +50 to +100 m on the X axis
        dplyr::mutate(
          base,
          xmin = 50 + c * cs, xmax = 50 + c * cs + cs,
          ymin = -cs + r * cs, ymax = -cs + r * cs + cs
        )
      } else {                   # O = West
        # from -100 to -50 m on the X axis
        dplyr::mutate(
          base,
          xmin = -100 + c * cs, xmax = -100 + c * cs + cs,
          ymin = -cs + r * cs, ymax = -cs + r * cs + cs
        )
      }
    }

    cells <- dplyr::bind_rows(
      mk_cells("N") %>% dplyr::mutate(arm = "N"),
      mk_cells("S") %>% dplyr::mutate(arm = "S"),
      mk_cells("L") %>% dplyr::mutate(arm = "L"),
      mk_cells("O") %>% dplyr::mutate(arm = "O")
    )

    centers <- cells %>%
      dplyr::mutate(
        cx = (xmin + xmax) / 2,
        cy = (ymin + ymax) / 2,
        col_from_center = dplyr::case_when(arm %in% c("N", "L") ~ c, TRUE ~ 4L - c),
        col_idx = col_from_center + 1L,
        base_num = (col_idx - 1L) * 2L + 1L,
        label = ifelse(col_idx %% 2L == 1L, base_num + r, base_num + (1L - r))
      )

    # Convert each rect to lat/lon
    for (k in seq_len(nrow(cells))) {
      XY <- with(cells[k, ], rbind(
        c(xmin, ymin), c(xmax, ymin), c(xmax, ymax), c(xmin, ymax), c(xmin, ymin)
      ))
      llp <- t(apply(XY, 1, function(v) .get_latlon_from_center(v[1], v[2], center)))
      grid_polys[[k]] <- list(lng = llp[, 2], lat = llp[, 1])
    }
    lab_ll <- t(apply(as.matrix(centers[, c("cx", "cy")]), 1, function(v)
      .get_latlon_from_center(v[1], v[2], center)))
    grid_labels <- data.frame(
      subplot = paste0(centers$arm, centers$label),
      lat = lab_ll[, 1],
      lon = lab_ll[, 2]
    )
  } else {
    # ForestPlots grid from four vertices
    # Helper to get the 4 corners of each subplot in local coordinates
    get_subplot_polygon <- function(col, row, size) {
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
        ncol = 2, byrow = TRUE
      )
    }

    grid_polys <- list()
    grid_labels <- data.frame()

    subplot_index <- 1
    for (row in 0:(n_rows - 1)) {
      for (col in 0:(n_rows - 1)) {
        local_poly <- get_subplot_polygon(col, row, subplot_size)
        center_x <- mean(local_poly[, 1])
        center_y <- mean(local_poly[, 2])

        coords_latlon <- t(mapply(
          .get_latlon,
          x = local_poly[, 1],
          y = local_poly[, 2],
          MoreArgs = list(p1 = p1, p2 = p2, p3 = p3, p4 = p4)
        ))

        # Store polygon
        grid_polys[[subplot_index]] <- list(
          lng = coords_latlon[, "lon"],
          lat = coords_latlon[, "lat"]
        )

        # Store label at center
        center_ll <- .get_latlon(center_x, center_y, p1, p2, p3, p4)
        t1_label <- if (col %% 2 == 0) {
          row + col * n_rows + 1
        } else {
          (n_rows - row - 1) + col * n_rows + 1
        }
        grid_labels <- rbind(
          grid_labels,
          data.frame(
            subplot = t1_label,
            lat = center_ll["lat"],
            lon = center_ll["lon"]
          )
        )

        subplot_index <- subplot_index + 1
      }
    }
  }

  # Assign color by collection status
  fp_coords$color <- dplyr::case_when(
    fp_coords$Family == "Arecaceae" ~ "gold",
    !is.na(fp_coords$Collected) & fp_coords$Collected != "" ~ "gray",
    TRUE ~ "red"
  )

  # ----------------------- herbaria LOOKUP  -----------------------
  if (isTRUE(herbaria_lookup)) {
    fp_coords$herbaria_link <- .herbaria_lookup_links(
      fp_df = fp_coords,
      herbariums = herbaria,
      cache_rds = jabot_cache_rds,
      force_refresh = jabot_force_refresh,
      verbose = FALSE,
      collector_fallback = collector
    )
  } else {
    fp_coords$herbaria_link <- NA_character_
  }

  # Final popup HTML (adds herbarium link only when available)
  fp_coords$popup <- paste0(
    "<b>Tag:</b> ", fp_coords[["New Tag No"]], "<br/>",
    "<b>Subplot:</b> ", fp_coords$T1, "<br/>",
    "<b>Family:</b> ", fp_coords$Family, "<br/>",
    "<b>Species:</b> <i class='taxon'>", fp_coords[["Original determination"]], "</i><br/>",
    "<b>DBH (mm):</b> ", round(fp_coords$D, 2), "<br/>",
    "<b>Voucher:</b> ",
    ifelse(
      is.na(fp_coords$Voucher) | !nzchar(fp_coords$Voucher),
      "NA", fp_coords$Voucher
    ),
    ifelse(
      isTRUE(herbaria_lookup) & !is.na(fp_coords$herbaria_link) & nzchar(fp_coords$herbaria_link),
      paste0(
        "<br/><b>Herbarium image:</b> ",
        "<a href='", fp_coords$herbaria_link, "' target='_blank'>open</a>"
      ),
      ""
    )
  )

  # -------------------------------------------------------------------------------

  plot_popup <- paste0(
    "<b>Plot Name:</b> ",
    plot_name, "<br/>",
    "<b>Plot Code:</b> ", plot_code, "<br/>",
    "<b>Team:</b> ", team
  )

  # Define the base folder where the original images are stored
  original_image_base <- here::here(voucher_imgs)

  # List all subdirectories once to improve performance
  leaf_dirs <- list.dirs(original_image_base, recursive = TRUE, full.names = TRUE)

  # Keep only those that do not contain other directories (i.e., leaf directories)
  leaf_dirs <- leaf_dirs[!sapply(leaf_dirs, function(dir) {
    any(file.info(list.dirs(dir, recursive = FALSE))$isdir)
  })]

  leaf_dirs <- sub(paste0(".*(", voucher_imgs, "/.*)"), "\\1", leaf_dirs)

  # -------------------------------------------------
  #  Build pop-ups with photo carousels
  handled <- character(0)
  # Ensure text types
  fp_coords$popup <- as.character(fp_coords$popup)
  fp_coords$Voucher <- as.character(fp_coords$Voucher)

  for (dir_path in leaf_dirs) {
    voucher_id <- basename(dir_path)

    # Skip if this voucher was already handled
    if (voucher_id %in% handled) next
    handled <- c(handled, voucher_id)

    # List image files in this directory
    img_files <- list.files(
      dir_path, full.names = TRUE,
      pattern = "\\.(jpg|jpeg|png|gif)$", ignore.case = TRUE
    )
    if (!length(img_files)) next

    rel_paths <- file.path(dir_path, basename(img_files))

    # Build HTML slides for the carousel
    slides <- paste0(
      "<div class='mySlides'>",
      "<a href='", rel_paths, "' target='_blank'>",
      "<img src='", rel_paths,
      "' style='max-width:200px; max-height:200px; display:block; margin:auto;'></a></div>",
      collapse = "\n"
    )

    tf <- which(!is.na(fp_coords$Voucher) & fp_coords$Voucher == voucher_id)
    if (!length(tf)) next

    fp_coords$popup[tf] <- paste0(
      fp_coords$popup[tf],
      "<br><span class='hasPhotoBadge'>Has Photo</span>",
      "<div class='slideshow-container' data-slide='1'>", slides, "</div>",
      "<div style='text-align:center;'>",
      "<a class='prev' onclick='plusSlides(-1)' style='margin-right:10px;'>&#10094;</a>",
      "<a class='next' onclick='plusSlides(1)'>&#10095;</a>",
      "</div>"
    )
  }

  #-------------------------------------------------------------------------------
  carousel_js_css <- htmltools::HTML("
<style>
.mySlides { display: none; }
.mySlides img { border-radius: 5px; }
.prev, .next {
  cursor: pointer;
  font-size: 18px;
  user-select: none;
  color: #333;
}

/* 'Send comment' button styling */
.send-comment-btn {
  margin-top:6px;
  padding:4px 8px;
  font-size:12px;
  border-radius:6px;
  border:none;
  background:#2e7d32;
  color:#fff;
  cursor:pointer;
}
.send-comment-btn:hover {
  background:#1b5e20;
}
</style>

<script>
function initSlidesInPopup(popupEl) {
  var container = popupEl.querySelector('.slideshow-container');
  if (!container) return;
  var slides = container.getElementsByClassName('mySlides');
  if (!slides.length) return;
  for (var i = 0; i < slides.length; i++) slides[i].style.display = 'none';
  slides[0].style.display = 'block';
  container.setAttribute('data-slide', '1');
}

function shiftSlides(offset) {
  var popup = document.querySelector('.leaflet-popup-content');
  if (!popup) return;
  var container = popup.querySelector('.slideshow-container');
  if (!container) return;
  var slides = container.getElementsByClassName('mySlides');
  if (!slides.length) return;

  var idx = parseInt(container.getAttribute('data-slide') || '1', 10);
  idx += offset;
  if (idx > slides.length) idx = 1;
  if (idx < 1) idx = slides.length;

  for (var i = 0; i < slides.length; i++) slides[i].style.display = 'none';
  slides[idx - 1].style.display = 'block';
  container.setAttribute('data-slide', String(idx));
}

// Delegated prev/next
document.addEventListener('click', function(e) {
  if (e.target.classList.contains('prev')) {
    e.preventDefault();
    shiftSlides(-1);
  }
  if (e.target.classList.contains('next')) {
    e.preventDefault();
    shiftSlides(1);
  }
});

// Initialize slides when popup opens
document.addEventListener('click', function(e) {
  if (e.target.closest('.leaflet-marker-icon') || e.target.closest('.leaflet-interactive')) {
    setTimeout(function() {
      var popup = document.querySelector('.leaflet-popup-content');
      if (popup) initSlidesInPopup(popup);
    }, 120);
  }
});
</script>
")

  sidebar_css_js <- htmltools::HTML("
<style>
/* ---------- SIDEBAR STYLES ---------- */
#sidebar {
  position:absolute; top:0; right:0;
  width:280px; height:100%;
  background:rgba(255,255,255,0.97);
  box-shadow:-4px 0 10px rgba(0,0,0,.3);
  overflow-y:auto; padding:12px 14px;
  transform:translateX(100%);
  transition:transform .28s ease-in-out;
  z-index:1001; font-size:14px;
}
#sidebar.open { transform:translateX(0); }

#sidebarToggle {
  position:absolute; top:10px; right:15px;
  width:34px; height:34px;
  border-radius:4px; background:#fff; color:#333;
  box-shadow:0 0 5px rgba(0,0,0,.35);
  cursor:pointer; z-index:1100;
  display:flex; align-items:center; justify-content:center;
  font-size:20px; user-select:none;
}

#searchInput {
  width:100%; padding:6px 8px; margin-bottom:8px;
  border:1px solid #999; border-radius:4px;
}

.leaflet-popup-content { max-width:260px; padding:4px; }
.leaflet-popup-content img { width:100%; height:auto; border-radius:4px; }

.mySlides{display:none;}
.prev,.next{cursor:pointer;font-size:18px;color:#444;}
</style>

<script>
function addFilterControl(el,x){
  /* ---------- BUILD SIDEBAR ---------- */
  const map=this, mapDiv=map.getContainer(),
        sb=document.createElement('div'),
        btn=document.createElement('div');
  sb.id='sidebar'; btn.id='sidebarToggle'; btn.innerHTML='&#9776;';
  mapDiv.appendChild(sb); mapDiv.appendChild(btn);
  btn.addEventListener('click',()=>sb.classList.toggle('open'));

  sb.innerHTML=`
    <input id='searchInput' type='text' placeholder='Search family or species'/>
    <div style='margin-bottom:6px;'><b>Enable Filters:</b><br/>
      <label><input type='checkbox' id='toggleFamily' checked> Family</label><br/>
      <label><input type='checkbox' id='toggleSpecies' checked> Species</label><br/>
      <label><input type='checkbox' id='toggleColl' checked> Collection&nbsp;Status</label><br/>
      <label><input type='checkbox' id='togglePhotos'> With&nbsp;Photos&nbsp;Only</label><br/>
      <label><input type='checkbox' id='toggleSubplot' checked> Subplots</label>
    </div>
    <div id='familyFilterContainer'>
      <b>Filter by Family:</b><br/>
      <select id='familyFilter' multiple size='6' style='width:100%;'></select>
    </div><br/>
    <div id='speciesFilterContainer'>
      <b>Filter by Species:</b><br/>
      <select id='speciesFilter' multiple size='6' style='width:100%;'></select>
    </div><br/>
    <div id='collFilterContainer'>
      <b>Filter by Collection Status:</b><br/>
      <select id='collFilter' multiple size='3' style='width:100%;'>
        <option value='collected'>Collected (Gray)</option>
        <option value='missing'>Not Collected (Red)</option>
        <option value='palm'>Palm (Yellow)</option>
      </select>
    </div><br/>
    <div id='subplotFilterContainer' style='display:none;'>
      <b>Filter by Subplot:</b><br/>
      <select id='subplotFilter' multiple size='6' style='width:100%;'></select>
    </div>`;

  /* ---------- COLLECT MARKER DATA ---------- */
  const markers=[];
  map.eachLayer(l=>{ if(l instanceof L.CircleMarker) markers.push(l); });

  const famSet=new Set(), spSet=new Set(), sbSet=new Set(),
        fam2sp={}, sp2fam={};

  markers.forEach(m=>{
    const html=m.getPopup().getContent();
    const fam =(html.match(/<b>Family:<\\/b>\\s*(.*?)<br\\/>/)||[])[1]||'';
    const sp =(html.match(/<b>Species:<\\/b>\\s*<i[^>]*>\\s*(.*?)<\\/i>/)||[])[1]||'';
    const sb =(html.match(/<b>Subplot:<\\/b>\\s*(\\d+)/)||[])[1]||'';

    if(fam) famSet.add(fam);
    if(sp)  spSet.add(sp);
    if(sb)  sbSet.add(sb);

    sp2fam[sp]=fam;
    (fam2sp[fam]=fam2sp[fam]||new Set()).add(sp);

    Object.assign(m,{ _fam:fam, _sp:sp, _subplot:sb, _stat:getStatus(m),
                      _hasPhoto:html.includes('slideshow-container') });
  });

  /* ---------- UTILS ---------- */
  const famSel=document.getElementById('familyFilter'),
        spSel =document.getElementById('speciesFilter'),
        csSel =document.getElementById('collFilter'),
        sbSel =document.getElementById('subplotFilter'),
        search=document.getElementById('searchInput');

  const fill=(sel,set,italic=false)=>{
    sel.innerHTML='';
    Array.from(set).sort((a,b)=>+a - +b || a.localeCompare(b)).forEach(v=>{
      const o=document.createElement('option');
      o.value=v; o.selected=true; o.innerHTML=italic?`<i>${v}</i>`:v;
      sel.appendChild(o);
    });
  };

  fill(famSel,famSet);
  fill(spSel ,spSet ,true);
  fill(sbSel ,sbSet ,false);
  Array.from(csSel.options).forEach(o=>o.selected=true);

  const getSel=sel=>Array.from(sel.selectedOptions).map(o=>o.value);
  function getStatus(m){
    const c=m.options.fillColor.toLowerCase();
    if(['gray','#808080'].includes(c)) return 'collected';
    if(['red','#ff0000'].includes(c)) return 'missing';
    if(['gold','yellow','#ffd700'].includes(c)) return 'palm';
    return '';
  }

  /* ---------- EVENT LISTENERS ---------- */
  famSel.addEventListener('change',()=>{
    if(document.getElementById('toggleSpecies').checked){
      const chosen=getSel(famSel);
      const subset=new Set();
      chosen.forEach(f=>fam2sp[f]?.forEach(sp=>subset.add(sp)));
      fill(spSel,subset,true);
    }
    updateMarkers();
  });

  spSel.addEventListener('change',updateMarkers);
  csSel.addEventListener('change',updateMarkers);
  sbSel.addEventListener('change',updateMarkers);
  search.addEventListener('keyup',updateMarkers);

  ['Family','Species','Coll','Photos','Subplot'].forEach(k=>{
    const cb=document.getElementById('toggle'+k),
          block=document.getElementById(k.toLowerCase()+'FilterContainer');
    cb.addEventListener('change',()=>{
      if(block) block.style.display=cb.checked?'block':'none';
      if(!cb.checked){
        if(k==='Species') fill(spSel,spSet,true);
        if(k==='Coll')    Array.from(csSel.options).forEach(o=>o.selected=true);
        if(k==='Subplot') fill(sbSel,sbSet,false);
      }
      updateMarkers();
    });
  });

  /* ---------- FILTER FUNCTION ---------- */
  function updateMarkers(){
    const fOn=document.getElementById('toggleFamily').checked,
          sOn=document.getElementById('toggleSpecies').checked,
          cOn=document.getElementById('toggleColl').checked,
          pOn=document.getElementById('togglePhotos').checked,
          bOn=document.getElementById('toggleSubplot').checked,
          term=search.value.trim().toLowerCase();

    const selFam=fOn?getSel(famSel):Array.from(famSet),
          selSp =sOn?getSel(spSel):Array.from(spSet),
          selSt =cOn?getSel(csSel):['collected','missing','palm'],
          selSb =bOn?getSel(sbSel):Array.from(sbSet);

    markers.forEach(m=>{
      const txt=m._fam.toLowerCase()+m._sp.toLowerCase();
      const show= selFam.includes(m._fam) &&
                  selSp .includes(m._sp ) &&
                  selSt .includes(m._stat) &&
                  selSb .includes(m._subplot) &&
                  (!pOn || m._hasPhoto) &&
                  txt.includes(term);

      if(show){ if(!map.hasLayer(m)) m.addTo(map); }
      else { if(map.hasLayer(m)) map.removeLayer(m); }
    });

    map.eachLayer(l => {
      if (l.options && l.options.group === 'Subplot Grid') {
        if (bOn) {
          if (!map.hasLayer(l)) map.addLayer(l);
        } else {
          if (map.hasLayer(l)) map.removeLayer(l);
        }
      }
    });

    if(sOn){
      const visibleSp=new Set();
      markers.filter(m=>map.hasLayer(m)).forEach(m=>visibleSp.add(m._sp));
      fill(spSel,visibleSp,true);
    }
  }
}
</script>
")

  #_____________________________________________________________________________
  # Define initial center/zoom
  initial_zoom <- 18
  lat0 <- mean(vertex_coords$latitude)
  lon0 <- mean(vertex_coords$longitude)

  # Build the Leaflet map: 5 base layers + reset button
  map <- leaflet::leaflet(
    fp_coords,
    options = leaflet::leafletOptions(minZoom = 2, maxZoom = 22)
  ) %>%

    # Reset-view button
    leaflet::addEasyButton(
      leaflet::easyButton(
        icon = "fa-crosshairs",
        title = "Back to plot",
        onClick = leaflet::JS(
          sprintf("function(btn, map){ map.setView([%f, %f], %d); }",
                  lat0, lon0, initial_zoom)
        )
      )
    ) %>%

    # Set title and subtitle
    leaflet::addControl(
      html = paste0(
        "<div style='text-align:center; line-height:1.2; padding:4px 8px; ",
        "background:rgba(255,255,255,0.85); border-radius:8px;",
        "box-shadow:0 1px 4px rgba(0,0,0,0.25);'>",
        "<div style='font-size:20px; font-weight:700;'>",
        htmltools::htmlEscape(plot_name),
        "</div>",
        "<div style='font-size:14px; font-weight:400;'>Plot Code: ",
        htmltools::htmlEscape(plot_code),
        "</div>",
        "</div>"
      ),
      position = "topleft",
      className = "custom-title"
    ) %>%

    # Add clickable project logo to top-left corner
    leaflet::addControl(
      html = "<a href='https://dboslab.github.io/forplotR-website/' target='_blank'>
                <img src='inst/figures/forplotR_hex_sticker.png' style='width: 70px; height: auto;'>
              </a>",
      position = "topleft",
      className = "project-logo"
    ) %>%

    htmlwidgets::onRender(
      "function(el, x){ this.whenReady(function(){ addFilterControl.call(this, el, x); }); }"
    ) %>%

    htmlwidgets::prependContent(
      carousel_js_css,
      sidebar_css_js
    ) %>%

    leaflet::addProviderTiles(leaflet::providers$OpenStreetMap, group = "OSM Street") %>%
    leaflet::addProviderTiles(leaflet::providers$Esri.WorldStreetMap, group = "ESRI Street") %>%
    leaflet::addProviderTiles(leaflet::providers$Esri.WorldImagery, group = "Satellite")  %>%
    leaflet::addProviderTiles(leaflet::providers$OpenTopoMap, group = "OpenTopo")   %>%
    leaflet::addProviderTiles(leaflet::providers$CartoDB.DarkMatter, group = "Dark Matter") %>%

    # Layer control
    leaflet::addLayersControl(
      baseGroups = c("OSM Street", "ESRI Street", "Satellite", "OpenTopo", "Dark Matter"),
      overlayGroups = c("Specimens", "Subplot Grid"),
      position = "bottomleft",
      options = leaflet::layersControlOptions(
        collapsed = TRUE,
        autoZIndex = TRUE
      )
    ) %>%

    # Add subplot labels (faint, watermark style)
    leaflet::addLabelOnlyMarkers(
      data = grid_labels,
      lng = ~lon, lat = ~lat,
      label = ~as.character(subplot),
      labelOptions = leaflet::labelOptions(
        noHide = TRUE,
        direction = "center",
        textOnly = TRUE,
        style = list(
          "font-size" = "10px",
          "color" = "#888888",
          "font-weight" = "300",
          "opacity" = "0.4"
        )
      ),
      group = "Subplot Grid"
    )

  # Plot polygon (only for field_sheet / 4 vertices)
  if (input_type != "monitora") {
    map <- map %>%
      leaflet::addPolygons(
        lng = c(p1[1], p2[1], p4[1], p3[1], p1[1]),
        lat = c(p1[2], p2[2], p4[2], p3[2], p1[2]),
        color = "darkgreen", fillColor = "lightgreen", fillOpacity = 0.15,
        popup = plot_popup,
        options = leaflet::pathOptions(interactive = FALSE)
      )
  }

  # Resume the pipe for the markers and the rest
  map <- map %>%
    leaflet::addCircleMarkers(
      lng = ~Longitude,
      lat = ~Latitude,
      radius = scales::rescale(
        fp_coords$D, to = c(2, 8),
        from = range(fp_coords$D, na.rm = TRUE)
      ),
      color = "black",
      weight = 0.5,
      fillColor = ~color,
      fillOpacity = 0.8,
      popup = ~popup,
      group = "Specimens"
    ) %>%
    leaflet::setView(lng = lon0, lat = lat0, zoom = initial_zoom)

  # Maltese cross axes (only for MONITORA)
  if (input_type == "monitora") {
    arm <- 100

    eo <- rbind(
      .get_latlon_from_center(-arm, 0, c(lon0, lat0)),
      .get_latlon_from_center(arm, 0, c(lon0, lat0))
    )
    ns <- rbind(
      .get_latlon_from_center(0, -arm, c(lon0, lat0)),
      .get_latlon_from_center(0, arm, c(lon0, lat0))
    )

    eo <- as.matrix(eo); ns <- as.matrix(ns)

    map <- map %>%
      leaflet::addPolylines(
        lng = eo[, 2], lat = eo[, 1],
        color = "#666666", weight = 2, opacity = 0.9,
        group = "Maltese Cross",
        options = leaflet::pathOptions(interactive = FALSE)
      ) %>%
      leaflet::addPolylines(
        lng = ns[, 2], lat = ns[, 1],
        color = "#666666", weight = 2, opacity = 0.9,
        group = "Maltese Cross",
        options = leaflet::pathOptions(interactive = FALSE)
      )
  }

  # Add subplot grid rectangles AFTER map is built
  for (i in seq_along(grid_polys)) {
    map <- map %>% leaflet::addPolygons(
      lng = grid_polys[[i]]$lng,
      lat = grid_polys[[i]]$lat,
      color = "#999999", weight = 0.5, fillOpacity = 0,
      group = "Subplot Grid",
      options = leaflet::pathOptions(interactive = FALSE)
    )
  }

  map <- map %>%
    # Local font stack (no internet required)
    htmlwidgets::prependContent(
      htmltools::tags$style(
        htmltools::HTML("
    .leaflet-container,
    .leaflet-popup-content,
    #sidebar {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica Neue',
                   Arial, 'Noto Sans', 'Liberation Sans', sans-serif, 'Apple Color Emoji',
                   'Segoe UI Emoji';
    }
    .taxon { font-style: italic; }
    .mapTitle { font-size: 22px; font-weight: 700; }
    .mapSubtitle { font-size: 16px; font-weight: 400; }
  ")
      )
    ) %>%
    htmlwidgets::prependContent(
      htmltools::tags$style("
      .leaflet-popup-content { font-size:14px; line-height:1.35; padding:6px 8px 8px; }
      .leaflet-popup-content-wrapper { border-radius:10px!important; box-shadow:0 3px 10px rgba(0,0,0,.25)!important; }
      .leaflet-popup-tip { display:none; }
      .leaflet-popup-content img { max-width:260px; width:100%; border-radius:6px; border:1px solid #ddd;
                                      box-shadow:0 1px 4px rgba(0,0,0,.15); margin-top:4px; }
      .hasPhotoBadge { display:inline-block; background:#388e3c; color:#fff; font-size:12px;
                                      padding:2px 6px; border-radius:12px; margin-bottom:4px; font-weight:500; }
      .prev, .next { color:#006699; font-size:22px; transition:color .2s; }
      .prev:hover, .next:hover { color:#004466; }
    ")
    )

  filename <- paste0("Results_", format(Sys.time(), "%d%b%Y_"), filename)
  output_path <- paste0(filename, ".html")
  message("Saving HTML map: '", output_path, "'")
  htmlwidgets::saveWidget(map, file = output_path, selfcontained = TRUE)
  # End function
}

#_______________________________________________________________________________
# Auxiliary function for coordinate interpolation ####
.get_latlon <- function(x, y, p1, p2, p3, p4) {
  left <- geosphere::destPoint(p1, geosphere::bearing(p1, p2), y)
  right <- geosphere::destPoint(p3, geosphere::bearing(p3, p4), y)
  interp <- geosphere::destPoint(left, geosphere::bearing(left, right), x)
  return(interp[1, c("lat", "lon")])
}

# Offset-based interpolation from center (used by MONITORA)
.get_latlon_from_center <- function(x, y, center_lonlat) {
  p1 <- geosphere::destPoint(center_lonlat, ifelse(x >= 0, 90, 270), abs(x))
  p2 <- geosphere::destPoint(c(p1[1,"lon"], p1[1,"lat"]), ifelse(y >= 0, 0, 180), abs(y))
  return(c(lat = p2[1, "lat"], lon = p2[1, "lon"]))
}

#_______________________________________________________________________________
# Given a plot census data frame and a set of herbaria, returns
.herbaria_lookup_links <- function(fp_df,
                                   herbariums,
                                   cache_rds = "jabot_index.rds",
                                   force_refresh = FALSE,
                                   verbose = FALSE,
                                   collector_fallback = NULL) {

  if (is.null(herbariums) || !length(herbariums)) {
    stop(
      "When `herbaria_lookup = TRUE`, you must provide at least one herbarium in `herbaria`, ",
      "for example: herbaria = c('RB','UPCB').",
      call. = FALSE
    )
  }
  herbariums <- toupper(herbariums)

  ## ---------------------- internal helpers ---------------------- ##

  .norm_txt <- function(x) {
    x <- tolower(iconv(x, to = "ASCII//TRANSLIT"))
    x <- gsub("\\s+", " ", x)
    trimws(x)
  }

  .norm_num <- function(x) {
    x <- gsub("[^0-9]", "", as.character(x))
    x[!nzchar(x)] <- NA_character_
    x
  }

  .strip_particles <- function(tokens) {
    parts <- c("de","da","do","das","dos","del","della","van","von")
    tokens[!(tokens %in% parts)]
  }

  .first_person <- function(rec_by_raw) {
    if (is.na(rec_by_raw) || !nzchar(rec_by_raw)) return(NA_character_)
    # split on "&", "e", "and"
    s <- gsub("\\s*&\\s*|\\s+e\\s+|\\s+and\\s+", ";", rec_by_raw, ignore.case = TRUE)
    p <- trimws(unlist(strsplit(s, ";")))
    p <- p[nzchar(p)]
    if (!length(p)) return(NA_character_)
    p[1]
  }

  .lastname_from_freeform <- function(name_raw) {
    if (is.na(name_raw) || !nzchar(name_raw)) return(NA_character_)
    n <- .norm_txt(name_raw)

    # "Lastname, Firstname" style
    if (grepl(",", n, fixed = TRUE)) {
      last <- trimws(strsplit(n, ",\\s*")[[1]][1])
      last <- .strip_particles(unlist(strsplit(last, "\\s+")))
      if (!length(last)) return(NA_character_)
      return(last[length(last)])
    }

    # Remove isolated initials (e.g. "A.", "J.F.")
    n <- gsub("\\b[a-z]\\.?\\b", "", n)
    toks <- .strip_particles(unlist(strsplit(n, "\\s+")))
    toks <- toks[nzchar(toks)]
    if (!length(toks)) return(NA_character_)
    toks[length(toks)]
  }

  .voucher_from_any_census <- function(fp_df, collector_fallback = NULL) {
    # Uses the combined Voucher field when available (e.g., "H. Medeiros 1234" or "GCO1095").
    v <- as.character(fp_df$Voucher)
    v[is.na(v)] <- ""

    # Split into trailing numeric part and prefix text (collector)
    num_from_voucher <- .norm_num(gsub(".*?(\\d[\\d\\./-]*)\\s*$", "\\1", v, perl = TRUE))
    col_from_voucher <- trimws(gsub("(.*?)(\\d[\\d\\./-]*\\s*)$", "\\1", v, perl = TRUE))
    col_from_voucher <- gsub("[\\.,;]+\\s*$", "", col_from_voucher)

    # Detect compact vouchers like "GCO1095": if fallback is provided, use it
    if (!is.null(collector_fallback) && nzchar(trimws(collector_fallback))) {
      is_compact <- grepl(
        "^[A-Za-z\\.]+\\s*\\d+$|^[A-Za-z]+\\d+$",
        gsub("\\s+", "", v)
      )
      col_from_voucher[is_compact] <- collector_fallback
    }

    # Fallback to census columns if needed
    .norm_name <- function(x) {
      tolower(gsub("\\s+", " ", iconv(x, to = "ASCII//TRANSLIT")))
    }
    .find_cols <- function(df, pats) {
      nms <- names(df)
      low <- .norm_name(nms)
      pat <- paste0("(", paste0(pats, collapse = "|"), ")")
      nms[grepl(pat, low)]
    }

    rec_cols <- .find_cols(
      fp_df,
      c("recorded\\s*by","collector","coletor","col\\.?", "coletora")
    )
    num_cols <- .find_cols(
      fp_df,
      c("record\\s*number","collector\\s*number","voucher\\s*number",
        "\\bno?\\.?\\s*(coletor|collector)","\\bnumero\\b","num\\b")
    )

    .first_non_empty <- function(df, cols) {
      if (!length(cols)) return(rep(NA_character_, nrow(df)))
      out <- rep(NA_character_, nrow(df))
      for (cl in cols) {
        v <- as.character(df[[cl]])
        v[is.na(v)] <- ""
        take <- !nzchar(out) & nzchar(trimws(v))
        out[take] <- trimws(v[take])
        if (all(nzchar(out))) break
      }
      out[!nzchar(out)] <- NA_character_
      out
    }

    rec_any <- .first_non_empty(fp_df, rec_cols)
    num_any <- .first_non_empty(fp_df, num_cols)

    collector <- ifelse(nzchar(col_from_voucher), col_from_voucher, rec_any)
    number <- ifelse(nzchar(num_from_voucher), num_from_voucher, num_any)

    list(collector = collector, number = number)
  }

  .read_dwca_min <- function(base_dir) {
    occ <- file.path(base_dir, "occurrence.txt")
    if (!file.exists(occ)) return(NULL)
    suppressWarnings(
      data.table::fread(
        occ,
        sep = "\t",
        select = c("catalogNumber","recordNumber","recordedBy"),
        na.strings = c("","NA"),
        showProgress = FALSE,
        quote = ""
      )
    )
  }

  .find_dwca_dir <- function(herbarium) {
    herbarium <- tolower(herbarium)
    jd <- "jabot_download"
    if (!dir.exists(jd)) return(NULL)
    subdirs <- list.dirs(jd, full.names = TRUE, recursive = FALSE)
    hits <- grep(paste0("dwca_.*_", herbarium, "_"), basename(subdirs), value = TRUE)
    if (!length(hits)) return(NULL)
    info <- file.info(file.path(jd, hits))
    rownames(info)[which.max(info$mtime)]
  }

  .build_herbaria_index_fast <- function(herbariums,
                                         needed_numbers = character(0),
                                         needed_lastnames = character(0),
                                         verbose = FALSE) {

    if (is.null(herbariums) || !length(herbariums)) {
      stop(
        "Argument `herbariums` must be a non-empty character vector (e.g. c('RB','UPCB')).",
        call. = FALSE
      )
    }

    if (!length(needed_numbers)) {
      if (verbose) message("herbaria: no record numbers found in input – skipping index build.")
      return(data.frame(
        key = character(0),
        catalogNumber = character(0),
        herbarium = character(0)
      ))
    }

    out <- list()
    for (h in toupper(herbariums)) {
      # Ensure we have a local DwC-A directory for this herbarium
      dir <- .find_dwca_dir(h)
      if (is.null(dir)) {
        if (verbose) message("Downloading DwC-A archive for ", h, "…")
        jabotR::jabot_download(herbarium = h, verbose = verbose)
        dir <- .find_dwca_dir(h)
        if (is.null(dir)) next
      }
      if (verbose) message("Reading minimal DwC-A occurrence table for ", h, "…")
      dt <- .read_dwca_min(dir)
      if (is.null(dt) || !nrow(dt)) next

      # Filter by needed record numbers
      dt[, rn := .norm_num(recordNumber)]
      dt <- dt[rn %in% needed_numbers]
      if (!nrow(dt)) next

      # Extract first collector and last name
      dt[, fp := vapply(recordedBy, .first_person, FUN.VALUE = character(1))]
      dt[, ln := vapply(fp, .lastname_from_freeform, FUN.VALUE = character(1))]
      dt <- dt[nzchar(ln)]
      if (length(needed_lastnames)) dt <- dt[ln %in% needed_lastnames]
      if (!nrow(dt)) next

      # Build key "lastname||number" and deduplicate
      dt[, key := paste(ln, rn, sep = "||")]
      dt <- dt[!duplicated(key)]
      out[[h]] <- data.frame(
        key = dt$key,
        catalogNumber = dt$catalogNumber,
        herbarium = h,
        stringsAsFactors = FALSE
      )
    }

    if (!length(out)) {
      return(data.frame(
        key = character(0),
        catalogNumber = character(0),
        herbarium = character(0)
      ))
    }

    do.call(rbind, out)
  }

  .get_or_build_herbaria_index <- function(fp_df,
                                           herbariums,
                                           cache_rds = "jabot_index.rds",
                                           force_refresh = FALSE,
                                           verbose = FALSE,
                                           collector_fallback = NULL) {

    if (is.null(herbariums) || !length(herbariums)) {
      stop(
        "Argument `herbariums` must be a non-empty character vector when running herbaria lookup.",
        call. = FALSE
      )
    }

    # Build the set of needed keys "lastname||number"
    vv <- .voucher_from_any_census(fp_df, collector_fallback = collector_fallback)
    last <- vapply(vv$collector, .lastname_from_freeform, FUN.VALUE = character(1))
    num <- .norm_num(vv$number)
    need_pairs <- unique(na.omit(
      ifelse(nzchar(last) & nzchar(num), paste(last, num, sep = "||"), NA_character_)
    ))

    if (!length(need_pairs)) {
      return(data.frame(
        key = character(0),
        catalogNumber = character(0),
        herbarium = character(0)
      ))
    }

    # Try cache
    idx <- NULL
    if (!force_refresh && file.exists(cache_rds)) {
      idx <- tryCatch(readRDS(cache_rds), error = function(e) NULL)
      if (!is.data.frame(idx) ||
          !all(c("key","catalogNumber","herbarium") %in% names(idx))) {
        idx <- NULL
      }
    }

    # Look for missing keys
    missing_pairs <- if (is.null(idx)) need_pairs else setdiff(need_pairs, idx$key)
    if (length(missing_pairs)) {
      miss_last <- unique(sub("\\|\\|.*$", "", missing_pairs))
      miss_num <- unique(sub("^.*\\|\\|", "", missing_pairs))

      inc <- .build_herbaria_index_fast(
        herbariums = herbariums,
        needed_numbers = miss_num,
        needed_lastnames = miss_last,
        verbose = verbose
      )
      if (is.null(idx)) {
        idx <- inc
      } else if (nrow(inc)) {
        idx <- unique(rbind(idx, inc))
      }

      if (nrow(idx)) {
        try(saveRDS(idx, cache_rds), silent = TRUE)
      }
    }

    idx
  }

  .make_herbaria_url <- function(catalogNumber, herbarium) {
    if (is.na(catalogNumber) || !nzchar(catalogNumber)) return(NA_character_)
    herbarium <- toupper(herbarium)
    digits <- gsub("\\D", "", catalogNumber)
    if (!nzchar(digits)) return(NA_character_)
    pad8 <- sprintf("%08d", as.integer(digits))
    if (herbarium == "RB") {
      paste0("https://rb.jbrj.gov.br/v2/regua/visualizador.php?r=true&colbot=RB",
             "&codtestemunho=", catalogNumber, "&arquivo=", pad8, ".dzi")
    } else if (herbarium == "UPCB") {
      paste0("https://rb.jbrj.gov.br/v2/regua/visualizador.php?r=true&colbot=UPCB",
             "&codtestemunho=", catalogNumber, "&arquivo=", "upcb", pad8, ".dzi")
    } else {
      NA_character_
    }
  }

  .link_for_voucher_vec <- function(collector, number, herbaria_index_df) {
    if (is.null(herbaria_index_df) || !nrow(herbaria_index_df)) {
      return(rep(NA_character_, length(collector)))
    }
    out <- rep(NA_character_, length(collector))
    last <- vapply(collector, .lastname_from_freeform, FUN.VALUE = character(1))
    num <- .norm_num(number)
    k <- ifelse(
      !is.na(last) & nzchar(last) & !is.na(num) & nzchar(num),
      paste(last, num, sep = "||"),
      NA_character_
    )
    pos <- match(k, herbaria_index_df$key)
    has <- !is.na(pos)
    out[has] <- mapply(
      .make_herbaria_url,
      herbaria_index_df$catalogNumber[pos[has]],
      herbaria_index_df$herbarium[pos[has]]
    )
    out
  }

  ## ---------------------- main pipeline ---------------------- ##

  herbaria_index <- .get_or_build_herbaria_index(
    fp_df = fp_df,
    herbariums = herbariums,
    cache_rds = cache_rds,
    force_refresh = force_refresh,
    verbose = verbose,
    collector_fallback = collector_fallback
  )

  vv <- .voucher_from_any_census(fp_df, collector_fallback = collector_fallback)

  .link_for_voucher_vec(vv$collector, vv$number, herbaria_index)
}
