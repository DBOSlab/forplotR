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
#'   \item harmonizes the data into a canonical internal schema;
#'   \item extracts plot metadata such as team, plot name, and plot code;
#'   \item converts local subplot coordinates to geographic coordinates using
#'   either plot vertices or a MONITORA center point;
#'   \item builds plot and subplot overlays;
#'   \item colors specimen markers by collection status, with palms shown as a
#'   separate visual category;
#'   \item optionally embeds voucher image carousels in specimen popups;
#'   \item optionally performs a local herbarium lookup using internal helper
#'   functions and adds herbarium image links when available; and
#'   \item saves the final map as an HTML file.
#' }
#'
#' Collection status is inferred from the canonical \code{Collected} column
#' after input harmonization. For some input types, this field is derived from
#' source-specific voucher or collection-status columns.
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
#' The function writes an HTML file to disk. When \code{herbaria_lookup = TRUE},
#' it may also build or refresh a local herbarium index from DwC-A downloads and
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
#' frame, a path to an Excel file, or a numeric vector of length 2.
#' For \code{"field_sheet"}, \code{"field_sheet_ti"}, and
#' \code{"fp_query_sheet"}, provide the four plot corners as latitude and
#' longitude. For \code{"monitora"}, provide a single central latitude/longitude.
#' A numeric vector of length 2 is interpreted as \code{c(lat, lon)} unless
#' named with \code{lat}/\code{latitude} and \code{lon}/\code{longitude}.
#'
#' @param plot_size Numeric scalar. Total plot size in hectares. Used for
#' non-MONITORA inputs.
#'
#' @param subplot_size Numeric scalar. Side length of each subplot in meters.
#'
#' @param voucher_imgs Character scalar or \code{NULL}. Directory containing
#' voucher image subfolders. Subdirectories are assumed to be named after
#' voucher IDs, each containing one or more image files to display in specimen
#' popups.
#'
#' @param filename Character scalar. Basename for the output HTML file, without
#' extension. The function prepends a date-stamped prefix before writing the file.
#'
#' @param station_name Character vector or \code{NULL}. For
#' \code{input_type = "monitora"}, optionally filter or iterate over one or more
#' station names. When a vector of length greater than 1 is supplied, the
#' function may process stations one by one via internal recursion.
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
#' herbarium lookup and related processing.
#'
#' @return Invisibly returns the path to the generated HTML file.
#'
#' @examples
#' \dontrun{
#' vertex_df <- data.frame(
#'   Latitude = c(-3.123, -3.123, -3.125, -3.125),
#'   Longitude = c(-60.012, -60.010, -60.012, -60.010)
#' )
#'
#' plot_html_map(
#'   fp_file_path = "data/tree_data.xlsx",
#'   input_type = "field_sheet",
#'   vertex_coords = vertex_df,
#'   voucher_imgs = "voucher_imgs",
#'   filename = "plot_map"
#' )
#'
#' plot_html_map(
#'   fp_file_path = "data/RUS_plot.xlsx",
#'   input_type = "field_sheet",
#'   vertex_coords = "data/vertices.xlsx",
#'   voucher_imgs = "voucher_imgs",
#'   collector = "G. C. Ottino",
#'   herbaria_lookup = TRUE,
#'   herbaria = "RB",
#'   filename = "rus_plot_map"
#' )
#'
#' plot_html_map(
#'   fp_file_path = "data/monitora.xlsx",
#'   input_type = "monitora",
#'   vertex_coords = c(lat = -3.10, lon = -60.02),
#'   station_name = "Station1",
#'   filename = "monitora_map"
#' )
#' }
#'
#' @importFrom leaflet leaflet addProviderTiles addPolygons addCircleMarkers
#' @importFrom leaflet addPolylines addLabelOnlyMarkers setView addControl
#' @importFrom htmltools HTML tags htmlEscape
#' @importFrom htmlwidgets onRender prependContent saveWidget
#' @importFrom geosphere destPoint bearing
#' @importFrom readxl read_excel
#' @importFrom here here
#' @importFrom dplyr select filter mutate case_when rename_with rename if_else
#' @importFrom magrittr "%>%"
#' @importFrom scales rescale
#' @importFrom openxlsx read.xlsx
#' @importFrom data.table fread :=
#' @importFrom stats na.omit
#' @importFrom DBI dbConnect dbDisconnect dbExecute dbGetQuery dbWriteTable
#' @importFrom duckdb duckdb
#'
#' @export

plot_html_map <- function(fp_file_path = NULL,
                          input_type = c("field_sheet", "field_sheet_ti", "fp_query_sheet", "monitora"),
                          vertex_coords = NULL,
                          plot_size = 1,
                          subplot_size = 10,
                          voucher_imgs = NULL,
                          filename = "plot_map",
                          station_name = NULL,
                          collector = NULL,
                          collector_map = NULL,
                          herbaria_lookup = FALSE,
                          herbaria = NULL,
                          herbaria_force_refresh = FALSE,
                          keep_herbaria_downloads = FALSE,
                          verbose = TRUE) {

  input_type <- tolower(trimws(as.character(input_type)))
  input_type <- match.arg(input_type, c("field_sheet", "field_sheet_ti", "fp_query_sheet", "monitora"))

  .validate_plot_size(plot_size)
  .validate_subplot_size(subplot_size)

  if (!is.character(fp_file_path) || length(fp_file_path) != 1L || !file.exists(fp_file_path)) {
    stop("The provided 'fp_file_path' does not exist.", call. = FALSE)
  }

  # Harmonize the plot input data type
  fp_loaded <- .harmonize_plot_input(
    fp_file_path = fp_file_path,
    input_type = input_type,
    station_name = station_name
  )

  fp_sheet  <- fp_loaded$fp_sheet
  team <- fp_loaded$team
  plot_name <- fp_loaded$plot_name
  plot_code <- fp_loaded$plot_code

  if (input_type == "monitora") {
    fp_clean <- fp_sheet %>%
      dplyr::mutate(
        T1 = as.character(T1),
        T2 = as.character(T2),
        X = .coerce_xy(X),
        Y = .coerce_xy(Y),
        D = .numify(D)
      ) %>%
      dplyr::filter(is.finite(X), is.finite(Y))

    if (!nrow(fp_clean)) {
      stop("No valid points to plot after cleaning MONITORA X/Y values.", call. = FALSE)
    }

    fp_coords <- .compute_monitora_geometry(
      fp_df = fp_clean,
      keep_only_cell = FALSE
    )

    if (!all(c("draw_x", "draw_y") %in% names(fp_coords))) {
      stop("MONITORA helper must add 'draw_x' and 'draw_y' columns (meters).", call. = FALSE)
    }

  } else {
    fp_clean <- fp_sheet %>%
      dplyr::mutate(
        T1 = .numify(T1),
        X = .coerce_xy(X),
        Y = .coerce_xy(Y),
        D = .numify(D)
      ) %>%
      dplyr::filter(is.finite(T1), is.finite(X), is.finite(Y))

    if (!nrow(fp_clean)) {
      stop("No valid points to plot: T1/X/Y could not be parsed from the input sheet.", call. = FALSE)
    }

    fp_coords <- .compute_global_coordinates_generic(
      fp_clean = fp_clean,
      plot_size = plot_size,
      subplot_size = subplot_size
    )
  }

  if (is.character(vertex_coords) && length(vertex_coords) == 1L && grepl("\\.xlsx?$", vertex_coords)) {
    vertex_coords <- readxl::read_excel(vertex_coords)
  }

  if (!is.null(vertex_coords) && !is.data.frame(vertex_coords)) {
    if (is.atomic(vertex_coords) && length(vertex_coords) == 2L) {
      nm <- tolower(names(vertex_coords))
      if (length(nm) == 2L && any(nm %in% c("lon", "longitude"))) {
        lon <- as.numeric(vertex_coords[which(nm %in% c("lon", "longitude"))][1])
        lat <- as.numeric(vertex_coords[which(nm %in% c("lat", "latitude"))][1])
      } else {
        lat <- as.numeric(vertex_coords[1])
        lon <- as.numeric(vertex_coords[2])
      }
      vertex_coords <- data.frame(latitude = lat, longitude = lon, check.names = FALSE)
    } else {
      stop(
        "vertex_coords must be a 1-row data.frame or a length-2 numeric vector (lat, lon).",
        call. = FALSE
      )
    }
  }

  vertex_coords <- vertex_coords %>%
    dplyr::rename_with(~tolower(gsub("\\s*\\(.*?\\)", "", .x)), dplyr::everything()) %>%
    dplyr::rename(latitude = latitude, longitude = longitude) %>%
    dplyr::mutate(
      longitude = as.numeric(gsub(",", ".", longitude)),
      latitude = as.numeric(gsub(",", ".", latitude))
    )

  if (input_type != "monitora" && nrow(vertex_coords) >= 4) {
    vertex_coords <- vertex_coords %>%
      dplyr::mutate(
        longitude = ifelse(longitude > 0, -longitude, longitude),
        latitude = ifelse(latitude > 0, -latitude, latitude)
      )
  }

  grid_polys <- list()
  grid_labels <- data.frame()

  if (input_type == "monitora") {
    if (nrow(vertex_coords) != 1L) {
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

    cells <- dplyr::bind_rows(
      .mk_cells("N") %>% dplyr::mutate(arm = "N"),
      .mk_cells("S") %>% dplyr::mutate(arm = "S"),
      .mk_cells("L") %>% dplyr::mutate(arm = "L"),
      .mk_cells("O") %>% dplyr::mutate(arm = "O")
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

    for (k in seq_len(nrow(cells))) {
      XY <- with(cells[k, ], rbind(
        c(xmin, ymin),
        c(xmax, ymin),
        c(xmax, ymax),
        c(xmin, ymax),
        c(xmin, ymin)
      ))
      llp <- t(apply(XY, 1, function(v) .get_latlon_from_center(v[1], v[2], center)))
      grid_polys[[k]] <- list(lng = llp[, 2], lat = llp[, 1])
    }

    lab_ll <- t(apply(
      as.matrix(centers[, c("cx", "cy")]),
      1,
      function(v) .get_latlon_from_center(v[1], v[2], center)
    ))

    grid_labels <- data.frame(
      subplot = paste0(centers$arm, centers$label),
      lat = lab_ll[, 1],
      lon = lab_ll[, 2]
    )

  } else {
    if (nrow(vertex_coords) < 4L) {
      stop(
        "For 'field_sheet' and 'fp_query_sheet' you must provide FOUR plot corners in 'vertex_coords'.",
        call. = FALSE
      )
    }

    p1 <- c(vertex_coords$longitude[1], vertex_coords$latitude[1])
    p2 <- c(vertex_coords$longitude[2], vertex_coords$latitude[2])
    p3 <- c(vertex_coords$longitude[3], vertex_coords$latitude[3])
    p4 <- c(vertex_coords$longitude[4], vertex_coords$latitude[4])

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

    fp_coords$Latitude <- suppressWarnings(as.numeric(coords_geo[1, , drop = TRUE]))
    fp_coords$Longitude <- suppressWarnings(as.numeric(coords_geo[2, , drop = TRUE]))

    keep <- is.finite(fp_coords$Latitude) & is.finite(fp_coords$Longitude)
    fp_coords <- fp_coords[keep, , drop = FALSE]

    if (!nrow(fp_coords)) {
      stop(
        "No valid points to plot: Latitude/Longitude are missing or non-numeric after conversion.",
        call. = FALSE
      )
    }

    total_area_m2 <- plot_size * 10000
    subplot_area_m2 <- subplot_size * subplot_size
    n_subplots <- as.integer(round(total_area_m2 / subplot_area_m2))
    n_cols <- max(1L, as.integer(round(100 / subplot_size)))
    n_rows <- ceiling(n_subplots / n_cols)

    subplot_index <- 1L
    for (row in 0:(n_rows - 1L)) {
      for (col in 0:(n_cols - 1L)) {
        if (subplot_index > n_subplots) next

        local_poly <- .get_subplot_polygon(col, row, subplot_size)
        center_x <- mean(local_poly[, 1])
        center_y <- mean(local_poly[, 2])

        coords_latlon <- t(mapply(
          .get_latlon,
          x = local_poly[, 1],
          y = local_poly[, 2],
          MoreArgs = list(p1 = p1, p2 = p2, p3 = p3, p4 = p4)
        ))

        grid_polys[[subplot_index]] <- list(
          lng = coords_latlon[, "lon"],
          lat = coords_latlon[, "lat"]
        )

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

        subplot_index <- subplot_index + 1L
      }
    }
  }

  fp_coords$color <- dplyr::case_when(
    fp_coords$Family == "Arecaceae" ~ "gold",
    !is.na(fp_coords$Collected) & fp_coords$Collected != "" ~ "gray",
    TRUE ~ "red"
  )

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
          warning(
            paste0("Herbarium lookup failed: ", conditionMessage(e)),
            call. = FALSE
          )
          rep(NA_character_, nrow(fp_coords))
        }
      )

      if (is.character(herbaria_links) && length(herbaria_links) == nrow(fp_coords)) {
        fp_coords$herbaria_link <- herbaria_links
      } else {
        warning(
          "Invalid result returned by herbarium lookup. Proceeding without herbarium links.",
          call. = FALSE
        )
      }
    }
  }

  fp_coords$popup <- paste0(
    "<b>Tag:</b> ", fp_coords[["New Tag No"]], "<br/>",
    "<b>Subplot:</b> ", fp_coords$T1, "<br/>",
    "<b>Family:</b> ", fp_coords$Family, "<br/>",
    "<b>Species:</b> <i class='taxon'>", fp_coords[["Original determination"]], "</i><br/>",
    "<b>DBH (cm):</b> ", round(fp_coords$D/10, 2), "<br/>",
    "<b>Voucher:</b> ",
    ifelse(
      is.na(fp_coords$Voucher) | !nzchar(fp_coords$Voucher),
      "Unvouchered",
      fp_coords$Voucher
    ),
    ifelse(
      !is.na(fp_coords$herbaria_link) &
        nzchar(as.character(fp_coords$herbaria_link)),
      as.character(fp_coords$herbaria_link),
      ""
    )
  )

  plot_popup <- paste0(
    "<b>Plot Name:</b> ", plot_name, "<br/>",
    "<b>Plot Code:</b> ", plot_code, "<br/>",
    "<b>Team:</b> ", team
  )

  original_image_base <- here::here(voucher_imgs)
  leaf_dirs <- list.dirs(original_image_base, recursive = TRUE, full.names = TRUE)

  leaf_dirs <- leaf_dirs[!sapply(leaf_dirs, function(dir) {
    any(file.info(list.dirs(dir, recursive = FALSE))$isdir)
  })]

  leaf_dirs <- sub(paste0(".*(", voucher_imgs, "/.*)"), "\\1", leaf_dirs)

  handled <- character(0)
  fp_coords$popup <- as.character(fp_coords$popup)
  fp_coords$Voucher <- as.character(fp_coords$Voucher)

  for (dir_path in leaf_dirs) {
    voucher_id <- basename(dir_path)

    if (voucher_id %in% handled) next
    handled <- c(handled, voucher_id)

    img_files <- list.files(
      dir_path,
      full.names = TRUE,
      pattern = "\\.(jpg|jpeg|png|gif)$",
      ignore.case = TRUE
    )
    if (!length(img_files)) next

    rel_paths <- file.path(dir_path, basename(img_files))

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

  initial_zoom <- 18
  lat0 <- mean(vertex_coords$latitude)
  lon0 <- mean(vertex_coords$longitude)

  map <- leaflet::leaflet(
    fp_coords,
    options = leaflet::leafletOptions(minZoom = 2, maxZoom = 22)
  ) %>%
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
    leaflet::addProviderTiles(leaflet::providers$Esri.WorldImagery, group = "Satellite") %>%
    leaflet::addProviderTiles(leaflet::providers$OpenTopoMap, group = "OpenTopo") %>%
    leaflet::addProviderTiles(leaflet::providers$CartoDB.DarkMatter, group = "Dark Matter") %>%
    leaflet::addLayersControl(
      baseGroups = c("OSM Street", "ESRI Street", "Satellite", "OpenTopo", "Dark Matter"),
      overlayGroups = c("Specimens", "Subplot Grid"),
      position = "bottomleft",
      options = leaflet::layersControlOptions(
        collapsed = TRUE,
        autoZIndex = TRUE
      )
    ) %>%
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

  if (input_type != "monitora") {
    map <- map %>%
      leaflet::addPolygons(
        lng = c(p1[1], p2[1], p4[1], p3[1], p1[1]),
        lat = c(p1[2], p2[2], p4[2], p3[2], p1[2]),
        color = "darkgreen",
        fillColor = "lightgreen",
        fillOpacity = 0.15,
        popup = plot_popup,
        options = leaflet::pathOptions(interactive = FALSE)
      )
  }

  map <- map %>%
    leaflet::addCircleMarkers(
      lng = ~Longitude,
      lat = ~Latitude,
      radius = scales::rescale(
        fp_coords$D,
        to = c(2, 8),
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

    eo <- as.matrix(eo)
    ns <- as.matrix(ns)

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

  for (i in seq_along(grid_polys)) {
    map <- map %>%
      leaflet::addPolygons(
        lng = grid_polys[[i]]$lng,
        lat = grid_polys[[i]]$lat,
        color = "#999999",
        weight = 0.5,
        fillOpacity = 0,
        group = "Subplot Grid",
        options = leaflet::pathOptions(interactive = FALSE)
      )
  }

  map <- map %>%
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
}

.validate_plot_size <- function(plot_size) {
  ok <- is.numeric(plot_size) && length(plot_size) == 1L && is.finite(plot_size) && plot_size > 0
  if (!ok) {
    stop("`plot_size` must be a single positive numeric value in hectares.", call. = FALSE)
  }
  invisible(TRUE)
}

.validate_subplot_size <- function(subplot_size) {
  ok <- is.numeric(subplot_size) && length(subplot_size) == 1L && is.finite(subplot_size) && subplot_size > 0
  if (!ok) {
    stop("`subplot_size` must be a single positive numeric value in meters.", call. = FALSE)
  }
  invisible(TRUE)
}

.compute_global_coordinates_generic <- function(fp_clean, plot_size, subplot_size) {
  total_area_m2 <- plot_size * 10000
  subplot_area_m2 <- subplot_size * subplot_size
  n_subplots <- total_area_m2 / subplot_area_m2

  if (!isTRUE(all.equal(n_subplots, round(n_subplots)))) {
    stop(
      "The combination of `plot_size` and `subplot_size` does not yield an integer number of subplots.",
      call. = FALSE
    )
  }

  n_subplots <- as.integer(round(n_subplots))
  n_cols <- max(1L, as.integer(round(100 / subplot_size)))
  n_rows <- ceiling(n_subplots / n_cols)

  fp_clean %>%
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

.get_latlon <- function(x, y, p1, p2, p3, p4) {
  left <- geosphere::destPoint(p1, geosphere::bearing(p1, p2), y)
  right <- geosphere::destPoint(p3, geosphere::bearing(p3, p4), y)
  interp <- geosphere::destPoint(left, geosphere::bearing(left, right), x)
  interp[1, c("lat", "lon")]
}

.get_latlon_from_center <- function(x, y, center_lonlat) {
  p1 <- geosphere::destPoint(center_lonlat, ifelse(x >= 0, 90, 270), abs(x))
  p2 <- geosphere::destPoint(c(p1[1, "lon"], p1[1, "lat"]), ifelse(y >= 0, 0, 180), abs(y))
  c(lat = p2[1, "lat"], lon = p2[1, "lon"])
}

.numify <- function(z) {
  if (is.numeric(z)) return(as.numeric(z))

  zc <- as.character(z)
  zc[is.na(zc)] <- ""
  zc <- trimws(zc)
  zc <- gsub("\\s+", "", zc)
  zc <- gsub("\\.(?=\\d{3}(\\D|$))", "", zc, perl = TRUE)
  zc <- gsub(",", ".", zc, fixed = TRUE)
  suppressWarnings(as.numeric(zc))
}

.coerce_xy <- function(z) {
  if (is.numeric(z)) return(as.numeric(z))

  zc <- as.character(z)
  zc[is.na(zc)] <- ""
  zc <- trimws(zc)
  zc <- gsub("\\s+", "", zc)
  zc <- sub("^([^;/|]+)[;/|].*$", "\\1", zc, perl = TRUE)
  zc <- gsub("\\.(?=\\d{3}(\\D|$))", "", zc, perl = TRUE)
  zc <- gsub(",", ".", zc, fixed = TRUE)

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

  if (arm == "N") {
    dplyr::mutate(
      base,
      xmin = -cs + r * cs, xmax = -cs + r * cs + cs,
      ymin = 50 + c * cs, ymax = 50 + c * cs + cs
    )
  } else if (arm == "S") {
    dplyr::mutate(
      base,
      xmin = -cs + r * cs, xmax = -cs + r * cs + cs,
      ymin = -100 + c * cs, ymax = -100 + c * cs + cs
    )
  } else if (arm == "L") {
    dplyr::mutate(
      base,
      xmin = 50 + c * cs, xmax = 50 + c * cs + cs,
      ymin = -cs + r * cs, ymax = -cs + r * cs + cs
    )
  } else {
    dplyr::mutate(
      base,
      xmin = -100 + c * cs, xmax = -100 + c * cs + cs,
      ymin = -cs + r * cs, ymax = -cs + r * cs + cs
    )
  }
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

