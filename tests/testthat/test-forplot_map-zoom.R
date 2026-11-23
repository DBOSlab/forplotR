# Zoom, subplot labels, marker size and report links of forplot_map()

# Widget data (the htmlwidgets JSON) of a saved forplot_map() page
.zoom_widget <- function(path) {
  h <- paste(readLines(path, warn = FALSE, encoding = "UTF-8"), collapse = "\n")
  m <- regmatches(h, regexpr(
    '<script type="application/json" data-for="htmlwidget-[^"]+">.*?</script>', h, perl = TRUE))
  json <- sub("^<script[^>]*>", "", sub("</script>$", "", m))
  obj <- jsonlite::fromJSON(json, simplifyVector = FALSE)
  list(html = h, x = obj$x, data = obj$jsHooks$render[[1]]$data)
}

.zoom_calls <- function(w, method) {
  Filter(function(cl) identical(cl$method, method), w$x$calls)
}

.zoom_rus01_map <- function(...) {
  f <- file.path(tempdir(), "zoom_RUS01.xlsx")
  if (!file.exists(f)) openxlsx::write.xlsx(RUS01, f)
  corners <- data.frame(Latitude = c(-19.9000, -19.9009, -19.9000, -19.9009),
                        Longitude = c(-40.5500, -40.5500, -40.5490, -40.5490))
  forplot_map(f, vertex_coords = corners, dir = file.path(tempdir(), "zoom_maps"),
              filename = "rus01", verbose = FALSE, ...)
}

.zoom_tum_map <- function(...) {
  f <- file.path(tempdir(), "zoom_TUM.xlsx")
  if (!file.exists(f)) openxlsx::write.xlsx(TUM, f)
  forplot_map(f, input_type = "monitora", station_name = "EA-1",
              vertex_coords = data.frame(Latitude = 1.5, Longitude = -52.5),
              dir = file.path(tempdir(), "zoom_maps"), filename = "tum", verbose = FALSE, ...)
}

# Each label must sit on the centre of one subplot polygon of the grid
.expect_labels_on_cells <- function(w, n_cells) {
  # addLabelOnlyMarkers() is stored as an addMarkers call
  lab <- .zoom_calls(w, "addMarkers")[[1]]$args
  lat <- unlist(lab[[1]]); lng <- unlist(lab[[2]])
  expect_length(lat, n_cells)

  grid <- Filter(function(cl) identical(cl$args[[3]], "Subplot Grid"),
                 .zoom_calls(w, "addPolygons"))
  expect_length(grid, n_cells)
  centres <- t(vapply(grid, function(cl) {
    ring <- cl$args[[1]][[1]][[1]][[1]]
    c(mean(unlist(ring$lat)[1:4]), mean(unlist(ring$lng)[1:4]))
  }, numeric(2)))

  # nearest cell centre, in metres
  d <- vapply(seq_along(lat), function(i) {
    min(sqrt(((centres[, 1] - lat[i]) * 110574)^2 +
             ((centres[, 2] - lng[i]) * 111320 * cos(lat[i] * pi / 180))^2))
  }, numeric(1))
  expect_lt(max(d), 0.5)
  # one label per cell
  nearest <- vapply(seq_along(lat), function(i) {
    which.min((centres[, 1] - lat[i])^2 + (centres[, 2] - lng[i])^2)
  }, integer(1))
  expect_setequal(nearest, seq_len(n_cells))
}

test_that("the map opens no closer than the initial zoom limit", {
  skip_if_not_installed("openxlsx")
  skip_on_cran()
  w <- .zoom_widget(.zoom_rus01_map())
  z <- .fp_map_init_max_zoom
  expect_identical(z, 17L)

  fit <- w$x$fitBounds
  expect_identical(fit[[5]]$maxZoom, z)

  sat <- Filter(function(cl) identical(cl$args[[3]], "Satellite"), .zoom_calls(w, "addProviderTiles"))
  expect_length(sat, 1L)
  expect_identical(sat[[1]]$args[[4]]$maxNativeZoom, z)

  # the "back to plot" button uses the same limit
  expect_match(w$html, sprintf("maxZoom:%d", z), fixed = TRUE)
  # and the page is told where to look for deeper Esri imagery
  im <- w$data$imagery
  expect_identical(im$nativeZoom, z)
  expect_gt(im$maxProbeZoom, z)
})

test_that("subplot labels sit on their cells in the square layout", {
  skip_if_not_installed("openxlsx")
  skip_on_cran()
  w <- .zoom_widget(.zoom_rus01_map())
  .expect_labels_on_cells(w, 100L)
  grid <- w$data$grid
  expect_equal(grid$cellMetres, 10)
  expect_equal(grid$minLabelPx, .fp_map_min_label_px)
})

test_that("subplot labels sit on their cells in the MONITORA cross", {
  skip_if_not_installed("openxlsx")
  skip_on_cran()
  w <- .zoom_widget(.zoom_tum_map())
  .expect_labels_on_cells(w, 40L)
  lab <- unlist(.zoom_calls(w, "addMarkers")[[1]]$args[[11]])
  expect_setequal(lab, paste0(rep(c("N", "S", "L", "O"), each = 10), 1:10))
})

test_that("embedded stylesheets keep data URIs on one line", {
  f <- tempfile(fileext = ".png")
  writeBin(as.raw(sample(0:255, 400, replace = TRUE)), f)
  uri <- .file_data_uri(f)
  expect_false(grepl("[\r\n]", uri))
  css <- .inline_css_urls(sprintf(".a{background:url(%s)}\n.leaflet-tooltip{position:absolute}",
                                  basename(f)), dirname(f))
  expect_match(css, '^\\.a\\{background:url\\("data:image/png;base64,[A-Za-z0-9+/=]+"\\)\\}\n')
  expect_match(css, ".leaflet-tooltip{position:absolute}", fixed = TRUE)
})

test_that("saved maps carry the whole leaflet stylesheet", {
  skip_if_not_installed("openxlsx")
  skip_on_cran()
  h <- .zoom_widget(.zoom_rus01_map())$html
  uris <- regmatches(h, gregexpr('url\\("data:[^"]*"\\)', h))[[1]]
  expect_false(any(grepl("[\r\n]", uris)))
  expect_match(h, ".leaflet-tooltip {", fixed = TRUE)
})

test_that("marker radius follows the zoom within the pixel limits", {
  sz <- .fp_map_marker_size
  rel <- .dbh_radius(c(5, 20, 80))
  # far out: everyone at the minimum, still clickable
  expect_equal(.fp_marker_px(rel, 0.1), rep(sz[["min_px"]], 3))
  # far in: capped
  expect_equal(.fp_marker_px(rel, 1000), rep(sz[["max_px"]], 3))
  # in between: radius in metres, proportional to the DBH scale
  px_per_m <- 7
  r <- .fp_marker_px(rel, px_per_m)
  expect_equal(r, rel * sz[["metres"]] * px_per_m)
  expect_equal(r / r[1], rel / rel[1])
  # doubling the zoom doubles the radius while within the limits
  expect_equal(.fp_marker_px(rel, 2 * px_per_m), 2 * r)
})

test_that("the page receives the marker size settings", {
  skip_if_not_installed("openxlsx")
  skip_on_cran()
  w <- .zoom_widget(.zoom_rus01_map())
  mk <- w$data$markers
  expect_equal(unlist(mk), .fp_map_marker_size)
  expect_match(w$html, "map.on('zoomend', resize)", fixed = TRUE)
})

test_that("report_url adds a link back to the report in the map", {
  skip_if_not_installed("openxlsx")
  skip_on_cran()
  ctrl <- function(w) paste(unlist(.zoom_calls(w, "addControl")[[1]]$args[[1]]), collapse = "")
  with_link <- .zoom_widget(.zoom_rus01_map(report_url = "index.html", language = "pt"))
  expect_match(ctrl(with_link), "class='fp-report-link' href='index.html'", fixed = TRUE)
  expect_match(ctrl(with_link), "Voltar ao relatório", fixed = TRUE)
  without <- .zoom_widget(.zoom_rus01_map())
  expect_false(grepl("fp-report-link", ctrl(without), fixed = TRUE))
})

test_that("map_url adds the map button to the field report", {
  skip_if_not_installed("openxlsx")
  skip_on_cran()
  f <- file.path(tempdir(), "zoom_RUS01.xlsx")
  if (!file.exists(f)) openxlsx::write.xlsx(RUS01, f)
  run <- function(...) {
    res <- forplot_balance(f, render_pdf = FALSE, write_xlsx = FALSE, verbose = FALSE,
                           dir = file.path(tempdir(), "zoom_report"), ...)
    paste(readLines(res$html_report, warn = FALSE, encoding = "UTF-8"), collapse = "\n")
  }
  h <- run(map_url = "map.html", filename = "with_map")
  expect_match(h, '<p class="map-cta"><a href="map.html">Open the interactive map', fixed = TRUE)
  expect_match(h, '<a href="map.html" class="to-map">', fixed = TRUE)
  h0 <- run(filename = "no_map")
  expect_false(grepl("map-cta\"><a", h0, fixed = TRUE))
  expect_false(grepl("class=\"to-map\"", h0, fixed = TRUE))
})

test_that("link arguments are validated", {
  expect_null(.arg_check_link_url(NULL, "map_url"))
  expect_identical(.arg_check_link_url(" map.html ", "map_url"), "map.html")
  for (bad in list("", NA_character_, c("a", "b"), 1)) {
    expect_error(.arg_check_link_url(bad, "map_url"), "`map_url` must be NULL")
  }
  expect_error(forplot_balance(map_url = ""), "`map_url` must be NULL")
  expect_error(forplot_map(report_url = 3), "`report_url` must be NULL")
})
