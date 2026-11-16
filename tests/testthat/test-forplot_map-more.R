# Additional forplot_map() tests: argument checks, vertex parsing, MONITORA
# stations and voucher image helpers. Offline; the HTML writer is mocked.

.mm_input <- function(env = parent.frame()) {
  f <- withr::local_tempfile(fileext = ".xlsx", .local_envir = env)
  file.create(f)
  f
}

.mm_vertices <- function() {
  data.frame(
    Latitude = c(-3.00000, -2.99991, -3.00000, -2.99991),
    Longitude = c(-60.00000, -60.00000, -59.99991, -59.99991)
  )
}

.mm_fp_loaded <- function(...) {
  list(
    fp_sheet = data.frame(
      T1 = c(1, 1, NA), X = c(1, 4, 8), Y = c(1, 5, 8), D = c(100, 180, 250),
      Family = c("Fabaceae", "Arecaceae", "Myrtaceae"),
      Collected = c("C001", NA, ""), `New Tag No` = c("1", "2", "3"),
      `Original determination` = c("Inga alba", "Euterpe precatoria", "Eugenia sp."),
      Voucher = c("V1", "", NA),
      check.names = FALSE, stringsAsFactors = FALSE
    ),
    team = "Team", plot_name = "Plot", plot_code = "P01", census_no_fp = "1"
  )
}

.mm_monitora_loaded <- function(...) {
  list(
    fp_sheet = data.frame(
      T1 = c("N1", "N1"), T2 = c("1", "1"), X = c("1,5", "5"), Y = c("2,5", "7"),
      D = c("100", "200"), Family = c("Fabaceae", "Arecaceae"),
      Collected = c("C001", NA), `New Tag No` = c("10", "11"),
      `Original determination` = c("Inga alba", "Euterpe precatoria"),
      Voucher = c("V10", ""), check.names = FALSE, stringsAsFactors = FALSE
    ),
    team = "", plot_name = "MON", plot_code = "", census_no_fp = "1"
  )
}

.mm_mock_save <- function(env = parent.frame()) {
  saved <- new.env(parent = emptyenv())
  saved$files <- character(0)
  testthat::local_mocked_bindings(
    .save_widget_selfcontained = function(widget, file, title = NULL) {
      saved$files <- c(saved$files, file)
      saved$widget <- widget
      writeLines("<html></html>", file)
      invisible(NULL)
    },
    .env = env
  )
  saved
}


# Argument checks --------------------------------------------------------------

test_that("forplot_map validates logical flags and filename", {
  f <- .mm_input()
  for (flag in c("herbaria_lookup", "verbose", "highlight_palms",
                 "embed_voucher_imgs")) {
    args <- list(fp_file_path = f, vertex_coords = .mm_vertices())
    args[[flag]] <- NA
    expect_error(do.call(forplot_map, args), paste0("`", flag, "` must be TRUE or FALSE"))
  }
  expect_error(forplot_map(f, vertex_coords = .mm_vertices(), filename = ""),
               "`filename` must be a single non-empty")
  expect_error(forplot_map(f, vertex_coords = .mm_vertices(), filename = c("a", "b")),
               "`filename` must be a single non-empty")
})

test_that("MONITORA-only column mappings warn for other inputs", {
  saved <- .mm_mock_save()
  testthat::local_mocked_bindings(.harmonize_plot_input = .mm_fp_loaded)
  expect_warning(
    suppressMessages(forplot_map(
      .mm_input(), input_type = "field_sheet", vertex_coords = .mm_vertices(),
      plot_size = 0.01, subplot_size = 10, voucher_imgs = NULL,
      colname_cap = "CAP", dir = withr::local_tempdir()
    )),
    "only used when input_type = 'monitora'"
  )
})

test_that("verbose mode reports rows dropped for missing T1/X/Y", {
  saved <- .mm_mock_save()
  testthat::local_mocked_bindings(.harmonize_plot_input = .mm_fp_loaded)
  out_dir <- withr::local_tempdir()
  msgs <- testthat::capture_messages(
    out <- forplot_map(
      .mm_input(), vertex_coords = .mm_vertices(), plot_size = 0.01,
      subplot_size = 10, voucher_imgs = NULL, verbose = TRUE,
      dir = file.path(out_dir, "sub/"), language = "pt"
    )
  )
  expect_true(any(grepl("1 row\\(s\\) without valid T1/X/Y were not mapped", msgs)))
  expect_identical(dirname(normalizePath(out, winslash = "/", mustWork = FALSE)),
                   normalizePath(file.path(out_dir, "sub"), winslash = "/"))
})

test_that("dir = NULL writes the map to tempdir()", {
  saved <- .mm_mock_save()
  testthat::local_mocked_bindings(.harmonize_plot_input = .mm_fp_loaded)
  out <- suppressMessages(forplot_map(
    .mm_input(), vertex_coords = .mm_vertices(), plot_size = 0.01,
    subplot_size = 10, voucher_imgs = NULL, dir = NULL, filename = "tmpmap"
  ))
  withr::defer(unlink(out))
  expect_identical(normalizePath(dirname(out), winslash = "/"),
                   normalizePath(tempdir(), winslash = "/"))
})

test_that("misplaced corners trigger a vertex geometry warning", {
  saved <- .mm_mock_save()
  testthat::local_mocked_bindings(.harmonize_plot_input = .mm_fp_loaded)
  far <- data.frame(Latitude = c(-3, -2.99, -3, -2.99),
                    Longitude = c(-60, -60, -59.99, -59.99))
  expect_warning(
    suppressMessages(forplot_map(
      .mm_input(), vertex_coords = far, plot_size = 0.01, subplot_size = 10,
      voucher_imgs = NULL, dir = withr::local_tempdir()
    )),
    "Vertex distances .* do not match"
  )
})


# Several MONITORA stations ----------------------------------------------------

test_that("several MONITORA stations produce one map each", {
  saved <- .mm_mock_save()
  stations_seen <- character(0)
  testthat::local_mocked_bindings(
    .harmonize_plot_input = function(..., station_name = NULL) {
      stations_seen <<- c(stations_seen, station_name)
      .mm_monitora_loaded()
    },
    .compute_monitora_geometry = function(fp_df, keep_only_cell = FALSE) {
      fp_df$draw_x <- .numify(fp_df$X)
      fp_df$draw_y <- .numify(fp_df$Y)
      fp_df
    }
  )
  centers <- list(`EA-2` = c(lat = -1, lon = -52), `EA-1` = c(lat = 1, lon = -52))
  out <- suppressMessages(forplot_map(
    .mm_input(), input_type = "monitora", vertex_coords = centers,
    station_name = c("EA-1", "EA-2"), voucher_imgs = NULL,
    dir = withr::local_tempdir(), filename = "mon"
  ))
  expect_length(out, 2L)
  expect_true(all(grepl("_mon_station_EA-[12]\\.html$", out)))
  expect_identical(stations_seen, c("EA-1", "EA-2"))
})

test_that(".vertex_for_station selects one center per station", {
  st <- c("EA-1", "EA-2")
  expect_identical(.vertex_for_station(list(c(1, 2), c(3, 4)), st, 2), c(3, 4))
  expect_identical(.vertex_for_station(list(`EA-2` = 9), st, 2), 9)
  expect_error(.vertex_for_station(list(1, 2, 3), st, 1),
               "one element per station")

  df <- data.frame(estacao = c("EA-2", "EA-1"), lat = c(1, 2), lon = c(3, 4))
  expect_identical(.vertex_for_station(df, st, 1)$lat, 2)
  expect_identical(.vertex_for_station(df, st, 2)$lat, 1)
  expect_error(.vertex_for_station(df, c("EA-7"), 1), "Station 'EA-7' was not found")

  rows <- data.frame(lat = c(1, 2), lon = c(3, 4))
  expect_identical(.vertex_for_station(rows, st, 2)$lat, 2)
  expect_error(.vertex_for_station(rows[1, ], st, 1),
               "must give one center per station")

  csv <- withr::local_tempfile(fileext = ".csv")
  utils::write.csv(df, csv, row.names = FALSE)
  expect_equal(.vertex_for_station(csv, st, 1)$lat, 2)

  xlsx <- withr::local_tempfile(fileext = ".xlsx")
  openxlsx::write.xlsx(df, xlsx)
  expect_equal(.vertex_for_station(xlsx, st, 2)$lat, 1)
})


# Vertex parsing ---------------------------------------------------------------

test_that(".parse_vertex_coords reads csv files and validates input", {
  csv <- withr::local_tempfile(fileext = ".csv")
  utils::write.csv(.mm_vertices(), csv, row.names = FALSE)
  expect_equal(.parse_vertex_coords(csv)$latitude, .mm_vertices()$Latitude)

  expect_error(.parse_vertex_coords("corners.txt"), "must be .xlsx, .xls or .csv")
  expect_identical(.parse_vertex_coords(c(-3, -60)),
                   data.frame(latitude = -3, longitude = -60))
  expect_error(.parse_vertex_coords(data.frame(a = 1, b = 2)),
               "Could not find latitude/longitude columns")
  expect_error(.parse_vertex_coords(data.frame(lat = c(-3, 95), lon = c(-60, -60))),
               "Invalid coordinates in `vertex_coords` \\(row\\(s\\) 2\\)")
})


# Images and popups ------------------------------------------------------------

test_that("voucher image sources are embedded or copied", {
  src <- withr::local_tempdir()
  dir.create(file.path(src, "V 1"))
  dir.create(file.path(src, "empty"))
  dir.create(file.path(src, "!!!", "deep"), recursive = TRUE)
  png <- file.path(src, "V 1", "a b.png")
  writeBin(as.raw(1:10), png)
  writeBin(as.raw(1:5), file.path(src, "!!!", "x.jpg"))

  idx <- .voucher_image_index(src)
  expect_identical(names(idx), "v1")

  expect_identical(.voucher_image_sources(list(), TRUE, src, "a"), list())

  emb <- .voucher_image_sources(idx, embed = TRUE, out_dir = src, asset_dirname = "a")
  expect_match(emb$v1, "^data:image/png;base64,")

  out <- withr::local_tempdir()
  cp <- .voucher_image_sources(idx, embed = FALSE, out_dir = out,
                               asset_dirname = "map_files")
  expect_identical(cp$v1, "map_files/v1/a%20b.png")
  expect_true(file.exists(file.path(out, "map_files", "v1", "a b.png")))
})

test_that("small HTML helpers handle edge cases", {
  expect_identical(.carousel_html(character(0)), "")
  one <- .carousel_html("a.png", embed = TRUE)
  expect_match(one, "<div class='fp-slide'><img loading='lazy' src='a.png'")
  expect_false(grepl("fp-nav", one))
  two <- .carousel_html(c("a.png", "b.png"))
  expect_match(two, "1 / 2")
  expect_match(two, "href='a.png'")

  expect_identical(.dbh_radius(c(NA, -1)), c(4, 4))
  expect_identical(.dbh_radius(c(9, 9)), c(6.5, 6.5))
  r <- .dbh_radius(c(1, 100, NA))
  expect_identical(r, c(3, 10, 4))

  m <- .get_latlon(c(0, 10), c(0, 10), c(-60, -3), c(-60, -2.99991),
                   c(-59.99991, -3), c(-59.99991, -2.99991))
  expect_identical(dim(m), c(2L, 2L))
  cm <- .get_latlon_from_center(c(0, 1), c(0, 1), c(-60, -3))
  expect_identical(dim(cm), c(2L, 2L))

  expect_identical(.map_ui_labels("xx"), .map_ui_labels("en"))
  expect_false(identical(.map_ui_labels("fr"), .map_ui_labels("en")))
})
