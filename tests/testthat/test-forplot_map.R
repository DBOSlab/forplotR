# Tests for forplot_map ------------------------------------------------------
#
# Design goals:
# - CRAN-safe: no internet access, no browser, no external herbarium downloads.
# - Deterministic: all filesystem writes go to temporary directories.
# - Fast: integration tests use a 0.01-ha plot (one 10 x 10 m subplot).
# - Unit tests cover parsing, geometry, validation and map orchestration.
#
# These tests assume testthat edition 3 and testthat >= 3.2.0 because
# local_mocked_bindings() is used to isolate external side effects.

# -------------------------------------------------------------------------
# Local test fixtures
# -------------------------------------------------------------------------

.map_test_fp_loaded <- function() {
  list(
    fp_sheet = data.frame(
      T1 = c(1, 1, 1),
      X = c(1, 4, 8),
      Y = c(1, 5, 8),
      D = c(100, 180, 250),
      Family = c("Fabaceae", "Arecaceae", "Myrtaceae"),
      Collected = c("C001", NA_character_, ""),
      `New Tag No` = c("1", "2", "3"),
      `Original determination` = c(
        "Inga alba",
        "Euterpe precatoria",
        "Eugenia sp."
      ),
      Voucher = c("V1", "", NA_character_),
      check.names = FALSE,
      stringsAsFactors = FALSE
    ),
    team = "Test Team",
    plot_name = "Test Plot",
    plot_code = "TEST01",
    census_no_fp = "1"
  )
}

.map_test_monitora_loaded <- function() {
  list(
    fp_sheet = data.frame(
      T1 = c("N1", "N1"),
      T2 = c("1", "1"),
      X = c("1,5", "5"),
      Y = c("2,5", "7"),
      D = c("100", "200"),
      Family = c("Fabaceae", "Arecaceae"),
      Collected = c("C001", NA_character_),
      `New Tag No` = c("10", "11"),
      `Original determination` = c("Inga alba", "Euterpe precatoria"),
      Voucher = c("V10", ""),
      check.names = FALSE,
      stringsAsFactors = FALSE
    ),
    team = "MONITORA Team",
    plot_name = "MONITORA Test",
    plot_code = "MON01",
    census_no_fp = "1"
  )
}

.map_test_vertices <- function() {
  data.frame(
    Latitude = c(-3.0000, -3.0000, -3.0010, -3.0010),
    Longitude = c(-60.0000, -59.9990, -60.0000, -59.9990)
  )
}

.map_test_vertices_with_units <- function() {
  data.frame(
    `Latitude (WGS84)` = c("-3,0000", "-3,0000", "-3,0010", "-3,0010"),
    `Longitude (WGS84)` = c("-60,0000", "-59,9990", "-60,0000", "-59,9990"),
    check.names = FALSE
  )
}

.map_test_input_file <- function() {
  f <- tempfile(fileext = ".xlsx")
  ok <- file.create(f)
  stopifnot(ok)
  f
}

.map_test_missing_voucher_dir <- function() {
  paste0("forplotR_nonexistent_voucher_dir_", Sys.getpid(), "_", sample.int(1e6, 1))
}

.map_contains_text <- function(x, pattern) {
  if (is.character(x)) {
    return(any(grepl(pattern, x, fixed = TRUE), na.rm = TRUE))
  }

  if (is.list(x)) {
    return(any(vapply(
      x,
      .map_contains_text,
      FUN.VALUE = logical(1),
      pattern = pattern
    )))
  }

  FALSE
}

.map_mock_standard_input <- function(...) {
  .map_test_fp_loaded()
}

.map_mock_monitora_input <- function(...) {
  .map_test_monitora_loaded()
}

.map_mock_monitora_geometry <- function(fp_df, keep_only_cell = FALSE) {
  dplyr::mutate(
    fp_df,
    draw_x = as.numeric(X),
    draw_y = as.numeric(Y)
  )
}


# -------------------------------------------------------------------------
# Argument validation
# -------------------------------------------------------------------------

test_that("forplot_map rejects invalid input_type values", {
  expect_error(
    forplot_map(
      fp_file_path = "irrelevant.xlsx",
      input_type = "not_a_type"
    ),
    "should be one of"
  )
})

test_that("forplot_map normalizes input_type case and whitespace", {
  input_file <- .map_test_input_file()
  saved <- new.env(parent = emptyenv())

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .package = "forplotR"
  )

  testthat::local_mocked_bindings(
    saveWidget = function(widget, file, selfcontained = TRUE, ...) {
      saved$file <- file
      saved$selfcontained <- selfcontained
      writeLines("<html></html>", file)
      invisible(NULL)
    },
    .package = "htmlwidgets"
  )

  old <- setwd(tempdir())
  on.exit(setwd(old), add = TRUE)

  out <- suppressMessages(
    forplot_map(
      fp_file_path = input_file,
      input_type = "  FIELD_SHEET  ",
      vertex_coords = .map_test_vertices(),
      plot_size = 0.01,
      subplot_size = 10,
      voucher_imgs = .map_test_missing_voucher_dir(),
      filename = "normalized_type"
    )
  )

  expect_true(file.exists(saved$file))
  expect_true(endsWith(saved$file, "_normalized_type.html"))
  expect_true(isTRUE(saved$selfcontained))

  # Contract documented by forplot_map(): invisibly return output path.
  expect_identical(
    normalizePath(out, winslash = "/", mustWork = FALSE),
    normalizePath(saved$file, winslash = "/", mustWork = FALSE)
  )
})

test_that("forplot_map validates plot_size before reading the input file", {
  expect_error(
    forplot_map(
      fp_file_path = "does-not-exist.xlsx",
      plot_size = 0
    ),
    "`plot_size` must be a single positive numeric value"
  )

  expect_error(
    forplot_map(
      fp_file_path = "does-not-exist.xlsx",
      plot_size = NA_real_
    ),
    "`plot_size` must be a single positive numeric value"
  )
})

test_that("forplot_map validates subplot_size before reading the input file", {
  expect_error(
    forplot_map(
      fp_file_path = "does-not-exist.xlsx",
      subplot_size = 0
    ),
    "`subplot_size` must be a single positive numeric value"
  )

  expect_error(
    forplot_map(
      fp_file_path = "does-not-exist.xlsx",
      subplot_size = Inf
    ),
    "`subplot_size` must be a single positive numeric value"
  )
})

test_that("forplot_map requires an existing scalar fp_file_path", {
  expect_error(
    forplot_map(fp_file_path = NULL),
    "does not exist"
  )

  expect_error(
    forplot_map(fp_file_path = c("a.xlsx", "b.xlsx")),
    "does not exist"
  )

  expect_error(
    forplot_map(fp_file_path = 123),
    "does not exist"
  )

  expect_error(
    forplot_map(fp_file_path = tempfile(fileext = ".xlsx")),
    "does not exist"
  )
})



test_that("forplot_map requires vertex_coords explicitly", {
  input_file <- .map_test_input_file()

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .package = "forplotR"
  )

  expect_error(
    forplot_map(
      fp_file_path = input_file,
      input_type = "field_sheet",
      vertex_coords = NULL,
      plot_size = 0.01,
      subplot_size = 10
    ),
    "`vertex_coords` must be provided"
  )
})

# -------------------------------------------------------------------------
# Validation helpers
# -------------------------------------------------------------------------

test_that(".validate_plot_size accepts positive finite scalar numeric values", {
  expect_invisible(forplotR:::.validate_plot_size(1))
  expect_invisible(forplotR:::.validate_plot_size(0.01))
  expect_invisible(forplotR:::.validate_plot_size(2.5))
})

test_that(".validate_plot_size rejects invalid values", {
  bad <- list(
    0,
    -1,
    NA_real_,
    NaN,
    Inf,
    -Inf,
    "1",
    c(1, 2),
    NULL
  )

  for (x in bad) {
    expect_error(
      forplotR:::.validate_plot_size(x),
      "`plot_size` must be a single positive numeric value",
      info = paste("value:", paste(x, collapse = ", "))
    )
  }
})

test_that(".validate_subplot_size accepts positive finite scalar numeric values", {
  expect_invisible(forplotR:::.validate_subplot_size(10))
  expect_invisible(forplotR:::.validate_subplot_size(20))
})

test_that(".validate_subplot_size rejects invalid values", {
  bad <- list(
    0,
    -10,
    NA_real_,
    NaN,
    Inf,
    -Inf,
    "10",
    c(10, 20),
    NULL
  )

  for (x in bad) {
    expect_error(
      forplotR:::.validate_subplot_size(x),
      "`subplot_size`",
      info = paste("value:", paste(x, collapse = ", "))
    )
  }
})


# -------------------------------------------------------------------------
# Numeric parsing helpers
# -------------------------------------------------------------------------

test_that(".numify preserves numeric values", {
  x <- c(1, 2.5, NA_real_, -3)
  expect_equal(forplotR:::.numify(x), x)
})

test_that(".numify parses decimal commas and thousands separators", {
  x <- c(
    "1,5",
    "  2,75 ",
    "1.234,56",
    "2 345,5",
    "-10,25",
    ""
  )

  expect_equal(
    forplotR:::.numify(x),
    c(1.5, 2.75, 1234.56, 2345.5, -10.25, NA_real_)
  )
})

test_that(".numify returns NA for non-numeric text", {
  expect_equal(
    forplotR:::.numify(c("abc", "-", "NA", NA_character_)),
    rep(NA_real_, 4)
  )
})

test_that(".coerce_xy preserves numeric values", {
  x <- c(1, 2.5, NA_real_, -3)
  expect_equal(forplotR:::.coerce_xy(x), x)
})

test_that(".coerce_xy parses common ForestPlots coordinate strings", {
  x <- c(
    "1,5",
    "2.5",
    "3,5/9,0",
    "4;8",
    "5|10",
    "X = 6,25 m",
    "-7.5 m",
    ""
  )

  expect_equal(
    forplotR:::.coerce_xy(x),
    c(1.5, 2.5, 3.5, 4, 5, 6.25, -7.5, NA_real_)
  )
})

test_that(".coerce_xy handles missing and non-numeric values", {
  expect_equal(
    forplotR:::.coerce_xy(c(NA_character_, "abc", "none")),
    rep(NA_real_, 3)
  )
})


# -------------------------------------------------------------------------
# Generic plot-coordinate geometry
# -------------------------------------------------------------------------

test_that(".compute_global_coordinates computes serpentine coordinates", {

  df <- data.frame(
    T1 = c(1, 10, 11, 20),
    X = c(1, 2, 3, 4),
    Y = c(1, 2, 3, 4)
  )

  out <- forplotR:::.compute_global_coordinates(
    fp_clean = df,
    subplot_size = 10,
    plot_width_m = 100,
    plot_length_m = 100
  )

  expect_equal(
    out$col,
    c(0, 0, 1, 1)
  )

  expect_equal(
    out$row,
    c(0, 9, 0, 9)
  )

  expect_equal(
    out$global_x,
    c(1, 2, 13, 14)
  )

  expect_equal(
    out$global_y,
    c(1, 92, 93, 4)
  )

  expect_equal(
    out$draw_x,
    out$global_x
  )

  expect_equal(
    out$draw_y,
    out$global_y
  )
})

test_that(".compute_global_coordinates supports a one-subplot plot", {

  df <- data.frame(
    T1 = 1,
    X = 2,
    Y = 3
  )

  out <- forplotR:::.compute_global_coordinates(
    fp_clean = df,
    subplot_size = 10,
    plot_width_m = 10,
    plot_length_m = 10
  )

  expect_equal(
    out$global_x,
    2
  )

  expect_equal(
    out$global_y,
    3
  )

  expect_equal(
    out$draw_x,
    2
  )

  expect_equal(
    out$draw_y,
    3
  )
})
test_that("shared geometry rejects incompatible plot and subplot combinations", {

  expect_error(
    forplotR:::.resolve_plot_geometry(
      plot_size = 1,
      subplot_size = 30
    ),
    "do not form an integer number of square subplots"
  )
})

test_that("shared geometry supports large rectangular plots", {

  geometry <- forplotR:::.resolve_plot_geometry(
    plot_size = 10,
    subplot_size = 10,
    plot_width_m = 200,
    plot_length_m = 500
  )

  expect_equal(
    geometry$plot_width_m,
    200
  )

  expect_equal(
    geometry$plot_length_m,
    500
  )

  expect_equal(
    geometry$n_cols,
    20L
  )

  expect_equal(
    geometry$n_rows,
    50L
  )

  expect_equal(
    geometry$n_subplots,
    1000L
  )

  df <- data.frame(
    T1 = c(
      1,
      50,
      51,
      1000
    ),
    X = c(
      1,
      2,
      3,
      4
    ),
    Y = c(
      1,
      2,
      3,
      4
    )
  )

  out <- forplotR:::.compute_global_coordinates(
    fp_clean = df,
    subplot_size = 10,
    plot_width_m = geometry$plot_width_m,
    plot_length_m = geometry$plot_length_m
  )

  expect_equal(
    nrow(out),
    4L
  )

  expect_true(
    all(
      out$global_x >= 0 &
        out$global_x <= 200
    )
  )

  expect_true(
    all(
      out$global_y >= 0 &
        out$global_y <= 500
    )
  )
})

# -------------------------------------------------------------------------
# Geographic conversion helpers
# -------------------------------------------------------------------------

test_that(".get_latlon_from_center returns the center for zero displacement", {
  center <- c(-60, -3)

  out <- forplotR:::.get_latlon_from_center(
    x = 0,
    y = 0,
    center_lonlat = center
  )

  expect_named(out, c("lat", "lon"))
  expect_equal(unname(out["lat"]), center[2], tolerance = 1e-7)
  expect_equal(unname(out["lon"]), center[1], tolerance = 1e-7)
})

test_that(".get_latlon_from_center moves east/west and north/south correctly", {
  center <- c(-60, -3)

  east <- forplotR:::.get_latlon_from_center(100, 0, center)
  west <- forplotR:::.get_latlon_from_center(-100, 0, center)
  north <- forplotR:::.get_latlon_from_center(0, 100, center)
  south <- forplotR:::.get_latlon_from_center(0, -100, center)

  expect_gt(east["lon"], center[1])
  expect_lt(west["lon"], center[1])
  expect_gt(north["lat"], center[2])
  expect_lt(south["lat"], center[2])
})

test_that(".get_latlon returns the first corner at x = y = 0", {
  vertices <- .map_test_vertices()

  p1 <- c(vertices$Longitude[1], vertices$Latitude[1])
  p2 <- c(vertices$Longitude[2], vertices$Latitude[2])
  p3 <- c(vertices$Longitude[3], vertices$Latitude[3])
  p4 <- c(vertices$Longitude[4], vertices$Latitude[4])

  out <- forplotR:::.get_latlon(0, 0, p1, p2, p3, p4)

  expect_named(out, c("lat", "lon"))
  expect_equal(unname(out["lat"]), p1[2], tolerance = 1e-7)
  expect_equal(unname(out["lon"]), p1[1], tolerance = 1e-7)
})

test_that(".get_latlon produces finite geographic coordinates", {
  vertices <- .map_test_vertices()

  p1 <- c(vertices$Longitude[1], vertices$Latitude[1])
  p2 <- c(vertices$Longitude[2], vertices$Latitude[2])
  p3 <- c(vertices$Longitude[3], vertices$Latitude[3])
  p4 <- c(vertices$Longitude[4], vertices$Latitude[4])

  out <- forplotR:::.get_latlon(5, 5, p1, p2, p3, p4)

  expect_true(all(is.finite(out)))
  expect_true(out["lat"] >= -90 && out["lat"] <= 90)
  expect_true(out["lon"] >= -180 && out["lon"] <= 180)
})


# -------------------------------------------------------------------------
# Subplot geometry helpers
# -------------------------------------------------------------------------

test_that(".get_subplot_polygon returns a closed square", {
  out <- forplotR:::.get_subplot_polygon(
    col = 2,
    row = 3,
    size = 10
  )

  expected <- matrix(
    c(
      20, 30,
      30, 30,
      30, 40,
      20, 40,
      20, 30
    ),
    ncol = 2,
    byrow = TRUE
  )

  expect_equal(out, expected)
  expect_equal(out[1, ], out[nrow(out), ])
})

test_that(".mk_cells creates ten cells per MONITORA arm", {
  for (arm in c("N", "S", "L", "O")) {
    cells <- forplotR:::.mk_cells(arm)

    expect_s3_class(cells, "data.frame")
    expect_equal(nrow(cells), 10)
    expect_true(all(c("c", "r", "xmin", "xmax", "ymin", "ymax") %in% names(cells)))
    expect_true(all((cells$xmax - cells$xmin) == 10))
    expect_true(all((cells$ymax - cells$ymin) == 10))
  }
})

test_that(".mk_cells places MONITORA arms in their expected quadrants", {
  north <- forplotR:::.mk_cells("N")
  south <- forplotR:::.mk_cells("S")
  east <- forplotR:::.mk_cells("L")
  west <- forplotR:::.mk_cells("O")

  expect_gte(min(north$ymin), 50)
  expect_lte(max(south$ymax), -50)
  expect_gte(min(east$xmin), 50)
  expect_lte(max(west$xmax), -50)
})


# -------------------------------------------------------------------------
# Input cleaning and vertex validation in the main function
# -------------------------------------------------------------------------

test_that("field-sheet input errors when no T1/X/Y values are usable", {
  input_file <- .map_test_input_file()

  bad_loaded <- .map_test_fp_loaded()
  bad_loaded$fp_sheet$T1 <- "bad"
  bad_loaded$fp_sheet$X <- "bad"
  bad_loaded$fp_sheet$Y <- "bad"

  testthat::local_mocked_bindings(
    .harmonize_plot_input = function(...) bad_loaded,
    .package = "forplotR"
  )

  expect_error(
    forplot_map(
      fp_file_path = input_file,
      input_type = "field_sheet",
      vertex_coords = .map_test_vertices(),
      plot_size = 0.01,
      subplot_size = 10
    ),
    "No valid points to plot: T1/X/Y could not be parsed"
  )
})

test_that("MONITORA input errors when no X/Y values are usable", {
  input_file <- .map_test_input_file()

  bad_loaded <- .map_test_monitora_loaded()
  bad_loaded$fp_sheet$X <- "bad"
  bad_loaded$fp_sheet$Y <- "bad"

  testthat::local_mocked_bindings(
    .harmonize_plot_input = function(...) bad_loaded,
    .package = "forplotR"
  )

  expect_error(
    forplot_map(
      fp_file_path = input_file,
      input_type = "monitora",
      vertex_coords = c(lat = -3, lon = -60),
      plot_size = 0.01,
      subplot_size = 10
    ),
    "No valid points to plot after cleaning MONITORA X/Y values"
  )
})

test_that("MONITORA geometry helper must provide draw_x and draw_y", {
  input_file <- .map_test_input_file()

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_monitora_input,
    .compute_monitora_geometry = function(fp_df, keep_only_cell = FALSE) fp_df,
    .package = "forplotR"
  )

  expect_error(
    forplot_map(
      fp_file_path = input_file,
      input_type = "monitora",
      vertex_coords = c(lat = -3, lon = -60),
      plot_size = 0.01,
      subplot_size = 10
    ),
    "must add 'draw_x' and 'draw_y'"
  )
})

test_that("invalid atomic vertex_coords is rejected", {
  input_file <- .map_test_input_file()

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .package = "forplotR"
  )

  expect_error(
    forplot_map(
      fp_file_path = input_file,
      input_type = "field_sheet",
      vertex_coords = c(-3, -60, 0),
      plot_size = 0.01,
      subplot_size = 10
    ),
    "vertex_coords must be a 1-row data.frame or a length-2 numeric vector"
  )
})

test_that("standard plot inputs require four corners", {
  input_file <- .map_test_input_file()

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .package = "forplotR"
  )

  expect_error(
    forplot_map(
      fp_file_path = input_file,
      input_type = "field_sheet",
      vertex_coords = c(-3, -60),
      plot_size = 0.01,
      subplot_size = 10
    ),
    "must provide FOUR plot corners"
  )
})

test_that("MONITORA requires exactly one central coordinate", {
  input_file <- .map_test_input_file()

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_monitora_input,
    .compute_monitora_geometry = .map_mock_monitora_geometry,
    .package = "forplotR"
  )

  expect_error(
    forplot_map(
      fp_file_path = input_file,
      input_type = "monitora",
      vertex_coords = .map_test_vertices(),
      plot_size = 0.01,
      subplot_size = 10
    ),
    "provide ONE central coordinate"
  )
})

test_that("named MONITORA coordinate vectors may be supplied lon before lat", {
  input_file <- .map_test_input_file()
  saved <- new.env(parent = emptyenv())

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_monitora_input,
    .compute_monitora_geometry = .map_mock_monitora_geometry,
    .package = "forplotR"
  )

  testthat::local_mocked_bindings(
    saveWidget = function(widget, file, selfcontained = TRUE, ...) {
      saved$file <- file
      writeLines("<html></html>", file)
      invisible(NULL)
    },
    .package = "htmlwidgets"
  )

  old <- setwd(tempdir())
  on.exit(setwd(old), add = TRUE)

  out <- suppressMessages(
    forplot_map(
      fp_file_path = input_file,
      input_type = "monitora",
      vertex_coords = c(lon = -60, lat = -3),
      plot_size = 0.01,
      subplot_size = 10,
      voucher_imgs = .map_test_missing_voucher_dir(),
      filename = "monitora_named_center"
    )
  )

  expect_true(file.exists(saved$file))
  expect_true(endsWith(saved$file, "_monitora_named_center.html"))

  expect_identical(
    normalizePath(out, winslash = "/", mustWork = FALSE),
    normalizePath(saved$file, winslash = "/", mustWork = FALSE)
  )
})

test_that("vertex data frame names with units and decimal commas are accepted", {
  input_file <- .map_test_input_file()
  saved <- new.env(parent = emptyenv())

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .package = "forplotR"
  )

  testthat::local_mocked_bindings(
    saveWidget = function(widget, file, selfcontained = TRUE, ...) {
      saved$file <- file
      writeLines("<html></html>", file)
      invisible(NULL)
    },
    .package = "htmlwidgets"
  )

  old <- setwd(tempdir())
  on.exit(setwd(old), add = TRUE)

  suppressMessages(
    out <- forplot_map(
      fp_file_path = input_file,
      input_type = "field_sheet",
      vertex_coords = .map_test_vertices_with_units(),
      plot_size = 0.01,
      subplot_size = 10,
      voucher_imgs = .map_test_missing_voucher_dir(),
      filename = "vertex_units"
    )
  )

  expect_true(file.exists(saved$file))
  expect_true(endsWith(saved$file, "_vertex_units.html"))

  expect_identical(
    normalizePath(out, winslash = "/", mustWork = FALSE),
    normalizePath(saved$file, winslash = "/", mustWork = FALSE)
  )
})

test_that("xlsx vertex coordinate input is delegated to readxl", {
  input_file <- .map_test_input_file()
  called <- new.env(parent = emptyenv())
  called$value <- FALSE

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .package = "forplotR"
  )

  testthat::local_mocked_bindings(
    read_excel = function(path, ...) {
      called$value <- TRUE
      data.frame(
        Latitude = c(-3, -3, -3.001),
        Longitude = c(-60, -59.999, -60)
      )
    },
    .package = "readxl"
  )

  expect_error(
    forplot_map(
      fp_file_path = input_file,
      input_type = "field_sheet",
      vertex_coords = "vertices.xlsx",
      plot_size = 0.01,
      subplot_size = 10
    ),
    "must provide FOUR plot corners"
  )

  expect_true(called$value)
})


# -------------------------------------------------------------------------
# Herbarium lookup behavior
# -------------------------------------------------------------------------

test_that("herbarium lookup is not called when herbaria_lookup is FALSE", {
  input_file <- .map_test_input_file()
  saved <- new.env(parent = emptyenv())

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .herbaria_lookup_links = function(...) {
      stop("Herbarium lookup should not have been called.")
    },
    .package = "forplotR"
  )

  testthat::local_mocked_bindings(
    saveWidget = function(widget, file, selfcontained = TRUE, ...) {
      saved$file <- file
      writeLines("<html></html>", file)
      invisible(NULL)
    },
    .package = "htmlwidgets"
  )

  old <- setwd(tempdir())
  on.exit(setwd(old), add = TRUE)

  expect_no_error(
    suppressMessages(
      forplot_map(
        fp_file_path = input_file,
        input_type = "field_sheet",
        vertex_coords = .map_test_vertices(),
        plot_size = 0.01,
        subplot_size = 10,
        voucher_imgs = .map_test_missing_voucher_dir(),
        filename = "no_herbarium_lookup",
        herbaria_lookup = FALSE
      )
    )
  )

  expect_true(file.exists(saved$file))
})

test_that("empty herbaria with lookup enabled gives a warning and continues", {
  input_file <- .map_test_input_file()
  saved <- new.env(parent = emptyenv())
  lookup_called <- new.env(parent = emptyenv())
  lookup_called$value <- FALSE

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .herbaria_lookup_links = function(...) {
      lookup_called$value <- TRUE
      stop("should not be reached")
    },
    .package = "forplotR"
  )

  testthat::local_mocked_bindings(
    saveWidget = function(widget, file, selfcontained = TRUE, ...) {
      saved$file <- file
      writeLines("<html></html>", file)
      invisible(NULL)
    },
    .package = "htmlwidgets"
  )

  old <- setwd(tempdir())
  on.exit(setwd(old), add = TRUE)

  expect_warning(
    suppressMessages(
      forplot_map(
        fp_file_path = input_file,
        input_type = "field_sheet",
        vertex_coords = .map_test_vertices(),
        plot_size = 0.01,
        subplot_size = 10,
        voucher_imgs = .map_test_missing_voucher_dir(),
        filename = "empty_herbaria",
        herbaria_lookup = TRUE,
        herbaria = NULL
      )
    ),
    "`herbaria_lookup = TRUE` but `herbaria` is NULL or empty"
  )

  expect_false(lookup_called$value)
  expect_true(file.exists(saved$file))
})

test_that("herbarium lookup failures warn and map generation continues", {
  input_file <- .map_test_input_file()
  saved <- new.env(parent = emptyenv())

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .herbaria_lookup_links = function(...) {
      stop("synthetic lookup failure")
    },
    .package = "forplotR"
  )

  testthat::local_mocked_bindings(
    saveWidget = function(widget, file, selfcontained = TRUE, ...) {
      saved$file <- file
      writeLines("<html></html>", file)
      invisible(NULL)
    },
    .package = "htmlwidgets"
  )

  old <- setwd(tempdir())
  on.exit(setwd(old), add = TRUE)

  expect_warning(
    suppressMessages(
      forplot_map(
        fp_file_path = input_file,
        input_type = "field_sheet",
        vertex_coords = .map_test_vertices(),
        plot_size = 0.01,
        subplot_size = 10,
        voucher_imgs = .map_test_missing_voucher_dir(),
        filename = "lookup_failure",
        herbaria_lookup = TRUE,
        herbaria = "RB"
      )
    ),
    "Herbarium lookup failed: synthetic lookup failure"
  )

  expect_true(file.exists(saved$file))
})

test_that("invalid herbarium lookup result warns and is ignored", {
  input_file <- .map_test_input_file()
  saved <- new.env(parent = emptyenv())

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .herbaria_lookup_links = function(...) "wrong-length-result",
    .package = "forplotR"
  )

  testthat::local_mocked_bindings(
    saveWidget = function(widget, file, selfcontained = TRUE, ...) {
      saved$file <- file
      writeLines("<html></html>", file)
      invisible(NULL)
    },
    .package = "htmlwidgets"
  )

  old <- setwd(tempdir())
  on.exit(setwd(old), add = TRUE)

  expect_warning(
    suppressMessages(
      forplot_map(
        fp_file_path = input_file,
        input_type = "field_sheet",
        vertex_coords = .map_test_vertices(),
        plot_size = 0.01,
        subplot_size = 10,
        voucher_imgs = .map_test_missing_voucher_dir(),
        filename = "invalid_lookup_result",
        herbaria_lookup = TRUE,
        herbaria = "RB"
      )
    ),
    "Invalid result returned by herbarium lookup"
  )

  expect_true(file.exists(saved$file))
})

test_that("valid herbarium links are incorporated into popup content", {
  input_file <- .map_test_input_file()
  saved <- new.env(parent = emptyenv())

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .herbaria_lookup_links = function(fp_df, ...) {
      c(
        "<br/><a href='https://example.invalid/record/1'>RB</a>",
        NA_character_,
        NA_character_
      )
    },
    .package = "forplotR"
  )

  testthat::local_mocked_bindings(
    saveWidget = function(widget, file, selfcontained = TRUE, ...) {
      saved$widget <- widget
      saved$file <- file
      writeLines("<html></html>", file)
      invisible(NULL)
    },
    .package = "htmlwidgets"
  )

  old <- setwd(tempdir())
  on.exit(setwd(old), add = TRUE)

  suppressWarnings(
    suppressMessages(
      forplot_map(
        fp_file_path = input_file,
        input_type = "field_sheet",
        vertex_coords = .map_test_vertices(),
        plot_size = 0.01,
        subplot_size = 10,
        voucher_imgs = .map_test_missing_voucher_dir(),
        filename = "valid_lookup_link",
        herbaria_lookup = TRUE,
        herbaria = "RB"
      )
    )
  )

  expect_true(.map_contains_text(saved$widget, "example.invalid/record/1"))
  expect_true(.map_contains_text(saved$widget, "RB"))
})



test_that("herbarium lookup receives all lookup-control arguments", {
  input_file <- .map_test_input_file()
  saved <- new.env(parent = emptyenv())
  received <- new.env(parent = emptyenv())

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .herbaria_lookup_links = function(fp_df,
                                      herbaria,
                                      force_refresh,
                                      keep_downloads,
                                      collector_fallback,
                                      collector_codes,
                                      verbose) {
      received$herbaria <- herbaria
      received$force_refresh <- force_refresh
      received$keep_downloads <- keep_downloads
      received$collector_fallback <- collector_fallback
      received$collector_codes <- collector_codes
      received$verbose <- verbose
      rep(NA_character_, nrow(fp_df))
    },
    .package = "forplotR"
  )

  testthat::local_mocked_bindings(
    saveWidget = function(widget, file, selfcontained = TRUE, ...) {
      saved$file <- file
      writeLines("<html></html>", file)
      invisible(NULL)
    },
    .package = "htmlwidgets"
  )

  old <- setwd(tempdir())
  on.exit(setwd(old), add = TRUE)

  suppressMessages(
    forplot_map(
      fp_file_path = input_file,
      input_type = "field_sheet",
      vertex_coords = .map_test_vertices(),
      plot_size = 0.01,
      subplot_size = 10,
      voucher_imgs = .map_test_missing_voucher_dir(),
      filename = "lookup_args",
      collector = "G. C. Ottino",
      collector_map = c(GCO = "G. C. Ottino"),
      herbaria_lookup = TRUE,
      herbaria = c("RB", "INPA"),
      herbaria_force_refresh = TRUE,
      keep_herbaria_downloads = TRUE,
      verbose = FALSE
    )
  )

  expect_identical(received$herbaria, c("RB", "INPA"))
  expect_true(received$force_refresh)
  expect_true(received$keep_downloads)
  expect_identical(received$collector_fallback, "G. C. Ottino")
  expect_identical(received$collector_codes, c(GCO = "G. C. Ottino"))
  expect_false(received$verbose)
  expect_true(file.exists(saved$file))
})

# -------------------------------------------------------------------------
# Voucher image carousel
# -------------------------------------------------------------------------

test_that("voucher image directories add a photo badge and slideshow to popup", {
  input_file <- .map_test_input_file()
  saved <- new.env(parent = emptyenv())

  img_root <- tempfile("voucher_imgs_")
  voucher_dir <- file.path(img_root, "V1")
  dir.create(voucher_dir, recursive = TRUE)
  image_file <- file.path(voucher_dir, "specimen.JPG")
  expect_true(file.create(image_file))

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .package = "forplotR"
  )

  # forplot_map() calls here::here(voucher_imgs). For this isolated filesystem
  # test we want the absolute temporary path unchanged.
  testthat::local_mocked_bindings(
    here = function(...) {
      args <- list(...)
      if (!length(args)) return(getwd())
      as.character(args[[1]])
    },
    .package = "here"
  )

  testthat::local_mocked_bindings(
    saveWidget = function(widget, file, selfcontained = TRUE, ...) {
      saved$widget <- widget
      saved$file <- file
      writeLines("<html></html>", file)
      invisible(NULL)
    },
    .package = "htmlwidgets"
  )

  old <- setwd(tempdir())
  on.exit(setwd(old), add = TRUE)

  suppressMessages(
    forplot_map(
      fp_file_path = input_file,
      input_type = "field_sheet",
      vertex_coords = .map_test_vertices(),
      plot_size = 0.01,
      subplot_size = 10,
      voucher_imgs = img_root,
      filename = "voucher_carousel"
    )
  )

  expect_true(.map_contains_text(saved$widget, "Has Photo"))
  expect_true(.map_contains_text(saved$widget, "slideshow-container"))
  expect_true(.map_contains_text(saved$widget, "specimen.JPG"))
})

test_that("voucher_imgs = NULL does not require image-directory traversal", {
  input_file <- .map_test_input_file()
  saved <- new.env(parent = emptyenv())
  here_called <- new.env(parent = emptyenv())
  here_called$value <- FALSE

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .package = "forplotR"
  )

  # A NULL image directory should bypass here::here() entirely.
  testthat::local_mocked_bindings(
    here = function(...) {
      here_called$value <- TRUE
      stop("here::here() should not be called when voucher_imgs is NULL")
    },
    .package = "here"
  )

  testthat::local_mocked_bindings(
    saveWidget = function(widget, file, selfcontained = TRUE, ...) {
      saved$file <- file
      writeLines("<html></html>", file)
      invisible(NULL)
    },
    .package = "htmlwidgets"
  )

  old <- setwd(tempdir())
  on.exit(setwd(old), add = TRUE)

  expect_no_error(
    suppressMessages(
      forplot_map(
        fp_file_path = input_file,
        input_type = "field_sheet",
        vertex_coords = .map_test_vertices(),
        plot_size = 0.01,
        subplot_size = 10,
        voucher_imgs = NULL,
        filename = "no_voucher_directory"
      )
    )
  )

  expect_false(here_called$value)
  expect_true(file.exists(saved$file))
})


# -------------------------------------------------------------------------
# HTML widget/output contract
# -------------------------------------------------------------------------

test_that("forplot_map creates a Leaflet htmlwidget and requests self-contained HTML", {
  input_file <- .map_test_input_file()
  saved <- new.env(parent = emptyenv())

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .package = "forplotR"
  )

  testthat::local_mocked_bindings(
    saveWidget = function(widget, file, selfcontained = TRUE, ...) {
      saved$widget <- widget
      saved$file <- file
      saved$selfcontained <- selfcontained
      writeLines("<html></html>", file)
      invisible(NULL)
    },
    .package = "htmlwidgets"
  )

  old <- setwd(tempdir())
  on.exit(setwd(old), add = TRUE)

  out <- suppressMessages(
    forplot_map(
      fp_file_path = input_file,
      input_type = "field_sheet",
      vertex_coords = .map_test_vertices(),
      plot_size = 0.01,
      subplot_size = 10,
      voucher_imgs = .map_test_missing_voucher_dir(),
      filename = "widget_contract"
    )
  )

  expect_s3_class(saved$widget, "leaflet")
  expect_true(isTRUE(saved$selfcontained))
  expect_true(file.exists(saved$file))
  expect_true(grepl("^Results_", basename(saved$file)))
  expect_true(endsWith(saved$file, "_widget_contract.html"))

  expect_identical(
    normalizePath(out, winslash = "/", mustWork = FALSE),
    normalizePath(saved$file, winslash = "/", mustWork = FALSE)
  )
})

test_that("generated widget contains specimen and subplot layers", {
  input_file <- .map_test_input_file()
  saved <- new.env(parent = emptyenv())

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .package = "forplotR"
  )

  testthat::local_mocked_bindings(
    saveWidget = function(widget, file, selfcontained = TRUE, ...) {
      saved$widget <- widget
      saved$file <- file
      writeLines("<html></html>", file)
      invisible(NULL)
    },
    .package = "htmlwidgets"
  )

  old <- setwd(tempdir())
  on.exit(setwd(old), add = TRUE)

  suppressMessages(
    forplot_map(
      fp_file_path = input_file,
      input_type = "field_sheet",
      vertex_coords = .map_test_vertices(),
      plot_size = 0.01,
      subplot_size = 10,
      voucher_imgs = .map_test_missing_voucher_dir(),
      filename = "layers"
    )
  )

  expect_true(.map_contains_text(saved$widget, "Specimens"))
  expect_true(.map_contains_text(saved$widget, "Subplot Grid"))
  expect_true(.map_contains_text(saved$widget, "OSM Street"))
  expect_true(.map_contains_text(saved$widget, "Satellite"))
})

test_that("generated widget contains expected popup metadata", {
  input_file <- .map_test_input_file()
  saved <- new.env(parent = emptyenv())

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .package = "forplotR"
  )

  testthat::local_mocked_bindings(
    saveWidget = function(widget, file, selfcontained = TRUE, ...) {
      saved$widget <- widget
      saved$file <- file
      writeLines("<html></html>", file)
      invisible(NULL)
    },
    .package = "htmlwidgets"
  )

  old <- setwd(tempdir())
  on.exit(setwd(old), add = TRUE)

  suppressMessages(
    forplot_map(
      fp_file_path = input_file,
      input_type = "field_sheet",
      vertex_coords = .map_test_vertices(),
      plot_size = 0.01,
      subplot_size = 10,
      voucher_imgs = .map_test_missing_voucher_dir(),
      filename = "popup_metadata"
    )
  )

  expect_true(.map_contains_text(saved$widget, "Test Plot"))
  expect_true(.map_contains_text(saved$widget, "TEST01"))
  expect_true(.map_contains_text(saved$widget, "Test Team"))
  expect_true(.map_contains_text(saved$widget, "Inga alba"))
  expect_true(.map_contains_text(saved$widget, "Unvouchered"))
})



test_that("specimen marker status colors are embedded in the Leaflet widget", {
  input_file <- .map_test_input_file()
  saved <- new.env(parent = emptyenv())

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .package = "forplotR"
  )

  testthat::local_mocked_bindings(
    saveWidget = function(widget, file, selfcontained = TRUE, ...) {
      saved$widget <- widget
      saved$file <- file
      writeLines("<html></html>", file)
      invisible(NULL)
    },
    .package = "htmlwidgets"
  )

  old <- setwd(tempdir())
  on.exit(setwd(old), add = TRUE)

  suppressMessages(
    forplot_map(
      fp_file_path = input_file,
      input_type = "field_sheet",
      vertex_coords = .map_test_vertices(),
      plot_size = 0.01,
      subplot_size = 10,
      voucher_imgs = .map_test_missing_voucher_dir(),
      filename = "status_colors"
    )
  )

  # Fixture contains one collected non-palm, one palm and one uncollected
  # non-palm individual.
  expect_true(.map_contains_text(saved$widget, "gray"))
  expect_true(.map_contains_text(saved$widget, "gold"))
  expect_true(.map_contains_text(saved$widget, "red"))
})

# -------------------------------------------------------------------------
# Side-effect safety
# -------------------------------------------------------------------------

test_that("CRAN smoke test does not require internet or a real Excel workbook", {
  input_file <- .map_test_input_file()
  saved <- new.env(parent = emptyenv())

  testthat::local_mocked_bindings(
    .harmonize_plot_input = .map_mock_standard_input,
    .package = "forplotR"
  )

  testthat::local_mocked_bindings(
    saveWidget = function(widget, file, selfcontained = TRUE, ...) {
      saved$file <- file
      writeLines("<html></html>", file)
      invisible(NULL)
    },
    .package = "htmlwidgets"
  )

  old <- setwd(tempdir())
  on.exit(setwd(old), add = TRUE)

  expect_no_error(
    suppressMessages(
      forplot_map(
        fp_file_path = input_file,
        input_type = "field_sheet",
        vertex_coords = .map_test_vertices(),
        plot_size = 0.01,
        subplot_size = 10,
        voucher_imgs = .map_test_missing_voucher_dir(),
        filename = "cran_smoke",
        herbaria_lookup = FALSE,
        verbose = FALSE
      )
    )
  )

  expect_true(file.exists(saved$file))
})
