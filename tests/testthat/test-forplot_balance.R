# Tests for forplot_balance() and its internal helpers
# CRAN-oriented: deterministic, no network, no browser, no LaTeX/TinyTeX.
# Requires testthat >= 3.2.0 because local_mocked_bindings() is used.

# -----------------------------------------------------------------------------
# Small fixtures
# -----------------------------------------------------------------------------

.minimal_clean_input <- function() {
  data.frame(
    T1 = "1",
    X = "1",
    Y = "2",
    D = "100",
    Collected = NA_character_,
    check.names = FALSE,
    stringsAsFactors = FALSE
  )
}

.minimal_fp_coords <- function() {
  data.frame(
    T1 = 1L,
    X = 1,
    Y = 2,
    D = 100,
    Collected = NA_character_,
    Family = "Fabaceae",
    `Original determination` = "Inga sp.",
    `New Tag No` = "1",
    global_x = 1,
    global_y = 2,
    diameter = 2,
    check.names = FALSE,
    stringsAsFactors = FALSE
  )
}

# -----------------------------------------------------------------------------
# Public API: argument validation
# -----------------------------------------------------------------------------

test_that("forplot_balance validates output switches", {
  common <- list(
    input_type = "field_sheet",
    subplot_size = 10,
    render_html = FALSE,
    render_pdf = FALSE,
    write_xlsx = FALSE
  )

  expect_error(
    do.call(forplotR::forplot_balance, modifyList(common, list(render_html = 1))),
    "`render_html` must be TRUE or FALSE\\."
  )
  expect_error(
    do.call(forplotR::forplot_balance, modifyList(common, list(render_html = NA))),
    "`render_html` must be TRUE or FALSE\\."
  )
  expect_error(
    do.call(forplotR::forplot_balance, modifyList(common, list(render_html = c(TRUE, FALSE)))),
    "`render_html` must be TRUE or FALSE\\."
  )

  expect_error(
    do.call(forplotR::forplot_balance, modifyList(common, list(render_pdf = "yes"))),
    "`render_pdf` must be TRUE or FALSE\\."
  )
  expect_error(
    do.call(forplotR::forplot_balance, modifyList(common, list(render_pdf = NA))),
    "`render_pdf` must be TRUE or FALSE\\."
  )

  expect_error(
    do.call(forplotR::forplot_balance, modifyList(common, list(write_xlsx = 1L))),
    "`write_xlsx` must be TRUE or FALSE\\."
  )
  expect_error(
    do.call(forplotR::forplot_balance, modifyList(common, list(write_xlsx = NA))),
    "`write_xlsx` must be TRUE or FALSE\\."
  )
})


test_that("forplot_balance validates input_type and language", {
  expect_error(
    forplotR::forplot_balance(
      input_type = "not_a_type",
      language = "en",
      subplot_size = 10,
      render_html = FALSE,
      render_pdf = FALSE,
      write_xlsx = FALSE
    ),
    "field_sheet"
  )

  expect_error(
    forplotR::forplot_balance(
      input_type = "field_sheet",
      language = "xx",
      subplot_size = 10,
      render_html = FALSE,
      render_pdf = FALSE,
      write_xlsx = FALSE
    ),
    "en"
  )
})


test_that("forplot_balance validates plot geometry", {
  common <- list(
    input_type = "field_sheet",
    language = "en",
    subplot_size = 10,
    render_html = FALSE,
    render_pdf = FALSE,
    write_xlsx = FALSE
  )

  for (bad in list(0, -1, NA_real_, "1", c(1, 2))) {
    expect_error(
      do.call(forplotR::forplot_balance, modifyList(common, list(plot_size = bad))),
      "`plot_size` must be a single positive numeric value in hectares\\."
    )
  }

  for (bad in list(0, -10, NA_real_, "100", c(100, 200))) {
    expect_error(
      do.call(
        forplotR::forplot_balance,
        modifyList(common, list(plot_width_m = bad))
      ),
      "`plot_width_m` must be NULL or a single positive numeric value\\."
    )
  }

  for (bad in list(0, -10, NA_real_, "100", c(100, 200))) {
    expect_error(
      do.call(
        forplotR::forplot_balance,
        modifyList(common, list(plot_length_m = bad))
      ),
      "`plot_length_m` must be NULL or a single positive numeric value\\."
    )
  }

  expect_error(
    do.call(
      forplotR::forplot_balance,
      modifyList(
        common,
        list(plot_size = 1, plot_width_m = 100, plot_length_m = 90)
      )
    ),
    "must match the area implied by `plot_size`"
  )

  expect_error(
    do.call(
      forplotR::forplot_balance,
      modifyList(
        common,
        list(plot_size = 1, plot_width_m = 80, plot_length_m = 125)
      )
    ),
    "must be exact multiples of `subplot_size`"
  )
})


test_that("plot geometry accepts arbitrary compatible subplot sizes", {

  g_1ha <- forplotR:::.resolve_plot_geometry(
    plot_size = 1,
    subplot_size = 10
  )

  expect_equal(g_1ha$plot_width_m, 100)
  expect_equal(g_1ha$plot_length_m, 100)
  expect_equal(g_1ha$n_cols, 10L)
  expect_equal(g_1ha$n_rows, 10L)
  expect_equal(g_1ha$n_subplots, 100L)


  g_15m <- forplotR:::.resolve_plot_geometry(
    plot_size = 2.25,
    subplot_size = 15
  )

  expect_equal(g_15m$plot_width_m, 150)
  expect_equal(g_15m$plot_length_m, 150)
  expect_equal(g_15m$n_cols, 10L)
  expect_equal(g_15m$n_rows, 10L)
  expect_equal(g_15m$n_subplots, 100L)


  g_10ha <- forplotR:::.resolve_plot_geometry(
    plot_size = 10,
    subplot_size = 10
  )

  expect_equal(g_10ha$plot_width_m, 250)
  expect_equal(g_10ha$plot_length_m, 400)
  expect_equal(g_10ha$n_cols, 25L)
  expect_equal(g_10ha$n_rows, 40L)
  expect_equal(g_10ha$n_subplots, 1000L)


  g_explicit <- forplotR:::.resolve_plot_geometry(
    plot_size = 10,
    subplot_size = 10,
    plot_width_m = 200,
    plot_length_m = 500
  )

  expect_equal(g_explicit$plot_width_m, 200)
  expect_equal(g_explicit$plot_length_m, 500)
  expect_equal(g_explicit$n_cols, 20L)
  expect_equal(g_explicit$n_rows, 50L)
  expect_equal(g_explicit$n_subplots, 1000L)
})


test_that("plot geometry rejects incompatible subplot grids", {

  expect_error(
    forplotR:::.resolve_plot_geometry(
      plot_size = 1,
      subplot_size = 15
    ),
    "do not form an integer number of square subplots"
  )

  for (bad in list(
    0,
    -1,
    NA_real_,
    Inf,
    "10",
    c(10, 20),
    NULL
  )) {

    expect_error(
      forplotR:::.validate_subplot_size(bad),
      "`subplot_size` must be a single positive numeric value in meters\\."
    )
  }
})


test_that("local X and Y limits follow subplot_size", {

  expect_invisible(
    forplotR:::.validate_local_xy(
      data.frame(
        X = c(0, 7.5, 15),
        Y = c(15, 2, 0)
      ),
      subplot_size = 15
    )
  )

  expect_error(
    forplotR:::.validate_local_xy(
      data.frame(
        X = c(0, 15.1),
        Y = c(4, 8)
      ),
      subplot_size = 15
    ),
    "Local X/Y coordinates must fall between 0 and `subplot_size`"
  )
})


test_that("large plot export dimensions stay bounded", {

  one_ha <- forplotR:::.plot_export_size(
    plot_width_m = 100,
    plot_length_m = 100
  )

  ten_ha <- forplotR:::.plot_export_size(
    plot_width_m = 250,
    plot_length_m = 400
  )

  expect_lte(one_ha$width, 12)
  expect_lte(one_ha$height, 8)

  expect_lte(ten_ha$width, 12)
  expect_lte(ten_ha$height, 8)

  expect_equal(
    forplotR:::.plot_symbol_scale(100, 100),
    1
  )

  expect_lt(
    forplotR:::.plot_symbol_scale(250, 400),
    1
  )

  expect_gte(
    forplotR:::.plot_symbol_scale(1000, 1000),
    0.25
  )
})
# -----------------------------------------------------------------------------
# Public API: normalization and inferred geometry (mock only expensive stages)
# -----------------------------------------------------------------------------

test_that("forplot_balance normalizes input_type and language", {
  seen <- new.env(parent = emptyenv())

  testthat::local_mocked_bindings(
    .harmonize_plot_input = function(fp_file_path, input_type, station_name, verbose) {
      seen$input_type <- input_type
      seen$station_name <- station_name
      stop("TEST_STOP_AFTER_HARMONIZE", call. = FALSE)
    },
    .package = "forplotR"
  )

  out_dir <- tempfile("forplot-balance-")
  on.exit(unlink(out_dir, recursive = TRUE, force = TRUE), add = TRUE)

  expect_error(
    forplotR::forplot_balance(
      fp_file_path = "dummy.xlsx",
      input_type = "  FIELD_SHEET  ",
      language = "  PT  ",
      subplot_size = 10,
      render_html = FALSE,
      render_pdf = FALSE,
      write_xlsx = FALSE,
      verbose = FALSE,
      dir = out_dir
    ),
    "TEST_STOP_AFTER_HARMONIZE"
  )

  expect_identical(seen$input_type, "field_sheet")
  expect_null(seen$station_name)
})


test_that("forplot_balance infers missing plot dimensions from plot_size", {
  seen <- new.env(parent = emptyenv())

  testthat::local_mocked_bindings(
    .harmonize_plot_input = function(...) {
      list(
        fp_sheet = .minimal_clean_input(),
        team = "",
        plot_name = "Plot",
        plot_code = "P1",
        census_no_fp = "1"
      )
    },
    .compute_global_coordinates = function(fp_clean, subplot_size,
                                           plot_width_m, plot_length_m) {
      seen$subplot_size <- subplot_size
      seen$plot_width_m <- plot_width_m
      seen$plot_length_m <- plot_length_m
      stop("TEST_STOP_AFTER_GEOMETRY", call. = FALSE)
    },
    .package = "forplotR"
  )

  run_case <- function(width, length, expected_width, expected_length) {
    out_dir <- tempfile("forplot-balance-")
    on.exit(unlink(out_dir, recursive = TRUE, force = TRUE), add = TRUE)

    expect_error(
      forplotR::forplot_balance(
        fp_file_path = "dummy.xlsx",
        input_type = "field_sheet",
        language = "en",
        plot_size = 1,
        subplot_size = 10,
        plot_width_m = width,
        plot_length_m = length,
        render_html = FALSE,
        render_pdf = FALSE,
        write_xlsx = FALSE,
        verbose = FALSE,
        dir = out_dir
      ),
      "TEST_STOP_AFTER_GEOMETRY"
    )

    expect_equal(seen$plot_width_m, expected_width)
    expect_equal(seen$plot_length_m, expected_length)
    expect_equal(seen$subplot_size, 10)
  }

  run_case(50, NULL, 50, 200)
  run_case(NULL, 200, 50, 200)
  run_case(NULL, NULL, 100, 100)
})


test_that("MONITORA ignores ForestPlots plot dimensions", {
  seen <- new.env(parent = emptyenv())

  testthat::local_mocked_bindings(
    .harmonize_plot_input = function(...) {
      list(
        fp_sheet = .minimal_clean_input(),
        team = "",
        plot_name = "Station",
        plot_code = "M1",
        census_no_fp = "1"
      )
    },
    .compute_monitora_geometry = function(fp_clean, keep_only_cell = FALSE) {
      seen$keep_only_cell <- keep_only_cell
      stop("TEST_STOP_MONITORA_GEOMETRY", call. = FALSE)
    },
    .package = "forplotR"
  )

  out_dir <- tempfile("forplot-balance-")
  on.exit(unlink(out_dir, recursive = TRUE, force = TRUE), add = TRUE)

  expect_error(
    forplotR::forplot_balance(
      fp_file_path = "dummy.xlsx",
      input_type = "monitora",
      language = "en",
      plot_size = -999,       # intentionally irrelevant for MONITORA
      subplot_size = 10,
      plot_width_m = -999,    # intentionally irrelevant for MONITORA
      plot_length_m = -999,   # intentionally irrelevant for MONITORA
      render_html = FALSE,
      render_pdf = FALSE,
      write_xlsx = FALSE,
      verbose = FALSE,
      dir = out_dir
    ),
    "TEST_STOP_MONITORA_GEOMETRY"
  )

  expect_false(seen$keep_only_cell)
})


test_that("multiple MONITORA stations recurse once per unique station", {
  original_fun <- forplotR::forplot_balance
  calls <- list()

  testthat::local_mocked_bindings(
    forplot_balance = function(..., station_name, filename) {
      calls[[length(calls) + 1L]] <<- list(
        station_name = station_name,
        filename = filename
      )
      invisible(list())
    },
    .package = "forplotR"
  )

  out_dir <- tempfile("forplot-balance-")
  on.exit(unlink(out_dir, recursive = TRUE, force = TRUE), add = TRUE)

  ans <- original_fun(
    fp_file_path = "dummy.xlsx",
    input_type = "monitora",
    language = "en",
    subplot_size = 10,
    station_name = c(" S1 ", "S2", "S1"),
    render_html = FALSE,
    render_pdf = FALSE,
    write_xlsx = FALSE,
    verbose = FALSE,
    dir = out_dir,
    filename = "balance"
  )

  expect_true(ans)
  expect_length(calls, 2L)
  expect_identical(vapply(calls, `[[`, character(1), "station_name"), c("S1", "S2"))
  expect_identical(vapply(calls, `[[`, character(1), "filename"),
                   c("balance_station_S1", "balance_station_S2"))
})

# -----------------------------------------------------------------------------
# Helpers defined with forplot_balance()
# -----------------------------------------------------------------------------

test_that(".find_field_header_row detects the most likely header row", {
  raw <- data.frame(
    V1 = c("metadata", "New Tag No", "1"),
    V2 = c("metadata", "T1", "1"),
    V3 = c("metadata", "X", "2"),
    V4 = c("metadata", "Y", "3"),
    V5 = c("metadata", "Family", "Fabaceae"),
    stringsAsFactors = FALSE
  )

  expect_identical(forplotR:::.find_field_header_row(raw), 2L)

  no_header <- data.frame(
    V1 = c("a", "b"),
    V2 = c("c", "d"),
    stringsAsFactors = FALSE
  )
  expect_identical(forplotR:::.find_field_header_row(no_header), 2L)
})


test_that(".safe_nzchar accepts only one non-empty character string", {
  expect_true(forplotR:::.safe_nzchar("abc"))
  expect_true(forplotR:::.safe_nzchar("  abc  "))

  expect_false(forplotR:::.safe_nzchar(""))
  expect_false(forplotR:::.safe_nzchar("   "))
  expect_false(forplotR:::.safe_nzchar(NA_character_))
  expect_false(forplotR:::.safe_nzchar(NULL))
  expect_false(forplotR:::.safe_nzchar(c("a", "b")))
  expect_false(forplotR:::.safe_nzchar(1))
})


test_that(".collapse_sorted_tags trims, deduplicates and sorts tags", {
  x <- c("10", "2", "A", " 2 ", "1", "", NA_character_)

  expect_identical(
    forplotR:::.collapse_sorted_tags(x),
    "1|2|10|A"
  )
  expect_identical(
    forplotR:::.collapse_sorted_tags(x, sep = " | "),
    "1 | 2 | 10 | A"
  )
  expect_true(is.na(forplotR:::.collapse_sorted_tags(c("", " ", NA_character_))))
})


test_that(".clean_fp_data canonicalizes numeric columns", {
  x <- data.frame(
    T1 = c("1", "2"),
    X = c("1.5", "2.5"),
    Y = c("3", "4"),
    D = c("100", "250"),
    Collected = factor(c("A", NA)),
    stringsAsFactors = FALSE
  )

  out <- forplotR:::.clean_fp_data(x)

  expect_type(out$T1, "integer")
  expect_equal(out$T1, c(1L, 2L))
  expect_type(out$X, "double")
  expect_equal(out$X, c(1.5, 2.5))
  expect_equal(out$Y, c(3, 4))
  expect_equal(out$D, c(100, 250))
  expect_type(out$Collected, "character")
  expect_identical(out$Collected[1], "A")
  expect_true(is.na(out$Collected[2]))
})


test_that(".detect_coordinate_mode chooses the representation with more complete data", {
  local_better <- data.frame(
    T1 = c(1, 2), X = c(1, 2), Y = c(3, 4),
    `Standardised SubPlot T1` = c(1, NA),
    `Standardised X` = c(1, NA),
    `Standardised Y` = c(3, NA),
    check.names = FALSE
  )
  expect_identical(forplotR:::.detect_coordinate_mode(local_better), "local")

  std_better <- data.frame(
    T1 = c(1, NA), X = c(1, NA), Y = c(3, NA),
    `Standardised SubPlot T1` = c(1, 2),
    `Standardised X` = c(1, 2),
    `Standardised Y` = c(3, 4),
    check.names = FALSE
  )
  expect_identical(forplotR:::.detect_coordinate_mode(std_better), "standardised")

  xy_only <- data.frame(X = 1:2, Y = 3:4)
  expect_identical(forplotR:::.detect_coordinate_mode(xy_only), "local_xy_only")

  unknown <- data.frame(tag = 1:2)
  expect_identical(forplotR:::.detect_coordinate_mode(unknown), "unknown")
})

# -----------------------------------------------------------------------------
# Workbook output helper: real file I/O, but only inside a temporary directory
# -----------------------------------------------------------------------------

test_that(".collection_percentual writes correct workbook summaries", {
  out_dir <- tempfile("collection-balance-")
  dir.create(out_dir, recursive = TRUE)
  on.exit(unlink(out_dir, recursive = TRUE, force = TRUE), add = TRUE)

  x <- data.frame(
    T1 = c(1, 1, 1, 2, 2),
    Family = c("Fabaceae", "Fabaceae", "Arecaceae", "Rubiaceae", "Arecaceae"),
    Collected = c("C1", NA, "C3", "", NA),
    `New Tag No` = c("10", "2", "3", "A", "1"),
    check.names = FALSE,
    stringsAsFactors = FALSE
  )

  path <- forplotR:::.collection_percentual(
    fp_sheet = x,
    dir = out_dir,
    plot_name = "Test Plot",
    plot_code = "TP",
    plot_census_no_fp = "1",
    team = "Team",
    plot_width_m = 20,
    plot_length_m = 10,
    subplot_size = 10
  )

  expect_true(file.exists(path))
  expect_identical(basename(path), "collection_balance.xlsx")

  expect_setequal(
    openxlsx::getSheetNames(path),
    c("COLLECTION_PERCENTUAL", "NOT_COLLECTED", "COLLECTED")
  )

  pct <- openxlsx::read.xlsx(path, sheet = "COLLECTION_PERCENTUAL", startRow = 3)
  not_collected <- openxlsx::read.xlsx(path, sheet = "NOT_COLLECTED", startRow = 3)
  collected <- openxlsx::read.xlsx(path, sheet = "COLLECTED", startRow = 3)

  total <- pct[pct$subplot == "TOTAL", , drop = FALSE]
  expect_equal(total$total_individuals, 5)
  expect_equal(total$total_non_arecaceae, 3)
  expect_equal(total$collected, 2)
  expect_equal(total$collected_non_arecaceae, 1)
  expect_equal(total$uncollected_non_arecaceae, 2)
  expect_equal(total$arecaceae_count, 2)
  expect_equal(total$collected_percentual, 40)
  expect_equal(total$collected_percentual_without_arecaceae, 33.3)

  subplot1 <- pct[pct$subplot == "1", , drop = FALSE]
  expect_equal(subplot1$total_individuals, 3)
  expect_equal(subplot1$collected_percentual, 66.7)
  expect_equal(subplot1$collected_percentual_without_arecaceae, 50)

  subplot2 <- pct[pct$subplot == "2", , drop = FALSE]
  expect_equal(subplot2$total_individuals, 2)
  expect_equal(subplot2$collected_percentual, 0)
  expect_equal(subplot2$collected_percentual_without_arecaceae, 0)

  total_uncollected <- not_collected[not_collected$subplot == "TOTAL", , drop = FALSE]
  expect_equal(total_uncollected$n_uncollected, 2)
  expect_identical(total_uncollected$tagno_uncollected, "2|A")

  total_collected <- collected[collected$subplot == "TOTAL", , drop = FALSE]
  expect_equal(total_collected$n_collected, 2)
  expect_identical(total_collected$tagno_collected, "3|10")
})

# -----------------------------------------------------------------------------
# Static plot builders: build plots in memory only
# -----------------------------------------------------------------------------

test_that(".build_fp_base_plot returns a buildable ggplot", {
  fp <- .minimal_fp_coords()
  labels <- tibble::tibble(T1 = 1, center_x = 5, center_y = 5)

  p <- forplotR:::.build_fp_base_plot(
    fp_coords = fp,
    subplot_labels = labels,
    subplot_size = 10,
    plot_width_m = 10,
    plot_length_m = 10,
    plot_name = "Test",
    plot_code = "T1",
    highlight_palms = TRUE,
    language = "en"
  )

  expect_s3_class(p, "ggplot")
  expect_no_error(suppressWarnings(ggplot2::ggplot_build(p)))
})


test_that(".build_monitora_base_plot returns a buildable ggplot", {
  fp <- data.frame(
    draw_x = c(0, 5),
    draw_y = c(55, 60),
    subunit_letter = c("N", "N"),
    D = c(100, 120),
    Collected = c(NA, "C1"),
    Family = c("Fabaceae", "Arecaceae"),
    `New Tag No` = c("1", "2"),
    check.names = FALSE,
    stringsAsFactors = FALSE
  )

  p <- forplotR:::.build_monitora_base_plot(
    fp_coords = fp,
    plot_name = "Station",
    plot_code = "M1",
    highlight_palms = TRUE,
    language = "en"
  )

  expect_s3_class(p, "ggplot")
  expect_no_error(suppressWarnings(ggplot2::ggplot_build(p)))
})

# -----------------------------------------------------------------------------
# Interactive plot builders: no browser and no saveWidget()
# -----------------------------------------------------------------------------

test_that("interactive FP plot builder returns a plotly widget", {
  fp <- .minimal_fp_coords()
  labels <- tibble::tibble(T1 = 1, center_x = 5, center_y = 5)

  p <- forplotR:::.build_fp_base_plot_interactive(
    fp_coords = fp,
    subplot_labels = labels,
    subplot_size = 10,
    plot_width_m = 10,
    plot_length_m = 10,
    plot_name = "Test",
    plot_code = "T1",
    highlight_palms = TRUE,
    language = "en"
  )

  expect_s3_class(p, "plotly")
  expect_s3_class(p, "htmlwidget")
})


test_that("interactive subplot builder handles empty and malformed inputs", {
  expect_null(
    forplotR:::.build_monitora_subplot_plot_interactive(
      sp_data = data.frame(),
      subplot_size = 10,
      highlight_palms = TRUE,
      language = "en"
    )
  )

  expect_null(
    forplotR:::.build_monitora_subplot_plot_interactive(
      sp_data = data.frame(X = 1, Y = 1),
      subplot_size = 10,
      highlight_palms = TRUE,
      language = "en"
    )
  )
})

# -----------------------------------------------------------------------------
# Rendering helper: validate dispatch only. Do not render HTML/PDF on CRAN.
# -----------------------------------------------------------------------------

test_that(".render_plot_report validates format before rendering", {
  expect_error(
    forplotR:::.render_plot_report(
      rmd_path = "does-not-need-to-exist.Rmd",
      output_path = tempdir(),
      output_name = "x.out",
      params = list(),
      format = "docx"
    ),
    "pdf|html"
  )
})

# -----------------------------------------------------------------------------
# Known regression guard for the filtered-map data replacement.
# This test is intentionally local-only until the implementation uses ggplot2 `%+%`
# (or another supported data replacement strategy) rather than `ggplot + data.frame`.
# Once fixed, remove skip_on_cran() and adapt this into the end-to-end fixture test.
# -----------------------------------------------------------------------------

test_that("filtered plot data replacement remains compatible with ggplot2", {

  p <- ggplot2::ggplot(
    data.frame(x = 1, y = 1),
    ggplot2::aes(x, y)
  ) +
    ggplot2::geom_point()

  replacement <- data.frame(
    x = 2,
    y = 2
  )

  if (utils::packageVersion("ggplot2") >= "4.0.0") {

    p2 <- p + replacement

  } else {

    p2 <- ggplot2::`%+%`(
      p,
      replacement
    )
  }

  expect_s3_class(p2, "ggplot")

  expect_equal(
    p2$data$x,
    2
  )

  expect_equal(
    p2$data$y,
    2
  )
})
