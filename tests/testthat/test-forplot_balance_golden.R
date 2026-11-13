# Golden-master (characterization) tests for forplot_balance()
#
# These tests freeze the *current* behaviour of the report pipeline so that the
# report code can be refactored (data preparation split from rendering, Rmd
# template moved to inst/templates/) without silently changing the output.
#
# What is frozen, for the bundled datasets RUS01 (field_sheet) and TUM
# (MONITORA):
#   1. the Rmd text produced by .create_rmd_content();
#   2. a plain-text dump of every data object passed as `params` to the report.
#
# Rendering itself (rmarkdown/LaTeX/plotly) and PNG export are mocked out, so
# the tests are deterministic and need no LaTeX, browser or network.
#
# If a change to the report is intentional, review the diff and accept it with
# testthat::snapshot_review() / snapshot_accept("forplot_balance_golden").

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------

.golden_fmt_col <- function(x) {
  if (is.list(x)) {
    return(vapply(x, function(el) paste(sort(as.character(unlist(el))), collapse = ";"),
                  character(1)))
  }
  if (is.factor(x)) x <- as.character(x)
  if (is.numeric(x)) x <- ifelse(is.na(x), NA_character_, format(signif(x, 8), trim = TRUE,
                                                                  scientific = FALSE))
  x <- as.character(x)
  x[is.na(x)] <- "NA"
  x
}

.golden_dump_df <- function(df) {
  df <- as.data.frame(df, stringsAsFactors = FALSE, check.names = FALSE)
  if (!ncol(df)) return("<0 columns>")
  cols <- lapply(df, .golden_fmt_col)
  header <- paste(names(df), collapse = "\t")
  body <- if (nrow(df)) do.call(paste, c(cols, sep = "\t")) else character(0)
  c(sprintf("<%d x %d>", nrow(df), ncol(df)), header, body)
}

# Data shown by a ggplot, restricted to the columns the report depends on.
.golden_plot_data <- function(p, cols = c("New Tag No", "T1", "T2", "subunit_letter",
                                          "Family", "Original determination",
                                          "Collected", "Status", "D", "diameter",
                                          "X", "Y", "global_x", "global_y",
                                          "x10", "y10", "draw_x", "draw_y")) {
  if (is.null(p)) return("<NULL>")
  d <- p$data
  if (!is.data.frame(d)) return(sprintf("<%s>", class(d)[1]))
  .golden_dump_df(d[, intersect(cols, names(d)), drop = FALSE])
}

.golden_dump_value <- function(x) {
  if (is.null(x)) return("<NULL>")
  if (inherits(x, "ggplot")) return(.golden_plot_data(x))
  if (inherits(x, "htmlwidget")) return(sprintf("<htmlwidget %s>", class(x)[1]))
  if (is.data.frame(x)) return(.golden_dump_df(x))
  if (is.list(x)) {
    out <- character(0)
    nms <- names(x)
    if (is.null(nms)) nms <- as.character(seq_along(x))
    for (i in seq_along(x)) {
      out <- c(out, paste0("$", nms[i]), paste0("  ", .golden_dump_value(x[[i]])))
    }
    return(out)
  }
  paste(.golden_fmt_col(x), collapse = " | ")
}

.golden_dump_params <- function(params) {
  keys <- sort(names(params))
  out <- character(0)
  for (k in keys) {
    out <- c(out, "", paste0("## ", k), .golden_dump_value(params[[k]]))
  }
  out
}

# Runs forplot_balance() with rendering and PNG export mocked, returning what
# would have been rendered.
.golden_run_balance <- function(data, input_type, ..., env = parent.frame()) {
  td <- withr::local_tempdir(.local_envir = env)
  input <- file.path(td, "input.xlsx")
  writexl::write_xlsx(data, input)

  cap <- new.env(parent = emptyenv())

  testthat::local_mocked_bindings(
    .render_plot_report = function(rmd_path, output_path, output_name, params,
                                   format = c("pdf", "html"), ...) {
      cap$rmd <- readLines(rmd_path, encoding = "UTF-8", warn = FALSE)
      cap$params <- params
      cap$format <- format
      invisible(NULL)
    },
    .package = "forplotR",
    .env = env
  )
  testthat::local_mocked_bindings(
    ggsave = function(filename, ...) invisible(filename),
    .package = "ggplot2",
    .env = env
  )

  res <- suppressWarnings(forplot_balance(
    fp_file_path = input,
    input_type = input_type,
    render_html = FALSE,
    render_pdf = TRUE,
    write_xlsx = FALSE,
    verbose = FALSE,
    dir = file.path(td, "out"),
    filename = "golden",
    ...
  ))

  list(result = res, rmd = cap$rmd, params = cap$params)
}

.golden_write <- function(lines, name) {
  path <- file.path(tempdir(), name)
  con <- file(path, open = "w", encoding = "UTF-8")
  on.exit(close(con))
  writeLines(lines, con, useBytes = FALSE)
  path
}

# RUS01 contains one individual without subplot (tag 1700), kept on purpose:
# forplot_balance() must exclude it with a warning, not abort.
.golden_rus01 <- function() RUS01

test_that("RUS01 keeps the field-sheet metadata headers readable", {
  expect_true("Plot Name: REBIO Augusto Ruschi" %in% names(RUS01))
  expect_true("Plotcode: RUS01" %in% names(RUS01))
})

# -----------------------------------------------------------------------------
# Tests
# -----------------------------------------------------------------------------

test_that("forplot_balance report is unchanged for a ForestPlots field sheet (RUS01)", {
  skip_on_cran()
  skip_if_not_installed("writexl")
  skip_if_not_installed("withr")

  out <- .golden_run_balance(.golden_rus01(), input_type = "field_sheet", language = "en")

  expect_type(out$rmd, "character")
  expect_type(out$params, "list")

  expect_snapshot_file(.golden_write(out$rmd, "rus01_report.Rmd"),
                       "rus01_report.Rmd")
  expect_snapshot_file(.golden_write(.golden_dump_params(out$params), "rus01_params.txt"),
                       "rus01_params.txt")
})

test_that("forplot_balance report is unchanged for a ForestPlots field sheet in Portuguese", {
  skip_on_cran()
  skip_if_not_installed("writexl")
  skip_if_not_installed("withr")

  out <- .golden_run_balance(.golden_rus01(), input_type = "field_sheet", language = "pt")

  expect_snapshot_file(.golden_write(out$rmd, "rus01_report_pt.Rmd"),
                       "rus01_report_pt.Rmd")
})

test_that("forplot_balance report is unchanged for MONITORA data (TUM)", {
  skip_on_cran()
  skip_if_not_installed("writexl")
  skip_if_not_installed("withr")

  out <- .golden_run_balance(TUM, input_type = "monitora", station_name = "EA-1",
                             language = "en")

  expect_type(out$rmd, "character")
  expect_type(out$params, "list")

  expect_snapshot_file(.golden_write(out$rmd, "tum_report.Rmd"),
                       "tum_report.Rmd")
  expect_snapshot_file(.golden_write(.golden_dump_params(out$params), "tum_params.txt"),
                       "tum_params.txt")
})

# -----------------------------------------------------------------------------
# Data layer: .prepare_balance_data() computes without writing to disk
# -----------------------------------------------------------------------------

test_that(".prepare_balance_data returns report data and writes no files", {
  skip_on_cran()
  skip_if_not_installed("writexl")
  skip_if_not_installed("withr")

  td <- withr::local_tempdir()
  input <- file.path(td, "input.xlsx")
  writexl::write_xlsx(.golden_rus01(), input)

  fp_loaded <- suppressWarnings(suppressMessages(
    .harmonize_plot_input(input, "field_sheet", verbose = FALSE)
  ))
  geom <- .resolve_plot_geometry(plot_size = 1, subplot_size = 10,
                                 plot_width_m = NULL, plot_length_m = NULL,
                                 auto_message = FALSE)

  wd_before <- list.files(getwd(), recursive = TRUE)
  td_before <- list.files(td, recursive = TRUE)

  out <- suppressWarnings(suppressMessages(.prepare_balance_data(
    fp_loaded = fp_loaded,
    input_type = "field_sheet",
    language = "en",
    plot_size = 1,
    subplot_size = 10,
    plot_width_m = geom$plot_width_m,
    plot_length_m = geom$plot_length_m,
    plot_geometry = geom,
    interactive = FALSE,
    verbose = FALSE
  )))

  expect_identical(list.files(getwd(), recursive = TRUE), wd_before)
  expect_identical(list.files(td, recursive = TRUE), td_before)

  expect_named(out, c("params", "fp_coords", "spec_df", "subplot_plots", "plots",
                      "dashboard", "tf_col", "tf_uncol", "tf_palm", "metadata"))
  expect_identical(out$metadata$plot_code, "RUS01")
  expect_identical(out$metadata$plot_name, "REBIO Augusto Ruschi")

  s <- out$params$stats
  expect_equal(s$total, nrow(out$fp_coords))
  expect_equal(s$collected + s$uncollected + s$palms, s$total)
  expect_length(out$subplot_plots, 100L)
  expect_s3_class(out$plots$base, "ggplot")
  expect_null(out$params$interactive_main_plot)
})

test_that(".resolve_balance_metadata prefers user values, then file, then defaults", {
  from_file <- list(plot_name = " File plot ", plot_code = "F1",
                    census_no_fp = NA_character_, team = "")

  m <- .resolve_balance_metadata(from_file)
  expect_identical(m, list(plot_name = "File plot", plot_code = "F1",
                           plot_census_no_fp = "1", team = ""))

  m <- .resolve_balance_metadata(from_file, plot_name = "User", plot_code = "  ",
                                 plot_census_no = "3", team = "Team A")
  expect_identical(m$plot_name, "User")
  expect_identical(m$plot_code, "F1")
  expect_identical(m$plot_census_no_fp, "3")
  expect_identical(m$team, "Team A")

  m <- .resolve_balance_metadata(list())
  expect_identical(m$plot_name, "Unknown Plot")
})
