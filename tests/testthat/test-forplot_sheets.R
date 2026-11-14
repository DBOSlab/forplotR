# Tests for forplot_sheets()
#
# All tests run offline: the bundled RUS01 (ForestPlots field sheet) and TUM
# (MONITORA) datasets are written to a temporary directory, and the two
# network calls (geodata::gadm() and geodata::elevation_global()) are either
# avoided (validate_locality = FALSE, alt supplied) or mocked.

.sheets_jabot_cols <- c(
  "family", "genus", "sp1", "number", "collector", "addcoll", "colldd",
  "collmm", "collyy", "country", "majorarea", "minorarea", "uc", "projeto",
  "locnotes", "latitude", "longitude", "altprof", "notes"
)

.sheets_fixtures <- function(env = parent.frame()) {
  td <- withr::local_tempdir(.local_envir = env)

  fp <- file.path(td, "RUS01.xlsx")
  openxlsx::write.xlsx(forplotR::RUS01, fp)

  mon <- file.path(td, "TUM.xlsx")
  openxlsx::write.xlsx(forplotR::TUM, mon)

  herb <- file.path(td, "jabot_template.xlsx")
  template <- as.data.frame(
    stats::setNames(
      replicate(length(.sheets_jabot_cols), character(0), simplify = FALSE),
      .sheets_jabot_cols
    )
  )
  openxlsx::write.xlsx(template, herb)

  list(dir = td, fp = fp, mon = mon, herb = herb)
}

.run_sheets <- function(fx, ...) {
  args <- utils::modifyList(
    list(
      fp_file_path = fx$fp,
      herb_file_path = fx$herb,
      input_type = "forestplots",
      majorarea = "ES",
      minorarea = "Santa Teresa",
      collector = "G.C. Ottino",
      lat = -19.9,
      long = -40.55,
      alt = 800,
      validate_locality = FALSE,
      dir = fx$dir
    ),
    list(...)
  )
  suppressMessages(do.call(forplot_sheets, args))
}


# Input validation ------------------------------------------------------------

test_that("forplot_sheets rejects missing or invalid file paths", {
  expect_error(forplot_sheets(), "valid path in 'fp_file_path'")
  expect_error(
    forplot_sheets(fp_file_path = c("a.xlsx", "b.xlsx")),
    "valid path in 'fp_file_path'"
  )
  expect_error(
    forplot_sheets(fp_file_path = "does-not-exist.xlsx"),
    "'fp_file_path' does not exist"
  )

  fx <- .sheets_fixtures()
  expect_error(
    forplot_sheets(fp_file_path = fx$fp),
    "valid path in 'herb_file_path'"
  )
  expect_error(
    forplot_sheets(fp_file_path = fx$fp, herb_file_path = "nope.xlsx"),
    "'herb_file_path' does not exist"
  )
})

test_that("forplot_sheets requires column_map for a custom herbarium format", {
  fx <- .sheets_fixtures()
  expect_error(
    forplot_sheets(
      fp_file_path = fx$fp,
      herb_file_path = fx$herb,
      herbarium_format = "custom"
    ),
    "must provide 'column_map'"
  )
})

test_that("forplot_sheets rejects an unknown input_type", {
  fx <- .sheets_fixtures()
  expect_error(
    forplot_sheets(
      fp_file_path = fx$fp,
      herb_file_path = fx$herb,
      input_type = "not_a_layout"
    ),
    "should be one of"
  )
})


# ForestPlots field sheet -----------------------------------------------------

test_that("forplot_sheets fills a JABOT sheet from a ForestPlots field sheet", {
  fx <- .sheets_fixtures()
  out <- .run_sheets(fx)

  expect_s3_class(out, "data.frame")
  expect_gt(nrow(out), 0)
  expect_true(all(c("family", "genus", "number", "notes") %in% names(out)))

  # Location arguments are recycled to every row
  expect_true(all(out$collector == "G.C. Ottino"))
  expect_true(all(out$latitude == -19.9))
  expect_true(all(out$altprof == 800))

  # Indeterminate names never leak into family/genus
  expect_false(any(grepl("indet", out$family, ignore.case = TRUE), na.rm = TRUE))
  expect_false(any(grepl("indet", out$genus, ignore.case = TRUE), na.rm = TRUE))

  # Voucher numbers keep digits only
  expect_true(all(grepl("^[0-9]*$", out$number[!is.na(out$number)])))

  # Notes are written in the requested language
  expect_true(any(grepl("DBH", out$notes, fixed = TRUE)))
})

test_that("forplot_sheets writes the spreadsheet inside a date-stamped folder of dir", {
  fx <- .sheets_fixtures()
  out <- .run_sheets(fx, filename = "rus01_labels")

  written <- list.files(fx$dir, pattern = "^rus01_labels\\.xlsx$",
                        recursive = TRUE, full.names = TRUE)
  expect_length(written, 1L)
  expect_identical(basename(dirname(written)), format(Sys.time(), "%d%b%Y"))

  back <- openxlsx::read.xlsx(written)
  expect_equal(nrow(back), nrow(out))
})

test_that("forplot_sheets writes notes in Portuguese when language = 'pt'", {
  fx <- .sheets_fixtures()
  out <- .run_sheets(fx, language = "pt")
  expect_true(any(grepl("DAP", out$notes, fixed = TRUE)))
})


# MONITORA --------------------------------------------------------------------

test_that("forplot_sheets processes a single MONITORA station", {
  fx <- .sheets_fixtures()
  out <- .run_sheets(
    fx,
    fp_file_path = fx$mon,
    input_type = "monitora",
    station_name = "EA-1",
    majorarea = "AP",
    minorarea = "Laranjal do Jari",
    lat = 1.5,
    long = -52.5,
    alt = 300
  )

  expect_s3_class(out, "data.frame")
  expect_gt(nrow(out), 0)
  expect_true(all(out$majorarea == "AP"))
  expect_true(all(out$altprof == 300))
})


# Offline safety --------------------------------------------------------------

test_that("forplot_sheets warns and continues when GADM boundaries cannot be downloaded", {
  fx <- .sheets_fixtures()

  testthat::local_mocked_bindings(
    gadm = function(...) stop("simulated network failure"),
    .package = "geodata"
  )

  expect_warning(
    out <- .run_sheets(fx, validate_locality = TRUE),
    "locality validation skipped"
  )
  expect_s3_class(out, "data.frame")
  expect_gt(nrow(out), 0)
})

test_that("forplot_sheets warns when geodata returns NULL for GADM (server down)", {
  fx <- .sheets_fixtures()

  testthat::local_mocked_bindings(
    gadm = function(...) NULL,
    .package = "geodata"
  )

  expect_warning(
    out <- .run_sheets(fx, validate_locality = TRUE),
    "locality validation skipped"
  )
  expect_s3_class(out, "data.frame")
})

test_that("forplot_sheets leaves altitude NA when elevation data cannot be downloaded", {
  fx <- .sheets_fixtures()

  testthat::local_mocked_bindings(
    elevation_global = function(...) stop("simulated network failure"),
    .package = "geodata"
  )

  expect_warning(
    out <- .run_sheets(fx, alt = NULL),
    "Could not download global elevation data"
  )
  expect_s3_class(out, "data.frame")
  expect_true(all(is.na(out$altprof)))
})

test_that("forplot_sheets does not write to the working directory", {
  fx <- .sheets_fixtures()
  wd <- withr::local_tempdir()
  withr::local_dir(wd)

  .run_sheets(fx)

  expect_length(list.files(wd, recursive = TRUE, all.files = TRUE, no.. = TRUE), 0L)
})

test_that(".check_point_in_admin validates coordinates against mocked boundaries", {
  # A square "state" polygon around (-40.5, -19.9); no download involved.
  square <- terra::vect(
    "POLYGON ((-41 -20.5, -40 -20.5, -40 -19.5, -41 -19.5, -41 -20.5))",
    crs = "EPSG:4326"
  )
  square$NAME_1 <- "Espírito Santo"

  testthat::local_mocked_bindings(
    gadm = function(...) square,
    .package = "geodata"
  )

  # Inside, matched by full (accented) name and by state abbreviation
  expect_no_error(.check_point_in_admin(-19.9, -40.55, "Espírito Santo", "Brazil"))
  expect_no_error(.check_point_in_admin(-19.9, -40.55, "ES", "Brazil"))

  # Outside the polygon
  expect_error(
    .check_point_in_admin(-10, -50, "ES", "Brazil"),
    "DO NOT fall within"
  )

  # Unknown administrative region
  expect_error(
    .check_point_in_admin(-19.9, -40.55, "Atlantis", "Brazil"),
    "Could not match"
  )
})


# Regression tests ------------------------------------------------------------

test_that(".fmt_num removes floating-point noise and keeps text", {
  expect_identical(.fmt_num("9.1999999999999993"), "9.2")
  expect_identical(.fmt_num("267.39999999999998"), "267.4")
  expect_identical(.fmt_num(c("3", "4.4000000000000004", NA, "", "abc")),
                   c("3", "4.4", "", "", "abc"))
  expect_identical(.fmt_num(13.369), "13.37")
})

test_that("forplot_sheets notes carry no floating-point noise", {
  fx <- .sheets_fixtures()
  out <- .run_sheets(fx)
  expect_false(any(grepl("[0-9][.][0-9]{5,}", out$notes)))
})

test_that("MONITORA vouchers recorded in an earlier census are carried over", {
  fx <- .sheets_fixtures()
  out <- .run_sheets(
    fx,
    fp_file_path = fx$mon,
    input_type = "monitora",
    station_name = "EA-1",
    majorarea = "AP",
    minorarea = "Laranjal do Jari",
    lat = 1.5,
    long = -52.5,
    alt = 300
  )

  # Most collected trees of EA-1 have their voucher in an earlier census
  expect_gt(mean(!is.na(out$number) & nzchar(out$number)), 0.5)

  # Tree 27 (subunit Leste, subplot 6) was vouchered as H. Medeiros 3239
  tum <- forplotR::TUM
  num_col <- grep("^voucher/n", names(tum), value = TRUE)
  expect_true("3239" %in% tum[[num_col]])
  expect_true("3239" %in% out$number)
})
