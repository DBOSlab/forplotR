# Additional tests for input harmonization, FP query conversion and the
# report dashboard helpers in R/aux_fxns_plots.R.

.ap_write <- function(sheets, env = parent.frame()) {
  path <- withr::local_tempfile(fileext = ".xlsx", .local_envir = env)
  writexl::write_xlsx(sheets, path, col_names = TRUE)
  path
}

.ap_write_raw <- function(rows, env = parent.frame()) {
  # rows: list of character vectors, written without a header row
  width <- max(lengths(rows))
  m <- t(vapply(rows, function(r) c(r, rep(NA_character_, width - length(r))),
                character(width)))
  df <- as.data.frame(m, stringsAsFactors = FALSE)
  path <- withr::local_tempfile(fileext = ".xlsx", .local_envir = env)
  writexl::write_xlsx(df, path, col_names = FALSE)
  path
}


# .harmonize_plot_input --------------------------------------------------------

test_that("field_sheet_ti input is mapped by column position", {
  skip_if_not_installed("writexl")
  hdr <- c("Tag", "Group", "T1", "X", "Y", "Family", "Species", "Genus",
           "Morpho", "D", "POM", "Flag", "Height", "Voucher", "Silica",
           "Collected", "Notes", "c18", "c19", "CAP")
  rows <- list(
    c("Plotcode: TI01", "Plot Name: Ilha", "Team: Ana, Bia"),
    hdr,
    c("1", "", "1", "2", "3", "Fabaceae", "alba", "Inga", "m1", "120", "1.3",
      "a", "12", "AB 1", "yes", "yes", "nota", "", "", "40"),
    c("2", "", "2", "5", "6", "Myrtaceae", "sp.", "Eugenia", "", "x", "1.3",
      "", "", "", "", "", "", "", "", "")
  )
  path <- .ap_write_raw(rows)

  msgs <- testthat::capture_messages(
    out <- .harmonize_plot_input(path, " Field_Sheet_TI ", verbose = TRUE)
  )
  fp <- out$fp_sheet
  expect_identical(out$plot_code, "TI01")
  expect_identical(out$plot_name, "Ilha")
  expect_identical(nrow(fp), 2L)
  expect_identical(fp$`Original determination`, c("Inga | alba", "Eugenia | sp."))
  expect_identical(fp$T1, c(1, 2))
  expect_identical(fp$D, c(120, NA))
  expect_identical(fp$Voucher, c("AB 1", NA))
  expect_identical(fp$CAP, c(40, NA))
  expect_true(any(grepl("MULTISTEM CONSOLIDATION", msgs)))

  # Mapped columns override the fixed TI positions
  cm <- .build_col_map(colname_species = "Species", colname_tagno = "Voucher")
  out2 <- .harmonize_plot_input(path, "field_sheet_ti", verbose = FALSE,
                                col_map = cm)
  expect_identical(out2$fp_sheet$`Original determination`, c("alba", "sp."))

  short <- .ap_write_raw(list(c("Plotcode: X"), hdr))
  expect_error(.harmonize_plot_input(short, "field_sheet_ti"),
               "must contain metadata, header, and data rows")
})

test_that("field_sheet input requires metadata unless the user supplies it", {
  skip_if_not_installed("writexl")
  rows <- list(
    c("New Tag No", "T1", "X", "Y", "D", "Family", "Original determination",
      "Voucher"),
    c("1", "1", "2", "3", "100", "Fabaceae", "Inga alba", "V1")
  )
  path <- .ap_write_raw(rows)

  expect_error(.harmonize_plot_input(path, "field_sheet"),
               "has no metadata row")

  msgs <- testthat::capture_messages(
    out <- .harmonize_plot_input(path, "field_sheet", has_user_meta = TRUE,
                                 verbose = TRUE)
  )
  expect_true(all(is.na(out$fp_sheet$Collected)))
  expect_true(any(grepl("Multistem consolidation skipped", msgs)))
  expect_identical(out$plot_code, "")
})

test_that("missing required columns are reported with the colname_ hint", {
  skip_if_not_installed("writexl")
  path <- .ap_write_raw(list(
    c("Plotcode: A1"),
    c("New Tag No", "T1", "X", "Y", "Family", "Original determination", "Voucher"),
    c("1", "1", "2", "3", "Fabaceae", "Inga alba", "V1")
  ))
  expect_error(.harmonize_plot_input(path, "field_sheet", verbose = FALSE),
               "missing required columns: D\\. .*`colname_dbh` \\(D\\)")
  expect_error(.harmonize_plot_input("nope.xlsx", "field_sheet"),
               "'fp_file_path' does not exist")
})


# .fp_query_to_field_sheet_df --------------------------------------------------

.ap_query <- function(...) {
  df <- tibble::tibble(
    `Plot Code` = c("AB12", "AB12"),
    `Tag No` = c("1", "2"),
    `Recommended Family` = c("Fabaceae", "Rubiaceae"),
    `Recommended Species` = c("Inga alba", "Coffea arabica"),
    `Voucher Code` = c("GO 1", "")
  )
  extra <- tibble::tibble(...)
  dplyr::bind_cols(df, extra)
}

test_that("standardised coordinates are used when local ones are absent", {
  skip_if_not_installed("writexl")
  path <- .ap_write(list(Data = .ap_query(
    `Standardised SubPlot T1` = c("1", "2"),
    `Standardised X` = c("1.5", "2"), `Standardised Y` = c("3", "4")
  )))
  out <- .fp_query_to_field_sheet_df(path)
  expect_identical(attr(out, "coord_mode"), "standardised")
  expect_equal(out$X, c(1.5, 2))
  expect_true(all(is.na(out$T2)))
})

test_that("FP query conversion reports unusable coordinates", {
  skip_if_not_installed("writexl")
  none <- .ap_write(list(Data = .ap_query(Other = c("a", "b"))))
  expect_error(.fp_query_to_field_sheet_df(none),
               "Could not detect FP query coordinates")

  na_xy <- .ap_write(list(Data = .ap_query(
    `Sub Plot T1` = c("1", "1"), X = c("a", "b"), Y = c("c", "d")
  )))
  expect_error(.fp_query_to_field_sheet_df(na_xy),
               "produced X/Y entirely as NA")

  same <- .ap_write(list(Data = .ap_query(
    `Sub Plot T1` = c("1", "1"), X = c("2", "3"), Y = c("2", "3")
  )))
  expect_warning(.fp_query_to_field_sheet_df(same),
                 "suspiciously identical")

  expect_error(.fp_query_to_field_sheet_df("nope.xlsx"), "Input file does not exist")
})


# Report dashboard -------------------------------------------------------------

.ap_dash_df <- function() {
  data.frame(
    T1 = c("N", "N", "S"), T2 = c("1", "2", "1"),
    Family = c("Fabaceae", "", "Arecaceae"),
    `Original determination` = c(
      "Inga averyveryverylongspeciesepithetname subsp. extra", NA, "Euterpe edulis"
    ),
    Collected = c("yes", "", NA), D = c(NA, NA, NA),
    check.names = FALSE, stringsAsFactors = FALSE
  )
}

test_that(".prepare_report_dashboard handles MONITORA subplots, TI tables and missing DBH", {
  out <- suppressMessages(.prepare_report_dashboard(.ap_dash_df(), input_type = "monitora",
                                   language = "xx"))
  # Subplots are combined as subunit-subplot (three sampling units)
  expect_setequal(out$phytosoc$species_metrics$n_subplots, 1L)
  expect_true(any(grepl("\\.\\.\\.$", out$species_metrics_tbl[[1]])))
  expect_s3_class(out$dbh_plot, "ggplot")
  expect_length(out$dbh_plot$layers, 0L)

  ti <- suppressMessages(.prepare_report_dashboard(.ap_dash_df(), input_type = "field_sheet_ti"))
  expect_identical(nrow(ti$metrics_tbl), 6L)
  expect_identical(ncol(ti$species_metrics_tbl), 4L)
  expect_identical(ncol(ti$family_metrics_tbl), 4L)
})

test_that(".calculate_phytosociological_metrics warns on empty input", {
  empty <- .ap_dash_df()[0, ]
  expect_warning(res <- .calculate_phytosociological_metrics(empty),
                 "No valid species data")
  expect_identical(res$diversity_metrics$total_species, 0)
})

test_that("number formatters and dictionary fall back sensibly", {
  expect_identical(.fmt_int(NA), NA_character_)
  expect_identical(.fmt_int(numeric(0)), NA_character_)
  expect_identical(.fmt_int(2.6), "3")
  expect_identical(.fmt_dec(NA), NA_character_)
  expect_identical(.fmt_dec(1 / 3, 2), "0.33")
  expect_identical(.tr_dict("collected", "zz", dict), .tr_dict("collected", "en", dict))
})


# .monitora_to_field_sheet_df errors -------------------------------------------

test_that(".monitora_to_field_sheet_df explains unreadable MONITORA sheets", {
  skip_if_not_installed("writexl")
  expect_error(.monitora_to_field_sheet_df("nope.xlsx"),
               "MONITORA file does not exist")

  no_schema <- .ap_write(list(Plan1 = data.frame(a = 1, b = 2)))
  expect_error(.monitora_to_field_sheet_df(no_schema),
               "MONITORA schema could not be detected")

  tum <- forplotR::TUM[1:5, setdiff(names(forplotR::TUM), c("X", "Y"))]
  no_xy <- .ap_write(list(Plan1 = tum))
  expect_error(.monitora_to_field_sheet_df(no_xy),
               "MONITORA coordinate columns could not be detected")
})
