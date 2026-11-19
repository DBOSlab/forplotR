# Additional forplot_vouchers() tests with a mocked harmonized input.

.vo_sheet <- function() {
  data.frame(
    Family = c("Fabaceae", "Myrtaceae", "Rubiaceae", "", "Lauraceae"),
    `Original determination` = c("Inga alba", "Eugenia sp.", "Psychotria",
                                 "Ocotea", NA),
    Voucher = c("V.1", "V2", "", "V4", "V5"),
    Collected = c("yes", "nao", "yes", "", "sim"),
    `New Tag No` = c("1", "2", "", "4", ""),
    check.names = FALSE, stringsAsFactors = FALSE
  )
}

test_that("forplot_vouchers counts skipped rows and builds Family/Genus/Voucher folders", {
  input <- withr::local_tempfile(fileext = ".xlsx")
  file.create(input)
  out_dir <- file.path(withr::local_tempdir(), "imgs")
  testthat::local_mocked_bindings(
    .harmonize_plot_input = function(...) list(fp_sheet = .vo_sheet())
  )

  msgs <- testthat::capture_messages(
    res <- forplot_vouchers(input, output_dir = paste0(out_dir, "/"))
  )

  expect_true(dir.exists(file.path(out_dir, "Fabaceae", "Inga", "V1_Tree#1")))
  expect_identical(list.dirs(out_dir, recursive = FALSE, full.names = FALSE),
                   "Fabaceae")
  expect_true(any(grepl("Created main output directory", msgs)))
  expect_true(any(grepl("Folders created: 1", msgs)))
  expect_true(any(grepl("Rows skipped \\(not collected\\): 1", msgs)))
  expect_true(any(grepl("Rows skipped \\(missing voucher/family/species\\): 3", msgs)))
  expect_identical(res$Voucher[1], "V.1 Tree#1")
  expect_identical(res$Voucher[4], "V4")
})

test_that("forplot_vouchers reports missing columns and fills a missing tag column", {
  input <- withr::local_tempfile(fileext = ".xlsx")
  file.create(input)
  testthat::local_mocked_bindings(
    .harmonize_plot_input = function(...) list(fp_sheet = .vo_sheet()[, 1:3])
  )
  expect_error(forplot_vouchers(input, verbose = FALSE),
               "missing required columns: Collected")

  testthat::local_mocked_bindings(
    .harmonize_plot_input = function(...) list(fp_sheet = .vo_sheet()[, 1:4])
  )
  res <- forplot_vouchers(input, output_dir = withr::local_tempdir(),
                          verbose = FALSE)
  expect_true(all(is.na(res$`New Tag No`)))
  expect_identical(res$Voucher[1], "V.1")
})

test_that("voucher helpers normalize collection flags and genus", {
  expect_identical(.extract_genus(c("Inga alba", "Fam | Eugenia sp.", "N/A", NA)),
                   c("Inga", "Eugenia", NA, NA))
  expect_identical(.is_not_collected_value(c("Não", "no", "yes")),
                   c(TRUE, TRUE, FALSE))
  expect_identical(.is_collected_value(c("", "Sim"), voucher = c("V1", NA)),
                   c(TRUE, TRUE))
  expect_identical(.clean_path_component(c("a/b:c", NA)), c("a_b_c", ""))
})

test_that("forplot_vouchers names folders in the language of the call", {
  input <- withr::local_tempfile(fileext = ".xlsx")
  file.create(input)
  out_dir <- withr::local_tempdir()
  testthat::local_mocked_bindings(
    .harmonize_plot_input = function(...) list(fp_sheet = .vo_sheet())
  )
  res <- suppressMessages(forplot_vouchers(input, output_dir = out_dir, language = "pt"))
  expect_true(dir.exists(file.path(out_dir, "Fabaceae", "Inga", "V1_Arvore#1")))
  expect_identical(res$Voucher[1], "V.1 Arvore#1")

  # default folder in tempdir(), named in the language
  def <- file.path(tempdir(), "imagens_testemunhos")
  withr::defer(unlink(def, recursive = TRUE))
  suppressMessages(forplot_vouchers(input, language = "pa"))
  expect_true(dir.exists(file.path(def, "Fabaceae", "Inga", "V1_Arvore#1")))

  expect_error(forplot_vouchers(input, output_dir = out_dir, language = "xx"))
})
