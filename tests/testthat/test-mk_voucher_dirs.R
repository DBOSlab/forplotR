# tests/testthat/test-mk_voucher_dirs.R

test_that("mk_voucher_dirs errors on missing fp_file_path", {
  expect_error(
    mk_voucher_dirs(fp_file_path = tempfile(fileext = ".xlsx")),
    "Please provide a valid Excel file path\\."
  )
})

.make_fp_xlsx_for_mk_voucher_dirs <- function(path) {
  skip_if_not_installed("openxlsx")

  meta_row <- c(
    "Plotcode: TEST",
    "Plot Name: Demo",
    "Date: 05/12/2025",
    "Team: Alice, Bob"
  )

  header_row <- c(
    "Family",
    "Voucher",
    "Collected",
    "Original determination"
  )

  data_row <- c(
    "Fabaceae",
    "ABC1234",
    "yes",
    "Inga sp."
  )

  dat <- as.data.frame(
    rbind(meta_row, header_row, data_row),
    stringsAsFactors = FALSE
  )

  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "Sheet1")
  openxlsx::writeData(wb, "Sheet1", dat, colNames = FALSE, rowNames = FALSE)
  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
}

test_that("mk_voucher_dirs creates correct Family/Genus/Voucher dirs for collected vouchers", {
  skip_if_not_installed("openxlsx")
  skip_if_not_installed("readxl")

  fp_path <- tempfile(fileext = ".xlsx")
  out_dir <- tempfile(pattern = "voucher_imgs_")
  dir.create(out_dir)

  .make_fp_xlsx_for_mk_voucher_dirs(fp_path)

  mk_voucher_dirs(
    fp_file_path = fp_path,
    input_type = "field_sheet",
    output_dir = out_dir
  )

  expect_true(dir.exists(file.path(out_dir, "Fabaceae", "Inga", "ABC1234")))
})

test_that("mk_voucher_dirs moves photos from outdated genus folder and deletes empty old folder", {
  skip_if_not_installed("openxlsx")
  skip_if_not_installed("readxl")

  fp_path <- tempfile(fileext = ".xlsx")
  out_dir <- tempfile(pattern = "voucher_imgs_")
  dir.create(out_dir)

  .make_fp_xlsx_for_mk_voucher_dirs(fp_path)

  old_dir <- file.path(out_dir, "Fabaceae", "Oldgenus", "ABC1234")
  dir.create(old_dir, recursive = TRUE)
  writeLines("x", file.path(old_dir, "img1.jpg"))
  writeLines("y", file.path(old_dir, "img2.jpg"))

  mk_voucher_dirs(
    fp_file_path = fp_path,
    input_type = "field_sheet",
    output_dir = out_dir
  )

  correct_dir <- file.path(out_dir, "Fabaceae", "Inga", "ABC1234")

  expect_true(dir.exists(correct_dir))
  expect_true(file.exists(file.path(correct_dir, "img1.jpg")))
  expect_true(file.exists(file.path(correct_dir, "img2.jpg")))
  expect_false(dir.exists(old_dir))
})

test_that("mk_voucher_dirs preserves existing destination files and moves non-duplicate files", {
  skip_if_not_installed("openxlsx")
  skip_if_not_installed("readxl")

  fp_path <- tempfile(fileext = ".xlsx")
  out_dir <- tempfile(pattern = "voucher_imgs_")
  dir.create(out_dir)

  .make_fp_xlsx_for_mk_voucher_dirs(fp_path)

  correct_dir <- file.path(out_dir, "Fabaceae", "Inga", "ABC1234")
  dir.create(correct_dir, recursive = TRUE)
  writeLines("correct", file.path(correct_dir, "img_correct.jpg"))

  old_dir <- file.path(out_dir, "Fabaceae", "Oldgenus", "ABC1234")
  dir.create(old_dir, recursive = TRUE)
  writeLines("old", file.path(old_dir, "img_old.jpg"))

  mk_voucher_dirs(
    fp_file_path = fp_path,
    input_type = "field_sheet",
    output_dir = out_dir
  )

  expect_true(file.exists(file.path(correct_dir, "img_correct.jpg")))
  expect_true(file.exists(file.path(correct_dir, "img_old.jpg")))
  expect_false(dir.exists(old_dir))
  expect_false(file.exists(file.path(old_dir, "img_old.jpg")))
})
