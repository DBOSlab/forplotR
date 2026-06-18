testthat::skip_if_not_installed("DBI")
testthat::skip_if_not_installed("duckdb")

make_fake_excel <- function() {
  p <- tempfile(fileext = ".xlsx")
  file.create(p)
  p
}

write_occurrence_txt <- function(dir, x) {
  subdir <- file.path(dir, "dwca_rb")
  dir.create(subdir, recursive = TRUE, showWarnings = FALSE)
  out <- file.path(subdir, "occurrence.txt")
  utils::write.table(
    x = x,
    file = out,
    sep = "\t",
    row.names = FALSE,
    col.names = TRUE,
    quote = FALSE,
    na = ""
  )
  out
}

fake_occurrence_df <- function() {
  data.frame(
    collectionCode = c("RB", "RB", "RB", "RB", "RB"),
    catalogNumber = c("RB100", "RB101", "RB102", "RB103", "RB104"),
    occurrenceID = c("occ100", "occ101", "occ102", "occ103", "occ104"),
    recordedBy = c("H. Medeiros", "H. Medeiros", "H. Medeiros", "H. Medeiros", "H. Medeiros"),
    recordNumber = c("100", "101", "102", "103", "104"),
    family = c("Olacaceae", "Fabaceae", "Apocynaceae", "Lauraceae", "Lauraceae"),
    genus = c("Heisteria", "Arapatiella", "Aspidosperma", "", "Ocotea"),
    specificEpithet = c("", "psilophylla", "Aspidosperma", "", "pulchella"),
    infraspecificEpithet = c("", "", "", "", ""),
    taxonRank = c("species", "species", "species", "family", "species"),
    scientificName = c(
      "Heisteria",
      "Arapatiella psilophylla (Harms) R.S.Cowan",
      "Aspidosperma aspidosperma",
      "Lauraceae",
      "Ocotea pulchella"
    ),
    acceptedScientificName = c(
      "Heisteria",
      "Arapatiella psilophylla (Harms) R.S.Cowan",
      "Aspidosperma aspidosperma",
      "Lauraceae",
      "Ocotea pulchella"
    ),
    verbatimScientificName = c(
      "Heisteria",
      "Arapatiella psilophylla (Harms) R.S.Cowan",
      "Aspidosperma aspidosperma",
      "Lauraceae",
      "Ocotea pulchella"
    ),
    identifiedBy = c("Det A", "Det B", "Det C", "Det D", "Det E"),
    dateIdentified = c("2020-01-01", "2020-01-02", "2020-01-03", "2020-01-04", "2020-01-05"),
    stringsAsFactors = FALSE
  )
}

fake_inventory_df <- function() {
  data.frame(
    Voucher = c("HM100", "HM101", "HM102", "HM103", "HM104", "HM999"),
    Family = c("Olacaceae", "Fabaceae", "Apocynaceae", "Lauraceae", "Lauraceae", "Myrtaceae"),
    `Original determination` = c(
      "Heisteria indet",
      "Arapatiella psilophylla",
      "Aspidosperma indet",
      "Lauraceae indet",
      "Ocotea indet",
      "Eugenia indet"
    ),
    Collected = c("Sim", "Sim", "Sim", "Sim", "Sim", "Sim"),
    stringsAsFactors = FALSE,
    check.names = FALSE
  )
}

fake_identity_parser <- function(voucher, collector_fallback = NULL, collector_codes = NULL) {
  nums <- gsub("\\D+", "", as.character(voucher))
  nums[nums == ""] <- NA_character_

  data.frame(
    collector_raw = rep("H. Medeiros", length(voucher)),
    number_raw = nums,
    number_clean = nums,
    primary_key = ifelse(is.na(nums), NA_character_, paste0("hmedeiros||", nums)),
    fallback_key = ifelse(is.na(nums), NA_character_, paste0("medeiros||", nums)),
    stringsAsFactors = FALSE
  )
}

fake_collector_tokens_one <- function(name_raw) {
  c(primary = "hmedeiros", fallback = "medeiros")
}

testthat::test_that("forplot_crossdet errors on missing input file", {
  testthat::local_mocked_bindings(
    .arg_check_herbarium = function(x) invisible(TRUE),
    .package = "forplotR"
  )

  testthat::expect_error(
    forplot_crossdet(
      fp_file_path = "does-not-exist.xlsx",
      input_type = "field_sheet",
      herbaria = "RB",
      save = FALSE
    ),
    "must be a valid existing file"
  )
})

testthat::test_that("forplot_crossdet errors when standardized inventory lacks required columns", {
  fp <- make_fake_excel()

  testthat::local_mocked_bindings(
    .arg_check_herbarium = function(x) invisible(TRUE),
    .fp_query_to_field_sheet_df = function(path, sheet = NULL) {
      data.frame(Voucher = "HM100", Collected = "Sim", stringsAsFactors = FALSE)
    },
    .package = "forplotR"
  )

  testthat::expect_error(
    forplot_crossdet(
      fp_file_path = fp,
      input_type = "field_sheet",
      herbaria = "RB",
      save = FALSE
    ),
    "missing required columns"
  )
})

testthat::test_that("forplot_crossdet normalizes taxon names and omits same_determination from matches", {
  fp <- make_fake_excel()
  occ_dir <- tempfile("occdir_")
  dir.create(occ_dir, recursive = TRUE, showWarnings = FALSE)
  write_occurrence_txt(occ_dir, fake_occurrence_df())

  testthat::local_mocked_bindings(
    .arg_check_herbarium = function(x) invisible(TRUE),
    .fp_query_to_field_sheet_df = function(path, sheet = NULL) fake_inventory_df(),
    .parse_census_identity = fake_identity_parser,
    .split_recordedby_people = function(recordedBy) "H. Medeiros",
    .collector_tokens_one = fake_collector_tokens_one,
    .package = "forplotR"
  )

  out <- forplot_crossdet(
    fp_file_path = fp,
    input_type = "field_sheet",
    herbaria = "RB",
    path = occ_dir,
    save = FALSE
  )

  testthat::expect_true(is.list(out))
  testthat::expect_true(all(c("matches", "diagnostics", "unmatched_inventory") %in% names(out)))

  testthat::expect_equal(nrow(out$matches), 1L)
  testthat::expect_equal(out$matches$voucher, "HM104")
  testthat::expect_equal(out$matches$inventory_species, "Ocotea indet")
  testthat::expect_equal(out$matches$herbarium_species, "Ocotea pulchella")
  testthat::expect_equal(out$matches$diagnostic_class, "inventory_indet_herbarium_det")

  diag_n <- stats::setNames(out$diagnostics$n, out$diagnostics$diagnostic_class)
  testthat::expect_equal(unname(diag_n["same_determination"]), 4)
  testthat::expect_equal(unname(diag_n["inventory_indet_herbarium_det"]), 1)

  testthat::expect_true("HM999" %in% out$unmatched_inventory$voucher)
})

testthat::test_that("forplot_crossdet uses fallback scientific name normalization when genus is missing", {
  fp <- make_fake_excel()
  occ_dir <- tempfile("occdir_")
  dir.create(occ_dir, recursive = TRUE, showWarnings = FALSE)

  occ <- fake_occurrence_df()
  occ$genus[4] <- ""
  occ$specificEpithet[4] <- ""
  occ$scientificName[4] <- "Lauraceae"
  occ$acceptedScientificName[4] <- "Lauraceae"
  occ$verbatimScientificName[4] <- "Lauraceae"

  write_occurrence_txt(occ_dir, occ)

  inv <- fake_inventory_df()[4, , drop = FALSE]

  testthat::local_mocked_bindings(
    .arg_check_herbarium = function(x) invisible(TRUE),
    .fp_query_to_field_sheet_df = function(path, sheet = NULL) inv,
    .parse_census_identity = fake_identity_parser,
    .split_recordedby_people = function(recordedBy) "H. Medeiros",
    .collector_tokens_one = fake_collector_tokens_one,
    .package = "forplotR"
  )

  out <- forplot_crossdet(
    fp_file_path = fp,
    input_type = "field_sheet",
    herbaria = "RB",
    path = occ_dir,
    save = FALSE
  )

  testthat::expect_equal(nrow(out$matches), 0L)
  diag_n <- stats::setNames(out$diagnostics$n, out$diagnostics$diagnostic_class)
  testthat::expect_equal(unname(diag_n["same_determination"]), 1)
})

testthat::test_that("forplot_crossdet can use auto-download branch via mocked IPT helpers", {
  fp <- make_fake_excel()
  dl_root <- tempfile("jabot_dl_")
  dir.create(dl_root, recursive = TRUE, showWarnings = FALSE)
  dwca_dir <- file.path(dl_root, "dwca_jabot_RB_jbrj_rb")
  dir.create(dwca_dir, recursive = TRUE, showWarnings = FALSE)
  write_occurrence_txt(dwca_dir, fake_occurrence_df()[1, , drop = FALSE])

  cache_dir <- tempfile("cache_")
  dir.create(cache_dir, recursive = TRUE, showWarnings = FALSE)

  testthat::local_mocked_bindings(
    .arg_check_herbarium = function(x) invisible(TRUE),
    .fp_query_to_field_sheet_df = function(path, sheet = NULL) fake_inventory_df()[1, , drop = FALSE],
    .parse_census_identity = fake_identity_parser,
    .split_recordedby_people = function(recordedBy) "H. Medeiros",
    .collector_tokens_one = fake_collector_tokens_one,
    .herbaria_cache_paths = function() {
      list(
        cache_dir = cache_dir,
        duckdb_path = file.path(cache_dir, "herbaria_index.duckdb")
      )
    },
    .get_ipt_info = function(herbarium, ipt = c("jabot", "reflora"), resource_map = NULL) {
      data.frame(
        ipt = "jabot",
        herbarium = "RB",
        resource_id = "jbrj_rb",
        archive_base = "https://example.org/archive.do?r=",
        resource_url = "https://example.org/resource?r=jbrj_rb",
        stringsAsFactors = FALSE
      )
    },
    .download_dwca_one = function(info_row, dir, verbose = FALSE, force_refresh = FALSE) {
      src <- file.path(dwca_dir, "occurrence.txt")
      out_dir <- file.path(dir, "dwca_jabot_RB_jbrj_rb")
      dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)
      file.copy(src, file.path(out_dir, "occurrence.txt"), overwrite = TRUE)
      out_dir
    },
    .package = "forplotR"
  )

  out <- forplot_crossdet(
    fp_file_path = fp,
    input_type = "field_sheet",
    herbaria = "RB",
    path = NULL,
    save = FALSE,
    verbose = FALSE
  )

  testthat::expect_equal(nrow(out$matches), 0L)
  diag_n <- stats::setNames(out$diagnostics$n, out$diagnostics$diagnostic_class)
  testthat::expect_equal(unname(diag_n["same_determination"]), 1)
})

testthat::test_that("forplot_crossdet writes workbook when save = TRUE", {
  testthat::skip_if_not_installed("writexl")

  fp <- make_fake_excel()
  occ_dir <- tempfile("occdir_")
  dir.create(occ_dir, recursive = TRUE, showWarnings = FALSE)
  write_occurrence_txt(occ_dir, fake_occurrence_df()[5, , drop = FALSE])

  out_dir <- tempfile("crossdet_out_")
  dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)

  inv <- fake_inventory_df()[5, , drop = FALSE]

  testthat::local_mocked_bindings(
    .arg_check_herbarium = function(x) invisible(TRUE),
    .fp_query_to_field_sheet_df = function(path, sheet = NULL) inv,
    .parse_census_identity = fake_identity_parser,
    .split_recordedby_people = function(recordedBy) "H. Medeiros",
    .collector_tokens_one = fake_collector_tokens_one,
    .package = "forplotR"
  )

  invisible(forplot_crossdet(
    fp_file_path = fp,
    input_type = "field_sheet",
    herbaria = "RB",
    path = occ_dir,
    save = TRUE,
    dir = out_dir,
    filename = "crossdet_test"
  ))

  testthat::expect_true(file.exists(file.path(out_dir, "crossdet_test.xlsx")))
})
