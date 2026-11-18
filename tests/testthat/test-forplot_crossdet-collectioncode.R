# forplot_crossdet(): herbarium of each voucher from the `collectionCode`
# column, and download of only those herbaria (jabotR::jabot_download()).

testthat::skip_if_not_installed("DBI")
testthat::skip_if_not_installed("duckdb")

.cc_identity <- function(voucher, collector_fallback = NULL, collector_codes = NULL) {
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

.cc_occ <- function(code, n, genus, epithet) {
  data.frame(
    collectionCode = code, catalogNumber = paste0(code, n), occurrenceID = paste0(code, "-", n),
    recordedBy = "H. Medeiros", recordNumber = n, family = "Lauraceae",
    genus = genus, specificEpithet = epithet, infraspecificEpithet = "", taxonRank = "species",
    scientificName = paste(genus, epithet), acceptedScientificName = paste(genus, epithet),
    verbatimScientificName = paste(genus, epithet), identifiedBy = "Det", dateIdentified = "2020",
    stringsAsFactors = FALSE
  )
}

# One jabot_download()-like folder per herbarium: occurrence.txt + <CODE>_Jabot.csv
.cc_write_herbarium <- function(dir, code, occ) {
  sub <- file.path(dir, paste0("dwca_jabot_", tolower(code), "_v1_0"))
  dir.create(sub, recursive = TRUE, showWarnings = FALSE)
  utils::write.table(occ, file.path(sub, "occurrence.txt"), sep = "\t",
                     row.names = FALSE, quote = FALSE, na = "")
  utils::write.csv(data.frame(collectionCode = code), file.path(sub, paste0(code, "_Jabot.csv")),
                   row.names = FALSE)
  sub
}

.cc_inventory <- function() {
  data.frame(
    Voucher = c("HM100", "HM101", "HM102"),
    Family = "Lauraceae",
    `Original determination` = c("Ocotea indet", "Ocotea indet", "Ocotea indet"),
    Collected = "Yes",
    collectionCode = c("RB", "inpa", NA),
    stringsAsFactors = FALSE, check.names = FALSE
  )
}

.cc_mocks <- function(env = parent.frame(), inventory = .cc_inventory()) {
  testthat::local_mocked_bindings(
    .fp_query_to_field_sheet_df = function(path, sheet = NULL) inventory,
    .parse_census_identity = .cc_identity,
    .split_recordedby_people = function(recordedBy) "H. Medeiros",
    .collector_tokens_one = function(name_raw) c(primary = "hmedeiros", fallback = "medeiros"),
    .package = "forplotR",
    .env = env
  )
}

testthat::test_that("collectionCode helpers read codes and find jabot_download() folders", {
  inv <- data.frame(a = 1:3, `Collection Code` = c(" rb", "", NA), check.names = FALSE)
  testthat::expect_identical(.inventory_collection_codes(inv), c("RB", NA, NA))
  testthat::expect_identical(.inventory_collection_codes(data.frame(a = 1)), NA_character_)

  td <- withr::local_tempdir()
  .cc_write_herbarium(td, "RB", .cc_occ("RB", "100", "Ocotea", "pulchella"))
  got <- .jabot_occurrence_files(c("RB", "INPA"), td)
  testthat::expect_length(got$RB, 1L)
  testthat::expect_length(got$INPA, 0L)

  withr::local_options(forplotR.offline = TRUE)
  testthat::expect_identical(.jabot_download_codes("RB", td), list())
})

testthat::test_that("each voucher is checked only in the herbarium of its collectionCode", {
  fp <- withr::local_tempfile(fileext = ".xlsx"); file.create(fp)
  td <- withr::local_tempdir()
  # HM100 is in RB (and a duplicate with another name is in INPA);
  # HM101 is only in INPA
  .cc_write_herbarium(td, "RB", .cc_occ("RB", c("100", "101"), "Ocotea", "pulchella"))
  .cc_write_herbarium(td, "INPA", .cc_occ("INPA", c("100", "101"), "Ocotea", "glomerata"))
  .cc_mocks()

  testthat::expect_message(
    out <- forplot_crossdet(fp, input_type = "field_sheet", path = td),
    "1 voucher\\(s\\) without `collectionCode` were not checked"
  )
  m <- out$matches[order(out$matches$voucher), ]
  testthat::expect_identical(m$voucher, c("HM100", "HM101"))
  testthat::expect_identical(m$herbarium, c("RB", "INPA"))
  testthat::expect_identical(m$herbarium_species, c("Ocotea pulchella", "Ocotea glomerata"))
  testthat::expect_false("HM102" %in% out$unmatched_inventory$voucher)
})

testthat::test_that("an explicit `herbaria` checks every voucher there, whatever collectionCode says", {
  fp <- withr::local_tempfile(fileext = ".xlsx"); file.create(fp)
  td <- withr::local_tempdir()
  .cc_write_herbarium(td, "RB", .cc_occ("RB", c("100", "101"), "Ocotea", "pulchella"))
  .cc_write_herbarium(td, "INPA", .cc_occ("INPA", "100", "Ocotea", "glomerata"))
  .cc_mocks()

  out <- forplot_crossdet(fp, input_type = "field_sheet", herbaria = "RB", path = td)
  testthat::expect_setequal(out$matches$voucher, c("HM100", "HM101"))
  testthat::expect_true(all(out$matches$herbarium == "RB"))
  testthat::expect_identical(out$unmatched_inventory$voucher, "HM102")
  testthat::expect_identical(out$unmatched_inventory$herbarium, "RB")
})

testthat::test_that("without herbaria nor collectionCode the user is told what to supply", {
  fp <- withr::local_tempfile(fileext = ".xlsx"); file.create(fp)
  inv <- .cc_inventory(); inv$collectionCode <- NA
  .cc_mocks(inventory = inv)
  testthat::expect_error(forplot_crossdet(fp, input_type = "field_sheet", path = tempdir()),
                         "supply `herbaria`")
})

testthat::test_that("only the herbaria of collectionCode are downloaded, with jabot_download()", {
  fp <- withr::local_tempfile(fileext = ".xlsx"); file.create(fp)
  cache <- withr::local_tempdir()
  asked <- NULL
  .cc_mocks()
  testthat::local_mocked_bindings(
    .herbaria_cache_paths = function() list(cache_dir = cache, duckdb_path = NULL),
    .jabot_download_codes = function(codes, dir, verbose = FALSE) {
      asked <<- codes
      .cc_write_herbarium(dir, "RB", .cc_occ("RB", "100", "Ocotea", "pulchella"))
      .cc_write_herbarium(dir, "INPA", .cc_occ("INPA", "101", "Ocotea", "glomerata"))
      .jabot_occurrence_files(codes, dir)
    },
    .get_ipt_info = function(...) stop("own downloader should not be used"),
    .package = "forplotR"
  )
  out <- suppressMessages(forplot_crossdet(fp, input_type = "field_sheet"))
  testthat::expect_identical(asked, c("INPA", "RB"))
  testthat::expect_setequal(out$matches$voucher, c("HM100", "HM101"))
})

testthat::test_that("herbaria jabot_download() could not fetch fall back on the own downloader", {
  fp <- withr::local_tempfile(fileext = ".xlsx"); file.create(fp)
  cache <- withr::local_tempdir()
  own <- character()
  .cc_mocks()
  testthat::local_mocked_bindings(
    .herbaria_cache_paths = function() list(cache_dir = cache, duckdb_path = NULL),
    .jabot_download_codes = function(codes, dir, verbose = FALSE) {
      .cc_write_herbarium(dir, "RB", .cc_occ("RB", "100", "Ocotea", "pulchella"))
      list(RB = file.path(list.files(dir, full.names = TRUE)[1], "occurrence.txt"), INPA = character())
    },
    .get_ipt_info = function(herbarium, ...) {
      own <<- c(own, herbarium)
      data.frame(herbarium = herbarium, resource_id = "x")
    },
    .download_dwca_one = function(info_row, dir, ...) {
      .cc_write_herbarium(file.path(dir, "own"), "INPA", .cc_occ("INPA", "101", "Ocotea", "glomerata"))
    },
    .package = "forplotR"
  )
  out <- suppressMessages(forplot_crossdet(fp, input_type = "field_sheet"))
  testthat::expect_identical(own, "INPA")
  testthat::expect_setequal(out$matches$herbarium, c("RB", "INPA"))
})

testthat::test_that("GigaFlor sheets are cross-checked through their collectionCode", {
  inv <- .cc_inventory()
  testthat::local_mocked_bindings(
    .harmonize_plot_input = function(fp_file_path, input_type, ...) {
      testthat::expect_identical(input_type, "gigaflor")
      list(fp_sheet = inv)
    },
    .parse_census_identity = .cc_identity,
    .split_recordedby_people = function(recordedBy) "H. Medeiros",
    .collector_tokens_one = function(name_raw) c(primary = "hmedeiros", fallback = "medeiros"),
    .package = "forplotR"
  )
  fp <- withr::local_tempfile(fileext = ".xlsx"); file.create(fp)
  td <- withr::local_tempdir()
  .cc_write_herbarium(td, "RB", .cc_occ("RB", "100", "Ocotea", "pulchella"))
  .cc_write_herbarium(td, "INPA", .cc_occ("INPA", "101", "Ocotea", "glomerata"))
  out <- suppressMessages(forplot_crossdet(fp, input_type = "gigaflor", path = td))
  testthat::expect_setequal(out$matches$voucher, c("HM100", "HM101"))
})
