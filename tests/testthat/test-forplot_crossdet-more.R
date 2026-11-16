# Additional forplot_crossdet() tests: taxon normalization, empty-result
# branches, input errors and the automatic-download path. All offline: the
# inventory reader and IPT helpers are mocked and occurrence files are local.

testthat::skip_if_not_installed("DBI")
testthat::skip_if_not_installed("duckdb")

.cd_occ <- function(n = 1, ...) {
  base <- data.frame(
    collectionCode = "RB",
    catalogNumber = paste0("RB", seq_len(n)),
    occurrenceID = paste0("occ", seq_len(n)),
    recordedBy = "H. Medeiros",
    recordNumber = as.character(100 + seq_len(n) - 1L),
    family = "Fabaceae",
    genus = "Inga",
    specificEpithet = "edulis",
    infraspecificEpithet = "",
    taxonRank = "species",
    scientificName = "Inga edulis Mart.",
    acceptedScientificName = "Inga edulis Mart.",
    verbatimScientificName = "Inga edulis",
    identifiedBy = "Det",
    dateIdentified = "2020-01-01",
    stringsAsFactors = FALSE
  )
  extra <- list(...)
  for (nm in names(extra)) base[[nm]] <- extra[[nm]]
  base
}

.cd_write <- function(df, sep = "\t", env = parent.frame()) {
  d <- withr::local_tempdir(.local_envir = env)
  sub <- file.path(d, "dwca_rb")
  dir.create(sub)
  utils::write.table(df, file.path(sub, "occurrence.txt"), sep = sep,
                     row.names = FALSE, quote = FALSE, na = "")
  d
}

.cd_inv <- function(species, voucher = NULL, collected = "Sim") {
  n <- length(species)
  if (is.null(voucher)) voucher <- paste("H. Medeiros", 100 + seq_len(n) - 1L)
  data.frame(
    Voucher = voucher,
    Family = "Fabaceae",
    `Original determination` = species,
    Collected = collected,
    stringsAsFactors = FALSE, check.names = FALSE
  )
}

.cd_run <- function(inv, occ_dir, ..., env = parent.frame()) {
  fp <- withr::local_tempfile(fileext = ".xlsx", .local_envir = env)
  file.create(fp)
  testthat::local_mocked_bindings(
    .fp_query_to_field_sheet_df = function(path, sheet = NULL) inv,
    .monitora_to_field_sheet_df = function(path, sheet = 1, station_name = NULL) inv,
    .env = env
  )
  args <- utils::modifyList(
    list(fp_file_path = fp, input_type = "field_sheet", herbaria = "RB",
         path = occ_dir, dbdir = withr::local_tempfile(fileext = ".duckdb",
                                                       .local_envir = env)),
    list(...)
  )
  do.call(forplot_crossdet, args)
}


test_that("inventory names are normalized before comparison", {
  species <- c(
    "Inga edulis Mart.",  # same after dropping the author
    "Inga",               # genus only
    "Inga sp.",           # indeterminate marker
    "Inga cf. edulis",    # qualifier kept
    "Inga cf.",           # qualifier without epithet
    "Inga aff. indet",    # qualifier + indet
    "Inga Fabaceae",      # family-like second token
    "Inga inga",          # repeated genus
    "Inga 3x",            # non-alphabetic epithet
    "Fabaceae",           # family only
    "Inga vera"           # different species
  )
  occ_dir <- .cd_write(.cd_occ(length(species)))
  out <- .cd_run(.cd_inv(species), occ_dir)

  m <- out$matches
  sp <- stats::setNames(m$inventory_species, m$voucher)
  cls <- stats::setNames(m$diagnostic_class, m$voucher)
  expect_false("H. Medeiros 100" %in% m$voucher)
  expect_true(all(m$herbarium_species == "Inga edulis"))
  expect_identical(unname(cls["H. Medeiros 103"]), "different_determination")
  expect_identical(unname(cls["H. Medeiros 110"]), "different_determination")
  indet_rows <- paste("H. Medeiros", c(101, 102, 104, 105, 106, 107, 108, 109))
  expect_true(all(cls[indet_rows] == "inventory_indet_herbarium_det"))

  d <- stats::setNames(out$diagnostics$n, out$diagnostics$diagnostic_class)
  expect_identical(as.integer(d[["same_determination"]]), 1L)
  expect_identical(as.integer(d[["inventory_indet_herbarium_det"]]), 8L)
  expect_identical(as.integer(d[["different_determination"]]), 2L)
  expect_identical(nrow(out$unmatched_inventory), 0L)
})

test_that("herbarium-indeterminate records are flagged", {
  occ <- .cd_occ(2, genus = c("", "Inga"), specificEpithet = c("", "inga"),
                 scientificName = c("Fabaceae", "Inga"),
                 acceptedScientificName = c("", ""),
                 verbatimScientificName = c("", ""))
  out <- .cd_run(.cd_inv(c("Inga edulis", "Inga edulis")), .cd_write(occ))
  expect_identical(out$matches$diagnostic_class,
                   rep("inventory_det_herbarium_indet", 2))
  expect_identical(sort(out$matches$herbarium_species),
                   c("Indet indet", "Inga indet"))
})

test_that("semicolon-delimited occurrence files are read", {
  out <- .cd_run(.cd_inv("Inga vera"), .cd_write(.cd_occ(1), sep = ";"))
  expect_identical(out$matches$herbarium_species, "Inga edulis")
})

test_that("no candidate records gives an empty result and can be saved", {
  skip_if_not_installed("writexl")
  out_dir <- withr::local_tempdir()
  inv <- .cd_inv(c("Inga edulis", "Inga vera"),
                 voucher = c("H. Medeiros 900", "H. Medeiros 901"))
  expect_message(
    out <- .cd_run(inv, .cd_write(.cd_occ(1)), save = TRUE, dir = out_dir,
                   filename = "none"),
    "Saved:"
  )
  expect_identical(nrow(out$matches), 0L)
  expect_identical(out$diagnostics$diagnostic_class, "no_matches_found")
  expect_match(out$diagnostics$summary_message, "herbarium RB")
  expect_identical(out$unmatched_inventory$voucher, inv$Voucher)
  expect_identical(
    readxl::excel_sheets(file.path(out_dir, "none.xlsx")),
    c("Matches", "Diagnostics", "Unmatched_inventory")
  )
})

test_that("candidates without collector names give no usable keys", {
  skip_if_not_installed("writexl")
  out_dir <- withr::local_tempdir()
  occ <- .cd_occ(1, recordedBy = "")
  expect_message(
    out <- .cd_run(.cd_inv("Inga edulis"), .cd_write(occ), save = TRUE,
                   dir = file.path(out_dir, "new"), filename = "nokeys"),
    "Saved:"
  )
  expect_match(out$diagnostics$summary_message, "usable identity keys")
  expect_identical(nrow(out$unmatched_inventory), 1L)
  expect_true(file.exists(file.path(out_dir, "new", "nokeys.xlsx")))
})

test_that("candidates by another collector are reported as unmatched", {
  occ <- .cd_occ(1, recordedBy = "J. Silva")
  out <- .cd_run(.cd_inv("Inga edulis"), .cd_write(occ))
  expect_identical(nrow(out$matches), 0L)
  expect_identical(out$diagnostics$diagnostic_class, "no_matches_found")
  expect_identical(out$unmatched_inventory$voucher, "H. Medeiros 100")
})

test_that("numeric vouchers use collector_fallback and surname keys", {
  occ <- .cd_occ(1, recordedBy = "Medeiros, Hersion")
  out <- .cd_run(.cd_inv("Inga vera", voucher = "100"), .cd_write(occ),
                 collector_fallback = "J. Medeiros")
  expect_identical(out$matches$voucher, "100")
  expect_identical(out$matches$inventory_collector, "J. Medeiros")
})

test_that("only_collected filters inventory rows", {
  inv <- .cd_inv(c("Inga vera", "Inga vera"),
                 voucher = c("H. Medeiros 100", ""), collected = c("nao", "nao"))
  out <- .cd_run(inv, .cd_write(.cd_occ(1)))
  expect_identical(out$matches$voucher, "H. Medeiros 100")

  inv2 <- .cd_inv("Inga vera", voucher = "", collected = "nao")
  expect_error(.cd_run(inv2, .cd_write(.cd_occ(1))),
               "No inventory rows remained")
})

test_that("forplot_crossdet stops on unusable input", {
  occ_dir <- .cd_write(.cd_occ(1))
  expect_error(.cd_run(.cd_inv("Inga"), occ_dir, herbaria = c("RB", "K")),
               "exactly one herbarium code")
  expect_error(.cd_run(.cd_inv("Inga", voucher = "sem numero"), occ_dir),
               "No usable voucher identity keys")
  expect_error(.cd_run(.cd_inv("Inga"), .cd_write(.cd_occ(1, collectionCode = "K"))),
               "No records were loaded for herbarium 'RB'")
})

test_that("MONITORA input is routed through the MONITORA reader", {
  expect_message(
    out <- .cd_run(.cd_inv("Inga vera"), .cd_write(.cd_occ(1)),
                   input_type = "MONITORA", station_name = "EA-1",
                   verbose = TRUE),
    "Using existing local occurrence.txt"
  )
  expect_identical(out$matches$voucher, "H. Medeiros 100")
})

test_that("automatic download errors are explained", {
  cache <- withr::local_tempdir()
  testthat::local_mocked_bindings(
    .herbaria_cache_paths = function() {
      list(cache_dir = cache, duckdb_path = file.path(cache, "x.duckdb"))
    },
    .get_ipt_info = function(herbarium, ipt, resource_map = NULL) {
      data.frame(ipt = character(0), herbarium = character(0),
                 resource_id = character(0), archive_base = character(0),
                 resource_url = character(0))
    }
  )
  expect_error(.cd_run(.cd_inv("Inga"), NULL),
               "No JABOT IPT resource was found for herbarium 'RB'")
  expect_true(dir.exists(file.path(cache, "jabot_download")))

  testthat::local_mocked_bindings(
    .get_ipt_info = function(herbarium, ipt, resource_map = NULL) {
      data.frame(ipt = "jabot", herbarium = "RB", resource_id = "r",
                 archive_base = "x", resource_url = "y")
    },
    .download_dwca_one = function(...) NULL
  )
  expect_error(.cd_run(.cd_inv("Inga"), NULL),
               "no 'occurrence.txt' files were obtained")
})
