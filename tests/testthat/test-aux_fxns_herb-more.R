# Additional tests for the herbarium lookup helpers in R/aux_fxns_herb.R.
# Everything runs offline: IPT pages are local files or mocked readLines(),
# downloads are mocked, and DuckDB databases live in temporary directories.

.localize <- function(fun, bindings = list()) {
  env <- list2env(bindings, parent = environment(fun))
  environment(fun) <- env
  fun
}

.empty_ipt <- function() {
  data.frame(ipt = character(0), herbarium = character(0),
             resource_id = character(0), archive_base = character(0),
             resource_url = character(0), stringsAsFactors = FALSE)
}

.write_occ <- function(path, df) {
  dir.create(dirname(path), recursive = TRUE, showWarnings = FALSE)
  utils::write.table(df, path, sep = "\t", quote = FALSE, row.names = FALSE,
                     col.names = TRUE, na = "")
  path
}

.info_row <- function(resource_url = "https://example.org/resource?r=rb_test") {
  data.frame(ipt = "jabot", herbarium = "RB", resource_id = "rb_test",
             archive_base = "https://example.org/archive.do?r=",
             resource_url = resource_url, stringsAsFactors = FALSE)
}


# Voucher parsing --------------------------------------------------------------

test_that("number tokens are extracted from free text", {
  expect_identical(.extract_number_token(NA), NA_character_)
  expect_identical(.extract_number_token("  "), NA_character_)
  expect_identical(.extract_number_token("Cardoso 1234"), "1234")
  expect_identical(.extract_number_token("12 de 3456 a"), "3456")
  expect_identical(.extract_number_token("sem numero"), NA_character_)
})

test_that(".split_recordedby_people handles empty input and separators", {
  expect_identical(.split_recordedby_people(NA), character(0))
  expect_identical(.split_recordedby_people("   "), character(0))
  expect_identical(
    .split_recordedby_people("Cardoso, D.; Ottino, G. et al."),
    c("Cardoso, D.", "Ottino, G.")
  )
  expect_identical(
    .split_recordedby_people("A. Silva, B. Souza, C. Lima and D. Rocha"),
    c("A. Silva", "B. Souza", "C. Lima", "D. Rocha")
  )
  expect_identical(.split_recordedby_people("Ana | Bia & Caio"),
                   c("Ana", "Bia", "Caio"))
})

test_that(".collector_tokens_one handles 'Surname, Initials' and empty names", {
  na_tok <- c(primary = NA_character_, fallback = NA_character_)
  expect_identical(.collector_tokens_one(NA), na_tok)
  expect_identical(.collector_tokens_one(" et al. "), na_tok)
  expect_identical(.collector_tokens_one("Cardoso, Domingos B. O. S."),
                   c(primary = "dboscardoso", fallback = "cardoso"))
  expect_identical(.collector_tokens_one("Ottino,"),
                   c(primary = "ottino", fallback = "ottino"))
  expect_identical(.collector_tokens_one("Maria da Silva"),
                   c(primary = "msilva", fallback = "silva"))
})

test_that(".expand_collector_code requires a named vector", {
  expect_identical(.expand_collector_code("DC", NULL), NA_character_)
  expect_error(.expand_collector_code("DC", c("D. Cardoso")),
               "must be a named character vector")
  codes <- c(dc = "D. Cardoso", GO = "  ")
  expect_identical(.expand_collector_code(" DC ", codes), "D. Cardoso")
  expect_identical(.expand_collector_code("GO", codes), NA_character_)
  expect_identical(.expand_collector_code("XX", codes), NA_character_)
})

test_that(".parse_voucher_one applies collector fallbacks", {
  na_v <- c(collector = NA_character_, number = NA_character_)
  expect_identical(.parse_voucher_one(NA), na_v)
  expect_identical(.parse_voucher_one("123", "G. Ottino"),
                   c(collector = "G. Ottino", number = "123"))
  expect_identical(.parse_voucher_one("123"),
                   c(collector = NA_character_, number = "123"))
  expect_identical(.parse_voucher_one("GCO1001", " G. Ottino "),
                   c(collector = "G. Ottino", number = "1001"))
  expect_identical(.parse_voucher_one("GCO1001", collector_codes = c(GCO = "G.C. Ottino")),
                   c(collector = "G.C. Ottino", number = "1001"))
  expect_identical(.parse_voucher_one("/ 55", "Fallback"),
                   c(collector = "Fallback", number = "55"))
  expect_identical(.parse_voucher_one("D. Cardoso et al. 4321"),
                   c(collector = "D. Cardoso", number = "4321"))
  expect_identical(.parse_voucher_one("sem coletor", "Fallback"),
                   c(collector = "Fallback", number = NA_character_))
  expect_identical(.parse_voucher_one("sem coletor"),
                   c(collector = NA_character_, number = NA_character_))
})

test_that("URL builders return NA for incomplete input and scan numeric blocks", {
  expect_identical(.make_jabot_url(NA, "RB"), NA_character_)
  expect_identical(.make_jabot_url("123", NA), NA_character_)
  expect_identical(.make_jabot_url("123", " "), NA_character_)
  expect_match(.make_reflora_url("RB00012345x"), "idFigura=00012345$")
  expect_identical(.make_reflora_url("RB-12"), NA_character_)
})


# IPT discovery ----------------------------------------------------------------

test_that(".get_latest_version_info returns NAs for an unreadable page", {
  out <- .get_latest_version_info(file.path(withr::local_tempdir(), "missing.html"))
  expect_identical(out, list(version = NA_character_, published_on = NA_character_,
                             records = NA_character_))
})

test_that(".get_ipt_info prefers the jbrj resource and honours resource_map", {
  pages <- list(
    "https://ipt.jbrj.gov.br/jabot/dcat" = c(
      "https://ipt.jbrj.gov.br/jabot/archive.do?r=rb_jabot",
      "https://ipt.jbrj.gov.br/jabot/archive.do?r=huefs_herbarium"
    ),
    "https://ipt.jbrj.gov.br/jbrj/dcat" = "https://ipt.jbrj.gov.br/jbrj/archive.do?r=jbrj_rb&v=1.2",
    "https://ipt.jbrj.gov.br/reflora/dcat" = "no links here"
  )
  f <- .localize(.get_ipt_info, list(
    readLines = function(con, ...) {
      if (is.null(pages[[con]])) stop("offline") else pages[[con]]
    }
  ))

  out <- f(c("rb", "HUEFS"), ipt = "jabot")
  expect_identical(out$herbarium, c("HUEFS", "RB"))
  expect_identical(out$resource_id, c("huefs_herbarium", "jbrj_rb"))
  expect_match(out$archive_base[2], "/jbrj/archive.do")

  expect_identical(f("RB", ipt = "reflora"), .empty_ipt())
  expect_identical(f("ZZ", ipt = "jabot"), .empty_ipt())

  over <- f("RB", ipt = "reflora", resource_map = c("RB::reflora" = "rb_reflora"))
  expect_identical(out$ipt[1], "jabot")
  expect_identical(over$resource_id, "rb_reflora")
  expect_identical(over$resource_url,
                   "https://ipt.jbrj.gov.br/reflora/resource?r=rb_reflora")

  over2 <- f("ZZ", ipt = "jabot", resource_map = c("ZZ::jabot" = "jbrj_zz"))
  expect_match(over2$archive_base, "/jbrj/archive.do")

  expect_error(f("RB", ipt = "jabot", resource_map = c("jbrj_rb")),
               "must be a named character vector")
})


# DwC-A download ---------------------------------------------------------------

test_that(".download_dwca_one downloads, unzips and versions the archive", {
  root <- withr::local_tempdir()
  page <- file.path(root, "resource.html")
  writeLines(c("latestVersion 1.5", "2025-01-01"), page)

  urls <- character(0)
  testthat::local_mocked_bindings(
    download.file = function(url, destfile, ...) {
      urls <<- c(urls, url)
      writeLines("zip", destfile)
      0L
    },
    unzip = function(zipfile, exdir, ...) {
      writeLines("id", file.path(exdir, "occurrence.txt"))
      invisible(file.path(exdir, "occurrence.txt"))
    },
    .package = "utils"
  )

  expect_message(
    out <- .download_dwca_one(.info_row(page), dir = file.path(root, "dl"),
                              verbose = TRUE),
    "Downloading DwC-A: RB \\[jabot\\]"
  )
  expect_true(file.exists(file.path(out, "occurrence.txt")))
  expect_identical(urls, "https://example.org/archive.do?r=rb_test&v=1.5")
  expect_length(list.files(file.path(root, "dl"), pattern = "\\.zip$"), 0L)

  # Cached copy is reused, unless force_refresh is requested
  .download_dwca_one(.info_row(page), dir = file.path(root, "dl"))
  expect_length(urls, 1L)
  .download_dwca_one(.info_row(page), dir = file.path(root, "dl"),
                     force_refresh = TRUE)
  expect_length(urls, 2L)
})

test_that(".download_dwca_one retries without version and cleans up on failure", {
  root <- withr::local_tempdir()
  page <- file.path(root, "resource.html")
  writeLines("latestVersion 2.0", page)

  urls <- character(0)
  testthat::local_mocked_bindings(
    download.file = function(url, destfile, ...) {
      urls <<- c(urls, url)
      if (grepl("&v=", url, fixed = TRUE)) stop("404")
      writeLines("zip", destfile)
      0L
    },
    unzip = function(zipfile, exdir, ...) {
      writeLines("id", file.path(exdir, "occurrence.txt"))
    },
    .package = "utils"
  )
  out <- .download_dwca_one(.info_row(page), dir = root)
  expect_identical(urls, c("https://example.org/archive.do?r=rb_test&v=2.0",
                           "https://example.org/archive.do?r=rb_test"))
  expect_true(dir.exists(out))
})

test_that(".download_dwca_one returns NULL when download or unzip fails", {
  root <- withr::local_tempdir()
  missing_page <- file.path(root, "nope.html")

  testthat::local_mocked_bindings(
    download.file = function(...) stop("no network"),
    .package = "utils"
  )
  expect_null(.download_dwca_one(.info_row(missing_page), dir = root))
  expect_false(dir.exists(file.path(root, "dwca_jabot_RB_rb_test")))

  testthat::local_mocked_bindings(
    download.file = function(url, destfile, ...) { writeLines("x", destfile); 0L },
    unzip = function(...) stop("corrupt"),
    .package = "utils"
  )
  expect_null(.download_dwca_one(.info_row(missing_page), dir = root))

  testthat::local_mocked_bindings(
    unzip = function(zipfile, exdir, ...) writeLines("x", file.path(exdir, "meta.xml")),
    .package = "utils"
  )
  expect_null(.download_dwca_one(.info_row(missing_page), dir = root))
  expect_false(dir.exists(file.path(root, "dwca_jabot_RB_rb_test")))
})


# DuckDB helpers ---------------------------------------------------------------

test_that(".herbaria_db_connect explains a missing DBI/duckdb", {
  f <- .localize(.herbaria_db_connect, list(
    requireNamespace = function(pkg, ...) pkg != "duckdb"
  ))
  expect_error(f(), "Package 'duckdb' is required for herbarium lookup")
})

test_that(".duckdb_load_occurrence maps alternative column names", {
  skip_if_not_installed("DBI")
  skip_if_not_installed("duckdb")
  root <- withr::local_tempdir()
  con <- DBI::dbConnect(.duckdb_driver(), dbdir = ":memory:")
  withr::defer(DBI::dbDisconnect(con, shutdown = TRUE))
  withr::local_options(forplotR.herbaria_cache_dir = root)

  f_connect <- .localize(.herbaria_db_connect, list(
    .herbaria_cache_paths = function() list(duckdb_path = ":memory:")
  ))
  con2 <- f_connect()
  withr::defer(DBI::dbDisconnect(con2, shutdown = TRUE))

  expect_false(.duckdb_load_occurrence(con2, file.path(root, "none.txt"),
                                       "RB", "jabot", "r1"))

  occ <- .write_occ(file.path(root, "a", "occurrence.txt"), data.frame(
    barcode = c("RB1", "RB2"), coletor = c("D. Cardoso", "G. Ottino"),
    numero = c("n. 12", "34"), stringsAsFactors = FALSE
  ))
  expect_message(
    expect_true(.duckdb_load_occurrence(con2, occ, "RB", "jabot", "r1",
                                        verbose = TRUE)),
    "Inserted 2 records into database"
  )
  got <- DBI::dbGetQuery(con2, "SELECT * FROM occ_raw ORDER BY catalogNumber")
  expect_identical(got$catalogNumber, c("RB1", "RB2"))
  expect_identical(got$recordNumber_clean, c("12", "34"))
  expect_true(all(is.na(got$occurrenceID)))

  # Reloading the same resource replaces its rows; a file without a number
  # column is still loaded with NULL numbers.
  occ2 <- .write_occ(file.path(root, "b", "occurrence.txt"),
                     data.frame(catalogNumber = "RB9", stringsAsFactors = FALSE))
  expect_true(.duckdb_load_occurrence(con2, occ2, "RB", "jabot", "r1"))
  got2 <- DBI::dbGetQuery(con2, "SELECT * FROM occ_raw")
  expect_identical(got2$catalogNumber, "RB9")
  expect_identical(got2$recordNumber_clean, "")
})

test_that(".duckdb_match_resource reports why nothing was indexed", {
  skip_if_not_installed("DBI")
  skip_if_not_installed("duckdb")
  con <- .localize(.herbaria_db_connect, list(
    .herbaria_cache_paths = function() list(duckdb_path = ":memory:")
  ))()
  withr::defer(DBI::dbDisconnect(con, shutdown = TRUE))

  DBI::dbWriteTable(con, "occ_raw", data.frame(
    catalogNumber = c("RB1", "RB2", "RB3"), occurrenceID = c("1", "2", "3"),
    recordedBy = c("", "D. Cardoso", "G. Ottino; ???"),
    recordNumber = c("10", "20", "30"), recordNumber_clean = c("10", "20", "30"),
    herbarium = "RB", source = "jabot", resource_id = "r1"
  ), append = TRUE)

  m <- function(p, fb = character(0), verbose = TRUE) {
    .duckdb_match_resource(con, "RB", "jabot", "r1", p, fb, verbose = verbose)
  }

  expect_false(m(NA_character_, verbose = FALSE))
  expect_message(expect_false(m("dcardoso||99")), "No candidate records found")
  # Only a record without people
  expect_false(m("x||10", verbose = FALSE))
  expect_message(expect_false(m("someone||20")),
                 "none matched requested keys")
  expect_message(expect_true(m("gottino||30", "cardoso||20")),
                 "Index built with 2 matched rows")
  idx <- DBI::dbGetQuery(con, "SELECT * FROM herbaria_index ORDER BY key")
  expect_identical(idx$catalogNumber, c("RB2", "RB3"))
})


# Full lookup workflow ---------------------------------------------------------

test_that(".herbaria_lookup_links falls back to REFLORA for unresolved vouchers", {
  skip_if_not_installed("DBI")
  skip_if_not_installed("duckdb")
  root <- withr::local_tempdir()
  withr::local_options(forplotR.herbaria_cache_dir = file.path(root, "cache"))
  withr::local_dir(root)
  dir.create("jabot_download")
  dir.create("reflora_download")

  dwca <- function(name, df) {
    d <- file.path(root, name)
    .write_occ(file.path(d, "occurrence.txt"), df)
    d
  }
  jabot_dir <- dwca("dw_jabot", data.frame(
    catalogNumber = "RB100", occurrenceID = "", recordedBy = "Cardoso, D.",
    recordNumber = "100", stringsAsFactors = FALSE
  ))
  reflora_dir <- dwca("dw_reflora", data.frame(
    catalogNumber = "", occurrenceID = "https://x.org/id/777",
    recordedBy = "G. Ottino", recordNumber = "200", stringsAsFactors = FALSE
  ))

  calls <- character(0)
  testthat::local_mocked_bindings(
    .get_ipt_info = function(herbarium, ipt, resource_map = NULL) {
      if (herbarium == "HUEFS") return(.empty_ipt())
      data.frame(ipt = ipt, herbarium = herbarium,
                 resource_id = paste0("rb_", ipt), archive_base = "x",
                 resource_url = "y", stringsAsFactors = FALSE)
    },
    .download_dwca_one = function(info_row, dir, ...) {
      calls <<- c(calls, info_row$ipt)
      if (info_row$ipt == "jabot") jabot_dir else reflora_dir
    }
  )

  fp <- data.frame(
    Voucher = c("D. Cardoso 100", "G. Ottino 200", "Z. Nobody 300"),
    stringsAsFactors = FALSE
  )
  msgs <- testthat::capture_messages(
    out <- .herbaria_lookup_links(fp, herbaria = c("RB", "HUEFS"),
                                  keep_downloads = FALSE, verbose = TRUE)
  )

  expect_match(out[1], "codtestemunho=RB100.*title='JABOT'")
  expect_match(out[2], "idFigura=777.*title='REFLORA'")
  expect_true(is.na(out[3]))
  expect_identical(calls, c("jabot", "reflora"))
  expect_true(any(grepl("HERBARIA LOOKUP", msgs)))
  expect_true(any(grepl("No JABOT resource found for HUEFS", msgs)))
  expect_true(any(grepl("No REFLORA resource found for HUEFS", msgs)))
  expect_false(dir.exists(file.path(root, "jabot_download")))
  expect_false(dir.exists(file.path(root, "reflora_download")))
})

test_that(".herbaria_lookup_links returns NA without keys or after failed loads", {
  skip_if_not_installed("DBI")
  skip_if_not_installed("duckdb")
  root <- withr::local_tempdir()
  withr::local_options(forplotR.herbaria_cache_dir = root)

  expect_identical(
    .herbaria_lookup_links(data.frame(Voucher = c("", NA)), "RB", verbose = FALSE),
    rep(NA_character_, 2)
  )

  testthat::local_mocked_bindings(
    .get_ipt_info = function(herbarium, ipt, resource_map = NULL) {
      data.frame(ipt = ipt, herbarium = herbarium, resource_id = "r",
                 archive_base = "x", resource_url = "y", stringsAsFactors = FALSE)
    },
    .download_dwca_one = function(...) NULL
  )
  out <- .herbaria_lookup_links(data.frame(Voucher = "D. Cardoso 1"), "RB",
                                keep_downloads = TRUE, verbose = FALSE)
  expect_identical(out, NA_character_)
})


# Link resolution --------------------------------------------------------------

test_that(".resolve_links_from_index uses fallback keys and skips invalid hits", {
  parsed <- data.frame(
    primary_key = c("dcardoso||1", NA, "gottino||3", "x||4"),
    fallback_key = c("cardoso||1", "silva||2", "ottino||3", "x||4"),
    stringsAsFactors = FALSE
  )
  fp <- data.frame(Voucher = 1:4)
  idx <- data.frame(
    key = c("cardoso||1", "silva||2", "gottino||3"),
    key_type = c("fallback", "fallback", "primary"),
    catalogNumber = c("RB1", "", "RB3"), occurrenceID = c("", "", "abc"),
    herbarium = "RB", source = "x", stringsAsFactors = FALSE
  )

  empty <- .resolve_links_from_index(idx[0, ], 1:4, "jabot", fp, parsed)
  expect_false(any(empty$resolved))

  j <- .resolve_links_from_index(idx, 1:4, "jabot", fp, parsed)
  expect_identical(j$resolved, c(TRUE, FALSE, TRUE, FALSE))
  expect_match(j$links[1], "codtestemunho=RB1'")

  r <- .resolve_links_from_index(idx, 1:4, "reflora", fp, parsed)
  # Row 3 has an occurrenceID but no numeric block, so no REFLORA URL
  expect_false(any(r$resolved))

  idx$occurrenceID[3] <- "4567"
  r2 <- .resolve_links_from_index(idx, 1:4, "reflora", fp, parsed)
  expect_identical(r2$resolved, c(FALSE, FALSE, TRUE, FALSE))
})

test_that(".load_resource_once reports a failed occurrence load", {
  root <- withr::local_tempdir()
  f <- .localize(.load_resource_once, list(
    .download_dwca_one = function(...) root,
    .duckdb_load_occurrence = function(...) FALSE
  ))
  res <- f(data.frame(herbarium = "RB", source = "jabot", resource_id = "r"),
           dir_name = root, already_loaded = "A::b::c", con = NULL)
  expect_false(res$ok)
  expect_identical(res$already_loaded, "A::b::c")
})

test_that(".parse_census_identity keeps going when some vouchers have no collector", {
  out <- .parse_census_identity(c("D. Cardoso 100", "", "123"))
  expect_s3_class(out, "data.frame")
  expect_equal(nrow(out), 3L)
  expect_false(anyNA(out$primary_key[1]))
  expect_true(all(is.na(out$primary_key[2:3])))
})

test_that("JABOT links open the image viewer of the specimen by its catalogue number", {
  expect_identical(
    .make_jabot_url(" RB01601078 ", "rb"),
    "https://jabot.jbrj.gov.br/v2/regua/visualizador.php?r=true&colbot=rb&codtestemunho=RB01601078"
  )
})
