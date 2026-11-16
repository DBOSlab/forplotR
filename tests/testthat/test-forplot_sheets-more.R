# Additional tests for forplot_sheets(): BRAHMS and custom herbarium formats,
# Spanish notes, collector inference and elevation extraction. All offline.

.more_fixtures <- function(env = parent.frame()) {
  td <- withr::local_tempdir(.local_envir = env)

  fp <- file.path(td, "RUS01.xlsx")
  openxlsx::write.xlsx(forplotR::RUS01, fp)

  mon <- file.path(td, "TUM.xlsx")
  openxlsx::write.xlsx(forplotR::TUM, mon)

  empty_template <- function(cols, path) {
    tpl <- as.data.frame(
      stats::setNames(replicate(length(cols), character(0), simplify = FALSE), cols)
    )
    openxlsx::write.xlsx(tpl, path)
    path
  }

  jabot <- empty_template(
    c("family", "genus", "sp1", "number", "collector", "addcoll", "colldd",
      "collmm", "collyy", "country", "majorarea", "minorarea", "uc", "projeto",
      "locnotes", "latitude", "longitude", "altprof", "notes"),
    file.path(td, "jabot.xlsx")
  )
  brahms <- empty_template(
    c("FamilyName", "GenusName", "SpeciesName", "FieldNumber", "Collectors",
      "AdditionalCollectors", "CollectionDay", "CollectionMonth",
      "CollectionYear", "CountryName", "MajorAdminName", "MinorAdminName",
      "LocalityName", "LocalityNotes", "Latitude", "Longitude", "Elevation",
      "DescriptionText"),
    file.path(td, "brahms.xlsx")
  )
  custom_cols <- c("Fam", "Gen", "Epi", "Num", "Coll", "AddColl", "Day",
                   "Month", "Year", "Ctry", "State", "Muni", "Area", "LocN",
                   "Proj", "Lat", "Lon", "Elev", "Remarks")
  custom <- empty_template(custom_cols, file.path(td, "custom.xlsx"))

  list(dir = td, fp = fp, mon = mon, jabot = jabot, brahms = brahms,
       custom = custom)
}

.more_run <- function(fx, ...) {
  args <- utils::modifyList(
    list(
      fp_file_path = fx$fp,
      herb_file_path = fx$jabot,
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

.custom_map <- list(
  family = "Fam", genus = "Gen", sp1 = "Epi", number = "Num",
  collector = "Coll", addcoll = "AddColl", colldd = "Day", collmm = "Month",
  collyy = "Year", country = "Ctry", majorarea = "State", minorarea = "Muni",
  protectedarea = "Area", locnotes = "LocN", project = "Proj", lat = "Lat",
  long = "Lon", alt = "Elev", notes = "Remarks"
)


# Herbarium formats -----------------------------------------------------------

test_that("forplot_sheets fills a BRAHMS template", {
  fx <- .more_fixtures()
  jab <- .more_run(fx)
  out <- .more_run(fx, herb_file_path = fx$brahms, herbarium_format = "brahms",
                   protectedarea = "REBIO Augusto Ruschi", locnotes = "trail")

  expect_s3_class(out, "data.frame")
  expect_equal(nrow(out), nrow(jab))
  expect_true(all(out$Collectors == "G.C. Ottino"))
  expect_true(all(out$CollectionDay == "26"))
  expect_true(all(out$CollectionMonth == "11"))
  expect_true(all(out$CollectionYear == "2014"))
  expect_true(all(out$LocalityName == "REBIO Augusto Ruschi"))
  expect_true(all(out$LocalityNotes == "trail"))
  expect_true(all(out$Elevation == 800))
  expect_identical(out$FamilyName, jab$family)
  expect_identical(out$GenusName, jab$genus)
  expect_identical(out$SpeciesName, jab$sp1)
  expect_identical(out$FieldNumber, jab$number)
  expect_identical(out$DescriptionText, jab$notes)
  expect_false(any(grepl("indet", out$GenusName, ignore.case = TRUE), na.rm = TRUE))

  written <- list.files(fx$dir, pattern = "plot_to_herb_sheet\\.xlsx$",
                        recursive = TRUE)
  expect_length(written, 1L)
})

test_that("forplot_sheets fills a custom template through column_map", {
  fx <- .more_fixtures()
  jab <- .more_run(fx, project = "PPBio")
  out <- .more_run(fx, herb_file_path = fx$custom, herbarium_format = "custom",
                   column_map = .custom_map, project = "PPBio",
                   protectedarea = "REBIO", filename = "custom_labels")

  expect_equal(nrow(out), nrow(jab))
  expect_true(all(out$Proj == "PPBio"))
  expect_true(all(out$Area == "REBIO"))
  expect_true(all(out$State == jab$majorarea))
  expect_true(all(out$Lat == -19.9))
  expect_true(all(out$Lon == -40.55))
  expect_identical(out$Fam, jab$family)
  expect_identical(out$Gen, jab$genus)
  expect_identical(out$Epi, jab$sp1)
  expect_identical(out$Num, jab$number)
  expect_identical(out$Remarks, jab$notes)

  written <- list.files(fx$dir, pattern = "^custom_labels\\.xlsx$",
                        recursive = TRUE, full.names = TRUE)
  expect_length(written, 1L)
  expect_true("Remarks" %in% names(openxlsx::read.xlsx(written)))
})


# Notes -----------------------------------------------------------------------

test_that("forplot_sheets writes Spanish notes", {
  fx <- .more_fixtures()
  out <- .more_run(fx, language = "es")
  expect_true(all(startsWith(out$notes, "Árbol")))
  expect_true(any(grepl("a la altura del pecho (DAP)", out$notes, fixed = TRUE)))
  expect_true(any(grepl("en la subparcelas", out$notes, fixed = TRUE)))
})

test_that("add_census_notes = FALSE drops census notes in every language", {
  fx <- .more_fixtures()
  # RUS01 tree 1 has the census note "face abaxial ferruginea."
  for (lang in c("en", "pt", "es")) {
    with_notes <- .more_run(fx, language = lang)
    without <- .more_run(fx, language = lang, add_census_notes = FALSE)
    expect_true(any(grepl("face abaxial", with_notes$notes)), info = lang)
    expect_false(any(grepl("face abaxial", without$notes)), info = lang)
    expect_true(all(nchar(without$notes) <= nchar(with_notes$notes)), info = lang)
  }
})

test_that(".build_notes translates flag and light codes and tidies punctuation", {
  fp <- data.frame(
    Flag1 = c("ac", NA, ""), LI = c("3b", "1", NA),
    `Census Notes` = c("note", "", NA), `New Tag No` = c("7", "", NA),
    T1 = c("2", "", NA), X = c("1.5000000001", NA, NA), Y = c("3", NA, NA),
    D = c(12.3456, NA, 5),
    check.names = FALSE, stringsAsFactors = FALSE
  )
  herb <- data.frame(notes = rep(NA_character_, 3))

  en <- .build_notes(fp, herb, notes = "notes", language = "en")$notes
  expect_identical(
    en[1],
    paste0("Tree, 12.35cm DBH, alive normal, leaning by ≥10%, note, ",
           "under high vertical illumination (>50%). Individual #7 in the ",
           "subplot 2, x = 1.5m, y = 3m.")
  )
  expect_identical(en[2], "Tree, not under direct light.")
  expect_identical(en[3], "Tree, 5cm DBH")

  pt <- .build_notes(fp, herb, notes = "notes", language = "pt")$notes
  expect_match(pt[1], "viva normal, inclinada", fixed = TRUE)
  expect_match(pt[1], "na subparcela 2", fixed = TRUE)
})

test_that("description tables exist for every supported language", {
  for (lang in c("en", "pt", "es")) {
    d <- .get_descriptions(language = lang)
    l <- .get_luminosity(language = lang)
    expect_equal(nrow(d), 21L)
    expect_equal(nrow(l), 8L)
    expect_false(anyNA(d$description))
  }
})

test_that(".convert_to_lowercase lowercases, swaps periods and blanks NA", {
  expect_identical(
    .convert_to_lowercase(c("Fruto Verde. Flor", NA)),
    c("fruto verde, flor", "")
  )
})


# Collector inference ---------------------------------------------------------

test_that("ForestPlots collectors are inferred from the Team header", {
  fx <- .more_fixtures()
  out <- .more_run(fx, collector = NULL)
  expect_true(all(out$collector == "GC Ottino"))
  expect_true(all(grepl("^S Pessoa, LF", out$addcoll)))
})

test_that("ForestPlots without a Team header warns and keeps given collector", {
  fx <- .more_fixtures()
  rus <- forplotR::RUS01
  names(rus)[grepl("^Team", names(rus))] <- "X13"
  fp2 <- file.path(fx$dir, "noteam.xlsx")
  openxlsx::write.xlsx(rus, fp2)

  expect_warning(
    out <- .more_run(fx, fp_file_path = fp2, addcoll = NULL),
    "Could not infer collector metadata"
  )
  expect_true(all(out$collector == "G.C. Ottino"))
})

test_that("ForestPlots sheets without a Date header are rejected", {
  fx <- .more_fixtures()
  rus <- forplotR::RUS01
  names(rus)[grepl("^Date", names(rus))] <- "X10"
  fp2 <- file.path(fx$dir, "nodate.xlsx")
  openxlsx::write.xlsx(rus, fp2)

  expect_error(.more_run(fx, fp_file_path = fp2),
               "collection date in the format")
})

test_that("MONITORA metadata fill protected area, project and census year", {
  fx <- .more_fixtures()
  out <- .more_run(fx, fp_file_path = fx$mon, input_type = "monitora",
                   station_name = "EA-1", majorarea = "AP",
                   minorarea = "Laranjal do Jari", lat = 1.5, long = -52.5,
                   alt = 300, collector = NULL)

  expect_true(all(grepl("Tumucumaque", out$uc)))
  expect_true(all(!is.na(out$projeto) & nzchar(out$projeto)))
  expect_true(all(out$collyy == as.character(max(forplotR::TUM$Ano))))
  expect_true(all(is.na(out$colldd)))
})

test_that("MONITORA collector falls back to team and then to the argument", {
  fx <- .more_fixtures()
  tum <- forplotR::TUM
  tum[["voucher/Coletor"]] <- NA_character_
  mon2 <- file.path(fx$dir, "tum_novoucher.xlsx")
  openxlsx::write.xlsx(tum, mon2)

  out <- .more_run(fx, fp_file_path = mon2, input_type = "monitora",
                   station_name = "EA-1", majorarea = "AP",
                   minorarea = "Laranjal do Jari", lat = 1.5, long = -52.5,
                   alt = 300, collector = "Fallback")
  # Team metadata exists in TUM, so it wins over the argument
  expect_false(any(out$collector == "Fallback", na.rm = TRUE))

  tum$nome_coletores <- NA_character_
  mon3 <- file.path(fx$dir, "tum_noteam.xlsx")
  openxlsx::write.xlsx(tum, mon3)

  out3 <- .more_run(fx, fp_file_path = mon3, input_type = "monitora",
                    station_name = "EA-1", majorarea = "AP",
                    minorarea = "Laranjal do Jari", lat = 1.5, long = -52.5,
                    alt = 300, collector = "Fallback")
  if (nrow(out3) > 0) expect_true(all(out3$collector == "Fallback"))

  expect_warning(
    .more_run(fx, fp_file_path = mon3, input_type = "monitora",
              station_name = "EA-1", majorarea = "AP",
              minorarea = "Laranjal do Jari", lat = 1.5, long = -52.5,
              alt = 300, collector = NULL),
    "Collector could not be inferred"
  )
})


# Elevation -------------------------------------------------------------------

test_that("altitude is extracted from a (mocked) elevation raster", {
  skip_if_not_installed("geodata")
  skip_if_not_installed("terra")
  fx <- .more_fixtures()
  r <- terra::rast(xmin = -41, xmax = -40, ymin = -20.5, ymax = -19.5,
                   nrows = 2, ncols = 2, crs = "EPSG:4326", vals = 123)

  testthat::local_mocked_bindings(
    elevation_global = function(...) r,
    .package = "geodata"
  )
  out <- .more_run(fx, alt = NULL)
  expect_true(all(out$altprof == 123))
})
