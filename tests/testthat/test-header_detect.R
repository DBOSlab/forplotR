# Header-row detection, plot metadata and bilingual column names.
#
# The Indigenous Land (Panara) field sheet has plot metadata on row 1 and a
# bilingual header on row 2, plus a duplicated diameter column. The package
# ships only the blank template of that sheet; the tests fill it with made-up
# trees. No real plot data is used here.

.hd_template <- function() {
  f <- system.file("extdata", "panara_field_sheet_template.xlsx", package = "forplotR")
  skip_if(!nzchar(f), "Panara template not installed")
  f
}

.hd_write_raw <- function(rows, env = parent.frame()) {
  width <- max(lengths(rows))
  m <- t(vapply(rows, function(r) c(r, rep(NA_character_, width - length(r))),
                character(width)))
  path <- withr::local_tempfile(fileext = ".xlsx", .local_envir = env)
  writexl::write_xlsx(as.data.frame(m, stringsAsFactors = FALSE), path, col_names = FALSE)
  path
}

# The template header with fictitious metadata and `n` fictitious trees:
# 10 x 10 m subplots numbered 1..20 without subplot 7, column 19 repeating the
# diameter of column 10 (as the helper column of the template does).
.hd_filled <- function(n = 60, d_dup = NULL, env = parent.frame()) {
  top <- suppressMessages(readxl::read_excel(.hd_template(), col_names = FALSE,
                                             col_types = "text", .name_repair = "minimal"))
  hdr <- as.character(unlist(top[2, ]))
  meta <- rep(NA_character_, length(hdr))
  lab <- as.character(unlist(top[1, ]))
  meta[1] <- paste(lab[1], "ZZ01")
  meta[2] <- paste(lab[2], "Mata de Teste")
  meta[10] <- paste(lab[10], "26/11/2014")
  meta[12] <- paste(lab[12], "Ana Souza, Bia Lima, Caio Reis & Davi Melo")

  sub <- setdiff(1:20, 7)
  i <- seq_len(n)
  d <- 100 + (i * 37) %% 900 + 0.5
  fam <- c("Fabaceae", "Myrtaceae", "Arecaceae", "Burseraceae")[(i %% 4) + 1]
  spp <- c("Inga alba", "Eugenia sp.", "Attalea maripa", "Protium decorum")[(i %% 4) + 1]
  body <- data.frame(
    tag = i, grp = ifelse(i %in% c(11, 12), 11, NA), t1 = sub[(i - 1) %% length(sub) + 1],
    x = (i * 3) %% 11, y = 0.1 + (i * 7) %% 10, fam = fam, spp = spp,
    loc = paste("nome", (i %% 4) + 1), morpho = NA, d = d, pom = 1300,
    flag = c("A", "B", "EQ", "CFH", "ZYB")[(i %% 5) + 1], height = NA,
    voucher = ifelse(i %% 3 == 0, NA, paste("X.Coletor", 100 + i)), silica = NA,
    coll = ifelse(i %% 3 == 0, NA, "yes"), notes = NA, vern = NA,
    d2 = if (is.null(d_dup)) d else d_dup, cap = round(d / 10 * pi, 1),
    stringsAsFactors = FALSE
  )
  rows <- c(list(meta, hdr), lapply(seq_len(n), function(k) as.character(unlist(body[k, ]))))
  .hd_write_raw(rows, env = env)
}


# Blank template ---------------------------------------------------------------

test_that("the blank Panara template has its header on row 2", {
  expect_identical(detect_header_row(.hd_template()), 2L)
  x <- read_plot_sheet(.hd_template(), verbose = FALSE)
  expect_identical(nrow(x), 0L)
  expect_identical(ncol(x), 20L)
  cols <- attr(x, "columns")
  expect_true(all(c("New Tag No", "New Stem Grouping", "T1", "X", "Y", "Family",
                    "Original determination", "D", "POM", "Flag1", "Voucher") %in%
                    cols$canonical))
  # metadata labels without values are not metadata, and not column names
  meta <- read_plot_metadata(.hd_template())
  expect_true(is.na(meta$plot_code))
  expect_length(meta$team, 0L)
  expect_identical(meta$header_row, 2L)
})


# Filled sheet (fictitious data) -----------------------------------------------

test_that("the header of a filled Panara sheet is found on row 2", {
  expect_identical(detect_header_row(.hd_filled()), 2L)
})

test_that("the metadata row is read into standard fields", {
  meta <- read_plot_metadata(.hd_filled())
  expect_identical(meta$plot_code, "ZZ01")
  expect_identical(meta$plot_name, "Mata de Teste")
  expect_identical(meta$date, as.Date("2014-11-26"))
  expect_identical(meta$team, c("Ana Souza", "Bia Lima", "Caio Reis", "Davi Melo"))
  expect_identical(meta$header_row, 2L)
})

test_that("the table is read from the detected header", {
  path <- .hd_filled(60)
  msgs <- testthat::capture_messages(x <- read_plot_sheet(path))
  expect_match(msgs[1], "Header detected on row 2; 4 metadata fields read (plot_code = ZZ01",
               fixed = TRUE)
  expect_match(msgs[1], "date = 2014-11-26", fixed = TRUE)

  expect_identical(nrow(x), 60L)
  expect_identical(attr(x, "header_row"), 2L)
  expect_length(unique(x$T1), 19L)
  expect_identical(setdiff(1:20, as.numeric(x$T1)), 7L)
  expect_identical(anyDuplicated(x$`New Tag No`), 0L)
  expect_length(unique(x$Family), 4L)
  expect_length(unique(x$`Original determination`), 4L)
  expect_false(anyNA(as.numeric(x$D)))
  expect_true(all(as.numeric(x$X) >= 0 & as.numeric(x$X) <= 10))
  expect_identical(sum(!is.na(x$Voucher)), 40L)

  # sparse and empty columns are fine; compound flag codes are kept as written
  expect_true(all(is.na(x$Height)))
  expect_identical(sum(!is.na(x$`New Stem Grouping`)), 2L)
  expect_true(all(c("EQ", "CFH", "ZYB") %in% x$Flag1))
})

test_that("bilingual columns are mapped without intervention", {
  x <- read_plot_sheet(.hd_filled(), verbose = FALSE)
  cols <- attr(x, "columns")
  expect_true(all(c("New Tag No", "New Stem Grouping", "T1", "X", "Y", "Family",
                    "Original determination", "D", "POM", "Flag1", "Voucher") %in%
                    cols$canonical))
  expect_true(startsWith(cols$header[match("T1", cols$canonical)], "T1 | "))
  expect_true(startsWith(cols$header[match("Collected", cols$canonical)], "Coletado | "))
  expect_true(startsWith(cols$header[match("Family", cols$canonical)], "Fam"))
})

test_that("the diameter used is 'D | Wy' (column J), not its duplicate", {
  path <- .hd_filled(12, d_dup = 9000 + seq_len(12))
  msgs <- testthat::capture_messages(x <- read_plot_sheet(path))
  cols <- attr(x, "columns")
  d <- cols[match("D", cols$column), ]
  expect_identical(d$position, 10L)
  expect_identical(d$header, "D | Wy")
  d2 <- cols[match("D_2", cols$column), ]
  expect_identical(d2$position, 19L)
  expect_identical(d2$duplicate_of, "D")
  expect_true(is.na(d2$canonical))
  expect_true(any(grepl("repeats the name 'D' and was read as 'D_2'", msgs, fixed = TRUE)))
  expect_true(all(as.numeric(x$D) < 1100))
  expect_true(all(as.numeric(x$D_2) > 9000))

  for (type in c("field_sheet_ti", "field_sheet")) {
    h <- .harmonize_plot_input(path, type, verbose = FALSE)
    expect_true(all(as.numeric(h$fp_sheet$D) < 1100), info = type)
  }
  msgs <- testthat::capture_messages(.harmonize_plot_input(path, "field_sheet_ti", verbose = TRUE))
  expect_true(any(grepl("repeats the name 'D'; using column 10 ('D | Wy')", msgs, fixed = TRUE)))

  # a column mapped by the user still wins
  cm <- .build_col_map(colname_dbh = "D")
  h <- .harmonize_plot_input(path, "field_sheet", verbose = FALSE, col_map = cm)
  expect_true(all(as.numeric(h$fp_sheet$D) > 9000))
})

test_that("forcing header_row = 2 gives the same result as detection", {
  path <- .hd_filled()
  auto <- read_plot_sheet(path, verbose = FALSE)
  forced <- read_plot_sheet(path, header_row = 2, verbose = FALSE)
  expect_identical(forced, auto)
  expect_identical(read_plot_metadata(path, header_row = 2), read_plot_metadata(path))
})

test_that("the sheet harmonizes the same way under both field-sheet types", {
  path <- .hd_filled(60)
  ti <- .harmonize_plot_input(path, "field_sheet_ti", verbose = FALSE)
  fs <- .harmonize_plot_input(path, "field_sheet", verbose = FALSE)

  for (h in list(ti, fs)) {
    expect_identical(h$plot_code, "ZZ01")
    expect_identical(h$plot_name, "Mata de Teste")
    expect_identical(h$census_date, as.Date("2014-11-26"))
    expect_length(h$team_members, 4L)
    expect_identical(h$header_row, 2L)
    s <- h$fp_sheet
    # 60 stems, one two-stem tree consolidated
    expect_identical(nrow(s), 59L)
    expect_length(unique(as.numeric(s$T1)), 19L)
    expect_length(unique(s$Family), 4L)
    expect_false(anyNA(as.numeric(s$D)))
    expect_false(any(grepl("|", s$`Original determination`, fixed = TRUE)))
  }
  expect_true("Local name" %in% names(ti$fp_sheet))
  forced <- .harmonize_plot_input(path, "field_sheet_ti", verbose = FALSE, header_row = 2)
  expect_identical(forced$fp_sheet, ti$fp_sheet)
})

test_that("forplot_balance() writes a report for the bilingual sheet", {
  out <- suppressMessages(suppressWarnings(forplot_balance(
    .hd_filled(60), input_type = "field_sheet_ti", plot_size = 0.2, plot_width_m = 20,
    plot_length_m = 100, render_pdf = FALSE, write_xlsx = FALSE,
    dir = withr::local_tempdir(), verbose = FALSE
  )))
  expect_true(file.exists(out$html_report))
  html <- paste(readLines(out$html_report, warn = FALSE, encoding = "UTF-8"), collapse = "\n")
  # grepl() rather than expect_match(): a failure would print the whole report
  expect_true(grepl("ZZ01", html, fixed = TRUE))
  expect_true(grepl("Date: 26/11/2014", html, fixed = TRUE))
  # 2 x 10 grid; subplot 7 has no trees
  expect_true(grepl("20 subplots of 10", html, fixed = TRUE))
  expect_true(grepl("in 19 subplots", html, fixed = TRUE))
})


# Sheets that already worked ---------------------------------------------------

test_that("the Ruschi field sheet (RUS01) is read as before", {
  fp <- withr::local_tempfile(fileext = ".xlsx")
  openxlsx::write.xlsx(forplotR::RUS01, fp)

  expect_identical(detect_header_row(fp), 2L)
  meta <- read_plot_metadata(fp)
  expect_identical(meta$plot_code, "RUS01")
  expect_identical(meta$plot_name, "REBIO Augusto Ruschi")

  h <- .harmonize_plot_input(fp, "field_sheet", verbose = FALSE)
  expect_identical(h$plot_code, "RUS01")
  expect_identical(h$header_row, 2L)
  expect_true(all(c("New Tag No", "T1", "X", "Y", "Family", "Original determination",
                    "D", "Voucher", "Collected") %in% names(h$fp_sheet)))
  # the diameter is the field-sheet D (column 10), as it always was
  raw <- as.data.frame(forplotR::RUS01)
  first_tag <- as.character(h$fp_sheet$`New Tag No`[1])
  expect_equal(as.numeric(h$fp_sheet$D[1]),
               as.numeric(raw[[10]][match(first_tag, as.character(raw[[1]]))]))
  expect_identical(.harmonize_plot_input(fp, "field_sheet", verbose = FALSE, header_row = 2)$fp_sheet,
                   h$fp_sheet)
})

test_that("a sheet with the header on row 1 keeps row 1", {
  path <- withr::local_tempfile(fileext = ".xlsx")
  writexl::write_xlsx(data.frame(
    `New Tag No` = c("1", "2"), T1 = c(1, 1), X = c(1, 2), Y = c(3, 4),
    Family = c("Fabaceae", "Myrtaceae"), `Original determination` = c("Inga alba", "Eugenia sp."),
    D = c(120, 150), Voucher = c("AB 1", NA), Collected = c("yes", NA),
    check.names = FALSE, stringsAsFactors = FALSE
  ), path)
  expect_silent(hr <- detect_header_row(path))
  expect_identical(hr, 1L)
  meta <- read_plot_metadata(path)
  expect_true(is.na(meta$plot_code))
  expect_length(meta$team, 0L)
  x <- read_plot_sheet(path, verbose = FALSE)
  expect_identical(nrow(x), 2L)
  expect_identical(names(x)[1:4], c("New Tag No", "T1", "X", "Y"))
  expect_match(testthat::capture_messages(read_plot_sheet(path))[1],
               "Header detected on row 1; no metadata rows above it.", fixed = TRUE)

  h <- .harmonize_plot_input(path, "field_sheet", verbose = FALSE, has_user_meta = TRUE)
  expect_identical(h$header_row, 1L)
  expect_identical(nrow(h$fp_sheet), 2L)
})


# Detection rules --------------------------------------------------------------

test_that("header detection needs three known names and warns otherwise", {
  path <- .hd_write_raw(list(c("a", "b", "c"), c("1", "2", "3"), c("4", "5", "6")))
  expect_warning(hr <- detect_header_row(path), "Could not find a header row")
  expect_identical(hr, 1L)

  # the first of two equally good rows wins
  tie <- .hd_write_raw(list(c("Tag", "T1", "X", "Y"), c("Tag", "T1", "X", "Y"), c("1", "1", "2", "3")))
  expect_identical(detect_header_row(tie), 1L)

  # metadata-looking cells do not count as column names
  meta_first <- .hd_write_raw(list(
    c("Tag: 1", "T1: 2", "X: 3", "Y: 4", "Family: z"),
    c("note", "", "", "", ""),
    c("Placa", "Subparcela", "X (m)", "Y (m)", "Familia", "Especie", "DAP"),
    c("1", "1", "2", "3", "Fabaceae", "Inga alba", "120")
  ))
  expect_identical(detect_header_row(meta_first), 3L)
  expect_warning(detect_header_row(meta_first, max_scan = 2), "Could not find a header row")
})

test_that("header_row and paths are validated", {
  path <- .hd_filled(5)
  expect_error(detect_header_row("no-such-file.xlsx"), "existing Excel file")
  expect_error(detect_header_row(path, max_scan = 0), "max_scan")
  expect_error(read_plot_sheet(path, header_row = 0), "header_row")
  expect_error(read_plot_sheet(path, header_row = c(1, 2)), "header_row")
  expect_error(read_plot_sheet(path, header_row = "2"), "header_row")
  expect_error(read_plot_sheet(path, header_row = 9999), "beyond the last row")
  expect_error(forplot_balance(path, header_row = 1.5), "header_row")
})

test_that("bilingual labels match on either side, ignoring case and accents", {
  expect_identical(
    .match_header_cells(c("FAMILIA BOTANICA", "Wy | D", "  isi cientifico ", "Rio", NA, "X (m)")),
    c("Family", "D", "Original determination", NA, NA, "X")
  )
  ch <- .canonical_header(c("Tag", "DAP", "D", "Outra | Other", "", "dbh"))
  expect_identical(ch$column, c("New Tag No", "D", "D_2", "Outra", "NA_col_5", "D_3"))
  expect_identical(ch$duplicate_of, c(NA, NA, "D", NA, NA, "D"))
})


# Metadata parsing -------------------------------------------------------------

test_that("metadata labels are accepted in English, Portuguese and Panara", {
  m <- .parse_plot_metadata(c("Codigo da parcela: AB-1", "Nome da parcela: Mata",
                              "Data: 03/02/2021", "Equipe: Ana; Bia & Caio",
                              "Censo: 2", "Obs | Note: chuva", "sem dois pontos",
                              "Team | Vazio:", NA))
  expect_identical(m$plot_code, "AB-1")
  expect_identical(m$plot_name, "Mata")
  expect_identical(m$date, as.Date("2021-02-03"))
  expect_identical(m$team, c("Ana", "Bia", "Caio"))
  expect_identical(m$census_no, "2")
  expect_identical(m$other, c(Obs = "chuva"))

  p <- .parse_plot_metadata(c("Kypa kukari: X1", "Haka: 2020-12-31", "Saperate: Um"))
  expect_identical(p$plot_code, "X1")
  expect_identical(p$date, as.Date("2020-12-31"))
  expect_identical(p$team, "Um")

  empty <- .parse_plot_metadata(character())
  expect_true(is.na(empty$plot_code))
  expect_match(.header_message(1L, empty), "no metadata rows")
  expect_match(.header_message(3L, m), "6 metadata fields read", fixed = TRUE)
})

test_that("dates are parsed day first, then in other orders", {
  expect_identical(.parse_meta_date("26/11/2014"), as.Date("2014-11-26"))
  expect_identical(.parse_meta_date("03/04/2020"), as.Date("2020-04-03"))
  expect_identical(.parse_meta_date("11/26/2014"), as.Date("2014-11-26"))
  expect_identical(.parse_meta_date("2014-11-26"), as.Date("2014-11-26"))
  expect_identical(.parse_meta_date("26.11.2014"), as.Date("2014-11-26"))
  expect_identical(.parse_meta_date("26/11/14"), as.Date("2014-11-26"))
  expect_identical(.parse_meta_date("41969"), as.Date("2014-11-26"))
  expect_true(is.na(.parse_meta_date("soon")))
  expect_true(is.na(.parse_meta_date(NA)))
  expect_identical(.split_team(" Ana ,Bia & Caio;  "), c("Ana", "Bia", "Caio"))
  expect_identical(.split_team(NA), character())
})


# Herbarium labels -------------------------------------------------------------

test_that("forplot_sheets() reads a bilingual sheet from its real header", {
  cols <- c("family", "genus", "sp1", "number", "collector", "addcoll", "colldd",
            "collmm", "collyy", "country", "majorarea", "minorarea", "uc", "projeto",
            "locnotes", "latitude", "longitude", "altprof", "notes")
  td <- withr::local_tempdir()
  herb <- file.path(td, "jabot_template.xlsx")
  openxlsx::write.xlsx(
    as.data.frame(stats::setNames(replicate(length(cols), character(0), simplify = FALSE), cols)),
    herb
  )
  path <- .hd_filled(60)
  run <- function(...) suppressMessages(forplot_sheets(
    path, herb, input_type = "forestplots", majorarea = "MT",
    minorarea = "Teste", lat = -10, long = -53, alt = 300,
    validate_locality = FALSE, dir = td, ...
  ))
  out <- run()
  # one label per stem with the Collected column filled
  expect_identical(nrow(out), 40L)
  expect_identical(unique(as.character(out$collyy)), "2014")
  expect_identical(unique(as.character(out$collmm)), "11")
  expect_identical(unique(as.character(out$colldd)), "26")
  expect_identical(unique(out$collector), "Ana Souza")
  expect_true(all(out$genus %in% c("Inga", "Eugenia", "Attalea", "Protium")))
  expect_true(grepl("cm DBH", out$notes[1], fixed = TRUE))
  expect_identical(run(header_row = 2), out)
})


# Column names for mapping -----------------------------------------------------

test_that("read_plot_header() lists the columns of the detected header, not the metadata", {
  # four metadata rows "Key | Translation: value" before the header; an empty
  # header cell over data, a repeated header, and a trailing empty column
  rows <- list(
    c("Plotcode | Codigo: ZZ01"),
    c("Plot Name | Nome: Mata de Teste"),
    c("Date | Data: 26/11/2014"),
    c("Team | Equipe: Ana Souza, Bia Lima & Caio Reis"),
    c("Placa | Tag", "Subparcela | T1", "X", "Y", "", "Familia | Family",
      "Especie | Species", "DAP | D", "DAP | D", "Coletado | Collected", ""),
    c("1", "1", "2", "3", "obs a", "Fabaceae", "Inga alba", "120", "999", "yes", ""),
    c("2", "2", "5", "6", "", "Myrtaceae", "Eugenia sp.", "150", "888", "", "")
  )
  path <- .hd_write_raw(rows)
  cols <- read_plot_header(path)

  expect_identical(attr(cols, "header_row"), 5L)
  expect_identical(cols$position, 1:10)
  expect_identical(cols$name, c("Placa | Tag", "Subparcela | T1", "X", "Y", "NA_col_5",
                                "Familia | Family", "Especie | Species", "DAP | D",
                                "DAP | D.1", "Coletado | Collected"))
  expect_identical(cols$header[5], "")
  expect_identical(cols$canonical[c(1, 2, 5, 8, 9)], c("New Tag No", "T1", NA, "D", NA))
  expect_true(all(cols$has_data))
  expect_false(any(grepl("ZZ01|Mata de Teste|26/11/2014|Ana Souza", cols$name)))

  # forcing the row gives that row, whatever it holds
  forced <- read_plot_header(path, header_row = 1)
  expect_identical(attr(forced, "header_row"), 1L)
  expect_identical(forced$name[1], "Plotcode | Codigo: ZZ01")

  # every listed name is exactly a column the readers know: mapping by it works
  for (type in c("field_sheet", "field_sheet_ti")) {
    for (nm in cols$name) {
      cm <- .build_col_map(colname_voucher = nm)
      h <- .harmonize_plot_input(path, type, verbose = FALSE, col_map = cm)
      expect_identical(nrow(h$fp_sheet), 2L, info = paste(type, nm))
    }
    # the repeated header is reachable under its own name
    # ... and so is the column whose header cell is empty
    cm <- .build_col_map(colname_dbh = "DAP | D.1", colname_voucher = "NA_col_5")
    h <- .harmonize_plot_input(path, type, verbose = FALSE, col_map = cm)
    expect_equal(as.numeric(h$fp_sheet$D), c(999, 888), info = type)
    expect_identical(h$fp_sheet$Voucher, c("obs a", NA), info = type)
  }
})

test_that("read_plot_header() keeps row 1 for a plain sheet and reads other sheets", {
  path <- withr::local_tempfile(fileext = ".xlsx")
  plain <- data.frame(
    `New Tag No` = c("1", "2"), T1 = c(1, 1), X = c(1, 2), Y = c(3, 4),
    Family = c("Fabaceae", "Myrtaceae"), `Original determination` = c("Inga alba", "Eugenia sp."),
    D = c(120, 150), Voucher = c("AB 1", NA), Collected = c("yes", NA),
    check.names = FALSE, stringsAsFactors = FALSE
  )
  other <- data.frame(Placa = "9", Subparcela = 3, X = 1, Y = 1, Familia = "Fabaceae",
                      Especie = "Inga alba", DAP = 100, stringsAsFactors = FALSE)
  writexl::write_xlsx(list(Data = plain, Outra = other), path)

  cols <- read_plot_header(path)
  expect_identical(attr(cols, "header_row"), 1L)
  expect_identical(cols$name, names(plain))
  expect_identical(cols$canonical, names(plain))

  second <- read_plot_header(path, sheet = 2)
  expect_identical(second$name, names(other))
  expect_identical(read_plot_header(path, sheet = "Outra"), second)

  # the blank template: header names, no data yet
  tpl <- read_plot_header(.hd_template())
  expect_identical(attr(tpl, "header_row"), 2L)
  expect_identical(nrow(tpl), 20L)
  expect_identical(tpl$name[c(10, 19)], c("D | Wy", "D"))
  expect_false(any(tpl$has_data))
  expect_error(read_plot_header("no-such-file.xlsx"), "existing Excel file")
})

test_that("read_plot_header() and read_plot_sheet() reject an empty worksheet", {
  path <- withr::local_tempfile(fileext = ".xlsx")
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "Vazia")
  openxlsx::saveWorkbook(wb, path)
  expect_error(read_plot_header(path), "worksheet is empty")
  expect_error(read_plot_sheet(path), "worksheet is empty")
})
