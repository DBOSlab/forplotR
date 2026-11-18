# GigaFlor field sheets: protocol 1 (25 subplots of 20 x 20 m), X/Y in
# decimal degrees, D in cm, "Tag Number" / "Species" / "Main Stem Tag".

# Simulated plot: trees at known positions inside a 100 x 100 m grid rotated
# by `rot` degrees and placed near the Equator, written as a GigaFlor sheet.
.make_gigaflor_sheet <- function(path, rot = 20, n_per_cell = 4, seed = 1,
                                 main_stem = TRUE, extra_sheet_first = TRUE) {
  set.seed(seed)
  s <- 20; nr <- 5
  T1 <- rep(1:25, each = n_per_cell)
  orig <- .t1_cell_origin(T1, s, nr)
  lx <- stats::runif(length(T1), 2, 18)
  ly <- stats::runif(length(T1), 2, 18)
  gx <- orig[, 1] + lx - 50
  gy <- orig[, 2] + ly - 50
  a <- rot * pi / 180
  east <- gx * cos(a) - gy * sin(a)
  north <- gx * sin(a) + gy * cos(a)
  lat0 <- 1.19; lon0 <- -52.37
  phi <- lat0 * pi / 180
  m_lat <- 111132.954 - 559.822 * cos(2 * phi) + 1.175 * cos(4 * phi)
  m_lon <- 111412.84 * cos(phi) - 93.5 * cos(3 * phi)
  n <- length(T1)
  dat <- data.frame(
    `Tag Number` = seq_len(n),
    `Main Stem Tag` = NA,
    T1 = T1,
    T2 = 1 + (seq_len(n) %% 4),
    X = lon0 + east / m_lon,
    Y = lat0 + north / m_lat,
    Family = rep(c("Fabaceae", "Lauraceae", "Arecaceae", "Sapotaceae"), length.out = n),
    Species = rep(c("Inga alba", "Indet indet", "Euterpe oleracea", "Pouteria sp."), length.out = n),
    `Original identification` = "vernacular",
    D = round(stats::runif(n, 10, 120), 2),
    POM = 1.3,
    Collected = ifelse(seq_len(n) %% 3 == 0, "Yes", NA),
    `CAP (cm)` = NA,
    check.names = FALSE
  )
  if (main_stem) {
    dat$`Main Stem Tag`[c(2, 3)] <- 1   # tags 2 and 3 are stems of tree 1
  }
  meta <- c("Plotcode: GIG01", "Plot Name: Test plot", "Date: 26/11/2026", "Team: A & B")
  top <- as.data.frame(matrix(NA, nrow = 2, ncol = ncol(dat)))
  top[1, 1:4] <- meta
  top[2, ] <- names(dat)
  body <- as.data.frame(lapply(dat, as.character), stringsAsFactors = FALSE)
  names(top) <- names(body) <- paste0("c", seq_len(ncol(dat)))
  sheet <- rbind(top, body)
  wb <- openxlsx::createWorkbook()
  if (extra_sheet_first) {
    openxlsx::addWorksheet(wb, "README")
    openxlsx::writeData(wb, "README", "Instructions", colNames = FALSE)
  }
  openxlsx::addWorksheet(wb, "Data")
  openxlsx::writeData(wb, "Data", sheet, colNames = FALSE)
  # numeric cells as numbers
  for (j in which(vapply(dat, is.numeric, logical(1)))) {
    openxlsx::writeData(wb, "Data", dat[[j]], startCol = j, startRow = 3, colNames = FALSE)
  }
  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
  list(path = path, local = cbind(lx, ly), dat = dat)
}

test_that("GigaFlor sheet is read from 'Data' with its own names and units", {
  skip_if_not_installed("openxlsx")
  f <- tempfile(fileext = ".xlsx")
  sim <- .make_gigaflor_sheet(f)

  fl <- .harmonize_plot_input(f, "gigaflor", verbose = FALSE)
  d <- fl$fp_sheet

  expect_equal(fl$plot_code, "GIG01")
  # stems 2 and 3 are consolidated into tree 1
  expect_equal(nrow(d), nrow(sim$dat) - 2L)
  expect_true(all(c("New Tag No", "Original determination", "Longitude", "Latitude") %in% names(d)))
  expect_setequal(unique(d$`Original determination`),
                  c("Inga alba", "Indet indet", "Euterpe oleracea", "Pouteria sp."))
  # D is cm in the sheet and mm in the package
  i <- match("10", d$`New Tag No`)
  expect_equal(d$D[i], sim$dat$D[10] * 10)
  # coordinates are kept in decimal degrees
  expect_equal(d$Latitude[i], sim$dat$Y[10], tolerance = 1e-9)
})

test_that("D falls back on CAP and empty template rows are dropped", {
  df <- data.frame(`New Tag No` = c("1", "2", NA), T1 = c(1, 1, NA),
                   X = c(-52, -52, NA), Y = c(1, 1, NA),
                   Family = c("Fabaceae", "Fabaceae", NA),
                   `Original determination` = c("A b", "A b", NA),
                   D = c(NA, "20", NA), CAP = c("31.4159", NA, NA),
                   check.names = FALSE, stringsAsFactors = FALSE)
  out <- .gigaflor_prepare(df, verbose = FALSE)
  expect_equal(nrow(out), 2L)
  expect_equal(out$D, c(100, 200))
})

test_that("GigaFlor decimal-degree coordinates are fitted onto the T1 grid", {
  skip_if_not_installed("openxlsx")
  f <- tempfile(fileext = ".xlsx")
  sim <- .make_gigaflor_sheet(f, rot = 35, main_stem = FALSE)
  fl <- .harmonize_plot_input(f, "gigaflor", verbose = FALSE)
  out <- .gigaflor_local_xy(fl$fp_sheet, 20, 100, 100, verbose = FALSE)

  fit <- attr(out, "gigaflor_fit")
  expect_lt(abs(fit$rotation_deg - 35), 2)
  expect_false(fit$mirrored)
  ord <- as.integer(out$`New Tag No`)
  # recovered within ~1 m of the true position inside the subplot
  expect_lt(max(abs(out$X - sim$local[ord, 1])), 1.5)
  expect_lt(max(abs(out$Y - sim$local[ord, 2])), 1.5)
  expect_true(all(out$X > 0 & out$X < 20 & out$Y > 0 & out$Y < 20))
})

test_that("GigaFlor coordinates: errors, poor fits and missing positions", {
  base <- data.frame(T1 = rep(1:25, each = 2), Longitude = NA_real_, Latitude = NA_real_)
  # no coordinates: centre of the subplot
  out <- .gigaflor_local_xy(base, 20, 100, 100, verbose = FALSE)
  expect_true(all(out$X == 10 & out$Y == 10))

  bad <- base
  bad$Longitude <- 347563; bad$Latitude <- 10131677   # UTM by mistake
  expect_error(.gigaflor_local_xy(bad, 20, 100, 100, verbose = FALSE), "decimal degrees")

  set.seed(2)
  rnd <- base
  rnd$Longitude <- -52.37 + stats::runif(50, 0, 0.0009)
  rnd$Latitude <- 1.19 + stats::runif(50, 0, 0.0009)
  expect_warning(.gigaflor_local_xy(rnd, 20, 100, 100, verbose = FALSE), "fit the T1 subplot grid poorly")
})

test_that("forplot_balance() runs on a GigaFlor sheet with 20 m subplots by default", {
  skip_if_not_installed("openxlsx")
  skip_on_cran()
  f <- tempfile(fileext = ".xlsx")
  .make_gigaflor_sheet(f)
  out <- file.path(tempdir(), "gigaflor_balance")
  res <- forplot_balance(f, input_type = "gigaflor", render_pdf = FALSE,
                         dir = out, verbose = FALSE)
  expect_true(file.exists(res$html_report))
  expect_true(file.exists(res$xlsx_report))
  summ <- openxlsx::read.xlsx(res$xlsx_report, sheet = "COLLECTION_PERCENTUAL", startRow = 3)
  # 25 subplots of 20 m + TOTAL
  expect_equal(nrow(summ), 26L)
  expect_equal(summ$total_individuals[summ$subplot == "TOTAL"], 98)
})

test_that("Collected = 'No' in a GigaFlor sheet is not a collection", {
  df <- data.frame(`New Tag No` = c("1", "2", "3"), T1 = 1, X = -52, Y = 1,
                   Family = "Fabaceae", `Original determination` = "A b", D = "20",
                   Collected = c("Yes", "No", NA),
                   check.names = FALSE, stringsAsFactors = FALSE)
  out <- .gigaflor_prepare(df, verbose = FALSE)
  expect_equal(out$Collected, c("Yes", NA, NA))
})

test_that("the GigaFlor fit is invertible and gives the plot corners", {
  skip_if_not_installed("openxlsx")
  f <- tempfile(fileext = ".xlsx")
  .make_gigaflor_sheet(f, rot = 20, main_stem = FALSE)
  fl <- .harmonize_plot_input(f, "gigaflor", verbose = FALSE)
  out <- .gigaflor_local_xy(fl$fp_sheet, 20, 100, 100, verbose = FALSE)
  fit <- attr(out, "gigaflor_fit")

  v <- .gigaflor_vertices(fit, 100, 100)
  expect_equal(nrow(v), 4L)
  # corners 100 m apart (local metres around the plot)
  m <- .lonlat_to_local_m(v$longitude, v$latitude, ref = fit$ref)
  side <- function(i, j) sqrt(sum((m[i, ] - m[j, ])^2))
  expect_equal(side(1, 2), 100, tolerance = 1e-3)
  expect_equal(side(1, 3), 100, tolerance = 1e-3)
  expect_equal(side(1, 4), 100 * sqrt(2), tolerance = 1e-3)
  # the corner (0, 0) of the grid sits where the simulation put it
  expect_equal(unname(.lonlat_to_local_m(v$longitude[1], v$latitude[1], ref = fit$ref)[1, ]),
               unname(m[1, ]))
  expect_error(.gigaflor_grid_to_lonlat(0, 0, list()), "No coordinate fit")
})

test_that("field report shows the GigaFlor X/Y as decimal degrees", {
  skip_if_not_installed("openxlsx")
  skip_on_cran()
  f <- tempfile(fileext = ".xlsx")
  sim <- .make_gigaflor_sheet(f, main_stem = FALSE)
  res <- forplot_balance(f, input_type = "gigaflor", render_pdf = FALSE, write_xlsx = FALSE,
                         dir = file.path(tempdir(), "gigaflor_report"), verbose = FALSE)
  h <- paste(readLines(res$html_report, warn = FALSE, encoding = "UTF-8"), collapse = "\n")
  expect_match(h, "X; Y<small>° (long.; lat.)</small>", fixed = TRUE)
  expect_match(h, sprintf("%.6f; %.6f", sim$dat$X[1], sim$dat$Y[1]), fixed = TRUE)
  expect_false(grepl("X; Y<small>m</small>", h, fixed = TRUE))
})

test_that("forplot_map() maps a GigaFlor sheet without vertices", {
  skip_if_not_installed("openxlsx")
  skip_if_not_installed("leaflet")
  skip_on_cran()
  f <- tempfile(fileext = ".xlsx")
  sim <- .make_gigaflor_sheet(f, main_stem = FALSE)
  out <- file.path(tempdir(), "gigaflor_map")
  expect_message(
    p <- forplot_map(f, input_type = "gigaflor", dir = out, filename = "gf", verbose = TRUE),
    "Plot corners estimated"
  )
  expect_true(file.exists(p))
  # trees keep the coordinates of the sheet
  h <- paste(readLines(p, warn = FALSE), collapse = "")
  lat1 <- sprintf("%.5f", trunc(sim$dat$Y[1] * 1e5) / 1e5)
  expect_match(h, lat1, fixed = TRUE)

  # corners marked in the survey can still be given
  fl <- .harmonize_plot_input(f, "gigaflor", verbose = FALSE)
  fit <- attr(.gigaflor_local_xy(fl$fp_sheet, 20, 100, 100, verbose = FALSE), "gigaflor_fit")
  v <- .gigaflor_vertices(fit, 100, 100)
  p2 <- forplot_map(f, input_type = "gigaflor", vertex_coords = v, dir = out,
                    filename = "gf_vertices", verbose = FALSE)
  expect_true(file.exists(p2))
  # other input types still need the corners
  expect_error(forplot_map(f, input_type = "field_sheet", dir = out, verbose = FALSE),
               "vertex_coords")
})

test_that("forplot_sheets() writes one label per collected GigaFlor tree", {
  skip_if_not_installed("openxlsx")
  f <- tempfile(fileext = ".xlsx")
  sim <- .make_gigaflor_sheet(f, main_stem = FALSE)
  cols <- c("family", "genus", "sp1", "number", "collector", "addcoll", "colldd",
            "collmm", "collyy", "country", "majorarea", "minorarea", "uc", "projeto",
            "locnotes", "latitude", "longitude", "altprof", "notes")
  herb <- tempfile(fileext = ".xlsx")
  openxlsx::write.xlsx(as.data.frame(stats::setNames(
    replicate(length(cols), character(0), simplify = FALSE), cols)), herb)

  h <- suppressMessages(forplot_sheets(f, herb, input_type = "gigaflor",
                                       majorarea = "Amapá", alt = 300,
                                       validate_locality = FALSE,
                                       dir = withr::local_tempdir()))
  coll <- which(!is.na(sim$dat$Collected))
  expect_equal(nrow(h), length(coll))
  expect_equal(h$latitude, round(sim$dat$Y[coll], 6))
  expect_equal(h$longitude, round(sim$dat$X[coll], 6))
  expect_equal(unique(h$collector), "A")
  expect_equal(unique(h$addcoll), "B")
  expect_equal(unique(paste(h$colldd, h$collmm, h$collyy, sep = "/")), "26/11/2026")
  # DBH in cm, as in the sheet, and no "x = ...m" for geographic X/Y
  expect_match(h$notes[1], paste0(format(sim$dat$D[coll[1]]), "cm DBH"), fixed = TRUE)
  expect_false(any(grepl("x = ", h$notes, fixed = TRUE)))
})

test_that("forplot_vouchers() accepts GigaFlor sheets", {
  skip_if_not_installed("openxlsx")
  f <- tempfile(fileext = ".xlsx")
  .make_gigaflor_sheet(f, main_stem = FALSE)
  img <- file.path(withr::local_tempdir(), "imgs")
  out <- forplot_vouchers(f, output_dir = img, input_type = "gigaflor", verbose = FALSE)
  expect_true(dir.exists(file.path(img, "Fabaceae")))
  expect_true("Original determination" %in% names(out))
})

test_that("GigaFlor helpers cope with few or missing coordinates", {
  skip_if_not_installed("openxlsx")
  f <- tempfile(fileext = ".xlsx")
  openxlsx::write.xlsx(data.frame(a = 1), f, sheetName = "Sheet1")
  expect_identical(.gigaflor_sheet(f), 1L)

  few <- data.frame(T1 = c(1, 2, 3), Longitude = c(-52.37, NA, NA),
                    Latitude = c(1.19, NA, NA), X = NA, Y = NA)
  expect_message(
    out <- .gigaflor_local_xy(few, 20, 100, 100, verbose = TRUE),
    "too few georeferenced stems"
  )
  expect_equal(out$X, rep(10, 3))
  expect_message(.gigaflor_local_xy(few, 20, 100, 100, verbose = TRUE),
                 "2 stem(s) without coordinates", fixed = TRUE)

  # a mirrored plot (T1 numbered the other way round) is still fitted
  s <- 20; nr <- 5
  T1 <- rep(1:25, each = 3)
  orig <- .t1_cell_origin(T1, s, nr)
  gx <- -(orig[, 1] + 10 - 50); gy <- orig[, 2] + 10 - 50
  mir <- data.frame(T1 = T1, Longitude = -52.37 + gx / 111000,
                    Latitude = 1.19 + gy / 110600, X = NA, Y = NA)
  expect_message(.gigaflor_local_xy(mir, 20, 100, 100, verbose = TRUE), "(mirrored)", fixed = TRUE)
})

test_that("forplot_sheets() explains GigaFlor sheets it cannot label", {
  skip_if_not_installed("openxlsx")
  cols <- c("family", "genus", "sp1", "number", "collector", "addcoll", "colldd",
            "collmm", "collyy", "country", "majorarea", "minorarea", "uc", "projeto",
            "locnotes", "latitude", "longitude", "altprof", "notes")
  herb <- tempfile(fileext = ".xlsx")
  openxlsx::write.xlsx(as.data.frame(stats::setNames(
    replicate(length(cols), character(0), simplify = FALSE), cols)), herb)
  edit_sheet <- function(f, fn) {
    wb <- openxlsx::loadWorkbook(f)
    fn(wb)
    openxlsx::saveWorkbook(wb, f, overwrite = TRUE)
  }
  run <- function(f) suppressMessages(forplot_sheets(f, herb, input_type = "gigaflor",
                                                     majorarea = "Amapa", alt = 300,
                                                     validate_locality = FALSE,
                                                     dir = withr::local_tempdir()))

  # no tree coordinates and no lat/long
  f1 <- tempfile(fileext = ".xlsx")
  sim <- .make_gigaflor_sheet(f1, main_stem = FALSE)
  n <- nrow(sim$dat)
  edit_sheet(f1, function(wb) openxlsx::deleteData(wb, "Data", cols = 5:6, rows = 3:(n + 2),
                                                   gridExpand = TRUE))
  expect_error(run(f1), "no X/Y coordinates")

  # no team and an unreadable date
  f2 <- tempfile(fileext = ".xlsx")
  .make_gigaflor_sheet(f2, main_stem = FALSE)
  edit_sheet(f2, function(wb) {
    openxlsx::writeData(wb, "Data", "Date: soon", startCol = 3, startRow = 1, colNames = FALSE)
    openxlsx::writeData(wb, "Data", "Team:", startCol = 4, startRow = 1, colNames = FALSE)
  })
  w <- character()
  h <- withCallingHandlers(run(f2), warning = function(cnd) {
    w <<- c(w, conditionMessage(cnd)); invokeRestart("muffleWarning")
  })
  expect_true(any(grepl("collector metadata", w)))
  expect_true(any(grepl("collection date", w)))
  expect_true(all(is.na(h$colldd)))

  # nothing collected
  f3 <- tempfile(fileext = ".xlsx")
  .make_gigaflor_sheet(f3, main_stem = FALSE)
  edit_sheet(f3, function(wb) openxlsx::deleteData(wb, "Data", cols = 12, rows = 3:(n + 2),
                                                   gridExpand = TRUE))
  expect_error(run(f3), "No collected specimens")
})
