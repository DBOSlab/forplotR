# Field report of forplot_balance() (R/report_field.R)

.rus01_coords <- function() {
  fp <- withr::local_tempfile(fileext = ".xlsx", .local_envir = parent.frame())
  openxlsx::write.xlsx(RUS01, fp)
  ld <- suppressMessages(.harmonize_plot_input(fp, "field_sheet", verbose = FALSE))
  suppressWarnings(.compute_global_coordinates(.clean_fp_data(ld$fp_sheet), 10, 100, 100))
}

.ids_unique <- function(html) {
  ids <- regmatches(html, gregexpr(' id="[^"]+"', html))[[1]]
  !anyDuplicated(ids)
}

.anchors_resolve <- function(html) {
  ids <- sub(' id="([^"]+)"', "\\1", regmatches(html, gregexpr(' id="[^"]+"', html))[[1]])
  refs <- sub('href="#([^"]+)"', "\\1", regmatches(html, gregexpr('href="#[^"]+"', html))[[1]])
  setdiff(unique(refs), ids)
}

test_that("strings exist in every language, with Panara falling back to Portuguese", {
  en <- .field_report_strings("en")
  for (lg in c("pt", "es", "fr", "ma", "pa")) {
    x <- .field_report_strings(lg)
    expect_identical(names(x), names(en))
    expect_false(any(is.na(x) | !nzchar(x)))
  }
  expect_identical(.field_report_strings("pa"), .field_report_strings("pt"))
  expect_identical(.field_report_strings("xx"), en)
  # placeholders survive translation
  for (lg in c("pt", "es", "fr", "ma")) {
    x <- .field_report_strings(lg)
    expect_identical(lengths(regmatches(x, gregexpr("%[ds]", x))),
                     lengths(regmatches(en, gregexpr("%[ds]", en))))
  }
})

test_that("layout counts stems, statuses and subplots for RUS01", {
  lay <- .field_report_layout(.rus01_coords(), "field_sheet", 10, 100, 100, TRUE)
  expect_equal(nrow(lay$cells), 100L)
  expect_equal(nrow(lay$trees), 936L)
  expect_equal(sum(lay$trees$st == "ok"), 688L)
  expect_equal(sum(lay$trees$st == "todo"), 11L)
  expect_equal(sum(lay$trees$st == "palm"), 237L)
  expect_equal(sum(lay$cells$todo > 0), 8L)
  expect_equal(sum(lay$cells$n), 936)
  # serpentine columns: subplot 11 sits at the top of the second column
  expect_equal(unlist(lay$cells[lay$cells$id == "11", c("x0", "y0")], use.names = FALSE), c(10, 90))

  lay2 <- .field_report_layout(.rus01_coords(), "field_sheet", 10, 100, 100, FALSE)
  expect_equal(sum(lay2$trees$st == "palm"), 0L)
  expect_equal(sum(lay2$trees$st %in% c("ok", "todo")), 936L)
})

test_that("HTML has every section, unique ids and no broken internal link", {
  lay <- .field_report_layout(.rus01_coords(), "field_sheet", 10, 100, 100, TRUE)
  meta <- list(plot_name = "REBIO Augusto Ruschi", plot_code = "RUS01", plot_census_no_fp = "1", team = NULL)
  html <- .field_report_html(lay, meta, "en", TRUE, plot_area_ha = 1)

  expect_true(grepl('<html lang="en">', html, fixed = TRUE))
  for (id in c("summary", "map-all", "map-ok", "map-todo", "map-palm", "route", "s-1", "s-100", "checklist")) {
    expect_true(grepl(sprintf(' id="%s"', id), html, fixed = TRUE), info = id)
  }
  expect_true(.ids_unique(html))
  expect_length(.anchors_resolve(html), 0)
  expect_true(grepl("688 / 699", html, fixed = TRUE))
  expect_false(grepl("http://|https://", html))           # nothing loaded from the network
  # every stem is labelled with its tag on the whole-plot map
  all_map <- regmatches(html, regexpr('<section class="page mapp" id="map-all">.*?</section>', html))
  placed <- sum(is.finite(lay$trees$gx) & is.finite(lay$trees$gy))  # stems without X/Y stay in the tables only
  expect_equal(placed, 930L)
  expect_equal(lengths(regmatches(all_map, gregexpr("<text x=", all_map))), placed)

  html2 <- .field_report_html(lay = .field_report_layout(.rus01_coords(), "field_sheet", 10, 100, 100, FALSE),
                              meta, "pt", FALSE)
  expect_false(grepl(' id="map-palm"', html2, fixed = TRUE))
  expect_true(grepl('<html lang="pt-BR">', html2, fixed = TRUE))
  expect_true(grepl("Pendentes de coleta", html2, fixed = TRUE))
})

test_that("MONITORA plots are laid out as four arms of ten subplots", {
  fp <- withr::local_tempfile(fileext = ".xlsx")
  openxlsx::write.xlsx(TUM, fp)
  ld <- suppressMessages(.harmonize_plot_input(fp, "monitora", verbose = FALSE))
  coords <- .compute_monitora_geometry(.clean_fp_data(ld$fp_sheet), keep_only_cell = FALSE)
  lay <- .field_report_layout(coords, "monitora", 10, NULL, NULL, TRUE)

  expect_equal(nrow(lay$cells), 40L)
  expect_setequal(lay$cells$id, paste0(rep(c("N", "S", "L", "O"), each = 10), 1:10))
  expect_equal(sum(lay$cells$n), nrow(lay$trees))
  # arms do not overlap, in the real and in the schematic position
  overlap <- function(x0, y0) {
    any(vapply(seq_along(x0), function(i) any(abs(x0[-i] - x0[i]) < 10 & abs(y0[-i] - y0[i]) < 10), logical(1)))
  }
  expect_false(overlap(lay$cells$x0, lay$cells$y0))
  expect_false(overlap(lay$cells$sx0, lay$cells$sy0))

  html <- .field_report_html(lay, list(plot_name = "Tumucumaque", plot_code = "EA-1"), "en", TRUE)
  expect_true(grepl(' id="s-N1"', html, fixed = TRUE))
  expect_true(grepl(' id="s-O10"', html, fixed = TRUE))
  expect_true(.ids_unique(html))
  expect_length(.anchors_resolve(html), 0)
})

test_that("labels avoid each other when there is room", {
  x <- c(10, 12, 50); y <- c(10, 10, 50); r <- c(2, 2, 2)
  lp <- .fr_place_labels(x, y, r, c("101", "102", "103"), fs = 8)
  expect_false(isTRUE(all.equal(lp[1, ], lp[2, ])))
  expect_equal(dim(lp), c(3L, 2L))
})

test_that("PDF is skipped with a warning when no browser is available", {
  html <- withr::local_tempfile(fileext = ".html")
  writeLines("<html><body>x</body></html>", html)
  local_mocked_bindings(.find_chrome = function() NULL)
  expect_warning(out <- .html_to_pdf(html, tempfile(fileext = ".pdf"), verbose = FALSE),
                 "No Chrome, Chromium or Edge")
  expect_null(out)
})

test_that("forplot_balance writes the field report without pandoc or LaTeX", {
  fp <- withr::local_tempfile(fileext = ".xlsx")
  openxlsx::write.xlsx(RUS01, fp)
  out <- withr::local_tempdir()
  local_mocked_bindings(.find_chrome = function() NULL)
  res <- suppressMessages(suppressWarnings(forplot_balance(
    fp, plot_size = 1, subplot_size = 10, render_html = TRUE, render_pdf = TRUE,
    write_xlsx = FALSE, dir = out, filename = "rus", verbose = FALSE)))
  expect_true(file.exists(res$html_report))
  expect_null(res$pdf_report)
  html <- paste(readLines(res$html_report, warn = FALSE, encoding = "UTF-8"), collapse = "\n")
  expect_true(grepl("REBIO Augusto Ruschi", html, fixed = TRUE))
  expect_lt(file.size(res$html_report), 3e6)
})
