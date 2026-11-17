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

test_that("strings exist in every language, Panara included", {
  en <- .field_report_strings("en")
  for (lg in c("pt", "es", "fr", "ma", "pa")) {
    x <- .field_report_strings(lg)
    expect_identical(names(x), names(en))
    expect_false(any(is.na(x) | !nzchar(x)))
  }
  expect_identical(names(.field_report_strings("pa")), names(.field_report_strings("pt")))
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

test_that("phytosociology section reports diversity, charts and full tables", {
  coords <- .rus01_coords()
  lay <- .field_report_layout(coords, "field_sheet", 10, 100, 100, TRUE)
  phy <- .calculate_phytosociological_metrics(coords, plot_size_ha = 1, subplot_size_m = 10)
  tr <- .field_report_strings("en")
  html <- .fr_phyto_html(phy, lay$trees$dbh, 1, tr)

  for (id in c("phyto", "phyto-2", "phyto-tables")) {
    expect_true(grepl(sprintf(' id="%s"', id), html, fixed = TRUE), info = id)
  }
  # one row per family and per taxon in the tables
  fam_tbl <- regmatches(html, regexpr("<caption>Table 1.*?</table>", html))
  sp_tbl <- regmatches(html, regexpr("<caption>Table 2.*?</table>", html))
  expect_equal(lengths(regmatches(fam_tbl, gregexpr("<tr><td>", fam_tbl))), nrow(phy$family_metrics))
  expect_equal(lengths(regmatches(sp_tbl, gregexpr("<tr><td>", sp_tbl))), nrow(phy$species_metrics))
  # IVI sums to 300 over all taxa
  expect_equal(sum(phy$species_metrics$importance_value), 300, tolerance = 0.5)
  expect_equal(sum(phy$family_metrics$importance_value), 300, tolerance = 0.5)
  # a name used in several families is one taxon per family
  expect_gt(sum(phy$species_metrics$species == "Indet indet"), 1)
  expect_true(grepl("Indet indet (Myrtaceae)", html, fixed = TRUE))
  # density per hectare and the three components in the legend
  expect_true(grepl(">936<", html, fixed = TRUE))
  expect_equal(lengths(regmatches(html, gregexpr('class="k-s[123]"', html))), 6L)  # two IVI figures
  expect_true(grepl('class="chart dbh"', html, fixed = TRUE))
})

test_that("phytosociology helpers handle edge cases", {
  expect_equal(.fr_nice_step(0), 1)
  expect_equal(.fr_nice_step(48), 10)
  expect_equal(.fr_nice_step(7), 2)
  expect_identical(.fr_trunc("abc", 30), "abc")
  expect_equal(nchar(.fr_trunc(strrep("a", 40), 30)), 30L)
  expect_null(.fr_dbh_classes(c(NA, -1, 0)))
  cl <- .fr_dbh_classes(c(10.2, 12, 25, 31, 99))
  expect_equal(sum(cl$n), 5L)
  expect_equal(cl$hi - cl$lo, rep(10, nrow(cl)))
  expect_identical(.fr_dbh_svg(numeric(), .field_report_strings("en")), "")
  expect_identical(.taxon_label(c("A b", "A b", "C d"), c("F1", "F2", "F1")), c("A b (F1)", "A b (F2)", "C d"))
  expect_identical(.taxon_label(c("x", "y")), c("x", "y"))

  # no diameters: no basal area, no relative dominance, a note in the report
  coords <- .rus01_coords()
  coords$D <- NA
  phy <- suppressMessages(.calculate_phytosociological_metrics(coords, 1, 10))
  html <- .fr_phyto_html(phy, rep(NA_real_, nrow(coords)), 1, .field_report_strings("pt"))
  expect_true(grepl("Sem di\u00e2metros v\u00e1lidos", html, fixed = TRUE))
  expect_false(grepl('class="chart dbh"', html, fixed = TRUE))
  expect_identical(.fr_phyto_html(list(species_metrics = data.frame()), numeric(), 1, .field_report_strings("en")), "")
})

test_that("forplot_balance puts the phytosociology section before the appendix and exports the charts", {
  fp <- withr::local_tempfile(fileext = ".xlsx")
  openxlsx::write.xlsx(RUS01, fp)
  out <- withr::local_tempdir()
  res <- suppressMessages(suppressWarnings(forplot_balance(
    fp, plot_size = 1, subplot_size = 10, render_pdf = FALSE, write_xlsx = FALSE,
    dir = out, filename = "rus", verbose = FALSE)))
  html <- paste(readLines(res$html_report, warn = FALSE, encoding = "UTF-8"), collapse = "\n")
  expect_lt(regexpr(' id="phyto"', html), regexpr(' id="checklist"', html))
  expect_gt(regexpr(' id="phyto"', html), regexpr(' id="s-100"', html))
  for (f in c("_families_ivi.png", "_species_abundance.png", "_species_ivi.png", "_dbh_classes.png")) {
    expect_true(file.exists(file.path(res$output_dir, paste0("rus", f))), info = f)
  }
})

test_that(".find_chrome honours the FORPLOTR_CHROME environment variable", {
  fake <- withr::local_tempfile(fileext = ".exe")
  writeLines("", fake)
  withr::local_envvar(FORPLOTR_CHROME = fake, CHROMOTE_CHROME = "",
                      PAGEDOWN_CHROME = "")
  expect_identical(.find_chrome(), fake)

  withr::local_envvar(FORPLOTR_CHROME = file.path(tempdir(), "no-such-browser"))
  found <- .find_chrome()
  expect_true(is.null(found) || file.exists(found))
})

test_that("a browser that fails to print gives a warning and no PDF", {
  html <- withr::local_tempfile(fileext = ".html")
  writeLines("<html><body>x</body></html>", html)
  pdf <- file.path(withr::local_tempdir(), "out.pdf")
  writeLines("stale", pdf)
  # Rscript stands in for a browser: it rejects the arguments and exits
  rscript <- file.path(R.home("bin"),
                       if (.Platform$OS.type == "windows") "Rscript.exe" else "Rscript")
  local_mocked_bindings(.find_chrome = function() rscript)

  expect_message(
    expect_warning(out <- .html_to_pdf(html, pdf, verbose = TRUE, timeout = 60),
                   "The PDF report could not be printed"),
    "Printing the PDF report with Rscript"
  )
  expect_null(out)
  expect_false(file.exists(pdf))
})

test_that("missing report assets raise an error", {
  expect_error(.fr_read_asset("does-not-exist.css"), "Report asset not found")
})


# Logos ------------------------------------------------------------------------

test_that("report logos resolve presets, names and files", {
  def <- .resolve_report_logos(NULL, "field_sheet")
  expect_identical(nrow(def), 2L)
  expect_identical(def$alt[1], "forplotR")
  expect_match(def$alt[2], "JBRJ", fixed = TRUE)
  expect_true(all(startsWith(def$src, "data:image/png;base64,")))

  # the Indigenous Land sheet suggests the Panara preset; naming logos turns it off
  ti <- .resolve_report_logos(NULL, "field_sheet_ti")
  expect_identical(nrow(ti), 4L)
  expect_match(ti$alt[3], "Conserva", fixed = TRUE)
  expect_match(ti$alt[4], "Iaki", fixed = TRUE)
  expect_identical(.resolve_report_logos(c("forplotR", "jbrj"), "field_sheet_ti"), def)
  expect_identical(.resolve_report_logos("panara", "field_sheet"), ti)
  expect_identical(.resolve_report_logos("default", "field_sheet_ti"), def)

  # no logos
  for (none in list(FALSE, "none", character(), "")) {
    expect_identical(nrow(.resolve_report_logos(none)), 0L)
  }
  expect_identical(.fr_logos_html(.resolve_report_logos(FALSE)), "")
  expect_identical(.fr_logos_html(NULL), "")

  # order is kept, repeats are dropped, custom files are accepted
  svg <- withr::local_tempfile(fileext = ".svg")
  writeLines('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2 1"><rect width="2" height="1"/></svg>', svg)
  mix <- .resolve_report_logos(c("JBRJ", svg, "default"))
  expect_identical(nrow(mix), 3L)
  expect_match(mix$alt[1], "JBRJ", fixed = TRUE)
  expect_true(startsWith(mix$src[2], "data:image/svg+xml;base64,"))
  expect_identical(mix$alt[3], "forplotR")

  expect_error(.resolve_report_logos("nope"), "Unknown logo 'nope'")
  expect_error(.resolve_report_logos(1), "`logos` must be")
  txt <- withr::local_tempfile(fileext = ".txt")
  writeLines("x", txt)
  expect_error(.resolve_report_logos(txt), "must be a PNG, SVG")
})

test_that("the report cover carries the logos", {
  lay <- .field_report_layout(.rus01_coords(), "field_sheet", 10, 100, 100, TRUE)
  meta <- list(plot_name = "Ruschi", plot_code = "RUS01", plot_census_no_fp = "1", team = "")
  html <- .field_report_html(lay, meta, "en", TRUE, plot_area_ha = 1,
                             logos = .resolve_report_logos("panara"))
  # cut by position: the strip holds base64 images, too long for a regex match
  from <- regexpr('<div class="logos">', html, fixed = TRUE)
  to <- regexpr('<header class="doc-head">', html, fixed = TRUE)
  expect_gt(from, 0)
  strip <- substr(html, from, to - 1L)
  expect_identical(lengths(regmatches(strip, gregexpr("<img ", strip, fixed = TRUE))), 4L)
  expect_true(grepl('alt="forplotR"', strip, fixed = TRUE))
  # the strip comes before the title, on the cover page
  expect_lt(regexpr('<div class="logos">', html, fixed = TRUE),
            regexpr('<header class="doc-head">', html, fixed = TRUE))
  none <- .field_report_html(lay, meta, "en", TRUE, plot_area_ha = 1)
  expect_false(grepl('<div class="logos">', none, fixed = TRUE))
  expect_error(forplot_balance("x.xlsx", logos = "nope"), "Unknown logo")
})

test_that("MONITORA and ForestPlots logos work as a preset or as a single logo", {
  def <- .resolve_report_logos("default")
  mon <- .resolve_report_logos("monitora")
  expect_identical(nrow(mon), 3L)
  expect_identical(mon[1:2, ], def)
  expect_match(mon$alt[3], "MONITORA", fixed = TRUE)
  fp <- .resolve_report_logos("ForestPlots")
  expect_identical(fp$alt, c(def$alt, "ForestPlots.net"))

  # in a list they are just that logo, in the order given
  one <- .resolve_report_logos(c("forestplots", "forplotR"))
  expect_identical(one$alt, c("ForestPlots.net", "forplotR"))
  all6 <- .resolve_report_logos(c("forplotR", "jbrj", "ci", "iakio", "monitora", "forestplots"))
  expect_identical(nrow(all6), 6L)
  expect_true(all(startsWith(all6$src, "data:image/png;base64,")))
  expect_true(all(nzchar(all6$url)))

  # `auto` is what NULL stands for when the input type does not decide
  expect_identical(.resolve_report_logos(NULL, "monitora", auto = "forplotR")$alt, "forplotR")
  expect_identical(nrow(.resolve_report_logos(NULL, "field_sheet_ti", auto = "forplotR")), 4L)
})


# Panara -----------------------------------------------------------------------

test_that("the Panara report takes its labels from the Panara dictionary", {
  pa <- .field_report_strings("pa")
  pt <- .field_report_strings("pt")
  expect_identical(names(pa), names(pt))

  # labels the dictionary covers come from it, word for word
  d <- function(key) trimws(paste(unlist(dict$pa[dict$key == key][[1]]), collapse = " "))
  expect_identical(pa[["team"]], d("team"))
  expect_identical(pa[["with_voucher"]], d("collected"))
  expect_identical(pa[["pending"]], d("uncollected"))
  expect_identical(pa[["palms"]], d("Palms"))
  expect_identical(pa[["tag"]], d("hover_tag"))
  expect_identical(pa[["subplot"]], d("subplot"))
  expect_identical(pa[["doc_title"]], d("Full Plot Report"))
  expect_identical(pa[["date"]], "Haka")
  expect_false(any(c("team", "with_voucher", "pending", "palms", "tag", "species",
                     "family", "individuals", "subplot") %in% names(pa)[pa == pt]))
  expect_gte(sum(pa != pt), 40L)
  expect_identical(pa[["fig_fam"]], d("Most Common Families"))
  expect_identical(pa[["tab_spp"]], d("Species Metrics"))
  expect_identical(sprintf(pa[["n_pending"]], 3L), paste("3", tolower(d("uncollected"))))

  # what the dictionary does not cover stays in Portuguese, never English
  expect_identical(pa[["fn_palms"]], pt[["fn_palms"]])
  expect_identical(pa[["route_lede"]], pt[["route_lede"]])
  # format strings keep their placeholders
  expect_identical(sprintf(pa[["subplots_of"]], 100L, "10", "10"),
                   paste0("100 ", tolower(d("Subplots")), " \u00b7 10 \u00d7 10 m"))
  expect_true(grepl("%d", pa[["in_subplots"]], fixed = TRUE))

  # dictionary entries that are not really translated are not used
  untranslated <- vapply(seq_len(nrow(dict)), function(i) {
    identical(unlist(dict$pa[i]), unlist(dict$pt[i]))
  }, logical(1))
  expect_false(any(.field_report_strings_pa() %in% unlist(dict$pa[untranslated])))
  # no private-use or replacement characters in the Panara column
  expect_false(any(grepl("[\ue000-\uf8ff\ufffd]", unlist(dict$pa))))

  lay <- .field_report_layout(.rus01_coords(), "field_sheet", 10, 100, 100, TRUE)
  meta <- list(plot_name = "Ruschi", plot_code = "RUS01", plot_census_no_fp = "1",
               team = "A, B", census_date = as.Date("2014-11-26"))
  html <- .field_report_html(lay, meta, "pa", TRUE, plot_area_ha = 1)
  expect_true(grepl(paste0(d("team"), ": A, B"), html, fixed = TRUE))
  expect_true(grepl("Haka: 26/11/2014", html, fixed = TRUE))
  expect_true(grepl(d("collected"), html, fixed = TRUE))
  # the label of the figure is Panara; the footnote that explains it is a
  # sentence, which the dictionary does not have
  expect_false(grepl("<span>Com voucher</span>", html, fixed = TRUE))
})
