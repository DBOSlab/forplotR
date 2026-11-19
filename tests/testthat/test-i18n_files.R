# Output file names (one dictionary, one helper) and the species charts.

test_that(".out_name builds '<prefix>_<name>.<ext>' in the language of the call", {
  expect_identical(.out_name("report", "pt", "RUS01", "html"), "RUS01_relatorio_completo.html")
  expect_identical(.out_name("report", "en", "RUS01", "pdf"), "RUS01_full_report.pdf")
  expect_identical(.out_name("balance_xlsx", "es", "p", "xlsx"), "p_balance_colecta.xlsx")
  expect_identical(.out_name("chart_species_ivi", "fr"), "especes_ivi")
  expect_identical(.out_name("default_prefix", "pt"), "parcela")
  expect_error(.out_name("nope", "en"), "Unknown output file key")
})

test_that("Panara uses the Portuguese names and Mandarin the English ones", {
  keys <- .file_name_dict()$key
  pt <- vapply(keys, .out_name, "", language = "pt")
  expect_identical(vapply(keys, .out_name, "", language = "pa"), pt)
  expect_identical(vapply(keys, .out_name, "", language = "ma"),
                   vapply(keys, .out_name, "", language = "en"))
})

test_that("file names are snake_case ASCII and unique within each language", {
  d <- .file_name_dict()
  d <- d[d$key != "tree", ]  # "Tree" is a word inside voucher folder names
  for (lg in c("en", "pt", "es", "fr")) {
    expect_true(all(grepl("^[a-z0-9]+(_[a-z0-9]+)*$", d[[lg]])), info = lg)
    expect_false(anyDuplicated(d[[lg]]) > 0, info = lg)
  }
})

test_that(".vernacular_names prefers 'Local name' over 'Vernacular name'", {
  x <- data.frame(`Local name` = c("A", "", NA), `Vernacular name` = c("x", "y", NA),
                  check.names = FALSE, stringsAsFactors = FALSE)
  expect_identical(.vernacular_names(x), c("A", "y", NA))
  expect_identical(.vernacular_names(data.frame(a = 1:2)), c(NA_character_, NA_character_))
})

test_that(".taxon_label_parts splits 'vernacular | scientific' and keeps repeated names apart", {
  p <- .taxon_label_parts(c("Kupa | Inga edulis", "Indet indet", "Indet indet", "Inga alba"),
                          c("Fabaceae", "Myrtaceae", "Rubiaceae", "Fabaceae"),
                          c(NA, NA, NA, "Nome"))
  expect_identical(p$sci, c("Inga edulis", "Indet indet (Myrtaceae)", "Indet indet (Rubiaceae)", "Inga alba"))
  expect_identical(p$vern, c("Kupa", NA, NA, "Nome"))
})

.panara_like <- function() {
  data.frame(
    `New Tag No` = as.character(1:6),
    T1 = c(1, 2, 3, 4, 4, 5),
    Family = c("Fabaceae", "Rubiaceae", "Fabaceae", "Arecaceae", "Fabaceae", "Myrtaceae"),
    `Original determination` = c("Inga alba", "Coffea arabica", "Inga alba", "Attalea sp", "Inga alba", "Indet indet"),
    `Local name` = c("Nome local", NA, "Nome local", NA, "", NA),
    D = c(100, 200, 300, 400, 150, 120),
    Collected = c("x", NA, NA, NA, "y", NA),
    check.names = FALSE, stringsAsFactors = FALSE
  )
}

test_that("species metrics carry the most frequent vernacular name", {
  phy <- suppressMessages(.calculate_phytosociological_metrics(.panara_like(), 1, 10))
  sp <- phy$species_metrics
  expect_true("vernacular" %in% names(sp))
  expect_identical(sp$vernacular[sp$species == "Inga alba"], "Nome local")
  expect_true(is.na(sp$vernacular[sp$species == "Coffea arabica"]))
})

test_that("the species charts label bars with their values and leave room for them", {
  out <- suppressMessages(.prepare_report_dashboard(.panara_like(), input_type = "field_sheet", language = "pt"))
  ab <- ggplot2::ggplot_build(out$species_abundance_plot)
  txt <- ab$data[[2]]
  sp <- out$phytosoc$species_metrics
  expect_setequal(as.integer(txt$label), utils::head(sort(sp$abundance, decreasing = TRUE), 15))
  expect_gt(ab$layout$panel_params[[1]]$x.range[2], max(sp$abundance))

  iv <- ggplot2::ggplot_build(out$species_plot)
  lab <- iv$data[[2]]$label
  expect_true(all(grepl("^[0-9]+\\.[0-9]$", lab)))
  expect_gt(iv$layout$panel_params[[1]]$x.range[2], max(as.numeric(lab)))
})

test_that("the report shows the ranking before the IVI chart, with vernacular | italic scientific", {
  phy <- suppressMessages(.calculate_phytosociological_metrics(.panara_like(), 1, 10))
  html <- .fr_phyto_html(phy, rep(30, 10), 1, .field_report_strings("pt"))
  expect_lt(regexpr("Figura 2. Espécies com mais indivíduos", html, fixed = TRUE),
            regexpr("Figura 3. Espécies com maior valor", html, fixed = TRUE))
  expect_true(grepl('class="chart abund"', html, fixed = TRUE))
  expect_true(grepl('Nome local | <tspan class="it">', html, fixed = TRUE))
  expect_true(grepl("Nomes: vernacular | científico.", html, fixed = TRUE))
})

test_that("forplot_balance names every output after the language", {
  fp <- withr::local_tempfile(fileext = ".xlsx")
  openxlsx::write.xlsx(RUS01, fp)
  run <- function(lg) {
    out <- withr::local_tempdir(.local_envir = parent.frame(2))
    res <- suppressMessages(suppressWarnings(forplot_balance(
      fp, language = lg, render_pdf = FALSE, dir = out, verbose = FALSE)))
    list(res = res, files = list.files(res$output_dir))
  }
  keys <- c("map_general", "map_collected", "map_uncollected", "map_uncollected_palms",
            "map_priority", "chart_species_abund", "chart_species_ivi", "chart_family_ivi",
            "chart_dbh_classes")
  for (lg in c("pt", "en")) {
    r <- run(lg)
    pre <- .out_name("default_prefix", lg)
    expected <- c(vapply(keys, .out_name, "", language = lg, prefix = pre, ext = "png"),
                  .out_name("report", lg, pre, "html"), .out_name("balance_xlsx", lg, pre, "xlsx"))
    expect_setequal(r$files, unname(expected))
    html <- paste(readLines(r$res$html_report, warn = FALSE, encoding = "UTF-8"), collapse = "\n")
    expect_true(grepl(sprintf('data-print-name="%s"', .out_name("report", lg, pre)), html, fixed = TRUE))
  }
})

test_that("the word 'Tree' of voucher folders exists in every language, as ASCII", {
  expect_identical(.tree_words(), c("Tree", "Arvore", "Arbol", "Arbre"))
  expect_identical(.out_name("tree", "pa"), "Arvore")
  expect_identical(.out_name("tree", "ma"), "Tree")
})
