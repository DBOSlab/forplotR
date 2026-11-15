# HTML report without pandoc (.render_html_report_knitr and helpers)

test_that(".strip_yaml_front_matter removes only a leading YAML block", {
  expect_identical(.strip_yaml_front_matter("---\noutput: html\n---\n# Title\n"), "# Title\n")
  expect_identical(.strip_yaml_front_matter("# Title\n\n---\n"), "# Title\n\n---\n")
})

test_that(".render_html_report_knitr builds a single self-contained report", {
  skip_if_not_installed("markdown")
  skip_if_not_installed("knitr")

  dir <- withr::local_tempdir()
  rmd <- file.path(dir, "report.Rmd")
  writeLines(c(
    "---",
    "output: html_document",
    "params:",
    "  plot_name: \"default\"",
    "  n: 3",
    "---",
    "",
    "# Overview",
    "",
    "```{r, echo=FALSE, results='asis'}",
    "cat(if (knitr::is_html_output()) '<p class=\"html-only\">HTML output</p>' else 'other')",
    "cat('\\n\\nPlot:', params$plot_name, '- n =', params$n, '\\n')",
    "```",
    "",
    "## Figure",
    "",
    "```{r, echo=FALSE, fig.width=3, fig.height=2}",
    "plot(seq_len(params$n))",
    "```",
    "",
    "# Widget",
    "",
    "```{r, echo=FALSE}",
    "leaflet::leaflet()",
    "```"
  ), rmd)

  out <- file.path(dir, "out", "report.html")
  .render_html_report_knitr(rmd, out, params = list(plot_name = "RUS01"))

  expect_true(file.exists(out))
  expect_length(list.files(dirname(out)), 1L)
  html <- .read_text_file(out)

  # parameters: supplied value overrides the YAML default, the rest is kept
  expect_match(html, "Plot: RUS01 - n = 3", fixed = TRUE)
  # knitr::is_html_output() is TRUE, as with rmarkdown
  expect_match(html, 'class="html-only"', fixed = TRUE)
  # table of contents and numbered sections
  expect_match(html, 'id="TOC"', fixed = TRUE)
  expect_match(html, 'class="section-number', fixed = TRUE)
  # figure and widget dependencies embedded
  expect_match(html, 'src="data:image/png;base64,', fixed = TRUE)
  expect_match(html, "HTMLWidgets", fixed = TRUE)
  expect_match(html, 'class="leaflet html-widget', fixed = TRUE)
  expect_false(grepl('(src|href)="(lib|figure)/', html))
  expect_false(grepl("html_preserve", html, fixed = TRUE))
})

test_that(".render_plot_report falls back to knitr when pandoc is missing", {
  called <- new.env()
  testthat::local_mocked_bindings(pandoc_available = function(...) FALSE, .package = "rmarkdown")
  testthat::local_mocked_bindings(
    .render_html_report_knitr = function(rmd_path, final_out, params = list(), envir = NULL) {
      called$args <- list(rmd_path = rmd_path, final_out = final_out, params = params)
      writeLines("<html></html>", final_out)
      invisible(final_out)
    }
  )
  dir <- withr::local_tempdir()
  rmd <- file.path(dir, "r.Rmd")
  writeLines("# x", rmd)

  out <- suppressMessages(.render_plot_report(rmd, dir, "r_full_report.html",
                                              params = list(language = "pt"), format = "html"))
  expect_identical(called$args$rmd_path, rmd)
  expect_identical(basename(called$args$final_out), "r_full_report.html")
  expect_identical(called$args$params, list(language = "pt"))
  expect_true(file.exists(out))
})
