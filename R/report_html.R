# HTML report without pandoc
# Author: Giulia Ottino & Domingos Cardoso


#' Render the R Markdown report to a single HTML file without pandoc
#'
#' Fallback for `forplot_balance(render_html = TRUE)` when pandoc is not
#' available (for example when forplotR runs in the browser through webR).
#' The report is knitted with \pkg{knitr} as HTML output (so
#' `knitr::is_html_output()` blocks behave as with rmarkdown), converted with
#' \pkg{markdown} (table of contents and numbered sections, as in the
#' rmarkdown version), and the htmlwidget dependencies, figures and logos are
#' embedded so the result is one self-contained file.
#'
#' @param rmd_path Path to the generated report `.Rmd`.
#' @param final_out Output HTML path.
#' @param params Named list of report parameters (merged over the YAML
#'   defaults, as `rmarkdown::render()` does).
#' @param envir Environment in which the report code is evaluated.
#'
#' @return `final_out`, invisibly.
#'
#' @keywords internal
#' @noRd
.render_html_report_knitr <- function(rmd_path, final_out, params = list(),
                                      envir = new.env(parent = globalenv())) {
  if (!requireNamespace("markdown", quietly = TRUE)) {
    stop(
      "Rendering the HTML report without pandoc requires the package 'markdown'. ",
      "Install it with install.packages('markdown'), install pandoc, ",
      "or use render_html = FALSE.",
      call. = FALSE
    )
  }

  work <- tempfile("fp_report_")
  dir.create(work)
  on.exit(unlink(work, recursive = TRUE, force = TRUE), add = TRUE)
  rmd <- file.path(work, "report.Rmd")
  file.copy(rmd_path, rmd)
  rmd_lines <- readLines(rmd, warn = FALSE, encoding = "UTF-8")

  # Parameters: YAML defaults overridden by the supplied values
  defaults <- lapply(knitr::knit_params(rmd_lines), function(p) p$value)
  supplied <- params %||% list()
  for (nm in names(supplied)) defaults[nm] <- list(supplied[[nm]])
  assign("params", defaults, envir = envir)

  # Knit as HTML output, with figures inside the work folder
  old_knit <- knitr::opts_knit$get(c("rmarkdown.pandoc.to", "base.dir"))
  knitr::opts_knit$set(rmarkdown.pandoc.to = "html", base.dir = work)
  on.exit(knitr::opts_knit$set(old_knit), add = TRUE)
  old_wd <- setwd(work)
  on.exit(setwd(old_wd), add = TRUE)

  knitr::knit_meta(clean = TRUE)
  knitr::opts_chunk$set(fig.path = "figure/", dev = "png", dpi = 150,
                        fig.width = 7, fig.height = 5)
  md_path <- knitr::knit(rmd, output = file.path(work, "report.md"),
                         envir = envir, quiet = TRUE, encoding = "UTF-8")
  meta <- knitr::knit_meta(clean = TRUE)

  md <- .read_text_file(md_path)
  md <- .strip_yaml_front_matter(md)

  # Raw HTML from widgets is kept aside so markdown cannot alter it
  preserved <- htmltools::extractPreserveChunks(md)
  html <- markdown::mark_html(
    text = preserved$value,
    options = "+toc+number_sections+embed_resources",
    meta = list(css = NULL, js = NULL)
  )
  html <- paste(html, collapse = "\n")
  html <- htmltools::restorePreserveChunks(html, preserved$chunks)

  head_extra <- c(
    paste0("<style>", .report_css(), "</style>"),
    .render_knit_dependencies(meta, work)
  )
  html <- sub("</head>", paste0(paste(head_extra, collapse = "\n"), "\n</head>"),
              html, fixed = TRUE)

  page <- file.path(work, "report.html")
  writeBin(charToRaw(enc2utf8(html)), page)
  out <- .inline_html_resources(page)

  dir.create(dirname(final_out), recursive = TRUE, showWarnings = FALSE)
  con <- file(final_out, open = "wb")
  on.exit(close(con), add = TRUE, after = FALSE)
  writeBin(charToRaw(enc2utf8(out)), con)
  invisible(final_out)
}


#' HTML tags for the dependencies collected while knitting (htmlwidgets)
#'
#' The dependency files are copied into `work/lib`; the tags reference them
#' relatively, and `.inline_html_resources()` embeds them afterwards.
#'
#' @keywords internal
#' @noRd
.render_knit_dependencies <- function(meta, work) {
  deps <- Filter(function(x) inherits(x, "html_dependency"), meta)
  if (!length(deps)) return(character(0))
  deps <- htmltools::resolveDependencies(deps)
  lib <- file.path(work, "lib")
  dir.create(lib, showWarnings = FALSE)
  deps <- lapply(deps, function(d) {
    d <- htmltools::copyDependencyToDir(d, lib)
    htmltools::makeDependencyRelative(d, work)
  })
  as.character(htmltools::renderDependencies(deps, srcType = "file"))
}


#' Remove a leading YAML front matter block from Markdown text
#'
#' @keywords internal
#' @noRd
.strip_yaml_front_matter <- function(md) {
  if (!grepl("^\\s*---\\s*\n", md)) return(md)
  sub("^\\s*---\\s*\n.*?\n---\\s*\n", "", md, perl = TRUE)
}


#' Stylesheet of the report rendered without pandoc
#'
#' @keywords internal
#' @noRd
.report_css <- function() {
  paste(
    "body { font-family: system-ui, -apple-system, 'Segoe UI', Roboto, sans-serif;",
    "  color: #222; line-height: 1.5; max-width: 1200px; margin: 0 auto; padding: 16px; }",
    "h1, h2, h3 { line-height: 1.25; margin-top: 1.6em; }",
    "h1 { border-bottom: 1px solid #ddd; padding-bottom: .2em; }",
    "#TOC { border: 1px solid #ddd; border-radius: 6px; padding: .5em 1.2em; margin: 1em 0 2em; }",
    "#TOC ul { list-style: none; padding-left: 1.2em; }",
    "#TOC > ul { padding-left: 0; }",
    "#TOC a { text-decoration: none; }",
    ".section-number { color: #666; margin-right: .4em; }",
    "table { border-collapse: collapse; margin: 1em 0; display: block; overflow-x: auto; }",
    "th, td { border: 1px solid #ddd; padding: .3em .6em; text-align: left; }",
    "thead th { background: #f5f5f5; }",
    "img { max-width: 100%; height: auto; }",
    "pre, code { font-family: ui-monospace, Consolas, monospace; font-size: .9em; }",
    "@media print {",
    "  body { max-width: none; padding: 0; }",
    "  #TOC { page-break-after: always; }",
    "  h1 { page-break-before: always; }",
    "  .modebar, .modebar-container { display: none !important; }",
    "  table, img, .plotly, .html-widget { page-break-inside: avoid; }",
    "}",
    sep = "\n"
  )
}
