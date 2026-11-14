# Self-contained HTML widgets without pandoc
# Author: Giulia Ottino & Domingos Cardoso


#' Save an htmlwidget as a single self-contained HTML file
#'
#' Uses `htmlwidgets::saveWidget(selfcontained = TRUE)` when pandoc is
#' available. Otherwise (for example when forplotR runs in the browser through
#' webR, where pandoc does not exist) the widget is saved with its
#' dependencies in a temporary folder and every local script, stylesheet,
#' image and font is embedded in the page, which gives the same single-file
#' result.
#'
#' @param widget An htmlwidget.
#' @param file Output HTML path.
#' @param title Page title.
#'
#' @return `file`, invisibly.
#'
#' @keywords internal
#' @noRd
.save_widget_selfcontained <- function(widget, file, title = NULL) {
  if (isTRUE(suppressWarnings(rmarkdown::pandoc_available()))) {
    htmlwidgets::saveWidget(widget, file = file, selfcontained = TRUE, title = title)
    return(invisible(file))
  }

  tmp <- tempfile("fp_widget_")
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE, force = TRUE), add = TRUE)
  page <- file.path(tmp, "index.html")
  htmlwidgets::saveWidget(widget, file = page, selfcontained = FALSE,
                          libdir = "lib", title = title)

  html <- .inline_html_resources(page)
  con <- file(file, open = "wb")
  on.exit(close(con), add = TRUE, after = FALSE)
  writeBin(charToRaw(enc2utf8(html)), con)
  invisible(file)
}


#' Embed the local resources referenced by an HTML page
#'
#' Scripts and stylesheets become inline `<script>` and `<style>` blocks;
#' other local files referenced through `src` or `href` (images, attachments)
#' become data URIs. `url()` references inside stylesheets are embedded as
#' data URIs as well. Remote URLs and missing files are left untouched.
#'
#' @param page Path to the HTML file.
#'
#' @return The self-contained HTML as a single string.
#'
#' @keywords internal
#' @noRd
.inline_html_resources <- function(page) {
  base <- dirname(page)
  html <- .read_text_file(page)

  local_file <- function(ref, dir) {
    if (grepl("^([a-z][a-z0-9+.-]*:|//|#)", ref, ignore.case = TRUE)) return(NULL)
    path <- file.path(dir, utils::URLdecode(sub("[?#].*$", "", ref)))
    if (file.exists(path) && !dir.exists(path)) path else NULL
  }

  # <script src="..."></script>  ->  inline script
  html <- .gsub_callback(
    '<script([^>]*?)\\ssrc="([^"]+)"([^>]*)>\\s*</script>', html,
    function(m) {
      path <- local_file(m[3], base)
      if (is.null(path)) return(m[1])
      # "\\\\" in the replacement is one literal backslash: </script -> <\/script
      js <- gsub("</script", "<\\\\/script", .read_text_file(path), ignore.case = TRUE)
      paste0("<script", m[2], m[4], ">", js, "</script>")
    }
  )

  # <link ... rel="stylesheet" ... href="..."> -> inline style
  html <- .gsub_callback(
    '<link\\b[^>]*>', html,
    function(m) {
      tag <- m[1]
      if (!grepl('rel="stylesheet"', tag, fixed = TRUE)) return(tag)
      href <- sub('.*\\bhref="([^"]+)".*', "\\1", tag)
      path <- local_file(href, base)
      if (is.null(path)) return(tag)
      paste0("<style>", .inline_css_urls(.read_text_file(path), dirname(path)), "</style>")
    }
  )

  # Any other local src/href (images, htmlwidgets attachments) -> data URI
  .gsub_callback(
    '\\b(src|href)="([^"]+)"', html,
    function(m) {
      path <- local_file(m[3], base)
      if (is.null(path)) return(m[1])
      paste0(m[2], '="', .file_data_uri(path), '"')
    }
  )
}


#' Embed `url()` references of a stylesheet as data URIs
#'
#' @keywords internal
#' @noRd
.inline_css_urls <- function(css, dir) {
  .gsub_callback(
    "url\\(\\s*['\"]?([^'\")]+)['\"]?\\s*\\)", css,
    function(m) {
      ref <- trimws(m[2])
      if (grepl("^([a-z][a-z0-9+.-]*:|//|#)", ref, ignore.case = TRUE)) return(m[1])
      path <- file.path(dir, utils::URLdecode(sub("[?#].*$", "", ref)))
      if (!file.exists(path) || dir.exists(path)) return(m[1])
      paste0('url("', .file_data_uri(path), '")')
    }
  )
}


#' Replace every regex match with the result of a function
#'
#' `fun` receives a character vector: the whole match followed by the
#' capture groups.
#'
#' @keywords internal
#' @noRd
.gsub_callback <- function(pattern, x, fun) {
  m <- gregexpr(pattern, x, perl = TRUE)[[1]]
  if (m[1] == -1) return(x)
  starts <- as.integer(m)
  lens <- attr(m, "match.length")
  cap_s <- attr(m, "capture.start")
  cap_l <- attr(m, "capture.length")

  out <- character(0)
  pos <- 1L
  for (i in seq_along(starts)) {
    whole <- substr(x, starts[i], starts[i] + lens[i] - 1L)
    groups <- if (is.null(cap_s)) character(0) else
      substring(x, cap_s[i, ], cap_s[i, ] + cap_l[i, ] - 1L)
    out <- c(out, substr(x, pos, starts[i] - 1L), fun(c(whole, groups)))
    pos <- starts[i] + lens[i]
  }
  paste0(c(out, substr(x, pos, nchar(x))), collapse = "")
}


#' Read a text file as a single UTF-8 string
#'
#' @keywords internal
#' @noRd
.read_text_file <- function(path) {
  raw <- readBin(path, what = "raw", n = file.info(path)$size)
  txt <- rawToChar(raw)
  Encoding(txt) <- "UTF-8"
  txt
}
