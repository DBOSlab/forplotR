# Self-contained HTML without pandoc (.save_widget_selfcontained and helpers)

.fake_page <- function() {
  dir <- withr::local_tempdir(.local_envir = parent.frame())
  lib <- file.path(dir, "lib", "dep-1.0")
  dir.create(file.path(lib, "images"), recursive = TRUE)
  writeLines("window.depLoaded = true; var s = '</script>';", file.path(lib, "dep.js"))
  writeLines(".icon { background: url(images/icon.png); }\n.ext { background: url('https://example.org/x.png'); }",
             file.path(lib, "dep.css"))
  writeBin(as.raw(c(0x89, 0x50, 0x4e, 0x47)), file.path(lib, "images", "icon.png"))
  writeBin(as.raw(c(0x89, 0x50, 0x4e, 0x47)), file.path(lib, "marker.png"))
  writeLines(c(
    "<!DOCTYPE html><html><head>",
    '<script src="lib/dep-1.0/dep.js"></script>',
    '<link href="lib/dep-1.0/dep.css" rel="stylesheet" />',
    '<link id="marker-attachment" rel="attachment" href="lib/dep-1.0/marker.png"/>',
    '<script src="https://cdn.example.org/remote.js"></script>',
    "</head><body><a href=\"#top\">top</a></body></html>"
  ), file.path(dir, "index.html"))
  file.path(dir, "index.html")
}

test_that(".inline_html_resources embeds local scripts, styles and attachments", {
  html <- .inline_html_resources(.fake_page())

  expect_false(grepl('src="lib/', html, fixed = TRUE))
  expect_false(grepl('href="lib/', html, fixed = TRUE))
  expect_match(html, "<script>window.depLoaded = true;", fixed = TRUE)
  # a literal </script> inside the inlined code must not close the tag
  expect_match(html, "var s = '<\\/script>';", fixed = TRUE)
  expect_match(html, "<style>.icon { background: url(\"data:image/png;base64,", fixed = TRUE)
  expect_match(html, 'rel="attachment" href="data:image/png;base64,', fixed = TRUE)
})

test_that(".inline_html_resources leaves remote URLs and anchors untouched", {
  html <- .inline_html_resources(.fake_page())

  expect_match(html, '<script src="https://cdn.example.org/remote.js"></script>', fixed = TRUE)
  expect_match(html, "url('https://example.org/x.png')", fixed = TRUE)
  expect_match(html, 'href="#top"', fixed = TRUE)
})

test_that(".gsub_callback passes the match and capture groups", {
  out <- .gsub_callback("(a)(\\d)", "xa1ya2z", function(m) paste0("[", m[3], m[2], "]"))
  expect_identical(out, "x[1a]y[2a]z")
  expect_identical(.gsub_callback("q", "abc", function(m) "!"), "abc")
})

test_that(".save_widget_selfcontained writes a single file without pandoc", {
  skip_if_not_installed("leaflet")
  testthat::local_mocked_bindings(pandoc_available = function(...) FALSE, .package = "rmarkdown")

  out <- file.path(withr::local_tempdir(), "map.html")
  .save_widget_selfcontained(leaflet::leaflet(), out, title = "Test")

  expect_true(file.exists(out))
  html <- .read_text_file(out)
  expect_false(grepl('(src|href)="lib/', html))
  expect_match(html, "<title>Test</title>", fixed = TRUE)
  expect_length(list.files(dirname(out)), 1L)
})

test_that(".save_widget_selfcontained delegates to htmlwidgets when pandoc exists", {
  calls <- list()
  local_mocked_bindings(pandoc_available = function(...) TRUE, .package = "rmarkdown")
  local_mocked_bindings(
    saveWidget = function(widget, file, selfcontained, title, ...) {
      calls[[length(calls) + 1L]] <<- list(file = file, selfcontained = selfcontained,
                                            title = title)
      invisible(NULL)
    },
    .package = "htmlwidgets"
  )
  out <- withr::local_tempfile(fileext = ".html")
  expect_identical(.save_widget_selfcontained("w", out, title = "T"), out)
  expect_identical(calls, list(list(file = out, selfcontained = TRUE, title = "T")))
})
