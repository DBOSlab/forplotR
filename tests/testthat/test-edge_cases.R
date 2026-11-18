# Edge cases and less common inputs of the internal helpers

# MONITORA sheets with other spellings of the column names --------------------

.write_monitora <- function(df) {
  path <- withr::local_tempfile(fileext = ".xlsx", .local_envir = parent.frame())
  writexl::write_xlsx(list(Sheet1 = df), path)
  path
}

test_that("MONITORA columns are found under their other spellings", {
  skip_if_not_installed("readxl")
  skip_if_not_installed("writexl")

  df <- tibble::tibble(
    Data = c("2023", "2024", "2024"),
    N_estacao = c("7", "7", "7"),
    CDUC = c("UC1", "UC1", "UC1"),
    Subunidades = c("N", "N", "S"),
    N_parcela = c("1", "1", "2"),
    N_arvore = c("1", "1", "2"),
    `Nome comum` = c("ingá", "ingá", ""),
    `individuo coletado` = c("", "sim", ""),
    `voucher/numero` = c("", "12", ""),
    `X(m)` = c("1", "1", "2"),
    `Y(m)` = c("10", "10", "20"),
    `circ total` = c("31.4", "31.4", "62.8"),
    `árvore_morta` = c("não", "não", "sim"),
    observacao = c("", "ok", ""),
    ABcap = c("0.1", "0.1", "0.2"),
    `POM(m)` = c("1.3", "1.3", "1.3"),
    Familia = c("Fabaceae", "Fabaceae", "Rubiaceae"),
    Genero = c("Inga", "Inga", "Coffea"),
    Especie = c("alba", "alba", "arabica")
  )
  out <- .monitora_to_field_sheet_df(.write_monitora(df))
  expect_equal(nrow(out), 2L)
  expect_equal(sort(out$`New Tag No`), c("1", "2"))
  expect_equal(out$X[out$`New Tag No` == "1"], 1)
  meta <- attr(out, "plot_meta")
  expect_equal(meta$plot_name, "UC1")
  expect_equal(meta$plot_code, "7")
  expect_equal(meta$census_years, c(2023L, 2024L))
})

test_that("MONITORA 'collected' and coordinate columns accept every variant", {
  skip_if_not_installed("readxl")
  skip_if_not_installed("writexl")

  base <- tibble::tibble(
    subunidade = c("N", "S"), N_parcela = c("1", "2"), N_arvore = c("1", "2"),
    Familia = c("Fabaceae", "Rubiaceae"), Genero = c("Inga", "Coffea"),
    Especie = c("alba", "arabica"), cap_tot = c("31.4", "62.8")
  )
  coll <- c("individuo coletado ?", "indivíduo coletado", "indivíduo coletado ?")
  for (i in seq_along(coll)) {
    df <- base
    df[[coll[i]]] <- c("sim", "")
    df[["X (m)"]] <- c("1", "2")
    df[["Y (m)"]] <- c("3", "4")
    out <- .monitora_to_field_sheet_df(.write_monitora(df))
    expect_equal(out$Collected, c("Sim", NA), info = coll[i])
    expect_equal(out$Y, c(3, 4), info = coll[i])
  }
})

test_that("MONITORA single-column overrides, no census year and no station", {
  skip_if_not_installed("readxl")
  skip_if_not_installed("writexl")

  df <- tibble::tibble(
    subunidade = c("N", "S"), N_parcela = c("1", "2"), N_arvore = c("1", "2"),
    X = c("1", "2"), Y = c("3", "4"), Familia = c("Fabaceae", "Rubiaceae"),
    nome_cientifico = c("Inga alba", ""), tombo = c("RB 1", ""), dap_mm = c("100", "50"),
    morta = c("sim", "")
  )
  col_map <- list(`Original determination` = "nome_cientifico", Voucher = "tombo", D = "dap_mm")
  out <- .monitora_to_field_sheet_df(.write_monitora(df), col_map = col_map)
  expect_equal(out$`Original determination`, c("Inga alba", "Indet indet"))
  expect_equal(out$Voucher, c("RB 1", NA))
  expect_equal(out$D, c(100, 50))
  meta <- attr(out, "plot_meta")
  expect_length(meta$census_years, 0L)
  expect_identical(meta$plot_code, "")
})

test_that("MONITORA stops when the coordinates are all empty", {
  skip_if_not_installed("readxl")
  skip_if_not_installed("writexl")
  df <- tibble::tibble(subunidade = "N", N_parcela = "1", N_arvore = "1",
                       X = "a", Y = "b", Familia = "Fabaceae")
  expect_error(.monitora_to_field_sheet_df(.write_monitora(df)), "entirely as NA")
})

# Multi-stemmed trees ------------------------------------------------------------

test_that("multi-stem groups without a tag or with small stems are handled", {
  df <- data.frame(
    `New Tag No` = c(NA, NA, "5", "6", "7", "8"),
    `New Stem Grouping` = c("g1", "g1", "g2", "g2", "g3", "g3"),
    D = c(100, 200, 2, 3, NA, NA),
    check.names = FALSE, stringsAsFactors = FALSE
  )
  expect_warning(
    out <- suppressMessages(.consolidate_multistem_trees(df, min_diameter = 5)),
    "No valid New Tag No"
  )
  # below the minimum: the largest stem is kept; no diameter at all: NA
  expect_equal(out$D[out$`New Stem Grouping` == "g2"], 3)
  expect_true(is.na(out$D[out$`New Stem Grouping` == "g3"]))
})

# Geometry and argument checks ------------------------------------------------------

test_that("plot geometry helpers reject what they cannot place", {
  expect_error(.validate_local_xy(data.frame(a = 1), 10), "requires `X` and `Y`")
  fp <- data.frame(T1 = 1, X = 1, Y = 1)
  expect_error(.compute_global_coordinates(fp, 10, 25, 50), "exact multiples")
  expect_error(.compute_global_coordinates(fp, 10, 0, 0), "Invalid plot dimensions")

  # more than ten individuals without a subplot are listed in short
  many <- data.frame(T1 = c(rep(NA, 12), 1), X = 1, Y = 1)
  expect_warning(.compute_global_coordinates(many, 10, 100, 100), "\\(2 more\\)")

  expect_null(.compute_monitora_geometry(NULL))
  expect_error(.compute_monitora_geometry(data.frame(T1 = "N")), "requires columns")
})

test_that("small helpers return empty values for empty input", {
  expect_identical(.safe_char_row(NULL), character())
  expect_identical(.safe_char_row(list(c("a", NA))), c("a", ""))
  expect_identical(.first_nonempty(data.frame(a = c("", "-")), "a", "-"), "")
  expect_true(is.na(.find_best_col(data.frame(), "x")))
})

test_that("priority route needs its columns and stops when nothing is left", {
  expect_error(.prioritize_uncollected_subplots(data.frame(T1 = 1)), "Missing required columns")

  fp <- data.frame(
    T1 = c("1", "1", "2", "3"),
    Family = "Fabaceae",
    Collected = NA,
    `Original determination` = c("A a", "B b", "C c", "D d"),
    `New Tag No` = as.character(1:4),
    check.names = FALSE, stringsAsFactors = FALSE
  )
  one <- .prioritize_uncollected_subplots(fp, max_subplots = 1)
  expect_equal(nrow(one$priority_table), 1L)

  fp$Collected <- "Sim"
  none <- .prioritize_uncollected_subplots(fp)
  expect_equal(nrow(none$priority_table), 0L)
})

test_that("forplot_balance() checks MONITORA-only columns and the file prefix", {
  f <- withr::local_tempfile(fileext = ".xlsx")
  openxlsx::write.xlsx(RUS01, f)
  expect_error(
    expect_warning(
      forplot_balance(f, input_type = "field_sheet", colname_cap = "CAP",
                      render_html = NA, dir = withr::local_tempdir()),
      "only used when input_type = 'monitora'"
    ),
    "`render_html` must be TRUE or FALSE"
  )
  expect_error(
    suppressMessages(forplot_balance(f, input_type = "field_sheet", filename = " ",
                                     dir = withr::local_tempdir())),
    "`filename` must be NULL"
  )
})

# Field report drawing ------------------------------------------------------------------

test_that("report charts adapt their scales to larger plots and ranges", {
  tr <- .field_report_strings("en")
  expect_identical(.fr_ivi_svg(data.frame(), "species", tr), "")
  expect_identical(.fr_count_svg(data.frame(), tr), "")

  expect_equal(.fr_dbh_classes(c(1, 150))$hi[1] - .fr_dbh_classes(c(1, 150))$lo[1], 20)
  expect_equal(unname(diff(unlist(.fr_dbh_classes(c(1, 300))[1, 1:2]))), 25)
  expect_equal(unname(diff(unlist(.fr_dbh_classes(c(1, 500))[1, 1:2]))), 50)
  # many classes: smaller axis labels
  expect_match(.fr_dbh_svg(seq(1, 399, by = 5), tr), "font-size:8.5px", fixed = TRUE)

  sd <- data.frame(lx = numeric(), ly = numeric())
  for (S in c(20, 40)) {
    svg <- .fr_subplot_svg(sd, S, "subplot")
    expect_match(svg, "<svg", fixed = TRUE)
  }
})

test_that("report maps and pages handle large plots, no pending stems and no metadata", {
  fp <- withr::local_tempfile(fileext = ".xlsx")
  openxlsx::write.xlsx(RUS01, fp)
  ld <- suppressMessages(.harmonize_plot_input(fp, "field_sheet", verbose = FALSE))
  coords <- suppressWarnings(.compute_global_coordinates(.clean_fp_data(ld$fp_sheet), 10, 100, 100))
  lay <- .field_report_layout(coords, "field_sheet", 10, 100, 100, TRUE)

  for (span in c(250, 500)) {
    big <- lay
    big$extent <- c(0, 0, span, span)
    expect_match(.fr_plot_svg(big, lay$trees, "plot"), "<svg", fixed = TRUE)
  }

  done <- lay
  done$trees$st[done$trees$st == "todo"] <- "ok"
  done$cells$todo <- 0
  html <- .field_report_html(done, list(plot_name = NULL, plot_code = NULL), "en", TRUE,
                             plot_area_ha = 1)
  expect_true(grepl('class="todo done"', html, fixed = TRUE))
  expect_true(grepl("<title>Collection balance</title>", html, fixed = TRUE))

  phy <- .calculate_phytosociological_metrics(coords, plot_size_ha = 1, subplot_size_m = 10)
  ph <- .fr_phyto_html(phy, lay$trees$dbh, NA, .field_report_strings("en"))
  expect_false(grepl(">936<", ph, fixed = TRUE))
})

# Header detection helpers --------------------------------------------------------------

test_that("header helpers cope with empty sheets and list columns", {
  expect_true(is.na(.score_header_rows(data.frame())$row))
  expect_identical(.resolve_header_row(data.frame(a = "x"), fallback = 2), 2L)
  expect_identical(.simplify_list_column(1:3), 1:3)
  expect_identical(.simplify_list_column(list(TRUE, NULL, FALSE)), c(TRUE, NA, FALSE))
  out <- .simplify_list_column(list(as.Date("2024-01-02"), 3.5, "a"))
  expect_identical(out, c("2024-01-02", "3.5", "a"))
})
