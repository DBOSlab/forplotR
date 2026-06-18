library(testthat)

# -----------------------------------------------------------------------------
# helpers
# -----------------------------------------------------------------------------

.dest_cols <- c(
  "New Tag No", "New Stem Grouping", "T1", "T2", "X", "Y", "Family",
  "Original determination", "Morphospecies", "D", "POM", "ExtraD", "ExtraPOM",
  "Flag1", "Flag2", "Flag3", "LI", "CI", "CF", "CD1", "nrdups", "Height",
  "Voucher", "Silica", "Collected", "Census Notes", "CAP", "Basal Area"
)

.write_xlsx <- function(path, sheets) {
  skip_if_not_installed("writexl")
  writexl::write_xlsx(sheets, path = path)
}

.make_field_sheet_df <- function() {
  tibble::tibble(
    `New Tag No` = c("1", "2", "3", "4"),
    `New Stem Grouping` = c(NA, NA, NA, NA),
    T1 = c(1, 2, 3, 4),
    T2 = c(1, 1, 1, 1),
    X = c(1, 2, 3, 4),
    Y = c(10, 20, 30, 40),
    Family = c("Fabaceae", "Rubiaceae", "Fabaceae", "Arecaceae"),
    `Original determination` = c("Inga alba", "Coffea arabica", "Inga alba", "Attalea sp"),
    Morphospecies = NA_character_,
    D = c(10, 20, 30, 40),
    POM = c(NA, NA, NA, NA),
    ExtraD = c(NA, NA, NA, NA),
    ExtraPOM = c(NA, NA, NA, NA),
    Flag1 = c(NA, NA, NA, NA),
    Flag2 = c(NA, NA, NA, NA),
    Flag3 = c(NA, NA, NA, NA),
    LI = c(NA, NA, NA, NA),
    CI = c(NA, NA, NA, NA),
    CF = c(NA, NA, NA, NA),
    CD1 = c(NA, NA, NA, NA),
    nrdups = c(NA, NA, NA, NA),
    Height = c(5, 6, 7, 8),
    Voucher = c("DC 100", "", "GO 42", ""),
    Silica = c(NA, NA, NA, NA),
    Collected = c("yes", NA, "yes", NA),
    `Census Notes` = c(NA, NA, NA, NA),
    CAP = c(31.4, 62.8, 94.2, 125.6),
    `Basal Area` = c(NA, NA, NA, NA)
  )
}

# -----------------------------------------------------------------------------
# basic parsing helpers
# -----------------------------------------------------------------------------

test_that(".norm_nm removes punctuation and lowercases", {
  out <- .norm_nm(c("Plot Code", NA))
  expect_equal(out[2], "")
  expect_match(out[1], "plot|lot")
  expect_false(grepl("[^a-z0-9]", out[1]))
})

test_that(".pick_colname returns NA when aliases are absent", {
  df <- tibble::tibble(`Plot Code` = 1, `X(m)` = 2)
  expect_true(is.na(.pick_colname(df, c("not_here", "also_missing"))))
})

test_that(".pick_colname can recover a unique partial alias with 2+ chars", {
  df <- tibble::tibble(`Voucher Code` = 1, `X(m)` = 2)
  expect_equal(.pick_colname(df, c("voucher")), "Voucher Code")
})

test_that(".has_any reflects whether aliases exist", {
  df <- tibble::tibble(`Standardised X` = 1)
  expect_true(.has_any(df, c("Standardised X", "Standardized X")))
  expect_false(.has_any(df, c("Y", "coord_y")))
})

test_that(".parse_num parses numeric text and drops slash-delimited values", {
  x <- c("1", " 2,5 ", "3 / 4", NA, "foo")
  expect_equal(.parse_num(x), c(1, 2.5, NA, NA, NA))
})

test_that(".best_numeric_vec picks the candidate with most finite values", {
  df <- tibble::tibble(
    x_bad = c("1/2", "a", NA),
    x_ok = c("1", "2", "3")
  )
  out <- .best_numeric_vec(df, candidates = list(c("x_bad"), c("x_ok")))
  expect_equal(out, c(1, 2, 3))
})

test_that(".best_numeric_vec returns default when no candidate exists", {
  df <- tibble::tibble(a = 1:3)
  out <- .best_numeric_vec(df, candidates = list(c("x", "y")), default = 99)
  expect_equal(out, rep(99, 3))
})

test_that(".clean_chr transliterates and normalizes whitespace", {
  x <- c("  João\u00A0Silva  ", NA)
  expect_equal(.clean_chr(x), c("Joao Silva", ""))
})

test_that(".parse_year extracts 4-digit years", {
  x <- c("census_2024", "19/08/2023", "foo", NA)
  expect_equal(.parse_year(x), c(2024L, 2023L, NA_integer_, NA_integer_))
})

test_that("station normalizers behave as expected", {
  expect_equal(.normalize_station_name(c("Estação Á")), "estacao a")
  expect_equal(.normalize_station_number(c(" 01 ", "002", "")), c("1", "2", NA_character_))
})

test_that(".field_sheet_cols returns canonical schema", {
  cols <- .field_sheet_cols()
  expect_true(is.character(cols))
  expect_equal(cols, .dest_cols)
})

# -----------------------------------------------------------------------------
# sheet scoring and workbook ingestion
# -----------------------------------------------------------------------------

test_that(".score_plot_sheet rewards canonical sheets more than unrelated sheets", {
  good <- as.data.frame(.make_field_sheet_df())
  bad <- data.frame(alpha = 1:3, beta = 4:6)
  expect_gt(.score_plot_sheet(good), .score_plot_sheet(bad))
})

test_that(".read_best_plot_sheet picks the sheet with strongest plot signature", {
  skip_if_not_installed("readxl")
  skip_if_not_installed("writexl")

  junk <- tibble::tibble(a = c("foo", "bar"), b = c("x", "y"))
  good <- .make_field_sheet_df()

  path <- withr::local_tempfile(fileext = ".xlsx")
  .write_xlsx(path, list(Junk = junk, Data = good))

  out <- .read_best_plot_sheet(path)
  expect_true(is.data.frame(out))
  expect_equal(attr(out, "sheet_name"), "Data")
  expect_true(any(names(out) %in% c("New Tag No", "Tag No", "T1", "X", "Y")))
})

test_that(".read_best_plot_sheet errors on non-Excel path", {
  expect_error(
    .read_best_plot_sheet("abc.csv"),
    "expects an Excel workbook path",
    fixed = FALSE
  )
})

test_that(".read_best_plot_sheet errors when requested sheet does not exist", {
  skip_if_not_installed("readxl")
  skip_if_not_installed("writexl")

  path <- withr::local_tempfile(fileext = ".xlsx")
  .write_xlsx(path, list(Sheet1 = .make_field_sheet_df()))

  expect_error(
    .read_best_plot_sheet(path, sheet = "Missing"),
    "Requested sheet 'Missing' was not found in workbook.",
    fixed = TRUE
  )
})

# -----------------------------------------------------------------------------
# FP query conversion
# -----------------------------------------------------------------------------

test_that(".fp_query_to_field_sheet_df returns canonical field sheet unchanged when already canonical", {
  skip_if_not_installed("readxl")
  skip_if_not_installed("writexl")

  path <- withr::local_tempfile(fileext = ".xlsx")
  .write_xlsx(path, list(Data = .make_field_sheet_df()))

  out <- .fp_query_to_field_sheet_df(path)
  expect_s3_class(out, "tbl_df")
  expect_equal(names(out), .dest_cols)
  expect_equal(attr(out, "coord_mode"), "local")
})

test_that(".fp_query_to_field_sheet_df converts query-like export with local coordinates", {
  skip_if_not_installed("readxl")
  skip_if_not_installed("writexl")

  df <- tibble::tibble(
    `Plot Code` = c("AB12", "AB12"),
    `Plot Name` = c("Alpha", "Alpha"),
    PI = c("Team X", "Team X"),
    `Tag No` = c("1", "2"),
    `Sub Plot T1` = c("1", "2"),
    `Sub Plot T2` = c("1", "1"),
    X = c("1.0", "2.0"),
    Y = c("10.0", "20.0"),
    `Recommended Family` = c("Fabaceae", "Rubiaceae"),
    `Recommended Species` = c("Inga alba", "Coffea arabica"),
    `Voucher Code` = c("GO 1", ""),
    `Voucher Collected` = c("", "yes")
  )

  path <- withr::local_tempfile(fileext = ".xlsx")
  .write_xlsx(path, list(`Plot Dump` = df))

  out <- .fp_query_to_field_sheet_df(path)
  expect_equal(names(out), .dest_cols)
  expect_equal(attr(out, "coord_mode"), "local")
  expect_equal(attr(out, "plot_meta")$plot_code, "AB-12")
  expect_equal(out$Collected, c("yes", "yes"))
})

test_that(".fp_query_to_field_sheet_df can split multiple plots", {
  skip_if_not_installed("readxl")
  skip_if_not_installed("writexl")

  df <- tibble::tibble(
    `Plot Code` = c("AB12", "CD34"),
    `Tag No` = c("1", "2"),
    `Sub Plot T1` = c("1", "1"),
    `Sub Plot T2` = c("1", "1"),
    X = c("1", "2"),
    Y = c("10", "20"),
    `Recommended Family` = c("Fabaceae", "Rubiaceae"),
    `Recommended Species` = c("Inga alba", "Coffea arabica")
  )

  path <- withr::local_tempfile(fileext = ".xlsx")
  .write_xlsx(path, list(Data = df))

  out <- .fp_query_to_field_sheet_df(path, split_plots = TRUE)
  expect_type(out, "list")
  expect_setequal(names(out), c("AB12", "CD34"))
  expect_true(all(vapply(out, inherits, logical(1), what = "tbl_df")))
})

test_that(".fp_query_to_field_sheet_df errors when plot_code is absent", {
  skip_if_not_installed("readxl")
  skip_if_not_installed("writexl")

  df <- tibble::tibble(
    `Plot Code` = c("AB12"),
    `Tag No` = c("1"),
    `Sub Plot T1` = c("1"),
    `Sub Plot T2` = c("1"),
    X = c("1"),
    Y = c("10"),
    `Recommended Family` = c("Fabaceae"),
    `Recommended Species` = c("Inga alba")
  )

  path <- withr::local_tempfile(fileext = ".xlsx")
  .write_xlsx(path, list(Data = df))

  expect_error(
    .fp_query_to_field_sheet_df(path, plot_code = "ZZ99"),
    "Requested `plot_code` was not found",
    fixed = TRUE
  )
})

# -----------------------------------------------------------------------------
# MONITORA conversion
# -----------------------------------------------------------------------------

test_that(".monitora_to_field_sheet_df returns latest census rows when station_name is NULL", {
  skip_if_not_installed("readxl")
  skip_if_not_installed("writexl")

  df <- tibble::tibble(
    Ano = c("2023", "2024", "2024"),
    Nome_estacao = c("A", "A", "B"),
    N_arvore = c("10", "1", "2"),
    N_parcela = c("1", "1", "1"),
    subunidade = c("N", "N", "S"),
    X = c("9", "1", "2"),
    Y = c("19", "10", "20"),
    cap_tot = c("20", "31.4", "31.4"),
    Familia = c("Fabaceae", "Fabaceae", "Rubiaceae"),
    Genero = c("Old", "Inga", "Coffea"),
    Especie = c("oldsp", "sp.", "arabica")
  )

  path <- withr::local_tempfile(fileext = ".xlsx")
  .write_xlsx(path, list(Sheet1 = df))

  out <- .monitora_to_field_sheet_df(path, sheet = 1, station_name = NULL)
  expect_true(is.data.frame(out))
  expect_true(all(.dest_cols %in% names(out)))
  expect_equal(nrow(out), 2)
  expect_equal(sort(out$`New Tag No`), c("1", "2"))
})

test_that(".monitora_to_field_sheet_df filters correctly when station_name exists", {
  skip_if_not_installed("readxl")
  skip_if_not_installed("writexl")

  df <- tibble::tibble(
    Ano = c("2024", "2024"),
    Nome_estacao = c("A", "B"),
    N_arvore = c("1", "2"),
    N_parcela = c("1", "1"),
    subunidade = c("N", "S"),
    X = c("1", "2"),
    Y = c("10", "20"),
    cap_tot = c("31.4", "31.4"),
    Familia = c("Fabaceae", "Rubiaceae"),
    Genero = c("Inga", "Coffea"),
    Especie = c("sp.", "arabica")
  )

  path <- withr::local_tempfile(fileext = ".xlsx")
  .write_xlsx(path, list(Sheet1 = df))

  out <- .monitora_to_field_sheet_df(path, sheet = 1, station_name = "A")

  expect_true(is.data.frame(out))
  expect_equal(nrow(out), 1)
  expect_equal(out$`New Tag No`, "1")
  expect_equal(out$T1, 1)
  expect_equal(out$T2, 1)
})

test_that(".monitora_to_field_sheet_df backfills missing coordinates from older census", {
  skip_if_not_installed("readxl")
  skip_if_not_installed("writexl")

  df <- tibble::tibble(
    Ano = c("2023", "2024"),
    Nome_estacao = c("A", "A"),
    N_arvore = c("1", "1"),
    N_parcela = c("1", "1"),
    subunidade = c("N", "N"),
    X = c("3", ""),
    Y = c("13", ""),
    cap_tot = c("31.4", "31.4"),
    Familia = c("Fabaceae", "Fabaceae"),
    Genero = c("Inga", "Inga"),
    Especie = c("alba", "alba")
  )

  path <- withr::local_tempfile(fileext = ".xlsx")
  .write_xlsx(path, list(Sheet1 = df))

  out <- .monitora_to_field_sheet_df(path)
  expect_equal(out$X, 3)
  expect_equal(out$Y, 13)
})

test_that(".monitora_to_field_sheet_df errors when requested station_name is not found", {
  skip_if_not_installed("readxl")
  skip_if_not_installed("writexl")

  df <- tibble::tibble(
    Ano = c("2024"),
    Nome_estacao = c("A"),
    N_arvore = c("1"),
    N_parcela = c("1"),
    subunidade = c("N"),
    X = c("1"),
    Y = c("10"),
    cap_tot = c("31.4"),
    Familia = c("Fabaceae"),
    Genero = c("Inga"),
    Especie = c("sp.")
  )

  path <- withr::local_tempfile(fileext = ".xlsx")
  .write_xlsx(path, list(Sheet1 = df))

  expect_error(
    .monitora_to_field_sheet_df(path, sheet = 1, station_name = "Z"),
    "Requested `station_name` not found in the most recent census.",
    fixed = TRUE
  )
})

# -----------------------------------------------------------------------------
# geometry helpers
# -----------------------------------------------------------------------------

test_that(".compute_global_coordinates computes serpentine plot coordinates", {
  fp_df <- tibble::tibble(
    T1 = c(1, 10, 11),
    X = c(1, 2, 3),
    Y = c(1, 2, 3)
  )

  res <- .compute_global_coordinates(
    fp_df,
    subplot_size = 10,
    plot_width_m = 100,
    plot_length_m = 100
  )

  expect_true(all(c("global_x", "global_y", "col", "row") %in% names(res)))
  expect_equal(res$global_x[1], 1)
  expect_equal(res$global_y[1], 1)
  expect_equal(res$global_x[3], 13)
})

test_that(".compute_global_coordinates filters out coordinates outside plot bounds", {
  fp_df <- tibble::tibble(
    T1 = c(1, 1),
    X = c(1, 200),
    Y = c(1, 1)
  )

  res <- .compute_global_coordinates(
    fp_df,
    subplot_size = 10,
    plot_width_m = 100,
    plot_length_m = 100
  )

  expect_equal(nrow(res), 1)
})

test_that(".compute_global_coordinates supports non-square plot dimensions", {
  fp_df <- tibble::tibble(
    T1 = c(1, 100, 101),
    X = c(1, 2, 3),
    Y = c(1, 2, 3)
  )

  res <- .compute_global_coordinates(
    fp_df,
    subplot_size = 10,
    plot_width_m = 100,
    plot_length_m = 1000
  )

  expect_true(nrow(res) > 0)
  expect_true(all(res$global_x >= 0))
  expect_true(all(res$global_y >= 0))
})

test_that(".compute_monitora_geometry works for full layout and local cell filtering", {
  fp_df <- tibble::tibble(
    T1 = c(1, 2, 3, 4),
    T2 = c(1, 2, 3, 4),
    X = c(5, -5, 0, 10),
    Y = c(10, 20, 30, 40)
  )

  full <- .compute_monitora_geometry(fp_df, keep_only_cell = FALSE)
  cell <- .compute_monitora_geometry(fp_df, keep_only_cell = TRUE)

  expect_s3_class(full, "tbl_df")
  expect_true(all(c("draw_x", "draw_y", "subunit_letter", "subplot_name", "x10", "y10") %in% names(full)))
  expect_true(all(grepl("^[NSLO]\\d+$", full$subplot_name)))
  expect_true(nrow(cell) <= nrow(full))
  expect_true(all(cell$x10 >= 0 & cell$x10 <= 10))
  expect_true(all(cell$y10 >= 0 & cell$y10 <= 10))
})

test_that(".compute_monitora_geometry clamps extreme coordinates and drops non-finite draw coords", {
  fp <- tibble::tibble(
    T1 = c(1, 2, 3, 4, NA),
    T2 = c(1, 1, 1, 1, 1),
    X = c(100, -100, 0, 0, 5),
    Y = c(-10, 10, 60, 10, NA)
  )

  res <- .compute_monitora_geometry(fp, keep_only_cell = FALSE)
  expect_true(all(res$X_loc >= -10 & res$X_loc <= 10))
  expect_true(all(res$Y_loc >= 0 & res$Y_loc <= 50))
  expect_true(all(is.finite(res$draw_x)))
  expect_true(all(is.finite(res$draw_y)))
})

# -----------------------------------------------------------------------------
# analytical helpers
# -----------------------------------------------------------------------------

test_that(".calculate_phytosociological_metrics computes richness and diversity", {
  fp <- .make_field_sheet_df()
  out <- .calculate_phytosociological_metrics(fp)

  expect_type(out, "list")
  expect_true(all(c("species_metrics", "family_metrics", "diversity_metrics") %in% names(out)))
  expect_true(nrow(out$species_metrics) >= 1)
  expect_equal(out$diversity_metrics$total_individuals, 4)
  expect_equal(out$diversity_metrics$total_families, 3)
})

test_that(".calculate_phytosociological_metrics errors on missing required columns", {
  expect_error(
    .calculate_phytosociological_metrics(tibble::tibble(Family = "Fabaceae")),
    "Missing required canonical column",
    fixed = TRUE
  )
})

test_that(".calculate_phytosociological_metrics recodes missing taxonomy to Indet/indet", {
  fp <- tibble::tibble(
    Family = c(NA, ""),
    `Original determination` = c(NA, ""),
    T1 = c(1, 2),
    D = c(NA, NA)
  )

  out <- .calculate_phytosociological_metrics(fp)

  expect_equal(out$diversity_metrics$total_species, 1)
  expect_equal(out$diversity_metrics$total_families, 1)
  expect_equal(out$diversity_metrics$total_individuals, 2)
  expect_equal(out$species_metrics$species, "indet")
  expect_equal(out$family_metrics$family, "Indet")
})

test_that(".prepare_report_dashboard returns translated tables and plots", {
  fp <- .make_field_sheet_df()
  out <- .prepare_report_dashboard(fp, language = "pt")

  expect_type(out, "list")
  expect_true(all(c("metrics_tbl", "species_metrics_tbl", "family_metrics_tbl", "family_plot", "species_plot", "subplot_plot", "dbh_plot", "phytosoc") %in% names(out)))
  expect_true(is.data.frame(out$metrics_tbl))
  expect_true(inherits(out$family_plot, "ggplot"))
})

# -----------------------------------------------------------------------------
# misc helpers from report pipeline
# -----------------------------------------------------------------------------

test_that("%||% returns fallback only for NULL", {
  expect_equal(NULL %||% 1, 1)
  expect_equal(0 %||% 1, 0)
  expect_equal(FALSE %||% TRUE, FALSE)
})

test_that(".create_rmd_content defaults language and warns on invalid language", {
  subplot_plots <- list(
    list(data = tibble::tibble(`New Tag No` = c("1", "2"), T1 = c(1, 1)), plot = "p1")
  )
  spec_df <- tibble::tibble(Family = "Fabaceae", Species_fmt = "Acacia mangium", tag_vec = list(c("1", "2")))

  rmd0 <- .create_rmd_content(
    subplot_plots,
    tf_col = TRUE,
    tf_uncol = FALSE,
    tf_palm = FALSE,
    plot_name = "Plot",
    plot_code = "P001",
    spec_df = spec_df
  )

  expect_type(rmd0, "character")
  expect_true(any(grepl("Full Plot Report", rmd0, fixed = TRUE)))

  expect_warning(
    rmd_bad <- .create_rmd_content(
      subplot_plots,
      tf_col = TRUE,
      tf_uncol = FALSE,
      tf_palm = FALSE,
      plot_name = "Plot",
      plot_code = "P001",
      spec_df = spec_df,
      language = "xx"
    ),
    "Invalid language"
  )

  expect_true(any(grepl("Full Plot Report", rmd_bad, fixed = TRUE)))
})

test_that(".create_rmd_content translates key headings", {
  subplot_plots <- list(list(data = tibble::tibble(`New Tag No` = "1", T1 = 1), plot = "p1"))
  spec_df <- tibble::tibble(Family = "Fabaceae", Species_fmt = "Acacia", tag_vec = list("1"))

  rmd_pt <- .create_rmd_content(
    subplot_plots,
    tf_col = FALSE,
    tf_uncol = FALSE,
    tf_palm = FALSE,
    plot_name = "Plot",
    plot_code = "P001",
    spec_df = spec_df,
    language = "pt"
  )
  expect_true(any(grepl("Relatório Completo da Parcela", rmd_pt, fixed = TRUE)))
  expect_true(any(grepl("## Metadados", rmd_pt, fixed = TRUE)))

  rmd_ma <- .create_rmd_content(
    subplot_plots,
    tf_col = FALSE,
    tf_uncol = FALSE,
    tf_palm = FALSE,
    plot_name = "Plot",
    plot_code = "P001",
    spec_df = spec_df,
    language = "ma"
  )
  expect_true(any(grepl("样地完整报告", rmd_ma, fixed = TRUE)))
})

test_that(".collapse_sorted_tags sorts numeric-like tags first and preserves character tags", {
  out <- .collapse_sorted_tags(c("10", "2", "A1", "2", NA, ""), sep = " | ")
  expect_equal(out, "2 | 10 | A1")
  expect_true(is.na(.collapse_sorted_tags(c(NA, ""))))
})

test_that(".replace_empty_with_na replaces empty strings only in character columns", {
  df <- data.frame(a = c("", "x"), b = c(1, 2), stringsAsFactors = FALSE)
  out <- .replace_empty_with_na(df)
  expect_true(is.na(out$a[1]))
  expect_equal(out$b, c(1, 2))
})

test_that(".clean_fp_data coerces coordinate and diameter fields to numeric", {
  fp <- tibble::tibble(T1 = c("1", "2"), X = c("1,5", "2"), Y = c("10", "20"), D = c("5", "6"), Collected = c("yes", NA))
  out <- .clean_fp_data(fp)
  expect_equal(out$T1, c(1L, 2L))
  expect_equal(out$X, c(1.5, 2))
  expect_equal(out$Y, c(10, 20))
  expect_equal(out$D, c(5, 6))
})

test_that(".safe_char_row extracts trimmed character rows", {
  raw <- data.frame(a = c(" x ", "z"), b = c(NA, " y "), stringsAsFactors = FALSE)
  expect_equal(.safe_char_row(raw, 1), c("x", ""))
  expect_equal(.safe_char_row(c(" a ", NA)), c("a", ""))
})

test_that(".detect_coordinate_mode distinguishes local and standardised coordinates", {
  df_local <- tibble::tibble(T1 = c("1", "2"), X = c("1", "2"), Y = c("3", "4"))
  df_std <- tibble::tibble(
    `Standardised SubPlot T1` = c("1", "2"),
    `Standardised X` = c("1", "2"),
    `Standardised Y` = c("3", "4")
  )
  df_xy <- tibble::tibble(X = c("1", "2"), Y = c("3", "4"))

  expect_equal(.detect_coordinate_mode(df_local), "local")
  expect_equal(.detect_coordinate_mode(df_std), "standardised")
  expect_equal(.detect_coordinate_mode(df_xy), "local_xy_only")
  expect_equal(.detect_coordinate_mode(tibble::tibble(a = 1)), "unknown")
})

test_that(".find_field_header_row finds the best matching header row", {
  raw <- data.frame(
    V1 = c("metadata", "New Tag No", "1"),
    V2 = c("metadata", "T1", "2"),
    V3 = c("metadata", "X", "3"),
    V4 = c("metadata", "Y", "4"),
    V5 = c("metadata", "Family", "Fabaceae"),
    V6 = c("metadata", "D", "10"),
    stringsAsFactors = FALSE
  )
  expect_equal(.find_field_header_row(raw), 2L)
})

# -----------------------------------------------------------------------------
# collection balance spreadsheet
# -----------------------------------------------------------------------------

test_that(".collection_percentual writes workbook with expected sheets", {
  out_dir <- tempdir()

  fp <- tibble::tibble(
    `New Tag No` = c("1", "2", "3"),
    T1 = c(1, 1, 2),
    Family = c("Fabaceae", "Arecaceae", "Myrtaceae"),
    Collected = c("yes", NA, "yes")
  )

  out_file <- .collection_percentual(
    fp_sheet = fp,
    dir = out_dir,
    plot_name = "Plot",
    plot_code = "P1",
    team = "Team",
    plot_width_m = 100,
    plot_length_m = 100,
    subplot_size = 10
  )

  expect_true(file.exists(out_file))

  sheets <- openxlsx::getSheetNames(out_file)
  expect_equal(
    sort(sheets),
    sort(c("COLLECTION_PERCENTUAL", "NOT_COLLECTED", "COLLECTED"))
  )
})

# -----------------------------------------------------------------------------
# herbarium / voucher helpers
# These are mostly deterministic and safe on CRAN. Network or DB dependent
# helpers should live in a separate test file with skip_on_cran().
# -----------------------------------------------------------------------------

test_that(".extract_number_token and .normalize_number_token normalize voucher numbers", {
  expect_equal(.extract_number_token("DC 00123"), "00123")
  expect_equal(
    .normalize_number_token(c("00123", "12-A", NA)),
    c("00123", "12", NA_character_)
  )
})

test_that(".split_recordedby_people splits common collector separators", {
  x <- .split_recordedby_people("Silva; Souza & Lima")
  expect_true(is.character(x))
  expect_true(length(x) >= 2)
})

test_that(".collector_tokens_one returns primary and fallback tokens", {
  tok <- .collector_tokens_one("Domingos Cardoso")
  expect_true(all(c("primary", "fallback") %in% names(tok)))
  expect_true(is.character(tok[["primary"]]))
})

test_that(".expand_collector_code expands mapped compact collector codes", {
  mp <- c(DC = "Domingos Cardoso", GO = "Giulia Ottino")
  expect_equal(.expand_collector_code("DC", collector_codes = mp), "Domingos Cardoso")
  expect_true(is.na(.expand_collector_code("XX", collector_codes = mp)))
})

test_that(".parse_voucher_one parses compact and plain collector-number strings", {
  mp <- c(DC = "Domingos Cardoso")

  v1 <- .parse_voucher_one("DC123", collector_codes = mp)
  expect_equal(unname(v1[["collector"]]), "Domingos Cardoso")
  expect_equal(unname(v1[["number"]]), "123")

  v2 <- .parse_voucher_one("Domingos Cardoso 123")
  expect_equal(unname(v2[["collector"]]), "Domingos Cardoso")
  expect_equal(unname(v2[["number"]]), "123")
})

test_that(".parse_census_identity builds raw and normalized keys", {
  mp <- c(DC = "Domingos Cardoso")
  out <- .parse_census_identity(c("DC123", "Domingos Cardoso 456"), collector_codes = mp)
  expect_true(is.data.frame(out))
  expect_true(all(c("collector_raw", "number_raw", "number_clean", "primary_key", "fallback_key") %in% names(out)))
  expect_equal(out$number_clean, c("123", "456"))
})

test_that(".resource_guess_herbarium infers herbarium code from resource id", {
  expect_equal(.resource_guess_herbarium("jbrj_rb_herbarium"), "RB")
  expect_equal(.resource_guess_herbarium("https://example.org/ipt/resource?r=ufmt_reflora"), "UFMT")
})

test_that(".get_latest_version_info parses local IPT-like HTML", {
  html <- c(
    "<html><body>",
    "latestVersion 1.2.3",
    "2026-03-10",
    "records 12,345",
    "</body></html>"
  )
  path <- withr::local_tempfile(fileext = ".html")
  writeLines(html, path)

  out <- .get_latest_version_info(path)
  expect_equal(out$version, "1.2.3")
  expect_equal(out$published_on, "2026-03-10")
  expect_true(is.character(out$records))
  expect_true(nzchar(out$records))
})

test_that("URL builders create valid JABOT and REFLORA links", {
  url1 <- .make_jabot_url("12345", "RB")
  url2 <- .make_reflora_url("12345")
  url3 <- .make_reflora_url(
    "https://reflora.jbrj.gov.br/reflora/geral/ExibeFiguraFSIUC/12345"
  )

  expect_true(is.character(url1) && length(url1) == 1)
  expect_true(is.character(url2) && length(url2) == 1)
  expect_true(is.character(url3) && length(url3) == 1)

  expect_match(url1, "codtestemunho=12345")
  expect_match(url1, "colbot=RB")

  expect_equal(
    url2,
    "https://reflora.jbrj.gov.br/reflora/geral/ExibeFiguraFSIUC/ExibeFiguraFSIUC.do?idFigura=12345"
  )

  expect_equal(url3, url2)

  expect_true(is.na(.make_reflora_url(NA_character_)))
  expect_true(is.na(.make_reflora_url("")))
})

# -----------------------------------------------------------------------------
# external / slow helpers
# -----------------------------------------------------------------------------

test_that("external herbarium helpers are callable under guarded skips", {
  skip_on_cran()
  skip_if_not_installed("curl")
  skip_if_offline()

  expect_true(is.function(.get_ipt_info))
  expect_true(is.function(.download_dwca_one))
  expect_true(is.function(.herbaria_db_connect))
  expect_true(is.function(.duckdb_load_occurrence))
  expect_true(is.function(.duckdb_match_resource))
  expect_true(is.function(.herbaria_lookup_links))
})
