library(testthat)

# =============================================================================
# CRAN-safe tests for aux_fxns_plots.R and aux_fxns_herb.R
#
# Scope:
# - all 71 top-level auxiliary functions currently defined in the two modules;
# - deterministic/offline tests by default;
# - no live HTTP download is required by this file;
# - JABOT download/cache behavior is reproduced with local fixtures/mocks.
# =============================================================================

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



.localize_function <- function(fun, bindings = list()) {
  env <- list2env(bindings, parent = environment(fun))
  environment(fun) <- env
  attr(fun, "test_env") <- env
  fun
}

.make_occurrence_df <- function() {
  data.frame(
    catalogNumber = c("RB123", "RB124"),
    occurrenceID = c("12345", "12346"),
    recordedBy = c("Domingos Cardoso", "Giulia Ottino"),
    recordNumber = c("123", "124"),
    stringsAsFactors = FALSE
  )
}

.make_memory_herbarium_db <- function() {
  skip_if_not_installed("DBI")
  skip_if_not_installed("duckdb")

  con <- DBI::dbConnect(duckdb::duckdb(), dbdir = ":memory:")

  DBI::dbExecute(con, "
    CREATE TABLE herbaria_index (
      key VARCHAR,
      key_type VARCHAR,
      catalogNumber VARCHAR,
      occurrenceID VARCHAR,
      herbarium VARCHAR,
      source VARCHAR,
      resource_id VARCHAR
    );
  ")

  con
}

.auxiliary_expected_helpers <- c(
  # aux_fxns_plots.R
  ".norm_nm",
  ".pick_colname",
  ".has_any",
  ".parse_num",
  ".best_numeric_vec",
  ".clean_chr",
  ".parse_year",
  ".normalize_station_name",
  ".normalize_station_number",
  ".field_sheet_cols",
  ".score_plot_sheet",
  ".read_best_plot_sheet",
  ".fp_query_to_field_sheet_df",
  ".pick_first_existing",
  ".get_col",
  ".parse_num_safe",
  ".meta_value",
  ".consolidate_multistem_trees",
  ".monitora_to_field_sheet_df",
  ".norm",
  ".find_best_col",
  ".find_best_numeric_col",
  ".clean_spaces",
  ".to_numeric",
  ".to_year",
  ".station_norm_name",
  ".station_norm_num",
  ".first_nonempty",
  ".map_sub",
  ".resolve_plot_geometry",
  ".validate_local_xy",
  ".plot_export_size",
  ".plot_symbol_scale",
  ".compute_global_coordinates",
  ".compute_monitora_geometry",
  ".calculate_phytosociological_metrics",
  ".prepare_report_dashboard",
  ".fmt_int",
  ".fmt_dec",
  "%||%",
  ".create_rmd_content",
  ".harmonize_plot_input",
  ".grab_meta_from_row",
  ".replace_empty_with_na",
  ".safe_char_row",
  ".prioritize_uncollected_subplots",
  ".build_uncollected_priority_plot",
  ".flatten_priority_species_checklist",
  ".get_lab",
  ".tr_dict",
  ".tr_dict_vec",

  # aux_fxns_herb.R
  ".herbaria_cache_paths",
  ".extract_number_token",
  ".normalize_number_token",
  ".split_recordedby_people",
  ".collector_tokens_one",
  ".expand_collector_code",
  ".parse_voucher_one",
  ".parse_census_identity",
  ".resource_guess_herbarium",
  ".get_latest_version_info",
  ".get_ipt_info",
  ".download_dwca_one",
  ".herbaria_db_connect",
  ".duckdb_load_occurrence",
  ".duckdb_match_resource",
  ".make_jabot_url",
  ".make_reflora_url",
  ".herbaria_lookup_links",
  ".load_resource_once",
  ".resolve_links_from_index"
)

test_that("auxiliary helper inventory contains all 71 expected functions", {
  expect_length(.auxiliary_expected_helpers, 71L)
  expect_equal(anyDuplicated(.auxiliary_expected_helpers), 0L)

  ns <- asNamespace("forplotR")

  for (nm in .auxiliary_expected_helpers) {
    expect_true(
      exists(nm, envir = ns, mode = "function", inherits = FALSE),
      info = paste("Missing auxiliary helper:", nm)
    )
  }
})



# -----------------------------------------------------------------------------
# basic parsing and schema helpers
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


test_that("remaining low-level plot helpers normalize, select and format correctly", {
  df <- data.frame(
    exact = c("1", "2"),
    numeric_bad = c("x", "3"),
    numeric_good = c("10", "20"),
    meta = c("", "Plot A"),
    stringsAsFactors = FALSE,
    check.names = FALSE
  )

  expect_equal(.pick_first_existing(df, c("missing", "exact")), "exact")
  expect_true(is.na(.pick_first_existing(df, c("missing1", "missing2"))))

  expect_equal(.get_col(df, "exact"), c("1", "2"))
  expect_equal(.get_col(df, "absent", default = 99), c(99, 99))

  expect_equal(
    .parse_num_safe(c("1,5", "2", "foo")),
    c(1.5, 2, NA_real_)
  )

  expect_equal(.meta_value(df, "meta"), "Plot A")

  expect_equal(
    .norm(c(" Família ", "X(m)")),
    c("familia", "xm")
  )

  df_alias <- data.frame(
    `Nome estação` = c("A", "B"),
    `x valor` = c("1", "2"),
    check.names = FALSE
  )

  expect_equal(
    .find_best_col(df_alias, c("nome_estacao", "nome estação")),
    "Nome estação"
  )

  numeric_df <- data.frame(
    bad = c("x", "1"),
    good = c("10", "20"),
    stringsAsFactors = FALSE
  )

  expect_equal(
    .find_best_numeric_col(
      numeric_df,
      list(c("bad"), c("good"))
    ),
    "good"
  )

  expect_equal(
    .clean_spaces(c("  A   B ", NA)),
    c("A B", "")
  )

  expect_equal(
    .to_numeric(c("1.234,5", "12/15", "foo")),
    c(1234.5, 12, NA_real_)
  )

  expect_equal(
    .to_year(c("Censo 2024", "2023-05-10", "")),
    c(2024L, 2023L, NA_integer_)
  )

  expect_equal(.station_norm_name("Estação Á"), "estacao a")
  expect_equal(
    .station_norm_num(c("001", "02", "")),
    c("1", "2", NA_character_)
  )

  first_df <- data.frame(
    x = c("", "NA", "Valor"),
    stringsAsFactors = FALSE
  )

  expect_equal(
    .first_nonempty(first_df, "x", lixo = c("", "NA")),
    "Valor"
  )

  expect_equal(
    .map_sub(c("N", "S", "L", "O", "NORTE", "SOUTH", "EAST", "WEST")),
    c(1, 2, 3, 4, 1, 2, 3, 4)
  )

  expect_equal(.fmt_int(10.7), "11")
  expect_equal(.fmt_dec(1.23456, 2), "1.23")
})



# -----------------------------------------------------------------------------
# workbook scoring and ForestPlots ingestion
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


test_that(".consolidate_multistem_trees consolidates duplicate stem groups", {
  df <- tibble::tibble(
    `New Tag No` = c("1", "1a", "2"),
    `New Stem Grouping` = c("G1", "G1", NA_character_),
    D = c(6, 8, 12),
    ExtraD = c(5, 12, NA_real_),
    Family = c("Fabaceae", "Fabaceae", "Rubiaceae")
  )

  out <- suppressMessages(
    .consolidate_multistem_trees(df, min_diameter = 5)
  )

  expect_equal(nrow(out), 2L)

  main <- out[  !is.na(out$`New Stem Grouping`) &
                  out$`New Stem Grouping` == "G1",
                , drop = FALSE
  ]
  expect_equal(
    nrow(main),
    1L
  )
  expect_equal(main$D, 10)
  expect_equal(main$ExtraD, 13)

  no_extra <- dplyr::select(df, -ExtraD)

  expect_no_error(
    suppressMessages(
      .consolidate_multistem_trees(no_extra, min_diameter = 5)
    )
  )

  incomplete <- tibble::tibble(Family = "Fabaceae")
  unchanged <- suppressMessages(
    .consolidate_multistem_trees(incomplete)
  )
  expect_equal(unchanged, incomplete)
})



# -----------------------------------------------------------------------------
# MONITORA conversion and harmonization
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


test_that(".harmonize_plot_input supports spatial and non-spatial schemas", {
  skip_if_not_installed("writexl")
  skip_if_not_installed("readxl")

  path <- withr::local_tempfile(fileext = ".xlsx")
  .write_xlsx(path, list(Data = .make_field_sheet_df()))

  spatial <- .harmonize_plot_input(
    fp_file_path = path,
    input_type = "fp_query_sheet",
    verbose = FALSE,
    require_spatial = TRUE
  )

  expect_type(spatial, "list")
  expect_true(all(c("fp_sheet", "team", "plot_name", "plot_code") %in% names(spatial)))
  expect_true(all(c("New Tag No", "T1", "X", "Y", "D") %in% names(spatial$fp_sheet)))

  core_only <- data.frame(
    Family = "Fabaceae",
    `Original determination` = "Inga alba",
    Voucher = "DC 100",
    Collected = "yes",
    check.names = FALSE,
    stringsAsFactors = FALSE
  )

  f <- .localize_function(
    .harmonize_plot_input,
    bindings = list(
      .fp_query_to_field_sheet_df = function(path) core_only
    )
  )

  non_spatial <- f(
    fp_file_path = path,
    input_type = "fp_query_sheet",
    verbose = FALSE,
    require_spatial = FALSE
  )

  expect_equal(non_spatial$fp_sheet$Family, "Fabaceae")

  expect_error(
    f(
      fp_file_path = path,
      input_type = "fp_query_sheet",
      verbose = FALSE,
      require_spatial = TRUE
    ),
    "missing required columns"
  )
})

test_that(".grab_meta_from_row extracts metadata values", {
  metadata <- c(
    "Plotcode: P001",
    "Plot Name: Test Plot",
    "Team: A; B"
  )

  expect_equal(.grab_meta_from_row(metadata, "Plotcode"), "P001")
  expect_equal(.grab_meta_from_row(metadata, "Plot Name"), "Test Plot")
  expect_equal(.grab_meta_from_row(metadata, "Team"), "A; B")
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


test_that(".resolve_plot_geometry resolves compatible rectangular grids", {
  g1 <- .resolve_plot_geometry(
    plot_size = 1,
    subplot_size = 10
  )

  expect_equal(g1$plot_width_m, 100)
  expect_equal(g1$plot_length_m, 100)
  expect_equal(g1$n_cols, 10L)
  expect_equal(g1$n_rows, 10L)
  expect_equal(g1$n_subplots, 100L)

  g2 <- .resolve_plot_geometry(
    plot_size = 2.25,
    subplot_size = 15
  )

  expect_equal(g2$plot_width_m, 150)
  expect_equal(g2$plot_length_m, 150)
  expect_equal(g2$n_subplots, 100L)

  g3 <- .resolve_plot_geometry(
    plot_size = 10,
    subplot_size = 10,
    plot_width_m = 200,
    plot_length_m = 500
  )

  expect_equal(g3$n_cols, 20L)
  expect_equal(g3$n_rows, 50L)
  expect_equal(g3$n_subplots, 1000L)

  expect_error(
    .resolve_plot_geometry(
      plot_size = 1,
      subplot_size = 15
    ),
    "do not form an integer number"
  )
})

test_that(".validate_local_xy follows the configured subplot size", {
  expect_invisible(
    .validate_local_xy(
      data.frame(
        X = c(0, 7.5, 15),
        Y = c(15, 2, 0)
      ),
      subplot_size = 15
    )
  )

  expect_error(
    .validate_local_xy(
      data.frame(
        X = c(0, 15.1),
        Y = c(4, 8)
      ),
      subplot_size = 15
    ),
    "Local X/Y coordinates must fall between"
  )
})

test_that(".compute_global_coordinates rejects local coordinates outside subplot bounds", {
  fp_df <- tibble::tibble(
    T1 = c(1, 1),
    X = c(1, 200),
    Y = c(1, 1)
  )

  expect_error(
    .compute_global_coordinates(
      fp_df,
      subplot_size = 10,
      plot_width_m = 100,
      plot_length_m = 100
    ),
    "Local X/Y coordinates must fall between"
  )
})

test_that("plot export helpers bound output dimensions and symbol scale", {
  one_ha <- .plot_export_size(
    plot_width_m = 100,
    plot_length_m = 100
  )

  ten_ha <- .plot_export_size(
    plot_width_m = 250,
    plot_length_m = 400
  )

  expect_lte(one_ha$width, 12)
  expect_lte(one_ha$height, 8)
  expect_lte(ten_ha$width, 12)
  expect_lte(ten_ha$height, 8)

  expect_equal(.plot_symbol_scale(100, 100), 1)
  expect_lt(.plot_symbol_scale(250, 400), 1)
  expect_gte(.plot_symbol_scale(1000, 1000), 0.25)
})



# -----------------------------------------------------------------------------
# analytical, dashboard, translation and report helpers
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
  out <- .prepare_report_dashboard(fp, input_type = "field_sheet", language = "pt")

  expect_type(out, "list")
  expect_true(all(c("metrics_tbl", "species_metrics_tbl", "family_metrics_tbl", "family_plot", "species_plot", "subplot_plot", "dbh_plot", "phytosoc") %in% names(out)))
  expect_true(is.data.frame(out$metrics_tbl))
  expect_true(inherits(out$family_plot, "ggplot"))
})

test_that("%||% returns fallback only for NULL", {
  expect_equal(NULL %||% 1, 1)
  expect_equal(0 %||% 1, 0)
  expect_equal(FALSE %||% TRUE, FALSE)
})

test_that(".replace_empty_with_na replaces empty strings only in character columns", {
  df <- data.frame(a = c("", "x"), b = c(1, 2), stringsAsFactors = FALSE)
  out <- .replace_empty_with_na(df)
  expect_true(is.na(out$a[1]))
  expect_equal(out$b, c(1, 2))
})

test_that(".safe_char_row extracts trimmed character rows", {
  raw <- data.frame(a = c(" x ", "z"), b = c(NA, " y "), stringsAsFactors = FALSE)
  expect_equal(.safe_char_row(raw, 1), c("x", ""))
  expect_equal(.safe_char_row(c(" a ", NA)), c("a", ""))
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
    spec_df = spec_df,
    dict = dict
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
      dict = dict,
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
    dict = dict,
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
    dict = dict,
    language = "ma"
  )
  expect_true(any(grepl("样地完整报告", rmd_ma, fixed = TRUE)))
})


test_that("translation helpers retrieve scalar and vector labels", {
  lab <- .get_lab(
    language = "pt",
    dict = dict
  )

  expect_type(lab, "list")
  expect_true(all(c("plot_name", "plot_code", "species_tbl") %in% names(lab)))
  expect_true(is.character(lab$plot_name))

  one <- .tr_dict(
    "plot_name",
    language = "pt",
    dict = dict
  )

  expect_true(is.character(one))
  expect_length(one, 1L)
  expect_false(identical(one, "plot_name"))

  many <- .tr_dict_vec(
    c("plot_name", "plot_code", "team"),
    language = "pt",
    dict = dict
  )

  expect_length(many, 3L)
  expect_true(all(nzchar(many)))

  expect_warning(
    missing <- .tr_dict(
      "this_key_does_not_exist",
      language = "pt",
      dict = dict
    ),
    "Translation key not found"
  )

  expect_equal(missing, "this_key_does_not_exist")
})

test_that("priority helpers select subplots, build a map and flatten the checklist", {
  fp <- tibble::tibble(
    T1 = c(1, 1, 2, 2, 3),
    Family = c(
      "Fabaceae",
      "Myrtaceae",
      "Fabaceae",
      "Arecaceae",
      "Rubiaceae"
    ),
    Collected = c(
      NA_character_,
      NA_character_,
      NA_character_,
      NA_character_,
      "yes"
    ),
    `Original determination` = c(
      "Inga alba",
      "Eugenia sp",
      "Inga alba",
      "Attalea sp",
      "Coffea arabica"
    ),
    `New Tag No` = c("1", "2", "3", "4", "5"),
    global_x = c(1, 2, 11, 12, 21),
    global_y = c(1, 2, 1, 2, 1),
    diameter = c(10, 20, 30, 40, 50)
  )

  obj <- .prioritize_uncollected_subplots(
    fp,
    exclude_palms = TRUE
  )

  expect_type(obj, "list")
  expect_true(all(c(
    "priority_table",
    "species_checklist",
    "subplot_summary",
    "covered_species",
    "remaining_species"
  ) %in% names(obj)))

  expect_gt(nrow(obj$priority_table), 0)
  expect_false("Attalea sp" %in% obj$covered_species)

  p <- .build_uncollected_priority_plot(
    fp_coords = fp,
    priority_obj = obj,
    subplot_size = 10,
    plot_width_m = 30,
    plot_length_m = 10,
    plot_name = "Test",
    plot_code = "P1",
    language = "en"
  )

  expect_s3_class(p, "ggplot")

  flat <- .flatten_priority_species_checklist(
    priority_obj = obj,
    original_data = fp,
    render_html = TRUE,
    language = "en"
  )

  expect_s3_class(flat, "tbl_df")
  expect_equal(nrow(flat), nrow(obj$species_checklist))

  empty <- .prioritize_uncollected_subplots(
    dplyr::mutate(fp, Collected = "yes"),
    exclude_palms = TRUE
  )

  expect_equal(nrow(empty$priority_table), 0L)
  expect_equal(
    nrow(.flatten_priority_species_checklist(empty, language = "en")),
    0L
  )
})



# -----------------------------------------------------------------------------
# herbarium identity, IPT and URL helpers
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
  expect_equal(out$records, "12,345")
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


test_that(".herbaria_cache_paths creates and reuses the configured cache", {
  root <- withr::local_tempdir()
  cache <- file.path(root, "forplotR-cache")

  withr::local_options(
    list(forplotR.herbaria_cache_dir = cache)
  )

  out1 <- .herbaria_cache_paths()
  out2 <- .herbaria_cache_paths()

  expect_true(dir.exists(out1$cache_dir))
  expect_equal(out1$cache_dir, cache)
  expect_equal(out1, out2)
  expect_equal(
    out1$duckdb_path,
    file.path(cache, "herbaria_index.duckdb")
  )
})

test_that(".get_ipt_info discovers a JABOT resource without live network access", {
  fake_dcat <- paste0(
    "https://ipt.jbrj.gov.br/jbrj/archive.do?r=jbrj_rb"
  )

  f <- .localize_function(
    .get_ipt_info,
    bindings = list(
      readLines = function(...) fake_dcat,
      .arg_check_herbarium = function(x) invisible(TRUE)
    )
  )

  out <- f(
    herbarium = "RB",
    ipt = "jabot"
  )

  expect_true(is.data.frame(out))
  expect_equal(nrow(out), 1L)
  expect_equal(out$herbarium, "RB")
  expect_equal(out$resource_id, "jbrj_rb")
  expect_match(out$archive_base, "/jbrj/archive.do")
})

test_that(".download_dwca_one reuses an existing cached occurrence file", {
  root <- withr::local_tempdir()

  info <- data.frame(
    ipt = "jabot",
    herbarium = "RB",
    resource_id = "jbrj_rb",
    archive_base = "https://example.org/archive.do?r=",
    resource_url = "https://example.org/resource?r=jbrj_rb",
    stringsAsFactors = FALSE
  )

  expected <- file.path(
    root,
    "dwca_jabot_RB_jbrj_rb"
  )

  dir.create(expected, recursive = TRUE)
  file.create(file.path(expected, "occurrence.txt"))

  out <- .download_dwca_one(
    info_row = info,
    dir = root,
    verbose = FALSE,
    force_refresh = FALSE
  )

  expect_equal(
    normalizePath(out, winslash = "/", mustWork = TRUE),
    normalizePath(expected, winslash = "/", mustWork = TRUE)
  )
})

test_that("DuckDB herbarium helpers connect, load occurrence data and match keys", {
  skip_if_not_installed("DBI")
  skip_if_not_installed("duckdb")

  root <- withr::local_tempdir()
  withr::local_options(
    list(forplotR.herbaria_cache_dir = root)
  )

  con <- .herbaria_db_connect()
  on.exit(
    try(DBI::dbDisconnect(con, shutdown = TRUE), silent = TRUE),
    add = TRUE
  )

  expect_true(
    all(c("occ_raw", "herbaria_index") %in% DBI::dbListTables(con))
  )

  occurrence <- file.path(root, "occurrence.txt")

  utils::write.table(
    .make_occurrence_df()[1, , drop = FALSE],
    occurrence,
    sep = "\t",
    quote = FALSE,
    row.names = FALSE,
    col.names = TRUE,
    na = ""
  )

  expect_true(
    .duckdb_load_occurrence(
      con = con,
      occurrence_path = occurrence,
      herbarium = "RB",
      source = "jabot",
      resource_id = "jbrj_rb",
      verbose = FALSE
    )
  )

  expect_true(
    .duckdb_match_resource(
      con = con,
      herbarium = "RB",
      source = "jabot",
      resource_id = "jbrj_rb",
      need_primary = "dcardoso||123",
      need_fallback = "cardoso||123",
      verbose = FALSE
    )
  )

  idx <- DBI::dbGetQuery(
    con,
    "SELECT * FROM herbaria_index"
  )

  expect_gt(nrow(idx), 0L)
  expect_true("RB123" %in% idx$catalogNumber)
})

test_that(".load_resource_once loads a resource once and caches its resource key", {

  root <- withr::local_tempdir()

  state <- new.env(parent = emptyenv())
  state$n_download <- 0L
  state$n_load <- 0L

  rowi <- data.frame(
    herbarium = "RB",
    source = "jabot",
    resource_id = "jbrj_rb",
    stringsAsFactors = FALSE
  )

  fake_con <- structure(
    list(),
    class = "fake_connection"
  )

  f <- .localize_function(
    .load_resource_once,
    bindings = list(

      .download_dwca_one = function(
    info_row,
    dir,
    verbose,
    force_refresh
      ) {

        state$n_download <- state$n_download + 1L

        out <- file.path(
          root,
          "dwca"
        )

        dir.create(
          out,
          recursive = TRUE,
          showWarnings = FALSE
        )

        file.create(
          file.path(
            out,
            "occurrence.txt"
          )
        )

        out
      },

    .duckdb_load_occurrence = function(...) {

      state$n_load <- state$n_load + 1L

      TRUE
    }
    )
  )

  already_loaded <- character(0)

  first <- f(
    rowi = rowi,
    dir_name = "jabot_download",
    already_loaded = already_loaded,
    con = fake_con,
    verbose = FALSE,
    force_refresh = FALSE
  )

  expect_true(first$ok)

  expect_equal(
    first$already_loaded,
    "RB::jabot::jbrj_rb"
  )

  already_loaded <- first$already_loaded

  second <- f(
    rowi = rowi,
    dir_name = "jabot_download",
    already_loaded = already_loaded,
    con = fake_con,
    verbose = FALSE,
    force_refresh = FALSE
  )

  expect_true(second$ok)

  expect_equal(
    second$already_loaded,
    "RB::jabot::jbrj_rb"
  )

  expect_equal(
    state$n_download,
    1L
  )

  expect_equal(
    state$n_load,
    1L
  )
})
test_that(".resolve_links_from_index resolves JABOT and REFLORA links", {
  fp_df <- data.frame(
    Voucher = c("DC123", "GO124"),
    stringsAsFactors = FALSE
  )

  parsed <- data.frame(
    primary_key = c("dcardoso||123", "gottino||124"),
    fallback_key = c("cardoso||123", "ottino||124"),
    stringsAsFactors = FALSE
  )

  index_jabot <- data.frame(
    key = "dcardoso||123",
    key_type = "primary",
    catalogNumber = "RB123",
    occurrenceID = NA_character_,
    herbarium = "RB",
    source = "jabot",
    stringsAsFactors = FALSE
  )

  f <- .resolve_links_from_index

  out_jabot <- f(
    index_df = index_jabot,
    rows_to_check = 1L,
    source_label = "jabot",
    fp_df = fp_df,
    parsed = parsed
  )

  expect_true(out_jabot$resolved[1])
  expect_match(out_jabot$links[1], "JABOT", fixed = TRUE)

  index_reflora <- data.frame(
    key = "gottino||124",
    key_type = "primary",
    catalogNumber = NA_character_,
    occurrenceID = "12346",
    herbarium = "RB",
    source = "reflora",
    stringsAsFactors = FALSE
  )

  out_reflora <- f(
    index_df = index_reflora,
    rows_to_check = 2L,
    source_label = "reflora",
    fp_df = fp_df,
    parsed = parsed
  )

  expect_true(out_reflora$resolved[2])
  expect_match(out_reflora$links[2], "REFLORA", fixed = TRUE)
})

test_that(".herbaria_lookup_links prioritizes JABOT in an offline mocked workflow", {
  skip_if_not_installed("DBI")
  skip_if_not_installed("duckdb")

  fp_df <- data.frame(
    Voucher = "Domingos Cardoso 123",
    stringsAsFactors = FALSE
  )

  make_con <- function() {
    .make_memory_herbarium_db()
  }

  f <- .localize_function(
    .herbaria_lookup_links,
    bindings = list(
      .arg_check_herbarium = function(x) invisible(TRUE),
      .herbaria_db_connect = make_con,
      .get_ipt_info = function(herbarium, ipt, resource_map = NULL) {
        if (identical(ipt, "reflora")) {
          return(data.frame(
            ipt = character(0),
            herbarium = character(0),
            resource_id = character(0),
            archive_base = character(0),
            resource_url = character(0),
            stringsAsFactors = FALSE
          ))
        }

        data.frame(
          ipt = "jabot",
          herbarium = herbarium,
          resource_id = "jbrj_rb",
          archive_base = "https://example.org/archive.do?r=",
          resource_url = "https://example.org/resource?r=jbrj_rb",
          stringsAsFactors = FALSE
        )
      },
      .load_resource_once = function(
    rowi,
    dir_name,
    already_loaded,
    con,
    verbose = FALSE,
    force_refresh = FALSE
      ) {

        rid_key <- paste(
          rowi$herbarium,
          rowi$source,
          rowi$resource_id,
          sep = "::"
        )

        list(
          ok = TRUE,
          already_loaded = unique(
            c(
              already_loaded,
              rid_key
            )
          )
        )
      },
    .duckdb_match_resource = function(
    con,
    herbarium,
    source,
    resource_id,
    need_primary,
    need_fallback,
    verbose = FALSE
    ) {
      row <- data.frame(
        key = need_primary[1],
        key_type = "primary",
        catalogNumber = "RB123",
        occurrenceID = "12345",
        herbarium = herbarium,
        source = source,
        resource_id = resource_id,
        stringsAsFactors = FALSE
      )
      DBI::dbWriteTable(con, "herbaria_index", row, append = TRUE)
      TRUE
    },
    .resolve_links_from_index = function(
    index_df,
    rows_to_check,
    source_label,
    fp_df,
    parsed
    ) {
      links <- rep(NA_character_, nrow(fp_df))
      resolved <- rep(FALSE, nrow(fp_df))

      if (nrow(index_df) && length(rows_to_check)) {
        i <- rows_to_check[1]
        links[i] <- "<br/><b>Herbarium image:</b> <a href='mock' target='_blank'>JABOT</a>"
        resolved[i] <- TRUE
      }

      list(links = links, resolved = resolved)
    }
    )
  )

  out <- f(
    fp_df = fp_df,
    herbaria = "RB",
    keep_downloads = TRUE,
    verbose = FALSE
  )

  expect_length(out, 1L)
  expect_match(out[1], "JABOT", fixed = TRUE)
})

test_that(".herbaria_lookup_links returns NA when Voucher is absent", {
  out <- .herbaria_lookup_links(
    fp_df = data.frame(x = 1:2),
    herbaria = "RB",
    verbose = FALSE
  )

  expect_equal(out, rep(NA_character_, 2))
})
