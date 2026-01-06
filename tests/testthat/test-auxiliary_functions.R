# --- helpers ---------------------------------------------------------------

.dest_cols <- c(
  "New Tag No","New Stem Grouping","T1","T2","X","Y","Family",
  "Original determination","Morphospecies","D","POM","ExtraD","ExtraPOM",
  "Flag1","Flag2","Flag3","LI","CI","CF","CD1","nrdups","Height",
  "Voucher","Silica","Collected","Census Notes","CAP","Basal Area"
)

.write_xlsx <- function(path, sheets) {
  skip_if_not_installed("writexl")
  writexl::write_xlsx(sheets, path = path)
}

test_that(".compute_global_coordinates_monitora works correctly", {
  # Create test MONITORA data
  fp_df <- tibble::tibble(
    T1 = c(1, 2, 3, 4),  # 1=N, 2=S, 3=L, 4=O
    T2 = c(1, 2, 3, 4),
    X = c(5, -5, 0, 10),
    Y = c(10, 20, 30, 40)
  )

  result <- .compute_global_coordinates_monitora(fp_df)

  expect_s3_class(result, "tbl_df")
  expect_true(all(c("draw_x", "draw_y", "subunit_letter", "subplot_name") %in% colnames(result)))
  expect_true(all(c("N", "S", "L", "O") %in% result$subunit_letter))

  # Check coordinate transformations
  expect_true(is.numeric(result$draw_x))
  expect_true(is.numeric(result$draw_y))

  # Check that local coordinates are clamped
  expect_true(all(result$X_loc >= -10 & result$X_loc <= 10))
  expect_true(all(result$Y_loc >= 0 & result$Y_loc <= 50))

  # Test with NAs
  fp_with_na <- tibble::tibble(
    T1 = c(1, NA, 3),
    T2 = c(1, 2, NA),
    X = c(5, 10, NA),
    Y = c(10, NA, 30)
  )

  result <- .compute_global_coordinates_monitora(fp_with_na)
  # Should filter out rows with non-finite coordinates
  expect_true(all(is.finite(result$draw_x)))
  expect_true(all(is.finite(result$draw_y)))

  # Test subplot naming
  expect_true(all(grepl("^[NSLO]\\d+$", result$subplot_name)))
})


# --- .compute_global_coordinates_monitora ----------------------------------

test_that(".compute_global_coordinates_monitora clamps and mirrors correctly", {
  fp <- tibble::tibble(
    # NOTE: mapping per code: 1->N, 2->S, 3->L, 4->O
    T1 = c(1, 2, 3, 4),
    T2 = c(1, 1, 1, 1),
    X  = c(100, -100,  0,  0),  # clamp to 10 / -10
    Y  = c(-10,  10, 60, 10)    # clamp to 0..50
  )

  res <- .compute_global_coordinates_monitora(fp)

  expect_true(all(c("X_loc","Y_loc","along_m","draw_x","draw_y","SubplotID","subunit_letter","subplot_name") %in% names(res)))
  expect_true(all(res$X_loc >= -10 & res$X_loc <= 10))
  expect_true(all(res$Y_loc >= 0 & res$Y_loc <= 50))

  # Check mirroring for S and O (arm in c("S","O")): along_m = 50 - Y_loc
  rS <- res[res$subunit_letter == "S", , drop = FALSE]
  expect_equal(rS$Y_loc, 10)
  expect_equal(rS$along_m, 40)

  # For S: draw_y = -100 + along_m => -60
  expect_equal(rS$draw_y, -60)

  # For N: Y_loc clamped from -10 to 0 => along_m=0, draw_y=50+0=50
  rN <- res[res$subunit_letter == "N", , drop = FALSE]
  expect_equal(rN$Y_loc, 0)
  expect_equal(rN$draw_y, 50)

  # SubplotID formula: (T1-1)*10 + T2
  expect_equal(rN$SubplotID, (1 - 1) * 10 + 1)
})

test_that(".compute_global_coordinates_monitora drops non-finite coords", {
  fp <- tibble::tibble(
    T1 = c(1, NA),
    T2 = c(1, 1),
    X  = c(5,  5),
    Y  = c(5,  5)
  )

  res <- .compute_global_coordinates_monitora(fp)
  expect_true(all(is.finite(res$draw_x)))
  expect_true(all(is.finite(res$draw_y)))
})

# --- .create_rmd_content ----------------------------------------------------

test_that(".create_rmd_content defaults language, warns on invalid, and toggles AGB", {
  subplot_plots <- list(
    list(data = tibble::tibble(`New Tag No` = c("1","2"), T1 = c(1,1)), plot = "p1")
  )
  tf_col   <- c(TRUE)
  tf_uncol <- c(FALSE)
  tf_palm  <- c(FALSE)

  spec_df <- tibble::tibble(
    Family = "Fabaceae",
    Species_fmt = "Acacia mangium",
    tag_vec = list(c("1","2"))
  )

  # default language (missing) => "en"
  rmd0 <- .create_rmd_content(
    subplot_plots, tf_col, tf_uncol, tf_palm,
    plot_name = "Plot", plot_code = "P001",
    spec_df = spec_df, has_agb = FALSE
  )
  expect_type(rmd0, "character")
  expect_true(any(grepl("Full Plot Report", rmd0, fixed = TRUE)))

  # invalid language => warning + fallback "en"
  expect_warning(
    rmd_bad <- .create_rmd_content(
      subplot_plots, tf_col, tf_uncol, tf_palm,
      plot_name = "Plot", plot_code = "P001",
      spec_df = spec_df, has_agb = FALSE,
      language = "xx"
    ),
    "Invalid language"
  )
  expect_true(any(grepl("Full Plot Report", rmd_bad, fixed = TRUE)))

  # has_agb toggles YAML and AGB section
  rmd_agb <- .create_rmd_content(
    subplot_plots, tf_col, tf_uncol, tf_palm,
    plot_name = "Plot", plot_code = "P001",
    spec_df = spec_df, has_agb = TRUE, language = "en"
  )
  expect_true(any(grepl("^  agb: NULL$", rmd_agb)))
  expect_true(any(grepl("## Above-ground Biomass", rmd_agb, fixed = TRUE)))
})

test_that(".create_rmd_content translation replaces key headings", {
  subplot_plots <- list(list(data = tibble::tibble(`New Tag No`="1", T1=1), plot="p1"))
  spec_df <- tibble::tibble(Family="Fabaceae", Species_fmt="Acacia", tag_vec=list("1"))

  rmd_pt <- .create_rmd_content(
    subplot_plots, tf_col = FALSE, tf_uncol = FALSE, tf_palm = FALSE,
    plot_name = "Plot", plot_code = "P001", spec_df = spec_df, has_agb = FALSE, language = "pt"
  )
  expect_true(any(grepl("Relatório Completo da Parcela", rmd_pt, fixed = TRUE)))
  expect_true(any(grepl("## Metadados", rmd_pt, fixed = TRUE)))
})


# --- .monitora_to_field_sheet_df -------------------------------------------

test_that(".monitora_to_field_sheet_df stops when multiple units in latest census and station_name missing", {
  skip_if_not_installed("readxl")
  skip_if_not_installed("writexl")

  df <- tibble::tibble(
    Ano = c("2024","2024"),
    Nome_estacao = c("A","B"),  # two units in most recent census
    N_arvore = c("1","2"),
    N_parcela = c("1","1"),
    Subunidade = c("N","S"),
    `X (m)` = c("1","2"),
    `Y (m)` = c("10","20"),
    CAP = c("31.4","31.4"),
    Familia = c("Fabaceae","Rubiaceae"),
    Genero = c("Inga","Coffea"),
    Especie = c("sp.","arabica")
  )

  path <- withr::local_tempfile(fileext = ".xlsx")
  .write_xlsx(path, sheets = list(Sheet1 = df))

  expect_error(
    .monitora_to_field_sheet_df(path, sheet = 1, station_name = NULL),
    "contains more than one sampling unit",
    fixed = TRUE
  )
})

test_that(".monitora_to_field_sheet_df errors when requested station_name not found", {
  skip_if_not_installed("readxl")
  skip_if_not_installed("writexl")

  df <- tibble::tibble(
    Ano = c("2024"),
    Nome_estacao = c("A"),
    N_arvore = c("1"),
    N_parcela = c("1"),
    Subunidade = c("N"),
    `X (m)` = c("1"),
    `Y (m)` = c("10")
  )

  path <- withr::local_tempfile(fileext = ".xlsx")
  .write_xlsx(path, sheets = list(Sheet1 = df))

  expect_error(
    .monitora_to_field_sheet_df(path, sheet = 1, station_name = "Z"),
    "station_name.*not found",
    fixed = FALSE
  )
})

# --- .create_rmd_content ----------------------------------------------------

test_that(".create_rmd_content supports Mandarin translation", {
  subplot_plots <- list(list(data = tibble::tibble(`New Tag No`="1", T1=1), plot="p1"))
  spec_df <- tibble::tibble(Family="Fabaceae", Species_fmt="Acacia", tag_vec=list("1"))

  rmd_ma <- .create_rmd_content(
    subplot_plots,
    tf_col = FALSE, tf_uncol = FALSE, tf_palm = FALSE,
    plot_name = "Plot", plot_code = "P001",
    spec_df = spec_df,
    has_agb = FALSE,
    language = "ma"
  )

  expect_true(any(grepl("样地完整报告", rmd_ma, fixed = TRUE)))
})


