# Additional tests for defensive helpers in R/def_fxns.R

test_that(".arg_check_herbarium validates herbarium codes", {
  expect_invisible(.arg_check_herbarium(c("rb", " HUEFS ", "K2")))
  expect_error(.arg_check_herbarium(NULL), "non-empty character vector")
  expect_error(.arg_check_herbarium(character(0)), "non-empty character vector")
  expect_error(.arg_check_herbarium(c("", "  ")), "non-empty character vector")
  expect_error(.arg_check_herbarium(c("RB", "R-B", "R B", "R-B")),
               "Invalid herbarium code\\(s\\): R-B, R B$")
})

test_that(".arg_check_colname accepts NULL or one non-empty string", {
  expect_invisible(.arg_check_colname(NULL, "colname_x"))
  expect_invisible(.arg_check_colname("Tag", "colname_x"))
  for (bad in list(1, c("a", "b"), NA_character_, "   ")) {
    expect_error(.arg_check_colname(bad, "colname_x"),
                 "`colname_x` must be NULL or a single non-empty")
  }
})

test_that(".build_col_map maps arguments to canonical names", {
  expect_null(.build_col_map())
  expect_null(.build_col_map(colname_x = NULL))
  expect_error(.build_col_map(colname_foo = "a"),
               "Unknown column-mapping argument\\(s\\): colname_foo")
  expect_error(.build_col_map(colname_x = 3), "must be NULL or a single")

  m <- .build_col_map(colname_x = " Xcoord ", colname_tagno = "Tag",
                      colname_station = "Estacao", colname_y = NULL)
  expect_identical(as.character(m), c("Xcoord", "Tag", "Estacao"))
  expect_identical(names(m), c("X", "New Tag No", ".station"))
  expect_identical(
    attr(m, "arg_names"),
    c(X = "colname_x", `New Tag No` = "colname_tagno", .station = "colname_station")
  )
})

test_that(".resolve_user_col matches exactly, trimmed and normalized", {
  df <- stats::setNames(data.frame(1, 2, 3), c("a", " Tag No ", "Espécie"))
  expect_identical(.resolve_user_col(df, "a"), "a")
  expect_identical(.resolve_user_col(df, "Tag No"), " Tag No ")
  expect_identical(.resolve_user_col(df, "especie"), "Espécie")

  expect_error(.resolve_user_col(df, "zzz", "colname_x"),
               "Column 'zzz' \\(from `colname_x`\\) was not found")
  expect_error(.resolve_user_col(df, "zzz"), "Column 'zzz' was not found")

  dup <- data.frame(`X.1` = 1, `x 1` = 2, check.names = FALSE)
  expect_error(.resolve_user_col(dup, "X_1", "colname_x"),
               "matches more than one input column")
})

test_that(".mapped_col returns fallback when unmapped", {
  df <- data.frame(Tag = 1, Xc = 2)
  m <- .build_col_map(colname_x = "xc")
  expect_identical(.mapped_col(df, NULL, "X", fallback = "X"), "X")
  expect_identical(.mapped_col(df, m, "Y"), NA_character_)
  expect_identical(.mapped_col(df, m, "X"), "Xc")

  m_noattr <- c(X = "nothere")
  expect_error(.mapped_col(df, m_noattr, "X"), "Column 'nothere' was not found")
})

test_that(".apply_col_map copies columns and supports swaps", {
  df <- data.frame(X = 1:2, Y = 3:4, Tag = c("a", "b"))
  expect_identical(.apply_col_map(df, NULL), df)

  station_only <- .build_col_map(colname_station = "Tag")
  expect_identical(.apply_col_map(df, station_only), df)

  swap <- .build_col_map(colname_x = "Y", colname_y = "X",
                         colname_tagno = "Tag")
  out <- .apply_col_map(df, swap)
  expect_identical(out$X, 3:4)
  expect_identical(out$Y, 1:2)
  expect_identical(out$`New Tag No`, c("a", "b"))
})

test_that(".arg_check_majorarea normalizes Brazilian states when country is 'Brasil'", {
  expect_identical(.arg_check_majorarea("ES", "Brasil"), "Espírito Santo")
  expect_identical(.arg_check_majorarea("Espirito Santo", "Brasil"),
                   "Espírito Santo")
  expect_error(.arg_check_majorarea("Atlantis", "Brasil"),
               "is not a Brazilian state")
  expect_identical(.arg_check_majorarea("Cusco", "Peru"), "Cusco")
  expect_error(.arg_check_majorarea(country = "Peru"), "'majorarea' is required")
})

test_that("lat, long and alt reject text and coerce numeric strings", {
  expect_identical(.arg_check_lat("-19.5"), -19.5)
  expect_identical(.arg_check_long("-40.1"), -40.1)
  expect_identical(.arg_check_alt("800"), 800)
  expect_error(.arg_check_lat("abc"), "valid numeric latitude")
  expect_error(.arg_check_long("abc"), "valid numeric longitude")
  expect_error(.arg_check_alt("high"), "must be a number")
  expect_error(.arg_check_lat(), "'lat' is required")
  expect_error(.arg_check_long(), "'long' is required")
  expect_error(.arg_check_minorarea(), "'minorarea' is required")
  expect_error(.arg_check_dir(1), "should be a character")
})

test_that(".check_point_in_admin skips validation offline or without spatial packages", {
  withr::local_options(forplotR.offline = TRUE)
  expect_warning(
    res <- .check_point_in_admin(-19.9, -40.5, "ES", "Brazil"),
    "Offline mode.*locality validation skipped"
  )
  expect_true(is.na(res))

  withr::local_options(forplotR.offline = FALSE)
  testthat::local_mocked_bindings(.missing_pkgs = function(pkgs) "terra")
  expect_warning(
    .check_point_in_admin(-19.9, -40.5, "ES", "Brazil"),
    "Locality validation requires the package\\(s\\) 'terra'"
  )
})

test_that(".check_point_in_admin errors on empty boundaries and partial-matches names", {
  skip_if_not_installed("geodata")
  skip_if_not_installed("terra")

  empty <- terra::vect(
    "POLYGON ((-41 -20.5, -40 -20.5, -40 -19.5, -41 -19.5, -41 -20.5))",
    crs = "EPSG:4326"
  )[0, ]
  testthat::local_mocked_bindings(gadm = function(...) empty, .package = "geodata")
  expect_error(.check_point_in_admin(-19.9, -40.5, "ES", "Brazil"),
               "No administrative boundaries found")

  square <- terra::vect(
    "POLYGON ((-41 -20.5, -40 -20.5, -40 -19.5, -41 -19.5, -41 -20.5))",
    crs = "EPSG:4326"
  )
  square$NAME_1 <- "Estado do Espirito Santo"
  testthat::local_mocked_bindings(gadm = function(...) square, .package = "geodata")
  expect_no_error(.check_point_in_admin(-19.9, -40.5, "ES", "Brazil"))
  expect_error(.check_point_in_admin(-19.9, -40.5, "ES", "Atlantida"),
               "Unknown country")
})
