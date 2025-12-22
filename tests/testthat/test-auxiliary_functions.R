test_that(".compute_global_coordinates works correctly", {
  # Create test data
  fp_clean <- tibble::tibble(
    T1 = c(1, 2, 3, 4, 5, 6),
    X = c(5, 10, 15, 20, 25, 30),
    Y = c(3, 6, 9, 12, 15, 18)
  )

  # Test with different plot and subplot sizes
  plot_size <- 1
  subplot_size <- 20

  result <- .compute_global_coordinates(fp_clean, plot_size, subplot_size)

  expect_s3_class(result, "tbl_df")
  expect_true(all(c("global_x", "global_y", "col", "row") %in% colnames(result)))

  # Test that coordinates are within bounds
  max_x <- 100
  max_y <- (plot_size / 1) * 100

  expect_true(all(result$global_x <= max_x & result$global_x >= 0))
  expect_true(all(result$global_y <= max_y & result$global_y >= 0))

  # Test with different plot sizes
  for (ps in c(0.2, 0.5, 1)) {
    for (ss in c(10, 20, 25)) {
      result <- .compute_global_coordinates(fp_clean, ps, ss)

      # Check that all required columns exist
      expect_true(all(c("global_x", "global_y") %in% colnames(result)))

      # Check that all coordinates are numeric
      expect_true(is.numeric(result$global_x))
      expect_true(is.numeric(result$global_y))
    }
  }

  # Test with missing coordinates
  fp_with_na <- tibble::tibble(
    T1 = c(1, 2, NA),
    X = c(5, NA, 15),
    Y = c(3, 6, NA)
  )

  result <- .compute_global_coordinates(fp_with_na, 1, 20)
  # Should handle NAs gracefully
  expect_s3_class(result, "tbl_df")
})

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

test_that("RMD content creation functions work correctly", {
  # Create mock data for testing
  subplot_plots <- list(
    list(data = tibble::tibble(`New Tag No` = c("1", "2"), T1 = c(1, 1)), plot = "plot1"),
    list(data = tibble::tibble(`New Tag No` = c("3", "4"), T1 = c(2, 2)), plot = "plot2")
  )

  tf_col <- c(TRUE, FALSE)
  tf_uncol <- c(FALSE, TRUE)
  tf_palm <- c(TRUE, TRUE)

  plot_name <- "Test Plot"
  plot_code <- "TP001"

  spec_df <- tibble::tibble(
    Family = c("Fabaceae", "Rubiaceae"),
    Species_fmt = c("Acacia mangium", "Coffea arabica"),
    tag_vec = list(c("1", "2"), c("3", "4"))
  )

  # Test English version
  rmd_en <- .create_rmd_content_en(
    subplot_plots, tf_col, tf_uncol, tf_palm,
    plot_name, plot_code, spec_df, has_agb = TRUE
  )

  expect_type(rmd_en, "character")
  expect_true(length(rmd_en) > 0)
  expect_true(any(grepl("Full Plot Report", rmd_en)))
  expect_true(any(grepl("---", rmd_en)))  # YAML header

  # Test Portuguese version
  rmd_pt <- .create_rmd_content_pt(
    subplot_plots, tf_col, tf_uncol, tf_palm,
    plot_name, plot_code, spec_df, has_agb = FALSE
  )

  expect_type(rmd_pt, "character")
  expect_true(any(grepl("Relatório Completo da Parcela", rmd_pt)))

  # Test Spanish version
  rmd_es <- .create_rmd_content_es(
    subplot_plots, tf_col, tf_uncol, tf_palm,
    plot_name, plot_code, spec_df, has_agb = TRUE
  )

  expect_type(rmd_es, "character")
  expect_true(any(grepl("Informe completo de la parcela", rmd_es)))

  # Test Mandarin version
  rmd_ma <- .create_rmd_content_ma(
    subplot_plots, tf_col, tf_uncol, tf_palm,
    plot_name, plot_code, spec_df, has_agb = FALSE
  )

  expect_type(rmd_ma, "character")
  expect_true(any(grepl("样地完整报告", rmd_ma)))

  # Test different combinations of collected/uncollected/palms
  test_combinations <- function(lang_func) {
    # All collected
    result1 <- lang_func(
      subplot_plots, c(TRUE, TRUE), c(FALSE, FALSE), c(FALSE, FALSE),
      plot_name, plot_code, spec_df
    )
    expect_true(any(grepl("collected", result1, ignore.case = TRUE)))

    # All uncollected
    result2 <- lang_func(
      subplot_plots, c(FALSE, FALSE), c(TRUE, TRUE), c(FALSE, FALSE),
      plot_name, plot_code, spec_df
    )
    expect_true(any(grepl("uncollected", result2, ignore.case = TRUE)))

    # Mixed with palms
    result3 <- lang_func(
      subplot_plots, c(TRUE, FALSE), c(FALSE, TRUE), c(TRUE, FALSE),
      plot_name, plot_code, spec_df
    )
    expect_true(any(grepl("palm", result3, ignore.case = TRUE)))
  }

  test_combinations(.create_rmd_content_en)
  test_combinations(.create_rmd_content_pt)
  test_combinations(.create_rmd_content_es)
  test_combinations(.create_rmd_content_ma)

  # Test that RMD content contains expected sections
  for (rmd in list(rmd_en, rmd_pt, rmd_es, rmd_ma)) {
    # Should have YAML header
    expect_true(grepl("^---$", rmd[1]))

    # Should have setup chunk
    expect_true(any(grepl("knitr::opts_chunk", rmd)))

    # Should have title
    expect_true(any(grepl("textbf", rmd)))

    # Should have table of contents
    expect_true(any(grepl("tableofcontents", rmd)))
  }
})
