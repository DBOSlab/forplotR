#' Prepare all data and plots for the collection-balance report
#'
#' Pure data-preparation step of [forplot_balance()]: takes the harmonized
#' input returned by `.harmonize_plot_input()` and computes the plot
#' coordinates, the static whole-plot maps exported as PNG and the dashboard
#' (species and family charts). It writes nothing to disk. The field report
#' itself is built from `fp_coords` by `.write_field_report()`.
#'
#' @param fp_loaded List returned by `.harmonize_plot_input()`.
#' @param input_type Canonical input type.
#' @param language Canonical language code.
#' @param plot_size,subplot_size,plot_width_m,plot_length_m Plot geometry, as
#'   resolved by [forplot_balance()] (`plot_width_m`/`plot_length_m` are `NULL`
#'   for MONITORA).
#' @param plot_geometry Result of `.resolve_plot_geometry()` or `NULL`.
#' @param highlight_palms Logical.
#' @param plot_name,plot_code,plot_census_no,team User overrides for the plot
#'   metadata (`NULL` keeps what was read from the file).
#' @param verbose Logical.
#'
#' @return A list with:
#' \describe{
#'   \item{fp_coords}{Cleaned individuals with plot coordinates.}
#'   \item{plots}{Static ggplots: `base`, `collected`, `uncollected`, `palms`,
#'     `priority` (each may be `NULL`).}
#'   \item{dashboard}{Result of `.prepare_report_dashboard()`.}
#'   \item{metadata}{Resolved `plot_name`, `plot_code`, `plot_census_no_fp`,
#'     `team`.}
#' }
#'
#' @keywords internal
#' @noRd
.prepare_balance_data <- function(fp_loaded,
                                  input_type,
                                  language,
                                  plot_size,
                                  subplot_size,
                                  plot_width_m = NULL,
                                  plot_length_m = NULL,
                                  plot_geometry = NULL,
                                  highlight_palms = TRUE,
                                  plot_name = NULL,
                                  plot_code = NULL,
                                  plot_census_no = NULL,
                                  team = NULL,
                                  verbose = TRUE) {

  # Metadata: user arguments override what was read from the file ----------
  meta <- .resolve_balance_metadata(
    fp_loaded = fp_loaded,
    plot_name = plot_name,
    plot_code = plot_code,
    plot_census_no = plot_census_no,
    team = team
  )
  plot_name <- meta$plot_name
  plot_code <- meta$plot_code

  fp_clean <- .clean_fp_data(fp_loaded$fp_sheet)

  # Coordinates -------------------------------------------------------------
  if (verbose) {
    message("Computing plot coordinates...")
  }
  if (input_type == "monitora") {
    fp_coords <- .compute_monitora_geometry(fp_clean, keep_only_cell = FALSE)
  } else {
    fp_coords <- .compute_global_coordinates(
      fp_clean = fp_clean,
      subplot_size = subplot_size,
      plot_width_m = plot_width_m,
      plot_length_m = plot_length_m
    )
  }

  fp_coords <- fp_coords %>%
    dplyr::mutate(
      diameter = (D / 100) * 2,
      `New Tag No` = trimws(as.character(`New Tag No`))
    )

  subplot_labels <- NULL

  if (input_type %in% c("field_sheet", "field_sheet_ti", "fp_query_sheet")) {
    subplot_labels <- tibble::tibble(T1 = sort(unique(fp_coords$T1))) %>%
      dplyr::mutate(
        T1 = suppressWarnings(as.numeric(T1)),
        n_subplots_per_col = floor(plot_length_m / subplot_size),
        col_index = (T1 - 1) %/% n_subplots_per_col,
        row_index = (T1 - 1) %% n_subplots_per_col,
        subplot_y = dplyr::if_else(
          col_index %% 2 == 0,
          row_index * subplot_size,
          (n_subplots_per_col - 1 - row_index) * subplot_size
        ),
        subplot_x = col_index * subplot_size,
        center_x = subplot_x + subplot_size / 2,
        center_y = subplot_y + subplot_size / 2
      ) %>%
      dplyr::select(T1, center_x, center_y)
  }

  # Whole-plot maps (PNG exports) -------------------------------------------
  if (verbose) {
    message("Building plot maps...")
  }

  if (input_type %in% c("field_sheet", "field_sheet_ti", "fp_query_sheet")) {
    base_plot_static <- .build_fp_base_plot(
      fp_coords = fp_coords,
      subplot_labels = subplot_labels,
      subplot_size = subplot_size,
      plot_width_m = plot_width_m,
      plot_length_m = plot_length_m,
      plot_name = plot_name,
      plot_code = plot_code,
      highlight_palms = highlight_palms,
      language = language
    )
  } else {
    base_plot_static <- .build_monitora_base_plot(
      fp_coords = fp_coords,
      plot_name = plot_name,
      plot_code = plot_code,
      highlight_palms = highlight_palms,
      language = language
    )
  }

  collected_plot <- NULL
  uncollected_plot <- NULL
  palms_plot <- NULL

  if (any(!is.na(fp_coords$Collected))) {
    collected_plot <- base_plot_static + (fp_coords %>% dplyr::filter(!is.na(Collected) & Family != "Arecaceae"))
  }
  if (any(is.na(fp_coords$Collected))) {
    uncollected_plot <- base_plot_static + (fp_coords %>% dplyr::filter(is.na(Collected) & Family != "Arecaceae"))
  }
  if (any(fp_coords$Family %in% "Arecaceae")) {
    palms_plot <- base_plot_static + (fp_coords %>% dplyr::filter(is.na(Collected) & Family == "Arecaceae"))
  }

  # Priority subplots (ForestPlots layouts only) ----------------------------
  priority_plot <- NULL

  if (input_type %in% c("field_sheet", "field_sheet_ti", "fp_query_sheet")) {
    priority_obj <- .prioritize_uncollected_subplots(
      fp_coords = fp_coords,
      exclude_palms = TRUE
    )

    if (!is.null(priority_obj$priority_table) && nrow(priority_obj$priority_table) > 0) {
      priority_plot <- .build_uncollected_priority_plot(
        fp_coords = fp_coords,
        priority_obj = priority_obj,
        subplot_size = subplot_size,
        plot_width_m = plot_width_m,
        plot_length_m = plot_length_m,
        plot_name = plot_name,
        plot_code = plot_code,
        language = language
      )
    }
  }

  dashboard_obj <- .prepare_report_dashboard(
    fp_sheet = fp_coords,
    input_type = input_type,
    plot_size_ha = plot_size,
    subplot_size_m = subplot_size,
    language = language
  )

  list(
    fp_coords = fp_coords,
    plots = list(
      base = base_plot_static,
      collected = collected_plot,
      uncollected = uncollected_plot,
      palms = palms_plot,
      priority = priority_plot
    ),
    dashboard = dashboard_obj,
    metadata = meta
  )
}


#' Resolve plot metadata: user arguments override values read from the file
#'
#' @keywords internal
#' @noRd
.resolve_balance_metadata <- function(fp_loaded,
                                      plot_name = NULL,
                                      plot_code = NULL,
                                      plot_census_no = NULL,
                                      team = NULL) {
  pick <- function(user, from_file, default) {
    if (.safe_nzchar(user)) {
      trimws(user)
    } else if (.safe_nzchar(from_file)) {
      trimws(from_file)
    } else {
      default
    }
  }

  list(
    plot_name = pick(plot_name, fp_loaded$plot_name, "Unknown Plot"),
    plot_code = pick(plot_code, fp_loaded$plot_code, ""),
    plot_census_no_fp = pick(plot_census_no, fp_loaded$census_no_fp, "1"),
    team = pick(team, fp_loaded$team, ""),
    census_date = if (inherits(fp_loaded$census_date, "Date")) fp_loaded$census_date else as.Date(NA)
  )
}


