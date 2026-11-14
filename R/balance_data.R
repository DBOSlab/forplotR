#' Prepare all data and plots for the collection-balance report
#'
#' Pure data-preparation step of [forplot_balance()]: takes the harmonized
#' input returned by `.harmonize_plot_input()` and computes everything the
#' report shows (coordinates, species list, counts, priority subplots,
#' dashboard tables and the ggplot/plotly objects). It writes nothing to disk.
#'
#' The returned `params` element is exactly the list passed as `params` to the
#' Rmd report, so it can also be reused by other consumers (tests, Excel
#' export, an MCP server, a template-based report).
#'
#' @param fp_loaded List returned by `.harmonize_plot_input()`.
#' @param input_type Canonical input type.
#' @param language Canonical language code.
#' @param plot_size,subplot_size,plot_width_m,plot_length_m Plot geometry, as
#'   resolved by [forplot_balance()] (`plot_width_m`/`plot_length_m` are `NULL`
#'   for MONITORA).
#' @param plot_geometry Result of `.resolve_plot_geometry()` or `NULL`.
#' @param highlight_palms Logical.
#' @param interactive Logical. If `TRUE`, also builds the plotly versions of
#'   the maps (needed only for the HTML report).
#' @param plot_name,plot_code,plot_census_no,team User overrides for the plot
#'   metadata (`NULL` keeps what was read from the file).
#' @param verbose Logical.
#'
#' @return A list with:
#' \describe{
#'   \item{params}{List passed to the Rmd report.}
#'   \item{fp_coords}{Cleaned individuals with plot coordinates.}
#'   \item{spec_df}{Species checklist grouped by family.}
#'   \item{subplot_plots}{Per-subplot plots and data.}
#'   \item{plots}{Static ggplots: `base`, `collected`, `uncollected`, `palms`,
#'     `priority` (each may be `NULL`).}
#'   \item{dashboard}{Result of `.prepare_report_dashboard()`.}
#'   \item{tf_col,tf_uncol,tf_palm}{Logical vectors over `fp_coords`.}
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
                                  interactive = TRUE,
                                  plot_name = NULL,
                                  plot_code = NULL,
                                  plot_census_no = NULL,
                                  team = NULL,
                                  verbose = TRUE) {

  tr <- setNames(.tr_dict_vec(dict$key, language = language), dict$key)

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
  plot_census_no_fp <- meta$plot_census_no_fp
  team <- meta$team

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

  # Species checklist -------------------------------------------------------
  spec_df <- fp_coords %>%
    dplyr::mutate(
      Family = as.character(Family),
      Family = dplyr::if_else(is.na(Family) | Family == "", "Indet", Family),
      `Original determination` = as.character(`Original determination`),
      Species = dplyr::if_else(
        is.na(`Original determination`) | `Original determination` == "",
        "indet",
        `Original determination`
      ),
      Species_fmt = paste0("*", Species, "*")
    ) %>%
    dplyr::group_by(Family, Species_fmt) %>%
    dplyr::summarise(
      tags = .collapse_sorted_tags(`New Tag No`, sep = " | "),
      tag_vec = list(unique(trimws(as.character(`New Tag No`)))),
      .groups = "drop"
    ) %>%
    dplyr::arrange(dplyr::if_else(Family == "Indet", 0, 1), Family)

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

  # Whole-plot maps ---------------------------------------------------------
  if (verbose) {
    message("Building plot maps...")
  }

  base_plot_interactive <- NULL

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
    if (isTRUE(interactive)) {
      base_plot_interactive <- .build_fp_base_plot_interactive(
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
    }
  } else {
    base_plot_static <- .build_monitora_base_plot(
      fp_coords = fp_coords,
      plot_name = plot_name,
      plot_code = plot_code,
      highlight_palms = highlight_palms,
      language = language
    )
    if (isTRUE(interactive)) {
      base_plot_interactive <- .build_monitora_base_plot_interactive(
        fp_coords = fp_coords,
        plot_name = plot_name,
        plot_code = plot_code,
        highlight_palms = highlight_palms,
        language = language
      )
    }
  }

  collected_plot <- NULL
  uncollected_plot <- NULL
  palms_plot <- NULL

  tf_col <- !is.na(fp_coords$Collected)
  if (any(tf_col)) {
    collected_plot <- base_plot_static + (fp_coords %>% dplyr::filter(!is.na(Collected) & Family != "Arecaceae"))
  }

  tf_uncol <- is.na(fp_coords$Collected)
  if (any(tf_uncol)) {
    uncollected_plot <- base_plot_static + (fp_coords %>% dplyr::filter(is.na(Collected) & Family != "Arecaceae"))
  }

  tf_palm <- fp_coords$Family %in% "Arecaceae"
  if (any(tf_palm)) {
    palms_plot <- base_plot_static + (fp_coords %>% dplyr::filter(is.na(Collected) & Family == "Arecaceae"))
  }

  # Priority subplots (ForestPlots layouts only) ----------------------------
  priority_obj <- NULL
  priority_plot <- NULL
  priority_plot_interactive <- NULL

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

      if (isTRUE(interactive)) {
        priority_plot_interactive <- .build_uncollected_priority_plot_interactive(
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
  }

  # Subplot maps ------------------------------------------------------------
  if (verbose) {
    message("Building subplot maps...")
  }
  subplot_plots <- .build_balance_subplot_plots(
    fp_coords = fp_coords,
    input_type = input_type,
    subplot_size = subplot_size,
    plot_width_m = plot_width_m,
    plot_length_m = plot_length_m,
    highlight_palms = highlight_palms,
    interactive = interactive,
    plot_name = plot_name,
    plot_code = plot_code,
    language = language,
    tr = tr
  )

  # Counts ------------------------------------------------------------------
  total_specimens <- nrow(fp_coords)
  collected_count <- sum(!is.na(fp_coords$Collected) & fp_coords$Family != "Arecaceae", na.rm = TRUE)
  uncollected_count <- sum(is.na(fp_coords$Collected) & fp_coords$Family != "Arecaceae", na.rm = TRUE)
  palms_count <- sum(fp_coords$Family == "Arecaceae", na.rm = TRUE)

  if (tolower(trimws(as.character(input_type))) == "monitora") {
    tag_to_subplot <- fp_coords %>%
      dplyr::transmute(
        `New Tag No` = trimws(as.character(`New Tag No`)),
        subunit_letter = as.character(subunit_letter),
        T2 = suppressWarnings(as.integer(T2)),
        subplot_index = dplyr::case_when(
          subunit_letter == "N" ~ T2,
          subunit_letter == "S" ~ 10L + T2,
          subunit_letter == "L" ~ 20L + T2,
          subunit_letter == "O" ~ 30L + T2,
          TRUE ~ NA_integer_
        )
      ) %>%
      dplyr::distinct()
  } else {
    tag_to_subplot <- fp_coords %>%
      dplyr::select(`New Tag No`, T1) %>%
      dplyr::distinct()
  }

  dashboard_obj <- .prepare_report_dashboard(
    fp_sheet = fp_coords,
    input_type = input_type,
    plot_size_ha = plot_size,
    subplot_size_m = subplot_size,
    language = language
  )

  # Report params -----------------------------------------------------------
  param_list <- list(
    language = language,
    input_type = input_type,
    metadata = list(
      plot_name = plot_name,
      plot_census_no_fp = plot_census_no_fp,
      plot_code = plot_code,
      team = team,
      census_years = fp_loaded$census_years,
      census_n = fp_loaded$census_n
    ),
    main_plot = base_plot_static,
    interactive_main_plot = base_plot_interactive,
    subplots_list = subplot_plots,
    subplot_size = subplot_size,
    plot_geometry = plot_geometry,
    stats = list(
      total = total_specimens,
      collected = collected_count,
      uncollected = uncollected_count,
      palms = palms_count,
      spec_df = spec_df,
      dead_since_first = fp_loaded$dead_since_first,
      recruits_since_first = fp_loaded$recruits_since_first
    ),
    dashboard = dashboard_obj,
    tag_list = unique(trimws(as.character(fp_coords$`New Tag No`))),
    tag_to_subplot = tag_to_subplot
  )

  if (!is.null(priority_plot)) {
    param_list$priority_uncollected_plot <- priority_plot
  }

  if (!is.null(priority_plot_interactive)) {
    param_list$interactive_priority_uncollected_plot <- priority_plot_interactive
  }

  if (!is.null(priority_obj)) {
    param_list$priority_obj <- priority_obj
    param_list$priority_species_tbl <- .flatten_priority_species_checklist(
      priority_obj,
      original_data = fp_coords,
      language = language
    )
  }

  if (!is.null(collected_plot)) param_list$collected_plot <- collected_plot
  if (!is.null(uncollected_plot)) param_list$uncollected_plot <- uncollected_plot
  if (!is.null(palms_plot)) param_list$uncollected_palm_plot <- palms_plot

  if (isTRUE(interactive)) {
    param_list <- c(
      param_list,
      .build_balance_filtered_interactive(
        fp_coords = fp_coords,
        input_type = input_type,
        subplot_labels = subplot_labels,
        subplot_size = subplot_size,
        plot_width_m = plot_width_m,
        plot_length_m = plot_length_m,
        plot_name = plot_name,
        plot_code = plot_code,
        language = language
      )
    )
  }

  list(
    params = param_list,
    fp_coords = fp_coords,
    spec_df = spec_df,
    subplot_plots = subplot_plots,
    plots = list(
      base = base_plot_static,
      collected = collected_plot,
      uncollected = uncollected_plot,
      palms = palms_plot,
      priority = priority_plot
    ),
    dashboard = dashboard_obj,
    tf_col = tf_col,
    tf_uncol = tf_uncol,
    tf_palm = tf_palm,
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
    team = pick(team, fp_loaded$team, "")
  )
}


#' Build one map (plot + data + optional plotly) per subplot
#'
#' @keywords internal
#' @noRd
.build_balance_subplot_plots <- function(fp_coords,
                                         input_type,
                                         subplot_size,
                                         plot_width_m,
                                         plot_length_m,
                                         highlight_palms,
                                         interactive,
                                         plot_name,
                                         plot_code,
                                         language,
                                         tr) {
  subplot_plots <- list()

  if (input_type == "monitora") {
    arms_order <- c("N", "S", "L", "O")
    arm_id <- stats::setNames(1:4, arms_order)

    full_idx <- do.call(rbind, lapply(arms_order, function(a) {
      data.frame(subunit_letter = a, T2 = 1:10, stringsAsFactors = FALSE)
    }))
    full_idx$T1 <- (arm_id[full_idx$subunit_letter] - 1L) * 10L + full_idx$T2

    .safe_breaks <- function(x) {
      x <- x[is.finite(x)]
      if (!length(x)) return(c(1, 3, 7, 10))
      rng <- range(x)
      br <- round(seq(rng[1], rng[2], length.out = 4), 0)
      br[1] <- rng[1]
      br[4] <- rng[2]
      br
    }

    for (k in seq_len(nrow(full_idx))) {
      sub <- full_idx$subunit_letter[k]
      t2 <- full_idx$T2[k]

      sp_data <- fp_coords %>%
        dplyr::filter(subunit_letter == sub, T2 == t2) %>%
        dplyr::mutate(
          Status = dplyr::case_when(
            highlight_palms & Family == "Arecaceae" ~ tr["palms"],
            !is.na(Collected) & Collected != "" ~ tr["collected"],
            TRUE ~ tr["uncollected"]
          ),
          Status = factor(Status, levels = c(tr["collected"], tr["uncollected"], tr["palms"])),
          diameter = dplyr::if_else(is.finite(D / 10), D / 10, NA_real_)
        )

      sp_plot <- .compute_monitora_geometry(sp_data, keep_only_cell = TRUE)

      brk <- .safe_breaks(sp_plot$diameter)
      brk <- unique(stats::na.omit(brk))
      if (length(brk) < 2) brk <- c(brk, brk + 0.01)
      if (anyDuplicated(brk)) brk <- unique(round(brk, 2))
      brk <- sort(brk)

      p <- ggplot2::ggplot(sp_plot, ggplot2::aes(x = x10, y = y10)) +
        ggplot2::geom_rect(ggplot2::aes(xmin = 0, xmax = 10, ymin = 0, ymax = 10),
                           fill = NA, color = "black", linewidth = 0.6) +
        ggplot2::geom_point(ggplot2::aes(size = diameter, fill = Status),
                            shape = 21, color = "black", stroke = 0.2, alpha = 0.9) +
        ggplot2::geom_text(ggplot2::aes(label = `New Tag No`), size = 1.7) +
        ggplot2::scale_fill_manual(
          values = stats::setNames(
            c("gray", "red", "gold"),
            c(tr["collected"], tr["uncollected"], tr["palms"])
          ),
          name = tr["status"]
        ) +
        ggplot2::scale_size_area(
          name = tr["dbh"],
          breaks = brk,
          labels = sprintf("%.1f", brk),
          limits = c(0, max(brk, na.rm = TRUE)),
          max_size = 10,
          guide = "legend"
        ) +
        ggplot2::labs(
          title = paste(tr["subunit"], sub, "\u2014", tr["subplot"], t2),
          subtitle = paste(tr["plot_name"], plot_name, "|", tr["plot_code"], plot_code),
          x = tr["local_x_m"],
          y = tr["local_y_m"]
        ) +
        ggplot2::coord_fixed() +
        ggplot2::theme_bw() +
        ggplot2::theme(
          panel.grid.minor = ggplot2::element_blank(),
          panel.grid.major = ggplot2::element_line(linewidth = 0.3, color = "gray80"),
          legend.position = "right"
        ) +
        ggplot2::guides(
          fill = ggplot2::guide_legend(override.aes = list(size = 5), order = 1),
          size = ggplot2::guide_legend(title = tr["dbh"], order = 2)
        )

      subplot_plots[[length(subplot_plots) + 1]] <- list(
        plot = p,
        data = sp_plot,
        plotly = if (isTRUE(interactive)) .build_monitora_subplot_plot_interactive(sp_plot,
                                                                                  subplot_size = subplot_size,
                                                                                  highlight_palms = highlight_palms,
                                                                                  language = language) else NULL
      )
    }

  } else {
    all_subplot_ids <- seq_len(
      floor(plot_width_m / subplot_size) * floor(plot_length_m / subplot_size)
    )
    for (i in all_subplot_ids) {
      sp_data <- fp_coords %>%
        dplyr::filter(T1 == i) %>%
        dplyr::mutate(
          Status = dplyr::case_when(
            highlight_palms & Family == "Arecaceae" ~ tr["palms"],
            !is.na(Collected) & Collected != "" ~ tr["collected"],
            TRUE ~ tr["uncollected"]
          ),
          Status = factor(Status, levels = c(tr["collected"], tr["uncollected"], tr["palms"])),
          diameter = dplyr::if_else(is.finite(D / 10), D / 10, NA_real_)
        )

      valid_d <- sp_data$diameter[is.finite(sp_data$diameter)]
      if (length(valid_d)) {
        min_d <- min(valid_d)
        max_d <- max(valid_d)
        breaks_d <- round(seq(min_d, max_d, length.out = 4), 0)
        breaks_d[1] <- min_d
        breaks_d[4] <- max_d
      } else {
        max_d <- 10
        breaks_d <- c(1, 3, 7, 10)
      }

      d_limit <- max_d
      breaks_d <- unique(stats::na.omit(round(breaks_d, 1)))
      if (length(breaks_d) < 2) breaks_d <- c(breaks_d, breaks_d + 0.01)

      p <- ggplot2::ggplot(sp_data, ggplot2::aes(x = X, y = Y)) +
        ggplot2::annotate("rect", xmin = 0, xmax = subplot_size, ymin = 0, ymax = subplot_size,
                          fill = NA, color = "darkolivegreen", linewidth = 0.6) +
        ggplot2::geom_point(ggplot2::aes(size = diameter, fill = Status),
                            shape = 21, color = "black", stroke = 0.2, alpha = 0.9) +
        ggplot2::geom_text(ggplot2::aes(label = `New Tag No`), size = 1.7) +
        ggplot2::scale_fill_manual(
          values = stats::setNames(
            c("gray", "red", "gold"),
            c(tr["collected"], tr["uncollected"], tr["palms"])
          ),
          name = tr["status"]
        ) +
        ggplot2::scale_size_area(
          name = tr["dbh"],
          breaks = breaks_d,
          labels = sprintf("%.1f", breaks_d),
          limits = c(0, d_limit),
          max_size = 10,
          guide = "legend"
        ) +
        ggplot2::labs(
          title = paste0(tr["plot_name"], ": ", plot_name),
          subtitle = paste0(tr["plot_code"], ": ", plot_code),
          x = tr["x_m"],
          y = tr["y_m"]
        ) +
        ggplot2::coord_fixed() +
        ggplot2::theme_bw() +
        ggplot2::theme(
          panel.grid.minor = ggplot2::element_blank(),
          panel.grid.major = ggplot2::element_line(linewidth = 0.3, color = "gray80"),
          panel.border = ggplot2::element_rect(color = "gray80", fill = NA, linewidth = 0.3),
          legend.position = "right",
        ) +
        ggplot2::guides(
          fill = ggplot2::guide_legend(override.aes = list(size = 5), order = 1),
          size = ggplot2::guide_legend(title = tr["dbh"], order = 2)
        )

      subplot_plots[[length(subplot_plots) + 1]] <- list(
        plot = p,
        data = sp_data,
        plotly = if (isTRUE(interactive)) .build_monitora_subplot_plot_interactive(sp_data,
                                                                                  subplot_size = subplot_size,
                                                                                  highlight_palms = highlight_palms,
                                                                                  language = language) else NULL
      )
    }
  }

  subplot_plots
}


#' Build plotly versions of the collected / uncollected / palm maps (HTML only)
#'
#' @return Named list with any of `interactive_collected_plot`,
#'   `interactive_uncollected_plot`, `interactive_palm_plot`.
#'
#' @keywords internal
#' @noRd
.build_balance_filtered_interactive <- function(fp_coords,
                                                input_type,
                                                subplot_labels,
                                                subplot_size,
                                                plot_width_m,
                                                plot_length_m,
                                                plot_name,
                                                plot_code,
                                                language) {
  out <- list()

  subsets <- list(
    interactive_collected_plot = list(
      data = fp_coords %>% dplyr::filter(!is.na(Collected) & Collected != "" & Family != "Arecaceae"),
      highlight_palms = FALSE
    ),
    interactive_uncollected_plot = list(
      data = fp_coords %>% dplyr::filter((is.na(Collected) | Collected == "") & Family != "Arecaceae"),
      highlight_palms = FALSE
    ),
    interactive_palm_plot = list(
      data = fp_coords %>% dplyr::filter((is.na(Collected) | Collected == "") & Family == "Arecaceae"),
      highlight_palms = TRUE
    )
  )

  for (nm in names(subsets)) {
    d <- subsets[[nm]]$data
    if (nrow(d) == 0) next

    if (input_type %in% c("field_sheet", "field_sheet_ti", "fp_query_sheet")) {
      out[[nm]] <- .build_fp_base_plot_interactive(
        fp_coords = d,
        subplot_labels = subplot_labels,
        subplot_size = subplot_size,
        plot_width_m = plot_width_m,
        plot_length_m = plot_length_m,
        plot_name = plot_name,
        plot_code = plot_code,
        highlight_palms = subsets[[nm]]$highlight_palms,
        language = language
      )
    } else if (input_type == "monitora") {
      out[[nm]] <- .build_monitora_base_plot_interactive(
        fp_coords = d,
        plot_name = plot_name,
        plot_code = plot_code,
        highlight_palms = subsets[[nm]]$highlight_palms,
        language = language
      )
    }
  }

  out
}
