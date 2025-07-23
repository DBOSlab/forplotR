#' Generate forest plot specimen map and collection balance
#'
#' @author Giulia Ottino & Domingos Cardoso
#'
#' @description Processes field data collected using the \href {https://forestplots.net/}{Forestplots format}
#' and generates a specimen map (2D plot) with collection status and spatial
#' distribution of individuals across subplots. Generate a PDF with a full plot
#' report with separate maps, for each subplot and a spreadsheet summarizing the
#' percentage of collected specimens per subplot. The function performs the
#' following steps: (i) validates all input arguments and checks/create output
#' folders; (ii) reads the xlsx {Forestplots format} sheet, extracts metadata
#' (team, plot name, plot code), and cleans the data; (iii) normalizes and cleans
#' coordinate and diameter values, computing global plot coordinates from
#' subplot-relative positions; (iv) creates a PDF report, including plot metadata,
#' the main map, collected and uncollected specimen maps, optionally a map
#' highlighting palms (Arecaceae) and navigable individual subplot  maps; (v)
#' optionally generates a xlsx spreadsheet with the collection
#' percentage per subplot (including totals), distinguishing collected vs.
#' uncollected, and palms specimen.
#'
#' @usage
#' plot_for_balance(fp_file_path = NULL,
#'                  plot_size = 1,
#'                  subplot_size = 10,
#'                  highlight_palms = TRUE,
#'                  dir = "Results_map_plot",
#'                  filename = "plot_specimen")
#'
#' @param fp_file_path File path to the forestplots dataset (Excel format).
#'
#' @param plot_size Overall plot size in hectares (default is 1 ha).
#'
#' @param subplot_size Side length of subplots in meters (default is 10 m).
#'
#' @param highlight_palms Logical. If TRUE, highlights Arecaceae in purple.
#'
#' @param dir Directory path where output will be saved
#' (default is "Results_map_plot").
#'
#' @param filename Name of the output file for the main plot (without extension).
#'
#' @return
#' - A full PDF report with:
#'   - plot metadata,
#'   - navigable sections including subplot maps;
#' - An Excel spreadsheet with:
#'   - collection percentages (overall and per subplot),
#'   - tag numbers for collected and not collected individuals.
#'
#' @importFrom readxl read_excel
#' @importFrom ggplot2 ggplot aes geom_point geom_text geom_rect
#' @importFrom ggplot2 scale_x_continuous scale_y_continuous coord_fixed
#' @importFrom ggplot2 scale_size_continuous scale_fill_identity labs theme
#' @importFrom ggplot2 element_line element_blank element_text ggsave
#' @importFrom dplyr mutate group_by summarise filter select arrange ungroup
#' @importFrom dplyr tibble if_else %>%
#' @importFrom grDevices pdf dev.off colorRampPalette
#' @importFrom utils capture.output
#' @importFrom rmarkdown render
#' @importFrom tools file_path_sans_ext
#' @importFrom openxlsx createWorkbook addWorksheet writeData saveWorkbook
#' @importFrom openxlsx createStyle addStyle
#'
#' @examples
#' \dontrun{
#' plot_for_balance(fp_file_path = "data/forestplot.xlsx",
#'                  plot_size = 1,
#'                  subplot_size = 10,
#'                  highlight_palms = TRUE,
#'                  dir = "Results_map_plot",
#'                  filename = "plot_specimen")
#' }
#'
#' @export
#'

plot_for_balance <- function(fp_file_path = NULL,
                             plot_size = 1,
                             subplot_size = 10,
                             highlight_palms = TRUE,
                             dir = "Results_map_plot",
                             filename = "plot_specimen") {

  # Plot size check
  .validate_plot_size(plot_size)

  # subplot size check
  .validate_subplot_size(subplot_size)

  # dir check
  dir <- .arg_check_dir(dir)

  # Creating the directory to save the file based on the current date
  foldername <- paste0(dir, "/", format(Sys.time(), "%d%b%Y"))
  if (!dir.exists(dir)) dir.create(dir)
  if (!dir.exists(foldername)) dir.create(foldername)

  # Read entire sheet without headers
  raw <- suppressMessages(readxl::read_excel(fp_file_path, sheet = 1, col_names = FALSE))

  # Extract metadata from the first row
  metadata_row <- raw[1, ] %>% unlist() %>% as.character()

  team <- metadata_row[grepl("^Team:", metadata_row, ignore.case = TRUE)] %>%
    sub("^Team:\\s*", "", ., ignore.case = TRUE)

  plot_name <- metadata_row[grepl("^Plot Name:", metadata_row, ignore.case = TRUE)] %>%
    sub("^Plot Name:\\s*", "", ., ignore.case = TRUE)

  plot_code <- metadata_row[grepl("^Plotcode:", metadata_row, ignore.case = TRUE)] %>%
    sub("^Plotcode:\\s*", "", ., ignore.case = TRUE)

  # Extract header from the second row
  header <- raw[2, ] %>% unlist() %>% as.character()
  header[is.na(header) | header == ""] <- paste0("NA_col_", seq_along(header))[is.na(header) | header == ""]

  # Data is from row 3 onwards
  data <- raw[-c(1,2), ]
  colnames(data) <- make.unique(header)

  # Remove columns with NA or empty header names (optional)
  data <- data[, !is.na(colnames(data)) & colnames(data) != ""]
  data$Collected <- gsub("^$", NA, data$Collected)

  # Clean data and compute coordinates
  fp_clean <- .clean_fp_data(data, subplot_size)
  fp_coords <- .compute_global_coordinates(fp_clean, plot_size, subplot_size)
  fp_coords <- fp_coords %>% mutate(diameter = (D / 100) * 2)

  # Calculate exact subplot centers based on subplot_size and subplot number (T1)
  subplot_labels <- tibble(T1 = sort(unique(fp_coords$T1))) %>%
    mutate(
      n_subplots_per_col = 100 / subplot_size,  # number of rows (assumed square plot)
      col_index = (T1 - 1) %/% n_subplots_per_col,
      row_index = (T1 - 1) %% n_subplots_per_col,
      subplot_y = if_else(col_index %% 2 == 0,
                          row_index * subplot_size,
                          (n_subplots_per_col - 1 - row_index) * subplot_size),
      subplot_x = col_index * subplot_size,

      center_x = subplot_x + subplot_size / 2,
      center_y = subplot_y + subplot_size / 2
    ) %>%
    select(T1, center_x, center_y)


# Base plot
  base_plot <- ggplot(fp_coords, aes(x = global_x, y = global_y)) +
    geom_vline(xintercept = seq(0, 100, subplot_size), color = "gray50", linewidth = 0.2) +
    geom_hline(yintercept = seq(0, 100, subplot_size), color = "gray50", linewidth = 0.2) +
    geom_rect(aes(xmin = 0, xmax = 100, ymin = 0, ymax = 100),
              fill = NA, color = "black", linewidth = 0.6) +

    geom_text(
      data = subplot_labels,
      aes(x = center_x, y = center_y, label = T1),
      color = "gray",
      size = 2,
      fontface = "bold",
      inherit.aes = FALSE
    ) +
    geom_point(aes(
      size = diameter,
      fill = factor(case_when(
        highlight_palms & Family == "Arecaceae" ~ "Palms",
        !is.na(Collected) & Collected != "" ~ "Collected",
        TRUE ~ "Uncollected"
      ), levels = c("Collected", "Uncollected", "Palms"))
    ),
    shape = 21,
    stroke = 0.2,
    color = "black",
    alpha = 0.9
    ) +
    geom_text(aes(label = `New Tag No`),
              vjust = 0.5,
              hjust = 0.5,
              size = 0.6) +
    scale_x_continuous(limits = c(0, 100), breaks = seq(0, 100, by = subplot_size)) +
    scale_y_continuous(limits = c(0, 100), breaks = seq(0, 100, by = subplot_size)) +
    coord_fixed(ratio = 1, clip = "off") +
    scale_fill_manual(
      values = c("Collected" = "gray", "Uncollected" = "red", "Palms" = "gold"),
      name = "Status"
    ) +
    scale_size_continuous(range = c(2, 6), guide = "none") +
    labs(
      x = "X (m)", y = "Y (m)",
      title = paste0("Collection Balance ", plot_name),
      subtitle = paste0("Plot Code: ", plot_code)
    ) +
    theme_bw() +
    theme(legend.position = "right")


  # Create a temporary .Rmd file path to save the report
  rmd_path <- tempfile(fileext = ".Rmd")

  # Define the final PDF output path
  final_pdf <- file.path(foldername, paste0(filename, "_full_report.pdf"))

  # Filter data for collected specimens (excluding family Arecaceae)
  tf_col <- !is.na(fp_coords$Collected)
  if (any(tf_col)) {
    filtered_collected_data <- fp_coords %>%
      filter(!is.na(Collected) & Family != "Arecaceae")

    # Create collected plot based on base_plot with filtered data
    collected_plot <- base_plot %+% filtered_collected_data
  }

  # Filter data for not collected specimens (excluding family Arecaceae)
  tf_uncol <- is.na(fp_coords$Collected)
  if (any(tf_uncol)) {
    filtered_uncollected_data <- fp_coords %>%
      filter((is.na(Collected)) & Family != "Arecaceae")

    # Create not collected plot based on base_plot with filtered data
    uncollected_plot <- base_plot %+% filtered_uncollected_data
  }

  # Filter data for not collected palm specimens
  tf_palm <- fp_coords$Family %in% "Arecaceae"
  if (any(tf_palm)) {
    filtered_uncollected_palm_data <- fp_coords %>%
      filter((is.na(Collected)) & Family == "Arecaceae")

    # Create not collected plot for palms based on base_plot with filtered data
    uncollected_palm_plot <- base_plot %+% filtered_uncollected_palm_data
  }

  # Generate subplot plots for each unique subplot (T1)
  subplot_plots <- list()
  unique_subplots <- sort(unique(fp_coords$T1))

  for (sp in unique_subplots) {
    sp_data <- fp_coords %>% filter(T1 == sp)

    # Build ggplot for each subplot
    p <- ggplot(sp_data, aes(x = X, y = Y)) +
      geom_rect(aes(xmin = 0, xmax = subplot_size, ymin = 0, ymax = subplot_size),
                fill = NA, color = "black", linewidth = 0.6) +
      geom_point(aes(
        size = diameter,
        fill = case_when(
          highlight_palms & Family == "Arecaceae" ~ "Palms",
          !is.na(Collected) & Collected != "" ~ "Collected",
          TRUE ~ "Uncollected"
        )
      ),
      shape = 21, color = "black", alpha = 0.9, show.legend = c(size = FALSE)) +
      scale_fill_manual(
        values = c("Collected" = "gray", "Uncollected" = "red", "Palms" = "gold"),
        name = "Status"
      ) +
      scale_size_continuous(range = c(4, 10)) +
      geom_text(aes(label = `New Tag No`), size = 1.5) +
      labs(
        title = paste("Subplot", sp),
        subtitle = paste("Plot Name:", plot_name),
        caption = paste("Team:", team),
        x = "X (m)", y = "Y (m)"
      ) +
      coord_fixed() +
      theme_bw() +
      theme(legend.position = "right")

    # Add the subplot ggplot object to the list
    subplot_plots[[length(subplot_plots) + 1]] <- p
  }


  # Calculate plot statistics
  total_specimens <- nrow(fp_coords)
  collected_count <- sum(!is.na(fp_coords$Collected) & fp_coords$Collected != "" & fp_coords$Family != "Arecaceae")
  uncollected_count <- sum((is.na(fp_coords$Collected) | fp_coords$Collected == "") & fp_coords$Family != "Arecaceae")
  palms_count <- sum(fp_coords$Family == "Arecaceae")

  # Prepare RMarkdown sections for each subplot with navigation and page breaks
  rmd_content <- .create_rmd_content(subplot_plots, tf_col, tf_uncol, tf_palm, plot_name, plot_code)

  # Write Rmd content to file
  writeLines(rmd_content, rmd_path)

options(tinytex.pdflatex.args = "--no-crop")

  # Render the Rmd file to PDF using rmarkdown::render with params
  invisible(
    capture.output(
      suppressMessages(
        suppressWarnings(
          rmarkdown::render(
            input = rmd_path,
            output_file = basename(final_pdf),
            output_dir = foldername,
            params = list(
              metadata = list(plot_name = plot_name, plot_code = plot_code, team = team),
              main_plot = base_plot,
              collected_plot = collected_plot,
              uncollected_plot = uncollected_plot,
              uncollected_palm_plot = uncollected_palm_plot,
              subplots_list = subplot_plots,
              subplot_size = subplot_size,
              stats = list(
                total = total_specimens,
                collected = collected_count,
                uncollected = uncollected_count,
                palms = palms_count
              )
            ),
            envir = new.env(parent = globalenv())
          )
        )
      ),
      type = "output"
    )
  )

  # Cleanup temporary files/folders created by knitting
  unlink(rmd_path)
  unlink(file.path(foldername, paste0(tools::file_path_sans_ext(basename(final_pdf)), "_files")), recursive = TRUE)

  cat(paste0("✅ Full report saved to: ", final_pdf, "\n"))


  # Generate collection summary table
  .collection_percentual(
    fp_sheet = fp_coords,
    dir = foldername,
    plot_name = plot_name,
    plot_code = plot_code,
    team = team
  )
}


# Data Clean
.clean_fp_data <- function(fp_sheet, subplot_size) {
  fp_sheet %>%
    mutate(

      # Convert to numeric, replace NA or invalid entries with 0
      X = suppressWarnings(as.numeric(X)),
      X = if_else(is.na(X), 0, X),
      Y = suppressWarnings(as.numeric(Y)),
      Y = if_else(is.na(Y), 0, Y),
      D = suppressWarnings(as.numeric(D)),
      D = if_else(is.na(D), 0, D),
      T1 = suppressWarnings(as.numeric(T1)),
      T1 = if_else(is.na(T1), 0, T1),

      # Ensure Collected is treated as character
      Collected = as.character(Collected)
    )
}


# Compute Global Coordinates
.compute_global_coordinates <- function(fp_clean, plot_size, subplot_size) {
  max_coord <- plot_size * 100
  n_rows <- floor(max_coord / subplot_size)

  fp_clean %>%
    mutate(
      T1 = as.numeric(T1),
      X = as.numeric(X),
      Y = as.numeric(Y),
      col = floor((T1 - 1) / n_rows),
      row = (T1 - 1) %% n_rows,
      global_x = col * subplot_size + X,
      global_y = if_else(
        col %% 2 == 0,
        row * subplot_size + Y,
        (n_rows - row - 1) * subplot_size + Y
      )
    ) %>%
    filter(
      global_x < max_coord,
      global_y < max_coord,
      global_x >= 0,
      global_y >= 0
    )
}


#
.collection_percentual <- function(fp_sheet, dir = getwd(),
                                   plot_name = plot_name, plot_code = plot_name,
                                   team = "") {
  if (!dir.exists(dir)) dir.create(dir, recursive = TRUE)

  output_file <- file.path(dir, "collection_balance.xlsx")

  # Clean Family column
  fp_sheet <- fp_sheet %>%
    mutate(
      Family = trimws(Family),
      Family = ifelse(is.na(Family) | Family == "", "Indet", Family)
    )

  # Summary by subplot
  resume <- fp_sheet %>%
    group_by(subplot = as.character(T1)) %>%
    summarise(
      total_individuals = n(),
      total_non_arecaceae = sum(Family != "Arecaceae"),
      collected = sum(!is.na(Collected) & Collected != ""),
      uncollected_non_arecaceae = sum((is.na(Collected) | Collected == "") & Family != "Arecaceae"),
      arecaceae_count = sum(Family == "Arecaceae"),
      collected_percentual = round(100 * collected / total_individuals, 1),
      collected_percentual_without_arecaceae = round(
        100 * sum(!is.na(Collected) & Collected != "" & Family != "Arecaceae") /
          total_non_arecaceae, 1
      ),
      .groups = "drop"
    )

  total <- resume %>%
    summarise(
      subplot = "TOTAL",
      total_individuals = sum(total_individuals),
      total_non_arecaceae = sum(total_non_arecaceae),
      collected = sum(collected),
      uncollected_non_arecaceae = sum(uncollected_non_arecaceae),
      arecaceae_count = sum(arecaceae_count),

    ) %>%
    mutate(
      collected_percentual = round(100 * collected / total_individuals, 1),
      collected_percentual_without_arecaceae = round(
        100 * sum(!is.na(fp_sheet$Collected) & fp_sheet$Collected != "" & fp_sheet$Family != "Arecaceae") /
          sum(fp_sheet$Family != "Arecaceae"), 1
      )
    )

  final_resume <- bind_rows(resume, total)

  # Not collected (non-palms only)
  uncollected_df <- fp_sheet %>%
    filter((is.na(Collected) | Collected == "") & Family != "Arecaceae") %>%
    group_by(subplot = as.character(T1)) %>%
    summarise(
      n_uncollected = n(),
      tagno_uncollected = paste(sort(unique(`New Tag No`)), collapse = "|"),
      .groups = "drop"
    )

  total_uncollected <- uncollected_df %>%
    summarise(
      subplot = "TOTAL",
      n_uncollected = sum(n_uncollected),
      tagno_uncollected = paste(
        sort(unique(unlist(strsplit(tagno_uncollected, "\\|")))),
        collapse = "|"
      )
    )

  final_uncollected <- bind_rows(uncollected_df, total_uncollected)

  # Collected
  collected_df <- fp_sheet %>%
    filter(!is.na(Collected) & Collected != "") %>%
    group_by(subplot = as.character(T1)) %>%
    summarise(
      n_collected = n(),
      tagno_collected = paste(sort(unique(`New Tag No`)), collapse = "|"),
      .groups = "drop"
    )

  total_collected <- collected_df %>%
    summarise(
      subplot = "TOTAL",
      n_collected = sum(n_collected),
      tagno_collected = paste(
        sort(unique(unlist(strsplit(tagno_collected, "\\|")))),
        collapse = "|"
      )
    )

  final_collected <- bind_rows(collected_df, total_collected)

  # Create xlsx
  wb <- openxlsx::createWorkbook()

  # Header string
  header_string <- paste("Plot Name:", plot_name,
                         "| Plot Code:", plot_code,
                         "| Team:", team)

  # Add each sheet with metadata row
  openxlsx::addWorksheet(wb, "COLLECTION_PERCENTUAL")
  openxlsx::writeData(wb, "COLLECTION_PERCENTUAL", header_string, startRow = 1, colNames = FALSE)
  openxlsx::writeData(wb, "COLLECTION_PERCENTUAL", final_resume, startRow = 3)

  openxlsx::addWorksheet(wb, "NOT_COLLECTED")
  openxlsx::writeData(wb, "NOT_COLLECTED", header_string, startRow = 1, colNames = FALSE)
  openxlsx::writeData(wb, "NOT_COLLECTED", final_uncollected, startRow = 3)

  openxlsx::addWorksheet(wb, "COLLECTED")
  openxlsx::writeData(wb, "COLLECTED", header_string, startRow = 1, colNames = FALSE)
  openxlsx::writeData(wb, "COLLECTED", final_collected, startRow = 3)

  # Color formatting on percentual columns
  color <- colorRampPalette(c("red", "yellow", "green"))(101)
  for (i in seq_len(nrow(resume))) {
    for (col_name in c("collected_percentual", "collected_percentual_without_arecaceae")) {
      valor <- final_resume[[col_name]][i]
      if (!is.na(valor)) {
        cor_hex <- color[round(valor) + 1]
        style <- openxlsx::createStyle(
          fgFill = cor_hex,
          halign = "CENTER",
          textDecoration = "bold",
          border = "TopBottomLeftRight"
        )
        openxlsx::addStyle(
          wb,
          sheet = "COLLECTION_PERCENTUAL",
          style = style,
          rows = i + 3,
          cols = which(names(final_resume) == col_name),
          gridExpand = FALSE,
          stack = TRUE
        )
      }
    }
  }

  # Save Excel
  openxlsx::saveWorkbook(wb, file = output_file, overwrite = TRUE)
}


#' Prepare RMarkdown sections for each subplot with navigation and page breaks
.create_rmd_content <- function(subplot_plots, tf_col, tf_uncol, tf_palm, plot_name, plot_code) {

  # Compose the RMarkdown document content as a character vector
  rmd_content <- c(
    "---",
    "title: \"ForplotR Full Report\"",
    "subtitle: \"`r paste0(params$metadata$plot_name, ' | Plot Code: ', params$metadata$plot_code)`\"",
    "output:",
    "  pdf_document:",
    "    toc: true",
    "    toc_depth: 2",
    "    number_sections: true",
    "fontsize: 11pt",
    "params:",
    "  metadata: NULL",
    "  main_plot: NULL"
  )

  mid_section <- c(
    "  subplots_list: NULL",
    "  subplot_size: NULL",
    "  stats: NULL",
    "---",
    "# Table of Contents {#contents}",
    "\\newpage",
    "",
    "```{r setup, include=FALSE}",
    "knitr::opts_chunk$set(echo = FALSE, warning = FALSE, message = FALSE)",
    "library(ggplot2)",
    "library(dplyr)",
    "```",
    "",
    "\\newpage",
    "# Metadata {#metadata}",
    "",
    "**Plot Name:** `r params$metadata$plot_name`",
    "",
    "**Plot Code:** `r params$metadata$plot_code`",
    "",
    "**Team:** `r params$metadata$team`",
    "\\",
    "",
    "\\hrule",
    "",
    "# Specimen Counts {#counts}",
    "",
    "- **Total Specimens:** `r params$stats$total`",
    "- **Collected (excluding palms):** `r params$stats$collected`",
    "- **Not Collected (excluding palms):** `r params$stats$uncollected`",
    "- **Palms (Arecaceae):** `r params$stats$palms`",
    "",
    "\\hrule"
  )

  index_section <- c(
    "",
    "\\newpage",
    "",
    "# Subplot Index {#subplot-index}",
    "",
    "```{r toc, results='asis', echo=FALSE}",
    "cols <- 5",
    "n <- length(params$subplots_list)",
    "per_col <- ceiling(n / cols)",
    "header <- paste(rep('Subplots', cols), collapse = ' | ')",
    "separator <- paste(rep('---', cols), collapse = ' | ')",
    "toc_lines <- c(header, separator)",
    "for (i in 1:per_col) {",
    "  row <- character(cols)",
    "  for (j in 0:(cols - 1)) {",
    "    idx <- i + j * per_col",
    "    if (idx <= n) {",
    "      row[j + 1] <- paste0('[Subplot ', idx, '](#subplot-', idx, ')')",
    "    } else {",
    "      row[j + 1] <- ' '",
    "    }",
    "  }",
    "  toc_lines <- c(toc_lines, paste(row, collapse = ' | '))",
    "}",
    "cat(paste(toc_lines, collapse = '\\n'))",
    "```",
    "",
    "\\newpage",
    "# General Plot {#general-plot}",
    "```{r general-plot, fig.width=14, fig.height=11}",
    "print(params$main_plot)",
    "```",
    "[Back to contents](#contents)",
    "",
    "\\newpage"
  )

  col_section <- c(
    "# Collected Only {#collected-only}",
    "```{r collected-only, fig.width=14, fig.height=11}",
    "print(params$collected_plot)",
    "```",
    "[Back to contents](#contents)",
    "",
    "\\newpage"
  )

  uncol_section <- c(
    "# Not Collected {#uncollected}",
    "```{r uncollected, fig.width=14, fig.height=11}",
    "print(params$uncollected_plot)",
    "```",
    "[Back to contents](#contents)",
    "",
    "\\newpage"
  )

  palm_section <- c(
    "# Not Collected Palms {#uncollected-palm}",
    "```{r uncollected-palm, fig.width=14, fig.height=11}",
    "print(params$uncollected_palm_plot)",
    "```",
    "[Back to contents](#contents)",
    "",
    "\\newpage"
  )

  subplot_sections <- unlist(lapply(seq_along(subplot_plots), function(i) {
    c(
      sprintf("\\subsection*{\\footnotesize Subplot %d} \\label{subplot-%d}", i, i),
      "",
      sprintf("```{r subplot-%d, fig.width=12, fig.height=9}", i),
      sprintf("print(params$subplots_list[[%d]])", i),
      "```",
      "",
      "[Back to subplot index](#subplot-index)",
      if (i < length(subplot_plots)) "\\newpage" else NULL,
      ""
    )
  }))

  # Add conditionally the correct sections
  if (any(tf_col) && !any(tf_uncol) && !any(tf_palm)) {
    rmd_content <- c(rmd_content,
                     "  collected_plot: NULL",
                     mid_section,
                     index_section,
                     col_section)
  } else if (!any(tf_col) && any(tf_uncol) && !any(tf_palm)) {
    rmd_content <- c(rmd_content,
                     "  uncollected_plot: NULL",
                     mid_section,
                     index_section,
                     uncol_section)
  } else if (any(tf_col) && any(tf_uncol) && any(tf_palm)) {
    rmd_content <- c(rmd_content,
                     "  collected_plot: NULL",
                     "  uncollected_plot: NULL",
                     "  uncollected_palm_plot: NULL",
                     mid_section,
                     index_section,
                     col_section,
                     uncol_section,
                     palm_section)
  } else if (any(tf_col) && any(tf_uncol) && !any(tf_palm)) {
    rmd_content <- c(rmd_content,
                     "  collected_plot: NULL",
                     "  uncollected_plot: NULL",
                     mid_section,
                     index_section,
                     col_section,
                     uncol_section)
  } else if (any(tf_col) && !any(tf_uncol) && any(tf_palm)) {
    rmd_content <- c(rmd_content,
                     "  collected_plot: NULL",
                     "  uncollected_palm_plot: NULL",
                     mid_section,
                     index_section,
                     col_section,
                     palm_section)
  } else if (!any(tf_col) && any(tf_uncol) && any(tf_palm)) {
    rmd_content <- c(rmd_content,
                     "  uncollected_plot: NULL",
                     "  uncollected_palm_plot: NULL",
                     mid_section,
                     index_section,
                     uncol_section,
                     palm_section)
  }

  # Final section with subplots
  rmd_content <- c(rmd_content,
                   "# Individual Subplots {#individual-subplots}",
                   subplot_sections)

  return(rmd_content)
}
