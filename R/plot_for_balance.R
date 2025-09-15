#' Generate forest plot specimen map and collection balance
#'
#' @author Giulia Ottino & Domingos Cardoso
#'
#' @description Processes field data collected using the \href{https://forestplots.net/}{Forestplots format}
#' and generates a specimen map (2D plot) with collection status and spatial
#' distribution of individuals across subplots. Generate a PDF with a full plot
#' report with separate maps, for each subplot and a spreadsheet summarizing the
#' percentage of collected specimens per subplot. and an optional above-ground
#' biomass (AGB) summary.  The function performs the following steps: (i)
#' validates all input arguments and checks/create output folders; (ii) reads the
#' Forestplots-formatted xlsx sheet, extracts metadata (team, plot name, plot
#' code), and cleans the data; (iii) normalizes and cleans coordinate and
#' diameter values, computing global plot coordinates from subplot-relative
#' positions; (iv) creates a PDF report, including plot metadata, the main map,
#' collected and uncollected specimen maps, optionally a map highlighting palms
#' (Arecaceae) and navigable individual subplot  maps; (v) optionally generates
#' a xlsx spreadsheet with the collection percentage per subplot (including
#' totals), distinguishing collected vs. uncollected, and palm individuals; (vi)
#' optionally, the above-ground biomass (AGB) per plot, computed with BiomasaFP
#' and inserted as an extra section in the report.
#'
#' @usage
#' plot_for_balance(fp_file_path = NULL,
#'                  input_type = c("field_sheet", "fp_query_sheet"),
#'                  plot_size = 1,
#'                  subplot_size = 10,
#'                  highlight_palms = TRUE,
#'                  calc_agb = FALSE,
#'                  trees_csv = NULL,
#'                  wd_csv = NULL,
#'                  md_csv = NULL,
#'                  dir = "Results_map_plot",
#'                  filename = "plot_specimen")
#'
#' @param fp_file_path Path to the Excel file (field or query sheet) in ForestPlots format.
#'
#' @param input_type One of `"field_sheet"` or `"fp_query_sheet"`; format of the input file.
#'
#' @param plot_size Total plot size in hectares (default = 1).
#'
#' @param subplot_size Side length of each subplot in meters (default = 10).
#'
#' @param highlight_palms Logical. If TRUE, highlights Arecaceae in the plots.
#'
#' @param calc_agb Logical. If TRUE, calculates above-ground biomass (AGB) using `BiomasaFP`.
#'
#' @param trees_csv Optional. Tree-level measurement data used for above-ground
#' biomass (AGB) estimation. Can be either a data frame or a file path (character)
#' to the "Data.csv" file downloaded via ForestPlots "Advanced Search".
#'
#' @param md_csv A metadata file downloaded from the "Query Library". Can be a
#' dataframe or a character indicating the file path.
#'
#' @param wd_csv A wood density file, with the wood density for each individual tree
#' by PlotViewID downloaded from the "Query Library". Can be a dataframe or a
#' character indicating the file path.
#'
#' @param dir Directory path where output will be saved
#' (default is "Results_map_plot").
#'
#' @param filename Basename for output files (without extension).
#'
#' @return A PDF report and an Excel file summarizing specimen collection
#' statistics per subplot.
#'
#' @import ggplot2
#' @import BiomasaFP
#' @importFrom readxl read_excel
#' @importFrom dplyr mutate if_else group_by arrange select distinct filter
#' @importFrom magrittr "%>%"
#' @importFrom grDevices pdf dev.off colorRampPalette
#' @importFrom utils capture.output
#' @importFrom rmarkdown render pdf_document
#' @importFrom openxlsx read.xlsx createWorkbook addWorksheet writeData createStyle addStyle saveWorkbook
#' @importFrom tinytex latexmk
#' @importFrom readr locale read_delim
#' @importFrom tibble as_tibble tibble
#' @importFrom here here
#'
#' @examples
#' \dontrun{
#' plot_for_balance(fp_file_path = "data/forestplot.xlsx",
#'                  input_type = c("field_sheet", "fp_query_sheet"),
#'                  plot_size = 1,
#'                  subplot_size = 10,
#'                  highlight_palms = TRUE,
#'                  calc_agb = FALSE,
#'                  trees_csv = NULL,
#'                  md_csv = NULL,
#'                  wd_csv = NULL,
#'                  dir = "Results_map_plot",
#'                  filename = "plot_specimen")
#' }
#'
#' @export
#'

plot_for_balance <- function(fp_file_path = NULL,
                             input_type = c("field_sheet", "fp_query_sheet"),
                             plot_size = 1,
                             subplot_size = 10,
                             highlight_palms = TRUE,
                             calc_agb = FALSE,
                             trees_csv = NULL,
                             wd_csv = NULL,
                             md_csv = NULL,
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

  agb_tbl <- NULL
  if (calc_agb) {
    agb_tbl <- .compute_plot_agb(
      trees = trees_csv,
      wd = wd_csv,
      md = md_csv,
      plot_area = plot_size
    )
  }

  if (input_type == "fp_query_sheet") {
    raw <- .fp_query_to_field_sheet_df(fp_file_path)
  }
  if (input_type == "field_sheet") {
    raw <- suppressMessages(openxlsx::read.xlsx(fp_file_path, sheet = 1, colNames = FALSE))
  }

  # Extract metadata from the first row
  metadata_row <- raw[1, ] %>% unlist() %>% as.character()

  team <- metadata_row[grepl("^Team:", metadata_row, ignore.case = TRUE)] %>%
    sub("^Team:\\s*", "", ., ignore.case = TRUE)

  plot_name <- metadata_row[grepl("^Plot Name:", metadata_row, ignore.case = TRUE)] %>%
    sub("^Plot Name:\\s*", "", ., ignore.case = TRUE)

  plot_code <- metadata_row[grepl("^Plotcode:", metadata_row, ignore.case = TRUE)] %>%
    sub("^Plotcode:\\s*", "", ., ignore.case = TRUE)
  if (!grepl("-", plot_code)) {
    plot_code <- gsub("(?<=[A-Z])(?=\\d)", "-", plot_code, perl = TRUE)
  }

  # Extract header from the second row
  header <- raw[2, ] %>% unlist() %>% as.character()
  header[is.na(header) | header == ""] <- paste0("NA_col_", seq_along(header))[is.na(header) | header == ""]

  # Data is from row 3 onwards
  data <- raw[-c(1,2), ]
  colnames(data) <- make.unique(header)

  # Remove columns with NA or empty header names (optional)
  data <- data[, !is.na(colnames(data)) & colnames(data) != ""]
  data$Collected <- gsub("^$", NA, data$Collected)

  data <- .replace_empty_with_na(data)

  # Clean data and compute coordinates
  fp_clean <- .clean_fp_data(data)
  fp_coords <- .compute_global_coordinates(fp_clean, plot_size, subplot_size)
  fp_coords <- fp_coords %>% dplyr::mutate(diameter = (D / 100) * 2)
  fp_coords <- fp_coords %>%
    dplyr::mutate(`New Tag No` = trimws(as.character(`New Tag No`)))

  # Build species check-list
  spec_df <- fp_coords %>%
    dplyr::mutate(
      Family = dplyr::if_else(is.na(Family) | Family == "", "Indet", Family),
      Species = dplyr::if_else(is.na(`Original determination`) |
                                 `Original determination` == "",
                               "indet", `Original determination`),
      Species_fmt  = paste0("*", Species, "*")
    ) %>%
    dplyr::group_by(Family, Species_fmt) %>%
    dplyr::summarise(
      tags = paste(sort(`New Tag No`), collapse = " | "),
      tag_vec = list(sort(trimws(as.character(`New Tag No`)))),
      .groups = "drop"
    ) %>%
    dplyr::arrange(dplyr::if_else(Family == "Indet", 0, 1), Family)

  # Calculate exact subplot centers based on subplot_size and subplot number (T1)
  subplot_labels <- tibble::tibble(T1 = sort(unique(fp_coords$T1))) %>%
    dplyr::mutate(
      n_subplots_per_col = (plot_size / 1) * 100 / subplot_size,
      col_index = (T1 - 1) %/% n_subplots_per_col,
      row_index = (T1 - 1) %% n_subplots_per_col,
      subplot_y = if_else(col_index %% 2 == 0,
                          row_index * subplot_size,
                          (n_subplots_per_col - 1 - row_index) * subplot_size),
      subplot_x = col_index * subplot_size,

      center_x = subplot_x + subplot_size / 2,
      center_y = subplot_y + subplot_size / 2
    ) %>%
    dplyr::select(T1, center_x, center_y)


  # Base plot

  max_x <- 100
  max_y <- (plot_size / 1) * 100

  base_plot <- ggplot2::ggplot(fp_coords, ggplot2::aes(x = global_x, y = global_y)) +
    ggplot2::geom_vline(xintercept = seq(0, max_x, subplot_size), color = "gray50", linewidth = 0.2) +
    ggplot2::geom_hline(yintercept = seq(0, max_y, subplot_size), color = "gray50", linewidth = 0.2) +
    ggplot2::geom_rect(ggplot2::aes(xmin = 0, xmax = max_x, ymin = 0, ymax = max_y),
                       fill = NA, color = "black", linewidth = 0.6) +
    ggplot2::geom_text(
      data = subplot_labels,
      ggplot2::aes(x = center_x, y = center_y, label = T1),
      color = "gray",
      size = 2,
      fontface = "bold",
      inherit.aes = FALSE
    ) +
    ggplot2::geom_point(aes(
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
    ggplot2::geom_text(ggplot2::aes(label = `New Tag No`),
                       vjust = 0.5,
                       hjust = 0.5,
                       size = 0.6) +
    ggplot2::scale_x_continuous(limits = c(0, max_x), breaks = seq(0, 100, by = subplot_size)) +
    ggplot2::scale_y_continuous(limits = c(0, max_y), breaks = seq(0, 100, by = subplot_size)) +
    ggplot2::coord_fixed(ratio = 1, clip = "off") +
    ggplot2::scale_fill_manual(
      values = c("Collected" = "gray", "Uncollected" = "red", "Palms" = "gold"),
      name = "Status"
    ) +
    ggplot2::scale_size_continuous(range = c(2, 6), guide = "none") +
    ggplot2::labs(
      x = "X (m)", y = "Y (m)",
      title = paste0("Collection Balance ", plot_name),
      subtitle = paste0("Plot Code: ", plot_code)
    ) +
    ggplot2::theme_bw() +
    ggplot2::theme(legend.position = "right") +
    ggplot2::guides(
      fill = guide_legend(override.aes = list(size = 5)),  # larger circles in Status legend
    )


  # Create a temporary .Rmd file path to save the report
  rmd_path  <- file.path(foldername, paste0(filename, "_temp.Rmd"))

  # Define the final PDF output path
  final_pdf <- file.path(foldername, paste0(filename, "_full_report.pdf"))

  collected_plot <- NULL
  uncollected_plot <- NULL
  uncollected_palm_plot <- NULL

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
  for (i in sort(unique(fp_coords$T1))) {
    sp_data <- fp_coords %>% filter(T1 == i)

    # Assign status for fill color
    sp_data <- sp_data %>%
      mutate(
        Status = case_when(
          highlight_palms & Family == "Arecaceae" ~ "Palms",
          !is.na(Collected) & Collected != "" ~ "Collected",
          TRUE ~ "Uncollected"
        )
      )

    # Compute diameter breaks
    sp_data$diameter <- sp_data$D / 10
    # Extract valid diameter values
    valid_d <- sp_data$diameter[!is.na(sp_data$diameter)]
    min_d <- min(valid_d)
    max_d <- max(valid_d)
    # Create 4 evenly spaced values: min, 1/3, 2/3, max
    breaks_d <- round(seq(min_d, max_d, length.out = 4), 0)
    breaks_d[1] <- min_d
    breaks_d[4] <- max_d

    # Build ggplot for each subplot
    p <- ggplot(sp_data, ggplot2::aes(x = X, y = Y)) +
      ggplot2::geom_rect(ggplot2::aes(xmin = 0, xmax = subplot_size, ymin = 0, ymax = subplot_size),
                         fill = NA, color = "black", linewidth = 0.6) +
      ggplot2::geom_point(
        aes(size = diameter, fill = Status),
        shape = 21, color = "black", stroke = 0.2, alpha = 0.9
      ) +

      #label every point with its tag number
      ggplot2::geom_text(
        aes(label = `New Tag No`),
        size = 1.7
      ) +
      ggplot2::scale_fill_manual(
        values = c("Collected" = "gray", "Uncollected" = "red", "Palms" = "gold"),
        name = "Status"
      ) +
      ggplot2::scale_size_area(
        name = "DBH (cm)",
        breaks = breaks_d,
        labels = breaks_d,
        limits = range(breaks_d),  # ← force legend to span all breaks
        max_size = 10,
        guide = "legend"
      ) +
      ggplot2::labs(
        title = paste("Plot Name:", plot_name),
        subtitle = paste0("Plot Code: ", plot_code),
        #caption = paste("Team:", team),
        x = "X (m)", y = "Y (m)"
      ) +
      coord_fixed() +
      theme_bw() +
      theme(
        panel.grid.minor = element_blank(),       # hides minor grid lines
        panel.grid.major = element_line(size = 0.3, color = "gray80")  # keeps major grid
      ) +
      theme(legend.position = "right") +
      ggplot2::guides(
        fill = guide_legend(override.aes = list(size = 5), order = 1),  # larger circles in Status legend
        size = guide_legend(title = "DBH (cm)", order = 2) # keep your diameter scale
      )

    # Add the subplot ggplot object to the list
    subplot_plots[[length(subplot_plots) + 1]] <- list(
      plot = p,
      data = sp_data  # must include tag info
    )
  }

  # Calculate plot statistics
  total_specimens <- nrow(fp_coords)
  collected_count <- sum(!is.na(fp_coords$Collected) & fp_coords$Family != "Arecaceae")
  uncollected_count <- sum(is.na(fp_coords$Collected) & fp_coords$Family != "Arecaceae")
  palms_count <- sum(fp_coords$Family == "Arecaceae")

  # Map tag → subplot
  tag_to_subplot <- fp_coords %>%
    dplyr::select(`New Tag No`, T1) %>%
    dplyr::distinct()

  # Prepare RMarkdown sections for each subplot with navigation and page breaks
  rmd_content <- .create_rmd_content(
    subplot_plots,
    tf_col, tf_uncol, tf_palm,
    plot_name, plot_code, spec_df,
    has_agb = !is.null(agb_tbl) #  <- new flag
  )

  # Write Rmd content to file
  writeLines(c(rmd_content, ""), rmd_path)

  options(tinytex.pdflatex.args = "--no-crop")

  param_list <- list(
    metadata = list(plot_name = plot_name,
                    plot_code = plot_code,
                    team = team),
    main_plot = base_plot,
    subplots_list = subplot_plots,
    subplot_size  = subplot_size,
    stats = list(
      total = total_specimens,
      collected = collected_count,
      uncollected = uncollected_count,
      palms = palms_count,
      spec_df = spec_df),
    tag_list = unique(trimws(as.character(fp_coords$`New Tag No`))),
    tag_to_subplot = tag_to_subplot
  )

  # Append the AGB on top, if calculated
  if (!is.null(agb_tbl)) {
    param_list$agb <- agb_tbl
  }
  if (!is.null(collected_plot)) {
    param_list$collected_plot <- collected_plot
  }
  if (!is.null(uncollected_plot)) {
    param_list$uncollected_plot <- uncollected_plot
  }
  if (!is.null(uncollected_palm_plot)) {
    param_list$uncollected_palm_plot <- uncollected_palm_plot
  }

  # Render the Rmd file to PDF using rmarkdown::render with params
  .render_plot_report(
    rmd_path = rmd_path,
    output_pdf_path = foldername,
    output_pdf_name = final_pdf,
    params = param_list
  )

  # Cleanup temporary files/folders created by knitting
  unlink(rmd_path)
  unlink(gsub("_temp[.]Rmd", "_temp.knit.md", rmd_path))

  # Generate collection summary table
  .collection_percentual(
    fp_sheet = fp_coords,
    dir = foldername,
    plot_name = plot_name,
    plot_code = plot_code,
    team = team
  )

}


# Rendering PDF with full plot report ####
.render_plot_report <- function(
    rmd_path,
    output_pdf_path,
    output_pdf_name,
    params) {
  render_env <- new.env(parent = globalenv())

  # Temp output name
  temp_pdf <- tempfile(fileext = ".pdf")

  rmarkdown::render(
    input = rmd_path,
    output_file = basename(temp_pdf),
    output_dir = tempdir(),
    params = params,
    envir = render_env,
    clean = FALSE,
    output_format = rmarkdown::pdf_document(
      keep_tex = TRUE,
      latex_engine = "pdflatex",
      fig_caption = TRUE)
  )

  # Write the LaTeX preamble commands
  preamble_lines <- c(
    "\\makeatletter",
    "\\providecommand{\\vcenteredhbox}[1]{\\begingroup",
    "  \\setbox0=\\hbox{#1}\\parbox{\\wd0}{\\box0}\\endgroup}",
    "\\renewcommand{\\@dotsep}{4}",
    "\\renewcommand{\\contentsname}{\\hspace*{-1cm}}",
    "\\makeatother"
  )

  # Locate .tex file
  tex_candidates <- list.files(tempdir(), pattern = "\\.tex$", full.names = TRUE)
  tex_file <- tex_candidates[which.max(file.info(tex_candidates)$mtime)]

  # Recompile .tex
  if (file.exists(tex_file)) {
    # Insert preamble before \begin{document}
    tex_lines <- readLines(tex_file)
    insert_idx <- grep("\\\\begin\\{document\\}", tex_lines)[1]
    new_tex_lines <- append(tex_lines, preamble_lines, after = insert_idx - 1)
    writeLines(new_tex_lines, tex_file)

    tinytex::latexmk(tex_file)
  }

  # Move output PDF to final location
  if (!dir.exists(output_pdf_path)) dir.create(output_pdf_path, recursive = TRUE)
  final_pdf_path <- file.path(output_pdf_path, basename(output_pdf_name))
  file.copy(sub("\\.tex$", ".pdf", tex_file), final_pdf_path, overwrite = TRUE)

  if (file.exists(final_pdf_path)) {
    message("✅ Full PDF report saved to: ", final_pdf_path)
  } else {
    warning("⚠️ PDF generation failed.")
  }

}


# Replace empty cells with NA ####
.replace_empty_with_na <- function(df) {
  df[] <- lapply(df, function(col) {
    if (is.character(col)) {
      col[col == ""] <- NA
    }
    col
  })
  df
}


# Data Clean ####
.clean_fp_data <- function(fp_sheet) {
  fp_sheet %>%
    dplyr::mutate(
      dplyr::across(
        .cols = any_of(c("X", "Y", "D", "T1")),
        .fns = ~ tidyr::replace_na(suppressWarnings(as.numeric(.)), 0)
      ),
      Collected = as.character(Collected)
    )
}


# Compute Global Coordinates ####
.compute_global_coordinates <- function(fp_clean, plot_size, subplot_size) {
  max_x <- 100
  max_y <- (plot_size / 1) * 100
  n_rows <- floor(max_y / subplot_size)
  fp_clean %>%
    dplyr::mutate(
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
    dplyr::filter(
      global_x <= max_x,
      global_y <= max_y,
      global_x >= 0,
      global_y >= 0
    )
}


# Get percentual values ####
.collection_percentual <- function(fp_sheet, dir = getwd(),
                                   plot_name = plot_name, plot_code = plot_name,
                                   team = "") {
  if (!dir.exists(dir)) dir.create(dir, recursive = TRUE)

  output_file <- file.path(dir, "collection_balance.xlsx")

  # Clean Family column
  fp_sheet <- fp_sheet %>%
    dplyr::mutate(
      Family = trimws(Family),
      Family = ifelse(is.na(Family) | Family == "", "Indet", Family)
    )

  # Summary by subplot
  resume <- fp_sheet %>%
    dplyr::group_by(subplot = as.character(T1)) %>%
    dplyr::summarise(
      total_individuals = n(),
      total_non_arecaceae = sum(Family != "Arecaceae"),
      collected = sum(!is.na(Collected) & Collected != ""),
      uncollected_non_arecaceae = sum((is.na(Collected) | Collected == "") & Family != "Arecaceae"),
      arecaceae_count = sum(Family == "Arecaceae"),
      collected_percentual = round(100 * collected / total_individuals, 1),
      collected_percentual_without_arecaceae = round(
        100 * sum(!is.na(Collected) & Collected != "" & Family != "Arecaceae") /
          total_non_arecaceae, 1), .groups = "drop")
  resume$subplot <- as.numeric(resume$subplot)
  resume <- resume %>% arrange(subplot)
  resume$subplot <- as.character(resume$subplot)

  total <- resume %>%
    dplyr::summarise(
      subplot = "TOTAL",
      total_individuals = sum(total_individuals),
      total_non_arecaceae = sum(total_non_arecaceae),
      collected = sum(collected),
      uncollected_non_arecaceae = sum(uncollected_non_arecaceae),
      arecaceae_count = sum(arecaceae_count)) %>%
    dplyr::mutate(
      collected_percentual = round(100 * collected / total_individuals, 1),
      collected_percentual_without_arecaceae = round(
        100 * sum(!is.na(fp_sheet$Collected) & fp_sheet$Collected != "" & fp_sheet$Family != "Arecaceae") /
          sum(fp_sheet$Family != "Arecaceae"), 1))

  final_resume <- bind_rows(resume, total)

  # Not collected (non-palms only)
  uncollected_df <- fp_sheet %>%
    dplyr::filter((is.na(Collected) | Collected == "") & Family != "Arecaceae") %>%
    dplyr::group_by(subplot = as.character(T1)) %>%
    dplyr::summarise(
      n_uncollected = n(),
      tagno_uncollected = paste(as.character(sort(as.numeric(unique(`New Tag No`)))), collapse = "|"),
      .groups = "drop"
    )
  uncollected_df$subplot <- as.numeric(uncollected_df$subplot)
  uncollected_df <- uncollected_df %>% arrange(subplot)
  uncollected_df$subplot <- as.character(uncollected_df$subplot)


  total_uncollected <- uncollected_df %>%
    dplyr::summarise(
      subplot = "TOTAL",
      n_uncollected = sum(n_uncollected),
      tagno_uncollected = paste(
        sort(unique(unlist(strsplit(tagno_uncollected, "\\|")))),
        collapse = "|"))
  #total_uncollected$tagno_uncollected <- NA
  final_uncollected <- bind_rows(uncollected_df, total_uncollected)

  # Collected
  collected_df <- fp_sheet %>%
    dplyr::filter(!is.na(Collected) & Collected != "") %>%
    dplyr::group_by(subplot = as.character(T1)) %>%
    dplyr::summarise(
      n_collected = n(),
      tagno_collected = paste(as.character(sort(as.numeric(unique(`New Tag No`)))), collapse = "|"),
      .groups = "drop")
  collected_df$subplot <- as.numeric(collected_df$subplot)
  collected_df <- collected_df %>% arrange(subplot)
  collected_df$subplot <- as.character(collected_df$subplot)

  total_collected <- collected_df %>%
    dplyr::summarise(
      subplot = "TOTAL",
      n_collected = sum(n_collected),
      tagno_collected = paste(
        sort(unique(unlist(strsplit(tagno_collected, "\\|")))),
        collapse = "|"))
  #total_collected$tagno_collected <- NA
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
  color <- grDevices::colorRampPalette(c("red", "yellow", "green"))(101)
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

  # Save Excel sheet
  openxlsx::saveWorkbook(wb, file = output_file, overwrite = TRUE)
}


# Prepare RMarkdown sections for each subplot with navigation and page breaks ####
.create_rmd_content <- function (subplot_plots, tf_col, tf_uncol, tf_palm,
                                 plot_name, plot_code, spec_df,
                                 has_agb = FALSE) {   # <‑ new argument ↑

  # -------------------------------------------------------------------------
  # 1) YAML HEADER – built dynamically.  **agb** is included ONLY when
  #    `has_agb` is TRUE (i.e., when calc_agb = TRUE in the main function).
  # -------------------------------------------------------------------------
  yaml_head <- c(
    "---",
    "output:",
    "  pdf_document:",
    "    toc: true",
    "    toc_depth: 2",
    "    number_sections: true",
    "fontsize: 11pt",
    "params:",
    "  metadata: NULL",
    "  main_plot: NULL",   # ← lines always present
    "  tag_list: NULL",
    "  tag_to_subplot: NULL"
  )

  if (has_agb) yaml_head <- c(yaml_head, "  agb: NULL")
  if (any(tf_col)) yaml_head <- c(yaml_head, "  collected_plot: NULL")
  if (any(tf_uncol)) yaml_head <- c(yaml_head, "  uncollected_plot: NULL")
  if (any(tf_palm)) yaml_head <- c(yaml_head, "  uncollected_palm_plot: NULL")

  yaml_head <- c(yaml_head,
                 "  subplots_list: NULL",
                 "  subplot_size: NULL",
                 "  stats: NULL",
                 "---"  # closes the YAML
  )

  # 2) COMMON BLOCK
  mid_section <- c(

    # -- SETUP --
    "```{r setup, include=FALSE}",
    "knitr::opts_chunk$set(echo = FALSE, warning = FALSE, message = FALSE)",
    "library(ggplot2)",
    "library(dplyr)",
    "spec_df <- params$stats$spec_df",
    "safe_id <- function(x) {",
    "  x <- trimws(as.character(x))",
    "  x <- gsub(\"\\\\s*\\\\(.*\\\\)$\", \"\", x)",  # remove ' (SXX)' if present
    "  x <- gsub(\"[^A-Za-z0-9]\", \"\", x)",   # remove any punctuation
    "  tolower(x)",  # makes it case-insensitive
    "}",
    "```",
    "",
    # -- INVISIBLE ANCHORS FOR TAGS --
    "```{r all-tag-anchors, results='asis', echo=FALSE}",
    "t2s_df <- params$tag_to_subplot",
    "tag_vec <- as.character(t2s_df$`New Tag No`)",
    "tag_vec <- trimws(tag_vec)",
    "tag_vec <- tag_vec[!is.na(tag_vec) & nzchar(tag_vec)]",
    "subplot_vec <- t2s_df$T1[match(tag_vec, t2s_df$`New Tag No`)]",
    "ids <- sprintf('tag-%s-S%s', tag_vec, subplot_vec)",
    "ids <- unique(ids)",
    "for (id in ids) cat(sprintf('\\\\hypertarget{%s}{}\\n', id))",
    "```",
    "",

    # ---------- TITLE BLOCK ----------
    "```{r title-block, echo=FALSE, results='asis'}",
    "cat('\\\\begin{center}')",
    "cat('\\\\Huge\\\\textbf{Full Plot Report} \\\\\\\\')",  # line break
    "cat('\\\\vspace{0.5em}')",
    "cat(paste0('\\\\normalsize ', params$metadata$plot_name, ' | ', params$metadata$plot_code, ' \\\\\\\\'))",
    "cat('\\\\end{center}')",
    "```",
    "",
    # -- LOGOS --
    "```{r logo, echo=FALSE, results='asis'}",
    "forplotr_logo_path    <- system.file('figures', 'forplotR_hex_sticker.png', package = 'forplotR')",
    "forestplots_logo_path <- system.file('figures', 'forestplotsnet_logo.png', package = 'forplotR')",
    "cat('\\\\vspace*{1cm}\\n')",
    "cat('\\\\begin{center}\\n')",
    "cat(sprintf('\\\\href{https://dboslab.github.io/forplotR-website/}{\\\\includegraphics[width=0.20\\\\linewidth]{%s}}\\n', forplotr_logo_path))",
    "cat('\\\\end{center}')",
    "cat('\\\\begin{center}')",
    "cat(sprintf('\\\\href{https://forestplots.net}{\\\\includegraphics[width=0.40\\\\linewidth]{%s}}\\n', forestplots_logo_path))",
    "cat('\\\\end{center}')",
    "cat('\\\\vspace*{2cm}\\n')",
    "```",
    "",

    # -- TABLE OF CONTENTS --
    "## Contents {#contents}",
    "\\vspace*{-1cm}",
    "\\thispagestyle{plain}",
    "\\tableofcontents",
    "\\newpage",
    "",

    # -- METADATA --
    "## Metadata {#metadata}",
    "",
    "**Plot Name:** `r params$metadata$plot_name`",
    "",
    "**Plot Code:** `r params$metadata$plot_code`",
    "",
    "**Team:** `r params$metadata$team`",
    "",
    "\\hrule",
    "",

    # -- SPECIMEN COUNTS --
    "## Specimen Counts {#counts}",
    "",
    "- **Total Specimens:** `r params$stats$total`",
    "- **Collected (excluding palms):** `r params$stats$collected`",
    "- **Not Collected (excluding palms):** `r params$stats$uncollected`",
    "- **Palms (Arecaceae):** `r params$stats$palms`",
    "",
    "\\hrule"
  )

  # -- AGB TABLE --
  if (has_agb) {
    mid_section <- c(mid_section,
                     "",
                     "## Above-ground Biomass {#agb}",
                     "",
                     "```{r agb-table, echo=FALSE, results='asis'}",
                     "if (!is.null(params$agb)) {",
                     "  agb_val <- round(params$agb[[2]][1], 2)",
                     "  cat(paste0(\"**Above-ground biomass:** \", agb_val, \" t ha$^{-1}$\\n\\n\"))",
                     "  cat('\\\\hrule\\n')",
                     "}",
                     "```"
    )
  }

  mid_section <- c(mid_section, "\\newpage")

  nav_targets <- c("[Back to Contents](#contents)",
                   "[Metadata](#metadata)",
                   "[Specimen Counts](#counts)")

  if (has_agb) nav_targets <- c(nav_targets, "[AGB](#agb)")
  nav_targets <- c(nav_targets, "[General Plot](#general-plot)")
  if (any(tf_col)) nav_targets <- c(nav_targets, "[Collected Only](#collected-only)")
  if (any(tf_uncol)) nav_targets <- c(nav_targets, "[Not Collected](#uncollected)")
  if (any(tf_palm)) nav_targets <- c(nav_targets, "[Palms](#uncollected-palm)")
  nav_targets <- c(nav_targets, "[Subplot Index](#subplot-index)", "[Checklist](#checklist)")

  nav_links <- c(
    "\\vspace{1em}",
    "\\begingroup\\tiny\\color{gray}",
    paste("«", paste(nav_targets, collapse = " | ")),
    "\\endgroup"
  )

  gencol_section <- c(
    "## General Plot {#general-plot}",
    "```{r general-plot, fig.width=14, fig.height=11}",
    "print(params$main_plot)",
    "```",
    "\n",
    "\n",
    nav_links,
    "",
    "\\newpage"
  )

  col_section <- c(
    "## Collected Only {#collected-only}",
    "```{r collected-only, fig.width=14, fig.height=11}",
    "print(params$collected_plot)",
    "```",
    "\n",
    "\n",
    nav_links,
    "",
    "\\newpage"
  )

  uncol_section <- c(
    "## Not Collected {#uncollected}",
    "```{r uncollected, fig.width=14, fig.height=11}",
    "print(params$uncollected_plot)",
    "```",
    "\n",
    "\n",
    nav_links,
    "",
    "\\newpage"
  )

  palm_section <- c(
    "## Not Collected Palms {#uncollected-palm}",
    "```{r uncollected-palm, fig.width=14, fig.height=11}",
    "print(params$uncollected_palm_plot)",
    "```",
    "\n",
    "\n",
    nav_links,
    "",
    "\\newpage"
  )

  index_section <- c(
    "",
    "\\newpage",
    "",
    "## Subplot Index {#subplot-index}",
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
    "\n",
    "\n",
    nav_links,
    "",
    "\\newpage"
  )

  subplot_sections <- unlist(lapply(seq_along(subplot_plots), function(i) {
    c(
      sprintf("\\hypertarget{subtag-%d}{}", i),

      sprintf("```{r anchors-%d, results='asis', echo=FALSE}", i),
      sprintf("tags_i <- params$subplots_list[[%d]]$data$`New Tag No`", i),
      "tags_i <- trimws(as.character(tags_i))",
      "tags_i <- tags_i[!is.na(tags_i) & nzchar(tags_i)]",
      sprintf("sp_number <- unique(params$subplots_list[[%d]]$data$T1)[1]", i),
      "anchors <- sprintf('\\\\hypertarget{tag-%s-S%s}{}', tags_i, sp_number)",
      "cat(paste(anchors, collapse = '\\n'))",
      "```",
      "",

      # subplot title and label
      sprintf("### Subplot %d {#subplot-%d .unlisted .unnumbered}", i, i),
      "",

      # subplot graphic
      sprintf("```{r subplot-%d, fig.width=12, fig.height=9}", i),
      sprintf("print(params$subplots_list[[%d]]$plot)", i),
      "```",
      "",

      # navigation bar (now also links to the checklist)
      "\\begingroup\\tiny\\color{gray}",
      "« [Back to Contents](#contents) | [Metadata](#metadata) | [Specimen Counts](#counts) | [General Plot](#general-plot) | [Collected Only](#collected-only) | [Not Collected](#uncollected) | [Palms](#uncollected-palm) | [Subplot Index](#subplot-index) | [Checklist](#checklist)",
      "\\endgroup",

      # page break (skip after the last subplot)
      if (i < length(subplot_plots)) "\\newpage" else NULL,
      ""
    )
  }))

  checklist_section <- c(
    "",
    "\\newpage",
    "",
    "## Species Check-list {#checklist}",
    "",
    "```{r checklist, results='asis', echo=FALSE}",
    "library(dplyr)",
    "cat('\\n')",
    "for (fam in unique(spec_df$Family)) {",
    "  cat(\"\\n\\n### \", fam, \"\\n\\n\", sep = \"\")",
    "  fam_df <- spec_df %>% filter(Family == fam)",
    "",
    "  for (i in seq_len(nrow(fam_df))) {",
    "    sp <- fam_df$Species_fmt[i]",
    "    tag_vec <- fam_df$tag_vec[[i]]",
    "",
    "    tag_vec <- as.character(tag_vec)",
    "    tag_vec <- trimws(tag_vec)",
    "    tag_vec <- tag_vec[!is.na(tag_vec) & nzchar(tag_vec)]",
    "    if (length(tag_vec) == 0) next",
    "",
    "    ids <- safe_id(tag_vec)",
    "    t2s_df <- params$tag_to_subplot",
    "    tag_df <- data.frame(",
    "      tag = tag_vec,",
    "      id  = ids,",
    "      subplot = sapply(tag_vec, function(tag) {",
    "         sp <- t2s_df$T1[t2s_df$`New Tag No` == tag]",
    "         if (length(sp) > 0) sp else NA_integer_",
    "      }),",
    "      stringsAsFactors = FALSE)",
    "tag_df <- tag_df[order(tag_df$subplot, as.numeric(tag_df$tag)), ]",

    "tag_links <- paste(sprintf(\"[ %s-S%s ](#subplot-%s)\",",
    "                           tag_df$tag,",
    "                           tag_df$subplot,",
    "                           tag_df$subplot),",
    "                   collapse = \" | \")",

    "",
    "    cat(\"* \", sp, \": \", tag_links, \"\\n\", sep = \"\")",
    "  }",
    "}",
    "```",
    "",
    nav_links,
    ""
  )

  rmd_content <- yaml_head
  add_body <- function(...) rmd_content <<- c(rmd_content, ...)

  if (any(tf_col) && !any(tf_uncol) && !any(tf_palm)) {
    add_body(mid_section, gencol_section, col_section, index_section)
  } else if (!any(tf_col) && any(tf_uncol) && !any(tf_palm)) {
    add_body(mid_section, gencol_section, uncol_section, index_section)
  } else if (any(tf_col) && any(tf_uncol) && any(tf_palm)) {
    add_body(mid_section, gencol_section, col_section, uncol_section, palm_section, index_section)
  } else if (any(tf_col) && any(tf_uncol) && !any(tf_palm)) {
    add_body(mid_section, gencol_section, col_section, uncol_section, index_section)
  } else if (any(tf_col) && !any(tf_uncol) && any(tf_palm)) {
    add_body(mid_section, gencol_section, col_section, palm_section, index_section)
  } else if (!any(tf_col) && any(tf_uncol) && any(tf_palm)) {
    add_body(mid_section, gencol_section, uncol_section, palm_section, index_section)
  } else {
    add_body(mid_section, gencol_section, index_section)
  }

  add_body("## Individual Subplots {#individual-subplots}", subplot_sections, checklist_section)
  return(rmd_content)
}

# Convert a ForestPlots Query-Library worksheet into compatible data-frame
.fp_query_to_field_sheet_df <- function(path,
                                        sheet = "Data",
                                        locale = readr::locale(decimal_mark = ".", grouping_mark = ",")) {

  ## 1. field template header -------------------------------------------------
  dest_cols <- c(
    "New Tag No","New Stem Grouping","T1","T2","X","Y","Family",
    "Original determination","Morphospecies","D","POM","ExtraD","ExtraPOM",
    "Flag1","Flag2","Flag3","LI","CI","CF","CD1","nrdups","Height",
    "Voucher","Silica","Collected","Census Notes","CAP","Basal Area"
  )

  ## 2. read the Query-Library file ------------------------------------------
  read_in <- function(f, sh) {
    if (grepl("\\.xlsx?$", f, ignore.case = TRUE))
      readxl::read_excel(f, sheet = sh, .name_repair = "minimal")
    else
      readr::read_delim(f, delim = ",", locale = locale,
                        show_col_types = FALSE, .name_repair = "minimal")
  }
  in_dat <- read_in(path, sheet)

  ## 3. prepare an empty field fp data-frame
  out <- as.data.frame(
    matrix(NA_character_, nrow = nrow(in_dat), ncol = length(dest_cols)),
    stringsAsFactors = FALSE
  )
  names(out) <- dest_cols

  ## 4. map source → destination columns -------------------------------------
  src2dest <- c(
    "Tag No"        = "New Tag No",
    "Stem Group ID" = "New Stem Grouping",
    "Sub Plot T1"   = "T1",
    "Sub Plot T2"   = "T2",
    "X"             = "X",
    "Y"             = "Y",
    "Family"        = "Family",
    "Species"       = "Original determination",
    "POM"           = "POM",
    "Extra D"       = "ExtraD",
    "Extra POM"     = "ExtraPOM",
    "F1"            = "Flag1",
    "F2"            = "Flag2",
    "F3"            = "Flag3",
    "LI"            = "LI",
    "CI"            = "CI",
    "CF"            = "CF",
    "CD1"           = "CD1",
    "CD2"           = "nrdups",
    "Height"        = "Height",
    "Voucher Code"  = "Voucher",
    "F5"            = "Silica",
    "Comments"      = "Census Notes"
  )
  for (src in names(src2dest)) {
    dest <- src2dest[[src]]
    if (src %in% names(in_dat))
      out[[dest]] <- in_dat[[src]]
  }

  ## 5. D comes strictly from column D4 --------------------------------------
  if ("D4" %in% names(in_dat))
    out$D <- in_dat$D4

  ## 6. Collected = "yes" when Voucher Code is present -----------------------
  if ("Voucher Code" %in% names(in_dat)) {
    vc <- in_dat[["Voucher Code"]]
    out$Collected <- ifelse(!is.na(vc) & vc != "", "yes", NA_character_)
  }

  ## 7. build metadata line (Plotcode | Plot Name | Date | Team[from PI]) ----
  first_val <- function(df, cols) {
    for (cl in cols)
      if (cl %in% names(df)) return(df[[cl]][1])
    NA_character_
  }

  meta <- rep("", length(dest_cols))
  meta[1]  <- paste0("Plotcode: ", first_val(in_dat, c("Plot Code","PlotCode")))
  meta[2]  <- paste0("Plot Name: ", first_val(in_dat, c("Plot Name","PlotName")))

  date_val <- first_val(in_dat, c("Date","Census Date"))
  if (!is.na(date_val) && inherits(date_val, "Date"))
    date_val <- format(date_val, "%d/%m/%Y")
  meta[10] <- paste0("Date: ", ifelse(is.na(date_val), "", date_val))

  team_val <- first_val(in_dat, "PI")   # the file has PI, not Team
  meta[15] <- paste0("Team: ", ifelse(is.na(team_val), "", team_val))

  ## 8. bind metadata + header + data and return -----------------------------
  field_temp <- tibble::as_tibble(rbind(meta, dest_cols, out))
  names(field_temp) <- NULL   # leave header in row 2
  field_temp
}


# Auxiliar function to compute plot aboveground biomass
.compute_plot_agb <- function(trees = NULL,
                              wd = NULL,
                              md = NULL,
                              allometric_region_id = 5,
                              plot_area = 1) {

  ## 0 ─ locate standard arquives if the path isn't informed -----------
  fallback <- list(
    trees = here::here("data", "Data.csv"),
    wd = here::here("data", "QL_Wood_Density_of_Individual_trees.csv"),
    md = here::here("data", "QL_Plot_Information_for_R_Package_V1.1.csv")
  )
  if (is.null(trees)) trees <- fallback$trees
  if (is.null(wd)) wd <- fallback$wd
  if (is.null(md)) md <- fallback$md

  ## 1 ─ read the data.frame  ------------------------------
  read_fp <- function(x) {
    if (is.data.frame(x)) return(x[, !duplicated(names(x))])
    stopifnot(is.character(x), length(x) == 1, file.exists(x))
    df <- read.csv(x, stringsAsFactors = FALSE)
    df[, !duplicated(names(df))]
  }
  trees <- read_fp(trees)
  wd <- read_fp(wd)
  md <- read_fp(md)

  ## 2 ─ Minimal cleaning -------------------------------------------------------
  md$AllometricRegionID[is.na(md$AllometricRegionID)] <- allometric_region_id
  trees$D4 <- as.numeric(gsub("=", "", trees$D4))
  wd$WD <- as.numeric(gsub("=", "", wd$WD))
  dedup_base <- function(df) {
    base <- sub("\\..*$", "", names(df))
    df[, !duplicated(base), drop = FALSE]
  }

  ## 3 ─ merge → dedup → CalcAGB ---------------------------------------------
  dat <- BiomasaFP::mergefp(trees, md, wd) %>% dedup_base()
  base_nms <- sub("\\..*$", "", names(dat))   # tira sufixos .x .y .1 …
  dup_troncos <- unique(base_nms[duplicated(base_nms)])
  dat[, base_nms %in% dup_troncos] |> names()
  dat <- BiomasaFP::CalcAGB(
    dat,
    dbh = "D4",
    height.data = NULL,  # curva Weibull global
    AGBFun = BiomasaFP::AGBChv14
  )

  ## 4 ─ AGB Sum plot (kg → tons per ha) -------------------------------
  tibble::as_tibble(dat) %>%
    dplyr::group_by(PlotCode) %>%
    dplyr::summarise(
      AGB_t_ha = sum(AGBind, na.rm = TRUE) / 1e3 / plot_area,
      .groups = "drop"
    )
}
