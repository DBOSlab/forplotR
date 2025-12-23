#' Generate forest plot specimen map and collection balance
#'
#' @author Giulia Ottino & Domingos Cardoso
#'
#' @description Processes field data collected using the
#' \href{https://forestplots.net/}{ForestPlots.net} format (field sheets or
#' Query Library output) or 
#' \href{https://www.gov.br/icmbio/pt-br/assuntos/monitoramento/programa-monitora}
#' {Monitora program} layouts, and generates a specimen map with 
#' collection status and spatial distribution of individuals across subplots.
#' It creates a PDF report with plot-level and subplot-level maps, an optional 
#' above-ground biomass summary, and an Excel spreadsheet summarizing the
#' percentage of collected specimens per subplot. The function performs the
#' following steps: (i) validates all input arguments and checks/creates output
#' folders; (ii) reads the input sheet, extracts metadata (team, plot name, plot 
#' code), and cleans the data; (iii) normalizes and cleans coordinate and
#' diameter values, computing global plot coordinates; (iv) builds a PDF report 
#' including plot metadata, the main map, collected and uncollected specimen 
#' maps, an optional palm map and navigable subplot maps; (v) optionally 
#' generates a spreadsheet with the collection percentage per subplot (including 
#' totals), separating palms from non-palms.
#'
#' @usage
#' plot_for_balance(fp_file_path = NULL,
#'                  language = c("en", "pt", "es", "ma"),
#'                  input_type = c("field_sheet", "fp_query_sheet", "monitora"),
#'                  plot_size = 1,
#'                  subplot_size = 10,
#'                  highlight_palms = TRUE,
#'                  station_name = NULL,
#'                  plot_name = NULL,
#'                  plot_code = NULL,
#'                  team = NULL,
#'                  calc_agb = FALSE,
#'                  trees_csv = NULL,
#'                  wd_csv = NULL,
#'                  md_csv = NULL,
#'                  dir = "Results_map_plot",
#'                  filename = "plot_specimen")
#'
#' @param fp_file_path Path to the Excel file (field or query sheet) in 
#' ForestPlots format.
#' 
#' @param language Character. One of en (english), pt (portuguese), es  
#' (spanish), ma (mandarin) (default = en)
#'
#' @param input_type Character. One of `"field_sheet"`, `"fp_query_sheet"` or
#'   `"monitora"`. Specifies the type/layout of the input file.
#'
#' @param plot_size Total plot size in hectares (default = 1).
#'
#' @param subplot_size Side length of each subplot in meters (default = 10).
#'
#' @param highlight_palms Logical. If `TRUE`, highlights Arecaceae individuals
#' in the plots.
#'
#' @param station_name Optional station identifier used when `input_type = "monitora"`.
#' If a vector of length > 1, the function will recursively generate a report
#' and summary spreadsheet for each station.
#'
#' @param plot_name Optional plot name. If provided, overrides the plot name
#' extracted from the input file metadata.
#'
#' @param plot_code Optional plot code. If provided, overrides the plot code
#' extracted from the input file metadata.
#'
#' @param team Optional team or PI name. If provided, overrides the team
#' extracted from the input file metadata.
#'
#' @param calc_agb Logical. If `TRUE`, calculates above-ground biomass (AGB)
#' using \code{BiomasaFP} and adds a summary section to the PDF report.
#'
#' @param trees_csv Optional. Tree-level measurement data used for above-ground
#' biomass (AGB) estimation. Can be either a data frame or a file path
#' (character) to the "Data.csv" file downloaded via ForestPlots "Advanced Search".
#'
#' @param md_csv Optional metadata file downloaded from the ForestPlots
#' "Query Library". Can be a data frame or a file path.
#'
#' @param wd_csv Optional wood density file, with wood density for each 
#' individual tree (by PlotViewID) downloaded from the ForestPlots
#' "Query Library". Can be a data frame or a file path.
#'
#' @param dir Directory path where output will be saved
#' (default is `"Results_map_plot"`). A date-stamped subfolder will be created
#' inside this directory.
#'
#' @param filename Basename for output files (without extension).
#'
#' @return Invisibly returns \code{NULL}. Called for its side effects of
#' generating a PDF report and an Excel file summarizing specimen collection
#' statistics per subplot.
#'
#' @import ggplot2
#' @import BiomasaFP
#' @importFrom readxl read_excel
#' @importFrom dplyr mutate if_else group_by arrange select distinct filter bind_rows any_of
#' @importFrom magrittr "%>%"
#' @importFrom grDevices pdf dev.off colorRampPalette
#' @importFrom utils capture.output download.file
#' @importFrom rmarkdown render pdf_document
#' @importFrom openxlsx read.xlsx createWorkbook addWorksheet writeData createStyle addStyle saveWorkbook
#' @importFrom tinytex latexmk
#' @importFrom readr locale read_delim
#' @importFrom tibble as_tibble tibble
#' @importFrom here here
#' @importFrom stringr str_to_lower str_trim str_detect str_remove word
#' @importFrom tidyr replace_na
#' @importFrom grid unit
#' @importFrom writexl write_xlsx
#'
#' @examples
#' \dontrun{
#' plot_for_balance(fp_file_path = "data/forestplot.xlsx",
#'                  language = c("en", "pt", "es", "ma"),
#'                  input_type = c("field_sheet", "fp_query_sheet", "monitora"),
#'                  plot_size = 1,
#'                  subplot_size = 10,
#'                  highlight_palms = TRUE,
#'                  station_name = NULL,
#'                  plot_name = NULL,
#'                  plot_code = NULL,
#'                  team = NULL,
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
                             language = c("en", "pt", "es", "ma"),
                             input_type = c("field_sheet", "fp_query_sheet", "monitora"),
                             plot_size = 1,
                             subplot_size = 10,
                             highlight_palms = TRUE,
                             station_name = NULL,
                             plot_name = NULL,
                             plot_code = NULL,
                             team = NULL,
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
  
  # language check / normalização
  language <- match.arg(tolower(trimws(as.character(language))),
                        c("en", "pt", "es", "ma"))
  
  # dir check
  dir <- .arg_check_dir(dir)
  
  # Creating the directory to save the file based on the current date
  foldername <- paste0(dir, "/", format(Sys.time(), "%d%b%Y"))
  if (!dir.exists(dir)) dir.create(dir)
  if (!dir.exists(foldername)) dir.create(foldername)
  
  # metadata (fallbacks)
  arg_plot_name <- plot_name
  arg_plot_code <- plot_code
  arg_team      <- team
  
  # aboveground biomass estimation
  agb_tbl <- NULL
  if (calc_agb) {
    agb_tbl <- .compute_plot_agb(
      trees = trees_csv,
      wd = wd_csv,
      md = md_csv,
      plot_area = plot_size
    )
  }
  # normalize input_type 
  input_type <- tolower(trimws(as.character(input_type)))
  input_type <- match.arg(input_type, c("field_sheet","fp_query_sheet","monitora"))
  if (input_type == "monitora") {
    raw <- .monitora_to_field_sheet_df(fp_file_path, station_name = station_name)
  }
  if (input_type == "fp_query_sheet") {
    raw <- .fp_query_to_field_sheet_df(fp_file_path)
  }
  if (input_type == "field_sheet") {
    raw <- suppressMessages(openxlsx::read.xlsx(fp_file_path, sheet = 1, colNames = FALSE))
  }
  if (input_type == "monitora" && !is.null(station_name) && length(station_name) > 1L) {
    for (st in unique(trimws(as.character(station_name)))) {
      plot_for_balance(fp_file_path = fp_file_path,
                       input_type = "monitora",
                       plot_size = plot_size,
                       subplot_size = subplot_size,
                       highlight_palms = highlight_palms,
                       calc_agb = calc_agb,
                       trees_csv = trees_csv,
                       wd_csv = wd_csv,
                       md_csv = md_csv,
                       dir = dir,
                       filename = paste0(filename, "_station_", st),
                       station_name = st,
                       plot_name = arg_plot_name,
                       plot_code = arg_plot_code,
                       team = arg_team)
    }
    return(invisible(TRUE))
  }
  
  metadata_row <- .safe_char_row(raw)
  
  # Detect if metadata row is present (typical patterns: Plotcode, Plot Name, Team, PI)
  has_meta <- any(grepl("^\\s*(Plotcode:|Plot Code:|Plot Name:|Team:|PI:)", 
                        metadata_row, ignore.case = TRUE))
  
  # Safe extractor for metadata fields with default fallback
  .extract_meta <- function(pattern, default = "") {
    if (!has_meta) return(default)
    hit <- metadata_row[grepl(pattern, metadata_row, ignore.case = TRUE)]
    if (length(hit) == 0) return(default)
    sub(paste0("^\\s*", pattern, "\\s*"), "", hit, ignore.case = TRUE)
  }
  
  # Extract metadata (com fallback nos argumentos, se planilha não tiver)
  team <- .extract_meta(
    "Team:",
    default = if (!is.null(arg_team)) arg_team else ""
  )
  
  plot_name <- .extract_meta(
    "Plot Name:",
    default = if (!is.null(arg_plot_name)) arg_plot_name else "Unknown Plot"
  )
  
  plot_code <- .extract_meta(
    "(Plotcode:|Plot Code:)",
    default = if (!is.null(arg_plot_code)) arg_plot_code else ""
  )
  if (!is.null(arg_plot_name) && nzchar(arg_plot_name)) {
    plot_name <- arg_plot_name
  }
  
  if (!is.null(arg_plot_code) && nzchar(arg_plot_code)) {
    plot_code <- arg_plot_code
  }
  
  if (!is.null(arg_team) && nzchar(arg_team)) {
    team <- arg_team
  }
  if (input_type == "fp_query_sheet" && exists("in_dat")) {
    if (is.null(arg_plot_code) && "Plot Code" %in% names(in_dat)) {
      pc <- in_dat[["Plot Code"]]
      if (length(pc) > 0 && any(nzchar(pc))) {
        plot_code <- pc[nzchar(pc)][1]
      } else {
        plot_code <- NA_character_
      }
    }
  }
  
  # Define header row index depending on metadata presence
  header_row_idx <- if (has_meta) 2 else 1
  data_start_idx <- header_row_idx + 1
  
  # Extract header row
  header <- raw[header_row_idx, ] |> unlist() |> as.character()
  header[is.na(header) | header == ""] <- paste0("NA_col_", seq_along(header))[is.na(header) | header == ""]
  
  # Extract data rows
  data <- raw[-seq_len(header_row_idx), , drop = FALSE]
  colnames(data) <- make.unique(header)
  
  # Drop empty columns
  data <- data[, !is.na(colnames(data)) & colnames(data) != "", drop = FALSE]
  
  # Ensure "Collected" column exists
  if (!("Collected" %in% names(data))) data$Collected <- NA_character_
  
  # Normalize empty values to NA
  data$Collected <- gsub("^$", NA, data$Collected)
  data <- .replace_empty_with_na(data)
  
  # Diagnostic message
  if (!has_meta) {
    message("No metadata row found; proceeding without Plot Name / Plot Code / Team.")
  }
  
  # Clean data
  fp_clean <- .clean_fp_data(data)
  
  # Clean data and compute coordinates
  if (input_type == "monitora") {
    fp_coords <- .compute_global_coordinates_monitora(fp_clean)
  } else {
    fp_coords <- .compute_global_coordinates(fp_clean, plot_size, subplot_size)
  }
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
  if (input_type == "field_sheet" | input_type == "fp_query_sheet") {
    base_plot <- .build_fp_base_plot(fp_coords = fp_coords,
                                     subplot_labels = subplot_labels,
                                     plot_size = plot_size,
                                     subplot_size = subplot_size,
                                     plot_name = plot_name,
                                     plot_code = plot_code,
                                     highlight_palms = highlight_palms)
  }
  
  if (input_type == "monitora") {
    base_plot <- .build_monitora_base_plot(fp_coords = fp_coords,
                                           plot_name = plot_name,
                                           plot_code = plot_code,
                                           highlight_palms = highlight_palms)
  }
  
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
  
  # usar o mesmo foldername e filename já utilizados para o PDF final
  png_dir <- foldername
  base_name <- filename
  
  # 1) General
  ggplot2::ggsave(
    filename = file.path(png_dir, paste0(base_name, "_general.png")),
    plot     = base_plot,
    width = 14, height = 11, units = "in", dpi = 300
  )
  
  # 2) Collected
  if (!is.null(collected_plot)) {
    ggplot2::ggsave(
      filename = file.path(png_dir, paste0(base_name, "_collected.png")),
      plot     = collected_plot,
      width = 14, height = 11, units = "in", dpi = 300
    )
  }
  
  # 3) Uncollected
  if (!is.null(uncollected_plot)) {
    ggplot2::ggsave(
      filename = file.path(png_dir, paste0(base_name, "_uncollected.png")),
      plot     = uncollected_plot,
      width = 14, height = 11, units = "in", dpi = 300
    )
  }
  
  # 4) Uncollected palms
  if (!is.null(uncollected_palm_plot)) {
    ggplot2::ggsave(
      filename = file.path(png_dir, paste0(base_name, "_uncollected_palms.png")),
      plot     = uncollected_palm_plot,
      width = 14, height = 11, units = "in", dpi = 300
    )
  }
  
  # Generate subplot plots for each unique subplot (T1)
  subplot_plots <- list()
  
  if (input_type == "monitora") {
    # >>> MONITORA: create 40 subplots (N,S,L,O × 1..10)
    arms_order <- c("N","S","L","O")
    arm_id <- setNames(1:4, arms_order)
    
    full_idx <- do.call(rbind, lapply(arms_order, function(a) {
      data.frame(subunit_letter = a, T2 = 1:10, stringsAsFactors = FALSE)
    }))
    full_idx$T1 <- (arm_id[ full_idx$subunit_letter ] - 1L) * 10L + full_idx$T2
    
    
    .safe_breaks <- function(x) {
      x <- x[is.finite(x)]
      if (!length(x)) return(c(1,3,7,10))  # fallback
      rng <- range(x); br <- round(seq(rng[1], rng[2], length.out = 4), 0)
      br[1] <- rng[1]; br[4] <- rng[2]; br
    }
    
    for (k in seq_len(nrow(full_idx))) {
      sub  <- full_idx$subunit_letter[k]
      t2   <- full_idx$T2[k]
      
      sp_data <- fp_coords %>%
        dplyr::filter(subunit_letter == sub, T2 == t2) %>%
        dplyr::mutate(
          Status   = dplyr::case_when(
            highlight_palms & Family == "Arecaceae" ~ "Palms",
            !is.na(Collected) & Collected != ""     ~ "Collected",
            TRUE                                     ~ "Uncollected"
          ),
          diameter = D/10
        )
      
      # be sure of local cols (se não vieram batizadas)
      if (!"X_loc" %in% names(sp_data)) sp_data$X_loc <- sp_data$X
      if (!"Y_loc" %in% names(sp_data)) sp_data$Y_loc <- sp_data$Y
      
      # convert coordinates to  0..10 × 0..10 inside subplot zig-zag
      sp_plot <- .monitora_to_cell_coords(sp_data)  # devolve colunas x10,y10
      brk <- .safe_breaks(sp_plot$diameter)
      brk <- unique(na.omit(brk))
      if (length(brk) < 2) brk <- c(brk, brk + 0.01)
      if (anyDuplicated(brk)) brk <- unique(round(brk, 2))
      brk <- sort(brk)
      p <- ggplot2::ggplot(sp_plot, ggplot2::aes(x = x10, y = y10)) +
        ggplot2::geom_rect(ggplot2::aes(xmin = 0, xmax = 10, ymin = 0, ymax = 10),
                           fill = NA, color = "black", linewidth = 0.6) +
        ggplot2::geom_point(ggplot2::aes(size = diameter, fill = Status),
                            shape = 21, color = "black", stroke = 0.2, alpha = 0.9) +
        ggplot2::geom_text(ggplot2::aes(label = `New Tag No`), size = 1.7) +
        ggplot2::scale_fill_manual(values = c(Collected="gray", Uncollected="red", Palms="gold"),
                                   name = "Status") +
        ggplot2::scale_size_area(name = "DBH (cm)",breaks = brk, 
                                 labels = sprintf("%.1f", brk),
                                 limits = range(brk, na.rm = TRUE),
                                 max_size = 10,
                                 guide = "legend"
        )+
        ggplot2::labs(
          title = paste("Subunit", sub, "— Subplot", t2),
          subtitle = paste("Plot Name:", plot_name, "| Plot Code:", plot_code),
          x = "Local X (m)", y = "Local Y (m)"
        )+
        ggplot2::coord_fixed() +
        ggplot2::theme_bw() +
        ggplot2::theme(panel.grid.minor = ggplot2::element_blank(),
                       panel.grid.major = ggplot2::element_line(size = 0.3, color = "gray80"),
                       legend.position = "right") +
        ggplot2::guides(fill = ggplot2::guide_legend(override.aes = list(size = 5), order = 1),
                        size = ggplot2::guide_legend(title = "DBH (cm)", order = 2))
      
      subplot_plots[[length(subplot_plots) + 1]] <- list(plot = p, data = sp_plot)
    }
    
  } else {
    # field_sheet or fp_query_sheet individual subplot mapping
    for (i in sort(unique(fp_coords$T1))) {
      sp_data <- fp_coords %>% dplyr::filter(T1 == i) %>%
        dplyr::mutate(
          Status = dplyr::case_when(
            highlight_palms & Family == "Arecaceae" ~ "Palms",
            !is.na(Collected) & Collected != ""     ~ "Collected",
            TRUE                                     ~ "Uncollected"
          ),
          diameter = D/10
        )
      
      valid_d <- sp_data$diameter[is.finite(sp_data$diameter)]
      if (length(valid_d)) {
        min_d <- min(valid_d); max_d <- max(valid_d)
        breaks_d <- round(seq(min_d, max_d, length.out = 4), 0)
        breaks_d[1] <- min_d; breaks_d[4] <- max_d
      } else {
        breaks_d <- c(1,3,7,10)
      }
      breaks_d <- unique(na.omit(round(breaks_d, 1)))
      if (length(breaks_d) < 2) breaks_d <- c(breaks_d, breaks_d + 0.01)
      p <- ggplot2::ggplot(sp_data, ggplot2::aes(x = X, y = Y)) +
        ggplot2::geom_rect(ggplot2::aes(xmin = 0, xmax = subplot_size, ymin = 0, ymax = subplot_size),
                           fill = NA, color = "black", linewidth = 0.6) +
        ggplot2::geom_point(ggplot2::aes(size = diameter, fill = Status),
                            shape = 21, color = "black", stroke = 0.2, alpha = 0.9) +
        ggplot2::geom_text(ggplot2::aes(label = `New Tag No`), size = 1.7) +
        ggplot2::scale_fill_manual(values = c(Collected="gray", Uncollected="red", Palms="gold"),
                                   name = "Status") +
        ggplot2::scale_size_area(
          name = "DBH (cm)",
          breaks = breaks_d,
          labels = sprintf("%.1f", breaks_d),
          limits = range(breaks_d, na.rm = TRUE),
          max_size = 10,
          guide = "legend")+
        ggplot2::labs(title = paste("Plot Name:", plot_name),
                      subtitle = paste0("Plot Code: ", plot_code),
                      x = "X (m)", y = "Y (m)") +
        ggplot2::coord_fixed() +
        ggplot2::theme_bw() +
        ggplot2::theme(panel.grid.minor = ggplot2::element_blank(),
                       panel.grid.major = ggplot2::element_line(size = 0.3, color = "gray80"),
                       legend.position = "right") +
        ggplot2::guides(fill = ggplot2::guide_legend(override.aes = list(size = 5), order = 1),
                        size = ggplot2::guide_legend(title = "DBH (cm)", order = 2))
      
      subplot_plots[[length(subplot_plots) + 1]] <- list(plot = p, data = sp_data)
    }
  }
  
  # Calculate plot statistics
  total_specimens <- nrow(fp_coords)
  collected_count <- sum(!is.na(fp_coords$Collected) & fp_coords$Family != "Arecaceae")
  uncollected_count <- sum(is.na(fp_coords$Collected) & fp_coords$Family != "Arecaceae")
  palms_count <- sum(fp_coords$Family == "Arecaceae")
  
  # Map tag → subplot (condicional para Monitora)
  if (tolower(trimws(as.character(input_type))) == "monitora") {
    tag_to_subplot <- fp_coords %>%
      dplyr::transmute(
        `New Tag No`   = trimws(as.character(`New Tag No`)),
        T1             = suppressWarnings(as.integer(T1)),  # 1=N, 2=S, 3=L, 4=O
        T2             = suppressWarnings(as.integer(T2)),  # 1..10
        subunit_letter = dplyr::case_when(
          T1 == 1 ~ "N",
          T1 == 2 ~ "S",
          T1 == 3 ~ "L",
          T1 == 4 ~ "O",
          TRUE    ~ NA_character_
        ),
        subplot_index  = dplyr::if_else(is.finite(T1) & is.finite(T2),
                                        (T1 - 1L) * 10L + T2, NA_integer_)
      ) %>%
      dplyr::distinct()
  } else {
    # comportamento original
    tag_to_subplot <- fp_coords %>%
      dplyr::select(`New Tag No`, T1) %>%
      dplyr::distinct()
  }
  
  
  #### Prepare RMarkdown sections for each subplot with navigation and page breaks ####
    rmd_content <- .create_rmd_content(
      subplot_plots,
      tf_col, tf_uncol, tf_palm,
      plot_name, plot_code, spec_df,
      has_agb = !is.null(agb_tbl)
      language = language
      
  # Write Rmd content to file
  writeLines(c(rmd_content, ""), rmd_path)
  
  options(tinytex.pdflatex.args = "--no-crop")
  
  param_list <- list(
    input_type = input_type,
    metadata = list(
      plot_name = plot_name,
      plot_code = plot_code,
      team      = team,
      census_years = if (exists("monitora_census_years", inherits = TRUE)) get("monitora_census_years") else NULL,
      census_n     = if (exists("monitora_census_n",     inherits = TRUE)) get("monitora_census_n") else NULL
    ),
    main_plot = base_plot,
    subplots_list = subplot_plots,
    subplot_size  = subplot_size,
    stats = list(
      total       = total_specimens,
      collected   = collected_count,
      uncollected = uncollected_count,
      palms       = palms_count,
      spec_df     = spec_df,
      dead_since_first     = if (exists("monitora_dead_since_first", inherits = TRUE)) get("monitora_dead_since_first") else NULL,
      recruits_since_first = if (exists("monitora_recruits_since_first", inherits = TRUE)) get("monitora_recruits_since_first") else NULL
    ),
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
  
  # Engine LaTeX: XeLaTeX just for mandarin
  latex_engine <- if (language == "ma") "xelatex" else "pdflatex"
  
  # Render the Rmd file to PDF using rmarkdown::render with params
  .render_plot_report(
    rmd_path = rmd_path,
    output_pdf_path = foldername,
    output_pdf_name = final_pdf,
    params = param_list,
    latex_engine = latex_engine
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
    params,
    latex_engine = "pdflatex") {   # default seguro
  
  render_env <- new.env(parent = globalenv())
  wd <- getwd()
  old_logs <- list.files(
    path = wd,
    pattern = "^file[0-9a-f]+\\.log$",
    full.names = TRUE
  )
  # Temp output name
  temp_pdf <- tempfile(fileext = ".pdf")
  extra_deps <- NULL
  if (identical(latex_engine, "xelatex")) {
    extra_deps <- rmarkdown::latex_dependency("ctex", options = "fontset=fandol")
  }
  
  rmarkdown::render(
    input       = rmd_path,
    output_file = basename(temp_pdf),
    output_dir  = tempdir(),
    params      = params,
    envir       = render_env,
    clean       = FALSE,
    output_format = rmarkdown::pdf_document(
      keep_tex          = TRUE,
      latex_engine      = latex_engine,
      fig_caption       = TRUE,
      extra_dependencies = extra_deps  
    )
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
    tex_lines  <- readLines(tex_file)
    insert_idx <- grep("\\\\begin\\{document\\}", tex_lines)[1]
    new_tex_lines <- append(tex_lines, preamble_lines, after = insert_idx - 1)
    writeLines(new_tex_lines, tex_file)
    tinytex::latexmk(tex_file, engine = latex_engine)
  }
  
  # Move output PDF to final location
  if (!dir.exists(output_pdf_path)) dir.create(output_pdf_path, recursive = TRUE)
  final_pdf_path <- file.path(output_pdf_path, basename(output_pdf_name))
  file.copy(sub("\\.tex$", ".pdf", tex_file), final_pdf_path, overwrite = TRUE)
  
  if (file.exists(final_pdf_path)) {
    message("Full PDF report saved to: ", final_pdf_path)
  } else {
    warning("PDF generation failed.")
  }
  new_logs <- list.files(
    path = wd,
    pattern = "^file[0-9a-f]+\\.log$",
    full.names = TRUE
  )
  extra_logs <- setdiff(new_logs, old_logs)
  if (length(extra_logs)) {
    unlink(extra_logs, force = TRUE)
  }
  log_files <- list.files(
    path        = output_pdf_path,
    pattern     = "^file[0-9a-f]+\\.log$",
    full.names  = TRUE
  )
  if (length(log_files)) {
    unlink(log_files, force = TRUE)
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

# Build fp base plot 
.build_fp_base_plot <- function(fp_coords,
                                subplot_labels,
                                plot_size,
                                subplot_size,
                                plot_name,
                                plot_code,
                                highlight_palms) {
  max_x <- 100
  max_y <- (plot_size / 1) * 100
  
  base_plot <- ggplot2::ggplot(fp_coords, ggplot2::aes(x = global_x, y = global_y)) +
    ggplot2::geom_vline(
      xintercept = seq(0, max_x, subplot_size),
      color = "gray50",
      linewidth = 0.2
    ) +
    ggplot2::geom_hline(
      yintercept = seq(0, max_y, subplot_size),
      color = "gray50",
      linewidth = 0.2
    ) +
    ggplot2::geom_rect(
      ggplot2::aes(xmin = 0, xmax = max_x, ymin = 0, ymax = max_y),
      fill = NA,
      color = "black",
      linewidth = 0.6
    ) +
    ggplot2::geom_text(
      data = subplot_labels,
      ggplot2::aes(x = center_x, y = center_y, label = T1),
      color = "gray",
      size = 2,
      fontface = "bold",
      inherit.aes = FALSE
    ) +
    ggplot2::geom_point(
      ggplot2::aes(
        size = diameter,
        fill = factor(
          dplyr::case_when(
            highlight_palms & Family == "Arecaceae" ~ "Palms",
            !is.na(Collected) & Collected != ""     ~ "Collected",
            TRUE                                    ~ "Uncollected"
          ),
          levels = c("Collected", "Uncollected", "Palms")
        )
      ),
      shape = 21,
      stroke = 0.2,
      color = "black",
      alpha = 0.9
    ) +
    ggplot2::geom_text(
      ggplot2::aes(label = `New Tag No`),
      vjust = 0.5,
      hjust = 0.5,
      size = 0.6
    ) +
    ggplot2::scale_x_continuous(
      limits = c(0, max_x),
      breaks = seq(0, 100, by = subplot_size)
    ) +
    ggplot2::scale_y_continuous(
      limits = c(0, max_y),
      breaks = seq(0, max_y, by = subplot_size)
    ) +
    ggplot2::coord_fixed(ratio = 1, clip = "off") +
    ggplot2::scale_fill_manual(
      values = c(
        "Collected"   = "gray",
        "Uncollected" = "red",
        "Palms"       = "gold"
      ),
      name = "Status"
    ) +
    ggplot2::scale_size_continuous(range = c(2, 6), guide = "none") +
    ggplot2::labs(
      x        = "X (m)",
      y        = "Y (m)",
      title    = paste0("Collection Balance ", plot_name),
      subtitle = paste0("Plot Code: ", plot_code)
    ) +
    ggplot2::theme_bw() +
    ggplot2::theme(
      legend.position      = "right", 
      legend.justification = c(0, 1),
      legend.box.margin    = ggplot2::margin(0, 5, 0, 5),
      plot.margin          = ggplot2::margin(5.5, 20, 5.5, 5.5) 
    ) +
    ggplot2::guides(
      fill = ggplot2::guide_legend(override.aes = list(size = 5))
    )
  
  base_plot
}

# Build monitora base plot (com título/subtítulo via annotate)
.build_monitora_base_plot <- function(
    fp_coords,
    plot_name,
    plot_code,
    highlight_palms,
    arm_offset = 18,
    pad        = 4
) {
  stopifnot(all(c("draw_x","draw_y","subunit_letter","D","Collected","Family") %in% names(fp_coords)))
  
  # ---- constants ----
  arm_length <- 50
  cell_size  <- 10
  gap_title  <- 3.0
  gap_sub    <- 2.5
  left_inset <- 0.8
  
  # grid per arm 
  mk_cells <- function(arm) {
    base <- expand.grid(c = 0:4, r = 0:1)
    if (arm == "N") {
      dplyr::mutate(base, arm="N",
                    xmin=-cell_size+r*cell_size, xmax=-cell_size+r*cell_size+cell_size,
                    ymin= arm_offset+c*cell_size, ymax= arm_offset+c*cell_size+cell_size)
    } else if (arm == "S") {
      dplyr::mutate(base, arm="S",
                    xmin=-cell_size+r*cell_size, xmax=-cell_size+r*cell_size+cell_size,
                    ymin=-(arm_offset+arm_length)+c*cell_size,
                    ymax=-(arm_offset+arm_length)+c*cell_size+cell_size)
    } else if (arm == "L") {
      dplyr::mutate(base, arm="L",
                    xmin= arm_offset+c*cell_size, xmax= arm_offset+c*cell_size+cell_size,
                    ymin=-cell_size+r*cell_size, ymax=-cell_size+r*cell_size+cell_size)
    } else { # O
      dplyr::mutate(base, arm="O",
                    xmin=-(arm_offset+arm_length)+c*cell_size,
                    xmax=-(arm_offset+arm_length)+c*cell_size+cell_size,
                    ymin=-cell_size+r*cell_size, ymax=-cell_size+r*cell_size+cell_size)
    }
  }
  cells <- dplyr::bind_rows(mk_cells("N"), mk_cells("S"), mk_cells("L"), mk_cells("O"))
  
  centers <- cells %>%
    dplyr::mutate(
      cx=(xmin+xmax)/2, cy=(ymin+ymax)/2,
      col_from_center = dplyr::case_when(arm %in% c("N","L") ~ c, arm %in% c("S","O") ~ 4L - c, TRUE ~ c),
      col_idx  = col_from_center + 1L,
      base_num = (col_idx - 1L) * 2L + 1L,
      label    = ifelse(col_idx %% 2L == 1L, base_num + r, base_num + (1L - r))
    ) %>% dplyr::select(-col_from_center, -col_idx, -base_num)
  
  # cross & limits 
  to_arm     <- arm_offset
  solid_half <- 10  # 10 m each side => total 20 m
  gx_min <- -(arm_offset + arm_length) - pad
  gx_max <-  (arm_offset + arm_length) + pad
  gy_min <- gx_min
  gy_max <- gx_max
  y_max_ext <- gy_max + gap_title + gap_sub + 0.5
  
  edge_pad_in <- max(0.6, pad/2)
  arm_labels <- data.frame(
    lab   = c("N","S","L","O"),
    lx    = c(0, 0, gx_max - edge_pad_in, gx_min + edge_pad_in),
    ly    = c(gy_max - edge_pad_in, gy_min + edge_pad_in, 0, 0),
    hjust = c(0.5, 0.5, 0.0, 1.0),
    vjust = c(0.0, 1.0, 0.5, 0.5)
  )
  
  # --- mini legend (define BEFORE building the plot) --------------------
  dashed_len <- max(to_arm - solid_half, 0)
  sx <- gx_max - 130   # more left  if you decrease
  sy <- gy_min + 10   # lower      if you decrease
  
  mini_legend_layers <- list(
    ggplot2::geom_segment(
      x = sx, y = sy, xend = sx + dashed_len, yend = sy,
      inherit.aes = FALSE, color = "gray50", linewidth = 0.6, linetype = "22",
      show.legend = FALSE
    ),
    ggplot2::annotate(
      "text", x = sx + dashed_len/2, y = sy - 1.6, label = "40 m",
      size = 3.4, hjust = 0.5, vjust = 1, color = "gray30"
    ),
    ggplot2::geom_segment(
      x = sx, y = sy - 5.0, xend = sx + 20, yend = sy - 5.0,
      inherit.aes = FALSE, color = "gray35", linewidth = 0.7,
      show.legend = FALSE
    ),
    ggplot2::annotate(
      "text", x = sx + 10, y = sy - 6.6, label = "20 m",
      size = 3.4, hjust = 0.5, vjust = 1, color = "gray30"
    )
  )
  
  # build plot -----------------------------------------------------------
  ggplot2::ggplot(
    fp_coords,
    ggplot2::aes(
      x = dplyr::case_when(
        subunit_letter %in% c("N","S") ~ draw_x,
        subunit_letter == "L"          ~ draw_x - 50 + arm_offset,
        subunit_letter == "O"          ~ draw_x + 100 - (arm_offset + arm_length),
        TRUE ~ draw_x
      ),
      y = dplyr::case_when(
        subunit_letter %in% c("L","O") ~ draw_y,
        subunit_letter == "N"          ~ draw_y - 50 + arm_offset,
        subunit_letter == "S"          ~ draw_y + 100 - (arm_offset + arm_length),
        TRUE ~ draw_y
      )
    )
  ) +
    ggplot2::geom_rect(
      data = cells,
      ggplot2::aes(xmin = xmin, xmax = xmax, ymin = ymin, ymax = ymax),
      inherit.aes = FALSE, fill = NA, color = "gray70", linewidth = 0.28, show.legend = FALSE
    ) +
    ggplot2::geom_text(
      data = centers,
      ggplot2::aes(x = cx, y = cy, label = label),
      inherit.aes = FALSE, size = 3.2, fontface = "bold", color = "gray35", show.legend = FALSE
    ) +
    ggplot2::geom_segment(x = -solid_half, xend =  solid_half, y = 0, yend = 0,
                          inherit.aes = FALSE, color = "gray35", linewidth = 0.7, show.legend = FALSE) +
    ggplot2::geom_segment(y = -solid_half, yend =  solid_half, x = 0, xend = 0,
                          inherit.aes = FALSE, color = "gray35", linewidth = 0.7, show.legend = FALSE) +
    ggplot2::geom_segment(x = -to_arm, xend = -solid_half, y = 0, yend = 0,
                          inherit.aes = FALSE, color = "gray50", linewidth = 0.6, linetype = "22", show.legend = FALSE) +
    ggplot2::geom_segment(x =  solid_half, xend =  to_arm, y = 0, yend = 0,
                          inherit.aes = FALSE, color = "gray50", linewidth = 0.6, linetype = "22", show.legend = FALSE) +
    ggplot2::geom_segment(y = -to_arm, yend = -solid_half, x = 0, xend = 0,
                          inherit.aes = FALSE, color = "gray50", linewidth = 0.6, linetype = "22", show.legend = FALSE) +
    ggplot2::geom_segment(y =  solid_half, yend =  to_arm, x = 0, xend = 0,
                          inherit.aes = FALSE, color = "gray50", linewidth = 0.6, linetype = "22", show.legend = FALSE) +
    ggplot2::geom_point(
      ggplot2::aes(
        size = (D/10),
        fill = factor(
          dplyr::case_when(
            highlight_palms & Family == "Arecaceae" ~ "Palms",
            !is.na(Collected) & Collected != ""     ~ "Collected",
            TRUE                                     ~ "Uncollected"
          ),
          levels = c("Collected","Uncollected","Palms")
        )
      ),
      shape = 21, stroke = 0.2, color = "black", alpha = 0.9
    ) +
    ggplot2::geom_text(ggplot2::aes(label = `New Tag No`), size = 0.65, show.legend = FALSE) +
    ggplot2::annotate("text",
                      x = gx_min + left_inset, y = gy_max + gap_title + gap_sub,
                      label = paste0("Collection Balance — ", plot_name),
                      size = 5, hjust = 0, vjust = 1
    ) +
    ggplot2::annotate("text",
                      x = gx_min + left_inset, y = gy_max + gap_title,
                      label = paste0("Plot Code: ", ifelse(is.null(plot_code),"",plot_code)),
                      size = 3.8, hjust = 0, vjust = 1
    ) +
    ggplot2::scale_x_continuous(limits = c(gx_min, gx_max),
                                expand = ggplot2::expansion(add = c(6, 0))) +
    ggplot2::scale_y_continuous(limits = c(gy_min, y_max_ext),
                                expand = ggplot2::expansion(add = c(8, 0))) +  # extra bottom room for "10 m"
    ggplot2::coord_fixed(clip = "off") +
    ggplot2::theme_void() +
    ggplot2::scale_fill_manual(values = c(Collected="gray", Uncollected="red", Palms="gold"),
                               name = "Status") +
    ggplot2::scale_size_continuous(range = c(2,6), guide = "none") +
    ggplot2::geom_text(
      data = arm_labels,
      ggplot2::aes(x = lx, y = ly, label = lab, hjust = hjust, vjust = vjust),
      inherit.aes = FALSE, size = 9, fontface = "bold", show.legend = FALSE
    ) +
    ggplot2::theme(
      legend.position      = c(0.98, 0.10),
      legend.justification = c(1, 0),
      legend.background    = ggplot2::element_rect(fill = "white", colour = NA),
      legend.title         = ggplot2::element_text(size = 13, face = "plain"),
      legend.text          = ggplot2::element_text(size = 12),
      legend.key.size      = grid::unit(1.1, "lines")
    ) +
    ggplot2::guides(fill = ggplot2::guide_legend(override.aes = list(size = 6), order = 1)) +
    # <- add the pre-built list of layers here
    mini_legend_layers
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

# Extract a character row from a data.frame / matrix / vector
.safe_char_row <- function(x, row = 1L) {

    if (is.null(x) || NROW(x) < row) return(character())
  
  if (is.data.frame(x) || is.matrix(x)) {
    v <- x[row, , drop = TRUE]
  } else if (is.atomic(x)) {
    v <- x
  } else {
    v <- x[[row]]
  }
  
  v <- as.character(v)
  v[is.na(v)] <- ""
  trimws(v)
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
  
  #### merge → dedup → CalcAGB ---------------------------------------------
  dat <- BiomasaFP::mergefp(trees, md, wd) %>% dedup_base()
  base_nms <- sub("\\..*$", "", names(dat))   
  dup_troncos <- unique(base_nms[duplicated(base_nms)])
  dat[, base_nms %in% dup_troncos] |> names()
  dat <- BiomasaFP::CalcAGB(
    dat,
    dbh = "D4",
    height.data = NULL,  
    AGBFun = BiomasaFP::AGBChv14
  )
  
  #### AGB Sum plot  -------------------------------
  tibble::as_tibble(dat) %>%
    dplyr::group_by(PlotCode) %>%
    dplyr::summarise(
      AGB_t_ha = sum(AGBind, na.rm = TRUE) / 1e3 / plot_area,
      .groups = "drop"
    )
}

 