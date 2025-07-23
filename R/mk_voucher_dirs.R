#' Organize voucher image directories by updated identification
#'
#' @author Giulia Ottino & Domingos Cardoso
#'
#' @description Organizes image folders of voucher specimens based on their
#' updated taxonomic identification provided in \href{https://forestplots.net/}{Forestplots format}
#' template. It is designed to be safely rerun multiple times,
#' adapting its behavior depending on the existing directory structure.On the
#' **first execution**, the function will read the voucher information from
#' the Forestplots standardized file, (i) create a main folder (default:
#' `"voucher_imgs"`), if it does not already exist.(ii) For each voucher marked
#' as "Collected", it will create a directory structure following the format
#' `Family/Genus/Voucher` to receive the corresponding images. On **subsequent
#' executions**, the function performs a reorganization: identifies and compares
#' existing folders that may have outdated or incorrect genus assignments. If it
#' finds that a voucher is in the wrong genus folder (e.g., due to taxonomic
#' reidentification),it moves the images to the correct folder. In case the
#' correct folder already contains images, the function avoids overwriting and
#' keeps the content intact.Once images are moved (or verified), the function
#' deletes any empty,outdated voucher directories.This ensures that the
#' directory structure always reflects the most up-to-date taxonomic information,
#' making image organization robust, repeatable, and easy to maintain.
#'
#' @param fp_file_path Character. Path to the Excel file containing voucher data.
#' @param output_dir Character. Directory to organize the voucher image folders.
#' Default is "voucher_imgs".
#'
#' @examples
#' \dontrun{
#' mk_voucher_dirs(fp_file_path = "data/RUS_plot.xlsx",
#'                          output_dir = "voucher_imgs")
#'                          }
#' @importFrom readxl read_excel
#' @export

mk_voucher_dirs <- function(fp_file_path = NULL,
                            output_dir = "voucher_imgs") {
  if (is.null(fp_file_path) || !file.exists(fp_file_path)) {
    stop("Please provide a valid Excel file path.")
  }

  # Read and clean the data
  fp_sheet <- suppressMessages(readxl::read_excel(fp_file_path, sheet = 1))
  header <- as.character(unlist(fp_sheet[1, ]))
  header[is.na(header)] <- paste0("NA_col_", seq_along(header))[is.na(header)]
  colnames(fp_sheet) <- make.unique(header)
  fp_sheet <- fp_sheet[-1, ]
  fp_sheet <- fp_sheet[, !is.na(colnames(fp_sheet)) & colnames(fp_sheet) != ""]

  # Normalize column types
  fp_sheet$Family <- as.character(fp_sheet$Family)
  fp_sheet$Voucher <- as.character(fp_sheet$Voucher)
  fp_sheet$Collected <- as.character(fp_sheet$Collected)

  # Create main directory if it doesn't exist
  if (!dir.exists(output_dir)) {
    dir.create(output_dir, recursive = TRUE)
    message(paste("Created main output directory:", output_dir))
  }

  # List all existing voucher directories
  existing_dirs <- list.dirs(output_dir, recursive = TRUE, full.names = TRUE)
  existing_voucher_dirs <- existing_dirs[grepl("[A-Z]{3}[0-9]{4}",
                                               basename(existing_dirs))]

  # Store paths of old voucher directories for later deletion check
  old_voucher_dirs_checked <- character(0)

  for (i in seq_len(nrow(fp_sheet))) {
    voucher <- fp_sheet$Voucher[i]
    collected <- fp_sheet$Collected[i]
    family <- fp_sheet$Family[i]
    genus_correct <- sub(" .*", "", fp_sheet$`Original determination`[i])

    if (is.na(voucher) || voucher == "" || is.na(collected) || collected == "") next

    correct_path <- file.path(output_dir, family, genus_correct, voucher)
    correct_path_norm <- normalizePath(correct_path, winslash = "/",
                                       mustWork = FALSE)

    # Find directories that contain this voucher
    voucher_matches <- existing_voucher_dirs[basename(existing_voucher_dirs) == voucher]

    for (vm in voucher_matches) {
      vm_norm <- normalizePath(vm, winslash = "/", mustWork = FALSE)

      if (!identical(vm_norm, correct_path_norm)) {
        photos <- list.files(vm, full.names = TRUE)

        if (length(photos) > 0) {
          if (!dir.exists(correct_path)) {
            dir.create(correct_path, recursive = TRUE)
            message(paste("Created correct directory:", correct_path))
          }

          if (length(list.files(correct_path)) == 0) {
            file.copy(photos, correct_path, overwrite = FALSE)
            message(paste("Moved photos from", vm, "to", correct_path))
          } else {
            message(paste("Correct folder already has content; skipped moving from",
                          vm))
          }
        }

        # Mark old voucher directory for possible deletion
        old_voucher_dirs_checked <- c(old_voucher_dirs_checked, vm)
      }
    }

    # Create the correct folder if it doesn't exist
    if (!dir.exists(correct_path)) {
      dir.create(correct_path, recursive = TRUE)
      message(paste("Created new voucher folder:", correct_path))
    }
  }

  # Check old voucher directories for possible deletion
  for (old_voucher_dir in unique(old_voucher_dirs_checked)) {
    # Check if the voucher folder is empty
    if (dir.exists(old_voucher_dir) && length(list.files(old_voucher_dir)) == 0) {
      unlink(old_voucher_dir, recursive = TRUE)
      message(paste("Deleted empty voucher folder:", old_voucher_dir))
    }
  }
}
