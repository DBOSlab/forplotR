# GigaFlor field sheets ####
#
# The GigaFlor network (Amazon giant forests) uses a field sheet close to the
# ForestPlots one, with three differences handled here:
#   * column names: "Tag Number" (= New Tag No), "Species"
#     (= Original determination), "Main Stem Tag" (stem grouping);
#   * D is given in centimetres (the package works in millimetres);
#   * X and Y are the geographic coordinates of each tree in decimal degrees
#     (X = longitude, Y = latitude), not local coordinates within the subplot.
#
# Protocol 1: a 1-ha plot of 25 subplots of 20 x 20 m (T1, numbered in the
# same serpentine order as ForestPlots), each split into four 10 x 10 m
# quadrants (T2).


#' Default subplot side (m) of each GigaFlor protocol
#' @keywords internal
#' @noRd
.gigaflor_protocols <- function() {
  list(`1` = list(plot_size = 1, subplot_size = 20))
}


#' Worksheet that holds the GigaFlor data ("Data" when present, else the first)
#' @keywords internal
#' @noRd
.gigaflor_sheet <- function(path) {
  sheets <- tryCatch(readxl::excel_sheets(path), error = function(e) character())
  hit <- sheets[tolower(trimws(sheets)) == "data"]
  if (length(hit)) hit[1L] else 1L
}


#' Harmonize the GigaFlor-specific columns of a field sheet
#'
#' Run on the harmonized field sheet before multistem consolidation:
#' drops empty template rows, converts D from cm to mm (D is taken from
#' "CAP (cm)" when empty), keeps the decimal-degree coordinates as
#' `Longitude` / `Latitude` and builds `New Stem Grouping` from
#' "Main Stem Tag".
#'
#' @param fp_sheet Harmonized field-sheet data frame.
#' @param verbose Logical.
#' @return The data frame, with `X`/`Y` still in decimal degrees.
#' @keywords internal
#' @noRd
.gigaflor_prepare <- function(fp_sheet, verbose = TRUE) {
  df <- as.data.frame(fp_sheet, stringsAsFactors = FALSE, check.names = FALSE)

  # Bilingual headers ("Main Stem Tag | Placa do fuste principal", as in the
  # Portuguese GigaFlor sheet): GigaFlor columns go by their English side
  short <- trimws(sub("\\s*\\|.*$", "", names(df)))
  rename <- grepl("|", names(df), fixed = TRUE) & nzchar(short) &
    !(short %in% names(df)) & !duplicated(short)
  names(df)[rename] <- short[rename]

  # The template is pre-formatted far below the data: drop rows with no tag,
  # family, species or measurement
  key <- intersect(c("New Tag No", "Family", "Original determination", "D", "CAP", "X", "Y"),
                   names(df))
  filled <- Reduce(`|`, lapply(df[key], function(v) {
    v <- trimws(as.character(v))
    !is.na(v) & nzchar(v)
  }), FALSE)
  df <- df[filled, , drop = FALSE]

  # D (cm) -> mm; fall back on the circumference
  d_cm <- .parse_num(df$D)
  if ("CAP" %in% names(df)) {
    cap <- .parse_num(df$CAP)
    miss <- !is.finite(d_cm) & is.finite(cap)
    d_cm[miss] <- round(cap[miss] / pi, 2)
  }
  df$D <- d_cm * 10

  # Collection columns are optional in the GigaFlor template
  for (cl in c("Voucher", "Collected")) {
    if (!(cl %in% names(df))) df[[cl]] <- NA_character_
  }
  # The template's Collected column is a Yes/No list: the package counts any
  # non-empty value as collected, so a negative answer becomes empty
  coll <- tolower(trimws(as.character(df$Collected)))
  negative <- coll %in% c("no", "n", "n\u00e3o", "nao", "0", "false", "-")
  df$Collected[negative] <- NA_character_

  # Geographic coordinates (decimal degrees)
  df$Longitude <- .parse_num(df$X)
  df$Latitude <- .parse_num(df$Y)

  # "Main Stem Tag": tag of the largest stem of a multi-stemmed tree. Every
  # stem of the tree (the main one included) gets that tag as its group.
  if ("Main Stem Tag" %in% names(df) &&
      !("New Stem Grouping" %in% names(df) && any(!is.na(df$`New Stem Grouping`)))) {
    main <- trimws(as.character(df$`Main Stem Tag`))
    tag <- trimws(as.character(df$`New Tag No`))
    grp <- ifelse(!is.na(main) & nzchar(main), main, tag)
    df$`New Stem Grouping` <- grp
    # main stem first within its group, so consolidation keeps its tag
    df <- df[order(match(grp, unique(grp)), tag != grp), , drop = FALSE]
  }

  if (isTRUE(verbose)) {
    message("GigaFlor sheet: ", nrow(df), " stems; D converted from cm to mm; ",
            "X/Y read as longitude/latitude (decimal degrees).")
  }
  rownames(df) <- NULL
  df
}


#' Longitude/latitude (decimal degrees) to metres around their centroid
#'
#' Local tangent-plane approximation (WGS84 / SIRGAS 2000 metres per degree at
#' the mean latitude); error is negligible over a few hundred metres.
#'
#' @return Matrix with columns `east`, `north` (m).
#' @keywords internal
#' @noRd
.lonlat_to_local_m <- function(lon, lat, ref = NULL) {
  if (is.null(ref)) {
    ok <- is.finite(lon) & is.finite(lat)
    ref <- .lonlat_ref(mean(lon[ok]), mean(lat[ok]))
  }
  out <- cbind(east = (lon - ref$lon0) * ref$m_lon,
               north = (lat - ref$lat0) * ref$m_lat)
  attr(out, "ref") <- ref
  out
}

#' Reference point and metres per degree for `.lonlat_to_local_m()`
#' @keywords internal
#' @noRd
.lonlat_ref <- function(lon0, lat0) {
  phi <- lat0 * pi / 180
  list(lon0 = lon0, lat0 = lat0,
       m_lat = 111132.954 - 559.822 * cos(2 * phi) + 1.175 * cos(4 * phi),
       m_lon = 111412.84 * cos(phi) - 93.5 * cos(3 * phi))
}


#' Plot-grid coordinates (m) back to longitude/latitude with a GigaFlor fit
#'
#' Inverse of the rigid fit made by `.gigaflor_local_xy()`.
#'
#' @param x,y Global plot coordinates in metres (origin at the corner of
#'   subplot 1, as in `.compute_global_coordinates()`).
#' @param fit The `gigaflor_fit` attribute of `.gigaflor_local_xy()`.
#' @return Matrix with columns `lat`, `lon`.
#' @keywords internal
#' @noRd
.gigaflor_grid_to_lonlat <- function(x, y, fit) {
  if (is.null(fit$R)) {
    stop("No coordinate fit available to locate the plot.", call. = FALSE)
  }
  B <- cbind(as.numeric(x), as.numeric(y))
  A <- sweep(sweep(B, 2, fit$mb) %*% t(fit$R), 2, fit$ma, "+")
  cbind(lat = fit$ref$lat0 + A[, 2] / fit$ref$m_lat,
        lon = fit$ref$lon0 + A[, 1] / fit$ref$m_lon)
}


#' Plot corners estimated from the trees' GigaFlor coordinates
#'
#' The four corners in the order `forplot_map()` expects for `vertex_coords`:
#' (0, 0), (0, length), (width, 0), (width, length).
#'
#' @return Data frame with `latitude`, `longitude`.
#' @keywords internal
#' @noRd
.gigaflor_vertices <- function(fit, plot_width_m, plot_length_m) {
  ll <- .gigaflor_grid_to_lonlat(c(0, 0, plot_width_m, plot_width_m),
                                 c(0, plot_length_m, 0, plot_length_m), fit)
  data.frame(latitude = ll[, "lat"], longitude = ll[, "lon"])
}


#' Centre of each T1 subplot in the serpentine ForestPlots grid (m)
#' @keywords internal
#' @noRd
.t1_cell_origin <- function(T1, subplot_size, n_rows) {
  col <- floor((T1 - 1) / n_rows)
  row <- (T1 - 1) %% n_rows
  gy <- ifelse(col %% 2 == 0, row, n_rows - row - 1)
  cbind(x = col * subplot_size, y = gy * subplot_size)
}


#' Rigid fit (rotation, optional mirror, translation; no scaling) of A onto B
#' @keywords internal
#' @noRd
.rigid_fit <- function(A, B, mirror = FALSE) {
  ma <- colMeans(A)
  mb <- colMeans(B)
  A0 <- sweep(A, 2, ma)
  B0 <- sweep(B, 2, mb)
  sv <- svd(crossprod(A0, B0))
  R <- sv$u %*% t(sv$v)
  if ((det(R) < 0) != isTRUE(mirror)) {
    sv$v[, 2] <- -sv$v[, 2]
    R <- sv$u %*% t(sv$v)
  }
  fitted <- sweep(A0 %*% R, 2, mb, "+")
  list(R = R, ma = ma, mb = mb, fitted = fitted,
       rmse = sqrt(mean(rowSums((fitted - B)^2))))
}


#' Local X/Y (m within the T1 subplot) from GigaFlor decimal-degree coordinates
#'
#' The trees' GPS positions are projected to metres and fitted onto the plot
#' grid by a rigid transformation (rotation and translation, no scaling) that
#' brings each tree as close as possible to the centre of its own T1 subplot.
#' This recovers the orientation of the plot on the ground. The fitted
#' position is then expressed as X/Y within the subplot and kept inside it
#' (GPS error under the canopy can place a tree a few metres outside).
#'
#' Trees without coordinates are placed at the centre of their subplot.
#'
#' @param fp_sheet Output of `.gigaflor_prepare()`.
#' @param subplot_size,plot_width_m,plot_length_m Plot geometry (m).
#' @param verbose Logical.
#' @return The data frame with `X`, `Y` in metres within the subplot and the
#'   attribute `gigaflor_fit` (rotation in degrees, mirrored, RMSE in metres).
#' @keywords internal
#' @noRd
.gigaflor_local_xy <- function(fp_sheet, subplot_size, plot_width_m, plot_length_m,
                               verbose = TRUE) {
  df <- fp_sheet
  n_rows <- as.integer(round(plot_length_m / subplot_size))
  T1 <- suppressWarnings(as.numeric(.parse_num(df$T1)))
  lon <- .parse_num(df$Longitude)
  lat <- .parse_num(df$Latitude)

  bad <- is.finite(lon) & is.finite(lat) & (abs(lat) > 90 | abs(lon) > 180)
  if (any(bad)) {
    stop("GigaFlor X/Y must be longitude/latitude in decimal degrees; ",
         sum(bad), " row(s) are out of range (e.g. X = ", lon[bad][1],
         ", Y = ", lat[bad][1], ").", call. = FALSE)
  }

  origin <- .t1_cell_origin(T1, subplot_size, n_rows)
  centre <- origin + subplot_size / 2
  use <- is.finite(lon) & is.finite(lat) & is.finite(T1) & is.finite(centre[, 1])

  X <- rep(subplot_size / 2, nrow(df))
  Y <- rep(subplot_size / 2, nrow(df))
  fit_info <- list(rotation_deg = NA_real_, mirrored = NA, rmse_m = NA_real_, n = sum(use))

  if (sum(use) >= 3L && length(unique(T1[use])) >= 2L) {
    P <- .lonlat_to_local_m(lon[use], lat[use])
    fits <- list(.rigid_fit(P, centre[use, , drop = FALSE], mirror = FALSE),
                 .rigid_fit(P, centre[use, , drop = FALSE], mirror = TRUE))
    best <- fits[[which.min(vapply(fits, `[[`, numeric(1), "rmse"))]]
    loc <- best$fitted - origin[use, , drop = FALSE]
    eps <- subplot_size * 0.01
    X[use] <- pmin(pmax(loc[, 1], eps), subplot_size - eps)
    Y[use] <- pmin(pmax(loc[, 2], eps), subplot_size - eps)
    fit_info <- list(rotation_deg = round(atan2(best$R[2, 1], best$R[1, 1]) * 180 / pi, 1),
                     mirrored = det(best$R) < 0, rmse_m = round(best$rmse, 1),
                     n = sum(use),
                     R = best$R, ma = best$ma, mb = best$mb, ref = attr(P, "ref"))

    # Trees spread uniformly in their subplot sit, on average, s/sqrt(6) from
    # its centre; much more than that means the T1 numbering does not match
    # the serpentine grid or the coordinates are wrong.
    expected <- subplot_size / sqrt(6)
    if (best$rmse > 1.75 * expected) {
      warning("GigaFlor coordinates fit the T1 subplot grid poorly (RMSE ",
              round(best$rmse, 1), " m). Check that X = longitude, Y = latitude, ",
              "that T1 follows the serpentine numbering and the plot geometry ",
              "(`plot_size`, `subplot_size`).", call. = FALSE)
    }
  } else if (isTRUE(verbose)) {
    message("GigaFlor sheet: too few georeferenced stems; stems drawn at the centre of their subplot.")
  }

  n_nocoord <- sum(is.finite(T1) & !use)
  if (n_nocoord > 0 && isTRUE(verbose)) {
    message(n_nocoord, " stem(s) without coordinates drawn at the centre of their subplot.")
  }

  df$X <- round(X, 2)
  df$Y <- round(Y, 2)

  if (isTRUE(verbose) && is.finite(fit_info$rmse_m)) {
    message("GigaFlor coordinates fitted to the ", subplot_size, " m subplot grid: plot rotated ",
            fit_info$rotation_deg, " deg", if (isTRUE(fit_info$mirrored)) " (mirrored)" else "",
            ", RMS distance to subplot centre ", fit_info$rmse_m, " m (", fit_info$n, " stems).")
  }
  attr(df, "gigaflor_fit") <- fit_info
  df
}
