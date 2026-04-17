# =============================================================================
# Auxiliary functions for fp_herb_converter
# Author: Giulia Ottino & Domingos Cardoso
# Revised: 08/Mar/2026
# =============================================================================

#' Resolve cache paths for herbarium lookup
#'
#' @return Named list with cache directory and DuckDB path.
#'
#' @keywords internal
#' @noRd
.herbaria_cache_paths <- function() {
  cache_dir <- getOption("forplotR.herbaria_cache_dir")

  if (!(is.character(cache_dir) && length(cache_dir) == 1L && nzchar(cache_dir))) {
    cache_dir <- tryCatch(
      tools::R_user_dir("forplotR", which = "cache"),
      error = function(e) NULL
    )
  }

  if (is.null(cache_dir) || !nzchar(cache_dir)) {
    cache_dir <- file.path(getwd(), ".forplotR_cache")
  }

  if (!dir.exists(cache_dir)) {
    dir.create(cache_dir, recursive = TRUE, showWarnings = FALSE)
  }

  list(
    cache_dir = cache_dir,
    duckdb_path = file.path(cache_dir, "herbaria_index.duckdb")
  )
}


#' Extract a collector number token from free text
#'
#' @param x Character vector.
#'
#' @return Character vector with extracted number token.
#'
#' @keywords internal
#' @noRd
.extract_number_token <- function(x) {
  if (is.na(x) || !nzchar(trimws(x))) {
    return(NA_character_)
  }

  x <- as.character(x)

  m <- regexec("(\\d+)\\s*$", x, perl = TRUE)
  if (m[[1]][1] > 0L) {
    hit <- regmatches(x, m)[[1]]
    if (length(hit) > 1L && nzchar(hit[2])) {
      return(trimws(hit[2]))
    }
  }

  nums <- regmatches(x, gregexpr("\\d+", x, perl = TRUE))[[1]]
  if (length(nums)) {
    return(nums[which.max(nchar(nums))])
  }

  NA_character_
}


#' Keep only digits in collector numbers
#'
#' @param x Character vector.
#'
#' @return Character vector with only digits.
#'
#' @keywords internal
#' @noRd
.normalize_number_token <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- ""
  x <- gsub("[^0-9]", "", x)
  x[!nzchar(x)] <- NA_character_
  x
}


#' Split recordedBy into candidate people
#'
#' @param recordedBy Character scalar.
#'
#' @return Character vector of candidate person strings.
#'
#' @keywords internal
#' @noRd
.split_recordedby_people <- function(recordedBy) {
  if (is.na(recordedBy) || !nzchar(trimws(recordedBy))) {
    return(character(0))
  }

  x <- as.character(recordedBy)
  x <- iconv(x, to = "ASCII//TRANSLIT", sub = "")
  x[is.na(x)] <- ""
  x <- trimws(x)

  if (!nzchar(x)) {
    return(character(0))
  }

  x <- gsub("\\bet\\s+al\\b\\.?", "", x, ignore.case = TRUE)
  x <- gsub("\\s+", " ", x)

  parts <- unlist(strsplit(
    x,
    "\\s*;\\s*|\\s*\\|\\s*|\\s*&\\s*|\\s+and\\s+",
    perl = TRUE
  ))
  parts <- trimws(parts)
  parts <- parts[nzchar(parts)]

  out <- character(0)

  for (p in parts) {
    comma_n <- lengths(regmatches(p, gregexpr(",", p, fixed = TRUE)))

    if (comma_n <= 1L && grepl("^[^,]+,[^,]+$", p)) {
      out <- c(out, p)
    } else if (comma_n >= 1L) {
      pp <- unlist(strsplit(p, "\\s*,\\s*", perl = TRUE))
      pp <- trimws(pp)
      pp <- pp[nzchar(pp)]
      out <- c(out, pp)
    } else {
      out <- c(out, p)
    }
  }

  out <- trimws(out)
  unique(out[nzchar(out)])
}


#' Build collector tokens from one person name
#'
#' @param name_raw Character scalar.
#'
#' @return Named character vector with primary and fallback tokens.
#'
#' @keywords internal
#' @noRd
.collector_tokens_one <- function(name_raw) {
  if (is.na(name_raw) || !nzchar(trimws(name_raw))) {
    return(c(primary = NA_character_, fallback = NA_character_))
  }

  x <- as.character(name_raw)
  x <- iconv(x, to = "ASCII//TRANSLIT", sub = "")
  x[is.na(x)] <- ""
  x <- trimws(tolower(x))
  x <- gsub("\\bet\\s+al\\b\\.?", "", x)
  x <- gsub("\\s+", " ", x)

  if (!nzchar(x)) {
    return(c(primary = NA_character_, fallback = NA_character_))
  }

  particles <- c("de", "da", "do", "das", "dos", "del", "della", "van", "von")

  words_clean <- function(z) {
    z <- gsub("[^a-z ]", " ", z)
    z <- unlist(strsplit(z, "\\s+"))
    z <- z[nzchar(z)]
    z <- z[!(z %in% particles)]
    z
  }

  surname <- ""
  initials <- ""

  if (grepl(",", x, fixed = TRUE)) {
    parts <- strsplit(x, ",", fixed = TRUE)[[1]]
    left <- words_clean(parts[1])
    right <- if (length(parts) >= 2L) words_clean(parts[2]) else character(0)

    if (length(left)) {
      surname <- left[length(left)]
    }
    if (length(right)) {
      initials <- paste(substr(right, 1L, 1L), collapse = "")
    }
  } else {
    w <- words_clean(x)
    if (length(w)) {
      surname <- w[length(w)]
      if (length(w) > 1L) {
        initials <- paste(substr(w[-length(w)], 1L, 1L), collapse = "")
      }
    }
  }

  surname <- gsub("[^a-z0-9]", "", surname)
  initials <- gsub("[^a-z0-9]", "", initials)

  fallback <- if (nzchar(surname)) surname else NA_character_
  primary <- if (nzchar(surname)) paste0(initials, surname) else NA_character_

  if (!nzchar(primary)) {
    primary <- fallback
  }

  c(primary = primary, fallback = fallback)
}


#' Expand compact collector code prefixes
#'
#' @param prefix Character scalar prefix extracted from a compact voucher.
#' @param collector_codes Optional named character vector mapping prefixes to collector names.
#'
#' @return Character scalar or NA.
#'
#' @keywords internal
#' @noRd
.expand_collector_code <- function(prefix, collector_codes = NULL) {
  if (is.null(collector_codes) || !length(collector_codes)) {
    return(NA_character_)
  }

  if (is.null(names(collector_codes)) || !length(names(collector_codes))) {
    stop(
      "`collector_codes` must be a named character vector. ",
      "Example: c('DC' = 'D. Cardoso', 'PWM' = 'P. W. Moonlight').",
      call. = FALSE
    )
  }

  nm <- toupper(trimws(names(collector_codes)))
  val <- as.character(collector_codes)
  key <- toupper(trimws(as.character(prefix)))

  hit <- match(key, nm, nomatch = 0L)
  if (hit > 0L) {
    out <- trimws(val[hit])
    if (nzchar(out)) {
      return(out)
    }
  }

  NA_character_
}


#' Parse one voucher string into collector and number
#'
#' @param voucher Character scalar.
#' @param collector_fallback Optional fallback collector name.
#' @param collector_codes Optional named character vector mapping compact prefixes to collector names.
#'
#' @return Named character vector with collector and number.
#'
#' @keywords internal
#' @noRd
.parse_voucher_one <- function(voucher,
                               collector_fallback = NULL,
                               collector_codes = NULL) {
  v <- as.character(voucher)
  v[is.na(v)] <- ""
  v <- trimws(v)

  if (!nzchar(v)) {
    return(c(collector = NA_character_, number = NA_character_))
  }

  if (grepl("^\\d+$", v)) {
    coll <- if (!is.null(collector_fallback) && nzchar(trimws(collector_fallback))) {
      trimws(collector_fallback)
    } else {
      NA_character_
    }
    return(c(collector = coll, number = v))
  }

  m_compact <- regexec("^([A-Za-z]+)\\s*([0-9]+)$", v, perl = TRUE)
  if (m_compact[[1]][1] > 0L) {
    hit <- regmatches(v, m_compact)[[1]]
    coll <- .expand_collector_code(hit[2], collector_codes = collector_codes)

    if (is.na(coll) && !is.null(collector_fallback) && nzchar(trimws(collector_fallback))) {
      coll <- trimws(collector_fallback)
    }

    return(c(collector = coll, number = hit[3]))
  }

  v2 <- gsub("[/;|]+", " ", v)
  v2 <- gsub("\\bet\\s+al\\b\\.?", "", v2, ignore.case = TRUE)
  v2 <- gsub("\\s+", " ", trimws(v2))

  m <- regexec("^(.*?)[[:space:]]*(\\d+)\\s*$", v2, perl = TRUE)
  if (m[[1]][1] > 0L) {
    hit <- regmatches(v2, m)[[1]]
    coll <- trimws(hit[2])
    num <- trimws(hit[3])

    if (!nzchar(coll) && !is.null(collector_fallback) && nzchar(trimws(collector_fallback))) {
      coll <- trimws(collector_fallback)
    }

    return(c(
      collector = if (nzchar(coll)) coll else NA_character_,
      number = if (nzchar(num)) num else NA_character_
    ))
  }

  c(
    collector = if (!is.null(collector_fallback) && nzchar(trimws(collector_fallback))) {
      trimws(collector_fallback)
    } else {
      NA_character_
    },
    number = .extract_number_token(v2)
  )
}


#' Parse census vouchers into normalized match keys
#'
#' @param voucher Character vector of voucher values.
#' @param collector_fallback Optional collector fallback.
#' @param collector_codes Optional named character vector mapping compact prefixes to collector names.
#'
#' @return Data frame with raw and normalized voucher identity fields.
#'
#' @keywords internal
#' @noRd
.parse_census_identity <- function(voucher,
                                   collector_fallback = NULL,
                                   collector_codes = NULL) {
  voucher <- as.character(voucher)
  voucher[is.na(voucher)] <- ""

  parsed <- lapply(
    voucher,
    .parse_voucher_one,
    collector_fallback = collector_fallback,
    collector_codes = collector_codes
  )

  collector_raw <- vapply(parsed, `[[`, character(1), "collector")
  number_raw <- vapply(parsed, `[[`, character(1), "number")
  number_clean <- .normalize_number_token(number_raw)

  tok <- t(vapply(
    collector_raw,
    .collector_tokens_one,
    FUN.VALUE = c(primary = NA_character_, fallback = NA_character_)
  ))

  primary_key <- ifelse(
    !is.na(tok[, "primary"]) & nzchar(tok[, "primary"]) &
      !is.na(number_clean) & nzchar(number_clean),
    paste0(tok[, "primary"], "||", number_clean),
    NA_character_
  )

  fallback_key <- ifelse(
    !is.na(tok[, "fallback"]) & nzchar(tok[, "fallback"]) &
      !is.na(number_clean) & nzchar(number_clean),
    paste0(tok[, "fallback"], "||", number_clean),
    NA_character_
  )

  data.frame(
    collector_raw = collector_raw,
    number_raw = number_raw,
    number_clean = number_clean,
    primary_key = primary_key,
    fallback_key = fallback_key,
    stringsAsFactors = FALSE
  )
}


#' Convert one IPT resource id into a guessed herbarium code
#'
#' @param resource_id Character scalar.
#'
#' @return Character scalar.
#'
#' @keywords internal
#' @noRd
.resource_guess_herbarium <- function(resource_id) {
  rid <- tolower(trimws(as.character(resource_id)))
  rid <- sub("^.*[?&]r=", "", rid)
  rid <- sub("&.*$", "", rid)

  out <- sub("^jbrj_", "", rid)
  out <- sub("_herbarium$", "", out)
  out <- sub("_reflora$", "", out)
  out <- sub("_jabot$", "", out)
  out <- sub("_hv$", "", out)
  out <- sub("_.*$", "", out)
  out <- toupper(out)

  if (grepl("(^|_)rb($|_)", rid)) {
    out <- "RB"
  }

  out
}


#' Read latest version info from an IPT resource page
#'
#' @param resource_url Character scalar.
#'
#' @return Named list with version, published_on and records.
#'
#' @keywords internal
#' @noRd
.get_latest_version_info <- function(resource_url) {
  x <- tryCatch(
    suppressWarnings(readLines(resource_url, encoding = "UTF-8", warn = FALSE)),
    error = function(e) character(0)
  )

  out <- list(
    version = NA_character_,
    published_on = NA_character_,
    records = NA_character_
  )

  if (!length(x)) {
    return(out)
  }

  txt <- paste(x, collapse = " ")

  ver <- sub(
    ".*latestVersion[^0-9]*([0-9]+(?:\\.[0-9]+)+).*",
    "\\1",
    txt,
    perl = TRUE
  )
  if (!identical(ver, txt)) {
    out$version <- ver
  }

  dates <- regmatches(txt, gregexpr("\\d{4}-\\d{2}-\\d{2}", txt, perl = TRUE))[[1]]
  if (length(dates)) {
    out$published_on <- dates[1]
  }

  recs <- regmatches(txt, gregexpr("\\b[0-9][0-9,]*\\b", txt, perl = TRUE))[[1]]
  if (length(recs)) {
    out$records <- recs[1]
  }

  out
}


#' Discover IPT resources for one or more herbaria
#'
#' @param herbarium Character vector of herbarium codes.
#' @param ipt One of "jabot" or "reflora".
#' @param resource_map Optional named character vector mapping "HERBARIUM::source" to resource_id.
#'
#' @return Data frame with discovered resources.
#'
#' @keywords internal
#' @noRd
.get_ipt_info <- function(herbarium,
                          ipt = c("jabot", "reflora"),
                          resource_map = NULL) {
  ipt <- match.arg(ipt)
  .arg_check_herbarium(herbarium)

  herbarium <- unique(toupper(trimws(as.character(herbarium))))

  empty <- data.frame(
    ipt = character(0),
    herbarium = character(0),
    resource_id = character(0),
    archive_base = character(0),
    resource_url = character(0),
    stringsAsFactors = FALSE
  )

  dcat_roots <- if (ipt == "jabot") {
    c("jabot", "jbrj")
  } else {
    "reflora"
  }

  out_list <- vector("list", length(dcat_roots))
  pos <- 0L

  for (root in dcat_roots) {
    dcat_url <- paste0("https://ipt.jbrj.gov.br/", root, "/dcat")

    lines <- tryCatch(
      suppressWarnings(readLines(dcat_url, encoding = "UTF-8", warn = FALSE)),
      error = function(e) character(0)
    )

    if (!length(lines)) {
      next
    }

    txt <- paste(lines, collapse = "\n")

    dl_pat <- paste0(
      "https://ipt\\.jbrj\\.gov\\.br/", root,
      "/archive\\.do\\?r=[A-Za-z0-9_\\-]+(?:&v=[0-9.]+)?"
    )

    download_urls <- unique(unlist(regmatches(txt, gregexpr(dl_pat, txt, perl = TRUE))))

    if (!length(download_urls)) {
      next
    }

    resource_ids <- sub("^.*[?&]r=", "", download_urls)
    resource_ids <- sub("&.*$", "", resource_ids)

    herb_guess <- vapply(resource_ids, .resource_guess_herbarium, character(1))

    one <- data.frame(
      ipt = rep(ipt, length(resource_ids)),
      herbarium = herb_guess,
      resource_id = resource_ids,
      archive_base = paste0("https://ipt.jbrj.gov.br/", root, "/archive.do?r="),
      resource_url = paste0("https://ipt.jbrj.gov.br/", root, "/resource?r=", resource_ids),
      stringsAsFactors = FALSE
    )

    one <- one[one$herbarium %in% herbarium, , drop = FALSE]

    if (nrow(one)) {
      pos <- pos + 1L
      out_list[[pos]] <- one
    }
  }

  out <- if (pos) {
    do.call(rbind, out_list[seq_len(pos)])
  } else {
    empty
  }

  if (!is.null(resource_map) && length(resource_map)) {
    if (is.null(names(resource_map)) || !length(names(resource_map))) {
      stop(
        "`herbaria_resource_map` must be a named character vector. ",
        "Example: c('RB::jabot' = 'jbrj_rb').",
        call. = FALSE
      )
    }

    wanted_names <- paste0(herbarium, "::", ipt)
    hit <- names(resource_map) %in% wanted_names

    if (any(hit)) {
      rid <- as.character(resource_map[hit])
      root <- ifelse(grepl("^jbrj_", rid), "jbrj", ipt)

      override_df <- data.frame(
        ipt = ipt,
        herbarium = sub("::.*$", "", names(resource_map)[hit]),
        resource_id = rid,
        archive_base = paste0("https://ipt.jbrj.gov.br/", root, "/archive.do?r="),
        resource_url = paste0("https://ipt.jbrj.gov.br/", root, "/resource?r=", rid),
        stringsAsFactors = FALSE
      )

      out <- rbind(out, override_df)
    }
  }

  if (!nrow(out)) {
    return(empty)
  }

  out <- out[!duplicated(paste(out$archive_base, out$resource_id, sep = "::")), , drop = FALSE]

  if (ipt == "jabot") {
    prio <- ifelse(grepl("/jbrj/", out$archive_base, fixed = FALSE), 0L, 1L)
    out <- out[order(out$herbarium, prio, out$resource_id), , drop = FALSE]
    out <- out[!duplicated(out$herbarium), , drop = FALSE]
  }

  rownames(out) <- NULL
  out
}


#' Download one DwC-A archive and extract occurrence.txt
#'
#' @param info_row One-row data frame returned by .get_ipt_info().
#' @param dir Output directory.
#' @param verbose Logical.
#' @param force_refresh Logical.
#'
#' @return Character path to extracted archive directory, or NULL.
#'
#' @keywords internal
#' @noRd
.download_dwca_one <- function(info_row,
                               dir,
                               verbose = FALSE,
                               force_refresh = FALSE) {
  stopifnot(is.data.frame(info_row), nrow(info_row) == 1L)

  if (!dir.exists(dir)) {
    dir.create(dir, recursive = TRUE, showWarnings = FALSE)
  }

  safe_id <- gsub("[^A-Za-z0-9]+", "_", info_row$resource_id)
  out_dir <- file.path(
    dir,
    paste0("dwca_", info_row$ipt, "_", info_row$herbarium, "_", safe_id)
  )

  if (isTRUE(force_refresh) && dir.exists(out_dir)) {
    unlink(out_dir, recursive = TRUE, force = TRUE)
  }

  if (dir.exists(out_dir) && file.exists(file.path(out_dir, "occurrence.txt"))) {
    return(out_dir)
  }

  if (!dir.exists(out_dir)) {
    dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)
  }

  vinfo <- .get_latest_version_info(info_row$resource_url)
  archive_url <- paste0(info_row$archive_base, info_row$resource_id)

  if (!is.na(vinfo$version) && nzchar(vinfo$version)) {
    archive_url <- paste0(
      archive_url,
      "&v=",
      utils::URLencode(vinfo$version, reserved = TRUE)
    )
  }

  zipfile <- tempfile(pattern = "dwca_", tmpdir = dir, fileext = ".zip")
  old_timeout <- getOption("timeout")
  options(timeout = max(600, old_timeout))
  on.exit(options(timeout = old_timeout), add = TRUE)

  if (isTRUE(verbose)) {
    message("Downloading DwC-A: ", info_row$herbarium, " [", info_row$ipt, "]")
  }

  ok <- tryCatch({
    suppressWarnings(
      utils::download.file(
        archive_url,
        zipfile,
        mode = "wb",
        quiet = !isTRUE(verbose)
      )
    )
    TRUE
  }, error = function(e) FALSE)

  if (!ok && grepl("&v=", archive_url, fixed = TRUE)) {
    ok <- tryCatch({
      suppressWarnings(
        utils::download.file(
          sub("&v=.*$", "", archive_url),
          zipfile,
          mode = "wb",
          quiet = !isTRUE(verbose)
        )
      )
      TRUE
    }, error = function(e) FALSE)
  }

  if (!ok || !file.exists(zipfile)) {
    unlink(out_dir, recursive = TRUE, force = TRUE)
    return(NULL)
  }

  unzip_ok <- tryCatch({
    utils::unzip(zipfile, exdir = out_dir)
    TRUE
  }, error = function(e) FALSE)
  unlink(zipfile)

  if (!unzip_ok || !file.exists(file.path(out_dir, "occurrence.txt"))) {
    unlink(out_dir, recursive = TRUE, force = TRUE)
    return(NULL)
  }

  out_dir
}


#' Connect to herbarium DuckDB cache
#'
#' @return DuckDB connection.
#'
#' @keywords internal
#' @noRd
.herbaria_db_connect <- function() {
  for (pkg in c("DBI", "duckdb")) {
    if (!requireNamespace(pkg, quietly = TRUE)) {
      stop(
        "Package '", pkg, "' is required for herbarium lookup.",
        call. = FALSE
      )
    }
  }

  db_path <- .herbaria_cache_paths()$duckdb_path
  con <- DBI::dbConnect(duckdb::duckdb(), dbdir = db_path)

  DBI::dbExecute(con, "
    CREATE TABLE IF NOT EXISTS occ_raw (
      catalogNumber VARCHAR,
      occurrenceID VARCHAR,
      recordedBy VARCHAR,
      recordNumber VARCHAR,
      recordNumber_clean VARCHAR,
      herbarium VARCHAR,
      source VARCHAR,
      resource_id VARCHAR
    );
  ")

  DBI::dbExecute(con, "
    CREATE TABLE IF NOT EXISTS herbaria_index (
      key VARCHAR,
      key_type VARCHAR,
      catalogNumber VARCHAR,
      occurrenceID VARCHAR,
      herbarium VARCHAR,
      source VARCHAR,
      resource_id VARCHAR
    );
  ")

  try(
    DBI::dbExecute(
      con,
      "CREATE INDEX IF NOT EXISTS idx_occ_num ON occ_raw(herbarium, source, resource_id, recordNumber_clean);"
    ),
    silent = TRUE
  )

  try(
    DBI::dbExecute(
      con,
      "CREATE INDEX IF NOT EXISTS idx_idx_key ON herbaria_index(key_type, key);"
    ),
    silent = TRUE
  )

  con
}


#' Load occurrence.txt into DuckDB
#'
#' @param con DuckDB connection.
#' @param occurrence_path Character path to occurrence.txt.
#' @param herbarium Herbarium code.
#' @param source One of "jabot" or "reflora".
#' @param resource_id IPT resource id.
#' @param verbose Logical.
#'
#' @return Logical.
#'
#' @keywords internal
#' @noRd
.duckdb_load_occurrence <- function(con,
                                    occurrence_path,
                                    herbarium,
                                    source,
                                    resource_id,
                                    verbose = FALSE) {
  if (!file.exists(occurrence_path)) {
    return(FALSE)
  }

  herb_q <- as.character(DBI::dbQuoteString(con, herbarium))
  src_q  <- as.character(DBI::dbQuoteString(con, source))
  rid_q  <- as.character(DBI::dbQuoteString(con, resource_id))
  occ_q  <- as.character(DBI::dbQuoteString(
    con,
    normalizePath(occurrence_path, winslash = "/", mustWork = FALSE)
  ))

  DBI::dbExecute(con, paste0(
    "DELETE FROM occ_raw WHERE herbarium = ", herb_q,
    " AND source = ", src_q,
    " AND resource_id = ", rid_q, ";"
  ))

  DBI::dbExecute(con, paste0(
    "DELETE FROM herbaria_index WHERE herbarium = ", herb_q,
    " AND source = ", src_q,
    " AND resource_id = ", rid_q, ";"
  ))

  probe <- tryCatch(
    DBI::dbGetQuery(
      con,
      paste0(
        "SELECT * FROM read_csv_auto(",
        occ_q,
        ", delim = '\\t', header = TRUE, quote = '', sample_size = 20000, ignore_errors = TRUE) LIMIT 0;"
      )
    ),
    error = function(e) NULL
  )

  if (is.null(probe)) {
    return(FALSE)
  }

  nms_raw <- names(probe)
  nms <- .norm_nm(nms_raw)

  pick_first <- function(cands) {
    cands <- .norm_nm(cands)
    hit <- match(cands, nms, nomatch = 0L)
    if (any(hit > 0L)) {
      return(nms_raw[hit[hit > 0L][1]])
    }
    for (z in cands) {
      ii <- which(grepl(z, nms, fixed = TRUE))
      if (length(ii)) {
        return(nms_raw[ii[1]])
      }
    }
    NA_character_
  }

  catalog_col <- pick_first(c("catalogNumber", "specimen", "barcode", "accession"))
  occid_col <- pick_first(c("occurrenceID", "id"))
  recorded_col <- pick_first(c("recordedBy", "collector", "coletor", "coletores"))
  record_col <- pick_first(c(
    "recordNumber", "collectorNumber", "voucherNumber",
    "verbatimCollectorNumber", "fieldNumber", "otherCatalogNumbers", "numero"
  ))

  sel <- function(col, alias) {
    if (is.na(col)) {
      paste0("NULL::VARCHAR AS ", alias)
    } else {
      paste0('CAST("', gsub('"', '""', col, fixed = TRUE), '" AS VARCHAR) AS ', alias)
    }
  }

  rec_expr <- if (is.na(record_col)) {
    "NULL::VARCHAR"
  } else {
    paste0('CAST("', gsub('"', '""', record_col, fixed = TRUE), '" AS VARCHAR)')
  }

  scan_sql <- paste0(
    "INSERT INTO occ_raw ",
    "SELECT ",
    sel(catalog_col, "catalogNumber"), ", ",
    sel(occid_col, "occurrenceID"), ", ",
    sel(recorded_col, "recordedBy"), ", ",
    sel(record_col, "recordNumber"), ", ",
    "regexp_replace(coalesce(", rec_expr, ", ''), '[^0-9]', '', 'g') AS recordNumber_clean, ",
    herb_q, " AS herbarium, ",
    src_q, " AS source, ",
    rid_q, " AS resource_id ",
    "FROM read_csv_auto(",
    occ_q,
    ", delim = '\\t', header = TRUE, quote = '', sample_size = 20000, ignore_errors = TRUE);"
  )

  ok <- tryCatch({
    DBI::dbExecute(con, scan_sql)
    TRUE
  }, error = function(e) FALSE)

  if (isTRUE(verbose) && isTRUE(ok)) {
    n_ins <- DBI::dbGetQuery(con, paste0(
      "SELECT COUNT(*) AS n FROM occ_raw WHERE herbarium = ", herb_q,
      " AND source = ", src_q,
      " AND resource_id = ", rid_q, ";"
    ))$n
    message("Inserted ", n_ins, " records into database")
  }

  ok
}


#' Match one resource against requested collector-number keys
#'
#' @param con DuckDB connection.
#' @param herbarium Herbarium code.
#' @param source Source string.
#' @param resource_id IPT resource id.
#' @param need_primary Character vector of requested primary keys.
#' @param need_fallback Character vector of requested fallback keys.
#' @param verbose Logical.
#'
#' @return Logical.
#'
#' @keywords internal
#' @noRd
.duckdb_match_resource <- function(con,
                                   herbarium,
                                   source,
                                   resource_id,
                                   need_primary,
                                   need_fallback,
                                   verbose = FALSE) {
  need_primary <- unique(stats::na.omit(as.character(need_primary)))
  need_fallback <- unique(stats::na.omit(as.character(need_fallback)))

  all_numbers <- unique(c(
    sub("^.*\\|\\|", "", need_primary),
    sub("^.*\\|\\|", "", need_fallback)
  ))
  all_numbers <- all_numbers[nzchar(all_numbers)]

  if (!length(all_numbers)) {
    return(FALSE)
  }

  herb_q <- as.character(DBI::dbQuoteString(con, herbarium))
  src_q  <- as.character(DBI::dbQuoteString(con, source))
  rid_q  <- as.character(DBI::dbQuoteString(con, resource_id))
  nums_sql <- paste(sprintf("'%s'", gsub("'", "''", all_numbers)), collapse = ", ")

  qry <- paste0(
    "SELECT catalogNumber, occurrenceID, recordedBy, recordNumber, recordNumber_clean ",
    "FROM occ_raw ",
    "WHERE herbarium = ", herb_q,
    " AND source = ", src_q,
    " AND resource_id = ", rid_q,
    " AND recordNumber_clean IN (", nums_sql, ");"
  )

  cand <- tryCatch(DBI::dbGetQuery(con, qry), error = function(e) NULL)

  if (is.null(cand) || !nrow(cand)) {
    if (isTRUE(verbose)) {
      message("No candidate records found for resource ", resource_id)
    }
    return(FALSE)
  }

  if (isTRUE(verbose)) {
    message("Candidate records found: ", nrow(cand))
  }

  idx_list <- vector("list", nrow(cand))
  pos <- 0L

  for (i in seq_len(nrow(cand))) {
    people <- .split_recordedby_people(cand$recordedBy[i])
    if (!length(people)) {
      next
    }

    tok <- t(vapply(
      people,
      .collector_tokens_one,
      FUN.VALUE = c(primary = NA_character_, fallback = NA_character_)
    ))

    primary_keys <- ifelse(
      !is.na(tok[, "primary"]) & nzchar(tok[, "primary"]),
      paste0(tok[, "primary"], "||", cand$recordNumber_clean[i]),
      NA_character_
    )

    fallback_keys <- ifelse(
      !is.na(tok[, "fallback"]) & nzchar(tok[, "fallback"]),
      paste0(tok[, "fallback"], "||", cand$recordNumber_clean[i]),
      NA_character_
    )

    one <- rbind(
      data.frame(
        key = primary_keys,
        key_type = "primary",
        catalogNumber = cand$catalogNumber[i],
        occurrenceID = cand$occurrenceID[i],
        herbarium = herbarium,
        source = source,
        resource_id = resource_id,
        stringsAsFactors = FALSE
      ),
      data.frame(
        key = fallback_keys,
        key_type = "fallback",
        catalogNumber = cand$catalogNumber[i],
        occurrenceID = cand$occurrenceID[i],
        herbarium = herbarium,
        source = source,
        resource_id = resource_id,
        stringsAsFactors = FALSE
      )
    )

    one <- one[!is.na(one$key) & nzchar(one$key), , drop = FALSE]
    if (!nrow(one)) {
      next
    }

    pos <- pos + 1L
    idx_list[[pos]] <- one
  }

  if (!pos) {
    return(FALSE)
  }

  idx <- do.call(rbind, idx_list[seq_len(pos)])
  idx <- unique(idx)

  keep <- rbind(
    if (length(need_primary)) {
      data.frame(key = need_primary, key_type = "primary", stringsAsFactors = FALSE)
    },
    if (length(need_fallback)) {
      data.frame(key = need_fallback, key_type = "fallback", stringsAsFactors = FALSE)
    }
  )

  idx <- merge(idx, keep, by = c("key", "key_type"), all = FALSE)
  idx <- idx[!duplicated(paste(
    idx$key, idx$key_type, idx$catalogNumber, idx$occurrenceID, idx$source
  )), , drop = FALSE]

  if (!nrow(idx)) {
    if (isTRUE(verbose)) {
      message("Candidates found, but none matched requested keys")
    }
    return(FALSE)
  }

  DBI::dbWriteTable(con, "herbaria_index", idx, append = TRUE)

  if (isTRUE(verbose)) {
    message("Index built with ", nrow(idx), " matched rows")
  }

  TRUE
}


#' Build a JABOT specimen URL
#'
#' @param catalogNumber Character scalar.
#' @param herbarium Character scalar.
#'
#' @return Character scalar URL or NA.
#'
#' @keywords internal
#' @noRd
.make_jabot_url <- function(catalogNumber, herbarium) {
  cat_str <- trimws(as.character(catalogNumber))
  if (is.na(catalogNumber) || !nzchar(cat_str)) {
    return(NA_character_)
  }

  herb_str <- toupper(trimws(as.character(herbarium)))
  if (is.na(herbarium) || !nzchar(herb_str)) {
    return(NA_character_)
  }

  paste0(
    "https://catalogo-ucs-brasil.jbrj.gov.br/regua/visualizador.php",
    "?codtestemunho=",
    utils::URLencode(cat_str, reserved = TRUE),
    "&colbot=",
    herb_str,
    "&db=Jabot&r=true"
  )
}


#' Build a REFLORA image URL
#'
#' @param occurrenceID Character scalar. Aceita URI completa ou ID numerico puro.
#'
#' @return Character scalar URL or NA.
#'
#' @keywords internal
#' @noRd
.make_reflora_url <- function(occurrenceID) {
  if (is.na(occurrenceID) || !nzchar(trimws(as.character(occurrenceID)))) {
    return(NA_character_)
  }

  id_raw <- trimws(as.character(occurrenceID))

  # Caso 1: URI com segmento numerico final  (ex: ".../ExibeFiguraFSIUC/12345")
  m_uri <- regexec("/([0-9]+)\\s*$", id_raw, perl = TRUE)
  if (m_uri[[1]][1] > 0L) {
    id <- regmatches(id_raw, m_uri)[[1]][2]
    if (nzchar(id)) {
      return(paste0(
        "https://reflora.jbrj.gov.br/reflora/geral/ExibeFiguraFSIUC/ExibeFiguraFSIUC.do?idFigura=",
        id
      ))
    }
  }

  # Caso 2: ID puramente numerico (sem nenhum outro caractere)
  if (grepl("^[0-9]+$", id_raw)) {
    return(paste0(
      "https://reflora.jbrj.gov.br/reflora/geral/ExibeFiguraFSIUC/ExibeFiguraFSIUC.do?idFigura=",
      id_raw
    ))
  }

  # Caso 3: fallback - maior bloco numerico continuo (minimo 3 digitos)
  # Cobre formatos como "RB00012345" onde nao ha separador "/"
  nums <- regmatches(id_raw, gregexpr("[0-9]{3,}", id_raw, perl = TRUE))[[1]]
  if (length(nums)) {
    id <- nums[which.max(nchar(nums))]
    return(paste0(
      "https://reflora.jbrj.gov.br/reflora/geral/ExibeFiguraFSIUC/ExibeFiguraFSIUC.do?idFigura=",
      id
    ))
  }

  NA_character_
}


#' Build herbarium links for voucher records using JABOT first and REFLORA only if needed
#'
#' @description
#' Matches voucher records from a field data frame against locally indexed
#' Darwin Core Archive (DwC-A) occurrence data and returns HTML fragments
#' linking to herbarium specimen pages.
#'
#' The lookup workflow is:
#' \enumerate{
#'   \item Parse voucher strings into normalized collector-number keys.
#'   \item Download and index JABOT resources for the requested herbaria.
#'   \item Resolve as many vouchers as possible using JABOT records.
#'   \item Only for still unresolved vouchers, download and index REFLORA
#'         resources and try again.
#' }
#'
#' JABOT links are built from `catalogNumber`, while REFLORA links are built
#' from `occurrenceID`.
#'
#' @param fp_df A data frame containing at least a `Voucher` column.
#' @param herbaria Character vector of herbarium codes, for example `"RB"`.
#' @param force_refresh Logical. If `TRUE`, forces redownload and reindexing
#'   of DwC-A resources.
#' @param keep_downloads Logical. If `FALSE`, temporary downloaded DwC-A folders
#'   are removed at the end of the function.
#' @param verbose Logical. If `TRUE`, prints detailed progress messages.
#' @param collector_fallback Optional character scalar. Collector name used when
#'   the voucher string does not explicitly include one.
#' @param collector_codes Optional named character vector mapping compact
#'   collector prefixes to full collector names.
#' @param resource_map Optional named character vector overriding IPT resource
#'   ids, using names such as `"RB::jabot"` or `"RB::reflora"`.
#'
#' @return A character vector of length `nrow(fp_df)`. Each element is either
#'   `NA_character_` or an HTML fragment containing a herbarium link.
#'
#' @keywords internal
#' @noRd
.herbaria_lookup_links <- function(fp_df,
                                   herbaria,
                                   force_refresh = FALSE,
                                   keep_downloads = FALSE,
                                   collector_fallback = NULL,
                                   collector_codes = NULL,
                                   resource_map = NULL,
                                   verbose = FALSE) {
  .arg_check_herbarium(herbaria)
  herbaria <- unique(toupper(trimws(as.character(herbaria))))

  if (!("Voucher" %in% names(fp_df))) {
    return(rep(NA_character_, nrow(fp_df)))
  }

  # Parse voucher strings into normalized identity keys.
  parsed <- .parse_census_identity(
    voucher = fp_df$Voucher,
    collector_fallback = collector_fallback,
    collector_codes = collector_codes
  )

  need_primary <- unique(stats::na.omit(parsed$primary_key))
  need_fallback <- unique(stats::na.omit(parsed$fallback_key))

  # Nothing to match.
  if (!length(need_primary) && !length(need_fallback)) {
    return(rep(NA_character_, nrow(fp_df)))
  }

  if (isTRUE(verbose)) {
    message("HERBARIA LOOKUP")
    message("Herbaria: ", paste(herbaria, collapse = ", "))
    message("Records: ", nrow(fp_df))
    message("Valid primary keys: ", length(need_primary))
    message("Valid fallback keys: ", length(need_fallback))
  }

  con <- .herbaria_db_connect()
  on.exit(try(DBI::dbDisconnect(con, shutdown = TRUE), silent = TRUE), add = TRUE)

  # Remove download folders at the end unless the user wants to keep them.
  if (!isTRUE(keep_downloads)) {
    on.exit({
      for (d in c("jabot_download", "reflora_download")) {
        if (dir.exists(d)) {
          unlink(d, recursive = TRUE, force = TRUE)
        }
      }
    }, add = TRUE)
  }

  already_loaded <- character(0)

  # Download and load one resource only once.
  load_resource_once <- function(rowi, dir_name) {
    rid_key <- paste(rowi$herbarium, rowi$source, rowi$resource_id, sep = "::")

    if (rid_key %in% already_loaded) {
      return(TRUE)
    }

    folder <- .download_dwca_one(
      info_row = rowi,
      dir = dir_name,
      verbose = verbose,
      force_refresh = force_refresh
    )

    if (is.null(folder)) {
      return(FALSE)
    }

    occ_path <- file.path(folder, "occurrence.txt")

    ok_load <- .duckdb_load_occurrence(
      con = con,
      occurrence_path = occ_path,
      herbarium = rowi$herbarium,
      source = rowi$source,
      resource_id = rowi$resource_id,
      verbose = verbose
    )

    if (isTRUE(ok_load)) {
      already_loaded <<- c(already_loaded, rid_key)
      return(TRUE)
    }

    FALSE
  }

  # Resolve links from an index table already stored in DuckDB.
  resolve_links_from_index <- function(index_df, rows_to_check, source_label) {
    out_local <- rep(NA_character_, nrow(fp_df))
    resolved_local <- rep(FALSE, nrow(fp_df))

    if (!nrow(index_df) || !length(rows_to_check)) {
      return(list(links = out_local, resolved = resolved_local))
    }

    for (i in rows_to_check) {
      pk <- parsed$primary_key[i]
      fk <- parsed$fallback_key[i]

      hit <- index_df[0, , drop = FALSE]

      # Try primary key first.
      if (!is.na(pk) && nzchar(pk)) {
        hit <- index_df[
          index_df$key == pk & index_df$key_type == "primary",
          ,
          drop = FALSE
        ]
      }

      # Fall back to surname + number if needed.
      if (!nrow(hit) && !is.na(fk) && nzchar(fk)) {
        hit <- index_df[
          index_df$key == fk & index_df$key_type == "fallback",
          ,
          drop = FALSE
        ]
      }

      if (!nrow(hit)) {
        next
      }

      if (identical(source_label, "jabot")) {
        # JABOT URLs require catalogNumber and herbarium.
        hit_valid <- hit[
          !is.na(hit$catalogNumber) & nzchar(trimws(hit$catalogNumber)) &
            !is.na(hit$herbarium) & nzchar(trimws(hit$herbarium)),
          ,
          drop = FALSE
        ]

        if (!nrow(hit_valid)) {
          next
        }

        urls <- unique(mapply(
          .make_jabot_url,
          hit_valid$catalogNumber,
          hit_valid$herbarium,
          USE.NAMES = FALSE
        ))

        urls <- urls[!is.na(urls) & nzchar(urls)]

        if (length(urls)) {
          out_local[i] <- paste0(
            "<br/><b>Herbarium image:</b> ",
            sprintf("<a href='%s' target='_blank' title='JABOT'>JABOT</a>", urls[1])
          )
          resolved_local[i] <- TRUE
        }

      } else if (identical(source_label, "reflora")) {
        # REFLORA URLs require occurrenceID.
        hit_valid <- hit[
          !is.na(hit$occurrenceID) & nzchar(trimws(hit$occurrenceID)),
          ,
          drop = FALSE
        ]

        if (!nrow(hit_valid)) {
          next
        }

        urls <- unique(vapply(
          hit_valid$occurrenceID,
          .make_reflora_url,
          character(1)
        ))

        urls <- urls[!is.na(urls) & nzchar(urls)]

        if (length(urls)) {
          out_local[i] <- paste0(
            "<br/><b>Herbarium image:</b> ",
            sprintf("<a href='%s' target='_blank' title='REFLORA'>REFLORA</a>", urls[1])
          )
          resolved_local[i] <- TRUE
        }
      }
    }

    list(links = out_local, resolved = resolved_local)
  }

  out <- rep(NA_character_, nrow(fp_df))
  unresolved <- rep(TRUE, nrow(fp_df))

  herb_sql <- paste(sprintf("'%s'", gsub("'", "''", herbaria)), collapse = ", ")

  # ============================================================================
  # Step 1: Load and match JABOT only
  # ============================================================================
  for (h in herbaria) {
    jabot_res <- .get_ipt_info(
      herbarium = h,
      ipt = "jabot",
      resource_map = resource_map
    )

    if (!nrow(jabot_res)) {
      if (isTRUE(verbose)) {
        message("No JABOT resource found for ", h)
      }
      next
    }

    jabot_res$source <- "jabot"

    if (isTRUE(verbose)) {
      message("JABOT resources found for ", h, ": ", nrow(jabot_res))
      print(jabot_res[, c("ipt", "herbarium", "resource_id", "source")], row.names = FALSE)
    }

    for (i in seq_len(nrow(jabot_res))) {
      rowi <- jabot_res[i, , drop = FALSE]

      ok_load <- load_resource_once(rowi, "jabot_download")
      if (!ok_load) {
        next
      }

      .duckdb_match_resource(
        con = con,
        herbarium = rowi$herbarium,
        source = rowi$source,
        resource_id = rowi$resource_id,
        need_primary = need_primary,
        need_fallback = need_fallback,
        verbose = verbose
      )
    }
  }

  # Query JABOT matches from the local index.
  idx_jabot <- DBI::dbGetQuery(
    con,
    paste0(
      "SELECT key, key_type, catalogNumber, occurrenceID, herbarium, source ",
      "FROM herbaria_index ",
      "WHERE herbarium IN (", herb_sql, ") AND source = 'jabot';"
    )
  )

  jabot_resolved <- resolve_links_from_index(
    index_df = idx_jabot,
    rows_to_check = which(unresolved),
    source_label = "jabot"
  )

  if (any(jabot_resolved$resolved)) {
    out[which(jabot_resolved$resolved)] <- jabot_resolved$links[which(jabot_resolved$resolved)]
    unresolved[which(jabot_resolved$resolved)] <- FALSE
  }

  # ============================================================================
  # Step 2: Only if needed, load and match REFLORA for unresolved records
  # ============================================================================
  if (any(unresolved)) {
    need_primary_rf <- unique(stats::na.omit(parsed$primary_key[unresolved]))
    need_fallback_rf <- unique(stats::na.omit(parsed$fallback_key[unresolved]))

    if (length(need_primary_rf) || length(need_fallback_rf)) {
      for (h in herbaria) {
        reflora_res <- .get_ipt_info(
          herbarium = h,
          ipt = "reflora",
          resource_map = resource_map
        )

        if (!nrow(reflora_res)) {
          if (isTRUE(verbose)) {
            message("No REFLORA resource found for ", h)
          }
          next
        }

        reflora_res$source <- "reflora"

        if (isTRUE(verbose)) {
          message("REFLORA resources found for ", h, ": ", nrow(reflora_res))
          print(reflora_res[, c("ipt", "herbarium", "resource_id", "source")], row.names = FALSE)
        }

        for (i in seq_len(nrow(reflora_res))) {
          rowi <- reflora_res[i, , drop = FALSE]

          ok_load <- load_resource_once(rowi, "reflora_download")
          if (!ok_load) {
            next
          }

          .duckdb_match_resource(
            con = con,
            herbarium = rowi$herbarium,
            source = rowi$source,
            resource_id = rowi$resource_id,
            need_primary = need_primary_rf,
            need_fallback = need_fallback_rf,
            verbose = verbose
          )
        }
      }

      # Query REFLORA matches from the local index.
      idx_reflora <- DBI::dbGetQuery(
        con,
        paste0(
          "SELECT key, key_type, catalogNumber, occurrenceID, herbarium, source ",
          "FROM herbaria_index ",
          "WHERE herbarium IN (", herb_sql, ") AND source = 'reflora';"
        )
      )

      reflora_resolved <- resolve_links_from_index(
        index_df = idx_reflora,
        rows_to_check = which(unresolved),
        source_label = "reflora"
      )

      if (any(reflora_resolved$resolved)) {
        out[which(reflora_resolved$resolved)] <- reflora_resolved$links[which(reflora_resolved$resolved)]
        unresolved[which(reflora_resolved$resolved)] <- FALSE
      }
    }
  }

  out
}
