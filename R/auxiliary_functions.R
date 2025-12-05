# Auxiliary functions
# Author: Giulia Ottino & Domingos Cardoso
# 05/Dec/2025

# Convert input data frames into readable sheets ####
.fp_query_to_field_sheet_df <- function(
    path,
    sheet  = "Data",
    locale = readr::locale(decimal_mark = ".", grouping_mark = ",")
) {
  dest_cols <- c(
    "New Tag No","New Stem Grouping","T1","T2","X","Y","Family",
    "Original determination","Morphospecies","D","POM","ExtraD","ExtraPOM",
    "Flag1","Flag2","Flag3","LI","CI","CF","CD1","nrdups","Height",
    "Voucher","Silica","Collected","Census Notes","CAP","Basal Area"
  )
    read_in <- function(f, sh = NULL) {
    if (grepl("\\.xlsx?$", f, ignore.case = TRUE)) {
      sheets <- readxl::excel_sheets(f)

      pref <- c(sh, "Data", "Plot Dump", sheets)
      pref <- unique(pref)

      df <- NULL
      used_sheet <- NULL

      for (s in pref) {
        if (!s %in% sheets) next
        temp <- suppressMessages(
          readxl::read_excel(f, sheet = s, .name_repair = "minimal")
        )
        # critério simples para "aba com dados"
        nrows_check <- min(10, nrow(temp))
        if (ncol(temp) >= 3 && nrows_check > 0 &&
            sum(!is.na(temp[1:nrows_check, ])) > 2) {
          df <- temp
          used_sheet <- s
          break
        }
      }

      if (is.null(df)) {
        stop("No data found in any worksheet.", call. = FALSE)
      }
      attr(df, "sheet_name") <- used_sheet
      return(df)

    } else {
      df <- readr::read_delim(
        f,
        delim          = ",",
        locale         = locale,
        show_col_types = FALSE,
        .name_repair   = "minimal"
      )
      attr(df, "sheet_name") <- NA_character_
      return(df)
    }
  }
  in_dat <- suppressWarnings(read_in(path, sheet))
  sheet_used <- attr(in_dat, "sheet_name")
  colnames(in_dat) <- stringr::str_trim(colnames(in_dat))
  is_plot_dump_header_row <-
    any(grepl("^Unnamed", names(in_dat))) &&
    any(as.character(unlist(in_dat[1, ])) == "Tree ID") &&
    any(as.character(unlist(in_dat[1, ])) == "Voucher Code")

  if (is_plot_dump_header_row) {
    hdr <- as.character(unlist(in_dat[1, ]))
    hdr <- stringr::str_trim(hdr)
    names(in_dat) <- hdr
    in_dat <- in_dat[-1, , drop = FALSE]
  }
  colnames(in_dat) <- stringr::str_trim(colnames(in_dat))
    is_field_sheet <- all(dest_cols %in% names(in_dat))
    is_plot_dump <- all(
    c("Tag No", "T1", "T2", "X", "Y",
      "Family",
      "Voucher Code",
      "Voucher Collected",
      "Recommended Voucher Species") %in% names(in_dat)
  )
  is_voucher_query <- all(
    c("Voucher Code", "TagNumber", "Plot Code") %in% names(in_dat)
  )

  if (is_field_sheet) {

    out <- in_dat

  } else if (is_plot_dump) {
    get_col <- function(df, candidates) {
      candidates <- as.character(candidates)
      for (cl in candidates) {
        if (cl %in% names(df)) {
          return(df[[cl]])
        }
      }
      return(rep(NA, nrow(df)))
    }
       out <- tibble::tibble(
      `New Tag No`            = get_col(in_dat, c("New Tag No", "Tag No", "Pv. Tag No", "Tree ID")),
      `New Stem Grouping`     = get_col(in_dat, c("New Stem Grouping", "Stem Group ID")),
      T1                      = get_col(in_dat, "T1"),
      T2                      = get_col(in_dat, "T2"),
      X                       = get_col(in_dat, "X"),
      Y                       = get_col(in_dat, "Y"),
      Family                  = get_col(in_dat, c("Recommended Voucher Family",
                                                  "Recommended Family",
                                                  "Family")),
      `Original determination`= get_col(in_dat, c("Recommended Voucher Species",
                                                  "Recommended Species",
                                                  "Species")),
      Morphospecies           = get_col(in_dat, c("Morphospecies", "MorphoSpecies", "Morpho")),
      D                       = get_col(in_dat, c("D", "DBH", "Dbh")),
      POM                     = get_col(in_dat, "POM"),
      ExtraD                  = get_col(in_dat, "ExtraD"),
      ExtraPOM                = get_col(in_dat, "ExtraPOM"),
      Flag1                   = get_col(in_dat, "Flag1"),
      Flag2                   = get_col(in_dat, "Flag2"),
      Flag3                   = get_col(in_dat, "Flag3"),
      LI                      = get_col(in_dat, "LI"),
      CI                      = get_col(in_dat, "CI"),
      CF                      = get_col(in_dat, "CF"),
      CD1                     = get_col(in_dat, "CD1"),
      nrdups                  = get_col(in_dat, c("nrdups", "Stem Count")),
      Height                  = get_col(in_dat, c("Height", "Ht")),
      Voucher                 = get_col(in_dat, c("Voucher", "Voucher Code")),
      Silica                  = get_col(in_dat, "Silica"),
      Collected               = get_col(in_dat, c("Collected", "Voucher Collected")),
      `Census Notes`          = get_col(in_dat, c("Census Notes", "Tree Notes", "Determination Comments")),
      CAP                     = get_col(in_dat, "CAP"),
      `Basal Area`            = get_col(in_dat, c("Basal Area", "BA"))
    )

  } else if (is_voucher_query) {
    # Formato Data(7).xlsx: uma linha por voucher, sem coordenadas.
    get_col <- function(df, candidates) {
      candidates <- as.character(candidates)
      for (cl in candidates) {
        if (cl %in% names(df)) {
          return(df[[cl]])
        }
      }
      return(rep(NA, nrow(df)))
    }
    tag <- get_col(in_dat, "TagNumber")
    fam <- get_col(in_dat, c("Recommended Family", "Family"))
    sp  <- get_col(in_dat, c("Recommended Species", "Species"))
    vch <- get_col(in_dat, "Voucher Code")

    collected_vec <- ifelse(!is.na(vch) & vch != "", "Yes", NA_character_)

    out <- tibble::tibble(
      `New Tag No`            = tag,
      `New Stem Grouping`     = NA,
      T1                      = NA,
      T2                      = NA,
      X                       = NA,
      Y                       = NA,
      Family                  = fam,
      `Original determination`= sp,
      Morphospecies           = NA,
      D                       = NA,
      POM                     = NA,
      ExtraD                  = NA,
      ExtraPOM                = NA,
      Flag1                   = NA,
      Flag2                   = NA,
      Flag3                   = NA,
      LI                      = NA,
      CI                      = NA,
      CF                      = NA,
      CD1                     = NA,
      nrdups                  = NA,
      Height                  = NA,
      Voucher                 = vch,
      Silica                  = NA,
      Collected               = collected_vec,
      `Census Notes`          = NA,
      CAP                     = NA,
      `Basal Area`            = NA
    )

  } else {
    missing <- setdiff(dest_cols, names(in_dat))
    stop(
      paste0(
        "Essential columns for a ForestPlots field sheet were not found. ",
        "Please check the column names in your worksheet. ",
        "Missing columns: ",
        paste(missing, collapse = ", ")
      ),
      call. = FALSE
    )
  }
    for (col in dest_cols) {
    if (!col %in% names(out)) out[[col]] <- NA
  }
  out <- out[, dest_cols]
    first_val <- function(df, cols) {
    for (cl in cols) {
      if (cl %in% names(df)) {
        v <- df[[cl]][1]
        return(v)
      }
    }
    NA_character_
  }

  meta <- rep("", length(dest_cols))
  meta[1]  <- paste0("Plotcode: ", first_val(in_dat, c("Plot Code", "PlotCode")))
  meta[2]  <- paste0("Plot Name: ", first_val(in_dat, c("Plot Name", "PlotName")))

  date_val <- first_val(in_dat, c("Date", "Census Date"))
  if (!is.null(date_val) && !all(is.na(date_val)) && inherits(date_val, "Date")) {
    date_val <- format(date_val, "%d/%m/%Y")
  }
  meta[10] <- paste0("Date: ", ifelse(is.na(date_val), "", date_val))

  team_val <- first_val(in_dat, c("PI", "Team"))
  meta[15] <- paste0("Team: ", ifelse(is.na(team_val), "", team_val))

  # Junta: linha de metadados + linha de cabeçalho + dados
  field_temp <- tibble::as_tibble(rbind(meta, dest_cols, out))
  names(field_temp) <- NULL

  return(field_temp)
}


.monitora_to_field_sheet_df <- function(path, sheet = 1, station_name = NULL) {
  # --------- helpers ----------
  .norm <- function(x){ x <- iconv(x, to="ASCII//TRANSLIT"); x <- tolower(x); gsub("[^a-z0-9]+","",x) }

  .find_first_col <- function(df, aliases){
    if (ncol(df) == 0) return(NA_character_)
    cn_raw <- names(df); cn_norm <- .norm(cn_raw); ali_norm <- .norm(aliases)
    hit_idx <- match(ali_norm, cn_norm, nomatch = 0L)
    if (any(hit_idx > 0L)) return(cn_raw[ hit_idx[which(hit_idx > 0L)[1]] ])
    NA_character_
  }

  .to_numeric <- function(x){
    if (is.factor(x)) x <- as.character(x)
    x <- gsub("\\s+","",x)
    x <- sub("^([^;/|]+)[;/|].*$","\\1",x, perl=TRUE)
    x <- gsub("\\.(?=\\d{3}(\\D|$))","",x, perl=TRUE)
    x <- gsub(",",".",x, fixed=TRUE)
    pat <- "[-+]?(?:\\d+\\.?\\d*|\\d*\\.?\\d+)"
    has_num <- grepl(pat,x, perl=TRUE)
    y <- rep(NA_character_, length(x))
    y[has_num] <- regmatches(x[has_num], regexpr(pat, x[has_num], perl=TRUE))
    suppressWarnings(as.numeric(y))
  }

  .to_year <- function(x){
    x <- as.character(x); x <- gsub("\\s+","",x); x <- gsub(",","",x, fixed=TRUE)
    m <- regexpr("(?:19|20)\\d{2}", x, perl=TRUE)
    y <- rep(NA_integer_, length(x)); hit <- m > 0
    y[hit] <- suppressWarnings(as.integer(substr(x[hit], m[hit], m[hit] + attr(m,"match.length")[hit]-1)))
    y
  }

  .station_norm_name <- function(v){ v <- trimws(as.character(v)); v <- iconv(v, to="ASCII//TRANSLIT"); tolower(v) }

  .station_norm_num  <- function(v){ v <- trimws(as.character(v)); v <- gsub("\\s+","",v); v <- sub("^0+([0-9])","\\1",v); v[v==""] <- NA_character_; v }

  clean_spaces <- function(x){
    x <- as.character(x)
    x <- gsub(intToUtf8(0x00A0), " ", x, fixed=TRUE)
    x <- gsub(intToUtf8(0x2007), " ", x, fixed=TRUE)
    x <- gsub(intToUtf8(0x202F), " ", x, fixed=TRUE)
    trimws(x)
  }
  resolve_col <- function(df, col){
    if (is.na(col)) return(NA_character_)
    nm <- names(df); hit <- which(tolower(trimws(nm)) == tolower(trimws(col)))
    if (length(hit)==1) nm[hit] else col
  }
  first_nonempty <- function(df, col, lixo=character(0)){
    if (is.na(col) || !(col %in% names(df)) || nrow(df)==0) return("")
    v <- clean_spaces(df[[col]])
    ok <- !is.na(v) & nzchar(v) & !(tolower(v) %in% tolower(lixo))
    v <- v[ok]; if (length(v)>0) v[1] else ""
  }

  # ---------- reading ----------
  df <- suppressMessages(readxl::read_excel(path, sheet=sheet, .name_repair="minimal", col_types="text"))
  if (!is.data.frame(df) || !nrow(df)) stop("Empty Monitora worksheet.")

  # ---------- aliases ----------
  col_subunidade <- .find_first_col(df, c("subunidade","subunidades","orientacao","unidade","ul"))
  col_nparcela   <- .find_first_col(df, c("n_parcela","nparcela","parcela","t2"))
  col_tag        <- .find_first_col(df, c("n_arvore","narvore","tag","newtag","numarvore"))
  col_coletores  <- .find_first_col(df, c("nome_coletores","coletores","equipe","team"))
  col_uc         <- .find_first_col(df, c("nomeuc","uc","unidadeconservacao","und.conservacao"))
  col_estacao    <- .find_first_col(df, c("nome_estacao","nome_estação","estacao","estação","nomeestacao"))
  col_estacao_n  <- .find_first_col(df, c("n_estacao","nestacao","num_estacao","numero_estacao","nºestacao","n°estacao"))
  col_familia    <- .find_first_col(df, c("familia","family"))
  col_genero     <- .find_first_col(df, c("genero","gênero","genero_sp","genus"))
  col_especie    <- .find_first_col(df, c("especie","espécie","species","sp"))
  col_coletado   <- .find_first_col(df, c("individuo_coletado","coletado","collected"))
  col_voucher_c  <- .find_first_col(df, c("voucher/coletor","vouchercoletor","coletor"))
  col_voucher_n  <- .find_first_col(df, c("voucher/numero","voucher/número","vouchernumero","voucher_n","voucher"))
  col_x          <- .find_first_col(df, c("x(m)","x","coordx","xm","x (m)"))
  col_y          <- .find_first_col(df, c("y(m)","y","coordy","ym","y (m)"))
  col_cap        <- .find_first_col(df, c("cap_tot","captot","cap","circunferencia"))
  col_nomecomum  <- .find_first_col(df, c("nomecomum","nome_comum","popular","morphospecies"))
  col_ano        <- .find_first_col(df, c("ano","anoo","censo","year"))
  col_dead       <- .find_first_col(df, c("arvore_morta","árvore_morta","arvore morta","arvore.morta","morta","dead"))

  # ---------- subset to most recent census ----------
  if (!is.na(col_ano) && col_ano %in% names(df)){
    ano_vec <- .to_year(df[[col_ano]])
    ano_max <- suppressWarnings(max(ano_vec, na.rm=TRUE))
    df_use  <- if (is.finite(ano_max)) df[ano_vec == ano_max, , drop=FALSE] else df
  } else {
    ano_vec <- rep(NA_integer_, nrow(df)); ano_max <- NA_integer_; df_use <- df
  }
  df_all <- df

  # ---------- normalize/filter by station ----------
  if (!is.na(col_estacao)   && col_estacao   %in% names(df_use)) df_use$.__st_name <- .station_norm_name(df_use[[col_estacao]])   else df_use$.__st_name <- NA_character_
  if (!is.na(col_estacao_n) && col_estacao_n %in% names(df_use)) df_use$.__st_num  <- .station_norm_num (df_use[[col_estacao_n]]) else df_use$.__st_num  <- NA_character_
  if (!is.na(col_estacao)   && col_estacao   %in% names(df_all)) df_all$.__st_name <- .station_norm_name(df_all[[col_estacao]])   else df_all$.__st_name <- NA_character_
  if (!is.na(col_estacao_n) && col_estacao_n %in% names(df_all)) df_all$.__st_num  <- .station_norm_num (df_all[[col_estacao_n]]) else df_all$.__st_num  <- NA_character_

  if (!is.na(col_estacao) || !is.na(col_estacao_n)){
    units_df   <- data.frame(name=df_use$.__st_name, num=df_use$.__st_num, stringsAsFactors=FALSE)
    units_df   <- units_df[ !(is.na(units_df$name) & is.na(units_df$num)), , drop=FALSE ]
    unit_keys  <- unique(paste0(ifelse(is.na(units_df$name),"",units_df$name),"||",ifelse(is.na(units_df$num),"",units_df$num)))
    multi_units <- length(unit_keys) > 1

    if (multi_units && (is.null(station_name) || !length(station_name))){
      names_set <- sort(unique(units_df$name[!is.na(units_df$name) & nzchar(units_df$name)]))
      nums_set  <- sort(unique(units_df$num [!is.na(units_df$num)  & nzchar(units_df$num) ]))
      stop(paste0(
        "This MONITORA sheet contains more than one sampling unit in the most recent census (",
        ifelse(is.finite(ano_max), ano_max, "unknown year"),
        "). Please provide 'station_name' (Nome_estacao or N_estacao). ",
        if (length(names_set)) paste0("Available names: ", paste(names_set, collapse=", "), ". ") else "",
        if (length(nums_set))  paste0("Available numbers: ", paste(nums_set,  collapse=", "), ".") else ""
      ), call. = FALSE)
    }

    if (!is.null(station_name) && length(station_name)){
      want_raw  <- trimws(as.character(station_name))
      want_name <- .station_norm_name(want_raw)
      want_num  <- .station_norm_num (want_raw)

      keep_recent <- (!is.na(df_use$.__st_name) & df_use$.__st_name %in% want_name) |
        (!is.na(df_use$.__st_num ) & df_use$.__st_num  %in% want_num)
      if (!any(keep_recent)) stop("Requested 'station_name' not found in the most recent census.", call. = FALSE)
      df_use <- df_use[keep_recent, , drop=FALSE]

      keep_all <- (!is.na(df_all$.__st_name) & df_all$.__st_name %in% want_name) |
        (!is.na(df_all$.__st_num ) & df_all$.__st_num  %in% want_num)
      if (any(keep_all)) df_all <- df_all[keep_all, , drop=FALSE]
    }
  }

  # ---------- fallback X/Y from earlier censuses ----------
  if ((!is.na(col_x) || !is.na(col_y)) && !is.na(col_tag) && !is.na(col_ano) &&
      all(c(col_x, col_y, col_tag, col_ano) %in% names(df))){
    need_xy <- rep(FALSE, nrow(df_use))
    if (!is.na(col_x)) need_xy <- need_xy | is.na(df_use[[col_x]])
    if (!is.na(col_y)) need_xy <- need_xy | is.na(df_use[[col_y]])

    if (any(need_xy)){
      ano_all <- .to_year(df[[col_ano]])
      older <- df[ano_all != suppressWarnings(max(ano_all, na.rm=TRUE)) & !is.na(ano_all), , drop=FALSE]
      if (nrow(older)){
        for (i in which(need_xy)){
          key <- as.character(df_use[[col_tag]][i]); if (!nzchar(key)) next
          hit <- older[ trimws(as.character(older[[col_tag]])) == trimws(key), , drop=FALSE ]
          if (nrow(hit)){
            if (!is.na(col_x) && is.na(df_use[[col_x]][i])) df_use[[col_x]][i] <- hit[[col_x]][1]
            if (!is.na(col_y) && is.na(df_use[[col_y]][i])) df_use[[col_y]][i] <- hit[[col_y]][1]
          }
        }
      }
    }
  }

  # ---------- metadata ----------
  col_uc      <- resolve_col(df_use, col_uc)
  col_estacao <- resolve_col(df_use, col_estacao)
  lixo_uc      <- c("","na","n/a","-","--",".","s/n","sn")
  lixo_estacao <- c("","na","n/a","-","--",".")
  uc_val    <- first_nonempty(df_use, col_uc, lixo_uc)
  plot_code <- first_nonempty(df_use, col_estacao, lixo_estacao)
  if (!nzchar(plot_code) && !is.na(col_estacao_n) && col_estacao_n %in% names(df_use)){
    plot_code <- first_nonempty(df_use, col_estacao_n, lixo_estacao)
  }
  plot_name <- uc_val

  team_val <- ""
  if (!is.na(col_coletores) && col_coletores %in% names(df_use)){
    coletores <- clean_spaces(df_use[[col_coletores]])
    coletores <- unique(coletores[!is.na(coletores) & nzchar(coletores)])
    team_val  <- if (length(coletores)>0) paste(coletores, collapse="; ") else ""
  }
  try(assign("plot_name", plot_name, envir=parent.frame()), silent=TRUE)
  try(assign("plot_code", plot_code, envir=parent.frame()), silent=TRUE)
  try(assign("team",      team_val,  envir=parent.frame()), silent=TRUE)

  # ---------- multi-census summary ----------
  census_years <- if (!is.na(col_ano) && col_ano %in% names(df_all)) sort(unique(.to_year(df_all[[col_ano]]))) else integer(0)
  dead_since_first <- NA_integer_
  if (!is.na(col_dead) && col_dead %in% names(df_all)){
    v <- tolower(trimws(as.character(df_all[[col_dead]])))
    dead_since_first <- sum(v %in% c("sim","s","yes","y","true","1"), na.rm=TRUE)
  }
  recruits_since_first <- NA_integer_
  if (!is.na(col_tag) && col_tag %in% names(df_all) && length(census_years) >= 2 && !is.na(col_ano)){
    y0 <- min(census_years, na.rm=TRUE); y1 <- max(census_years, na.rm=TRUE)
    year_all  <- .to_year(df_all[[col_ano]])
    tag_first <- unique(trimws(as.character(df_all[[col_tag]][ year_all==y0 ])))
    tag_last  <- unique(trimws(as.character(df_all[[col_tag]][ year_all==y1 ])))
    tag_first <- tag_first[!is.na(tag_first) & nzchar(tag_first)]
    tag_last  <- tag_last [!is.na(tag_last)  & nzchar(tag_last)]
    recruits_since_first <- length(setdiff(tag_last, tag_first))
  }
  try(assign("monitora_census_years", census_years, envir=parent.frame()), silent=TRUE)
  try(assign("monitora_census_n",     length(census_years), envir=parent.frame()), silent=TRUE)
  try(assign("monitora_dead_since_first",     dead_since_first, envir=parent.frame()), silent=TRUE)
  try(assign("monitora_recruits_since_first", recruits_since_first, envir=parent.frame()), silent=TRUE)

  # ---------- subunit mapping ----------
  map_sub <- function(v){
    v <- toupper(trimws(as.character(v)))
    dplyr::case_when(
      v %in% c("N","NORTE")              ~ 1,
      v %in% c("S","SUL","SOUTH")        ~ 2,
      v %in% c("L","LESTE","E","EAST")   ~ 3,
      v %in% c("O","OESTE","W","WEST")   ~ 4,
      TRUE ~ NA_real_
    )
  }

  # ---------- destination (28 columns) ----------
  dest_cols <- c(
    "New Tag No","New Stem Grouping","T1","T2","X","Y","Family",
    "Original determination","Morphospecies","D","POM","ExtraD","ExtraPOM",
    "Flag1","Flag2","Flag3","LI","CI","CF","CD1","nrdups","Height",
    "Voucher","Silica","Collected","Census Notes","CAP","Basal Area"
  )
  out <- as.data.frame(matrix(NA_character_, nrow=nrow(df_use), ncol=length(dest_cols)))
  names(out) <- dest_cols

  if (!is.na(col_tag)        && col_tag        %in% names(df_use)) out[["New Tag No"]] <- df_use[[col_tag]]
  if (!is.na(col_subunidade) && col_subunidade %in% names(df_use)) out[["T1"]]         <- map_sub(df_use[[col_subunidade]])
  if (!is.na(col_nparcela)   && col_nparcela   %in% names(df_use)) out[["T2"]]         <- df_use[[col_nparcela]]
  if (!is.na(col_x)          && col_x          %in% names(df_use)) out[["X"]]          <- .to_numeric(df_use[[col_x]])
  if (!is.na(col_y)          && col_y          %in% names(df_use)) out[["Y"]]          <- .to_numeric(df_use[[col_y]])
  if (!is.na(col_familia)    && col_familia    %in% names(df_use)) out[["Family"]]     <- df_use[[col_familia]]

  gen <- if (!is.na(col_genero) && col_genero %in% names(df_use)) as.character(df_use[[col_genero]]) else NA
  esp <- if (!is.na(col_especie) && col_especie %in% names(df_use)) as.character(df_use[[col_especie]]) else NA
  od  <- ifelse(is.na(esp) | !nzchar(esp),
                ifelse(is.na(gen) | !nzchar(gen), NA, paste0(gen," indet")),
                paste0(gen," ",esp))
  out[["Original determination"]] <- od

  if (!is.na(col_nomecomum) && col_nomecomum %in% names(df_use)) out[["Morphospecies"]] <- as.character(df_use[[col_nomecomum]])

  if (!is.na(col_cap) && col_cap %in% names(df_use)){
    cap_cm <- .to_numeric(df_use[[col_cap]])
    d_mm   <- (cap_cm / pi) * 10
    out[["D"]]   <- round(d_mm, 2)
    out[["CAP"]] <- cap_cm
  }

  if (!is.na(col_coletado) && col_coletado %in% names(df_use)){
    cc_raw <- as.character(df_use[[col_coletado]])
    cc <- tolower(iconv(cc_raw, to="ASCII//TRANSLIT")); cc <- trimws(cc)
    yes_vals <- c("sim","s","yes","y","true","1","coletado","coletada")
    out[["Collected"]] <- ifelse(cc %in% yes_vals, "yes", NA_character_)
  }

  # ---------- VOUCHER: use detected columns and fill from past censuses ----------
  vc  <- if (!is.na(col_voucher_c) && col_voucher_c %in% names(df_use)) as.character(df_use[[col_voucher_c]]) else NA
  vnr <- if (!is.na(col_voucher_n) && col_voucher_n %in% names(df_use)) as.character(df_use[[col_voucher_n]]) else NA
  vc  <- clean_spaces(vc);  vnr <- clean_spaces(vnr)
  cur_voucher <- ifelse((is.na(vc) | vc=="") & (is.na(vnr) | vnr==""),
                        NA_character_,
                        trimws(paste(vc, vnr)))

  # voucher map by TAG using the MOST RECENT non-empty entry in df_all
  if (!is.na(col_tag)){
    year_all <- if (!is.na(col_ano) && col_ano %in% names(df_all)) .to_year(df_all[[col_ano]]) else rep(NA_integer_, nrow(df_all))
    tag_all  <- trimws(as.character(df_all[[col_tag]]))
    vc_all   <- if (!is.na(col_voucher_c) && col_voucher_c %in% names(df_all)) clean_spaces(df_all[[col_voucher_c]]) else NA
    vnr_all  <- if (!is.na(col_voucher_n) && col_voucher_n %in% names(df_all)) clean_spaces(df_all[[col_voucher_n]]) else NA
    comb_all <- ifelse((is.na(vc_all) | vc_all=="") & (is.na(vnr_all) | vnr_all==""),
                       NA_character_,
                       trimws(paste(vc_all, vnr_all)))

    ord <- order(year_all, na.last = TRUE, decreasing = TRUE)
    tag_all  <- tag_all [ord]
    comb_all <- comb_all[ord]

    voucher_by_tag <- tapply(comb_all, tag_all, function(v) {
      v <- v[!is.na(v) & nzchar(v)]
      if (length(v)) v[1] else NA_character_
    })

    tag_use <- trimws(as.character(df_use[[col_tag]]))
    fill_from_past <- voucher_by_tag[tag_use]
    fill_from_past[is.na(fill_from_past)] <- NA_character_

    # if current census has NA, use voucher from past censuses
    cur_voucher <- ifelse(is.na(cur_voucher) | cur_voucher=="",
                          as.character(fill_from_past),
                          cur_voucher)
  }

  out[["Voucher"]] <- cur_voucher

  # ---------- minimal taxonomic cleaning ----------
  to_clean_taxa <- function(x, is_species=FALSE){
    v <- trimws(as.character(x))
    bad <- tolower(v) %in% c("na","n/a","indeterminada","inteterminada","interminada","","s/n","sn")
    v[bad] <- if (is_species) "indet" else "Indet"
    v
  }
  if ("Family" %in% names(out)) out$Family <- to_clean_taxa(out$Family, FALSE)

  # ---------- enforce column schema ----------
  missing_cols <- setdiff(dest_cols, names(out))
  for (nm in missing_cols) out[[nm]] <- NA
  out <- out[, dest_cols, drop=FALSE]

  # ---------- metadata + header + data ----------
  meta <- rep("", length(dest_cols))
  meta[2]  <- paste0("Plot Name: ", plot_name)
  meta[3]  <- paste0("Plot Code: ", plot_code)
  if (is.finite(ano_max)) meta[10] <- paste0("Date: ", ano_max)
  meta[15] <- paste0("Team: ", team_val)

  field_temp <- tibble::as_tibble(rbind(meta, dest_cols, out))
  names(field_temp) <- NULL
  field_temp
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

.compute_global_coordinates_monitora <- function(fp_df) {
  # 1) Safely convert to numeric
  fp_df$T1 <- suppressWarnings(as.numeric(fp_df$T1))   # 1=N, 2=L, 3=S, 4=O
  fp_df$T2 <- suppressWarnings(as.numeric(fp_df$T2))   # 1..10 (zigzag)
  fp_df$X  <- suppressWarnings(as.numeric(fp_df$X))    # local: transversal
  fp_df$Y  <- suppressWarnings(as.numeric(fp_df$Y))    # local: along

  # 2) Map T1 → arm letter
  base <- data.frame(T1 = c(1,2,3,4), arm = c("N","S","L","O"))
  fp2  <- dplyr::left_join(fp_df, base, by = "T1")

  # 3) Ensure local coordinate ranges for ALL subunits
  #    X: [-10, 10] (transversal); Y: [0, 50] (along)
  fp2 <- fp2 %>%
    dplyr::mutate(
      X_loc = pmax(pmin(X,  10), -10),
      Y_loc = pmax(pmin(Y,  50),   0)
    )

  # 4) Build drawing coordinates (with “along” axis mirroring in S/O)
  fp3 <- fp2 %>%
    dplyr::mutate(
      # local clamping
      X_loc = pmax(pmin(X,  10), -10),
      Y_loc = pmax(pmin(Y,  50),   0),

      # mirror the “along” axis in S and O: 0 (center) ↔ 50 (edge)
      along_m = dplyr::if_else(arm %in% c("S","O"), 50 - Y_loc, Y_loc),

      # X axis in the drawing
      draw_x = dplyr::case_when(
        arm %in% c("N","S") ~ X_loc,             # vertical arms: transversal → X
        arm == "L"          ~ 50   + along_m,    # L to +X
        arm == "O"          ~ -100 + along_m     # O to -X
      ),

      # Y axis in the drawing
      draw_y = dplyr::case_when(
        arm == "N"          ~  50   + along_m,   # N to +Y
        arm == "S"          ~ -100  + along_m,   # S to -Y
        arm %in% c("L","O") ~ X_loc              # horizontal arms: transversal → Y
      ),

      SubplotID      = (as.integer(T1) - 1L) * 10L + as.integer(T2),
      subunit_letter = arm,
      subplot_name   = paste0(arm, T2)
    ) %>%
    dplyr::filter(is.finite(draw_x), is.finite(draw_y))

  fp3
}


# functions to plot_for_balance translations ####
#### English version ####
.create_rmd_content_en <- function(subplot_plots, tf_col, tf_uncol, tf_palm,
                                   plot_name, plot_code, spec_df,
                                   has_agb = FALSE) {

  # 1) YAML HEADER – built dynamically
  yaml_head <- c(
    "---",
    "output:",
    "  pdf_document:",
    "    toc: true",
    "    toc_depth: 2",
    "    number_sections: true",
    "fontsize: 12pt",
    "params:",
    "  input_type: \"forestplots\"",
    "  metadata: NULL",
    "  main_plot: NULL",
    "  subplots_list: NULL",
    "  subplot_size: 70",
    "  stats: NULL",
    "  tag_list: NULL",
    "  tag_to_subplot: NULL"
  )

  if (has_agb)       yaml_head <- c(yaml_head, "  agb: NULL")
  if (any(tf_col))   yaml_head <- c(yaml_head, "  collected_plot: NULL")
  if (any(tf_uncol)) yaml_head <- c(yaml_head, "  uncollected_plot: NULL")
  if (any(tf_palm))  yaml_head <- c(yaml_head, "  uncollected_palm_plot: NULL")

  yaml_head <- c(yaml_head, "---")

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
    "  x <- gsub(\"[^A-Za-z0-9]\", \"\", x)",        # remove punctuation
    "  tolower(x)",
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
    "cat('\\\\Huge\\\\textbf{Full Plot Report} \\\\\\\\')",
    "cat('\\\\vspace{0.5em}')",
    "cat(paste0('\\\\normalsize ', params$metadata$plot_name, ' | ', params$metadata$plot_code, ' \\\\\\\\'))",
    "cat('\\\\end{center}')",
    "```",
    "",
    # -- LOGOS --
    "```{r logo, echo=FALSE, results='asis'}",
    "norm_path <- function(p) {",
    "  p <- as.character(p); p[is.na(p)] <- \"\"",
    "  tryCatch(normalizePath(p, winslash = '/', mustWork = FALSE), error = function(e) p)",
    "}",
    "",
    "forplotr_logo_path    <- norm_path(system.file('figures', 'forplotR_hex_sticker.png', package = 'forplotR'))",
    "forestplots_logo_path <- norm_path(system.file('figures', 'forestplotsnet_logo.png',   package = 'forplotR'))",
    "",
    "raw_it <- if (!is.null(params$input_type)) params$input_type else if (!is.null(params$metadata$input_type)) params$metadata$input_type else ''",
    "input_type  <- tolower(trimws(as.character(raw_it)))",
    "is_monitora <- (input_type == 'monitora')",
    "",
    "monitora_candidates <- character(0)",
    "if (length(forestplots_logo_path) && nzchar(forestplots_logo_path)) {",
    "  mon_dir <- dirname(forestplots_logo_path[1])",
    "  monitora_candidates <- c(monitora_candidates, file.path(mon_dir, 'monitora_logo.png'))",
    "}",
    "monitora_candidates <- c(",
    "  monitora_candidates,",
    "  'figures/monitora_logo.png',",
    "  system.file('figures', 'monitora_logo.png', package = 'forplotR')",
    ")",
    "monitora_candidates <- norm_path(monitora_candidates)",
    "monitora_candidates <- monitora_candidates[nzchar(monitora_candidates)]",
    "",
    "monitora_logo_path <- ''",
    "if (is_monitora) {",
    "  idx <- which(file.exists(monitora_candidates))",
    "  if (length(idx) > 0) {",
    "    monitora_logo_path <- monitora_candidates[idx[1]]",
    "  } else {",
    "    url <- 'https://www.gov.br/icmbio/pt-br/assuntos/monitoramento/programa-monitora/@@collective.cover.banner/af7c14b7-6562-4da9-8056-dc56a2b69f7d/@@images/f86d904c-f11b-4b41-a283-84790263c284.png'",
    "    tmp <- file.path(tempdir(), 'monitora_logo.png')",
    "    try(utils::download.file(url, tmp, mode = 'wb', quiet = TRUE), silent = TRUE)",
    "    if (file.exists(tmp)) monitora_logo_path <- norm_path(tmp)",
    "  }",
    "}",
    "",
    "partner_logo_path <- if (is_monitora && nzchar(monitora_logo_path)) monitora_logo_path else forestplots_logo_path",
    "partner_href      <- if (is_monitora) 'https://www.gov.br/icmbio/pt-br/assuntos/monitoramento/programa-monitora' else 'https://forestplots.net'",
    "partner_width     <- if (is_monitora) \"0.28\\\\linewidth\" else \"0.40\\\\linewidth\"",
    "",
    "cat('\\\\vspace*{1cm}\n')",
    "cat('\\\\begin{center}\n')",
    "cat(sprintf('\\\\href{https://dboslab.github.io/forplotR-website/}{\\\\includegraphics[width=0.20\\\\linewidth]{%s}}', forplotr_logo_path), '\n')",
    "cat('\\\\end{center}\n')",
    "",
    "cat('\\\\begin{center}\n')",
    "if (is_monitora && !nzchar(partner_logo_path)) {",
    "  cat('MONITORA Program', '\\n')",
    "} else {",
    "  cat(sprintf('\\\\href{%s}{\\\\includegraphics[width=%s,keepaspectratio]{%s}}',",
    "              partner_href, partner_width, partner_logo_path), '\n')",
    "}",
    "cat('\\\\end{center}\n')",
    "cat('\\\\vspace*{2cm}\n')",
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
    "\\vspace{0.6\\baselineskip}",
    "\\hrule",
    "",
    # -- MULTI-CENSUS INFO --
    "```{r meta-census, echo=FALSE, results='asis'}",
    "yrs <- params$metadata$census_years; yrs <- yrs[is.finite(yrs)]",
    "if (length(yrs) > 1) {",
    "  cat(paste0('**Number of census:** ', length(yrs), '\\n\\n'))",
    "  cat(paste0('**Dates of census:** ', paste(yrs, collapse = ' | '), '\\n\\n'))",
    "}",
    "```",
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
    "\\hrule",
    "```{r counts-mc, echo=FALSE, results='asis'}",
    "yrs <- params$metadata$census_years",
    "dead_total <- params$stats$dead_since_first",
    "recr_total <- params$stats$recruits_since_first",
    "if (!is.null(yrs) && length(yrs) > 1) {",
    "  if (!is.null(dead_total) && is.finite(dead_total))",
    "    cat(paste0('- **Dead trees since first census:** ', dead_total, '\\n'))",
    "  if (!is.null(recr_total) && is.finite(recr_total))",
    "    cat(paste0('- **Recruits since first census:** ', recr_total, '\\n'))",
    "}",
    "```"
  )

  if (has_agb) {
    mid_section <- c(
      mid_section,
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

  # NAVIGATION LINKS (FOOTER)
  nav_targets <- c(
    "[Back to Contents](#contents)",
    "[Metadata](#metadata)",
    "[Specimen Counts](#counts)"
  )
  if (has_agb) nav_targets <- c(nav_targets, "[AGB](#agb)")
  nav_targets <- c(nav_targets, "[General Plot](#general-plot)")
  if (any(tf_col))   nav_targets <- c(nav_targets, "[Collected Only](#collected-only)")
  if (any(tf_uncol)) nav_targets <- c(nav_targets, "[Not Collected](#uncollected)")
  if (any(tf_palm))  nav_targets <- c(nav_targets, "[Palms](#uncollected-palm)")
  nav_targets <- c(nav_targets, "[Subplot Index](#subplot-index)", "[Checklist](#checklist)")

  nav_links <- c(
    "\\begingroup\\tiny\\color{gray}",
    paste("«", paste(nav_targets, collapse = " | ")),
    "\\endgroup"
  )

  gencol_section <- c(
    "",
    "## General Plot {#general-plot}",
    "```{r general-plot, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$main_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  col_section <- c(
    "## Collected Only {#collected-only}",
    "```{r collected-only, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$collected_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  uncol_section <- c(
    "## Not Collected {#uncollected}",
    "```{r uncollected,fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$uncollected_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  palm_section <- c(
    "## Not Collected Palms {#uncollected-palm}",
    "```{r uncollected-palm, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$uncollected_palm_plot)",
    "```",
    "",
    "",
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
    "",
    "",
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
      sprintf("### Subplot %d {#subplot-%d .unlisted .unnumbered}", i, i),
      "",
      sprintf("```{r subplot-%d, fig.width=12, fig.height=9}", i),
      sprintf("print(params$subplots_list[[%d]]$plot)", i),
      "```",
      "",
      "\\begingroup\\tiny\\color{gray}",
      "« [Back to Contents](#contents) | [Metadata](#metadata) | [Specimen Counts](#counts) | [General Plot](#general-plot) | [Collected Only](#collected-only) | [Not Collected](#uncollected) | [Palms](#uncollected-palm) | [Subplot Index](#subplot-index) | [Checklist](#checklist)",
      "\\endgroup",
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
    "      stringsAsFactors = FALSE)",
    "",
    "# subplot information",
    "t2s_df <- params$tag_to_subplot",
    "",
    "# input_type normalised",
    "is_monitora <- tolower(trimws(as.character(params$input_type))) == \"monitora\"",
    "",
    "if (is_monitora) {",
    "  tag_df <- merge(tag_df, t2s_df, by.x = \"tag\", by.y = \"New Tag No\", all.x = TRUE, sort = FALSE)",
    "  tag_df$target_idx <- ifelse(!is.na(tag_df$subplot_index), tag_df$subplot_index, tag_df$T1)",
    "  has_fields <- !is.na(tag_df$subunit_letter) & !is.na(tag_df$T2)",
    "  tag_df$subplot_code <- ifelse(has_fields, paste0(tag_df$subunit_letter, tag_df$T2), paste0(\"S\", tag_df$target_idx))",
    "  o <- order(tag_df$target_idx, suppressWarnings(as.numeric(tag_df$tag)))",
    "  tag_df <- tag_df[o, , drop = FALSE]",
    "  tag_links <- paste(sprintf(\"[ %s → %s ](#subplot-%s)\",",
    "                           tag_df$tag, tag_df$subplot_code, tag_df$target_idx),",
    "                     collapse = \" | \")",
    "} else {",
    "  tag_df <- merge(tag_df, t2s_df, by.x = \"tag\", by.y = \"New Tag No\", all.x = TRUE, sort = FALSE)",
    "  tag_df$target_idx <- tag_df$T1",
    "  o <- order(tag_df$target_idx, suppressWarnings(as.numeric(tag_df$tag)))",
    "  tag_df <- tag_df[o, , drop = FALSE]",
    "  tag_links <- paste(sprintf(\"[ %s - S%s ](#subplot-%s)\",",
    "                           tag_df$tag, tag_df$target_idx, tag_df$target_idx),",
    "                     collapse = \" | \")",
    "}",
    "",
    "cat(\"* \", sp, \": \", tag_links, \"\\n\", sep = \"\")",
    "}",
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

.create_rmd_content_pt <- function (subplot_plots, tf_col, tf_uncol, tf_palm,
                                    plot_name, plot_code, spec_df,
                                    has_agb = FALSE) {

  yaml_head <- c(
    "---",
    "output:",
    "  pdf_document:",
    "    toc: true",
    "    toc_depth: 2",
    "    number_sections: true",
    "fontsize: 12pt",
    "params:",
    "  input_type: \"forestplots\"",
    "  metadata: NULL",
    "  main_plot: NULL",
    "  subplots_list: NULL",
    "  subplot_size: 70",
    "  stats: NULL",
    "  tag_list: NULL",
    "  tag_to_subplot: NULL"
  )

  if (has_agb)       yaml_head <- c(yaml_head, "  agb: NULL")
  if (any(tf_col))   yaml_head <- c(yaml_head, "  collected_plot: NULL")
  if (any(tf_uncol)) yaml_head <- c(yaml_head, "  uncollected_plot: NULL")
  if (any(tf_palm))  yaml_head <- c(yaml_head, "  uncollected_palm_plot: NULL")

  yaml_head <- c(yaml_head, "---")

  mid_section <- c(
    "```{r setup, include=FALSE}",
    "knitr::opts_chunk$set(echo = FALSE, warning = FALSE, message = FALSE)",
    "library(ggplot2)",
    "library(dplyr)",
    "spec_df <- params$stats$spec_df",
    "safe_id <- function(x) {",
    "  x <- trimws(as.character(x))",
    "  x <- gsub(\"\\\\s*\\\\(.*\\\\)$\", \"\", x)",
    "  x <- gsub(\"[^A-Za-z0-9]\", \"\", x)",
    "  tolower(x)",
    "}",
    "```",
    "",
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
    "```{r title-block, echo=FALSE, results='asis'}",
    "cat('\\\\begin{center}')",
    "cat('\\\\Huge\\\\textbf{Relatório Completo da Parcela} \\\\\\\\')",
    "cat('\\\\vspace{0.5em}')",
    "cat(paste0('\\\\normalsize ', params$metadata$plot_name, ' | ', params$metadata$plot_code, ' \\\\\\\\'))",
    "cat('\\\\end{center}')",
    "```",
    "",
    "```{r logo, echo=FALSE, results='asis'}",
    "norm_path <- function(p) {",
    "  p <- as.character(p); p[is.na(p)] <- \"\"",
    "  tryCatch(normalizePath(p, winslash = '/', mustWork = FALSE), error = function(e) p)",
    "}",
    "",
    "forplotr_logo_path    <- norm_path(system.file('figures', 'forplotR_hex_sticker.png', package = 'forplotR'))",
    "forestplots_logo_path <- norm_path(system.file('figures', 'forestplotsnet_logo.png',   package = 'forplotR'))",
    "",
    "raw_it <- if (!is.null(params$input_type)) params$input_type else if (!is.null(params$metadata$input_type)) params$metadata$input_type else ''",
    "input_type  <- tolower(trimws(as.character(raw_it)))",
    "is_monitora <- (input_type == 'monitora')",
    "",
    "monitora_candidates <- character(0)",
    "if (length(forestplots_logo_path) && nzchar(forestplots_logo_path)) {",
    "  mon_dir <- dirname(forestplots_logo_path[1])",
    "  monitora_candidates <- c(monitora_candidates, file.path(mon_dir, 'monitora_logo.png'))",
    "}",
    "monitora_candidates <- c(",
    "  monitora_candidates,",
    "  'figures/monitora_logo.png',",
    "  system.file('figures', 'monitora_logo.png', package = 'forplotR')",
    ")",
    "monitora_candidates <- norm_path(monitora_candidates)",
    "monitora_candidates <- monitora_candidates[nzchar(monitora_candidates)]",
    "",
    "monitora_logo_path <- ''",
    "if (is_monitora) {",
    "  idx <- which(file.exists(monitora_candidates))",
    "  if (length(idx) > 0) {",
    "    monitora_logo_path <- monitora_candidates[idx[1]]",
    "  } else {",
    "    url <- 'https://www.gov.br/icmbio/pt-br/assuntos/monitoramento/programa-monitora/@@collective.cover.banner/af7c14b7-6562-4da9-8056-dc56a2b69f7d/@@images/f86d904c-f11b-4b41-a283-84790263c284.png'",
    "    tmp <- file.path(tempdir(), 'monitora_logo.png')",
    "    try(utils::download.file(url, tmp, mode = 'wb', quiet = TRUE), silent = TRUE)",
    "    if (file.exists(tmp)) monitora_logo_path <- norm_path(tmp)",
    "  }",
    "}",
    "",
    "partner_logo_path <- if (is_monitora && nzchar(monitora_logo_path)) monitora_logo_path else forestplots_logo_path",
    "partner_href      <- if (is_monitora) 'https://www.gov.br/icmbio/pt-br/assuntos/monitoramento/programa-monitora' else 'https://forestplots.net'",
    "partner_width     <- if (is_monitora) \"0.28\\\\linewidth\" else \"0.40\\\\linewidth\"",
    "",
    "cat('\\\\vspace*{1cm}\n')",
    "cat('\\\\begin{center}\n')",
    "cat(sprintf('\\\\href{https://dboslab.github.io/forplotR-website/}{\\\\includegraphics[width=0.20\\\\linewidth]{%s}}', forplotr_logo_path), '\n')",
    "cat('\\\\end{center}\n')",
    "",
    "cat('\\\\begin{center}\n')",
    "if (is_monitora && !nzchar(partner_logo_path)) {",
    "  cat('Programa MONITORA', '\\n')",
    "} else {",
    "  cat(sprintf('\\\\href{%s}{\\\\includegraphics[width=%s,keepaspectratio]{%s}}',",
    "              partner_href, partner_width, partner_logo_path), '\n')",
    "}",
    "cat('\\\\end{center}\n')",
    "cat('\\\\vspace*{2cm}\n')",
    "```",
    "",
    "## Sumário {#contents}",
    "\\vspace*{-1cm}",
    "\\thispagestyle{plain}",
    "\\tableofcontents",
    "\\newpage",
    "",
    "## Metadados {#metadata}",
    "",
    "**Nome da parcela:** `r params$metadata$plot_name`",
    "",
    "**Código da parcela:** `r params$metadata$plot_code`",
    "",
    "**Equipe:** `r params$metadata$team`",
    "",
    "\\vspace{0.6\\baselineskip}",
    "\\hrule",
    "",
    "```{r meta-census, echo=FALSE, results='asis'}",
    "yrs <- params$metadata$census_years; yrs <- yrs[is.finite(yrs)]",
    "if (length(yrs) > 1) {",
    "  cat(paste0('**Número de censos:** ', length(yrs), '\\n\\n'))",
    "  cat(paste0('**Datas dos censos:** ', paste(yrs, collapse = ' | '), '\\n\\n'))",
    "}",
    "```",
    "\\hrule",
    "",
    "## Contagem de indivíduos {#counts}",
    "",
    "- **Total de indivíduos:** `r params$stats$total`",
    "- **Coletados (excluindo palmeiras):** `r params$stats$collected`",
    "- **Não coletados (excluindo palmeiras):** `r params$stats$uncollected`",
    "- **Palmeiras (Arecaceae):** `r params$stats$palms`",
    "",
    "\\hrule",
    "```{r counts-mc, echo=FALSE, results='asis'}",
    "yrs <- params$metadata$census_years",
    "dead_total <- params$stats$dead_since_first",
    "recr_total <- params$stats$recruits_since_first",
    "if (!is.null(yrs) && length(yrs) > 1) {",
    "  if (!is.null(dead_total) && is.finite(dead_total))",
    "    cat(paste0('- **Árvores mortas desde o primeiro censo:** ', dead_total, '\\n'))",
    "  if (!is.null(recr_total) && is.finite(recr_total))",
    "    cat(paste0('- **Recrutas desde o primeiro censo:** ', recr_total, '\\n'))",
    "}",
    "```"
  )

  if (has_agb) {
    mid_section <- c(
      mid_section,
      "",
      "## Biomassa acima do solo {#agb}",
      "",
      "```{r agb-table, echo=FALSE, results='asis'}",
      "if (!is.null(params$agb)) {",
      "  agb_val <- round(params$agb[[2]][1], 2)",
      "  cat(paste0(\"**Biomassa acima do solo:** \", agb_val, \" t ha$^{-1}$\\n\\n\"))",
      "  cat('\\\\hrule\\n')",
      "}",
      "```"
    )
  }

  mid_section <- c(mid_section, "\\newpage")

  nav_targets <- c(
    "[Voltar ao Sumário](#contents)",
    "[Metadados](#metadata)",
    "[Contagem de indivíduos](#counts)"
  )
  if (has_agb) nav_targets <- c(nav_targets, "[Biomassa acima do solo](#agb)")
  nav_targets <- c(nav_targets, "[Mapa geral da parcela](#general-plot)")
  if (any(tf_col))   nav_targets <- c(nav_targets, "[Apenas coletados](#collected-only)")
  if (any(tf_uncol)) nav_targets <- c(nav_targets, "[Não coletados](#uncollected)")
  if (any(tf_palm))  nav_targets <- c(nav_targets, "[Palmeiras](#uncollected-palm)")
  nav_targets <- c(nav_targets, "[Índice de subparcelas](#subplot-index)", "[Lista de espécies](#checklist)")

  nav_links <- c(
    "\\begingroup\\tiny\\color{gray}",
    paste("«", paste(nav_targets, collapse = " | ")),
    "\\endgroup"
  )

  gencol_section <- c(
    "",
    "## Mapa geral da parcela {#general-plot}",
    "```{r general-plot, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$main_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  col_section <- c(
    "## Apenas coletados {#collected-only}",
    "```{r collected-only, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$collected_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  uncol_section <- c(
    "## Não coletados {#uncollected}",
    "```{r uncollected,fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$uncollected_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  palm_section <- c(
    "## Palmeiras não coletadas {#uncollected-palm}",
    "```{r uncollected-palm, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$uncollected_palm_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  index_section <- c(
    "",
    "\\newpage",
    "",
    "## Índice de subparcelas {#subplot-index}",
    "",
    "```{r toc, results='asis', echo=FALSE}",
    "cols <- 5",
    "n <- length(params$subplots_list)",
    "per_col <- ceiling(n / cols)",
    "header <- paste(rep('Subparcelas', cols), collapse = ' | ')",
    "separator <- paste(rep('---', cols), collapse = ' | ')",
    "toc_lines <- c(header, separator)",
    "for (i in 1:per_col) {",
    "  row <- character(cols)",
    "  for (j in 0:(cols - 1)) {",
    "    idx <- i + j * per_col",
    "    if (idx <= n) {",
    "      row[j + 1] <- paste0('[Subparcela ', idx, '](#subplot-', idx, ')')",
    "    } else {",
    "      row[j + 1] <- ' '",
    "    }",
    "  }",
    "  toc_lines <- c(toc_lines, paste(row, collapse = ' | '))",
    "}",
    "cat(paste(toc_lines, collapse = '\\n'))",
    "```",
    "",
    "",
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
      sprintf("### Subparcela %d {#subplot-%d .unlisted .unnumbered}", i, i),
      "",
      sprintf("```{r subplot-%d, fig.width=12, fig.height=9}", i),
      sprintf("print(params$subplots_list[[%d]]$plot)", i),
      "```",
      "",
      "\\begingroup\\tiny\\color{gray}",
      "« [Voltar ao Sumário](#contents) | [Metadados](#metadata) | [Contagem de indivíduos](#counts) | [Mapa geral da parcela](#general-plot) | [Apenas coletados](#collected-only) | [Não coletados](#uncollected) | [Palmeiras](#uncollected-palm) | [Índice de subparcelas](#subplot-index) | [Lista de espécies](#checklist)",
      "\\endgroup",
      if (i < length(subplot_plots)) "\\newpage" else NULL,
      ""
    )
  }))

  checklist_section <- c(
    "",
    "\\newpage",
    "",
    "## Lista de espécies {#checklist}",
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
    "      stringsAsFactors = FALSE)",
    "",
    "t2s_df <- params$tag_to_subplot",
    "is_monitora <- tolower(trimws(as.character(params$input_type))) == \"monitora\"",
    "",
    "if (is_monitora) {",
    "  tag_df <- merge(tag_df, t2s_df, by.x = \"tag\", by.y = \"New Tag No\", all.x = TRUE, sort = FALSE)",
    "  tag_df$target_idx <- ifelse(!is.na(tag_df$subplot_index), tag_df$subplot_index, tag_df$T1)",
    "  has_fields <- !is.na(tag_df$subunit_letter) & !is.na(tag_df$T2)",
    "  tag_df$subplot_code <- ifelse(has_fields, paste0(tag_df$subunit_letter, tag_df$T2), paste0(\"S\", tag_df$target_idx))",
    "  o <- order(tag_df$target_idx, suppressWarnings(as.numeric(tag_df$tag)))",
    "  tag_df <- tag_df[o, , drop = FALSE]",
    "  tag_links <- paste(sprintf(\"[ %s → %s ](#subplot-%s)\",",
    "                           tag_df$tag, tag_df$subplot_code, tag_df$target_idx),",
    "                     collapse = \" | \")",
    "} else {",
    "  tag_df <- merge(tag_df, t2s_df, by.x = \"tag\", by.y = \"New Tag No\", all.x = TRUE, sort = FALSE)",
    "  tag_df$target_idx <- tag_df$T1",
    "  o <- order(tag_df$target_idx, suppressWarnings(as.numeric(tag_df$tag)))",
    "  tag_df <- tag_df[o, , drop = FALSE]",
    "  tag_links <- paste(sprintf(\"[ %s - S%s ](#subplot-%s)\",",
    "                           tag_df$tag, tag_df$target_idx, tag_df$target_idx),",
    "                     collapse = \" | \")",
    "}",
    "",
    "cat(\"* \", sp, \": \", tag_links, \"\\n\", sep = \"\")",
    "}",
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

  add_body("## Subparcelas individuais {#individual-subplots}", subplot_sections, checklist_section)
  return(rmd_content)
}

#### Portuguese version ####
.create_rmd_content_pt <- function (subplot_plots, tf_col, tf_uncol, tf_palm,
                                    plot_name, plot_code, spec_df,
                                    has_agb = FALSE) {

  yaml_head <- c(
    "---",
    "output:",
    "  pdf_document:",
    "    toc: true",
    "    toc_depth: 2",
    "    number_sections: true",
    "fontsize: 12pt",
    "params:",
    "  input_type: \"forestplots\"",
    "  metadata: NULL",
    "  main_plot: NULL",
    "  subplots_list: NULL",
    "  subplot_size: 70",
    "  stats: NULL",
    "  tag_list: NULL",
    "  tag_to_subplot: NULL"
  )

  if (has_agb)       yaml_head <- c(yaml_head, "  agb: NULL")
  if (any(tf_col))   yaml_head <- c(yaml_head, "  collected_plot: NULL")
  if (any(tf_uncol)) yaml_head <- c(yaml_head, "  uncollected_plot: NULL")
  if (any(tf_palm))  yaml_head <- c(yaml_head, "  uncollected_palm_plot: NULL")

  yaml_head <- c(yaml_head, "---")

  mid_section <- c(
    "```{r setup, include=FALSE}",
    "knitr::opts_chunk$set(echo = FALSE, warning = FALSE, message = FALSE)",
    "library(ggplot2)",
    "library(dplyr)",
    "spec_df <- params$stats$spec_df",
    "safe_id <- function(x) {",
    "  x <- trimws(as.character(x))",
    "  x <- gsub(\"\\\\s*\\\\(.*\\\\)$\", \"\", x)",
    "  x <- gsub(\"[^A-Za-z0-9]\", \"\", x)",
    "  tolower(x)",
    "}",
    "```",
    "",
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
    "```{r title-block, echo=FALSE, results='asis'}",
    "cat('\\\\begin{center}')",
    "cat('\\\\Huge\\\\textbf{Relatório Completo da Parcela} \\\\\\\\')",
    "cat('\\\\vspace{0.5em}')",
    "cat(paste0('\\\\normalsize ', params$metadata$plot_name, ' | ', params$metadata$plot_code, ' \\\\\\\\'))",
    "cat('\\\\end{center}')",
    "```",
    "",
    "```{r logo, echo=FALSE, results='asis'}",
    "norm_path <- function(p) {",
    "  p <- as.character(p); p[is.na(p)] <- \"\"",
    "  tryCatch(normalizePath(p, winslash = '/', mustWork = FALSE), error = function(e) p)",
    "}",
    "",
    "forplotr_logo_path    <- norm_path(system.file('figures', 'forplotR_hex_sticker.png', package = 'forplotR'))",
    "forestplots_logo_path <- norm_path(system.file('figures', 'forestplotsnet_logo.png',   package = 'forplotR'))",
    "",
    "raw_it <- if (!is.null(params$input_type)) params$input_type else if (!is.null(params$metadata$input_type)) params$metadata$input_type else ''",
    "input_type  <- tolower(trimws(as.character(raw_it)))",
    "is_monitora <- (input_type == 'monitora')",
    "",
    "monitora_candidates <- character(0)",
    "if (length(forestplots_logo_path) && nzchar(forestplots_logo_path)) {",
    "  mon_dir <- dirname(forestplots_logo_path[1])",
    "  monitora_candidates <- c(monitora_candidates, file.path(mon_dir, 'monitora_logo.png'))",
    "}",
    "monitora_candidates <- c(",
    "  monitora_candidates,",
    "  'figures/monitora_logo.png',",
    "  system.file('figures', 'monitora_logo.png', package = 'forplotR')",
    ")",
    "monitora_candidates <- norm_path(monitora_candidates)",
    "monitora_candidates <- monitora_candidates[nzchar(monitora_candidates)]",
    "",
    "monitora_logo_path <- ''",
    "if (is_monitora) {",
    "  idx <- which(file.exists(monitora_candidates))",
    "  if (length(idx) > 0) {",
    "    monitora_logo_path <- monitora_candidates[idx[1]]",
    "  } else {",
    "    url <- 'https://www.gov.br/icmbio/pt-br/assuntos/monitoramento/programa-monitora/@@collective.cover.banner/af7c14b7-6562-4da9-8056-dc56a2b69f7d/@@images/f86d904c-f11b-4b41-a283-84790263c284.png'",
    "    tmp <- file.path(tempdir(), 'monitora_logo.png')",
    "    try(utils::download.file(url, tmp, mode = 'wb', quiet = TRUE), silent = TRUE)",
    "    if (file.exists(tmp)) monitora_logo_path <- norm_path(tmp)",
    "  }",
    "}",
    "",
    "partner_logo_path <- if (is_monitora && nzchar(monitora_logo_path)) monitora_logo_path else forestplots_logo_path",
    "partner_href      <- if (is_monitora) 'https://www.gov.br/icmbio/pt-br/assuntos/monitoramento/programa-monitora' else 'https://forestplots.net'",
    "partner_width     <- if (is_monitora) \"0.28\\\\linewidth\" else \"0.40\\\\linewidth\"",
    "",
    "cat('\\\\vspace*{1cm}\n')",
    "cat('\\\\begin{center}\n')",
    "cat(sprintf('\\\\href{https://dboslab.github.io/forplotR-website/}{\\\\includegraphics[width=0.20\\\\linewidth]{%s}}', forplotr_logo_path), '\n')",
    "cat('\\\\end{center}\n')",
    "",
    "cat('\\\\begin{center}\n')",
    "if (is_monitora && !nzchar(partner_logo_path)) {",
    "  cat('Programa MONITORA', '\\n')",
    "} else {",
    "  cat(sprintf('\\\\href{%s}{\\\\includegraphics[width=%s,keepaspectratio]{%s}}',",
    "              partner_href, partner_width, partner_logo_path), '\n')",
    "}",
    "cat('\\\\end{center}\n')",
    "cat('\\\\vspace*{2cm}\n')",
    "```",
    "",
    "## Sumário {#contents}",
    "\\vspace*{-1cm}",
    "\\thispagestyle{plain}",
    "\\tableofcontents",
    "\\newpage",
    "",
    "## Metadados {#metadata}",
    "",
    "**Nome da parcela:** `r params$metadata$plot_name`",
    "",
    "**Código da parcela:** `r params$metadata$plot_code`",
    "",
    "**Equipe:** `r params$metadata$team`",
    "",
    "\\vspace{0.6\\baselineskip}",
    "\\hrule",
    "",
    "```{r meta-census, echo=FALSE, results='asis'}",
    "yrs <- params$metadata$census_years; yrs <- yrs[is.finite(yrs)]",
    "if (length(yrs) > 1) {",
    "  cat(paste0('**Número de censos:** ', length(yrs), '\\n\\n'))",
    "  cat(paste0('**Datas dos censos:** ', paste(yrs, collapse = ' | '), '\\n\\n'))",
    "}",
    "```",
    "\\hrule",
    "",
    "## Contagem de indivíduos {#counts}",
    "",
    "- **Total de indivíduos:** `r params$stats$total`",
    "- **Coletados (excluindo palmeiras):** `r params$stats$collected`",
    "- **Não coletados (excluindo palmeiras):** `r params$stats$uncollected`",
    "- **Palmeiras (Arecaceae):** `r params$stats$palms`",
    "",
    "\\hrule",
    "```{r counts-mc, echo=FALSE, results='asis'}",
    "yrs <- params$metadata$census_years",
    "dead_total <- params$stats$dead_since_first",
    "recr_total <- params$stats$recruits_since_first",
    "if (!is.null(yrs) && length(yrs) > 1) {",
    "  if (!is.null(dead_total) && is.finite(dead_total))",
    "    cat(paste0('- **Árvores mortas desde o primeiro censo:** ', dead_total, '\\n'))",
    "  if (!is.null(recr_total) && is.finite(recr_total))",
    "    cat(paste0('- **Recrutas desde o primeiro censo:** ', recr_total, '\\n'))",
    "}",
    "```"
  )

  if (has_agb) {
    mid_section <- c(
      mid_section,
      "",
      "## Biomassa acima do solo {#agb}",
      "",
      "```{r agb-table, echo=FALSE, results='asis'}",
      "if (!is.null(params$agb)) {",
      "  agb_val <- round(params$agb[[2]][1], 2)",
      "  cat(paste0(\"**Biomassa acima do solo:** \", agb_val, \" t ha$^{-1}$\\n\\n\"))",
      "  cat('\\\\hrule\\n')",
      "}",
      "```"
    )
  }

  mid_section <- c(mid_section, "\\newpage")

  nav_targets <- c(
    "[Voltar ao Sumário](#contents)",
    "[Metadados](#metadata)",
    "[Contagem de indivíduos](#counts)"
  )
  if (has_agb) nav_targets <- c(nav_targets, "[Biomassa acima do solo](#agb)")
  nav_targets <- c(nav_targets, "[Mapa geral da parcela](#general-plot)")
  if (any(tf_col))   nav_targets <- c(nav_targets, "[Apenas coletados](#collected-only)")
  if (any(tf_uncol)) nav_targets <- c(nav_targets, "[Não coletados](#uncollected)")
  if (any(tf_palm))  nav_targets <- c(nav_targets, "[Palmeiras](#uncollected-palm)")
  nav_targets <- c(nav_targets, "[Índice de subparcelas](#subplot-index)", "[Lista de espécies](#checklist)")

  nav_links <- c(
    "\\begingroup\\tiny\\color{gray}",
    paste("«", paste(nav_targets, collapse = " | ")),
    "\\endgroup"
  )

  gencol_section <- c(
    "",
    "## Mapa geral da parcela {#general-plot}",
    "```{r general-plot, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$main_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  col_section <- c(
    "## Apenas coletados {#collected-only}",
    "```{r collected-only, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$collected_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  uncol_section <- c(
    "## Não coletados {#uncollected}",
    "```{r uncollected,fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$uncollected_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  palm_section <- c(
    "## Palmeiras não coletadas {#uncollected-palm}",
    "```{r uncollected-palm, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$uncollected_palm_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  index_section <- c(
    "",
    "\\newpage",
    "",
    "## Índice de subparcelas {#subplot-index}",
    "",
    "```{r toc, results='asis', echo=FALSE}",
    "cols <- 5",
    "n <- length(params$subplots_list)",
    "per_col <- ceiling(n / cols)",
    "header <- paste(rep('Subparcelas', cols), collapse = ' | ')",
    "separator <- paste(rep('---', cols), collapse = ' | ')",
    "toc_lines <- c(header, separator)",
    "for (i in 1:per_col) {",
    "  row <- character(cols)",
    "  for (j in 0:(cols - 1)) {",
    "    idx <- i + j * per_col",
    "    if (idx <= n) {",
    "      row[j + 1] <- paste0('[Subparcela ', idx, '](#subplot-', idx, ')')",
    "    } else {",
    "      row[j + 1] <- ' '",
    "    }",
    "  }",
    "  toc_lines <- c(toc_lines, paste(row, collapse = ' | '))",
    "}",
    "cat(paste(toc_lines, collapse = '\\n'))",
    "```",
    "",
    "",
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
      sprintf("### Subparcela %d {#subplot-%d .unlisted .unnumbered}", i, i),
      "",
      sprintf("```{r subplot-%d, fig.width=12, fig.height=9}", i),
      sprintf("print(params$subplots_list[[%d]]$plot)", i),
      "```",
      "",
      "\\begingroup\\tiny\\color{gray}",
      "« [Voltar ao Sumário](#contents) | [Metadados](#metadata) | [Contagem de indivíduos](#counts) | [Mapa geral da parcela](#general-plot) | [Apenas coletados](#collected-only) | [Não coletados](#uncollected) | [Palmeiras](#uncollected-palm) | [Índice de subparcelas](#subplot-index) | [Lista de espécies](#checklist)",
      "\\endgroup",
      if (i < length(subplot_plots)) "\\newpage" else NULL,
      ""
    )
  }))

  checklist_section <- c(
    "",
    "\\newpage",
    "",
    "## Lista de espécies {#checklist}",
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
    "      stringsAsFactors = FALSE)",
    "",
    "t2s_df <- params$tag_to_subplot",
    "is_monitora <- tolower(trimws(as.character(params$input_type))) == \"monitora\"",
    "",
    "if (is_monitora) {",
    "  tag_df <- merge(tag_df, t2s_df, by.x = \"tag\", by.y = \"New Tag No\", all.x = TRUE, sort = FALSE)",
    "  tag_df$target_idx <- ifelse(!is.na(tag_df$subplot_index), tag_df$subplot_index, tag_df$T1)",
    "  has_fields <- !is.na(tag_df$subunit_letter) & !is.na(tag_df$T2)",
    "  tag_df$subplot_code <- ifelse(has_fields, paste0(tag_df$subunit_letter, tag_df$T2), paste0(\"S\", tag_df$target_idx))",
    "  o <- order(tag_df$target_idx, suppressWarnings(as.numeric(tag_df$tag)))",
    "  tag_df <- tag_df[o, , drop = FALSE]",
    "  tag_links <- paste(sprintf(\"[ %s → %s ](#subplot-%s)\",",
    "                           tag_df$tag, tag_df$subplot_code, tag_df$target_idx),",
    "                     collapse = \" | \")",
    "} else {",
    "  tag_df <- merge(tag_df, t2s_df, by.x = \"tag\", by.y = \"New Tag No\", all.x = TRUE, sort = FALSE)",
    "  tag_df$target_idx <- tag_df$T1",
    "  o <- order(tag_df$target_idx, suppressWarnings(as.numeric(tag_df$tag)))",
    "  tag_df <- tag_df[o, , drop = FALSE]",
    "  tag_links <- paste(sprintf(\"[ %s - S%s ](#subplot-%s)\",",
    "                           tag_df$tag, tag_df$target_idx, tag_df$target_idx),",
    "                     collapse = \" | \")",
    "}",
    "",
    "cat(\"* \", sp, \": \", tag_links, \"\\n\", sep = \"\")",
    "}",
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

  add_body("## Subparcelas individuais {#individual-subplots}", subplot_sections, checklist_section)
  return(rmd_content)
}

#### Spanish version ####
.create_rmd_content_es <- function (subplot_plots, tf_col, tf_uncol, tf_palm,
                                    plot_name, plot_code, spec_df,
                                    has_agb = FALSE) {

  yaml_head <- c(
    "---",
    "output:",
    "  pdf_document:",
    "    toc: true",
    "    toc_depth: 2",
    "    number_sections: true",
    "fontsize: 12pt",
    "params:",
    "  input_type: \"forestplots\"",
    "  metadata: NULL",
    "  main_plot: NULL",
    "  subplots_list: NULL",
    "  subplot_size: 70",
    "  stats: NULL",
    "  tag_list: NULL",
    "  tag_to_subplot: NULL"
  )

  if (has_agb)       yaml_head <- c(yaml_head, "  agb: NULL")
  if (any(tf_col))   yaml_head <- c(yaml_head, "  collected_plot: NULL")
  if (any(tf_uncol)) yaml_head <- c(yaml_head, "  uncollected_plot: NULL")
  if (any(tf_palm))  yaml_head <- c(yaml_head, "  uncollected_palm_plot: NULL")

  yaml_head <- c(yaml_head, "---")

  mid_section <- c(
    "```{r setup, include=FALSE}",
    "knitr::opts_chunk$set(echo = FALSE, warning = FALSE, message = FALSE)",
    "library(ggplot2)",
    "library(dplyr)",
    "spec_df <- params$stats$spec_df",
    "safe_id <- function(x) {",
    "  x <- trimws(as.character(x))",
    "  x <- gsub(\"\\\\s*\\\\(.*\\\\)$\", \"\", x)",
    "  x <- gsub(\"[^A-Za-z0-9]\", \"\", x)",
    "  tolower(x)",
    "}",
    "```",
    "",
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
    "```{r title-block, echo=FALSE, results='asis'}",
    "cat('\\\\begin{center}')",
    "cat('\\\\Huge\\\\textbf{Informe completo de la parcela} \\\\\\\\')",
    "cat('\\\\vspace{0.5em}')",
    "cat(paste0('\\\\normalsize ', params$metadata$plot_name, ' | ', params$metadata$plot_code, ' \\\\\\\\'))",
    "cat('\\\\end{center}')",
    "```",
    "",
    "```{r logo, echo=FALSE, results='asis'}",
    "norm_path <- function(p) {",
    "  p <- as.character(p); p[is.na(p)] <- \"\"",
    "  tryCatch(normalizePath(p, winslash = '/', mustWork = FALSE), error = function(e) p)",
    "}",
    "",
    "forplotr_logo_path    <- norm_path(system.file('figures', 'forplotR_hex_sticker.png', package = 'forplotR'))",
    "forestplots_logo_path <- norm_path(system.file('figures', 'forestplotsnet_logo.png',   package = 'forplotR'))",
    "",
    "raw_it <- if (!is.null(params$input_type)) params$input_type else if (!is.null(params$metadata$input_type)) params$metadata$input_type else ''",
    "input_type  <- tolower(trimws(as.character(raw_it)))",
    "is_monitora <- (input_type == 'monitora')",
    "",
    "monitora_candidates <- character(0)",
    "if (length(forestplots_logo_path) && nzchar(forestplots_logo_path)) {",
    "  mon_dir <- dirname(forestplots_logo_path[1])",
    "  monitora_candidates <- c(monitora_candidates, file.path(mon_dir, 'monitora_logo.png'))",
    "}",
    "monitora_candidates <- c(",
    "  monitora_candidates,",
    "  'figures/monitora_logo.png',",
    "  system.file('figures', 'monitora_logo.png', package = 'forplotR')",
    ")",
    "monitora_candidates <- norm_path(monitora_candidates)",
    "monitora_candidates <- monitora_candidates[nzchar(monitora_candidates)]",
    "",
    "monitora_logo_path <- ''",
    "if (is_monitora) {",
    "  idx <- which(file.exists(monitora_candidates))",
    "  if (length(idx) > 0) {",
    "    monitora_logo_path <- monitora_candidates[idx[1]]",
    "  } else {",
    "    url <- 'https://www.gov.br/icmbio/pt-br/assuntos/monitoramento/programa-monitora/@@collective.cover.banner/af7c14b7-6562-4da9-8056-dc56a2b69f7d/@@images/f86d904c-f11b-4b41-a283-84790263c284.png'",
    "    tmp <- file.path(tempdir(), 'monitora_logo.png')",
    "    try(utils::download.file(url, tmp, mode = 'wb', quiet = TRUE), silent = TRUE)",
    "    if (file.exists(tmp)) monitora_logo_path <- norm_path(tmp)",
    "  }",
    "}",
    "",
    "partner_logo_path <- if (is_monitora && nzchar(monitora_logo_path)) monitora_logo_path else forestplots_logo_path",
    "partner_href      <- if (is_monitora) 'https://www.gov.br/icmbio/pt-br/assuntos/monitoramento/programa-monitora' else 'https://forestplots.net'",
    "partner_width     <- if (is_monitora) \"0.28\\\\linewidth\" else \"0.40\\\\linewidth\"",
    "",
    "cat('\\\\vspace*{1cm}\n')",
    "cat('\\\\begin{center}\n')",
    "cat(sprintf('\\\\href{https://dboslab.github.io/forplotR-website/}{\\\\includegraphics[width=0.20\\\\linewidth]{%s}}', forplotr_logo_path), '\n')",
    "cat('\\\\end{center}\n')",
    "",
    "cat('\\\\begin{center}\n')",
    "if (is_monitora && !nzchar(partner_logo_path)) {",
    "  cat('Programa MONITORA', '\\n')",
    "} else {",
    "  cat(sprintf('\\\\href{%s}{\\\\includegraphics[width=%s,keepaspectratio]{%s}}',",
    "              partner_href, partner_width, partner_logo_path), '\n')",
    "}",
    "cat('\\\\end{center}\n')",
    "cat('\\\\vspace*{2cm}\n')",
    "```",
    "",
    "## Contenido {#contents}",
    "\\vspace*{-1cm}",
    "\\thispagestyle{plain}",
    "\\tableofcontents",
    "\\newpage",
    "",
    "## Metadatos {#metadata}",
    "",
    "**Nombre de la parcela:** `r params$metadata$plot_name`",
    "",
    "**Código de la parcela:** `r params$metadata$plot_code`",
    "",
    "**Equipo:** `r params$metadata$team`",
    "",
    "\\vspace{0.6\\baselineskip}",
    "\\hrule",
    "",
    "```{r meta-census, echo=FALSE, results='asis'}",
    "yrs <- params$metadata$census_years; yrs <- yrs[is.finite(yrs)]",
    "if (length(yrs) > 1) {",
    "  cat(paste0('**Número de censos:** ', length(yrs), '\\n\\n'))",
    "  cat(paste0('**Fechas de los censos:** ', paste(yrs, collapse = ' | '), '\\n\\n'))",
    "}",
    "```",
    "\\hrule",
    "",
    "## Conteo de individuos {#counts}",
    "",
    "- **Total de individuos:** `r params$stats$total`",
    "- **Colectados (excluyendo palmas):** `r params$stats$collected`",
    "- **No colectados (excluyendo palmas):** `r params$stats$uncollected`",
    "- **Palmas (Arecaceae):** `r params$stats$palms`",
    "",
    "\\hrule",
    "```{r counts-mc, echo=FALSE, results='asis'}",
    "yrs <- params$metadata$census_years",
    "dead_total <- params$stats$dead_since_first",
    "recr_total <- params$stats$recruits_since_first",
    "if (!is.null(yrs) && length(yrs) > 1) {",
    "  if (!is.null(dead_total) && is.finite(dead_total))",
    "    cat(paste0('- **Árboles muertos desde el primer censo:** ', dead_total, '\\n'))",
    "  if (!is.null(recr_total) && is.finite(recr_total))",
    "    cat(paste0('- **Reclutas desde el primer censo:** ', recr_total, '\\n'))",
    "}",
    "```"
  )

  if (has_agb) {
    mid_section <- c(
      mid_section,
      "",
      "## Biomasa aérea {#agb}",
      "",
      "```{r agb-table, echo=FALSE, results='asis'}",
      "if (!is.null(params$agb)) {",
      "  agb_val <- round(params$agb[[2]][1], 2)",
      "  cat(paste0(\"**Biomasa aérea:** \", agb_val, \" t ha$^{-1}$\\n\\n\"))",
      "  cat('\\\\hrule\\n')",
      "}",
      "```"
    )
  }

  mid_section <- c(mid_section, "\\newpage")

  nav_targets <- c(
    "[Volver al contenido](#contents)",
    "[Metadatos](#metadata)",
    "[Conteo de individuos](#counts)"
  )
  if (has_agb) nav_targets <- c(nav_targets, "[Biomasa aérea](#agb)")
  nav_targets <- c(nav_targets, "[Mapa general de la parcela](#general-plot)")
  if (any(tf_col))   nav_targets <- c(nav_targets, "[Solo colectados](#collected-only)")
  if (any(tf_uncol)) nav_targets <- c(nav_targets, "[No colectados](#uncollected)")
  if (any(tf_palm))  nav_targets <- c(nav_targets, "[Palmas](#uncollected-palm)")
  nav_targets <- c(nav_targets, "[Índice de subparcelas](#subplot-index)", "[Lista de especies](#checklist)")

  nav_links <- c(
    "\\begingroup\\tiny\\color{gray}",
    paste("«", paste(nav_targets, collapse = " | ")),
    "\\endgroup"
  )

  gencol_section <- c(
    "",
    "## Mapa general de la parcela {#general-plot}",
    "```{r general-plot, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$main_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  col_section <- c(
    "## Solo colectados {#collected-only}",
    "```{r collected-only, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$collected_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  uncol_section <- c(
    "## No colectados {#uncollected}",
    "```{r uncollected,fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$uncollected_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  palm_section <- c(
    "## Palmas no colectadas {#uncollected-palm}",
    "```{r uncollected-palm, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$uncollected_palm_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  index_section <- c(
    "",
    "\\newpage",
    "",
    "## Índice de subparcelas {#subplot-index}",
    "",
    "```{r toc, results='asis', echo=FALSE}",
    "cols <- 5",
    "n <- length(params$subplots_list)",
    "per_col <- ceiling(n / cols)",
    "header <- paste(rep('Subparcelas', cols), collapse = ' | ')",
    "separator <- paste(rep('---', cols), collapse = ' | ')",
    "toc_lines <- c(header, separator)",
    "for (i in 1:per_col) {",
    "  row <- character(cols)",
    "  for (j in 0:(cols - 1)) {",
    "    idx <- i + j * per_col",
    "    if (idx <= n) {",
    "      row[j + 1] <- paste0('[Subparcela ', idx, '](#subplot-', idx, ')')",
    "    } else {",
    "      row[j + 1] <- ' '",
    "    }",
    "  }",
    "  toc_lines <- c(toc_lines, paste(row, collapse = ' | '))",
    "}",
    "cat(paste(toc_lines, collapse = '\\n'))",
    "```",
    "",
    "",
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
      sprintf("### Subparcela %d {#subplot-%d .unlisted .unnumbered}", i, i),
      "",
      sprintf("```{r subplot-%d, fig.width=12, fig.height=9}", i),
      sprintf("print(params$subplots_list[[%d]]$plot)", i),
      "```",
      "",
      "\\begingroup\\tiny\\color{gray}",
      "« [Volver al contenido](#contents) | [Metadatos](#metadata) | [Conteo de individuos](#counts) | [Mapa general de la parcela](#general-plot) | [Solo colectados](#collected-only) | [No colectados](#uncollected) | [Palmas](#uncollected-palm) | [Índice de subparcelas](#subplot-index) | [Lista de especies](#checklist)",
      "\\endgroup",
      if (i < length(subplot_plots)) "\\newpage" else NULL,
      ""
    )
  }))

  checklist_section <- c(
    "",
    "\\newpage",
    "",
    "## Lista de especies {#checklist}",
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
    "      stringsAsFactors = FALSE)",
    "",
    "t2s_df <- params$tag_to_subplot",
    "is_monitora <- tolower(trimws(as.character(params$input_type))) == \"monitora\"",
    "",
    "if (is_monitora) {",
    "  tag_df <- merge(tag_df, t2s_df, by.x = \"tag\", by.y = \"New Tag No\", all.x = TRUE, sort = FALSE)",
    "  tag_df$target_idx <- ifelse(!is.na(tag_df$subplot_index), tag_df$subplot_index, tag_df$T1)",
    "  has_fields <- !is.na(tag_df$subunit_letter) & !is.na(tag_df$T2)",
    "  tag_df$subplot_code <- ifelse(has_fields, paste0(tag_df$subunit_letter, tag_df$T2), paste0(\"S\", tag_df$target_idx))",
    "  o <- order(tag_df$target_idx, suppressWarnings(as.numeric(tag_df$tag)))",
    "  tag_df <- tag_df[o, , drop = FALSE]",
    "  tag_links <- paste(sprintf(\"[ %s → %s ](#subplot-%s)\",",
    "                           tag_df$tag, tag_df$subplot_code, tag_df$target_idx),",
    "                     collapse = \" | \")",
    "} else {",
    "  tag_df <- merge(tag_df, t2s_df, by.x = \"tag\", by.y = \"New Tag No\", all.x = TRUE, sort = FALSE)",
    "  tag_df$target_idx <- tag_df$T1",
    "  o <- order(tag_df$target_idx, suppressWarnings(as.numeric(tag_df$tag)))",
    "  tag_df <- tag_df[o, , drop = FALSE]",
    "  tag_links <- paste(sprintf(\"[ %s - S%s ](#subplot-%s)\",",
    "                           tag_df$tag, tag_df$target_idx, tag_df$target_idx),",
    "                     collapse = \" | \")",
    "}",
    "",
    "cat(\"* \", sp, \": \", tag_links, \"\\n\", sep = \"\")",
    "}",
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
    add_body(mid_section, gencol_section, col_section, uncollected_section = uncol_section, index_section)
  } else if (any(tf_col) && !any(tf_uncol) && any(tf_palm)) {
    add_body(mid_section, gencol_section, col_section, palm_section, index_section)
  } else if (!any(tf_col) && any(tf_uncol) && any(tf_palm)) {
    add_body(mid_section, gencol_section, uncol_section, palm_section, index_section)
  } else {
    add_body(mid_section, gencol_section, index_section)
  }

  add_body("## Subparcelas individuales {#individual-subplots}", subplot_sections, checklist_section)
  return(rmd_content)
}

#### Mandarin version ####
.create_rmd_content_ma <- function (subplot_plots, tf_col, tf_uncol, tf_palm,
                                    plot_name, plot_code, spec_df,
                                    has_agb = FALSE) {

  yaml_head <- c(
    "---",
    "output:",
    "  pdf_document:",
    "    toc: true",
    "    toc_depth: 2",
    "    number_sections: true",
    "fontsize: 12pt",
    "params:",
    "  input_type: \"forestplots\"",
    "  metadata: NULL",
    "  main_plot: NULL",
    "  subplots_list: NULL",
    "  subplot_size: 70",
    "  stats: NULL",
    "  tag_list: NULL",
    "  tag_to_subplot: NULL"
  )

  if (has_agb)       yaml_head <- c(yaml_head, "  agb: NULL")
  if (any(tf_col))   yaml_head <- c(yaml_head, "  collected_plot: NULL")
  if (any(tf_uncol)) yaml_head <- c(yaml_head, "  uncollected_plot: NULL")
  if (any(tf_palm))  yaml_head <- c(yaml_head, "  uncollected_palm_plot: NULL")

  yaml_head <- c(yaml_head, "---")

  mid_section <- c(
    "```{r setup, include=FALSE}",
    "knitr::opts_chunk$set(echo = FALSE, warning = FALSE, message = FALSE)",
    "library(ggplot2)",
    "library(dplyr)",
    "spec_df <- params$stats$spec_df",
    "safe_id <- function(x) {",
    "  x <- trimws(as.character(x))",
    "  x <- gsub(\"\\\\s*\\\\(.*\\\\)$\", \"\", x)",
    "  x <- gsub(\"[^A-Za-z0-9]\", \"\", x)",
    "  tolower(x)",
    "}",
    "```",
    "",
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
    "```{r title-block, echo=FALSE, results='asis'}",
    "cat('\\\\begin{center}')",
    "cat('\\\\Huge\\\\textbf{样地完整报告} \\\\\\\\')",
    "cat('\\\\vspace{0.5em}')",
    "cat(paste0('\\\\normalsize ', params$metadata$plot_name, ' | ', params$metadata$plot_code, ' \\\\\\\\'))",
    "cat('\\\\end{center}')",
    "```",
    "",
    "```{r logo, echo=FALSE, results='asis'}",
    "norm_path <- function(p) {",
    "  p <- as.character(p); p[is.na(p)] <- \"\"",
    "  tryCatch(normalizePath(p, winslash = '/', mustWork = FALSE), error = function(e) p)",
    "}",
    "",
    "forplotr_logo_path    <- norm_path(system.file('figures', 'forplotR_hex_sticker.png', package = 'forplotR'))",
    "forestplots_logo_path <- norm_path(system.file('figures', 'forestplotsnet_logo.png',   package = 'forplotR'))",
    "",
    "raw_it <- if (!is.null(params$input_type)) params$input_type else if (!is.null(params$metadata$input_type)) params$metadata$input_type else ''",
    "input_type  <- tolower(trimws(as.character(raw_it)))",
    "is_monitora <- (input_type == 'monitora')",
    "",
    "monitora_candidates <- character(0)",
    "if (length(forestplots_logo_path) && nzchar(forestplots_logo_path)) {",
    "  mon_dir <- dirname(forestplots_logo_path[1])",
    "  monitora_candidates <- c(monitora_candidates, file.path(mon_dir, 'monitora_logo.png'))",
    "}",
    "monitora_candidates <- c(",
    "  monitora_candidates,",
    "  'figures/monitora_logo.png',",
    "  system.file('figures', 'monitora_logo.png', package = 'forplotR')",
    ")",
    "monitora_candidates <- norm_path(monitora_candidates)",
    "monitora_candidates <- monitora_candidates[nzchar(monitora_candidates)]",
    "",
    "monitora_logo_path <- ''",
    "if (is_monitora) {",
    "  idx <- which(file.exists(monitora_candidates))",
    "  if (length(idx) > 0) {",
    "    monitora_logo_path <- monitora_candidates[idx[1]]",
    "  } else {",
    "    url <- 'https://www.gov.br/icmbio/pt-br/assuntos/monitoramento/programa-monitora/@@collective.cover.banner/af7c14b7-6562-4da9-8056-dc56a2b69f7d/@@images/f86d904c-f11b-4b41-a283-84790263c284.png'",
    "    tmp <- file.path(tempdir(), 'monitora_logo.png')",
    "    try(utils::download.file(url, tmp, mode = 'wb', quiet = TRUE), silent = TRUE)",
    "    if (file.exists(tmp)) monitora_logo_path <- norm_path(tmp)",
    "  }",
    "}",
    "",
    "partner_logo_path <- if (is_monitora && nzchar(monitora_logo_path)) monitora_logo_path else forestplots_logo_path",
    "partner_href      <- if (is_monitora) 'https://www.gov.br/icmbio/pt-br/assuntos/monitoramento/programa-monitora' else 'https://forestplots.net'",
    "partner_width     <- if (is_monitora) \"0.28\\\\linewidth\" else \"0.40\\\\linewidth\"",
    "",
    "cat('\\\\vspace*{1cm}\n')",
    "cat('\\\\begin{center}\n')",
    "cat(sprintf('\\\\href{https://dboslab.github.io/forplotR-website/}{\\\\includegraphics[width=0.20\\\\linewidth]{%s}}', forplotr_logo_path), '\n')",
    "cat('\\\\end{center}\n')",
    "",
    "cat('\\\\begin{center}\n')",
    "if (is_monitora && !nzchar(partner_logo_path)) {",
    "  cat('MONITORA 项目', '\\n')",
    "} else {",
    "  cat(sprintf('\\\\href{%s}{\\\\includegraphics[width=%s,keepaspectratio]{%s}}',",
    "              partner_href, partner_width, partner_logo_path), '\n')",
    "}",
    "cat('\\\\end{center}\n')",
    "cat('\\\\vspace*{2cm}\n')",
    "```",
    "",
    "## 目录 {#contents}",
    "\\vspace*{-1cm}",
    "\\thispagestyle{plain}",
    "\\tableofcontents",
    "\\newpage",
    "",
    "## 元数据 {#metadata}",
    "",
    "**样地名称：** `r params$metadata$plot_name`",
    "",
    "**样地代码：** `r params$metadata$plot_code`",
    "",
    "**团队：** `r params$metadata$team`",
    "",
    "\\vspace{0.6\\baselineskip}",
    "\\hrule",
    "",
    "```{r meta-census, echo=FALSE, results='asis'}",
    "yrs <- params$metadata$census_years; yrs <- yrs[is.finite(yrs)]",
    "if (length(yrs) > 1) {",
    "  cat(paste0('**普查次数：** ', length(yrs), '\\n\\n'))",
    "  cat(paste0('**普查日期：** ', paste(yrs, collapse = ' | '), '\\n\\n'))",
    "}",
    "```",
    "\\hrule",
    "",
    "## 个体统计 {#counts}",
    "",
    "- **个体总数：** `r params$stats$total`",
    "- **已采集（不含棕榈科）：** `r params$stats$collected`",
    "- **未采集（不含棕榈科）：** `r params$stats$uncollected`",
    "- **棕榈科个体 (Arecaceae)：** `r params$stats$palms`",
    "",
    "\\hrule",
    "```{r counts-mc, echo=FALSE, results='asis'}",
    "yrs <- params$metadata$census_years",
    "dead_total <- params$stats$dead_since_first",
    "recr_total <- params$stats$recruits_since_first",
    "if (!is.null(yrs) && length(yrs) > 1) {",
    "  if (!is.null(dead_total) && is.finite(dead_total))",
    "    cat(paste0('- **自第一次普查以来死亡的树木：** ', dead_total, '\\n'))",
    "  if (!is.null(recr_total) && is.finite(recr_total))",
    "    cat(paste0('- **自第一次普查以来新增个体：** ', recr_total, '\\n'))",
    "}",
    "```"
  )

  if (has_agb) {
    mid_section <- c(
      mid_section,
      "",
      "## 地上生物量 {#agb}",
      "",
      "```{r agb-table, echo=FALSE, results='asis'}",
      "if (!is.null(params$agb)) {",
      "  agb_val <- round(params$agb[[2]][1], 2)",
      "  cat(paste0(\"**地上生物量：** \", agb_val, \" t ha$^{-1}$\\n\\n\"))",
      "  cat('\\\\hrule\\n')",
      "}",
      "```"
    )
  }

  mid_section <- c(mid_section, "\\newpage")

  nav_targets <- c(
    "[返回目录](#contents)",
    "[元数据](#metadata)",
    "[个体统计](#counts)"
  )
  if (has_agb) nav_targets <- c(nav_targets, "[地上生物量](#agb)")
  nav_targets <- c(nav_targets, "[样地总体地图](#general-plot)")
  if (any(tf_col))   nav_targets <- c(nav_targets, "[仅已采集个体](#collected-only)")
  if (any(tf_uncol)) nav_targets <- c(nav_targets, "[未采集个体](#uncollected)")
  if (any(tf_palm))  nav_targets <- c(nav_targets, "[棕榈科](#uncollected-palm)")
  nav_targets <- c(nav_targets, "[子样地索引](#subplot-index)", "[物种清单](#checklist)")

  nav_links <- c(
    "\\begingroup\\tiny\\color{gray}",
    paste("«", paste(nav_targets, collapse = " | ")),
    "\\endgroup"
  )

  gencol_section <- c(
    "",
    "## 样地总体地图 {#general-plot}",
    "```{r general-plot, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$main_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  col_section <- c(
    "## 仅已采集个体 {#collected-only}",
    "```{r collected-only, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$collected_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  uncol_section <- c(
    "## 未采集个体 {#uncollected}",
    "```{r uncollected,fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$uncollected_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  palm_section <- c(
    "## 未采集棕榈科个体 {#uncollected-palm}",
    "```{r uncollected-palm, fig.width=12, fig.height=12, out.width='\\\\textwidth', fig.align='center'}",
    "print(params$uncollected_palm_plot)",
    "```",
    "",
    "",
    nav_links,
    "",
    "\\newpage"
  )

  index_section <- c(
    "",
    "\\newpage",
    "",
    "## 子样地索引 {#subplot-index}",
    "",
    "```{r toc, results='asis', echo=FALSE}",
    "cols <- 5",
    "n <- length(params$subplots_list)",
    "per_col <- ceiling(n / cols)",
    "header <- paste(rep('子样地', cols), collapse = ' | ')",
    "separator <- paste(rep('---', cols), collapse = ' | ')",
    "toc_lines <- c(header, separator)",
    "for (i in 1:per_col) {",
    "  row <- character(cols)",
    "  for (j in 0:(cols - 1)) {",
    "    idx <- i + j * per_col",
    "    if (idx <= n) {",
    "      row[j + 1] <- paste0('[子样地 ', idx, '](#subplot-', idx, ')')",
    "    } else {",
    "      row[j + 1] <- ' '",
    "    }",
    "  }",
    "  toc_lines <- c(toc_lines, paste(row, collapse = ' | '))",
    "}",
    "cat(paste(toc_lines, collapse = '\\n'))",
    "```",
    "",
    "",
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
      sprintf("### 子样地 %d {#subplot-%d .unlisted .unnumbered}", i, i),
      "",
      sprintf("```{r subplot-%d, fig.width=12, fig.height=9}", i),
      sprintf("print(params$subplots_list[[%d]]$plot)", i),
      "```",
      "",
      "\\begingroup\\tiny\\color{gray}",
      "« [返回目录](#contents) | [元数据](#metadata) | [个体统计](#counts) | [样地总体地图](#general-plot) | [仅已采集个体](#collected-only) | [未采集个体](#uncollected) | [棕榈科](#uncollected-palm) | [子样地索引](#subplot-index) | [物种清单](#checklist)",
      "\\endgroup",
      if (i < length(subplot_plots)) "\\newpage" else NULL,
      ""
    )
  }))

  checklist_section <- c(
    "",
    "\\newpage",
    "",
    "## 物种清单 {#checklist}",
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
    "      stringsAsFactors = FALSE)",
    "",
    "t2s_df <- params$tag_to_subplot",
    "is_monitora <- tolower(trimws(as.character(params$input_type))) == \"monitora\"",
    "",
    "if (is_monitora) {",
    "  tag_df <- merge(tag_df, t2s_df, by.x = \"tag\", by.y = \"New Tag No\", all.x = TRUE, sort = FALSE)",
    "  tag_df$target_idx <- ifelse(!is.na(tag_df$subplot_index), tag_df$subplot_index, tag_df$T1)",
    "  has_fields <- !is.na(tag_df$subunit_letter) & !is.na(tag_df$T2)",
    "  tag_df$subplot_code <- ifelse(has_fields, paste0(tag_df$subunit_letter, tag_df$T2), paste0(\"S\", tag_df$target_idx))",
    "  o <- order(tag_df$target_idx, suppressWarnings(as.numeric(tag_df$tag)))",
    "  tag_df <- tag_df[o, , drop = FALSE]",
    "  tag_links <- paste(sprintf(\"[ %s → %s ](#subplot-%s)\",",
    "                           tag_df$tag, tag_df$subplot_code, tag_df$target_idx),",
    "                     collapse = \" | \")",
    "} else {",
    "  tag_df <- merge(tag_df, t2s_df, by.x = \"tag\", by.y = \"New Tag No\", all.x = TRUE, sort = FALSE)",
    "  tag_df$target_idx <- tag_df$T1",
    "  o <- order(tag_df$target_idx, suppressWarnings(as.numeric(tag_df$tag)))",
    "  tag_df <- tag_df[o, , drop = FALSE]",
    "  tag_links <- paste(sprintf(\"[ %s - S%s ](#subplot-%s)\",",
    "                           tag_df$tag, tag_df$target_idx, tag_df$target_idx),",
    "                     collapse = \" | \")",
    "}",
    "",
    "cat(\"* \", sp, \": \", tag_links, \"\\n\", sep = \"\")",
    "}",
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

  add_body("## 单个子样地 {#individual-subplots}", subplot_sections, checklist_section)
  return(rmd_content)
}
