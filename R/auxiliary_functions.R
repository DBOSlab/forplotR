# Small functions to convert user input for standard data frame
# Author: Giulia Ottino & Domingos Cardoso
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
  
  # Helper to read either Excel or CSV using a flexible approach
  read_in <- function(f, sh = NULL) {
    if (grepl("\\.xlsx?$", f, ignore.case = TRUE)) {
      sheets <- readxl::excel_sheets(f)
      if (!is.null(sh) && sh %in% sheets) {
        df <- suppressMessages(
          readxl::read_excel(f, sheet = sh, .name_repair = "minimal")
        )
      } else {
        df <- NULL
        for (s in sheets) {
          temp <- suppressMessages(
            readxl::read_excel(f, sheet = s, .name_repair = "minimal")
          )
          if (ncol(temp) >= 3 && sum(!is.na(temp[1:10, ])) > 2) {
            df <- temp
            break
          }
        }
        if (is.null(df)) stop("No data found in any worksheet.", call. = FALSE)
      }
      return(df)
    } else {
      return(
        readr::read_delim(
          f,
          delim         = ",",
          locale        = locale,
          show_col_types = FALSE,
          .name_repair  = "minimal"
        )
      )
    }
  }
  
  # Read input
  in_dat <- suppressWarnings(read_in(path, sheet))
  colnames(in_dat) <- str_trim(colnames(in_dat))
  
  ## If already in the standard ForestPlots field-sheet format, use it directly
  if (all(dest_cols %in% names(in_dat))) {
    out <- in_dat
  } else {
    # If essential columns are not present, stop and ask the user to check the worksheet
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
  
  # Ensure all destination columns exist (safety check; normally no-op if we stopped above)
  for (col in dest_cols) {
    if (!col %in% names(out)) out[[col]] <- NA
  }
  out <- out[, dest_cols]
  
  # Small helper: first non-NA value in a set of possible metadata columns
  first_val <- function(df, cols) {
    for (cl in cols) {
      if (cl %in% names(df)) {
        v <- df[[cl]][1]
        return(v)
      }
    }
    NA_character_
  }
  
  # Build metadata row (first row)
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
  
  # Bind metadata row + header row + data
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
