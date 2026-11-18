# Names of output files ####
# Authors: Giulia Ottino & Domingos Cardoso
#
# The one place where the names of the files written by forplot_balance() are
# defined. A file name is "<prefix>_<suffix>.<ext>", where the suffix comes
# from the table below in the language of the call. Names are snake_case ASCII
# (no accents, no spaces) so that they work in URLs. Panara ("pa") uses the
# Portuguese names; Mandarin ("ma") uses the English names, because Chinese
# characters cannot be written in ASCII.


#' Dictionary of output file names
#'
#' One row per key; one column per language with its own file names.
#'
#' @return Data frame with columns `key`, `en`, `pt`, `es`, `fr`.
#' @keywords internal
#' @noRd
.file_name_dict <- function() {
  rows <- list(
    #                              en                              pt                                     es                                     fr
    default_prefix        = c("plot",                         "parcela",                             "parcela",                             "parcelle"),
    station               = c("station",                      "estacao",                             "estacion",                            "station"),
    report                = c("full_report",                  "relatorio_completo",                  "informe_completo",                    "rapport_complet"),
    map_general           = c("general",                      "geral",                               "general",                             "general"),
    map_collected         = c("collected",                    "coletados",                           "colectados",                          "collectes"),
    map_uncollected       = c("uncollected",                  "nao_coletados",                       "no_colectados",                       "non_collectes"),
    map_uncollected_palms = c("uncollected_palms",            "palmeiras_nao_coletadas",             "palmeras_no_colectadas",              "palmiers_non_collectes"),
    map_priority          = c("priority_uncollected_species", "especies_prioritarias_nao_coletadas", "especies_prioritarias_no_colectadas", "especes_prioritaires_non_collectees"),
    chart_species_abund   = c("species_abundance",            "especies_abundancia",                 "especies_abundancia",                 "especes_abondance"),
    chart_species_ivi     = c("species_ivi",                  "especies_ivi",                        "especies_ivi",                        "especes_ivi"),
    chart_family_ivi      = c("families_ivi",                 "familias_ivi",                        "familias_ivi",                        "familles_ivi"),
    chart_dbh_classes     = c("dbh_classes",                  "classes_dap",                         "clases_dap",                          "classes_dhp"),
    balance_xlsx          = c("collection_balance",           "balanco_coleta",                      "balance_colecta",                     "bilan_collecte")
  )
  m <- do.call(rbind, rows)
  data.frame(key = rownames(m), en = m[, 1], pt = m[, 2], es = m[, 3], fr = m[, 4],
             row.names = NULL, stringsAsFactors = FALSE)
}


#' Language column of the file-name dictionary
#'
#' @keywords internal
#' @noRd
.file_name_lang <- function(language) {
  language <- tolower(trimws(as.character(language)[1]))
  switch(language, pt = "pt", pa = "pt", es = "es", fr = "fr", "en")
}


#' Output file name for a key in a language
#'
#' @param key Key of `.file_name_dict()`.
#' @param language Language code of the call ("en", "pt", "es", "fr", "ma", "pa").
#' @param prefix Optional prefix (the `filename` argument); joined with "_".
#' @param ext Optional extension without the dot.
#' @return Character scalar, e.g. `"RUS01_relatorio_completo.html"`.
#' @keywords internal
#' @noRd
.out_name <- function(key, language = "en", prefix = NULL, ext = NULL) {
  d <- .file_name_dict()
  i <- match(key, d$key)
  if (is.na(i)) stop("Unknown output file key: ", key, call. = FALSE)
  out <- d[[.file_name_lang(language)]][[i]]
  if (!is.null(prefix) && nzchar(prefix)) out <- paste0(prefix, "_", out)
  if (!is.null(ext) && nzchar(ext)) out <- paste0(out, ".", ext)
  out
}
