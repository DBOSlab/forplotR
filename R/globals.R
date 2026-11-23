# Global variables and package-wide tables ####
# Authors: Giulia Ottino & Domingos Cardoso

utils::globalVariables(
  c(".ord",
    "D_cm",
    "X_loc",
    "abundance",
    "along_m",
    "collected_non_arecaceae",
    "count",
    "dbh_class",
    "family",
    "fill_group",
    "first_visit_order",
    "importance_value",
    "n_new_species",
    "n_species",
    "n_stems",
    "n_subplots",
    "n_trees",
    "new_species",
    "pct",
    "priority_subplots",
    "relative_density",
    "relative_frequency",
    "selected",
    "species",
    "species_set",
    "subplot_num",
    "subplots",
    "total",
    "visit_order",
    "New Stem Grouping",
    # From .build_fp_base_plot
    "global_x", "global_y", "center_x", "center_y", "T1", "diameter", "New Tag No",

    # From .build_monitora_base_plot
    "r", "xmin", "xmax", "ymin", "ymax", "col_from_center", "col_idx", "base_num",
    "cx", "cy", "label", "D", "lx", "ly", "lab", "hjust", "vjust",

    # From .clean_fp_data
    "Collected",

    # From .collection_percentual
    "Family", "T1", "n", "Collected", "collected", "total_individuals",
    "total_non_arecaceae", "subplot", "uncollected_non_arecaceae",
    "arecaceae_count", "New Tag No", "n_uncollected", "tagno_uncollected",
    "n_collected", "tagno_collected",

    # From .compute_global_coordinates
    "X", "Y", "global_x", "global_y",

    # From .compute_global_coordinates_monitora
    "arm", "Y_loc", "T2", "draw_x", "draw_y",

    # From .compute_plot_agb
    "PlotCode", "AGBind",

    # From .herbaria_lookup_links
    ":=", "rn", "recordNumber", "fp", "recordedBy", "ln", "key",

    # From fp_herb_converter
    "column_map",

    # From plot_for_balance
    "in_dat", "D", "Family", "Original determination", "Species", "Species_fmt",
    "n_subplots_per_col", "col_index", "row_index", "subplot_x", "subplot_y",
    "center_x", "center_y", "subunit_letter", "x10", "y10",
    "Status", "X", "Y", ".monitora_to_cell_coords",

    # From plot_html_map
    "latitude", "longitude",

    # From .calculate_phytosociological_metrics
    "absolute_frequency", "relative_dominance", "basal_area_m2",
    "Vernacular", "component", "value",

    NULL
  )
)


#===============================================================================
# Names of output files
#===============================================================================

# Names of output files ####
# Authors: Giulia Ottino & Domingos Cardoso
#
# The one place where the names of the files written by the package are
# defined (forplot_balance(), forplot_map(), forplot_vouchers() and
# forplot_sheets()). A file name is "<prefix>_<suffix>.<ext>", where the suffix comes
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
    balance_xlsx          = c("collection_balance",           "balanco_coleta",                      "balance_colecta",                     "bilan_collecte"),
    # forplot_map()
    results               = c("results",                      "resultados",                          "resultados",                          "resultats"),
    map_default           = c("plot_map",                     "mapa_parcela",                        "mapa_parcela",                        "carte_parcelle"),
    map_vouchers          = c("vouchers",                     "testemunhos",                         "testigos",                            "temoins"),
    # forplot_vouchers(); "tree" names the folder "<Voucher> <tree>#<tag>"
    voucher_imgs          = c("voucher_imgs",                 "imagens_testemunhos",                 "imagenes_testigos",                   "images_temoins"),
    tree                  = c("Tree",                         "Arvore",                              "Arbol",                               "Arbre"),
    # forplot_sheets()
    herb_dir              = c("herbarium_labels",             "etiquetas_herbario",                  "etiquetas_herbario",                  "etiquettes_herbier"),
    herb_sheet            = c("plot_to_herb_sheet",           "planilha_herbario",                   "planilla_herbario",                   "feuille_herbier")
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


#' The word "Tree" of voucher folders in every language
#'
#' `forplot_vouchers()` names the folder of a collected tree
#' "<Voucher> <Tree>#<tag>" in the language of the call; `forplot_map()`
#' accepts any of them when it links the images.
#'
#' @return Character vector, English first.
#' @keywords internal
#' @noRd
.tree_words <- function() {
  d <- .file_name_dict()
  unique(unlist(d[d$key == "tree", c("en", "pt", "es", "fr")], use.names = FALSE))
}
