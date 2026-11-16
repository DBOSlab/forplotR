# Field report for forplot_balance() ####
# Authors: Giulia Ottino & Domingos Cardoso
#
# A single self-contained HTML file meant for the field: a summary of the
# collection balance, a clickable plot grid, labelled whole-plot maps, a
# priority route, one A4 page per subplot (map, tree table and blank columns
# for remeasurement) and a checklist appendix. Maps are inline SVG drawn
# directly from the data, so the file is small, opens offline and prints the
# same way it looks on screen. The PDF is the same HTML printed by a headless
# Chrome, Chromium or Edge (see .html_to_pdf()); no pandoc or LaTeX is needed.


# Strings ----------------------------------------------------------------------

#' Report strings in one language
#'
#' Panara ("pa") has no translation of these strings yet and falls back to
#' Portuguese.
#'
#' @param language One of "en", "pt", "es", "fr", "ma", "pa".
#' @return Named character vector.
#' @keywords internal
#' @noRd
.field_report_strings <- function(language = "en") {
  s <- list(
    doc_title       = c("Collection balance", "Balan\u00e7o de coleta", "Balance de colecta", "Bilan de collecte", "\u91c7\u96c6\u5e73\u8861"),
    census          = c("census", "censo", "censo", "recensement", "\u666e\u67e5"),
    team            = c("Team", "Equipe", "Equipo", "\u00c9quipe", "\u56e2\u961f"),
    subplots_of     = c("%d subplots of %s \u00d7 %s m", "%d subparcelas de %s \u00d7 %s m", "%d subparcelas de %s \u00d7 %s m", "%d sous-parcelles de %s \u00d7 %s m", "%d \u4e2a %s \u00d7 %s \u7c73\u5b50\u6837\u5730"),
    individuals     = c("Individuals", "Indiv\u00edduos", "Individuos", "Individus", "\u4e2a\u4f53"),
    in_subplots     = c("in %d subplots", "em %d subparcelas", "en %d subparcelas", "dans %d sous-parcelles", "\u5206\u5e03\u4e8e %d \u4e2a\u5b50\u6837\u5730"),
    families        = c("Families", "Fam\u00edlias", "Familias", "Familles", "\u79d1"),
    identified      = c("identified", "identificadas", "identificadas", "identifi\u00e9es", "\u5df2\u9274\u5b9a"),
    with_voucher    = c("With voucher", "Com voucher", "Con voucher", "Avec voucher", "\u5df2\u6709\u51ed\u8bc1\u6807\u672c"),
    pct_nonpalms    = c("%s%% of non-palms", "%s%% das n\u00e3o palmeiras", "%s%% de las no palmeras", "%s%% des non-palmiers", "\u5360\u975e\u68d5\u6988\u4e2a\u4f53\u7684 %s%%"),
    pct_all         = c("%s%% of individuals", "%s%% dos indiv\u00edduos", "%s%% de los individuos", "%s%% des individus", "\u5360\u5168\u90e8\u4e2a\u4f53\u7684 %s%%"),
    pending         = c("Pending collection", "Pendentes de coleta", "Pendientes de colecta", "\u00c0 collecter", "\u5f85\u91c7\u96c6"),
    pending_ctx_np  = c("non-palms without voucher", "n\u00e3o palmeiras sem voucher", "no palmeras sin voucher", "non-palmiers sans voucher", "\u65e0\u51ed\u8bc1\u7684\u975e\u68d5\u6988\u4e2a\u4f53"),
    pending_ctx     = c("without voucher", "sem voucher", "sin voucher", "sans voucher", "\u65e0\u51ed\u8bc1\u6807\u672c"),
    palms           = c("Palms", "Palmeiras", "Palmeras", "Palmiers", "\u68d5\u6988"),
    counted_apart   = c("(counted apart)", "(contadas \u00e0 parte)", "(contadas aparte)", "(compt\u00e9s \u00e0 part)", "\uff08\u5355\u72ec\u8ba1\u6570\uff09"),
    subplots_visit  = c("Subplots to visit", "Subparcelas a visitar", "Subparcelas por visitar", "Sous-parcelles \u00e0 visiter", "\u5f85\u8bbf\u95ee\u5b50\u6837\u5730"),
    with_pending    = c("with \u2265 1 pending", "com \u2265 1 pendente", "con \u2265 1 pendiente", "avec \u2265 1 \u00e0 collecter", "\u81f3\u5c11 1 \u682a\u5f85\u91c7\u96c6"),
    fn_indet        = c("%d individuals without a family (Indet) are not counted as a family.",
                        "%d indiv\u00edduos sem fam\u00edlia (Indet) n\u00e3o s\u00e3o contados como fam\u00edlia.",
                        "%d individuos sin familia (Indet) no se cuentan como familia.",
                        "%d individus sans famille (Indet) ne sont pas compt\u00e9s comme famille.",
                        "%d \u4e2a\u672a\u5b9a\u79d1\uff08Indet\uff09\u4e2a\u4f53\u4e0d\u8ba1\u4e3a\u4e00\u4e2a\u79d1\u3002"),
    fn_voucher_np   = c("\u201cWith voucher\u201d means the <i>Collected</i> column is filled in. Denominator: %d individuals that are not palms (%d with voucher + %d pending).",
                        "\u201cCom voucher\u201d significa que a coluna <i>Collected</i> est\u00e1 preenchida. Denominador: %d indiv\u00edduos que n\u00e3o s\u00e3o palmeiras (%d com voucher + %d pendentes).",
                        "\u201cCon voucher\u201d significa que la columna <i>Collected</i> est\u00e1 rellenada. Denominador: %d individuos que no son palmeras (%d con voucher + %d pendientes).",
                        "\u00ab Avec voucher \u00bb signifie que la colonne <i>Collected</i> est remplie. D\u00e9nominateur : %d individus qui ne sont pas des palmiers (%d avec voucher + %d \u00e0 collecter).",
                        "\u201c\u5df2\u6709\u51ed\u8bc1\u6807\u672c\u201d\u6307 <i>Collected</i> \u5217\u5df2\u586b\u5199\u3002\u5206\u6bcd\uff1a%d \u4e2a\u975e\u68d5\u6988\u4e2a\u4f53\uff08%d \u4e2a\u5df2\u6709\u51ed\u8bc1 + %d \u4e2a\u5f85\u91c7\u96c6\uff09\u3002"),
    fn_voucher_all  = c("\u201cWith voucher\u201d means the <i>Collected</i> column is filled in. Denominator: all %d individuals (%d with voucher + %d pending).",
                        "\u201cCom voucher\u201d significa que a coluna <i>Collected</i> est\u00e1 preenchida. Denominador: todos os %d indiv\u00edduos (%d com voucher + %d pendentes).",
                        "\u201cCon voucher\u201d significa que la columna <i>Collected</i> est\u00e1 rellenada. Denominador: los %d individuos (%d con voucher + %d pendientes).",
                        "\u00ab Avec voucher \u00bb signifie que la colonne <i>Collected</i> est remplie. D\u00e9nominateur : les %d individus (%d avec voucher + %d \u00e0 collecter).",
                        "\u201c\u5df2\u6709\u51ed\u8bc1\u6807\u672c\u201d\u6307 <i>Collected</i> \u5217\u5df2\u586b\u5199\u3002\u5206\u6bcd\uff1a\u5168\u90e8 %d \u4e2a\u4e2a\u4f53\uff08%d \u4e2a\u5df2\u6709\u51ed\u8bc1 + %d \u4e2a\u5f85\u91c7\u96c6\uff09\u3002"),
    fn_palms        = c("Palms (Arecaceae) are left out of the voucher balance because they are usually not collected as herbarium vouchers; %d + %d + %d = %d individuals.",
                        "Palmeiras (Arecaceae) ficam fora do balan\u00e7o de vouchers porque em geral n\u00e3o s\u00e3o coletadas como exsicatas; %d + %d + %d = %d indiv\u00edduos.",
                        "Las palmeras (Arecaceae) quedan fuera del balance de vouchers porque en general no se colectan como exsicatas; %d + %d + %d = %d individuos.",
                        "Les palmiers (Arecaceae) sont exclus du bilan des vouchers car ils sont rarement r\u00e9colt\u00e9s en herbier ; %d + %d + %d = %d individus.",
                        "\u68d5\u6988\u79d1\uff08Arecaceae\uff09\u901a\u5e38\u4e0d\u91c7\u96c6\u4e3a\u814a\u53f6\u6807\u672c\uff0c\u56e0\u6b64\u4e0d\u8ba1\u5165\u51ed\u8bc1\u5e73\u8861\uff1b%d + %d + %d = %d \u4e2a\u4e2a\u4f53\u3002"),
    grid_cap        = c("Subplots \u00b7 tap to open", "Subparcelas \u00b7 toque para abrir", "Subparcelas \u00b7 toque para abrir", "Sous-parcelles \u00b7 touchez pour ouvrir", "\u5b50\u6837\u5730 \u00b7 \u70b9\u51fb\u6253\u5f00"),
    leg_many        = c("3+ pending", "3+ pendentes", "3+ pendientes", "3+ \u00e0 collecter", "3 \u682a\u4ee5\u4e0a\u5f85\u91c7\u96c6"),
    leg_done        = c("done", "conclu\u00edda", "completa", "termin\u00e9e", "\u5df2\u5b8c\u6210"),
    leg_empty       = c("no stems", "sem fustes", "sin fustes", "sans tiges", "\u65e0\u830e\u5e72"),
    maps_cap        = c("Plot maps \u00b7 every stem with its tag", "Mapas da parcela \u00b7 cada fuste com a plaqueta", "Mapas de la parcela \u00b7 cada fuste con su placa", "Cartes de la parcelle \u00b7 chaque tige avec sa plaque", "\u6837\u5730\u56fe \u00b7 \u6bcf\u4e2a\u830e\u5e72\u53ca\u5176\u6807\u7b7e"),
    map_all         = c("All stems", "Todos os fustes", "Todos los fustes", "Toutes les tiges", "\u5168\u90e8\u830e\u5e72"),
    plot_map        = c("plot map", "mapa da parcela", "mapa de la parcela", "carte de la parcelle", "\u6837\u5730\u56fe"),
    stems           = c("stems", "fustes", "fustes", "tiges", "\u4e2a\u830e\u5e72"),
    lede_all        = c("Every stem in the plot with its tag number.", "Todos os fustes da parcela com o n\u00famero da plaqueta.", "Todos los fustes de la parcela con el n\u00famero de placa.", "Toutes les tiges de la parcelle avec leur num\u00e9ro de plaque.", "\u6837\u5730\u5185\u6240\u6709\u830e\u5e72\u53ca\u5176\u6807\u7b7e\u53f7\u3002"),
    lede_ok         = c("Only stems already collected (<i>Collected</i> column filled in).", "S\u00f3 os fustes j\u00e1 coletados (coluna <i>Collected</i> preenchida).", "Solo los fustes ya colectados (columna <i>Collected</i> rellenada).", "Uniquement les tiges d\u00e9j\u00e0 r\u00e9colt\u00e9es (colonne <i>Collected</i> remplie).", "\u4ec5\u663e\u793a\u5df2\u91c7\u96c6\u7684\u830e\u5e72\uff08<i>Collected</i> \u5217\u5df2\u586b\u5199\uff09\u3002"),
    lede_todo       = c("Only stems without a voucher.", "S\u00f3 os fustes sem voucher.", "Solo los fustes sin voucher.", "Uniquement les tiges sans voucher.", "\u4ec5\u663e\u793a\u65e0\u51ed\u8bc1\u6807\u672c\u7684\u830e\u5e72\u3002"),
    lede_todo_palms = c(" Palms are shown on their own map.", " As palmeiras t\u00eam mapa pr\u00f3prio.", " Las palmeras tienen su propio mapa.", " Les palmiers ont leur propre carte.", "\u68d5\u6988\u53e6\u89c1\u5355\u72ec\u5730\u56fe\u3002"),
    lede_palm       = c("Only palms (Arecaceae), counted apart from the voucher balance.", "S\u00f3 as palmeiras (Arecaceae), contadas \u00e0 parte do balan\u00e7o de vouchers.", "Solo las palmeras (Arecaceae), contadas aparte del balance de vouchers.", "Uniquement les palmiers (Arecaceae), compt\u00e9s \u00e0 part du bilan des vouchers.", "\u4ec5\u663e\u793a\u68d5\u6988\u79d1\uff08Arecaceae\uff09\uff0c\u4e0d\u8ba1\u5165\u51ed\u8bc1\u5e73\u8861\u3002"),
    tap_subplot     = c(" Tap a subplot to open its page.", " Toque numa subparcela para abrir a p\u00e1gina dela.", " Toque una subparcela para abrir su p\u00e1gina.", " Touchez une sous-parcelle pour ouvrir sa page.", "\u70b9\u51fb\u5b50\u6837\u5730\u53ef\u6253\u5f00\u5176\u9875\u9762\u3002"),
    leg_pending     = c("pending", "pendente", "pendiente", "\u00e0 collecter", "\u5f85\u91c7\u96c6"),
    leg_voucher     = c("with voucher", "com voucher", "con voucher", "avec voucher", "\u5df2\u6709\u51ed\u8bc1"),
    leg_palm        = c("palm", "palmeira", "palmera", "palmier", "\u68d5\u6988"),
    leg_size        = c("size = DBH", "tamanho = DAP", "tama\u00f1o = DAP", "taille = DHP", "\u5927\u5c0f = \u80f8\u5f84"),
    leg_tag         = c("numbers = stem tag", "n\u00fameros = plaqueta", "n\u00fameros = placa", "nombres = plaque", "\u6570\u5b57 = \u6807\u7b7e\u53f7"),
    route           = c("Priority route", "Rota de prioridade", "Ruta de prioridad", "Itin\u00e9raire prioritaire", "\u4f18\u5148\u8def\u7ebf"),
    route_lede      = c("Subplots with stems still missing a voucher, most first.", "Subparcelas com fustes ainda sem voucher, da que tem mais para a que tem menos.", "Subparcelas con fustes a\u00fan sin voucher, de m\u00e1s a menos.", "Sous-parcelles avec des tiges encore sans voucher, de la plus \u00e0 la moins charg\u00e9e.", "\u4ecd\u6709\u830e\u5e72\u7f3a\u5c11\u51ed\u8bc1\u6807\u672c\u7684\u5b50\u6837\u5730\uff0c\u6309\u6570\u91cf\u4ece\u591a\u5230\u5c11\u6392\u5217\u3002"),
    route_palms     = c(" Palms are counted apart.", " Palmeiras contadas \u00e0 parte.", " Palmeras contadas aparte.", " Palmiers compt\u00e9s \u00e0 part.", "\u68d5\u6988\u5355\u72ec\u8ba1\u6570\u3002"),
    n_pending       = c("%d pending", "%d pendentes", "%d pendientes", "%d \u00e0 collecter", "%d \u682a\u5f85\u91c7\u96c6"),
    nothing_pending = c("No stem is pending collection.", "Nenhum fuste pendente de coleta.", "Ning\u00fan fuste pendiente de colecta.", "Aucune tige \u00e0 collecter.", "\u6ca1\u6709\u5f85\u91c7\u96c6\u7684\u830e\u5e72\u3002"),
    spp_pending     = c("Species to collect", "Esp\u00e9cies a coletar", "Especies por colectar", "Esp\u00e8ces \u00e0 collecter", "\u5f85\u91c7\u96c6\u7269\u79cd"),
    family          = c("Family", "Fam\u00edlia", "Familia", "Famille", "\u79d1"),
    species         = c("Species", "Esp\u00e9cie", "Especie", "Esp\u00e8ce", "\u7269\u79cd"),
    n_stems         = c("Stems", "Fustes", "Fustes", "Tiges", "\u830e\u5e72\u6570"),
    tags_sub        = c("Tags \u00b7 subplot", "Plaquetas \u00b7 subparcela", "Placas \u00b7 subparcela", "Plaques \u00b7 sous-parcelle", "\u6807\u7b7e \u00b7 \u5b50\u6837\u5730"),
    subplot         = c("Subplot", "Subparcela", "Subparcela", "Sous-parcelle", "\u5b50\u6837\u5730"),
    index           = c("Index", "\u00cdndice", "\u00cdndice", "Index", "\u7d22\u5f15"),
    prev            = c("Previous subplot", "Subparcela anterior", "Subparcela anterior", "Sous-parcelle pr\u00e9c\u00e9dente", "\u4e0a\u4e00\u4e2a\u5b50\u6837\u5730"),
    nxt             = c("Next subplot", "Pr\u00f3xima subparcela", "Siguiente subparcela", "Sous-parcelle suivante", "\u4e0b\u4e00\u4e2a\u5b50\u6837\u5730"),
    back_grid       = c("Back to the plot grid", "Voltar \u00e0 grade da parcela", "Volver a la cuadr\u00edcula", "Retour \u00e0 la grille", "\u8fd4\u56de\u6837\u5730\u7f51\u683c"),
    col_row         = c("column %d \u00b7 row %d", "coluna %d \u00b7 linha %d", "columna %d \u00b7 fila %d", "colonne %d \u00b7 ligne %d", "\u7b2c %d \u5217 \u00b7 \u7b2c %d \u884c"),
    arm_sub         = c("arm %s \u00b7 subplot %d", "bra\u00e7o %s \u00b7 subparcela %d", "brazo %s \u00b7 subparcela %d", "bras %s \u00b7 sous-parcelle %d", "%s \u81c2 \u00b7 \u7b2c %d \u5b50\u6837\u5730"),
    pending_short   = c("Pending", "Pendentes", "Pendientes", "\u00c0 collecter", "\u5f85\u91c7\u96c6"),
    no_pending_sub  = c("No pending collection in this subplot.", "Nenhuma coleta pendente nesta subparcela.", "Ninguna colecta pendiente en esta subparcela.", "Rien \u00e0 collecter dans cette sous-parcelle.", "\u6b64\u5b50\u6837\u5730\u6ca1\u6709\u5f85\u91c7\u96c6\u7684\u830e\u5e72\u3002"),
    tag             = c("Tag", "Plaqueta", "Placa", "Plaque", "\u6807\u7b7e"),
    dbh             = c("DBH", "DAP", "DAP", "DHP", "\u80f8\u5f84"),
    new_dbh         = c("New DBH", "Novo DAP", "Nuevo DAP", "Nouveau DHP", "\u65b0\u80f8\u5f84"),
    st              = c("St", "St", "Est", "St", "\u72b6\u6001"),
    coll_no         = c("\u2713 / no.", "\u2713 / n\u00ba", "\u2713 / n\u00ba", "\u2713 / n\u00ba", "\u2713 / \u7f16\u53f7"),
    notes           = c("Notes", "Observa\u00e7\u00f5es", "Notas", "Notes", "\u5907\u6ce8"),
    st_legend       = c("St: A alive \u00b7 D dead \u00b7 R recruit \u00b7 N not found", "St: V viva \u00b7 M morta \u00b7 R recruta \u00b7 N n\u00e3o encontrada", "Est: V viva \u00b7 M muerta \u00b7 R recluta \u00b7 N no encontrada", "St : V vivant \u00b7 M mort \u00b7 R recrue \u00b7 N non trouv\u00e9", "\u72b6\u6001\uff1aA \u5b58\u6d3b \u00b7 D \u6b7b\u4ea1 \u00b7 R \u65b0\u589e \u00b7 N \u672a\u627e\u5230"),
    no_stems        = c("No stems recorded in this subplot.", "Nenhum fuste registrado nesta subparcela.", "Ning\u00fan fuste registrado en esta subparcela.", "Aucune tige enregistr\u00e9e dans cette sous-parcelle.", "\u6b64\u5b50\u6837\u5730\u65e0\u830e\u5e72\u8bb0\u5f55\u3002"),
    appendix        = c("Appendix \u00b7 checklist", "Ap\u00eandice \u00b7 checklist", "Ap\u00e9ndice \u00b7 lista de especies", "Annexe \u00b7 liste des esp\u00e8ces", "\u9644\u5f55 \u00b7 \u6e05\u5355"),
    chk_lede        = c("Family \u2192 species \u2192 tags. Tap a tag to open its row on the subplot page.", "Fam\u00edlia \u2192 esp\u00e9cie \u2192 plaquetas. Toque numa plaqueta para abrir a linha dela na p\u00e1gina da subparcela.", "Familia \u2192 especie \u2192 placas. Toque una placa para abrir su fila en la p\u00e1gina de la subparcela.", "Famille \u2192 esp\u00e8ce \u2192 plaques. Touchez une plaque pour ouvrir sa ligne sur la page de la sous-parcelle.", "\u79d1 \u2192 \u7269\u79cd \u2192 \u6807\u7b7e\u3002\u70b9\u51fb\u6807\u7b7e\u53ef\u6253\u5f00\u5176\u6240\u5728\u5b50\u6837\u5730\u9875\u9762\u4e2d\u7684\u884c\u3002"),
    find_ph         = c("Tag or %s", "Plaqueta ou %s", "Placa o %s", "Plaque ou %s", "\u6807\u7b7e\u6216 %s"),
    find_label      = c("Find a tag or subplot", "Buscar plaqueta ou subparcela", "Buscar placa o subparcela", "Chercher une plaque ou une sous-parcelle", "\u67e5\u627e\u6807\u7b7e\u6216\u5b50\u6837\u5730"),
    not_found       = c("Tag not found", "Plaqueta n\u00e3o encontrada", "Placa no encontrada", "Plaque introuvable", "\u672a\u627e\u5230\u6807\u7b7e"),
    maps            = c("Maps", "Mapas", "Mapas", "Cartes", "\u5730\u56fe"),
    route_short     = c("Route", "Rota", "Ruta", "Itin\u00e9raire", "\u8def\u7ebf"),
    checklist       = c("Checklist", "Checklist", "Especies", "Esp\u00e8ces", "\u6e05\u5355"),
    arms_note       = c("Arms are drawn closer to the centre than in the field; distances within each arm are to scale.",
                        "Os bra\u00e7os aparecem mais pr\u00f3ximos do centro do que no campo; dentro de cada bra\u00e7o as dist\u00e2ncias est\u00e3o em escala.",
                        "Los brazos se dibujan m\u00e1s cerca del centro que en el campo; dentro de cada brazo las distancias est\u00e1n a escala.",
                        "Les bras sont dessin\u00e9s plus pr\u00e8s du centre que sur le terrain ; \u00e0 l'int\u00e9rieur de chaque bras, les distances sont \u00e0 l'\u00e9chelle.",
                        "\u5404\u81c2\u7ed8\u5236\u5f97\u6bd4\u5b9e\u5730\u66f4\u9760\u8fd1\u4e2d\u5fc3\uff1b\u6bcf\u4e2a\u81c2\u5185\u7684\u8ddd\u79bb\u6309\u6bd4\u4f8b\u7ed8\u5236\u3002"),
    print_pdf       = c("Print or save as PDF", "Imprimir ou salvar em PDF", "Imprimir o guardar como PDF", "Imprimer ou enregistrer en PDF", "\u6253\u5370\u6216\u4fdd\u5b58\u4e3a PDF")
  )
  langs <- c("en", "pt", "es", "fr", "ma")
  language <- tolower(trimws(as.character(language)[1]))
  if (identical(language, "pa")) language <- "pt"
  if (!language %in% langs) language <- "en"
  vapply(s, function(v) v[[match(language, langs)]], character(1))
}


# Helpers ----------------------------------------------------------------------

.fr_esc <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- ""
  x <- gsub("&", "&amp;", x, fixed = TRUE)
  x <- gsub("<", "&lt;", x, fixed = TRUE)
  x <- gsub(">", "&gt;", x, fixed = TRUE)
  gsub("\"", "&quot;", x, fixed = TRUE)
}

.fr_num <- function(x) suppressWarnings(as.numeric(as.character(x)))

.fr_fmt <- function(x, digits = 1) {
  ifelse(is.finite(x), formatC(x, format = "f", digits = digits), "\u2013")
}

# Plain coordinate as typed (2, 7.5) without trailing zeros.
.fr_coord <- function(x) {
  x <- .fr_num(x)
  ifelse(is.finite(x), format(round(x, 2), trim = TRUE, drop0trailing = TRUE), "\u2013")
}

.fr_read_asset <- function(name) {
  f <- system.file("report", name, package = "forplotR")
  if (!nzchar(f) || !file.exists(f)) {
    stop("Report asset not found: ", name, call. = FALSE)
  }
  paste(readLines(f, warn = FALSE, encoding = "UTF-8"), collapse = "\n")
}


# Layout -----------------------------------------------------------------------

#' Geometry of the report: subplot cells and stems in drawing coordinates
#'
#' Cells are rectangles in a common coordinate system (metres, y pointing up).
#' Stems carry local coordinates inside their subplot (0..cell size) for the
#' subplot maps and global coordinates for the whole-plot maps.
#'
#' @param fp_coords Data frame returned by `.prepare_balance_data()`.
#' @return List with `cells`, `trees`, `cell_size`, `extent`, `kind`.
#' @keywords internal
#' @noRd
.field_report_layout <- function(fp_coords,
                                 input_type,
                                 subplot_size = 10,
                                 plot_width_m = NULL,
                                 plot_length_m = NULL,
                                 highlight_palms = TRUE) {
  d <- as.data.frame(fp_coords, stringsAsFactors = FALSE)

  get_col <- function(nm) if (nm %in% names(d)) d[[nm]] else rep(NA, nrow(d))
  sp <- as.character(get_col("Original determination"))
  if (all(is.na(sp)) && "Species" %in% names(d)) sp <- as.character(d$Species)
  fam <- trimws(as.character(get_col("Family")))
  fam[is.na(fam) | !nzchar(fam)] <- "Indet"
  sp <- trimws(sp)
  sp[is.na(sp) | !nzchar(sp)] <- "Indet indet"
  coll <- trimws(as.character(get_col("Collected")))
  st <- ifelse(isTRUE(highlight_palms) & fam == "Arecaceae", "palm",
               ifelse(!is.na(coll) & nzchar(coll), "ok", "todo"))
  tag <- trimws(as.character(get_col("New Tag No")))
  tag[is.na(tag)] <- ""

  trees <- data.frame(
    tag = tag,
    fam = fam,
    spp = sp,
    dbh = .fr_num(get_col("D")) / 10,
    st = st,
    x_raw = get_col("X"),
    y_raw = get_col("Y"),
    stringsAsFactors = FALSE
  )

  if (identical(input_type, "monitora")) {
    arms <- c("N", "S", "L", "O")
    cells <- expand.grid(T2 = 1:10, arm = arms, stringsAsFactors = FALSE)[, c("arm", "T2")]
    col_idx <- ceiling(cells$T2 / 2)
    off <- cells$T2 - ((col_idx - 1L) * 2L + 1L)
    row_in_col <- ifelse(col_idx %% 2L == 0L, 1L - off, off)
    a0 <- (col_idx - 1) * 10
    c0 <- ifelse(row_in_col == 0L, -10, 0)
    cells$x0 <- ifelse(cells$arm %in% c("N", "S"), c0, ifelse(cells$arm == "L", 50 + a0, -100 + a0))
    cells$y0 <- ifelse(cells$arm == "N", 50 + a0, ifelse(cells$arm == "S", -100 + a0, c0))
    cells$x1 <- cells$x0 + 10
    cells$y1 <- cells$y0 + 10
    # Schematic position for the clickable grids: arms pulled towards the
    # centre (gap of 12 m instead of 50 m) so the cells stay large enough to tap.
    shift <- 38
    cells$sx0 <- cells$x0 + ifelse(cells$arm == "L", -shift, ifelse(cells$arm == "O", shift, 0))
    cells$sy0 <- cells$y0 + ifelse(cells$arm == "N", -shift, ifelse(cells$arm == "S", shift, 0))
    cells$id <- paste0(cells$arm, cells$T2)
    cells$label <- cells$id
    cells$where_a <- cells$arm
    cells$where_b <- cells$T2
    cell_size <- 10
    extent <- c(-100, -100, 100, 100)
    grid_extent <- c(-62, -62, 62, 62)

    trees$cell <- paste0(toupper(trimws(as.character(get_col("T1")))), .fr_num(get_col("T2")))
    trees$lx <- .fr_num(get_col("x10"))
    trees$ly <- .fr_num(get_col("y10"))
    trees$gx <- .fr_num(get_col("draw_x"))
    trees$gy <- .fr_num(get_col("draw_y"))
    arm_t <- toupper(trimws(as.character(get_col("T1"))))
    trees$sgx <- trees$gx + ifelse(arm_t == "L", -shift, ifelse(arm_t == "O", shift, 0))
    trees$sgy <- trees$gy + ifelse(arm_t == "N", -shift, ifelse(arm_t == "S", shift, 0))
    kind <- "monitora"
  } else {
    S <- subplot_size
    nr <- as.integer(round(plot_length_m / S))
    nc <- as.integer(round(plot_width_m / S))
    t1 <- seq_len(nr * nc)
    col <- floor((t1 - 1) / nr)
    row <- (t1 - 1) %% nr
    gy <- ifelse(col %% 2 == 0, row, nr - row - 1)
    cells <- data.frame(id = as.character(t1), label = as.character(t1),
                        x0 = col * S, y0 = gy * S, x1 = col * S + S, y1 = gy * S + S,
                        where_a = col + 1, where_b = gy + 1, stringsAsFactors = FALSE)
    cell_size <- S
    extent <- c(0, 0, nc * S, nr * S)
    grid_extent <- extent

    trees$cell <- as.character(.fr_num(get_col("T1")))
    trees$lx <- .fr_num(get_col("X"))
    trees$ly <- .fr_num(get_col("Y"))
    trees$gx <- .fr_num(get_col("global_x"))
    trees$gy <- .fr_num(get_col("global_y"))
    kind <- "grid"
  }

  cells$order <- seq_len(nrow(cells))
  trees <- trees[trees$cell %in% cells$id, , drop = FALSE]
  tn <- .fr_num(trees$tag)
  trees <- trees[order(is.na(tn), tn, trees$tag), , drop = FALSE]
  trees$row_id <- paste0("r", seq_len(nrow(trees)))

  cnt <- function(cond) vapply(cells$id, function(i) sum(trees$cell == i & cond), numeric(1))
  cells$n <- cnt(rep(TRUE, nrow(trees)))
  cells$todo <- cnt(trees$st == "todo")
  cells$palm <- cnt(trees$st == "palm")

  if (is.null(cells$sx0)) { cells$sx0 <- cells$x0; cells$sy0 <- cells$y0 }
  if (is.null(trees$sgx)) { trees$sgx <- trees$gx; trees$sgy <- trees$gy }
  list(cells = cells, trees = trees, cell_size = cell_size, extent = extent,
       grid_extent = grid_extent, kind = kind)
}


# SVG --------------------------------------------------------------------------

.fr_cell_class <- function(n, todo) {
  ifelse(n == 0, "c-empty", ifelse(todo == 0, "c-done", ifelse(todo <= 2, "c-few", "c-many")))
}

# Scale cells of the layout into a box of `size` units (y flipped). The
# clickable grids use the schematic position (`schematic = TRUE`); the
# whole-plot maps use the real one.
.fr_cell_boxes <- function(lay, size, pad = 0, schematic = TRUE) {
  e <- if (schematic) lay$grid_extent else lay$extent
  k <- (size - 2 * pad) / max(e[3] - e[1], e[4] - e[2])
  c <- lay$cells
  x0 <- if (schematic) c$sx0 else c$x0
  y0 <- if (schematic) c$sy0 else c$y0
  cw <- c$x1 - c$x0; ch <- c$y1 - c$y0
  list(k = k,
       w = (e[3] - e[1]) * k + 2 * pad, h = (e[4] - e[2]) * k + 2 * pad,
       x = pad + (x0 - e[1]) * k, y = pad + (e[4] - (y0 + ch)) * k,
       cw = cw * k, ch = ch * k)
}

.fr_grid_svg <- function(lay, size = 300, labels = TRUE) {
  b <- .fr_cell_boxes(lay, size)
  c <- lay$cells
  fs <- max(5, min(b$cw[1], b$ch[1]) * 0.27)
  cls <- .fr_cell_class(c$n, c$todo)
  lab <- if (labels) sprintf('<text x="%.1f" y="%.1f" style="font-size:%.1fpx">%s</text>',
                             b$x + b$cw / 2, b$y + b$ch / 2 + fs * 0.36, fs, .fr_esc(c$label)) else ""
  cells <- sprintf('<a href="#s-%s" aria-label="%s: %d"><rect class="%s" x="%.1f" y="%.1f" width="%.1f" height="%.1f"/>%s</a>',
                   c$id, .fr_esc(c$label), c$todo, cls, b$x + .5, b$y + .5, b$cw - 1, b$ch - 1, lab)
  sprintf('<svg class="grid" viewBox="0 0 %.1f %.1f" role="img">%s</svg>', b$w, b$h, paste(cells, collapse = ""))
}

# Mini grid drawn once as a <symbol>; each subplot page reuses it and adds a frame.
.fr_mini_symbol <- function(lay, size = 120) {
  b <- .fr_cell_boxes(lay, size)
  c <- lay$cells
  rects <- sprintf('<rect class="%s" x="%.1f" y="%.1f" width="%.1f" height="%.1f"/>',
                   .fr_cell_class(c$n, c$todo), b$x + .4, b$y + .4, b$cw - .8, b$ch - .8)
  sprintf('<svg width="0" height="0" style="position:absolute" aria-hidden="true"><symbol id="g-mini" viewBox="0 0 %.1f %.1f">%s</symbol></svg>',
          b$w, b$h, paste(rects, collapse = ""))
}

.fr_mini_svg <- function(lay, i, back_label, size = 120) {
  b <- .fr_cell_boxes(lay, size)
  sprintf('<a href="#summary" aria-label="%s"><svg class="grid mini" viewBox="0 0 %.1f %.1f"><use href="#g-mini"/><rect class="cur" x="%.1f" y="%.1f" width="%.1f" height="%.1f"/></svg></a>',
          .fr_esc(back_label), b$w, b$h, b$x[i], b$y[i], b$cw[i], b$ch[i])
}

.fr_shape <- function(x, y, r, st) {
  ifelse(st == "palm",
         sprintf('<path class="p-palm" d="M%.1f %.1fl%.1f %.1fl%.1f %.1fl%.1f %.1fz"/>', x, y - r, r, r, -r, r, -r, -r),
         sprintf('<circle class="p-%s" cx="%.1f" cy="%.1f" r="%.1f"/>', st, x, y, r))
}

# Place one label per point: above, below, right or left, whichever overlaps
# the fewest labels and points already placed (vectorised check per point).
.fr_place_labels <- function(x, y, r, tags, fs) {
  n <- length(x)
  cw <- fs * 0.6
  bx <- matrix(NA_real_, n, 4)
  out <- matrix(NA_real_, n, 2)
  cx0 <- x - r; cx1 <- x + r; cy0 <- y - r; cy1 <- y + r
  for (j in seq_len(n)) {
    w <- max(1, nchar(tags[j])) * cw
    cand <- rbind(c(x[j] - w / 2, y[j] - r[j] - 1.5),
                  c(x[j] - w / 2, y[j] + r[j] + fs - 0.5),
                  c(x[j] + r[j] + 1.5, y[j] + fs / 2.8),
                  c(x[j] - r[j] - 1.5 - w, y[j] + fs / 2.8))
    best <- 1; best_cost <- Inf
    for (q in 1:4) {
      b0 <- cand[q, 1]; b2 <- b0 + w; b1 <- cand[q, 2] - fs * 0.78; b3 <- cand[q, 2] + 1
      cost <- 2 * sum(b0 < bx[, 3] & b2 > bx[, 1] & b1 < bx[, 4] & b3 > bx[, 2], na.rm = TRUE) +
        sum(b0 < cx1[-j] & b2 > cx0[-j] & b1 < cy1[-j] & b3 > cy0[-j])
      if (cost < best_cost) { best_cost <- cost; best <- q }
      if (cost == 0) break
    }
    out[j, ] <- cand[best, ]
    bx[j, ] <- c(cand[best, 1], cand[best, 2] - fs * 0.78, cand[best, 1] + w, cand[best, 2] + 1)
  }
  out
}

.fr_subplot_svg <- function(sd, S, aria) {
  P <- 300; m <- 26; k <- (P - 2 * m) / S
  step <- if (S <= 12) 1 else if (S <= 30) 5 else 10
  ts <- seq(0, S, by = step)
  grid <- paste(sprintf('<line class="gl" x1="%.1f" y1="%d" x2="%.1f" y2="%d"/><line class="gl" x1="%d" y1="%.1f" x2="%d" y2="%.1f"/>',
                        m + ts * k, m, m + ts * k, P - m, m, m + (S - ts) * k, P - m, m + (S - ts) * k), collapse = "")
  lab_t <- ts[seq(1, length(ts), by = if (length(ts) > 6) 2 else 1)]
  axes <- paste(sprintf('<text class="ax" x="%.1f" y="%d">%s</text><text class="ax" x="%d" y="%.1f">%s</text>',
                        m + lab_t * k, P - m + 13, lab_t, m - 9, m + (S - lab_t) * k + 3, lab_t), collapse = "")
  sd <- sd[is.finite(sd$lx) & is.finite(sd$ly), , drop = FALSE]
  pts <- ""
  if (nrow(sd)) {
    sd <- sd[order(-ifelse(is.finite(sd$dbh), sd$dbh, 0)), , drop = FALSE]
    x <- m + sd$lx * k; y <- m + (S - sd$ly) * k
    r <- ifelse(is.finite(sd$dbh), 3 + sqrt(pmax(sd$dbh, 0)) * 1.15, 3) * (10 / max(S, 10))^0.5
    lp <- .fr_place_labels(x, y, r, sd$tag, 8.5)
    pts <- paste(sprintf('<g class="t" data-tag="%s" tabindex="0"><title>%s \u00b7 %s \u00b7 %s cm</title>%s</g>',
                         .fr_esc(sd$tag), .fr_esc(sd$tag), .fr_esc(sd$spp), .fr_fmt(sd$dbh),
                         .fr_shape(x, y, r, sd$st)), collapse = "")
    pts <- paste0(pts, '<g class="tls">',
                  paste(sprintf('<text x="%.1f" y="%.1f">%s</text>', lp[, 1], lp[, 2], .fr_esc(sd$tag)), collapse = ""), '</g>')
  }
  sprintf('<svg class="map" viewBox="0 0 %d %d" role="img" aria-label="%s"><rect class="frame" x="%d" y="%d" width="%d" height="%d"/>%s%s%s</svg>',
          P, P, .fr_esc(aria), m, m, P - 2 * m, P - 2 * m, grid, axes, pts)
}

.fr_plot_svg <- function(lay, dd, aria) {
  # MONITORA: arms in the schematic position (closer to the centre), no axes
  sch <- identical(lay$kind, "monitora")
  e <- if (sch) lay$grid_extent else lay$extent
  if (sch) { dd$gx <- dd$sgx; dd$gy <- dd$sgy }
  span <- max(e[3] - e[1], e[4] - e[2])
  P <- 1000; m <- 34; k <- (P - 2 * m) / span
  fs <- 8
  b <- .fr_cell_boxes(lay, P, pad = m, schematic = sch)
  c <- lay$cells
  cells <- paste(sprintf('<a href="#s-%s"><rect class="pc" x="%.1f" y="%.1f" width="%.1f" height="%.1f"/><text class="pn-l" x="%.1f" y="%.1f">%s</text></a>',
                         c$id, b$x, b$y, b$cw, b$ch, b$x + 3, b$y + 10, .fr_esc(c$label)), collapse = "")
  step <- if (span <= 120) 20 else if (span <= 300) 50 else 100
  tx <- seq(ceiling(e[1] / step) * step, e[3], by = step)
  ty <- seq(ceiling(e[2] / step) * step, e[4], by = step)
  axes <- if (sch) "" else paste0(
    paste(sprintf('<text class="ax" x="%.1f" y="%.1f">%s</text>', m + (tx - e[1]) * k, b$h - m + 16, tx), collapse = ""),
    paste(sprintf('<text class="ax" x="%d" y="%.1f">%s</text>', m - 14, m + (e[4] - ty) * k + 3, ty), collapse = ""))
  dd <- dd[is.finite(dd$gx) & is.finite(dd$gy), , drop = FALSE]
  body <- ""
  if (nrow(dd)) {
    dd <- dd[order(dd$st == "todo", -ifelse(is.finite(dd$dbh), dd$dbh, 0)), , drop = FALSE]
    x <- m + (dd$gx - e[1]) * k; y <- m + (e[4] - dd$gy) * k
    zoom <- (100 / span)^0.5
    r <- ifelse(is.finite(dd$dbh), 2 + sqrt(pmax(dd$dbh, 0)) * 0.55, 2) * zoom * ifelse(dd$st == "todo", 1.5, 1)
    lp <- .fr_place_labels(x, y, r, dd$tag, fs)
    body <- paste0(paste(.fr_shape(x, y, r, dd$st), collapse = ""), '<g class="wl">',
                   paste(sprintf('<text x="%.1f" y="%.1f">%s</text>', lp[, 1], lp[, 2], .fr_esc(dd$tag)), collapse = ""), '</g>')
  }
  sprintf('<svg class="whole" viewBox="0 0 %.0f %.0f" role="img" aria-label="%s">%s%s%s</svg>',
          b$w, b$h, .fr_esc(aria), cells, axes, body)
}


# HTML -------------------------------------------------------------------------

#' Build the field report as one HTML string
#'
#' @param lay Layout from `.field_report_layout()`.
#' @param meta List with `plot_name`, `plot_code`, `plot_census_no_fp`, `team`.
#' @keywords internal
#' @noRd
.field_report_html <- function(lay, meta, language = "en", highlight_palms = TRUE,
                               plot_area_ha = NULL) {
  tr <- .field_report_strings(language)
  d <- lay$trees
  c <- lay$cells
  hp <- isTRUE(highlight_palms)
  lang_attr <- switch(language, pt = "pt-BR", pa = "pt-BR", ma = "zh", language)
  pc <- .fr_esc(if (.safe_nzchar(meta$plot_code)) meta$plot_code else "")
  pn <- .fr_esc(if (.safe_nzchar(meta$plot_name)) meta$plot_name else "")
  sub_prefix <- if (lay$kind == "grid") "S" else ""
  slab <- function(lbl) paste0(sub_prefix, lbl)

  n_ok <- sum(d$st == "ok"); n_todo <- sum(d$st == "todo"); n_palm <- sum(d$st == "palm")
  n_all <- nrow(d); denom <- n_ok + n_todo
  pct <- if (denom > 0) 100 * n_ok / denom else NA_real_
  fam_id <- length(unique(d$fam[d$fam != "Indet"]))
  n_indet <- sum(d$fam == "Indet")
  n_visit <- sum(c$todo > 0)
  mk <- '<span class="mk" aria-hidden="true">\u25b2</span>'

  kpi <- function(label, value, ctx = "", cls = "") sprintf(
    '<div class="kpi-i %s"><dt><span>%s</span></dt><dd><span class="v">%s</span><span class="c">%s</span></dd></div>',
    cls, label, value, ctx)

  # head ----
  census <- if (.safe_nzchar(as.character(meta$plot_census_no_fp))) paste0(" \u00b7 ", tr[["census"]], " ", .fr_esc(meta$plot_census_no_fp)) else ""
  size_txt <- if (lay$kind == "grid") {
    paste0(if (!is.null(plot_area_ha)) paste0(format(plot_area_ha, trim = TRUE, drop0trailing = TRUE), " ha \u00b7 ") else "",
           sprintf(tr[["subplots_of"]], nrow(c), lay$cell_size, lay$cell_size))
  } else {
    paste0("MONITORA \u00b7 ", sprintf(tr[["subplots_of"]], nrow(c), 10, 10))
  }
  team_txt <- if (.safe_nzchar(as.character(meta$team))) paste0('<p class="meta">', tr[["team"]], ": ", .fr_esc(meta$team), "</p>") else ""
  ver <- tryCatch(as.character(utils::packageVersion("forplotR")), error = function(e) "")

  fns <- c(
    if (n_indet > 0) sprintf(tr[["fn_indet"]], n_indet),
    if (hp) sprintf(tr[["fn_voucher_np"]], denom, n_ok, n_todo) else sprintf(tr[["fn_voucher_all"]], denom, n_ok, n_todo),
    if (hp) sprintf(tr[["fn_palms"]], n_ok, n_todo, n_palm, n_all)
  )
  fn_no <- function(key) {
    k <- c(indet = if (n_indet > 0) 1L else NA, voucher = if (n_indet > 0) 2L else 1L,
           palms = if (n_indet > 0) 3L else 2L)
    sprintf("<sup>%d</sup>", k[[key]])
  }

  map_links <- paste0(
    sprintf('<li><a href="#map-all">%s<span>%d</span></a></li>', tr[["map_all"]], n_all),
    sprintf('<li><a href="#map-ok">%s<span>%d</span></a></li>', tr[["with_voucher"]], n_ok),
    sprintf('<li><a href="#map-todo">%s<span>%d</span></a></li>', tr[["pending"]], n_todo),
    if (hp) sprintf('<li><a href="#map-palm">%s<span>%d</span></a></li>', tr[["palms"]], n_palm) else "")

  summary_html <- paste0(
    '<section class="page cover" id="summary">',
    '<header class="doc-head">',
    sprintf('<p class="kicker">%s%s</p>', tr[["doc_title"]], census),
    sprintf('<h1>%s <span>%s</span></h1>', pn, pc),
    sprintf('<p class="meta">%s%s</p>', size_txt, if (nzchar(ver)) paste0(" \u00b7 forplotR ", ver) else ""),
    team_txt, '</header>',
    '<dl class="kpi">',
    kpi(tr[["individuals"]], n_all, sprintf(tr[["in_subplots"]], sum(c$n > 0))),
    kpi(tr[["families"]], fam_id, paste0(tr[["identified"]], if (n_indet > 0) fn_no("indet") else "")),
    kpi(tr[["with_voucher"]], sprintf("%d / %d", n_ok, denom),
        paste0(sprintf(if (hp) tr[["pct_nonpalms"]] else tr[["pct_all"]], .fr_fmt(pct)), fn_no("voucher"))),
    kpi(tr[["pending"]], paste0(mk, n_todo), if (hp) tr[["pending_ctx_np"]] else tr[["pending_ctx"]], "pend"),
    if (hp) kpi(paste0(tr[["palms"]], ' <span class="q">', tr[["counted_apart"]], '</span>'), n_palm, paste0("Arecaceae", fn_no("palms"))) else "",
    kpi(tr[["subplots_visit"]], sprintf("%d / %d", n_visit, nrow(c)), tr[["with_pending"]]),
    '</dl>',
    '<ol class="fn">', paste(sprintf("<li>%s</li>", fns), collapse = ""), '</ol>',
    '<div class="two"><figure><figcaption>', tr[["grid_cap"]], '</figcaption>',
    .fr_grid_svg(lay, size = 300),
    sprintf('<p class="leg"><i class="c-many"></i>%s <i class="c-few"></i>1\u20132 <i class="c-done"></i>%s <i class="c-empty"></i>%s</p>',
            tr[["leg_many"]], tr[["leg_done"]], tr[["leg_empty"]]),
    '</figure><figure><figcaption>', tr[["maps_cap"]], '</figcaption><ul class="maplinks">', map_links, '</ul></figure></div>',
    '</section>')

  # whole-plot maps ----
  legend_for <- function(keys) {
    all <- c(todo = paste0('<i class="k-todo"></i>', tr[["leg_pending"]]),
             ok = paste0('<i class="k-ok"></i>', tr[["leg_voucher"]]),
             palm = paste0('<i class="k-palm"></i>', tr[["leg_palm"]]))
    keys <- intersect(keys, if (hp) names(all) else c("todo", "ok"))
    paste0('<p class="leg">', paste(all[keys], collapse = " "), " \u00b7 ", tr[["leg_size"]], " \u00b7 ", tr[["leg_tag"]], "</p>",
           if (lay$kind == "monitora") paste0('<p class="leg">', tr[["arms_note"]], "</p>") else "")
  }
  map_page <- function(id, title, dd, lede, keys) paste0(
    sprintf('<section class="page mapp" id="%s"><header class="sub-head"><div><p class="kicker">%s \u00b7 %s</p><h2>%s <span class="nn">%d %s</span></h2></div><nav class="pn"><a href="#summary" class="up">%s</a></nav></header>',
            id, pc, tr[["plot_map"]], title, nrow(dd), tr[["stems"]], tr[["index"]]),
    sprintf('<p class="lede">%s%s</p>', lede, tr[["tap_subplot"]]),
    '<div class="mapwrap">', .fr_plot_svg(lay, dd, title), '</div>', legend_for(keys), '</section>')
  maps_html <- paste0(
    map_page("map-all", tr[["map_all"]], d, tr[["lede_all"]], c("todo", "ok", "palm")),
    map_page("map-ok", tr[["with_voucher"]], d[d$st == "ok", , drop = FALSE], tr[["lede_ok"]], "ok"),
    map_page("map-todo", tr[["pending"]], d[d$st == "todo", , drop = FALSE],
             paste0(tr[["lede_todo"]], if (hp) tr[["lede_todo_palms"]] else ""), "todo"),
    if (hp) map_page("map-palm", tr[["palms"]], d[d$st == "palm", , drop = FALSE], tr[["lede_palm"]], "palm") else "")

  # priority route ----
  pri <- c[c$todo > 0, , drop = FALSE]
  pri <- pri[order(-pri$todo, pri$order), , drop = FALSE]
  pend <- d[d$st == "todo", , drop = FALSE]
  route_body <- if (nrow(pri)) {
    lbl_of <- stats::setNames(c$label, c$id)
    keys <- unique(pend[, c("fam", "spp")])
    keys$n <- vapply(seq_len(nrow(keys)), function(i) sum(pend$fam == keys$fam[i] & pend$spp == keys$spp[i]), numeric(1))
    keys <- keys[order(-keys$n, keys$fam, keys$spp), , drop = FALSE]
    rows <- vapply(seq_len(nrow(keys)), function(i) {
      x <- pend[pend$fam == keys$fam[i] & pend$spp == keys$spp[i], , drop = FALSE]
      sprintf('<tr><td>%s</td><td><i>%s</i></td><td class="r">%d</td><td class="tags">%s</td></tr>',
              .fr_esc(keys$fam[i]), .fr_esc(keys$spp[i]), keys$n[i],
              paste(sprintf('<a href="#%s">%s</a><small>%s</small>', x$row_id, .fr_esc(x$tag), .fr_esc(slab(lbl_of[x$cell]))), collapse = " "))
    }, character(1))
    paste0(
      '<ol class="route">', paste(sprintf('<li><a href="#s-%s"><b>%s</b><span>%s</span></a></li>',
                                          pri$id, .fr_esc(slab(pri$label)), sprintf(tr[["n_pending"]], pri$todo)), collapse = ""), '</ol>',
      '<h3>', tr[["spp_pending"]], '</h3><table class="tbl"><thead><tr><th>', tr[["family"]], '</th><th>', tr[["species"]],
      '</th><th class="r">', tr[["n_stems"]], '</th><th>', tr[["tags_sub"]], '</th></tr></thead><tbody>',
      paste(rows, collapse = ""), '</tbody></table>')
  } else {
    sprintf('<p class="todo done"><b>%s</b></p>', tr[["nothing_pending"]])
  }
  route_html <- paste0('<section class="page" id="route"><h2>', tr[["route"]], '</h2><p class="lede">', tr[["route_lede"]],
                       if (hp) tr[["route_palms"]] else "", '</p>', route_body, '</section>')

  # subplot pages ----
  legend_sm <- paste0('<p class="leg sm"><i class="k-todo"></i>', tr[["leg_pending"]], ' <i class="k-ok"></i>', tr[["leg_voucher"]],
                      if (hp) paste0(' <i class="k-palm"></i>', tr[["leg_palm"]]) else "", '</p>')
  th <- paste0('<thead><tr><th>', tr[["tag"]], '</th><th>', tr[["species"]], '</th><th class="r">', tr[["dbh"]], '<small>cm</small></th>',
               '<th class="r">X; Y<small>m</small></th><th class="f">', tr[["new_dbh"]], '</th><th class="f s">', tr[["st"]],
               '</th><th class="f">', tr[["coll_no"]], '</th><th class="f n">', tr[["notes"]], '</th></tr></thead>')
  by_cell <- split(d, factor(d$cell, levels = c$id))
  sub_page <- function(i) {
    ci <- c[i, ]
    sd <- by_cell[[ci$id]]
    tn <- .fr_num(sd$tag)
    sd <- sd[order(sd$st != "todo", is.na(tn), tn, sd$tag), , drop = FALSE]
    prev <- if (i > 1) sprintf('<a href="#s-%s" aria-label="%s">\u2039 %s</a>', c$id[i - 1], tr[["prev"]], .fr_esc(slab(c$label[i - 1]))) else ""
    nxt <- if (i < nrow(c)) sprintf('<a href="#s-%s" aria-label="%s">%s \u203a</a>', c$id[i + 1], tr[["nxt"]], .fr_esc(slab(c$label[i + 1]))) else ""
    where <- if (lay$kind == "grid") sprintf(tr[["col_row"]], as.integer(ci$where_a), as.integer(ci$where_b)) else sprintf(tr[["arm_sub"]], ci$where_a, as.integer(ci$where_b))
    rows <- if (nrow(sd)) paste(sprintf(
      '<tr id="%s" data-tag="%s" class="r-%s"><td class="tag">%s</td><td><i>%s</i><small>%s</small></td><td class="r">%s</td><td class="r xy">%s; %s</td><td class="f"></td><td class="f s"></td><td class="f"></td><td class="f n"></td></tr>',
      sd$row_id, .fr_esc(sd$tag), sd$st, .fr_esc(sd$tag), .fr_esc(sd$spp), .fr_esc(sd$fam), .fr_fmt(sd$dbh),
      .fr_coord(sd$x_raw), .fr_coord(sd$y_raw)), collapse = "") else sprintf('<tr><td colspan="8" class="none">%s</td></tr>', tr[["no_stems"]])
    todo <- sd[sd$st == "todo", , drop = FALSE]
    title <- paste(tr[["subplot"]], .fr_esc(ci$label))
    paste0(
      sprintf('<section class="page sub%s" id="s-%s">', if (nrow(sd) > 18) " dense" else "", ci$id),
      sprintf('<header class="sub-head"><div><p class="kicker">%s \u00b7 %s</p><h2>%s</h2></div><nav class="pn">%s<a href="#summary" class="up">%s</a>%s</nav></header>',
              pc, pn, title, prev, tr[["index"]], nxt),
      '<div class="sub-top">', .fr_subplot_svg(sd, lay$cell_size, title),
      '<aside>', .fr_mini_svg(lay, i, tr[["back_grid"]]),
      sprintf('<p class="where">%s</p>', where),
      '<dl class="kpi sm">', kpi(tr[["individuals"]], ci$n),
      kpi(tr[["pending_short"]], if (ci$todo > 0) paste0(mk, ci$todo) else "0", "", if (ci$todo > 0) "pend" else ""),
      if (hp) kpi(tr[["palms"]], ci$palm) else "", '</dl>', legend_sm, '</aside></div>',
      if (nrow(todo)) sprintf('<p class="todo"><b>%s%s</b> %s</p>', mk, tr[["pending"]],
                              paste(sprintf('<a href="#%s">%s</a>', todo$row_id, .fr_esc(todo$tag)), collapse = " "))
      else sprintf('<p class="todo done"><b>%s</b></p>', tr[["no_pending_sub"]]),
      '<table class="tbl trees">', th, '<tbody>', rows, '</tbody></table>',
      '<p class="foot">', tr[["st_legend"]], '</p></section>')
  }
  subs_html <- paste(vapply(seq_len(nrow(c)), sub_page, character(1)), collapse = "\n")

  # checklist ----
  fams <- sort(unique(d$fam))
  chk <- vapply(fams, function(f) {
    x <- d[d$fam == f, , drop = FALSE]
    spp <- sort(unique(x$spp))
    paste0('<div class="fam"><h3>', .fr_esc(f), '</h3>', paste(vapply(spp, function(s) {
      y <- x[x$spp == s, , drop = FALSE]
      sprintf('<p><i>%s</i> %s</p>', .fr_esc(s), paste(sprintf('<a class="k-%s" href="#%s">%s</a>', y$st, y$row_id, .fr_esc(y$tag)), collapse = " "))
    }, character(1)), collapse = ""), '</div>')
  }, character(1))
  chk_html <- paste0('<section class="page appx" id="checklist"><h2>', tr[["appendix"]], '</h2><p class="lede">', tr[["chk_lede"]],
                     '</p><div class="chk">', paste(chk, collapse = ""), '</div></section>')

  first_lbl <- slab(c$label[min(nrow(c), if (lay$kind == "grid") 37 else 3)])
  paste0(
    '<!doctype html><html lang="', lang_attr, '"><head><meta charset="utf-8">',
    '<meta name="viewport" content="width=device-width,initial-scale=1">',
    '<meta name="generator" content="forplotR ', ver, '">',
    '<title>', pc, if (nzchar(pc)) " \u00b7 " else "", tr[["doc_title"]], '</title>',
    '<style>', .fr_read_asset("field_report.css"), '</style></head>',
    sprintf('<body data-not-found="%s" data-sub-prefix="%s">', .fr_esc(tr[["not_found"]]), sub_prefix),
    '<div class="bar"><a href="#summary" class="brand">', if (nzchar(pc)) pc else tr[["index"]], '</a>',
    '<a href="#map-all">', tr[["maps"]], '</a><a href="#route">', tr[["route_short"]], '</a><a href="#checklist">', tr[["checklist"]], '</a>',
    sprintf('<form class="find" onsubmit="return false"><input id="find" type="search" placeholder="%s" aria-label="%s" autocomplete="off"></form>',
            .fr_esc(sprintf(tr[["find_ph"]], first_lbl)), .fr_esc(tr[["find_label"]])),
    sprintf('<button class="print" type="button" onclick="window.print()" title="%s">PDF</button></div>', .fr_esc(tr[["print_pdf"]])),
    .fr_mini_symbol(lay),
    '<main>', summary_html, maps_html, route_html, subs_html, chk_html, '</main>',
    '<script>', .fr_read_asset("field_report.js"), '</script></body></html>')
}

#' Write the field report
#'
#' @return The path to the HTML file, invisibly.
#' @keywords internal
#' @noRd
.write_field_report <- function(file, fp_coords, input_type, subplot_size,
                                plot_width_m, plot_length_m, highlight_palms,
                                meta, language, plot_area_ha = NULL) {
  lay <- .field_report_layout(fp_coords, input_type, subplot_size,
                              plot_width_m, plot_length_m, highlight_palms)
  html <- .field_report_html(lay, meta, language, highlight_palms, plot_area_ha)
  con <- file(file, open = "wb")
  on.exit(close(con), add = TRUE)
  writeBin(charToRaw(enc2utf8(html)), con)
  invisible(file)
}


# PDF --------------------------------------------------------------------------

#' Find a Chrome, Chromium or Edge executable
#'
#' Looks at the `FORPLOTR_CHROME`, `CHROMOTE_CHROME` and `PAGEDOWN_CHROME`
#' environment variables, then the usual program names and install locations.
#'
#' @return Path to the executable, or `NULL`.
#' @keywords internal
#' @noRd
.find_chrome <- function() {
  env <- Sys.getenv(c("FORPLOTR_CHROME", "CHROMOTE_CHROME", "PAGEDOWN_CHROME"))
  env <- env[nzchar(env) & file.exists(env)]
  if (length(env)) return(unname(env[[1]]))

  if (identical(.Platform$OS.type, "windows")) {
    roots <- Sys.getenv(c("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"))
    roots <- roots[nzchar(roots)]
    cand <- c(file.path(roots, "Google", "Chrome", "Application", "chrome.exe"),
              file.path(roots, "Microsoft", "Edge", "Application", "msedge.exe"),
              file.path(roots, "Chromium", "Application", "chrome.exe"))
  } else if (identical(Sys.info()[["sysname"]], "Darwin")) {
    cand <- c("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
              "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
              "/Applications/Chromium.app/Contents/MacOS/Chromium")
  } else {
    cand <- Sys.which(c("google-chrome", "google-chrome-stable", "chromium",
                        "chromium-browser", "microsoft-edge", "microsoft-edge-stable"))
  }
  cand <- unname(cand[nzchar(cand) & file.exists(cand)])
  if (length(cand)) cand[[1]] else NULL
}

#' Print an HTML file to PDF with a headless browser
#'
#' @return The PDF path, or `NULL` (with a warning) when no browser is found or
#'   printing fails.
#' @keywords internal
#' @noRd
.html_to_pdf <- function(html_path, pdf_path, verbose = TRUE, timeout = 180) {
  chrome <- .find_chrome()
  if (is.null(chrome)) {
    warning(
      "No Chrome, Chromium or Edge browser was found, so the PDF report was not ",
      "generated. Open the HTML report and print it to PDF (A4), or set the ",
      "FORPLOTR_CHROME environment variable to the browser executable.",
      call. = FALSE
    )
    return(NULL)
  }
  html_path <- normalizePath(html_path, winslash = "/", mustWork = TRUE)
  url <- paste0(if (startsWith(html_path, "/")) "file://" else "file:///",
                utils::URLencode(html_path))
  pdf_abs <- file.path(normalizePath(dirname(pdf_path), winslash = "/", mustWork = TRUE), basename(pdf_path))
  if (file.exists(pdf_abs)) unlink(pdf_abs)
  profile <- tempfile("forplotR-chrome-")
  dir.create(profile)
  on.exit(unlink(profile, recursive = TRUE, force = TRUE), add = TRUE)
  args <- c("--headless", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
            "--disable-extensions", "--no-pdf-header-footer",
            if (identical(.Platform$OS.type, "unix") && !identical(Sys.info()[["sysname"]], "Darwin")) "--no-sandbox",
            shQuote(paste0("--user-data-dir=", profile)),
            shQuote(paste0("--print-to-pdf=", pdf_abs)), shQuote(url))
  if (verbose) message("Printing the PDF report with ", basename(chrome), "...")
  status <- tryCatch(
    suppressWarnings(system2(chrome, args, stdout = FALSE, stderr = FALSE, timeout = timeout)),
    error = function(e) 1L
  )
  if (!file.exists(pdf_abs) || file.size(pdf_abs) == 0) {
    warning("The PDF report could not be printed (browser exit status ", status,
            "). Open the HTML report and print it to PDF instead.", call. = FALSE)
    return(NULL)
  }
  pdf_abs
}
