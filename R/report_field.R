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
    date            = c("Date", "Data", "Fecha", "Date", "\u65e5\u671f"),
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
    phyto_title     = c("Floristic and phytosociological overview", "Flor\u00edstica e fitossociologia", "Flor\u00edstica y fitosociolog\u00eda", "Floristique et phytosociologie", "\u690d\u7269\u533a\u7cfb\u4e0e\u690d\u7269\u7fa4\u843d\u5b66\u6982\u51b5"),
    phyto_short     = c("Floristics", "Flor\u00edstica", "Flor\u00edstica", "Floristique", "\u690d\u7269\u533a\u7cfb"),
    phyto_lede      = c("Parameters computed from all stems of the census (Mueller-Dombois & Ellenberg 1974). The sampling unit for frequency is the subplot.",
                        "Par\u00e2metros calculados com todos os fustes do censo (Mueller-Dombois & Ellenberg 1974). A unidade amostral da frequ\u00eancia \u00e9 a subparcela.",
                        "Par\u00e1metros calculados con todos los fustes del censo (Mueller-Dombois & Ellenberg 1974). La unidad de muestreo de la frecuencia es la subparcela.",
                        "Param\u00e8tres calcul\u00e9s sur toutes les tiges du recensement (Mueller-Dombois & Ellenberg 1974). L'unit\u00e9 d'\u00e9chantillonnage de la fr\u00e9quence est la sous-parcelle.",
                        "\u53c2\u6570\u57fa\u4e8e\u672c\u6b21\u666e\u67e5\u7684\u5168\u90e8\u830e\u5e72\u8ba1\u7b97\uff08Mueller-Dombois & Ellenberg 1974\uff09\u3002\u9891\u5ea6\u7684\u62bd\u6837\u5355\u5143\u4e3a\u5b50\u6837\u5730\u3002"),
    k_species       = c("Species", "Esp\u00e9cies", "Especies", "Esp\u00e8ces", "\u7269\u79cd"),
    k_genera        = c("Genera", "G\u00eaneros", "G\u00e9neros", "Genres", "\u5c5e"),
    k_density       = c("Density", "Densidade", "Densidad", "Densit\u00e9", "\u5bc6\u5ea6"),
    u_density       = c("stems/ha", "fustes/ha", "fustes/ha", "tiges/ha", "\u830e\u5e72/\u516c\u9877"),
    k_ba            = c("Basal area", "\u00c1rea basal", "\u00c1rea basal", "Surface terri\u00e8re", "\u80f8\u9ad8\u65ad\u9762\u79ef"),
    u_ba            = c("m\u00b2/ha", "m\u00b2/ha", "m\u00b2/ha", "m\u00b2/ha", "\u5e73\u65b9\u7c73/\u516c\u9877"),
    k_taxa_ctx      = c("names as determined", "nomes como determinados", "nombres como determinados", "noms tels que d\u00e9termin\u00e9s", "\u6309\u9274\u5b9a\u540d\u79f0"),
    k_div_ctx       = c("diversity", "diversidade", "diversidad", "diversit\u00e9", "\u591a\u6837\u6027"),
    k_even_ctx      = c("evenness", "equabilidade", "equitatividad", "\u00e9quitabilit\u00e9", "\u5747\u5300\u5ea6"),
    fn_taxa         = c("S counts the names as written in the determination column; each name with \u201cindet\u201d counts as one taxon.",
                        "S conta os nomes como est\u00e3o na coluna de determina\u00e7\u00e3o; cada nome com \u201cindet\u201d conta como um t\u00e1xon.",
                        "S cuenta los nombres tal como est\u00e1n en la columna de determinaci\u00f3n; cada nombre con \u201cindet\u201d cuenta como un tax\u00f3n.",
                        "S compte les noms tels qu'\u00e9crits dans la colonne de d\u00e9termination ; chaque nom avec \u00ab indet \u00bb compte comme un taxon.",
                        "S \u6309\u9274\u5b9a\u5217\u4e2d\u7684\u540d\u79f0\u8ba1\u6570\uff1b\u6bcf\u4e2a\u542b\u201cindet\u201d\u7684\u540d\u79f0\u8ba1\u4e3a\u4e00\u4e2a\u5206\u7c7b\u5355\u5143\u3002"),
    fn_area         = c("Density and basal area per hectare use %s ha sampled; basal area is computed from the diameter of each stem.",
                        "Densidade e \u00e1rea basal por hectare usam %s ha amostrados; a \u00e1rea basal \u00e9 calculada com o di\u00e2metro de cada fuste.",
                        "La densidad y el \u00e1rea basal por hect\u00e1rea usan %s ha muestreadas; el \u00e1rea basal se calcula con el di\u00e1metro de cada fuste.",
                        "La densit\u00e9 et la surface terri\u00e8re par hectare utilisent %s ha \u00e9chantillonn\u00e9s ; la surface terri\u00e8re est calcul\u00e9e \u00e0 partir du diam\u00e8tre de chaque tige.",
                        "\u6bcf\u516c\u9877\u5bc6\u5ea6\u548c\u80f8\u9ad8\u65ad\u9762\u79ef\u57fa\u4e8e %s \u516c\u9877\u6837\u5730\u9762\u79ef\uff1b\u80f8\u9ad8\u65ad\u9762\u79ef\u7531\u6bcf\u4e2a\u830e\u5e72\u7684\u76f4\u5f84\u8ba1\u7b97\u3002"),
    fn_div          = c("H\u2032 = \u2212\u03a3 p\u1d62 ln p\u1d62; 1 \u2212 D = 1 \u2212 \u03a3 p\u1d62\u00b2; J\u2032 = H\u2032 / ln S, where p\u1d62 is the proportion of stems of taxon i.",
                        "H\u2032 = \u2212\u03a3 p\u1d62 ln p\u1d62; 1 \u2212 D = 1 \u2212 \u03a3 p\u1d62\u00b2; J\u2032 = H\u2032 / ln S, em que p\u1d62 \u00e9 a propor\u00e7\u00e3o de fustes do t\u00e1xon i.",
                        "H\u2032 = \u2212\u03a3 p\u1d62 ln p\u1d62; 1 \u2212 D = 1 \u2212 \u03a3 p\u1d62\u00b2; J\u2032 = H\u2032 / ln S, donde p\u1d62 es la proporci\u00f3n de fustes del tax\u00f3n i.",
                        "H\u2032 = \u2212\u03a3 p\u1d62 ln p\u1d62 ; 1 \u2212 D = 1 \u2212 \u03a3 p\u1d62\u00b2 ; J\u2032 = H\u2032 / ln S, o\u00f9 p\u1d62 est la proportion de tiges du taxon i.",
                        "H\u2032 = \u2212\u03a3 p\u1d62 ln p\u1d62\uff1b1 \u2212 D = 1 \u2212 \u03a3 p\u1d62\u00b2\uff1bJ\u2032 = H\u2032 / ln S\uff0c\u5176\u4e2d p\u1d62 \u4e3a\u5206\u7c7b\u5355\u5143 i \u7684\u830e\u5e72\u6bd4\u4f8b\u3002"),
    fn_ivi          = c("RD relative density, RF relative frequency, RDo relative dominance (basal area), AF absolute frequency (% of subplots); IVI = RD + RF + RDo, summing to 300 over all taxa.",
                        "DR densidade relativa, FR frequ\u00eancia relativa, DoR domin\u00e2ncia relativa (\u00e1rea basal), FA frequ\u00eancia absoluta (% das subparcelas); IVI = DR + FR + DoR, que soma 300 no conjunto dos t\u00e1xons.",
                        "DR densidad relativa, FR frecuencia relativa, DoR dominancia relativa (\u00e1rea basal), FA frecuencia absoluta (% de las subparcelas); IVI = DR + FR + DoR, que suma 300 en el conjunto de los taxones.",
                        "DR densit\u00e9 relative, FR fr\u00e9quence relative, DoR dominance relative (surface terri\u00e8re), FA fr\u00e9quence absolue (% des sous-parcelles) ; IVI = DR + FR + DoR, dont la somme vaut 300 sur l'ensemble des taxons.",
                        "RD \u76f8\u5bf9\u5bc6\u5ea6\uff0cRF \u76f8\u5bf9\u9891\u5ea6\uff0cRDo \u76f8\u5bf9\u663e\u8457\u5ea6\uff08\u80f8\u9ad8\u65ad\u9762\u79ef\uff09\uff0cAF \u7edd\u5bf9\u9891\u5ea6\uff08\u5b50\u6837\u5730\u767e\u5206\u6bd4\uff09\uff1bIVI = RD + RF + RDo\uff0c\u5168\u90e8\u5206\u7c7b\u5355\u5143\u4e4b\u548c\u4e3a 300\u3002"),
    ab_rd           = c("RD", "DR", "DR", "DR", "RD"),
    ab_rf           = c("RF", "FR", "FR", "FR", "RF"),
    ab_rdo          = c("RDo", "DoR", "DoR", "DoR", "RDo"),
    ab_af           = c("AF", "FA", "FA", "FA", "AF"),
    ab_ba           = c("BA", "AB", "AB", "ST", "BA"),
    comp_rd         = c("relative density", "densidade relativa", "densidad relativa", "densit\u00e9 relative", "\u76f8\u5bf9\u5bc6\u5ea6"),
    comp_rf         = c("relative frequency", "frequ\u00eancia relativa", "frecuencia relativa", "fr\u00e9quence relative", "\u76f8\u5bf9\u9891\u5ea6"),
    comp_rdo        = c("relative dominance", "domin\u00e2ncia relativa", "dominancia relativa", "dominance relative", "\u76f8\u5bf9\u663e\u8457\u5ea6"),
    fig_fam         = c("Figure 1. Families with the highest importance value (IVI)", "Figura 1. Fam\u00edlias com maior valor de import\u00e2ncia (IVI)", "Figura 1. Familias con mayor valor de importancia (IVI)", "Figure 1. Familles \u00e0 plus forte valeur d'importance (IVI)", "\u56fe 1. \u91cd\u8981\u503c\uff08IVI\uff09\u6700\u9ad8\u7684\u79d1"),
    fig_spp         = c("Figure 2. Species with the highest importance value (IVI)", "Figura 2. Esp\u00e9cies com maior valor de import\u00e2ncia (IVI)", "Figura 2. Especies con mayor valor de importancia (IVI)", "Figure 2. Esp\u00e8ces \u00e0 plus forte valeur d'importance (IVI)", "\u56fe 2. \u91cd\u8981\u503c\uff08IVI\uff09\u6700\u9ad8\u7684\u7269\u79cd"),
    fig_dbh         = c("Figure 3. Stems per diameter class", "Figura 3. Fustes por classe de di\u00e2metro", "Figura 3. Fustes por clase diam\u00e9trica", "Figure 3. Tiges par classe de diam\u00e8tre", "\u56fe 3. \u5404\u5f84\u7ea7\u830e\u5e72\u6570"),
    dbh_axis        = c("DBH class (cm)", "Classe de DAP (cm)", "Clase de DAP (cm)", "Classe de DHP (cm)", "\u80f8\u5f84\u7ea7\uff08\u5398\u7c73\uff09"),
    top_note        = c("First %d by IVI; every taxon is in the tables below.", "Os %d primeiros por IVI; todos os t\u00e1xons est\u00e3o nas tabelas abaixo.", "Los %d primeros por IVI; todos los taxones est\u00e1n en las tablas siguientes.", "Les %d premiers par IVI ; tous les taxons figurent dans les tableaux ci-dessous.", "\u6309 IVI \u6392\u540d\u524d %d\uff1b\u5168\u90e8\u5206\u7c7b\u5355\u5143\u89c1\u4e0b\u8868\u3002"),
    tab_fam         = c("Table 1. Phytosociological parameters of the families, by IVI", "Tabela 1. Par\u00e2metros fitossociol\u00f3gicos das fam\u00edlias, por IVI", "Tabla 1. Par\u00e1metros fitosociol\u00f3gicos de las familias, por IVI", "Tableau 1. Param\u00e8tres phytosociologiques des familles, par IVI", "\u8868 1. \u5404\u79d1\u7684\u690d\u7269\u7fa4\u843d\u5b66\u53c2\u6570\uff08\u6309 IVI \u6392\u5e8f\uff09"),
    tab_spp         = c("Table 2. Phytosociological parameters of the species, by IVI", "Tabela 2. Par\u00e2metros fitossociol\u00f3gicos das esp\u00e9cies, por IVI", "Tabla 2. Par\u00e1metros fitosociol\u00f3gicos de las especies, por IVI", "Tableau 2. Param\u00e8tres phytosociologiques des esp\u00e8ces, par IVI", "\u8868 2. \u5404\u7269\u79cd\u7684\u690d\u7269\u7fa4\u843d\u5b66\u53c2\u6570\uff08\u6309 IVI \u6392\u5e8f\uff09"),
    col_spp         = c("Spp.", "Spp.", "Spp.", "Esp.", "\u79cd\u6570"),
    col_sub         = c("Subpl.", "Subp.", "Subp.", "S.-p.", "\u5b50\u6837\u5730"),
    no_dbh          = c("No valid diameters: relative dominance and basal area are not available, so IVI = RD + RF.", "Sem di\u00e2metros v\u00e1lidos: domin\u00e2ncia relativa e \u00e1rea basal indispon\u00edveis, ent\u00e3o IVI = DR + FR.", "Sin di\u00e1metros v\u00e1lidos: dominancia relativa y \u00e1rea basal no disponibles, por lo que IVI = DR + FR.", "Pas de diam\u00e8tres valides : dominance relative et surface terri\u00e8re indisponibles, donc IVI = DR + FR.", "\u65e0\u6709\u6548\u80f8\u5f84\uff1a\u76f8\u5bf9\u663e\u8457\u5ea6\u548c\u80f8\u9ad8\u65ad\u9762\u79ef\u4e0d\u53ef\u7528\uff0c\u56e0\u6b64 IVI = RD + RF\u3002"),
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


# Floristics and phytosociology ------------------------------------------------
#
# Charts are inline SVG with a validated three-colour palette for the IVI
# components (relative density, relative frequency, relative dominance; it
# passes colour-vision-deficiency checks and is reinforced by gaps between
# segments, a legend, tooltips and the full tables). Every figure has its
# table: Table 1 (families) and Table 2 (species).

.fr_nice_step <- function(maxv, target = 5) {
  if (!is.finite(maxv) || maxv <= 0) return(1)
  raw <- maxv / target
  mag <- 10^floor(log10(raw))
  for (m in c(1, 2, 2.5, 5, 10)) if (m * mag >= raw) return(m * mag)
  10 * mag
}

.fr_trunc <- function(x, n = 30) {
  x <- as.character(x)
  ifelse(nchar(x) > n, paste0(substr(x, 1, n - 1), "\u2026"), x)
}

# Horizontal stacked bars of the IVI components for the first `n` taxa.
.fr_ivi_svg <- function(df, name_col, tr, n = 10, italic = FALSE, aria = "") {
  df <- as.data.frame(df)
  if (!nrow(df)) return("")
  df[[name_col]] <- .taxon_label(df[[name_col]], if (name_col == "species") df$family else NULL)
  df <- utils::head(df[order(-df$importance_value), , drop = FALSE], n)
  comp <- cbind(df$relative_density, df$relative_frequency, df$relative_dominance)
  comp[!is.finite(comp)] <- 0
  tot <- rowSums(comp)
  W <- 680; lw <- 200; rw <- 48; top <- 6; bh <- 15; gap <- 9
  H <- top + nrow(df) * (bh + gap) + 22
  step <- .fr_nice_step(max(tot))
  vmax <- max(step, ceiling(max(tot) / step) * step)
  k <- (W - lw - rw) / vmax
  ticks <- seq(0, vmax, by = step)
  base_y <- H - 20
  grid <- paste(sprintf('<line class="gl" x1="%.1f" y1="%d" x2="%.1f" y2="%.1f"/><text class="ax" x="%.1f" y="%.1f">%s</text>',
                        lw + ticks * k, top, lw + ticks * k, base_y, lw + ticks * k, base_y + 14,
                        format(ticks, trim = TRUE)), collapse = "")
  lab <- c(tr[["ab_rd"]], tr[["ab_rf"]], tr[["ab_rdo"]])
  rows <- vapply(seq_len(nrow(df)), function(i) {
    y <- top + (i - 1) * (bh + gap)
    nm <- as.character(df[[name_col]][i])
    x0 <- lw
    segs <- ""
    for (j in 1:3) {
      w <- comp[i, j] * k
      if (w <= 0) next
      segs <- paste0(segs, sprintf('<rect class="s%d" x="%.1f" y="%d" width="%.1f" height="%d"><title>%s \u00b7 %s %s</title></rect>',
                                   j, x0, y, w, bh, .fr_esc(nm), lab[j], formatC(comp[i, j], format = "f", digits = 1)))
      x0 <- x0 + w
    }
    sprintf('<g><text class="nm%s" x="%d" y="%.1f">%s<title>%s</title></text>%s<text class="vl" x="%.1f" y="%.1f">%s</text></g>',
            if (italic) " it" else "", lw - 8, y + bh * 0.75, .fr_esc(.fr_trunc(nm)), .fr_esc(nm), segs,
            lw + tot[i] * k + 6, y + bh * 0.75, formatC(tot[i], format = "f", digits = 1))
  }, character(1))
  sprintf('<svg class="chart ivi" viewBox="0 0 %d %d" role="img" aria-label="%s">%s%s<line class="bl" x1="%d" y1="%d" x2="%d" y2="%.1f"/></svg>',
          W, H, .fr_esc(aria), grid, paste(rows, collapse = ""), lw, top - 2, lw, base_y)
}

# Diameter classes (cm) with an adaptive class width.
.fr_dbh_classes <- function(dbh_cm) {
  d <- dbh_cm[is.finite(dbh_cm) & dbh_cm > 0]
  if (!length(d)) return(NULL)
  rng <- ceiling(max(d)) - floor(min(d))
  w <- if (rng <= 100) 10 else if (rng <= 200) 20 else if (rng <= 400) 25 else 50
  b0 <- floor(min(d) / w) * w
  breaks <- seq(b0, ceiling(max(d) / w) * w + w, by = w)
  cls <- cut(d, breaks = breaks, right = FALSE, include.lowest = TRUE)
  n <- as.integer(table(cls))
  keep <- seq_len(max(which(n > 0)))
  data.frame(lo = breaks[keep], hi = breaks[keep + 1], n = n[keep])
}

.fr_dbh_svg <- function(dbh_cm, tr) {
  cl <- .fr_dbh_classes(dbh_cm)
  if (is.null(cl)) return("")
  W <- 680; H <- 250; l <- 48; r <- 12; t <- 10; b <- 44
  slot <- (W - l - r) / nrow(cl)
  step <- .fr_nice_step(max(cl$n))
  vmax <- max(step, ceiling(max(cl$n) / step) * step)
  k <- (H - t - b) / vmax
  ticks <- seq(0, vmax, by = step)
  grid <- paste(sprintf('<line class="gl" x1="%d" y1="%.1f" x2="%d" y2="%.1f"/><text class="ax end" x="%d" y="%.1f">%s</text>',
                        l, H - b - ticks * k, W - r, H - b - ticks * k, l - 6, H - b - ticks * k + 3, format(ticks, trim = TRUE)), collapse = "")
  bw <- slot * 0.74
  fs <- if (nrow(cl) > 14) 8.5 else 10
  bars <- paste(sprintf('<rect class="s1" x="%.1f" y="%.1f" width="%.1f" height="%.1f" rx="2"><title>%s\u2013%s cm \u00b7 %d %s</title></rect><text class="ax" x="%.1f" y="%d" style="font-size:%.1fpx">%s\u2013%s</text>',
                        l + (seq_len(nrow(cl)) - 1) * slot + (slot - bw) / 2, H - b - cl$n * k, bw, pmax(cl$n * k, 0),
                        cl$lo, cl$hi, cl$n, tr[["stems"]],
                        l + (seq_len(nrow(cl)) - 0.5) * slot, H - b + 14, fs, cl$lo, cl$hi), collapse = "")
  sprintf('<svg class="chart dbh" viewBox="0 0 %d %d" role="img" aria-label="%s">%s%s<line class="bl" x1="%d" y1="%d" x2="%d" y2="%d"/><text class="ax" x="%.1f" y="%d">%s</text><text class="ax" transform="translate(12 %.1f) rotate(-90)">%s</text></svg>',
          W, H, .fr_esc(tr[["fig_dbh"]]), grid, bars, l, H - b, W - r, H - b,
          l + (W - l - r) / 2, H - 6, tr[["dbh_axis"]], t + (H - t - b) / 2, tr[["stems"]])
}

.fr_num_cell <- function(x, digits = 1) {
  ifelse(is.finite(x), formatC(x, format = "f", digits = digits, big.mark = ","), "\u2013")
}

#' Floristic and phytosociological section of the field report
#'
#' @param phyto Result of `.calculate_phytosociological_metrics()`.
#' @param dbh_cm Diameters (cm) of the stems in the report.
#' @param area_ha Sampled area in hectares (for density and basal area).
#' @keywords internal
#' @noRd
.fr_phyto_html <- function(phyto, dbh_cm, area_ha, tr, n_fig_fam = 10, n_fig_spp = 15) {
  sp <- as.data.frame(phyto$species_metrics)
  fa <- as.data.frame(phyto$family_metrics)
  if (!nrow(sp)) return("")
  dv <- phyto$diversity_metrics
  has_ba <- any(is.finite(sp$relative_dominance))

  S <- nrow(sp)
  gen <- sub(" .*", "", sp$species)
  n_gen <- length(unique(gen[!grepl("^indet$", gen, ignore.case = TRUE)]))
  n_fam <- length(unique(fa$family[fa$family != "Indet"]))
  N <- sum(sp$abundance)
  ba_tot <- sum(sp$basal_area_m2, na.rm = TRUE)
  pielou <- if (S > 1) dv$shannon_index / log(S) else NA_real_
  area_ok <- !is.null(area_ha) && is.finite(area_ha) && area_ha > 0

  kpi <- function(label, value, ctx = "") sprintf(
    '<div class="kpi-i"><dt><span>%s</span></dt><dd><span class="v">%s</span><span class="c">%s</span></dd></div>', label, value, ctx)
  strip <- paste0(
    '<dl class="kpi k4">',
    kpi(paste0(tr[["k_species"]], " (S)"), S, paste0(tr[["k_taxa_ctx"]], "<sup>1</sup>")),
    kpi(tr[["k_genera"]], n_gen, tr[["identified"]]),
    kpi(tr[["families"]], n_fam, tr[["identified"]]),
    kpi(tr[["k_density"]], if (area_ok) .fr_num_cell(N / area_ha, 0) else "\u2013", paste0(tr[["u_density"]], "<sup>2</sup>")),
    kpi(tr[["k_ba"]], if (area_ok && has_ba) .fr_num_cell(ba_tot / area_ha, 2) else "\u2013", paste0(tr[["u_ba"]], "<sup>2</sup>")),
    kpi("Shannon H\u2032", .fr_num_cell(dv$shannon_index, 2), paste0(tr[["k_div_ctx"]], "<sup>3</sup>")),
    kpi("Simpson 1 \u2212 D", .fr_num_cell(dv$simpson_index, 3), paste0(tr[["k_div_ctx"]], "<sup>3</sup>")),
    kpi("Pielou J\u2032", .fr_num_cell(pielou, 2), paste0(tr[["k_even_ctx"]], "<sup>3</sup>")),
    '</dl>')
  fns <- c(tr[["fn_taxa"]],
           if (area_ok) sprintf(tr[["fn_area"]], format(round(area_ha, 4), trim = TRUE, drop0trailing = TRUE)) else tr[["fn_area"]],
           tr[["fn_div"]], tr[["fn_ivi"]], if (!has_ba) tr[["no_dbh"]])
  fns[2] <- sub("%s", "\u2013", fns[2], fixed = TRUE)

  legend <- paste0('<p class="leg"><i class="k-s1"></i>', tr[["ab_rd"]], " ", tr[["comp_rd"]],
                   ' <i class="k-s2"></i>', tr[["ab_rf"]], " ", tr[["comp_rf"]],
                   if (has_ba) paste0(' <i class="k-s3"></i>', tr[["ab_rdo"]], " ", tr[["comp_rdo"]]) else "", '</p>')
  fig <- function(cap, svg, note) paste0('<figure class="fig"><figcaption>', cap, '</figcaption>', legend, svg,
                                         '<p class="note">', note, '</p></figure>')

  th_num <- function(x, title) sprintf('<th class="r" title="%s">%s</th>', .fr_esc(title), x)
  fam_rows <- paste(sprintf('<tr><td>%s</td><td class="r">%d</td><td class="r">%d</td><td class="r">%d</td><td class="r">%s</td><td class="r">%s</td><td class="r">%s</td><td class="r b">%s</td></tr>',
                            .fr_esc(fa$family), as.integer(fa$abundance), as.integer(fa$n_species), as.integer(fa$n_subplots),
                            .fr_num_cell(fa$relative_density, 2), .fr_num_cell(fa$relative_frequency, 2),
                            .fr_num_cell(fa$relative_dominance, 2), .fr_num_cell(fa$importance_value, 2)), collapse = "")
  fam_tbl <- paste0('<table class="tbl phy"><caption>', tr[["tab_fam"]], '</caption><thead><tr><th>', tr[["family"]], '</th>',
                    th_num("N", tr[["n_stems"]]), th_num(tr[["col_spp"]], tr[["k_species"]]), th_num(tr[["col_sub"]], tr[["subplot"]]),
                    th_num(tr[["ab_rd"]], tr[["comp_rd"]]), th_num(tr[["ab_rf"]], tr[["comp_rf"]]), th_num(tr[["ab_rdo"]], tr[["comp_rdo"]]),
                    th_num("IVI", "IVI"), '</tr></thead><tbody>', fam_rows, '</tbody></table>')
  sp_rows <- paste(sprintf('<tr><td><i>%s</i></td><td>%s</td><td class="r">%d</td><td class="r">%d</td><td class="r">%s</td><td class="r">%s</td><td class="r">%s</td><td class="r">%s</td><td class="r">%s</td><td class="r b">%s</td></tr>',
                           .fr_esc(sp$species), .fr_esc(sp$family), as.integer(sp$abundance), as.integer(sp$n_subplots),
                           .fr_num_cell(sp$absolute_frequency, 1), .fr_num_cell(sp$relative_density, 2), .fr_num_cell(sp$relative_frequency, 2),
                           .fr_num_cell(sp$relative_dominance, 2), .fr_num_cell(sp$basal_area_m2, 3), .fr_num_cell(sp$importance_value, 2)), collapse = "")
  sp_tbl <- paste0('<table class="tbl phy"><caption>', tr[["tab_spp"]], '</caption><thead><tr><th>', tr[["species"]], '</th><th>', tr[["family"]], '</th>',
                   th_num("N", tr[["n_stems"]]), th_num(tr[["col_sub"]], tr[["subplot"]]), th_num(paste0(tr[["ab_af"]], " %"), tr[["ab_af"]]),
                   th_num(tr[["ab_rd"]], tr[["comp_rd"]]), th_num(tr[["ab_rf"]], tr[["comp_rf"]]), th_num(tr[["ab_rdo"]], tr[["comp_rdo"]]),
                   th_num(paste0(tr[["ab_ba"]], " (m\u00b2)"), tr[["k_ba"]]), th_num("IVI", "IVI"), '</tr></thead><tbody>', sp_rows, '</tbody></table>')

  nf <- min(n_fig_fam, nrow(fa)); ns <- min(n_fig_spp, nrow(sp))
  dbh_svg <- .fr_dbh_svg(dbh_cm, tr)
  paste0(
    '<section class="page phyto" id="phyto"><h2>', tr[["phyto_title"]], '</h2><p class="lede">', tr[["phyto_lede"]], '</p>',
    strip, '<ol class="fn">', paste(sprintf("<li>%s</li>", fns), collapse = ""), '</ol>',
    fig(tr[["fig_fam"]], .fr_ivi_svg(fa, "family", tr, nf, FALSE, tr[["fig_fam"]]), sprintf(tr[["top_note"]], nf)),
    '</section>',
    '<section class="page phyto" id="phyto-2">',
    fig(tr[["fig_spp"]], .fr_ivi_svg(sp, "species", tr, ns, TRUE, tr[["fig_spp"]]), sprintf(tr[["top_note"]], ns)),
    if (nzchar(dbh_svg)) paste0('<figure class="fig"><figcaption>', tr[["fig_dbh"]], '</figcaption>', dbh_svg, '</figure>') else "",
    '</section>',
    '<section class="page phyto" id="phyto-tables">', fam_tbl, sp_tbl, '</section>')
}


# Logos ------------------------------------------------------------------------

#' Official logos shipped with the package (inst/logos)
#'
#' @return Named list; each element has `file`, `alt` and `url`.
#' @keywords internal
#' @noRd
.report_logo_catalog <- function() {
  list(
    forplotr = list(file = "forplotR.png", alt = "forplotR",
                    url = "https://dboslab.github.io/forplotR-website/"),
    jbrj = list(file = "jbrj.png",
                alt = "Instituto de Pesquisas Jardim Bot\u00e2nico do Rio de Janeiro (JBRJ)",
                url = "https://www.gov.br/jbrj/pt-br"),
    ci = list(file = "ci.png", alt = "Conserva\u00e7\u00e3o Internacional Brasil (CI)",
              url = "https://brasil.conservation.org/"),
    iakio = list(file = "iakio.png", alt = "Associa\u00e7\u00e3o Iaki\u00f4 Panar\u00e1",
                 url = "https://www.instagram.com/associacao.iakio/"),
    monitora = list(file = "monitora.png", alt = "Programa MONITORA (ICMBio)",
                    url = "https://www.gov.br/icmbio/pt-br/assuntos/monitoramento/programa-monitora"),
    forestplots = list(file = "forestplots.png", alt = "ForestPlots.net",
                       url = "https://forestplots.net/")
  )
}

#' Logo presets: a name that stands for several logos
#'
#' @keywords internal
#' @noRd
.report_logo_presets <- function() {
  list(
    default = c("forplotr", "jbrj"),
    panara = c("forplotr", "jbrj", "ci", "iakio"),
    monitora = c("forplotr", "jbrj", "monitora"),
    forestplots = c("forplotr", "jbrj", "forestplots")
  )
}

#' Resolve the `logos` argument into images ready to embed
#'
#' @param logos `NULL` (automatic), `FALSE`/`"none"`/`character(0)` (no logos),
#'   or a character vector mixing logo names (`"forplotR"`, `"jbrj"`, `"ci"`,
#'   `"iakio"`, `"monitora"`, `"forestplots"`), presets (`"default"`,
#'   `"panara"`, `"monitora"`, `"forestplots"`) and paths to image files.
#' @param input_type Canonical input type; with `logos = NULL`, the Indigenous
#'   Land field sheet gets the `"panara"` preset.
#' @param auto What `logos = NULL` stands for when `input_type` does not decide
#'   it (a preset or logo name).
#' @return Data frame with `src` (data URI), `alt` and `url`; zero rows for none.
#' @keywords internal
#' @noRd
.resolve_report_logos <- function(logos = NULL, input_type = "field_sheet",
                                  auto = "default") {
  empty <- data.frame(src = character(), alt = character(), url = character(),
                      stringsAsFactors = FALSE)
  presets <- .report_logo_presets()
  catalog <- .report_logo_catalog()

  if (is.null(logos)) {
    logos <- if (identical(input_type, "field_sheet_ti")) "panara" else auto
  }
  if (isFALSE(logos)) return(empty)
  if (!is.character(logos) || anyNA(logos)) {
    stop("`logos` must be NULL, FALSE or a character vector of logo names, presets ",
         "and image file paths.", call. = FALSE)
  }
  logos <- trimws(logos)
  logos <- logos[nzchar(logos)]
  if (!length(logos) || identical(tolower(logos), "none")) return(empty)

  # presets expand in place; a logo named twice is shown once. "monitora" and
  # "forestplots" are both a logo and a preset: alone they are the preset (the
  # partner logo next to forplotR and JBRJ), in a list they are just the logo.
  items <- unlist(lapply(logos, function(x) {
    key <- tolower(x)
    is_preset <- key %in% names(presets) && (length(logos) == 1L || !(key %in% names(catalog)))
    if (is_preset) presets[[key]] else x
  }), use.names = FALSE)
  items <- items[!duplicated(tolower(items))]

  rows <- lapply(items, function(x) {
    key <- tolower(x)
    if (key %in% names(catalog)) {
      lg <- catalog[[key]]
      f <- system.file("logos", lg$file, package = "forplotR")
      if (!nzchar(f)) return(NULL)
      return(data.frame(src = .file_data_uri(f), alt = lg$alt, url = lg$url,
                        stringsAsFactors = FALSE))
    }
    if (file.exists(x) && !dir.exists(x)) {
      ext <- tolower(tools::file_ext(x))
      if (!ext %in% c("png", "svg", "jpg", "jpeg", "gif", "webp")) {
        stop("Logo file '", x, "' must be a PNG, SVG, JPEG, GIF or WebP image.",
             call. = FALSE)
      }
      alt <- gsub("[_-]+", " ", tools::file_path_sans_ext(basename(x)))
      return(data.frame(src = .file_data_uri(x), alt = alt, url = "",
                        stringsAsFactors = FALSE))
    }
    stop("Unknown logo '", x, "'. Use one of: ",
         paste(c("forplotR", "jbrj", "ci", "iakio", "monitora", "forestplots"), collapse = ", "),
         "; a preset (", paste(names(presets), collapse = ", "),
         "); \"none\"; or the path to an image file.", call. = FALSE)
  })
  rows <- rows[!vapply(rows, is.null, logical(1))]
  if (!length(rows)) return(empty)
  do.call(rbind, rows)
}

#' Logo strip of the report cover
#'
#' @keywords internal
#' @noRd
.fr_logos_html <- function(logos) {
  if (is.null(logos) || !nrow(logos)) return("")
  img <- sprintf('<img src="%s" alt="%s">', logos$src, .fr_esc(logos$alt))
  has_url <- nzchar(logos$url)
  img[has_url] <- sprintf('<a href="%s" target="_blank" rel="noopener">%s</a>',
                          .fr_esc(logos$url[has_url]), img[has_url])
  paste0('<div class="logos">', paste(img, collapse = ""), '</div>')
}


# HTML -------------------------------------------------------------------------

#' Build the field report as one HTML string
#'
#' @param lay Layout from `.field_report_layout()`.
#' @param meta List with `plot_name`, `plot_code`, `plot_census_no_fp`, `team`
#'   and, optionally, `census_date` (a `Date`).
#' @param logos Data frame from `.resolve_report_logos()`, or `NULL` for none.
#' @keywords internal
#' @noRd
.field_report_html <- function(lay, meta, language = "en", highlight_palms = TRUE,
                               plot_area_ha = NULL, phyto = NULL, logos = NULL) {
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
  date_txt <- if (inherits(meta$census_date, "Date") && length(meta$census_date) == 1L && !is.na(meta$census_date)) {
    paste0(" \u00b7 ", tr[["date"]], ": ", format(meta$census_date, "%d/%m/%Y"))
  } else {
    ""
  }

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
    .fr_logos_html(logos),
    '<header class="doc-head">',
    sprintf('<p class="kicker">%s%s</p>', tr[["doc_title"]], census),
    sprintf('<h1>%s <span>%s</span></h1>', pn, pc),
    sprintf('<p class="meta">%s%s%s</p>', size_txt, date_txt, if (nzchar(ver)) paste0(" \u00b7 forplotR ", ver) else ""),
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

  # floristics and phytosociology (after the subplot pages, before the appendix)
  area_ha <- if (lay$kind == "monitora") 0.4 else plot_area_ha
  phyto_html <- if (!is.null(phyto)) .fr_phyto_html(phyto, d$dbh, area_ha, tr) else ""

  first_lbl <- slab(c$label[min(nrow(c), if (lay$kind == "grid") 37 else 3)])
  paste0(
    '<!doctype html><html lang="', lang_attr, '"><head><meta charset="utf-8">',
    '<meta name="viewport" content="width=device-width,initial-scale=1">',
    '<meta name="generator" content="forplotR ', ver, '">',
    '<title>', pc, if (nzchar(pc)) " \u00b7 " else "", tr[["doc_title"]], '</title>',
    '<style>', .fr_read_asset("field_report.css"), '</style></head>',
    sprintf('<body data-not-found="%s" data-sub-prefix="%s">', .fr_esc(tr[["not_found"]]), sub_prefix),
    '<div class="bar"><a href="#summary" class="brand">', if (nzchar(pc)) pc else tr[["index"]], '</a>',
    '<a href="#map-all">', tr[["maps"]], '</a><a href="#route">', tr[["route_short"]], '</a>', if (nzchar(phyto_html)) paste0('<a href="#phyto">', tr[["phyto_short"]], '</a>') else "", '<a href="#checklist">', tr[["checklist"]], '</a>',
    sprintf('<form class="find" onsubmit="return false"><input id="find" type="search" placeholder="%s" aria-label="%s" autocomplete="off"></form>',
            .fr_esc(sprintf(tr[["find_ph"]], first_lbl)), .fr_esc(tr[["find_label"]])),
    sprintf('<button class="print" type="button" onclick="window.print()" title="%s">PDF</button></div>', .fr_esc(tr[["print_pdf"]])),
    .fr_mini_symbol(lay),
    '<main>', summary_html, maps_html, route_html, subs_html, phyto_html, chk_html, '</main>',
    '<script>', .fr_read_asset("field_report.js"), '</script></body></html>')
}

#' Write the field report
#'
#' @return The path to the HTML file, invisibly.
#' @keywords internal
#' @noRd
.write_field_report <- function(file, fp_coords, input_type, subplot_size,
                                plot_width_m, plot_length_m, highlight_palms,
                                meta, language, plot_area_ha = NULL, phyto = NULL,
                                logos = NULL) {
  lay <- .field_report_layout(fp_coords, input_type, subplot_size,
                              plot_width_m, plot_length_m, highlight_palms)
  html <- .field_report_html(lay, meta, language, highlight_palms, plot_area_ha, phyto,
                             logos = logos)
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
