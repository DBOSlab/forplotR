utils::globalVariables(
  c(
    # From .build_fp_base_plot
    "global_x", "global_y", "center_x", "center_y", "T1", "diameter", "New Tag No",

    # From .build_monitora_base_plot
    "r", "xmin", "xmax", "ymin", "ymax", "col_from_center", "col_idx", "base_num",
    "cx", "cy", "label", "D", "lx", "ly", "lab", "hjust", "vjust",

    # From .check_point_in_admin
    "stri_trim_both", "crs", "relate",

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
    "read.csv", "PlotCode", "AGBind",

    # From .herbaria_lookup_links
    ":=", "rn", "recordNumber", "fp", "recordedBy", "ln", "key", "na.omit",

    # From fp_herb_converter
    "column_map",

    # From plot_for_balance
    "in_dat", "D", "Family", "Original determination", "Species", "Species_fmt",
    "n_subplots_per_col", "col_index", "row_index", "subplot_x", "subplot_y",
    "center_x", "center_y", "setNames", "subunit_letter", "x10", "y10",
    "Status", "X", "Y", ".monitora_to_cell_coords",

    # From plot_html_map
    "latitude", "longitude",

    NULL
  )
)
