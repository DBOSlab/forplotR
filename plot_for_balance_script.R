# Plot Forest Balance - Main Script
# Authors: Giulia Ottino & Domingos Cardoso

#### Load Required Packages ####
library(forplotR)

#### ForestPlots / Field Sheet ####

# 1 ha plot (100 x 100 m) - Panará

mk_voucher_dirs(
  fp_file_path = "data/SKR_01.xlsx",
  input_type = "field_sheet_ti",
  output_dir = "voucher_imgs_SKR01"
)


plot_html_map(
  fp_file_path = "data/SKR_01.xlsx",
  input_type = "field_sheet_ti",
  vertex_coords = "data/vertices_SKR_01.xlsx",
  voucher_imgs = "voucher_imgs_SKR01",
  collector = "D.Cardoso",
  herbaria_lookup = FALSE,
  herbaria = "RB",
  filename = "SKR01_plot_map",
  ortho_image = "data/Parcela_Panara_True_Ortho.tif",
  ortho_opacity = 0.7
)




mk_voucher_dirs(
  fp_file_path = "data/RUS_01.xlsx",
  input_type = "field_sheet",
  output_dir = "voucher_imgs_RUS01"
)


plot_html_map(
  fp_file_path = "data/RUS_01.xlsx",
  input_type = "field_sheet",
  vertex_coords = "data/vertices.xlsx",
  voucher_imgs = "voucher_imgs_RUS01",
  collector = "G. C. Ottino",
  herbaria_lookup = TRUE,
  herbaria = "RB",
  filename = "rus_plot_map",
  verbose = TRUE
)






plot_for_balance(
  fp_file_path   = "data/SKR_01.xlsx",
  input_type     = "field_sheet_ti",
  language       = "pa",
  plot_size      = 1,
  subplot_size   = 10,
  highlight_palms = TRUE,
  render_html    = TRUE,
  render_pdf     = TRUE,
  write_xlsx     = TRUE,
  dir            = "Results_map_plot_Panara_PA",
  filename       = "SKR01_plot_specimen_PA"
)


plot_for_balance(
  fp_file_path   = "data/SKR_01.xlsx",
  input_type     = "field_sheet_ti",
  language       = "pt",
  plot_size      = 1,
  subplot_size   = 10,
  highlight_palms = TRUE,
  render_html    = TRUE,
  render_pdf     = TRUE,
  write_xlsx     = TRUE,
  verbose = TRUE,
  dir            = "Results_map_plot_Panara_PT",
  filename       = "SKR01_plot_specimen_PT"
)


plot_for_balance(
  fp_file_path   = "data/SKR_01.xlsx",
  input_type     = "field_sheet_ti",
  language       = "en",
  plot_size      = 1,
  subplot_size   = 10,
  highlight_palms = TRUE,
  render_html    = TRUE,
  render_pdf     = TRUE,
  write_xlsx     = TRUE,
  dir            = "Results_map_plot_Panara_EN",
  filename       = "SKR01_plot_specimen_EN"
)






fp_file_path   = "data/SKR_01.xlsx"
input_type     = "field_sheet_ti"
language       = "pt"
plot_size      = 1
subplot_size   = 10
highlight_palms = TRUE
render_html    = TRUE
render_pdf     = TRUE
write_xlsx     = TRUE
dir            = "Results_map_plot_Panara_PT__"
filename       = "SKR01_plot_specimen_PT"


files <- list.files("data/forestplot_ppbio_data_1Apr2026")
i=9
for (i in seq_along(files)) {

  filename <- gsub("__all_censuses|[.]xlsx", "", files[i])

  plot_for_balance(
    fp_file_path   = paste0("data/forestplot_ppbio_data_1Apr2026/", files[i]),
    input_type     = "fp_query_sheet",
    language       = "PT",
    plot_size      = 0.5,
    subplot_size   = 10,
    highlight_palms = TRUE,
    plot_census_no = 1,
    render_html    = TRUE,
    render_pdf     = TRUE,
    write_xlsx     = TRUE,
    dir            = "Results_map_plot_PPBio",
    filename       = filename
  )

}



