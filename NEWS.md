# forplotR (development version)

- The species chart of `forplot_balance()` is now two charts: the species with
  the most individuals (the simple ranking used before the IVI chart, now with
  the first 15 species and the same taxon rule as the IVI: a name within its
  family) followed by the species with the highest importance value (IVI).
  Each has its own figure and caption in the field report (Figures 2 and 3;
  stems per diameter class is now Figure 4) and its own PNG file. Every bar of
  the species and family charts ends with its value (integer counts, IVI with
  one decimal), and the x axis is extended so the value is never cut.
- When the field sheet has a vernacular name (the `Local name` column of the
  Indigenous Land template, e.g. the Panará name, or a `Vernacular name`
  column), species are labelled "vernacular | scientific" in both charts, in
  the PNG files and in the report, with the scientific name in italics.
  Without a vernacular name, only the scientific name is shown. In Panará,
  the dictionary's "most abundant species" now titles the ranking; the IVI
  species chart has no equivalent in the dictionary and stays in Portuguese.
- Output files of `forplot_balance()` are named in the language of the call,
  from one dictionary of file names (`R/i18n_files.R`): Portuguese for `"pt"`
  and `"pa"`, Spanish, French, and English for `"en"` and `"ma"` (Chinese
  characters cannot be used in file names). Names are snake_case ASCII. Old
  and new names (English): `_full_report` (unchanged), `_general`,
  `_collected`, `_uncollected`, `_uncollected_palms`,
  `_priority_uncollected_species`, `_dbh_classes` (unchanged);
  `_species_plot` -> `_species_ivi`; `_family_plot` -> `_families_ivi`;
  new `_species_abundance`. In Portuguese: `_relatorio_completo`, `_geral`,
  `_coletados`, `_nao_coletados`, `_palmeiras_nao_coletadas`,
  `_especies_prioritarias_nao_coletadas`, `_especies_abundancia`,
  `_especies_ivi`, `_familias_ivi`, `_classes_dap`, `_balanco_coleta`.
- The Excel summary now carries the prefix too
  (`<filename>_collection_balance.xlsx`, `<filename>_balanco_coleta.xlsx`),
  so two runs in the same folder no longer overwrite each other.
- `filename` now defaults to `NULL`, which uses the word for "plot" in the
  report language (`plot`, `parcela`, `parcelle`) instead of
  `"plot_specimen"`. Calls that pass `filename` are not affected. With
  several MONITORA stations, the station suffix follows the language too
  (`_station_`, `_estacao_`, `_estacion_`).
- A PDF printed from the HTML report in a browser is offered with the name of
  the report file (e.g. `RUS01_relatorio_completo.pdf`) instead of the page
  title ("RUS01 · Balanço de coleta"); the tab keeps the title.
- `forplot_balance()` reads GigaFlor network field sheets with
  `input_type = "gigaflor"` (protocol 1: a 1-ha plot of 25 subplots of
  20 x 20 m in `T1`, each split into four 10 x 10 m quadrants in `T2`).
  `subplot_size` then defaults to 20 m. The sheet is read from its `"Data"`
  worksheet; `"Tag Number"` is the tag, `"Species"` the determination and
  `"Main Stem Tag"` groups the stems of a multi-stemmed tree. `D` is in
  centimetres (converted to millimetres internally, or taken from
  `"CAP (cm)"` / pi when empty). `X` and `Y` are the tree's longitude and
  latitude in decimal degrees: they are projected to metres and fitted onto
  the subplot grid by a rotation and translation (which recovers the
  orientation of the plot on the ground), then drawn at that position inside
  each tree's own subplot. A warning is given when the coordinates fit the
  grid poorly. In the field report, the "X; Y" column of the subplot tables
  shows the original longitude; latitude in decimal degrees. A `"No"` in
  `Collected` (the template's Yes/No list) is not counted as collected.
- `forplot_map()`, `forplot_sheets()` and `forplot_vouchers()` also accept
  `input_type = "gigaflor"`. `forplot_map()` draws each tree at its own
  coordinates; `vertex_coords` becomes optional, the plot corners (and so the
  subplot grid) being estimated from the trees' coordinates, their `T1`
  subplots and the plot dimensions (give `vertex_coords` to use the corners
  marked in the survey). `forplot_sheets()` writes each tree's latitude and
  longitude on its label (`lat`/`long` become optional), reads `D` in cm,
  and takes the collector, additional collectors and collection date from
  the `"Team:"` and `"Date:"` cells; label notes no longer print
  `x = ...m, y = ...m` when the position within the subplot is unknown.
  `forplot_sheets()` now stops with a clear message when no specimen is
  marked as collected.
- `forplot_crossdet()`: `herbaria` is now optional. With `herbaria = NULL`
  (the default) the herbarium of each voucher is read from the inventory
  column `collectionCode` (Darwin Core; e.g. `"RB"`, as in the GigaFlor
  sheet), only those herbaria are downloaded, and each voucher is checked
  only in its own herbarium. Vouchers without a `collectionCode` are left out
  with a message. Archives are downloaded with
  `jabotR::jabot_download(herbarium = <codes>)` when jabotR is installed
  (now in Suggests), falling back on the package's own JABOT downloader
  (also used with `force_refresh = TRUE` or `resource_map`). An explicit
  `herbaria` (now one or more codes) keeps the previous behaviour: every
  voucher is checked there. `input_type = "gigaflor"` is accepted, and the
  `unmatched_inventory` table gains a `herbarium` column.

- The field report in Panará (`language = "pa"`) is no longer all in
  Portuguese: its labels (title, team, date, individuals, families, with
  voucher, pending, palms, subplot, tag, DBH, species, checklist and others)
  now come from the package's Panará dictionary, as do the titles of the
  figures and tables of the floristics section (reworded to what the
  dictionary has: the IVI figures are titled "most common families" and "most
  abundant species"). Explanatory sentences and the labels the dictionary
  does not have stay in Portuguese.

- `read_plot_header()` lists the columns of a field sheet from its detected
  header row, with the exact name each column has when the sheet is read
  (the name a `colname_*` argument must give). Rows of plot metadata above
  the header are never listed. An empty header cell is `NA_col_<position>`
  and a repeated header gets `.1`, `.2`. The field-sheet readers now name
  their columns with this same rule.

- Field sheets that start with rows of plot metadata are now read from their
  real header row. `detect_header_row()` scores the first rows of a sheet by
  the column names it recognises and returns the header row;
  `read_plot_metadata()` turns the `"Label | Translation: value"` cells above
  it into `plot_code`, `plot_name`, `date` (day/month/year first) and `team`
  (one element per person); `read_plot_sheet()` reads the table from that row.
  `forplot_balance()`, `forplot_map()`, `forplot_vouchers()` and
  `forplot_sheets()` gain `header_row = NULL` (automatic) to force the row.
- Bilingual column names are understood: either side of
  `"Name | Translation"` is matched, ignoring case and accents, with English,
  Portuguese and Panará labels (`"Família botânica | Ksyapiâhapiâra pârikiã"`
  is `Family`, `"D | Wy"` is `D`). The Indigenous Land (Panará) field sheet is
  now mapped by column name under both `input_type = "field_sheet_ti"` and
  `"field_sheet"`, with the fixed column order as a fallback. When that sheet
  has separate columns for the scientific and the Panará name, the
  determination is the scientific name and the Panará name is kept in
  `Local name`.
- When two columns read as the same name (the Panará sheet has `"D | Wy"` and
  a helper column `"D"`), the first one in the sheet is used and the other is
  kept as `D_2`; a message says which is which.
- The census date read from the sheet is shown in the field report.
- The field report (HTML and PDF) has a header with official logos again, and
  the map can show them too. `forplot_balance()` and `forplot_map()` gain
  `logos`: logo names, presets and paths to your own image files, in the
  order they should appear. Logos shipped in `inst/logos/`: `"forplotR"`,
  `"jbrj"`, `"ci"` (Conservação Internacional), `"iakio"` (Associação Iakiô
  Panará), `"monitora"` (Programa MONITORA) and `"forestplots"`
  (ForestPlots.net). Presets: `"default"` (forplotR + JBRJ), `"panara"`
  (plus CI and Iakiô), `"monitora"` and `"forestplots"` (the two standard
  logos plus that partner). In the report the default is forplotR + JBRJ,
  or `"panara"` for `input_type = "field_sheet_ti"`; in the map it is the
  forplotR logo alone. `FALSE` shows none. Logos keep their proportions and
  share one height.
- `LICENSE.md` with the full text of the MIT licence.
- `inst/extdata/panara_field_sheet_template.xlsx`: the blank field-sheet
  template of the Panará Indigenous Land (metadata labels and bilingual
  header, no data), used in the examples.

- The field report of `forplot_balance()` ends with a floristic and
  phytosociological overview: richness, density and basal area per hectare,
  Shannon, Simpson and Pielou indices, the families and species with the
  highest importance value (IVI, stacked by relative density, frequency and
  dominance), stems per diameter class, and full tables of phytosociological
  parameters for every family and species. The PNG charts use the same
  palette, and the diameter classes are also saved
  (`<filename>_dbh_classes.png`).
- Phytosociological parameters now treat a name used in more than one family
  (e.g. "Indet indet") as one taxon per family instead of merging them.
- Pedro Lage Viana (Instituto Nacional da Mata Atlântica) added as thesis
  co-advisor.

- `forplot_balance()` writes a new field report: a single self-contained HTML
  file for use on a phone (offline) or printed, one A4 page per subplot. It
  has a summary of the collection balance with footnotes on how each number
  is computed, a clickable subplot grid, whole-plot maps with the tag of every
  stem (all stems, with voucher, pending and palms), a priority route, subplot
  pages with a tree table and blank columns for remeasurement, and a
  checklist appendix. MONITORA plots are drawn as the cross of four arms.
  Maps are inline SVG instead of plotly widgets, so the report is several
  times smaller (RUS01: about 1 MB instead of 5.8 MB) and builds in seconds.
- The PDF report is the same HTML printed by a headless Chrome, Chromium or
  Edge (found automatically or through `FORPLOTR_CHROME`); LaTeX, `tinytex`
  and pandoc are no longer used by `forplot_balance()`. Without a browser the
  PDF is skipped with a warning.
- `plotly` is no longer a dependency; `markdown` and `tinytex` are no longer
  suggested.

# forplotR 1.0.0

## Initial release

`forplotR` streamlines the management of permanent forest plot data from
ForestPlots.net field sheets, Query Library exports and MONITORA
spreadsheets.

### Functions

- `forplot_map()`: interactive HTML map of tagged trees, with voucher
  photographs, optional JABOT herbarium lookup, and filters for taxonomy,
  collection status and image availability.
- `forplot_balance()`: plot- and subplot-level collection-balance maps,
  HTML/PDF report, species checklist and Excel summary of collection
  completeness.
- `forplot_vouchers()`: creates and updates voucher image folders organized
  by updated family and species.
- `forplot_sheets()`: fills herbarium label spreadsheets (JABOT, BRAHMS or a
  custom layout) from field data, with multilingual notes.
- `forplot_crossdet()`: cross-checks field determinations against a JABOT
  herbarium Darwin Core Archive.

### Data

- `RUS01`: ForestPlots.net field sheet of the 1-ha plot RUS01 (REBIO Augusto
  Ruschi, Espirito Santo, Brazil).
- `TUM`: MONITORA data from the Parque Nacional Montanhas do Tumucumaque
  (Amapa, Brazil).

### Behaviour

- Output files are written to `dir` (or `output_dir`); by default they go to a
  session temporary directory, so nothing is written to the working directory
  unless a path is supplied. `forplot_crossdet()` no longer saves a workbook
  unless `save = TRUE`.
- Downloads of GADM boundaries and WorldClim elevation in `forplot_sheets()`
  fail gracefully with a warning when offline.
- `verbose = FALSE` now silences all informational messages.
- PDF rendering requires the suggested package 'tinytex'.
- 'terra', 'geodata', 'duckdb' and 'DBI' are now suggested packages, loaded
  only by the features that need them (locality validation and elevation in
  `forplot_sheets()`, herbarium lookup in `forplot_map()` and
  `forplot_crossdet()`). Without them those features are skipped with a
  warning that says what to install, and `library(forplotR)` loads faster.
- The HTML report of `forplot_balance()` no longer needs pandoc: without it,
  the report is rendered with knitr and the suggested package 'markdown',
  with the same sections, table of contents, interactive plots and figures,
  in a single HTML file.
- `forplot_map()` no longer needs pandoc: without it, the map's scripts,
  styles and images are embedded directly, giving the same single-file HTML.
- New option `options(forplotR.offline = TRUE)`: no function accesses the
  internet. Locality validation, elevation download and herbarium downloads
  are skipped with a warning; already cached herbarium data is still used.

### Infrastructure

- MIT license.
- Vignette with a complete worked example using the bundled data.
- GitHub Actions: R-CMD-check and test coverage.
- Hosted documentation: [forplotR-website](https://dboslab.github.io/forplotR-website/).

Please report bugs at <https://github.com/DBOSlab/forplotR/issues>.
