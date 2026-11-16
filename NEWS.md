# forplotR (development version)

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
