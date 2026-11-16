# forplotR (development version)

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
