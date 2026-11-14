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

### Infrastructure

- MIT license.
- Vignette with a complete worked example using the bundled data.
- GitHub Actions: R-CMD-check and test coverage.
- Hosted documentation: [forplotR-website](https://dboslab.github.io/forplotR-website/).

Please report bugs at <https://github.com/DBOSlab/forplotR/issues>.
