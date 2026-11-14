#' @name RUS01
#'
#' @docType data
#'
#' @title ForestPlots.net field sheet of plot RUS01 (REBIO Augusto Ruschi, Brazil)
#'
#' @description
#' Census data of the 1-ha permanent plot RUS01 in the Reserva Biologica
#' Augusto Ruschi (Santa Teresa, Espirito Santo, Brazil; Atlantic Forest),
#' stored exactly as a ForestPlots.net standard field sheet: the first row of
#' the sheet holds plot metadata and the second row holds the column headers.
#' Write it to an Excel file with \code{openxlsx::write.xlsx()} to use it as
#' \code{fp_file_path} with \code{input_type = "field_sheet"} in every
#' \code{forplot_*()} function.
#'
#' @format A data frame with 938 rows and 29 character columns. Column names
#' carry the plot metadata (\code{"Plotcode: RUS01"},
#' \code{"Plot Name: REBIO Augusto Ruschi"}, \code{"Date: 26/11/2014"},
#' \code{"Team: ..."}); the first row holds the field-sheet headers and the
#' remaining 937 rows are trees with DBH >= 10 cm, distributed in 100
#' subplots of 10 x 10 m. Main fields:
#' \itemize{
#'   \item \strong{New Tag No}: Tree tag number.
#'   \item \strong{New Stem Grouping}: Stem group identifier for multi-stemmed trees.
#'   \item \strong{T1}: Subplot number (1--100).
#'   \item \strong{X, Y}: Local coordinates within the subplot (m).
#'   \item \strong{Family, Original determination, Morphospecies}: Taxonomic identification.
#'   \item \strong{D}: Diameter at the point of measurement (mm).
#'   \item \strong{POM}: Height of the point of measurement (mm).
#'   \item \strong{Flag1--Flag3, LI, CI}: ForestPlots.net condition and light codes.
#'   \item \strong{Voucher, Silica, Collected}: Voucher number and collection status.
#'   \item \strong{Census Notes}: Free-text field notes.
#' }
#'
#' @source Field census by G. C. Ottino, D. Cardoso and collaborators
#' (November 2014).
#'
#' @usage data(RUS01)
#'
#' @examples
#' dim(RUS01)
#' # Field-sheet headers are stored in the first row
#' unlist(RUS01[1, 1:8])
#'
#' # Use it as an input file
#' fp <- file.path(tempdir(), "RUS01.xlsx")
#' openxlsx::write.xlsx(RUS01, fp)
#'
NULL

#' @name TUM
#'
#' @docType data
#'
#' @title MONITORA forest monitoring data from Parque Nacional Montanhas do Tumucumaque
#'
#' @description
#' Tree census data from three sampling stations (\code{"EA-1"},
#' \code{"EA-2"}, \code{"EA-3"}) of the MONITORA biodiversity monitoring
#' program (ICMBio) in the Parque Nacional Montanhas do Tumucumaque (Amapa,
#' Brazil; Amazonia), in the original MONITORA spreadsheet layout. Write it to
#' an Excel file with \code{openxlsx::write.xlsx()} to use it as
#' \code{fp_file_path} with \code{input_type = "monitora"}.
#'
#' @format A data frame with 1675 rows and 47 columns. Main fields:
#' \itemize{
#'   \item \strong{CDUC, NOMEUC, bioma}: Protected-area code, name and biome.
#'   \item \strong{N_estacao, Nome_estacao}: Station number and name.
#'   \item \strong{N_UA}: Sampling unit number.
#'   \item \strong{Data, Ano}: Census date (Excel serial number) and year (2014--2021).
#'   \item \strong{nome_coletores}: Field team.
#'   \item \strong{Subunidades, N_parcela}: Subunit (Norte, Sul, Leste, Oeste) and
#'   subplot number.
#'   \item \strong{N_arvore}: Tree tag number.
#'   \item \strong{CAP_1--CAP_8, cap_tot}: Circumference at breast height of each stem
#'   and total (cm).
#'   \item \strong{altura}: Tree height (m).
#'   \item \strong{arvore_morta}: Whether the tree is dead.
#'   \item \strong{Familia, Genero, Especie}: Taxonomic identification (column names
#'   carry Portuguese accents in the data).
#'   \item \strong{individuo.coletado}: Whether a voucher was collected.
#'   \item \strong{voucher/Coletor, voucher/numero, N.tombo}: Voucher collector,
#'   collector number and herbarium accession number.
#'   \item \strong{X, Y}: Local coordinates within the subplot (m).
#' }
#'
#' @source MONITORA program, Instituto Chico Mendes de Conservacao da
#' Biodiversidade (ICMBio).
#'
#' @usage data(TUM)
#'
#' @examples
#' dim(TUM)
#' table(TUM$Nome_estacao)
#'
NULL
