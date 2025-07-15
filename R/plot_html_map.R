#' Plot and save vouchered forest plot map as HTML
#'
#' @author Giulia Ottino & Domingos Cardoso
#'
#' @description Generates an interactive Leaflet map of a forest plot with
#' coordinates transformed from local subplot coordinates to geographic
#' coordinates, using four corner vertices. The polygon is built using angular
#' ordering to prevent self-intersecting shapes. The function saves the resulting
#' map as a date-stamped self-contained HTML file in working directory.
#'
#' @usage plot_html_map(fp_file_path = NULL,
#'                      vertex_coords = NULL,
#'                      plot_size = 1,
#'                      subplot_size = 10,
#'                      map_type = c("satellite", "street"),
#'                      voucher_imgs = NULL,
#'                      dir = "Results_plot_map_html",
#'                      filename = "plot_map.html")
#'
#' @param fp_file_path Character. Path to the Excel file containing tree data.
#'
#' @param vertex_coords Data frame or path to Excel file with columns `Latitude`
#' and `Longitude` for the four plot corners.
#'
#' @param plot_size Numeric. Total plot size in hectares (default is 1).
#'
#' @param subplot_size Numeric. Subplot size in meters (default is 10).
#'
#' @param dir Character. Directory to save the output HTML file
#' (default is "Results_plot_map_html").
#'
#' @param filename Character. Name of the output HTML file (default is
#' "plot_map.html").
#'
#' @param map_type Character. Base map type: "satellite" or "street" (default
#' is "satellite").
#'
#' @param voucher_imgs Character. Directory path where voucher images are stored.
#'
#' @examples
#' \dontrun{
#' vertex_df <- data.frame(
#'   Latitude = c(-3.123, -3.123, -3.125, -3.125),
#'   Longitude = c(-60.012, -60.010, -60.012, -60.010)
#' )
#' plot_html_map("data/tree_data.xlsx", vertex_coords = vertex_df)
#' }
#'
#' @importFrom leaflet leaflet addProviderTiles addPolygons addCircleMarkers setView addControl
#' @importFrom htmltools HTML
#' @importFrom htmlwidgets onRender prependContent saveWidget
#' @importFrom geosphere destPoint bearing
#' @importFrom readxl read_excel
#' @importFrom here here
#' @importFrom dplyr select filter mutate case_when rename_with rename if_else
#' @importFrom dplyr %>%
#' @importFrom scales rescale
#'
#' @export
#'

plot_html_map <- function(fp_file_path = NULL,
                          vertex_coords = NULL,
                          plot_size = 1,
                          subplot_size = 10,
                          map_type = c("satellite", "street"),
                          voucher_imgs = NULL,
                          dir = "Results_plot_map_html",
                          filename = "plot_map.html") {

  # Check input file
  if (!file.exists(fp_file_path)) {
    stop("The provided fp_file_path does not exist.")
  }

  # Read the Excel file without headers to process metadata and headers correctly
  fp_sheet_raw <- suppressMessages(readxl::read_excel(fp_file_path, sheet = 1, col_names = FALSE))

  # Extract metadata from the first row (Team and Plot Name)
  metadata_row <- fp_sheet_raw[1, ]
  team <- metadata_row %>%
    select(where(~any(grepl("^Team[:]", ., ignore.case = TRUE)))) %>%
    as.character() %>%
    sub("^Team:\\s*", "", ., ignore.case = TRUE)
  plot_name <- metadata_row %>%
    select(where(~any(grepl("^Plot Name[:]", ., ignore.case = TRUE)))) %>%
    as.character() %>%
    sub("^Plot Name:\\s*", "", ., ignore.case = TRUE)
  plot_code <- metadata_row %>%
    select(where(~any(grepl("^Plotcode[:]", ., ignore.case = TRUE)))) %>%
    as.character() %>%
    sub("^Plotcode:\\s*", "", ., ignore.case = TRUE)

  # Extract headers from the second row
  header_row <- as.character(fp_sheet_raw[2, ])
  header_row[is.na(header_row)] <- paste0("NA_col_", seq_along(header_row))[is.na(header_row)]
  colnames(fp_sheet_raw) <- make.unique(header_row)

  # Data starts from the third row
  fp_sheet <- fp_sheet_raw[-(1:2), ]

  # Clean and convert necessary columns with suppressed warnings
  fp_clean <- fp_sheet %>%
    mutate(
      T1 = suppressWarnings(as.numeric(T1)),
      X = suppressWarnings(as.numeric(X)),
      Y = suppressWarnings(as.numeric(Y)),
      D = suppressWarnings(as.numeric(D))
    ) %>%
    filter(!is.na(T1), !is.na(X), !is.na(Y))

  # Geometry calculations
  max_coord <- plot_size * 100
  n_rows <- floor(max_coord / subplot_size)

  # Convert local to global coordinates
  fp_coords <- fp_clean %>%
    mutate(
      col = floor((T1 - 1) / n_rows),
      row = (T1 - 1) %% n_rows,
      global_x = col * subplot_size + X,
      global_y = if_else(
        col %% 2 == 0,
        row * subplot_size + Y,
        (n_rows - row - 1) * subplot_size + Y
      )
    )

  # Read vertex_coords if given as path
  if (is.character(vertex_coords) && grepl("\\.xlsx?$", vertex_coords)) {
    vertex_coords <- readxl::read_excel(vertex_coords)
  }

  # Standardize and clean vertex coordinates
  vertex_coords <- vertex_coords %>%
    rename_with(~tolower(gsub("\\s*\\(.*?\\)", "", .x)), everything()) %>%
    rename(latitude = latitude, longitude = longitude) %>%
    mutate(
      longitude = as.numeric(gsub(",", ".", longitude)),
      latitude  = as.numeric(gsub(",", ".", latitude))
    )

  # Correct signals to ensure coordinates are in southern hemisphere
  vertex_coords <- vertex_coords %>%
    mutate(
      longitude = ifelse(longitude > 0, -longitude, longitude),
      latitude  = ifelse(latitude > 0, -latitude, latitude)
    )

  # Use coordinates after cleaning
  p1 <- c(vertex_coords$longitude[1], vertex_coords$latitude[1])
  p2 <- c(vertex_coords$longitude[2], vertex_coords$latitude[2])
  p3 <- c(vertex_coords$longitude[3], vertex_coords$latitude[3])
  p4 <- c(vertex_coords$longitude[4], vertex_coords$latitude[4])

  coords_geo <- mapply(.get_latlon, fp_coords$global_x, fp_coords$global_y)
  fp_coords$Latitude <- coords_geo[1, ]
  fp_coords$Longitude <- coords_geo[2, ]

  # Assign color by collection status
  fp_coords$color <- case_when(
    fp_coords$Family == "Arecaceae" ~ "gold",
    !is.na(fp_coords$Collected) & fp_coords$Collected != "" ~ "gray",
    TRUE ~ "red"
  )

  fp_coords$popup <- paste0(
    "<b>Tag:</b> ", fp_coords[[ "New Tag No" ]], "<br/>",
    "<b>Subplot:</b> ", fp_coords$T1, "<br/>",
    "<b>Family:</b> ", fp_coords$Family, "<br/>",
    "<b>Species:</b> <i> ", fp_coords[["Original determination"]], "</i><br/>",
    "<b>DBH (mm):</b> ", round(fp_coords$D, 2), "<br/>",
    "<b>Voucher:</b> ", fp_coords$Voucher
  )

  plot_popup <- paste0("<b>Plot Name:</b> ",
                       plot_name, "<br/>",
                       "<b>Team:</b> ", team)

  # Define the base folder where the original images are stored
  original_image_base <- here(voucher_imgs)

  # List all subdirectories once to improve performance
  leaf_dirs <- list.dirs(original_image_base, recursive = TRUE, full.names = TRUE)

  # Keep only those that do not contain other directories (i.e., leaf directories)
  leaf_dirs <- leaf_dirs[!sapply(leaf_dirs, function(dir) {
    any(file.info(list.dirs(dir, recursive = FALSE))$isdir)
  })]

  leaf_dirs <- sub(paste0(".*(", voucher_imgs, "/.*)"), "\\1", leaf_dirs)

  for (i in seq_along(leaf_dirs)) {
    # List only image files in the current leaf directory
    img_files <- list.files(leaf_dirs[i], full.names = TRUE, pattern = "\\.(jpg|jpeg|png|gif)$", ignore.case = TRUE)

    # Skip if no images found
    if (length(img_files) == 0) next

    rel_paths <- file.path(leaf_dirs[i], basename(img_files))

    # Build HTML carousel slides
    slides <- paste0(
      "<div class='mySlides'>",
      "<a href='", rel_paths, "' target='_blank'>",
      "<img src='", rel_paths, "' style='max-width:200px; max-height:200px; display:block; margin:auto;'></a></div>",
      collapse = "\n"
    )

    # Match the voucher
    tf <- fp_coords$Voucher %in% basename(leaf_dirs[i])

    # Append popup content and flag for hasPhoto
    fp_coords$popup[tf] <- paste0(
      fp_coords$popup[tf],
      "<br/><b>HasPhoto:</b> yes<br/>",
      "<div class='slideshow-container' data-slide='1'>",
      slides,
      "</div>
    <div style='text-align:center;'>
      <a class='prev' onclick='plusSlides(-1)' style='margin-right:10px;'>&#10094;</a>
      <a class='next' onclick='plusSlides(1)'>&#10095;</a>
    </div>"
    )
  }

  base_tiles <- if (map_type == "satellite") "Esri.WorldImagery" else "OpenStreetMap"

  #_____________________________________________________________________________
  # Filter code in JavaScript + CSS
  filter_script <- htmltools::HTML("
function addFilterControl(el, x) {
  var map = this;

  // Create the filter container positioned in the top right corner
  var filterContainer = L.control({position: 'topright'});

  filterContainer.onAdd = function () {
    var div = L.DomUtil.create('div', 'info legend');

    // HTML filter structure: checkboxes and dropdowns
    div.innerHTML = `
      <div style='margin-bottom:5px;'>
        <b>Enable Filters:</b><br/>
        <label><input type='checkbox' id='toggleFamily'> Family</label><br/>
        <label><input type='checkbox' id='toggleSpecies'> Species</label><br/>
        <label><input type='checkbox' id='toggleColl'> Collection Status</label><br/>
        <label><input type='checkbox' id='togglePhotos'> With Photos Only</label>
      </div>
      <div id='familyFilterContainer' style='display:none; margin-bottom:5px;'>
        <label><b>Filter by Family:</b></label><br/>
        <select id='familyFilter' multiple size='5'></select>
      </div>
      <div id='speciesFilterContainer' style='display:none; margin-bottom:5px;'>
        <label><b>Filter by Species:</b></label><br/>
        <select id='speciesFilter' multiple size='5' style='width:100%; line-height:1.5em; font-size:13px;'></select>
      </div>
      <div id='collFilterContainer' style='display:none; margin-bottom:5px;'>
        <label><b>Filter by Collection Status:</b></label><br/>
        <select id='collFilter' multiple size='3'>
          <option value='collected'>Collected (Gray)</option>
          <option value='missing'>Not Collected (Red)</option>
          <option value='palm'>Palm (Yellow)</option>
        </select>
      </div>
    `;
    L.DomEvent.disableClickPropagation(div);
    return div;
  };

  filterContainer.addTo(map);

  // Collect all CircleMarkers on the map
  var allMarkers = [];
  map.eachLayer(function(layer) {
    if (layer instanceof L.CircleMarker) {
      allMarkers.push(layer);
    }
  });

  // Initialize structures for filters
  var speciesToFamily = {};
  var familyToSpecies = {};
  var speciesSet = new Set();
  var familySet = new Set();
  var markersWithPhotos = new Set();

  // Extract species, family, voucher, and photo status for each marker
  allMarkers.forEach(marker => {
    var html = marker.getPopup()?.getContent?.() || '';
    var speciesMatch = html.match(/<b>Species:<\\/b>\\s*<i>\\s*(.*?)<\\/i><br\\/>/);
    var familyMatch = html.match(/<b>Family:<\\/b>\\s*(.*?)<br\\/>/);
    var voucherMatch = html.match(/<b>Voucher:<\\/b>\\s*(.*?)</);
    var hasPhoto = html.includes('slideshow-container');

    var species = speciesMatch?.[1]?.trim() || '';
    var family = familyMatch?.[1]?.trim() || '';
    var voucher = voucherMatch?.[1]?.trim() || '';

    if (species) speciesSet.add(species);
    if (family) familySet.add(family);
    if (hasPhoto && voucher) markersWithPhotos.add(voucher);

    if (species && family) {
      speciesToFamily[species] = family;
      if (!familyToSpecies[family]) familyToSpecies[family] = new Set();
      familyToSpecies[family].add(species);
    }

    // Store marker data as custom properties
    marker._voucher = voucher;
    marker._species = species;
    marker._family = family;
    marker._status = getMarkerStatus(marker);
    marker._hasPhoto = hasPhoto;
  });

  // DOM references
  var familyFilter = document.getElementById('familyFilter');
  var speciesFilter = document.getElementById('speciesFilter');
  var collFilter = document.getElementById('collFilter');
  var photoCheckbox = document.getElementById('togglePhotos');

  // Populate dropdown with options
  function populateSelect(selectElement, items, italicize = false) {
    selectElement.innerHTML = '';
    Array.from(items).sort().forEach(item => {
      var option = document.createElement('option');
      option.value = item;
      option.selected = true;
      option.innerHTML = italicize ? `<i>${item}</i>` : item;
      selectElement.appendChild(option);
    });
  }

  // Get selected values from a <select>
  function getSelectedOptions(select) {
    return Array.from(select.selectedOptions).map(opt => opt.value);
  }

  // Define collection status from marker color
  function getMarkerStatus(marker) {
    var fill = marker.options.fillColor.toLowerCase();
    if (['gray', '#808080'].includes(fill)) return 'collected';
    if (['red', '#ff0000'].includes(fill)) return 'missing';
    if (['yellow', 'gold', '#ffd700'].includes(fill)) return 'palm';
    return '';
  }

  // Filter species options based on selected families
  function updateSpeciesOptions(selectedFamilies) {
    let filteredSpecies = new Set();
    selectedFamilies.forEach(fam => {
      if (familyToSpecies[fam]) {
        familyToSpecies[fam].forEach(sp => filteredSpecies.add(sp));
      }
    });
    populateSelect(speciesFilter, filteredSpecies, true);
  }

  // Populate dropdowns on initial load
  populateSelect(familyFilter, familySet);
  populateSelect(speciesFilter, speciesSet, true);
  Array.from(collFilter.options).forEach(opt => opt.selected = true);

  // Update visibility of markers based on filters
  function updateMarkers() {
    var useFamily = document.getElementById('toggleFamily').checked;
    var useSpecies = document.getElementById('toggleSpecies').checked;
    var useColl = document.getElementById('toggleColl').checked;
    var usePhotos = document.getElementById('togglePhotos').checked;

    var selectedFamilies = useFamily ? getSelectedOptions(familyFilter) : Array.from(familySet);
    var selectedSpecies = useSpecies ? getSelectedOptions(speciesFilter) : Array.from(speciesSet);
    var selectedStatus = useColl ? getSelectedOptions(collFilter) : ['collected', 'missing', 'palm'];

    allMarkers.forEach(marker => {
      var show = selectedFamilies.includes(marker._family) &&
                 selectedSpecies.includes(marker._species) &&
                 selectedStatus.includes(marker._status) &&
                 (!usePhotos || marker._hasPhoto);

      marker.setStyle({
        opacity: show ? 1 : 0,
        fillOpacity: show ? 0.8 : 0
      });
    });
  }

  // Toggle filter visibility and reset selections when disabled
  ['Family','Species','Coll','Photos'].forEach(key => {
    var checkbox = document.getElementById('toggle' + key);
    var container = document.getElementById(key.toLowerCase() + 'FilterContainer');

    if (container) {
      checkbox.addEventListener('change', () => {
        container.style.display = checkbox.checked ? 'block' : 'none';

        if (!checkbox.checked) {
          if (key === 'Family') {
            populateSelect(familyFilter, familySet);
          } else if (key === 'Species') {
            populateSelect(speciesFilter, speciesSet, true);
          } else if (key === 'Coll') {
            Array.from(collFilter.options).forEach(opt => opt.selected = true);
          }
        }

        // If family filter is active and species filter is also active, update species list
        if (key === 'Family' && checkbox.checked && document.getElementById('toggleSpecies').checked) {
          updateSpeciesOptions(getSelectedOptions(familyFilter));
        }

        updateMarkers();
      });
    } else if (key === 'Photos') {
      checkbox.addEventListener('change', updateMarkers);
    }
  });

  // React to dropdown changes
  familyFilter.addEventListener('change', () => {
    if (document.getElementById('toggleSpecies').checked) {
      updateSpeciesOptions(getSelectedOptions(familyFilter));
    }
    updateMarkers();
  });

  speciesFilter.addEventListener('change', updateMarkers);
  collFilter.addEventListener('change', updateMarkers);
}
")

#-------------------------------------------------------------------------------
  # Carousel-like navigation code in JavaScript + CSS
  carousel_js_css <- htmltools::HTML("
<style>
.mySlides { display: none; }
.mySlides img { border-radius: 5px; }
.prev, .next {
  cursor: pointer;
  font-size: 18px;
  user-select: none;
  color: #333;
}
</style>

<script>
function showSlides(n, container) {
  var slides = container.getElementsByClassName('mySlides');
  if (slides.length === 0) return;
  if (n > slides.length) n = 1;
  if (n < 1) n = slides.length;

  for (let i = 0; i < slides.length; i++) {
    slides[i].style.display = 'none';
  }
  slides[n - 1].style.display = 'block';
  container.setAttribute('data-slide', n);
}

function plusSlides(n) {
  var popup = document.querySelector('.leaflet-popup-content');
  if (!popup) return;
  var container = popup.querySelector('.slideshow-container');
  if (!container) return;

  var current = parseInt(container.getAttribute('data-slide')) || 1;
  showSlides(current + n, container);
}

document.addEventListener('click', function(e) {
  if (e.target.closest('.leaflet-marker-icon') || e.target.closest('.leaflet-interactive')) {
    setTimeout(function() {
      var container = document.querySelector('.slideshow-container');
      if (container) showSlides(1, container);
    }, 100);
  }
});
</script>
")

  # End JavaScript and CSS code
  #_____________________________________________________________________________

  map <- leaflet(fp_coords) %>%
    addProviderTiles(base_tiles, options = providerTileOptions(maxZoom = 31)) %>%
    addPolygons(
      lng = c(p1[1], p2[1], p4[1], p3[1], p1[1]),
      lat = c(p1[2], p2[2], p4[2], p3[2], p1[2]),
      color = "darkgreen", fillColor = "lightgreen", fillOpacity = 0.15,
      popup = plot_popup
    ) %>%
    addCircleMarkers(
      lng = ~Longitude,
      lat = ~Latitude,
      radius = rescale(fp_coords$D, to = c(2, 8), from = range(fp_coords$D, na.rm = TRUE)),
      color = "black",
      weight = 0.5,
      fillColor = ~color,
      fillOpacity = 0.8,
      popup = ~popup
    ) %>%
    setView(lng = mean(vertex_coords$longitude),
            lat = mean(vertex_coords$latitude),
            zoom = 19) %>%
    addControl(
      html = paste0(
        "<div style='text-align:center;'>",
        "<div style='font-size:22px; font-weight:bold;'>", plot_name, "</div>",
        "</div>",
        "<div style='font-size:18px;'>", plot_code, "</div>",
        "</div>"
      ),
      position = "topleft"
    ) %>%
    htmlwidgets::onRender(filter_script) %>%
    htmlwidgets::prependContent(carousel_js_css)

  filename <- paste0("Results_", format(Sys.time(), "%d%b%Y_"), filename)

  message("Saving HTML map in: ", file.path(paste0(filename, ".html")))
  htmlwidgets::saveWidget(map, file = file.path(paste0(filename, ".html")), selfcontained = TRUE)
  unlink(file.path(paste0(filename, "_files")), recursive = TRUE)

  # End function
}


#_______________________________________________________________________________
# Auxiliary function for coordinate interpolation ####

.get_latlon <- function(x, y) {
  left <- geosphere::destPoint(p1, geosphere::bearing(p1, p2), y)
  right <- geosphere::destPoint(p3, geosphere::bearing(p3, p4), y)
  interp <- geosphere::destPoint(left, geosphere::bearing(left, right), x)
  return(interp[1, c("lat", "lon")])
}
