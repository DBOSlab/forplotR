#' Plot and save vouchered forest plot map as HTML
#'
#' @author Giulia Ottino & Domingos Cardoso
#'
#' @description
#' Generates an interactive and self-contained HTML map following the
#' \href{https://forestplots.net/}{ForestPlots plot protocol}, using Leaflet
#' and based on vouchered tree data and the geographical coordinates of the
#' plot corners. The function:(i) Converts local subplot coordinates (X, Y) to
#' geographic coordinates using geospatial interpolation from four corner vertices;
#' (ii) automatically extracts metadata like team name, plot name, and plot code
#' from the input xlsx file; (iii) draws the plot boundary polygon with angular
#' ordering to prevent self-intersection, over a selectable basemap ("satellite"
#' or "street"); (iv) colors tree points based on collection status and palms:
#' palms (yellow), collected (gray), and missing (red); (v) embeds image
#' carousels in the popups for vouchers when corresponding images are available
#' in the specified folder; (vi) adds an interactive filter sidebar with
#' checkboxes, multi-select inputs, and search box for filtering by family,
#' species, collection status, and presence of photos; (vii) displays informative
#' popups for each tree,showing tag number, subplot,family, species, DBH,
#' voucher code, and photos (if present); (viii) saves the resulting map
#' as a date-stamped standalone HTML file to the specified directory for easy
#' sharing, archiving, or field consultation.
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
#' plot_html_map("data/tree_data.xlsx",
#'                vertex_coords = vertex_df,
#'                voucher_imgs = "voucher_imgs")
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
    "<b>Species:</b> <i class='taxon'>",
    fp_coords[["Original determination"]],   "</i><br/>",
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

  # replace the old choice logic
  base_tiles <- if (map_type == "satellite") {
    "Esri.WorldImagery"
  } else {
    "OpenTopoMap"   # verde, sem ícones de árvore, gratuito
  }

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
  sidebar_css_js <- htmltools::HTML("
<!-- ---------- SIDEBAR + SEARCH ------------ -->
<style>
#sidebar {
  position: absolute; top: 0; right: 0;
  width: 280px; height: 100%;
  background: rgba(255,255,255,0.97);
  box-shadow: -4px 0 10px rgba(0,0,0,0.3);
  overflow-y: auto; padding: 12px 14px;
  transform: translateX(100%);
  transition: transform .28s ease-in-out;
  z-index: 1001;
  font-size: 14px;
}
#sidebar.open { transform: translateX(0); }

#sidebarToggle {
  position: absolute; top: 10px; right: 15px;
  width: 34px; height: 34px;
  border-radius: 4px; background: #fff; color: #333;
  box-shadow: 0 0 5px rgba(0,0,0,.35);
  cursor: pointer; z-index: 1100;
  display:flex; align-items:center; justify-content:center;
  font-size: 20px; user-select:none;
}

#searchInput {
  width: 100%; padding: 6px 8px; margin-bottom: 8px;
  border: 1px solid #999; border-radius: 4px;
}

.leaflet-popup-content { max-width: 260px; padding: 4px; }
.leaflet-popup-content img { width: 100%; height: auto; border-radius: 4px; }

.mySlides { display:none; }
.prev,.next { cursor:pointer; font-size:18px; color:#444; }
</style>

<script>
function addFilterControl(el, x) {
  // Create slide-out sidebar elements
  var map = this,
      mapDiv = map.getContainer(),
      sb   = document.createElement('div'),
      btn  = document.createElement('div');

  sb.id = 'sidebar';
  btn.id = 'sidebarToggle';
  btn.innerHTML = '&#9776;';  // hamburger icon

  mapDiv.appendChild(sb);
  mapDiv.appendChild(btn);

  // Toggle sidebar
  btn.addEventListener('click',()=> sb.classList.toggle('open'));

  // Sidebar HTML content
  sb.innerHTML = `
    <input id='searchInput' type='text' placeholder='Search family or species'/>
    <div style='margin-bottom:6px;'><b>Enable Filters:</b><br/>
      <label><input type='checkbox' id='toggleFamily' checked> Family</label><br/>
      <label><input type='checkbox' id='toggleSpecies' checked> Species</label><br/>
      <label><input type='checkbox' id='toggleColl' checked> Collection&nbsp;Status</label><br/>
      <label><input type='checkbox' id='togglePhotos'> With&nbsp;Photos&nbsp;Only</label>
    </div>
    <div id='familyFilterContainer'>
      <b>Filter by Family:</b><br/>
      <select id='familyFilter' multiple size='6' style='width:100%;'></select>
    </div><br/>
    <div id='speciesFilterContainer'>
      <b>Filter by Species:</b><br/>
      <select id='speciesFilter' multiple size='6' style='width:100%;'></select>
    </div><br/>
    <div id='collFilterContainer'>
      <b>Filter by Collection Status:</b><br/>
      <select id='collFilter' multiple size='3' style='width:100%;'>
        <option value='collected'>Collected (Gray)</option>
        <option value='missing'>Not Collected (Red)</option>
        <option value='palm'>Palm (Yellow)</option>
      </select>
    </div>
  `;

  // Gather marker info
  var markers = [];
  map.eachLayer(l=>{
     if(l instanceof L.CircleMarker){ markers.push(l); }
  });

  var famSet = new Set(), spSet  = new Set(),
      sp2fam = {}, fam2sp = {}, photoVouchers = new Set();

  markers.forEach(m=>{
    var html = m.getPopup().getContent(),
        fam  = /<b>Family:<\\/b>\\s*(.*?)<br\\/>/.exec(html)?.[1]?.trim() || '',
        sp   = /<b>Species:<\\/b>\\s*<i[^>]*>\\s*(.*?)<\\/i>/.exec(html)?.[1]?.trim() || '',
        vou  = /<b>Voucher:<\\/b>\\s*(.*?)</.exec(html)?.[1]?.trim() || '',
        photo= html.includes('slideshow-container');

    if(fam) famSet.add(fam);
    if(sp)  spSet.add(sp);
    if(photo && vou) photoVouchers.add(vou);

    sp2fam[sp]=fam;
    fam2sp[fam] = fam2sp[fam] || new Set();
    fam2sp[fam].add(sp);

    Object.assign(m,{ _fam:fam, _sp:sp, _vou:vou, _hasPhoto:photo, _stat:getStatus(m) });
  });

  const famSel=document.getElementById('familyFilter'),
        spSel =document.getElementById('speciesFilter'),
        csSel =document.getElementById('collFilter');

  function fill(sel, set, italic=false){
     sel.innerHTML='';
     Array.from(set).sort().forEach(v=>{
        var o=document.createElement('option');
        o.value=v; o.selected=true;
        o.innerHTML=italic?`<i>${v}</i>`:v;
        sel.appendChild(o);
     });
  }

  fill(famSel,famSet); fill(spSel,spSet,true);

  const searchBox=document.getElementById('searchInput');
  searchBox.addEventListener('keyup',updateMarkers);

  famSel.addEventListener('change',()=>{
     if(document.getElementById('toggleSpecies').checked){
        let famChosen=getSel(famSel);
        let tmp=new Set();
        famChosen.forEach(f=>fam2sp[f]?.forEach(s=>tmp.add(s)));
        fill(spSel,tmp,true);
     }
     updateMarkers();
  });
  spSel.addEventListener('change',updateMarkers);
  csSel.addEventListener('change',updateMarkers);

  ['Family','Species','Coll','Photos'].forEach(k=>{
     var cb=document.getElementById('toggle'+k),
         block=document.getElementById(k.toLowerCase()+'FilterContainer');
     if(block) block.style.display='block';
     cb.addEventListener('change',()=>{
        if(block) block.style.display=cb.checked?'block':'none';
        if(!cb.checked){
           if(k==='Family') fill(famSel,famSet);
           if(k==='Species') fill(spSel,spSet,true);
           if(k==='Coll') Array.from(csSel.options).forEach(o=>o.selected=true);
        }
        updateMarkers();
     });
  });

  function getSel(sel){ return Array.from(sel.selectedOptions).map(o=>o.value); }
  function getStatus(m){
      var c=m.options.fillColor.toLowerCase();
      if(['gray','#808080'].includes(c)) return 'collected';
      if(['red','#ff0000'].includes(c)) return 'missing';
      if(['gold','yellow','#ffd700'].includes(c)) return 'palm';
      return '';
  }

  // Filter both markers and dropdowns based on all active filters + search term
  function updateMarkers(){
     var fOn=document.getElementById('toggleFamily').checked,
         sOn=document.getElementById('toggleSpecies').checked,
         cOn=document.getElementById('toggleColl').checked,
         pOn=document.getElementById('togglePhotos').checked,
         term=searchBox.value.trim().toLowerCase();

     var selFam=fOn?getSel(famSel):Array.from(famSet),
         selSp =sOn?getSel(spSel):Array.from(spSet),
         selSt =cOn?getSel(csSel):['collected','missing','palm'];

     var filteredMarkers = [];

     markers.forEach(m=>{
        var matchText = term === '' || m._fam.toLowerCase().includes(term) ||
                                       m._sp.toLowerCase().includes(term);
        var show = selFam.includes(m._fam) &&
                   selSp.includes(m._sp) &&
                   selSt.includes(m._stat) &&
                   (!pOn || m._hasPhoto) &&
                   matchText;

        if(show){
           if(!map.hasLayer(m)) m.addTo(map);
           filteredMarkers.push(m);
        } else {
           if(map.hasLayer(m)) map.removeLayer(m);
        }
     });

     // Update dropdowns to reflect only options in visible markers
     let visibleFams = new Set(), visibleSps = new Set();
     filteredMarkers.forEach(m => {
        visibleFams.add(m._fam);
        visibleSps.add(m._sp);
     });

     if(fOn) fill(famSel, visibleFams);
     if(sOn) fill(spSel, visibleSps, true);
  }
}
</script>
")

  # End JavaScript and CSS code
  #_____________________________________________________________________________

  map <- leaflet(fp_coords) %>%
    addProviderTiles(base_tiles, options = providerTileOptions(maxZoom = 31,
                                                               maxNativeZoom = 17  )) %>%
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
            zoom = 18) %>%
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
    htmlwidgets::onRender(
      "function(el, x){ this.whenReady(function(){ addFilterControl.call(this, el, x); }); }"
    ) %>%

    htmlwidgets::prependContent(
      carousel_js_css,
      sidebar_css_js
    )

  map <- map %>%
    # 1) Load the Roboto family (light, regular, medium, bold)
    htmlwidgets::prependContent(
      htmltools::tags$link(
        href = "https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,300;0,400;0,500;0,700;1,300;1,400;1,500;1,700&display=swap",
        rel  = "stylesheet"
      ),
      # 2) Apply Roboto everywhere and set some helper classes
      htmltools::tags$style(htmltools::HTML("
      /* Global font replacement */
      .leaflet-container,
      .leaflet-popup-content,
      #sidebar {
        font-family: 'Roboto', sans-serif;
      }

      /* Italic for scientific names */
      .taxon { font-style: italic; }

      /* Map title and subtitle helpers */
      .mapTitle    { font-size: 22px; font-weight: 700; }
      .mapSubtitle { font-size: 16px; font-weight: 400; }
    "))
    )
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
