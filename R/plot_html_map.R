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
#' voucher code, and photos; (viii) saves the resulting map as a date-stamped
#' standalone HTML file with multiple interactive leaflet map options to the
#' specified directory for easy sharing, archiving, or field consultation.
#'
#' @usage plot_html_map(fp_file_path = NULL,
#'                      vertex_coords = NULL,
#'                      plot_size = 1,
#'                      subplot_size = 10,
#'                      voucher_imgs = NULL,
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
#' @param filename Character. Name of the output HTML file (default is
#' "plot_map.html")..
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
                          voucher_imgs = NULL,
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
    dplyr::mutate(
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
    dplyr::mutate(
      col = floor((T1 - 1) / n_rows),
      row = (T1 - 1) %% n_rows,
      global_x = col * subplot_size + X,
      global_y = dplyr::if_else(
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
    dplyr::rename_with(~tolower(gsub("\\s*\\(.*?\\)", "", .x)), everything()) %>%
    dplyr::rename(latitude = latitude, longitude = longitude) %>%
    dplyr::mutate(
      longitude = as.numeric(gsub(",", ".", longitude)),
      latitude  = as.numeric(gsub(",", ".", latitude))
    )

  # Correct signals to ensure coordinates are in southern hemisphere
  vertex_coords <- vertex_coords %>%
    dplyr::mutate(
      longitude = ifelse(longitude > 0, -longitude, longitude),
      latitude  = ifelse(latitude > 0, -latitude, latitude)
    )

  # Use coordinates after cleaning
  p1 <- c(vertex_coords$longitude[1], vertex_coords$latitude[1])
  p2 <- c(vertex_coords$longitude[2], vertex_coords$latitude[2])
  p3 <- c(vertex_coords$longitude[3], vertex_coords$latitude[3])
  p4 <- c(vertex_coords$longitude[4], vertex_coords$latitude[4])

  coords_geo <- mapply(
    .get_latlon,
    x = fp_coords$global_x,
    y = fp_coords$global_y,
    MoreArgs = list(p1 = p1, p2 = p2, p3 = p3, p4 = p4)
  )
  fp_coords$Latitude <- coords_geo[1, ]
  fp_coords$Longitude <- coords_geo[2, ]

  # Assign color by collection status
  fp_coords$color <- dplyr::case_when(
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
                       "<b>Plot Code:</b> ", plot_code, "<br/>",
                       "<b>Team:</b> ", team)


  # Define the base folder where the original images are stored
  original_image_base <- here::here(voucher_imgs)

  # List all subdirectories once to improve performance
  leaf_dirs <- list.dirs(original_image_base, recursive = TRUE, full.names = TRUE)

  # Keep only those that do not contain other directories (i.e., leaf directories)
  leaf_dirs <- leaf_dirs[!sapply(leaf_dirs, function(dir) {
    any(file.info(list.dirs(dir, recursive = FALSE))$isdir)
  })]

  leaf_dirs <- sub(paste0(".*(", voucher_imgs, "/.*)"), "\\1", leaf_dirs)

  # -------------------------------------------------
  #  build pop-ups with photo carousels
  handled <- character(0)

  for (dir_path in leaf_dirs) {
    voucher_id <- basename(dir_path)

    # skip if this voucher was already handled
    if (voucher_id %in% handled) next
    handled <- c(handled, voucher_id)

    # list image files in this directory
    img_files <- list.files(
      dir_path, full.names = TRUE,
      pattern = "\\.(jpg|jpeg|png|gif)$", ignore.case = TRUE
    )
    if (!length(img_files)) next      # no photos → nothing to add

    rel_paths <- file.path(dir_path, basename(img_files))

    # build HTML slides for the carousel
    slides <- paste0(
      "<div class='mySlides'>",
      "<a href='", rel_paths, "' target='_blank'>",
      "<img src='", rel_paths,
      "' style='max-width:200px; max-height:200px; display:block; margin:auto;'></a></div>",
      collapse = "\n"
    )

    # Append the carousel to the correct tree
    tf <- fp_coords$Voucher == voucher_id
    fp_coords$popup[tf] <- paste0(
      fp_coords$popup[tf],
      "<br>
      <span class='hasPhotoBadge'>Has Photo</span>",
      "<div class='slideshow-container' data-slide='1'>", slides, "</div>",
      "<div style='text-align:center;'>",
      "<a class='prev' onclick='plusSlides(-1)' style='margin-right:10px;'>&#10094;</a>",
      "<a class='next' onclick='plusSlides(1)'>&#10095;</a>",
      "</div>"
    )
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
<style>
/* ---------- SIDEBAR STYLES ---------- */
#sidebar {
  position:absolute; top:0; right:0;
  width:280px; height:100%;
  background:rgba(255,255,255,0.97);
  box-shadow:-4px 0 10px rgba(0,0,0,.3);
  overflow-y:auto; padding:12px 14px;
  transform:translateX(100%);
  transition:transform .28s ease-in-out;
  z-index:1001; font-size:14px;
}
#sidebar.open { transform:translateX(0); }

#sidebarToggle {
  position:absolute; top:10px; right:15px;
  width:34px; height:34px;
  border-radius:4px; background:#fff; color:#333;
  box-shadow:0 0 5px rgba(0,0,0,.35);
  cursor:pointer; z-index:1100;
  display:flex; align-items:center; justify-content:center;
  font-size:20px; user-select:none;
}

#searchInput {
  width:100%; padding:6px 8px; margin-bottom:8px;
  border:1px solid #999; border-radius:4px;
}

.leaflet-popup-content { max-width:260px; padding:4px; }
.leaflet-popup-content img { width:100%; height:auto; border-radius:4px; }

.mySlides{display:none;}
.prev,.next{cursor:pointer;font-size:18px;color:#444;}
</style>

<script>
function addFilterControl(el,x){
  /* ---------- BUILD SIDEBAR ---------- */
  const map=this, mapDiv=map.getContainer(),
        sb=document.createElement('div'),
        btn=document.createElement('div');
  sb.id='sidebar'; btn.id='sidebarToggle'; btn.innerHTML='&#9776;';
  mapDiv.appendChild(sb); mapDiv.appendChild(btn);
  btn.addEventListener('click',()=>sb.classList.toggle('open'));

  sb.innerHTML=`
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

  /* ---------- COLLECT MARKER DATA ---------- */
  const markers=[];
  map.eachLayer(l=>{ if(l instanceof L.CircleMarker) markers.push(l); });

  const famSet=new Set(), spSet=new Set(),
        fam2sp={}, sp2fam={};

  markers.forEach(m=>{
    const html=m.getPopup().getContent();
    const fam =(html.match(/<b>Family:<\\/b>\\s*(.*?)<br\\/>/)||[])[1]||'';
    const sp  =(html.match(/<b>Species:<\\/b>\\s*<i[^>]*>\\s*(.*?)<\\/i>/)||[])[1]||'';

    if(fam) famSet.add(fam);
    if(sp)  spSet.add(sp);

    sp2fam[sp]=fam;
    (fam2sp[fam]=fam2sp[fam]||new Set()).add(sp);

    Object.assign(m,{ _fam:fam, _sp:sp, _stat:getStatus(m),
                      _hasPhoto:html.includes('slideshow-container') });
  });

  /* ---------- UTILS ---------- */
  const famSel=document.getElementById('familyFilter'),
        spSel =document.getElementById('speciesFilter'),
        csSel =document.getElementById('collFilter'),
        search=document.getElementById('searchInput');

  const fill=(sel,set,italic=false)=>{
    sel.innerHTML='';
    Array.from(set).sort().forEach(v=>{
      const o=document.createElement('option');
      o.value=v; o.selected=true; o.innerHTML=italic?`<i>${v}</i>`:v;
      sel.appendChild(o);
    });
  };

  fill(famSel,famSet);               // keep full family list always
  fill(spSel ,spSet ,true);          // initial full species list
  Array.from(csSel.options).forEach(o=>o.selected=true);

  const getSel=sel=>Array.from(sel.selectedOptions).map(o=>o.value);
  function getStatus(m){
    const c=m.options.fillColor.toLowerCase();
    if(['gray','#808080'].includes(c)) return 'collected';
    if(['red','#ff0000'].includes(c)) return 'missing';
    if(['gold','yellow','#ffd700'].includes(c)) return 'palm';
    return '';
  }

  /* ---------- EVENT LISTENERS ---------- */
  famSel.addEventListener('change',()=>{
    if(document.getElementById('toggleSpecies').checked){
      const chosen=getSel(famSel);
      const subset=new Set();
      chosen.forEach(f=>fam2sp[f]?.forEach(sp=>subset.add(sp)));
      fill(spSel,subset,true);       // rebuild species list only
    }
    updateMarkers();
  });

  spSel.addEventListener('change',updateMarkers);
  csSel.addEventListener('change',updateMarkers);
  search.addEventListener('keyup',updateMarkers);

  ['Family','Species','Coll','Photos'].forEach(k=>{
    const cb=document.getElementById('toggle'+k),
          block=document.getElementById(k.toLowerCase()+'FilterContainer');
    cb.addEventListener('change',()=>{
      if(block) block.style.display=cb.checked?'block':'none';
      if(!cb.checked){
        if(k==='Species') fill(spSel,spSet,true);
        if(k==='Coll')    Array.from(csSel.options).forEach(o=>o.selected=true);
      }
      updateMarkers();
    });
  });

  /* ---------- FILTER FUNCTION ---------- */
  function updateMarkers(){
    const fOn=document.getElementById('toggleFamily').checked,
          sOn=document.getElementById('toggleSpecies').checked,
          cOn=document.getElementById('toggleColl').checked,
          pOn=document.getElementById('togglePhotos').checked,
          term=search.value.trim().toLowerCase();

    const selFam=fOn?getSel(famSel):Array.from(famSet),
          selSp =sOn?getSel(spSel):Array.from(spSet),
          selSt =cOn?getSel(csSel):['collected','missing','palm'];

    markers.forEach(m=>{
      const txt=m._fam.toLowerCase()+m._sp.toLowerCase();
      const show= selFam.includes(m._fam) &&
                  selSp .includes(m._sp ) &&
                  selSt .includes(m._stat) &&
                  (!pOn || m._hasPhoto) &&
                  txt.includes(term);

      if(show){ if(!map.hasLayer(m)) m.addTo(map); }
      else     { if(map.hasLayer(m)) map.removeLayer(m); }
    });

    /* rebuild species list (but NOT family list) */
    if(sOn){
      const visibleSp=new Set();
      markers.filter(m=>map.hasLayer(m)).forEach(m=>visibleSp.add(m._sp));
      fill(spSel,visibleSp,true);
    }
  }
}
</script>
")


  # End JavaScript and CSS code
  #_____________________________________________________________________________
  # Define initial centre/zoom
  initial_zoom <- 18   # keep your current value
  lat0 <- mean(vertex_coords$latitude)
  lon0 <- mean(vertex_coords$longitude)

  # Build the Leaflet map 5 base layers + reset button
  map <- leaflet(fp_coords, options = leaflet::leafletOptions(minZoom = 2,  maxZoom = 22)) %>%

    # Set title and subtitle
    leaflet::addControl(
      html = paste0(
        "<div style='text-align:center; line-height:1.2; padding:4px 8px; " ,
        "background:rgba(255,255,255,0.85); border-radius:8px;" ,
        "box-shadow:0 1px 4px rgba(0,0,0,0.25);'>",
        "<div style='font-size:20px; font-weight:700;'>",
        htmltools::htmlEscape(plot_name),
        "</div>",
        "<div style='font-size:14px; font-weight:400;'>Plot Code: ",
        htmltools::htmlEscape(plot_code),
        "</div>",
        "</div>"),
      position = "topleft",
      className = "custom-title"
    ) %>%

    htmlwidgets::onRender(
      "function(el, x){ this.whenReady(function(){ addFilterControl.call(this, el, x); }); }"
    ) %>%

    htmlwidgets::prependContent(
      carousel_js_css,
      sidebar_css_js
    ) %>%

    leaflet::addProviderTiles(leaflet::providers$OpenStreetMap,       group = "OSM Street") |>
    leaflet::addProviderTiles(leaflet::providers$Esri.WorldStreetMap, group = "ESRI Street") |>
    leaflet::addProviderTiles(leaflet::providers$Esri.WorldImagery,   group = "Satellite")  |>
    leaflet::addProviderTiles(leaflet::providers$OpenTopoMap,         group = "OpenTopo")   |>
    leaflet::addProviderTiles(leaflet::providers$CartoDB.DarkMatter,  group = "Dark Matter") |>

    # Layer control
    leaflet::addLayersControl(
      baseGroups    = c("OSM Street", "ESRI Street", "Satellite", "OpenTopo", "Dark Matter"),
      overlayGroups = c("Specimens"),
      position      = "bottomleft",
      options       = leaflet::layersControlOptions(
        collapsed = TRUE,
        autoZIndex = TRUE
      )
    ) %>%

    # Plot polygon
    leaflet::addPolygons(
      lng = c(p1[1], p2[1], p4[1], p3[1], p1[1]),
      lat = c(p1[2], p2[2], p4[2], p3[2], p1[2]),
      color = "darkgreen", fillColor = "lightgreen", fillOpacity = 0.15,
      popup = plot_popup
    )  %>%

    # Specimmen points
    leaflet::addCircleMarkers(
      lng = ~Longitude,
      lat = ~Latitude,
      radius = scales::rescale(fp_coords$D, to = c(2, 8),
                               from = range(fp_coords$D, na.rm = TRUE)),
      color = "black", weight = 0.5,
      fillColor = ~color, fillOpacity = 0.8,
      popup = ~popup,
      group = "Specimens"
    )  %>%

    # Reset-view buttom
    leaflet::addEasyButton(
      leaflet::easyButton(
        icon    = "fa-crosshairs",
        title   = "Back to plot",
        onClick = leaflet::JS(
          sprintf("function(btn, map){ map.setView([%f, %f], %d); }",
                  lat0, lon0, initial_zoom)
        )
      )
    ) %>%

    # Set initial view
    leaflet::setView(lng = lon0, lat = lat0, zoom = initial_zoom)

  map <- map %>%
    # Load the Roboto font family
    htmlwidgets::prependContent(
      htmltools::tags$link(
        href = "https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,300;0,400;0,500;0,700;1,300;1,400;1,500;1,700&display=swap",
        rel  = "stylesheet"
      ),
      # Apply Roboto everywhere and set some helper classes
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
    ) %>%
    htmlwidgets::prependContent(
      htmltools::tags$style("
      .leaflet-popup-content        { font-size:14px; line-height:1.35; padding:6px 8px 8px; }
      .leaflet-popup-content-wrapper{ border-radius:10px!important; box-shadow:0 3px 10px rgba(0,0,0,.25)!important; }
      .leaflet-popup-tip            { display:none; }
      .leaflet-popup-content img    { max-width:260px; width:100%; border-radius:6px; border:1px solid #ddd;
                                      box-shadow:0 1px 4px rgba(0,0,0,.15); margin-top:4px; }
      .hasPhotoBadge                { display:inline-block; background:#388e3c; color:#fff; font-size:12px;
                                      padding:2px 6px; border-radius:12px; margin-bottom:4px; font-weight:500; }
      .prev, .next                  { color:#006699; font-size:22px; transition:color .2s; }
      .prev:hover, .next:hover      { color:#004466; }
    ")
    )
  filename <- paste0("Results_", format(Sys.time(), "%d%b%Y_"), filename)
  message("Saving HTML map: '", file.path(paste0(filename, ".html'")))
  output_path <- paste0(filename, ".html")
  htmlwidgets::saveWidget(map, file = output_path, selfcontained = TRUE)

  # End function
}


#_______________________________________________________________________________
# Auxiliary function for coordinate interpolation ####

.get_latlon <- function(x, y, p1, p2, p3, p4) {
  left   <- geosphere::destPoint(p1, geosphere::bearing(p1, p2), y)
  right  <- geosphere::destPoint(p3, geosphere::bearing(p3, p4), y)
  interp <- geosphere::destPoint(left, geosphere::bearing(left, right), x)
  return(interp[1, c("lat", "lon")])
}

