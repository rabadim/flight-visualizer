// Part 1: Globals, Constants, and Initialization

// --- Global Variables (alphabetized & grouped) ---
// --- Configuration Constants ---
let headerKeywords = ['departure', 'arrival', 'aircraft', 'duration'];
const dateFormats = [
  { regex: /^\d{1,2}\/\d{1,2}\/\d{2,4}$/, format: 'MM/dd/yyyy' },
  { regex: /^\d{4}-\d{2}-\d{2}$/, format: 'yyyy-MM-dd' },
  { regex: /^\d{2}-\d{2}-\d{4}$/, format: 'MM-dd-yyyy' },
  { regex: /^\d{2} \w{3} \d{4}$/, format: 'dd MMM yyyy' },
  { regex: /^\w{3} \d{1,2}, \d{4}$/, format: 'MMM dd, yyyy' }
];

// --- Debug Logger ---
const DEBUG_MODE = true;
function debugLog(...args) {
  if (DEBUG_MODE) console.log(...args);
}
let airportCache = new Map();
let airportFrequency = {};
let airportsData = { airports: [], code_to_airport_id: {} };
let currentLogbook = 'Logbook 1';
let currentMapStyle = 'light';
let distanceCache = new Map();
let endDate = null;
let isDarkMode = false;
let logbooks = { 'Logbook 1': { flights: [], stats: {} } };
let map;
let missingAirports = new Set();
let pendingCSVText = null;
let pendingLogbookData = null;
let routeFrequency = new Map();
let showLocalFlights = true;
let startDate = null;
let tileLayer;
let totalActualIFR = 0;
let totalCrossCountry = 0;
let totalFlightTime = 0;
let totalNight = 0;
let totalSolo = 0;
let userDefinedFromIndex = -1;
let userDefinedToIndex = -1;
let validAirportCodes = new Set();

const BATCH_SIZE = 20;
const INVALIDATE_TIMEOUT = 200;
const TOAST_TIMEOUT = 3000;
const ERROR_TIMEOUT = 7000;
const TILE_ERROR_TIMEOUT = 5000;

// Predefined fields for LogTen-style CSV import
const logbookFields = [
    { name: 'Date', key: 'date', mandatory: true, validator: (value, dateFormat) => {
        if (!value) return false;
        const formats = {
            'MM/DD/YYYY': /^\d{2}\/\d{2}\/\d{4}$/,
            'DD/MMM/YY': /^\d{2}\/[A-Za-z]{3}\/\d{2}$/,
            'YYYY-MM-DD': /^\d{4}-\d{2}-\d{2}$/,
            'DD-MM-YYYY': /^\d{2}-\d{2}-\d{4}$/
        };
        return formats[dateFormat]?.test(value) && !isNaN(new Date(value).getTime());
    }},
    { name: 'From', key: 'from', mandatory: true, validator: code => validAirportCodes.has(code.toUpperCase()) },
    { name: 'To', key: 'to', mandatory: true, validator: code => validAirportCodes.has(code.toUpperCase()) },
    { name: 'Duration', key: 'duration', mandatory: false, validator: val => !isNaN(parseFloat(val)) && parseFloat(val) >= 0 },
    { name: 'Aircraft Type', key: 'aircraftType', mandatory: false },
    { name: 'Registration', key: 'registration', mandatory: false },
    { name: 'Notes', key: 'notes', mandatory: false }
];

let fieldMappings = {};
let pendingRows = null;
let importOptions = { dateFormat: 'MM/DD/YYYY', ignoreFirstRow: true, timesInMinutes: false };
let currentPreviewRow = 0;

const mapStyles = {
    light: {
        light: L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
            noWrap: true,
            minZoom: 2,
            maxZoom: 18
        }),
        dark: L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}.png', {
            attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors © <a href="https://carto.com/attributions">CARTO</a>',
            noWrap: true,
            minZoom: 2,
            maxZoom: 18
        })
    },
    dark: L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}.png', {
        attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors © <a href="https://carto.com/attributions">CARTO</a>',
        noWrap: true,
        minZoom: 2,
        maxZoom: 18
    }),
    satellite: {
        light: L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
            attribution: 'Tiles © Esri — Source: Esri, i-cubed, USDA, USGS, AEX, GeoEye, Getmapping, Aerogrid, IGN, IGP, UPR-EGP, and the GIS User Community',
            noWrap: true,
            minZoom: 2,
            maxZoom: 18
        }),
        dark: L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
            attribution: 'Tiles © Esri — Source: Esri, i-cubed, USDA, USGS, AEX, GeoEye, Getmapping, Aerogrid, IGN, IGP, UPR-EGP, and the GIS User Community',
            noWrap: true,
            minZoom: 2,
            maxZoom: 18
        })
    }
};

const fallbackTileLayer = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
    noWrap: true,
    minZoom: 2,
    maxZoom: 18
});

// Function to calculate optimal zoom level based on screen size
function calculateDynamicZoom() {
    const mapContainer = document.getElementById('map');
    const width = mapContainer.clientWidth; // Screen width in pixels
    const height = mapContainer.clientHeight; // Screen height in pixels

    // Approximate world width in degrees (360° longitude)
    const worldWidthDeg = 360;
    const tileSize = 256; // Leaflet tile size in pixels

    // Adjust for Mercator projection: effective width at equator
    // At zoom 0, world is 256px wide; doubles each zoom level
    let zoom = Math.log2(width / tileSize) + 0.5; // Base zoom to fit width
    zoom = Math.max(2, Math.min(4, Math.ceil(zoom))); // Clamp between 2 and 6

    debugLog(`[calculateDynamicZoom] Screen: ${width}x${height}, Calculated Zoom: ${zoom}`);
    return zoom;
}

// Debounce function to limit resize event frequency
function debounce(func, wait) {
    let timeout;
    return function (...args) {
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(this, args), wait);
    };
}

function showFileError(message, category = 'general') {
    const fileError = document.getElementById('file-error');
    let detailedMessage = message;

    switch (category) {
        case 'file-format':
            detailedMessage += ' Please check the file type and ensure it is a CSV, PDF, TXT, JPG, PNG, or JPEG.';
            break;
        case 'parsing':
            detailedMessage += ' Ensure your logbook has valid FROM and TO columns or try mapping them manually.';
            break;
        case 'airport-code':
            detailedMessage += ' These codes may be aircraft types, simulator codes, or missing airports. Add them using the form below if they are airports.';
            break;
        case 'date-parsing':
            detailedMessage += ' Please ensure dates are in a supported format like DD.MMM.YY, MM/DD/YYYY, or YYYY-MM-DD.';
            break;
    }

    fileError.querySelector('p').textContent = detailedMessage;
    fileError.classList.remove('hidden');
    setTimeout(() => fileError.classList.add('hidden'), ERROR_TIMEOUT);
}

function showToast(message) {
    const existing = document.querySelector('.fixed.bottom-4.right-4');
    if (existing) existing.remove();
    const toast = document.createElement('div');
    toast.className = 'fixed bottom-4 right-4 bg-green-500 text-white px-4 py-2 rounded shadow-lg z-1300 transition-opacity duration-300';
    toast.textContent = message;
    document.body.appendChild(toast);
    setTimeout(() => {
        toast.style.opacity = '0';
        setTimeout(() => toast.remove(), 300);
    }, TOAST_TIMEOUT);
}

function updateMapStyle(style) {
    const mapStyleLoading = document.getElementById('map-style-loading');
    const tileError = document.getElementById('tile-error');
    mapStyleLoading.classList.remove('hidden');
    tileError.classList.add('hidden');

    if (tileLayer) {
        try { map.removeLayer(tileLayer); } catch (_) {}
    }

    let newTileLayer;
    if (style === 'dark') {
        newTileLayer = mapStyles.dark;
    } else {
        const mode = isDarkMode ? 'dark' : 'light';
        newTileLayer = mapStyles[style][mode];
    }

    newTileLayer.on('tileerror', () => {
        if (tileLayer) {
            try { map.removeLayer(tileLayer); } catch (_) {}
        }
        tileLayer = fallbackTileLayer.addTo(map);
        tileError.classList.remove('hidden');
        setTimeout(() => {
            tileError.classList.add('hidden');
        }, TILE_ERROR_TIMEOUT);
    });

    tileLayer = newTileLayer.addTo(map);

    setTimeout(() => {
        map.invalidateSize();
        map.panBy([1, 1], { animate: false });
        mapStyleLoading.classList.add('hidden');
    }, INVALIDATE_TIMEOUT);
}

// --- Theme Toggle Handler ---
function handleThemeToggle() {
    document.body.classList.toggle('dark-mode');
    isDarkMode = document.body.classList.contains('dark-mode');
    this.querySelector('span:not([data-icon])').textContent =
      isDarkMode ? 'Toggle Light Mode' : 'Toggle Dark Mode';
    updateMapStyle(currentMapStyle);
}

// --- Map and UI Initialization ---
window.addEventListener('load', () => {
    // Cache frequently accessed DOM elements
    const themeToggleBtn = document.getElementById('theme-toggle');
    const mapStyleSelect = document.getElementById('map-style');
    const loadingIndicator = document.getElementById('loading-indicator');
    map = L.map('map', {
        zoomControl: true,
        attributionControl: true,
        fadeAnimation: true,
        worldCopyJump: false,
        maxBounds: [[-90, -180], [90, 180]],
        maxBoundsViscosity: 1.0,
        minZoom: 2,
        maxZoom: 18
    }).setView([20, -60], calculateDynamicZoom());

    tileLayer = mapStyles.light.light.addTo(map);

    setTimeout(() => {
        map.invalidateSize();
    }, INVALIDATE_TIMEOUT);

    // Update zoom on window resize with debounce
    const updateZoomOnResize = debounce(() => {
        const newZoom = calculateDynamicZoom();
        map.setZoom(newZoom, { animate: true });
        map.invalidateSize();
        debugLog(`[resize] Updated Zoom: ${newZoom}`);
    }, 200);

    window.addEventListener('resize', updateZoomOnResize);

    mapStyleSelect.addEventListener('change', e => {
        currentMapStyle = e.target.value;
        updateMapStyle(currentMapStyle);
    });

    themeToggleBtn.addEventListener('click', handleThemeToggle);

    const menuOverlay = document.getElementById('menu-overlay');
    const menuBackdrop = document.getElementById('menu-backdrop');
    const menuToggle = document.getElementById('menu-toggle');
    // loadingIndicator already cached above
    const mapStyleLoading = document.getElementById('map-style-loading');
    const tileError = document.getElementById('tile-error');
    const fileError = document.getElementById('file-error');

    const progressBar = document.createElement('div');
    progressBar.id = 'progress-bar';
    progressBar.style.display = 'none';
    progressBar.style.width = '100%';
    progressBar.style.height = '5px';
    progressBar.style.backgroundColor = '#ddd';
    const progress = document.createElement('div');
    progress.style.height = '100%';
    progress.style.backgroundColor = '#4CAF50';
    progress.style.width = '0%';
    progressBar.appendChild(progress);
    document.body.appendChild(progressBar);

    let isDownloading = false;

    const toggleMenu = () => {
        menuOverlay.classList.toggle('translate-x-full');
        menuBackdrop.classList.toggle('hidden');
        const isOpen = !menuOverlay.classList.contains('translate-x-full');
        document.getElementById('menu-icon-open').classList.toggle('hidden', isOpen);
        document.getElementById('menu-icon-close').classList.toggle('hidden', !isOpen);
        menuToggle.setAttribute('aria-label', isOpen ? 'Close Menu' : 'Open Menu');
    };

    menuToggle.addEventListener('click', toggleMenu);
    menuBackdrop.addEventListener('click', toggleMenu);

    const toggleBtn = (id, containerId, chevron, openText, closeText) => {
        document.getElementById(id).addEventListener('click', function () {
            const container = document.getElementById(containerId);
            container.classList.toggle('hidden');
            this.classList.toggle('expanded');
            this.querySelector(`span[data-chevron="${chevron}"]`).classList.toggle('expanded');
            this.querySelector('.text-label').textContent = container.classList.contains('hidden') ? openText : closeText;
        });
    };

    toggleBtn('toggle-form-btn', 'flight-form-container', 'form', 'Add Flight', 'Hide Form');
    toggleBtn('toggle-logbook-btn', 'logbook-upload-container', 'logbook', 'Upload Logbook', 'Hide Logbook Upload');
    toggleBtn('toggle-date-filter-btn', 'date-filter-container', 'date-filter', 'Filter by Date Range', 'Hide Date Filter');
    toggleBtn('toggle-filter-flights-btn', 'filter-flights-container', 'filter-flights', 'Filter Flights', 'Hide Flight Filters');

    document.getElementById('submit-flight').addEventListener('click', event => {
        event.preventDefault();
        const departure = document.getElementById('departure').value.toUpperCase();
        const arrival = document.getElementById('arrival').value.toUpperCase();
        const notes = document.getElementById('flight-notes')?.value || '';

        const depAirport = getAirport(departure);
        const arrAirport = getAirport(arrival);

        if (depAirport && arrAirport) {
            logbooks[currentLogbook].flights.push({
                coords: [depAirport.coords, arrAirport.coords],
                codes: [depAirport.iata, arrAirport.iata],
                date: 'N/A',
                distance: calculateDistance(depAirport.coords, arrAirport.coords),
                duration: 0,
                notes,
                aircraftType: 'Unknown',
                registration: 'Unknown',
                crossCountry: 0,
                night: 0,
                solo: 0,
                actualIFR: 0,
                isLocal: depAirport.iata === arrAirport.iata
            });
            updateAirportFrequency(depAirport.iata, arrAirport.iata);
            updateRouteFrequency(depAirport.iata, arrAirport.iata);
            drawFlights(true);
            showToast(`Flight from ${depAirport.iata} to ${arrAirport.iata} added successfully!`);
        } else {
            showFileError(`Invalid or missing airport code! Please ensure the codes (${departure}, ${arrival}) exist in airports.json.`, 'airport-code');
        }

        document.getElementById('departure').value = '';
        document.getElementById('arrival').value = '';
        document.getElementById('flight-notes')?.value && (document.getElementById('flight-notes').value = '');
    });

    document.getElementById('clear-flights-btn').addEventListener('click', () => {
        logbooks[currentLogbook].flights = [];
        airportFrequency = {};
        routeFrequency.clear();
        missingAirports.clear();
        totalFlightTime = totalCrossCountry = totalNight = totalSolo = totalActualIFR = 0;
        drawFlights();
    });

    document.getElementById('download-btn').addEventListener('click', () => {
        if (isDownloading) return;
        isDownloading = true;
        loadingIndicator.classList.remove('hidden');
        document.getElementById('download-btn').disabled = true;

        leafletImage(map, (err, canvas) => {
            loadingIndicator.classList.add('hidden');
            document.getElementById('download-btn').disabled = false;
            isDownloading = false;

            if (err) {
                showFileError('Failed to download map. Please try again.', 'general');
                return;
            }
            const link = document.createElement('a');
            link.download = 'flight_map.png';
            link.href = canvas.toDataURL();
            link.click();
        });
    });

    document.getElementById('export-flights-btn').addEventListener('click', () => {
        const flights = logbooks[currentLogbook].flights;
        if (flights.length === 0) return showFileError('No flights to export.', 'general');

        const csv = [
            ['Departure', 'Arrival', 'Date', 'Distance (km)', 'Duration (hrs)', 'Aircraft Type', 'Registration', 'Cross-Country (hrs)', 'Night (hrs)', 'Solo (hrs)', 'Actual IFR (hrs)', 'Notes'].join(','),
            ...flights.map(flight => [
                flight.codes[0],
                flight.codes[1],
                flight.date,
                flight.distance.toFixed(2),
                flight.duration.toFixed(2),
                flight.aircraftType,
                flight.registration,
                flight.crossCountry.toFixed(2),
                flight.night.toFixed(2),
                flight.solo.toFixed(2),
                flight.actualIFR.toFixed(2),
                `"${flight.notes.replace(/"/g, '""')}"`
            ].join(','))
        ].join('\n');

        const blob = new Blob([csv], { type: 'text/csv' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `${currentLogbook}_flights.csv`;
        link.click();
    });

    document.getElementById('toggle-local-flights-btn').addEventListener('click', () => {
        showLocalFlights = !showLocalFlights;
        document.getElementById('toggle-local-flights-btn').querySelector('span:not([data-icon])').textContent = showLocalFlights ? 'Show/Hide Local Flights' : 'Show/Hide Local Flights';
        drawFlights();
    });

    document.getElementById('apply-date-filter').addEventListener('click', () => {
        startDate = document.getElementById('start-date').value;
        endDate = document.getElementById('end-date').value;
        drawFlights();
    });

    document.getElementById('apply-filter-sort').addEventListener('click', () => {
        const departureFilter = document.getElementById('filter-departure').value.toUpperCase();
        const arrivalFilter = document.getElementById('filter-arrival').value.toUpperCase();
        const sortOption = document.getElementById('sort-flights').value;

        let filteredFlights = logbooks[currentLogbook].flights.slice();

        if (startDate || endDate) {
            filteredFlights = filteredFlights.filter(flight => {
                const flightDate = new Date(flight.date);
                const start = startDate ? new Date(startDate) : new Date('1900-01-01');
                const end = endDate ? new Date(endDate) : new Date('9999-12-31');
                return flightDate >= start && flightDate <= end;
            });
        }

        if (departureFilter) {
            const depAirport = getAirport(departureFilter);
            depAirport && (filteredFlights = filteredFlights.filter(flight => flight.codes[0] === depAirport.iata));
        }
        if (arrivalFilter) {
            const arrAirport = getAirport(arrivalFilter);
            arrAirport && (filteredFlights = filteredFlights.filter(flight => flight.codes[1] === arrAirport.iata));
        }

        const sortFunctions = {
            'date-asc': (a, b) => new Date(a.date) - new Date(b.date),
            'date-desc': (a, b) => new Date(b.date) - new Date(a.date),
            'distance-asc': (a, b) => a.distance - b.distance,
            'distance-desc': (a, b) => b.distance - a.distance,
            'frequency-desc': (a, b) => (routeFrequency.get(`${b.codes[0]}-${b.codes[1]}`) || 0) - (routeFrequency.get(`${a.codes[0]}-${a.codes[1]}`) || 0),
        };

        if (sortFunctions[sortOption]) filteredFlights.sort(sortFunctions[sortOption]);

        const originalFlights = logbooks[currentLogbook].flights;
        logbooks[currentLogbook].flights = filteredFlights;
        drawFlights();
        logbooks[currentLogbook].flights = originalFlights;
    });

    document.getElementById('add-missing-airport').addEventListener('click', () => {
        const iata = document.getElementById('missing-airport-iata').value.toUpperCase();
        const icao = document.getElementById('missing-airport-icao').value.toUpperCase();
        const identifier = document.getElementById('missing-airport-identifier').value.toUpperCase() || null;
        const country = document.getElementById('missing-airport-country').value.toUpperCase();
        const lat = parseFloat(document.getElementById('missing-airport-lat').value);
        const lon = parseFloat(document.getElementById('missing-airport-lon').value);
        const name = document.getElementById('missing-airport-name').value || `Custom Airport (${iata})`;

        const isValidIATA = /^[A-Z]{3}$/.test(iata);
        const isValidICAO = /^[A-Z]{4}$/.test(icao);
        const isValidCountry = /^[A-Z]{2}$/.test(country);
        const isValidIdentifier = identifier ? /^[A-Z]{3,4}$/.test(identifier) : true;

        if (isValidIATA && isValidICAO && isValidCountry && isValidIdentifier && !isNaN(lat) && !isNaN(lon)) {
            const airportId = airportsData.airports.length;
            const newAirport = {
                id: airportId,
                iata,
                icao,
                identifier: isValidIdentifier ? identifier : null,
                country,
                coords: [lat, lon],
                name
            };
            airportsData.airports.push(newAirport);
            airportsData.code_to_airport_id[iata] = airportId;
            airportsData.code_to_airport_id[icao] = airportId;
            if (identifier) airportsData.code_to_airport_id[identifier] = airportId;
            validAirportCodes.add(iata).add(icao);
            if (identifier) validAirportCodes.add(identifier);
            missingAirports.delete(iata).delete(icao);
            if (identifier) missingAirports.delete(identifier);
            showFileError(`Added airport ${iata} (${name}) at (${lat}, ${lon}). Re-upload your logbook to include flights with this airport.`, 'general');
        } else {
            showFileError('Please provide valid IATA (3 letters), ICAO (4 letters), country code (2 letters), latitude, longitude, and optionally an identifier (3-4 letters).', 'general');
        }

        ['missing-airport-iata', 'missing-airport-icao', 'missing-airport-identifier', 'missing-airport-country', 'missing-airport-lat', 'missing-airport-lon', 'missing-airport-name'].forEach(id => document.getElementById(id).value = '');
    });

    document.getElementById('toggle-missing-airport-btn').addEventListener('click', function () {
        const formContainer = document.getElementById('missing-airport-form-container');
        formContainer.classList.toggle('hidden');
        const chevron = this.querySelector('span[data-chevron="missing-airport"]');
        chevron.classList.toggle('fa-chevron-right');
        chevron.classList.toggle('fa-chevron-down');
        this.querySelector('span:not([data-chevron])').textContent = formContainer.classList.contains('hidden') ? 'Add Missing Airport' : 'Hide Missing Airport';
    });

    document.getElementById('logbook-select').addEventListener('change', function () {
        currentLogbook = this.value;
        drawFlights();
    });

    document.getElementById('add-logbook').addEventListener('click', () => {
        const logbookCount = Object.keys(logbooks).length + 1;
        const newLogbookName = `Logbook ${logbookCount}`;
        logbooks[newLogbookName] = { flights: [], stats: {} };
        const option = document.createElement('option');
        option.value = newLogbookName;
        option.textContent = newLogbookName;
        document.getElementById('logbook-select').appendChild(option);
        document.getElementById('logbook-select').value = newLogbookName;
        currentLogbook = newLogbookName;
        drawFlights();
    });

    document.getElementById('logbook-upload').addEventListener('change', function (event) {
        const file = event.target.files[0];
        if (!file) {
            showFileError('No file selected.', 'file-format');
            return;
        }

        const fileExtension = file.name.split('.').pop().toLowerCase();
        if (['csv', 'pdf', 'txt', 'jpg', 'png', 'jpeg'].includes(fileExtension)) {
            const reader = new FileReader();
            if (fileExtension === 'csv' || fileExtension === 'txt') {
                reader.onload = e => fileExtension === 'csv' ? parseCSV(e.target.result) : parseText(e.target.result);
                reader.readAsText(file);
            } else if (fileExtension === 'pdf') {
                parsePDF(file);
            } else {
                reader.onload = e => parseImage(e.target.result);
                reader.readAsDataURL(file);
            }
        } else {
            showFileError('Unsupported file format. Please upload a CSV, PDF, TXT, JPG, PNG, or JPEG file.', 'file-format');
        }

        this.value = '';
    });

    fetch('airports.json')
        .then(response => response.json())
        .then(data => {
            airportsData = data;
            airportsData.code_to_airport = {};
            Object.entries(airportsData.code_to_airport_id).forEach(([code, id]) => {
                airportsData.code_to_airport[code] = airportsData.airports[id];
            });
            validAirportCodes = new Set(Object.keys(airportsData.code_to_airport_id));
        })
        .catch(() => showFileError('Failed to load airport data. Please try again.', 'file-format'));

    const editMenuBtn = document.getElementById('edit-menu-btn');
    const reorderMenuOverlay = document.getElementById('reorder-menu-overlay');
    const closeReorderMenuBtn = document.getElementById('close-reorder-menu');
    const reorderList = document.getElementById('reorder-list');
    const sectionsContainer = document.getElementById('menu-sections');

    if (editMenuBtn && reorderMenuOverlay && closeReorderMenuBtn && reorderList && sectionsContainer) {
        const sectionsList = [
            { id: 'flight-management', name: 'Flight Management' },
            { id: 'map-controls', name: 'Map Controls' },
            { id: 'filters-sorting', name: 'Filters & Sorting' },
            { id: 'logbook-management', name: 'Logbook Management' },
            { id: 'add-missing-airport', name: 'Add Missing Airport' },
            { id: 'flight-stats', name: 'Flight Stats' },
            { id: 'column-mapping', name: 'Column Mapping' }
        ];

        function loadSectionOrder() {
            const savedOrder = localStorage.getItem('menuSectionOrder');
            if (savedOrder) {
                const order = JSON.parse(savedOrder);
                const sections = Array.from(sectionsContainer.children);
                order.forEach(sectionId => {
                    const section = sections.find(s => s.dataset.sectionId === sectionId);
                    section && sectionsContainer.appendChild(section);
                });
            }
            populateReorderList();
        }

        function populateReorderList() {
            reorderList.innerHTML = '';
            Array.from(sectionsContainer.children).forEach(section => {
                const sectionId = section.dataset.sectionId;
                const sectionName = sectionsList.find(s => s.id === sectionId).name;
                const item = document.createElement('div');
                item.className = 'reorder-item flex items-center p-2 bg-gray-100 rounded-md cursor-move';
                item.draggable = true;
                item.dataset.sectionId = sectionId;
                const handle = document.createElement('span');
                handle.className = 'drag-handle mr-2 text-gray-600';
                handle.innerHTML = '<span data-icon="bars"></span>';
                const label = document.createElement('span');
                label.textContent = sectionName;
                item.appendChild(handle);
                item.appendChild(label);
                reorderList.appendChild(item);
            });
        }

        function saveSectionOrder() {
            const sections = Array.from(sectionsContainer.children);
            const order = sections.map(section => section.dataset.sectionId);
            localStorage.setItem('menuSectionOrder', JSON.stringify(order));
        }

        editMenuBtn.addEventListener('click', () => reorderMenuOverlay.classList.add('open'));

        closeReorderMenuBtn.addEventListener('click', () => reorderMenuOverlay.classList.remove('open'));

        let draggedItem = null;

        reorderList.addEventListener('dragstart', e => {
            draggedItem = e.target.closest('.reorder-item');
            draggedItem?.classList.add('dragging');
        });

        reorderList.addEventListener('dragend', () => {
            if (draggedItem) {
                draggedItem.classList.remove('dragging');
                const reorderedItems = Array.from(reorderList.children);
                const sections = Array.from(sectionsContainer.children);
                reorderedItems.forEach(item => {
                    const sectionId = item.dataset.sectionId;
                    const section = sections.find(s => s.dataset.sectionId === sectionId);
                    section && sectionsContainer.appendChild(section);
                });
                saveSectionOrder();
                draggedItem = null;
            }
        });

        reorderList.addEventListener('dragover', e => e.preventDefault());

        reorderList.addEventListener('drop', e => {
            e.preventDefault();
            if (!draggedItem) return;
            const targetItem = e.target.closest('.reorder-item');
            if (!targetItem || targetItem === draggedItem) return;
            const allItems = Array.from(reorderList.children);
            const draggedIndex = allItems.indexOf(draggedItem);
            const targetIndex = allItems.indexOf(targetItem);
            if (draggedIndex < targetIndex) {
                targetItem.after(draggedItem);
            } else {
                targetItem.before(draggedItem);
            }
        });

        loadSectionOrder();
    }
});

// --- Functions and Parsing Logic ---

// --- CSV Parsing Logic ---

headerKeywords = [
    'DATE', 'AIRCRAFT', 'FROM', 'TO', 'ROUTE', 'TYPE', 'REGISTRATION',
    'PILOT', 'COMMAND', 'FLIGHT', 'NUMBER', 'MULTI', 'ASEL', 'ASES',
    'AMEL', 'AMES', 'JET', 'TURBO', 'PROP', 'ROTOR', 'TAKEOFFS', 'LANDINGS',
    'DAY', 'NIGHT', 'D', 'A', 'Y', 'N', 'I', 'G', 'H', 'T', 'dd.MMM.yy',
    'OUT', 'IN', 'DEP', 'ARR', 'DEPARTURE', 'ARRIVAL'
];

summaryIndicators = ['TOTAL', 'FORWARDED', 'PAGE', 'THIS', 'REPORT', 'SUBTOTAL', 'CARRIED FORWARD'];

function isAirportCode(str) {
    const excludedTerms = [
        'ILS', 'LOC', 'VOR', 'NDB', 'GPS', 'DME', 'RNAV', 'TACAN',
        'CRJ', 'FAA', 'A320', 'B737', 'B747', 'B757', 'B767', 'B777', 'B787',
        'E170', 'E175', 'E190', 'E195', 'FRASCA', 'AVENGER', 'SIM', 'FTD',
        'DAY', 'NIGHT', 'TYPE', 'TOTAL', 'OUT', 'IN', 'PIC', 'SIC', 'SOLO',
        'CFI', 'DPE', 'MEI', 'VMC', 'PPL', 'AND'
    ];
    const upperStr = str.toUpperCase();
    return /^[A-Z]{3,4}$/.test(str) &&
           !excludedTerms.includes(upperStr) &&
           validAirportCodes.has(upperStr);
}

function isLikelyHeader(row) {
    const upperRow = row.map(col => col.toUpperCase());
    let keywordCount = 0;
    let hasAirportCode = false;
    let isUppercase = true;
    let hasNumbers = false;

    for (const cell of upperRow) {
        if (headerKeywords.includes(cell)) keywordCount++;
        if (isAirportCode(cell)) hasAirportCode = true;
        if (cell !== cell.toUpperCase()) isUppercase = false;
        if (/\d/.test(cell)) hasNumbers = true;
    }

    const hasKeyHeaderTerms = upperRow.some(cell => ['FROM', 'TO', 'OUT', 'IN', 'ROUTE'].includes(cell));
    return (keywordCount >= 2 || (keywordCount >= 1 && hasKeyHeaderTerms && upperRow.length <= 10)) && !hasAirportCode && (!hasNumbers || isUppercase);
}

function getAirport(code, contextAirports = []) {
    if (!code) return null;
    const upperCode = code.toUpperCase();
    const matchingAirportIds = [];

    for (const [key, id] of Object.entries(airportsData.code_to_airport_id)) {
        if (key === upperCode || key.startsWith(`${upperCode}_`)) {
            const airport = airportsData.airports[id];
            if (airport) matchingAirportIds.push({ id, airport });
        }
    }

    if (matchingAirportIds.length === 0) {
        if (isAirportCode(upperCode)) {
            debugLog(`Unmatched airport code: ${upperCode}`);
            missingAirports.add(upperCode);
        }
        return null;
    }

    function validateAirportCode(airport, code, contextAirports) {
        let score = 0;
        const reasons = [];

        if (airport.icao === code) {
            score += 100;
            reasons.push('ICAO match');
        }
        if (airport.identifier === code) {
            score += 80;
            reasons.push('Identifier match');
        }
        if (airport.iata === code) {
            score += 50;
            reasons.push('IATA match');
        }
        if (airport.country === 'US' && code.length === 4 && code.startsWith('K')) {
            score += 30;
            reasons.push('US airport with 4-letter code');
        } else if (airport.country === 'US') {
            score += 20;
            reasons.push('US airport');
        }
        if (airport.iso_region === 'US-CA') {
            score += 10;
            reasons.push('California airport');
        }

        const kmyfCoords = [32.8157, -117.1396];
        const distance = calculateDistance(airport.coords, kmyfCoords);
        if (distance < 1000) {
            score += 50 / (distance + 1);
            reasons.push(`Close to KMYF (${distance.toFixed(2)} km)`);
        }

        if (contextAirports.length > 0 && contextAirports.includes(code)) {
            score += 70;
            reasons.push('Found in Remarks');
        }

        return { score, reasons };
    }

    const scoredCandidates = matchingAirportIds.map(({ id, airport }) => {
        const { score, reasons } = validateAirportCode(airport, upperCode, contextAirports);
        return { id, airport, score, reasons };
    });

    scoredCandidates.sort((a, b) => b.score - a.score);
    const bestCandidate = scoredCandidates[0];

    if (bestCandidate.score > 0) {
        debugLog(`Selected airport for ${upperCode}: ${bestCandidate.airport.name} (${bestCandidate.airport.iata || bestCandidate.airport.icao}), Score: ${bestCandidate.score}, Reasons: ${bestCandidate.reasons.join(', ')}`);
        return bestCandidate.airport;
    }

    if (isAirportCode(upperCode)) {
        debugLog(`Unmatched airport code: ${upperCode}`);
        missingAirports.add(upperCode);
    }
    return null;
}
// Utility: Check if both departure and arrival are valid airport codes and exist
function isValidAirportPair(depCode, arrCode) {
    return isAirportCode(depCode) && isAirportCode(arrCode) &&
           getAirport(depCode) && getAirport(arrCode);
}

// Utility: Normalize a 3-letter airport code to its ICAO equivalent if present in icaoList
function normalizeAirportCode(code, icaoList) {
    if (code.length === 3 && icaoList?.length) {
      const match = icaoList.find(icao => icao.slice(1) === code);
      return match || code;
    }
    return code;
}

// Utility: Parse a CSV row by delimiter, trimming and removing quotes
function parseRow(line, delimiter) {
    return line.split(delimiter).map(cell => cell.trim().replace(/"/g, ''));
}


function getCoords(airport) {
    return airport ? airport.coords : null;
}

function getAirportName(airport) {
    return airport ? airport.name : 'Unknown Airport';
}

function updateAirportFrequency(departure, arrival) {
    airportFrequency[departure] = (airportFrequency[departure] || 0) + 1;
    airportFrequency[arrival] = (airportFrequency[arrival] || 0) + 1;
}

function updateRouteFrequency(departure, arrival) {
    const routeKey = `${departure}-${arrival}`;
    routeFrequency.set(routeKey, (routeFrequency.get(routeKey) || 0) + 1);
}

function calculateDistance(coords1, coords2) {
    const key = `${coords1[0]},${coords1[1]}-${coords2[0]},${coords2[1]}`;
    if (distanceCache.has(key)) {
        return distanceCache.get(key);
    }
    const R = 6371;
    const lat1 = coords1[0] * Math.PI / 180;
    const lon1 = coords1[1] * Math.PI / 180;
    const lat2 = coords2[0] * Math.PI / 180;
    const lon2 = coords2[1] * Math.PI / 180;
    const dLat = lat2 - lat1;
    const dLon = lon2 - lon1;
    const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
              Math.cos(lat1) * Math.cos(lat2) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    const distance = R * c;
    distanceCache.set(key, distance);
    return distance;
}

// Detect and validate dates in multiple formats
function findDateInRow(row) {
    const dateFormats = [
        { regex: /^\d{1,2}\.\w{3}\.\d{2}$/, splitter: '.', parser: parts => {
            const month = parts[1].toLowerCase();
            const validMonths = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
            if (validMonths.includes(month)) {
                const day = parts[0].padStart(2, '0');
                const year = `20${parts[2]}`;
                const monthNum = (validMonths.indexOf(month) + 1).toString().padStart(2, '0');
                return `${year}-${monthNum}-${day}`;
            }
        }},
        { regex: /^\d{1,2}\/\d{1,2}\/\d{4}$/, splitter: '/', parser: parts => {
            const month = parseInt(parts[0], 10).toString().padStart(2, '0');
            const day = parseInt(parts[1], 10).toString().padStart(2, '0');
            const year = parts[2];
            if (parseInt(day, 10) <= 31 && parseInt(month, 10) <= 12) return `${year}-${month}-${day}`;
        }},
        { regex: /^\d{1,2}\/\d{1,2}\/\d{2}$/, splitter: '/', parser: parts => {
            const month = parseInt(parts[0], 10).toString().padStart(2, '0');
            const day = parseInt(parts[1], 10).toString().padStart(2, '0');
            let year = parts[2];
            year = parseInt(year, 10) < 50 ? `20${year}` : `19${year}`;
            if (parseInt(day, 10) <= 31 && parseInt(month, 10) <= 12) return `${year}-${month}-${day}`;
        }},
        { regex: /^\d{1,2}-\d{1,2}-\d{2}$/, splitter: '-', parser: parts => {
            const month = parseInt(parts[0], 10).toString().padStart(2, '0');
            const day = parseInt(parts[1], 10).toString().padStart(2, '0');
            let year = parts[2];
            year = parseInt(year, 10) < 50 ? `20${year}` : `19${year}`;
            if (parseInt(day, 10) <= 31 && parseInt(month, 10) <= 12) return `${year}-${month}-${day}`;
        }},
        { regex: /^\d{4}-\d{1,2}-\d{1,2}$/, splitter: '-', parser: parts => {
            const year = parts[0];
            const month = parseInt(parts[1], 10).toString().padStart(2, '0');
            const day = parseInt(parts[2], 10).toString().padStart(2, '0');
            if (parseInt(day, 10) <= 31 && parseInt(month, 10) <= 12) return `${year}-${month}-${day}`;
        }},
        { regex: /^\d{4}\/\d{1,2}\/\d{1,2}$/, splitter: '/', parser: parts => {
            const year = parts[0];
            const month = parseInt(parts[1], 10).toString().padStart(2, '0');
            const day = parseInt(parts[2], 10).toString().padStart(2, '0');
            if (parseInt(day, 10) <= 31 && parseInt(month, 10) <= 12) return `${year}-${month}-${day}`;
        }},
        { regex: /^\d{1,2}-\d{1,2}-\d{4}$/, splitter: '-', parser: parts => {
            const day = parseInt(parts[0], 10).toString().padStart(2, '0');
            const month = parseInt(parts[1], 10).toString().padStart(2, '0');
            const year = parts[2];
            if (parseInt(day, 10) <= 31 && parseInt(month, 10) <= 12) return `${year}-${month}-${day}`;
        }},
        { regex: /^(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4}$/i, splitter: null, parser: cell => {
            const parts = cell.match(/(\w+)\s+(\d{1,2}),\s+(\d{4})/);
            if (parts) {
                const month = parts[1].toLowerCase();
                const validMonths = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
                if (validMonths.includes(month)) {
                    const day = parseInt(parts[2], 10).toString().padStart(2, '0');
                    const year = parts[3];
                    const monthNum = (validMonths.indexOf(month) + 1).toString().padStart(2, '0');
                    return `${year}-${monthNum}-${day}`;
                }
            }
        }},
        { regex: /^\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}$/i, splitter: null, parser: cell => {
            const parts = cell.match(/(\d{1,2})\s+(\w+)\s+(\d{4})/);
            if (parts) {
                const day = parseInt(parts[1], 10).toString().padStart(2, '0');
                const month = parts[2].toLowerCase();
                const year = parts[3];
                const validMonths = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
                if (validMonths.includes(month)) {
                    const monthNum = (validMonths.indexOf(month) + 1).toString().padStart(2, '0');
                    return `${year}-${monthNum}-${day}`;
                }
            }
        }},
        { regex: /^\d{4}\.\d{1,2}\.\d{1,2}$/, splitter: '.', parser: parts => {
            const year = parts[0];
            const month = parseInt(parts[1], 10).toString().padStart(2, '0');
            const day = parseInt(parts[2], 10).toString().padStart(2, '0');
            if (parseInt(day, 10) <= 31 && parseInt(month, 10) <= 12) return `${year}-${month}-${day}`;
        }},
        { regex: /^\d{1,2}\.\d{1,2}\.\d{4}$/, splitter: '.', parser: parts => {
            const month = parseInt(parts[0], 10).toString().padStart(2, '0');
            const day = parseInt(parts[1], 10).toString().padStart(2, '0');
            const year = parts[2];
            if (parseInt(day, 10) <= 31 && parseInt(month, 10) <= 12) return `${year}-${month}-${day}`;
        }}
    ];

    debugLog(`[findDateInRow] Scanning row: ${row.join(' | ')}`);

    for (const cell of row) {
        if (!cell || typeof cell !== 'string') continue;
        const trimmedCell = cell.trim();
        for (const { regex, splitter, parser } of dateFormats) {
            if (regex.test(trimmedCell)) {
                let parsedDate;
                if (splitter) {
                    const parts = trimmedCell.split(splitter);
                    parsedDate = parser(parts);
                } else {
                    parsedDate = parser(trimmedCell);
                }
                if (parsedDate) {
                    const dateObj = new Date(parsedDate);
                    if (!isNaN(dateObj.getTime())) {
                        debugLog(`[findDateInRow] Found valid date: ${parsedDate} from cell: ${trimmedCell}`);
                        return parsedDate;
                    } else {
                        debugLog(`[findDateInRow] Invalid date parsed: ${parsedDate} from cell: ${trimmedCell}`);
                    }
                }
            }
        }
    }

    debugLog(`[findDateInRow] No valid date found in row: ${row.join(' | ')}`);
    const fallbackDate = new Date().toISOString().split('T')[0];
    debugLog(`[findDateInRow] Using fallback date: ${fallbackDate}`);
    return fallbackDate;
}

function extractDuration(row) {
    const durationRegex = /^\d+:\d{2}$/;
    for (const cell of row) {
        if (durationRegex.test(cell)) {
            const [hours, minutes] = cell.split(':').map(Number);
            return hours + minutes / 60;
        }
        if (!isNaN(parseFloat(cell)) && parseFloat(cell) > 0) {
            return parseFloat(cell);
        }
    }
    return 0;
}

function extractTime(row, label) {
    const durationRegex = /^\d+:\d{2}$/;
    for (let i = 0; i < row.length; i++) {
        if (row[i].toLowerCase().includes(label) && durationRegex.test(row[i + 1])) {
            const [hours, minutes] = row[i + 1].split(':').map(Number);
            return hours + minutes / 60;
        }
        if (row[i].toLowerCase().includes(label) && !isNaN(parseFloat(row[i + 1])) && parseFloat(row[i + 1]) > 0) {
            return parseFloat(row[i + 1]);
        }
    }
    return 0;
}

function extractAircraftInfo(row) {
    let aircraftType = 'Unknown';
    let registration = 'Unknown';
    const aircraftTerms = ['CRJ', 'A320', 'B737', 'E170', 'E175', 'FRASCA', 'AVENGER', 'FAA', 'P28A'];
    for (let i = 0; i < row.length; i++) {
        const cell = row[i].toUpperCase();
        let foundAircraft = aircraftTerms.find(term => cell.includes(term));
        if (foundAircraft) {
            if (cell.includes('FRASCA') || cell.includes('AVENGER') || cell.includes('FAA')) {
                aircraftType = row[i];
            } else {
                aircraftType = foundAircraft;
            }
            if (i > 0 && row[i - 1].match(/^N\d+[A-Z]+$/)) {
                registration = row[i - 1];
            } else if (i + 1 < row.length && row[i + 1].match(/^N\d+[A-Z]+$/)) {
                registration = row[i + 1];
            }
            break;
        }
        if (cell.match(/^[A-Z][0-9]+$/)) {
            aircraftType = cell;
            if (i + 1 < row.length && row[i + 1].match(/^N\d+[A-Z]+$/)) {
                registration = row[i + 1];
            }
            break;
        }
    }
    return { aircraftType, registration };
}

function cleanOCRText(text) {
    text = text.replace(/([A-Z]{3,4})[-→]([A-Z]{3,4})/g, '$1___SEPARATOR___$2');
    text = text.replace(/\$\{[^}]*\}|\$\mathbb{[A-Z]}/g, ' ')
               .replace(/COMPILYY/g, 'COMPANY')
               .replace(/FLOT IN/g, 'FLIGHT IN')
               .replace(/FUGHT/g, 'FLIGHT')
               .replace(/REPGKT/g, 'REPORT')
               .replace(/[^\w\s.:-]/g, ' ')
               .replace(/\s+/g, ' ')
               .trim();
    text = text.replace(/___SEPARATOR___/g, '-');
    return text;
}

function debounce(func, wait) {
    let timeout;
    return function (...args) {
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(this, args), wait);
    };
}

const debounceParseCSV = debounce(parseCSV, 100);

function parseCSV(csvText) {
    const lines = csvText.split('\n').filter(line => line.trim() !== '');
    if (lines.length < 1) {
        showFileError('Empty CSV file. Please ensure the file contains flight data.', 'file-format');
        return;
    }

    const possibleDelimiters = [',', ';', '\t'];
    let delimiter = possibleDelimiters.find(d => lines[0].includes(d)) || ',';

    const headers = lines[0].split(delimiter).map(h => h.trim().toLowerCase());
    let depIndex = userDefinedFromIndex !== -1 ? userDefinedFromIndex : headers.findIndex(h => ['from', 'departure', 'dep', 'route'].includes(h));
    let arrIndex = userDefinedToIndex !== -1 ? userDefinedToIndex : headers.findIndex(h => ['to', 'arrival', 'arr', 'route'].includes(h));
    let dateIndex = headers.findIndex(h => ['date'].includes(h));

    if (depIndex === -1 || arrIndex === -1) {
        depIndex = userDefinedFromIndex !== -1 ? userDefinedFromIndex : headers.findIndex(h => h.includes('from') || h.includes('dep') || h.includes('route'));
        arrIndex = userDefinedToIndex !== -1 ? userDefinedToIndex : headers.findIndex(h => h.includes('to') || h.includes('arr') || h.includes('route'));
    }

    const newFlights = [];
    if (depIndex !== -1 && arrIndex !== -1) {
        for (let i = 1; i < lines.length; i++) {
            const row = parseRow(lines[i], delimiter);
            if (row.length < Math.max(depIndex, arrIndex) + 1) continue;

            let departure = '', arrival = '';
            const routeCell = depIndex === arrIndex ? row[depIndex] : null;
            if (routeCell && routeCell.includes('-')) {
                const parts = routeCell.split('-').map(part => part.trim().toUpperCase());
                if (parts.length >= 2 && isAirportCode(parts[0]) && isAirportCode(parts[1])) {
                    departure = parts[0];
                    arrival = parts[1];
                }
            } else {
                departure = row[depIndex].toUpperCase();
                arrival = row[arrIndex].toUpperCase();
            }
            const date = dateIndex !== -1 ? row[dateIndex] : findDateInRow(row);

            const depAirport = getAirport(departure);
            const arrAirport = getAirport(arrival);
            const aircraftInfo = extractAircraftInfo(row);

            if (isAirportCode(departure) && isAirportCode(arrival) && depAirport && arrAirport) {
                const distance = calculateDistance(depAirport.coords, arrAirport.coords);
                newFlights.push({
                    coords: [depAirport.coords, arrAirport.coords],
                    codes: [depAirport.iata || depAirport.identifier, arrAirport.iata || arrAirport.identifier],
                    date: date,
                    distance: distance,
                    duration: extractDuration(row),
                    notes: '',
                    aircraftType: aircraftInfo.aircraftType,
                    registration: aircraftInfo.registration,
                    crossCountry: extractTime(row, 'cross country'),
                    night: extractTime(row, 'night'),
                    solo: extractTime(row, 'solo'),
                    actualIFR: extractTime(row, 'actual ifr'),
                    isLocal: depAirport.iata === arrAirport.iata || depAirport.identifier === arrAirport.identifier
                });
                updateAirportFrequency(depAirport.iata || depAirport.identifier, arrAirport.iata || arrAirport.identifier);
                updateRouteFrequency(depAirport.iata || depAirport.identifier, arrAirport.iata || arrAirport.identifier);
            } else {
                if (isAirportCode(departure) && !depAirport) missingAirports.add(departure);
                if (isAirportCode(arrival) && !arrAirport) missingAirports.add(arrival);
            }
        }
    }

    if (newFlights.length > 0) {
        logbooks[currentLogbook].flights.push(...newFlights);
        drawFlights(true);
        userDefinedFromIndex = -1;
        userDefinedToIndex = -1;
        pendingCSVText = null;
    } else {
        const rows = lines.map(line => parseRow(line, delimiter));
        showColumnMappingUI(rows, true, csvText);
    }
}

async function showColumnMappingUI(rows, isCSV = false, csvText = null) {
    const previewContainer = document.getElementById('logbook-preview');
    const columnMappingContainer = document.getElementById('column-mapping-container');
    const menuBackdrop = document.getElementById('menu-backdrop');
    const menuToggle = document.getElementById('menu-toggle');

    userDefinedFromIndex = -1;
    userDefinedToIndex = -1;

    if (isCSV && csvText) {
        pendingCSVText = csvText;
    } else {
        pendingCSVText = null;
    }

    columnMappingContainer.classList.remove('hidden');
    menuBackdrop.classList.remove('hidden');
    menuToggle.classList.remove('hidden');

    const maxColumns = Math.max(...rows.map(row => row.length));
    let previewHTML = '<table><thead><tr>';
    for (let i = 0; i < maxColumns; i++) {
        previewHTML += `<th class="clickable-header" data-index="${i}" title="Click to assign as FROM or TO">Col ${i + 1}</th>`;
    }
    previewHTML += '</tr></thead><tbody>';
    const previewRows = rows.slice(0, 5);
    for (const row of previewRows) {
        previewHTML += '<tr>';
        for (let i = 0; i < maxColumns; i++) {
            previewHTML += `<td>${row[i] || ''}</td>`;
        }
        previewHTML += '</tr>';
    }
    previewHTML += '</tbody></table>';
    previewContainer.innerHTML = previewHTML;

    const headers = previewContainer.querySelectorAll('.clickable-header');
    headers.forEach(header => {
        header.addEventListener('click', function () {
            const index = this.getAttribute('data-index');
            const assignAs = prompt('Assign this column as (FROM or TO):').toUpperCase();
            if (assignAs === 'FROM') {
                userDefinedFromIndex = parseInt(index);
                this.style.backgroundColor = '#d4edda';
            } else if (assignAs === 'TO') {
                userDefinedToIndex = parseInt(index);
                this.style.backgroundColor = '#cce5ff';
            } else {
                alert('Invalid assignment. Please enter FROM or TO.');
            }
        });
    });

    const confirmBtn = document.createElement('button');
    confirmBtn.textContent = 'Confirm Mapping';
    confirmBtn.className = 'w-full bg-blue-600 text-white p-2 rounded-md hover:bg-blue-700 mt-4';
    confirmBtn.addEventListener('click', () => {
        if (userDefinedFromIndex !== -1 && userDefinedToIndex !== -1 && userDefinedFromIndex !== userDefinedToIndex) {
            columnMappingContainer.classList.add('hidden');
            menuBackdrop.classList.add('hidden');
            menuToggle.classList.remove('hidden');
            if (isCSV && pendingCSVText) {
                parseCSV(pendingCSVText);
            } else if (pendingLogbookData) {
                parsePDF(pendingLogbookData);
            } else {
                showFileError('No logbook data available to reprocess. Please upload the logbook again.', 'parsing');
            }
        } else {
            showFileError('Please assign different columns for FROM and TO.', 'parsing');
        }
    });
    previewContainer.appendChild(confirmBtn);
}

// Process a single row asynchronously
async function processRow(row, i, j, fromIndex, toIndex) {
    return new Promise(resolve => {
        let departure = '', arrival = '';
        // Try to find date in a column labeled 'DATE' first
        let date = 'N/A';
        if (fromIndex !== -1 && toIndex !== -1) {
            const headers = row.slice(0, Math.max(fromIndex, toIndex) + 1).map(col => col.toUpperCase());
            const dateIndex = headers.findIndex(col => col.includes('DATE'));
            if (dateIndex !== -1 && row[dateIndex]) {
                const tempRow = [row[dateIndex]];
                date = findDateInRow(tempRow);
                console.log(`Date from DATE column (index ${dateIndex}): ${date}`);
            }
        }
        // Fall back to scanning all cells
        if (date === 'N/A') {
            date = findDateInRow(row);
            console.log(`Date from full row scan: ${date}`);
        }

        const duration = extractDuration(row);
        const crossCountry = extractTime(row, 'cross country');
        const night = extractTime(row, 'night');
        const solo = extractTime(row, 'solo');
        const actualIFR = extractTime(row, 'actual ifr');
        const { aircraftType, registration } = extractAircraftInfo(row);
        const upperRow = row.map(col => col.toUpperCase());

        if (upperRow.includes('TOTAL') || upperRow.includes('FORWARDED') || 
            upperRow.includes('PAGE') || upperRow.includes('CERTIFICATES') ||
            isLikelyHeader(row)) {
            resolve(null);
            return;
        }

        let remarksAirports = [];
        const narrativeText = row.join(' ').toUpperCase();
        const airportMatches = narrativeText.match(/\b([A-Z]{4})\b(?=\s*(?:,|to|-|→|back to|\s|$))/g) || [];
        const countryPrefixes = ['K', 'C', 'EG', 'L', 'E', 'Y', 'Z', 'V', 'W', 'T', 'U', 'O', 'F', 'S', 'R', 'M', 'N', 'B', 'D', 'G', 'H', 'I', 'P', 'A'];
        remarksAirports = airportMatches.filter(code => 
            /^[A-Z]{4}$/.test(code) && validAirportCodes.has(code) &&
            countryPrefixes.some(prefix => code.startsWith(prefix))
        );
        if (remarksAirports.length > 0) {
            debugLog(`Row ${j} Remarks airports:`, remarksAirports);
        }

        if (remarksAirports.length >= 2) {
            const flights = [];
            for (let p = 0; p < remarksAirports.length - 1; p++) {
                const depCode = remarksAirports[p];
                const arrCode = remarksAirports[p + 1];
                const depAirport = getAirport(depCode, remarksAirports);
                const arrAirport = getAirport(arrCode, remarksAirports);
                if (depAirport && arrAirport) {
                    flights.push({
                        departure: depAirport.iata || depAirport.identifier,
                        arrival: arrAirport.iata || arrAirport.identifier,
                        depAirport,
                        arrAirport,
                        date,
                        distance: calculateDistance(depAirport.coords, arrAirport.coords),
                        duration: duration / (remarksAirports.length - 1),
                        crossCountry,
                        night,
                        solo,
                        actualIFR,
                        aircraftType,
                        registration,
                        notes: `From remarks: ${depCode} to ${arrCode}`,
                        isLocal: depCode === arrCode
                    });
                }
            }
            if (flights.length > 0) {
                debugLog(`Row ${j} airports from remarks:`, flights.map(f => `${f.departure} to ${f.arrival}`));
                resolve(flights);
                return;
            }
        }

        if (userDefinedFromIndex !== -1 && userDefinedToIndex !== -1 && 
            row.length > Math.max(userDefinedFromIndex, userDefinedToIndex)) {
            let fromCode = row[userDefinedFromIndex].toUpperCase();
            let toCode = row[userDefinedToIndex].toUpperCase();
            fromCode = normalizeAirportCode(fromCode, remarksAirports);
            toCode = normalizeAirportCode(toCode, remarksAirports);
            if (isAirportCode(fromCode) && validAirportCodes.has(fromCode) &&
                isAirportCode(toCode) && validAirportCodes.has(toCode)) {
                departure = fromCode;
                arrival = toCode;
                debugLog(`Row ${j} FROM/TO airports: ${departure} to ${arrival}`);
            }
        }

        if (!departure && fromIndex !== -1 && toIndex !== -1 && 
            row.length > Math.max(fromIndex, toIndex)) {
            if (fromIndex === toIndex) {
                const routeCell = row[fromIndex].toUpperCase();
                const parts = routeCell.split('-').map(part => part.trim());
                if (parts.length >= 2 && isAirportCode(parts[0]) && validAirportCodes.has(parts[0]) &&
                    isAirportCode(parts[1]) && validAirportCodes.has(parts[1])) {
                    let depCode = normalizeAirportCode(parts[0], remarksAirports);
                    let arrCode = normalizeAirportCode(parts[1], remarksAirports);
                    departure = depCode;
                    arrival = arrCode;
                    if (parts.length > 2) {
                        const flights = [];
                        for (let p = 0; p < parts.length - 1; p++) {
                            let depCode = normalizeAirportCode(parts[p], remarksAirports);
                            let arrCode = normalizeAirportCode(parts[p + 1], remarksAirports);
                            if (isAirportCode(depCode) && validAirportCodes.has(depCode) &&
                                isAirportCode(arrCode) && validAirportCodes.has(arrCode)) {
                                const depAirport = getAirport(depCode, remarksAirports);
                                const arrAirport = getAirport(arrCode, remarksAirports);
                                if (depAirport && arrAirport) {
                                    flights.push({
                                        departure: depAirport.iata || depAirport.identifier,
                                        arrival: arrAirport.iata || arrAirport.identifier,
                                        depAirport,
                                        arrAirport,
                                        date,
                                        distance: calculateDistance(depAirport.coords, arrAirport.coords),
                                        duration: duration / (parts.length - 1),
                                        crossCountry,
                                        night,
                                        solo,
                                        actualIFR,
                                        aircraftType,
                                        registration,
                                        notes: `Segment ${p + 1}: ${depCode} to ${arrCode}`,
                                        isLocal: depCode === arrCode
                                    });
                                }
                            }
                        }
                        if (flights.length > 0) {
                            debugLog(`Row ${j} multi-segment route:`, flights.map(f => `${f.departure} to ${f.arrival}`));
                            resolve(flights);
                            return;
                        }
                    }
                } else if (parts.length === 1 && isAirportCode(parts[0]) && validAirportCodes.has(parts[0])) {
                    let routeAirport = normalizeAirportCode(parts[0], remarksAirports);
                    if (fromIndex !== -1 && row[fromIndex]) {
                        let fromCode = row[fromIndex].toUpperCase();
                        fromCode = normalizeAirportCode(fromCode, remarksAirports);
                        if (isAirportCode(fromCode) && validAirportCodes.has(fromCode)) {
                            departure = fromCode;
                            arrival = routeAirport;
                            console.log(`Row ${j} FROM/Route airports: ${departure} to ${arrival}`);
                        }
                    }
                }
                if (parts.length > 0) {
                    debugLog(`Row ${j} Route airports:`, parts.filter(part => isAirportCode(part) && validAirportCodes.has(part)));
                }
            } else {
                let fromCode = row[fromIndex].toUpperCase();
                let toCode = row[toIndex].toUpperCase();
                fromCode = normalizeAirportCode(fromCode, remarksAirports);
                toCode = normalizeAirportCode(toCode, remarksAirports);
                if (isAirportCode(fromCode) && validAirportCodes.has(fromCode)) {
                    departure = fromCode;
                }
                if (isAirportCode(toCode) && validAirportCodes.has(toCode)) {
                    arrival = toCode;
                }
                debugLog(`Row ${j} FROM/TO airports: ${departure || 'none'} to ${arrival || 'none'}`);
            }
        }

        if (!departure || !arrival) {
            let foundAirports = [];
            let routeSegments = [];

            for (let k = 0; k < row.length; k++) {
                let cell = row[k].toUpperCase();
                if (cell.includes('-')) {
                    const parts = cell.split('-').map(part => part.trim());
                    parts.forEach(part => {
                        if (isAirportCode(part) && validAirportCodes.has(part)) {
                            let code = normalizeAirportCode(part, remarksAirports);
                            foundAirports.push({ code, index: k });
                        }
                    });
                    routeSegments.push(...parts.filter(part => isAirportCode(part) && validAirportCodes.has(part)).map(part => normalizeAirportCode(part, remarksAirports)));
                } else {
                    const countryPrefixes = ['K', 'C', 'EG', 'L', 'E', 'Y', 'Z', 'V', 'W', 'T', 'U', 'O', 'F', 'S', 'R', 'M', 'N', 'B', 'D', 'G', 'H', 'I', 'P', 'A'];
                    if (/^[A-Z]{4}$/.test(cell) && validAirportCodes.has(cell) && 
                        countryPrefixes.some(prefix => cell.startsWith(prefix))) {
                        foundAirports.push({ code: cell, index: k });
                    }
                }
            }

            if (routeSegments.length >= 2) {
                const flights = [];
                for (let p = 0; p < routeSegments.length - 1; p++) {
                    const depCode = routeSegments[p];
                    const arrCode = routeSegments[p + 1];
                    const depAirport = getAirport(depCode, remarksAirports);
                    const arrAirport = getAirport(arrCode, remarksAirports);
                    if (depAirport && arrAirport) {
                        flights.push({
                            departure: depAirport.iata || depAirport.identifier,
                            arrival: arrAirport.iata || arrAirport.identifier,
                            depAirport,
                            arrAirport,
                            date,
                            distance: calculateDistance(depAirport.coords, arrAirport.coords),
                            duration: duration / (routeSegments.length - 1),
                            crossCountry,
                            night,
                            solo,
                            actualIFR,
                            aircraftType,
                            registration,
                            notes: `Segment ${p + 1}: ${depCode} to ${arrCode}`,
                            isLocal: depCode === arrCode
                        });
                    }
                }
                if (flights.length > 0) {
                    debugLog(`Row ${j} multi-segment route:`, flights.map(f => `${f.departure} to ${f.arrival}`));
                    resolve(flights);
                    return;
                }
            }

            if (departure && !arrival && (foundAirports.length > 0 || remarksAirports.length > 0)) {
                const allAirports = [...foundAirports.map(a => a.code), ...remarksAirports];
                const nextAirport = allAirports.find(code => code !== departure);
                if (nextAirport) {
                    arrival = nextAirport;
                }
            } else if (!departure && (foundAirports.length >= 2 || remarksAirports.length >= 2)) {
                const allAirports = [...foundAirports.map(a => ({ code: a.code, index: a.index })), ...remarksAirports.map((code, idx) => ({ code, index: idx }))];
                allAirports.sort((a, b) => a.index - b.index);
                departure = allAirports[0]?.code;
                arrival = allAirports[1]?.code;
            } else if (!arrival && remarksAirports.length === 1) {
                arrival = remarksAirports[0];
                debugLog(`Row ${j} using single Remarks airport as arrival: ${arrival}`);
            }
        }

        if (departure && arrival) {
            const depAirport = getAirport(departure, remarksAirports);
            const arrAirport = getAirport(arrival, remarksAirports);
            if (depAirport && arrAirport) {
                debugLog(`Row ${j} final airports: ${departure} to ${arrival}, Date: ${date}`);
                resolve({
                    departure: depAirport.iata || depAirport.identifier,
                    arrival: arrAirport.iata || arrAirport.identifier,
                    depAirport,
                    arrAirport,
                    date,
                    distance: calculateDistance(depAirport.coords, arrAirport.coords),
                    duration,
                    crossCountry,
                    night,
                    solo,
                    actualIFR,
                    aircraftType,
                    registration,
                    notes: remarksAirports.length > 0 ? `From remarks: ${remarksAirports.join(', ')}` : '',
                    isLocal: departure === arrival
                });
            } else {
                debugLog(`Unmatched airports in row ${j}: ${departure}, ${arrival}`, row);
                resolve(null);
            }
        } else {
            debugLog(`No valid airports in row ${j}:`, row);
            resolve(null);
        }
    });
}

async function parsePDF(input) {
    try {
        let arrayBuffer;
        if (input instanceof File) {
            const fileType = input.name.split('.').pop().toLowerCase();
            if (fileType !== 'pdf') {
                showFileError(`Non-PDF file detected (${fileType}). Please upload a PDF file.`, 'file-format');
                return;
            }
            arrayBuffer = await input.arrayBuffer();
        } else {
            arrayBuffer = input;
        }

        if (arrayBuffer.byteLength === 0) {
            throw new Error("File or data is empty or could not be read");
        }

        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.worker.min.js';

        let pdf;
        try {
            pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        } catch (parseError) {
            throw new Error("Invalid or corrupted PDF file. Please ensure the file is a valid PDF.");
        }

        let allRows = [];
        missingAirports.clear();
        totalFlightTime = 0;
        totalCrossCountry = 0;
        totalNight = 0;
        totalSolo = 0;
        totalActualIFR = 0;

        const totalPages = pdf.numPages;
        let processedPages = 0;

        document.getElementById('progress-bar').style.display = 'block';
        const progressElement = document.getElementById('progress-bar').firstChild;

        for (let i = 1; i <= totalPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            let pageText = '';
            let currentLine = '';

            for (const item of textContent.items) {
                currentLine += item.str;
                if (item.hasEOL) {
                    pageText += currentLine + '\n';
                    currentLine = '';
                } else {
                    currentLine += ' ';
                }
            }
            if (currentLine) pageText += currentLine + '\n';

            const lines = pageText.split('\n').filter(line => line.trim() !== '');
            let rows = [];

            for (let line of lines) {
                line = cleanOCRText(line);
                if (line.includes('|')) {
                    rows.push(line.split('|').map(col => col.trim()));
                } else {
                    rows.push(line.split(/\s+/));
                }
            }

            let isSummaryPage = false;
            for (const row of rows) {
                const upperRow = row.map(col => col.toUpperCase());
                if (upperRow.includes('CONDITIONS OF FLIGHT') || upperRow.includes('SIMULATED') || upperRow.includes('CROSS COUNTRY') || upperRow.includes('CERTIFICATES')) {
                    isSummaryPage = true;
                    break;
                }
            }
            if (isSummaryPage) {
                processedPages++;
                progressElement.style.width = `${(processedPages / totalPages) * 100}%`;
                continue;
            }

            let headerIndices = [];
            for (let j = 0; j < rows.length; j++) {
                if (isLikelyHeader(rows[j])) {
                    headerIndices.push(j);
                } else {
                    break;
                }
            }

            let fromIndex = -1;
            let toIndex = -1;
            for (const j of headerIndices) {
                const row = rows[j];
                const upperRow = row.map(col => col.toUpperCase());
                const departureLabels = ['FROM', 'DEP', 'DEPARTURE', 'ROUTE'];
                const arrivalLabels = ['TO', 'ARR', 'ARRIVAL', 'ROUTE'];
                for (let k = 0; k < upperRow.length; k++) {
                    if (departureLabels.includes(upperRow[k]) && fromIndex === -1) fromIndex = k;
                    if (arrivalLabels.includes(upperRow[k]) && toIndex === -1) toIndex = k;
                }
                if (fromIndex !== -1 && toIndex !== -1) {
                    if (j + 1 < rows.length) {
                        const nextRow = rows[j + 1];
                        if (nextRow[fromIndex] && isAirportCode(nextRow[fromIndex]) && nextRow[toIndex] && isAirportCode(nextRow[toIndex])) {
                            break;
                        } else {
                            fromIndex = -1;
                            toIndex = -1;
                        }
                    }
                }
            }

            for (let j = 0; j < rows.length; j += BATCH_SIZE) {
                const batch = rows.slice(j, j + BATCH_SIZE).filter((_, idx) => !headerIndices.includes(j + idx));
                const batchPromises = batch.map((row, index) => {
                    const rowIndex = j + index;
                    const upperRow = row.map(col => col.toUpperCase());
                    if (summaryIndicators.some(indicator => upperRow.includes(indicator))) {
                        return Promise.resolve(null);
                    }
                    if (isLikelyHeader(row)) {
                        return Promise.resolve(null);
                    }
                    return processRow(row, i, rowIndex, fromIndex, toIndex).catch(error => {
                        console.error(`Error processing row ${rowIndex}:`, error);
                        return null;
                    });
                });

                const batchResults = await Promise.all(batchPromises);
                const validRows = batchResults.flat().filter(row => row !== null);
                allRows.push(...validRows);
                progressElement.style.width = `${((processedPages + j / rows.length) / totalPages) * 100}%`;
            }

            processedPages++;
        }

        document.getElementById('progress-bar').style.display = 'none';

        if (allRows.length > 0) {
            const newFlights = allRows.map(row => ({
                coords: [row.depAirport.coords, row.arrAirport.coords],
                codes: [row.departure, row.arrival],
                date: row.date,
                distance: row.distance,
                duration: row.duration,
                crossCountry: row.crossCountry,
                night: row.night,
                solo: row.solo,
                actualIFR: row.actualIFR,
                aircraftType: row.aircraftType,
                registration: row.registration,
                notes: row.notes,
                isLocal: row.isLocal,
                depAirport: row.depAirport,
                arrAirport: row.arrAirport
            }));
            logbooks[currentLogbook].flights.push(...newFlights);
            newFlights.forEach(flight => {
                updateAirportFrequency(flight.codes[0], flight.codes[1]);
                updateRouteFrequency(flight.codes[0], flight.codes[1]);
                totalFlightTime += flight.duration;
                totalCrossCountry += flight.crossCountry;
                totalNight += flight.night;
                totalSolo += flight.solo;
                totalActualIFR += flight.actualIFR;
            });
            drawFlights(true);
            const filteredMissingAirports = Array.from(missingAirports).filter(code => !headerKeywords.includes(code.toUpperCase()));
            if (filteredMissingAirports.length > 0) {
                const missingList = filteredMissingAirports.join(', ');
                showFileError(`The following codes were not recognized as airports: ${missingList}`, 'airport-code');
            }
        } else {
            await showColumnMappingUI(rows);
        }
        userDefinedFromIndex = -1;
        userDefinedToIndex = -1;
        pendingLogbookData = null;
    } catch (error) {
        showFileError(`Failed to parse PDF file: ${error.message}. Please ensure the file is a valid PDF.`, 'file-format');
        document.getElementById('progress-bar').style.display = 'none';
        userDefinedFromIndex = -1;
        userDefinedToIndex = -1;
        pendingLogbookData = null;
    }
}

async function parseImage(dataUrl) {
    try {
        const { data: { text } } = await Tesseract.recognize(dataUrl, 'eng');
        const cleanedText = cleanOCRText(text);
        const lines = cleanedText.split('\n').filter(line => line.trim() !== '');
        parseLinesAsText(lines.join('\n'));
    } catch (error) {
        showFileError('Failed to process image. Please ensure it contains clear, readable text and try again.', 'file-format');
    }
}

function parseText(text) {
    const cleanedText = cleanOCRText(text);
    parseLinesAsText(cleanedText);
}

function parseLinesAsText(text) {
    const lines = text.split('\n').filter(line => line.trim() !== '');
    if (lines.length < 1) {
        showFileError('Empty file. Please ensure the file contains flight data.', 'file-format');
        return;
    }

    const newFlights = [];
    let fromIdx = userDefinedFromIndex !== -1 ? userDefinedFromIndex : -1;
    let toIdx = userDefinedToIndex !== -1 ? userDefinedToIndex : -1;
    let dateIdx = -1;

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const words = line.split(/\s+/).map(word => word.toUpperCase());
        if (fromIdx === -1 && toIdx === -1 && words.includes('FROM') && words.includes('TO')) {
            fromIdx = words.indexOf('FROM');
            toIdx = words.indexOf('TO');
            dateIdx = words.indexOf('DATE');
            for (let j = i + 1; j < lines.length; j++) {
                const row = lines[j].split(/\s+/);
                if (row.length < Math.max(fromIdx, toIdx) + 1) continue;

                const from = row[fromIdx].toUpperCase();
                const to = row[toIdx].toUpperCase();
                const date = dateIdx !== -1 && row[dateIdx] ? row[dateIdx] : findDateInRow(row);

                const fromAirport = getAirport(from);
                const toAirport = getAirport(to);

                if (isAirportCode(from) && isAirportCode(to) && fromAirport && toAirport) {
                    const distance = calculateDistance(fromAirport.coords, toAirport.coords);
                    newFlights.push({
                        coords: [fromAirport.coords, toAirport.coords],
                        codes: [fromAirport.iata || fromAirport.identifier, toAirport.iata || toAirport.identifier],
                        date: date,
                        distance: distance,
                        duration: 0,
                        notes: '',
                        aircraftType: 'Unknown',
                        registration: 'Unknown',
                        crossCountry: 0,
                        night: 0,
                        solo: 0,
                        actualIFR: 0,
                        isLocal: fromAirport.iata === toAirport.iata || fromAirport.identifier === toAirport.identifier,
                        depAirport: fromAirport,
                        arrAirport: toAirport
                    });
                    updateAirportFrequency(fromAirport.iata || fromAirport.identifier, toAirport.iata || toAirport.identifier);
                    updateRouteFrequency(fromAirport.iata || fromAirport.identifier, toAirport.iata || toAirport.identifier);
                } else {
                    if (!fromAirport && isAirportCode(from)) missingAirports.add(from);
                    if (!toAirport && isAirportCode(to)) missingAirports.add(to);
                }
            }
            break;
        }
    }

    if (newFlights.length > 0) {
        logbooks[currentLogbook].flights.push(...newFlights);
        drawFlights(true);
        userDefinedFromIndex = -1;
        userDefinedToIndex = -1;
        pendingCSVText = null;
        return;
    }

    lines.forEach((line, index) => {
        const words = line.split(/\s+/);
        const airportCodes = words.filter(isAirportCode);

        if (fromIdx !== -1 && toIdx !== -1 && airportCodes.length >= 2) {
            const row = words;
            if (row.length < Math.max(fromIdx, toIdx) + 1) return;

            const from = row[fromIdx].toUpperCase();
            const to = row[toIdx].toUpperCase();
            const date = dateIdx !== -1 && row[dateIdx] ? row[dateIdx] : findDateInRow(row);

            const fromAirport = getAirport(from);
            const toAirport = getAirport(to);

            if (isAirportCode(from) && isAirportCode(to) && fromAirport && toAirport) {
                const distance = calculateDistance(fromAirport.coords, toAirport.coords);
                newFlights.push({
                    coords: [fromAirport.coords, toAirport.coords],
                    codes: [fromAirport.iata || fromAirport.identifier, toAirport.iata || toAirport.identifier],
                    date: date,
                    distance: distance,
                    duration: 0,
                    notes: '',
                    aircraftType: 'Unknown',
                    registration: 'Unknown',
                    crossCountry: 0,
                    night: 0,
                    solo: 0,
                    actualIFR: 0,
                    isLocal: fromAirport.iata === toAirport.iata || fromAirport.identifier === toAirport.identifier,
                    depAirport: fromAirport,
                    arrAirport: toAirport
                });
                updateAirportFrequency(fromAirport.iata || fromAirport.identifier, toAirport.iata || toAirport.identifier);
                updateRouteFrequency(fromAirport.iata || fromAirport.identifier, toAirport.iata || toAirport.identifier);
            } else {
                if (!fromAirport && isAirportCode(from)) missingAirports.add(from);
                if (!toAirport && isAirportCode(to)) missingAirports.add(to);
            }
        } else if (airportCodes.length >= 2) {
            let departureIndex = -1, arrivalIndex = -1;
            for (let j = 0; j < words.length; j++) {
                if (isAirportCode(words[j])) {
                    if (departureIndex === -1) {
                        departureIndex = j;
                    } else if (arrivalIndex === -1) {
                        arrivalIndex = j;
                        break;
                    }
                }
            }

            if (departureIndex !== -1 && arrivalIndex !== -1) {
                const departure = words[departureIndex].toUpperCase();
                const arrival = words[arrivalIndex].toUpperCase();
                const date = words.find(w => /\d{1,2}\.\w{3}\.\d{2}/.test(w)) || findDateInRow(words);

                const depAirport = getAirport(departure);
                const arrAirport = getAirport(arrival);

                if (depAirport && arrAirport) {
                    const distance = calculateDistance(depAirport.coords, arrAirport.coords);
                    newFlights.push({
                        coords: [depAirport.coords, arrAirport.coords],
                        codes: [depAirport.iata || depAirport.identifier, arrAirport.iata || arrAirport.identifier],
                        date: date,
                        distance: distance,
                        duration: 0,
                        notes: '',
                        aircraftType: 'Unknown',
                        registration: 'Unknown',
                        crossCountry: 0,
                        night: 0,
                        solo: 0,
                        actualIFR: 0,
                        isLocal: depAirport.iata === arrAirport.iata || depAirport.identifier === arrAirport.identifier,
                        depAirport: depAirport,
                        arrAirport: arrAirport
                    });
                    updateAirportFrequency(depAirport.iata || depAirport.identifier, arrAirport.iata || arrAirport.identifier);
                    updateRouteFrequency(depAirport.iata || depAirport.identifier, arrAirport.iata || arrAirport.identifier);
                } else {
                    if (!depAirport && isAirportCode(departure)) missingAirports.add(departure);
                    if (!arrAirport && isAirportCode(arrival)) missingAirports.add(arrival);
                }
            }
        }
    });

    if (newFlights.length > 0) {
        logbooks[currentLogbook].flights.push(...newFlights);
        drawFlights(true);
        userDefinedFromIndex = -1;
        userDefinedToIndex = -1;
        pendingCSVText = null;
    } else {
        showFileError('No valid flight data found in file. Ensure it contains valid airport codes.', 'parsing');
        userDefinedFromIndex = -1;
        userDefinedToIndex = -1;
        pendingCSVText = null;
    }

    if (missingAirports.size > 0) {
        const missingList = Array.from(missingAirports).join(', ');
        showFileError(`The following codes were not recognized as airports in airports.json: ${missingList}. They may be aircraft or simulator types, or missing airports.`, 'airport-code');
    }
}

// Draw flights and markers
function drawFlights(autoOpenPopups = false) {
    map.eachLayer(layer => layer instanceof L.Marker && map.removeLayer(layer));

    const bounds = L.latLngBounds();
    const uniqueAirports = new Set();
    let totalDistance = 0;
    const airportVisits = new Map();

    let displayFlights = logbooks[currentLogbook].flights.filter(flight => {
        if (!showLocalFlights && flight.isLocal) return false;
        if (startDate || endDate) {
            const flightDate = new Date(flight.date);
            const start = startDate ? new Date(startDate) : new Date('1900-01-01');
            const end = endDate ? new Date(endDate) : new Date('9999-12-31');
            return flightDate >= start && flightDate <= end;
        }
        return true;
    });

    displayFlights.forEach(({ depAirport, arrAirport, codes: [fromCode, toCode], date, distance, isLocal }) => {
        if (!depAirport || !arrAirport) return console.warn(`Skipping flight: ${fromCode} to ${toCode}`);

        [fromCode, toCode].forEach((code, idx) => {
            if (!airportVisits.has(code)) {
                const airport = idx === 0 ? depAirport : arrAirport;
                airportVisits.set(code, { coords: airport.coords, dates: [], visits: 0, name: getAirportName(airport) });
            }
            const visits = airportVisits.get(code);
            visits.visits++;
            if (date && date !== 'N/A' && !isNaN(new Date(date).getTime())) visits.dates.push(date);
            uniqueAirports.add(code);
        });

        totalDistance += distance;
        bounds.extend(depAirport.coords).extend(arrAirport.coords);
    });

    airportVisits.forEach((data, code) => {
        const marker = L.marker(data.coords, { title: `${data.name} (${code})` }).addTo(map);
        const datesList = data.dates.length ? `<ul class="max-h-40 overflow-y-auto">${data.dates.map(d => `<li>${d}</li>`).join('')}</ul>` : 'No dates recorded';
        marker.bindPopup(`<b>${data.name} (${code})</b><br>Visits: ${data.visits}<br><b>Dates:</b><br>${datesList}`);
        if (autoOpenPopups) marker.openPopup();
    });

    document.getElementById('total-flights').textContent = `Total Flights: ${displayFlights.length}`;
    document.getElementById('unique-airports').textContent = `Unique Airports: ${uniqueAirports.size}`;
    document.getElementById('total-distance').textContent = `Total Distance: ${totalDistance.toFixed(2)} km`;

    airportVisits.size > 0 ? map.fitBounds(bounds, { padding: [50, 50] }) : map.setView([0, 0], 2);
}