console.log('[script2.js] Script loaded');

const timelineSlider = document.getElementById('timeline-slider');
const playButton = document.getElementById('play-button');
const pauseButton = document.getElementById('pause-button');
const resetButton = document.getElementById('reset-button');
const currentDateLabel = document.getElementById('current-date-label');

let animationInterval = null;
let airplaneMarker = null;
let flightPaths = [];
let currentFlightIndex = 0;
let isPlaying = false;

// Ensure timeline controls remain visible
function ensureTimelineVisible() {
    const timelineControls = document.getElementById('timeline-controls');
    if (timelineControls) {
        timelineControls.classList.remove('hidden');
        timelineControls.style.display = 'flex';
        timelineControls.style.visibility = 'visible';
        timelineControls.style.opacity = '1';
        timelineControls.style.zIndex = '1100';
        console.log('[ensureTimelineVisible] Timeline controls set to visible, style:', timelineControls.style.cssText);
        console.log('[ensureTimelineVisible] DOM state:', timelineControls.outerHTML);
    } else {
        console.error('[ensureTimelineVisible] Timeline controls not found in DOM');
    }
}

// Disable timeline controls
function disableTimelineControls() {
    if (timelineSlider) timelineSlider.disabled = true;
    if (playButton) playButton.disabled = true;
    if (pauseButton) pauseButton.disabled = true;
    if (resetButton) resetButton.disabled = true;
    if (pauseButton) pauseButton.classList.add('hidden');
    if (playButton) playButton.classList.remove('hidden');
    console.log('[disableTimelineControls] Timeline controls disabled');
}

// Enable timeline controls
function enableTimelineControls() {
    if (timelineSlider) timelineSlider.disabled = false;
    if (playButton) playButton.disabled = false;
    if (pauseButton) pauseButton.disabled = false;
    if (resetButton) resetButton.disabled = false;
    console.log('[enableTimelineControls] Timeline controls enabled');
}

// Initialize default slider (disabled, no flights)
function initializeDefaultSlider() {
    if (!timelineSlider || !currentDateLabel) {
        console.error('[initializeDefaultSlider] Slider or label element missing');
        return;
    }
    const defaultMinDate = new Date('2021-01-01').getTime();
    const defaultMaxDate = new Date().getTime();
    timelineSlider.min = defaultMinDate;
    timelineSlider.max = defaultMaxDate;
    timelineSlider.step = 24 * 60 * 60 * 1000; // One day
    timelineSlider.value = defaultMinDate;
    updateDateLabel(defaultMinDate);
    disableTimelineControls();
    ensureTimelineVisible();
    console.log('[initializeDefaultSlider] Initialized default slider range (disabled)');
}

// Initialize timeline with flight data
function initializeTimeline() {
    console.log('[initializeTimeline] Starting initialization');
    ensureTimelineVisible();

    console.log('[initializeTimeline] Logbooks state:', window.logbooks, 'Current logbook:', window.currentLogbook);
    console.log('[initializeTimeline] Map state:', window.map);

    if (!window.logbooks || !window.logbooks[window.currentLogbook] || !window.logbooks[window.currentLogbook].flights) {
        console.error('[initializeTimeline] Invalid logbooks or flights data');
        if (window.showFileError) {
            window.showFileError('Logbook data not loaded. Please upload a valid logbook.', 'general');
        }
        initializeDefaultSlider();
        return;
    }

    const flights = window.logbooks[window.currentLogbook].flights
        .filter(flight => flight.date !== 'N/A' && !isNaN(new Date(flight.date).getTime()));
    console.log(`[initializeTimeline] Found ${flights.length} flights with valid dates`);

    if (flights.length === 0) {
        console.warn('[initializeTimeline] No valid flight dates found');
        if (window.showFileError) {
            window.showFileError('No valid flight dates found for animation. Please ensure your logbook has flights with valid dates.', 'date-parsing');
        }
        initializeDefaultSlider();
        return;
    }

    // Sort flights chronologically
    flights.sort((a, b) => new Date(a.date) - new Date(b.date));

    // Get date range
    const dates = flights.map(flight => new Date(flight.date));
    const minDate = new Date(Math.min(...dates));
    let maxDate = new Date(Math.max(...dates));
    
    if (minDate.getTime() === maxDate.getTime()) {
        maxDate = new Date(minDate.getTime() + 24 * 60 * 60 * 1000); // Add one day
        console.log('[initializeTimeline] Single date detected, extending maxDate by one day');
    }

    const minTimestamp = minDate.getTime();
    const maxTimestamp = maxDate.getTime();
    console.log(`[initializeTimeline] Date range: ${minDate.toISOString()} to ${maxDate.toISOString()}`);

    // Set slider range
    timelineSlider.min = minTimestamp;
    timelineSlider.max = maxTimestamp;
    timelineSlider.step = 24 * 60 * 60 * 1000; // One day
    timelineSlider.value = minTimestamp;

    // Update date label
    updateDateLabel(minTimestamp);

    // Calculate overall bounds for all flights
    const overallBounds = window.L.latLngBounds();
    flights.forEach(flight => {
        if (flight.depAirport) overallBounds.extend(flight.depAirport.coords);
        if (flight.arrAirport) overallBounds.extend(flight.arrAirport.coords);
    });

    // Draw initial state and set initial map view
    if (window.map && typeof window.map.addLayer === 'function') {
        updateMapForDate(minTimestamp);
        if (overallBounds.isValid()) {
            window.map.fitBounds(overallBounds, { padding: [50, 50], maxZoom: window.calculateDynamicZoom ? window.calculateDynamicZoom() : 3 });
            console.log('[initializeTimeline] Initial map view set to overall bounds');
        }
    } else {
        console.warn('[initializeTimeline] Map not ready, enabling slider without map update');
        if (window.showFileError) {
            window.showFileError('Map not initialized. Slider enabled, but flight paths may not display.', 'general');
        }
    }

    // Enable controls
    enableTimelineControls();

    // Remove existing event listeners to prevent duplicates
    timelineSlider.removeEventListener('input', handleSliderInput);
    playButton.removeEventListener('click', startAnimation);
    pauseButton.removeEventListener('click', pauseAnimation);
    resetButton.removeEventListener('click', resetAnimation);

    // Add event listeners
    timelineSlider.addEventListener('input', handleSliderInput);
    playButton.addEventListener('click', startAnimation);
    pauseButton.addEventListener('click', pauseAnimation);
    resetButton.addEventListener('click', resetAnimation);

    console.log('[initializeTimeline] Completed initialization');
}

// Handle slider input
function handleSliderInput() {
    const timestamp = parseInt(timelineSlider.value);
    updateDateLabel(timestamp);
    updateMapForDate(timestamp);
    if (isPlaying) {
        pauseAnimation();
    }
}

function updateDateLabel(timestamp) {
    if (!currentDateLabel) return;
    const date = new Date(timestamp);
    currentDateLabel.textContent = date.toISOString().split('T')[0];
    console.log(`[updateDateLabel] Set date to ${currentDateLabel.textContent}`);
}

function updateMapForDate(timestamp) {
    console.log(`[updateMapForDate] Updating map for timestamp: ${new Date(timestamp).toISOString()}`);
    if (!window.map || !window.L || typeof window.map.addLayer !== 'function') {
        console.warn('[updateMapForDate] Map not initialized, skipping update');
        return;
    }

    try {
        console.time('[updateMapForDate] Rendering');

        // Clear existing paths and airplane
        flightPaths.forEach(path => {
            if (path && typeof path.remove === 'function') {
                path.remove();
            }
        });
        if (airplaneMarker && typeof airplaneMarker.remove === 'function') {
            airplaneMarker.remove();
        }
        flightPaths = [];
        currentFlightIndex = 0;

        // Filter flights up to the current date
        const flights = window.logbooks && window.logbooks[window.currentLogbook] && window.logbooks[window.currentLogbook].flights
            ? window.logbooks[window.currentLogbook].flights
                .filter(flight => {
                    if (flight.date === 'N/A' || isNaN(new Date(flight.date).getTime())) return false;
                    return new Date(flight.date).getTime() <= timestamp;
                })
                .sort((a, b) => new Date(a.date) - new Date(b.date))
            : [];

        console.log(`[updateMapForDate] Rendering ${flights.length} flights`);

        // Draw completed flights
        flights.forEach((flight, index) => {
            if (!flight.depAirport || !flight.arrAirport) return;
            try {
                const path = window.L.polyline([flight.depAirport.coords, flight.arrAirport.coords], {
                    color: '#3388ff',
                    weight: 2
                }).addTo(window.map);
                flightPaths.push(path);
                if (index === flights.length - 1) {
                    currentFlightIndex = index;
                }
            } catch (e) {
                console.error('[updateMapForDate] Error adding polyline:', e);
            }
        });

        // Animate airplane for the latest flight and pan to arrival
        if (flights.length > 0) {
            const latestFlight = flights[flights.length - 1];
            animateAirplane(latestFlight);
            if (latestFlight.arrAirport) {
                window.map.panTo(latestFlight.arrAirport.coords);
                console.log('[updateMapForDate] Panned to arrival airport:', latestFlight.arrAirport.coords);
            }
        }

        console.timeEnd('[updateMapForDate] Rendering');
    } catch (e) {
        console.error('[updateMapForDate] Error updating map:', e);
        console.timeEnd('[updateMapForDate] Rendering');
    }

    ensureTimelineVisible();
}

function animateAirplane(flight) {
    if (!window.map || !window.L || !flight.depAirport || !flight.arrAirport) return;

    const start = flight.depAirport.coords;
    const end = flight.arrAirport.coords;
    const duration = 1500; // 1.5 seconds
    let startTime = null;

    console.log(`[animateAirplane] Animating from ${start} to ${end}`);

    try {
        if (airplaneMarker) airplaneMarker.remove();
        airplaneMarker = window.L.marker(start, {
            icon: window.L.divIcon({
                className: 'airplane-marker',
                html: '<i class="fas fa-plane fa-lg" style="color: #ff4444;"></i>',
                iconSize: [20, 20],
                iconAnchor: [10, 10]
            })
        }).addTo(window.map);

        function animate(timestamp) {
            if (!startTime) startTime = timestamp;
            const progress = (timestamp - startTime) / duration;
            if (progress < 1) {
                const lat = start[0] + (end[0] - start[0]) * progress;
                const lng = start[1] + (end[1] - start[1]) * progress;
                airplaneMarker.setLatLng([lat, lng]);
                requestAnimationFrame(animate);
            } else {
                airplaneMarker.setLatLng(end);
            }
        }

        requestAnimationFrame(animate);
    } catch (e) {
        console.error('[animateAirplane] Error animating airplane:', e);
    }
}

function startAnimation() {
    if (isPlaying || timelineSlider.disabled) return;
    isPlaying = true;
    playButton.classList.add('hidden');
    pauseButton.classList.remove('hidden');

    const step = parseInt(timelineSlider.step);
    const max = parseInt(timelineSlider.max);
    let current = parseInt(timelineSlider.value);

    animationInterval = setInterval(() => {
        current += step;
        if (current > max) {
            current = max;
            pauseAnimation();
        }
        timelineSlider.value = current;
        updateDateLabel(current);
        updateMapForDate(current);
    }, 1000); // 1 second per day

    console.log('[startAnimation] Animation started');
}

function pauseAnimation() {
    if (!isPlaying) return;
    isPlaying = false;
    clearInterval(animationInterval);
    playButton.classList.remove('hidden');
    pauseButton.classList.add('hidden');
    console.log('[pauseAnimation] Animation paused');
}

function resetAnimation() {
    if (timelineSlider.disabled) return;
    pauseAnimation();
    timelineSlider.value = timelineSlider.min;
    updateDateLabel(parseInt(timelineSlider.min));
    updateMapForDate(parseInt(timelineSlider.min));
    console.log('[resetAnimation] Animation reset');
}

// Initialize on load
document.addEventListener('DOMContentLoaded', () => {
    console.log('[script2.js] DOMContentLoaded fired');
    ensureTimelineVisible();
    initializeDefaultSlider();

    document.getElementById('logbook-upload').addEventListener('change', () => {
        console.log('[logbook-upload] Logbook upload detected via input');
    });

    document.addEventListener('logbookLoaded', () => {
        console.log('[logbookLoaded] Logbook loaded event received');
        let attempts = 0;
        const maxAttempts = 50;
        const initCheck = setInterval(() => {
            attempts++;
            console.log('[logbookLoaded] Attempt', attempts, 'Map state:', window.map);
            if (window.map && typeof window.map.addLayer === 'function') {
                console.log('[logbookLoaded] Map ready, initializing timeline');
                initializeTimeline();
                clearInterval(initCheck);
            } else if (attempts >= maxAttempts) {
                console.warn('[logbookLoaded] Map not initialized after', maxAttempts, 'attempts, enabling slider without map');
                initializeTimeline();
                if (window.showFileError) {
                    window.showFileError('Map initialization timed out. Slider enabled, but flight paths may not display.', 'general');
                }
                clearInterval(initCheck);
            } else {
                console.log('[logbookLoaded] Waiting for map to initialize...');
            }
        }, 100);
    });

    window.addEventListener('resize', ensureTimelineVisible);

    const visibilityCheck = setInterval(ensureTimelineVisible, 1000);
    setTimeout(() => clearInterval(visibilityCheck), 10000);

    const mapCheck = setInterval(() => {
        console.log('[mapCheck] Map state:', window.map);
        if (window.map && typeof window.map.on === 'function') {
            console.log('[script2.js] Map ready, attaching event listeners');
            window.map.on('moveend zoomend', ensureTimelineVisible);
            clearInterval(mapCheck);
        } else {
            console.log('[script2.js] Waiting for map initialization...');
        }
    }, 500);
});

