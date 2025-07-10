const fs = require('fs');
const csv = require('csv-parser');

const airports = [];
const code_to_airport_id = {};
const duplicates = {}; // Track duplicate codes
let skippedByType = 0; // Track rows skipped due to type

function isValidCode(code, type) {
  if (!code) return false;
  if (type === 'iata') return /^[A-Z]{3}$/.test(code);
  return /^[A-Z0-9]{3,4}$/.test(code); // ICAO, gps_code, local_code
}

function isValidCoords(lat, lon) {
  return !isNaN(lat) && !isNaN(lon) && lat >= -90 && lat <= 90 && lon >= -180 && lon <= 180;
}

fs.createReadStream('airports.csv')
  .pipe(csv())
  .on('data', (row) => {
    // Filter for airplane-relevant airport types
    const validTypes = ['small_airport', 'medium_airport', 'large_airport'];
    if (!validTypes.includes(row.type)) {
      console.log(`Skipping row ${row.id} (${row.ident}): Invalid type (${row.type})`);
      skippedByType++;
      return;
    }

    // Validate required fields
    if (!isValidCoords(row.latitude_deg, row.longitude_deg) ||
        (!row.gps_code && !row.icao_code && !row.local_code && !row.iata_code)) {
      console.log(`Skipping row ${row.id} (${row.ident}): Invalid coords or missing codes`);
      return;
    }

    const gps_code = row.gps_code && isValidCode(row.gps_code, 'other') ? row.gps_code.toUpperCase() : null;
    const icao = row.icao_code && isValidCode(row.icao_code, 'other') ? row.icao_code.toUpperCase() : null;
    const local_code = row.local_code && isValidCode(row.local_code, 'other') ? row.local_code.toUpperCase() : null;
    const iata = row.iata_code && isValidCode(row.iata_code, 'iata') ? row.iata_code.toUpperCase() : null;
    const identifier = iata || icao || gps_code || local_code || null;
    const country = row.iso_country && /^[A-Z]{2}$/.test(row.iso_country) ? row.iso_country.toUpperCase() : 'Unknown';
    const coords = [parseFloat(row.latitude_deg), parseFloat(row.longitude_deg)];
    const name = row.name || 'Unknown Airport';

    const airportId = airports.length;
    const airport = {
      id: airportId,
      iata,
      icao,
      identifier,
      country,
      coords,
      name
    };
    airports.push(airport);

    // Map codes and track duplicates
    const mapCode = (code, id) => {
      if (code) {
        if (code_to_airport_id[code]) {
          // Log duplicate
          if (!duplicates[code]) duplicates[code] = [code_to_airport_id[code]];
          duplicates[code].push(id);
          code_to_airport_id[`${code}_${id}`] = id; // Unique key
        } else {
          code_to_airport_id[code] = id;
        }
      }
    };

    mapCode(identifier, airportId);
    if (gps_code && gps_code !== identifier) mapCode(gps_code, airportId);
    if (icao && icao !== identifier) mapCode(icao, airportId);
    if (local_code && local_code !== identifier) mapCode(local_code, airportId);
    if (iata && iata !== identifier) mapCode(iata, airportId);
  })
  .on('end', () => {
    const output = {
      airports,
      code_to_airport_id,
      duplicates // Include duplicates for runtime resolution
    };
    fs.writeFileSync('airports.json', JSON.stringify(output, null, 2));
    console.log(`airports.json generated: ${airports.length} airports, ${skippedByType} rows skipped due to type, ${Object.keys(duplicates).length} duplicate codes`);
  });