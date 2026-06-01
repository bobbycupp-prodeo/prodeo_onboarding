/***************************************
 * addressCleanup.gs
 *
 * Reads AM and PM raw address data from sm_reg
 * Uses Google Maps Geocoding to find exact real-world address AND Geocode
 * Copies AM to PM if the house number and street name start match
 * Writes the corrected addresses to DJ/DK and Geocodes to DL/DM
 * Re-checks addresses if they are flagged with <FIX IN SCHOOLMINT>
 ***************************************/

const ADDRESS_CLEANUP_CONFIG = {
  sheetName: 'sm_reg',
  startRow: 3,
  
  // AM Columns
  colAmStreet: 49,  // AW
  colAmCity: 50,    // AX
  colAmZip: 51,     // AY
  colAmOutput: 120, // DP
  colAmGeocode: 122,// DR

  // PM Columns
  colPmStreet: 53,  // BA
  colPmCity: 54,    // BB
  colPmZip: 55,     // BC
  colPmOutput: 121, // DQ
  colPmGeocode: 123  // DS
};

const ERROR_FLAG = "<FIX IN SCHOOLMINT>";

function cleanAddresses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ADDRESS_CLEANUP_CONFIG.sheetName);
  
  if (!sheet) {
    Logger.log(`Sheet not found: ${ADDRESS_CLEANUP_CONFIG.sheetName}`);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < ADDRESS_CLEANUP_CONFIG.startRow) {
    Logger.log("No data to process.");
    return;
  }

  // Read all data in bulk to minimize SpreadsheetApp calls
  const numRows = lastRow - ADDRESS_CLEANUP_CONFIG.startRow + 1;
  
  const amStreetVals = sheet.getRange(ADDRESS_CLEANUP_CONFIG.startRow, ADDRESS_CLEANUP_CONFIG.colAmStreet, numRows, 1).getValues();
  const amCityVals = sheet.getRange(ADDRESS_CLEANUP_CONFIG.startRow, ADDRESS_CLEANUP_CONFIG.colAmCity, numRows, 1).getValues();
  const amZipVals = sheet.getRange(ADDRESS_CLEANUP_CONFIG.startRow, ADDRESS_CLEANUP_CONFIG.colAmZip, numRows, 1).getValues();
  
  const pmStreetVals = sheet.getRange(ADDRESS_CLEANUP_CONFIG.startRow, ADDRESS_CLEANUP_CONFIG.colPmStreet, numRows, 1).getValues();
  const pmCityVals = sheet.getRange(ADDRESS_CLEANUP_CONFIG.startRow, ADDRESS_CLEANUP_CONFIG.colPmCity, numRows, 1).getValues();
  const pmZipVals = sheet.getRange(ADDRESS_CLEANUP_CONFIG.startRow, ADDRESS_CLEANUP_CONFIG.colPmZip, numRows, 1).getValues();
  
  const amOutputRange = sheet.getRange(ADDRESS_CLEANUP_CONFIG.startRow, ADDRESS_CLEANUP_CONFIG.colAmOutput, numRows, 1);
  const amOutputVals = amOutputRange.getValues();
  const amGeocodeRange = sheet.getRange(ADDRESS_CLEANUP_CONFIG.startRow, ADDRESS_CLEANUP_CONFIG.colAmGeocode, numRows, 1);
  const amGeocodeVals = amGeocodeRange.getValues();
  
  const pmOutputRange = sheet.getRange(ADDRESS_CLEANUP_CONFIG.startRow, ADDRESS_CLEANUP_CONFIG.colPmOutput, numRows, 1);
  const pmOutputVals = pmOutputRange.getValues();
  const pmGeocodeRange = sheet.getRange(ADDRESS_CLEANUP_CONFIG.startRow, ADDRESS_CLEANUP_CONFIG.colPmGeocode, numRows, 1);
  const pmGeocodeVals = pmGeocodeRange.getValues();

  let amUpdatesMade = false;
  let pmUpdatesMade = false;

  for (let i = 0; i < numRows; i++) {
    // ------------------------------------
    // Process AM Address
    // ------------------------------------
    const amStreet = String(amStreetVals[i][0]).trim();
    const amExistingOut = String(amOutputVals[i][0]).trim();
    
    // Process if it is blank OR if it was previously flagged as an error
    if (amStreet && (amExistingOut === "" || amExistingOut === ERROR_FLAG)) {
      const amCity = String(amCityVals[i][0]).trim();
      const amZip = String(amZipVals[i][0]).trim();
      
      const amSearchCity = amCity ? amCity : "Twin Cities Metro";
      const amRawAddress = `${amStreet}, ${amSearchCity}, MN ${amZip}`.trim();
      
      try {
        const amResult = verifyAddressWithMaps_(amRawAddress);
        if (amResult) {
          // Check if we are updating an existing value or error flag to save unnecessary writes
          if (amOutputVals[i][0] !== amResult.address || amGeocodeVals[i][0] !== amResult.geocode) {
            amOutputVals[i][0] = amResult.address;
            amGeocodeVals[i][0] = amResult.geocode;
            amUpdatesMade = true;
          }
          Logger.log(`Row ${i + ADDRESS_CLEANUP_CONFIG.startRow} (AM): ${amResult.address}`);
          Utilities.sleep(1000); 
        }
      } catch (e) {
        Logger.log(`Error on row ${i + ADDRESS_CLEANUP_CONFIG.startRow} (AM): ${e.message}`);
      }
    }

    // ------------------------------------
    // Process PM Address
    // ------------------------------------
    const pmStreet = String(pmStreetVals[i][0]).trim();
    const pmExistingOut = String(pmOutputVals[i][0]).trim();
    
    // Process if it is blank OR if it was previously flagged as an error
    if (pmStreet && (pmExistingOut === "" || pmExistingOut === ERROR_FLAG)) {
      const isSameAsAm = isSubstantiallySame_(amStreet, pmStreet);
      const currentAmOutput = String(amOutputVals[i][0]).trim(); 
      const currentAmGeocode = String(amGeocodeVals[i][0]).trim();
      
      // If addresses match, copy from AM (this includes copying the ERROR_FLAG if AM failed)
      if (isSameAsAm && currentAmOutput) {
        if (pmOutputVals[i][0] !== currentAmOutput || pmGeocodeVals[i][0] !== currentAmGeocode) {
          pmOutputVals[i][0] = currentAmOutput;
          pmGeocodeVals[i][0] = currentAmGeocode;
          pmUpdatesMade = true;
        }
        Logger.log(`Row ${i + ADDRESS_CLEANUP_CONFIG.startRow} (PM): Skipped API, copied from AM -> ${currentAmOutput}`);
      } else {
        const pmCity = String(pmCityVals[i][0]).trim();
        const pmZip = String(pmZipVals[i][0]).trim();
        
        const pmSearchCity = pmCity ? pmCity : "Twin Cities Metro";
        const pmRawAddress = `${pmStreet}, ${pmSearchCity}, MN ${pmZip}`.trim();
        
        try {
          const pmResult = verifyAddressWithMaps_(pmRawAddress);
          if (pmResult) {
            if (pmOutputVals[i][0] !== pmResult.address || pmGeocodeVals[i][0] !== pmResult.geocode) {
              pmOutputVals[i][0] = pmResult.address;
              pmGeocodeVals[i][0] = pmResult.geocode;
              pmUpdatesMade = true;
            }
            Logger.log(`Row ${i + ADDRESS_CLEANUP_CONFIG.startRow} (PM): ${pmResult.address}`);
            Utilities.sleep(1000); 
          }
        } catch (e) {
          Logger.log(`Error on row ${i + ADDRESS_CLEANUP_CONFIG.startRow} (PM): ${e.message}`);
        }
      }
    }
  }

  // Write all new addresses back to the sheet at once
  if (amUpdatesMade) {
    amOutputRange.setValues(amOutputVals);
    amGeocodeRange.setValues(amGeocodeVals);
    Logger.log("AM Addresses & Geocodes written to sheet.");
  }
  if (pmUpdatesMade) {
    pmOutputRange.setValues(pmOutputVals);
    pmGeocodeRange.setValues(pmGeocodeVals);
    Logger.log("PM Addresses & Geocodes written to sheet.");
  }
  
  if (!amUpdatesMade && !pmUpdatesMade) {
    Logger.log("No new addresses needed cleaning.");
  }
}

/**
 * Checks if two raw street addresses are substantially the same
 */
function isSubstantiallySame_(amStr, pmStr) {
  if (!amStr || !pmStr) return false;
  const amParts = String(amStr).trim().toLowerCase().split(/\s+/);
  const pmParts = String(pmStr).trim().toLowerCase().split(/\s+/);
  if (amParts.length < 1 || pmParts.length < 1) return false;
  
  const firstPartMatches = amParts[0] === pmParts[0];
  const amSecondPart = (amParts[1] || "").substring(0, 3);
  const pmSecondPart = (pmParts[1] || "").substring(0, 3);
  const secondPartMatches = amSecondPart === pmSecondPart;
  
  return firstPartMatches && secondPartMatches;
}

/**
 * Helper function to verify and clean addresses using Google Maps Data
 * Returns an object with the formatted address and the Lat/Long geocode.
 */
function verifyAddressWithMaps_(rawAddress) {
  const response = Maps.newGeocoder().geocode(rawAddress);
  
  if (response.status === 'OK' && response.results.length > 0) {
    const result = response.results[0];
    
    if (result.geometry.location_type === 'ROOFTOP' || result.geometry.location_type === 'RANGE_INTERPOLATED') {
      
      let streetNum = "", route = "", city = "", state = "", zip = "";
      
      result.address_components.forEach(comp => {
        if (comp.types.includes("street_number")) streetNum = comp.short_name;
        if (comp.types.includes("route")) route = comp.short_name; 
        if (comp.types.includes("locality") || comp.types.includes("neighborhood")) city = comp.long_name;
        if (comp.types.includes("administrative_area_level_1")) state = comp.short_name;
        if (comp.types.includes("postal_code")) zip = comp.short_name;
      });
      
      if (streetNum && route && city && zip) {
        const formattedAddress = `${streetNum} ${route}, ${city}, ${state} ${zip}`;
        const geocode = `${result.geometry.location.lat}, ${result.geometry.location.lng}`;
        
        return {
          address: formattedAddress,
          geocode: geocode
        };
      }
    }
  }
  
  // If it failed to find an exact match, output the new error flag
  return {
    address: ERROR_FLAG,
    geocode: ""
  };
}

/**
 * A utility function to fetch geocodes for addresses 
 * that were already cleaned prior to the geocode script update.
 */
function catchUpGeocodes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ADDRESS_CLEANUP_CONFIG.sheetName);
  
  if (!sheet) {
    Logger.log(`Sheet not found: ${ADDRESS_CLEANUP_CONFIG.sheetName}`);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < ADDRESS_CLEANUP_CONFIG.startRow) return;

  const numRows = lastRow - ADDRESS_CLEANUP_CONFIG.startRow + 1;
  
  // Notice we reuse the columns from ADDRESS_CLEANUP_CONFIG
  const amAddressVals = sheet.getRange(ADDRESS_CLEANUP_CONFIG.startRow, ADDRESS_CLEANUP_CONFIG.colAmOutput, numRows, 1).getValues();
  const amGeocodeRange = sheet.getRange(ADDRESS_CLEANUP_CONFIG.startRow, ADDRESS_CLEANUP_CONFIG.colAmGeocode, numRows, 1);
  const amGeocodeVals = amGeocodeRange.getValues();
  
  const pmAddressVals = sheet.getRange(ADDRESS_CLEANUP_CONFIG.startRow, ADDRESS_CLEANUP_CONFIG.colPmOutput, numRows, 1);
  const pmGeocodeRange = sheet.getRange(ADDRESS_CLEANUP_CONFIG.startRow, ADDRESS_CLEANUP_CONFIG.colPmGeocode, numRows, 1);
  const pmGeocodeVals = pmGeocodeRange.getValues();

  let amUpdatesMade = false;
  let pmUpdatesMade = false;

  for (let i = 0; i < numRows; i++) {
    const amAddress = String(amAddressVals[i][0]).trim();
    const existingAmGeocode = String(amGeocodeVals[i][0]).trim();
    
    if (amAddress && amAddress !== ERROR_FLAG && existingAmGeocode === "") {
      try {
        const geo = getGeocodeOnly_(amAddress);
        if (geo) {
          amGeocodeVals[i][0] = geo;
          amUpdatesMade = true;
          Logger.log(`Row ${i + ADDRESS_CLEANUP_CONFIG.startRow} (AM Geocode): ${geo}`);
          Utilities.sleep(500); 
        }
      } catch (e) {
        Logger.log(`Error on row ${i + ADDRESS_CLEANUP_CONFIG.startRow} (AM): ${e.message}`);
      }
    }

    const pmAddress = String(pmAddressVals[i][0]).trim();
    const existingPmGeocode = String(pmGeocodeVals[i][0]).trim();

    if (pmAddress && pmAddress !== ERROR_FLAG && existingPmGeocode === "") {
      const currentAmGeo = String(amGeocodeVals[i][0]).trim();
      
      if (pmAddress === amAddress && currentAmGeo !== "") {
        pmGeocodeVals[i][0] = currentAmGeo;
        pmUpdatesMade = true;
        Logger.log(`Row ${i + ADDRESS_CLEANUP_CONFIG.startRow} (PM Geocode): Copied from AM -> ${currentAmGeo}`);
      } else {
        try {
          const geo = getGeocodeOnly_(pmAddress);
          if (geo) {
            pmGeocodeVals[i][0] = geo;
            pmUpdatesMade = true;
            Logger.log(`Row ${i + ADDRESS_CLEANUP_CONFIG.startRow} (PM Geocode): ${geo}`);
            Utilities.sleep(500); 
          }
        } catch (e) {
          Logger.log(`Error on row ${i + ADDRESS_CLEANUP_CONFIG.startRow} (PM): ${e.message}`);
        }
      }
    }
  }

  if (amUpdatesMade) {
    amGeocodeRange.setValues(amGeocodeVals);
    Logger.log("AM Geocodes written to sheet.");
  }
  if (pmUpdatesMade) {
    pmGeocodeRange.setValues(pmGeocodeVals);
    Logger.log("PM Geocodes written to sheet.");
  }
}

/**
 * Helper function purely for fetching Lat/Lng strings
 */
function getGeocodeOnly_(address) {
  const response = Maps.newGeocoder().geocode(address);
  
  if (response.status === 'OK' && response.results.length > 0) {
    const location = response.results[0].geometry.location;
    return `${location.lat}, ${location.lng}`;
  }
  
  return "NOT FOUND";
}