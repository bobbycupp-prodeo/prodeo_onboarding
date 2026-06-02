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
    const rowNumber = i + ADDRESS_CLEANUP_CONFIG.startRow;

    // ------------------------------------
    // Process AM Address
    // ------------------------------------
    const amStreet = String(amStreetVals[i][0]).trim();

    if (amStreet) {
      const amCity = String(amCityVals[i][0]).trim();
      const amZip = String(amZipVals[i][0]).trim();

      const amSearchCity = amCity ? amCity : "Twin Cities Metro";
      const amRawAddress = `${amStreet}, ${amSearchCity}, MN ${amZip}`.trim();

      try {
        const amResult = verifyAddressWithMaps_(amRawAddress);

        if (amResult) {
          const existingAmOutput = String(amOutputVals[i][0]).trim();
          const existingAmGeocode = String(amGeocodeVals[i][0]).trim();

          if (existingAmOutput !== amResult.address || existingAmGeocode !== amResult.geocode) {
            amOutputVals[i][0] = amResult.address;
            amGeocodeVals[i][0] = amResult.geocode;
            amUpdatesMade = true;

            Logger.log(`Row ${rowNumber} (AM): Updated -> ${amResult.address}`);
          } else {
            Logger.log(`Row ${rowNumber} (AM): No change -> ${amResult.address}`);
          }

          Utilities.sleep(1000);
        }
      } catch (e) {
        Logger.log(`Error on row ${rowNumber} (AM): ${e.message}`);
      }
    } else {
      // If the raw AM street is now blank, clear the cleaned output/geocode.
      if (String(amOutputVals[i][0]).trim() !== "" || String(amGeocodeVals[i][0]).trim() !== "") {
        amOutputVals[i][0] = "";
        amGeocodeVals[i][0] = "";
        amUpdatesMade = true;
        Logger.log(`Row ${rowNumber} (AM): Raw street blank, cleared output/geocode.`);
      }
    }

    // ------------------------------------
    // Process PM Address
    // ------------------------------------
    const pmStreet = String(pmStreetVals[i][0]).trim();

    if (pmStreet) {
      const currentAmOutput = String(amOutputVals[i][0]).trim();
      const currentAmGeocode = String(amGeocodeVals[i][0]).trim();
      const isSameAsAm = isSubstantiallySame_(amStreet, pmStreet);

      // If PM appears to be the same as AM, copy the current AM result.
      if (isSameAsAm && currentAmOutput) {
        const existingPmOutput = String(pmOutputVals[i][0]).trim();
        const existingPmGeocode = String(pmGeocodeVals[i][0]).trim();

        if (existingPmOutput !== currentAmOutput || existingPmGeocode !== currentAmGeocode) {
          pmOutputVals[i][0] = currentAmOutput;
          pmGeocodeVals[i][0] = currentAmGeocode;
          pmUpdatesMade = true;

          Logger.log(`Row ${rowNumber} (PM): Copied from AM -> ${currentAmOutput}`);
        } else {
          Logger.log(`Row ${rowNumber} (PM): No change, already matches AM -> ${currentAmOutput}`);
        }
      } else {
        const pmCity = String(pmCityVals[i][0]).trim();
        const pmZip = String(pmZipVals[i][0]).trim();

        const pmSearchCity = pmCity ? pmCity : "Twin Cities Metro";
        const pmRawAddress = `${pmStreet}, ${pmSearchCity}, MN ${pmZip}`.trim();

        try {
          const pmResult = verifyAddressWithMaps_(pmRawAddress);

          if (pmResult) {
            const existingPmOutput = String(pmOutputVals[i][0]).trim();
            const existingPmGeocode = String(pmGeocodeVals[i][0]).trim();

            if (existingPmOutput !== pmResult.address || existingPmGeocode !== pmResult.geocode) {
              pmOutputVals[i][0] = pmResult.address;
              pmGeocodeVals[i][0] = pmResult.geocode;
              pmUpdatesMade = true;

              Logger.log(`Row ${rowNumber} (PM): Updated -> ${pmResult.address}`);
            } else {
              Logger.log(`Row ${rowNumber} (PM): No change -> ${pmResult.address}`);
            }

            Utilities.sleep(1000);
          }
        } catch (e) {
          Logger.log(`Error on row ${rowNumber} (PM): ${e.message}`);
        }
      }
    } else {
      // If the raw PM street is now blank, clear the cleaned output/geocode.
      if (String(pmOutputVals[i][0]).trim() !== "" || String(pmGeocodeVals[i][0]).trim() !== "") {
        pmOutputVals[i][0] = "";
        pmGeocodeVals[i][0] = "";
        pmUpdatesMade = true;
        Logger.log(`Row ${rowNumber} (PM): Raw street blank, cleared output/geocode.`);
      }
    }
  }

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
    Logger.log("No address/geocode changes found.");
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

function installDailyCleanAddressesTrigger() {
  const functionName = "cleanAddresses";

  // Remove existing daily triggers for this function so duplicates do not pile up.
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Runs daily between 5 AM and 6 AM in the script project's timezone.
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  Logger.log(`Daily trigger installed for ${functionName}.`);
}