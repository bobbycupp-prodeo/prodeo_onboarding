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

  colStudentId: 2, // B

  // AM Columns
  colAmStreet: 49,   // AW
  colAmCity: 50,     // AX
  colAmZip: 51,      // AY
  colAmOutput: 120,  // DP
  colAmGeocode: 122, // DR

  // PM Columns
  colPmStreet: 53,   // BA
  colPmCity: 54,     // BB
  colPmZip: 55,      // BC
  colPmOutput: 121,  // DQ
  colPmGeocode: 123, // DS

  // Cleaned Address Tracking
  colCleanedAddressStuId: 124 // DT
};

const ERROR_FLAG = "<FIX IN SCHOOLMINT>";

const MAX_DISTANCE_FROM_INPUT_ZIP_MILES = 2;
const ZIP_CENTER_CACHE = {};

function cleanAddresses() {
  cleanAddressesByMode_("all");
}

function cleanAddressesByMode_(mode) {
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

  const studentIdVals = sheet.getRange(
    ADDRESS_CLEANUP_CONFIG.startRow,
    ADDRESS_CLEANUP_CONFIG.colStudentId,
    numRows,
    1
  ).getValues();

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

  const cleanedStuIdRange = sheet.getRange(
    ADDRESS_CLEANUP_CONFIG.startRow,
    ADDRESS_CLEANUP_CONFIG.colCleanedAddressStuId,
    numRows,
    1
  );
  const cleanedStuIdVals = cleanedStuIdRange.getValues();

  let amUpdatesMade = false;
  let pmUpdatesMade = false;
  let cleanedStuIdUpdatesMade = false;

  for (let i = 0; i < numRows; i++) {
    const rowNumber = i + ADDRESS_CLEANUP_CONFIG.startRow;
    const studentId = String(studentIdVals[i][0]).trim();

    // ------------------------------------
    // Process AM Address
    // ------------------------------------
    const amStreet = String(amStreetVals[i][0]).trim();
    const existingAmOutputAtStart = String(amOutputVals[i][0]).trim();

    if (
      mode === "missingOrFlagged" &&
      existingAmOutputAtStart !== "" &&
      existingAmOutputAtStart !== ERROR_FLAG
    ) {
      Logger.log(`Row ${rowNumber} (AM): Skipped, already cleaned.`);
    } else if (amStreet) {
      const amCity = String(amCityVals[i][0]).trim();
      const amZip = String(amZipVals[i][0]).trim();

      const amSearchCity = amCity ? amCity : "Twin Cities Metro";
      const amStreetForSearch = normalizeStreetForGeocoding_(amStreet);
      const amRawAddress = `${amStreetForSearch}, ${amSearchCity}, MN ${amZip}`.trim();

      try {
        const amResult = verifyAddressWithMaps_(amRawAddress, amZip);

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
    const existingPmOutputAtStart = String(pmOutputVals[i][0]).trim();

    if (
      mode === "missingOrFlagged" &&
      existingPmOutputAtStart !== "" &&
      existingPmOutputAtStart !== ERROR_FLAG
    ) {
      Logger.log(`Row ${rowNumber} (PM): Skipped, already cleaned.`);
    } else if (pmStreet) {
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
        const pmStreetForSearch = normalizeStreetForGeocoding_(pmStreet);
        const pmRawAddress = `${pmStreetForSearch}, ${pmSearchCity}, MN ${pmZip}`.trim();

        try {
          const pmResult = verifyAddressWithMaps_(pmRawAddress, pmZip);

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

    // ------------------------------------
    // Track which student owns the cleaned address data
    // ------------------------------------
    const hasCleanedAddressData =
      String(amOutputVals[i][0]).trim() !== "" ||
      String(amGeocodeVals[i][0]).trim() !== "" ||
      String(pmOutputVals[i][0]).trim() !== "" ||
      String(pmGeocodeVals[i][0]).trim() !== "";

    const existingCleanedStuId = String(cleanedStuIdVals[i][0]).trim();

    if (hasCleanedAddressData && studentId && existingCleanedStuId !== studentId) {
      cleanedStuIdVals[i][0] = studentId;
      cleanedStuIdUpdatesMade = true;
      Logger.log(`Row ${rowNumber}: Cleaned Address StuID set to ${studentId}.`);
    }

    if (!hasCleanedAddressData && existingCleanedStuId !== "") {
      cleanedStuIdVals[i][0] = "";
      cleanedStuIdUpdatesMade = true;
      Logger.log(`Row ${rowNumber}: Cleaned Address StuID cleared.`);
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

  if (cleanedStuIdUpdatesMade) {
    cleanedStuIdRange.setValues(cleanedStuIdVals);
    Logger.log("Cleaned Address StuIDs written to sheet.");
  }

  if (!amUpdatesMade && !pmUpdatesMade && !cleanedStuIdUpdatesMade) {
    Logger.log("No address/geocode/student ID changes found.");
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
/**
 * Helper function to verify and clean addresses using Google Maps Data.
 * Returns an object with the formatted address and the Lat/Long geocode.
 *
 * If Google returns a result in a different ZIP area than the submitted ZIP,
 * and the returned address is more than MAX_DISTANCE_FROM_INPUT_ZIP_MILES
 * from the submitted ZIP area, this returns ERROR_FLAG instead.
 */
function verifyAddressWithMaps_(rawAddress, inputZip) {
  const normalizedInputZip = normalizeZip_(inputZip);
  const response = Maps.newGeocoder().geocode(rawAddress);

  if (response.status === 'OK' && response.results.length > 0) {
    const result = response.results[0];

    const allowedLocationTypes = [
      "ROOFTOP",
      "RANGE_INTERPOLATED",
      "GEOMETRIC_CENTER"
      ];

    if (allowedLocationTypes.includes(result.geometry.location_type)) {
      let streetNum = "";
      let route = "";
      let city = "";
      let state = "";
      let zip = "";

      result.address_components.forEach(comp => {
        if (comp.types.includes("street_number")) streetNum = comp.short_name;
        if (comp.types.includes("route")) route = comp.short_name;
        if (comp.types.includes("locality") || comp.types.includes("neighborhood")) city = comp.long_name;
        if (comp.types.includes("administrative_area_level_1")) state = comp.short_name;
        if (comp.types.includes("postal_code")) zip = comp.short_name;
      });

      const normalizedResultZip = normalizeZip_(zip);
      const resultLocation = result.geometry.location;

      if (normalizedInputZip && normalizedResultZip && normalizedInputZip !== normalizedResultZip) {
        const inputZipCenter = getZipCenter_(normalizedInputZip);

        if (!inputZipCenter) {
          Logger.log(
            `ZIP mismatch for "${rawAddress}". Input ZIP ${normalizedInputZip}, result ZIP ${normalizedResultZip}. Could not verify ZIP center, flagging.`
          );

          return {
            address: ERROR_FLAG,
            geocode: ""
          };
        }

        const milesFromInputZip = getDistanceMiles_(
          resultLocation.lat,
          resultLocation.lng,
          inputZipCenter.lat,
          inputZipCenter.lng
        );

        if (milesFromInputZip > MAX_DISTANCE_FROM_INPUT_ZIP_MILES) {
          Logger.log(
            `ZIP/distance mismatch for "${rawAddress}". Input ZIP ${normalizedInputZip}, result ZIP ${normalizedResultZip}, distance ${milesFromInputZip.toFixed(2)} miles. Flagging.`
          );

          return {
            address: ERROR_FLAG,
            geocode: ""
          };
        }

        Logger.log(
          `ZIP mismatch allowed for "${rawAddress}". Input ZIP ${normalizedInputZip}, result ZIP ${normalizedResultZip}, distance ${milesFromInputZip.toFixed(2)} miles.`
        );
      }

      if (streetNum && route && city && zip) {
        const formattedAddress = `${streetNum} ${route}, ${city}, ${state} ${zip}`;
        const geocode = `${resultLocation.lat}, ${resultLocation.lng}`;

        return {
          address: formattedAddress,
          geocode: geocode
        };
      }
    }
  }

  return {
    address: ERROR_FLAG,
    geocode: ""
  };
}

/**
 * Normalizes ZIPs to their first 5 digits.
 */
function normalizeZip_(zip) {
  const match = String(zip || "").match(/\d{5}/);
  return match ? match[0] : "";
}

/**
 * Gets an approximate center point for a submitted ZIP.
 * Cached during the script run to avoid repeated ZIP geocoding calls.
 */
function getZipCenter_(zip) {
  const normalizedZip = normalizeZip_(zip);

  if (!normalizedZip) return null;

  if (ZIP_CENTER_CACHE[normalizedZip]) {
    return ZIP_CENTER_CACHE[normalizedZip];
  }

  const response = Maps.newGeocoder().geocode(`${normalizedZip}, MN`);

  if (response.status === 'OK' && response.results.length > 0) {
    const location = response.results[0].geometry.location;

    ZIP_CENTER_CACHE[normalizedZip] = {
      lat: location.lat,
      lng: location.lng
    };

    return ZIP_CENTER_CACHE[normalizedZip];
  }

  return null;
}

/**
 * Calculates distance between two latitude/longitude points in miles.
 */
function getDistanceMiles_(lat1, lng1, lat2, lng2) {
  const earthRadiusMiles = 3958.8;

  const dLat = degreesToRadians_(lat2 - lat1);
  const dLng = degreesToRadians_(lng2 - lng1);

  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(degreesToRadians_(lat1)) *
      Math.cos(degreesToRadians_(lat2)) *
      Math.sin(dLng / 2) *
      Math.sin(dLng / 2);

  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

  return earthRadiusMiles * c;
}

function degreesToRadians_(degrees) {
  return degrees * Math.PI / 180;
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

function normalizeStreetForGeocoding_(street) {
  return String(street || "")
    .trim()
    .replace(/\bstreet\b/gi, "St")
    .replace(/\bavenue\b/gi, "Ave")
    .replace(/\broad\b/gi, "Rd")
    .replace(/\bboulevard\b/gi, "Blvd")
    .replace(/\bdrive\b/gi, "Dr")
    .replace(/\blane\b/gi, "Ln")
    .replace(/\bplace\b/gi, "Pl")
    .replace(/\bcourt\b/gi, "Ct")
    .replace(/\bnorth\b/gi, "N")
    .replace(/\bsouth\b/gi, "S")
    .replace(/\beast\b/gi, "E")
    .replace(/\bwest\b/gi, "W");
}

function cleanMissingOrFlaggedAddresses() {
  cleanAddressesByMode_("missingOrFlagged");
}