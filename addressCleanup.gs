/***************************************
 * addressCleanup.gs
 *
 * Reads AM and PM raw address data from sm_reg
 * Skips processing if the street is blank or the destination is already filled
 * Uses Gemini to standardize the addresses based on strict Minneapolis/Saint Paul rules
 * Copies AM to PM if the house number and street name start match
 * Writes the corrected addresses to DJ and DK
 ***************************************/

const ADDRESS_CLEANUP_CONFIG = {
  sheetName: 'sm_reg',
  startRow: 2,
  
  // AM Columns
  colAmStreet: 49, // AW
  colAmCity: 50,   // AX
  colAmZip: 52,    // AZ
  colAmOutput: 114,// DJ

  // PM Columns
  colPmStreet: 53, // BA
  colPmCity: 54,   // BB
  colPmZip: 55,    // BC
  colPmOutput: 115,// DK
  
  modelId: 'gemini-3.5-flash'
};

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
  
  const pmOutputRange = sheet.getRange(ADDRESS_CLEANUP_CONFIG.startRow, ADDRESS_CLEANUP_CONFIG.colPmOutput, numRows, 1);
  const pmOutputVals = pmOutputRange.getValues();

  let amUpdatesMade = false;
  let pmUpdatesMade = false;

  for (let i = 0; i < numRows; i++) {
    // Establish AM Variables
    const amStreet = String(amStreetVals[i][0]).trim();
    const amCity = String(amCityVals[i][0]).trim();
    const amZip = String(amZipVals[i][0]).trim();
    const amExistingOut = String(amOutputVals[i][0]).trim();
    // If city is blank, default to "Twin Cities Metro" to anchor the Google Maps search
    const amSearchCity = amCity ? amCity : "Twin Cities Metro";
    const amRawAddress = `${amStreet}, ${amSearchCity}, MN ${amZip}`.trim();
    
    // --- Process AM Address ---
    if (amStreet && amExistingOut === "") {
      try {
        const amCleaned = verifyAddressWithMaps_(amRawAddress);
        if (amCleaned) {
          amOutputVals[i][0] = amCleaned;
          amUpdatesMade = true;
          Logger.log(`Row ${i + ADDRESS_CLEANUP_CONFIG.startRow} (AM): ${amCleaned}`);
          Utilities.sleep(1000); // Respect rate limits
        }
      } catch (e) {
        Logger.log(`Error on row ${i + ADDRESS_CLEANUP_CONFIG.startRow} (AM): ${e.message}`);
      }
    }

    // Establish PM Variables
    const pmStreet = String(pmStreetVals[i][0]).trim();
    const pmCity = String(pmCityVals[i][0]).trim();
    const pmZip = String(pmZipVals[i][0]).trim();
    const pmExistingOut = String(pmOutputVals[i][0]).trim();
    // If city is blank, default to "Twin Cities Metro" to anchor the Google Maps search
    const pmSearchCity = pmCity ? pmCity : "Twin Cities Metro";
    const pmRawAddress = `${pmStreet}, ${pmSearchCity}, MN ${pmZip}`.trim();
    
    // --- Process PM Address ---
    if (pmStreet && pmExistingOut === "") {
      
      // Optimization: Check if PM address shares the same house number and street name
      const isSameAsAm = isSubstantiallySame_(amStreet, pmStreet);
      const currentAmOutput = String(amOutputVals[i][0]).trim(); // Gets the pre-existing OR newly cleaned AM address
      
      if (isSameAsAm && currentAmOutput) {
        // Just copy the result over
        pmOutputVals[i][0] = currentAmOutput;
        pmUpdatesMade = true;
        Logger.log(`Row ${i + ADDRESS_CLEANUP_CONFIG.startRow} (PM): Skipped API, copied from AM -> ${currentAmOutput}`);
      } else {
        // Call the API if they are different
        try {
          const pmCleaned = verifyAddressWithMaps_(pmRawAddress);
          if (pmCleaned) {
            pmOutputVals[i][0] = pmCleaned;
            pmUpdatesMade = true;
            Logger.log(`Row ${i + ADDRESS_CLEANUP_CONFIG.startRow} (PM): ${pmCleaned}`);
            Utilities.sleep(1000); // Respect rate limits
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
    Logger.log("AM Addresses written to sheet.");
  }
  if (pmUpdatesMade) {
    pmOutputRange.setValues(pmOutputVals);
    Logger.log("PM Addresses written to sheet.");
  }
  
  if (!amUpdatesMade && !pmUpdatesMade) {
    Logger.log("No new addresses needed cleaning.");
  }
}

/**
 * Checks if two raw street addresses are substantially the same
 * by comparing the house number and the first 3 letters of the street name.
 */
function isSubstantiallySame_(amStr, pmStr) {
  if (!amStr || !pmStr) return false;
  
  // Split by any whitespace
  const amParts = String(amStr).trim().toLowerCase().split(/\s+/);
  const pmParts = String(pmStr).trim().toLowerCase().split(/\s+/);
  
  if (amParts.length < 1 || pmParts.length < 1) return false;
  
  // Compare the first part (e.g., "1234")
  const firstPartMatches = amParts[0] === pmParts[0];
  
  // Compare the first 3 characters of the second part (e.g., "mai" for "Main")
  const amSecondPart = (amParts[1] || "").substring(0, 3);
  const pmSecondPart = (pmParts[1] || "").substring(0, 3);
  const secondPartMatches = amSecondPart === pmSecondPart;
  
  return firstPartMatches && secondPartMatches;
}
/**
 * Helper function to verify and clean addresses using Google Maps Data
 */
function verifyAddressWithMaps_(rawAddress) {
  // Use Google's built-in geocoder to find the real address
  const response = Maps.newGeocoder().geocode(rawAddress);
  
  if (response.status === 'OK' && response.results.length > 0) {
    const result = response.results[0];
    
    // Check if Google found a precise building (ROOFTOP or RANGE_INTERPOLATED)
    // If it's APPROXIMATE, it means it only found the street or city (no house number)
    if (result.geometry.location_type === 'ROOFTOP' || result.geometry.location_type === 'RANGE_INTERPOLATED') {
      
      let streetNum = "", route = "", city = "", state = "", zip = "";
      
      // Extract exactly what we want, completely ignoring unit/apartment numbers
      result.address_components.forEach(comp => {
        if (comp.types.includes("street_number")) streetNum = comp.short_name;
        if (comp.types.includes("route")) route = comp.short_name; // e.g., "Main St"
        if (comp.types.includes("locality") || comp.types.includes("neighborhood")) city = comp.long_name;
        if (comp.types.includes("administrative_area_level_1")) state = comp.short_name;
        if (comp.types.includes("postal_code")) zip = comp.short_name;
      });
      
      // Ensure we have the minimum required pieces of a valid address
      if (streetNum && route && city && zip) {
        // Formats perfectly to Title Case, extracting the real city and zip!
        return `${streetNum} ${route}, ${city}, ${state} ${zip}`;
      }
    }
  }
  
  // If it failed to find an exact match (e.g. missing house number or bad spelling), flag it
  return "<NEED VERIFICATION>";
}
// /**
//  * Helper function to call the Gemini API
//  */
// function callGeminiToCleanAddress_(rawAddress) {
//   const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
//   if (!apiKey) {
//     throw new Error("GEMINI_API_KEY not found in Script Properties. Please add it via Project Settings.");
//   }

//   const url = `https://generativelanguage.googleapis.com/v1beta/models/${ADDRESS_CLEANUP_CONFIG.modelId}:generateContent?key=${apiKey}`;
  
//  const prompt = `You are a strict data standardization assistant. 
//   Take the following raw address and format it into a clean, standardized format on a single line. 
//   All addresses are in the Minneapolis/Saint Paul metro area, so addresses without a clear city or zip should be determined with that in mind.
//   Ignore any unit numbers or apartment numbers. We only want street addresses.
//   Fix any obvious spelling errors in the city name. 
//   CRITICAL: Use Title Case for the street and city (e.g., "123 Main St, Minneapolis, MN 55401"). Do not use ALL CAPS.
//   DO NOT output any conversational text, explanations, or markdown formatting. ONLY return the standardized address string.
  
//   Raw Address: ${rawAddress}`;
  
//   const payload = {
//     "contents": [{
//       "parts": [{"text": prompt}]
//     }],
//     "generationConfig": {
//       "temperature": 0.1 
//     }
//   };
  
//   const options = {
//     "method": "post",
//     "contentType": "application/json",
//     "payload": JSON.stringify(payload),
//     "muteHttpExceptions": true
//   };
  
//   const response = UrlFetchApp.fetch(url, options);
//   const json = JSON.parse(response.getContentText());
  
//   if (json.error) {
//      throw new Error(json.error.message);
//   }
  
//   if (json.candidates && json.candidates.length > 0) {
//     return json.candidates[0].content.parts[0].text.trim().replace(/\n/g, ""); 
//   }
  
//   return null;
// }