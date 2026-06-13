/**
 * USE THIS WHEN THE NEW GENERATOR BREAKS. OLD TEMPLATE IS KEPT
 *
 * CUSMA Certificate Generator
 * Auto-fills product data and producer information based on SKUs
 */


function populateCUSMA_old() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cusmaSheet = ss.getSheetByName("[DEFUNCT] Manual CUSMA Generator");
  const ui = SpreadsheetApp.getUi();
 
  if (!cusmaSheet) {
    ui.alert("Error: '[DEFUNCT] Manual CUSMA Generator' sheet not found!");
    return;
  }
 
  // FIRST: Clear all previous data before populating new data
  cusmaSheet.getRange("C19:O25").clearContent();  // Clear producers
  cusmaSheet.getRange("C27:G34").clearContent();  // Clear products (not SKUs in column B)
 
  // Get SKUs from column B starting at row 27
  const lastRow = cusmaSheet.getLastRow();
  const startRow = 27;
  const maxRow = 34; // Maximum row for products
 
  if (lastRow < startRow) {
    ui.alert("No SKUs found. Please enter SKUs in column B starting at row 27.");
    return;
  }
 
  const skuRange = cusmaSheet.getRange(`B${startRow}:B${maxRow}`);
  const skuValues = skuRange.getValues().flat().filter(sku => sku !== "");
 
  if (skuValues.length === 0) {
    ui.alert("No SKUs found. Please enter SKUs in column B starting at row 27.");
    return;
  }
 
  // Check if too many SKUs
  if (skuValues.length > 8) {
    ui.alert(`Warning: Found ${skuValues.length} SKUs but only 8 rows available (rows 27-34). Only the first 8 will be processed.`);
  }
 
  // Look up product data
  const productData = lookupProductData(skuValues);
 
  if (productData.length === 0) {
    ui.alert("Could not find product data. Please check your SKUs.");
    return;
  }
 
  // Check for NOT FOUND products
  const notFoundSkus = productData.filter(p => p.description === "NOT FOUND" || p.description === "INCOMPLETE DATA").map(p => p.sku);
  if (notFoundSkus.length > 0) {
    const proceed = ui.alert(
      'SKUs Not Found',
      `The following SKUs were not found in the inventory:\n${notFoundSkus.join(', ')}\n\nContinue anyway?`,
      ui.ButtonSet.YES_NO
    );
    if (proceed !== ui.Button.YES) {
      return;
    }
  }
 
  // Fill product table
  fillProductTable(cusmaSheet, productData, startRow);
 
  // Get unique manufacturers and fill Section 5
  const manufacturers = getUniqueManufacturers(productData);
  fillProducerSection(cusmaSheet, manufacturers);
 
  ui.alert(`Success! Populated ${productData.length} product(s) and ${manufacturers.length} producer(s).`);
}


/**
 * Looks up product data from Master Product Inventory
 * Sheet ID: 1mkf1z2MHsHUSJsw8eVQmRSNRXGFxsGnDeBiOMkc10ak, gid 318026309
 */
function lookupProductData(skus) {
  const masterSheet = SpreadsheetApp.openById("1mkf1z2MHsHUSJsw8eVQmRSNRXGFxsGnDeBiOMkc10ak").getSheetById(318026309);
 
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert("Error: Master Product Inventory sheet not found!");
    return [];
  }
 
  const lastRow = masterSheet.getLastRow();
 
  // Get columns from Master Product Inventory
  const skuCol = masterSheet.getRange(`D2:D${lastRow}`).getValues().flat();         // Column D: Short SKU
  const supplierCol = masterSheet.getRange(`M2:M${lastRow}`).getValues().flat();    // Column M: Supplier Name
  const countryCol = masterSheet.getRange(`P2:P${lastRow}`).getValues().flat();     // Column P: Country of Origin
  const htsCol = masterSheet.getRange(`R2:R${lastRow}`).getValues().flat();         // Column R: HTS Code
  const descCol = masterSheet.getRange(`S2:S${lastRow}`).getValues().flat();        // Column S: Goods Description
  const originCritCol = masterSheet.getRange(`V2:V${lastRow}`).getValues().flat();  // Column V: Origin Criterion
 
  const products = [];
 
  skus.forEach(sku => {
    // Trim whitespace from SKU for matching
    const cleanSku = String(sku).trim();
    const index = skuCol.findIndex(v => String(v).trim() === cleanSku);
   
    if (index !== -1) {
      // Found the SKU, but check if data is valid
      const desc = descCol[index];
      const hts = htsCol[index];
      const country = countryCol[index];
      const supplier = supplierCol[index];
     
      // Check for #N/A errors or empty critical fields
      const isInvalid =
        desc === "#N/A" || desc === "" || desc === null || desc === undefined ||
        hts === "#N/A" || hts === "" || hts === null || hts === undefined ||
        country === "" || country === null || country === undefined ||
        supplier === "" || supplier === null || supplier === undefined;
     
      if (isInvalid) {
        Logger.log(`SKU found but has invalid/missing data: "${cleanSku}"`);
        products.push({
          sku: cleanSku,
          description: "INCOMPLETE DATA",
          htsCode: "MISSING",
          originCriterion: "",
          country: "MISSING",
          supplierName: ""
        });
      } else {
        products.push({
          sku: cleanSku,
          description: desc,
          htsCode: hts,
          originCriterion: originCritCol[index] || "",
          country: country,
          supplierName: supplier
        });
      }
    } else {
      // Debug: Log what we're searching for
      Logger.log(`SKU not found: "${cleanSku}"`);
      products.push({
        sku: cleanSku,
        description: "NOT FOUND",
        htsCode: "",
        originCriterion: "",
        country: "",
        supplierName: ""
      });
    }
  });
 
  return products;
}


function getUniqueManufacturers(productData) {
  const mfgSheet = SpreadsheetApp.openById("1mkf1z2MHsHUSJsw8eVQmRSNRXGFxsGnDeBiOMkc10ak").getSheetById(1865583938);
 
  if (!mfgSheet) {
    SpreadsheetApp.getUi().alert("Error: Manufacturer database not found!");
    return [];
  }
 
  const lastRow = mfgSheet.getLastRow();
 
  // Get manufacturer data columns
  const mfgIdCol = mfgSheet.getRange(`A2:A${lastRow}`).getValues().flat();      // Column L: Manufacturer ID (also used as Name)
  const mfgStreetCol = mfgSheet.getRange(`D2:D${lastRow}`).getValues().flat();  // Column N: Street
  const mfgCityCol = mfgSheet.getRange(`E2:E${lastRow}`).getValues().flat();    // Column O: City
  const mfgPostalCol = mfgSheet.getRange(`G2:G${lastRow}`).getValues().flat();  // Column P: Postal Code
  const mfgProvinceCol = mfgSheet.getRange(`F2:F${lastRow}`).getValues().flat(); // Column Q: Province
  const mfgPhoneCol = mfgSheet.getRange(`I2:I${lastRow}`).getValues().flat();   // Column R: Phone
  const mfgEmailCol = mfgSheet.getRange(`J2:J${lastRow}`).getValues().flat();   // Column S: Email
 
  // Get unique supplier names from products (excluding empty and NOT FOUND)
  const uniqueSuppliers = [...new Set(
    productData
      .map(p => p.supplierName)
      .filter(name => name !== "" && name !== undefined)
  )];
 
  const manufacturers = [];
 
  // For each unique supplier, find their info in the manufacturer sheet
  uniqueSuppliers.forEach(supplierName => {
    // Find first occurrence of this manufacturer ID
    const index = mfgIdCol.indexOf(supplierName);
   
    if (index !== -1) {
      manufacturers.push({
        id: supplierName,
        name: supplierName, // Using manufacturer ID as name per requirements
        street: mfgStreetCol[index] || "",
        city: mfgCityCol[index] || "",
        postal: mfgPostalCol[index] || "",
        province: mfgProvinceCol[index] || "",
        phone: mfgPhoneCol[index] || "",
        email: mfgEmailCol[index] || ""
      });
    }
  });
 
  return manufacturers;
}


/**
 * Fills the product table (7b, 8, 9, 10)
 * Column B (7a. SKU) is left alone - user enters these manually
 */
function fillProductTable(sheet, products, startRow) {
  products.forEach((product, index) => {
    const row = startRow + index;
   
    // Column C: 7b. Description
    sheet.getRange(row, 3).setValue(product.description || "");
   
    // Column E: 8. HTS Tariff Classification
    sheet.getRange(row, 5).setValue(product.htsCode || "");
   
    // Column F: 9. Origin Criterion
    sheet.getRange(row, 6).setValue(product.originCriterion || "");
   
    // Column G: 10. Country of Origin
    sheet.getRange(row, 7).setValue(product.country || "");
  });
}


/**
 * Fills Section 5 - Producers' Name and Address
 * Each manufacturer in a separate column (C, D, E, etc.)
 */
function fillProducerSection(sheet, manufacturers) {
  // Clear existing producer data
  sheet.getRange("C19:O25").clearContent();
 
  if (manufacturers.length === 0) {
    return;
  }
 
  // Fill each manufacturer in a separate column
  manufacturers.forEach((mfg, index) => {
    const col = 3 + index; // Column C = 3, D = 4, E = 5, etc.
   
    sheet.getRange(19, col).setValue(mfg.name || "");        // Row 19: Name
    sheet.getRange(20, col).setValue(mfg.street || "");      // Row 20: Street
    sheet.getRange(21, col).setValue(mfg.city || "");        // Row 21: City
    sheet.getRange(22, col).setValue(mfg.postal || "");      // Row 22: Postal Code
    sheet.getRange(23, col).setValue(mfg.province || "");    // Row 23: Province Code
    sheet.getRange(24, col).setValue(mfg.phone || "");       // Row 24: Telephone
    sheet.getRange(25, col).setValue(mfg.email || "");       // Row 25: E-Mail Address
  });
}


/**
 * Create menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('CUSMA Tools')
    .addItem('Populate CUSMA', 'populateCUSMA_old')
    .addItem('Clear SKUs', 'clearSKUs')
    .addSeparator()
    .addToUi();
}
