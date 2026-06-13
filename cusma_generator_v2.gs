/**
 * CUSMA Certificate Generator
 * Auto-fills product data and producer information based on SKUs
 */


function populateCUSMA() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cusmaSheet = ss.getSheetByName("[NEW] Manual CUSMA Generator");
  const ui = SpreadsheetApp.getUi();
 
  if (!cusmaSheet) {
    ui.alert("Error: '[NEW] Manual CUSMA Generator' sheet not found!");
    return;
  }
 
  // FIRST: Clear all previous data before populating new data
  cusmaSheet.getRange("C19:G39").clearContent();  // Clear products (not SKUs in column B)
 
  // Get SKUs from column B starting at row 27
  const lastRow = cusmaSheet.getLastRow();
  const startRow = 19;
  const maxRow = 39; // Maximum row for products
 
  if (lastRow < startRow) {
    ui.alert("No SKUs found. Please enter SKUs in column B starting at row 19.");
    return;
  }
 
  const skuRange = cusmaSheet.getRange(`B${startRow}:B${maxRow}`);
  const skuValues = skuRange.getValues().flat().filter(sku => sku !== "");
 
  if (skuValues.length === 0) {
    ui.alert("No SKUs found. Please enter SKUs in column B starting at row 19.");
    return;
  }
 
  // Check if too many SKUs
  if (skuValues.length > 20) {
    ui.alert(`Warning: Found ${skuValues.length} SKUs but only 20 rows available (rows 19-39). Only the first 20 will be processed.`);
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
 
  ui.alert(`Success! Populated ${productData.length} product(s).`);
}


/**
 * Looks up product data from Master Product Inventory
 * Note: The master inventory sheet ID references a private company spreadsheet and is included for documentation purposes only. The script will not run without access to the original data source. 
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


/**
 * Fills the product table (7b, 8, 9, 10)
 * Column B (6a. SKU) is left alone - user enters these manually
 */
function fillProductTable(sheet, products, startRow) {
  products.forEach((product, index) => {
    const row = startRow + index;
   
    // Column C: 6b. Description
    sheet.getRange(row, 3).setValue(product.description || "");
   
    // Column D: 7. Producer
    sheet.getRange(row, 4).setValue(product.supplierName || "");
   
    // Column E: 8. HTS Tariff Classification
    sheet.getRange(row, 5).setValue(product.htsCode || "");
   
    // Column F: 9. Origin Criterion
    sheet.getRange(row, 6).setValue(product.originCriterion || "");
   
    // Column G: 10. Country of Origin
    sheet.getRange(row, 7).setValue(product.country || "");
  });
}
