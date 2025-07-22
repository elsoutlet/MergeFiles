// processFiles.js
// Requires SheetJS (xlsx) to be loaded in the page

// Define the expected header row for output
const EXPECTED_HEADER = [
    "Location", "Receipt #", "Inmar Order #", "Item", "UPC", "Department", "Department Name", "Item description",
    "Quantity", "Last Known Price", "Extended Price", "Liquidation %", "Ext Liquidation Prc", "Category Code",
    "Category Code Description", "Vendor Name", "Container ID", "Container", "Parent Container", "Universal Id",
    "Description", "Retail Cost", "Store Cost", "Invoice Cost", "Actual UPC"
];

// Helper: Read a file as ArrayBuffer
function readFileAsync(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Helper: Find header row containing a specific column name
function findHeaderRow(sheet, headerName, maxRows = 20) {
    headerName = headerName.toLowerCase().trim();
    for (let i = 0; i < Math.min(maxRows, sheet.length); i++) {
        if (sheet[i].some(cell => (cell || '').toString().toLowerCase().trim() === headerName)) return i;
    }
    return -1;
}

// Helper: Map header names to lowercased/trimmed for robust matching
function normalizeHeader(header) {
    return header.map(h => (h || '').toString().toLowerCase().trim());
}

// Helper: Extract order ID from first column
function getOrderId(sheet, maxRows = 20) {
    for (let i = 0; i < Math.min(maxRows, sheet.length); i++) {
        if ((sheet[i][0] || '').toString().trim() === 'Order ID:') {
            return sheet[i][2];
        }
    }
    return null;
}

// Helper: Clean dataframe starting from header row
function cleanSheet(sheet, headerRow) {
    const header = sheet[headerRow];
    const normHeader = normalizeHeader(header);
    const data = sheet.slice(headerRow + 1);
    // Remove unnamed columns and empty columns
    const validIndexes = header.map((col, idx) => col && !/^Unnamed/.test(col) ? idx : null).filter(idx => idx !== null);
    const cleanHeader = validIndexes.map(idx => header[idx].toString());
    const cleanNormHeader = validIndexes.map(idx => normHeader[idx]);
    const cleanData = data.map(row => validIndexes.map(idx => row[idx]));
    // Remove columns that are completely empty
    const nonEmptyCols = cleanNormHeader.map((col, idx) => cleanData.some(row => row[idx] !== undefined && row[idx] !== null && row[idx] !== '') ? idx : null).filter(idx => idx !== null);
    const finalHeader = nonEmptyCols.map(idx => cleanHeader[idx]);
    const finalNormHeader = nonEmptyCols.map(idx => cleanNormHeader[idx]);
    const finalData = cleanData.map(row => nonEmptyCols.map(idx => row[idx]));
    // Log headers for debugging
    console.log('Detected header:', finalHeader);
    return [finalHeader, finalNormHeader, ...finalData];
}

// Helper: Format UPC to 12-digit string with checksum
function formatUPC(upc) {
    if (upc == null || upc === '') return '';
    let upcStr = upc.toString().trim();
    let base = upcStr.slice(-11).padStart(11, '0');
    let checksum = calculateChecksum(base);
    return base + checksum;
}
function calculateChecksum(upcStr) {
    upcStr = upcStr.padStart(11, '0');
    let total = 0;
    for (let i = 0; i < 11; i++) {
        total += ((i % 2 === 0) ? 3 : 1) * parseInt(upcStr[i] || '0', 10);
    }
    return (10 - (total % 10)) % 10;
}

// Helper: Convert array-of-arrays to array-of-objects (robust to header variants)
function arrayToObjects(sheet) {
    const [header, normHeader, ...rows] = sheet;
    return rows.map(row => {
        const obj = {};
        normHeader.forEach((h, i) => {
            obj[h] = row[i];
        });
        // Attach original header for output
        obj.__originalHeader = header;
        return obj;
    });
}
// Helper: Convert array-of-objects to array-of-arrays
function objectsToArray(objs) {
    if (!objs.length) return [EXPECTED_HEADER];
    // Map each object to the expected header order
    const rows = objs.map(obj =>
        EXPECTED_HEADER.map(h => {
            // Try to find the value by normalized header
            const norm = h.toLowerCase().trim();
            // Find the key in obj that matches norm
            const key = Object.keys(obj).find(k => k.toLowerCase().trim() === norm);
            return key ? obj[key] : "";
        })
    );
    return [EXPECTED_HEADER, ...rows];
}

// Group by key, aggregate sum for 'quantity', first for others
function groupBy(data, key) {
    key = key.toLowerCase().trim();
    const groups = {};
    data.forEach(row => {
        const k = row[key];
        if (!groups[k]) groups[k] = [];
        groups[k].push(row);
    });
    const result = [];
    for (const k in groups) {
        const rows = groups[k];
        const base = { ...rows[0] };
        if ('quantity' in base) {
            base['quantity'] = rows.reduce((sum, r) => sum + (parseFloat(r['quantity']) || 0), 0);
        }
        result.push(base);
    }
    return result;
}

// Merge logic
async function merge(files) {
    // Read all files as SheetJS arrays
    const sheets = [];
    for (const file of files) {
        const ab = await readFileAsync(file);
        const wb = XLSX.read(ab, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const sheet = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false });
        if (sheet.length) sheets.push(sheet);
    }
    if (sheets.length < 2) return [];
    // Collect header row info and Order IDs
    const altidInfo = {}, itemInfo = {};
    sheets.forEach((sheet, i) => {
        const altidRow = findHeaderRow(sheet, 'Alt Universal Id');
        const itemRow = findHeaderRow(sheet, 'Item');
        console.log('File', i, 'altidRow', altidRow, 'row content:', sheet[altidRow]);
        console.log('File', i, 'itemRow', itemRow, 'row content:', sheet[itemRow]);
        console.log('File', i, 'row above altidRow', altidRow > 0 ? sheet[altidRow - 1] : null);
        console.log('File', i, 'row above itemRow', itemRow > 0 ? sheet[itemRow - 1] : null);
        if (altidRow !== -1) {
            const orderId = getOrderId(sheet);
            altidInfo[i] = { row: altidRow, orderId }; // Use altidRow, not altidRow - 1
        }
        if (itemRow !== -1) {
            itemInfo[i] = itemRow; // Use itemRow, not itemRow - 1
        }
    });
    // Build cleaned data
    const altidData = {}, itemData = {};
    for (const i in altidInfo) {
        altidData[i] = arrayToObjects(cleanSheet(sheets[i], altidInfo[i].row));
    }
    for (const i in itemInfo) {
        itemData[i] = arrayToObjects(cleanSheet(sheets[i], itemInfo[i]));
    }
    // Match item and altid by order ID
    const mergedList = [];
    for (const itemIdx in itemData) {
        const itemArr = itemData[itemIdx];
        let orderVal = itemArr.length && (itemArr[0]['inmar order #'] || itemArr[0]['Inmar Order #']);
        if (!orderVal) continue;
        for (const altidIdx in altidInfo) {
            if (orderVal == altidInfo[altidIdx].orderId && altidData[altidIdx]) {
                let altidArr = altidData[altidIdx].map(row => ({ ...row }));
                let itemArrCopy = itemArr.map(row => ({ ...row }));
                // Grouping
                if (altidArr.length && ('alt universal id' in altidArr[0] || 'Alt Universal Id' in altidArr[0])) {
                    altidArr = groupBy(altidArr, 'alt universal id');
                }
                if (itemArrCopy.length && ('item' in itemArrCopy[0] || 'Item' in itemArrCopy[0])) {
                    itemArrCopy = groupBy(itemArrCopy, 'item');
                }
                // Merge
                const merged = [];
                itemArrCopy.forEach(itemRow => {
                    const match = altidArr.find(a => a['alt universal id'] == itemRow['item']);
                    if (match) {
                        let mergedRow = { ...itemRow, ...match };
                        // Clean up columns
                        if ('alt universal id' in mergedRow) mergedRow['alt universal id'] = mergedRow['item'];
                        // Extended Price
                        if ('quantity' in mergedRow && 'last known price' in mergedRow) {
                            mergedRow['extended price'] = (parseFloat(mergedRow['quantity']) || 0) * (parseFloat(mergedRow['last known price']) || 0);
                        }
                        // Ext Liquidation Prc
                        if ('extended price' in mergedRow && 'liquidation %' in mergedRow) {
                            mergedRow['ext liquidation prc'] = (parseFloat(mergedRow['extended price']) || 0) * (parseFloat(mergedRow['liquidation %']) || 0) * 0.01;
                        }
                        // Actual UPC
                        if ('universal id' in mergedRow) {
                            mergedRow['actual upc'] = formatUPC(mergedRow['universal id']);
                        }
                        merged.push(mergedRow);
                    }
                });
                if (merged.length) {
                    // Convert to CSV string
                    const arr = objectsToArray(merged);
                    const csv = XLSX.utils.sheet_to_csv(XLSX.utils.aoa_to_sheet(arr));
                    mergedList.push([csv, orderVal]);
                }
            }
        }
    }
    if (!mergedList.length) {
        console.warn('No merged data produced. Check header names and order IDs.');
    }
    return mergedList;
}

// Combine all merged files into a single workbook
async function combineMergedFiles(files, masterFile = null) {
    const mergedList = await merge(files);
    const allRows = [];
    let header = null;
    for (const [csv, orderId] of mergedList) {
        // Use XLSX.read to parse CSV string (SheetJS Community Edition)
        const wbCsv = XLSX.read(csv, { type: 'string' });
        const wsCsv = wbCsv.Sheets[wbCsv.SheetNames[0]];
        const arr = XLSX.utils.sheet_to_json(wsCsv, { header: 1 });
        if (!header) header = arr[0];
        allRows.push(...arr.slice(1));
    }
    // If masterFile is provided, prepend its rows
    if (masterFile) {
        const ab = await readFileAsync(masterFile);
        const wb = XLSX.read(ab, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const masterArr = XLSX.utils.sheet_to_json(ws, { header: 1 });
        if (masterArr.length) {
            header = masterArr[0];
            allRows.unshift(...masterArr.slice(1));
        }
    }
    if (!header) {
        console.warn('No header found for combined file.');
        return null;
    }
    const finalArr = [header, ...allRows];
    const ws = XLSX.utils.aoa_to_sheet(finalArr);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'MergedData');
    return wb;
}

// Export
window.processFiles = { merge, combineMergedFiles }; 