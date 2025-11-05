import * as XLSX from 'xlsx';
import Papa from 'papaparse';

// Dynamic imports for pdfjs-dist and mammoth to avoid Vite build issues
let pdfjsLib = null;
let mammothLib = null;

const loadPdfJs = async () => {
  if (!pdfjsLib) {
    // Check if already loaded
    if (window.pdfjsLib || window.pdfjs || (window.pdfjsDist && window.pdfjsDist.getDocument)) {
      pdfjsLib = window.pdfjsLib || window.pdfjs || window.pdfjsDist;
      if (pdfjsLib && pdfjsLib.GlobalWorkerOptions) {
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
      }
      return pdfjsLib;
    }

    // Check if script is already loading
    const existingScript = document.querySelector('script[data-pdfjs]');
    if (existingScript) {
      return new Promise((resolve) => {
        existingScript.addEventListener('load', () => {
          pdfjsLib = window.pdfjsLib || window.pdfjs || window.pdfjsDist;
          if (pdfjsLib && pdfjsLib.GlobalWorkerOptions) {
            pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
          }
          resolve(pdfjsLib);
        });
      });
    }

    // Load PDF.js from CDN using script tag - most reliable method
    return new Promise((resolve, reject) => {
      const script = document.createElement('script');
      // Use unpkg UMD build which properly exposes pdfjsLib
      script.src = 'https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.min.js';
      script.async = true;
      script.crossOrigin = 'anonymous';
      script.setAttribute('data-pdfjs', 'true');
      
      script.onload = () => {
        // Wait a bit for the library to fully initialize
        setTimeout(() => {
          // PDF.js UMD build exposes as pdfjsLib
          pdfjsLib = window.pdfjsLib || window.pdfjs || window.pdfjsDist;
          
          // Try accessing the default export if it's a module
          if (!pdfjsLib && window.pdfjsLib && typeof window.pdfjsLib === 'object') {
            pdfjsLib = window.pdfjsLib.default || window.pdfjsLib;
          }

          // Last resort: search window for PDF.js
          if (!pdfjsLib || !pdfjsLib.getDocument) {
            for (const key in window) {
              try {
                const obj = window[key];
                if (obj && typeof obj === 'object' && obj.getDocument && typeof obj.getDocument === 'function') {
                  pdfjsLib = obj;
                  break;
                }
              } catch (e) {
                // Skip
              }
            }
          }

          if (!pdfjsLib || !pdfjsLib.getDocument) {
            // Try one more time - sometimes it needs a moment
            setTimeout(() => {
              pdfjsLib = window.pdfjsLib || window.pdfjs;
              if (!pdfjsLib || !pdfjsLib.getDocument) {
                reject(new Error('PDF.js loaded but getDocument method not found. Please refresh the page or use CSV/XLSX format.'));
                return;
              }
              configureWorker();
              resolve(pdfjsLib);
            }, 100);
            return;
          }

          configureWorker();
          resolve(pdfjsLib);
        }, 50);
      };
      
      const configureWorker = () => {
        if (pdfjsLib && pdfjsLib.GlobalWorkerOptions) {
          pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
        }
      };
      
      script.onerror = () => {
        reject(new Error('Failed to load PDF.js from CDN. Please check your internet connection or use CSV/XLSX format.'));
      };
      
      document.head.appendChild(script);
    });
  }
  return Promise.resolve(pdfjsLib);
};

const loadMammoth = async () => {
  if (!mammothLib) {
    try {
      // Use Function constructor to avoid Vite static analysis
      const importFn = new Function('specifier', 'return import(specifier)');
      const mammothModule = await importFn('mammoth');
      mammothLib = mammothModule.default || mammothModule;
    } catch (error) {
      console.error('Error loading mammoth:', error);
      throw new Error('DOCX parsing library failed to load. Please refresh the page or use CSV/XLSX format.');
    }
  }
  return mammothLib;
};

// Convert Excel serial date number to readable date string
const convertExcelDate = (value) => {
  // If it's already a string (formatted date), return as is
  if (typeof value === 'string') {
    // Check if it already looks like a date string (contains date patterns)
    if (value.match(/\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}/) || value.match(/\d{1,2}\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/i)) {
      return value;
    }
    // If it's a string but doesn't look like a date, try to parse as number
    const numValue = parseFloat(value);
    if (!isNaN(numValue)) {
      value = numValue;
    } else {
      return value; // Return original string if not a number
    }
  }
  
  // Check if it's a number that looks like an Excel serial date (typically 1-100000 range for modern dates)
  // Excel dates can be like 45727 (days since Jan 1, 1900) or 45727.00011574074 (with time)
  if (typeof value === 'number' && value > 1 && value < 1000000) {
    try {
      // Excel epoch: January 1, 1900 = serial date 1
      // Excel incorrectly treats 1900 as a leap year, so we need to account for that
      // The correct calculation: (serialDate - 1) days from Dec 30, 1899
      
      // Excel serial date calculation
      // Excel date 1 = Jan 1, 1900
      // JavaScript Date(1900, 0, 1) = Jan 1, 1900
      // But Excel incorrectly counts Feb 29, 1900, so we adjust
      
      const excelEpoch = new Date(1899, 11, 30); // Dec 30, 1899 (one day before Excel's epoch)
      
      // Get the integer part (days) and fractional part (time of day)
      const days = Math.floor(value);
      const fraction = value - days;
      
      // Calculate the date: Dec 30, 1899 + (days - 1) days
      // Subtract 1 because Excel date 1 = Jan 1, 1900 = Dec 30, 1899 + 1 day
      const jsDate = new Date(excelEpoch.getTime() + (days - 1) * 24 * 60 * 60 * 1000);
      
      // If there's a fractional part (time), add it (but we'll only use date for display)
      if (fraction > 0) {
        const milliseconds = Math.round(fraction * 24 * 60 * 60 * 1000);
        jsDate.setTime(jsDate.getTime() + milliseconds);
      }
      
      // Format as DD-MMM-YY (e.g., "03-Nov-25")
      const day = String(jsDate.getDate()).padStart(2, '0');
      const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      const month = monthNames[jsDate.getMonth()];
      const year = jsDate.getFullYear();
      const shortYear = String(year).slice(-2);
      
      return `${day}-${month}-${shortYear}`;
    } catch (error) {
      // If conversion fails, return original value as string
      return String(value);
    }
  }
  
  // If it's already a string or not a date-like number, return as is
  return value;
};

export const parseFile = (file, billingType) => {
  return new Promise((resolve, reject) => {
    const fileExtension = file.name.split('.').pop().toLowerCase();
    
    if (fileExtension === 'csv') {
      parseCSV(file, billingType, resolve, reject);
    } else if (fileExtension === 'xlsx' || fileExtension === 'xls') {
      parseXLSX(file, billingType, resolve, reject);
    } else if (fileExtension === 'pdf') {
      parsePDF(file, billingType).then(resolve).catch(reject);
    } else if (fileExtension === 'doc' || fileExtension === 'docx') {
      parseDOC(file, billingType, fileExtension).then(resolve).catch(reject);
    } else {
      reject(new Error('Unsupported file format. Please upload CSV, XLSX, PDF, or DOC/DOCX files.'));
    }
  });
};

const parseCSV = (file, billingType, resolve, reject) => {
  Papa.parse(file, {
    header: true,
    skipEmptyLines: true,
    complete: (results) => {
      try {
        // Filter metadata rows before processing
        // Note: 'description' is excluded as it's a valid transaction column
        const metadataKeywords = [
          'account number', 'currency', 'corporate address',
          'rate of interest', 'ifs code', 'book balance', 'available balance',
          'hold value', 'mod balance', 'uncleared', 'balance on', 'start date', 'end date'
        ];
        
        let extractedStartDate = null;
        let extractedEndDate = null;
        
        // First pass: Extract Start Date and End Date from metadata
        for (const row of results.data) {
          const rowValues = Object.values(row).map(v => String(v || '').trim()).filter(v => v.length > 0);
          if (rowValues.length === 0) continue;
          
          const rowText = rowValues.join(' ').toLowerCase();
          
          // Extract Start Date
          if (rowText.includes('start date') && !extractedStartDate) {
            const dateMatch = rowValues.join(' ').match(/start\s+date\s*[:]?\s*(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}|\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4})/i);
            if (dateMatch && dateMatch[1]) {
              extractedStartDate = dateMatch[1].trim();
            } else {
              for (const cell of rowValues) {
                const cellDateMatch = cell.match(/(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}|\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4})/i);
                if (cellDateMatch && cellDateMatch[1]) {
                  extractedStartDate = cellDateMatch[1].trim();
                  break;
                }
              }
            }
          }
          
          // Extract End Date
          if (rowText.includes('end date') && !extractedEndDate) {
            const dateMatch = rowValues.join(' ').match(/end\s+date\s*[:]?\s*(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}|\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4})/i);
            if (dateMatch && dateMatch[1]) {
              extractedEndDate = dateMatch[1].trim();
            } else {
              for (const cell of rowValues) {
                const cellDateMatch = cell.match(/(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}|\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4})/i);
                if (cellDateMatch && cellDateMatch[1]) {
                  extractedEndDate = cellDateMatch[1].trim();
                  break;
                }
              }
            }
          }
        }
        
        const filteredData = results.data.filter(row => {
          const rowValues = Object.values(row).map(v => String(v || '').trim()).filter(v => v.length > 0);
          if (rowValues.length === 0) return false;
          
          const rowText = rowValues.join(' ').toLowerCase();
          const isMetadataRow = metadataKeywords.some(keyword => {
            const keywordPattern = new RegExp(`^\\s*${keyword.replace(/\s+/g, '\\s+')}\\s*[:]`, 'i');
            return keywordPattern.test(rowText) || 
                   (rowValues.length <= 2 && rowText.includes(keyword) && 
                    !rowText.match(/\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}/));
          });
          
          return !isMetadataRow;
        });
        
        const processed = processData(filteredData, billingType);
        
        // Add extracted dates to the result
        processed.startDate = extractedStartDate;
        processed.endDate = extractedEndDate;
        
        resolve(processed);
      } catch (error) {
        reject(error);
      }
    },
    error: (error) => reject(error),
  });
};

const parseXLSX = (file, billingType, resolve, reject) => {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      // Get raw data as array of arrays to find header row
      // Use raw: false to get formatted text (as displayed in Excel) instead of raw values
      const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: false });
      
      // Expected header columns (case-insensitive)
      const expectedHeaders = ['Txn Date', 'Value Date', 'Description', 'Ref No./Cheque No.', 'Branch Code', 'Debit', 'Credit', 'Balance'];
      const normalizedExpectedHeaders = expectedHeaders.map(h => h.toLowerCase().replace(/[^a-z0-9]/g, ''));
      
      // Metadata keywords to skip (excluding 'description' as it's a valid column name)
      const metadataKeywords = [
        'account number', 'name', 'currency', 'corporate address',
        'rate of interest', 'ifs code', 'book balance', 'available balance',
        'hold value', 'mod balance', 'uncleared', 'balance on', 'start date', 'end date'
      ];
      
      let headerRowIndex = -1;
      let cleanedRawData = [];
      let extractedStartDate = null;
      let extractedEndDate = null;
      
      // First pass: Extract Start Date and End Date from metadata before filtering
      for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        if (!row || row.length === 0) continue;
        
        const rowValues = row.map(cell => String(cell || '').trim()).filter(v => v.length > 0);
        if (rowValues.length === 0) continue;
        
        const rowText = rowValues.join(' ').toLowerCase();
        
        // Extract Start Date
        if (rowText.includes('start date') && !extractedStartDate) {
          // Look for date value in the row (usually after "Start Date :")
          const dateMatch = rowValues.join(' ').match(/start\s+date\s*[:]?\s*(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}|\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4})/i);
          if (dateMatch && dateMatch[1]) {
            extractedStartDate = dateMatch[1].trim();
          } else {
            // Check if any cell contains a date pattern
            for (const cell of rowValues) {
              const cellDateMatch = cell.match(/(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}|\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4})/i);
              if (cellDateMatch && cellDateMatch[1]) {
                extractedStartDate = cellDateMatch[1].trim();
                break;
              }
            }
          }
        }
        
        // Extract End Date
        if (rowText.includes('end date') && !extractedEndDate) {
          // Look for date value in the row (usually after "End Date :")
          const dateMatch = rowValues.join(' ').match(/end\s+date\s*[:]?\s*(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}|\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4})/i);
          if (dateMatch && dateMatch[1]) {
            extractedEndDate = dateMatch[1].trim();
          } else {
            // Check if any cell contains a date pattern
            for (const cell of rowValues) {
              const cellDateMatch = cell.match(/(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}|\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4})/i);
              if (cellDateMatch && cellDateMatch[1]) {
                extractedEndDate = cellDateMatch[1].trim();
                break;
              }
            }
          }
        }
      }
      
      // Find the header row
      for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        if (!row || row.length === 0) continue;
        
        // Convert row to string array for comparison
        const rowValues = row.map(cell => String(cell || '').trim()).filter(v => v.length > 0);
        if (rowValues.length === 0) continue;
        
        // Check if this row matches metadata pattern (but not if it looks like a header row)
        const rowText = rowValues.join(' ').toLowerCase();
        // Skip rows that look like metadata (e.g., "Account Number : 123" or "Branch : NELSON")
        // But don't skip if it contains valid headers like "Branch Code"
        
        // Special check for "Branch :" metadata (but not "Branch Code" header)
        const isBranchMetadata = rowText.match(/^branch\s*[:]/i) && !rowText.includes('branch code');
        
        // Check other metadata keywords
        const isOtherMetadata = metadataKeywords.some(keyword => {
          // Check if row starts with metadata keyword followed by colon/colon-like pattern
          const keywordPattern = new RegExp(`^\\s*${keyword.replace(/\s+/g, '\\s+')}\\s*[:]`, 'i');
          return keywordPattern.test(rowText);
        });
        
        const isMetadata = isBranchMetadata || isOtherMetadata;
        
        if (isMetadata) {
          continue; // Skip metadata rows
        }
        
        // Check if this row matches the expected header
        const normalizedRowHeaders = rowValues.map(h => h.toLowerCase().replace(/[^a-z0-9]/g, ''));
        
        // Check if all expected headers are present (allow partial matches)
        const matches = normalizedExpectedHeaders.filter(expected => 
          normalizedRowHeaders.some(rowH => rowH.includes(expected) || expected.includes(rowH))
        );
        
        if (matches.length >= 5) { // At least 5 out of 8 columns should match
          headerRowIndex = i;
          cleanedRawData = rawData.slice(i + 1); // Start from row after header
          break;
        }
      }
      
      // If header not found, use all rows but filter metadata
      if (headerRowIndex === -1) {
        cleanedRawData = rawData.filter(row => {
          if (!row || row.length === 0) return false;
          const rowValues = row.map(cell => String(cell || '').trim()).filter(v => v.length > 0);
          if (rowValues.length === 0) return false;
          const rowText = rowValues.join(' ').toLowerCase();
          return !metadataKeywords.some(keyword => rowText.includes(keyword));
        });
      }
      
      // Convert cleaned data to JSON with header from first row or expected headers
      let jsonData = [];
      if (headerRowIndex !== -1 && rawData[headerRowIndex]) {
        // Use the found header row
        const headers = rawData[headerRowIndex].map(cell => String(cell || '').trim());
        
        // Don't convert dates - use Excel's formatted text as-is
        jsonData = cleanedRawData.map(row => {
          const obj = {};
          headers.forEach((header, idx) => {
            if (header && row[idx] !== null && row[idx] !== undefined) {
              // Use value as-is (already formatted by Excel)
              obj[header] = row[idx];
            }
          });
          return obj;
        }).filter(obj => Object.keys(obj).length > 0);
      } else {
        // Fallback: use default XLSX parsing but filter metadata
        // Use raw: false to get formatted text (as displayed in Excel) instead of raw values
        const allJsonData = XLSX.utils.sheet_to_json(worksheet, { defval: null, cellDates: false, raw: false });
        
        // Don't convert dates - use Excel's formatted text as-is
        jsonData = allJsonData.filter(row => {
          const rowValues = Object.values(row).map(v => String(v || '').trim()).filter(v => v.length > 0);
          const rowText = rowValues.join(' ').toLowerCase();
          return !metadataKeywords.some(keyword => rowText.includes(keyword));
        });
      }
      
      const processed = processData(jsonData, billingType);
      
      // Add extracted dates to the result
      processed.startDate = extractedStartDate;
      processed.endDate = extractedEndDate;
      
      resolve(processed);
    } catch (error) {
      reject(error);
    }
  };
  reader.onerror = () => reject(new Error('Error reading file'));
  reader.readAsArrayBuffer(file);
};

const processData = (data, billingType) => {
  if (!data || data.length === 0) {
    throw new Error('No data found in file');
  }

  // Metadata keywords to filter out (additional safety check)
  // Note: 'description' is excluded as it's a valid transaction column
  const metadataKeywords = [
    'account number', 'currency', 'corporate address',
    'rate of interest', 'ifs code', 'book balance', 'available balance',
    'hold value', 'mod balance', 'uncleared', 'balance on', 'start date', 'end date'
  ];

  // Filter out metadata rows that might have slipped through
  const filteredData = data.filter(row => {
    const rowValues = Object.values(row).map(v => String(v || '').trim()).filter(v => v.length > 0);
    if (rowValues.length === 0) return false;
    
    const rowText = rowValues.join(' ').toLowerCase();
    
    // Check if row contains only metadata keywords (like "Account Number : 123")
    const isMetadataRow = metadataKeywords.some(keyword => {
      // Check if the row starts with or is primarily metadata
      const keywordPattern = new RegExp(`^\\s*${keyword.replace(/\s+/g, '\\s+')}\\s*[:]`, 'i');
      return keywordPattern.test(rowText) || 
             (rowValues.length <= 2 && rowText.includes(keyword) && 
              !rowText.match(/\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}/)); // Exclude if it has dates
    });
    
    return !isMetadataRow;
  });

  if (filteredData.length === 0) {
    throw new Error('No valid data found after filtering metadata');
  }

  // Normalize column names (trim, lowercase, remove special chars)
  const normalized = filteredData.map(row => {
    const normalizedRow = {};
    Object.keys(row).forEach(key => {
      const normalizedKey = key.trim().toLowerCase()
        .replace(/[^a-z0-9]/g, '')
        .replace(/\s+/g, '');
      normalizedRow[normalizedKey] = row[key];
    });
    return normalizedRow;
  });

  // Get the identifier column based on billing type
  const idColumn = billingType === 'student' 
    ? ['invoiceno', 'invoicenumber', 'invoice', 'voucherno']
    : ['voucherno', 'vouchernumber', 'voucher', 'invoiceno'];

  // Find the actual column name
  let identifierColumn = null;
  for (const row of normalized) {
    for (const key of Object.keys(row)) {
      if (idColumn.some(id => key.includes(id))) {
        identifierColumn = key;
        break;
      }
    }
    if (identifierColumn) break;
  }

  // Remove duplicates based on identifier column
  const uniqueData = removeDuplicates(normalized, identifierColumn);

  return {
    data: uniqueData,
    originalCount: normalized.length,
    uniqueCount: uniqueData.length,
    identifierColumn,
    startDate: null, // Will be set by parser if found
    endDate: null, // Will be set by parser if found
  };
};

const parsePDF = async (file, billingType) => {
  try {
    const pdfjs = await loadPdfJs();
    if (!pdfjs || !pdfjs.getDocument) {
      throw new Error('PDF.js library is not properly loaded');
    }
    
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjs.getDocument({ 
      data: arrayBuffer,
      useSystemFonts: true,
    }).promise;
    
    let allRows = [];
    
    // Extract text with positions from all pages
    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      const page = await pdf.getPage(pageNum);
      const textContent = await page.getTextContent();
      
      // Extract text items with their positions
      const textItems = textContent.items.map(item => ({
        text: item.str,
        x: item.transform[4],
        y: item.transform[5],
        width: item.width,
        height: item.height,
      }));
      
      // Group text by Y position (rows) and sort by X (columns)
      const rowsByY = {};
      textItems.forEach(item => {
        const y = Math.round(item.y * 10) / 10; // Round to nearest 0.1 for row grouping
        if (!rowsByY[y]) {
          rowsByY[y] = [];
        }
        rowsByY[y].push(item);
      });
      
      // Process each row - preserve all text with spacing
      const pageRows = Object.keys(rowsByY)
        .sort((a, b) => parseFloat(b) - parseFloat(a)) // Sort from top to bottom
        .map(y => {
          const items = rowsByY[y].sort((a, b) => a.x - b.x); // Sort from left to right
          
          // Build row with proper spacing to preserve all text
          let rowText = '';
          for (let i = 0; i < items.length; i++) {
            if (i > 0) {
              // Calculate spacing between items to preserve column structure
              const spaceWidth = items[i].x - (items[i-1].x + items[i-1].width);
              // Use tabs for large gaps (column separators), spaces for small gaps
              if (spaceWidth > items[i-1].width * 1.5) {
                rowText += '\t'; // Tab for column separation
              } else {
                rowText += ' '; // Single space for normal text
              }
            }
            rowText += items[i].text;
          }
          return rowText;
        });
      
      allRows.push(...pageRows);
    }

    // Now parse the rows to extract structured data
    const rows = [];
    let headers = [];
    let headerFound = false;
    let dataStarted = false;
    
    // Common column headers for bank statements
    const headerPatterns = {
      txndate: ['txn date', 'transaction date', 'date', 'trans date'],
      valuedate: ['value date', 'val date'],
      description: ['description', 'particulars', 'narration', 'details', 'desc'],
      refno: ['ref no', 'reference no', 'ref number', 'cheque no', 'chq no', 'utr no', 'utr'],
      branchcode: ['branch code', 'branch', 'br code', 'br'],
      debit: ['debit', 'dr', 'withdrawal', 'paid'],
      credit: ['credit', 'cr', 'deposit', 'received'],
      balance: ['balance', 'bal', 'closing balance']
    };
    
    for (let i = 0; i < allRows.length; i++) {
      const line = allRows[i].trim();
      
      if (!line) continue;
      
      // Look for header row
      if (!headerFound && i < 30) {
        const lowerLine = line.toLowerCase();
        let foundHeaders = false;
        
        // Check if this line contains header keywords
        for (const [key, patterns] of Object.entries(headerPatterns)) {
          if (patterns.some(pattern => lowerLine.includes(pattern))) {
            foundHeaders = true;
            break;
          }
        }
        
        if (foundHeaders) {
          // Try to split the header row
          const headerParts = line.split(/\s{2,}|\t/).map(h => h.trim()).filter(h => h.length > 0);
          if (headerParts.length >= 3) {
            headers = headerParts;
            headerFound = true;
            dataStarted = true;
            continue;
          }
        }
      }
      
      // Process data rows
      if (dataStarted || i > 10) {
        // Try to parse as a data row
        // Bank statement rows typically have:
        // Date Date Description Ref No. Branch Code Debit Credit Balance
        
        // Pattern: Date pattern (dd MMM yyyy or dd/MM/yyyy)
        const datePattern = /(\d{1,2}\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4})|(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4})/i;
        const dateMatch = line.match(datePattern);
        
        if (dateMatch || line.match(/\d{1,2}\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/i)) {
          // This looks like a data row
          const row = {};
          
          // Split by tabs or multiple spaces (at least 2 spaces) to separate columns
          // Preserve all text by using both tabs and multiple spaces as delimiters
          // Also handle single spaces more carefully for better column separation
          let parts = line.split(/\t/).map(p => p.trim()).filter(p => p.length > 0);
          
          // If no tabs found, split by multiple spaces (2+)
          if (parts.length < 4) {
            parts = line.split(/\s{2,}/).map(p => p.trim()).filter(p => p.length > 0);
          }
          
          // If still not enough parts, try splitting by single space but preserve structure
          if (parts.length < 3 && line.includes(' ')) {
            // For lines with mixed spacing, be smarter about splitting
            // Look for pattern: Date Date Description... RefNo BranchCode Amount Amount Amount
            const allWords = line.split(/\s+/).filter(w => w.length > 0);
            if (allWords.length >= 5) {
              // Try to reconstruct parts by grouping consecutive words that aren't amounts
              parts = [];
              let currentPart = '';
              for (let w = 0; w < allWords.length; w++) {
                const word = allWords[w];
                const cleanWord = word.replace(/,/g, '').trim();
                const isAmount = /^[\d,]+\.?\d{0,2}$/.test(cleanWord) || /^\d+\.?\d{2}$/.test(cleanWord);
                
                // If it's an amount, start a new part
                if (isAmount && currentPart) {
                  parts.push(currentPart.trim());
                  currentPart = word;
                } else {
                  if (currentPart) {
                    currentPart += ' ' + word;
                  } else {
                    currentPart = word;
                  }
                }
              }
              if (currentPart) {
                parts.push(currentPart.trim());
              }
            } else {
              parts = allWords;
            }
          }
          
          if (parts.length >= 2) {
            let partIndex = 0;
            
            // Extract dates (usually first 1-2 parts)
            if (parts[partIndex] && datePattern.test(parts[partIndex])) {
              row.txndate = parts[partIndex].trim();
              partIndex++;
              
              // Second date might be value date
              if (parts[partIndex] && datePattern.test(parts[partIndex])) {
                row.valuedate = parts[partIndex].trim();
                partIndex++;
              } else {
                row.valuedate = row.txndate; // Same as transaction date if not provided
              }
            }
            
            // Description might span multiple parts - extract ALL text that's not clearly a number or reference
            // We want to capture everything in description field - even if it spans many parts
            let descriptionParts = [];
            
            // First pass: Identify structured fields (refno, branchcode, amounts) and their positions
            const structuredFields = [];
            let amountStartIndex = -1;
            
            // Scan all parts to identify refno, branchcode, and amounts
            for (let j = partIndex; j < parts.length; j++) {
              const part = parts[j];
              const cleanPart = part.replace(/,/g, '').trim();
              const partTrimmed = part.trim();
              
              // Check if it's a clear amount pattern (numbers with optional decimals)
              const isAmount = /^[\d,]+\.?\d{0,2}$/.test(cleanPart) || /^\d+\.?\d{2}$/.test(cleanPart);
              
              // IMPROVED Reference Number Detection:
              // - UTR numbers: 10-15 digits (like "4897736162097", "99826044309")
              // - Cheque numbers: 6-8 digits (like "264852")
              // - RTGS/NEFT refs: Alphanumeric patterns (like "SBINR12025102903", "CR0AOXXGS6", "CR0APIOOL9")
              // - UPI refs: Long numeric strings in UPI transactions
              const isRefNo = 
                // Long numeric strings (10-15 digits) - UTR, account numbers (like "4897736162097", "99826044309")
                (/^\d{10,15}$/.test(cleanPart)) ||
                // Medium numeric strings (6-11 digits) - Cheque numbers, shorter refs
                (/^\d{6,11}$/.test(cleanPart) && j > partIndex + 2 && !isAmount) ||
                // Alphanumeric patterns with letters and numbers (RTGS/NEFT) - like "CR0AOXXGS6", "CR0APIOOL9"
                (/^[A-Z0-9]{8,20}$/i.test(partTrimmed) && 
                 /[A-Z]/i.test(partTrimmed) && 
                 /[0-9]/.test(partTrimmed) &&
                 !isAmount &&
                 partTrimmed.length >= 8 && partTrimmed.length <= 20 &&
                 // Exclude common words
                 !/^(TRANSFER|ELLEN|INFORMATION|TECHNOLOGY|SOLUTI|INB|FROM|TO|ELLENINFORMATI)$/i.test(partTrimmed));
              
              // IMPROVED Branch Code Detection:
              // Branch codes are short (3-6 chars), alphanumeric, appear before amounts
              // Examples: "11606", "99922", "889574", "10395", "8085", "4292", "4430"
              const isBranchCode = 
                // Short alphanumeric (3-6 chars)
                (/^[A-Z0-9]{3,6}$/i.test(partTrimmed) && 
                 partTrimmed.length >= 3 && 
                 partTrimmed.length <= 6 && 
                 !isAmount &&
                 // Should appear before amounts (usually in last 2-3 columns before amounts)
                 j < parts.length - 3);
              
              // If we find a clear amount, mark this as where structured data starts
              if (isAmount && amountStartIndex === -1) {
                // Look ahead: if next part is also an amount, this is likely structured data section
                if (j + 1 < parts.length) {
                  const nextPart = parts[j + 1].replace(/,/g, '').trim();
                  if (/^[\d,]+\.?\d{0,2}$/.test(nextPart) || /^\d+\.?\d{2}$/.test(nextPart)) {
                    amountStartIndex = j;
                    
                    // Before amounts, check for refno and branchcode
                    if (isRefNo && !structuredFields.find(f => f.type === 'refno')) {
                      structuredFields.push({ type: 'refno', value: partTrimmed, index: j });
                    }
                    if (isBranchCode && !structuredFields.find(f => f.type === 'branchcode')) {
                      structuredFields.push({ type: 'branchcode', value: partTrimmed, index: j });
                    }
                    
                    break;
                  }
                }
              }
              
              // Collect refno and branchcode candidates before we reach amounts
              if (amountStartIndex === -1) {
                // Look for refno (usually appears in description area but is structured data)
                if (isRefNo && !structuredFields.find(f => f.type === 'refno')) {
                  // Prefer refnos that are clearly not description text (long, alphanumeric)
                  structuredFields.push({ type: 'refno', value: partTrimmed, index: j });
                }
                // Look for branch code (usually appears right before amounts)
                if (isBranchCode && !structuredFields.find(f => f.type === 'branchcode')) {
                  structuredFields.push({ type: 'branchcode', value: partTrimmed, index: j });
                }
              }
            }
            
            // Assign refno and branchcode from structured fields found BEFORE amounts
            const refNoField = structuredFields.find(f => f.type === 'refno');
            const branchCodeField = structuredFields.find(f => f.type === 'branchcode');
            
            if (refNoField && !row.refno) {
              row.refno = refNoField.value;
            }
            if (branchCodeField && !row.branchcode) {
              row.branchcode = branchCodeField.value;
            }
            
            // Extract description: everything from current index to amount start
            // But exclude parts that are clearly refno or branchcode
            const descriptionEndIndex = amountStartIndex > -1 ? amountStartIndex : parts.length;
            
            for (let j = partIndex; j < descriptionEndIndex; j++) {
              const part = parts[j];
              
              // Check if this part is identified as a structured field (refno or branchcode)
              const isStructuredField = structuredFields.some(f => f.index === j);
              
              // Also check if it's a clear refno pattern (long numeric or alphanumeric)
              const cleanPart = part.replace(/,/g, '').trim();
              const isStandaloneRefNo = 
                /^\d{12,}$/.test(cleanPart) || // Long numeric UTR/UPI
                (/^[A-Z0-9\/\-]{8,}$/i.test(part.trim()) && /[A-Z]/i.test(part.trim()) && /[0-9]/.test(part.trim())); // RTGS/NEFT pattern
              
              // Check if it's a standalone branch code
              const isStandaloneBranchCode = /^[A-Z0-9]{3,6}$/i.test(part.trim()) && 
                                            part.trim().length >= 3 && 
                                            part.trim().length <= 6 &&
                                            j >= descriptionEndIndex - 3; // Appears in last few parts before amounts
              
              if (!isStructuredField && !isStandaloneRefNo && !isStandaloneBranchCode) {
                descriptionParts.push(part);
              }
            }
            
            // Join all description parts - preserve ALL text with spaces
            // Normalize multiple spaces to single space for better pattern matching
            const fullDescription = descriptionParts.join(' ').replace(/\s+/g, ' ').trim();
            row.description = fullDescription;
            
            // Extract refno from description if not found yet
            // Handle multi-line patterns where "RTGS INB:" and ref number might be on different lines
            if (!row.refno && fullDescription) {
              // Clean description - normalize whitespace but preserve structure
              const normalizedDesc = fullDescription.replace(/\s+/g, ' ').trim();
              
              // Pattern 1: RTGS INB: followed by alphanumeric ref (handles multi-line)
              // Examples: "RTGS INB: CR0AOXXGS6", "RTGS INB CR0APIOOL9"
              const rtgsInbPattern = /RTGS\s+INB[:\s]+([A-Z0-9]{8,20})/i;
              const rtgsInbMatch = normalizedDesc.match(rtgsInbPattern);
              if (rtgsInbMatch && rtgsInbMatch[1]) {
                row.refno = rtgsInbMatch[1].trim();
              }
              
              // Pattern 2: RTGS/NEFT/UTR followed by ref (broader pattern)
              if (!row.refno) {
                const rtgsNeftPattern = /(?:RTGS|NEFT|UTR)[\s]+(?:INB|NO)?[:\s]*([A-Z0-9]{8,20})/i;
                const rtgsMatch = normalizedDesc.match(rtgsNeftPattern);
                if (rtgsMatch && rtgsMatch[1]) {
                  row.refno = rtgsMatch[1].trim();
                }
              }
              
              // Pattern 3: TRANSFER FROM followed by long numeric (account/UTR numbers)
              // Examples: "TRANSFER FROM 4897736162097", "TRANSFER FROM 99826044309"
              if (!row.refno) {
                const transferFromPattern = /TRANSFER\s+FROM\s+(\d{10,})/i;
                const transferMatch = normalizedDesc.match(transferFromPattern);
                if (transferMatch && transferMatch[1]) {
                  row.refno = transferMatch[1].trim();
                }
              }
              
              // Pattern 4: TRANSFER TO followed by long numeric
              if (!row.refno) {
                const transferToPattern = /TRANSFER\s+TO\s+(\d{10,})/i;
                const transferToMatch = normalizedDesc.match(transferToPattern);
                if (transferToMatch && transferToMatch[1]) {
                  row.refno = transferToMatch[1].trim();
                }
              }
              
              // Pattern 5: Cheque numbers (Chq 264852, Cheque 264852, etc.)
              if (!row.refno) {
                const chequePattern = /(?:Chq|Cheque|Chq\.)[\s]*(\d{6,})/i;
                const chequeMatch = normalizedDesc.match(chequePattern);
                if (chequeMatch && chequeMatch[1]) {
                  row.refno = chequeMatch[1].trim();
                }
              }
              
              // Pattern 6: Standalone alphanumeric refs (CR0AOXXGS6, CR0APIOOL9, SBINR12025102903, etc.)
              // Look for patterns with both letters and numbers, 8-20 chars
              if (!row.refno) {
                // Match all potential alphanumeric refs
                const alphanumericMatches = normalizedDesc.matchAll(/([A-Z0-9]{8,20})/gi);
                for (const match of alphanumericMatches) {
                  const candidate = match[1];
                  // Must have both letters and numbers, and reasonable length
                  if (/[A-Z]/i.test(candidate) && 
                      /[0-9]/.test(candidate) && 
                      candidate.length >= 8 && 
                      candidate.length <= 20 &&
                      // Should not be part of common words
                      !/^(TRANSFER|ELLEN|INFORMATION|TECHNOLOGY|SOLUTI|INB|FROM|TO)$/i.test(candidate)) {
                    row.refno = candidate.trim();
                    break;
                  }
                }
              }
              
              // Pattern 7: Long numeric strings (10+ digits) - UTR/UPI refs, account numbers
              // Examples: "4897736162097", "99826044309"
              if (!row.refno) {
                // Match all long numeric strings
                const longNumericMatches = normalizedDesc.matchAll(/(\d{10,})/g);
                for (const match of longNumericMatches) {
                  const candidate = match[1];
                  // Prefer 10-15 digit numbers (likely UTR/account refs)
                  // Avoid very long numbers that might be part of description
                  if (candidate.length >= 10 && candidate.length <= 15) {
                    // Check if it's not part of a date or amount
                    const candidateIndex = normalizedDesc.indexOf(candidate);
                    const before = normalizedDesc.substring(Math.max(0, candidateIndex - 10), candidateIndex);
                    const after = normalizedDesc.substring(candidateIndex + candidate.length, candidateIndex + candidate.length + 10);
                    
                    // Should not be immediately after amount-like patterns
                    if (!before.match(/[,\d]{3,}\s*$/) && !after.match(/^[,\d]{3,}/)) {
                      row.refno = candidate.trim();
                      break;
                    }
                  }
                }
              }
              
              // Pattern 8: Medium numeric strings (8-11 digits) that appear after keywords
              if (!row.refno) {
                const mediumNumericPattern = /(?:UTR|UPI|REF|REFERENCE|NO|NUMBER)[\s:]*(\d{8,11})/i;
                const mediumMatch = normalizedDesc.match(mediumNumericPattern);
                if (mediumMatch && mediumMatch[1]) {
                  row.refno = mediumMatch[1].trim();
                }
              }
            }
            
            // Update partIndex to where we stopped
            partIndex = descriptionEndIndex;
            
            // Extract reference number, branch code, debit, credit, balance
            // Process remaining parts more carefully to capture ALL data without missing anything
            const remainingParts = parts.slice(partIndex);
            
            // First, identify ALL numeric amounts
            // Also check for refno and branchcode in remaining parts
            const amounts = [];
            
            for (let j = 0; j < remainingParts.length; j++) {
              const part = remainingParts[j];
              const cleanPart = part.replace(/,/g, '').trim();
              const partTrimmed = part.trim();
              
              // Check for amounts (numbers with optional decimals) - Indian number format with commas
              if (/^[\d,]+\.?\d{0,2}$/.test(cleanPart) || /^\d+\.?\d{2}$/.test(cleanPart)) {
                const numericValue = parseFloat(cleanPart.replace(/,/g, ''));
                if (!isNaN(numericValue) && numericValue > 0) {
                  amounts.push({
                    value: numericValue,
                    index: j,
                    original: part
                  });
                }
              }
              // If refno not found yet, check for it in remaining parts (before amounts)
              else if (!row.refno && j < remainingParts.length - 2) {
                // Check for alphanumeric refs (RTGS/NEFT patterns)
                if (/^[A-Z0-9]{8,}$/i.test(partTrimmed) && 
                    /[A-Z]/i.test(partTrimmed) && 
                    /[0-9]/.test(partTrimmed) &&
                    !isNaN(parseFloat(cleanPart)) === false) {
                  row.refno = partTrimmed;
                }
                // Check for long numeric refs (UTR)
                else if (/^\d{12,}$/.test(cleanPart)) {
                  row.refno = partTrimmed;
                }
                // Check for medium numeric refs (cheque numbers)
                else if (/^\d{6,11}$/.test(cleanPart) && j === 0) {
                  row.refno = partTrimmed;
                }
              }
              // If branchcode not found yet, check for it in remaining parts (usually before amounts)
              else if (!row.branchcode && j < remainingParts.length - 2 && amounts.length === 0) {
                // Branch code is usually short numeric or alphanumeric (3-6 chars)
                // Examples: "99922", "10395", "11606", "8085", "4292", "4430"
                const isBranchCodePattern = 
                  /^[A-Z0-9]{3,6}$/i.test(partTrimmed) && 
                  partTrimmed.length >= 3 && 
                  partTrimmed.length <= 6 &&
                  // Should be numeric or mostly numeric (like "99922", "11606")
                  (/^\d{3,6}$/.test(partTrimmed) || 
                   (/\d/.test(partTrimmed) && partTrimmed.length <= 6));
                
                if (isBranchCodePattern && !isNaN(parseFloat(cleanPart.replace(/[^0-9]/g, '')))) {
                  row.branchcode = partTrimmed;
                }
              }
            }
            
            // If branchcode still not found, check description for patterns
            if (!row.branchcode && fullDescription) {
              // Look for branch code patterns at the end of description (before amounts)
              // Usually appears as standalone number: "99922", "10395", etc.
              // Try multiple patterns
              
              // Pattern 1: Standalone number at the end (after description text)
              const branchCodeAtEnd = fullDescription.match(/\b([0-9]{3,6})\s*$/);
              if (branchCodeAtEnd && branchCodeAtEnd[1]) {
                const code = branchCodeAtEnd[1];
                // Verify it's not part of a larger number or account number
                if (code.length >= 3 && code.length <= 6 && 
                    !fullDescription.includes(code + '.') &&
                    !fullDescription.match(new RegExp(code + '\\d{4,}'))) { // Not part of account number
                  row.branchcode = code;
                }
              }
              
              // Pattern 2: Look for branch code patterns before common keywords
              // Usually appears before "TRANSFER TO", "ELLEN", etc. or after account numbers
              if (!row.branchcode) {
                // Match standalone 3-6 digit numbers that appear isolated
                const allNumbers = fullDescription.match(/\b(\d{3,6})\b/g);
                if (allNumbers) {
                  // Filter out numbers that are likely part of account numbers, amounts, or dates
                  for (let i = allNumbers.length - 1; i >= 0; i--) {
                    const num = allNumbers[i];
                    // Check if it appears isolated (not part of larger number)
                    const numPattern = new RegExp(`\\b${num}\\b`);
                    const beforeMatch = fullDescription.substring(0, fullDescription.search(numPattern));
                    const afterMatch = fullDescription.substring(fullDescription.search(numPattern) + num.length);
                    
                    // Should not have digits immediately before or after (except spaces)
                    const isolated = (!/\d/.test(beforeMatch.slice(-1)) && !/\d/.test(afterMatch[0]));
                    
                    if (isolated && num.length >= 3 && num.length <= 6 && 
                        !num.includes('.') && !num.includes(',')) {
                      // This looks like a branch code
                      row.branchcode = num;
                      break;
                    }
                  }
                }
              }
            }
            
            // Final fallback: Look in original parts array for branchcode
            if (!row.branchcode && parts.length > partIndex) {
              // Check parts right before amounts (usually last 2-3 parts)
              for (let k = Math.max(partIndex - 3, 0); k < partIndex; k++) {
                const part = parts[k];
                const cleanPart = part.replace(/,/g, '').trim();
                
                // Branch code is usually 3-6 digit numeric
                if (/^\d{3,6}$/.test(cleanPart) && 
                    cleanPart.length >= 3 && cleanPart.length <= 6 &&
                    !cleanPart.includes('.') && !cleanPart.includes(',')) {
                  // Check if it's not already in description (avoid duplicate)
                  if (!fullDescription.includes(cleanPart) || 
                      fullDescription.split(cleanPart).length === 2) {
                    row.branchcode = cleanPart;
                    break;
                  }
                }
              }
            }
            
            // Assign amounts based on position and context
            // Bank statement format: typically has amounts in order (debit, credit, balance) or (credit, balance) or (amount, balance)
            const descLower = (row.description || '').toLowerCase();
            const isCreditTransaction = descLower.includes('credit') || descLower.includes('cr') ||
                                       descLower.includes('deposit') || descLower.includes('received') ||
                                       descLower.includes('transfer from') || descLower.includes('by transfer') ||
                                       descLower.includes('chq');
            const isDebitTransaction = descLower.includes('debit') || descLower.includes('dr') ||
                                       descLower.includes('withdrawal') || descLower.includes('paid') ||
                                       descLower.includes('transfer to') || descLower.includes('atm') ||
                                       descLower.includes('to transfer') || descLower.includes('wdl');
            
            if (amounts.length > 0) {
              // Bank statement format: usually has 3 amounts (debit, credit, balance) or 2 (amount, balance)
              if (amounts.length >= 3) {
                // Three amounts: typically debit, credit, balance (or credit, debit, balance)
                // Based on description, determine which is which
                if (isDebitTransaction && !isCreditTransaction) {
                  row.debit = amounts[0].value.toFixed(2);
                  row.credit = amounts[1].value.toFixed(2);
                  row.balance = amounts[2].value.toFixed(2);
                } else if (isCreditTransaction && !isDebitTransaction) {
                  row.credit = amounts[0].value.toFixed(2);
                  row.debit = amounts[1].value.toFixed(2);
                  row.balance = amounts[2].value.toFixed(2);
                } else {
                  // Unknown: assign first two as debit/credit, last as balance
                  // Check which is larger to guess debit vs credit
                  if (amounts[0].value > amounts[1].value) {
                    row.credit = amounts[0].value.toFixed(2);
                    row.debit = amounts[1].value.toFixed(2);
                  } else {
                    row.debit = amounts[0].value.toFixed(2);
                    row.credit = amounts[1].value.toFixed(2);
                  }
                  row.balance = amounts[2].value.toFixed(2);
                }
              } else if (amounts.length === 2) {
                // Two amounts: could be (debit/credit, balance) or (debit, credit)
                // Usually the last amount is balance
                if (isDebitTransaction) {
                  row.debit = amounts[0].value.toFixed(2);
                  row.balance = amounts[1].value.toFixed(2);
                  row.credit = '';
                } else if (isCreditTransaction) {
                  row.credit = amounts[0].value.toFixed(2);
                  row.balance = amounts[1].value.toFixed(2);
                  row.debit = '';
                } else {
                  // Unknown: assume first is amount, second is balance
                  row.credit = amounts[0].value.toFixed(2);
                  row.balance = amounts[1].value.toFixed(2);
                  row.debit = '';
                }
              } else if (amounts.length === 1) {
                // Single amount: could be credit or debit based on description
                if (isDebitTransaction) {
                  row.debit = amounts[0].value.toFixed(2);
                  row.credit = '';
                } else if (isCreditTransaction) {
                  row.credit = amounts[0].value.toFixed(2);
                  row.debit = '';
                } else {
                  // Default to credit for single amount
                  row.credit = amounts[0].value.toFixed(2);
                  row.debit = '';
                }
              }
            }
            
            // Clean up the row - ensure all standard fields exist
            // Preserve ALL description text - don't truncate or modify
            const cleanRow = {
              txndate: row.txndate || '',
              valuedate: row.valuedate || row.txndate || '',
              description: row.description || '', // Keep full description - no truncation
              refno: row.refno || '',
              branchcode: row.branchcode || '',
              debit: row.debit || '',
              credit: row.credit || '',
              balance: row.balance || '',
            };
            
            // Only add if it has meaningful data (date or description or amounts)
            if (cleanRow.txndate || cleanRow.description || cleanRow.debit || cleanRow.credit || cleanRow.balance) {
              rows.push(cleanRow);
            }
          }
        }
      }
    }

    if (rows.length > 0) {
      const processed = processData(rows, billingType);
      return processed;
    } else {
      // Fallback: try simpler parsing
      const fullText = allRows.join('\n');
      const fallbackRows = [];
      
      // Extract patterns: Date, Description, Amount
      const dateMatches = fullText.matchAll(/(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4})/gi);
      const dates = Array.from(dateMatches).map(m => m[1]);
      
      // Try to match dates with amounts
      for (let i = 0; i < dates.length; i++) {
        const dateStr = dates[i];
        const dateIndex = fullText.indexOf(dateStr);
        if (dateIndex !== -1) {
          const afterDate = fullText.substring(dateIndex + dateStr.length, dateIndex + 500);
          const amountMatch = afterDate.match(/[\d,]+\.?\d{0,2}/);
          
          if (amountMatch) {
            fallbackRows.push({
              txndate: dateStr,
              valuedate: dateStr,
              description: afterDate.trim(), // Don't truncate - preserve full description
              amount: amountMatch[0],
            });
          }
        }
      }
      
      if (fallbackRows.length > 0) {
        const processed = processData(fallbackRows, billingType);
        return processed;
      }
      
      // Last resort: return all text as a single row
      const processed = processData([{
        content: fullText.substring(0, 1000),
        extractedtext: fullText
      }], billingType);
      return processed;
    }
  } catch (error) {
    throw new Error(`Error parsing PDF: ${error.message}`);
  }
};

const parseDOC = async (file, billingType, fileExtension) => {
  try {
    if (fileExtension === 'docx') {
      // Load mammoth library dynamically
      const mammoth = await loadMammoth();
      
      // Parse DOCX file
      const arrayBuffer = await file.arrayBuffer();
      
      // Extract raw text
      const result = await mammoth.extractRawText({ arrayBuffer });
      const text = result.value;
      
      // Try to extract table data - mammoth has better table support
      const tableResult = await mammoth.convertToHtml({ arrayBuffer });
      
      // Parse text into structured data
      const lines = text.split('\n').filter(line => line.trim().length > 0);
      const rows = [];
      let headers = [];
      
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        
        // Identify headers (look for common invoice/voucher keywords)
        if (i < 15 && !headers.length) {
          const headerKeywords = ['date', 'amount', 'name', 'invoice', 'voucher', 'particulars', 'credit', 'debit', 'transaction', 'description'];
          const lowerLine = line.toLowerCase();
          if (headerKeywords.some(keyword => lowerLine.includes(keyword))) {
            // Try splitting by tabs, multiple spaces, or commas
            headers = line.split(/\t|\s{3,}|,\s*/).map(h => h.trim()).filter(h => h.length > 0);
            if (headers.length >= 2) {
              continue;
            }
          }
        }
        
        // Parse data rows - split by tabs, multiple spaces, or commas
        const parts = line.split(/\t|\s{3,}|,\s*/).map(p => p.trim()).filter(p => p.length > 0);
        
        if (parts.length >= 2) {
          const row = {};
          if (headers.length > 0) {
            headers.forEach((header, idx) => {
              if (parts[idx] && parts[idx] !== '') {
                const normalizedHeader = header.toLowerCase().replace(/[^a-z0-9]/g, '');
                row[normalizedHeader] = parts[idx];
              }
            });
          } else {
            // Auto-generate headers based on common patterns
            parts.forEach((part, idx) => {
              const possibleKeys = ['date', 'name', 'amount', 'description', 'invoice', 'voucher'];
              const key = idx < possibleKeys.length ? possibleKeys[idx] : `column${idx + 1}`;
              row[key] = part;
            });
          }
          
          if (Object.keys(row).length > 0) {
            rows.push(row);
          }
        }
      }

      if (rows.length > 0) {
        const processed = processData(rows, billingType);
        return processed;
      } else {
        // Fallback: try to extract any structured information from text
        // Look for date patterns, amounts, etc.
        const fallbackRow = {
          content: text.substring(0, 1000),
          rawText: text.substring(0, 5000),
          extractedtext: text
        };
        
        // Try to find dates
        const dateMatches = text.match(/\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}/g);
        if (dateMatches && dateMatches.length > 0) {
          fallbackRow.date = dateMatches[0];
        }
        
        // Try to find amounts
        const amountMatches = text.match(/\d+[.,]\d{2}|\d+/g);
        if (amountMatches && amountMatches.length > 0) {
          fallbackRow.amount = amountMatches[amountMatches.length - 1];
        }
        
        const processed = processData([fallbackRow], billingType);
        return processed;
      }
    } else {
      // .doc files (older format) - limited support
      // Try to read as binary and extract text if possible
      throw new Error('DOC files (older format) have limited support. Please convert to DOCX or use PDF/CSV/XLSX format for better results.');
    }
  } catch (error) {
    throw new Error(`Error parsing DOC/DOCX: ${error.message}`);
  }
};

const removeDuplicates = (data, identifierColumn) => {
  if (!identifierColumn) return data;

  const seen = new Set();
  const unique = [];

  for (const row of data) {
    const identifier = String(row[identifierColumn] || '').trim();
    if (identifier && !seen.has(identifier)) {
      seen.add(identifier);
      unique.push(row);
    }
  }

  return unique;
};

