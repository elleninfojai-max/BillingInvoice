import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  AlignmentType, 
  HeadingLevel,
  BorderStyle,
  WidthType,
  Table,
  TableRow,
  TableCell,
  VerticalAlign,
  ImageRun,
  Media
} from 'docx';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';
import { generateDOCXBlobDebit } from './exportUtilsDebit';
import { exportToPDF, generatePDFBlob } from './exportUtilsPDF';

export const exportToCSV = (data, filename = 'invoice-data.csv') => {
  if (!data || data.length === 0) {
    alert('No data to export');
    return;
  }

  const csv = Papa.unparse(data, {
    header: true,
  });

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  saveAs(blob, filename);
};

export const exportToXLSX = (data, filename = 'invoice-data.xlsx') => {
  if (!data || data.length === 0) {
    alert('No data to export');
    return;
  }

  try {
    // Ensure data is properly formatted
    const formattedData = data.map(row => {
      const formattedRow = {};
      Object.keys(row).forEach(key => {
        // Preserve the exact value as it appears in the data
        formattedRow[key] = row[key] !== null && row[key] !== undefined ? row[key] : '';
      });
      return formattedRow;
    });

    const ws = XLSX.utils.json_to_sheet(formattedData);
    
    // Set column widths for better readability
    const colWidths = Object.keys(formattedData[0] || {}).map(key => ({ wch: Math.max(key.length, 15) }));
    ws['!cols'] = colWidths;
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, filename);
  } catch (error) {
    console.error('Error exporting to XLSX:', error);
    alert('Error exporting to Excel. Please try again.');
  }
};

// Helper function to convert number to words (Indian format)
const convertToWords = (num) => {
  if (num === 0) return '';
  if (num < 10) {
    const ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
    return ones[num];
  }
  if (num < 20) {
    const teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
    return teens[num - 10];
  }
  if (num < 100) {
    const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
    const tensPlace = Math.floor(num / 10);
    const onesPlace = num % 10;
    return tens[tensPlace] + (onesPlace > 0 ? ' ' + convertToWords(onesPlace) : '');
  }
  return '';
};

const numberToWords = (num) => {
  if (!num || num === 0) return 'Rupee Zero Only';
  
  const crores = Math.floor(num / 10000000);
  const lakhs = Math.floor((num % 10000000) / 100000);
  const thousands = Math.floor((num % 100000) / 1000);
  const hundreds = Math.floor((num % 1000) / 100);
  const remainder = num % 100;
  
  let words = [];
  
  if (crores > 0) words.push(`${convertToWords(crores)} Crore`);
  if (lakhs > 0) words.push(`${convertToWords(lakhs)} Lakh`);
  if (thousands > 0) words.push(`${convertToWords(thousands)} Thousand`);
  if (hundreds > 0) words.push(`${convertToWords(hundreds)} Hundred`);
  if (remainder > 0) words.push(convertToWords(remainder));
  
  // Add "Rupee" prefix and proper formatting
  const amountWords = words.length > 0 ? words.join(' ') : 'Zero';
  return `Rupee ${amountWords} Only`;
};

// Helper to format date - return exactly as it appears in Excel/data without any conversion
const formatDate = (dateStr) => {
  if (!dateStr || dateStr === null || dateStr === undefined) {
    // Only return default if completely empty
    return '-';
  }
  
  // Return date exactly as it appears in the source data (Excel format preserved)
  return String(dateStr);
};

// Helper to get amount value (prioritizes credit for invoices, then debit)
const getAmount = (row, amountKeys = ['amount', 'total', 'amt', 'sum', 'price', 'credit', 'debit', 'balance']) => {
  // First try to find credit amount (for invoice generation)
  for (const key of Object.keys(row)) {
    const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (normalizedKey.includes('credit') || normalizedKey.includes('cr')) {
      const value = row[key];
      if (value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
        const amt = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
        if (amt > 0) return amt;
      }
    }
  }
  
  // Then try debit amount
  for (const key of Object.keys(row)) {
    const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (normalizedKey.includes('debit') || normalizedKey.includes('dr')) {
      const value = row[key];
      if (value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
        const amt = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
        if (amt > 0) return amt;
      }
    }
  }
  
  // If no credit or debit, try general amount columns
  for (const key of Object.keys(row)) {
    const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (amountKeys.some(amtKey => normalizedKey.includes(amtKey)) && 
        !normalizedKey.includes('debit') && !normalizedKey.includes('dr') &&
        !normalizedKey.includes('credit') && !normalizedKey.includes('cr')) {
      const value = row[key];
      if (value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
        return Math.abs(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')));
      }
    }
  }
  return 0;
};

// Helper to check if transaction is credit
const isCreditTransaction = (row) => {
  let creditValue = 0;
  let debitValue = 0;
  
  for (const key of Object.keys(row)) {
    const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
    const value = row[key];
    
    if (normalizedKey.includes('credit') || normalizedKey.includes('cr')) {
      if (value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
        creditValue = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
      }
    } else if (normalizedKey.includes('debit') || normalizedKey.includes('dr')) {
      if (value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
        debitValue = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
      }
    }
  }
  
  // If credit has value, it's a credit transaction
  if (creditValue > 0) return true;
  // If debit has value and credit doesn't, it's a debit transaction
  if (debitValue > 0 && creditValue === 0) return false;
  // Default to credit if both are 0 (for invoice generation)
  return true;
};

// Extract last part from serial number (e.g., "001" from "ELLEN/UPI/2025/001")
const extractSerialNumberSequence = (serialNumber) => {
  if (!serialNumber) return null;
  const serialStr = String(serialNumber).trim();
  
  // Extract last part after last "/"
  const parts = serialStr.split('/');
  if (parts.length > 0) {
    const lastPart = parts[parts.length - 1].trim();
    // If it's just numbers, return padded to 3 digits
    if (/^\d+$/.test(lastPart)) {
      return lastPart.padStart(3, '0');
    }
  }
  
  // Fallback: try to extract last sequence of digits
  const match = serialStr.match(/(\d+)$/);
  if (match && match[1]) {
    return match[1].padStart(3, '0');
  }
  
  return null;
};

// Generate invoice number based on sequential numbering (format: ELLEN/UPI/2025/001)
// If serialNumber is provided, extract sequence from it; otherwise use counter
const generateInvoiceNumber = (invoiceCounter, serialNumber = null) => {
  const currentYear = new Date().getFullYear();
  
  // If serial number is provided, extract sequence from it
  if (serialNumber) {
    const extractedSequence = extractSerialNumberSequence(serialNumber);
    if (extractedSequence) {
      return `ELLEN/UPI/${currentYear}/${extractedSequence}`;
    }
  }
  
  // Fallback to counter-based generation
  if (!invoiceCounter.count) {
    invoiceCounter.count = 0;
  }
  invoiceCounter.count++;
  
  const sequence = String(invoiceCounter.count).padStart(3, '0');
  return `ELLEN/UPI/${currentYear}/${sequence}`;
};

// Detect payment mode (UPI, Net Banking, Cash on Delivery)
const detectPaymentMode = (description, refNo) => {
  if (!description) description = '';
  if (!refNo) refNo = '';
  
  const descLower = String(description).toLowerCase();
  const refLower = String(refNo).toLowerCase();
  
  // Check for UPI
  if (descLower.includes('upi') || descLower.includes('upi/') || refLower.includes('upi') || 
      descLower.match(/@\w+/) || descLower.match(/[a-z0-9]+@[a-z]+/)) {
    return 'UPI';
  }
  
  // Check for Net Banking (NEFT, RTGS, IMPS)
  if (descLower.includes('neft') || descLower.includes('rtgs') || descLower.includes('imps') ||
      descLower.includes('net banking') || descLower.includes('online transfer') ||
      refLower.includes('neft') || refLower.includes('rtgs') || refLower.includes('imps')) {
    return 'Net Banking';
  }
  
  // Check for Cash on Delivery
  if (descLower.includes('cod') || descLower.includes('cash on delivery') ||
      descLower.includes('cash delivery')) {
    return 'Cash on Delivery';
  }
  
  // Default to UPI if no match
  return 'UPI';
};

// Generate voucher number for debit transactions (format: ELLEN/PV/2025/VIGNESH001)
const generateVoucherNumber = (payeeName, index, payeeNameMap = {}) => {
  if (!payeeName || payeeName === 'N/A') {
    const sequence = String(index + 1).padStart(3, '0');
    const currentYear = new Date().getFullYear();
    return `ELLEN/PV/${currentYear}/PAYEE${sequence}`;
  }
  
  // Extract payee name prefix (first word)
  let namePrefix = payeeName
    .split('/')[0]  // Take part before "/" if exists
    .trim()
    .split(' ')[0]  // Take first word
    .replace(/[^A-Za-z0-9]/g, '')  // Remove special characters
    .toUpperCase()
    .substring(0, 10);  // Max 10 characters
  
  if (!namePrefix || namePrefix.length === 0) {
    namePrefix = 'PAYEE';
  }
  
  // Track sequence number for each payee
  if (!payeeNameMap[namePrefix]) {
    payeeNameMap[namePrefix] = 0;
  }
  payeeNameMap[namePrefix]++;
  
  const sequence = String(payeeNameMap[namePrefix]).padStart(3, '0');
  const currentYear = new Date().getFullYear();
  return `ELLEN/PV/${currentYear}/${namePrefix}${sequence}`;
};

export const exportToDOCX = async (data, billingType, transactionType = 'all', filename = 'invoice-report.docx') => {
  if (!data || data.length === 0) {
    alert('No data to export');
    return;
  }

  // Company details
  const companyName = 'ELLEN INFORMATION TECHNOLOGY SOLUTIONS PRIVATE LIMITED';
  const companyGSTIN = '33AAHCE0984H1ZN';
  const companyAddress = 'Registered Office : 8/3, Athreyapuram 2nd Street, Choolaimedu, Chennai – 600094';
  const companyEmail = 'Email : support@learnsconnect.com';
  const companyPhone = 'Phone : +91 8489357705';

  // Invoice counter for sequential numbering
  const invoiceCounter = { count: 0 };
  const payeeNameMap = {};

  // Batch processing: 50 invoices per file for better performance
  const BATCH_SIZE = 50;
  const batches = Math.ceil(data.length / BATCH_SIZE);

  // Load logo image once (for all receipts)
  let logoData = null;
  try {
    // Try to load logo from assets folder
    const logoResponse = await fetch('/src/assets/images/logo.jpg');
    if (logoResponse.ok) {
      logoData = await logoResponse.arrayBuffer();
    }
  } catch (error) {
    // Try alternative paths
    try {
      const altResponse = await fetch('/assets/images/logo.jpg');
      if (altResponse.ok) {
        logoData = await altResponse.arrayBuffer();
      } else {
        // Try public folder
        const pubResponse = await fetch('/logo.jpg');
        if (pubResponse.ok) {
          logoData = await pubResponse.arrayBuffer();
        }
      }
    } catch (err) {
      console.warn('Logo image not found, continuing without logo');
    }
  }

  for (let batchIndex = 0; batchIndex < batches; batchIndex++) {
    const startIndex = batchIndex * BATCH_SIZE;
    const endIndex = Math.min(startIndex + BATCH_SIZE, data.length);
    const batchData = data.slice(startIndex, endIndex);

    const children = [];

    // Process each record as individual invoice/voucher
    for (let i = 0; i < batchData.length; i++) {
      const row = batchData[i];
      
      // Debug: Log row data to help identify issues
      console.log('Processing row:', row);
      
      const amount = getAmount(row);
      const amountInWords = numberToWords(Math.floor(amount));
      
      // Get serial number from row data
      const serialNoKeys = ['serialno', 'serialnumber', 'serial no', 's.no', 'sno', 'serialno'];
      const serialNoKey = Object.keys(row).find(key => 
        serialNoKeys.some(sk => key.toLowerCase().replace(/[^a-z0-9]/g, '').includes(sk.toLowerCase().replace(/[^a-z0-9]/g, '')))
      );
      const serialNumber = serialNoKey ? String(row[serialNoKey] || '').trim() : '';

      // Get date from various possible column names (try multiple date columns)
      const dateKeys = ['date', 'txndate', 'transactiondate', 'paymentdate', 'valuedate', 'transaction date', 'value date'];
      let dateKey = Object.keys(row).find(key => 
        dateKeys.some(dk => key.toLowerCase().includes(dk.toLowerCase()))
      );
      
      // If no date key found, try to find any column that looks like a date (contains date pattern)
      if (!dateKey) {
        for (const key of Object.keys(row)) {
          const value = String(row[key] || '').trim();
          // Check if value looks like a date (e.g., "3/11/25", "11/3/25", etc.)
          if (value.match(/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/)) {
            dateKey = key;
            break;
          }
        }
      }
      
      const invoiceDate = formatDate(row[dateKey] || null);

      // Get description to extract client/payer and payment details
      // Try multiple possible column names for description
      const descKeys = ['description', 'particulars', 'narration', 'details', 'remarks', 'note', 'transaction details', 'narration'];
      let descKey = Object.keys(row).find(key => 
        descKeys.some(dk => key.toLowerCase().includes(dk.toLowerCase()))
      );
      
      // If no description key found, try to find column with longest text content (likely description)
      if (!descKey) {
        let maxLength = 0;
        for (const key of Object.keys(row)) {
          const value = String(row[key] || '').trim();
          // Skip if it's a date, number, serial number, or empty
          if (value && 
              !value.match(/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/) && 
              !value.match(/^\-?\d+[.,]?\d*$/) && 
              !value.match(/^[A-Z]+\/\d+\/\d+\/\d+$/) &&
              !value.match(/^[\/\-\s]+$/)) { // Skip if it's just separators like "- / -"
            if (value.length > maxLength && value.length > 10) {
              maxLength = value.length;
              descKey = key;
            }
          }
        }
      }
      
      // Get description - preserve full text including special characters
      const description = descKey ? String(row[descKey] || '').trim() : '';
      
      // Debug: Log extracted values (only in development)
      if (process.env.NODE_ENV === 'development') {
        console.log('Extracted values:', {
          dateKey,
          invoiceDate,
          descKey,
          description,
          amount,
          serialNumber
        });
      }

      // Extract client/payer name from description
      // Pattern: "Tacklerz Innovations / tgpradz3@ybl" or similar
      let partyName = 'N/A';
      
      // Try direct name fields first
      const nameKeys = ['name', 'partyname', 'client', 'payer', 'studentname', 'payeename', 'vendorname'];
      const nameKey = Object.keys(row).find(key => 
        nameKeys.some(nk => key.toLowerCase().includes(nk.toLowerCase()))
      );
      
      if (nameKey && row[nameKey] && String(row[nameKey]).trim().length > 0) {
        partyName = String(row[nameKey]).trim();
      } else if (description) {
        // Extract from description - look for patterns like "Company Name / email" or "Company Name"
        const clientPatterns = [
          /([A-Z][A-Za-z\s&]+(?:\s+\/[A-Za-z0-9@.]+)?)/,  // "Company / email"
          /(?:FROM|TO|BY)\s+([A-Z][A-Za-z\s]+)/,  // "FROM Company" or "TO Company"
          /(\d+\s*\/\s*[A-Z][A-Za-z\s]+)/  // "123 / Company"
        ];
        
        for (const pattern of clientPatterns) {
          const match = description.match(pattern);
          if (match && match[1]) {
            partyName = match[1].trim();
            break;
          }
        }
        
        // If still not found, extract first meaningful words
        if (partyName === 'N/A') {
          const words = description.split(/[\/\s]+/).filter(w => w.length > 2 && !/^\d+$/.test(w));
          if (words.length > 0) {
            partyName = words.slice(0, 3).join(' ').trim();
          }
        }
      }

      // Determine if this is credit (Invoice) or debit (Payment Voucher)
      const isCredit = isCreditTransaction(row);
      
      // Generate invoice/voucher number (only increment counter for credit transactions)
      // Use serial number if available to extract sequence
      const invoiceNumber = isCredit 
        ? generateInvoiceNumber(invoiceCounter, serialNumber)
        : generateVoucherNumber(partyName, startIndex + i, payeeNameMap);

      // Get Ref No from various columns
      const refKeys = ['refno', 'refno', 'ref', 'chequeno', 'utrno', 'utr', 'branchcode'];
      const refKey = Object.keys(row).find(key => 
        refKeys.some(rk => key.toLowerCase().includes(rk.toLowerCase()))
      );
      let refNo = refKey ? String(row[refKey] || '') : '';
      
      // Detect payment mode (UPI, Net Banking, Cash on Delivery)
      const paymentMode = detectPaymentMode(description, refNo);

      if (isCredit) {
        // PAYMENT RECEIPT FORMAT (Credit/Student Fee Collection) - Exact template matching image
        // Get description for "Received from Payer (Name and Details)"
        const payerDetails = description || partyName || '-';
        
        children.push(
          // Page break before for new receipts (except first)
          new Paragraph({
            children: [new TextRun({ text: '' })],
            pageBreakBefore: i > 0 || batchIndex > 0,
            spacing: { after: 0, before: 0 },
          }),
          
          // Header Section with Logo and Contact Info (using Table for side-by-side layout, no borders)
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: {
                size: 0,
                style: BorderStyle.NONE,
              },
              bottom: {
                size: 0,
                style: BorderStyle.NONE,
              },
              left: {
                size: 0,
                style: BorderStyle.NONE,
              },
              right: {
                size: 0,
                style: BorderStyle.NONE,
              },
              insideHorizontal: {
                size: 0,
                style: BorderStyle.NONE,
              },
              insideVertical: {
                size: 0,
                style: BorderStyle.NONE,
              },
            },
            rows: [
              new TableRow({
                children: [
                  // Left Column: Logo and LEARNSCONNECT
                  new TableCell({
                    borders: {
                      top: { size: 0, style: BorderStyle.NONE },
                      bottom: { size: 0, style: BorderStyle.NONE },
                      left: { size: 0, style: BorderStyle.NONE },
                      right: { size: 0, style: BorderStyle.NONE },
                    },
                    children: [
                      // Logo Image at the top
                      ...(logoData ? [
                        new Paragraph({
                          children: [
                            new ImageRun({
                              data: new Uint8Array(logoData),
                              transformation: {
                                width: 80,
                                height: 80,
                              },
                            }),
                          ],
                          alignment: AlignmentType.LEFT,
                          spacing: { after: 80, before: 0 },
                        }),
                      ] : []),
                      // LEARNSCONNECT Branding
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: 'LEARNSCONNECT',
                            bold: true,
                            size: 28,
                            color: '0000FF',
                          }),
                        ],
                        alignment: AlignmentType.LEFT,
                        spacing: { after: 30 },
                      }),
                      // Tagline
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: 'WHERE LEARNING MEETS OPPORTUNITY',
                            size: 20,
                            color: '333333',
                            bold: true,
                            font: 'Calibri',
                          }),
                        ],
                        alignment: AlignmentType.LEFT,
                        spacing: { after: 0 },
                      }),
                    ],
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.TOP,
                  }),
                  // Right Column: Contact Information
                  new TableCell({
                    borders: {
                      top: { size: 0, style: BorderStyle.NONE },
                      bottom: { size: 0, style: BorderStyle.NONE },
                      left: { size: 0, style: BorderStyle.NONE },
                      right: { size: 0, style: BorderStyle.NONE },
                    },
                    children: [
                      // Address
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: '8/3, Athreyapuram 2nd Street, Choolaimedu, Chennai - 600094',
                            size: 22,
                            font: 'Calibri',
                          }),
                        ],
                        alignment: AlignmentType.RIGHT,
                        spacing: { after: 60, before: 0 },
                      }),
                      // Phone
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: '📞 +91 84893 57705',
                            size: 22,
                            font: 'Calibri',
                          }),
                        ],
                        alignment: AlignmentType.RIGHT,
                        spacing: { after: 60 },
                      }),
                      // Email
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: '✉ support@learnsconnect.com',
                            size: 22,
                            font: 'Calibri',
                          }),
                        ],
                        alignment: AlignmentType.RIGHT,
                        spacing: { after: 60 },
                      }),
                      // Website
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: '🌐 www.learnsconnect.com',
                            size: 22,
                            font: 'Calibri',
                          }),
                        ],
                        alignment: AlignmentType.RIGHT,
                        spacing: { after: 0 },
                      }),
                    ],
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.TOP,
                  }),
                ],
              }),
            ],
          }),
          
          // Spacing after header table (compact)
          new Paragraph({
            children: [new TextRun({ text: '' })],
            spacing: { after: 150 },
          }),

          // Company Name - Centered, Bold, Uppercase
          new Paragraph({
            children: [
              new TextRun({
                text: 'ELLEN INFORMATION TECHNOLOGY SOLUTIONS PRIVATE LIMITED',
                bold: true,
                size: 22,
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
          }),
          
          // Teal Separator Line (using underline or border effect)
          new Paragraph({
            children: [
              new TextRun({
                text: '________________________________________________________________________________',
                size: 22,
                color: '008080', // Teal color
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
          }),

          // Payment Receipt Title - "LearnsConnect – Payment Receipt" (centered, bold, larger font)
          new Paragraph({
            children: [
              new TextRun({
                text: 'LearnsConnect – Payment Receipt',
                bold: true,
                size: 32,
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 600 },
          }),

          // Payment Receipt Content - Table format for proper alignment
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { size: 0, style: BorderStyle.NONE },
              bottom: { size: 0, style: BorderStyle.NONE },
              left: { size: 0, style: BorderStyle.NONE },
              right: { size: 0, style: BorderStyle.NONE },
              insideHorizontal: { size: 0, style: BorderStyle.NONE },
              insideVertical: { size: 0, style: BorderStyle.NONE },
            },
            rows: [
              // Payment Receipt Voucher No.
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: 'Payment Receipt Voucher No.:',
                            bold: true,
                            size: 22,
                          }),
                        ],
                        spacing: { after: 200 },
                      }),
                    ],
                    width: { size: 40, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.TOP,
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: invoiceNumber || '-',
                            size: 22,
                          }),
                        ],
                        spacing: { after: 200 },
                      }),
                    ],
                    width: { size: 60, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.TOP,
                  }),
                ],
              }),
              // Date
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: 'Date:',
                            bold: true,
                            size: 22,
                          }),
                        ],
                        spacing: { after: 200 },
                      }),
                    ],
                    width: { size: 40, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.TOP,
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: invoiceDate || '-',
                            size: 22,
                          }),
                        ],
                        spacing: { after: 200 },
                      }),
                    ],
                    width: { size: 60, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.TOP,
                  }),
                ],
              }),
              // Mode of Payment
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: 'Mode of Payment:',
                            bold: true,
                            size: 22,
                          }),
                        ],
                        spacing: { after: 200 },
                      }),
                    ],
                    width: { size: 40, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.TOP,
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: 'UPI',
                            size: 22,
                          }),
                        ],
                        spacing: { after: 200 },
                      }),
                    ],
                    width: { size: 60, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.TOP,
                  }),
                ],
              }),
              // Received from Payer (Name and Details)
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: 'Received from Payer (Name and Details):',
                            bold: true,
                            size: 22,
                          }),
                        ],
                        spacing: { after: 200 },
                      }),
                    ],
                    width: { size: 40, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.TOP,
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: payerDetails,
                            size: 22,
                          }),
                        ],
                        spacing: { after: 200 },
                      }),
                    ],
                    width: { size: 60, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.TOP,
                  }),
                ],
              }),
              // Purpose of Payment
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: 'Purpose of Payment:',
                            bold: true,
                            size: 22,
                          }),
                        ],
                        spacing: { after: 200 },
                      }),
                    ],
                    width: { size: 40, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.TOP,
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: 'Payment collected from students for Course/Batch',
                            size: 22,
                          }),
                        ],
                        spacing: { after: 200 },
                      }),
                    ],
                    width: { size: 60, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.TOP,
                  }),
                ],
              }),
              // Amount (₹)
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: 'Amount (₹):',
                            bold: true,
                            size: 22,
                          }),
                        ],
                        spacing: { after: 400 },
                      }),
                    ],
                    width: { size: 40, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.TOP,
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: `₹ ${amount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
                            size: 22,
                            bold: true,
                          }),
                        ],
                        spacing: { after: 400 },
                      }),
                    ],
                    width: { size: 60, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.TOP,
                  }),
                ],
              }),
              // Amount in Words
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: 'Amount in Words:',
                            bold: true,
                            size: 22,
                          }),
                        ],
                        spacing: { after: 800 },
                      }),
                    ],
                    width: { size: 40, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.TOP,
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: amountInWords,
                            size: 22,
                          }),
                        ],
                        spacing: { after: 800 },
                      }),
                    ],
                    width: { size: 60, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.TOP,
                  }),
                ],
              }),
            ],
          }),

          // Spacer to push footer 2.5 inches + 3 pixels down (3600 + 45 = 3645 twips)
          new Paragraph({
            children: [new TextRun({ text: '' })],
            spacing: { before: 3645, after: 0 }, // 2.5 inches + 3 pixels = 3645 twips
          }),

          // Footer Disclaimer (matching image exactly)
          new Paragraph({
            children: [
              new TextRun({
                text: 'This receipt is computer generated by for acknowledgement of payment received. Non-refundable unless of technical issues. All disputes subject to Chennai jurisdiction.',
                size: 18,
                italics: true,
                color: '808080', // Gray color
              }),
            ],
            alignment: AlignmentType.LEFT,
            spacing: { before: 0, after: 200 },
          }),
          
          // Bottom Teal Border Line (matching image)
          new Paragraph({
            children: [
              new TextRun({
                text: '________________________________________________________________________________',
                size: 22,
                color: '008080', // Teal color
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
          }),
        );
        
        // Add page break after receipt (except for last receipt in last batch)
        if (!(i === batchData.length - 1 && batchIndex === batches - 1)) {
          children.push(
            new Paragraph({
              children: [new TextRun({ text: '', break: 1 })],
            })
          );
        }
      } else {
        // PAYMENT VOUCHER FORMAT (Debit/Fees Paid to Tutors)
        children.push(
          // Company Header (with page break before for new vouchers)
          new Paragraph({
            children: [
              new TextRun({
                text: companyName,
                bold: true,
                size: 28,
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
            pageBreakBefore: i > 0 || batchIndex > 0, // Start new page for each voucher except the first one
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: companyAddress,
                size: 22,
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 100 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: companyEmail,
                size: 22,
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
          }),

          // Title: Payment Voucher / Supplier Invoice
          new Paragraph({
            children: [
              new TextRun({
                text: 'PAYMENT VOUCHER / SUPPLIER INVOICE',
                bold: true,
                size: 32,
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
          }),

          // Voucher Details - Table format
          new Table({
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Voucher No.', bold: true, size: 22 })] })],
                    width: { size: 40, type: WidthType.PERCENTAGE },
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: invoiceNumber, size: 22 })] })],
                    width: { size: 60, type: WidthType.PERCENTAGE },
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Date', bold: true, size: 22 })] })],
                    width: { size: 40, type: WidthType.PERCENTAGE },
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: invoiceDate || '-', size: 22 })] })],
                    width: { size: 60, type: WidthType.PERCENTAGE },
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Mode of Payment', bold: true, size: 22 })] })],
                    width: { size: 40, type: WidthType.PERCENTAGE },
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: paymentMode || '-', size: 22 })] })],
                    width: { size: 60, type: WidthType.PERCENTAGE },
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Payee Name', bold: true, size: 22 })] })],
                    width: { size: 40, type: WidthType.PERCENTAGE },
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: partyName || '-', size: 22 })] })],
                    width: { size: 60, type: WidthType.PERCENTAGE },
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Purpose of Payment', bold: true, size: 22 })] })],
                    width: { size: 40, type: WidthType.PERCENTAGE },
                  }),
                  new TableCell({
                    children: [new Paragraph({ 
                      children: [new TextRun({ 
                        text: description || 'Advance Payment to Dealer / Vendor for Procurement of Learning Kit Materials', 
                        size: 22 
                      })] 
                    })],
                    width: { size: 60, type: WidthType.PERCENTAGE },
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Payment Type', bold: true, size: 22 })] })],
                    width: { size: 40, type: WidthType.PERCENTAGE },
                  }),
                  new TableCell({
                    children: [new Paragraph({ 
                      children: [new TextRun({ 
                        text: 'Advance to Supplier (Deductible against final invoice)', 
                        size: 22 
                      })] 
                    })],
                    width: { size: 60, type: WidthType.PERCENTAGE },
                  }),
                ],
              }),
            ],
            width: { size: 100, type: WidthType.PERCENTAGE },
          }),

          new Paragraph({ spacing: { after: 400 } }),

          // Break-up of Amount
          new Paragraph({
            children: [
              new TextRun({
                text: 'Break-up of Amount:',
                bold: true,
                size: 22,
              }),
            ],
            spacing: { after: 200 },
          }),
          
          // Amount Table
          new Table({
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Particulars', bold: true, size: 22 })] })],
                    width: { size: 70, type: WidthType.PERCENTAGE },
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Amount (₹)', bold: true, size: 22 })] })],
                    width: { size: 30, type: WidthType.PERCENTAGE },
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Material / Service Advance', size: 22 })] })],
                    width: { size: 70, type: WidthType.PERCENTAGE },
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: amount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }), size: 22 })] })],
                    width: { size: 30, type: WidthType.PERCENTAGE },
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Total', bold: true, size: 22 })] })],
                    width: { size: 70, type: WidthType.PERCENTAGE },
                  }),
                  new TableCell({
                    children: [new Paragraph({ 
                      children: [new TextRun({ 
                        text: `₹ ${amount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} Only`, 
                        size: 22,
                        bold: true
                      })] 
                    })],
                    width: { size: 30, type: WidthType.PERCENTAGE },
                  }),
                ],
              }),
            ],
            width: { size: 100, type: WidthType.PERCENTAGE },
          }),

          new Paragraph({ spacing: { after: 400 } }),

          // Amount in Words
          new Paragraph({
            children: [
              new TextRun({
                text: `Amount in Words: ${amountInWords}`,
                size: 22,
              }),
            ],
            spacing: { after: 800 },
          }),

          // Footer Disclaimer
          new Paragraph({
            children: [
              new TextRun({
                text: 'All payments are made via SBI RTGS/NEFT from Ellen Information Technology Solutions Pvt. Ltd. (A/c No. ending 5037) towards business services rendered under the LearnsConnect operations. Each transfer is supported by work orders, GST invoices, and digital approval records.',
                size: 18,
                italics: true,
              }),
            ],
            alignment: AlignmentType.LEFT,
            spacing: { before: 400, after: 400 },
          }),
          // Page break after voucher (ensures next voucher/invoice starts on new page)
          new Paragraph({
            children: [new TextRun({ text: '', break: 1 })],
          }),
        );
      }
    }

    // Ensure we have content before creating document
    if (children.length === 0) {
      console.warn('No content to export');
      return;
    }

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: {
              top: 720, // 0.5 inch in twips (reduced from 1 inch for no top space)
              right: 1440,
              bottom: 1440,
              left: 1440,
            },
          },
        },
        children,
      }],
    });

    const blob = await Packer.toBlob(doc);
    const batchFilename = batches > 1 
      ? filename.replace('.docx', `_part${batchIndex + 1}.docx`)
      : filename;
    saveAs(blob, batchFilename);
  }
};

// Export mixed credit and debit transactions - generates credit invoices and debit vouchers in one document
export const exportToDOCXMixed = async (data, billingType, filename = 'invoice-mixed-report.docx') => {
  if (!data || data.length === 0) {
    alert('No data to export');
    return;
  }

  // Separate credit and debit rows - improved detection
  const creditRows = [];
  const debitRows = [];
  
  for (const row of data) {
    let hasCredit = false;
    let hasDebit = false;
    
    // Check for credit value
    for (const key of Object.keys(row)) {
      const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
      const value = row[key];
      
      if ((normalizedKey.includes('credit') || normalizedKey.includes('cr')) && 
          value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
        const amt = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
        if (amt > 0) {
          hasCredit = true;
          break;
        }
      }
    }
    
    // Check for debit value
    for (const key of Object.keys(row)) {
      const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
      const value = row[key];
      
      if ((normalizedKey.includes('debit') || normalizedKey.includes('dr')) && 
          value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
        const amt = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
        if (amt > 0) {
          hasDebit = true;
          break;
        }
      }
    }
    
    // Classify row: if has credit and no debit = credit, if has debit and no credit = debit
    if (hasCredit && !hasDebit) {
      creditRows.push(row);
    } else if (hasDebit && !hasCredit) {
      debitRows.push(row);
    }
    // If both or neither, skip or classify based on amount comparison (skip for now to be safe)
  }
  
  console.log(`Exporting mixed: ${creditRows.length} credit rows, ${debitRows.length} debit rows`);

  const companyName = 'ELLEN INFORMATION TECHNOLOGY SOLUTIONS PRIVATE LIMITED';
  const companyGSTIN = '33AAHCE0984H1ZN';
  const companyAddress = 'Registered Office : 8/3, Athreyapuram 2nd Street, Choolaimedu, Chennai – 600094';
  const companyEmail = 'Email : support@learnsconnect.com';
  const companyPhone = 'Phone : +91 8489357705';

  const invoiceCounter = { count: 0 };
  const voucherCounter = { count: 0 };
  const payeeNameMap = {};
  const children = [];

  // Generate credit invoices first
  for (let i = 0; i < creditRows.length; i++) {
    const row = creditRows[i];
    const amount = getAmount(row);
    const amountInWords = numberToWords(Math.floor(amount));
    
    // Get serial number from row data
    const serialNoKeys = ['serialno', 'serialnumber', 'serial no', 's.no', 'sno'];
    const serialNoKey = Object.keys(row).find(key => 
      serialNoKeys.some(sk => key.toLowerCase().replace(/[^a-z0-9]/g, '').includes(sk.toLowerCase().replace(/[^a-z0-9]/g, '')))
    );
    const serialNumber = serialNoKey ? String(row[serialNoKey] || '') : '';
    
    const dateKeys = ['date', 'txndate', 'transactiondate', 'paymentdate', 'valuedate'];
    const dateKey = Object.keys(row).find(key => 
      dateKeys.some(dk => key.toLowerCase().includes(dk.toLowerCase()))
    );
    const invoiceDate = formatDate(row[dateKey] || new Date());

    const descKey = Object.keys(row).find(key => 
      key.toLowerCase().includes('description') || 
      key.toLowerCase().includes('particulars') ||
      key.toLowerCase().includes('narration')
    );
    const description = descKey ? String(row[descKey] || '') : '';

    let partyName = 'N/A';
    const nameKeys = ['name', 'partyname', 'client', 'payer', 'studentname', 'payeename', 'vendorname'];
    const nameKey = Object.keys(row).find(key => 
      nameKeys.some(nk => key.toLowerCase().includes(nk.toLowerCase()))
    );
    
    if (nameKey && row[nameKey]) {
      partyName = String(row[nameKey]).trim();
    } else if (description) {
      const nameMatch = description.match(/^([^\/]+?)(?:\s*\/|\s*@|$)/);
      if (nameMatch && nameMatch.length > 0) {
        partyName = nameMatch[1].trim();
      }
    }

    const invoiceNumber = generateInvoiceNumber(invoiceCounter, serialNumber);
    const paymentMode = detectPaymentMode(description, row);

    // Get Ref No
    const refKeys = ['refno', 'ref', 'referenceno', 'referencenumber', 'chequeno', 'cheque'];
    const refKey = Object.keys(row).find(key => 
      refKeys.some(rk => key.toLowerCase().includes(rk.toLowerCase()))
    );
    const refNo = refKey ? String(row[refKey] || '') : '';

    // Credit Invoice Content
    if (i === 0) {
      // First credit invoice - add page break before header
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: companyName,
              bold: true,
              size: 28,
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: { after: 200 },
        }),
      );
    } else {
      // Subsequent credit invoices - add page break
      children.push(
        new Paragraph({
          pageBreakBefore: true,
          children: [
            new TextRun({
              text: companyName,
              bold: true,
              size: 28,
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: { after: 200 },
        }),
      );
    }

    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `GSTIN : ${companyGSTIN}`,
            size: 22,
          }),
        ],
        alignment: AlignmentType.LEFT,
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: companyAddress,
            size: 22,
          }),
        ],
        alignment: AlignmentType.LEFT,
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: companyEmail,
            size: 22,
          }),
        ],
        alignment: AlignmentType.LEFT,
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: companyPhone,
            size: 22,
          }),
        ],
        alignment: AlignmentType.LEFT,
        spacing: { after: 600 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'PAYMENT RECEIPT',
            bold: true,
            size: 32,
          }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 600 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'Payment Receipt Voucher No.: ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: invoiceNumber,
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'Date: ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: invoiceDate || '-',
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'Mode of Payment: ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: 'UPI',
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'Received from Payer (Name and Details): ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: description || partyName || '-',
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'Purpose of Payment: ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: 'Payment collected from students for Course/Batch',
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'Amount (₹): ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: `₹ ${amount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
            size: 22,
            bold: true,
          }),
        ],
        spacing: { after: 400 },
      }),
      new Paragraph({
        spacing: { after: 400 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: `Amount in Words: ${amountInWords}`,
            size: 22,
          }),
        ],
        spacing: { after: 800 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'All payments are made via SBI RTGS/NEFT from Ellen Information Technology Solutions Pvt. Ltd. (A/c No. ending 5037) towards business services rendered under the LearnsConnect operations. Each transfer is supported by work orders, GST invoices, and digital approval records.',
            size: 20,
            italics: true,
          }),
        ],
        spacing: { after: 800 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: '',
            break: 1,
          }),
        ],
      }),
    );
  }

  // Generate debit vouchers inline
  for (let i = 0; i < debitRows.length; i++) {
    const row = debitRows[i];
    
    // Get serial number from row data
    const serialNoKeys = ['serialno', 'serialnumber', 'serial no', 's.no', 'sno'];
    const serialNoKey = Object.keys(row).find(key => 
      serialNoKeys.some(sk => key.toLowerCase().replace(/[^a-z0-9]/g, '').includes(sk.toLowerCase().replace(/[^a-z0-9]/g, '')))
    );
    const serialNumber = serialNoKey ? String(row[serialNoKey] || '') : '';
    
    const debitAmount = getDebitAmount(row);
    const amountInWords = numberToWords(Math.floor(debitAmount));
    
    const dateKeys = ['date', 'txndate', 'transactiondate', 'paymentdate', 'valuedate'];
    const dateKey = Object.keys(row).find(key => 
      dateKeys.some(dk => key.toLowerCase().includes(dk.toLowerCase()))
    );
    const invoiceDate = formatDate(row[dateKey] || new Date());

    const descKey = Object.keys(row).find(key => 
      key.toLowerCase().includes('description') || 
      key.toLowerCase().includes('particulars') ||
      key.toLowerCase().includes('narration')
    );
    const description = descKey ? String(row[descKey] || '') : '';

    // Extract payee name - improved logic
    let partyName = 'N/A';
    const nameKeys = ['name', 'partyname', 'payeename', 'vendorname'];
    const nameKey = Object.keys(row).find(key => 
      nameKeys.some(nk => key.toLowerCase().includes(nk.toLowerCase()))
    );
    
    if (nameKey && row[nameKey] && String(row[nameKey]).trim().length > 0) {
      partyName = String(row[nameKey]).trim();
    } else if (description) {
      // Extract from description - pattern: "Name / account@upi" or "Name / Account Number"
      const nameMatch = description.match(/^([^\/@]+?)(?:\s*\/|\s*@|$)/);
      if (nameMatch && nameMatch.length > 0) {
        partyName = nameMatch[1].trim();
      }
    }
    
    // Extract bank account - improved logic
    let bankAccount = '-';
    if (description) {
      // Try to find account number patterns
      const accountPatterns = [
        /(?:A\/c|Account|Acc)[\s]*[No\.]*[\s]*[:]*[\s]*(\d{8,})/i,  // "A/c No. 1234567890"
        /(?:A\/c|Account|Acc)[\s]*[No\.]*[\s]*[:]*[\s]*(\d{4,}[\d\s]{4,})/i,  // "Account No: 1234 5678 90"
        /(?:ending|a\/c|account)[\s]*(\d{4,})/i,  // "ending 5037" or "account 123456"
        /[\/@](\d{8,})/,  // Numbers after / or @
        /\b(\d{10,})\b/,  // Any 10+ digit number
      ];
      
      for (const pattern of accountPatterns) {
        const match = description.match(pattern);
        if (match && match[1]) {
          bankAccount = match[1].replace(/\s+/g, '');
          break;
        }
      }
    }

    // Generate voucher number - sequential format: ELLEN/PV/2025/01
    const currentYear = new Date().getFullYear();
    const voucherNumber = `ELLEN/PV/${currentYear}/${String(voucherCounter.count + 1).padStart(2, '0')}`;
    voucherCounter.count++;

    // Debit Voucher Content
    // Add page break before first debit voucher if there are credit invoices, or before subsequent vouchers
    children.push(
      new Paragraph({
        pageBreakBefore: creditRows.length > 0 || i > 0, // Page break if credit rows exist or if not first debit voucher
        children: [
          new TextRun({
            text: companyName,
            bold: true,
            size: 28,
          }),
        ],
        alignment: AlignmentType.LEFT,
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: `GSTIN : ${companyGSTIN}`,
            size: 22,
          }),
        ],
        alignment: AlignmentType.LEFT,
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: companyAddress,
            size: 22,
          }),
        ],
        alignment: AlignmentType.LEFT,
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: companyEmail,
            size: 22,
          }),
        ],
        alignment: AlignmentType.LEFT,
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: companyPhone,
            size: 22,
          }),
        ],
        alignment: AlignmentType.LEFT,
        spacing: { after: 600 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'PAYMENT VOUCHER / Vendor INVOICE',
            bold: true,
            size: 32,
          }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 600 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'Invoice Details:',
            bold: true,
            size: 24,
          }),
        ],
        spacing: { after: 200 },
      }),
      // Serial Number - Line by line (if available)
      ...(serialNumber ? [
        new Paragraph({
          children: [
            new TextRun({
              text: 'Serial Number: ',
              bold: true,
              size: 22,
            }),
            new TextRun({
              text: serialNumber,
              size: 22,
            }),
          ],
          spacing: { after: 200 },
        }),
      ] : []),
      new Paragraph({
        children: [
          new TextRun({
            text: 'Voucher No.: ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: voucherNumber,
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'Date: ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: invoiceDate || '-',
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'Mode of Payment: ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: 'UPI',
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'Payee Name: ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: partyName || '-',
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'Bank Account: ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: bankAccount || '-',
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'Purpose of Payment: ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: 'Payment collected from students is paid to tutors/institutions.',
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'Amount (₹): ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: `₹ ${debitAmount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
            size: 22,
            bold: true,
          }),
        ],
        spacing: { after: 400 },
      }),
      new Paragraph({
        spacing: { after: 400 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: `Amount in Words: ${amountInWords}`,
            size: 22,
          }),
        ],
        spacing: { after: 800 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'All payments are made via SBI RTGS/NEFT from Ellen Information Technology Solutions Pvt. Ltd. (A/c No. ending 5037) towards business services rendered under the LearnsConnect operations. Each transfer is supported by work orders, GST invoices, and digital approval records.',
            size: 20,
            italics: true,
          }),
        ],
        spacing: { after: 800 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: '',
            break: 1,
          }),
        ],
      }),
    );
  }

  // Create document
  const doc = new Document({
    sections: [{
      children,
    }],
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, filename);
};

// Helper function to get debit amount
const getDebitAmount = (row) => {
  for (const key of Object.keys(row)) {
    const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (normalizedKey.includes('debit') || normalizedKey.includes('dr')) {
      const value = row[key];
      if (value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
        const amt = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
        if (amt > 0) return amt;
      }
    }
  }
  return 0;
};

const formatHeader = (header) => {
  return header
    .replace(/([A-Z])/g, ' $1')
    .replace(/^./, str => str.toUpperCase())
    .trim()
    .split(' ')
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');
};

// Helper function to detect if data contains debit transactions
const detectIsDebit = (dataToCheck) => {
  if (!dataToCheck || dataToCheck.length === 0) return false;
  
  let creditCount = 0;
  let debitCount = 0;
  
  for (const row of dataToCheck) {
    let rowHasCredit = false;
    let rowHasDebit = false;
    
    for (const key of Object.keys(row)) {
      const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
      const value = row[key];
      
      if ((normalizedKey.includes('credit') || normalizedKey.includes('cr')) && 
          value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
        const amt = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
        if (amt > 0) rowHasCredit = true;
      }
      
      if ((normalizedKey.includes('debit') || normalizedKey.includes('dr')) && 
          value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
        const amt = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
        if (amt > 0) rowHasDebit = true;
      }
    }
    
    if (rowHasCredit && !rowHasDebit) creditCount++;
    if (rowHasDebit && !rowHasCredit) debitCount++;
  }
  
  return debitCount >= creditCount;
};

  // Export all formats (CSV, XLSX, DOCX) as a ZIP file
  export const exportToZIP = async (data, billingType, transactionType = 'all', filename = 'ellen-invoice-all.zip') => {
    if (!data || data.length === 0) {
      alert('No data to export');
      return;
    }

    const zip = new JSZip();
    const timestamp = new Date().toISOString().slice(0, 10);
    const baseFilename = filename.replace('.zip', '').replace('ellen-invoice-all', 'ellen-invoice');

    try {
      // 1. Add CSV file to ZIP
      const csv = Papa.unparse(data, { header: true });
      zip.file(`${baseFilename}-${timestamp}.csv`, csv);

      // 2. Add XLSX file to ZIP
      const ws = XLSX.utils.json_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
      const xlsxBuffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
      zip.file(`${baseFilename}-${timestamp}.xlsx`, xlsxBuffer);

      // 3. Detect if data contains both credit and debit rows (like exportToDOCXMixed)
      const creditRows = [];
      const debitRows = [];
      
      for (const row of data) {
        let hasCredit = false;
        let hasDebit = false;
        
        // Check for credit value
        for (const key of Object.keys(row)) {
          const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
          const value = row[key];
          
          if ((normalizedKey.includes('credit') || normalizedKey.includes('cr')) && 
              value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
            const amt = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
            if (amt > 0) {
              hasCredit = true;
              break;
            }
          }
        }
        
        // Check for debit value
        for (const key of Object.keys(row)) {
          const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
          const value = row[key];
          
          if ((normalizedKey.includes('debit') || normalizedKey.includes('dr')) && 
              value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
            const amt = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
            if (amt > 0) {
              hasDebit = true;
              break;
            }
          }
        }
        
        // Classify row: if has credit and no debit = credit, if has debit and no credit = debit
        if (hasCredit && !hasDebit) {
          creditRows.push(row);
        } else if (hasDebit && !hasCredit) {
          debitRows.push(row);
        }
      }
      
      const hasCredit = creditRows.length > 0;
      const hasDebit = debitRows.length > 0;
      const isMixed = hasCredit && hasDebit;

      // Generate individual PDF files for each record
      // Process all records (credit and debit) in order
      const allRows = [...data]; // Keep original order
      
      for (let i = 0; i < allRows.length; i++) {
        const row = allRows[i];
        
        // Determine if this is credit or debit
        let isRowCredit = false;
        let isRowDebit = false;
        
        for (const key of Object.keys(row)) {
          const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
          const value = row[key];
          
          if ((normalizedKey.includes('credit') || normalizedKey.includes('cr')) && 
              value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
            const amt = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
            if (amt > 0) {
              isRowCredit = true;
              break;
            }
          }
        }
        
        if (!isRowCredit) {
          for (const key of Object.keys(row)) {
            const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
            const value = row[key];
            
            if ((normalizedKey.includes('debit') || normalizedKey.includes('dr')) && 
                value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
              const amt = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
              if (amt > 0) {
                isRowDebit = true;
                break;
              }
            }
          }
        }
        
        // Generate PDF for this single record
        const singleRecordData = [row];
        const pdfBuffer = await generatePDFBlob(singleRecordData, billingType, transactionType);
        
        if (pdfBuffer) {
          // Generate filename based on record type and index
          const recordType = isRowCredit ? 'Payment-Receipt' : 'Vendor-Invoice';
          const pdfFilename = `${recordType}-${timestamp}-${String(i + 1).padStart(3, '0')}.pdf`;
          zip.file(pdfFilename, pdfBuffer);
        }
      }

      /* Old DOCX generation code removed - now using PDF only
      if (isMixed) {
        // Mixed data: Use exportToDOCXMixed logic - generate both credit invoices and debit vouchers in one DOCX
        const docxFilename = `invoice-mixed-report-${timestamp}.docx`;
        
        // Create a blob using the mixed export logic
        const companyName = 'ELLEN INFORMATION TECHNOLOGY SOLUTIONS PRIVATE LIMITED';
        const companyGSTIN = '33AAHCE0984H1ZN';
        const companyAddress = 'Registered Office : 8/3, Athreyapuram 2nd Street, Choolaimedu, Chennai – 600094';
        const companyEmail = 'Email : support@learnsconnect.com';
        const companyPhone = 'Phone : +91 8489357705';

        const invoiceCounter = { count: 0 };
        const voucherCounter = { count: 0 };
        const children = [];

        // Generate credit invoices first
        for (let i = 0; i < creditRows.length; i++) {
          const row = creditRows[i];
          const amount = getAmount(row);
          const amountInWords = numberToWords(Math.floor(amount));
          
          // Get serial number from row data
          const serialNoKeys = ['serialno', 'serialnumber', 'serial no', 's.no', 'sno'];
          const serialNoKey = Object.keys(row).find(key => 
            serialNoKeys.some(sk => key.toLowerCase().replace(/[^a-z0-9]/g, '').includes(sk.toLowerCase().replace(/[^a-z0-9]/g, '')))
          );
          const serialNumber = serialNoKey ? String(row[serialNoKey] || '') : '';
          
          const dateKeys = ['date', 'txndate', 'transactiondate', 'paymentdate', 'valuedate'];
          const dateKey = Object.keys(row).find(key => 
            dateKeys.some(dk => key.toLowerCase().includes(dk.toLowerCase()))
          );
          const invoiceDate = formatDate(row[dateKey] || null);

          const descKey = Object.keys(row).find(key => 
            key.toLowerCase().includes('description') || 
            key.toLowerCase().includes('particulars') ||
            key.toLowerCase().includes('narration')
          );
          const description = descKey ? String(row[descKey] || '') : '';

          let partyName = 'N/A';
          const nameKeys = ['name', 'partyname', 'client', 'payer', 'studentname', 'payeename', 'vendorname'];
          const nameKey = Object.keys(row).find(key => 
            nameKeys.some(nk => key.toLowerCase().includes(nk.toLowerCase()))
          );
          
          if (nameKey && row[nameKey]) {
            partyName = String(row[nameKey]).trim();
          } else if (description) {
            const nameMatch = description.match(/^([^\/]+?)(?:\s*\/|\s*@|$)/);
            if (nameMatch && nameMatch.length > 0) {
              partyName = nameMatch[1].trim();
            }
          }

          const invoiceNumber = generateInvoiceNumber(invoiceCounter, serialNumber);
          const paymentMode = detectPaymentMode(description, row);

          // Credit Invoice Content
          children.push(
            new Paragraph({
              pageBreakBefore: i > 0,
              children: [
                new TextRun({
                  text: companyName,
                  bold: true,
                  size: 28,
                }),
              ],
              alignment: AlignmentType.LEFT,
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: `GSTIN : ${companyGSTIN}`,
                  size: 22,
                }),
              ],
              alignment: AlignmentType.LEFT,
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: companyAddress,
                  size: 22,
                }),
              ],
              alignment: AlignmentType.LEFT,
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: companyEmail,
                  size: 22,
                }),
              ],
              alignment: AlignmentType.LEFT,
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: companyPhone,
                  size: 22,
                }),
              ],
              alignment: AlignmentType.LEFT,
              spacing: { after: 600 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: 'PAYMENT RECEIPT',
                  bold: true,
                  size: 32,
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: { after: 600 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: 'Payment Receipt Voucher No.: ',
                  bold: true,
                  size: 22,
                }),
                new TextRun({
                  text: invoiceNumber,
                  size: 22,
                }),
              ],
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: 'Date: ',
                  bold: true,
                  size: 22,
                }),
                new TextRun({
                  text: invoiceDate || '-',
                  size: 22,
                }),
              ],
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: 'Mode of Payment: ',
                  bold: true,
                  size: 22,
                }),
                new TextRun({
                  text: 'UPI',
                  size: 22,
                }),
              ],
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: 'Received from Payer (Name and Details): ',
                  bold: true,
                  size: 22,
                }),
                new TextRun({
                  text: description || partyName || '-',
                  size: 22,
                }),
              ],
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: 'Purpose of Payment: ',
                  bold: true,
                  size: 22,
                }),
                new TextRun({
                  text: 'Payment collected from students for Course/Batch',
                  size: 22,
                }),
              ],
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: 'Amount (₹): ',
                  bold: true,
                  size: 22,
                }),
                new TextRun({
                  text: `₹ ${amount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
                  size: 22,
                  bold: true,
                }),
              ],
              spacing: { after: 400 },
            }),
            new Paragraph({
              spacing: { after: 400 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: `Amount in Words: ${amountInWords}`,
                  size: 22,
                }),
              ],
              spacing: { after: 800 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: 'All payments are made via SBI RTGS/NEFT from Ellen Information Technology Solutions Pvt. Ltd. (A/c No. ending 5037) towards business services rendered under the LearnsConnect operations. Each transfer is supported by work orders, GST invoices, and digital approval records.',
                  size: 20,
                  italics: true,
                }),
              ],
              spacing: { after: 800 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: '',
                  break: 1,
                }),
              ],
            }),
          );
        }

        // Generate debit vouchers
        for (let i = 0; i < debitRows.length; i++) {
          const row = debitRows[i];
          
          // Get serial number from row data
          const serialNoKeys = ['serialno', 'serialnumber', 'serial no', 's.no', 'sno'];
          const serialNoKey = Object.keys(row).find(key => 
            serialNoKeys.some(sk => key.toLowerCase().replace(/[^a-z0-9]/g, '').includes(sk.toLowerCase().replace(/[^a-z0-9]/g, '')))
          );
          const serialNumber = serialNoKey ? String(row[serialNoKey] || '') : '';
          
          const debitAmount = getDebitAmount(row);
          const amountInWords = numberToWords(Math.floor(debitAmount));
          
          const dateKeys = ['date', 'txndate', 'transactiondate', 'paymentdate', 'valuedate'];
          const dateKey = Object.keys(row).find(key => 
            dateKeys.some(dk => key.toLowerCase().includes(dk.toLowerCase()))
          );
          const invoiceDate = formatDate(row[dateKey] || null);

          const descKey = Object.keys(row).find(key => 
            key.toLowerCase().includes('description') || 
            key.toLowerCase().includes('particulars') ||
            key.toLowerCase().includes('narration')
          );
          const description = descKey ? String(row[descKey] || '') : '';

          // Extract payee name
          let partyName = 'N/A';
          const nameKeys = ['name', 'partyname', 'payeename', 'vendorname'];
          const nameKey = Object.keys(row).find(key => 
            nameKeys.some(nk => key.toLowerCase().includes(nk.toLowerCase()))
          );
          
          if (nameKey && row[nameKey] && String(row[nameKey]).trim().length > 0) {
            partyName = String(row[nameKey]).trim();
          } else if (description) {
            const nameMatch = description.match(/^([^\/@]+?)(?:\s*\/|\s*@|$)/);
            if (nameMatch && nameMatch.length > 0) {
              partyName = nameMatch[1].trim();
            }
          }
          
          // Extract bank account
          let bankAccount = '-';
          if (description) {
            const accountPatterns = [
              /(?:A\/c|Account|Acc)[\s]*[No\.]*[\s]*[:]*[\s]*(\d{8,})/i,
              /(?:A\/c|Account|Acc)[\s]*[No\.]*[\s]*[:]*[\s]*(\d{4,}[\d\s]{4,})/i,
              /(?:ending|a\/c|account)[\s]*(\d{4,})/i,
              /[\/@](\d{8,})/,
              /\b(\d{10,})\b/,
            ];
            
            for (const pattern of accountPatterns) {
              const match = description.match(pattern);
              if (match && match[1]) {
                bankAccount = match[1].replace(/\s+/g, '');
                break;
              }
            }
          }

          const currentYear = new Date().getFullYear();
          const voucherNumber = `ELLEN/PV/${currentYear}/${String(voucherCounter.count + 1).padStart(2, '0')}`;
          voucherCounter.count++;

          // Debit Voucher Content
          children.push(
            new Paragraph({
              pageBreakBefore: creditRows.length > 0 || i > 0,
              children: [
                new TextRun({
                  text: companyName,
                  bold: true,
                  size: 28,
                }),
              ],
              alignment: AlignmentType.LEFT,
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: `GSTIN : ${companyGSTIN}`,
                  size: 22,
                }),
              ],
              alignment: AlignmentType.LEFT,
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: companyAddress,
                  size: 22,
                }),
              ],
              alignment: AlignmentType.LEFT,
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: companyEmail,
                  size: 22,
                }),
              ],
              alignment: AlignmentType.LEFT,
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: companyPhone,
                  size: 22,
                }),
              ],
              alignment: AlignmentType.LEFT,
              spacing: { after: 600 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: 'PAYMENT VOUCHER / Vendor INVOICE',
                  bold: true,
                  size: 32,
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: { after: 600 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: 'Invoice Details:',
                  bold: true,
                  size: 24,
                }),
              ],
              spacing: { after: 200 },
            }),
            // Serial Number - Line by line (if available)
            ...(serialNumber ? [
              new Paragraph({
                children: [
                  new TextRun({
                    text: 'Serial Number: ',
                    bold: true,
                    size: 22,
                  }),
                  new TextRun({
                    text: serialNumber,
                    size: 22,
                  }),
                ],
                spacing: { after: 200 },
              }),
            ] : []),
            new Paragraph({
              children: [
                new TextRun({
                  text: 'Voucher No.: ',
                  bold: true,
                  size: 22,
                }),
                new TextRun({
                  text: voucherNumber,
                  size: 22,
                }),
              ],
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: 'Date: ',
                  bold: true,
                  size: 22,
                }),
                new TextRun({
                  text: invoiceDate || '-',
                  size: 22,
                }),
              ],
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: 'Mode of Payment: ',
                  bold: true,
                  size: 22,
                }),
                new TextRun({
                  text: 'UPI',
                  size: 22,
                }),
              ],
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: 'Payee Name: ',
                  bold: true,
                  size: 22,
                }),
                new TextRun({
                  text: partyName || '-',
                  size: 22,
                }),
              ],
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: 'Bank Account: ',
                  bold: true,
                  size: 22,
                }),
                new TextRun({
                  text: bankAccount || '-',
                  size: 22,
                }),
              ],
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: 'Purpose of Payment: ',
                  bold: true,
                  size: 22,
                }),
                new TextRun({
                  text: 'Payment collected from students is paid to tutors/institutions.',
                  size: 22,
                }),
              ],
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: 'Amount (₹): ',
                  bold: true,
                  size: 22,
                }),
                new TextRun({
                  text: `₹ ${debitAmount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
                  size: 22,
                  bold: true,
                }),
              ],
              spacing: { after: 400 },
            }),
            new Paragraph({
              spacing: { after: 400 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: `Amount in Words: ${amountInWords}`,
                  size: 22,
                }),
              ],
              spacing: { after: 800 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: 'All payments are made via SBI RTGS/NEFT from Ellen Information Technology Solutions Pvt. Ltd. (A/c No. ending 5037) towards business services rendered under the LearnsConnect operations. Each transfer is supported by work orders, GST invoices, and digital approval records.',
                  size: 20,
                  italics: true,
                }),
              ],
              spacing: { after: 800 },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: '',
                  break: 1,
                }),
              ],
            }),
          );
        }

        // Create document and add to ZIP
        const doc = new Document({
          sections: [{ children }],
        });

        const docxBlob = await Packer.toBlob(doc);
        const docxArrayBuffer = await docxBlob.arrayBuffer();
        zip.file(docxFilename, docxArrayBuffer);
        
      } else if (hasDebit) {
        // Only debit data - Use debit blob generator
        const voucherCounter = { count: 0 };
        const batches = Math.ceil(debitRows.length / BATCH_SIZE);
        
        for (let batchIndex = 0; batchIndex < batches; batchIndex++) {
          const startIndex = batchIndex * BATCH_SIZE;
          const endIndex = Math.min(startIndex + BATCH_SIZE, debitRows.length);
          const batchData = debitRows.slice(startIndex, endIndex);

          const docxFilename = batches > 1 
            ? `payment-voucher-${timestamp}_part${batchIndex + 1}.docx`
            : `payment-voucher-${timestamp}.docx`;
          
          const docxBlob = await generateDOCXBlobDebit(batchData, voucherCounter);
          const docxArrayBuffer = await docxBlob.arrayBuffer();
          zip.file(docxFilename, docxArrayBuffer);
        }
      } else {
        // Only credit data - Use credit blob generator
        const companyName = 'ELLEN INFORMATION TECHNOLOGY SOLUTIONS PRIVATE LIMITED';
        const companyGSTIN = '33AAHCE0984H1ZN';
        const companyAddress = 'Registered Office : 8/3, Athreyapuram 2nd Street, Choolaimedu, Chennai – 600094';
        const companyEmail = 'Email : support@learnsconnect.com';
        const companyPhone = 'Phone : +91 8489357705';

        const invoiceCounter = { count: 0 };
        const payeeNameMap = {};
        const batches = Math.ceil(creditRows.length / BATCH_SIZE);

        for (let batchIndex = 0; batchIndex < batches; batchIndex++) {
          const startIndex = batchIndex * BATCH_SIZE;
          const endIndex = Math.min(startIndex + BATCH_SIZE, creditRows.length);
          const batchData = creditRows.slice(startIndex, endIndex);

          const docxFilename = batches > 1 
            ? `${baseFilename}-${timestamp}_part${batchIndex + 1}.docx`
            : `${baseFilename}-${timestamp}.docx`;
          
          const docxBlob = await generateDOCXBlob(batchData, billingType, transactionType, invoiceCounter, payeeNameMap, companyName, companyGSTIN, companyAddress, companyEmail, companyPhone);
          const docxArrayBuffer = await docxBlob.arrayBuffer();
          zip.file(docxFilename, docxArrayBuffer);
        }
      }
      */ // End of old DOCX code

      // Generate and download ZIP file
      const zipBlob = await zip.generateAsync({ type: 'blob' });
      saveAs(zipBlob, filename);
    } catch (error) {
      console.error('Error creating ZIP file:', error);
      alert('Error creating ZIP file. Please try again.');
    }
  };

// Helper function to generate DOCX blob (reused from exportToDOCX logic)
const generateDOCXBlob = async (batchData, billingType, transactionType, invoiceCounter, payeeNameMap, companyName, companyGSTIN, companyAddress, companyEmail, companyPhone) => {
  const children = [];

  for (let i = 0; i < batchData.length; i++) {
    const row = batchData[i];
    const amount = getAmount(row);
    const amountInWords = numberToWords(Math.floor(amount));
    
    // Get serial number from row data
    const serialNoKeys = ['serialno', 'serialnumber', 'serial no', 's.no', 'sno'];
    const serialNoKey = Object.keys(row).find(key => 
      serialNoKeys.some(sk => key.toLowerCase().replace(/[^a-z0-9]/g, '').includes(sk.toLowerCase().replace(/[^a-z0-9]/g, '')))
    );
    const serialNumber = serialNoKey ? String(row[serialNoKey] || '') : '';
    
    const dateKeys = ['date', 'txndate', 'transactiondate', 'paymentdate', 'valuedate'];
    const dateKey = Object.keys(row).find(key => 
      dateKeys.some(dk => key.toLowerCase().includes(dk.toLowerCase()))
    );
    const invoiceDate = formatDate(row[dateKey] || new Date());

    const descKey = Object.keys(row).find(key => 
      key.toLowerCase().includes('description') || 
      key.toLowerCase().includes('particulars') ||
      key.toLowerCase().includes('narration')
    );
    const description = descKey ? String(row[descKey] || '') : '';

    let partyName = 'N/A';
    const nameKeys = ['name', 'partyname', 'client', 'payer', 'studentname', 'payeename', 'vendorname'];
    const nameKey = Object.keys(row).find(key => 
      nameKeys.some(nk => key.toLowerCase().includes(nk.toLowerCase()))
    );
    
    if (nameKey && row[nameKey]) {
      partyName = String(row[nameKey]).trim();
    } else if (description) {
      const nameMatch = description.split('/')[0].trim();
      if (nameMatch && nameMatch.length > 0) {
        partyName = nameMatch;
      }
    }

    const isCredit = isCreditTransaction(row);
    const invoiceNumber = isCredit 
      ? generateInvoiceNumber(invoiceCounter, serialNumber)
      : generateVoucherNumber(partyName, i, payeeNameMap);
    
    // Get Ref No
    const refKeys = ['refno', 'refno', 'ref', 'chequeno', 'utrno', 'utr', 'branchcode'];
    const refKey = Object.keys(row).find(key => 
      refKeys.some(rk => key.toLowerCase().includes(rk.toLowerCase()))
    );
    let refNo = refKey ? String(row[refKey] || '') : '';
    
    // Detect payment mode
    const paymentMode = detectPaymentMode(description, refNo);

    if (isCredit) {
      // PAYMENT RECEIPT FORMAT - Line by line format
      // Get description for "Received from Payer (Name and Details)"
      const payerDetails = description || partyName || '-';
      
      children.push(
        // Company Name (with page break before for new receipts)
        new Paragraph({
          children: [
            new TextRun({
              text: companyName,
              bold: true,
              size: 28,
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: { after: 200 },
          pageBreakBefore: i > 0, // Start new page for each receipt except the first one
        }),
        // GSTIN
        new Paragraph({
          children: [
            new TextRun({
              text: `GSTIN : ${companyGSTIN}`,
              size: 22,
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: { after: 100 },
        }),
        // Address
        new Paragraph({
          children: [
            new TextRun({
              text: companyAddress,
              size: 22,
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: { after: 100 },
        }),
        // Email
        new Paragraph({
          children: [
            new TextRun({
              text: companyEmail,
              size: 22,
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: { after: 100 },
        }),
        // Phone
        new Paragraph({
          children: [
            new TextRun({
              text: companyPhone,
              size: 22,
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: { after: 400 },
        }),

        // Dividing Line
        new Paragraph({
          children: [
            new TextRun({
              text: '________________________________________________________________________________',
              size: 22,
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: { after: 400 },
        }),

        // Payment Receipt Title
        new Paragraph({
          children: [
            new TextRun({
              text: 'PAYMENT RECEIPT',
              bold: true,
              size: 32,
            }),
          ],
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
        }),

        // Payment Receipt Voucher No. - Line by line (ref from excel - uses serial number)
        new Paragraph({
          children: [
            new TextRun({
              text: 'Payment Receipt Voucher No.: ',
              bold: true,
              size: 22,
            }),
            new TextRun({
              text: invoiceNumber,
              size: 22,
            }),
          ],
          spacing: { after: 200 },
        }),

        // Date - Line by line (ref from excel)
        new Paragraph({
          children: [
            new TextRun({
              text: 'Date: ',
              bold: true,
              size: 22,
            }),
            new TextRun({
              text: invoiceDate || '-',
              size: 22,
            }),
          ],
          spacing: { after: 200 },
        }),

        // Mode of Payment - Line by line (default: UPI)
        new Paragraph({
          children: [
            new TextRun({
              text: 'Mode of Payment: ',
              bold: true,
              size: 22,
            }),
            new TextRun({
              text: 'UPI',
              size: 22,
            }),
          ],
          spacing: { after: 200 },
        }),

        // Received from Payer (Name and Details) - Line by line (ref from excel - description)
        new Paragraph({
          children: [
            new TextRun({
              text: 'Received from Payer (Name and Details): ',
              bold: true,
              size: 22,
            }),
            new TextRun({
              text: payerDetails,
              size: 22,
            }),
          ],
          spacing: { after: 200 },
        }),

        // Purpose of Payment - Line by line (default: Payment collected from students for Course/Batch)
        new Paragraph({
          children: [
            new TextRun({
              text: 'Purpose of Payment: ',
              bold: true,
              size: 22,
            }),
            new TextRun({
              text: 'Payment collected from students for Course/Batch',
              size: 22,
            }),
          ],
          spacing: { after: 200 },
        }),

        // Amount (₹) - Line by line (ref from excel)
        new Paragraph({
          children: [
            new TextRun({
              text: 'Amount (₹): ',
              bold: true,
              size: 22,
            }),
            new TextRun({
              text: `₹ ${amount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
              size: 22,
              bold: true,
            }),
          ],
          spacing: { after: 400 },
        }),

        // Amount in Words (calculated from amount)
        new Paragraph({
          children: [
            new TextRun({
              text: `Amount in Words: ${amountInWords}`,
              size: 22,
            }),
          ],
          spacing: { after: 800 },
        }),

        // Footer Disclaimer
        new Paragraph({
          children: [
            new TextRun({
              text: 'All payments are made via SBI RTGS/NEFT from Ellen Information Technology Solutions Pvt. Ltd. (A/c No. ending 5037) towards business services rendered under the LearnsConnect operations. Each transfer is supported by work orders, GST invoices, and digital approval records.',
              size: 18,
              italics: true,
            }),
          ],
          alignment: AlignmentType.LEFT,
          spacing: { before: 400, after: 400 },
        }),
        // Page break after receipt (ensures next receipt starts on new page)
        new Paragraph({
          children: [new TextRun({ text: '', break: 1 })],
        }),
      );
    } else {
      children.push(
        // Company Header (with page break before for new vouchers)
        new Paragraph({
          children: [new TextRun({ text: companyName, bold: true, size: 28 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          pageBreakBefore: i > 0, // Start new page for each voucher except the first one
        }),
        new Paragraph({
          children: [new TextRun({ text: `GSTIN : ${companyGSTIN}`, size: 22 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
        }),
        new Paragraph({
          children: [new TextRun({ text: companyAddress, size: 22 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
        }),
        new Paragraph({
          children: [new TextRun({ text: companyEmail, size: 22 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
        }),
        new Paragraph({
          children: [new TextRun({ text: companyPhone, size: 22 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
        }),
        new Paragraph({
          children: [new TextRun({ text: 'PAYMENT VOUCHER / SUPPLIER INVOICE', bold: true, size: 32 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
        }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Voucher Number', bold: true })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: invoiceNumber })] })] }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Date', bold: true })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: invoiceDate })] })] }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Paid To', bold: true })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: partyName })] })] }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Payment Mode', bold: true })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: paymentMode || '-' })] })] }),
              ],
            }),
          ],
        }),
        new Paragraph({
          children: [new TextRun({ text: 'Break-up of Amount', bold: true, size: 24 })],
          spacing: { before: 400, after: 200 },
        }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Particulars', bold: true })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Amount (INR)', bold: true })] })] }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: description || 'N/A' })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `₹${amount.toFixed(2)}` })] })] }),
              ],
            }),
          ],
        }),
        new Paragraph({
          children: [new TextRun({ text: `Amount in Words: ${amountInWords} Only`, italics: true, size: 22 })],
          spacing: { before: 200, after: 400 },
        }),
        new Paragraph({
          children: [new TextRun({ text: 'This is a system-generated payment voucher. No signature required.', size: 18 })],
          alignment: AlignmentType.CENTER,
          spacing: { before: 600 },
        }),
        // Page break after voucher (ensures next voucher/invoice starts on new page)
        new Paragraph({
          children: [new TextRun({ text: '', break: 1 })],
        }),
      );
    }
  }

  const doc = new Document({
    sections: [{ children }],
  });

  return await Packer.toBlob(doc);
};
