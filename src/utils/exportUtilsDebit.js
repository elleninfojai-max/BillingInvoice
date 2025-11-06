import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  AlignmentType, 
  WidthType,
  Table,
  TableRow,
  TableCell,
} from 'docx';
import { saveAs } from 'file-saver';
import { findBeneficiaryData, getBeneficiaryField } from './excelDataLoader';

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
  if (!num || num === 0) return 'Rupees Zero Only';
  
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
  
  // Add "Rupees" prefix and proper formatting
  const amountWords = words.length > 0 ? words.join(' ') : 'Zero';
  return `Rupees ${amountWords} Only`;
};

// Helper to format date (handles "30 Oct 2025" format from PDF)
const formatDate = (dateStr) => {
  if (!dateStr) {
    const today = new Date();
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return `${today.getDate()} ${months[today.getMonth()]} ${today.getFullYear()}`;
  }
  
  try {
    // Check if it's already in "dd MMM yyyy" format (e.g., "30 Oct 2025")
    const ddmmyyyyMatch = String(dateStr).match(/(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})/i);
    if (ddmmyyyyMatch) {
      return `${ddmmyyyyMatch[1]} ${ddmmyyyyMatch[2]} ${ddmmyyyyMatch[3]}`;
    }
    
    // Try standard Date parsing
    const date = new Date(dateStr);
    if (!isNaN(date.getTime())) {
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      return `${date.getDate()} ${months[date.getMonth()]} ${date.getFullYear()}`;
    }
    
    return String(dateStr);
  } catch {
    return String(dateStr);
  }
};

// Helper to get debit amount value
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

// Generate voucher number for debit transactions (format: ELLEN/PV/2025/(S.No)) - Using Serial No
const generateVoucherNumber = (row) => {
  // Try to get Serial No from row
  const serialNoKeys = ['Serial No', 'Serial No.', 'SerialNo', 'Serial Number', 'S.No', 'S.No.', 'SNo'];
  const serialNoKey = Object.keys(row).find(key => 
    serialNoKeys.some(sk => key.toLowerCase().replace(/[^a-z0-9]/g, '') === sk.toLowerCase().replace(/[^a-z0-9]/g, ''))
  );
  
  let serialNo = '';
  if (serialNoKey && row[serialNoKey]) {
    serialNo = String(row[serialNoKey]).trim();
  }
  
  const currentYear = new Date().getFullYear();
  if (serialNo) {
    return `ELLEN/PV/${currentYear}/${serialNo}`;
  }
  
  // Fallback to sequential numbering if Serial No not found
  return `ELLEN/PV/${currentYear}/-`;
};

// Extract bank account number from description
const extractBankAccount = (description) => {
  if (!description) return '-';
  
  // Try to find account number patterns
  // Common patterns: A/c No., Account No., Acc No., A/c:, Account:, etc.
  const accountPatterns = [
    /(?:A\/c|Account|Acc)[\s]*[No\.]*[\s]*[:]*[\s]*(\d{8,})/i,  // "A/c No. 1234567890"
    /(?:A\/c|Account|Acc)[\s]*[No\.]*[\s]*[:]*[\s]*(\d{4,}[\d\s]{4,})/i,  // "Account No: 1234 5678 90"
    /(?:ending|a\/c|account)[\s]*(\d{4,})/i,  // "ending 5037" or "account 123456"
    /\b(\d{10,})\b/,  // Any 10+ digit number (likely account number)
  ];
  
  for (const pattern of accountPatterns) {
    const match = String(description).match(pattern);
    if (match && match[1]) {
      // Return the account number, removing spaces
      return match[1].replace(/\s+/g, '');
    }
  }
  
  return '-';
};

// Extract payee name from description (only the name portion)
const extractPayeeName = (description, row) => {
  let partyName = 'N/A';
  
  // Try direct name fields first
  const nameKeys = ['name', 'partyname', 'payeename', 'vendorname'];
  const nameKey = Object.keys(row).find(key => 
    nameKeys.some(nk => key.toLowerCase().includes(nk.toLowerCase()))
  );
  
  if (nameKey && row[nameKey] && String(row[nameKey]).trim().length > 0) {
    partyName = String(row[nameKey]).trim();
    return partyName;
  }
  
  if (description) {
    const desc = String(description).trim();
    
    // First, try to extract full company name before common separators
    // Look for patterns like "COMPANY NAME /", "COMPANY NAME @", "COMPANY NAME A/c", etc.
    const beforeSeparatorMatch = desc.match(/^([A-Z][A-Z\s&.,]+?)(?:\s*[/@|]\s*|A\/c|Account|UPI|@|\d{10,})/i);
    if (beforeSeparatorMatch && beforeSeparatorMatch[1]) {
      const extracted = beforeSeparatorMatch[1].trim();
      // Check if it looks like a company name (has multiple words, contains letters)
      const words = extracted.split(/\s+/).filter(w => /[A-Za-z]/.test(w));
      if (words.length >= 2 && extracted.length > 5) {
        partyName = extracted;
        return partyName;
      }
    }
    
    // Split by common separators and try each part
    const separators = /[\/|,@]/;
    const parts = desc.split(separators).map(p => p.trim()).filter(p => p.length > 0);
    
    // Find the part that looks like a company name (longest meaningful text)
    let longestPart = '';
    for (const part of parts) {
      const cleanPart = part
        .replace(/A\/c[\s]*[No\.]*[\s]*[:]*[\s]*\d+/i, '')
        .replace(/Account[\s]*[No\.]*[\s]*[:]*[\s]*\d+/i, '')
        .replace(/@\w+/g, '')
        .replace(/\d{10,}/g, '')
        .replace(/UPI/i, '')
        .replace(/RTGS|NEFT|IMPS/i, '')
        .trim();
      
      // Check if it looks like a company name
      const wordCount = cleanPart.split(/\s+/).filter(w => /[A-Za-z]/.test(w)).length;
      if (cleanPart.length > longestPart.length && 
          cleanPart.length > 5 && 
          wordCount >= 2 && 
          /[A-Za-z]/.test(cleanPart)) {
        longestPart = cleanPart;
      }
    }
    
    if (longestPart.length > 0) {
      partyName = longestPart;
      return partyName;
    }
    
    // Fallback: Extract first meaningful words (at least 3-4 words for company names)
    if (partyName === 'N/A') {
      const words = desc.split(/[\/\s,@]+/).filter(w => 
        w.length > 2 && 
        !/^\d+$/.test(w) && 
        !/@/.test(w) &&
        !/^upi$/i.test(w) &&
        !/^rtgs|neft|imps$/i.test(w) &&
        /[A-Za-z]/.test(w)
      );
      if (words.length >= 3) {
        // Take first 4-5 words for company names
        partyName = words.slice(0, 5).join(' ').trim();
      } else if (words.length > 0) {
        partyName = words.join(' ').trim();
      }
    }
  }
  
  return partyName;
};

// Main export function for debit invoices
export const exportToDOCXDebit = async (data, filename = 'payment-voucher-report.docx') => {
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

  // Voucher counter for sequential numbering
  const voucherCounter = { count: 0 };

  // Batch processing: 50 vouchers per file for better performance
  const BATCH_SIZE = 50;
  const batches = Math.ceil(data.length / BATCH_SIZE);

  for (let batchIndex = 0; batchIndex < batches; batchIndex++) {
    const startIndex = batchIndex * BATCH_SIZE;
    const endIndex = Math.min(startIndex + BATCH_SIZE, data.length);
    const batchData = data.slice(startIndex, endIndex);

    const children = [];

    // Process each record as individual payment voucher
    for (let i = 0; i < batchData.length; i++) {
      const row = batchData[i];
      const amount = getDebitAmount(row);
      const amountInWords = numberToWords(Math.floor(amount));
      
      // Get description to extract payee name
      const descKey = Object.keys(row).find(key => 
        key.toLowerCase().includes('description') || 
        key.toLowerCase().includes('particulars') ||
        key.toLowerCase().includes('narration')
      );
      const description = descKey ? String(row[descKey] || '') : '';

      // Extract payee name from description
      const partyName = extractPayeeName(description, row);
      
      // Find matching beneficiary data from Excel
      const beneficiaryData = await findBeneficiaryData(partyName);
      
      // Get date - prefer from transaction data (more accurate), fallback to Excel DB date
      let invoiceDate = '-';
      const dateKeys = ['date', 'txndate', 'transactiondate', 'paymentdate', 'valuedate'];
      const dateKey = Object.keys(row).find(key => 
        dateKeys.some(dk => key.toLowerCase().includes(dk.toLowerCase()))
      );
      invoiceDate = formatDate(row[dateKey] || null);
      
      // If no date from transaction, try Excel DB
      if (invoiceDate === '-' && beneficiaryData) {
        invoiceDate = getBeneficiaryField(beneficiaryData, ['Date', 'Transaction Date', 'Payment Date', 'Value Date']);
      }
      
      // Final fallback to current date
      if (invoiceDate === '-' || invoiceDate === null) {
        invoiceDate = formatDate(new Date());
      }

      // Get data from Excel or use "Refer Excel DB" as fallback
      // Updated to use new column names from database
      let payeeName = beneficiaryData ? getBeneficiaryField(beneficiaryData, ['Vendor Name', 'Name', 'Payee Name', 'Beneficiary Name']) : 'Refer Excel DB';
      if (payeeName === '-') payeeName = 'Refer Excel DB';
      
      let bankAccountNo = beneficiaryData ? getBeneficiaryField(beneficiaryData, ['Bank A/c No.', 'Bank Account No', 'Bank Account', 'Account No', 'Account Number', 'A/c No']) : 'Refer Excel DB';
      if (bankAccountNo === '-') bankAccountNo = 'Refer Excel DB';
      
      let bankNameAndDetails = beneficiaryData ? getBeneficiaryField(beneficiaryData, ['Bank Name & Branch (Verified)', 'Bank Name and Details', 'Bank Name', 'Bank Details', 'Bank']) : 'Refer Excel DB';
      if (bankNameAndDetails === '-') bankNameAndDetails = 'Refer Excel DB';
      
      let purposeOfPayment = beneficiaryData ? getBeneficiaryField(beneficiaryData, ['Description', 'Purpose of Payment', 'Purpose', 'Payment Purpose']) : 'Refer Excel DB';
      if (purposeOfPayment === '-') purposeOfPayment = 'Refer Excel DB';
      
      let categoryOfPayment = beneficiaryData ? getBeneficiaryField(beneficiaryData, ['Purpose / Category', 'Category of Payment', 'Category', 'Payment Category']) : 'Refer Excel DB';
      if (categoryOfPayment === '-') categoryOfPayment = 'Refer Excel DB';
      
      let modeOfPayment = beneficiaryData ? getBeneficiaryField(beneficiaryData, ['Mode of Payment', 'Payment Mode', 'Mode']) : 'RTGS/NEFT';
      
      // If mode of payment is still '-', default to RTGS/NEFT
      if (modeOfPayment === '-') {
        modeOfPayment = 'RTGS/NEFT';
      }

      // Generate voucher number using Serial No
      const voucherNumber = generateVoucherNumber(row);

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
              text: `GSTIN : ${companyGSTIN}`,
              size: 22,
            }),
          ],
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
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
          spacing: { after: 100 },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: companyPhone,
              size: 22,
            }),
          ],
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
        }),

        // Title: Vendor Invoice
        new Paragraph({
          children: [
            new TextRun({
              text: 'Vendor Invoice',
              bold: true,
              size: 32,
            }),
          ],
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
        }),

        // Voucher Details - Line by line format
        // Voucher No. - Line by line
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

        // Date - Line by line (from Excel)
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

        // Mode of Payment - Line by line (RTGS/NEFT)
        new Paragraph({
          children: [
            new TextRun({
              text: 'Mode of Payment: ',
              bold: true,
              size: 22,
            }),
            new TextRun({
              text: modeOfPayment === '-' ? 'RTGS/NEFT' : modeOfPayment,
              size: 22,
            }),
          ],
          spacing: { after: 200 },
        }),

        // Payee Name - Line by line (from Excel)
        new Paragraph({
          children: [
            new TextRun({
              text: 'Payee Name: ',
              bold: true,
              size: 22,
            }),
            new TextRun({
              text: payeeName || '-',
              size: 22,
            }),
          ],
          spacing: { after: 200 },
        }),

        // Bank Account No. - Line by line (from Excel)
        new Paragraph({
          children: [
            new TextRun({
              text: 'Bank Account No.: ',
              bold: true,
              size: 22,
            }),
            new TextRun({
              text: bankAccountNo || '-',
              size: 22,
            }),
          ],
          spacing: { after: 200 },
        }),

        // Bank Name and Details - Line by line (from Excel)
        new Paragraph({
          children: [
            new TextRun({
              text: 'Bank Name and Details: ',
              bold: true,
              size: 22,
            }),
            new TextRun({
              text: bankNameAndDetails || '-',
              size: 22,
            }),
          ],
          spacing: { after: 200 },
        }),

        // Purpose of Payment - Line by line (from Excel)
        new Paragraph({
          children: [
            new TextRun({
              text: 'Purpose of Payment: ',
              bold: true,
              size: 22,
            }),
            new TextRun({
              text: purposeOfPayment || '-',
              size: 22,
            }),
          ],
          spacing: { after: 200 },
        }),

        // Category of Payment - Line by line (from Excel)
        new Paragraph({
          children: [
            new TextRun({
              text: 'Category of Payment: ',
              bold: true,
              size: 22,
            }),
            new TextRun({
              text: categoryOfPayment || '-',
              size: 22,
            }),
          ],
          spacing: { after: 200 },
        }),

        // Amount (₹) - Line by line
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
        // Page break after voucher (ensures next voucher starts on new page)
        new Paragraph({
          children: [new TextRun({ text: '', break: 1 })],
        }),
      );
    }

    const doc = new Document({
      sections: [{
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

// Helper function to generate DOCX blob for ZIP export (reused from exportToDOCXDebit logic)
export const generateDOCXBlobDebit = async (batchData, voucherCounter) => {
  const companyName = 'ELLEN INFORMATION TECHNOLOGY SOLUTIONS PRIVATE LIMITED';
  const companyGSTIN = '33AAHCE0984H1ZN';
  const companyAddress = 'Registered Office : 8/3, Athreyapuram 2nd Street, Choolaimedu, Chennai – 600094';
  const companyEmail = 'Email : support@learnsconnect.com';
  const companyPhone = 'Phone : +91 8489357705';

  const children = [];

  for (let i = 0; i < batchData.length; i++) {
    const row = batchData[i];
    const amount = getDebitAmount(row);
    const amountInWords = numberToWords(Math.floor(amount));
    
    const descKey = Object.keys(row).find(key => 
      key.toLowerCase().includes('description') || 
      key.toLowerCase().includes('particulars') ||
      key.toLowerCase().includes('narration')
    );
    const description = descKey ? String(row[descKey] || '') : '';

    const partyName = extractPayeeName(description, row);
    
    // Find matching beneficiary data from Excel
    const beneficiaryData = await findBeneficiaryData(partyName);
    
    // Get date - prefer from transaction data (more accurate), fallback to Excel DB date
    let invoiceDate = '-';
    const dateKeys = ['date', 'txndate', 'transactiondate', 'paymentdate', 'valuedate'];
    const dateKey = Object.keys(row).find(key => 
      dateKeys.some(dk => key.toLowerCase().includes(dk.toLowerCase()))
    );
    invoiceDate = formatDate(row[dateKey] || null);
    
    // If no date from transaction, try Excel DB
    if (invoiceDate === '-' && beneficiaryData) {
      invoiceDate = getBeneficiaryField(beneficiaryData, ['Date', 'Transaction Date', 'Payment Date', 'Value Date']);
    }
    
    // Final fallback to current date
    if (invoiceDate === '-' || invoiceDate === null) {
      invoiceDate = formatDate(new Date());
    }

    // Get data from Excel or use "Refer Excel DB" as fallback
    // Updated to use new column names from database
    let payeeName = beneficiaryData ? getBeneficiaryField(beneficiaryData, ['Vendor Name', 'Name', 'Payee Name', 'Beneficiary Name']) : 'Refer Excel DB';
    if (payeeName === '-') payeeName = 'Refer Excel DB';
    
    let bankAccountNo = beneficiaryData ? getBeneficiaryField(beneficiaryData, ['Bank A/c No.', 'Bank Account No', 'Bank Account', 'Account No', 'Account Number', 'A/c No']) : 'Refer Excel DB';
    if (bankAccountNo === '-') bankAccountNo = 'Refer Excel DB';
    
    let bankNameAndDetails = beneficiaryData ? getBeneficiaryField(beneficiaryData, ['Bank Name & Branch (Verified)', 'Bank Name and Details', 'Bank Name', 'Bank Details', 'Bank']) : 'Refer Excel DB';
    if (bankNameAndDetails === '-') bankNameAndDetails = 'Refer Excel DB';
    
    let purposeOfPayment = beneficiaryData ? getBeneficiaryField(beneficiaryData, ['Description', 'Purpose of Payment', 'Purpose', 'Payment Purpose']) : 'Refer Excel DB';
    if (purposeOfPayment === '-') purposeOfPayment = 'Refer Excel DB';
    
    let categoryOfPayment = beneficiaryData ? getBeneficiaryField(beneficiaryData, ['Purpose / Category', 'Category of Payment', 'Category', 'Payment Category']) : 'Refer Excel DB';
    if (categoryOfPayment === '-') categoryOfPayment = 'Refer Excel DB';
    
    let modeOfPayment = beneficiaryData ? getBeneficiaryField(beneficiaryData, ['Mode of Payment', 'Payment Mode', 'Mode']) : 'RTGS/NEFT';
    
    // If mode of payment is still '-', default to RTGS/NEFT
    if (modeOfPayment === '-') {
      modeOfPayment = 'RTGS/NEFT';
    }

    // Generate voucher number using Serial No
    const voucherNumber = generateVoucherNumber(row);

    // PAYMENT VOUCHER FORMAT
    children.push(
      new Paragraph({
        children: [new TextRun({ text: companyName, bold: true, size: 28 })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        pageBreakBefore: i > 0,
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
        children: [new TextRun({ text: 'Vendor Invoice', bold: true, size: 32 })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
      }),
      // Voucher Details - Line by line format
      // Voucher No. - Line by line
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
      // Date - Line by line (from Excel)
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
      // Mode of Payment - Line by line (RTGS/NEFT)
      new Paragraph({
        children: [
          new TextRun({
            text: 'Mode of Payment: ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: modeOfPayment === '-' ? 'RTGS/NEFT' : modeOfPayment,
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
      // Payee Name - Line by line (from Excel)
      new Paragraph({
        children: [
          new TextRun({
            text: 'Payee Name: ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: payeeName || '-',
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
      // Bank Account No. - Line by line (from Excel)
      new Paragraph({
        children: [
          new TextRun({
            text: 'Bank Account No.: ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: bankAccountNo || '-',
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
      // Bank Name and Details - Line by line (from Excel)
      new Paragraph({
        children: [
          new TextRun({
            text: 'Bank Name and Details: ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: bankNameAndDetails || '-',
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
      // Purpose of Payment - Line by line (from Excel)
      new Paragraph({
        children: [
          new TextRun({
            text: 'Purpose of Payment: ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: purposeOfPayment || '-',
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
      // Category of Payment - Line by line (from Excel)
      new Paragraph({
        children: [
          new TextRun({
            text: 'Category of Payment: ',
            bold: true,
            size: 22,
          }),
          new TextRun({
            text: categoryOfPayment || '-',
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
      // Amount (₹) - Line by line
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
        children: [new TextRun({ text: `Amount in Words: ${amountInWords}`, size: 22 })],
        spacing: { after: 800 },
      }),
      new Paragraph({
        children: [new TextRun({ text: 'All payments are made via SBI RTGS/NEFT from Ellen Information Technology Solutions Pvt. Ltd. (A/c No. ending 5037) towards business services rendered under the LearnsConnect operations. Each transfer is supported by work orders, GST invoices, and digital approval records.', size: 18, italics: true })],
        alignment: AlignmentType.LEFT,
        spacing: { before: 400, after: 400 },
      }),
      new Paragraph({
        children: [new TextRun({ text: '', break: 1 })],
      }),
    );
  }

  const doc = new Document({
    sections: [{ children }],
  });

  return await Packer.toBlob(doc);
};
