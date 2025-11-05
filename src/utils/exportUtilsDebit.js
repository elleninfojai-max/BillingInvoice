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

// Generate voucher number for debit transactions (format: ELLEN/PV/2025/01) - Sequential numbering
const generateVoucherNumber = (voucherCounter) => {
  // Track global voucher counter
  if (!voucherCounter.count) {
    voucherCounter.count = 0;
  }
  voucherCounter.count++;
  
  const sequence = String(voucherCounter.count).padStart(2, '0');
  const currentYear = new Date().getFullYear();
  return `ELLEN/PV/${currentYear}/${sequence}`;
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
  } else if (description) {
    // Extract from description - look for patterns like "Company Name / email" or "Company Name"
    // Remove account numbers, UPI IDs, and other non-name text
    const desc = String(description);
    
    // Split by common separators
    const parts = desc.split(/[\/|,]/).map(p => p.trim());
    
    // Find the part that looks like a name (has letters, not just numbers/special chars)
    for (const part of parts) {
      const cleanPart = part.replace(/A\/c[\s]*[No\.]*[\s]*[:]*[\s]*\d+/i, '')
                           .replace(/Account[\s]*[No\.]*[\s]*[:]*[\s]*\d+/i, '')
                           .replace(/@\w+/g, '')
                           .replace(/\d{10,}/g, '')
                           .trim();
      
      if (cleanPart.length > 2 && /[A-Za-z]/.test(cleanPart)) {
        partyName = cleanPart;
        break;
      }
    }
    
    // If still not found, extract first meaningful words
    if (partyName === 'N/A') {
      const words = description.split(/[\/\s]+/).filter(w => 
        w.length > 2 && 
        !/^\d+$/.test(w) && 
        !/@/.test(w) &&
        /[A-Za-z]/.test(w)
      );
      if (words.length > 0) {
        partyName = words.slice(0, 3).join(' ').trim();
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
      
      // Get date from various possible column names
      const dateKeys = ['date', 'txndate', 'transactiondate', 'paymentdate', 'valuedate'];
      const dateKey = Object.keys(row).find(key => 
        dateKeys.some(dk => key.toLowerCase().includes(dk.toLowerCase()))
      );
      const invoiceDate = formatDate(row[dateKey] || new Date());

      // Get description to extract payee name and bank account
      const descKey = Object.keys(row).find(key => 
        key.toLowerCase().includes('description') || 
        key.toLowerCase().includes('particulars') ||
        key.toLowerCase().includes('narration')
      );
      const description = descKey ? String(row[descKey] || '') : '';

      // Extract payee name from description
      const partyName = extractPayeeName(description, row);
      
      // Extract bank account from description
      const bankAccount = extractBankAccount(description);

      // Generate voucher number
      const voucherNumber = generateVoucherNumber(voucherCounter);

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

        // Title: Payment Voucher / Vendor Invoice
        new Paragraph({
          children: [
            new TextRun({
              text: 'PAYMENT VOUCHER / Vendor INVOICE',
              bold: true,
              size: 32,
            }),
          ],
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
        }),

        // Invoice Details
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

        // Date - Line by line
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

        // Mode of Payment - Line by line
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

        // Payee Name - Line by line
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

        // Bank Account - Line by line
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

        // Purpose of Payment - Line by line
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

    const partyName = extractPayeeName(description, row);
    const bankAccount = extractBankAccount(description);
    const voucherNumber = generateVoucherNumber(voucherCounter);

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
        children: [new TextRun({ text: 'PAYMENT VOUCHER / Vendor INVOICE', bold: true, size: 32 })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
      }),
      new Paragraph({
        children: [new TextRun({ text: 'Invoice Details:', bold: true, size: 24 })],
        spacing: { after: 200 },
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
      // Date - Line by line
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
      // Mode of Payment - Line by line
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
      // Payee Name - Line by line
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
      // Bank Account - Line by line
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
      // Purpose of Payment - Line by line
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
