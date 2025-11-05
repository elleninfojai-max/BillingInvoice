import jsPDF from 'jspdf';
import { saveAs } from 'file-saver';

// Helper functions (same as exportUtils.js)
const getAmount = (row) => {
  // First try to find credit amount
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
  
  return 0;
};

const isCreditTransaction = (row) => {
  let creditValue = 0;
  let debitValue = 0;
  
  for (const key of Object.keys(row)) {
    const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
    const value = row[key];
    
    if ((normalizedKey.includes('credit') || normalizedKey.includes('cr')) && 
        value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
      creditValue = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
    }
    
    if ((normalizedKey.includes('debit') || normalizedKey.includes('dr')) && 
        value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
      debitValue = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
    }
  }
  
  return creditValue > 0 && debitValue === 0;
};

const numberToWords = (num) => {
  const ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten',
    'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  
  if (num === 0) return 'Zero';
  if (num < 20) return ones[num];
  if (num < 100) return tens[Math.floor(num / 10)] + (num % 10 ? ' ' + ones[num % 10] : '');
  if (num < 1000) return ones[Math.floor(num / 100)] + ' Hundred' + (num % 100 ? ' ' + numberToWords(num % 100) : '');
  if (num < 100000) return numberToWords(Math.floor(num / 1000)) + ' Thousand' + (num % 1000 ? ' ' + numberToWords(num % 1000) : '');
  if (num < 10000000) return numberToWords(Math.floor(num / 100000)) + ' Lakh' + (num % 100000 ? ' ' + numberToWords(num % 100000) : '');
  return numberToWords(Math.floor(num / 10000000)) + ' Crore' + (num % 10000000 ? ' ' + numberToWords(num % 10000000) : '');
};

const formatDate = (dateStr) => {
  if (!dateStr || dateStr === null || dateStr === undefined) {
    return '-';
  }
  return String(dateStr);
};

const extractSerialNumberSequence = (serialNumber) => {
  if (!serialNumber || serialNumber.trim() === '') return '';
  const parts = String(serialNumber).split('/');
  if (parts.length > 0) {
    return parts[parts.length - 1].trim();
  }
  return '';
};

const generateInvoiceNumber = (counter, serialNumber) => {
  const sequence = extractSerialNumberSequence(serialNumber);
  if (sequence) {
    counter.count = parseInt(sequence) || counter.count;
  }
  const year = new Date().getFullYear();
  const invoiceNum = String(counter.count).padStart(3, '0');
  counter.count++;
  return `ELLEN/UPI/${year}/${invoiceNum}`;
};

const generateVoucherNumber = (partyName, index, payeeNameMap) => {
  if (!partyName || partyName === 'N/A') {
    partyName = 'UNKNOWN';
  }
  const key = partyName.toUpperCase().substring(0, 10).replace(/[^A-Z0-9]/g, '');
  if (!payeeNameMap[key]) {
    payeeNameMap[key] = 0;
  }
  payeeNameMap[key]++;
  const year = new Date().getFullYear();
  return `ELLEN/PV/${year}/${String(payeeNameMap[key]).padStart(2, '0')}`;
};

const detectPaymentMode = (description, refNo) => {
  if (!description) description = '';
  if (!refNo) refNo = '';
  
  const descLower = String(description).toLowerCase();
  const refLower = String(refNo).toLowerCase();
  
  if (descLower.includes('upi') || refLower.includes('upi')) return 'UPI';
  if (descLower.includes('neft') || descLower.includes('rtgs') || refLower.includes('neft') || refLower.includes('rtgs')) return 'RTGS/NEFT';
  if (descLower.includes('cheque') || descLower.includes('cheque') || refLower.includes('cheque')) return 'Cheque';
  if (descLower.includes('debit card') || descLower.includes('card') || descLower.includes('pos')) return 'Debit Card';
  if (descLower.includes('credit card')) return 'Credit Card';
  return 'RTGS/NEFT'; // Default for vendor invoices
};

// Vendor beneficiary database - matches vendor names to bank details
const vendorBeneficiaryDatabase = {
  'Arshyaa Deegital Marketing Agency LLP': {
    bankAccount: '925020028194859',
    bankName: 'Axis Bank, Nandanvan',
    description: 'Payment for Digital Marketing Services – Meta & Google Campaign Management for LearnsConnect Franchise & Student Promotions (Month: Nov 2025)',
    category: 'Marketing & Advertising'
  },
  'Maa Furniture Store': {
    bankAccount: '924020036833394',
    bankName: 'Axis Bank, Surya Nagar',
    description: 'Payment for Supply of Furniture and Interior Setup for LearnsConnect Centers – Batch Deliveries (Location-wise as per Work Order for Nov 2025)',
    category: 'Infrastructure & Furnishings'
  },
  'Selvaanand Solutions Private Limited': {
    bankAccount: '925020042321084',
    bankName: 'Axis Bank, Hoshangabad Road, Bhopal',
    description: 'Payment for Software, CRM & Portal Maintenance Services – LearnsConnect Learning & Franchise Management System (Project Phase II, Nov 2025)',
    category: 'Software Development & Maintenance'
  },
  'Preet S Earth Movers': {
    bankAccount: '925020008797504',
    bankName: 'Axis Bank, Hoshangabad Road, Bhopal',
    description: 'Payment for Logistics & Equipment Handling Services for Dispatch and Center Setup Activities – LearnsConnect Network (Nov 2025)',
    category: 'Transport & Logistics'
  },
  'Earthlygoods Agro Private Limited': {
    bankAccount: '201035788730',
    bankName: 'IndusInd Bank, Calcutta',
    description: 'Payment for Procurement & Supply Chain Support – Learning Kits, Stationery & IT Materials',
    category: 'Procurement & Operations'
  },
  'Banavat Constructions Private Limited': {
    bankAccount: '259211081730',
    bankName: 'IndusInd Bank, Preet Vihar',
    description: 'Payment for Civil & Electrical Setup Work at LearnsConnect Training Centers (Work Ref: LC/TN/2025/Phase-II)',
    category: 'Construction & Site Development'
  },
  'Extreame Horizons Tours and Travels': {
    bankAccount: '258459188735',
    bankName: 'IndusInd Bank, Kasturba Road, Borivali East',
    description: 'Payment for Travel, Coordination & Event Management Services – Franchise Onboarding & Staff Training Programs (Nov 2025)',
    category: 'Travel & Coordination'
  },
  'Unique Vision Multiservices Private Limited': {
    bankAccount: '921020006021559',
    bankName: 'Axis Bank, Gondia',
    description: 'Payment for Printing, Branding & Promotional Material Design – LearnsConnect Franchise Marketing Collaterals (Batch ID: Q4-2025)',
    category: 'Printing & Marketing Collateral'
  },
  'Golden Pettle Trading Private Limited': {
    bankAccount: '201035783063',
    bankName: 'IndusInd Bank, Hingna Road, Nagpur',
    description: 'Payment for Procurement of Electronic Accessories, Network Tools, and Support Materials for Franchise Setup & Equipment Maintenance (Work Order for Nov25)',
    category: 'Procurement & Technical Supplies'
  },
  'Elight Property Solution': {
    bankAccount: '925020024559766',
    bankName: 'Axis Bank, Ashoka Garden, Bhopal',
    description: 'Payment for Office Rental Assistance, Facility Management, and Property Maintenance Services (as per Lease/Service Agreement – Nov 2025)',
    category: 'Property & Facilities'
  }
};

// Function to match vendor name and get beneficiary details
const getVendorDetails = (vendorName) => {
  if (!vendorName) return null;
  
  // Try exact match first
  const exactMatch = vendorBeneficiaryDatabase[vendorName];
  if (exactMatch) return exactMatch;
  
  // Try partial match (case insensitive)
  const vendorNameLower = vendorName.toLowerCase().trim();
  for (const [key, value] of Object.entries(vendorBeneficiaryDatabase)) {
    if (key.toLowerCase().includes(vendorNameLower) || vendorNameLower.includes(key.toLowerCase())) {
      return value;
    }
  }
  
  return null;
};

// Extract bank account from description or row data
const extractBankAccountFromData = (row, description) => {
  // Try direct bank account fields first
  const bankAccountKeys = ['bankaccount', 'accountno', 'accountnumber', 'bankacno', 'acno', 'bankac', 'refnochequeno', 'refno', 'chequeno'];
  const bankAccountKey = Object.keys(row).find(key => {
    const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
    return bankAccountKeys.some(bk => {
      const normalizedBk = bk.toLowerCase().replace(/[^a-z0-9]/g, '');
      return normalizedKey.includes(normalizedBk) || normalizedBk.includes(normalizedKey);
    });
  });
  
  if (bankAccountKey && row[bankAccountKey]) {
    const accountValue = String(row[bankAccountKey]).trim();
    // Check if it looks like an account number (has digits, at least 10 digits)
    if (accountValue && /^\d{10,}$/.test(accountValue.replace(/[^\d]/g, ''))) {
      // Return only digits (remove any non-digit characters)
      return accountValue.replace(/[^\d]/g, '');
    }
    // If it's a valid account number format, return as is
    if (accountValue && accountValue.length >= 10) {
      return accountValue;
    }
  }
  
  // Try to extract from description
  if (description) {
    // Look for patterns like "A/c No: 123456" or "Account: 123456"
    const accountPatterns = [
      /A\/c[\s]*[No\.]*[\s]*[:]*[\s]*(\d{10,})/i,
      /Account[\s]*[No\.]*[\s]*[:]*[\s]*(\d{10,})/i,
      /Bank[\s]*A\/c[\s]*[No\.]*[\s]*[:]*[\s]*(\d{10,})/i,
      /\b(\d{12,})\b/g  // 12 or more digits
    ];
    
    for (const pattern of accountPatterns) {
      const match = description.match(pattern);
      if (match && match[1]) {
        return match[1].trim();
      }
    }
  }
  
  return null;
};

// Extract bank name from description or row data
const extractBankNameFromData = (row, description) => {
  // Try direct bank name fields
  const bankNameKeys = ['bankname', 'bank', 'bankbranch', 'bankdetails', 'branchname'];
  const bankNameKey = Object.keys(row).find(key => 
    bankNameKeys.some(bk => key.toLowerCase().includes(bk.toLowerCase()))
  );
  
  if (bankNameKey && row[bankNameKey]) {
    return String(row[bankNameKey]).trim();
  }
  
  // Try to extract from description
  if (description) {
    // Look for bank names
    const bankPatterns = [
      /(Axis Bank[^,]*)/i,
      /(IndusInd Bank[^,]*)/i,
      /(SBI[^,]*)/i,
      /(HDFC Bank[^,]*)/i,
      /(ICICI Bank[^,]*)/i
    ];
    
    for (const pattern of bankPatterns) {
      const match = description.match(pattern);
      if (match && match[1]) {
        return match[1].trim();
      }
    }
  }
  
  return null;
};

// Extract category from row data
const extractCategoryFromData = (row) => {
  const categoryKeys = ['category', 'purposecategory', 'paymentcategory', 'type'];
  const categoryKey = Object.keys(row).find(key => 
    categoryKeys.some(ck => key.toLowerCase().includes(ck.toLowerCase()))
  );
  
  if (categoryKey && row[categoryKey]) {
    return String(row[categoryKey]).trim();
  }
  
  return null;
};

// Load logo as base64
const loadLogoBase64 = async () => {
  try {
    const logoResponse = await fetch('/src/assets/images/logo.jpg');
    if (logoResponse.ok) {
      const blob = await logoResponse.blob();
      return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.readAsDataURL(blob);
      });
    }
  } catch (error) {
    try {
      const altResponse = await fetch('/assets/images/logo.jpg');
      if (altResponse.ok) {
        const blob = await altResponse.blob();
        return new Promise((resolve) => {
          const reader = new FileReader();
          reader.onloadend = () => resolve(reader.result);
          reader.readAsDataURL(blob);
        });
      }
    } catch (err) {
      try {
        const pubResponse = await fetch('/logo.jpg');
        if (pubResponse.ok) {
          const blob = await pubResponse.blob();
          return new Promise((resolve) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result);
            reader.readAsDataURL(blob);
          });
        }
      } catch (e) {
        console.warn('Logo not found');
      }
    }
  }
  return null;
};

export const exportToPDF = async (data, billingType, transactionType = 'all', filename = 'invoice-report.pdf') => {
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

  // Load logo
  const logoBase64 = await loadLogoBase64();

  // Create PDF document (single document for all records)
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'mm',
    format: 'a4'
  });

  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();

  // Process each record
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const amount = getAmount(row);
    const amountInWords = numberToWords(Math.floor(amount));
    
    // Get serial number
    const serialNoKeys = ['serialno', 'serialnumber', 'serial no', 's.no', 'sno'];
    const serialNoKey = Object.keys(row).find(key => 
      serialNoKeys.some(sk => key.toLowerCase().replace(/[^a-z0-9]/g, '').includes(sk.toLowerCase().replace(/[^a-z0-9]/g, '')))
    );
    const serialNumber = serialNoKey ? String(row[serialNoKey] || '').trim() : '';

    // Get date
    const dateKeys = ['date', 'txndate', 'transactiondate', 'paymentdate', 'valuedate', 'transaction date', 'value date'];
    let dateKey = Object.keys(row).find(key => 
      dateKeys.some(dk => key.toLowerCase().includes(dk.toLowerCase()))
    );
    
    if (!dateKey) {
      for (const key of Object.keys(row)) {
        const value = String(row[key] || '').trim();
        if (value.match(/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/)) {
          dateKey = key;
          break;
        }
      }
    }
    
    const invoiceDate = formatDate(row[dateKey] || null);

    // Get description
    const descKeys = ['description', 'particulars', 'narration', 'details', 'remarks', 'note'];
    let descKey = Object.keys(row).find(key => 
      descKeys.some(dk => key.toLowerCase().includes(dk.toLowerCase()))
    );
    
    if (!descKey) {
      let maxLength = 0;
      for (const key of Object.keys(row)) {
        const value = String(row[key] || '').trim();
        if (value && 
            !value.match(/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/) && 
            !value.match(/^\-?\d+[.,]?\d*$/) && 
            !value.match(/^[A-Z]+\/\d+\/\d+\/\d+$/) &&
            !value.match(/^[\/\-\s]+$/)) {
          if (value.length > maxLength && value.length > 10) {
            maxLength = value.length;
            descKey = key;
          }
        }
      }
    }
    
    const description = descKey ? String(row[descKey] || '').trim() : '';

    // Extract party name
    let partyName = 'N/A';
    const nameKeys = ['name', 'partyname', 'client', 'payer', 'studentname', 'payeename', 'vendorname'];
    const nameKey = Object.keys(row).find(key => 
      nameKeys.some(nk => key.toLowerCase().includes(nk.toLowerCase()))
    );
    
    if (nameKey && row[nameKey] && String(row[nameKey]).trim().length > 0) {
      partyName = String(row[nameKey]).trim();
    } else if (description) {
      const words = description.split(/[\/\s]+/).filter(w => w.length > 2 && !/^\d+$/.test(w));
      if (words.length > 0) {
        partyName = words.slice(0, 3).join(' ').trim();
      }
    }

    const isCredit = isCreditTransaction(row);
    
    // For debit transactions, use serial number in voucher number
    let invoiceNumber;
    if (isCredit) {
      invoiceNumber = generateInvoiceNumber(invoiceCounter, serialNumber);
    } else {
      // For vendor invoice, use format: ELLEN/PV/2025/(serial number)
      const year = new Date().getFullYear();
      const serialSequence = extractSerialNumberSequence(serialNumber);
      invoiceNumber = serialSequence 
        ? `ELLEN/PV/${year}/${serialSequence}`
        : generateVoucherNumber(partyName, i, payeeNameMap);
    }

    const refKeys = ['refno', 'refno', 'ref', 'chequeno', 'utrno', 'utr', 'branchcode'];
    const refKey = Object.keys(row).find(key => 
      refKeys.some(rk => key.toLowerCase().includes(rk.toLowerCase()))
    );
    let refNo = refKey ? String(row[refKey] || '') : '';
    
    const paymentMode = detectPaymentMode(description, refNo);
    
    // Get vendor details from database
    const vendorDetails = getVendorDetails(partyName);
    
    // Extract bank account and bank name from data or vendor database
    let bankAccount = extractBankAccountFromData(row, description);
    let bankName = extractBankNameFromData(row, description);
    
    // Use vendor database if available
    if (vendorDetails) {
      bankAccount = bankAccount || vendorDetails.bankAccount;
      bankName = bankName || vendorDetails.bankName;
    }
    
    // Extract category from data or vendor database
    let category = extractCategoryFromData(row);
    if (vendorDetails && !category) {
      category = vendorDetails.category;
    }
    
    // Use description from vendor database if available, otherwise use row description
    const finalDescription = vendorDetails ? vendorDetails.description : description;

    // Add new page if not first record
    if (i > 0) {
      pdf.addPage();
    }
    
    let yPos = 20;

    // Header with logo and contact info
    if (logoBase64) {
      try {
        pdf.addImage(logoBase64, 'JPEG', 20, yPos, 15, 15); // Reduced logo size from 20x20 to 15x15
      } catch (e) {
        console.warn('Could not add logo image');
      }
    }

    // LEARNSCONNECT branding (positioned to avoid logo overlap, size reduced)
    pdf.setFontSize(12); // Reduced from 16
    pdf.setTextColor(0, 0, 255); // Blue
    pdf.setFont('helvetica', 'bold');
    pdf.text('LEARNSCONNECT', 40, yPos + 6); // Positioned after logo (20mm + 15mm logo + 5mm gap), adjusted position

    // Tagline (positioned below LEARNSCONNECT, size reduced)
    pdf.setFontSize(9); // Reduced from 11
    pdf.setTextColor(51, 51, 51); // Dark gray
    pdf.setFont('helvetica', 'bold');
    pdf.text('WHERE LEARNING MEETS OPPORTUNITY', 40, yPos + 12); // Positioned below LEARNSCONNECT, adjusted position

    // Contact info on right (positioned to avoid overlap, emojis removed, size reduced)
    pdf.setFontSize(8); // Reduced from 9
    pdf.setTextColor(0, 0, 0);
    pdf.setFont('helvetica', 'normal');
    const contactInfo = [
      '8/3, Athreyapuram 2nd Street, Choolaimedu, Chennai - 600094',
      'Phone: +91 84893 57705',
      'Email: support@learnsconnect.com',
      'Website: www.learnsconnect.com'
    ];
    // Right align contact info, starting from top with proper spacing (reduced spacing)
    let contactY = yPos + 1;
    contactInfo.forEach((line, idx) => {
      // Split long lines if needed
      const splitLines = pdf.splitTextToSize(line, pageWidth - 110); // Leave space for logo
      splitLines.forEach((splitLine, lineIdx) => {
        pdf.text(splitLine, pageWidth - 20, contactY + (idx * 5) + (lineIdx * 5), { align: 'right' });
      });
      contactY += (splitLines.length - 1) * 5; // Adjust for multi-line text
    });
    
    // Calculate actual header end position (reduced height)
    const headerEndY = Math.max(yPos + 15, contactY + 3); // End of contact info or logo + tagline, reduced height

    // Move yPos down to avoid header overlap with content
    // Use calculated header end position and add proper spacing
    yPos = headerEndY + 15; // Add extra 15mm spacing after header to ensure no overlap

    // Company Name
    pdf.setFontSize(14);
    pdf.setFont('helvetica', 'bold');
    pdf.text(companyName, pageWidth / 2, yPos, { align: 'center' });
    yPos += 12; // Increased spacing

    // Teal separator line
    pdf.setDrawColor(0, 128, 128); // Teal
    pdf.setLineWidth(0.5);
    pdf.line(20, yPos, pageWidth - 20, yPos);
    yPos += 12; // Increased spacing

    // Title based on transaction type (Credit = Payment Receipt, Debit = Vendor Invoice)
    pdf.setFontSize(18);
    pdf.setFont('helvetica', 'bold');
    if (isCredit) {
      pdf.text('LearnsConnect – Payment Receipt', pageWidth / 2, yPos, { align: 'center' });
    } else {
      pdf.text('LearnsConnect - Vendor Invoice', pageWidth / 2, yPos, { align: 'center' });
    }
    yPos += 18; // Increased spacing before content

    if (isCredit) {
      // Payment Receipt Content - Table Format
      const lineHeight = 10; // Spacing between lines
      const leftMargin = 20;
      const rightMargin = pageWidth - 20;
      const labelWidth = 75; // Label column width
      const valueStart = leftMargin + labelWidth;
      const valueWidth = rightMargin - valueStart; // Available width for values
      const cellPadding = 3; // Padding inside cells
      const rowHeight = 12; // Minimum row height
      
      // Helper function to draw table row with borders
      const drawTableRow = (labelText, valueText, isBoldLabel = true, isBoldValue = false, isMultiLine = false) => {
        const startY = yPos;
        let currentY = yPos;
        
        // Draw label cell
        pdf.setFont('helvetica', isBoldLabel ? 'bold' : 'normal');
        pdf.setFontSize(11);
        const labelSplit = pdf.splitTextToSize(labelText, labelWidth - (cellPadding * 2));
        let labelHeight = Math.max(rowHeight, labelSplit.length * lineHeight);
        
        // Draw value cell
        pdf.setFont('helvetica', isBoldValue ? 'bold' : 'normal');
        const valueSplit = pdf.splitTextToSize(valueText, valueWidth - (cellPadding * 2));
        let valueHeight = Math.max(rowHeight, valueSplit.length * lineHeight);
        
        // Use the maximum height for both cells
        const cellHeight = Math.max(labelHeight, valueHeight);
        
        // Draw cell borders
        pdf.setDrawColor(0, 0, 0); // Black borders
        pdf.setLineWidth(0.1);
        
        // Left border
        pdf.line(leftMargin, startY, leftMargin, startY + cellHeight);
        // Right border
        pdf.line(rightMargin, startY, rightMargin, startY + cellHeight);
        // Top border
        pdf.line(leftMargin, startY, rightMargin, startY);
        // Bottom border
        pdf.line(leftMargin, startY + cellHeight, rightMargin, startY + cellHeight);
        // Middle border (between label and value)
        pdf.line(valueStart, startY, valueStart, startY + cellHeight);
        
        // Draw label text
        pdf.setFont('helvetica', isBoldLabel ? 'bold' : 'normal');
        pdf.setFontSize(11);
        labelSplit.forEach((line, idx) => {
          pdf.text(line, leftMargin + cellPadding, currentY + (idx * lineHeight) + 5);
        });
        
        // Draw value text
        pdf.setFont('helvetica', isBoldValue ? 'bold' : 'normal');
        pdf.setFontSize(11);
        valueSplit.forEach((line, idx) => {
          pdf.text(line, valueStart + cellPadding, currentY + (idx * lineHeight) + 5);
        });
        
        yPos = startY + cellHeight;
      };

      // Payment Receipt Voucher No.
      drawTableRow('Payment Receipt Voucher No.:', invoiceNumber || '-', true, false);

      // Date
      drawTableRow('Date:', invoiceDate || '-', true, false);

      // Mode of Payment
      drawTableRow('Mode of Payment:', 'UPI', true, false);

      // Received from Payer
      const payerDetails = description || partyName || '-';
      drawTableRow('Received from Payer (Name and Details):', payerDetails, true, false, true);

      // Purpose of Payment
      drawTableRow('Purpose of Payment:', 'Payment collected from students for Course/Batch', true, false, true);

      // Amount
      const amountText = `INR ${amount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
      drawTableRow('Amount (INR):', amountText, true, true);

      // Amount in Words
      drawTableRow('Amount in Words:', amountInWords, true, false, true);
      
      yPos += 5; // Extra spacing after table

      // Footer section - position at bottom of first page
      const footerYPos = pageHeight - 30; // Position footer near bottom of page

      // Footer Disclaimer
      pdf.setFontSize(9);
      pdf.setFont('helvetica', 'italic');
      pdf.setTextColor(128, 128, 128); // Gray
      const disclaimerText = 'This receipt is computer generated by for acknowledgement of payment received. Non-refundable unless of technical issues. All disputes subject to Chennai jurisdiction.';
      const disclaimerSplit = pdf.splitTextToSize(disclaimerText, pageWidth - 40);
      pdf.text(disclaimerSplit, 20, footerYPos, { align: 'left' });
      const footerTextHeight = disclaimerSplit.length * 5 + 3; // Approximate height
      
      // Bottom Teal Border Line
      pdf.setDrawColor(0, 128, 128); // Teal
      pdf.setLineWidth(0.5);
      pdf.line(20, footerYPos + footerTextHeight, pageWidth - 20, footerYPos + footerTextHeight);
      
      // Update yPos to footer position for proper spacing
      yPos = footerYPos + footerTextHeight + 5;
    } else {
      // Vendor Invoice format (table structure matching image template)
      const lineHeight = 10; // Spacing between lines
      const leftMargin = 20;
      const rightMargin = pageWidth - 20;
      const labelWidth = 75; // Label column width
      const valueStart = leftMargin + labelWidth;
      const valueWidth = rightMargin - valueStart; // Available width for values
      const cellPadding = 3; // Padding inside cells
      const rowHeight = 12; // Minimum row height
      
      // Helper function to draw table row with borders
      const drawTableRow = (labelText, valueText, isBoldLabel = true, isBoldValue = false, isMultiLine = false) => {
        const startY = yPos;
        let currentY = yPos;
        
        // Draw label cell
        pdf.setFont('helvetica', isBoldLabel ? 'bold' : 'normal');
        pdf.setFontSize(11);
        const labelSplit = pdf.splitTextToSize(labelText, labelWidth - (cellPadding * 2));
        let labelHeight = Math.max(rowHeight, labelSplit.length * lineHeight);
        
        // Draw value cell
        pdf.setFont('helvetica', isBoldValue ? 'bold' : 'normal');
        const valueSplit = pdf.splitTextToSize(valueText, valueWidth - (cellPadding * 2));
        let valueHeight = Math.max(rowHeight, valueSplit.length * lineHeight);
        
        // Use the maximum height for both cells
        const cellHeight = Math.max(labelHeight, valueHeight);
        
        // Draw cell borders
        pdf.setDrawColor(0, 0, 0); // Black borders
        pdf.setLineWidth(0.1);
        
        // Left border
        pdf.line(leftMargin, startY, leftMargin, startY + cellHeight);
        // Right border
        pdf.line(rightMargin, startY, rightMargin, startY + cellHeight);
        // Top border
        pdf.line(leftMargin, startY, rightMargin, startY);
        // Bottom border
        pdf.line(leftMargin, startY + cellHeight, rightMargin, startY + cellHeight);
        // Middle border (between label and value)
        pdf.line(valueStart, startY, valueStart, startY + cellHeight);
        
        // Draw label text
        pdf.setFont('helvetica', isBoldLabel ? 'bold' : 'normal');
        pdf.setFontSize(11);
        labelSplit.forEach((line, idx) => {
          pdf.text(line, leftMargin + cellPadding, currentY + (idx * lineHeight) + 5);
        });
        
        // Draw value text
        pdf.setFont('helvetica', isBoldValue ? 'bold' : 'normal');
        pdf.setFontSize(11);
        valueSplit.forEach((line, idx) => {
          pdf.text(line, valueStart + cellPadding, currentY + (idx * lineHeight) + 5);
        });
        
        yPos = startY + cellHeight;
      };

              // Vendor Invoice fields
              drawTableRow('Voucher No.:', invoiceNumber || '-', true, false);
              drawTableRow('Date:', invoiceDate || '-', true, false);
              drawTableRow('Mode of Payment:', paymentMode || 'RTGS/NEFT', true, false);
              drawTableRow('Bank Account No:', bankAccount || '-', true, false);
              drawTableRow('Bank Name and Details:', bankName || '-', true, false);
              // Purpose of Payment - always use default value
              const purposeText = 'Payment collected from students is paid to tutors/institutions';
              drawTableRow('Purpose of Payment:', purposeText, true, false, true);
      const amountText = `INR ${amount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
      drawTableRow('Amount (INR):', amountText, true, true);
      drawTableRow('Amount in Words:', amountInWords, true, false, true);
      
      yPos += 10; // Extra spacing after main table
      
      // Footer section - position at bottom of first page
      const footerYPos = pageHeight - 30; // Position footer near bottom of page
      
      // Footer disclaimer
      pdf.setFontSize(9);
      pdf.setFont('helvetica', 'italic');
      pdf.setTextColor(0, 0, 0);
      const footerText = 'All payments are made via SBI RTGS/NEFT from Ellen Information Technology Solutions Pvt. Ltd. (A/c No. ending 5037) towards business services rendered under the LearnsConnect operations. Each transfer is supported by work orders, GST invoices, and digital approval records.';
      const footerSplit = pdf.splitTextToSize(footerText, pageWidth - 40);
      pdf.text(footerSplit, leftMargin, footerYPos, { align: 'left' });
      const footerTextHeight = footerSplit.length * 5 + 3;
      
      // Bottom teal border line
      pdf.setDrawColor(0, 128, 128); // Teal
      pdf.setLineWidth(0.5);
      pdf.line(leftMargin, footerYPos + footerTextHeight, pageWidth - 20, footerYPos + footerTextHeight);
      
      // Update yPos to footer position for proper spacing
      yPos = footerYPos + footerTextHeight + 5;
    }

    // Save PDF after processing all records
    if (i === data.length - 1) {
      pdf.save(filename);
      window.pdfDoc = null; // Clean up
    }
  }
};

// Helper function to generate PDF blob for ZIP export (without saving)
export const generatePDFBlob = async (data, billingType, transactionType = 'all') => {
  if (!data || data.length === 0) {
    return null;
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

  // Load logo
  const logoBase64 = await loadLogoBase64();

  // Create PDF document (single document for all records)
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'mm',
    format: 'a4'
  });

  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();

  // Process each record (same logic as exportToPDF)
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const amount = getAmount(row);
    const amountInWords = numberToWords(Math.floor(amount));
    
    // Get serial number
    const serialNoKeys = ['serialno', 'serialnumber', 'serial no', 's.no', 'sno'];
    const serialNoKey = Object.keys(row).find(key => 
      serialNoKeys.some(sk => key.toLowerCase().replace(/[^a-z0-9]/g, '').includes(sk.toLowerCase().replace(/[^a-z0-9]/g, '')))
    );
    const serialNumber = serialNoKey ? String(row[serialNoKey] || '').trim() : '';

    // Get date
    const dateKeys = ['date', 'txndate', 'transactiondate', 'paymentdate', 'valuedate', 'transaction date', 'value date'];
    let dateKey = Object.keys(row).find(key => 
      dateKeys.some(dk => key.toLowerCase().includes(dk.toLowerCase()))
    );
    
    if (!dateKey) {
      for (const key of Object.keys(row)) {
        const value = String(row[key] || '').trim();
        if (value.match(/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/)) {
          dateKey = key;
          break;
        }
      }
    }
    
    const invoiceDate = formatDate(row[dateKey] || null);

    // Get description
    const descKeys = ['description', 'particulars', 'narration', 'details', 'remarks', 'note'];
    let descKey = Object.keys(row).find(key => 
      descKeys.some(dk => key.toLowerCase().includes(dk.toLowerCase()))
    );
    
    if (!descKey) {
      let maxLength = 0;
      for (const key of Object.keys(row)) {
        const value = String(row[key] || '').trim();
        if (value && 
            !value.match(/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/) && 
            !value.match(/^\-?\d+[.,]?\d*$/) && 
            !value.match(/^[A-Z]+\/\d+\/\d+\/\d+$/) &&
            !value.match(/^[\/\-\s]+$/)) {
          if (value.length > maxLength && value.length > 10) {
            maxLength = value.length;
            descKey = key;
          }
        }
      }
    }
    
    const description = descKey ? String(row[descKey] || '').trim() : '';

    // Extract party name
    let partyName = 'N/A';
    const nameKeys = ['name', 'partyname', 'client', 'payer', 'studentname', 'payeename', 'vendorname'];
    const nameKey = Object.keys(row).find(key => 
      nameKeys.some(nk => key.toLowerCase().includes(nk.toLowerCase()))
    );
    
    if (nameKey && row[nameKey] && String(row[nameKey]).trim().length > 0) {
      partyName = String(row[nameKey]).trim();
    } else if (description) {
      const words = description.split(/[\/\s]+/).filter(w => w.length > 2 && !/^\d+$/.test(w));
      if (words.length > 0) {
        partyName = words.slice(0, 3).join(' ').trim();
      }
    }

    const isCredit = isCreditTransaction(row);
    
    // For debit transactions, use serial number in voucher number
    let invoiceNumber;
    if (isCredit) {
      invoiceNumber = generateInvoiceNumber(invoiceCounter, serialNumber);
    } else {
      // For vendor invoice, use format: ELLEN/PV/2025/(serial number)
      const year = new Date().getFullYear();
      const serialSequence = extractSerialNumberSequence(serialNumber);
      invoiceNumber = serialSequence 
        ? `ELLEN/PV/${year}/${serialSequence}`
        : generateVoucherNumber(partyName, i, payeeNameMap);
    }

    const refKeys = ['refno', 'refno', 'ref', 'chequeno', 'utrno', 'utr', 'branchcode'];
    const refKey = Object.keys(row).find(key => 
      refKeys.some(rk => key.toLowerCase().includes(rk.toLowerCase()))
    );
    let refNo = refKey ? String(row[refKey] || '') : '';
    
    const paymentMode = detectPaymentMode(description, refNo);
    
    // Get vendor details from database
    const vendorDetails = getVendorDetails(partyName);
    
    // Extract bank account and bank name from data or vendor database
    let bankAccount = extractBankAccountFromData(row, description);
    let bankName = extractBankNameFromData(row, description);
    
    // Use vendor database if available
    if (vendorDetails) {
      bankAccount = bankAccount || vendorDetails.bankAccount;
      bankName = bankName || vendorDetails.bankName;
    }
    
    // Extract category from data or vendor database
    let category = extractCategoryFromData(row);
    if (vendorDetails && !category) {
      category = vendorDetails.category;
    }
    
    // Use description from vendor database if available, otherwise use row description
    const finalDescription = vendorDetails ? vendorDetails.description : description;

    // Add new page if not first record
    if (i > 0) {
      pdf.addPage();
    }
    
    let yPos = 20;

    // Header with logo and contact info (same as exportToPDF)
    if (logoBase64) {
      try {
        pdf.addImage(logoBase64, 'JPEG', 20, yPos, 15, 15);
      } catch (e) {
        console.warn('Could not add logo image');
      }
    }

    pdf.setFontSize(12);
    pdf.setTextColor(0, 0, 255);
    pdf.setFont('helvetica', 'bold');
    pdf.text('LEARNSCONNECT', 40, yPos + 6);

    pdf.setFontSize(9);
    pdf.setTextColor(51, 51, 51);
    pdf.setFont('helvetica', 'bold');
    pdf.text('WHERE LEARNING MEETS OPPORTUNITY', 40, yPos + 12);

    pdf.setFontSize(8);
    pdf.setTextColor(0, 0, 0);
    pdf.setFont('helvetica', 'normal');
    const contactInfo = [
      '8/3, Athreyapuram 2nd Street, Choolaimedu, Chennai - 600094',
      'Phone: +91 84893 57705',
      'Email: support@learnsconnect.com',
      'Website: www.learnsconnect.com'
    ];
    let contactY = yPos + 1;
    contactInfo.forEach((line, idx) => {
      const splitLines = pdf.splitTextToSize(line, pageWidth - 110);
      splitLines.forEach((splitLine, lineIdx) => {
        pdf.text(splitLine, pageWidth - 20, contactY + (idx * 5) + (lineIdx * 5), { align: 'right' });
      });
      contactY += (splitLines.length - 1) * 5;
    });
    
    const headerEndY = Math.max(yPos + 15, contactY + 3);
    yPos = headerEndY + 15;

    pdf.setFontSize(14);
    pdf.setFont('helvetica', 'bold');
    pdf.text(companyName, pageWidth / 2, yPos, { align: 'center' });
    yPos += 12;

    pdf.setDrawColor(0, 128, 128);
    pdf.setLineWidth(0.5);
    pdf.line(20, yPos, pageWidth - 20, yPos);
    yPos += 12;

    pdf.setFontSize(18);
    pdf.setFont('helvetica', 'bold');
    if (isCredit) {
      pdf.text('LearnsConnect – Payment Receipt', pageWidth / 2, yPos, { align: 'center' });
    } else {
      pdf.text('LearnsConnect - Vendor Invoice', pageWidth / 2, yPos, { align: 'center' });
    }
    yPos += 18;

    if (isCredit) {
      // Payment Receipt Content (same as exportToPDF)
      const lineHeight = 10;
      const leftMargin = 20;
      const rightMargin = pageWidth - 20;
      const labelWidth = 75;
      const valueStart = leftMargin + labelWidth;
      const valueWidth = rightMargin - valueStart;
      const cellPadding = 3;
      const rowHeight = 12;
      
      const drawTableRow = (labelText, valueText, isBoldLabel = true, isBoldValue = false, isMultiLine = false) => {
        const startY = yPos;
        
        pdf.setFont('helvetica', isBoldLabel ? 'bold' : 'normal');
        pdf.setFontSize(11);
        const labelSplit = pdf.splitTextToSize(labelText, labelWidth - (cellPadding * 2));
        let labelHeight = Math.max(rowHeight, labelSplit.length * lineHeight);
        
        pdf.setFont('helvetica', isBoldValue ? 'bold' : 'normal');
        const valueSplit = pdf.splitTextToSize(valueText, valueWidth - (cellPadding * 2));
        let valueHeight = Math.max(rowHeight, valueSplit.length * lineHeight);
        
        const cellHeight = Math.max(labelHeight, valueHeight);
        
        pdf.setDrawColor(0, 0, 0);
        pdf.setLineWidth(0.1);
        pdf.line(leftMargin, startY, leftMargin, startY + cellHeight);
        pdf.line(rightMargin, startY, rightMargin, startY + cellHeight);
        pdf.line(leftMargin, startY, rightMargin, startY);
        pdf.line(leftMargin, startY + cellHeight, rightMargin, startY + cellHeight);
        pdf.line(valueStart, startY, valueStart, startY + cellHeight);
        
        pdf.setFont('helvetica', isBoldLabel ? 'bold' : 'normal');
        pdf.setFontSize(11);
        labelSplit.forEach((line, idx) => {
          pdf.text(line, leftMargin + cellPadding, startY + (idx * lineHeight) + 5);
        });
        
        pdf.setFont('helvetica', isBoldValue ? 'bold' : 'normal');
        pdf.setFontSize(11);
        valueSplit.forEach((line, idx) => {
          pdf.text(line, valueStart + cellPadding, startY + (idx * lineHeight) + 5);
        });
        
        yPos = startY + cellHeight;
      };

      drawTableRow('Payment Receipt Voucher No.:', invoiceNumber || '-', true, false);
      drawTableRow('Date:', invoiceDate || '-', true, false);
      drawTableRow('Mode of Payment:', 'UPI', true, false);
      const payerDetails = description || partyName || '-';
      drawTableRow('Received from Payer (Name and Details):', payerDetails, true, false, true);
      drawTableRow('Purpose of Payment:', 'Payment collected from students for Course/Batch', true, false, true);
      const amountText = `INR ${amount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
      drawTableRow('Amount (INR):', amountText, true, true);
      drawTableRow('Amount in Words:', amountInWords, true, false, true);
      
      yPos += 5;

      const footerYPos = pageHeight - 30;
      pdf.setFontSize(9);
      pdf.setFont('helvetica', 'italic');
      pdf.setTextColor(128, 128, 128);
      const disclaimerText = 'This receipt is computer generated by for acknowledgement of payment received. Non-refundable unless of technical issues. All disputes subject to Chennai jurisdiction.';
      const disclaimerSplit = pdf.splitTextToSize(disclaimerText, pageWidth - 40);
      pdf.text(disclaimerSplit, 20, footerYPos, { align: 'left' });
      const footerTextHeight = disclaimerSplit.length * 5 + 3;
      
      pdf.setDrawColor(0, 128, 128);
      pdf.setLineWidth(0.5);
      pdf.line(20, footerYPos + footerTextHeight, pageWidth - 20, footerYPos + footerTextHeight);
      
      yPos = footerYPos + footerTextHeight + 5;
    } else {
      // Vendor Invoice format (same as exportToPDF)
      const lineHeight = 10;
      const leftMargin = 20;
      const rightMargin = pageWidth - 20;
      const labelWidth = 75;
      const valueStart = leftMargin + labelWidth;
      const valueWidth = rightMargin - valueStart;
      const cellPadding = 3;
      const rowHeight = 12;
      
      const drawTableRow = (labelText, valueText, isBoldLabel = true, isBoldValue = false, isMultiLine = false) => {
        const startY = yPos;
        
        pdf.setFont('helvetica', isBoldLabel ? 'bold' : 'normal');
        pdf.setFontSize(11);
        const labelSplit = pdf.splitTextToSize(labelText, labelWidth - (cellPadding * 2));
        let labelHeight = Math.max(rowHeight, labelSplit.length * lineHeight);
        
        pdf.setFont('helvetica', isBoldValue ? 'bold' : 'normal');
        const valueSplit = pdf.splitTextToSize(valueText, valueWidth - (cellPadding * 2));
        let valueHeight = Math.max(rowHeight, valueSplit.length * lineHeight);
        
        const cellHeight = Math.max(labelHeight, valueHeight);
        
        pdf.setDrawColor(0, 0, 0);
        pdf.setLineWidth(0.1);
        pdf.line(leftMargin, startY, leftMargin, startY + cellHeight);
        pdf.line(rightMargin, startY, rightMargin, startY + cellHeight);
        pdf.line(leftMargin, startY, rightMargin, startY);
        pdf.line(leftMargin, startY + cellHeight, rightMargin, startY + cellHeight);
        pdf.line(valueStart, startY, valueStart, startY + cellHeight);
        
        pdf.setFont('helvetica', isBoldLabel ? 'bold' : 'normal');
        pdf.setFontSize(11);
        labelSplit.forEach((line, idx) => {
          pdf.text(line, leftMargin + cellPadding, startY + (idx * lineHeight) + 5);
        });
        
        pdf.setFont('helvetica', isBoldValue ? 'bold' : 'normal');
        pdf.setFontSize(11);
        valueSplit.forEach((line, idx) => {
          pdf.text(line, valueStart + cellPadding, startY + (idx * lineHeight) + 5);
        });
        
        yPos = startY + cellHeight;
      };

      drawTableRow('Voucher No.:', invoiceNumber || '-', true, false);
      drawTableRow('Date:', invoiceDate || '-', true, false);
      drawTableRow('Mode of Payment:', paymentMode || 'RTGS/NEFT', true, false);
      drawTableRow('Bank Account No:', bankAccount || '-', true, false);
      drawTableRow('Bank Name and Details:', bankName || '-', true, false);
      const purposeText = 'Payment collected from students is paid to tutors/institutions';
      drawTableRow('Purpose of Payment:', purposeText, true, false, true);
      const amountText = `INR ${amount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
      drawTableRow('Amount (INR):', amountText, true, true);
      drawTableRow('Amount in Words:', amountInWords, true, false, true);
      
      yPos += 10;
      
      const footerYPos = pageHeight - 30;
      pdf.setFontSize(9);
      pdf.setFont('helvetica', 'italic');
      pdf.setTextColor(0, 0, 0);
      const footerText = 'All payments are made via SBI RTGS/NEFT from Ellen Information Technology Solutions Pvt. Ltd. (A/c No. ending 5037) towards business services rendered under the LearnsConnect operations. Each transfer is supported by work orders, GST invoices, and digital approval records.';
      const footerSplit = pdf.splitTextToSize(footerText, pageWidth - 40);
      pdf.text(footerSplit, leftMargin, footerYPos, { align: 'left' });
      const footerTextHeight = footerSplit.length * 5 + 3;
      
      pdf.setDrawColor(0, 128, 128);
      pdf.setLineWidth(0.5);
      pdf.line(leftMargin, footerYPos + footerTextHeight, pageWidth - 20, footerYPos + footerTextHeight);
      
      yPos = footerYPos + footerTextHeight + 5;
    }
  }

  // Return PDF as array buffer for ZIP
  return pdf.output('arraybuffer');
};

