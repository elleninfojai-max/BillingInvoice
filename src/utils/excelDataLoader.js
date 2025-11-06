// Hardcoded beneficiary database from Excel data
const BENEFICIARY_DATABASE = [
  {
    'Vendor Name': 'Arshyaa Deegital Marketing Agency LLP',
    'Bank A/c No.': '925020028194859',
    'Bank Name & Branch (Verified)': 'Axis Bank, Nandanvan',
    'Description': 'Payment for Digital Marketing Services – Meta & Google Campaign Management for LearnsConnect Franchise & Student Promotions (Month: Nov 2025)',
    'Purpose / Category': 'Marketing & Advertising',
    'Mode of Payment': 'RTGS/NEFT',
    'Date': '11 Mar 2025'
  },
  {
    'Vendor Name': 'Maa Furniture Store',
    'Bank A/c No.': '924020036833394',
    'Bank Name & Branch (Verified)': 'Axis Bank, Surya Nagar',
    'Description': 'Payment for Supply of Furniture and Interior Setup for LearnsConnect Centers – Batch Deliveries (Location-wise as per Work Order for Nov 2025)',
    'Purpose / Category': 'Infrastructure & Furnishings',
    'Mode of Payment': 'RTGS/NEFT',
    'Date': '11 Mar 2025'
  },
  {
    'Vendor Name': 'Selvaanand Solutions Private Limited',
    'Bank A/c No.': '925020042321084',
    'Bank Name & Branch (Verified)': 'Axis Bank, Hoshangabad Road, Bhopal',
    'Description': 'Payment for Software, CRM & Portal Maintenance Services – LearnsConnect Learning & Franchise Management System (Project Phase II, Nov 2025)',
    'Purpose / Category': 'Software Development & Maintenance',
    'Mode of Payment': 'RTGS/NEFT',
    'Date': '11 Mar 2025'
  },
  {
    'Vendor Name': 'Preet S Earth Movers',
    'Bank A/c No.': '925020008797504',
    'Bank Name & Branch (Verified)': 'Axis Bank, Hoshangabad Road, Bhopal',
    'Description': 'Payment for Logistics & Equipment Handling Services for Dispatch and Center Setup Activities – LearnsConnect Network (Nov 2025)',
    'Purpose / Category': 'Transport & Logistics',
    'Mode of Payment': 'RTGS/NEFT',
    'Date': '11 Mar 2025'
  },
  {
    'Vendor Name': 'Earthlygoods Agro Private Limited',
    'Bank A/c No.': '201035788730',
    'Bank Name & Branch (Verified)': 'IndusInd Bank, Calcutta',
    'Description': 'Payment for Procurement & Supply Chain Support – Learning Kits, Stationery & IT Materials',
    'Purpose / Category': 'Procurement & Operations',
    'Mode of Payment': 'RTGS/NEFT',
    'Date': '11 Mar 2025'
  },
  {
    'Vendor Name': 'Banavat Constructions Private Limited',
    'Bank A/c No.': '259211081730',
    'Bank Name & Branch (Verified)': 'IndusInd Bank, Preet Vihar',
    'Description': 'Payment for Civil & Electrical Setup Work at LearnsConnect Training Centers (Work Ref: LC/TN/2025/Phase-II)',
    'Purpose / Category': 'Construction & Site Development',
    'Mode of Payment': 'RTGS/NEFT',
    'Date': '11 Mar 2025'
  },
  {
    'Vendor Name': 'Extreame Horizons Tours and Travels',
    'Bank A/c No.': '258459188735',
    'Bank Name & Branch (Verified)': 'IndusInd Bank, Kasturba Road, Borivali East',
    'Description': 'Payment for Travel, Coordination & Event Management Services – Franchise Onboarding & Staff Training Programs (Nov 2025)',
    'Purpose / Category': 'Travel & Coordination',
    'Mode of Payment': 'RTGS/NEFT',
    'Date': '11 Mar 2025'
  },
  {
    'Vendor Name': 'Unique Vision Multiservices Private Limited',
    'Bank A/c No.': '921020006021559',
    'Bank Name & Branch (Verified)': 'Axis Bank, Gondia',
    'Description': 'Payment for Printing, Branding & Promotional Material Design – LearnsConnect Franchise Marketing Collaterals (Batch ID: Q4-2025)',
    'Purpose / Category': 'Printing & Marketing Collateral',
    'Mode of Payment': 'RTGS/NEFT',
    'Date': '11 Mar 2025'
  },
  {
    'Vendor Name': 'Golden Pettle Trading Private Limited',
    'Bank A/c No.': '201035783063',
    'Bank Name & Branch (Verified)': 'IndusInd Bank, Hingna Road, Nagpur',
    'Description': 'Payment for Procurement of Electronic Accessories, Network Tools, and Support Materials for Franchise Setup & Equipment Maintenance (Work Order for Nov25)',
    'Purpose / Category': 'Procurement & Technical Supplies',
    'Mode of Payment': 'RTGS/NEFT',
    'Date': '11 Mar 2025'
  },
  {
    'Vendor Name': 'Elight Property Solution',
    'Bank A/c No.': '925020024559766',
    'Bank Name & Branch (Verified)': 'Axis Bank, Ashoka Garden, Bhopal',
    'Description': 'Payment for Office Rental Assistance, Facility Management, and Property Maintenance Services (as per Lease/Service Agreement – Nov 2025)',
    'Purpose / Category': 'Property & Facilities',
    'Mode of Payment': 'RTGS/NEFT',
    'Date': '11 Mar 2025'
  }
];

/**
 * Load beneficiary data from hardcoded database
 * @returns {Promise<Array>} Array of objects with beneficiary data
 */
export const loadSBIBeneficiaryData = async () => {
  // Return the hardcoded database directly
  return BENEFICIARY_DATABASE;
};

/**
 * Normalize name for matching (remove extra spaces, convert to lowercase, etc.)
 * @param {string} name - Name to normalize
 * @returns {string} Normalized name
 */
const normalizeName = (name) => {
  if (!name) return '';
  return String(name)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ') // Replace multiple spaces with single space
    .replace(/[^\w\s]/g, '') // Remove special characters (but keep spaces)
    .replace(/\s+/g, ' ') // Replace multiple spaces again after removing special chars
    .trim();
};

/**
 * Extract key words from name for fuzzy matching
 * @param {string} name - Name to extract keywords from
 * @returns {Array} Array of key words (3+ characters)
 */
const extractKeywords = (name) => {
  if (!name) return [];
  const normalized = normalizeName(name);
  // Split by spaces and filter words that are 3+ characters
  const words = normalized.split(/\s+/).filter(word => word.length >= 3);
  return words;
};

/**
 * Calculate similarity score between two names (0-1)
 * @param {string} name1 - First name
 * @param {string} name2 - Second name
 * @returns {number} Similarity score (0-1)
 */
const calculateSimilarity = (name1, name2) => {
  const normalized1 = normalizeName(name1);
  const normalized2 = normalizeName(name2);
  
  // Exact match
  if (normalized1 === normalized2) return 1.0;
  
  // One contains the other (very high score)
  if (normalized1.includes(normalized2) || normalized2.includes(normalized1)) {
    const longer = normalized1.length > normalized2.length ? normalized1 : normalized2;
    const shorter = normalized1.length > normalized2.length ? normalized2 : normalized1;
    const containsScore = shorter.length / longer.length;
    // Boost contains matches - they're very likely correct
    return Math.min(0.95, containsScore + 0.2);
  }
  
  // Keyword matching - improved algorithm
  const keywords1 = extractKeywords(name1);
  const keywords2 = extractKeywords(name2);
  
  if (keywords1.length === 0 || keywords2.length === 0) return 0;
  
  let matchCount = 0;
  let totalKeywords = Math.max(keywords1.length, keywords2.length);
  
  for (const keyword of keywords1) {
    if (keywords2.some(kw => {
      // Exact keyword match
      if (kw === keyword) return true;
      // One contains the other
      if (kw.includes(keyword) || keyword.includes(kw)) return true;
      return false;
    })) {
      matchCount++;
    }
  }
  
  // Calculate score based on keyword matches
  // If most keywords match, it's likely a match
  const keywordScore = matchCount / totalKeywords;
  
  // Also check if important keywords (longer words) match
  const importantKeywords1 = keywords1.filter(kw => kw.length >= 5);
  const importantKeywords2 = keywords2.filter(kw => kw.length >= 5);
  let importantMatches = 0;
  
  if (importantKeywords1.length > 0 && importantKeywords2.length > 0) {
    for (const keyword of importantKeywords1) {
      if (importantKeywords2.some(kw => kw.includes(keyword) || keyword.includes(kw))) {
        importantMatches++;
      }
    }
    const importantScore = importantMatches / Math.max(importantKeywords1.length, importantKeywords2.length);
    // Boost score if important keywords match
    return Math.max(keywordScore, keywordScore * 0.7 + importantScore * 0.3);
  }
  
  return keywordScore;
};

/**
 * Find matching beneficiary data from database based on payee name
 * Uses fuzzy matching to handle unwanted characters and variations
 * @param {string} payeeName - Name to search for (may contain unwanted characters)
 * @returns {Object|null} Matching beneficiary data or null
 */
export const findBeneficiaryData = async (payeeName) => {
  if (!payeeName) return null;

  const beneficiaryData = await loadSBIBeneficiaryData();
  if (!beneficiaryData || beneficiaryData.length === 0) {
    return null;
  }

  const normalizedSearchName = normalizeName(payeeName);
  let bestMatch = null;
  let bestScore = 0;
  const MIN_SIMILARITY_THRESHOLD = 0.25; // Lowered to 25% similarity to be more lenient

  // Try to find best match using fuzzy matching
  for (const row of beneficiaryData) {
    // Primary column: Vendor Name
    const vendorName = row['Vendor Name'];
    if (vendorName) {
      const similarity = calculateSimilarity(payeeName, vendorName);
      
      // Also try direct contains check (case insensitive)
      const vendorNameLower = normalizeName(vendorName);
      if (normalizedSearchName.includes(vendorNameLower) || vendorNameLower.includes(normalizedSearchName)) {
        const containsScore = 0.8; // High score for contains match
        if (containsScore > bestScore) {
          bestScore = containsScore;
          bestMatch = row;
        }
      }
      
      if (similarity > bestScore) {
        bestScore = similarity;
        bestMatch = row;
      }
      
      // If we have a very high match (80%+), return immediately
      if (similarity >= 0.8 || (normalizedSearchName.includes(vendorNameLower) || vendorNameLower.includes(normalizedSearchName))) {
        return row;
      }
    }
    
    // Also check other columns that might contain the name
    const nameColumns = ['Name', 'Payee Name', 'Beneficiary Name', 'Payee', 'Beneficiary', 'Description'];
    for (const colName of nameColumns) {
      const cellValue = row[colName];
      if (cellValue && typeof cellValue === 'string') {
        const similarity = calculateSimilarity(payeeName, cellValue);
        if (similarity > bestScore) {
          bestScore = similarity;
          bestMatch = row;
        }
        if (similarity >= 0.9) {
          return row;
        }
      }
    }
  }

  // Return best match if it meets the threshold
  if (bestScore >= MIN_SIMILARITY_THRESHOLD && bestMatch) {
    return bestMatch;
  }

  return null;
};

/**
 * Get a specific field from beneficiary data
 * @param {Object} beneficiaryData - Beneficiary data object
 * @param {string|Array} fieldNames - Field name(s) to look for (in order of preference)
 * @returns {string} Field value or '-'
 */
export const getBeneficiaryField = (beneficiaryData, fieldNames) => {
  if (!beneficiaryData) return '-';
  
  const fields = Array.isArray(fieldNames) ? fieldNames : [fieldNames];
  
  for (const field of fields) {
    // Try exact match first
    if (beneficiaryData[field] !== undefined && beneficiaryData[field] !== null) {
      const value = String(beneficiaryData[field]).trim();
      if (value.length > 0) {
        return value;
      }
    }
    
    // Try case-insensitive match with normalized keys
    for (const key of Object.keys(beneficiaryData)) {
      const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
      const normalizedField = field.toLowerCase().replace(/[^a-z0-9]/g, '');
      
      if (normalizedKey === normalizedField) {
        const value = String(beneficiaryData[key] || '').trim();
        if (value.length > 0) {
          return value;
        }
      }
      
      // Also check for partial matches (e.g., "Bank Account No" matches "Bank A/c No.")
      if (normalizedKey.includes(normalizedField) || normalizedField.includes(normalizedKey)) {
        const value = String(beneficiaryData[key] || '').trim();
        if (value.length > 0 && normalizedField.length >= 3) {
          return value;
        }
      }
    }
  }
  
  return '-';
};

