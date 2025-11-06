import { useState } from 'react';
import { motion } from 'framer-motion';
import Navbar from './components/Navbar';
import Footer from './components/Footer';
import UploadSection from './components/UploadSection';
import TransactionTypeSelector from './components/TransactionTypeSelector';
import SummaryCards from './components/SummaryCards';
import DataTable from './components/DataTable';
import DownloadButtons from './components/DownloadButtons';
import SelectedDetailsCard from './components/SelectedDetailsCard';
import { parseFile } from './utils/fileParser';
import { exportToCSV, exportToXLSX, exportToDOCX, exportToZIP, exportToDOCXMixed } from './utils/exportUtils';
import { exportToDOCXDebit } from './utils/exportUtilsDebit';
import { exportToPDF, generatePDFBlob } from './utils/exportUtilsPDF';

function App() {
  const [billingType, setBillingType] = useState('student');
  const [data, setData] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [summary, setSummary] = useState({ 
    totalCount: 0, 
    creditTotal: 0, 
    debitTotal: 0, 
    netBalance: 0 
  });
  const [startDate, setStartDate] = useState(null);
  const [endDate, setEndDate] = useState(null);
  const [selectedRows, setSelectedRows] = useState([]);
  const [selectedRowIndices, setSelectedRowIndices] = useState(new Set());
  const [transactionType, setTransactionType] = useState('all'); // 'all', 'credit', or 'debit'
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage] = useState(100); // Number of items per page

  const handleFileUpload = async (file) => {
    setIsProcessing(true);
    try {
      const result = await parseFile(file, billingType);
      
      // Add serial number to each row
      const dataWithSerial = result.data.map((row, index) => ({
        ...row,
        'Serial No': index + 1, // Add serial number starting from 1
      }));
      
      setData(dataWithSerial);
      setTransactionType('all'); // Reset to show all transactions when new file is uploaded
      setCurrentPage(1); // Reset to first page
      setSelectedRows([]); // Reset selected rows
      setSelectedRowIndices(new Set()); // Reset selected row indices
      
      // Store extracted dates from metadata
      setStartDate(result.startDate || null);
      setEndDate(result.endDate || null);
      
      // Calculate summary with credit and debit totals
      const { creditTotal, debitTotal, netBalance } = calculateCreditDebitTotals(dataWithSerial);
      
      setSummary({
        totalCount: result.uniqueCount,
        creditTotal: creditTotal,
        debitTotal: debitTotal,
        netBalance: netBalance,
      });
    } catch (error) {
      console.error('Error parsing file:', error);
      alert(`Error: ${error.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleTransactionTypeSelect = (type) => {
    setTransactionType(type);
    setCurrentPage(1); // Reset to first page when filter changes
    // Don't clear selections when switching filters - preserve user selections
  };

  const getFilteredData = () => {
    if (!data) return null;
    
    // If 'all', return all data with original indices
    if (transactionType === 'all') {
      return data.map((row, index) => ({ row, originalIndex: index }));
    }
    
    // Filter data based on credit/debit column values with original indices
    // Check if Credit column has value (for credit) or Debit column has value (for debit)
    const filtered = [];
    data.forEach((row, index) => {
      // Find credit and debit columns
      const creditKeys = ['credit', 'cr', 'deposit', 'received', 'income', 'in'];
      const debitKeys = ['debit', 'dr', 'withdrawal', 'paid', 'expense', 'out'];
      
      let hasCredit = false;
      let hasDebit = false;
      
      // Check for credit value
      for (const key of Object.keys(row)) {
        const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
        if (creditKeys.some(ck => normalizedKey.includes(ck))) {
          const value = row[key];
          if (value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
            const amt = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
            if (amt > 0) {
              hasCredit = true;
              break;
            }
          }
        }
      }
      
      // Check for debit value
      for (const key of Object.keys(row)) {
        const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
        if (debitKeys.some(dk => normalizedKey.includes(dk))) {
          const value = row[key];
          if (value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
            const amt = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
            if (amt > 0) {
              hasDebit = true;
              break;
            }
          }
        }
      }
      
      // Check if row should be included based on transaction type
      let shouldInclude = false;
      if (transactionType === 'credit') {
        shouldInclude = hasCredit && !hasDebit; // Show only rows with credit values
      } else if (transactionType === 'debit') {
        shouldInclude = hasDebit && !hasCredit; // Show only rows with debit values
      }
      
      if (shouldInclude) {
        filtered.push({ row, originalIndex: index });
      }
    });
    
    return filtered;
  };

  const handleBillingTypeChange = (type) => {
    setBillingType(type);
    // Clear data when billing type changes
    if (data) {
      setData(null);
      setSummary({ totalCount: 0, creditTotal: 0, debitTotal: 0, netBalance: 0 });
      setSelectedRows([]);
      setSelectedRowIndices(new Set());
      setStartDate(null);
      setEndDate(null);
      setTransactionType('all'); // Reset to show all transactions
    }
  };

  const handleSelectionChange = (selectedData, selectedIndices) => {
    setSelectedRows(selectedData);
    setSelectedRowIndices(selectedIndices || new Set());
  };

  const handleClearAll = () => {
    setSelectedRows([]);
    setSelectedRowIndices(new Set());
  };

  // Preview selected rows as PDF in a new tab
  const handleViewSelected = async () => {
    try {
      if (!selectedRows || selectedRows.length === 0) {
        alert('Please select at least one record to view');
        return;
      }

      const arrayBuffer = await generatePDFBlob(selectedRows, billingType, transactionType || 'all');
      if (!arrayBuffer) {
        alert('Unable to generate PDF preview');
        return;
      }

      const blob = new Blob([arrayBuffer], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      window.open(url, '_blank', 'noopener,noreferrer');
      // Revoke after a minute to free memory
      setTimeout(() => URL.revokeObjectURL(url), 60 * 1000);
    } catch (err) {
      console.error('Error generating PDF preview:', err);
      alert('Error generating PDF preview');
    }
  };

  // Helper function to detect if data contains credit or debit transactions
  const detectTransactionType = (dataToCheck) => {
    if (!dataToCheck || dataToCheck.length === 0) return 'credit'; // Default to credit
    
    let hasCredit = false;
    let hasDebit = false;
    
    for (const row of dataToCheck) {
      for (const key of Object.keys(row)) {
        const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
        const value = row[key];
        
        if ((normalizedKey.includes('credit') || normalizedKey.includes('cr')) && 
            value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
          const amt = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
          if (amt > 0) hasCredit = true;
        }
        
        if ((normalizedKey.includes('debit') || normalizedKey.includes('dr')) && 
            value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
          const amt = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
          if (amt > 0) hasDebit = true;
        }
      }
      
      // If we found both, prioritize based on which is more common in the dataset
      if (hasCredit && hasDebit) {
        // Count credit vs debit transactions
        let creditCount = 0;
        let debitCount = 0;
        
        for (const r of dataToCheck) {
          let rowHasCredit = false;
          let rowHasDebit = false;
          
          for (const k of Object.keys(r)) {
            const nKey = k.toLowerCase().replace(/[^a-z0-9]/g, '');
            const v = r[k];
            
            if ((nKey.includes('credit') || nKey.includes('cr')) && 
                v && String(v).trim() && !isNaN(parseFloat(String(v).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
              const amt = parseFloat(String(v).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
              if (amt > 0) rowHasCredit = true;
            }
            
            if ((nKey.includes('debit') || nKey.includes('dr')) && 
                v && String(v).trim() && !isNaN(parseFloat(String(v).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
              const amt = parseFloat(String(v).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
              if (amt > 0) rowHasDebit = true;
            }
          }
          
          if (rowHasCredit && !rowHasDebit) creditCount++;
          if (rowHasDebit && !rowHasCredit) debitCount++;
        }
        
        return creditCount >= debitCount ? 'credit' : 'debit';
      }
      
      if (hasCredit && !hasDebit) return 'credit';
      if (hasDebit && !hasCredit) return 'debit';
    }
    
    return 'credit'; // Default to credit
  };

  // Helper function to check if a row is credit transaction
  const isRowCredit = (row) => {
    let hasCredit = false;
    let hasDebit = false;
    
    for (const key of Object.keys(row)) {
      const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
      const value = row[key];
      
      if ((normalizedKey.includes('credit') || normalizedKey.includes('cr')) && 
          value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
        const amt = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
        if (amt > 0) hasCredit = true;
      }
      
      if ((normalizedKey.includes('debit') || normalizedKey.includes('dr')) && 
          value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
        const amt = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
        if (amt > 0) hasDebit = true;
      }
    }
    
    return hasCredit && !hasDebit; // Credit row if has credit and no debit
  };

    const handleDownload = async (type, downloadData, billingType, useSelection = false) => {
      const timestamp = new Date().toISOString().slice(0, 10);
      
      // Automatically use selected rows if any are selected (useSelection will be true)
      // Otherwise use ALL uploaded data (not filtered)
      let dataToExport;
      if (useSelection && selectedRows.length > 0) {
        // Use selected rows
        dataToExport = selectedRows;
      } else {
        // Use ALL uploaded data (not filtered) when no selection
        dataToExport = data || downloadData;
      }
    
    if (!dataToExport || dataToExport.length === 0) {
      alert('No data to download');
      return;
    }
    
    // Check if data contains both credit and debit rows (for mixed export)
    let hasCredit = false;
    let hasDebit = false;
    for (const row of dataToExport) {
      if (isRowCredit(row)) {
        hasCredit = true;
      } else {
        // Check if it's a debit row
        let rowHasDebit = false;
        for (const key of Object.keys(row)) {
          const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
          const value = row[key];
          if ((normalizedKey.includes('debit') || normalizedKey.includes('dr')) && 
              value && String(value).trim() && !isNaN(parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, '')))) {
            const amt = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(/,/g, ''));
            if (amt > 0) {
              rowHasDebit = true;
              break;
            }
          }
        }
        if (rowHasDebit) hasDebit = true;
      }
      if (hasCredit && hasDebit) break; // Found both, no need to continue
    }
    
    const isMixed = hasCredit && hasDebit;
    
    switch (type) {
      case 'zip':
        await exportToZIP(dataToExport, billingType, transactionType || 'all', `ellen-invoice-all-${timestamp}.zip`);
        break;
      case 'csv':
        exportToCSV(dataToExport, `ellen-invoice-${timestamp}.csv`);
        break;
      case 'xlsx':
        exportToXLSX(dataToExport, `ellen-invoice-${timestamp}.xlsx`);
        break;
      case 'docx':
        // If mixed data, use combined export; otherwise use specific export
        if (isMixed) {
          await exportToDOCXMixed(dataToExport, billingType, `ellen-invoice-mixed-${timestamp}.docx`);
        } else if (hasDebit || transactionType === 'debit') {
          await exportToDOCXDebit(dataToExport, `payment-voucher-${timestamp}.docx`);
        } else {
          await exportToDOCX(dataToExport, billingType, transactionType || 'all', `ellen-invoice-${timestamp}.docx`);
        }
        break;
      case 'pdf':
        await exportToPDF(dataToExport, billingType, transactionType || 'all', `ellen-invoice-${timestamp}.pdf`);
        break;
      default:
        console.error('Unknown download type');
    }
  };

  // Helper function to calculate credit and debit totals separately
  const calculateCreditDebitTotals = (data) => {
    if (!data || data.length === 0) {
      return { creditTotal: 0, debitTotal: 0, netBalance: 0 };
    }

    let creditTotal = 0;
    let debitTotal = 0;

    // Keywords to identify credit and debit columns
    const creditKeys = ['credit', 'cr', 'deposit', 'received', 'income', 'in'];
    const debitKeys = ['debit', 'dr', 'withdrawal', 'paid', 'expense', 'out'];

    // Find credit and debit columns by checking column names (not values)
    const firstRow = data[0];
    let creditColumn = null;
    let debitColumn = null;

    // Find credit column by name only - check all columns
    // Note: fileParser normalizes column names to lowercase and removes special chars
    for (const key of Object.keys(firstRow)) {
      const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
      
      // Exact match or contains credit keyword (priority to exact match)
      if (normalizedKey === 'credit' || normalizedKey === 'cr') {
        creditColumn = key;
        break;
      }
      // Check if column name contains credit keywords (exclude 'debit' keywords)
      if (creditKeys.some(ck => normalizedKey.includes(ck)) && 
          !debitKeys.some(dk => normalizedKey.includes(dk))) {
        if (!creditColumn) { // Take first match
          creditColumn = key;
        }
      }
    }

    // Find debit column by name only - check all columns
    for (const key of Object.keys(firstRow)) {
      const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
      
      // Exact match or contains debit keyword (priority to exact match)
      if (normalizedKey === 'debit' || normalizedKey === 'dr') {
        debitColumn = key;
        break;
      }
      // Check if column name contains debit keywords (exclude 'credit' keywords)
      if (debitKeys.some(dk => normalizedKey.includes(dk)) && 
          !creditKeys.some(ck => normalizedKey.includes(ck))) {
        if (!debitColumn) { // Take first match
          debitColumn = key;
        }
      }
    }

    // Debug: Log found columns
    console.log('=== Credit/Debit Column Detection ===');
    console.log('Available columns:', Object.keys(firstRow));
    console.log('Credit column found:', creditColumn);
    console.log('Debit column found:', debitColumn);
    
    // Show sample values from first few rows
    if (creditColumn) {
      const sampleCredits = data.slice(0, 5).map(r => r[creditColumn]).filter(v => v);
      console.log('Sample credit values:', sampleCredits);
    }
    if (debitColumn) {
      const sampleDebits = data.slice(0, 5).map(r => r[debitColumn]).filter(v => v);
      console.log('Sample debit values:', sampleDebits);
    }

    // Calculate totals from all rows
    let creditCount = 0;
    let debitCount = 0;
    
    data.forEach((row, index) => {
      // Calculate credit - check all columns that might contain credit values
      if (creditColumn) {
        // Use identified credit column
        const value = row[creditColumn];
        if (value !== null && value !== undefined) {
          const valueStr = String(value).trim();
          if (valueStr !== '' && valueStr !== '-') {
            // Clean the string: remove currency symbols, keep numbers, dots, commas, minus
            const cleanedStr = valueStr.replace(/[^\d.,-]/g, '').replace(/,/g, '');
            const creditValue = parseFloat(cleanedStr);
            if (!isNaN(creditValue) && creditValue > 0) {
              creditTotal += Math.abs(creditValue);
              creditCount++;
              // Log first few credit transactions for debugging
              if (creditCount <= 3) {
                console.log(`Credit row ${index}: value="${valueStr}" -> ${creditValue}`);
              }
            }
          }
        }
      } else {
        // Fallback: check all columns for credit keywords
        for (const key of Object.keys(row)) {
          const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
          if (creditKeys.some(ck => normalizedKey.includes(ck)) && 
              !debitKeys.some(dk => normalizedKey.includes(dk))) {
            const value = row[key];
            if (value !== null && value !== undefined) {
              const valueStr = String(value).trim();
              if (valueStr !== '' && valueStr !== '-') {
                const cleanedStr = valueStr.replace(/[^\d.,-]/g, '').replace(/,/g, '');
                const creditValue = parseFloat(cleanedStr);
                if (!isNaN(creditValue) && creditValue > 0) {
                  creditTotal += Math.abs(creditValue);
                  creditCount++;
                  if (creditCount <= 3) {
                    console.log(`Credit row ${index} (fallback, column: ${key}): value="${valueStr}" -> ${creditValue}`);
                  }
                  break; // Only count once per row
                }
              }
            }
          }
        }
      }

      // Calculate debit - check all columns that might contain debit values
      if (debitColumn) {
        // Use identified debit column
        const value = row[debitColumn];
        if (value !== null && value !== undefined) {
          const valueStr = String(value).trim();
          if (valueStr !== '' && valueStr !== '-') {
            // Clean the string: remove currency symbols, keep numbers, dots, commas, minus
            const cleanedStr = valueStr.replace(/[^\d.,-]/g, '').replace(/,/g, '');
            const debitValue = parseFloat(cleanedStr);
            if (!isNaN(debitValue) && debitValue > 0) {
              debitTotal += Math.abs(debitValue);
              debitCount++;
              // Log first few debit transactions for debugging
              if (debitCount <= 3) {
                console.log(`Debit row ${index}: value="${valueStr}" -> ${debitValue}`);
              }
            }
          }
        }
      } else {
        // Fallback: check all columns for debit keywords
        for (const key of Object.keys(row)) {
          const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
          if (debitKeys.some(dk => normalizedKey.includes(dk)) && 
              !creditKeys.some(ck => normalizedKey.includes(ck))) {
            const value = row[key];
            if (value !== null && value !== undefined) {
              const valueStr = String(value).trim();
              if (valueStr !== '' && valueStr !== '-') {
                const cleanedStr = valueStr.replace(/[^\d.,-]/g, '').replace(/,/g, '');
                const debitValue = parseFloat(cleanedStr);
                if (!isNaN(debitValue) && debitValue > 0) {
                  debitTotal += Math.abs(debitValue);
                  debitCount++;
                  if (debitCount <= 3) {
                    console.log(`Debit row ${index} (fallback, column: ${key}): value="${valueStr}" -> ${debitValue}`);
                  }
                  break; // Only count once per row
                }
              }
            }
          }
        }
      }
    });

    const netBalance = creditTotal - debitTotal;

    // Debug: Log calculated totals
    console.log('=== Calculation Summary ===');
    console.log('Credit Total:', creditTotal);
    console.log('Debit Total:', debitTotal);
    console.log('Net Balance:', netBalance);
    console.log('Total rows processed:', data.length);
    console.log('Credit transactions found:', creditCount);
    console.log('Debit transactions found:', debitCount);

    return {
      creditTotal: creditTotal,
      debitTotal: debitTotal,
      netBalance: netBalance,
    };
  };

  return (
    <div className="min-h-screen flex flex-col" style={{ position: 'relative' }}>
      {/* Glassmorphism Background Layer */}
      <div
        style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(255, 255, 255, 0.15)',
          backdropFilter: 'blur(30px)',
          WebkitBackdropFilter: 'blur(30px)',
          zIndex: 0,
          pointerEvents: 'none',
        }}
      />
      
      <div style={{ position: 'relative', zIndex: 1, minHeight: '100vh', display: 'flex', flexDirection: 'column' }}>
        <Navbar />
        
        <main className="flex-1 container mx-auto px-4 sm:px-6 lg:px-8 py-6 sm:py-8 max-w-7xl">
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            transition={{ duration: 0.5 }}
          >
          {/* Upload Section */}
          <UploadSection
            onFileUpload={handleFileUpload}
            onBillingTypeChange={handleBillingTypeChange}
            billingType={billingType}
            isProcessing={isProcessing}
          />

          {/* Transaction Type Selector - Always visible after file upload */}
          {data && data.length > 0 && (
            <TransactionTypeSelector
              onSelect={handleTransactionTypeSelect}
              selectedType={transactionType}
            />
          )}

          {/* Summary Cards */}
          {data && data.length > 0 && (
            <>
              <SummaryCards
                totalCount={summary.totalCount}
                creditTotal={summary.creditTotal}
                debitTotal={summary.debitTotal}
                netBalance={summary.netBalance}
              />

              {/* Download Buttons */}
              <DownloadButtons
                data={getFilteredData() || data}
                billingType={billingType}
                selectedRows={selectedRows}
                onDownload={handleDownload}
              />

              {/* Selected Details Card */}
              <SelectedDetailsCard
                selectedRows={selectedRows}
                onClearAll={handleClearAll}
                onView={handleViewSelected}
              />

              {/* Data Table with Search and Pagination */}
              {(() => {
                const filteredData = getFilteredData() || data || [];
                const totalItems = filteredData.length;

                return (
                  <DataTable
                    data={filteredData}
                    originalData={data}
                    currentPage={currentPage}
                    totalPages={Math.max(1, Math.ceil(totalItems / itemsPerPage))}
                    totalItems={totalItems}
                    itemsPerPage={itemsPerPage}
                    onPageChange={setCurrentPage}
                    selectedRowIndices={selectedRowIndices}
                    onSelectionChange={handleSelectionChange}
                  />
                );
              })()}
            </>
          )}

          {/* Empty State */}
          {!data && !isProcessing && (
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              style={{
                textAlign: 'center',
                padding: '3rem 1rem',
                background: 'rgba(255, 255, 255, 0.9)',
                backdropFilter: 'blur(25px)',
                WebkitBackdropFilter: 'blur(25px)',
                border: '1px solid rgba(255, 255, 255, 0.6)',
                boxShadow: '0 8px 32px 0 rgba(31, 38, 135, 0.2)',
                borderRadius: '20px',
              }}
            >
              <p style={{ 
                color: '#666666', 
                fontSize: '1.125rem',
                fontWeight: 500,
                margin: 0,
              }}>
                Upload a file to get started
              </p>
            </motion.div>
          )}
          </motion.div>
        </main>

        <Footer />
      </div>
    </div>
  );
}

export default App;

