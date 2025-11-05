import { useState, useMemo, useEffect, useRef } from 'react';

function DataTable({ data, originalData, currentPage, totalPages, totalItems, itemsPerPage, onPageChange, selectedRowIndices = new Set(), onSelectionChange }) {
  const [searchQuery, setSearchQuery] = useState('');
  const [debouncedSearchQuery, setDebouncedSearchQuery] = useState('');
  const searchInputRef = useRef(null);

  // Debounce search query for better performance
  useEffect(() => {
    const timer = setTimeout(() => {
      setDebouncedSearchQuery(searchQuery);
      if (searchQuery && currentPage !== 1) {
        onPageChange(1);
      }
    }, 300);

    return () => clearTimeout(timer);
  }, [searchQuery, currentPage, onPageChange]);

  // Keyboard shortcut: Ctrl+F or Cmd+F to focus search
  useEffect(() => {
    const handleKeyDown = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
        e.preventDefault();
        searchInputRef.current?.focus();
      }
      // Escape to clear search
      if (e.key === 'Escape' && searchQuery) {
        setSearchQuery('');
        onPageChange(1);
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [searchQuery, onPageChange]);

  if (!data || data.length === 0) {
    return (
      <div className="table-container p-8 text-center text-professional-gray">
        No data to display. Please apply a filter or upload a file.
      </div>
    )
  }

  // Get all unique keys from the data (handle both formats)
  // Ensure Serial No appears first if it exists
  const getColumns = () => {
    if (!data || data.length === 0) return [];
    const firstItem = data[0];
    let columns;
    // Check if data is in { row, originalIndex } format
    if (firstItem && typeof firstItem === 'object' && 'row' in firstItem) {
      columns = Object.keys(firstItem.row || {});
    } else {
      columns = Object.keys(firstItem || {});
    }
    
    // Move Serial No to the beginning if it exists
    const serialNoIndex = columns.findIndex(col => 
      col.toLowerCase().replace(/[^a-z0-9]/g, '') === 'serialno' || 
      col.toLowerCase().replace(/[^a-z0-9]/g, '') === 'serialnumber'
    );
    if (serialNoIndex > 0) {
      const serialNo = columns[serialNoIndex];
      columns.splice(serialNoIndex, 1);
      columns.unshift(serialNo);
    }
    
    return columns;
  };
  const columns = useMemo(() => getColumns(), [data]);

  // Convert selectedRowIndices to Set if it's an array (for backward compatibility)
  const selectedSet = useMemo(() => {
    if (selectedRowIndices instanceof Set) {
      return selectedRowIndices;
    }
    return new Set(selectedRowIndices || []);
  }, [selectedRowIndices]);

  // Check if a row is selected by index (O(1) lookup)
  const isRowSelected = (originalIndex) => {
    return selectedSet.has(originalIndex);
  };

  // Filter data based on search query - include original indices
  // Data may already be in { row, originalIndex } format from parent, or plain row objects
  const filteredData = useMemo(() => {
    // Check if data is already in { row, originalIndex } format
    const isDataWithIndices = data.length > 0 && data[0] && typeof data[0] === 'object' && 'row' in data[0] && 'originalIndex' in data[0];
    
    if (!debouncedSearchQuery.trim()) {
      // Return data with indices when no filter
      if (isDataWithIndices) {
        return data; // Already in correct format
      }
      return data.map((row, index) => ({ row, originalIndex: index }));
    }

    const query = debouncedSearchQuery.toLowerCase().trim();
    const filtered = [];
    
    if (isDataWithIndices) {
      // Data is already in { row, originalIndex } format
      data.forEach(({ row, originalIndex }) => {
        const matches = columns.some(column => {
          const cellValue = String(row[column] || '').toLowerCase();
          return cellValue.includes(query);
        });
        if (matches) {
          filtered.push({ row, originalIndex });
        }
      });
    } else {
      // Data is plain row objects
      data.forEach((row, index) => {
        const matches = columns.some(column => {
          const cellValue = String(row[column] || '').toLowerCase();
          return cellValue.includes(query);
        });
        if (matches) {
          filtered.push({ row, originalIndex: index });
        }
      });
    }
    return filtered;
  }, [data, debouncedSearchQuery, columns]);

  // Highlight search matches in cell content
  const highlightText = (text, query) => {
    if (!query || !text) return text;
    
    const textStr = String(text);
    const queryLower = query.toLowerCase();
    const textLower = textStr.toLowerCase();
    const index = textLower.indexOf(queryLower);
    
    if (index === -1) return textStr;
    
    const before = textStr.substring(0, index);
    const match = textStr.substring(index, index + query.length);
    const after = textStr.substring(index + query.length);
    
    return (
      <>
        {before}
        <mark className="bg-yellow-200 text-gray-900 px-1 rounded font-medium">
          {match}
        </mark>
        {after}
      </>
    );
  };

  // Calculate pagination for filtered data
  const filteredTotalItems = filteredData.length;
  const filteredTotalPages = Math.max(1, Math.ceil(filteredTotalItems / itemsPerPage));
  const filteredCurrentPage = Math.min(currentPage, filteredTotalPages);
  const startIndex = (filteredCurrentPage - 1) * itemsPerPage;
  const endIndex = startIndex + itemsPerPage;
  const paginatedData = filteredData.slice(startIndex, endIndex);

  // Check if all filtered rows are selected (O(n) but optimized)
  const isAllFilteredSelected = useMemo(() => {
    if (!filteredData || filteredData.length === 0) return false;
    return filteredData.every(({ originalIndex }) => selectedSet.has(originalIndex));
  }, [filteredData, selectedSet]);

  // Handle individual row checkbox
  const handleRowSelect = (originalIndex, checked) => {
    const newSelectedSet = new Set(selectedSet);
    
    if (checked) {
      newSelectedSet.add(originalIndex);
    } else {
      newSelectedSet.delete(originalIndex);
    }
    
    if (onSelectionChange) {
      // Convert Set to Array for backward compatibility, and get actual row data from originalData
      const dataSource = originalData || data;
      const selectedData = Array.from(newSelectedSet).map(idx => {
        // If originalData is provided, use it; otherwise try to get from current data
        if (originalData && originalData[idx]) {
          return originalData[idx];
        }
        // Fallback: if data is in { row, originalIndex } format, find the row
        if (data && data.length > 0 && data[0] && typeof data[0] === 'object' && 'originalIndex' in data[0]) {
          const found = data.find(item => item.originalIndex === idx);
          return found ? found.row : null;
        }
        return data[idx] || null;
      }).filter(Boolean);
      onSelectionChange(selectedData, newSelectedSet);
    }
  };

  // Handle select all checkbox
  const handleSelectAll = (checked) => {
    const newSelectedSet = new Set(selectedSet);
    
    if (checked) {
      // Add all filtered row indices
      filteredData.forEach(({ originalIndex }) => {
        newSelectedSet.add(originalIndex);
      });
    } else {
      // Remove all filtered row indices
      filteredData.forEach(({ originalIndex }) => {
        newSelectedSet.delete(originalIndex);
      });
    }
    
    if (onSelectionChange) {
      // Convert Set to Array for backward compatibility, and get actual row data from originalData
      const dataSource = originalData || data;
      const selectedData = Array.from(newSelectedSet).map(idx => {
        // If originalData is provided, use it; otherwise try to get from current data
        if (originalData && originalData[idx]) {
          return originalData[idx];
        }
        // Fallback: if data is in { row, originalIndex } format, find the row
        if (data && data.length > 0 && data[0] && typeof data[0] === 'object' && 'originalIndex' in data[0]) {
          const found = data.find(item => item.originalIndex === idx);
          return found ? found.row : null;
        }
        return data[idx] || null;
      }).filter(Boolean);
      onSelectionChange(selectedData, newSelectedSet);
    }
  };

  // Handle search input change
  const handleSearchChange = (e) => {
    setSearchQuery(e.target.value);
  };

  // Clear search
  const handleClearSearch = () => {
    setSearchQuery('');
    setDebouncedSearchQuery('');
    onPageChange(1);
    searchInputRef.current?.focus();
  };

  const handlePrevPage = () => {
    if (filteredCurrentPage > 1) {
      onPageChange(filteredCurrentPage - 1);
    }
  }

  const handleNextPage = () => {
    if (filteredCurrentPage < filteredTotalPages) {
      onPageChange(filteredCurrentPage + 1);
    }
  }

  // Get search hints based on columns
  const searchHint = useMemo(() => {
    if (columns.length === 0) return '';
    const sampleColumns = columns.slice(0, 3).join(', ');
    return `Search in ${sampleColumns}${columns.length > 3 ? '...' : ''}`;
  }, [columns]);

  // Calculate minimum table height to show 100 rows
  // Each row: padding (1.25rem * 2 = 40px) + line-height (~25px) + border (1px) ≈ 66-70px per row
  // Header: ~70px
  // For 100 rows: 100 * 70px = 7000px + header 70px = 7070px minimum
  const minTableHeight = useMemo(() => {
    const rowHeight = 70; // Approximate height per row in pixels
    const headerHeight = 70;
    const minRows = 100;
    return (minRows * rowHeight) + headerHeight; // 7070px minimum
  }, []);

  return (
    <div className="table-container">
      {/* Enhanced Search Bar */}
      <div className="p-4 bg-gradient-to-r from-gray-50 to-gray-100 border-b border-gray-200">
        <div className="relative">
          <input
            ref={searchInputRef}
            type="text"
            placeholder={searchHint || "Search across all columns... (Press Ctrl+F to focus)"}
            value={searchQuery}
            onChange={handleSearchChange}
            className="w-full px-4 py-3 pl-11 pr-12 border-2 border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-purple-500 focus:border-purple-500 transition-all shadow-sm hover:border-gray-400 text-base"
          />
          <svg
            className="absolute left-3.5 top-1/2 transform -translate-y-1/2 w-5 h-5 text-gray-400"
            fill="none"
            stroke="currentColor"
            viewBox="0 0 24 24"
          >
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
          </svg>
          {searchQuery && (
            <button
              onClick={handleClearSearch}
              className="absolute right-3 top-1/2 transform -translate-y-1/2 text-gray-400 hover:text-red-500 transition-colors p-1 rounded-full hover:bg-gray-200"
              title="Clear search (Esc)"
            >
              <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
              </svg>
            </button>
          )}
        </div>
        
        {/* Search Info */}
        <div className="mt-3 flex items-center justify-between flex-wrap gap-2">
          {debouncedSearchQuery ? (
            <div className="flex items-center gap-3">
              <div className="flex items-center gap-2 px-3 py-1.5 bg-purple-100 text-purple-700 rounded-full text-sm font-medium">
                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
                Found <span className="font-bold">{filteredTotalItems}</span> result{filteredTotalItems !== 1 ? 's' : ''}
              </div>
              <span className="text-sm text-gray-600">
                for "<span className="font-semibold text-gray-800">{debouncedSearchQuery}</span>"
              </span>
            </div>
          ) : (
            <div className="flex items-center gap-2 text-sm text-gray-500">
              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
              <span>Tip: Use <kbd className="px-2 py-1 bg-gray-200 rounded text-xs font-mono">Ctrl+F</kbd> or <kbd className="px-2 py-1 bg-gray-200 rounded text-xs font-mono">Cmd+F</kbd> to focus search, <kbd className="px-2 py-1 bg-gray-200 rounded text-xs font-mono">Esc</kbd> to clear</span>
            </div>
          )}
          {debouncedSearchQuery && (
            <button
              onClick={handleClearSearch}
              className="text-sm text-purple-600 hover:text-purple-800 font-medium underline flex items-center gap-1"
            >
              Clear all filters
            </button>
          )}
        </div>
      </div>

      {/* Table */}
      <div className="overflow-x-auto overflow-y-auto" style={{ minHeight: `${minTableHeight}px`, maxHeight: 'calc(100vh - 250px)' }}>
        <table className="w-full">
          <thead className="table-header sticky-header">
            <tr>
              <th className="table-cell font-semibold text-left text-professional-dark uppercase text-xs tracking-wider" style={{ width: '50px', textAlign: 'center' }}>
                Select
              </th>
              {columns.map((col, idx) => (
                <th
                  key={idx}
                  className="table-cell font-semibold text-left text-professional-dark uppercase text-xs tracking-wider"
                >
                  {col}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {paginatedData.length > 0 ? (
              paginatedData.map(({ row, originalIndex }, rowIdx) => (
                <tr key={originalIndex} className="hover:bg-purple-50 transition-colors border-b border-gray-100">
                  <td className="table-cell" style={{ width: '50px', textAlign: 'center' }}>
                    <input
                      type="checkbox"
                      checked={isRowSelected(originalIndex)}
                      onChange={(e) => handleRowSelect(originalIndex, e.target.checked)}
                      className="h-5 w-5 text-purple-600 focus:ring-purple-500 border-gray-300 rounded cursor-pointer"
                      style={{ cursor: 'pointer' }}
                    />
                  </td>
                  {columns.map((col, colIdx) => (
                    <td key={colIdx} className="table-cell">
                      {debouncedSearchQuery 
                        ? highlightText(row[col] || '-', debouncedSearchQuery)
                        : (row[col] || '-')
                      }
                    </td>
                  ))}
                </tr>
              ))
            ) : (
              <tr>
                <td colSpan={columns.length + 1} className="table-cell text-center py-12">
                  <div className="flex flex-col items-center justify-center gap-4">
                    <svg className="w-16 h-16 text-gray-300" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9.172 16.172a4 4 0 015.656 0M9 10h.01M15 10h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                    </svg>
                    <div className="text-center">
                      <p className="text-lg font-semibold text-gray-700 mb-2">No results found</p>
                      <p className="text-sm text-gray-500 mb-4">
                        No data matches your search for "<span className="font-semibold text-gray-700">"{debouncedSearchQuery}"</span>"
                      </p>
                      <button
                        onClick={handleClearSearch}
                        className="px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition-colors text-sm font-medium"
                      >
                        Clear search and show all results
                      </button>
                    </div>
                  </div>
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      
      {/* Enhanced Pagination */}
      <div className="bg-gray-50 px-4 py-3 border-t border-gray-200 flex items-center justify-between flex-wrap gap-2">
        <div className="text-sm text-professional-gray flex items-center gap-2 flex-wrap">
          <span>
            Showing <span className="font-semibold text-gray-800">{filteredTotalItems > 0 ? startIndex + 1 : 0}</span> to{' '}
            <span className="font-semibold text-gray-800">{Math.min(endIndex, filteredTotalItems)}</span> of{' '}
            <span className="font-semibold text-gray-800">{filteredTotalItems}</span> entries
          </span>
          {debouncedSearchQuery && (
            <span className="px-2 py-1 bg-blue-100 text-blue-700 rounded text-xs font-medium">
              (Filtered from {totalItems} total)
            </span>
          )}
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={handlePrevPage}
            disabled={filteredCurrentPage === 1}
            className="px-4 py-2 text-sm font-medium text-professional-dark bg-white border-2 border-gray-300 rounded-md hover:bg-gray-50 hover:border-purple-500 disabled:opacity-50 disabled:cursor-not-allowed disabled:hover:border-gray-300 transition-all"
          >
            Previous
          </button>
          <span className="px-4 py-2 text-sm font-medium text-professional-dark bg-white border border-gray-300 rounded-md">
            Page <span className="font-bold">{filteredCurrentPage}</span> of <span className="font-bold">{filteredTotalPages}</span>
          </span>
          <button
            onClick={handleNextPage}
            disabled={filteredCurrentPage === filteredTotalPages}
            className="px-4 py-2 text-sm font-medium text-professional-dark bg-white border-2 border-gray-300 rounded-md hover:bg-gray-50 hover:border-purple-500 disabled:opacity-50 disabled:cursor-not-allowed disabled:hover:border-gray-300 transition-all"
          >
            Next
          </button>
        </div>
      </div>
    </div>
  )
}

export default DataTable
