import { useState } from 'react';
import { Button, Box, Typography, Paper } from '@mui/material';
import { motion } from 'framer-motion';
import FileDownloadIcon from '@mui/icons-material/FileDownload';
import PictureAsPdfIcon from '@mui/icons-material/PictureAsPdf';

const DownloadButtons = ({ data, billingType, selectedRows = [], onDownload }) => {
  const [isDownloading, setIsDownloading] = useState(false);

  const handleDownload = async (type) => {
    // Automatically use selected rows if any are selected, otherwise use all data
    const hasSelection = selectedRows.length > 0;
    const dataToUse = hasSelection ? selectedRows : data;
    
    if (!dataToUse || dataToUse.length === 0) {
      alert(hasSelection 
        ? 'No records selected. Please select records using checkboxes.' 
        : 'No data to download'
      );
      return;
    }

    setIsDownloading(true);
    try {
      // Pass hasSelection to indicate we should use selected rows
      await onDownload(type, data, billingType, hasSelection);
    } catch (error) {
      console.error('Download error:', error);
      alert('Error downloading file. Please try again.');
    } finally {
      setIsDownloading(false);
    }
  };

  const selectedCount = selectedRows.length;
  const hasSelection = selectedCount > 0;

  const buttons = [
    {
      label: 'Download All as ZIP',
      type: 'zip',
      icon: FileDownloadIcon,
      color: 'from-blue-500 to-blue-600',
    },
    {
      label: 'Download XLSX',
      type: 'xlsx',
      icon: FileDownloadIcon,
      color: 'from-green-500 to-green-600',
    },
    {
      label: 'Download as PDF',
      type: 'pdf',
      icon: PictureAsPdfIcon,
      color: 'from-red-500 to-red-600',
    },
  ];

  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.5 }}
      className="w-full mb-6"
    >
      {/* Selection Info */}
      {hasSelection && (
        <Paper
          elevation={0}
          sx={{
            background: 'rgba(76, 175, 80, 0.15)',
            backdropFilter: 'blur(20px)',
            WebkitBackdropFilter: 'blur(20px)',
            border: '1px solid rgba(76, 175, 80, 0.3)',
            boxShadow: '0 4px 16px 0 rgba(76, 175, 80, 0.2)',
            p: 2,
            mb: 2,
            borderRadius: '12px',
          }}
        >
          <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 1, flexWrap: 'wrap' }}>
            <Typography sx={{ color: '#ffffff', fontWeight: 600, fontSize: '0.9375rem' }}>
              ✓ {selectedCount} record{selectedCount !== 1 ? 's' : ''} selected
            </Typography>
            <Typography sx={{ color: '#ffffff', fontSize: '0.875rem' }}>
              (Only selected records will be downloaded)
            </Typography>
          </Box>
        </Paper>
      )}

      {/* Download Buttons */}
      <Box className="flex flex-wrap gap-3 sm:gap-4 justify-center">
        {buttons.map((button, index) => (
          <motion.div
            key={button.type}
            whileHover={{ scale: 1.05 }}
            whileTap={{ scale: 0.95 }}
          >
            <Button
              variant="contained"
              onClick={() => handleDownload(button.type)}
              disabled={isDownloading || !data || data.length === 0}
              className={`bg-gradient-to-r ${button.color} text-white px-4 sm:px-6 py-2 sm:py-3 rounded-lg shadow-lg hover:shadow-xl transition-all`}
              sx={{
                textTransform: 'none',
                fontSize: '0.875rem',
                fontWeight: 600,
                minWidth: '140px',
              }}
              startIcon={<button.icon />}
            >
              {isDownloading 
                ? 'Processing...' 
                : hasSelection
                ? `${button.label.replace('Download ', '')} (${selectedCount})`
                : button.label
              }
            </Button>
          </motion.div>
        ))}
      </Box>

      {/* Help Text */}
      {!hasSelection && (
        <Typography 
          variant="body2" 
          sx={{ 
            color: '#666666', 
            textAlign: 'center', 
            mt: 2,
            fontSize: '0.875rem',
            background: 'rgba(255, 255, 255, 0.9)',
            backdropFilter: 'blur(20px)',
            p: 1.5,
            borderRadius: '12px',
            border: '1px solid rgba(102, 126, 234, 0.2)',
          }}
        >
          💡 Select records using checkboxes to download only selected entries, or download all records
        </Typography>
      )}
    </motion.div>
  );
};

export default DownloadButtons;

