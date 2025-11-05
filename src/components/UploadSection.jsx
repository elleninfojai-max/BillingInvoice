import { useState, useRef } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import {
  Box,
  Typography,
  Paper,
} from '@mui/material';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';

const UploadSection = ({ onFileUpload, onBillingTypeChange, billingType, isProcessing }) => {
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef(null);

  const handleDragEnter = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);

    const files = Array.from(e.dataTransfer.files);
    if (files.length > 0) {
      handleFileSelect(files[0]);
    }
  };

  const handleFileSelect = (file) => {
    const validExtensions = ['.csv', '.xlsx', '.xls', '.pdf', '.doc', '.docx'];
    const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
    
    if (file && validExtensions.includes(fileExtension)) {
      onFileUpload(file);
    } else {
      alert('Please upload a CSV, XLSX, PDF, or DOC/DOCX file');
    }
  };

  const handleFileInputChange = (e) => {
    const files = Array.from(e.target.files);
    if (files.length > 0) {
      handleFileSelect(files[0]);
    }
  };

  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.5 }}
      className="w-full mb-6"
    >
      <Paper
        elevation={0}
        sx={{
          background: 'rgba(255, 255, 255, 0.9)',
          backdropFilter: 'blur(25px)',
          WebkitBackdropFilter: 'blur(25px)',
          border: '1px solid rgba(255, 255, 255, 0.6)',
          boxShadow: '0 8px 32px 0 rgba(31, 38, 135, 0.2)',
          p: { xs: 2, sm: 3 },
          borderRadius: '20px',
        }}
      >
        {/* File Upload Dropzone */}
        <div
          onDragEnter={handleDragEnter}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
          style={{
            border: isDragging ? '2px dashed #667eea' : '2px dashed rgba(102, 126, 234, 0.4)',
            borderRadius: '16px',
            padding: '2rem 1.5rem',
            textAlign: 'center',
            cursor: isProcessing ? 'not-allowed' : 'pointer',
            transition: 'all 0.3s ease',
            backgroundColor: isDragging ? 'rgba(102, 126, 234, 0.1)' : 'rgba(102, 126, 234, 0.05)',
            transform: isDragging ? 'scale(1.02)' : 'scale(1)',
            opacity: isProcessing ? 0.6 : 1,
            pointerEvents: isProcessing ? 'none' : 'auto',
          }}
        >
          <input
            ref={fileInputRef}
            type="file"
            accept=".csv,.xlsx,.xls,.pdf,.doc,.docx"
            onChange={handleFileInputChange}
            className="hidden"
            disabled={isProcessing}
          />
          
          <motion.div
            whileHover={{ scale: 1.1 }}
            whileTap={{ scale: 0.95 }}
          >
            <CloudUploadIcon sx={{ fontSize: 60, color: '#667eea', mb: 2 }} />
          </motion.div>
          
          <Typography 
            variant="h6" 
            sx={{ 
              color: '#333333', 
              mb: 2, 
              fontSize: { xs: '1rem', sm: '1.125rem' },
              fontWeight: 600,
            }}
          >
            {isDragging ? 'Drop file here' : 'Click or drag file to upload'}
          </Typography>
          
          <Typography 
            variant="body2" 
            sx={{ 
              color: '#666666', 
              fontSize: '0.875rem',
            }}
          >
            Supports CSV, XLSX, PDF, and DOC/DOCX files (1 lakh+ records)
          </Typography>

          <AnimatePresence>
            {isProcessing && (
              <motion.div
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                className="mt-4"
              >
                <Typography variant="body2" sx={{ color: '#667eea', fontWeight: 500 }}>
                  Processing file...
                </Typography>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </Paper>
    </motion.div>
  );
};

export default UploadSection;

