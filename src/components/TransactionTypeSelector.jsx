import { motion } from 'framer-motion';
import { Box, Typography, Paper } from '@mui/material';
import TrendingUpIcon from '@mui/icons-material/TrendingUp';
import TrendingDownIcon from '@mui/icons-material/TrendingDown';

import ViewModuleIcon from '@mui/icons-material/ViewModule';

const TransactionTypeSelector = ({ onSelect, selectedType }) => {
  const options = [
    {
      type: 'all',
      label: 'All Transactions',
      icon: ViewModuleIcon,
      description: 'View all credit and debit transactions',
      color: '#667eea',
    },
    {
      type: 'credit',
      label: 'Credit Details',
      icon: TrendingUpIcon,
      description: 'View income/received transactions',
      color: '#4caf50',
    },
    {
      type: 'debit',
      label: 'Debit Details',
      icon: TrendingDownIcon,
      description: 'View expense/paid transactions',
      color: '#f44336',
    },
  ];

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
        <Typography
          sx={{
            color: '#333333',
            fontWeight: 600,
            fontSize: '1.125rem',
            mb: 3,
            textAlign: 'center',
          }}
        >
          Select Transaction Type
        </Typography>
        <Box
          sx={{
            display: 'grid',
            gridTemplateColumns: { xs: '1fr', sm: 'repeat(3, 1fr)' },
            gap: 2,
          }}
        >
          {options.map((option) => (
            <motion.div
              key={option.type}
              whileHover={{ scale: 1.02 }}
              whileTap={{ scale: 0.98 }}
            >
              <Paper
                elevation={0}
                onClick={() => onSelect(option.type)}
                sx={{
                  background: selectedType === option.type
                    ? `linear-gradient(135deg, ${option.color}15 0%, ${option.color}25 100%)`
                    : 'rgba(255, 255, 255, 0.8)',
                  backdropFilter: 'blur(20px)',
                  WebkitBackdropFilter: 'blur(20px)',
                  border: selectedType === option.type
                    ? `2px solid ${option.color}`
                    : '2px solid rgba(102, 126, 234, 0.3)',
                  borderRadius: '16px',
                  p: 3,
                  cursor: 'pointer',
                  transition: 'all 0.3s ease',
                  position: 'relative',
                  '&:hover': {
                    background: `linear-gradient(135deg, ${option.color}10 0%, ${option.color}20 100%)`,
                    borderColor: option.color,
                    boxShadow: `0 8px 24px ${option.color}30`,
                  },
                }}
              >
                <Box sx={{ display: 'flex', flexDirection: 'column', alignItems: 'center', textAlign: 'center' }}>
                  <Box
                    sx={{
                      background: `linear-gradient(135deg, ${option.color} 0%, ${option.color}dd 100%)`,
                      borderRadius: '12px',
                      p: 2,
                      mb: 2,
                    }}
                  >
                    <option.icon sx={{ fontSize: 40, color: '#ffffff' }} />
                  </Box>
                  <Typography
                    sx={{
                      color: '#333333',
                      fontWeight: 700,
                      fontSize: '1.125rem',
                      mb: 1,
                    }}
                  >
                    {option.label}
                  </Typography>
                  <Typography
                    sx={{
                      color: '#666666',
                      fontSize: '0.875rem',
                    }}
                  >
                    {option.description}
                  </Typography>
                </Box>
              </Paper>
            </motion.div>
          ))}
        </Box>
      </Paper>
    </motion.div>
  );
};

export default TransactionTypeSelector;

