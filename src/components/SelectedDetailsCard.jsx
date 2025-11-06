import { motion } from 'framer-motion';
import { Paper, Typography, Box, Button } from '@mui/material';
import CheckCircleIcon from '@mui/icons-material/CheckCircle';
import ClearIcon from '@mui/icons-material/Clear';
import VisibilityIcon from '@mui/icons-material/Visibility';

const SelectedDetailsCard = ({ selectedRows, onClearAll, onView }) => {
  if (!selectedRows || selectedRows.length === 0) {
    return null;
  }

  return (
    <motion.div
      initial={{ opacity: 0, y: -20 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.5 }}
      className="mb-6"
    >
      <Paper
        elevation={0}
        sx={{
          background: 'rgba(255, 255, 255, 0.95)',
          backdropFilter: 'blur(25px)',
          WebkitBackdropFilter: 'blur(25px)',
          border: '1px solid rgba(147, 51, 234, 0.3)',
          boxShadow: '0 8px 32px 0 rgba(147, 51, 234, 0.15)',
          p: { xs: 2, sm: 3 },
          borderRadius: '16px',
        }}
      >
        <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 2 }}>
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 2, flex: 1 }}>
            <motion.div
              whileHover={{ scale: 1.1 }}
              className="bg-gradient-to-br from-purple-500 to-purple-600 p-3 rounded-xl"
            >
              <CheckCircleIcon sx={{ fontSize: 32, color: 'white' }} />
            </motion.div>
            <Typography
              variant="h6"
              sx={{
                color: '#333333',
                fontWeight: 700,
                fontSize: { xs: '1.125rem', sm: '1.25rem' },
              }}
            >
              Selected: {selectedRows.length} {selectedRows.length === 1 ? 'Record' : 'Records'}
            </Typography>
          </Box>
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
            <Button
              variant="contained"
              startIcon={<VisibilityIcon />}
              onClick={onView}
              sx={{
                backgroundColor: '#7c3aed',
                '&:hover': { backgroundColor: '#6d28d9' },
                fontWeight: 600,
                px: 3,
                py: 1,
                borderRadius: '12px',
                textTransform: 'none',
              }}
            >
              View
            </Button>
          <Button
            variant="outlined"
            startIcon={<ClearIcon />}
            onClick={onClearAll}
            sx={{
              borderColor: '#ef4444',
              color: '#ef4444',
              '&:hover': {
                borderColor: '#dc2626',
                backgroundColor: 'rgba(239, 68, 68, 0.1)',
              },
              fontWeight: 600,
              px: 3,
              py: 1,
              borderRadius: '12px',
              textTransform: 'none',
            }}
          >
            Clear All
          </Button>
          </Box>
        </Box>
      </Paper>
    </motion.div>
  );
};

export default SelectedDetailsCard;

