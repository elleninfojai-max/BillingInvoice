import { motion } from 'framer-motion';
import { Paper, Typography, Box } from '@mui/material';
import AttachMoneyIcon from '@mui/icons-material/AttachMoney';
import ReceiptIcon from '@mui/icons-material/Receipt';
import TrendingUpIcon from '@mui/icons-material/TrendingUp';
import TrendingDownIcon from '@mui/icons-material/TrendingDown';

const SummaryCards = ({ totalCount = 0, creditTotal = 0, debitTotal = 0, netBalance = 0 }) => {
  const cards = [
    {
      title: 'Total Records',
      value: (totalCount || 0).toLocaleString(),
      icon: ReceiptIcon,
      color: 'from-blue-500 to-blue-600',
    },
    {
      title: 'Credit Total',
      value: `₹${(creditTotal || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
      icon: TrendingUpIcon,
      color: 'from-green-500 to-green-600',
    },
    {
      title: 'Debit Total',
      value: `₹${(debitTotal || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
      icon: TrendingDownIcon,
      color: 'from-red-500 to-red-600',
    },
    {
      title: 'Net Balance',
      value: `₹${(netBalance || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
      icon: AttachMoneyIcon,
      color: (netBalance || 0) >= 0 ? 'from-purple-500 to-purple-600' : 'from-orange-500 to-orange-600',
    },
  ];

  return (
    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 sm:gap-6 mb-6">
      {cards.map((card, index) => (
        <motion.div
          key={card.title}
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.5, delay: index * 0.1 }}
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
              borderRadius: '16px',
            }}
          >
            <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
              <Box>
                <Typography 
                  variant="body2" 
                  sx={{ 
                    color: '#666666', 
                    mb: 1,
                    fontSize: { xs: '0.875rem', sm: '1rem' },
                    fontWeight: 500,
                  }}
                >
                  {card.title}
                </Typography>
                <Typography 
                  variant="h5" 
                  sx={{ 
                    color: '#333333', 
                    fontWeight: 700,
                    fontSize: { xs: '1.25rem', sm: '1.5rem' },
                  }}
                >
                  {card.value}
                </Typography>
              </Box>
              <motion.div
                whileHover={{ scale: 1.1, rotate: 5 }}
                className={`bg-gradient-to-br ${card.color} p-3 sm:p-4 rounded-xl`}
              >
                <card.icon sx={{ fontSize: 32, color: 'white' }} />
              </motion.div>
            </Box>
          </Paper>
        </motion.div>
      ))}
    </div>
  );
};

export default SummaryCards;

