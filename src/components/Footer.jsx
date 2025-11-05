import { motion } from 'framer-motion';
import { Box, Typography } from '@mui/material';

const Footer = () => {
  return (
    <motion.footer
      initial={{ y: 100, opacity: 0 }}
      animate={{ y: 0, opacity: 1 }}
      transition={{ duration: 0.5, delay: 0.2 }}
      style={{
        width: '100%',
        padding: '1.5rem 1rem',
        marginTop: '2rem',
        background: 'rgba(255, 255, 255, 0.9)',
        backdropFilter: 'blur(25px)',
        WebkitBackdropFilter: 'blur(25px)',
        borderTop: '1px solid rgba(255, 255, 255, 0.6)',
        boxShadow: '0 -8px 32px 0 rgba(31, 38, 135, 0.15)',
      }}
    >
      <Box
        sx={{
          maxWidth: '1280px',
          margin: '0 auto',
          textAlign: 'center',
        }}
      >
        <Typography
          sx={{
            fontSize: { xs: '1rem', sm: '1.125rem' },
            fontWeight: 600,
            mb: 1,
            color: '#333333',
          }}
        >
          ELLEN INFORMATION TECHNOLOGY SOLUTIONS PRIVATE LIMITED
        </Typography>
        <Typography
          sx={{
            fontSize: { xs: '0.875rem', sm: '1rem' },
            mb: 0.5,
            color: '#666666',
          }}
        >
          8/3, Athreyapuram 2nd Street, Choolaimedu, Chennai – 600094
        </Typography>
        <Typography
          sx={{
            fontSize: { xs: '0.875rem', sm: '1rem' },
            color: '#666666',
          }}
        >
          Email:{' '}
          <a
            href="mailto:elleninfojai@gmail.com"
            style={{
              color: '#667eea',
              textDecoration: 'underline',
              fontWeight: 500,
            }}
            onMouseEnter={(e) => (e.target.style.color = '#764ba2')}
            onMouseLeave={(e) => (e.target.style.color = '#667eea')}
          >
            elleninfojai@gmail.com
          </a>
        </Typography>
      </Box>
    </motion.footer>
  );
};

export default Footer;

