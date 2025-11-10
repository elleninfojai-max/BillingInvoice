import { motion } from 'framer-motion';
import { Box, Button, Typography } from '@mui/material';
import logoImage from '../assets/images/logo.jpg';

const Navbar = ({ isAuthenticated = false, onLogout }) => {
  return (
    <motion.nav
      initial={{ y: -100, opacity: 0 }}
      animate={{ y: 0, opacity: 1 }}
      transition={{ duration: 0.5 }}
      style={{
        width: '100%',
        padding: '1rem 1rem',
        background: 'rgba(255, 255, 255, 0.85)',
        backdropFilter: 'blur(20px)',
        WebkitBackdropFilter: 'blur(20px)',
        borderBottom: '1px solid rgba(255, 255, 255, 0.5)',
        boxShadow: '0 8px 32px 0 rgba(31, 38, 135, 0.15)',
      }}
    >
      <Box
        sx={{
          maxWidth: '1280px',
          margin: '0 auto',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          position: 'relative',
          gap: { xs: 1, sm: 2, md: 3 },
          px: { xs: 0.5, sm: 2 },
        }}
      >
        {/* Logo - Left */}
        <motion.div
          whileHover={{ scale: 1.05 }}
          transition={{ type: 'spring', stiffness: 400, damping: 10 }}
          style={{ 
            flexShrink: 0,
            zIndex: 2,
          }}
        >
          <Box
            component="img"
            src={logoImage}
            alt="ELLEN Logo"
            sx={{
              height: { xs: '40px', sm: '45px', md: '50px' },
              width: 'auto',
              objectFit: 'contain',
              borderRadius: '8px',
              maxWidth: { xs: '70px', sm: '80px' },
            }}
          />
        </motion.div>

        {/* Title - Center */}
        <Typography
          variant="h4"
          component="h1"
          sx={{
            position: { xs: 'relative', sm: 'absolute' },
            left: { sm: '50%' },
            transform: { sm: 'translateX(-50%)' },
            fontSize: { xs: '0.875rem', sm: '1.25rem', md: '1.75rem', lg: '2.25rem' },
            fontWeight: 700,
            background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
            WebkitBackgroundClip: 'text',
            WebkitTextFillColor: 'transparent',
            backgroundClip: 'text',
            letterSpacing: { xs: '0.01em', sm: '0.02em', md: '0.05em' },
            lineHeight: 1.2,
            whiteSpace: { xs: 'normal', sm: 'nowrap' },
            textAlign: { xs: 'center', sm: 'left' },
            flex: { xs: 1, sm: 0 },
            px: { xs: 1, sm: 0 },
            maxWidth: { xs: 'calc(100% - 100px)', sm: 'none' },
            zIndex: 1,
          }}
        >
          <Box
            component="span"
            sx={{
              display: { xs: 'block', sm: 'inline' },
            }}
          >
            ELLEN
          </Box>
          <Box
            component="span"
            sx={{
              display: { xs: 'block', sm: 'inline' },
            }}
          >
            {' '}INVOICE MANAGER
          </Box>
        </Typography>

        {/* Spacer to balance the layout - Only on larger screens */}
        <Box
          sx={{
            display: 'flex',
            alignItems: 'center',
            gap: 1,
          }}
        >
          {isAuthenticated && (
            <Button
              variant="contained"
              onClick={onLogout}
              sx={{
                borderRadius: '12px',
                background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                boxShadow: '0 10px 20px rgba(102, 126, 234, 0.2)',
                fontWeight: 600,
                textTransform: 'none',
                px: 2.5,
                py: 0.75,
                '&:hover': {
                  background: 'linear-gradient(135deg, #5a67d8 0%, #6b46c1 100%)',
                  boxShadow: '0 12px 24px rgba(102, 126, 234, 0.25)',
                },
              }}
            >
              Logout
            </Button>
          )}
        </Box>
      </Box>
    </motion.nav>
  );
};

export default Navbar;

