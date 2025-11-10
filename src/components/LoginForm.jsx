import { useState } from 'react';
import { motion } from 'framer-motion';
import {
  Box,
  Button,
  Paper,
  TextField,
  Typography,
  CircularProgress,
} from '@mui/material';

const LoginForm = ({ onSubmit, isSubmitting = false }) => {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [formError, setFormError] = useState('');
  const [usernameError, setUsernameError] = useState('');
  const [passwordError, setPasswordError] = useState('');

  const handleSubmit = async (event) => {
    event.preventDefault();
    setFormError('');
    setUsernameError('');
    setPasswordError('');

    const trimmedUsername = username.trim();
    const trimmedPassword = password;

    let hasError = false;

    if (!trimmedUsername) {
      setUsernameError('Username is required.');
      hasError = true;
    }

    if (!trimmedPassword) {
      setPasswordError('Password is required.');
      hasError = true;
    }

    if (hasError) {
      setFormError('Please fix the highlighted fields.');
      return;
    }

    try {
      const result = await onSubmit({
        username: trimmedUsername,
        password: trimmedPassword,
      });

      if (!result?.success) {
        setFormError(result?.message || 'Invalid username or password.');
        setUsernameError('Check your username.');
        setPasswordError('Check your password.');
      }
    } catch (err) {
      console.error('Login error:', err);
      setFormError('Something went wrong. Try again.');
    }
  };

  return (
    <Box
      sx={{
        display: 'flex',
        justifyContent: 'center',
        alignItems: 'center',
        minHeight: '70vh',
      }}
    >
      <motion.div
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.5 }}
        style={{ width: '100%', maxWidth: 420 }}
      >
        <Paper
          elevation={0}
          sx={{
            p: 4,
            borderRadius: '20px',
            background: 'rgba(255, 255, 255, 0.9)',
            backdropFilter: 'blur(20px)',
            WebkitBackdropFilter: 'blur(20px)',
            border: '1px solid rgba(255, 255, 255, 0.6)',
            boxShadow: '0 8px 32px 0 rgba(31, 38, 135, 0.2)',
          }}
        >
          <Typography
            variant="h5"
            sx={{
              mb: 2,
              textAlign: 'center',
              fontWeight: 600,
              color: '#333333',
            }}
          >
            Sign In
          </Typography>
          <Typography
            variant="body2"
            sx={{ mb: 3, textAlign: 'center', color: '#666666' }}
          >
            Enter your credentials to access the invoice generator.
          </Typography>
          <Box
            component="form"
            onSubmit={handleSubmit}
            sx={{ display: 'flex', flexDirection: 'column', gap: 2 }}
          >
            <TextField
              label="Username"
              value={username}
              onChange={(event) => {
                setUsername(event.target.value);
                if (formError) setFormError('');
                if (usernameError) setUsernameError('');
              }}
              fullWidth
              autoComplete="username"
              disabled={isSubmitting}
              error={Boolean(usernameError)}
              helperText={usernameError}
            />
            <TextField
              label="Password"
              type="password"
              value={password}
              onChange={(event) => {
                setPassword(event.target.value);
                if (formError) setFormError('');
                if (passwordError) setPasswordError('');
              }}
              fullWidth
              autoComplete="current-password"
              disabled={isSubmitting}
              error={Boolean(passwordError)}
              helperText={passwordError}
            />

            {formError && (
              <Typography
                variant="body2"
                sx={{ color: '#e57373', textAlign: 'center' }}
              >
                {formError}
              </Typography>
            )}

            <Button
              type="submit"
              variant="contained"
              disabled={isSubmitting}
              sx={{
                mt: 1,
                py: 1.2,
                borderRadius: '12px',
                background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                boxShadow: '0 10px 20px rgba(102, 126, 234, 0.2)',
                '&:hover': {
                  background: 'linear-gradient(135deg, #5a67d8 0%, #6b46c1 100%)',
                  boxShadow: '0 12px 24px rgba(102, 126, 234, 0.25)',
                },
              }}
            >
              {isSubmitting ? (
                <CircularProgress size={24} sx={{ color: 'white' }} />
              ) : (
                'Login'
              )}
            </Button>
          </Box>
        </Paper>
      </motion.div>
    </Box>
  );
};

export default LoginForm;

