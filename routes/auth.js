const express = require('express');
const passport = require('passport');
const router = express.Router();

// Login Page
router.get('/login', (req, res) => {
  res.render('login'); // views/auth/login.ejs
});

// Google OAuth Login
router.get('/google', passport.authenticate('google', {
  scope: ['profile', 'email']
}));

// OAuth Callback
router.get('/google/callback',
  passport.authenticate('google', {
    failureRedirect: '/auth/login',
    failureMessage: true
  }),
  (req, res) => {
    const email = req.user.email;

    // ✅ Admin check
    if (req.user.isAdmin || email.endsWith('@svsu.ac.in')) {
      return res.redirect('/step1');
    }

    // ✅ Temporarily allow gmail and others for testing
    return res.redirect('/user/timetable');

    // ❌ In final version, you can uncomment this to restrict access:
    /*
    req.logout(() => {
      res.send('Access Denied: Only @svsu.ac.in users allowed.');
    });
    */
  }
);

// Logout
router.get('/logout', (req, res) => {
  req.logout(() => {
    res.redirect('/auth/login');
  });
});

module.exports = router;