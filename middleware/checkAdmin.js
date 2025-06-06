module.exports = function checkAdmin(req, res, next) {
  if (!req.isAuthenticated()) {
    return res.redirect('/auth/login');
  }

  const email = req.user?.email || '';
  const domain = email.split('@')[1];

  // ✅ Allow if: (1) exact admin email OR (2) belongs to svsu.ac.in
  if (email === process.env.ADMIN_EMAIL || domain === 'svsu.ac.in') {
    return next();
  }

  return res.status(403).send('Access Denied: Admins only');
};