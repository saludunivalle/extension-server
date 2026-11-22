const express = require('express');
const { googleAuth, verifySession } = require('../controllers/authController');
const { verifyToken } = require('../middleware/auth');
const router = express.Router();

router.post('/google', googleAuth);
router.get('/verify', verifyToken, verifySession);

module.exports = router;
