const express = require('express');
const { saveUser } = require('../controllers/userController');
const router = express.Router();

router.post('/saveUser', saveUser);

module.exports = router;