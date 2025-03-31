const express = require('express');
const { getProgress, updateProgress } = require('../controllers/progressController');
const router = express.Router();

router.get('/:id_solicitud', getProgress);
router.put('/:id_solicitud', updateProgress);

module.exports = router;