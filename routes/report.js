const express = require('express');
const { generateReport, downloadReport } = require('../controllers/reportController');
const router = express.Router();

router.post('/generateReport', generateReport);
router.post('/downloadReport', downloadReport);

module.exports = router;
