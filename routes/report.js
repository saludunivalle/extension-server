const express = require('express');
const { generateReport, downloadReport } = require('../controllers/reportController');
const { generateReport1 } = require('../controllers/report1Controller');
const { generateReport2 } = require('../controllers/report2Controller');
const { generateReport3 } = require('../controllers/report3Controller');
const { generateReport4 } = require('../controllers/report4Controller');
const router = express.Router();

// Rutas existentes
router.post('/generateReport', generateReport);
router.post('/downloadReport', downloadReport);

// Rutas específicas
router.post('/generateReport1', generateReport1);
router.post('/generateReport2', generateReport2);
router.post('/generateReport3', generateReport3);
router.post('/generateReport4', generateReport4);

module.exports = router;
