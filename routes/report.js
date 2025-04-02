const express = require('express');
const { generateReport, downloadReport } = require('../controllers/reportController');
const { generateReport1 } = require('../controllers/report1Controller');
const { generateReport2 } = require('../controllers/report2Controller');
const { generateReport3 } = require('../controllers/report3Controller');


const router = express.Router();

// Rutas existentes
router.post('/generateReport', generateReport);
router.post('/downloadReport', downloadReport);

// Rutas específicas para reportes
router.post('/generateReport1', generateReport1);
router.post('/generateReport2', generateReport2);
router.post('/generateReport3', generateReport3);

module.exports = router;
