const express = require('express');
const { generateReport, downloadReport } = require('../controllers/reportController');
const { generateReport1 } = require('../controllers/report1Controller');
const router = express.Router();

// Rutas existentes
router.post('/generateReport', generateReport);
router.post('/downloadReport', downloadReport);

// Nueva ruta específica para reporte 1
router.post('/generateReport1', generateReport1);

module.exports = router;
