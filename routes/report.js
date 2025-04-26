const express = require('express');
const { generateReport, downloadReport } = require('../controllers/reportController');
const { generateReport1 } = require('../controllers/report1Controller');
const { generateReport2 } = require('../controllers/report2Controller');
const { generateReport3 } = require('../controllers/report3Controller');
const { generateReport4 } = require('../controllers/report4Controller');
const riskReportService = require('../services/riskReportService');
const router = express.Router();

// Rutas existentes
router.post('/generateReport', generateReport);
router.post('/downloadReport', downloadReport);

// Rutas específicas
router.post('/generateReport1', generateReport1);
router.post('/generateReport2', generateReport2);
router.post('/generateReport3', generateReport3);
router.post('/generateReport4', generateReport4);

// Ruta de prueba para riesgos dinámicos
router.post('/testRiskRows', async (req, res) => {
  try {
    const { fileId, riesgos } = req.body;
    
    if (!fileId || !riesgos) {
      return res.status(400).json({ error: 'fileId y riesgos son requeridos' });
    }
    
    // Clasificar riesgos por categoría
    const riesgosPorCategoria = {
      diseno: [],
      locacion: [],
      desarrollo: [],
      cierre: [],
      otros: []
    };
    
    // Categorizar según el campo 'categoria'
    riesgos.forEach(riesgo => {
      const cat = riesgo.categoria?.toLowerCase() || '';
      
      if (cat.includes('dise')) {
        riesgosPorCategoria.diseno.push(riesgo);
      } else if (cat.includes('loca')) {
        riesgosPorCategoria.locacion.push(riesgo);
      } else if (cat.includes('desa')) {
        riesgosPorCategoria.desarrollo.push(riesgo);
      } else if (cat.includes('cier')) {
        riesgosPorCategoria.cierre.push(riesgo);
      } else {
        riesgosPorCategoria.otros.push(riesgo);
      }
    });
    
    // Generar datos de filas dinámicas para cada categoría
    const reportData = {};
    
    // Posiciones de inserción para cada categoría
    const posiciones = {
      diseno: 'B18:H18',
      locacion: 'B24:H24',
      desarrollo: 'B35:H35',
      cierre: 'B38:H38',
      otros: 'B41:H41'
    };
    
    // Generar datos para cada categoría
    Object.keys(riesgosPorCategoria).forEach(categoria => {
      const riesgosCat = riesgosPorCategoria[categoria];
      
      if (riesgosCat.length > 0) {
        const { generateRiskRows } = require('../services/dynamicRows');
        const dynamicRowsData = generateRiskRows(riesgosCat, categoria, posiciones[categoria]);
        
        if (dynamicRowsData) {
          const key = `__FILAS_DINAMICAS_${categoria.toUpperCase()}__`;
          reportData[key] = dynamicRowsData;
        }
      }
    });
    
    // Generar reporte con riesgos dinámicos
    const result = await riskReportService.generateReportWithRisks(fileId, reportData);
    
    res.status(200).json({
      success: result,
      message: `Procesamiento de riesgos completado: ${result ? 'exitoso' : 'fallido'}`,
      stats: Object.keys(riesgosPorCategoria).reduce((acc, cat) => {
        acc[cat] = riesgosPorCategoria[cat].length;
        return acc;
      }, {})
    });
  } catch (error) {
    console.error('Error en prueba de riesgos dinámicos:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

module.exports = router;