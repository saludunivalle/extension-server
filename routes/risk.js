const express = require('express');
const { 
  getRiesgos,
  addRiesgo,
  updateRiesgo,
  deleteRiesgo,
  getCategoriasRiesgo,
  migrarRiesgosForm3
} = require('../controllers/riskController');
const router = express.Router();

// Rutas para riesgos
router.get('/riesgos', getRiesgos);
router.post('/riesgos', addRiesgo);
router.put('/riesgos', updateRiesgo);
router.delete('/riesgos/:id_riesgo', deleteRiesgo);
router.get('/categorias-riesgo', getCategoriasRiesgo);
router.post('/migrar-riesgos-form3', migrarRiesgosForm3);

module.exports = router;