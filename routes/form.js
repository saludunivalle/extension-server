const express = require('express');
const { 
  guardarProgreso, 
  createNewRequest, 
  getRequests, 
  getActiveRequests, 
  getCompletedRequests, 
  getFormDataForm2,
  guardarGastos,
  actualizarPasoMaximo,
  validarProgresion,       
  actualizarProgresoGlobal,
  getLastId,
  guardarForm2Paso1,
  guardarForm2Paso2,
  guardarForm2Paso3  // Añadir guardarForm2Paso3 aquí
} = require('../controllers/formController');
const { verifyToken, isAdmin } = require('../middleware/auth');
const { loadProgressMiddleware } = require('../middleware/progressMiddleware');
const router = express.Router();
const multer = require('multer');
const upload = multer({ dest: '/tmp/uploads/' });

router.use(loadProgressMiddleware);

router.post('/guardarProgreso', upload.single('pieza_grafica'), guardarProgreso);
router.post('/guardarGastos', guardarGastos);
router.post('/createNewRequest', createNewRequest);
router.get('/getRequests', getRequests);
router.get('/getActiveRequests', getActiveRequests);
router.get('/getCompletedRequests', getCompletedRequests);
router.get('/getFormDataForm2', getFormDataForm2);
router.post('/actualizarPasoMaximo', actualizarPasoMaximo);
router.post('/progreso-actual', validarProgresion);
router.post('/actualizacion-progreso', actualizarProgresoGlobal);
router.get('/getLastId', getLastId);
router.post('/guardarForm2Paso1', guardarForm2Paso1); // Añadir esta ruta también
router.post('/guardarForm2Paso2', guardarForm2Paso2);
router.post('/guardarForm2Paso3', guardarForm2Paso3); // Añadir la ruta para guardarForm2Paso3

module.exports = router;