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
  guardarForm2Paso3,
  enviarSolicitudRevision,
  getSolicitudesRevisionAdmin,
  aprobarSolicitudAdmin,
  aprobarSolicitudCompletaAdmin,
  enviarCorreccionesAdmin,
  getEstadoRevisionSolicitud,
  guardarComentarioPaso
} = require('../controllers/formController');
const { verifyToken, isAdmin } = require('../middleware/auth');
const { loadProgressMiddleware } = require('../middleware/progressMiddleware');
const router = express.Router();
const multer = require('multer');
const upload = multer({ dest: '/tmp/uploads/' });
const uploadAdjuntos = upload.fields([
  { name: 'pieza_grafica', maxCount: 1 },
  { name: 'archivo_fondo_comun', maxCount: 1 }
]);

router.use(loadProgressMiddleware);

router.post('/guardarProgreso', uploadAdjuntos, guardarProgreso);
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
router.post('/guardarForm2Paso3', uploadAdjuntos, guardarForm2Paso3); // Añadir la ruta para guardarForm2Paso3
router.post('/enviarFormulariosRevision', enviarSolicitudRevision);
router.post('/enviarSolicitudRevision', enviarSolicitudRevision);
router.get('/admin/solicitudesRevision', getSolicitudesRevisionAdmin);
router.post('/admin/aprobarFormularios', aprobarSolicitudAdmin);
router.post('/admin/aprobarSolicitud', aprobarSolicitudAdmin);
router.post('/admin/aprobarSolicitudCompleta', aprobarSolicitudCompletaAdmin);
router.post('/admin/enviarCorrecciones', enviarCorreccionesAdmin);
router.get('/estadoRevisionSolicitud', getEstadoRevisionSolicitud);
router.post('/guardarComentarioPaso', guardarComentarioPaso);

module.exports = router;