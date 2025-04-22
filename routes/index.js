const express = require('express');
const router = express.Router();
const formRoutes = require('./form');
const reportRoutes = require('./report');
const userRoutes = require('./user');
const authRoutes = require('./auth');
const riskRoutes = require('./risk');
const otherRoutes = require('./other');
const formController = require('../controllers/formController');
const multer = require('multer');
const upload = multer({ dest: '/tmp/uploads/' });

// Middleware para verificar el token JWT
const { verifyToken } = require('../middleware/auth');

// Rutas que no requieren autenticación
router.use('/auth', authRoutes);
router.post('/saveUser', require('../controllers/userController').saveUser);

// Agregar las rutas de reporte aquí, antes del middleware de autenticación
router.post('/report/generateReport', require('../controllers/reportController').generateReport);
router.post('/report/downloadReport', require('../controllers/reportController').downloadReport);

// Otras rutas sin autenticación
router.get('/getActiveRequests', formController.getActiveRequests);
router.get('/getCompletedRequests', formController.getCompletedRequests);
router.get('/getFormDataForm2', formController.getFormDataForm2);
router.post('/actualizarPasoMaximo', formController.actualizarPasoMaximo);
router.post('/progreso-actual', formController.validarProgresion);
router.post('/actualizacion-progreso', formController.actualizarProgresoGlobal);
// Añadir esta línea para el endpoint que falta
router.post('/actualizacion-progreso-global', formController.actualizarProgresoGlobal);
router.get('/getLastId', formController.getLastId);
router.post('/guardarProgreso', upload.single('pieza_grafica'), formController.guardarProgreso);
router.post('/guardarGastos', formController.guardarGastos);
router.post('/createNewRequest', formController.createNewRequest);
router.get('/getRequests', formController.getRequests);
router.get('/getProgramasYOficinas', require('../controllers/otherController').getProgramasYOficinas);
router.get('/getSolicitud', require('../controllers/otherController').getSolicitud);
router.post('/guardarForm2Paso2', formController.guardarForm2Paso2);

// Añade esta línea como ruta pública (antes del middleware de autenticación)
router.get('/getGastos', formController.getGastos);

// Rutas de formulario (para mantener compatibilidad)
router.use('/form', formRoutes);

// Para mantener compatibilidad, deja las rutas de reporte también después
router.use('/report', reportRoutes);

// Rutas de usuario
router.use('/user', userRoutes);

// Añadir antes de las líneas que definen router.use('/risk', riskRoutes)
router.get('/riesgos', require('../controllers/riskController').getRiesgos);
router.post('/riesgos', require('../controllers/riskController').addRiesgo);
router.put('/riesgos', require('../controllers/riskController').updateRiesgo);
router.delete('/riesgos/:id_riesgo', require('../controllers/riskController').deleteRiesgo);

// Rutas de riesgo
router.use('/risk', riskRoutes);

router.use('/other', otherRoutes);

// Ruta de ejemplo
router.get('/', (req, res) => {
  res.send('¡Hola desde la ruta principal!');
});

module.exports = router;