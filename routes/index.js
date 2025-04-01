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
const upload = multer({ dest: '/tmp/uploads/' });  // Añade esta línea

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
router.post('/progreso-actual', async (req, res) => {
  try {
    const { id_solicitud } = req.body;
    
    // Evitar hacer consultas innecesarias a Google Sheets si no hay ID
    if (!id_solicitud) {
      console.log('No se recibió ID de solicitud, devolviendo valores por defecto');
      return res.status(200).json({
        success: true,
        data: {
          etapa_actual: 1,
          paso: 1,
          estado: 'En progreso',
          estado_formularios: {
            "1": "En progreso", "2": "En progreso",
            "3": "En progreso", "4": "En progreso"
          }
        },
        isNewRequest: true
      });
    }
    
    try {
      // Intento interno de acceder a Google Sheets con su propio try-catch
      const client = require('../services/sheetsService').getClient();
      const etapasResponse = await client.spreadsheets.values.get({
        spreadsheetId: require('../services/sheetsService').spreadsheetId,
        range: 'ETAPAS!A2:I'
      });
      
      const etapasRows = etapasResponse.data.values || [];
      const filaActual = etapasRows.find(row => row[0] === id_solicitud?.toString());
      
      // Si no existe la solicitud, devolver valores por defecto
      if (!filaActual) {
        console.log(`No se encontró la solicitud ${id_solicitud}, devolviendo valores por defecto`);
        return res.status(200).json({
          success: true,
          data: {
            etapa_actual: 1,
            paso: 1,
            estado: 'En progreso',
            estado_formularios: {
              "1": "En progreso", "2": "En progreso",
              "3": "En progreso", "4": "En progreso"
            }
          },
          isNewRequest: true
        });
      }
      
      // Si existe, continuar con la validación normal
      return formController.validarProgresion(req, res);
      
    } catch (sheetError) {
      // Capturar errores específicos de Google Sheets aquí
      console.log('Error al acceder a Google Sheets:', sheetError.message);
      
      // Comprobar específicamente si es un error de cuota
      if (sheetError.code === 429 || 
          (sheetError.response && sheetError.response.status === 429) ||
          sheetError.message?.includes('Quota exceeded')) {
        console.log('Error de cuota excedida en Google Sheets, continuando con valores por defecto');
      }
      
      // Siempre retornar datos por defecto en caso de error de Sheets
      return res.status(200).json({
        success: true,
        data: {
          etapa_actual: 1,
          paso: 1,
          estado: 'En progreso',
          estado_formularios: {
            "1": "En progreso", "2": "En progreso",
            "3": "En progreso", "4": "En progreso"
          }
        },
        quota_warning: 'Se alcanzó el límite de solicitudes a Google Sheets.'
      });
    }
    
  } catch (outerError) {
    // Capturar cualquier otro error inesperado
    console.error('Error general en progreso-actual:', outerError);
    
    // Siempre devolver estado 200 con datos por defecto
    return res.status(200).json({
      success: true,
      data: {
        etapa_actual: 1,
        paso: 1,
        estado: 'En progreso',
        estado_formularios: {
          "1": "En progreso", "2": "En progreso",
          "3": "En progreso", "4": "En progreso"
        }
      },
      error_message: outerError.message
    });
  }
});
router.post('/actualizacion-progreso', formController.actualizarProgresoGlobal);
router.get('/getLastId', formController.getLastId);
router.post('/guardarProgreso', upload.single('pieza_grafica'), formController.guardarProgreso);
router.post('/guardarGastos', formController.guardarGastos);
router.post('/createNewRequest', formController.createNewRequest);
router.get('/getRequests', formController.getRequests);
router.get('/getProgramasYOficinas', require('../controllers/otherController').getProgramasYOficinas);
router.get('/getSolicitud', require('../controllers/otherController').getSolicitud);
router.post('/guardarForm2Paso2', formController.guardarForm2Paso2);
router.get('/getGastos', formController.getGastos);
router.post('/report/previewReport', require('../controllers/reportController').previewReport);

// Rutas de formulario (para mantener compatibilidad)
router.use('/form', formRoutes);
router.use('/report', reportRoutes);

// Rutas de usuario
router.use('/user', userRoutes);

// Rutas de riesgo
router.use('/risk', riskRoutes);

router.use('/other', otherRoutes);

// Ruta de ejemplo (asegúrate de que todas las rutas tengan un controlador)
router.get('/', (req, res) => {
  res.send('¡Hola desde la ruta principal!');
});

module.exports = router;