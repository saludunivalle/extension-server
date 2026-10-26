const express = require('express');
const multer = require('multer');
const os = require('os');
const path = require('path');
const { verifyToken } = require('../middleware/auth');
const controller = require('../controllers/workflowController');
const router = express.Router();
const upload = multer({ dest: path.join(os.tmpdir(), 'extension-workflow-uploads') });

router.use(verifyToken);
router.get('/cola', controller.getQueue);
router.get('/solicitudes/consulta', controller.getConsultationQueue);
router.get('/firmas/pendientes', controller.getPendingSignatures);
router.post('/:idSolicitud/tomar', controller.take);
router.post('/:idSolicitud/decision', controller.decide);
router.get('/:idSolicitud/observaciones', controller.getObservations);
router.post('/:idSolicitud/observaciones', controller.addObservation);
router.patch('/:idSolicitud/observaciones/:idObservacion', controller.updateObservation);
router.post('/:idSolicitud/firma', upload.single('evidencia_firma'), controller.sign);
router.post('/:idSolicitud/enviar-decanato', controller.sendToDecanato);
router.post('/:idSolicitud/firma/recordatorio', controller.remindSignature);
router.post('/cortes-decanato', upload.single('pdf_resolucion'), controller.createCorte);

module.exports = router;
