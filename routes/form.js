const express = require('express');
const { 
  guardarProgreso, 
  createNewRequest, 
  getRequests, 
  getActiveRequests, 
  getCompletedRequests, 
  getFormDataForm2,
  guardarGastos,
} = require('../controllers/formController');
const { verifyToken, isAdmin } = require('../middleware/auth');
const router = express.Router();
const multer = require('multer');
const upload = multer({ dest: '/tmp/uploads/' });

router.post('/guardarProgreso', upload.single('pieza_grafica'), guardarProgreso);
router.post('/guardarGastos', guardarGastos);
router.post('/createNewRequest', createNewRequest);
router.get('/getRequests', getRequests);
router.get('/getActiveRequests', getActiveRequests);
router.get('/getCompletedRequests', getCompletedRequests);
router.get('/getFormDataForm2', getFormDataForm2);


module.exports = router;