const express = require('express');
const { getLastId, getProgramasYOficinas, getSolicitud } = require('../controllers/otherController');
const router = express.Router();

router.get('/getLastId', getLastId);
router.get('/getProgramasYOficinas', getProgramasYOficinas);
router.get('/getSolicitud', getSolicitud);

module.exports = router;