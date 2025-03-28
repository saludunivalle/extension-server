const express = require('express');
const router = express.Router();
const progressController = require('../controllers/progressController');
const authRoutes = require('./auth');
const userRoutes = require('./user');
const formRoutes = require('./form');
const reportRoutes = require('./report');
const otherRoutes = require('./other');
const riskRoutes = require('./risk');

// SIN prefijos para mantener compatibilidad con el frontend
router.use('/', authRoutes);
router.use('/', userRoutes);
router.use('/', formRoutes);
router.use('/', reportRoutes);
router.use('/', otherRoutes);
router.use('/', riskRoutes);
router.get('/:id_solicitud', progressController.getProgress);
router.put('/:id_solicitud', progressController.updateProgress);

module.exports = router;