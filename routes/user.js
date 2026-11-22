const express = require('express');
const {
  saveUser,
  getDirectors,
  addDirector,
  updateDirector,
  deleteDirector,
} = require('../controllers/userController');
const { verifyToken, isRevisor1 } = require('../middleware/auth');
const router = express.Router();

router.post('/saveUser', saveUser);

// Módulo de administración de directores (exclusivo para Revisor 1)
router.get('/directors', verifyToken, isRevisor1, getDirectors);
router.post('/directors', verifyToken, isRevisor1, addDirector);
router.put('/directors', verifyToken, isRevisor1, updateDirector);
router.delete('/directors/:email', verifyToken, isRevisor1, deleteDirector);
router.delete('/directors', verifyToken, isRevisor1, deleteDirector);

module.exports = router;
