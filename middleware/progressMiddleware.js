const progressService = require('../services/progressStateService');

const loadProgressMiddleware = async (req, res, next) => {
  const id_solicitud = req.body.id_solicitud || req.query.id_solicitud;

  if (id_solicitud) {
    try {
      const progressState = await progressService.getProgress(id_solicitud);
      req.progressState = progressState;
    } catch (error) {
      console.error(`Error loading progress for ${id_solicitud}:`, error);
      // Manejar el error apropiadamente, tal vez enviando un mensaje al cliente
      return res.status(500).json({ success: false, error: 'Failed to load progress' });
    }
  } else {
    req.progressState = null;
  }
  next();
};

module.exports = {
  loadProgressMiddleware
};