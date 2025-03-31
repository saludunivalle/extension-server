const loadProgressMiddleware = (req, res, next) => {
  req.progressState = req.session.progressState || {
    etapa_actual: 1,
    paso: 1,
    estado: 'En progreso',
    estado_formularios: {
      "1": "En progreso", "2": "En progreso",
      "3": "En progreso", "4": "En progreso"
    }
  };
  next();
};

module.exports = {
  loadProgressMiddleware
};