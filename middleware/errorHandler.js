/**
 * Middleware centralizado para manejo de errores
 * Permite una gestión consistente de errores en toda la aplicación
*/

// Clase base para errores personalizados de la aplicación
class AppError extends Error {
    constructor(message, statusCode, errorCode = 'ERROR_GENERAL') {
      super(message);
      this.statusCode = statusCode;
      this.errorCode = errorCode;
      this.isOperational = true; // Errores controlados vs no controlados
      Error.captureStackTrace(this, this.constructor);
    }
  }
  
  // Errores específicos que extienden la clase base
  class ValidationError extends AppError {
    constructor(message = 'Datos inválidos') {
      super(message, 400, 'VALIDATION_ERROR');
      this.name = 'ValidationError';
    }
  }
  
  class AuthenticationError extends AppError {
    constructor(message = 'No autenticado') {
      super(message, 401, 'AUTHENTICATION_ERROR');
      this.name = 'AuthenticationError';
    }
  }
  
  class AuthorizationError extends AppError {
    constructor(message = 'No autorizado') {
      super(message, 403, 'AUTHORIZATION_ERROR');
      this.name = 'AuthorizationError';
    }
  }
  
  class NotFoundError extends AppError {
    constructor(message = 'Recurso no encontrado') {
      super(message, 404, 'NOT_FOUND');
      this.name = 'NotFoundError';
    }
  }
  
  class GoogleAPIError extends AppError {
    constructor(message = 'Error en la API de Google') {
      super(message, 500, 'GOOGLE_API_ERROR');
      this.name = 'GoogleAPIError';
    }
  }
  
  /**
   * Middleware para capturar errores y responder de manera consistente
  */
  const errorHandler = (err, req, res, next) => {
    console.error('Error capturado:', err);
  
    // Si ya es uno de nuestros errores personalizados, usamos su información
    if (err instanceof AppError) {
      return res.status(err.statusCode).json({
        success: false,
        errorCode: err.errorCode,
        error: err.message,
        ...(process.env.NODE_ENV === 'development' && { stack: err.stack })
      });
    }
  
    // Manejar errores específicos de Express o Node
    if (err.name === 'SyntaxError' && err.message.includes('JSON')) {
      return res.status(400).json({
        success: false,
        errorCode: 'INVALID_JSON',
        error: 'Formato JSON inválido'
      });
    }
  
    if (err.name === 'MulterError') {
      return res.status(400).json({
        success: false,
        errorCode: 'FILE_UPLOAD_ERROR',
        error: `Error al subir archivo: ${err.message}`
      });
    }
  
    // Para errores en la API de Google
    if (err.message && (
      err.message.includes('Google') || 
      err.message.includes('sheets') ||
      err.message.includes('drive')
    )) {
      return res.status(500).json({
        success: false,
        errorCode: 'GOOGLE_API_ERROR',
        error: 'Error en la comunicación con servicios de Google'
      });
    }
  
    // Error genérico (no controlado)
    // En producción no mostramos detalles técnicos
    const isProd = process.env.NODE_ENV === 'production';
    
    return res.status(500).json({
      success: false,
      errorCode: 'INTERNAL_SERVER_ERROR',
      error: isProd ? 'Error interno del servidor' : err.message,
      ...(process.env.NODE_ENV === 'development' && { stack: err.stack })
    });
  };
  
  /**
   * Middleware para capturar rutas no encontradas
 */
  const notFoundHandler = (req, res, next) => {
    const error = new NotFoundError(`Ruta no encontrada: ${req.originalUrl}`);
    next(error);
  };
  
  module.exports = {
    errorHandler,
    notFoundHandler,
    AppError,
    ValidationError,
    AuthenticationError,
    AuthorizationError,
    NotFoundError,
    GoogleAPIError
  };