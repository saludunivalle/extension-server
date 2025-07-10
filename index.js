// Importaciones y Configuraciones Iniciales
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const { config } = require('dotenv');
const multer = require('multer');
const os = require('os');
const path = require('path');
const session = require('express-session');
const tempDir = path.join(os.tmpdir(), 'extension-server-uploads');

const { jwtClient, oAuth2Client } = require('./config/google');
const routes = require('./routes');

config();

// Verificar variables de entorno críticas
if (process.env.NODE_ENV === 'production' && !process.env.SESSION_SECRET) {
  console.error('FATAL ERROR: La variable de entorno SESSION_SECRET no está definida.');
  process.exit(1); // Detener el servidor si el secreto no está en producción
}


// Inicializar express y middleware
const app = express();
const upload = multer({ dest: tempDir });
const PORT = process.env.PORT || 3001;

// ELIMINAR CONFIGURACIÓN DUPLICADA - Mantener solo UNA configuración CORS
app.use(cors({
  origin: ['http://localhost:5173', 'https://siac-extension-form.vercel.app'],
  credentials: true,
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With', 'Accept', 'Origin']
}));

// Configurar middleware
app.use(bodyParser.json());

// Configurar middleware de sesión
app.use(session({
  secret: process.env.SESSION_SECRET,
  resave: false,
  saveUninitialized: true,
  cookie: { secure: process.env.NODE_ENV === 'production' } // Solo HTTPS en producción
}));

// Importar servicios
const driveService = require('./services/driveService');
const reportService = require('./services/reportService');
const sheetsService = require('./services/sheetsService');
const progressRoutes = require('./routes/progress');

const { errorHandler, notFoundHandler } = require('./middleware/errorHandler');

// Usar el enrutador principal para todas las rutas
app.use('/', routes);

// Middleware para rutas no encontradas
app.use(notFoundHandler);

// Middleware para manejo de errores
app.use(errorHandler);

// Iniciar el servidor
app.listen(PORT, () => {
  console.log(`Servidor de extensión escuchando en el puerto ${PORT}`);
});