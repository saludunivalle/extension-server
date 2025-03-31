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

// Inicializar express y middleware
const app = express();
const upload = multer({ dest: tempDir });
const PORT = process.env.PORT || 3001;

// Mover esta configuración antes de cualquier otro middleware
app.use(cors({
  origin: ['http://localhost:5173', 'https://siac-extension-form.vercel.app'], 
  credentials: true,
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With', 'Accept', 'Origin']
}));

// Configurar middleware
app.use(bodyParser.json());

// Configuración mejorada de CORS
app.use(cors({
  origin: ['http://localhost:5173', 'https://siac-extension-form.vercel.app'], 
  credentials: true,
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With', 'Accept', 'Origin']
}));

// Middleware específico para solicitudes OPTIONS
app.options('*', (req, res) => {
  res.header('Access-Control-Allow-Origin', req.headers.origin);
  res.header('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE,OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With, Accept, Origin');
  res.header('Access-Control-Allow-Credentials', 'true');
  res.status(200).end();
});

// Asegurar que las cabeceras CORS estén en todas las respuestas
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', req.headers.origin);
  res.header('Access-Control-Allow-Credentials', 'true');
  res.header('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE,OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
  next();
});

// Configurar middleware de sesión
app.use(session({
  secret: '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',
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