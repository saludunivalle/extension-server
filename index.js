// Importaciones y Configuraciones Iniciales
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const { config } = require('dotenv');
const multer = require('multer');
const os = require('os');
const path = require('path');
const tempDir = path.join(os.tmpdir(), 'extension-server-uploads');

const { jwtClient, oAuth2Client } = require('./config/google');
const routes = require('./routes');

config();

// Inicializar express y middleware
const app = express();
const upload = multer({ dest: tempDir });
const PORT = process.env.PORT || 3001;

// Configurar middleware
app.use(bodyParser.json());
app.use(cors({
  origin: '*', // Permite todas las orígenes durante desarrollo
  credentials: true
}));

// Importar servicios
const driveService = require('./services/driveService');
const reportService = require('./services/reportService');
const sheetsService = require('./services/sheetsService');

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