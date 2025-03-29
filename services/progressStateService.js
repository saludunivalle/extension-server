const redis = require('redis');
const sheetsService = require('./sheetsService'); // Importar sheetsService
const fs = require('fs'); // Para el caché local
const path = require('path');

class ProgressStateService {
  constructor() {
    this.redisAvailable = true; // Asumir que Redis está disponible inicialmente
    this.client = redis.createClient();
    this.client.connect().then(() => {
      console.log('Connected to Redis!');
      this.redisAvailable = true;
      this.scheduleRedisReconnect(); // Intentar reconectar periódicamente
      this.schedulePeriodicSync(); // Iniciar sincronización periódica
    }).catch(err => {
      console.error('Failed to connect to Redis, using fallback:', err);
      this.redisAvailable = false;
    });
    this.syncing = false; // Flag para evitar sincronizaciones concurrentes
    this.cacheDir = path.join(__dirname, 'progress_cache'); // Directorio para el caché local
    if (!fs.existsSync(this.cacheDir)) {
      fs.mkdirSync(this.cacheDir);
    }
    this.scheduleCacheCleanup(); // Limpieza del caché
    this.sheetsQueue = [];
    this.isProcessingSheets = false;
  }

  // Intenta reconectar a Redis periódicamente
  scheduleRedisReconnect() {
    setInterval(() => {
      if (!this.redisAvailable) {
        console.log('Attempting to reconnect to Redis...');
        this.client.connect().then(() => {
          console.log('Successfully reconnected to Redis!');
          this.redisAvailable = true;
        }).catch(err => {
          console.error('Failed to reconnect to Redis:', err);
        });
      }
    }, 60000); // Intenta reconectar cada 60 segundos
  }

  // Sincronización periódica con Google Sheets
  schedulePeriodicSync() {
    setInterval(async () => {
      if (!this.syncing) {
        this.syncing = true;
        try {
          console.log('Starting periodic sync with Google Sheets...');
          const keys = await this.getAllKeys();
          for (const key of keys) {
            const solicitudId = key.replace('progress:', '');
            try {
              await this.syncWithSheets(solicitudId);
            } catch (syncError) {
              console.error(`Error syncing solicitud ${solicitudId}:`, syncError);
              this.logError(`Error syncing solicitud ${solicitudId}: ${syncError.message}`);
            }
          }
          console.log('Periodic sync completed.');
        } catch (error) {
          console.error('Error during periodic sync:', error);
          this.logError(`Error during periodic sync: ${error.message}`);
        } finally {
          this.syncing = false;
        }
      } else {
        console.log('Sync already in progress, skipping...');
      }
    }, 300000); // Sincroniza cada 5 minutos
  }

  // Obtener todas las claves de Redis
  async getAllKeys() {
    return new Promise((resolve, reject) => {
      this.client.keys('progress:*', (err, keys) => {
        if (err) {
          reject(err);
        } else {
          resolve(keys);
        }
      });
    });
  }

  async getProgress(solicitudId) {
    let progressData = null;

    // 1. Intentar obtener de Redis
    if (this.redisAvailable) {
      try {
        const data = await this.client.get(`progress:${solicitudId}`);
        if (data) {
          progressData = JSON.parse(data);
          console.log(`✅ Progreso obtenido de Redis para ${solicitudId}`);
          return progressData;
        }
      } catch (error) {
        console.error(`Error getting progress from Redis for ${solicitudId}:`, error);
        this.redisAvailable = false; // Marcar Redis como no disponible
      }
    }

    // 2. Si Redis falla, intentar obtener del caché local
    if (!progressData) {
      try {
        progressData = this.loadFromCache(solicitudId);
        if (progressData) {
          console.log(`✅ Progreso obtenido del caché local para ${solicitudId}`);
          return progressData;
        }
      } catch (error) {
        console.error(`Error getting progress from cache for ${solicitudId}:`, error);
        this.logError(`Error getting progress from cache for ${solicitudId}: ${error.message}`);
      }
    }

    // 3. Si el caché local falla, obtener de Google Sheets
    if (!progressData) {
      console.log(`Redis y caché local no disponibles, obteniendo progreso de Google Sheets para ${solicitudId}`);
      try {
        progressData = await this.loadFromSheets(solicitudId);
        if (progressData) {
          console.log(`✅ Progreso obtenido de Google Sheets para ${solicitudId}`);
          return progressData;
        }
      } catch (error) {
        console.error(`Error getting progress from Google Sheets for ${solicitudId}:`, error);
        this.logError(`Error getting progress from Google Sheets for ${solicitudId}: ${error.message}`);
      }
    }

    // 4. Si todo falla, devolver valores por defecto
    console.warn(`⚠️ No se pudo obtener el progreso para ${solicitudId}, devolviendo valores por defecto.`);
    return {
      etapa_actual: 1,
      paso: 1,
      estado: 'En progreso',
      estado_formularios: {
        "1": "En progreso", "2": "En progreso",
        "3": "En progreso", "4": "En progreso"
      },
      version: 0 // Inicializar la versión
    };
  }

  async setProgress(solicitudId, progressData) {
    let success = false;
    let attempts = 0;
    const maxAttempts = 3;

    while (!success && attempts < maxAttempts) {
      attempts++;
      try {
        // 1. Obtener el estado actual con la versión
        const currentProgress = await this.getProgress(solicitudId);
        const expectedVersion = currentProgress.version || 0;

        // 2. Incrementar la versión para este intento
        const newVersion = expectedVersion + 1;
        progressData.version = newVersion;

        // 3. Guardar en Google Sheets
        try {
          await this.saveToSheets(solicitudId, progressData);
        } catch (saveToSheetsError) {
          console.error(`Error saving to sheets: ${saveToSheetsError}`);
          this.logError(`Error saving to sheets for ${solicitudId}: ${saveToSheetsError.message}`);
        }

        // 4. Guardar en Redis
        if (this.redisAvailable) {
          try {
            // Intentar guardar solo si la versión coincide
            const setResult = await this.client.set(`progress:${solicitudId}`, JSON.stringify(progressData), { XX: true });

            if (setResult === null) {
              console.warn(`Conflicto de versión detectado para ${solicitudId}. Reintento...`);
              continue; // Reintentar si hay un conflicto
            }
            console.log(`✅ Progreso guardado en Redis para ${solicitudId} con versión ${newVersion}`);
          } catch (error) {
            console.error(`Error setting progress in Redis for ${solicitudId}:`, error);
            this.redisAvailable = false; // Marcar Redis como no disponible
          }
        }

        // 5. Guardar en caché local
        try {
          this.saveToCache(solicitudId, progressData);
          console.log(`✅ Progreso guardado en caché local para ${solicitudId}`);
        } catch (error) {
          console.error(`Error saving to cache for ${solicitudId}:`, error);
        }

        success = true; // Marcar como éxito si llegamos aquí
      } catch (error) {
        console.error(`Error setting progress (attempt ${attempts}):`, error);
        this.logError(`Error setting progress for ${solicitudId} (attempt ${attempts}): ${error.message}`);
        // Puedes agregar una espera aquí antes de reintentar
        await new Promise(resolve => setTimeout(resolve, 100));
      }
    }

    if (!success) {
      console.error(`❌ No se pudo establecer el progreso para ${solicitudId} después de ${maxAttempts} intentos.`);
    }

    return success;
  }

  async deleteProgress(solicitudId) {
    if (this.redisAvailable) {
      try {
        await this.client.del(`progress:${solicitudId}`);
        return true;
      } catch (error) {
        console.error(`Error deleting progress from Redis for ${solicitudId}:`, error);
        this.redisAvailable = false; // Marcar Redis como no disponible
        return false;
      }
    }
    return false;
  }

  async saveToSheets(solicitudId, progressData, retryCount = 0) {
    try {
      const { etapa_actual, paso, estado, estado_formularios } = progressData;

      // Obtener los datos actuales de ETAPAS
      const client = sheetsService.getClient();
      const etapasResponse = await client.spreadsheets.values.get({
        spreadsheetId: sheetsService.spreadsheetId,
        range: 'ETAPAS!A:I'
      });
      const etapasRows = etapasResponse.data.values || [];

      // Buscar la fila que corresponde al id_solicitud
      let filaEtapas = etapasRows.findIndex(row => row[0] === solicitudId.toString());

      if (filaEtapas === -1) {
        console.log(`No se encontró la solicitud con ID ${solicitudId} en ETAPAS.`);
        return false;
      }

      filaEtapas += 1; // Ajustar índice a 1-based para Google Sheets

      // Actualizar la fila en Google Sheets
      await client.spreadsheets.values.update({
        spreadsheetId: sheetsService.spreadsheetId,
        range: `ETAPAS!E${filaEtapas}:I${filaEtapas}`,
        valueInputOption: 'RAW',
        resource: {
          values: [[
            etapa_actual,
            estado,
            etapasRows[filaEtapas - 1][6] || 'N/A', // Mantener nombre_actividad
            paso,
            JSON.stringify(estado_formularios)
          ]]
        }
      });

      console.log(`✅ Progreso guardado en Google Sheets para ${solicitudId}`);
      return true;

    } catch (error) {
      console.error(`Error saving progress to Google Sheets for ${solicitudId}:`, error);
      // Lógica de reintento con retroceso exponencial
      if (retryCount < 3) {
        const delay = Math.pow(2, retryCount) * 1000; // 1s, 2s, 4s
        console.log(`Retrying in ${delay}ms... (attempt ${retryCount + 1})`);
        await new Promise(resolve => setTimeout(resolve, delay));
        return this.saveToSheets(solicitudId, progressData, retryCount + 1);
      } else {
        console.error('Max retries reached, failing.');
        throw error;
      }
    }
  }

  async loadFromSheets(solicitudId) {
    try {
      // Obtener los datos actuales de ETAPAS
      const client = sheetsService.getClient();
      const etapasResponse = await client.spreadsheets.values.get({
        spreadsheetId: sheetsService.spreadsheetId,
        range: 'ETAPAS!A:I'
      });
      const etapasRows = etapasResponse.data.values || [];

      // Buscar la fila que corresponde al id_solicitud
      const filaEtapas = etapasRows.find(row => row[0] === solicitudId.toString());

      if (!filaEtapas) {
        console.log(`No se encontró la solicitud con ID ${solicitudId} en ETAPAS.`);
        return { // Valores por defecto
          etapa_actual: 1,
          paso: 1,
          estado: 'En progreso',
          estado_formularios: {
            "1": "En progreso", "2": "En progreso",
            "3": "En progreso", "4": "En progreso"
          },
          version: 0
        };
      }

      const etapa_actual = parseInt(filaEtapas[4]) || 1;
      const estado = filaEtapas[5] || 'En progreso';
      const paso = parseInt(filaEtapas[7]) || 1;
      const estado_formularios = filaEtapas[8] ? JSON.parse(filaEtapas[8]) : {
        "1": "En progreso", "2": "En progreso",
        "3": "En progreso", "4": "En progreso"
      };
      const version = parseInt(filaEtapas[9]) || 0;

      return {
        etapa_actual,
        paso,
        estado,
        estado_formularios,
        version
      };

    } catch (error) {
      console.error(`Error loading progress from Google Sheets for ${solicitudId}:`, error);
      return { // Valores por defecto
        etapa_actual: 1,
        paso: 1,
        estado: 'En progreso',
        estado_formularios: {
          "1": "En progreso", "2": "En progreso",
          "3": "En progreso", "4": "En progreso"
        },
        version: 0
      };
    }
  }

  // Sincronizar con Google Sheets
  async syncWithSheets(solicitudId) {
    try {
      console.log(`Syncing progress for ${solicitudId} with Google Sheets...`);
      const redisData = await this.getProgress(solicitudId);
      const sheetsData = await this.loadFromSheets(solicitudId);

      // Comparar datos y actualizar Redis si es necesario
      if (JSON.stringify(redisData) !== JSON.stringify(sheetsData)) {
        console.log(`Detected difference for ${solicitudId}, updating Redis...`);
        await this.setProgress(solicitudId, sheetsData);
      } else {
        console.log(`No difference detected for ${solicitudId}, skipping update.`);
      }
    } catch (error) {
      console.error(`Error syncing progress for ${solicitudId}:`, error);
      this.logError(`Error syncing progress for ${solicitudId}: ${error.message}`);
      throw error; // Propagar el error para que schedulePeriodicSync lo capture
    }
  }

  // Guardar en caché local
  saveToCache(solicitudId, progressData) {
    const cacheFile = path.join(this.cacheDir, `${solicitudId}.json`);
    try {
      fs.writeFileSync(cacheFile, JSON.stringify(progressData));
    } catch (error) {
      console.error(`Error writing to cache file ${cacheFile}:`, error);
      this.logError(`Error writing to cache file ${cacheFile}: ${error.message}`);
    }
  }

  // Cargar del caché local
  loadFromCache(solicitudId) {
    const cacheFile = path.join(this.cacheDir, `${solicitudId}.json`);
    if (fs.existsSync(cacheFile)) {
      try {
        const data = fs.readFileSync(cacheFile, 'utf8');
        return JSON.parse(data);
      } catch (error) {
        console.error(`Error reading cache file ${cacheFile}:`, error);
        this.logError(`Error reading cache file ${cacheFile}: ${error.message}`);
        return null;
      }
    }
    return null;
  }

  // Registrar errores en un archivo
  logError(message) {
    const logFile = path.join(__dirname, 'error.log');
    const timestamp = new Date().toISOString();
    fs.appendFileSync(logFile, `[${timestamp}] ${message}\n`);
  }

  // Limpieza del caché
  scheduleCacheCleanup() {
    setInterval(() => {
      fs.readdir(this.cacheDir, (err, files) => {
        if (err) {
          console.error("Error reading cache directory:", err);
          this.logError(`Error reading cache directory: ${err.message}`);
          return;
        }

        files.forEach(file => {
          const filePath = path.join(this.cacheDir, file);
          fs.stat(filePath, (statErr, stats) => {
            if (statErr) {
              console.error(`Error getting file stats for ${filePath}:`, statErr);
              this.logError(`Error getting file stats for ${filePath}: ${statErr.message}`);
              return;
            }

            const now = new Date();
            const fileAge = now.getTime() - stats.mtime.getTime();
            const oneDay = 24 * 60 * 60 * 1000;

            if (fileAge > oneDay) {
              fs.unlink(filePath, unlinkErr => {
                if (unlinkErr) {
                  console.error(`Error deleting cache file ${filePath}:`, unlinkErr);
                  this.logError(`Error deleting cache file ${filePath}: ${unlinkErr.message}`);
                } else {
                  console.log(`Cache file ${filePath} deleted.`);
                }
              });
            }
          });
        });
      });
    }, 24 * 60 * 60 * 1000); // Limpiar cada 24 horas
  }

  enqueueSheetsUpdate(solicitudId, progressData) {
    this.sheetsQueue.push({ solicitudId, progressData });
    this.processSheetsQueue();
  }

  async processSheetsQueue() {
    if (this.isProcessingSheets) return;
    this.isProcessingSheets = true;

    while (this.sheetsQueue.length > 0) {
      const { solicitudId, progressData } = this.sheetsQueue.shift();
      try {
        await this.saveToSheets(solicitudId, progressData);
      } catch (error) {
        console.error(`Error processing sheets update for ${solicitudId}:`, error);
        this.logError(`Error processing sheets update for ${solicitudId}: ${error.message}`);
      }
    }

    this.isProcessingSheets = false;
  }
}

module.exports = new ProgressStateService();