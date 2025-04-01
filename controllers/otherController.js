const { google } = require('googleapis');
const { jwtClient } = require('../config/google');
const sheetsService = require('../services/sheetsService');
const reportService = require('../services/reportService');
const { getDataWithCache, invalidateCache } = require('../utils/cacheUtils');

/**
 * Mapa de caché local para reducir llamadas a Google Sheets
 * @type {Map<string, {data: any, timestamp: number}>}
 */
const localCache = new Map();

/**
 * Obtiene datos con caché local (respaldo adicional al sistema de caché global)
 * @param {string} key - Clave de caché
 * @param {Function} fetchData - Función para obtener datos si no están en caché
 * @param {number} ttlSeconds - Tiempo de vida en segundos (60 segundos por defecto)
 * @returns {Promise<any>} - Datos desde caché o frescos
 */
const getWithLocalCache = async (key, fetchData, ttlSeconds = 60) => {
  const now = Date.now();
  
  // Primero intentar usar la caché local (más rápida que Redis/caché global)
  if (localCache.has(key)) {
    const cached = localCache.get(key);
    if (now - cached.timestamp < ttlSeconds * 1000) {
      console.log(`📋 CACHÉ LOCAL: Usando datos en caché local para ${key}`);
      return cached.data;
    }
  }
  
  // Si no está en caché local, intentar con caché global
  try {
    return await getDataWithCache(key, fetchData, ttlSeconds/60); // ttlSeconds/60 convierte a minutos
  } catch (error) {
    // Si hay error en caché global pero tenemos datos en caché local (incluso expirados)
    if (localCache.has(key)) {
      console.warn(`⚠️ Error en caché global, usando caché local expirado para ${key}`);
      return localCache.get(key).data;
    }
    throw error;
  }
};

/**
 * Obtiene el último ID de una hoja específica
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
*/
const getLastId = async (req, res) => {
  const { sheetName } = req.query;
  
  if (!sheetName) {
    return res.status(400).json({ error: 'Se requiere el nombre de la hoja' });
  }
  
  try {
    // Usar caché para esta operación
    const lastId = await getWithLocalCache(
      `lastId_${sheetName}`,
      async () => {
        const result = await sheetsService.getLastId(sheetName);
        return result;
      },
      300 // TTL de 5 minutos para últimos IDs (cambian lentamente)
    );
    
    res.status(200).json({ lastId });
  } catch (error) {
    console.error(`Error al obtener el último ID de ${sheetName}:`, error);
    res.status(500).json({ error: `Error al obtener el último ID de ${sheetName}` });
  }
};

/**
 * Obtiene datos de programas y oficinas desde la hoja de cálculo
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
*/
const getProgramasYOficinas = async (req, res) => {
  try {
    // Usar caché para esta operación (datos estáticos que casi nunca cambian)
    const datos = await getWithLocalCache(
      'programas_oficinas',
      async () => {
        const spreadsheetId = sheetsService.spreadsheetId;
        const client = sheetsService.getClient();

        const response = await client.spreadsheets.values.get({
          spreadsheetId,
          range: 'PROGRAMAS!A2:K500',
        });

        const rows = response.data.values;
        if (!rows || rows.length === 0) {
          throw new Error('No se encontraron datos en la hoja de Google Sheets');
        }

        const programas = [];
        const oficinas = new Set();

        rows.forEach(row => {
          if (row[4] || row[5] || row[6] || row[7]) {
            programas.push({
              Programa: row[0],
              Snies: row[1],
              Sede: row[2],
              Facultad: row[3],
              Escuela: row[4],
              Departamento: row[5],
              Sección: row[6] || 'General',
              PregradoPosgrado: row[7],
            });
          }

          if (row[9]) {
            oficinas.add(row[9]);
          }
        });

        return {
          programas,
          oficinas: Array.from(oficinas),
        };
      },
      3600 // TTL de 1 hora (datos muy estáticos)
    );
    
    res.status(200).json(datos);
  } catch (error) {
    console.error('Error al obtener datos de la hoja de Google Sheets:', error);
    res.status(500).json({ error: 'Error al obtener datos de la hoja de Google Sheets' });
  }
};

/**
 * Obtiene datos de una solicitud específica
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
*/
const getSolicitud = async (req, res) => {
  try {
    const { id_solicitud } = req.query;
    
    if (!id_solicitud) {
      return res.status(400).json({ error: 'El ID de la solicitud es requerido' });
    }
    
    // Clave de caché que incluye el ID de solicitud
    const cacheKey = `solicitud_data_${id_solicitud}`;
    
    try {
      // Usar caché con ambos sistemas (local y global)
      const resultados = await getWithLocalCache(
        cacheKey,
        async () => {
          // Usar el servicio de hojas
          const hojas = sheetsService.reportSheetDefinitions || reportService.reportSheetDefinitions;
          const datos = await sheetsService.getSolicitudData(id_solicitud, hojas);
          
          // Guardar en caché local antes de retornar
          localCache.set(cacheKey, {
            data: datos,
            timestamp: Date.now()
          });
          
          return datos;
        },
        120 // TTL de 2 minutos para datos de solicitud
      );
      
      // Log para debugging
      console.log(`📋 Datos obtenidos para solicitud ${id_solicitud} (${Object.keys(resultados).length} hojas)`);
      
      res.status(200).json(resultados);
    } catch (sheetError) {
      console.error('Error específico de Google Sheets:', sheetError);
      
      // Si es un error de cuota, intentar devolver datos de caché local
      if (sheetError.code === 429 || 
          (sheetError.response && sheetError.response.status === 429) ||
          sheetError.message?.includes('Quota exceeded')) {
        
        console.log('Límite de API excedido, verificando datos en caché local...');
        
        // Intentar recuperar de caché local cualquier dato disponible (incluso expirado)
        if (localCache.has(cacheKey)) {
          console.log('✅ Encontrados datos en caché local para respaldo');
          const cachedData = localCache.get(cacheKey).data;
          
          return res.status(200).json({
            ...cachedData,
            _fromLocalCache: true,
            _cacheTime: localCache.get(cacheKey).timestamp
          });
        }
        
        // Si no hay caché, devolver respuesta mínima
        return res.status(200).json({ 
          message: 'Datos no disponibles temporalmente por límites de API', 
          data: {},
          limitExceeded: true
        });
      }
      
      // Para otros errores, seguir devolviendo 500
      throw sheetError;
    }
  } catch (error) {
    console.error('Error al obtener los datos de la solicitud:', error);
    res.status(500).json({ 
      error: 'Error al obtener los datos de la solicitud',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};

/**
 * Limpia la caché para una solicitud específica
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */
const clearSolicitudCache = async (req, res) => {
  try {
    const { id_solicitud } = req.params;
    
    if (!id_solicitud) {
      return res.status(400).json({ error: 'El ID de la solicitud es requerido' });
    }
    
    // Limpiar la caché para esta solicitud
    const cacheKey = `solicitud_data_${id_solicitud}`;
    localCache.delete(cacheKey);
    invalidateCache(cacheKey);
    
    console.log(`🗑️ Caché limpiado para solicitud ${id_solicitud}`);
    res.status(200).json({ success: true, message: 'Caché limpiado exitosamente' });
  } catch (error) {
    console.error('Error al limpiar la caché:', error);
    res.status(500).json({ error: 'Error al limpiar la caché' });
  }
};

/**
 * Obtiene estadísticas del uso de caché
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */
const getCacheStats = async (req, res) => {
  try {
    const cacheStats = {
      localCacheSize: localCache.size,
      localCacheKeys: Array.from(localCache.keys()),
      // Si tienes el estado global de la caché disponible, puedes incluirlo aquí
    };
    
    res.status(200).json(cacheStats);
  } catch (error) {
    console.error('Error al obtener estadísticas de caché:', error);
    res.status(500).json({ error: 'Error al obtener estadísticas de caché' });
  }
};

module.exports = {
  getLastId,
  getProgramasYOficinas,
  getSolicitud,
  clearSolicitudCache, // Nueva función
  getCacheStats // Nueva función
};