/**
 * Sistema de caché en memoria para reducir llamadas a la API de Google Sheets
 */

// Caché global
const memoryCache = {
  data: new Map(),
  ttls: new Map(),
  stats: {
    hits: 0,
    misses: 0,
    saved: 0 // Llamadas ahorradas estimadas
  }
};

/**
 * Obtiene datos con caché
 * @param {string} key - Clave única para los datos
 * @param {Function} dataFetcher - Función que obtiene los datos si no están en caché
 * @param {number} ttlMinutes - Tiempo de vida en minutos (por defecto: 5)
 * @returns {Promise<any>} Datos, ya sea de caché o recién obtenidos
 */
const getDataWithCache = async (key, dataFetcher, ttlMinutes = 5) => {
  // Convertir minutos a milisegundos
  const ttl = ttlMinutes * 60 * 1000;
  const now = Date.now();
  
  // Verificar si tenemos datos en caché y si aún son válidos
  if (memoryCache.data.has(key)) {
    const timestamp = memoryCache.ttls.get(key) || 0;
    
    // Si el caché es válido
    if (now - timestamp < ttl) {
      memoryCache.stats.hits++;
      console.log(`🔄 CACHÉ: Usando datos en caché para '${key}' (ahorro de llamada API)`);
      return memoryCache.data.get(key);
    }
    
    console.log(`⏱️ CACHÉ: Datos expirados para '${key}', actualizando...`);
  } else {
    console.log(`🔍 CACHÉ: Dato '${key}' no encontrado en caché, obteniendo...`);
  }
  
  memoryCache.stats.misses++;
  
  try {
    // Llamar a la función para obtener los datos
    const data = await dataFetcher();
    
    // Guardar en caché con timestamp actual
    memoryCache.data.set(key, data);
    memoryCache.ttls.set(key, now);
    
    // Registrar estadísticas
    logCacheStats();
    
    return data;
  } catch (error) {
    // En caso de error, intentar usar datos caducados como respaldo
    if (memoryCache.data.has(key)) {
      console.warn(`⚠️ CACHÉ: Error al obtener datos nuevos para '${key}', usando caché antiguo como respaldo`);
      return memoryCache.data.get(key);
    }
    
    // Si no hay respaldo, propagar el error
    throw error;
  }
};

/**
 * Invalida manualmente una entrada de caché
 * @param {string} key - Clave a invalidar
 */
const invalidateCache = (key) => {
  memoryCache.data.delete(key);
  memoryCache.ttls.delete(key);
  console.log(`🗑️ CACHÉ: Entrada '${key}' invalidada manualmente`);
};

/**
 * Registra estadísticas de uso de caché en consola
 */
const logCacheStats = () => {
  memoryCache.stats.saved = memoryCache.stats.hits;
  
  // Solo mostrar cada 10 operaciones (para no saturar logs)
  if ((memoryCache.stats.hits + memoryCache.stats.misses) % 10 === 0) {
    console.log(`📊 CACHÉ STATS - Hits: ${memoryCache.stats.hits}, Misses: ${memoryCache.stats.misses}, Ahorradas: ${memoryCache.stats.saved}`);
    console.log(`📊 Tamaño caché: ${memoryCache.data.size} entradas`);
  }
};

module.exports = {
  getDataWithCache,
  invalidateCache,
  memoryCache
};