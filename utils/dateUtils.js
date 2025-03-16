/**
 * Utilidades para manipulación y formato de fechas
 * Centraliza las operaciones con fechas para garantizar consistencia en toda la aplicación
 */

/**
 * Formatea una fecha en formato día/mes/año
 * @param {Date|string|number} date - Fecha en cualquier formato reconocible por Date()
 * @param {Object} options - Opciones de formato
 * @param {boolean} [options.leadingZeros=true] - Si debe incluir ceros iniciales (ej: 01/02/2023)
 * @param {string} [options.separator='/'] - Separador entre día, mes y año
 * @returns {string} - Fecha formateada como dd/mm/yyyy
 */
const formatDate = (date, options = {}) => {
    const { leadingZeros = true, separator = '/' } = options;
    const dateObj = new Date(date);
    
    if (isNaN(dateObj.getTime())) {
      return '';
    }
    
    const day = leadingZeros 
      ? dateObj.getDate().toString().padStart(2, '0') 
      : dateObj.getDate();
    
    const month = leadingZeros 
      ? (dateObj.getMonth() + 1).toString().padStart(2, '0') 
      : (dateObj.getMonth() + 1);
    
    const year = dateObj.getFullYear();
    
    return `${day}${separator}${month}${separator}${year}`;
  };
  
  /**
   * Obtiene la fecha actual formateada
   * @param {Object} options - Opciones de formato (ver formatDate)
   * @returns {string} - Fecha actual formateada
   */
  const getCurrentDate = (options = {}) => {
    return formatDate(new Date(), options);
  };
  
  /**
   * Devuelve un objeto con día, mes y año separados
   * @param {Date|string|number} date - Fecha a formatear
   * @param {boolean} [withLeadingZeros=true] - Si incluye ceros iniciales
   * @returns {Object} - { dia, mes, anio }
   */
  const getDateParts = (date, withLeadingZeros = true) => {
    const dateObj = new Date(date);
    
    if (isNaN(dateObj.getTime())) {
      return { dia: '', mes: '', anio: '' };
    }
    
    return {
      dia: withLeadingZeros 
        ? dateObj.getDate().toString().padStart(2, '0') 
        : dateObj.getDate().toString(),
      mes: withLeadingZeros 
        ? (dateObj.getMonth() + 1).toString().padStart(2, '0') 
        : (dateObj.getMonth() + 1).toString(),
      anio: dateObj.getFullYear().toString()
    };
  };
  
  /**
   * Convierte una fecha en formato español (dd/mm/yyyy) a objeto Date
   * @param {string} dateStr - Fecha en formato dd/mm/yyyy
   * @returns {Date} - Objeto Date
   */
  const parseSpanishDate = (dateStr) => {
    if (!dateStr || typeof dateStr !== 'string') return null;
    
    // Soporta tanto / como - como separadores
    const parts = dateStr.split(/[/-]/);
    if (parts.length !== 3) return null;
    
    // En formato español: día/mes/año
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1; // Meses en JS: 0-11
    const year = parseInt(parts[2], 10);
    
    const date = new Date(year, month, day);
    
    // Verificar que la fecha sea válida
    if (
      isNaN(date.getTime()) ||
      date.getDate() !== day ||
      date.getMonth() !== month ||
      date.getFullYear() !== year
    ) {
      return null;
    }
    
    return date;
  };
  
  /**
   * Verifica si una fecha es válida
   * @param {Date|string|number} date - Fecha a validar
   * @returns {boolean} - true si es válida
   */
  const isValidDate = (date) => {
    const d = new Date(date);
    return !isNaN(d.getTime());
  };
  
  /**
   * Compara dos fechas
   * @param {Date|string|number} date1 - Primera fecha
   * @param {Date|string|number} date2 - Segunda fecha
   * @returns {number} - -1 si date1 < date2, 0 si iguales, 1 si date1 > date2
   */
  const compareDates = (date1, date2) => {
    const d1 = new Date(date1).getTime();
    const d2 = new Date(date2).getTime();
    
    if (isNaN(d1) || isNaN(d2)) return NaN;
    
    if (d1 < d2) return -1;
    if (d1 > d2) return 1;
    return 0;
  };
  
  /**
   * Formatea una fecha para Google Sheets (yyyy-mm-dd)
   * @param {Date|string|number} date - Fecha a formatear
   * @returns {string} - Fecha en formato yyyy-mm-dd
   */
  const formatDateForSheets = (date) => {
    const d = new Date(date);
    if (isNaN(d.getTime())) return '';
    
    return d.toISOString().split('T')[0];
  };
  
  module.exports = {
    formatDate,
    getCurrentDate,
    getDateParts,
    parseSpanishDate,
    isValidDate,
    compareDates,
    formatDateForSheets
  };