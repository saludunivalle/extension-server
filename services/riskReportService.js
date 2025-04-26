/**
 * Servicio específico para la generación de reportes con riesgos dinámicos 
 */
const { google } = require('googleapis');
const { jwtClient } = require('../config/google');
const fs = require('fs');
const path = require('path');
const risksGenerator = require('./dynamicRows/risksGenerator');

class RiskReportService {
  /**
   * Genera un reporte con filas dinámicas de riesgos
   * @param {string} fileId - ID del archivo de Google Sheets donde insertar riesgos
   * @param {object} reportData - Datos del reporte con secciones de riesgos
   * @returns {Promise<boolean>} - Éxito de la operación
   */
  async generateReportWithRisks(fileId, reportData) {
    try {
      console.log(`🔄 Iniciando procesamiento de riesgos dinámicos para el reporte ${fileId}`);
      
      if (!fileId) {
        throw new Error('ID de archivo de reporte no proporcionado');
      }
      
      // Verificar acceso a API de Google
      const sheets = google.sheets({version: 'v4', auth: jwtClient});
      
      // Verificar que el archivo existe
      try {
        await sheets.spreadsheets.get({
          spreadsheetId: fileId,
          fields: 'spreadsheetId'
        });
        console.log(`✅ Archivo de reporte verificado: ${fileId}`);
      } catch (error) {
        console.error(`❌ Error al verificar el archivo de reporte: ${error.message}`);
        return false;
      }
      
      // Procesar cada sección de riesgos dinámicos
      const seccionesRiesgos = [
        { key: '__FILAS_DINAMICAS_DISENO__', name: 'Diseño', insertarEn: 'B18:H18' },
        { key: '__FILAS_DINAMICAS_LOCACION__', name: 'Locación', insertarEn: 'B24:H24' },
        { key: '__FILAS_DINAMICAS_DESARROLLO__', name: 'Desarrollo', insertarEn: 'B35:H35' },
        { key: '__FILAS_DINAMICAS_CIERRE__', name: 'Cierre', insertarEn: 'B38:H38' },
        { key: '__FILAS_DINAMICAS_OTROS__', name: 'Otros', insertarEn: 'B41:H41' }
      ];
      
      let totalRiesgosInsertados = 0;
      
      for (const seccion of seccionesRiesgos) {
        if (reportData[seccion.key]) {
          console.log(`🔄 Procesando riesgos de ${seccion.name}`);
          
          const dynamicRowsData = reportData[seccion.key];
          
          // Asegurarse de que tengamos todo lo necesario
          if (!dynamicRowsData.insertarEn) {
            dynamicRowsData.insertarEn = seccion.insertarEn;
          }
          
          // Intentar insertar las filas dinámicas
          try {
            const success = await risksGenerator.insertDynamicRows(fileId, dynamicRowsData);
            
            if (success) {
              const numRiesgos = dynamicRowsData.riesgos?.length || 0;
              totalRiesgosInsertados += numRiesgos;
              console.log(`✅ Insertados ${numRiesgos} riesgos de ${seccion.name}`);
            } else {
              console.error(`❌ Error al insertar riesgos de ${seccion.name}`);
            }
          } catch (error) {
            console.error(`❌ Error al procesar riesgos de ${seccion.name}:`, error);
          }
        } else {
          console.log(`ℹ️ No hay riesgos para la sección ${seccion.name}`);
        }
      }
      
      console.log(`🏁 Finalizada inserción de riesgos dinámicos. Total: ${totalRiesgosInsertados} riesgos insertados`);
      return totalRiesgosInsertados > 0;
    } catch (error) {
      console.error('❌ Error en generación de reporte con riesgos dinámicos:', error);
      return false;
    }
  }
  
  /**
   * Actualiza un reporte con filas dinámicas de riesgos
   * @param {string} reportURL - URL del reporte
   * @param {object} reportData - Datos del reporte con secciones de riesgos
   * @returns {Promise<boolean>} - Éxito de la operación
   */
  async updateReportWithRisks(reportURL, reportData) {
    try {
      // Extraer el ID del reporte desde la URL
      const fileId = this.extractFileIdFromUrl(reportURL);
      
      if (!fileId) {
        throw new Error('No se pudo extraer el ID del archivo de la URL');
      }
      
      return await this.generateReportWithRisks(fileId, reportData);
    } catch (error) {
      console.error('❌ Error al actualizar reporte con riesgos:', error);
      return false;
    }
  }
  
  /**
   * Extrae el ID del archivo desde una URL de Google Sheets
   * @param {string} url - URL del archivo de Google Sheets
   * @returns {string|null} - ID del archivo o null si no se pudo extraer
   */
  extractFileIdFromUrl(url) {
    try {
      if (!url) return null;
      
      // Patrones para URLs de Google Sheets
      const patterns = [
        /\/d\/([a-zA-Z0-9-_]+)/, // Formato: .../d/ID/...
        /id=([a-zA-Z0-9-_]+)/, // Formato: ...?id=ID...
        /^([a-zA-Z0-9-_]+)$/ // Solo ID
      ];
      
      for (const pattern of patterns) {
        const match = url.match(pattern);
        if (match && match[1]) {
          return match[1];
        }
      }
      
      return null;
    } catch (error) {
      console.error('Error al extraer ID del archivo:', error);
      return null;
    }
  }
}

module.exports = new RiskReportService();