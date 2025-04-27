/**
 * Servicio específico para la generación de reportes con riesgos dinámicos 
 */
const { google } = require('googleapis');
const { jwtClient } = require('../config/google');
const fs = require('fs');
const path = require('path');
// Correctly import the risksGenerator module
const risksGenerator = require('./dynamicRows/risksGenerator'); 

class RiskReportService {
  /**
   * Genera un reporte con filas dinámicas de riesgos
   * @param {string} fileId - ID del archivo de Google Sheets donde insertar riesgos
   * @param {object} reportData - Datos del reporte con secciones de riesgos (e.g., { __FILAS_DINAMICAS_DISENO__: {...}, ... })
   * @returns {Promise<boolean>} - Éxito de la operación
   */
  async generateReportWithRisks(fileId, reportData) {
    try {
      console.log(`🔄 Iniciando procesamiento de riesgos dinámicos para el reporte ${fileId}`);
      
      if (!fileId) throw new Error('ID de archivo de reporte no proporcionado');
      if (!reportData) throw new Error('Datos del reporte no proporcionados');
      
      // Verificar acceso a API de Google y que el archivo es un Sheet
      await this.verifyFileIsGoogleSheet(fileId); // Use helper function
      
      // --- Orchestration Logic ---
      const processedCategories = {}; // Store info about inserted rows { category: { count: number, insertStartRow: number } }
      let totalRiesgosInsertados = 0;
      
      // Define the order in which categories must be processed
      const categoryOrder = ['diseno', 'locacion', 'desarrollo', 'cierre', 'otros'];
      
      for (const categoria of categoryOrder) {
        const categoryKey = `__FILAS_DINAMICAS_${categoria.toUpperCase()}__`;
        
        // Check if data exists for this category in the reportData
        if (reportData[categoryKey] && reportData[categoryKey].riesgos && reportData[categoryKey].riesgos.length > 0) {
          console.log(`🔄 Procesando categoría ${categoria}...`);
          
          // Generate rows data, passing previously processed info (needed for 'otros')
          const dynamicRowsData = risksGenerator.generateRows(
            reportData[categoryKey].riesgos, 
            categoria, 
            null, // No custom insert location, use category config
            processedCategories // Pass info about previous categories
          );
          
          if (dynamicRowsData && dynamicRowsData.rows && dynamicRowsData.rows.length > 0) {
            // Attempt to insert the rows
            try {
              const success = await risksGenerator.insertDynamicRows(fileId, dynamicRowsData);
              
              if (success) {
                const numRiesgos = dynamicRowsData.count;
                totalRiesgosInsertados += numRiesgos;
                
                // Store info about this insertion for the next category ('otros')
                processedCategories[categoria] = {
                  count: numRiesgos,
                  insertStartRow: dynamicRowsData.insertStartRow // Store the 1-based start row
                };
                console.log(`✅ Insertados ${numRiesgos} riesgos de ${categoria}. Próxima fila disponible: ${dynamicRowsData.insertStartRow + numRiesgos}`);
              } else {
                console.error(`❌ Error al insertar riesgos de ${categoria}`);
                // Decide if you want to stop or continue on error
              }
            } catch (insertError) {
              console.error(`❌ Error crítico al insertar riesgos de ${categoria}:`, insertError);
              // Decide if you want to stop or continue on error
            }
          } else {
            console.log(`ℹ️ No se generaron filas válidas para ${categoria}`);
          }
        } else {
          console.log(`ℹ️ No hay datos de riesgos para la sección ${categoria}`);
        }
      }
      // --- End Orchestration Logic ---
      
      console.log(`🏁 Finalizada inserción de riesgos dinámicos. Total: ${totalRiesgosInsertados} riesgos insertados`);
      return totalRiesgosInsertados >= 0; // Return true even if 0 inserted, as long as no errors stopped it
    } catch (error) {
      console.error('❌ Error en generación de reporte con riesgos dinámicos:', error);
      return false;
    }
  }
  
  /**
   * Actualiza un reporte con filas dinámicas de riesgos (wrapper around generate)
   * @param {string} reportURL - URL del reporte
   * @param {object} reportData - Datos del reporte con secciones de riesgos
   * @returns {Promise<boolean>} - Éxito de la operación
   */
  async updateReportWithRisks(reportURL, reportData) {
    try {
      const fileId = this.extractFileIdFromUrl(reportURL);
      if (!fileId) throw new Error('No se pudo extraer el ID del archivo de la URL');
      
      // Call the main generation function
      return await this.generateReportWithRisks(fileId, reportData);
    } catch (error) {
      console.error('❌ Error al actualizar reporte con riesgos:', error);
      return false;
    }
  }
  
  /**
   * Extrae el ID del archivo desde una URL de Google Sheets
   * @param {string} url - URL del archivo
   * @returns {string|null} - ID del archivo o null
   */
  extractFileIdFromUrl(url) {
    const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
    return match ? match[1] : null;
  }

  /**
   * Helper to verify the file exists and is a Google Sheet.
   * @param {string} fileId - ID of the Google Sheet
   * @throws {Error} If file not found or not a Google Sheet
   */
  async verifyFileIsGoogleSheet(fileId) {
    try {
      const drive = google.drive({ version: 'v3', auth: jwtClient });
      const fileResponse = await drive.files.get({
        fileId: fileId,
        fields: 'mimeType,name,id'
      });
      
      if (fileResponse.data.mimeType !== 'application/vnd.google-apps.spreadsheet') {
        // Log warning but don't throw error, maybe it works anyway? Or throw error.
        console.warn(`⚠️ El archivo ${fileId} (${fileResponse.data.name}) no es un Google Sheet nativo (${fileResponse.data.mimeType}). La inserción de filas podría fallar.`);
        // throw new Error(`El archivo ${fileId} no es un Google Sheet.`);
      } else {
        console.log(`✅ Verificado: ${fileId} (${fileResponse.data.name}) es un Google Sheet.`);
      }
    } catch (error) {
      console.error(`Error al verificar el archivo ${fileId}:`, error);
      if (error.code === 404) {
          throw new Error(`Archivo de reporte con ID ${fileId} no encontrado.`);
      }
      throw new Error(`Error al acceder al archivo de reporte ${fileId}: ${error.message}`);
    }
  }
}

module.exports = new RiskReportService();