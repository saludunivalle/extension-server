const { google } = require('googleapis');
const { jwtClient } = require('../config/google');
const excelUtils = require('../utils/excelUtils');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const os = require('os');

/**
 * Servicio para manejar operaciones con Google Drive
 */

class DriveService {
  constructor() {
    this.drive = google.drive({ version: 'v3', auth: jwtClient });
    this.templateFolderId = '12bxb0XEArXMLvc7gX2ndqJVqS_sTiiUE'; //Folder con plantillas
    this.uploadsFolder = '1iDJTcUYCV7C7dTsa0Y3rfBAjFUUelu-x';
    
    // IDs de plantillas de formularios
    this.templateIds = {
      1: '1xsz9YSnYEOng56eNKGV9it9EgTn0mZw1', //Plantilla de formulario 1
      2: '1JY-4IfJqEWLqZ_wrq_B_bfIlI9MeVzgF', //Plantilla de formulario 2
      3: '1WoPUZYusNl2u3FpmZ1qiO5URBUqHIwKF', //Plantilla de formulario 3
      4: '1FTC7Vq3O4ultexRPXYrJKOpL9G0071-0', //Plantilla de formulario 4
    };
  }

  /**
   * Sube un archivo a Google Drive
   * @param {Object} file - Objeto de archivo de Multer
   * @param {String} folderDestination - ID de carpeta destino (opcional)
   * @returns {Promise<String>} - URL del archivo subido
   */

  async uploadFile(file, folderDestination = this.uploadsFolder) {
    try {
      if (!file) {
        throw new Error('No se proporcionó ningún archivo para subir');
      }

      const fileMetadata = {
        name: file.originalname,
        parents: [folderDestination]
      };
      
      const media = {
        mimeType: file.mimetype,
        body: fs.createReadStream(file.path)
      };
      
      const uploadedFile = await this.drive.files.create({
        requestBody: fileMetadata,
        media: media,
        fields: 'id'
      });
      
      const fileId = uploadedFile.data.id;
      
      // Hacer el archivo accesible públicamente
      await this.makeFilePublic(fileId);
      
      return `https://drive.google.com/file/d/${fileId}/view`;
    } catch (error) {
      console.error('Error al subir archivo a Google Drive:', error);
      throw new Error(`Error al subir archivo: ${error.message}`);
    }
  }

  /**
   * Hace público un archivo en Google Drive
   * @param {String} fileId - ID del archivo
   */

  async makeFilePublic(fileId) {
    try {
      await this.drive.permissions.create({
        fileId,
        requestBody: {
          role: 'reader',
          type: 'anyone'
        }
      });
      
      return true;
    } catch (error) {
      console.error('Error al hacer público el archivo:', error);
      throw new Error(`Error al modificar permisos: ${error.message}`);
    }
  }

  /**
   * Procesa un archivo XLSX sustituyendo marcadores con datos
   * @param {String} templateId - ID de la plantilla
   * @param {Object} data - Datos para reemplazar marcadores
   * @param {String} fileName - Nombre del archivo resultante
   * @returns {Promise<String>} - URL del archivo resultante
   */

  async processXLSXWithStyles(templateId, data, fileName) {
    try {
      console.log(`Descargando la plantilla: ${templateId}`);
      const fileResponse = await this.drive.files.get(
        { fileId: templateId, alt: 'media' },
        { responseType: 'stream' }
      );
  
      // Cargar el libro desde el stream
      const workbook = await excelUtils.loadWorkbookFromStream(fileResponse.data);
      console.log('Libro cargado desde stream correctamente');
  
      // Reemplazar marcadores
      excelUtils.replaceMarkers(workbook, data);
      console.log('Marcadores reemplazados correctamente');
  
      // Guardar a archivo temporal
      const tempFilePath = await excelUtils.saveToTempFile(workbook, fileName);
      console.log(`Archivo guardado temporalmente en ${tempFilePath}`);
  
      // Subir a Google Drive
      const uploadResponse = await this.drive.files.create({
        requestBody: {
          name: fileName,
          parents: [this.templateFolderId],
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        },
        media: {
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          body: fs.createReadStream(tempFilePath),
        },
      });
  
      const fileId = uploadResponse.data.id;
      await this.makeFilePublic(fileId);
  
      // Limpiar archivos temporales
      excelUtils.cleanupTempFiles([tempFilePath]);
  
      return `https://drive.google.com/file/d/${fileId}/view`;
    } catch (error) {
      console.error('Error al procesar archivo XLSX:', error);
      throw new Error(`Error al procesar archivo XLSX: ${error.message}`);
    }
  }

  /**
   * Genera un reporte basado en plantilla y datos
   * @param {Number} formNumber - Número de formulario (1-4)
   * @param {String} solicitudId - ID de la solicitud
   * @param {Object} data - Datos para rellenar la plantilla
   * @param {String} mode - Modo de visualización (view/edit)
   * @returns {Promise<String>} - URL del reporte generado
   */

  async generateReport(formNumber, solicitudId, data, mode = 'view') {
    try {
      
      const templateId = this.templateIds[formNumber];
      
      if (!templateId) {
        throw new Error(`Plantilla no encontrada para formulario ${formNumber}`);
      }
      
      const fileName = `Formulario${formNumber}_${solicitudId}`;
      const reportLink = await this.processXLSXWithStyles(templateId, data, fileName);
      
      // Si se solicita el modo "edit", modificamos el enlace
      if (mode === 'edit') {
        return reportLink.replace('/view', '/view?usp=drivesdk&edit=true');
      }
      
      return reportLink;
    } catch (error) {
      console.error('Error al generar reporte:', error);
      throw new Error(`Error al generar reporte: ${error.message}`);
    }
  }

  /**
   * Elimina un archivo de Google Drive
   * @param {String} fileId - ID del archivo a eliminar
   */

  async deleteFile(fileId) {
    try {
      await this.drive.files.delete({ fileId });
      return true;
    } catch (error) {
      console.error('Error al eliminar archivo:', error);
      throw new Error(`Error al eliminar archivo: ${error.message}`);
    }
  }
}

module.exports = new DriveService();