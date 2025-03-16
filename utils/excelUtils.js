const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const os = require('os');

/**
 * Utilidades para trabajar con archivos Excel
 * Proporciona métodos para crear, modificar y estilizar documentos Excel
 */

/**
 * Crea un libro de Excel desde una plantilla
 * @param {string} templatePath - Ruta al archivo de plantilla
 * @returns {Promise<ExcelJS.Workbook>} - Libro de Excel cargado
 */
const loadTemplateWorkbook = async (templatePath) => {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(templatePath);
      return workbook;
    } catch (error) {
      console.error('Error al cargar la plantilla Excel:', error);
      throw new Error(`Error al cargar plantilla Excel: ${error.message}`);
    }
  };
  
  /**
   * Crea un libro de Excel desde un stream de datos
   * @param {ReadableStream} stream - Stream de datos de Excel
   * @returns {Promise<ExcelJS.Workbook>} - Libro de Excel cargado
   */
  const loadWorkbookFromStream = async (stream) => {
    try {
      const workbook = new ExcelJS.Workbook();
      return await workbook.xlsx.read(stream);
    } catch (error) {
      console.error('Error al cargar Excel desde stream:', error);
      throw new Error(`Error al cargar Excel desde stream: ${error.message}`);
    }
  };
  
  /**
   * Crea un nuevo libro de Excel
   * @returns {ExcelJS.Workbook} - Nuevo libro de Excel
   */
  const createNewWorkbook = () => {
    return new ExcelJS.Workbook();
  };
  
  /**
   * Reemplaza marcadores en un libro de Excel con datos reales
   * @param {ExcelJS.Workbook} workbook - Libro de Excel a procesar
   * @param {Object} data - Objeto con datos para reemplazar marcadores
   * @param {Object} options - Opciones de configuración
   * @param {string} [options.markerFormat='{{%s}}'] - Formato de los marcadores (usa %s para la variable)
   * @returns {ExcelJS.Workbook} - Libro de Excel con marcadores reemplazados
   */
  const replaceMarkers = (workbook, data, options = {}) => {
    const { markerFormat = '{{%s}}' } = options;
    
    workbook.eachSheet((sheet) => {
      sheet.eachRow((row) => {
        row.eachCell((cell) => {
          if (typeof cell.value === 'string') {
            Object.entries(data).forEach(([key, value]) => {
              const marker = markerFormat.replace('%s', key);
              if (cell.value.includes(marker)) {
                const newValue = cell.value.replace(marker, value || '');
                cell.value = newValue;
              }
            });
          }
        });
      });
    });
    
    return workbook;
  };
  
  /**
   * Guarda un libro de Excel en un archivo temporal
   * @param {ExcelJS.Workbook} workbook - Libro de Excel a guardar
   * @param {string} fileName - Nombre del archivo (sin extensión)
   * @returns {Promise<string>} - Ruta del archivo guardado
   */
  const saveToTempFile = async (workbook, fileName) => {
    try {
      const tempDir = os.tmpdir();
      const filePath = path.join(tempDir, `${fileName}.xlsx`);
      
      await workbook.xlsx.writeFile(filePath);
      return filePath;
    } catch (error) {
      console.error('Error al guardar archivo Excel temporal:', error);
      throw new Error(`Error al guardar Excel temporal: ${error.message}`);
    }
  };
  
  /**
   * Aplica estilos básicos a un rango de celdas
   * @param {ExcelJS.Worksheet} sheet - Hoja de Excel
   * @param {string} range - Rango de celdas (ej: 'A1:C5')
   * @param {Object} styles - Estilos a aplicar
   */
  const applyStyleToRange = (sheet, range, styles) => {
    const rangeRegex = /([A-Z]+)(\d+):([A-Z]+)(\d+)/;
    const match = range.match(rangeRegex);
    
    if (!match) {
      throw new Error(`Formato de rango inválido: ${range}`);
    }
    
    const [, startCol, startRow, endCol, endRow] = match;
    const startColIndex = colNameToIndex(startCol);
    const endColIndex = colNameToIndex(endCol);
    
    for (let row = parseInt(startRow); row <= parseInt(endRow); row++) {
      for (let colIdx = startColIndex; colIdx <= endColIndex; colIdx++) {
        const col = indexToColName(colIdx);
        const cellAddress = `${col}${row}`;
        const cell = sheet.getCell(cellAddress);
        
        Object.assign(cell, styles);
      }
    }
  };
  
  /**
   * Convierte nombre de columna a índice (A=0, B=1, etc.)
   * @param {string} colName - Nombre de columna (A, B, AA, etc.)
   * @returns {number} - Índice de columna
   */
  const colNameToIndex = (colName) => {
    let index = 0;
    for (let i = 0; i < colName.length; i++) {
      index = index * 26 + colName.charCodeAt(i) - 64;
    }
    return index - 1; // 0-based index
  };
  
  /**
   * Convierte índice a nombre de columna (0=A, 1=B, etc.)
   * @param {number} index - Índice de columna
   * @returns {string} - Nombre de columna
   */
  const indexToColName = (index) => {
    let name = '';
    index += 1; // 1-based
    
    while (index > 0) {
      const modulo = (index - 1) % 26;
      name = String.fromCharCode(65 + modulo) + name;
      index = Math.floor((index - modulo) / 26);
    }
    
    return name;
  };
  
  /**
   * Crea una tabla formateada en Excel
   * @param {ExcelJS.Worksheet} sheet - Hoja de Excel
   * @param {Array<string>} headers - Cabeceras de la tabla
   * @param {Array<Array>} data - Datos para la tabla
   * @param {Object} options - Opciones de formato
   */
  const createFormattedTable = (sheet, headers, data, options = {}) => {
    const { 
      startRow = 1, 
      startColumn = 'A',
      headerStyle = {
        font: { bold: true, color: { argb: 'FFFFFFFF' } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } }
      },
      rowStyle = {
        font: { color: { argb: 'FF000000' } }
      },
      alternatingRowStyle = {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } }
      }
    } = options;
  
    // Insertar cabeceras
    const headerRow = sheet.getRow(startRow);
    headers.forEach((header, idx) => {
      const cell = headerRow.getCell(colNameToIndex(startColumn) + idx + 1);
      cell.value = header;
      Object.assign(cell, headerStyle);
    });
  
    // Insertar datos
    data.forEach((rowData, rowIdx) => {
      const excelRow = sheet.getRow(startRow + rowIdx + 1);
      
      rowData.forEach((cellData, colIdx) => {
        const cell = excelRow.getCell(colNameToIndex(startColumn) + colIdx + 1);
        cell.value = cellData;
        
        // Aplicar estilo base para todas las filas
        Object.assign(cell, rowStyle);
        
        // Aplicar estilo alternado para filas pares
        if (rowIdx % 2 === 1 && alternatingRowStyle) {
          Object.assign(cell, alternatingRowStyle);
        }
      });
    });
  
    // Auto-ajustar columnas
    const maxColumn = colNameToIndex(startColumn) + headers.length;
    for (let i = colNameToIndex(startColumn); i < maxColumn; i++) {
      const column = sheet.getColumn(i + 1);
      let maxLength = headers[i - colNameToIndex(startColumn)].length;
      
      // Obtener longitud máxima para ajustar columna
      data.forEach(row => {
        const cellValue = String(row[i - colNameToIndex(startColumn)] || '');
        if (cellValue && cellValue.length > maxLength) {
          maxLength = cellValue.length;
        }
      });
      
      column.width = maxLength + 2; // Agregar espacio extra
    }
  };
  
  /**
   * Limpia los archivos temporales
   * @param {Array<string>} filePaths - Rutas de archivos a eliminar
   */
  const cleanupTempFiles = (filePaths) => {
    filePaths.forEach(filePath => {
      try {
        if (fs.existsSync(filePath)) {
          fs.unlinkSync(filePath);
        }
      } catch (error) {
        console.warn(`Error al eliminar archivo temporal ${filePath}:`, error);
      }
    });
  };
  
  /**
   * Verifica si un directorio temporal existe, lo crea si no
   * @param {string} dirPath - Ruta del directorio
   * @returns {string} - Ruta del directorio
   */
  const ensureTempDir = (dirPath = path.join(os.tmpdir(), 'extension-server-temp')) => {
    if (!fs.existsSync(dirPath)) {
      fs.mkdirSync(dirPath, { recursive: true });
    }
    return dirPath;
  };
  
  /**
   * Convierte datos de una hoja plana a un objeto
   * @param {Array<Array>} rows - Filas de datos
   * @param {Array<string>} headers - Cabeceras (opcional, usa primera fila si no se proporciona)
   * @returns {Array<Object>} - Datos convertidos a objetos
   */
  const rowsToObjects = (rows, headers = null) => {
    if (!rows || !rows.length) return [];
    
    const headerRow = headers || rows[0];
    const dataRows = headers ? rows : rows.slice(1);
    
    return dataRows.map(row => {
      const obj = {};
      headerRow.forEach((header, idx) => {
        obj[header] = row[idx] || '';
      });
      return obj;
    });
  };
  
  module.exports = {
    loadTemplateWorkbook,
    loadWorkbookFromStream,
    createNewWorkbook,
    replaceMarkers,
    saveToTempFile,
    applyStyleToRange,
    colNameToIndex,
    indexToColName,
    createFormattedTable,
    cleanupTempFiles,
    ensureTempDir,
    rowsToObjects
  };