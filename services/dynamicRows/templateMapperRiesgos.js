/**
 * Template mapper for riesgos (risks)
 * Defines how risk data should be mapped to the report template
 */

const fs = require('fs');
const path = require('path');

// Try to load template configuration
let templateConfig;
try {
  const templateFile = path.join(__dirname, '../../templates/risks.json');
  templateConfig = JSON.parse(fs.readFileSync(templateFile, 'utf8'));
} catch (error) {
  console.error('Warning: Could not load risks template configuration in mapper');
  templateConfig = null; // Fallback handled below
}

const templateMapperRiesgos = {
  // Default location (fallback if config fails or category doesn't specify)
  // Actualizado para reflejar el nuevo rango B:E
  defaultInsertLocation: templateConfig?.templateRow?.range || 'B43:E43',

  // Column mappings (zero-based index from the start column B)
  columns: {
    // Use columnConfig helper to read from risks.json or use defaults
    descripcion: columnConfig('descripcion', 'B', 0), // Column B -> index 0
    aplica: columnConfig('aplica', 'C', 1),           // Column C -> index 1
    mitigacion: columnConfig('mitigacion', 'D', 2)    // Column D -> index 2 (span handled by merge)
  },

  // Number of columns to span in the template (B to E = 4 columns)
  columnSpan: 4, // <-- Actualizado a 4

  // Function to format data before insertion
  formatData: function(riesgo) {
    // Use getDataValue to handle potential missing fields and mapping
    return {
      // Prioritize specific fields, provide fallbacks
      descripcion: getDataValue(riesgo, 'descripcion') || getDataValue(riesgo, 'nombre_riesgo') || '',
      // Ensure 'aplica' is 'Sí' or 'No' (or similar expected values)
      aplica: getDataValue(riesgo, 'aplica') === 'Sí' ? 'Sí' : 'No',
      mitigacion: getDataValue(riesgo, 'mitigacion') || getDataValue(riesgo, 'estrategia') || '' // Añadido fallback a 'estrategia'
    };
  },

  // Function to create a row with the correct structure (B, C, D, E)
  createRow: function(riesgo) {
    const formattedData = this.formatData(riesgo);

    // Create a row with 4 empty cells (for columns B, C, D, E)
    const row = new Array(this.columnSpan).fill(''); // <-- Usa this.columnSpan

    // Fill specific cells with data based on index
    row[this.columns.descripcion.index] = formattedData.descripcion; // Index 0 -> Col B
    row[this.columns.aplica.index] = formattedData.aplica;           // Index 1 -> Col C
    row[this.columns.mitigacion.index] = formattedData.mitigacion;     // Index 2 -> Col D (API pone valor aquí)

    return row;
  },

  // Get the template configuration
  getTemplateConfig: function() {
    return templateConfig;
  }
};

// --- Helper Functions ---

/**
 * Helper function to get column configuration
 * @param {String} key - Column key in the config ('descripcion', 'aplica', 'mitigacion')
 * @param {String} defaultColumn - Default column letter (B, C, D)
 * @param {Number} defaultIndex - Default column index (0, 1, 2)
 * @returns {Object} Column configuration object
 */
function columnConfig(key, defaultColumn, defaultIndex) {
  // Check if config loaded and has the specific column definition
  if (templateConfig && templateConfig.columns && templateConfig.columns[key]) {
    const configCol = templateConfig.columns[key].column || defaultColumn;
    // Determine the starting column letter (e.g., 'D' from 'D:E')
    const startColLetter = configCol.split(':')[0];
    // Calculate index relative to the start of the *template range* (which should be B)
    const templateStartCol = templateConfig?.templateRow?.range?.match(/([A-Z]+)/)?.[1] || 'B';

    return {
      column: configCol, // Keep original definition like "D:E"
      index: columnToIndex(startColLetter) - columnToIndex(templateStartCol), // Index relative to start column B
      span: templateConfig.columns[key].span || 1
    };
  }

  // Fallback if config is missing
  return {
    column: defaultColumn,
    index: defaultIndex,
    span: 1
  };
}

/**
 * Convert column letter to index (0-based)
 * @param {String} column - Column letter (e.g., "A", "BC")
 * @returns {Number|null} Column index (0-based) or null if invalid
 */
function columnToIndex(column) {
  if (!column || typeof column !== 'string') return null;
  // Handle simple case first
  const colLetter = column.split(':')[0].toUpperCase(); // Use only the start column if range (e.g., F from F:V)

  let result = 0;
  for (let i = 0; i < colLetter.length; i++) {
    const charCode = colLetter.charCodeAt(i);
    if (charCode < 65 || charCode > 90) return null; // Ensure it's A-Z
    result = result * 26 + (charCode - 64);
  }
  return result - 1; // Convert to 0-based index
}

/**
 * Get data value using mapping if available
 * @param {Object} data - Data object
 * @param {String} key - Key to get from data ('descripcion', 'aplica', 'mitigacion')
 * @param {String} defaultValue - Default value if not found
 * @returns {String} Data value
 */
function getDataValue(data, key, defaultValue = '') {
  let mappedKey = key;

  // 1. Try using data mapping from template config first
  if (templateConfig && templateConfig.dataMapping && templateConfig.dataMapping[key]) {
    mappedKey = templateConfig.dataMapping[key];
    if (data[mappedKey] !== undefined && data[mappedKey] !== null) {
      return data[mappedKey]?.toString() || defaultValue;
    }
  }

  // 2. Try direct access using the original key if mapping failed or didn't exist
  if (data[key] !== undefined && data[key] !== null) {
    return data[key]?.toString() || defaultValue;
  }

  // 3. (Optional but good) Try common alternative keys as a fallback
  const alternatives = {
    descripcion: ['nombre_riesgo', 'description'],
    aplica: ['aplica'], // Already specific
    mitigacion: ['mitigacion', 'estrategia', 'strategy']
  };

  if (alternatives[key]) {
    for (const alt of alternatives[key]) {
      if (data[alt] !== undefined && data[alt] !== null) {
        return data[alt]?.toString() || defaultValue;
      }
    }
  }

  // Return default if nothing found
  return defaultValue;
}

module.exports = templateMapperRiesgos;
