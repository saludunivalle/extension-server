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
  templateConfig = null;
}

const templateMapperRiesgos = {
  // Default location where rows should be inserted
  defaultInsertLocation: templateConfig?.templateRow?.range || 'B44:H44',
  
  // Column mappings (zero-based index from the start column)
  columns: {
    id: columnConfig('id', 'B', 0),               // Column B - ID of risk
    descripcion: columnConfig('description', 'C', 1), // Column C - Description of risk
    impacto: columnConfig('impact', 'E', 3),        // Column E - Impact level
    probabilidad: columnConfig('probability', 'F', 4), // Column F - Probability
    estrategia: columnConfig('strategy', 'G', 5)     // Column G - Strategy to mitigate
  },
  
  // Number of columns to span in the template
  columnSpan: 7,        // From B to H (7 columns)
  
  // Function to format data before insertion
  formatData: function(riesgo) {
    return {
      id: getDataValue(riesgo, 'id'),
      descripcion: getDataValue(riesgo, 'descripcion'),
      impacto: getDataValue(riesgo, 'impacto'),
      probabilidad: getDataValue(riesgo, 'probabilidad'),
      estrategia: getDataValue(riesgo, 'estrategia')
    };
  },
  
  // Function to create a row with the correct structure
  createRow: function(riesgo) {
    const formattedData = this.formatData(riesgo);
    
    // Create a row with empty cells spanning the full width
    const row = new Array(this.columnSpan).fill('');
    
    // Fill specific cells with data
    row[this.columns.id.index] = formattedData.id;
    row[this.columns.descripcion.index] = formattedData.descripcion;
    row[this.columns.impacto.index] = formattedData.impacto;
    row[this.columns.probabilidad.index] = formattedData.probabilidad;
    row[this.columns.estrategia.index] = formattedData.estrategia;
    
    return row;
  },
  
  // Get the template configuration
  getTemplateConfig: function() {
    return templateConfig;
  }
};

/**
 * Helper function to get column configuration
 * @param {String} key - Column key in the config
 * @param {String} defaultColumn - Default column letter if config not available
 * @param {Number} defaultIndex - Default column index if config not available
 * @returns {Object} Column configuration object
 */
function columnConfig(key, defaultColumn, defaultIndex) {
  if (templateConfig && templateConfig.columns && templateConfig.columns[key]) {
    return {
      column: templateConfig.columns[key].column || defaultColumn,
      index: columnToIndex(templateConfig.columns[key].column) || defaultIndex,
      span: templateConfig.columns[key].span || 1
    };
  }
  
  return {
    column: defaultColumn,
    index: defaultIndex,
    span: 1
  };
}

/**
 * Convert column letter to index (0-based)
 * @param {String} column - Column letter (e.g., "A", "BC")
 * @returns {Number} Column index (0-based)
 */
function columnToIndex(column) {
  if (!column) return null;
  
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 64);
  }
  return result - 1; // Convert to 0-based index
}

/**
 * Get data value using mapping if available
 * @param {Object} data - Data object
 * @param {String} key - Key to get from data
 * @param {String} defaultValue - Default value if not found
 * @returns {String} Data value
 */
function getDataValue(data, key, defaultValue = '') {
  // Try direct access first
  if (data[key] !== undefined) {
    return data[key]?.toString() || defaultValue;
  }
  
  // Try using data mapping from template config
  if (templateConfig && templateConfig.dataMapping) {
    const mappedKey = templateConfig.dataMapping[key];
    if (mappedKey && data[mappedKey] !== undefined) {
      return data[mappedKey]?.toString() || defaultValue;
    }
  }
  
  // Try alternative keys based on common patterns
  const alternatives = {
    id: ['id', 'codigo'],
    descripcion: ['descripcion', 'description'],
    impacto: ['impacto', 'impact'],
    probabilidad: ['probabilidad', 'probability'],
    estrategia: ['estrategia', 'mitigacion', 'strategy']
  };
  
  if (alternatives[key]) {
    for (const alt of alternatives[key]) {
      if (data[alt] !== undefined) {
        return data[alt]?.toString() || defaultValue;
      }
    }
  }
  
  return defaultValue;
}

module.exports = templateMapperRiesgos;
