/**
 * Template mapper for gastos (expenses)
 * Defines how expense data should be mapped to the report template
 */

const fs = require('fs');
const path = require('path');

// Try to load template configuration
let templateConfig;
try {
  const templateFile = path.join(__dirname, '../../templates/expenses.json');
  templateConfig = JSON.parse(fs.readFileSync(templateFile, 'utf8'));
} catch (error) {
  console.error('Warning: Could not load expenses template configuration in mapper');
  templateConfig = null;
}

const templateMapperGastos = {
  // Default location where rows should be inserted
  defaultInsertLocation: templateConfig?.templateRow?.range || 'A44:AK44',
  
  // Column mappings (zero-based index from the start column)
  columns: {
    id: columnConfig('id', 'E', 0),               // Column E - ID of expense concept (e.g. "8.1")
    descripcion: columnConfig('description', 'F', 1),  // Column F - Description text
    cantidad: columnConfig('quantity', 'X', 23),    // Column X - Quantity
    valorUnit: columnConfig('unitValue', 'Z', 25),  // Column Z - Unit value
    valorTotal: columnConfig('totalValue', 'AC', 28) // Column AC - Total value
  },
  
  // Number of columns to span in the template (from A to AK)
  columnSpan: 37,
  
  // Function to format data before insertion
  formatData: function(gasto) {
    return {
      id: getDataValue(gasto, 'id'),
      descripcion: getDataValue(gasto, 'descripcion'),
      cantidad: getDataValue(gasto, 'cantidad', '0'),
      valorUnit: getDataValue(gasto, 'valorUnit'),
      valorTotal: getDataValue(gasto, 'valorTotal')
    };
  },
  
  // Function to create a row with the correct structure
  createRow: function(gasto) {
    const formattedData = this.formatData(gasto);
    
    // Create a row with empty cells spanning the full width
    const row = new Array(this.columnSpan).fill('');
    
    // Fill specific cells with data
    row[this.columns.id.index] = formattedData.id;
    row[this.columns.descripcion.index] = formattedData.descripcion;
    row[this.columns.cantidad.index] = formattedData.cantidad;
    row[this.columns.valorUnit.index] = formattedData.valorUnit;
    row[this.columns.valorTotal.index] = formattedData.valorTotal;
    
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
    id: ['id_concepto', 'id'],
    descripcion: ['descripcion', 'concepto', 'description'],
    cantidad: ['cantidad', 'quantity'],
    valorUnit: ['valorUnit_formatted', 'valor_unit_formatted', 'valorUnit', 'valor_unit'],
    valorTotal: ['valorTotal_formatted', 'valor_total_formatted', 'valorTotal', 'valor_total']
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

module.exports = templateMapperGastos;
