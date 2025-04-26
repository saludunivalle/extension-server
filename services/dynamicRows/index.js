const expensesGenerator = require('./expensesGenerator');
const risksGenerator = require('./risksGenerator');

/**
 * Genera filas dinámicas para gastos
 * @param {Array|Object} expenses - Datos de gastos
 * @param {String} insertLocation - Ubicación de inserción opcional
 * @returns {Object} Objeto con datos formateados para inserción
 */
const generateExpenseRows = (expenses, insertLocation = null) => {
  return expensesGenerator.generateRows(expenses, insertLocation);
};

/**
 * Genera filas dinámicas para riesgos
 * @param {Array|Object} risks - Datos de riesgos
 * @param {String} categoria - Categoría de riesgos
 * @param {String} insertLocation - Ubicación de inserción opcional
 * @returns {Object} Objeto con datos formateados para inserción
 */
const generateRiskRows = (risks, categoria = null, insertLocation = null) => {
  return risksGenerator.generateRows(risks, categoria, insertLocation);
};

module.exports = {
  generateExpenseRows,
  insertExpenseDynamicRows: expensesGenerator.insertDynamicRows,
  generateRiskRows,
  insertRiskDynamicRows: risksGenerator.insertDynamicRows
};