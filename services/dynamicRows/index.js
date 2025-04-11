/**
 * Dynamic Rows Services
 * Exports generators for dynamic rows in reports
 */

const expensesGenerator = require('./expensesGenerator');
const risksGenerator = require('./risksGenerator');

module.exports = {
  generateExpenseRows: expensesGenerator.generateRows,
  generateRiskRows: risksGenerator.generateRows
};
