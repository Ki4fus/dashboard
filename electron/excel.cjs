const ExcelJS = require('exceljs');
function scalar(value) {
  if (value && typeof value === 'object' && !(value instanceof Date)) {
    if ('formula' in value || 'sharedFormula' in value) return scalar(value.result);
    if ('richText' in value) return value.richText.map(t => t.text).join('');
  }
  return value;
}
function numericCell(cell) {
  const value = cell.type === ExcelJS.ValueType.Formula ? scalar(cell.model.result) : scalar(cell.value);
  const valid = typeof value === 'number' && Number.isFinite(value);
  return {value: valid ? value : null, issue: valid ? null : value && typeof value === 'object' && 'error' in value ? `Excel error: ${value.error}` : 'No saved numeric total'};
}
async function readDashboard(file) {
  const book = new ExcelJS.Workbook();
  await book.xlsx.readFile(file);
  const sheet = book.getWorksheet('RUB Deals Rates');
  if (!sheet) throw new Error('Sheet RUB Deals Rates was not found.');
  const points = [];
  const operations = [];
  let dates = new Set();
  sheet.eachRow(row => {
    const a = scalar(row.getCell('A').value);
    if (a instanceof Date && Number.isFinite(a.getTime())) {
      const date = a.toISOString().slice(0, 10);
      dates.add(date);
      const currency = scalar(row.getCell('L').value);
      const rate = numericCell(row.getCell('E')).value;
      const amount = numericCell(row.getCell('B')).value;
      const fixedUsdAmount = numericCell(row.getCell('U')).value;
      if (rate !== null || (amount !== null && amount !== 0) || (fixedUsdAmount !== null && fixedUsdAmount !== 0)) operations.push({date, row: row.number, currency: typeof currency === 'string' ? currency.trim().toUpperCase() : '', rate, fixedUsdAmount, hasAmount: amount !== null && amount !== 0});
    }
    if (typeof a !== 'string' || a.trim().toUpperCase() !== 'TOTAL') return;
    const y = numericCell(row.getCell('Y'));
    const t = numericCell(row.getCell('T'));
    const u = numericCell(row.getCell('U'));
    const dateIssue = dates.size !== 1 ? 'The block date is missing or ambiguous' : null;
    points.push({date: dates.size === 1 ? [...dates][0] : null, total: y.value, fxPnl: t.value, fixedUsdAmount: u.value, row: row.number,
      issue: dateIssue || y.issue, fxIssue: dateIssue || t.issue, fixedIssue: dateIssue || u.issue});
    dates = new Set();
  });
  if (!points.length) throw new Error('No TOTAL rows found.');
  return {sheet: sheet.name, points, operations};
}
module.exports = {readDashboard};
