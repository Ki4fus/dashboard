import ExcelJS from 'exceljs';
import type {Data} from './data';
function scalar(value: any): any {
  if (value && typeof value === 'object' && !(value instanceof Date)) {
    if ('formula' in value || 'sharedFormula' in value) return scalar(value.result);
    if ('richText' in value) return value.richText.map((t: {text: string}) => t.text).join('');
  }
  return value;
}
function numericCell(cell: ExcelJS.Cell) {
  const value = cell.type === ExcelJS.ValueType.Formula ? scalar(cell.model.result) : scalar(cell.value);
  const valid = typeof value === 'number' && Number.isFinite(value);
  return {value: valid ? value : null, issue: valid ? null : value && typeof value === 'object' && 'error' in value ? `Ошибка Excel: ${value.error}` : 'Нет сохранённого числового итога'};
}
export async function readDashboard(buffer: ArrayBuffer, name: string): Promise<Data> {
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(buffer);
  const sheet = book.getWorksheet('RUB Deals Rates');
  if (!sheet) throw new Error('Не найден лист «RUB Deals Rates».');
  const points: Data['points'] = [];
  const operations: Data['operations'] = [];
  let dates = new Set<string>();
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
    const u = numericCell(row.getCell('U'));
    const dateIssue = dates.size !== 1 ? 'Не удалось однозначно определить дату блока' : null;
    points.push({date: dates.size === 1 ? [...dates][0] : null, total: y.value, fixedUsdAmount: u.value, row: row.number,
      issue: dateIssue || y.issue, fixedIssue: dateIssue || u.issue});
    dates = new Set<string>();
  });
  if (!points.length) throw new Error('Строки TOTAL не найдены.');
  return {sheet: sheet.name, points, operations, name, updated: new Date().toISOString()};
}

