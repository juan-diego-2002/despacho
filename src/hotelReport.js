const ExcelJS = require('exceljs');

const money = '#,##0.00;[Red]-#,##0.00';
const border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
const stateColors = { PPD: 'FCDB57', COMPLEMENTO: '9FB0FF', NC: 'F8CBAD', CANCELADA: 'FC3030', REVISAR: 'F4CCCC' };
const monthNames = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];

function stateKey(item) {
  const status = String(item.estatus || '').toUpperCase();
  const method = String(item.tipoMetodo || '').toUpperCase();
  if (status === 'CANCELADA') return 'CANCELADA';
  if (method === 'NC') return 'NC';
  if (method === 'COMPLEMENTO') return 'COMPLEMENTO';
  if (method === 'PPD') return 'PPD';
  if (status === 'REVISAR') return 'REVISAR';
  return '';
}

function colorStateRow(row, item) {
  const color = stateColors[stateKey(item)];
  if (!color) return;
  row.eachCell(cell => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${color}` } }; });
}

function title(sheet, row, text, endColumn) {
  sheet.mergeCells(row, 1, row, endColumn);
  const cell = sheet.getCell(row, 1);
  cell.value = text;
  cell.font = { bold: true, size: row === 1 ? 15 : 12 };
  cell.alignment = { horizontal: 'center' };
}

function header(row) {
  row.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1D66D0' } };
  row.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  row.eachCell(cell => { cell.border = border; });
}

function periodFromDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return { year: value.getFullYear(), month: value.getMonth() };
  const text = String(value || '').trim();
  let match = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
  if (match) return { year: Number(match[3]), month: Number(match[2]) - 1 };
  match = text.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/);
  if (match) return { year: Number(match[1]), month: Number(match[2]) - 1 };
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : { year: parsed.getFullYear(), month: parsed.getMonth() };
}

function groupByPeriod(hotel) {
  const groups = new Map();
  const add = (kind, item) => {
    const period = periodFromDate(item.fecha);
    if (!period || period.month < 0 || period.month > 11) return;
    const key = `${period.year}-${String(period.month + 1).padStart(2, '0')}`;
    if (!groups.has(key)) groups.set(key, { ...period, emitidas: [], recibidas: [] });
    groups.get(key)[kind].push(item);
  };
  for (const item of hotel.emitidas || []) add('emitidas', item);
  for (const item of hotel.recibidas || []) add('recibidas', item);
  return Array.from(groups.values()).sort((a, b) => (a.year - b.year) || (a.month - b.month));
}

function numericTotal(items, field) {
  return items.reduce((sum, item) => sum + (Number(item[field]) || 0), 0);
}

function createSummarySheet(workbook, reporte, periods) {
  const sheet = workbook.addWorksheet('RESUMEN');
  title(sheet, 1, reporte.cliente?.nombre || 'NOMBRE DEL CLIENTE', 6);
  title(sheet, 2, `RFC: ${reporte.cliente?.rfc || ''}`, 6);
  title(sheet, 3, `REPORTE RESUMEN - ${reporte.periodo || ''}`, 6);
  sheet.getRow(5).values = ['MES', 'CFDI EMITIDAS', 'CFDI RECIBIDAS', 'TOTAL EMITIDAS', 'TOTAL RECIBIDAS', 'TOTAL GENERAL'];
  header(sheet.getRow(5));

  let row = 6;
  let emittedCount = 0;
  let receivedCount = 0;
  let emittedTotal = 0;
  let receivedTotal = 0;
  for (const period of periods) {
    const monthEmitted = numericTotal(period.emitidas, 'total');
    const monthReceived = numericTotal(period.recibidas, 'pagoNeto');
    emittedCount += period.emitidas.length;
    receivedCount += period.recibidas.length;
    emittedTotal += monthEmitted;
    receivedTotal += monthReceived;
    sheet.getRow(row).values = [`${monthNames[period.month]} ${period.year}`, period.emitidas.length, period.recibidas.length, monthEmitted, monthReceived, monthEmitted + monthReceived];
    row++;
  }
  sheet.getRow(row).values = ['TOTAL GENERAL', emittedCount, receivedCount, emittedTotal, receivedTotal, emittedTotal + receivedTotal];
  sheet.getRow(row).font = { bold: true };
  sheet.getRow(row).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
  for (let r = 5; r <= row; r++) sheet.getRow(r).eachCell(cell => { cell.border = border; });
  for (let col = 4; col <= 6; col++) sheet.getColumn(col).numFmt = money;
  sheet.columns.forEach(column => { column.width = 21; });
  sheet.views = [{ state: 'frozen', ySplit: 5 }];
  sheet.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 1 };
}

function createPeriodSheet(workbook, reporte, period, tasa) {
  const periodName = `${monthNames[period.month]} ${period.year}`;
  const sheet = workbook.addWorksheet(periodName.substring(0, 31));
  title(sheet, 1, reporte.cliente?.nombre || 'NOMBRE DEL CLIENTE', 19);
  title(sheet, 2, `RFC: ${reporte.cliente?.rfc || ''}`, 19);
  title(sheet, 4, `CÉDULA DE INGRESOS - ${periodName} | IVA ${tasa * 100}%`, 19);
  let row = 6;
  sheet.getRow(row).values = ['FECHA', 'FACTURA', 'CONCEPTO', 'HOSPEDAJE', 'ALIMENTOS', 'OTROS', 'SUBTOTAL', '2% S/HOSP', 'IVA', 'TOTAL', 'TOTAL CFDI', 'UUID', 'ESTATUS'];
  header(sheet.getRow(row)); row++;
  const incomeStart = row;
  for (const item of period.emitidas) {
    const current = row;
    sheet.getRow(row).values = [item.fecha, item.factura, item.nombre, item.hospedaje, item.alimentos, item.otros];
    sheet.getCell(current, 7).value = { formula: `ROUND(SUM(D${current}:F${current}),2)`, result: item.subtotal };
    sheet.getCell(current, 8).value = { formula: `ROUND(D${current}*2%,2)`, result: item.impuestoHospedaje };
    sheet.getCell(current, 9).value = { formula: `ROUND(G${current}*${tasa},2)`, result: item.iva };
    sheet.getCell(current, 10).value = { formula: `ROUND(SUM(G${current}:I${current}),2)`, result: item.total };
    sheet.getCell(current, 11).value = item.totalReferencia;
    sheet.getCell(current, 12).value = item.uuid;
    sheet.getCell(current, 13).value = item.estatus;
    colorStateRow(sheet.getRow(row), item);
    row++;
  }
  sheet.getCell(row, 3).value = 'TOTAL';
  for (let col = 4; col <= 11; col++) {
    sheet.getCell(row, col).value = row > incomeStart
      ? { formula: `SUM(${sheet.getColumn(col).letter}${incomeStart}:${sheet.getColumn(col).letter}${row - 1})` }
      : 0;
  }
  sheet.getRow(row).font = { bold: true }; row += 3;

  title(sheet, row, `CÉDULA DE EGRESOS - ${periodName}`, 19); row += 2;
  sheet.getRow(row).values = ['FECHA', 'FACTURA', 'NOMBRE', 'CASO', 'GRAVADO 8%', 'GRAVADO 16%', 'TASA 0%', 'EXENTO', 'IVA 8%', 'IVA 16%', 'IEPS', 'SUBTOTAL', 'RET. ISR', 'RET. IVA', 'TOTAL', 'TIPO / MÉTODO', 'RFC TERCERO', 'UUID', 'ESTATUS'];
  header(sheet.getRow(row)); row++;
  const expenseStart = row;
  for (const item of period.recibidas) {
    sheet.getRow(row).values = [item.fecha, item.factura, item.nombre, item.caso, item.base8, item.base16, item.tasa0, item.exento, item.iva8, item.iva16, item.ieps, item.subtotal, item.retencionIsr, item.retencionIva, item.pagoNeto, item.tipoMetodo, item.rfc, item.uuid, item.estatus];
    colorStateRow(sheet.getRow(row), item);
    row++;
  }
  sheet.getCell(row, 3).value = 'TOTALES';
  for (let col = 5; col <= 15; col++) {
    sheet.getCell(row, col).value = row > expenseStart
      ? { formula: `SUM(${sheet.getColumn(col).letter}${expenseStart}:${sheet.getColumn(col).letter}${row - 1})` }
      : 0;
  }
  sheet.getRow(row).font = { bold: true };

  for (let r = 1; r <= row; r++) sheet.getRow(r).eachCell(cell => { cell.border = border; });
  for (let col = 4; col <= 11; col++) sheet.getColumn(col).numFmt = money;
  for (let col = 5; col <= 15; col++) sheet.getColumn(col).numFmt = money;
  sheet.columns.forEach((column, index) => { column.width = index === 2 ? 28 : index === 17 ? 38 : 16; });
  sheet.views = [];
  sheet.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
}

async function crearExcelHoteleria(reporte) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Cédulas CFDI - Hotelería';
  workbook.created = new Date();
  const hotel = reporte.hoteleria || {};
  const tasa = Number(hotel.tasaIva) || 0.08;
  let periods = groupByPeriod(hotel);
  if (periods.length === 0) periods = [{ year: new Date().getFullYear(), month: 0, emitidas: hotel.emitidas || [], recibidas: hotel.recibidas || [] }];
  createSummarySheet(workbook, reporte, periods);
  for (const period of periods) createPeriodSheet(workbook, reporte, period, tasa);
  return workbook.xlsx.writeBuffer();
}

module.exports = { crearExcelHoteleria };
