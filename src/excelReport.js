async function crearExcel(reporte) {
  const html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: Arial, sans-serif; }
    h1, h2, h3 { text-align: center; margin: 4px 0; }
    table { border-collapse: collapse; margin-top: 14px; width: 100%; }
    th, td { border: 1px solid #000; padding: 5px; font-size: 11px; }
    th { font-weight: bold; text-align: center; }
    .emitidas th { background: #D9EAF7; }
    .recibidas th { background: #FCE4D6; }
    .number { mso-number-format: "\\$\\#\\,\\#\\#0\\.00"; text-align: right; }
    .total td { background: #FFF2CC; font-weight: bold; }
    .PPD td { background: #FFF2CC; }
    .COMPLEMENTO td { background: #FFE699; }
    .CANCELADA td { background: #EA9999; }
    .REVISAR td { background: #F4CCCC; }
  </style>
</head>
<body>
  <h1>${escapeHtml(reporte.cliente?.nombre || 'NOMBRE DEL CLIENTE')}</h1>
  <h2>RFC: ${escapeHtml(reporte.cliente?.rfc || 'XAXX010101000')}</h2>
  <h3>REPORTE FINANCIERO - ${escapeHtml(reporte.periodo || '')} - VERSION WEB</h3>
  ${diagnosticoHtml(reporte.diagnostico || {})}
  ${tablaHtml('FACTURAS EMITIDAS', reporte.emitidas, 'emitidas', 3, 6)}
  ${tablaHtml('FACTURAS RECIBIDAS', reporte.recibidas, 'recibidas', 3, 9)}
</body>
</html>`;
  return Buffer.from(html, 'utf8');
}

function diagnosticoHtml(diagnostico) {
  return `<table>
    <tbody>
      <tr><td><b>XML emitidas encontrados</b></td><td>${Number(diagnostico.xmlEmitidas || 0)}</td><td><b>CFDI emitidas antes del filtro</b></td><td>${Number(diagnostico.emitidasAntesFiltro || 0)}</td><td><b>Fechas emitidas</b></td><td>${escapeHtml(diagnostico.fechasEmitidas || '')}</td></tr>
      <tr><td><b>XML recibidas encontrados</b></td><td>${Number(diagnostico.xmlRecibidas || 0)}</td><td><b>CFDI recibidas antes del filtro</b></td><td>${Number(diagnostico.recibidasAntesFiltro || 0)}</td><td><b>Fechas recibidas</b></td><td>${escapeHtml(diagnostico.fechasRecibidas || '')}</td></tr>
    </tbody>
  </table>`;
}

function tablaHtml(titulo, tabla, clase, moneyStart, moneyEnd) {
  const filas = tabla.filas?.length
    ? tabla.filas.map((fila) => {
      const rowClass = String(fila[fila.length - 1] === 'REVISAR' || fila[fila.length - 1] === 'CANCELADA' ? fila[fila.length - 1] : fila[clase === 'emitidas' ? 7 : 10]).toUpperCase();
      return `<tr class="${escapeHtml(rowClass)}">${fila.map((valor, index) => celda(valor, index, moneyStart, moneyEnd)).join('')}</tr>`;
    }).join('')
    : `<tr><td colspan="${tabla.encabezados.length}">Sin CFDI aplicables</td></tr>`;
  const totales = tabla.filas?.length
    ? `<tr class="total">${tabla.encabezados.map((_, index) => {
      if (index === 2) return '<td>TOTAL</td>';
      if (index >= moneyStart && index <= moneyEnd) return `<td class="number">${Number(tabla.totales?.[index] || 0)}</td>`;
      return '<td></td>';
    }).join('')}</tr>`
    : '';

  return `<table class="${clase}">
    <caption><b>${titulo}</b></caption>
    <thead><tr>${tabla.encabezados.map((header) => `<th>${escapeHtml(header)}</th>`).join('')}</tr></thead>
    <tbody>${filas}${totales}</tbody>
  </table>`;
}

function celda(valor, index, moneyStart, moneyEnd) {
  const clase = index >= moneyStart && index <= moneyEnd ? ' class="number"' : '';
  return `<td${clase}>${escapeHtml(valor)}</td>`;
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, (char) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  })[char]);
}

module.exports = { crearExcel };
