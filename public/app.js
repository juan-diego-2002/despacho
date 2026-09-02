const form = {
  modoCedula: document.querySelector('#modoCedula'),
  cliente: document.querySelector('#cliente'),
  ivaHotel: document.querySelector('#ivaHotel'),
  modo: document.querySelector('#modo'),
  anio: document.querySelector('#anio'),
  mes: document.querySelector('#mes'),
  trimestre: document.querySelector('#trimestre'),
  emitidas: document.querySelector('#emitidas'),
  recibidas: document.querySelector('#recibidas'),
  tipoCarga: document.querySelector('#tipoCarga'),
  quickFiles: document.querySelector('#quickFiles')
};

const processBtn = document.querySelector('#processBtn');
const clearBtn = document.querySelector('#clearBtn');
const exportBtn = document.querySelector('#exportBtn');
const addQuickFilesBtn = document.querySelector('#addQuickFiles');
const statusEl = document.querySelector('#status');
const summaryEl = document.querySelector('#summary');
const tablesEl = document.querySelector('#tables');
let ultimoReporte = null;
const API_BASE = window.location.protocol === 'file:' ? 'http://localhost:3000' : '';

const hoy = new Date();
form.modo.value = 'auto';
form.anio.value = hoy.getFullYear();
form.mes.value = hoy.getMonth() + 1;

form.modo.addEventListener('change', actualizarCamposPeriodo);
form.cliente.addEventListener('change', actualizarTipoCedula);
document.querySelectorAll('.mode-tab').forEach(tab => tab.addEventListener('click', cambiarModoCedula));
form.emitidas.addEventListener('change', actualizarConteos);
form.recibidas.addEventListener('change', actualizarConteos);
addQuickFilesBtn.addEventListener('click', agregarCargaRapida);
processBtn.addEventListener('click', procesar);
clearBtn.addEventListener('click', limpiar);
exportBtn.addEventListener('click', exportar);

actualizarCamposPeriodo();
actualizarTipoCedula();
actualizarConteos();

function actualizarTipoCedula() {
  const usaHotel = ['casa_nicho', 'amanda_sanchez', 'inmobiliaria_saraguato'].includes(form.cliente.value) || form.modoCedula.value === 'personalizada';
  document.querySelectorAll('.hotel-field').forEach((field) => {
    field.classList.toggle('hidden', !usaHotel);
  });
}

function cambiarModoCedula(event) {
  const modo = event.currentTarget.dataset.mode;
  form.modoCedula.value = modo;
  document.querySelectorAll('.mode-tab').forEach(tab => tab.classList.toggle('active', tab.dataset.mode === modo));
  document.querySelector('.client-field').classList.toggle('hidden', modo === 'personalizada');
  document.querySelector('#customColumns').classList.toggle('hidden', modo !== 'personalizada');
  actualizarTipoCedula();
}

function actualizarCamposPeriodo() {
  document.querySelectorAll('.period-field').forEach((field) => {
    field.classList.toggle('hidden', field.dataset.for !== form.modo.value);
  });
}

function actualizarConteos() {
  document.querySelector('#emitidasCount').textContent = `${form.emitidas.files.length} archivos`;
  document.querySelector('#recibidasCount').textContent = `${form.recibidas.files.length} archivos`;
}

function agregarCargaRapida() {
  if (!form.quickFiles.files.length) {
    statusEl.textContent = 'Selecciona XML o ZIP en la carga rapida.';
    return;
  }
  const destino = form.tipoCarga.value === 'recibidas' ? form.recibidas : form.emitidas;
  const dataTransfer = new DataTransfer();
  [...destino.files, ...form.quickFiles.files].forEach((file) => dataTransfer.items.add(file));
  destino.files = dataTransfer.files;
  form.quickFiles.value = '';
  actualizarConteos();
  statusEl.textContent = 'Archivos agregados al tipo seleccionado.';
}

async function procesar() {
  if (!form.emitidas.files.length && !form.recibidas.files.length) {
    statusEl.textContent = 'Selecciona al menos un XML o ZIP.';
    return;
  }

  const data = new FormData();
  [...form.emitidas.files].forEach((file) => data.append('emitidas', file));
  [...form.recibidas.files].forEach((file) => data.append('recibidas', file));
  data.append('modo', form.modo.value);
  data.append('anio', form.anio.value);
  data.append('mes', form.mes.value);
  data.append('trimestre', form.trimestre.value);
  data.append('modoCedula', form.modoCedula.value);
  data.append('cliente', form.cliente.value);
  data.append('columnasEmitidas', JSON.stringify([...document.querySelectorAll('[name="colEmitidas"]:checked')].map(el => el.value)));
  data.append('columnasRecibidas', JSON.stringify([...document.querySelectorAll('[name="colRecibidas"]:checked')].map(el => el.value)));
  data.append('ivaHotel', form.ivaHotel.value);

  processBtn.disabled = true;
  exportBtn.disabled = true;
  statusEl.textContent = 'Procesando CFDI...';

  try {
    const response = await fetch(`${API_BASE}/api/procesar`, { method: 'POST', body: data });
    const json = await response.json();
    if (!response.ok) throw new Error(json.error || 'No se pudo procesar.');
    ultimoReporte = json;
    renderReporte(json);
    exportBtn.disabled = false;
    statusEl.textContent = 'Proceso terminado.';
  } catch (error) {
    statusEl.textContent = mensajeConexion(error);
  } finally {
    processBtn.disabled = false;
  }
}

async function exportar() {
  if (!ultimoReporte) return;
  exportBtn.disabled = true;
  statusEl.textContent = 'Generando Excel...';
  try {
    const response = await fetch(`${API_BASE}/api/exportar`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(ultimoReporte)
    });
    if (!response.ok) {
      const json = await response.json();
      throw new Error(json.error || 'No se pudo exportar.');
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `cedula-cfdi-${ultimoReporte.periodo.toLowerCase().replaceAll(' ', '-')}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
    statusEl.textContent = 'Excel generado.';
  } catch (error) {
    statusEl.textContent = mensajeConexion(error);
  } finally {
    exportBtn.disabled = false;
  }
}

function limpiar() {
  form.emitidas.value = '';
  form.recibidas.value = '';
  form.quickFiles.value = '';
  ultimoReporte = null;
  tablesEl.innerHTML = '';
  summaryEl.innerHTML = '';
  summaryEl.classList.add('hidden');
  exportBtn.disabled = true;
  statusEl.textContent = 'Listo para cargar XML o ZIP.';
  actualizarConteos();
}

function renderReporte(reporte) {
  summaryEl.classList.remove('hidden');
  summaryEl.innerHTML = `
    <div class="metric"><strong>${escapeHtml(reporte.emitidas.filas.length)}</strong><span>Emitidas integradas</span></div>
    <div class="metric"><strong>${escapeHtml(reporte.recibidas.filas.length)}</strong><span>Recibidas integradas</span></div>
    <div class="metric"><strong>${escapeHtml((reporte.diagnostico?.xmlEmitidas || 0) + (reporte.diagnostico?.xmlRecibidas || 0))}</strong><span>XML encontrados</span></div>
    <div class="metric"><strong>${escapeHtml((reporte.diagnostico?.emitidasFueraPeriodo || 0) + (reporte.diagnostico?.recibidasFueraPeriodo || 0))}</strong><span>Fuera del periodo</span></div>
    <div class="metric"><strong>${escapeHtml(reporte.periodo)}</strong><span>${escapeHtml(reporte.cliente.rfc || 'RFC no detectado')}</span></div>
  `;

  if (reporte.tipoCedula === 'hoteleria') {
    renderHoteleria(reporte);
    return;
  }

  tablesEl.innerHTML = [
    renderTabla('FACTURAS EMITIDAS', reporte.emitidas, 'emitidas'),
    renderTabla('FACTURAS RECIBIDAS', reporte.recibidas, 'recibidas')
  ].join('');
}

function renderHoteleria(reporte) {
  const hotel = reporte.hoteleria;
  tablesEl.innerHTML = `
    <section class="table-section">
      <div class="table-title"><h2>INGRESOS DE HOTELERÍA</h2><p>Clasificados automáticamente desde los XML</p></div>
      <div class="hotel-help">Los importes se colocan automáticamente en Hospedaje, Alimentos u Otros. El subtotal, 2% de hospedaje, IVA y total se calculan solos.</div>
      <div class="table-wrap"><table class="hotel-table">
        <thead><tr><th>FECHA</th><th>FACTURA</th><th>CLIENTE</th><th>HOSPEDAJE</th><th>ALIMENTOS</th><th>OTROS</th><th>SUBTOTAL</th><th>2% HOSP.</th><th>IVA ${hotel.tasaIva * 100}%</th><th>TOTAL</th><th>TOTAL CFDI</th></tr></thead>
        <tbody>${hotel.emitidas.map((item, index) => `<tr>
          <td>${escapeHtml(item.fecha)}</td><td>${escapeHtml(item.factura)}</td><td>${escapeHtml(item.nombre)}</td>
          ${hotelInput('emitidas', index, 'hospedaje', item.hospedaje)}${hotelInput('emitidas', index, 'alimentos', item.alimentos)}${hotelInput('emitidas', index, 'otros', item.otros)}
          <td class="number calc" data-out="emitidas-${index}-subtotal">${formatMoney(item.subtotal)}</td>
          <td class="number calc" data-out="emitidas-${index}-impuestoHospedaje">${formatMoney(item.impuestoHospedaje)}</td>
          <td class="number calc" data-out="emitidas-${index}-iva">${formatMoney(item.iva)}</td>
          <td class="number calc" data-out="emitidas-${index}-total">${formatMoney(item.total)}</td>
          <td class="number reference">${formatMoney(item.totalReferencia)}</td>
        </tr>`).join('')}</tbody>
      </table></div>
    </section>
    <section class="table-section">
      <div class="table-title"><h2>EGRESOS DE HOTELERÍA</h2><p>Selecciona el caso y captura únicamente las celdas amarillas necesarias</p></div>
      <div class="table-wrap"><table class="hotel-table hotel-expenses">
        <thead><tr><th>FECHA</th><th>FACTURA</th><th>PROVEEDOR</th><th>CASO</th><th>TOTAL</th><th>GRAVADO 8%</th><th>IVA 8%</th><th>GRAVADO 16%</th><th>IVA 16%</th><th>TASA 0%</th><th>EXENTO</th><th>IEPS</th><th>SUBTOTAL</th><th>RET. ISR</th><th>RET. IVA</th><th>TIPO / MÉTODO</th><th>RFC TERCERO</th><th>UUID</th><th>ESTATUS</th></tr></thead>
        <tbody>${hotel.recibidas.map((item, index) => renderRecibidaHotel(item, index)).join('')}</tbody>
      </table></div>
    </section>`;
  tablesEl.querySelectorAll('[data-hotel-input]').forEach(input => input.addEventListener('input', actualizarHotelDesdePantalla));
  tablesEl.querySelectorAll('[data-hotel-case]').forEach(input => input.addEventListener('change', actualizarHotelDesdePantalla));
}

function hotelInput(grupo, index, campo, valor) {
  return `<td class="editable"><input type="number" step="0.01" data-hotel-input data-group="${grupo}" data-index="${index}" data-field="${campo}" value="${Number(valor) || 0}"></td>`;
}

function renderRecibidaHotel(item, index) {
  const casos = [
    ['0', 'Complemento informativo'],
    ['1', '1 · Gravado 8%'], ['2', '2 · Gravado 16%'], ['3', '3 · Tasa 0% / retención'],
    ['4', '4 · Mixto 8% y 16%'], ['5', '5 · Exento con ISR 1.25%'], ['IEPS', 'IEPS'], ['REVISAR', 'Revisar clasificación']
  ];
  const rowClass = claseEstadoHotel(item);
  return `<tr class="${escapeHtml(rowClass)}"><td>${escapeHtml(item.fecha)}</td><td>${escapeHtml(item.factura)}</td><td>${escapeHtml(item.nombre)}</td>
    <td class="editable"><select data-hotel-case data-group="recibidas" data-index="${index}" data-field="caso">${casos.map(([value, label]) => `<option value="${value}" ${item.caso === value ? 'selected' : ''}>${label}</option>`).join('')}</select></td>
    ${hotelInput('recibidas', index, 'pagoNeto', item.pagoNeto)}
    ${hotelInput('recibidas', index, 'base8', item.base8)}${hotelInput('recibidas', index, 'iva8', item.iva8)}
    ${hotelInput('recibidas', index, 'base16', item.base16)}${hotelInput('recibidas', index, 'iva16', item.iva16)}
    ${hotelInput('recibidas', index, 'tasa0', item.tasa0)}${hotelInput('recibidas', index, 'exento', item.exento)}
    ${hotelInput('recibidas', index, 'ieps', item.ieps)}
    <td class="number calc" data-out="recibidas-${index}-subtotal">${formatMoney(item.subtotal)}</td>
    ${hotelInput('recibidas', index, 'retencionIsr', item.retencionIsr)}${hotelInput('recibidas', index, 'retencionIva', item.retencionIva)}
    <td>${escapeHtml(item.tipoMetodo)}</td><td>${escapeHtml(item.rfc)}</td><td>${escapeHtml(item.uuid)}</td><td>${escapeHtml(item.estatus)}</td></tr>`;
}

function claseEstadoHotel(item) {
  const estado = String(item.estatus || '').toUpperCase();
  const metodo = String(item.tipoMetodo || '').toUpperCase();
  if (estado === 'CANCELADA') return 'CANCELADA';
  if (metodo === 'NC') return 'NC';
  if (metodo === 'COMPLEMENTO') return 'COMPLEMENTO';
  if (metodo === 'PPD') return 'PPD';
  if (estado === 'REVISAR') return 'REVISAR';
  return '';
}

function actualizarHotelDesdePantalla(event) {
  const el = event.target;
  const grupo = el.dataset.group;
  const index = Number(el.dataset.index);
  const campo = el.dataset.field;
  const item = ultimoReporte.hoteleria[grupo][index];
  item[campo] = campo === 'caso' ? el.value : Number(el.value || 0);
  if (grupo === 'emitidas') calcularEmitidaHotel(item);
  else calcularRecibidaHotel(item);
  actualizarFilaHotel(grupo, index, item);
}

function actualizarFilaHotel(grupo, index, item) {
  Object.entries(item).forEach(([campo, valor]) => {
    const output = tablesEl.querySelector(`[data-out="${grupo}-${index}-${campo}"]`);
    if (output) {
      output.textContent = formatMoney(Number(valor) || 0);
      if (campo === 'diferencia') {
        output.classList.toggle('mismatch', Math.abs(Number(valor) || 0) > 0.02);
        output.classList.toggle('ok', Math.abs(Number(valor) || 0) <= 0.02);
      }
    }
    const input = tablesEl.querySelector(`[data-hotel-input][data-group="${grupo}"][data-index="${index}"][data-field="${campo}"]`);
    if (input && document.activeElement !== input) input.value = Number(valor) || 0;
  });
}

function calcularEmitidaHotel(item) {
  item.subtotal = moneyRound(item.hospedaje + item.alimentos + item.otros);
  item.impuestoHospedaje = moneyRound(item.hospedaje * 0.02);
  item.iva = moneyRound(item.subtotal * ultimoReporte.hoteleria.tasaIva);
  item.total = moneyRound(item.subtotal + item.impuestoHospedaje + item.iva);
}

function calcularRecibidaHotel(item) {
  const neto = Number(item.pagoNeto) || 0;
  if (item.caso === '0') {
    item.pagoNeto = item.base8 = item.iva8 = item.base16 = item.iva16 = 0;
    item.tasa0 = item.exento = item.retencionIsr = item.retencionIva = 0;
  } else if (item.caso === '1') {
    item.base8 = moneyRound(neto / 1.08); item.iva8 = moneyRound(item.base8 * 0.08);
    item.base16 = item.iva16 = item.tasa0 = item.exento = 0;
  } else if (item.caso === '2') {
    item.base16 = moneyRound(neto / 1.16); item.iva16 = moneyRound(item.base16 * 0.16);
    item.base8 = item.iva8 = item.tasa0 = item.exento = 0;
  } else if (item.caso === '4') {
    const restante = Math.max(neto - item.base8 - item.iva8, 0);
    item.base16 = moneyRound(restante / 1.16); item.iva16 = moneyRound(item.base16 * 0.16);
    item.tasa0 = item.exento = 0;
  }
  item.subtotal = moneyRound(item.base8 + item.base16 + item.tasa0 + item.exento);
  if (item.caso === '5') item.retencionIsr = moneyRound(item.subtotal * 0.0125);
  const sinIeps = moneyRound(item.subtotal + item.iva8 + item.iva16 - item.retencionIsr - item.retencionIva);
  const conIeps = moneyRound(sinIeps + (Number(item.ieps) || 0));
  const calculado = Math.abs(neto - conIeps) < Math.abs(neto - sinIeps) ? conIeps : sinIeps;
  item.netoCalculado = calculado;
  item.diferencia = moneyRound(neto - calculado);
}

function moneyRound(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}

function renderTabla(titulo, tabla, clase) {
  const noMonetarias = new Set(['FECHA', 'FACTURA', 'CONCEPTO', 'TERCERO / CONCEPTO', 'CASO', 'TIPO / METODO', 'RFC TERCERO', 'UUID', 'ESTATUS']);
  const esMonetaria = index => index >= 3 && !noMonetarias.has(String(tabla.encabezados[index] || '').toUpperCase());
  const rows = tabla.filas.length
    ? tabla.filas.map((fila) => {
      const rowClass = String(fila[fila.length - 1] === 'REVISAR' || fila[fila.length - 1] === 'CANCELADA' ? fila[fila.length - 1] : fila[clase === 'emitidas' ? 7 : 11]).toUpperCase();
      return `<tr class="${escapeHtml(rowClass)}">${fila.map((valor, index) => `<td class="${esMonetaria(index) ? 'number' : ''}">${formatValue(valor, esMonetaria(index))}</td>`).join('')}</tr>`;
    }).join('')
    : `<tr><td colspan="${tabla.encabezados.length}">Sin CFDI aplicables para el periodo seleccionado.</td></tr>`;

  const totals = tabla.filas.length
    ? `<tfoot><tr>${tabla.encabezados.map((_, index) => {
      if (index === 2) return '<td>TOTAL</td>';
      if (esMonetaria(index)) return `<td class="number">${formatMoney(tabla.totales[index] || 0)}</td>`;
      return '<td></td>';
    }).join('')}</tr></tfoot>`
    : '';

  return `
    <section class="table-section">
      <div class="table-title">
        <h2>${titulo}</h2>
        <p>${tabla.filas.length} registros</p>
      </div>
      <div class="table-wrap ${clase}">
        <table>
          <thead><tr>${tabla.encabezados.map((header) => `<th>${escapeHtml(header)}</th>`).join('')}</tr></thead>
          <tbody>${rows}</tbody>
          ${totals}
        </table>
      </div>
    </section>
  `;
}

function formatValue(value, money) {
  return money ? formatMoney(Number(value) || 0) : escapeHtml(value);
}

function formatMoney(value) {
  return value.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' });
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

function mensajeConexion(error) {
  if (String(error.message || '').toLowerCase().includes('fetch')) {
    return 'No se pudo conectar con Node. Abre http://localhost:3000 o ejecuta npm start.';
  }
  return error.message;
}
