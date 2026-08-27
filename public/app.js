const form = {
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
form.emitidas.addEventListener('change', actualizarConteos);
form.recibidas.addEventListener('change', actualizarConteos);
addQuickFilesBtn.addEventListener('click', agregarCargaRapida);
processBtn.addEventListener('click', procesar);
clearBtn.addEventListener('click', limpiar);
exportBtn.addEventListener('click', exportar);

actualizarCamposPeriodo();
actualizarConteos();

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
    a.download = `cedula-cfdi-${ultimoReporte.periodo.toLowerCase().replaceAll(' ', '-')}.xls`;
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

  tablesEl.innerHTML = [
    renderTabla('FACTURAS EMITIDAS', reporte.emitidas, 'emitidas'),
    renderTabla('FACTURAS RECIBIDAS', reporte.recibidas, 'recibidas')
  ].join('');
}

function renderTabla(titulo, tabla, clase) {
  const moneyStart = clase === 'emitidas' ? 3 : 3;
  const moneyEnd = clase === 'emitidas' ? 6 : 9;
  const rows = tabla.filas.length
    ? tabla.filas.map((fila) => {
      const rowClass = String(fila[fila.length - 1] === 'REVISAR' || fila[fila.length - 1] === 'CANCELADA' ? fila[fila.length - 1] : fila[clase === 'emitidas' ? 7 : 10]).toUpperCase();
      return `<tr class="${escapeHtml(rowClass)}">${fila.map((valor, index) => `<td class="${index >= moneyStart && index <= moneyEnd ? 'number' : ''}">${formatValue(valor, index >= moneyStart && index <= moneyEnd)}</td>`).join('')}</tr>`;
    }).join('')
    : `<tr><td colspan="${tabla.encabezados.length}">Sin CFDI aplicables para el periodo seleccionado.</td></tr>`;

  const totals = tabla.filas.length
    ? `<tfoot><tr>${tabla.encabezados.map((_, index) => {
      if (index === 2) return '<td>TOTAL</td>';
      if (index >= moneyStart && index <= moneyEnd) return `<td class="number">${formatMoney(tabla.totales[index] || 0)}</td>`;
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
