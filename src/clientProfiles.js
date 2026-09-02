const CLIENTES = {
  casa_nicho: { nombre: 'Casa Nicho', emitidas: 'johan', recibidas: 'saraguato' },
  corpes: { nombre: 'Corpes', emitidas: 'johan', recibidas: 'johan' },
  amanda_sanchez: { nombre: 'Amanda Sanchez', emitidas: 'saraguato', recibidas: 'saraguato' },
  gesell: { nombre: 'Gesell', emitidas: 'johan', recibidas: 'johan' },
  inmobiliaria_saraguato: { nombre: 'Inmobiliaria Saraguato', emitidas: 'saraguato', recibidas: 'saraguato' },
  julia_angelica: { nombre: 'Julia Angelica', emitidas: 'johan', recibidas: 'johan' },
  leonardo: { nombre: 'Leonardo (Leo)', emitidas: 'johan', recibidas: 'johan' },
  johan: { nombre: 'Johan', emitidas: 'johan', recibidas: 'johan' },
  jose_francisco: { nombre: 'Jose Francisco', emitidas: 'johan', recibidas: 'johan' },
  ruben: { nombre: 'Ruben', emitidas: 'johan', recibidas: 'johan' },
  minerva: { nombre: 'Minerva', emitidas: 'johan', recibidas: 'johan' },
  teresa: { nombre: 'Teresa', emitidas: 'johan', recibidas: 'johan' },
  jose_carlos: { nombre: 'Jose Carlos', emitidas: 'johan', recibidas: 'johan' },
  alvaro: { nombre: 'Alvaro', emitidas: 'johan', recibidas: 'johan' },
  luis_morales: { nombre: 'Luis Morales', emitidas: 'johan', recibidas: 'johan' },
  laura_kristell: { nombre: 'Laura Kristell', emitidas: 'laura', recibidas: 'johan' },
  maria_kairene: { nombre: 'Maria kairene', emitidas: 'johan', recibidas: 'johan' },
  daniel: { nombre: 'Daniel (Dani)', emitidas: 'dani', recibidas: 'dani' }
};

const EMITIDAS = {
  ingresos: ['INGRESOS', 3], exento: ['EXENTO', 4], iva16: ['IVA 16%', 5],
  subtotal: ['SUBTOTAL', fila => numero(fila[3]) + numero(fila[4])],
  subtotalLaura: ['SUBTOTAL', fila => numero(fila[3]) + numero(fila[5])],
  retIsr: ['RET. ISR', (_, extra) => numero(extra?.retencionIsr)],
  retIva: ['RET. IVA', (_, extra) => numero(extra?.retencionIva)], total: ['TOTAL', 6]
};
const RECIBIDAS = {
  egresos: ['EGRESOS', (_, extra) => numero(extra?.base8) + numero(extra?.base16) + numero(extra?.tasa0)],
  iva: ['IVA', (_, extra) => numero(extra?.iva8) + numero(extra?.iva16)],
  gravado8: ['GRAVADO 8%', (_, extra) => numero(extra?.base8)],
  gravado16: ['GRAVADO 16%', (_, extra) => numero(extra?.base16)],
  tasa0: ['TASA 0%', (_, extra) => numero(extra?.tasa0)], exento: ['EXENTO', 5],
  subtotal: ['SUBTOTAL', 6], iva8: ['IVA 8%', 7], iva16: ['IVA 16%', 8], ieps: ['IEPS', 9],
  retIsr: ['RET. ISR', (_, extra) => numero(extra?.retencionIsr)],
  retIva: ['RET. IVA', (_, extra) => numero(extra?.retencionIva)], total: ['TOTAL', 10]
};

function numero(valor) { return Math.round((Number(valor) || 0) * 100) / 100; }
function sumar(filas, inicio, fin) {
  const totales = [];
  for (let i = inicio; i <= fin; i += 1) totales[i] = numero(filas.reduce((s, f) => s + numero(f[i]), 0));
  return totales;
}
function tablaSeleccionada(tabla, mapa, seleccion, extras = []) {
  const columnas = seleccion.filter(clave => mapa[clave]);
  const encabezados = ['FECHA', 'FACTURA', tabla.encabezados[2], ...columnas.map(clave => mapa[clave][0]), 'TIPO / METODO', 'RFC TERCERO', 'UUID', 'ESTATUS'];
  const filas = tabla.filas.map((fila, index) => {
    const extra = extras[index] || {};
    const valores = columnas.map(clave => {
      const origen = mapa[clave][1];
      return typeof origen === 'function' ? origen(fila, extra) : numero(fila[origen]);
    });
    return [fila[0], fila[1], fila[2], ...valores, ...fila.slice(-4)];
  });
  return { encabezados, filas, registros: tabla.registros, totales: sumar(filas, 3, 3 + columnas.length - 1) };
}

function aplicarPerfilCliente(resultado, clienteId, personalizado = null) {
  const cliente = CLIENTES[clienteId] || CLIENTES.johan;
  const extrasE = (resultado.emitidas.registros || []).map(r => r.detalleFiscal || {});
  const extrasR = (resultado.recibidas.registros || []).map(r => r.hotelRecibida || {});
  resultado.clienteSeleccionado = { id: clienteId, nombre: cliente.nombre };
  resultado.perfilCedula = { emitidas: cliente.emitidas, recibidas: cliente.recibidas };
  if (personalizado) {
    resultado.tipoCedula = 'personalizada';
    resultado.emitidas = tablaSeleccionada(resultado.emitidas, EMITIDAS, personalizado.emitidas || [], extrasE);
    resultado.recibidas = tablaSeleccionada(resultado.recibidas, RECIBIDAS, personalizado.recibidas || [], extrasR);
    return resultado;
  }
  if (cliente.emitidas === 'laura') {
    resultado.emitidas = tablaSeleccionada(resultado.emitidas, EMITIDAS, ['ingresos', 'iva16', 'subtotalLaura', 'retIsr', 'retIva', 'total'], extrasE);
    resultado.emitidas.encabezados[3] = 'IMPORTE';
  }
  if (cliente.emitidas === 'dani') {
    resultado.emitidas = tablaSeleccionada(resultado.emitidas, EMITIDAS, ['ingresos', 'exento', 'retIsr', 'iva16', 'total'], extrasE);
    resultado.emitidas.encabezados[6] = 'IVA';
  }
  if (cliente.recibidas === 'dani') resultado.recibidas = tablaSeleccionada(resultado.recibidas, RECIBIDAS, ['exento', 'egresos', 'iva', 'total'], extrasR);
  if (cliente.emitidas === 'johan' && cliente.recibidas === 'saraguato') {
    const hotel = resultado.hoteleria?.recibidas || [];
    const encabezados = ['FECHA', 'FACTURA', 'TERCERO / CONCEPTO', 'CASO', 'GRAVADO 8%', 'GRAVADO 16%', 'TASA 0%', 'EXENTO', 'IVA 8%', 'IVA 16%', 'IEPS', 'SUBTOTAL', 'RET. ISR', 'RET. IVA', 'TOTAL', 'TIPO / METODO', 'RFC TERCERO', 'UUID', 'ESTATUS'];
    const filas = hotel.map(item => [item.fecha, item.factura, item.nombre, item.caso, item.base8, item.base16, item.tasa0, item.exento, item.iva8, item.iva16, item.ieps, item.subtotal, item.retencionIsr, item.retencionIva, item.pagoNeto, item.tipoMetodo, item.rfc, item.uuid, item.estatus]);
    resultado.recibidas = { encabezados, filas, totales: sumar(filas, 4, 14), registros: resultado.recibidas.registros };
  }
  resultado.tipoCedula = cliente.emitidas === 'saraguato' && cliente.recibidas === 'saraguato' ? 'hoteleria' : (cliente.emitidas === 'johan' && cliente.recibidas === 'johan' ? 'general' : 'cliente');
  return resultado;
}

module.exports = { CLIENTES, aplicarPerfilCliente };
