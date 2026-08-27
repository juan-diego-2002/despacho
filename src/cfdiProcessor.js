const JSZip = require('jszip');
const crypto = require('crypto');
const { XMLParser } = require('fast-xml-parser');

const CONFIG_CEDULAS = {
  TOLERANCIA: 0.005,
  EVITAR_PAGO_TOTAL_DUPLICADO: true,
  CFDI_CANCELADOS: []
};

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  removeNSPrefix: true,
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: true
});

async function procesarLoteCfdi({ emitidas, recibidas, filtro }) {
  const emitidasXml = await extraerXmls(emitidas || []);
  const recibidasXml = await extraerXmls(recibidas || []);
  const emitidasProcesadas = procesarGrupo(emitidasXml, 'emitida');
  const recibidasProcesadas = procesarGrupo(recibidasXml, 'recibida');

  const cliente = validarClienteUnico(emitidasProcesadas, recibidasProcesadas);
  const emitidasFiltradas = filtrarRegistros(emitidasProcesadas.registros, filtro);
  const recibidasFiltradas = filtrarRegistros(recibidasProcesadas.registros, filtro);
  const periodo = describirPeriodo(filtro);
  const advertencias = emitidasProcesadas.advertencias.concat(recibidasProcesadas.advertencias);
  agregarAdvertenciaPeriodo(advertencias, 'emitida', emitidasXml.length, emitidasProcesadas.registros, emitidasFiltradas, filtro);
  agregarAdvertenciaPeriodo(advertencias, 'recibida', recibidasXml.length, recibidasProcesadas.registros, recibidasFiltradas, filtro);

  return {
    cliente,
    periodo,
    filtro,
    diagnostico: {
      xmlEmitidas: emitidasXml.length,
      xmlRecibidas: recibidasXml.length,
      emitidasAntesFiltro: emitidasProcesadas.registros.length,
      recibidasAntesFiltro: recibidasProcesadas.registros.length,
      emitidasFueraPeriodo: Math.max(emitidasProcesadas.registros.length - emitidasFiltradas.length, 0),
      recibidasFueraPeriodo: Math.max(recibidasProcesadas.registros.length - recibidasFiltradas.length, 0),
      fechasEmitidas: describirFechasRegistros(emitidasProcesadas.registros),
      fechasRecibidas: describirFechasRegistros(recibidasProcesadas.registros)
    },
    emitidas: {
      encabezados: ['FECHA', 'FACTURA', 'CONCEPTO', 'INGRESOS', 'EXENTO', 'IVA 16%', 'TOTAL', 'TIPO / METODO', 'RFC TERCERO', 'UUID', 'ESTATUS'],
      filas: emitidasFiltradas.map((registro) => registro.fila),
      totales: sumarColumnas(emitidasFiltradas.map((registro) => registro.fila), 3, 6)
    },
    recibidas: {
      encabezados: ['FECHA', 'FACTURA', 'TERCERO / CONCEPTO', 'GTO 8%', 'GTO 16%', 'EXENTO', 'SUBTOTAL', 'IVA 8%', 'IVA 16%', 'TOTAL', 'TIPO / METODO', 'RFC TERCERO', 'UUID', 'ESTATUS'],
      filas: recibidasFiltradas.map((registro) => registro.fila),
      totales: sumarColumnas(recibidasFiltradas.map((registro) => registro.fila), 3, 9)
    },
    advertencias
  };
}

async function extraerXmls(files) {
  const xmls = [];
  for (const file of files) {
    const nombre = file.originalname || file.name || 'archivo';
    if (nombre.toLowerCase().endsWith('.xml')) {
      xmls.push({ contenido: file.buffer.toString('utf8'), origen: nombre });
    } else if (nombre.toLowerCase().endsWith('.zip')) {
      const zip = await JSZip.loadAsync(file.buffer);
      const entries = Object.keys(zip.files);
      for (const entryName of entries) {
        const entry = zip.files[entryName];
        if (!entry.dir && entryName.toLowerCase().endsWith('.xml')) {
          xmls.push({ contenido: await entry.async('string'), origen: `${nombre} > ${entryName}` });
        }
      }
    }
  }
  return xmls;
}

function procesarGrupo(xmlPendientes, tipo) {
  const registros = [];
  const advertencias = [];
  const uuidVistos = {};
  const clientesEncontrados = {};
  let clienteInfo = { nombre: '', rfc: '' };
  const ingresosPorUuid = {};

  for (const xml of xmlPendientes) {
    try {
      const resumen = obtenerResumenCfdi(xml.contenido);
      if (resumen.tipo === 'I' && resumen.uuid) ingresosPorUuid[resumen.uuid] = resumen;
    } catch (_) {}
  }

  for (const xml of xmlPendientes) {
    try {
      const registro = procesarXmlCfdi(xml.contenido, tipo, xml.origen, ingresosPorUuid);
      if (registro.uuid && uuidVistos[registro.uuid]) {
        advertencias.push(['DUPLICADO', tipo, registro.factura, xml.origen, 'UUID repetido; se integro una sola vez.']);
        continue;
      }
      if (registro.uuid) uuidVistos[registro.uuid] = true;
      if (!clienteInfo.nombre && registro.clienteInfo.nombre) clienteInfo = registro.clienteInfo;
      if (registro.clienteInfo.rfc) clientesEncontrados[registro.clienteInfo.rfc] = registro.clienteInfo.nombre || '';
      advertencias.push(...registro.advertencias);
      if (registro.incluir) registros.push(registro);
    } catch (error) {
      advertencias.push(['XML INVALIDO', tipo, '', xml.origen, error.message]);
    }
  }

  registros.sort((a, b) => a.fechaOrden - b.fechaOrden || String(a.factura).localeCompare(String(b.factura), undefined, { numeric: true }));
  return { registros, clienteInfo, clientesEncontrados, advertencias };
}

function procesarXmlCfdi(contenidoXml, tipo, origen, ingresosPorUuid) {
  const root = obtenerRoot(contenidoXml);
  const emisor = child(root, 'Emisor');
  const receptor = child(root, 'Receptor');
  const nombreEmisor = atributoTexto(emisor, 'Nombre', 'DESCONOCIDO');
  const nombreReceptor = atributoTexto(receptor, 'Nombre', 'DESCONOCIDO');
  const rfcEmisor = atributoTexto(emisor, 'Rfc', '');
  const rfcReceptor = atributoTexto(receptor, 'Rfc', '');
  const metodoPago = atributoTexto(root, 'MetodoPago', '');
  const estadoXml = atributoTexto(root, 'Estatus', atributoTexto(root, 'Estado', ''));
  const tipoComprobante = atributoTexto(root, 'TipoDeComprobante', '');
  const total = atributoNumero(root, 'Total');
  const subtotalCfdi = atributoNumero(root, 'SubTotal');
  const descuento = atributoNumero(root, 'Descuento');
  const fechaIso = atributoTexto(root, 'Fecha', '');
  const folio = atributoTexto(root, 'Folio', 'S/F');
  const serie = atributoTexto(root, 'Serie', '');
  const factura = serie ? `${serie} ${folio}` : folio;
  const uuid = obtenerUuid(root) || huellaXml(contenidoXml);
  const advertencias = [];
  const clienteInfo = {
    nombre: tipo === 'emitida' ? nombreEmisor : nombreReceptor,
    rfc: tipo === 'emitida' ? rfcEmisor : rfcReceptor
  };

  if (estaCancelado(factura, uuid, estadoXml)) {
    advertencias.push(['CANCELADA', tipo, factura, origen, 'CFDI marcado como cancelado.']);
    return crearRegistroCancelado(tipo, tipoComprobante, fechaIso, factura, uuid, clienteInfo, nombreEmisor, nombreReceptor, rfcEmisor, rfcReceptor, metodoPago, advertencias);
  }

  if (tipoComprobante === 'P') {
    return procesarComplementoPago(root, tipo, origen, ingresosPorUuid, uuid, factura, fechaIso, clienteInfo, nombreEmisor, nombreReceptor, advertencias);
  }

  if (tipoComprobante !== 'I') {
    advertencias.push([`EXCLUIDO TIPO ${tipoComprobante || 'VACIO'}`, tipo, factura, origen, 'La cedula integra CFDI I y complementos P; este tipo se excluye.']);
    return registroExcluido(uuid, factura, fechaIso, clienteInfo, advertencias);
  }

  if (total <= 0) {
    advertencias.push(['EXCLUIDO TOTAL CERO', tipo, factura, origen, 'El CFDI tipo I no tiene importe positivo.']);
    return registroExcluido(uuid, factura, fechaIso, clienteInfo, advertencias);
  }

  const iva = obtenerIvaCfdi(root);
  const conceptos = analizarConceptos(root);
  if (descuento !== 0) advertencias.push(['DESCUENTO', tipo, factura, origen, `El CFDI tiene descuento por ${descuento}.`]);
  if (conceptos.tieneTasaCero) advertencias.push(['TASA 0%', tipo, factura, origen, 'Existe IVA a tasa 0%; se presenta en EXENTO.']);
  if (conceptos.tieneNoObjeto) advertencias.push(['NO OBJETO', tipo, factura, origen, 'Existe concepto no objeto; se presenta en EXENTO.']);
  if (conceptos.tieneOtraTasa) advertencias.push(['TASA NO SOPORTADA', tipo, factura, origen, 'Se encontro una tasa de IVA distinta de 8%, 16%, 0% o exento.']);

  const fila = tipo === 'emitida'
    ? calcularEmitida(total, iva, conceptos, factura, origen, advertencias, formatearFecha(fechaIso), nombreReceptor)
    : calcularRecibida(total, iva, conceptos, factura, origen, advertencias, formatearFecha(fechaIso), nombreEmisor);
  fila.push(clasificarMetodo(metodoPago));
  fila.push(tipo === 'emitida' ? rfcReceptor : rfcEmisor);
  fila.push(uuid);

  const baseCedula = tipo === 'emitida' ? fila[3] + fila[4] : fila[6];
  if (Math.abs(subtotalCfdi - baseCedula) >= CONFIG_CEDULAS.TOLERANCIA) {
    advertencias.push(['SUBTOTAL DIFERENTE', tipo, factura, origen, `SubTotal CFDI: ${subtotalCfdi}; base de cedula: ${baseCedula}.`]);
  }
  fila.push(determinarEstatus(advertencias));
  return { incluir: true, fila, fechaOrden: fechaAOrden(fechaIso), fechaIso, factura, uuid, clienteInfo, advertencias };
}

function obtenerResumenCfdi(contenidoXml) {
  const root = obtenerRoot(contenidoXml);
  return {
    tipo: atributoTexto(root, 'TipoDeComprobante', ''),
    uuid: String(obtenerUuid(root) || '').toUpperCase(),
    total: atributoNumero(root, 'Total')
  };
}

function procesarComplementoPago(root, tipo, origen, ingresosPorUuid, uuid, factura, fechaIso, clienteInfo, nombreEmisor, nombreReceptor, advertencias) {
  const documentos = buscarElementos(root, 'DoctoRelacionado');
  const grupos = {};

  documentos.forEach((documento, i) => {
    const idRelacionado = atributoTexto(documento, 'IdDocumento', '').toUpperCase();
    const clave = idRelacionado || `SIN-UUID-${i}`;
    if (!grupos[clave]) grupos[clave] = { id: idRelacionado, pagado: 0, iva8: 0, iva16: 0, folios: [] };
    const grupo = grupos[clave];
    grupo.pagado += atributoNumero(documento, 'ImpPagado');
    grupo.folios.push(`${atributoTexto(documento, 'Serie', '')} ${atributoTexto(documento, 'Folio', '')}`.trim());
    buscarElementos(documento, 'TrasladoDR').forEach((traslado) => {
      if (atributoTexto(traslado, 'ImpuestoDR', '') !== '002') return;
      const factor = atributoTexto(traslado, 'TipoFactorDR', '');
      const tasa = atributoNumero(traslado, 'TasaOCuotaDR');
      const importe = atributoNumero(traslado, 'ImporteDR');
      if (factor !== 'Exento' && tasasIguales(tasa, 0.16)) grupo.iva16 += importe;
      else if (factor !== 'Exento' && tasasIguales(tasa, 0.08)) grupo.iva8 += importe;
    });
  });

  const pago = { total: 0, iva8: 0, iva16: 0, omitidos: [], incluidos: [] };
  Object.keys(grupos).forEach((clave) => {
    const actual = grupos[clave];
    const ingresoOriginal = actual.id ? ingresosPorUuid[actual.id] : null;
    const esDuplicado = CONFIG_CEDULAS.EVITAR_PAGO_TOTAL_DUPLICADO && ingresoOriginal && Math.abs(actual.pagado - ingresoOriginal.total) < CONFIG_CEDULAS.TOLERANCIA;
    if (esDuplicado) pago.omitidos.push(actual.folios.join(', ') || actual.id);
    else {
      pago.total += actual.pagado;
      pago.iva8 += actual.iva8;
      pago.iva16 += actual.iva16;
      pago.incluidos.push(actual.folios.join(', ') || actual.id);
    }
  });

  if (!documentos.length) {
    const totales = buscarElemento(root, 'Totales');
    pago.total = atributoNumero(totales, 'MontoTotalPagos');
    pago.iva8 = atributoNumero(totales, 'TotalTrasladosImpuestoIVA8');
    pago.iva16 = atributoNumero(totales, 'TotalTrasladosImpuestoIVA16');
    pago.incluidos.push('SIN DOCUMENTO RELACIONADO');
    advertencias.push(['PAGO SIN RELACION', tipo, factura, origen, 'Se uso el total global del complemento porque no se encontro DoctoRelacionado.']);
  }

  if (pago.omitidos.length) advertencias.push(['PAGO NO DUPLICADO', tipo, factura, origen, `No se sumo el pago total de: ${pago.omitidos.join('; ')}.`]);
  if (pago.total < CONFIG_CEDULAS.TOLERANCIA) return registroExcluido(uuid, factura, fechaIso, clienteInfo, advertencias);

  const base8 = pago.iva8 ? pago.iva8 / 0.08 : 0;
  const base16 = pago.iva16 ? pago.iva16 / 0.16 : 0;
  const exento = normalizarCero(pago.total - base8 - base16 - pago.iva8 - pago.iva16);
  let fila;
  if (tipo === 'emitida') {
    if (pago.iva8) advertencias.push(['PAGO EMITIDO IVA 8%', tipo, factura, origen, 'Emitidas no tiene columnas al 8%; requiere revision.']);
    fila = [formatearFecha(fechaIso), `${factura} (P)`, nombreReceptor, base16, normalizarCero(pago.total - base16 - pago.iva16), pago.iva16, pago.total, 'COMPLEMENTO', atributoTexto(child(root, 'Receptor'), 'Rfc', ''), uuid, determinarEstatus(advertencias)];
  } else {
    fila = [formatearFecha(fechaIso), `${factura} (P)`, nombreEmisor, base8, base16, exento, base8 + base16 + exento, pago.iva8, pago.iva16, pago.total, 'COMPLEMENTO', atributoTexto(child(root, 'Emisor'), 'Rfc', ''), uuid, determinarEstatus(advertencias)];
  }
  advertencias.push(['PAGO INTEGRADO', tipo, factura, origen, `Complemento P integrado por ${pago.total}. Relacionado con: ${pago.incluidos.join('; ')}.`]);
  return { incluir: true, fila, fechaOrden: fechaAOrden(fechaIso), fechaIso, factura, uuid, clienteInfo, advertencias };
}

function calcularEmitida(total, iva, conceptos, factura, origen, advertencias, fecha, receptor) {
  let ingresos = 0;
  let exento = 0;
  let iva16 = 0;
  const tieneNoGravado = conceptos.tieneExento || conceptos.tieneTasaCero || conceptos.tieneNoObjeto;
  if (conceptos.tieneIva8 || Math.abs(iva.iva8) >= CONFIG_CEDULAS.TOLERANCIA) advertencias.push(['EMITIDA IVA 8%', 'emitida', factura, origen, 'La cedula de emitidas no tiene columnas de base e IVA al 8%; revisar manualmente.']);
  if (!tieneNoGravado && conceptos.tieneIva16 && !conceptos.tieneIva8) {
    ingresos = total / 1.16;
    iva16 = ingresos * 0.16;
  } else if (conceptos.tieneIva16 || Math.abs(iva.iva16) >= CONFIG_CEDULAS.TOLERANCIA) {
    iva16 = iva.iva16;
    ingresos = iva16 / 0.16;
    exento = normalizarCero(total - ingresos - iva16 - iva.iva8);
  } else {
    exento = total;
  }
  return [fecha, factura, receptor, ingresos, exento, iva16, total];
}

function calcularRecibida(total, iva, conceptos, factura, origen, advertencias, fecha, emisor) {
  let gto8 = 0, gto16 = 0, exento = 0, iva8 = 0, iva16 = 0;
  const tieneNoGravado = conceptos.tieneExento || conceptos.tieneTasaCero || conceptos.tieneNoObjeto;
  const solo16 = conceptos.tieneIva16 && !conceptos.tieneIva8 && !tieneNoGravado;
  const solo8 = conceptos.tieneIva8 && !conceptos.tieneIva16 && !tieneNoGravado;
  if (solo16) {
    gto16 = total / 1.16;
    iva16 = gto16 * 0.16;
  } else if (solo8) {
    gto8 = total / 1.08;
    iva8 = gto8 * 0.08;
  } else if (conceptos.tieneIva8 || conceptos.tieneIva16 || Math.abs(iva.iva8) >= CONFIG_CEDULAS.TOLERANCIA || Math.abs(iva.iva16) >= CONFIG_CEDULAS.TOLERANCIA) {
    iva8 = iva.iva8;
    iva16 = iva.iva16;
    gto8 = iva8 === 0 ? 0 : iva8 / 0.08;
    gto16 = iva16 === 0 ? 0 : iva16 / 0.16;
    exento = normalizarCero(total - gto8 - gto16 - iva8 - iva16);
  } else {
    exento = total;
  }
  if (!solo8 && !solo16 && !tieneNoGravado && conceptos.tieneIva8 && conceptos.tieneIva16) advertencias.push(['IVA 8% Y 16%', 'recibida', factura, origen, 'Factura con ambas tasas; las bases se reconstruyeron desde cada IVA.']);
  return [fecha, factura, emisor, gto8, gto16, exento, gto8 + gto16 + exento, iva8, iva16, total];
}

function analizarConceptos(root) {
  const resultado = { tieneIva8: false, tieneIva16: false, tieneExento: false, tieneTasaCero: false, tieneNoObjeto: false, tieneOtraTasa: false };
  for (const concepto of children(child(root, 'Conceptos'), 'Concepto')) {
    const objetoImp = atributoTexto(concepto, 'ObjetoImp', '');
    const importeNeto = atributoNumero(concepto, 'Importe') - atributoNumero(concepto, 'Descuento');
    const tieneImporteNeto = Math.abs(importeNeto) >= CONFIG_CEDULAS.TOLERANCIA;
    let encontroIva = false;
    for (const traslado of children(child(child(concepto, 'Impuestos'), 'Traslados'), 'Traslado')) {
      if (atributoTexto(traslado, 'Impuesto', '') !== '002') continue;
      encontroIva = true;
      const tipoFactor = atributoTexto(traslado, 'TipoFactor', '');
      const tasa = atributoNumero(traslado, 'TasaOCuota');
      if (tipoFactor === 'Exento') resultado.tieneExento = true;
      else if (tasasIguales(tasa, 0.16)) resultado.tieneIva16 = true;
      else if (tasasIguales(tasa, 0.08)) resultado.tieneIva8 = true;
      else if (tasasIguales(tasa, 0)) resultado.tieneTasaCero = true;
      else resultado.tieneOtraTasa = true;
    }
    if (tieneImporteNeto && (objetoImp === '01' || objetoImp === '03' || objetoImp === '04')) resultado.tieneNoObjeto = true;
    else if (tieneImporteNeto && !encontroIva && objetoImp !== '02') resultado.tieneNoObjeto = true;
  }
  return resultado;
}

function obtenerIvaCfdi(root) {
  const global = sumarIvaEnImpuestos(child(root, 'Impuestos'));
  if (global.encontroIva) return global;
  const resultado = { iva8: 0, iva16: 0, encontroIva: false };
  for (const concepto of children(child(root, 'Conceptos'), 'Concepto')) {
    const parcial = sumarIvaEnImpuestos(child(concepto, 'Impuestos'));
    resultado.iva8 += parcial.iva8;
    resultado.iva16 += parcial.iva16;
    resultado.encontroIva = resultado.encontroIva || parcial.encontroIva;
  }
  return resultado;
}

function sumarIvaEnImpuestos(impuestosNode) {
  const resultado = { iva8: 0, iva16: 0, encontroIva: false };
  for (const traslado of children(child(impuestosNode, 'Traslados'), 'Traslado')) {
    if (atributoTexto(traslado, 'Impuesto', '') !== '002' || atributoTexto(traslado, 'TipoFactor', '') === 'Exento') continue;
    const tasa = atributoNumero(traslado, 'TasaOCuota');
    const importe = atributoNumero(traslado, 'Importe');
    if (tasasIguales(tasa, 0.08)) {
      resultado.iva8 += importe;
      resultado.encontroIva = true;
    } else if (tasasIguales(tasa, 0.16)) {
      resultado.iva16 += importe;
      resultado.encontroIva = true;
    }
  }
  return resultado;
}

function obtenerRoot(xml) {
  const parsed = parser.parse(xml);
  const root = parsed.Comprobante || parsed['cfdi:Comprobante'];
  if (!root) throw new Error('La raiz no es cfdi:Comprobante.');
  return root;
}

function child(node, name) {
  const value = node && node[name];
  return Array.isArray(value) ? value[0] : value;
}

function children(node, name) {
  if (!node || !node[name]) return [];
  return Array.isArray(node[name]) ? node[name] : [node[name]];
}

function buscarElemento(elemento, nombreBuscado) {
  return buscarElementos(elemento, nombreBuscado)[0] || null;
}

function buscarElementos(elemento, nombreBuscado) {
  const resultados = [];
  recorrer(elemento, (valor, clave) => {
    if (clave === nombreBuscado) resultados.push(...(Array.isArray(valor) ? valor : [valor]));
  });
  return resultados;
}

function recorrer(valor, cb, claveActual = '') {
  if (!valor || typeof valor !== 'object') return;
  Object.keys(valor).forEach((clave) => {
    const hijo = valor[clave];
    if (hijo && typeof hijo === 'object') {
      cb(hijo, clave);
      if (Array.isArray(hijo)) hijo.forEach((item) => recorrer(item, cb, clave));
      else recorrer(hijo, cb, clave);
    }
  });
}

function filtrarRegistros(registros, filtro) {
  return registros.filter((registro) => {
    const fecha = new Date(registro.fechaIso || registro.fechaOrden || 0);
    if (Number.isNaN(fecha.getTime())) return false;
    const anio = fecha.getFullYear();
    const mes = fecha.getMonth() + 1;
    if (anio !== filtro.anio) return false;
    if (filtro.modo === 'anual') return true;
    if (filtro.modo === 'trimestral') {
      const inicio = (filtro.trimestre - 1) * 3 + 1;
      return mes >= inicio && mes <= inicio + 2;
    }
    return mes === filtro.mes;
  });
}

function describirPeriodo(filtro) {
  const meses = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
  if (filtro.modo === 'anual') return `ANUAL ${filtro.anio}`;
  if (filtro.modo === 'trimestral') return `TRIMESTRE ${filtro.trimestre} ${filtro.anio}`;
  return `${meses[filtro.mes - 1] || 'MES'} ${filtro.anio}`;
}

function agregarAdvertenciaPeriodo(advertencias, tipo, xmlEncontrados, registros, filtrados, filtro) {
  if (!xmlEncontrados || !registros.length || filtrados.length) return;
  advertencias.push([
    'FUERA DE PERIODO',
    tipo,
    '',
    '',
    `Se encontraron ${registros.length} CFDI aplicables, pero ninguno cae dentro de ${describirPeriodo(filtro)}. Fechas detectadas: ${describirFechasRegistros(registros)}.`
  ]);
}

function describirFechasRegistros(registros) {
  const meses = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
  const periodos = {};
  registros.forEach((registro) => {
    const fecha = new Date(registro.fechaIso || registro.fechaOrden || 0);
    if (Number.isNaN(fecha.getTime())) return;
    const clave = `${fecha.getFullYear()}-${String(fecha.getMonth() + 1).padStart(2, '0')}`;
    periodos[clave] = `${meses[fecha.getMonth()]} ${fecha.getFullYear()}`;
  });
  const valores = Object.keys(periodos).sort().map((clave) => periodos[clave]);
  return valores.length ? valores.join(', ') : 'sin fechas validas';
}

function sumarColumnas(filas, desde, hasta) {
  const totales = {};
  for (let i = desde; i <= hasta; i++) totales[i] = filas.reduce((suma, fila) => suma + (Number(fila[i]) || 0), 0);
  return totales;
}

function validarClienteUnico(emitidas, recibidas) {
  const clientes = { ...emitidas.clientesEncontrados, ...recibidas.clientesEncontrados };
  const encontrados = Object.keys(clientes);
  if (encontrados.length > 1) throw new Error(`Las cargas contienen CFDI de varios clientes: ${encontrados.join(', ')}. Separe los XML por RFC.`);
  if (encontrados.length === 1) return { rfc: encontrados[0], nombre: clientes[encontrados[0]] };
  return emitidas.clienteInfo.nombre ? emitidas.clienteInfo : recibidas.clienteInfo;
}

function registroExcluido(uuid, factura, fechaIso, clienteInfo, advertencias) {
  return { incluir: false, fila: null, fechaOrden: fechaAOrden(fechaIso), fechaIso, factura, uuid, clienteInfo, advertencias };
}

function crearRegistroCancelado(tipo, tipoComprobante, fechaIso, factura, uuid, clienteInfo, nombreEmisor, nombreReceptor, rfcEmisor, rfcReceptor, metodoPago, advertencias) {
  const clasificacion = tipoComprobante === 'P' ? 'COMPLEMENTO' : clasificarMetodo(metodoPago);
  const fila = tipo === 'emitida'
    ? [formatearFecha(fechaIso), factura, nombreReceptor, 0, 0, 0, 0, clasificacion, rfcReceptor, uuid, 'CANCELADA']
    : [formatearFecha(fechaIso), factura, nombreEmisor, 0, 0, 0, 0, 0, 0, 0, clasificacion, rfcEmisor, uuid, 'CANCELADA'];
  return { incluir: true, fila, fechaOrden: fechaAOrden(fechaIso), fechaIso, factura, uuid, clienteInfo, advertencias };
}

function estaCancelado(factura, uuid, estadoXml) {
  const objetivoFolio = normalizarIdentificador(factura);
  const objetivoUuid = String(uuid || '').toUpperCase();
  if (String(estadoXml || '').toUpperCase().includes('CANCEL')) return true;
  return CONFIG_CEDULAS.CFDI_CANCELADOS.some((valor) => normalizarIdentificador(valor) === objetivoFolio || String(valor).toUpperCase() === objetivoUuid);
}

function atributoTexto(elemento, nombre, predeterminado) {
  return elemento && Object.prototype.hasOwnProperty.call(elemento, nombre) ? String(elemento[nombre]) : predeterminado;
}

function atributoNumero(elemento, nombre) {
  const numero = Number(atributoTexto(elemento, nombre, '0'));
  return Number.isFinite(numero) ? numero : 0;
}

function tasasIguales(a, b) {
  return Math.abs(a - b) < 0.000001;
}

function normalizarCero(numero) {
  return Math.abs(numero) < 0.000000001 ? 0 : numero;
}

function formatearFecha(fechaIso) {
  if (!fechaIso) return '';
  const partes = fechaIso.substring(0, 10).split('-');
  return partes.length === 3 ? `${partes[2]}/${partes[1]}/${partes[0]}` : fechaIso;
}

function fechaAOrden(fechaIso) {
  const tiempo = fechaIso ? new Date(fechaIso).getTime() : 0;
  return Number.isNaN(tiempo) ? 0 : tiempo;
}

function obtenerUuid(root) {
  const timbre = buscarElemento(root, 'TimbreFiscalDigital');
  return timbre ? atributoTexto(timbre, 'UUID', '') : '';
}

function normalizarIdentificador(valor) {
  return String(valor || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function clasificarMetodo(metodoPago) {
  const metodo = String(metodoPago || '').toUpperCase();
  return metodo === 'PUE' || metodo === 'PPD' ? metodo : 'REVISAR';
}

function determinarEstatus(advertencias) {
  const revisar = new Set(['TASA NO SOPORTADA', 'EMITIDA IVA 8%', 'PAGO EMITIDO IVA 8%', 'PAGO SIN RELACION']);
  return advertencias.some((advertencia) => revisar.has(advertencia[0])) ? 'REVISAR' : 'VIGENTE';
}

function huellaXml(contenidoXml) {
  return crypto.createHash('sha256').update(contenidoXml, 'utf8').digest('hex');
}

module.exports = { procesarLoteCfdi };
