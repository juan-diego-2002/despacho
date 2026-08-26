/**
 * CEDULAS CFDI - VERSION 3.4
 *
 * Reglas principales:
 * 1) Integra CFDI tipo I y complementos de pago tipo P sin duplicar pagos totales
 *    cuya factura relacionada ya se encuentre en la misma carpeta.
 * 2) Excluye notas de credito/egreso (E), traslados (T) y nomina (N).
 * 3) Distingue facturas totalmente gravadas de facturas con conceptos no gravados.
 * 4) Totalmente gravada: parte del Total (Total/1.16 o Total/1.08).
 * 5) Con exento/no gravado: parte del IVA del CFDI (IVA/tasa) y obtiene el
 *    no gravado como diferencia.
 * 6) Conserva todos los decimales internos y solo muestra dos decimales.
 */

var CONFIG_CEDULAS = {
  // Puede pegarse solamente el ID o la liga completa de la carpeta de Google Drive.
  ID_CARPETA_EMITIDAS: '1eA0E6P-X4OacKcfgk26KhXn9TEl7EMAw',
  ID_CARPETA_RECIBIDAS: '1osfJue2c6CFtVZ3en_3l5SCjGasKyTt2',
  TOLERANCIA: 0.005,
  FORMATO_MONEDA: '$#,##0.00',
  // true: no suma un P cuando liquida totalmente un CFDI I presente en la carpeta.
  EVITAR_PAGO_TOTAL_DUPLICADO: true,
  // El XML no informa su estado actual en el SAT. Agregar folios o UUID cancelados del periodo.
  CFDI_CANCELADOS: []
};

function procesarFacturasPorRutaDirecta() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var emitidas = leerCarpetaCfdi_(CONFIG_CEDULAS.ID_CARPETA_EMITIDAS, 'emitida');
  var recibidas = leerCarpetaCfdi_(CONFIG_CEDULAS.ID_CARPETA_RECIBIDAS, 'recibida');
  var cliente = validarClienteUnico_(emitidas, recibidas);
  var periodo = detectarPeriodo_(emitidas.filas.concat(recibidas.filas));
  var advertencias = emitidas.advertencias.concat(recibidas.advertencias);

  sheet.clear();
  var fila = construirEncabezado_(sheet, cliente, periodo);
  fila = escribirEmitidas_(sheet, fila, emitidas.filas) + 2;
  fila = escribirRecibidas_(sheet, fila, recibidas.filas) + 2;
  aplicarPresentacionGeneral_(sheet);

  Logger.log('Emitidas integradas: ' + emitidas.filas.length);
  Logger.log('Recibidas integradas: ' + recibidas.filas.length);
  Logger.log('Advertencias/exclusiones: ' + advertencias.length);
  for (var i = 0; i < advertencias.length; i++) Logger.log(advertencias[i].join(' | '));
}

function leerCarpetaCfdi_(idCarpeta, tipo) {
  var registros = [];
  var clienteInfo = { nombre: '', rfc: '' };
  var advertencias = [];
  var uuidVistos = {};
  var xmlPendientes = [];
  var clientesEncontrados = {};

  try {
    var idLimpio = limpiarIdCarpeta_(idCarpeta);
    var carpeta = DriveApp.getFolderById(idLimpio);
    var archivos = carpeta.getFiles();

    while (archivos.hasNext()) {
      var archivo = archivos.next();
      xmlPendientes = xmlPendientes.concat(extraerXmls_(archivo, advertencias));
    }

    // Primer recorrido: conocer los CFDI I disponibles para no duplicarlos con pagos totales P.
    var ingresosPorUuid = {};
    for (var p = 0; p < xmlPendientes.length; p++) {
      try {
        var resumen = obtenerResumenCfdi_(xmlPendientes[p].contenido);
        if (resumen.tipo === 'I' && resumen.uuid) ingresosPorUuid[resumen.uuid] = resumen;
      } catch (errorResumen) {
        // El error completo se mostrara durante el recorrido principal.
      }
    }

    for (var i = 0; i < xmlPendientes.length; i++) {
        try {
          var registro = procesarXmlCfdi_(xmlPendientes[i].contenido, tipo,
            xmlPendientes[i].origen, ingresosPorUuid);

          if (registro.uuid && uuidVistos[registro.uuid]) {
            advertencias.push(['DUPLICADO', tipo, registro.factura, xmlPendientes[i].origen,
              'UUID repetido; se integro una sola vez.']);
            continue;
          }
          if (registro.uuid) uuidVistos[registro.uuid] = true;

          if (!clienteInfo.nombre && registro.clienteInfo.nombre) {
            clienteInfo = registro.clienteInfo;
          }
          if (registro.clienteInfo.rfc) {
            clientesEncontrados[registro.clienteInfo.rfc] = registro.clienteInfo.nombre || '';
          }

          advertencias = advertencias.concat(registro.advertencias);
          if (registro.incluir) registros.push(registro);
        } catch (errorXml) {
          advertencias.push(['XML INVALIDO', tipo, '', xmlPendientes[i].origen, errorXml.message]);
        }
    }
  } catch (errorCarpeta) {
    throw new Error('No se pudo leer la carpeta ' + idCarpeta + ': ' + errorCarpeta.message);
  }

  registros.sort(function (a, b) {
    if (a.fechaOrden !== b.fechaOrden) return a.fechaOrden - b.fechaOrden;
    return String(a.factura).localeCompare(String(b.factura), undefined, { numeric: true });
  });

  return {
    filas: registros.map(function (registro) { return registro.fila; }),
    clienteInfo: clienteInfo,
    clientesEncontrados: clientesEncontrados,
    advertencias: advertencias
  };
}

function procesarXmlCfdi_(contenidoXml, tipo, origen, ingresosPorUuid) {
  var documento = XmlService.parse(contenidoXml);
  var root = documento.getRootElement();
  var ns = root.getNamespace();

  if (root.getName() !== 'Comprobante') {
    throw new Error('La raiz no es cfdi:Comprobante.');
  }

  var emisor = root.getChild('Emisor', ns);
  var receptor = root.getChild('Receptor', ns);
  var nombreEmisor = atributoTexto_(emisor, 'Nombre', 'DESCONOCIDO');
  var nombreReceptor = atributoTexto_(receptor, 'Nombre', 'DESCONOCIDO');
  var rfcEmisor = atributoTexto_(emisor, 'Rfc', '');
  var rfcReceptor = atributoTexto_(receptor, 'Rfc', '');
  var metodoPago = atributoTexto_(root, 'MetodoPago', '');
  var estadoXml = atributoTexto_(root, 'Estatus', atributoTexto_(root, 'Estado', ''));
  var tipoComprobante = atributoTexto_(root, 'TipoDeComprobante', '');
  var total = atributoNumero_(root, 'Total');
  var subtotalCfdi = atributoNumero_(root, 'SubTotal');
  var descuento = atributoNumero_(root, 'Descuento');
  var fechaIso = atributoTexto_(root, 'Fecha', '');
  var folio = atributoTexto_(root, 'Folio', 'S/F');
  var serie = atributoTexto_(root, 'Serie', '');
  var factura = serie ? serie + ' ' + folio : folio;
  var uuid = obtenerUuid_(root) || huellaXml_(contenidoXml);
  var advertencias = [];

  var clienteInfo = {
    nombre: tipo === 'emitida' ? nombreEmisor : nombreReceptor,
    rfc: tipo === 'emitida' ? rfcEmisor : rfcReceptor
  };

  if (estaCancelado_(factura, uuid, estadoXml)) {
    advertencias.push(['CANCELADA', tipo, factura, origen,
      'CFDI marcado como cancelado en CONFIG_CEDULAS.CFDI_CANCELADOS.']);
    return crearRegistroCancelado_(tipo, tipoComprobante, fechaIso, factura, uuid,
      clienteInfo, nombreEmisor, nombreReceptor, rfcEmisor, rfcReceptor, metodoPago, advertencias);
  }

  if (tipoComprobante === 'P') {
    return procesarComplementoPago_(root, tipo, origen, ingresosPorUuid, uuid,
      factura, fechaIso, clienteInfo, nombreEmisor, nombreReceptor, advertencias);
  }

  if (tipoComprobante !== 'I') {
    advertencias.push(['EXCLUIDO TIPO ' + (tipoComprobante || 'VACIO'), tipo, factura, origen,
      'La cedula integra CFDI I y complementos P; este tipo se excluye.']);
    return registroExcluido_(uuid, factura, fechaIso, clienteInfo, advertencias);
  }

  if (total <= 0) {
    advertencias.push(['EXCLUIDO TOTAL CERO', tipo, factura, origen,
      'El CFDI tipo I no tiene importe positivo.']);
    return registroExcluido_(uuid, factura, fechaIso, clienteInfo, advertencias);
  }

  var iva = obtenerIvaCfdi_(root, ns);
  var conceptos = analizarConceptos_(root, ns);

  if (descuento !== 0) {
    advertencias.push(['DESCUENTO', tipo, factura, origen,
      'El CFDI tiene descuento por ' + descuento + '. La regla parte del Total para facturas totalmente gravadas.']);
  }
  if (conceptos.tieneTasaCero) {
    advertencias.push(['TASA 0%', tipo, factura, origen,
      'Existe IVA a tasa 0%; se presenta en EXENTO por falta de una columna separada, pero fiscalmente no es exento.']);
  }
  if (conceptos.tieneNoObjeto) {
    advertencias.push(['NO OBJETO', tipo, factura, origen,
      'Existe concepto no objeto; se presenta en EXENTO por la estructura actual de la cedula.']);
  }
  if (conceptos.tieneOtraTasa) {
    advertencias.push(['TASA NO SOPORTADA', tipo, factura, origen,
      'Se encontro una tasa de IVA distinta de 8%, 16%, 0% o exento. Revisar manualmente.']);
  }

  var fila;
  if (tipo === 'emitida') {
    fila = calcularEmitida_(total, iva, conceptos, factura, origen, advertencias,
      formatearFecha_(fechaIso), nombreReceptor);
  } else {
    fila = calcularRecibida_(total, iva, conceptos, factura, origen, advertencias,
      formatearFecha_(fechaIso), nombreEmisor);
  }

  fila.push(clasificarMetodo_(metodoPago));
  fila.push(tipo === 'emitida' ? rfcReceptor : rfcEmisor);
  fila.push(uuid);

  var baseCedula = tipo === 'emitida' ? fila[3] + fila[4] : fila[6];
  if (Math.abs(subtotalCfdi - baseCedula) >= CONFIG_CEDULAS.TOLERANCIA) {
    advertencias.push(['SUBTOTAL DIFERENTE', tipo, factura, origen,
      'SubTotal CFDI: ' + subtotalCfdi + '; base de cedula: ' + baseCedula +
      '. Puede deberse a IEPS, descuento, retencion o al metodo manual basado en Total.']);
  }
  fila.push(determinarEstatus_(advertencias));

  return {
    incluir: true,
    fila: fila,
    fechaOrden: fechaAOrden_(fechaIso),
    factura: factura,
    uuid: uuid,
    clienteInfo: clienteInfo,
    advertencias: advertencias
  };
}

function obtenerResumenCfdi_(contenidoXml) {
  var root = XmlService.parse(contenidoXml).getRootElement();
  return {
    tipo: atributoTexto_(root, 'TipoDeComprobante', ''),
    uuid: String(obtenerUuid_(root) || '').toUpperCase(),
    total: atributoNumero_(root, 'Total')
  };
}

function estaCancelado_(factura, uuid, estadoXml) {
  var objetivoFolio = normalizarIdentificador_(factura);
  var objetivoUuid = String(uuid || '').toUpperCase();
  if (String(estadoXml || '').toUpperCase().indexOf('CANCEL') >= 0) return true;
  for (var i = 0; i < CONFIG_CEDULAS.CFDI_CANCELADOS.length; i++) {
    var valor = String(CONFIG_CEDULAS.CFDI_CANCELADOS[i] || '');
    if (normalizarIdentificador_(valor) === objetivoFolio || valor.toUpperCase() === objetivoUuid) return true;
  }
  return false;
}

function procesarComplementoPago_(root, tipo, origen, ingresosPorUuid, uuid,
    factura, fechaIso, clienteInfo, nombreEmisor, nombreReceptor, advertencias) {
  var documentos = buscarElementos_(root, 'DoctoRelacionado');
  var grupos = {};

  for (var i = 0; i < documentos.length; i++) {
    var documento = documentos[i];
    var idRelacionado = atributoTexto_(documento, 'IdDocumento', '').toUpperCase();
    var clave = idRelacionado || ('SIN-UUID-' + i);
    if (!grupos[clave]) {
      grupos[clave] = { id: idRelacionado, pagado: 0, iva8: 0, iva16: 0,
        tieneTasaCero: false, tieneExento: false, folios: [] };
    }
    var grupo = grupos[clave];
    grupo.pagado += atributoNumero_(documento, 'ImpPagado');
    grupo.folios.push((atributoTexto_(documento, 'Serie', '') + ' ' +
      atributoTexto_(documento, 'Folio', '')).trim());

    var traslados = buscarElementos_(documento, 'TrasladoDR');
    for (var t = 0; t < traslados.length; t++) {
      if (atributoTexto_(traslados[t], 'ImpuestoDR', '') !== '002') continue;
      var factor = atributoTexto_(traslados[t], 'TipoFactorDR', '');
      var tasa = atributoNumero_(traslados[t], 'TasaOCuotaDR');
      var importe = atributoNumero_(traslados[t], 'ImporteDR');
      if (factor === 'Exento') grupo.tieneExento = true;
      else if (tasasIguales_(tasa, 0.16)) grupo.iva16 += importe;
      else if (tasasIguales_(tasa, 0.08)) grupo.iva8 += importe;
      else if (tasasIguales_(tasa, 0)) grupo.tieneTasaCero = true;
    }
  }

  var pago = { total: 0, iva8: 0, iva16: 0, omitidos: [], incluidos: [] };
  var claves = Object.keys(grupos);
  for (var g = 0; g < claves.length; g++) {
    var actual = grupos[claves[g]];
    var ingresoOriginal = actual.id ? ingresosPorUuid[actual.id] : null;
    var esPagoTotalDuplicado = CONFIG_CEDULAS.EVITAR_PAGO_TOTAL_DUPLICADO && ingresoOriginal &&
      Math.abs(actual.pagado - ingresoOriginal.total) < CONFIG_CEDULAS.TOLERANCIA;

    if (esPagoTotalDuplicado) {
      pago.omitidos.push(actual.folios.join(', ') || actual.id);
    } else {
      pago.total += actual.pagado;
      pago.iva8 += actual.iva8;
      pago.iva16 += actual.iva16;
      pago.incluidos.push(actual.folios.join(', ') || actual.id);
    }
  }

  // Respaldo para complementos sin DoctoRelacionado legible.
  if (!claves.length) {
    var totales = buscarElemento_(root, 'Totales');
    pago.total = atributoNumero_(totales, 'MontoTotalPagos');
    pago.iva8 = atributoNumero_(totales, 'TotalTrasladosImpuestoIVA8');
    pago.iva16 = atributoNumero_(totales, 'TotalTrasladosImpuestoIVA16');
    pago.incluidos.push('SIN DOCUMENTO RELACIONADO');
    advertencias.push(['PAGO SIN RELACION', tipo, factura, origen,
      'Se uso el total global del complemento porque no se encontro DoctoRelacionado.']);
  }

  if (pago.omitidos.length) {
    advertencias.push(['PAGO NO DUPLICADO', tipo, factura, origen,
      'No se sumo el pago total de: ' + pago.omitidos.join('; ') +
      ', porque la factura I relacionada ya esta incluida.']);
  }

  if (pago.total < CONFIG_CEDULAS.TOLERANCIA) {
    return registroExcluido_(uuid, factura, fechaIso, clienteInfo, advertencias);
  }

  var base8 = pago.iva8 ? pago.iva8 / 0.08 : 0;
  var base16 = pago.iva16 ? pago.iva16 / 0.16 : 0;
  var exento = normalizarCero_(pago.total - base8 - base16 - pago.iva8 - pago.iva16);
  var fecha = formatearFecha_(fechaIso);
  var fila;

  if (tipo === 'emitida') {
    if (pago.iva8) {
      advertencias.push(['PAGO EMITIDO IVA 8%', tipo, factura, origen,
        'Emitidas no tiene columnas al 8%; el importe restante aparece en EXENTO y requiere revision.']);
    }
    var estatusPagoEmitido = determinarEstatus_(advertencias);
    fila = [fecha, factura + ' (P)', nombreReceptor, base16,
      normalizarCero_(pago.total - base16 - pago.iva16), pago.iva16, pago.total,
      'COMPLEMENTO', atributoTexto_(root.getChild('Receptor', root.getNamespace()), 'Rfc', ''), uuid,
      estatusPagoEmitido];
  } else {
    var estatusPagoRecibido = determinarEstatus_(advertencias);
    fila = [fecha, factura + ' (P)', nombreEmisor, base8, base16, exento,
      base8 + base16 + exento, pago.iva8, pago.iva16, pago.total,
      'COMPLEMENTO', atributoTexto_(root.getChild('Emisor', root.getNamespace()), 'Rfc', ''), uuid,
      estatusPagoRecibido];
  }

  advertencias.push(['PAGO INTEGRADO', tipo, factura, origen,
    'Complemento P integrado por ' + pago.total + '. Relacionado con: ' + pago.incluidos.join('; ') + '.']);
  return {
    incluir: true,
    fila: fila,
    fechaOrden: fechaAOrden_(fechaIso),
    factura: factura,
    uuid: uuid,
    clienteInfo: clienteInfo,
    advertencias: advertencias
  };
}

function calcularEmitida_(total, iva, conceptos, factura, origen, advertencias, fecha, receptor) {
  var ingresos = 0;
  var exento = 0;
  var iva16 = 0;
  var tieneNoGravado = conceptos.tieneExento || conceptos.tieneTasaCero || conceptos.tieneNoObjeto;

  if (conceptos.tieneIva8 || Math.abs(iva.iva8) >= CONFIG_CEDULAS.TOLERANCIA) {
    advertencias.push(['EMITIDA IVA 8%', 'emitida', factura, origen,
      'La cedula de emitidas no tiene columnas de base e IVA al 8%; revisar manualmente.']);
  }

  if (!tieneNoGravado && conceptos.tieneIva16 && !conceptos.tieneIva8) {
    // Factura completamente gravada al 16%: reproduce el Excel.
    ingresos = total / 1.16;
    iva16 = ingresos * 0.16;
    exento = 0;
  } else if (conceptos.tieneIva16 || Math.abs(iva.iva16) >= CONFIG_CEDULAS.TOLERANCIA) {
    // Factura mixta 16% + no gravado: parte del IVA informado.
    iva16 = iva.iva16;
    ingresos = iva16 / 0.16;
    exento = normalizarCero_(total - ingresos - iva16 - iva.iva8);
  } else {
    // Sin IVA 16%: toda la factura va a la columna disponible de EXENTO.
    ingresos = 0;
    iva16 = 0;
    exento = total;
  }

  return [fecha, factura, receptor, ingresos, exento, iva16, total];
}

function calcularRecibida_(total, iva, conceptos, factura, origen, advertencias, fecha, emisor) {
  var gto8 = 0;
  var gto16 = 0;
  var exento = 0;
  var iva8 = 0;
  var iva16 = 0;
  var tieneNoGravado = conceptos.tieneExento || conceptos.tieneTasaCero || conceptos.tieneNoObjeto;
  var solo16 = conceptos.tieneIva16 && !conceptos.tieneIva8 && !tieneNoGravado;
  var solo8 = conceptos.tieneIva8 && !conceptos.tieneIva16 && !tieneNoGravado;

  if (solo16) {
    // Totalmente gravada al 16%.
    gto16 = total / 1.16;
    iva16 = gto16 * 0.16;
  } else if (solo8) {
    // Totalmente gravada al 8%.
    gto8 = total / 1.08;
    iva8 = gto8 * 0.08;
  } else if (conceptos.tieneIva8 || conceptos.tieneIva16 ||
             Math.abs(iva.iva8) >= CONFIG_CEDULAS.TOLERANCIA ||
             Math.abs(iva.iva16) >= CONFIG_CEDULAS.TOLERANCIA) {
    // Mixta: 8%, 16% y/o no gravado. Parte de los IVA informados.
    iva8 = iva.iva8;
    iva16 = iva.iva16;
    gto8 = iva8 === 0 ? 0 : iva8 / 0.08;
    gto16 = iva16 === 0 ? 0 : iva16 / 0.16;
    exento = normalizarCero_(total - gto8 - gto16 - iva8 - iva16);
  } else {
    // Sin IVA: toda la factura se presenta en EXENTO.
    exento = total;
  }

  if (!solo8 && !solo16 && !tieneNoGravado && conceptos.tieneIva8 && conceptos.tieneIva16) {
    advertencias.push(['IVA 8% Y 16%', 'recibida', factura, origen,
      'Factura con ambas tasas; las bases se reconstruyeron desde cada IVA.']);
  }

  var subtotal = gto8 + gto16 + exento;
  return [fecha, factura, emisor, gto8, gto16, exento, subtotal, iva8, iva16, total];
}

function analizarConceptos_(root, ns) {
  var resultado = {
    tieneIva8: false,
    tieneIva16: false,
    tieneExento: false,
    tieneTasaCero: false,
    tieneNoObjeto: false,
    tieneOtraTasa: false
  };
  var conceptosNode = root.getChild('Conceptos', ns);
  if (!conceptosNode) return resultado;

  var conceptos = conceptosNode.getChildren('Concepto', ns);
  for (var i = 0; i < conceptos.length; i++) {
    var concepto = conceptos[i];
    var objetoImp = atributoTexto_(concepto, 'ObjetoImp', '');
    var importeNeto = atributoNumero_(concepto, 'Importe') - atributoNumero_(concepto, 'Descuento');
    var tieneImporteNeto = Math.abs(importeNeto) >= CONFIG_CEDULAS.TOLERANCIA;
    var encontroIva = false;
    var impuestos = concepto.getChild('Impuestos', ns);
    var trasladosNode = impuestos ? impuestos.getChild('Traslados', ns) : null;
    var traslados = trasladosNode ? trasladosNode.getChildren('Traslado', ns) : [];

    for (var j = 0; j < traslados.length; j++) {
      var traslado = traslados[j];
      if (atributoTexto_(traslado, 'Impuesto', '') !== '002') continue;
      encontroIva = true;
      var tipoFactor = atributoTexto_(traslado, 'TipoFactor', '');
      var tasa = atributoNumero_(traslado, 'TasaOCuota');

      if (tipoFactor === 'Exento') resultado.tieneExento = true;
      else if (tasasIguales_(tasa, 0.16)) resultado.tieneIva16 = true;
      else if (tasasIguales_(tasa, 0.08)) resultado.tieneIva8 = true;
      else if (tasasIguales_(tasa, 0)) resultado.tieneTasaCero = true;
      else resultado.tieneOtraTasa = true;
    }

    // ObjetoImp 01 = no objeto. 03/04 tambien requieren revision y no son base IVA normal.
    // Un concepto bonificado/descontado al 100% no debe convertir toda la factura en mixta.
    if (tieneImporteNeto && (objetoImp === '01' || objetoImp === '03' || objetoImp === '04')) {
      resultado.tieneNoObjeto = true;
    } else if (tieneImporteNeto && !encontroIva && objetoImp !== '02') {
      resultado.tieneNoObjeto = true;
    }
  }
  return resultado;
}

function obtenerIvaCfdi_(root, ns) {
  var global = sumarIvaEnImpuestos_(root.getChild('Impuestos', ns), ns);
  if (global.encontroIva) return global;

  var resultado = { iva8: 0, iva16: 0, encontroIva: false };
  var conceptosNode = root.getChild('Conceptos', ns);
  if (!conceptosNode) return resultado;

  var conceptos = conceptosNode.getChildren('Concepto', ns);
  for (var i = 0; i < conceptos.length; i++) {
    var parcial = sumarIvaEnImpuestos_(conceptos[i].getChild('Impuestos', ns), ns);
    resultado.iva8 += parcial.iva8;
    resultado.iva16 += parcial.iva16;
    resultado.encontroIva = resultado.encontroIva || parcial.encontroIva;
  }
  return resultado;
}

function sumarIvaEnImpuestos_(impuestosNode, ns) {
  var resultado = { iva8: 0, iva16: 0, encontroIva: false };
  if (!impuestosNode) return resultado;
  var trasladosNode = impuestosNode.getChild('Traslados', ns);
  if (!trasladosNode) return resultado;
  var traslados = trasladosNode.getChildren('Traslado', ns);

  for (var i = 0; i < traslados.length; i++) {
    if (atributoTexto_(traslados[i], 'Impuesto', '') !== '002') continue;
    var tipoFactor = atributoTexto_(traslados[i], 'TipoFactor', '');
    if (tipoFactor === 'Exento') continue;
    var tasa = atributoNumero_(traslados[i], 'TasaOCuota');
    var importe = atributoNumero_(traslados[i], 'Importe');
    if (tasasIguales_(tasa, 0.08)) {
      resultado.iva8 += importe;
      resultado.encontroIva = true;
    } else if (tasasIguales_(tasa, 0.16)) {
      resultado.iva16 += importe;
      resultado.encontroIva = true;
    }
  }
  return resultado;
}

function extraerXmls_(archivo, advertencias) {
  var nombre = archivo.getName();
  var minusculas = nombre.toLowerCase();
  var xmls = [];

  if (minusculas.endsWith('.xml')) {
    xmls.push({ contenido: archivo.getBlob().getDataAsString('UTF-8'), origen: nombre });
  } else if (minusculas.endsWith('.zip')) {
    try {
      var blobs = Utilities.unzip(archivo.getBlob());
      for (var i = 0; i < blobs.length; i++) {
        if (blobs[i].getName().toLowerCase().endsWith('.xml')) {
          xmls.push({
            contenido: blobs[i].getDataAsString('UTF-8'),
            origen: nombre + ' > ' + blobs[i].getName()
          });
        }
      }
    } catch (errorZip) {
      advertencias.push(['ZIP INVALIDO', '', '', nombre, errorZip.message]);
    }
  }
  return xmls;
}

function escribirEmitidas_(sheet, filaInicial, filas) {
  var encabezados = ['FECHA', 'FACTURA', 'CONCEPTO', 'INGRESOS', 'EXENTO', 'IVA 16%', 'TOTAL',
    'TIPO / METODO', 'RFC TERCERO', 'UUID', 'ESTATUS'];
  return escribirTabla_(sheet, filaInicial, 'FACTURAS EMITIDAS', encabezados, filas,
    11, 4, 7, 8, 11, '#D9EAF7');
}

function escribirRecibidas_(sheet, filaInicial, filas) {
  var encabezados = ['FECHA', 'FACTURA', 'TERCERO / CONCEPTO', 'GTO 8%', 'GTO 16%',
    'EXENTO', 'SUBTOTAL', 'IVA 8%', 'IVA 16%', 'TOTAL', 'TIPO / METODO', 'RFC TERCERO', 'UUID', 'ESTATUS'];
  return escribirTabla_(sheet, filaInicial, 'FACTURAS RECIBIDAS', encabezados, filas,
    14, 4, 10, 11, 14, '#FCE4D6');
}

function escribirTabla_(sheet, filaInicial, titulo, encabezados, filas, columnas,
    primeraMoneda, ultimaMoneda, columnaTipo, columnaEstatus, color) {
  var fila = filaInicial;
  sheet.getRange(fila, 1).setValue(titulo).setFontWeight('bold');
  fila++;
  var filaEncabezado = fila;
  sheet.getRange(fila, 1, 1, columnas).setValues([encabezados])
    .setFontWeight('bold').setHorizontalAlignment('center').setBackground(color);
  fila++;
  var inicioDatos = fila;

  if (filas.length) {
    sheet.getRange(fila, 1, filas.length, columnas).setValues(filas);
    sheet.getRange(fila, columnaTipo, filas.length, 3).setNumberFormat('@');
    sheet.getRange(fila, columnaTipo, filas.length, 2).setHorizontalAlignment('center');
    sheet.getRange(fila, columnaEstatus, filas.length, 1).setHorizontalAlignment('center');
    aplicarColoresClasificacion_(sheet, fila, filas, columnas, columnaTipo, columnaEstatus);
    fila += filas.length;
  } else {
    var vacia = new Array(columnas).fill('');
    vacia[2] = 'Sin CFDI aplicables';
    sheet.getRange(fila, 1, 1, columnas).setValues([vacia]);
    fila++;
  }
  var finDatos = fila - 1;

  if (filas.length) {
    var filaTotales = new Array(columnas).fill('');
    filaTotales[2] = titulo === 'FACTURAS EMITIDAS' ? 'TOTAL EMITIDAS' : 'TOTAL RECIBIDAS';
    sheet.getRange(fila, 1, 1, columnas).setValues([filaTotales]);
    var cantidadMonedas = ultimaMoneda - primeraMoneda + 1;
    var formulas = [];
    for (var c = primeraMoneda; c <= ultimaMoneda; c++) {
      var letra = columnaALetra_(c);
      formulas.push('=SUM(' + letra + inicioDatos + ':' + letra + finDatos + ')');
    }
    sheet.getRange(fila, primeraMoneda, 1, cantidadMonedas).setFormulas([formulas]);
    sheet.getRange(fila, 1, 1, columnas).setFontWeight('bold').setBackground('#FFF2CC');
    var filaSuma = fila;
    fila++;

    sheet.getRange(fila, 3).setValue('TOTAL CALCULADO');
    if (ultimaMoneda === 7) {
      sheet.getRange(fila, 7).setFormula('=D' + filaSuma + '+E' + filaSuma + '+F' + filaSuma);
    } else {
      sheet.getRange(fila, 10).setFormula('=G' + filaSuma + '+H' + filaSuma + '+I' + filaSuma);
    }
    fila++;
    sheet.getRange(fila, 3).setValue('DIFERENCIA');
    var totalLetra = columnaALetra_(ultimaMoneda);
    sheet.getRange(fila, ultimaMoneda).setFormula('=' + totalLetra + filaSuma + '-' + totalLetra + (fila - 1));
    sheet.getRange(fila, 3, 1, columnas - 2).setFontWeight('bold');
  }

  sheet.getRange(inicioDatos, primeraMoneda, Math.max(1, fila - inicioDatos + 1), ultimaMoneda - primeraMoneda + 1)
    .setNumberFormat(CONFIG_CEDULAS.FORMATO_MONEDA);
  sheet.getRange(filaEncabezado, 1, fila - filaEncabezado + 1, columnas)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  return fila + 1;
}

function construirEncabezado_(sheet, cliente, periodo) {
  sheet.getRange('A1:N1').merge().setValue(cliente.nombre || 'NOMBRE DEL CLIENTE');
  sheet.getRange('A2:N2').merge().setValue('RFC: ' + (cliente.rfc || 'XAXX010101000'));
  sheet.getRange('A3:N3').merge().setValue('REPORTE FINANCIERO - ' + periodo + ' - VERSION 3.4');
  sheet.getRange('A1:N3').setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange('J4:N4').setValues([['PUE', 'PPD', 'COMPLEMENTO', 'CANCELADA', 'REVISAR']])
    .setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange('J4').setBackground('#FFFFFF');
  sheet.getRange('K4').setBackground('#FFF2CC');
  sheet.getRange('L4').setBackground('#FFE699');
  sheet.getRange('M4').setBackground('#EA9999');
  sheet.getRange('N4').setBackground('#F4CCCC');
  return 5;
}

function aplicarPresentacionGeneral_(sheet) {
  sheet.setFrozenRows(3);
  sheet.getDataRange().setVerticalAlignment('middle');
  sheet.setColumnWidth(1, 95);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 300);
  for (var c = 4; c <= 10; c++) sheet.setColumnWidth(c, 110);
  sheet.setColumnWidth(8, 130);
  sheet.setColumnWidth(9, 150);
  sheet.setColumnWidth(10, 220);
  sheet.setColumnWidth(11, 130);
  sheet.setColumnWidth(12, 150);
  sheet.setColumnWidth(13, 280);
  sheet.setColumnWidth(14, 120);
}

function aplicarColoresClasificacion_(sheet, filaInicial, filas, columnas, columnaTipo, columnaEstatus) {
  var colores = {
    'PUE': '#FFFFFF',
    'PPD': '#FFF2CC',
    'COMPLEMENTO': '#FFE699',
    'CANCELADA': '#EA9999',
    'REVISAR': '#F4CCCC'
  };
  for (var i = 0; i < filas.length; i++) {
    var tipo = String(filas[i][columnaTipo - 1] || 'REVISAR').toUpperCase();
    var estatus = String(filas[i][columnaEstatus - 1] || 'REVISAR').toUpperCase();
    var color = estatus === 'CANCELADA' || estatus === 'REVISAR' ? colores[estatus] :
      (colores[tipo] || colores.REVISAR);
    sheet.getRange(filaInicial + i, 1, 1, columnas).setBackground(color);
    sheet.getRange(filaInicial + i, columnaTipo).setFontWeight('bold');
    sheet.getRange(filaInicial + i, columnaEstatus).setFontWeight('bold');
  }
}

function registroExcluido_(uuid, factura, fechaIso, clienteInfo, advertencias) {
  return {
    incluir: false,
    fila: null,
    fechaOrden: fechaAOrden_(fechaIso),
    factura: factura,
    uuid: uuid,
    clienteInfo: clienteInfo,
    advertencias: advertencias
  };
}

function crearRegistroCancelado_(tipo, tipoComprobante, fechaIso, factura, uuid,
    clienteInfo, nombreEmisor, nombreReceptor, rfcEmisor, rfcReceptor, metodoPago, advertencias) {
  var fecha = formatearFecha_(fechaIso);
  var clasificacion = tipoComprobante === 'P' ? 'COMPLEMENTO' : clasificarMetodo_(metodoPago);
  var fila;
  if (tipo === 'emitida') {
    fila = [fecha, factura, nombreReceptor, 0, 0, 0, 0, clasificacion,
      rfcReceptor, uuid, 'CANCELADA'];
  } else {
    fila = [fecha, factura, nombreEmisor, 0, 0, 0, 0, 0, 0, 0, clasificacion,
      rfcEmisor, uuid, 'CANCELADA'];
  }
  return {
    incluir: true,
    fila: fila,
    fechaOrden: fechaAOrden_(fechaIso),
    factura: factura,
    uuid: uuid,
    clienteInfo: clienteInfo,
    advertencias: advertencias
  };
}

function determinarEstatus_(advertencias) {
  var requierenRevision = {
    'TASA NO SOPORTADA': true,
    'EMITIDA IVA 8%': true,
    'PAGO EMITIDO IVA 8%': true,
    'PAGO SIN RELACION': true
  };
  for (var i = 0; i < advertencias.length; i++) {
    if (requierenRevision[advertencias[i][0]]) return 'REVISAR';
  }
  return 'VIGENTE';
}

function validarClienteUnico_(emitidas, recibidas) {
  var clientes = {};
  var fuentes = [emitidas.clientesEncontrados || {}, recibidas.clientesEncontrados || {}];
  for (var i = 0; i < fuentes.length; i++) {
    var rfcs = Object.keys(fuentes[i]);
    for (var j = 0; j < rfcs.length; j++) clientes[rfcs[j]] = fuentes[i][rfcs[j]];
  }
  var encontrados = Object.keys(clientes);
  if (encontrados.length > 1) {
    throw new Error('Las carpetas contienen CFDI de varios clientes: ' + encontrados.join(', ') +
      '. Separe los XML por RFC antes de ejecutar Lya.');
  }
  if (encontrados.length === 1) return { rfc: encontrados[0], nombre: clientes[encontrados[0]] };
  return elegirCliente_(emitidas.clienteInfo, recibidas.clienteInfo);
}

function detectarPeriodo_(filas) {
  var meses = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
    'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
  var periodos = {};
  for (var i = 0; i < filas.length; i++) {
    var partes = String(filas[i][0] || '').split('/');
    if (partes.length !== 3) continue;
    var mes = Number(partes[1]);
    var anio = partes[2];
    if (mes >= 1 && mes <= 12) periodos[anio + '-' + partes[1]] = meses[mes - 1] + ' ' + anio;
  }
  var claves = Object.keys(periodos).sort();
  if (!claves.length) return 'PERIODO NO DETECTADO';
  if (claves.length === 1) return periodos[claves[0]];
  return 'VARIOS PERIODOS (' + claves.map(function (clave) { return periodos[clave]; }).join(', ') + ')';
}

function limpiarIdCarpeta_(valor) {
  var texto = String(valor || '').trim();
  var coincidencia = texto.match(/\/folders\/([^\/?#]+)/i);
  if (coincidencia) return coincidencia[1];
  texto = texto.split('?')[0].split('#')[0];
  var partes = texto.split('/');
  return partes[partes.length - 1].trim();
}

function elegirCliente_(a, b) {
  if (a && a.nombre) return a;
  if (b && b.nombre) return b;
  return { nombre: '', rfc: '' };
}

function atributoTexto_(elemento, nombre, predeterminado) {
  if (!elemento) return predeterminado;
  var atributo = elemento.getAttribute(nombre);
  return atributo ? atributo.getValue() : predeterminado;
}

function atributoNumero_(elemento, nombre) {
  var numero = Number(atributoTexto_(elemento, nombre, '0'));
  return isFinite(numero) ? numero : 0;
}

function tasasIguales_(a, b) {
  return Math.abs(a - b) < 0.000001;
}

function normalizarCero_(numero) {
  return Math.abs(numero) < 0.000000001 ? 0 : numero;
}

function formatearFecha_(fechaIso) {
  if (!fechaIso) return '';
  var partes = fechaIso.substring(0, 10).split('-');
  return partes.length === 3 ? partes[2] + '/' + partes[1] + '/' + partes[0] : fechaIso;
}

function fechaAOrden_(fechaIso) {
  var tiempo = fechaIso ? new Date(fechaIso).getTime() : 0;
  return isNaN(tiempo) ? 0 : tiempo;
}

function obtenerUuid_(root) {
  var timbre = buscarElemento_(root, 'TimbreFiscalDigital');
  return timbre ? atributoTexto_(timbre, 'UUID', '') : '';
}

function buscarElemento_(elemento, nombreBuscado) {
  if (elemento.getName() === nombreBuscado) return elemento;
  var hijos = elemento.getChildren();
  for (var i = 0; i < hijos.length; i++) {
    var encontrado = buscarElemento_(hijos[i], nombreBuscado);
    if (encontrado) return encontrado;
  }
  return null;
}

function buscarElementos_(elemento, nombreBuscado) {
  var resultados = [];
  if (elemento.getName() === nombreBuscado) resultados.push(elemento);
  var hijos = elemento.getChildren();
  for (var i = 0; i < hijos.length; i++) {
    resultados = resultados.concat(buscarElementos_(hijos[i], nombreBuscado));
  }
  return resultados;
}

function normalizarIdentificador_(valor) {
  return String(valor || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function clasificarMetodo_(metodoPago) {
  var metodo = String(metodoPago || '').toUpperCase();
  if (metodo === 'PUE' || metodo === 'PPD') return metodo;
  return 'REVISAR';
}

function huellaXml_(contenidoXml) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, contenidoXml, Utilities.Charset.UTF_8);
  return digest.map(function (byte) {
    var valor = byte < 0 ? byte + 256 : byte;
    return ('0' + valor.toString(16)).slice(-2);
  }).join('');
}

function columnaALetra_(columna) {
  var resultado = '';
  while (columna > 0) {
    var residuo = (columna - 1) % 26;
    resultado = String.fromCharCode(65 + residuo) + resultado;
    columna = Math.floor((columna - 1) / 26);
  }
  return resultado;
}