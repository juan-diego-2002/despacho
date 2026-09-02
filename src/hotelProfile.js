function redondear(valor) {
  return Math.round((Number(valor) || 0) * 100) / 100;
}

function calcularEmitida(fila, tasaIva) {
  const hospedaje = Number(fila.hospedaje) || 0;
  const alimentos = Number(fila.alimentos) || 0;
  const otros = Number(fila.otros) || 0;
  const subtotal = redondear(hospedaje + alimentos + otros);
  const impuestoHospedaje = redondear(hospedaje * 0.02);
  const iva = redondear(subtotal * tasaIva);
  return { ...fila, hospedaje, alimentos, otros, subtotal, impuestoHospedaje, iva, total: redondear(subtotal + impuestoHospedaje + iva) };
}

function calcularRecibida(fila) {
  const caso = String(fila.caso || '1');
  const pagoNeto = Number(fila.pagoNeto) || 0;
  let base8 = Number(fila.base8) || 0;
  let base16 = Number(fila.base16) || 0;
  let tasa0 = Number(fila.tasa0) || 0;
  let exento = Number(fila.exento) || 0;
  let iva8 = Number(fila.iva8) || 0;
  let iva16 = Number(fila.iva16) || 0;
  let retencionIsr = Number(fila.retencionIsr) || 0;
  const retencionIva = Number(fila.retencionIva) || 0;

  if (caso === '1' && !fila.desdeXml) {
    base8 = redondear(pagoNeto / 1.08);
    iva8 = redondear(base8 * 0.08);
    base16 = iva16 = tasa0 = exento = 0;
  } else if (caso === '2' && !fila.desdeXml) {
    base16 = redondear(pagoNeto / 1.16);
    iva16 = redondear(base16 * 0.16);
    base8 = iva8 = tasa0 = exento = 0;
  } else if (caso === '4' && !fila.desdeXml) {
    const restante = Math.max(pagoNeto - base8 - iva8, 0);
    base16 = redondear(restante / 1.16);
    iva16 = redondear(base16 * 0.16);
    tasa0 = exento = 0;
  }

  const subtotal = redondear(base8 + base16 + tasa0 + exento);
  if (caso === '5') retencionIsr = redondear(subtotal * 0.0125);
  const ieps = Number(fila.ieps) || 0;
  const sinIeps = redondear(subtotal + iva8 + iva16 - retencionIsr - retencionIva);
  const conIeps = redondear(sinIeps + ieps);
  // El IEPS puede venir fuera del subtotal (PepsiCo/combustibles) o ya incluido
  // en la base de IVA (algunos CFDI mixtos). Se usa la conciliación más cercana
  // al Total sellado, sin ocultar el IEPS en su columna propia.
  const iepsSeSuma = Math.abs(pagoNeto - conIeps) < Math.abs(pagoNeto - sinIeps);
  const netoCalculado = iepsSeSuma ? conIeps : sinIeps;
  return { ...fila, caso, pagoNeto, base8, base16, tasa0, exento, iva8, iva16, ieps, subtotal, retencionIsr, retencionIva, iepsSeSuma, netoCalculado, diferencia: redondear(pagoNeto - netoCalculado) };
}

function crearPerfilHoteleria(resultado, opciones = {}) {
  const tasaIva = Number(opciones.ivaHotel) === 16 ? 0.16 : 0.08;
  const registrosEmitidos = resultado.emitidas?.registros || [];
  const emitidas = (resultado.emitidas?.filas || []).map((fila, index) => {
    const registro = registrosEmitidos[index] || {};
    const importes = registro.hotelEmitida || {};
    const advertenciasReales = (registro.advertencias || []).filter(advertencia =>
      !['EMITIDA IVA 8%', 'SUBTOTAL DIFERENTE'].includes(String(advertencia?.[0] || ''))
    );
    const estatusOriginal = String(fila[10] || 'VIGENTE').toUpperCase();
    const estatus = estatusOriginal === 'CANCELADA'
      ? 'CANCELADA'
      : (advertenciasReales.length ? 'REVISAR' : 'VIGENTE');
    return calcularEmitida({
      fecha: fila[0], factura: fila[1], nombre: fila[2],
      hospedaje: Number(importes.hospedaje) || 0,
      alimentos: Number(importes.alimentos) || 0,
      otros: Number(importes.otros) || 0,
      totalReferencia: Number(fila[6]) || 0, uuid: fila[9], estatus
    }, tasaIva);
  });
  const registrosRecibidos = resultado.recibidas?.registros || [];
  const recibidas = (resultado.recibidas?.filas || []).map((fila, index) => {
    const metodo = String(fila[11] || '').toUpperCase();
    const xml = registrosRecibidos[index]?.hotelRecibida || {};
    const esComplemento = metodo === 'COMPLEMENTO';
    let base8Xml = esComplemento ? 0 : Number(xml.base8) || 0;
    const base16Xml = esComplemento ? 0 : Number(xml.base16) || 0;
    let tasa0Xml = esComplemento ? 0 : Number(xml.tasa0) || 0;
    const exentoXml = esComplemento ? 0 : Number(xml.exento) || 0;
    let iva8Xml = esComplemento ? 0 : Number(xml.iva8) || 0;
    const iva16Xml = esComplemento ? 0 : Number(xml.iva16) || 0;
    let iepsXml = esComplemento ? 0 : Number(xml.ieps) || 0;
    const descuentoXml = esComplemento ? 0 : Number(xml.descuento) || 0;
    const retIsrXml = esComplemento ? 0 : Number(xml.retencionIsr) || 0;
    const retIvaXml = esComplemento ? 0 : Number(xml.retencionIva) || 0;
    let caso = 'REVISAR';
    if (esComplemento) caso = '0';
    else if (iepsXml) caso = 'IEPS';
    else if (exentoXml && retIsrXml && !base8Xml && !base16Xml && !tasa0Xml) caso = '5';
    else if (base8Xml && base16Xml && !tasa0Xml && !exentoXml && !iepsXml && !descuentoXml) caso = '4';
    else if (base8Xml && !base16Xml && !tasa0Xml && !exentoXml && !iepsXml && !descuentoXml) caso = '1';
    else if (base16Xml && !base8Xml && !tasa0Xml && !exentoXml && !iepsXml && !descuentoXml) caso = '2';
    else if (!base8Xml && !base16Xml && (tasa0Xml || exentoXml || retIsrXml || retIvaXml) && !iepsXml && !descuentoXml) caso = '3';
    return calcularRecibida({
      fecha: fila[0], factura: fila[1], nombre: fila[2], rfc: fila[12], caso,
      pagoNeto: esComplemento ? 0 : Number(xml.pagoNeto) || 0,
      base8: base8Xml, base16: base16Xml, tasa0: tasa0Xml, exento: exentoXml,
      iva8: iva8Xml, iva16: iva16Xml, ieps: iepsXml,
      retencionIsr: retIsrXml, retencionIva: retIvaXml,
      uuid: fila[13], estatus: fila[14], tipoMetodo: fila[11], desdeXml: true
    });
  });
  return { tasaIva, emitidas, recibidas };
}

module.exports = { crearPerfilHoteleria, calcularEmitida, calcularRecibida };
