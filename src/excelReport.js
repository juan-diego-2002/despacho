const ExcelJS = require('exceljs');

const COLORES_ESTADO = {
    PPD: 'FCDB57',
    COMPLEMENTO: '9FB0FF',
    CANCELADA: 'FC3030',
    REVISAR: 'F4CCCC'
};

/**
 * ============================================================
 * CREAR EXCEL
 * ============================================================
 *
 * Genera:
 *
 * MES:
 *   RESUMEN
 *   JUNIO
 *
 * TRIMESTRAL:
 *   RESUMEN
 *   ENERO
 *   FEBRERO
 *   MARZO
 *
 * ANUAL:
 *   RESUMEN
 *   ENERO
 *   FEBRERO
 *   MARZO
 *   ...
 *   DICIEMBRE
 *
 */
async function crearExcel(reporte) {

    const workbook = new ExcelJS.Workbook();

    workbook.creator = 'CFDI Web';
    workbook.created = new Date();

    // =========================================================
    // 1. MESES
    // =========================================================

    const meses = [
        'ENERO',
        'FEBRERO',
        'MARZO',
        'ABRIL',
        'MAYO',
        'JUNIO',
        'JULIO',
        'AGOSTO',
        'SEPTIEMBRE',
        'OCTUBRE',
        'NOVIEMBRE',
        'DICIEMBRE'
    ];

    // =========================================================
    // 2. DETERMINAR TIPO DE PERIODO
    // =========================================================

    const tipoPeriodo = String(
        reporte.tipoPeriodo || obtenerTipoDesdeFiltro(reporte.filtro) || 'MES'
    ).toUpperCase();

    let mesesSeleccionados = [];

    // =========================================================
    // 3. DETERMINAR MESES
    // =========================================================

    if (tipoPeriodo === 'MES') {

        let mes = Number(reporte.mes || 0);

        // Si viene como nombre
        if (!mes) {

            const nombreMes = String(
                reporte.mes ||
                reporte.periodo ||
                ''
            ).toUpperCase();

            const indice = meses.indexOf(nombreMes);

            if (indice !== -1) {
                mes = indice + 1;
            }
        }

        if (mes >= 1 && mes <= 12) {
            mesesSeleccionados = [mes - 1];
        }

    } else if (tipoPeriodo === 'TRIMESTRAL') {

        const trimestre = Number(
            reporte.trimestre ||
            reporte.filtro?.trimestre ||
            1
        );

        const inicio = (trimestre - 1) * 3;

        mesesSeleccionados = [
            inicio,
            inicio + 1,
            inicio + 2
        ];

    } else if (tipoPeriodo === 'ANUAL') {

        mesesSeleccionados = meses.map(
            (_, index) => index
        );

    }

    // Si no se detectó correctamente
    if (mesesSeleccionados.length === 0) {

        // Intentar obtener los meses directamente
        // de las facturas existentes.
        const emitidas = reporte.emitidas?.filas || [];
        const recibidas = reporte.recibidas?.filas || [];

        const mesesEncontrados = new Set();

        for (const fila of [...emitidas, ...recibidas]) {

            const fecha = obtenerFechaFila(fila);

            if (fecha) {
                mesesEncontrados.add(fecha.getMonth());
            }
        }

        mesesSeleccionados = Array.from(
            mesesEncontrados
        ).sort((a, b) => a - b);
    }

    // Si definitivamente no hay fechas
    if (mesesSeleccionados.length === 0) {
        mesesSeleccionados = [0];
    }

    // =========================================================
    // 4. OBTENER FILAS
    // =========================================================

    const filasEmitidas =
        reporte.emitidas?.filas || [];

    const filasRecibidas =
        reporte.recibidas?.filas || [];

    // =========================================================
    // 5. SEPARAR POR MES
    // =========================================================

    const emitidasPorMes =
        separarPorMes(filasEmitidas);

    const recibidasPorMes =
        separarPorMes(filasRecibidas);

    // =========================================================
    // 6. CREAR RESUMEN
    // =========================================================

    crearHojaResumen(
        workbook,
        reporte,
        meses,
        mesesSeleccionados,
        emitidasPorMes,
        recibidasPorMes
    );

    // =========================================================
    // 7. CREAR UNA HOJA POR MES
    // =========================================================

    for (const indiceMes of mesesSeleccionados) {

        const nombreMes =
            meses[indiceMes];

        const datosEmitidas =
            emitidasPorMes[indiceMes] || [];

        const datosRecibidas =
            recibidasPorMes[indiceMes] || [];

        crearHojaMes(
            workbook,
            nombreMes,
            reporte,
            datosEmitidas,
            datosRecibidas
        );
    }

    // =========================================================
    // 8. GENERAR EXCEL
    // =========================================================

    return await workbook.xlsx.writeBuffer();
}


/**
 * ============================================================
 * OBTENER TIPO DESDE FILTRO
 * ============================================================
 */
function obtenerTipoDesdeFiltro(filtro) {

    if (!filtro) {
        return 'MES';
    }

    const modo =
        String(filtro.modo || '').toLowerCase();

    if (modo === 'anual') {
        return 'ANUAL';
    }

    if (modo === 'trimestral') {
        return 'TRIMESTRAL';
    }

    return 'MES';
}


/**
 * ============================================================
 * SEPARAR FILAS POR MES
 * ============================================================
 */
function separarPorMes(filas) {

    const resultado = {};

    for (let i = 0; i < 12; i++) {
        resultado[i] = [];
    }

    for (const fila of filas) {

        const fecha =
            obtenerFechaFila(fila);

        if (!fecha) {
            continue;
        }

        const mes =
            fecha.getMonth();

        resultado[mes].push(fila);
    }

    return resultado;
}


/**
 * ============================================================
 * OBTENER FECHA DE LA FILA
 * ============================================================
 *
 * IMPORTANTE:
 *
 * La fecha de tus CFDI está en la primera columna:
 *
 * [0] FECHA
 *
 * Ejemplo:
 *
 * 15/06/2026
 *
 */
function obtenerFechaFila(fila) {

    if (!fila || !fila.length) {
        return null;
    }

    // ---------------------------------------------------------
    // PRIMERO: TOMAR EXCLUSIVAMENTE LA COLUMNA FECHA
    // ---------------------------------------------------------

    const valor =
        fila[0];

    if (valor instanceof Date) {

        if (!isNaN(valor.getTime())) {
            return new Date(
                valor.getFullYear(),
                valor.getMonth(),
                valor.getDate()
            );
        }
    }

    if (typeof valor !== 'string') {
        return null;
    }

    const texto =
        valor.trim();

    if (!texto) {
        return null;
    }

    // ---------------------------------------------------------
    // DD/MM/YYYY
    // ---------------------------------------------------------

    let match =
        texto.match(
            /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
        );

    if (match) {

        const dia =
            Number(match[1]);

        const mes =
            Number(match[2]);

        const anio =
            Number(match[3]);

        const fecha =
            new Date(
                anio,
                mes - 1,
                dia
            );

        if (
            fecha.getFullYear() === anio &&
            fecha.getMonth() === mes - 1 &&
            fecha.getDate() === dia
        ) {
            return fecha;
        }
    }

    // ---------------------------------------------------------
    // DD-MM-YYYY
    // ---------------------------------------------------------

    match =
        texto.match(
            /^(\d{1,2})-(\d{1,2})-(\d{4})$/
        );

    if (match) {

        const dia =
            Number(match[1]);

        const mes =
            Number(match[2]);

        const anio =
            Number(match[3]);

        const fecha =
            new Date(
                anio,
                mes - 1,
                dia
            );

        if (
            fecha.getFullYear() === anio &&
            fecha.getMonth() === mes - 1 &&
            fecha.getDate() === dia
        ) {
            return fecha;
        }
    }

    // ---------------------------------------------------------
    // YYYY-MM-DD
    // ---------------------------------------------------------

    match =
        texto.match(
            /^(\d{4})-(\d{1,2})-(\d{1,2})$/
        );

    if (match) {

        const anio =
            Number(match[1]);

        const mes =
            Number(match[2]);

        const dia =
            Number(match[3]);

        const fecha =
            new Date(
                anio,
                mes - 1,
                dia
            );

        if (
            fecha.getFullYear() === anio &&
            fecha.getMonth() === mes - 1 &&
            fecha.getDate() === dia
        ) {
            return fecha;
        }
    }

    // ---------------------------------------------------------
    // FECHA ISO CON HORA
    // Ejemplo:
    // 2026-06-15T10:30:00
    // ---------------------------------------------------------

    if (
        /^\d{4}-\d{2}-\d{2}T/.test(texto)
    ) {

        const partes =
            texto.substring(0, 10)
                .split('-');

        const anio =
            Number(partes[0]);

        const mes =
            Number(partes[1]);

        const dia =
            Number(partes[2]);

        const fecha =
            new Date(
                anio,
                mes - 1,
                dia
            );

        if (!isNaN(fecha.getTime())) {
            return fecha;
        }
    }

    return null;
}


/**
 * ============================================================
 * CREAR HOJA DEL MES
 * ============================================================
 */
function crearHojaMes(
    workbook,
    nombreMes,
    reporte,
    filasEmitidas,
    filasRecibidas
) {

    const hoja =
        workbook.addWorksheet(
            nombreMes.substring(0, 31)
        );

    // =========================================================
    // ENCABEZADO
    // =========================================================

    hoja.mergeCells('A1:N1');

    hoja.getCell('A1').value =
        reporte.cliente?.nombre ||
        'NOMBRE DEL CLIENTE';

    hoja.getCell('A1').font = {
        bold: true,
        size: 16
    };

    hoja.getCell('A1').alignment = {
        horizontal: 'center'
    };

    hoja.mergeCells('A2:N2');

    hoja.getCell('A2').value =
        `RFC: ${
            reporte.cliente?.rfc ||
            'XAXX010101000'
        }`;

    hoja.getCell('A2').font = {
        bold: true,
        size: 12
    };

    hoja.getCell('A2').alignment = {
        horizontal: 'center'
    };

    hoja.mergeCells('A3:N3');

    hoja.getCell('A3').value =
        `REPORTE FINANCIERO - ${nombreMes}`;

    hoja.getCell('A3').font = {
        bold: true,
        size: 13
    };

    hoja.getCell('A3').alignment = {
        horizontal: 'center'
    };

    // =========================================================
    // INFORMACIÓN DEL MES
    // =========================================================

    hoja.mergeCells('A4:N4');

    hoja.getCell('A4').value =
        `CFDI emitidas: ${filasEmitidas.length} | ` +
        `CFDI recibidas: ${filasRecibidas.length}`;

    hoja.getCell('A4').alignment = {
        horizontal: 'center'
    };

    hoja.getCell('A4').font = {
        italic: true
    };

    // =========================================================
    // FACTURAS EMITIDAS
    // =========================================================

    let filaActual = 6;

    filaActual =
        agregarTabla(
            hoja,
            filaActual,
            'FACTURAS EMITIDAS',
            reporte.emitidas?.encabezados || [],
            filasEmitidas,
            'emitidas'
        );

    // =========================================================
    // ESPACIO
    // =========================================================

    filaActual += 2;

    // =========================================================
    // FACTURAS RECIBIDAS
    // =========================================================

    agregarTabla(
        hoja,
        filaActual,
        'FACTURAS RECIBIDAS',
        reporte.recibidas?.encabezados || [],
        filasRecibidas,
        'recibidas'
    );

    // =========================================================
    // AJUSTAR COLUMNAS
    // =========================================================

    hoja.columns.forEach(column => {

        let maxLength = 10;

        column.eachCell(
            {
                includeEmpty: false
            },
            cell => {

                let valor = '';

                if (
                    cell.value instanceof Date
                ) {
                    valor =
                        formatearFechaExcel(
                            cell.value
                        );
                } else {
                    valor =
                        String(
                            cell.value ?? ''
                        );
                }

                if (
                    valor.length > maxLength
                ) {
                    maxLength =
                        valor.length;
                }
            }
        );

        column.width =
            Math.min(
                maxLength + 2,
                35
            );
    });

    // =========================================================
    // CONGELAR ENCABEZADOS
    // =========================================================

    hoja.views = [
        {
            state: 'frozen',
            ySplit: 6
        }
    ];

    // =========================================================
    // FILTRO
    // =========================================================

    hoja.autoFilter = {
        from: 'A7',
        to: {
            row: Math.max(
                hoja.rowCount,
                7
            ),
            column:
                Math.max(
                    reporte.recibidas?.encabezados?.length || 0,
                    reporte.emitidas?.encabezados?.length || 0
                )
        }
    };
}


/**
 * ============================================================
 * AGREGAR TABLA
 * ============================================================
 */
function agregarTabla(
    hoja,
    filaInicial,
    titulo,
    encabezados,
    filas,
    tipo
) {

    let fila =
        filaInicial;

    const cantidadColumnas =
        Math.max(
            encabezados.length,
            1
        );

    // =========================================================
    // TÍTULO
    // =========================================================

    hoja.mergeCells(
        fila,
        1,
        fila,
        cantidadColumnas
    );

    const tituloCell =
        hoja.getCell(
            fila,
            1
        );

    tituloCell.value =
        titulo;

    tituloCell.font = {
        bold: true,
        size: 12
    };

    tituloCell.alignment = {
        horizontal: 'center'
    };

    tituloCell.border =
        crearBorde();

    fila++;

    // =========================================================
    // ENCABEZADOS
    // =========================================================

    encabezados.forEach(
        (header, index) => {

            const cell =
                hoja.getCell(
                    fila,
                    index + 1
                );

            cell.value =
                header;

            cell.font = {
                bold: true
            };

            cell.alignment = {
                horizontal: 'center',
                vertical: 'middle',
                wrapText: true
            };

            cell.border =
                crearBorde();

            if (tipo === 'emitidas') {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: {
                        argb: 'D9EAF7'
                    }
                };
            }

            if (tipo === 'recibidas') {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: {
                        argb: 'FCE4D6'
                    }
                };
            }
        }
    );

    fila++;

    // =========================================================
    // SIN DATOS
    // =========================================================

    if (!filas.length) {

        hoja.mergeCells(
            fila,
            1,
            fila,
            cantidadColumnas
        );

        const cell =
            hoja.getCell(
                fila,
                1
            );

        cell.value =
            'Sin CFDI aplicables';

        cell.alignment = {
            horizontal: 'center'
        };

        cell.border =
            crearBorde();

        return fila + 1;
    }

    // =========================================================
    // DATOS
    // =========================================================

    for (const datos of filas) {

        const estado = obtenerEstadoFila(datos, tipo);

        datos.forEach(
            (valor, index) => {

                const cell =
                    hoja.getCell(
                        fila,
                        index + 1
                    );

                cell.value =
                    convertirValor(valor);

                cell.border =
                    crearBorde();

                cell.alignment = {
                    vertical: 'middle'
                };

                // -------------------------------------------------
                // FECHA
                // -------------------------------------------------

                if (index === 0) {

                    const fecha =
                        convertirFecha(valor);

                    if (fecha) {

                        cell.value =
                            fecha;

                        cell.numFmt =
                            'dd/mm/yyyy';
                    }
                }

                // -------------------------------------------------
                // MONEDA EMITIDAS
                // -------------------------------------------------

                if (
                    tipo === 'emitidas' &&
                    index >= 3 &&
                    index <= 6
                ) {

                    cell.numFmt =
                        '$#,##0.00';

                    cell.alignment = {
                        horizontal: 'right'
                    };
                }

                // -------------------------------------------------
                // MONEDA RECIBIDAS
                // -------------------------------------------------

                if (
                    tipo === 'recibidas' &&
                    index >= 3 &&
                    index <= 10
                ) {

                    cell.numFmt =
                        '$#,##0.00';

                    cell.alignment = {
                        horizontal: 'right'
                    };
                }

                // -------------------------------------------------
                // ESTATUS
                // -------------------------------------------------

                if (estado) {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: {
                            argb: COLORES_ESTADO[estado]
                        }
                    };
                }
            }
        );

        fila++;
    }

    // =========================================================
    // TOTAL DEL MES
    // =========================================================

    const totales =
        calcularTotalesTabla(
            filas,
            tipo
        );

    encabezados.forEach(
        (_, index) => {

            const cell =
                hoja.getCell(
                    fila,
                    index + 1
                );

            if (index === 2) {

                cell.value =
                    'TOTAL';
            }

            if (
                tipo === 'emitidas' &&
                index >= 3 &&
                index <= 6
            ) {

                cell.value =
                    Number(
                        totales[index] || 0
                    );

                cell.numFmt =
                    '$#,##0.00';
            }

            if (
                tipo === 'recibidas' &&
                index >= 3 &&
                index <= 10
            ) {

                cell.value =
                    Number(
                        totales[index] || 0
                    );

                cell.numFmt =
                    '$#,##0.00';
            }

            cell.font = {
                bold: true
            };

            cell.border =
                crearBorde();

            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: {
                    argb: 'FFF2CC'
                }
            };
        }
    );

    fila++;

    return fila;
}


function obtenerEstadoFila(fila, tipo) {

    const estado =
        fila[fila.length - 1] === 'REVISAR' ||
        fila[fila.length - 1] === 'CANCELADA'
            ? fila[fila.length - 1]
            : fila[tipo === 'emitidas' ? 7 : 11];

    const texto =
        String(estado ?? '').toUpperCase();

    return COLORES_ESTADO[texto]
        ? texto
        : null;
}


/**
 * ============================================================
 * CALCULAR TOTALES DE LA TABLA
 * ============================================================
 */
function calcularTotalesTabla(
    filas,
    tipo
) {

    const totales = {};

    const inicio =
        tipo === 'emitidas'
            ? 3
            : 3;

    const fin =
        tipo === 'emitidas'
            ? 6
            : 10;

    for (
        let i = inicio;
        i <= fin;
        i++
    ) {

        totales[i] =
            filas.reduce(
                (suma, fila) => {

                    return suma +
                        convertirNumero(
                            fila[i]
                        );

                },
                0
            );
    }

    return totales;
}


/**
 * ============================================================
 * CREAR HOJA RESUMEN
 * ============================================================
 */
function crearHojaResumen(
    workbook,
    reporte,
    meses,
    mesesSeleccionados,
    emitidasPorMes,
    recibidasPorMes
) {

    const hoja =
        workbook.addWorksheet(
            'RESUMEN'
        );

    // =========================================================
    // ENCABEZADO
    // =========================================================

    hoja.mergeCells('A1:F1');

    hoja.getCell('A1').value =
        reporte.cliente?.nombre ||
        'NOMBRE DEL CLIENTE';

    hoja.getCell('A1').font = {
        bold: true,
        size: 16
    };

    hoja.getCell('A1').alignment = {
        horizontal: 'center'
    };

    hoja.mergeCells('A2:F2');

    hoja.getCell('A2').value =
        `RFC: ${
            reporte.cliente?.rfc ||
            'XAXX010101000'
        }`;

    hoja.getCell('A2').alignment = {
        horizontal: 'center'
    };

    hoja.mergeCells('A3:F3');

    hoja.getCell('A3').value =
        `REPORTE RESUMEN - ${
            reporte.periodo || ''
        }`;

    hoja.getCell('A3').font = {
        bold: true,
        size: 13
    };

    hoja.getCell('A3').alignment = {
        horizontal: 'center'
    };

    // =========================================================
    // ENCABEZADOS
    // =========================================================

    const encabezados = [
        'MES',
        'CFDI EMITIDAS',
        'CFDI RECIBIDAS',
        'TOTAL EMITIDAS',
        'TOTAL RECIBIDAS',
        'TOTAL GENERAL'
    ];

    const filaHeader = 5;

    encabezados.forEach(
        (header, index) => {

            const cell =
                hoja.getCell(
                    filaHeader,
                    index + 1
                );

            cell.value =
                header;

            cell.font = {
                bold: true
            };

            cell.alignment = {
                horizontal: 'center',
                vertical: 'middle'
            };

            cell.border =
                crearBorde();

            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: {
                    argb: 'D9EAF7'
                }
            };
        }
    );

    // =========================================================
    // DATOS
    // =========================================================

    let fila =
        6;

    let granTotalEmitidas =
        0;

    let granTotalRecibidas =
        0;

    let cantidadEmitidas =
        0;

    let cantidadRecibidas =
        0;

    for (
        const indice of mesesSeleccionados
    ) {

        const nombreMes =
            meses[indice];

        const emitidas =
            emitidasPorMes[indice] || [];

        const recibidas =
            recibidasPorMes[indice] || [];

        // -----------------------------------------------------
        // CALCULAR TOTAL REAL DEL MES
        // -----------------------------------------------------

        const totalEmitidas =
            calcularTotalFilas(
                emitidas,
                6
            );

        const totalRecibidas =
            calcularTotalFilas(
                recibidas,
                10
            );

        granTotalEmitidas +=
            totalEmitidas;

        granTotalRecibidas +=
            totalRecibidas;

        cantidadEmitidas +=
            emitidas.length;

        cantidadRecibidas +=
            recibidas.length;

        const datos = [
            nombreMes,
            emitidas.length,
            recibidas.length,
            totalEmitidas,
            totalRecibidas,
            totalEmitidas +
                totalRecibidas
        ];

        datos.forEach(
            (valor, index) => {

                const cell =
                    hoja.getCell(
                        fila,
                        index + 1
                    );

                cell.value =
                    valor;

                cell.border =
                    crearBorde();

                if (index >= 3) {

                    cell.numFmt =
                        '$#,##0.00';

                    cell.alignment = {
                        horizontal: 'right'
                    };
                }
            }
        );

        fila++;
    }

    // =========================================================
    // TOTAL GENERAL
    // =========================================================

    hoja.getCell(
        fila,
        1
    ).value =
        'TOTAL GENERAL';

    hoja.getCell(
        fila,
        2
    ).value =
        cantidadEmitidas;

    hoja.getCell(
        fila,
        3
    ).value =
        cantidadRecibidas;

    hoja.getCell(
        fila,
        4
    ).value =
        granTotalEmitidas;

    hoja.getCell(
        fila,
        5
    ).value =
        granTotalRecibidas;

    hoja.getCell(
        fila,
        6
    ).value =
        granTotalEmitidas +
        granTotalRecibidas;

    for (
        let i = 1;
        i <= 6;
        i++
    ) {

        const cell =
            hoja.getCell(
                fila,
                i
            );

        cell.font = {
            bold: true
        };

        cell.border =
            crearBorde();

        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'FFF2CC'
            }
        };

        if (i >= 4) {

            cell.numFmt =
                '$#,##0.00';
        }
    }

    // =========================================================
    // ANCHO
    // =========================================================

    hoja.columns.forEach(
        column => {

            column.width =
                20;
        }
    );

    hoja.views = [
        {
            state: 'frozen',
            ySplit: 5
        }
    ];
}


/**
 * ============================================================
 * CALCULAR TOTAL DE UNA COLUMNA
 * ============================================================
 */
function calcularTotalFilas(
    filas,
    indice
) {

    let total =
        0;

    for (
        const fila of filas
    ) {

        total +=
            convertirNumero(
                fila[indice]
            );
    }

    return total;
}


/**
 * ============================================================
 * CONVERTIR FECHA
 * ============================================================
 */
function convertirFecha(valor) {

    if (
        valor instanceof Date
    ) {

        return valor;
    }

    if (
        typeof valor !== 'string'
    ) {

        return null;
    }

    const texto =
        valor.trim();

    // DD/MM/YYYY
    let match =
        texto.match(
            /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
        );

    if (match) {

        return new Date(
            Number(match[3]),
            Number(match[2]) - 1,
            Number(match[1])
        );
    }

    // DD-MM-YYYY
    match =
        texto.match(
            /^(\d{1,2})-(\d{1,2})-(\d{4})$/
        );

    if (match) {

        return new Date(
            Number(match[3]),
            Number(match[2]) - 1,
            Number(match[1])
        );
    }

    // YYYY-MM-DD
    match =
        texto.match(
            /^(\d{4})-(\d{1,2})-(\d{1,2})$/
        );

    if (match) {

        return new Date(
            Number(match[1]),
            Number(match[2]) - 1,
            Number(match[3])
        );
    }

    return null;
}


/**
 * ============================================================
 * FORMATEAR FECHA PARA EXCEL
 * ============================================================
 */
function formatearFechaExcel(
    fecha
) {

    const dia =
        String(
            fecha.getDate()
        ).padStart(2, '0');

    const mes =
        String(
            fecha.getMonth() + 1
        ).padStart(2, '0');

    const anio =
        fecha.getFullYear();

    return `${dia}/${mes}/${anio}`;
}


/**
 * ============================================================
 * CONVERTIR NUMERO
 * ============================================================
 */
function convertirNumero(
    valor
) {

    if (
        typeof valor === 'number'
    ) {

        return Number.isFinite(valor)
            ? valor
            : 0;
    }

    if (
        valor === null ||
        valor === undefined
    ) {

        return 0;
    }

    const numero =
        Number(
            String(valor)
                .replace(/[$,\s]/g, '')
        );

    return Number.isFinite(numero)
        ? numero
        : 0;
}


/**
 * ============================================================
 * CONVERTIR VALOR PARA EXCEL
 * ============================================================
 */
function convertirValor(
    valor
) {

    if (
        valor instanceof Date
    ) {

        return valor;
    }

    if (
        typeof valor === 'number' ||
        typeof valor === 'boolean'
    ) {

        return valor;
    }

    return String(
        valor ?? ''
    );
}


/**
 * ============================================================
 * BORDE
 * ============================================================
 */
function crearBorde() {

    return {
        top: {
            style: 'thin'
        },
        left: {
            style: 'thin'
        },
        bottom: {
            style: 'thin'
        },
        right: {
            style: 'thin'
        }
    };
}


/**
 * ============================================================
 * EXPORTAR
 * ============================================================
 */

module.exports = {
    crearExcel
};
