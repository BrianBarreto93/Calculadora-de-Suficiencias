// Prueba del lector y escritor de Excel (assets/js/excel.js).
// Escribe un libro, lo vuelve a leer, revisa que cada XML esté bien formado y lee la plantilla real guardada por Excel.
// Uso: node --test tests/excel.test.js
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const X = require('../assets/js/excel.js');

// Revisa que cada etiqueta abra y cierre en orden (así se detecta un XML que Excel no podría abrir).
function errorXml(xml) {
  const cuerpo = xml.replace(/^<\?xml[^>]*\?>\s*/, '');
  const re = /<(\/?)([A-Za-z_][\w:.-]*)((?:\s+[\w:.-]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>/g;
  const pila = [];
  let ultimo = 0;
  let m;
  while ((m = re.exec(cuerpo))) {
    if (/[<>]/.test(cuerpo.slice(ultimo, m.index))) return 'texto con < o > antes de <' + m[2] + '>';
    ultimo = re.lastIndex;
    if (m[1]) { const abierta = pila.pop(); if (abierta !== m[2]) return '</' + m[2] + '> cierra <' + abierta + '>'; }
    else if (!m[4]) pila.push(m[2]);
  }
  if (/[<>]/.test(cuerpo.slice(ultimo))) return 'texto inválido al final';
  return pila.length ? 'sin cerrar: ' + pila.join(', ') : '';
}
// Orden de los elementos de una hoja que exige Excel (solo los que usa el escritor)
const ORDEN_HOJA = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'mergeCells', 'dataValidations', 'hyperlinks', 'pageMargins', 'pageSetup'];

function libroDePrueba() {
  const h = new X.Hoja('Resumen');
  h.anchos = [6, 40, 14];
  h.color = 'DA291C';
  h.congelar = { filas: 2, columnas: 2 };
  h.poner(0, 0, 'Acentos y símbolos: ¿Qué? «Ñandú» & <menor> > "comillas" \'apóstrofo\'', { negrita: true, tam: 18 }).alto(0, 30);
  h.poner(1, 1, '  espacios al inicio y al final  ');
  h.poner(2, 1, 'línea 1\nlínea 2', { ajustar: true });
  h.poner(3, 2, 0.8543, { formato: '0%', fondo: 'FFF2CC', borde: 'BFBFBF', h: 'center' });
  h.poner(4, 2, 0.68344769230769231, { formato: '0.00" pp"' });
  h.poner(5, 2, -3.5);
  h.poner(6, 2, 1e-7);
  h.poner(7, 2, 12000, { formato: '#,##0' });
  h.poner(8, 2, true);
  h.poner(9, 2, '');
  h.poner(10, 2, null, { fondo: 'DDEBF7' });
  h.poner(11, 2, X.serieDeFecha('2027-03-31'), { formato: 'fecha' });
  h.poner(12, 1, 'control\u0001fuera');
  h.unir(13, 1, 13, 2);
  h.validar(3, 2, 4, 2, { tipo: 'decimal', min: 0, max: 1, titulo: 'Porcentaje', mensaje: 'Entre 0 % y 100 %' });
  h.validar(7, 2, 7, 2, { tipo: 'entero', min: 0 });
  h.validar(1, 1, 1, 1, { tipo: 'lista', valores: ['Que suba', 'Que baje'] });
  const otra = X.nombreHoja("Datos: 'raros' [x]", ['Resumen']);
  h.enlazar(14, 1, otra, 'Ver hoja');
  const h2 = new X.Hoja(otra);
  h2.poner(0, 0, 'Segunda hoja');
  return { libro: { hojas: [h, h2], titulo: 'Prueba & <libro>' }, otra };
}

test('Escribe un libro y lo vuelve a leer igual', async () => {
  const { libro, otra } = libroDePrueba();
  const bytes = await X.crearXlsx(libro);
  const leido = await X.leerXlsx(bytes);
  assert.deepEqual(leido.nombres, ['Resumen', otra]);
  const c = leido.hojas.Resumen;
  assert.equal(c.A1.v, 'Acentos y símbolos: ¿Qué? «Ñandú» & <menor> > "comillas" \'apóstrofo\'');
  assert.equal(c.B2.v, '  espacios al inicio y al final  ');
  assert.equal(c.B3.v, 'línea 1\nlínea 2');
  assert.equal(c.C4.v, 0.8543);
  assert.equal(c.C5.v, 0.68344769230769231);
  assert.equal(c.C6.v, -3.5);
  assert.equal(c.C7.v, 1e-7);
  assert.equal(c.C8.v, 12000);
  assert.equal(c.C9.v, true);
  assert.equal(c.C10, undefined, 'el texto vacío no crea celda');
  assert.equal(c.C11, undefined, 'una celda solo con estilo no tiene valor');
  assert.equal(X.fechaDeSerie(c.C12.v), '2027-03-31');
  assert.equal(c.B13.v, 'controlfuera', 'se quitan los caracteres de control que el XML no admite');
  assert.equal(leido.hojas[otra].A1.v, 'Segunda hoja');
});

test('Cada XML del archivo está bien formado y en el orden que exige Excel', async () => {
  const bytes = await X.crearXlsx(libroDePrueba().libro);
  const partes = await X.leerPartes(bytes);
  for (const n of ['[Content_Types].xml', '_rels/.rels', 'docProps/core.xml', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']) {
    assert.ok(partes[n], 'falta ' + n);
    assert.equal(errorXml(partes[n]), '', n);
  }
  const hoja = partes['xl/worksheets/sheet1.xml'];
  const posiciones = ORDEN_HOJA.map((e) => hoja.indexOf('<' + e + (e === 'cols' ? '>' : ''))).filter((p) => p >= 0);
  assert.deepEqual(posiciones, [...posiciones].sort((a, b) => a - b), 'orden de los elementos de la hoja');
  assert.match(hoja, /<pane xSplit="2" ySplit="2" topLeftCell="C3" activePane="bottomRight" state="frozen"\/>/);
  assert.match(partes['xl/workbook.xml'], /<sheet name="Datos 'raros' x" sheetId="2"/);
});

test('Sin compresión del navegador el archivo también se puede leer', async () => {
  const original = global.CompressionStream;
  global.CompressionStream = undefined;
  try {
    const bytes = await X.crearXlsx(libroDePrueba().libro);
    const leido = await X.leerXlsx(bytes);
    assert.equal(leido.hojas.Resumen.C8.v, 12000);
  } finally { global.CompressionStream = original; }
});

test('Nombres de hoja válidos para Excel', () => {
  const usados = [];
  assert.equal(X.nombreHoja('01 Churn: mensual/pospago [x]', usados), '01 Churn mensual pospago x');
  assert.equal(X.nombreHoja("'Con apóstrofos'", usados), 'Con apóstrofos');
  assert.equal(X.nombreHoja('', usados), 'Hoja');
  assert.equal(X.nombreHoja('Un nombre de indicador larguísimo que no cabe', usados), 'Un nombre de indicador…', 'recorta en una palabra completa');
  assert.equal(X.nombreHoja('UN NOMBRE DE INDICADOR LARGUÍSIMO que no cabe', usados), 'UN NOMBRE DE INDICADOR… (2)', 'no repite nombres aunque cambien mayúsculas');
  assert.equal(X.nombreHoja('02 Instalaciones en fecha (FTTH)', []), '02 Instalaciones en fecha…', 'no deja un paréntesis abierto');
  assert.ok(usados.every((u) => u.length <= 31));
});

test('Fechas de Excel', () => {
  assert.equal(X.serieDeFecha('2027-03-31'), 46477);
  assert.equal(X.fechaDeSerie(46477), '2027-03-31');
  assert.equal(X.fechaDeSerie(46477.75), '2027-03-31');
  assert.equal(X.serieDeFecha('31/03/2027'), null);
  assert.equal(X.fechaDeSerie('x'), '');
});

test('Lee la plantilla real guardada por Excel (textos compartidos, fórmulas y fechas)', async () => {
  const bytes = fs.readFileSync(path.join(__dirname, '..', 'Plantilla_Suficiencia.xlsx'));
  const libro = await X.leerXlsx(bytes);
  assert.deepEqual(libro.nombres, ['Guia', 'Resumen', 'IND01_Churn', 'IND02_Instalaciones', 'IND_PLANTILLA', 'Parametros']);
  const churn = libro.hojas.IND01_Churn;
  assert.equal(churn.C5.v, 'Churn mensual pospago');
  assert.equal(churn.C8.v, 2.8);
  assert.equal(churn.C17.v, 'Precio / oferta de la competencia');
  assert.equal(X.fechaDeSerie(churn.L29.v), '2027-03-31');
  assert.ok(Math.abs(churn.D42.v - 0.85430961538461558) < 1e-12, 'lee el valor calculado de una fórmula');
  assert.equal(libro.hojas.Parametros.C15.v, 0.99);
});
