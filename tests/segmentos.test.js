// Prueba de los dolores e iniciativas por segmento, de la parte del problema de cada segmento y de lo que se
// guarda en el navegador y en el respaldo .json.
// Uso: node --test tests/segmentos.test.js
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const Motor = require('../assets/js/engine.js');
const crearEjemplos = require('../assets/js/ejemplos.js');

// Churn de prueba con dos segmentos: problema 2,8 pp (hoy 2,8 → ideal 0), brecha 0,8 pp (meta 2,0).
// Hogar: 7.000 casos (70 % del problema) con dos dolores iguales; Pospago: 3.000 casos (30 %) con un dolor.
function dosSegmentos(extra) {
  return Motor.normalizarIndicador(Object.assign({
    nombre: 'Churn de prueba', sentido: 'baja', unidad: '%', ejecutado: 2.8, meta: 2, ideal: null,
    tipoSegmento: 'UMM', segmentos: [{ nombre: 'Hogar', peso: null, casos: 7000 }, { nombre: 'Pospago', peso: null, casos: 3000 }],
    dolores: [
      { id: 'a', segmento: 'Hogar', nombre: 'Precio', casos: 3500 },
      { id: 'b', segmento: 'Hogar', nombre: 'Red', casos: 3500 },
      { id: 'c', segmento: 'Pospago', nombre: 'Factura', casos: 1500 }
    ],
    iniciativas: [
      // Solo Hogar: su 90 % sobre el dolor de Pospago no cuenta
      { id: 'i1', nombre: 'Plan Hogar', segmentos: ['Hogar'], efectividad: { a: 0.5, c: 0.9 }, ejecucion: 1 },
      { id: 'i2', nombre: 'Factura clara', segmentos: [], efectividad: { c: 0.5 }, ejecucion: 0.5 }
    ]
  }, extra || {}));
}
const cerca = (a, b, m) => assert.ok(Math.abs(a - b) < 1e-9, (m || '') + ' → ' + a + ' ≠ ' + b);
function sinSegmentos(r) { const x = Object.assign({}, r); delete x.segmentos; return x; }

test('Un solo segmento: el cálculo es idéntico al de antes (con o sin segmento, crudo o normalizado)', () => {
  for (const ej of crearEjemplos()) {
    const antes = Motor.calcular(Object.assign({}, ej, { tipoSegmento: '', segmentos: [] }));
    const crudo = Motor.calcular(ej);
    const norm = Motor.normalizarIndicador(ej);
    assert.deepEqual(sinSegmentos(crudo), sinSegmentos(antes), ej.nombre + ' crudo');
    assert.deepEqual(sinSegmentos(Motor.calcular(norm)), sinSegmentos(antes), ej.nombre + ' normalizado');
    // Al normalizar, los dolores pasan al único segmento con los casos analizados del indicador
    assert.ok(norm.dolores.every((d) => d.segmento === ej.segmentos[0].nombre));
    assert.equal(norm.segmentos[0].casos, ej.totalCasos);
    assert.equal(norm.totalCasos, null);
    assert.equal(Motor.calcular(norm).segmentos.varios, false);
  }
});

test('Dos segmentos: cada uno pesa según sus casos y el impacto se reparte dentro de cada segmento', () => {
  const r = Motor.calcular(dosSegmentos());
  const [h, p] = r.segmentos.filas;
  cerca(h.parte, 0.7, 'Hogar 7.000 de 10.000');
  cerca(p.parte, 0.3, 'Pospago 3.000 de 10.000');
  cerca(r.dolores[0].atribucion, 0.35);
  cerca(r.dolores[1].atribucion, 0.35);
  cerca(r.dolores[2].atribucion, 0.3);
  cerca(r.dolores[2].prevalencia, 0.5, '% dolor = casos ÷ casos de SU segmento');
  assert.equal(r.chequeoAtribucion, 'OK');
  assert.equal(r.segmentos.chequeo, 'OK');
});

test('Una iniciativa solo resuelve dolores de los segmentos a los que aplica', () => {
  const ind = dosSegmentos();
  const r = Motor.calcular(ind);
  assert.equal(Motor.aplica(ind, ind.iniciativas[0], ind.dolores[2]), false);
  // Pospago: solo «Factura clara» al 50 % × 50 % de ejecución = 25 % (el 90 % de «Plan Hogar» no cuenta)
  cerca(r.dolores[2].resueltoReal, 0.25);
  // pp = 2,8 × (0,35 × 0,5 + 0,3 × 0,25) = 0,7 → suficiencia 0,7 ÷ 0,8
  cerca(r.real.pp, 0.7);
  cerca(r.real.suficiencia, 0.875);
  // Lo de cada segmento suma lo que mueve el indicador
  cerca(r.segmentos.filas[0].pp, 2.8 * 0.35 * 0.5);
  cerca(r.segmentos.filas[1].pp, 2.8 * 0.3 * 0.25);
  cerca(r.segmentos.filas[0].pp + r.segmentos.filas[1].pp, r.real.pp, 'segmentos = total');
  cerca(r.segmentos.filas[0].suficiencia + r.segmentos.filas[1].suficiencia, r.real.suficiencia, 'suficiencias suman la total');
  // Si «Plan Hogar» aplicara a todos, su 90 % sobre Pospago sí contaría
  ind.iniciativas[0].segmentos = [];
  assert.ok(Motor.calcular(ind).real.pp > r.real.pp);
});

test('Parte del problema escrita a mano: manda sobre los casos y el resto se reparte por casos', () => {
  const ind = dosSegmentos();
  ind.segmentos[0].peso = 0.5;
  const r = Motor.calcular(ind);
  cerca(r.segmentos.filas[0].parte, 0.5);
  cerca(r.segmentos.filas[1].parte, 0.5, 'Pospago toma lo que falta');
  cerca(r.dolores[0].atribucion, 0.25);
  cerca(r.dolores[2].atribucion, 0.5);
  assert.equal(r.segmentos.chequeo, 'OK');
  // Las dos a mano que no suman 100 % se avisan
  ind.segmentos[1].peso = 0.3;
  assert.equal(Motor.calcular(ind).segmentos.chequeo, 'NO SUMA 100');
  // Y si pasan de 100 % se rechazan
  ind.segmentos[1].peso = 0.6;
  assert.equal(Motor.calcular(ind).segmentos.chequeo, 'PASA DE 100');
  assert.match(Motor.revisarPesos([0.5, 0.6]), /más de 100 %/);
});

test('Un segmento sin dolores deja su parte del problema sin resolver (y el motor lo muestra)', () => {
  const ind = dosSegmentos();
  ind.dolores = ind.dolores.filter((d) => d.segmento === 'Hogar');
  const r = Motor.calcular(ind);
  cerca(r.sumaAtribucion, 0.7);
  assert.equal(r.chequeoAtribucion, 'NO SUMA 100');
  assert.equal(r.segmentos.filas[1].dolores, 0);
  cerca(r.segmentos.filas[1].parte, 0.3);
  // Sin casos en ningún segmento no hay cómo repartir el problema
  ind.segmentos.forEach((s) => { s.casos = null; });
  assert.equal(Motor.calcular(ind).segmentos.chequeo, 'SIN CASOS');
});

test('Pesos que suman 100 % exacto (con decimales) y «partes iguales» (1/3) se aceptan', () => {
  assert.equal(Motor.revisarPesos([0.3333, 0.3333, 0.3334]), '');
  assert.match(Motor.revisarPesos([1.2]), /0 % a 100 %/);
  assert.match(Motor.revisarPesos([-0.1]), /0 % a 100 %/);
  const ind = dosSegmentos({ tipoSegmento: 'Transversal' });
  ind.segmentos.push({ nombre: 'Pyme Fijo', peso: null, casos: 100 });
  ind.segmentos.forEach((s) => { s.peso = 1 / 3; });
  const r = Motor.calcular(ind);
  assert.equal(r.segmentos.chequeo, 'OK');
  cerca(r.segmentos.sumaPartes, 1);
});

test('Sin segmento marcado el motor lo avisa', () => {
  assert.equal(Motor.calcular(Motor.nuevoIndicador()).segmentos.chequeo, 'SIN SEGMENTO');
});

test('Al limpiar datos: solo segmentos de la unidad elegida, con sus casos; iniciativas solo con segmentos marcados', () => {
  const umm = Motor.normalizarIndicador({ tipoSegmento: 'UMM', segmentos: [{ nombre: 'Prepago', peso: 0.5, casos: 10 }, { nombre: 'Pyme Fijo' }, { nombre: 'Hogar', casos: 20 }, 'x'] });
  assert.deepEqual(umm.segmentos, [{ nombre: 'Hogar', peso: null, casos: 20 }, { nombre: 'Prepago', peso: 0.5, casos: 10 }]);
  const tr = Motor.normalizarIndicador({ tipoSegmento: 'Transversal', segmentos: [{ nombre: 'Pyme Móvil', peso: 7, casos: null }, { nombre: 'Hogar', peso: 0.4, casos: null }] });
  assert.deepEqual(tr.segmentos.map((s) => s.peso), [0.4, 1]);
  assert.equal(Motor.normalizarIndicador({ tipoSegmento: 'Otro', segmentos: [{ nombre: 'Hogar' }] }).segmentos.length, 0);
  const ind = dosSegmentos();
  const otro = Motor.normalizarIndicador(Object.assign({}, ind, { iniciativas: [
    { id: 'x', segmentos: ['Pospago', 'Prepago', 'Hogar'] }, { id: 'y', segmentos: ['Pospago', 'Nada'] }] }));
  assert.deepEqual(otro.iniciativas[0].segmentos, [], 'todos los marcados = aplica a todos');
  assert.deepEqual(otro.iniciativas[1].segmentos, ['Pospago']);
  // Un dolor con un segmento que no está marcado queda sin segmento
  const suelto = Motor.normalizarIndicador(Object.assign({}, ind, { dolores: [{ id: 'z', segmento: 'Prepago', casos: 5 }] }));
  assert.equal(suelto.dolores[0].segmento, '');
});

test('Máximo 15 dolores por segmento y 15 iniciativas por segmento marcado', () => {
  const ind = dosSegmentos();
  const muchos = [];
  for (let k = 0; k < 20; k++) muchos.push({ id: 'h' + k, segmento: 'Hogar' }, { id: 'p' + k, segmento: 'Pospago' });
  const n = Motor.normalizarIndicador(Object.assign({}, ind, { dolores: muchos, iniciativas: Array.from({ length: 40 }, (_, k) => ({ id: 'i' + k })) }));
  assert.equal(n.dolores.filter((d) => d.segmento === 'Hogar').length, 15);
  assert.equal(n.dolores.filter((d) => d.segmento === 'Pospago').length, 15);
  assert.equal(n.dolores[14].segmento, 'Hogar', 'agrupados por segmento');
  assert.equal(n.dolores[15].segmento, 'Pospago');
  assert.equal(n.iniciativas.length, 30);
  assert.equal(Motor.maxIniciativas(n), 30);
});

test('Datos de antes (pesos Transversales sin casos): los pesos no se usan y el resultado no cambia', () => {
  const ej = crearEjemplos()[0];
  const antes = Motor.calcular(Object.assign({}, ej, { tipoSegmento: '', segmentos: [] }));
  const viejo = Object.assign({}, ej, { tipoSegmento: 'Transversal', segmentos: [{ nombre: 'Hogar', peso: 0.6 }, { nombre: 'Pospago', peso: 0.4 }] });
  const n = Motor.normalizarIndicador(JSON.parse(JSON.stringify(viejo)));
  assert.deepEqual(n.segmentos, [{ nombre: 'Hogar', peso: null, casos: null }, { nombre: 'Pospago', peso: null, casos: null }]);
  assert.ok(n.dolores.every((d) => d.segmento === ''), 'con 2 segmentos no se adivina el segmento de cada dolor');
  assert.equal(n.totalCasos, ej.totalCasos);
  const r = Motor.calcular(n);
  assert.deepEqual(sinSegmentos(r), sinSegmentos(antes));
  assert.deepEqual(r.segmentos.filas.map((f) => f.nombre), ['Hogar', 'Pospago', ''], 'los dolores quedan en «sin segmento»');
  cerca(r.segmentos.filas[2].parte, 1);
});

test('Ordenar y acomodar: un segmento recibe los dolores sueltos; con varios se agrupan', () => {
  const ind = Motor.normalizarIndicador({ tipoSegmento: 'UMM', totalCasos: 50, dolores: [{ id: 'a', casos: 5 }, { id: 'b', casos: 7 }] });
  assert.equal(ind.dolores[0].segmento, '');
  ind.segmentos.push({ nombre: 'Prepago', peso: null, casos: null });
  Motor.acomodarSegmentos(ind);
  assert.deepEqual(ind.dolores.map((d) => d.segmento), ['Prepago', 'Prepago']);
  assert.equal(ind.segmentos[0].casos, 50);
  ind.segmentos.unshift({ nombre: 'Hogar', peso: null, casos: 10 });
  ind.dolores.push(Motor.nuevoDolor('Hogar'));
  Motor.ordenarDolores(ind);
  assert.deepEqual(ind.dolores.map((d) => d.segmento), ['Hogar', 'Prepago', 'Prepago']);
});

test('Persistencia: el formato anterior (versión 2) se migra al nuevo sin perder nada', () => {
  const v2 = { version: 2, sel: 'ej-instalaciones', params: Motor.PARAMS_DEFECTO, indicadores: crearEjemplos() };
  const e = Motor.normalizarEspacio(JSON.parse(JSON.stringify(v2)));
  assert.equal(e.version, 3);
  assert.equal(e.actual.sel, 'ej-instalaciones');
  assert.equal(e.actual.indicadores.length, 2);
  assert.deepEqual(e.calculos, []);
  assert.equal(e.activo, null);
  assert.equal(e.base, null, 'lo migrado cuenta como cambios sin guardar');
  // Antes de la versión 2 no venía la marca de ejemplo
  const v1 = Motor.normalizarEspacio({ indicadores: [{ id: 'ej-churn', nombre: 'Churn' }] });
  assert.equal(v1.actual.indicadores[0].ejemplo, true);
});

test('Persistencia: el espacio (cálculos, segmentos, pesos y cálculo activo) vuelve igual', () => {
  const ind = dosSegmentos();
  ind.segmentos[0].peso = 0.6;
  const datos = { sel: ind.id, params: Motor.PARAMS_DEFECTO, indicadores: [ind] };
  const espacio = {
    version: 3, activo: 'calc-1', base: 'x',
    actual: datos,
    calculos: [{ id: 'calc-1', nombre: 'Churn · 2 nov 2026', fecha: '2026-11-02T15:00:00.000Z', datos: datos }]
  };
  const vuelta = Motor.normalizarEspacio(JSON.parse(JSON.stringify(espacio)));
  assert.equal(vuelta.activo, 'calc-1');
  assert.equal(vuelta.base, 'x');
  assert.equal(vuelta.calculos[0].nombre, 'Churn · 2 nov 2026');
  assert.deepEqual(vuelta.actual.indicadores[0].segmentos, ind.segmentos);
  assert.deepEqual(vuelta.calculos[0].datos.indicadores[0].segmentos, ind.segmentos);
  assert.equal(vuelta.actual.indicadores[0].tipoSegmento, 'UMM');
  assert.deepEqual(vuelta.actual.indicadores[0].dolores.map((d) => d.segmento), ['Hogar', 'Hogar', 'Pospago']);
  assert.deepEqual(vuelta.actual.indicadores[0].iniciativas.map((x) => x.segmentos), [['Hogar'], []]);
  assert.deepEqual(Motor.calcular(vuelta.actual.indicadores[0]), Motor.calcular(ind), 'mismo resultado al volver');
});

test('Respaldo .json: se limpian datos raros, ids repetidos y cálculos vacíos', () => {
  const ind = Motor.normalizarIndicador(crearEjemplos()[0]);
  const datos = { sel: ind.id, params: Motor.PARAMS_DEFECTO, indicadores: [ind] };
  const e = Motor.normalizarEspacio({
    app: 'Suficiencia 2027', tipo: 'respaldo', version: 3, activo: 'no-existe', actual: datos,
    calculos: [{ id: 'a', nombre: '', datos: datos }, { id: 'a', nombre: 'Repetido', datos: datos }, { id: 'b', datos: { indicadores: [] } }, null, 'x']
  });
  assert.equal(e.calculos.length, 2);
  assert.notEqual(e.calculos[0].id, e.calculos[1].id, 'ids únicos');
  assert.equal(e.calculos[0].nombre, 'Cálculo sin nombre');
  assert.equal(e.activo, null, 'un cálculo activo que no existe se suelta');
  assert.equal(Motor.normalizarEspacio('hola'), null);
  assert.equal(Motor.normalizarEspacio({ version: 3 }), null);
});

test('El estado de la iniciativa ya no cuenta: misma etiqueta y mismo resultado con cualquier estado', () => {
  const base = crearEjemplos()[0];
  const resultados = Motor.ESTADOS.map((e) => Motor.calcular(Object.assign({}, base, { iniciativas: base.iniciativas.map((x) => Object.assign({}, x, { estado: e })) })));
  resultados.forEach((r) => {
    assert.ok(r.iniciativas.every((x) => x.etiqueta !== 'EN RIESGO'));
    assert.deepEqual(r, resultados[0]);
  });
});

test('Sin límite fijo de indicadores', () => {
  const muchos = Array.from({ length: 80 }, (_, i) => ({ id: 'i' + i, nombre: 'Indicador ' + i }));
  assert.equal(Motor.normalizarDatos({ indicadores: muchos }).indicadores.length, 80);
});

test('El aporte de las iniciativas suma lo que mueve el indicador (se reparte lo que comen de cada dolor)', () => {
  const ind = dosSegmentos();
  ind.iniciativas.push({ id: 'i3', nombre: 'Red Hogar', segmentos: [], efectividad: { a: 0.3 }, ejecucion: 1 });
  const r = Motor.calcular(ind);
  // Dolor a (35 % de 2,8 pp): queda (1 − 0,5) × (1 − 0,3) = 0,35 → recupera 0,98 × 0,65 = 0,637 pp, repartido 0,5 : 0,3
  cerca(r.dolores[0].ppReal, 0.637);
  cerca(r.iniciativas[0].aporte, 0.637 * 0.5 / 0.8, 'Plan Hogar');
  cerca(r.iniciativas[2].aporte, 0.637 * 0.3 / 0.8, 'Red Hogar');
  // Dolor c (30 %): solo Factura clara (50 % × 50 % de ejecución) → 0,84 × 0,25 = 0,21 pp
  cerca(r.iniciativas[1].aporte, 0.21, 'Factura clara');
  cerca(r.iniciativas.reduce((t, x) => t + x.aporte, 0), r.real.pp, 'suma = lo que mueve');
  cerca(r.iniciativas[1].aportePct, 0.21 / 0.8);
});
