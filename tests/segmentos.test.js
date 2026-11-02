// Prueba del segmento, del peso por segmento y de lo que se guarda en el navegador y en el respaldo .json.
// Uso: node --test tests/segmentos.test.js
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const Motor = require('../assets/js/engine.js');
const crearEjemplos = require('../assets/js/ejemplos.js');

function transversal(pesos) {
  const ind = Motor.normalizarIndicador(crearEjemplos()[0]);
  ind.tipoSegmento = 'Transversal';
  ind.segmentos = [{ nombre: 'Hogar', peso: pesos[0] }, { nombre: 'Pospago', peso: pesos[1] }, { nombre: 'Pyme Fijo', peso: pesos[2] }];
  return ind;
}
function sinSegmentos(r) { const x = Object.assign({}, r); delete x.segmentos; return x; }

test('Con pesos: aporte = peso × lo que mueve el indicador, y aportes + sin asignar = total', () => {
  const ind = transversal([0.6, 0.25, null]);
  const r = Motor.calcular(ind);
  const s = r.segmentos;
  assert.equal(s.total, r.real.pp);
  assert.equal(s.filas[0].pp, 0.6 * r.real.pp);
  assert.equal(s.filas[1].pp, 0.25 * r.real.pp);
  assert.equal(s.filas[2].pp, null, 'sin peso no hay aporte');
  assert.ok(Math.abs(s.sinAsignar.peso - 0.15) < 1e-12);
  const suma = s.filas.reduce((t, f) => t + (f.pp || 0), 0) + s.sinAsignar.pp;
  assert.ok(Math.abs(suma - s.total) < 1e-12, 'aportes + sin asignar = total');
  assert.equal(s.chequeo, 'FALTA REPARTIR');
});

test('El ejemplo del pedido: indicador de 3 pp y Hogar 60 % → aporta 1,8 pp', () => {
  const s = Motor.repartirSegmentos({ tipoSegmento: 'Transversal', segmentos: [{ nombre: 'Hogar', peso: 0.6 }] }, 3);
  assert.ok(Math.abs(s.filas[0].pp - 1.8) < 1e-12);
  assert.ok(Math.abs(s.sinAsignar.pp - 1.2) < 1e-12);
  assert.ok(Math.abs(s.sinAsignar.peso - 0.4) < 1e-12);
});

test('Pesos que suman 100 % exacto (con decimales) se aceptan', () => {
  assert.equal(Motor.revisarPesos([0.3333, 0.3333, 0.3334]), '');
  const r = Motor.calcular(transversal([0.3333, 0.3333, 0.3334]));
  assert.equal(r.segmentos.chequeo, 'OK');
  assert.ok(Math.abs(r.segmentos.sinAsignar.pp) < 1e-12);
});

test('Si los pesos suman más de 100 % se rechazan', () => {
  assert.match(Motor.revisarPesos([0.6, 0.5]), /más de 100 %/);
  assert.match(Motor.revisarPesos([1.2]), /0 % a 100 %/);
  assert.match(Motor.revisarPesos([-0.1]), /0 % a 100 %/);
  const r = Motor.calcular(transversal([0.6, 0.5, null]));
  assert.equal(r.segmentos.chequeo, 'PASA DE 100');
  assert.ok(r.segmentos.filas.every((f) => f.pp === null), 'no se reparten puntos con pesos inválidos');
});

test('Sin pesos el cálculo es idéntico al de antes', () => {
  for (const ej of crearEjemplos()) {
    const antes = Motor.calcular(Object.assign({}, ej, { tipoSegmento: '', segmentos: [] }));
    const conSegmento = Motor.calcular(transversal([null, null, null]));
    const mismo = Motor.calcular(Object.assign({}, ej, { tipoSegmento: 'Transversal', segmentos: [{ nombre: 'Hogar', peso: null }] }));
    assert.deepEqual(sinSegmentos(mismo), sinSegmentos(antes), ej.nombre);
    assert.equal(conSegmento.segmentos.conPesos, false);
    assert.equal(conSegmento.segmentos.sinAsignar.peso, 1);
    assert.equal(mismo.segmentos.sinAsignar.pp, mismo.real.pp, 'todo queda sin asignar');
  }
  // Y con pesos, la suficiencia tampoco cambia
  assert.deepEqual(sinSegmentos(Motor.calcular(transversal([0.6, 0.4, null]))), sinSegmentos(Motor.calcular(transversal([null, null, null]))));
});

test('Sin segmento marcado el motor lo avisa', () => {
  assert.equal(Motor.calcular(Motor.nuevoIndicador()).segmentos.chequeo, 'SIN SEGMENTO');
});

test('Al limpiar datos: solo segmentos de la unidad elegida y peso solo en Transversal', () => {
  const umm = Motor.normalizarIndicador({ tipoSegmento: 'UMM', segmentos: [{ nombre: 'Prepago', peso: 0.5 }, { nombre: 'Pyme Fijo' }, { nombre: 'Hogar' }, 'x'] });
  assert.deepEqual(umm.segmentos, [{ nombre: 'Hogar', peso: null }, { nombre: 'Prepago', peso: null }]);
  const tr = Motor.normalizarIndicador({ tipoSegmento: 'Transversal', segmentos: [{ nombre: 'Pyme Móvil', peso: 7 }, { nombre: 'Hogar', peso: 0.4 }] });
  assert.deepEqual(tr.segmentos, [{ nombre: 'Hogar', peso: 0.4 }, { nombre: 'Pyme Móvil', peso: 1 }]);
  assert.equal(Motor.normalizarIndicador({ tipoSegmento: 'Otro', segmentos: [{ nombre: 'Hogar' }] }).segmentos.length, 0);
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
  const ind = transversal([0.6, 0.25, null]);
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
  assert.equal(vuelta.actual.indicadores[0].tipoSegmento, 'Transversal');
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

test('Sin límite fijo de indicadores', () => {
  const muchos = Array.from({ length: 80 }, (_, i) => ({ id: 'i' + i, nombre: 'Indicador ' + i }));
  assert.equal(Motor.normalizarDatos({ indicadores: muchos }).indicadores.length, 80);
});
