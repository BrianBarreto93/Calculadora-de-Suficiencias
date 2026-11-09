// Comprueba que el motor web da exactamente lo mismo que Plantilla_Suficiencia.xlsx.
// Los valores esperados se leyeron del Excel recalculado (hojas IND01_Churn e IND02_Instalaciones).
// Única diferencia a propósito: desde el 2 de nov el tablero ya no usa el Estado, así que donde la plantilla
// dice EN RIESGO el tablero dice APORTA (la suficiencia y todo lo demás siguen iguales).
// Segunda diferencia a propósito (9 de nov): el aporte de cada iniciativa ya no es «lo que se pierde si falla»
// (la plantilla) sino su parte de lo que mueve el indicador, así que los aportes suman el total. Por eso
// aporte, aportePct y ranking se comparan con los valores del tablero, no con los del Excel.
// Uso: node --test tests/paridad_excel.test.js
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const Motor = require('../assets/js/engine.js');
const crearEjemplos = require('../assets/js/ejemplos.js');

const EXCEL = {
  'ej-churn': {
    problema: 2.8, brecha: 0.79999999999999982,
    ppPlan: 1.0418692307692305, ppReal: 0.68344769230769231,
    sufPlan: 1.3023365384615384, sufReal: 0.85430961538461558,
    semPlan: 'VERDE', semReal: 'AMARILLO',
    proyPlan: 1.7581307692307693, proyReal: 2.1165523076923076,
    restantePlan: 0, restanteReal: 0.11655230769230751,
    principal: 'Precio / oferta de la competencia',
    atribucion: [0.38461538461538458, 0.23076923076923075, 0.19230769230769229, 0.11538461538461538, 0.076923076923076927],
    resPlan: [0.5099499999999999, 0.30000000000000004, 0.405, 0.25, 0],
    resReal: [0.26938000000000006, 0.17999999999999994, 0.37950000000000006, 0.22499999999999998, 0],
    ppDolorPlan: [0.54917692307692278, 0.19384615384615386, 0.21807692307692303, 0.08076923076923076, 0],
    ppDolorReal: [0.29010153846153847, 0.11630769230769225, 0.20434615384615382, 0.072692307692307681, 0],
    alertas: ['SIN COBERTURA', 'SIN COBERTURA', 'CUBIERTO', 'PESO BAJO', 'PESO BAJO'],
    aporte: [0.18006302387267906, 0.17301578249336869, 0.11630769230769225, 0.20405769230769227, 0.010003501326259955],
    aportePct: [0.22507877984084887, 0.2162697281167109, 0.14538461538461533, 0.2550721153846154, 0.012504376657824947],
    etiquetas: ['APORTA', 'APORTA', 'APORTA', 'APORTA', 'REEMPLAZAR'],
    revision: ['OK', 'OK', 'OK', 'OK', 'SIN RESPALDO'],
    ranking: ['Resolución en el primer contacto', 'Nuevo portafolio de planes competitivos', 'Ofertas de retención con modelo predictivo', 'Densificación de red en zonas críticas', 'Campaña de imagen de marca en redes']
  },
  'ej-instalaciones': {
    problema: 18, brecha: 10,
    ppPlan: 12.632727272727275, ppReal: 12.235090909090912,
    sufPlan: 1.2632727272727275, sufReal: 1.2235090909090913,
    semPlan: 'VERDE', semReal: 'VERDE',
    proyPlan: 94.63272727272728, proyReal: 94.235090909090914,
    restantePlan: 0, restanteReal: 0,
    principal: 'Falta de técnicos en la zona',
    atribucion: [0.40909090909090917, 0.27272727272727276, 0.18181818181818185, 0.13636363636363638],
    resPlan: [0.6399999999999999, 0.73, 0.8, 0.7],
    resReal: [0.58600000000000008, 0.73, 0.8, 0.7],
    ppDolorPlan: [4.7127272727272729, 3.5836363636363644, 2.6181818181818191, 1.7181818181818185],
    ppDolorReal: [4.3150909090909106, 3.5836363636363644, 2.6181818181818191, 1.7181818181818185],
    alertas: ['CUBIERTO', 'CUBIERTO', 'CUBIERTO', 'CUBIERTO'],
    aporte: [3.1356818181818187, 3.640857954545456, 3.2924147727272732, 2.166136363636364],
    aportePct: [0.31356818181818186, 0.3640857954545456, 0.32924147727272735, 0.2166136363636364],
    etiquetas: ['APORTA', 'APORTA', 'APORTA', 'APORTA'],
    revision: ['OK', 'OK', 'OK', 'OK'],
    ranking: ['Contratación de técnicos tercerizados', 'Reposición automática de inventario', 'Agendamiento con confirmación por WhatsApp', 'Validación de factibilidad en el punto de venta']
  }
};

function igual(real, esperado, que) {
  const tolerancia = 1e-12 * Math.max(1, Math.abs(esperado));
  assert.ok(Math.abs(real - esperado) <= tolerancia, `${que}: web ${real} ≠ Excel ${esperado}`);
}
function igualLista(real, esperado, que) {
  assert.equal(real.length, esperado.length, `${que}: largo distinto`);
  esperado.forEach((v, i) => igual(real[i], v, `${que}[${i}]`));
}

for (const ind of crearEjemplos()) {
  const e = EXCEL[ind.id];
  test(`${ind.nombre}: mismo resultado que el Excel`, () => {
    const r = Motor.calcular(ind);
    igual(r.problema, e.problema, 'Problema total');
    igual(r.brecha, e.brecha, 'Brecha');
    igual(r.plan.pp, e.ppPlan, 'Puntos plan');
    igual(r.real.pp, e.ppReal, 'Puntos realista');
    igual(r.plan.suficiencia, e.sufPlan, 'Suficiencia plan');
    igual(r.real.suficiencia, e.sufReal, 'Suficiencia realista');
    assert.equal(r.plan.semaforo, e.semPlan);
    assert.equal(r.real.semaforo, e.semReal);
    igual(r.plan.proyectado, e.proyPlan, 'Proyectado plan');
    igual(r.real.proyectado, e.proyReal, 'Proyectado realista');
    igual(r.plan.brechaRestante, e.restantePlan, 'Brecha restante plan');
    igual(r.real.brechaRestante, e.restanteReal, 'Brecha restante realista');
    assert.equal(ind.dolores[r.principal].nombre, e.principal);
    igualLista(r.dolores.map(d => d.atribucion), e.atribucion, 'Atribución');
    igualLista(r.dolores.map(d => d.resueltoPlan), e.resPlan, 'Resuelto plan');
    igualLista(r.dolores.map(d => d.resueltoReal), e.resReal, 'Resuelto realista');
    igualLista(r.dolores.map(d => d.ppPlan), e.ppDolorPlan, 'Puntos por dolor plan');
    igualLista(r.dolores.map(d => d.ppReal), e.ppDolorReal, 'Puntos por dolor realista');
    assert.deepEqual(r.dolores.map(d => d.alerta), e.alertas);
    igualLista(r.iniciativas.map(x => x.aporte), e.aporte, 'Aporte');
    igualLista(r.iniciativas.map(x => x.aportePct), e.aportePct, 'Aporte % brecha');
    igual(r.iniciativas.reduce((t, x) => t + x.aporte, 0), r.real.pp, 'Los aportes suman lo que mueve');
    assert.deepEqual(r.iniciativas.map(x => x.etiqueta), e.etiquetas);
    assert.deepEqual(r.iniciativas.map(x => x.revision), e.revision);
    const nombre = id => ind.iniciativas.find(x => x.id === id).nombre;
    assert.deepEqual(r.ranking.map(nombre), e.ranking);
  });
}

test('Indicador vacío: sin datos y sin errores', () => {
  const r = Motor.calcular(Motor.nuevoIndicador());
  assert.equal(r.real.semaforo, 'SIN DATOS');
  assert.equal(r.brecha, 0);
  assert.equal(r.chequeoAtribucion, 'SIN DOLORES');
  assert.equal(r.chequeoIniciativas, 'SIN INICIATIVAS');
});

test('Un archivo con datos raros se limpia sin romper el cálculo', () => {
  const datos = Motor.normalizarDatos({
    indicadores: [{ nombre: 42, sentido: 'x', ejecutado: '2', dolores: [{ id: 'a', casos: 5, nombre: '<b>hola</b>' }], iniciativas: [{ nombre: 'X', efectividad: { a: 7, z: 0.5 }, ejecucion: -3 }] }]
  });
  const ind = datos.indicadores[0];
  assert.equal(ind.nombre, '42');
  assert.equal(ind.sentido, '');
  assert.equal(ind.ejecutado, null);
  assert.deepEqual(ind.iniciativas[0].efectividad, { a: 1 });
  assert.equal(ind.iniciativas[0].ejecucion, 0);
  assert.doesNotThrow(() => Motor.calcular(ind, datos.params));
});
