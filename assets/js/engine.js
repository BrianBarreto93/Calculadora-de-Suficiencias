/*
 * Motor de suficiencia.
 * Replica, fórmula por fórmula, la sección CÁLCULOS de Plantilla_Suficiencia.xlsx
 * para que la calculadora web y el Excel den siempre el mismo resultado.
 * Funciona en el navegador (window.Motor) y en Node (require).
 */
(function (raiz, fabrica) {
  if (typeof module === 'object' && module.exports) module.exports = fabrica();
  else raiz.Motor = fabrica();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // La plantilla de Excel trae espacio para 6 dolores y 8 iniciativas; el tablero y su exportación, para 15.
  var MAX_DOLORES = 15;
  var MAX_INICIATIVAS = 15;
  var ESTADOS = ['En plan', 'En riesgo', 'Retrasada'];
  var ESTADOS_RIESGO = ['Retrasada', 'En riesgo'];

  var PARAMS_DEFECTO = {
    semaforo: [
      { desde: 0, color: 'ROJO' },
      { desde: 0.7, color: 'AMARILLO' },
      { desde: 1, color: 'VERDE' }
    ],
    umbralReemplazar: 0.05,
    coberturaAtribMin: 0.2,
    coberturaResueltoMin: 0.3,
    tope: 0.99
  };

  function vacio(v) { return v === null || v === undefined || v === ''; }
  // Igual que N() de Excel: un número se queda, cualquier otra cosa vale 0.
  function n(v) { return typeof v === 'number' && isFinite(v) ? v : 0; }
  function suma(lista) { var t = 0; for (var i = 0; i < lista.length; i++) t += lista[i]; return t; }
  function sumaProducto(a, b) { var t = 0; for (var i = 0; i < a.length; i++) t += a[i] * b[i]; return t; }

  // BUSCAR aproximado de Excel sobre la tabla del semáforo (ordenada de menor a mayor).
  function colorSemaforo(suf, p) {
    var color = p.semaforo[0].color;
    for (var i = 0; i < p.semaforo.length; i++) if (suf >= p.semaforo[i].desde) color = p.semaforo[i].color;
    return color;
  }

  function calcular(ind, params) {
    var p = params || PARAMS_DEFECTO;
    var dolores = ind.dolores || [];
    var inis = ind.iniciativas || [];
    var sube = ind.sentido === 'sube';

    // Paso 1 · el pastel (problema total) y la brecha a la meta
    var idealUsado = vacio(ind.ideal) ? (sube ? 100 : 0) : n(ind.ideal);
    var problema = vacio(ind.ejecutado) ? 0 : Math.abs(idealUsado - n(ind.ejecutado));
    var brecha = (vacio(ind.sentido) || vacio(ind.ejecutado) || vacio(ind.meta)) ? 0
      : Math.max(0, sube ? n(ind.meta) - n(ind.ejecutado) : n(ind.ejecutado) - n(ind.meta));

    // Paso 2 · prevalencia y parte del pastel (atribución) de cada dolor
    var total = n(ind.totalCasos);
    var prevalencia = dolores.map(function (d) {
      return (n(d.casos) === 0 || total === 0) ? 0 : n(d.casos) / total;
    });
    var sumaPrev = suma(prevalencia);
    var atribAuto = prevalencia.map(function (v) { return sumaPrev === 0 ? 0 : v / sumaPrev; });
    var atribucion = dolores.map(function (d, j) {
      return vacio(d.atribucionManual) ? atribAuto[j] : n(d.atribucionManual);
    });
    var sumaAtrib = suma(atribucion);

    // Paso 3 · lo que DEJA cada iniciativa en cada dolor = 1 − efectividad (× ejecución), con tope
    function efectividad(ini, d) { return n((ini.efectividad || {})[d.id]); }
    var dejaPlan = inis.map(function (ini) {
      return dolores.map(function (d) { return 1 - Math.min(p.tope, efectividad(ini, d)); });
    });
    var dejaReal = inis.map(function (ini) {
      return dolores.map(function (d) { return 1 - Math.min(p.tope, efectividad(ini, d) * n(ini.ejecucion)); });
    });
    // Lo que queda del dolor = se multiplica la columna (la 2.ª iniciativa solo come lo que dejó la 1.ª)
    function queda(deja) {
      return dolores.map(function (_, j) {
        var q = 1;
        for (var i = 0; i < deja.length; i++) q *= deja[i][j];
        return q;
      });
    }
    var quedaPlan = queda(dejaPlan);
    var quedaReal = queda(dejaReal);
    var resPlan = quedaPlan.map(function (q) { return 1 - q; });
    var resReal = quedaReal.map(function (q) { return 1 - q; });
    var ppDolorPlan = dolores.map(function (_, j) { return problema * atribucion[j] * resPlan[j]; });
    var ppDolorReal = dolores.map(function (_, j) { return problema * atribucion[j] * resReal[j]; });

    // Paso 4 · resultado
    var ejecutado = n(ind.ejecutado);
    function escenario(pp) {
      var suf = brecha === 0 ? 0 : pp / brecha;
      var semaforo = vacio(ind.ejecutado) ? 'SIN DATOS' : (brecha === 0 ? 'META CUMPLIDA' : colorSemaforo(suf, p));
      return {
        pp: pp,
        suficiencia: suf,
        semaforo: semaforo,
        proyectado: vacio(ind.ejecutado) ? null : (sube ? ejecutado + pp : ejecutado - pp),
        brechaRestante: Math.max(0, brecha - pp)
      };
    }
    var plan = escenario(suma(ppDolorPlan));
    var real = escenario(suma(ppDolorReal));

    var maxAtrib = atribucion.length ? Math.max.apply(null, atribucion) : 0;
    var principal = maxAtrib === 0 ? -1 : atribucion.indexOf(maxAtrib);

    var resultadoDolores = dolores.map(function (d, j) {
      var alerta = '';
      if (!vacio(d.nombre)) {
        if (atribucion[j] > p.coberturaAtribMin && resReal[j] < p.coberturaResueltoMin) alerta = 'SIN COBERTURA';
        else alerta = resReal[j] >= p.coberturaResueltoMin ? 'CUBIERTO' : 'PESO BAJO';
      }
      return {
        id: d.id,
        prevalencia: prevalencia[j],
        atribucionAuto: atribAuto[j],
        atribucion: atribucion[j],
        manual: !vacio(d.atribucionManual),
        resueltoPlan: resPlan[j],
        resueltoReal: resReal[j],
        ppPlan: ppDolorPlan[j],
        ppReal: ppDolorReal[j],
        alerta: alerta
      };
    });

    // Si la iniciativa falla: se divide para deshacer su multiplicación (escenario realista)
    var resultadoIniciativas = inis.map(function (ini, i) {
      var resSin = dolores.map(function (_, j) { return 1 - quedaReal[j] / dejaReal[i][j]; });
      var ppSin = problema * sumaProducto(atribucion, resSin);
      var sufSin = brecha === 0 ? 0 : ppSin / brecha;
      var aporte = real.pp - ppSin;
      var aportePct = brecha === 0 ? 0 : aporte / brecha;
      var etiqueta = '';
      if (!vacio(ini.nombre)) {
        if (n(ini.ejecucion) === 0 || aportePct < p.umbralReemplazar) etiqueta = 'REEMPLAZAR';
        else etiqueta = ESTADOS_RIESGO.indexOf(ini.estado) >= 0 ? 'EN RIESGO' : 'APORTA';
      }
      var revision = '';
      if (!vacio(ini.nombre)) {
        if (vacio(ini.ejecucion)) revision = 'FALTA EJECUCION';
        else revision = vacio(ini.evidencia) ? 'SIN RESPALDO' : 'OK';
      }
      return {
        id: ini.id,
        ppSin: ppSin,
        suficienciaSin: sufSin,
        aporte: aporte,
        aportePct: aportePct,
        etiqueta: etiqueta,
        revision: revision,
        // Mismo desempate que el Excel: a igual aporte, primero la que está más arriba
        orden: (vacio(ini.nombre) ? -1 : aporte) - (104 + i) / 1e9
      };
    });
    var ranking = resultadoIniciativas
      .map(function (x, i) { return { id: x.id, orden: x.orden, nombrada: !vacio(inis[i].nombre) }; })
      .sort(function (a, b) { return b.orden - a.orden; })
      .filter(function (x) { return x.nombrada; })
      .map(function (x) { return x.id; });

    var nombradas = inis.filter(function (x) { return !vacio(x.nombre); }).length;
    var faltanDatos = resultadoIniciativas.some(function (x) { return x.revision !== '' && x.revision !== 'OK'; });

    return {
      idealUsado: idealUsado,
      problema: problema,
      brecha: brecha,
      sumaPrevalencia: sumaPrev,
      sumaAtribucion: sumaAtrib,
      plan: plan,
      real: real,
      principal: principal,
      dolores: resultadoDolores,
      iniciativas: resultadoIniciativas,
      ranking: ranking,
      chequeoAtribucion: sumaAtrib === 0 ? 'SIN DOLORES' : (Math.abs(sumaAtrib - 1) < 0.0005 ? 'OK' : 'NO SUMA 100'),
      chequeoIniciativas: nombradas === 0 ? 'SIN INICIATIVAS' : (faltanDatos ? 'FALTAN DATOS' : 'OK'),
      posiblePorcentaje: ind.unidad === '%' && n(ind.ejecutado) > 0 && n(ind.ejecutado) < 1 && n(ind.meta) < 1
    };
  }

  // ------------------------------------------------------------------ datos
  var contador = 0;
  function nuevoId(prefijo) {
    contador += 1;
    return prefijo + Date.now().toString(36) + contador.toString(36) + Math.random().toString(36).slice(2, 6);
  }

  function nuevoIndicador() {
    return {
      id: nuevoId('ind-'), nombre: '', sentido: '', unidad: '%',
      ejecutado: null, meta: null, ideal: null, totalCasos: null,
      ejemplo: false, dolores: [], iniciativas: []
    };
  }
  function nuevoDolor() { return { id: nuevoId('d-'), nombre: '', casos: null, atribucionManual: null, evidencia: '' }; }
  function nuevaIniciativa() {
    return { id: nuevoId('i-'), nombre: '', efectividad: {}, ejecucion: 1, estado: 'En plan', responsable: '', fecha: '', evidencia: '' };
  }

  function texto(v, max) { return vacio(v) ? '' : String(v).slice(0, max || 200); }
  function numero(v) { return typeof v === 'number' && isFinite(v) ? v : null; }
  function fraccion(v) { var x = numero(v); return x === null ? null : Math.min(1, Math.max(0, x)); }
  function fecha(v) { return typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : ''; }

  // Limpia lo que llega de un archivo (JSON o Excel) para que siempre tenga la forma esperada.
  function normalizarIndicador(o) {
    o = o && typeof o === 'object' ? o : {};
    var ind = nuevoIndicador();
    if (typeof o.id === 'string' && o.id) ind.id = o.id.slice(0, 60);
    ind.nombre = texto(o.nombre, 120);
    ind.sentido = o.sentido === 'sube' || o.sentido === 'baja' ? o.sentido : '';
    ind.unidad = texto(o.unidad, 20);
    ind.ejecutado = numero(o.ejecutado);
    ind.meta = numero(o.meta);
    ind.ideal = numero(o.ideal);
    ind.totalCasos = numero(o.totalCasos);
    ind.ejemplo = o.ejemplo === true;
    var vistos = {};
    ind.dolores = (Array.isArray(o.dolores) ? o.dolores : []).slice(0, MAX_DOLORES).map(function (d) {
      d = d && typeof d === 'object' ? d : {};
      var id = typeof d.id === 'string' && d.id && !vistos[d.id] ? d.id.slice(0, 60) : nuevoId('d-');
      vistos[id] = true;
      return { id: id, nombre: texto(d.nombre, 120), casos: numero(d.casos), atribucionManual: fraccion(d.atribucionManual), evidencia: texto(d.evidencia, 300) };
    });
    ind.iniciativas = (Array.isArray(o.iniciativas) ? o.iniciativas : []).slice(0, MAX_INICIATIVAS).map(function (x) {
      x = x && typeof x === 'object' ? x : {};
      var efect = {};
      var origen = x.efectividad && typeof x.efectividad === 'object' ? x.efectividad : {};
      ind.dolores.forEach(function (d) { var v = fraccion(origen[d.id]); if (v !== null) efect[d.id] = v; });
      return {
        id: typeof x.id === 'string' && x.id ? x.id.slice(0, 60) : nuevoId('i-'),
        nombre: texto(x.nombre, 120),
        efectividad: efect,
        ejecucion: fraccion(x.ejecucion),
        estado: ESTADOS.indexOf(x.estado) >= 0 ? x.estado : 'En plan',
        responsable: texto(x.responsable, 80),
        fecha: fecha(x.fecha),
        evidencia: texto(x.evidencia, 300)
      };
    });
    return ind;
  }

  function copiarParams(p) { return JSON.parse(JSON.stringify(p)); }

  function normalizarParams(o) {
    var p = copiarParams(PARAMS_DEFECTO);
    if (!o || typeof o !== 'object') return p;
    if (Array.isArray(o.semaforo) && o.semaforo.length === 3) {
      var desdes = o.semaforo.map(function (x) { return numero(x && x.desde); });
      if (desdes.every(function (x) { return x !== null; }) && desdes[0] <= desdes[1] && desdes[1] <= desdes[2]) {
        p.semaforo.forEach(function (x, i) { x.desde = desdes[i]; });
      }
    }
    ['umbralReemplazar', 'coberturaAtribMin', 'coberturaResueltoMin'].forEach(function (k) {
      var v = fraccion(o[k]); if (v !== null) p[k] = v;
    });
    var tope = numero(o.tope);
    if (tope !== null && tope > 0 && tope <= 0.999) p.tope = tope;
    return p;
  }

  function normalizarDatos(o) {
    if (!o || typeof o !== 'object' || !Array.isArray(o.indicadores)) return null;
    var vistos = {};
    var inds = o.indicadores.slice(0, 60).map(normalizarIndicador).map(function (ind) {
      if (vistos[ind.id]) ind.id = nuevoId('ind-');
      vistos[ind.id] = true;
      return ind;
    });
    return { sel: typeof o.sel === 'string' ? o.sel : null, params: normalizarParams(o.params), indicadores: inds };
  }

  return {
    calcular: calcular,
    PARAMS_DEFECTO: PARAMS_DEFECTO,
    MAX_DOLORES: MAX_DOLORES,
    MAX_INICIATIVAS: MAX_INICIATIVAS,
    ESTADOS: ESTADOS,
    vacio: vacio,
    nuevoId: nuevoId,
    nuevoIndicador: nuevoIndicador,
    nuevoDolor: nuevoDolor,
    nuevaIniciativa: nuevaIniciativa,
    normalizarIndicador: normalizarIndicador,
    normalizarParams: normalizarParams,
    normalizarDatos: normalizarDatos,
    copiarParams: copiarParams
  };
});
