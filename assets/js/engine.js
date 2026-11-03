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

  // La plantilla de Excel trae espacio para 6 dolores y 8 iniciativas; el tablero y su exportación, para 15
  // dolores por segmento y 15 iniciativas por segmento marcado (una iniciativa puede servir a varios).
  var MAX_DOLORES = 15;
  var MAX_INICIATIVAS = 15;
  function maxIniciativas(ind) { return MAX_INICIATIVAS * Math.max(1, ((ind && ind.segmentos) || []).length); }
  // El estado se sigue leyendo de la plantilla y de los archivos, pero ya no cambia ningún resultado
  var ESTADOS = ['En plan', 'En riesgo', 'Retrasada'];
  // Segmentos por unidad de negocio. Un indicador Transversal puede marcar segmentos de las dos.
  var SEGMENTOS = { UMM: ['Hogar', 'Pospago', 'Prepago'], UMC: ['Pyme Fijo', 'Pyme Móvil'] };
  var TIPOS_SEGMENTO = ['UMM', 'UMC', 'Transversal'];
  function segmentosDe(tipo) { return tipo === 'Transversal' ? SEGMENTOS.UMM.concat(SEGMENTOS.UMC) : (SEGMENTOS[tipo] || []); }

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

  // Los grupos del pastel: un grupo por segmento marcado y, si hay dolores sin segmento (de antes o de la
  // plantilla), un grupo «sin segmento» que usa los casos analizados del indicador.
  // Parte del problema de cada grupo: la que escribe el usuario (peso) o, si no, la que sale de sus casos
  // analizados (casos del grupo ÷ casos de los grupos automáticos, sobre lo que no se repartió a mano).
  // Con un solo grupo, ese grupo es todo el problema: así el cálculo de un segmento es el de siempre.
  function segmentoDe(ind, d) {
    var s = d && d.segmento;
    return s && (ind.segmentos || []).some(function (x) { return x.nombre === s; }) ? s : '';
  }
  function grupos(ind) {
    var lista = (ind.segmentos || []).map(function (s) { return { nombre: s.nombre, casos: s.casos, peso: s.peso }; });
    if ((ind.dolores || []).some(function (d) { return segmentoDe(ind, d) === ''; }) || !lista.length) {
      lista.push({ nombre: '', casos: ind.totalCasos, peso: null });
    }
    var varios = lista.length > 1;
    lista.forEach(function (g) { g.manual = varios && typeof g.peso === 'number' && isFinite(g.peso) ? g.peso : null; });
    var sumaManual = suma(lista.map(function (g) { return g.manual || 0; }));
    var autos = lista.filter(function (g) { return g.manual === null; });
    var casosAutos = suma(autos.map(function (g) { return Math.max(0, n(g.casos)); }));
    var resto = Math.max(0, 1 - sumaManual);
    lista.forEach(function (g) {
      g.parteAuto = !varios ? 1 : (casosAutos === 0 ? 0 : resto * Math.max(0, n(g.casos)) / casosAutos);
      g.parte = g.manual !== null ? g.manual : g.parteAuto;
    });
    return lista;
  }
  // ¿La iniciativa trabaja sobre este dolor? Sí, si es para todos los segmentos ([]), si el dolor no tiene
  // segmento o si el segmento del dolor está entre los de la iniciativa.
  function aplica(ind, ini, d) {
    var seg = segmentoDe(ind, d);
    var lista = Array.isArray(ini.segmentos) ? ini.segmentos : [];
    return seg === '' || !lista.length || lista.indexOf(seg) >= 0;
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

    // Paso 2 · prevalencia (casos ÷ casos analizados de su segmento) y parte del pastel (atribución):
    // dentro de cada segmento se reparte su parte del problema según el % dolor.
    var gs = grupos(ind);
    var grupoDe = {};
    gs.forEach(function (g) { grupoDe[g.nombre] = g; });
    var segDolor = dolores.map(function (d) { return segmentoDe(ind, d); });
    var prevalencia = dolores.map(function (d, j) {
      var total = n(grupoDe[segDolor[j]].casos);
      return (n(d.casos) === 0 || total === 0) ? 0 : n(d.casos) / total;
    });
    var sumaPrev = suma(prevalencia);
    var prevGrupo = {};
    gs.forEach(function (g) { prevGrupo[g.nombre] = 0; });
    prevalencia.forEach(function (v, j) { prevGrupo[segDolor[j]] += v; });
    var atribAuto = prevalencia.map(function (v, j) {
      var sg = prevGrupo[segDolor[j]];
      return sg === 0 ? 0 : grupoDe[segDolor[j]].parte * v / sg;
    });
    var atribucion = dolores.map(function (d, j) {
      return vacio(d.atribucionManual) ? atribAuto[j] : n(d.atribucionManual);
    });
    var sumaAtrib = suma(atribucion);

    // Paso 3 · lo que DEJA cada iniciativa en cada dolor = 1 − efectividad (× ejecución), con tope.
    // Una iniciativa no toca los dolores de los segmentos donde no aplica.
    function efectividad(ini, d) { return aplica(ind, ini, d) ? n((ini.efectividad || {})[d.id]) : 0; }
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
        // Desde el 2 de nov el estado (En plan / En riesgo / Retrasada) ya no cuenta: el retraso se expresa con la ejecución
        else etiqueta = 'APORTA';
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
    var segmentos = resumenSegmentos(ind, gs, function (nombre) {
      var r = { real: 0, plan: 0, dolores: 0, atribucion: 0 };
      dolores.forEach(function (d, j) {
        if (segDolor[j] !== nombre) return;
        r.real += ppDolorReal[j];
        r.plan += ppDolorPlan[j];
        r.atribucion += atribucion[j];
        r.dolores += 1;
      });
      return r;
    }, brecha, real.pp);
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
      posiblePorcentaje: ind.unidad === '%' && n(ind.ejecutado) > 0 && n(ind.ejecutado) < 1 && n(ind.meta) < 1,
      segmentos: segmentos
    };
  }

  // Paso 5 · resultado por segmento. Cada segmento recupera lo que resuelven las iniciativas en SUS dolores,
  // así que la suma de los segmentos es exactamente lo que mueve el indicador (y su suficiencia, la del total).
  // filas: una por grupo (nombre '' = dolores sin segmento) con su parte del problema, sus puntos y su
  // aporte a la suficiencia. chequeo: SIN SEGMENTO · PASA DE 100 (pesos) · SIN CASOS · NO SUMA 100 · OK.
  function resumenSegmentos(ind, gs, sumar, brecha, total) {
    var varios = gs.length > 1;
    var filas = gs.map(function (g) {
      var x = sumar(g.nombre);
      return {
        nombre: g.nombre,
        casos: vacio(g.casos) ? null : n(g.casos),
        peso: g.manual,
        parteAuto: g.parteAuto,
        parte: g.parte,
        dolores: x.dolores,
        atribucion: x.atribucion,
        pp: x.real,
        ppPlan: x.plan,
        suficiencia: brecha === 0 ? 0 : x.real / brecha,
        suficienciaPlan: brecha === 0 ? 0 : x.plan / brecha
      };
    });
    var sumaPesos = suma(filas.map(function (f) { return f.peso || 0; }));
    var sumaPartes = suma(filas.map(function (f) { return f.parte; }));
    var chequeo = 'OK';
    if (!(ind.segmentos || []).length) chequeo = 'SIN SEGMENTO';
    else if (varios && sumaPesos > 1 + 1e-9) chequeo = 'PASA DE 100';
    else if (varios && sumaPartes === 0) chequeo = 'SIN CASOS';
    else if (varios && Math.abs(sumaPartes - 1) > 0.0005) chequeo = 'NO SUMA 100';
    return {
      filas: filas,
      varios: varios,
      total: total,
      conPesos: filas.some(function (f) { return f.peso !== null; }),
      sumaPesos: sumaPesos,
      sumaPartes: sumaPartes,
      chequeo: chequeo
    };
  }

  // ¿Se pueden usar estos pesos? (fracciones de 0 a 1 o null). Devuelve '' si sí, o el motivo en palabras simples.
  function revisarPesos(pesos) {
    var total = 0;
    for (var i = 0; i < pesos.length; i++) {
      var p = pesos[i];
      if (p === null || p === undefined) continue;
      if (typeof p !== 'number' || !isFinite(p) || p < 0 || p > 1) return 'Cada peso va de 0 % a 100 %.';
      total += p;
    }
    return total > 1 + 1e-9 ? 'Los pesos suman más de 100 %. Baja alguno.' : '';
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
      tipoSegmento: '', segmentos: [],
      ejemplo: false, dolores: [], iniciativas: []
    };
  }
  function nuevoDolor(segmento) { return { id: nuevoId('d-'), segmento: segmento || '', nombre: '', casos: null, atribucionManual: null, evidencia: '' }; }
  // segmentos: [] = aplica a todos los segmentos del indicador
  function nuevaIniciativa() {
    return { id: nuevoId('i-'), nombre: '', segmentos: [], efectividad: {}, ejecucion: 1, estado: 'En plan', responsable: '', fecha: '', evidencia: '' };
  }

  // Deja los dolores agrupados en el orden de los segmentos (los sin segmento al final), sin cambiar el
  // orden dentro de cada segmento: así la numeración D1, D2… y las columnas del paso 3 van por segmento.
  function ordenarDolores(ind) {
    var orden = (ind.segmentos || []).map(function (s) { return s.nombre; });
    var lugar = function (d) { var k = orden.indexOf(segmentoDe(ind, d)); return k < 0 ? orden.length : k; };
    ind.dolores = ind.dolores.map(function (d, j) { return { d: d, j: j }; })
      .sort(function (a, b) { return lugar(a.d) - lugar(b.d) || a.j - b.j; })
      .map(function (x) { return x.d; });
    return ind;
  }
  // Con un solo segmento marcado, los dolores sin segmento (de antes, de la plantilla o de un segmento que
  // se quitó) pasan a ese segmento con sus casos analizados: el resultado no cambia.
  function acomodarSegmentos(ind) {
    ind.dolores.forEach(function (d) { if (!segmentoDe(ind, d)) d.segmento = ''; });
    if (ind.segmentos.length === 1) {
      var s = ind.segmentos[0];
      var sueltos = ind.dolores.filter(function (d) { return d.segmento === ''; });
      if (sueltos.length || !vacio(ind.totalCasos)) {
        sueltos.forEach(function (d) { d.segmento = s.nombre; });
        if (vacio(s.casos)) s.casos = ind.totalCasos;
        ind.totalCasos = null;
      }
    }
    // Los casos del indicador solo cuentan para los dolores sin segmento: si no queda ninguno, sobran
    if (ind.segmentos.length && !ind.dolores.some(function (d) { return d.segmento === ''; })) ind.totalCasos = null;
    var marcados = ind.segmentos.map(function (s) { return s.nombre; });
    ind.iniciativas.forEach(function (x) {
      x.segmentos = (x.segmentos || []).filter(function (nombre) { return marcados.indexOf(nombre) >= 0; });
      if (x.segmentos.length === marcados.length) x.segmentos = [];
    });
    return ordenarDolores(ind);
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
    // Segmentos: solo los de la unidad elegida, en el orden del catálogo, con sus casos analizados y su
    // peso (parte del problema escrita a mano, opcional).
    ind.tipoSegmento = TIPOS_SEGMENTO.indexOf(o.tipoSegmento) >= 0 ? o.tipoSegmento : '';
    var marcados = Array.isArray(o.segmentos) ? o.segmentos : [];
    // Antes del 3 de nov el peso solo repartía los puntos al final y los segmentos no traían casos: ese peso
    // no significa «parte del problema», así que no se usa (si se usara, cambiaría el resultado guardado).
    var pesoDeAntes = marcados.some(function (x) { return x && typeof x === 'object'; }) &&
      marcados.every(function (x) { return !x || typeof x !== 'object' || x.casos === undefined; });
    ind.segmentos = segmentosDe(ind.tipoSegmento).map(function (nombre) {
      var s = marcados.filter(function (x) { return x && x.nombre === nombre; })[0];
      return s ? { nombre: nombre, peso: pesoDeAntes ? null : fraccion(s.peso), casos: numero(s.casos) } : null;
    }).filter(Boolean);
    var nombresSeg = ind.segmentos.map(function (s) { return s.nombre; });
    var vistos = {};
    var porGrupo = {};
    ind.dolores = (Array.isArray(o.dolores) ? o.dolores : []).map(function (d) {
      d = d && typeof d === 'object' ? d : {};
      var seg = nombresSeg.indexOf(d.segmento) >= 0 ? d.segmento : '';
      porGrupo[seg] = (porGrupo[seg] || 0) + 1;
      if (porGrupo[seg] > MAX_DOLORES) return null;
      var id = typeof d.id === 'string' && d.id && !vistos[d.id] ? d.id.slice(0, 60) : nuevoId('d-');
      vistos[id] = true;
      return { id: id, segmento: seg, nombre: texto(d.nombre, 120), casos: numero(d.casos), atribucionManual: fraccion(d.atribucionManual), evidencia: texto(d.evidencia, 300) };
    }).filter(Boolean);
    ind.iniciativas = (Array.isArray(o.iniciativas) ? o.iniciativas : []).slice(0, maxIniciativas(ind)).map(function (x) {
      x = x && typeof x === 'object' ? x : {};
      var efect = {};
      var origen = x.efectividad && typeof x.efectividad === 'object' ? x.efectividad : {};
      ind.dolores.forEach(function (d) { var v = fraccion(origen[d.id]); if (v !== null) efect[d.id] = v; });
      var segs = Array.isArray(x.segmentos) ? x.segmentos : [];
      return {
        id: typeof x.id === 'string' && x.id ? x.id.slice(0, 60) : nuevoId('i-'),
        nombre: texto(x.nombre, 120),
        segmentos: nombresSeg.filter(function (nombre) { return segs.indexOf(nombre) >= 0; }),
        efectividad: efect,
        ejecucion: fraccion(x.ejecucion),
        estado: ESTADOS.indexOf(x.estado) >= 0 ? x.estado : 'En plan',
        responsable: texto(x.responsable, 80),
        fecha: fecha(x.fecha),
        evidencia: texto(x.evidencia, 300)
      };
    });
    return acomodarSegmentos(ind);
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
    var inds = o.indicadores.map(normalizarIndicador).map(function (ind) {
      if (vistos[ind.id]) ind.id = nuevoId('ind-');
      vistos[ind.id] = true;
      return ind;
    });
    return { sel: typeof o.sel === 'string' ? o.sel : null, params: normalizarParams(o.params), indicadores: inds };
  }

  // Todo lo que se guarda en el navegador y en el respaldo .json (versión 3):
  //   { version: 3, activo: id del cálculo guardado que está abierto (o null),
  //     actual: el tablero que se ve (sel, params, indicadores) + base (cómo estaba al abrirlo, para saber si cambió),
  //     calculos: [{ id, nombre, fecha, datos: { sel, params, indicadores } }] }
  // También acepta el formato anterior (versión 2: un solo tablero) y lo migra. Devuelve null si no sirve.
  var VERSION_ESPACIO = 3;
  function normalizarEspacio(o) {
    if (!o || typeof o !== 'object') return null;
    if (!(o.version >= 3)) {
      var viejo = normalizarDatos(o);
      // Antes de la versión 2 los datos no traían la marca de ejemplo
      if (viejo && !(o.version >= 2)) viejo.indicadores.forEach(function (x) { if (/^ej-/.test(x.id)) x.ejemplo = true; });
      return viejo ? { version: VERSION_ESPACIO, activo: null, actual: viejo, base: null, calculos: [] } : null;
    }
    var actual = normalizarDatos(o.actual);
    var vistos = {};
    var calculos = (Array.isArray(o.calculos) ? o.calculos : []).map(function (c) {
      var datos = c && typeof c === 'object' ? normalizarDatos(c.datos) : null;
      if (!datos || !datos.indicadores.length) return null;
      var id = typeof c.id === 'string' && c.id && !vistos[c.id] ? c.id.slice(0, 60) : nuevoId('calc-');
      vistos[id] = true;
      return { id: id, nombre: texto(c.nombre, 80) || 'Cálculo sin nombre', fecha: typeof c.fecha === 'string' ? c.fecha.slice(0, 40) : '', datos: datos };
    }).filter(Boolean);
    var activo = typeof o.activo === 'string' && vistos[o.activo] ? o.activo : null;
    if (!actual) return calculos.length ? { version: VERSION_ESPACIO, activo: null, actual: null, base: null, calculos: calculos } : null;
    return { version: VERSION_ESPACIO, activo: activo, actual: actual, base: typeof o.base === 'string' ? o.base : null, calculos: calculos };
  }

  return {
    calcular: calcular,
    PARAMS_DEFECTO: PARAMS_DEFECTO,
    MAX_DOLORES: MAX_DOLORES,
    MAX_INICIATIVAS: MAX_INICIATIVAS,
    ESTADOS: ESTADOS,
    SEGMENTOS: SEGMENTOS,
    TIPOS_SEGMENTO: TIPOS_SEGMENTO,
    segmentosDe: segmentosDe,
    segmentoDe: segmentoDe,
    aplica: aplica,
    maxIniciativas: maxIniciativas,
    ordenarDolores: ordenarDolores,
    acomodarSegmentos: acomodarSegmentos,
    revisarPesos: revisarPesos,
    normalizarEspacio: normalizarEspacio,
    VERSION_ESPACIO: VERSION_ESPACIO,
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
