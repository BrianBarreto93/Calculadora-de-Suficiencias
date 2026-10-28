/*
 * Suficiencia 2027 · tablero.
 * Toda la matemática vive en engine.js (Motor); este archivo solo pinta y edita.
 * Los textos que escribe el usuario se insertan siempre con textContent, nunca como HTML.
 */
(function () {
  'use strict';

  var Motor = window.Motor;
  var crearEjemplos = window.SuficienciaEjemplos;
  var CLAVE_DATOS = 'suficiencia.datos.v1';
  var CLAVE_INTRO = 'suficiencia.intro.v1';
  var SHEETJS = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  var movimientoReducido = window.matchMedia('(prefers-reduced-motion: reduce)');

  var SEMAFORO = {
    'VERDE': { txt: 'Alcanza', tono: 'ok', icono: 'check' },
    'AMARILLO': { txt: 'Cerca', tono: 'alerta', icono: 'triangulo' },
    'ROJO': { txt: 'No alcanza', tono: 'critico', icono: 'x' },
    'META CUMPLIDA': { txt: 'Meta cumplida', tono: 'ok', icono: 'bandera' },
    'SIN DATOS': { txt: 'Sin datos', tono: 'neutro', icono: 'menos' }
  };
  var ETIQUETA = {
    'APORTA': { txt: 'Aporta', tono: 'ok', icono: 'check', que: 'Va bien. Seguimiento normal.' },
    'EN RIESGO': { txt: 'En riesgo', tono: 'alerta', icono: 'reloj', que: 'Aporta, pero está retrasada o en riesgo. Si falla, la suficiencia cae: dale seguimiento cercano.' },
    'REEMPLAZAR': { txt: 'Reemplazar', tono: 'critico', icono: 'cambiar', que: 'Casi no aporta o su ejecución es 0 %. Cámbiala o rediséñala.' }
  };
  var ALERTA_DOLOR = {
    'SIN COBERTURA': { txt: 'Sin cobertura', tono: 'critico', icono: 'triangulo', que: 'Un dolor importante casi no se resuelve. Falta una iniciativa para él.' },
    'CUBIERTO': { txt: 'Cubierto', tono: 'ok', icono: 'check', que: 'Las iniciativas resuelven una parte suficiente del dolor.' },
    'PESO BAJO': { txt: 'Peso bajo', tono: 'neutro', icono: 'menos', que: 'Se resuelve poco, pero pesa poco en el problema.' }
  };
  var REVISION = {
    'SIN RESPALDO': { txt: 'Sin respaldo', tono: 'alerta', icono: 'ayuda', que: 'Falta la evidencia que justifica los %. Toda cifra debe tener su fuente.' },
    'FALTA EJECUCION': { txt: 'Falta % de ejecución', tono: 'alerta', icono: 'ayuda', que: 'Sin % de ejecución la iniciativa cuenta como 0 %.' }
  };
  var AYUDAS = {
    ideal: 'El mejor valor posible. Si lo dejas vacío: 100 si es mejor que suba, 0 si es mejor que baje.',
    casos: 'Total de casos del análisis. Ej.: clientes que se dieron de baja en el periodo. Sirve para calcular la prevalencia.',
    prevalencia: 'Casos con este dolor ÷ casos analizados. La suma puede pasar de 100 % porque un caso puede tener varios dolores.',
    atribucion: 'Lo que le toca a este dolor del problema total. Se calcula sola con los casos; escribe un % solo si quieres cambiarla.',
    resuelto: 'Qué parte del dolor se comen tus iniciativas con su ejecución esperada. La rayita marca lo que se comerían si todo se cumple.',
    efectividad: 'Qué % del dolor resuelve la iniciativa si se hace completa. Vacío = 0 %.',
    ejecucion: 'Qué % de la iniciativa estará listo de verdad en 2027. 100 % = a tiempo y completa.',
    aporte: 'Puntos que se pierden si la iniciativa falla, y cómo quedaría la suficiencia sin ella.'
  };

  // ---------------------------------------------------------------- utilidades
  function $(sel, raiz) { return (raiz || document).querySelector(sel); }
  function $$(sel, raiz) { return Array.prototype.slice.call((raiz || document).querySelectorAll(sel)); }
  function vacio(v) { return v === null || v === undefined || v === ''; }
  function n(v) { return typeof v === 'number' && isFinite(v) ? v : 0; }
  function limitar(v, a, b) { return Math.min(b, Math.max(a, v)); }

  function poner(e, c) {
    if (c === null || c === undefined || c === false) return;
    if (Array.isArray(c)) { c.forEach(function (x) { poner(e, x); }); return; }
    e.appendChild(c instanceof Node ? c : document.createTextNode(String(c)));
  }
  function aplicar(e, props) {
    Object.keys(props).forEach(function (k) {
      var v = props[k];
      if (v === null || v === undefined || v === false) return;
      if (k === 'class') e.setAttribute('class', v);
      else if (k === 'style') e.style.cssText = v;
      else if (k === 'data') Object.keys(v).forEach(function (d) { if (v[d] !== undefined && v[d] !== null) e.dataset[d] = v[d]; });
      else if (k === 'value') e.value = v;
      else if (v === true) e.setAttribute(k, '');
      else e.setAttribute(k, v);
    });
  }
  function h(tag, props) {
    var e = document.createElement(tag);
    if (props) aplicar(e, props);
    for (var i = 2; i < arguments.length; i++) poner(e, arguments[i]);
    return e;
  }
  var SVGNS = 'http://www.w3.org/2000/svg';
  function s(tag, attrs) {
    var e = document.createElementNS(SVGNS, tag);
    if (attrs) Object.keys(attrs).forEach(function (k) { if (attrs[k] !== null && attrs[k] !== undefined) e.setAttribute(k, attrs[k]); });
    for (var i = 2; i < arguments.length; i++) poner(e, arguments[i]);
    return e;
  }
  function fijar(e) { e.replaceChildren(); for (var i = 1; i < arguments.length; i++) poner(e, arguments[i]); return e; }
  function icono(nombre, clase) {
    var u = document.createElementNS(SVGNS, 'use');
    u.setAttribute('href', '#i-' + nombre);
    var sv = document.createElementNS(SVGNS, 'svg');
    sv.setAttribute('class', 'ic' + (clase ? ' ' + clase : ''));
    sv.setAttribute('aria-hidden', 'true');
    sv.appendChild(u);
    return sv;
  }
  var cuentaIds = 0;
  function idUnico(p) { cuentaIds += 1; return (p || 'c') + '-' + cuentaIds; }
  function dos(x) { return (x < 10 ? '0' : '') + x; }

  // formato colombiano: 2,80 · 12.000 · 85 %
  var formatos = {};
  function nf(min, max, agrupar) {
    var k = min + '-' + max + (agrupar ? 'g' : '');
    if (!formatos[k]) formatos[k] = new Intl.NumberFormat('es-CO', { minimumFractionDigits: min, maximumFractionDigits: max, useGrouping: agrupar });
    return formatos[k];
  }
  function dec(v, d) { return nf(d, d, true).format(v); }
  function pct(v, d) { return nf(d || 0, d || 0, true).format(v * 100) + ' %'; }
  function editableDec(v) { return vacio(v) ? '' : nf(0, 4, false).format(v); }
  function editablePct(v) { return vacio(v) ? '' : nf(0, 2, false).format(v * 100); }
  function editableEntero(v) { return vacio(v) ? '' : nf(0, 0, true).format(v); }
  function formatoEditable(tipo, v) {
    if (tipo === 'decimal') return editableDec(v);
    if (tipo === 'entero') return editableEntero(v);
    if (tipo === 'pct') return editablePct(v);
    return vacio(v) ? '' : String(v);
  }
  // Acepta 2,8 · 2.8 · 1.234,5 — la coma siempre es decimal.
  function leerNumero(t) {
    t = String(t).replace(/[\s %]/g, '');
    if (t === '') return null;
    if (t.indexOf(',') >= 0) t = t.replace(/\./g, '').replace(',', '.');
    else if ((t.match(/\./g) || []).length > 1) t = t.replace(/\./g, '');
    var v = Number(t);
    return isFinite(v) ? v : NaN;
  }
  function leerEntero(t) {
    t = String(t).trim();
    if (t === '') return null;
    if (/[^\d.,\s ]/.test(t)) return NaN;
    var d = t.replace(/\D/g, '');
    return d === '' ? null : Number(d);
  }

  function nombreDe(ind) { return ind.nombre ? ind.nombre : 'Indicador sin nombre'; }
  function unidadPuntos(ind) { return ind.unidad === '%' ? 'pp' : (ind.unidad || 'puntos'); }
  function valorConUnidad(v, ind) {
    if (vacio(v)) return '—';
    var u = ind.unidad || '';
    return dec(v, 2) + (u ? ' ' + u : '');
  }
  function puntos(v, ind) { return dec(v, 2) + ' ' + unidadPuntos(ind); }
  function textoSentido(ind) { return ind.sentido === 'sube' ? 'Sube es mejor' : ind.sentido === 'baja' ? 'Baja es mejor' : 'Sin sentido elegido'; }
  function unirLista(lista) { return lista.length < 2 ? lista.join('') : lista.slice(0, -1).join(', ') + ' y ' + lista[lista.length - 1]; }

  // ---------------------------------------------------------------- estado
  var estado = cargar();
  var abiertos = {};
  var refs = null;
  var deshacer = null;
  var tAviso = 0;
  var tGuardar = 0;

  function datosIniciales() {
    var inds = crearEjemplos();
    return { sel: inds[0].id, params: Motor.copiarParams(Motor.PARAMS_DEFECTO), indicadores: inds };
  }
  function cargar() {
    var datos = null;
    try {
      var crudo = window.localStorage.getItem(CLAVE_DATOS);
      if (crudo) datos = Motor.normalizarDatos(JSON.parse(crudo));
    } catch (e) { datos = null; }
    if (!datos || !datos.indicadores.length) datos = datosIniciales();
    if (!datos.indicadores.some(function (x) { return x.id === datos.sel; })) datos.sel = datos.indicadores[0].id;
    return datos;
  }
  function guardar() {
    try {
      window.localStorage.setItem(CLAVE_DATOS, JSON.stringify({ version: 1, sel: estado.sel, params: estado.params, indicadores: estado.indicadores }));
    } catch (e) { /* sin almacenamiento: la sesión sigue funcionando */ }
  }
  function guardarLuego() { clearTimeout(tGuardar); tGuardar = setTimeout(guardar, 250); }
  function foto() { return JSON.stringify({ sel: estado.sel, params: estado.params, indicadores: estado.indicadores }); }
  function volverA(f) {
    var d = JSON.parse(f);
    estado.sel = d.sel; estado.params = d.params; estado.indicadores = d.indicadores;
    pintarTodo(true);
  }
  function actual() {
    for (var i = 0; i < estado.indicadores.length; i++) if (estado.indicadores[i].id === estado.sel) return estado.indicadores[i];
    return estado.indicadores[0];
  }
  function buscar(lista, id) { for (var i = 0; i < lista.length; i++) if (lista[i].id === id) return lista[i]; return null; }

  // ---------------------------------------------------------------- referencias del DOM
  var D = {
    cuerpo: document.body,
    barraResumen: $('#barra-resumen'),
    riel: $('#riel-pista'),
    rielCuenta: $('#riel-cuenta'),
    intro: $('#intro'),
    escenario: $('#escenario'),
    cartel: $('#cartel'),
    titulo: $('#titulo-indicador'),
    metaLinea: $('#meta-linea'),
    frase: $('#frase'),
    kpis: $('#kpis'),
    pastelSvg: $('#pastel-svg'),
    pastelAlt: $('#pastel-alt'),
    pastelCentro: $('#pastel-centro'),
    leyenda: $('#leyenda'),
    vigilar: $('#vigilar'),
    vigilarCuenta: $('#vigilar-cuenta'),
    aportes: $('#aportes'),
    pasos: $('#pasos'),
    taller: $('#taller-cuerpo'),
    ajustes: $('#ajustes-cuerpo'),
    dlg: $('#dlg-como'),
    tip: $('#tip'),
    aviso: $('#aviso'),
    archivo: $('#archivo')
  };

  function chip(meta, chico) {
    if (!meta) return null;
    return h('span', { class: 'chip chip--' + meta.tono + (chico ? ' chip--chico' : '') }, icono(meta.icono), meta.txt);
  }

  // ---------------------------------------------------------------- el pastel y los anillos
  var GEO = { c: 180, r0: 92, R: 146, anillo: 170 };
  var CIRC = 2 * Math.PI * GEO.anillo;
  function pt(r, a) { return (GEO.c + r * Math.sin(a)).toFixed(2) + ' ' + (GEO.c - r * Math.cos(a)).toFixed(2); }
  function sector(r0, r1, a0, a1) {
    var g = a1 - a0 > Math.PI ? 1 : 0;
    return 'M' + pt(r1, a0) + 'A' + r1.toFixed(2) + ' ' + r1.toFixed(2) + ' 0 ' + g + ' 1 ' + pt(r1, a1) +
      'L' + pt(r0, a1) + 'A' + r0.toFixed(2) + ' ' + r0.toFixed(2) + ' 0 ' + g + ' 0 ' + pt(r0, a0) + 'Z';
  }
  function arco(r, a0, a1) {
    var g = a1 - a0 > Math.PI ? 1 : 0;
    return 'M' + pt(r, a0) + 'A' + r.toFixed(2) + ' ' + r.toFixed(2) + ' 0 ' + g + ' 1 ' + pt(r, a1);
  }
  // El área pintada de cada porción es proporcional a lo resuelto (no su radio).
  function radioRelleno(fr) { return Math.sqrt(GEO.r0 * GEO.r0 + limitar(fr, 0, 1) * (GEO.R * GEO.R - GEO.r0 * GEO.r0)); }

  function porciones(ind, r) {
    var lista = [];
    var tot = r.sumaAtribucion;
    if (tot <= 0) return lista;
    var a = 0;
    ind.dolores.forEach(function (d, j) {
      var at = r.dolores[j].atribucion;
      if (at <= 0) return;
      var ancho = 2 * Math.PI * at / tot;
      lista.push({ d: d, j: j, a0: a, a1: a + Math.min(ancho, 2 * Math.PI - 1e-4) });
      a += ancho;
    });
    return lista;
  }

  function dibujarPastel(svg, ind, r, op) {
    svg.replaceChildren();
    svg.classList.toggle('entrando', !!op.entrando);
    if (op.anillo) {
      var suf = r.real.suficiencia;
      var tono = (SEMAFORO[r.real.semaforo] || SEMAFORO['SIN DATOS']).tono;
      var g = s('g', { class: 'anillo tono-' + tono, transform: 'rotate(-90 180 180)' });
      g.appendChild(s('circle', { class: 'anillo-pista', cx: 180, cy: 180, r: GEO.anillo }));
      var v1 = limitar(suf, 0, 1);
      var v2 = limitar(suf - 1, 0, 1);
      if (v1 > 0.002) g.appendChild(s('circle', { class: 'anillo-vuelta', cx: 180, cy: 180, r: GEO.anillo, 'stroke-dasharray': (v1 * CIRC).toFixed(2) + ' ' + CIRC.toFixed(2) }));
      if (v2 > 0.002) g.appendChild(s('circle', { class: 'anillo-vuelta anillo-vuelta--2', cx: 180, cy: 180, r: GEO.anillo, 'stroke-dasharray': (v2 * CIRC).toFixed(2) + ' ' + CIRC.toFixed(2) }));
      svg.appendChild(g);
    }
    var gp = s('g', { class: 'porciones' });
    var lista = porciones(ind, r);
    if (!lista.length) gp.appendChild(s('path', { class: 'porcion-base vacia', d: sector(GEO.r0, GEO.R, 0, 2 * Math.PI - 1e-4) }));
    lista.forEach(function (x) {
      var dr = r.dolores[x.j];
      var g1 = s('g', { class: 'porcion', 'data-dolor': x.d.id });
      g1.appendChild(s('path', { class: 'porcion-base', d: sector(GEO.r0, GEO.R, x.a0, x.a1) }));
      if (dr.resueltoReal > 0.0005) g1.appendChild(s('path', { class: 'porcion-resuelta', d: sector(GEO.r0, radioRelleno(dr.resueltoReal), x.a0, x.a1) }));
      if (op.plan && dr.resueltoPlan - dr.resueltoReal > 0.004 && x.a1 - x.a0 > 0.09) {
        g1.appendChild(s('path', { class: 'porcion-plan', d: arco(radioRelleno(dr.resueltoPlan), x.a0 + 0.03, x.a1 - 0.03) }));
      }
      if (op.etiquetas && x.a1 - x.a0 > 0.2) {
        var am = (x.a0 + x.a1) / 2;
        g1.appendChild(s('text', { class: 'porcion-num', x: (180 + 156 * Math.sin(am)).toFixed(1), y: (180 - 156 * Math.cos(am)).toFixed(1) }, String(x.j + 1)));
      }
      gp.appendChild(g1);
    });
    svg.appendChild(gp);
  }

  function anilloMini(suf, tono) {
    var C = 2 * Math.PI * 14;
    var sv = s('svg', { viewBox: '0 0 36 36', 'aria-hidden': 'true', class: 'tono-' + tono });
    sv.appendChild(s('circle', { cx: 18, cy: 18, r: 14, fill: 'none', stroke: 'rgba(255,255,255,.14)', 'stroke-width': 4.5 }));
    var v = limitar(suf, 0, 1);
    if (v > 0.002) {
      sv.appendChild(s('circle', {
        cx: 18, cy: 18, r: 14, fill: 'none', style: 'stroke: var(--anillo)', 'stroke-width': 4.5, 'stroke-linecap': 'round',
        'stroke-dasharray': (v * C).toFixed(2) + ' ' + C.toFixed(2), transform: 'rotate(-90 18 18)'
      }));
    }
    return sv;
  }

  // ---------------------------------------------------------------- riel (Tus indicadores)
  function tarjeta(ind, r) {
    var sem = SEMAFORO[r.real.semaforo];
    var arte = s('svg', { class: 'tarjeta-arte', viewBox: '0 0 360 360', 'aria-hidden': 'true' });
    dibujarPastel(arte, ind, r, { anillo: true });
    var valor = r.real.semaforo === 'SIN DATOS' ? '—' : (r.real.semaforo === 'META CUMPLIDA' ? '✓' : pct(r.real.suficiencia));
    return h('button', {
      class: 'tarjeta', type: 'button', 'aria-pressed': String(ind.id === estado.sel),
      'aria-label': nombreDe(ind) + '. ' + sem.txt + (r.brecha > 0 ? ', suficiencia ' + pct(r.real.suficiencia) : ''),
      data: { accion: 'seleccionar', id: ind.id }
    },
    arte,
    h('span', { class: 'tarjeta-sentido' }, ind.sentido ? icono(ind.sentido === 'sube' ? 'sube' : 'baja') : null, textoSentido(ind) + (ind.unidad ? ' · ' + ind.unidad : '')),
    h('span', { class: 'tarjeta-nombre' }, nombreDe(ind)),
    h('span', { class: 'tarjeta-pie' }, anilloMini(r.real.suficiencia, sem.tono), h('b', null, valor), sem.txt));
  }
  function pintarRiel() {
    var foco = document.activeElement && document.activeElement.classList && document.activeElement.classList.contains('tarjeta');
    D.riel.replaceChildren();
    estado.indicadores.forEach(function (ind) { D.riel.appendChild(tarjeta(ind, Motor.calcular(ind, estado.params))); });
    D.riel.appendChild(h('button', { class: 'tarjeta tarjeta--nueva', type: 'button', data: { accion: 'nuevo-indicador' } },
      icono('plus'), h('b', null, 'Nuevo indicador'), h('span', null, 'Empieza desde cero')));
    var k = estado.indicadores.length;
    D.rielCuenta.textContent = k + (k === 1 ? ' indicador' : ' indicadores');
    var sel = tarjetaDe(estado.sel);
    if (foco && sel) sel.focus({ preventScroll: true });
    if (sel) {
      var izq = sel.offsetLeft - D.riel.offsetLeft;
      if (izq < D.riel.scrollLeft || izq + sel.offsetWidth > D.riel.scrollLeft + D.riel.clientWidth) D.riel.scrollLeft = Math.max(0, izq - 24);
    }
    actualizarFlechas();
  }
  function tarjetaDe(id) { return $$('.tarjeta', D.riel).filter(function (t) { return t.dataset.id === id; })[0] || null; }
  function actualizarTarjeta(ind, r) {
    var vieja = tarjetaDe(ind.id);
    if (!vieja) return;
    var nueva = tarjeta(ind, r);
    var conFoco = document.activeElement === vieja;
    vieja.replaceWith(nueva);
    if (conFoco) nueva.focus({ preventScroll: true });
  }
  function actualizarFlechas() {
    var p = D.riel;
    $('[data-accion="riel-izq"]').disabled = p.scrollLeft <= 4;
    $('[data-accion="riel-der"]').disabled = p.scrollLeft + p.clientWidth >= p.scrollWidth - 4;
  }

  // ---------------------------------------------------------------- escenario (resultado)
  function pintarEscenario(ind, r, entrando) {
    D.titulo.textContent = nombreDe(ind);
    document.title = nombreDe(ind) + ' · Suficiencia 2027';
    if (entrando) $$('.respuesta, .pastel-zona', D.cartel).forEach(function (e) { e.classList.remove('aparece'); void e.offsetWidth; e.classList.add('aparece'); });

    var meta = [h('span', { class: 'pildora' }, ind.sentido ? icono(ind.sentido === 'sube' ? 'sube' : 'baja') : null, textoSentido(ind))];
    if (!vacio(ind.ejecutado)) meta.push(h('span', { class: 'pildora' }, 'Hoy ', h('b', null, valorConUnidad(ind.ejecutado, ind))));
    if (!vacio(ind.meta)) meta.push(h('span', { class: 'pildora' }, 'Meta 2027 ', h('b', null, valorConUnidad(ind.meta, ind))));
    if (r.brecha > 0) meta.push(h('span', { class: 'pildora' }, 'Brecha ', h('b', null, puntos(r.brecha, ind))));
    fijar(D.metaLinea, meta);

    fijar(D.frase, frase(ind, r));
    pintarKpis(ind, r);
    dibujarPastel(D.pastelSvg, ind, r, { anillo: true, plan: true, etiquetas: true, entrando: entrando });
    pintarCentro(r);
    D.pastelAlt.textContent = textoAlternativo(ind, r);
    pintarLeyenda(ind, r);
    pintarVigilar(ind, r);
    pintarAportes(ind, r);
  }

  function frase(ind, r) {
    var fuerte = function (t) { return h('strong', null, t); };
    if (vacio(ind.ejecutado) || vacio(ind.meta) || !ind.sentido) {
      return ['Completa el ', fuerte('paso 1'), ': escribe el valor de hoy, la meta 2027 y si es mejor que el indicador suba o que baje.'];
    }
    if (r.brecha === 0) return [fuerte('La meta ya se cumple hoy.'), ' No hay brecha que cerrar.'];
    if (r.chequeoAtribucion === 'SIN DOLORES') {
      return ['Te faltan ', fuerte(puntos(r.brecha, ind)), ' para la meta. Ahora cuenta qué causa el problema en el ', fuerte('paso 2'), '.'];
    }
    if (!r.ranking.length) return ['Ya tienes el pastel del problema. Agrega tus iniciativas en el ', fuerte('paso 3'), ' para saber si te alcanza.'];
    var R = r.real.suficiencia;
    var P = r.plan.suficiencia;
    var partes = ['Con la ejecución esperada cubres el ', fuerte(pct(R)), ' de la brecha. '];
    if (R >= 1) partes.push('Alcanza, incluso con los retrasos que tienes previstos.');
    else if (P >= 1) partes.push('Si todo se cumple a tiempo llegas al ', fuerte(pct(P)), ': los retrasos te cuestan ', fuerte(dec((P - R) * 100, 0) + ' puntos'), ' de suficiencia.');
    else partes.push('Aunque todo se cumpla al 100 % llegarías al ', fuerte(pct(P)), ': faltan iniciativas o más efectividad.');
    return partes;
  }

  function pintarKpis(ind, r) {
    var items;
    if (r.real.semaforo === 'SIN DATOS') {
      items = [['Si todo se cumple', '—', 'Falta el paso 1'], ['Así quedarías en 2027', '—', ''], ['Te faltaría', '—', '']];
    } else if (r.real.semaforo === 'META CUMPLIDA') {
      items = [['Hoy', valorConUnidad(ind.ejecutado, ind), ''], ['Meta 2027', valorConUnidad(ind.meta, ind), ''], ['Te faltaría', 'Nada', 'La meta ya se cumple']];
    } else {
      items = [
        ['Si todo se cumple', pct(r.plan.suficiencia), 'de la brecha, con 100 % de ejecución'],
        ['Así quedarías en 2027', valorConUnidad(r.real.proyectado, ind), 'Meta ' + valorConUnidad(ind.meta, ind)],
        ['Te faltaría', r.real.brechaRestante > 0 ? puntos(r.real.brechaRestante, ind) : 'Nada',
          r.real.brechaRestante > 0 ? 'de ' + puntos(r.brecha, ind) + ' de brecha' : 'Cierras toda la brecha']
      ];
    }
    fijar(D.kpis, items.map(function (it) {
      return h('div', { class: 'kpi' }, h('dt', null, it[0]), h('dd', null, it[1], it[2] ? h('span', { class: 'sub' }, it[2]) : null));
    }));
  }

  function pintarCentro(r) {
    var sem = SEMAFORO[r.real.semaforo];
    var valor;
    var etq;
    if (r.real.semaforo === 'SIN DATOS') { valor = '—'; etq = 'Completa el paso 1'; }
    else if (r.real.semaforo === 'META CUMPLIDA') { valor = '✓'; etq = 'No hay brecha que cerrar'; }
    else { valor = [nf(0, 0, true).format(r.real.suficiencia * 100), h('small', null, '%')]; etq = 'de la brecha cubierta'; }
    fijar(D.pastelCentro, h('span', { class: 'pastel-valor' }, valor), h('span', { class: 'pastel-etq' }, etq), chip(sem));
  }

  function textoAlternativo(ind, r) {
    if (r.real.semaforo === 'SIN DATOS') return 'Pastel del problema: todavía sin datos.';
    var partes = ind.dolores.map(function (d, j) {
      var dr = r.dolores[j];
      return (d.nombre || 'Dolor ' + (j + 1)) + ': ' + pct(dr.atribucion) + ' del problema, ' + pct(dr.resueltoReal) + ' resuelto';
    });
    return 'Pastel del problema con ' + ind.dolores.length + ' dolores. ' + partes.join('; ') +
      '. Suficiencia con la ejecución esperada: ' + pct(r.real.suficiencia) + ' (' + SEMAFORO[r.real.semaforo].txt + ').';
  }

  function pintarLeyenda(ind, r) {
    var items = [];
    ind.dolores.forEach(function (d, j) {
      var dr = r.dolores[j];
      if (dr.atribucion <= 0 && !d.nombre) return;
      var conPlan = dr.resueltoPlan - dr.resueltoReal > 0.004;
      items.push(h('li', { data: { dolor: d.id } },
        h('button', { class: 'ley-btn', type: 'button', data: { accion: 'ir', tipo: 'dolor', id: d.id }, 'aria-label': 'Ir al dolor ' + (j + 1) + ' en el taller' },
          h('span', { class: 'num-d', 'aria-hidden': 'true' }, String(j + 1)),
          h('span', { class: 'ley-nombre' }, d.nombre || 'Dolor sin nombre'),
          h('span', { class: 'ley-dato' }, pct(dr.resueltoReal)),
          h('span', { class: 'ley-barra', style: '--v:' + dr.resueltoReal.toFixed(4) + ';--p:' + dr.resueltoPlan.toFixed(4), 'aria-hidden': 'true' }, h('i'), conPlan ? h('b') : null),
          h('span', { class: 'ley-sub' }, pct(dr.atribucion) + ' del problema' + (conPlan ? ' · ' + pct(dr.resueltoPlan) + ' si todo se cumple' : '')))));
    });
    if (!items.length) items.push(h('li', { class: 'vacio-oscuro' }, 'Aún no hay dolores. Agrégalos en el paso 2.'));
    fijar(D.leyenda, items);
  }

  function alertas(ind, r) {
    var lista = [];
    var nom = function (t) { return h('strong', null, '«' + t + '»'); };
    var suf = r.real.suficiencia;
    if (vacio(ind.ejecutado) || vacio(ind.meta) || !ind.sentido) {
      lista.push({ tono: 'neutro', icono: 'info', texto: ['Completa el ', h('strong', null, 'paso 1'), ': hoy, meta 2027 y si es mejor que suba o que baje.'], ir: { tipo: 'paso', id: '1' } });
      return lista;
    }
    if (r.posiblePorcentaje) lista.push({ tono: 'alerta', icono: 'triangulo', texto: ['Revisa hoy y la meta: parece que escribiste el símbolo %. Si tu valor es 2,8 %, escribe solo 2,8.'], ir: { tipo: 'paso', id: '1' } });
    if (r.brecha === 0) return lista;
    if (r.chequeoAtribucion === 'SIN DOLORES') lista.push({ tono: 'neutro', icono: 'info', texto: ['Agrega los dolores y sus casos en el ', h('strong', null, 'paso 2'), '.'], ir: { tipo: 'paso', id: '2' } });
    if (r.chequeoAtribucion === 'NO SUMA 100') lista.push({ tono: 'critico', icono: 'triangulo', texto: ['La parte del pastel suma ', h('strong', null, pct(r.sumaAtribucion, 1)), '; debe sumar 100 %.'], ir: { tipo: 'paso', id: '2' } });
    ind.dolores.forEach(function (d, j) {
      var dr = r.dolores[j];
      if (dr.alerta !== 'SIN COBERTURA') return;
      lista.push({ tono: 'critico', icono: 'triangulo', dolor: d.id, texto: [nom(d.nombre), ' explica el ' + pct(dr.atribucion) + ' del problema y solo se resuelve el ' + pct(dr.resueltoReal) + '. Falta una iniciativa para este dolor.'], ir: { tipo: 'dolor', id: d.id } });
    });
    if (!r.ranking.length && r.chequeoAtribucion !== 'SIN DOLORES') lista.push({ tono: 'neutro', icono: 'info', texto: ['Agrega tus iniciativas en el ', h('strong', null, 'paso 3'), ' para saber si te alcanza.'], ir: { tipo: 'paso', id: '3' } });
    r.ranking.forEach(function (id) {
      var x = buscar(r.iniciativas, id);
      var ini = buscar(ind.iniciativas, id);
      if (x.etiqueta === 'EN RIESGO') {
        lista.push({ tono: 'alerta', icono: 'reloj', ini: id, texto: [nom(ini.nombre), ' está ' + (ini.estado === 'Retrasada' ? 'retrasada' : 'en riesgo') + '. Si falla, la suficiencia baja de ' + pct(suf) + ' a ' + pct(x.suficienciaSin) + '.'], ir: { tipo: 'ini', id: id } });
      }
    });
    r.ranking.forEach(function (id) {
      var x = buscar(r.iniciativas, id);
      var ini = buscar(ind.iniciativas, id);
      if (x.etiqueta === 'REEMPLAZAR') {
        lista.push({ tono: 'critico', icono: 'cambiar', ini: id, texto: n(ini.ejecucion) === 0
          ? [nom(ini.nombre), ' no se ejecutará (0 %). Cámbiala o reactívala.']
          : [nom(ini.nombre), ' aporta solo el ' + pct(x.aportePct) + ' de la brecha. Cámbiala o rediséñala.'], ir: { tipo: 'ini', id: id } });
      }
    });
    r.ranking.forEach(function (id) {
      var x = buscar(r.iniciativas, id);
      var ini = buscar(ind.iniciativas, id);
      if (x.revision === 'SIN RESPALDO') lista.push({ tono: 'alerta', icono: 'ayuda', ini: id, texto: [nom(ini.nombre), ' no tiene evidencia que respalde sus %.'], ir: { tipo: 'detalle', id: id } });
      if (x.revision === 'FALTA EJECUCION') lista.push({ tono: 'alerta', icono: 'ayuda', ini: id, texto: ['A ', nom(ini.nombre), ' le falta el % de ejecución.'], ir: { tipo: 'ini', id: id } });
    });
    return lista;
  }

  var verTodasAlertas = false;
  function pintarVigilar(ind, r) {
    var lista = alertas(ind, r);
    D.vigilarCuenta.textContent = lista.length ? lista.length + (lista.length === 1 ? ' alerta' : ' alertas') : '';
    if (!lista.length) {
      fijar(D.vigilar, h('li', { class: 'vig vig--ok' }, icono('check'), h('p', null, h('strong', null, 'Todo en orden.'), ' No hay alertas para este indicador.')));
      return;
    }
    var visibles = verTodasAlertas ? lista : lista.slice(0, 4);
    var items = visibles.map(function (a) {
      return h('li', { class: 'vig vig--' + a.tono, data: { dolor: a.dolor, ini: a.ini } },
        icono(a.icono), h('p', null, a.texto),
        h('button', { class: 'btn-chico', type: 'button', data: { accion: 'ir', tipo: a.ir.tipo, id: a.ir.id } }, 'Ver'));
    });
    if (lista.length > 4) items.push(h('li', null, h('button', { class: 'btn-texto', type: 'button', data: { accion: 'mas-alertas' } }, verTodasAlertas ? 'Ver menos' : 'Ver ' + (lista.length - 4) + ' más')));
    fijar(D.vigilar, items);
  }

  function pintarAportes(ind, r) {
    if (r.real.semaforo === 'SIN DATOS' || r.real.semaforo === 'META CUMPLIDA') {
      fijar(D.aportes, h('li', { class: 'vacio-oscuro' }, r.real.semaforo === 'SIN DATOS' ? 'Completa el paso 1 para ver cuánto aporta cada iniciativa.' : 'La meta ya se cumple: no hay brecha que repartir.'));
      return;
    }
    if (!r.ranking.length) { fijar(D.aportes, h('li', { class: 'vacio-oscuro' }, 'Aún no hay iniciativas. Agrégalas en el paso 3.')); return; }
    var max = Math.max.apply(null, r.ranking.map(function (id) { return buscar(r.iniciativas, id).aporte; }).concat([1e-9]));
    fijar(D.aportes, r.ranking.map(function (id) {
      var x = buscar(r.iniciativas, id);
      var ini = buscar(ind.iniciativas, id);
      return h('li', { data: { ini: id } },
        h('button', { class: 'ap-btn', type: 'button', data: { accion: 'ir', tipo: 'ini', id: id }, 'aria-label': ini.nombre + ': aporta ' + puntos(x.aporte, ind) + '. Ir a la iniciativa en el taller.' },
          h('span', { class: 'ap-nombre' }, ini.nombre),
          h('span', { class: 'ap-valor' }, puntos(x.aporte, ind), h('small', null, pct(x.aportePct))),
          h('span', { class: 'ap-barra', 'aria-hidden': 'true' }, h('i', { style: '--v:' + limitar(x.aporte / max, 0, 1).toFixed(4) })),
          h('span', { class: 'ap-pie' }, chip(ETIQUETA[x.etiqueta], true), h('span', null, 'Si falla, la suficiencia baja a ', h('b', null, pct(x.suficienciaSin))))));
    }));
  }

  function pintarBarraResumen(ind, r) {
    var sem = SEMAFORO[r.real.semaforo];
    fijar(D.barraResumen, anilloMini(r.real.suficiencia, sem.tono), h('span', { class: 'br-nombre' }, nombreDe(ind)),
      h('span', { class: 'br-valor' }, r.brecha > 0 ? pct(r.real.suficiencia) : '—'), chip(sem, true));
  }

  // ---------------------------------------------------------------- taller (edición)
  function botonAyuda(clave, sobre) {
    return h('button', { class: 'ayuda', type: 'button', data: { ayuda: clave }, 'aria-label': 'Qué significa ' + sobre }, '?');
  }
  function etq(texto, ayuda, paraId) {
    return h('div', { class: 'etq' }, h('label', { for: paraId }, texto), ayuda ? botonAyuda(ayuda, texto) : null);
  }
  function entrada(tipo, valor, props) {
    var p = { class: 'entrada', type: 'text', autocomplete: 'off', spellcheck: 'false' };
    Object.keys(props || {}).forEach(function (k) { p[k] = props[k]; });
    p.data = Object.assign({ tipo: tipo }, p.data || {});
    if (tipo === 'decimal' || tipo === 'pct') p.inputmode = 'decimal';
    if (tipo === 'entero') p.inputmode = 'numeric';
    var e = h('input', p);
    e.value = formatoEditable(tipo, valor);
    return e;
  }
  // Nombres largos: área de texto de una o dos líneas (Enter no crea saltos de línea).
  function entradaArea(valor, props) {
    var p = { class: 'entrada entrada--area', rows: '2', spellcheck: 'false' };
    Object.keys(props || {}).forEach(function (k) { p[k] = props[k]; });
    p.data = Object.assign({ tipo: 'texto' }, p.data || {});
    var e = h('textarea', p);
    e.value = vacio(valor) ? '' : String(valor);
    return e;
  }
  function conUnidad(input, texto, lista) {
    var u = h('span', { class: 'u', 'aria-hidden': 'true' }, texto);
    if (lista) lista.push(u);
    return h('div', { class: 'con-unidad' }, input, u);
  }
  function hoja(num, titulo, ayuda, id) {
    var numEl = h('span', { class: 'hoja-n', 'aria-hidden': 'true' }, String(num));
    var estadoEl = h('span', { class: 'hoja-estado' });
    var cuerpo = h('div');
    var sec = h('section', { class: 'hoja', id: id, 'aria-labelledby': id + '-t' },
      h('div', { class: 'hoja-cab' }, numEl, h('h3', { class: 'hoja-titulo', id: id + '-t' }, 'Paso ' + num + ' · ' + titulo), h('p', { class: 'hoja-ayuda' }, ayuda), estadoEl),
      cuerpo);
    return { sec: sec, cuerpo: cuerpo, estadoEl: estadoEl, numEl: numEl };
  }

  function pintarTaller() {
    var ind = actual();
    refs = { unidades: [], dolores: {}, inis: {}, cabDolores: [] };
    D.taller.replaceChildren(paso1(ind), paso2(ind), paso3(ind));
  }

  function paso1(ind) {
    var hj = hoja(1, '¿Cuánto me falta?', 'Dónde estás hoy, a dónde debes llegar en 2027 y cuál sería el mejor valor posible.', 'paso-1');
    var ids = { n: idUnico(), s: idUnico(), u: idUnico(), h: idUnico(), m: idUnico(), i: idUnico() };
    var segs = h('div', { class: 'segmentos' + (ind.sentido ? '' : ' sin-eleccion'), role: 'group', 'aria-labelledby': ids.s },
      h('button', { type: 'button', 'aria-pressed': String(ind.sentido === 'sube'), data: { accion: 'sentido', valor: 'sube' } }, icono('sube'), 'Que suba'),
      h('button', { type: 'button', 'aria-pressed': String(ind.sentido === 'baja'), data: { accion: 'sentido', valor: 'baja' } }, icono('baja'), 'Que baje'));
    refs.ideal = entrada('decimal', ind.ideal, { id: ids.i, data: { campo: 'ideal' } });
    refs.pista1 = h('p', { class: 'pista' });
    var unidad = ind.unidad || '';
    var campos = h('div', { class: 'campos' },
      h('div', { class: 'campo c-nombre' }, etq('Nombre del indicador', null, ids.n),
        entrada('texto', ind.nombre, { id: ids.n, placeholder: 'Ej.: Churn mensual pospago', maxlength: 120, data: { campo: 'nombre' } })),
      h('div', { class: 'campo c-sentido' }, h('div', { class: 'etq', id: ids.s }, '¿Qué es mejor para este indicador?'), segs),
      h('div', { class: 'campo c-unidad' }, etq('Unidad', null, ids.u),
        entrada('texto', ind.unidad, { id: ids.u, placeholder: '%', maxlength: 20, data: { campo: 'unidad' } })),
      h('div', { class: 'campo c-hoy' }, etq('Hoy', null, ids.h),
        conUnidad(entrada('decimal', ind.ejecutado, { id: ids.h, placeholder: 'Ej.: 2,8', data: { campo: 'ejecutado' } }), unidad, refs.unidades)),
      h('div', { class: 'campo c-meta' }, etq('Meta 2027', null, ids.m),
        conUnidad(entrada('decimal', ind.meta, { id: ids.m, placeholder: 'Ej.: 2,0', data: { campo: 'meta' } }), unidad, refs.unidades)),
      h('div', { class: 'campo c-ideal' }, etq('Valor ideal', 'ideal', ids.i), conUnidad(refs.ideal, unidad, refs.unidades)),
      h('div', { class: 'campo c-nombre' }, refs.pista1));
    refs.segmentos = segs;
    refs.bala = h('div', { class: 'bala' });
    hj.cuerpo.appendChild(h('div', { class: 'paso1' }, campos, refs.bala));
    refs.paso1 = hj;
    return hj.sec;
  }

  function paso2(ind) {
    var hj = hoja(2, '¿Qué causa el problema?', 'Los dolores salen del análisis de datos del indicador. Un caso puede tener varios dolores.', 'paso-2');
    var idC = idUnico();
    refs.notaPrev = h('p', { class: 'pista' });
    hj.cuerpo.appendChild(h('div', { class: 'casos-fila' },
      h('div', { class: 'campo' }, etq('Casos analizados', 'casos', idC),
        entrada('entero', ind.totalCasos, { id: idC, placeholder: 'Ej.: 12.000', data: { campo: 'totalCasos' } })),
      refs.notaPrev));
    var cab = h('div', { class: 'tabla-cab' },
      h('span', null, '#'), h('span', null, 'Dolor'), h('span', null, 'Casos'),
      h('span', null, 'Prevalencia', botonAyuda('prevalencia', 'prevalencia')),
      h('span', null, 'Parte del pastel', botonAyuda('atribucion', 'parte del pastel')),
      h('span', null, 'Resuelto', botonAyuda('resuelto', 'resuelto')),
      h('span', null, 'Fuente del dato'), h('span'));
    hj.cuerpo.appendChild(cab);
    ind.dolores.forEach(function (d, j) { hj.cuerpo.appendChild(filaDolor(d, j)); });
    if (!ind.dolores.length) hj.cuerpo.appendChild(h('p', { class: 'nota' }, icono('info'), h('span', null, 'Agrega el primer dolor: una causa del problema que salga del análisis de datos.')));
    refs.notaAtrib = h('p', { class: 'nota nota--alerta', hidden: true });
    hj.cuerpo.appendChild(refs.notaAtrib);
    var lleno = ind.dolores.length >= Motor.MAX_DOLORES;
    hj.cuerpo.appendChild(h('button', { class: 'btn-agregar', type: 'button', data: { accion: 'agregar-dolor' }, disabled: lleno },
      icono('plus'), lleno ? 'Máximo ' + Motor.MAX_DOLORES + ' dolores por indicador' : 'Agregar dolor'));
    refs.paso2 = hj;
    return hj.sec;
  }

  function filaDolor(d, j) {
    var ids = { n: idUnico(), c: idUnico(), a: idUnico(), f: idUnico() };
    var x = {
      prev: h('span', { class: 'valor-calc' }),
      atribIn: entrada('pct', d.atribucionManual, { id: ids.a, 'aria-label': 'Parte del pastel del dolor ' + (j + 1) + ' (vacío = automática)', data: { campo: 'atribucionManual', dolor: d.id } }),
      atribModo: h('span', { class: 'modo' }),
      atribBarra: h('span', { class: 'barra-fina barra-fina--rojo', 'aria-hidden': 'true' }, h('i')),
      resValor: h('span', { class: 'valor-calc' }),
      resChip: h('span'),
      resBarra: h('span', { class: 'barra-fina', 'aria-hidden': 'true' }, h('i'), h('b'))
    };
    refs.dolores[d.id] = x;
    return h('div', { class: 'dolor', data: { dolorFila: d.id } },
      h('div', { class: 'd-n' }, h('span', { class: 'num-claro', 'aria-hidden': 'true' }, String(j + 1))),
      h('div', { class: 'd-nombre' }, h('label', { class: 'sr', for: ids.n }, 'Nombre del dolor ' + (j + 1)),
        entradaArea(d.nombre, { id: ids.n, placeholder: 'Ej.: Precio / oferta de la competencia', maxlength: 120, data: { campo: 'nombre', dolor: d.id } })),
      h('div', { class: 'd-datos' },
        h('div', { class: 'd-casos' }, h('label', { class: 'et-movil', for: ids.c }, 'Casos'),
          entrada('entero', d.casos, { id: ids.c, placeholder: '0', 'aria-label': 'Casos con el dolor ' + (j + 1), data: { campo: 'casos', dolor: d.id } })),
        h('div', { class: 'd-prev' }, h('span', { class: 'et-movil' }, 'Prevalencia'), x.prev),
        h('div', { class: 'd-atrib' }, h('span', { class: 'et-movil' }, 'Parte del pastel'),
          h('div', { class: 'atrib-ctl' }, conUnidad(x.atribIn, '%'), x.atribModo), x.atribBarra),
        h('div', { class: 'd-res' }, h('span', { class: 'et-movil' }, 'Resuelto'),
          h('div', { class: 'res-ctl' }, h('div', { class: 'res-linea' }, x.resValor, x.resChip), x.resBarra))),
      h('div', { class: 'd-fuente' }, h('label', { class: 'et-movil', for: ids.f }, 'Fuente del dato'),
        entradaArea(d.evidencia, { id: ids.f, placeholder: '¿De dónde sale el dato?', maxlength: 300, 'aria-label': 'Fuente del dato del dolor ' + (j + 1), data: { campo: 'evidencia', dolor: d.id } })),
      h('div', { class: 'd-mas' }, h('button', { class: 'btn-icono peligro', type: 'button', data: { accion: 'borrar-dolor', id: d.id }, 'aria-label': 'Borrar el dolor ' + (j + 1) }, icono('borrar'))));
  }

  function paso3(ind) {
    var hj = hoja(3, '¿Qué vas a hacer?', 'En cada dolor escribe qué % resuelve la iniciativa si se hace completa (vacío = 0 %). Luego ajusta cuánto estará ejecutado en 2027.', 'paso-3');
    var matriz = h('div', { class: 'matriz', style: '--nd:' + Math.max(1, ind.dolores.length) });
    var cab = h('div', { class: 'matriz-cab' }, h('span', null, 'N°'), h('span', null, 'Iniciativa', botonAyuda('efectividad', 'efectividad')));
    if (ind.dolores.length) {
      ind.dolores.forEach(function (d, j) {
        var nombre = h('span', null, d.nombre || 'Dolor ' + (j + 1));
        refs.cabDolores.push({ id: d.id, j: j, el: nombre });
        cab.appendChild(h('span', { class: 'cab-dolor' }, h('span', { class: 'num-claro', 'aria-hidden': 'true' }, String(j + 1)), nombre));
      });
    } else {
      cab.appendChild(h('span', null, 'Dolores'));
    }
    cab.appendChild(h('span', null, 'Ejecución a 2027', botonAyuda('ejecucion', 'ejecución')));
    cab.appendChild(h('span', null, 'Estado'));
    cab.appendChild(h('span', null, 'Aporte esperado', botonAyuda('aporte', 'aporte')));
    cab.appendChild(h('span'));
    matriz.appendChild(cab);
    ind.iniciativas.forEach(function (x, i) { matriz.appendChild(filaIni(ind, x, i)); });
    hj.cuerpo.appendChild(matriz);
    if (!ind.iniciativas.length) hj.cuerpo.appendChild(h('p', { class: 'nota' }, icono('info'), h('span', null, 'Agrega la primera iniciativa: algo concreto que vas a hacer para resolver uno o varios dolores.')));
    var lleno = ind.iniciativas.length >= Motor.MAX_INICIATIVAS;
    hj.cuerpo.appendChild(h('button', { class: 'btn-agregar', type: 'button', data: { accion: 'agregar-ini' }, disabled: lleno },
      icono('plus'), lleno ? 'Máximo ' + Motor.MAX_INICIATIVAS + ' iniciativas por indicador (igual que la plantilla)' : 'Agregar iniciativa'));
    refs.paso3 = hj;
    return hj.sec;
  }

  function filaIni(ind, x, i) {
    var ids = { n: idUnico(), r: idUnico(), f: idUnico(), e: idUnico(), det: idUnico('det') };
    var q = { celdas: {}, nombresCelda: {} };
    var celdas = h('div', { class: 'i-celdas' });
    if (!ind.dolores.length) celdas.appendChild(h('div', { class: 'sin-dolores' }, 'Primero agrega dolores en el paso 2.'));
    ind.dolores.forEach(function (d, j) {
      var v = x.efectividad[d.id];
      var inp = entrada('pct', v, { class: '', placeholder: '—', 'aria-label': 'Efectividad de la iniciativa ' + (i + 1) + ' sobre el dolor ' + (j + 1), data: { campo: 'efectividad', ini: x.id, dolor: d.id } });
      var caja = h('div', { class: 'celda-caja', style: '--e:' + n(v) }, inp, h('span', { class: 'u', 'aria-hidden': 'true' }, '%'));
      var nombre = h('span', null, d.nombre || 'Dolor ' + (j + 1));
      q.celdas[d.id] = caja;
      q.nombresCelda[d.id] = nombre;
      celdas.appendChild(h('div', { class: 'celda' }, h('span', { class: 'celda-et', 'aria-hidden': 'true' }, h('span', { class: 'num-claro' }, String(j + 1)), nombre), caja));
    });
    var ejec = n(x.ejecucion);
    q.rango = h('input', {
      class: 'rango', type: 'range', min: '0', max: '100', step: '1', value: String(Math.round(ejec * 100)),
      style: '--v:' + (ejec * 100) + '%', 'aria-label': 'Ejecución a 2027 de la iniciativa ' + (i + 1), data: { campo: 'ejecucion', ini: x.id, tipo: 'rango' }
    });
    q.ejecIn = entrada('pct', x.ejecucion, { class: 'entrada', placeholder: '—', 'aria-label': 'Ejecución a 2027 de la iniciativa ' + (i + 1) + ' en porcentaje', data: { campo: 'ejecucion', ini: x.id } });
    var sel = h('select', { class: 'entrada entrada--chica', 'aria-label': 'Estado de la iniciativa ' + (i + 1), data: { campo: 'estado', ini: x.id, tipo: 'texto' } },
      Motor.ESTADOS.map(function (e) { return h('option', { value: e, selected: e === x.estado }, e); }));
    q.apValor = h('b');
    q.apPct = h('small');
    q.apBarra = h('span', { class: 'barra-fina', 'aria-hidden': 'true' }, h('i'));
    q.apChip = h('span');
    q.apFalla = h('span');
    q.revision = h('div', { class: 'revision' });
    var abierto = !!abiertos[x.id];
    q.detalle = h('div', { class: 'i-detalle', id: ids.det, hidden: !abierto },
      h('div', { class: 'campo' }, etq('Responsable', null, ids.r), entrada('texto', x.responsable, { id: ids.r, placeholder: 'Ej.: Marketing', maxlength: 80, data: { campo: 'responsable', ini: x.id } })),
      h('div', { class: 'campo' }, etq('Fecha comprometida', null, ids.f), h('input', { class: 'entrada', type: 'date', id: ids.f, value: x.fecha || '', data: { campo: 'fecha', ini: x.id, tipo: 'texto' } })),
      h('div', { class: 'campo c-evidencia' }, etq('Evidencia que respalda los %', null, ids.e),
        entrada('texto', x.evidencia, { id: ids.e, placeholder: 'Ej.: Piloto zona sur: −30 % de reclamos reiterados', maxlength: 300, data: { campo: 'evidencia', ini: x.id } })));
    refs.inis[x.id] = q;
    return h('div', { class: 'ini', data: { iniFila: x.id } },
      h('div', { class: 'i-n' }, h('span', { class: 'num-ini', 'aria-hidden': 'true' }, 'I' + (i + 1))),
      h('div', { class: 'i-nombre' }, h('label', { class: 'sr', for: ids.n }, 'Nombre de la iniciativa ' + (i + 1)),
        entradaArea(x.nombre, { id: ids.n, placeholder: 'Ej.: Resolución en el primer contacto', maxlength: 120, data: { campo: 'nombre', ini: x.id } }), q.revision),
      celdas,
      h('div', { class: 'i-ejec' }, h('span', { class: 'et-movil' }, 'Ejecución a 2027'), h('div', { class: 'ejec-ctl' }, q.rango, conUnidad(q.ejecIn, '%'))),
      h('div', { class: 'i-estado' }, h('span', { class: 'et-movil' }, 'Estado'), sel),
      h('div', { class: 'i-aporte' }, h('span', { class: 'et-movil' }, 'Aporte con la ejecución esperada'),
        h('div', { class: 'ap-linea' }, q.apValor, q.apPct), q.apBarra, h('div', { class: 'ap-mini' }, q.apChip, q.apFalla)),
      h('div', { class: 'i-mas' },
        h('button', { class: 'btn-icono', type: 'button', 'aria-expanded': String(abierto), 'aria-controls': ids.det, 'aria-label': 'Responsable, fecha y evidencia de la iniciativa ' + (i + 1), data: { accion: 'detalle-ini', id: x.id } }, icono('abajo', 'girar')),
        h('button', { class: 'btn-icono peligro', type: 'button', 'aria-label': 'Borrar la iniciativa ' + (i + 1), data: { accion: 'borrar-ini', id: x.id } }, icono('borrar'))),
      q.detalle);
  }

  // ---------------------------------------------------------------- salidas del taller
  function actualizarTaller(ind, r) {
    // paso 1
    refs.unidades.forEach(function (u) { u.textContent = ind.unidad || ''; });
    refs.ideal.placeholder = (ind.sentido === 'sube' ? '100' : '0') + ' (automático)';
    $$('button', refs.segmentos).forEach(function (b) { b.setAttribute('aria-pressed', String(b.dataset.valor === ind.sentido)); });
    refs.segmentos.classList.toggle('sin-eleccion', !ind.sentido);
    var pista = '';
    var alerta = false;
    if (!ind.sentido) { pista = 'Elige si es mejor que el indicador suba o que baje: sin eso no se puede calcular la brecha.'; alerta = true; }
    else if (r.posiblePorcentaje) { pista = '¿Escribiste el símbolo %? Si tu valor es 2,8 %, escribe solo 2,8.'; alerta = true; }
    refs.pista1.textContent = pista;
    refs.pista1.className = 'pista' + (alerta ? ' pista--alerta' : '');
    refs.pista1.hidden = !pista;
    pintarBala(ind, r);

    // paso 2
    ind.dolores.forEach(function (d, j) {
      var x = refs.dolores[d.id];
      var dr = r.dolores[j];
      if (!x) return;
      x.prev.textContent = pct(dr.prevalencia);
      x.atribIn.placeholder = editablePct(dr.atribucionAuto);
      x.atribModo.textContent = dr.manual ? 'manual' : 'auto';
      x.atribModo.classList.toggle('manual', dr.manual);
      x.atribBarra.style.setProperty('--v', dr.atribucion);
      x.resValor.textContent = pct(dr.resueltoReal);
      x.resBarra.style.setProperty('--v', dr.resueltoReal);
      x.resBarra.style.setProperty('--p', dr.resueltoPlan);
      x.resBarra.lastChild.hidden = !(dr.resueltoPlan - dr.resueltoReal > 0.004);
      fijar(x.resChip, chip(ALERTA_DOLOR[dr.alerta], true));
    });
    refs.notaPrev.textContent = r.sumaPrevalencia > 0
      ? 'La suma de prevalencias es ' + pct(r.sumaPrevalencia) + '. Puede pasar de 100 % porque un caso puede tener varios dolores: es normal.'
      : 'Escribe cuántos casos tiene cada dolor: con eso se calcula su parte del pastel.';
    var noSuma = r.chequeoAtribucion === 'NO SUMA 100';
    refs.notaAtrib.hidden = !noSuma;
    if (noSuma) fijar(refs.notaAtrib, icono('triangulo'), h('span', null, 'La parte del pastel suma ', h('b', null, pct(r.sumaAtribucion, 1)), ' y debe sumar 100 %. Si cambias una a mano, cámbialas todas, o borra las manuales para volver a la automática.'));

    // paso 3
    refs.cabDolores.forEach(function (c) { var d = buscar(ind.dolores, c.id); if (d) c.el.textContent = d.nombre || 'Dolor ' + (c.j + 1); });
    ind.iniciativas.forEach(function (x, i) {
      var q = refs.inis[x.id];
      var res = r.iniciativas[i];
      if (!q) return;
      var activa = r.brecha > 0 && !vacio(x.nombre);
      q.apValor.textContent = activa ? puntos(res.aporte, ind) : '—';
      q.apPct.textContent = activa ? pct(res.aportePct) + ' de la brecha' : '';
      q.apBarra.style.setProperty('--v', activa ? limitar(res.aportePct, 0, 1) : 0);
      fijar(q.apChip, activa ? chip(ETIQUETA[res.etiqueta], true) : null);
      fijar(q.apFalla, activa ? ['Si falla: ', h('b', null, pct(res.suficienciaSin))] : null);
      fijar(q.revision, res.revision && res.revision !== 'OK'
        ? h('button', { type: 'button', data: { accion: 'detalle-ini', id: x.id, abrir: '1' }, 'aria-label': REVISION[res.revision].txt + ': abrir los detalles' }, chip(REVISION[res.revision], true))
        : null);
      Object.keys(q.celdas).forEach(function (did) {
        q.celdas[did].style.setProperty('--e', n(x.efectividad[did]));
        var d = buscar(ind.dolores, did);
        if (d) q.nombresCelda[did].textContent = d.nombre || 'Dolor ' + (ind.dolores.indexOf(d) + 1);
      });
      var v = Math.round(n(x.ejecucion) * 100);
      if (document.activeElement !== q.rango) q.rango.value = String(v);
      q.rango.style.setProperty('--v', (n(x.ejecucion) * 100) + '%');
      if (document.activeElement !== q.ejecIn) q.ejecIn.value = editablePct(x.ejecucion);
    });
    pintarPasos(ind, r);
  }

  function pintarBala(ind, r) {
    var b = refs.bala;
    var partes = [h('p', { class: 'bala-titulo' }, 'Del hoy al ideal')];
    if (vacio(ind.ejecutado) || !ind.sentido) {
      partes.push(h('p', { class: 'bala-ayuda' }, 'Escribe el valor de hoy, la meta y elige si es mejor que suba o que baje. Aquí verás el tamaño del problema y de la brecha.'));
      fijar(b, partes);
      return;
    }
    if (r.problema === 0) {
      partes.push(h('p', { class: 'bala-ayuda' }, 'Hoy ya estás en el valor ideal: no hay problema que resolver.'));
      fijar(b, partes);
      return;
    }
    var w = function (v) { return limitar(v / r.problema, 0, 1); };
    var xMeta = w(r.brecha);
    var xReal = w(r.real.pp);
    var xPlan = w(r.plan.pp);
    var conPlan = xPlan > xReal + 0.002;
    partes.push(h('p', { class: 'bala-ayuda' }, 'Toda la barra es el problema total. Lo rojo es la brecha que debes cerrar; lo negro, lo que recuperan tus iniciativas.'));
    partes.push(h('div', { class: 'bala-extremos' },
      h('span', null, 'Hoy ', h('b', null, valorConUnidad(ind.ejecutado, ind))),
      h('span', null, 'Ideal ', h('b', null, valorConUnidad(r.idealUsado, ind)))));
    var pista = h('div', { class: 'bala-pista', 'aria-hidden': 'true' },
      r.brecha > 0 ? h('div', { class: 'bala-brecha', style: '--w:' + xMeta.toFixed(4) }) : null,
      conPlan ? h('div', { class: 'bala-plan', style: '--w:' + xPlan.toFixed(4) }) : null,
      xReal > 0 ? h('div', { class: 'bala-real', style: '--w:' + xReal.toFixed(4) }) : null,
      r.brecha > 0 ? h('div', { class: 'bala-meta', style: '--x:' + xMeta.toFixed(4) }) : null);
    var zona = h('div', { class: 'bala-zona' }, pista);
    if (r.brecha > 0) zona.appendChild(etiquetaBala('arriba', xMeta, ['Meta 2027 ', h('b', null, valorConUnidad(ind.meta, ind))]));
    else zona.appendChild(etiquetaBala('arriba', 0, [h('b', null, 'La meta ya se cumple')]));
    if (r.real.pp > 0 && r.brecha > 0) zona.appendChild(etiquetaBala('abajo', xReal, ['Con tus iniciativas: ', h('b', null, valorConUnidad(r.real.proyectado, ind))]));
    partes.push(zona);
    var ley = function (clase, texto) { return h('span', null, h('i', { class: clase }), texto); };
    partes.push(h('div', { class: 'bala-ley' },
      ley('k-real', 'Recuperas ' + puntos(r.real.pp, ind)),
      conPlan ? ley('k-plan', 'Si todo se cumple ' + puntos(r.plan.pp, ind)) : null,
      r.brecha > 0 ? ley('k-brecha', 'Brecha ' + puntos(r.brecha, ind)) : null,
      ley('k-pastel', 'Problema total ' + puntos(r.problema, ind))));
    fijar(b, partes);
  }
  function etiquetaBala(pos, x, contenido) {
    var al = x < 0.18 ? 'izq' : (x > 0.82 ? 'der' : 'centro');
    return h('span', { class: 'bala-et bala-et--' + pos + ' ' + al, style: 'left:' + (x * 100).toFixed(2) + '%' }, contenido);
  }

  function estadoPasos(ind, r) {
    var f1 = [];
    if (!ind.nombre) f1.push('el nombre');
    if (!ind.sentido) f1.push('qué es mejor');
    if (vacio(ind.ejecutado)) f1.push('hoy');
    if (vacio(ind.meta)) f1.push('la meta');
    var p2ok = ind.dolores.some(function (d) { return d.nombre; }) && n(ind.totalCasos) > 0 && r.chequeoAtribucion === 'OK';
    var falta2 = !ind.dolores.length ? 'Sin dolores' : (n(ind.totalCasos) === 0 ? 'Faltan los casos analizados' : (r.chequeoAtribucion !== 'OK' ? 'Revisa la parte del pastel' : 'Faltan nombres'));
    return [
      { ok: !f1.length, falta: 'Falta ' + unirLista(f1) },
      { ok: p2ok, falta: falta2 },
      { ok: r.chequeoIniciativas === 'OK', falta: r.chequeoIniciativas === 'SIN INICIATIVAS' ? 'Sin iniciativas' : 'Faltan datos' }
    ];
  }
  function pintarPasos(ind, r) {
    var p = estadoPasos(ind, r);
    var nombres = ['Brecha', 'Dolores', 'Iniciativas'];
    fijar(D.pasos, nombres.map(function (nom, k) {
      return h('li', null, h('button', { class: 'paso-chip' + (p[k].ok ? ' hecho' : ''), type: 'button', data: { accion: 'ir', tipo: 'paso', id: String(k + 1) } },
        h('span', { class: 'n', 'aria-hidden': 'true' }, p[k].ok ? '✓' : String(k + 1)), nom, h('small', null, p[k].ok ? 'listo' : p[k].falta)));
    }));
    [refs.paso1, refs.paso2, refs.paso3].forEach(function (hj, k) {
      hj.numEl.classList.toggle('hecho', p[k].ok);
      fijar(hj.estadoEl, p[k].ok ? chip({ txt: 'Listo', tono: 'ok', icono: 'check' }, true) : chip({ txt: p[k].falta, tono: 'neutro', icono: 'menos' }, true));
    });
  }

  // ---------------------------------------------------------------- ciclo de pintado
  function recalcular(entrando) {
    var ind = actual();
    var r = Motor.calcular(ind, estado.params);
    pintarEscenario(ind, r, entrando);
    actualizarTaller(ind, r);
    actualizarTarjeta(ind, r);
    pintarBarraResumen(ind, r);
    guardarLuego();
  }
  function pintarTodo(entrando) {
    verTodasAlertas = false;
    pintarRiel();
    pintarTaller();
    pintarAjustes();
    recalcular(entrando);
  }

  // ---------------------------------------------------------------- edición
  function manejarEntrada(e) {
    var t = e.target;
    var campo = t.dataset ? t.dataset.campo : null;
    if (!campo) return;
    var ind = actual();
    var tipo = t.dataset.tipo || 'texto';
    var v;
    if (tipo === 'rango') v = Number(t.value) / 100;
    else if (tipo === 'texto') v = t.tagName === 'TEXTAREA' ? t.value.replace(/\s*[\r\n]+\s*/g, ' ') : t.value;
    else {
      v = tipo === 'entero' ? leerEntero(t.value) : leerNumero(t.value);
      var malo = Number.isNaN(v) || (tipo === 'pct' && v !== null && (v < 0 || v > 100));
      if (malo) { t.setAttribute('aria-invalid', 'true'); return; }
      if (tipo === 'pct' && v !== null) v = v / 100;
    }
    t.removeAttribute('aria-invalid');
    if (campo === 'fecha' && v && !/^\d{4}-\d{2}-\d{2}$/.test(v)) v = '';
    if (t.dataset.ini) {
      var ini = buscar(ind.iniciativas, t.dataset.ini);
      if (!ini) return;
      if (campo === 'efectividad') {
        if (v === null) delete ini.efectividad[t.dataset.dolor];
        else ini.efectividad[t.dataset.dolor] = v;
      } else ini[campo] = v;
    } else if (t.dataset.dolor) {
      var d = buscar(ind.dolores, t.dataset.dolor);
      if (!d) return;
      d[campo] = v;
    } else ind[campo] = v;
    recalcular(false);
  }
  function alSalir(e) {
    var t = e.target;
    if (!t.dataset || !t.dataset.campo) return;
    var tipo = t.dataset.tipo || 'texto';
    if (t.getAttribute('aria-invalid') === 'true') return;
    var ind = actual();
    var obj = t.dataset.ini ? buscar(ind.iniciativas, t.dataset.ini) : (t.dataset.dolor ? buscar(ind.dolores, t.dataset.dolor) : ind);
    if (!obj) return;
    if (tipo === 'texto' && (t.tagName === 'TEXTAREA' || (t.tagName === 'INPUT' && t.type === 'text'))) {
      var limpio = t.value.replace(/\s*[\r\n]+\s*/g, ' ').trim();
      if (limpio !== t.value) { t.value = limpio; obj[t.dataset.campo] = limpio; recalcular(false); }
      return;
    }
    if (tipo === 'rango' || tipo === 'texto') return;
    var valor = t.dataset.campo === 'efectividad' ? obj.efectividad[t.dataset.dolor] : obj[t.dataset.campo];
    t.value = formatoEditable(tipo, valor);
  }

  // ---------------------------------------------------------------- acciones
  function nuevoIndicador() {
    var ind = Motor.nuevoIndicador();
    estado.indicadores.push(ind);
    estado.sel = ind.id;
    abiertos = {};
    pintarTodo(true);
    var p1 = $('#paso-1');
    irA(p1, p1 ? $('input', p1) : null);
    avisar('Creaste un indicador nuevo. Empieza por el paso 1.');
  }
  function duplicar() {
    var ind = actual();
    var copia = JSON.parse(JSON.stringify(ind));
    copia.id = Motor.nuevoId('ind-');
    copia.nombre = (ind.nombre || 'Indicador') + ' (copia)';
    copia.iniciativas.forEach(function (x) { x.id = Motor.nuevoId('i-'); });
    estado.indicadores.splice(estado.indicadores.indexOf(ind) + 1, 0, copia);
    estado.sel = copia.id;
    pintarTodo(true);
    avisar('Duplicaste «' + nombreDe(ind) + '».');
  }
  function eliminar() {
    var antes = foto();
    var ind = actual();
    var k = estado.indicadores.indexOf(ind);
    estado.indicadores.splice(k, 1);
    if (!estado.indicadores.length) estado.indicadores.push(Motor.nuevoIndicador());
    estado.sel = estado.indicadores[Math.min(k, estado.indicadores.length - 1)].id;
    pintarTodo(true);
    conDeshacer('Eliminaste «' + nombreDe(ind) + '».', antes);
  }
  function agregarDolor() {
    var ind = actual();
    if (ind.dolores.length >= Motor.MAX_DOLORES) return;
    var d = Motor.nuevoDolor();
    ind.dolores.push(d);
    pintarTaller();
    recalcular(false);
    var fila = filaPorDato('dolorFila', d.id);
    if (fila) irA(fila, $('input', fila));
  }
  function borrarDolor(id) {
    var antes = foto();
    var ind = actual();
    var d = buscar(ind.dolores, id);
    if (!d) return;
    ind.dolores.splice(ind.dolores.indexOf(d), 1);
    ind.iniciativas.forEach(function (x) { delete x.efectividad[id]; });
    pintarTaller();
    recalcular(false);
    conDeshacer('Borraste el dolor «' + (d.nombre || 'sin nombre') + '».', antes);
  }
  function agregarIni() {
    var ind = actual();
    if (ind.iniciativas.length >= Motor.MAX_INICIATIVAS) return;
    var x = Motor.nuevaIniciativa();
    ind.iniciativas.push(x);
    pintarTaller();
    recalcular(false);
    var fila = filaPorDato('iniFila', x.id);
    if (fila) irA(fila, $('input', fila));
  }
  function borrarIni(id) {
    var antes = foto();
    var ind = actual();
    var x = buscar(ind.iniciativas, id);
    if (!x) return;
    ind.iniciativas.splice(ind.iniciativas.indexOf(x), 1);
    delete abiertos[id];
    pintarTaller();
    recalcular(false);
    conDeshacer('Borraste la iniciativa «' + (x.nombre || 'sin nombre') + '».', antes);
  }
  function alternarDetalle(id, forzarAbrir) {
    var q = refs.inis[id];
    if (!q) return;
    var abrir = forzarAbrir ? true : q.detalle.hidden;
    q.detalle.hidden = !abrir;
    abiertos[id] = abrir;
    var btn = $$('[data-accion="detalle-ini"][aria-controls]').filter(function (b) { return b.dataset.id === id; })[0];
    if (btn) btn.setAttribute('aria-expanded', String(abrir));
    return q.detalle;
  }
  function filaPorDato(clave, id) {
    return $$('[data-' + (clave === 'dolorFila' ? 'dolor-fila' : 'ini-fila') + ']', D.taller).filter(function (e) { return e.dataset[clave] === id; })[0] || null;
  }
  function irA(el, enfocarEl) {
    if (!el) return;
    el.scrollIntoView({ behavior: movimientoReducido.matches ? 'auto' : 'smooth', block: 'center' });
    el.classList.remove('destello');
    void el.offsetWidth;
    el.classList.add('destello');
    if (enfocarEl) setTimeout(function () { enfocarEl.focus({ preventScroll: true }); }, movimientoReducido.matches ? 0 : 450);
  }
  function irADestino(tipo, id) {
    if (tipo === 'paso') { var sec = $('#paso-' + id); irA(sec, sec ? $('input, button', sec) : null); return; }
    if (tipo === 'dolor') { var fd = filaPorDato('dolorFila', id); irA(fd, fd ? $('input', fd) : null); return; }
    if (tipo === 'ini' || tipo === 'detalle') {
      var fi = filaPorDato('iniFila', id);
      if (!fi) return;
      if (tipo === 'detalle') { var det = alternarDetalle(id, true); irA(fi, det ? $('.c-evidencia input', det) : null); return; }
      irA(fi, $('input', fi));
    }
  }

  function conDeshacer(texto, antes) { deshacer = antes; avisar(texto, 'Deshacer'); }
  function avisar(texto, boton) {
    clearTimeout(tAviso);
    if (!boton) deshacer = null;
    fijar(D.aviso, h('span', null, texto), boton ? h('button', { type: 'button', data: { accion: 'deshacer' } }, boton) : null);
    D.aviso.classList.add('ver');
    tAviso = setTimeout(function () { D.aviso.classList.remove('ver'); deshacer = null; }, boton ? 8000 : 4500);
  }

  // menús
  function cerrarMenus(excepto) {
    $$('.menu').forEach(function (m) {
      if (m === excepto) return;
      m.hidden = true;
      var b = $('[aria-controls="' + m.id + '"]');
      if (b) b.setAttribute('aria-expanded', 'false');
    });
  }
  function alternarMenu(btn) {
    var m = document.getElementById(btn.getAttribute('aria-controls'));
    var abrir = m.hidden;
    cerrarMenus();
    m.hidden = !abrir;
    btn.setAttribute('aria-expanded', String(abrir));
    if (abrir) { var primero = $('button', m); if (primero) primero.focus(); }
  }

  // ---------------------------------------------------------------- foco cruzado pastel ↔ iniciativas
  function enfocar(tipo, id) {
    var ind = actual();
    var dol = {};
    var inis = {};
    if (tipo === 'dolor') {
      dol[id] = true;
      ind.iniciativas.forEach(function (x) { if (n(x.efectividad[id]) > 0) inis[x.id] = true; });
    } else {
      inis[id] = true;
      var ini = buscar(ind.iniciativas, id);
      if (ini) Object.keys(ini.efectividad).forEach(function (k) { if (n(ini.efectividad[k]) > 0) dol[k] = true; });
    }
    D.escenario.classList.add('hay-foco');
    $$('.porcion', D.pastelSvg).forEach(function (g) { g.classList.toggle('es-foco', !!dol[g.getAttribute('data-dolor')]); });
    $$('#leyenda > li').forEach(function (li) { li.classList.toggle('es-foco', !!dol[li.dataset.dolor]); });
    $$('#aportes > li').forEach(function (li) { li.classList.toggle('es-foco', !!inis[li.dataset.ini]); });
  }
  function desenfocar() {
    D.escenario.classList.remove('hay-foco');
    $$('.es-foco', D.escenario).forEach(function (e) { e.classList.remove('es-foco'); });
  }

  // tooltip
  function mostrarTip(nodos, x, y) {
    fijar(D.tip, nodos);
    D.tip.hidden = false;
    var w = D.tip.offsetWidth;
    var alto = D.tip.offsetHeight;
    var izq = x + 16;
    var arriba = y + 16;
    if (izq + w > window.innerWidth - 8) izq = x - w - 16;
    if (arriba + alto > window.innerHeight - 8) arriba = y - alto - 16;
    D.tip.style.left = Math.max(8, izq) + 'px';
    D.tip.style.top = Math.max(8, arriba) + 'px';
  }
  function tipJunto(el, nodos) { var b = el.getBoundingClientRect(); mostrarTip(nodos, b.left - 8, b.bottom - 8); }
  function ocultarTip() { D.tip.hidden = true; }
  function tipDolor(id) {
    var ind = actual();
    var r = Motor.calcular(ind, estado.params);
    var j = -1;
    ind.dolores.forEach(function (d, k) { if (d.id === id) j = k; });
    if (j < 0) return null;
    var d = ind.dolores[j];
    var dr = r.dolores[j];
    return [
      h('span', { class: 'tip-titulo' }, h('span', { class: 'num-d' }, String(j + 1)), d.nombre || 'Dolor sin nombre'),
      h('span', { class: 'tip-valor' }, pct(dr.resueltoReal) + ' resuelto'),
      h('span', { class: 'tip-linea' }, 'Si todo se cumple: ', h('b', null, pct(dr.resueltoPlan))),
      h('span', { class: 'tip-linea' }, 'Parte del pastel: ', h('b', null, pct(dr.atribucion, 1))),
      h('span', { class: 'tip-linea' }, 'Recupera: ', h('b', null, puntos(dr.ppReal, ind)))
    ];
  }

  // ---------------------------------------------------------------- ajustes (parámetros)
  var PARAMS_UI = [
    { k: 'verde', t: 'Verde (alcanza) desde', get: function (p) { return p.semaforo[2].desde; }, set: function (p, v) { p.semaforo[2].desde = v; } },
    { k: 'amarillo', t: 'Amarillo (cerca) desde', get: function (p) { return p.semaforo[1].desde; }, set: function (p, v) { p.semaforo[1].desde = v; } },
    { k: 'umbralReemplazar', t: 'Reemplazar si aporta menos de (% de la brecha)' },
    { k: 'coberturaAtribMin', t: 'Sin cobertura: dolor con parte del pastel mayor a' },
    { k: 'coberturaResueltoMin', t: '… y que queda resuelto menos de' },
    { k: 'tope', t: 'Tope de efectividad real (ninguna iniciativa resuelve el 100 %)' }
  ];
  function leerParam(def, p) { return def.get ? def.get(p) : p[def.k]; }
  function pintarAjustes() {
    fijar(D.ajustes, PARAMS_UI.map(function (def) {
      var id = idUnico('p');
      return h('div', { class: 'campo' }, h('div', { class: 'etq' }, h('label', { for: id }, def.t)),
        conUnidad(entrada('pct', leerParam(def, estado.params), { id: id, data: { param: def.k } }), '%'));
    }));
  }
  function manejarParam(e) {
    var t = e.target;
    if (!t.dataset || !t.dataset.param) return;
    var v = leerNumero(t.value);
    var def = PARAMS_UI.filter(function (x) { return x.k === t.dataset.param; })[0];
    var p = Motor.copiarParams(estado.params);
    var ok = v !== null && !Number.isNaN(v) && v >= 0 && v <= 100;
    if (ok) {
      v = v / 100;
      if (def.set) def.set(p, v); else p[def.k] = v;
      ok = p.semaforo[0].desde <= p.semaforo[1].desde && p.semaforo[1].desde <= p.semaforo[2].desde && p.tope > 0 && p.tope <= 0.999;
    }
    if (!ok) { t.setAttribute('aria-invalid', 'true'); return; }
    t.removeAttribute('aria-invalid');
    estado.params = p;
    pintarRiel();
    recalcular(false);
  }

  // ---------------------------------------------------------------- importar y exportar
  function exportar() {
    var datos = { app: 'Suficiencia 2027', version: 1, exportado: new Date().toISOString(), sel: estado.sel, params: estado.params, indicadores: estado.indicadores };
    var blob = new Blob([JSON.stringify(datos, null, 2)], { type: 'application/json' });
    var url = URL.createObjectURL(blob);
    var hoy = new Date();
    var a = h('a', { href: url, download: 'suficiencia-' + hoy.getFullYear() + '-' + dos(hoy.getMonth() + 1) + '-' + dos(hoy.getDate()) + '.json' });
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
    avisar('Guardaste una copia con ' + estado.indicadores.length + ' indicadores.');
  }

  function cargarSheetJS() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    return new Promise(function (ok, mal) {
      var sc = document.createElement('script');
      sc.src = SHEETJS;
      sc.onload = function () { if (window.XLSX) ok(window.XLSX); else mal(new Error('sin XLSX')); };
      sc.onerror = function () { mal(new Error('red')); };
      document.head.appendChild(sc);
    });
  }

  function fechaExcel(celda, XLSX) {
    if (!celda) return '';
    if (typeof celda.v === 'number' && XLSX.SSF && XLSX.SSF.parse_date_code) {
      var p = XLSX.SSF.parse_date_code(celda.v);
      if (p && p.y) return p.y + '-' + dos(p.m) + '-' + dos(p.d);
    }
    if (typeof celda.v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(celda.v)) return celda.v.slice(0, 10);
    return '';
  }

  // Lee las hojas IND… de Plantilla_Suficiencia.xlsx (mismas celdas que la plantilla).
  function leerPlantilla(XLSX, libro) {
    var inds = [];
    var columnas = ['C', 'D', 'E', 'F', 'G', 'H'];
    libro.SheetNames.forEach(function (nombreHoja) {
      if (!/^IND/i.test(nombreHoja)) return;
      var ws = libro.Sheets[nombreHoja];
      var celda = function (a) { return ws[a]; };
      var num = function (a) { var c = celda(a); return c && typeof c.v === 'number' && isFinite(c.v) ? c.v : null; };
      var txt = function (a) { var c = celda(a); return c && c.v !== undefined && c.v !== null ? String(c.v).trim() : ''; };
      var nombre = txt('C5');
      if (!nombre) return;
      var ind = {
        nombre: nombre,
        sentido: /sube/i.test(txt('C6')) ? 'sube' : (/baja/i.test(txt('C6')) ? 'baja' : ''),
        unidad: txt('C7'), ejecutado: num('C8'), meta: num('C9'), ideal: num('C10'), totalCasos: num('C13'),
        dolores: [], iniciativas: []
      };
      var idDe = {};
      columnas.forEach(function (c, j) {
        var dn = txt(c + '17');
        var casos = num(c + '18');
        var manual = num(c + '21');
        if (!dn && casos === null && manual === null) return;
        idDe[c] = 'd' + (j + 1);
        ind.dolores.push({ id: idDe[c], nombre: dn, casos: casos, atribucionManual: manual, evidencia: txt(c + '23') });
      });
      for (var fila = 29; fila <= 36; fila++) {
        var efect = {};
        var alguna = false;
        columnas.forEach(function (c) { var v = num(c + fila); if (v !== null && idDe[c]) { efect[idDe[c]] = v; alguna = true; } });
        var inom = txt('B' + fila);
        if (!inom && !alguna) continue;
        ind.iniciativas.push({
          id: 'i' + (fila - 28), nombre: inom, efectividad: efect, ejecucion: num('I' + fila),
          estado: txt('J' + fila) || 'En plan', responsable: txt('K' + fila), fecha: fechaExcel(celda('L' + fila), XLSX), evidencia: txt('M' + fila)
        });
      }
      inds.push(Motor.normalizarIndicador(ind));
    });
    var params = null;
    var wp = libro.Sheets.Parametros;
    if (wp) {
      var pn = function (a) { var c = wp[a]; return c && typeof c.v === 'number' ? c.v : null; };
      params = Motor.normalizarParams({
        semaforo: [{ desde: pn('B6') }, { desde: pn('B7') }, { desde: pn('B8') }],
        umbralReemplazar: pn('C12'), coberturaAtribMin: pn('C13'), coberturaResueltoMin: pn('C14'), tope: pn('C15')
      });
    }
    return { indicadores: inds, params: params };
  }

  function incorporar(inds, params, nombreArchivo) {
    var antes = foto();
    var nuevos = 0;
    var actualizados = 0;
    inds.forEach(function (ind) {
      var clave = ind.nombre.trim().toLowerCase();
      var existente = estado.indicadores.filter(function (x) { return x.nombre.trim().toLowerCase() === clave; })[0];
      if (existente) { ind.id = existente.id; estado.indicadores[estado.indicadores.indexOf(existente)] = ind; actualizados++; }
      else { ind.id = Motor.nuevoId('ind-'); estado.indicadores.push(ind); nuevos++; }
    });
    if (params) estado.params = params;
    estado.sel = inds[0].id;
    pintarTodo(true);
    var partes = [];
    if (nuevos) partes.push(nuevos + (nuevos === 1 ? ' indicador nuevo' : ' indicadores nuevos'));
    if (actualizados) partes.push(actualizados + (actualizados === 1 ? ' actualizado' : ' actualizados'));
    conDeshacer('Importaste ' + partes.join(' y ') + ' desde ' + nombreArchivo + '.', antes);
  }

  function importar(archivo) {
    if (!archivo) return;
    var nombre = archivo.name;
    if (/\.json$/i.test(nombre)) {
      archivo.text().then(function (texto) {
        var datos = Motor.normalizarDatos(JSON.parse(texto));
        if (!datos || !datos.indicadores.length) throw new Error('vacío');
        var antes = foto();
        estado.indicadores = datos.indicadores;
        estado.params = datos.params;
        estado.sel = datos.sel && datos.indicadores.some(function (x) { return x.id === datos.sel; }) ? datos.sel : datos.indicadores[0].id;
        pintarTodo(true);
        conDeshacer('Abriste la copia ' + nombre + ' (' + datos.indicadores.length + ' indicadores).', antes);
      }).catch(function () {
        avisar('No se pudo abrir ' + nombre + ': no es una copia de esta calculadora. Usa un archivo creado con «Guardar una copia».');
      });
      return;
    }
    avisar('Leyendo ' + nombre + '…');
    cargarSheetJS().then(function (XLSX) {
      return archivo.arrayBuffer().then(function (buf) {
        var res = leerPlantilla(XLSX, XLSX.read(buf, { type: 'array' }));
        if (!res.indicadores.length) {
          avisar('No encontré hojas de indicador (IND…) con nombre en ' + nombre + '. Usa la plantilla Plantilla_Suficiencia.xlsx.');
          return;
        }
        incorporar(res.indicadores, res.params, nombre);
      });
    }).catch(function (err) {
      avisar(err && err.message === 'red'
        ? 'No se pudo cargar el lector de Excel. Revisa tu conexión a internet e intenta de nuevo.'
        : 'No se pudo leer ' + nombre + '. Revisa que sea la plantilla de suficiencia y que no esté dañada.');
    });
  }

  function volverAEjemplos() {
    var antes = foto();
    var d = datosIniciales();
    estado.indicadores = d.indicadores;
    estado.params = d.params;
    estado.sel = d.sel;
    pintarTodo(true);
    conDeshacer('Volviste a los indicadores de ejemplo.', antes);
  }

  // ---------------------------------------------------------------- "Cómo funciona"
  function pintarDemos() {
    var base = {
      sentido: 'baja', unidad: '%', ejecutado: 10, meta: 6, ideal: 0, totalCasos: 100,
      dolores: [{ id: 'a', nombre: 'A', casos: 50 }, { id: 'b', nombre: 'B', casos: 30 }, { id: 'c', nombre: 'C', casos: 20 }],
      iniciativas: []
    };
    var entero = Object.assign({}, base, { dolores: [{ id: 't', nombre: 'Todo', casos: 100 }] });
    var conIni = Object.assign({}, base, {
      iniciativas: [
        { id: 'x', nombre: 'X', efectividad: { a: 0.6, c: 0.3 }, ejecucion: 0.7, estado: 'En plan', evidencia: 'e' },
        { id: 'y', nombre: 'Y', efectividad: { b: 0.5, a: 0.2 }, ejecucion: 1, estado: 'En plan', evidencia: 'e' }
      ]
    });
    var casos = { 1: [entero, {}], 2: [base, { etiquetas: true }], 3: [conIni, { plan: true, etiquetas: true }], 4: [conIni, { plan: true, anillo: true }] };
    $$('[data-demo]', D.dlg).forEach(function (svg) {
      var c = casos[svg.dataset.demo];
      dibujarPastel(svg, c[0], Motor.calcular(c[0], estado.params), c[1]);
    });
  }
  function pintarEtiquetasLista() {
    var filas = [];
    [SEMAFORO.VERDE, SEMAFORO.AMARILLO, SEMAFORO.ROJO].forEach(function (m, k) {
      var desde = estado.params.semaforo[2 - k].desde;
      filas.push([m, k === 0 ? 'Suficiencia de ' + pct(desde) + ' o más: las iniciativas alcanzan.' : (k === 1 ? 'Desde ' + pct(desde) + ': cerca, con ajustes puede alcanzar.' : 'Menos de ' + pct(estado.params.semaforo[1].desde) + ': faltan iniciativas o más efectividad.')]);
    });
    Object.keys(ETIQUETA).forEach(function (k) { filas.push([ETIQUETA[k], ETIQUETA[k].que]); });
    filas.push([ALERTA_DOLOR['SIN COBERTURA'], ALERTA_DOLOR['SIN COBERTURA'].que]);
    filas.push([REVISION['SIN RESPALDO'], REVISION['SIN RESPALDO'].que]);
    fijar($('#etiquetas-lista'), filas.map(function (f) { return h('li', null, h('span', null, chip(f[0])), h('span', null, f[1])); }));
  }
  function abrirComo() {
    pintarDemos();
    pintarEtiquetasLista();
    cerrarIntro();
    if (typeof D.dlg.showModal === 'function') D.dlg.showModal();
    else D.dlg.setAttribute('open', '');
  }
  function cerrarDialogo() { if (typeof D.dlg.close === 'function') D.dlg.close(); else D.dlg.removeAttribute('open'); }
  function cerrarIntro() {
    D.intro.hidden = true;
    try { window.localStorage.setItem(CLAVE_INTRO, '1'); } catch (e) { /* nada */ }
  }

  // ---------------------------------------------------------------- eventos
  document.addEventListener('click', function (e) {
    var b = e.target.closest('[data-accion]');
    if (!e.target.closest('.menu-envoltura')) cerrarMenus();
    if (e.target.closest('.ayuda')) return;
    if (!b) return;
    var acc = b.dataset.accion;
    if (b.closest('.menu')) cerrarMenus();
    switch (acc) {
      case 'menu': alternarMenu(b); break;
      case 'seleccionar':
        if (b.dataset.id !== estado.sel) { estado.sel = b.dataset.id; abiertos = {}; pintarTodo(true); }
        break;
      case 'nuevo-indicador': nuevoIndicador(); break;
      case 'duplicar': duplicar(); break;
      case 'eliminar': eliminar(); break;
      case 'sentido': actual().sentido = b.dataset.valor; recalcular(false); break;
      case 'agregar-dolor': agregarDolor(); break;
      case 'borrar-dolor': borrarDolor(b.dataset.id); break;
      case 'agregar-ini': agregarIni(); break;
      case 'borrar-ini': borrarIni(b.dataset.id); break;
      case 'detalle-ini': alternarDetalle(b.dataset.id, b.dataset.abrir === '1'); break;
      case 'ir': irADestino(b.dataset.tipo, b.dataset.id); break;
      case 'mas-alertas': verTodasAlertas = !verTodasAlertas; recalcular(false); break;
      case 'riel-izq': D.riel.scrollBy({ left: -D.riel.clientWidth * 0.8 }); break;
      case 'riel-der': D.riel.scrollBy({ left: D.riel.clientWidth * 0.8 }); break;
      case 'como-funciona': abrirComo(); break;
      case 'cerrar-dialogo': cerrarDialogo(); break;
      case 'cerrar-intro': cerrarIntro(); break;
      case 'importar': D.archivo.value = ''; D.archivo.click(); break;
      case 'exportar': exportar(); break;
      case 'imprimir': window.print(); break;
      case 'ejemplos': volverAEjemplos(); break;
      case 'deshacer':
        if (deshacer) { var f = deshacer; deshacer = null; volverA(f); avisar('Listo, se deshizo el cambio.'); }
        break;
    }
  });

  D.taller.addEventListener('input', manejarEntrada);
  D.taller.addEventListener('change', manejarEntrada);
  D.taller.addEventListener('focusout', alSalir);
  D.ajustes.addEventListener('input', manejarParam);
  D.ajustes.addEventListener('focusout', function (e) {
    var t = e.target;
    if (!t.dataset || !t.dataset.param || t.getAttribute('aria-invalid') === 'true') return;
    var def = PARAMS_UI.filter(function (x) { return x.k === t.dataset.param; })[0];
    t.value = editablePct(leerParam(def, estado.params));
  });
  D.archivo.addEventListener('change', function () { importar(D.archivo.files && D.archivo.files[0]); });
  D.riel.addEventListener('scroll', function () { window.requestAnimationFrame(actualizarFlechas); }, { passive: true });
  window.addEventListener('resize', actualizarFlechas);

  // ayudas "?"
  function mostrarAyuda(btn) { tipJunto(btn, [h('span', { class: 'tip-linea' }, AYUDAS[btn.dataset.ayuda] || '')]); }
  document.addEventListener('pointerover', function (e) { var b = e.target.closest && e.target.closest('.ayuda'); if (b) mostrarAyuda(b); });
  document.addEventListener('pointerout', function (e) { if (e.target.closest && e.target.closest('.ayuda')) ocultarTip(); });
  document.addEventListener('focusin', function (e) { if (e.target.classList && e.target.classList.contains('ayuda')) mostrarAyuda(e.target); });
  document.addEventListener('focusout', function (e) { if (e.target.classList && e.target.classList.contains('ayuda')) ocultarTip(); });
  document.addEventListener('click', function (e) { var b = e.target.closest('.ayuda'); if (b) mostrarAyuda(b); });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') { ocultarTip(); cerrarMenus(); desenfocar(); }
    if (e.key === 'Enter' && e.target.tagName === 'TEXTAREA') e.preventDefault();
  });

  // foco cruzado y tooltip del pastel
  D.pastelSvg.addEventListener('pointermove', function (e) {
    var g = e.target.closest && e.target.closest('.porcion');
    if (!g) { ocultarTip(); desenfocar(); return; }
    var id = g.getAttribute('data-dolor');
    enfocar('dolor', id);
    var nodos = tipDolor(id);
    if (nodos) mostrarTip(nodos, e.clientX, e.clientY);
  });
  D.pastelSvg.addEventListener('pointerleave', function () { ocultarTip(); desenfocar(); });
  function marcarDesde(e) {
    var x = e.target.closest && e.target.closest('#leyenda > li[data-dolor], #aportes > li[data-ini], #vigilar > li');
    if (!x) return false;
    if (x.dataset.dolor) enfocar('dolor', x.dataset.dolor);
    else if (x.dataset.ini) enfocar('ini', x.dataset.ini);
    else return false;
    return true;
  }
  D.escenario.addEventListener('pointerover', function (e) { if (!e.target.closest('#pastel-svg') && !marcarDesde(e)) desenfocar(); });
  D.escenario.addEventListener('pointerleave', desenfocar);
  D.escenario.addEventListener('focusin', function (e) { if (!marcarDesde(e)) desenfocar(); });

  // resumen fijo en la barra cuando el resultado sale de la vista
  function revisarResumen() {
    D.cuerpo.classList.toggle('ver-resumen', D.cartel.getBoundingClientRect().bottom < 60);
  }
  window.addEventListener('scroll', revisarResumen, { passive: true });

  // ---------------------------------------------------------------- inicio
  var primeraVez = true;
  try { primeraVez = !window.localStorage.getItem(CLAVE_INTRO); } catch (e) { primeraVez = true; }
  D.intro.hidden = !primeraVez;
  pintarTodo(true);
  revisarResumen();
})();
