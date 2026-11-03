/*
 * Suficiencia 2027 · tablero.
 * Toda la matemática vive en engine.js (Motor); este archivo pinta, edita, anima los pasteles
 * y arma el contenido del Excel que se exporta (el archivo .xlsx lo escribe y lo lee excel.js).
 * Los textos que escribe el usuario se insertan siempre con textContent, nunca como HTML.
 */
(function () {
  'use strict';

  var Motor = window.Motor;
  var Excel = window.SuficienciaExcel;
  var crearEjemplos = window.SuficienciaEjemplos;
  // Clave versionada: si el formato cambia, se sube la versión y Motor.normalizarEspacio migra lo anterior
  var CLAVE_DATOS = 'suficiencia.espacio.v3';
  var CLAVE_ANTERIOR = 'suficiencia.datos.v1';
  var CLAVE_INTRO = 'suficiencia.intro.v1';
  var CLAVE_RESPALDO = 'suficiencia.respaldo.v1';   // fecha del último respaldo descargado (para el recordatorio semanal)
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
    casos: 'Total de casos del análisis. Ej.: clientes que se dieron de baja en el periodo. Sirve para calcular el % dolor. Con varios segmentos, cada uno lleva los suyos y de ahí sale su parte del problema.',
    prevalencia: '% dolor: casos con este dolor ÷ casos analizados. La suma puede pasar de 100 % porque un caso puede tener varios dolores.',
    atribucion: 'Impacto del dolor: qué parte del problema total explica este dolor (su porción del pastel). Se calcula solo con los casos; escribe un % solo si quieres cambiarlo.',
    resuelto: 'Qué parte del dolor se comen tus iniciativas con su ejecución esperada. La rayita marca lo que se comerían si todo se cumple.',
    efectividad: 'Qué % del dolor resuelve la iniciativa si se hace completa. Vacío = 0 %.',
    ejecucion: 'Qué % de la iniciativa estará listo de verdad en 2027. 100 % = a tiempo y completa.',
    aporte: 'Puntos que se pierden si la iniciativa falla, y cómo quedaría la suficiencia sin ella.',
    // «¿Qué significa este número?» junto a cada resultado
    suficiencia: 'El número del centro es la suficiencia: qué parte de la brecha cubren tus iniciativas. 100 % o más = alcanza. Ej.: 85 % = te falta cubrir el 15 % de la brecha.',
    mueve: 'Cuántos puntos se mueve el indicador con tus iniciativas y su ejecución esperada. Ej.: churn de 2,80 % que mueve 0,68 pp → queda en 2,12 %.',
    quedarias: 'El valor del indicador en 2027 si todo sale como esperas: hoy más (o menos) lo que mueve. Compáralo con la meta.',
    faltaria: 'Puntos de la brecha que tus iniciativas no alcanzan a cubrir. «Nada» = llegas a la meta.',
    todoCumple: 'La suficiencia si todas las iniciativas se ejecutan al 100 %. Si aquí alcanzas y con la ejecución esperada no, lo que falla son los retrasos.',
    aporteSegmento: 'Cada segmento mueve lo que resuelven las iniciativas en SUS dolores; la suma es lo que mueve el indicador. «De la brecha» es su parte de la suficiencia. Ej.: brecha de 1 pp, Hogar mueve 0,5 pp y Pospago 0,3 pp → 50 % + 30 % = 80 % de suficiencia.',
    parteSegmento: 'Qué parte del problema total es de cada segmento. Vacío = sale de los casos analizados de cada segmento (paso 2). Ej.: 7.000 bajas en Hogar y 3.000 en Pospago → 70 % y 30 %. Escríbela solo si la conoces de otra fuente.',
    aplicaA: 'A qué segmentos llega la iniciativa. Solo resuelve dolores de esos segmentos; en los demás la celda dice «No aplica».'
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
  // Pesos: con los decimales que escribió el usuario (33,33 %), sin redondear a entero
  function pctLibre(v) { return nf(0, 2, true).format(v * 100) + ' %'; }
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

  // Texto plano de una lista de textos y nodos (para el Excel)
  function textoPlano(partes) {
    return [].concat(partes).map(function (p) { return p instanceof Node ? p.textContent : (p === null || p === undefined ? '' : String(p)); }).join('');
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
  function textoSegmentos(ind) { return ind.segmentos.map(function (s) { return s.nombre; }).join(', '); }
  function nombreGrupo(nombre) { return nombre || 'Sin segmento'; }
  // Los grupos del paso 2: un bloque por segmento marcado y, si hay dolores sin segmento (o ningún segmento), uno más
  function gruposTaller(ind) {
    var lista = ind.segmentos.map(function (s) { return s.nombre; });
    if (!lista.length || ind.dolores.some(function (d) { return !Motor.segmentoDe(ind, d); })) lista.push('');
    return lista;
  }
  function doloresDe(ind, nombre) { return ind.dolores.filter(function (d) { return Motor.segmentoDe(ind, d) === nombre; }); }
  function casosDeGrupo(ind, nombre) { var s = nombre ? buscarSegmento(ind, nombre) : null; return s ? s.casos : ind.totalCasos; }
  function filaSeg(r, nombre) { return r.segmentos.filas.filter(function (f) { return f.nombre === nombre; })[0] || null; }
  function unirLista(lista) { return lista.length < 2 ? lista.join('') : lista.slice(0, -1).join(', ') + ' y ' + lista[lista.length - 1]; }

  // ---------------------------------------------------------------- estado
  // estado = el tablero que se ve. calculoId: el cálculo de «Mis cálculos» del que salió (o null);
  // base: cómo estaba al abrirlo o guardarlo, para saber si tiene cambios sin guardar.
  var estado = { sel: null, params: null, indicadores: [], calculoId: null, base: null };
  var calculos = [];
  var sinAlmacen = false;
  var falloGuardar = false;
  aplicarEspacio(leerGuardado());
  var abiertos = {};
  var refs = null;
  var deshacer = null;
  var tAviso = 0;
  var tGuardar = 0;

  // Se empieza con un solo ejemplo lleno para guiar; cada cálculo nuevo empieza desde cero.
  function datosIniciales() {
    var inds = crearEjemplos().slice(0, 1).map(Motor.normalizarIndicador);
    return { sel: inds[0].id, params: Motor.copiarParams(Motor.PARAMS_DEFECTO), indicadores: inds };
  }
  // Los datos en una forma fija (mismo orden de campos) para comparar si algo cambió
  function textoDatos(d) { return JSON.stringify(Motor.normalizarDatos({ params: d.params, indicadores: d.indicadores })); }
  function leerGuardado() {
    var crudo = null;
    var e = null;
    try {
      crudo = window.localStorage.getItem(CLAVE_DATOS);
      if (crudo === null) crudo = window.localStorage.getItem(CLAVE_ANTERIOR);
    } catch (err) { sinAlmacen = true; }
    try { e = crudo ? Motor.normalizarEspacio(JSON.parse(crudo)) : null; } catch (err) { e = null; }
    if (!e) e = { activo: null, actual: null, base: null, calculos: [] };
    if (!e.actual || !e.actual.indicadores.length) {
      e.actual = datosIniciales();
      e.activo = null;
      e.base = textoDatos(e.actual);
    }
    return e;
  }
  function aplicarEspacio(e) {
    calculos = e.calculos;
    estado.sel = e.actual.sel;
    estado.params = e.actual.params;
    estado.indicadores = e.actual.indicadores;
    estado.calculoId = e.activo;
    estado.base = e.base;
    if (!estado.indicadores.some(function (x) { return x.id === estado.sel; })) estado.sel = estado.indicadores[0].id;
  }
  function espacio() {
    return {
      version: Motor.VERSION_ESPACIO, activo: estado.calculoId, base: estado.base, calculos: calculos,
      actual: { sel: estado.sel, params: estado.params, indicadores: estado.indicadores }
    };
  }
  // Se guarda todo a cada cambio. Si el navegador no deja (lleno o bloqueado), se avisa y se sigue:
  // el guardado automático avisa una vez; un guardado que pidió el usuario (explicito) avisa siempre.
  // Devuelve true si quedó guardado.
  function guardar(explicito) {
    clearTimeout(tGuardar);
    tGuardar = 0;
    try {
      window.localStorage.setItem(CLAVE_DATOS, JSON.stringify(espacio()));
      falloGuardar = false;
      return true;
    } catch (e) {
      if (!falloGuardar || explicito) avisar('No se pudo guardar en este navegador (está lleno o bloqueado). Puedes seguir trabajando: usa «Descargar respaldo» en Mis cálculos para no perder nada.', null, 10000);
      falloGuardar = true;
      return false;
    }
  }
  function guardarLuego() { clearTimeout(tGuardar); tGuardar = setTimeout(guardar, 250); }
  function foto() { return JSON.stringify(espacio()); }
  function volverA(f) {
    aplicarEspacio(JSON.parse(f));
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
    panelSeg: $('#panel-segmentos'),
    segAyuda: $('#seg-ayuda'),
    tablaSeg: $('#tabla-seg'),
    calculos: $('#menu-calculos'),
    dlgPregunta: $('#dlg-pregunta'),
    pregTitulo: $('#preg-titulo'),
    pregTexto: $('#preg-texto'),
    pregCampo: $('#preg-campo'),
    pregNombre: $('#preg-nombre'),
    pregBotones: $('#preg-botones'),
    pasos: $('#pasos'),
    taller: $('#taller-cuerpo'),
    ajustes: $('#ajustes-cuerpo'),
    dlg: $('#dlg-como'),
    dlgExportar: $('#dlg-exportar'),
    formExportar: $('#form-exportar'),
    expLista: $('#exp-lista'),
    expOk: $('#exp-ok'),
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

  // ---------------------------------------------------------------- animaciones
  // Cada animación llama a paso(t) con t de 0 a 1. Respeta "reducir movimiento" y termina sola al imprimir.
  var animando = [];
  function animar(duracion, demora, paso, fin) {
    var tw = { vivo: true, inicio: null };
    // Seguro: si el navegador no pinta cuadros (pestaña en segundo plano), la animación igual termina a tiempo
    var seguro = setTimeout(function () { tw.terminar(); }, demora + duracion + 250);
    function quitar() { var i = animando.indexOf(tw); if (i >= 0) animando.splice(i, 1); }
    tw.cancelar = function () { tw.vivo = false; clearTimeout(seguro); quitar(); };
    tw.terminar = function () { if (!tw.vivo) return; tw.cancelar(); paso(1); if (fin) fin(); };
    animando.push(tw);
    paso(0);
    window.requestAnimationFrame(function cuadro(ahora) {
      if (!tw.vivo) return;
      if (tw.inicio === null) tw.inicio = ahora + demora;
      var t = (ahora - tw.inicio) / duracion;
      if (t >= 1) { tw.terminar(); return; }
      if (t > 0) paso(t);
      window.requestAnimationFrame(cuadro);
    });
    return tw;
  }
  function terminarAnimaciones() { animando.slice().forEach(function (tw) { tw.terminar(); }); }
  function suave(t) { return 1 - Math.pow(1 - t, 3); }
  function tramo(t, desde, hasta) { return limitar((t - desde) / (hasta - desde), 0, 1); }
  function entre(a, b, t) { return a + (b - a) * t; }

  // Lo que se dibuja: ángulos y partes resueltas de cada porción, y la suficiencia del anillo.
  function modeloPastel(ind, r, conAnillo) {
    return {
      suf: conAnillo ? r.real.suficiencia : 0,
      tono: (SEMAFORO[r.real.semaforo] || SEMAFORO['SIN DATOS']).tono,
      porciones: porciones(ind, r).map(function (x) {
        var dr = r.dolores[x.j];
        return { id: x.d.id, num: x.j + 1, a0: x.a0, a1: x.a1, real: dr.resueltoReal, plan: dr.resueltoPlan };
      })
    };
  }
  function mismoModelo(a, b) {
    if (Math.abs(a.suf - b.suf) > 1e-9 || a.tono !== b.tono || a.porciones.length !== b.porciones.length) return false;
    return a.porciones.every(function (p, k) {
      var q = b.porciones[k];
      return p.id === q.id && p.num === q.num && Math.abs(p.a0 - q.a0) < 1e-9 && Math.abs(p.a1 - q.a1) < 1e-9 && Math.abs(p.real - q.real) < 1e-9 && Math.abs(p.plan - q.plan) < 1e-9;
    });
  }
  // Empareja las porciones de antes y de ahora. La que aparece nace con ancho cero junto a su vecina
  // y la que se va se encoge hasta cero: así el pastel siempre se ve cerrado mientras se mueve.
  function emparejar(antes, ahora) {
    var enAhora = {};
    var enAntes = {};
    ahora.forEach(function (p) { enAhora[p.id] = p; });
    antes.forEach(function (p) { enAntes[p.id] = p; });
    var pares = [];
    var i = 0;
    ahora.forEach(function (p) {
      while (i < antes.length && !enAhora[antes[i].id]) pares.push({ de: antes[i++], a: null });
      if (i < antes.length && antes[i].id === p.id) i++;
      pares.push({ de: enAntes[p.id] || null, a: p });
    });
    for (; i < antes.length; i++) if (!enAhora[antes[i].id]) pares.push({ de: antes[i], a: null });
    var finDe = 0;
    var finA = 0;
    pares.forEach(function (x) {
      if (!x.de) x.de = { id: x.a.id, num: x.a.num, a0: finDe, a1: finDe, real: 0, plan: 0 };
      if (!x.a) x.a = { id: x.de.id, num: x.de.num, a0: finA, a1: finA, real: x.de.real, plan: x.de.plan };
      finDe = x.de.a1;
      finA = x.a.a1;
    });
    return pares;
  }
  // Al entrar, primero se despliegan las porciones y luego se llenan; al editar, todo se mueve a la vez.
  function mezclar(pares, sufDe, sufA, t, entrada) {
    var tA = suave(entrada ? tramo(t, 0, 0.7) : t);
    var tR = suave(entrada ? tramo(t, 0.25, 1) : t);
    var tS = suave(entrada ? tramo(t, 0.1, 1) : t);
    return {
      suf: entre(sufDe, sufA, tS),
      porciones: pares.map(function (x) {
        return { id: x.a.id, num: x.a.num, a0: entre(x.de.a0, x.a.a0, tA), a1: entre(x.de.a1, x.a.a1, tA), real: entre(x.de.real, x.a.real, tR), plan: entre(x.de.plan, x.a.plan, tR) };
      })
    };
  }

  // Esqueleto del SVG: se crea una vez y en cada cuadro solo cambian los atributos.
  function esqueleto(svg, op) {
    var clave = (op.anillo ? 'a' : '') + (op.plan ? 'p' : '') + (op.etiquetas ? 'e' : '');
    if (svg._dib && svg._dib.clave === clave) return svg._dib;
    svg.replaceChildren();
    var d = { clave: clave, porciones: {} };
    if (op.anillo) {
      d.anillo = s('g', { class: 'anillo', transform: 'rotate(-90 180 180)' },
        s('circle', { class: 'anillo-pista', cx: 180, cy: 180, r: GEO.anillo }));
      d.v1 = s('circle', { class: 'anillo-vuelta', cx: 180, cy: 180, r: GEO.anillo });
      d.v2 = s('circle', { class: 'anillo-vuelta anillo-vuelta--2', cx: 180, cy: 180, r: GEO.anillo });
      d.anillo.appendChild(d.v1);
      d.anillo.appendChild(d.v2);
      svg.appendChild(d.anillo);
    }
    d.vacia = s('path', { class: 'porcion-base vacia', d: sector(GEO.r0, GEO.R, 0, 2 * Math.PI - 1e-4) });
    d.gp = s('g', { class: 'porciones' }, d.vacia);
    svg.appendChild(d.gp);
    svg._dib = d;
    svg._modelo = null;
    return d;
  }
  function porcionDe(d, p) {
    var g = d.porciones[p.id];
    if (g) return g;
    g = s('g', { class: 'porcion', 'data-dolor': p.id });
    g._base = g.appendChild(s('path', { class: 'porcion-base' }));
    g._res = g.appendChild(s('path', { class: 'porcion-resuelta' }));
    g._plan = g.appendChild(s('path', { class: 'porcion-plan' }));
    g._num = g.appendChild(s('text', { class: 'porcion-num' }));
    d.gp.appendChild(g);
    d.porciones[p.id] = g;
    return g;
  }
  function ver(el, si) { el.style.display = si ? '' : 'none'; }
  function pintarModelo(d, m, op) {
    if (d.anillo) {
      d.anillo.setAttribute('class', 'anillo tono-' + m.tono);
      var v1 = limitar(m.suf, 0, 1);
      var v2 = limitar(m.suf - 1, 0, 1);
      d.v1.setAttribute('stroke-dasharray', (v1 * CIRC).toFixed(2) + ' ' + CIRC.toFixed(2));
      d.v2.setAttribute('stroke-dasharray', (v2 * CIRC).toFixed(2) + ' ' + CIRC.toFixed(2));
      ver(d.v1, v1 > 0.002);
      ver(d.v2, v2 > 0.002);
    }
    var vivas = {};
    var hay = false;
    m.porciones.forEach(function (p) {
      var g = porcionDe(d, p);
      var ancho = p.a1 - p.a0;
      vivas[p.id] = true;
      ver(g, ancho > 1e-4);
      if (ancho <= 1e-4) return;
      hay = true;
      var medio = (p.a0 + p.a1) / 2;
      g._base.setAttribute('d', sector(GEO.r0, GEO.R, p.a0, p.a1));
      ver(g._res, p.real > 0.0005);
      if (p.real > 0.0005) g._res.setAttribute('d', sector(GEO.r0, radioRelleno(p.real), p.a0, p.a1));
      var conPlan = !!op.plan && p.plan - p.real > 0.004 && ancho > 0.09;
      ver(g._plan, conPlan);
      if (conPlan) g._plan.setAttribute('d', arco(radioRelleno(p.plan), p.a0 + 0.03, p.a1 - 0.03));
      var conNum = !!op.etiquetas && ancho > 0.2;
      ver(g._num, conNum);
      if (conNum) {
        g._num.textContent = String(p.num);
        g._num.setAttribute('x', (180 + 156 * Math.sin(medio)).toFixed(1));
        g._num.setAttribute('y', (180 - 156 * Math.cos(medio)).toFixed(1));
      }
      // Hacia dónde se asoma la porción cuando la señalas
      g.style.setProperty('--dx', (7 * Math.sin(medio)).toFixed(2) + 'px');
      g.style.setProperty('--dy', (-7 * Math.cos(medio)).toFixed(2) + 'px');
    });
    Object.keys(d.porciones).forEach(function (id) {
      if (!vivas[id]) { d.porciones[id].remove(); delete d.porciones[id]; }
    });
    ver(d.vacia, !hay);
  }

  // op: anillo, plan, etiquetas · animar: 'entrada' (se despliega desde cero), 'cambio' (pasa del dibujo
  // anterior al nuevo) o nada · demora y duracion en ms · alPaso(suficiencia) para contar el número del centro.
  function dibujarPastel(svg, ind, r, op) {
    var d = esqueleto(svg, op);
    var destino = modeloPastel(ind, r, op.anillo);
    if (svg._tw) svg._tw.cancelar();
    svg._tw = null;
    var entrada = op.animar === 'entrada';
    var antes = entrada || !svg._modelo ? { suf: 0, tono: destino.tono, porciones: [] } : svg._modelo;
    var quieto = !op.animar || movimientoReducido.matches || (!entrada && svg._modelo && mismoModelo(antes, destino));
    if (quieto) {
      svg._modelo = destino;
      pintarModelo(d, destino, op);
      if (op.alPaso) op.alPaso(destino.suf);
      if (op.alTerminar) op.alTerminar(false);
      return;
    }
    var pares = emparejar(antes.porciones, destino.porciones);
    svg._tw = animar(op.duracion || (entrada ? 1100 : 480), op.demora || 0, function (t) {
      var m = t >= 1 ? destino : mezclar(pares, antes.suf, destino.suf, t, entrada);
      m.tono = destino.tono;
      svg._modelo = m;
      pintarModelo(d, m, op);
      if (op.alPaso) op.alPaso(m.suf);
    }, function () {
      svg._tw = null;
      if (op.alTerminar) op.alTerminar(entrada);
    });
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
  // arte: el pastel de la tarjeta anterior, para que siga su animación en vez de empezar de nuevo.
  // orden: posición entre las tarjetas que acaban de aparecer (escalona su entrada).
  function tarjeta(ind, r, arte, orden) {
    var sem = SEMAFORO[r.real.semaforo];
    var nueva = !arte;
    if (nueva) arte = s('svg', { class: 'tarjeta-arte', viewBox: '0 0 360 360', 'aria-hidden': 'true' });
    dibujarPastel(arte, ind, r, nueva ? { anillo: true, animar: 'entrada', demora: 160 + (orden || 0) * 90, duracion: 950 } : { anillo: true, animar: 'cambio' });
    var valor = r.real.semaforo === 'SIN DATOS' ? '—' : (r.real.semaforo === 'META CUMPLIDA' ? '✓' : pct(r.real.suficiencia));
    return h('button', {
      class: 'tarjeta' + (nueva ? ' entra' : ''), type: 'button', 'aria-pressed': String(ind.id === estado.sel),
      style: nueva ? '--i:' + (orden || 0) : null,
      'aria-label': nombreDe(ind) + (ind.ejemplo ? ' (ejemplo)' : '') + '. ' + sem.txt + (r.brecha > 0 ? ', suficiencia ' + pct(r.real.suficiencia) : ''),
      data: { accion: 'seleccionar', id: ind.id }
    },
    arte,
    h('span', { class: 'tarjeta-sentido' }, ind.ejemplo ? h('span', { class: 'etq-ejemplo' }, 'Ejemplo') : null,
      ind.sentido ? icono(ind.sentido === 'sube' ? 'sube' : 'baja') : null, textoSentido(ind) + (ind.unidad ? ' · ' + ind.unidad : '')),
    h('span', { class: 'tarjeta-nombre' }, nombreDe(ind)),
    h('span', { class: 'tarjeta-pie' }, anilloMini(r.real.suficiencia, sem.tono), h('b', null, valor), sem.txt));
  }
  function pintarRiel() {
    var foco = document.activeElement && document.activeElement.classList && document.activeElement.classList.contains('tarjeta');
    var primera = !D.riel.firstChild;
    var artes = {};
    $$('.tarjeta[data-id]', D.riel).forEach(function (t) { artes[t.dataset.id] = $('.tarjeta-arte', t); });
    D.riel.replaceChildren();
    var nuevas = 0;
    // Cada tarjeta va con su papelera al lado (un botón no puede ir dentro de otro)
    estado.indicadores.forEach(function (ind) {
      var arte = artes[ind.id] || null;
      D.riel.appendChild(h('div', { class: 'tarjeta-celda' },
        tarjeta(ind, Motor.calcular(ind, estado.params), arte, arte ? 0 : nuevas++),
        h('button', { class: 'tarjeta-borrar', type: 'button', title: 'Eliminar indicador', 'aria-label': 'Eliminar el indicador «' + nombreDe(ind) + '»', data: { accion: 'eliminar', id: ind.id } }, icono('borrar'))));
    });
    D.riel.appendChild(h('button', { class: 'tarjeta tarjeta--nueva' + (primera ? ' entra' : ''), type: 'button', style: primera ? '--i:' + nuevas : null, data: { accion: 'agregar-indicador' } },
      h('b', null, '+ Agregar indicador'), h('span', null, 'Suma otro indicador a este tablero')));
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
    var nueva = tarjeta(ind, r, $('.tarjeta-arte', vieja), 0);
    var conFoco = document.activeElement === vieja;
    var borrar = vieja.nextElementSibling;
    if (borrar) borrar.setAttribute('aria-label', 'Eliminar el indicador «' + nombreDe(ind) + '»');
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
    if (ind.ejemplo) meta.unshift(h('span', { class: 'pildora pildora--ejemplo' }, icono('info'), 'Ejemplo con datos de muestra'));
    if (!vacio(ind.ejecutado)) meta.push(h('span', { class: 'pildora' }, 'Hoy ', h('b', null, valorConUnidad(ind.ejecutado, ind))));
    if (!vacio(ind.meta)) meta.push(h('span', { class: 'pildora' }, 'Meta 2027 ', h('b', null, valorConUnidad(ind.meta, ind))));
    if (r.brecha > 0) meta.push(h('span', { class: 'pildora' }, 'Brecha ', h('b', null, puntos(r.brecha, ind))));
    if (ind.segmentos.length) meta.push(h('span', { class: 'pildora' }, ind.tipoSegmento + ' ', h('b', null, textoSegmentos(ind))));
    fijar(D.metaLinea, meta);

    fijar(D.frase, frase(ind, r));
    pintarKpis(ind, r);
    var numero = pintarCentro(r);
    dibujarPastel(D.pastelSvg, ind, r, {
      anillo: true, plan: true, etiquetas: true, animar: entrando ? 'entrada' : 'cambio', demora: entrando ? 120 : 0,
      alPaso: numero ? function (suf) { numero.textContent = nf(0, 0, true).format(suf * 100); } : null,
      alTerminar: function (fueEntrada) { if (fueEntrada && r.real.semaforo === 'VERDE') celebrar(); }
    });
    D.pastelAlt.textContent = textoAlternativo(ind, r);
    pintarLeyenda(ind, r, entrando);
    pintarVigilar(ind, r);
    pintarAportes(ind, r, entrando);
    pintarSegmentosResumen(ind, r);
  }

  // Tabla del resultado por segmento (con 2 o más segmentos): segmento | parte del problema | mueve | de la brecha, y total
  function pintarSegmentosResumen(ind, r) {
    var rs = r.segmentos;
    D.panelSeg.hidden = !rs.varios;
    if (D.panelSeg.hidden) return;
    var con = r.brecha > 0;
    D.segAyuda.textContent = rs.chequeo === 'PASA DE 100' ? 'Las partes del problema suman ' + pctLibre(rs.sumaPesos) + ': deben sumar 100 %. Corrígelas en el paso 1.'
      : 'Cada segmento mueve lo que resuelven las iniciativas en sus propios dolores. La suma es lo que mueve el indicador (' + puntos(rs.total, ind) + ').';
    var filas = rs.filas.map(function (f) {
      return h('tr', { class: f.nombre ? '' : 'seg-sin' }, h('th', { scope: 'row' }, nombreGrupo(f.nombre)),
        h('td', null, pct(f.parte, 1), h('small', null, f.peso === null ? ' auto' : ' manual')),
        h('td', null, puntos(f.pp, ind)), h('td', null, con ? pct(f.suficiencia) : '—'));
    });
    filas.push(h('tr', { class: 'seg-total' }, h('th', { scope: 'row' }, 'Total'), h('td', null, pct(rs.sumaPartes, 1)), h('td', null, puntos(rs.total, ind)),
      h('td', null, con ? pct(r.real.suficiencia) : '—')));
    fijar(D.tablaSeg, h('thead', null, h('tr', null, h('th', { scope: 'col' }, 'Segmento'), h('th', { scope: 'col' }, 'Parte del problema'),
      h('th', { scope: 'col' }, 'Mueve'), h('th', { scope: 'col' }, 'De la brecha'))), h('tbody', null, filas));
  }
  // Un destello en el anillo cuando la suficiencia alcanza la meta
  function celebrar() {
    D.pastelSvg.classList.remove('celebra');
    void D.pastelSvg.getBoundingClientRect();
    D.pastelSvg.classList.add('celebra');
    setTimeout(function () { D.pastelSvg.classList.remove('celebra'); }, 1600);
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

  // Cuántos puntos mueve el indicador la suficiencia calculada (con la ejecución esperada)
  function textoMueve(ind, r) {
    if (r.real.semaforo === 'SIN DATOS') return 'Falta el paso 1';
    if (r.real.semaforo === 'META CUMPLIDA') return r.real.pp > 0 ? 'La meta ya se cumple: esto es extra' : 'Aún sin iniciativas';
    if (!r.ranking.length) return 'Agrega iniciativas en el paso 3';
    return 'el ' + pct(r.real.suficiencia) + ' de la brecha de ' + puntos(r.brecha, ind);
  }
  var AYUDA_KPI = { 'Mueve el indicador': 'mueve', 'Así quedarías en 2027': 'quedarias', 'Te faltaría': 'faltaria', 'Si todo se cumple': 'todoCumple' };
  function pintarKpis(ind, r) {
    var tono = (SEMAFORO[r.real.semaforo] || SEMAFORO['SIN DATOS']).tono;
    var mueve = r.real.semaforo === 'SIN DATOS' ? '—'
      : h('span', { class: 'kpi-mov' }, ind.sentido && r.real.pp > 0 ? icono(ind.sentido === 'sube' ? 'sube' : 'baja') : null, puntos(r.real.pp, ind));
    var items;
    if (r.real.semaforo === 'SIN DATOS') {
      items = [['Mueve el indicador', mueve, textoMueve(ind, r)], ['Así quedarías en 2027', '—', ''], ['Te faltaría', '—', ''], ['Si todo se cumple', '—', '']];
    } else if (r.real.semaforo === 'META CUMPLIDA') {
      items = [['Mueve el indicador', mueve, textoMueve(ind, r)], ['Hoy', valorConUnidad(ind.ejecutado, ind), ''], ['Meta 2027', valorConUnidad(ind.meta, ind), ''], ['Te faltaría', 'Nada', 'La meta ya se cumple']];
    } else {
      items = [
        ['Mueve el indicador', mueve, textoMueve(ind, r)],
        ['Así quedarías en 2027', valorConUnidad(r.real.proyectado, ind), 'Meta ' + valorConUnidad(ind.meta, ind)],
        ['Te faltaría', r.real.brechaRestante > 0 ? puntos(r.real.brechaRestante, ind) : 'Nada',
          r.real.brechaRestante > 0 ? 'de ' + puntos(r.brecha, ind) + ' de brecha' : 'Cierras toda la brecha'],
        ['Si todo se cumple', pct(r.plan.suficiencia), 'de la brecha, con 100 % de ejecución']
      ];
    }
    fijar(D.kpis, items.map(function (it, k) {
      var ayuda = AYUDA_KPI[it[0]];
      return h('div', { class: 'kpi' + (k === 0 ? ' kpi--mueve tono-' + tono : '') }, h('dt', null, it[0], ayuda ? botonAyuda(ayuda, it[0]) : null), h('dd', null, it[1], it[2] ? h('span', { class: 'sub' }, it[2]) : null));
    }));
  }

  // Devuelve el elemento del número para que la animación del anillo lo vaya contando.
  function pintarCentro(r) {
    var sem = SEMAFORO[r.real.semaforo];
    var valor;
    var etq;
    var numero = null;
    if (r.real.semaforo === 'SIN DATOS') { valor = '—'; etq = 'Completa el paso 1'; }
    else if (r.real.semaforo === 'META CUMPLIDA') { valor = '✓'; etq = 'No hay brecha que cerrar'; }
    else {
      numero = h('span', null, nf(0, 0, true).format(r.real.suficiencia * 100));
      valor = [numero, h('small', null, '%')];
      etq = 'de la brecha cubierta';
    }
    fijar(D.pastelCentro, h('span', { class: 'pastel-valor' }, valor), h('span', { class: 'pastel-etq' }, etq), chip(sem));
    return numero;
  }

  function textoAlternativo(ind, r) {
    if (r.real.semaforo === 'SIN DATOS') return 'Pastel del problema: todavía sin datos.';
    var partes = ind.dolores.map(function (d, j) {
      var dr = r.dolores[j];
      return (d.nombre || 'Dolor ' + (j + 1)) + ': impacto del dolor ' + pct(dr.atribucion) + ', ' + pct(dr.resueltoReal) + ' resuelto';
    });
    return 'Pastel del problema con ' + ind.dolores.length + ' dolores. ' + partes.join('; ') +
      '. Suficiencia con la ejecución esperada: ' + pct(r.real.suficiencia) + ' (' + SEMAFORO[r.real.semaforo].txt + ').';
  }

  // Con muchos dolores la leyenda muestra los primeros y un botón para ver todos.
  var MAX_LEYENDA = 6;
  var verTodaLeyenda = false;
  function pintarLeyenda(ind, r, entrando) {
    var items = [];
    ind.dolores.forEach(function (d, j) {
      var dr = r.dolores[j];
      if (dr.atribucion <= 0 && !d.nombre) return;
      var conPlan = dr.resueltoPlan - dr.resueltoReal > 0.004;
      items.push(h('li', { data: { dolor: d.id }, style: '--k:' + items.length },
        h('button', { class: 'ley-btn', type: 'button', data: { accion: 'ir', tipo: 'dolor', id: d.id }, 'aria-label': 'Ir al dolor ' + (j + 1) + ' en el taller' },
          h('span', { class: 'num-d', 'aria-hidden': 'true' }, String(j + 1)),
          h('span', { class: 'ley-nombre' }, d.nombre || 'Dolor sin nombre'),
          h('span', { class: 'ley-dato' }, pct(dr.resueltoReal)),
          h('span', { class: 'ley-barra', style: '--v:' + dr.resueltoReal.toFixed(4) + ';--p:' + dr.resueltoPlan.toFixed(4), 'aria-hidden': 'true' }, h('i'), conPlan ? h('b') : null),
          h('span', { class: 'ley-sub' }, r.segmentos.varios ? h('span', { class: 'ley-seg' }, nombreGrupo(Motor.segmentoDe(ind, d))) : null,
            h('span', null, 'Impacto del dolor ' + pct(dr.atribucion)), conPlan ? h('span', null, pct(dr.resueltoPlan) + ' si todo se cumple') : null))));
    });
    var total = items.length;
    if (!total) items.push(h('li', { class: 'vacio-oscuro' }, 'Aún no hay dolores. Agrégalos en el paso 2.'));
    if (total > MAX_LEYENDA) {
      if (!verTodaLeyenda) items = items.slice(0, MAX_LEYENDA);
      items.push(h('li', { class: 'ley-mas' }, h('button', { class: 'btn-texto', type: 'button', data: { accion: 'mas-leyenda' }, 'aria-expanded': String(verTodaLeyenda) },
        verTodaLeyenda ? 'Ver menos' : 'Ver los ' + total + ' dolores')));
    }
    D.leyenda.classList.toggle('entrando', !!entrando);
    fijar(D.leyenda, items);
  }

  function alertas(ind, r) {
    var lista = [];
    var nom = function (t) { return h('strong', null, '«' + t + '»'); };
    if (vacio(ind.ejecutado) || vacio(ind.meta) || !ind.sentido) {
      lista.push({ tono: 'neutro', icono: 'info', texto: ['Completa el ', h('strong', null, 'paso 1'), ': hoy, meta 2027 y si es mejor que suba o que baje.'], ir: { tipo: 'paso', id: '1' } });
      return lista;
    }
    if (!ind.segmentos.length) lista.push({ tono: 'alerta', icono: 'info', texto: ['Marca al menos un ', h('strong', null, 'segmento'), ' en el paso 1 (por ejemplo, Hogar o Pospago).'], ir: { tipo: 'paso', id: '1' } });
    var rs = r.segmentos;
    if (rs.chequeo === 'PASA DE 100') lista.push({ tono: 'critico', icono: 'triangulo', texto: ['Las partes del problema de los segmentos suman ', h('strong', null, pctLibre(rs.sumaPesos)), '. Deben sumar 100 %.'], ir: { tipo: 'paso', id: '1' } });
    if (r.posiblePorcentaje) lista.push({ tono: 'alerta', icono: 'triangulo', texto: ['Revisa hoy y la meta: parece que escribiste el símbolo %. Si tu valor es 2,8 %, escribe solo 2,8.'], ir: { tipo: 'paso', id: '1' } });
    if (r.brecha === 0) return lista;
    if (r.chequeoAtribucion === 'SIN DOLORES' && !rs.varios) lista.push({ tono: 'neutro', icono: 'info', texto: ['Agrega los dolores y sus casos en el ', h('strong', null, 'paso 2'), '.'], ir: { tipo: 'paso', id: '2' } });
    // Con varios segmentos: a cada uno le deben llegar sus casos y sus dolores, si no su parte del problema queda sin explicar
    var faltaSegmento = false;
    if (rs.varios) {
      rs.filas.forEach(function (f) {
        var seg = h('strong', null, '«' + nombreGrupo(f.nombre) + '»');
        if (!f.nombre) {
          lista.push({ tono: 'alerta', icono: 'info', texto: [f.dolores === 1 ? 'Hay 1 dolor ' : 'Hay ' + f.dolores + ' dolores ', 'sin segmento. Elige a qué segmento pertenece cada uno en el paso 2.'], ir: { tipo: 'grupo', id: '' } });
          faltaSegmento = true;
        } else if (!f.dolores) {
          lista.push({ tono: f.parte > 0 ? 'critico' : 'neutro', icono: f.parte > 0 ? 'triangulo' : 'info', ir: { tipo: 'grupo', id: f.nombre },
            texto: f.parte > 0 ? ['Faltan los dolores de ', seg, ': es el ' + pct(f.parte) + ' del problema y sin dolores no se resuelve nada ahí.'] : ['Agrega los casos analizados y los dolores de ', seg, ' en el paso 2.'] });
          faltaSegmento = true;
        } else if (f.peso === null && !(f.casos > 0)) {
          lista.push({ tono: 'critico', icono: 'triangulo', texto: ['Escribe los casos analizados de ', seg, ' en el paso 2: sin ellos su parte del problema es 0 %.'], ir: { tipo: 'grupo', id: f.nombre } });
          faltaSegmento = true;
        }
      });
      if (rs.chequeo === 'NO SUMA 100' && !faltaSegmento) lista.push({ tono: 'critico', icono: 'triangulo', texto: ['Las partes del problema de los segmentos suman ', h('strong', null, pct(rs.sumaPartes, 1)), '; deben sumar 100 %. Deja vacía la parte de algún segmento para que salga de sus casos.'], ir: { tipo: 'paso', id: '1' } });
    }
    if (r.chequeoAtribucion === 'NO SUMA 100' && !faltaSegmento && (rs.chequeo === 'OK' || rs.chequeo === 'SIN SEGMENTO')) lista.push({ tono: 'critico', icono: 'triangulo', texto: ['El impacto de los dolores suma ', h('strong', null, pct(r.sumaAtribucion, 1)), '; debe sumar 100 %.'], ir: { tipo: 'paso', id: '2' } });
    ind.dolores.forEach(function (d, j) {
      var dr = r.dolores[j];
      if (dr.alerta !== 'SIN COBERTURA') return;
      lista.push({ tono: 'critico', icono: 'triangulo', dolor: d.id, texto: [nom(d.nombre), ' explica el ' + pct(dr.atribucion) + ' del problema y solo se resuelve el ' + pct(dr.resueltoReal) + '. Falta una iniciativa para este dolor.'], ir: { tipo: 'dolor', id: d.id } });
    });
    if (!r.ranking.length && r.chequeoAtribucion !== 'SIN DOLORES') lista.push({ tono: 'neutro', icono: 'info', texto: ['Agrega tus iniciativas en el ', h('strong', null, 'paso 3'), ' para saber si te alcanza.'], ir: { tipo: 'paso', id: '3' } });
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
    // En un cálculo que empieza desde cero, una pista para ver el ejemplo lleno
    if (r.real.semaforo === 'SIN DATOS' && !ind.ejemplo) {
      var ej = estado.indicadores.filter(function (x) { return x.ejemplo; })[0];
      lista.push({ tono: 'neutro', icono: 'ojo', boton: 'Ver ejemplo', ir: { tipo: 'ejemplo', id: ej ? ej.id : '' },
        texto: ['¿Quieres ver un cálculo lleno? Mira el ejemplo ', h('strong', null, '«' + (ej ? nombreDe(ej) : crearEjemplos()[0].nombre) + '»'), ' y vuelve cuando quieras.'] });
    }
    if (!lista.length) {
      fijar(D.vigilar, h('li', { class: 'vig vig--ok' }, icono('check'), h('p', null, h('strong', null, 'Todo en orden.'), ' No hay alertas para este indicador.')));
      return;
    }
    var visibles = verTodasAlertas ? lista : lista.slice(0, 4);
    var items = visibles.map(function (a) {
      return h('li', { class: 'vig vig--' + a.tono, data: { dolor: a.dolor, ini: a.ini } },
        icono(a.icono), h('p', null, a.texto),
        h('button', { class: 'btn-chico', type: 'button', data: { accion: 'ir', tipo: a.ir.tipo, id: a.ir.id } }, a.boton || 'Ver'));
    });
    if (lista.length > 4) items.push(h('li', null, h('button', { class: 'btn-texto', type: 'button', data: { accion: 'mas-alertas' } }, verTodasAlertas ? 'Ver menos' : 'Ver ' + (lista.length - 4) + ' más')));
    fijar(D.vigilar, items);
  }

  function pintarAportes(ind, r, entrando) {
    D.aportes.classList.toggle('entrando', !!entrando);
    if (r.real.semaforo === 'SIN DATOS' || r.real.semaforo === 'META CUMPLIDA') {
      fijar(D.aportes, h('li', { class: 'vacio-oscuro' }, r.real.semaforo === 'SIN DATOS' ? 'Completa el paso 1 para ver cuánto aporta cada iniciativa.' : 'La meta ya se cumple: no hay brecha que repartir.'));
      return;
    }
    if (!r.ranking.length) { fijar(D.aportes, h('li', { class: 'vacio-oscuro' }, 'Aún no hay iniciativas. Agrégalas en el paso 3.')); return; }
    var max = Math.max.apply(null, r.ranking.map(function (id) { return buscar(r.iniciativas, id).aporte; }).concat([1e-9]));
    fijar(D.aportes, r.ranking.map(function (id, k) {
      var x = buscar(r.iniciativas, id);
      var ini = buscar(ind.iniciativas, id);
      return h('li', { data: { ini: id }, style: '--k:' + k },
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
      campoSegmento(ind),
      h('div', { class: 'campo c-nombre' }, refs.pista1));
    refs.segmentos = segs;
    refs.bala = h('div', { class: 'bala' });
    hj.cuerpo.appendChild(h('div', { class: 'paso1' }, campos, refs.bala));
    refs.paso1 = hj;
    return hj.sec;
  }

  // Segmento: primero la unidad (UMM, UMC o Transversal) y luego los segmentos. Transversal admite los de las dos.
  // Con 2 o más segmentos, cada uno puede llevar su parte del problema escrita a mano (opcional).
  function campoSegmento(ind) {
    var idT = idUnico();
    refs.tiposSeg = h('div', { class: 'segmentos', role: 'group', 'aria-labelledby': idT },
      Motor.TIPOS_SEGMENTO.map(function (t) { return h('button', { type: 'button', 'aria-pressed': String(ind.tipoSegmento === t), data: { accion: 'tipo-segmento', valor: t } }, t); }));
    refs.segLista = h('div', { class: 'seg-lista' });
    refs.segPista = h('p', { class: 'pista' });
    refs.segAviso = h('p', { class: 'pista' });
    refs.segError = h('p', { class: 'pista pista--alerta', role: 'alert', hidden: true });
    pintarSegmentos(ind);
    return h('div', { class: 'campo c-segmento' }, h('div', { class: 'etq', id: idT }, '¿De qué segmento es el indicador?'),
      refs.tiposSeg, refs.segPista, refs.segLista, refs.segAviso, refs.segError);
  }
  function buscarSegmento(ind, nombre) { return ind.segmentos.filter(function (s) { return s.nombre === nombre; })[0] || null; }
  function pintarSegmentos(ind) {
    var trans = ind.tipoSegmento === 'Transversal';
    var varios = ind.segmentos.length > 1;
    refs.segAportes = {};
    refs.segPesos = {};
    refs.segError.hidden = true;
    if (!ind.tipoSegmento) { refs.segLista.replaceChildren(); return; }
    fijar(refs.segLista, (trans ? ['UMM', 'UMC'] : [ind.tipoSegmento]).map(function (g) {
      return h('div', { class: 'seg-grupo', role: 'group', 'aria-label': 'Segmentos ' + g }, trans ? h('span', { class: 'seg-unidad' }, g) : null,
        Motor.SEGMENTOS[g].map(function (nombre) {
          var s = buscarSegmento(ind, nombre);
          var id = idUnico('seg');
          var fila = h('div', { class: 'seg-fila' + (s ? ' marcado' : '') },
            h('label', { class: 'seg-marca', for: id }, h('input', { type: 'checkbox', id: id, value: nombre, checked: !!s, data: { campo: 'segmento' } }), nombre));
          if (varios && s) {
            var idp = idUnico('peso');
            refs.segAportes[nombre] = h('span', { class: 'seg-aporte' });
            refs.segPesos[nombre] = entrada('pct', s.peso, { id: idp, placeholder: 'Auto', data: { campo: 'peso', segmento: nombre } });
            fila.appendChild(h('div', { class: 'seg-peso' },
              h('div', { class: 'etq' }, h('label', { class: 'seg-peso-etq', for: idp }, 'Parte del problema (%)'), botonAyuda('parteSegmento', 'parte del problema')),
              conUnidad(refs.segPesos[nombre], '%'),
              refs.segAportes[nombre]));
          }
          return fila;
        }));
    }));
    // Solo si el usuario lo pide: la herramienta nunca reparte a mano por su cuenta (lo automático sale de los casos)
    if (varios) {
      refs.segLista.appendChild(h('button', { class: 'btn btn-repartir', type: 'button', data: { accion: 'repartir-pesos' } }, 'Repartir en partes iguales'));
      refs.segLista.appendChild(h('p', { class: 'pista seg-ayuda' }, 'Con varios segmentos, en los pasos 2 y 3 llenas los dolores de cada segmento y a cuáles aplica cada iniciativa. ' +
        'La parte del problema de cada segmento sale de sus casos analizados; escríbela aquí solo si la conoces de otra fuente.'));
    }
  }

  // Un bloque por segmento: sus casos analizados, su parte del problema y sus dolores.
  // Con un solo segmento se ve como siempre (un solo bloque, sin título).
  function paso2(ind) {
    var grupos = gruposTaller(ind);
    var varios = grupos.length > 1;
    var hj = hoja(2, '¿Qué causa el problema?', varios
      ? 'Llena los dolores de cada segmento con sus propios casos. Los dolores salen del análisis de datos; un caso puede tener varios dolores.'
      : 'Los dolores salen del análisis de datos del indicador. Un caso puede tener varios dolores.', 'paso-2');
    refs.grupos = {};
    grupos.forEach(function (nombre) { hj.cuerpo.appendChild(bloqueDolores(ind, nombre, varios)); });
    refs.notaAtrib = h('p', { class: 'nota nota--alerta', hidden: true });
    hj.cuerpo.appendChild(refs.notaAtrib);
    refs.paso2 = hj;
    return hj.sec;
  }

  function bloqueDolores(ind, nombre, varios) {
    var idC = idUnico();
    var lista = doloresDe(ind, nombre);
    var g = { notaPrev: h('p', { class: 'pista' }), parte: h('span', { class: 'seg-bloque-parte' }) };
    refs.grupos[nombre] = g;
    var bloque = h('div', { class: 'seg-bloque' + (varios ? ' seg-bloque--varios' : '') + (nombre ? '' : ' seg-bloque--sin'), data: { grupoBloque: nombre } });
    if (varios) {
      bloque.appendChild(h('div', { class: 'seg-bloque-cab' }, h('h4', { class: 'seg-bloque-titulo' }, nombre ? 'Dolores de ' + nombre : 'Dolores sin segmento'), g.parte));
      if (!nombre) bloque.appendChild(h('p', { class: 'nota nota--alerta' }, icono('info'), h('span', null, 'Estos dolores vienen de antes de separar por segmento. Elige el segmento de cada uno en la lista «Segmento».')));
    }
    var etiquetaCasos = varios ? (nombre ? 'Casos analizados en ' + nombre : 'Casos analizados (sin segmento)') : 'Casos analizados';
    bloque.appendChild(h('div', { class: 'casos-fila' },
      h('div', { class: 'campo' }, etq(etiquetaCasos, 'casos', idC),
        entrada('entero', casosDeGrupo(ind, nombre), { id: idC, placeholder: 'Ej.: 12.000', data: { campo: 'casosGrupo', segmento: nombre } })),
      g.notaPrev));
    if (lista.length) {
      bloque.appendChild(h('div', { class: 'tabla-cab' },
        h('span', null, '#'), h('span', null, 'Dolor'), h('span', null, 'Casos'),
        h('span', null, '% dolor', botonAyuda('prevalencia', '% dolor')),
        h('span', null, 'Impacto del dolor', botonAyuda('atribucion', 'impacto del dolor')),
        h('span', null, 'Resuelto', botonAyuda('resuelto', 'resuelto')),
        h('span', null, 'Fuente del dato'), h('span')));
    }
    lista.forEach(function (d) { bloque.appendChild(filaDolor(ind, d, ind.dolores.indexOf(d), varios)); });
    if (!lista.length) {
      bloque.appendChild(h('p', { class: 'nota' }, icono('info'), h('span', null, nombre && varios
        ? 'Agrega el primer dolor de ' + nombre + ': una causa del problema en este segmento, con sus casos.'
        : 'Agrega el primer dolor: una causa del problema que salga del análisis de datos.')));
    }
    // Los dolores sin segmento solo se reparten; los nuevos siempre van a un segmento (si hay alguno marcado)
    if (nombre || !ind.segmentos.length) {
      var lleno = lista.length >= Motor.MAX_DOLORES;
      var acciones = h('div', { class: 'seg-bloque-acciones' },
        h('button', { class: 'btn-agregar', type: 'button', data: { accion: 'agregar-dolor', segmento: nombre }, disabled: lleno },
          icono('plus'), lleno ? 'Llegaste al máximo de ' + Motor.MAX_DOLORES + ' dolores' + (varios ? ' en ' + nombre : ' por segmento') : (varios ? 'Agregar dolor de ' + nombre : 'Agregar dolor'),
          h('small', null, lista.length + ' de ' + Motor.MAX_DOLORES)));
      // Para no escribir dos veces la misma lista: copia los nombres (y fuentes) de los dolores de otro segmento
      var otros = ind.segmentos.filter(function (s) { return s.nombre !== nombre && doloresDe(ind, s.nombre).length; });
      if (varios && nombre && otros.length && !lleno) {
        var idS = idUnico();
        acciones.appendChild(h('div', { class: 'copiar-dolores' },
          h('label', { for: idS }, 'Copiar los dolores de'),
          h('select', { class: 'entrada', id: idS, data: { campo: 'copiarDolores', segmento: nombre } },
            h('option', { value: '' }, 'Elige un segmento'),
            otros.map(function (s) { return h('option', { value: s.nombre }, s.nombre + ' (' + doloresDe(ind, s.nombre).length + ')'); }))));
      }
      bloque.appendChild(acciones);
    }
    return bloque;
  }

  function filaDolor(ind, d, j, varios) {
    var ids = { n: idUnico(), c: idUnico(), a: idUnico(), f: idUnico() };
    var x = {
      prev: h('span', { class: 'valor-calc' }),
      atribIn: entrada('pct', d.atribucionManual, { id: ids.a, 'aria-label': 'Impacto del dolor ' + (j + 1) + ' (vacío = automático)', data: { campo: 'atribucionManual', dolor: d.id } }),
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
        entradaArea(d.nombre, { id: ids.n, placeholder: 'Ej.: Precio / oferta de la competencia', maxlength: 120, data: { campo: 'nombre', dolor: d.id } }),
        varios ? selectorSegmento(ind, d, j) : null),
      h('div', { class: 'd-datos' },
        h('div', { class: 'd-casos' }, h('label', { class: 'et-movil', for: ids.c }, 'Casos'),
          entrada('entero', d.casos, { id: ids.c, placeholder: '0', 'aria-label': 'Casos con el dolor ' + (j + 1), data: { campo: 'casos', dolor: d.id } })),
        h('div', { class: 'd-prev' }, h('span', { class: 'et-movil' }, '% dolor'), x.prev),
        h('div', { class: 'd-atrib' }, h('span', { class: 'et-movil' }, 'Impacto del dolor'),
          h('div', { class: 'atrib-ctl' }, conUnidad(x.atribIn, '%'), x.atribModo), x.atribBarra),
        h('div', { class: 'd-res' }, h('span', { class: 'et-movil' }, 'Resuelto'),
          h('div', { class: 'res-ctl' }, h('div', { class: 'res-linea' }, x.resValor, x.resChip), x.resBarra))),
      h('div', { class: 'd-fuente' }, h('label', { class: 'et-movil', for: ids.f }, 'Fuente del dato'),
        entradaArea(d.evidencia, { id: ids.f, placeholder: '¿De dónde sale el dato?', maxlength: 300, 'aria-label': 'Fuente del dato del dolor ' + (j + 1), data: { campo: 'evidencia', dolor: d.id } })),
      h('div', { class: 'd-mas' }, h('button', { class: 'btn-icono peligro', type: 'button', data: { accion: 'borrar-dolor', id: d.id }, 'aria-label': 'Borrar el dolor ' + (j + 1) }, icono('borrar'))));
  }

  // Para mover un dolor a otro segmento (o darle uno a los que vienen sin segmento)
  function selectorSegmento(ind, d, j) {
    var id = idUnico();
    var seg = Motor.segmentoDe(ind, d);
    var opciones = ind.segmentos.map(function (s) { return h('option', { value: s.nombre, selected: s.nombre === seg }, s.nombre); });
    if (!seg) opciones.unshift(h('option', { value: '', selected: true }, 'Elige el segmento'));
    return h('div', { class: 'd-seg' + (seg ? '' : ' d-seg--falta') }, h('label', { for: id }, 'Segmento'),
      h('select', { class: 'entrada', id: id, 'aria-label': 'Segmento del dolor ' + (j + 1), data: { campo: 'segmentoDolor', dolor: d.id } }, opciones));
  }

  function paso3(ind) {
    var hj = hoja(3, '¿Qué vas a hacer?', 'En cada dolor escribe qué % resuelve la iniciativa si se hace completa (vacío = 0 %). Luego ajusta cuánto estará ejecutado en 2027.', 'paso-3');
    // Ancho mínimo de la tabla = suma de los mínimos de sus columnas; si no cabe, se desliza de lado.
    var nd = Math.max(1, ind.dolores.length);
    var matriz = h('div', { class: 'matriz', style: '--nd:' + nd + ';--min-ancho:' + (706 + 68 * nd) + 'px' });
    var varios = ind.segmentos.length > 1;
    var cab = h('div', { class: 'matriz-cab' }, h('span', null, 'N°'), h('span', null, 'Iniciativa', botonAyuda('efectividad', 'efectividad')));
    if (ind.dolores.length) {
      ind.dolores.forEach(function (d, j) {
        var nombre = h('span', null, d.nombre || 'Dolor ' + (j + 1));
        var seg = varios ? h('small', { class: 'cab-seg' }, nombreGrupo(Motor.segmentoDe(ind, d))) : null;
        var celda = h('span', { class: 'cab-dolor', title: d.nombre || 'Dolor ' + (j + 1) }, h('span', { class: 'cab-num' }, h('span', { class: 'num-claro', 'aria-hidden': 'true' }, String(j + 1)), seg), nombre);
        refs.cabDolores.push({ id: d.id, j: j, el: nombre, celda: celda });
        cab.appendChild(celda);
      });
    } else {
      cab.appendChild(h('span', null, 'Dolores'));
    }
    cab.appendChild(h('span', null, 'Ejecución a 2027', botonAyuda('ejecucion', 'ejecución')));
    cab.appendChild(h('span', null, 'Aporte esperado', botonAyuda('aporte', 'aporte')));
    cab.appendChild(h('span'));
    matriz.appendChild(cab);
    ind.iniciativas.forEach(function (x, i) { matriz.appendChild(filaIni(ind, x, i)); });
    if (varios) hj.cuerpo.appendChild(h('p', { class: 'nota' }, icono('info'), h('span', null, 'En cada iniciativa marca ', h('b', null, 'a qué segmentos aplica'),
      ': solo resuelve dolores de esos segmentos y en los demás la celda dice «No aplica». Las columnas de dolores van agrupadas por segmento.')));
    if (ind.dolores.length > 6) hj.cuerpo.appendChild(h('p', { class: 'pista pista--matriz' }, 'Con muchos dolores la tabla se desliza hacia la derecha; el nombre de la iniciativa se queda fijo.'));
    hj.cuerpo.appendChild(h('div', { class: 'matriz-scroll' }, matriz));
    if (!ind.iniciativas.length) hj.cuerpo.appendChild(h('p', { class: 'nota' }, icono('info'), h('span', null, 'Agrega la primera iniciativa: algo concreto que vas a hacer para resolver uno o varios dolores.')));
    var max = Motor.maxIniciativas(ind);
    var lleno = ind.iniciativas.length >= max;
    hj.cuerpo.appendChild(h('button', { class: 'btn-agregar', type: 'button', data: { accion: 'agregar-ini' }, disabled: lleno },
      icono('plus'), lleno ? 'Llegaste al máximo de ' + max + ' iniciativas' + (varios ? ' (15 por segmento)' : ' por indicador') : 'Agregar iniciativa',
      h('small', null, ind.iniciativas.length + ' de ' + max)));
    refs.paso3 = hj;
    return hj.sec;
  }

  function filaIni(ind, x, i) {
    var ids = { n: idUnico(), r: idUnico(), f: idUnico(), e: idUnico(), det: idUnico('det') };
    var q = { celdas: {}, nombresCelda: {} };
    var celdas = h('div', { class: 'i-celdas' });
    var varios = ind.segmentos.length > 1;
    if (!ind.dolores.length) celdas.appendChild(h('div', { class: 'sin-dolores' }, 'Primero agrega dolores en el paso 2.'));
    ind.dolores.forEach(function (d, j) {
      if (!Motor.aplica(ind, x, d)) {
        celdas.appendChild(h('div', { class: 'celda' },
          h('span', { class: 'celda-et', 'aria-hidden': 'true' }, h('span', { class: 'num-claro' }, String(j + 1)), h('span', null, d.nombre || 'Dolor ' + (j + 1))),
          h('div', { class: 'celda-caja celda-caja--no', title: 'La iniciativa no aplica a ' + Motor.segmentoDe(ind, d) }, h('span', null, 'No aplica'))));
        return;
      }
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
        entradaArea(x.nombre, { id: ids.n, placeholder: 'Ej.: Resolución en el primer contacto', maxlength: 120, data: { campo: 'nombre', ini: x.id } }),
        varios ? aplicaA(ind, x, i) : null, q.revision),
      celdas,
      h('div', { class: 'i-ejec' }, h('span', { class: 'et-movil' }, 'Ejecución a 2027'), h('div', { class: 'ejec-ctl' }, q.rango, conUnidad(q.ejecIn, '%'))),
      h('div', { class: 'i-aporte' }, h('span', { class: 'et-movil' }, 'Aporte con la ejecución esperada'),
        h('div', { class: 'ap-linea' }, q.apValor, q.apPct), q.apBarra, h('div', { class: 'ap-mini' }, q.apChip, q.apFalla)),
      h('div', { class: 'i-mas' },
        h('button', { class: 'btn-icono', type: 'button', 'aria-expanded': String(abierto), 'aria-controls': ids.det, 'aria-label': 'Responsable, fecha y evidencia de la iniciativa ' + (i + 1), data: { accion: 'detalle-ini', id: x.id } }, icono('abajo', 'girar')),
        h('button', { class: 'btn-icono peligro', type: 'button', 'aria-label': 'Borrar la iniciativa ' + (i + 1), data: { accion: 'borrar-ini', id: x.id } }, icono('borrar'))),
      q.detalle);
  }

  // «Aplica a»: un botón por segmento (todos encendidos = aplica a todos, también a los que se marquen después)
  function aplicaA(ind, x, i) {
    var id = idUnico();
    var todos = !x.segmentos.length;
    return h('div', { class: 'aplica', role: 'group', 'aria-labelledby': id },
      h('span', { class: 'aplica-etq', id: id }, 'Aplica a', botonAyuda('aplicaA', 'aplica a')),
      ind.segmentos.map(function (s) {
        var si = todos || x.segmentos.indexOf(s.nombre) >= 0;
        return h('button', { class: 'aplica-seg', type: 'button', 'aria-pressed': String(si), 'aria-label': s.nombre + ': la iniciativa ' + (i + 1) + (si ? ' aplica' : ' no aplica'),
          data: { accion: 'aplica-ini', id: x.id, segmento: s.nombre } }, s.nombre);
      }));
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
    // segmento
    $$('button', refs.tiposSeg).forEach(function (b) { b.setAttribute('aria-pressed', String(b.dataset.valor === ind.tipoSegmento)); });
    refs.tiposSeg.classList.toggle('sin-eleccion', !ind.segmentos.length);
    refs.segPista.textContent = !ind.tipoSegmento ? 'Elige UMM, UMC o Transversal (Transversal deja marcar segmentos de las dos).'
      : (!ind.segmentos.length ? 'Marca al menos un segmento.' : '');
    refs.segPista.className = 'pista' + (ind.segmentos.length ? '' : ' pista--alerta');
    refs.segPista.hidden = !refs.segPista.textContent;
    var rs = r.segmentos;
    rs.filas.forEach(function (f) {
      var el = refs.segAportes[f.nombre];
      if (el) el.textContent = (f.peso === null ? 'Auto (de sus casos): ' + pct(f.parteAuto, 1) : 'Escrita a mano') + ' · mueve ' + puntos(f.pp, ind);
      var inp = refs.segPesos[f.nombre];
      if (inp) inp.placeholder = rs.chequeo === 'SIN CASOS' ? 'Auto' : editablePct(f.parteAuto);
    });
    refs.segAviso.textContent = rs.chequeo === 'NO SUMA 100' ? 'Las partes del problema suman ' + pct(rs.sumaPartes, 1) + ': deben sumar 100 %. Deja vacía la de algún segmento (sale de sus casos) o completa sus casos en el paso 2.'
      : (rs.chequeo === 'SIN CASOS' ? 'Escribe los casos analizados de cada segmento en el paso 2: de ahí sale la parte del problema de cada uno.' : '');
    refs.segAviso.hidden = !refs.segAviso.textContent;
    if (rs.chequeo === 'PASA DE 100') { refs.segError.textContent = 'Las partes del problema suman ' + pctLibre(rs.sumaPesos) + '. Deben sumar máximo 100 %: baja alguna.'; refs.segError.hidden = false; }
    pintarBala(ind, r);

    // paso 2: por cada segmento, su parte del problema y la suma de sus % dolor
    Object.keys(refs.grupos).forEach(function (nombre) {
      var g = refs.grupos[nombre];
      var f = filaSeg(r, nombre);
      var prev = 0;
      ind.dolores.forEach(function (d, j) { if (Motor.segmentoDe(ind, d) === nombre) prev += r.dolores[j].prevalencia; });
      g.notaPrev.textContent = prev > 0
        ? 'Los % dolor ' + (rs.varios ? 'de ' + nombreGrupo(nombre) + ' ' : '') + 'suman ' + pct(prev) + '. Puede pasar de 100 % porque un caso puede tener varios dolores: es normal.'
        : 'Escribe cuántos casos tiene cada dolor: con eso se calcula su % dolor y su impacto.';
      if (f && rs.varios) {
        fijar(g.parte, 'Parte del problema ', h('b', null, pct(f.parte, 1)), f.peso === null ? ' (de sus casos)' : ' (escrita en el paso 1)', ' · mueve ', h('b', null, puntos(f.pp, ind)));
      }
    });
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
    // Si falta un segmento entero (sin dolores o sin casos), el aviso va en su bloque y en «Lo que debes vigilar»
    var segIncompleto = rs.varios && (rs.chequeo !== 'OK' || rs.filas.some(function (f) { return !f.nombre || !f.dolores || (f.peso === null && !(f.casos > 0)); }));
    var noSuma = r.chequeoAtribucion === 'NO SUMA 100' && !segIncompleto;
    refs.notaAtrib.hidden = !noSuma;
    if (noSuma) fijar(refs.notaAtrib, icono('triangulo'), h('span', null, 'El impacto de los dolores suma ', h('b', null, pct(r.sumaAtribucion, 1)), ' y debe sumar 100 %. Si cambias uno a mano, cámbialos todos, o borra los manuales para volver al automático.'));

    // paso 3
    refs.cabDolores.forEach(function (c) {
      var d = buscar(ind.dolores, c.id);
      if (!d) return;
      c.el.textContent = d.nombre || 'Dolor ' + (c.j + 1);
      c.celda.title = c.el.textContent;
    });
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
    if (!ind.segmentos.length) f1.push('el segmento');
    // Paso 2: cada segmento con sus dolores (con nombre) y sus casos analizados, y el impacto sumando 100 %
    var rs = r.segmentos;
    var falta2 = '';
    rs.filas.forEach(function (f) {
      if (falta2) return;
      var donde = rs.varios ? ' de ' + nombreGrupo(f.nombre) : '';
      var lista = doloresDe(ind, f.nombre);
      if (rs.varios && !f.nombre) falta2 = 'Elige el segmento de cada dolor';
      else if (!lista.length) falta2 = rs.varios ? 'Faltan dolores' + donde : 'Sin dolores';
      else if (!(f.casos > 0) && f.peso === null) falta2 = 'Faltan los casos analizados' + donde;
      else if (!lista.some(function (d) { return d.nombre; })) falta2 = 'Faltan nombres' + donde;
    });
    if (!falta2 && rs.chequeo !== 'OK' && rs.chequeo !== 'SIN SEGMENTO') falta2 = 'Revisa la parte de cada segmento';
    if (!falta2 && r.chequeoAtribucion !== 'OK') falta2 = 'Revisa el impacto del dolor';
    var p2ok = !falta2;
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
    // Al entrar, el riel se acaba de pintar con los mismos datos: no se toca su animación
    if (!entrando) actualizarTarjeta(ind, r);
    pintarBarraResumen(ind, r);
    guardarLuego();
  }
  function pintarTodo(entrando) {
    verTodasAlertas = false;
    verTodaLeyenda = false;
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
    if (campo === 'segmento') { if (e.type === 'change') marcarSegmento(t.value, t.checked); return; }
    if (campo === 'peso') { cambiarPeso(t, e.type === 'change'); return; }
    if (campo === 'segmentoDolor') { if (e.type === 'change') moverDolor(t.dataset.dolor, t.value); return; }
    if (campo === 'copiarDolores') { if (e.type === 'change' && t.value) copiarDolores(t.value, t.dataset.segmento); return; }
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
    } else if (campo === 'casosGrupo') {
      // Casos analizados de un segmento (o, sin segmento, los del indicador)
      var sg = t.dataset.segmento ? buscarSegmento(ind, t.dataset.segmento) : null;
      if (sg) sg.casos = v; else ind.totalCasos = v;
    } else {
      // Al ponerle tu propio nombre, el ejemplo pasa a ser tu cálculo
      if (campo === 'nombre') ind.ejemplo = false;
      ind[campo] = v;
    }
    recalcular(false);
  }
  function alSalir(e) {
    var t = e.target;
    if (!t.dataset || !t.dataset.campo) return;
    var tipo = t.dataset.tipo || 'texto';
    if (t.getAttribute('aria-invalid') === 'true') {
      if (t.dataset.campo !== 'peso') return;
      // Un peso rechazado vuelve al último aceptado (el mensaje sigue a la vista): lo que se ve es lo que se calcula
      t.removeAttribute('aria-invalid');
    }
    var ind = actual();
    var obj = t.dataset.ini ? buscar(ind.iniciativas, t.dataset.ini) : (t.dataset.dolor ? buscar(ind.dolores, t.dataset.dolor) : ind);
    if (!obj) return;
    if (tipo === 'texto' && (t.tagName === 'TEXTAREA' || (t.tagName === 'INPUT' && t.type === 'text'))) {
      var limpio = t.value.replace(/\s*[\r\n]+\s*/g, ' ').trim();
      if (limpio !== t.value) { t.value = limpio; obj[t.dataset.campo] = limpio; recalcular(false); }
      return;
    }
    if (tipo === 'rango' || tipo === 'texto') return;
    var valor = t.dataset.campo === 'efectividad' ? obj.efectividad[t.dataset.dolor]
      : (t.dataset.campo === 'peso' ? (buscarSegmento(ind, t.dataset.segmento) || {}).peso
        : (t.dataset.campo === 'casosGrupo' ? casosDeGrupo(ind, t.dataset.segmento) : obj[t.dataset.campo]));
    t.value = formatoEditable(tipo, valor);
  }

  // ---------------------------------------------------------------- segmento
  // Al quitar segmentos, sus dolores se borran (se puede deshacer). Si no queda ningún segmento, en cambio
  // quedan «sin segmento» con sus casos: así cambiar de segmento (quitar uno y marcar otro) no borra el trabajo.
  // Devuelve cuántos dolores se borraron o quedaron sin segmento.
  function quitarSegmentos(ind, quitados) {
    var nombres = quitados.map(function (s) { return s.nombre; });
    var afectados = ind.dolores.filter(function (d) { return nombres.indexOf(d.segmento) >= 0; });
    ind.segmentos = ind.segmentos.filter(function (s) { return nombres.indexOf(s.nombre) < 0; });
    if (!ind.segmentos.length) {
      var casos = quitados.filter(function (s) { return !vacio(s.casos); });
      if (vacio(ind.totalCasos) && casos.length) ind.totalCasos = casos.reduce(function (t, s) { return t + n(s.casos); }, 0);
      afectados.forEach(function (d) { d.segmento = ''; });
    } else {
      afectados.forEach(function (d) {
        ind.dolores.splice(ind.dolores.indexOf(d), 1);
        ind.iniciativas.forEach(function (x) { delete x.efectividad[d.id]; });
      });
    }
    return afectados.length;
  }
  function textoQuitados(ind, nombres, cuantos) {
    if (!cuantos) return '';
    var uno = cuantos === 1;
    if (ind.segmentos.length) return uno ? ' Se borró su dolor.' : ' Se borraron sus ' + cuantos + ' dolores.';
    return (uno ? ' Su dolor quedó' : ' Sus ' + cuantos + ' dolores quedaron') + ' sin segmento: marca el nuevo segmento y pasan a él.';
  }
  // Se repinta todo el taller (los pasos 2 y 3 dependen de los segmentos) y el foco vuelve al control que se usó
  function repintarSegmentos(selector, valor) {
    pintarTaller();
    recalcular(false);
    var vuelve = $$(selector, D.taller).filter(function (x) { return valor === undefined || x.value === valor; })[0];
    if (vuelve) vuelve.focus({ preventScroll: true });
  }
  function cambiarTipoSegmento(tipo) {
    var ind = actual();
    if (ind.tipoSegmento === tipo) return;
    var antes = foto();
    var permitidos = Motor.segmentosDe(tipo);
    var quitados = ind.segmentos.filter(function (s) { return permitidos.indexOf(s.nombre) < 0; });
    ind.tipoSegmento = tipo;
    var cuantos = quitarSegmentos(ind, quitados);
    Motor.acomodarSegmentos(ind);
    repintarSegmentos('[data-accion="tipo-segmento"][data-valor="' + tipo + '"]');
    if (quitados.length) conDeshacer('Cambiaste a ' + tipo + ': se quitaron los segmentos que no aplican (' + unirLista(quitados.map(function (s) { return s.nombre; })) + ').' + textoQuitados(ind, quitados, cuantos), antes, 9000);
  }
  function marcarSegmento(nombre, si) {
    var ind = actual();
    var s = buscarSegmento(ind, nombre);
    if (si && !s) {
      var orden = Motor.segmentosDe(ind.tipoSegmento);
      var conSueltos = ind.dolores.length && !ind.segmentos.length;
      ind.segmentos.push({ nombre: nombre, peso: null, casos: null });
      ind.segmentos.sort(function (a, b) { return orden.indexOf(a.nombre) - orden.indexOf(b.nombre); });
      Motor.acomodarSegmentos(ind);
      repintarSegmentos('input[data-campo="segmento"]', nombre);
      if (conSueltos) avisar('Tus dolores pasaron a «' + nombre + '». Si marcas otro segmento, llenarás sus dolores aparte.');
    } else if (!si && s) {
      var antes = foto();
      var cuantos = quitarSegmentos(ind, [s]);
      Motor.acomodarSegmentos(ind);
      repintarSegmentos('input[data-campo="segmento"]', nombre);
      if (s.peso !== null || cuantos) conDeshacer('Quitaste «' + nombre + '».' + textoQuitados(ind, [s], cuantos), antes, 9000);
    }
  }
  // «Repartir en partes iguales»: 100 % ÷ número de segmentos marcados
  function repartirIguales() {
    var ind = actual();
    var n = ind.segmentos.length;
    if (n < 2) return;
    var antes = foto();
    var habia = ind.segmentos.some(function (s) { return s.peso !== null; });
    ind.segmentos.forEach(function (s) { s.peso = 1 / n; });
    repintarSegmentos('[data-accion="repartir-pesos"]');
    var txt = 'Repartiste el problema en partes iguales: ' + pctLibre(1 / n) + ' para cada segmento.';
    if (habia) conDeshacer(txt, antes); else avisar(txt);
  }
  // La parte del problema es opcional (vacío = sale de los casos). No se acepta si la suma pasa de 100 %.
  // Mientras se escribe solo se revisa; se guarda al confirmar (change), para no guardar números a medio escribir.
  function cambiarPeso(t, aplicar) {
    var ind = actual();
    var s = buscarSegmento(ind, t.dataset.segmento);
    if (!s) return;
    var v = leerNumero(t.value);
    var motivo = Number.isNaN(v) ? 'Escribe solo el número. Ej.: 60' : '';
    if (!motivo && v !== null) v = v / 100;
    var suma = function (lista) { return lista.reduce(function (a, p) { return a + (p || 0); }, 0); };
    var nuevos = ind.segmentos.map(function (x) { return x === s ? v : x.peso; });
    if (!motivo) motivo = Motor.revisarPesos(nuevos);
    // Si ya venían pasados de 100 % (de un Excel o un respaldo), bajar un peso siempre se acepta
    if (motivo && v !== null && v >= 0 && v <= 1 && suma(nuevos) <= suma(ind.segmentos.map(function (x) { return x.peso; })) + 1e-9) motivo = '';
    refs.segError.textContent = motivo + (motivo && aplicar ? ' No se guardó: queda el peso anterior.' : '');
    refs.segError.hidden = !motivo;
    if (motivo) { t.setAttribute('aria-invalid', 'true'); return; }
    t.removeAttribute('aria-invalid');
    if (!aplicar) return;
    s.peso = v;
    recalcular(false);
  }

  // ---------------------------------------------------------------- acciones
  // + Agregar indicador: suma uno vacío al mismo tablero y deja el taller listo para llenarlo
  function agregarIndicador() {
    var ind = Motor.nuevoIndicador();
    estado.indicadores.push(ind);
    estado.sel = ind.id;
    abiertos = {};
    pintarTodo(true);
    irAlPaso1();
    avisar('Agregaste un indicador nuevo. Llena sus 3 pasos en el taller.');
  }
  function irAlPaso1() { var p1 = $('#paso-1'); irA(p1, p1 ? $('input', p1) : null); }

  // ---------------------------------------------------------------- Mis cálculos
  // El tablero que se ve se guarda solo a cada cambio. «Guardar» lo copia a Mis cálculos para volver a abrirlo.
  function calculoDe(id) { return id ? buscar(calculos, id) : null; }
  function hayCambios() { return textoDatos(estado) !== estado.base; }
  function fechaCorta(d) { return d.toLocaleDateString('es-CO', { day: 'numeric', month: 'short', year: 'numeric' }); }
  function nombreSugerido() {
    var propios = estado.indicadores.filter(function (x) { return x.nombre && !x.ejemplo; });
    var nombre = propios.length ? propios[0].nombre + (propios.length > 1 ? ' y ' + (propios.length - 1) + ' más' : '') : 'Mi cálculo';
    return (nombre + ' · ' + fechaCorta(new Date())).slice(0, 80);
  }
  function guardarCalculo(nombre) {
    var c = calculoDe(estado.calculoId);
    if (!c) { c = { id: Motor.nuevoId('calc-'), nombre: '' }; calculos.push(c); estado.calculoId = c.id; }
    c.nombre = String(nombre || '').trim().slice(0, 80) || c.nombre || nombreSugerido();
    c.fecha = new Date().toISOString();
    c.datos = JSON.parse(JSON.stringify({ sel: estado.sel, params: estado.params, indicadores: estado.indicadores }));
    estado.base = textoDatos(estado);
    // Si el navegador no dejó guardar, ya se avisó: quien llama no debe decir «Guardaste»
    return guardar(true) ? c : null;
  }
  function abrirTablero(datos, id) {
    estado.indicadores = JSON.parse(JSON.stringify(datos.indicadores));
    estado.params = Motor.copiarParams(datos.params);
    estado.sel = estado.indicadores.some(function (x) { return x.id === datos.sel; }) ? datos.sel : estado.indicadores[0].id;
    estado.calculoId = id || null;
    estado.base = textoDatos(estado);
    abiertos = {};
    pintarTodo(true);
    guardar();
  }
  // Si el tablero tiene cambios sin guardar, pregunta antes de cambiarlo. seguir(antes, guardado):
  // antes = cómo estaba todo (para Deshacer si eligió «No guardar»); guardado = el cálculo que se guardó.
  function pedirGuardar(seguir) {
    if (!hayCambios()) { seguir(null, null); return; }
    var c = calculoDe(estado.calculoId);
    preguntar({
      titulo: '¿Guardar los cambios del tablero actual?',
      texto: c ? 'Se actualiza «' + c.nombre + '» en Mis cálculos.' : 'Quedará en Mis cálculos para abrirlo cuando quieras.',
      campo: c ? null : nombreSugerido(),
      botones: [{ valor: 'guardar', txt: 'Guardar', clase: 'btn--rojo' }, { valor: 'no', txt: 'No guardar' }, { valor: 'cancelar', txt: 'Cancelar' }]
    }, function (r, nombre) {
      if (r === 'no') { seguir(foto(), null); return; }
      if (r !== 'guardar') return;
      var g = guardarCalculo(nombre);
      // Sin guardar de verdad no se cambia de tablero: así no se pierde lo que había
      if (g) seguir(null, g);
    });
  }
  function textoGuardado(g) { return g ? 'Guardaste «' + g.nombre + '» en Mis cálculos. ' : ''; }
  // Nuevo cálculo: abre un tablero nuevo con un indicador vacío
  function nuevoCalculo() {
    pedirGuardar(function (antes, guardado) {
      var ind = Motor.nuevoIndicador();
      abrirTablero({ sel: ind.id, params: estado.params, indicadores: [ind] }, null);
      irAlPaso1();
      var txt = textoGuardado(guardado) + 'Empezaste un cálculo nuevo: llena el paso 1.';
      if (antes) conDeshacer(txt + ' No guardaste el anterior.', antes, 12000);
      else avisar(txt, null, 7000);
    });
  }
  function abrirCalculo(id) {
    if (id === estado.calculoId && !hayCambios()) return;
    pedirGuardar(function (antes, guardado) {
      var c = calculoDe(id);
      if (!c) return;
      abrirTablero(c.datos, c.id);
      window.scrollTo({ top: 0, behavior: movimientoReducido.matches ? 'auto' : 'smooth' });
      var txt = textoGuardado(guardado) + 'Abriste «' + c.nombre + '».';
      if (antes) conDeshacer(txt, antes, 12000);
      else avisar(txt);
    });
  }
  function guardarDesdeMenu() {
    var avisarSi = function (g) { if (g) avisar(textoGuardado(g)); };
    if (calculoDe(estado.calculoId)) { avisarSi(guardarCalculo()); return; }
    preguntar({
      titulo: 'Guardar en Mis cálculos', texto: 'Ponle un nombre para encontrarlo después.', campo: nombreSugerido(),
      botones: [{ valor: 'guardar', txt: 'Guardar', clase: 'btn--rojo' }, { valor: 'cancelar', txt: 'Cancelar' }]
    }, function (r, nombre) { if (r === 'guardar') avisarSi(guardarCalculo(nombre)); });
  }
  function eliminarCalculo(id) {
    var c = calculoDe(id);
    if (!c) return;
    preguntar({
      titulo: '¿Eliminar «' + c.nombre + '»?', texto: 'Se borra de Mis cálculos y no se puede recuperar. El tablero que estás viendo no cambia.',
      botones: [{ valor: 'eliminar', txt: 'Eliminar', clase: 'btn--rojo' }, { valor: 'cancelar', txt: 'Cancelar' }]
    }, function (r) {
      if (r !== 'eliminar') return;
      calculos.splice(calculos.indexOf(c), 1);
      // El tablero sigue abierto, pero ya no está guardado en ningún lado
      if (estado.calculoId === id) { estado.calculoId = null; estado.base = null; }
      if (guardar(true)) avisar('Eliminaste «' + c.nombre + '» de Mis cálculos.');
    });
  }
  function renombrarCalculo(entradaNombre) {
    var c = calculoDe(entradaNombre.dataset.calculo);
    if (!c) return;
    var nombre = entradaNombre.value.trim().slice(0, 80);
    if (!nombre) { entradaNombre.value = c.nombre; return; }
    c.nombre = nombre;
    if (guardar(true)) avisar('Renombraste el cálculo: «' + nombre + '».');
  }
  // Semáforo de un cálculo guardado: cuántos indicadores alcanzan, están cerca o no alcanzan (sin abrirlo)
  function semaforoDe(datos) {
    var cuenta = {};
    datos.indicadores.forEach(function (ind) { var s = Motor.calcular(ind, datos.params).real.semaforo; cuenta[s] = (cuenta[s] || 0) + 1; });
    return Object.keys(SEMAFORO).filter(function (s) { return cuenta[s]; }).map(function (s) {
      return chip({ txt: cuenta[s] + ' ' + SEMAFORO[s].txt.toLowerCase(), tono: SEMAFORO[s].tono, icono: SEMAFORO[s].icono }, true);
    });
  }
  function duplicarCalculo(id) {
    var c = calculoDe(id);
    if (!c) return;
    calculos.push({ id: Motor.nuevoId('calc-'), nombre: (c.nombre + ' (copia)').slice(0, 80), fecha: new Date().toISOString(), datos: JSON.parse(JSON.stringify(c.datos)) });
    if (guardar(true)) avisar('Duplicaste «' + c.nombre + '». La copia está en Mis cálculos.');
  }
  function pintarCalculos() {
    var c = calculoDe(estado.calculoId);
    var cambios = hayCambios();
    var lista = calculos.slice().sort(function (a, b) { return a.fecha < b.fecha ? 1 : (a.fecha > b.fecha ? -1 : 0); });
    fijar(D.calculos,
      h('div', { class: 'calc-cab' },
        h('p', null, h('small', null, 'Estás viendo'), h('b', null, c ? c.nombre : 'Cálculo sin guardar'),
          h('small', { class: cambios ? 'calc-cambios' : '' }, cambios ? 'Tiene cambios sin guardar' : 'Todo guardado')),
        h('button', { class: 'btn btn--rojo', type: 'button', data: { accion: 'guardar-calculo' } }, 'Guardar')),
      lista.length ? h('ul', { class: 'calc-lista', 'aria-label': 'Cálculos guardados' }, lista.map(function (x) {
        var id = idUnico('calc');
        var fecha = new Date(x.fecha);
        return h('li', { class: 'calc-item' + (x.id === estado.calculoId ? ' abierto' : '') },
          h('label', { class: 'sr', for: id }, 'Nombre del cálculo'),
          h('input', { class: 'calc-nombre', id: id, type: 'text', maxlength: 80, autocomplete: 'off', value: x.nombre, data: { calculo: x.id } }),
          h('small', null, (x.fecha && !isNaN(fecha) ? 'Guardado el ' + fechaLarga(fecha) : 'Sin fecha') + ' · ' + x.datos.indicadores.length + (x.datos.indicadores.length === 1 ? ' indicador' : ' indicadores')),
          h('span', { class: 'calc-semaforo' }, semaforoDe(x.datos)),
          h('div', { class: 'calc-botones' },
            x.id === estado.calculoId ? h('span', { class: 'calc-abierto' }, 'Abierto') : h('button', { class: 'btn-chico-claro', type: 'button', data: { accion: 'abrir-calculo', id: x.id } }, 'Abrir'),
            h('button', { class: 'btn-icono', type: 'button', 'aria-label': 'Duplicar «' + x.nombre + '»', title: 'Duplicar', data: { accion: 'duplicar-calculo', id: x.id } }, icono('copiar')),
            h('button', { class: 'btn-icono peligro', type: 'button', 'aria-label': 'Eliminar «' + x.nombre + '»', title: 'Eliminar', data: { accion: 'eliminar-calculo', id: x.id } }, icono('borrar'))));
      })) : h('p', { class: 'calc-vacio' }, 'Aún no tienes cálculos guardados. Pulsa «Guardar» o, al empezar un «Nuevo cálculo», elige guardar el actual.'),
      h('hr'),
      h('button', { type: 'button', data: { accion: 'descargar-respaldo' } }, icono('bajar'), 'Descargar respaldo (.json)'),
      h('button', { type: 'button', data: { accion: 'cargar-respaldo' } }, icono('subir'), 'Cargar respaldo (.json)'),
      h('p', { class: 'calc-nota' }, 'Todo se guarda solo en este navegador y en esta dirección. Para pasarlo a otro equipo, o del archivo descargado a la página publicada, descarga un respaldo y cárgalo allá.'));
  }

  // Pregunta con botones (y un campo de nombre si hace falta). listo(valor del botón, nombre); Esc = 'cancelar'.
  var alResponder = null;
  function preguntar(op, listo) {
    D.pregTitulo.textContent = op.titulo;
    D.pregTexto.textContent = op.texto || '';
    D.pregCampo.hidden = !op.campo;
    D.pregNombre.value = op.campo || '';
    fijar(D.pregBotones, op.botones.map(function (b) { return h('button', { class: 'btn' + (b.clase ? ' ' + b.clase : ''), type: 'submit', value: b.valor }, b.txt); }));
    alResponder = listo;
    abrirDialogo(D.dlgPregunta);
    if (op.campo) { D.pregNombre.focus(); D.pregNombre.select(); }
  }
  // Muestra el ejemplo lleno; si lo borraste, lo vuelve a agregar sin tocar tus cálculos
  function verEjemplo() {
    var ej = estado.indicadores.filter(function (x) { return x.ejemplo; })[0];
    if (!ej) {
      ej = Motor.normalizarIndicador(crearEjemplos()[0]);
      if (estado.indicadores.some(function (x) { return x.id === ej.id; })) ej.id = Motor.nuevoId('ind-');
      estado.indicadores.unshift(ej);
      avisar('Agregaste el ejemplo lleno «' + ej.nombre + '». Cámbialo sin miedo: tus cálculos no se tocan.');
    }
    estado.sel = ej.id;
    abiertos = {};
    pintarTodo(true);
    window.scrollTo({ top: 0, behavior: movimientoReducido.matches ? 'auto' : 'smooth' });
  }
  function duplicar() {
    var ind = actual();
    var copia = JSON.parse(JSON.stringify(ind));
    copia.id = Motor.nuevoId('ind-');
    copia.ejemplo = false;
    copia.nombre = (ind.nombre || 'Indicador') + ' (copia)';
    copia.iniciativas.forEach(function (x) { x.id = Motor.nuevoId('i-'); });
    estado.indicadores.splice(estado.indicadores.indexOf(ind) + 1, 0, copia);
    estado.sel = copia.id;
    pintarTodo(true);
    avisar('Duplicaste «' + nombreDe(ind) + '».');
  }
  // Elimina el indicador elegido (o el que se está viendo). Si era el último, queda uno vacío.
  function eliminar(id) {
    var antes = foto();
    var ind = (id && buscar(estado.indicadores, id)) || actual();
    var k = estado.indicadores.indexOf(ind);
    var eraElVisto = ind.id === estado.sel;
    estado.indicadores.splice(k, 1);
    if (!estado.indicadores.length) estado.indicadores.push(Motor.nuevoIndicador());
    if (eraElVisto) { estado.sel = estado.indicadores[Math.min(k, estado.indicadores.length - 1)].id; abiertos = {}; }
    pintarTodo(eraElVisto);
    conDeshacer('Eliminaste «' + nombreDe(ind) + '».', antes);
  }
  function agregarDolor(segmento) {
    var ind = actual();
    segmento = buscarSegmento(ind, segmento) ? segmento : '';
    if (doloresDe(ind, segmento).length >= Motor.MAX_DOLORES) return;
    var d = Motor.nuevoDolor(segmento);
    ind.dolores.push(d);
    Motor.ordenarDolores(ind);
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
  // Pasa un dolor a otro segmento: conserva sus datos; sus % dolor e impacto se recalculan con los casos del nuevo
  function moverDolor(id, segmento) {
    var ind = actual();
    var d = buscar(ind.dolores, id);
    if (!d || !buscarSegmento(ind, segmento) || d.segmento === segmento) return;
    if (doloresDe(ind, segmento).length >= Motor.MAX_DOLORES) { avisar('«' + segmento + '» ya tiene ' + Motor.MAX_DOLORES + ' dolores: borra uno antes de mover este.'); pintarTaller(); recalcular(false); return; }
    var antes = foto();
    d.segmento = segmento;
    Motor.ordenarDolores(ind);
    pintarTaller();
    recalcular(false);
    var fila = filaPorDato('dolorFila', id);
    if (fila) irA(fila, $('select', fila));
    conDeshacer('Pasaste «' + (d.nombre || 'el dolor') + '» a ' + segmento + '. Revisa sus casos: el % dolor ahora se mide con los casos de ' + segmento + '.', antes, 8000);
  }
  // Copia la lista de dolores de un segmento a otro: nombre y fuente, sin casos (cada segmento tiene los suyos)
  function copiarDolores(desde, hacia) {
    var ind = actual();
    if (!buscarSegmento(ind, desde) || !buscarSegmento(ind, hacia)) return;
    var antes = foto();
    var ya = doloresDe(ind, hacia).map(function (d) { return d.nombre.trim().toLowerCase(); });
    var lugar = Motor.MAX_DOLORES - doloresDe(ind, hacia).length;
    var nuevos = doloresDe(ind, desde).filter(function (d) { return !d.nombre || ya.indexOf(d.nombre.trim().toLowerCase()) < 0; }).slice(0, Math.max(0, lugar));
    nuevos.forEach(function (d) {
      var c = Motor.nuevoDolor(hacia);
      c.nombre = d.nombre;
      c.evidencia = d.evidencia;
      ind.dolores.push(c);
    });
    Motor.ordenarDolores(ind);
    pintarTaller();
    recalcular(false);
    var bloque = $$('[data-grupo-bloque]', D.taller).filter(function (b) { return b.dataset.grupoBloque === hacia; })[0];
    var primero = nuevos.length ? $$('[data-dolor-fila]', bloque).slice(-nuevos.length)[0] : null;
    if (primero) irA(primero, $('input[data-campo="casos"]', primero));
    if (nuevos.length) conDeshacer('Copiaste ' + (nuevos.length === 1 ? '1 dolor' : nuevos.length + ' dolores') + ' de ' + desde + ' a ' + hacia + '. Escribe sus casos en ' + hacia + '.', antes, 8000);
    else avisar(hacia + ' ya tiene esos dolores (o no le cabe ninguno más).');
  }
  // «Aplica a» de una iniciativa: enciende o apaga un segmento. Siempre queda al menos uno.
  function alternarAplica(id, segmento) {
    var ind = actual();
    var x = buscar(ind.iniciativas, id);
    if (!x || !buscarSegmento(ind, segmento)) return;
    var todos = ind.segmentos.map(function (s) { return s.nombre; });
    var ahora = x.segmentos.length ? x.segmentos.slice() : todos.slice();
    var k = ahora.indexOf(segmento);
    if (k >= 0) ahora.splice(k, 1); else ahora.push(segmento);
    if (!ahora.length) { avisar('La iniciativa debe aplicar al menos a un segmento.'); return; }
    x.segmentos = ahora.length === todos.length ? [] : todos.filter(function (nombre) { return ahora.indexOf(nombre) >= 0; });
    pintarTaller();
    recalcular(false);
    var vuelve = $$('[data-accion="aplica-ini"]', D.taller).filter(function (b) { return b.dataset.id === id && b.dataset.segmento === segmento; })[0];
    if (vuelve) vuelve.focus({ preventScroll: true });
  }
  function agregarIni() {
    var ind = actual();
    if (ind.iniciativas.length >= Motor.maxIniciativas(ind)) return;
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
    if (tipo === 'ejemplo') { verEjemplo(); return; }
    if (tipo === 'paso') { var sec = $('#paso-' + id); irA(sec, sec ? $('input, button', sec) : null); return; }
    if (tipo === 'dolor') { var fd = filaPorDato('dolorFila', id); irA(fd, fd ? $('input', fd) : null); return; }
    if (tipo === 'grupo') {
      var bl = $$('[data-grupo-bloque]', D.taller).filter(function (b) { return b.dataset.grupoBloque === (id || ''); })[0] || $('#paso-2');
      irA(bl, bl ? $('input, select', bl) : null);
      return;
    }
    if (tipo === 'ini' || tipo === 'detalle') {
      var fi = filaPorDato('iniFila', id);
      if (!fi) return;
      if (tipo === 'detalle') { var det = alternarDetalle(id, true); irA(fi, det ? $('.c-evidencia input', det) : null); return; }
      irA(fi, $('input', fi));
    }
  }

  function conDeshacer(texto, antes, ms) { deshacer = antes; avisar(texto, 'Deshacer', ms); }
  // accion: lo que hace el botón del aviso (por defecto, Deshacer)
  function avisar(texto, boton, ms, accion) {
    clearTimeout(tAviso);
    if (!boton || accion) deshacer = null;
    fijar(D.aviso, h('span', null, texto), boton ? h('button', { type: 'button', data: { accion: accion || 'deshacer' } }, boton) : null);
    D.aviso.classList.add('ver');
    tAviso = setTimeout(function () { D.aviso.classList.remove('ver'); deshacer = null; }, ms || (boton ? 8000 : 4500));
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
      ind.segmentos.length > 1 ? h('span', { class: 'tip-linea' }, 'Segmento: ', h('b', null, nombreGrupo(Motor.segmentoDe(ind, d)))) : null,
      h('span', { class: 'tip-valor' }, pct(dr.resueltoReal) + ' resuelto'),
      h('span', { class: 'tip-linea' }, 'Si todo se cumple: ', h('b', null, pct(dr.resueltoPlan))),
      h('span', { class: 'tip-linea' }, 'Impacto del dolor: ', h('b', null, pct(dr.atribucion, 1))),
      h('span', { class: 'tip-linea' }, 'Recupera: ', h('b', null, puntos(dr.ppReal, ind)))
    ];
  }

  // ---------------------------------------------------------------- ajustes (parámetros)
  var PARAMS_UI = [
    { k: 'verde', t: 'Verde (alcanza) desde', que: 'Desde esta suficiencia el indicador queda en «Alcanza».', get: function (p) { return p.semaforo[2].desde; }, set: function (p, v) { p.semaforo[2].desde = v; } },
    { k: 'amarillo', t: 'Amarillo (cerca) desde', que: 'Desde aquí y hasta el verde: «Cerca». Por debajo: «No alcanza».', get: function (p) { return p.semaforo[1].desde; }, set: function (p, v) { p.semaforo[1].desde = v; } },
    { k: 'umbralReemplazar', t: 'Reemplazar si aporta menos de (% de la brecha)', que: 'Una iniciativa que aporta menos que esto (o con 0 % de ejecución) se marca «Reemplazar».' },
    { k: 'coberturaAtribMin', t: 'Sin cobertura: dolor con impacto mayor a', que: 'Un dolor con más impacto que esto…' },
    { k: 'coberturaResueltoMin', t: '… y que queda resuelto menos de', que: '…y que queda resuelto menos que esto se marca «Sin cobertura».' },
    { k: 'tope', t: 'Tope de efectividad real (ninguna iniciativa resuelve el 100 %)', que: 'Ninguna iniciativa resuelve el 100 % de un dolor. También evita divisiones entre cero.' }
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
  function fechaISO(d) { return d.getFullYear() + '-' + dos(d.getMonth() + 1) + '-' + dos(d.getDate()); }
  function descargar(blob, nombre) {
    var url = URL.createObjectURL(blob);
    var a = h('a', { href: url, download: nombre });
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }
  // Respaldo: todo lo de este navegador (Mis cálculos y el tablero actual) en un .json
  function descargarRespaldo() {
    var datos = Object.assign({ app: 'Suficiencia 2027', tipo: 'respaldo', exportado: new Date().toISOString() }, espacio());
    descargar(new Blob([JSON.stringify(datos, null, 2)], { type: 'application/json' }), 'suficiencia-respaldo-' + fechaISO(new Date()) + '.json');
    try { window.localStorage.setItem(CLAVE_RESPALDO, datos.exportado); } catch (e) { /* sin almacenamiento: el recordatorio vuelve a salir */ }
    avisar('Descargaste el respaldo: el tablero actual y ' + calculos.length + (calculos.length === 1 ? ' cálculo guardado' : ' cálculos guardados') + '. Guárdalo en una carpeta segura.');
  }

  // Exportar a Excel: primero se eligen los indicadores. Si ya tienes cálculos propios, el ejemplo queda sin marcar.
  function abrirExportar() {
    var propios = estado.indicadores.some(function (x) { return !x.ejemplo; });
    fijar(D.expLista, estado.indicadores.map(function (ind) {
      var r = Motor.calcular(ind, estado.params);
      var id = idUnico('exp');
      return h('li', null, h('label', { class: 'exp-item', for: id },
        h('input', { type: 'checkbox', id: id, value: ind.id, checked: propios ? !ind.ejemplo : true }),
        h('span', { class: 'exp-nombre' }, nombreDe(ind), ind.ejemplo ? h('span', { class: 'etq-ejemplo' }, 'Ejemplo') : null),
        h('span', { class: 'exp-dato' }, r.brecha > 0 ? pct(r.real.suficiencia) : ''),
        chip(SEMAFORO[r.real.semaforo], true)));
    }));
    actualizarBotonExportar();
    abrirDialogo(D.dlgExportar);
  }
  function idsExportar() { return $$('input[type="checkbox"]', D.expLista).filter(function (c) { return c.checked; }).map(function (c) { return c.value; }); }
  function marcarExportar(si) { $$('input[type="checkbox"]', D.expLista).forEach(function (c) { c.checked = si; }); actualizarBotonExportar(); }
  function actualizarBotonExportar() {
    var k = idsExportar().length;
    D.expOk.disabled = !k;
    D.expOk.textContent = k ? 'Exportar ' + k + (k === 1 ? ' indicador' : ' indicadores') : 'Elige al menos uno';
  }
  function nombreArchivo(t) {
    return String(t).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\w-]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 40) || 'indicador';
  }
  function exportarExcel(ids) {
    var inds = estado.indicadores.filter(function (x) { return ids.indexOf(x.id) >= 0; });
    if (!inds.length) return;
    var ahora = new Date();
    var archivo = 'Suficiencia_CX_' + (inds.length === 1 ? nombreArchivo(nombreDe(inds[0])) + '_' : '') + fechaISO(ahora) + '.xlsx';
    var libro;
    try { libro = libroExcel(inds, ahora); } catch (e) { avisar('No se pudo armar el Excel. Intenta de nuevo.'); return; }
    Excel.crearXlsx(libro).then(function (bytes) {
      descargar(new Blob([bytes], { type: Excel.TIPO }), archivo);
      avisar('Exportaste ' + inds.length + (inds.length === 1 ? ' indicador' : ' indicadores') + ' a ' + archivo + '. Ábrelo en Excel: empieza por la hoja Resumen.');
    }).catch(function () { avisar('No se pudo crear el archivo de Excel. Intenta de nuevo.'); });
  }

  // ---------------------------------------------------------------- el Excel que se exporta
  // Resumen, una hoja por indicador (el resultado arriba y los datos abajo), Ajustes y Glosario.
  // Celdas azules = datos; grises = resultados. El importador encuentra cada dato por su título, así que
  // el archivo se puede editar en Excel y volver a abrir aquí.
  var MARCA_HOJA = 'SUFICIENCIA 2027 · HOJA DE INDICADOR';
  var MARCA_AJUSTES = 'SUFICIENCIA 2027 · AJUSTES';
  var XC = { rojo: 'DA291C', noche: '141414', gris: '737373', linea: 'BFBFBF', suave: 'D9D9D9', calc: 'F2F2F2', cab: 'E7E6E6', azul: 'DDEBF7', azulTinta: '003399' };
  var XTONO = { ok: ['E2EFDA', '006100'], alerta: ['FFF2CC', '9C6500'], critico: ['FDE2E1', 'C00000'], neutro: ['F2F2F2', '595959'] };
  // Símbolos que Excel dibuja con su fuente (⏱ sale como un cuadro vacío; ◷ sí se ve)
  var XICONO = { check: '✔', triangulo: '⚠', x: '✖', bandera: '⚑', menos: '○', reloj: '◷', cambiar: '⇄', ayuda: '⚠', info: 'ℹ', ojo: 'ℹ' };
  var XE = {
    cejilla: { negrita: true, tam: 9, color: XC.rojo },
    titulo: { negrita: true, tam: 20, color: XC.noche },
    frase: { tam: 11, color: '404040' },
    nota: { tam: 9, color: XC.gris },
    banda: { negrita: true, tam: 11, color: 'FFFFFF', fondo: XC.noche },
    cab: { negrita: true, color: XC.noche, fondo: XC.cab, borde: XC.linea, ajustar: true, h: 'center' },
    cabIzq: { negrita: true, color: XC.noche, fondo: XC.cab, borde: XC.linea, ajustar: true },
    etiqueta: { negrita: true }
  };
  var REGLA_PCT = { tipo: 'decimal', min: 0, max: 1, titulo: 'Porcentaje', mensaje: 'Escribe un porcentaje entre 0 % y 100 % (por ejemplo 40 %).' };
  var REGLA_CASOS = { tipo: 'entero', min: 0, titulo: 'Casos', mensaje: 'Escribe un número entero mayor o igual a 0.' };
  function xEntrada(formato, extra) { return Object.assign({ fondo: XC.azul, color: XC.azulTinta, borde: XC.linea, formato: formato || undefined }, extra); }
  function xCalc(formato, extra) { return Object.assign({ fondo: XC.calc, borde: XC.suave, formato: formato || undefined }, extra); }
  function xChip(tono) { var t = XTONO[tono] || XTONO.neutro; return { negrita: true, fondo: t[0], color: t[1], borde: XC.suave, h: 'center' }; }
  function xTextoChip(meta) { return meta ? (XICONO[meta.icono] || '') + ' ' + meta.txt : ''; }
  function xFormato(unidad) { var u = String(unidad || '').replace(/"/g, '').trim(); return u ? '0.00" ' + u + '"' : '0.00'; }
  // Alto de fila para que se lea un texto ajustado a un ancho (en caracteres): parte por palabras, como Excel
  function xAlto(texto, ancho, negrita) {
    var cabe = Math.max(4, ancho * (negrita ? 0.95 : 1.1));
    var lineas = 0;
    String(texto === null || texto === undefined ? '' : texto).split('\n').forEach(function (parrafo) {
      var n = 1;
      var largo = 0;
      parrafo.split(/\s+/).forEach(function (palabra) {
        if (!palabra) return;
        if (largo && largo + 1 + palabra.length > cabe) { n += 1; largo = palabra.length; }
        else largo = largo ? largo + 1 + palabra.length : palabra.length;
        while (largo > cabe) { n += 1; largo -= cabe; }
      });
      lineas += n;
    });
    return Math.ceil(lineas * 12.75 + 4);
  }
  function xBanda(X, f, texto, ultima) {
    X.poner(f, 0, texto, XE.banda);
    X.pintar(f, 1, f, ultima, XE.banda);
    X.alto(f, 22);
    return f + 1;
  }
  // celdas: [[columna, texto, alineado a la izquierda]]
  function xCabecera(X, f, celdas) {
    var alto = 18;
    celdas.forEach(function (c) {
      X.poner(f, c[0], c[1], c[2] ? XE.cabIzq : XE.cab);
      alto = Math.max(alto, xAlto(c[1], (X.anchos[c[0]] || 9) - 1, true));
    });
    X.alto(f, alto);
    return f + 1;
  }
  function fechaLarga(d) {
    return d.toLocaleDateString('es-CO', { day: 'numeric', month: 'long', year: 'numeric' }) + ' a las ' + d.toLocaleTimeString('es-CO', { hour: 'numeric', minute: '2-digit' });
  }

  function libroExcel(inds, ahora) {
    var usados = ['Resumen', 'Ajustes', 'Glosario'];
    var datos = inds.map(function (ind, k) {
      return { ind: ind, r: Motor.calcular(ind, estado.params), hoja: Excel.nombreHoja(dos(k + 1) + ' ' + nombreDe(ind), usados) };
    });
    var hojas = [hojaResumen(datos, ahora)]
      .concat(datos.map(function (x) { return hojaIndicador(x.ind, x.r, x.hoja); }))
      .concat([hojaAjustes(), hojaGlosario()]);
    return { hojas: hojas, activa: 0, titulo: 'Suficiencia 2027 · ' + (inds.length === 1 ? nombreDe(inds[0]) : inds.length + ' indicadores') };
  }

  function hojaResumen(datos, ahora) {
    var X = new Excel.Hoja('Resumen');
    X.anchos = [5, 42, 20, 12, 11, 11, 11, 13, 15, 16, 14, 12, 13, 9, 11, 9, 13];
    X.color = XC.noche;
    X.horizontal = true;
    var ultima = X.anchos.length - 1;
    var f = 0;
    X.poner(f++, 0, 'SUFICIENCIA 2027 · RESUMEN', XE.cejilla);
    X.poner(f, 0, '¿Me alcanzan las iniciativas para la meta 2027?', XE.titulo);
    X.alto(f++, 32);
    X.poner(f++, 0, 'Exportado el ' + fechaLarga(ahora) + ' desde la calculadora de suficiencia · ' + datos.length + (datos.length === 1 ? ' indicador' : ' indicadores') + '.', XE.nota);
    f = xBanda(X, f + 1, 'CÓMO LEER ESTE ARCHIVO', ultima);
    [
      'Suficiencia = puntos que recuperan tus iniciativas ÷ brecha a la meta. 100 % o más quiere decir que alcanza.',
      '«Esperada» usa el % de ejecución que escribiste. «Si todo se cumple» supone todas las iniciativas al 100 %.',
      '«Mueve el indicador» = cuántos puntos se mueve el indicador con tus iniciativas (suficiencia × brecha).',
      'Cada indicador tiene su hoja con todo el detalle: resultado, el cálculo paso a paso, alertas, segmento, dolores e iniciativas. Usa «Ver hoja ›».',
      'Celdas azules = datos que escribes. Celdas grises = resultados. Si cambias datos, abre el archivo en la calculadora (Importar Excel) para recalcular.'
    ].forEach(function (t) { X.poner(f++, 1, '•  ' + t); });
    f = xBanda(X, f + 1, 'COMPARATIVO DE INDICADORES', ultima);
    f = xCabecera(X, f, ['N°', 'Indicador', 'Segmento', '¿Qué es mejor?', 'Hoy', 'Meta 2027', 'Brecha', 'Suficiencia esperada', 'Semáforo', 'Mueve el indicador',
      'Así quedarías en 2027', 'Te faltaría', 'Si todo se cumple', 'Dolores', 'Iniciativas', 'Alertas', 'Detalle'].map(function (t, c) { return [c, t, c === 1 || c === 2]; }));
    datos.forEach(function (x, k) {
      var ind = x.ind;
      var r = x.r;
      var sem = SEMAFORO[r.real.semaforo];
      var con = r.real.semaforo !== 'SIN DATOS';
      var cumple = r.real.semaforo === 'META CUMPLIDA';
      var fPts = xFormato(unidadPuntos(ind));
      var fVal = xFormato(ind.unidad);
      var seg = ind.segmentos.length ? ind.tipoSegmento + ' · ' + textoSegmentos(ind) : 'Sin segmento';
      X.poner(f, 0, k + 1, xCalc(null, { h: 'center', negrita: true }));
      X.poner(f, 1, nombreDe(ind) + (ind.ejemplo ? '  (ejemplo)' : ''), xCalc(null, { negrita: true, ajustar: true }));
      X.poner(f, 2, seg, xCalc(null, { ajustar: true }));
      X.poner(f, 3, ind.sentido === 'sube' ? '↑ Que suba' : (ind.sentido === 'baja' ? '↓ Que baje' : '—'), xCalc(null, { h: 'center' }));
      X.poner(f, 4, vacio(ind.ejecutado) ? '—' : ind.ejecutado, xCalc(fVal, { h: 'center' }));
      X.poner(f, 5, vacio(ind.meta) ? '—' : ind.meta, xCalc(fVal, { h: 'center' }));
      X.poner(f, 6, con ? r.brecha : '—', xCalc(fPts, { h: 'center' }));
      X.poner(f, 7, con && !cumple ? r.real.suficiencia : '—', xCalc('0%', { h: 'center', negrita: true, tam: 11 }));
      X.poner(f, 8, xTextoChip(sem), xChip(sem.tono));
      X.poner(f, 9, con ? r.real.pp : '—', xCalc(fPts, { h: 'center', negrita: true }));
      X.poner(f, 10, con ? r.real.proyectado : '—', xCalc(fVal, { h: 'center' }));
      X.poner(f, 11, con ? r.real.brechaRestante : '—', xCalc(fPts, { h: 'center' }));
      X.poner(f, 12, con && !cumple ? r.plan.suficiencia : '—', xCalc('0%', { h: 'center' }));
      X.poner(f, 13, ind.dolores.length, xCalc(null, { h: 'center' }));
      X.poner(f, 14, ind.iniciativas.length, xCalc(null, { h: 'center' }));
      X.poner(f, 15, alertas(ind, r).length, xCalc(null, { h: 'center' }));
      X.poner(f, 16, 'Ver hoja ›', { negrita: true, subrayado: true, color: XC.rojo, fondo: XC.calc, borde: XC.suave, h: 'center' });
      X.enlazar(f, 16, x.hoja, 'Ver hoja ›');
      X.alto(f++, Math.max(20, xAlto(nombreDe(ind) + (ind.ejemplo ? '  (ejemplo)' : ''), 41, true), xAlto(seg, 19)));
    });
    f = xBanda(X, f + 1, 'QUÉ SIGNIFICA EL SEMÁFORO', ultima);
    etiquetasExplicadas().slice(0, 3).concat([
      [SEMAFORO['META CUMPLIDA'], 'Hoy ya se cumple la meta: no hay brecha que cerrar.'],
      [SEMAFORO['SIN DATOS'], 'Faltan los datos del paso 1 (hoy, meta y qué es mejor).']
    ]).forEach(function (x) {
      X.poner(f, 1, xTextoChip(x[0]), xChip(x[0].tono));
      X.poner(f++, 2, x[1]);
    });
    return X;
  }

  function hojaIndicador(ind, r, nombre) {
    var X = new Excel.Hoja(nombre);
    var nd = ind.dolores.length;
    var D0 = 2;
    var cx = D0 + nd;
    var rs = r.segmentos;
    // Con varios segmentos, el paso 3 lleva la columna «Aplica a» junto a la ejecución
    var claves = ['ejec'].concat(rs.varios ? ['aplica'] : [], ['resp', 'fecha', 'evid', 'aporte', 'aportePct', 'sinElla', 'etq', 'rev']);
    var anchoDe = { ejec: 12, aplica: 22, resp: 18, fecha: 15, evid: 44, aporte: 12, aportePct: 12, sinElla: 13, etq: 15, rev: 20 };
    var col = {};
    claves.forEach(function (k, i) { col[k] = cx + i; });
    var ultima = col.rev;
    var fPts = xFormato(unidadPuntos(ind));
    var fVal = xFormato(ind.unidad);
    var sem = SEMAFORO[r.real.semaforo];
    var con = r.real.semaforo !== 'SIN DATOS';
    var cumple = r.real.semaforo === 'META CUMPLIDA';
    X.anchos = [6, 46];
    for (var k = 0; k < nd; k++) X.anchos[D0 + k] = 17;
    claves.forEach(function (k) { X.anchos[col[k]] = anchoDe[k]; });
    X.color = ind.ejemplo ? 'A6A6A6' : XC.rojo;
    X.congelar = { columnas: 2 };
    X.horizontal = true;
    var f = 0;
    function dato(etq, valor, estilo, nota) {
      X.poner(f, 1, etq, XE.etiqueta);
      X.poner(f, 2, valor, estilo);
      if (nota) X.poner(f, 3, nota, XE.nota);
      return f++;
    }

    // Encabezado: la primera celda marca la hoja para poder importarla de vuelta
    X.poner(f++, 0, MARCA_HOJA, XE.cejilla);
    X.poner(f, 0, nombreDe(ind), XE.titulo);
    X.alto(f++, 32);
    if (ind.ejemplo) {
      X.poner(f, 0, '⚠ EJEMPLO con datos de muestra: reemplázalos por los tuyos.', { negrita: true, color: XTONO.alerta[1], fondo: XTONO.alerta[0] });
      X.pintar(f, 1, f, 6, { fondo: XTONO.alerta[0] });
      f++;
    }
    X.poner(f++, 0, textoPlano(frase(ind, r)), XE.frase);
    X.poner(f++, 0, 'Celdas azules = datos que escribes · Celdas grises = resultados. Si cambias un dato, importa este archivo en la calculadora para recalcular.', XE.nota);

    // Resultado
    f = xBanda(X, f + 1, '¿ME ALCANZA PARA LA META 2027?   ·   RESULTADO', ultima);
    var t = XTONO[sem.tono] || XTONO.neutro;
    X.poner(f, 1, 'Suficiencia esperada', XE.etiqueta);
    X.poner(f, 2, con && !cumple ? r.real.suficiencia : (cumple ? '✔' : '—'), { negrita: true, tam: 16, color: t[1], fondo: t[0], borde: XC.suave, formato: '0%', h: 'center' });
    X.poner(f, 3, xTextoChip(sem), xChip(sem.tono));
    X.poner(f, 4, 'Puntos que recuperan tus iniciativas ÷ brecha. 100 % o más = alcanza.', XE.nota);
    X.alto(f++, 28);
    dato('Mueve el indicador', con ? r.real.pp : '—', xCalc(fPts, { negrita: true, h: 'center' }),
      (con && ind.sentido && r.real.pp > 0 ? (ind.sentido === 'sube' ? '↑ Sube ' : '↓ Baja ') + puntos(r.real.pp, ind) + ': ' : '') + textoMueve(ind, r));
    dato('Así quedarías en 2027', con ? r.real.proyectado : '—', xCalc(fVal, { h: 'center' }), vacio(ind.meta) ? '' : 'Meta 2027: ' + valorConUnidad(ind.meta, ind));
    dato('Te faltaría', con ? r.real.brechaRestante : '—', xCalc(fPts, { h: 'center' }),
      r.brecha > 0 ? (r.real.brechaRestante > 0 ? 'de ' + puntos(r.brecha, ind) + ' de brecha' : 'Cierras toda la brecha') : (cumple ? 'La meta ya se cumple' : ''));
    dato('Si todo se cumple', con && !cumple ? r.plan.suficiencia : '—', xCalc('0%', { h: 'center' }),
      con && !cumple ? 'Con todas las iniciativas al 100 % recuperarías ' + puntos(r.plan.pp, ind) + '.' : '');

    // El cálculo paso a paso, con los números de este indicador
    f = xBanda(X, f + 1, 'CÓMO SE CALCULA LA SUFICIENCIA   ·   PASO A PASO', ultima);
    f = xCabecera(X, f, [[0, '#'], [1, 'Qué se calcula', true], [2, 'Resultado'], [3, 'De dónde sale', true]]);
    var nPaso = 0;
    function pasoCalculo(etq, valor, estilo, como) {
      X.poner(f, 0, ++nPaso, { negrita: true, h: 'center', color: XC.rojo });
      X.poner(f, 1, etq, XE.etiqueta);
      X.poner(f, 2, valor, estilo);
      X.poner(f, 3, como, { color: '404040' });
      X.alto(f++, 18);
    }
    var hoyTxt = valorConUnidad(ind.ejecutado, ind);
    pasoCalculo('Problema total (el pastel)', vacio(ind.ejecutado) ? '—' : r.problema, xCalc(fPts, { h: 'center' }),
      vacio(ind.ejecutado) ? 'Falta el valor de hoy (paso 1).' : 'Distancia entre hoy (' + hoyTxt + ') y el valor ideal (' + valorConUnidad(r.idealUsado, ind) + ').');
    pasoCalculo('Brecha a la meta', con ? r.brecha : '—', xCalc(fPts, { h: 'center' }),
      con ? 'Distancia entre hoy (' + hoyTxt + ') y la meta 2027 (' + valorConUnidad(ind.meta, ind) + '): lo que debes recuperar.' : 'Faltan datos del paso 1.');
    if (rs.varios) {
      pasoCalculo('Parte del problema de cada segmento', 'Ver paso 1', xCalc(null, { h: 'center' }),
        'Casos analizados del segmento ÷ casos de todos los segmentos (o la parte escrita a mano). Hoy suman ' + pct(rs.sumaPartes, 1) + '.');
    }
    pasoCalculo('Impacto de cada dolor', nd ? 'Ver paso 2' : '—', xCalc(null, { h: 'center' }), rs.varios
      ? 'Casos de cada dolor ÷ casos analizados de su segmento; la parte del segmento se reparte entre sus dolores (entre todos suman ' + pct(r.sumaAtribucion, 1) + ').'
      : 'Casos de cada dolor ÷ casos analizados, repartido para que entre todos sumen 100 % (hoy suman ' + pct(r.sumaAtribucion, 1) + ').');
    pasoCalculo('Lo que resuelven las iniciativas', ind.iniciativas.length ? 'Ver paso 3' : '—', xCalc(null, { h: 'center' }),
      'En cada dolor, cada iniciativa resuelve efectividad × ejecución de lo que dejó la anterior.' + (rs.varios ? ' Solo cuenta en los segmentos a los que aplica.' : ''));
    pasoCalculo('Puntos que recuperas (mueve el indicador)', con ? r.real.pp : '—', xCalc(fPts, { h: 'center', negrita: true }),
      'Problema × impacto del dolor × lo resuelto, sumado en todos los dolores.');
    pasoCalculo('Suficiencia = puntos ÷ brecha', con && !cumple ? r.real.suficiencia : '—', xCalc('0%', { h: 'center', negrita: true }),
      con && !cumple ? puntos(r.real.pp, ind) + ' ÷ ' + puntos(r.brecha, ind) + ' = ' + pct(r.real.suficiencia) + '.' : (cumple ? 'La meta ya se cumple: no hay brecha.' : 'Faltan datos del paso 1.'));
    pasoCalculo('Semáforo', xTextoChip(sem), xChip(sem.tono),
      'Desde ' + pct(estado.params.semaforo[2].desde) + ' alcanza; desde ' + pct(estado.params.semaforo[1].desde) + ' está cerca; menos, no alcanza.');
    if (rs.varios) {
      pasoCalculo('Resultado por segmento', 'Ver paso 1', xCalc(null, { h: 'center' }),
        'Cada segmento mueve lo que resuelven las iniciativas en sus dolores; la suma es ' + puntos(rs.total, ind) + ' (lo que mueve el indicador).');
    }
    X.poner(f, 1, 'En palabras simples', XE.etiqueta);
    X.poner(f, 2, textoPlano(frase(ind, r)), XE.frase);
    X.alto(f++, 22);

    // Lo que debes vigilar
    f = xBanda(X, f + 1, 'LO QUE DEBES VIGILAR', ultima);
    var lista = alertas(ind, r);
    if (!lista.length) lista = [{ tono: 'ok', icono: 'check', texto: 'Todo en orden. No hay alertas para este indicador.' }];
    var hasta = Math.min(ultima, 8);
    var ancho = 0;
    for (var c = 1; c <= hasta; c++) ancho += X.anchos[c] || 9;
    lista.forEach(function (a) {
      var tt = XTONO[a.tono] || XTONO.neutro;
      var txt = textoPlano(a.texto);
      X.poner(f, 0, XICONO[a.icono] || '•', { negrita: true, tam: 11, color: tt[1], h: 'center' });
      X.poner(f, 1, txt, { ajustar: true });
      X.unir(f, 1, f, hasta);
      X.alto(f++, Math.max(18, xAlto(txt, ancho)));
    });

    // Quién aporta más
    f = xBanda(X, f + 1, 'QUIÉN APORTA MÁS   ·   RANKING DE INICIATIVAS', ultima);
    if (con && !cumple && r.ranking.length) {
      f = xCabecera(X, f, [[0, '#'], [1, 'Iniciativa', true], [2, 'Aporte (pp)'], [3, 'Aporte (% de la brecha)'], [4, 'Si falla, la suficiencia baja a'], [5, 'Etiqueta']]);
      r.ranking.forEach(function (id, puesto) {
        var x = buscar(r.iniciativas, id);
        var ini = buscar(ind.iniciativas, id);
        var et = ETIQUETA[x.etiqueta];
        X.poner(f, 0, puesto + 1, xCalc(null, { h: 'center', negrita: true }));
        X.poner(f, 1, ini.nombre, xCalc(null, { ajustar: true }));
        X.poner(f, 2, x.aporte, xCalc(fPts, { h: 'center', negrita: true }));
        X.poner(f, 3, x.aportePct, xCalc('0%', { h: 'center' }));
        X.poner(f, 4, x.suficienciaSin, xCalc('0%', { h: 'center' }));
        X.poner(f, 5, xTextoChip(et), xChip(et ? et.tono : 'neutro'));
        X.alto(f++, Math.max(20, xAlto(ini.nombre, 45)));
      });
    } else {
      X.poner(f++, 1, cumple ? 'La meta ya se cumple: no hay brecha que repartir.' : (con ? 'Aún no hay iniciativas.' : 'Completa el paso 1 para ver cuánto aporta cada iniciativa.'), XE.nota);
    }

    // Paso 1
    f = xBanda(X, f + 1, 'PASO 1 · ¿CUÁNTO ME FALTA?', ultima);
    X.poner(f, 1, 'Nombre del indicador', XE.etiqueta);
    X.poner(f++, 2, ind.nombre, xEntrada(null, { negrita: true }));
    var fSentido = dato('¿Qué es mejor?', ind.sentido === 'sube' ? 'Que suba' : (ind.sentido === 'baja' ? 'Que baje' : ''), xEntrada(null, { h: 'center' }), 'Elige de la lista: Que suba o Que baje.');
    X.validar(fSentido, 2, fSentido, 2, { tipo: 'lista', valores: ['Que suba', 'Que baje'] });
    dato('Unidad', ind.unidad, xEntrada(null, { h: 'center' }), 'Ej.: %, minutos, puntos NPS.');
    dato('Hoy', ind.ejecutado, xEntrada(null, { h: 'center' }), (ind.unidad ? 'En ' + ind.unidad + '. ' : '') + 'Valor real de hoy. Escribe solo el número, sin el símbolo %.');
    dato('Meta 2027', ind.meta, xEntrada(null, { h: 'center' }), 'Valor que debes alcanzar en 2027.');
    dato('Valor ideal', ind.ideal, xEntrada(null, { h: 'center' }), 'El mejor valor posible. Vacío = automático (' + (ind.sentido === 'sube' ? '100' : '0') + ').');
    dato('Problema total (el pastel)', vacio(ind.ejecutado) ? '—' : r.problema, xCalc(fPts, { h: 'center' }), 'Distancia entre hoy y el valor ideal: todo el problema que explican los dolores.');
    dato('Brecha a la meta', con ? r.brecha : '—', xCalc(fPts, { h: 'center' }), 'Distancia entre hoy y la meta 2027: lo que tus iniciativas deben recuperar.');
    var fTipo = dato('Tipo de segmento', ind.tipoSegmento, xEntrada(null, { h: 'center' }), 'Elige de la lista: UMM, UMC o Transversal (puede tener segmentos de las dos).');
    X.validar(fTipo, 2, fTipo, 2, { tipo: 'lista', valores: Motor.TIPOS_SEGMENTO });
    var segTxt = textoSegmentos(ind);
    var fSeg = dato('Segmentos', segTxt, xEntrada(null, { ajustar: true }), 'Separados por coma. UMM: Hogar, Pospago, Prepago · UMC: Pyme Fijo, Pyme Móvil.');
    X.alto(fSeg, Math.max(18, xAlto(segTxt, (X.anchos[2] || 9) - 1)));
    // Con varios segmentos: casos analizados y parte del problema de cada uno, y lo que mueve cada uno
    if (rs.varios) {
      f = xCabecera(X, f, [[1, 'Segmento', true], [2, 'Casos analizados'], [3, 'Parte del problema escrita (%)'], [4, 'Parte del problema usada'], [5, 'Mueve (pp)'], [6, 'De la brecha']]);
      var fP = f;
      rs.filas.forEach(function (s) {
        X.poner(f, 1, nombreGrupo(s.nombre), XE.etiqueta);
        X.poner(f, 2, s.casos, xEntrada('#,##0', { h: 'center' }));
        X.poner(f, 3, s.nombre ? s.peso : null, s.nombre ? xEntrada('0.0%', { h: 'center' }) : xCalc(null, { h: 'center' }));
        X.poner(f, 4, s.parte, xCalc('0.0%', { h: 'center' }));
        X.poner(f, 5, s.pp, xCalc(fPts, { h: 'center' }));
        X.poner(f++, 6, r.brecha > 0 ? s.suficiencia : '—', xCalc('0%', { h: 'center' }));
      });
      X.validar(fP, 2, f - 1, 2, REGLA_CASOS);
      X.validar(fP, 3, f - 1, 3, REGLA_PCT);
      X.poner(f, 1, 'Total', XE.etiqueta);
      X.poner(f, 4, rs.sumaPartes, xCalc('0.0%', { h: 'center', negrita: true }));
      X.poner(f, 5, rs.total, xCalc(fPts, { h: 'center', negrita: true }));
      X.poner(f++, 6, r.brecha > 0 ? r.real.suficiencia : '—', xCalc('0%', { h: 'center', negrita: true }));
      X.poner(f++, 1, 'Parte del problema escrita: vacía = sale de los casos analizados. Cada segmento mueve lo que resuelven las iniciativas en sus dolores.', XE.nota);
    }

    // Paso 2: los dolores van en columnas (las mismas columnas del paso 3)
    f = xBanda(X, f + 1, 'PASO 2 · ¿QUÉ CAUSA EL PROBLEMA?   (DOLORES)', ultima);
    if (!rs.varios) {
      var fCasos = dato('Casos analizados', rs.filas.length ? rs.filas[0].casos : ind.totalCasos, xEntrada('#,##0', { h: 'center' }), 'Total de casos del análisis de datos. Ej.: clientes que se dieron de baja en el periodo.');
      X.validar(fCasos, 2, fCasos, 2, REGLA_CASOS);
    } else {
      X.poner(f++, 1, 'Los casos analizados de cada segmento están en el paso 1. La fila Segmento dice de qué segmento es cada dolor.', XE.nota);
    }
    f++;
    if (!nd) X.poner(f++, 1, 'Aún no hay dolores. Agrégalos en la calculadora (paso 2).', XE.nota);
    else {
      f = xCabecera(X, f, [[1, 'Dato del dolor', true]].concat(ind.dolores.map(function (d, j) { return [D0 + j, 'D' + (j + 1)]; })));
      var filaDolor = function (etq, valor, estilo, ajustar) {
        X.poner(f, 1, etq, XE.etiqueta);
        var alto = 0;
        ind.dolores.forEach(function (d, j) {
          var v = valor(d, r.dolores[j]);
          X.poner(f, D0 + j, v, estilo(r.dolores[j]));
          if (ajustar) alto = Math.max(alto, xAlto(v, 16, etq === 'Dolor'));
        });
        if (alto) X.alto(f, alto);
        return f++;
      };
      var entradaCentro = function (formato) { return function () { return xEntrada(formato, { h: 'center' }); }; };
      var calcCentro = function (formato, negrita) { return function () { return xCalc(formato, { h: 'center', negrita: !!negrita }); }; };
      filaDolor('Dolor', function (d) { return d.nombre; }, function () { return xEntrada(null, { negrita: true, ajustar: true, h: 'center' }); }, true);
      var fS = rs.varios ? filaDolor('Segmento', function (d) { return nombreGrupo(Motor.segmentoDe(ind, d)); }, entradaCentro(null)) : -1;
      var fC = filaDolor('Casos', function (d) { return d.casos; }, entradaCentro('#,##0'));
      filaDolor('% dolor', function (d, dr) { return dr.prevalencia; }, calcCentro('0%'));
      filaDolor('Impacto del dolor', function (d, dr) { return dr.atribucion; }, calcCentro('0.0%', true));
      var fM = filaDolor('Impacto manual (opcional)', function (d) { return d.atribucionManual; }, entradaCentro('0.0%'));
      filaDolor('Resuelto esperado', function (d, dr) { return dr.resueltoReal; }, calcCentro('0%'));
      filaDolor('Resuelto si todo se cumple', function (d, dr) { return dr.resueltoPlan; }, calcCentro('0%'));
      filaDolor('Puntos que recupera', function (d, dr) { return dr.ppReal; }, calcCentro(fPts));
      filaDolor('Estado', function (d, dr) { return xTextoChip(ALERTA_DOLOR[dr.alerta]); }, function (dr) { return xChip((ALERTA_DOLOR[dr.alerta] || SEMAFORO['SIN DATOS']).tono); });
      filaDolor('Fuente del dato', function (d) { return d.evidencia; }, function () { return xEntrada(null, { ajustar: true, v: 'top' }); }, true);
      X.validar(fC, D0, fC, D0 + nd - 1, REGLA_CASOS);
      if (fS >= 0) X.validar(fS, D0, fS, D0 + nd - 1, { tipo: 'lista', valores: ind.segmentos.map(function (sg) { return sg.nombre; }) });
      X.validar(fM, D0, fM, D0 + nd - 1, REGLA_PCT);
      X.poner(f++, 1, 'Los % dolor pueden sumar más de 100 % porque un caso puede tener varios dolores. El impacto del dolor siempre suma 100 % (hoy suma ' +
        pct(r.sumaAtribucion, 1) + ') y es el tamaño de cada porción del pastel.', XE.nota);
    }

    // Paso 3: la matriz de iniciativas × dolores
    f = xBanda(X, f + 1, 'PASO 3 · ¿QUÉ VAS A HACER?   (INICIATIVAS)', ultima);
    X.poner(f++, 1, 'En cada columna D escribe qué % de ese dolor resuelve la iniciativa si se hace completa (vacío = 0 %). La ejecución es cuánto estará listo de verdad en 2027.' +
      (rs.varios ? ' «Aplica a» dice a qué segmentos llega la iniciativa: Todos, o los nombres separados por coma.' : ''), XE.nota);
    f = xCabecera(X, f, [[0, 'N°'], [1, 'Iniciativa', true]]
      .concat(ind.dolores.map(function (d, j) { return [D0 + j, 'D' + (j + 1) + ' · ' + (d.nombre || 'Dolor ' + (j + 1))]; }))
      .concat([[col.ejec, 'Ejecución a 2027']], rs.varios ? [[col.aplica, 'Aplica a (segmentos)', true]] : [], [[col.resp, 'Responsable'], [col.fecha, 'Fecha comprometida'],
        [col.evid, 'Evidencia que respalda los %', true], [col.aporte, 'Aporte (pp)'], [col.aportePct, 'Aporte (% de la brecha)'],
        [col.sinElla, 'Si falla, la suficiencia baja a'], [col.etq, 'Etiqueta'], [col.rev, 'Revisión']]));
    var f0 = f;
    ind.iniciativas.forEach(function (x, i) {
      var res = r.iniciativas[i];
      var activa = r.brecha > 0 && !vacio(x.nombre);
      var et = activa ? ETIQUETA[res.etiqueta] : null;
      var rev = res.revision === 'OK' ? { txt: 'Completa', tono: 'ok', icono: 'check' } : REVISION[res.revision];
      X.poner(f, 0, 'I' + (i + 1), { negrita: true, h: 'center' });
      X.poner(f, 1, x.nombre, xEntrada(null, { ajustar: true }));
      ind.dolores.forEach(function (d, j) {
        if (!Motor.aplica(ind, x, d)) X.poner(f, D0 + j, 'No aplica', xCalc(null, { h: 'center', color: XC.gris }));
        else X.poner(f, D0 + j, vacio(x.efectividad[d.id]) ? null : x.efectividad[d.id], xEntrada('0%', { h: 'center' }));
      });
      X.poner(f, col.ejec, x.ejecucion, xEntrada('0%', { h: 'center', negrita: true }));
      if (rs.varios) X.poner(f, col.aplica, x.segmentos.length ? x.segmentos.join(', ') : 'Todos', xEntrada(null, { ajustar: true }));
      X.poner(f, col.resp, x.responsable, xEntrada(null, { ajustar: true }));
      X.poner(f, col.fecha, Excel.serieDeFecha(x.fecha), xEntrada('fecha', { h: 'center' }));
      X.poner(f, col.evid, x.evidencia, xEntrada(null, { ajustar: true }));
      X.poner(f, col.aporte, activa ? res.aporte : '—', xCalc(fPts, { h: 'center', negrita: true }));
      X.poner(f, col.aportePct, activa ? res.aportePct : '—', xCalc('0%', { h: 'center' }));
      X.poner(f, col.sinElla, activa ? res.suficienciaSin : '—', xCalc('0%', { h: 'center' }));
      X.poner(f, col.etq, xTextoChip(et), xChip(et ? et.tono : 'neutro'));
      X.poner(f, col.rev, xTextoChip(rev), xChip(rev ? rev.tono : 'neutro'));
      X.alto(f++, Math.max(20, xAlto(x.nombre, 45), xAlto(x.evidencia, 43), xAlto(x.responsable, 17)));
    });
    if (ind.iniciativas.length) {
      if (nd) X.validar(f0, D0, f - 1, D0 + nd - 1, REGLA_PCT);
      X.validar(f0, col.ejec, f - 1, col.ejec, REGLA_PCT);
    } else {
      X.poner(f + 1, 1, 'Aún no hay iniciativas. Agrégalas en la calculadora (paso 3).', XE.nota);
    }
    return X;
  }

  function hojaAjustes() {
    var X = new Excel.Hoja('Ajustes');
    X.anchos = [5, 60, 13, 80];
    X.color = 'A6A6A6';
    var f = 0;
    X.poner(f++, 0, MARCA_AJUSTES, XE.cejilla);
    X.poner(f, 0, 'Ajustes del semáforo y de las alertas', XE.titulo);
    X.alto(f++, 32);
    X.poner(f++, 0, 'Aplican a todos los indicadores. Si los cambias aquí, importa el archivo en la calculadora para recalcular.', XE.nota);
    f = xCabecera(X, f + 1, [[1, 'Ajuste', true], [2, 'Valor'], [3, 'Qué significa', true]]);
    var f0 = f;
    PARAMS_UI.forEach(function (def) {
      X.poner(f, 1, def.t, XE.etiqueta);
      X.poner(f, 2, leerParam(def, estado.params), xEntrada('0%', { h: 'center' }));
      X.poner(f, 3, def.que, { ajustar: true });
      X.alto(f++, 20);
    });
    X.validar(f0, 2, f - 1, 2, REGLA_PCT);
    return X;
  }

  function hojaGlosario() {
    var X = new Excel.Hoja('Glosario');
    X.anchos = [5, 30, 100];
    X.color = 'A6A6A6';
    var f = 0;
    X.poner(f++, 0, 'SUFICIENCIA 2027 · GLOSARIO', XE.cejilla);
    X.poner(f, 0, 'Las palabras de la suficiencia', XE.titulo);
    X.alto(f++, 32);
    f = xBanda(X, f + 1, 'GLOSARIO', 2);
    $$('#dlg-como .glosario dl > dt').forEach(function (dt) {
      X.poner(f, 1, dt.textContent, XE.etiqueta);
      X.poner(f++, 2, dt.nextElementSibling ? dt.nextElementSibling.textContent : '', { ajustar: true });
    });
    f = xBanda(X, f + 1, 'QUÉ SIGNIFICA CADA ETIQUETA', 2);
    etiquetasExplicadas().forEach(function (x) {
      X.poner(f, 1, xTextoChip(x[0]), xChip(x[0].tono));
      X.poner(f, 2, x[1], { ajustar: true });
      X.alto(f++, 18);
    });
    f = xBanda(X, f + 1, 'CÓMO USAR ESTE ARCHIVO', 2);
    [
      '1. En cada hoja de indicador cambia solo las celdas azules (los datos). Las grises son resultados.',
      '2. Abre la calculadora (suficiencia_cx.html) › Importar Excel › elige este archivo. Todo se recalcula.',
      '3. Los indicadores con el mismo nombre se actualizan y los demás se agregan. Si te equivocas, usa Deshacer.',
      '4. Para agregar un dolor o una iniciativa, hazlo en la calculadora, o inserta una columna o una fila dentro de su tabla.'
    ].forEach(function (t) { X.poner(f++, 1, t); });
    return X;
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

  // Lee el Excel sin internet. Si no se puede (por ejemplo un .xls antiguo), usa SheetJS, que se descarga de internet.
  // Los dos devuelven la misma forma: { nombres: [hojas], hojas: { nombre: { 'A1': { v } } } }.
  function leerLibro(archivo) {
    return archivo.arrayBuffer().then(function (buf) {
      var nativo = /\.xls$/i.test(archivo.name) ? Promise.reject(new Error('xls')) : Excel.leerXlsx(buf);
      return nativo.catch(function () {
        return cargarSheetJS().then(function (XLSX) {
          var wb = XLSX.read(buf, { type: 'array' });
          return { nombres: wb.SheetNames, hojas: wb.Sheets };
        });
      });
    });
  }

  function fechaCelda(celda) {
    if (!celda) return '';
    if (typeof celda.v === 'number') return Excel.fechaDeSerie(celda.v);
    var t = String(celda.v).trim();
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(t);
    if (m) return m[1] + '-' + m[2] + '-' + m[3];
    m = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/.exec(t);
    return m ? m[3] + '-' + dos(Number(m[2])) + '-' + dos(Number(m[1])) : '';
  }

  // Lee las hojas IND… de Plantilla_Suficiencia.xlsx (mismas celdas que la plantilla).
  function leerPlantilla(libro) {
    var inds = [];
    var columnas = ['C', 'D', 'E', 'F', 'G', 'H'];
    libro.nombres.forEach(function (nombreHoja) {
      if (!/^IND/i.test(nombreHoja)) return;
      var ws = libro.hojas[nombreHoja] || {};
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
          estado: txt('J' + fila) || 'En plan', responsable: txt('K' + fila), fecha: fechaCelda(celda('L' + fila)), evidencia: txt('M' + fila)
        });
      }
      // La plantilla no tiene segmento: al actualizar un indicador se conserva el que tenía
      var listo = Motor.normalizarIndicador(ind);
      listo.sinSegmento = true;
      inds.push(listo);
    });
    var params = null;
    var wp = libro.hojas.Parametros;
    if (wp) {
      var pn = function (a) { var c = wp[a]; return c && typeof c.v === 'number' ? c.v : null; };
      params = Motor.normalizarParams({
        semaforo: [{ desde: pn('B6') }, { desde: pn('B7') }, { desde: pn('B8') }],
        umbralReemplazar: pn('C12'), coberturaAtribMin: pn('C13'), coberturaResueltoMin: pn('C14'), tope: pn('C15')
      });
    }
    return { indicadores: inds, params: params };
  }

  // Lee un Excel exportado por esta calculadora: cada dato se busca por su título, no por su celda,
  // así que sigue funcionando aunque el usuario inserte filas o columnas.
  function normal(t) {
    return String(t === null || t === undefined ? '' : t).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9%]+/g, ' ').trim();
  }
  function limites(ws) {
    var lim = { filas: 0, cols: 0 };
    Object.keys(ws).forEach(function (k) {
      var p = k.charAt(0) === '!' ? null : Excel.partirRef(k);
      if (p) { lim.filas = Math.max(lim.filas, p.fila); lim.cols = Math.max(lim.cols, p.col); }
    });
    return lim;
  }
  function valorEn(ws, f, c) {
    if (f < 0 || c < 0) return null;
    var x = ws[Excel.ref(f, c)];
    return x && x.v !== undefined && x.v !== null && x.v !== '' ? x.v : null;
  }
  function numeroCelda(v) {
    if (typeof v === 'number') return isFinite(v) ? v : null;
    if (typeof v !== 'string' || !v.trim()) return null;
    var x = leerNumero(v);
    return x === null || Number.isNaN(x) ? null : x;
  }
  function enteroCelda(v) {
    if (typeof v === 'number') return isFinite(v) ? Math.round(v) : null;
    var x = typeof v === 'string' ? leerEntero(v) : null;
    return x === null || Number.isNaN(x) ? null : x;
  }
  // Porcentajes: 0,45 · «45 %» · 45 → 0,45
  function fraccionCelda(v) {
    var x = numeroCelda(v);
    if (x === null) return null;
    return (typeof v === 'string' && v.indexOf('%') >= 0) || (x > 1 && x <= 100) ? x / 100 : x;
  }
  function estadoDe(t) {
    var buscado = normal(t);
    return Motor.ESTADOS.filter(function (e) { return normal(e) === buscado; })[0] || 'En plan';
  }

  function leerExportado(libro) {
    var inds = [];
    var params = null;
    libro.nombres.forEach(function (nombre) {
      var ws = libro.hojas[nombre] || {};
      var marca = normal(valorEn(ws, 0, 0));
      if (marca === normal(MARCA_HOJA)) {
        var ind = leerHojaIndicador(ws);
        if (ind) inds.push(ind);
      } else if (marca === normal(MARCA_AJUSTES)) params = leerHojaAjustes(ws);
    });
    return { indicadores: inds, params: params };
  }

  function leerHojaIndicador(ws) {
    var lim = limites(ws);
    var v = function (f, c) { return valorEn(ws, f, c); };
    var txt = function (f, c) { var x = v(f, c); return x === null ? '' : String(x).trim(); };
    // Las bandas «PASO 1/2/3» de la columna A separan las secciones
    var paso = {};
    for (var f = 0; f <= lim.filas; f++) {
      var m = /^paso ([123])\b/.exec(normal(v(f, 0)));
      if (m && paso[m[1]] === undefined) paso[m[1]] = f;
    }
    if (paso[1] === undefined) return null;
    var fin = lim.filas + 1;
    var fin1 = paso[2] !== undefined ? paso[2] : fin;
    var fin2 = paso[3] !== undefined ? paso[3] : fin;
    // Fila cuyo título (columna B) es la etiqueta buscada
    function fila(desde, hasta, etiqueta, alInicio) {
      if (desde === undefined) return -1;
      for (var k = desde; k < hasta; k++) {
        var e = normal(v(k, 1));
        if (alInicio ? e.indexOf(etiqueta) === 0 : e === etiqueta) return k;
      }
      return -1;
    }
    function dato(desde, hasta, etiqueta) { return v(fila(desde, hasta, etiqueta), 2); }
    var sentido = normal(dato(paso[1], fin1, 'que es mejor'));
    var ind = {
      nombre: txt(fila(paso[1], fin1, 'nombre del indicador'), 2),
      sentido: /sub/.test(sentido) ? 'sube' : (/baj/.test(sentido) ? 'baja' : ''),
      unidad: txt(fila(paso[1], fin1, 'unidad'), 2),
      ejecutado: numeroCelda(dato(paso[1], fin1, 'hoy')),
      meta: numeroCelda(dato(paso[1], fin1, 'meta 2027')),
      ideal: numeroCelda(dato(paso[1], fin1, 'valor ideal')),
      totalCasos: enteroCelda(dato(paso[2], fin2, 'casos analizados')),
      dolores: [],
      iniciativas: []
    };
    // Segmento: «Tipo de segmento», «Segmentos» (separados por coma) y, con varios segmentos, la tabla con los
    // casos analizados y la parte del problema de cada uno. Un Excel de antes (tabla de pesos sin casos) no trae
    // casos: sus pesos significaban otra cosa y no se usan.
    var tipoSeg = normal(dato(paso[1], fin1, 'tipo de segmento'));
    ind.tipoSegmento = Motor.TIPOS_SEGMENTO.filter(function (t) { return normal(t) === tipoSeg; })[0] || '';
    var fTab = fila(paso[1], fin1, 'segmento');
    var colSeg = {};
    if (fTab >= 0) {
      for (var cs = 2; cs <= lim.cols; cs++) {
        var es = normal(v(fTab, cs));
        if (es.indexOf('casos analizados') === 0) colSeg.casos = cs;
        if (es.indexOf('parte del problema escrita') === 0) colSeg.peso = cs;
      }
    }
    var tabla = colSeg.casos !== undefined;
    var datoSeg = function (nombre, c) { var fs = tabla ? fila(fTab + 1, fin1, normal(nombre)) : -1; return fs < 0 || c === undefined ? null : v(fs, c); };
    ind.segmentos = txt(fila(paso[1], fin1, 'segmentos'), 2).split(/[,;]/).map(function (s) {
      var nombre = Motor.segmentosDe('Transversal').filter(function (x) { return normal(x) === normal(s); })[0];
      if (!nombre) return null;
      return tabla ? { nombre: nombre, peso: fraccionCelda(datoSeg(nombre, colSeg.peso)), casos: enteroCelda(datoSeg(nombre, colSeg.casos)) } : { nombre: nombre, peso: null };
    }).filter(Boolean);
    if (tabla) ind.totalCasos = enteroCelda(datoSeg('Sin segmento', colSeg.casos));
    // Paso 3: el encabezado dice en qué columna va cada dato
    var fCab = fila(paso[3], fin, 'iniciativa');
    var colDe = {};
    if (fCab >= 0) {
      for (var c = 2; c <= lim.cols; c++) {
        var e = normal(v(fCab, c));
        ['ejecucion', 'estado', 'responsable', 'fecha', 'evidencia', 'aplica a'].forEach(function (k) { if (colDe[k] === undefined && e.indexOf(k) === 0) colDe[k] = c; });
      }
    }
    // Paso 2: un dolor por columna, desde la C hasta antes de «Ejecución a 2027» (son las mismas columnas en los dos pasos)
    var columnas = [];
    var fD = fila(paso[2], fin2, 'dolor');
    if (fD >= 0) {
      var fC = fila(paso[2], fin2, 'casos');
      var fS = fila(paso[2], fin2, 'segmento');
      var fM = fila(paso[2], fin2, 'impacto manual', true);
      var fF = fila(paso[2], fin2, 'fuente del dato');
      var tope = colDe.ejecucion !== undefined ? colDe.ejecucion : lim.cols + 1;
      var segDe = function (t) { return ind.segmentos.filter(function (sg) { return normal(sg.nombre) === normal(t); }).map(function (sg) { return sg.nombre; })[0] || ''; };
      for (var cd = 2; cd < tope; cd++) {
        var casos = v(fC, cd);
        var manual = v(fM, cd);
        if (!txt(fD, cd) && casos === null && manual === null && !txt(fF, cd)) continue;
        var id = 'd' + (columnas.length + 1);
        columnas.push({ col: cd, id: id });
        ind.dolores.push({ id: id, segmento: fS >= 0 ? segDe(txt(fS, cd)) : undefined, nombre: txt(fD, cd), casos: enteroCelda(casos), atribucionManual: fraccionCelda(manual), evidencia: txt(fF, cd) });
      }
    }
    var segDeIni = function (t) { return ind.segmentos.filter(function (sg) { return normal(sg.nombre) === normal(t); }).map(function (sg) { return sg.nombre; })[0] || ''; };
    if (fCab >= 0) {
      var col = function (k) { return colDe[k] === undefined ? -1 : colDe[k]; };
      for (var fi = fCab + 1; fi <= lim.filas; fi++) {
        var efect = {};
        var alguna = false;
        columnas.forEach(function (x) {
          var e = fraccionCelda(v(fi, x.col));
          if (e !== null) { efect[x.id] = e; alguna = true; }
        });
        var ejec = v(fi, col('ejecucion'));
        if (!txt(fi, 1) && !alguna && ejec === null) {
          var vacia = true;
          for (var cc = 0; cc <= lim.cols; cc++) if (v(fi, cc) !== null) { vacia = false; break; }
          if (vacia) break;
          continue;
        }
        var aplicaTxt = col('aplica a') >= 0 ? txt(fi, col('aplica a')) : '';
        ind.iniciativas.push({
          id: 'i' + (ind.iniciativas.length + 1), nombre: txt(fi, 1), efectividad: efect, ejecucion: fraccionCelda(ejec),
          segmentos: /^todos?$/.test(normal(aplicaTxt)) ? [] : aplicaTxt.split(/[,;]/).map(function (t) { return segDeIni(t); }).filter(Boolean),
          estado: estadoDe(txt(fi, col('estado'))), responsable: txt(fi, col('responsable')),
          fecha: col('fecha') >= 0 ? fechaCelda(ws[Excel.ref(fi, col('fecha'))]) : '',
          evidencia: txt(fi, col('evidencia'))
        });
      }
    }
    var listo = Motor.normalizarIndicador(ind);
    if (fila(paso[1], fin1, 'tipo de segmento') < 0) listo.sinSegmento = true;
    return listo;
  }

  function leerHojaAjustes(ws) {
    var lim = limites(ws);
    var p = Motor.copiarParams(estado.params);
    var alguno = false;
    for (var f = 0; f <= lim.filas; f++) {
      var etq = normal(valorEn(ws, f, 1));
      if (!etq) continue;
      PARAMS_UI.forEach(function (def) {
        if (etq !== normal(def.t)) return;
        var val = fraccionCelda(valorEn(ws, f, 2));
        if (val === null) return;
        if (def.set) def.set(p, val); else p[def.k] = val;
        alguno = true;
      });
    }
    return alguno ? Motor.normalizarParams(p) : null;
  }

  function incorporar(inds, params, nombreArchivo) {
    var antes = foto();
    var nuevos = 0;
    var actualizados = 0;
    // Cada indicador del tablero se actualiza una sola vez (dos hojas con el mismo nombre, o sin nombre, no se pisan)
    var usados = {};
    inds.forEach(function (ind) {
      var clave = ind.nombre.trim().toLowerCase();
      var existente = estado.indicadores.filter(function (x) { return !usados[x.id] && x.nombre.trim().toLowerCase() === clave; })[0];
      // Un Excel exportado antes de existir el segmento no trae esas filas: se conserva el que había
      if (existente && ind.sinSegmento) { ind.tipoSegmento = existente.tipoSegmento; ind.segmentos = existente.segmentos; Motor.acomodarSegmentos(ind); }
      delete ind.sinSegmento;
      if (existente) { ind.id = existente.id; estado.indicadores[estado.indicadores.indexOf(existente)] = ind; actualizados++; }
      else { ind.id = Motor.nuevoId('ind-'); estado.indicadores.push(ind); nuevos++; }
      usados[ind.id] = true;
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
        var crudo = JSON.parse(texto);
        var e = Motor.normalizarEspacio(crudo);
        if (!e) throw new Error('vacío');
        var antes = foto();
        var copiaVieja = !(crudo.version >= 3);
        if (copiaVieja && !e.actual.indicadores.length) throw new Error('vacío');
        // Un respaldo suma sus cálculos a Mis cálculos (el mismo cálculo se reemplaza) y abre su tablero
        e.calculos.forEach(function (c) {
          var viejo = calculoDe(c.id);
          if (viejo) calculos[calculos.indexOf(viejo)] = c; else calculos.push(c);
        });
        if (e.actual && e.actual.indicadores.length) {
          estado.indicadores = e.actual.indicadores;
          estado.params = e.actual.params;
          estado.sel = e.actual.indicadores.some(function (x) { return x.id === e.actual.sel; }) ? e.actual.sel : e.actual.indicadores[0].id;
          estado.calculoId = e.activo;
          estado.base = e.base;
        }
        abiertos = {};
        pintarTodo(true);
        guardar();
        conDeshacer(copiaVieja ? 'Abriste la copia ' + nombre + ' (' + e.actual.indicadores.length + ' indicadores).'
          : 'Cargaste el respaldo ' + nombre + ': ' + e.calculos.length + (e.calculos.length === 1 ? ' cálculo guardado' : ' cálculos guardados') + ' y su tablero.', antes, 12000);
      }).catch(function () {
        avisar('No se pudo abrir ' + nombre + ': no es un respaldo de esta calculadora. Usa un archivo creado con «Descargar respaldo».');
      });
      return;
    }
    avisar('Leyendo ' + nombre + '…');
    leerLibro(archivo).then(function (libro) {
      // Primero un Excel exportado por esta calculadora; si no, la plantilla Plantilla_Suficiencia.xlsx
      var res = leerExportado(libro);
      if (!res.indicadores.length) res = leerPlantilla(libro);
      if (!res.indicadores.length) {
        avisar('No encontré indicadores en ' + nombre + '. Usa un Excel exportado desde esta calculadora o la plantilla Plantilla_Suficiencia.xlsx.');
        return;
      }
      incorporar(res.indicadores, res.params, nombre);
    }).catch(function () {
      avisar(/\.xls$/i.test(nombre)
        ? 'Para abrir un .xls antiguo se necesita internet. Ábrelo en Excel, guárdalo como .xlsx e inténtalo de nuevo.'
        : 'No se pudo leer ' + nombre + '. Revisa que sea un Excel (.xlsx) y que no esté dañado: ábrelo en Excel y guárdalo de nuevo.');
    });
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
    // Los cuatro pasteles se despliegan uno tras otro al abrir el diálogo
    $$('[data-demo]', D.dlg).forEach(function (svg, k) {
      var c = casos[svg.dataset.demo];
      dibujarPastel(svg, c[0], Motor.calcular(c[0], estado.params), Object.assign({ animar: 'entrada', demora: 150 + k * 170, duracion: 950 }, c[1]));
    });
  }
  // Cada etiqueta con su explicación (para el diálogo y para el Excel)
  function etiquetasExplicadas() {
    var filas = [];
    [SEMAFORO.VERDE, SEMAFORO.AMARILLO, SEMAFORO.ROJO].forEach(function (m, k) {
      var desde = estado.params.semaforo[2 - k].desde;
      filas.push([m, k === 0 ? 'Suficiencia de ' + pct(desde) + ' o más: las iniciativas alcanzan.' : (k === 1 ? 'Desde ' + pct(desde) + ': cerca, con ajustes puede alcanzar.' : 'Menos de ' + pct(estado.params.semaforo[1].desde) + ': faltan iniciativas o más efectividad.')]);
    });
    Object.keys(ETIQUETA).forEach(function (k) { filas.push([ETIQUETA[k], ETIQUETA[k].que]); });
    filas.push([ALERTA_DOLOR['SIN COBERTURA'], ALERTA_DOLOR['SIN COBERTURA'].que]);
    filas.push([REVISION['SIN RESPALDO'], REVISION['SIN RESPALDO'].que]);
    return filas;
  }
  function pintarEtiquetasLista() {
    fijar($('#etiquetas-lista'), etiquetasExplicadas().map(function (f) { return h('li', null, h('span', null, chip(f[0])), h('span', null, f[1])); }));
  }
  function abrirDialogo(dlg) {
    if (typeof dlg.showModal === 'function') dlg.showModal();
    else dlg.setAttribute('open', '');
  }
  function cerrarDialogo(dlg) {
    dlg = dlg || D.dlg;
    if (typeof dlg.close === 'function') dlg.close();
    else dlg.removeAttribute('open');
  }
  function abrirComo() {
    pintarEtiquetasLista();
    cerrarIntro();
    abrirDialogo(D.dlg);
    pintarDemos();
  }
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
      case 'menu':
        if (b.getAttribute('aria-controls') === 'menu-calculos') pintarCalculos();
        alternarMenu(b);
        break;
      case 'seleccionar':
        if (b.dataset.id !== estado.sel) { estado.sel = b.dataset.id; abiertos = {}; pintarTodo(true); }
        break;
      case 'nuevo-calculo': nuevoCalculo(); break;
      case 'agregar-indicador': agregarIndicador(); break;
      case 'guardar-calculo': guardarDesdeMenu(); break;
      case 'abrir-calculo': abrirCalculo(b.dataset.id); break;
      case 'eliminar-calculo': eliminarCalculo(b.dataset.id); break;
      case 'duplicar-calculo': duplicarCalculo(b.dataset.id); break;
      case 'repartir-pesos': repartirIguales(); break;
      case 'descargar-respaldo': descargarRespaldo(); break;
      case 'cargar-respaldo': D.archivo.accept = '.json'; D.archivo.value = ''; D.archivo.click(); break;
      case 'tipo-segmento': cambiarTipoSegmento(b.dataset.valor); break;
      case 'duplicar': duplicar(); break;
      case 'eliminar': eliminar(b.dataset.id); break;
      case 'sentido': actual().sentido = b.dataset.valor; recalcular(false); break;
      case 'agregar-dolor': agregarDolor(b.dataset.segmento); break;
      case 'aplica-ini': alternarAplica(b.dataset.id, b.dataset.segmento); break;
      case 'borrar-dolor': borrarDolor(b.dataset.id); break;
      case 'agregar-ini': agregarIni(); break;
      case 'borrar-ini': borrarIni(b.dataset.id); break;
      case 'detalle-ini': alternarDetalle(b.dataset.id, b.dataset.abrir === '1'); break;
      case 'ir': irADestino(b.dataset.tipo, b.dataset.id); break;
      case 'mas-alertas': verTodasAlertas = !verTodasAlertas; recalcular(false); break;
      case 'mas-leyenda': verTodaLeyenda = !verTodaLeyenda; recalcular(false); break;
      case 'riel-izq': D.riel.scrollBy({ left: -D.riel.clientWidth * 0.8 }); break;
      case 'riel-der': D.riel.scrollBy({ left: D.riel.clientWidth * 0.8 }); break;
      case 'como-funciona': abrirComo(); break;
      case 'cerrar-dialogo': cerrarDialogo(b.closest('dialog')); break;
      case 'cerrar-intro': cerrarIntro(); break;
      case 'importar': D.archivo.accept = '.xlsx,.xlsm,.xls,.json'; D.archivo.value = ''; D.archivo.click(); break;
      case 'exportar-excel': abrirExportar(); break;
      case 'exp-todos': marcarExportar(true); break;
      case 'exp-ninguno': marcarExportar(false); break;
      case 'imprimir': terminarAnimaciones(); window.print(); break;
      case 'ver-ejemplo': verEjemplo(); break;
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
  D.calculos.addEventListener('change', function (e) { if (e.target.classList.contains('calc-nombre')) renombrarCalculo(e.target); });
  // La respuesta se da al pulsar el botón (no al cerrarse el diálogo, que llega después): así no se cruza con otra pregunta
  $('form', D.dlgPregunta).addEventListener('submit', function (e) {
    e.preventDefault();
    var listo = alResponder;
    alResponder = null;
    cerrarDialogo(D.dlgPregunta);
    if (listo) listo((e.submitter && e.submitter.value) || 'cancelar', D.pregNombre.value);
  });
  // Esc o cerrar sin elegir = Cancelar
  D.dlgPregunta.addEventListener('close', function () {
    if (D.dlgPregunta.open) return;
    var listo = alResponder;
    alResponder = null;
    if (listo) listo('cancelar', '');
  });
  // Lo último que se escribió se guarda aunque se cierre la página antes de la pausa del guardado
  window.addEventListener('pagehide', function () { if (tGuardar) guardar(); });
  // Con la página abierta en otra pestaña: se toma su lista de Mis cálculos para no borrarla al guardar aquí
  window.addEventListener('storage', function (e) {
    if (e.key !== CLAVE_DATOS || !e.newValue) return;
    try {
      var otro = Motor.normalizarEspacio(JSON.parse(e.newValue));
      if (otro) calculos = otro.calculos;
    } catch (err) { /* lo de la otra pestaña no se pudo leer: se sigue con lo de esta */ }
  });
  D.expLista.addEventListener('change', actualizarBotonExportar);
  D.formExportar.addEventListener('submit', function (e) {
    e.preventDefault();
    var ids = idsExportar();
    cerrarDialogo(D.dlgExportar);
    exportarExcel(ids);
  });
  D.riel.addEventListener('scroll', function () { window.requestAnimationFrame(actualizarFlechas); }, { passive: true });
  window.addEventListener('resize', actualizarFlechas);
  window.addEventListener('beforeprint', terminarAnimaciones);

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
  if (sinAlmacen) avisar('Este navegador no deja guardar datos: lo que hagas se pierde al cerrar. Usa «Descargar respaldo» en Mis cálculos para no perderlo.', null, 10000);
  else recordarRespaldo();

  // Si tienes cálculos guardados y hace más de 7 días (o nunca) que no descargas un respaldo, se recuerda al abrir
  function recordarRespaldo() {
    if (!calculos.length) return;
    var ultimo = null;
    try { ultimo = window.localStorage.getItem(CLAVE_RESPALDO); } catch (e) { return; }
    var dias = ultimo ? (Date.now() - new Date(ultimo).getTime()) / 86400000 : Infinity;
    if (dias < 7) return;
    avisar(ultimo ? 'Hace más de una semana que no descargas un respaldo de tus cálculos.' : 'Aún no tienes un respaldo de tus cálculos: si se borra el navegador, se pierden.',
      'Descargar respaldo', 15000, 'descargar-respaldo');
  }
})();
