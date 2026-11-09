// Prueba en celular (390 px de ancho). La ejecuta herramientas/probar_tablero.ps1.
// El navegador sin ventana no baja de ~500 px, así que el tablero se abre dentro de un marco de 390 px.
(function () {
  var errores = [];
  window.addEventListener('error', function (e) { errores.push(e.message + ' @' + e.lineno); });
  var out = [];
  function ok(c, m) { out.push((c ? 'OK    ' : 'FALLA ') + m); }
  function terminar() {
    out.push('errores JS: ' + JSON.stringify(errores));
    var pre = document.createElement('pre');
    pre.id = 'resultado-prueba';
    pre.textContent = out.join('\n');
    document.body.appendChild(pre);
  }
  var ANCHO = 390;
  var marco = document.createElement('iframe');
  marco.style.cssText = 'position:fixed;left:0;top:0;width:' + ANCHO + 'px;height:844px;border:0;z-index:999;background:#fff';
  marco.src = new URL('suficiencia_cx.html', document.baseURI).href;
  marco.addEventListener('load', function () { setTimeout(revisar, 400); });
  document.body.appendChild(marco);

  function revisar() {
    var w = marco.contentWindow;
    var d = marco.contentDocument;
    if (!d) { ok(false, 'no se pudo abrir el tablero en el marco'); terminar(); return; }
    w.addEventListener('error', function (e) { errores.push('marco: ' + e.message + ' @' + e.lineno); });
    function $(s) { return d.querySelector(s); }
    function $$(s) { return Array.prototype.slice.call(d.querySelectorAll(s)); }
    function sinDesborde(que) { ok(d.documentElement.scrollWidth <= ANCHO, que + ': sin desplazamiento de lado (' + d.documentElement.scrollWidth + ' px)'); }
    function dentro(el, que) {
      var r = el.getBoundingClientRect();
      ok(r.width > 0 && r.left >= 0 && r.right <= ANCHO + 0.5, que + ' cabe en la pantalla (' + Math.round(r.left) + '–' + Math.round(r.right) + ' px)');
    }
    try {
      ok(w.innerWidth === ANCHO, 'ancho de la pantalla → ' + w.innerWidth + ' px');
      sinDesborde('al abrir');
      $$('.barra-acciones > button, .barra-acciones .menu-envoltura > button').filter(function (b) { return w.getComputedStyle(b).display !== 'none'; }).forEach(function (b) { dentro(b, 'botón «' + b.getAttribute('aria-label') + '»'); });
      var salirMenu = $('#menu-mas [data-cerrar-sesion]');
      ok(w.getComputedStyle($('#cerrar-sesion')).display === 'none' && !!salirMenu && w.getComputedStyle(salirMenu).display !== 'none', 'en celular «Cerrar sesión» pasa al menú «Más opciones»');
      $('[aria-controls="menu-calculos"]').click();
      dentro($('#menu-calculos'), 'Mis cálculos');
      $('.tarjeta--nueva').click();
      $('[data-accion="tipo-segmento"][data-valor="Transversal"]').click();
      ['Hogar', 'Pospago'].forEach(function (n) { $$('#paso-1 input[data-campo="segmento"]').filter(function (x) { return x.value === n; })[0].click(); });
      ok($$('#paso-1 input[data-campo="peso"]').length === 2, 'Transversal con 2 segmentos muestra 2 pesos');
      dentro($('#paso-1 .c-segmento'), 'el bloque de segmento');
      $$('#paso-1 input[data-campo="peso"]').forEach(function (x) { dentro(x, 'el campo de peso'); });
      sinDesborde('con el segmento Transversal');
      $('.barra [data-accion="nuevo-calculo"]').click();
      dentro($('#dlg-pregunta'), 'la pregunta «¿Guardar los cambios?»');
      $$('#preg-botones button').forEach(function (b) { dentro(b, 'botón «' + b.textContent + '»'); });
    } catch (e) { out.push('EXCEPCION ' + e.message + ' | ' + e.stack); }
    terminar();
  }
})();
