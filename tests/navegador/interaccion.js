// Prueba de interacción del tablero. La ejecuta herramientas/probar_tablero.ps1 dentro de index.html.
// Cada resultado se compara con el motor de cálculo (Motor) para que pantalla y cálculo coincidan.
(function () {
  var errores = [];
  window.addEventListener('error', function (e) { errores.push(e.message + ' @' + e.lineno); });
  function $(s, r) { return (r || document).querySelector(s); }
  function $$(s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); }
  var out = [];
  function ok(c, m) { out.push((c ? 'OK    ' : 'FALLA ') + m); }
  function disparar(el, tipo) { el.dispatchEvent(new Event(tipo, { bubbles: true })); }
  function escribir(el, v) { el.focus(); el.value = v; disparar(el, 'input'); el.blur(); }
  function fila(tipo, id) { return $$('[data-' + tipo + ']').filter(function (e) { return e.getAttribute('data-' + tipo) === id; })[0]; }
  function heroe() { return $('.pastel-valor').textContent.replace(/\s/g, ''); }
  function pctTxt(v) { return Math.round(v * 100) + '%'; }

  setTimeout(function () {
    try {
      var ej = SuficienciaEjemplos();
      ok($('#titulo-indicador').textContent === 'Churn mensual pospago', 'título inicial');
      ok(heroe() === '85%', 'héroe inicial → ' + heroe());
      ok($$('.tarjeta').length === 3, 'riel: 2 tarjetas + nueva');
      ok($$('#leyenda > li').length === 5, 'leyenda con 5 dolores');
      ok($$('#aportes > li').length === 5, 'ranking con 5 iniciativas');
      ok($('#aportes > li .ap-nombre').textContent === 'Resolución en el primer contacto', 'primera del ranking igual al Excel');
      ok($$('#vigilar .vig').length === 4 && $('#vigilar-cuenta').textContent === '6 alertas', 'vigilar: 4 visibles de 6');

      var rango = $('input.rango', fila('ini-fila', 'i1'));
      rango.value = '100'; disparar(rango, 'input');
      ej[0].iniciativas[0].ejecucion = 1;
      var esperado = pctTxt(Motor.calcular(ej[0]).real.suficiencia);
      ok(heroe() === esperado, 'deslizador I1 a 100 % → ' + heroe() + ' (motor: ' + esperado + ')');
      var ejecTxt = $('.ejec-ctl input.entrada', fila('ini-fila', 'i1')).value;
      ok(ejecTxt === '100', 'campo de ejecución sincronizado → ' + ejecTxt);

      var celda = $$('input[data-campo="efectividad"]').filter(function (i) { return i.dataset.ini === 'i5' && i.dataset.dolor === 'd2'; })[0];
      escribir(celda, '40');
      ej[0].iniciativas[4].efectividad.d2 = 0.4;
      esperado = pctTxt(Motor.calcular(ej[0]).real.suficiencia);
      ok(heroe() === esperado, 'celda I5 × D2 = 40 → ' + heroe() + ' (motor: ' + esperado + ')');
      ok(celda.parentElement.style.getPropertyValue('--e') === '0.4', 'color de calor de la celda → ' + celda.parentElement.style.getPropertyValue('--e'));

      escribir(celda, '140');
      ok(celda.getAttribute('aria-invalid') === 'true', '140 % se marca como inválido');
      escribir(celda, '40');
      ok(!celda.hasAttribute('aria-invalid'), 'al corregir se quita la marca');

      var nom = $('#paso-1 input[data-campo="nombre"]');
      escribir(nom, 'Churn prueba');
      ok($('#titulo-indicador').textContent === 'Churn prueba', 'el título sigue al nombre');
      ok($('.tarjeta[aria-pressed="true"] .tarjeta-nombre').textContent === 'Churn prueba', 'la tarjeta del riel sigue al nombre');

      var hoy = $('#paso-1 input[data-campo="ejecutado"]');
      escribir(hoy, '2,6');
      ok($$('.pildora b')[0].textContent.indexOf('2,60') === 0, 'hoy con coma decimal → ' + $$('.pildora b')[0].textContent);
      escribir(hoy, '2.8');
      ok(hoy.value === '2,8', 'hoy con punto se muestra 2,8 → ' + hoy.value);

      var nomIni = $('textarea[data-campo="nombre"]', fila('ini-fila', 'i2'));
      escribir(nomIni, 'Ofertas de\nretención');
      ok($$('#aportes .ap-nombre').some(function (e) { return e.textContent === 'Ofertas de retención'; }), 'salto de línea pegado en un nombre se vuelve espacio');

      var nDol = $$('.dolor').length;
      $('[data-accion="agregar-dolor"]').click();
      ok($$('.dolor').length === nDol + 1, 'agregar dolor');
      $('[data-accion="borrar-dolor"]', $$('.dolor').pop()).click();
      ok($$('.dolor').length === nDol, 'borrar el dolor nuevo');
      $('[data-accion="borrar-dolor"]', $$('.dolor')[0]).click();
      ok($$('.dolor').length === nDol - 1, 'borrar el dolor 1');
      $('#aviso [data-accion="deshacer"]').click();
      ok($$('.dolor').length === nDol && heroe() === esperado, 'deshacer devuelve el dolor y el resultado → ' + heroe());

      var nIni = $$('.ini').length;
      $('[data-accion="agregar-ini"]').click();
      ok($$('.ini').length === nIni + 1, 'agregar iniciativa');

      var det = $('.i-detalle', fila('ini-fila', 'i5'));
      $('[data-accion="detalle-ini"][aria-controls]', fila('ini-fila', 'i5')).click();
      ok(!det.hidden, 'abrir responsable, fecha y evidencia');

      $$('.tarjeta').filter(function (t) { return t.dataset.id === 'ej-instalaciones'; })[0].click();
      ok($('#titulo-indicador').textContent === 'Instalaciones en fecha (FTTH)', 'seleccionar Instalaciones');
      ok(heroe() === '122%', 'héroe Instalaciones → ' + heroe());

      $('[data-accion="sentido"][data-valor="baja"]').click();
      ok($('.pastel-valor').textContent === '✓', '"baja es mejor" deja la meta cumplida → ' + $('.pastel-valor').textContent);
      $('[data-accion="sentido"][data-valor="sube"]').click();
      ok(heroe() === '122%', 'volver a "sube es mejor" → ' + heroe());

      $('[data-accion="nuevo-indicador"]').click();
      ok($('#titulo-indicador').textContent === 'Indicador sin nombre', 'nuevo indicador');
      ok($('.pastel-valor').textContent === '—', 'nuevo indicador sin datos');
      ok($$('.tarjeta').length === 4, 'riel: 3 tarjetas + nueva');
      ok($$('.paso-chip.hecho').length === 0, 'nuevo indicador sin pasos listos');

      $('#ajustes').open = true;
      var verde = $('[data-param="verde"]');
      escribir(verde, '50');
      ok(verde.getAttribute('aria-invalid') === 'true', 'verde por debajo de amarillo se rechaza');
      escribir(verde, '100');

      $('[data-accion="como-funciona"]').click();
      ok($('#dlg-como').open && $$('#dlg-como [data-demo] .porcion, #dlg-como [data-demo] .porcion-base').length > 0, 'diálogo con los 4 pasteles de ejemplo');
      $('#dlg-como [data-accion="cerrar-dialogo"]').click();

      setTimeout(function () {
        var guardado = localStorage.getItem('suficiencia.datos.v1');
        ok(!!guardado && guardado.indexOf('Churn prueba') >= 0, 'se guarda en el navegador');
        terminar();
      }, 500);
    } catch (e) { out.push('EXCEPCION ' + e.message + ' | ' + e.stack); terminar(); }
  }, 300);

  function terminar() {
    out.push('errores JS: ' + JSON.stringify(errores));
    var pre = document.createElement('pre');
    pre.id = 'resultado-prueba';
    pre.textContent = out.join('\n');
    document.body.appendChild(pre);
  }
})();
