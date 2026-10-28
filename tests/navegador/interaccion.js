// Prueba de interacción del tablero. La ejecuta herramientas/probar_tablero.ps1 dentro de suficiencia_cx.html.
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
  // Los números llevan espacios que no se parten (85 %): se comparan como espacios normales
  function texto(el) { return el ? el.textContent.replace(/\s+/g, ' ').trim() : ''; }
  function titulo() { return $('#titulo-indicador').textContent; }
  function pctTxt(v) { return Math.round(v * 100) + '%'; }

  setTimeout(function () {
    try {
      var ej = SuficienciaEjemplos();
      ok(titulo() === 'Churn mensual pospago', 'título inicial');
      ok(heroe() === '85%', 'héroe inicial → ' + heroe());
      ok($$('.tarjeta').length === 2, 'riel: un solo ejemplo + «Nuevo cálculo»');
      ok(!!$('.tarjeta[data-id] .etq-ejemplo') && !!$('.pildora--ejemplo'), 'el ejemplo lleva su marca en la tarjeta y en el resultado');
      ok($$('#leyenda > li').length === 5, 'leyenda con 5 dolores');
      ok($$('#aportes > li').length === 5, 'ranking con 5 iniciativas');
      ok($('#aportes > li .ap-nombre').textContent === 'Resolución en el primer contacto', 'primera del ranking igual al Excel');
      ok($$('#vigilar .vig').length === 4 && $('#vigilar-cuenta').textContent === '6 alertas', 'vigilar: 4 visibles de 6');

      // Las 4 tarjetas de KPI, con los puntos que mueve el indicador
      var kpis = $$('#kpis .kpi');
      var r0 = Motor.calcular(ej[0]);
      ok(kpis.length === 4, '4 tarjetas de KPI');
      ok(texto($('dt', kpis[0])) === 'Mueve el indicador' && texto($('.kpi-mov', kpis[0])) === '0,68 pp' && !!$('.kpi-mov .ic', kpis[0]),
        'KPI «Mueve el indicador» → ' + texto($('.kpi-mov', kpis[0])) + ' (motor: ' + r0.real.pp.toFixed(4) + ')');
      ok(texto($('.sub', kpis[0])) === 'el 85 % de la brecha de 0,80 pp', 'explica de dónde salen los puntos → ' + texto($('.sub', kpis[0])));
      ok($$('#kpis dt').map(function (d) { return d.textContent; }).join('|') === 'Mueve el indicador|Así quedarías en 2027|Te faltaría|Si todo se cumple', 'orden de los KPI');

      // Nombres nuevos: «% dolor» e «Impacto del dolor»
      var cab = $('.tabla-cab').textContent;
      ok(cab.indexOf('% dolor') >= 0 && cab.indexOf('Impacto del dolor') >= 0, 'encabezados «% dolor» e «Impacto del dolor»');
      var visibles = ['#escenario', '#taller', '#dlg-como', '.pie'].map(function (s) { return $(s).textContent; }).join(' ');
      ok(!/[Pp]revalencia|[Pp]arte del pastel/.test(visibles), 'ya no aparecen «prevalencia» ni «parte del pastel»');
      ok(/Impacto del dolor/.test($('#leyenda').textContent), 'la leyenda del pastel habla de impacto del dolor');

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
      ok(titulo() === 'Churn prueba', 'el título sigue al nombre');
      ok($('.tarjeta[aria-pressed="true"] .tarjeta-nombre').textContent === 'Churn prueba', 'la tarjeta del riel sigue al nombre');
      ok(!$('.pildora--ejemplo') && !$('.tarjeta[aria-pressed="true"] .etq-ejemplo'), 'al ponerle tu nombre deja de ser ejemplo');

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
      ok(/6 de 15/.test($('[data-accion="agregar-ini"]').textContent), 'el botón cuenta las iniciativas → ' + $('[data-accion="agregar-ini"]').textContent);

      var det = $('.i-detalle', fila('ini-fila', 'i5'));
      $('[data-accion="detalle-ini"][aria-controls]', fila('ini-fila', 'i5')).click();
      ok(!det.hidden, 'abrir responsable, fecha y evidencia');

      // Papelera de cada tarjeta: duplicar, borrar la copia desde el riel y deshacer
      $('[data-accion="duplicar"]').click();
      ok($$('.tarjeta[data-id]').length === 2 && $$('.tarjeta-borrar').length === 2, 'cada tarjeta tiene su papelera');
      var copiaId = $('.tarjeta[aria-pressed="true"]').dataset.id;
      $$('.tarjeta-borrar').filter(function (b) { return b.dataset.id !== copiaId; })[0].click();
      ok($$('.tarjeta[data-id]').length === 1 && $('.tarjeta[data-id]').dataset.id === copiaId && titulo() === 'Churn prueba (copia)', 'la papelera borra ese indicador y deja ver el otro');
      ok(/Eliminaste «Churn prueba»/.test($('#aviso').textContent), 'avisa qué se borró → ' + $('#aviso').textContent.replace('Deshacer', ''));
      $('#aviso [data-accion="deshacer"]').click();
      ok($$('.tarjeta[data-id]').length === 2, 'deshacer devuelve el indicador borrado');
      $$('.tarjeta-borrar').filter(function (b) { return b.dataset.id === copiaId; })[0].click();
      ok($$('.tarjeta[data-id]').length === 1 && titulo() === 'Churn prueba', 'borrar el indicador que estás viendo pasa al siguiente');

      // Nuevo cálculo (botón de la barra): borra lo que había y empieza desde cero; se puede deshacer
      $('.barra [data-accion="nuevo-indicador"]').click();
      ok(titulo() === 'Indicador sin nombre', 'nuevo cálculo');
      ok($('.pastel-valor').textContent === '—', 'el cálculo nuevo empieza sin datos');
      ok($$('.tarjeta').length === 2, 'riel: solo el cálculo nuevo + «Nuevo cálculo» (se borró lo anterior)');
      ok(/se borró el indicador que había/.test($('#aviso').textContent) && !!$('#aviso [data-accion="deshacer"]'), 'avisa que se borró lo anterior y ofrece deshacer');
      $('#aviso [data-accion="deshacer"]').click();
      ok($$('.tarjeta[data-id]').length === 1 && titulo() === 'Churn prueba', 'deshacer recupera lo que había');
      $('.tarjeta--nueva').click();
      ok(titulo() === 'Indicador sin nombre' && $$('.tarjeta[data-id]').length === 1, 'la tarjeta «Nuevo cálculo» del riel hace lo mismo');
      ok($$('.paso-chip.hecho').length === 0, 'cálculo nuevo sin pasos listos');
      ok($$('.dolor').length === 0 && $$('.ini').length === 0, 'sin dolores ni iniciativas: todo se define desde cero');
      ok($$('#vigilar [data-tipo="ejemplo"]').length === 1, 'ofrece ver el ejemplo lleno');

      var ins = ej[1];
      escribir($('#paso-1 input[data-campo="nombre"]'), ins.nombre);
      $('[data-accion="sentido"][data-valor="sube"]').click();
      escribir($('#paso-1 input[data-campo="ejecutado"]'), '82');
      escribir($('#paso-1 input[data-campo="meta"]'), '92');
      escribir($('#paso-2 input[data-campo="totalCasos"]'), '9.000');
      ins.dolores.forEach(function (d) {
        $('[data-accion="agregar-dolor"]').click();
        var fd = $$('.dolor').pop();
        escribir($('textarea[data-campo="nombre"]', fd), d.nombre);
        escribir($('input[data-campo="casos"]', fd), String(d.casos));
      });
      var idsDolor = $$('.dolor').map(function (f) { return f.getAttribute('data-dolor-fila'); });
      ins.iniciativas.forEach(function (x) {
        $('[data-accion="agregar-ini"]').click();
        var fi = $$('.ini').pop();
        escribir($('textarea[data-campo="nombre"]', fi), x.nombre);
        Object.keys(x.efectividad).forEach(function (did) {
          var c = $$('input[data-campo="efectividad"]', fi).filter(function (i) { return i.dataset.dolor === idsDolor[Number(did.slice(1)) - 1]; })[0];
          escribir(c, String(Math.round(x.efectividad[did] * 100)));
        });
        escribir($('.ejec-ctl input.entrada', fi), String(Math.round(x.ejecucion * 100)));
        var sel = $('select[data-campo="estado"]', fi);
        sel.value = x.estado;
        disparar(sel, 'change');
      });
      var r1 = Motor.calcular(ins);
      ok(heroe() === pctTxt(r1.real.suficiencia), 'cálculo hecho desde cero igual al motor → ' + heroe() + ' (motor: ' + pctTxt(r1.real.suficiencia) + ')');
      ok($$('.paso-chip.hecho').length === 2, 'pasos 1 y 2 listos (al 3 le falta la evidencia)');
      ok(texto($('.kpi-mov')) === r1.real.pp.toFixed(2).replace('.', ',') + ' pp', 'mueve el indicador → ' + texto($('.kpi-mov')));

      $('[data-accion="sentido"][data-valor="baja"]').click();
      ok($('.pastel-valor').textContent === '✓', '"baja es mejor" deja la meta cumplida → ' + $('.pastel-valor').textContent);
      ok($('#kpis dt').textContent === 'Mueve el indicador' && $$('#kpis .kpi').length === 4, 'con la meta cumplida siguen las 4 tarjetas');
      $('[data-accion="sentido"][data-valor="sube"]').click();
      ok(heroe() === pctTxt(r1.real.suficiencia), 'volver a "sube es mejor" → ' + heroe());

      // Hasta 15 dolores y 15 iniciativas
      while ($$('.dolor').length < 15) $('[data-accion="agregar-dolor"]').click();
      $$('.dolor').slice(4).forEach(function (fd, k) {
        escribir($('textarea[data-campo="nombre"]', fd), 'Dolor extra ' + (k + 1));
        escribir($('input[data-campo="casos"]', fd), '100');
      });
      var btnD = $('[data-accion="agregar-dolor"]');
      ok(btnD.disabled && /máximo de 15 dolores/.test(btnD.textContent) && /15 de 15/.test(btnD.textContent), 'máximo 15 dolores → ' + btnD.textContent);
      while ($$('.ini').length < 15) $('[data-accion="agregar-ini"]').click();
      var btnI = $('[data-accion="agregar-ini"]');
      ok(btnI.disabled && /15 de 15/.test(btnI.textContent), 'máximo 15 iniciativas → ' + btnI.textContent);
      ok($('.matriz').style.getPropertyValue('--min-ancho') === (840 + 68 * 15) + 'px' && !!$('.matriz-scroll .matriz'), 'la tabla de 15 dolores se desliza de lado');
      ok($$('#pastel-svg .porcion').length === 15, 'el pastel muestra 15 porciones');
      ok($$('#leyenda > li[data-dolor]').length === 6 && $('#leyenda [data-accion="mas-leyenda"]').textContent === 'Ver los 15 dolores', 'la leyenda muestra 6 y ofrece ver los 15');
      $('#leyenda [data-accion="mas-leyenda"]').click();
      ok($$('#leyenda > li[data-dolor]').length === 15, 'al pedirlo muestra los 15 dolores');

      // «Ver el ejemplo lleno»: el ejemplo se renombró (ya es un cálculo propio), así que se vuelve a agregar
      $('[data-accion="ver-ejemplo"]').click();
      ok(titulo() === 'Churn mensual pospago' && !!$('.pildora--ejemplo') && heroe() === '85%', 'Ver el ejemplo lleno lo trae de vuelta');
      ok($$('.tarjeta[data-id]').length === 2, 'sin tocar tus cálculos');

      $('#ajustes').open = true;
      var verde = $('[data-param="verde"]');
      escribir(verde, '50');
      ok(verde.getAttribute('aria-invalid') === 'true', 'verde por debajo de amarillo se rechaza');
      escribir(verde, '100');

      $('.barra [data-accion="como-funciona"]').click();
      ok($('#dlg-como').open && $$('#dlg-como [data-demo] .porcion').length > 0, 'diálogo con los 4 pasteles de ejemplo');
      $('#dlg-como [data-accion="cerrar-dialogo"]').click();

      setTimeout(function () {
        var guardado = localStorage.getItem('suficiencia.datos.v1');
        ok(!!guardado && guardado.indexOf('Instalaciones en fecha (FTTH)') >= 0 && guardado.indexOf('Churn prueba') < 0 && JSON.parse(guardado).version === 2, 'se guarda en el navegador');
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
