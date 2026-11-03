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
  // Como en el navegador: «input» con cada tecla y «change» al salir del campo
  function escribir(el, v) { el.focus(); el.value = v; disparar(el, 'input'); disparar(el, 'change'); el.blur(); }
  function fila(tipo, id) { return $$('[data-' + tipo + ']').filter(function (e) { return e.getAttribute('data-' + tipo) === id; })[0]; }
  function heroe() { return $('.pastel-valor').textContent.replace(/\s/g, ''); }
  // Los números llevan espacios que no se parten (85 %): se comparan como espacios normales
  function dt(kpi) { return texto($('dt', kpi)).replace(/\s*\?$/, ''); }
  function texto(el) { return el ? el.textContent.replace(/\s+/g, ' ').trim() : ''; }
  function titulo() { return $('#titulo-indicador').textContent; }
  function pctTxt(v) { return Math.round(v * 100) + '%'; }

  setTimeout(function () {
    try {
      var ej = SuficienciaEjemplos();
      ok(titulo() === 'Churn mensual pospago', 'título inicial');
      ok(heroe() === '85%', 'héroe inicial → ' + heroe());
      ok($$('.tarjeta').length === 2, 'riel: un solo ejemplo + «Agregar indicador»');
      ok(!!$('.tarjeta[data-id] .etq-ejemplo') && !!$('.pildora--ejemplo'), 'el ejemplo lleva su marca en la tarjeta y en el resultado');
      ok($$('#leyenda > li').length === 5, 'leyenda con 5 dolores');
      ok($$('#aportes > li').length === 5, 'ranking con 5 iniciativas');
      ok($('#aportes > li .ap-nombre').textContent === 'Resolución en el primer contacto', 'primera del ranking igual al Excel');
      ok($$('#vigilar .vig').length === 4 && $('#vigilar-cuenta').textContent === '4 alertas', 'vigilar: 4 alertas → ' + $('#vigilar-cuenta').textContent);
      ok(!/En riesgo|retrasada/.test($('#escenario').textContent), 'el estado ya no cuenta: sin «En riesgo» ni «está retrasada»');

      // Las 4 tarjetas de KPI, con los puntos que mueve el indicador
      var kpis = $$('#kpis .kpi');
      var r0 = Motor.calcular(ej[0]);
      ok(kpis.length === 4, '4 tarjetas de KPI');
      ok(dt(kpis[0]) === 'Mueve el indicador' && texto($('.kpi-mov', kpis[0])) === '0,68 pp' && !!$('.kpi-mov .ic', kpis[0]),
        'KPI «Mueve el indicador» → ' + texto($('.kpi-mov', kpis[0])) + ' (motor: ' + r0.real.pp.toFixed(4) + ')');
      ok(texto($('.sub', kpis[0])) === 'el 85 % de la brecha de 0,80 pp', 'explica de dónde salen los puntos → ' + texto($('.sub', kpis[0])));
      ok($$('#kpis .kpi').map(dt).join('|') === 'Mueve el indicador|Así quedarías en 2027|Te faltaría|Si todo se cumple', 'orden de los KPI');
      // «¿Qué significa este número?»: cada resultado tiene su ayuda
      ok($$('#kpis .ayuda').length === 4 && $$('#escenario .ayuda').length >= 6, 'cada resultado tiene su «?» de ayuda');
      $('#kpis .ayuda').click();
      ok(!$('#tip').hidden && /puntos se mueve el indicador/.test($('#tip').textContent), 'la ayuda explica el número con un ejemplo → ' + $('#tip').textContent);
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));

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

      // + Agregar indicador (tarjeta del riel): suma uno vacío al mismo tablero, sin borrar nada
      ok($('.tarjeta--nueva b').textContent === '+ Agregar indicador', 'la tarjeta del riel se llama «+ Agregar indicador»');
      $('.tarjeta--nueva').click();
      ok(titulo() === 'Indicador sin nombre' && $$('.tarjeta[data-id]').length === 2, 'agregar indicador suma uno vacío y conserva el anterior');
      ok($('.pastel-valor').textContent === '—', 'el indicador nuevo empieza sin datos');
      ok($$('.paso-chip.hecho').length === 0, 'indicador nuevo sin pasos listos');
      ok($$('.dolor').length === 0 && $$('.ini').length === 0, 'sin dolores ni iniciativas: todo se define desde cero');
      ok($$('#vigilar [data-tipo="ejemplo"]').length === 1, 'ofrece ver el ejemplo lleno');
      ok(!$('#aviso [data-accion="deshacer"]') && /Agregaste un indicador nuevo/.test($('#aviso').textContent), 'avisa que se agregó (no borró nada)');
      ok(!/Estado/.test($('.matriz-cab').textContent) && !$('select[data-campo="estado"]'), 'el paso 3 ya no tiene la columna «Estado»');

      // Nuevo cálculo (botón rojo): pregunta si guardas el tablero actual
      $('.barra [data-accion="nuevo-calculo"]').click();
      var dlg = $('#dlg-pregunta');
      ok(dlg.open && $('#preg-titulo').textContent === '¿Guardar los cambios del tablero actual?', 'nuevo cálculo pregunta si guardas');
      ok($$('#preg-botones button').map(function (b) { return b.textContent; }).join('|') === 'Guardar|No guardar|Cancelar' && !$('#preg-campo').hidden, 'botones Guardar / No guardar / Cancelar y campo de nombre');
      $$('#preg-botones button')[2].click();
    } catch (e) { out.push('EXCEPCION ' + e.message + ' | ' + e.stack); terminar(); return; }
    setTimeout(paso2, 60);
  }, 300);

  function paso2() {
    try {
      ok(!$('#dlg-pregunta').open && $$('.tarjeta[data-id]').length === 2, 'Cancelar no cambia nada');
      $('.barra [data-accion="nuevo-calculo"]').click();
      $('#preg-nombre').value = 'Mi tablero de prueba';
      $$('#preg-botones button')[0].click();
    } catch (e) { out.push('EXCEPCION ' + e.message + ' | ' + e.stack); terminar(); return; }
    setTimeout(paso3, 60);
  }

  function marcar(nombre) { $$('#paso-1 input[data-campo="segmento"]').filter(function (x) { return x.value === nombre; })[0].click(); }
  function peso(nombre) { return $$('#paso-1 input[data-campo="peso"]').filter(function (x) { return x.dataset.segmento === nombre; })[0]; }
  function filasSeg() { return $$('#tabla-seg tbody tr').map(function (tr) { return texto(tr).replace(/[  ]/g, ''); }); }
  function bloques() { return $$('#paso-2 [data-grupo-bloque]').map(function (b) { return b.dataset.grupoBloque; }); }
  function bloque(nombre) { return $$('#paso-2 [data-grupo-bloque]').filter(function (b) { return b.dataset.grupoBloque === nombre; })[0]; }
  function guardado() { return JSON.parse(localStorage.getItem('suficiencia.espacio.v3') || 'null'); }
  function dosDec(v) { return v.toFixed(2).replace('.', ','); }

  function paso3() {
    try {
      var ej = SuficienciaEjemplos();
      ok(titulo() === 'Indicador sin nombre' && $$('.tarjeta[data-id]').length === 1, 'Guardar: abre un tablero nuevo con un indicador vacío');
      ok(/Guardaste «Mi tablero de prueba» en Mis cálculos/.test($('#aviso').textContent), 'avisa que se guardó → ' + texto($('#aviso')));
      var g = guardado();
      ok(g && g.calculos.length === 1 && g.calculos[0].nombre === 'Mi tablero de prueba' && g.calculos[0].datos.indicadores.length === 2, 'el cálculo guardado tiene sus 2 indicadores');

      var ins = ej[1];
      escribir($('#paso-1 input[data-campo="nombre"]'), ins.nombre);
      $('[data-accion="sentido"][data-valor="sube"]').click();
      escribir($('#paso-1 input[data-campo="ejecutado"]'), '82');
      escribir($('#paso-1 input[data-campo="meta"]'), '92');
      ok(/el segmento/.test($('#pasos').textContent) && /Elige UMM, UMC o Transversal/.test($('#paso-1').textContent), 'sin segmento el paso 1 no está listo y lo pide');
      $('[data-accion="tipo-segmento"][data-valor="UMM"]').click();
      ok($$('#paso-1 input[data-campo="segmento"]').map(function (x) { return x.value; }).join('|') === 'Hogar|Pospago|Prepago', 'UMM: Hogar, Pospago y Prepago');
      ok(/Marca al menos un segmento/.test($('#paso-1').textContent), 'mensaje simple si falta el segmento');
      marcar('Hogar');
      ok(!$('#paso-1 input[data-campo="peso"]'), 'con un solo segmento no se pide la parte del problema');
      ok($$('#paso-2 [data-grupo-bloque]').length === 1 && !$('#paso-2 .seg-bloque-titulo') && !$('#paso-2 select[data-campo="segmentoDolor"]'), 'con un solo segmento el paso 2 se ve como siempre (un bloque sin título)');
      escribir($('#paso-2 input[data-campo="casosGrupo"]'), '9.000');
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
      });
      // Sin la columna Estado, las iniciativas nuevas quedan «En plan»
      ins.iniciativas.forEach(function (x) { x.estado = 'En plan'; });
      var r1 = Motor.calcular(ins);
      ok(heroe() === pctTxt(r1.real.suficiencia), 'cálculo hecho desde cero igual al motor → ' + heroe() + ' (motor: ' + pctTxt(r1.real.suficiencia) + ')');
      ok($$('.paso-chip.hecho').length === 2, 'pasos 1 y 2 listos (al 3 le falta la evidencia)');
      ok(texto($('.kpi-mov')) === dosDec(r1.real.pp) + ' pp', 'mueve el indicador → ' + texto($('.kpi-mov')));
      ok(/UMM\s*Hogar/.test(texto($('#meta-linea'))), 'el resumen muestra el segmento → ' + texto($('#meta-linea')));
      ok($('#panel-segmentos').hidden, 'con un solo segmento no hay tabla por segmento');

      // Transversal con 2 segmentos: dolores e iniciativas por segmento
      $('[data-accion="tipo-segmento"][data-valor="Transversal"]').click();
      ok($$('#paso-1 input[data-campo="segmento"]').length === 5 && $$('#paso-1 input[data-campo="segmento"]:checked').length === 1, 'Transversal muestra los 5 segmentos y conserva Hogar');
      marcar('Pyme Fijo');
      ok(document.activeElement && document.activeElement.value === 'Pyme Fijo', 'el foco vuelve a la casilla que marcaste');
      ok(bloques().join('|') === 'Hogar|Pyme Fijo', 'paso 2: un bloque por segmento → ' + bloques().join('|'));
      ok($$('.dolor', bloque('Hogar')).length === 4 && !$('.dolor', bloque('Pyme Fijo')), 'los dolores que ya tenías quedan en Hogar');
      ok(heroe() === pctTxt(r1.real.suficiencia), 'sin casos en Pyme Fijo, Hogar sigue siendo todo el problema');
      ok(/Faltan dolores de Pyme Fijo/.test(texto($('#pasos'))), 'el paso 2 dice qué falta → ' + texto($('#pasos')));
      ok(/Agrega los casos analizados y los dolores de «Pyme Fijo»/.test(texto($('#vigilar'))), 'lo que debes vigilar pide los dolores de Pyme Fijo');
      ok($$('#paso-3 .aplica').length === 4 && $$('.aplica-seg', $('#paso-3 .ini')).map(texto).join('|') === 'Hogar|Pyme Fijo', 'cada iniciativa tiene «Aplica a» con los 2 segmentos');

      // Casos de Pyme Fijo: 1.000 de 10.000 → 10 % del problema, aún sin dolores
      escribir($('input[data-campo="casosGrupo"]', bloque('Pyme Fijo')), '1.000');
      ok(heroe() === pctTxt(0.9 * r1.real.suficiencia), 'Pyme Fijo sin dolores: su 10 % del problema queda sin resolver → ' + heroe());
      ok(/Faltan los dolores de «Pyme Fijo»: es el 10 %/.test(texto($('#vigilar'))), 'alerta: faltan los dolores de Pyme Fijo (10 %)');
      ok(/Parte del problema 90,0 %/.test(texto($('.seg-bloque-parte', bloque('Hogar')))), 'el bloque de Hogar dice su parte → ' + texto($('.seg-bloque-parte', bloque('Hogar'))));
      $('[data-accion="agregar-dolor"]', bloque('Pyme Fijo')).click();
      var fp = $('.dolor', bloque('Pyme Fijo'));
      escribir($('textarea[data-campo="nombre"]', fp), 'Demora en la visita');
      escribir($('input[data-campo="casos"]', fp), '500');
      var idPyme = fp.getAttribute('data-dolor-fila');
      ok($$('#paso-3 .cab-seg').map(texto).join('|') === 'Hogar|Hogar|Hogar|Hogar|Pyme Fijo', 'paso 3: cada columna de dolor dice su segmento → ' + $$('#paso-3 .cab-seg').map(texto).join('|'));
      // La iniciativa 1 resuelve 50 % del dolor de Pyme Fijo; luego deja de aplicar a Hogar
      var i1 = $('.ini');
      escribir($$('input[data-campo="efectividad"]', i1).filter(function (c) { return c.dataset.dolor === idPyme; })[0], '50');
      var esperado = JSON.parse(JSON.stringify(ins));
      esperado.tipoSegmento = 'Transversal';
      esperado.segmentos = [{ nombre: 'Hogar', peso: null, casos: 9000 }, { nombre: 'Pyme Fijo', peso: null, casos: 1000 }];
      esperado.dolores.forEach(function (d, k) { d.segmento = 'Hogar'; d.id = idsDolor[k]; });
      esperado.dolores.push({ id: idPyme, segmento: 'Pyme Fijo', nombre: 'Demora en la visita', casos: 500 });
      esperado.iniciativas.forEach(function (x) { var e = {}; Object.keys(x.efectividad).forEach(function (k) { e[idsDolor[Number(k.slice(1)) - 1]] = x.efectividad[k]; }); x.efectividad = e; x.segmentos = []; });
      esperado.iniciativas[0].efectividad[idPyme] = 0.5;
      var r2 = Motor.calcular(Motor.normalizarIndicador(esperado));
      ok(heroe() === pctTxt(r2.real.suficiencia), 'cálculo con 2 segmentos igual al motor → ' + heroe() + ' (motor: ' + pctTxt(r2.real.suficiencia) + ')');
      ok(!$('#panel-segmentos').hidden && filasSeg().length === 3 && filasSeg()[0].indexOf('Hogar90,0%auto') === 0 && filasSeg()[1].indexOf('PymeFijo10,0%auto' + dosDec(r2.segmentos.filas[1].pp) + 'pp') === 0,
        'tabla por segmento: parte, mueve y de la brecha → ' + filasSeg().join(' | '));
      $$('.aplica-seg', i1)[0].click();
      ok($$('.aplica-seg', $('.ini'))[0].getAttribute('aria-pressed') === 'false' && $$('.celda-caja--no', $('.ini')).length === 4, 'si no aplica a Hogar, sus 4 celdas dicen «No aplica»');
      esperado.iniciativas[0].segmentos = ['Pyme Fijo'];
      var r3 = Motor.calcular(Motor.normalizarIndicador(esperado));
      ok(r3.real.pp < r2.real.pp && heroe() === pctTxt(r3.real.suficiencia), 'y ya no resuelve dolores de Hogar → ' + heroe());
      $$('.aplica-seg', $('.ini'))[1].click();
      ok(/al menos a un segmento/.test(texto($('#aviso'))) && $$('.aplica-seg', $('.ini'))[1].getAttribute('aria-pressed') === 'true', 'no deja una iniciativa sin segmentos');
      $$('.aplica-seg', $('.ini'))[0].click();
      ok(heroe() === pctTxt(r2.real.suficiencia), 'al volver a aplicar a Hogar vuelve el resultado');
      esperado.iniciativas[0].segmentos = [];

      // Parte del problema escrita a mano (opcional)
      ok(peso('Hogar').placeholder === '90' && peso('Pyme Fijo').placeholder === '10', 'la parte automática se ve en el campo → ' + peso('Hogar').placeholder);
      escribir(peso('Hogar'), '60');
      ok(/Escrita a mano/.test(texto($$('#paso-1 .seg-aporte')[0])) && /Auto \(de sus casos\): 40,0 %/.test(texto($$('#paso-1 .seg-aporte')[1])), 'Hogar 60 % a mano y Pyme Fijo toma el 40 % que falta');
      var pf = peso('Pyme Fijo');
      pf.focus(); pf.value = '5'; disparar(pf, 'input');
      ok(!pf.hasAttribute('aria-invalid') && filasSeg()[1].indexOf('auto') > 0, 'mientras escribes no se guarda un número a medio escribir');
      pf.value = '50'; disparar(pf, 'input');
      ok(pf.getAttribute('aria-invalid') === 'true' && /más de 100 %/.test(texto($('#paso-1 .pista--alerta[role="alert"]'))), 'si la suma pasa de 100 % se marca al escribir, con mensaje');
      disparar(pf, 'change'); pf.blur();
      ok(pf.value === '' && !pf.hasAttribute('aria-invalid') && /No se guardó/.test(texto($('#paso-1 .pista--alerta[role="alert"]'))) && filasSeg()[1].indexOf('auto') > 0,
        'al confirmar se rechaza: el campo vuelve al valor anterior y lo dice');
      ok($('[data-accion="repartir-pesos"]').textContent === 'Repartir en partes iguales', 'hay un botón para repartir en partes iguales');
      $('[data-accion="repartir-pesos"]').click();
      ok(peso('Hogar').value === '50' && peso('Pyme Fijo').value === '50' && filasSeg()[1].indexOf('PymeFijo50,0%manual') === 0 && !!$('#aviso [data-accion="deshacer"]'),
        'repartir en partes iguales: 50 % y 50 %, y se puede deshacer → ' + filasSeg().join(' | '));
      esperado.segmentos.forEach(function (sg) { sg.peso = 0.5; });
      var r4 = Motor.calcular(Motor.normalizarIndicador(esperado));
      ok(heroe() === pctTxt(r4.real.suficiencia), 'la parte del problema sí cambia la suficiencia, igual que el motor → ' + heroe());

      // Copiar dolores de Hogar a Pyme Fijo (sin casos) y deshacer
      var sel = $('select[data-campo="copiarDolores"]', bloque('Pyme Fijo'));
      ok(!!sel && !!$('select[data-campo="copiarDolores"]', bloque('Hogar')), 'cada bloque ofrece copiar los dolores del otro');
      sel.value = 'Hogar'; disparar(sel, 'change');
      var copiados = $$('.dolor', bloque('Pyme Fijo'));
      ok(copiados.length === 5 && $('textarea[data-campo="nombre"]', copiados[1]).value === ins.dolores[0].nombre && $('input[data-campo="casos"]', copiados[1]).value === '',
        'copia los 4 nombres de Hogar a Pyme Fijo, sin casos');
      ok(/Copiaste 4 dolores de Hogar a Pyme Fijo/.test(texto($('#aviso'))), 'avisa y ofrece deshacer → ' + texto($('#aviso')));
      $('#aviso [data-accion="deshacer"]').click();
      ok($$('.dolor', bloque('Pyme Fijo')).length === 1, 'deshacer quita los copiados');
      // Mover un dolor a otro segmento y deshacer
      var mover = $('select[data-campo="segmentoDolor"]', bloque('Pyme Fijo'));
      mover.value = 'Hogar'; disparar(mover, 'change');
      ok($$('.dolor', bloque('Hogar')).length === 5 && !$('.dolor', bloque('Pyme Fijo')), 'el dolor pasa a Hogar');
      $('#aviso [data-accion="deshacer"]').click();
      // Quitar Pyme Fijo borra sus dolores (se puede deshacer)
      marcar('Pyme Fijo');
      ok(bloques().join('|') === 'Hogar' && /Quitaste «Pyme Fijo». Se borró su dolor/.test(texto($('#aviso'))), 'quitar Pyme Fijo borra su dolor y avisa → ' + texto($('#aviso')));
      $('#aviso [data-accion="deshacer"]').click();
      ok(bloques().join('|') === 'Hogar|Pyme Fijo' && $$('.dolor', bloque('Pyme Fijo')).length === 1 && heroe() === pctTxt(r4.real.suficiencia), 'deshacer devuelve Pyme Fijo con su dolor → ' + bloques().join('|') + ' ' + heroe());

      $('[data-accion="sentido"][data-valor="baja"]').click();
      ok($('.pastel-valor').textContent === '✓', '"baja es mejor" deja la meta cumplida → ' + $('.pastel-valor').textContent);
      ok(dt($('#kpis .kpi')) === 'Mueve el indicador' && $$('#kpis .kpi').length === 4, 'con la meta cumplida siguen las 4 tarjetas');
      $('[data-accion="sentido"][data-valor="sube"]').click();
      ok(heroe() === pctTxt(r4.real.suficiencia), 'volver a "sube es mejor" → ' + heroe());

      // Hasta 15 dolores por segmento y 15 iniciativas por segmento (30 con 2 segmentos)
      while ($$('.dolor', bloque('Hogar')).length < 15) $('[data-accion="agregar-dolor"]', bloque('Hogar')).click();
      $$('.dolor', bloque('Hogar')).slice(4).forEach(function (fd, k) {
        escribir($('textarea[data-campo="nombre"]', fd), 'Dolor extra ' + (k + 1));
        escribir($('input[data-campo="casos"]', fd), '100');
      });
      var btnD = $('[data-accion="agregar-dolor"]', bloque('Hogar'));
      ok(btnD.disabled && /máximo de 15 dolores en Hogar/.test(btnD.textContent) && /15 de 15/.test(btnD.textContent), 'máximo 15 dolores por segmento → ' + btnD.textContent);
      ok(!$('[data-accion="agregar-dolor"]', bloque('Pyme Fijo')).disabled, 'Pyme Fijo todavía puede agregar dolores');
      while ($$('.ini').length < 30) $('[data-accion="agregar-ini"]').click();
      var btnI = $('[data-accion="agregar-ini"]');
      ok(btnI.disabled && /30 de 30/.test(btnI.textContent), 'máximo 15 iniciativas por segmento → ' + btnI.textContent);
      ok($('.matriz').style.getPropertyValue('--min-ancho') === (706 + 68 * 16) + 'px' && !!$('.matriz-scroll .matriz'), 'la tabla de 16 dolores se desliza de lado');
      ok($$('#pastel-svg .porcion').length === 16, 'el pastel muestra 16 porciones');
      ok($$('#leyenda > li[data-dolor]').length === 6 && $('#leyenda [data-accion="mas-leyenda"]').textContent === 'Ver los 16 dolores', 'la leyenda muestra 6 y ofrece ver los 16');
      $('#leyenda [data-accion="mas-leyenda"]').click();
      ok($$('#leyenda > li[data-dolor]').length === 16 && /Pyme Fijo/.test(texto($$('#leyenda > li[data-dolor]').pop())), 'al pedirlo muestra los 16 dolores, con su segmento');

      // Mis cálculos: renombrar y abrir (pregunta antes porque hay cambios)
      $('[aria-controls="menu-calculos"]').click();
      var menu = $('#menu-calculos');
      ok(!menu.hidden && $$('.calc-item', menu).length === 1 && /Cálculo sin guardar/.test(menu.textContent) && /cambios sin guardar/.test(menu.textContent), 'Mis cálculos lista el guardado y dice que el actual no está guardado');
      var nom = $('.calc-nombre', menu);
      nom.value = 'Tablero renombrado';
      disparar(nom, 'change');
      ok(guardado().calculos[0].nombre === 'Tablero renombrado', 'el nombre se edita en la lista');
      ok(/\d{4}/.test($('.calc-item small', menu).textContent), 'muestra la fecha → ' + $('.calc-item small', menu).textContent);
      $('[data-accion="abrir-calculo"]', menu).click();
      ok($('#dlg-pregunta').open, 'abrir otro cálculo pregunta si guardas los cambios');
      $$('#preg-botones button')[1].click();
    } catch (e) { out.push('EXCEPCION ' + e.message + ' | ' + e.stack); terminar(); return; }
    setTimeout(paso4, 60);
  }

  function paso4() {
    try {
      ok($$('.tarjeta[data-id]').length === 2 && titulo() === 'Indicador sin nombre', 'No guardar: abre el cálculo guardado tal como quedó');
      ok(/Abriste «Tablero renombrado»/.test($('#aviso').textContent) && !!$('#aviso [data-accion="deshacer"]'), 'avisa y ofrece deshacer');
      $('#aviso [data-accion="deshacer"]').click();
      ok(titulo() === 'Instalaciones en fecha (FTTH)' && $$('.dolor').length === 16, 'deshacer devuelve el tablero que no guardaste');
      $('[aria-controls="menu-calculos"]').click();
      $('[data-accion="eliminar-calculo"]', $('#menu-calculos')).click();
      ok($('#dlg-pregunta').open && /¿Eliminar «Tablero renombrado»\?/.test($('#preg-titulo').textContent), 'eliminar pide confirmación');
      $$('#preg-botones button')[0].click();
    } catch (e) { out.push('EXCEPCION ' + e.message + ' | ' + e.stack); terminar(); return; }
    setTimeout(paso5, 60);
  }

  function paso5() {
    try {
      ok(guardado().calculos.length === 0 && /Eliminaste «Tablero renombrado»/.test($('#aviso').textContent), 'eliminado de Mis cálculos');
      // Mis cálculos: semáforo de cada cálculo y duplicar
      $('[aria-controls="menu-calculos"]').click();
      $('[data-accion="guardar-calculo"]').click();
      $('#preg-nombre').value = 'Base';
      $$('#preg-botones button')[0].click();
      $('[aria-controls="menu-calculos"]').click();
      var item = $('#menu-calculos .calc-item');
      ok(item && /^\d+ (alcanza|cerca|no alcanza)$/.test(texto($('.calc-semaforo', item))), 'cada cálculo guardado muestra su semáforo → ' + texto(item && $('.calc-semaforo', item)));
      $('[data-accion="duplicar-calculo"]', item).click();
      var nombres = guardado().calculos.map(function (c) { return c.nombre; }).sort().join('|');
      ok(nombres === 'Base|Base (copia)' && /Duplicaste «Base»/.test($('#aviso').textContent), 'duplicar un cálculo → ' + nombres);

      // «Ver el ejemplo lleno»: el ejemplo no está en este tablero, así que se vuelve a agregar
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
    } catch (e) { out.push('EXCEPCION ' + e.message + ' | ' + e.stack); terminar(); return; }

    setTimeout(function () {
      var g = guardado();
      var ins = g && g.actual.indicadores.filter(function (x) { return x.nombre === 'Instalaciones en fecha (FTTH)'; })[0];
      ok(!!g && g.version === 3 && !!ins && JSON.stringify(g).indexOf('Churn prueba') < 0, 'se guarda en el navegador (versión 3)');
      ok(ins && ins.tipoSegmento === 'Transversal' && JSON.stringify(ins.segmentos) === JSON.stringify([{ nombre: 'Hogar', peso: 0.5, casos: 9000 }, { nombre: 'Pyme Fijo', peso: 0.5, casos: 1000 }]), 'se guardan los segmentos, sus casos y su parte → ' + JSON.stringify(ins && ins.segmentos));
      ok(ins && ins.dolores.filter(function (d) { return d.segmento === 'Pyme Fijo'; }).length === 1 && ins.dolores.filter(function (d) { return d.segmento === 'Hogar'; }).length === 15, 'y el segmento de cada dolor');
      terminar();
    }, 500);
  }

  function terminar() {
    out.push('errores JS: ' + JSON.stringify(errores));
    var pre = document.createElement('pre');
    pre.id = 'resultado-prueba';
    pre.textContent = out.join('\n');
    document.body.appendChild(pre);
  }
})();
