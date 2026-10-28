// Prueba de las animaciones de los pasteles. La ejecuta herramientas/probar_tablero.ps1 con el movimiento normal.
// Revisa que el pastel se despliegue desde cero, que el número del centro vaya contando y que todo termine
// exactamente en el valor del motor.
(function () {
  var errores = [];
  window.addEventListener('error', function (e) { errores.push(e.message + ' @' + e.lineno); });
  function $(s, r) { return (r || document).querySelector(s); }
  function $$(s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); }
  var out = [];
  function ok(c, m) { out.push((c ? 'OK    ' : 'FALLA ') + m); }
  function heroe() { return $('.pastel-valor').textContent.replace(/\s/g, ''); }
  function numero() { return parseInt(heroe(), 10); }
  function vuelta() { var c = $('#pastel-svg .anillo-vuelta'); return parseFloat(c.getAttribute('stroke-dasharray')); }
  var CIRC = 2 * Math.PI * 170;
  function terminar() {
    out.push('errores JS: ' + JSON.stringify(errores));
    var pre = document.createElement('pre');
    pre.id = 'resultado-prueba';
    pre.textContent = out.join('\n');
    document.body.appendChild(pre);
  }

  try {
    ok(!window.matchMedia('(prefers-reduced-motion: reduce)').matches, 'la prueba corre con movimiento normal');
    ok(heroe() === '0%', 'al cargar, el número del centro empieza en 0 → ' + heroe());
    ok($$('.tarjeta.entra').length === 2, 'las tarjetas del riel entran con animación');
    // Sin ventana, el navegador casi no pinta cuadros con tiempo virtual: se los damos con un temporizador de 16 ms
    window.requestAnimationFrame = function (cb) { return setTimeout(function () { cb(performance.now()); }, 16); };
    // Volver a entrar al ejemplo reinicia el despliegue, ya con esos cuadros
    $('[data-accion="ver-ejemplo"]').click();
    ok(heroe() === '0%', 'al entrar a un indicador el pastel vuelve a desplegarse desde 0');
  } catch (e) { out.push('EXCEPCION ' + e.message); }

  setTimeout(function () {
    try {
      var n = numero();
      ok(n > 0 && n < 85, 'a mitad de la animación el número va subiendo → ' + n);
      ok(vuelta() > 0 && vuelta() < 0.854 * CIRC, 'el anillo se va dibujando → ' + vuelta().toFixed(1));
    } catch (e) { out.push('EXCEPCION ' + e.message); }
    setTimeout(fin, 1700);
  }, 650);

  function fin() {
    try {
      var r = Motor.calcular(SuficienciaEjemplos()[0]);
      ok(heroe() === Math.round(r.real.suficiencia * 100) + '%', 'al terminar muestra la suficiencia del motor → ' + heroe());
      ok(Math.abs(vuelta() - r.real.suficiencia * CIRC) < 0.01, 'el anillo termina en la suficiencia → ' + vuelta().toFixed(2));
      ok($$('#pastel-svg .porcion').length === 5 && $$('#pastel-svg .porcion-resuelta').filter(function (p) { return p.style.display !== 'none'; }).length === 4, 'el pastel termina con sus 5 porciones');
      ok($$('.tarjeta-arte .porcion').length === 5, 'el pastel de la tarjeta también se despliega');

      // Al señalar un dolor su porción se asoma hacia afuera (sin la transición, para leer la posición final)
      var g = $('#pastel-svg .porcion');
      g.style.transition = 'none';
      g.dispatchEvent(new PointerEvent('pointermove', { bubbles: true, clientX: 200, clientY: 200 }));
      ok(g.classList.contains('es-foco') && $('#escenario').classList.contains('hay-foco'), 'la porción señalada queda resaltada');
      setTimeout(function () {
        try {
          var t = getComputedStyle(g).transform;
          var m = /matrix\(1, 0, 0, 1, ([-\d.]+), ([-\d.]+)\)/.exec(t);
          ok(m && Math.hypot(Number(m[1]), Number(m[2])) > 6.5, 'la porción señalada se asoma 7 px hacia afuera → ' + t);
          $('#pastel-svg').dispatchEvent(new PointerEvent('pointerleave'));

          // Un cambio en vivo pasa suavemente del valor anterior al nuevo
          var rango = $('input.rango');
          rango.value = '100';
          rango.dispatchEvent(new Event('input', { bubbles: true }));
          var ej = SuficienciaEjemplos()[0];
          ej.iniciativas[0].ejecucion = 1;
          var meta = Math.round(Motor.calcular(ej).real.suficiencia * 100);
          setTimeout(function () {
            var medio = numero();
            ok(medio > 85 && medio < meta, 'al mover el deslizador el número pasa por valores intermedios → ' + medio);
            setTimeout(function () {
              ok(numero() === meta, 'y termina en el nuevo valor → ' + numero() + ' %');
              terminar();
            }, 900);
          }, 120);
        } catch (e) { out.push('EXCEPCION ' + e.message); terminar(); }
      }, 600);
    } catch (e) { out.push('EXCEPCION ' + e.message + ' | ' + e.stack); terminar(); }
  }
})();
